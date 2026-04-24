/**
 * IFC Color Studio v2.0
 * ─────────────────────────────────────
 * Mỗi slot: nhiều file Excel, màu tùy chọn
 * Logic: v7/v9 đã proven (ẩn hết → hiện+màu → hiện lại)
 * Developed by Le Van Thao
 */

var RETRY_MAX   = 7;
var RETRY_DELAY = 2000;
var BATCH_CVT   = 500;
var BATCH_CLR   = 300;
var PAINT_DELAY = 150;

var _api = null;

// Mỗi slot: { color, files: [{name, guids}] }
var slots = [
  { color: "#00FF00", files: [] },
  { color: "#0088FF", files: [] }
];

var PRESETS = ["#00FF00","#0088FF","#FF6600","#FFD700","#FF0000","#CC00FF","#FF69B4","#00FFFF"];

/* ═══ UI ═══ */
function log(m,t){var e=document.getElementById("log");if(!e){console.log(m);return;}var s=document.createElement("span");if(t)s.className=t;s.textContent=m+"\n";e.appendChild(s);e.scrollTop=e.scrollHeight;}
function clearLog(){var e=document.getElementById("log");if(e)e.innerHTML="";}
function setStat(id,v){var e=document.getElementById(id);if(e)e.textContent=(v!=null)?v:"—";}
function setProgress(p){var w=document.getElementById("progWrap"),b=document.getElementById("progBar");if(!w||!b)return;if(p<=0){w.classList.remove("on");b.style.width="0%";return;}w.classList.add("on");b.style.width=Math.min(p,100)+"%";}
function lockUI(y){["applyBtn","resetBtn","saveBtn"].forEach(function(id){var e=document.getElementById(id);if(e)e.disabled=y;});}
function sleep(ms){return new Promise(function(r){setTimeout(r,ms);});}
function pad2(n){return String(n).padStart(2,"0");}
function fmtN(n){return typeof n==="number"?n.toLocaleString():String(n);}

function checkApplyBtn(){
  var hasGuids=slots.some(function(s){return s.files.some(function(f){return f.guids.length>0;});});
  document.getElementById("applyBtn").disabled=!hasGuids;
}

function isLight(hex){
  var r=parseInt(hex.substr(1,2),16),g=parseInt(hex.substr(3,2),16),b=parseInt(hex.substr(5,2),16);
  return(r*299+g*587+b*114)/1000>160;
}
function isValidHex(h){return/^#[0-9a-fA-F]{6}$/.test(h);}

/* ═══ Color picker setup ═══ */
function setColor(slotIdx, hex){
  hex=hex.toUpperCase();
  if(!isValidHex(hex))return;
  slots[slotIdx].color=hex;
  var n=slotIdx+1;
  document.getElementById("swatch"+n).style.background=hex;
  document.getElementById("picker"+n).value=hex;
  document.getElementById("hex"+n).value=hex;
  document.getElementById("s-c"+n).style.color=hex;
  var num=document.getElementById("num"+n);
  num.style.background=hex;
  num.style.color=isLight(hex)?"#000":"#fff";
}

function buildPresets(slotIdx){
  var n=slotIdx+1;
  var wrap=document.getElementById("presets"+n);
  wrap.innerHTML="";
  PRESETS.forEach(function(c){
    var d=document.createElement("div");
    d.className="preset";
    d.style.background=c;
    d.title=c;
    d.onclick=function(){setColor(slotIdx,c);};
    wrap.appendChild(d);
  });
}

function setupColorPicker(slotIdx){
  var n=slotIdx+1;
  var picker=document.getElementById("picker"+n);
  var hexInput=document.getElementById("hex"+n);
  var swatch=document.getElementById("swatch"+n);

  picker.addEventListener("input",function(){setColor(slotIdx,picker.value);});
  hexInput.addEventListener("input",function(){if(isValidHex(hexInput.value))setColor(slotIdx,hexInput.value);});
  hexInput.addEventListener("blur",function(){if(!isValidHex(hexInput.value))hexInput.value=slots[slotIdx].color;});
  swatch.addEventListener("click",function(e){if(e.target!==picker)picker.click();});
  buildPresets(slotIdx);
}

/* ═══ File list UI ═══ */
function renderFileList(slotIdx){
  var n=slotIdx+1;
  var list=document.getElementById("flist"+n);
  list.innerHTML="";
  var totalGuids=0;
  slots[slotIdx].files.forEach(function(f,fi){
    totalGuids+=f.guids.length;
    var item=document.createElement("div");
    item.className="file-item";
    item.innerHTML=
      '<span class="file-item-ok">✓</span>'+
      '<span class="file-item-name" title="'+f.name+'">'+f.name+'</span>'+
      '<span class="file-item-count">'+fmtN(f.guids.length)+' GUID</span>'+
      '<button class="file-item-del" title="Xóa">✕</button>';
    item.querySelector(".file-item-del").onclick=function(){
      slots[slotIdx].files.splice(fi,1);
      renderFileList(slotIdx);
      checkApplyBtn();
    };
    list.appendChild(item);
  });
  var fileCount=slots[slotIdx].files.length;
  document.getElementById("count"+n).textContent=fileCount+" file · "+fmtN(totalGuids)+" GUID";
}

function getAllGuids(slotIdx){
  var seen={},out=[];
  slots[slotIdx].files.forEach(function(f){
    f.guids.forEach(function(g){if(!seen[g]){seen[g]=1;out.push(g);}});
  });
  return out;
}

/* ═══ UUID ↔ IFC GUID ═══ */
var B64="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_$";
function to64(n,d){var r=[];for(var i=0;i<d;i++){r.push(B64.charAt(n%64));n=Math.floor(n/64);}return r.reverse().join("");}
function from64(s){var r=0;for(var i=0;i<s.length;i++){var x=B64.indexOf(s.charAt(i));if(x<0)return-1;r=r*64+x;}return r;}
function uuid2ifc(u){if(!u)return null;var h=String(u).replace(/-/g,"").toLowerCase();if(h.length!==32||!/^[0-9a-f]{32}$/.test(h))return null;var n=[parseInt(h.substr(0,2),16)];for(var i=0;i<5;i++)n.push(parseInt(h.substr(2+i*6,6),16));var r=to64(n[0],2);for(var i=1;i<6;i++)r+=to64(n[i],4);return r;}
function ifc2uuid(c){if(!c||c.length!==22)return null;var p=[from64(c.substr(0,2))];for(var i=0;i<5;i++)p.push(from64(c.substr(2+i*4,4)));if(p.some(function(x){return x<0;}))return null;var h=p[0].toString(16).padStart(2,"0");for(var i=1;i<6;i++)h+=p[i].toString(16).padStart(6,"0");return h.substr(0,8)+"-"+h.substr(8,4)+"-"+h.substr(12,4)+"-"+h.substr(16,4)+"-"+h.substr(20,12);}
function detectFmt(g){if(!g)return"x";var s=String(g).trim();if(s.length===36&&/^[0-9a-f]{8}-/i.test(s))return"uuid";if(s.length===32&&/^[0-9a-f]{32}$/i.test(s))return"nd";if(s.length===22)return"ifc";return"x";}

/* ═══ API ═══ */
async function getAPI(){if(_api)return _api;_api=await TrimbleConnectWorkspace.connect(window.parent,function(e,d){console.log("[T]",e,d);});log("Đã kết nối Trimble API.","ok");return _api;}

/* ═══ Excel ═══ */
function readWB(f){return new Promise(function(ok,no){var r=new FileReader();r.onload=function(e){try{ok(XLSX.read(e.target.result,{type:"array"}));}catch(err){no(err);}};r.onerror=no;r.readAsArrayBuffer(f);});}

function extractGuids(wb){
  if(!wb||!wb.SheetNames||!wb.SheetNames.length)throw new Error("Không có sheet.");
  var sn=wb.SheetNames[0];
  var rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{defval:""});
  if(!rows.length)throw new Error("Sheet trống.");
  var keys=Object.keys(rows[0]);
  var gk=keys.find(function(k){return k.trim().toUpperCase()==="GUID";});
  if(!gk)gk=keys[0];
  var seen={},out=[];
  rows.forEach(function(r){var g=String(r[gk]||"").trim();if(g&&!seen[g]){seen[g]=true;out.push(g);}});
  return out;
}

/* ═══ Model ═══ */
async function getModelIds(){
  var api=await getAPI();
  for(var a=1;a<=RETRY_MAX;a++){
    var raw;
    try{raw=await api.viewer.getObjects();}catch(e){if(a<RETRY_MAX){await sleep(RETRY_DELAY);continue;}throw e;}
    if(!Array.isArray(raw)||!raw.length){if(a<RETRY_MAX){await sleep(RETRY_DELAY);continue;}throw new Error("Viewer trống.");}
    var total=0,mids=[];
    raw.forEach(function(g){
      if(!g||!g.modelId)return;
      if(mids.indexOf(g.modelId)===-1)mids.push(g.modelId);
      if(Array.isArray(g.objects))total+=g.objects.length;
      else if(Array.isArray(g.objectRuntimeIds))total+=g.objectRuntimeIds.length;
      else if(Array.isArray(g.ids))total+=g.ids.length;
    });
    if(!mids.length){if(a<RETRY_MAX){await sleep(RETRY_DELAY);continue;}throw new Error("Không thấy modelId.");}
    return{modelIds:mids,total:total};
  }
}

/* ═══ Convert ═══ */
function flat(v){if(v==null)return[];if(typeof v==="number")return[v];if(Array.isArray(v)){var o=[];v.forEach(function(x){if(typeof x==="number")o.push(x);else if(Array.isArray(x))x.forEach(function(y){if(typeof y==="number")o.push(y);});});return o;}return[];}

async function batchConvert(api,mid,guids){
  var out=[];
  for(var i=0;i<guids.length;i+=BATCH_CVT){
    var c=guids.slice(i,i+BATCH_CVT);var r;
    try{r=await api.viewer.convertToObjectRuntimeIds(mid,c);}catch(e){for(var k=0;k<c.length;k++)out.push(null);continue;}
    if(!Array.isArray(r)){for(var k=0;k<c.length;k++)out.push(null);continue;}
    out=out.concat(r);
  }
  return out;
}

async function convertAll(api,modelIds,guids,label){
  var result=new Map();
  if(!guids.length)return result;
  var uuids=[],ifcs=[],others=[];
  guids.forEach(function(g){var f=detectFmt(g);if(f==="uuid"||f==="nd")uuids.push(g);else if(f==="ifc")ifcs.push(g);else others.push(g);});
  var u2i=uuids.map(uuid2ifc).filter(Boolean);
  var i2u=ifcs.map(ifc2uuid).filter(Boolean);

  for(var mi=0;mi<modelIds.length;mi++){
    var mid=modelIds[mi];var all=[];
    async function tryL(list,lbl){
      if(!list.length)return;var conv=await batchConvert(api,mid,list);var hit=0;
      for(var i=0;i<list.length;i++){var ids=flat(conv[i]);if(ids.length){hit++;all=all.concat(ids);}}
      if(hit>0)log("  ["+label+"/"+lbl+"] "+hit+" GUIDs matched","ok");
    }
    await tryL(uuids,"UUID");await tryL(ifcs,"IFC");await tryL(u2i,"U→I");await tryL(i2u,"I→U");await tryL(others,"RAW");
    if(all.length){var u={};all.forEach(function(id){u[id]=1;});result.set(mid,Object.keys(u).map(Number));}
  }
  return result;
}

/* ═══ Paint ═══ */
async function paintBatch(api,mid,ids,state){
  for(var i=0;i<ids.length;i+=BATCH_CLR){
    var chunk=ids.slice(i,i+BATCH_CLR);
    try{await api.viewer.setObjectState({modelObjectIds:[{modelId:mid,objectRuntimeIds:chunk}]},state);}catch(e){}
    if(i+BATCH_CLR<ids.length)await sleep(PAINT_DELAY);
  }
}

/* ═══ MAIN ═══ */
async function applyColors(){
  lockUI(true);clearLog();setProgress(5);
  try{
    var api=await getAPI();

    log("Reset...","info");
    try{await api.viewer.setObjectState(undefined,{color:"reset",visible:"reset"});}catch(e){}
    await sleep(500);
    setProgress(10);

    var mi=await getModelIds();
    setStat("s-total",fmtN(mi.total));
    setProgress(18);

    // Convert GUIDs cho từng slot
    log("Map GUIDs...","info");
    var maps=[];
    var totals=[];
    for(var si=0;si<slots.length;si++){
      var guids=getAllGuids(si);
      log("  Slot "+(si+1)+": "+fmtN(guids.length)+" GUID duy nhất","info");
      var m=await convertAll(api,mi.modelIds,guids,"Slot"+(si+1));
      maps.push(m);
      var cnt=0;m.forEach(function(ids){cnt+=ids.length;});
      totals.push(cnt);
      setStat("s-c"+(si+1),cnt>0?fmtN(cnt):"—");
    }
    setProgress(35);

    if(totals.every(function(t){return t===0;})){
      log("✗ Không match object nào!","err");setProgress(0);lockUI(false);checkApplyBtn();return;
    }

    // ẨN TẤT CẢ
    log("Ẩn toàn bộ model...","info");
    try{await api.viewer.setObjectState(undefined,{visible:false});}catch(e){}
    await sleep(800);
    setProgress(45);

    // TÔ TỪNG SLOT
    var step=0;var totalSteps=slots.length;
    for(var si=0;si<slots.length;si++){
      if(totals[si]===0){step++;continue;}
      var color=slots[si].color;
      log("Tô Slot "+(si+1)+" ("+color+"): "+fmtN(totals[si])+" objects...","info");
      for(var mi2=0;mi2<mi.modelIds.length;mi2++){
        var mid=mi.modelIds[mi2];
        var ids=maps[si].get(mid);
        if(!ids||!ids.length)continue;
        await paintBatch(api,mid,ids,{visible:true,color:color});
        log("  ▪ "+fmtN(ids.length)+" objects","ok");
      }
      step++;
      setProgress(45+step/totalSteps*35);
      await sleep(300);
    }

    // HIỆN LẠI PHẦN CÒN LẠI
    log("Hiện phần còn lại (màu gốc)...","info");
    try{await api.viewer.setObjectState(undefined,{visible:true});}catch(e){}
    await sleep(500);
    setProgress(100);

    log("","info");
    log("✓ HOÀN TẤT!","ok");
    for(var si=0;si<slots.length;si++){
      if(totals[si]>0) log("  Slot "+(si+1)+" ("+slots[si].color+"): "+fmtN(totals[si])+" cấu kiện","ok");
    }
    setTimeout(function(){setProgress(0);},2000);

  }catch(err){
    log("✗ "+(err&&err.message?err.message:String(err)),"err");setProgress(0);
  }finally{lockUI(false);checkApplyBtn();}
}

/* ═══ Reset ═══ */
async function resetViewer(){
  lockUI(true);clearLog();setProgress(10);
  try{var api=await getAPI();try{await api.viewer.setObjectState(undefined,{color:"reset",visible:"reset"});}catch(e){}await api.viewer.reset();
  setStat("s-total","—");setStat("s-c1","—");setStat("s-c2","—");
  setProgress(100);log("✓ Reset OK.","ok");setTimeout(function(){setProgress(0);},1000);}
  catch(e){log("✗ "+(e&&e.message?e.message:String(e)),"err");setProgress(0);}
  finally{lockUI(false);checkApplyBtn();}
}

/* ═══ Save View ═══ */
async function saveView(){
  try{var api=await getAPI();var inp=document.getElementById("viewName");var name=inp?inp.value.trim():"";
  if(!name){var n=new Date();name="ColorStudio "+n.getFullYear()+"-"+pad2(n.getMonth()+1)+"-"+pad2(n.getDate())+" "+pad2(n.getHours())+":"+pad2(n.getMinutes());if(inp)inp.value=name;}
  var c=await api.view.createView({name:name,description:"IFC Color Studio v2.0 | Le Van Thao"});
  if(!c||!c.id)throw new Error("No view ID.");await api.view.updateView({id:c.id});await api.view.selectView(c.id);
  log('✓ View: "'+name+'"',"ok");}
  catch(e){log("✗ "+(e&&e.message?e.message:String(e)),"err");}
}

/* ═══ File handling ═══ */
async function handleFiles(fileList, slotIdx){
  var files=Array.from(fileList);
  if(!files.length)return;
  log("Đọc "+files.length+" file cho Slot "+(slotIdx+1)+"...","info");
  for(var i=0;i<files.length;i++){
    var f=files[i];
    try{
      var wb=await readWB(f);
      var guids=extractGuids(wb);
      // Kiểm tra trùng tên file trong slot
      var existing=slots[slotIdx].files.findIndex(function(x){return x.name===f.name;});
      if(existing>=0){
        slots[slotIdx].files[existing]={name:f.name,guids:guids};
        log("  ↻ Cập nhật: "+f.name+" ("+fmtN(guids.length)+" GUID)","info");
      }else{
        slots[slotIdx].files.push({name:f.name,guids:guids});
        log("  ✓ "+f.name+" ("+fmtN(guids.length)+" GUID)","ok");
      }
    }catch(e){
      log("  ✗ "+f.name+": "+(e&&e.message?e.message:String(e)),"err");
    }
  }
  renderFileList(slotIdx);
  checkApplyBtn();
}

/* ═══ Init ═══ */
window.addEventListener("load",function(){
  // Setup color pickers
  setupColorPicker(0);
  setupColorPicker(1);

  // File inputs
  document.getElementById("file1").addEventListener("change",function(){handleFiles(this.files,0);this.value="";});
  document.getElementById("file2").addEventListener("change",function(){handleFiles(this.files,1);this.value="";});

  // Drag & drop
  function setupDrop(zoneId,slotIdx){
    var z=document.getElementById(zoneId);
    z.addEventListener("dragover",function(e){e.preventDefault();z.classList.add("over");});
    z.addEventListener("dragleave",function(){z.classList.remove("over");});
    z.addEventListener("drop",function(e){
      e.preventDefault();z.classList.remove("over");
      if(e.dataTransfer&&e.dataTransfer.files&&e.dataTransfer.files.length)
        handleFiles(e.dataTransfer.files,slotIdx);
    });
  }
  setupDrop("zone1",0);
  setupDrop("zone2",1);

  document.getElementById("applyBtn").addEventListener("click",applyColors);
  document.getElementById("resetBtn").addEventListener("click",resetViewer);
  document.getElementById("saveBtn").addEventListener("click",saveView);
});
