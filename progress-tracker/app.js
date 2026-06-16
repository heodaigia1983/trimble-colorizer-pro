/**
 * Tiến Độ Lắp Dựng T2 v1.0
 * ─────────────────────────────────────
 * Theo dõi tiến độ thi công kết cấu thép qua 3 trạng thái:
 * Chưa lắp / Đang lắp / Đã lắp.
 * Gán trạng thái bằng cách CHỌN TRỰC TIẾP trên Viewer.
 * Lưu / nạp lại tiến độ bằng file Excel (GUID + Trạng thái).
 * Developed by Le Van Thao
 */

var RETRY_MAX   = 7;
var RETRY_DELAY = 2000;
var BATCH_CVT   = 500;
var BATCH_CLR   = 300;
var PAINT_DELAY = 150;

var STATUS = [
  {label:"Chưa lắp", color:"#6b7280"},
  {label:"Đang lắp", color:"#FFD700"},
  {label:"Đã lắp",   color:"#00C853"}
];

var _api = null;
var _state = new Map();   // guid (string) -> statusIdx (0|1|2)
var _modelInfo = null;    // {modelIds:[...], total:N} - cache
var _busy = false;

/* ═══ UI helpers ═══ */
function log(m,t){var e=document.getElementById("log");if(!e){console.log(m);return;}var s=document.createElement("span");if(t)s.className=t;s.textContent=m+"\n";e.appendChild(s);e.scrollTop=e.scrollHeight;console.log("["+(t||"")+"] "+m);}
function clearLog(){var e=document.getElementById("log");if(e)e.innerHTML="";}
function setStat(id,v){var e=document.getElementById(id);if(e)e.textContent=(v!=null)?v:"—";}
function setProgress(p){var w=document.getElementById("progWrap"),b=document.getElementById("progBar");if(!w||!b)return;if(p<=0){w.classList.remove("on");b.style.width="0%";return;}w.classList.add("on");b.style.width=Math.min(p,100)+"%";}
function lockUI(y){_busy=y;["btnS0","btnS1","btnS2","exportBtn","resetBtn"].forEach(function(id){var e=document.getElementById(id);if(e)e.disabled=y;});}
function sleep(ms){return new Promise(function(r){setTimeout(r,ms);});}
function pad2(n){return String(n).padStart(2,"0");}
function fmtN(n){return typeof n==="number"?n.toLocaleString():String(n);}

/* ═══ UUID ↔ IFC GUID (giữ nguyên logic từ IFC Color Studio) ═══ */
var B64="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_$";
function to64(n,d){var r=[];for(var i=0;i<d;i++){r.push(B64.charAt(n%64));n=Math.floor(n/64);}return r.reverse().join("");}
function from64(s){var r=0;for(var i=0;i<s.length;i++){var x=B64.indexOf(s.charAt(i));if(x<0)return-1;r=r*64+x;}return r;}
function uuid2ifc(u){if(!u)return null;var h=String(u).replace(/-/g,"").toLowerCase();if(h.length!==32||!/^[0-9a-f]{32}$/.test(h))return null;var n=[parseInt(h.substr(0,2),16)];for(var i=0;i<5;i++)n.push(parseInt(h.substr(2+i*6,6),16));var r=to64(n[0],2);for(var i=1;i<6;i++)r+=to64(n[i],4);return r;}
function ifc2uuid(c){if(!c||c.length!==22)return null;var p=[from64(c.substr(0,2))];for(var i=0;i<5;i++)p.push(from64(c.substr(2+i*4,4)));if(p.some(function(x){return x<0;}))return null;var h=p[0].toString(16).padStart(2,"0");for(var i=1;i<6;i++)h+=p[i].toString(16).padStart(6,"0");return h.substr(0,8)+"-"+h.substr(8,4)+"-"+h.substr(12,4)+"-"+h.substr(16,4)+"-"+h.substr(20,12);}
function detectFmt(g){if(!g)return"x";var s=String(g).trim();if(s.length===36&&/^[0-9a-f]{8}-/i.test(s))return"uuid";if(s.length===32&&/^[0-9a-f]{32}$/i.test(s))return"nd";if(s.length===22)return"ifc";return"x";}

/* ═══ API ═══ */
async function getAPI(){if(_api)return _api;_api=await TrimbleConnectWorkspace.connect(window.parent,function(e,d){console.log("[T]",e,d);});log("Đã kết nối Trimble API.","ok");return _api;}

/* ═══ Model info (tổng số cấu kiện trong model) ═══ */
async function getModelInfo(force){
  if(_modelInfo&&!force)return _modelInfo;
  var api=await getAPI();
  for(var a=1;a<=RETRY_MAX;a++){
    var raw;
    try{raw=await api.viewer.getObjects();}catch(e){if(a<RETRY_MAX){await sleep(RETRY_DELAY);continue;}throw e;}
    if(!Array.isArray(raw)||!raw.length){if(a<RETRY_MAX){await sleep(RETRY_DELAY);continue;}throw new Error("Viewer trống. Đợi model load.");}
    var total=0,mids=[];
    raw.forEach(function(g){
      if(!g||!g.modelId)return;
      if(mids.indexOf(g.modelId)===-1)mids.push(g.modelId);
      var ids=flat(g.objects||g.objectRuntimeIds||g.ids);
      total+=ids.length;
    });
    if(!mids.length){if(a<RETRY_MAX){await sleep(RETRY_DELAY);continue;}throw new Error("Không thấy modelId.");}
    _modelInfo={modelIds:mids,total:total};
    return _modelInfo;
  }
}

/* ═══ Convert helpers ═══ */
function flat(v){if(v==null)return[];if(typeof v==="number")return[v];if(Array.isArray(v)){var o=[];v.forEach(function(x){if(typeof x==="number")o.push(x);else if(Array.isArray(x))x.forEach(function(y){if(typeof y==="number")o.push(y);});});return o;}return[];}

async function batchConvertToRuntime(api,mid,guids){
  var out=[];
  for(var i=0;i<guids.length;i+=BATCH_CVT){
    var c=guids.slice(i,i+BATCH_CVT);var r;
    try{r=await api.viewer.convertToObjectRuntimeIds(mid,c);}catch(e){for(var k=0;k<c.length;k++)out.push(null);continue;}
    if(!Array.isArray(r)){for(var k=0;k<c.length;k++)out.push(null);continue;}
    out=out.concat(r);
  }
  return out;
}

async function convertGuidsToRuntime(api,modelIds,guids,label){
  var result=new Map();
  if(!guids.length)return result;
  var uuids=[],ifcs=[],others=[];
  guids.forEach(function(g){var f=detectFmt(g);if(f==="uuid"||f==="nd")uuids.push(g);else if(f==="ifc")ifcs.push(g);else others.push(g);});
  var u2i=uuids.map(uuid2ifc).filter(Boolean);
  var i2u=ifcs.map(ifc2uuid).filter(Boolean);
  for(var mi=0;mi<modelIds.length;mi++){
    var mid=modelIds[mi];var all=[];
    async function tryL(list,lbl){
      if(!list.length)return;var conv=await batchConvertToRuntime(api,mid,list);var hit=0;
      for(var i=0;i<list.length;i++){var ids=flat(conv[i]);if(ids.length){hit++;all=all.concat(ids);}}
      if(hit>0)log("  ["+label+"/"+lbl+"] "+hit+" GUID khớp","ok");
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

/* ═══ Stats ═══ */
function refreshStats(){
  var counts=[0,0,0];
  _state.forEach(function(idx){if(counts[idx]!=null)counts[idx]++;});
  var total=_modelInfo?_modelInfo.total:null;
  var haveData=_state.size>0||total!=null;
  setStat("s-total",total!=null?fmtN(total):"—");
  for(var i=0;i<3;i++){
    var c=counts[i];
    setStat("s-c"+i, haveData?fmtN(c):"—");
    var p=document.getElementById("s-p"+i);
    if(p)p.textContent=(total&&c)?Math.round(c/total*100)+"%":"";
  }
}

/* ═══ Gán trạng thái cho lựa chọn hiện tại trên Viewer ═══ */
async function assignStatus(idx){
  if(_busy)return;
  lockUI(true);
  try{
    var api=await getAPI();
    var raw=await api.viewer.getSelection();
    if(!Array.isArray(raw)||!raw.length){log("✗ Chưa chọn cấu kiện nào trên viewer.","warn");return;}

    var groups=[];
    raw.forEach(function(g){
      if(!g||!g.modelId)return;
      var ids=flat(g.objects||g.objectRuntimeIds||g.ids);
      if(ids.length)groups.push({modelId:g.modelId,ids:ids});
    });
    if(!groups.length){log("✗ Chưa chọn cấu kiện nào trên viewer.","warn");return;}

    // lấy tổng số cấu kiện model (cho % thống kê) — không chặn nếu lỗi
    if(!_modelInfo){try{await getModelInfo();}catch(e){}}

    var st=STATUS[idx];
    var done=0,savedGuid=0;
    for(var i=0;i<groups.length;i++){
      var grp=groups[i];
      // tô màu ngay bằng runtime id đã có sẵn (không cần convert)
      await paintBatch(api,grp.modelId,grp.ids,{color:st.color});
      done+=grp.ids.length;
      // lấy GUID bền (portable) để lưu trạng thái — dùng convertToObjectIds
      var guids=null;
      try{guids=await api.viewer.convertToObjectIds(grp.modelId,grp.ids);}catch(e){guids=null;}
      if(Array.isArray(guids)){
        guids.forEach(function(g){if(g){_state.set(String(g),idx);savedGuid++;}});
      }
    }
    if(savedGuid===0){
      log('⚠ Đã tô màu "'+st.label+'" nhưng KHÔNG lưu được GUID (API convertToObjectIds có thể không khả dụng) — trạng thái sẽ KHÔNG xuất được ra Excel cho lượt này.',"warn");
    }
    refreshStats();
    log('✓ Gán "'+st.label+'" cho '+fmtN(done)+' cấu kiện.',"ok");
  }catch(e){
    log("✗ "+(e&&e.message?e.message:String(e)),"err");
  }finally{
    lockUI(false);
  }
}

/* ═══ Xuất Excel (lưu tiến độ) ═══ */
function exportExcel(){
  if(!_state.size){log("✗ Chưa có dữ liệu tiến độ để xuất.","warn");return;}
  var n=new Date();
  var ts=n.getFullYear()+"-"+pad2(n.getMonth()+1)+"-"+pad2(n.getDate())+" "+pad2(n.getHours())+":"+pad2(n.getMinutes());
  var rows=[];
  _state.forEach(function(idx,guid){
    rows.push({GUID:guid,"Trạng thái":STATUS[idx].label,"Cập nhật":ts});
  });
  var ws=XLSX.utils.json_to_sheet(rows);
  var wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"TienDo");
  var fname="TienDoT2_"+n.getFullYear()+pad2(n.getMonth()+1)+pad2(n.getDate())+"_"+pad2(n.getHours())+pad2(n.getMinutes())+".xlsx";
  XLSX.writeFile(wb,fname);
  log('✓ Đã xuất "'+fname+'" ('+fmtN(rows.length)+' dòng).',"ok");
}

/* ═══ Nhập Excel (nạp lại tiến độ + tự tô màu) ═══ */
function readWB(f){return new Promise(function(ok,no){var r=new FileReader();r.onload=function(e){try{ok(XLSX.read(e.target.result,{type:"array"}));}catch(err){no(err);}};r.onerror=no;r.readAsArrayBuffer(f);});}

function statusIdxFromLabel(v){
  var s=String(v||"").trim().toLowerCase();
  for(var i=0;i<STATUS.length;i++){if(STATUS[i].label.toLowerCase()===s)return i;}
  if(s==="0"||s==="1"||s==="2")return parseInt(s,10);
  return -1;
}

async function importExcel(file){
  if(_busy)return;
  lockUI(true);clearLog();setProgress(5);
  try{
    log('Đọc "'+file.name+'"...',"info");
    var wb=await readWB(file);
    var sn=wb.SheetNames[0];
    var rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{defval:""});
    if(!rows.length)throw new Error("File Excel trống.");
    var keys=Object.keys(rows[0]);
    var gk=keys.find(function(k){return k.trim().toUpperCase()==="GUID";})||keys[0];
    var sk=keys.find(function(k){return k.trim().toLowerCase().indexOf("trạng thái")>=0||k.trim().toLowerCase()==="status";})||keys[1];

    var newState=new Map();
    var skipped=0;
    rows.forEach(function(r){
      var g=String(r[gk]||"").trim();
      var idx=statusIdxFromLabel(r[sk]);
      if(!g||idx<0){skipped++;return;}
      newState.set(g,idx);
    });
    if(!newState.size)throw new Error("Không đọc được dòng hợp lệ nào (kiểm tra cột GUID / Trạng thái).");
    log("  ✓ Đọc được "+fmtN(newState.size)+" dòng"+(skipped?(", bỏ qua "+skipped+" dòng lỗi"):"")+".","ok");
    setProgress(15);

    var api=await getAPI();
    log("Reset màu hiện tại...","info");
    try{await api.viewer.setObjectState(undefined,{color:"reset"});}catch(e){}
    await sleep(400);
    setProgress(25);

    var mi=await getModelInfo(true);
    setProgress(35);

    // nhóm GUID theo trạng thái
    var byStatus=[[],[],[]];
    newState.forEach(function(idx,g){if(byStatus[idx])byStatus[idx].push(g);});

    var stepBase=35,stepRange=55;
    for(var s=0;s<STATUS.length;s++){
      var list=byStatus[s];
      if(!list.length)continue;
      log('Tô màu "'+STATUS[s].label+'" ('+STATUS[s].color+'): '+fmtN(list.length)+' GUID...',"info");
      var map=await convertGuidsToRuntime(api,mi.modelIds,list,STATUS[s].label);
      var painted=0;
      for(var m=0;m<mi.modelIds.length;m++){
        var mid=mi.modelIds[m];var ids=map.get(mid);
        if(!ids||!ids.length)continue;
        await paintBatch(api,mid,ids,{color:STATUS[s].color});
        painted+=ids.length;
      }
      log("  ▪ Tô được "+fmtN(painted)+"/"+fmtN(list.length)+" cấu kiện","ok");
      setProgress(stepBase+stepRange*((s+1)/STATUS.length));
    }

    _state=newState;
    refreshStats();
    setProgress(100);
    log("✓ HOÀN TẤT nhập tiến độ.","ok");
    setTimeout(function(){setProgress(0);},1500);
  }catch(e){
    log("✗ "+(e&&e.message?e.message:String(e)),"err");setProgress(0);
  }finally{
    lockUI(false);
  }
}

/* ═══ Reset ═══ */
async function resetAll(){
  if(_busy)return;
  lockUI(true);clearLog();setProgress(10);
  try{
    var api=await getAPI();
    try{await api.viewer.setObjectState(undefined,{color:"reset",visible:"reset"});}catch(e){}
    await api.viewer.reset();
    _state=new Map();
    _modelInfo=null;
    refreshStats();
    setProgress(100);
    log("✓ Đã reset toàn bộ tiến độ và màu sắc.","ok");
    setTimeout(function(){setProgress(0);},1000);
  }catch(e){
    log("✗ "+(e&&e.message?e.message:String(e)),"err");setProgress(0);
  }finally{
    lockUI(false);
  }
}

/* ═══ Events ═══ */
document.getElementById("btnS0").addEventListener("click",function(){assignStatus(0);});
document.getElementById("btnS1").addEventListener("click",function(){assignStatus(1);});
document.getElementById("btnS2").addEventListener("click",function(){assignStatus(2);});
document.getElementById("exportBtn").addEventListener("click",exportExcel);
document.getElementById("resetBtn").addEventListener("click",resetAll);
document.getElementById("fileImport").addEventListener("change",function(){
  var f=this.files&&this.files[0];
  if(f)importExcel(f);
  this.value="";
});

refreshStats();
