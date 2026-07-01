/**
 * Model Control Center v2.0
 * ─────────────────────────────────────
 * 2 file Excel, màu tùy chọn cho mỗi file
 * Còn lại giữ màu gốc
 * Logic dựa trên v7/v9 đã chứng minh hoạt động (không đổi)
 * UI: Tailwind CSS, theme Technical Midnight, KPI badges, mobile-first
 * Developed by Le Van Thao
 */

var RETRY_MAX   = 7;
var RETRY_DELAY = 2000;
var BATCH_CVT   = 500;
var BATCH_CLR   = 300;
var PAINT_DELAY = 150;

var _api = null;
var _guids1 = [];
var _guids2 = [];
var _color1 = "#00FF00";
var _color2 = "#0088FF";
var _color3 = "#FF0000";
var _selMap = null;
var _map1 = null;
var _map2 = null;
var _noteMarkupId = {1:null,2:null,3:null};
var MARKUP_COLOR = "#FF1493";
var _colorLedger = new Map();
var _colorLedgerSummary = null;
var _currentViewId = null;
var _ledgerViewId = null;
var _viewPollInterval = null;
var _loadingView = false;
var _lastQty = {};

/* ═══ Option 4 — Assembly Position Prefix Filter (highlight tạm, KHÔNG ledger) ═══ */
var _apCache = null;          // Map<objectKey, {modelId, runtimeId, ap, name}>
var _apScanning = false;
var _apSelected = new Set();  // objectKeys đang highlight tạm thời
var AP_HIGHLIGHT = "#9334e6";

function makeObjectKey(modelId, runtimeId){
  return modelId+":"+runtimeId;
}
function splitObjectKey(key){
  var idx=String(key).indexOf(":");
  return {modelId:key.substr(0,idx), runtimeId:key.substr(idx+1)};
}
function getNoteForColor(hex){
  var note="";
  for(var entry of _colorLedger){
    if(entry[1]&&entry[1].color===hex && entry[1].note){
      note=entry[1].note;
    }
  }
  return note;
}
function updateColorNoteForColor(hex, note){
  for(var entry of _colorLedger){
    if(entry[1]&&entry[1].color===hex){
      entry[1].note = note;
    }
  }
  if(_colorLedgerSummary && _colorLedgerSummary.has(hex)){
    var info=_colorLedgerSummary.get(hex);
    info.note = note;
  }
  saveLedger();
}
function buildColorGroups(){
  var groups = new Map();
  for(var entry of _colorLedger){
    var key=entry[0], data=entry[1];
    if(!data||!data.color) continue;
    var color=data.color;
    var info=groups.get(color);
    if(!info){
      info={note:"", ids:new Map()};
      groups.set(color, info);
    }
    if(data.note!=null) info.note=data.note;
    var obj=splitObjectKey(key);
    var modelMap=info.ids.get(obj.modelId);
    if(!modelMap){
      modelMap=[];
      info.ids.set(obj.modelId, modelMap);
    }
    modelMap.push(obj.runtimeId);
  }
  return groups;
}
function buildLedgerMapForColor(color){
  var groups=buildColorGroups();
  var info=groups.get(color);
  if(!info) return null;
  var map=new Map();
  info.ids.forEach(function(ids,modelId){map.set(modelId,ids.slice());});
  return map;
}
function buildLedgerMapAll(){
  var map=new Map();
  for(var entry of _colorLedger){
    var key=entry[0];
    var obj=splitObjectKey(key);
    var ids=map.get(obj.modelId);
    if(!ids){ids=[];map.set(obj.modelId,ids);}
    ids.push(obj.runtimeId);
  }
  return map;
}
async function rebuildColorLedgerSummary(){
  _colorLedgerSummary = new Map();
  var groups=buildColorGroups();
  for(var entry of groups){
    var color=entry[0], info=entry[1];
    var qty=0;
    info.ids.forEach(function(ids){qty+=ids.length;});
    var weight=0;
    try{
      var api=await getAPI();
      var q=await fetchQuantities(api, buildLedgerMapForColor(color));
      q.groups.forEach(function(g){weight += g.weight;});
    }catch(e){weight=0;}
    _colorLedgerSummary.set(color, {count:qty, weight:weight, note:info.note});
  }
}

/* ═══ UI ═══ */
/* Pattern → câu giải thích thân thiện (chỉ áp dụng cho hiển thị #log, console.log vẫn giữ message gốc) */
var ERROR_HINTS=[
  {pattern:/dispatcher/i,hint:"(Lỗi kết nối với viewer, thử lại sau vài giây)"},
  {pattern:/timeout/i,hint:"(Lỗi kết nối với viewer, thử lại sau vài giây)"},
  {pattern:/viewer trống|đợi model load/i,hint:"(Model 3D chưa load xong, đợi vài giây rồi thử lại)"},
  {pattern:/không thấy modelid/i,hint:"(Không tìm thấy model nào đang mở trong viewer)"},
  {pattern:/không có sheet|sheet trống/i,hint:"(File Excel không có dữ liệu, kiểm tra lại file)"},
  {pattern:/no view id/i,hint:"(Không tạo được view, thử lại sau)"}
];
function friendlyHint(msg){if(!msg)return null;for(var i=0;i<ERROR_HINTS.length;i++){if(ERROR_HINTS[i].pattern.test(msg))return ERROR_HINTS[i].hint;}return null;}
var LOG_CLR={ok:"text-sys-green",info:"text-accent",warn:"text-amber-400",err:"text-sys-red"};
function log(m,t){console.log("["+(t||"")+"] "+m);var e=document.getElementById("log");if(!e)return;var s=document.createElement("span");s.className=LOG_CLR[t]||"";var hint=friendlyHint(m);s.textContent=m+(hint?"\n"+hint:"")+"\n";e.appendChild(s);e.scrollTop=e.scrollHeight;}
function clearLog(){var e=document.getElementById("log");if(e)e.innerHTML="";}
function setStat(id,v){var e=document.getElementById(id);if(e)e.textContent=(v!=null)?v:"—";}
function setProgress(p){var w=document.getElementById("progWrap"),b=document.getElementById("progBar");if(!w||!b)return;if(p<=0){w.classList.add("hidden");b.style.width="0%";return;}w.classList.remove("hidden");b.style.width=Math.min(p,100)+"%";}
function lockUI(y){["resetBtn","saveBtn"].forEach(function(id){var e=document.getElementById(id);if(e)e.disabled=y;});}
function sleep(ms){return new Promise(function(r){setTimeout(r,ms);});}
function pad2(n){return String(n).padStart(2,"0");}
function fmtN(n){return typeof n==="number"?n.toLocaleString():String(n);}
function checkApplyBtn(slot){var g=slot===1?_guids1:_guids2;document.getElementById("applyBtn"+slot).disabled=!g.length;}

/* ═══ KPI status badge (Pending/Ready/Error) ═══ */
var BADGE_CFG={
  pending:{text:"Pending",cls:"bg-slate-200 text-slate-600"},
  error:{text:"Error",cls:"bg-sys-red/20 text-sys-red"}
};
function setBadge(slot,status){
  var el=document.getElementById("badge-s"+slot);
  if(!el)return;
  var cfg;
  if(status==="ready"){var c=slot===1?"green":"blue";cfg={text:"Ready",cls:"bg-sys-"+c+"/20 text-sys-"+c};}
  else cfg=BADGE_CFG[status]||BADGE_CFG.pending;
  el.textContent=cfg.text;
  el.className="text-[9px] font-semibold px-1.5 py-0.5 rounded-full "+cfg.cls;
}

/* ═══ COLOR PICKER ═══ */
function isValidHex(h){return/^#[0-9a-fA-F]{6}$/.test(h);}

function setColor(slot, hex){
  hex = hex.toUpperCase();
  if(!isValidHex(hex)) return;
  if(slot===1){
    _color1=hex;
    document.getElementById("swatch1").style.background=hex;
    document.getElementById("picker1").value=hex;
    document.getElementById("hex1").value=hex;
    document.getElementById("s-c1").style.color=hex;
    document.getElementById("num1").style.background=hex;
    document.getElementById("num1").style.color=isLightColor(hex)?"#000":"#fff";
  } else if(slot===2){
    _color2=hex;
    document.getElementById("swatch2").style.background=hex;
    document.getElementById("picker2").value=hex;
    document.getElementById("hex2").value=hex;
    document.getElementById("s-c2").style.color=hex;
    document.getElementById("num2").style.background=hex;
    document.getElementById("num2").style.color=isLightColor(hex)?"#000":"#fff";
  } else {
    _color3=hex;
    document.getElementById("swatch3").style.background=hex;
    document.getElementById("picker3").value=hex;
    document.getElementById("hex3").value=hex;
    document.getElementById("num3").style.background=hex;
    document.getElementById("num3").style.color=isLightColor(hex)?"#000":"#fff";
    var ni=document.getElementById("noteInput");
    if(ni) ni.value=getNoteForColor(hex);
  }
}

function isLightColor(hex){
  var r=parseInt(hex.substr(1,2),16),g=parseInt(hex.substr(3,2),16),b=parseInt(hex.substr(5,2),16);
  return (r*299+g*587+b*114)/1000>128;
}

// Init color pickers
window.addEventListener("load",function(){
  // Picker 1
  var p1=document.getElementById("picker1");
  var h1=document.getElementById("hex1");
  p1.addEventListener("input",function(){setColor(1,p1.value);});
  h1.addEventListener("input",function(){if(isValidHex(h1.value))setColor(1,h1.value);});
  h1.addEventListener("blur",function(){if(!isValidHex(h1.value))h1.value=_color1;});
  document.getElementById("swatch1").addEventListener("click",function(){p1.click();});

  // Picker 2
  var p2=document.getElementById("picker2");
  var h2=document.getElementById("hex2");
  p2.addEventListener("input",function(){setColor(2,p2.value);});
  h2.addEventListener("input",function(){if(isValidHex(h2.value))setColor(2,h2.value);});
  h2.addEventListener("blur",function(){if(!isValidHex(h2.value))h2.value=_color2;});
  document.getElementById("swatch2").addEventListener("click",function(){p2.click();});

  // Picker 3 (viewer selection)
  var p3=document.getElementById("picker3");
  var h3=document.getElementById("hex3");
  p3.addEventListener("input",function(){setColor(3,p3.value);});
  h3.addEventListener("input",function(){if(isValidHex(h3.value))setColor(3,h3.value);});
  h3.addEventListener("blur",function(){if(!isValidHex(h3.value))h3.value=_color3;});
  document.getElementById("swatch3").addEventListener("click",function(){p3.click();});

  // Init colors
  setColor(1,_color1);
  setColor(2,_color2);
  setColor(3,_color3);

  // Init KPI badges
  setBadge(1,"pending");
  setBadge(2,"pending");

  // Drag & drop
  setupDrop("zone1","file1");
  setupDrop("zone2","file2");
});

function setupDrop(zoneId, fileId){
  var z=document.getElementById(zoneId);
  var overCls=["border-accent","bg-accent/5"];
  z.addEventListener("dragover",function(e){e.preventDefault();z.classList.add.apply(z.classList,overCls);});
  z.addEventListener("dragleave",function(){z.classList.remove.apply(z.classList,overCls);});
  z.addEventListener("drop",function(e){
    e.preventDefault();z.classList.remove.apply(z.classList,overCls);
    var f=e.dataTransfer&&e.dataTransfer.files&&e.dataTransfer.files[0];
    if(f){document.getElementById(fileId).files=e.dataTransfer.files;document.getElementById(fileId).dispatchEvent(new Event("change"));}
  });
}

/* ═══ UUID ↔ IFC GUID ═══ */
var B64="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_$";
function to64(n,d){var r=[];for(var i=0;i<d;i++){r.push(B64.charAt(n%64));n=Math.floor(n/64);}return r.reverse().join("");}
function from64(s){var r=0;for(var i=0;i<s.length;i++){var x=B64.indexOf(s.charAt(i));if(x<0)return-1;r=r*64+x;}return r;}
function uuid2ifc(u){if(!u)return null;var h=String(u).replace(/-/g,"").toLowerCase();if(h.length!==32||!/^[0-9a-f]{32}$/.test(h))return null;var n=[parseInt(h.substr(0,2),16)];for(var i=0;i<5;i++)n.push(parseInt(h.substr(2+i*6,6),16));var r=to64(n[0],2);for(var i=1;i<6;i++)r+=to64(n[i],4);return r;}
function ifc2uuid(c){if(!c||c.length!==22)return null;var p=[from64(c.substr(0,2))];for(var i=0;i<5;i++)p.push(from64(c.substr(2+i*4,4)));if(p.some(function(x){return x<0;}))return null;var h=p[0].toString(16).padStart(2,"0");for(var i=1;i<6;i++)h+=p[i].toString(16).padStart(6,"0");return h.substr(0,8)+"-"+h.substr(8,4)+"-"+h.substr(12,4)+"-"+h.substr(16,4)+"-"+h.substr(20,12);}
function detectFmt(g){if(!g)return"x";var s=String(g).trim();if(s.length===36&&/^[0-9a-f]{8}-/i.test(s))return"uuid";if(s.length===32&&/^[0-9a-f]{32}$/i.test(s))return"nd";if(s.length===22)return"ifc";return"x";}

/* ═══ Ledger persistence (per-view) ═══ */
function serializeColorLedger(){
  return Array.from(_colorLedger.entries()).map(function(entry){
    return {key:entry[0], value:entry[1]};
  });
}
function deserializeColorLedger(raw){
  var map=new Map();
  if(!raw) return map;
  var arr;
  if(typeof raw==="string"){
    try{arr=JSON.parse(raw);}catch(e){arr=null;}
  } else arr=raw;
  if(Array.isArray(arr)){
    arr.forEach(function(item){
      if(!item||!item.key) return;
      var value=item.value||{};
      map.set(item.key,{color:String(value.color||"").trim(), note:String(value.note||"").trim()});
    });
    return map;
  }
  if(typeof arr==="object" && arr!==null){
    Object.keys(arr).forEach(function(key){
      var value=arr[key]||{};
      if(typeof value==="object"){
        map.set(key,{color:String(value.color||"").trim(), note:String(value.note||"").trim()});
      }
    });
    return map;
  }
  return map;
}
function getLedgerKey(){return _ledgerViewId?"colorstudio_"+_ledgerViewId:null;}
function saveLedger(){var key=getLedgerKey();if(!key)return;try{localStorage.setItem(key,JSON.stringify(serializeColorLedger()));}catch(e){}}
async function loadLedger(viewId){
  _ledgerViewId=viewId;
  try{var raw=localStorage.getItem("colorstudio_"+viewId);_colorLedger=raw?deserializeColorLedger(raw):new Map();}catch(e){_colorLedger=new Map();}
  await rebuildColorLedgerSummary();
  await renderColorLedger();
  await repaintLedger();
}
async function repaintLedger(){
  try{
    var api=await getAPI();
    try{await api.viewer.setObjectState(undefined,{color:"reset",visible:"reset"});}catch(e){}
    var groups=buildColorGroups();
    for(var entry of groups){
      var color=entry[0], info=entry[1];
      for(var modelId of info.ids.keys()){
        var ids=info.ids.get(modelId);
        if(!Array.isArray(ids)||!ids.length) continue;
        await paintBatch(api,modelId,ids,{visible:true,color:color});
      }
    }
  }catch(e){}
}
async function initActiveView(api){
  try{
    var view=await api.viewer.getActiveView();
    var id=view&&(view.id||view.name);
    if(id){_currentViewId=String(id);loadLedger(_currentViewId);}
  }catch(e){}
  if(!_viewPollInterval){
    _viewPollInterval=setInterval(async function(){
      try{
        if(!_api)return;
        if(_loadingView)return;
        var view=await _api.viewer.getActiveView();
        var id=view&&(view.id||view.name);
        if(id&&String(id)!==_currentViewId){_currentViewId=String(id);loadLedger(_currentViewId);}
      }catch(e){}
    },5000);
  }
}

/* ═══ API ═══ */
async function getAPI(){if(_api)return _api;_api=await TrimbleConnectWorkspace.connect(window.parent,function(e,d){console.log("[T]",e,d);});log("Đã kết nối Trimble API.","ok");initActiveView(_api);return _api;}

/* ═══ SPM (.xsr) input ═══ */
function readText(f){return new Promise(function(ok,no){var r=new FileReader();r.onload=function(e){ok(String(e.target.result||""));};r.onerror=no;r.readAsText(f);});}
function parseSPM(text){
  var lines=String(text||"").split(/\r?\n/);
  if(lines.length)lines.shift();                 // bỏ dòng đầu (header "SPM2019: N")
  var set=new Set();
  lines.forEach(function(line){
    if(!line)return;
    var cols=line.trim().split(/\s+/);            // fixed-width → tách theo khoảng trắng liên tiếp
    if(cols[0]!=="ASS")return;                    // chỉ xử lý dòng loại ASS, bỏ qua PART
    // TODO: xác nhận cột GUID sau khi có kết quả snippet match
    //       (mặc định cols[1] = assembly GUID gốc Tekla; có thể đổi sang cols[2] = part GUID)
    var guid=cols[1];
    if(guid)set.add(guid);                         // dedupe bằng Set
  });
  return Array.from(set);
}

/* ═══ Model ═══ */
async function getModelIds(){
  var api=await getAPI();
  for(var a=1;a<=RETRY_MAX;a++){
    var raw;
    try{raw=await api.viewer.getObjects();}catch(e){if(a<RETRY_MAX){await sleep(RETRY_DELAY);continue;}throw e;}
    if(!Array.isArray(raw)||!raw.length){if(a<RETRY_MAX){await sleep(RETRY_DELAY);continue;}throw new Error("Viewer trống. Đợi model load.");}
    var total=0,mids=[];
    raw.forEach(function(g){
      if(!g||!g.modelId)return;
      if(mids.indexOf(g.modelId)===-1)mids.push(g.modelId);
      if(Array.isArray(g.objects))total+=g.objects.length;
      else if(Array.isArray(g.objectRuntimeIds))total+=g.objectRuntimeIds.length;
      else if(Array.isArray(g.ids))total+=g.ids.length;
    });
    if(!mids.length){if(a<RETRY_MAX){await sleep(RETRY_DELAY);continue;}throw new Error("Không thấy modelId.");}
    mids.forEach(function(id){log("  model "+id.substr(0,12)+"...","info");});
    return{modelIds:mids,total:total};
  }
}

/* ═══ Convert ═══ */
function flat(v){if(v==null)return[];if(typeof v==="number")return[v];if(Array.isArray(v)){var o=[];v.forEach(function(x){if(typeof x==="number")o.push(x);else if(Array.isArray(x))x.forEach(function(y){if(typeof y==="number")o.push(y);});});return o;}return[];}
async function batchConvert(api,mid,guids){var out=[];for(var i=0;i<guids.length;i+=BATCH_CVT){var c=guids.slice(i,i+BATCH_CVT);var r;try{r=await api.viewer.convertToObjectRuntimeIds(mid,c);}catch(e){for(var k=0;k<c.length;k++)out.push(null);continue;}if(!Array.isArray(r)){for(var k=0;k<c.length;k++)out.push(null);continue;}out=out.concat(r);}return out;}

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

/* ═══ MAIN (paint per slot) ═══ */
async function paintSlot(slot){
  var guids = slot===1 ? _guids1 : _guids2;
  var color = slot===1 ? _color1 : _color2;
  var btn = document.getElementById("applyBtn"+slot);
  if(!guids.length){log("✗ Chưa có file Excel #"+slot+".","warn");return;}
  btn.disabled=true;clearLog();setProgress(10);
  try{
    var api=await getAPI();
    var mi=await getModelIds();
    setStat("s-total",fmtN(mi.total));
    setProgress(30);
    log("Map GUIDs File #"+slot+"...","info");
    var map=await convertAll(api,mi.modelIds,guids,"File"+slot);
    var cnt=0;map.forEach(function(ids){cnt+=ids.length;});
    setStat("s-c"+slot,cnt>0?fmtN(cnt):"—");
    if(cnt===0){log("✗ Không match object nào cho File #"+slot+"!","err");setProgress(0);btn.disabled=false;return;}
    setProgress(60);
    log("Tô màu File #"+slot+" ("+color+"): "+fmtN(cnt)+" objects...","info");
    for(var i=0;i<mi.modelIds.length;i++){
      var mid=mi.modelIds[i];var ids=map.get(mid);
      if(!ids||!ids.length)continue;
      await paintBatch(api,mid,ids,{visible:true,color:color});
      log("  ▪ "+fmtN(ids.length)+" objects","ok");
    }
    if(slot===1){_map1=map;}else{_map2=map;}
    setBadge(slot,"ready");
    setProgress(100);
    log("✓ Hoàn tất File #"+slot+".","ok");
    setTimeout(function(){setProgress(0);},1500);
    document.getElementById("qtyBtn"+slot).disabled=false;
  }catch(err){
    log("✗ "+(err&&err.message?err.message:String(err)),"err");setProgress(0);
  }finally{btn.disabled=false;}
}

/* ═══ Quantities (Qto) ═══ */
var BATCH_PROP = 200;
var QTY_PATTERNS = {
  weight: /NetWeight|GrossWeight|^Weight$/i,
  length: /^Length$/i,
  profile: /^PROFILE$|profile|member.*size|section/i,
  area: /NetArea|GrossArea|^Area$/i
};
var ASSEMBLY_TIERS = [
  /^ASSEMBLY_POS$/i,
  /^assembly\s*position$/i,
  /assembly.*(mark|position|name)/i,
  /^mark$/i,
  /^position$/i,
  /drawing\s*no/i
];

/* Gom theo tên cấu kiện Assembly: { count, length, weight } cho mỗi nhóm */
async function fetchQuantities(api, map){
  var groups=new Map(), hit=0, miss=0;
  for(var entry of map){
    var mid=entry[0], ids=entry[1];
    for(var i=0;i<ids.length;i+=BATCH_PROP){
      var chunk=ids.slice(i,i+BATCH_PROP);
      var props;
      try{ props = await api.viewer.getObjectProperties(mid, chunk); }
      catch(e){ miss+=chunk.length; continue; }
      if(!Array.isArray(props)){ miss+=chunk.length; continue; }
      props.forEach(function(op){
        if(!window.__diagLogged){
          window.__diagLogged=true;
          console.log("=== DIAG: full properties of first object ===");
          if(Array.isArray(op.properties)){
            op.properties.forEach(function(ps){
              console.log("--- PropertySet:", ps.name||ps.pset||"(no name)");
              if(Array.isArray(ps.properties)){
                ps.properties.forEach(function(p){
                  console.log("   ", p.name, "=", p.value);
                });
              }
            });
          }
        }
        var weight=null, length=null, profile=null, area=null, found=false;
        var assemblyTierVals=new Array(ASSEMBLY_TIERS.length).fill(null);
        if(Array.isArray(op.properties)){
          op.properties.forEach(function(ps){
            if(!Array.isArray(ps.properties))return;
            ps.properties.forEach(function(p){
              for(var t=0;t<ASSEMBLY_TIERS.length;t++){
                if(assemblyTierVals[t]==null && ASSEMBLY_TIERS[t].test(p.name) && p.value!=null && String(p.value).trim()!==""){
                  assemblyTierVals[t]=String(p.value).trim();
                }
              }
              if(profile==null && QTY_PATTERNS.profile.test(p.name) && p.value!=null && String(p.value).trim()!==""){
                profile=String(p.value).trim();
              }
              var v=parseFloat(p.value);
              if(isNaN(v))return;
              if(weight==null && QTY_PATTERNS.weight.test(p.name)){weight=v;found=true;}
              if(length==null && QTY_PATTERNS.length.test(p.name)){length=v;found=true;}
              if(QTY_PATTERNS.area.test(p.name)){area=(area||0)+v;}
            });
          });
        }
        var assembly=assemblyTierVals.find(function(v){return v!=null;})||null;
        if(found)hit++; else miss++;
        var key=assembly||"Không xác định";
        var g=groups.get(key);
        if(!g){g={count:0,length:null,weight:0,profile:null,area:null};groups.set(key,g);}
        g.count+=1;
        if(g.length==null && length!=null)g.length=length;
        if(weight!=null)g.weight+=weight;
        if(g.profile==null && profile!=null)g.profile=profile;
        if(area!=null)g.area=(g.area||0)+area;
      });
    }
  }
  return {groups:groups, hit:hit, miss:miss};
}

function escapeHtml(s){return String(s).replace(/[&<>"]/g,function(c){return({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"})[c];});}

async function showQty(slot, map){
  var btn=document.getElementById("qtyBtn"+slot);
  var resEl=document.getElementById("qtyResult"+slot);
  var noteEl=document.getElementById("noteInput"+slot);
  if(!map||!map.size){log("✗ Slot "+slot+" chưa có dữ liệu để xem khối lượng.","warn");return;}
  btn.disabled=true;resEl.classList.remove("hidden");
  resEl.innerHTML='<div class="text-[11px] font-mono text-slate-400">Đang đọc khối lượng...</div>';
  try{
    var api=await getAPI();
    var sel=[],total=0;
    map.forEach(function(ids,mid){sel.push({modelId:mid,objectRuntimeIds:ids});total+=ids.length;});
    await api.viewer.setSelection({modelObjectIds:sel}, "set");
    var q=await fetchQuantities(api, map);
    _lastQty[slot]=q;
    if(!q.groups.size){
      resEl.innerHTML='<div class="text-[11px] font-mono text-slate-400">Model này không có dữ liệu Qto/Assembly để đọc.</div>';
    } else {
      var rows=Array.from(q.groups.entries()).sort(function(a,b){return a[0].localeCompare(b[0]);});
      var totalCount=0, totalWeight=0;
      var html='<table class="w-full text-[10.5px] font-mono border-collapse">'
        +'<thead><tr style="background:#f0f2f5">'
        +'<th class="text-left py-1.5 px-2" style="border:1px solid #d0d3d8;color:#5f6368;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.05em">Tên</th>'
        +'<th class="text-right py-1.5 px-2" style="border:1px solid #d0d3d8;color:#5f6368;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.05em">SL</th>'
        +'<th class="text-right py-1.5 px-2" style="border:1px solid #d0d3d8;color:#5f6368;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.05em">Dài</th>'
        +'<th class="text-right py-1.5 px-2" style="border:1px solid #d0d3d8;color:#5f6368;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.05em">KL (tấn)</th>'
        +'</tr></thead><tbody>';
      rows.forEach(function(r){
        var name=r[0], g=r[1];
        totalCount+=g.count; totalWeight+=g.weight;
        html+='<tr>'
          +'<td class="py-1 px-2" style="border:1px solid #e2e5ea;color:#1a1c1e">'+escapeHtml(name)+'</td>'
          +'<td class="text-right py-1 px-2" style="border:1px solid #e2e5ea;color:#1a1c1e">'+fmtN(g.count)+'</td>'
          +'<td class="text-right py-1 px-2" style="border:1px solid #e2e5ea;color:#1a1c1e">'+(g.length!=null?g.length.toLocaleString(undefined,{maximumFractionDigits:1}):'—')+'</td>'
          +'<td class="text-right py-1 px-2" style="border:1px solid #e2e5ea;color:#1a1c1e">'+(g.weight/1000).toLocaleString(undefined,{maximumFractionDigits:3})+'</td>'
          +'</tr>';
      });
      html+='</tbody><tfoot><tr style="background:#e8f0fe">'
        +'<td class="py-1.5 px-2" style="border:1px solid #d0d3d8;color:#1a73e8;font-weight:700">TỔNG</td>'
        +'<td class="text-right py-1.5 px-2" style="border:1px solid #d0d3d8;color:#1a73e8;font-weight:700">'+fmtN(totalCount)+'</td>'
        +'<td class="text-right py-1.5 px-2" style="border:1px solid #d0d3d8;color:#1a73e8;font-weight:700">—</td>'
        +'<td class="text-right py-1.5 px-2" style="border:1px solid #d0d3d8;color:#1a73e8;font-weight:700">'+(totalWeight/1000).toLocaleString(undefined,{maximumFractionDigits:3})+'</td>'
        +'</tr></tfoot></table>'
        +'<div class="text-[10px] text-slate-500 mt-1">'+q.hit+' object có dữ liệu, '+q.miss+' object không có.</div>'
        +'<button onclick="exportMTO('+slot+')" class="w-full mt-2 px-1.5 py-2 rounded-lg font-sans text-xs font-semibold cursor-pointer transition-all duration-150 active:scale-95" style="background:#e8f0fe;border:1px solid #c5d5f7;color:#1a73e8">⬇ Xuất khối lượng (MTO)</button>';
      resEl.innerHTML=html;
    }
    if(noteEl){noteEl.classList.remove("hidden");var saved=localStorage.getItem("mcc_note_slot"+slot);if(saved!=null)noteEl.value=saved;}
    log("✓ Slot "+slot+": đã sáng "+fmtN(total)+" object trên model.","ok");
  }catch(e){
    resEl.innerHTML='<div class="text-[11px] font-mono text-sys-red">✗ '+escapeHtml(e&&e.message?e.message:String(e))+'</div>';
    log("✗ "+(e&&e.message?e.message:String(e)),"err");
  }finally{btn.disabled=false;}
}

/* ═══ Reset ═══ */
async function resetViewer(){
  lockUI(true);clearLog();setProgress(10);
  try{var api=await getAPI();try{await api.viewer.setObjectState(undefined,{color:"reset",visible:"reset"});}catch(e){}await api.viewer.reset();
  setStat("s-total","—");setStat("s-c1","—");setStat("s-c2","—");
  _map1=null;_map2=null;_selMap=null;
  _colorLedger=new Map();_ledgerViewId=null;_colorLedgerSummary=null;await renderColorLedger();
  [1,2].forEach(function(n){
    document.getElementById("qtyBtn"+n).disabled=true;
    document.getElementById("qtyResult"+n).classList.add("hidden");
    document.getElementById("noteInput"+n).classList.add("hidden");
  });
  document.getElementById("qtyBtn3").disabled=true;
  document.getElementById("qtyResult3").classList.add("hidden");
  // Option 4: đồng bộ highlight tạm sau khi reset toàn bộ màu (giữ nguyên cache)
  _apSelected.clear();
  var apL=document.getElementById("apList");
  if(apL){Array.prototype.forEach.call(apL.querySelectorAll(".ap-check"),function(cb){cb.checked=false;});}
  setProgress(100);log("✓ Reset OK.","ok");setTimeout(function(){setProgress(0);},1000);}
  catch(e){log("✗ "+(e&&e.message?e.message:String(e)),"err");setProgress(0);}
  finally{lockUI(false);checkApplyBtn(1);checkApplyBtn(2);}
}

/* ═══ Viewer Selection ═══ */
async function captureSelection(){
  var btn=document.getElementById("captureSelBtn");
  var info=document.getElementById("selInfo");
  btn.disabled=true;
  try{
    var api=await getAPI();
    var raw=await api.viewer.getSelection();
    var map=new Map(),total=0;
    if(Array.isArray(raw)){
      raw.forEach(function(g){
        if(!g||!g.modelId)return;
        var ids=flat(g.objects||g.objectRuntimeIds||g.ids);
        if(!ids.length)return;
        map.set(g.modelId,ids);
        total+=ids.length;
      });
    }
    if(!total){
      _selMap=null;
      info.textContent='Chưa chọn gì — click cấu kiện trên viewer rồi bấm "Lấy lựa chọn"';
      document.getElementById("paintSelBtn").disabled=true;
      log("✗ Chưa chọn cấu kiện nào trên viewer.","warn");
      return;
    }
    _selMap=map;
    info.textContent="Đã chọn: "+fmtN(total)+" cấu kiện";
    document.getElementById("paintSelBtn").disabled=false;
    document.getElementById("qtyBtn3").disabled=false;
    log("✓ Lấy lựa chọn: "+fmtN(total)+" cấu kiện","ok");
  }catch(e){
    log("✗ "+(e&&e.message?e.message:String(e)),"err");
  }finally{
    btn.disabled=false;
  }
}

async function paintSelection(){
  if(!_selMap||!_selMap.size){log("✗ Chưa có lựa chọn nào.","warn");return;}
  var btn=document.getElementById("paintSelBtn");
  btn.disabled=true;
  try{
    var api=await getAPI();
    log("Tô màu lựa chọn ("+_color3+")...","info");
    var done=0;
    var note=(document.getElementById("noteInput")||{}).value;
    note=note?note.trim():"";
    for(var entry of _selMap){
      await paintBatch(api,entry[0],entry[1],{color:_color3});
      done+=entry[1].length;
      entry[1].forEach(function(rid){
        var key=makeObjectKey(entry[0],rid);
        _colorLedger.set(key,{color:_color3,note:note});
      });
    }
    log("✓ Đã tô "+fmtN(done)+" cấu kiện.","ok");
    document.getElementById("qtyBtn3").disabled=false;
    await rebuildColorLedgerSummary();
    await renderColorLedger();
    saveLedger();
  }catch(e){
    log("✗ "+(e&&e.message?e.message:String(e)),"err");
  }finally{
    btn.disabled=false;
  }
}

/* ═══ Markup ghi chú (auto) ═══ */
function hex2rgba(hex){var h=hex.replace("#","");return{r:parseInt(h.substr(0,2),16),g:parseInt(h.substr(2,2),16),b:parseInt(h.substr(4,2),16),a:255};}
function mergeBoundingBox(existing, bbox){
  if(!bbox||bbox.min==null||bbox.max==null)return existing;
  if(!existing)return {min:{x:bbox.min.x,y:bbox.min.y,z:bbox.min.z},max:{x:bbox.max.x,y:bbox.max.y,z:bbox.max.z}};
  existing.min.x = Math.min(existing.min.x,bbox.min.x);
  existing.min.y = Math.min(existing.min.y,bbox.min.y);
  existing.min.z = Math.min(existing.min.z,bbox.min.z);
  existing.max.x = Math.max(existing.max.x,bbox.max.x);
  existing.max.y = Math.max(existing.max.y,bbox.max.y);
  existing.max.z = Math.max(existing.max.z,bbox.max.z);
  return existing;
}
function getMarkupColor(slot){return slot===3?_color3:MARKUP_COLOR;}
async function getMapBoundingBox(api,map){
  if(!map||!map.size)return null;
  var merged=null;
  for(var entry of map){
    var modelId=entry[0], ids=entry[1];
    if(!Array.isArray(ids)||!ids.length)continue;
    var boxes=null;
    if(typeof api.viewer.getObjectBoundingBoxes==="function"){
      try{boxes=await api.viewer.getObjectBoundingBoxes(modelId,ids);}catch(e){boxes=null;}
    }
    if((!boxes||!boxes.length) && typeof api.viewer.getObjectBoundingBox==="function"){
      try{var single=await api.viewer.getObjectBoundingBox({modelId:modelId,objectRuntimeIds:ids});
        if(single){boxes=Array.isArray(single)?single:[single];}
      }catch(e){boxes=null;}
    }
    if(!boxes||!boxes.length){
      try{
        var raw=await api.viewer.getSelection();
        if(Array.isArray(raw)){
          var selIds=[];
          raw.forEach(function(item){
            if(!item||item.modelId!==modelId)return;
            var found=(item.objects||item.objectRuntimeIds||item.ids);
            if(Array.isArray(found)) selIds = selIds.concat(found);
          });
          if(selIds.length && typeof api.viewer.getObjectBoundingBoxes==="function"){
            try{boxes=await api.viewer.getObjectBoundingBoxes(modelId,selIds);}catch(e){boxes=null;}
          }
        }
      }catch(e){boxes=null;}
    }
    if(Array.isArray(boxes)){
      boxes.forEach(function(item){
        if(!item) return;
        var bbox = item.boundingBox || item;
        if(bbox && bbox.min!=null && bbox.max!=null){
          merged = mergeBoundingBox(merged,bbox);
        }
      });
    }
  }
  return merged;
}
function getFirstAnchor(map){if(!map||!map.size)return null;for(var entry of map){if(entry[1]&&entry[1].length)return{modelId:entry[0],id:entry[1][0]};}return null;}

async function addNoteMarkup(slot){
  var map = slot===1?_map1:slot===2?_map2:_selMap;
  var noteEl = document.getElementById("noteInput"+slot);
  var text = noteEl.value.trim();
  try{
    var api=await getAPI();
    if(!text){
      if(_noteMarkupId[slot]!=null){await api.markup.removeMarkups([_noteMarkupId[slot]]);_noteMarkupId[slot]=null;}
      return;
    }
    var bb = await getMapBoundingBox(api,map);
    if(!bb||bb.min==null||bb.max==null){
      log("✗ Không lấy được vị trí vùng chọn để gắn markup.","warn");
      return;
    }
    var cx=(bb.min.x+bb.max.x)/2, cy=(bb.min.y+bb.max.y)/2, cz=(bb.min.z+bb.max.z)/2;
    var offset = Math.max(Math.abs(bb.max.x-bb.min.x), Math.abs(bb.max.y-bb.min.y), Math.abs(bb.max.z-bb.min.z))/2 || 1000;
    var m={
      start:{positionX:cx,positionY:cy,positionZ:cz},
      end:{positionX:cx+offset,positionY:cy+offset,positionZ:cz+offset},
      text:text,
      color:hex2rgba(getMarkupColor(slot))
    };
    if(_noteMarkupId[slot]!=null)m.id=_noteMarkupId[slot];
    var added=await api.markup.addTextMarkup([m]);
    if(added&&added[0]&&added[0].id!=null)_noteMarkupId[slot]=added[0].id;
    log("✓ Đã gắn markup ghi chú (Slot "+slot+").","ok");
  }catch(e){
    log("✗ "+(e&&e.message?e.message:String(e)),"err");
  }
}

/* ═══ Color Ledger ═══ */
async function renderColorLedger(){
  var el=document.getElementById("colorLedger");
  var expBtn=document.getElementById("exportLedgerBtn");
  if(!el)return;
  var groups=buildColorGroups();
  if(!groups.size){el.classList.add("hidden");if(expBtn)expBtn.classList.add("hidden");return;}
  await rebuildColorLedgerSummary();
  var colors=Array.from(groups.keys()).sort();
  var html='<table style="width:100%;border-collapse:collapse">'
    +'<thead><tr style="background:#f0f2f5;position:sticky;top:0;z-index:2">'
    +'<th style="text-align:left;padding:6px 8px;border:1px solid #d0d3d8;font-size:10px;color:#5f6368;font-weight:700;text-transform:uppercase;letter-spacing:.05em;white-space:nowrap">STT</th>'
    +'<th style="text-align:left;padding:6px 8px;border:1px solid #d0d3d8;font-size:10px;color:#5f6368;font-weight:700;text-transform:uppercase;letter-spacing:.05em;white-space:nowrap">Màu</th>'
    +'<th style="text-align:right;padding:6px 8px;border:1px solid #d0d3d8;font-size:10px;color:#5f6368;font-weight:700;text-transform:uppercase;letter-spacing:.05em;white-space:nowrap">SL</th>'
    +'<th style="text-align:right;padding:6px 8px;border:1px solid #d0d3d8;font-size:10px;color:#5f6368;font-weight:700;text-transform:uppercase;letter-spacing:.05em;white-space:nowrap">KL (Tấn)</th>'
    +'<th style="text-align:left;padding:6px 8px;border:1px solid #d0d3d8;font-size:10px;color:#5f6368;font-weight:700;text-transform:uppercase;letter-spacing:.05em">Ghi chú</th>'
    +'</tr></thead><tbody>';
  colors.forEach(function(color,idx){
    var info=groups.get(color);
    var summary=_colorLedgerSummary&&_colorLedgerSummary.get(color);
    var count=info?Array.from(info.ids.values()).reduce(function(acc,ids){return acc+(Array.isArray(ids)?ids.length:0);},0):0;
    var weight=summary?summary.weight:0;
    var note=summary?summary.note:(info?info.note||"":"");
    var bg=idx%2===0?"#ffffff":"#f7f8fa";
    html+='<tr style="background:'+bg+'">'
      +'<td style="padding:5px 8px;border:1px solid #e2e5ea;font-size:10.5px;color:#1a1c1e">'+(idx+1)+'</td>'
      +'<td style="padding:5px 8px;border:1px solid #e2e5ea">'
      +'<div style="width:16px;height:16px;border-radius:3px;background:'+color+';border:1px solid rgba(0,0,0,0.12)"></div></td>'
      +'<td style="padding:5px 8px;border:1px solid #e2e5ea;text-align:right;font-family:\'JetBrains Mono\',monospace;font-size:10.5px;color:#1a1c1e">'+fmtN(count)+'</td>'
      +'<td style="padding:5px 8px;border:1px solid #e2e5ea;text-align:right;font-family:\'JetBrains Mono\',monospace;font-size:10.5px;color:#1a1c1e">'+(weight/1000).toLocaleString(undefined,{minimumFractionDigits:3,maximumFractionDigits:3})+'</td>'
      +'<td style="padding:5px 8px;border:1px solid #e2e5ea"><input type="text" class="ledger-note-input" data-color="'+color+'" value="'+escapeHtml(note)+'" placeholder="Ghi chú..." onblur="updateColorNoteForColor(this.dataset.color,this.value.trim())" onkeydown="if(event.key===\'Enter\')this.blur()"/></td>'
      +'</tr>';
  });
  html+='</tbody></table>';
  el.innerHTML=html;
  el.classList.remove("hidden");
  if(expBtn)expBtn.classList.remove("hidden");
}

async function exportLedger(){
  await rebuildColorLedgerSummary();
  var groups=buildColorGroups();
  if(!groups.size){log("✗ Bảng kê trống.","warn");return;}
  var data=[["STT","Màu","SL","KL (Tấn)","Ghi chú"]];
  var colors=Array.from(groups.keys()).sort();
  colors.forEach(function(color,idx){
    var info=groups.get(color);
    var count=info?Array.from(info.ids.values()).reduce(function(acc,ids){return acc+(Array.isArray(ids)?ids.length:0);},0):0;
    var summary=_colorLedgerSummary&&_colorLedgerSummary.get(color);
    var weight=summary?summary.weight:0;
    var note=summary?summary.note:(info?info.note||"":"");
    data.push([idx+1,color,count,parseFloat((weight/1000).toFixed(3)),note]);
  });
  var ws=XLSX.utils.aoa_to_sheet(data);
  var wb2=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2,ws,"Bang ke");
  var d=new Date();
  XLSX.writeFile(wb2,"bang-ke-"+d.getFullYear()+"-"+pad2(d.getMonth()+1)+"-"+pad2(d.getDate())+".xlsx");
  log("✓ Xuất bảng kê thành công.","ok");
}

function round2(n){return Math.round(n*100)/100;}

async function exportMTO(slot){
  var api=await getAPI();
  var data;
  if(slot===3){
    var map = buildLedgerMapAll();
    if(!map||!map.size){log("✗ Chưa có dữ liệu khối lượng (Slot 3 chưa tô màu).","warn");return;}
    data = await fetchQuantities(api, map);
  } else {
    data=_lastQty&&_lastQty[slot];
  }
  if(!data||!data.groups||!data.groups.size){log("✗ Chưa có dữ liệu khối lượng (bấm Xem khối lượng trước).","warn");return;}
  var now=new Date();
  var dt=pad2(now.getDate())+"."+pad2(now.getMonth()+1)+"."+now.getFullYear()+" / "+pad2(now.getHours())+":"+pad2(now.getMinutes())+":"+pad2(now.getSeconds());
  var aoa=[];
  aoa.push(["ASSEMBLY LIST (MTO)"]);
  aoa.push(["Project number:","","DDC-0033"]);
  aoa.push(["Project name:","","PHU QUOC INTERNATIONAL AIRPORT"]);
  aoa.push(["Builder:","",""]);
  aoa.push(["Object:","",""]);
  aoa.push(["Rev.:"]);
  aoa.push(["Date/time of issue","",dt]);
  aoa.push([]);
  aoa.push(["NO.","Drawing No.","MEMBER SIZE","QTY","LENGTH","Area(m2)","T/Area(m2)","Weight(kg)","T/Weight(kg)","GRID LOCATION","FIRE PROOFED"]);
  var rows=Array.from(data.groups.entries()).sort(function(a,b){return String(a[0]).localeCompare(String(b[0]));});
  var no=0,totalWeight=0;
  rows.forEach(function(r){
    var name=r[0],g=r[1];
    no++;
    var qty=g.count||0;
    var wTot=g.weight||0;
    var wPer=qty?wTot/qty:wTot;
    var aTot=(g.area!=null)?g.area:null;
    var aPer=(aTot!=null&&qty)?aTot/qty:aTot;
    totalWeight+=wTot;
    aoa.push([no,name,g.profile||"",qty,g.length!=null?round2(g.length):"",aPer!=null?round2(aPer):"",aTot!=null?round2(aTot):"",round2(wPer),round2(wTot),"",""]);
  });
  aoa.push(["","","","","","","","Total",round2(totalWeight),"",""]);
  var ws=XLSX.utils.aoa_to_sheet(aoa);
  var wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"MTO");
  var fn="MTO_"+now.getFullYear()+pad2(now.getMonth()+1)+pad2(now.getDate())+"_"+pad2(now.getHours())+pad2(now.getMinutes())+".xlsx";
  XLSX.writeFile(wb,fn);
  log('✓ Đã xuất "'+fn+'".',"ok");
}

/* ═══ Save View ═══ */
async function saveView(){
  try{
    var api=await getAPI();
    var inp=document.getElementById("viewName");
    var name=inp?inp.value.trim():"";
    if(!name){
      var n=new Date();
      name="ColorStudio "+n.getFullYear()+"-"+pad2(n.getMonth()+1)+"-"+pad2(n.getDate())+" "+pad2(n.getHours())+":"+pad2(n.getMinutes());
      if(inp)inp.value=name;
    }
    var c=await api.view.createView({name:name,description:"Model Control Center v2.0 | Le Van Thao"});
    if(!c||!c.id)throw new Error("No view ID.");
    await api.view.updateView({id:c.id});
    await api.view.selectView(c.id);

    // Lưu colorLedger vào view mới
    _currentViewId=String(c.id);
    _ledgerViewId=_currentViewId;
    saveLedger();

    // Lưu state vào localStorage
    var state={
      name:name,
      color1:_color1,
      color2:_color2,
      color3:_color3,
      guids1:_guids1||[],
      guids2:_guids2||[],
      note1:document.getElementById("noteInput1").value||"",
      note2:document.getElementById("noteInput2").value||"",
      note3:(document.getElementById("noteInput")||{}).value||"",
      selMapSerialized:_selMap?Array.from(_selMap.entries()).map(function(e){return{modelId:e[0],ids:e[1]};}):[],
      colorLedger:serializeColorLedger()
    };
    localStorage.setItem("mcc_view_"+c.id,JSON.stringify(state));

    // Cập nhật danh sách view
    var viewList=JSON.parse(localStorage.getItem("mcc_view_list")||"[]");
    viewList.unshift({id:c.id,name:name,ts:Date.now()});
    viewList=viewList.slice(0,20);
    localStorage.setItem("mcc_view_list",JSON.stringify(viewList));

    // Reset canvas + bảng kê sau khi lưu
    _colorLedger=new Map();
    _colorLedgerSummary=null;
    _ledgerViewId=null;
    await renderColorLedger();
    try{await api.viewer.setObjectState(undefined,{color:"reset",visible:"reset"});}catch(e3){}

    log('✓ View đã lưu: "'+name+'" — Canvas đã reset, sẵn sàng tô cho view kế tiếp.',"ok");
    renderViewList();
  }catch(e){log("✗ "+(e&&e.message?e.message:String(e)),"err");}
}

async function loadView(viewId){
  var raw=localStorage.getItem("mcc_view_"+viewId);
  if(!raw){log("✗ Không tìm thấy dữ liệu view này.","warn");return;}
  var state;
  try{state=JSON.parse(raw);}catch(e){log("✗ Dữ liệu view bị lỗi.","err");return;}

  log('Đang load view "'+state.name+'"...',"info");

  // Reset sạch màu viewer và ledger trước khi tô view mới (tránh màu cộng dồn từ view trước)
  try{var api=await getAPI();await api.viewer.setObjectState(undefined,{color:"reset"});}catch(e){}
  await sleep(300);
  _colorLedger=new Map();
  _colorLedgerSummary=null;
  await renderColorLedger();

  // Khoá poll active-view trong lúc load để repaintLedger không tô đè/nạp ledger cũ,
  // rồi rebuild _colorLedger từ đúng màu thực tế sắp tô (1 nguồn duy nhất, tránh cộng KL trùng)
  _loadingView=true;
  _currentViewId=String(viewId);
  _ledgerViewId=String(viewId);

  // Khôi phục màu
  if(state.color1)setColor(1,state.color1);
  if(state.color2)setColor(2,state.color2);
  if(state.color3)setColor(3,state.color3);

  // Khôi phục ghi chú
  var noteIds={1:"noteInput1",2:"noteInput2",3:"noteInput"};
  [1,2,3].forEach(function(n){
    var el=document.getElementById(noteIds[n]);
    if(el&&state["note"+n]!=null)el.value=state["note"+n];
  });

  // Chuyển viewer sang view đó
  try{
    var api=await getAPI();
    await api.view.selectView(viewId);
    // Đợi viewer load xong view mới rồi mới tô màu
    await sleep(2000);
  }catch(e){log("⚠ Không chuyển được view trên viewer: "+(e&&e.message?e.message:String(e)),"warn");}

  // Khôi phục guids và tô màu lại
  if(state.guids1&&state.guids1.length){
    _guids1=state.guids1;
    checkApplyBtn(1);
    log("Tô lại màu Slot 1 ("+state.guids1.length+" GUID)...","info");
    await paintSlot(1);
  }
  if(state.guids2&&state.guids2.length){
    _guids2=state.guids2;
    checkApplyBtn(2);
    log("Tô lại màu Slot 2 ("+state.guids2.length+" GUID)...","info");
    await paintSlot(2);
  }
  if(state.colorLedger&&Array.isArray(state.colorLedger)){
    _colorLedger=deserializeColorLedger(state.colorLedger);
    await repaintLedger();
    await rebuildColorLedgerSummary();
    await renderColorLedger();
  }
  if(state.selMapSerialized&&state.selMapSerialized.length){
    _selMap=new Map();
    state.selMapSerialized.forEach(function(e){_selMap.set(e.modelId,e.ids);});
    var total=0;state.selMapSerialized.forEach(function(e){total+=e.ids.length;});
    document.getElementById("selInfo").textContent="Đã chọn: "+fmtN(total)+" cấu kiện";
    document.getElementById("paintSelBtn").disabled=false;
  }

  // Ledger giờ đã khớp đúng màu thực tế trên viewer → lưu lại & mở poll trở lại
  _loadingView=false;
  renderColorLedger();
  saveLedger();
  log('✓ Load view "'+state.name+'" hoàn tất.',"ok");
}

function renderViewList(){
  var container=document.getElementById("viewListContainer");
  if(!container)return;
  var viewList=JSON.parse(localStorage.getItem("mcc_view_list")||"[]");
  if(!viewList.length){
    container.innerHTML='<div class="text-[10.5px] text-slate-400 italic">Chưa có view nào được lưu.</div>';
    return;
  }
  container.innerHTML=viewList.map(function(v){
    var d=new Date(v.ts);
    var ts=d.getFullYear()+"-"+pad2(d.getMonth()+1)+"-"+pad2(d.getDate())+" "+pad2(d.getHours())+":"+pad2(d.getMinutes());
    return '<div class="flex items-center gap-1.5 py-1 border-b border-slate-100 last:border-0">'
      +'<button onclick="loadView(\''+v.id+'\')" class="flex-1 text-left text-[10.5px] font-mono text-blue-600 hover:underline truncate" title="'+escapeHtml(v.name)+'">📂 '+escapeHtml(v.name)+'</button>'
      +'<span class="text-[9px] text-slate-400 whitespace-nowrap">'+ts+'</span>'
      +'<button onclick="deleteViewRecord(\''+v.id+'\')" class="text-[10px] text-slate-300 hover:text-red-400 ml-1" title="Xoá">✕</button>'
      +'</div>';
  }).join("");
}

function deleteViewRecord(viewId){
  localStorage.removeItem("mcc_view_"+viewId);
  var viewList=JSON.parse(localStorage.getItem("mcc_view_list")||"[]");
  viewList=viewList.filter(function(v){return v.id!==viewId;});
  localStorage.setItem("mcc_view_list",JSON.stringify(viewList));
  log("✓ Đã xoá view khỏi danh sách.","ok");
  renderViewList();
}

/* ═══ Option 4 — Lọc theo Assembly Position (Prefix) ═══ */
/* Trích prefix thật từ Assembly Position kiểu "Z213-TS59" → "TS" (chữ ngay trước số cuối) */
function extractAssemblyPrefix(pos){
  if(!pos) return null;
  var m=String(pos).match(/([A-Za-z]+)\d+$/);
  return m?m[1]:null;
}
/* Đọc Assembly Position của 1 object (tái dùng ASSEMBLY_TIERS — KHÔNG đụng fetchQuantities) */
function apExtractAssembly(op){
  var tierVals=new Array(ASSEMBLY_TIERS.length).fill(null);
  if(op && Array.isArray(op.properties)){
    op.properties.forEach(function(ps){
      if(!Array.isArray(ps.properties))return;
      ps.properties.forEach(function(p){
        for(var t=0;t<ASSEMBLY_TIERS.length;t++){
          if(tierVals[t]==null && ASSEMBLY_TIERS[t].test(p.name) && p.value!=null && String(p.value).trim()!==""){
            tierVals[t]=String(p.value).trim();
          }
        }
      });
    });
  }
  return tierVals.find(function(v){return v!=null;})||null;
}
/* Tên hiển thị: ưu tiên op.name → property "Name" → IFC type/class */
function apExtractName(op){
  if(!op)return"";
  if(op.name && String(op.name).trim())return String(op.name).trim();
  if(Array.isArray(op.properties)){
    var found=null;
    op.properties.forEach(function(ps){
      if(!Array.isArray(ps.properties))return;
      ps.properties.forEach(function(p){
        if(found==null && /^name$/i.test(p.name) && p.value!=null && String(p.value).trim()!==""){found=String(p.value).trim();}
      });
    });
    if(found)return found;
  }
  if(op.type && String(op.type).trim())return String(op.type).trim();
  if(op.class && String(op.class).trim())return String(op.class).trim();
  return"";
}
/* Riêng cho Option 4: raw[0].objects trả về [{id:N},...] chứ không phải number thuần như flat() gốc kỳ vọng */
function apFlatIds(v){
  if(!Array.isArray(v))return [];
  var out=[];
  v.forEach(function(x){
    if(typeof x==="number")out.push(x);
    else if(x && typeof x==="object" && typeof x.id==="number")out.push(x.id);
  });
  return out;
}
function apSetProgress(p){
  var w=document.getElementById("apProgWrap"),b=document.getElementById("apProgBar");
  if(!w||!b)return;
  if(p<=0){w.classList.add("hidden");b.style.width="0%";return;}
  w.classList.remove("hidden");b.style.width=Math.min(p,100)+"%";
}
function apSetStatus(t){var e=document.getElementById("apStatus");if(e)e.textContent=t||"";}
/* Quét TOÀN BỘ object 1 lần, build cache RAM (key giống _colorLedger) */
async function apBuildCache(){
  if(_apScanning)return;
  _apScanning=true;
  var btn=document.getElementById("apSearchBtn");
  if(btn)btn.disabled=true;
  apSetProgress(1);
  apSetStatus("Đang quét toàn bộ model lần đầu...");
  try{
    var api=await getAPI();
    var raw=null;
    for(var a=1;a<=RETRY_MAX;a++){
      try{raw=await api.viewer.getObjects();}catch(e){raw=null;}
      if(Array.isArray(raw)&&raw.length)break;
      if(a<RETRY_MAX)await sleep(RETRY_DELAY);
    }
    if(!Array.isArray(raw)||!raw.length)throw new Error("Viewer trống. Đợi model load.");
    var models=[],total=0;
    raw.forEach(function(g){
      if(!g||!g.modelId)return;
      var ids=apFlatIds(g.objects||g.objectRuntimeIds||g.ids);
      if(!ids.length)return;
      models.push({modelId:g.modelId,ids:ids});
      total+=ids.length;
    });
    if(!total)throw new Error("Không thấy object nào trong model.");
    var cache=new Map(),done=0;
    for(var mi=0;mi<models.length;mi++){
      var mid=models[mi].modelId,ids=models[mi].ids;
      for(var i=0;i<ids.length;i+=BATCH_PROP){
        var chunk=ids.slice(i,i+BATCH_PROP);
        var props=null;
        try{props=await api.viewer.getObjectProperties(mid,chunk);}catch(e){props=null;}
        for(var j=0;j<chunk.length;j++){
          var op=Array.isArray(props)?props[j]:null;
          var rid=(op&&op.id!=null)?op.id:chunk[j];
          var ap=apExtractAssembly(op);
          var nm=apExtractName(op);
          cache.set(makeObjectKey(mid,rid),{modelId:mid,runtimeId:rid,ap:ap||"",apPrefix:extractAssemblyPrefix(ap),name:nm||""});
        }
        done+=chunk.length;
        apSetProgress(total?Math.round(done/total*100):100);
      }
    }
    _apCache=cache;
    apSetStatus("Đã quét "+fmtN(cache.size)+" cấu kiện. Lọc tức thời từ giờ.");
    log("✓ Option 4: cache "+fmtN(cache.size)+" cấu kiện.","ok");
  }catch(e){
    _apCache=null;
    apSetStatus("✗ "+(e&&e.message?e.message:String(e)));
    log("✗ Option 4: "+(e&&e.message?e.message:String(e)),"err");
  }finally{
    _apScanning=false;
    if(btn)btn.disabled=false;
    setTimeout(function(){apSetProgress(0);},800);
  }
}
async function apSearch(){
  if(_apScanning)return;
  if(!_apCache){await apBuildCache();if(!_apCache)return;}
  apFilter();
}
/* Lọc trong cache RAM — tức thời, không gọi lại getObjectProperties */
function apFilter(){
  var input=document.getElementById("apPrefix");
  var prefix=input?input.value.replace(/^\s+|\s+$/g,""):"";
  var countEl=document.getElementById("apCount");
  if(!_apCache){apRenderResults([]);if(countEl)countEl.textContent="";return;}
  if(!prefix){apRenderResults([]);if(countEl)countEl.textContent="";apSetStatus("Nhập prefix rồi bấm Tìm.");return;}
  var matches=[];
  _apCache.forEach(function(v,key){
    if(v.apPrefix!=null && v.apPrefix===prefix){matches.push({key:key,ap:v.ap,name:v.name});} // so khớp tuyệt đối prefix đã trích, phân biệt hoa thường
  });
  matches.sort(function(a,b){var x=a.ap.localeCompare(b.ap);return x!==0?x:String(a.name).localeCompare(String(b.name));});
  if(countEl)countEl.textContent=fmtN(matches.length)+" kết quả";
  apRenderResults(matches);
}
function apRenderResults(matches){
  var listEl=document.getElementById("apList");
  if(!listEl)return;
  if(!matches.length){
    listEl.innerHTML='<div class="text-[10.5px] text-slate-400 italic px-2 py-3">Không có cấu kiện khớp prefix (phân biệt hoa thường).</div>';
    return;
  }
  listEl.innerHTML=matches.map(function(m){
    var checked=_apSelected.has(m.key)?"checked":"";
    return '<label class="ap-row flex items-center gap-2 px-2 py-1.5 cursor-pointer" style="border-bottom:1px solid #f0f2f5">'
      +'<input type="checkbox" class="ap-check" data-key="'+escapeHtml(m.key)+'" '+checked+' style="width:14px;height:14px;flex-shrink:0;accent-color:#9334e6;cursor:pointer"/>'
      +'<span class="font-mono text-[10.5px] font-semibold flex-shrink-0" style="color:#1a1c1e">'+escapeHtml(m.ap||"—")+'</span>'
      +'<span class="text-[10px] truncate" style="color:#5f6368">'+escapeHtml(m.name||"")+'</span>'
      +'</label>';
  }).join("");
  Array.prototype.forEach.call(listEl.querySelectorAll(".ap-check"),function(cb){
    cb.addEventListener("change",function(){apToggle(this.dataset.key,this.checked);});
  });
}
/* Highlight tạm thời 1 object — KHÔNG ghi _colorLedger, KHÔNG lưu, KHÔNG tính MTO */
async function apToggle(key,isChecked){
  var entry=_apCache&&_apCache.get(key);
  if(!entry)return;
  try{
    var api=await getAPI();
    if(isChecked){
      _apSelected.add(key);
      await api.viewer.setObjectState({modelObjectIds:[{modelId:entry.modelId,objectRuntimeIds:[Number(entry.runtimeId)]}]},{color:AP_HIGHLIGHT});
    }else{
      _apSelected.delete(key);
      await api.viewer.setObjectState({modelObjectIds:[{modelId:entry.modelId,objectRuntimeIds:[Number(entry.runtimeId)]}]},{color:"reset"});
    }
  }catch(e){log("✗ "+(e&&e.message?e.message:String(e)),"err");}
}
/* Bỏ chọn tất cả → trả màu gốc model cho mọi object Option 4 đang highlight */
async function apClearAll(){
  var listEl=document.getElementById("apList");
  if(listEl){Array.prototype.forEach.call(listEl.querySelectorAll(".ap-check"),function(cb){cb.checked=false;});}
  if(!_apSelected.size)return;
  var byModel=new Map();
  _apSelected.forEach(function(key){
    var o=splitObjectKey(key);
    var arr=byModel.get(o.modelId);if(!arr){arr=[];byModel.set(o.modelId,arr);}
    arr.push(Number(o.runtimeId));
  });
  _apSelected.clear();
  try{
    var api=await getAPI();
    for(var entry of byModel){await paintBatch(api,entry[0],entry[1],{color:"reset"});}
    log("✓ Option 4: đã xoá highlight tạm thời.","ok");
  }catch(e){log("✗ "+(e&&e.message?e.message:String(e)),"err");}
}

/* ═══ File Events ═══ */
async function handleFile(inputEl,fnameId,slot,setGuids){
  var f=inputEl.files&&inputEl.files[0];if(!f)return;
  var fnEl=document.getElementById(fnameId);
  fnEl.textContent=f.name;fnEl.classList.remove("hidden");
  setBadge(slot,"pending");
  log('Đọc [File#'+slot+'] "'+f.name+'"...',"info");
  try{
    var text=await readText(f);
    var guids=parseSPM(text);
    setGuids(guids);
    checkApplyBtn(slot);
    setBadge(slot,"ready");
    log('  ✓ '+guids.length+' GUID',"ok");
  }catch(e){
    log("  ✗ "+(e&&e.message?e.message:String(e)),"err");
    setGuids([]);checkApplyBtn(slot);
    setBadge(slot,"error");
  }
}

document.getElementById("file1").addEventListener("change",function(){handleFile(this,"fname1",1,function(g){_guids1=g;});});
document.getElementById("file2").addEventListener("change",function(){handleFile(this,"fname2",2,function(g){_guids2=g;});});
document.getElementById("applyBtn1").addEventListener("click",function(){paintSlot(1);});
document.getElementById("applyBtn2").addEventListener("click",function(){paintSlot(2);});
document.getElementById("resetBtn").addEventListener("click",resetViewer);
document.getElementById("saveBtn").addEventListener("click",saveView);
document.getElementById("captureSelBtn").addEventListener("click",captureSelection);
document.getElementById("paintSelBtn").addEventListener("click",paintSelection);
document.getElementById("noteInput1").addEventListener("blur",function(){localStorage.setItem("mcc_note_slot1",this.value);addNoteMarkup(1);});
document.getElementById("noteInput2").addEventListener("blur",function(){localStorage.setItem("mcc_note_slot2",this.value);addNoteMarkup(2);});
document.getElementById("qtyBtn1").addEventListener("click",function(){showQty(1,_map1);});
document.getElementById("qtyBtn2").addEventListener("click",function(){showQty(2,_map2);});
document.getElementById("qtyBtn3").addEventListener("click",async function(){await captureSelection();if(_selMap)showQty(3,_selMap);});
/* Option 4 wiring */
(function(){
  var sb=document.getElementById("apSearchBtn");if(sb)sb.addEventListener("click",apSearch);
  var pf=document.getElementById("apPrefix");if(pf)pf.addEventListener("keydown",function(e){if(e.key==="Enter"){e.preventDefault();apSearch();}});
  var cb=document.getElementById("apClearBtn");if(cb)cb.addEventListener("click",apClearAll);
})();
renderViewList();
