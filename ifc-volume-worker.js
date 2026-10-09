/* IFC geometry volume worker. Geometry is read only from the loaded IFC source. */
var WEB_IFC_VERSION="0.0.78";
var WEB_IFC_ROOT="https://unpkg.com/web-ifc@"+WEB_IFC_VERSION+"/";
var ifcApiPromise=null;
var loadedModels=new Map();

async function getIfcApi(){
  if(!ifcApiPromise){
    ifcApiPromise=import(WEB_IFC_ROOT+"web-ifc-api.js").then(async function(WebIFC){
      var api=new WebIFC.IfcAPI();
      await api.Init(function(path){return WEB_IFC_ROOT+path;},true);
      return api;
    });
  }
  return ifcApiPromise;
}

function transformPoint(p,m){
  if(!m||m.length<16)return p;
  return {
    x:m[0]*p.x+m[4]*p.y+m[8]*p.z+m[12],
    y:m[1]*p.x+m[5]*p.y+m[9]*p.z+m[13],
    z:m[2]*p.x+m[6]*p.y+m[10]*p.z+m[14]
  };
}

function transformNormal(n,m){
  if(!m||m.length<16)return n;
  var x=m[0]*n.x+m[4]*n.y+m[8]*n.z;
  var y=m[1]*n.x+m[5]*n.y+m[9]*n.z;
  var z=m[2]*n.x+m[6]*n.y+m[10]*n.z;
  var len=Math.sqrt(x*x+y*y+z*z)||1;
  return {x:x/len,y:y/len,z:z/len};
}

function dot(a,b){return a.x*b.x+a.y*b.y+a.z*b.z;}
function cross(a,b){return{x:a.y*b.z-a.z*b.y,y:a.z*b.x-a.x*b.z,z:a.x*b.y-a.y*b.x};}
function sub(a,b){return{x:a.x-b.x,y:a.y-b.y,z:a.z-b.z};}

function placedGeometryVolume(api,modelId,placed){
  var geometry=api.GetGeometry(modelId,placed.geometryExpressID);
  try{
    var vertices=api.GetVertexArray(geometry.GetVertexData(),geometry.GetVertexDataSize());
    var indices=api.GetIndexArray(geometry.GetIndexData(),geometry.GetIndexDataSize());
    if(vertices.length%6!==0)throw new Error("Định dạng vertex IFC không phải XYZ+normal; dừng tính để tránh sai thể tích.");
    var stride=6;
    var matrix=placed.flatTransformation;
    function point(index){
      var at=index*stride;
      return transformPoint({x:vertices[at],y:vertices[at+1],z:vertices[at+2]},matrix);
    }
    function normal(index){
      if(stride<6)return null;
      var at=index*stride+3;
      return transformNormal({x:vertices[at],y:vertices[at+1],z:vertices[at+2]},matrix);
    }
    if(!indices||indices.length<3)return 0;
    var origin=point(indices[0]),volume=0;
    for(var i=0;i+2<indices.length;i+=3){
      var ia=indices[i],ib=indices[i+1],ic=indices[i+2];
      var a=sub(point(ia),origin),b=sub(point(ib),origin),c=sub(point(ic),origin);
      var n1=normal(ia),n2=normal(ib),n3=normal(ic);
      if(n1&&n2&&n3){
        var geometric=cross(sub(b,a),sub(c,a));
        var average={x:n1.x+n2.x+n3.x,y:n1.y+n2.y+n3.y,z:n1.z+n2.z+n3.z};
        if(dot(geometric,average)<0){var tmp=b;b=c;c=tmp;}
      }
      volume+=dot(a,cross(b,c))/6;
    }
    return volume;
  }finally{geometry.delete();}
}

function elementVolume(api,modelId,expressId){
  var mesh=api.GetFlatMesh(modelId,expressId),signed=0;
  try{
    if(!mesh||!mesh.geometries||!mesh.geometries.size())return null;
    for(var i=0;i<mesh.geometries.size();i++){
      signed+=placedGeometryVolume(api,modelId,mesh.geometries.get(i));
    }
  }finally{if(mesh&&typeof mesh.delete==="function")mesh.delete();}
  var volume=Math.abs(signed);
  return Number.isFinite(volume)&&volume>0?volume:null;
}

async function openModel(sourceKey,blob){
  if(loadedModels.has(sourceKey))return loadedModels.get(sourceKey);
  var api=await getIfcApi();
  if(!(blob instanceof Blob))throw new Error("Trimble Viewer không cung cấp nội dung IFC dạng Blob.");
  var buffer=await blob.arrayBuffer();
  var bytes=new Uint8Array(buffer);
  var modelId=api.OpenModel(bytes);
  if(modelId<0)throw new Error("WebIFC không mở được model nguồn.");
  var header=new TextDecoder("utf-8").decode(bytes.subarray(0,Math.min(bytes.length,8*1024*1024)));
  var unitMatch=header.match(/IFCSIUNIT\s*\(\s*\*\s*,\s*\.LENGTHUNIT\.\s*,\s*\.(MILLI|CENTI|DECI|KILO)?\.\s*,\s*\.METRE\.\s*\)/i);
  var unit=unitMatch?(unitMatch[1]||"").toUpperCase():null;
  var toMetres=unit==="MILLI"?0.001:unit==="CENTI"?0.01:unit==="DECI"?0.1:unit==="KILO"?1000:unit===""?1:null;
  if(toMetres==null){api.CloseModel(modelId);throw new Error("Không xác định được đơn vị chiều dài IFC dạng SI; không thể báo thể tích m³ an toàn.");}
  var state={api:api,modelId:modelId,toMetres:toMetres};
  loadedModels.set(sourceKey,state);
  return state;
}

self.onmessage=async function(event){
  var data=event.data||{};
  if(data.action!=="volumes")return;
  try{
    var state=await openModel(data.sourceKey,data.blob);
    var api=state.api,modelId=state.modelId,results=[];
    var guids=data.guids||[];
    for(var i=0;i<guids.length;i++){
      var guid=guids[i];
      var expressId=api.GetExpressIdFromGuid(modelId,guid);
      var nativeVolume=expressId!=null?elementVolume(api,modelId,Number(expressId)):null;
      var volume=nativeVolume==null?null:nativeVolume*Math.pow(state.toMetres,3);
      results.push({guid:guid,volumeM3:volume});
      if(i%10===0||i===guids.length-1){
        self.postMessage({id:data.id,type:"progress",done:i+1,total:guids.length});
      }
    }
    self.postMessage({id:data.id,type:"result",results:results});
  }catch(error){
    self.postMessage({id:data.id,type:"error",message:error&&error.message?error.message:String(error)});
  }
};
