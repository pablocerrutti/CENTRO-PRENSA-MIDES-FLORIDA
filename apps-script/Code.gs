const CONFIG = {
  spreadsheetId: '',
  sheetName: 'Comunicados',
  adminKey: 'CAMBIAR-ESTA-CLAVE'
};

const HEADERS = ['ID','Estado','Fecha','Categoria','Tag','Titulo','Resumen','Contenido','FechaCreacion','FechaActualizacion'];

function getSpreadsheet_() {
  return CONFIG.spreadsheetId ? SpreadsheetApp.openById(CONFIG.spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
}
function getSheet_() {
  const ss=getSpreadsheet_();
  let sh=ss.getSheetByName(CONFIG.sheetName);
  if(!sh) sh=ss.insertSheet(CONFIG.sheetName);
  if(sh.getLastRow()===0) sh.appendRow(HEADERS);
  return sh;
}
function setup(){const sh=getSheet_();sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);sh.setFrozenRows(1);return json_({ok:true,sheet:sh.getName(),headers:HEADERS});}
function doGet(e){
  try{
    const a=(e&&e.parameter&&e.parameter.action)||'list';
    if(a==='ping') return json_({ok:true,service:'Centro de Prensa MIDES Florida'});
    if(a==='list') return json_({ok:true,items:listPublished_()});
    if(a==='get'){const item=findById_(e.parameter.id);const publicItem=item&&item.estado==='Publicado'?item:null;return json_({ok:!!publicItem,item:publicItem});}
    return json_({ok:false,error:'Acción GET no válida'});
  }catch(err){return json_({ok:false,error:String(err)});}
}
function doPost(e){
  try{
    const p=e&&e.parameter?e.parameter:{};
    if(p.key!==CONFIG.adminKey)return json_({ok:false,error:'Clave editorial incorrecta'});
    if(p.action==='adminList')return json_({ok:true,items:readAll_().sort((a,b)=>String(b.fecha).localeCompare(String(a.fecha)))});\n    if(p.action==='adminGet')return json_({ok:true,item:findById_(p.id)});\n    if(p.action==='save')return json_({ok:true,item:save_(p)});
    if(p.action==='delete')return json_({ok:delete_(p.id)});
    if(p.action==='publish')return json_({ok:setStatus_(p.id,'Publicado'),item:findById_(p.id)});
    if(p.action==='unpublish')return json_({ok:setStatus_(p.id,'Borrador'),item:findById_(p.id)});
    return json_({ok:false,error:'Acción POST no válida'});
  }catch(err){return json_({ok:false,error:String(err)});}
}
function listPublished_(){return readAll_().filter(x=>x.estado==='Publicado').sort((a,b)=>String(b.fecha).localeCompare(String(a.fecha)));}
function readAll_(){const v=getSheet_().getDataRange().getValues();return v.length<2?[]:v.slice(1).filter(r=>r[0]).map(rowToObject_);}
function rowToObject_(r){return{id:String(r[0]),estado:String(r[1]),fecha:formatDate_(r[2]),categoria:String(r[3]),tag:String(r[4]),titulo:String(r[5]),resumen:String(r[6]),contenido:String(r[7]),fechaCreacion:formatDateTime_(r[8]),fechaActualizacion:formatDateTime_(r[9])};}
function findById_(id){return id?readAll_().find(x=>x.id===String(id))||null:null;}
function save_(p){
  const sh=getSheet_(),id=p.id||Utilities.getUuid(),now=new Date(),old=findById_(id);
  const item={id:id,estado:old?old.estado:'Borrador',fecha:p.fecha||'',categoria:p.categoria||'Comunicado',tag:p.tag||'',titulo:p.titulo||'',resumen:p.resumen||'',contenido:p.contenido||'',fechaCreacion:old&&old.fechaCreacion?old.fechaCreacion:formatDateTime_(now),fechaActualizacion:formatDateTime_(now)};
  const row=[item.id,item.estado,item.fecha,item.categoria,item.tag,item.titulo,item.resumen,item.contenido,item.fechaCreacion,item.fechaActualizacion];
  const values=sh.getDataRange().getValues();let n=-1;
  for(let i=1;i<values.length;i++)if(String(values[i][0])===id){n=i+1;break;}
  if(n<0)sh.appendRow(row);else sh.getRange(n,1,1,row.length).setValues([row]);
  return item;
}
function setStatus_(id,status){
  const sh=getSheet_(),v=sh.getDataRange().getValues();
  for(let i=1;i<v.length;i++)if(String(v[i][0])===String(id)){sh.getRange(i+1,2).setValue(status);sh.getRange(i+1,10).setValue(formatDateTime_(new Date()));return true;}
  return false;
}
function delete_(id){
  const sh=getSheet_(),v=sh.getDataRange().getValues();
  for(let i=1;i<v.length;i++)if(String(v[i][0])===String(id)){sh.deleteRow(i+1);return true;}
  return false;
}
function formatDate_(v){if(!v)return '';if(Object.prototype.toString.call(v)==='[object Date]')return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd');return String(v);}
function formatDateTime_(v){if(!v)return '';if(Object.prototype.toString.call(v)==='[object Date]')return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss');return String(v);}
function json_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}