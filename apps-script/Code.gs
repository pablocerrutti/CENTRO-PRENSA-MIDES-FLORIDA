const CONFIG = {
  spreadsheetId: '',
  sheetName: 'Comunicados',
  adminKey: 'Ik3r2026',
  driveRootFolderId: '1bgzF1n5ufGlIQ84ykL90pAnJqvoWL2ET'
};

const HEADERS = ['ID','Estado','Fecha','Categoria','Tag','Titulo','Resumen','Contenido','FotoPrincipal','Fotos','VideoURL','VideoDriveFileId','VideoDriveUrl','VideoDriveFolderId','FechaCreacion','FechaActualizacion'];

function getSpreadsheet_() {
  return CONFIG.spreadsheetId ? SpreadsheetApp.openById(CONFIG.spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
}
function getSheet_() {
  const ss=getSpreadsheet_();
  let sh=ss.getSheetByName(CONFIG.sheetName);
  if(!sh) sh=ss.insertSheet(CONFIG.sheetName);
  if(sh.getLastRow()===0) sh.appendRow(HEADERS);
  ensureHeaders_(sh);
  return sh;
}
function ensureHeaders_(sh){
  const current=sh.getLastColumn()?sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]:[];
  HEADERS.forEach((h,i)=>{if(current[i]!==h)sh.getRange(1,i+1).setValue(h);});
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
    if(p.action==='adminList')return json_({ok:true,items:readAll_().sort((a,b)=>String(b.fecha).localeCompare(String(a.fecha)))});
    if(p.action==='adminGet')return json_({ok:true,item:findById_(p.id)});
    if(p.action==='save')return json_({ok:true,item:save_(p)});
    if(p.action==='savePublish')return json_({ok:true,item:saveAndPublish_(p)});
    if(p.action==='uploadVideoStart')return json_({ok:true,upload:uploadVideoStart_(p)});
    if(p.action==='uploadVideoChunk')return json_({ok:true,upload:uploadVideoChunk_(p)});
    if(p.action==='uploadVideoComplete')return json_({ok:true,upload:uploadVideoComplete_(p)});
    if(p.action==='uploadImages')return json_({ok:true,images:uploadImages_(p)});
    if(p.action==='delete')return json_({ok:delete_(p.id)});
    if(p.action==='publish'){const item=setStatus_(p.id,'Publicado');return json_({ok:!!item,item:item,error:item?null:'No se encontró el comunicado para publicar.'});}
    if(p.action==='unpublish')return json_({ok:setStatus_(p.id,'Borrador'),item:findById_(p.id)});
    return json_({ok:false,error:'Acción POST no válida'});
  }catch(err){return json_({ok:false,error:String(err)});}
}
function listPublished_(){return readAll_().filter(x=>x.estado==='Publicado').sort((a,b)=>String(b.fecha).localeCompare(String(a.fecha)));}
function readAll_(){const v=getSheet_().getDataRange().getValues();return v.length<2?[]:v.slice(1).filter(r=>r[0]).map(rowToObject_);}
function rowToObject_(r){
  return {id:String(r[0]),estado:String(r[1]),fecha:formatDate_(r[2]),categoria:String(r[3]),tag:String(r[4]),titulo:String(r[5]),resumen:String(r[6]),contenido:String(r[7]),fotoPrincipal:String(r[8]||''),fotos:splitPhotos_(r[9]),videoUrl:String(r[10]||''),videoDriveFileId:String(r[11]||''),videoDriveUrl:String(r[12]||''),videoDriveFolderId:String(r[13]||''),fechaCreacion:formatDateTime_(r[14]),fechaActualizacion:formatDateTime_(r[15])};
}
function splitPhotos_(v){return String(v||'').split(/\n+/).map(x=>x.trim()).filter(Boolean);}
function findById_(id){if(!id)return null;const sh=getSheet_(),row=findRow_(sh,id);return row<0?null:rowToObject_(sh.getRange(row,1,1,HEADERS.length).getValues()[0]);}
function save_(p){
  const sh=getSheet_(),id=p.id||Utilities.getUuid(),now=new Date(),old=findById_(id);
  const item={id:id,estado:old?old.estado:'Borrador',fecha:p.fecha||'',categoria:p.categoria||'Comunicado',tag:p.tag||'',titulo:p.titulo||'',resumen:p.resumen||'',contenido:p.contenido||'',fotoPrincipal:p.fotoPrincipal||'',fotos:splitPhotos_(p.fotos),videoUrl:p.videoUrl||'',videoDriveFileId:old?.videoDriveFileId||'',videoDriveUrl:old?.videoDriveUrl||'',videoDriveFolderId:old?.videoDriveFolderId||'',fechaCreacion:old&&old.fechaCreacion?old.fechaCreacion:formatDateTime_(now),fechaActualizacion:formatDateTime_(now)};
  const row=[item.id,item.estado,item.fecha,item.categoria,item.tag,item.titulo,item.resumen,item.contenido,item.fotoPrincipal,item.fotos.join('\n'),item.videoUrl,item.videoDriveFileId,item.videoDriveUrl,item.videoDriveFolderId,item.fechaCreacion,item.fechaActualizacion];
  const values=sh.getDataRange().getValues();let n=-1;
  for(let i=1;i<values.length;i++)if(String(values[i][0])===id){n=i+1;break;}
  if(n<0)sh.appendRow(row);else sh.getRange(n,1,1,row.length).setValues([row]);
  return item;
}
function uploadVideoStart_(p){
  if(!p.id)throw new Error('Primero guardá el comunicado como borrador.');
  if(!p.fileName)throw new Error('No se recibió el nombre del video.');
  const root=getDriveRoot_(),titleFolder=getOrCreateCommunicationFolder_(root,p.titulo||'Comunicado '+p.id),mime=p.mimeType||'video/mp4';
  const response=UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable',{method:'post',contentType:'application/json; charset=UTF-8',headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken(),'X-Upload-Content-Type':mime,'X-Upload-Content-Length':String(p.totalSize||0)},payload:JSON.stringify({name:p.fileName,mimeType:mime,parents:[titleFolder.getId()]}),muteHttpExceptions:true,followRedirects:false});
  const status=response.getResponseCode(),headers=response.getAllHeaders();
  if(status<200||status>=300)throw new Error('No se pudo iniciar la carga en Google Drive: HTTP '+status+' '+response.getContentText());
  const location=headers.Location||headers.location;
  if(!location)throw new Error('Google Drive no devolvió la URL de carga reanudable.');
  return {sessionUrl:String(location),folderId:titleFolder.getId(),folderName:titleFolder.getName()};
}
function uploadVideoComplete_(p){
  if(!p.id||!p.fileId)throw new Error('Faltan datos para finalizar la carga del video.');
  const sh=getSheet_(),row=findRow_(sh,p.id);
  if(row<0)throw new Error('No se encontró el comunicado para asociar el video.');
  const file=DriveApp.getFileById(p.fileId);
  try{file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);}catch(err){}
  const fileUrl='https://drive.google.com/file/d/'+p.fileId+'/view',previewUrl='https://drive.google.com/file/d/'+p.fileId+'/preview';
  sh.getRange(row,12,1,4).setValues([[p.fileId,fileUrl,p.folderId||'',formatDateTime_(new Date())]]);
  sh.getRange(row,16).setValue(formatDateTime_(new Date()));
  return {complete:true,fileId:p.fileId,fileUrl:fileUrl,previewUrl:previewUrl,folderId:p.folderId||'',folderName:p.folderName||'',name:file.getName()};
}
function uploadVideoChunk_(p){
  if(!p.sessionUrl)throw new Error('Falta la sesión de carga de Google Drive.');
  const total=Number(p.totalSize||0),start=Number(p.start||0),end=Number(p.end||0);
  if(!total||end<=start||end>total)throw new Error('Rango de video inválido.');
  const data=String(p.chunkData||'').replace(/^data:[^;]+;base64,/,'');
  if(!data)throw new Error('No se recibió el fragmento del video.');
  const bytes=Utilities.base64Decode(data);
  const response=UrlFetchApp.fetch(p.sessionUrl,{method:'put',headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken(),'Content-Length':String(bytes.length),'Content-Range':'bytes '+start+'-'+(end-1)+'/'+total},payload:bytes,muteHttpExceptions:true,followRedirects:false});
  const status=response.getResponseCode();
  if(status===308){const h=response.getAllHeaders(),range=h.Range||h.range||'';return {complete:false,nextStart:range?parseInt(String(range).split('-').pop(),10)+1:end};}
  if(status!==200&&status!==201)throw new Error('Google Drive rechazó el fragmento: HTTP '+status+' '+response.getContentText());
  const file=JSON.parse(response.getContentText());
  if(!file.id)throw new Error('Google Drive no devolvió el ID del archivo.');
  try{DriveApp.getFileById(file.id).setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);}catch(err){}
  const sh=getSheet_(),row=findRow_(sh,p.id);
  if(row<0)throw new Error('No se encontró el comunicado para asociar el video.');
  const fileUrl='https://drive.google.com/file/d/'+file.id+'/view',previewUrl='https://drive.google.com/file/d/'+file.id+'/preview';
  sh.getRange(row,12,1,4).setValues([[file.id,fileUrl,p.folderId||'',formatDateTime_(new Date())]]);
  sh.getRange(row,16).setValue(formatDateTime_(new Date()));
  return {complete:true,fileId:file.id,fileUrl:fileUrl,previewUrl:previewUrl,folderId:p.folderId||'',folderName:p.folderName||'',name:file.name||''};
}
function uploadImages_(p){
  if(!p.id)throw new Error('Primero guardá el comunicado como borrador.');
  if(!p.files)throw new Error('No se recibieron imágenes.');
  const files=JSON.parse(p.files),max=10*1024*1024,root=getDriveRoot_(),titleFolder=getOrCreateCommunicationFolder_(root,p.titulo||'Comunicado '+p.id),uploaded=[];
  files.forEach(f=>{const data=String(f.data||'').replace(/^data:[^;]+;base64,/,'');const bytes=Utilities.base64Decode(data);if(bytes.length>max)throw new Error('La imagen '+f.name+' supera el límite de 10 MB.');const file=titleFolder.createFile(Utilities.newBlob(bytes,f.type||'image/jpeg',f.name));try{file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);}catch(err){}uploaded.push({id:file.getId(),name:file.getName(),url:file.getUrl(),directUrl:'https://drive.google.com/uc?export=view&id='+file.getId()});});
  const sh=getSheet_(),row=findRow_(sh,p.id);if(row<0)throw new Error('No se encontró el comunicado para asociar las imágenes.');
  const current=splitPhotos_(sh.getRange(row,10).getValue());sh.getRange(row,10).setValue(current.concat(uploaded.map(x=>x.directUrl)).join('\n'));sh.getRange(row,16).setValue(formatDateTime_(new Date()));
  return {folderId:titleFolder.getId(),folderName:titleFolder.getName(),items:uploaded};
}
function getDriveRoot_(){return CONFIG.driveRootFolderId?DriveApp.getFolderById(CONFIG.driveRootFolderId):DriveApp.getRootFolder();}
function getOrCreateFolder_(parent,name){const it=parent.getFoldersByName(name);return it.hasNext()?it.next():parent.createFolder(name);}
function getOrCreateCommunicationFolder_(parent,title){return getOrCreateFolder_(parent,safeFolderName_(title));}
function safeFolderName_(name){return String(name).replace(/[\\/:*?"<>|#%{}~&]/g,' ').replace(/\s+/g,' ').trim().slice(0,150)||'Comunicado';}
function findRow_(sh,id){const v=sh.getDataRange().getValues();for(let i=1;i<v.length;i++)if(String(v[i][0])===String(id))return i+1;return -1;}
function setStatus_(id,status){const sh=getSheet_(),v=sh.getDataRange().getValues();for(let i=1;i<v.length;i++)if(String(v[i][0])===String(id)){const now=formatDateTime_(new Date());sh.getRange(i+1,2).setValue(status);sh.getRange(i+1,16).setValue(now);return rowToObject_(sh.getRange(i+1,1,1,HEADERS.length).getValues()[0]);}return null;}
function saveAndPublish_(p){const item=save_(p);const sh=getSheet_(),row=findRow_(sh,item.id);if(row<0)throw new Error('No se encontró el comunicado recién guardado.');const now=formatDateTime_(new Date());sh.getRange(row,2).setValue('Publicado');sh.getRange(row,16).setValue(now);return rowToObject_(sh.getRange(row,1,1,HEADERS.length).getValues()[0]);}
function delete_(id){const sh=getSheet_(),v=sh.getDataRange().getValues();for(let i=1;i<v.length;i++)if(String(v[i][0])===String(id)){sh.deleteRow(i+1);return true;}return false;}
function formatDate_(v){if(!v)return '';if(Object.prototype.toString.call(v)==='[object Date]')return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd');return String(v);}
function formatDateTime_(v){if(!v)return '';if(Object.prototype.toString.call(v)==='[object Date]')return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss');return String(v);}
function json_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}