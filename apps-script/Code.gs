const CONFIG = {
  spreadsheetId: '',
  sheetName: 'Comunicados',
  adminKey: String.fromCharCode(73,107,51,114,50,48,50,54),
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
function ensureHeaders_(sh,headers){
  headers=headers||HEADERS;
  const current=sh.getLastColumn()?sh.getRange(1,1,1,Math.max(sh.getLastColumn(),headers.length)).getValues()[0]:[];
  headers.forEach((h,i)=>{if(current[i]!==h)sh.getRange(1,i+1).setValue(h);});
}
function autorizarCentroPrensa(){
  // Ejecutar manualmente una vez desde el editor de Apps Script con la cuenta propietaria.
  // Esto fuerza la solicitud de permisos de Sheets + Drive antes de usar el Web App.
  const ss=getSpreadsheet_();
  if(!ss)throw new Error('No se pudo acceder a la hoja de cálculo.');
  const root=getDriveRoot_();
  root.getName();
  if(CONFIG.driveRootFolderId){
    const folder=DriveApp.getFolderById(CONFIG.driveRootFolderId);
    folder.getName();
  }
  UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/about?fields=user',{
    headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},
    muteHttpExceptions:true
  });
  return 'Autorización de Google Sheets, Drive y solicitudes externas completada correctamente.';
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

const CONTACT_HEADERS=['ID','Nombre','Apellido','NombreCompleto','Medio','Cargo','Email','Telefono','Localidad','Listas','Estado','Observaciones','FechaActualizacion'];
const LIST_HEADERS=['ID','Nombre','Descripcion','Estado','FechaCreacion','FechaActualizacion'];
const MEMBERSHIP_HEADERS=['ContactoID','ListaID','FechaAsignacion'];
const CAMPAIGN_HEADERS=['ID','Fecha','Asunto','ComunicadoID','ListaIDs','Destinatarios','Estado','Notas'];
const CONTACT_SOURCE_SHEET='Contactos', CONTACTS_SHEET='Contactos_Normalizados', LISTS_SHEET='ListasMailing', MEMBERSHIPS_SHEET='Contactos_Listas', CAMPAIGNS_SHEET='Mailing_Campañas';
function getNamedSheet_(name,headers){const ss=getSpreadsheet_();let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);if(sh.getLastRow()===0)sh.appendRow(headers);ensureHeaders_(sh,headers);sh.setFrozenRows(1);return sh;}
function normalizeHeader_(s){return String(s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]/g,'');}
function contactHeaderMap_(sh){const h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0],m={};h.forEach((x,i)=>m[normalizeHeader_(x)]=i);return m;}
function firstCol_(m,names){for(const n of names){const k=normalizeHeader_(n);if(m[k]!==undefined)return m[k];}return -1;}
function contactSource_(){const sh=getSpreadsheet_().getSheetByName(CONTACT_SOURCE_SHEET);if(!sh||sh.getLastRow()<2)return {sheet:sh,rows:[],map:{}};return {sheet:sh,rows:sh.getRange(1,1,sh.getLastRow(),sh.getLastColumn()).getValues(),map:contactHeaderMap_(sh)};}
function sourceContact_(r,m){const val=a=>{const c=firstCol_(m,a);return c>=0?String(r[c]??'').trim():''};const nombre=val(['Nombre','First Name','Given Name']),apellido=val(['Apellido','Last Name','Family Name']),nombreCompleto=val(['NombreCompleto','Nombre completo','Name','Display Name'])||[nombre,apellido].filter(Boolean).join(' ');return {nombre,apellido,nombreCompleto,email:val(['Email','E-mail','Correo electrónico','Correo','Email 1 - Value','E-mail 1 - Value']),telefono:val(['Telefono','Teléfono','Phone','Mobile Phone','Phone 1 - Value']),medio:val(['Medio','Organization','Organización','Empresa','Company','Organization 1 - Name']),cargo:val(['Cargo','Title','Puesto','Organization 1 - Title']),localidad:val(['Localidad','Ciudad','City','Address 1 - City'])};}
function contactRow_(r){return {id:String(r[0]),nombre:String(r[1]||''),apellido:String(r[2]||''),nombreCompleto:String(r[3]||''),medio:String(r[4]||''),cargo:String(r[5]||''),email:String(r[6]||''),telefono:String(r[7]||''),localidad:String(r[8]||''),listas:String(r[9]||'').split(';').map(x=>x.trim()).filter(Boolean),estado:String(r[10]||'Activo'),observaciones:String(r[11]||''),fechaActualizacion:formatDateTime_(r[12])};}
function readContacts_(){
  const src=contactSource_();
  if(!src.sheet||src.rows.length<2)return [];
  const normalized=[];
  const membershipsSheet=getSpreadsheet_().getSheetByName(MEMBERSHIPS_SHEET);
  const listMap={};
  if(membershipsSheet&&membershipsSheet.getLastRow()>1){
    const mv=membershipsSheet.getRange(2,1,membershipsSheet.getLastRow()-1,MEMBERSHIP_HEADERS.length).getValues();
    mv.forEach(r=>{const cid=String(r[0]||''),lid=String(r[1]||'');if(cid&&lid)(listMap[cid]||(listMap[cid]=[])).push(lid);});
  }
  const lists=listMailingLists_();
  const namesById={};lists.forEach(l=>namesById[l.id]=l.nombre);
  src.rows.slice(1).forEach((r,i)=>{
    const o=sourceContact_(r,src.map);
    if(!o.nombreCompleto&&!o.email&&!o.telefono)return;
    const key=(o.email||'').toLowerCase().trim()||((o.telefono||'').replace(/\\D/g,'')?'tel-'+(o.telefono||'').replace(/\\D/g,''):'row-'+(i+2));
    const id='C-'+Utilities.base64EncodeWebSafe(key).replace(/=+$/,'').slice(0,20);
    const ids=listMap[id]||[];
    normalized.push({id,nombre:o.nombre,apellido:o.apellido,nombreCompleto:o.nombreCompleto,medio:o.medio,cargo:o.cargo,email:o.email,telefono:o.telefono,localidad:o.localidad,listas:ids.map(x=>namesById[x]).filter(Boolean),estado:'Activo',observaciones:'',fechaActualizacion:formatDateTime_(new Date())});
  });
  return normalized;
}
function importContacts_(){
  const src=contactSource_();
  if(!src.sheet)throw new Error('No existe la hoja Contactos.');
  const total=src.rows.slice(1).filter(r=>{const o=sourceContact_(r,src.map);return !!(o.nombreCompleto||o.email||o.telefono);}).length;
  ensureContactsStructure_();
  return {imported:total,updated:0,skipped:Math.max(0,src.rows.length-1-total),total};
}
function listContacts_(){
  const ss=getSpreadsheet_();
  const source=ss.getSheetByName(CONTACT_SOURCE_SHEET);
  if(!source)throw new Error('No existe la hoja Contactos. Creá o importá primero la hoja Contactos en este mismo archivo de Google Sheets.');
  ensureContactsStructure_();
  return readContacts_();
}
function saveContact_(p){const sh=getNamedSheet_(CONTACTS_SHEET,CONTACT_HEADERS),id=p.id||'C-'+Utilities.getUuid().slice(0,8),old=readContacts_().find(x=>x.id===id),item=[id,p.nombre||'',p.apellido||'',p.nombreCompleto||[p.nombre,p.apellido].filter(Boolean).join(' '),p.medio||'',p.cargo||'',p.email||'',p.telefono||'',p.localidad||'',p.listas||'',p.estado||old?.estado||'Activo',p.observaciones||old?.observaciones||'',formatDateTime_(new Date())];const rows=sh.getDataRange().getValues();let n=-1;for(let i=1;i<rows.length;i++)if(String(rows[i][0])===id){n=i+1;break;}if(n<0)sh.appendRow(item);else sh.getRange(n,1,1,CONTACT_HEADERS.length).setValues([item]);return contactRow_(item);}
function setContactStatus_(id,status){const c=readContacts_().find(x=>x.id===id);if(!c)return null;c.estado=status;return saveContact_(c);}
function ensureLists_(){const sh=getNamedSheet_(LISTS_SHEET,LIST_HEADERS),have=sh.getDataRange().getValues().slice(1).map(r=>String(r[1]).toLowerCase());['Prensa Florida','Radios','Televisión','Prensa escrita','Medios digitales','Medios nacionales','Institucional','Prioritarios'].forEach(n=>{if(!have.includes(n.toLowerCase()))sh.appendRow(['L-'+Utilities.getUuid().slice(0,8),n,'','Activa',formatDateTime_(new Date()),formatDateTime_(new Date())]);});}
function listMailingLists_(){ensureLists_();const sh=getSpreadsheet_().getSheetByName(LISTS_SHEET);return sh.getLastRow()<2?[]:sh.getRange(2,1,sh.getLastRow()-1,LIST_HEADERS.length).getValues().filter(r=>r[0]).map(r=>({id:String(r[0]),nombre:String(r[1]),descripcion:String(r[2]||''),estado:String(r[3]||'Activa'),fechaCreacion:formatDateTime_(r[4]),fechaActualizacion:formatDateTime_(r[5])}));}
function saveMailingList_(p){const sh=getNamedSheet_(LISTS_SHEET,LIST_HEADERS),id=p.id||'L-'+Utilities.getUuid().slice(0,8),rows=sh.getDataRange().getValues(),now=formatDateTime_(new Date());let n=-1;for(let i=1;i<rows.length;i++)if(String(rows[i][0])===id){n=i+1;break;}const v=[id,p.nombre||'',p.descripcion||'',p.estado||'Activa',p.fechaCreacion||now,now];if(n<0)sh.appendRow(v);else sh.getRange(n,1,1,LIST_HEADERS.length).setValues([v]);return {id,nombre:v[1],descripcion:v[2],estado:v[3]};}
function setMailingListStatus_(id,status){const l=listMailingLists_().find(x=>x.id===id);return l?saveMailingList_({...l,estado:status}):null;}
function saveMemberships_(contactId,listIds){const sh=getNamedSheet_(MEMBERSHIPS_SHEET,MEMBERSHIP_HEADERS),rows=sh.getDataRange().getValues(),ids=String(listIds||'').split(',').map(x=>x.trim()).filter(Boolean);for(let i=rows.length-1;i>=1;i--)if(String(rows[i][0])===String(contactId))sh.deleteRow(i+1);ids.forEach(id=>sh.appendRow([contactId,id,formatDateTime_(new Date())]));const c=readContacts_().find(x=>x.id===contactId);if(c){c.listas=listMailingLists_().filter(l=>ids.includes(l.id)).map(l=>l.nombre).join('; ');saveContact_(c);}return {contactId,listIds:ids};}
function previewCampaign_(p){const ids=String(p.listIds||'').split(',').map(x=>x.trim()).filter(Boolean),lists=listMailingLists_().filter(l=>ids.includes(l.id)),contacts=readContacts_().filter(c=>c.estado==='Activo'&&c.email),set={};contacts.forEach(c=>{const a=c.listas.map(x=>x.toLowerCase());if(lists.some(l=>a.includes(l.nombre.toLowerCase())))set[c.email.toLowerCase()]=c;});return {total:Object.keys(set).length,destinatarios:Object.values(set).map(c=>({id:c.id,nombre:c.nombreCompleto,email:c.email,medio:c.medio}))};}
function saveCampaign_(p){const sh=getNamedSheet_(CAMPAIGNS_SHEET,CAMPAIGN_HEADERS),preview=previewCampaign_(p),id=p.id||'M-'+Utilities.getUuid().slice(0,8),now=formatDateTime_(new Date());sh.appendRow([id,p.fecha||now,p.asunto||'',p.comunicadoId||'',p.listIds||'',preview.total,'Preparada',p.notas||'']);return {id,estado:'Preparada',destinatarios:preview.total};}

function doPost(e){
  try{
    const p=e&&e.parameter?e.parameter:{};
    if(p.key!==CONFIG.adminKey)return json_({ok:false,error:'Clave editorial incorrecta'});
    if(p.action==='adminList')return json_({ok:true,items:readAll_().sort((a,b)=>String(b.fecha).localeCompare(String(a.fecha)))});
    if(p.action==='adminGet')return json_({ok:true,item:findById_(p.id)});
    if(p.action==='save')return json_({ok:true,item:save_(p)});
    if(p.action==='savePublish')return json_({ok:true,item:saveAndPublish_(p)});
    if(p.action==='contactsList')return json_({ok:true,items:listContacts_()});
    if(p.action==='contactsImport')return json_({ok:true,result:importContacts_()});
    if(p.action==='contactSave')return json_({ok:true,item:saveContact_(p)});
    if(p.action==='contactDelete')return json_({ok:true,item:setContactStatus_(p.id,'Inactivo')});
    if(p.action==='listsList')return json_({ok:true,items:listMailingLists_()});
    if(p.action==='listSave')return json_({ok:true,item:saveMailingList_(p)});
    if(p.action==='listDelete')return json_({ok:true,item:setMailingListStatus_(p.id,'Inactiva')});
    if(p.action==='membershipSave')return json_({ok:true,item:saveMemberships_(p.contactId,p.listIds)});
    if(p.action==='campaignPreview')return json_({ok:true,preview:previewCampaign_(p)});
    if(p.action==='campaignSave')return json_({ok:true,item:saveCampaign_(p)});
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
function listPublished_(){return readAll_().filter(x=>String(x.estado||'').trim().toLowerCase()==='publicado').sort((a,b)=>String(b.fecha).localeCompare(String(a.fecha)));}
function readAll_(){const v=getSheet_().getDataRange().getValues();return v.length<2?[]:v.slice(1).filter(r=>r[0]).map(rowToObject_);}
function rowToObject_(r){
  return {id:String(r[0]),estado:String(r[1]||'').trim(),fecha:formatDate_(r[2]),categoria:String(r[3]),tag:String(r[4]),titulo:String(r[5]),resumen:String(r[6]),contenido:String(r[7]),fotoPrincipal:String(r[8]||''),fotos:splitPhotos_(r[9]),videoUrl:String(r[10]||''),videoDriveFileId:String(r[11]||''),videoDriveUrl:String(r[12]||''),videoDriveFolderId:String(r[13]||''),fechaCreacion:formatDateTime_(r[14]),fechaActualizacion:formatDateTime_(r[15])};
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