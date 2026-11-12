const CONFIG = {
  SPREADSHEET_ID: '1ZglfWKES6sHWcn2vi9JVHXHVJImPqzcWGb_rYHaj0Uk',
  DRIVE_FOLDER_ID: '1KlTFhCQhSX40oqEps8cP-BVXRNxhglUM',
  MAX_FILES: 5,
  MAX_FILE_BYTES: 5 * 1024 * 1024,
};
const TYPES = {
  personal: { sheet: '個人渡眾', fields: ['country','unit','district','guide','eventDate','count'], photo: true },
  family: { sheet: '家堂渡眾', fields: ['country','unit','district','temple','host','eventDate','count'], photo: true, photoRequired: false },
  public: { sheet: '公堂渡眾', fields: ['country','unit','district','reporterName','temple','eventDate','count'], photo: true, photoRequired: false },
  personalClass: { sheet: '個人成全進階班', fields: ['country','unit','district','classType','mentor','student','eventDate','session'], photo: true },
  publicClass: { sheet: '公堂進階班', fields: ['country','unit','district','reporterName','temple','classType','count'], photo: true },
  temple: { sheet: '成全安堂完成', fields: ['country','unit','district','mentor','host','eventDate'], photo: true },
};
// 六種成績各自使用專屬欄位；沒有出現在該類表單的欄位不建立於該工作表。
const BASE_HEADERS = ['回報編號','送出時間','回報類別'];
const FIELD_HEADERS = {
  country: '國別', unit: '單位', district: '區域',
  guide: '引師姓名', reporterName: '回報人姓名',
  temple: '佛堂／公堂堂號', host: '壇主姓名',
  mentor: '成全者姓名', student: '班員姓名',
  classType: '進階班類別', eventDate: '辦道／上課／正式安堂日期',
  count: '渡眾／達標人數', session: '上課次數'
};
const TYPE_LABELS = {
  personal: { eventDate: '辦道日期', count: '渡眾人數' },
  family: { temple: '佛堂堂號', eventDate: '辦道日期', count: '渡眾人數' },
  public: { temple: '公堂堂號', eventDate: '辦道日期', count: '渡眾人數' },
  personalClass: { eventDate: '上課日期' },
  publicClass: { temple: '公堂堂號', count: '堂數達標之班員人數' },
  temple: { eventDate: '正式安堂日期' }
};
function headersFor_(type) {
  const spec = TYPES[type], names = TYPE_LABELS[type] || {};
  if (!spec) throw new Error('未知回報類別');
  return BASE_HEADERS.concat(spec.fields.map(key => names[key] || FIELD_HEADERS[key]), spec.photo ? ['照片連結'] : []);
}
function rowFor_(type,p,id,urls) {
  const spec=TYPES[type];
  const row=[id,new Date(),spec.sheet];
  spec.fields.forEach(key => {
    if (key==='count' || key==='session') row.push(Number(p[key]));
    else row.push(safeText_(p[key]));
  });
  if (spec.photo) row.push(urls.join('\n'));
  return row;
}
function ensureSchema_(sh,type) {
  const expected = headersFor_(type);
  if (sh.getLastRow()===0) {
    sh.getRange(1,1,1,expected.length).setValues([expected]);
    return;
  }
  const existing = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),expected.length)).getDisplayValues()[0];
  // 家堂／公堂渡眾原本沒有照片欄。若其餘欄名完全一致，只在末尾安全補上照片連結。
  if (['family','public'].includes(type) &&
      expected.length >= 1 && expected[expected.length-1] === '照片連結' &&
      expected.slice(0,-1).every((label,i)=>label===existing[i]) &&
      existing.slice(expected.length-1).every(v=>v==='')) {
    sh.getRange(1,expected.length).setValue('照片連結');
    return;
  }
  const matches = expected.every((label,i)=>label===existing[i]) &&
    existing.slice(expected.length).every(v=>v==='');
  if (!matches) throw new Error('「'+sh.getName()+'」仍是舊版欄位格式；為保護原有成績，請先備份並執行 migrateReportSheets()，不可直接混用新舊版。');
}
// 先備份後轉換舊資料。此函式需由管理員於 Apps Script 手動執行，切勿在回報時自動搬移。
// 以舊版共同欄位標題與回報類別識別各欄；不刪除原始內容，會建立隱藏的備份分頁。
function migrateReportSheets() {
  const ss=sheet_();
  const legacy = {
    '回報編號':'id','送出時間':'submitted','回報類別':'type',
    '國別':'country','單位':'unit','區別':'district','區域':'district',
    '引師姓名':'guide','佛堂或公堂堂號':'temple','佛堂／公堂堂號':'temple',
    '公堂堂號':'temple','佛堂堂號':'temple','壇主姓名':'host',
    '回報人姓名':'reporterName','成全者姓名':'mentor','班員姓名':'student',
    '進階班類別':'classType','辦道／上課／安堂日期':'eventDate',
    '辦道／上課／正式安堂日期':'eventDate','辦道日期':'eventDate',
    '上課日期':'eventDate','正式安堂日期':'eventDate',
    '渡眾／達標人數':'count','渡眾人數':'count',
    '堂數達標之班員人數':'count','上課次數':'session','照片連結':'files'
  };
  Object.keys(TYPES).forEach(type => {
    const sh=ss.getSheetByName(TYPES[type].sheet);
    if (!sh || sh.getLastRow()===0) return;
    const expected=headersFor_(type), width=sh.getLastColumn();
    const rows=sh.getDataRange().getValues();
    const heads=rows[0].map(h=>String(h).trim());
    if (heads.length===expected.length && heads.every((v,i)=>v===expected[i])) return;
    const mapped=heads.map(h=>legacy[h]||null);
    const keys=['id','submitted','type'].concat(TYPES[type].fields,TYPES[type].photo?['files']:[]);
    if (!keys.every(k=>mapped.includes(k) || (k==='files'&&['family','public'].includes(type)))) {
      throw new Error('無法辨認「'+sh.getName()+'」舊版欄位，停止轉換；請手動核對。');
    }
    const converted=rows.slice(1).filter(row=>row.some(x=>x!==''&&x!==null))
      .map(row=>keys.map(key=>mapped.includes(key)?row[mapped.indexOf(key)]:(key==='files'?'':null)));
    // 先建立完整備份，原始分頁內容保留在備份中。
    const backupName=('備份_'+sh.getName()+'_'+Utilities.formatDate(new Date(),'Asia/Taipei','MMdd_HHmmss')).slice(0,99);
    const backup=sh.copyTo(ss).setName(backupName);
    backup.hideSheet();
    sh.getRange(1,1,Math.max(sh.getLastRow(),converted.length+1),Math.max(width,expected.length)).clearContent();
    sh.getRange(1,1,1,expected.length).setValues([expected]);
    if(converted.length)sh.getRange(2,1,converted.length,expected.length).setValues(converted);
  });
}
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'options') return json_(getOptions_());
  return json_({ ok: true, message: '白陽慈航2027 API 已啟用' });
}
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error('沒有收到資料');
    const payload = JSON.parse(e.postData.contents);
    const action=payload.action || 'submit';
    let result;
    if(action==='submit') result=submit_(payload);
    else if(action==='lookup') result=lookupReport_(payload);
    else if(action==='changeRequest') result=requestChange_(payload);
    else throw new Error('不支援的操作');
    return json_(result);
  } catch (err) {
    return json_({ok:false,error:String(err.message||err)});
  }
}

function json_(value) { return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON); }
function sheet_() { return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID); }
function getOptions_() {
  const ss = sheet_();
  const sh = ss.getSheetByName('選項設定');
  if (!sh) throw new Error('找不到「選項設定」工作表');
  const rows = sh.getDataRange().getDisplayValues().slice(1);
  const options = rows.filter(r => r[0] && r[1] && r[2]).map(r => ({country:r[0].trim(),unit:r[1].trim(),district:r[2].trim()}));
  return {ok:true,options};
}
function validate_(p) {
  const spec = TYPES[p.type];
  if (!spec) throw new Error('無效的回報類別');
  for (const key of spec.fields) if (p[key] === undefined || p[key] === null || String(p[key]).trim() === '') throw new Error('缺少必填欄位：' + key);
  if (p.classType && !['禮行','義理','仁德'].includes(p.classType)) throw new Error('進階班類別錯誤');
  if (p.session && (!Number.isInteger(Number(p.session)) || +p.session < 1 || +p.session > 24)) throw new Error('上課次數必須為1至24');
  if (spec.fields.includes('count') && (!Number.isInteger(Number(p.count)) || Number(p.count) < 0)) throw new Error('人數必須是非負整數');
  if (p.eventDate && (!/^\d{4}-\d{2}-\d{2}$/.test(p.eventDate) || isNaN(Date.parse(p.eventDate)))) throw new Error('日期格式錯誤');
  const files = p.files || [];
  if (!Array.isArray(files) || files.length > CONFIG.MAX_FILES) throw new Error('最多可上傳5張照片');
  if (spec.photo && spec.photoRequired !== false && !files.length) throw new Error('此回報類別必須上傳照片');
  files.forEach(f => {
    if (!/^image\/(jpeg|png|webp|heic|heif)$/.test(f.mimeType)) throw new Error('僅接受照片檔案');
    if (!f.data || f.data.length > Math.ceil(CONFIG.MAX_FILE_BYTES * 4 / 3) + 100) throw new Error('照片過大或內容無效');
  });
  const opts = getOptions_().options;
  if (!opts.some(o => o.country === p.country && o.unit === p.unit && (!spec.fields.includes('district') || o.district === p.district))) throw new Error('國別、單位或區別不在選項設定內');
  return spec;
}
function safeText_(x) {
  const s = String(x === undefined || x === null ? '' : x).trim();
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}
/**
 * 2026/11/12 hotfix: 儲存報表照片。
 * 先完成所有照片建立，再將網址寫入回報列；上傳中途失敗時，
 * 僅回收「這次」剛建立的照片，絕不碰觸既有回報檔案。
 */
/**
 * 管理員在 Apps Script 編輯器手動執行此函式，讓 Google 顯示
 * 試算表／Drive OAuth 授權畫面；這個檢查不會建立檔案或修改成績。
 * 網頁使用者不需要，也不應取得整個 Drive 的編輯權限。
 */
function authorizeServices(){
  const spreadsheet=SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const folder=DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  const result={
    spreadsheet:spreadsheet.getName(),
    folder:folder.getName(),
    message:'已取得試算表與 Google Drive 資料夾的存取權。'
  };
  Logger.log(JSON.stringify(result));
  return result;
}
function storePhotos_(files,reportId,created){
  const list=files||[];
  if(!list.length)return [];
  if(!Array.isArray(list)||list.length>CONFIG.MAX_FILES)throw new Error('照片張數不符合限制');
  const extension={
    'image/jpeg':'jpg','image/png':'png','image/webp':'webp',
    'image/heic':'heic','image/heif':'heif'
  };
  let folder;
  try{folder=DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);}
  catch(err){
    const message=String(err&&err.message||err);
    if(/permission|authorization|權限|授權|scope/i.test(message))
      throw new Error('照片上傳尚未取得 Google Drive 權限；請管理員在 Apps Script 編輯器執行 authorizeServices() 並完成授權，再重新部署。');
    throw err;
  }
  const urls=[];
  try{
    list.forEach(function(file,index){
      const ext=extension[file.mimeType];
      if(!ext||!file.data||typeof file.data!=='string')throw new Error('照片格式或內容不正確');
      const bytes=Utilities.base64Decode(file.data);
      if(!bytes.length||bytes.length>CONFIG.MAX_FILE_BYTES)throw new Error('單張照片不可超過5MB');
      const name=reportId+'_'+(index+1)+'.'+ext;
      const blob=Utilities.newBlob(bytes,file.mimeType,name);
      const saved=folder.createFile(blob);
      created.push(saved);
      urls.push(saved.getUrl());
    });
    return urls;
  }catch(err){
    // 只回收本次未完成的照片，避免留下孤兒檔案。
    created.forEach(function(file){try{file.setTrashed(true);}catch(ignore){}});
    created.length=0;
    throw err;
  }
}
function submit_(p) {
  validate_(p);
  requireSecuritySetup_();
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  const created=[];
  try {
    const spec=TYPES[p.type],sh=reportSheet_(p.type),id=Utilities.getUuid();
    const files=storePhotos_(p.files||[],id,created);
    const pass=issueSecret_();
    // 工作表仍保持現有每類專屬欄位；憑證及狀態分開存放在「回報索引」。
    sh.appendRow(rowFor_(p.type,p,id,files));
    try {
      indexSheet_().appendRow([id,p.type,pass.salt,pass.hash,'已回報',1,new Date(),new Date()]);
    }catch(err){
      // 主表已寫入時不刪除原始資料，必須由管理員執行修復索引。
      throw new Error('成績可能已儲存，但無法建立查詢索引。請勿重複送出，聯絡管理員並提供回報時間。');
    }
    return {ok:true,id,lookupPassword:pass.password,reviewStatus:'已回報'};
  } finally {lock.releaseLock();}
}

function setupSheets() {
  initSecurity_();
  indexSheet_();
  changeSheet_();
  auditSheet_();
  const ss = sheet_();
  let options = ss.getSheetByName('選項設定');
  if (!options) options = ss.insertSheet('選項設定');
  if (options.getLastRow() === 0) options.appendRow(['國別','單位','區別']);
  Object.keys(TYPES).forEach(type => {
    const spec=TYPES[type];
    let sh=ss.getSheetByName(spec.sheet);
    if(!sh)sh=ss.insertSheet(spec.sheet);
    ensureSchema_(sh,type);
  });
}

/**
 * 查詢與更正模組（V2）
 * 公開路由僅 submit / lookup / changeRequest。
 * 管理員核准或拒絕只可從受權 Apps Script 編輯器執行 reviewChangeRequest()。
 */
const INDEX_HEADERS=['回報編號','回報類別','查詢鹽值','查詢密碼雜湊','審核狀態','版本','建立時間','更新時間'];
const CHANGE_HEADERS=['申請編號','回報編號','申請類型','原始版本','原內容JSON','申請內容JSON','申請理由','新照片連結JSON','狀態','申請時間','處理時間','管理員','處理說明'];
const AUDIT_HEADERS=['時間','動作','回報編號','申請編號','管理員','修改前JSON','修改後JSON','備註'];
function initSecurity_(){
  const props=PropertiesService.getScriptProperties();
  if(!props.getProperty('LOOKUP_PEPPER')){
    const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,Utilities.getUuid()+Utilities.getUuid());
    props.setProperty('LOOKUP_PEPPER',Utilities.base64EncodeWebSafe(bytes));
  }
}
function requireSecuritySetup_(){
  if(!PropertiesService.getScriptProperties().getProperty('LOOKUP_PEPPER'))throw new Error('尚未完成查詢安全初始化，請管理員執行 setupSheets()');
}
function managementSheet_(name,headers){
  const ss=sheet_();let sh=ss.getSheetByName(name);
  if(!sh)sh=ss.insertSheet(name);
  if(sh.getLastRow()===0)sh.getRange(1,1,1,headers.length).setValues([headers]);
  else {
    const got=sh.getRange(1,1,1,headers.length).getDisplayValues()[0];
    if(headers.some((h,i)=>h!==got[i]))throw new Error(name+' 欄位結構不符，停止操作');
  }
  return sh;
}
function indexSheet_(){return managementSheet_('回報索引',INDEX_HEADERS);}
function changeSheet_(){return managementSheet_('修改申請',CHANGE_HEADERS);}
function auditSheet_(){return managementSheet_('操作紀錄',AUDIT_HEADERS);}
function digest_(password,salt){
  const pepper=PropertiesService.getScriptProperties().getProperty('LOOKUP_PEPPER');
  const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,salt+':'+password+':'+pepper,Utilities.Charset.UTF_8);
  return bytes.map(b=>('0'+(b&255).toString(16)).slice(-2)).join('');
}
function issueSecret_(){
  const salt=Utilities.getUuid()+Utilities.getUuid();
  const password=Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,'');
  return {password,salt,hash:digest_(password,salt)};
}
function safeEquals_(a,b){
  const x=String(a),y=String(b);let diff=x.length^y.length;
  for(let i=0;i<Math.max(x.length,y.length);i++)diff|=(x.charCodeAt(i)||0)^(y.charCodeAt(i)||0);
  return diff===0;
}
function findIndex_(id){
  const sh=indexSheet_(),last=sh.getLastRow();
  if(last<2)return null;
  const finder=sh.getRange(2,1,last-1,1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  if(!finder)return null;
  const row=finder.getRow();
  return {sheet:sh,row,values:sh.getRange(row,1,1,INDEX_HEADERS.length).getValues()[0]};
}
function reportSheet_(type){
  const spec=TYPES[type];if(!spec)throw new Error('不支援的回報類別');
  const ss=sheet_();let sh=ss.getSheetByName(spec.sheet);
  if(!sh)sh=ss.insertSheet(spec.sheet);
  ensureSchema_(sh,type);return sh;
}
function findReport_(type,id){
  const sh=reportSheet_(type),last=sh.getLastRow();
  if(last<2)return null;
  const found=sh.getRange(2,1,last-1,1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  if(!found)return null;
  const row=found.getRow(),names=headersFor_(type);
  const vals=sh.getRange(row,1,1,names.length).getValues()[0];
  const values={};
  TYPES[type].fields.forEach((key,i)=>{
    const raw=vals[i+3];
    values[key]=raw instanceof Date?Utilities.formatDate(raw,'Asia/Taipei','yyyy-MM-dd'):String(raw===null||raw===undefined?'':raw);
  });
  return {sheet:sh,row,values,photos:TYPES[type].photo?String(vals[vals.length-1]||'').split('\n').filter(Boolean):[]};
}
function checkAccess_(p){
  requireSecuritySetup_();
  const id=String(p.id||'').trim(),password=String(p.lookupPassword||'');
  if(!/^[a-f0-9-]{36}$/i.test(id)||password.length<20||password.length>200)throw new Error('查詢資料不正確');
  const cache=CacheService.getScriptCache(),key='attempt:'+id;
  const attempts=Number(cache.get(key)||0);
  if(attempts>=8)throw new Error('驗證失敗次數過多，請15分鐘後再試');
  const rec=findIndex_(id),ok=rec&&safeEquals_(digest_(password,String(rec.values[2])),String(rec.values[3]));
  if(!ok){cache.put(key,String(attempts+1),900);throw new Error('回報編號或查詢密碼錯誤');}
  cache.remove(key);
  return rec;
}
function pendingChange_(id){
  const sh=changeSheet_(),last=sh.getLastRow();
  if(last<2)return false;
  return sh.getRange(2,1,last-1,CHANGE_HEADERS.length).getValues().some(row=>String(row[1])===id&&String(row[8])==='待審核');
}
function lookupReport_(p){
  const rec=checkAccess_(p),type=String(rec.values[1]),report=findReport_(type,p.id);
  if(!report)throw new Error('回報資料不存在，請聯絡管理員');
  return {ok:true,id:p.id,type,version:Number(rec.values[5]),status:String(rec.values[4]),
    fields:report.values,photoCount:report.photos.length,hasPendingRequest:pendingChange_(p.id)};
}
function validateChange_(type,oldValues,patch){
  if(!patch||typeof patch!=='object'||Array.isArray(patch))throw new Error('修改欄位無效');
  const fields=TYPES[type].fields;
  const keys=Object.keys(patch);
  if(!keys.length)throw new Error('請至少修改一個欄位');
  if(keys.some(k=>!fields.includes(k)||typeof patch[k]!=='string'))throw new Error('包含不允許修改的欄位');
  const merged=Object.assign({},oldValues,patch,{type,files:[]});
  // 原本必須附照片的回報，在修改既有文字內容時不要求重新上傳舊照片。
  validateChangeFields_(type,merged);
  return {merged,keys};
}
function validateChangeFields_(type,p){
  const spec=TYPES[type];
  spec.fields.forEach(k=>{if(!String(p[k]===undefined?'':p[k]).trim())throw new Error(k+' 為必填');});
  if(p.classType&&!['禮行','義理','仁德'].includes(p.classType))throw new Error('進階班類別錯誤');
  if(p.session&&(!Number.isInteger(Number(p.session))||Number(p.session)<1||Number(p.session)>24))throw new Error('上課次數錯誤');
  if(p.count!==undefined&&(!/^\d+$/.test(String(p.count))||!Number.isSafeInteger(Number(p.count))))throw new Error('人數必須為非負整數');
  if(p.eventDate){
    const m=/^(\d{4})-(\d{2})-(\d{2})$/.exec(String(p.eventDate));
    if(!m||Number(m[1])<1000||Number(m[1])>9999)throw new Error('日期必須為有效的四位數年份');
    const d=new Date(0);d.setFullYear(+m[1],+m[2]-1,+m[3]);
    if(d.getFullYear()!==+m[1]||d.getMonth()!==+m[2]-1||d.getDate()!==+m[3])throw new Error('日期不正確');
  }
  const list=getOptions_().options;
  if(!list.some(x=>x.country===p.country&&x.unit===p.unit&&x.district===p.district))throw new Error('國家、單位或區域不在設定名單內');
}
function requestChange_(p){
  const rec=checkAccess_(p),type=String(rec.values[1]),id=String(rec.values[0]);
  if(String(rec.values[4])==='已取消')throw new Error('已取消的回報不可申請修改');
  const action=String(p.changeType||'');
  if(!['modify','cancel'].includes(action))throw new Error('申請類型無效');
  const reason=String(p.reason||'').trim().slice(0,500);
  if(reason.length<4)throw new Error('請填寫修改或取消原因（至少4字）');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    // 重新核對最新狀態與版本，避免同時送出多個修改申請。
    const latest=findIndex_(id);
    if(!latest||String(latest.values[4])==='已取消'||pendingChange_(id))throw new Error('已有待審核申請或回報已取消');
    const report=findReport_(type,id);
    if(!report)throw new Error('找不到原始資料');
    const patch=action==='modify'?p.changes:{};
    if(action==='modify')validateChange_(type,report.values,patch);
    if(Number(p.version)!==Number(latest.values[5]))throw new Error('回報資料已更新，請重新查詢後修改');
    const requestId=Utilities.getUuid();
    changeSheet_().appendRow([requestId,id,action,Number(latest.values[5]),JSON.stringify(report.values),
      JSON.stringify(patch),safeText_(reason),'[]','待審核',new Date(),'','','']);
    return {ok:true,requestId,status:'待審核'};
  }finally{lock.releaseLock();}
}
/**
 * 只可由管理員登入 Apps Script 編輯器，從函式手動呼叫。
 * 不存在任何公開 API action 可執行核准。
 * 例如：在編輯器執行 reviewFirstPending() 會拒絕（僅為防誤操作的占位函式）。
 * 真正執行：在 Apps Script 編輯器建立暫時性執行函式，
 *   function approveSelected(){reviewChangeRequest('申請UUID','approve','已核對');}
 * 結束後刪除暫時函式，避免共用編輯者誤操作。
 */
function requireAdmin_(){
  const raw=PropertiesService.getScriptProperties().getProperty('ADMIN_EMAILS')||'';
  const allowed=raw.split(',').map(x=>x.trim().toLowerCase()).filter(Boolean);
  const current=(Session.getEffectiveUser().getEmail()||'').trim().toLowerCase();
  if(!current||!allowed.length||!allowed.includes(current))
    throw new Error('管理員權限未設定或目前帳號不在 ADMIN_EMAILS 允許清單');
  return current;
}
function reviewChangeRequest(requestId,decision,note){
  const adminEmail=requireAdmin_();
  requireSecuritySetup_();
  if(!['approve','reject'].includes(decision))throw new Error('只能核准或拒絕');
  if(!/^[a-f0-9-]{36}$/i.test(String(requestId)))throw new Error('申請編號格式錯誤');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try {
    const cs=changeSheet_(),last=cs.getLastRow();
    if(last<2)throw new Error('找不到申請');
    const cell=cs.getRange(2,1,last-1,1).createTextFinder(requestId).matchEntireCell(true).findNext();
    if(!cell)throw new Error('找不到申請');
    const row=cell.getRow(),r=cs.getRange(row,1,1,CHANGE_HEADERS.length).getValues()[0];
    if(String(r[8])!=='待審核')throw new Error('此申請已經處理');
    const idx=findIndex_(r[1]);
    if(!idx)throw new Error('回報索引遺失');
    if(Number(idx.values[5])!==Number(r[3]))throw new Error('原始資料已變更，請重新核對申請');
    const report=findReport_(String(idx.values[1]),String(r[1]));
    if(!report)throw new Error('原始回報不見了');
    const admin=adminEmail;
    const after=JSON.parse(JSON.stringify(report.values));
    const patches=JSON.parse(String(r[5]));
    if(decision==='approve'&&String(r[2])==='modify'){
      validateChange_(String(idx.values[1]),report.values,patches);
      Object.assign(after,patches);
    }
    const stamp=new Date();
    // 先留下不可省略的審核軌跡，再修改原始工作表。
    auditSheet_().appendRow([stamp,decision,String(r[1]),requestId,admin,JSON.stringify(report.values),JSON.stringify(after),safeText_(note||'')]);
    if(decision==='approve'&&String(r[2])==='modify'){
      const fields=TYPES[String(idx.values[1])].fields;
      const write=fields.map(k=>(k==='count'||k==='session')?Number(after[k]):safeText_(after[k]));
      report.sheet.getRange(report.row,4,1,fields.length).setValues([write]);
    }
    const status=decision==='approve'&&String(r[2])==='cancel'?'已取消':String(idx.values[4]);
    idx.sheet.getRange(idx.row,5).setValue(status);
    idx.sheet.getRange(idx.row,6).setValue(Number(idx.values[5])+1);
    idx.sheet.getRange(idx.row,8).setValue(stamp);
    cs.getRange(row,9).setValue(decision==='approve'?'已核准':'已拒絕');
    cs.getRange(row,11,1,3).setValues([[stamp,admin,safeText_(note||'')]]);
    return {ok:true,requestId,status:decision==='approve'?'已核准':'已拒絕'};
  }finally{lock.releaseLock();}
}
/**
 * 舊版回報沒有查詢密碼。此函式只能由 Apps Script 編輯者手動簽發，
 * 必須先由管理員確認申請人就是原回報人，再以安全管道交付一次性查詢憑證。
 * 不寫入原始成績工作表，不對外開放 API。
 */
function enrollLegacyReport(reportId){
  requireAdmin_();
  requireSecuritySetup_();
  if(!/^[a-f0-9-]{36}$/i.test(String(reportId)))throw new Error('回報編號格式錯誤');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    if(findIndex_(reportId))throw new Error('此回報已有查詢憑證，不可再次簽發');
    const matches=Object.keys(TYPES).filter(type=>findReport_(type,reportId));
    if(matches.length!==1)throw new Error('查無唯一對應的歷史回報');
    const secret=issueSecret_(),now=new Date();
    indexSheet_().appendRow([reportId,matches[0],secret.salt,secret.hash,'已回報',1,now,now]);
    return {id:reportId,lookupPassword:secret.password};
  }finally{lock.releaseLock();}
}

/**
 * ADMINISTRATOR ONLY. Never add these functions to doGet/doPost.
 * Uses Session.getEffectiveUser() plus ADMIN_EMAILS Script Property.
 * The admin editor must verify the requester independently before resetting credentials.
 * Query by UUID, not by name. Do not expose photo links to the public API.
 */
function adminLookupReport(reportId){
  const admin=requireAdmin_();
  requireSecuritySetup_();
  const id=String(reportId||'').trim();
  if(!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id))
    throw new Error('請提供完整的回報 UUID');
  const index=findIndex_(id);
  // Older reports may not be enrolled yet. The admin can see them but cannot reset
  // until verification/enrollment, avoiding creating a secret for unknown people.
  const types=index?[String(index.values[1])]:Object.keys(TYPES);
  const matches=types.map(type=>({type,report:findReport_(type,id)})).filter(x=>x.report);
  if(matches.length!==1)throw new Error('找不到唯一的原始回報；請先核對各類工作表');
  const item=matches[0],info={
    id:id,type:item.type,sheet:TYPES[item.type].sheet,
    status:index?String(index.values[4]):'舊版尚未建立查詢憑證',
    hasLookupCredentials:!!index,
    version:index?Number(index.values[5]):null,
    fields:item.report.values,photoCount:item.report.photos.length,
    pendingChange:index?pendingChange_(id):false
  };
  // Administrative access is logged. No password, salt or hash in log entries.
  auditSheet_().appendRow([new Date(),'admin_lookup',id,'',admin,'',JSON.stringify({
    type:info.type,status:info.status,photoCount:info.photoCount
  }),'管理員查詢回報']);
  return info;
}
/**
 * Identity verification is an out-of-band responsibility, not self-assertion.
 * Caller must supply a specific note recording how the requester was verified.
 * Password plaintext is returned ONCE, never written to Sheets or logs.
 * Existing reports and photos are preserved, previous password is invalidated.
 * Pending change requests remain pending but their version is unaffected.
 */
function resetLookupPassword(reportId,verificationNote){
  const admin=requireAdmin_();
  requireSecuritySetup_();
  const id=String(reportId||'').trim();
  const reason=String(verificationNote||'').trim();
  if(!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id))
    throw new Error('請提供完整的回報 UUID');
  if(reason.length<12||reason.length>500)
    throw new Error('請填寫至少12字的身分核實記錄（確認方式、對象與原因），最多500字');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    let idx=findIndex_(id);
    const reportTypes=idx?[String(idx.values[1])]:Object.keys(TYPES);
    const matches=reportTypes.filter(type=>findReport_(type,id));
    if(matches.length!==1)throw new Error('找不到唯一的原始回報，停止重設');
    if(idx&&String(idx.values[4])==='已取消')throw new Error('已取消回報不可重設查詢密碼');
    const secret=issueSecret_();
    const now=new Date();
    const indexSh=indexSheet_();
    // Audit before modifying the credential to avoid unlogged rotations.
    // The audit record contains neither the new secret nor the old hash/salt.
    auditSheet_().appendRow([now,idx?'credential_reset':'legacy_credential_issue',
      id,'',admin,'',JSON.stringify({type:matches[0],version:idx?Number(idx.values[5]):1}),
      safeText_(reason)]);
    if(idx){
      // Update only salt and hash. Preserve status, creation timestamp and
      // report version so pending correction requests are not invalidated.
      indexSh.getRange(idx.row,3,1,2).setValues([[secret.salt,secret.hash]]);
      indexSh.getRange(idx.row,8).setValue(now);
    }else{
      indexSh.appendRow([id,matches[0],secret.salt,secret.hash,'已回報',1,now,now]);
    }
    return {id,lookupPassword:secret.password,
      message:'新密碼僅於本次執行結果顯示，原密碼已失效。請以安全方式交付經核實的回報者。'};
  }finally{lock.releaseLock();}
}
/**
 * Admin editor convenience: select report and verification note in Script
 * Properties: ADMIN_TARGET_ID and ADMIN_VERIFICATION_NOTE (never a password).
 * Use adminLookupSelected() to inspect, then resetLookupPasswordSelected().
 * This avoids editing source or accidentally exposing a password in GitHub.
 * Do not store the returned password in Script Properties.
 */
function adminLookupSelected(){
  requireAdmin_();
  const id=PropertiesService.getScriptProperties().getProperty('ADMIN_TARGET_ID');
  const result=adminLookupReport(id);
  Logger.log(JSON.stringify(result));
  return result;
}
function resetLookupPasswordSelected(){
  requireAdmin_();
  const props=PropertiesService.getScriptProperties();
  const id=props.getProperty('ADMIN_TARGET_ID');
  const note=props.getProperty('ADMIN_VERIFICATION_NOTE');
  const result=resetLookupPassword(id,note);
  // Clear the target and reason to reduce accidental repeated rotations.
  props.deleteProperty('ADMIN_TARGET_ID');
  props.deleteProperty('ADMIN_VERIFICATION_NOTE');
  Logger.log('已重設回報 '+result.id+' 的密碼。新密碼僅於本次函式回傳值可取得。');
  return result;
}
