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
// 只輸出不含權限資訊的部署識別，用於核對「網站實際連到哪一版」。
const API_BUILD='regional-enhancements-20261112-v3';
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'options') return json_(getOptions_());
  if (action === 'health') return json_({ok:true,build:API_BUILD,
    actions:['submit','lookup','changeRequest','regionalLookup','regionalReport','regionalChangeRequest','regionalRequests']});
  return json_({ok:true,message:'白陽慈航2027 API 已啟用',build:API_BUILD});
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
    else if(action==='regionalLookup') result=regionalLookup_(payload);
    else if(action==='regionalReport') result=regionalReport_(payload);
    else if(action==='regionalChangeRequest') result=regionalChangeRequest_(payload);
    else if(action==='regionalRequests') result=regionalRequests_(payload);
    else throw new Error('不支援的操作：'+String(action).slice(0,40)+'；請確認部署版本與網站 API_URL 一致');
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
  // 新增 D 欄作為區域負責人手動輸入密碼；不覆蓋已存在的 D1 標題。
  if (!String(options.getRange('D1').getDisplayValue()||'').trim())
    options.getRange('D1').setValue('區域查詢密碼');
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
    const previousSnapshot=JSON.parse(String(r[4]));
    if(TYPES[String(idx.values[1])].fields.some(key=>
      String(report.values[key]||'')!==String(previousSnapshot[key]||'')))
      throw new Error('原始資料已被其他人修改，請先核對，不能直接審核');
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
 * 管理者視窗功能：只能從「與試算表綁定」的 Apps Script，
 * 在 Google 試算表上方「管理員工具」選單開啟。
 * 注意：直接在 Apps Script 編輯器按「執行」，不會有試算表 UI 可顯示。
 * 獨立型 Apps Script 專案無法使用 SpreadsheetApp.getUi()。
 */
function onOpen(e){
  SpreadsheetApp.getUi().createMenu('管理員工具')
    .addItem('查詢回報編號…','adminLookupSelected')
    .addItem('重設回報查詢密碼…','resetLookupPasswordSelected')
    .addSeparator()
    .addItem('審核修改／取消申請…','adminReviewPending')
    .addItem('成績、索引與照片健檢…','adminIntegrityCheckPrompt')
    .addItem('重新結算個人與區域成績…','adminRecalculateScoresPrompt')
    .addItem('啟用每5分鐘自動結算…','adminInstallAutoSettlementPrompt')
    .addSeparator()
    .addItem('核發舊式區域密碼（選用）…','adminIssueRegionalAccessPrompt')
    .addItem('停用區域查詢密碼…','adminRevokeRegionalAccessPrompt')
    .addToUi();
}
/**
 * 管理員對話框：查看待審核的修改／取消申請，再決定核准或拒絕。
 * 僅能從已綁定試算表的「管理員工具」選單使用；公開 API 不提供審核操作。
 * onOpen 建立選單並不代表授權，查詢與核准時都會再次要求 ADMIN_EMAILS。
 */
function adminPendingRequests_(){
  requireAdmin_();
  const sh=changeSheet_(),last=sh.getLastRow();
  if(last<2)return [];
  return sh.getRange(2,1,last-1,CHANGE_HEADERS.length).getValues()
    .filter(row=>String(row[8])==='待審核')
    .map(row=>({
      requestId:String(row[0]),reportId:String(row[1]),
      type:String(row[2]),version:Number(row[3]),
      reason:String(row[6]),submittedAt:row[9] instanceof Date
        ?Utilities.formatDate(row[9],'Asia/Taipei','yyyy/MM/dd HH:mm')
        :String(row[9])
    }));
}
function adminReviewDetails_(requestId){
  requireAdmin_();
  const id=String(requestId||'').trim();
  if(!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id))
    throw new Error('申請編號格式錯誤');
  const sh=changeSheet_(),last=sh.getLastRow();
  if(last<2)throw new Error('沒有修改申請');
  const hit=sh.getRange(2,1,last-1,1).createTextFinder(id).matchEntireCell(true).findNext();
  if(!hit)throw new Error('查無此申請編號');
  const vals=sh.getRange(hit.getRow(),1,1,CHANGE_HEADERS.length).getValues()[0];
  if(String(vals[8])!=='待審核')throw new Error('這筆申請已經處理，請重新整理待審核清單');
  const idx=findIndex_(String(vals[1]));
  if(!idx)throw new Error('找不到原始回報索引');
  const type=String(idx.values[1]);
  const original=JSON.parse(String(vals[4])),patch=JSON.parse(String(vals[5]));
  if(!original||typeof original!=='object'||!patch||typeof patch!=='object')
    throw new Error('申請資料格式錯誤');
  const current=findReport_(type,String(vals[1]));
  if(!current)throw new Error('找不到原始回報');
  const stale=Number(idx.values[5])!==Number(vals[3]) ||
    TYPES[type].fields.some(key=>String(current.values[key]||'')!==String(original[key]||''));
  const labelMap=Object.assign({},FIELD_HEADERS,TYPE_LABELS[type]||{});
  const changes=[];
  if(String(vals[2])==='modify'){
    Object.keys(patch).forEach(key=>{
      if(!TYPES[type].fields.includes(key))throw new Error('申請包含不合法的欄位，不能審核');
      changes.push((labelMap[key]||key)+'：'+String(original[key]||'')+'  →  '+String(patch[key]||''));
    });
  }else if(String(vals[2])==='cancel'){
    changes.push('申請取消整筆回報（原始紀錄保留，停止認列）');
  }else throw new Error('未知申請類型');
  return {
    requestId:id,reportId:String(vals[1]),
    type:String(vals[2]),sheet:TYPES[type].sheet,
    reason:String(vals[6]),version:Number(vals[3]),
    changes:changes,stale:stale,
    currentSummary:TYPES[type].fields
      .map(key=>(labelMap[key]||key)+'：'+String(current.values[key]||''))
  };
}
function adminReviewPending(){
  requireAdmin_();
  const ui=SpreadsheetApp.getUi();
  try{
    const pending=adminPendingRequests_();
    if(!pending.length){
      ui.alert('目前沒有待審核申請','「修改申請」分頁目前沒有狀態為待審核的紀錄。',ui.ButtonSet.OK);
      return {ok:true,pending:0};
    }
    const ordered=pending.slice().sort((a,b)=>String(b.submittedAt).localeCompare(String(a.submittedAt)));
    const display=ordered.slice(0,15);
    const modifications=pending.filter(item=>item.type==='modify').length;
    const cancellations=pending.filter(item=>item.type==='cancel').length;
    const intro='待審核共 '+pending.length+' 筆（修改 '+modifications+'、取消 '+cancellations+'）。'+
      '\n以下列出最新 '+display.length+' 筆：\n\n'+
      display.map((item,i)=>(i+1)+'. '+(item.type==='cancel'?'取消':'修改')+
        '｜'+item.submittedAt+'\n申請編號：'+item.requestId+
        '\n回報編號：'+item.reportId).join('\n\n')+
      '\n\n輸入 1～'+display.length+' 可選擇以上申請；也可直接輸入任一待審核申請 UUID。';
    ui.alert('待審核申請清單',intro,ui.ButtonSet.OK);
    const response=ui.prompt('選擇要審核的申請','輸入清單序號（例如 1），或貼上完整申請 UUID：',ui.ButtonSet.OK_CANCEL);
    if(response.getSelectedButton()!==ui.Button.OK)return {ok:false,cancelled:true};
    const selected=response.getResponseText().trim();
    const choice=/^\d{1,2}$/.test(selected)?display[Number(selected)-1]:null;
    const id=choice?choice.requestId:selected;
    if(!pending.some(item=>item.requestId===id)){
      ui.alert('找不到待審核申請','請確認輸入的是清單中的申請編號，並重新操作。',ui.ButtonSet.OK);
      return {ok:false,error:'not_pending'};
    }
    const detail=adminReviewDetails_(id);
    const message=[
      '回報類別：'+detail.sheet,
      '回報編號：'+detail.reportId,
      '申請編號：'+detail.requestId,
      '申請類型：'+(detail.type==='cancel'?'取消回報':'修改資料'),
      '申請原因：'+detail.reason,
      '',
      '申請內容：',
      detail.changes.join('\n'),
      '',
      '目前資料（供核對）：',
      detail.currentSummary.join('\n'),
      '',
      detail.stale?'⚠ 原始資料／版本已變更，不能直接核准。':
        '請核對辦道簿、求道卡或出席表等證明後再決定。',
      '',
      '按「是」＝核准；按「否」＝拒絕；按「取消」＝不處理。'
    ].join('\n');
    const decision=ui.alert('審核申請：請仔細核對',message,ui.ButtonSet.YES_NO_CANCEL);
    if(decision!==ui.Button.YES&&decision!==ui.Button.NO)return {ok:false,cancelled:true};
    if(detail.stale){
      ui.alert('資料已變更','原始回報內容或版本已與申請當時不同。請先人工核對，這次不會更動成績。',ui.ButtonSet.OK);
      return {ok:false,stale:true};
    }
    const action=decision===ui.Button.YES?'approve':'reject';
    const noteResponse=ui.prompt(action==='approve'?'核准說明':'拒絕原因',
      '請輸入審核依據與處理說明（至少 4 字），確認後才會儲存：',
      ui.ButtonSet.OK_CANCEL);
    if(noteResponse.getSelectedButton()!==ui.Button.OK)return {ok:false,cancelled:true};
    const note=noteResponse.getResponseText().trim();
    if(note.length<4||note.length>500){
      ui.alert('審核說明不完整','請填寫 4～500 字的審核依據，這次未做任何修改。',ui.ButtonSet.OK);
      return {ok:false,error:'note'};
    }
    const confirm=ui.alert('最後確認',
      '申請：'+detail.requestId+'\n決定：'+(action==='approve'?'核准':'拒絕')+
      '\n說明：'+note+'\n\n執行後會留下稽核紀錄。確定繼續嗎？',
      ui.ButtonSet.YES_NO);
    if(confirm!==ui.Button.YES)return {ok:false,cancelled:true};
    const result=reviewChangeRequest(id,action,note);
    ui.alert('審核已完成','申請編號：'+id+'\n處理結果：'+
      (action==='approve'?'已核准':'已拒絕')+
      '\n請確認「修改申請」與「操作紀錄」分頁。',ui.ButtonSet.OK);
    return result;
  }catch(err){
    ui.alert('審核未完成','未確認成功，請檢查資料與操作紀錄後再處理。\n'+String(err.message||err),ui.ButtonSet.OK);
    throw err;
  }
}
function adminPromptId_(ui,title){
  const response=ui.prompt(title,'請輸入完整回報編號（UUID）：',ui.ButtonSet.OK_CANCEL);
  if(response.getSelectedButton()!==ui.Button.OK)return null;
  const id=response.getResponseText().trim();
  if(!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id)){
    ui.alert('格式不正確','請輸入完整的回報編號，格式為 8-4-4-4-12。',ui.ButtonSet.OK);
    return null;
  }
  return id;
}
function adminLookupSelected(){
  requireAdmin_(); // 任何查詢前均須後端驗證管理員
  const ui=SpreadsheetApp.getUi();
  const id=adminPromptId_(ui,'管理員查詢回報');
  if(!id)return;
  try{
    const info=adminLookupReport(id);
    const labels={country:'國別',unit:'單位',district:'區域',guide:'引師姓名',
      reporterName:'回報人姓名',temple:'佛堂／公堂堂號',host:'壇主姓名',
      mentor:'成全者姓名',student:'班員姓名',classType:'進階班類別',
      eventDate:'日期',count:'人數',session:'上課次數'};
    const lines=['回報編號：'+info.id,'回報類別：'+info.sheet,
      '回報狀態：'+info.status,'照片張數：'+info.photoCount];
    Object.keys(info.fields).forEach(key=>{
      lines.push((labels[key]||key)+'：'+info.fields[key]);
    });
    ui.alert('查詢結果',lines.join('\n'),ui.ButtonSet.OK);
    return info;
  }catch(err){
    ui.alert('查詢失敗',String(err.message||err),ui.ButtonSet.OK);
    throw err;
  }
}
function resetLookupPasswordSelected(){
  requireAdmin_(); // 不經公開 doPost，且先檢查白名單
  const ui=SpreadsheetApp.getUi();
  const id=adminPromptId_(ui,'重設回報查詢密碼');
  if(!id)return;
  const reasonPrompt=ui.prompt('核實回報者身分',
    '請先以其他方式核實本人，再輸入至少 12 字的核實方式與重設原因（勿填敏感證件號碼）：',
    ui.ButtonSet.OK_CANCEL);
  if(reasonPrompt.getSelectedButton()!==ui.Button.OK)return;
  const reason=reasonPrompt.getResponseText().trim();
  if(reason.length<12||reason.length>500){
    ui.alert('核實紀錄不足','原因請填 12 至 500 字，未進行重設。',ui.ButtonSet.OK);
    return;
  }
  const confirmation=ui.alert('確認重設',
    '即將重設回報 '+id+' 的查詢密碼。原密碼會立即失效，原始成績及照片不會刪除。確認已核實本人？',
    ui.ButtonSet.YES_NO);
  if(confirmation!==ui.Button.YES)return;
  try{
    const result=resetLookupPassword(id,reason);
    // 只在受限的試算表管理介面顯示新密碼，不寫入執行記錄。
    // 請管理員立即安全交付；視窗關閉後無法取回原始密碼。
    ui.alert('密碼已重設（請立即保存）',
      '回報編號：'+result.id+'\n新查詢密碼：'+result.lookupPassword+
      '\n\n原密碼已失效。請透過安全管道交給經核實的回報者。關閉視窗後不能再次查看這組密碼。',
      ui.ButtonSet.OK);
    return {ok:true,id:result.id};
  }catch(err){
    ui.alert('重設失敗',String(err.message||err),ui.ButtonSet.OK);
    throw err;
  }
}




/**
 * 區域負責人專用檢視：每個檔案只包含指定「國別＋單位＋區域」。
 * 只讓 ADMIN_EMAILS 中的管理員產生／更新，並由管理員手動分享給該區域負責人（檢視者）。
 * 原始回報、照片網址、密碼雜湊、管理員操作紀錄絕不複製到檢視檔。
 */
const REGIONAL_VIEW_MARKER='BY2027_SINGLE_DISTRICT_V1';
const REGIONAL_SUMMARY_HEADERS=['國別','單位','區域','個人渡眾人數','家堂渡眾人數','公堂渡眾人數','個人成全進階班回報筆數','公堂進階班達標人數','成全安堂完成筆數','有效回報總筆數'];
const REGIONAL_DETAIL_HEADERS=['回報編號','送出時間','回報類別','國別','單位','區域','目前狀態','人數／達標數','進階班類別','上課次數','日期','引師姓名','回報人姓名','佛堂／公堂堂號','壇主姓名','成全者姓名','班員姓名','照片張數'];
function safeRegionCell_(x){
  const v=x instanceof Date?Utilities.formatDate(x,'Asia/Taipei','yyyy/MM/dd HH:mm:ss'):
    String(x===undefined||x===null?'':x);
  return /^[=+\-@\t\r]/.test(v)?"'"+v:v;
}
function regionKey_(country,unit,district){return JSON.stringify([country,unit,district]);}
function validRegion_(country,unit,district){
  const c=String(country||'').trim(),u=String(unit||'').trim(),d=String(district||'').trim();
  if(!c||!u||!d||!getOptions_().options.some(o=>o.country===c&&o.unit===u&&o.district===d))
    throw new Error('指定國別／單位／區域不在選項設定中，停止建立報表');
  return {country:c,unit:u,district:d};
}
function regionViewRows_(country,unit,district){
  const target=validRegion_(country,unit,district);
  const ss=sheet_(),index=ss.getSheetByName('回報索引');
  const statusById={};
  if(index&&index.getLastRow()>1)
    index.getRange(2,1,index.getLastRow()-1,5).getValues().forEach(row=>{
      if(row[0])statusById[String(row[0])]=String(row[4]||'已回報');
    });
  const total=[target.country,target.unit,target.district,0,0,0,0,0,0,0],details=[];
  Object.keys(TYPES).forEach(type=>{
    const spec=TYPES[type],sh=ss.getSheetByName(spec.sheet);
    if(!sh||sh.getLastRow()<2)return;
    const header=headersFor_(type),actual=sh.getRange(1,1,1,header.length).getDisplayValues()[0];
    if(header.some((h,i)=>h!==actual[i]))
      throw new Error('工作表「'+spec.sheet+'」欄位格式不符，為避免誤送資料已停止');
    sh.getRange(2,1,sh.getLastRow()-1,header.length).getValues().forEach(row=>{
      if(!row[0])return;
      const p={};spec.fields.forEach((k,i)=>p[k]=row[i+3]);
      // 先比對完整三段鍵值，才允許任何紀錄出現在輸出陣列。
      if(p.country!==target.country||p.unit!==target.unit||p.district!==target.district)return;
      const id=String(row[0]),state=statusById[id]||'舊版未建立查詢憑證';
      const count=Number(p.count||0);
      const photoCount=spec.photo?String(row[row.length-1]||'').split('\n').filter(Boolean).length:0;
      details.push([
        id,row[1],spec.sheet,target.country,target.unit,target.district,state,
        spec.fields.includes('count')?count:'',
        p.classType||'',p.session||'',p.eventDate||'',p.guide||'',
        p.reporterName||'',p.temple||'',p.host||'',p.mentor||'',p.student||'',photoCount
      ].map(safeRegionCell_));
      if(state==='已取消')return;
      const index={personal:3,family:4,public:5,personalClass:6,publicClass:7,temple:8}[type];
      total[index]+=['personalClass','temple'].includes(type)?1:count;
      total[9]++;
    });
  });
  details.sort((a,b)=>String(a[1]).localeCompare(String(b[1])));
  return {target,summary:[total],details};
}
function districtViewIdKey_(country,unit,district){
  const data=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,regionKey_(country,unit,district));
  return 'DISTRICT_VIEW_'+data.map(v=>('0'+(v&255).toString(16)).slice(-2)).join('').slice(0,36);
}
function writeDistrictView_(view,tabName,header,rows){
  let sh=view.getSheetByName(tabName);if(!sh)sh=view.insertSheet(tabName);
  if(sh.getMaxColumns()<header.length)sh.insertColumnsAfter(sh.getMaxColumns(),header.length-sh.getMaxColumns());
  if(sh.getMaxRows()<rows.length+1)sh.insertRowsAfter(sh.getMaxRows(),rows.length+1-sh.getMaxRows());
  sh.getRange(1,1,sh.getMaxRows(),header.length).clearContent();
  sh.getRange(1,1,1,header.length).setValues([header]);
  if(rows.length)sh.getRange(2,1,rows.length,header.length).setValues(rows);
  sh.setFrozenRows(1);sh.getRange(1,1,1,header.length).setFontWeight('bold').setBackground('#e4f0e7');
  if(!sh.getFilter())sh.getRange(1,1,sh.getMaxRows(),header.length).createFilter();
}
function makeDistrictView_(country,unit,district){
  const admin=requireAdmin_(),data=regionViewRows_(country,unit,district);
  const key=districtViewIdKey_(country,unit,district);
  const props=PropertiesService.getScriptProperties();
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const existing=props.getProperty(key);
    let view;
    if(existing){
      view=SpreadsheetApp.openById(existing);
      const meta=view.getSheetByName('檢視說明');
      if(!meta||meta.getRange('A1').getDisplayValue()!==REGIONAL_VIEW_MARKER||
         meta.getRange('B1').getDisplayValue()!==regionKey_(country,unit,district))
        throw new Error('檢視表身分或區域不符，停止更新以免洩漏其他區域資料');
      // 每次更新前檢查檢視表只有已知分頁，防止留下先前非本區域資料。
      if(view.getSheets().some(sh=>!['檢視說明','區域彙整','回報明細'].includes(sh.getName())))
        throw new Error('檢視表存在非預期分頁，為保護資料已停止更新');
    }else{
      view=SpreadsheetApp.create('白陽慈航2027｜'+data.target.country+'｜'+data.target.unit+'｜'+data.target.district+'｜區域成績（唯讀分享）');
      const meta=view.getSheets()[0];meta.setName('檢視說明');
      meta.getRange('A1').setValue(REGIONAL_VIEW_MARKER);
      meta.getRange('B1').setValue(regionKey_(country,unit,district));
      meta.getRange('A2').setValue('此檔案僅供所列區域負責人檢視，請勿公開分享或授予編輯權限。');
      meta.getRange('A3').setValue('已取消回報保留在明細，但不計入區域彙整；照片檔案與密碼資料不會提供。');
      meta.getRange('A4').setValue('如需更新成績，請由管理員重新執行同步；本檔案不會自動更新。');
    }
    writeDistrictView_(view,'區域彙整',REGIONAL_SUMMARY_HEADERS,data.summary);
    writeDistrictView_(view,'回報明細',REGIONAL_DETAIL_HEADERS,data.details);
    view.getSheetByName('檢視說明').getRange('A5').setValue('資料更新時間：'+Utilities.formatDate(new Date(),'Asia/Taipei','yyyy/MM/dd HH:mm:ss'));
    if(!existing)props.setProperty(key,view.getId());
    auditSheet_().appendRow([new Date(),existing?'district_view_refresh':'district_view_create','','',admin,'',JSON.stringify({
      country,unit,district,rows:data.details.length
    }),'指定區域專用檢視表，權限需由管理員另行設定']);
    return {url:view.getUrl(),target:data.target,rows:data.details.length,isNew:!existing};
  }finally{lock.releaseLock();}
}
/** 由綁定的試算表「管理員工具」執行，逐項輸入國別、單位、區域。 */
function adminDistrictViewPrompt(){
  requireAdmin_();
  const ui=SpreadsheetApp.getUi();
  const inputs=[];
  for(const [label,desc] of [
    ['國別','請輸入選項設定中的完整國別（例如：台灣）'],
    ['單位','請輸入選項設定中的完整單位'],
    ['區域','請輸入選項設定中的完整區域（例如：中和001）']
  ]){
    const response=ui.prompt('建立／更新單一區域報表｜'+label,desc,ui.ButtonSet.OK_CANCEL);
    if(response.getSelectedButton()!==ui.Button.OK)return;
    inputs.push(response.getResponseText().trim());
  }
  let result;
  try{result=makeDistrictView_(inputs[0],inputs[1],inputs[2]);}
  catch(err){ui.alert('建立失敗',String(err.message||err),ui.ButtonSet.OK);throw err;}
  ui.alert('單一區域報表已準備完成',
    '範圍：'+inputs.join('／')+'\n回報筆數：'+result.rows+
    '\n\n請至執行記錄取得此檔案網址，並由管理員手動分享給該區域負責人（設為檢視者）。不要分享原始成績試算表。',
    ui.ButtonSet.OK);
  Logger.log('區域報表（請限該區域負責人檢視）：'+result.url);
  return result;
}


/**
 * 區域負責人前端查詢（V3）：同一組獨立查詢密碼可綁定多個指定區域。
 * 不能以公開 API 發密碼／更改權限；只有 ADMIN_EMAILS 的 Apps Script 管理員可核發。
 * 不使用原回報查詢密碼，不返回照片連結或其他區域資料。
 */
const REGIONAL_ACCESS_HEADERS=['憑證編號','查詢鹽值','查詢雜湊','允許區域JSON','狀態','建立時間','更新時間','核發管理員'];
function regionalAccessSheet_(){return managementSheet_('區域查詢授權',REGIONAL_ACCESS_HEADERS);}
function parseRegionalScope_(raw){
  if(!Array.isArray(raw)||!raw.length||raw.length>100)throw new Error('請指定1～100個區域');
  const seen=new Set(),regions=[];
  raw.forEach(item=>{
    if(!item||typeof item!=='object')throw new Error('區域格式不符');
    const country=String(item.country||'').trim(),unit=String(item.unit||'').trim(),
      district=String(item.district||'').trim();
    const key=regionKey_(country,unit,district);
    if(!seen.has(key)){
      regions.push(validRegion_(country,unit,district));
      seen.add(key);
    }
  });
  return regions;
}
function regionalSecret_(){
  const password=Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,'');
  const salt=Utilities.getUuid()+Utilities.getUuid();
  return {password,salt,hash:digest_(password,salt)};
}
function adminIssueRegionalAccess_(areas){
  const actor=requireAdmin_();requireSecuritySetup_();
  const regions=parseRegionalScope_(areas);
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const sh=regionalAccessSheet_(),secret=regionalSecret_(),id=Utilities.getUuid(),now=new Date();
    sh.appendRow([id,secret.salt,secret.hash,JSON.stringify(regions),'有效',now,now,actor]);
    auditSheet_().appendRow([now,'regional_access_issue','',id,actor,'',
      JSON.stringify({regions:regions.length}), '核發指定區域負責人唯讀密碼']);
    return {id,password:secret.password,regions};
  }finally{lock.releaseLock();}
}
function adminRevokeRegionalAccess_(credentialId){
  const actor=requireAdmin_();const id=String(credentialId||'').trim();
  if(!/^[0-9a-f-]{36}$/i.test(id))throw new Error('憑證編號格式不正確');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const sh=regionalAccessSheet_(),end=sh.getLastRow();
    if(end<2)throw new Error('沒有授權紀錄');
    const hit=sh.getRange(2,1,end-1,1).createTextFinder(id).matchEntireCell(true).findNext();
    if(!hit)throw new Error('查無授權憑證');
    sh.getRange(hit.getRow(),5).setValue('停用');
    sh.getRange(hit.getRow(),7).setValue(new Date());
    auditSheet_().appendRow([new Date(),'regional_access_revoke','',id,actor,'','','停用區域查詢密碼']);
    return {ok:true,id};
  }finally{lock.releaseLock();}
}
function adminIssueRegionalAccessPrompt(){
  requireAdmin_();
  const ui=SpreadsheetApp.getUi();
  const p=ui.prompt('核發區域負責人查詢密碼',
    '請輸入要授權的區域，每行一個，格式為：國別/單位/區域\n例如：台灣/中和/中和001\n可輸入多行以授權同一位負責人查看多個區域。',
    ui.ButtonSet.OK_CANCEL);
  if(p.getSelectedButton()!==ui.Button.OK)return;
  const lines=p.getResponseText().split(/\r?\n/).map(x=>x.trim()).filter(Boolean);
  let regions;
  try{
    regions=parseRegionalScope_(lines.map(line=>{
      const parts=line.split('/').map(x=>x.trim());
      if(parts.length!==3)throw new Error('每行必須是 國別/單位/區域');
      return {country:parts[0],unit:parts[1],district:parts[2]};
    }));
  }catch(err){ui.alert('授權區域錯誤',String(err.message||err),ui.ButtonSet.OK);return;}
  const confirm=ui.alert('確認核發',
    '確定允許此查詢密碼查看以下 '+regions.length+' 個區域？\n'+
    regions.map(x=>[x.country,x.unit,x.district].join('／')).join('\n')+
    '\n\n密碼請私下交付負責人，勿張貼於群組。',ui.ButtonSet.YES_NO);
  if(confirm!==ui.Button.YES)return;
  const result=adminIssueRegionalAccess_(regions);
  ui.alert('區域查詢密碼（請立即保存）',
    '授權編號：'+result.id+'\n查詢密碼：'+result.password+
    '\n\n請立即複製並私下交付區域負責人。關閉後無法找回原密碼；如需停用請記下授權編號。',
    ui.ButtonSet.OK);
  return {ok:true,id:result.id};
}
function adminRevokeRegionalAccessPrompt(){
  requireAdmin_();const ui=SpreadsheetApp.getUi();
  const p=ui.prompt('停用區域查詢密碼','請輸入當初核發的授權編號（不是回報編號）：',ui.ButtonSet.OK_CANCEL);
  if(p.getSelectedButton()!==ui.Button.OK)return;
  const yes=ui.alert('確認停用','停用後該查詢密碼將無法再查看區域資料，確定？',ui.ButtonSet.YES_NO);
  if(yes!==ui.Button.YES)return;
  const result=adminRevokeRegionalAccess_(p.getResponseText().trim());
  ui.alert('已停用','授權編號：'+result.id,ui.ButtonSet.OK);
}
/**
 * 「選項設定」D 欄：區域查詢密碼（手動 KEY）。
 * A~C 欄原樣保留。D 欄只在伺服器讀取，永遠不回傳到公開選項 API。
 * 同一密碼出現在多列，即代表可以查看那幾列對應的所有區域。
 * 明文 D 欄有讀取風險：原始試算表務必只供授權管理員存取。
 */
function optionPasswordRegions_(password){
  const sh=sheet_().getSheetByName('選項設定');
  if(!sh||sh.getLastRow()<2||sh.getLastColumn()<4)return [];
  // 相容現有工作表的「密碼」標題，也支援完整的「區域查詢密碼」。
  const passwordHeader=String(sh.getRange('D1').getDisplayValue()||'').trim();
  if(!['區域查詢密碼','密碼'].includes(passwordHeader))return [];
  // A~D 範圍，只由伺服器讀取。
  const records=sh.getRange(2,1,sh.getLastRow()-1,4).getDisplayValues();
  const regions=[],seen={};
  const requested=digest_(password,'regional-manual-option');
  records.forEach(row=>{
    // 同一區域可在 D 欄填寫多組密碼，以換行、半形／全形逗號或分號分隔。
    // 維持舊版單一密碼格式相容；精確比對完整密碼，不能只比對部分字串。
    const passwords=String(row[3]||'').split(/[\r\n,，;；]+/)
      .map(value=>value.trim()).filter(Boolean);
    if(passwords.length>20)throw new Error('同一區域最多可設定20組查詢密碼');
    const matched=passwords.some(stored=>{
      if(stored.length<6||stored.length>128||stored[0]==='=')return false;
      return safeEquals_(requested,digest_(stored,'regional-manual-option'));
    });
    if(!matched)return;
    const country=String(row[0]||'').trim(),unit=String(row[1]||'').trim(),
      district=String(row[2]||'').trim();
    if(!country||!unit||!district)return;
    const key=regionKey_(country,unit,district);
    if(!seen[key]){seen[key]=true;regions.push({country,unit,district});}
  });
  if(regions.length>100)throw new Error('此密碼對應超過100個區域，請聯絡管理員');
  return regions;
}
/**
 * Best-effort shared failed-authentication guard in addition to per-password throttling.
 * GAS web apps do not reliably expose a trusted caller IP; this global limiter may
 * temporarily block legitimate users during an attack. Never log entered passwords.
 * Cache increment is not atomic, so deploy-time quotas/network throttling are still needed.
 */
function regionalFailedThrottle_(){
  const cache=CacheService.getScriptCache();
  const key='region_invalid_global_15min';
  if(Number(cache.get(key)||0)>=80)throw new Error('區域查詢暫時繁忙，請15分鐘後再試或聯絡管理員');
}
function recordRegionalFailure_(){
  const cache=CacheService.getScriptCache(),key='region_invalid_global_15min';
  cache.put(key,String(Math.min(1000,Number(cache.get(key)||0)+1)),900);
}
function regionalAllowed_(passwordInput){
  requireSecuritySetup_();
  const password=String(passwordInput||'').trim();
  // 相容舊版核發的64位密碼，以及選項設定D欄可手動輸入的6～128字密碼。
  if(password.length<6||password.length>128)throw new Error('查詢密碼錯誤');
  regionalFailedThrottle_();
  const cache=CacheService.getScriptCache();
  // 限制同一錯誤密碼的重複嘗試；不把原始密碼放進快取鍵。
  const cacheKey='regional_fail_'+digest_(password,'region-attempt').slice(0,40);
  if(Number(cache.get(cacheKey)||0)>=8)throw new Error('嘗試次數過多，請稍後再試');
  // 優先使用「選項設定」D欄的手動密碼；不從前端讀取任何授權範圍。
  const optionRegions=optionPasswordRegions_(password);
  let allowed=optionRegions;
  // 保留舊版管理員核發的密碼，避免既有憑證升級後失效。
  if(!allowed.length){
    const sh=sheet_().getSheetByName('區域查詢授權');
    if(sh&&sh.getLastRow()>1&&/^[a-f0-9]{64}$/i.test(password)){
      const rows=sh.getRange(2,1,sh.getLastRow()-1,REGIONAL_ACCESS_HEADERS.length).getValues();
      for(const row of rows){
        if(String(row[4])!=='有效')continue;
        if(safeEquals_(digest_(password,String(row[1])),String(row[2]))){
          allowed=parseRegionalScope_(JSON.parse(String(row[3])));break;
        }
      }
    }
  }
  if(!allowed.length){
    recordRegionalFailure_();
    cache.put(cacheKey,String(Number(cache.get(cacheKey)||0)+1),900);
    throw new Error('查詢密碼錯誤或已停用');
  }
  cache.remove(cacheKey);
  return allowed;
}
function regionalLookup_(payload){
  const allowed=regionalAllowed_(payload.regionalPassword);
  const results=allowed.map(x=>regionViewRows_(x.country,x.unit,x.district));
  const details=[],summary=[];
  results.forEach(x=>{
    summary.push(x.summary[0]);x.details.forEach(row=>details.push(row));
  });
  return {ok:true,regions:allowed,summaryHeaders:REGIONAL_SUMMARY_HEADERS,
    detailHeaders:REGIONAL_DETAIL_HEADERS,summary,details,updatedAt:new Date().toISOString()};
}


/**
 * 區域負責人先經密碼授權，再取回單筆目前欄位供編輯。
 * 不可由前端傳來的區域或類別自行提升權限。
 */
/**
 * 取得唯一原始回報。舊版回報可能尚未有「回報索引」，
 * 仍必須先找到唯一存在的紀錄，才能讓區域負責人申請更正。
 * 使用前須先以區域密碼驗證，並於返回前比對原始區域。
 */
function regionalReportSource_(id,idx){
  const types=idx?[String(idx.values[1])]:Object.keys(TYPES);
  if(types.some(type=>!TYPES[type]))throw new Error('回報索引類別有誤，請管理員核對');
  const matches=[];
  types.forEach(type=>{
    // 不建立不存在的原始分頁；只讀取有紀錄的工作表。
    if(!sheet_().getSheetByName(TYPES[type].sheet))return;
    const report=findReport_(type,id);
    if(report)matches.push({type,report});
  });
  if(matches.length!==1)throw new Error('找不到唯一對應的原始回報，請管理員核對');
  return matches[0];
}
function regionalReport_(p){
  const allowed=regionalAllowed_(p.regionalPassword);
  const id=String(p.id||'').trim();
  if(!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id))
    throw new Error('回報編號格式不正確');
  const idx=findIndex_(id);
  const source=regionalReportSource_(id,idx);
  const f=source.report.values;
  if(!allowed.some(x=>x.country===f.country&&x.unit===f.unit&&x.district===f.district))
    throw new Error('此回報不在授權區域內');
  return {ok:true,id,type:source.type,version:idx?Number(idx.values[5]):0,
    status:idx?String(idx.values[4]):'已回報',
    fields:f,photoCount:source.report.photos.length,hasPendingRequest:pendingChange_(id),
    legacyWithoutIndex:!idx};
}
/**
 * 區域負責人只可申請修改自己目前授權區域內的既有回報文字資料。
 * 不可修改 country/unit/district：跨區調整必須聯絡管理員處理。
 * 與個別回報修改走相同的「修改申請」審核工作表。
 */
function regionalChangeRequest_(p){
  const allowed=regionalAllowed_(p.regionalPassword);
  const id=String(p.id||'').trim();
  if(!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id))
    throw new Error('回報編號格式不正確');
  const changeType=String(p.changeType||'modify');
  if(!['modify','cancel'].includes(changeType))throw new Error('申請類型無效');
  const patch=changeType==='cancel'?{}:p.changes;
  if(changeType==='modify'){
    if(!patch||typeof patch!=='object'||Array.isArray(patch)||!Object.keys(patch).length)
      throw new Error('請至少修改一個欄位');
    if(Object.keys(patch).some(k=>['country','unit','district'].includes(k)))
      throw new Error('區域負責人無法跨區修改國別、單位或區域，請聯絡管理員');
  }
  const reason=String(p.reason||'').trim();
  if(reason.length<4||reason.length>500)throw new Error('請填寫4至500字的修改或取消原因');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    // 在持鎖的寫入階段重新讀取 D 欄密碼與授權範圍，避免管理員剛撤銷權限。
    const latestAllowed=regionalAllowed_(p.regionalPassword);
    let idx=findIndex_(id);
    if(idx&&String(idx.values[4])==='已取消')throw new Error('已取消的回報不可提出修改申請');
    if(Number(p.version)!==(idx?Number(idx.values[5]):0))
      throw new Error('回報資料已更新，請重新查詢');
    if(pendingChange_(id))throw new Error('已有待審核申請，請等管理員處理');
    const source=regionalReportSource_(id,idx);
    const type=source.type,report=source.report,f=report.values;
    if(!latestAllowed.some(x=>x.country===f.country&&x.unit===f.unit&&x.district===f.district))
      throw new Error('此回報已不在密碼授權範圍');
    if(changeType==='modify'){
      validateChange_(type,f,patch);
      if(Object.keys(patch).every(k=>String(patch[k])===String(f[k])))throw new Error('修改內容與原始資料相同');
    }
    // 只有真正提交有效申請，才補建缺少的回報索引。
    // 隨機密碼不向區域負責人顯示，也不寫入日誌；以後如需個別查詢，
    // 應由管理員核實原回報者後再安全重設其專屬查詢密碼。
    if(!idx){
      const secret=issueSecret_(),now=new Date();
      auditSheet_().appendRow([now,'regional_legacy_index',id,'','系統','',
        JSON.stringify({type,location:[f.country,f.unit,f.district]}),
        '區域負責人提出修改或取消申請，為既有回報補建索引，不更動原始成績']);
      indexSheet_().appendRow([id,type,secret.salt,secret.hash,'已回報',1,now,now]);
      idx=findIndex_(id);
      if(!idx)throw new Error('建立回報索引失敗，請管理員核對，勿重複送出');
    }
    const requestId=Utilities.getUuid();
    changeSheet_().appendRow([requestId,id,changeType,Number(idx.values[5]),JSON.stringify(f),
      JSON.stringify(patch),safeText_('[區域負責人'+(changeType==='cancel'?'取消':'修改')+'申請] '+reason),'[]','待審核',new Date(),'','','']);
    return {ok:true,requestId,status:'待審核'};
  }finally{lock.releaseLock();}
}


/**
 * Show change-request statuses only for records that STILL belong to password-
 * authorized districts. No original/modified JSON, photos, password hashes or
 * admin-only audit data is returned.
 */
function regionalRequests_(payload){
  const allowed=regionalAllowed_(payload.regionalPassword);
  const allowedIds={};
  allowed.forEach(area=>regionViewRows_(area.country,area.unit,area.district)
    .details.forEach(row=>{allowedIds[String(row[0])]=true;}));
  const sh=sheet_().getSheetByName('修改申請');
  if(!sh||sh.getLastRow()<2)return {ok:true,requests:[],total:0};
  const rows=sh.getRange(2,1,sh.getLastRow()-1,CHANGE_HEADERS.length).getValues();
  const matched=rows.filter(row=>allowedIds[String(row[1])])
    .map(row=>({
      requestId:String(row[0]),reportId:String(row[1]),
      changeType:String(row[2]),reason:String(row[6]),
      status:String(row[8]),
      submittedAt:row[9] instanceof Date
        ?Utilities.formatDate(row[9],'Asia/Taipei','yyyy/MM/dd HH:mm'):String(row[9]),
      handledAt:row[10] instanceof Date
        ?Utilities.formatDate(row[10],'Asia/Taipei','yyyy/MM/dd HH:mm'):String(row[10]||''),
      reviewNote:String(row[12]||'')
    })).reverse();
  return {ok:true,requests:matched.slice(0,200),total:matched.length,truncated:matched.length>200};
}
/**
 * Read-only administrator health scan. Does not create indices, touch photos,
 * change permissions, or overwrite existing data. Photo URL existence check
 * is deliberately opt-in and limited to avoid expensive Drive calls.
 */
function auditReportIntegrity(includeDriveChecks){
  requireAdmin_();
  const ss=sheet_(),issues=[],ids={},summary={reports:0,indexed:0,photos:0,
    missingIndex:0,orphanIndex:0,duplicateId:0,missingPhotos:0,
    malformedPhotoLinks:0,missingDriveFiles:0,driveFilesChecked:0};
  const add=(kind,id,detail)=>{if(issues.length<150)issues.push({kind,id,detail});};
  const index=ss.getSheetByName('回報索引'),indexRows=index&&index.getLastRow()>1
    ?index.getRange(2,1,index.getLastRow()-1,INDEX_HEADERS.length).getValues():[];
  const indexIds={};
  indexRows.forEach(row=>{if(!row[0])return;summary.indexed++;const id=String(row[0]);
    if(indexIds[id]){summary.duplicateId++;add('重複索引',id,'回報索引有重複編號');}
    indexIds[id]={type:String(row[1]),status:String(row[4])};
  });
  Object.keys(TYPES).forEach(type=>{
    const spec=TYPES[type],sh=ss.getSheetByName(spec.sheet);
    if(!sh||sh.getLastRow()<2)return;
    const expected=headersFor_(type);
    const cols=sh.getRange(1,1,1,expected.length).getDisplayValues()[0];
    if(expected.some((x,i)=>x!==cols[i])){
      add('欄位格式',type,'該工作表的欄名與目前程式不一致，跳過內容檢查');return;
    }
    sh.getRange(2,1,sh.getLastRow()-1,expected.length).getValues().forEach(row=>{
      if(!row[0])return;
      summary.reports++;
      const id=String(row[0]);
      if(ids[id]){summary.duplicateId++;add('重複回報編號',id,'不同回報紀錄使用相同編號');}
      ids[id]=type;
      if(!indexIds[id]){summary.missingIndex++;add('缺少索引',id,spec.sheet);}
      else if(indexIds[id].type!==type)add('索引類別不一致',id,'索引與原始分頁類別不同');
      if(!spec.photo)return;
      const links=String(row[row.length-1]||'').split('\n').filter(Boolean);
      summary.photos+=links.length;
      if(spec.photoRequired!==false&&!links.length){
        summary.missingPhotos++;add('缺少照片連結',id,spec.sheet);
      }
      links.forEach(link=>{
        const match=/^https:\/\/drive\.google\.com\/(?:file\/d\/|open\?id=)([-\w]+)/.exec(link);
        if(!match){summary.malformedPhotoLinks++;add('照片連結格式',id,'照片連結格式不符');return;}
        if(includeDriveChecks&&summary.driveFilesChecked<50){
          summary.driveFilesChecked++;
          try{if(DriveApp.getFileById(match[1]).isTrashed()){
            summary.missingDriveFiles++;add('照片已移至垃圾桶',id,'請由管理員人工核對');
          }}catch(e){
            summary.missingDriveFiles++;add('照片存取異常',id,'檔案不存在或目前執行帳號無存取權');
          }
        }
      });
    });
  });
  Object.keys(indexIds).forEach(id=>{
    if(!ids[id]){summary.orphanIndex++;add('孤立索引',id,'索引存在，但六類原始回報均找不到');}
  });
  const result={ok:true,checkedAt:Utilities.formatDate(new Date(),'Asia/Taipei','yyyy/MM/dd HH:mm:ss'),
    summary,issues,truncated:issues.length>=150};
  Logger.log(JSON.stringify(result));
  return result;
}
function adminIntegrityCheckPrompt(){
  requireAdmin_();
  const ui=SpreadsheetApp.getUi();
  const confirm=ui.alert('成績資料健檢（唯讀）',
    '檢查六類回報、回報索引及照片連結格式。不會修改任何成績或照片。\n按「是」額外抽查前50個照片檔案是否存在（須Drive讀取權限）；按「否」只檢查資料表。',
    ui.ButtonSet.YES_NO_CANCEL);
  if(confirm!==ui.Button.YES&&confirm!==ui.Button.NO)return;
  const result=auditReportIntegrity(confirm===ui.Button.YES);
  const x=result.summary;
  ui.alert('健檢結果',
    '回報筆數：'+x.reports+'\n索引筆數：'+x.indexed+'\n缺少索引：'+x.missingIndex+
    '\n孤立索引：'+x.orphanIndex+'\n重複編號：'+x.duplicateId+
    '\n缺少照片連結：'+x.missingPhotos+'\n異常照片網址：'+x.malformedPhotoLinks+
    '\n抽查照片：'+x.driveFilesChecked+'\n照片存取異常：'+x.missingDriveFiles+
    '\n\n詳細問題已輸出至 Apps Script 執行記錄（最多150筆），未修改原始資料。',
    ui.ButtonSet.OK);
  return result;
}


/**
 * 管理員專用成績結算。只在管理員主動執行時，建立／更新兩張衍生表：
 * - 區域成績結算：六類成績獨立統計，不混合人數與筆數
 * - 個人成績結算：按真實欄位角色分組，不能把引師、成全者、班員、壇主混為一人得分
 * 取消已核准不計入；待審核照原始值計入，另標待審核筆數。
 * 不刪除或改寫六類原始回報、照片、索引與修改申請。
 */
const SETTLEMENT_REGION_HEADERS=['區域','個人渡眾人數','家堂渡眾人數','公堂渡眾人數',
  '個人成全進階班回報筆數','公堂進階班達標班員人數','成全安堂完成筆數',
  '有效回報筆數','待審核申請筆數','結算時間'];
const SETTLEMENT_PERSON_HEADERS=['區域','姓名','角色',
  '個人渡眾人數','個人成全進階班回報筆數','成全安堂完成筆數',
  '家堂渡眾人數（壇主記錄）','公堂渡眾人數（回報人記錄）',
  '公堂進階班達標人數（回報人記錄）',
  '有效相關回報筆數','結算時間'];
const SETTLEMENT_ROLE_FIELDS={
  personal:[['guide','引師',5]],
  family:[['host','壇主',8]],
  public:[['reporterName','回報人',9]],
  // 班員不計入個人成績；只保留成全者角色的紀錄。
  personalClass:[['mentor','成全者',6]],
  publicClass:[['reporterName','回報人',10]],
  temple:[['mentor','成全者',7]]
};
function settlementRows_(){
  const ss=sheet_(),regions={},people={},seenIds={},districtOwners={};
  const idxSh=ss.getSheetByName('回報索引'),states={};
  if(idxSh&&idxSh.getLastRow()>1){
    idxSh.getRange(2,1,idxSh.getLastRow()-1,6).getValues().forEach(row=>{
      if(!row[0])return;
      const id=String(row[0]);
      if(states[id])throw new Error('回報索引有重複 UUID，請先執行資料健檢：'+id);
      states[id]={type:String(row[1]),status:String(row[4])};
    });
  }
  const change=ss.getSheetByName('修改申請'),pending={};
  if(change&&change.getLastRow()>1)
    change.getRange(2,1,change.getLastRow()-1,9).getValues().forEach(row=>{
      if(String(row[8])==='待審核'&&row[1])pending[String(row[1])]=(pending[String(row[1])]||0)+1;
    });
  const stamp=Utilities.formatDate(new Date(),'Asia/Taipei','yyyy/MM/dd HH:mm:ss');
  Object.keys(TYPES).forEach(type=>{
    const spec=TYPES[type],sh=ss.getSheetByName(spec.sheet);
    if(!sh||sh.getLastRow()<2)return;
    const header=headersFor_(type),observed=sh.getRange(1,1,1,header.length).getDisplayValues()[0];
    if(header.some((h,i)=>h!==observed[i]))
      throw new Error('「'+spec.sheet+'」欄位格式不符，停止結算以免計錯');
    sh.getRange(2,1,sh.getLastRow()-1,header.length).getValues().forEach(row=>{
      if(!row[0])return;
      const id=String(row[0]);
      if(seenIds[id])throw new Error('原始回報 UUID 重複，停止結算：'+id);
      seenIds[id]=true;
      if(states[id]&&states[id].type!==type)throw new Error('回報索引類別與原始成績不一致：'+id);
      const obj={};spec.fields.forEach((field,i)=>obj[field]=row[i+3]);
      const c=String(obj.country||'').trim(),u=String(obj.unit||'').trim(),
        d=String(obj.district||'').trim();
      if(!c||!u||!d)throw new Error('回報缺少國別、單位或區域，停止結算：'+id);
      const areaKey=JSON.stringify([c,u,d]);
      // 顯示欄不含國別/單位時，重名區域不應看似只有一個區。
      if(districtOwners[d]&&districtOwners[d]!==areaKey)
        throw new Error('「'+d+'」在不同國別或單位重複，無法安全產生不含國別／單位的結算表；請管理員統一區域名稱');
      districtOwners[d]=areaKey;
      if(!regions[areaKey])regions[areaKey]=[d,c,u,0,0,0,0,0,0,0,0,stamp];
      const area=regions[areaKey];
      if(pending[id])area[10]+=pending[id];
      if(states[id]&&states[id].status==='已取消')return;
      // 未有回報索引的歷史資料暫按有效統計，須在正式封存前人工健檢。
      const count=Number(obj.count||0);
      if(spec.fields.includes('count')&&(!Number.isInteger(count)||count<0))
        throw new Error('回報人數不是有效整數，停止結算：'+id);
      const metric={personal:3,family:4,public:5,personalClass:6,publicClass:7,temple:8}[type];
      const score=['personalClass','temple'].includes(type)?1:count;
      area[metric]+=score;
      area[9]+=1;
      (SETTLEMENT_ROLE_FIELDS[type]||[]).forEach(role=>{
        const person=String(obj[role[0]]||'').trim();
        if(!person)return;
        const pkey=JSON.stringify([c,u,d,person,role[1]]);
        if(!people[pkey])people[pkey]=[c,u,d,person,role[1],0,0,0,0,0,0,0,stamp];
        const out=people[pkey];
        out[role[2]]+=score;
        out[11]+=1;
      });
    });
  });
  const sort=(a,b)=>{
    for(let i=0;i<Math.min(a.length,b.length,5);i++){
      const n=String(a[i]).localeCompare(String(b[i]),'zh-Hant');
      if(n)return n;
    }
    return 0;
  };
  return {regional:Object.keys(regions).map(k=>regions[k]).sort(sort).map(row=>[row[0]].concat(row.slice(3))),
    personal:Object.keys(people).map(k=>people[k]).sort(sort).map(row=>row.slice(2)),
    pendingCount:Object.values(pending).reduce((n,v)=>n+v,0),
    missingIndex:Object.keys(seenIds).filter(id=>!states[id]).length,stamp};
}
// 舊版兩張結算表是系統衍生資料；只允許這兩種已知表頭升級。
// 更換欄位前備份既有結算分頁，原始六類回報分頁永不動。
const PREVIOUS_REGION_HEADERS=['國別','單位','區域','個人渡眾人數','家堂渡眾人數','公堂渡眾人數',
  '個人成全進階班回報筆數','公堂進階班達標班員人數','成全安堂完成筆數',
  '有效回報筆數','待審核申請筆數','結算時間'];
const PREVIOUS_PERSON_HEADERS=['國別','單位','區域','姓名','角色',
  '個人渡眾人數','個人成全進階班回報筆數','成全安堂完成筆數',
  '家堂渡眾人數（壇主記錄）','公堂渡眾人數（回報人記錄）',
  '公堂進階班達標人數（回報人記錄）','個人成全進階班回報筆數（班員記錄）',
  '有效相關回報筆數','結算時間'];
const PREVIOUS_V2_REGION_HEADERS=['區域','國別','單位','個人渡眾人數','家堂渡眾人數','公堂渡眾人數',
  '個人成全進階班回報筆數','公堂進階班達標班員人數','成全安堂完成筆數',
  '有效回報筆數','待審核申請筆數','結算時間'];
const PREVIOUS_V2_PERSON_HEADERS=['國別','單位','區域','姓名','角色',
  '個人渡眾人數','個人成全進階班回報筆數','成全安堂完成筆數',
  '家堂渡眾人數（壇主記錄）','公堂渡眾人數（回報人記錄）',
  '公堂進階班達標人數（回報人記錄）','有效相關回報筆數','結算時間'];
function settlementSchema_(ss,title,expected){
  const sh=ss.getSheetByName(title);
  if(!sh)return {sheet:null,upgrade:false};
  const width=Math.max(sh.getLastColumn(),expected.length);
  const got=sh.getRange(1,1,1,width).getDisplayValues()[0];
  const matches=(headers)=>headers.every((x,i)=>String(got[i]||'')===x)&&
    got.slice(headers.length).every(x=>String(x||'')==='');
  if(matches(expected))return {sheet:sh,upgrade:false};
  const legacy=title==='區域成績結算'?
    [PREVIOUS_V2_REGION_HEADERS,PREVIOUS_REGION_HEADERS]:
    [PREVIOUS_V2_PERSON_HEADERS,PREVIOUS_PERSON_HEADERS];
  if(legacy.some(headers=>matches(headers)))return {sheet:sh,upgrade:true};
  throw new Error('「'+title+'」表頭不符合目前或前一版結算格式，為保護資料停止更新');
}
function backupSettlementBeforeUpgrade_(sh){
  const ss=sh.getParent();
  const stamp=Utilities.formatDate(new Date(),'Asia/Taipei','MMdd_HHmmss');
  const name=('備份_'+sh.getName()+'_'+stamp).slice(0,99);
  if(ss.getSheetByName(name))throw new Error('已存在同名結算備份，請稍後重試');
  const backup=sh.copyTo(ss).setName(name);
  backup.hideSheet();
}
function settlementSheet_(ss,title,headers,rows,checked){
  let sh=checked.sheet;
  if(!sh)sh=ss.insertSheet(title);
  // 僅清除系統衍生結算表；前版的「班員」舊欄也須清空避免殘留。
  const previousWidth=Math.max(sh.getLastColumn(),headers.length);
  const needed=rows.length+1,width=headers.length;
  if(sh.getMaxColumns()<width)sh.insertColumnsAfter(sh.getMaxColumns(),width-sh.getMaxColumns());
  if(sh.getMaxRows()<needed)sh.insertRowsAfter(sh.getMaxRows(),needed-sh.getMaxRows());
  sh.getRange(1,1,sh.getMaxRows(),previousWidth).clearContent();
  sh.getRange(1,1,1,width).setValues([headers]);
  if(rows.length)sh.getRange(2,1,rows.length,width).setValues(rows.map(row=>row.map(value=>
    typeof value==='string'?safeRegionCell_(value):value)));
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,width).setFontWeight('bold').setBackground('#e6f1e8');
  if(sh.getFilter())sh.getFilter().remove();
  sh.getRange(1,1,sh.getMaxRows(),width).createFilter();
}
function recalculateScores(options){
  const admin=requireAdmin_();
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const ss=sheet_(),data=settlementRows_();
    // 先驗證兩個結算分頁，確保非系統資料不會被覆寫。
    const personalSchema=settlementSchema_(ss,'個人成績結算',SETTLEMENT_PERSON_HEADERS);
    const regionSchema=settlementSchema_(ss,'區域成績結算',SETTLEMENT_REGION_HEADERS);
    // 升級既有衍生表前保留快照，不影響原始成績與照片。
    if(personalSchema.upgrade)backupSettlementBeforeUpgrade_(personalSchema.sheet);
    if(regionSchema.upgrade)backupSettlementBeforeUpgrade_(regionSchema.sheet);
    settlementSheet_(ss,'個人成績結算',SETTLEMENT_PERSON_HEADERS,data.personal,personalSchema);
    settlementSheet_(ss,'區域成績結算',SETTLEMENT_REGION_HEADERS,data.regional,regionSchema);
    if(!options||!options.silent)
      auditSheet_().appendRow([new Date(),'settlement_refresh','','',admin,'',JSON.stringify({
        people:data.personal.length,regions:data.regional.length,pending:data.pendingCount,
        legacyWithoutIndex:data.missingIndex
      }),'重新結算個人與區域成績；原始紀錄未更動']);
    return {ok:true,personalRows:data.personal.length,regionalRows:data.regional.length,
      pendingCount:data.pendingCount,missingIndex:data.missingIndex,updatedAt:data.stamp};
  }finally{lock.releaseLock();}
}
function adminRecalculateScoresPrompt(){
  requireAdmin_();
  const ui=SpreadsheetApp.getUi();
  const yes=ui.alert('重新結算成績',
    '將建立／更新「個人成績結算」及「區域成績結算」兩個衍生分頁。已核准取消的回報不計入；未完成審核的申請仍以目前原始成績統計。原始成績、照片、索引不會刪除。確定嗎？',
    ui.ButtonSet.YES_NO);
  if(yes!==ui.Button.YES)return {ok:false,cancelled:true};
  try{
    const result=recalculateScores();
    ui.alert('結算完成',
      '個人角色明細：'+result.personalRows+' 列\n區域：'+result.regionalRows+
      ' 區\n待審核申請：'+result.pendingCount+
      ' 筆\n無索引舊回報：'+result.missingIndex+
      ' 筆\n更新：'+result.updatedAt+
      '\n\n請至「個人成績結算」及「區域成績結算」查看。',
      ui.ButtonSet.OK);
    return result;
  }catch(err){
    ui.alert('結算未完成',String(err.message||err)+'\n請檢查成績分頁，不要手動刪除舊資料。',ui.ButtonSet.OK);
    throw err;
  }
}


/**
 * Run once as a spreadsheet administrator to authorize an installable time
 * trigger. A scheduled trigger runs as its installer (never as an anonymous
 * reporter); no publicly callable web action is provided.
 */
function installAutoSettlement(){
  requireAdmin_();
  const props=PropertiesService.getScriptProperties();
  // recalculateScores() already acquires ScriptLock; avoid nesting the same lock.
  const active=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='autoRecalculateScores_');
  {
    if(active.length)return {ok:true,created:false,count:active.length,message:'已存在自動結算排程，無須重複建立'};
    // First calculate and validate both tables, then create a trigger.
    recalculateScores({silent:true});
    ScriptApp.newTrigger('autoRecalculateScores_').timeBased().everyMinutes(5).create();
    props.setProperty('AUTO_SETTLEMENT_INTERVAL_MIN','5');
    return {ok:true,created:true,count:1,message:'已啟用每5分鐘自動更新個人／區域成績'};
  }
}
function autoRecalculateScores_(){
  // Scheduled handler, not a public web action. Any failure surfaces in Apps
  // Script's trigger execution logs / failure emails; never discard the data.
  return recalculateScores({silent:true});
}
function adminInstallAutoSettlementPrompt(){
  requireAdmin_();
  const ui=SpreadsheetApp.getUi();
  const choice=ui.alert('啟用自動成績結算',
    '將授權 Apps Script 每5分鐘重新計算「個人成績結算」和「區域成績結算」，無須手動按重新結算。首次啟用會先執行結算。\n不會更動原始成績或照片。\n確定啟用嗎？',
    ui.ButtonSet.YES_NO);
  if(choice!==ui.Button.YES)return {ok:false,cancelled:true};
  try{
    const result=installAutoSettlement();
    ui.alert('自動結算設定',result.message+
      '\n未來的新回報、核准修改、核准取消會在下一次排程更新（約5分鐘內，實際以Google排程為準）。',
      ui.ButtonSet.OK);
    return result;
  }catch(err){
    ui.alert('啟用失敗',String(err.message||err)+
      '\n請確認執行帳號有 Apps Script 觸發器及原始試算表權限。',ui.ButtonSet.OK);
    throw err;
  }
}
