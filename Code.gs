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
    const payload = JSON.parse(e.postData.contents);
    return json_(submit_(payload));
  } catch (err) {
    return json_({ok:false,error:String(err.message || err)});
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
function submit_(p) {
  const spec = validate_(p);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = sheet_();
    let sh = ss.getSheetByName(spec.sheet);
    if (!sh) sh = ss.insertSheet(spec.sheet);
    ensureSchema_(sh,p.type);
    const id = Utilities.getUuid();
    const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    const urls = [];
    for (const f of (p.files || [])) {
      const bytes = Utilities.base64Decode(f.data);
      if (bytes.length > CONFIG.MAX_FILE_BYTES) throw new Error('單張照片不可超過5MB');
      const ext = ({'image/jpeg':'jpg','image/png':'png','image/webp':'webp','image/heic':'heic','image/heif':'heif'})[f.mimeType];
      const blob = Utilities.newBlob(bytes, f.mimeType, id + '_' + (urls.length + 1) + '.' + ext);
      urls.push(folder.createFile(blob).getUrl());
    }
    sh.appendRow(rowFor_(p.type,p,id,urls));
    return {ok:true,id};
  } finally { lock.releaseLock(); }
}
function setupSheets() {
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
