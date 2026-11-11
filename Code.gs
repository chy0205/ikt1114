const CONFIG = {
  SPREADSHEET_ID: '1ZglfWKES6sHWcn2vi9JVHXHVJImPqzcWGb_rYHaj0Uk',
  DRIVE_FOLDER_ID: '請填入GoogleDrive資料夾ID',
  MAX_FILES: 5,
  MAX_FILE_BYTES: 5 * 1024 * 1024,
};
const TYPES = {
  personal: { sheet: '個人渡眾', fields: ['country','unit','district','guide','eventDate','count'], photo: true },
  family: { sheet: '家堂渡眾', fields: ['country','unit','district','temple','host','eventDate','count'], photo: false },
  public: { sheet: '公堂渡眾', fields: ['country','unit','temple','eventDate','count'], photo: false },
  personalClass: { sheet: '個人成全進階班', fields: ['country','unit','district','classType','mentor','student','eventDate','session'], photo: true },
  publicClass: { sheet: '公堂進階班', fields: ['country','unit','temple','classType','count'], photo: true },
  temple: { sheet: '成全安堂完成', fields: ['country','unit','district','mentor','host','eventDate'], photo: true },
};
const HEADERS = ['回報編號','送出時間','回報類別','國別','單位','區別','引師姓名','佛堂或公堂堂號','壇主姓名','成全者姓名','班員姓名','進階班類別','辦道／上課／安堂日期','渡眾／達標人數','上課次數','照片連結'];
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
  if (spec.photo && !files.length) throw new Error('此回報類別必須上傳照片');
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
    if (!sh) { sh = ss.insertSheet(spec.sheet); sh.appendRow(HEADERS); }
    if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
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
    sh.appendRow([id,new Date(),spec.sheet,p.country,p.unit,p.district,p.guide,p.temple,p.host,p.mentor,p.student,p.classType,p.eventDate,p.count,p.session,urls.join('\n')].map(safeText_));
    return {ok:true,id};
  } finally { lock.releaseLock(); }
}
function setupSheets() {
  const ss = sheet_();
  let options = ss.getSheetByName('選項設定');
  if (!options) options = ss.insertSheet('選項設定');
  if (options.getLastRow() === 0) options.appendRow(['國別','單位','區別']);
  Object.values(TYPES).forEach(spec => {
    let sh = ss.getSheetByName(spec.sheet);
    if (!sh) sh = ss.insertSheet(spec.sheet);
    if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  });
}