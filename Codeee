const APP = Object.freeze({
  sheetName: 'Stroke_Assessments_Demo',
  maxRecords: 1000,
  headers: ['Record_ID','Created_At','Assessment_Date','Reference_Code','Assessor','Visit','NIHSS_Total','Barthel_Total','NIHSS_JSON','Barthel_JSON','Created_By']
});

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Stroke & ADL Assessment System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Run once from the Apps Script editor to create/check the sheet and its columns. */
function SetupSheet() {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getSpreadsheet_();
    let sh = ss.getSheetByName(APP.sheetName);
    if (!sh) sh = ss.insertSheet(APP.sheetName);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, APP.headers.length).setValues([APP.headers]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, APP.headers.length).setFontWeight('bold').setBackground('#d9e1f2');
      sh.getRange('A:K').setNumberFormat('@');
      sh.autoResizeColumns(1, APP.headers.length);
    } else {
      assertHeaders_(sh);
    }
    return { ok: true, message: 'สร้าง/ตรวจสอบชีตและคอลัมน์เรียบร้อย: ' + APP.sheetName };
  } finally { lock.releaseLock(); }
}

function getAppData() {
  const sh = getSheet_();
  const count = Math.max(0, sh.getLastRow() - 1);
  const values = count ? sh.getRange(2, 1, count, APP.headers.length).getDisplayValues() : [];
  const records = values.filter(r => r[0]).map(r => ({
    id:r[0], createdAt:r[1], date:r[2], patient:r[3], assessor:r[4], visit:Number(r[5]),
    nihss:Number(r[6]), barthel:Number(r[7]), n:parseJson_(r[8]), b:parseJson_(r[9])
  })).sort((a,b) => b.date.localeCompare(a.date) || b.createdAt.localeCompare(a.createdAt));
  return { ok:true, records };
}

function saveAssessment(payload) {
  const data = validatePayload_(payload);
  const email = Session.getActiveUser().getEmail() || 'anonymous';
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = getSheet_();
    const count = Math.max(0, sh.getLastRow() - 1);
    if (count) {
      const match = sh.getRange(2,1,count,1).createTextFinder(data.id).matchEntireCell(true).findNext();
      if (match) return { ok:true, id:data.id, duplicate:true };
    }
    if (count >= APP.maxRecords) throw new Error('ระบบสาธิตรองรับสูงสุด ' + APP.maxRecords + ' รายการ');
    const row = [textCell_(data.id),textCell_(new Date().toISOString()),textCell_(data.date),textCell_(data.patient),textCell_(data.assessor),textCell_(String(data.visit)),data.nihss,data.barthel,JSON.stringify(data.n),JSON.stringify(data.b),textCell_(email)];
    sh.getRange(sh.getLastRow()+1,1,1,row.length).setValues([row]);
    SpreadsheetApp.flush();
    return { ok:true, id:data.id, duplicate:false };
  } finally { lock.releaseLock(); }
}

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('ตั้งค่า Script Property SPREADSHEET_ID หรือผูกสคริปต์กับ Google Sheet ก่อน');
  return ss;
}
function getSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(APP.sheetName);
  if (!sh) {
    sh = ss.insertSheet(APP.sheetName);
    sh.getRange(1,1,1,APP.headers.length).setValues([APP.headers]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,APP.headers.length).setFontWeight('bold').setBackground('#d9e1f2');
    sh.getRange('A:K').setNumberFormat('@');
  }
  assertHeaders_(sh);
  return sh;
}
function assertHeaders_(sh) {
  const headers = sh.getRange(1,1,1,APP.headers.length).getDisplayValues()[0];
  if (!APP.headers.every((header,i) => headers[i] === header)) throw new Error('หัวคอลัมน์ในชีตไม่ตรงกับโครงสร้างระบบ');
}
function validatePayload_(p) {
  if (!p || typeof p !== 'object') throw new Error('ข้อมูลไม่ถูกต้อง');
  const data = {
    id:String(p.id || 'REC-'+Date.now()), date:String(p.date || ''), patient:String(p.patient || '').trim(),
    assessor:String(p.assessor || '').trim(), visit:Number(p.visit), nihss:Number(p.nihss), barthel:Number(p.barthel),
    n:p.n && typeof p.n === 'object' ? p.n : {}, b:p.b && typeof p.b === 'object' ? p.b : {}
  };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(data.date) || !data.patient || !data.assessor || !Number.isInteger(data.visit) || data.visit < 1 || !Number.isFinite(data.nihss) || data.nihss < 0 || data.nihss > 42 || !Number.isFinite(data.barthel) || data.barthel < 0 || data.barthel > 100) throw new Error('กรุณาตรวจสอบข้อมูลที่จำเป็นและคะแนนรวม');
  return data;
}
function parseJson_(value) { try { return JSON.parse(value || '{}'); } catch (e) { return {}; } }
function textCell_(value) { return "'" + String(value); }

