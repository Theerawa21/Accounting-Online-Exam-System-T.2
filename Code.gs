const SPREADSHEET_ID = '1OQDzQ9FiMKTtI6v7wQpEWusjf2gJG4DEWV6aqsJuKNs';
const SHEET_NAME = 'ชีต1';
const STUDENT_SHEET_NAME = 'นักเรียน ม.4/3';
const TEACHER_PIN = PropertiesService.getScriptProperties().getProperty('TEACHER_PIN') || '';
const HEADERS = [
  'วันที่และเวลาที่ส่ง', 'รหัสนักเรียน', 'ชื่อ–นามสกุล', 'ห้องเรียน',
  'ชุดข้อสอบ', 'คะแนน', 'ร้อยละ', 'เวลาที่ใช้ (วินาที)',
  'ส่งอัตโนมัติ', 'จำนวนข้อที่ตอบ', 'สถานะ'
];

/** รับข้อมูลจากหน้าเว็บที่โฮสต์แยกจาก Apps Script */
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action !== 'save') throw new Error('Unsupported action');
    return jsonOutput_(saveResult(body.record));
  } catch (error) {
    return jsonOutput_({ ok: false, error: error.message });
  }
}

/** อ่านผลคะแนนผ่าน JSON/JSONP สำหรับหน้าครู */
function doGet(e) {
  const params = (e && e.parameter) || {};
  let payload;
  try {
    if (params.action === 'list') payload = { ok: true, results: getResults(params.pin) };
    else if (params.action === 'lookup') payload = lookupStudent(params.sid);
    else payload = { ok: true, service: 'online-accounting-exam', sheet: SHEET_NAME };
  } catch (error) {
    payload = { ok: false, error: error.message, results: [] };
  }
  if (params.callback && /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(params.callback)) {
    return ContentService
      .createTextOutput(params.callback + '(' + JSON.stringify(payload) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonOutput_(payload);
}

/** เรียกโดย google.script.run หรือ doPost */
function saveResult(record) {
  const clean = validateRecord_(record);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getSheet_();
    const values = [
      new Date(clean.submitted), clean.sid, safeText_(clean.name), clean.room,
      clean.set, clean.score, clean.percent, clean.duration,
      clean.auto ? 'ใช่' : 'ไม่ใช่', clean.answeredCount, 'ส่งแล้ว'
    ];
    const lastRow = sheet.getLastRow();
    let targetRow = lastRow + 1;
    if (lastRow > 1) {
      const ids = sheet.getRange(2, 2, lastRow - 1, 1).getDisplayValues();
      const found = ids.findIndex(row => String(row[0]).trim() === clean.sid);
      if (found >= 0) targetRow = found + 2;
    }
    sheet.getRange(targetRow, 2).setNumberFormat('@');
    sheet.getRange(targetRow, 1, 1, HEADERS.length).setValues([values]);
    sheet.getRange(targetRow, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
    sheet.getRange(targetRow, 6, 1, 2).setNumberFormats([['0', '0.00']]);
    return { ok: true, row: targetRow, savedAt: new Date().toISOString() };
  } finally {
    lock.releaseLock();
  }
}

/** เรียกโดย google.script.run หรือ doGet */
function getResults(pin) {
  if (!TEACHER_PIN || String(pin || '') !== TEACHER_PIN) throw new Error('รหัสสำหรับครูไม่ถูกต้อง');
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues()
    .filter(row => row[1])
    .map(row => ({
      submitted: row[0] instanceof Date ? row[0].toISOString() : String(row[0]),
      sid: String(row[1]), name: String(row[2]), room: String(row[3]),
      set: Number(row[4]), score: Number(row[5]), percent: Number(row[6]),
      duration: Number(row[7]), auto: String(row[8]) === 'ใช่',
      answeredCount: Number(row[9]), status: String(row[10])
    }))
    .sort((a, b) => new Date(b.submitted) - new Date(a.submitted));
}

/** ค้นหาข้อมูลนักเรียนจากแท็บรายชื่อโดยไม่ส่งรายชื่อทั้งหมดออกไปยังหน้าเว็บ */
function lookupStudent(sid) {
  const cleanSid = String(sid || '').trim();
  if (!/^\d{5}$/.test(cleanSid)) return { ok: true, found: false };
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(STUDENT_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, found: false };
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getDisplayValues();
  const row = rows.find(values => String(values[0]).trim() === cleanSid);
  if (!row) return { ok: true, found: false };
  return {
    ok: true,
    found: true,
    student: { sid: cleanSid, name: String(row[1]).trim(), room: String(row[2]).trim() }
  };
}

function getSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  spreadsheet.setSpreadsheetTimeZone('Asia/Bangkok');
  const sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
  const currentHeader = sheet.getRange(1, 1, 1, HEADERS.length).getDisplayValues()[0];
  if (currentHeader.join('').trim() === '') {
    const header = sheet.getRange(1, 1, 1, HEADERS.length);
    header.setValues([HEADERS])
      .setBackground('#062f67')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.setRowHeight(1, 34);
    [145, 120, 220, 95, 90, 80, 85, 135, 115, 130, 90]
      .forEach((width, index) => sheet.setColumnWidth(index + 1, width));
    sheet.getRange('B:B').setNumberFormat('@');
  }
  return sheet;
}

function validateRecord_(record) {
  if (!record || typeof record !== 'object') throw new Error('ไม่พบข้อมูลผลสอบ');
  const result = {
    sid: String(record.sid || '').trim(),
    name: String(record.name || '').trim().slice(0, 120),
    room: String(record.room || '').trim(),
    set: Number(record.set), score: Number(record.score), percent: Number(record.percent),
    duration: Math.max(0, Number(record.duration) || 0),
    submitted: record.submitted || new Date().toISOString(),
    auto: Boolean(record.auto), answeredCount: Number(record.answeredCount) || 0
  };
  if (!/^\d{4,13}$/.test(result.sid)) throw new Error('รหัสนักเรียนไม่ถูกต้อง');
  const rosterResult = lookupStudent(result.sid);
  if (!rosterResult.found) throw new Error('ไม่พบรหัสนักเรียนในรายชื่อ');
  result.name = rosterResult.student.name;
  result.room = rosterResult.student.room;
  if (!Number.isInteger(result.score) || result.score < 0 || result.score > 40) throw new Error('คะแนนไม่ถูกต้อง');
  if (!Number.isFinite(result.percent) || result.percent < 0 || result.percent > 100) throw new Error('ร้อยละไม่ถูกต้อง');
  if (!Number.isInteger(result.set) || result.set < 1 || result.set > 3) throw new Error('ชุดข้อสอบไม่ถูกต้อง');
  return result;
}

function safeText_(value) {
  const text = String(value || '');
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function jsonOutput_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

