/**
 * สภานักเรียนโรงเรียนบ้านไผ่ — เชื่อม Google Sheet เป็นหลังบ้าน
 *
 * วิธีติดตั้ง (ทำครั้งเดียว / อัปเดต):
 * 1) เปิดชีต: https://docs.google.com/spreadsheets/d/1eI-A2RUYkDu2D6v952ikNwKzTVt6TRLFJ_6Wr96Tw1o/edit
 * 2) เมนู Extensions → Apps Script
 * 3) ลบโค้ดเดิม แล้ววางไฟล์นี้ทั้งหมด → Save
 * 4) กด Run → เลือกฟังก์ชัน setupSheets → Allow permissions
 * 5) Deploy → New deployment (หรือ Manage deployments → Edit → Version: New)
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6) Copy Web app URL ไปวางในระบบ (ถ้า URL เปลี่ยน)
 */

var SPREADSHEET_ID = '1eI-A2RUYkDu2D6v952ikNwKzTVt6TRLFJ_6Wr96Tw1o';

var BEHAVIOR_SHEET = 'รายงานความประพฤติ';
var TASK_SHEET = 'มอบหมายงานสภา';
var LOG_SHEET = 'บันทึกงาน';
var LF_SHEET = 'ของหายได้คืน';
var MSG_SHEET = 'แจ้งเรื่อง';

var BEHAVIOR_HEADERS = [
  'เวลาบันทึก',
  'รหัสรายการ',
  'ประเภทหลัก',
  'รหัสนักเรียน',
  'ทะเบียนรถ',
  'ประเภทย่อย',
  'รายละเอียด',
  'ผู้บันทึก',
  'วันที่เกิดเหตุ',
  'สถานะ',
  'ผู้รับรอง',
  'วันที่รับรอง',
];

var TASK_HEADERS = [
  'เวลาบันทึก',
  'รหัสงาน',
  'ชื่องาน',
  'ฝ่าย',
  'ผู้รับมอบหมาย',
  'กำหนดส่ง',
  'ความสำคัญ',
  'รายละเอียด',
  'สถานะ',
  'ผู้มอบหมาย',
  'วันที่มอบหมาย',
];

var LOG_HEADERS = [
  'เวลาบันทึก',
  'รหัสรายการ',
  'ประเภท',
  'หัวข้อ',
  'รายละเอียด',
  'ผู้บันทึก',
  'วันที่',
  'สถานะ',
];

var LF_HEADERS = [
  'เวลาบันทึก',
  'รหัสรายการ',
  'ประเภท',
  'หมวดหมู่',
  'ชื่อรายการ',
  'สถานที่',
  'วันที่',
  'รายละเอียด',
  'ผู้แจ้ง',
  'รหัสผู้แจ้ง',
  'ช่องทางติดต่อ',
  'สถานะ',
  'ผู้ขอรับ',
  'รหัสผู้ขอรับ',
  'หมายเหตุขอรับ',
  'ผู้ยืนยันคืน',
  'วันที่คืน',
];

var MSG_HEADERS = [
  'เวลาบันทึก',
  'รหัสรายการ',
  'หัวข้อ',
  'รายละเอียด',
  'ผู้ส่ง',
  'รหัสผู้ส่ง',
  'วันที่',
  'สถานะ',
  'คำตอบ',
  'ผู้ตอบ',
  'วันที่ตอบ',
];

// ชื่อเก่า (รองรับโค้ดเดิม)
var MAIN_SHEET = BEHAVIOR_SHEET;
var HEADERS = BEHAVIOR_HEADERS;

function styleHeader_(sh, headers) {
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#5e2d96')
    .setFontColor('#ffffff');
}

function ensureSheet_(name, headers) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  var lastCol = Math.max(sh.getLastColumn(), 1);
  var existing = sh.getRange(1, 1, 1, Math.max(headers.length, lastCol)).getDisplayValues()[0];
  var needsHeader = sh.getLastRow() === 0 || String(existing[0] || '') !== headers[0];

  if (needsHeader && sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeader_(sh, headers);
  } else if (needsHeader && sh.getLastRow() > 0 && String(existing[1] || '') === '') {
    // แถวแรกว่างจริง ๆ — ใส่หัวใหม่
    sh.clear();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeader_(sh, headers);
  } else if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeader_(sh, headers);
  }

  return sh;
}

function setupSheets() {
  ensureSheet_(BEHAVIOR_SHEET, BEHAVIOR_HEADERS);
  ensureSheet_(TASK_SHEET, TASK_HEADERS);
  ensureSheet_(LOG_SHEET, LOG_HEADERS);
  ensureSheet_(LF_SHEET, LF_HEADERS);
  ensureSheet_(MSG_SHEET, MSG_HEADERS);

  var keep = {};
  keep[BEHAVIOR_SHEET] = true;
  keep[TASK_SHEET] = true;
  keep[LOG_SHEET] = true;
  keep[LF_SHEET] = true;
  keep[MSG_SHEET] = true;

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ss.getSheets();
  if (sheets.length > 5) {
    sheets.forEach(function (s) {
      var n = s.getName();
      if (!keep[n] && s.getLastRow() === 0) {
        try {
          ss.deleteSheet(s);
        } catch (e) {}
      }
    });
  }
  return (
    'พร้อมใช้งาน: ชีต "' +
    BEHAVIOR_SHEET +
    '", "' +
    TASK_SHEET +
    '", "' +
    LOG_SHEET +
    '", "' +
    LF_SHEET +
    '", "' +
    MSG_SHEET +
    '"'
  );
}

function getBehaviorSheet_() {
  return ensureSheet_(BEHAVIOR_SHEET, BEHAVIOR_HEADERS);
}

function getTaskSheet_() {
  return ensureSheet_(TASK_SHEET, TASK_HEADERS);
}

function getLogSheet_() {
  return ensureSheet_(LOG_SHEET, LOG_HEADERS);
}

function getLfSheet_() {
  return ensureSheet_(LF_SHEET, LF_HEADERS);
}

function getMsgSheet_() {
  return ensureSheet_(MSG_SHEET, MSG_HEADERS);
}

function getMainSheet_() {
  return getBehaviorSheet_();
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function typeLabel_(type) {
  if (type === 'helmet') return 'ไม่สวมหมวกกันน็อค';
  if (type === 'late') return 'กักแถวสาย';
  if (type === 'misconduct') return 'พฤติกรรมไม่ถูกต้อง';
  return type || '';
}

function priorityLabel_(p) {
  if (p === 'high' || p === 'สูง') return 'สูง';
  if (p === 'low' || p === 'ต่ำ') return 'ต่ำ';
  return 'กลาง';
}

function rowToObj_(row) {
  return {
    timestamp: row[0] || '',
    id: String(row[1] || ''),
    type: String(row[2] || ''),
    typeKey: '',
    studentId: String(row[3] || ''),
    plate: String(row[4] || ''),
    category: String(row[5] || ''),
    detail: String(row[6] || ''),
    by: String(row[7] || ''),
    date: String(row[8] || ''),
    status: String(row[9] || 'pending'),
    reviewedBy: String(row[10] || ''),
    reviewedAt: String(row[11] || ''),
  };
}

function labelToType_(label) {
  if (label === 'ไม่สวมหมวกกันน็อค') return 'helmet';
  if (label === 'กักแถวสาย') return 'late';
  if (label === 'พฤติกรรมไม่ถูกต้อง') return 'misconduct';
  if (label === 'helmet' || label === 'late' || label === 'misconduct') return label;
  return 'misconduct';
}

function listReports_() {
  var sh = getBehaviorSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last, BEHAVIOR_HEADERS.length).getDisplayValues();
  return values
    .filter(function (r) {
      return r[1];
    })
    .map(function (r) {
      var o = rowToObj_(r);
      o.typeKey = labelToType_(o.type);
      o.type = o.typeKey;
      o.note = o.detail;
      o.recorder = o.by;
      return o;
    });
}

function addReport_(payload) {
  var sh = getBehaviorSheet_();
  var id = payload.id || 'r' + Date.now().toString(36);
  var typeKey = payload.type || 'misconduct';
  var row = [
    new Date(),
    id,
    typeLabel_(typeKey),
    payload.studentId || '',
    payload.plate || '',
    payload.category || '',
    payload.detail || payload.note || '',
    payload.by || payload.recorder || '',
    payload.date || Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd'),
    payload.status || 'pending',
    payload.reviewedBy || '',
    payload.reviewedAt || '',
  ];
  sh.appendRow(row);
  return { ok: true, id: id };
}

function updateStatus_(payload) {
  var sh = getBehaviorSheet_();
  var last = sh.getLastRow();
  if (last < 2) return { ok: false, error: 'ไม่พบข้อมูล' };
  var ids = sh.getRange(2, 2, last, 2).getDisplayValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(payload.id)) {
      var rowNum = i + 2;
      sh.getRange(rowNum, 10).setValue(payload.status || 'reviewed');
      sh.getRange(rowNum, 11).setValue(payload.reviewedBy || '');
      sh.getRange(rowNum, 12).setValue(
        payload.reviewedAt || Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd')
      );
      return { ok: true };
    }
  }
  return { ok: false, error: 'ไม่พบรหัสรายการ' };
}

function taskRowToObj_(row) {
  return {
    timestamp: row[0] || '',
    id: String(row[1] || ''),
    title: String(row[2] || ''),
    dept: String(row[3] || ''),
    assignee: String(row[4] || ''),
    due: String(row[5] || ''),
    priority: String(row[6] || 'กลาง'),
    note: String(row[7] || ''),
    status: String(row[8] || 'todo'),
    createdBy: String(row[9] || ''),
    createdAt: String(row[10] || ''),
  };
}

function listTasks_() {
  var sh = getTaskSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last, TASK_HEADERS.length).getDisplayValues();
  return values
    .filter(function (r) {
      return r[1];
    })
    .map(taskRowToObj_);
}

function addTask_(payload) {
  var sh = getTaskSheet_();
  var id = payload.id || 't' + Date.now().toString(36);
  var row = [
    new Date(),
    id,
    payload.title || '',
    payload.dept || '',
    payload.assignee || '',
    payload.due || '',
    priorityLabel_(payload.priority),
    payload.note || '',
    payload.status || 'todo',
    payload.createdBy || '',
    payload.createdAt || Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd'),
  ];
  sh.appendRow(row);
  return { ok: true, id: id };
}

function deleteTask_(payload) {
  var sh = getTaskSheet_();
  var last = sh.getLastRow();
  if (last < 2) return { ok: false, error: 'ไม่พบข้อมูล' };
  var ids = sh.getRange(2, 2, last, 2).getDisplayValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(payload.id)) {
      sh.deleteRow(i + 2);
      return { ok: true };
    }
  }
  return { ok: false, error: 'ไม่พบรหัสงาน' };
}

function addLog_(payload) {
  var sh = getLogSheet_();
  var id = payload.id || 'i' + Date.now().toString(36);
  var row = [
    new Date(),
    id,
    payload.type || 'บันทึกงาน',
    payload.title || '',
    payload.detail || '',
    payload.by || '',
    payload.date || Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd'),
    payload.status || 'open',
  ];
  sh.appendRow(row);
  return { ok: true, id: id };
}

function deleteLog_(payload) {
  var sh = getLogSheet_();
  var last = sh.getLastRow();
  if (last < 2) return { ok: false, error: 'ไม่พบข้อมูล' };
  var ids = sh.getRange(2, 2, last, 2).getDisplayValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(payload.id)) {
      sh.deleteRow(i + 2);
      return { ok: true };
    }
  }
  return { ok: false, error: 'ไม่พบรหัสรายการ' };
}

function listLogs_() {
  var sh = getLogSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last, LOG_HEADERS.length).getDisplayValues();
  return values
    .filter(function (r) {
      return r[1];
    })
    .map(function (r) {
      return {
        timestamp: r[0] || '',
        id: String(r[1] || ''),
        type: String(r[2] || ''),
        title: String(r[3] || ''),
        detail: String(r[4] || ''),
        by: String(r[5] || ''),
        date: String(r[6] || ''),
        status: String(r[7] || 'open'),
      };
    });
}

function lfTypeLabel_(type) {
  if (type === 'found' || type === 'ของพบ') return 'ของพบ';
  if (type === 'lost' || type === 'แจ้งหาย') return 'แจ้งหาย';
  return type || '';
}

function addLostFound_(payload) {
  var sh = getLfSheet_();
  var id = payload.id || 'lf' + Date.now().toString(36);
  var row = [
    new Date(),
    id,
    lfTypeLabel_(payload.type),
    payload.category || '',
    payload.title || '',
    payload.location || '',
    payload.date || Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd'),
    payload.description || '',
    payload.reporter || '',
    payload.reporterId || '',
    payload.contact || '',
    payload.status || 'open',
    payload.claimBy || '',
    payload.claimId || '',
    payload.claimNote || '',
    payload.returnedBy || '',
    payload.returnedAt || '',
  ];
  sh.appendRow(row);
  return { ok: true, id: id };
}

function findLfRow_(id) {
  var sh = getLfSheet_();
  var last = sh.getLastRow();
  if (last < 2) return null;
  var ids = sh.getRange(2, 2, last, 2).getDisplayValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return null;
}

function updateLostFound_(payload) {
  var rowNum = findLfRow_(payload.id);
  if (!rowNum) return { ok: false, error: 'ไม่พบรหัสรายการ' };
  var sh = getLfSheet_();
  if (payload.status != null) sh.getRange(rowNum, 12).setValue(payload.status);
  if (payload.claimBy != null) sh.getRange(rowNum, 13).setValue(payload.claimBy);
  if (payload.claimId != null) sh.getRange(rowNum, 14).setValue(payload.claimId);
  if (payload.claimNote != null) sh.getRange(rowNum, 15).setValue(payload.claimNote);
  if (payload.returnedBy != null) sh.getRange(rowNum, 16).setValue(payload.returnedBy);
  if (payload.returnedAt != null) sh.getRange(rowNum, 17).setValue(payload.returnedAt);
  return { ok: true };
}

function deleteLostFound_(payload) {
  var rowNum = findLfRow_(payload.id);
  if (!rowNum) return { ok: false, error: 'ไม่พบรหัสรายการ' };
  getLfSheet_().deleteRow(rowNum);
  return { ok: true };
}

function listLostFound_() {
  var sh = getLfSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last, LF_HEADERS.length).getDisplayValues();
  return values
    .filter(function (r) {
      return r[1];
    })
    .map(function (r) {
      var type = String(r[2] || '');
      return {
        timestamp: r[0] || '',
        id: String(r[1] || ''),
        type: type === 'ของพบ' ? 'found' : type === 'แจ้งหาย' ? 'lost' : type,
        category: String(r[3] || ''),
        title: String(r[4] || ''),
        location: String(r[5] || ''),
        date: String(r[6] || ''),
        description: String(r[7] || ''),
        reporter: String(r[8] || ''),
        reporterId: String(r[9] || ''),
        contact: String(r[10] || ''),
        status: String(r[11] || 'open'),
        claimBy: String(r[12] || ''),
        claimId: String(r[13] || ''),
        claimNote: String(r[14] || ''),
        returnedBy: String(r[15] || ''),
        returnedAt: String(r[16] || ''),
      };
    });
}

function addMessage_(payload) {
  var sh = getMsgSheet_();
  var id = payload.id || 'm' + Date.now().toString(36);
  var row = [
    new Date(),
    id,
    payload.topic || '',
    payload.detail || '',
    payload.from || '',
    payload.fromId || '',
    payload.date || Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd'),
    payload.status || 'open',
    payload.reply || '',
    payload.repliedBy || '',
    payload.repliedAt || '',
  ];
  sh.appendRow(row);
  return { ok: true, id: id };
}

function findMsgRow_(id) {
  var sh = getMsgSheet_();
  var last = sh.getLastRow();
  if (last < 2) return null;
  var ids = sh.getRange(2, 2, last, 2).getDisplayValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return null;
}

function updateMessage_(payload) {
  var rowNum = findMsgRow_(payload.id);
  if (!rowNum) return { ok: false, error: 'ไม่พบรหัสรายการ' };
  var sh = getMsgSheet_();
  if (payload.status != null) sh.getRange(rowNum, 8).setValue(payload.status);
  if (payload.reply != null) sh.getRange(rowNum, 9).setValue(payload.reply);
  if (payload.repliedBy != null) sh.getRange(rowNum, 10).setValue(payload.repliedBy);
  if (payload.repliedAt != null) sh.getRange(rowNum, 11).setValue(payload.repliedAt);
  return { ok: true };
}

function listMessages_() {
  var sh = getMsgSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last, MSG_HEADERS.length).getDisplayValues();
  return values
    .filter(function (r) {
      return r[1];
    })
    .map(function (r) {
      return {
        timestamp: r[0] || '',
        id: String(r[1] || ''),
        topic: String(r[2] || ''),
        detail: String(r[3] || ''),
        from: String(r[4] || ''),
        fromId: String(r[5] || ''),
        date: String(r[6] || ''),
        status: String(r[7] || 'open'),
        reply: String(r[8] || ''),
        repliedBy: String(r[9] || ''),
        repliedAt: String(r[10] || ''),
      };
    });
}

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || 'list';
    if (action === 'setup') {
      return jsonOut_({ ok: true, message: setupSheets() });
    }
    if (action === 'list') {
      return jsonOut_({ ok: true, reports: listReports_() });
    }
    if (action === 'listTasks') {
      return jsonOut_({ ok: true, tasks: listTasks_() });
    }
    if (action === 'listLogs') {
      return jsonOut_({ ok: true, logs: listLogs_() });
    }
    if (action === 'listLostFound') {
      return jsonOut_({ ok: true, items: listLostFound_() });
    }
    if (action === 'listMessages') {
      return jsonOut_({ ok: true, messages: listMessages_() });
    }
    return jsonOut_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var raw = (e && e.postData && e.postData.contents) || '{}';
    var data = JSON.parse(raw);
    var action = data.action || 'add';
    if (action === 'setup') {
      return jsonOut_({ ok: true, message: setupSheets() });
    }
    if (action === 'add') {
      return jsonOut_(addReport_(data.report || data));
    }
    if (action === 'updateStatus') {
      return jsonOut_(updateStatus_(data));
    }
    if (action === 'list') {
      return jsonOut_({ ok: true, reports: listReports_() });
    }
    if (action === 'addTask') {
      return jsonOut_(addTask_(data.task || data));
    }
    if (action === 'listTasks') {
      return jsonOut_({ ok: true, tasks: listTasks_() });
    }
    if (action === 'deleteTask') {
      return jsonOut_(deleteTask_(data));
    }
    if (action === 'addLog') {
      return jsonOut_(addLog_(data.log || data));
    }
    if (action === 'listLogs') {
      return jsonOut_({ ok: true, logs: listLogs_() });
    }
    if (action === 'deleteLog') {
      return jsonOut_(deleteLog_(data));
    }
    if (action === 'addLostFound') {
      return jsonOut_(addLostFound_(data.item || data));
    }
    if (action === 'updateLostFound') {
      return jsonOut_(updateLostFound_(data));
    }
    if (action === 'deleteLostFound') {
      return jsonOut_(deleteLostFound_(data));
    }
    if (action === 'listLostFound') {
      return jsonOut_({ ok: true, items: listLostFound_() });
    }
    if (action === 'addMessage') {
      return jsonOut_(addMessage_(data.message || data));
    }
    if (action === 'updateMessage') {
      return jsonOut_(updateMessage_(data));
    }
    if (action === 'listMessages') {
      return jsonOut_({ ok: true, messages: listMessages_() });
    }
    return jsonOut_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}
