// --- Code.gs ---

// ─── CORS helper ───────────────────────────────────────────────
function _corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  };
}

function _json(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

// 1. doGet — รองรับทั้ง serve HTML (เดิม) และ REST API
function doGet(e) {
  const action = e.parameter.action;

  // ── REST API mode ──────────────────────────────────────────
  if (action === 'getDropdownData')   return _json(getDropdownData());
  if (action === 'getCalendarEvents') return _json(getCalendarEvents());
  if (action === 'getRequests')       return _json(getRequests());
  if (action === 'getScriptUrl')      return _json({ url: ScriptApp.getService().getUrl() });

  // ── HTML serve mode (Apps Script Web App ดั้งเดิม) ────────
  let page = e.parameter.page || 'index';
  let html = HtmlService.createTemplateFromFile(page).evaluate();
  html.setTitle('ระบบจองรถโรงพยาบาล').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

// 2. doPost — รองรับ action ที่ต้องการ write
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'saveBooking')  return _json(saveBooking(body.data));
    if (action === 'updateStatus') return _json({ result: updateStatus(body.id, body.status) });
    if (action === 'verifyAdminPassword') return _json({ ok: verifyAdminPassword(body.password) });

    return _json({ error: 'Unknown action: ' + action });
  } catch(err) {
    return _json({ error: err.message });
  }
}

// 2b. ฟังก์ชันดึง URL (ยังคงไว้สำหรับ google.script.run fallback)
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// 3. ฟังก์ชันดึงรายชื่อรถและคนขับไปแสดงใน Dropdown
function getDropdownData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const vehicles = ss.getSheetByName('Vehicles').getDataRange().getValues().slice(1).map(r => r[0]).filter(String);
  const drivers = ss.getSheetByName('Drivers').getDataRange().getValues().slice(1).map(r => r[0]).filter(String);
  return { vehicles: vehicles, drivers: drivers };
}

// --- Helper: แปลงค่าจาก Sheets cell เป็น timestamp (ms) อย่างปลอดภัย ---
// dateCell = row[3] หรือ row[5]  (Date object หรือ string)
// timeCell = row[4] หรือ row[6]  (Date object ของ epoch 1899-12-30, string "HH:mm" หรือ "H:mm")
function cellsToTimestamp_(dateCell, timeCell) {
  // 1) แปลงวันที่
  let dateStr;
  if (dateCell instanceof Date) {
    dateStr = Utilities.formatDate(dateCell, "GMT+7", "yyyy-MM-dd");
  } else {
    // Sheets อาจส่งมาเป็น "29/3/2026" หรือ "2026-03-29"
    const s = dateCell.toString().trim();
    // ถ้าเป็น dd/MM/yyyy → แปลงเป็น yyyy-MM-dd
    const dmyMatch = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (dmyMatch) {
      dateStr = dmyMatch[3] + '-' +
                dmyMatch[2].padStart(2,'0') + '-' +
                dmyMatch[1].padStart(2,'0');
    } else {
      dateStr = s; // สมมติ yyyy-MM-dd อยู่แล้ว
    }
  }

  // 2) แปลงเวลา
  let timeStr;
  if (timeCell instanceof Date) {
    // Sheets เก็บ time-only เป็น Date ที่ epoch = 1899-12-30T00:00:00
    // ดึงเฉพาะ HH:mm จาก timezone GMT+7
    timeStr = Utilities.formatDate(timeCell, "GMT+7", "HH:mm");
  } else {
    // string เช่น "9:05", "09:05", "23:58"
    const t = timeCell.toString().trim();
    const parts = t.split(':');
    timeStr = parts[0].padStart(2,'0') + ':' + (parts[1] || '00').padStart(2,'0');
  }

  // 3) รวมกันแล้วแปลง — ใช้ new Date() กับ ISO string เพื่อหลีกเลี่ยง locale
  const iso = dateStr + 'T' + timeStr + ':00+07:00';
  const ts = new Date(iso).getTime();
  if (isNaN(ts)) throw new Error('Invalid date: ' + iso);
  return ts;
}

// 4. ฟังก์ชันบันทึกข้อมูลการจอง (พร้อมตรวจสอบการจองซ้ำ)
function saveBooking(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Requests');
  const vehicle = (data.vehicle || '').trim();

  // --- แปลงช่วงเวลาของคำขอใหม่ ---
  // ฟอร์มส่ง dateOut="2026-03-29", timeOut="23:58" มาเป็น string เสมอ
  let newStart, newEnd;
  try {
    const newStartIso = data.dateOut + 'T' + data.timeOut + ':00+07:00';
    const newEndIso   = data.dateIn  + 'T' + data.timeIn  + ':00+07:00';
    newStart = new Date(newStartIso).getTime();
    newEnd   = new Date(newEndIso).getTime();
    if (isNaN(newStart) || isNaN(newEnd)) throw new Error('Invalid form date');
  } catch(e) {
    return { success: false, message: '⚠️ รูปแบบวันที่หรือเวลาไม่ถูกต้อง กรุณาตรวจสอบอีกครั้ง' };
  }

  // ตรวจว่า เวลาออก < เวลากลับ
  if (newEnd <= newStart) {
    return { success: false, message: '⚠️ วันเวลากลับต้องมากกว่าวันเวลาออกเดินทาง' };
  }

  // --- ตรวจสอบการจองซ้ำ ---
  if (vehicle) {
    const existing = sheet.getDataRange().getValues(); // getValues() เพื่อได้ Date object จริง
    existing.shift(); // ตัดหัวข้อ

    for (let i = 0; i < existing.length; i++) {
      const row       = existing[i];
      const rowStatus  = (row[11] || '').toString().trim();
      const rowVehicle = (row[9]  || '').toString().trim();

      // ข้ามรายการที่ยกเลิกแล้ว
      if (rowStatus === 'ยกเลิก') continue;

      // ตรวจเฉพาะรถคันเดียวกัน
      if (rowVehicle !== vehicle) continue;

      // แปลงวันที่+เวลาของรายการเดิมใน Sheets
      let exStart, exEnd;
      try {
        exStart = cellsToTimestamp_(row[3], row[4]);
        exEnd   = cellsToTimestamp_(row[5], row[6]);
      } catch(err) {
        Logger.log('Skip row ' + row[0] + ': ' + err.message);
        continue;
      }

      // ตรวจ overlap: ทับซ้อนเมื่อ newStart < exEnd && newEnd > exStart
      if (newStart < exEnd && newEnd > exStart) {
        // สร้างข้อความแสดงเวลาที่ชนกัน
        const fmtDate = d => Utilities.formatDate(new Date(d), "GMT+7", "d/M/yyyy HH:mm");
        return {
          success: false,
          message:
            `⚠️ รถ "${vehicle}" ถูกจองในช่วงเวลาดังกล่าวแล้ว\n` +
            `📋 REQ: ${row[0]}\n` +
            `👤 ผู้จอง: ${row[1]} (${row[2]})\n` +
            `🕐 ช่วงเวลา: ${fmtDate(exStart)} → ${fmtDate(exEnd)}\n\n` +
            `กรุณาเลือกรถคันอื่น หรือเปลี่ยนช่วงเวลา`
        };
      }
    }
  }
  // --- จบการตรวจสอบ ---

  const newId = 'REQ' + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMddHHmmss");
  sheet.appendRow([
    newId, data.name, data.department, data.dateOut, data.timeOut,
    data.dateIn, data.timeIn, data.destination, data.phone,
    data.vehicle, data.driver, 'รอดำเนินการ'
  ]);
  return { success: true, message: '✅ บันทึกข้อมูลสำเร็จ!' };
}

// 5. ฟังก์ชันดึงข้อมูลคำขอทั้งหมด พร้อม timestamp เพื่อให้ฝั่ง Admin ตรวจสอบการอัปเดต
function getRequests() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Requests');
  const data = sheet.getDataRange().getDisplayValues();
  data.shift();
  return {
    rows: data,
    serverTime: new Date().getTime()  // ส่ง timestamp กลับด้วย
  };
}

// 6. ฟังก์ชันอัปเดตสถานะการจอง — บันทึกสถานะ + timestamp ที่คอลัมน์ M
function updateStatus(id, status) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Requests');
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.getRange(i + 1, 12).setValue(status);
      // บันทึกเวลาที่อัปเดต → ใช้โดย weeklyCleanup เพื่อตัดสินใจลบ
      sheet.getRange(i + 1, 13).setValue(new Date());
      return "อัปเดตสถานะสำเร็จ";
    }
  }
  return "ไม่พบรายการ";
}

// ─────────────────────────────────────────────────────────────────
// 6b. ลบแถวที่สถานะ "ดำเนินการเรียบร้อย" / "ยกเลิก"
//     โดยอัตโนมัติเมื่อขึ้นสัปดาห์ถัดไป (ทุกวันจันทร์ 00:00 น.)
//
//  วิธีติดตั้ง Trigger (ทำครั้งเดียว):
//    รัน setupWeeklyCleanupTrigger() ใน Apps Script Editor
//    หรือตั้งด้วยมือ: Triggers > + Add Trigger >
//      Function: weeklyCleanup | Time-driven > Week timer > Every Monday
// ─────────────────────────────────────────────────────────────────

function weeklyCleanup() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Requests');
  const data  = sheet.getDataRange().getValues();

  // หาต้นสัปดาห์ปัจจุบัน (จันทร์ 00:00 GMT+7)
  const tz        = "GMT+7";
  const todayStr  = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  const today     = new Date(todayStr + "T00:00:00+07:00");
  const dow       = today.getDay(); // 0=อาทิตย์
  const diffMon   = (dow === 0) ? -6 : 1 - dow;
  const weekStart = new Date(today);
  weekStart.setDate(today.getDate() + diffMon); // จันทร์ 00:00 น. ของสัปดาห์นี้

  Logger.log('weeklyCleanup — weekStart: ' + weekStart);

  const deleteStatuses = ['ดำเนินการเรียบร้อย', 'ยกเลิก'];
  const rowsToDelete   = [];

  for (let i = 1; i < data.length; i++) {
    const status    = (data[i][11] || '').toString().trim();
    const updatedAt = data[i][12]; // คอลัมน์ M — Date หรือ empty

    if (!deleteStatuses.includes(status)) continue;

    // ใช้ timestamp อัปเดต (col M) เป็นหลัก
    // ถ้ายังไม่มี (ข้อมูลเก่าก่อนมี col M) → fallback ใช้วันออกเดินทาง (col D)
    let refDate;
    if (updatedAt instanceof Date && !isNaN(updatedAt.getTime())) {
      refDate = updatedAt;
    } else {
      const raw = data[i][3];
      refDate = (raw instanceof Date) ? raw : new Date(raw);
    }

    // ลบถ้า refDate อยู่ก่อนต้นสัปดาห์ปัจจุบัน
    if (!isNaN(refDate) && refDate < weekStart) {
      rowsToDelete.push(i + 1); // sheet row = array index + 1
    }
  }

  // ลบจากล่างขึ้นบนเพื่อไม่ให้ index เลื่อน
  rowsToDelete.reverse().forEach(rowNum => {
    sheet.deleteRow(rowNum);
    Logger.log('Deleted row: ' + rowNum);
  });

  Logger.log('weeklyCleanup done — deleted ' + rowsToDelete.length + ' rows');
}

// สร้าง Trigger ทุกวันจันทร์ 00:00–01:00 น. (รันครั้งเดียว)
function setupWeeklyCleanupTrigger() {
  // ลบ trigger เก่าชื่อเดิมก่อน ป้องกัน duplicate
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'weeklyCleanup')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('weeklyCleanup')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(0)
    .create();

  Logger.log('weeklyCleanup trigger created successfully');
}

// 7. ฟังก์ชันดึงข้อมูลสำหรับ FullCalendar
function getCalendarEvents() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Requests');
  const data = sheet.getDataRange().getValues();
  data.shift();

  const events = [];

  data.forEach(row => {
    if (!row[3] || !row[5]) return;

    try {
      let startDate = new Date(row[3]);
      let endDate   = new Date(row[5]);
      let startStr  = Utilities.formatDate(startDate, "GMT+7", "yyyy-MM-dd");
      let endStr    = Utilities.formatDate(endDate,   "GMT+7", "yyyy-MM-dd");

      let startTime = (row[4] instanceof Date)
        ? Utilities.formatDate(row[4], "GMT+7", "HH:mm:ss")
        : row[4] + ":00";
      let endTime = (row[6] instanceof Date)
        ? Utilities.formatDate(row[6], "GMT+7", "HH:mm:ss")
        : row[6] + ":00";

      let status = row[11];
      let eventColor = '#6c757d';
      if (status === 'ดำเนินการเรียบร้อย') eventColor = '#198754';
      if (status === 'ยกเลิก')             eventColor = '#dc3545';

      events.push({
        title: (row[9] || 'ไม่ระบุรถ') + ' (' + row[7] + ')',
        start: startStr + 'T' + startTime,
        end:   endStr   + 'T' + endTime,
        color: eventColor,
        extendedProps: {
          booker:     row[1],
          department: row[2],
          driver:     row[10] || 'ไม่ระบุ',
          status:     status
        }
      });
    } catch(e) { /* ข้ามแถวที่ผิดพลาด */ }
  });

  return events;
}

// 8. ฟังก์ชันตรวจสอบรหัสผ่าน Admin
function verifyAdminPassword(password) {
  // เก็บรหัสผ่านใน Script Properties (ตั้งค่าใน Project Settings > Script Properties)
  // key: ADMIN_PASSWORD  value: รหัสผ่านที่ต้องการ
  const stored = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  return password === stored;
}