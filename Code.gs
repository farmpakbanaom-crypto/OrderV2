/**
 * ฟาร์มผักบ้านออม — ระบบรับออเดอร์
 * Google Apps Script backend
 *
 * วิธีติดตั้ง:
 * 1. เปิดสเปรดชีตออเดอร์ของคุณ -> Extensions > Apps Script
 * 2. ลบโค้ดเดิมในไฟล์ Code.gs แล้ววางโค้ดนี้ทั้งหมดแทน
 * 3. ให้แน่ใจว่ามีไฟล์ HTML ชื่อ "Index" ในโปรเจกต์เดียวกัน (วางเนื้อหาจาก index.html)
 * 4. กด Deploy > New deployment > เลือกประเภท "Web app"
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. กด Deploy แล้วเปิด Web app URL ที่ได้ใช้งานได้เลย
 *    (ไม่ต้องตั้งค่า URL ใดๆ เพิ่มเติมในไฟล์ HTML แล้ว
 *     เพราะหน้าเว็บเรียกกลับมาที่สคริปต์นี้โดยตรงผ่าน google.script.run)
 */

const SHEET_NAME = 'ชีต1'; // เปลี่ยนชื่อได้ถ้าชีตของคุณชื่ออื่น
const HEADERS = [
  'วันที่สั่งซื้อ',        // A
  'ชื่อผู้สั่ง',            // B
  'ต้องการผักสลัด',        // C
  'ต้องการหน่อไม้ต้ม',      // D
  'กำหนดจัดส่ง',           // E
  'สถานที่ส่ง',             // F
  'ราคารวม',               // G
  'เบอร์โทร',              // H
  'จำนวนผักสลัด(ชุด)',      // I
  'จำนวนหน่อไม้ต้ม(กก.)',   // J
  'สถานะ',                 // K
  'รหัสออเดอร์'             // L
];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const isEmpty = firstRow.every(function (v) { return v === '' || v === null; });
  if (isEmpty) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function fmtDate_(d) {
  if (!d) return '';
  if (Object.prototype.toString.call(d) === '[object Date]') {
    return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd');
  }
  return String(d);
}

function rowToOrder_(row, rowIndex) {
  return {
    row: rowIndex,
    orderDate: fmtDate_(row[0]),
    customerName: row[1] || '',
    saladDetail: row[2] || '',
    bambooDetail: row[3] || '',
    deliveryDate: fmtDate_(row[4]),
    location: row[5] || '',
    totalPrice: Number(row[6]) || 0,
    phone: row[7] || '',
    saladQty: Number(row[8]) || 0,
    bambooQty: Number(row[9]) || 0,
    status: row[10] || 'รอจัดส่ง',
    orderId: row[11] || ''
  };
}

function getAllOrders_() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const orders = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (!row[1] && !row[2] && !row[3]) continue; // แถวว่าง
    orders.push(rowToOrder_(row, i + 2));
  }
  return orders;
}

function isoWeekKey_(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = (d.getUTCDay() + 6) % 7;
  d.setUTCDate(d.getUTCDate() - dayNum + 3);
  const firstThursday = new Date(Date.UTC(d.getUTCFullYear(), 0, 4));
  const weekNum = 1 + Math.round(((d - firstThursday) / 86400000 - 3 + ((firstThursday.getUTCDay() + 6) % 7)) / 7);
  return d.getUTCFullYear() + '-W' + String(weekNum).padStart(2, '0');
}

function buildSummary_(orders) {
  const now = new Date();
  const todayStr = fmtDate_(now);
  const thisWeekKey = isoWeekKey_(now);
  const thisMonthKey = Utilities.formatDate(now, Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM');

  let week = { total: 0, count: 0 };
  let month = { total: 0, count: 0 };
  let today = { total: 0, count: 0 };

  orders.forEach(function (o) {
    if (!o.orderDate) return;
    const d = new Date(o.orderDate + 'T00:00:00');
    if (isNaN(d)) return;
    const wk = isoWeekKey_(d);
    const mo = Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM');
    if (wk === thisWeekKey) { week.total += o.totalPrice; week.count += 1; }
    if (mo === thisMonthKey) { month.total += o.totalPrice; month.count += 1; }
    if (o.orderDate === todayStr) { today.total += o.totalPrice; today.count += 1; }
  });

  return { today: today, week: week, month: month };
}

/* =========================================================
 * ฟังก์ชันที่หน้าเว็บ (Index.html) เรียกใช้โดยตรงผ่าน
 * google.script.run — ไม่ต้องพึ่ง fetch() หรือ URL ใดๆ
 * เพราะ HtmlService serve หน้าเว็บผ่าน iframe คนละโดเมน
 * การ fetch กลับไปที่ window.location.href จึงใช้งานไม่ได้
 * ========================================================= */

function listOrdersForClient() {
  return getAllOrders_();
}

function getSummaryForClient() {
  return buildSummary_(getAllOrders_());
}

function addOrderForClient(d) {
  d = d || {};
  const sheet = getSheet_();
  const orderId = 'ORD' + new Date().getTime();
  sheet.appendRow([
    fmtDate_(new Date()), // วันที่สั่งซื้อ = วันที่บันทึกจริงเสมอ
    d.customerName || '',
    d.saladDetail || '',
    d.bambooDetail || '',
    d.deliveryDate || '',
    d.location || '',
    Number(d.totalPrice) || 0,
    d.phone || '',
    Number(d.saladQty) || 0,
    Number(d.bambooQty) || 0,
    'รอจัดส่ง',
    orderId
  ]);
  return { ok: true, orderId: orderId };
}

function updateStatusForClient(orderId, status) {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  const ids = sheet.getRange(2, 12, Math.max(lastRow - 1, 0), 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === orderId) {
      sheet.getRange(i + 2, 11).setValue(status);
      return { ok: true };
    }
  }
  return { ok: false, error: 'ไม่พบออเดอร์นี้' };
}

/* ---------------------------------------------------------
 * doGet / doPost ยังคงไว้เผื่อเรียกผ่าน URL ภายนอกโดยตรง
 * (ไม่จำเป็นสำหรับหน้าเว็บหลัก ซึ่งใช้ google.script.run แล้ว)
 * --------------------------------------------------------- */

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  try {
    if (action === 'summary') {
      return jsonOut_({ ok: true, summary: getSummaryForClient() });
    }
    if (action === 'list') {
      return jsonOut_({ ok: true, orders: getAllOrders_() });
    }
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('ฟาร์มผักบ้านออม')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover, maximum-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'addOrder') {
      return jsonOut_(addOrderForClient(body.data || {}));
    }

    if (action === 'updateStatus') {
      return jsonOut_(updateStatusForClient(body.orderId, body.status));
    }

    return jsonOut_({ ok: false, error: 'ไม่รู้จักคำสั่ง: ' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}
