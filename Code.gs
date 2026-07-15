/**
 * ==================================================================
 * ระบบจัดการรถ (Fleet Management) — Backend เต็มรูปแบบ (Google Apps Script)
 * ==================================================================
 * ก๊อปปี้ไฟล์นี้ทั้งหมดวางแทนที่ไฟล์ Code.gs ในโปรเจกต์ Apps Script ของคุณได้เลย
 * รองรับข้อมูลครบทั้ง 6 ประเภท: Trips, Refuels, Vehicles, Jobs, Maintenance, Tires
 * ตรงกับ action ที่หน้าเว็บ (GAS_WEB_APP_URL) เรียกใช้อยู่แล้วทุกจุด ไม่ต้องแก้โค้ดฝั่งเว็บ
 *
 * ==================================================================
 * วิธีติดตั้ง (ทำครั้งเดียว)
 * ==================================================================
 * 1) เปิด Google Sheet ที่จะใช้เป็นฐานข้อมูล (สร้างชีตเปล่าใหม่ก็ได้ หรือใช้ของเดิม)
 * 2) เมนู Extensions > Apps Script
 * 3) ลบโค้ดเดิมทั้งหมดในไฟล์ Code.gs แล้ววางโค้ดทั้งไฟล์นี้แทน
 * 4) ถ้าสคริปต์นี้ "ผูกกับชีต" อยู่แล้ว (เปิดผ่าน Extensions > Apps Script จากในชีต) ข้ามข้อนี้ได้เลย
 *    ถ้าเป็นสคริปต์แบบ standalone ให้ใส่ Spreadsheet ID ที่ตัวแปร SPREADSHEET_ID ด้านล่าง
 *    (ID คือส่วนที่อยู่ระหว่าง /d/ กับ /edit ใน URL ของ Google Sheet)
 * 5) กดรันฟังก์ชัน setupSheets() หนึ่งครั้ง (เลือกฟังก์ชันจาก dropdown ด้านบน แล้วกด Run)
 *    - ครั้งแรกจะขอ authorize สิทธิ์การเข้าถึง Sheet ให้กด Allow
 *    - ฟังก์ชันนี้จะสร้างชีตทั้ง 6 แผ่นพร้อมหัวคอลัมน์ให้อัตโนมัติ (รันซ้ำได้ ไม่ทับข้อมูลเดิม)
 * 6) Deploy > New deployment (หรือถ้ามี deployment เดิมอยู่แล้วให้ไปที่ Manage deployments > Edit
 *    แล้ว deploy ทับของเดิม เพื่อให้ URL เดิมยังใช้ได้ ไม่ต้องแก้ GAS_WEB_APP_URL ในหน้าเว็บ)
 *    - Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 7) คัดลอก Web app URL ไปวางที่ตัวแปร GAS_WEB_APP_URL ในไฟล์ HTML (ถ้ายังไม่เคยตั้งค่า)
 * ==================================================================
 */

// ถ้าสคริปต์นี้ "ผูกกับชีต" อยู่แล้วปล่อยว่างไว้ได้เลย — ถ้าเป็น standalone script ให้ใส่ Spreadsheet ID
const SPREADSHEET_ID = '';

// หัวคอลัมน์ของแต่ละชีต (แถวที่ 1) — ต้องตรงกับชื่อ field ที่หน้าเว็บส่งมาทุกตัวอักษร
const SHEET_HEADERS = {
  Trips: ['id', 'category', 'date', 'plate', 'origin', 'dest', 'dist', 'fuelRate', 'currentFuelPrice', 'fuelCost', 'driverFee', 'revenue', 'weight', 'profit'],
  Refuels: ['id', 'category', 'date', 'plate', 'liters', 'price', 'total'],
  Vehicles: ['id', 'category', 'plate', 'driver', 'type', 'model', 'year', 'baselineOdometer', 'baselineDate'],
  Jobs: ['id', 'category', 'origin', 'dest', 'dist', 'revenue', 'fuelRate', 'currentFuelPrice', 'fuelCost', 'driverFee'],
  Maintenance: ['id', 'plate', 'date', 'category', 'odometer', 'description', 'cost', 'partsChanged', 'technician'],
  Tires: ['id', 'plate', 'position', 'brand', 'serial', 'installDate', 'installOdometer', 'status', 'removeDate', 'removeOdometer', 'notes']
};

// action ที่หน้าเว็บเรียกมา -> ชีตปลายทาง + ชนิดการทำงาน (เพิ่ม/แก้ไข/ลบ)
const ACTION_MAP = {
  addTrip: { sheet: 'Trips', op: 'add' }, editTrip: { sheet: 'Trips', op: 'edit' }, deleteTrip: { sheet: 'Trips', op: 'delete' },
  addRefuel: { sheet: 'Refuels', op: 'add' }, editRefuel: { sheet: 'Refuels', op: 'edit' }, deleteRefuel: { sheet: 'Refuels', op: 'delete' },
  addVehicle: { sheet: 'Vehicles', op: 'add' }, editVehicle: { sheet: 'Vehicles', op: 'edit' }, deleteVehicle: { sheet: 'Vehicles', op: 'delete' },
  addJob: { sheet: 'Jobs', op: 'add' }, editJob: { sheet: 'Jobs', op: 'edit' }, deleteJob: { sheet: 'Jobs', op: 'delete' },
  addMaintenance: { sheet: 'Maintenance', op: 'add' }, editMaintenance: { sheet: 'Maintenance', op: 'edit' }, deleteMaintenance: { sheet: 'Maintenance', op: 'delete' },
  addTire: { sheet: 'Tires', op: 'add' }, editTire: { sheet: 'Tires', op: 'edit' }, deleteTire: { sheet: 'Tires', op: 'delete' }
};

// ==================================================================
// Entry points หลักที่ Apps Script Web App เรียกอัตโนมัติ
// ==================================================================

function doGet(e) {
  const ss = getSpreadsheet_();
  const result = { status: 'success' };
  Object.keys(SHEET_HEADERS).forEach(name => {
    const key = name === 'Maintenance' ? 'maintenance'
              : name === 'Tires' ? 'tires'
              : name.toLowerCase(); // Trips->trips, Refuels->refuels, Vehicles->vehicles, Jobs->jobs
    result[key] = sheetToObjects_(getSheet_(ss, name));
  });
  return jsonResponse_(result);
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const mapping = ACTION_MAP[action];
    if (!mapping) {
      return jsonResponse_({ status: 'error', message: 'ไม่รู้จัก action: ' + action });
    }

    const ss = getSpreadsheet_();
    const sheet = getSheet_(ss, mapping.sheet);

    if (mapping.op === 'add') {
      appendObjectRow_(sheet, payload.data || {});
    } else if (mapping.op === 'edit') {
      updateObjectRowById_(sheet, payload.data || {});
    } else if (mapping.op === 'delete') {
      // รองรับทั้ง 2 รูปแบบ payload ที่หน้าเว็บอาจส่งมา: {id, sheetType} หรือ {data:{id}}
      const targetId = payload.id !== undefined ? payload.id : (payload.data && payload.data.id);
      deleteObjectRowById_(sheet, targetId);
    }

    return jsonResponse_({ status: 'success' });
  } catch (err) {
    return jsonResponse_({ status: 'error', message: String(err) });
  }
}

// รันฟังก์ชันนี้ "หนึ่งครั้ง" จากเมนู Apps Script (Run) เพื่อสร้างชีตทั้งหมดพร้อมหัวคอลัมน์
// รันซ้ำได้อย่างปลอดภัย — ถ้าชีตไหนมีอยู่แล้วจะไม่ถูกสร้างซ้ำหรือทับข้อมูล
function setupSheets() {
  const ss = getSpreadsheet_();
  const created = [];
  Object.keys(SHEET_HEADERS).forEach(name => {
    const existed = !!ss.getSheetByName(name);
    getSheet_(ss, name);
    if (!existed) created.push(name);
  });
  Logger.log(created.length > 0
    ? 'สร้างชีตใหม่: ' + created.join(', ')
    : 'มีชีตครบทั้งหมดอยู่แล้ว ไม่ต้องสร้างเพิ่ม');
}

// ==================================================================
// ตัวช่วยภายใน
// ==================================================================

function getSpreadsheet_() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(SHEET_HEADERS[name]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, SHEET_HEADERS[name].length).setFontWeight('bold');
  }
  return sheet;
}

function sheetToObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1)
    .filter(r => r.some(cell => cell !== '' && cell !== null))
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = r[i]; });
      return obj;
    });
}

// เขียนทั้งแถวโดยบังคับ format เป็น "ข้อความล้วน" (@) ก่อนใส่ค่า
// ป้องกัน Google Sheets ตีความวันที่แบบไทย (เช่น "14/7/2569") เป็นวันที่จริงแล้วแปลงเพี้ยน
function writeRowAsText_(sheet, rowIndex, values) {
  const range = sheet.getRange(rowIndex, 1, 1, values.length);
  range.setNumberFormat('@');
  range.setValues([values]);
}

function appendObjectRow_(sheet, dataObj) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => (dataObj[h] !== undefined ? dataObj[h] : ''));
  writeRowAsText_(sheet, sheet.getLastRow() + 1, row);
}

function updateObjectRowById_(sheet, dataObj) {
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('id');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(dataObj.id)) {
      const row = headers.map(h => (dataObj[h] !== undefined ? dataObj[h] : ''));
      writeRowAsText_(sheet, i + 1, row);
      return true;
    }
  }
  // ไม่เจอ id เดิม (เช่นแก้ไขรายการที่ยังไม่เคย sync ขึ้นชีต) → เพิ่มเป็นแถวใหม่แทน กันข้อมูลหาย
  appendObjectRow_(sheet, dataObj);
  return false;
}

function deleteObjectRowById_(sheet, id) {
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('id');
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][idCol]) === String(id)) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
