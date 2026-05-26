// ============================================================
//  MPS Stock — Google Apps Script Backend  v1.5.1
//  วิธีติดตั้ง:
//  1. เปิด Google Sheets → Extensions → Apps Script
//  2. วางโค้ดนี้ทั้งหมดแทนที่โค้ดเดิม → Save (Ctrl+S)
//  3. Deploy → New deployment → Web app
//     - Execute as: Me
//     - Who has access: Anyone
//  4. กด Deploy → Copy URL → วางใน ⚙️ Apps Script URL
// ============================================================

const BORROW_SHEET = 'BorrowReturn';  // ชื่อ sheet ยืม-คืน
const USERS_SHEET  = 'Users';          // ชื่อ sheet users
// ⚠️ ใส่ Spreadsheet ID ของคุณตรงนี้
const SPREADSHEET_ID = '1d32oqw7pGZMlFgO7hISoNmJdKfHKifkEsa8Paxwds84';

// คอลัมน์ใน sheet BorrowReturn (เรียงตามลำดับ)
const BORROW_COLS = ['id','borrower','dept','itemId','itemName','qty',
                     'borrowDate','dueDate','returnDate','status','purpose','returnNote'];

// ── เปิด Spreadsheet (ใช้ ID ที่กำหนดไว้ หรือจาก payload)
function getSpreadsheet(data) {
  const id = (data && data.sheetId) ? data.sheetId : SPREADSHEET_ID;
  return SpreadsheetApp.openById(id);
}

// ── รับ POST request จาก Dashboard
function doPost(e) {
  const res = ContentService.createTextOutput();
  res.setMimeType(ContentService.MimeType.JSON);

  try {
    const data = JSON.parse(e.postData.contents);
    let result;

    if      (data.action === 'addBorrow')    result = addBorrow(data);
    else if (data.action === 'updateReturn') result = updateReturn(data);
    else                                     result = { ok: false, msg: 'Unknown action' };

    res.setContent(JSON.stringify(result));
  } catch(err) {
    res.setContent(JSON.stringify({ ok: false, msg: err.toString() }));
  }
  return res;
}

// ── รับ GET request (ทดสอบว่า deploy ถูกต้อง)
function doGet() {
  return ContentService.createTextOutput(
    JSON.stringify({ ok: true, msg: 'MPS Stock GAS v1.3.0 — ready' })
  ).setMimeType(ContentService.MimeType.JSON);
}



// ── เพิ่ม/อัปเดต stock item จาก Internal Test
function updateStockItem(data) {
  try {
    const ss    = getSpreadsheet(data);
    const sheetName = data.stockSheet || 'Stock';
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { ok: false, msg: `Sheet "${sheetName}" not found` };

    const rows = sheet.getDataRange().getValues();
    const headers = rows[0].map(h => String(h).toLowerCase().trim());
    const idIdx  = headers.findIndex(h => h === 'id');
    const qtyIdx = headers.findIndex(h => h === 'qty');
    const nameIdx = headers.findIndex(h => h === 'name');
    if (idIdx < 0 || qtyIdx < 0) return { ok: false, msg: 'Missing id/qty columns' };

    const delta = parseInt(data.delta || 1);

    // Find existing row
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][idIdx]).trim() === String(data.id).trim()) {
        const cur = parseInt(rows[i][qtyIdx] || 0);
        sheet.getRange(i + 1, qtyIdx + 1).setValue(Math.max(0, cur + delta));
        return { ok: true, msg: `Updated ${data.id} qty +${delta}` };
      }
    }

    // Not found — append new row
    const STOCK_COLS = ['id','name','category','partNo','brand','qty','minQty','unit','location','status'];
    if (sheet.getLastRow() <= 1) {
      // ensure header
      const headerRow = rows[0].join(',').toLowerCase();
      if (!headerRow.includes('id')) {
        sheet.getRange(1, 1, 1, STOCK_COLS.length).setValues([STOCK_COLS]);
      }
    }
    const newRow = STOCK_COLS.map(col => {
      if (col === 'id')       return data.id || '';
      if (col === 'name')     return data.name || '';
      if (col === 'category') return data.category || 'flow';
      if (col === 'brand')    return data.brand || 'Emerson';
      if (col === 'qty')      return String(delta);
      if (col === 'minQty')   return '1';
      if (col === 'unit')     return data.unit || 'ชิ้น';
      if (col === 'status')   return 'ok';
      return '';
    });
    sheet.appendRow(newRow);
    return { ok: true, msg: `Added new stock item ${data.id}` };
  } catch(e) {
    return { ok: false, msg: e.toString() };
  }
}

// ── ตรวจสอบ login จาก sheet Users
function verifyLogin(data) {
  try {
    const ss    = getSpreadsheet(data);
    const sheet = ss.getSheetByName(USERS_SHEET);
    if (!sheet) return { ok: false, msg: 'Users sheet not found' };

    const rows    = sheet.getDataRange().getValues();
    if (rows.length < 2) return { ok: false, msg: 'No users found' };

    const headers = rows[0].map(h => String(h).toLowerCase().trim());
    const uIdx = headers.indexOf('username');
    const pIdx = headers.indexOf('password');
    const rIdx = headers.indexOf('role');
    const nIdx = headers.indexOf('name');
    const aIdx = headers.indexOf('active');

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const active = aIdx >= 0 ? String(row[aIdx]).toLowerCase() : 'true';
      if (active === 'false') continue;
      if (String(row[uIdx]).toLowerCase() === String(data.username).toLowerCase() &&
          String(row[pIdx]) === String(data.password)) {
        return {
          ok: true,
          username: row[uIdx],
          name:     nIdx >= 0 ? row[nIdx] : row[uIdx],
          role:     rIdx >= 0 ? String(row[rIdx]).toLowerCase() : 'viewer'
        };
      }
    }
    return { ok: false, msg: 'Invalid credentials' };
  } catch(e) {
    return { ok: false, msg: e.toString() };
  }
}

// ── เพิ่มแถวยืมใหม่
function addBorrow(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(BORROW_SHEET);
  if (!sheet) return { ok: false, msg: `Sheet "${BORROW_SHEET}" not found` };

  // สร้าง header row ถ้ายังไม่มี
  ensureHeader(sheet);

  const row = BORROW_COLS.map(col => data[col] ?? '');
  sheet.appendRow(row);

  // อัปเดตสต็อกใน Stock sheet ด้วย
  updateStockQty(data.itemId, -parseInt(data.qty || 0));

  return { ok: true, msg: 'addBorrow success', id: data.id };
}

// ── อัปเดตสถานะคืน
function updateReturn(data) {
  const ss    = getSpreadsheet(data);
  const sheet = ss.getSheetByName(BORROW_SHEET);
  if (!sheet) return { ok: false, msg: `Sheet "${BORROW_SHEET}" not found` };

  const rows = sheet.getDataRange().getValues();
  const idCol = BORROW_COLS.indexOf('id');         // col index of 'id'
  const statusCol = BORROW_COLS.indexOf('status');
  const returnDateCol = BORROW_COLS.indexOf('returnDate');
  const noteCol = BORROW_COLS.indexOf('returnNote');
  const qtyCol  = BORROW_COLS.indexOf('qty');
  const itemCol = BORROW_COLS.indexOf('itemId');

  let found = false;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]) === String(data.id)) {
      const rowNum = i + 1; // 1-indexed
      sheet.getRange(rowNum, statusCol + 1).setValue('returned');
      sheet.getRange(rowNum, returnDateCol + 1).setValue(data.returnDate);
      if (noteCol >= 0) sheet.getRange(rowNum, noteCol + 1).setValue(data.note || '');

      // คืนสต็อก
      const qty    = parseInt(rows[i][qtyCol] || 0);
      const itemId = rows[i][itemCol];
      updateStockQty(itemId, qty);

      found = true;
      break;
    }
  }

  if (!found) return { ok: false, msg: `BorrowID ${data.id} not found` };
  return { ok: true, msg: 'updateReturn success' };
}

// ── อัปเดตจำนวนคงเหลือใน Stock sheet
function updateStockQty(itemId, delta, data) {
  try {
    const ss    = getSpreadsheet(data);
    const sheet = ss.getSheetByName('Stock');
    if (!sheet) return;

    const rows = sheet.getDataRange().getValues();
    // หา column 'id' จาก header row
    const headers = rows[0].map(h => String(h).toLowerCase().trim());
    const idIdx  = headers.findIndex(h => h === 'id');
    const qtyIdx = headers.findIndex(h => h === 'qty');
    if (idIdx < 0 || qtyIdx < 0) return;

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][idIdx]).trim() === String(itemId).trim()) {
        const cur = parseInt(rows[i][qtyIdx] || 0);
        const newQty = Math.max(0, cur + delta);
        sheet.getRange(i + 1, qtyIdx + 1).setValue(newQty);
        break;
      }
    }
  } catch(e) {
    console.log('updateStockQty error:', e);
  }
}

// ── สร้าง header row ถ้า sheet ใหม่หรือว่างอยู่
function ensureHeader(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(BORROW_COLS);
  } else {
    const firstCell = sheet.getRange(1, 1).getValue();
    if (String(firstCell).toLowerCase() !== 'id') {
      sheet.insertRowBefore(1);
      sheet.getRange(1, 1, 1, BORROW_COLS.length).setValues([BORROW_COLS]);
    }
  }
}
