// ============================================================
//  MPS Stock — Google Apps Script Backend  v1.7.0
//  วิธีติดตั้ง:
//  1. เปิด Google Sheets → Extensions → Apps Script
//  2. วางโค้ดนี้ทั้งหมดแทนที่โค้ดเดิม → Save (Ctrl+S)
//  3. Deploy → Manage deployments → ✏️ → New version → Deploy
// ============================================================

const BORROW_SHEET = 'BorrowReturn';
const USERS_SHEET  = 'Users';
const SPREADSHEET_ID = '1c5tiZzbtOpp_5PuRKpw4PbKcmHei_LWg9YZP3XhxcaE'; // ← Spreadsheet ID ของคุณ

const BORROW_COLS = ['id','borrower','dept','itemId','itemName','qty',
                     'borrowDate','dueDate','returnDate','status','purpose','returnNote'];

// ── เปิด Spreadsheet ด้วย ID (รองรับทั้ง hardcode และจาก payload)
function getSpreadsheet(data) {
  const id = (data && data.sheetId) ? data.sheetId : SPREADSHEET_ID;
  return SpreadsheetApp.openById(id);
}

// ── รับ GET request (Dashboard ส่งแบบ GET + URL params เพราะ no-cors)
function doGet(e) {
  const res = ContentService.createTextOutput();
  res.setMimeType(ContentService.MimeType.JSON);
  try {
    const p = e.parameter || {};
    // Health check
    if (!p.action) {
      res.setContent(JSON.stringify({ ok: true, msg: 'MPS Stock GAS v1.7.0 — ready' }));
      return res;
    }
    // Map params to data object
    const data = {};
    Object.keys(p).forEach(k => { data[k] = p[k]; });

    let result;
    if      (data.action === 'addBorrow')    result = addBorrow(data);
    else if (data.action === 'updateReturn') result = updateReturn(data);
    else if (data.action === 'login')        result = verifyLogin(data);
    else if (data.action === 'updateStock')  result = updateStockItem(data);
    else if (data.action === 'updateStockItem') result = updateStockRow(data);
    else if (data.action === 'addTestHistory')  result = addTestHistory(data);
    else if (data.action === 'addStockItem')    result = addStockItem(data);
    else if (data.action === 'updateStockRow')  result = updateStockRow(data);
    else if (data.action === 'deleteStockItem') result = deleteStockItem(data);
    else if (data.action === 'addInstrument')   result = addInstrument(data);
    else if (data.action === 'updateInstrument')result = updateInstrument(data);
    else if (data.action === 'deleteInstrument')result = deleteInstrument(data);
    else                                     result = { ok: false, msg: 'Unknown action: ' + data.action };

    res.setContent(JSON.stringify(result));
  } catch(err) {
    res.setContent(JSON.stringify({ ok: false, msg: err.toString() }));
  }
  return res;
}

// ── doPost รองรับไว้เผื่อ (redirect ไป doGet)
function doPost(e) {
  return doGet(e);
}

// ──────────── BORROW ────────────
function addBorrow(data) {
  try {
    const ss    = getSpreadsheet(data);
    const sheet = ss.getSheetByName(data.borrowSheet || BORROW_SHEET);
    if (!sheet) return { ok: false, msg: `Sheet "${BORROW_SHEET}" not found` };
    ensureHeader(sheet, BORROW_COLS);
    const row = BORROW_COLS.map(col => data[col] || '');
    sheet.appendRow(row);
    updateStockQty(data.itemId, -parseInt(data.qty || 0), data);
    return { ok: true, msg: 'addBorrow success', id: data.id };
  } catch(e) {
    return { ok: false, msg: e.toString() };
  }
}

// ──────────── RETURN ────────────
function updateReturn(data) {
  try {
    const ss    = getSpreadsheet(data);
    const sheet = ss.getSheetByName(data.borrowSheet || BORROW_SHEET);
    if (!sheet) return { ok: false, msg: `Sheet "${BORROW_SHEET}" not found` };

    const rows = sheet.getDataRange().getValues();
    const idCol         = BORROW_COLS.indexOf('id');
    const statusCol     = BORROW_COLS.indexOf('status');
    const returnDateCol = BORROW_COLS.indexOf('returnDate');
    const noteCol       = BORROW_COLS.indexOf('returnNote');
    const qtyCol        = BORROW_COLS.indexOf('qty');
    const itemCol       = BORROW_COLS.indexOf('itemId');

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][idCol]) === String(data.id)) {
        const rowNum = i + 1;
        sheet.getRange(rowNum, statusCol + 1).setValue('returned');
        sheet.getRange(rowNum, returnDateCol + 1).setValue(data.returnDate);
        if (noteCol >= 0) sheet.getRange(rowNum, noteCol + 1).setValue(data.note || '');
        updateStockQty(rows[i][itemCol], parseInt(rows[i][qtyCol] || 0), data);
        return { ok: true, msg: 'updateReturn success' };
      }
    }
    return { ok: false, msg: `BorrowID ${data.id} not found` };
  } catch(e) {
    return { ok: false, msg: e.toString() };
  }
}

// ──────────── STOCK QTY UPDATE ────────────
function updateStockQty(itemId, delta, data) {
  try {
    if (!itemId || delta === 0) return;
    const ss    = getSpreadsheet(data);
    const sheet = ss.getSheetByName('Stock');
    if (!sheet) return;
    const rows    = sheet.getDataRange().getValues();
    const headers = rows[0].map(h => String(h).toLowerCase().trim());
    const idIdx   = headers.findIndex(h => h === 'id');
    const qtyIdx  = headers.findIndex(h => h === 'qty');
    if (idIdx < 0 || qtyIdx < 0) return;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][idIdx]).trim() === String(itemId).trim()) {
        const newQty = Math.max(0, parseInt(rows[i][qtyIdx] || 0) + delta);
        sheet.getRange(i + 1, qtyIdx + 1).setValue(newQty);
        return;
      }
    }
  } catch(e) { console.log('updateStockQty error:', e.toString()); }
}

// ──────────── UPDATE STOCK ITEM (from Internal Test) ────────────
function updateStockItem(data) {
  try {
    const ss        = getSpreadsheet(data);
    const sheetName = data.stockSheet || 'Stock';
    const sheet     = ss.getSheetByName(sheetName);
    if (!sheet) return { ok: false, msg: `Sheet "${sheetName}" not found` };

    const rows    = sheet.getDataRange().getValues();
    const headers = rows[0].map(h => String(h).toLowerCase().trim());
    const idIdx   = headers.findIndex(h => h === 'id');
    const qtyIdx  = headers.findIndex(h => h === 'qty');
    if (idIdx < 0 || qtyIdx < 0) return { ok: false, msg: 'Missing id/qty columns in Stock sheet' };

    const delta = parseInt(data.delta || 1);

    // Find and update existing row
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][idIdx]).trim() === String(data.id).trim()) {
        const newQty = Math.max(0, parseInt(rows[i][qtyIdx] || 0) + delta);
        sheet.getRange(i + 1, qtyIdx + 1).setValue(newQty);
        return { ok: true, msg: `Updated ${data.id} qty to ${newQty}` };
      }
    }

    // Not found → append new row
    const STOCK_COLS = ['id','name','category','partNo','brand','qty','minQty','unit','location','status'];
    ensureHeader(sheet, STOCK_COLS);
    const newRow = STOCK_COLS.map(col => {
      const map = { id:data.id||'', name:data.name||'', category:data.category||'flow',
                    partNo:'—', brand:data.brand||'Emerson', qty:String(delta),
                    minQty:'1', unit:data.unit||'ชิ้น', location:'—', status:'ok' };
      return map[col] !== undefined ? map[col] : '';
    });
    sheet.appendRow(newRow);
    return { ok: true, msg: `Added new item ${data.id}` };
  } catch(e) {
    return { ok: false, msg: e.toString() };
  }
}

// ──────────── LOGIN ────────────
function verifyLogin(data) {
  try {
    const ss    = getSpreadsheet(data);
    const sheet = ss.getSheetByName(USERS_SHEET);
    if (!sheet) return { ok: false, msg: 'Users sheet not found — using local fallback' };

    const rows    = sheet.getDataRange().getValues();
    if (rows.length < 2) return { ok: false, msg: 'No users in sheet' };
    const headers = rows[0].map(h => String(h).toLowerCase().trim());
    const uIdx = headers.indexOf('username');
    const pIdx = headers.indexOf('password');
    const rIdx = headers.indexOf('role');
    const nIdx = headers.indexOf('name');
    const aIdx = headers.indexOf('active');

    for (let i = 1; i < rows.length; i++) {
      const row    = rows[i];
      const active = aIdx >= 0 ? String(row[aIdx]).toLowerCase() : 'true';
      if (active === 'false') continue;
      if (String(row[uIdx]).toLowerCase() === String(data.username).toLowerCase() &&
          String(row[pIdx]) === String(data.password)) {
        return { ok: true, username: row[uIdx],
                 name: nIdx >= 0 ? row[nIdx] : row[uIdx],
                 role: rIdx >= 0 ? String(row[rIdx]).toLowerCase() : 'viewer' };
      }
    }
    return { ok: false, msg: 'Invalid credentials' };
  } catch(e) {
    return { ok: false, msg: e.toString() };
  }
}

// ──────────── HELPER ────────────
function ensureHeader(sheet, cols) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(cols);
    return;
  }
  const firstCell = String(sheet.getRange(1,1).getValue()).toLowerCase().trim();
  if (firstCell !== cols[0].toLowerCase()) {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, cols.length).setValues([cols]);
  }
}

// ──────────── STOCK CRUD ────────────
function addStockItem(data) {
  try {
    const ss    = getSpreadsheet(data);
    const sheet = ss.getSheetByName(data.stockSheet || 'Stock');
    if (!sheet) return { ok:false, msg:'Stock sheet not found' };
    const COLS = ['id','name','category','partNo','brand','qty','minQty','unit','location','status'];
    ensureHeader(sheet, COLS);
    // Check duplicate
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0].map(h => String(h).toLowerCase().trim());
    const idIdx = headers.indexOf('id');
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][idIdx]).trim() === String(data.id).trim())
        return { ok:false, msg:'Duplicate ID: ' + data.id };
    }
    sheet.appendRow(COLS.map(c => data[c] || ''));
    return { ok:true, msg:'addStockItem success' };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}

function updateStockRow(data) {
  // Full row update (not just qty delta)
  try {
    const ss    = getSpreadsheet(data);
    const sheet = ss.getSheetByName(data.stockSheet || 'Stock');
    if (!sheet) return { ok:false, msg:'Stock sheet not found' };
    const COLS = ['id','name','category','partNo','brand','qty','minQty','unit','location','status'];
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0].map(h => String(h).toLowerCase().trim());
    const idIdx = headers.indexOf('id');
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][idIdx]).trim() === String(data.id).trim()) {
        const newRow = COLS.map(c => {
          const hi = headers.indexOf(c.toLowerCase());
          return data[c] !== undefined ? data[c] : (hi >= 0 ? rows[i][hi] : '');
        });
        sheet.getRange(i+1, 1, 1, newRow.length).setValues([newRow]);
        return { ok:true, msg:'updateStockRow success' };
      }
    }
    return { ok:false, msg:'Item not found: ' + data.id };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}

function deleteStockItem(data) {
  try {
    const ss    = getSpreadsheet(data);
    const sheet = ss.getSheetByName(data.stockSheet || 'Stock');
    if (!sheet) return { ok:false, msg:'Stock sheet not found' };
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0].map(h => String(h).toLowerCase().trim());
    const idIdx = headers.indexOf('id');
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][idIdx]).trim() === String(data.id).trim()) {
        sheet.deleteRow(i + 1);
        return { ok:true, msg:'deleteStockItem success' };
      }
    }
    return { ok:false, msg:'Item not found: ' + data.id };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}

// ──────────── INSTRUMENTS CRUD ────────────
const INST_COLS = ['id','name','owner','mfr','model','sn','range','tol','freq','type','expiry','remark'];

// Map จาก field key ที่ dashboard ส่งมา → ชื่อ header ที่เป็นไปได้ใน Google Sheets
const INST_HEADER_ALIASES = {
  'id':     ['id','รหัส','code'],
  'name':   ['name','ชื่อเครื่องมือ','ชื่อ','instrument'],
  'owner':  ['owner','ผู้เก็บรักษา','keeper','ผู้รับผิดชอบ'],
  'mfr':    ['mfr','manufacturer','brand','ยี่ห้อ'],
  'model':  ['model','รุ่น'],
  'sn':     ['sn','s/n','serial','serialnumber'],
  'range':  ['range','ช่วง'],
  'tol':    ['tol','tolerance','ค่าความคลาดเคลื่อน','error'],
  'freq':   ['freq','frequency','interval','ความถี่'],
  'type':   ['type','calibration','ลักษณะ','method'],
  'expiry': ['expiry','expire','วันหมดอายุ','หมดอายุ','expirydate'],
  'remark': ['remark','หมายเหตุ','note','remark/note'],
};

// หา column index ใน sheet จาก aliases
function findInstCol(headers, field) {
  const aliases = INST_HEADER_ALIASES[field] || [field];
  for (const alias of aliases) {
    const idx = headers.findIndex(h => h.replace(/[\s\/\-_]/g,'').toLowerCase() === alias.replace(/[\s\/\-_]/g,'').toLowerCase());
    if (idx >= 0) return idx;
  }
  return -1;
}

function addInstrument(data) {
  try {
    const ss    = getSpreadsheet(data);
    const sheet = ss.getSheetByName(data.instSheet || 'Instruments');
    if (!sheet) return { ok:false, msg:'Instruments sheet not found' };

    const rows    = sheet.getDataRange().getValues();
    const headers = rows[0].map(h => String(h).toLowerCase().trim());

    // Check duplicate by id
    const idColIdx = findInstCol(headers, 'id');
    if (idColIdx >= 0) {
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][idColIdx]).trim() === String(data.id).trim())
          return { ok:false, msg:'Duplicate ID: ' + data.id };
      }
    }

    // Build new row matching existing headers
    const newRow = rows[0].map((h, colIdx) => {
      const hNorm = String(h).toLowerCase().trim();
      // Find which field this header belongs to
      for (const [field, aliases] of Object.entries(INST_HEADER_ALIASES)) {
        const match = aliases.some(a => a.replace(/[\s\/\-_]/g,'').toLowerCase() === hNorm.replace(/[\s\/\-_]/g,'').toLowerCase());
        if (match && data[field] !== undefined) return data[field];
      }
      return '';
    });
    sheet.appendRow(newRow);
    return { ok:true, msg:'addInstrument success' };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}

function updateInstrument(data) {
  try {
    const ss    = getSpreadsheet(data);
    const sheet = ss.getSheetByName(data.instSheet || 'Instruments');
    if (!sheet) return { ok:false, msg:'Instruments sheet not found' };

    const rows    = sheet.getDataRange().getValues();
    const headers = rows[0].map(h => String(h).toLowerCase().trim());
    const idColIdx = findInstCol(headers, 'id');
    if (idColIdx < 0) return { ok:false, msg:'Cannot find id column in Instruments sheet' };

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][idColIdx]).trim() !== String(data.id).trim()) continue;

      // Update only columns that have a matching field in data
      rows[0].forEach((h, colIdx) => {
        const hNorm = String(h).toLowerCase().trim();
        for (const [field, aliases] of Object.entries(INST_HEADER_ALIASES)) {
          const match = aliases.some(a => a.replace(/[\s\/\-_]/g,'').toLowerCase() === hNorm.replace(/[\s\/\-_]/g,'').toLowerCase());
          if (match && data[field] !== undefined) {
            sheet.getRange(i+1, colIdx+1).setValue(data[field]);
            break;
          }
        }
      });
      return { ok:true, msg:'updateInstrument success' };
    }
    return { ok:false, msg:'Instrument not found: ' + data.id };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}

function deleteInstrument(data) {
  try {
    const ss    = getSpreadsheet(data);
    const sheet = ss.getSheetByName(data.instSheet || 'Instruments');
    if (!sheet) return { ok:false, msg:'Instruments sheet not found' };
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0].map(h => String(h).toLowerCase().trim());
    const idIdx = headers.indexOf('id');
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][idIdx]).trim() === String(data.id).trim()) {
        sheet.deleteRow(i + 1);
        return { ok:true, msg:'deleteInstrument success' };
      }
    }
    return { ok:false, msg:'Instrument not found: ' + data.id };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}

// ──────────── TEST HISTORY ────────────
function addTestHistory(data) {
  try {
    const ss    = getSpreadsheet(data);
    const sheet = ss.getSheetByName('TestHistory') || ss.insertSheet('TestHistory');
    const COLS  = ['testId','date','module','serialNumbers','tester','passCount','failCount','total','overall'];
    ensureHeader(sheet, COLS);
    sheet.appendRow(COLS.map(c => data[c] || ''));
    return { ok:true, msg:'addTestHistory success' };
  } catch(e) { return { ok:false, msg:e.toString() }; }
}
