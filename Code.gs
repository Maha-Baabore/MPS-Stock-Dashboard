const BORROW_SHEET = 'BorrowReturn';
const SPREADSHEET_ID = '1c5tiZzbtOpp_5PuRKpw4PbKcmHei_LWg9YZP3XhxcaE';
const BORROW_COLS = ['id','borrower','dept','itemId','itemName','qty',
                     'borrowDate','dueDate','returnDate','status','purpose','returnNote'];

function getSpreadsheet(data) {
  const id = (data && data.sheetId) ? data.sheetId : SPREADSHEET_ID;
  return SpreadsheetApp.openById(id);
}

function doGet(e) {
  const res = ContentService.createTextOutput();
  res.setMimeType(ContentService.MimeType.JSON);
  try {
    const p = e.parameter;
    if (!p.action) {
      res.setContent(JSON.stringify({ ok: true, msg: 'MPS Stock GAS v1.3.6 — ready' }));
      return res;
    }
    const data = {};
    Object.keys(p).forEach(k => { data[k] = p[k]; });
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

function doPost(e) { return doGet(e); }

function addBorrow(data) {
  const ss    = getSpreadsheet(data);
  const sheet = ss.getSheetByName(BORROW_SHEET);
  if (!sheet) return { ok: false, msg: `Sheet "${BORROW_SHEET}" not found` };
  ensureHeader(sheet);
  const row = BORROW_COLS.map(col => data[col] ?? '');
  sheet.appendRow(row);
  updateStockQty(data.itemId, -parseInt(data.qty || 0), data);
  return { ok: true, msg: 'addBorrow success', id: data.id };
}

function updateReturn(data) {
  const ss    = getSpreadsheet(data);
  const sheet = ss.getSheetByName(BORROW_SHEET);
  if (!sheet) return { ok: false, msg: `Sheet "${BORROW_SHEET}" not found` };
  const rows = sheet.getDataRange().getValues();
  const idCol = BORROW_COLS.indexOf('id');
  const statusCol = BORROW_COLS.indexOf('status');
  const returnDateCol = BORROW_COLS.indexOf('returnDate');
  const noteCol = BORROW_COLS.indexOf('returnNote');
  const qtyCol  = BORROW_COLS.indexOf('qty');
  const itemCol = BORROW_COLS.indexOf('itemId');
  let found = false;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]) === String(data.id)) {
      const rowNum = i + 1;
      sheet.getRange(rowNum, statusCol + 1).setValue('returned');
      sheet.getRange(rowNum, returnDateCol + 1).setValue(data.returnDate);
      if (noteCol >= 0) sheet.getRange(rowNum, noteCol + 1).setValue(data.note || '');
      updateStockQty(rows[i][itemCol], parseInt(rows[i][qtyCol] || 0), data);
      found = true;
      break;
    }
  }
  if (!found) return { ok: false, msg: `BorrowID ${data.id} not found` };
  return { ok: true, msg: 'updateReturn success' };
}

function updateStockQty(itemId, delta, data) {
  try {
    const ss    = getSpreadsheet(data);
    const sheet = ss.getSheetByName('Stock');
    if (!sheet) return;
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0].map(h => String(h).toLowerCase().trim());
    const idIdx  = headers.findIndex(h => h === 'id');
    const qtyIdx = headers.findIndex(h => h === 'qty');
    if (idIdx < 0 || qtyIdx < 0) return;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][idIdx]).trim() === String(itemId).trim()) {
        const newQty = Math.max(0, parseInt(rows[i][qtyIdx] || 0) + delta);
        sheet.getRange(i + 1, qtyIdx + 1).setValue(newQty);
        break;
      }
    }
  } catch(e) { console.log('updateStockQty error:', e); }
}

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
