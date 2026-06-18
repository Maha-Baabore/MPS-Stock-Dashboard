/* ════════════════════════════════════════════════════════════════
   MPS Stock Management System — Google Apps Script backend
   Code.gs  v1.15.4
   ----------------------------------------------------------------
   Updated from v1.11.0:
   - addTestHistory_  : uses testHistSheet param (not testSheet);
                        added fc_sn + inStock columns; seeds correct headers
   - updateTestHistory: NEW — flips inStock flag after unit enters inventory
   - clearTestHistory : NEW — deletes all rows in TestHistory sheet
   - deleteTestHistory: NEW — deletes one row by testId
   - isDataKey_       : added testHistSheet to control-key exclusion list
   ----------------------------------------------------------------
   This is a FULL replacement file. Select all in your Code.gs,
   delete, paste this, Save, then DEPLOY A NEW VERSION:
     Deploy → Manage deployments → (pencil/Edit) → Version: New version → Deploy
   Saving alone does NOT update the live web app.
   ----------------------------------------------------------------
   Design notes:
   - All writes come in as GET requests with URL params (no-cors POST
     cannot send a body), so everything is read from e.parameter.
   - Column mapping is HEADER-BASED: the script reads row 1 of each
     sheet and matches by column name, so column order can change
     without breaking writes. New columns the frontend sends are
     auto-appended to the header row.
   ════════════════════════════════════════════════════════════════ */

function doGet(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); // serialize writes to avoid race conditions
  } catch (err) {
    return jsonOut_({ ok: false, error: 'busy, try again' });
  }

  try {
    var action = (e && e.parameter && e.parameter.action) || '';
    switch (action) {

      // ── Stock ──────────────────────────────────────────────────
      case 'updateStock':                 // delta-based qty change (legacy)
        return updateStockDelta_(e);
      case 'addStockItem':
        return upsertRow_(e, e.parameter.stockSheet || 'Stock');
      case 'updateStockItem':
        return upsertRow_(e, e.parameter.stockSheet || 'Stock');
      case 'deleteStockItem':
        return deleteRow_(e, e.parameter.stockSheet || 'Stock');

      // ── Instruments ────────────────────────────────────────────
      case 'addInstrument':
        return upsertRow_(e, e.parameter.instSheet || 'Instruments');
      case 'updateInstrument':
        return upsertRow_(e, e.parameter.instSheet || 'Instruments');
      case 'deleteInstrument':
        return deleteRow_(e, e.parameter.instSheet || 'Instruments');

      // ── Borrow / Return ────────────────────────────────────────
      case 'addBorrow':
        return upsertRow_(e, e.parameter.borrowSheet || 'BorrowReturn');
      case 'updateReturn':
        return updateReturn_(e);

      // ── Test History ───────────────────────────────────────────
      case 'addTestHistory':
        return addTestHistory_(e);
      case 'updateTestHistory':           // NEW v1.15.4: flip inStock flag
        return updateTestHistory_(e);
      case 'clearTestHistory':            // NEW v1.15.4: clear all rows
        return clearTestHistory_(e);
      case 'deleteTestHistory':           // NEW v1.15.4: delete one row
        return deleteTestHistoryRow_(e);

      // ── Material Master ────────────────────────────────────────
      case 'addMaterial':
      case 'updateMaterial':
        return upsertRow_(e, e.parameter.materialsSheet || 'Materials',
                          ['id','name','category','partNo','brand','unit','location']);
      case 'deleteMaterial':
        return deleteRow_(e, e.parameter.materialsSheet || 'Materials');

      // ── Stock Transactions ledger ──────────────────────────────
      case 'addTxn':
        return appendTxn_(e);

      default:
        return jsonOut_({ ok: false, error: 'unknown action: ' + action });
    }
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ════════════════════════════════════════════════════════════════
   CORE HELPERS
   ════════════════════════════════════════════════════════════════ */

function getSheet_(e, sheetName, seedHeaders) {
  var ss = SpreadsheetApp.openById(e.parameter.sheetId);
  var sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    if (seedHeaders && seedHeaders.length) sh.appendRow(seedHeaders);
  }
  return sh;
}

// Return the header row (row 1) as an array of trimmed strings.
function getHeaders_(sh) {
  if (sh.getLastRow() < 1) return [];
  var lastCol = Math.max(1, sh.getLastColumn());
  return sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h).trim();
  });
}

// Find the column index (0-based) of a header, case/space-insensitive.
function headerIndex_(headers, name) {
  var target = String(name).toLowerCase().replace(/[\s_]/g, '');
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).toLowerCase().replace(/[\s_]/g, '') === target) return i;
  }
  return -1;
}

/* Generic upsert: writes all incoming params (except control keys) to a row,
   keyed by the "id" column. Inserts if id not found, updates otherwise.
   If a param has no matching header column, the header is auto-appended. */
function upsertRow_(e, sheetName, seedHeaders) {
  var p = e.parameter;
  var id = String(p.id || '').trim();
  if (!id) return jsonOut_({ ok: false, error: 'missing id' });

  var sh = getSheet_(e, sheetName, seedHeaders);
  if (sh.getLastRow() === 0 && seedHeaders) sh.appendRow(seedHeaders);

  var headers = getHeaders_(sh);
  if (!headers.length) {
    // No header row yet — seed from incoming keys
    headers = Object.keys(p).filter(isDataKey_);
    if (headers.indexOf('id') < 0) headers.unshift('id');
    sh.appendRow(headers);
  }

  // Collect the data fields the frontend sent (skip control keys)
  var dataKeys = Object.keys(p).filter(isDataKey_);

  // Auto-append any new columns the sheet doesn't have yet
  var appended = false;
  dataKeys.forEach(function (k) {
    if (headerIndex_(headers, k) < 0) {
      headers.push(k);
      appended = true;
    }
  });
  if (appended) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  // Build the row in header order
  var idCol = headerIndex_(headers, 'id');

  // Find existing row by id
  var lastRow = sh.getLastRow();
  var foundRow = -1;
  if (lastRow >= 2) {
    var idColValues = sh.getRange(2, idCol + 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < idColValues.length; i++) {
      if (String(idColValues[i][0]).trim().toLowerCase() === id.toLowerCase()) {
        foundRow = i + 2;
        break;
      }
    }
  }

  if (foundRow > 0) {
    // UPDATE: overwrite only the columns the frontend sent; keep others intact
    var existing = sh.getRange(foundRow, 1, 1, headers.length).getValues()[0];
    dataKeys.forEach(function (k) {
      var ci = headerIndex_(headers, k);
      if (ci >= 0) existing[ci] = p[k];
    });
    sh.getRange(foundRow, 1, 1, headers.length).setValues([existing]);
    return jsonOut_({ ok: true, id: id, mode: 'update', row: foundRow });
  } else {
    // INSERT
    var row = headers.map(function (h) {
      var key = dataKeys.filter(function (k) {
        return String(k).toLowerCase().replace(/[\s_]/g, '') === String(h).toLowerCase().replace(/[\s_]/g, '');
      })[0];
      return key ? p[key] : '';
    });
    sh.appendRow(row);
    return jsonOut_({ ok: true, id: id, mode: 'add', row: sh.getLastRow() });
  }
}

// Delete a row by id from any sheet.
function deleteRow_(e, sheetName) {
  var id = String(e.parameter.id || '').trim();
  if (!id) return jsonOut_({ ok: false, error: 'missing id' });
  var sh = getSheet_(e, sheetName);
  var headers = getHeaders_(sh);
  var idCol = headerIndex_(headers, 'id');
  if (idCol < 0) return jsonOut_({ ok: false, error: 'no id column' });
  var lastRow = sh.getLastRow();
  for (var r = 2; r <= lastRow; r++) {
    var cell = String(sh.getRange(r, idCol + 1).getValue()).trim();
    if (cell.toLowerCase() === id.toLowerCase()) {
      sh.deleteRow(r);
      return jsonOut_({ ok: true, id: id, deleted: true });
    }
  }
  return jsonOut_({ ok: false, error: 'not found', id: id });
}

/* Delta-based stock update (legacy — used before ledger architecture).
   Increments qty for an existing id, or creates the row if missing. */
function updateStockDelta_(e) {
  var p = e.parameter;
  var id = String(p.id || '').trim();
  if (!id) return jsonOut_({ ok: false, error: 'missing id' });
  var delta = parseInt(p.delta, 10) || 0;

  var sh = getSheet_(e, p.stockSheet || 'Stock',
                     ['id','name','category','partNo','brand','qty','minQty','unit','location','status']);
  var headers = getHeaders_(sh);
  var idCol = headerIndex_(headers, 'id');
  var qtyCol = headerIndex_(headers, 'qty');
  var lastRow = sh.getLastRow();

  for (var r = 2; r <= lastRow; r++) {
    if (String(sh.getRange(r, idCol + 1).getValue()).trim().toLowerCase() === id.toLowerCase()) {
      var cur = parseInt(sh.getRange(r, qtyCol + 1).getValue(), 10) || 0;
      sh.getRange(r, qtyCol + 1).setValue(cur + delta);
      return jsonOut_({ ok: true, id: id, qty: cur + delta, mode: 'increment' });
    }
  }
  // Not found → create
  var row = headers.map(function (h) {
    var key = h.toLowerCase();
    if (key === 'id')       return id;
    if (key === 'name')     return p.name || '';
    if (key === 'category') return p.category || '';
    if (key === 'brand')    return p.brand || '';
    if (key === 'unit')     return p.unit || '';
    if (key === 'qty')      return delta;
    if (key === 'minqty')   return 1;
    if (key === 'status')   return 'ok';
    return '';
  });
  sh.appendRow(row);
  return jsonOut_({ ok: true, id: id, qty: delta, mode: 'create' });
}

/* Mark a borrow record as returned. Sets status + returnDate + note. */
function updateReturn_(e) {
  var p = e.parameter;
  var id = String(p.id || '').trim();
  if (!id) return jsonOut_({ ok: false, error: 'missing id' });

  var sh = getSheet_(e, p.borrowSheet || 'BorrowReturn');
  var headers = getHeaders_(sh);
  var idCol = headerIndex_(headers, 'id');
  var lastRow = sh.getLastRow();

  for (var r = 2; r <= lastRow; r++) {
    if (String(sh.getRange(r, idCol + 1).getValue()).trim().toLowerCase() === id.toLowerCase()) {
      setIfHeader_(sh, headers, r, 'status', 'returned');
      setIfHeader_(sh, headers, r, 'returnDate', p.returnDate || '');
      setIfHeader_(sh, headers, r, 'returnNote', p.note || '');
      return jsonOut_({ ok: true, id: id, mode: 'return' });
    }
  }
  return jsonOut_({ ok: false, error: 'borrow id not found', id: id });
}

function setIfHeader_(sh, headers, row, headerName, value) {
  var ci = headerIndex_(headers, headerName);
  if (ci >= 0) sh.getRange(row, ci + 1).setValue(value);
}

/* ════════════════════════════════════════════════════════════════
   TEST HISTORY
   ════════════════════════════════════════════════════════════════ */

/* Append a test-history record.
   Updated v1.15.4:
   - Uses testHistSheet param (frontend sends testHistSheet, not testSheet)
   - Seed headers now include fc_sn and inStock columns
   - map keys cover all fields the frontend sends                    */
function addTestHistory_(e) {
  var p  = e.parameter;
  var sh = getSheet_(e, p.testHistSheet || 'TestHistory',
    ['testId','date','module','serialNumbers','fc_sn','tester',
     'passCount','failCount','total','overall','inStock','timestamp']);

  // If the sheet was just created and has no header row, seed it
  if (sh.getLastRow() === 0) {
    sh.appendRow(['testId','date','module','serialNumbers','fc_sn','tester',
                  'passCount','failCount','total','overall','inStock','timestamp']);
  }

  var headers = getHeaders_(sh);

  // Auto-append columns the frontend sends that the sheet doesn't have yet
  var incomingKeys = ['testId','date','module','serialNumbers','fc_sn','tester',
                      'passCount','failCount','total','overall','inStock','timestamp'];
  var appended = false;
  incomingKeys.forEach(function(k) {
    if (headerIndex_(headers, k) < 0) {
      headers.push(k);
      appended = true;
    }
  });
  if (appended) sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  var map = {
    testid:        p.testId        || '',
    date:          p.date          || '',
    module:        p.module        || '',
    serialnumbers: p.serialNumbers || '',
    fcsn:          p.fc_sn         || '',
    tester:        p.tester        || '',
    passcount:     p.passCount     || '',
    failcount:     p.failCount     || '',
    total:         p.total         || '',
    overall:       p.overall       || '',
    instock:       p.inStock       || 'false',
    timestamp:     p.timestamp     || ''
  };

  var row = headers.map(function(h) {
    return map[String(h).toLowerCase().replace(/[\s_]/g, '')] || '';
  });
  sh.appendRow(row);
  SpreadsheetApp.flush();
  return jsonOut_({ ok: true, testId: p.testId, mode: 'add' });
}

/* NEW v1.15.4 — Flip inStock to TRUE for a specific testId.
   Called after recordTxn('IN') succeeds in the dashboard.           */
function updateTestHistory_(e) {
  var p        = e.parameter;
  var targetId = String(p.testId || '').trim();
  if (!targetId) return jsonOut_({ ok: false, error: 'missing testId' });

  var sh      = getSheet_(e, p.testHistSheet || 'TestHistory');
  var headers = getHeaders_(sh);
  var colId   = headerIndex_(headers, 'testId');
  var colStk  = headerIndex_(headers, 'inStock');

  if (colId < 0)  return jsonOut_({ ok: false, error: 'testId column not found' });
  if (colStk < 0) {
    // inStock column missing → auto-append it
    headers.push('inStock');
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    colStk = headers.length - 1;
  }

  var lastRow = sh.getLastRow();
  for (var r = 2; r <= lastRow; r++) {
    var cell = String(sh.getRange(r, colId + 1).getValue()).trim();
    if (cell === targetId) {
      var newVal = String(p.inStock || 'true').toLowerCase() === 'true';
      sh.getRange(r, colStk + 1).setValue(newVal);
      SpreadsheetApp.flush();
      return jsonOut_({ ok: true, testId: targetId, inStock: newVal, row: r });
    }
  }
  return jsonOut_({ ok: false, error: 'testId not found', testId: targetId });
}

/* NEW v1.15.4 — Delete ALL data rows from TestHistory (keep header row).
   Called when Admin presses the clear-history button.               */
function clearTestHistory_(e) {
  var p  = e.parameter;
  var sh = getSheet_(e, p.testHistSheet || 'TestHistory');
  var lastRow = sh.getLastRow();
  if (lastRow <= 1) return jsonOut_({ ok: true, deleted: 0 }); // nothing to clear
  sh.deleteRows(2, lastRow - 1);
  SpreadsheetApp.flush();
  return jsonOut_({ ok: true, deleted: lastRow - 1 });
}

/* NEW v1.15.4 — Delete one row from TestHistory by testId.
   Called when user clicks the ✕ button on a history row.           */
function deleteTestHistoryRow_(e) {
  var p        = e.parameter;
  var targetId = String(p.testId || '').trim();
  if (!targetId) return jsonOut_({ ok: false, error: 'missing testId' });

  var sh      = getSheet_(e, p.testHistSheet || 'TestHistory');
  var headers = getHeaders_(sh);
  var colId   = headerIndex_(headers, 'testId');
  if (colId < 0) return jsonOut_({ ok: false, error: 'testId column not found' });

  var lastRow = sh.getLastRow();
  for (var r = 2; r <= lastRow; r++) {
    var cell = String(sh.getRange(r, colId + 1).getValue()).trim();
    if (cell === targetId) {
      sh.deleteRow(r);
      SpreadsheetApp.flush();
      return jsonOut_({ ok: true, testId: targetId, deleted: true, row: r });
    }
  }
  return jsonOut_({ ok: false, error: 'testId not found', testId: targetId });
}

/* ════════════════════════════════════════════════════════════════
   STOCK TRANSACTIONS LEDGER
   ════════════════════════════════════════════════════════════════ */

/* Append a stock-movement transaction to the StockTransactions ledger. */
function appendTxn_(e) {
  var p  = e.parameter;
  var sh = getSheet_(e, p.txnSheet || 'StockTransactions',
    ['txnId','timestamp','type','itemId','itemName','qty','balanceAfter','reason','reference','user','note']);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['txnId','timestamp','type','itemId','itemName','qty','balanceAfter','reason','reference','user','note']);
  }
  var headers = getHeaders_(sh);
  var map = {
    txnid:        p.txnId        || '',
    timestamp:    p.timestamp    || '',
    type:         p.type         || '',
    itemid:       p.itemId       || '',
    itemname:     p.itemName     || '',
    qty:          p.qty          || '',
    balanceafter: p.balanceAfter || '',
    reason:       p.reason       || '',
    reference:    p.reference    || '',
    user:         p.user         || '',
    note:         p.note         || ''
  };
  var row = headers.map(function(h) {
    return map[String(h).toLowerCase().replace(/[\s_]/g, '')] || '';
  });
  sh.appendRow(row);
  SpreadsheetApp.flush();
  return jsonOut_({ ok: true, txnId: p.txnId, type: p.type, mode: 'add' });
}

/* ════════════════════════════════════════════════════════════════
   UTILITIES
   ════════════════════════════════════════════════════════════════ */

/* Control keys are NOT written as data columns. */
function isDataKey_(k) {
  var control = ['action','sheetId','stockSheet','borrowSheet','instSheet',
                 'materialsSheet','testSheet','testHistSheet','txnSheet','delta','note'];
  return control.indexOf(k) < 0;
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
