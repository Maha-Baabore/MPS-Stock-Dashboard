/* ════════════════════════════════════════════════════════════════
   MPS Stock Management System — Google Apps Script backend
   Complete Code.gs for dashboard v1.10.0
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

      // ── Stock ──────────────────────────────────────────────
      case 'updateStock':                 // delta-based qty change (from Internal Test pass)
        return updateStockDelta_(e);
      case 'addStockItem':
        return upsertRow_(e, e.parameter.stockSheet || 'Stock');
      case 'updateStockItem':
        return upsertRow_(e, e.parameter.stockSheet || 'Stock');
      case 'deleteStockItem':
        return deleteRow_(e, e.parameter.stockSheet || 'Stock');

      // ── Instruments ────────────────────────────────────────
      case 'addInstrument':
        return upsertRow_(e, e.parameter.instSheet || 'Instruments');
      case 'updateInstrument':
        return upsertRow_(e, e.parameter.instSheet || 'Instruments');
      case 'deleteInstrument':
        return deleteRow_(e, e.parameter.instSheet || 'Instruments');

      // ── Borrow / Return ────────────────────────────────────
      case 'addBorrow':
        return upsertRow_(e, e.parameter.borrowSheet || 'BorrowReturn');
      case 'updateReturn':
        return updateReturn_(e);

      // ── Test History ───────────────────────────────────────
      case 'addTestHistory':
        return addTestHistory_(e);

      // ── Material Master ────────────────────────────────────
      case 'addMaterial':
      case 'updateMaterial':
        return upsertRow_(e, e.parameter.materialsSheet || 'Materials',
                          ['id','name','category','partNo','brand','unit','location']);
      case 'deleteMaterial':
        return deleteRow_(e, e.parameter.materialsSheet || 'Materials');

      // ── Stock Transactions ledger ──────────────────────────
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

/* Delta-based stock update (used when an Internal Test passes).
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
    if (key === 'id') return id;
    if (key === 'name') return p.name || '';
    if (key === 'category') return p.category || '';
    if (key === 'brand') return p.brand || '';
    if (key === 'unit') return p.unit || '';
    if (key === 'qty') return delta;
    if (key === 'minqty') return 1;
    if (key === 'status') return 'ok';
    return '';
  });
  sh.appendRow(row);
  return jsonOut_({ ok: true, id: id, qty: delta, mode: 'create' });
}

/* Mark a borrow record as returned. Sets status + returnDate (+ note/timestamps if columns exist). */
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

/* Append a test-history record. */
function addTestHistory_(e) {
  var p = e.parameter;
  var sh = getSheet_(e, p.testSheet || 'TestHistory',
    ['testId','date','module','serialNumbers','tester','passCount','failCount','total','overall','timestamp']);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['testId','date','module','serialNumbers','tester','passCount','failCount','total','overall','timestamp']);
  }
  var headers = getHeaders_(sh);
  var map = {
    testid: p.testId || '',
    date: p.date || '',
    module: p.module || '',
    serialnumbers: p.serialNumbers || '',
    tester: p.tester || '',
    passcount: p.passCount || '',
    failcount: p.failCount || '',
    total: p.total || '',
    overall: p.overall || '',
    timestamp: p.timestamp || ''
  };
  var row = headers.map(function (h) {
    return map[String(h).toLowerCase().replace(/[\s_]/g, '')] || '';
  });
  sh.appendRow(row);
  return jsonOut_({ ok: true, testId: p.testId, mode: 'add' });
}

/* Append a stock-movement transaction to the StockTransactions ledger. */
function appendTxn_(e) {
  var p = e.parameter;
  var sh = getSheet_(e, p.txnSheet || 'StockTransactions',
    ['txnId','timestamp','type','itemId','itemName','qty','balanceAfter','reason','reference','user','note']);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['txnId','timestamp','type','itemId','itemName','qty','balanceAfter','reason','reference','user','note']);
  }
  var headers = getHeaders_(sh);
  var map = {
    txnid: p.txnId || '',
    timestamp: p.timestamp || '',
    type: p.type || '',
    itemid: p.itemId || '',
    itemname: p.itemName || '',
    qty: p.qty || '',
    balanceafter: p.balanceAfter || '',
    reason: p.reason || '',
    reference: p.reference || '',
    user: p.user || '',
    note: p.note || ''
  };
  var row = headers.map(function (h) {
    return map[String(h).toLowerCase().replace(/[\s_]/g, '')] || '';
  });
  sh.appendRow(row);
  return jsonOut_({ ok: true, txnId: p.txnId, type: p.type, mode: 'add' });
}

/* Control keys are NOT written as data columns. */
function isDataKey_(k) {
  var control = ['action','sheetId','stockSheet','borrowSheet','instSheet',
                 'materialsSheet','testSheet','txnSheet','delta','note'];
  return control.indexOf(k) < 0;
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
