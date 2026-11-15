/**
 * ==============================================================================
 * CUB SCOUT PACK 170 — COMPLETE BIDIRECTIONAL BACKEND SCRIPT
 * ==============================================================================
 * Spreadsheet URL:
 * https://docs.google.com/spreadsheets/d/1moBG34-2_9bXo4ZDxrI8PfNvFZ63spCJJQhHMRTk8Aw/edit
 *
 * This script provides full 2-way integration between the website & Google Sheets:
 * 1. doGet(e): Reads live budget figures from the "Proposed Budget 2026-2027 Cub Scouting"
 *              tab and outputs the JSON needed by budget.html (ending balance, income,
 *              expenses, surplus inventory, and scout counts).
 * 2. doPost(e): Receives payment submissions from portal.html, auto-creates the
 *               "Payments" tab if missing, and appends the transaction with timestamps,
 *               Scout names, and generated Venmo reference codes.
 *
 * ------------------------------------------------------------------------------
 * HOW TO UPDATE IN GOOGLE APPS SCRIPT:
 * ------------------------------------------------------------------------------
 * 1. Open your Google Sheet -> Click "Extensions" -> "Apps Script".
 * 2. Select everything in Code.gs (Ctrl+A / Cmd+A) and replace it with THIS ENTIRE FILE.
 * 3. Click "Save" (disk icon).
 * 4. Click "Deploy" -> "Manage Deployments".
 * 5. Click the Pencil (Edit) icon next to your active Web App deployment.
 * 6. Under "Version", select "New version" (MANDATORY for changes to take effect).
 * 7. Confirm "Execute as": "Me" and "Who has access": "Anyone".
 * 8. Click "Deploy".
 * ==============================================================================
 */

// ==============================================================================
// 1. GET HANDLER: SERVES LIVE BUDGET DATA TO budget.html
// ==============================================================================
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'getBudget';

    if (action === 'getPayments') {
      return handleGetPayments();
    }

    return handleGetBudget();

  } catch (err) {
    console.error('Error in doGet:', err);
    return createJsonResponse({
      status: 'error',
      message: err.toString()
    }, 500);
  }
}

/**
 * Extracts live budget numbers from the proposed budget sheet
 */
function handleGetBudget() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Find the budget sheet
  var sheet = ss.getSheetByName('Proposed Budget 2026-2027 Cub Scouting') || 
              ss.getSheetByName('Budget') || 
              ss.getSheetByName('Summary') || 
              ss.getSheets()[0];
  
  var values = sheet.getDataRange().getValues();

  function parseVal(v) {
    if (typeof v === 'number') return v;
    if (!v) return 0;
    var str = String(v).replace(/[$,\s]/g, '');
    if (str.indexOf('(') !== -1 && str.indexOf(')') !== -1) {
      str = '-' + str.replace(/[()]/g, '');
    }
    var n = parseFloat(str);
    return isNaN(n) ? 0 : n;
  }

  // 1. Parse Summary (Bank balance, projected income, expenses, ending balance)
  var bankBalance = 0;
  var projectedIncome = 0;
  var projectedExpenses = 0;
  var endingBalance = 0;

  for (var r = 0; r < Math.min(10, values.length); r++) {
    var label = String(values[r][1] || '').trim().toLowerCase();
    var val = parseVal(values[r][2]);
    if (label.indexOf('bank balance') !== -1) {
      bankBalance = Math.abs(val);
    } else if (label.indexOf('projected income') !== -1) {
      projectedIncome = Math.abs(val);
    } else if (label.indexOf('projected expenses') !== -1) {
      projectedExpenses = Math.abs(val);
    } else if (label.indexOf('ending balance') !== -1) {
      endingBalance = val;
    }
  }

  // 2. Parse Scout Counts
  var scoutCounts = [];
  var scoutCountsRow = -1;
  for (var r = 0; r < values.length; r++) {
    var txt = String(values[r][1] || '').trim().toLowerCase();
    if (txt === 'scout counts') {
      scoutCountsRow = r;
      break;
    }
  }

  if (scoutCountsRow !== -1) {
    for (var r = scoutCountsRow + 1; r < values.length; r++) {
      var rank = String(values[r][1] || '').trim();
      if (!rank) break;
      var fullPrice = parseInt(values[r][2]) || 0;
      var reduced = parseInt(values[r][3]) || 0;
      scoutCounts.push({
        rank: rank,
        fullPrice: fullPrice,
        reduced: reduced
      });
    }
  }

  // 3. Parse Income Line Items
  var income = [];
  var incomeRow = -1;
  for (var r = 0; r < values.length; r++) {
    var txt = String(values[r][1] || '').trim().toLowerCase();
    if (txt === 'income') {
      incomeRow = r;
      break;
    }
  }

  if (incomeRow !== -1) {
    for (var r = incomeRow + 1; r < values.length; r++) {
      var item = String(values[r][1] || '').trim();
      if (!item || item.toUpperCase() === 'TOTAL') break;
      var qtyOnHand = parseInt(values[r][2]) || 0;
      var qtyNeeded = parseInt(values[r][3]) || 0;
      var amount = Math.abs(parseVal(values[r][5] !== '' && values[r][5] !== undefined ? values[r][5] : values[r][4]));
      income.push({
        item: item,
        qtyOnHand: qtyOnHand,
        qtyNeeded: qtyNeeded,
        amount: amount
      });
    }
  }

  // 4. Parse Expenses Line Items
  var expenses = [];
  var expensesRow = -1;
  for (var r = 0; r < values.length; r++) {
    var txt = String(values[r][1] || '').trim().toLowerCase();
    if (txt === 'expenses') {
      expensesRow = r;
      break;
    }
  }

  if (expensesRow !== -1) {
    for (var r = expensesRow + 1; r < values.length; r++) {
      var item = String(values[r][1] || '').trim();
      if (!item || item.toUpperCase() === 'TOTAL') break;
      var qtyOnHand = parseInt(values[r][2]) || 0;
      var qtyNeeded = parseInt(values[r][3]) || 0;
      var unitCost = Math.abs(parseVal(values[r][4]));
      var rawTotal = parseVal(values[r][5]);
      var note = String(values[r][6] || values[r][7] || '').trim();

      // Negative in sheet = cost (represented as positive in budget.html)
      // Positive in sheet with qtyOnHand > qtyNeeded = surplus (represented as negative in budget.html)
      var total = 0;
      if (rawTotal < 0) {
        total = Math.abs(rawTotal);
      } else if (rawTotal > 0 && qtyOnHand > qtyNeeded) {
        total = -Math.abs(rawTotal);
      } else {
        total = rawTotal;
      }

      expenses.push({
        item: item,
        qtyOnHand: qtyOnHand,
        qtyNeeded: qtyNeeded,
        unitCost: unitCost,
        total: total,
        note: note
      });
    }
  }

  var responseData = {
    updatedAt: new Date().toISOString(),
    summary: {
      bankBalance: bankBalance,
      projectedIncome: projectedIncome,
      projectedExpenses: projectedExpenses,
      endingBalance: endingBalance
    },
    scoutCounts: scoutCounts,
    income: income,
    expenses: expenses
  };

  return createJsonResponse(responseData);
}

// ==============================================================================
// 2. POST HANDLER: RECEIVES PAYMENTS FROM portal.html & LOGS TO "Payments"
// ==============================================================================
function doPost(e) {
  var lock = LockService.getScriptLock();
  var hasLock = lock.tryLock(30000);
  
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({ status: 'error', message: 'No post data received' }, 400);
    }

    var payload = {};
    try {
      payload = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      payload = e.parameter || {};
    }

    var action = payload.action || 'recordPayment';
    if (action === 'recordPayment') {
      return handleRecordPayment(payload);
    }

    return createJsonResponse({ status: 'error', message: 'Unknown action: ' + action }, 400);

  } catch (err) {
    console.error('Error in doPost:', err);
    return createJsonResponse({ status: 'error', message: err.toString() }, 500);
  } finally {
    if (hasLock) {
      lock.releaseLock();
    }
  }
}

/**
 * Appends payment to the "Payments" tab
 */
function handleRecordPayment(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'Payments';
  var sheet = ss.getSheetByName(sheetName);

  // Auto-create "Payments" tab if missing
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);

    var headers = [
      'Timestamp',
      'Payment ID',
      'Parent / Payer Name',
      'Email Address',
      'Scout Name',
      'Rank / Den',
      'Payment Purpose',
      'Amount ($)',
      'Venmo Memo Reference',
      'Status',
      'Notes'
    ];

    sheet.appendRow(headers);

    // Pack 170 Navy header styling
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#15446B');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    
    for (var c = 1; c <= headers.length; c++) {
      sheet.autoResizeColumn(c);
    }
  }

  function clean(val) {
    if (val === null || val === undefined) return '';
    var str = String(val).trim();
    if (/^[=+\-@\t\r]/.test(str)) {
      str = "'" + str;
    }
    return str;
  }

  var timestamp = data.timestamp ? new Date(data.timestamp) : new Date();
  var paymentId = clean(data.paymentId || ('DUES-' + Utilities.formatDate(timestamp, 'America/New_York', 'MMdd') + '-' + Math.random().toString(36).substring(2, 6).toUpperCase()));
  var parentName = clean(data.parentName || '');
  var email = clean(data.email || '');
  var scoutName = clean(data.scoutName || '');
  var rank = clean(data.rank || '');
  var purpose = clean(data.category || data.purpose || 'Pack Annual Dues');
  var amount = Number(data.amount) || 0;
  var venmoMemo = clean(data.venmoMemo || '');
  var status = clean(data.status || 'Pending Venmo Transfer');
  var notes = clean(data.notes || '');

  var row = [
    timestamp,
    paymentId,
    parentName,
    email,
    scoutName,
    rank,
    purpose,
    amount,
    venmoMemo,
    status,
    notes
  ];

  sheet.appendRow(row);

  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.getRange(lastRow, 8).setNumberFormat('$#,##0.00');
  sheet.getRange(lastRow, 2).setHorizontalAlignment('center');
  sheet.getRange(lastRow, 6).setHorizontalAlignment('center');
  sheet.getRange(lastRow, 10).setHorizontalAlignment('center');

  return createJsonResponse({
    status: 'success',
    message: 'Payment logged successfully',
    paymentId: paymentId,
    venmoMemo: venmoMemo,
    amount: amount,
    row: lastRow
  });
}

/**
 * Optional readout of recorded payments
 */
function handleGetPayments() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Payments');
  if (!sheet) {
    return createJsonResponse({ payments: [] });
  }

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return createJsonResponse({ payments: [] });
  }

  var headers = data[0];
  var rows = data.slice(1);
  var payments = rows.map(function(r) {
    var obj = {};
    for (var i = 0; i < headers.length; i++) {
      var key = String(headers[i]).toLowerCase().replace(/[^a-z0-9]/g, '');
      obj[key] = r[i];
    }
    return obj;
  });

  return createJsonResponse({
    payments: payments,
    totalRecords: payments.length,
    updatedAt: new Date().toISOString()
  });
}

/**
 * Helper to build JSON output
 */
function createJsonResponse(data, statusCode) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
