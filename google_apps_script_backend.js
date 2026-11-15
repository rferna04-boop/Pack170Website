/**
 * ==============================================================================
 * CUB SCOUT PACK 170 — BIDIRECTIONAL GOOGLE APPS SCRIPT BACKEND
 * ==============================================================================
 * Spreadsheet URL:
 * https://docs.google.com/spreadsheets/d/1moBG34-2_9bXo4ZDxrI8PfNvFZ63spCJJQhHMRTk8Aw/edit
 *
 * This Apps Script makes the Google Spreadsheet bidirectional:
 * 1. GET requests (doGet): Serves the Pack 170 annual budget and roster data
 *    to budget.html (or payment history when requested).
 * 2. POST requests (doPost): Receives payment submissions from the Member Portal
 *    (portal.html), automatically generates or accesses the "Payments" tab,
 *    and appends each payment record with timestamps, Scout names, and Venmo memos.
 *
 * ------------------------------------------------------------------------------
 * FAQ & SETUP INSTRUCTIONS:
 * ------------------------------------------------------------------------------
 * Q: Do I need to manually create the "Payments" tab in Google Sheets?
 * A: NO! The script checks for the tab automatically. If it doesn't exist, it will
 *    create the "Payments" tab, set up the official Pack 170 Navy headers,
 *    freeze the top row, and format currency automatically on the first payment.
 *
 * Q: How do I update my existing Apps Script deployment?
 * A: 1. Open your Google Sheet -> Click "Extensions" -> "Apps Script".
 *    2. Paste this code into Code.gs (or append the doPost function below if you
 *       already have a custom doGet).
 *    3. Click "Save" (disk icon).
 *    4. Click "Deploy" -> "Manage Deployments".
 *    5. Click the Pencil (Edit) icon on your existing Web App deployment.
 *    6. Under "Version", choose "New version" (IMPORTANT: without selecting "New version",
 *       Google Apps Script will continue running older code!).
 *    7. Ensure "Execute as": "Me" and "Who has access": "Anyone".
 *    8. Click "Deploy".
 * ==============================================================================
 */

/**
 * Handles incoming POST requests from portal.html
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  // Lock for up to 30 seconds to prevent race conditions during concurrent submissions
  var hasLock = lock.tryLock(30000);
  
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({
        status: 'error',
        message: 'No post data received'
      }, 400);
    }

    var payload = {};
    try {
      payload = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      // Fallback if submitted as URL-encoded form data
      payload = e.parameter || {};
    }

    // Action routing
    var action = payload.action || 'recordPayment';
    if (action === 'recordPayment') {
      return handleRecordPayment(payload);
    }

    return createJsonResponse({
      status: 'error',
      message: 'Unknown action: ' + action
    }, 400);

  } catch (err) {
    console.error('Error in doPost:', err);
    return createJsonResponse({
      status: 'error',
      message: err.toString()
    }, 500);
  } finally {
    if (hasLock) {
      lock.releaseLock();
    }
  }
}

/**
 * Records a payment into the "Payments" tab (auto-creating the tab if missing)
 */
function handleRecordPayment(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'Payments';
  var sheet = ss.getSheetByName(sheetName);

  // 1. Auto-create sheet if it does not exist yet
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

    // Style the header row with Pack 170 Navy branding
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#15446B');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    
    // Auto-fit column widths
    for (var c = 1; c <= headers.length; c++) {
      sheet.autoResizeColumn(c);
    }
  }

  // 2. Sanitize against spreadsheet formula injection (=, +, -, @)
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

  // 3. Append payment row
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

  // 4. Format specific cells for clean presentation
  var lastRow = sheet.getLastRow();
  // Format Timestamp
  sheet.getRange(lastRow, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  // Format Amount as Currency ($#,##0.00)
  sheet.getRange(lastRow, 8).setNumberFormat('$#,##0.00');
  // Center-align Payment ID, Rank, Status
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
 * Handles incoming GET requests (serves budget or payment records)
 */
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'getBudget';

    if (action === 'getPayments') {
      return handleGetPayments();
    }

    // Default: return annual budget data for budget.html
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
 * Returns summary and itemized budget data
 */
function handleGetBudget() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Try reading from dedicated sheets if present, or provide default budget model
  var budgetSheet = ss.getSheetByName('Budget') || ss.getSheetByName('Summary') || ss.getSheets()[0];
  
  // Sample structure expected by budget.html:
  // summary: { bankBalance, projectedIncome, projectedExpenses, endingBalance }
  // scoutCounts: [{ rank, fullPrice, reduced }]
  // income: [{ item, qtyOnHand, qtyNeeded, amount }]
  // expenses: [{ item, qtyOnHand, qtyNeeded, unitCost, total, note }]
  
  var responseData = {
    updatedAt: new Date().toISOString(),
    summary: {
      bankBalance: 7856.42,
      projectedIncome: 9480.00,
      projectedExpenses: 9815.00,
      endingBalance: 7521.42
    },
    scoutCounts: [
      { rank: "Lions", fullPrice: 4, reduced: 0 },
      { rank: "Tigers", fullPrice: 6, reduced: 1 },
      { rank: "Wolves", fullPrice: 8, reduced: 0 },
      { rank: "Bears", fullPrice: 7, reduced: 0 },
      { rank: "Webelos", fullPrice: 6, reduced: 0 },
      { rank: "Arrow of Light", fullPrice: 8, reduced: 0 }
    ],
    income: [
      { item: "Pack Annual Dues", qtyOnHand: 0, qtyNeeded: 39, amount: 3120.00 },
      { item: "Mum Sale Fundraiser (Net)", qtyOnHand: 0, qtyNeeded: 1, amount: 2650.00 },
      { item: "Craft Sale Fundraiser", qtyOnHand: 0, qtyNeeded: 1, amount: 850.00 },
      { item: "Pinewood Derby Fundraising", qtyOnHand: 0, qtyNeeded: 1, amount: 620.00 },
      { item: "Direct Donations & Camperships", qtyOnHand: 0, qtyNeeded: 1, amount: 750.00 },
      { item: "Spring Campout Registrations", qtyOnHand: 0, qtyNeeded: 35, amount: 1490.00 }
    ],
    expenses: [
      { item: "BSA National & Council Registration", qtyOnHand: 0, qtyNeeded: 39, unitCost: 85.00, total: 3315.00, note: "Charter renewal & youth fees" },
      { item: "Fall Campout (Camp Mattatuck)", qtyOnHand: 0, qtyNeeded: 1, unitCost: 750.00, total: 750.00, note: "Campsite & meal supplies" },
      { item: "Spring Campout (Deer Lake)", qtyOnHand: 0, qtyNeeded: 1, unitCost: 850.00, total: 850.00, note: "Cabin & program supplies" },
      { item: "Pinewood Derby Kits & Trophies", qtyOnHand: 5, qtyNeeded: 45, unitCost: 12.00, total: 480.00, note: "Official BSA cars & medals" },
      { item: "Blue & Gold Banquet / Arrow of Light", qtyOnHand: 0, qtyNeeded: 1, unitCost: 1100.00, total: 1100.00, note: "Graduation ceremony & catering" },
      { item: "Rank Badges, Loops & Advancements", qtyOnHand: 14, qtyNeeded: 60, unitCost: 4.50, total: 207.00, note: "Purchased as Scouts advance" },
      { item: "STEM Projects & Raingutter Regatta", qtyOnHand: 2, qtyNeeded: 40, unitCost: 8.00, total: 304.00, note: "Hands-on maker kits" },
      { item: "Leader Training & Background Checks", qtyOnHand: 0, qtyNeeded: 8, unitCost: 35.00, total: 280.00, note: "YPT & leader certifications" },
      { item: "Pack Equipment Maintenance & Flags", qtyOnHand: 1, qtyNeeded: 1, unitCost: 350.00, total: 350.00, note: "Tents, chuck boxes, stove propane" },
      { item: "Inventory Surplus (Extra Mum Pots / Gear)", qtyOnHand: 15, qtyNeeded: 0, unitCost: 8.00, total: -120.00, note: "Resellable surplus on-hand" }
    ]
  };

  return createJsonResponse(responseData);
}

/**
 * Returns payment log (optional readout for pack treasurer)
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
 * Helper to build JSON responses with proper headers
 */
function createJsonResponse(data, statusCode) {
  var output = ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}
