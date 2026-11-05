// ============================================================
//  REDBIRD CRM — Master Apps Script  (consolidated)
//  Deploy as: Execute as "Me", Access "Anyone"
//
//  Folds in every portal-reservation change:
//    • writeReservation  — writes customer Notes to column 12
//    • updateReservation — syncs Notes across the reservation's rows
//    • cancelReservation — sends a cancellation confirmation email
//    • getReservation    — NEW doGet endpoint: returns a saved reservation
//    • sendCancellationEmail — NEW
//    • sendOrderConfirmationEmail — portal link carries &reservation=<id>
//    • Email logo replaced with a lightweight text wordmark
//
//  AFTER PASTING: Deploy → Manage deployments → edit → Version: New version.
// ============================================================

// ── CONFIG ───────────────────────────────────────────────────
const ACTIVITY_LOG_TAB   = 'Activity_Log';
const ORDERS_TAB         = 'Orders Log';
const TASKS_TAB          = 'tasks';
const ACCOUNTS_DEST_TAB  = 'Accounts Mirror';
const INVENTORY_DEST_TAB = 'Inventory Projections Mirror';

// Every reservation email (create / update / cancel) is BCC'd to this
// address so the Redbird team has a copy of all reservation activity.
// Comma-separated if more than one; edit here to change who gets copied.
const RESERVATION_BCC = 'redbird@redbird-cannabis.com';

const ACCOUNTS_SOURCE_ID    = '15dJswwQ1Pyt8VsAqtH17PJnVTYPWrkXZnZlWGCIKjd0';
const ACCOUNTS_SOURCE_TAB   = 'Active WA Retailers';

const INVENTORY_SOURCE_ID   = '1IuSfAblyYWSO5gYZhdgL_KdaWSfxG7sjvBWv2zWURJ4';
const INVENTORY_SOURCE_TAB  = 'Product Allocation';
const INVENTORY_DATE_COL    = 0;
const INVENTORY_NUM_COLS    = 31;

const ORDERS_HEADERS = [
  'Timestamp', 'Account Name', 'Opportunity Name',
  'Close Date', 'Stage', 'Product Name', 'Sales Price', 'Quantity'
];

const TASKS_HEADERS = [
  'Account Name', 'Subject', 'Related To: Name', 'Completed Date/Time', 'Date'
];

// ── RESPONSE BUILDER ─────────────────────────────────────────
function buildResponse(data, callback) {
  var json = JSON.stringify(data);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ── SHEET HELPERS ─────────────────────────────────────────────
function getActiveSheet(name) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function getSheetByGid(spreadsheet, gid) {
  return spreadsheet.getSheets().find(function(s) {
    return s.getSheetId() === gid;
  }) || null;
}

function styleHeader(sheet, numCols) {
  var header = sheet.getRange(1, 1, 1, numCols);
  header.setBackground('#C8102E');
  header.setFontColor('#FFFFFF');
  header.setFontWeight('bold');
}

// ── ACCOUNT LOOKUP — pulls account details from CRM Accounts ──
function getAccountInfo(accountName) {
  if (!accountName) return null;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var target = String(accountName).toLowerCase().trim();

    var crm = ss.getSheetByName('CRM Accounts');
    if (crm && crm.getLastRow() >= 2) {
      var crmData = crm.getDataRange().getValues();
      var ch = crmData[0];
      var idx = {
        trade:   ch.indexOf('Trade Name'),
        lic:     ch.indexOf('License #'),
        ubi:     ch.indexOf('UBI'),
        addr:    ch.indexOf('Mailing Street')  > -1 ? ch.indexOf('Mailing Street')  : ch.indexOf('Address'),
        city:    ch.indexOf('Mailing City')    > -1 ? ch.indexOf('Mailing City')    : ch.indexOf('City'),
        state:   ch.indexOf('Mailing State')   > -1 ? ch.indexOf('Mailing State')   : ch.indexOf('State'),
        zip:     ch.indexOf('Mailing Zip')     > -1 ? ch.indexOf('Mailing Zip')     : ch.indexOf('Zip'),
        county:  ch.indexOf('County'),
        source:  ch.indexOf('Source'),
        firstO:  ch.indexOf('First Order'),
        lastO:   ch.indexOf('Last Order'),
        orderCt: ch.indexOf('Order Count'),
        notes:   ch.indexOf('Notes'),
        manual:  ch.indexOf('Manual Edit')
      };
      if (idx.trade > -1) {
        for (var i = 1; i < crmData.length; i++) {
          var name = String(crmData[i][idx.trade] || '').toLowerCase().trim();
          if (name === target) {
            function pick(k) { return idx[k] > -1 ? (crmData[i][idx[k]] || '') : ''; }
            return {
              tradeName:  crmData[i][idx.trade] || '',
              license:    pick('lic'),
              ubi:        pick('ubi'),
              address:    pick('addr'),
              city:       pick('city'),
              county:     pick('county'),
              state:      pick('state'),
              zip:        pick('zip'),
              source:     pick('source'),
              firstOrder: pick('firstO'),
              lastOrder:  pick('lastO'),
              orderCount: pick('orderCt'),
              notes:      pick('notes'),
              manualEdit: idx.manual > -1 ? !!crmData[i][idx.manual] : false
            };
          }
        }
      }
    }

    var sheet = ss.getSheetByName(ACCOUNTS_DEST_TAB);
    if (!sheet) { Logger.log('[getAccountInfo] No CRM Accounts and no Accounts Mirror'); return null; }
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;
    var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0] || '').toLowerCase().trim() === target) {
        return {
          tradeName: data[i][0] || '',
          license:   data[i][1] || '',
          ubi:       data[i][2] || '',
          address:   data[i][3] || '',
          city:      data[i][4] || '',
          county:    data[i][5] || '',
          state:     data[i][6] || '',
          zip:       data[i][7] || '',
          source:    'Mirror'
        };
      }
    }
    Logger.log('[getAccountInfo] No match for: ' + accountName);
    return null;
  } catch(err) {
    Logger.log('[getAccountInfo] error: ' + err.message);
    return null;
  }
}

// ── ACTIVITY LOG SHEET SETUP ──────────────────────────────────
function getActivityLogSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ACTIVITY_LOG_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(ACTIVITY_LOG_TAB);
    sheet.getRange(1, 1, 1, 6).setValues([['Timestamp', 'Account', 'Type', 'Notes', 'User', 'Date']]);
    styleHeader(sheet, 6);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 220);
    sheet.setColumnWidth(3, 90);
    sheet.setColumnWidth(4, 400);
    sheet.setColumnWidth(5, 100);
    sheet.setColumnWidth(6, 110);
  }
  return sheet;
}

// ── PURCHASE ORDERS SHEET SETUP ───────────────────────────────
function _getOrCreatePOSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var name  = 'Purchase Orders';
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, 10).setValues([[
      'Timestamp', 'PO Number', 'Account', 'Recipients',
      'Item Count', 'Unit Count', 'Total $', 'Items JSON', 'Sent By', 'Date Label'
    ]]);
    styleHeader(sheet, 10);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 140);
    sheet.setColumnWidth(3, 220);
    sheet.setColumnWidth(4, 280);
    sheet.setColumnWidth(5, 80);
    sheet.setColumnWidth(6, 80);
    sheet.setColumnWidth(7, 90);
    sheet.setColumnWidth(8, 80);
    sheet.setColumnWidth(9, 100);
    sheet.setColumnWidth(10, 140);
  }
  return sheet;
}

// ── ONE-TIME MIGRATION: Purchase Orders backfill ──────────────
function migratePurchaseOrders() {
  var t0 = Date.now();
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var activity  = ss.getSheetByName(ACTIVITY_LOG_TAB);
  if (!activity) {
    console.log('No Activity Log tab found — nothing to migrate.');
    return { ok: false, error: 'no activity log' };
  }

  var poSheet = _getOrCreatePOSheet();

  var existing = {};
  var poLast = poSheet.getLastRow();
  if (poLast > 1) {
    var existingData = poSheet.getRange(2, 1, poLast - 1, 3).getValues();
    existingData.forEach(function(r) {
      var key = String(r[0]) + '|' + String(r[2]);
      existing[key] = true;
    });
  }

  var actLast = activity.getLastRow();
  if (actLast < 2) {
    console.log('Activity log is empty.');
    return { ok: true, scanned: 0, migrated: 0, skipped: 0 };
  }
  var actData = activity.getRange(2, 1, actLast - 1, 6).getValues();

  var scanned = 0, migrated = 0, skipped = 0, parsedFailed = 0;
  var rowsToAppend = [];

  for (var i = 0; i < actData.length; i++) {
    var row = actData[i];
    var ts     = row[0];
    var acct   = row[1];
    var type   = row[2];
    var notes  = String(row[3] || '');
    var user   = row[4] || 'Unknown';
    var dLabel = row[5] || '';

    if (type !== 'email') continue;
    var isPO = /\[Purchase Order Proposal\]/i.test(notes)
            || /^Purchase Order Proposal sent/i.test(notes);
    if (!isPO) continue;
    scanned++;

    var key = String(ts) + '|' + String(acct);
    if (existing[key]) { skipped++; continue; }

    var recipMatch = notes.match(/Sent to (\d+) recipient/i)
                  || notes.match(/sent to (\d+) recipient/i);
    var recipListMatch = notes.match(/recipient\(s\):\s*([^.]+?)\.\s*\d+\s+SKUs/i);
    var skuMatch = notes.match(/(\d+)\s+SKUs/i);
    var unitMatch = notes.match(/(\d+)\s+units/i);
    var totalMatch = notes.match(/\$([\d,]+(?:\.\d{2})?)\s*total/i);

    if (!skuMatch && !unitMatch && !totalMatch) { parsedFailed++; continue; }

    var recipList  = recipListMatch ? recipListMatch[1].trim() : '';
    var itemCount  = skuMatch  ? parseInt(skuMatch[1])   : 0;
    var unitCount  = unitMatch ? parseInt(unitMatch[1])  : 0;
    var totalNum   = totalMatch ? parseFloat(totalMatch[1].replace(/,/g,'')) : 0;

    rowsToAppend.push([
      ts, 'PROP-MIGRATED-' + (i + 1), acct, recipList,
      itemCount, unitCount, totalNum, '', user, dLabel
    ]);
    migrated++;
  }

  if (rowsToAppend.length > 0) {
    var startRow = poSheet.getLastRow() + 1;
    poSheet.getRange(startRow, 1, rowsToAppend.length, 10).setValues(rowsToAppend);
  }

  console.log('Migration complete in ' + Math.round((Date.now()-t0)/1000) + 's');
  console.log('Scanned: ' + scanned + ' | Migrated: ' + migrated + ' | Skipped: ' + skipped + ' | Parse failed: ' + parsedFailed);
  return { ok: true, scanned: scanned, migrated: migrated, skipped: skipped, parsedFailed: parsedFailed };
}

// ── ORDERS LOG SHEET SETUP ────────────────────────────────────
function getOrdersSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ORDERS_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(ORDERS_TAB);
    sheet.getRange(1, 1, 1, ORDERS_HEADERS.length).setValues([ORDERS_HEADERS]);
    styleHeader(sheet, ORDERS_HEADERS.length);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── TASKS SHEET SETUP ─────────────────────────────────────────
function getTasksSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TASKS_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(TASKS_TAB);
    sheet.getRange(1, 1, 1, TASKS_HEADERS.length).setValues([TASKS_HEADERS]);
    styleHeader(sheet, TASKS_HEADERS.length);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 220);
    sheet.setColumnWidth(2, 300);
    sheet.setColumnWidth(3, 220);
    sheet.setColumnWidth(4, 160);
    sheet.setColumnWidth(5, 120);
  }
  return sheet;
}

// ── PORTAL RESERVATIONS SHEET SETUP ────────────────────────────
function getReservationsSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var tabName = 'Portal Reservations';
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    var headers = ['Timestamp', 'Reservation ID', 'Account', 'Delivery Week', 'Strain', 'Product Size', 'Qty', 'Unit Price', 'Line Total', 'Status', 'Cancelled At', 'Notes'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeader(sheet, headers.length);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 140);
    sheet.setColumnWidth(3, 220);
    sheet.setColumnWidth(4, 140);
    sheet.setColumnWidth(5, 180);
    sheet.setColumnWidth(6, 120);
    sheet.setColumnWidth(7, 60);
    sheet.setColumnWidth(8, 90);
    sheet.setColumnWidth(9, 100);
    sheet.setColumnWidth(10, 90);
    sheet.setColumnWidth(11, 160);
    sheet.setColumnWidth(12, 320);
  } else {
    var lastCol = sheet.getLastColumn();
    if (lastCol < 12) {
      sheet.getRange(1, 12).setValue('Notes');
      sheet.getRange(1, 12).setBackground('#C8102E').setFontColor('#FFFFFF').setFontWeight('bold');
      sheet.setColumnWidth(12, 320);
    }
  }
  return sheet;
}

// ── TASK COMPLETIONS SHEET SETUP ───────────────────────────────
function getTaskCompletionsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tabName = 'Task Completions';
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    var headers = ['Task Key', 'Account', 'Task Name', 'Due Date', 'Completed At', 'Completed By', 'Notes'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeader(sheet, headers.length);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 360);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 280);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 160);
    sheet.setColumnWidth(6, 120);
    sheet.setColumnWidth(7, 320);
  }
  return sheet;
}

function upsertTaskCompletion(taskKey, account, taskName, dueDate, user, notes) {
  var sheet = getTaskCompletionsSheet();
  var lastRow = sheet.getLastRow();
  var now = new Date();
  var completedAt = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a');

  if (lastRow >= 2) {
    var keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (String(keys[i][0]).toLowerCase() === String(taskKey).toLowerCase()) {
        var rowNum = i + 2;
        sheet.getRange(rowNum, 1, 1, 7).setValues([[taskKey, account, taskName, dueDate, completedAt, user, notes]]);
        return 'updated';
      }
    }
  }
  sheet.appendRow([taskKey, account, taskName, dueDate, completedAt, user, notes]);
  return 'inserted';
}

function deleteTaskCompletion(taskKey) {
  var sheet = getTaskCompletionsSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  var keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < keys.length; i++) {
    if (String(keys[i][0]).toLowerCase() === String(taskKey).toLowerCase()) {
      sheet.deleteRow(i + 2);
      return true;
    }
  }
  return false;
}

// ── CONTACTS SHEET SETUP ───────────────────────────────────────
function getContactsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tabName = 'Contacts';
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    var headers = ['Contact ID', 'Trade Name', 'First Name', 'Last Name', 'Title', 'Email', 'Phone', 'Mobile',
                   'Mailing Street', 'Mailing City', 'Mailing State', 'Mailing Zip',
                   'Notes', 'Created At', 'Updated At'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeader(sheet, headers.length);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 130);
    sheet.setColumnWidth(2, 220);
    sheet.setColumnWidth(6, 240);
    sheet.setColumnWidth(9, 220);
    sheet.setColumnWidth(13, 320);
  } else {
    _ensureContactsSchemaUpgrade(sheet);
  }
  return sheet;
}

function _ensureContactsSchemaUpgrade(sheet) {
  if (sheet.getLastColumn() < 1) return;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return String(h).trim(); });
  if (headers.indexOf('Mailing Street') !== -1) return;

  var notesIdx = headers.indexOf('Notes');
  if (notesIdx === -1) return;

  sheet.insertColumnsBefore(notesIdx + 1, 4);
  var insertHeaders = [['Mailing Street', 'Mailing City', 'Mailing State', 'Mailing Zip']];
  sheet.getRange(1, notesIdx + 1, 1, 4).setValues(insertHeaders);
  styleHeader(sheet, sheet.getLastColumn());
  console.log('[Contacts] Schema upgraded — added Mailing Street/City/State/Zip columns');
}

// Resolve the Contacts data sheet robustly. A messy migration history left
// several contact-like tabs in CRM Master; if the canonical "Contacts" tab
// is missing or empty, fall back to whichever sheet actually holds the
// contact rows (identified by having both a Contact ID and an Email header).
function findBestContactsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var primary = ss.getSheetByName('Contacts');
  if (primary && primary.getLastRow() > 1) return primary;
  var sheets = ss.getSheets();
  var best = null, bestRows = 0;
  for (var s = 0; s < sheets.length; s++) {
    var sh = sheets[s];
    var lc = sh.getLastColumn(), lr = sh.getLastRow();
    if (lc < 1 || lr <= 1) continue;
    var hdr = sh.getRange(1, 1, 1, lc).getValues()[0].map(function(h) {
      return String(h).trim().toLowerCase();
    });
    // Must look like a real Contacts table — and NOT the 3-col
    // "Contact Accounts" join tab (which has no Email column).
    if (hdr.indexOf('contact id') === -1 || hdr.indexOf('email') === -1) continue;
    if (lr - 1 > bestRows) { best = sh; bestRows = lr - 1; }
  }
  if (best) return best;
  return primary || getContactsSheet();
}

function _generateContactId() {
  return 'C' + Date.now().toString(36).toUpperCase() + Math.floor(Math.random() * 1000).toString(36).toUpperCase();
}

function upsertContact(contactId, tradeName, firstName, lastName, title, email, phone, mobile, notes) {
  var sheet = getContactsSheet();
  var lastRow = sheet.getLastRow();
  var now = new Date();
  var nowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  if (contactId && lastRow >= 2) {
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(contactId)) {
        var rowNum = i + 2;
        var origCreated = sheet.getRange(rowNum, 10).getValue() || nowStr;
        sheet.getRange(rowNum, 1, 1, 11).setValues([[
          contactId, tradeName || '', firstName || '', lastName || '',
          title || '', email || '', phone || '', mobile || '',
          notes || '', origCreated, nowStr
        ]]);
        return contactId;
      }
    }
  }
  var newId = contactId || _generateContactId();
  sheet.appendRow([
    newId, tradeName || '', firstName || '', lastName || '',
    title || '', email || '', phone || '', mobile || '',
    notes || '', nowStr, nowStr
  ]);
  return newId;
}

function deleteContact(contactId) {
  var sheet = getContactsSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(contactId)) {
      sheet.deleteRow(i + 2);
      return true;
    }
  }
  return false;
}

// ── CONTACT ACCOUNTS (many-to-many join) ──────────────────────
// A contact's PRIMARY account is its Trade Name in the Contacts tab.
// This "Contact Accounts" tab holds ADDITIONAL accounts a contact also
// buys for, so one contact can be linked to multiple retailers.
function getContactAccountsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tabName = 'Contact Accounts';
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    var headers = ['Contact ID', 'Trade Name', 'Created At'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeader(sheet, headers.length);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 130);
    sheet.setColumnWidth(2, 240);
    sheet.setColumnWidth(3, 160);
  }
  return sheet;
}

// ── doGet — Activity Log + Tasks + Reservations (read) ────────
function doGet(e) {
  var params   = e.parameter || {};
  var action   = params.action   || 'read';
  var callback = params.callback || null;

  try {

    if (action === 'readTasks') {
      var sheet   = getTasksSheet();
      var lastRow = sheet.getLastRow();
      if (lastRow <= 1) return buildResponse({ ok: true, tasks: [] }, callback);
      var data  = sheet.getRange(2, 1, lastRow - 1, TASKS_HEADERS.length).getValues();
      var tasks = data
        .filter(function(row) { return row[0] || row[1]; })
        .map(function(row) {
          return {
            accountName   : row[0] || '',
            subject       : row[1] || '',
            relatedTo     : row[2] || '',
            completedDate : row[3] || '',
            date          : row[4] || ''
          };
        });
      return buildResponse({ ok: true, tasks: tasks }, callback);
    }

    if (action === 'readTaskCompletions') {
      var sheet = getTaskCompletionsSheet();
      var lastRow = sheet.getLastRow();
      if (lastRow <= 1) return buildResponse({ ok: true, completions: [] }, callback);
      var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
      var completions = [];
      data.forEach(function(row) {
        if (!row[0]) return;
        completions.push({
          taskKey:      row[0] || '',
          account:      row[1] || '',
          taskName:     row[2] || '',
          dueDate:      row[3] || '',
          completedAt:  row[4] || '',
          completedBy:  row[5] || '',
          notes:        row[6] || ''
        });
      });
      return buildResponse({ ok: true, completions: completions }, callback);
    }

    if (action === 'markTaskDone') {
      var taskKey  = (params.taskKey  || '').trim();
      var account  = (params.account  || '').trim();
      var taskName = (params.taskName || '').trim();
      var dueDate  = (params.dueDate  || '').trim();
      var user     = (params.user     || 'Unknown').trim();
      var notes    = (params.notes    || '').trim();
      if (!taskKey) return buildResponse({ ok: false, error: 'taskKey required' }, callback);
      var result = upsertTaskCompletion(taskKey, account, taskName, dueDate, user, notes);
      return buildResponse({ ok: true, action: result }, callback);
    }

    if (action === 'markTaskUndone') {
      var taskKey = (params.taskKey || '').trim();
      if (!taskKey) return buildResponse({ ok: false, error: 'taskKey required' }, callback);
      var deleted = deleteTaskCompletion(taskKey);
      return buildResponse({ ok: true, deleted: deleted }, callback);
    }

    if (action === 'readContacts') {
      var sheet = findBestContactsSheet();
      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      if (lastRow <= 1) {
        return buildResponse({ ok: true, contacts: [], count: 0, sheet: sheet.getName() }, callback);
      }
      // Read by header name so this works whether the Contacts tab has the
      // old 11-column layout or the upgraded 15-column layout (Mailing fields).
      var all = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      var hdr = all[0].map(function(h) { return String(h).trim().toLowerCase(); });
      function ci(name) { return hdr.indexOf(name.toLowerCase()); }
      function g(r, i) { return (i >= 0 && r[i] != null) ? r[i] : ''; }
      var iId = ci('Contact ID'), iTrade = ci('Trade Name'), iFirst = ci('First Name'),
          iLast = ci('Last Name'), iTitle = ci('Title'), iEmail = ci('Email'),
          iPhone = ci('Phone'), iMobile = ci('Mobile'), iNotes = ci('Notes'),
          iStreet = ci('Mailing Street'), iCity = ci('Mailing City'),
          iState = ci('Mailing State'), iZip = ci('Mailing Zip'),
          iCreated = ci('Created At'), iUpdated = ci('Updated At');
      var contacts = [];
      for (var ri = 1; ri < all.length; ri++) {
        var r = all[ri];
        // Keep a row if it has an ID, an email, or a name. A contact with a
        // blank Trade Name (not paired to an account) is still included.
        if (!g(r, iId) && !g(r, iEmail) && !g(r, iFirst) && !g(r, iLast)) continue;
        contacts.push({
          contactId:     g(r, iId),
          tradeName:     g(r, iTrade),
          firstName:     g(r, iFirst),
          lastName:      g(r, iLast),
          title:         g(r, iTitle),
          email:         g(r, iEmail),
          phone:         g(r, iPhone),
          mobile:        g(r, iMobile),
          mailingStreet: g(r, iStreet),
          mailingCity:   g(r, iCity),
          mailingState:  g(r, iState),
          mailingZip:    g(r, iZip),
          notes:         g(r, iNotes),
          createdAt:     g(r, iCreated),
          updatedAt:     g(r, iUpdated)
        });
      }
      return buildResponse({ ok: true, contacts: contacts, count: contacts.length, sheet: sheet.getName() }, callback);
    }

    if (action === 'upsertContact') {
      var contactId = (params.contactId || '').trim();
      var tradeName = (params.tradeName || '').trim();
      var firstName = (params.firstName || '').trim();
      var lastName  = (params.lastName  || '').trim();
      var title     = (params.title     || '').trim();
      var email     = (params.email     || '').trim();
      var phone     = (params.phone     || '').trim();
      var mobile    = (params.mobile    || '').trim();
      var notes     = (params.notes     || '').trim();
      if (!firstName && !lastName && !email) {
        return buildResponse({ ok: false, error: 'name or email required' }, callback);
      }
      var savedId = upsertContact(contactId, tradeName, firstName, lastName, title, email, phone, mobile, notes);
      return buildResponse({ ok: true, contactId: savedId }, callback);
    }

    if (action === 'deleteContact') {
      var contactId = (params.contactId || '').trim();
      if (!contactId) return buildResponse({ ok: false, error: 'contactId required' }, callback);
      var deleted = deleteContact(contactId);
      return buildResponse({ ok: true, deleted: deleted }, callback);
    }

    // ── CONTACT ACCOUNTS — read all contact↔account links ──
    if (action === 'readContactAccounts') {
      var caSheet = getContactAccountsSheet();
      var caLast = caSheet.getLastRow();
      if (caLast <= 1) return buildResponse({ ok: true, links: [] }, callback);
      var caRows = caSheet.getRange(2, 1, caLast - 1, 2).getValues();
      var links = [];
      caRows.forEach(function(r) {
        if (!r[0] || !r[1]) return;
        links.push({ contactId: String(r[0]).trim(), tradeName: String(r[1]).trim() });
      });
      return buildResponse({ ok: true, links: links }, callback);
    }

    // ── CONTACT ACCOUNTS — link a contact to an additional account ──
    if (action === 'addContactAccount') {
      var cacId = (params.contactId || '').trim();
      var cacTrade = (params.tradeName || '').trim();
      if (!cacId || !cacTrade) {
        return buildResponse({ ok: false, error: 'contactId and tradeName required' }, callback);
      }
      var caSheetA = getContactAccountsSheet();
      var caLastA = caSheetA.getLastRow();
      if (caLastA >= 2) {
        var existingA = caSheetA.getRange(2, 1, caLastA - 1, 2).getValues();
        for (var ai = 0; ai < existingA.length; ai++) {
          if (String(existingA[ai][0]).trim() === cacId &&
              String(existingA[ai][1]).trim().toLowerCase() === cacTrade.toLowerCase()) {
            return buildResponse({ ok: true, alreadyLinked: true }, callback);
          }
        }
      }
      var caNow = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      caSheetA.appendRow([cacId, cacTrade, caNow]);
      return buildResponse({ ok: true }, callback);
    }

    // ── CONTACT ACCOUNTS — remove a contact↔account link ──
    if (action === 'removeContactAccount') {
      var carId = (params.contactId || '').trim();
      var carTrade = (params.tradeName || '').trim();
      if (!carId || !carTrade) {
        return buildResponse({ ok: false, error: 'contactId and tradeName required' }, callback);
      }
      var caSheetR = getContactAccountsSheet();
      var caLastR = caSheetR.getLastRow();
      var caRemoved = false;
      if (caLastR >= 2) {
        var existingR = caSheetR.getRange(2, 1, caLastR - 1, 2).getValues();
        for (var ri = existingR.length - 1; ri >= 0; ri--) {
          if (String(existingR[ri][0]).trim() === carId &&
              String(existingR[ri][1]).trim().toLowerCase() === carTrade.toLowerCase()) {
            caSheetR.deleteRow(ri + 2);
            caRemoved = true;
          }
        }
      }
      return buildResponse({ ok: true, removed: caRemoved }, callback);
    }

    // ── PORTAL — full context for a contact: who they are + all stores ──
    // The portal calls this on load when given ?c=<contactId>.
    // Returns: contact { contactId, firstName, lastName, email },
    //          accounts [{ tradeName, isPrimary }, ...]   primary first.
    if (action === 'getContactPortalContext') {
      var cpcId = (params.contactId || params.c || '').trim();
      if (!cpcId) return buildResponse({ ok: false, error: 'contactId required' }, callback);

      // 1) Find the contact in the Contacts tab.
      var cpcSheet = findBestContactsSheet();
      var cpcLast = cpcSheet.getLastRow();
      var cpcCols = cpcSheet.getLastColumn();
      if (cpcLast <= 1) return buildResponse({ ok: false, error: 'contact not found' }, callback);
      var cpcHdr = cpcSheet.getRange(1, 1, 1, cpcCols).getValues()[0].map(function(h) {
        return String(h).trim().toLowerCase();
      });
      var idxId        = cpcHdr.indexOf('contact id');
      var idxTrade     = cpcHdr.indexOf('trade name');
      var idxFirst     = cpcHdr.indexOf('first name');
      var idxLast      = cpcHdr.indexOf('last name');
      var idxEmail     = cpcHdr.indexOf('email');
      if (idxId < 0) return buildResponse({ ok: false, error: 'Contacts sheet missing Contact ID column' }, callback);

      var cpcRows = cpcSheet.getRange(2, 1, cpcLast - 1, cpcCols).getValues();
      var cpcRow = null;
      for (var ci = 0; ci < cpcRows.length; ci++) {
        if (String(cpcRows[ci][idxId]).trim() === cpcId) { cpcRow = cpcRows[ci]; break; }
      }
      if (!cpcRow) return buildResponse({ ok: false, error: 'contact not found' }, callback);

      var primaryTrade = (idxTrade >= 0 ? String(cpcRow[idxTrade]).trim() : '');
      var contactOut = {
        contactId: cpcId,
        firstName: idxFirst >= 0 ? String(cpcRow[idxFirst]).trim() : '',
        lastName:  idxLast  >= 0 ? String(cpcRow[idxLast]).trim()  : '',
        email:     idxEmail >= 0 ? String(cpcRow[idxEmail]).trim() : ''
      };

      // 2) Collect additional accounts from the Contact Accounts join tab.
      var caSheet = getContactAccountsSheet();
      var caLast = caSheet.getLastRow();
      var additional = [];   // trade names
      if (caLast > 1) {
        var caRows = caSheet.getRange(2, 1, caLast - 1, 2).getValues();
        for (var ai = 0; ai < caRows.length; ai++) {
          if (String(caRows[ai][0]).trim() === cpcId) {
            var t = String(caRows[ai][1]).trim();
            if (t) additional.push(t);
          }
        }
      }

      // 3) Build accounts list: primary first, then additional alphabetically.
      //    Dedupe by lowercase trade name so an accidental duplicate row
      //    in the join tab doesn't show twice in the switcher.
      var seen = {};
      var accounts = [];
      if (primaryTrade) {
        accounts.push({ tradeName: primaryTrade, isPrimary: true });
        seen[primaryTrade.toLowerCase()] = true;
      }
      additional.sort(function(a, b) { return a.toLowerCase() < b.toLowerCase() ? -1 : 1; });
      additional.forEach(function(t) {
        var key = t.toLowerCase();
        if (seen[key]) return;
        accounts.push({ tradeName: t, isPrimary: false });
        seen[key] = true;
      });

      return buildResponse({ ok: true, contact: contactOut, accounts: accounts }, callback);
    }

    // ── ACCOUNT INFO — account-detail lookup for the CRM ──
    if (action === 'getAccountInfo') {
      var aiName = (params.account || params.tradeName || params.name || '').trim();
      if (!aiName) return buildResponse({ ok: false, error: 'account required' }, callback);
      var aiInfo = getAccountInfo(aiName);
      return buildResponse({ ok: true, account: aiInfo, found: !!aiInfo }, callback);
    }

    // ── RECONCILE ACCOUNTS — rebuild the unified CRM Accounts tab ──
    if (action === 'reconcileAccounts') {
      var rcT0 = Date.now();
      var rcResult = reconcileAccounts() || {};
      var rcElapsed = Date.now() - rcT0;
      return buildResponse({
        ok: true,
        total: rcResult.total || 0,
        mirror: rcResult.mirror || 0,
        cultiveraOnly: rcResult.cultiveraOnly || 0,
        elapsedMs: rcElapsed,
        result: rcResult
      }, callback);
    }

    // ── PORTAL — single-line reservation write (legacy GET path) ──
    if (action === 'reservation') {
      var reservationId = (params.reservationId || '').trim();
      var account       = (params.account || '').trim();
      var deliveryWeek  = (params.deliveryWeek || '').trim();
      var strain        = (params.strain || '').trim();
      var productSize   = (params.productSize || '').trim();
      var qty           = parseInt(params.qty) || 0;
      var unitPrice     = parseFloat(params.unitPrice) || 0;
      var lineTotal     = parseFloat(params.lineTotal) || 0;
      if (!account || !strain || !qty) {
        return buildResponse({ ok: false, error: 'account, strain, and qty are required' }, callback);
      }
      var sheet = getReservationsSheet();
      var now = new Date();
      var timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      sheet.appendRow([timestamp, reservationId, account, deliveryWeek, strain, productSize, qty, unitPrice, lineTotal, 'Pending', '', '']);
      return buildResponse({ ok: true, timestamp: timestamp }, callback);
    }

    // ── PORTAL — read back a reservation by ID (for "View or modify") ──
    if (action === 'getReservation') {
      var resId = (params.reservationId || '').trim();
      if (!resId) return buildResponse({ ok: false, error: 'reservationId required' }, callback);
      var sheet = getReservationsSheet();
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) return buildResponse({ ok: true, found: false, reservationId: resId }, callback);
      var data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
      var items = [], account = '', notes = '';
      for (var ri = 0; ri < data.length; ri++) {
        if (String(data[ri][1]).trim() !== resId) continue;
        if (String(data[ri][9]).toLowerCase() === 'cancelled') continue;
        if (data[ri][2]) account = data[ri][2];
        if (data[ri][11]) notes = String(data[ri][11]);
        items.push({
          deliveryWeek: String(data[ri][3] || ''),
          strain:       String(data[ri][4] || ''),
          productSize:  String(data[ri][5] || ''),
          qty:          Number(data[ri][6]) || 0,
          unitPrice:    Number(data[ri][7]) || 0,
          lineTotal:    Number(data[ri][8]) || 0
        });
      }
      return buildResponse({
        ok: true, found: items.length > 0, reservationId: resId,
        account: account, notes: notes, items: items
      }, callback);
    }

    if (action === 'log') {
      var account = (params.account || '').trim();
      var type    = (params.type    || '').trim();
      var notes   = (params.notes   || '').trim();
      var user    = (params.user    || 'Unknown').trim();
      if (!account || !notes) {
        return buildResponse({ ok: false, error: 'account and notes are required' }, callback);
      }
      var now       = new Date();
      var timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      var dateLabel = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a');
      getActivityLogSheet().appendRow([timestamp, account, type, notes, user, dateLabel]);
      return buildResponse({ ok: true, timestamp: timestamp }, callback);
    }

    if (action === 'logPurchaseOrder') {
      var poAccount    = (params.account    || '').trim();
      var poNumber     = (params.poNumber   || '').trim();
      var poRecipients = (params.recipients || '').trim();
      var poItemCount  = parseInt(params.itemCount)  || 0;
      var poUnitCount  = parseInt(params.unitCount)  || 0;
      var poTotal      = parseFloat(params.total)    || 0;
      var poItemsJson  = (params.itemsJson  || '').trim();
      var poUser       = (params.user       || 'Unknown').trim();
      if (!poAccount || !poItemsJson) {
        return buildResponse({ ok: false, error: 'account and itemsJson are required' }, callback);
      }
      var nowPO = new Date();
      var tsPO   = Utilities.formatDate(nowPO, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      var dlPO   = Utilities.formatDate(nowPO, Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a');
      _getOrCreatePOSheet().appendRow([
        tsPO, poNumber, poAccount, poRecipients,
        poItemCount, poUnitCount, poTotal, poItemsJson, poUser, dlPO
      ]);
      return buildResponse({ ok: true, timestamp: tsPO, poNumber: poNumber }, callback);
    }

    if (action === 'readPurchaseOrders') {
      var poSheet = _getOrCreatePOSheet();
      var poLastRow = poSheet.getLastRow();
      if (poLastRow <= 1) return buildResponse({ ok: true, orders: [] }, callback);
      var poData = poSheet.getRange(2, 1, poLastRow - 1, 10).getValues();
      var orders = [];
      for (var p = poData.length - 1; p >= 0 && orders.length < 500; p--) {
        var row = poData[p];
        if (!row[0]) continue;
        orders.push({
          timestamp: row[0], poNumber: row[1], account: row[2], recipients: row[3],
          itemCount: row[4], unitCount: row[5], total: row[6], itemsJson: row[7],
          user: row[8], dateLabel: row[9]
        });
      }
      return buildResponse({ ok: true, orders: orders }, callback);
    }

    // ── default: activity log read ──
    var account = (params.account || '').trim();
    var limit   = Math.min(parseInt(params.limit) || 100, 500);
    var sheet   = getActivityLogSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return buildResponse({ ok: true, entries: [] }, callback);
    }
    var data    = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    var entries = [];
    for (var i = data.length - 1; i >= 0; i--) {
      var row = data[i];
      if (account && row[1] !== account) continue;
      entries.push({
        timestamp: row[0], account: row[1], type: row[2],
        notes: row[3], user: row[4], date: row[5]
      });
      if (entries.length >= limit) break;
    }
    return buildResponse({ ok: true, entries: entries }, callback);

  } catch(err) {
    return buildResponse({ ok: false, error: err.toString() }, callback);
  }
}

// ── doPost — Orders + Tasks + Batches + Reservations (write) ───
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);

    if (payload.action === 'writeReservation') {
      var items = payload.items || [];
      var reservationId = payload.reservationId || '';
      var account = payload.account || '';
      var contactName = payload.contactName || '';
      var contactEmail = payload.contactEmail || '';
      var grandTotal = payload.grandTotal || 0;
      var sendConfirmation = payload.sendConfirmation || false;
      var notes = (payload.notes || '').toString().trim();
      var sheet = getReservationsSheet();
      var now = new Date();
      var timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

      var rows = items.map(function(item) {
        return [
          timestamp, reservationId, account,
          item.deliveryWeek || '', item.strain || '', item.productSize || '',
          item.qty || 0, item.unitPrice || 0, item.lineTotal || 0,
          'Pending', contactEmail || '', notes
        ];
      });

      if (rows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 12).setValues(rows);
      }

      if (account && items.length > 0) {
        var summary = items.map(function(i) { return i.strain + ' ' + i.productSize + ' x' + i.qty; }).join(', ');
        var total = items.reduce(function(s, i) { return s + (parseFloat(i.lineTotal) || 0); }, 0);
        var notesSnippet = notes ? ' | Notes: ' + (notes.length > 80 ? notes.substring(0, 80) + '...' : notes) : '';
        var logNotes = 'Portal reservation ' + reservationId + ': ' + summary + ' - Total: $' + total.toFixed(2) + notesSnippet;
        getActivityLogSheet().appendRow([timestamp, account, 'order', logNotes, 'Portal', Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a')]);
      }

      if (sendConfirmation) {
        if (contactEmail) {
          try {
            sendOrderConfirmationEmail(contactEmail, contactName, account, reservationId, items, grandTotal, notes);
            Logger.log('[OrderConfirm] Sent to: ' + contactEmail + ' for ' + reservationId);
          } catch(emailErr) {
            Logger.log('[OrderConfirm] ERROR: ' + emailErr.message);
          }
        } else {
          try {
            sendOrderConfirmationEmail('josh.ruhlman@redbird-cannabis.com', contactName || 'Unknown', account, reservationId, items, grandTotal, notes);
          } catch(emailErr) {
            Logger.log('[OrderConfirm] Fallback ERROR: ' + emailErr.message);
          }
        }
      }

      Logger.log('Portal reservation written: ' + reservationId + ', ' + rows.length + ' items' + (notes ? ', with notes' : ''));
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, count: rows.length, reservationId: reservationId }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (payload.action === 'updateReservation') {
      var reservationId = payload.reservationId || '';
      var updatedItems = payload.items || [];
      var account = payload.account || '';
      // Notes are reservation-level. Only touch column 12 when the payload
      // actually includes the field — guards against an older cached portal
      // build (which omits it) wiping an existing note on an unrelated edit.
      var notesProvided = (typeof payload.notes !== 'undefined' && payload.notes !== null);
      var notes = notesProvided ? String(payload.notes).trim() : '';
      var sheet = getReservationsSheet();
      var lastRow = sheet.getLastRow();
      var now = new Date();
      var timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

      if (lastRow > 1) {
        var data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
        var updatedMap = {};
        updatedItems.forEach(function(item) {
          var key = (item.strain || '') + '|' + (item.productSize || '');
          updatedMap[key] = item;
        });

        for (var i = 0; i < data.length; i++) {
          if (data[i][1] !== reservationId) continue;
          var rowNum = i + 2;
          var key = data[i][4] + '|' + data[i][5];

          if (updatedMap[key]) {
            var item = updatedMap[key];
            sheet.getRange(rowNum, 7).setValue(item.qty);
            sheet.getRange(rowNum, 9).setValue(item.lineTotal);
            sheet.getRange(rowNum, 1).setValue(timestamp);
            delete updatedMap[key];
          } else {
            sheet.getRange(rowNum, 10).setValue('Cancelled');
            sheet.getRange(rowNum, 11).setValue(timestamp);
          }
          if (notesProvided) sheet.getRange(rowNum, 12).setValue(notes);
        }

        var newKeys = Object.keys(updatedMap);
        if (newKeys.length > 0) {
          var newRows = newKeys.map(function(key) {
            var item = updatedMap[key];
            return [timestamp, reservationId, account, item.deliveryWeek || '', item.strain || '', item.productSize || '', item.qty || 0, item.unitPrice || 0, item.lineTotal || 0, 'Pending', '', notes];
          });
          sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 12).setValues(newRows);
        }
      }

      var editSummary = updatedItems.map(function(i) { return i.strain + ' ' + i.productSize + ' x' + i.qty; }).join(', ');
      var editTotal = updatedItems.reduce(function(s, i) { return s + (parseFloat(i.lineTotal) || 0); }, 0);
      var editNotesSnippet = notes ? ' | Notes: ' + (notes.length > 80 ? notes.substring(0, 80) + '...' : notes) : '';
      getActivityLogSheet().appendRow([timestamp, account, 'order-edit', 'Edited reservation ' + reservationId + ': ' + editSummary + ' - New total: $' + editTotal.toFixed(2) + editNotesSnippet, 'Portal', Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a')]);

      // Send an "updated" confirmation email, mirroring writeReservation.
      // The portal sends contactName / contactEmail / grandTotal / sendConfirmation
      // on the updateReservation POST too, so reuse them here.
      var sendConfirmation = payload.sendConfirmation || false;
      var contactName = payload.contactName || '';
      var contactEmail = payload.contactEmail || '';
      var grandTotal = payload.grandTotal || editTotal;
      if (sendConfirmation) {
        var updTo = contactEmail || 'josh.ruhlman@redbird-cannabis.com';
        try {
          sendOrderConfirmationEmail(updTo, contactName, account, reservationId, updatedItems, grandTotal, notes, true);
          Logger.log('[UpdateEmail] Sent to ' + updTo + ' for ' + reservationId);
        } catch(emailErr) {
          Logger.log('[UpdateEmail] ERROR: ' + emailErr.message);
        }
      }

      Logger.log('Reservation updated: ' + reservationId + (notesProvided ? ', notes synced' : ''));
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, reservationId: reservationId }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (payload.action === 'cancelReservation') {
      var reservationId = payload.reservationId || '';
      var account       = payload.account || '';
      var contactName   = payload.contactName || '';
      var contactEmail  = payload.contactEmail || '';
      var sheet = getReservationsSheet();
      var lastRow = sheet.getLastRow();
      var now = new Date();
      var timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      var cancelled = 0;
      var cancelledItems = [];

      if (lastRow > 1) {
        var data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
        for (var i = 0; i < data.length; i++) {
          if (data[i][1] !== reservationId) continue;
          if (data[i][9] === 'Cancelled') continue;
          var rowNum = i + 2;
          // Capture items + an email fallback from the in-memory snapshot
          // BEFORE overwriting column 11 (writeReservation stashes email there).
          cancelledItems.push({ strain: data[i][4], productSize: data[i][5], qty: data[i][6] });
          if (!contactEmail && data[i][10] && String(data[i][10]).indexOf('@') > -1) {
            contactEmail = String(data[i][10]);
          }
          if (!account && data[i][2]) account = data[i][2];
          sheet.getRange(rowNum, 10).setValue('Cancelled');
          sheet.getRange(rowNum, 11).setValue(timestamp);
          cancelled++;
        }
      }

      getActivityLogSheet().appendRow([timestamp, account, 'order-cancel', 'Cancelled reservation ' + reservationId + ' (' + cancelled + ' items)', 'Portal', Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a')]);

      if (cancelled > 0) {
        var cancelTo = contactEmail || 'josh.ruhlman@redbird-cannabis.com';
        try {
          sendCancellationEmail(cancelTo, contactName, account, reservationId, cancelledItems);
          Logger.log('[CancelEmail] Sent to ' + cancelTo + ' for ' + reservationId);
        } catch(emailErr) {
          Logger.log('[CancelEmail] ERROR: ' + emailErr.message);
        }
      }

      Logger.log('Reservation cancelled: ' + reservationId + ', ' + cancelled + ' items');
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, reservationId: reservationId, cancelled: cancelled }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (payload.action === 'mirrorBatchesInStock') {
      var result = mirrorBatchesInStock(payload.data || []);
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }

    if (payload.action === 'writeTasks') {
      var result = writeTasks(payload.tasks || []);
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }

    var orders = payload.orders || [];
    var result = logOrders(orders);
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── WRITE TASKS ───────────────────────────────────────────────
function writeTasks(tasks) {
  try {
    var sheet = getTasksSheet();
    tasks.forEach(function(t) {
      sheet.appendRow([
        t.accountName || '', t.subject || '', t.relatedTo || '',
        t.completedDate || '', t.date || ''
      ]);
    });
    Logger.log('Tasks written: ' + tasks.length + ' rows');
    return { success: true, count: tasks.length };
  } catch(err) {
    Logger.log('Error writing tasks: ' + err.message);
    return { success: false, error: err.message };
  }
}

// ── LOG ORDERS ────────────────────────────────────────────────
function logOrders(orders) {
  try {
    var sheet     = getOrdersSheet();
    var timestamp = new Date();
    orders.forEach(function(order) {
      sheet.appendRow([
        timestamp, order.accountName || '', order.opportunityName || '',
        order.closeDate || '', order.stage || '', order.productName || '',
        order.salesPrice || '', order.quantity || ''
      ]);
    });
    Logger.log('Orders logged: ' + orders.length + ' rows');
    return { success: true, count: orders.length };
  } catch(err) {
    Logger.log('Error logging orders: ' + err.message);
    return { success: false, error: err.message };
  }
}

// ── MIRROR: BATCHES IN STOCK ──────────────────────────────────
function mirrorBatchesInStock(data) {
  try {
    var dest = SpreadsheetApp.openById(INVENTORY_SOURCE_ID);
    var tabName = 'batches in stock';
    var sheet = dest.getSheetByName(tabName);
    if (!sheet) sheet = dest.insertSheet(tabName);

    sheet.clearContents();
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a');

    if (data && data.length > 0) {
      var headers = [
        'Last Updated', 'Barcode', 'Product', 'Product Line', 'Sub-Product Line',
        'Category', 'Sub-Category', 'Batch Date',
        'THCA', 'THC', 'CBD', 'QA Total',
        'Availability', 'Units For Sale', 'Units On Hold', 'Units Allocated', 'Units In Stock'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      styleHeader(sheet, headers.length);
      sheet.setFrozenRows(1);
      var rows = data.map(function(r) {
        return [
          timestamp, r.barcode || '', r.product || '', r.productLine || '',
          r.subProductLine || '', r.category || '', r.subCategory || '',
          r.batchDate || '', r.thca || '', r.thc || '', r.cbd || '',
          r.qaTotal || '', r.availability || '',
          r.unitsForSale || 0, r.unitsOnHold || 0,
          r.unitsAllocated || 0, r.unitsInStock || 0
        ];
      });
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
      Logger.log('Batches in stock mirrored from POST: ' + rows.length + ' rows');
      return { success: true, count: rows.length, timestamp: timestamp };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var invSheet = ss.getSheetByName('Inventory');
    if (!invSheet) {
      Logger.log('Inventory tab not found in CRM sheet');
      return { success: false, error: 'Inventory tab not found' };
    }
    var invData = invSheet.getDataRange().getValues();
    if (invData.length < 2) return { success: true, count: 0 };

    var headers = ['Last Updated'].concat(invData[0]);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeader(sheet, headers.length);
    sheet.setFrozenRows(1);
    var rows = invData.slice(1).map(function(row) { return [timestamp].concat(row); });
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
    Logger.log('Batches in stock mirrored from Inventory tab: ' + rows.length + ' rows');
    return { success: true, count: rows.length, timestamp: timestamp };
  } catch(e) {
    Logger.log('Error mirroring batches in stock: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ── MIRROR: ACCOUNTS ──────────────────────────────────────────
function mirrorAccounts() {
  try {
    var source      = SpreadsheetApp.openById(ACCOUNTS_SOURCE_ID);
    var sourceSheet = source.getSheetByName(ACCOUNTS_SOURCE_TAB);
    if (!sourceSheet) { Logger.log('Accounts source tab not found.'); return; }
    var data = sourceSheet.getDataRange().getValues();
    if (!data.length) return;
    var destSheet = getActiveSheet(ACCOUNTS_DEST_TAB);
    destSheet.clearContents();
    destSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    styleHeader(destSheet, data[0].length);
    destSheet.setFrozenRows(1);
    Logger.log('Accounts mirror updated: ' + data.length + ' rows');
  } catch(e) {
    Logger.log('Error syncing Accounts: ' + e.message);
  }
}

// ── MIRROR: INVENTORY PROJECTIONS ─────────────────────────────
/**
 * ============================================================================
 *  REDBIRD INVENTORY PIPELINE
 *  Product Allocation + Propagation  →  "Inventory Projections Mirror"
 *  (feeds the CRM Projected tab, Share Projected emails, and the retailer portal)
 * ============================================================================
 *
 *  WHAT IT DOES (every hour via syncAll(), automatically)
 *    1. Reads Product Allocation (Chris's sheet) and Propagation (Cultivation
 *       Overview) BY HEADER NAME. Columns can be moved or inserted in either
 *       sheet without breaking anything. If a required column disappears, the
 *       run stops, the last good data stays live, and you get an email.
 *    2. Merges them:
 *         - Product Allocation is authoritative for any harvest week it has
 *           (it carries Chris's 7g / 28g / wholesale / fresh-frozen decisions).
 *         - Every later week comes from Propagation, with units projected using
 *           the same math as Product Allocation and per-tote yields taken from
 *           the most recent Product Allocation harvests (self-calibrating).
 *         - Tote count uses the best number available:
 *           FM2 Actual → FM1 Actual → Flower Ready → Target.
 *    3. Validates before writing. It never replaces good data with an empty or
 *       broken set.
 *    4. Writes the mirror in ONE fixed layout (MIRROR_SCHEMA below). That layout
 *       is the contract the CRM and portal read. It does not change when the
 *       source sheets change.
 *    5. Records every run on a "Pipeline Status" tab (the CRM reads it for its
 *       freshness indicator) and a "Pipeline Log" tab (run history).
 *    6. Emails ALERT_EMAILS on failures or warnings (same issue repeats at most
 *       every 12 hours), and once more when things recover.
 *    7. A watchdog runs every 6 hours. If the hourly sync hasn't succeeded in
 *       3+ hours (trigger deleted, auth expired, quota), it emails you.
 *
 *  INSTALL (one time)
 *    1. Run  installInventoryPipeline()  once and approve the permissions.
 *       It confirms the hourly syncAll trigger, adds the watchdog, and runs
 *       a first sync.
 *    2. Check the "Pipeline Status" tab in the CRM Master sheet. It should say OK.
 *
 *  Everything in this file is self-contained. It does not depend on any other
 *  helper function in the project.
 * ============================================================================
 */

var INV_PIPELINE = {
  CRM_SHEET_ID: '1Nrz05Rr8B1-5-2HXIXByy_pL6TH9U6wz-HUgu9LSkjw',      // CRM Master
  MIRROR_TAB: 'Inventory Projections Mirror',
  STATUS_TAB: 'Pipeline Status',
  LOG_TAB: 'Pipeline Log',
  LOG_KEEP_ROWS: 1000,

  PA_SHEET_ID: '1IuSfAblyYWSO5gYZhdgL_KdaWSfxG7sjvBWv2zWURJ4',       // Product Allocation workbook
  PA_TAB: 'Product Allocation',

  PROP_SHEET_ID: '1UVNIp5DWbTabg21fGdABFP6KYAw4qNgU3M852LE6JhY',     // Cultivation Overview
  PROP_TAB: 'Propagation',

  ALERT_EMAILS: ['Josh.Ruhlman@redbird-cannabis.com'],
  ALERT_REPEAT_HOURS: 12,

  FLOWER_DAYS_AFTER_WEEK: 14,   // Week of → flower delivery (matches Product Allocation)
  JOINT_DAYS_AFTER_WEEK: 21,    // Week of → joint delivery
  MIN_RUNWAY_WEEKS: 6,          // warn if projected deliveries end sooner than this
  STALE_HOURS: 3,               // watchdog alerts if no successful sync in this long

  // Fallbacks used only if Product Allocation history can't provide a rate
  DEFAULT_P35_PER_TOTE: 70,
  DEFAULT_MICROS_PER_TOTE: 12,
  DEFAULT_GRAMS_PER_TOTE: 850,
  DEFAULT_TRIM: 0.35,
  RATE_LOOKBACK_WEEKS: 8
};

/**
 * The mirror layout. Column positions are the contract read by
 * sheets-loader.js (CRM) and index.html (portal). Don't reorder.
 * Add new fields only at the end.
 */
var MIRROR_SCHEMA = [
  'Week Of',            // 0
  'Grow Room',          // 1
  'Totes',              // 2
  'Plant Count',        // 3
  'Total Days',         // 4
  'Strain',             // 5
  'Flower Delivery',    // 6
  'Joint Delivery',     // 7
  'Proj 3.5g',          // 8   ← CRM p35 / portal p35
  'Grams Per Tote',     // 9
  'Actual 3.5g',        // 10  ← CRM a35
  '',                   // 11
  'Target 7g',          // 12
  'Target 28g',         // 13
  'Wholesale Tier 1',   // 14
  'Proj Micros 14g',    // 15  ← pm
  'Actual Micros 14g',  // 16
  'Wholesale Micros',   // 17
  'Post Grind/Sift g',  // 18
  'Trim %',             // 19
  '1g Joints',          // 20  ← j1
  '2pk Joints',         // 21  ← j2
  '10pk Joints',        // 22  ← j10
  'Infused',            // 23
  'Batch Count',        // 24
  'Total Proj 1g',      // 25
  'Total Proj 0.5g',    // 26
  'Fresh Frozen',       // 27  ← live resin flag
  'Cured',              // 28
  'Synced At',          // 29
  'Source'              // 30
];
var MC = {}; MIRROR_SCHEMA.forEach(function (h, i) { if (h) MC[h] = i; });

/**
 * Source headers, matched case- and punctuation-insensitively, with aliases
 * so routine renames ("Proj. 3.5g" → "Projected 3.5g") don't break the sync.
 * required: true → the run stops and alerts if the column can't be found.
 */
var PA_FIELDS = {
  week:      { names: ['Week of Date', 'Week of', 'Week'], required: true },
  room:      { names: ['Grow Room', 'Room'], required: true },
  totes:     { names: ['Totes'], required: true },
  plants:    { names: ['Plant Count', 'Plants'] },
  totalDays: { names: ['Total Days'] },
  strain:    { names: ['Strain', 'Strain Name'], required: true },
  fd:        { names: ['Prj. Flower Delivery Date', 'Proj Flower Delivery Date', 'Flower Delivery Date', 'Flower Delivery'] },
  jd:        { names: ['Prj. Joint Delivery Date', 'Proj Joint Delivery Date', 'Joint Delivery Date', 'Joint Delivery'] },
  perTote:   { names: ['Per Tote', 'Grams Per Tote'] },
  p35:       { names: ['Proj. 3.5g', 'Projected 3.5g', 'Proj 3.5g'], required: true },
  a35:       { names: ['3.5-g', 'Actual 3.5g', '3.5g'] },
  t7:        { names: ['Target 7g'] },
  t28:       { names: ['Target 28g'] },
  wt1:       { names: ['Wholesale Tier 1 Flower', 'Wholesale Tier 1'] },
  micros:    { names: ['Proj. Micros 14g', 'Projected Micros 14g', 'Proj Micros 14g'], required: true },
  aMicros:   { names: ['Micros 14g', 'Actual Micros 14g'] },
  wMicros:   { names: ['Wholesale Micros'] },
  grind:     { names: ['Post Grind/Sift', 'Post Grind Sift'] },
  trim:      { names: ['% Trim', 'Trim %', 'Trim'] },
  j1:        { names: ['1g', '1g Joints', '1g Singles'] },
  j2:        { names: ['2pk', '2pk Joints'] },
  j10:       { names: ['10pks', '10pk', '10pk Joints'], required: true },
  infused:   { names: ['Infused'] },
  batch:     { names: ['Batch Count'] },
  total1g:   { names: ['Total Proj. 1g', 'Total Proj 1g'] },
  totalHalf: { names: ['Total Proj. 0.5g', 'Total Proj 0.5g'] },
  ff:        { names: ['Fresh Frozen', 'Live Resin'] },
  cured:     { names: ['Cured', 'Cured Resin'] }
};

var PROP_FIELDS = {
  strain:       { names: ['Strain Name', 'Strain'], required: true },
  room:         { names: ['Grow Room', 'Room'], required: true },
  weekOf:       { names: ['Week of Date', 'Week of'], required: true },
  totalDays:    { names: ['Total Days'] },
  tgtTotes:     { names: ['Tgt Totes', 'Target Totes'], required: true },
  flowerTotes:  { names: ['Flower Ready Totes'] },
  flowerPlants: { names: ['Flower Ready Plants'] },
  fm1Totes:     { names: ['FM1 Actual Totes'] },
  fm2Totes:     { names: ['FM2 Actual Totes'] },
  kill:         { names: ['Kill'] },
  killDone:     { names: ['Kill Completed Date'] }
};

// ============================================================================
//  ENTRY POINTS
// ============================================================================

/**
 * One-time setup. Safe to re-run.
 *  - Makes sure the hourly syncAll() trigger exists (syncAll calls mirrorInventory)
 *  - Removes any stray standalone mirrorInventory triggers (avoids double runs)
 *  - Installs the 6-hour watchdog
 *  - Runs a first sync and reports the result
 */
function installInventoryPipeline() {
  var hasSyncAll = false;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === 'mirrorInventory' || fn === 'inventoryPipelineWatchdog') ScriptApp.deleteTrigger(t);
    else if (fn === 'syncAll') {
      if (hasSyncAll) ScriptApp.deleteTrigger(t);   // keep exactly one
      hasSyncAll = true;
    }
  });
  if (!hasSyncAll) ScriptApp.newTrigger('syncAll').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('inventoryPipelineWatchdog').timeBased().everyHours(6).create();
  var result = mirrorInventory();
  Logger.log('Hourly syncAll trigger ' + (hasSyncAll ? 'already present' : 'created') +
             '; 6-hour watchdog installed. First run: ' + JSON.stringify(result));
}

/** Hourly sync. Called by syncAll() every hour. Same name as before, so anything already calling it keeps working. */
function mirrorInventory() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(60 * 1000)) {
    Logger.log('Another inventory sync is already running. Skipping.');
    return { status: 'SKIPPED' };
  }
  var run = { started: new Date(), status: 'OK', warnings: [], errors: [], stats: {} };
  try {
    _invRun(run);
  } catch (e) {
    run.status = 'ERROR';
    run.errors.push('Unexpected error: ' + e.message);
    Logger.log(e.stack);
  } finally {
    try { _invRecord(run); } catch (e2) { Logger.log('Could not write status: ' + e2.message); }
    try { _invAlert(run); } catch (e3) { Logger.log('Could not send alert: ' + e3.message); }
    lock.releaseLock();
  }
  return { status: run.status, warnings: run.warnings, errors: run.errors, stats: run.stats };
}

/** Every 6 hours: alert if the hourly sync has stopped succeeding. */
function inventoryPipelineWatchdog() {
  var props = PropertiesService.getScriptProperties();
  var last = Number(props.getProperty('INV_LAST_SUCCESS') || 0);
  var hours = last ? (Date.now() - last) / 36e5 : Infinity;
  var hasHourly = ScriptApp.getProjectTriggers().some(function (t) {
    var fn = t.getHandlerFunction();
    return fn === 'syncAll' || fn === 'mirrorInventory';
  });
  var problems = [];
  if (!hasHourly) problems.push('The hourly syncAll trigger is missing. Run installInventoryPipeline() to restore it.');
  if (hours > INV_PIPELINE.STALE_HOURS) {
    problems.push('No successful inventory sync in ' + (isFinite(hours) ? hours.toFixed(1) + ' hours' : 'ever') +
                  '. The CRM Projected tab and the retailer portal are showing stale projections.');
  }
  if (problems.length) {
    _invSendAlert('WATCHDOG', 'Inventory sync has stopped', problems, []);
  }
}

// ============================================================================
//  CORE
// ============================================================================

function _invRun(run) {
  var tz = Session.getScriptTimeZone();
  var today = _invDay(new Date());
  var runwayDays = INV_PIPELINE.MIN_RUNWAY_WEEKS * 7;

  // ---- Read sources ----
  var pa = _invReadTable(INV_PIPELINE.PA_SHEET_ID, INV_PIPELINE.PA_TAB, PA_FIELDS, 'Product Allocation');
  var prop = _invReadTable(INV_PIPELINE.PROP_SHEET_ID, INV_PIPELINE.PROP_TAB, PROP_FIELDS, 'Propagation');
  [pa, prop].forEach(function (t) {
    if (t.error) run.errors.push(t.error);
    t.missingOptional.forEach(function (m) { run.warnings.push(t.label + ': optional column "' + m + '" not found. Using a default.'); });
  });
  if (run.errors.length) { run.status = 'ERROR'; return; }

  // ---- Product Allocation rows ----
  var paRows = [];
  pa.rows.forEach(function (r) {
    var w = _invDate(pa.get(r, 'week'));
    var strain = String(pa.get(r, 'strain') || '').trim();
    if (!w || !strain || /^no strains/i.test(strain)) return;
    var rec = {
      w: w, strain: strain,
      room: pa.get(r, 'room'), totes: _invNum(pa.get(r, 'totes')),
      plants: pa.get(r, 'plants'), totalDays: pa.get(r, 'totalDays'),
      fd: _invDate(pa.get(r, 'fd')) || _invAdd(w, INV_PIPELINE.FLOWER_DAYS_AFTER_WEEK),
      jd: _invDate(pa.get(r, 'jd')) || _invAdd(w, INV_PIPELINE.JOINT_DAYS_AFTER_WEEK),
      perTote: _invNum(pa.get(r, 'perTote')),
      p35: _invNum(pa.get(r, 'p35')), a35: _invNum(pa.get(r, 'a35')),
      t7: _invNum(pa.get(r, 't7')), t28: _invNum(pa.get(r, 't28')), wt1: _invNum(pa.get(r, 'wt1')),
      micros: _invNum(pa.get(r, 'micros')), aMicros: _invNum(pa.get(r, 'aMicros')), wMicros: _invNum(pa.get(r, 'wMicros')),
      grind: _invNum(pa.get(r, 'grind')), trim: _invPct(pa.get(r, 'trim')),
      j1: _invNum(pa.get(r, 'j1')), j2: _invNum(pa.get(r, 'j2')), j10raw: pa.get(r, 'j10'),
      infused: _invBool(pa.get(r, 'infused')), batch: _invNum(pa.get(r, 'batch')),
      total1g: _invNum(pa.get(r, 'total1g')), totalHalf: _invNum(pa.get(r, 'totalHalf')),
      ff: _invBool(pa.get(r, 'ff')), cured: _invBool(pa.get(r, 'cured')),
      source: 'Product Allocation'
    };
    rec.j10 = _invNum(rec.j10raw);
    paRows.push(rec);
  });

  // Rates calibrated from recent Product Allocation harvests
  var rates = _invRates(paRows);
  run.stats.rates = rates.summary;

  var paWeeks = {};
  paRows.forEach(function (r) { paWeeks[_invKey(r.w, tz)] = true; });

  var keep = [];
  paRows.forEach(function (r) {
    var latest = Math.max(r.w.getTime(), r.fd.getTime(), r.jd.getTime());
    if (latest < today.getTime()) return;
    if (/over\s*alloc/i.test(String(r.j10raw))) {
      run.warnings.push('Product Allocation: ' + r.strain + ' (week of ' + _invFmt(r.w, tz) + ') joints are OVER ALLOCATED. 10pk shown as 0.');
    }
    if (r.totes > 0 && r.p35 <= 0 && r.micros <= 0 && !r.ff) {
      run.warnings.push('Product Allocation: ' + r.strain + ' (week of ' + _invFmt(r.w, tz) + ') has ' + r.totes + ' totes but 0 projected 3.5g and Micros.');
    }
    keep.push(r);
  });
  run.stats.paRows = keep.length;

  // ---- Propagation rows for weeks Product Allocation doesn't cover ----
  var merged = {}, propStrainsByWeek = {};
  prop.rows.forEach(function (r) {
    var w = _invDate(prop.get(r, 'weekOf'));
    var strain = String(prop.get(r, 'strain') || '').trim();
    var room = String(prop.get(r, 'room') || '').trim();
    if (!w || !strain) return;
    if (_invBool(prop.get(r, 'kill')) || String(prop.get(r, 'killDone') || '').trim()) return;
    if (!room || /^n\/?a$/i.test(room)) return;
    var jd = _invAdd(w, INV_PIPELINE.JOINT_DAYS_AFTER_WEEK);
    if (jd < today) return;

    var totes = _invNum(prop.get(r, 'fm2Totes')) || _invNum(prop.get(r, 'fm1Totes')) ||
                _invNum(prop.get(r, 'flowerTotes')) || _invNum(prop.get(r, 'tgtTotes'));
    if (totes <= 0) return;

    var wk = _invKey(w, tz);
    (propStrainsByWeek[wk] = propStrainsByWeek[wk] || {})[strain.toLowerCase()] = strain;
    if (paWeeks[wk]) return;   // Product Allocation owns this week

    var k = wk + '|' + strain.toLowerCase();
    if (!merged[k]) merged[k] = { w: w, strain: strain, room: room, totes: 0, plants: 0, totalDays: prop.get(r, 'totalDays') };
    merged[k].totes += totes;
    merged[k].plants += _invNum(prop.get(r, 'flowerPlants'));
  });

  // Flag weeks where the two sheets disagree on strains (PA wins, but tell someone)
  var paStrainsByWeek = {};
  keep.forEach(function (r) { (paStrainsByWeek[_invKey(r.w, tz)] = paStrainsByWeek[_invKey(r.w, tz)] || {})[r.strain.toLowerCase()] = r.strain; });
  Object.keys(paStrainsByWeek).forEach(function (wk) {
    var ps = propStrainsByWeek[wk];
    if (!ps) return;
    var onlyProp = Object.keys(ps).filter(function (s) { return !paStrainsByWeek[wk][s]; }).map(function (s) { return ps[s]; });
    if (onlyProp.length) {
      run.warnings.push('Week of ' + wk + ': in Propagation but not Product Allocation → ' + onlyProp.join(', ') + '. Product Allocation is used for this week.');
    }
  });

  var propCount = 0;
  Object.keys(merged).forEach(function (k) {
    keep.push(_invProject(merged[k], rates));
    propCount++;
  });
  run.stats.propagationRows = propCount;

  // ---- Validate ----
  keep.sort(function (a, b) { return a.w - b.w || String(a.strain).localeCompare(String(b.strain)); });
  var futureJoint = keep.filter(function (r) { return r.jd >= today; });
  var weeks = {};
  futureJoint.forEach(function (r) { weeks[_invKey(r.w, tz)] = true; });
  run.stats.harvestWeeks = Object.keys(weeks).length;
  run.stats.totalRows = keep.length;
  var furthest = futureJoint.length ? new Date(Math.max.apply(null, futureJoint.map(function (r) { return r.jd.getTime(); }))) : null;
  run.stats.furthestDelivery = furthest ? _invFmt(furthest, tz) : '';

  if (!keep.length) {
    run.status = 'ERROR';
    run.errors.push('Both sources produced zero upcoming harvests. The mirror was NOT overwritten; the CRM and portal still show the last good data.');
    return;
  }
  if (!furthest || (furthest - today) / 864e5 < runwayDays) {
    run.warnings.push('Projected deliveries end ' + (furthest ? _invFmt(furthest, tz) : 'now') +
      ', less than ' + INV_PIPELINE.MIN_RUNWAY_WEEKS + ' weeks out. Add upcoming harvests to Propagation.');
  }

  // ---- Write (overwrite in place; never leaves the tab empty) ----
  var stamp = new Date();
  var values = [MIRROR_SCHEMA.slice()].concat(keep.map(function (r) { return _invToMirrorRow(r, stamp); }));
  var ss = SpreadsheetApp.openById(INV_PIPELINE.CRM_SHEET_ID);
  var sh = ss.getSheetByName(INV_PIPELINE.MIRROR_TAB) || ss.insertSheet(INV_PIPELINE.MIRROR_TAB);
  if (sh.getMaxColumns() < MIRROR_SCHEMA.length) sh.insertColumnsAfter(sh.getMaxColumns(), MIRROR_SCHEMA.length - sh.getMaxColumns());
  if (sh.getMaxRows() < values.length) sh.insertRowsAfter(sh.getMaxRows(), values.length - sh.getMaxRows());
  sh.getRange(1, 1, values.length, MIRROR_SCHEMA.length).setValues(values);
  var extra = sh.getLastRow() - values.length;
  if (extra > 0) sh.getRange(values.length + 1, 1, extra, sh.getMaxColumns()).clearContent();
  if (values.length > 1) {
    [MC['Week Of'], MC['Flower Delivery'], MC['Joint Delivery']].forEach(function (c) {
      sh.getRange(2, c + 1, values.length - 1, 1).setNumberFormat('M/d/yyyy');
    });
    sh.getRange(2, MC['Synced At'] + 1, values.length - 1, 1).setNumberFormat('M/d/yyyy h:mm am/pm');
  }
  sh.getRange(1, 1, 1, MIRROR_SCHEMA.length).setFontWeight('bold').setBackground('#f3f3f3');
  sh.setFrozenRows(1);
  SpreadsheetApp.flush();

  if (run.warnings.length && run.status === 'OK') run.status = 'WARN';
}

/** Project units for a Propagation harvest using Product Allocation's math. */
function _invProject(h, rates) {
  var trim = rates.trimByStrain[h.strain.toLowerCase()] || rates.trim;
  var grind = rates.perTote * h.totes * trim;
  return {
    w: h.w, strain: h.strain, room: h.room, totes: h.totes, plants: h.plants || '', totalDays: h.totalDays || '',
    fd: _invAdd(h.w, INV_PIPELINE.FLOWER_DAYS_AFTER_WEEK),
    jd: _invAdd(h.w, INV_PIPELINE.JOINT_DAYS_AFTER_WEEK),
    perTote: rates.perTote,
    p35: Math.round(h.totes * rates.p35), a35: 0, t7: 0, t28: 0, wt1: 0,
    micros: Math.round(h.totes * rates.micros), aMicros: 0, wMicros: 0,
    grind: Math.round(grind), trim: trim,
    j1: 0, j2: 0, j10: Math.round(grind * 0.75 / 5 / 10) * 10,
    infused: false, batch: 0, total1g: 0, totalHalf: Math.floor(grind * 0.75 / 5) * 10,
    ff: false, cured: false, source: 'Propagation'
  };
}

/** Per-tote yields from the last N weeks of Product Allocation. */
function _invRates(paRows) {
  var C = INV_PIPELINE;
  var rows = paRows.filter(function (r) { return r.totes > 0; }).sort(function (a, b) { return a.w - b.w; });
  var weekKeys = [];
  rows.forEach(function (r) { var k = r.w.getTime(); if (weekKeys.indexOf(k) < 0) weekKeys.push(k); });
  var recentSet = weekKeys.slice(-C.RATE_LOOKBACK_WEEKS);
  var recent = rows.filter(function (r) { return recentSet.indexOf(r.w.getTime()) >= 0 && !r.ff; });

  var p35 = [], micros = [];
  recent.forEach(function (r) {
    if (!r.t7 && !r.t28 && !r.wt1 && r.p35 > 0) p35.push(r.p35 / r.totes);
    if (!r.wMicros && r.micros > 0) micros.push(r.micros / r.totes);
  });
  var perTote = 0, trim = 0, trimByStrain = {};
  rows.forEach(function (r) {
    if (r.perTote > 0) perTote = r.perTote;
    if (r.trim > 0) { trim = r.trim; trimByStrain[r.strain.toLowerCase()] = r.trim; }
  });
  var out = {
    p35: _invMedian(p35) || C.DEFAULT_P35_PER_TOTE,
    micros: _invMedian(micros) || C.DEFAULT_MICROS_PER_TOTE,
    perTote: perTote || C.DEFAULT_GRAMS_PER_TOTE,
    trim: trim || C.DEFAULT_TRIM,
    trimByStrain: trimByStrain
  };
  out.summary = '3.5g/tote ' + out.p35 + ' · Micros/tote ' + out.micros + ' · g/tote ' + out.perTote +
                ' · trim ' + Math.round(out.trim * 100) + '% (from ' + recentSet.length + ' recent weeks)';
  return out;
}

function _invToMirrorRow(r, stamp) {
  var row = MIRROR_SCHEMA.map(function () { return ''; });
  row[MC['Week Of']] = r.w;
  row[MC['Grow Room']] = r.room;
  row[MC['Totes']] = r.totes;
  row[MC['Plant Count']] = r.plants;
  row[MC['Total Days']] = r.totalDays;
  row[MC['Strain']] = r.strain;
  row[MC['Flower Delivery']] = r.fd;
  row[MC['Joint Delivery']] = r.jd;
  row[MC['Proj 3.5g']] = r.p35;
  row[MC['Grams Per Tote']] = r.perTote;
  row[MC['Actual 3.5g']] = r.a35;
  row[MC['Target 7g']] = r.t7;
  row[MC['Target 28g']] = r.t28;
  row[MC['Wholesale Tier 1']] = r.wt1;
  row[MC['Proj Micros 14g']] = r.micros;
  row[MC['Actual Micros 14g']] = r.aMicros;
  row[MC['Wholesale Micros']] = r.wMicros;
  row[MC['Post Grind/Sift g']] = r.grind;
  row[MC['Trim %']] = r.trim;
  row[MC['1g Joints']] = r.j1;
  row[MC['2pk Joints']] = r.j2;
  row[MC['10pk Joints']] = r.j10;
  row[MC['Infused']] = r.infused ? 'TRUE' : 'FALSE';
  row[MC['Batch Count']] = r.batch;
  row[MC['Total Proj 1g']] = r.total1g;
  row[MC['Total Proj 0.5g']] = r.totalHalf;
  row[MC['Fresh Frozen']] = r.ff ? 'TRUE' : 'FALSE';
  row[MC['Cured']] = r.cured ? 'TRUE' : 'FALSE';
  row[MC['Synced At']] = stamp;
  row[MC['Source']] = r.source;
  return row;
}

// ============================================================================
//  STATUS, LOG, ALERTS
// ============================================================================

function _invRecord(run) {
  var tz = Session.getScriptTimeZone();
  var ss = SpreadsheetApp.openById(INV_PIPELINE.CRM_SHEET_ID);
  var props = PropertiesService.getScriptProperties();
  if (run.status === 'OK' || run.status === 'WARN') props.setProperty('INV_LAST_SUCCESS', String(Date.now()));
  var lastSuccess = Number(props.getProperty('INV_LAST_SUCCESS') || 0);

  var st = ss.getSheetByName(INV_PIPELINE.STATUS_TAB) || ss.insertSheet(INV_PIPELINE.STATUS_TAB);
  var s = run.stats;
  var rows = [
    ['Field', 'Value'],
    ['Status', run.status],
    ['Last Run', run.started],
    ['Last Success', lastSuccess ? new Date(lastSuccess) : ''],
    ['Product Allocation Rows', s.paRows || 0],
    ['Propagation Rows', s.propagationRows || 0],
    ['Harvest Weeks', s.harvestWeeks || 0],
    ['Furthest Delivery', s.furthestDelivery || ''],
    ['Yield Rates', s.rates || ''],
    ['Errors', run.errors.join('\n')],
    ['Warnings', run.warnings.join('\n')]
  ];
  st.clearContents();
  st.getRange(1, 1, rows.length, 2).setValues(rows);
  st.getRange('B3:B4').setNumberFormat('M/d/yyyy h:mm am/pm');
  st.getRange('A1:B1').setFontWeight('bold');
  st.getRange('B2').setBackground(run.status === 'OK' ? '#d9ead3' : run.status === 'WARN' ? '#fff2cc' : '#f4cccc');
  st.setColumnWidth(2, 700);
  st.getRange('B10:B11').setWrap(true);

  var lg = ss.getSheetByName(INV_PIPELINE.LOG_TAB) || ss.insertSheet(INV_PIPELINE.LOG_TAB);
  if (lg.getLastRow() === 0) {
    lg.appendRow(['Run At', 'Status', 'PA Rows', 'Propagation Rows', 'Harvest Weeks', 'Furthest Delivery', 'Seconds', 'Errors', 'Warnings']);
    lg.setFrozenRows(1);
  }
  lg.appendRow([run.started, run.status, s.paRows || 0, s.propagationRows || 0, s.harvestWeeks || 0,
                s.furthestDelivery || '', Math.round((Date.now() - run.started) / 1000),
                run.errors.join(' | '), run.warnings.join(' | ')]);
  var excess = lg.getLastRow() - 1 - INV_PIPELINE.LOG_KEEP_ROWS;
  if (excess > 0) lg.deleteRows(2, excess);
}

function _invAlert(run) {
  var props = PropertiesService.getScriptProperties();
  var wasAlerting = props.getProperty('INV_ALERTING') === '1';
  if (run.status === 'OK') {
    if (wasAlerting) {
      _invSendAlert('RESOLVED', 'Inventory sync is healthy again',
        ['Latest run succeeded: ' + (run.stats.harvestWeeks || 0) + ' harvest weeks, deliveries through ' + (run.stats.furthestDelivery || 'n/a') + '.'], []);
      props.deleteProperty('INV_ALERTING');
      props.deleteProperty('INV_ALERT_SIG');
    }
    return;
  }
  if (run.status === 'SKIPPED') return;
  // Same problems → email at most every ALERT_REPEAT_HOURS
  var sig = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, run.status + run.errors.join('|') + run.warnings.join('|')));
  var lastSig = props.getProperty('INV_ALERT_SIG');
  var lastAt = Number(props.getProperty('INV_ALERT_AT') || 0);
  if (sig === lastSig && Date.now() - lastAt < INV_PIPELINE.ALERT_REPEAT_HOURS * 36e5) return;

  _invSendAlert(run.status, run.status === 'ERROR' ? 'Inventory sync FAILED — portal is showing stale data' : 'Inventory sync needs attention',
                run.errors, run.warnings);
  props.setProperty('INV_ALERT_SIG', sig);
  props.setProperty('INV_ALERT_AT', String(Date.now()));
  props.setProperty('INV_ALERTING', '1');
}

function _invSendAlert(level, subject, errors, warnings) {
  var url = 'https://docs.google.com/spreadsheets/d/' + INV_PIPELINE.CRM_SHEET_ID;
  var body = '<div style="font-family:Arial,sans-serif;font-size:14px">';
  body += '<p><b>' + subject + '</b></p>';
  if (errors.length) body += '<p><b style="color:#C8102E">Errors</b></p><ul>' + errors.map(function (e) { return '<li>' + _invEsc(e) + '</li>'; }).join('') + '</ul>';
  if (warnings.length) body += '<p><b style="color:#B45309">Warnings</b></p><ul>' + warnings.map(function (e) { return '<li>' + _invEsc(e) + '</li>'; }).join('') + '</ul>';
  body += '<p>Status tab: <a href="' + url + '">CRM Master → Pipeline Status</a></p>';
  body += '<p style="color:#888;font-size:12px">Sent by the Redbird inventory pipeline (Apps Script). Repeats of the same issue are limited to once every ' +
          INV_PIPELINE.ALERT_REPEAT_HOURS + ' hours.</p></div>';
  MailApp.sendEmail({ to: INV_PIPELINE.ALERT_EMAILS.join(','), subject: '[Redbird Inventory ' + level + '] ' + subject, htmlBody: body });
}

// ============================================================================
//  SOURCE READING (by header name)
// ============================================================================

function _invReadTable(sheetId, tabName, fields, label) {
  var out = { label: label, rows: [], missingOptional: [], error: '' };
  var sh;
  try {
    sh = SpreadsheetApp.openById(sheetId).getSheetByName(tabName);
  } catch (e) {
    out.error = label + ': cannot open the spreadsheet (' + e.message + '). Check that the script owner still has access.';
    return out;
  }
  if (!sh) { out.error = label + ': tab "' + tabName + '" not found. Was it renamed?'; return out; }
  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) { out.error = label + ': tab is empty.'; return out; }
  var data = sh.getRange(1, 1, lastRow, lastCol).getValues();

  // Find the header row: the first of the top 10 rows that matches the most field names
  var best = -1, bestHits = 0, bestMap = null;
  for (var i = 0; i < Math.min(10, data.length); i++) {
    var map = _invMapHeaders(data[i], fields);
    var hits = Object.keys(map).length;
    if (hits > bestHits) { best = i; bestHits = hits; bestMap = map; }
  }
  if (best < 0) { out.error = label + ': could not find a header row.'; return out; }

  var missingReq = [];
  Object.keys(fields).forEach(function (k) {
    if (bestMap[k] === undefined) {
      if (fields[k].required) missingReq.push(fields[k].names[0]);
      else out.missingOptional.push(fields[k].names[0]);
    }
  });
  if (missingReq.length) {
    out.error = label + ': required column(s) not found: ' + missingReq.join(', ') +
                '. A column was likely renamed or deleted. Sync stopped; last good data is still live.';
    return out;
  }
  out.rows = data.slice(best + 1);
  out.get = function (row, key) { var c = bestMap[key]; return c === undefined ? '' : row[c]; };
  return out;
}

function _invMapHeaders(headerRow, fields) {
  var norm = headerRow.map(_invNorm);
  var map = {};
  Object.keys(fields).forEach(function (k) {
    var names = fields[k].names.map(_invNorm);
    for (var n = 0; n < names.length; n++) {
      var idx = norm.indexOf(names[n]);       // first match wins (e.g. duplicate "Per Tote")
      if (idx >= 0) { map[k] = idx; break; }
    }
  });
  return map;
}

// ============================================================================
//  SMALL HELPERS
// ============================================================================

function _invNorm(s) { return String(s || '').toLowerCase().replace(/[^a-z0-9%]/g, ''); }
function _invDay(d) { var x = new Date(d); x.setHours(0, 0, 0, 0); return x; }
function _invAdd(d, n) { var x = new Date(d); x.setDate(x.getDate() + n); x.setHours(0, 0, 0, 0); return x; }
function _invDate(v) {
  if (v instanceof Date && !isNaN(v)) return _invDay(v);
  if (typeof v === 'string' && /\d{1,4}[\/\-]\d{1,2}[\/\-]\d{1,4}/.test(v)) { var p = new Date(v); if (!isNaN(p)) return _invDay(p); }
  return null;
}
function _invKey(d, tz) { return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); }
function _invFmt(d, tz) { return Utilities.formatDate(d, tz, 'M/d/yyyy'); }
function _invNum(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  var n = parseFloat(String(v || '').replace(/[,$\s]/g, ''));
  return isNaN(n) ? 0 : n;
}
function _invPct(v) {
  var n = _invNum(String(v).replace('%', ''));
  if (typeof v === 'string' && v.indexOf('%') >= 0) n = n / 100;
  else if (n > 1) n = n / 100;
  return n;
}
function _invBool(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function _invMedian(a) {
  if (!a.length) return 0;
  var s = a.slice().sort(function (x, y) { return x - y; }), m = Math.floor(s.length / 2);
  return Math.round((s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2) * 100) / 100;
}
function _invEsc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

// ── MAIN SYNC (hourly trigger) ────────────────────────────────
function syncAll() {
  mirrorAccounts();
  mirrorInventory();
  mirrorBatchesInStock();
  reconcileAccounts();
  Logger.log('Redbird CRM sync complete: ' + new Date());
}

function createSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'syncAll') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncAll').timeBased().everyHours(1).create();
  Logger.log('Hourly sync trigger created.');
}

// ── RECONCILE ACCOUNTS — unifies Mirror + Cultivera history ───
function reconcileAccounts() {
  var t0 = Date.now();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mirror = ss.getSheetByName(ACCOUNTS_DEST_TAB);
  var orders = ss.getSheetByName('OrdersLog');
  if (!mirror) { console.log('[Reconcile] Accounts Mirror not found'); return; }
  if (!orders) { console.log('[Reconcile] OrdersLog not found'); return; }

  function findCol(headers, candidates) {
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i]).toLowerCase().trim();
      for (var j = 0; j < candidates.length; j++) {
        if (h === candidates[j].toLowerCase()) return i;
      }
    }
    return -1;
  }

  var mirrorData = mirror.getDataRange().getValues();
  if (mirrorData.length < 2) { console.log('[Reconcile] Mirror is empty'); return; }
  var mh = mirrorData[0];
  var mTradeIdx   = findCol(mh, ['Trade Name', 'Account', 'Account Name', 'Name']);
  var mLicenseIdx = findCol(mh, ['License #', 'License', 'License Number']);
  var mUbiIdx     = findCol(mh, ['UBI']);
  var mAddrIdx    = findCol(mh, ['Address', 'Mailing Street', 'Street']);
  var mCityIdx    = findCol(mh, ['City', 'Mailing City']);
  var mCountyIdx  = findCol(mh, ['County']);
  var mStateIdx   = findCol(mh, ['State', 'Mailing State']);
  var mZipIdx     = findCol(mh, ['Zip Code', 'Zip', 'Mailing Zip', 'Postal']);
  if (mTradeIdx === -1 || mLicenseIdx === -1) {
    console.log('[Reconcile] Mirror missing Trade Name or License # column'); return;
  }

  var mirrorByLicense = {};
  for (var i = 1; i < mirrorData.length; i++) {
    var lic = String(mirrorData[i][mLicenseIdx] || '').trim();
    if (lic) mirrorByLicense[lic] = mirrorData[i];
  }

  var ordersData = orders.getDataRange().getValues();
  if (ordersData.length < 2) { console.log('[Reconcile] OrdersLog is empty'); return; }
  var oh = ordersData[0];
  var oLicenseIdx = findCol(oh, ['License #', 'License', 'License Number']);
  var oClientIdx  = findCol(oh, ['Client', 'Account', 'Account Name', 'Customer']);
  var oDateIdx    = findCol(oh, ['Submitted Date', 'Date', 'Order Date']);
  var oOrderIdx   = findCol(oh, ['Order #', 'Order Number', 'Order Id']);
  if (oLicenseIdx === -1) { console.log('[Reconcile] OrdersLog missing License # column'); return; }

  var licenseStats = {};
  for (var i = 1; i < ordersData.length; i++) {
    var r = ordersData[i];
    var lic = String(r[oLicenseIdx] || '').trim();
    if (!lic) continue;
    var client = oClientIdx > -1 ? r[oClientIdx] : '';
    var date = oDateIdx > -1 ? r[oDateIdx] : null;
    var orderId = oOrderIdx > -1 ? r[oOrderIdx] : '';
    if (!licenseStats[lic]) {
      licenseStats[lic] = { client: client, lastDate: date, orderIds: {} };
    }
    if (orderId) licenseStats[lic].orderIds[orderId] = true;
    if (date instanceof Date) {
      if (!(licenseStats[lic].lastDate instanceof Date) || date > licenseStats[lic].lastDate) {
        licenseStats[lic].lastDate = date;
      }
    } else if (date && !licenseStats[lic].lastDate) {
      licenseStats[lic].lastDate = date;
    }
  }

  var crmTab = ss.getSheetByName('CRM Accounts');
  var existingAddresses = {};
  if (crmTab) {
    var existingData = crmTab.getDataRange().getValues();
    if (existingData.length >= 2) {
      var ch = existingData[0];
      var cLicIdx = findCol(ch, ['License #']);
      var cAddrIdx = findCol(ch, ['Address']);
      var cCityIdx = findCol(ch, ['City']);
      var cCountyIdx = findCol(ch, ['County']);
      var cStateIdx = findCol(ch, ['State']);
      var cZipIdx = findCol(ch, ['Zip']);
      var cUbiIdx = findCol(ch, ['UBI']);
      var cSourceIdx = findCol(ch, ['Source']);
      for (var i = 1; i < existingData.length; i++) {
        var lic = String(existingData[i][cLicIdx] || '').trim();
        if (!lic) continue;
        var src = cSourceIdx > -1 ? existingData[i][cSourceIdx] : '';
        if (src === 'Cultivera-only') {
          existingAddresses[lic] = {
            addr:   cAddrIdx   > -1 ? existingData[i][cAddrIdx]   : '',
            city:   cCityIdx   > -1 ? existingData[i][cCityIdx]   : '',
            county: cCountyIdx > -1 ? existingData[i][cCountyIdx] : '',
            state:  cStateIdx  > -1 ? existingData[i][cStateIdx]  : '',
            zip:    cZipIdx    > -1 ? existingData[i][cZipIdx]    : '',
            ubi:    cUbiIdx    > -1 ? existingData[i][cUbiIdx]    : ''
          };
        }
      }
    }
  }

  var allLicenses = {};
  Object.keys(mirrorByLicense).forEach(function(k) { allLicenses[k] = true; });
  Object.keys(licenseStats).forEach(function(k) { allLicenses[k] = true; });

  var crmRows = [];
  Object.keys(allLicenses).forEach(function(lic) {
    var mirrorRow = mirrorByLicense[lic];
    var stats = licenseStats[lic];
    var preserved = existingAddresses[lic] || {};
    var tradeName, addr, city, county, state, zip, ubi, source;
    if (mirrorRow) {
      tradeName = mirrorRow[mTradeIdx];
      addr   = mAddrIdx   > -1 ? mirrorRow[mAddrIdx]   : '';
      city   = mCityIdx   > -1 ? mirrorRow[mCityIdx]   : '';
      county = mCountyIdx > -1 ? mirrorRow[mCountyIdx] : '';
      state  = mStateIdx  > -1 ? mirrorRow[mStateIdx]  : '';
      zip    = mZipIdx    > -1 ? mirrorRow[mZipIdx]    : '';
      ubi    = mUbiIdx    > -1 ? mirrorRow[mUbiIdx]    : '';
      source = 'Mirror';
    } else {
      tradeName = stats ? stats.client : '';
      addr   = preserved.addr   || '';
      city   = preserved.city   || '';
      county = preserved.county || '';
      state  = preserved.state  || '';
      zip    = preserved.zip    || '';
      ubi    = preserved.ubi    || '';
      source = 'Cultivera-only';
    }
    var lastInv = stats ? stats.lastDate : '';
    var invCount = stats ? Object.keys(stats.orderIds).length : 0;
    crmRows.push([tradeName, lic, ubi, addr, city, county, state, zip, source, lastInv, invCount]);
  });

  crmRows.sort(function(a, b) {
    var srcA = a[8], srcB = b[8];
    if (srcA !== srcB) return srcA === 'Mirror' ? -1 : 1;
    return String(a[0] || '').toLowerCase().localeCompare(String(b[0] || '').toLowerCase());
  });

  if (!crmTab) crmTab = ss.insertSheet('CRM Accounts');
  else crmTab.clear();
  var crmHeaders = ['Trade Name', 'License #', 'UBI', 'Address', 'City', 'County', 'State', 'Zip', 'Source', 'Last Invoice', 'Invoice Count'];
  var crmOutput = [crmHeaders].concat(crmRows);
  crmTab.getRange(1, 1, crmOutput.length, crmHeaders.length).setValues(crmOutput);
  crmTab.getRange(1, 1, 1, crmHeaders.length).setBackground('#C8102E').setFontColor('#FFFFFF').setFontWeight('bold');
  crmTab.setFrozenRows(1);

  var crmByLicense = {};
  crmRows.forEach(function(r) { crmByLicense[String(r[1])] = r[0]; });

  var mirrorAcctColIdx = findCol(oh, ['Mirror Account']);
  if (mirrorAcctColIdx === -1) {
    mirrorAcctColIdx = oh.length;
    orders.getRange(1, mirrorAcctColIdx + 1).setValue('Mirror Account')
      .setBackground('#C8102E').setFontColor('#FFFFFF').setFontWeight('bold');
  }

  var stampValues = [];
  for (var i = 1; i < ordersData.length; i++) {
    var lic = String(ordersData[i][oLicenseIdx] || '').trim();
    stampValues.push([crmByLicense[lic] || '']);
  }
  if (stampValues.length > 0) {
    orders.getRange(2, mirrorAcctColIdx + 1, stampValues.length, 1).setValues(stampValues);
  }

  var elapsed = ((Date.now() - t0) / 1000).toFixed(1);
  var mirrorCount = crmRows.filter(function(r) { return r[8] === 'Mirror'; }).length;
  var cultiveraOnly = crmRows.filter(function(r) { return r[8] === 'Cultivera-only'; }).length;
  console.log('[Reconcile] Build complete in ' + elapsed + 's: total ' + crmRows.length + ', Mirror ' + mirrorCount + ', Cultivera-only ' + cultiveraOnly);
  return { total: crmRows.length, mirror: mirrorCount, cultiveraOnly: cultiveraOnly };
}

// ── ACCOUNT CONTACTS — joins curated contacts to CRM Accounts ──
function buildAccountContacts() {
  var t0 = Date.now();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var crmAccounts = ss.getSheetByName('CRM Accounts');
  var accounts = ss.getSheetByName('Accounts');
  if (!crmAccounts) { console.log('[AccountContacts] CRM Accounts tab not found — run reconcileAccounts first'); return; }
  if (!accounts) { console.log('[AccountContacts] Accounts tab not found'); return; }

  function normAddr(s) {
    var raw = String(s == null ? '' : s).toUpperCase().replace(/[.,#]/g, ' ').replace(/\s+/g, ' ').trim();
    if (!raw) return '';
    var expansions = {
      'N': 'NORTH', 'S': 'SOUTH', 'E': 'EAST', 'W': 'WEST',
      'NE': 'NORTHEAST', 'NW': 'NORTHWEST', 'SE': 'SOUTHEAST', 'SW': 'SOUTHWEST',
      'ST': 'STREET', 'AVE': 'AVENUE', 'BLVD': 'BOULEVARD', 'RD': 'ROAD',
      'DR': 'DRIVE', 'LN': 'LANE', 'CT': 'COURT', 'PL': 'PLACE',
      'PKWY': 'PARKWAY', 'TER': 'TERRACE', 'CIR': 'CIRCLE',
      'HWY': 'HIGHWAY', 'TRL': 'TRAIL', 'PT': 'POINT', 'SQ': 'SQUARE',
      'PLZ': 'PLAZA', 'XING': 'CROSSING', 'WA': 'WAY', 'WY': 'WAY',
      'STE': 'SUITE', 'APT': 'APARTMENT', 'BLDG': 'BUILDING'
    };
    var tokens = raw.split(' ');
    for (var i = 0; i < tokens.length; i++) {
      if (expansions[tokens[i]]) tokens[i] = expansions[tokens[i]];
    }
    var dedup = [];
    for (var j = 0; j < tokens.length; j++) {
      if (j >= 2 && tokens[j] === tokens[j - 2] && tokens[j - 1] === dedup[dedup.length - 1]) continue;
      if (dedup.length > 0 && dedup[dedup.length - 1] === tokens[j]) continue;
      dedup.push(tokens[j]);
    }
    var unitMarkers = { 'SUITE': 1, 'APARTMENT': 1, 'BUILDING': 1, 'UNIT': 1, 'LOT': 1,
                        'FLOOR': 1, 'FL': 1, 'OFFICE': 1, 'ROOM': 1, 'RM': 1,
                        'SMP': 1, 'SPC': 1, 'TRLR': 1, 'DEPT': 1 };
    for (var k = 0; k < dedup.length; k++) {
      if (unitMarkers[dedup[k]]) { dedup = dedup.slice(0, k); break; }
    }
    return dedup.join(' ');
  }
  function normCity(s) {
    return String(s == null ? '' : s).toUpperCase().replace(/[.,]/g, '').replace(/\s+/g, ' ').trim();
  }
  function normName(s) {
    return String(s == null ? '' : s).toLowerCase().replace(/\s+/g, ' ').trim();
  }
  function findCol(headers, candidates) {
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i]).toLowerCase().trim();
      for (var j = 0; j < candidates.length; j++) {
        if (h === candidates[j].toLowerCase()) return i;
      }
    }
    return -1;
  }

  var crmData = crmAccounts.getDataRange().getValues();
  if (crmData.length < 2) { console.log('[AccountContacts] CRM Accounts is empty'); return; }
  var crmHeaders = crmData[0];
  var cTradeIdx   = findCol(crmHeaders, ['Trade Name']);
  var cLicenseIdx = findCol(crmHeaders, ['License #']);
  var cAddrIdx    = findCol(crmHeaders, ['Address']);
  var cCityIdx    = findCol(crmHeaders, ['City']);
  var cStateIdx   = findCol(crmHeaders, ['State']);
  var cZipIdx     = findCol(crmHeaders, ['Zip']);
  var cSourceIdx  = findCol(crmHeaders, ['Source']);
  if (cTradeIdx === -1 || cAddrIdx === -1) {
    console.log('[AccountContacts] CRM Accounts missing Trade Name or Address column'); return;
  }

  var crmByAddrCity = {}, crmByAddrOnly = {}, crmByName = {};
  for (var i = 1; i < crmData.length; i++) {
    var row = crmData[i];
    var addr = normAddr(row[cAddrIdx]);
    var city = cCityIdx > -1 ? normCity(row[cCityIdx]) : '';
    var trade = row[cTradeIdx];
    if (addr) {
      if (city) {
        var ckey = addr + '|' + city;
        if (!crmByAddrCity[ckey]) crmByAddrCity[ckey] = row;
      }
      if (!crmByAddrOnly[addr]) crmByAddrOnly[addr] = row;
    }
    if (trade) {
      var n = normName(trade);
      if (n && !crmByName[n]) crmByName[n] = row;
    }
  }

  var acctRange = accounts.getDataRange().getValues();
  if (acctRange.length < 2) { console.log('[AccountContacts] Accounts tab is empty'); return; }
  var acctHeaders = acctRange[0];
  var aFirstIdx  = findCol(acctHeaders, ['First Name', 'FirstName', 'Given Name']);
  var aLastIdx   = findCol(acctHeaders, ['Last Name', 'LastName', 'Surname', 'Family Name']);
  var aFullIdx   = findCol(acctHeaders, ['Full Name', 'Name', 'Contact Name']);
  var aTitleIdx  = findCol(acctHeaders, ['Title', 'Job Title', 'Role']);
  var aAcctIdx   = findCol(acctHeaders, ['Account Name', 'Account', 'Company']);
  var aStreetIdx = findCol(acctHeaders, ['Mailing Street', 'Street', 'Address']);
  var aCityIdx   = findCol(acctHeaders, ['Mailing City', 'City']);
  var aPhoneIdx  = findCol(acctHeaders, ['Phone', 'Phone Number']);
  var aEmailIdx  = findCol(acctHeaders, ['Email', 'Email Address']);
  if (aAcctIdx === -1) { console.log('[AccountContacts] Accounts missing Account Name column'); return; }

  var matched = [], unmatched = [];
  var byAddrCity = 0, byAddrOnly = 0, byName = 0;

  for (var ci = 1; ci < acctRange.length; ci++) {
    var r = acctRange[ci];
    var acctName = aAcctIdx > -1 ? r[aAcctIdx] : '';
    var first = aFirstIdx > -1 ? r[aFirstIdx] : '';
    var last  = aLastIdx > -1 ? r[aLastIdx] : '';
    var fullName = aFullIdx > -1 ? r[aFullIdx] : ((first || last) ? String(first || '').trim() + (first && last ? ' ' : '') + String(last || '').trim() : '');
    var title = aTitleIdx > -1 ? r[aTitleIdx] : '';
    var street = aStreetIdx > -1 ? r[aStreetIdx] : '';
    var city = aCityIdx > -1 ? r[aCityIdx] : '';
    var email = aEmailIdx > -1 ? r[aEmailIdx] : '';
    var phone = aPhoneIdx > -1 ? r[aPhoneIdx] : '';
    if (!acctName && !fullName && !email) continue;

    var crmMatch = null, matchType = '';
    if (street && city) {
      var k1 = normAddr(street) + '|' + normCity(city);
      if (crmByAddrCity[k1]) { crmMatch = crmByAddrCity[k1]; matchType = 'address+city'; byAddrCity++; }
    }
    if (!crmMatch && street) {
      var k2 = normAddr(street);
      if (k2 && crmByAddrOnly[k2]) { crmMatch = crmByAddrOnly[k2]; matchType = 'address'; byAddrOnly++; }
    }
    if (!crmMatch && acctName) {
      var k3 = normName(acctName);
      if (k3 && crmByName[k3]) { crmMatch = crmByName[k3]; matchType = 'name'; byName++; }
    }

    if (crmMatch) {
      matched.push([
        crmMatch[cTradeIdx],
        cLicenseIdx > -1 ? crmMatch[cLicenseIdx] : '',
        cSourceIdx  > -1 ? crmMatch[cSourceIdx]  : '',
        fullName, title, email, phone,
        cAddrIdx  > -1 ? crmMatch[cAddrIdx]  : '',
        cCityIdx  > -1 ? crmMatch[cCityIdx]  : '',
        cStateIdx > -1 ? crmMatch[cStateIdx] : '',
        cZipIdx   > -1 ? crmMatch[cZipIdx]   : '',
        matchType, acctName
      ]);
    } else {
      unmatched.push([acctName, fullName, title, email, phone, street, city]);
    }
  }

  matched.sort(function(a, b) {
    var n1 = String(a[0] || '').toLowerCase();
    var n2 = String(b[0] || '').toLowerCase();
    if (n1 !== n2) return n1 < n2 ? -1 : 1;
    return String(a[3] || '').toLowerCase().localeCompare(String(b[3] || '').toLowerCase());
  });
  unmatched.sort(function(a, b) {
    return String(a[0] || '').toLowerCase().localeCompare(String(b[0] || '').toLowerCase());
  });

  var contactsHeaders = ['Trade Name', 'License #', 'Source', 'Full Name', 'Title', 'Email', 'Phone',
    'Address', 'City', 'State', 'Zip', 'Match Type', 'Source Account Name'];
  var unmatchedHeaders = ['Account Name', 'Full Name', 'Title', 'Email', 'Phone', 'Address', 'City'];

  var contactsTab = ss.getSheetByName('Account Contacts');
  if (!contactsTab) contactsTab = ss.insertSheet('Account Contacts');
  else contactsTab.clear();
  var contactsOutput = [contactsHeaders].concat(matched);
  contactsTab.getRange(1, 1, contactsOutput.length, contactsHeaders.length).setValues(contactsOutput);
  contactsTab.getRange(1, 1, 1, contactsHeaders.length).setBackground('#C8102E').setFontColor('#FFFFFF').setFontWeight('bold');
  contactsTab.setFrozenRows(1);

  var unmatchedTab = ss.getSheetByName('Unmatched Contacts');
  if (!unmatchedTab) unmatchedTab = ss.insertSheet('Unmatched Contacts');
  else unmatchedTab.clear();
  var unmatchedOutput = [unmatchedHeaders].concat(unmatched);
  unmatchedTab.getRange(1, 1, unmatchedOutput.length, unmatchedHeaders.length).setValues(unmatchedOutput);
  unmatchedTab.getRange(1, 1, 1, unmatchedHeaders.length).setBackground('#C8102E').setFontColor('#FFFFFF').setFontWeight('bold');
  unmatchedTab.setFrozenRows(1);

  var elapsed = ((Date.now() - t0) / 1000).toFixed(1);
  console.log('[AccountContacts] Build complete in ' + elapsed + 's: matched ' + matched.length + ', unmatched ' + unmatched.length);
  return { matched: matched.length, unmatched: unmatched.length };
}

// ============================================================
//  ORDER CONFIRMATION EMAIL
// ============================================================
function sendOrderConfirmationEmail(toEmail, contactName, account, orderId, items, grandTotal, notes, isUpdate) {
  var now = new Date();
  var orderDate = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMMM d, yyyy');
  notes = (notes || '').toString().trim();

  var acct = getAccountInfo(account);

  var weekMap = {}, weekOrder = [];
  items.forEach(function(item) {
    var wk = item.deliveryWeek || 'TBD';
    if (!weekMap[wk]) { weekMap[wk] = []; weekOrder.push(wk); }
    weekMap[wk].push(item);
  });

  var totalAmt = items.reduce(function(s, i) { return s + (parseFloat(i.lineTotal) || 0); }, 0);
  if (grandTotal) totalAmt = parseFloat(grandTotal) || totalAmt;
  var totalQty = items.reduce(function(s, i) { return s + (parseInt(i.qty) || 0); }, 0);

  var itemsContent = weekOrder.map(function(wk) {
    var lineItems = weekMap[wk];
    var itemCount = lineItems.reduce(function(s, i) { return s + (parseInt(i.qty) || 0); }, 0);
    var rows = lineItems.map(function(it) {
      var qty = parseInt(it.qty) || 0;
      var lineTotal = parseFloat(it.lineTotal) || 0;
      return '<tr><td style="padding:12px 0;border-bottom:1px solid #f0f0f0">'
        + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>'
        + '<td style="vertical-align:top">'
        + '<div style="font-size:14px;color:#1a1a1a;font-weight:600;line-height:1.3">' + escapeHtml(it.strain || '') + '</div>'
        + '<div style="font-size:12px;color:#888;margin-top:3px">' + escapeHtml(it.productSize || '')
        + ' \u00b7 Qty ' + qty + ' \u00b7 ' + formatMoney(parseFloat(it.unitPrice) || 0) + ' ea</div></td>'
        + '<td style="text-align:right;font-size:14px;font-weight:700;color:#1a1a1a;white-space:nowrap;vertical-align:top;padding-left:16px">'
        + formatMoney(lineTotal) + '</td>'
        + '</tr></table></td></tr>';
    }).join('');
    return '<div style="font-size:11px;font-weight:700;color:#D2232A;margin:14px 0 4px;text-transform:uppercase;letter-spacing:0.06em">'
      + escapeHtml(wk) + ' \u00b7 ' + itemCount + ' item' + (itemCount === 1 ? '' : 's') + '</div>'
      + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">' + rows + '</table>';
  }).join('');

  var deliveryLocationContent;
  if (acct) {
    deliveryLocationContent = '<div style="font-size:14px;font-weight:700;color:#1a1a1a;margin-bottom:4px">' + escapeHtml(acct.tradeName) + '</div>'
      + (acct.address ? '<div style="font-size:13px;color:#555;line-height:1.6">' + escapeHtml(acct.address) + '</div>' : '')
      + (acct.city || acct.state || acct.zip ? '<div style="font-size:13px;color:#555;line-height:1.6">'
          + escapeHtml(acct.city) + (acct.city && acct.state ? ', ' : '') + escapeHtml(acct.state) + ' '
          + escapeHtml(String(acct.zip || '')) + '</div>' : '')
      + (acct.county ? '<div style="font-size:12px;color:#888;line-height:1.6">' + escapeHtml(acct.county) + ' County</div>' : '')
      + (acct.license || acct.ubi ? '<div style="font-size:11px;color:#999;margin-top:10px;line-height:1.5">'
          + (acct.license ? 'License # ' + escapeHtml(String(acct.license)) : '')
          + (acct.license && acct.ubi ? '<br>' : '')
          + (acct.ubi ? 'UBI ' + escapeHtml(String(acct.ubi)) : '')
          + '</div>' : '');
  } else {
    deliveryLocationContent = '<div style="font-size:14px;font-weight:700;color:#1a1a1a">' + escapeHtml(account || '') + '</div>'
      + '<div style="font-size:12px;color:#999;margin-top:4px;font-style:italic">Address details on file</div>';
  }

  var buyerContactContent = '<div style="font-size:14px;color:#1a1a1a;font-weight:600">' + escapeHtml(contactName || 'Buyer') + '</div>'
    + (toEmail ? '<div style="font-size:13px;margin-top:4px"><a href="mailto:' + escapeAttr(toEmail) + '" style="color:#D2232A;text-decoration:none">' + escapeHtml(toEmail) + '</a></div>' : '');

  function section(label, content) {
    return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom:36px"><tr>'
      + '<td valign="top" width="160" style="padding-right:24px"><div style="font-size:14px;font-weight:700;color:#1a1a1a;line-height:1.3">' + label + '</div></td>'
      + '<td valign="top">' + content + '</td>'
      + '</tr></table>';
  }

  var summaryHtml = '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">'
    + '<tr><td style="padding:8px 0;font-size:13px;color:#666;text-align:right">Subtotal (' + totalQty + ' items)</td>'
    + '<td style="padding:8px 0 8px 32px;font-size:13px;color:#1a1a1a;text-align:right;font-weight:600;width:120px">' + formatMoney(totalAmt) + '</td></tr>'
    + '<tr><td style="padding:14px 0 0;font-size:15px;color:#1a1a1a;text-align:right;font-weight:700;border-top:2px solid #1a1a1a">Reservation Total</td>'
    + '<td style="padding:14px 0 0 32px;font-size:20px;color:#D2232A;text-align:right;font-weight:800;border-top:2px solid #1a1a1a">' + formatMoney(totalAmt) + '</td></tr>'
    + '</table>';

  var portalUrl = 'https://joshruhlman-creator.github.io/redbird-crm/'
    + '?acct=' + encodeURIComponent(account || '')
    + '&contact=' + encodeURIComponent(contactName || '')
    + '&email=' + encodeURIComponent(toEmail || '')
    + '&reservation=' + encodeURIComponent(orderId || '');

  var html = '<table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" align="center" style="margin:0 auto;background:#ffffff;font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Helvetica,Arial,sans-serif">'
    + '<tr><td style="padding:36px 40px 4px;text-align:center">'
    + '<div style="font-size:30px;font-weight:800;color:#1a1a1a;line-height:1.2">' + (isUpdate ? 'Your reservation has been updated.' : 'Thanks for your reservation.') + '</div>'
    + '</td></tr>'
    + '<tr><td style="padding:0 40px 28px;text-align:center;border-bottom:1px solid #e8e8e6">'
    + '<div style="font-size:13px;color:#666;line-height:1.7;margin-top:8px">Order Number: <span style="color:#D2232A;font-weight:600">' + escapeHtml(orderId || '') + '</span></div>'
    + '<div style="font-size:13px;color:#666;line-height:1.7">Ordered on: ' + escapeHtml(orderDate) + '</div>'
    + '</td></tr>'
    + '<tr><td style="padding:32px 40px 0">' + section('Items for<br>Delivery', itemsContent) + '</td></tr>'
    + '<tr><td style="padding:0 40px">' + section('Delivery<br>Location', deliveryLocationContent) + '</td></tr>'
    + '<tr><td style="padding:0 40px">' + section('Buyer<br>Contact', buyerContactContent) + '</td></tr>'
    + (notes ? '<tr><td style="padding:0 40px">' + section('Order<br>Notes', '<div style="font-size:13px;color:#1a1a1a;line-height:1.6;background:#fafaf9;border-left:3px solid #D2232A;padding:12px 14px;border-radius:4px;white-space:pre-wrap">' + escapeHtml(notes) + '</div>') + '</td></tr>' : '')
    + '<tr><td style="padding:0 40px 36px">' + summaryHtml + '</td></tr>'
    + '<tr><td style="padding:0 40px 40px;text-align:center">'
    + '<a href="' + escapeAttr(portalUrl) + '" style="display:inline-block;background:#1a1a1a;color:#ffffff;text-decoration:none;padding:13px 32px;border-radius:7px;font-size:13px;font-weight:700;letter-spacing:0.02em">View or modify reservation \u2192</a>'
    + '</td></tr>'
    + '<tr><td style="background:#fafaf9;padding:24px 40px;border-top:1px solid #eeede8">'
    + '<div style="font-size:11px;color:#888;line-height:1.6;text-align:center">'
    + 'Questions? Reply directly to this email or contact <a href="mailto:josh.ruhlman@redbird-cannabis.com" style="color:#D2232A;text-decoration:none;font-weight:600">josh.ruhlman@redbird-cannabis.com</a><br>'
    + 'Redbird Cannabis \u00b7 TrueAeroGrow LLC \u00b7 Washington i502 Producer/Processor</div>'
    + '</td></tr></table>';

  var subject = (isUpdate ? 'Reservation Updated - ' : 'Reservation Confirmed - ') + (orderId || '') + ' - Redbird Cannabis';
  var plainText = 'Reservation ' + (orderId || '') + ' confirmed for ' + (account || '') + '. Total: ' + formatMoney(totalAmt) + '.';

  GmailApp.sendEmail(toEmail, subject, plainText, {
    htmlBody: html,
    name: 'Redbird Cannabis',
    replyTo: 'josh.ruhlman@redbird-cannabis.com',
    bcc: RESERVATION_BCC
  });
  Logger.log('[OrderConfirm] Sent to ' + toEmail + ' for ' + orderId);
}

// ============================================================
//  CANCELLATION CONFIRMATION EMAIL
// ============================================================
function sendCancellationEmail(toEmail, contactName, account, reservationId, items) {
  var firstName = (contactName || '').split(' ')[0] || 'there';
  var cancelDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM d, yyyy');
  items = items || [];

  var itemRows = items.map(function(it) {
    return '<tr><td style="padding:8px 0;border-bottom:1px solid #f0f0f0;font-size:13px;color:#555">'
      + escapeHtml(String(it.strain || '')) + ' \u00b7 ' + escapeHtml(String(it.productSize || ''))
      + ' \u00b7 Qty ' + (parseInt(it.qty) || 0) + '</td></tr>';
  }).join('');

  var html = '<table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" align="center" style="margin:0 auto;background:#ffffff;font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Helvetica,Arial,sans-serif">'
    + '<tr><td style="padding:36px 40px 4px;text-align:center">'
    + '<div style="font-size:28px;font-weight:800;color:#1a1a1a;line-height:1.2">Reservation cancelled</div></td></tr>'
    + '<tr><td style="padding:8px 40px 28px;text-align:center;border-bottom:1px solid #e8e8e6">'
    + '<div style="font-size:13px;color:#666;line-height:1.7">Hi ' + escapeHtml(firstName) + ', your reservation has been cancelled and the reserved quantities released.</div>'
    + '<div style="font-size:13px;color:#666;line-height:1.7;margin-top:8px">Reservation: <span style="color:#D2232A;font-weight:600">' + escapeHtml(String(reservationId || '')) + '</span></div>'
    + '<div style="font-size:13px;color:#666;line-height:1.7">Cancelled on: ' + escapeHtml(cancelDate) + '</div></td></tr>'
    + (itemRows ? '<tr><td style="padding:24px 40px 8px"><div style="font-size:11px;font-weight:700;color:#D2232A;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:6px">Cancelled items</div>'
        + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">' + itemRows + '</table></td></tr>' : '')
    + '<tr><td style="padding:24px 40px 36px;text-align:center">'
    + '<div style="font-size:13px;color:#666;line-height:1.6">Changed your mind? You can place a new reservation any time.</div></td></tr>'
    + '<tr><td style="background:#fafaf9;padding:24px 40px;border-top:1px solid #eeede8">'
    + '<div style="font-size:11px;color:#888;line-height:1.6;text-align:center">'
    + 'Questions? Reply to this email or contact <a href="mailto:josh.ruhlman@redbird-cannabis.com" style="color:#D2232A;text-decoration:none;font-weight:600">josh.ruhlman@redbird-cannabis.com</a><br>'
    + 'Redbird Cannabis \u00b7 TrueAeroGrow LLC \u00b7 Washington i502 Producer/Processor</div></td></tr></table>';

  GmailApp.sendEmail(toEmail, 'Reservation Cancelled - ' + (reservationId || '') + ' - Redbird Cannabis',
    'Reservation ' + (reservationId || '') + ' for ' + (account || '') + ' has been cancelled.',
    { htmlBody: html, name: 'Redbird Cannabis', replyTo: 'josh.ruhlman@redbird-cannabis.com', bcc: RESERVATION_BCC });
}

// ── HELPERS for email rendering ──────────────────────────────
function formatMoney(n) {
  var num = parseFloat(n) || 0;
  return '$' + num.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function escapeHtml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function escapeAttr(s) {
  return String(s == null ? '' : s).replace(/"/g, '&quot;');
}

// ── TESTS (run from the editor) ──────────────────────────────
function TEST_sendConfirmationEmail() {
  var TEST_EMAIL = 'josh.ruhlman@redbird-cannabis.com';
  var fakeItems = [
    { deliveryWeek: '2026-05-18', strain: 'Agate Dreams',        productSize: '3.5g Flower',    qty: 12, unitPrice: 10,   lineTotal: 120 },
    { deliveryWeek: '2026-05-18', strain: 'Lemon Cherry Gelato', productSize: '28g Flower',     qty: 5,  unitPrice: 60,   lineTotal: 300 },
    { deliveryWeek: '2026-05-25', strain: 'Zaitai',              productSize: '0.5g Vape Cart', qty: 20, unitPrice: 15,   lineTotal: 300 }
  ];
  var grandTotal = fakeItems.reduce(function(s, i) { return s + i.lineTotal; }, 0);
  var fakeNotes = 'Please deliver before 10am. Ask for Mike at the back door.';
  sendOrderConfirmationEmail(TEST_EMAIL, 'Test Buyer', 'Green Lady East',
    'TEST-' + new Date().getTime().toString().slice(-6), fakeItems, grandTotal, fakeNotes);
  Logger.log('Test confirmation email triggered to: ' + TEST_EMAIL);
}

function TEST_sendCancellationEmail() {
  sendCancellationEmail('josh.ruhlman@redbird-cannabis.com', 'Test Buyer', 'Green Lady East',
    'TEST-' + new Date().getTime().toString().slice(-6),
    [{ strain: 'Agate Dreams', productSize: '3.5g Flower', qty: 12 },
     { strain: 'Zaitai', productSize: '0.5g Vape Cart', qty: 20 }]);
  Logger.log('Test cancellation email triggered.');
}

function TEST_taskCompletions() {
  var key = 'test account|test task|2026-05-04';
  console.log('Inserting:', upsertTaskCompletion(key, 'Test Account', 'Test Task', '2026-05-04', 'Josh', 'Testing'));
  console.log('Re-upserting (should update):', upsertTaskCompletion(key, 'Test Account', 'Test Task', '2026-05-04', 'Josh', 'Updated note'));
  console.log('Deleting:', deleteTaskCompletion(key));
}

function TEST_contacts() {
  var id = upsertContact('', 'Test Account', 'Jane', 'Doe', 'Buyer', 'jane@example.com', '555-1234', '', 'Test note');
  console.log('Created contact:', id);
  upsertContact(id, 'Test Account', 'Jane', 'Smith', 'Senior Buyer', 'jane@example.com', '555-1234', '555-5678', 'Updated');
  console.log('Delete result:', deleteContact(id));
}

// ============================================================
//  CONTACTS_IMPORT / SALESFORCE_REBUILD
// ============================================================
var CONTACTS_IMPORT_TAB = 'Contacts Import';
var SF_ORDERS_TAB       = 'Salesforce Order History';

function _normalizeOrderNum(v) {
  if (v === null || v === undefined) return '';
  var s = String(v).trim();
  s = s.replace(/^ORD\s*[-\u2013\u2014]?\s*/i, '');
  s = s.replace(/\.0+$/, '');
  return s.trim();
}

function linkOrdersFromSalesforce() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sf = ss.getSheetByName(SF_ORDERS_TAB);
  var orders = ss.getSheetByName('OrdersLog');
  if (!sf || !orders) { console.log('[LinkOrders] Missing required tab(s)'); return { ok: false }; }
  var sfData = sf.getDataRange().getValues();
  var sfHeaderRowIdx = -1;
  for (var r = 0; r < Math.min(25, sfData.length); r++) {
    var rowStr = sfData[r].map(function(c) { return String(c || ''); }).join('|');
    if (rowStr.indexOf('Opportunity Name') > -1 && rowStr.indexOf('Account Name') > -1) { sfHeaderRowIdx = r; break; }
  }
  if (sfHeaderRowIdx === -1) return { ok: false, error: 'SF header not found' };
  var sfHeaders = sfData[sfHeaderRowIdx].map(function(h) { return String(h || '').trim(); });
  var oppIdx = -1, acctIdx = -1;
  for (var i = 0; i < sfHeaders.length; i++) {
    var clean = sfHeaders[i].replace(/[\u2191\u2193\s]+$/, '').trim();
    if (clean === 'Opportunity Name') oppIdx = i;
    else if (clean === 'Account Name') acctIdx = i;
  }
  var orderToAccount = {};
  for (var r = sfHeaderRowIdx + 1; r < sfData.length; r++) {
    var opp = sfData[r][oppIdx];
    var acct = String(sfData[r][acctIdx] || '').trim();
    if (!opp || !acct) continue;
    var oppStr = _normalizeOrderNum(opp);
    if (oppStr) orderToAccount[oppStr] = acct;
  }
  var oData = orders.getDataRange().getValues();
  var oh = oData[0];
  var oOrderIdx = -1, oMirrorIdx = -1;
  for (var i = 0; i < oh.length; i++) {
    var h = String(oh[i]).toLowerCase().trim();
    if (h === 'order #' || h === 'order number') oOrderIdx = i;
    else if (h === 'mirror account') oMirrorIdx = i;
  }
  if (oOrderIdx === -1 || oMirrorIdx === -1) return { ok: false, error: 'OrdersLog cols missing' };
  var matched = 0, unmatched = 0, changed = false;
  for (var i = 1; i < oData.length; i++) {
    var ordNum = _normalizeOrderNum(oData[i][oOrderIdx]);
    var hit = orderToAccount[ordNum];
    if (hit) { matched++; if (oData[i][oMirrorIdx] !== hit) { oData[i][oMirrorIdx] = hit; changed = true; } }
    else unmatched++;
  }
  if (changed) {
    var col = oData.slice(1).map(function(r) { return [r[oMirrorIdx]]; });
    orders.getRange(2, oMirrorIdx + 1, col.length, 1).setValues(col);
  }
  console.log('[LinkOrders] Matched ' + matched + ' / ' + (matched + unmatched));
  return { ok: true, matched: matched, unmatched: unmatched };
}

function diagnoseOrderJoin() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var orders = ss.getSheetByName('OrdersLog');
  var sf = ss.getSheetByName(SF_ORDERS_TAB);
  if (!orders || !sf) { console.log('Missing tabs'); return; }
  var oData = orders.getRange(1, 1, Math.min(orders.getLastRow(), 6), orders.getLastColumn()).getValues();
  console.log('OrdersLog headers: ' + JSON.stringify(oData[0]));
  if (oData.length > 1) console.log('Sample row: ' + JSON.stringify(oData[1]));
  var sfData = sf.getRange(1, 1, Math.min(sf.getLastRow(), 18), sf.getLastColumn()).getValues();
  var sfHeaderRowIdx = -1;
  for (var r = 0; r < sfData.length; r++) {
    var rowStr = sfData[r].map(function(c) { return String(c || ''); }).join('|');
    if (rowStr.indexOf('Opportunity Name') > -1) { sfHeaderRowIdx = r; break; }
  }
  console.log('SF header row index: ' + sfHeaderRowIdx);
  if (sfHeaderRowIdx > -1) console.log('SF headers: ' + JSON.stringify(sfData[sfHeaderRowIdx]));
}