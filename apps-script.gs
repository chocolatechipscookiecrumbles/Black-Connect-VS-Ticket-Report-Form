// ============================================================
//  ISSUE REPORT — Google Apps Script (apps-script.gs)
//  Paste this entire file into script.google.com
// ============================================================
//
//  SETUP STEPS (takes ~10 minutes):
//
//  1. Open your Google Sheet.
//  2. Click Extensions → Apps Script.
//  3. Delete everything in the editor and paste this whole file.
//  4. Fill in SHEET_NAME and SECRET_TOKEN below.
//  5. Click Deploy → New deployment.
//       - Type: Web app
//       - Execute as: Me
//       - Who has access: Anyone   ← lets the form POST without login
//  6. Click Deploy, then copy the Web App URL.
//  7. Paste that URL into issue-report-form.html where it says
//     APPS_SCRIPT_URL, and paste the same SECRET_TOKEN.
//  8. Host issue-report-form.html on GitHub Pages or Netlify.
//
//  RE-DEPLOYING AFTER EDITS:
//  Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy.
//  The URL stays the same.
// ============================================================

// ── CONFIG ────────────────────────────────────────────────────
var SHEET_NAME   = 'Reports';          // Tab name in your Sheet
var SECRET_TOKEN = 'YOUR_SECRET_TOKEN_HERE'; // Must match the token in the HTML form
                                             // Set both to '' to disable auth
// External ticket source (optional)
// If SOURCE_SPREADSHEET_ID is set, the server will validate incoming
// submissions against the ticket IDs read from that spreadsheet.
var SOURCE_SPREADSHEET_ID = 'YOUR_SOURCE_SPREADSHEET_ID_HERE'; // Spreadsheet ID (not URL)
var SOURCE_SHEET_NAME = 'Tickets'; // Tab name in the source spreadsheet
var TICKET_COLUMN = 1; // 1 = column A
// ─────────────────────────────────────────────────────────────


// Called by Google when the form POSTs data.
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);

    // ── Token check (skip if both tokens are empty strings) ──
    if (SECRET_TOKEN !== '' && payload.token !== SECRET_TOKEN) {
      return respond({ status: 'error', message: 'Unauthorized' }, 401);
    }

    // ── Optional: validate ticketId against external ticket list ──
    if (SOURCE_SPREADSHEET_ID && payload.ticketId) {
      var allowed = getTicketListFromExternalSpreadsheet();
      if (allowed.length && allowed.indexOf(payload.ticketId) === -1) {
        return respond({ status: 'error', message: 'Invalid ticketId' }, 400);
      }
    }

    var sheet = getOrCreateSheet(SHEET_NAME);

    // Write one row per feature tested, with ticket metadata on every row
    // so you can filter/pivot by ticket in Sheets easily.
    var rowsWritten = 0;
    var featureRows = payload.rows || [];

    featureRows.forEach(function(r) {
      sheet.appendRow([
        payload.ticketId   || '',
        payload.testerName || '',
        payload.moduleName || '',
        payload.testDate   || '',
        r.feature          || '',
        r.status           || 'Untested',
        r.notes            || '',
        new Date()           // timestamp of submission
      ]);
      rowsWritten++;
    });

    return respond({ status: 'ok', rowsWritten: rowsWritten });

  } catch (err) {
    return respond({ status: 'error', message: err.message }, 500);
  }
}


// ── Helpers ───────────────────────────────────────────────────

// Returns the named sheet, creating it with headers if it doesn't exist.
function getOrCreateSheet(name) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
    // Write column headers on first use
    sheet.appendRow([
      'Ticket ID',
      'Tester',
      'Module',
      'Date',
      'Feature',
      'Pass / Fail',
      'Notes',
      'Submitted at'
    ]);

    // Style the header row
    var header = sheet.getRange(1, 1, 1, 8);
    header.setFontWeight('bold');
    header.setBackground('#f3f3f3');
    sheet.setFrozenRows(1);

    // Set column widths for readability
    sheet.setColumnWidth(1, 110); // Ticket ID
    sheet.setColumnWidth(2, 120); // Tester
    sheet.setColumnWidth(3, 140); // Module
    sheet.setColumnWidth(4, 100); // Date
    sheet.setColumnWidth(5, 160); // Feature
    sheet.setColumnWidth(6, 90);  // Pass/Fail
    sheet.setColumnWidth(7, 280); // Notes
    sheet.setColumnWidth(8, 150); // Submitted at
  }

  return sheet;
}

// Wraps JSON response with CORS headers so the browser form can receive it.
function respond(data, statusCode) {
  var output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}


// Returns ticket IDs from the external spreadsheet (simple, non-cached)
function getTicketListFromExternalSpreadsheet() {
  if (!SOURCE_SPREADSHEET_ID) return [];
  try {
    var other = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
    var sheet = other.getSheetByName(SOURCE_SHEET_NAME);
    if (!sheet) return [];
    var last = sheet.getLastRow();
    if (last < 2) return [];
    var vals = sheet.getRange(2, TICKET_COLUMN, last - 1, 1).getValues();
    var out = [];
    var seen = {};
    for (var i = 0; i < vals.length; i++) {
      var v = (vals[i][0] || '').toString().trim();
      if (!v) continue;
      if (!seen[v]) { seen[v] = true; out.push(v); }
    }
    out.sort();
    return out;
  } catch (err) {
    return [];
  }
}

// Simple GET endpoint to return the ticket list as JSON.
// Call the web app URL with `?action=tickets` to receive { tickets: [...] }
function doGet(e) {
  try {
    var action = e && e.parameter && e.parameter.action;
    if (action === 'tickets') {
      var list = getTicketListFromExternalSpreadsheet();
      var output = ContentService.createTextOutput(JSON.stringify({ tickets: list })).setMimeType(ContentService.MimeType.JSON);
      return output;
    }
    return respond({ status: 'ok' });
  } catch (err) {
    return respond({ status: 'error', message: err.message }, 500);
  }
}


// ── Optional: auto-color Pass/Fail cells ──────────────────────
// This trigger runs every time a new row is added to the sheet
// and colors the Pass/Fail column green or red automatically.
//
// To enable:
//   Triggers (clock icon) → Add trigger → onEdit, From spreadsheet, On edit.

function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== SHEET_NAME) return;

  var range  = e.range;
  var col    = range.getColumn();
  var pfCol  = 6; // Column F = Pass / Fail

  if (col !== pfCol) return;

  var val = range.getValue().toString().toLowerCase();
  if (val === 'pass') {
    range.setBackground('#c6efce');
    range.setFontColor('#276221');
  } else if (val === 'fail') {
    range.setBackground('#ffc7ce');
    range.setFontColor('#9c0006');
  } else {
    range.setBackground(null);
    range.setFontColor(null);
  }
}
