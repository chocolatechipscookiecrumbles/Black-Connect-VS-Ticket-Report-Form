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
// ─────────────────────────────────────────────────────────────


// Called by Google when the form POSTs data.
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);

    // ── Token check (skip if both tokens are empty strings) ──
    if (SECRET_TOKEN !== '' && payload.token !== SECRET_TOKEN) {
      return respond({ status: 'error', message: 'Unauthorized' }, 401);
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
