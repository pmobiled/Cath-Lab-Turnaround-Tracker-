/**
 * Cath Lab Turnaround Tracker — Google Sheets receiver
 *
 * Setup (5 minutes):
 *  1. Open the Google Sheet you want the data in.
 *  2. Extensions → Apps Script. Delete whatever is in Code.gs and paste this file.
 *  3. Save. Then Deploy → New deployment → gear icon → Web app.
 *       Execute as:       Me
 *       Who has access:   Anyone
 *  4. Authorize when prompted (you'll get an "unverified app" screen — Advanced →
 *     Go to project → Allow. It's your own script talking to your own sheet).
 *  5. Copy the /exec URL it gives you and paste it into the app's Setup screen.
 *
 * Re-deploy note: after editing this file, use Deploy → Manage deployments →
 * pencil → Version: New version. Editing without re-deploying does nothing.
 *
 * NOTE: This file is the version-controlled copy for history/recovery. The
 * copy that actually runs lives inside the bound Google Sheet's Apps Script
 * editor (Extensions → Apps Script). Keep them in sync manually — pushing to
 * this repo does not redeploy the live script.
 */
var SHEET_NAME = 'Runs';

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var body = JSON.parse(e.postData.contents);
    var columns = body.columns || [];
    var rows = body.rows || (body.row ? [body.row] : []);
    var sheet = getSheet_(columns);

    // Ping / connection test sends an empty rows array.
    if (!rows.length) {
      return json_({ ok: true, added: 0, message: 'Connected' });
    }

    var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var existing = getExistingIds_(sheet, header);
    var toAppend = [];
    var skipped = 0;

    rows.forEach(function (row) {
      if (row.run_id && existing[row.run_id]) { skipped++; return; }   // retry-safe
      if (row.run_id) existing[row.run_id] = true;
      toAppend.push(header.map(function (col) {
        var v = row[col];
        return (v === undefined || v === null) ? '' : v;
      }));
    });

    if (toAppend.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, toAppend.length, header.length).setValues(toAppend);
    }

    return json_({ ok: true, added: toAppend.length, skipped: skipped });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function doGet() {
  return json_({ ok: true, message: 'Turnaround Tracker receiver is live. Post runs here.' });
}

function getSheet_(columns) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    if (!columns.length) throw new Error('No header to write yet — send a run first.');
    sheet.getRange(1, 1, 1, columns.length).setValues([columns]);
    sheet.getRange(1, 1, 1, columns.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getExistingIds_(sheet, header) {
  var idCol = header.indexOf('run_id') + 1;
  var seen = {};
  if (idCol < 1 || sheet.getLastRow() < 2) return seen;
  sheet.getRange(2, idCol, sheet.getLastRow() - 1, 1).getValues().forEach(function (r) {
    if (r[0]) seen[String(r[0])] = true;
  });
  return seen;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
