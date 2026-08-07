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

    var header = syncHeader_(sheet, columns);
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

// Serves every row in Runs as JSON for the analysis dashboard (analysis.html).
// Supports JSONP via ?callback=name so the dashboard can fetch cross-origin
// from GitHub Pages without hitting CORS.
function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var columns = [], rows = [];
  if (sheet && sheet.getLastRow() > 1) {
    var values = sheet.getDataRange().getValues();
    columns = values[0];
    rows = values.slice(1).map(function (r) {
      var obj = {};
      columns.forEach(function (col, i) { obj[col] = normalizeValue_(col, r[i]); });
      return obj;
    });
  }
  var payload = { ok: true, columns: columns, rows: rows };
  var callback = e && e.parameter && e.parameter.callback;
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(payload) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(payload);
}

// Sheets auto-converts well-formatted text ("TRUE", date-like strings) into
// native booleans/dates on write, so reads need to normalize back to the
// plain strings/numbers buildRow() originally sent.
function normalizeValue_(col, val) {
  if (val instanceof Date) {
    var tz = Session.getScriptTimeZone();
    if (col === 'date') return Utilities.formatDate(val, tz, 'yyyy-MM-dd');
    if (col === 'depart_time') return Utilities.formatDate(val, tz, 'HH:mm');
    return val.toISOString();
  }
  if (typeof val === 'boolean') return val ? 'TRUE' : 'FALSE';
  return val;
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

// Appends any columns the app knows about but the sheet doesn't, so a schema
// addition doesn't silently drop data. Only ever appends — existing columns
// keep their positions, so old rows stay aligned.
function syncHeader_(sheet, columns) {
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var missing = columns.filter(function (col) { return header.indexOf(col) === -1; });
  if (missing.length) {
    sheet.getRange(1, header.length + 1, 1, missing.length).setValues([missing]);
    sheet.getRange(1, header.length + 1, 1, missing.length).setFontWeight('bold');
    header = header.concat(missing);
  }
  return header;
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
