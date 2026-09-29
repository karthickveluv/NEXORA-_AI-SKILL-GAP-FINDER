/**
 * Form -> Google Sheets backend
 * Handles both doGet(e) and doPost(e).
 * Reads name, email, phone from query params and appends a row to the
 * "Responses" tab (Timestamp, Name, Email, Phone).
 *
 * SETUP
 * 1. Open your Google Sheet -> Extensions -> Apps Script, paste this file.
 * 2. Deploy -> New deployment -> type: Web app
 *      Execute as: Me
 *      Who has access: Anyone
 * 3. Copy the Web app URL (ends with /exec).
 *
 * If the script is NOT bound to a sheet (created from script.google.com),
 * put your spreadsheet ID in SPREADSHEET_ID below.
 */

var SHEET_NAME = "Responses";
var SPREADSHEET_ID = ""; // optional: only needed for standalone scripts
var HEADERS = ["Timestamp", "Name", "Email", "Phone"];

function doGet(e) {
  return handleRequest_(e);
}

function doPost(e) {
  return handleRequest_(e);
}

function handleRequest_(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var params = (e && e.parameter) ? e.parameter : {};
    var name = String(params.name || "").trim();
    var email = String(params.email || "").trim();
    var phone = String(params.phone || "").trim();

    if (!name && !email && !phone) {
      return json_({ success: false, error: "No data received" });
    }

    var sheet = getOrCreateSheet_();
    sheet.appendRow([new Date(), name, email, phone]);

    return json_({ success: true });
  } catch (err) {
    return json_({ success: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function getOrCreateSheet_() {
  var ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error("No spreadsheet found. Bind the script to a sheet or set SPREADSHEET_ID.");
  }

  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setValues([HEADERS])
      .setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
