/**
 * Nexora — Contact Form Backend (Google Apps Script)
 * -----------------------------------------------------------------
 * Handles both GET and POST requests, reads name/email/phone from
 * query params (or POST body/params), appends a row
 * (Timestamp, Name, Email, Phone) to a sheet tab named "Responses"
 * (auto-created with bold headers if it doesn't exist), and returns
 * a JSON response: { "success": true } or { "success": false, "error": "..." }
 *
 * DEPLOYMENT:
 * 1. Open (or create) a Google Sheet.
 * 2. Extensions > Apps Script, paste this file in as Code.gs.
 * 3. Deploy > New deployment > Select type "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Copy the deployed Web App URL — you'll need it for the HTML page.
 */

var SHEET_NAME = "Responses";
var HEADERS = ["Timestamp", "Name", "Email", "Phone"];

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    var params = extractParams(e);

    var name = (params.name || "").toString().trim();
    var email = (params.email || "").toString().trim();
    var phone = (params.phone || "").toString().trim();

    if (!name || !email || !phone) {
      return jsonResponse({
        success: false,
        error: "Missing required field(s): name, email, phone are all required."
      });
    }

    var sheet = getOrCreateResponsesSheet();

    sheet.appendRow([new Date(), name, email, phone]);

    return jsonResponse({ success: true });
  } catch (err) {
    return jsonResponse({ success: false, error: err && err.message ? err.message : String(err) });
  }
}

/**
 * Pulls parameters from either query string (e.parameter, used for GET
 * and for POST requests sent as x-www-form-urlencoded/no-cors) or from
 * a JSON POST body, whichever is present.
 */
function extractParams(e) {
  var params = {};

  if (e && e.parameter) {
    for (var key in e.parameter) {
      params[key] = e.parameter[key];
    }
  }

  if (e && e.postData && e.postData.contents) {
    try {
      var body = JSON.parse(e.postData.contents);
      for (var k in body) {
        params[k] = body[k];
      }
    } catch (parseErr) {
      // Not JSON — ignore; e.parameter (if any) already covers form-encoded data.
    }
  }

  return params;
}

function getOrCreateResponsesSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setValues([HEADERS]);
    headerRange.setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, HEADERS.length);
  }

  return sheet;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
