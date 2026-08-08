// Campus event request form -> Google Sheets
// Nothing here needs editing. Paste this into the Apps Script editor
// attached to your Sheet (Extensions > Apps Script) and deploy.
// Want an email per request? Put your address between the quotes below.

var SHEET_NAME = "Requests";
var NOTIFY_EMAIL = "";

var HEADERS = [
  "Received at",
  "Reference",
  "Request type",
  "Name",
  "Phone",
  "Email",
  "Venue",
  "Address",
  "Event date",
  "Expected crowd",
  "Notes",
  "Status"
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getSheet_();

    sheet.appendRow([
      new Date(),
      data.ref || "",
      data.requestType || "",
      data.name || "",
      "'" + (data.phone || ""),
      data.email || "",
      data.venue || "",
      data.address || "",
      data.eventDate || "",
      data.crowd || "",
      data.notes || "",
      "New"
    ]);

    if (NOTIFY_EMAIL) {
      var lines = [
        "Reference: " + (data.ref || ""),
        "Type: " + (data.requestType || ""),
        "Name: " + (data.name || ""),
        "Phone: " + (data.phone || ""),
        "Email: " + (data.email || ""),
        "Venue: " + (data.venue || ""),
        "Address: " + (data.address || ""),
        "Event date: " + (data.eventDate || ""),
        "Expected crowd: " + (data.crowd || ""),
        "Notes: " + (data.notes || "")
      ];
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        subject: "New request: " + (data.name || "Unknown"),
        body: lines.join("\n")
      });
    }

    return json_({ ok: true, ref: data.ref || "" });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput("Request endpoint is live.");
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight("bold")
      .setBackground("#111111")
      .setFontColor("#8CFF33");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(8, 280);
    sheet.setColumnWidth(11, 280);
  }
  return sheet;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Run this once from the editor to drop a test row into the Sheet.
function testWrite() {
  getSheet_().appendRow([new Date(), "TEST-0001", "Collab", "Test Person",
    "'9999999999", "test@example.com", "Test venue", "Test address",
    "2026-09-01", "500", "Test row", "New"]);
}
