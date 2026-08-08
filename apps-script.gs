/**
 * Monster truck request form -> Google Sheets
 *
 * SETUP
 * 1. Create a Google Sheet. Copy its ID from the URL:
 *    https://docs.google.com/spreadsheets/d/<THIS_PART>/edit
 * 2. Paste that ID into SHEET_ID below.
 * 3. In the Sheet: Extensions > Apps Script. Delete the sample code,
 *    paste this whole file, save.
 * 4. Deploy > New deployment > type "Web app".
 *      Execute as: Me
 *      Who has access: Anyone
 * 5. Authorize when prompted. Copy the /exec URL it gives you.
 * 6. Paste that URL into SHEET_ENDPOINT in index.html.
 *
 * Optional: set NOTIFY_EMAIL to get an email on every new request.
 */

var SHEET_ID = 'PASTE_YOUR_SHEET_ID_HERE';
var SHEET_NAME = 'Requests';
var NOTIFY_EMAIL = ''; // e.g. 'you@gmail.com' - leave empty to disable

var HEADERS = [
  'Received at',
  'Reference',
  'Request type',
  'Name',
  'Phone',
  'Email',
  'Venue',
  'Address',
  'Event date',
  'Expected crowd',
  'Notes',
  'Status'
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getSheet_();

    sheet.appendRow([
      new Date(),
      data.ref || '',
      data.requestType || '',
      data.name || '',
      "'" + (data.phone || ''),   // leading quote keeps the number as text
      data.email || '',
      data.venue || '',
      data.address || '',
      data.eventDate || '',
      data.crowd || '',
      data.notes || '',
      'New'
    ]);

    if (NOTIFY_EMAIL) {
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        subject: 'New truck request: ' + (data.name || 'Unknown') + ' (' + (data.requestType || '') + ')',
        body: [
          'Reference: ' + (data.ref || ''),
          'Type: ' + (data.requestType || ''),
          'Name: ' + (data.name || ''),
          'Phone: ' + (data.phone || ''),
          'Email: ' + (data.email || ''),
          'Venue: ' + (data.venue || ''),
          'Address: ' + (data.address || ''),
          'Event date: ' + (data.eventDate || ''),
          'Expected crowd: ' + (data.crowd || ''),
          'Notes: ' + (data.notes || '')
        ].join('\n')
      });
    }

    return json_({ ok: true, ref: data.ref || '' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('Monster request endpoint is live.');
}

function getSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#111111')
      .setFontColor('#8CFF33');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(8, 280); // address
    sheet.setColumnWidth(11, 280); // notes
  }
  return sheet;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
