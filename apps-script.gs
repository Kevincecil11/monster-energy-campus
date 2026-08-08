/**
 * Monster Energy campus forms -> Google Sheets
 *
 * Handles BOTH forms:
 *   1. Event requests  -> "Requests" tab
 *   2. Raid reports    -> "Raid Reports" tab + photos in Google Drive
 *
 * SETUP
 *   1. Open your Google Sheet.
 *   2. Extensions > Apps Script. Select all existing code, delete it, paste this whole file.
 *   3. Save (Ctrl+S).
 *   4. Run the function setupAll once. Approve the Google permission prompts.
 *   5. Deploy > Manage deployments > pencil icon > Version: New version > Deploy.
 *      Your /exec URL does not change.
 *
 * Optional: put an email address in NOTIFY_EMAIL below to get an alert per submission.
 */

var REQUESTS_SHEET = 'Requests';
var RAID_SHEET = 'Raid Reports';
var PHOTO_ROOT = 'Raid Report Photos';
var NOTIFY_EMAIL = '';

var REQUEST_HEADERS = [
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

var RAID_HEADERS = [
  'Submitted at',
  'Reference',
  'Event date',
  'Raid end time',
  '18h deadline',
  'Timeline',
  'Event name',
  'University / College',
  'Venue',
  'Type of activation',
  'Brand',
  'Cans out',
  'Cans in',
  'Cans sampled',
  'Cans to organisers',
  'Stock difference',
  'MAT members and hours',
  'Total hours',
  'Submitted by',
  'What worked',
  'What did not work',
  'Photo folder',
  'Photo links'
];


/* ------------------------------------------------------------------ *
 *  SETUP
 * ------------------------------------------------------------------ */

function setupAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  ensureSheet_(REQUESTS_SHEET, REQUEST_HEADERS);
  ensureSheet_(RAID_SHEET, RAID_HEADERS);
  getOrCreateFolder_(PHOTO_ROOT);

  Logger.log('Setup complete. Sheet: ' + ss.getUrl());
}


/* ------------------------------------------------------------------ *
 *  WEB APP ENTRY POINTS
 * ------------------------------------------------------------------ */

function doGet() {
  return ContentService.createTextOutput('Endpoint is live.');
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var data = JSON.parse(e.postData.contents);

    if (data.formType === 'raidReport') {
      return saveRaidReport_(data);
    }
    return saveRequest_(data);

  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    lock.releaseLock();
  }
}


/* ------------------------------------------------------------------ *
 *  FORM 1: EVENT REQUESTS
 * ------------------------------------------------------------------ */

function saveRequest_(data) {
  var sheet = ensureSheet_(REQUESTS_SHEET, REQUEST_HEADERS);

  sheet.appendRow([
    new Date(),
    data.ref || '',
    data.requestType || '',
    data.name || '',
    "'" + (data.phone || ''),
    data.email || '',
    data.venue || '',
    data.address || '',
    data.eventDate || '',
    data.crowd || '',
    data.notes || '',
    'New'
  ]);

  if (NOTIFY_EMAIL) {
    notify_('New event request: ' + (data.name || 'Unknown'), [
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
    ]);
  }

  return json_({ ok: true, ref: data.ref || '' });
}


/* ------------------------------------------------------------------ *
 *  FORM 2: RAID REPORTS
 * ------------------------------------------------------------------ */

function saveRaidReport_(data) {
  var images = data.images || [];
  if (images.length !== 10) {
    throw new Error('Exactly 10 pictures are required. Received ' + images.length + '.');
  }

  var members = readMembers_(data);
  if (!members.length) {
    throw new Error('Add at least one MAT member.');
  }

  // Upload photos into a folder named after this report.
  var folder = createReportFolder_(data);
  var links = [];
  for (var i = 0; i < images.length; i++) {
    var image = images[i];
    var blob = Utilities.newBlob(
      Utilities.base64Decode(image.data),
      image.mime || 'image/jpeg',
      safeName_(pad2_(i + 1) + '-' + (image.name || 'photo.jpg'))
    );
    links.push(folder.createFile(blob).getUrl());
  }

  var submitted = data.submittedAt ? new Date(data.submittedAt) : new Date();
  var deadline = deadlineFor_(data.reportDate, data.raidEnd);
  var timeline = (deadline && submitted > deadline) ? 'Late' : 'On time';

  var out = num_(data.cansOut);
  var back = num_(data.cansIn);
  var sampled = num_(data.cansSampled);
  var handed = num_(data.cansOrganisers);
  var difference = out - back - sampled - handed;

  var memberLines = members.map(function (m) {
    return m.name + ': ' + m.hours + ' hrs';
  }).join('\n');

  var totalHours = members.reduce(function (sum, m) {
    return sum + m.hours;
  }, 0);

  var sheet = ensureSheet_(RAID_SHEET, RAID_HEADERS);
  sheet.appendRow([
    submitted,
    data.ref || '',
    data.reportDate || '',
    data.raidEnd || '',
    deadline || '',
    timeline,
    data.eventName || '',
    data.college || '',
    data.venue || '',
    data.activation || '',
    data.brand || '',
    out,
    back,
    sampled,
    handed,
    difference,
    memberLines,
    totalHours,
    data.submittedBy || '',
    data.worked || '',
    data.didnt || '',
    folder.getUrl(),
    links.join('\n')
  ]);

  // Colour the two columns that matter at a glance.
  var row = sheet.getLastRow();
  sheet.getRange(row, 6).setBackground(timeline === 'On time' ? '#d9ead3' : '#f4cccc');
  sheet.getRange(row, 16).setBackground(difference === 0 ? '#d9ead3' : '#f4cccc');

  if (NOTIFY_EMAIL) {
    notify_('Raid report: ' + (data.eventName || 'Unknown event'), [
      'Reference: ' + (data.ref || ''),
      'Event: ' + (data.eventName || ''),
      'College: ' + (data.college || ''),
      'Date: ' + (data.reportDate || ''),
      'Timeline: ' + timeline,
      'Submitted by: ' + (data.submittedBy || ''),
      '',
      'MATs (' + totalHours + ' hrs total):',
      memberLines,
      '',
      'Stock difference: ' + difference,
      'Photos: ' + folder.getUrl()
    ]);
  }

  return json_({ ok: true, ref: data.ref || '', folder: folder.getUrl() });
}

/**
 * Accepts the new matMembers array, and still understands the older
 * mats + totalHours payload so nothing breaks mid-rollout.
 */
function readMembers_(data) {
  var members = [];

  if (data.matMembers && data.matMembers.length) {
    for (var i = 0; i < data.matMembers.length; i++) {
      var m = data.matMembers[i] || {};
      var name = String(m.name || '').trim();
      if (name) members.push({ name: name, hours: num_(m.hours) });
    }
    return members;
  }

  if (data.mats) {
    var names = String(data.mats).split('\n');
    for (var j = 0; j < names.length; j++) {
      var single = names[j].trim();
      if (single) members.push({ name: single, hours: 0 });
    }
    if (members.length && data.totalHours) {
      members[0].hours = num_(data.totalHours);
    }
  }

  return members;
}


/* ------------------------------------------------------------------ *
 *  HELPERS
 * ------------------------------------------------------------------ */

function getSpreadsheet_() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) {
    return SpreadsheetApp.openById(id);
  }
  // Fallback so submissions still save if setupAll was never run.
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) {
    throw new Error('No spreadsheet found. Run setupAll once from the Apps Script editor.');
  }
  return active;
}

function ensureSheet_(name, headers) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  var needed = headers.length;
  if (sheet.getMaxColumns() < needed) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), needed - sheet.getMaxColumns());
  }

  // Write the header row only when the sheet is empty or the header changed.
  var current = sheet.getRange(1, 1, 1, needed).getValues()[0];
  if (current.join('|') !== headers.join('|')) {
    sheet.getRange(1, 1, 1, needed).setValues([headers]);
    sheet.getRange(1, 1, 1, needed)
      .setFontWeight('bold')
      .setBackground('#111111')
      .setFontColor('#8CFF33');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function deadlineFor_(dateText, timeText) {
  if (!dateText || !timeText) return null;
  var parts = String(dateText).split('-');
  var clock = String(timeText).split(':');
  if (parts.length < 3 || clock.length < 2) return null;

  var ended = new Date(
    Number(parts[0]),
    Number(parts[1]) - 1,
    Number(parts[2]),
    Number(clock[0]),
    Number(clock[1])
  );
  if (isNaN(ended.getTime())) return null;

  return new Date(ended.getTime() + 18 * 60 * 60 * 1000);
}

function getOrCreateFolder_(name) {
  var found = DriveApp.getFoldersByName(name);
  return found.hasNext() ? found.next() : DriveApp.createFolder(name);
}

function createReportFolder_(data) {
  var root = getOrCreateFolder_(PHOTO_ROOT);
  var label = [data.reportDate, data.eventName, data.ref].filter(String).join(' - ');
  return root.createFolder(safeName_(label || 'Raid report'));
}

function safeName_(text) {
  return String(text || 'file').replace(/[\\\/:*?"<>|]/g, '-').slice(0, 120);
}

function pad2_(n) {
  return n < 10 ? '0' + n : String(n);
}

function num_(value) {
  var n = Number(value);
  return isNaN(n) ? 0 : n;
}

function notify_(subject, lines) {
  try {
    MailApp.sendEmail(NOTIFY_EMAIL, subject, lines.join('\n'));
  } catch (err) {
    Logger.log('Email failed: ' + err);
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/* ------------------------------------------------------------------ *
 *  TESTS (run these from the editor)
 * ------------------------------------------------------------------ */

function testRequestWrite() {
  saveRequest_({
    ref: 'TEST-REQUEST',
    requestType: 'Collab',
    name: 'Test Person',
    phone: '9999999999',
    email: 'test@example.com',
    venue: 'Test venue',
    address: 'Test address',
    eventDate: '2026-09-01',
    crowd: '500',
    notes: 'Test row'
  });
  Logger.log('Wrote a test row to ' + REQUESTS_SHEET);
}

function testRaidSheet() {
  var sheet = ensureSheet_(RAID_SHEET, RAID_HEADERS);
  Logger.log(RAID_SHEET + ' is ready with ' + sheet.getLastColumn() + ' columns.');
}
