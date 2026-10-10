// Onboarding Form to Google Sheet bridge.
// Paste this whole file in your Google Sheet: Extensions > Apps Script.
// Save, then: Deploy > New deployment > Web app
//   Execute as: Me  |  Who has access: Anyone
// Copy the Web app URL and paste it in index.html (SCRIPT_URL).
const SHEET_NAME = 'Responses';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    const data = (e && e.parameter) ? e.parameter : {};

    let headers = sheet.getLastColumn() > 0
      ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].filter(String)
      : [];
    if (headers.length === 0) headers = ['Submitted at'];
    Object.keys(data).forEach(function (key) {
      if (headers.indexOf(key) === -1) headers.push(key);
    });

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);

    const row = headers.map(function (h) {
      return h === 'Submitted at' ? new Date() : (data[h] || '');
    });
    sheet.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Opening the Web app URL in a browser shows this, so you know it's live.
function doGet() {
  return ContentService.createTextOutput('Onboarding form bridge is live.');
}
