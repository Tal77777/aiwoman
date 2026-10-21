// Google Apps Script: receives sign-ups from the landing page and adds a row to this sheet.
// Setup steps are in README.md.

const SHEET_NAME = 'הרשמות';
const HEADERS = ['תאריך', 'שם', 'טלפון', 'מייל', 'איפה עם AI', 'עדכונים', 'מקור'];

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.website) return reply({ ok: true }); // honeypot filled: a bot

    if (!d.name || !d.phone || !d.email) return reply({ ok: false, error: 'missing fields' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
      sheet.setRightToLeft(true);
    }

    sheet.appendRow([
      new Date(),
      clean(d.name),
      clean(d.phone),
      clean(d.email),
      clean(d.level),
      clean(d.updates),
      clean(d.source),
    ]);
    return reply({ ok: true });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  }
}

// Stops a value like "=HYPERLINK(...)" from running as a formula in the sheet,
// and keeps the leading 0 of phone numbers.
function clean(v) {
  const s = String(v || '').slice(0, 300);
  return /^[=+\-@0]/.test(s) ? "'" + s : s;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
