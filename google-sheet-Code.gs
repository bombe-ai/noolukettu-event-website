/**
 * Noolukettu RSVP -> Google Sheet
 * Paste into Extensions > Apps Script of the sheet that should collect replies,
 * then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 */
const SHEET_NAME = 'RSVP';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Time', 'Name', 'Attending', 'Guests', 'Note']);
      sheet.setFrozenRows(1);
    }
    const yes = d.attending === 'yes';
    const guests = yes ? Math.max(1, Math.min(15, Number(d.guests) || 1)) : 0;
    sheet.appendRow([new Date(), clean(d.name, 80), yes ? 'Yes' : 'No', guests, clean(d.note, 500)]);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json({ ok: true, message: 'RSVP endpoint is running' });
}

// Trim, cap the length, and stop a reply from being read as a spreadsheet formula.
function clean(value, max) {
  let s = String(value == null ? '' : value).trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
