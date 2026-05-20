/**
 * Launchpad Waitlist — Google Apps Script backend
 *
 * Setup:
 * 1. Create a new Google Sheet with two tabs named exactly: "Founders" and "Talent"
 * 2. In the Sheet: Extensions → Apps Script. Paste this whole file in.
 * 3. Save, then Deploy → New deployment → Type: Web app.
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Authorize when prompted.
 * 5. Copy the deployed Web app URL and paste it into index.html as WAITLIST_ENDPOINT.
 */

const FOUNDER_HEADERS = ['Timestamp', 'Name', 'Email', 'Startup', 'Industry', 'Stage'];
const TALENT_HEADERS  = ['Timestamp', 'Name', 'Email', 'School', 'Skill', 'Looking For'];

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const type = body.type;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let sheet, row;

    if (type === 'founder') {
      sheet = ensureSheet(ss, 'Founders', FOUNDER_HEADERS);
      row = [
        new Date(),
        body.name || '',
        body.email || '',
        body.startup || '',
        body.industry || '',
        body.stage || '',
      ];
    } else if (type === 'talent') {
      sheet = ensureSheet(ss, 'Talent', TALENT_HEADERS);
      row = [
        new Date(),
        body.name || '',
        body.email || '',
        body.school || '',
        body.skill || '',
        body.looking || '',
      ];
    } else {
      return jsonOut({ ok: false, error: 'Unknown type' });
    }

    sheet.appendRow(row);
    return jsonOut({ ok: true });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doGet() {
  return jsonOut({ ok: true, service: 'launchpad-waitlist' });
}

function ensureSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
