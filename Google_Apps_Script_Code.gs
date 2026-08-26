/**
 * Anu apology website response collector.
 *
 * 1. Create a Google Sheet.
 * 2. Extensions -> Apps Script.
 * 3. Replace the default code with this file.
 * 4. Deploy as a Web app:
 *      Execute as: Me
 *      Who has access: Anyone
 * 5. Copy the Web app URL into index.html as RESPONSE_ENDPOINT.
 */

const SHEET_NAME = 'Responses';

function doGet() {
  return ContentService
    .createTextOutput('Anu response collector is running.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    const body = e && e.postData && e.postData.contents
      ? JSON.parse(e.postData.contents)
      : {};

    const choice = String(body.choice || '').trim();
    const message = String(body.message || '').trim();
    const page = String(body.page || '').trim();
    const submittedAt = String(body.submittedAt || '').trim();

    // Only accept the two choices used by the website.
    if (choice !== 'yes' && choice !== 'time') {
      return jsonResponse({ ok: false, error: 'Invalid choice' });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Answer', 'Message', 'Page', 'Client submitted time']);
      sheet.setFrozenRows(1);
    }

    const answer = choice === 'yes' ? 'Yes ❤️' : 'I need time';

    sheet.appendRow([
      new Date(),
      answer,
      message,
      page,
      submittedAt
    ]);

    return jsonResponse({ ok: true });
  } catch (err) {
    console.error(err);
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
