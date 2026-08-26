const SPREADSHEET_ID = '1LyOmp2J7-DqBC283R8cudrp0EwWHWo9rToMu48z_0NM';
const SHEET_NAME = 'Responses';

function doGet() {
  return ContentService
    .createTextOutput('Anu response collector is running.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No POST data received');
    }

    const body = JSON.parse(e.postData.contents);

    const choice = String(body.choice || '').trim();
    const message = String(body.message || '').trim();
    const page = String(body.page || '').trim();
    const submittedAt = String(body.submittedAt || '').trim();

    if (choice !== 'yes' && choice !== 'time') {
      throw new Error('Invalid choice: ' + choice);
    }

    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);

    let sheet = spreadsheet.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAME);
    }

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Timestamp',
        'Answer',
        'Message',
        'Page',
        'Client submitted time'
      ]);
    }

    const answer = choice === 'yes'
      ? 'Yes ❤️'
      : 'I need time';

    sheet.appendRow([
      new Date(),
      answer,
      message,
      page,
      submittedAt
    ]);

    SpreadsheetApp.flush();

    return jsonResponse({
      ok: true,
      answer: answer
    });

  } catch (error) {

    console.error(error);

    return jsonResponse({
      ok: false,
      error: String(error)
    });
  }
}


function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}


function testResponse() {

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);

  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  sheet.appendRow([
    new Date(),
    'TEST ❤️',
    'TEST RESPONSE FROM APPS SCRIPT',
    'TEST',
    new Date().toISOString()
  ]);

  SpreadsheetApp.flush();

  console.log('TEST ROW WRITTEN SUCCESSFULLY');
}
