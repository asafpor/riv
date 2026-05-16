// ==== הדביקי את הקוד הזה ב-Google Apps Script ====
// Extensions > Apps Script > הדביקי > Deploy > New Deployment > Web App
// Execute as: Me, Who has access: Anyone

const KIDS = ['דנ', 'אב', 'יר', 'ג'];

function getCurrentMonth() {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
}

function getOrCreateSheet(month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(month);
  if (!sheet) {
    sheet = ss.insertSheet(month);
    sheet.appendRow(['name', 'count']);
    KIDS.forEach(kid => sheet.appendRow([kid, 0]));
  }
  return sheet;
}

function getData(month) {
  const sheet = getOrCreateSheet(month);
  const rows = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < rows.length; i++) {
    result[rows[i][0]] = rows[i][1];
  }
  return result;
}

function setCount(month, name, count) {
  const sheet = getOrCreateSheet(month);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === name) {
      sheet.getRange(i + 1, 2).setValue(count);
      return;
    }
  }
}

function doGet(e) {
  const month = (e && e.parameter && e.parameter.month) || getCurrentMonth();
  const data = getData(month);
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const month = body.month || getCurrentMonth();

  if (body.action === 'add') {
    const data = getData(month);
    setCount(month, body.name, (data[body.name] || 0) + 1);
  } else if (body.action === 'remove') {
    const data = getData(month);
    const current = data[body.name] || 0;
    setCount(month, body.name, Math.max(0, current - 1));
  } else if (body.action === 'reset') {
    const data = getData(month);
    Object.keys(data).forEach(name => setCount(month, name, 0));
  }

  const updated = getData(month);
  return ContentService.createTextOutput(JSON.stringify(updated))
    .setMimeType(ContentService.MimeType.JSON);
}
