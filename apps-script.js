// ==== הדביקי את הקוד הזה ב-Google Apps Script ====
// Extensions > Apps Script > הדביקי > Deploy > New Deployment > Web App
// Execute as: Me, Who has access: Anyone
// ** חשוב: אחרי עדכון, צריכה לעשות Deploy חדש (New deployment) **

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
  const params = e && e.parameter ? e.parameter : {};
  const month = params.month || getCurrentMonth();
  const action = params.action || '';
  const name = params.name || '';

  if (action === 'add' && name) {
    const data = getData(month);
    setCount(month, name, (data[name] || 0) + 1);
  } else if (action === 'remove' && name) {
    const data = getData(month);
    const current = data[name] || 0;
    setCount(month, name, Math.max(0, current - 1));
  } else if (action === 'reset') {
    const data = getData(month);
    Object.keys(data).forEach(n => setCount(month, n, 0));
  }

  const updated = getData(month);
  return ContentService.createTextOutput(JSON.stringify(updated))
    .setMimeType(ContentService.MimeType.JSON);
}
