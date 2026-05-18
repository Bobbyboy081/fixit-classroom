// FixIt Classroom — Google Apps Script Backend
// วิธีใช้:
//   1. เปิด Google Sheets → Extensions → Apps Script → วางโค้ดนี้
//   2. Deploy → New deployment → Web app
//      Execute as: Me | Who has access: Anyone
//   3. Copy the Web app URL ไปวางในหน้าเว็บ (ช่อง Apps Script URL)

const SHEET_NAME = 'Repairs';

function doGet() {
  try {
    const sheet = getSheet();
    const rows  = sheet.getDataRange().getValues();
    if (rows.length <= 1) return json([]);

    const headers = rows[0];
    const data    = rows.slice(1).reverse().map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i] instanceof Date ? row[i].toISOString() : row[i]; });
      return obj;
    });
    return json(data);
  } catch (err) {
    return json({ error: err.message });
  }
}

function doPost(e) {
  try {
    const p     = JSON.parse(e.postData.contents);
    const sheet = getSheet();
    sheet.appendRow([
      Utilities.getUuid(),
      new Date().toISOString(),
      p.equipment || '',
      p.room      || '',
      p.reporter  || '',
      p.note      || '',
      p.imageUrl  || '',
      'รอดำเนินการ'
    ]);
    return json({ ok: true });
  } catch (err) {
    return json({ error: err.message });
  }
}

function getSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id', 'date', 'equipment', 'room', 'reporter', 'note', 'imageUrl', 'status']);
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, 8, 160);
  }
  return sheet;
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
