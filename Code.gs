// ============================================================
// SI-DEMANG KREBET — Google Apps Script Backend
// REST API for Village Demographic & UMKM Mapping
// ============================================================

const SHEET_NAME = 'Data';
const COLUMNS = ['Timestamp', 'NIK', 'Nama_Lengkap', 'RT', 'Usia', 'Pekerjaan', 'Jenis_UMKM', 'Latitude', 'Longitude'];

function doGet(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return buildJsonResponse({ success: false, error: 'Sheet not found' });

    const lastRow = sheet.getLastRow();
    if (lastRow < 1) return buildJsonResponse({ success: true, data: [], count: 0 });

    const data = sheet.getDataRange().getValues();

    // Detect if first row is a header row by checking if first cell looks like a date or "Timestamp"
    let startIndex = 0;
    const firstCell = String(data[0][0]).trim().toLowerCase();
    if (firstCell === 'timestamp' || firstCell === 'Timestamp') {
      startIndex = 1; // Skip header row
    }

    const rows = data.slice(startIndex);
    const result = rows.map(row => {
      const obj = {};
      COLUMNS.forEach((col, i) => {
        let val = row[i];
        // Convert Date objects to ISO string
        if (val instanceof Date) {
          val = val.toISOString();
        }
        // Ensure Latitude/Longitude are proper numbers
        if (col === 'Latitude' || col === 'Longitude') {
          val = parseFloat(String(val).replace(',', '.'));
          if (isNaN(val)) val = 0;
        }
        // Ensure Usia is a number
        if (col === 'Usia') {
          val = parseInt(String(val).replace(',', '.'), 10);
          if (isNaN(val)) val = 0;
        }
        obj[col] = val;
      });
      return obj;
    });

    return buildJsonResponse({ success: true, data: result, count: result.length });
  } catch (error) {
    return buildJsonResponse({ success: false, error: error.message });
  }
}

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return buildJsonResponse({ success: false, error: 'Sheet not found' });

    const body = JSON.parse(e.postData.contents);
    const required = ['NIK', 'Nama_Lengkap', 'RT', 'Usia', 'Pekerjaan', 'Latitude', 'Longitude'];
    for (const f of required) {
      if (!body[f] && body[f] !== 0) return buildJsonResponse({ success: false, error: 'Field "' + f + '" wajib diisi.' });
    }

    // Parse coordinates properly (handle comma decimal separator)
    const lat = parseFloat(String(body.Latitude).replace(',', '.'));
    const lng = parseFloat(String(body.Longitude).replace(',', '.'));

    if (isNaN(lat) || isNaN(lng)) {
      return buildJsonResponse({ success: false, error: 'Koordinat tidak valid.' });
    }

    sheet.appendRow([
      new Date(),
      String(body.NIK),
      body.Nama_Lengkap,
      String(body.RT),
      Number(body.Usia),
      body.Pekerjaan,
      body.Jenis_UMKM || '-',
      lat,
      lng
    ]);

    return buildJsonResponse({ success: true, message: 'Data berhasil disimpan!' });
  } catch (error) {
    return buildJsonResponse({ success: false, error: error.message });
  }
}

function buildJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function initializeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  // Insert header row at row 1 (push existing data down)
  const lastRow = sheet.getLastRow();
  if (lastRow > 0) {
    // Check if headers already exist
    const firstCell = String(sheet.getRange(1, 1).getValue()).trim().toLowerCase();
    if (firstCell !== 'timestamp') {
      sheet.insertRowBefore(1);
    }
  }

  sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]).setFontWeight('bold').setBackground('#065f46').setFontColor('#fff');
  COLUMNS.forEach((_, i) => sheet.autoResizeColumn(i + 1));
  SpreadsheetApp.getUi().alert('Sheet "Data" berhasil diinisialisasi! Header sudah ditambahkan.');
}
