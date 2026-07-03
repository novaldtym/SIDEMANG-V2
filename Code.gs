// ============================================================
// SI-DEMANG KREBET — Google Apps Script Backend
// REST API for Village Demographic & UMKM Mapping
// ============================================================

const SHEET_NAME = 'Data';
const COLUMNS = [
  'Timestamp',
  'NIK',
  'Nama_Lengkap',
  'Jenis_Kelamin',
  'RT',
  'Usia',
  'Pendidikan',
  'Pekerjaan',
  'Status_Nikah',
  'Jenis_UMKM',
  'Pembatik',
  'Jumlah_Pembatik',
  'Bantuan',
  'Latitude',
  'Longitude'
];

function doGet(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return buildJsonResponse({ success: false, error: 'Sheet not found' });

    const lastRow = sheet.getLastRow();
    if (lastRow < 1) return buildJsonResponse({ success: true, data: [], count: 0 });

    const data = sheet.getDataRange().getValues();

    let startIndex = 0;
    const firstCell = String(data[0][0]).trim().toLowerCase();
    if (firstCell === 'timestamp') {
      startIndex = 1;
    }

    const rows = data.slice(startIndex);
    const result = rows.map(row => {
      const obj = {};
      COLUMNS.forEach((col, i) => {
        let val = row[i];
        if (val instanceof Date) {
          val = val.toISOString();
        }
        if (col === 'Latitude' || col === 'Longitude') {
          val = parseFloat(String(val).replace(',', '.'));
          if (isNaN(val)) val = 0;
        }
        if (col === 'Usia') {
          val = parseInt(String(val).replace(',', '.'), 10);
          if (isNaN(val)) val = 0;
        }
        if (col === 'Jumlah_Pembatik') {
          val = parseInt(String(val), 10);
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
    const required = ['NIK', 'Nama_Lengkap', 'Jenis_Kelamin', 'RT', 'Usia', 'Pekerjaan', 'Latitude', 'Longitude'];
    for (const f of required) {
      if (!body[f] && body[f] !== 0) return buildJsonResponse({ success: false, error: 'Field "' + f + '" wajib diisi.' });
    }

    const lat = parseFloat(String(body.Latitude).replace(',', '.'));
    const lng = parseFloat(String(body.Longitude).replace(',', '.'));

    if (isNaN(lat) || isNaN(lng)) {
      return buildJsonResponse({ success: false, error: 'Koordinat tidak valid.' });
    }

    // Auto-uppercase Pekerjaan and Jenis_UMKM for database consistency
    const pekerjaan = String(body.Pekerjaan || '-').toUpperCase();
    const jenisUmkm = String(body.Jenis_UMKM || '-').toUpperCase();

    sheet.appendRow([
      new Date(),
      String(body.NIK),
      body.Nama_Lengkap,
      body.Jenis_Kelamin || '-',
      String(body.RT),
      Number(body.Usia),
      body.Pendidikan || '-',
      pekerjaan,
      body.Status_Nikah || '-',
      jenisUmkm,
      body.Pembatik || 'Tidak',
      Number(body.Jumlah_Pembatik) || 0,
      body.Bantuan || 'Tidak Ada',
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

  const lastRow = sheet.getLastRow();
  if (lastRow > 0) {
    const firstCell = String(sheet.getRange(1, 1).getValue()).trim().toLowerCase();
    if (firstCell !== 'timestamp') {
      sheet.insertRowBefore(1);
    }
  }

  sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]).setFontWeight('bold').setBackground('#069C67').setFontColor('#fff');
  COLUMNS.forEach((_, i) => sheet.autoResizeColumn(i + 1));
  SpreadsheetApp.getUi().alert('Sheet "Data" berhasil diinisialisasi dengan ' + COLUMNS.length + ' kolom!');
}
