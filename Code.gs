// ============================================================
// SI-DEMANG KREBET — Google Apps Script Backend
// REST API for Village Demographic & UMKM Mapping
// Supports CREATE and UPDATE actions
// ============================================================

const SHEET_NAME = 'Data';
const COLUMNS = [
  'Timestamp',
  'No_KK',
  'NIK',
  'Nama_Lengkap',
  'Jenis_Kelamin',
  'RT',
  'Usia',
  'Tanggal_Lahir',
  'Pendidikan',
  'Pekerjaan',
  'Status_Nikah',
  'Status_Kematian',
  'Jenis_UMKM',
  'Pembatik',
  'Jumlah_Pembatik',
  'Bantuan',
  'Latitude',
  'Longitude',
  'Status_Kependudukan'
];

// ============================================================
// GET — Read all data
// ============================================================
function doGet(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return buildJsonResponse({ success: false, error: 'Sheet not found' });

    const lastRow = sheet.getLastRow();
    if (lastRow < 1) return buildJsonResponse({ success: true, data: [], count: 0 });

    const data = sheet.getDataRange().getValues();

    // Detect header row
    let startIndex = 0;
    const firstCell = String(data[0][0]).trim().toLowerCase();
    if (firstCell === 'timestamp') {
      startIndex = 1;
    }

    const rows = data.slice(startIndex);
    const result = [];

    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];

      // Skip completely empty rows
      const nikVal = String(row[2] || '').trim();
      const namaVal = String(row[3] || '').trim();
      if (!nikVal && !namaVal) continue;

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

      // Include the actual sheet row number for reference
      obj._rowIndex = startIndex + r + 1; // 1-indexed sheet row
      result.push(obj);
    }

    return buildJsonResponse({ success: true, data: result, count: result.length });
  } catch (error) {
    return buildJsonResponse({ success: false, error: error.message });
  }
}

// ============================================================
// POST — Create or Update data
// ============================================================
function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return buildJsonResponse({ success: false, error: 'Sheet not found' });

    const body = JSON.parse(e.postData.contents);
    const action = String(body.action || 'create').toLowerCase();

    // Accept flexible field names
    const namaLengkap = body.Nama_Lengkap || body.Nama || '-';

    // Soft validation — only Nama_Lengkap is truly required (except for delete)
    if (action !== 'delete' && (!namaLengkap || namaLengkap === '-')) {
      return buildJsonResponse({ success: false, error: 'Field "Nama_Lengkap" wajib diisi.' });
    }

    // Build the row values with defaults for missing fields
    const noKK = String(body.No_KK || '-').trim();
    const nik = String(body.NIK || '-').trim();
    const jenisKelamin = body.Jenis_Kelamin || '-';
    const rt = String(body.RT || '-').trim();
    const usia = (body.Usia !== undefined && body.Usia !== '') ? Number(body.Usia) : 0;
    const tanggalLahir = body.Tanggal_Lahir || '-';
    const pendidikan = body.Pendidikan || '-';

    // Auto-uppercase for consistency
    const pekerjaan = body.Pekerjaan ? String(body.Pekerjaan).toUpperCase() : '-';
    const statusNikah = body.Status_Nikah || '-';
    const statusKematian = body.Status_Kematian || 'Hidup';
    const jenisUmkm = body.Jenis_UMKM ? String(body.Jenis_UMKM).toUpperCase() : '-';

    const pembatik = body.Pembatik || body.Keluarga_Pembatik || 'Tidak';
    const jumlahPembatik = Number(body.Jumlah_Pembatik) || 0;
    const bantuan = body.Bantuan || 'Tidak Ada';

    // Parse coordinates
    const lat = parseFloat(String(body.Latitude || '0').replace(',', '.'));
    const lng = parseFloat(String(body.Longitude || '0').replace(',', '.'));
    const latitude = isNaN(lat) ? 0 : lat;
    const longitude = isNaN(lng) ? 0 : lng;

    const statusKependudukan = body.Status_Kependudukan || 'Warga Asli (KTP/KK Krebet)';

    // ============ CREATE ============
    if (action === 'create') {
      // Validasi NIK tidak boleh sama
      if (nik && nik !== '-' && nik.length > 5) {
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          const niks = sheet.getRange(2, 3, lastRow - 1, 1).getValues(); // Column C (NIK)
          for (let i = 0; i < niks.length; i++) {
            if (String(niks[i][0]).trim() === nik) {
              return buildJsonResponse({ success: false, error: 'Data gagal disimpan: NIK ini sudah terdaftar di database.' });
            }
          }
        }
      }

      const rowData = [
        new Date(),        // Timestamp
        noKK,              // No_KK
        nik,               // NIK
        namaLengkap,       // Nama_Lengkap
        jenisKelamin,      // Jenis_Kelamin
        rt,                // RT
        usia,              // Usia
        tanggalLahir,      // Tanggal_Lahir
        pendidikan,        // Pendidikan
        pekerjaan,         // Pekerjaan
        statusNikah,       // Status_Nikah
        statusKematian,    // Status_Kematian
        jenisUmkm,         // Jenis_UMKM
        pembatik,          // Pembatik
        jumlahPembatik,    // Jumlah_Pembatik
        bantuan,           // Bantuan
        latitude,          // Latitude
        longitude,         // Longitude
        statusKependudukan // Status_Kependudukan
      ];

      sheet.appendRow(rowData);

      // Auto-sort by No_KK (column 2) to group families together
      const lastRow = sheet.getLastRow();
      if (lastRow > 2) {
        sheet.getRange(2, 1, lastRow - 1, COLUMNS.length).sort(2);
      }

      return buildJsonResponse({ success: true, message: 'Data berhasil disimpan!' });
    }

    // ============ UPDATE ============
    if (action === 'update') {
      const targetTimestamp = body.Timestamp;
      if (!targetTimestamp) {
        return buildJsonResponse({ success: false, error: 'Timestamp diperlukan untuk update.' });
      }

      // Find the row by matching Timestamp
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) {
        return buildJsonResponse({ success: false, error: 'Tidak ada data untuk diupdate.' });
      }

      const timestamps = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); // Column A (Timestamp)
      let targetRow = -1;

      for (let i = 0; i < timestamps.length; i++) {
        const cellVal = timestamps[i][0];
        const cellStr = (cellVal instanceof Date) ? cellVal.toISOString() : String(cellVal);
        if (cellStr === targetTimestamp) {
          targetRow = i + 2; // +2 because: +1 for header, +1 for 1-indexed
          break;
        }
      }

      if (targetRow === -1) {
        return buildJsonResponse({ success: false, error: 'Data tidak ditemukan untuk diupdate.' });
      }

      // Build updated row (keep original timestamp)
      const updatedRow = [
        timestamps[targetRow - 2][0], // Keep original Timestamp
        noKK,
        nik,
        namaLengkap,
        jenisKelamin,
        rt,
        usia,
        tanggalLahir,
        pendidikan,
        pekerjaan,
        statusNikah,
        statusKematian,
        jenisUmkm,
        pembatik,
        jumlahPembatik,
        bantuan,
        latitude,
        longitude,
        statusKependudukan
      ];

      sheet.getRange(targetRow, 1, 1, COLUMNS.length).setValues([updatedRow]);

      return buildJsonResponse({ success: true, message: 'Data berhasil diperbarui!' });
    }

    // ============ DELETE ============
    if (action === 'delete') {
      const targetTimestamp = body.Timestamp;
      if (!targetTimestamp) {
        return buildJsonResponse({ success: false, error: 'Timestamp diperlukan untuk delete.' });
      }

      const lastRow = sheet.getLastRow();
      if (lastRow < 2) {
        return buildJsonResponse({ success: false, error: 'Tidak ada data untuk dihapus.' });
      }

      const timestamps = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      let targetRow = -1;

      for (let i = 0; i < timestamps.length; i++) {
        const cellVal = timestamps[i][0];
        const cellStr = (cellVal instanceof Date) ? cellVal.toISOString() : String(cellVal);
        if (cellStr === targetTimestamp) {
          targetRow = i + 2;
          break;
        }
      }

      if (targetRow === -1) {
        return buildJsonResponse({ success: false, error: 'Data tidak ditemukan untuk dihapus.' });
      }

      sheet.deleteRow(targetRow);
      return buildJsonResponse({ success: true, message: 'Data berhasil dihapus!' });
    }

    return buildJsonResponse({ success: false, error: 'Action tidak dikenal: ' + action });

  } catch (error) {
    return buildJsonResponse({ success: false, error: error.message });
  }
}

// ============================================================
// Helpers
// ============================================================
function buildJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function initializeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  sheet.clearContents();

  sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]).setFontWeight('bold').setBackground('#069C67').setFontColor('#fff');
  COLUMNS.forEach((_, i) => sheet.autoResizeColumn(i + 1));
  SpreadsheetApp.getUi().alert('Sheet "Data" berhasil diinisialisasi dengan ' + COLUMNS.length + ' kolom!\nKolom: ' + COLUMNS.join(', '));
}
