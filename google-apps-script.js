/**
 * =========================================================================
 * GOOGLE APPS SCRIPT: RSVP & UCAPAN PERNIKAHAN KIRANA & JAROT
 * =========================================================================
 * Spreadsheet ID: 144ha50NKgdG1JR2cto0nupENZ2W-mDt-KkGWn9xAwD8
 * Sheet Name    : Sheet1
 *
 * Kolom yang digunakan:
 * 1. timestamp
 * 2. nama tamu
 * 3. ucapan
 * 4. konfirmasi kehadiran
 * 5. jumlah tamu
 * =========================================================================
 */

var SPREADSHEET_ID = "144ha50NKgdG1JR2cto0nupENZ2W-mDt-KkGWn9xAwD8";
var SHEET_NAME = "Sheet1";

/**
 * Endpoint GET:
 * - Mengambil daftar ucapan (doGet)
 * - Menerima fallback submit via GET jika diperlukan (CORS friendly)
 */
function doGet(e) {
  try {
    // Cek jika ada request submit lewat GET (fallback CORS)
    if (e && e.parameter && e.parameter.action === "submit") {
      return handleSubmission(e.parameter);
    }

    return handleGetWishes();
  } catch (error) {
    return createJsonResponse({
      status: "error",
      message: error.toString()
    });
  }
}

/**
 * Endpoint POST:
 * - Menyimpan RSVP & Ucapan baru ke Spreadsheet
 */
function doPost(e) {
  try {
    var data = {};
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    return handleSubmission(data);
  } catch (error) {
    return createJsonResponse({
      status: "error",
      message: error.toString()
    });
  }
}

/**
 * Helper untuk memproses penyimpanan ke sheet
 */
function handleSubmission(data) {
  var sheet = getSheet();
  ensureHeader(sheet);

  var now = new Date();
  var timestampStr = Utilities.formatDate(now, "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");

  var namaTamu = (data.nama || data.nama_tamu || data["nama tamu"] || "Tamu Undangan").toString().trim();
  var ucapan = (data.ucapan || data.pesan || data.doa || "-").toString().trim();
  var konfirmasi = (data.kehadiran || data.konfirmasi || data["konfirmasi kehadiran"] || "Hadir").toString().trim();
  var jumlahTamu = (konfirmasi.toLowerCase().indexOf("tidak") !== -1) ? "0" : (data.jumlah || data.jumlah_tamu || data["jumlah tamu"] || "1");

  // Tambahkan baris baru sesuai urutan kolom yang diminta
  sheet.appendRow([
    timestampStr,
    namaTamu,
    ucapan,
    konfirmasi,
    jumlahTamu
  ]);

  return createJsonResponse({
    status: "success",
    message: "Terima kasih, ucapan & konfirmasi kehadiran Anda berhasil disimpan!",
    data: {
      timestamp: timestampStr,
      nama: namaTamu,
      ucapan: ucapan,
      kehadiran: konfirmasi,
      jumlah: jumlahTamu
    }
  });
}

/**
 * Helper untuk mengambil seluruh data ucapan
 */
function handleGetWishes() {
  var sheet = getSheet();
  ensureHeader(sheet);

  var values = sheet.getDataRange().getValues();

  // Jika hanya ada header
  if (values.length <= 1) {
    return createJsonResponse({
      status: "success",
      total: 0,
      data: []
    });
  }

  var rows = values.slice(1);
  var wishes = [];

  // Tampilkan urutan dari yang terbaru (paling bawah) ke terlama
  for (var i = rows.length - 1; i >= 0; i--) {
    var row = rows[i];
    var timestamp = row[0];
    var namaTamu = row[1];
    var ucapan = row[2];
    var konfirmasi = row[3];
    var jumlahTamu = row[4];

    if (namaTamu || ucapan) {
      var formattedTime = "";
      if (timestamp instanceof Date) {
        formattedTime = Utilities.formatDate(timestamp, "Asia/Jakarta", "dd MMM yyyy, HH:mm");
      } else {
        formattedTime = String(timestamp || "");
      }

      wishes.push({
        timestamp: formattedTime,
        nama: String(namaTamu || "").trim(),
        ucapan: String(ucapan || "").trim(),
        kehadiran: String(konfirmasi || "").trim(),
        jumlah: jumlahTamu || 1
      });
    }
  }

  return createJsonResponse({
    status: "success",
    total: wishes.length,
    data: wishes
  });
}

/**
 * Ambil Sheet berdasarkan Spreadsheet ID & Sheet Name
 */
function getSheet() {
  var ss;
  if (SPREADSHEET_ID && SPREADSHEET_ID.length > 10) {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  return sheet;
}

/**
 * Pastikan Baris Header Kolom sudah ada
 */
function ensureHeader(sheet) {
  if (sheet.getLastRow() === 0) {
    var headers = [
      "timestamp",
      "nama tamu",
      "ucapan",
      "konfirmasi kehadiran",
      "jumlah tamu"
    ];
    sheet.appendRow(headers);

    // Format header agar rapi
    var headerRange = sheet.getRange(1, 1, 1, 5);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#9c8366");
    headerRange.setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
}

/**
 * Helper untuk format response JSON dengan CORS header
 */
function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * FUNGSI SETUP (Jalankan sekali jika ingin inisialisasi sheet secara otomatis)
 */
function setupSheet() {
  var sheet = getSheet();
  ensureHeader(sheet);
  Logger.log("Inisialisasi sheet selesai!");
}
