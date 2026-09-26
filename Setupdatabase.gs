/**
 * Konfigurasi Utama Database & Storage
 * Sesuaikan FOLDER_ID dengan ID folder Google Drive tempat menyimpan foto selfie.
 */
var CONFIG = {
  SHEET_NAME: "Sheet1",
  FOLDER_ID: "11e-JC6twgN2wF9teqEfEjUfFE_LN439X"
};

/**
 * Fungsi utama untuk inisialisasi database dan struktur sheet secara otomatis.
 * Jalankan fungsi ini sekali saat pertama kali menyiapkan aplikasi.
 */
function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  // Jika sheet belum ada, buat baru
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    Logger.log("Sheet '" + CONFIG.SHEET_NAME + "' berhasil dibuat.");
  }
  
  // 1. Setup Header Kolom untuk Absensi
  var headers = [
    "Timestamp", 
    "Nama Lengkap", 
    "ID Karyawan", 
    "Lokasi", 
    "Latitude, Longitude", 
    "Link Foto Drive", 
    "Status"
  ];
  
  // Periksa apakah baris pertama masih kosong atau belum sesuai
  var currentHeader = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  var isHeaderEmpty = currentHeader.every(function(cell) { return cell === ""; });
  
  if (isHeaderEmpty) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format Styling Header (Opsional agar terlihat rapi)
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#2563EB"); // Warna biru Tailwind
    headerRange.setFontColor("#FFFFFF");
    headerRange.setFontWeight("bold");
    sheet.setFrozenRows(1);
    
    Logger.log("Header kolom berhasil diatur dan diformat.");
  } else {
    Logger.log("Header kolom sudah ada sebelumnya, melewati pembuatan header.");
  }
  
  // 2. Verifikasi Folder Google Drive
  try {
    var folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    Logger.log("Koneksi Google Drive berhasil! Folder ditemukan: " + folder.getName());
  } catch (e) {
    Logger.log("PERINGATAN: Gagal mengakses Folder Google Drive. Pastikan FOLDER_ID sudah diisi dengan benar dan bot/akun memiliki akses.");
  }
  
  SpreadsheetApp.getUi().alert("Setup Database Selesai!", "Struktur database dan folder Google Drive telah diverifikasi.", SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Fungsi utilitas untuk mengecek status koneksi dan isi database cepat via log.
 */
function testDatabaseConnection() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var lastRow = sheet.getLastRow();
  
  Logger.log("Total baris data saat ini (termasuk header): " + lastRow);
}