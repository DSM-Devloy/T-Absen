var CONFIG = {
  SHEET_NAME: "Sheet1",
  FOLDER_ID: "11e-JC6twgN2wF9teqEfEjUfFE_LN439X" // ID Folder Google Drive Anda
};

function doGet(e) {
  var action = e.parameter.action;
  
  if (action === "getDashboard") {
    return getDashboardData();
  }
  
  // Default response untuk pengecekan akses Web App
  return ContentService.createTextOutput(JSON.stringify({status: "active", message: "API Absensi Aktif"}))
                       .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
    
    // Jika sheet belum ada, buat otomatis
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(CONFIG.SHEET_NAME);
    }
    
    // Simpan Foto ke Google Drive
    var folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    var base64Data = data.photo.split(',')[1];
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), data.photoType, "Absensi_" + data.employeeId + "_" + new Date().getTime() + ".jpg");
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileUrl = file.getUrl();
    
    // Simpan data ke Google Sheets
    var timestamp = new Date();
    sheet.appendRow([
      timestamp,
      data.name,
      data.employeeId,
      data.location,
      data.latLng,
      fileUrl,
      data.status
    ]);
    
    // Menggunakan Header CORS agar bisa diakses dari Vercel
    return ContentService.createTextOutput(JSON.stringify({status: "success", message: "Absensi berhasil disimpan!"}))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: error.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

function getDashboardData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  var totalAbsen = 0;
  
  if (sheet) {
    var rows = sheet.getDataRange().getValues();
    totalAbsen = rows.length > 1 ? rows.length - 1 : 0;
  }
  
  var output = ContentService.createTextOutput(JSON.stringify({
    total: totalAbsen
  }));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

/**
 * Fungsi inisialisasi awal database (opsional dijalankan manual dari editor Apps Script)
 */
function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }
  
  var headers = ["Timestamp", "Nama Lengkap", "ID Karyawan", "Lokasi", "Latitude, Longitude", "Link Foto Drive", "Status"];
  var currentHeader = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  var isHeaderEmpty = currentHeader.every(function(cell) { return cell === ""; });
  
  if (isHeaderEmpty) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#2563EB");
    headerRange.setFontColor("#FFFFFF");
    headerRange.setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  
  SpreadsheetApp.getUi().alert("Setup Selesai!", "Database dan Header berhasil disiapkan.", SpreadsheetApp.getUi().ButtonSet.OK);
}
