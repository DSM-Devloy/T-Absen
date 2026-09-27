var SHEET_NAME = "Sheet1";
var FOLDER_ID = "https://docs.google.com/spreadsheets/d/1MmpAigYNP7kGJHvd0wkNSkNEyD0IxqPeXu0UQDuAjXQ/edit?usp=sharing"; 

function doGet(e) {
  var action = e.parameter.action;
  
  if (action === "getDashboard") {
    return getDashboardData();
  } else if (action === "getReport") {
    var startDate = e.parameter.start;
    var endDate = e.parameter.end;
    return getReportData(startDate, endDate);
  }
  
  // Default merender halaman utama (index.html)
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Aplikasi Absensi')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    
    // Simpan Foto ke Google Drive
    var folder = DriveApp.getFolderById(FOLDER_ID);
    var blob = Utilities.newBlob(Utilities.base64Decode(data.photo.split(',')[1]), data.photoType, "Absensi_" + data.employeeId + "_" + new Date().getTime() + ".jpg");
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileUrl = file.getUrl();
    
    // Simpan ke Google Sheets
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
    
    return ContentService.createTextOutput(JSON.stringify({status: "success", message: "Absensi berhasil disimpan!"}))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: error.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

function getDashboardData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var rows = sheet.getDataRange().getValues();
  var totalAbsen = rows.length > 1 ? rows.length - 1 : 0;
  
  // Kirim data ringkasan atau baris terakhir untuk dashboard
  var recent = rows.slice(1).reverse().slice(0, 5); // 5 absensi terakhir
  
  return ContentService.createTextOutput(JSON.stringify({
    total: totalAbsen,
    recent: recent
  })).setMimeType(ContentService.MimeType.JSON);
}
