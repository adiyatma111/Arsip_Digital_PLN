// Membuka halaman web utama
function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Sistem Arsip Digital')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Menyertakan file HTML/JS lain (jika dipisah)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Fungsi inisialisasi awal (membuat tabel/sheet otomatis)
function getorceretmaindra() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetArsip = ss.getSheetByName("Data_Arsip");
  if (!sheetArsip) {
    sheetArsip = ss.insertSheet("Data_Arsip");
    sheetArsip.appendRow(["Divisi", "Nama Dokumen", "Tempat Peletakan", "Masa Retensi", "Link File", "User Input"]);
  }
}

// Fungsi untuk mengambil semua data dari semua sheet lemari untuk ditampilkan di web
function ambilDataArsip() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var semuaData = [];
    
    // Looping semua sheet kecuali sheet 'Users' dan 'Data_Arsip' jika ada
    for (var i = 0; i < sheets.length; i++) {
      var name = sheets[i].getName();
      if (name !== "Users" && name !== "Data_Arsip") {
        var sheet = sheets[i];
        var data = sheet.getDataRange().getValues();
        
        // Ambil data dari baris ke-2 (skip header)
        for (var j = 1; j < data.length; j++) {
          var formattedDate = "";
          if (data[j][3] instanceof Date) {
            // Format date ke YYYY-MM-DD agar seragam
            formattedDate = Utilities.formatDate(data[j][3], ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
          } else {
            formattedDate = data[j][3];
          }
          
          semuaData.push({
            divisi: data[j][0],
            namaDokumen: data[j][1],
            kategori: data[j][2],
            masaRetensi: formattedDate,
            linkFile: data[j][4],
            userInput: data[j][5]
          });
        }
      }
    }
    
    // Sort di sisi backend: Masa Retensi Terbaru (paling jauh ke depan) -> Terlama
    semuaData.sort(function(a, b) {
      return new Date(b.masaRetensi) - new Date(a.masaRetensi);
    });
    
    return { status: "success", data: semuaData };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

// Fungsi untuk menyimpan data arsip baru
function simpanArsip(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var namaSheetTujuan = data.kategori; 
    var sheet = ss.getSheetByName(namaSheetTujuan);
    
    if (!sheet) {
      sheet = ss.insertSheet(namaSheetTujuan);
      sheet.appendRow(["Divisi", "Nama Dokumen", "Tempat Peletakan", "Masa Retensi", "Link File", "User Input"]);
    }
    
    var objekMasaRetensi = new Date(data.masaRetensi);
    
    sheet.appendRow([
      data.divisi, 
      data.namaDokumen, 
      namaSheetTujuan.replace("_", " "), 
      objekMasaRetensi, 
      data.linkFile, 
      data.userInput
    ]);
    
    var jumlahBarisData = sheet.getLastRow() - 1;
    if (jumlahBarisData > 1) {
      var rangeData = sheet.getRange(2, 1, jumlahBarisData, sheet.getLastColumn());
      rangeData.sort({column: 4, ascending: false}); // Urutkan Kolom 4 (Masa Retensi) Descending
      sheet.getRange(2, 4, jumlahBarisData, 1).setNumberFormat("yyyy-mm-dd");
    }
    
    return { status: "success", message: "Data arsip berhasil disimpan di sheet " + namaSheetTujuan.replace("_", " ") + "!" };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}
