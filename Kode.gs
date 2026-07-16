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

// Fungsi inisialisasi awal (membuat tabel/sheet otomatis jika belum ada)
function getorceretmaindra() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetUsers = ss.getSheetByName("Users");
  if (!sheetUsers) {
    sheetUsers = ss.insertSheet("Users");
    sheetUsers.appendRow(["Username", "Password", "Role", "Divisi"]);
    sheetUsers.appendRow(["admin", "admin123", "Super Admin", "All"]);
  }
}

// Fungsi untuk memeriksa username dan password dari sheet Users
function prosesLogin(username, password) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Users");
    if (!sheet) {
      getorceretmaindra();
      sheet = ss.getSheetByName("Users");
    }
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === username.toLowerCase() && data[i][1].toString() === password) {
        return { status: "success", role: data[i][2], divisi: data[i][3], username: data[i][0] };
      }
    }
    return { status: "failed", message: "Username atau Password salah!" };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

// Fungsi untuk mengambil semua data & menghitung statistik ringkasan dengan pembatasan hak akses divisi
function ambilDataArsip(userDivisi, userRole) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var semuaData = [];
    
    // Objek untuk menyimpan jumlah data per divisi dan lemari
    var statistikData = {
      "Total": 0,
      "SDM": 0,
      "Keuangan": 0,
      "Sistem Informasi": 0,
      "Distribusi": 0,
      "Lemari A": 0,
      "Lemari B": 0,
      "Lemari C": 0,
      "Lemari D": 0
    };
    
    // Cek apakah user adalah Admin/Super Admin
    var isAdmin = (userRole === "Admin" || userRole === "Super Admin" || userRole === "admin");

    for (var i = 0; i < sheets.length; i++) {
      var name = sheets[i].getName();
      if (name !== "Users" && name !== "Data_Arsip") {
        var sheet = sheets[i];
        var data = sheet.getDataRange().getValues();
        
        for (var j = 1; j < data.length; j++) {
          var div = data[j][0] ? data[j][0].toString().trim() : "";
          var lemariRaw = data[j][2] ? data[j][2].toString().replace("_", " ") : "";
          var formattedDate = "";
          
          if (data[j][3] instanceof Date) {
            formattedDate = Utilities.formatDate(data[j][3], ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
          } else {
            formattedDate = data[j][3];
          }
          
          // FILTER HAK AKSES DIVISI
          if (!isAdmin && div.toLowerCase() !== userDivisi.toLowerCase()) {
            continue;
          }

          semuaData.push({
            divisi: div,
            namaDokumen: data[j][1],
            kategori: lemariRaw,
            masaRetensi: formattedDate,
            linkFile: data[j][4],
            userInput: data[j][5]
          });

          // Hitung statistik
          statistikData["Total"]++;
          if (statistikData.hasOwnProperty(div)) {
            statistikData[div]++;
          }
          if (statistikData.hasOwnProperty(lemariRaw)) {
            statistikData[lemariRaw]++;
          }
        }
      }
    }
    
    // Urutkan: Lemari A -> B -> C -> D
    semuaData.sort(function(a, b) {
      if (a.kategori < b.kategori) return -1;
      if (a.kategori > b.kategori) return 1;
      return new Date(b.masaRetensi) - new Date(a.masaRetensi);
    });
    
    return { status: "success", data: semuaData, statistik: statistikData };
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
      rangeData.sort({column: 4, ascending: false});
      sheet.getRange(2, 4, jumlahBarisData, 1).setNumberFormat("yyyy-mm-dd");
    }
    
    return { status: "success", message: "Data arsip berhasil disimpan di sheet " + namaSheetTujuan.replace("_", " ") + "!" };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

// Tambahkan fungsi baru ini di paling bawah file Kode.gs Anda

// Fungsi untuk menghapus arsip langsung dari Web
function hapusArsipWeb(namaDokumen, kategori, userInputDokumen, userRole, currentUsername) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var namaSheet = kategori.replace(" ", "_");
    var sheet = ss.getSheetByName(namaSheet);
    
    if (!sheet) return { status: "error", message: "Sheet tidak ditemukan." };
    
    // Proteksi Keamanan Backend: Cek apakah user berhak menghapus
    var isAdmin = (userRole === "Admin" || userRole === "Super Admin" || userRole === "admin");
    if (!isAdmin && userInputDokumen !== currentUsername) {
      return { status: "error", message: "Anda tidak memiliki akses untuk menghapus dokumen ini." };
    }
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      // Validasi berdasarkan nama dokumen dan user input untuk memastikan baris yang tepat
      if (data[i][1].toString() === namaDokumen && data[i][5].toString() === userInputDokumen) {
        sheet.deleteRow(i + 1);
        return { status: "success", message: "Dokumen '" + namaDokumen + "' berhasil dihapus!" };
      }
    }
    return { status: "error", message: "Data dokumen tidak ditemukan di spreadsheet." };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

// Fungsi untuk mengedit/update arsip dari Web
function updateArsipWeb(dataLama, dataBaru, userRole, currentUsername) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Proteksi Keamanan Backend
    var isAdmin = (userRole === "Admin" || userRole === "Super Admin" || userRole === "admin");
    if (!isAdmin && dataLama.userInput !== currentUsername) {
      return { status: "error", message: "Anda tidak memiliki akses untuk mengubah dokumen ini." };
    }
    
    var sheetLama = ss.getSheetByName(dataLama.kategori.replace(" ", "_"));
    var sheetBaru = ss.getSheetByName(dataBaru.kategori.replace(" ", "_"));
    
    if (!sheetLama) return { status: "error", message: "Sheet asal tidak ditemukan." };
    
    var valuesLama = sheetLama.getDataRange().getValues();
    var barisDitemukan = -1;
    
    for (var i = 1; i < valuesLama.length; i++) {
      if (valuesLama[i][1].toString() === dataLama.namaDokumen && valuesLama[i][5].toString() === dataLama.userInput) {
        barisDitemukan = i + 1;
        break;
      }
    }
    
    if (barisDitemukan === -1) return { status: "error", message: "Data asli tidak ditemukan." };
    
    var objekMasaRetensi = new Date(dataBaru.masaRetensi);
    
    // Jika pindah lemari/kategori sheet
    if (dataLama.kategori !== dataBaru.kategori) {
      sheetLama.deleteRow(barisDitemukan); // Hapus dari sheet lama
      
      if (!sheetBaru) {
        sheetBaru = ss.insertSheet(dataBaru.kategori.replace(" ", "_"));
        sheetBaru.appendRow(["Divisi", "Nama Dokumen", "Tempat Peletakan", "Masa Retensi", "Link File", "User Input"]);
      }
      sheetBaru.appendRow([
        dataBaru.divisi, dataBaru.namaDokumen, dataBaru.kategori.replace("_", " "), objekMasaRetensi, dataBaru.linkFile, dataLama.userInput
      ]);
      sortSheetRetensi(sheetBaru);
    } else {
      // Jika tetap di lemari yang sama, cukup update barisnya
      sheetLama.getRange(barisDitemukan, 1, 1, 6).setValues([[
        dataBaru.divisi, dataBaru.namaDokumen, dataBaru.kategori.replace("_", " "), objekMasaRetensi, dataBaru.linkFile, dataLama.userInput
      ]]);
      sortSheetRetensi(sheetLama);
    }
    
    return { status: "success", message: "Dokumen berhasil diperbarui!" };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

// Fungsi pembantu untuk otomatis mengurutkan ulang tanggal setelah diubah
function sortSheetRetensi(sheet) {
  var jumlahBarisData = sheet.getLastRow() - 1;
  if (jumlahBarisData > 1) {
    var rangeData = sheet.getRange(2, 1, jumlahBarisData, sheet.getLastColumn());
    rangeData.sort({column: 4, ascending: false});
    sheet.getRange(2, 4, jumlahBarisData, 1).setNumberFormat("yyyy-mm-dd");
  }
}
