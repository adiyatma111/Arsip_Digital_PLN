function doGet() {
  DriveApp.getRootFolder();

  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Sistem Arsip Digital')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getorceretmaindra() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetUsers = ss.getSheetByName("Users");
  if (!sheetUsers) {
    sheetUsers = ss.insertSheet("Users");
    sheetUsers.appendRow(["Username", "Password", "Role", "Folder"]);
    sheetUsers.appendRow(["admin", "admin123", "Super Admin", "All"]);
  }
  
  var sheetLemari = ss.getSheetByName("Master_Lemari");
  if (!sheetLemari) {
    sheetLemari = ss.insertSheet("Master_Lemari");
    sheetLemari.appendRow(["Nama Lemari", "Kode Sheet"]);
    sheetLemari.appendRow(["Lemari A", "Lemari_A"]);
    sheetLemari.appendRow(["Lemari B", "Lemari_B"]);
    sheetLemari.appendRow(["Lemari C", "Lemari_C"]);
    sheetLemari.appendRow(["Lemari D", "Lemari_D"]);
  }
}

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

function ambilDaftarLemari() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Master_Lemari");
    if (!sheet) {
      getorceretmaindra();
      sheet = ss.getSheetByName("Master_Lemari");
    }
    var data = sheet.getDataRange().getValues();
    var lemariList = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        lemariList.push({
          nama: data[i][0].toString(),
          kode: data[i][1] ? data[i][1].toString() : data[i][0].toString().replace(/\s+/g, '_')
        });
      }
    }
    return { status: "success", data: lemariList };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

function tambahLemariBaruWeb(namaLemariBaru, currentUserRole) {
  try {
    if (currentUserRole !== "Super Admin") {
      return { status: "error", message: "Akses ditolak! Hanya Super Admin yang bisa menambah lemari." };
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetMaster = ss.getSheetByName("Master_Lemari");
    if (!sheetMaster) {
      getorceretmaindra();
      sheetMaster = ss.getSheetByName("Master_Lemari");
    }
    
    var namaClean = namaLemariBaru.trim();
    var kodeSheet = namaClean.replace(/\s+/g, '_');
    
    var data = sheetMaster.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === namaClean.toLowerCase()) {
        return { status: "error", message: "Lemari dengan nama tersebut sudah ada!" };
      }
    }
    
    sheetMaster.appendRow([namaClean, kodeSheet]);
    
    var sheetBaru = ss.getSheetByName(kodeSheet);
    if (!sheetBaru) {
      sheetBaru = ss.insertSheet(kodeSheet);
      sheetBaru.appendRow(["Nama Folder Berkas", "Nama Dokumen", "Tempat Peletakan", "Masa Retensi", "Link File", "User Input"]);
    }
    
    return { status: "success", message: "Lemari '" + namaClean + "' berhasil ditambahkan!" };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

function ambilDataArsip(userDivisi, userRole) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var resLemari = ambilDaftarLemari();
    var lemariList = resLemari.status === "success" ? resLemari.data : [];
    
    var semuaData = [];
    var statistikData = { "Total": 0, "PFK": 0, "ESTETIKA": 0, "Pelanggan TM": 0, "SPKLU": 0 };
    
    lemariList.forEach(function(l) {
      statistikData[l.nama] = 0;
    });

    var isAdmin = (userRole === "Admin" || userRole === "Super Admin" || userRole === "admin");

    for (var k = 0; k < lemariList.length; k++) {
      var sheetName = lemariList[k].kode;
      var sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        var data = sheet.getDataRange().getValues();
        for (var j = 1; j < data.length; j++) {
          var div = data[j][0] ? data[j][0].toString().trim() : "";
          var lemariRaw = lemariList[k].nama;
          var formattedDate = "";
          
          if (data[j][3] instanceof Date) {
            formattedDate = Utilities.formatDate(data[j][3], ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
          } else {
            formattedDate = data[j][3];
          }
          
          if (!isAdmin && userDivisi !== "All" && div.toLowerCase() !== userDivisi.toLowerCase()) {
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

          statistikData["Total"]++;
          if (statistikData.hasOwnProperty(div)) { statistikData[div]++; }
          if (statistikData.hasOwnProperty(lemariRaw)) { statistikData[lemariRaw]++; }
        }
      }
    }
    
    semuaData.sort(function(a, b) {
      if (a.kategori < b.kategori) return -1;
      if (a.kategori > b.kategori) return 1;
      return new Date(b.masaRetensi) - new Date(a.masaRetensi);
    });
    
    return { status: "success", data: semuaData, statistik: statistikData, lemariList: lemariList };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

// FUNGSI UNTUK UNGGAN FILE PDF KE GOOGLE DRIVE DENGAN FOLDER OTOMATIS
function uploadFileToDrive(base64Data, fileName, divisi, masaRetensi) {
  try {
    var splitData = base64Data.split(',');
    var contentType = splitData[0].match(/:(.*?);/)[1];
    var fileData = Utilities.base64Decode(splitData[1]);
    var blob = Utilities.newBlob(fileData, contentType, fileName);
    
    var tahun = new Date().getFullYear().toString();
    if (masaRetensi) {
      var thnRetensi = masaRetensi.split('-')[0];
      if (thnRetensi && thnRetensi.length === 4) {
        tahun = thnRetensi;
      }
    }
    
    var targetFolder = getFolderTujuan(divisi, tahun);
    var file = targetFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return {
      status: "success",
      url: file.getUrl()
    };
  } catch (error) {
    return {
      status: "error",
      message: "Gagal mengunggah file PDF: " + error.toString()
    };
  }
}

// HELPER FUNCTION: Dapatkan/Buat Hirarki Folder di Drive
function getFolderTujuan(divisi, tahun) {
  var parentName = "Arsip Digital PLN";
  var divisiName = divisi ? "Folder " + divisi : "Umum";
  
  var mainFolders = DriveApp.getFoldersByName(parentName);
  var mainFolder = mainFolders.hasNext() ? mainFolders.next() : DriveApp.createFolder(parentName);
  
  var divisiFolders = mainFolder.getFoldersByName(divisiName);
  var divisiFolder = divisiFolders.hasNext() ? divisiFolders.next() : mainFolder.createFolder(divisiName);
  
  var tahunFolders = divisiFolder.getFoldersByName(tahun);
  var tahunFolder = tahunFolders.hasNext() ? tahunFolders.next() : divisiFolder.createFolder(tahun);
  
  return tahunFolder;
}

function simpanArsip(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var namaSheetTujuan = data.kategori.replace(/\s+/g, '_'); 
    var sheet = ss.getSheetByName(namaSheetTujuan);
    
    if (!sheet) {
      sheet = ss.insertSheet(namaSheetTujuan);
      sheet.appendRow(["Nama Folder Berkas", "Nama Dokumen", "Tempat Peletakan", "Masa Retensi", "Link File", "User Input"]);
    }
    
    sheet.appendRow([
      data.divisi, data.namaDokumen, data.kategori, data.masaRetensi, data.linkFile, data.userInput
    ]);
    
    sortSheetRetensi(sheet);
    return { status: "success", message: "Data arsip berhasil disimpan di " + data.kategori + "!" };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

function hapusArsipWeb(namaDokumen, kategori, userInputDokumen, userRole, currentUsername) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var namaSheet = kategori.replace(/\s+/g, '_');
    var sheet = ss.getSheetByName(namaSheet);
    
    if (!sheet) return { status: "error", message: "Sheet tidak ditemukan." };
    
    var isAdmin = (userRole === "Admin" || userRole === "Super Admin" || userRole === "admin");
    if (!isAdmin && userInputDokumen !== currentUsername) {
      return { status: "error", message: "Anda tidak memiliki akses untuk menghapus dokumen ini." };
    }
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1].toString() === namaDokumen && data[i][5].toString() === userInputDokumen) {
        sheet.deleteRow(i + 1);
        return { status: "success", message: "Dokumen '" + namaDokumen + "' berhasil dihapus!" };
      }
    }
    return { status: "error", message: "Data dokumen tidak ditemukan." };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

function updateArsipWeb(dataLama, dataBaru, userRole, currentUsername) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var isAdmin = (userRole === "Admin" || userRole === "Super Admin" || userRole === "admin");
    if (!isAdmin && dataLama.userInput !== currentUsername) {
      return { status: "error", message: "Anda tidak memiliki akses untuk mengubah dokumen ini." };
    }
    
    var sheetLama = ss.getSheetByName(dataLama.kategori.replace(/\s+/g, '_'));
    var sheetBaru = ss.getSheetByName(dataBaru.kategori.replace(/\s+/g, '_'));
    
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
    
    if (dataLama.kategori !== dataBaru.kategori) {
      sheetLama.deleteRow(barisDitemukan);
      if (!sheetBaru) {
        sheetBaru = ss.insertSheet(dataBaru.kategori.replace(/\s+/g, '_'));
        sheetBaru.appendRow(["Nama Folder Berkas", "Nama Dokumen", "Tempat Peletakan", "Masa Retensi", "Link File", "User Input"]);
      }
      sheetBaru.appendRow([
        dataBaru.divisi, dataBaru.namaDokumen, dataBaru.kategori, dataBaru.masaRetensi, dataBaru.linkFile, dataLama.userInput
      ]);
      sortSheetRetensi(sheetBaru);
    } else {
      sheetLama.getRange(barisDitemukan, 1, 1, 6).setValues([[
        dataBaru.divisi, dataBaru.namaDokumen, dataBaru.kategori, dataBaru.masaRetensi, dataBaru.linkFile, dataLama.userInput
      ]]);
      sortSheetRetensi(sheetLama);
    }
    
    return { status: "success", message: "Dokumen berhasil diperbarui!" };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

function sortSheetRetensi(sheet) {
  var jumlahBarisData = sheet.getLastRow() - 1;
  if (jumlahBarisData > 1) {
    var rangeData = sheet.getRange(2, 1, jumlahBarisData, sheet.getLastColumn());
    rangeData.sort({column: 4, ascending: false});
    sheet.getRange(2, 4, jumlahBarisData, 1).setNumberFormat("yyyy-mm-dd");
  }
}

function ambilSemuaUser(currentUserRole) {
  try {
    var isSuperAdmin = (currentUserRole === "Super Admin");
    if (!isSuperAdmin) return { status: "error", message: "Akses ditolak!" };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Users");
    if (!sheet) { getorceretmaindra(); sheet = ss.getSheetByName("Users"); }

    var data = sheet.getDataRange().getValues();
    var daftarUser = [];
    for (var i = 1; i < data.length; i++) {
      daftarUser.push({ username: data[i][0], password: data[i][1], role: data[i][2], divisi: data[i][3] });
    }
    return { status: "success", data: daftarUser };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

function tambahUserBaruWeb(dataUserBaru, currentUserRole) {
  try {
    var isSuperAdmin = (currentUserRole === "Super Admin");
    if (!isSuperAdmin) return { status: "error", message: "Akses ditolak!" };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === dataUserBaru.username.toLowerCase()) {
        return { status: "error", message: "Username sudah terdaftar!" };
      }
    }

    sheet.appendRow([dataUserBaru.username, dataUserBaru.password, dataUserBaru.role, dataUserBaru.divisi]);
    return { status: "success", message: "User berhasil ditambahkan!" };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

function updateUserDataWeb(usernameLama, dataBaru, currentUserRole) {
  try {
    var isSuperAdmin = (currentUserRole === "Super Admin");
    if (!isSuperAdmin) return { status: "error", message: "Akses ditolak!" };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    var barisDitemukan = -1;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === usernameLama.toLowerCase()) {
        barisDitemukan = i + 1;
        break;
      }
    }

    if (barisDitemukan === -1) return { status: "error", message: "User tidak ditemukan." };

    sheet.getRange(barisDitemukan, 1, 1, 4).setValues([[
      dataBaru.username, dataBaru.password, dataBaru.role, dataBaru.divisi
    ]]);

    return { status: "success", message: "Data user berhasil diperbarui!" };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

function hapusUserWeb(usernameYangDihapus, currentUserRole, currentUsername) {
  try {
    var isSuperAdmin = (currentUserRole === "Super Admin");
    if (!isSuperAdmin) return { status: "error", message: "Akses ditolak!" };

    if (usernameYangDihapus.toLowerCase() === currentUsername.toLowerCase()) {
      return { status: "error", message: "Tidak bisa menghapus diri sendiri!" };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === usernameYangDihapus.toLowerCase()) {
        sheet.deleteRow(i + 1);
        return { status: "success", message: "User berhasil dihapus!" };
      }
    }
    return { status: "error", message: "User tidak ditemukan." };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}
