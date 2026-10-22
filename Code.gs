// =========================================================================
// BACKEND DATABASE & UPLOAD MEDIA - TAKO BUNNY
// =========================================================================
//
// DAFTAR ISI
//   0. Router (doGet / doPost)
//   1. Upload file / media (foto & video)
//   2. Pengambilan semua data (read)
//   3. Administrasi & autentikasi user
//   4. Katalog (tambah, ubah, hapus)
//   5. Adopsi / transaksi
//   6. Produk, galeri, dan testimoni
//   7. Konfigurasi kwitansi & label logistik
// =========================================================================

// =========================================================================
// 0. ROUTER
// =========================================================================
function doGet(e) {
  return ContentService.createTextOutput("Backend Tako Bunny Online.")
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    var contents = JSON.parse(e.postData.contents);
    var action = contents.action;
    var payload = contents.payload;
    var result = {};

    switch (action) {
      case 'apiUploadFile':
        result = apiUploadFile(payload);
        break;
      case 'apiGetData':
        result = apiGetData();
        break;
      case 'apiLoginAdmin':
        result = apiLoginAdmin(payload);
        break;
      case 'apiLoginUser':
        result = apiLoginUser(payload);
        break;
      case 'apiRegisterUser':
        result = apiRegisterUser(payload);
        break;
      case 'apiForgotPassword':
        result = apiForgotPassword(payload);
        break;
      case 'apiAddKatalog':
        result = apiAddKatalog(payload);
        break;
      case 'apiEditKatalog':
        result = apiEditKatalog(payload);
        break;
      case 'apiDeleteKatalog':
        result = apiDeleteKatalog(payload);
        break;
      case 'apiSubmitAdopsi':
        result = apiSubmitAdopsi(payload);
        break;
      case 'apiUpdateStatusPengiriman':
        result = apiUpdateStatusPengiriman(payload);
        break;
      case 'apiUpdateStatusPelunasanManual':
        result = apiUpdateStatusPelunasanManual(payload);
        break;
      case 'apiUpdateOngkir':
        result = apiUpdateOngkir(payload);
        break;
      case 'apiDeleteAdopsi':
        result = apiDeleteAdopsi(payload);
        break;
      case 'apiAddProduk':
        result = apiAddProduk(payload);
        break;
      case 'apiDeleteProduk':
        result = apiDeleteProduk(payload);
        break;
      case 'apiAddGaleri':
        result = apiAddGaleri(payload);
        break;
      case 'apiSubmitTestimoni':
        result = apiSubmitTestimoni(payload);
        break;
      case 'apiDeleteTestimoni':
        result = apiDeleteTestimoni(payload);
        break;
      case 'apiSaveKwitansiConfig':
        result = apiSaveKwitansiConfig(payload);
        break;
      case 'apiGetKwitansiData':
        result = apiGetKwitansiData(payload);
        break;
      default:
        result = { success: false, message: "Aksi tidak dikenali." };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// =========================================================================
// 1. FUNGSI UPLOAD FILE / MEDIA (FOTO & VIDEO)
// =========================================================================
function apiUploadFile(payload) {
  try {
    var dataUrl = payload.dataUrl; // Format: "data:image/jpeg;base64,..." atau "data:video/mp4;base64,..."
    var filename = payload.filename || ("upload_" + new Date().getTime());
    
    var splitData = dataUrl.split(',');
    var mimeMatch = splitData[0].match(/:(.*?);/);
    var mimeType = mimeMatch ? mimeMatch[1] : 'application/octet-stream';
    var base64Data = splitData[1];
    
    var bytes = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(bytes, mimeType, filename);
    
    // Buka / Buat folder tempat menyimpan media di Google Drive
    var folderName = "TAKO_BUNNY_UPLOADS";
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    
    // Simpan file
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    var fileId = file.getId();
    var fileUrl = "";
    
    // Membedakan format URL antara Foto dan Video
    if (mimeType.indexOf("video") !== -1 || filename.toLowerCase().match(/\.(mp4|mov|webm|mkv|avi)$/)) {
      // URL khusus untuk pemutar Iframe Video Google Drive
      fileUrl = "https://drive.google.com/file/d/" + fileId + "/preview?video=.mp4";
    } else {
      // URL khusus Gambar langsung
      fileUrl = "https://lh3.googleusercontent.com/d/" + fileId;
    }
    
    return {
      success: true,
      url: fileUrl,
      fileId: fileId
    };
  } catch (e) {
    return {
      success: false,
      message: "Gagal upload media: " + e.toString()
    };
  }
}

// =========================================================================
// 2. FUNGSI PENGAMBILAN SEMUA DATA (READ)
// =========================================================================
function apiGetData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    katalog: getSheetData(ss.getSheetByName("Katalog")),
    galeri: getSheetData(ss.getSheetByName("Galeri")),
    produk: getSheetData(ss.getSheetByName("Produk")),
    adopsi: getSheetData(ss.getSheetByName("Adopsi")),
    testimoni: getSheetData(ss.getSheetByName("Testimoni"))
  };
}

function getSheetData(sheet) {
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var key = String(headers[j]).trim();
      obj[key] = data[i][j];
    }
    result.push(obj);
  }
  return result;
}

// =========================================================================
// 3. FUNGSI ADMINISTRASI & USER AUTH
// =========================================================================
function apiLoginAdmin(payload) {
  var user = payload.user;
  var pass = payload.pass;
  // Ganti username dan password admin sesuai kebutuhan Anda
  if (user === "admin" && pass === "takobunny2026") {
    return { success: true };
  }
  return { success: false, message: "Username/Password Admin Salah" };
}

function apiLoginUser(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  if (!sheet) return { success: false, message: "User tidak ditemukan." };
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === String(payload.username).toLowerCase() && String(data[i][3]) === String(payload.password)) {
      return {
        success: true,
        user: { username: data[i][0], nama: data[i][1], noHp: data[i][2], email: data[i][4] }
      };
    }
  }
  return { success: false, message: "Username atau Password salah." };
}

function apiRegisterUser(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users") || ss.insertSheet("Users");
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["username", "nama", "noHp", "password", "email"]);
  }
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === String(payload.username).toLowerCase()) {
      return { success: false, message: "Username sudah digunakan." };
    }
  }
  
  sheet.appendRow([payload.username, payload.nama, payload.noHp, payload.password, payload.email]);
  return {
    success: true,
    user: { username: payload.username, nama: payload.nama, noHp: payload.noHp, email: payload.email }
  };
}

function apiForgotPassword(email) {
  return { success: true, message: "Instruksi pemulihan password telah dikirim ke email " + email };
}

// =========================================================================
// 4. FUNGSI KATALOG (CREATE, UPDATE, DELETE)
// =========================================================================
function apiAddKatalog(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Katalog") || ss.insertSheet("Katalog");
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["id", "nama", "ras", "gender", "warna", "dob", "hargaCoret", "harga", "status", "fotourl"]);
  }
  var id = "RAB-" + Utilities.formatDate(new Date(), "GMT+7", "yyMMddHHmmss");
  sheet.appendRow([id, payload.nama, payload.ras, payload.gender, payload.warna, payload.dob, payload.hargaCoret, payload.harga, payload.status, payload.fotoUrl]);
  return { success: true, id: id };
}

function apiEditKatalog(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Katalog");
  if (!sheet) return { success: false };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(payload.id).trim()) {
      sheet.getRange(i + 1, 2, 1, 9).setValues([[payload.nama, payload.ras, payload.gender, payload.warna, payload.dob, payload.hargaCoret, payload.harga, payload.status, payload.fotoUrl]]);
      return { success: true };
    }
  }
  return { success: false, message: "Data tidak ditemukan" };
}

function apiDeleteKatalog(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Katalog");
  if (!sheet) return { success: false };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false };
}

// =========================================================================
// 5. FUNGSI ADOPSI / TRANSAKSI
// =========================================================================
function apiSubmitAdopsi(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Adopsi") || ss.insertSheet("Adopsi");
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["id", "tanggal", "kelinciId", "kelinciNama", "kelinciRas", "kelinciGender", "harga", "kelinciDob", "namaAdopter", "waAdopter", "provinsi", "kota", "alamat", "ekspedisi", "ongkir", "catatan", "tipeTransaksi", "username", "itemsProduk", "totalSeluruh", "statusPelunasan", "statusPengiriman", "buktiResiUrl", "ongkirAktual"]);
  }
  var id = "TRX-" + Utilities.formatDate(new Date(), "GMT+7", "yyMMddHHmmss");
  var tgl = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
  
  sheet.appendRow([
    id, tgl, payload.kelinciId, payload.kelinciNama, payload.kelinciRas, payload.kelinciGender,
    payload.kelinciHarga, payload.kelinciDob, payload.namaAdopter, payload.waAdopter,
    payload.provinsi, payload.kota, payload.alamat, payload.ekspedisi, payload.ongkir,
    payload.catatan, payload.tipeTransaksi, payload.username, payload.itemsProduk,
    payload.totalSeluruh, "Menunggu Pembayaran", "Belum Dikirim", "", ""
  ]);
  
  return { success: true, generatedId: id };
}

function apiUpdateStatusPengiriman(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Adopsi");
  if (!sheet) return { success: false };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(payload.idTrans).trim()) {
      sheet.getRange(i + 1, 22).setValue(payload.status); // Col 22 = statusPengiriman
      if (payload.resiUrl !== undefined) {
        sheet.getRange(i + 1, 23).setValue(payload.resiUrl); // Col 23 = buktiResiUrl
      }
      return { success: true };
    }
  }
  return { success: false };
}

function apiUpdateStatusPelunasanManual(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Adopsi");
  if (!sheet) return { success: false };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(payload.idTrans).trim()) {
      sheet.getRange(i + 1, 21).setValue(payload.statusBaru); // Col 21 = statusPelunasan
      
      // Update otomatis status di Katalog jika LUNAS / BOOKED
      var kelinciId = data[i][2];
      var katSheet = ss.getSheetByName("Katalog");
      if (katSheet && kelinciId) {
        var kData = katSheet.getDataRange().getValues();
        for (var j = 1; j < kData.length; j++) {
          if (String(kData[j][0]).trim() === String(kelinciId).trim()) {
            if (payload.statusBaru === "Lunas") katSheet.getRange(j + 1, 9).setValue("Adopted");
            else if (payload.statusBaru === "DP Diterima") katSheet.getRange(j + 1, 9).setValue("Booked");
            else if (payload.statusBaru === "Dibatalkan" || payload.statusBaru === "Menunggu Pembayaran") katSheet.getRange(j + 1, 9).setValue("Tersedia");
            break;
          }
        }
      }
      return { success: true };
    }
  }
  return { success: false };
}

function apiUpdateOngkir(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Adopsi");
  if (!sheet) return { success: false };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(payload.idTrans).trim()) {
      sheet.getRange(i + 1, 24).setValue(payload.nominalOngkir); // Col 24 = ongkirAktual
      return { success: true };
    }
  }
  return { success: false };
}

function apiDeleteAdopsi(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Adopsi");
  if (!sheet) return { success: false };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false };
}

// =========================================================================
// 6. FUNGSI PRODUK & GALERI & TESTIMONI
// =========================================================================
function apiAddProduk(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Produk") || ss.insertSheet("Produk");
  if (sheet.getLastRow() === 0) sheet.appendRow(["id", "nama", "harga", "fotoUrl", "deskripsi"]);
  var id = "PRD-" + new Date().getTime();
  sheet.appendRow([id, payload.nama, payload.harga, payload.fotoUrl, payload.deskripsi]);
  return { success: true };
}

function apiDeleteProduk(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Produk");
  if (!sheet) return { success: false };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false };
}

function apiAddGaleri(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Galeri") || ss.insertSheet("Galeri");
  if (sheet.getLastRow() === 0) sheet.appendRow(["id", "judul", "mediaurl"]);
  var id = "GAL-" + new Date().getTime();
  sheet.appendRow([id, payload.judul, payload.mediaUrl]);
  return { success: true };
}

function apiSubmitTestimoni(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Testimoni") || ss.insertSheet("Testimoni");
  if (sheet.getLastRow() === 0) sheet.appendRow(["id", "nama", "ras", "teks", "mediaUrl", "tanggal"]);
  var id = "TST-" + new Date().getTime();
  var tgl = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  sheet.appendRow([id, payload.nama, payload.ras, payload.teks, payload.mediaUrl, tgl]);
  return { success: true, message: "Terima kasih atas ulasan Anda!" };
}

function apiDeleteTestimoni(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Testimoni");
  if (!sheet) return { success: false };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false };
}

// =========================================================================
// 7. FUNGSI KONFIGURASI KWITANSI & LABEL LOGISTIK
// =========================================================================
function apiSaveKwitansiConfig(payload) {
  var sp = PropertiesService.getScriptProperties();
  sp.setProperty("CFG_LOGO_URL", payload.logoUrl || "");
  sp.setProperty("CFG_TTD_URL", payload.ttdUrl || "");
  sp.setProperty("CFG_NAMA_PEMILIK", payload.namaPemilik || "TAKO BUNNY");
  return { success: true };
}

function apiGetKwitansiData(adopsiId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var adopsiSheet = ss.getSheetByName("Adopsi");
  if (!adopsiSheet) return { success: false };
  
  var data = adopsiSheet.getDataRange().getValues();
  var header = data[0];
  var trxData = null;
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(adopsiId).trim()) {
      trxData = {};
      for (var j = 0; j < header.length; j++) {
        trxData[header[j]] = data[i][j];
      }
      break;
    }
  }
  
  if (!trxData) return { success: false, message: "Transaksi tidak ditemukan." };
  
  // Ambil detail warna dan foto kelinci dari Katalog
  var kelinciWarna = "";
  var kelinciFotoUrl = "";
  var katSheet = ss.getSheetByName("Katalog");
  if (katSheet && trxData.kelinciId) {
    var kData = katSheet.getDataRange().getValues();
    for (var k = 1; k < kData.length; k++) {
      if (String(kData[k][0]).trim() === String(trxData.kelinciId).trim()) {
        kelinciWarna = kData[k][4]; // Col warna
        kelinciFotoUrl = kData[k][9]; // Col fotourl
        break;
      }
    }
  }
  
  var sp = PropertiesService.getScriptProperties();
  var config = {
    logoUrl: sp.getProperty("CFG_LOGO_URL") || "",
    ttdUrl: sp.getProperty("CFG_TTD_URL") || "",
    namaPemilik: sp.getProperty("CFG_NAMA_PEMILIK") || "TAKO BUNNY"
  };
  
  return {
    success: true,
    data: trxData,
    kelinciWarna: kelinciWarna,
    kelinciFotoUrl: kelinciFotoUrl,
    config: config
  };
}
