function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('TAKO BUNNY - Peternakan Kelinci Hias')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function apiUploadFile(dataUrl, filename) {
  try {
    var folderId = "1VJb5Iy_fdD9U4DiQoKqTDgApTlI0eKpW";
    var folder = DriveApp.getFolderById(folderId);
    var splitData = dataUrl.split(',');
    var base64Str = splitData[1];
    var mimeType = splitData[0].split(':')[1].split(';')[0];
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Str), mimeType, filename);
    var file = folder.createFile(blob);
    var directUrl = "https://lh3.googleusercontent.com/d/" + file.getId();
    return { success: true, url: directUrl, message: "Upload berhasil!" };
  } catch (e) { return { success: false, message: "Error upload: " + e.toString() }; }
}

function apiGetData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetKatalog = ss.getSheetByName("Katalog"); var katalog = [];
  if (sheetKatalog) {
    var data = sheetKatalog.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        var dobVal = data[i][9] ? String(data[i][9]) : ''; if (data[i][9] instanceof Date) dobVal = Utilities.formatDate(data[i][9], "GMT+7", "yyyy-MM-dd");
        katalog.push({ id: String(data[i][0]), nama: String(data[i][1] || ''), ras: String(data[i][2] || ''), gender: String(data[i][3] || ''), warna: String(data[i][4] || ''), harga: Number(data[i][5] || 0), status: String(data[i][6] || 'Tersedia'), fotourl: String(data[i][7] || ''), hargaCoret: Number(data[i][8] || 0), dob: dobVal });
      }
    }
  }

  var sheetGaleri = ss.getSheetByName("Galeri"); var galeri = [];
  if (sheetGaleri) {
    var dataG = sheetGaleri.getDataRange().getValues();
    for (var j = 1; j < dataG.length; j++) { if (dataG[j][0] || dataG[j][1]) galeri.push({ id: String(dataG[j][0] || j), judul: String(dataG[j][1] || ''), mediaurl: String(dataG[j][2] || '') }); }
  }

  var sheetProduk = ss.getSheetByName("Produk"); var produk = [];
  if (sheetProduk) {
    var dataP = sheetProduk.getDataRange().getValues();
    for (var p = 1; p < dataP.length; p++) { if (dataP[p][0]) produk.push({ id: String(dataP[p][0]), nama: String(dataP[p][1] || ''), deskripsi: String(dataP[p][2] || ''), harga: Number(dataP[p][3] || 0), fotoUrl: String(dataP[p][4] || '') }); }
  }

  var sheetAdopsi = ss.getSheetByName("Adopsi"); var adopsi = [];
  if (sheetAdopsi) {
    var dataA = sheetAdopsi.getDataRange().getValues();
    for (var k = 1; k < dataA.length; k++) {
      if (dataA[k][0]) {
        var dobAdopsi = dataA[k][15] ? String(dataA[k][15]) : ''; if (dataA[k][15] instanceof Date) dobAdopsi = Utilities.formatDate(dataA[k][15], "GMT+7", "yyyy-MM-dd");
        adopsi.push({
          id: String(dataA[k][0]), tanggal: String(dataA[k][1] || ''), kelinciId: String(dataA[k][2] || ''), kelinciNama: String(dataA[k][3] || ''), kelinciRas: String(dataA[k][4] || ''), kelinciGender: String(dataA[k][5] || ''), harga: String(dataA[k][6] || ''), namaAdopter: String(dataA[k][7] || ''), waAdopter: String(dataA[k][8] || ''), provinsi: String(dataA[k][9] || ''), kota: String(dataA[k][10] || ''), alamat: String(dataA[k][11] || ''), ekspedisi: String(dataA[k][12] || ''), ongkir: String(dataA[k][13] || ''), catatan: String(dataA[k][14] || ''), dob: dobAdopsi, tipeTransaksi: String(dataA[k][16] || 'Adopsi'), username: String(dataA[k][17] || ''), statusPelunasan: String(dataA[k][18] || 'Menunggu Pembayaran'), tglPelunasan: String(dataA[k][19] || ''), itemsProduk: String(dataA[k][20] || ''), totalSeluruh: String(dataA[k][21] || dataA[k][6] || '0'), statusPengiriman: String(dataA[k][22] || 'Belum Dikirim'), buktiResiUrl: String(dataA[k][23] || ''), ongkirAktual: String(dataA[k][24] || '0')
        });
      }
    }
  }

  var sheetTesti = ss.getSheetByName("Testimoni"); var testimoni = [];
  if (sheetTesti) {
    var dataT = sheetTesti.getDataRange().getValues();
    for (var t = 1; t < dataT.length; t++) { if (dataT[t][0]) testimoni.push({ id: String(dataT[t][0]), nama: String(dataT[t][1] || ''), ras: String(dataT[t][2] || ''), teks: String(dataT[t][3] || ''), mediaUrl: String(dataT[t][4] || ''), tanggal: String(dataT[t][5] || '') }); }
  }

  return { katalog: katalog, galeri: galeri, produk: produk, adopsi: adopsi, testimoni: testimoni };
}

function apiSubmitTestimoni(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("Testimoni");
  if (!sheet) { sheet = ss.insertSheet("Testimoni"); sheet.appendRow(["ID", "Nama Adopter", "Ras Kelinci", "Ulasan", "MediaURL", "Tanggal"]); }
  var newId = "TST-" + Utilities.formatDate(new Date(), "GMT+7", "yyMMddHHmmss"); var tglStr = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy");
  sheet.appendRow([newId, payload.nama, payload.ras, payload.teks, payload.mediaUrl || '', tglStr]);
  return { success: true, message: "Ulasan berhasil dikirim! Terima kasih." };
}

function apiDeleteTestimoni(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("Testimoni");
  if (!sheet) return { success: false, message: "Sheet Testimoni tidak ditemukan" };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { if (String(data[i][0]).trim() === String(id).trim()) { sheet.deleteRow(i + 1); return { success: true, message: "Testimoni berhasil dihapus!" }; } }
  return { success: false, message: "Data tidak ditemukan!" };
}

function apiSubmitAdopsi(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheetAdopsi = ss.getSheetByName("Adopsi");
  if (!sheetAdopsi) {
    sheetAdopsi = ss.insertSheet("Adopsi");
    sheetAdopsi.appendRow(["ID Transaksi", "Tanggal", "ID Kelinci", "Nama Kelinci", "Ras", "Gender", "Harga Kelinci", "Nama Adopter", "WA Adopter", "Provinsi", "Kota", "Alamat Lengkap", "Ekspedisi", "Estimasi Ongkir", "Catatan", "DOB", "Tipe Transaksi", "Username", "Status Pelunasan", "Tanggal Pelunasan", "Items Produk", "Total Seluruh", "Status Pengiriman", "Bukti Resi URL", "ongkirAktual"]);
  }
  
  var statusPrefix = (payload.tipeTransaksi === "Booking") ? "BK" : "TB";
  var rasClean = (payload.kelinciRas || "").toLowerCase(); var rasCode = "LAIN";
  if (rasClean.indexOf("holland") !== -1) rasCode = "HL"; else if (rasClean.indexOf("netherland") !== -1 || rasClean.indexOf("nd") !== -1) rasCode = "ND";

  var now = new Date(); var dateStr = Utilities.formatDate(now, "GMT+7", "ddMMyy"); var tglFullStr = Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy HH:mm");
  var sequenceNumber = sheetAdopsi.getLastRow(); var seqFormatted = ("00" + sequenceNumber).slice(-3);
  var newId = statusPrefix + "-" + rasCode + "-" + dateStr + "-" + seqFormatted;
  var initialStatusPelunasan = "Menunggu Pembayaran"; 

  sheetAdopsi.appendRow([
    newId, tglFullStr, payload.kelinciId, payload.kelinciNama, payload.kelinciRas, payload.kelinciGender, payload.kelinciHarga, payload.namaAdopter, payload.waAdopter, payload.provinsi, payload.kota, payload.alamat, payload.ekspedisi, payload.ongkir, payload.catatan, payload.kelinciDob || '', payload.tipeTransaksi || 'Adopsi', payload.username || '', initialStatusPelunasan, "", payload.itemsProduk || '', payload.totalSeluruh || payload.kelinciHarga, "Belum Dikirim", "", "0"
  ]);

  return { success: true, message: "Data pendaftaran berhasil disimpan!", generatedId: newId };
}

function apiAddProduk(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("Produk");
  if (!sheet) { sheet = ss.insertSheet("Produk"); sheet.appendRow(["ID", "Nama Produk", "Deskripsi", "Harga", "FotoURL"]); }
  var newId = "PRD-" + Utilities.formatDate(new Date(), "GMT+7", "yyMMddHHmmss"); sheet.appendRow([newId, payload.nama, payload.deskripsi, Number(payload.harga || 0), payload.fotoUrl]);
  return { success: true, message: "Produk perlengkapan berhasil ditambahkan!" };
}

function apiDeleteProduk(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("Produk");
  if (!sheet) return { success: false, message: "Sheet Produk tidak ditemukan" };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { if (String(data[i][0]).trim() === String(id).trim()) { sheet.deleteRow(i + 1); return { success: true, message: "Produk berhasil dihapus!" }; } }
  return { success: false, message: "Data produk tidak ditemukan!" };
}

function apiUpdateStatusPengiriman(idTrans, status, resiUrl) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("Adopsi");
  if (!sheet) return { success: false, message: "Sheet Adopsi tidak ditemukan" };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(idTrans).trim()) {
      sheet.getRange(i + 1, 23).setValue(status); if (resiUrl !== undefined) sheet.getRange(i + 1, 24).setValue(resiUrl);
      return { success: true, message: "Status pengiriman berhasil diperbarui!" };
    }
  }
  return { success: false, message: "Transaksi tidak ditemukan!" };
}

function apiUpdateStatusPelunasanManual(idTrans, statusBaru) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("Adopsi");
  if (!sheet) return { success: false, message: "Sheet Adopsi tidak ditemukan" };
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(idTrans).trim()) {
      var tglLunas = (statusBaru === "Lunas" || statusBaru === "DP Diterima") ? Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm") : "";
      sheet.getRange(i + 1, 19).setValue(statusBaru); 
      sheet.getRange(i + 1, 20).setValue(tglLunas);
      
      var kelinciId = data[i][2]; 
      var sheetKatalog = ss.getSheetByName("Katalog");
      if (sheetKatalog && kelinciId) {
        var dataK = sheetKatalog.getDataRange().getValues();
        for (var j = 1; j < dataK.length; j++) {
          if (String(dataK[j][0]).trim() === String(kelinciId).trim()) { 
            var statusKatalogBaru = "";
            if (statusBaru === "Lunas") statusKatalogBaru = "Adopted";
            else if (statusBaru === "DP Diterima") statusKatalogBaru = "Booked";
            else if (statusBaru === "Menunggu Pembayaran" || statusBaru === "Dibatalkan") statusKatalogBaru = "Tersedia";

            if (statusKatalogBaru !== "") {
              sheetKatalog.getRange(j + 1, 7).setValue(statusKatalogBaru); 
            }
            break; 
          }
        }
      }
      return { success: true, message: "Status Pelunasan diubah menjadi: " + statusBaru };
    }
  }
  return { success: false, message: "Transaksi tidak ditemukan!" };
}

function apiRegisterUser(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("Users");
  if (!sheet) { sheet = ss.insertSheet("Users"); sheet.appendRow(["Username", "Nama Lengkap", "No HP", "Password", "Tanggal Daftar", "Email"]); }
  var data = sheet.getDataRange().getValues(); var userClean = String(payload.username).trim().toLowerCase(); var emailClean = payload.email ? String(payload.email).trim().toLowerCase() : "";
  for (var i = 1; i < data.length; i++) { if (String(data[i][0]).trim().toLowerCase() === userClean) return { success: false, message: "Username sudah terdaftar!" }; }
  var tglStr = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm"); sheet.appendRow([userClean, payload.nama, payload.noHp, payload.password, tglStr, emailClean]);
  return { success: true, message: "Pendaftaran berhasil!", user: { username: userClean, nama: payload.nama, noHp: payload.noHp, email: emailClean } };
}

function apiLoginUser(username, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("Users");
  if (!sheet) return { success: false, message: "Belum ada user terdaftar." };
  var data = sheet.getDataRange().getValues(); var userClean = String(username).trim().toLowerCase();
  for (var i = 1; i < data.length; i++) { if (String(data[i][0]).trim().toLowerCase() === userClean && String(data[i][3]) === String(password)) { return { success: true, message: "Login berhasil!", user: { username: userClean, nama: String(data[i][1] || ''), noHp: String(data[i][2] || '') } }; } }
  return { success: false, message: "Username atau Password salah!" };
}

function apiForgotPassword(email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("Users");
  if (!sheet) return { success: false, message: "Belum ada user terdaftar." };
  var data = sheet.getDataRange().getValues(); var emailClean = String(email).trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (data[i][5] && String(data[i][5]).trim().toLowerCase() === emailClean) {
      try { MailApp.sendEmail(emailClean, "TAKO BUNNY - Pemulihan Akun", "Halo " + data[i][1] + ",\n\nBerikut adalah informasi login Anda:\n\nUsername: " + data[i][0] + "\nPassword: " + data[i][3] + "\n\nHarap simpan informasi ini dengan aman.\n\nSalam,\nTAKO BUNNY."); return { success: true, message: "Informasi login telah dikirim ke email Anda." }; } catch (e) { return { success: false, message: "Error sistem: " + e.toString() }; }
    }
  } return { success: false, message: "Email tidak ditemukan dalam sistem." };
}

function apiLoginAdmin(user, pass) { if (user === "admin" && pass === "takobunny123") return { success: true, message: "Login Berhasil" }; else return { success: false, message: "Username atau Password Salah!" }; }

function apiAddKatalog(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("Katalog");
  if (!sheet) { sheet = ss.insertSheet("Katalog"); sheet.appendRow(["ID", "Nama", "Ras", "Gender", "Warna", "Harga", "Status", "FotoURL", "HargaCoret", "DOB"]); }
  var newId = "RAB-" + Utilities.formatDate(new Date(), "GMT+7", "yyMMddHHmmss"); sheet.appendRow([newId, payload.nama, payload.ras, payload.gender, payload.warna, Number(payload.harga || 0), payload.status, payload.fotoUrl, Number(payload.hargaCoret || 0), payload.dob || '']);
  return { success: true, message: "Kelinci berhasil ditambahkan ke katalog!" };
}

function apiEditKatalog(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("Katalog");
  if (!sheet) return { success: false, message: "Sheet Katalog tidak ditemukan" };
  var data = sheet.getDataRange().getValues(); var targetRow = -1;
  for (var i = 1; i < data.length; i++) { if (String(data[i][0]).trim() === String(payload.id).trim()) { targetRow = i + 1; break; } }
  if (targetRow === -1) return { success: false, message: "Data kelinci tidak ditemukan!" };
  sheet.getRange(targetRow, 2).setValue(payload.nama); sheet.getRange(targetRow, 3).setValue(payload.ras); sheet.getRange(targetRow, 4).setValue(payload.gender); sheet.getRange(targetRow, 5).setValue(payload.warna); sheet.getRange(targetRow, 6).setValue(Number(payload.harga || 0)); sheet.getRange(targetRow, 7).setValue(payload.status); sheet.getRange(targetRow, 8).setValue(payload.fotoUrl); sheet.getRange(targetRow, 9).setValue(Number(payload.hargaCoret || 0)); sheet.getRange(targetRow, 10).setValue(payload.dob || '');
  return { success: true, message: "Data kelinci berhasil diperbarui!" };
}

function apiDeleteKatalog(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("Katalog");
  if (!sheet) return { success: false, message: "Sheet Katalog tidak ditemukan" };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { if (String(data[i][0]).trim() === String(id).trim()) { sheet.deleteRow(i + 1); return { success: true, message: "Kelinci berhasil dihapus!" }; } } return { success: false, message: "Data tidak ditemukan!" };
}

function apiAddGaleri(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("Galeri");
  if (!sheet) { sheet = ss.insertSheet("Galeri"); sheet.appendRow(["ID", "Judul", "MediaURL"]); }
  var newId = "GAL-" + Utilities.formatDate(new Date(), "GMT+7", "yyMMddHHmmss"); sheet.appendRow([newId, payload.judul, payload.mediaUrl]);
  return { success: true, message: "Galeri berhasil ditambahkan!" };
}

function apiDeleteAdopsi(idTrans) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("Adopsi");
  if (!sheet) return { success: false, message: "Sheet Adopsi tidak ditemukan" };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { if (String(data[i][0]).trim() === String(idTrans).trim()) { sheet.deleteRow(i + 1); return { success: true, message: "Riwayat berhasil dihapus!" }; } } return { success: false, message: "Data tidak ditemukan!" };
}

function apiSaveKwitansiConfig(logoUrl, ttdUrl, namaPemilik) {
  var props = PropertiesService.getScriptProperties(); props.setProperty('CONFIG_LOGO_URL', logoUrl || ''); props.setProperty('CONFIG_TTD_URL', ttdUrl || ''); props.setProperty('CONFIG_NAMA_PEMILIK', namaPemilik || 'TAKO BUNNY');
  return { success: true, message: "Pengaturan Kwitansi berhasil disimpan!" };
}

function apiGetKwitansiData(adopsiId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheetAdopsi = ss.getSheetByName("Adopsi"); var sheetKatalog = ss.getSheetByName("Katalog");
  if (!sheetAdopsi) return { success: false, message: "Data tidak ditemukan." };
  var dataA = sheetAdopsi.getDataRange().getValues(); var itemA = null;
  for (var i = 1; i < dataA.length; i++) {
    if (String(dataA[i][0]).trim() === String(adopsiId).trim()) {
      var dobVal = dataA[i][15] ? String(dataA[i][15]) : ''; if (dataA[i][15] instanceof Date) dobVal = Utilities.formatDate(dataA[i][15], "GMT+7", "yyyy-MM-dd");
      itemA = { id: dataA[i][0], tanggal: dataA[i][1], kelinciId: dataA[i][2], kelinciNama: dataA[i][3], kelinciRas: dataA[i][4], kelinciGender: dataA[i][5], harga: dataA[i][6], namaAdopter: dataA[i][7], waAdopter: dataA[i][8], provinsi: dataA[i][9], kota: dataA[i][10], alamat: dataA[i][11], ekspedisi: dataA[i][12], ongkir: dataA[i][13], catatan: dataA[i][14], dob: dobVal, tipeTransaksi: String(dataA[i][16] || 'Adopsi'), username: String(dataA[i][17] || ''), statusPelunasan: String(dataA[i][18] || 'Menunggu Pembayaran'), tglPelunasan: String(dataA[i][19] || ''), itemsProduk: String(dataA[i][20] || ''), totalSeluruh: String(dataA[i][21] || dataA[k][6] || '0'), ongkirAktual: String(dataA[i][24] || '0') }; break;
    }
  }
  if (!itemA) return { success: false, message: "Transaksi tidak ditemukan." };
  var kelinciFotoUrl = ""; var kelinciWarna = "";
  if (sheetKatalog) {
    var dataK = sheetKatalog.getDataRange().getValues();
    for (var j = 1; j < dataK.length; j++) {
      if (String(dataK[j][0]).trim() === String(itemA.kelinciId).trim()) { kelinciWarna = dataK[j][4] || ""; kelinciFotoUrl = dataK[j][7] || ""; if (!itemA.dob && dataK[j][9]) itemA.dob = dataK[j][9] instanceof Date ? Utilities.formatDate(dataK[j][9], "GMT+7", "yyyy-MM-dd") : String(dataK[j][9]); break; }
    }
  }
  var props = PropertiesService.getScriptProperties(); var config = { logoUrl: props.getProperty('CONFIG_LOGO_URL') || '', ttdUrl: props.getProperty('CONFIG_TTD_URL') || '', namaPemilik: props.getProperty('CONFIG_NAMA_PEMILIK') || 'TAKO BUNNY' }; return { success: true, data: itemA, kelinciFotoUrl: kelinciFotoUrl, kelinciWarna: kelinciWarna, config: config };
}

function apiUpdateOngkir(idTrans, nominalOngkir) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("Adopsi");
    if (!sheet) return { success: false, message: "Sheet Adopsi tidak ditemukan" };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(idTrans).trim()) {
        sheet.getRange(i + 1, 25).setValue(nominalOngkir);
        return { success: true, message: "Ongkir berhasil diupdate" };
      }
    }
    return { success: false, message: "Transaksi tidak ditemukan" };
  } catch (e) { return { success: false, message: e.message }; }
}

// ====================================================================================
// FUNGSI BARU UNTUK MENERIMA REQUEST DARI GITHUB LALU MENGIRIM BALIK DATA SEBAGAI JSON
// ====================================================================================
function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    var act = req.action;
    var pay = req.payload;
    var res = {};

    if (act === 'apiUploadFile') res = apiUploadFile(pay.dataUrl, pay.filename);
    else if (act === 'apiGetData') res = apiGetData();
    else if (act === 'apiSubmitTestimoni') res = apiSubmitTestimoni(pay);
    else if (act === 'apiDeleteTestimoni') res = apiDeleteTestimoni(pay);
    else if (act === 'apiSubmitAdopsi') res = apiSubmitAdopsi(pay);
    else if (act === 'apiAddProduk') res = apiAddProduk(pay);
    else if (act === 'apiDeleteProduk') res = apiDeleteProduk(pay);
    else if (act === 'apiUpdateStatusPengiriman') res = apiUpdateStatusPengiriman(pay.idTrans, pay.status, pay.resiUrl);
    else if (act === 'apiUpdateStatusPelunasanManual') res = apiUpdateStatusPelunasanManual(pay.idTrans, pay.statusBaru);
    else if (act === 'apiRegisterUser') res = apiRegisterUser(pay);
    else if (act === 'apiLoginUser') res = apiLoginUser(pay.username, pay.password);
    else if (act === 'apiForgotPassword') res = apiForgotPassword(pay);
    else if (act === 'apiLoginAdmin') res = apiLoginAdmin(pay.user, pay.pass);
    else if (act === 'apiAddKatalog') res = apiAddKatalog(pay);
    else if (act === 'apiEditKatalog') res = apiEditKatalog(pay);
    else if (act === 'apiDeleteKatalog') res = apiDeleteKatalog(pay);
    else if (act === 'apiAddGaleri') res = apiAddGaleri(pay);
    else if (act === 'apiDeleteAdopsi') res = apiDeleteAdopsi(pay);
    else if (act === 'apiSaveKwitansiConfig') res = apiSaveKwitansiConfig(pay.logoUrl, pay.ttdUrl, pay.namaPemilik);
    else if (act === 'apiGetKwitansiData') res = apiGetKwitansiData(pay);
    else if (act === 'apiUpdateOngkir') res = apiUpdateOngkir(pay.idTrans, pay.nominalOngkir);

    return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({success: false, message: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}