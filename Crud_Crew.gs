function getCrewData() {
  const ss = getDB();
  const sheet = ss.getSheetByName('Crew');
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return []; 
  return data.slice(1);
}

function hitungUsia(tglLahir) {
  if(!tglLahir) return "-";
  var dob = new Date(tglLahir);
  var today = new Date();
  var age = today.getFullYear() - dob.getFullYear();
  var m = today.getMonth() - dob.getMonth();
  if (m < 0 || (m === 0 && today.getDate() < dob.getDate())) { age--; }
  return age + " Tahun";
}

function deleteCrewData(kodeCrew) {
  const ss = getDB();
  const sheetCrew = ss.getSheetByName('Crew');
  if (!sheetCrew) throw new Error("Sheet Crew tidak ditemukan.");
  
  // 1. Cari baris kru di Sheet Utama
  const tf = sheetCrew.getRange("A:A").createTextFinder(kodeCrew.toString()).matchEntireCell(true).findNext();
  if(!tf) throw new Error("Data kru tidak ditemukan untuk dihapus.");
  
  const rowIndex = tf.getRow();
  const namaKru = sheetCrew.getRange(rowIndex, 2).getValue(); // Ambil nama untuk log

  // 2. HAPUS DATA RELASI (Keluarga, Dokumen, Pengalaman)
  // Kode crew ada di Kolom B (kolom ke-2) pada ketiga sheet ini
  const sheetsToClean = ['Crew_Fam', 'Crew_Doc', 'Crew_SeaServices'];
  
  sheetsToClean.forEach(sheetName => {
    let sheetRelated = ss.getSheetByName(sheetName);
    if (sheetRelated) {
      // Cari semua baris yang mengandung kode_crew anak
      let finds = sheetRelated.getRange("B:B").createTextFinder(kodeCrew.toString()).matchEntireCell(true).findAll();
      
      // Kunci Penting: Urutkan baris dari BAWAH ke ATAS agar saat dihapus, index baris tidak bergeser/kacau
      let rowsToDelete = finds.map(cell => cell.getRow()).sort((a, b) => b - a);
      
      rowsToDelete.forEach(r => sheetRelated.deleteRow(r));
    }
  });

  // 3. HAPUS FOLDER DRIVE (Bersih-bersih Kuota)
  try {
    const mainFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    // Cari folder yang mengandung kode crew tersebut
    const folders = mainFolder.searchFolders("title contains '" + kodeCrew.toString() + "'");
    while (folders.hasNext()) {
      folders.next().setTrashed(true); // Pindahkan ke kotak sampah Drive
    }
  } catch(e) {
    console.error("Gagal menghapus folder drive: " + e.message);
  }

  // 4. Hapus data utama di sheet Crew
  sheetCrew.deleteRow(rowIndex);
  
  // 5. Catat Log
  catatLogPerubahan(kodeCrew, namaKru, "Database Utama", "Menghapus profil kru dan seluruh data relasinya secara permanen dari sistem");
  
  return "Data kru " + namaKru + " beserta dokumen dan riwayatnya berhasil dihapus bersih!";
}

function updateCrewData(formObject) {
  const sheet = getDB().getSheetByName('Crew');
  
  // OPTIMASI: TextFinder untuk mencari baris
  const tf = sheet.getRange("A:A").createTextFinder(formObject.edit_kode_crew.toString()).matchEntireCell(true).findNext();
  if (!tf) throw new Error("Data kru tidak ditemukan.");
  
  let rowIndex = tf.getRow();
  let dataBaris = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  let linkKtp = dataBaris[9];
  let linkFoto = dataBaris[10];
  let linkTtd = dataBaris[11];

  const mainFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  let targetFolder = mainFolder;
  const searchString = "title contains '" + formObject.edit_kode_crew + "' and trashed = false";
  const folderIterator = mainFolder.searchFolders(searchString);
  if (folderIterator.hasNext()) { targetFolder = folderIterator.next(); } 
  else { targetFolder = mainFolder.createFolder(formObject.edit_kode_crew + " - " + formObject.edit_nama_lengkap); }

  function gantiFileKru(base64Data, fileName, oldFileUrl) {
    if (!base64Data || !fileName) return oldFileUrl;
    if (oldFileUrl && oldFileUrl.includes("drive.google.com")) {
      try {
        const oldFileIdMatch = oldFileUrl.match(/[-\w]{25,}/);
        if (oldFileIdMatch && oldFileIdMatch[0]) DriveApp.getFileById(oldFileIdMatch[0]).setTrashed(true);
      } catch(e) {}
    }
    const contentType = base64Data.substring(5, base64Data.indexOf(';'));
    const bytes = Utilities.base64Decode(base64Data.split(',')[1]);
    const blob = Utilities.newBlob(bytes, contentType, fileName);
    const file = targetFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  }

  linkFoto = gantiFileKru(formObject.fileFotoData, formObject.fileFotoName, linkFoto);
  linkKtp = gantiFileKru(formObject.fileKtpData, formObject.fileKtpName, linkKtp);
  linkTtd = gantiFileKru(formObject.fileTtdData, formObject.fileTtdName, linkTtd);
  
  const nikAman = formObject.edit_nik ? "'" + formObject.edit_nik : '';
  const usiaOtomatis = hitungUsia(formObject.edit_tanggal_lahir);

  // OPTIMASI: Batch update per baris agar lebih cepat
  sheet.getRange(rowIndex, 2, 1, 16).setValues([[
    formObject.edit_nama_lengkap, formObject.edit_jabatan, formObject.edit_tempat_lahir,
    formObject.edit_tanggal_lahir, formObject.edit_telepon, formObject.edit_email,
    formObject.edit_asal_daerah, nikAman, linkKtp, linkFoto, linkTtd,
    formObject.edit_status, usiaOtomatis, formObject.edit_skill, dataBaris[15] || "", formObject.edit_tanggal_input
  ]]);

  catatLogPerubahan(formObject.edit_kode_crew, formObject.edit_nama_lengkap, "Profil Utama", "Melakukan update data profil/berkas");
  return "Profil " + formObject.edit_nama_lengkap + " berhasil diupdate!";
}

function simpanBulkOnboardSatuKru(p) {
 const ss = getDB();
 const sCrew = ss.getSheetByName('Crew');
 const sDoc = ss.getSheetByName('Crew_Doc');
 
 if (!sCrew || !sDoc) throw new Error("Sheet database tidak ditemukan.");

 // ==========================================
 // PERBAIKAN: CEK DUPLIKAT KODE CREW (ANTI-GANDA)
 // ==========================================
 const tf = sCrew.getRange("A:A").createTextFinder(p.kode_crew.toString()).matchEntireCell(true).findNext();
 if (tf) throw new Error("Gagal! Kode Crew '" + p.kode_crew + "' sudah terdaftar di sistem.");
 // ==========================================
 
 const rootFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
 const kruFolder = rootFolder.createFolder(p.kode_crew + " - " + p.nama_lengkap);
 
 const nikFormat = p.nik ? "'" + p.nik : "";
 const tglLahir = p.tanggal_lahir || "";
 let usiaStr = "-";
 if (tglLahir) {
  let thn = new Date().getFullYear() - new Date(tglLahir).getFullYear();
  usiaStr = thn + " Tahun";
 }

 const tglInput = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");

 sCrew.appendRow([
  p.kode_crew, p.nama_lengkap, p.jabatan || "", p.tempat_lahir || "", 
  tglLahir, p.telepon || "", p.email || "", p.asal_daerah || "",
  nikFormat, "", "", "", p.status, usiaStr, "", "", tglInput
 ]);

 // OPTIMASI: Batch Write (setValues) untuk bulk dokumen agar tidak timeout
 if (p.dokumen && p.dokumen.length > 0) {
  let barisDokumenBaru = [];
  p.dokumen.forEach(doc => {
   let fileUrl = "";
   if (doc.fileData && doc.fileName) {
    try {
     const splitData = doc.fileData.split(',');
     const contentType = splitData[0].substring(5, splitData[0].indexOf(';'));
     const bytes = Utilities.base64Decode(splitData[1]);
     const blob = Utilities.newBlob(bytes, contentType, doc.fileName);
     const file = kruFolder.createFile(blob);
     file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
     fileUrl = file.getUrl();
    } catch(e) { console.error("Gagal Upload: " + doc.fileName); }
   }
   const docID = "DOC-" + new Date().getTime() + "-" + Math.floor(Math.random() * 100);
   barisDokumenBaru.push([
    docID, p.kode_crew, doc.tipe || "Dokumen", doc.nomor || "", 
    doc.tempat || "", doc.issued || "", doc.expiry || "", fileUrl
   ]);
  });
  
 if (barisDokumenBaru.length > 0) {
  sDoc.getRange(sDoc.getLastRow() + 1, 1, barisDokumenBaru.length, barisDokumenBaru[0].length).setValues(barisDokumenBaru);
  }
 }
 // Catat ke Log
 catatLogPerubahan(p.kode_crew, p.nama_lengkap, "Database Utama", "Mendaftarkan kru baru via fitur Bulk Upload");
 return true;
}

function simpanKruManual(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sCrew = ss.getSheetByName('Crew');
  
  // OPTIMASI: TextFinder Duplicate Check
  const tf = sCrew.getRange("A:A").createTextFinder(p.kode_crew).matchEntireCell(true).findNext();
  if (tf) throw new Error("Kode Crew '" + p.kode_crew + "' sudah terdaftar!");

  const rootFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const kruFolder = rootFolder.createFolder(p.kode_crew + " - " + p.nama_lengkap);
  
  const uploadKeDrive = (fileData, fileName) => {
    if (!fileData || !fileName) return "";
    try {
      let splitData = fileData.split(',');
      let contentType = splitData[0].substring(5, splitData[0].indexOf(';'));
      let bytes = Utilities.base64Decode(splitData[1]);
      let blob = Utilities.newBlob(bytes, contentType, fileName);
      let file = kruFolder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return file.getUrl();
    } catch (e) { return ""; }
  };

  const urlFoto = uploadKeDrive(p.fileFotoData, p.fileFotoName);
  const urlKtp = uploadKeDrive(p.fileKtpData, p.fileKtpName);
  const urlTtd = uploadKeDrive(p.fileTtdData, p.fileTtdName);

  let usiaStr = "-";
  if (p.tanggal_lahir) {
    let thn = new Date().getFullYear() - new Date(p.tanggal_lahir).getFullYear();
    usiaStr = thn + " Tahun";
  }

  const nikFormat = p.nik ? "'" + p.nik : "";
  const tglInput = p.tanggal_input || Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");

  sCrew.appendRow([
    p.kode_crew, p.nama_lengkap, p.jabatan, p.tempat_lahir, 
    p.tanggal_lahir, p.telepon, p.email, p.asal_daerah,
    nikFormat, urlKtp, urlFoto, urlTtd, p.status, usiaStr, p.skill || "", "", tglInput
  ]);
  catatLogPerubahan(p.kode_crew, p.nama_lengkap, "Database Utama", "Mendaftarkan kru baru ke sistem secara manual");
  return "Data Kru " + p.nama_lengkap + " berhasil ditambahkan beserta berkasnya!";
}

function prosesBlacklistBackend(kode, alasan) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Crew');
  
  const tf = sheet.getRange("A:A").createTextFinder(kode.toString()).matchEntireCell(true).findNext();
  if(tf) {
      let rowIndex = tf.getRow();
      const namaKru = sheet.getRange(rowIndex, 2).getValue();
      const folderAsal = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      const folderTujuan = DriveApp.getFolderById(BLACKLIST_FOLDER_ID);
      const listFolder = folderAsal.getFoldersByName(kode + " - " + namaKru);
      if (listFolder.hasNext()) listFolder.next().moveTo(folderTujuan);

      sheet.getRange(rowIndex, 13).setValue("Blacklisted");
      sheet.getRange(rowIndex, 16).setValue(alasan);
      catatLogPerubahan(kode, namaKru, "Status Kru", "Memasukkan kru ke daftar Blacklist. Alasan: " + alasan);
      return "Sukses! " + namaKru + " masuk daftar Blacklist.";
  }
}

function pulihkanKruBackend(kode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Crew');
  
  const tf = sheet.getRange("A:A").createTextFinder(kode.toString()).matchEntireCell(true).findNext();
  if(tf) {
      let rowIndex = tf.getRow();
      const namaKru = sheet.getRange(rowIndex, 2).getValue();
      const folderAsal = DriveApp.getFolderById(BLACKLIST_FOLDER_ID);
      const folderTujuan = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      const listFolder = folderAsal.getFoldersByName(kode + " - " + namaKru);
      if (listFolder.hasNext()) listFolder.next().moveTo(folderTujuan);

      sheet.getRange(rowIndex, 13).setValue("Standby");
      sheet.getRange(rowIndex, 16).setValue("");
      catatLogPerubahan(kode, namaKru, "Status Kru", "Memulihkan status kru dari Blacklist menjadi Standby");
      return "Kru " + namaKru + " berhasil dipulihkan.";
  }
}

function getSeaServicesSummary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Crew_SeaServices');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  let summary = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][1]) summary.push([data[i][1], data[i][10]]);
  }
  return summary;
}

function getDocsSummary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Crew_Doc');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  let summary = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][1]) summary.push([data[i][1].toString().trim(), data[i][2].toString().trim().toUpperCase()]);
  }
  return summary;
}

// ==========================================
// OPTIMASI DATABASE CREW: BATCHING FETCH
// Ditaruh di paling bawah Crud_Crew.gs
// ==========================================
function getCrewListMasterData() {
  const ss = getDB();
  
  // 1. Ambil Data Crew Utama
  const sheetCrew = ss.getSheetByName('Crew');
  let crewData = [];
  if (sheetCrew) {
    let data = sheetCrew.getDataRange().getDisplayValues();
    if (data.length > 1) crewData = data.slice(1);
  }

  // 2. Ambil Summary Pengalaman Kapal (Sea Services)
  const sheetSea = ss.getSheetByName('Crew_SeaServices');
  let seaSummary = [];
  if (sheetSea) {
    let data = sheetSea.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1]) seaSummary.push([data[i][1], data[i][10]]);
    }
  }

  // 3. Ambil Summary Dokumen
  const sheetDoc = ss.getSheetByName('Crew_Doc');
  let docSummary = [];
  if (sheetDoc) {
    let data = sheetDoc.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1]) docSummary.push([data[i][1].toString().trim(), data[i][2].toString().trim().toUpperCase()]);
    }
  }

  // 4. Ambil Dropdown Jabatan & Status (Memanfaatkan cache dari Dropdown_Leader.gs)
  let jabatanList = [];
  let statusList = [];
  try { jabatanList = getJabatanList(); } catch(e){}
  try { statusList = getStatusList(); } catch(e){}

  // Kirim semua dalam satu kotak paket
  return {
    crewData: crewData,
    seaSummary: seaSummary,
    docSummary: docSummary,
    jabatanList: jabatanList,
    statusList: statusList
  };
}