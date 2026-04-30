function getCrewProfile(kodeCrew) {
  const sheet = getDB().getSheetByName('Crew');
  // OPTIMASI: TextFinder sangat cepat untuk 1 baris
  const tf = sheet.getRange("A:A").createTextFinder(kodeCrew.toString()).matchEntireCell(true).findNext();
  if(tf) {
    return sheet.getRange(tf.getRow(), 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  }
  return null; 
}

function getRelatedData(kodeCrew, sheetName) {
  const sheet = getDB().getSheetByName(sheetName);
  if(!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  if(data.length <= 1) return [];
  const result = [];
  for(let i = 1; i < data.length; i++) {
    if(data[i][1].toString() === kodeCrew.toString()) result.push(data[i]);
  }
  return result;
}

function addFamilyData(formObject) {
  const sheet = getDB().getSheetByName('Crew_Fam');
  if (!sheet) throw new Error("Sheet Crew_Fam tidak ditemukan.");
  const uniqueId = "FAM-" + new Date().getTime();
  sheet.appendRow([ uniqueId, formObject.fam_kode_crew, formObject.fam_nama, formObject.fam_hubungan, formObject.fam_tgl_lahir || '', formObject.fam_telepon || '', formObject.fam_asal || '' ]);
  return "Data keluarga atas nama " + formObject.fam_nama + " berhasil ditambahkan!";
}

function getShipTypes() {
  const cache = CacheService.getScriptCache();
  if (cache.get('shipTypes')) return JSON.parse(cache.get('shipTypes'));

  const sheet = getDB().getSheetByName('Kapal_Reference');
  if(!sheet) return ["Lainnya"];
  const lastRow = sheet.getLastRow();
  if(lastRow < 2) return ["Lainnya"];
  const list = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(String);
  
  cache.put('shipTypes', JSON.stringify(list), 21600);
  return list;
}

function addSeaServiceData(formObject) {
  const sheet = getDB().getSheetByName('Crew_SeaServices');
  if (!sheet) throw new Error("Sheet Crew_SeaServices tidak ditemukan.");
  const durasiOtomatis = hitungSelisihTanggal(formObject.sea_on, formObject.sea_off);
  const uniqueId = "SEA-" + new Date().getTime();
  sheet.appendRow([ uniqueId, formObject.sea_kode_crew, formObject.sea_kapal, formObject.sea_bendera || '', formObject.sea_grt || '', formObject.sea_rank || '', formObject.sea_perusahaan || '', formObject.sea_on || '', formObject.sea_off || '', durasiOtomatis, formObject.sea_jenis_kapal || '' ]);
  return "Data pengalaman berhasil ditambahkan!";
}

function hitungSelisihTanggal(start, end) {
  if (!start || !end) return "-";
  var d1 = new Date(start);
  var d2 = new Date(end);
  if (d2 < d1) return "Error: Tanggal tdk valid";
  var months = (d2.getFullYear() - d1.getFullYear()) * 12;
  months -= d1.getMonth();
  months += d2.getMonth();
  var days = d2.getDate() - d1.getDate();
  if (days < 0) {
    months--;
    var lastDayPrevMonth = new Date(d2.getFullYear(), d2.getMonth(), 0).getDate();
    days += lastDayPrevMonth;
  }
  var hasil = "";
  if (months > 0) hasil += months + " Bulan ";
  if (days > 0) hasil += days + " Hari";
  return hasil === "" ? "0 Hari" : hasil;
}

function addDocumentData(dataObj) {
  const sheet = getDB().getSheetByName('Crew_Doc');
  if (!sheet) throw new Error("Sheet Crew_Doc tidak ditemukan.");
  let fileUrl = "";
  if (dataObj.fileData && dataObj.fileName) {
    try {
      const mainFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      let targetFolder = mainFolder; 
      const searchString = "title contains '" + dataObj.doc_kode_crew + "' and trashed = false";
      const folderIterator = mainFolder.searchFolders(searchString);
      if (folderIterator.hasNext()) { targetFolder = folderIterator.next(); } 
      else { targetFolder = mainFolder.createFolder(dataObj.doc_kode_crew + " - Dokumen Susulan"); }
      
      const contentType = dataObj.fileData.substring(5, dataObj.fileData.indexOf(';'));
      const bytes = Utilities.base64Decode(dataObj.fileData.split(',')[1]);
      const blob = Utilities.newBlob(bytes, contentType, dataObj.fileName);
      const file = targetFolder.createFile(blob); 
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
    } catch(e) { throw new Error("Gagal mengupload file ke Drive. Error: " + e.message); }
  }
  const uniqueId = "DOC-" + new Date().getTime();
  sheet.appendRow([ uniqueId, dataObj.doc_kode_crew, dataObj.doc_tipe, dataObj.doc_nomor, dataObj.doc_tempat || '', dataObj.doc_issued || '', dataObj.doc_expiry || '', fileUrl ]);
  catatLogPerubahan(dataObj.doc_kode_crew, "-", "Dokumen", "Menambahkan dokumen baru tipe: " + dataObj.doc_tipe);
  return "Dokumen " + dataObj.doc_tipe + " berhasil disimpan!";
}

function updateDocumentData(dataObj) {
  const sheet = getDB().getSheetByName('Crew_Doc');
  if (!sheet) throw new Error("Sheet Crew_Doc tidak ditemukan.");

  const tf = sheet.getRange("A:A").createTextFinder(dataObj.doc_id.toString()).matchEntireCell(true).findNext();
  if (!tf) throw new Error("Data dokumen tidak ditemukan.");
  let rowIndex = tf.getRow();
  
  let fileUrl = sheet.getRange(rowIndex, 8).getValue();
  if (dataObj.fileData && dataObj.fileName) {
    if (fileUrl && fileUrl.includes("drive.google.com")) {
      try {
        const oldFileIdMatch = fileUrl.match(/[-\w]{25,}/);
        if (oldFileIdMatch && oldFileIdMatch[0]) { DriveApp.getFileById(oldFileIdMatch[0]).setTrashed(true); }
      } catch(e) {}
    }
    try {
      const mainFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      let targetFolder = mainFolder;
      const searchString = "title contains '" + dataObj.doc_kode_crew + "' and trashed = false";
      const folderIterator = mainFolder.searchFolders(searchString);
      if (folderIterator.hasNext()) targetFolder = folderIterator.next();
      
      const contentType = dataObj.fileData.substring(5, dataObj.fileData.indexOf(';'));
      const bytes = Utilities.base64Decode(dataObj.fileData.split(',')[1]);
      const blob = Utilities.newBlob(bytes, contentType, dataObj.fileName);
      const file = targetFolder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
    } catch(e) { throw new Error("Gagal upload file baru: " + e.message); }
  }

  sheet.getRange(rowIndex, 3, 1, 6).setValues([[ dataObj.doc_tipe, dataObj.doc_nomor, dataObj.doc_tempat, dataObj.doc_issued, dataObj.doc_expiry, fileUrl ]]);
  return "Dokumen " + dataObj.doc_tipe + " berhasil diperbarui!";
}

function updateFamilyData(obj) {
  const sheet = getDB().getSheetByName('Crew_Fam');
  const tf = sheet.getRange("A:A").createTextFinder(obj.fam_id.toString()).matchEntireCell(true).findNext();
  if(tf) {
      const row = tf.getRow();
      sheet.getRange(row, 3, 1, 5).setValues([[ obj.fam_nama, obj.fam_hubungan, obj.fam_tgl_lahir, obj.fam_telepon, obj.fam_asal ]]);
      return "Data keluarga " + obj.fam_nama + " berhasil diupdate!";
  }
}

function updateSeaServiceData(obj) {
  const sheet = getDB().getSheetByName('Crew_SeaServices');
  const durasiBaru = hitungSelisihTanggal(obj.sea_on, obj.sea_off);
  const tf = sheet.getRange("A:A").createTextFinder(obj.sea_id.toString()).matchEntireCell(true).findNext();
  if(tf) {
      const row = tf.getRow();
      sheet.getRange(row, 3, 1, 9).setValues([[ obj.sea_kapal, obj.sea_bendera, obj.sea_grt, obj.sea_rank, obj.sea_perusahaan, obj.sea_on, obj.sea_off, durasiBaru, obj.sea_jenis_kapal ]]);
      return "Riwayat kapal " + obj.sea_kapal + " berhasil diupdate!";
  }
  throw new Error("Data tidak ditemukan.");
}

function deleteFamilyData(id) {
  const sheet = getDB().getSheetByName('Crew_Fam');
  const tf = sheet.getRange("A:A").createTextFinder(id.toString()).matchEntireCell(true).findNext();
  if(tf) {
      sheet.deleteRow(tf.getRow());
      return "Data keluarga berhasil dihapus!";
  }
  throw new Error("Data tidak ditemukan.");
}

function deleteSeaServiceData(id) {
  const sheet = getDB().getSheetByName('Crew_SeaServices');
  const tf = sheet.getRange("A:A").createTextFinder(id.toString()).matchEntireCell(true).findNext();
  if(tf) {
      sheet.deleteRow(tf.getRow());
      return "Riwayat pengalaman berhasil dihapus!";
  }
  throw new Error("Data tidak ditemukan.");
}

function deleteDocumentData(id) {
  const sheet = getDB().getSheetByName('Crew_Doc');
  const tf = sheet.getRange("A:A").createTextFinder(id.toString()).matchEntireCell(true).findNext();
  if(tf) {
      const row = tf.getRow();
      const fileUrl = sheet.getRange(row, 8).getValue();
      if (fileUrl && fileUrl.includes("drive.google.com")) {
        try {
          const fileIdMatch = fileUrl.match(/[-\w]{25,}/);
          if (fileIdMatch && fileIdMatch[0]) DriveApp.getFileById(fileIdMatch[0]).setTrashed(true); 
        } catch(e) {}
      }
      sheet.deleteRow(row);
      return "Dokumen berhasil dihapus dari sistem dan Google Drive!";
  }
  throw new Error("Data dokumen tidak ditemukan.");
}

// 1. Mengambil Nomor Surat Baru
function getNomorSuratBaru() {
  const props = PropertiesService.getScriptProperties();
  let urut = props.getProperty('NOMOR_SURAT_TTD');
  if (!urut) { urut = 1; } else { urut = parseInt(urut) + 1; }
  const formatUrut = urut.toString().padStart(3, '0');
  const date = new Date();
  const romawi = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"];
  const bulanRomawi = romawi[date.getMonth()];
  const tahun = date.getFullYear();
  const nomorSuratLengkap = `${formatUrut}/TTD/ADM/PPM/${bulanRomawi}/${tahun}`;
  return { angkaUrut: urut, nomorLengkap: nomorSuratLengkap };
}

// 2. Fungsi Utama Generate PDF Tanda Terima
function generatePDFTandaTerima(kodeCrew, nomorSuratEdit, angkaSimpan, idKopSurat) {
  PropertiesService.getScriptProperties().setProperty('NOMOR_SURAT_TTD', angkaSimpan.toString());
  const profilKru = getCrewProfile(kodeCrew);
  if (!profilKru) throw new Error("Data kru tidak ditemukan.");
  const namaKru = profilKru[1] || "-";
  const jabatanKru = profilKru[2] || "-";
  const teleponKru = profilKru[5] || "-";
  const dokumenKru = getRelatedData(kodeCrew, 'Crew_Doc');
  let imgKopHTML = "";
  if (idKopSurat) {
    try {
      const fileKop = DriveApp.getFileById(idKopSurat);
      const base64 = Utilities.base64Encode(fileKop.getBlob().getBytes());
      const mime = fileKop.getMimeType();
      imgKopHTML = `<img src="data:${mime};base64,${base64}" style="width: 100%; max-height: 150px; object-fit: contain; margin-bottom: 20px;">`;
    } catch(e) { console.error("Gagal meload Kop Surat: " + e.message); }
  }

  let barisTabel = "";
  if (dokumenKru.length === 0) {
    barisTabel = `<tr><td colspan="6" style="text-align: center; padding: 10px;">Belum ada dokumen yang diserahkan.</td></tr>`;
  } else {
    dokumenKru.forEach((doc, index) => {
      barisTabel += `
        <tr>
          <td style="text-align: center; padding: 5px;">${index + 1}</td>
          <td style="padding: 5px;">${doc[2] || '-'}</td>
          <td style="padding: 5px; text-align: center;">${doc[3] || '-'}</td>
          <td style="padding: 5px; text-align: center;">${doc[4] || '-'}</td>
          <td style="padding: 5px; text-align: center;">${formatTanggalIndo(doc[5])}</td>
          <td style="padding: 5px; text-align: center;">${formatTanggalIndo(doc[6])}</td>
        </tr>
      `;
    });
  }

  const tanggalHariIni = formatTanggalIndo(new Date().toISOString().split('T')[0]);
  const htmlTemplate = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        @page { size: A4; margin: 0.5in; }
        body { font-family: 'Calibri', sans-serif; color: #000; font-size: 14px; }
        .center { text-align: center; }
        .bold { font-weight: bold; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        table, th, td { border: 1px solid black; }
        th { background-color: #f2f2f2; padding: 8px; text-align: center; }
        .header-title { font-size: 18px; margin-top: 10px; margin-bottom: 5px; text-decoration: underline; }
        .info-table { width: 60%; margin: 20px auto; border: none; }
        .info-table td { border: none; padding: 3px; text-align: left; }
        .signature-table { width: 100%; border: none; margin-top: 50px; text-align: center; }
        .signature-table td { border: none; width: 50%; }
      </style>
    </head>
    <body>
      <div class="center">
        ${imgKopHTML}
        
        <div class="bold header-title">TANDA TERIMA DOKUMEN</div>
        <div>Nomor: ${nomorSuratEdit}</div>
        
        <table class="info-table">
          <tr><td style="width: 150px;">Nama Kru</td><td>: <b>${namaKru}</b></td></tr>
          <tr><td>Jabatan</td><td>: ${jabatanKru}</td></tr>
          <tr><td>Nomor Telepon</td><td>: ${teleponKru}</td></tr>
        </table>
        
        <p>Telah diterima dokumen asli / salinan dari nama tersebut di atas dengan rincian sebagai berikut:</p>
        
        <table>
          <thead>
            <tr>
              <th style="width: 5%;">No</th>
              <th style="width: 25%;">Nama Dokumen</th>
              <th style="width: 20%;">Nomor Dokumen</th>
              <th style="width: 15%;">Tempat Terbit</th>
              <th style="width: 15%;">Issued</th>
              <th style="width: 20%;">Expiry</th>
            </tr>
          </thead>
          <tbody>
            ${barisTabel}
          </tbody>
        </table>
        
        <table class="signature-table">
          <tr>
            <td>Yang Menyerahkan,<br><br><br><br><br><b>(${namaKru})</b></td>
            <td>Penerima,<br><br><br><br><br><b>( Admin Dokumen )</b></td>
          </tr>
        </table>
      </div>
    </body>
    </html>
  `;
  const blob = Utilities.newBlob(htmlTemplate, MimeType.HTML).getAs(MimeType.PDF);
  blob.setName(`Tanda Terima Dokumen - ${namaKru}.pdf`);
  const mainFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const pdfFile = mainFolder.createFile(blob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return pdfFile.getDownloadUrl();
}

function formatTanggalIndo(str) {
  if (!str || str === '-') return '-';
  if (str.includes('-')) {
    const p = str.split('-');
    if (p[0].length === 4) return `${p[2]}/${p[1]}/${p[0]}`;
  }
  return str;
}

// ==========================================
// MESIN CETAK CURRICULUM VITAE (CV)
// ==========================================
function generateCVPDF(kodeCrew, idKopSurat) {
  const profil = getCrewProfile(kodeCrew);
  if (!profil) throw new Error("Data kru tidak ditemukan.");

  const jabatan = (profil[2] || "").toUpperCase();
  const isDeck = jabatan.includes("ANT") || /MASTER|OFFICER|DECK|BOSUN|AB|OS|PUMPMAN/.test(jabatan);
  const isEngine = jabatan.includes("ATT") || /ENGINEER|ENGINE|OILER|WIPER|FITTER/.test(jabatan);

  let kelasRomawi = "";
  let matchRomawi = jabatan.match(/\b(I{1,3}|IV|V)\b/);
  if (matchRomawi) { kelasRomawi = " Class " + matchRomawi[0]; }

  const docMap = {
    "PP": "Passport", "CDC": "Seaman's Book", "YELLOW": "Yellow Fever Certificate",
    "BST": "Basic Safety Training", "AFF": "Advanced Fire Fighting", "SCRB": "Survival Craft and Rescue Boats",
    "MFA": "Medical First Aid", "SAT": "Security Awareness Training", "SDSD": "Seafarers with Designated Security Duties",
    "SSO": "Ship Security Officer", "MCOB": "Medical Care On Board", "ERRM": "Engine Room Resource Management",
    "BRM": "Bridge Resource Management", "ARPA": "ARPA Simulator", "RADAR": "Radar Simulator",
    "ECDIS": "Electronic Chart Display and Information System", "BOCT": "Basic Oil and Chemical Tanker",
    "BLGT": "Basic Liquefied Gas Tanker", "AOT": "Advanced Oil Tanker", "ACT": "Advanced Chemical Tanker",
    "ALGT": "Advanced Liquefied Gas Tanker", "ABLE": "Able Seafarer", "RATING": "Rating Forming"
  };

  const rawDocs = getRelatedData(kodeCrew, 'Crew_Doc');
  const seas = getRelatedData(kodeCrew, 'Crew_SeaServices');

  let travelDocs = []; let cocDocs = []; let copDocs = [];

  rawDocs.forEach(d => {
    let tipeAsli = (d[2] || "").toString().toUpperCase().trim();
    if (tipeAsli === "ONLINE") return; 

    let namaInggris = docMap[tipeAsli] || d[2]; 

    if (tipeAsli === "COC") {
      if (isDeck) namaInggris = "Certificate of Competency (Deck Officer" + kelasRomawi + ")";
      else if (isEngine) namaInggris = "Certificate of Competency (Engineer Officer" + kelasRomawi + ")";
      else namaInggris = "Certificate of Competency";
    } else if (tipeAsli === "COE") {
      if (isDeck) namaInggris = "Certificate of Endorsement (Deck Officer" + kelasRomawi + ")";
      else if (isEngine) namaInggris = "Certificate of Endorsement (Engineer Officer" + kelasRomawi + ")";
      else namaInggris = "Certificate of Endorsement";
    } else if (tipeAsli === "GMDSS") { namaInggris = "GMDSS Radio Operator";
    } else if (tipeAsli === "GMDSE") { namaInggris = "GMDSS Endorsement"; }

    let docObj = { nama: namaInggris, no: d[3], tempat: d[4], issued: d[5], expiry: d[6] };
    if (["PP", "CDC", "YELLOW"].includes(tipeAsli)) { travelDocs.push(docObj); } 
    else if (["COC", "COE", "GMDSS", "GMDSE"].includes(tipeAsli)) { cocDocs.push(docObj); } 
    else { copDocs.push(docObj); }
  });

  let imgKopHTML = "";
  if (idKopSurat) {
    try {
      const fileKop = DriveApp.getFileById(idKopSurat);
      const base64 = Utilities.base64Encode(fileKop.getBlob().getBytes());
      const mime = fileKop.getMimeType();
      imgKopHTML = `<img src="data:${mime};base64,${base64}" style="width: 100%; max-height: 130px; object-fit: contain; margin-bottom: 10px;">`;
    } catch(e) { console.error("Kop gagal dimuat."); }
  }

  let imgFotoHTML = "";
  let linkFotoDrive = profil[10] || "";
  if (linkFotoDrive && linkFotoDrive.includes("drive.google.com")) {
    try {
      const fotoIdMatch = linkFotoDrive.match(/[-\w]{25,}/);
      if (fotoIdMatch && fotoIdMatch[0]) {
        const fileFoto = DriveApp.getFileById(fotoIdMatch[0]);
        const base64Foto = Utilities.base64Encode(fileFoto.getBlob().getBytes());
        const mimeFoto = fileFoto.getMimeType();
        imgFotoHTML = `<img src="data:${mimeFoto};base64,${base64Foto}" style="width: 3cm; height: 4cm; object-fit: cover; border: 1px solid #333;">`;
      }
    } catch(e) { console.error("Foto kru gagal dimuat."); }
  } else {
    imgFotoHTML = `<div style="width: 3cm; height: 4cm; border: 1px solid #333; display: inline-block; line-height: 4cm; text-align: center; background-color: #f2f2f2; color: #888;">No Photo</div>`;
  }

  const formatDurasiInggris = (durasiIndo) => {
    if (!durasiIndo || durasiIndo === "-") return "-";
    return durasiIndo.replace(/Bulan/g, "Months").replace(/Hari/g, "Days");
  };
  
  let usiaInggris = profil[13] ? profil[13].replace(/Tahun/g, "Years Old") : "-";

  const renderTabel = (arr) => {
    return arr.map(x => `
      <tr>
        <td style="padding:4px;">${x.nama}</td>
        <td class="center" style="padding:4px;">${x.no || '-'}</td>
        <td class="center" style="padding:4px;">${x.tempat || '-'}</td>
        <td class="center" style="padding:4px;">${formatTanggalIndo(x.issued)}</td>
        <td class="center" style="padding:4px;">${formatTanggalIndo(x.expiry)}</td>
      </tr>
    `).join('');
  };

  const renderTabelLayar = (arr) => {
    return arr.map(x => `
      <tr>
        <td style="padding:4px;">${x[2] || '-'}</td>
        <td class="center" style="padding:4px;">${x[10] || '-'}</td>
        <td class="center" style="padding:4px;">${x[5] || '-'}</td>
        <td class="center" style="padding:4px;">${x[6] || '-'}</td>
        <td class="center" style="padding:4px;">${formatTanggalIndo(x[7])}</td>
        <td class="center" style="padding:4px;">${formatTanggalIndo(x[8])}</td>
        <td class="center" style="padding:4px;">${formatDurasiInggris(x[9])}</td>
      </tr>
    `).join('');
  };

  let sectionCounter = 1;
  let htmlTravel = ""; let htmlCOC = ""; let htmlCOP = ""; let htmlSea = "";

  if (travelDocs.length > 0) {
    htmlTravel = `
      <div class="section-title">${sectionCounter}. TRAVEL DOCUMENTS</div>
      <table class="data-table">
        <tr><th style="width: 30%;">Document Name</th><th style="width: 20%;">Number</th><th style="width: 20%;">Place of Issue</th><th style="width: 15%;">Date of Issue</th><th style="width: 15%;">Expiry Date</th></tr>
        ${renderTabel(travelDocs)}
      </table>
    `;
    sectionCounter++;
  }

  if (cocDocs.length > 0) {
    htmlCOC = `
      <div class="section-title">${sectionCounter}. CERTIFICATE OF COMPETENCY</div>
      <table class="data-table">
        <tr><th style="width: 30%;">Certificate Name</th><th style="width: 20%;">Number</th><th style="width: 20%;">Place of Issue</th><th style="width: 15%;">Date of Issue</th><th style="width: 15%;">Expiry Date</th></tr>
        ${renderTabel(cocDocs)}
      </table>
    `;
    sectionCounter++;
  }

  if (copDocs.length > 0) {
    htmlCOP = `
      <div class="section-title">${sectionCounter}. CERTIFICATE OF PROFICIENCY</div>
      <table class="data-table">
        <tr><th style="width: 30%;">Certificate Name</th><th style="width: 20%;">Number</th><th style="width: 20%;">Place of Issue</th><th style="width: 15%;">Date of Issue</th><th style="width: 15%;">Expiry Date</th></tr>
        ${renderTabel(copDocs)}
      </table>
    `;
    sectionCounter++;
  }

  if (seas.length > 0) {
    htmlSea = `
      <div class="section-title">${sectionCounter}. SEA SERVICE RECORD</div>
      <table class="data-table">
        <tr><th style="width: 20%;">Vessel Name</th><th style="width: 15%;">Type</th><th style="width: 10%;">Rank</th><th style="width: 15%;">Company</th><th style="width: 12%;">Sign On</th><th style="width: 12%;">Sign Off</th><th style="width: 16%;">Duration</th></tr>
        ${renderTabelLayar(seas)}
      </table>
    `;
    sectionCounter++;
  }

  const htmlTemplate = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        @page { size: A4; margin: 0.4in; }
        body { font-family: 'Calibri', sans-serif; color: #000; font-size: 11px; }
        .center { text-align: center; }
        .bold { font-weight: bold; }
        .title { font-size: 16px; text-decoration: underline; margin-bottom: 15px; }
        .section-title { background-color: #d9d9d9; padding: 4px; font-weight: bold; margin-top: 15px; border: 1px solid black; }
        table.data-table { width: 100%; border-collapse: collapse; margin-top: 0px; }
        table.data-table, table.data-table th, table.data-table td { border: 1px solid black; }
        table.data-table th { background-color: #f2f2f2; padding: 5px; text-align: center; font-size: 10px; }
        
        .layout-container { width: 100%; display: table; margin-bottom: 10px; }
        .layout-text { display: table-cell; width: 75%; vertical-align: top; }
        .layout-photo { display: table-cell; width: 25%; vertical-align: top; text-align: right; }
        
        .profile-table { width: 100%; border: none; }
        .profile-table td { border: none; padding: 3px; font-size: 12px; text-align: left; }
      </style>
    </head>
    <body>
      ${imgKopHTML}
      <div class="center bold title">CURRICULUM VITAE</div>
      
      <div class="layout-container">
        <div class="layout-text">
          <table class="profile-table">
            <tr><td style="width: 35%;"><b>Name</b></td><td>: ${profil[1] || '-'}</td></tr>
            <tr><td><b>Rank</b></td><td>: <b>${profil[2] || '-'}</b></td></tr>
            <tr><td><b>Place, Date of Birth</b></td><td>: ${profil[3] || '-'}, ${formatTanggalIndo(profil[4])}</td></tr>
            <tr><td><b>Age</b></td><td>: ${usiaInggris}</td></tr>
            <tr><td valign="top"><b>Address</b></td><td>: ${profil[7] || '-'}</td></tr>
          </table>
        </div>
        <div class="layout-photo">
          ${imgFotoHTML}
        </div>
      </div>

      ${htmlTravel}
      ${htmlCOC}
      ${htmlCOP}
      ${htmlSea}

    </body>
    </html>
  `;
  const blob = Utilities.newBlob(htmlTemplate, MimeType.HTML).getAs(MimeType.PDF);
  blob.setName(`CV_${profil[1]}_${profil[2]}.pdf`);
  
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const pdfFile = folder.createFile(blob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return pdfFile.getDownloadUrl();
}

// ==========================================
// FITUR DOWNLOAD SEMUA DOKUMEN (ZIP)
// ==========================================
function generateZipDokumenKru(kodeCrew) {
  const profil = getCrewProfile(kodeCrew);
  if (!profil) throw new Error("Data kru tidak ditemukan.");
  
  const namaKru = profil[1] || "Kru";
  const docs = getRelatedData(kodeCrew, 'Crew_Doc');
  
  if (!docs || docs.length === 0) throw new Error("Kru ini belum memiliki dokumen yang diunggah.");

  let blobs = [];
  
  // Masukkan juga Foto Profil, KTP, dan TTD jika ada
  const personalFiles = [
    { url: profil[9], nama: "KTP" },
    { url: profil[10], nama: "FOTO" },
    { url: profil[11], nama: "TTD" }
  ];

  personalFiles.forEach(pf => {
    if (pf.url && pf.url.includes("drive.google.com")) {
      try {
        let idMatch = pf.url.match(/[-\w]{25,}/);
        if (idMatch && idMatch[0]) {
          let file = DriveApp.getFileById(idMatch[0]);
          let blob = file.getBlob();
          let ext = file.getName().split('.').pop();
          blob.setName(`00_PROFIL_${pf.nama}.${ext}`);
          blobs.push(blob);
        }
      } catch(e) {} // Abaikan jika file tidak ditemukan
    }
  });

  // Ambil semua dokumen sertifikat
  docs.forEach(d => {
    let url = d[7]; // Kolom URL file
    if (url && url.includes("drive.google.com")) {
      try {
        let idMatch = url.match(/[-\w]{25,}/);
        if (idMatch && idMatch[0]) {
          let file = DriveApp.getFileById(idMatch[0]);
          let blob = file.getBlob();
          
          let ext = file.getName().split('.').pop();
          // Bersihkan nama tipe dokumen agar aman untuk nama file
          let docType = (d[2] || "Dokumen").toString().replace(/[^a-zA-Z0-9 ]/g, "").trim().toUpperCase();
          
          blob.setName(`${docType}_${file.getName()}`);
          blobs.push(blob);
        }
      } catch(e) {
        console.error("Gagal mengambil file: " + url);
      }
    }
  });

  if (blobs.length === 0) throw new Error("Tidak ada file valid yang bisa di-compress ke dalam ZIP.");

  // Proses Zipping
  const namaZip = `Dokumen_${kodeCrew}_${namaKru.replace(/[^a-zA-Z0-9]/g, "_")}.zip`;
  const zipBlob = Utilities.zip(blobs, namaZip);
  
  // Simpan sementara di Drive dan berikan linknya
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const zipFile = folder.createFile(zipBlob);
  zipFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return zipFile.getDownloadUrl();
}

// ==========================================
// OPTIMASI DETAIL CREW: MASTER FETCH
// Ditaruh di paling bawah Crud_Detail.gs
// ==========================================
function getCrewFullDetail(kodeCrew) {
  return {
    profil: getCrewProfile(kodeCrew),
    famData: getRelatedData(kodeCrew, 'Crew_Fam'),
    docData: getRelatedData(kodeCrew, 'Crew_Doc'),
    seaData: getRelatedData(kodeCrew, 'Crew_SeaServices')
  };
}