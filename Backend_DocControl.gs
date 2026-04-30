// ==========================================
// BACKEND: SIRKULASI DOKUMEN (DOC CONTROL)
// ==========================================

// 1. Ambil daftar Kru (HANYA STANDBY / PLOTTING) untuk Pencarian
function getListKruSirkulasi() {
  const sheet = getDB().getSheetByName('Crew'); 
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const list = [];
  
  // =======================================================
  // KUNCI ABSOLUT: GANTI ANGKA DI BAWAH SESUAI KOLOM STATUS ANDA!
  // (A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7, I=8, J=9)
  // =======================================================
  const indexStatus = 12; // <--- UBAH ANGKA INI!
  
  // Kita mulai dari baris ke-2 (index 1) untuk melewati Header
  for (let i = 1; i < data.length; i++) {
    const kode = data[i][0];
    const nama = data[i][1];
    
    if (kode !== "") { 
       // Tarik data status dan paksa jadi HURUF BESAR agar cocok dengan format Anda
       const statusKru = data[i][indexStatus] ? data[i][indexStatus].toString().toUpperCase() : "";
       
       // HANYA masukkan ke daftar jika teks status mengandung kata STANDBY atau PLOTTING
       if (statusKru.includes('STANDBY') || statusKru.includes('PLOTTING')) {
           list.push({ kode: kode, nama: nama });
       }
    }
  }
  return list;
}

// 2. Ambil seluruh dokumen milik kru tersebut (Anti-Duplikat & Filter Sedang Dipinjam)
function getDokumenKruTersedia(kode) {
  try {
      const db = getDB();
      const docs = [];
      let namaKru = "Kru Tidak Ditemukan";
      let statusKru = "";

      // ==========================================
      // OPTIMASI: A. Cari Data Kru via TextFinder (Bukan Looping)
      // ==========================================
      const sheetMain = db.getSheetByName('Crew');
      if (!sheetMain) throw new Error("Sheet Crew tidak ditemukan.");
      
      const tfKru = sheetMain.getRange("A:A").createTextFinder(kode.toString()).matchEntireCell(true).findNext();
      if (!tfKru) return { nama: "Error", dokumen: [], error: "Kru tidak ditemukan di sistem." };
      
      const rowKru = tfKru.getRow();
      // Asumsi: Nama di kolom 2, Status di kolom 13
      namaKru = sheetMain.getRange(rowKru, 2).getValue();
      statusKru = String(sheetMain.getRange(rowKru, 13).getValue()).toUpperCase().trim();

      // SABUK PENGAMAN 1: BLOKIR STATUS ILEGAL
      if (statusKru === "ONBOARD") return { nama: namaKru, dokumen: [], error: "Ditolak! Kru ini sedang ONBOARD di kapal." };
      if (statusKru === "TANPA BERKAS") return { nama: namaKru, dokumen: [], error: "Ditolak! Kru ini sudah berstatus TANPA BERKAS." };
      if (!statusKru.includes("STANDBY") && !statusKru.includes("PLOTTING")) return { nama: namaKru, dokumen: [], error: `Ditolak! Status kru saat ini: ${statusKru}` };

      // ==========================================
      // B. RADAR PINTAR: CEK DOKUMEN YANG SEDANG DIPINJAM
      // ==========================================
      const docSedangDipinjam = new Set();
      const sheetLog = db.getSheetByName('Log_Dokumen');
      if (sheetLog) {
          const dataLog = sheetLog.getDataRange().getDisplayValues(); 
          for (let i = 1; i < dataLog.length; i++) {
              if (String(dataLog[i][2]).trim() === String(kode).trim() && String(dataLog[i][7]).includes("Sedang Dipinjam")) {
                  const listPinjam = String(dataLog[i][5]).split(',');
                  listPinjam.forEach(d => docSedangDipinjam.add(d.trim()));
              }
          }
      }

      // ==========================================
      // OPTIMASI: C. Tarik murni dari sheet 'Crew_Doc' menggunakan TextFinder FindAll
      // ==========================================
      const sheetDoc = db.getSheetByName('Crew_Doc');
      if (sheetDoc) {
          // Cari semua baris yang punya kode kru ini di kolom B
          const tfDocs = sheetDoc.getRange("B:B").createTextFinder(kode.toString()).matchEntireCell(true).findAll();
          const docTercatat = new Set(); 

          tfDocs.forEach((cell, index) => {
              const r = cell.getRow();
              // Asumsi Crew_Doc: Nama Dok di kolom 3, Nomor di kolom 4
              const namaDokumen = String(sheetDoc.getRange(r, 3).getValue()).trim();
              const nomorDokumen = String(sheetDoc.getRange(r, 4).getValue()) || "-";
              
              if (!docTercatat.has(namaDokumen) && namaDokumen !== "" && !docSedangDipinjam.has(namaDokumen)) {
                  docTercatat.add(namaDokumen);
                  docs.push({ id: 'doc-' + index, namaDoc: namaDokumen, nomor: nomorDokumen });
              }
          });
      }

      return { nama: namaKru, dokumen: docs };
  } catch (e) {
      console.error("Error getDokumenKruTersedia: " + e.message);
      return { nama: "Error", dokumen: [], error: "Sistem gagal memproses data: " + e.message }; 
  }
}

// 3. Simpan Transaksi Peminjaman / Penarikan
function prosesSirkulasiKeluar(data) {
  const db = getDB();
  const sheetLog = db.getSheetByName('Log_Dokumen');
  const sheetCrew = db.getSheetByName('Crew');
  
  const idTrx = "TRXD-" + new Date().getTime().toString().slice(-6);
  const waktu = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
  const admin = dapatkanNamaAdminAktif(); 
  
  // ==========================================
  // SABUK PENGAMAN 2: CEK STATUS TERAKHIR DI EXCEL SEBELUM SAVE
  // ==========================================
  const dataCrew = sheetCrew.getDataRange().getValues();
  const indexStatus = 12; // Kolom M
  let statusKruAktif = "";
  let barisKru = -1;
  
  for (let i = 1; i < dataCrew.length; i++) {
      if (String(dataCrew[i][0]).trim() === String(data.kode).trim()) {
          statusKruAktif = String(dataCrew[i][indexStatus]).toUpperCase().trim();
          barisKru = i + 1;
          break;
      }
  }

  // JIKA KRU DITEMUKAN, LAKUKAN VALIDASI KERAS
  if (barisKru !== -1) {
      if (statusKruAktif === "ONBOARD") {
          throw new Error("Transaksi ditolak! Kru saat ini berstatus ONBOARD.");
      }
      if (statusKruAktif === "TANPA BERKAS") {
          throw new Error("Kru ini sudah berstatus TANPA BERKAS di brankas.");
      }
      
      // JIKA LOLOS VALIDASI DAN AKSI ADALAH PENARIKAN TOTAL -> UBAH STATUS
      if (data.aksi === "Penarikan Total") {
          sheetCrew.getRange(barisKru, indexStatus + 1).setValue("TANPA BERKAS");
      }
  }

  // 2. Tentukan status di Log_Dokumen
  let statusLog = (data.aksi === "Pinjam Sementara") ? "Sedang Dipinjam" : "Selesai (Ditarik Total)";

  // 3. Simpan ke Log_Dokumen
  sheetLog.appendRow([
     idTrx, waktu, data.kode, data.nama, data.aksi, data.dokumen, data.alasan, statusLog, admin
  ]);
  
  catatLogPerubahan(data.kode, data.nama, "Sirkulasi", `Melakukan ${data.aksi}: ${data.dokumen}`);
  
  return "Berhasil! Transaksi " + data.aksi + " tercatat dan status database diperbarui.";
}

// ==========================================
// 3. Ambil Riwayat Sirkulasi untuk Tabel
// ==========================================
function getRiwayatSirkulasi() {
  try {
      const db = getDB(); 
      const sheet = db.getSheetByName('Log_Dokumen');
      
      if (!sheet) return [];
      
      // MAGIC: Gunakan getDisplayValues() agar SEMUA data (termasuk Tanggal) ditarik sebagai TEKS murni!
      const data = sheet.getDataRange().getDisplayValues(); 
      const list = [];
      
      // Looping terbalik (dari bawah ke atas)
      for (let i = data.length - 1; i > 0; i--) {
        if (data[i][0] !== "") {
          list.push({
            idTrx: data[i][0],
            waktu: data[i][1],
            kode: data[i][2],
            nama: data[i][3],
            aksi: data[i][4],
            dokumen: data[i][5],
            alasan: data[i][6],
            status: data[i][7],
            admin: data[i][8] || "Sistem"
          });
        }
      }
      return list;
  } catch(e) {
      console.error(e);
      return []; // Kembalikan array kosong jika terjadi error parah
  }
}

// ==========================================
// 5. Ambil Daftar Kru yang SEDANG MEMINJAM Dokumen
// ==========================================
function getListKruPinjam() {
  const sheet = getDB().getSheetByName('Log_Dokumen');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getDisplayValues(); // Gunakan DisplayValues agar aman
  const list = [];
  const filterDuplikat = new Set();
  
  for (let i = 1; i < data.length; i++) {
    // Paksa huruf kecil semua agar tidak peduli "Sedang Dipinjam" atau "SEDANG DIPINJAM"
    const statusLog = String(data[i][7]).toLowerCase();
    
    if (statusLog.includes('sedang dipinjam') && data[i][0] !== "") {
      const kode = String(data[i][2]).trim(); // Hilangkan spasi gaib
      const nama = String(data[i][3]).trim();
      
      if (!filterDuplikat.has(kode)) {
          filterDuplikat.add(kode);
          list.push({ kode: kode, nama: nama });
      }
    }
  }
  return list;
}

// ==========================================
// 6. Ambil Daftar Dokumen + Detailnya untuk Smart Return
// ==========================================
function getDokumenDipinjam(kode) {
  try {
      const db = getDB();
      const sheetLog = db.getSheetByName('Log_Dokumen');
      const sheetDoc = db.getSheetByName('Crew_Doc');
      
      const dataLog = sheetLog.getDataRange().getDisplayValues();
      const dataDoc = sheetDoc ? sheetDoc.getDataRange().getDisplayValues() : [];
      
      const docsDetail = [];
      const idTrxLama = [];
      const targetKode = String(kode).trim(); // Kunci baja target kode dari HTML
      
      // A. Cari apa saja yang sedang dipinjam di Log
      for (let i = 1; i < dataLog.length; i++) {
        const logKode = String(dataLog[i][2]).trim();
        const logStatus = String(dataLog[i][7]).toLowerCase();

        if (logKode === targetKode && logStatus.includes('sedang dipinjam')) {
           idTrxLama.push(dataLog[i][0]);
           
           // Pecah dokumen yang koma-koma
           const listDocName = String(dataLog[i][5]).split(',');
           
           // B. Untuk setiap dokumen yang dipinjam, cari detailnya di Crew_Doc
           listDocName.forEach(name => {
             const cleanName = name.trim();
             if (cleanName === "") return; // Lewati jika kosong

             let detail = { tipe: cleanName, nomor: '-', tempat: '-', issued: '-', expiry: '-' };
             
             // Cek ke database Crew_Doc jika ada
             if (dataDoc.length > 0) {
                 for (let j = 1; j < dataDoc.length; j++) {
                   const docKode = String(dataDoc[j][1]).trim();
                   const docTipe = String(dataDoc[j][2]).trim();

                   if (docKode === targetKode && docTipe === cleanName) {
                     detail = {
                       tipe: cleanName,
                       nomor: dataDoc[j][3] || '-',
                       tempat: dataDoc[j][4] || '-',
                       issued: dataDoc[j][5] || '-',
                       expiry: dataDoc[j][6] || '-'
                     };
                     break;
                   }
                 }
             }
             docsDetail.push(detail);
           });
        }
      }
      return { trxIds: idTrxLama.join(','), dokumen: docsDetail };
  } catch(e) {
      console.error("Error getDokumenDipinjam: " + e.message);
      return { trxIds: "", dokumen: [] };
  }
}

// ==========================================
// 7. Proses Smart Return (Update Database & Log)
// ==========================================
function prosesSirkulasiKembali(data) {
  const db = getDB();
  const sheetLog = db.getSheetByName('Log_Dokumen');
  const sheetDoc = db.getSheetByName('Crew_Doc');
  
  // A. Update Status di Log_Dokumen
  const dataLog = sheetLog.getDataRange().getValues();
  const ids = data.trxIds.split(',');
  for (let i = 1; i < dataLog.length; i++) {
      if (ids.includes(dataLog[i][0].toString())) {
          sheetLog.getRange(i + 1, 8).setValue('Selesai (Dikembalikan)');
      }
  }

  // B. SMART UPDATE: Update data di Crew_Doc untuk setiap dokumen
  const dataDoc = sheetDoc.getDataRange().getValues();
  data.updateList.forEach(item => {
    for (let j = 1; j < dataDoc.length; j++) {
      if (dataDoc[j][1].toString() === data.kode.toString() && dataDoc[j][2].toString().trim() === item.tipe.trim()) {
        // Update Kolom D, E, F, G (Index 4, 5, 6, 7 di Spreadsheet)
        sheetDoc.getRange(j + 1, 4, 1, 4).setValues([[
          item.nomor, item.tempat, item.issued, item.expiry
        ]]);
        break;
      }
    }
  });

  // C. Catat Log Masuk Baru
  const idTrxBaru = "TRXR-" + new Date().getTime().toString().slice(-6);
  const waktu = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
  const admin = dapatkanNamaAdminAktif(); 
  
  sheetLog.appendRow([
     idTrxBaru, waktu, data.kode, data.nama, "Pengembalian Dokumen", 
     data.updateList.map(d => d.tipe).join(', '), "Update otomatis via Smart Return", "Selesai", admin
  ]);
  
  return "Data dokumen telah diperbarui dan dikembalikan ke brankas!";
}

// ==========================================
// FUNGSI BARU: GENERATOR NOMOR SURAT SIRKULASI
// ==========================================
function getNomorSuratSirkulasi(aksi) {
  const props = PropertiesService.getScriptProperties();
  let urut = props.getProperty('NOMOR_SURAT_SIRKULASI');
  if (!urut) { urut = 1; } else { urut = parseInt(urut) + 1; }

  const formatUrut = urut.toString().padStart(3, '0');
  const date = new Date();
  const romawi = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"];
  const bulanRomawi = romawi[date.getMonth()];
  const tahun = date.getFullYear();

  // Penentuan Kode Tengah Berdasarkan Aksi
  let kodeTengah = "TTPD"; // Default: Pinjam Sementara
  if (aksi === "Pengembalian Dokumen") {
      kodeTengah = "TTP";
  } else if (aksi === "Penarikan Total") {
      kodeTengah = "TTAD";
  }

  const nomorSuratLengkap = `${formatUrut}/${kodeTengah}/ADM/PPM/${bulanRomawi}/${tahun}`;
  return { angkaUrut: urut, nomorLengkap: nomorSuratLengkap };
}

// ==========================================
// 8. GENERATE PDF SIRKULASI (REVISI PARAMETER)
// ==========================================
function cetakPDFSirkulasiBackend(dataTrx, nomorSuratEdit, angkaSimpan, idKopSurat) {
  // Simpan angka urut terbaru ke memori Google
  if(angkaSimpan) {
      PropertiesService.getScriptProperties().setProperty('NOMOR_SURAT_SIRKULASI', angkaSimpan.toString());
  }

  const DRIVE_FOLDER_ID = "1lnLiJsHFK3DixXuldroguymCM2nJGCUT"; 

  // 1. Ambil Kop Surat (Dinamis dari Modal)
  let imgKopHTML = "";
  if (idKopSurat && idKopSurat.trim() !== "") {
    try {
      const fileKop = DriveApp.getFileById(idKopSurat);
      const base64 = Utilities.base64Encode(fileKop.getBlob().getBytes());
      const mime = fileKop.getMimeType();
      imgKopHTML = `<img src="data:${mime};base64,${base64}" style="width: 100%; max-height: 150px; object-fit: contain; margin-bottom: 20px;">`;
    } catch(e) { console.error("Gagal meload Kop Surat: " + e.message); }
  }

  // 2. Ambil Detail Dokumen dari Crew_Doc
  const db = getDB();
  const sheetDoc = db.getSheetByName('Crew_Doc');
  const dataDoc = sheetDoc ? sheetDoc.getDataRange().getDisplayValues() : [];
  
  const listNamaDokumen = dataTrx.dokumen.split(',');
  let barisTabel = "";

  if (listNamaDokumen.length === 0 || dataTrx.dokumen === "") {
    barisTabel = `<tr><td colspan="6" style="text-align: center; padding: 10px;">Tidak ada dokumen spesifik.</td></tr>`;
  } else {
    listNamaDokumen.forEach((namaDoc, index) => {
      const cleanName = namaDoc.trim();
      let dNomor = "-", dTempat = "-", dIssued = "-", dExpiry = "-";
      
      for (let j = 1; j < dataDoc.length; j++) {
        if (dataDoc[j][1].toString() === dataTrx.kode.toString() && dataDoc[j][2].toString().trim() === cleanName) {
           dNomor = dataDoc[j][3] || "-";
           dTempat = dataDoc[j][4] || "-";
           dIssued = dataDoc[j][5] ? formatTanggalIndo(dataDoc[j][5]) : "-";
           dExpiry = dataDoc[j][6] ? formatTanggalIndo(dataDoc[j][6]) : "-";
           break;
        }
      }

      barisTabel += `
        <tr>
          <td style="text-align: center; padding: 5px;">${index + 1}</td>
          <td style="padding: 5px; font-weight: bold;">${cleanName}</td>
          <td style="padding: 5px; text-align: center;">${dNomor}</td>
          <td style="padding: 5px; text-align: center;">${dTempat}</td>
          <td style="padding: 5px; text-align: center;">${dIssued}</td>
          <td style="padding: 5px; text-align: center;">${dExpiry}</td>
        </tr>
      `;
    });
  }

  // 3. Logika Judul & Tanda Tangan
  let judulSurat = "TANDA TERIMA DOKUMEN";
  let kalimatPembuka = "Telah diterima dokumen asli / salinan dengan rincian sebagai berikut:";
  let penyerah = "Pihak Pertama";
  let penerima = "Pihak Kedua";

  if (dataTrx.aksi === "Pinjam Sementara") {
      judulSurat = "TANDA TERIMA PEMINJAMAN DOKUMEN";
      kalimatPembuka = `Pada hari ini, ${dataTrx.waktu.split(' ')[0]}, telah diserahkan dokumen asli milik pelaut untuk dipinjam sementara:`;
      penyerah = dataTrx.admin; 
      penerima = dataTrx.nama;
  } else if (dataTrx.aksi === "Penarikan Total") {
      judulSurat = "BERITA ACARA PENARIKAN DOKUMEN";
      kalimatPembuka = `Pada hari ini, ${dataTrx.waktu.split(' ')[0]}, telah diserahkan kembali seluruh dokumen asli kepada pelaut (Resign/Keluar):`;
      penyerah = dataTrx.admin; 
      penerima = dataTrx.nama;
  } else if (dataTrx.aksi === "Pengembalian Dokumen") {
      judulSurat = "TANDA TERIMA PENGEMBALIAN DOKUMEN";
      kalimatPembuka = `Pada hari ini, ${dataTrx.waktu.split(' ')[0]}, telah diterima kembali dokumen asli dari pelaut ke brankas perusahaan:`;
      penyerah = dataTrx.nama;  
      penerima = dataTrx.admin;
  }

  // 4. Rakit HTML Template (MEMASUKKAN NOMOR SURAT EDIT)
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
        .info-table { width: 70%; margin: 20px auto; border: none; }
        .info-table td { border: none; padding: 4px; text-align: left; }
        .signature-table { width: 100%; border: none; margin-top: 50px; text-align: center; }
        .signature-table td { border: none; width: 50%; }
      </style>
    </head>
    <body>
      <div class="center">
        ${imgKopHTML}
        
        <div class="bold header-title">${judulSurat}</div>
        <div>Nomor: ${nomorSuratEdit}</div>
        
        <table class="info-table">
          <tr><td style="width: 150px;">Nama Pelaut</td><td>: <b>${dataTrx.nama}</b></td></tr>
          <tr><td>ID / Kode Kru</td><td>: ${dataTrx.kode}</td></tr>
          <tr><td>Alasan / Keterangan</td><td>: ${dataTrx.alasan}</td></tr>
        </table>
        
        <p style="text-align: left;">${kalimatPembuka}</p>
        
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
            <td>Yang Menyerahkan,<br><br><br><br><br><b>( ${penyerah} )</b></td>
            <td>Penerima,<br><br><br><br><br><b>( ${penerima} )</b></td>
          </tr>
        </table>
      </div>
    </body>
    </html>
  `;

  // 5. Ubah HTML ke PDF dan Simpan di Drive
  const blob = Utilities.newBlob(htmlTemplate, MimeType.HTML).getAs(MimeType.PDF);
  blob.setName(`${judulSurat} - ${dataTrx.nama}.pdf`);
  
  try {
      const mainFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      const pdfFile = mainFolder.createFile(blob);
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return pdfFile.getDownloadUrl();
  } catch (err) {
      throw new Error("Gagal menyimpan ke Drive. Error: " + err.message);
  }
}

// Fungsi pembantu jika belum ada formatTanggalIndo di global
function formatTanggalIndo(dateStr) {
  if (!dateStr || dateStr === "-" || dateStr === "") return "-";
  // Asumsi format data displayValues Excel adalah string, kembalikan saja langsung
  // Jika ingin diformat khusus, bisa dipecah di sini.
  return dateStr;
}