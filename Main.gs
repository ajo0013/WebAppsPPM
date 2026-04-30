function doGet(e) {
  let html = HtmlService.createTemplateFromFile('Index');
  return html.evaluate()
    .setTitle('Crew Management PPM') 
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// ==========================================
// SISTEM OTENTIKASI GOOGLE & ROLE (UPDATE FOTO SIDEBAR)
// ==========================================
function verifikasiLoginGoogle() {
  const email = Session.getActiveUser().getEmail();
  if (!email) return { status: false, pesan: "Tidak dapat mendeteksi email." };

  const sheet = getDB().getSheetByName('Data_Admin');
  if (!sheet) return { status: false, pesan: "Sheet 'Data_Admin' tidak ditemukan." };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === email.toLowerCase()) {
      const role = data[i][1];
      const nama = data[i][2] || email;
      
      // --- SUNTIKAN ANTI-GAMBAR-HILANG UNTUK SIDEBAR ---
      let rawFoto = data[i][3];
      let finalFoto = "";
      
      // Kita gunakan fungsi ekstrakIdDrive yang sudah Bapak pasang sebelumnya
      let fileId = ekstrakIdDrive(rawFoto); 
      
      // Jika itu link Google Drive, ubah jadi link Thumbnail (Anti-Blokir)
      if (fileId !== "") {
        finalFoto = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w500";
      } 
      // Jika sudah berupa link gambar biasa / link ui-avatars
      else if (rawFoto && rawFoto.includes("http")) {
        finalFoto = rawFoto; 
      } 
      // Jika kosong, berikan Avatar inisial
      else {
        finalFoto = "https://ui-avatars.com/api/?name=" + encodeURIComponent(nama) + "&background=random";
      }

      return { 
          status: true, 
          role: role, 
          nama: nama, 
          email: email, 
          foto: finalFoto // <- Sekarang URL Foto yang dikirim ke Sidebar sudah matang!
      };
    }
  }
  return { status: false, pesan: "Akses ditolak. Email Anda tidak terdaftar sebagai Admin/HR." };
}

// ==========================================
// FUNGSI PEMBANTU: BONGKAR ID DARI LINK DRIVE
// ==========================================
function ekstrakIdDrive(url) {
  if (!url || typeof url !== 'string' || url === "") return "";
  let id = "";
  if (url.includes("/file/d/")) {
    id = url.split("/file/d/")[1].split("/")[0];
  } else if (url.includes("?id=")) {
    id = url.split("?id=")[1].split("&")[0];
  }
  return id;
}

// ==========================================
// BACKEND: PENGATURAN PENGGUNA (ADMIN ONLY)
// ==========================================
function getAdminUsers() {
  const sheet = getDB().getSheetByName('Data_Admin');
  const data = sheet.getDataRange().getValues();
  data.shift(); // Buang header
  
  return data.map((row, index) => {
    let rawFoto = row[3];
    let finalFoto = "";
    let fileId = ekstrakIdDrive(rawFoto);
    
    // Jika itu link Google Drive, ubah jadi link Thumbnail (Anti-Blokir)
    if (fileId !== "") {
      finalFoto = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w500";
    } 
    // Jika sudah berupa link gambar biasa / link ui-avatars
    else if (rawFoto && rawFoto.includes("http")) {
      finalFoto = rawFoto; 
    } 
    // Jika kosong, berikan Avatar inisial
    else {
      finalFoto = "https://ui-avatars.com/api/?name=" + encodeURIComponent(row[2] || row[0]) + "&background=random";
    }

    return {
      baris: index + 2, 
      email: row[0], 
      role: row[1], 
      nama: row[2], 
      foto: finalFoto
    };
  }).filter(user => user.email !== "");
}

// ==========================================
// BACKEND: SIMPAN & HAPUS PENGGUNA (ADMIN ONLY)
// ==========================================
function simpanAdminUser(email, role, nama, fotoUrl) {
  const sheet = getDB().getSheetByName('Data_Admin');
  const data = sheet.getDataRange().getValues();
  
  // Cek apakah email sudah ada (Update)
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === email.toLowerCase()) {
      sheet.getRange(i + 1, 2, 1, 3).setValues([[role, nama, fotoUrl]]);
      catatLogPerubahan("-", nama, "Sistem", `Mengubah data akses pengguna: ${email}`);
      return "Data pengguna berhasil diperbarui!";
    }
  }
  
  // Jika tidak ada, tambah baru
  sheet.appendRow([email, role, nama, fotoUrl]);
  catatLogPerubahan("-", nama, "Sistem", `Menambahkan akses pengguna baru: ${email}`);
  return "Pengguna baru berhasil ditambahkan!";
}

function hapusAdminUser(baris, email) {
  const sheet = getDB().getSheetByName('Data_Admin');
  sheet.deleteRow(baris);
  catatLogPerubahan("-", email, "Sistem", `Menghapus akses pengguna: ${email}`);
  return "Akses pengguna berhasil dicabut!";
}

/**
 * FUNGSI GENERATE PDF MCU MANUAL
 * Taruh di paling bawah file Main.gs 1xAPtnbZBBtFIrN4VEqnwvRk0OfAfQnjN
 */
function buatPDFMCUManual(d) {
  try {
    var ID_KOP_SURAT = "1xAPtnbZBBtFIrN4VEqnwvRk0OfAfQnjN"; // Ganti dengan ID Kop Surat Bapak
    var ID_FOTO_TTD  = "1bjEuciP73K-GjI_QWBSiaA8tLCJPt614"; 
    var ID_FOLDER_PDF = "1lnLiJsHFK3DixXuldroguymCM2nJGCUT"; 

    var doc = DocumentApp.create("TEMP_MCU_" + d.nama.toUpperCase());
    var body = doc.getBody();
    
    // Set Margin & Ukuran A4
    var lebarKertas = 595.27;
    var marginSamping = 40;
    body.setPageHeight(841.89).setPageWidth(lebarKertas);
    body.setMarginTop(30).setMarginBottom(30).setMarginLeft(marginSamping).setMarginRight(marginSamping);

    // 1. KOP SURAT (Otomatis selebar margin cetak)
    try {
      var kopImg = DriveApp.getFileById(ID_KOP_SURAT).getBlob();
      var kop = body.insertImage(0, kopImg);
      var ratioKop = kop.getHeight() / kop.getWidth();
      
      var lebarMaksimal = lebarKertas - (marginSamping * 2); // Menghitung lebar area cetak (515.27 pt)
      kop.setWidth(lebarMaksimal); 
      kop.setHeight(lebarMaksimal * ratioKop); 
      body.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(0).setSpacingBefore(0);
    } catch(e) { 
      body.appendParagraph("PT PANCA PASIFIK MANDIRI").setBold(true).setFontSize(16).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    }

    body.appendHorizontalRule();

    // 2. PERIHAL & NOMOR
    var pPerihal = body.appendParagraph("\nPerihal: SURAT PENGANTAR MEDICAL CHECK UP");
    pPerihal.setBold(true).setSpacingAfter(0).setLineSpacing(1.0);
    
    var pNomor = body.appendParagraph("Nomor: " + d.noSurat);
    pNomor.setSpacingAfter(10).setLineSpacing(1.0);

    // 3. TUJUAN
    body.appendParagraph("Kepada Yth,").setSpacingAfter(0).setLineSpacing(1.0);
    body.appendParagraph("Oilia Medical Centre").setBold(true).setSpacingAfter(0).setLineSpacing(1.0);
    body.appendParagraph("Jl. Enggano No.11 O Blok. C, Tanjung Priok\nJakarta Utara").setSpacingAfter(10).setLineSpacing(1.0);

    body.appendParagraph("Dengan Hormat,").setSpacingAfter(5).setLineSpacing(1.0);
    body.appendParagraph("Yang bertanda tangan dibawah ini:").setSpacingAfter(5).setLineSpacing(1.0);

    // 4. DATA PENGIRIM (DIANA)
    var tabAdmin = body.appendTable([
      ["Nama", ": Diana Sugiyanthi, S.T., M.M"],
      ["Jabatan", ": Direktur Utama"],
      ["Perusahaan", ": PT Panca Pasifik Mandiri"]
    ]);
    aturTabelRapat(tabAdmin, 130); // Angka 130 adalah lebar kolom kiri

    // 5. ISI SURAT & DATA KRU
    body.appendParagraph("\nMemohon untuk dilakukannya Medical Check Up " + d.tipeMedical + " kepada kru dengan detail sebagai berikut:")
        .setSpacingBefore(5).setSpacingAfter(5).setLineSpacing(1.0);
    
    var tabKru = body.appendTable([
      ["Nama Lengkap", ": " + d.nama],
      ["Tempat, Tanggal Lahir", ": " + d.ttl],
      ["No. Passport", ": " + d.paspor],
      ["NIK", ": " + d.nik],
      ["Rank", ": " + d.rank]
    ]);
    aturTabelRapat(tabKru, 130); // Angka 130 menjepit teks ke kiri

    // 6. PENUTUP
    body.appendParagraph("\nDemikian surat pengantar ini kami sampaikan untuk dapat dipergunakan sebagaimana mestinya. Atas perhatian dan kerjasamanya kami ucapkan terima kasih.")
        .setSpacingBefore(5).setLineSpacing(1.0);
    
    // 7. TANDA TANGAN
    var tglStr = Utilities.formatDate(new Date(), "GMT+7", "dd MMMM yyyy");
    body.appendParagraph("\nBogor, " + tglStr).setSpacingAfter(0).setLineSpacing(1.0);
    body.appendParagraph("PT Panca Pasifik Mandiri").setSpacingAfter(0).setLineSpacing(1.0);

    try {
      var ttdBlob = DriveApp.getFileById(ID_FOTO_TTD).getBlob();
      var ttdImg = body.appendImage(ttdBlob);
      var ratio = ttdImg.getHeight() / ttdImg.getWidth();
      ttdImg.setWidth(130); 
      ttdImg.setHeight(130 * ratio);
    } catch(e) { body.appendParagraph("\n\n"); }

    body.appendParagraph("Diana Sugiyanthi, S.T., M.M")
        .setBold(true).setUnderline(true).setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1.0);
    body.appendParagraph("Direktur Utama").setSpacingBefore(0).setLineSpacing(1.0);

    // 8. KONVERSI PDF
    doc.saveAndClose();
    Utilities.sleep(1500);
    var pdfBlob = DriveApp.getFileById(doc.getId()).getBlob().getAs(MimeType.PDF);
    var pdfFile = DriveApp.getFolderById(ID_FOLDER_PDF).createFile(pdfBlob);
    pdfFile.setName("MCU - " + d.nama.toUpperCase() + ".pdf");
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    DriveApp.getFileById(doc.getId()).setTrashed(true);

    return "https://drive.google.com/file/d/" + pdfFile.getId() + "/preview";

  } catch (err) {
    throw new Error("Gagal: " + err.message);
  }
}

/**
 * FUNGSI MERAPATKAN TABEL & MENJEPIT KOLOM KIRI
 */
function aturTabelRapat(table, col0Width) {
  table.setBorderWidth(0);
  for (var i = 0; i < table.getNumRows(); i++) {
    var row = table.getRow(i);
    
    // Menjepit kolom kiri agar titik dua bergeser ke kiri
    if (col0Width) {
      row.getCell(0).setWidth(col0Width);
    }

    for (var j = 0; j < row.getNumCells(); j++) {
      var cell = row.getCell(j);
      cell.setPaddingTop(0);    
      cell.setPaddingBottom(0); 
      cell.setPaddingLeft(0);  // Hilangkan jarak dari batas kiri pinggir tabel
      cell.setPaddingRight(0);
      var para = cell.getChild(0).asParagraph();
      para.setSpacingBefore(0);
      para.setSpacingAfter(0);
      para.setLineSpacing(1.0); 
    }
  }
}