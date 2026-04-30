// ==========================================
// FUNGSI BANTUAN: MENDAPATKAN NAMA ADMIN SAAT INI (REAL-TIME)
// ==========================================
function dapatkanNamaAdminAktif() {
  const emailAktif = Session.getActiveUser().getEmail();
  if (!emailAktif) return "Sistem";
  
  // KITA HAPUS CACHE DI SINI AGAR SISTEM SELALU MEMBACA NAMA TERBARU DARI EXCEL
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetAdmin = ss.getSheetByName('Data_Admin');
    if (sheetAdmin) {
      const dataAdmin = sheetAdmin.getDataRange().getValues();
      for (let i = 1; i < dataAdmin.length; i++) {
        if (dataAdmin[i][0].toString().toLowerCase() === emailAktif.toLowerCase()) {
          return dataAdmin[i][2] || emailAktif;
        }
      }
    }
  } catch(e) {}
  
  return emailAktif;
}

// ==========================================
// MESIN PENCATAT RIWAYAT PERUBAHAN (DENGAN DETEKSI USER)
// ==========================================
function catatLogPerubahan(kodeCrew, namaKru, modul, aktivitas) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetLog = ss.getSheetByName('Log_Perubahan');
    if (!sheetLog) return; 
    
    // 1. Dapatkan nama Admin asli yang sedang bertugas
    const namaAdmin = dapatkanNamaAdminAktif();
    
    // 2. KITA PAKSA SELALU TAMPILKAN NAMA ADMIN DI KOLOM AKTIVITAS
    let aktivitasFinal = `[Oleh: ${namaAdmin}] ${aktivitas}`;
    
    // 3. Pastikan kolom "Nama Kru" untuk event Sistem/Login menggunakan nama Admin
    let namaKruFinal = namaKru;
    if (modul === "Sistem") {
        namaKruFinal = namaAdmin;
    }
    
    // 4. Catat waktu saat ini (WIB)
    const waktu = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
    
    // 5. Tambahkan log baru ke baris paling bawah
    sheetLog.appendRow([waktu, kodeCrew, namaKruFinal, modul, aktivitasFinal]);
    
    // 6. Pembersih Otomatis
    const BATAS_MAKSIMAL = 100;
    if (sheetLog.getLastRow() > BATAS_MAKSIMAL + 1) { 
       sheetLog.deleteRow(2); 
    }
  } catch(e) {
    console.error("Gagal mencatat log: " + e.message);
  }
}

// MENGAMBIL DATA LOG UNTUK DITAMPILKAN DI HTML
function getLogPerubahanData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Log_Perubahan');
  
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return []; 
  
  // Ambil data tanpa baris header (index 0), lalu balik urutannya (terbaru di atas)
  return data.slice(1).reverse();
}

// FITUR NOTIFIKASI DOKUMEN EXPIRED (SOP BARU: PP < 18 Bulan, Lainnya < 12 Bulan)
function getNotifikasiExpired() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetDoc = ss.getSheetByName('Crew_Doc');
  const sheetCrew = ss.getSheetByName('Crew');
  
  if (!sheetDoc || !sheetCrew) return [];
  
  const dataDoc = sheetDoc.getDataRange().getDisplayValues();
  const dataCrew = sheetCrew.getDataRange().getDisplayValues();
  
  // Buat kamus (map) Kode Crew -> Nama Crew agar pencarian cepat
  let mapCrew = {};
  for(let i = 1; i < dataCrew.length; i++) {
     mapCrew[dataCrew[i][0]] = dataCrew[i][1];
  }
  
  let notifList = [];
  const today = new Date();
  today.setHours(0,0,0,0);
  
  for(let i = 1; i < dataDoc.length; i++) {
     let expiryStr = dataDoc[i][6]; // Kolom Expiry di Crew_Doc
     let tipeDoc = (dataDoc[i][2] || "").toString().toUpperCase().trim();
     
     if(expiryStr && expiryStr !== "" && expiryStr !== "-") {
        let expiryDate = new Date(expiryStr);
        if(!isNaN(expiryDate.getTime())) {
           
           // Hitung selisih bulan murni (mempertimbangkan tahun dan bulan)
           let diffMonths = (expiryDate.getFullYear() - today.getFullYear()) * 12 + (expiryDate.getMonth() - today.getMonth());
           
           // Kurangi 1 bulan jika tanggal hari ini sudah melewati tanggal kedaluwarsa di bulan tersebut
           if (today.getDate() > expiryDate.getDate()) {
               diffMonths--;
           }

           // Logika SOP: Paspor/PP = 18 Bulan, Sertifikat Lain = 12 Bulan
           let isPP = tipeDoc === "PP" || tipeDoc.includes("PASPOR");
           let batasBulan = isPP ? 18 : 12;

           // Jika sisa bulan KURANG DARI batas SOP, masukkan ke daftar notifikasi!
           if (diffMonths < batasBulan) { 
              let kode = dataDoc[i][1];
              let nama = mapCrew[kode] || "Kru Tidak Ditemukan";
              
              // Hitung total sisa hari untuk menentukan warna merah/kuning di frontend
              let diffDays = Math.ceil((expiryDate.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
              
              notifList.push({
                 kode: kode,
                 nama: nama,
                 tipe: tipeDoc,
                 sisaBulan: diffMonths,
                 sisaHari: diffDays, 
                 batasSOP: batasBulan,
                 tanggal: expiryStr
              });
           }
        }
     }
  }
  
  // Urutkan dari yang paling mendesak (Sisa hari paling kecil / sudah minus)
  notifList.sort((a, b) => a.sisaHari - b.sisaHari);
  return notifList;
}

// FITUR NOTIFIKASI KONTRAK KRU (Batas Peringatan: > 9 Bulan)
function getNotifikasiKontrak() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSea = ss.getSheetByName('Crew_SeaServices');
  const sheetCrew = ss.getSheetByName('Crew');
  
  if (!sheetSea || !sheetCrew) return [];
  
  const dataSea = sheetSea.getDataRange().getDisplayValues();
  const dataCrew = sheetCrew.getDataRange().getDisplayValues();
  
  // Mapping Master Kru untuk memastikan dia berstatus Onboard
  let mapCrew = {};
  for(let i = 1; i < dataCrew.length; i++) {
     mapCrew[dataCrew[i][0]] = {
         nama: dataCrew[i][1],
         status: dataCrew[i][12] // Status di Index 12
     };
  }
  
  let notifList = [];
  const today = new Date();
  today.setHours(0,0,0,0);
  
  // Loop data pengalaman layar
  for(let i = 1; i < dataSea.length; i++) {
     let kode = dataSea[i][1];
     let kapal = dataSea[i][2];
     let signOnStr = dataSea[i][7]; // Index 7: Sign On
     let signOffStr = dataSea[i][8]; // Index 8: Sign Off
     
     let infoKru = mapCrew[kode];
     
     // Syarat aktif: Sign Off Kosong DAN Status Master "Onboard"
     if(signOffStr === "" && signOnStr !== "" && infoKru && infoKru.status.toLowerCase() === 'onboard') {
        let signOnDate = new Date(signOnStr);
        if(!isNaN(signOnDate.getTime())) {
           
           // Hitung selisih hari & bulan
           let diffTime = today.getTime() - signOnDate.getTime();
           let diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
           let diffMonths = diffDays / 30.44; // Rata-rata hari dalam sebulan
           
           // TRIGGER: Jika sudah berlayar lebih dari 9 bulan (Batas persiapan)
           if(diffMonths >= 9) {
              let bulanBulat = Math.floor(diffMonths);
              let sisaHari = Math.floor(diffDays - (bulanBulat * 30.44));
              
              notifList.push({
                 kode: kode,
                 nama: infoKru.nama,
                 kapal: kapal,
                 signOn: signOnStr,
                 durasiBulan: bulanBulat,
                 durasiHari: sisaHari,
                 totalHari: diffDays
              });
           }
        }
     }
  }
  
  // Urutkan dari yang PALING LAMA di kapal (Paling Mendesak)
  notifList.sort((a, b) => b.totalHari - a.totalHari);
  return notifList;
}

// ==========================================
// FITUR NOTIFIKASI DOKUMEN SEDANG DIPINJAM
// ==========================================
function getNotifikasiPinjaman() {
  try {
      const sheetLog = getDB().getSheetByName('Log_Dokumen');
      if (!sheetLog) return [];

      const data = sheetLog.getDataRange().getDisplayValues();
      let listPinjam = [];

      // Looping terbalik agar yang paling baru dipinjam ada di atas
      for (let i = data.length - 1; i > 0; i--) { 
        // Ubah jadi huruf kecil semua untuk pengecekan yang aman
        const statusLog = String(data[i][7]).toLowerCase();
        
        if (statusLog.includes('sedang dipinjam') && data[i][0] !== "") {
          listPinjam.push({
            idTrx: data[i][0],
            waktu: data[i][1],
            kode: data[i][2],
            nama: data[i][3],
            dokumen: data[i][5],
            admin: data[i][8] || "Sistem"
          });
        }
      }
      return listPinjam;
  } catch(e) {
      console.error("Error Notif Pinjaman: " + e.message);
      return [];
  }
}

// ==========================================
// OPTIMASI DASHBOARD: MASTER DATA BATCHING
// Ditaruh di paling bawah Crud_Dashboard.gs
// ==========================================

function getDashboardMasterData() {
  // Ambil statistik kapal (fallback jika fungsi belum ada)
  let statsKapal = { totalKapal: 0, totalPrincipal: 0 };
  try { statsKapal = getVesselStatsDashboard(); } catch(e) {}

  return {
    summaryKru: hitungSummaryKruOptimal(),
    notifExpired: getNotifikasiExpired(),
    notifKontrak: getNotifikasiKontrak(),
    notifPinjaman: getNotifikasiPinjaman(),
    logData: getLogPerubahanData(),
    statsKapal: statsKapal
  };
}

function hitungSummaryKruOptimal() {
  try {
    const sheet = getDB().getSheetByName('Crew');
    if (!sheet) return { tTotal: 0, tOnboard: 0, tStandby: 0, tPlotting: 0 };
    
    const data = sheet.getDataRange().getValues();
    let tTotal = 0, tOnboard = 0, tStandby = 0, tPlotting = 0;

    // Hitung langsung di server, browser tidak perlu load ribuan baris data
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) { // Pastikan baris memiliki Kode Crew
        let statusKru = (data[i][12] || '').toString().trim();
        if (statusKru.toLowerCase() !== 'blacklisted') {
          tTotal++;
          if (statusKru === 'Onboard') tOnboard++;
          else if (statusKru === 'Plotting') tPlotting++;
          else tStandby++;
        }
      }
    }
    return { tTotal: tTotal, tOnboard: tOnboard, tStandby: tStandby, tPlotting: tPlotting };
  } catch(e) {
    return { tTotal: 0, tOnboard: 0, tStandby: 0, tPlotting: 0 };
  }
}