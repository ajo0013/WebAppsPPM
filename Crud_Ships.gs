// ==========================================
// BAGIAN 1: MASTER PRINCIPAL (AGENCY)
// ==========================================
function getPrincipalsData() {
  try {
    const sheet = getDB().getSheetByName('Master_Principal');
    if (!sheet) return [];
    const data = sheet.getDataRange().getDisplayValues();
    return data.length > 1 ? data.slice(1) : [];
  } catch(e) { return []; }
}

function simpanPrincipalData(dataForm) {
  try {
    const sheet = getDB().getSheetByName('Master_Principal');
    if (!sheet) throw new Error("Sheet 'Master_Principal' tidak ditemukan!");

    const isUpdate = dataForm.id_principal && dataForm.id_principal.trim() !== "";
    const idPrnc = isUpdate ? dataForm.id_principal : "PRC-" + new Date().getTime().toString().slice(-6);
    
    // Susunan: ID, Nama, Negara, PIC, Kontak, Kontrak, Status
    const rowData = [
      idPrnc, 
      dataForm.nama_principal.toUpperCase(), 
      dataForm.negara.toUpperCase(),
      dataForm.pic, 
      dataForm.kontak, 
      dataForm.kontrak, 
      dataForm.status
    ];

    if (isUpdate) {
      // OPTIMASI: Gunakan TextFinder untuk pencarian kilat
      const tf = sheet.getRange("A:A").createTextFinder(idPrnc.toString()).matchEntireCell(true).findNext();
      if (tf) {
        sheet.getRange(tf.getRow(), 1, 1, rowData.length).setValues([rowData]);
        catatLogPerubahan("-", dataForm.nama_principal, "Data Principal", "Mengupdate data agency/principal");
        return "Data Principal berhasil diperbarui!";
      }
      throw new Error("Data Principal tidak ditemukan di database.");
    } else {
      sheet.appendRow(rowData);
      catatLogPerubahan("-", dataForm.nama_principal, "Data Principal", "Menambahkan agency/principal baru");
      return "Berhasil menambahkan Principal baru!";
    }
  } catch(e) { throw new Error(e.message); }
}

function hapusPrincipalData(idPrnc, namaPrnc) {
  const sheet = getDB().getSheetByName('Master_Principal');
  const tf = sheet.getRange("A:A").createTextFinder(idPrnc.toString()).matchEntireCell(true).findNext();
  if (tf) {
    sheet.deleteRow(tf.getRow());
    catatLogPerubahan("-", namaPrnc, "Data Principal", "Menghapus data agency/principal");
    return `Principal ${namaPrnc} berhasil dihapus!`;
  }
  throw new Error("ID Principal tidak ditemukan.");
}

// ==========================================
// BAGIAN 2: MASTER KAPAL (VESSELS)
// ==========================================
function getVesselsData() {
  try {
    const sheet = getDB().getSheetByName('Master_Vessel');
    if (!sheet) return [];
    const data = sheet.getDataRange().getDisplayValues();
    return data.length > 1 ? data.slice(1) : [];
  } catch(e) { return []; }
}

function simpanVesselData(dataForm) {
  try {
    const sheet = getDB().getSheetByName('Master_Vessel');
    if (!sheet) throw new Error("Sheet 'Master_Vessel' tidak ditemukan!");

    const isUpdate = dataForm.id_kapal && dataForm.id_kapal.trim() !== "";
    const idKapal = isUpdate ? dataForm.id_kapal : "VSL-" + new Date().getTime().toString().slice(-6);
    
    // Susunan: ID, Nama, Principal, Jenis, Bendera, GRT, Kapasitas, Status, Tahun Kapal
    const rowData = [
      idKapal, 
      dataForm.nama_kapal.toUpperCase(), 
      dataForm.principal.toUpperCase(),
      dataForm.jenis_kapal, 
      dataForm.bendera.toUpperCase(), 
      dataForm.grt,
      dataForm.kapasitas, 
      dataForm.status, 
      dataForm.tahun_kapal
    ];

    if (isUpdate) {
      // OPTIMASI: Gunakan TextFinder untuk pencarian kilat
      const tf = sheet.getRange("A:A").createTextFinder(idKapal.toString()).matchEntireCell(true).findNext();
      if (tf) {
        sheet.getRange(tf.getRow(), 1, 1, rowData.length).setValues([rowData]);
        catatLogPerubahan("-", dataForm.nama_kapal, "Data Kapal", "Mengupdate data spesifikasi kapal");
        return "Data Kapal berhasil diperbarui!";
      }
      throw new Error("Kapal tidak ditemukan di database.");
    } else {
      sheet.appendRow(rowData);
      catatLogPerubahan("-", dataForm.nama_kapal, "Data Kapal", "Menambahkan armada kapal baru");
      return "Berhasil menambahkan armada kapal baru!";
    }
  } catch(e) { throw new Error(e.message); }
}

function hapusVesselData(idKapal, namaKapal) {
  const sheet = getDB().getSheetByName('Master_Vessel');
  const tf = sheet.getRange("A:A").createTextFinder(idKapal.toString()).matchEntireCell(true).findNext();
  if (tf) {
    sheet.deleteRow(tf.getRow());
    catatLogPerubahan("-", namaKapal, "Data Kapal", "Menghapus data kapal");
    return `Kapal ${namaKapal} berhasil dihapus!`;
  }
  throw new Error("ID Kapal tidak ditemukan.");
}

// ==========================================
// BAGIAN 3: FUNGSI RELASIONAL (PENGHUBUNG)
// ==========================================
// Fungsi ini akan dipanggil oleh UI Kapal untuk membuat Dropdown nama Principal
function getDropdownPrincipalAktif() {
  const sheet = getDB().getSheetByName('Master_Principal');
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  let list = [];
  
  for (let i = 1; i < data.length; i++) {
    // Index 6 adalah Kolom G (Status). Kita hanya ambil Principal yang AKTIF
    if(String(data[i][6]).toUpperCase() === "AKTIF") {
      list.push(data[i][1]); // Masukkan namanya
    }
  }
  return list.sort(); // Urutkan sesuai abjad
}

// ==========================================
// BAGIAN 4: KONEKSI DATA KRU ONBOARD
// ==========================================
function getKruOnboardByKapal(namaKapal) {
  try {
    const db = getDB();
    const sheetSea = db.getSheetByName('Crew_SeaServices');
    const sheetCrew = db.getSheetByName('Crew');

    if (!sheetSea || !sheetCrew) throw new Error("Database Kru/Layar tidak ditemukan.");

    const dataSea = sheetSea.getDataRange().getDisplayValues();
    const dataCrew = sheetCrew.getDataRange().getDisplayValues();

    // 1. Buat Peta Status Kru (Hanya ambil yang statusnya ONBOARD)
    let mapKruAktif = {};
    for (let i = 1; i < dataCrew.length; i++) {
        const kode = String(dataCrew[i][0]).trim();
        const status = String(dataCrew[i][12]).toUpperCase().trim(); // Kolom M
        if (kode !== "" && status === "ONBOARD") {
            mapKruAktif[kode] = { nama: dataCrew[i][1] };
        }
    }

    // 2. Cari di Riwayat Layar yang belum Sign-Off
    let listKruOnboard = [];
    for (let i = 1; i < dataSea.length; i++) {
        const kodeKru = String(dataSea[i][1]).trim();
        const kapal = String(dataSea[i][2]).trim().toUpperCase();
        const signOff = String(dataSea[i][8]).trim(); // Kolom I (Sign Off)

        // Jika Kapal cocok, dan belum Sign Off (Kosong), dan status profilnya ONBOARD
        if (kapal === String(namaKapal).toUpperCase().trim() && signOff === "") {
            if (mapKruAktif[kodeKru]) {
                listKruOnboard.push({
                    kode: kodeKru,
                    nama: mapKruAktif[kodeKru].nama,
                    jabatan: dataSea[i][5] || "-",
                    sign_on: dataSea[i][7] || "-"
                });
            }
        }
    }
    return listKruOnboard;
  } catch(e) {
    return { error: e.message };
  }
}

// ==========================================
// BAGIAN 5: KONEKSI DATA PRINCIPAL -> KAPAL -> KRU
// ==========================================
function getArmadaDanKruByPrincipal(namaPrincipal) {
  try {
    const db = getDB();
    const sheetVessel = db.getSheetByName('Master_Vessel');
    const sheetSea = db.getSheetByName('Crew_SeaServices');
    const sheetCrew = db.getSheetByName('Crew');

    if (!sheetVessel || !sheetSea || !sheetCrew) throw new Error("Database tidak lengkap.");

    const dataVessel = sheetVessel.getDataRange().getDisplayValues();
    const dataSea = sheetSea.getDataRange().getDisplayValues();
    const dataCrew = sheetCrew.getDataRange().getDisplayValues();

    // 1. Cari semua kapal milik Principal ini
    let mapKapal = {}; // Menyimpan data kapal agar cepat dicari
    let totalKapasitas = 0;
    let listKapal = [];

    for (let i = 1; i < dataVessel.length; i++) {
        const principalTabel = String(dataVessel[i][2]).trim().toUpperCase();
        if (principalTabel === String(namaPrincipal).toUpperCase().trim()) {
            const namaKapal = String(dataVessel[i][1]).trim().toUpperCase();
            const kapasitas = parseInt(dataVessel[i][6]) || 0;
            
            mapKapal[namaKapal] = {
                nama_asli: dataVessel[i][1],
                jenis: dataVessel[i][3] || "-",
                kapasitas: kapasitas
            };
            listKapal.push(mapKapal[namaKapal]);
            totalKapasitas += kapasitas;
        }
    }

    // 2. Buat Peta Status Kru (Hanya ambil yang statusnya ONBOARD)
    let mapKruAktif = {};
    for (let i = 1; i < dataCrew.length; i++) {
        const kode = String(dataCrew[i][0]).trim();
        const status = String(dataCrew[i][12]).toUpperCase().trim();
        if (kode !== "" && status === "ONBOARD") {
            mapKruAktif[kode] = { nama: dataCrew[i][1] };
        }
    }

    // 3. Cari kru di SeaServices yang berada di kapal-kapal milik Principal ini
    let listKruOnboard = [];
    for (let i = 1; i < dataSea.length; i++) {
        const kodeKru = String(dataSea[i][1]).trim();
        const kapalSea = String(dataSea[i][2]).trim().toUpperCase();
        const signOff = String(dataSea[i][8]).trim(); 

        // Jika kapalnya ada di daftar milik Principal, Sign Off kosong, dan profilnya ONBOARD
        if (mapKapal[kapalSea] && signOff === "" && mapKruAktif[kodeKru]) {
            listKruOnboard.push({
                kode: kodeKru,
                nama: mapKruAktif[kodeKru].nama,
                kapal: mapKapal[kapalSea].nama_asli, // Gunakan nama asli kapal
                jabatan: dataSea[i][5] || "-",
                sign_on: dataSea[i][7] || "-"
            });
        }
    }

    return {
        kapal: listKapal,
        kru: listKruOnboard,
        totalKapasitas: totalKapasitas,
        totalKru: listKruOnboard.length
    };
  } catch(e) {
    return { error: e.message };
  }
}

// ==========================================
// BAGIAN 6: PENYEDIA DROPDOWN KAPAL UNTUK FORM MUTASI
// ==========================================
function getDropdownKapalAktif() {
  try {
    const sheet = getDB().getSheetByName('Master_Vessel');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getDisplayValues();
    let listKapal = [];
    
    for (let i = 1; i < data.length; i++) {
      const namaKapal = String(data[i][1]).trim();
      const principal = String(data[i][2]).trim();
      const status = String(data[i][7]).toUpperCase().trim(); // Kolom H (Status)
      
      // Hanya ambil kapal yang statusnya AKTIF
      if (status === "AKTIF" && namaKapal !== "") {
        listKapal.push(`${namaKapal} - ${principal}`);
      }
    }
    
    return listKapal.sort(); // Urutkan sesuai abjad A-Z
  } catch (e) {
    return [];
  }
}

// ==========================================
// BAGIAN 7: STATISTIK UNTUK DASHBOARD UTAMA
// ==========================================
function getVesselStatsDashboard() {
  const db = getDB();
  let totalKapalAktif = 0;
  let totalPrincipalAktif = 0;

  // 1. Hitung Mitra Principal yang statusnya AKTIF
  try {
    const sheetPrnc = db.getSheetByName('Master_Principal');
    if (sheetPrnc) {
      const dataPrnc = sheetPrnc.getDataRange().getDisplayValues();
      for (let i = 1; i < dataPrnc.length; i++) {
         // Kolom G (index 6) adalah Status Principal
         if (String(dataPrnc[i][6]).toUpperCase() === "AKTIF") totalPrincipalAktif++;
      }
    }
  } catch(e) {}

  // 2. Hitung Armada Kapal yang statusnya AKTIF
  try {
    const sheetVsl = db.getSheetByName('Master_Vessel');
    if (sheetVsl) {
      const dataVsl = sheetVsl.getDataRange().getDisplayValues();
      for (let i = 1; i < dataVsl.length; i++) {
         // Kolom H (index 7) adalah Status Kapal
         if (String(dataVsl[i][7]).toUpperCase() === "AKTIF") totalKapalAktif++;
      }
    }
  } catch(e) {}

  return { totalKapal: totalKapalAktif, totalPrincipal: totalPrincipalAktif };
}