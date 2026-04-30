// ==========================================
// FITUR MASSAL: ONBOARD, SIGN OFF & PLOTTING (OPTIMIZED)
// ==========================================

function prosesMutasiOnboardMassal(kruList, dataKapal) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCrew = ss.getSheetByName('Crew');
  const sheetDeploy = ss.getSheetByName('Crew_Deploy');
  const sheetSea = ss.getSheetByName('Crew_SeaServices');

  let deployDataToInsert = [];
  let seaDataToInsert = [];

  for(let i = 0; i < kruList.length; i++) {
    let kode = kruList[i].kode;
    let rankOnboard = kruList[i].rank; 
    
    // OPTIMASI: Gunakan TextFinder
    const tf = sheetCrew.getRange("A:A").createTextFinder(kode.toString()).matchEntireCell(true).findNext();
    
    if(tf) {
      let rowIndex = tf.getRow();
      let namaCrew = sheetCrew.getRange(rowIndex, 2).getValue();
      
      sheetCrew.getRange(rowIndex, 13).setValue("Onboard");
      sheetCrew.getRange(rowIndex, 18).setValue(""); 
      catatLogPerubahan(kode, namaCrew, "Mutasi (Onboard)", "Sign On ke kapal " + dataKapal.kapal + " dengan jabatan " + rankOnboard);
      
      let timestamp = new Date().getTime() + i; 
      
      deployDataToInsert.push([
        "DEP-" + timestamp, kode, namaCrew, dataKapal.kapal, dataKapal.bendera, dataKapal.grt, rankOnboard, 
        dataKapal.negara_tujuan, dataKapal.perusahaan, dataKapal.tgl_onboard, "", "", dataKapal.bulan_laporan
      ]);

      seaDataToInsert.push([
        "SEA-" + timestamp, kode, dataKapal.kapal, dataKapal.bendera, dataKapal.grt, rankOnboard, 
        dataKapal.perusahaan, dataKapal.tgl_onboard, "", "", dataKapal.jenis_kapal 
      ]);
    }
  }

  // OPTIMASI: Batch Write (Hanya 1 kali tulis ke sheet)
  if(deployDataToInsert.length > 0) {
    sheetDeploy.getRange(sheetDeploy.getLastRow() + 1, 1, deployDataToInsert.length, deployDataToInsert[0].length).setValues(deployDataToInsert);
    sheetSea.getRange(sheetSea.getLastRow() + 1, 1, seaDataToInsert.length, seaDataToInsert[0].length).setValues(seaDataToInsert);
  }
  
  return "Sukses! " + deployDataToInsert.length + " kru berhasil diberangkatkan.";
}


function prosesMutasiSignOffMassal(kruList, dataSignOff) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCrew = ss.getSheetByName('Crew');
  const sheetDeploy = ss.getSheetByName('Crew_Deploy');
  const sheetSea = ss.getSheetByName('Crew_SeaServices');

  const dataSea = sheetSea.getDataRange().getValues();
  let deployDataToInsert = [];
  let countSukses = 0;

  for(let i = 0; i < kruList.length; i++) {
    let kode = kruList[i].kode;
    let ketCustom = kruList[i].keterangan;

    const tf = sheetCrew.getRange("A:A").createTextFinder(kode.toString()).matchEntireCell(true).findNext();

    if(tf) {
      let rowCrew = tf.getRow();
      let namaCrew = sheetCrew.getRange(rowCrew, 2).getValue();

      // 1. Ubah Status
      sheetCrew.getRange(rowCrew, 13).setValue("Standby");

      // 2. Cari Data Kontrak Aktif di SeaServices
      let rowSea = -1;
      let dataKapal = {};

      for(let j = dataSea.length - 1; j >= 1; j--) {
        if(String(dataSea[j][1]).trim() === String(kode).trim() && dataSea[j][8] == "") {
           rowSea = j + 1;
           dataKapal.kapal = dataSea[j][2];
           dataKapal.bendera = dataSea[j][3];
           dataKapal.grt = dataSea[j][4];
           dataKapal.rank = dataSea[j][5];
           dataKapal.perusahaan = dataSea[j][6];
           dataKapal.tglOn = dataSea[j][7];
           break;
        }
      }

      if(rowSea !== -1) {
         let durasi = hitungSelisihTanggal(dataKapal.tglOn, dataSignOff.tgl_signoff);
         
         // A. TUTUP KONTRAK DI SEASERVICES (Hanya update sel terkait)
         sheetSea.getRange(rowSea, 9, 1, 2).setValues([[dataSignOff.tgl_signoff, durasi]]);

         // --- TAMBAH BARIS INI UNTUK FORMATTING TANGGAL ONBOARD ---
         // Paksa format dataKapal.tglOn menjadi yyyy-mm-dd
         let tglOnFormatted = "";
         if (dataKapal.tglOn) {
            let tglObj = new Date(dataKapal.tglOn);
            if (!isNaN(tglObj.getTime())) {
               // Sesuaikan dengan Timezone GMT+7 (WIB)
               tglOnFormatted = Utilities.formatDate(tglObj, "GMT+7", "yyyy-MM-dd");
            }
         }
         // ---------------------------------------------------------

         // B. TAMBAH KE ARRAY BATCH DEPLOY
         const deployId = "DEP-OFF-" + new Date().getTime() + "-" + i;
         deployDataToInsert.push([
            deployId, kode, namaCrew, dataKapal.kapal, dataKapal.bendera, dataKapal.grt, dataKapal.rank,
            "", dataKapal.perusahaan, tglOnFormatted, dataSignOff.tgl_signoff, ketCustom, dataSignOff.bulan_laporan
            //                        ^^^^^^^^^^^^^^ (Ganti dataKapal.tglOn menjadi tglOnFormatted)
         ]);

         catatLogPerubahan(kode, namaCrew, "Mutasi (Sign Off)", "Turun dari kapal " + dataKapal.kapal + ". Alasan: " + ketCustom);
         countSukses++;
      }
    }
  }

  // OPTIMASI: Menulis semua baris mutasi turun sekaligus!
  if(deployDataToInsert.length > 0) {
    sheetDeploy.getRange(sheetDeploy.getLastRow() + 1, 1, deployDataToInsert.length, deployDataToInsert[0].length).setValues(deployDataToInsert);
  }

  return "Sukses! " + countSukses + " kru berhasil Sign Off dan riwayat telah diperbarui.";
}

function prosesMutasiPlottingMassal(kruList, namaKapal) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCrew = ss.getSheetByName('Crew');
  let count = 0;

  for(let i = 0; i < kruList.length; i++) {
    let kode = kruList[i];
    const tf = sheetCrew.getRange("A:A").createTextFinder(kode.toString()).matchEntireCell(true).findNext();

    if(tf) {
      let rowIndex = tf.getRow();
      let namaKru = sheetCrew.getRange(rowIndex, 2).getValue();

      sheetCrew.getRange(rowIndex, 13).setValue("Plotting");
      let teksCatatan = namaKapal ? "Plotting Kapal: " + namaKapal : "";
      sheetCrew.getRange(rowIndex, 18).setValue(teksCatatan); 

      catatLogPerubahan(kode, namaKru, "Mutasi (Plotting)", "Status diubah menjadi Plotting. Rencana Kapal: " + (namaKapal ? namaKapal : "Belum ditentukan"));
      count++;
    }
  }

  return "Sukses! " + count + " kru berhasil diubah statusnya menjadi Plotting.";
}

function getCrewDeployData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Crew_Deploy');
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  return data.length > 1 ? data.slice(1) : [];
}

// ==========================================
// PENGAMBIL KOP SURAT DARI GOOGLE DRIVE (Taruh di file .gs)
// ==========================================
function getKopSuratBase64() {
  try {
    // INI DIA ID GOOGLE DRIVE ANDA:
    const fileId = '1xAPtnbZBBtFIrN4VEqnwvRk0OfAfQnjN'; 
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());
    const mimeType = blob.getContentType();
    
    return 'data:' + mimeType + ';base64,' + base64;
  } catch (e) {
    throw new Error("Gagal mengambil Kop Surat. Pastikan ID benar dan file sudah disetting 'Siapa saja yang memiliki link dapat melihat'.");
  }
}