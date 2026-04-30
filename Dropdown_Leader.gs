// OPTIMASI: Menyimpan referensi di Cache selama 6 jam agar Dasbor instan
function getJabatanList() {
  const cache = CacheService.getScriptCache();
  if (cache.get('jabatanList')) return JSON.parse(cache.get('jabatanList'));

  const sheet = getDB().getSheetByName('Jabatan_Reference');
  if (!sheet) return []; 
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; 
  const jabatanList = [];
  for (let i = 1; i < data.length; i++) { if (data[i][1] !== '') jabatanList.push(data[i][1]); }
  
  cache.put('jabatanList', JSON.stringify(jabatanList), 21600);
  return jabatanList;
}

function getStatusList() {
  const cache = CacheService.getScriptCache();
  if (cache.get('statusList')) return JSON.parse(cache.get('statusList'));

  const sheet = getDB().getSheetByName('Status_Reference');
  if (!sheet) return []; 
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const statusList = [];
  for (let i = 1; i < data.length; i++) { if (data[i][0] !== '') statusList.push(data[i][0]); }
  
  cache.put('statusList', JSON.stringify(statusList), 21600);
  return statusList;
}

function bersihkanCacheManual() {
  const cache = CacheService.getScriptCache();
  cache.remove('statusList');
  cache.remove('jabatanList');
  Logger.log("Cache telah dihapus. Silakan refresh halaman web Anda.");
}