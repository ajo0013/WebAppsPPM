const SPREADSHEET_ID = '1mDzbZSXrlLUrXKsy70kEj_FGeetGDxiy7Q7KioDiD_c'; 
const DRIVE_FOLDER_ID = '1lnLiJsHFK3DixXuldroguymCM2nJGCUT';
const BLACKLIST_FOLDER_ID = '1PKIWBJWMP1z5n8Wwcm4MWCs8Obf0nXxL';
function getDB() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}