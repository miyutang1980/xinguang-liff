/**
 * 一次設好「實景照進件_2026」父資料夾權限
 * 設成「擁有連結的人 - 編輯者」、子資料夾自動繼承
 * 跑一次就好、之後不會再跳「需要存取權」
 */
function shareIntakeFolders() {
  // 找父資料夾(在你 Drive 根目錄)
  const folders = DriveApp.getFoldersByName('實景照進件_2026');
  if (!folders.hasNext()) {
    SpreadsheetApp.getUi().alert('找不到「實景照進件_2026」資料夾、請先跑 buildIntakeFolders');
    return;
  }
  const parent = folders.next();
  
  // 設父資料夾為「任何人有連結都能編輯」
  try {
    parent.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
    Logger.log('父資料夾權限設定完成: ' + parent.getName() + ' (' + parent.getId() + ')');
  } catch (e) {
    Logger.log('父資料夾失敗: ' + e.message);
  }
  
  // 同時逐一把 79 個子資料夾也設一次(保險)
  const subFolders = parent.getFolders();
  let count = 0;
  let failed = 0;
  while (subFolders.hasNext()) {
    const f = subFolders.next();
    try {
      f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
      count++;
    } catch (e) {
      failed++;
      Logger.log('子資料夾失敗 ' + f.getName() + ': ' + e.message);
    }
    if ((count + failed) % 10 === 0) Utilities.sleep(500); // 避免配額
  }
  
  const msg = '完成: ' + count + ' 個資料夾權限設定成功' + (failed ? '、' + failed + ' 個失敗' : '') +
              '\n\n父資料夾URL: ' + parent.getUrl() +
              '\n\n之後從後台「📁 進件」開啟、不會再跳需要存取權';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}
