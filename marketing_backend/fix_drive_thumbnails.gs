/**
 * 一鍵把 Drive 視覺底圖資料夾裡所有圖檔改成「知道連結的人都能檢視」
 * 解決後台縮圖載不出來的問題
 * 
 * 用法：
 *   1. 貼進 Apps Script
 *   2. 執行 fixDriveThumbnails()
 *   3. 完成後後台 F5、縮圖就出來
 */

const FIX_FOLDER_ID = '15rYYJ5ZEJ6rTYEy_QthJuEAf6OU652D9';

function fixDriveThumbnails() {
  const folder = DriveApp.getFolderById(FIX_FOLDER_ID);
  
  // Step 1：把資料夾本身改公開
  try {
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    Logger.log('✅ 資料夾已設為「知道連結的人都能檢視」');
  } catch (e) {
    Logger.log('⚠️ 資料夾權限設定失敗：' + e.message);
  }

  // Step 2：遞迴改所有圖檔（含子資料夾）
  let processed = 0;
  let errors = [];
  
  processFolder_(folder, processed, errors);
  
  function processFolder_(f, _, __) {
    // 處理當前資料夾的所有檔
    const files = f.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        processed++;
        if (processed % 20 === 0) {
          Logger.log(`進度：${processed} 個檔已改公開`);
          Utilities.sleep(200);
        }
      } catch (e) {
        errors.push(`${file.getName()}: ${e.message}`);
      }
    }
    
    // 遞迴子資料夾
    const subs = f.getFolders();
    while (subs.hasNext()) {
      processFolder_(subs.next(), 0, []);
    }
  }
  
  Logger.log(`====== 完成 ======`);
  Logger.log(`✅ 已改公開的檔案：${processed}`);
  Logger.log(`❌ 錯誤：${errors.length}`);
  if (errors.length) errors.slice(0, 5).forEach(e => Logger.log(e));
  
  SpreadsheetApp.getUi().alert(
    `Drive 權限修補完成\n\n` +
    `已改公開：${processed} 個檔\n` +
    `錯誤：${errors.length}\n\n` +
    `請到後台 F5，縮圖應該會出來。`
  );
}


/**
 * 測試用：只改一個檔，驗證 setSharing API 能用
 */
function testFixOneFile() {
  // 拿資料夾的第一個檔
  const folder = DriveApp.getFolderById(FIX_FOLDER_ID);
  const files = folder.getFiles();
  if (!files.hasNext()) {
    Logger.log('資料夾是空的，去子資料夾找');
    const subs = folder.getFolders();
    if (subs.hasNext()) {
      const sub = subs.next();
      const subFiles = sub.getFiles();
      if (subFiles.hasNext()) {
        const file = subFiles.next();
        Logger.log('找到檔：' + file.getName() + ' (ID: ' + file.getId() + ')');
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        Logger.log('✅ 已改公開');
        Logger.log('請打開瀏覽器無痕視窗測這 URL 是否能看到圖：');
        Logger.log('https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1600');
      }
    }
    return;
  }
  const file = files.next();
  Logger.log('找到檔：' + file.getName() + ' (ID: ' + file.getId() + ')');
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  Logger.log('✅ 已改公開');
  Logger.log('測試縮圖 URL: https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1600');
}
