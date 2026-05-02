/**
 * Layer 3：實景照進件掃描器
 *
 * 每 5 分鐘觸發一次：
 * 1. 讀 Posting_Queue 全表(95 列)
 * 2. 對每列 W 欄(進件資料夾URL)、找對應 Drive 資料夾
 * 3. 掃資料夾裡未處理的圖檔(檔名沒有「[已收]_」前綴)
 * 4. 看 Y 欄(發布類型):
 *    - 'single' 或空白 → 取第 1 張寫進 G/H 欄、其他存到 Z 欄(備檔池)
 *    - 'carousel'      → 全部(最多10張)寫進 Z 欄候選池、第1張寫 G/H、Y 欄寫所有 file_ids
 * 5. 處理完的檔案前綴加「[已收]_」避免重複
 * 6. 更新狀態欄 P=「圖片就位」、備註欄 V 加註
 *
 * 新欄結構:
 *   X(24)  發布類型: single / carousel
 *   Y(25)  輪播圖 file_ids (逗號分隔、最多10張、依顯示順序)
 *   Z(26)  備檔 file_ids (逗號分隔、未上稿的)
 */

const SCAN_SS_ID = '1DybgWBdCyvkEijMyaE46rKLtQD9J2ImjU8xeYCKSKnA';
const SCAN_QUEUE_NAME = '排程佇列 Posting_Queue';
const SCAN_VISUAL_FOLDER_ID = '15rYYJ5ZEJ6rTYEy_QthJuEAf6OU652D9'; // 視覺底圖總資料夾
const SCAN_PROCESSED_PREFIX = '[已收]_';
const SCAN_MAX_CAROUSEL = 10;

/* =========================================================
 *  欄位確保: 確認 X/Y/Z 欄頭存在
 * ========================================================= */
function ensureScanColumns_() {
  const sh = SpreadsheetApp.openById(SCAN_SS_ID).getSheetByName(SCAN_QUEUE_NAME);
  const headers = sh.getRange(1, 1, 1, Math.max(26, sh.getLastColumn())).getValues()[0];
  if (headers[23] !== '發布類型') sh.getRange(1, 24).setValue('發布類型');
  if (headers[24] !== '輪播圖file_ids') sh.getRange(1, 25).setValue('輪播圖file_ids');
  if (headers[25] !== '備檔file_ids') sh.getRange(1, 26).setValue('備檔file_ids');
  return sh;
}

/* =========================================================
 *  主掃描器
 * ========================================================= */
function scanIntakeFolders() {
  const sh = ensureScanColumns_();
  const last = sh.getLastRow();
  if (last < 2) return;

  // 讀 26 欄全表
  const data = sh.getRange(2, 1, last - 1, 26).getValues();
  let scanned = 0, found = 0, errors = 0;
  const startTs = new Date().getTime();

  for (let i = 0; i < data.length; i++) {
    // 6 分鐘逾時保護
    if (new Date().getTime() - startTs > 5 * 60 * 1000) {
      Logger.log('逾時、處理到列 ' + (i + 2));
      break;
    }

    const r = data[i];
    const rowNum = i + 2;
    const folderUrl = r[22];        // W 進件資料夾URL
    const status = r[15];           // P 排程狀態
    const publishType = r[23] || 'single'; // X 發布類型(預設 single)

    if (!folderUrl) continue;
    if (status === '已發布' || status === '已排程') continue;

    scanned++;
    try {
      const result = scanOneFolder_(sh, rowNum, folderUrl, publishType);
      if (result.found) found++;
    } catch (e) {
      errors++;
      Logger.log('列 ' + rowNum + ' 失敗: ' + e.message);
      sh.getRange(rowNum, 20).setValue('掃描失敗: ' + e.message); // T 錯誤訊息
    }
  }

  Logger.log('掃描完成: 掃 ' + scanned + ' 列、收圖 ' + found + ' 列、錯 ' + errors + ' 列');
}

/* =========================================================
 *  掃單一資料夾
 * ========================================================= */
function scanOneFolder_(sh, rowNum, folderUrl, publishType) {
  const folderId = (folderUrl.match(/folders\/([-\w]+)/) || [])[1];
  if (!folderId) return { found: false, reason: '無法解析資料夾 ID' };

  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  const newFiles = [];
  while (files.hasNext()) {
    const f = files.next();
    const name = f.getName();
    // 跳過已處理的
    if (name.indexOf(SCAN_PROCESSED_PREFIX) === 0) continue;
    // 只處理圖檔
    const mime = f.getMimeType();
    if (mime.indexOf('image/') !== 0) continue;
    newFiles.push(f);
  }

  if (newFiles.length === 0) return { found: false };

  // 依檔名排序
  newFiles.sort(function(a, b) {
    return a.getName().localeCompare(b.getName());
  });

  // 視覺底圖總資料夾
  const visualFolder = DriveApp.getFolderById(SCAN_VISUAL_FOLDER_ID);

  // 移動所有檔到視覺底圖總資料夾(避免被誤刪、且已收前綴)
  const ids = [];
  for (let i = 0; i < newFiles.length; i++) {
    const f = newFiles[i];
    f.setName(SCAN_PROCESSED_PREFIX + f.getName());
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    ids.push(f.getId());
  }

  // 讀現有 G/H/Y/Z 欄
  const existing = sh.getRange(rowNum, 7, 1, 20).getValues()[0]; // G(7) 到 Z(26)
  const existingY = String(existing[18] || '').split(',').filter(function(x){return x.trim();});
  const existingZ = String(existing[19] || '').split(',').filter(function(x){return x.trim();});

  let mainId, allIds;
  if (publishType === 'carousel') {
    // 輪播: 把新圖加到既有清單後面、最多 10 張
    allIds = existingY.concat(ids).slice(0, SCAN_MAX_CAROUSEL);
    const overflow = existingY.concat(ids).slice(SCAN_MAX_CAROUSEL);
    mainId = allIds[0];
    
    // G/H = 主圖
    sh.getRange(rowNum, 7).setValue('https://drive.google.com/file/d/' + mainId + '/view');
    sh.getRange(rowNum, 8).setValue('https://drive.google.com/thumbnail?id=' + mainId + '&sz=w400');
    // Y = 全部輪播圖
    sh.getRange(rowNum, 25).setValue(allIds.join(','));
    // Z = 溢出 + 既有備檔
    sh.getRange(rowNum, 26).setValue(existingZ.concat(overflow).join(','));
  } else {
    // 單張: 第 1 張當主圖、其他存備檔
    mainId = ids[0];
    const backup = ids.slice(1);
    
    sh.getRange(rowNum, 7).setValue('https://drive.google.com/file/d/' + mainId + '/view');
    sh.getRange(rowNum, 8).setValue('https://drive.google.com/thumbnail?id=' + mainId + '&sz=w400');
    sh.getRange(rowNum, 26).setValue(existingZ.concat(backup).join(','));
  }

  // P 排程狀態 → 「圖片就位」(若還是草稿)
  const currentStatus = sh.getRange(rowNum, 16).getValue();
  if (!currentStatus || currentStatus === '草稿' || currentStatus === '失敗') {
    sh.getRange(rowNum, 16).setValue('圖片就位');
  }
  // N 圖片審核 → 待審
  sh.getRange(rowNum, 14).setValue('待審');
  // V 備註加註
  const oldNote = sh.getRange(rowNum, 22).getValue() || '';
  const ts = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
  sh.getRange(rowNum, 22).setValue(oldNote + (oldNote ? ' | ' : '') + ts + ' 收實景照 ' + ids.length + ' 張');

  return { found: true, count: ids.length };
}

/* =========================================================
 *  觸發器安裝
 * ========================================================= */
function installScanTrigger() {
  // 移除既有
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'scanIntakeFolders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('scanIntakeFolders').timeBased().everyMinutes(5).create();
  try { SpreadsheetApp.getUi().alert('掃描器觸發器已安裝、每 5 分鐘執行一次'); } catch (e) {}
}

function uninstallScanTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'scanIntakeFolders') ScriptApp.deleteTrigger(t);
  });
  try { SpreadsheetApp.getUi().alert('掃描器觸發器已移除'); } catch (e) {}
}

/* =========================================================
 *  測試: 立即掃一次
 * ========================================================= */
function testScanOnce() {
  scanIntakeFolders();
  SpreadsheetApp.getUi().alert('掃描完成、查看 Logger 結果與 Sheet 變動');
}
