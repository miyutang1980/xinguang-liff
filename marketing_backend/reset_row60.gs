/**
 * 重置列 60（JUL_D31_1x1, 6/8 12:30）狀態
 *
 * 因為 testPublishRow60 真發測試後、列 60 的 P/Q/R/S 已被填成「已發布」、
 * 觸發器 6/8 不會再發、所以這裡清掉、讓 6/8 觸發器照排程重發
 */

function resetRow60() {
  const ss = SpreadsheetApp.openById('1DybgWBdCyvkEijMyaE46rKLtQD9J2ImjU8xeYCKSKnA');
  const sh = ss.getSheetByName('排程佇列 Posting_Queue');

  // 先看現在狀態
  const before = sh.getRange(60, 1, 1, 22).getValues()[0];
  Logger.log('===== 重置前 =====');
  Logger.log('A queue_id: ' + before[0]);
  Logger.log('I 主題: ' + before[8]);
  Logger.log('P 排程狀態: ' + before[15]);
  Logger.log('Q 發文時間: ' + before[16]);
  Logger.log('R post_id: ' + before[17]);
  Logger.log('S post_url: ' + before[18]);
  Logger.log('T 錯誤訊息: ' + before[19]);

  // 直接重置、不跳確認框（從 Apps Script 編輯器跳不出 UI）
  // P (col 16) 草稿、Q (17) R (18) S (19) T (20) 清空
  sh.getRange(60, 16).setValue('草稿');
  sh.getRange(60, 17).clearContent();
  sh.getRange(60, 18).clearContent();
  sh.getRange(60, 19).clearContent();
  sh.getRange(60, 20).clearContent();
  // 同時把 N/O 文圖審核設回「待審」（避免 6/8 觸發器以為雙審已過、反而跳過；你可以之後再到後台點「圖過」「文過」）
  // 這裡保留原審核狀態、不動

  const after = sh.getRange(60, 1, 1, 22).getValues()[0];
  Logger.log('');
  Logger.log('===== 重置後 =====');
  Logger.log('P 排程狀態: ' + after[15]);
  Logger.log('Q 發文時間: ' + (after[16] || '(空)'));
  Logger.log('R post_id: ' + (after[17] || '(空)'));
  Logger.log('S post_url: ' + (after[18] || '(空)'));
  Logger.log('T 錯誤訊息: ' + (after[19] || '(空)'));
  Logger.log('');
  Logger.log('✅ 列 60 已重置、需手動重點「圖過/文過」才會重新變「已排程」。');
}
