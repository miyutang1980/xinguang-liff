/**
 * F 步驟測試：publishOneRow(60) 真發 AUG_D01_1x1
 *
 * 列 60 = Posting_Queue 第 60 列 = 6/9 12:30 八月預告 D01 拉力與張力 1x1
 *
 * 注意：
 *   - 真的會發到 IG @eagle__xinguang 和 FB 粉專
 *   - 發完馬上去兩處手動刪除（不刪 8/1 觸發器會再發一次）
 *   - 看執行記錄看 IG/FB API 回應
 */

function testPublishRow60() {
  // 把列 60 排程狀態改成「等待發布」、然後叫 publishOneRow(60)
  const ss = SpreadsheetApp.openById('1DybgWBdCyvkEijMyaE46rKLtQD9J2ImjU8xeYCKSKnA');
  const sh = ss.getSheetByName('排程佇列 Posting_Queue');

  // 先檢查列 60 圖文都齊
  const row = sh.getRange(60, 1, 1, 22).getValues()[0];
  Logger.log('=== 列 60 內容檢查 ===');
  Logger.log('A queue_id: ' + row[0]);
  Logger.log('B 排程日期: ' + row[1]);
  Logger.log('D 平台: ' + row[3]);
  Logger.log('E 比例: ' + row[4]);
  Logger.log('G 圖片URL: ' + row[6]);
  Logger.log('I 主題: ' + row[8]);
  Logger.log('J 主標長度: ' + String(row[9] || '').length);
  Logger.log('K 內文長度: ' + String(row[10] || '').length);
  Logger.log('P 排程狀態: ' + row[15]);

  if (!row[6]) { Logger.log('❌ G 欄圖片URL為空、無法發文'); return; }
  if (!row[10]) { Logger.log('❌ K 欄文案內文為空、無法發文'); return; }

  // 直接叫 publishOneRow（你 posting_engine.gs 裡的函式）
  Logger.log('=== 開始呼叫 publishOneRow(60) ===');
  publishOneRow(60);
  Logger.log('=== publishOneRow 完成、檢查 Sheet 狀態 ===');

  const after = sh.getRange(60, 16, 1, 5).getValues()[0];
  Logger.log('P 排程狀態: ' + after[0]);
  Logger.log('Q 發文時間: ' + after[1]);
  Logger.log('R post_id: ' + after[2]);
  Logger.log('S post_url: ' + after[3]);
  Logger.log('T 錯誤訊息: ' + after[4]);
}
