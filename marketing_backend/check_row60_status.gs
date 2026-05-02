/**
 * 跑前先確認列 60 的圖文 + 發布類型
 */
function checkRow60Status() {
  const sh = SpreadsheetApp.openById('1DybgWBdCyvkEijMyaE46rKLtQD9J2ImjU8xeYCKSKnA').getSheetByName('排程佇列 Posting_Queue');
  const r = sh.getRange(60, 1, 1, 26).getValues()[0];
  
  Logger.log('===== 列 60 狀態 =====');
  Logger.log('A queue_id: ' + r[0]);
  Logger.log('B 排程日期: ' + r[1]);
  Logger.log('C 排程時間: ' + r[2]);
  Logger.log('D 平台: ' + r[3]);
  Logger.log('E 比例: ' + r[4]);
  Logger.log('G 圖片URL: ' + r[6]);
  Logger.log('H 縮圖URL: ' + r[7]);
  Logger.log('I 主題: ' + r[8]);
  Logger.log('J 主標(' + String(r[9]||'').length + '字): ' + String(r[9]||'').substring(0, 60));
  Logger.log('K 內文(' + String(r[10]||'').length + '字): ' + String(r[10]||'').substring(0, 80));
  Logger.log('L Hashtags: ' + r[11]);
  Logger.log('N 圖片審核: ' + r[13]);
  Logger.log('O 文案審核: ' + r[14]);
  Logger.log('P 排程狀態: ' + r[15]);
  Logger.log('R post_id: ' + r[17]);
  Logger.log('S post_url: ' + r[18]);
  Logger.log('X 發布類型: ' + r[23]);
  Logger.log('Y 輪播file_ids: ' + r[24]);
  Logger.log('Z 備檔file_ids: ' + r[25]);
  
  Logger.log('');
  Logger.log('===== 發文前檢查 =====');
  const checks = [];
  if (!r[6]) checks.push('❌ G 欄圖片URL 為空');
  if (!r[10]) checks.push('❌ K 欄文案內文 為空');
  if (r[15] === '已發布') checks.push('⚠️ 已經是「已發布」、再發會重複');
  if (r[15] === '失敗') checks.push('⚠️ 上次發失敗、發前可能要清 P/T 欄');
  if (r[23] === 'carousel') {
    const yIds = String(r[24]||'').split(',').filter(function(s){return s.trim();});
    if (yIds.length < 2) checks.push('❌ X=carousel 但 Y 欄少於 2 張');
    else checks.push('✓ carousel 模式、輪播 ' + yIds.length + ' 張');
  } else {
    checks.push('✓ single 模式');
  }
  if (checks.length === 1 && checks[0].indexOf('✓') === 0) {
    Logger.log('✓ 通過、可以發');
  }
  checks.forEach(function(c){ Logger.log(c); });
}
