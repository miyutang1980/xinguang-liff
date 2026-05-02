/**
 * 檢查 Posting_Queue 每列的圖片配置
 * 顯示：單張/輪播、主圖、輪播 N 張、備檔 N 張
 * 跑完看「執行記錄」、會列出每列詳情 + 統計
 */
function checkImageConfig() {
  const SS_ID = '1DybgWBdCyvkEijMyaE46rKLtQD9J2ImjU8xeYCKSKnA';
  const sh = SpreadsheetApp.openById(SS_ID).getSheetByName('排程佇列 Posting_Queue');
  const last = sh.getLastRow();
  if (last < 2) { Logger.log('沒資料'); return; }
  
  // 讀 A(asset_id), G(圖片URL), X(發布類型,24), Y(輪播file_ids,25), Z(備檔file_ids,26)
  const data = sh.getRange(2, 1, last - 1, 26).getValues();
  
  let single = 0, carousel = 0, noImage = 0, hasBackup = 0;
  let lines = [];
  
  data.forEach(function(r, i) {
    const row = i + 2;
    const assetId = String(r[0] || '');  // A
    const imgUrl = String(r[6] || '');   // G 主圖URL
    const type = String(r[23] || 'single'); // X
    const yIds = String(r[24] || '').split(',').filter(function(s){return s.trim();});
    const zIds = String(r[25] || '').split(',').filter(function(s){return s.trim();});
    
    const hasMain = imgUrl.length > 0;
    const isCarousel = type === 'carousel';
    
    if (!hasMain && yIds.length === 0) { noImage++; }
    else if (isCarousel) { carousel++; }
    else { single++; }
    if (zIds.length > 0) hasBackup++;
    
    lines.push(
      '列' + row + ' | ' + assetId.substring(0, 25).padEnd(28) +
      ' | 類型:' + (type || 'single').padEnd(9) +
      ' | 主圖:' + (hasMain ? '✓' : '✗') +
      ' | 輪播:' + yIds.length + '張' +
      ' | 備檔:' + zIds.length + '張'
    );
  });
  
  Logger.log('===== 每列詳情 =====');
  lines.forEach(function(line){ Logger.log(line); });
  
  Logger.log('');
  Logger.log('===== 統計總計 =====');
  Logger.log('總列數: ' + data.length);
  Logger.log('單張 (single): ' + single + ' 列');
  Logger.log('輪播 (carousel): ' + carousel + ' 列');
  Logger.log('沒圖: ' + noImage + ' 列');
  Logger.log('有備檔池的列: ' + hasBackup + ' 列');
}
