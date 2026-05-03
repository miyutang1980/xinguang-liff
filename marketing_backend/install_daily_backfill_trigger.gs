/**
 * 一鍵安裝每天自動 backfill trigger
 * - 每天凌晨 4:00 自動跑 backfillAllHistorical（IG + FB 新貼文 + 互動補抓）
 * - 先清掉舊的同名 trigger、避免重複
 */
function installDailyBackfillTrigger() {
  // 清掉舊的
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  triggers.forEach(function(t){
    if (t.getHandlerFunction() === 'dailyBackfillJob') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  
  // 裝新的：每天 4:00 跑
  ScriptApp.newTrigger('dailyBackfillJob')
    .timeBased()
    .everyDays(1)
    .atHour(4)
    .create();
  
  Logger.log('===== Trigger 安裝完成 =====');
  Logger.log('清掉舊 trigger: ' + removed + ' 個');
  Logger.log('新 trigger: 每天凌晨 4:00 自動跑 dailyBackfillJob');
  Logger.log('（會自動補抓昨天的 IG + FB 新貼文互動）');
}

/**
 * 每日自動跑的工作：IG + FB backfill + 互動補抓
 * 寫成 try/catch、單一函式失敗不影響其他
 */
function dailyBackfillJob() {
  const startTime = new Date();
  Logger.log('===== Daily Backfill Job 開始 =====');
  Logger.log('開始時間: ' + Utilities.formatDate(startTime, 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss'));
  
  // 1. IG 歷史（含新貼文）
  try {
    Logger.log('[1/3] backfillIGHistorical 開始...');
    backfillIGHistorical();
    Logger.log('[1/3] backfillIGHistorical 完成');
  } catch (e) {
    Logger.log('[1/3] IG backfill 失敗: ' + e.message);
  }
  
  // 2. FB 歷史（含新貼文）
  try {
    Logger.log('[2/3] backfillFBHistorical 開始...');
    backfillFBHistorical();
    Logger.log('[2/3] backfillFBHistorical 完成');
  } catch (e) {
    Logger.log('[2/3] FB backfill 失敗: ' + e.message);
  }
  
  // 3. FB 互動補抓（針對 likes+comments+shares=0 的舊貼文）
  try {
    Logger.log('[3/3] fbBackfillInteractions 開始...');
    fbBackfillInteractions();
    Logger.log('[3/3] fbBackfillInteractions 完成');
  } catch (e) {
    Logger.log('[3/3] FB 互動補抓失敗: ' + e.message);
  }
  
  const endTime = new Date();
  const elapsed = Math.round((endTime - startTime) / 1000);
  Logger.log('===== Daily Backfill Job 結束 =====');
  Logger.log('耗時: ' + elapsed + ' 秒');
}

/**
 * 列出所有 trigger（檢查用）
 */
function listAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  Logger.log('===== 全部 Trigger =====');
  Logger.log('總數: ' + triggers.length);
  triggers.forEach(function(t, i){
    Logger.log((i+1) + '. ' + t.getHandlerFunction() + ' | ' + t.getEventType() + ' | ID: ' + t.getUniqueId());
  });
}
