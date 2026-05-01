/**
 * 重置列 60（JUL_D31_1x1, 6/8 12:30）狀態
 *
 * 因為 testPublishRow60 真發測試後、列 60 的 P/Q/R/S 已被填成「已發布」、
 * 觸發器 6/8 不會再發、所以這裡清掉、讓 6/8 觸發器照排程重發
 */

function resetRow60() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById('1DybgWBdCyvkEijMyaE46rKLtQD9J2ImjU8xeYCKSKnA');
  const sh = ss.getSheetByName('排程佇列 Posting_Queue');

  // 先看看現在狀態
  const before = sh.getRange(60, 1, 1, 22).getValues()[0];
  const ans = ui.alert(
    '重置列 60 確認\n\n' +
    'A queue_id: ' + before[0] + '\n' +
    'I 主題: ' + before[8] + '\n' +
    'P 排程狀態: ' + before[15] + '\n' +
    'R post_id: ' + before[17] + '\n\n' +
    '會清空 P/Q/R/S/T 五欄、把 P 改回「草稿」。\n' +
    '6/8 12:30 觸發器會重新發這篇。\n\n' +
    '按 OK 確認',
    ui.ButtonSet.OK_CANCEL
  );
  if (ans !== ui.Button.OK) return;

  // 清 P (col 16) Q (17) R (18) S (19) T (20)
  sh.getRange(60, 16).setValue('草稿');
  sh.getRange(60, 17).clearContent();
  sh.getRange(60, 18).clearContent();
  sh.getRange(60, 19).clearContent();
  sh.getRange(60, 20).clearContent();

  ui.alert('✅ 列 60 已重置\n\nP=草稿，Q/R/S/T 已清空。\n6/8 觸發器會重新發 JUL_D31 這篇。');
}
