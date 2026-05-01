/**
 * 診斷：列出 Posting_Queue 哪些列 G 欄（圖片URL）為空
 *
 * 執行 → 看執行記錄輸出，會分組列出每個 Phase 缺圖列數
 */

const DIAG_SS_ID = '1DybgWBdCyvkEijMyaE46rKLtQD9J2ImjU8xeYCKSKnA';
const DIAG_QUEUE = '排程佇列 Posting_Queue';

function diagnoseMissingImages() {
  const ss = SpreadsheetApp.openById(DIAG_SS_ID);
  const sh = ss.getSheetByName(DIAG_QUEUE);
  const last = sh.getLastRow();
  if (last < 2) { Logger.log('Posting_Queue 是空的'); return; }

  // A=qid, B=date, E=ratio, G=img_url, V=stage
  const data = sh.getRange(2, 1, last - 1, 22).getValues();

  const missing = [];
  const phases = { p1: 0, p2: 0, p3: 0, other: 0 };
  data.forEach((row, idx) => {
    const qid = String(row[0] || '');
    const dateStr = row[1] instanceof Date ? Utilities.formatDate(row[1], 'GMT+8', 'yyyy-MM-dd') : String(row[1]);
    const ratio = row[4];
    const img = String(row[6] || '');
    const stage = String(row[21] || '');
    if (!img) {
      missing.push(`列 ${idx + 2} | ${dateStr} | ${qid} | ${ratio} | ${stage}`);
      if (qid.startsWith('P1_')) phases.p1++;
      else if (qid.startsWith('JUL_')) phases.p2++;
      else if (qid.startsWith('AUG_')) phases.p3++;
      else phases.other++;
    }
  });

  Logger.log(`=== 診斷結果 ===`);
  Logger.log(`總列數: ${data.length}`);
  Logger.log(`G 欄為空: ${missing.length} 列`);
  Logger.log(`  P1 暖身缺圖: ${phases.p1}`);
  Logger.log(`  P2 七月預告缺圖: ${phases.p2}`);
  Logger.log(`  P3 八月預告缺圖: ${phases.p3}`);
  Logger.log(`  其他: ${phases.other}`);
  Logger.log(`---`);
  missing.slice(0, 50).forEach(m => Logger.log(m));
  if (missing.length > 50) Logger.log(`... 還有 ${missing.length - 50} 列`);

  SpreadsheetApp.getUi().alert(
    `診斷完成\n\n總 ${data.length} 列\nG 欄空白 ${missing.length} 列\n\n` +
    `P1 暖身: ${phases.p1}\nP2 七月: ${phases.p2}\nP3 八月: ${phases.p3}\n\n` +
    `詳細看執行記錄`
  );
}
