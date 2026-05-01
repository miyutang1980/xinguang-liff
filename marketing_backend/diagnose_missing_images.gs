/**
 * 診斷：列出 Posting_Queue 哪些列 G 欄（圖片URL）為空
 * 極簡版：只讀 G 欄、只 log、不 alert（避免逾時）
 */

const DIAG_SS_ID = '1DybgWBdCyvkEijMyaE46rKLtQD9J2ImjU8xeYCKSKnA';
const DIAG_QUEUE = '排程佇列 Posting_Queue';

function diagnoseMissingImages() {
  const t0 = Date.now();
  const ss = SpreadsheetApp.openById(DIAG_SS_ID);
  const sh = ss.getSheetByName(DIAG_QUEUE);
  const last = sh.getLastRow();
  Logger.log('開啟試算表用時 ' + (Date.now() - t0) + 'ms, last row=' + last);

  if (last < 2) { Logger.log('Posting_Queue 是空的'); return; }

  // 只讀 A B G 三欄（最快）
  const t1 = Date.now();
  const A = sh.getRange(2, 1, last - 1, 1).getValues(); // qid
  const B = sh.getRange(2, 2, last - 1, 1).getValues(); // date
  const G = sh.getRange(2, 7, last - 1, 1).getValues(); // img_url
  Logger.log('讀 A/B/G 三欄用時 ' + (Date.now() - t1) + 'ms, 共 ' + A.length + ' 列');

  let p1 = 0, p2 = 0, p3 = 0, other = 0, total = 0;
  for (let i = 0; i < A.length; i++) {
    const qid = String(A[i][0] || '');
    const img = String(G[i][0] || '');
    if (!img) {
      total++;
      if (qid.startsWith('P1_')) p1++;
      else if (qid.startsWith('JUL_')) p2++;
      else if (qid.startsWith('AUG_')) p3++;
      else other++;
    }
  }

  Logger.log('===== 診斷結果 =====');
  Logger.log('總列數 ' + A.length);
  Logger.log('G 欄空白 ' + total + ' 列');
  Logger.log('  P1 暖身 ' + p1);
  Logger.log('  P2 七月 ' + p2);
  Logger.log('  P3 八月 ' + p3);
  Logger.log('  其他 ' + other);
  Logger.log('總用時 ' + (Date.now() - t0) + 'ms');
}
