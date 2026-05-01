/**
 * 建立 79 個實景照進件資料夾
 *
 * 結構：
 *   /實景照進件_2026/
 *     /P1_001_4-19/
 *     /P1_002_4-20/
 *     ...
 *     /JUL_D03_5-11/
 *     ...
 *     /AUG_D28_7-06/
 *
 * 同時會把每個資料夾 ID 寫回 Sheet W 欄（新增「進件資料夾」欄）
 *
 * 用法：
 *   1. 貼進 Apps Script
 *   2. 執行 buildIntakeFolders()
 *   3. 等 ~2 分鐘
 *   4. 跳出 ✅ 顯示根目錄連結
 */

const INTAKE_SS_ID = '1DybgWBdCyvkEijMyaE46rKLtQD9J2ImjU8xeYCKSKnA';
const INTAKE_QUEUE = '排程佇列 Posting_Queue';
const INTAKE_ROOT_NAME = '實景照進件_2026';

function buildIntakeFolders() {
  const ui = SpreadsheetApp.getUi();
  const ans = ui.alert(
    '建立 79 個實景照進件資料夾\n\n' +
    '會在 Drive 根目錄建一個 「' + INTAKE_ROOT_NAME + '」、底下 79 個按 queue_id 命名的子資料夾。\n\n' +
    '同時把每個資料夾 ID 寫到 Sheet 第 W 欄。\n\n' +
    '預估 2 分鐘。按 OK 確認',
    ui.ButtonSet.OK_CANCEL
  );
  if (ans !== ui.Button.OK) return;

  // 找 root，不存在就建
  let root;
  const rootIter = DriveApp.getFoldersByName(INTAKE_ROOT_NAME);
  if (rootIter.hasNext()) {
    root = rootIter.next();
    Logger.log('找到既有 root: ' + root.getId());
  } else {
    root = DriveApp.createFolder(INTAKE_ROOT_NAME);
    Logger.log('建立新 root: ' + root.getId());
  }

  // 確保 W1 表頭
  const ss = SpreadsheetApp.openById(INTAKE_SS_ID);
  const sh = ss.getSheetByName(INTAKE_QUEUE);
  if (!sh.getRange('W1').getValue()) sh.getRange('W1').setValue('進件資料夾URL');

  const last = sh.getLastRow();
  const A = sh.getRange(2, 1, last - 1, 1).getValues(); // qid
  const B = sh.getRange(2, 2, last - 1, 1).getValues(); // date

  // 收集每個 queue_id 對應的列號（同一 qid 只建一個資料夾、給 1x1 列、其他 9x16 列指同一個）
  const qidMap = {};
  for (let i = 0; i < A.length; i++) {
    const qid = String(A[i][0] || '');
    let baseQid = qid;
    if (qid.endsWith('_1x1') || qid.endsWith('_9x16') || qid.endsWith('_16x9')) {
      baseQid = qid.replace(/_(1x1|9x16|16x9)$/, '');
    }
    if (!qidMap[baseQid]) qidMap[baseQid] = { rows: [], date: B[i][0] };
    qidMap[baseQid].rows.push(i + 2);
  }

  const baseKeys = Object.keys(qidMap);
  Logger.log('共 ' + baseKeys.length + ' 個 base queue_id 要建資料夾');

  // 建資料夾 + 寫 W 欄
  let ok = 0, skipped = 0, failed = 0;
  baseKeys.forEach((baseQid, i) => {
    const meta = qidMap[baseQid];
    const d = meta.date instanceof Date ? meta.date : new Date(meta.date);
    const folderName = baseQid + '_' + (d.getMonth() + 1) + '-' + String(d.getDate()).padStart(2, '0');

    try {
      let folder;
      const iter = root.getFoldersByName(folderName);
      if (iter.hasNext()) {
        folder = iter.next();
        skipped++;
      } else {
        folder = root.createFolder(folderName);
        ok++;
      }
      const folderUrl = folder.getUrl();

      // 把 URL 寫到所有對應列的 W 欄
      meta.rows.forEach(rowNum => {
        sh.getRange(rowNum, 23).setValue(folderUrl);
      });

      if (i % 10 === 0) Logger.log((i + 1) + ' / ' + baseKeys.length + ' ...');
    } catch (e) {
      failed++;
      Logger.log('FAIL: ' + baseQid + ' - ' + e.message);
    }
  });

  ui.alert(
    '✅ 完成\n\n' +
    'root: ' + root.getUrl() + '\n\n' +
    '新建 ' + ok + ' / 已存在跳過 ' + skipped + ' / 失敗 ' + failed + '\n\n' +
    '回 dashboard 看 W 欄、每列都有進件資料夾連結。'
  );
}
