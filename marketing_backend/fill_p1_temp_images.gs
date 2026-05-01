/**
 * Phase 1 暖身 22 列暫頂補圖
 *
 * 策略：用七月 D03-D24 的 1x1 file_id 暫時填入 P1 22 列的 G/H 欄
 *      圖內容雖然是七月課程主題、但以「暖身期視覺佔位」為目的可接受
 *      之後拍完實體照、改用 fillP1RealImages(map) 覆蓋
 *
 * 用法：
 *   1. 貼進 Apps Script
 *   2. 執行 fillP1TempImages()
 *   3. 看 alert 確認 22 列 G/H 欄已填
 */

const P1FILL_SS_ID = '1DybgWBdCyvkEijMyaE46rKLtQD9J2ImjU8xeYCKSKnA';
const P1FILL_QUEUE = '排程佇列 Posting_Queue';

// 七月 D03-D24 的 1x1 file_id（按順序對應 P1_001 ~ P1_022）
const P1_TEMP_IDS = [
  '1HDukFK2QUZ3XYobC5muOBFGY2MNbLhYI', // d03_1x1 → P1_001
  '1AQQW7BQP9TuK10WDWkMW87Et_loxX8Ls', // d04 → P1_002
  '1D6okM5KXpR49NUnbeOGkHvkHsyxitvJv', // d05 → P1_003
  '1x0NtyUebYTS18czyaXWQuHzmnz-KYa1X', // d06 → P1_004
  '1pk7kNOv0Fdsh1Oj_jY0yPr_D8HssSS1s', // d07 → P1_005
  '1RzSGb8QVDIHD7eBxox3_WHUhJTUvl1Ed', // d08 → P1_006
  '1qDiEwUxXbHi0nH8PDBp-nBqoMlbEZCXb', // d09 → P1_007
  '1qGYIXHoHnod-oB54JC5T0IP_2OJzJVI7', // d10 → P1_008
  '13jLIiIb09IwpfkYaZM00CXJXSmXCGcnZ', // d11 → P1_009
  '14Wg5c9gBrJKahwUUG2Ttp_wcaOf_HJj_', // d12 → P1_010
  '1l6dmpk6KKphMqio0nilSv4NaxHAHMXfh', // d13 → P1_011
  '10Uyg1VntuMjEYNAPQR5Nc-Jtt47dBU2g', // d14 → P1_012
  '1G-vMDZpEKChcYGPJiZ9OOXGjtofl06qy', // d15 → P1_013
  '1f08hSoYDheL6znXeYZ0JXdF-znHL76o_', // d16 → P1_014
  '1f5r3W6hu7oIOLYZ8OMYSZJtDEkcvde9l', // d17 → P1_015
  '1B-_TLKyacHyRuTWXHR9-3TfWhxdHe0i_', // d18 → P1_016
  '1lr1vKREyN_8QqR_8GFYhEcFjZbIsKh5F', // d19 → P1_017
  '1qEiDK150vA4aSuJKgbJ55XMyVwvCPUzc', // d20 → P1_018
  '1tw5iwdxtwSLxEMQAT2tlr8ADeOz34nFq', // d21 → P1_019
  '1Yz6wUR8L3GeG1oKdPLpRvUotshWuqAXK', // d22 → P1_020
  '1xxojHGZXMBK8bVGWUCpH4TD2sIs6I7GF', // d23 → P1_021
  '1v7bWYaOg9eJf-W0avHaBx0LA4Inm6iz5', // d24 → P1_022
];

function fillP1TempImages() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(P1FILL_SS_ID);
  const sh = ss.getSheetByName(P1FILL_QUEUE);
  const last = sh.getLastRow();

  // 找出 P1_001 ~ P1_022 的列號（A 欄）
  const A = sh.getRange(2, 1, last - 1, 1).getValues();
  const p1Rows = [];
  for (let i = 0; i < A.length; i++) {
    const qid = String(A[i][0] || '');
    if (qid.startsWith('P1_')) p1Rows.push(i + 2); // sheet row number
  }

  if (p1Rows.length !== 22) {
    ui.alert('警告：P1 列數不對（找到 ' + p1Rows.length + '，預期 22）');
    return;
  }

  const ans = ui.alert(
    '⚠️ Phase 1 暫頂補圖確認\n\n' +
    '將為 P1_001 ~ P1_022 共 22 列填入七月 D03-D24 的 1x1 圖。\n\n' +
    '這是視覺佔位、不是最終圖。\n' +
    '等你拍實體照後、再用 fillP1RealImages 覆蓋。\n\n' +
    '按 OK 確認',
    ui.ButtonSet.OK_CANCEL
  );
  if (ans !== ui.Button.OK) return;

  // 一列一列寫 G/H
  for (let i = 0; i < 22; i++) {
    const fid = P1_TEMP_IDS[i];
    const rowNum = p1Rows[i];
    const imgUrl = 'https://drive.google.com/file/d/' + fid + '/view?usp=drivesdk';
    const thumbUrl = 'https://drive.google.com/thumbnail?id=' + fid + '&sz=w400';
    sh.getRange(rowNum, 7).setValue(imgUrl);
    sh.getRange(rowNum, 8).setValue(thumbUrl);
  }

  ui.alert('✅ 完成\n\nP1 22 列暫頂圖已填入。\n之後拍完實體照、跑 fillP1RealImages 覆蓋。');
}


/**
 * 拍完實體照後用這個覆蓋 P1 22 列
 *
 * 用法：
 *   把 P1_REAL_IDS 改成你拍完上傳到 Drive 後的 22 個 file_id
 *   執行 fillP1RealImages()
 */
const P1_REAL_IDS = [
  // P1_001 4/19 開站公告
  // P1_002 4/20 七月梯介紹
  // ...22 個 file_id 等你提供
];

function fillP1RealImages() {
  const ui = SpreadsheetApp.getUi();
  if (P1_REAL_IDS.length !== 22) {
    ui.alert('請先把 P1_REAL_IDS 填滿 22 個 file_id 再執行');
    return;
  }
  const ss = SpreadsheetApp.openById(P1FILL_SS_ID);
  const sh = ss.getSheetByName(P1FILL_QUEUE);
  const last = sh.getLastRow();
  const A = sh.getRange(2, 1, last - 1, 1).getValues();
  const p1Rows = [];
  for (let i = 0; i < A.length; i++) {
    if (String(A[i][0] || '').startsWith('P1_')) p1Rows.push(i + 2);
  }
  for (let i = 0; i < 22; i++) {
    const fid = P1_REAL_IDS[i];
    sh.getRange(p1Rows[i], 7).setValue('https://drive.google.com/file/d/' + fid + '/view?usp=drivesdk');
    sh.getRange(p1Rows[i], 8).setValue('https://drive.google.com/thumbnail?id=' + fid + '&sz=w400');
  }
  ui.alert('✅ P1 22 列已換成實體照');
}
