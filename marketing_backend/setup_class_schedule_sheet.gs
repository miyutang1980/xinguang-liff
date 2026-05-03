/**
 * 一鍵建立「Class_Schedule 班別課表」試算表分頁
 * 5/3 學期調整版、永久 source of truth
 *
 * 用法：
 * 1. 在 Apps Script 編輯器、選 setupClassScheduleSheet、執行
 * 2. 自動建立 / 更新分頁、寫入完整課表
 */
function setupClassScheduleSheet() {
  const SS_ID = '1DybgWBdCyvkEijMyaE46rKLtQD9J2ImjU8xeYCKSKnA';
  const ss = SpreadsheetApp.openById(SS_ID);
  const SHEET_NAME = '班別課表 Class_Schedule';

  let sh = ss.getSheetByName(SHEET_NAME);
  if (sh) {
    // 已存在 → 清空重寫
    sh.clear();
  } else {
    sh = ss.insertSheet(SHEET_NAME);
  }

  // 標題列
  const headers = ['時段', '時間', '班別', '上課日', '備註'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // 完整資料（依據 5/3 PDF 公告）
  const rows = [
    // 時段一
    ['時段一', '14:00-16:00', 'SA',  '週一、三、四、五', '國小低年級／銜接班'],
    ['時段一', '14:00-16:00', '1A',  '週一、三、四、五', '國小低年級／銜接班'],
    ['時段一', '13:30-16:15', 'G1',  '週一、三、四、五', '國小低年級／銜接班'],
    // 時段二（中高年級 3 天班、G2 浸潤）
    ['時段二', '16:30-19:00', '1A',  '週二、三、四', '中高年級 3 天班'],
    ['時段二', '16:30-19:00', '3A',  '週一、三、五', '週三改 R102'],
    ['時段二', '16:30-19:00', '1C',  '週一、二、四', '中高年級 3 天班'],
    ['時段二', '16:30-19:00', '4A',  '週二、四、五', '週四改 R102'],
    ['時段二', '16:30-18:45', 'G2',  '週一至週五', '浸潤班'],
    ['時段二', '16:30-19:00', '2B',  '週一、二、四', '中高年級 3 天班'],
    ['時段二', '16:30-19:00', '3C',  '週一、三、五', '週一改 R102'],
    // 時段三（全民英檢／幼兒）
    ['時段三', '19:00-21:00', '4B 全民英檢中級初試', '週一、四', '全民英檢備考'],
    ['時段三', '19:00-21:00', '全民英檢中級複試',     '週二、五', '全民英檢備考'],
    ['時段三', '19:00-21:00', '3B 全民英檢初級初試', '週二、五', '全民英檢備考'],
    ['時段三', '19:00-21:00', '全民英檢初級複試',     '週一、三', '全民英檢備考'],
    ['時段三', '19:00-20:30', 'ES 幼兒美語',          '週一、二、五', '幼兒美語']
  ];

  sh.getRange(2, 1, rows.length, headers.length).setValues(rows);

  // 樣式
  sh.setColumnWidth(1, 80);
  sh.setColumnWidth(2, 120);
  sh.setColumnWidth(3, 220);
  sh.setColumnWidth(4, 180);
  sh.setColumnWidth(5, 180);

  // 時段分組底色
  for (let i = 0; i < rows.length; i++) {
    const r = i + 2;
    const seg = rows[i][0];
    let bg = '#ffffff';
    if (seg === '時段一') bg = '#eaf2f8';
    else if (seg === '時段二') bg = '#fdf2e3';
    else if (seg === '時段三') bg = '#eafaf1';
    sh.getRange(r, 1, 1, headers.length).setBackground(bg);
  }
  sh.setFrozenRows(1);
  sh.getRange(1, 1, rows.length + 1, headers.length).setBorder(true, true, true, true, true, true);

  // 在最下方加一行學校資訊
  const infoRow = rows.length + 3;
  sh.getRange(infoRow, 1).setValue('校址');
  sh.getRange(infoRow, 2, 1, 4).merge().setValue('台中市太平區新福路 880 號');
  sh.getRange(infoRow + 1, 1).setValue('電話');
  sh.getRange(infoRow + 1, 2, 1, 4).merge().setValue('04-2396-0585');
  sh.getRange(infoRow + 2, 1).setValue('LINE OA');
  sh.getRange(infoRow + 2, 2, 1, 4).merge().setValue('@143qbory');
  sh.getRange(infoRow + 3, 1).setValue('學期重點');
  sh.getRange(infoRow + 3, 2, 1, 4).merge().setValue('時段二中高年級為 3 天班（G2 為浸潤班）、週時數 7.5 小時不變、同教材同進度同師資');
  sh.getRange(infoRow, 1, 4, 1).setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange(infoRow, 2, 4, 4).setBackground('#fafafa');
  sh.getRange(infoRow, 1, 4, 5).setBorder(true, true, true, true, true, true);

  SpreadsheetApp.getUi().alert('✅ 班別課表 已建立 / 更新完成、共 ' + rows.length + ' 筆班別資料');
}
