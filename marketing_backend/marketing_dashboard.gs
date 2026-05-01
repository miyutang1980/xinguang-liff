/**
 * 太平新光社群行銷後台 Dashboard（Apps Script Web App 版）
 *
 * 部署：Apps Script → 部署 → Web App → 任何人
 * 取得 URL → 嵌入 admin.taipingxinguang.org/marketing 或直接做書籤
 *
 * 功能分頁：
 *   1. 待審佇列（圖文一鍵審核 / 退回）
 *   2. 排程看板（七天行事曆視圖）
 *   3. 互動中心（最近留言、未回覆）
 *   4. 數據儀表（KPI + 觸及趨勢）
 *   5. 規則編輯（Auto_Reply）
 *   6. 預約管理
 */

const DASH_SS_ID = '1DybgWBdCyvkEijMyaE46rKLtQD9J2ImjU8xeYCKSKnA';

function doGet(e) {
  const tpl = HtmlService.createTemplateFromFile('dashboard');
  return tpl.evaluate()
    .setTitle('太平新光社群行銷後台')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  // LIFF 預約寫入端點
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'writeBooking') return writeBooking_(body);
    if (body.action === 'approveRow')   return approveRowApi_(body);
    if (body.action === 'reject')       return rejectRowApi_(body);
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'unknown action' })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  }
}

/* ========== Server functions for client to call ========== */

function getQueueData() {
  const sh = SpreadsheetApp.openById(DASH_SS_ID).getSheetByName('排程佇列 Posting_Queue');
  const last = sh.getLastRow();
  if (last < 2) return [];
  const lastCol = sh.getLastColumn(); // 自動讀到 W 欄(進件資料夾URL)、未來新增欄也不需再改
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = sh.getRange(2, 1, last - 1, lastCol).getValues();
  return data.map((r, i) => {
    const o = { _row: i + 2 };
    headers.forEach((h, j) => {
      o[h] = (r[j] instanceof Date) ? Utilities.formatDate(r[j], 'Asia/Taipei', 'yyyy-MM-dd HH:mm') : r[j];
    });
    return o;
  });
}

function getInteractionsData() {
  const sh = SpreadsheetApp.openById(DASH_SS_ID).getSheetByName('互動紀錄 Interactions');
  const last = sh.getLastRow();
  if (last < 2) return [];
  const headers = sh.getRange(1, 1, 1, 15).getValues()[0];
  const data = sh.getRange(Math.max(2, last - 99), 1, Math.min(100, last - 1), 15).getValues();
  return data.map(r => {
    const o = {};
    headers.forEach((h, j) => {
      o[h] = (r[j] instanceof Date) ? Utilities.formatDate(r[j], 'Asia/Taipei', 'yyyy-MM-dd HH:mm') : r[j];
    });
    return o;
  });
}

function getInsightsKPIs() {
  const sh = SpreadsheetApp.openById(DASH_SS_ID).getSheetByName('Insights');
  const last = sh.getLastRow();
  if (last < 2) return { total: 0, reach7d: 0, eng7d: 0, posts: 0 };
  const data = sh.getRange(2, 1, last - 1, 17).getValues();
  const sevenDaysAgo = new Date(); sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);

  let reach = 0, eng = 0, postSet = new Set();
  for (const r of data) {
    const d = (r[1] instanceof Date) ? r[1] : new Date(r[1]);
    if (d < sevenDaysAgo) continue;
    reach += Number(r[6] || 0);
    eng += Number(r[7] || 0) + Number(r[8] || 0) + Number(r[9] || 0) + Number(r[10] || 0);
    postSet.add(r[3]);
  }
  return { total: data.length, reach7d: reach, eng7d: eng, posts: postSet.size };
}

function getBookingsData() {
  const sh = SpreadsheetApp.openById(DASH_SS_ID).getSheetByName('預約 Bookings');
  const last = sh.getLastRow();
  if (last < 2) return [];
  const headers = sh.getRange(1, 1, 1, 15).getValues()[0];
  const data = sh.getRange(2, 1, last - 1, 15).getValues();
  return data.map(r => {
    const o = {};
    headers.forEach((h, j) => {
      o[h] = (r[j] instanceof Date) ? Utilities.formatDate(r[j], 'Asia/Taipei', 'yyyy-MM-dd HH:mm') : r[j];
    });
    return o;
  });
}

function approveRow(rowNum, target) {
  const sh = SpreadsheetApp.openById(DASH_SS_ID).getSheetByName('排程佇列 Posting_Queue');
  if (target === 'image') sh.getRange(rowNum, 14).setValue('過');
  if (target === 'copy')  sh.getRange(rowNum, 15).setValue('過');
  // 如果雙審過 → 自動設成已排程
  const r = sh.getRange(rowNum, 14, 1, 2).getValues()[0];
  if (r[0] === '過' && r[1] === '過') {
    if (sh.getRange(rowNum, 16).getValue() === '草稿') {
      sh.getRange(rowNum, 16).setValue('已排程');
    }
  }
  return { ok: true };
}

function rejectRow(rowNum, target, reason) {
  const sh = SpreadsheetApp.openById(DASH_SS_ID).getSheetByName('排程佇列 Posting_Queue');
  if (target === 'image') sh.getRange(rowNum, 14).setValue('退回');
  if (target === 'copy')  sh.getRange(rowNum, 15).setValue('退回');
  sh.getRange(rowNum, 22).setValue('退回原因：' + reason);
  return { ok: true };
}

function updateCopyText(rowNum, headline, body, hashtags) {
  const sh = SpreadsheetApp.openById(DASH_SS_ID).getSheetByName('排程佇列 Posting_Queue');
  if (headline !== undefined) sh.getRange(rowNum, 10).setValue(headline);
  if (body !== undefined) sh.getRange(rowNum, 11).setValue(body);
  if (hashtags !== undefined) sh.getRange(rowNum, 12).setValue(hashtags);
  // 改文後重設文案審核為待審
  sh.getRange(rowNum, 15).setValue('待審');
  return { ok: true };
}

function updateAutoReplyRule(ruleId, field, value) {
  const sh = SpreadsheetApp.openById(DASH_SS_ID).getSheetByName('自動回覆 Auto_Reply');
  const last = sh.getLastRow();
  const data = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === ruleId) {
      const headers = sh.getRange(1, 1, 1, 14).getValues()[0];
      const colIdx = headers.indexOf(field);
      if (colIdx >= 0) {
        sh.getRange(i + 2, colIdx + 1).setValue(value);
        return { ok: true };
      }
    }
  }
  return { ok: false, error: '找不到 ruleId/field' };
}

/* ========== LIFF Booking 端點 ========== */
function writeBooking_(body) {
  const sh = SpreadsheetApp.openById(DASH_SS_ID).getSheetByName('預約 Bookings');
  const id = 'B' + Date.now();
  const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
  sh.appendRow([
    id, now, 'LINE LIFF',
    body.parentName || '', body.kidName || '', body.kidGrade || '',
    body.phone || '', body.lineUserId || '',
    body.preferDate || '', body.preferTime || '',
    body.note || '', '新進', '', '', ''
  ]);
  return ContentService.createTextOutput(JSON.stringify({ ok: true, id: id }))
    .setMimeType(ContentService.MimeType.JSON);
}

function approveRowApi_(body) {
  approveRow(body.rowNum, body.target);
  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}

function rejectRowApi_(body) {
  rejectRow(body.rowNum, body.target, body.reason || '');
  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}
