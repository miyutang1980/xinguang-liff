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
  const COL = 23;
  const headers = sh.getRange(1, 1, 1, COL).getValues()[0].map(function(h){ return String(h || ''); });
  const data = sh.getRange(2, 1, last - 1, COL).getValues();
  const out = [];
  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    var o = { _row: i + 2 };
    for (var j = 0; j < headers.length; j++) {
      var v = r[j];
      if (v === null || v === undefined || v === '') {
        o[headers[j]] = '';
      } else if (v instanceof Date) {
        o[headers[j]] = Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
      } else if (typeof v === 'number' || typeof v === 'boolean') {
        o[headers[j]] = v;
      } else {
        o[headers[j]] = String(v);
      }
    }
    out.push(o);
  }
  // 強制走純 JSON 序列化、繞過 HtmlService 對複雜物件的緩慢序列化
  return JSON.parse(JSON.stringify(out));
}

function getInteractionsData() {
  const sh = SpreadsheetApp.openById(DASH_SS_ID).getSheetByName('互動紀錄 Interactions');
  const last = sh.getLastRow();
  if (last < 2) return [];
  const headers = sh.getRange(1, 1, 1, 15).getValues()[0].map(function(h){return String(h||'');});
  const data = sh.getRange(Math.max(2, last - 99), 1, Math.min(100, last - 1), 15).getValues();
  const out = [];
  for (var i = 0; i < data.length; i++) {
    var r = data[i], o = {};
    for (var j = 0; j < headers.length; j++) {
      var v = r[j];
      o[headers[j]] = (v === null || v === undefined || v === '') ? '' :
        (v instanceof Date) ? Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM-dd HH:mm') :
        (typeof v === 'number' || typeof v === 'boolean') ? v : String(v);
    }
    out.push(o);
  }
  return JSON.parse(JSON.stringify(out));
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
  const headers = sh.getRange(1, 1, 1, 15).getValues()[0].map(function(h){return String(h||'');});
  const data = sh.getRange(2, 1, last - 1, 15).getValues();
  const out = [];
  for (var i = 0; i < data.length; i++) {
    var r = data[i], o = {};
    for (var j = 0; j < headers.length; j++) {
      var v = r[j];
      o[headers[j]] = (v === null || v === undefined || v === '') ? '' :
        (v instanceof Date) ? Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM-dd HH:mm') :
        (typeof v === 'number' || typeof v === 'boolean') ? v : String(v);
    }
    out.push(o);
  }
  return JSON.parse(JSON.stringify(out));
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

/* ========== Carousel/PublishType 操作 (Layer 3) ========== */

// 設定該列發布類型 (single / carousel)
function setPublishType(rowNum, type) {
  if (type !== 'single' && type !== 'carousel') {
    return { ok: false, error: '類型必須是 single 或 carousel' };
  }
  const sh = SpreadsheetApp.openById(DASH_SS_ID).getSheetByName('排程佇列 Posting_Queue');
  // 確保 X 欄(24)欄頭存在
  const headers = sh.getRange(1, 1, 1, Math.max(26, sh.getLastColumn())).getValues()[0];
  if (headers[23] !== '發布類型') sh.getRange(1, 24).setValue('發布類型');
  if (headers[24] !== '輪播圖file_ids') sh.getRange(1, 25).setValue('輪播圖file_ids');
  if (headers[25] !== '備檔file_ids') sh.getRange(1, 26).setValue('備檔file_ids');
  
  sh.getRange(rowNum, 24).setValue(type);
  
  // 切到 carousel 時、若 Y 欄空、自動把 G 欄主圖+Z 欄備檔合併到 Y
  if (type === 'carousel') {
    const r = sh.getRange(rowNum, 7, 1, 20).getValues()[0]; // G(7) ~ Z(26)
    const yIds = String(r[18] || '').split(',').filter(function(s){return s.trim();});
    if (yIds.length === 0) {
      const mainUrl = String(r[0] || '');
      const mainId = (mainUrl.match(/[-\w]{25,}/) || [])[0];
      const zIds = String(r[19] || '').split(',').filter(function(s){return s.trim();});
      const merged = [];
      if (mainId) merged.push(mainId);
      zIds.forEach(function(id){ if (id !== mainId) merged.push(id); });
      sh.getRange(rowNum, 25).setValue(merged.slice(0, 10).join(','));
      sh.getRange(rowNum, 26).setValue(merged.slice(10).join(','));
    }
  }
  // 切到 single 時、把 Y 欄第 1 張當主圖、其他丟回 Z
  if (type === 'single') {
    const r = sh.getRange(rowNum, 25, 1, 2).getValues()[0]; // Y, Z
    const yIds = String(r[0] || '').split(',').filter(function(s){return s.trim();});
    const zIds = String(r[1] || '').split(',').filter(function(s){return s.trim();});
    if (yIds.length > 0) {
      const newMain = yIds[0];
      sh.getRange(rowNum, 7).setValue('https://drive.google.com/file/d/' + newMain + '/view');
      sh.getRange(rowNum, 8).setValue('https://drive.google.com/thumbnail?id=' + newMain + '&sz=w400');
      // 把 Y 第 2+ 張丟回 Z
      const restToZ = yIds.slice(1).concat(zIds);
      sh.getRange(rowNum, 25).setValue('');
      sh.getRange(rowNum, 26).setValue(restToZ.join(','));
    }
  }
  
  // 重設圖片審核
  sh.getRange(rowNum, 14).setValue('待審');
  return { ok: true };
}

// 重新排序輪播圖 (傳入 file_ids 陣列、依序為輪播順序)
function reorderCarousel(rowNum, orderedIds) {
  const sh = SpreadsheetApp.openById(DASH_SS_ID).getSheetByName('排程佇列 Posting_Queue');
  if (!Array.isArray(orderedIds) || orderedIds.length === 0) {
    return { ok: false, error: '需傳入 file_ids 陣列' };
  }
  const ids = orderedIds.slice(0, 10);
  // Y 欄
  sh.getRange(rowNum, 25).setValue(ids.join(','));
  // G/H 欄主圖 = 第 1 張
  const mainId = ids[0];
  sh.getRange(rowNum, 7).setValue('https://drive.google.com/file/d/' + mainId + '/view');
  sh.getRange(rowNum, 8).setValue('https://drive.google.com/thumbnail?id=' + mainId + '&sz=w400');
  // 重設圖片審核
  sh.getRange(rowNum, 14).setValue('待審');
  return { ok: true };
}

// 取出該列所有可用圖 (Y + Z)、給後台 UI 顯示
function getAvailableImages(rowNum) {
  const sh = SpreadsheetApp.openById(DASH_SS_ID).getSheetByName('排程佇列 Posting_Queue');
  const r = sh.getRange(rowNum, 7, 1, 20).getValues()[0];
  const mainUrl = String(r[0] || '');
  const mainId = (mainUrl.match(/[-\w]{25,}/) || [])[0];
  const yIds = String(r[18] || '').split(',').filter(function(s){return s.trim();});
  const zIds = String(r[19] || '').split(',').filter(function(s){return s.trim();});
  
  return {
    mainId: mainId,
    carouselIds: yIds,
    backupIds: zIds,
    publishType: r[17] || 'single'
  };
}

// 從備檔池(Z)拉一張到主圖位置
function promoteBackupToMain(rowNum, fileId) {
  const sh = SpreadsheetApp.openById(DASH_SS_ID).getSheetByName('排程佇列 Posting_Queue');
  const r = sh.getRange(rowNum, 7, 1, 20).getValues()[0];
  const mainUrl = String(r[0] || '');
  const oldMainId = (mainUrl.match(/[-\w]{25,}/) || [])[0];
  const zIds = String(r[19] || '').split(',').filter(function(s){return s.trim();});
  
  if (zIds.indexOf(fileId) < 0) return { ok: false, error: '此 file_id 不在備檔池' };
  
  // 新主圖
  sh.getRange(rowNum, 7).setValue('https://drive.google.com/file/d/' + fileId + '/view');
  sh.getRange(rowNum, 8).setValue('https://drive.google.com/thumbnail?id=' + fileId + '&sz=w400');
  // 從 Z 移除新主圖、加回舊主圖
  const newZ = zIds.filter(function(id){return id !== fileId;});
  if (oldMainId && oldMainId !== fileId) newZ.push(oldMainId);
  sh.getRange(rowNum, 26).setValue(newZ.join(','));
  
  sh.getRange(rowNum, 14).setValue('待審');
  return { ok: true };
}
