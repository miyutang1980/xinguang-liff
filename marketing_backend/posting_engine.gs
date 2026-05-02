/**
 * 太平新光社群發文引擎
 *
 * 觸發器：每 10 分鐘檢查一次 Posting_Queue，找到「圖片審核=過 ∧ 文案審核=過 ∧ 排程狀態=已排程 ∧ 排程時間≤現在」的列就發
 *
 * 核心 functions：
 *   processPostingQueue()    — 主排程器（每 10 分鐘 trigger 一次）
 *   approveAllImages()       — 一鍵把 Queue 全部圖片審過（測試用）
 *   approveAllCopies()       — 一鍵把 Queue 全部文案審過
 *   scheduleAllApproved()    — 把雙審過的列從「草稿」設為「已排程」
 *   publishOneRow(rowNum)    — 手動發某一列（測試用）
 *   installTriggers()        — 安裝時觸發器
 *   uninstallTriggers()      — 移除觸發器
 */

const PE_SS_ID = '1DybgWBdCyvkEijMyaE46rKLtQD9J2ImjU8xeYCKSKnA';
const PE_QUEUE_NAME = '排程佇列 Posting_Queue';
const PE_TZ = 'Asia/Taipei';

// 從「設定 Settings」讀 token
function pe_getSetting_(key) {
  const ss = SpreadsheetApp.openById(PE_SS_ID);
  const sh = ss.getSheetByName('Settings') || ss.getSheetByName('設定 Settings');
  if (!sh) throw new Error('找不到 Settings 分頁');
  const data = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  for (const r of data) {
    if (String(r[0]).trim() === key) return String(r[1]).trim();
  }
  throw new Error('設定 ' + key + ' 不存在');
}

function pe_pageToken_() { return pe_getSetting_('FB_PAGE_TOKEN'); }
function pe_igId_()      { return pe_getSetting_('IG_BUSINESS_ACCOUNT_ID'); }
function pe_pageId_()    { return pe_getSetting_('FB_PAGE_ID'); }

/* =========================================================
 *  主排程器：每 10 分鐘執行一次
 * ========================================================= */
function processPostingQueue() {
  const ss = SpreadsheetApp.openById(PE_SS_ID);
  const sh = ss.getSheetByName(PE_QUEUE_NAME);
  if (!sh) { Logger.log('找不到 Posting_Queue'); return; }

  const last = sh.getLastRow();
  if (last < 2) return;
  const lastCol = Math.max(26, sh.getLastColumn());
  const data = sh.getRange(2, 1, last - 1, lastCol).getValues();
  const now = new Date();

  let posted = 0, failed = 0, skipped = 0;
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const rowNum = i + 2;
    const date = r[1];     // B 排程日期
    const time = r[2];     // C 排程時間
    const platform = r[3]; // D 平台
    const imgApprove = r[13];
    const copyApprove = r[14];
    const status = r[15];  // 排程狀態

    if (status !== '已排程') { skipped++; continue; }
    if (imgApprove !== '過' || copyApprove !== '過') { skipped++; continue; }

    // 組合排程時間（台灣時區）
    const scheduledAt = parseDateTime_(date, time);
    if (!scheduledAt) { skipped++; continue; }
    if (scheduledAt > now) { skipped++; continue; } // 還沒到時間

    // 觸發發文
    try {
      const result = publishRow_(sh, rowNum, r);
      if (result.ok) {
        sh.getRange(rowNum, 16).setValue('已發布');                              // 排程狀態
        sh.getRange(rowNum, 17).setValue(Utilities.formatDate(now, PE_TZ, 'yyyy-MM-dd HH:mm:ss')); // 發文時間
        sh.getRange(rowNum, 18).setValue(result.post_id);
        sh.getRange(rowNum, 19).setValue(result.post_url);
        sh.getRange(rowNum, 20).setValue('');
        posted++;
      } else {
        sh.getRange(rowNum, 16).setValue('失敗');
        sh.getRange(rowNum, 20).setValue(result.error);
        failed++;
      }
    } catch (e) {
      sh.getRange(rowNum, 16).setValue('失敗');
      sh.getRange(rowNum, 20).setValue(String(e));
      failed++;
    }

    Utilities.sleep(2000); // API rate limit 緩衝
  }
  Logger.log(`processPostingQueue: 發 ${posted}、失 ${failed}、跳 ${skipped}`);
}

function parseDateTime_(date, time) {
  try {
    const dStr = (date instanceof Date) ? Utilities.formatDate(date, PE_TZ, 'yyyy-MM-dd') : String(date);
    const tStr = (time instanceof Date) ? Utilities.formatDate(time, PE_TZ, 'HH:mm') : String(time);
    return new Date(dStr + 'T' + tStr + ':00+08:00');
  } catch (e) {
    return null;
  }
}

/* =========================================================
 *  發某一列：根據平台分流
 * ========================================================= */
function publishRow_(sh, rowNum, r) {
  const platform = r[3];
  const driveUrl = r[6];
  const headline = r[9];
  const body = r[10];
  const hashtags = r[11];
  const cta = r[12];
  const publishType = r[23] || 'single';   // X 發布類型
  const carouselIds = String(r[24] || '');  // Y 輪播 file_ids

  const caption = `${headline}\n\n${body}\n\n${hashtags}\n\n${cta}`;

  // 輪播分支
  if (publishType === 'carousel') {
    const ids = carouselIds.split(',').map(function(s){return s.trim();}).filter(function(s){return s;});
    if (ids.length < 2) return { ok: false, error: '輪播需 2-10 張、目前只有 ' + ids.length };
    const urls = ids.map(function(id){return `https://drive.google.com/thumbnail?id=${id}&sz=w1600`;});
    return publishCarousel_(platform, urls, caption);
  }

  // 單張分支 (原邏輯)
  const fileId = extractDriveFileId_(driveUrl);
  if (!fileId) return { ok: false, error: '無法從 Drive URL 取得 file_id' };
  const directUrl = `https://drive.google.com/thumbnail?id=${fileId}&sz=w1600`;

  if (platform.indexOf('IG Reels') >= 0) {
    return publishIGReel_(directUrl, caption);
  } else if (platform.indexOf('IG') >= 0 && platform.indexOf('FB') >= 0) {
    const ig = publishIGPost_(directUrl, caption);
    const fb = publishFBPhoto_(directUrl, caption);
    if (ig.ok && fb.ok) {
      return { ok: true, post_id: `IG:${ig.post_id} FB:${fb.post_id}`, post_url: ig.post_url + ' | ' + fb.post_url };
    }
    return { ok: false, error: `IG: ${ig.error||'OK'} | FB: ${fb.error||'OK'}` };
  } else if (platform.indexOf('IG') >= 0) {
    return publishIGPost_(directUrl, caption);
  } else if (platform.indexOf('FB') >= 0) {
    return publishFBPhoto_(directUrl, caption);
  }
  return { ok: false, error: '未知平台 ' + platform };
}

/* =========================================================
 *  輪播發布 (IG + FB)
 * ========================================================= */
function publishCarousel_(platform, imageUrls, caption) {
  const wantIG = platform.indexOf('IG') >= 0;
  const wantFB = platform.indexOf('FB') >= 0;

  let igRes = { ok: true }, fbRes = { ok: true };
  if (wantIG) igRes = publishIGCarousel_(imageUrls, caption);
  if (wantFB) fbRes = publishFBCarousel_(imageUrls, caption);

  if (wantIG && wantFB) {
    if (igRes.ok && fbRes.ok) {
      return { ok: true, post_id: `IG:${igRes.post_id} FB:${fbRes.post_id}`, post_url: igRes.post_url + ' | ' + fbRes.post_url };
    }
    return { ok: false, error: `IG: ${igRes.error||'OK'} | FB: ${fbRes.error||'OK'}` };
  } else if (wantIG) {
    return igRes;
  } else if (wantFB) {
    return fbRes;
  }
  return { ok: false, error: '未知平台 ' + platform };
}

/* -------- IG 輪播 (Carousel) -------- */
function publishIGCarousel_(imageUrls, caption) {
  const igId = pe_igId_();
  const token = pe_pageToken_();

  // 1. 建每張子容器(is_carousel_item)
  const childIds = [];
  for (let i = 0; i < imageUrls.length; i++) {
    const res = UrlFetchApp.fetch(`https://graph.facebook.com/v19.0/${igId}/media`, {
      method: 'post',
      payload: { image_url: imageUrls[i], is_carousel_item: 'true', access_token: token },
      muteHttpExceptions: true
    });
    const d = JSON.parse(res.getContentText());
    if (!d.id) return { ok: false, error: 'IG carousel child[' + i + '] fail: ' + res.getContentText() };
    childIds.push(d.id);
    Utilities.sleep(1500);
  }

  // 2. 建輪播主容器
  const createRes = UrlFetchApp.fetch(`https://graph.facebook.com/v19.0/${igId}/media`, {
    method: 'post',
    payload: { media_type: 'CAROUSEL', children: childIds.join(','), caption: caption, access_token: token },
    muteHttpExceptions: true
  });
  const createData = JSON.parse(createRes.getContentText());
  if (!createData.id) return { ok: false, error: 'IG carousel create fail: ' + createRes.getContentText() };
  const creationId = createData.id;

  Utilities.sleep(8000);

  // 3. 發布
  const pubRes = UrlFetchApp.fetch(`https://graph.facebook.com/v19.0/${igId}/media_publish`, {
    method: 'post',
    payload: { creation_id: creationId, access_token: token },
    muteHttpExceptions: true
  });
  const pubData = JSON.parse(pubRes.getContentText());
  if (!pubData.id) return { ok: false, error: 'IG carousel publish fail: ' + pubRes.getContentText() };
  return { ok: true, post_id: pubData.id, post_url: `https://www.instagram.com/p/${pubData.id}/` };
}

/* -------- FB 輪播 (多圖貼文) -------- */
function publishFBCarousel_(imageUrls, caption) {
  const pageId = pe_pageId_();
  const token = pe_pageToken_();

  // 1. 上傳每張为 unpublished photo、拿到 photo_id
  const photoIds = [];
  for (let i = 0; i < imageUrls.length; i++) {
    const res = UrlFetchApp.fetch(`https://graph.facebook.com/v19.0/${pageId}/photos`, {
      method: 'post',
      payload: { url: imageUrls[i], published: 'false', access_token: token },
      muteHttpExceptions: true
    });
    const d = JSON.parse(res.getContentText());
    if (!d.id) return { ok: false, error: 'FB carousel photo[' + i + '] fail: ' + res.getContentText() };
    photoIds.push(d.id);
    Utilities.sleep(1000);
  }

  // 2. 發貼文、attached_media 帶入所有照片
  const attached = photoIds.map(function(id){return JSON.stringify({media_fbid: id});});
  const postRes = UrlFetchApp.fetch(`https://graph.facebook.com/v19.0/${pageId}/feed`, {
    method: 'post',
    payload: {
      message: caption,
      attached_media: '[' + attached.join(',') + ']',
      access_token: token
    },
    muteHttpExceptions: true
  });
  const postData = JSON.parse(postRes.getContentText());
  if (!postData.id) return { ok: false, error: 'FB carousel feed fail: ' + postRes.getContentText() };
  const fbPostId = postData.id;
  return { ok: true, post_id: fbPostId, post_url: `https://www.facebook.com/${pageId}/posts/${fbPostId.split('_')[1] || fbPostId}` };
}

function extractDriveFileId_(url) {
  const m = url.match(/[-\w]{25,}/);
  return m ? m[0] : null;
}

/* -------- IG 一般貼文 (Single image) -------- */
function publishIGPost_(imageUrl, caption) {
  const igId = pe_igId_();
  const token = pe_pageToken_();

  // 1. 建容器
  const createRes = UrlFetchApp.fetch(`https://graph.facebook.com/v19.0/${igId}/media`, {
    method: 'post',
    payload: { image_url: imageUrl, caption: caption, access_token: token },
    muteHttpExceptions: true
  });
  const createData = JSON.parse(createRes.getContentText());
  if (!createData.id) return { ok: false, error: 'IG create container fail: ' + createRes.getContentText() };
  const creationId = createData.id;

  // 等 5 秒讓容器處理完
  Utilities.sleep(5000);

  // 2. 發布
  const pubRes = UrlFetchApp.fetch(`https://graph.facebook.com/v19.0/${igId}/media_publish`, {
    method: 'post',
    payload: { creation_id: creationId, access_token: token },
    muteHttpExceptions: true
  });
  const pubData = JSON.parse(pubRes.getContentText());
  if (!pubData.id) return { ok: false, error: 'IG publish fail: ' + pubRes.getContentText() };

  return { ok: true, post_id: pubData.id, post_url: `https://www.instagram.com/p/${pubData.id}/` };
}

/* -------- IG Reels (9x16 video) -------- */
function publishIGReel_(videoUrl, caption) {
  // 注意：IG Reels API 需要 video_url + media_type=REELS
  // 我們的素材是 PNG，不是影片。降級為一般貼文 + 9x16 比例（IG 接受非方形）
  return publishIGPost_(videoUrl, caption);
}

/* -------- FB 粉專照片 -------- */
function publishFBPhoto_(imageUrl, caption) {
  const pageId = pe_pageId_();
  const token = pe_pageToken_();
  const res = UrlFetchApp.fetch(`https://graph.facebook.com/v19.0/${pageId}/photos`, {
    method: 'post',
    payload: { url: imageUrl, caption: caption, access_token: token },
    muteHttpExceptions: true
  });
  const data = JSON.parse(res.getContentText());
  if (!data.id) return { ok: false, error: 'FB publish fail: ' + res.getContentText() };
  return { ok: true, post_id: data.post_id || data.id, post_url: `https://www.facebook.com/${pageId}/posts/${(data.post_id||data.id).split('_')[1] || data.id}` };
}

/* =========================================================
 *  審核與排程操作
 * ========================================================= */
function approveAllImages() {
  const sh = SpreadsheetApp.openById(PE_SS_ID).getSheetByName(PE_QUEUE_NAME);
  const last = sh.getLastRow();
  if (last < 2) return;
  const range = sh.getRange(2, 14, last - 1, 1);
  const vals = range.getValues().map(r => r[0] === '待審' ? ['過'] : r);
  range.setValues(vals);
  SpreadsheetApp.getUi().alert(`圖片全部審過 (${last - 1} 列)`);
}

function approveAllCopies() {
  const sh = SpreadsheetApp.openById(PE_SS_ID).getSheetByName(PE_QUEUE_NAME);
  const last = sh.getLastRow();
  if (last < 2) return;
  const range = sh.getRange(2, 15, last - 1, 1);
  const vals = range.getValues().map(r => r[0] === '待審' ? ['過'] : r);
  range.setValues(vals);
  SpreadsheetApp.getUi().alert(`文案全部審過 (${last - 1} 列)`);
}

function scheduleAllApproved() {
  const sh = SpreadsheetApp.openById(PE_SS_ID).getSheetByName(PE_QUEUE_NAME);
  const last = sh.getLastRow();
  if (last < 2) return;
  const data = sh.getRange(2, 1, last - 1, 16).getValues();
  let n = 0;
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    if (r[13] === '過' && r[14] === '過' && r[15] === '草稿') {
      sh.getRange(i + 2, 16).setValue('已排程');
      n++;
    }
  }
  SpreadsheetApp.getUi().alert(`已排程 ${n} 列（雙審過的草稿）`);
}

function publishOneRow(rowNum) {
  const sh = SpreadsheetApp.openById(PE_SS_ID).getSheetByName(PE_QUEUE_NAME);
  const lastCol = Math.max(26, sh.getLastColumn());
  const r = sh.getRange(rowNum, 1, 1, lastCol).getValues()[0];
  const result = publishRow_(sh, rowNum, r);
  if (result.ok) {
    sh.getRange(rowNum, 16).setValue('已發布');
    sh.getRange(rowNum, 17).setValue(Utilities.formatDate(new Date(), PE_TZ, 'yyyy-MM-dd HH:mm:ss'));
    sh.getRange(rowNum, 18).setValue(result.post_id);
    sh.getRange(rowNum, 19).setValue(result.post_url);
    Logger.log('已發：' + result.post_url);
  } else {
    sh.getRange(rowNum, 16).setValue('失敗');
    sh.getRange(rowNum, 20).setValue(result.error);
    Logger.log('失敗：' + result.error);
  }
  return result;
}

/* =========================================================
 *  時觸發器安裝
 * ========================================================= */
function installTriggers() {
  uninstallTriggers();
  ScriptApp.newTrigger('processPostingQueue').timeBased().everyMinutes(10).create();
  ScriptApp.newTrigger('snapshotInsightsDaily').timeBased().atHour(23).everyDays(1).create();
  SpreadsheetApp.getUi().alert('觸發器安裝完成：\n• processPostingQueue 每 10 分\n• snapshotInsightsDaily 每日 23:00');
}

function uninstallTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['processPostingQueue', 'snapshotInsightsDaily'].indexOf(t.getHandlerFunction()) >= 0) {
      ScriptApp.deleteTrigger(t);
    }
  });
}

/* =========================================================
 *  Insights 抓取（IG/FB）
 * ========================================================= */
function snapshotInsightsDaily() {
  const ss = SpreadsheetApp.openById(PE_SS_ID);
  const qSh = ss.getSheetByName(PE_QUEUE_NAME);
  const iSh = ss.getSheetByName('Insights');
  if (!qSh || !iSh) return;

  const last = qSh.getLastRow();
  if (last < 2) return;
  const data = qSh.getRange(2, 1, last - 1, 22).getValues();
  const today = Utilities.formatDate(new Date(), PE_TZ, 'yyyy-MM-dd');

  for (const r of data) {
    if (r[15] !== '已發布') continue;
    const queueId = r[0];
    const platform = r[3];
    const postId = r[17];
    if (!postId) continue;

    // 抓 IG
    const igPart = (postId.match(/IG:(\d+)/) || [])[1];
    if (igPart) fetchAndAppendIG_(iSh, today, queueId, igPart);
    // 抓 FB
    const fbPart = (postId.match(/FB:([\d_]+)/) || [])[1] || (platform === 'FB Post' ? postId : null);
    if (fbPart) fetchAndAppendFB_(iSh, today, queueId, fbPart);

    Utilities.sleep(1000);
  }
}

function fetchAndAppendIG_(iSh, today, queueId, mediaId) {
  try {
    const url = `https://graph.facebook.com/v19.0/${mediaId}/insights?metric=impressions,reach,likes,comments,saved,shares&access_token=${pe_pageToken_()}`;
    const res = JSON.parse(UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText());
    if (!res.data) return;
    const m = {};
    res.data.forEach(d => { m[d.name] = d.values[0].value; });
    iSh.appendRow([
      `IG_${mediaId}_${today}`, today, 'IG', mediaId, queueId,
      m.impressions || 0, m.reach || 0, m.likes || 0, m.comments || 0, m.saved || 0, m.shares || 0,
      0, 0,
      ((m.likes || 0) + (m.comments || 0) + (m.saved || 0)) / Math.max(m.reach || 1, 1),
      '', '', ''
    ]);
  } catch (e) { Logger.log('IG insight fail: ' + e); }
}

function fetchAndAppendFB_(iSh, today, queueId, postId) {
  try {
    const url = `https://graph.facebook.com/v19.0/${postId}/insights?metric=post_impressions,post_impressions_unique,post_reactions_like_total&access_token=${pe_pageToken_()}`;
    const res = JSON.parse(UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText());
    if (!res.data) return;
    const m = {};
    res.data.forEach(d => { m[d.name] = d.values[0].value; });
    iSh.appendRow([
      `FB_${postId}_${today}`, today, 'FB', postId, queueId,
      m.post_impressions || 0, m.post_impressions_unique || 0, m.post_reactions_like_total || 0, 0, 0, 0,
      0, 0,
      0, '', '', ''
    ]);
  } catch (e) { Logger.log('FB insight fail: ' + e); }
}
