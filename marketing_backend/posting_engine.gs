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
  const publishType = r[23] || 'single';   // X 發布類型：single / carousel / reel
  const carouselIds = String(r[24] || '');  // Y 輪播 file_ids、或 Reel 的影片公開 URL（並存不衝突）

  const caption = `${headline}\n\n${body}\n\n${hashtags}\n\n${cta}`;

  // ---- Reel 分支（短影音）----
  if (publishType === 'reel') {
    // Y 欄裝 .mp4 公開 URL（例如 GitHub Releases https 直連）
    const videoUrl = String(carouselIds || driveUrl || '').trim();
    if (!/^https?:\/\//i.test(videoUrl)) {
      return { ok: false, error: 'Reel 需要公開的 .mp4 URL、請填在 Y 欄（影片連結）' };
    }
    // 取得 D 列「平台」、Q 列若有「排程時間」字串、傳給 Buffer 排程
    const dPart2 = r[1] ? Utilities.formatDate(new Date(r[1]), PE_TZ, 'yyyy-MM-dd') : '';
    const tPart2 = r[2] ? (typeof r[2] === 'string' ? r[2] : Utilities.formatDate(new Date(r[2]), PE_TZ, 'HH:mm')) : '';
    const scheduleAtIso = (dPart2 && tPart2) ? `${dPart2}T${tPart2}` : null;
    return publishReel_(platform, videoUrl, caption, scheduleAtIso);
  }

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

/* =========================================================
 *  Reel 發布（IG Reels + FB Reels）
 * ========================================================= */
function publishReel_(platform, videoUrl, caption, scheduleAtIso) {
  // 路由策略（全部走 Buffer）：
  //   IG / FB / TT 三平台都透過 Buffer GraphQL 排程
  //   原因：
  //     - IG：App Review 沒過、API 需透過 Buffer
  //     - FB：自家以 R2 URL 送、FB crawler 對無 robots.txt 的 R2 回 403
  //     - TT：API 不公開、Buffer 是唯一選項
  const wantIG = platform.indexOf('IG') >= 0;
  const wantFB = platform.indexOf('FB') >= 0;
  const wantTT = platform.indexOf('TT') >= 0 || platform.indexOf('TikTok') >= 0;

  let igRes = { ok: true, skipped: true };
  let fbRes = { ok: true, skipped: true };
  let ttRes = { ok: true, skipped: true };

  const bufferTargets = [];
  if (wantIG) bufferTargets.push('IG');
  if (wantFB) bufferTargets.push('FB');
  if (wantTT) bufferTargets.push('TT');
  if (bufferTargets.length > 0) {
    const bRes = publishViaBuffer_(bufferTargets, videoUrl, caption, scheduleAtIso);
    if (wantIG) igRes = bRes.IG || { ok: false, error: 'Buffer 無 IG 回應' };
    if (wantFB) fbRes = bRes.FB || { ok: false, error: 'Buffer 無 FB 回應' };
    if (wantTT) ttRes = bRes.TT || { ok: false, error: 'Buffer 無 TT 回應' };
  }

  // 彙整結果
  const parts = [], errs = [], ids = [], urls = [];
  [['IG', wantIG, igRes], ['FB', wantFB, fbRes], ['TT', wantTT, ttRes]].forEach(([k, want, r]) => {
    if (!want) return;
    parts.push(k);
    if (r.ok) {
      if (r.post_id) ids.push(`${k}:${r.post_id}`);
      if (r.post_url) urls.push(r.post_url);
    } else {
      errs.push(`${k}: ${r.error || '未知錯誤'}`);
    }
  });
  if (parts.length === 0) return { ok: false, error: 'Reel 未指定平台' };
  if (errs.length > 0)  return { ok: false, error: errs.join(' | ') };
  return { ok: true, post_id: ids.join(' '), post_url: urls.join(' | ') };
}

/* =========================================================
 *  Buffer GraphQL 發布 / 排程
 *  targets: ['IG'] / ['TT'] / ['IG','TT']
 *  scheduleAtIso：'2026-05-08T19:00' (台灣時區字串、自動轉 UTC ms)
 *  回傳：{ IG: {ok,post_id}, TT: {ok,post_id} }
 * ========================================================= */
function publishViaBuffer_(targets, videoUrl, caption, scheduleAtIso) {
  const apiKey = pe_getSetting_('BUFFER_API_KEY');
  const orgId  = pe_getSetting_('BUFFER_ORG_ID');
  const igCh   = pe_getSetting_('BUFFER_IG_CHANNEL_ID');
  const ttCh   = pe_getSetting_('BUFFER_TIKTOK_CHANNEL_ID');
  if (!apiKey) {
    const err = { ok: false, error: '未設定 BUFFER_API_KEY' };
    return { IG: err, TT: err };
  }

  // 排程模式：有未來時間 → customScheduled；無/過期 → 現在+5分鐘
  // 原因：Buffer 拒絕過期 dueAt、所以一律保證未來時間
  let mode = 'customScheduled';
  let dueAt = null;
  const nowMs = Date.now();
  const minFuture = nowMs + 5 * 60 * 1000; // 5 分鐘後、讓 Buffer 有足夠間距
  if (scheduleAtIso) {
    const m = String(scheduleAtIso).replace('T', ' ').match(/^(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2})/);
    if (m) {
      const localStr = `${m[1]}-${m[2]}-${m[3]}T${m[4]}:${m[5]}:00+08:00`;
      const d = new Date(localStr);
      if (!isNaN(d.getTime()) && d.getTime() > minFuture) {
        dueAt = d.toISOString();
      }
    }
  }
  // 未能解析或時間已過、退回「現在+5分鐘」
  if (!dueAt) {
    dueAt = new Date(minFuture).toISOString();
  }

  const fbCh   = pe_getSetting_('BUFFER_FB_CHANNEL_ID');
  const result = {};
  targets.forEach(t => {
    let channelId, metadata;
    if (t === 'IG') {
      channelId = igCh;
      metadata = { instagram: { type: 'reel', shouldShareToFeed: true } };
    } else if (t === 'FB') {
      channelId = fbCh;
      metadata = { facebook: { type: 'reel' } };
    } else { // TT
      channelId = ttCh;
      metadata = { tiktok: { title: (caption.split('\n')[0] || '').substring(0, 90) } };
    }
    if (!channelId) {
      result[t] = { ok: false, error: `未設定 Buffer ${t} channel ID` };
      return;
    }

    const input = {
      channelId: channelId,
      text: caption,
      mode: mode,
      schedulingType: 'automatic',
      assets: { videos: [{ url: videoUrl }] },
      metadata: metadata,
      source: 'xinguang_dashboard'
    };
    if (dueAt) input.dueAt = dueAt;

    const res = UrlFetchApp.fetch('https://api.buffer.com/graphql', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + apiKey },
      payload: JSON.stringify({
        query: 'mutation($input: CreatePostInput!){ createPost(input:$input){ __typename ... on PostActionSuccess { post { id } } ... on NotFoundError { message } ... on UnauthorizedError { message } ... on UnexpectedError { message } ... on RestProxyError { message } ... on LimitReachedError { message } ... on InvalidInputError { message } } }',
        variables: { input: input }
      }),
      muteHttpExceptions: true
    });
    const txt = res.getContentText();
    let data;
    try { data = JSON.parse(txt); } catch (e) { result[t] = { ok: false, error: 'Buffer 回傳非 JSON：' + txt.substring(0,200) }; return; }

    if (data.errors && data.errors.length) {
      result[t] = { ok: false, error: 'Buffer GraphQL 錯誤：' + JSON.stringify(data.errors).substring(0,300) };
      return;
    }
    const cp = data.data && data.data.createPost;
    if (cp && cp.__typename === 'PostActionSuccess' && cp.post && cp.post.id) {
      result[t] = { ok: true, post_id: cp.post.id, post_url: '', queued: true };
    } else if (cp && cp.message) {
      result[t] = { ok: false, error: `Buffer ${cp.__typename || ''}：${cp.message}` };
    } else {
      result[t] = { ok: false, error: 'Buffer 未知回應：' + txt.substring(0,300) };
    }
  });
  return result;
}

/* -------- IG Reels (真的 Reels API) -------- */
function publishIGReel_(videoUrl, caption) {
  const igId = pe_igId_();
  const token = pe_pageToken_();

  // 1. 建容器：media_type=REELS
  const createRes = UrlFetchApp.fetch(`https://graph.facebook.com/v19.0/${igId}/media`, {
    method: 'post',
    payload: {
      media_type: 'REELS',
      video_url: videoUrl,
      caption: caption,
      share_to_feed: 'true',
      access_token: token
    },
    muteHttpExceptions: true
  });
  const createData = JSON.parse(createRes.getContentText());
  if (!createData.id) return { ok: false, error: 'IG Reel 建容器失敗：' + createRes.getContentText() };
  const creationId = createData.id;

  // 2. 輪詢狀態（影片處理需時間）、最多等 90 秒
  let status = 'IN_PROGRESS', tries = 0;
  while (status === 'IN_PROGRESS' && tries < 30) {
    Utilities.sleep(3000);
    const sRes = UrlFetchApp.fetch(`https://graph.facebook.com/v19.0/${creationId}?fields=status_code&access_token=${token}`, { muteHttpExceptions: true });
    const sData = JSON.parse(sRes.getContentText());
    status = sData.status_code || 'ERROR';
    tries++;
  }
  if (status !== 'FINISHED') {
    return { ok: false, error: 'IG Reel 影片處理未完成（狀態：' + status + '）' };
  }

  // 3. 發布
  const pubRes = UrlFetchApp.fetch(`https://graph.facebook.com/v19.0/${igId}/media_publish`, {
    method: 'post',
    payload: { creation_id: creationId, access_token: token },
    muteHttpExceptions: true
  });
  const pubData = JSON.parse(pubRes.getContentText());
  if (!pubData.id) return { ok: false, error: 'IG Reel publish fail: ' + pubRes.getContentText() };
  return { ok: true, post_id: pubData.id, post_url: `https://www.instagram.com/reel/${pubData.id}/` };
}

/* -------- FB Reels -------- */
// FB Reels 需 resumable upload 三步：start → upload → finish
function publishFBReel_(videoUrl, caption) {
  const pageId = pe_pageId_();
  const token = pe_pageToken_();

  // 1. 「start」取得 video_id + upload_url
  const startRes = UrlFetchApp.fetch(`https://graph.facebook.com/v19.0/${pageId}/video_reels`, {
    method: 'post',
    payload: { upload_phase: 'start', access_token: token },
    muteHttpExceptions: true
  });
  const startData = JSON.parse(startRes.getContentText());
  if (!startData.video_id) return { ok: false, error: 'FB Reel start fail: ' + startRes.getContentText() };
  const videoId = startData.video_id;

  // 2. 「upload」讓 FB 從我們的 URL 抓影片（hosted file_url 模式）
  const uploadRes = UrlFetchApp.fetch(`https://rupload.facebook.com/video-upload/v19.0/${videoId}`, {
    method: 'post',
    headers: {
      Authorization: 'OAuth ' + token,
      file_url: videoUrl
    },
    muteHttpExceptions: true
  });
  const uploadData = JSON.parse(uploadRes.getContentText());
  if (uploadData.success !== true && !uploadData.success) {
    return { ok: false, error: 'FB Reel upload fail: ' + uploadRes.getContentText() };
  }

  // 3. 「finish」發布
  const finishUrl = `https://graph.facebook.com/v19.0/${pageId}/video_reels?` +
    'access_token=' + encodeURIComponent(token) +
    '&video_id=' + encodeURIComponent(videoId) +
    '&upload_phase=finish' +
    '&video_state=PUBLISHED' +
    '&description=' + encodeURIComponent(caption);
  const finishRes = UrlFetchApp.fetch(finishUrl, { method: 'post', muteHttpExceptions: true });
  const finishData = JSON.parse(finishRes.getContentText());
  if (finishData.success !== true && !finishData.success) {
    return { ok: false, error: 'FB Reel finish fail: ' + finishRes.getContentText() };
  }
  return { ok: true, post_id: videoId, post_url: `https://www.facebook.com/reel/${videoId}` };
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

/* =========================================================
 *  appendReelToQueue：後台「新增 Reel 預排」表單入口
 *  payload = {
 *    videoUrl: '公開 .mp4 URL（必須）',
 *    platform: 'IG+FB' | 'IG' | 'FB',
 *    headline: '標題',
 *    body: '文案',
 *    hashtags: '#tag1 #tag2',
 *    cta: 'CTA 文字',
 *    scheduleAt: 'yyyy-MM-dd HH:mm'（台灣時區）
 *  }
 * ========================================================= */
function appendReelToQueue(payload) {
  try {
    const p = payload || {};
    const videoUrl = String(p.videoUrl || '').trim();
    if (!/^https?:\/\/.+\.mp4(\?.*)?$/i.test(videoUrl)) {
      return { ok: false, error: '影片 URL 必須是公開的 .mp4 直連（例：GitHub Releases）' };
    }
    const plat = String(p.platform || 'IG+FB').trim();
    const headline = String(p.headline || '').trim();
    const body = String(p.body || '').trim();
    const hashtags = String(p.hashtags || '').trim();
    const cta = String(p.cta || '').trim();
    const sAt = String(p.scheduleAt || '').trim();
    if (!sAt || !/^\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}/.test(sAt)) {
      return { ok: false, error: '排程時間格式錯（需 yyyy-MM-dd HH:mm）' };
    }
    const dPart = sAt.substring(0, 10);
    const tPart = sAt.substring(11, 16);

    const ss = SpreadsheetApp.openById(PE_SS_ID);
    const sh = ss.getSheetByName(PE_QUEUE_NAME);
    if (!sh) return { ok: false, error: '找不到 Posting_Queue 工作表' };

    // 確保 X/Y 欄頭存在
    const lastCol = Math.max(26, sh.getLastColumn());
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    if (headers[23] !== '發布類型') sh.getRange(1, 24).setValue('發布類型');
    if (headers[24] !== '輪播圖file_ids') sh.getRange(1, 25).setValue('輪播圖file_ids');

    // 產生列號 ID
    const last = sh.getLastRow();
    const newId = 'R' + Utilities.formatDate(new Date(), PE_TZ, 'yyyyMMddHHmmss');

    // 寫入一列：A=ID, B=日期, C=時間, D=平台, E=主題, F=形式, G=主圖URL(借位放影片URL方便預覽), H=縮圖, I=備註,
    //        J=headline, K=body, L=hashtags, M=cta, N=圖片審(過), O=文案審(過), P=排程狀態(已排程), Q=發文時間, R=post_id, S=post_url, T=錯誤,
    //        U/V/W=保留, X=發布類型(reel), Y=影片URL
    // 預排表單提交 = 已決定發該貼文、自動雙審過 + 已排程、不需手動過審
    const row = [
      newId, dPart, tPart, plat, '短影音 Reel', 'Reel', videoUrl, '', '',
      headline, body, hashtags, cta,
      '過', '過', '已排程', '', '', '', '',
      '', '', '',
      'reel', videoUrl, ''
    ];
    sh.appendRow(row);
    return { ok: true, id: newId, row: sh.getLastRow() };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

/* =========================================================
 *  testPublish：手動觸發某列發布、繞過時間檢查
 *  用法：把下面 rowNum 改成 Sheet 的列號 → ▶ 執行 → 看執行記錄
 * ========================================================= */
function testPublish() {
  const rowNum = 14;   // ⚠️ 改成你那列在 Sheet 的列號
  publishOneRow(rowNum);
}
