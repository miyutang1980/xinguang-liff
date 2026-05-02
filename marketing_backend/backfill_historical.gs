/**
 * 歷史數據回填：抓 IG + FB 帳號全部歷史貼文（到 2026/5/1 之前）
 * 寫入 Sheet 分頁「Historical_Posts」、每篇貼文一列、含基本互動 + 觸及 + 影片指標
 *
 * 使用方式：
 *   1. backfillIGHistorical()    — 抓 IG 全部歷史貼文（IMAGE / VIDEO / CAROUSEL_ALBUM / REELS）
 *   2. backfillFBHistorical()    — 抓 FB 全部歷史貼文
 *   3. backfillAllHistorical()   — 兩個一次跑
 *
 * 注意：
 *   - 跑大量資料可能 6 分鐘超時、超時就再跑一次（會從 LastCursor 繼續）
 *   - 4/30 之後的不抓（CUTOFF）
 *   - 每篇貼文同步抓 insights
 */

const HIST_SS_ID = '1DybgWBdCyvkEijMyaE46rKLtQD9J2ImjU8xeYCKSKnA';
const HIST_SHEET = 'Historical_Posts';
const HIST_CUTOFF = '2026-05-01';  // 不抓 5/1 以後
const HIST_TZ = 'Asia/Taipei';

function hist_getSetting_(key) {
  const ss = SpreadsheetApp.openById(HIST_SS_ID);
  const sh = ss.getSheetByName('Settings') || ss.getSheetByName('設定 Settings');
  const data = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  for (const r of data) {
    if (String(r[0]).trim() === key) return String(r[1]).trim();
  }
  throw new Error('設定 ' + key + ' 不存在');
}

function hist_ensureSheet_() {
  const ss = SpreadsheetApp.openById(HIST_SS_ID);
  let sh = ss.getSheetByName(HIST_SHEET);
  if (!sh) {
    sh = ss.insertSheet(HIST_SHEET);
    const headers = [
      '抓取時間', '平台', '貼文ID', '發布時間', '類型',
      '永久連結', '主圖URL', '說明文字',
      '按讚', '留言', '分享', '儲存',
      '觸及', '曝光',
      '影片觀看', '平均觀看秒', '完看率',
      'caption_前60字'
    ];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1f3864').setFontColor('#fff');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 130);
    sh.setColumnWidth(6, 200);
    sh.setColumnWidth(8, 250);
    sh.setColumnWidth(18, 250);
  }
  return sh;
}

function hist_appendRows_(rows) {
  if (rows.length === 0) return;
  const sh = hist_ensureSheet_();
  const last = sh.getLastRow();
  sh.getRange(last + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function hist_existingPostIds_() {
  const sh = hist_ensureSheet_();
  const last = sh.getLastRow();
  if (last < 2) return new Set();
  const ids = sh.getRange(2, 3, last - 1, 1).getValues().map(function(r){return String(r[0]);});
  return new Set(ids);
}

/* =========================================================
 *  IG 歷史回填
 * ========================================================= */
function backfillIGHistorical() {
  const igId = hist_getSetting_('IG_BUSINESS_ACCOUNT_ID');
  const token = hist_getSetting_('FB_PAGE_TOKEN');
  const cutoffMs = new Date(HIST_CUTOFF + 'T00:00:00+08:00').getTime();
  const existing = hist_existingPostIds_();
  const startTime = new Date().getTime();
  const TIMEOUT_MS = 5 * 60 * 1000;  // 5 分鐘

  let url = 'https://graph.facebook.com/v19.0/' + igId + '/media?fields=id,caption,media_type,media_url,thumbnail_url,permalink,timestamp,like_count,comments_count,is_shared_to_feed&limit=100&access_token=' + token;
  let totalAdded = 0, totalSkipped = 0, pages = 0;

  while (url) {
    if (new Date().getTime() - startTime > TIMEOUT_MS) {
      Logger.log('⚠️ 5 分鐘超時、已新增 ' + totalAdded + ' 篇、再跑一次會從下一頁繼續');
      break;
    }
    pages++;
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(res.getContentText());
    if (data.error) {
      Logger.log('❌ IG API 錯誤: ' + JSON.stringify(data.error));
      break;
    }
    const items = data.data || [];
    Logger.log('第 ' + pages + ' 頁、' + items.length + ' 筆');
    if (items.length > 0 && pages === 1) {
      Logger.log('  最新一篇 timestamp (UTC): ' + items[0].timestamp);
      Logger.log('  最舊一篇 timestamp (UTC): ' + items[items.length - 1].timestamp);
      Logger.log('  cutoff (台北): ' + HIST_CUTOFF + ' 00:00 = UTC ' + new Date(cutoffMs).toISOString());
      Logger.log('  existing set 大小: ' + existing.size);
    }
    
    const rowsToAdd = [];
    let pageReasonAfter = 0, pageReasonExist = 0, pageReasonAdded = 0;
    for (const m of items) {
      const ts = new Date(m.timestamp).getTime();
      if (ts >= cutoffMs) { totalSkipped++; pageReasonAfter++; continue; }
      if (existing.has('IG:' + m.id)) { totalSkipped++; pageReasonExist++; continue; }
      
      // 抓 insights
      const ins = igGetInsights_(m.id, m.media_type, token);
      
      rowsToAdd.push([
        Utilities.formatDate(new Date(), HIST_TZ, 'yyyy-MM-dd HH:mm:ss'),
        'IG',
        'IG:' + m.id,
        Utilities.formatDate(new Date(m.timestamp), HIST_TZ, 'yyyy-MM-dd HH:mm'),
        m.media_type || '',
        m.permalink || '',
        m.media_url || m.thumbnail_url || '',
        String(m.caption || '').substring(0, 500),
        m.like_count || 0,
        m.comments_count || 0,
        ins.shares || 0,
        ins.saved || 0,
        ins.reach || 0,
        ins.impressions || 0,
        ins.video_views || 0,
        ins.avg_view_seconds || 0,
        ins.completion_rate || 0,
        String(m.caption || '').substring(0, 60)
      ]);
      existing.add('IG:' + m.id);
      pageReasonAdded++;
    }
    Logger.log('  本頁細目：新增=' + pageReasonAdded + '、晚於cutoff被略=' + pageReasonAfter + '、已存在被略=' + pageReasonExist);
    if (rowsToAdd.length > 0) {
      hist_appendRows_(rowsToAdd);
      totalAdded += rowsToAdd.length;
    }
    
    // 全部 items 都早於 cutoff、停
    if (items.length > 0) {
      const oldestTs = new Date(items[items.length - 1].timestamp).getTime();
      if (oldestTs < cutoffMs && items.every(function(m){return new Date(m.timestamp).getTime() < cutoffMs;})) {
        // 繼續往下抓
      }
    }
    
    url = (data.paging && data.paging.next) || null;
  }
  
  Logger.log('===== IG 完成 =====');
  Logger.log('新增 ' + totalAdded + ' 篇、略過 ' + totalSkipped + ' 篇、共 ' + pages + ' 頁');
  if (totalAdded === 0 && totalSkipped > 0) {
    Logger.log('⚠️ 全部被略、請看上面「本頁細目」判斷原因：');
    Logger.log('   - 若「晚於cutoff」多：historical 貼文都看似是 5/1 以後、可能 cutoff 太早');
    Logger.log('   - 若「已存在」多：Historical_Posts 裡已有資料');
  }
}

function igGetInsights_(mediaId, mediaType, token) {
  // IG insights metrics 依類型而定
  let metrics = 'reach,saved,shares';
  if (mediaType === 'VIDEO' || mediaType === 'REELS') {
    metrics += ',video_views,avg_time_watched';
  } else if (mediaType === 'IMAGE' || mediaType === 'CAROUSEL_ALBUM') {
    metrics += ',impressions';
  }
  
  const url = 'https://graph.facebook.com/v19.0/' + mediaId + '/insights?metric=' + metrics + '&access_token=' + token;
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(res.getContentText());
    if (data.error) return {};
    const out = {};
    (data.data || []).forEach(function(d){
      const v = d.values && d.values[0] && d.values[0].value;
      if (d.name === 'reach') out.reach = v;
      else if (d.name === 'impressions') out.impressions = v;
      else if (d.name === 'saved') out.saved = v;
      else if (d.name === 'shares') out.shares = v;
      else if (d.name === 'video_views') out.video_views = v;
      else if (d.name === 'avg_time_watched') out.avg_view_seconds = Math.round((v || 0) / 1000);
    });
    // 完看率（只有 video）若有 length
    return out;
  } catch (e) {
    return {};
  }
}

/* =========================================================
 *  FB 歷史回填
 * ========================================================= */
function backfillFBHistorical() {
  const pageId = hist_getSetting_('FB_PAGE_ID');
  const token = hist_getSetting_('FB_PAGE_TOKEN');
  const cutoffMs = new Date(HIST_CUTOFF + 'T00:00:00+08:00').getTime();
  const existing = hist_existingPostIds_();
  const startTime = new Date().getTime();
  const TIMEOUT_MS = 5 * 60 * 1000;

  // Meta 對這個 Page 特別敏感、fields 只拿最基本三欄、limit 限 10
  // full_picture / message 都改為單篇查詢
  const fbFields = 'id,created_time,permalink_url';
  let url = 'https://graph.facebook.com/v19.0/' + pageId + '/posts?fields=' + fbFields + '&limit=10&access_token=' + token;
  let totalAdded = 0, totalSkipped = 0, pages = 0;

  while (url) {
    if (new Date().getTime() - startTime > TIMEOUT_MS) {
      Logger.log('⚠️ 5 分鐘超時、已新增 ' + totalAdded + ' 篇');
      break;
    }
    pages++;
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(res.getContentText());
    if (data.error) {
      Logger.log('❌ FB API 錯誤: ' + JSON.stringify(data.error));
      break;
    }
    const items = data.data || [];
    Logger.log('FB 第 ' + pages + ' 頁、' + items.length + ' 筆');
    if (items.length > 0 && pages === 1) {
      Logger.log('  最新一篇 created_time: ' + items[0].created_time);
      Logger.log('  cutoff (台北): ' + HIST_CUTOFF + ' 00:00 = UTC ' + new Date(cutoffMs).toISOString());
    }
    
    const rowsToAdd = [];
    let pageReasonAfter = 0, pageReasonExist = 0, pageReasonAdded = 0;
    for (const p of items) {
      const ts = new Date(p.created_time).getTime();
      if (ts >= cutoffMs) { totalSkipped++; pageReasonAfter++; continue; }
      if (existing.has('FB:' + p.id)) { totalSkipped++; pageReasonExist++; continue; }
      
      const ins = fbGetInsights_(p.id, token);
      // 一次拿 message + full_picture + attachments.media_type、單篇查不會被拒
      let pMessage = '', pPicture = '', mediaType = 'STATUS';
      try {
        const u3 = 'https://graph.facebook.com/v19.0/' + p.id + '?fields=' + encodeURIComponent('message,full_picture,attachments{media_type}') + '&access_token=' + token;
        const r3 = JSON.parse(UrlFetchApp.fetch(u3, {muteHttpExceptions:true}).getContentText());
        pMessage = r3.message || '';
        pPicture = r3.full_picture || '';
        if (r3.attachments && r3.attachments.data && r3.attachments.data[0]) {
          mediaType = r3.attachments.data[0].media_type || 'STATUS';
        } else if (pPicture) {
          mediaType = 'IMAGE';
        }
      } catch (eMt) {}
      // 套回 p 以便下面原來的 push code 能用
      p.message = pMessage;
      p.full_picture = pPicture;
      
      rowsToAdd.push([
        Utilities.formatDate(new Date(), HIST_TZ, 'yyyy-MM-dd HH:mm:ss'),
        'FB',
        'FB:' + p.id,
        Utilities.formatDate(new Date(p.created_time), HIST_TZ, 'yyyy-MM-dd HH:mm'),
        mediaType,
        p.permalink_url || '',
        p.full_picture || '',
        String(p.message || '').substring(0, 500),
        ins.likes || 0,
        ins.comments || 0,
        ins.shares || 0,
        0,  // FB 沒有 saved
        ins.reach || 0,
        ins.impressions || 0,
        ins.video_views || 0,
        ins.avg_view_seconds || 0,
        ins.completion_rate || 0,
        String(p.message || '').substring(0, 60)
      ]);
      existing.add('FB:' + p.id);
      pageReasonAdded++;
    }
    Logger.log('  本頁細目：新增=' + pageReasonAdded + '、晚於cutoff被略=' + pageReasonAfter + '、已存在被略=' + pageReasonExist);
    if (rowsToAdd.length > 0) {
      hist_appendRows_(rowsToAdd);
      totalAdded += rowsToAdd.length;
    }
    
    url = (data.paging && data.paging.next) || null;
  }
  
  Logger.log('===== FB 完成 =====');
  Logger.log('新增 ' + totalAdded + ' 篇、略過 ' + totalSkipped + ' 篇、共 ' + pages + ' 頁');
}

function fbGetInsights_(postId, token) {
  const out = { likes: 0, comments: 0, shares: 0, reach: 0, impressions: 0, video_views: 0, avg_view_seconds: 0, completion_rate: 0 };
  // 第一步：用 reactions.summary / comments.summary / shares 抓互動（必拿得到）
  try {
    const f1 = encodeURIComponent('reactions.summary(total_count).limit(0),comments.summary(total_count).limit(0),shares');
    const u1 = 'https://graph.facebook.com/v19.0/' + postId + '?fields=' + f1 + '&access_token=' + token;
    const r1 = UrlFetchApp.fetch(u1, { muteHttpExceptions: true });
    const d1 = JSON.parse(r1.getContentText());
    if (d1.error) {
      Logger.log('  fbGetInsights step1 error post=' + postId + ': ' + JSON.stringify(d1.error));
    } else {
      if (d1.reactions && d1.reactions.summary) out.likes = d1.reactions.summary.total_count || 0;
      if (d1.comments && d1.comments.summary) out.comments = d1.comments.summary.total_count || 0;
      if (d1.shares && d1.shares.count !== undefined) out.shares = d1.shares.count || 0;
    }
  } catch (e1) {
    Logger.log('  fbGetInsights step1 exception post=' + postId + ': ' + e1);
  }
  // 第二步：抓 insights（reach / impressions / video_views）— 失敗不影響上面
  try {
    const metrics = 'post_impressions,post_impressions_unique,post_video_views,post_video_avg_time_watched';
    const u2 = 'https://graph.facebook.com/v19.0/' + postId + '/insights?metric=' + metrics + '&access_token=' + token;
    const r2 = UrlFetchApp.fetch(u2, { muteHttpExceptions: true });
    const d2 = JSON.parse(r2.getContentText());
    if (!d2.error) {
      (d2.data || []).forEach(function(d){
        const v = d.values && d.values[0] && d.values[0].value;
        if (d.name === 'post_impressions') out.impressions = v || 0;
        else if (d.name === 'post_impressions_unique') out.reach = v || 0;
        else if (d.name === 'post_video_views') out.video_views = v || 0;
        else if (d.name === 'post_video_avg_time_watched') out.avg_view_seconds = Math.round((v || 0) / 1000);
      });
    }
  } catch (e2) {}
  return out;
}

/**
 * 補抓所有 FB 歷史貼文互動數（reactions/comments/shares）
 * 用 Historical_Posts 的 FB 列、回填欄位 I/J/K（讚/留言/分享）+ M/N（觸及/曝光）+ O（影片觀看）
 */
function fbBackfillInteractions() {
  const ss = SpreadsheetApp.openById(HIST_SS_ID);
  const sh = ss.getSheetByName('Historical_Posts');
  if (!sh || sh.getLastRow() < 2) { Logger.log('無資料'); return; }
  const token = hist_getSetting_('FB_PAGE_TOKEN');
  if (!token) { Logger.log('FB_PAGE_TOKEN 未設'); return; }
  
  const last = sh.getLastRow();
  const data = sh.getRange(2, 1, last - 1, 18).getValues();
  const startTime = new Date().getTime();
  const TIMEOUT_MS = 5 * 60 * 1000;
  
  let updated = 0, skipped = 0, errors = 0;
  const updates = []; // {row, likes, comments, shares, reach, impressions, video_views, avg}
  
  for (let i = 0; i < data.length; i++) {
    if (new Date().getTime() - startTime > TIMEOUT_MS) {
      Logger.log('5 分鐘超時、本輪處理 ' + updated + ' 筆、再跑一次補完剩下');
      break;
    }
    const r = data[i];
    if (r[1] !== 'FB') continue;
    const id = String(r[2] || '').replace(/^FB:/, '');
    if (!id) { skipped++; continue; }
    // 已經有互動就跳過
    const curLikes = Number(r[8]) || 0;
    const curComments = Number(r[9]) || 0;
    const curShares = Number(r[10]) || 0;
    if (curLikes + curComments + curShares > 0) { skipped++; continue; }
    
    const ins = fbGetInsights_(id, token);
    if (ins.likes + ins.comments + ins.shares + ins.reach + ins.video_views > 0) {
      updates.push({
        row: i + 2,
        likes: ins.likes, comments: ins.comments, shares: ins.shares,
        reach: ins.reach, impressions: ins.impressions,
        video_views: ins.video_views, avg: ins.avg_view_seconds
      });
      updated++;
    } else {
      errors++;
    }
    // 每 50 筆批次寫
    if (updates.length >= 50) {
      _fb_flushUpdates(sh, updates);
      updates.length = 0;
    }
  }
  if (updates.length > 0) _fb_flushUpdates(sh, updates);
  
  Logger.log('===== FB 互動補抓完成 =====');
  Logger.log('更新 ' + updated + ' 筆、跳過 ' + skipped + '（已有互動或非FB）、抓不到 ' + errors + ' 筆');
}

function _fb_flushUpdates(sh, updates) {
  updates.forEach(function(u){
    sh.getRange(u.row, 9, 1, 7).setValues([[u.likes, u.comments, u.shares, 0, u.reach, u.impressions, u.video_views]]);
    if (u.avg) sh.getRange(u.row, 16).setValue(u.avg);
  });
  SpreadsheetApp.flush();
}

/* =========================================================
 *  受眾洞察（粉絲、人口統計、地區）
 * ========================================================= */
function fetchAudienceDemographics() {
  const ss = SpreadsheetApp.openById(HIST_SS_ID);
  let sh = ss.getSheetByName('Audience_Snapshot');
  if (!sh) {
    sh = ss.insertSheet('Audience_Snapshot');
    sh.getRange(1, 1, 1, 6).setValues([['抓取時間', '平台', '指標類型', '維度', '值', '備註']])
      .setFontWeight('bold').setBackground('#1f3864').setFontColor('#fff');
    sh.setFrozenRows(1);
  }
  
  const igId = hist_getSetting_('IG_BUSINESS_ACCOUNT_ID');
  const pageId = hist_getSetting_('FB_PAGE_ID');
  const token = hist_getSetting_('FB_PAGE_TOKEN');
  const now = Utilities.formatDate(new Date(), HIST_TZ, 'yyyy-MM-dd HH:mm:ss');
  const rows = [];

  // === IG 受眾（IG 受眾 demographics 需要至少 100 粉、且只給 lifetime）===
  try {
    // 粉絲總數
    const u1 = 'https://graph.facebook.com/v19.0/' + igId + '?fields=followers_count,media_count,name,username&access_token=' + token;
    const r1 = JSON.parse(UrlFetchApp.fetch(u1, {muteHttpExceptions:true}).getContentText());
    if (r1.followers_count !== undefined) {
      rows.push([now, 'IG', '粉絲總數', '當前', r1.followers_count, '@' + (r1.username || '')]);
      rows.push([now, 'IG', '貼文總數', '當前', r1.media_count, '']);
    }
    
    // IG 受眾人口（age_gender / city / country）
    const dimensions = ['age', 'gender', 'city', 'country'];
    for (const dim of dimensions) {
      try {
        const u = 'https://graph.facebook.com/v19.0/' + igId + '/insights?metric=follower_demographics&period=lifetime&breakdown=' + dim + '&metric_type=total_value&access_token=' + token;
        const r = JSON.parse(UrlFetchApp.fetch(u, {muteHttpExceptions:true}).getContentText());
        if (r.data && r.data[0] && r.data[0].total_value && r.data[0].total_value.breakdowns) {
          const breakdowns = r.data[0].total_value.breakdowns[0];
          if (breakdowns && breakdowns.results) {
            for (const item of breakdowns.results) {
              rows.push([now, 'IG', '受眾_' + dim, item.dimension_values.join('/'), item.value, '']);
            }
          }
        }
      } catch (e) {
        Logger.log('IG ' + dim + ' 失敗: ' + e.message);
      }
    }
    
    // 近 30 天每日粉絲變化
    try {
      const u = 'https://graph.facebook.com/v19.0/' + igId + '/insights?metric=follower_count&period=day&access_token=' + token;
      const r = JSON.parse(UrlFetchApp.fetch(u, {muteHttpExceptions:true}).getContentText());
      if (r.data && r.data[0] && r.data[0].values) {
        for (const v of r.data[0].values) {
          rows.push([now, 'IG', '日新增粉絲', v.end_time.substring(0, 10), v.value, '']);
        }
      }
    } catch (e) {}
  } catch (e) {
    Logger.log('IG 受眾失敗: ' + e.message);
  }

  // === FB 受眾 ===
  try {
    const u1 = 'https://graph.facebook.com/v19.0/' + pageId + '?fields=fan_count,name,about&access_token=' + token;
    const r1 = JSON.parse(UrlFetchApp.fetch(u1, {muteHttpExceptions:true}).getContentText());
    if (r1.fan_count !== undefined) {
      rows.push([now, 'FB', '粉絲總數', '當前', r1.fan_count, r1.name || '']);
    }
    
    // FB Page fans 城市分布、性別年齡（注意：Meta 已逐步淘汰 page_fans_*、可能拿不到、用 page_fan_adds_unique 替代）
    const fbMetrics = [
      ['page_fans', 'lifetime', '總粉絲(lifetime)'],
      ['page_fan_adds_unique', 'days_28', '28天新增粉絲'],
      ['page_impressions_unique', 'days_28', '28天觸及人數'],
      ['page_post_engagements', 'days_28', '28天貼文互動']
    ];
    for (const [m, period, label] of fbMetrics) {
      try {
        const u = 'https://graph.facebook.com/v19.0/' + pageId + '/insights?metric=' + m + '&period=' + period + '&access_token=' + token;
        const r = JSON.parse(UrlFetchApp.fetch(u, {muteHttpExceptions:true}).getContentText());
        if (r.data && r.data[0] && r.data[0].values && r.data[0].values.length > 0) {
          const last = r.data[0].values[r.data[0].values.length - 1];
          rows.push([now, 'FB', label, last.end_time ? last.end_time.substring(0, 10) : '', last.value, m]);
        }
      } catch (e) {
        Logger.log('FB ' + m + ' 失敗: ' + e.message);
      }
    }
  } catch (e) {
    Logger.log('FB 受眾失敗: ' + e.message);
  }

  if (rows.length > 0) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
  }
  Logger.log('✓ Audience_Snapshot 寫入 ' + rows.length + ' 列');
}

/* =========================================================
 *  診斷工具：看实際資料兩端的 timestamp
 * ========================================================= */
function diagnoseHistoricalRange() {
  const igId = hist_getSetting_('IG_BUSINESS_ACCOUNT_ID');
  const pageId = hist_getSetting_('FB_PAGE_ID');
  const token = hist_getSetting_('FB_PAGE_TOKEN');
  
  Logger.log('===== IG 帳號認定 =====');
  const u0 = 'https://graph.facebook.com/v19.0/' + igId + '?fields=username,name,followers_count,media_count&access_token=' + token;
  Logger.log(JSON.parse(UrlFetchApp.fetch(u0, {muteHttpExceptions:true}).getContentText()));
  
  Logger.log('===== IG 最新 5 篇 =====');
  const u1 = 'https://graph.facebook.com/v19.0/' + igId + '/media?fields=id,timestamp,permalink,caption,media_type&limit=5&access_token=' + token;
  const r1 = JSON.parse(UrlFetchApp.fetch(u1, {muteHttpExceptions:true}).getContentText());
  (r1.data || []).forEach(function(m){
    Logger.log(m.timestamp + ' | ' + m.media_type + ' | ' + m.permalink);
  });
  
  Logger.log('===== FB Page 認定 =====');
  const u2 = 'https://graph.facebook.com/v19.0/' + pageId + '?fields=name,fan_count,about&access_token=' + token;
  Logger.log(JSON.parse(UrlFetchApp.fetch(u2, {muteHttpExceptions:true}).getContentText()));
  
  Logger.log('===== FB 最新 5 篇 =====');
  const u3 = 'https://graph.facebook.com/v19.0/' + pageId + '/posts?fields=id,created_time,permalink_url,message&limit=5&access_token=' + token;
  const r3 = JSON.parse(UrlFetchApp.fetch(u3, {muteHttpExceptions:true}).getContentText());
  (r3.data || []).forEach(function(p){
    Logger.log(p.created_time + ' | ' + p.permalink_url);
  });
}

/* =========================================================
 *  全部一次跑
 * ========================================================= */
function backfillAllHistorical() {
  Logger.log('===== 1/3 IG 歷史貼文 =====');
  backfillIGHistorical();
  Logger.log('');
  Logger.log('===== 2/3 FB 歷史貼文 =====');
  backfillFBHistorical();
  Logger.log('');
  Logger.log('===== 3/3 受眾洞察 =====');
  fetchAudienceDemographics();
  Logger.log('');
  Logger.log('🎉 全部完成');
}

/**
 * 診斷 FB token 類型 + 確認是否能讀自己 Page 的互動
 * 跑完看執行記錄
 */
function fbDiagnoseToken() {
  const token = hist_getSetting_('FB_PAGE_TOKEN');
  const pageId = hist_getSetting_('FB_PAGE_ID');
  Logger.log('========== FB Token 診斷 ==========');
  Logger.log('Page ID: ' + pageId);
  Logger.log('Token 長度: ' + (token ? token.length : 0));
  
  // 1. /me 端點看 token 主體
  try {
    const u = 'https://graph.facebook.com/v19.0/me?access_token=' + token;
    const r = JSON.parse(UrlFetchApp.fetch(u, {muteHttpExceptions:true}).getContentText());
    Logger.log('1. /me 結果: ' + JSON.stringify(r));
    if (r.id === pageId) {
      Logger.log('   ✅ 這是 Page Access Token、主體就是 Page、應該能讀自己貼文');
    } else if (r.id) {
      Logger.log('   ⚠️ 這是 User Access Token (id=' + r.id + ')、不是 Page Token、會被 PPCA 擋');
    }
  } catch (e) { Logger.log('1. /me 失敗: ' + e); }
  
  // 2. debug_token 看細節
  try {
    const u = 'https://graph.facebook.com/v19.0/debug_token?input_token=' + token + '&access_token=' + token;
    const r = JSON.parse(UrlFetchApp.fetch(u, {muteHttpExceptions:true}).getContentText());
    Logger.log('2. debug_token: ' + JSON.stringify(r.data || r));
  } catch (e) {}
  
  // 3. 試 me/feed（用 Page Token 時最不會被擋）
  try {
    const u = 'https://graph.facebook.com/v19.0/me/feed?fields=id,created_time&limit=2&access_token=' + token;
    const r = JSON.parse(UrlFetchApp.fetch(u, {muteHttpExceptions:true}).getContentText());
    Logger.log('3. me/feed 結果筆數: ' + (r.data ? r.data.length : 0));
    if (r.error) Logger.log('   錯誤: ' + JSON.stringify(r.error));
    else if (r.data && r.data[0]) Logger.log('   首筆: ' + JSON.stringify(r.data[0]));
  } catch (e) {}
  
  // 4. 試 me/posts + 一次帶 reactions/comments/shares fields（最關鍵測試）
  try {
    const fields = encodeURIComponent('id,created_time,reactions.summary(total_count).limit(0),comments.summary(total_count).limit(0),shares');
    const u = 'https://graph.facebook.com/v19.0/me/posts?fields=' + fields + '&limit=3&access_token=' + token;
    const r = JSON.parse(UrlFetchApp.fetch(u, {muteHttpExceptions:true}).getContentText());
    Logger.log('4. me/posts + reactions/comments/shares:');
    if (r.error) {
      Logger.log('   ❌ 錯誤: ' + JSON.stringify(r.error));
    } else if (r.data) {
      r.data.slice(0, 3).forEach(function(p, i){
        const reactions = p.reactions && p.reactions.summary ? p.reactions.summary.total_count : 'N/A';
        const comments = p.comments && p.comments.summary ? p.comments.summary.total_count : 'N/A';
        const shares = p.shares && p.shares.count !== undefined ? p.shares.count : 'N/A';
        Logger.log('   ' + (i+1) + '. ' + p.id + ' | ' + p.created_time + ' | reactions=' + reactions + '、comments=' + comments + '、shares=' + shares);
      });
      Logger.log('   ✅ 用 me/posts 可以拿到互動！下一步用 fbBackfillInteractionsViaMeFeed');
    }
  } catch (e) { Logger.log('4. 失敗: ' + e); }
  
  Logger.log('=====================================');
}

/**
 * 用 me/posts edge 補抓 FB 互動（避開 PPCA 限制）
 * 一次抓 100 筆、含 reactions/comments/shares 全部 fields、寫回 sheet
 */
function fbBackfillInteractionsViaMeFeed() {
  const ss = SpreadsheetApp.openById(HIST_SS_ID);
  const sh = ss.getSheetByName('Historical_Posts');
  if (!sh || sh.getLastRow() < 2) { Logger.log('無資料'); return; }
  const token = hist_getSetting_('FB_PAGE_TOKEN');
  if (!token) { Logger.log('FB_PAGE_TOKEN 未設'); return; }
  
  // 建 id => row 的索引
  const last = sh.getLastRow();
  const data = sh.getRange(2, 1, last - 1, 18).getValues();
  const idToRow = {}; // 'FB:xxxx' => row number
  data.forEach(function(r, i){
    if (r[1] === 'FB' && r[2]) idToRow[String(r[2])] = i + 2;
  });
  Logger.log('Sheet 有 ' + Object.keys(idToRow).length + ' 筆 FB 列待補');
  
  const fields = encodeURIComponent('id,created_time,reactions.summary(total_count).limit(0),comments.summary(total_count).limit(0),shares');
  let url = 'https://graph.facebook.com/v19.0/me/posts?fields=' + fields + '&limit=100&access_token=' + token;
  
  const startTime = new Date().getTime();
  const TIMEOUT_MS = 5 * 60 * 1000;
  let totalUpdated = 0, totalSeen = 0, pages = 0;
  
  while (url) {
    if (new Date().getTime() - startTime > TIMEOUT_MS) {
      Logger.log('5 分鐘超時、本輪更新 ' + totalUpdated + ' 筆');
      break;
    }
    pages++;
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data2 = JSON.parse(res.getContentText());
    if (data2.error) {
      Logger.log('❌ FB API 錯誤: ' + JSON.stringify(data2.error));
      break;
    }
    const items = data2.data || [];
    Logger.log('第 ' + pages + ' 頁、' + items.length + ' 筆');
    
    const updates = [];
    items.forEach(function(p){
      totalSeen++;
      const key = 'FB:' + p.id;
      const row = idToRow[key];
      if (!row) return; // sheet 沒這筆、跳過
      const likes = (p.reactions && p.reactions.summary) ? (p.reactions.summary.total_count || 0) : 0;
      const comments = (p.comments && p.comments.summary) ? (p.comments.summary.total_count || 0) : 0;
      const shares = (p.shares && p.shares.count !== undefined) ? (p.shares.count || 0) : 0;
      if (likes + comments + shares > 0) {
        updates.push({ row: row, likes: likes, comments: comments, shares: shares });
      }
    });
    
    // 批次寫入
    updates.forEach(function(u){
      sh.getRange(u.row, 9, 1, 3).setValues([[u.likes, u.comments, u.shares]]);
    });
    SpreadsheetApp.flush();
    totalUpdated += updates.length;
    
    url = (data2.paging && data2.paging.next) || null;
  }
  
  Logger.log('===== me/posts 補抓完成 =====');
  Logger.log('總共看過 ' + totalSeen + ' 筆 API 結果、更新 sheet 上 ' + totalUpdated + ' 筆');
}
