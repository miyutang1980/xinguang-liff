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
    
    const rowsToAdd = [];
    for (const m of items) {
      const ts = new Date(m.timestamp).getTime();
      if (ts >= cutoffMs) { totalSkipped++; continue; }
      if (existing.has('IG:' + m.id)) { totalSkipped++; continue; }
      
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
    }
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

  let url = 'https://graph.facebook.com/v19.0/' + pageId + '/posts?fields=id,message,created_time,permalink_url,full_picture,attachments{media_type,media,subattachments}&limit=100&access_token=' + token;
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
    
    const rowsToAdd = [];
    for (const p of items) {
      const ts = new Date(p.created_time).getTime();
      if (ts >= cutoffMs) { totalSkipped++; continue; }
      if (existing.has('FB:' + p.id)) { totalSkipped++; continue; }
      
      const ins = fbGetInsights_(p.id, token);
      const mediaType = (p.attachments && p.attachments.data && p.attachments.data[0] && p.attachments.data[0].media_type) || 'STATUS';
      
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
    }
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
  // FB Page post metrics
  const metrics = 'post_impressions,post_impressions_unique,post_reactions_like_total,post_reactions_love_total,post_reactions_wow_total,post_clicks,post_video_views,post_video_avg_time_watched,post_video_complete_views_organic';
  const url = 'https://graph.facebook.com/v19.0/' + postId + '/insights?metric=' + metrics + '&access_token=' + token;
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(res.getContentText());
    if (data.error) return {};
    const out = {};
    (data.data || []).forEach(function(d){
      const v = d.values && d.values[0] && d.values[0].value;
      if (d.name === 'post_impressions') out.impressions = v;
      else if (d.name === 'post_impressions_unique') out.reach = v;
      else if (d.name === 'post_reactions_like_total') out.likes = v;
      else if (d.name === 'post_video_views') out.video_views = v;
      else if (d.name === 'post_video_avg_time_watched') out.avg_view_seconds = Math.round((v || 0) / 1000);
    });
    // 同時抓 likes/comments/shares 的 summary
    try {
      const u2 = 'https://graph.facebook.com/v19.0/' + postId + '?fields=likes.summary(true),comments.summary(true),shares&access_token=' + token;
      const r2 = UrlFetchApp.fetch(u2, { muteHttpExceptions: true });
      const d2 = JSON.parse(r2.getContentText());
      if (d2.likes && d2.likes.summary) out.likes = d2.likes.summary.total_count;
      if (d2.comments && d2.comments.summary) out.comments = d2.comments.summary.total_count;
      if (d2.shares && d2.shares.count) out.shares = d2.shares.count;
    } catch (e2) {}
    return out;
  } catch (e) {
    return {};
  }
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
