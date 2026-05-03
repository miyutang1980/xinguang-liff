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
  
  // 在後端組裝縮圖 URL、前端完全不用拼
  function toItem_(id) { return { id: id, thumb: 'https://drive.google.com/thumbnail?id=' + id + '&sz=w400' }; }
  
  return {
    main: mainId ? toItem_(mainId) : null,
    carousel: yIds.map(toItem_),
    backup: zIds.map(toItem_),
    publishType: r[17] || 'single'
  };
}

/* ========== 歷史分析（Historical_Posts + Audience_Snapshot 串連） ========== */
function tsToString_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
  }
  return String(v || '');
}

function getHistoricalAnalytics() {
  const ss = SpreadsheetApp.openById(DASH_SS_ID);
  const out = {
    summary: { total: 0, ig: 0, fb: 0, byType: {} },
    monthly: [],
    topPosts: [],
    weekdayHeat: [],
    audience: { ig: {}, fb: {}, dailyFollowers: [] },
    error: null
  };
  try {
    const sh = ss.getSheetByName('Historical_Posts');
    if (!sh || sh.getLastRow() < 2) {
      out.error = '尚未跑歷史回填、Historical_Posts 沒資料';
      return out;
    }
    const last = sh.getLastRow();
    const data = sh.getRange(2, 1, last - 1, 18).getValues();
    // 欄位：0抓取時間 1平台 2貼文ID 3發布時間 4類型 5永久連結 6主圖URL 7說明文字
    //       8按讚 9留言 10分享 11儲存 12觸及 13曝光 14影片觀看 15平均觀看秒 16完看率 17caption_前60字
    const monthly = {};      // YYYY-MM => { posts, reach, engagement, byPlatform: {IG, FB} }
    const weekdayMap = [0,0,0,0,0,0,0]; // 週一~週日 互動
    const weekdayCount = [0,0,0,0,0,0,0];
    const topByEng = [];
    
    data.forEach(function(r){
      const platform = String(r[1] || '');
      const ts = tsToString_(r[3]);
      const type = String(r[4] || '');
      const permalink = String(r[5] || '');
      const thumb = String(r[6] || '');
      const captionShort = String(r[17] || '');
      const likes = Number(r[8]) || 0;
      const comments = Number(r[9]) || 0;
      const shares = Number(r[10]) || 0;
      const saved = Number(r[11]) || 0;
      const reach = Number(r[12]) || 0;
      const videoViews = Number(r[14]) || 0;
      const eng = likes + comments + shares + saved;
      
      out.summary.total++;
      if (platform === 'IG') out.summary.ig++;
      else if (platform === 'FB') out.summary.fb++;
      out.summary.byType[type] = (out.summary.byType[type] || 0) + 1;
      
      // monthly
      if (ts.length >= 7) {
        const ym = ts.substring(0, 7);
        if (!monthly[ym]) monthly[ym] = { posts: 0, reach: 0, engagement: 0, ig: 0, fb: 0, videoViews: 0 };
        monthly[ym].posts++;
        monthly[ym].reach += reach;
        monthly[ym].engagement += eng;
        monthly[ym].videoViews += videoViews;
        if (platform === 'IG') monthly[ym].ig++; else if (platform === 'FB') monthly[ym].fb++;
      }
      
      // weekday（用發布時間推算）
      if (ts.length >= 10) {
        try {
          const d = new Date(ts.replace(' ', 'T') + ':00+08:00');
          if (!isNaN(d.getTime())) {
            const wd = (d.getDay() + 6) % 7; // 週一=0、週日=6
            weekdayMap[wd] += eng;
            weekdayCount[wd]++;
          }
        } catch (e) {}
      }
      
      topByEng.push({
        platform: platform,
        ts: ts,
        type: type,
        permalink: permalink,
        thumb: thumb,
        caption: captionShort,
        likes: likes, comments: comments, shares: shares, saved: saved,
        reach: reach, videoViews: videoViews,
        engagement: eng
      });
    });
    
    // monthly array 排序
    out.monthly = Object.keys(monthly).sort().map(function(ym){
      return Object.assign({ ym: ym }, monthly[ym]);
    });
    
    // weekday heatmap
    const wdNames = ['週一','週二','週三','週四','週五','週六','週日'];
    out.weekdayHeat = wdNames.map(function(name, i){
      return {
        name: name,
        avgEng: weekdayCount[i] > 0 ? Math.round(weekdayMap[i] / weekdayCount[i] * 10) / 10 : 0,
        posts: weekdayCount[i]
      };
    });
    
    // top 10 posts by engagement
    topByEng.sort(function(a, b){ return b.engagement - a.engagement; });
    out.topPosts = topByEng.slice(0, 10);
  } catch (e) {
    out.error = '歷史貼文讀取失敗: ' + String(e);
  }
  
  // 受眾洞察
  try {
    const ash = ss.getSheetByName('Audience_Snapshot');
    if (ash && ash.getLastRow() > 1) {
      const adata = ash.getRange(2, 1, ash.getLastRow() - 1, 6).getValues();
      // 取每個指標最新一筆
      const latest = {}; // key = platform|metric|dim => row
      adata.forEach(function(r){
        const ts = tsToString_(r[0]);
        const platform = String(r[1] || '');
        const metric = String(r[2] || '');
        const dim = String(r[3] || '');
        const value = r[4];
        const note = String(r[5] || '');
        const key = platform + '|' + metric + '|' + dim;
        if (!latest[key] || latest[key].ts < ts) {
          latest[key] = { ts: ts, platform: platform, metric: metric, dim: dim, value: value, note: note };
        }
      });
      
      const igAge = [], igGender = [], igCity = [], igCountry = [], dailyFollowers = [];
      Object.keys(latest).forEach(function(k){
        const v = latest[k];
        if (v.platform === 'IG') {
          if (v.metric === '粉絲總數' && v.dim === '當前') out.audience.ig.followers = v.value;
          else if (v.metric === '貼文總數') out.audience.ig.mediaCount = v.value;
          else if (v.metric === '受眾_age') igAge.push({ dim: v.dim, value: Number(v.value) || 0 });
          else if (v.metric === '受眾_gender') igGender.push({ dim: v.dim, value: Number(v.value) || 0 });
          else if (v.metric === '受眾_city') igCity.push({ dim: v.dim, value: Number(v.value) || 0 });
          else if (v.metric === '受眾_country') igCountry.push({ dim: v.dim, value: Number(v.value) || 0 });
          else if (v.metric === '日新增粉絲') dailyFollowers.push({ date: v.dim, value: Number(v.value) || 0 });
        } else if (v.platform === 'FB') {
          if (v.metric === '粉絲總數' && v.dim === '當前') out.audience.fb.followers = v.value;
          else if (v.metric === '28天觸及人數') out.audience.fb.reach28d = v.value;
          else if (v.metric === '28天新增粉絲') out.audience.fb.fanAdds28d = v.value;
          else if (v.metric === '28天貼文互動') out.audience.fb.engagement28d = v.value;
        }
      });
      igAge.sort(function(a,b){return String(a.dim).localeCompare(String(b.dim));});
      igCity.sort(function(a,b){return b.value - a.value;});
      igCountry.sort(function(a,b){return b.value - a.value;});
      dailyFollowers.sort(function(a,b){return String(a.date).localeCompare(String(b.date));});
      out.audience.ig.age = igAge;
      out.audience.ig.gender = igGender;
      out.audience.ig.city = igCity.slice(0, 10);
      out.audience.ig.country = igCountry.slice(0, 10);
      out.audience.dailyFollowers = dailyFollowers.slice(-30);
    }
  } catch (e) {
    // 忽略
  }
  
  // force JSON 序列化、避免 GAS Date 物件不能傳
  return JSON.parse(JSON.stringify(out));
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

/* ========== 中文自動洞察（核心功能：開後台 5 分鐘看完就知道下一步）========== */
function getInsightsAuto() {
  const ss = SpreadsheetApp.openById(DASH_SS_ID);
  const sh = ss.getSheetByName('Historical_Posts');
  const out = { insights: [], generated_at: Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm') };
  if (!sh || sh.getLastRow() < 2) {
    out.insights.push({ type: 'warning', icon: '⚠️', title: '尚無歷史數據', detail: '先跑 backfillIGHistorical + backfillFBHistorical' });
    return out;
  }
  
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 18).getValues();
  
  // 統計
  const stats = { ig: { count: 0, eng: 0, video: 0, image: 0, video_eng: 0, image_eng: 0 }, fb: { count: 0, eng: 0 } };
  const byHour = {}; // hour => { ig_count, ig_eng, fb_count, fb_eng }
  const byWk = {};
  const byMonth = {};
  const allPosts = [];
  
  data.forEach(function(r){
    const platform = String(r[1] || '');
    const ts = tsToString_(r[3]);
    const type = String(r[4] || '');
    const likes = Number(r[8]) || 0;
    const comments = Number(r[9]) || 0;
    const shares = Number(r[10]) || 0;
    const eng = likes + comments + shares;
    
    let dt;
    try {
      dt = new Date(ts.replace(' ', 'T') + ':00+08:00');
      if (isNaN(dt.getTime())) return;
    } catch(e) { return; }
    
    const hour = dt.getHours();
    const wk = ['日','一','二','三','四','五','六'][dt.getDay()];
    const month = Utilities.formatDate(dt, 'Asia/Taipei', 'yyyy-MM');
    
    allPosts.push({ ts: ts, platform: platform, type: type, eng: eng, hour: hour, wk: wk, month: month, permalink: String(r[5]||''), caption: String(r[17]||'') });
    
    if (platform === 'IG') {
      stats.ig.count++; stats.ig.eng += eng;
      if (type === 'VIDEO' || type === 'REELS') { stats.ig.video++; stats.ig.video_eng += eng; }
      else { stats.ig.image++; stats.ig.image_eng += eng; }
    } else if (platform === 'FB') {
      stats.fb.count++; stats.fb.eng += eng;
    }
    
    if (!byHour[hour]) byHour[hour] = { ig_c:0, ig_e:0, fb_c:0, fb_e:0 };
    if (!byWk[wk]) byWk[wk] = { ig_c:0, ig_e:0, fb_c:0, fb_e:0 };
    if (!byMonth[month]) byMonth[month] = { ig_c:0, ig_e:0, fb_c:0, fb_e:0 };
    
    const key = platform === 'IG' ? 'ig' : 'fb';
    byHour[hour][key+'_c']++; byHour[hour][key+'_e'] += eng;
    byWk[wk][key+'_c']++; byWk[wk][key+'_e'] += eng;
    byMonth[month][key+'_c']++; byMonth[month][key+'_e'] += eng;
  });
  
  // 洞察 1：IG vs FB 戰場差異
  const igAvg = stats.ig.count ? stats.ig.eng / stats.ig.count : 0;
  const fbAvg = stats.fb.count ? stats.fb.eng / stats.fb.count : 0;
  const ratio = fbAvg ? (igAvg / fbAvg).toFixed(1) : '∞';
  out.insights.push({
    type: 'critical', icon: '🎯',
    title: 'IG 是主場、FB 是擺設',
    detail: 'IG 單篇 ' + igAvg.toFixed(0) + ' 互動、FB 單篇 ' + fbAvg.toFixed(1) + ' 互動。IG 強 ' + ratio + ' 倍。',
    action: '5 月起戰力 70% IG / 30% FB。FB 不再衝互動、改成「品牌信任面」。'
  });
  
  // 洞察 2：IG 影片 vs 圖片
  const igVidAvg = stats.ig.video ? stats.ig.video_eng / stats.ig.video : 0;
  const igImgAvg = stats.ig.image ? stats.ig.image_eng / stats.ig.image : 0;
  if (igVidAvg && igImgAvg) {
    const vRatio = (igVidAvg / igImgAvg).toFixed(1);
    out.insights.push({
      type: 'opportunity', icon: '🎬',
      title: 'IG Reels 比圖文強 ' + vRatio + ' 倍',
      detail: 'Reels 單篇 ' + igVidAvg.toFixed(0) + ' 互動、圖文 ' + igImgAvg.toFixed(0) + ' 互動。',
      action: '每週 5 條 Reels + 2 條圖文（不要再 50/50 分配）。'
    });
  }
  
  // 洞察 3：黃金時段（IG）
  const igHours = Object.keys(byHour).map(function(h){
    const d = byHour[h];
    return { h: parseInt(h), avg: d.ig_c ? d.ig_e / d.ig_c : 0, cnt: d.ig_c };
  }).filter(function(x){ return x.cnt >= 3; }).sort(function(a,b){ return b.avg - a.avg; });
  if (igHours.length) {
    const top = igHours[0];
    out.insights.push({
      type: 'opportunity', icon: '⏰',
      title: 'IG 黃金時段：' + (top.h<10?'0':'') + top.h + ':00',
      detail: '此時段平均 ' + top.avg.toFixed(0) + ' 互動（共 ' + top.cnt + ' 篇）。',
      action: '所有 IG Reels 強制排在 ' + top.h + ':00 整點發、不要早不要晚。'
    });
  }
  
  // 洞察 4：黃金週幾（IG）
  const wkOrder = ['一','二','三','四','五','六','日'];
  const igWks = wkOrder.map(function(w){
    const d = byWk[w] || { ig_c:0, ig_e:0 };
    return { w: w, avg: d.ig_c ? d.ig_e / d.ig_c : 0, cnt: d.ig_c };
  }).filter(function(x){ return x.cnt >= 3; }).sort(function(a,b){ return b.avg - a.avg; });
  if (igWks.length >= 2) {
    const top2 = igWks.slice(0,2).map(function(x){ return '週' + x.w + '(' + x.avg.toFixed(0) + ')'; }).join('、');
    const bot = igWks[igWks.length-1];
    out.insights.push({
      type: 'opportunity', icon: '📅',
      title: 'IG 最強週幾：' + top2,
      detail: '最弱：週' + bot.w + '（單篇 ' + bot.avg.toFixed(0) + '）',
      action: '把弱週幾的 IG 全停發、補到強週幾雙發。'
    });
  }
  
  // 洞察 5：FB 月度趨勢警示
  const fbMonths = Object.keys(byMonth).sort();
  if (fbMonths.length >= 3) {
    const recent3 = fbMonths.slice(-3);
    const trend = recent3.map(function(m){
      const d = byMonth[m];
      return { m: m, avg: d.fb_c ? d.fb_e / d.fb_c : 0 };
    });
    if (trend[0].avg > trend[2].avg * 1.5) {
      out.insights.push({
        type: 'warning', icon: '📉',
        title: 'FB 連續 3 月下滑',
        detail: trend.map(function(t){return t.m+'='+t.avg.toFixed(0);}).join(' → '),
        action: 'FB 死亡時段（週二/三/五/六中午）全砍、只發週一/四/日 18-20h 共 3 篇/週。'
      });
    }
  }
  
  // 洞察 6：Top 3 爆紅文公式
  const top3 = allPosts.filter(function(p){ return p.platform === 'IG'; }).sort(function(a,b){ return b.eng - a.eng; }).slice(0, 3);
  if (top3.length === 3) {
    const sample = top3.map(function(p){
      return '互動 ' + p.eng + ' | ' + p.ts.substring(0,10) + ' ' + (p.hour<10?'0':'') + p.hour + 'h | ' + (p.caption ? p.caption.substring(0,30) : '');
    }).join(' / ');
    out.insights.push({
      type: 'opportunity', icon: '🔥',
      title: 'IG Top 3 爆紅文都長這樣',
      detail: sample,
      action: '套這公式：3 秒問題式 hook + 生活感官主題 + 套科學實驗 + 18:00 整點發。'
    });
  }
  
  // 洞察 7：受眾畫像對應
  const ash = ss.getSheetByName('Audience_Snapshot');
  if (ash && ash.getLastRow() > 1) {
    const adata = ash.getRange(2, 1, ash.getLastRow()-1, 6).getValues();
    let female = 0, male = 0, age25_34 = 0, age35_44 = 0, age45_54 = 0;
    adata.forEach(function(r){
      const dim = String(r[3]||'');
      const val = Number(r[4]) || 0;
      const metric = String(r[2]||'');
      if (metric.indexOf('性別') >= 0 || metric.indexOf('gender') >= 0) {
        if (dim === 'F' || dim === 'female') female += val;
        else if (dim === 'M' || dim === 'male') male += val;
      }
      if (metric.indexOf('年齡') >= 0 || metric.indexOf('age') >= 0) {
        if (dim === '25-34') age25_34 += val;
        else if (dim === '35-44') age35_44 += val;
        else if (dim === '45-54') age45_54 += val;
      }
    });
    const total = female + male;
    if (total > 0) {
      const pctF = (female/total*100).toFixed(0);
      out.insights.push({
        type: 'critical', icon: '👩',
        title: '受眾女性佔 ' + pctF + '%、25-34 為主',
        detail: '25-34 歲 ' + age25_34 + '、35-44 歲 ' + age35_44 + '、45-54 歲 ' + age45_54 + '。',
        action: '文案對「25-34 媽媽」說話、不對「孩子」說話。重寫所有 CTA 用「精細檢測 50 分鐘」+「E 小編 1 對 1」。'
      });
    }
  }
  
  return out;
}

/* ========== 未來排程（從 Posting_Queue 撈未來 14 天）========== */
function getFutureSchedule() {
  const sh = SpreadsheetApp.openById(DASH_SS_ID).getSheetByName('排程佇列 Posting_Queue');
  if (!sh || sh.getLastRow() < 2) return [];
  
  const data = sh.getRange(2, 1, sh.getLastRow()-1, 26).getValues();
  const now = new Date();
  const future14 = new Date(now.getTime() + 14*24*60*60*1000);
  const out = [];
  
  data.forEach(function(r, i){
    const date = r[1]; const time = r[2];
    if (!date) return;
    
    let dt;
    if (date instanceof Date) {
      dt = new Date(date);
      if (time && typeof time === 'string') {
        const parts = time.split(':');
        dt.setHours(parseInt(parts[0])||0, parseInt(parts[1])||0);
      }
    } else if (typeof date === 'string' && date) {
      const dStr = String(date) + ' ' + String(time||'00:00');
      dt = new Date(dStr.replace(' ','T') + ':00+08:00');
    }
    if (!dt || isNaN(dt.getTime())) return;
    if (dt < now || dt > future14) return;
    
    const status = String(r[15]||'');
    if (status === '已發' || status === '失敗') return;
    
    out.push({
      _row: i + 2,
      queue_id: String(r[0]||''),
      date: Utilities.formatDate(dt, 'Asia/Taipei', 'MM-dd'),
      time: String(r[2]||'').substring(0,5) || (dt.getHours()<10?'0':'')+dt.getHours()+':00',
      weekday: ['日','一','二','三','四','五','六'][dt.getDay()],
      platform: String(r[3]||''),
      ratio: String(r[4]||''),
      theme: String(r[8]||''),
      headline: String(r[9]||''),
      status: status,
      img_review: String(r[13]||''),
      copy_review: String(r[14]||''),
      thumb: String(r[7]||r[6]||'')
    });
  });
  
  out.sort(function(a,b){ return (a.date+a.time).localeCompare(b.date+b.time); });
  return JSON.parse(JSON.stringify(out));
}

/* ========== 7/30/90 天 KPI 切換 ========== */
function getKpisRange(days) {
  days = parseInt(days) || 30;
  const ss = SpreadsheetApp.openById(DASH_SS_ID);
  const sh = ss.getSheetByName('Historical_Posts');
  const out = { days: days, ig: { posts:0, eng:0, avg:0 }, fb: { posts:0, eng:0, avg:0 }, daily: [] };
  if (!sh || sh.getLastRow() < 2) return out;
  
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);
  
  const data = sh.getRange(2, 1, sh.getLastRow()-1, 18).getValues();
  const dailyMap = {};
  
  data.forEach(function(r){
    const ts = tsToString_(r[3]);
    let dt;
    try {
      dt = new Date(ts.replace(' ','T') + ':00+08:00');
      if (isNaN(dt.getTime())) return;
    } catch(e) { return; }
    if (dt < cutoff) return;
    
    const platform = String(r[1]||'');
    const eng = (Number(r[8])||0) + (Number(r[9])||0) + (Number(r[10])||0);
    const dayKey = Utilities.formatDate(dt, 'Asia/Taipei', 'MM-dd');
    
    if (platform === 'IG') { out.ig.posts++; out.ig.eng += eng; }
    else if (platform === 'FB') { out.fb.posts++; out.fb.eng += eng; }
    
    if (!dailyMap[dayKey]) dailyMap[dayKey] = { ig: 0, fb: 0 };
    if (platform === 'IG') dailyMap[dayKey].ig += eng;
    else if (platform === 'FB') dailyMap[dayKey].fb += eng;
  });
  
  out.ig.avg = out.ig.posts ? Math.round(out.ig.eng / out.ig.posts) : 0;
  out.fb.avg = out.fb.posts ? Math.round(out.fb.eng / out.fb.posts) : 0;
  
  out.daily = Object.keys(dailyMap).sort().map(function(k){
    return { date: k, ig: dailyMap[k].ig, fb: dailyMap[k].fb };
  });
  
  return out;
}

/* ========== Phase 3: 高潛家長名單 ========== */
function getHotLeads() {
  const ss = SpreadsheetApp.openById(DASH_SS_ID);
  const sh = ss.getSheetByName('高潛家長 Hot_Leads');
  if (!sh || sh.getLastRow() < 2) return [];
  const data = sh.getRange(2, 1, sh.getLastRow()-1, 11).getValues();
  return data.map(function(r){
    return {
      user_id: String(r[0]||''),
      user_name: String(r[1]||''),
      platform: String(r[2]||''),
      count: Number(r[3])||0,
      first: r[4] instanceof Date ? Utilities.formatDate(r[4], 'Asia/Taipei', 'MM-dd HH:mm') : String(r[4]||''),
      last: r[5] instanceof Date ? Utilities.formatDate(r[5], 'Asia/Taipei', 'MM-dd HH:mm') : String(r[5]||''),
      tag: String(r[6]||''),
      topic: String(r[7]||''),
      action: String(r[8]||''),
      line_link: String(r[9]||''),
      contacted: String(r[10]||'')
    };
  }).filter(function(x){ return x.user_id; });
}

function refreshHotLeadsFromDashboard() {
  // 觸發 auto_reply_engine_v2 的 refreshHotLeads_
  if (typeof refreshHotLeads_ === 'function') {
    refreshHotLeads_();
    return { ok: true, msg: '已重新掃描' };
  }
  return { ok: false, msg: '請先部署 auto_reply_engine_v2.gs' };
}

/* ========== Phase 4: 完整 6 區塊拆解 ========== */
function getDeepAnalytics() {
  const ss = SpreadsheetApp.openById(DASH_SS_ID);
  const sh = ss.getSheetByName('Historical_Posts');
  const out = {
    generated_at: Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm'),
    block1_platform: null,
    block2_type: null,
    block3_heatmap: null,
    block4_topics: null,
    block5_growth: null,
    block6_frequency: null,
    warnings: []
  };
  if (!sh || sh.getLastRow() < 2) {
    out.warnings.push('Historical_Posts 無資料、請先跑 backfillIGHistorical / backfillFBHistorical');
    return out;
  }

  const data = sh.getRange(2, 1, sh.getLastRow()-1, 18).getValues();
  const posts = [];
  data.forEach(function(r){
    const ts = tsToString_(r[3]);
    let dt;
    try {
      dt = new Date(ts.replace(' ','T') + ':00+08:00');
      if (isNaN(dt.getTime())) return;
    } catch(e) { return; }
    const likes = Number(r[8]) || 0;
    const comments = Number(r[9]) || 0;
    const shares = Number(r[10]) || 0;
    posts.push({
      platform: String(r[1]||''),
      type: String(r[4]||''),
      dt: dt,
      hour: dt.getHours(),
      wd: dt.getDay(),
      ym: Utilities.formatDate(dt, 'Asia/Taipei', 'yyyy-MM'),
      yw: hist_yearWeek_(dt),
      likes: likes, comments: comments, shares: shares,
      eng: likes + comments + shares,
      reach: Number(r[12]) || 0,
      caption: String(r[17]||'')
    });
  });

  if (!posts.length) { out.warnings.push('資料解析後為空'); return out; }

  // 偵測 FB 互動全 0 的問題
  const fb = posts.filter(function(p){ return p.platform === 'FB'; });
  const fbZero = fb.length && fb.every(function(p){ return p.eng === 0; });
  if (fbZero) out.warnings.push('FB ' + fb.length + ' 篇互動全 0、請到 Apps Script 跑 fbBackfillInteractions 補抓');

  /* ===== Block 1: 平台戰力對比 ===== */
  const ig = posts.filter(function(p){ return p.platform === 'IG'; });
  const igAvg = ig.length ? ig.reduce(function(s,p){return s+p.eng;}, 0) / ig.length : 0;
  const fbAvg = fb.length ? fb.reduce(function(s,p){return s+p.eng;}, 0) / fb.length : 0;
  const ratio = fbAvg ? (igAvg / fbAvg) : (igAvg ? 999 : 0);
  // 近 30 天
  const cutoff30 = new Date(); cutoff30.setDate(cutoff30.getDate() - 30);
  const ig30 = ig.filter(function(p){return p.dt >= cutoff30;});
  const fb30 = fb.filter(function(p){return p.dt >= cutoff30;});
  out.block1_platform = {
    ig_total: ig.length, fb_total: fb.length,
    ig_avg: Math.round(igAvg), fb_avg: Math.round(fbAvg),
    ig_eng_total: ig.reduce(function(s,p){return s+p.eng;}, 0),
    fb_eng_total: fb.reduce(function(s,p){return s+p.eng;}, 0),
    ratio: ratio >= 999 ? '∞' : ratio.toFixed(1),
    ig_30d_posts: ig30.length, fb_30d_posts: fb30.length,
    interpretation: ratio >= 5
      ? 'IG 是絕對主場、戰力強 ' + (ratio >= 999 ? '無限' : ratio.toFixed(1)) + ' 倍。FB 平均互動 ' + Math.round(fbAvg) + ' 接近 0、再投時間進去 CP 值極差。'
      : (ratio >= 2 ? 'IG 略強、但 FB 仍有戰力。' : 'FB 表現意外不錯、可保留資源。'),
    action: ratio >= 5
      ? '5 月戰力 70% IG / 30% FB。FB 不衝互動、改用「品牌信任面」（家長分享、活動紀錄）'
      : '5 月戰力 60% IG / 40% FB、保留 FB 經營'
  };

  /* ===== Block 2: 內容類型拆解（IG）===== */
  const igTypes = {};
  ig.forEach(function(p){
    const t = (p.type === 'VIDEO' || p.type === 'REELS') ? 'Reels影片' : (p.type === 'CAROUSEL_ALBUM' ? '輪播' : '圖文');
    if (!igTypes[t]) igTypes[t] = { count: 0, eng: 0, top_eng: 0 };
    igTypes[t].count++; igTypes[t].eng += p.eng;
    if (p.eng > igTypes[t].top_eng) igTypes[t].top_eng = p.eng;
  });
  const typesArr = Object.keys(igTypes).map(function(t){
    return { type: t, count: igTypes[t].count, avg: Math.round(igTypes[t].eng / igTypes[t].count), top: igTypes[t].top_eng };
  }).sort(function(a,b){ return b.avg - a.avg; });

  let typeAction = '無資料';
  if (typesArr.length >= 2) {
    const win = typesArr[0]; const loss = typesArr[typesArr.length-1];
    const winRatio = loss.avg ? (win.avg / loss.avg).toFixed(1) : '∞';
    typeAction = win.type + ' 比 ' + loss.type + ' 強 ' + winRatio + ' 倍。每週 5 條 ' + win.type + ' + 2 條 ' + loss.type + '、不要 50/50 平均分。';
  }
  out.block2_type = { items: typesArr, action: typeAction };

  /* ===== Block 3: 黃金時段熱力圖（IG）===== */
  // 24 小時 × 7 天 = 168 格
  const heat = []; // { wd, hour, count, avg }
  const heatMap = {};
  ig.forEach(function(p){
    const k = p.wd + '_' + p.hour;
    if (!heatMap[k]) heatMap[k] = { wd: p.wd, hour: p.hour, count: 0, eng: 0 };
    heatMap[k].count++; heatMap[k].eng += p.eng;
  });
  Object.keys(heatMap).forEach(function(k){
    const c = heatMap[k];
    heat.push({ wd: c.wd, hour: c.hour, count: c.count, avg: c.count ? Math.round(c.eng/c.count) : 0 });
  });
  // 取 Top 5
  const top5 = heat.filter(function(x){ return x.count >= 2; }).sort(function(a,b){ return b.avg - a.avg; }).slice(0, 5);
  const wkName = ['日','一','二','三','四','五','六'];
  out.block3_heatmap = {
    grid: heat,
    top5: top5.map(function(x){ return { label: '週' + wkName[x.wd] + ' ' + (x.hour<10?'0':'') + x.hour + ':00', avg: x.avg, count: x.count }; }),
    action: top5.length ? '黃金時段 ' + top5.slice(0,2).map(function(x){return '週'+wkName[x.wd]+' '+(x.hour<10?'0':'')+x.hour+'h';}).join('、') + ' 必發 Reels' : '資料不足'
  };

  /* ===== Block 4: 主題效能排名 ===== */
  const topicKeywords = {
    '科學/實驗': ['科學','實驗','化學','物理','生物','觀察','顯微','光學','磁','電'],
    '英文/單字': ['英文','english','單字','發音','phonics','閱讀','文法'],
    '課程/開課': ['課程','開課','梯次','報名','上課','班級'],
    '活動/體驗': ['活動','體驗','講座','派對','派發','親子','闖關','遊戲'],
    '師資/環境': ['老師','師資','教室','分校','環境','設施'],
    '家長/見證': ['家長','分享','回饋','心得','見證','故事','成長']
  };
  const topicStats = {};
  Object.keys(topicKeywords).forEach(function(t){ topicStats[t] = { count: 0, eng: 0 }; });
  topicStats['其他'] = { count: 0, eng: 0 };
  ig.forEach(function(p){
    const cap = (p.caption || '').toLowerCase();
    let matched = false;
    for (const t in topicKeywords) {
      if (topicKeywords[t].some(function(kw){ return cap.indexOf(kw) >= 0; })) {
        topicStats[t].count++; topicStats[t].eng += p.eng; matched = true; break;
      }
    }
    if (!matched) { topicStats['其他'].count++; topicStats['其他'].eng += p.eng; }
  });
  const topicsArr = Object.keys(topicStats).map(function(t){
    const s = topicStats[t];
    return { topic: t, count: s.count, avg: s.count ? Math.round(s.eng/s.count) : 0 };
  }).filter(function(x){ return x.count > 0; }).sort(function(a,b){ return b.avg - a.avg; });
  let topicAction = '無資料';
  if (topicsArr.length >= 2) {
    const w = topicsArr[0]; const l = topicsArr[topicsArr.length-1];
    topicAction = '「' + w.topic + '」單篇平均 ' + w.avg + '、最強。「' + l.topic + '」最弱（' + l.avg + '）。多寫前 3 名、少碰最後 1 名。';
  }
  out.block4_topics = { items: topicsArr, action: topicAction };

  /* ===== Block 5: 成長曲線（近 8 週）===== */
  const wkBuckets = {};
  ig.forEach(function(p){
    if (!wkBuckets[p.yw]) wkBuckets[p.yw] = { count: 0, eng: 0 };
    wkBuckets[p.yw].count++; wkBuckets[p.yw].eng += p.eng;
  });
  const sortedWks = Object.keys(wkBuckets).sort();
  const last8 = sortedWks.slice(-8).map(function(w){
    return { week: w, posts: wkBuckets[w].count, eng: wkBuckets[w].eng, avg: Math.round(wkBuckets[w].eng / wkBuckets[w].count) };
  });
  // 預測下週（線性回歸：簡化為近 4 週平均 + 趨勢）
  let predict = 0; let trend = '持平';
  if (last8.length >= 4) {
    const recent4 = last8.slice(-4);
    const avg4 = recent4.reduce(function(s,x){return s+x.avg;}, 0) / 4;
    const slope = (recent4[3].avg - recent4[0].avg) / 3;
    predict = Math.max(0, Math.round(avg4 + slope));
    if (slope > avg4 * 0.1) trend = '上升';
    else if (slope < -avg4 * 0.1) trend = '下滑';
  }
  out.block5_growth = {
    weeks: last8, predict_next: predict, trend: trend,
    action: trend === '上升' ? '近 4 週呈上升、預估下週單篇平均 ' + predict + '。維持節奏、加碼 Reels。' :
            trend === '下滑' ? '⚠️ 近 4 週呈下滑、預估下週 ' + predict + '。改 hook 公式、換主題、增加 Reels 比例。' :
            '近 4 週持平、預估 ' + predict + '。試新題材推一波。'
  };

  /* ===== Block 6: 發文頻率 vs 互動 ===== */
  // 把 ig 依週分組、看「該週發文數」對應「平均互動」
  const freqCorr = sortedWks.slice(-12).map(function(w){
    return { week: w, posts_per_week: wkBuckets[w].count, avg_eng: Math.round(wkBuckets[w].eng / wkBuckets[w].count) };
  });
  // 找最佳節奏
  const byPosts = {};
  freqCorr.forEach(function(x){
    const bucket = x.posts_per_week >= 7 ? '7+' : (x.posts_per_week >= 5 ? '5-6' : (x.posts_per_week >= 3 ? '3-4' : '1-2'));
    if (!byPosts[bucket]) byPosts[bucket] = { weeks: 0, total_avg: 0 };
    byPosts[bucket].weeks++; byPosts[bucket].total_avg += x.avg_eng;
  });
  const buckets = Object.keys(byPosts).map(function(b){
    return { bucket: b + '篇/週', weeks: byPosts[b].weeks, avg: Math.round(byPosts[b].total_avg / byPosts[b].weeks) };
  }).sort(function(a,b){ return b.avg - a.avg; });
  let freqAction = '無資料';
  if (buckets.length >= 2) {
    freqAction = '最佳節奏：' + buckets[0].bucket + '（單篇平均 ' + buckets[0].avg + '）、避免 ' + buckets[buckets.length-1].bucket + '（' + buckets[buckets.length-1].avg + '）';
  }
  out.block6_frequency = { items: buckets, recent: freqCorr, action: freqAction };

  return out;
}

function hist_yearWeek_(dt) {
  const d = new Date(dt);
  d.setHours(0,0,0,0);
  d.setDate(d.getDate() + 4 - (d.getDay()||7));
  const yearStart = new Date(d.getFullYear(),0,1);
  const wn = Math.ceil((((d - yearStart) / 86400000) + 1)/7);
  return d.getFullYear() + '-W' + (wn<10?'0':'') + wn;
}

/**
 * 重新抓取 Top N 高互動貼文的縮圖（IG/FB CDN URL 約 24h 失效）
 * 從 Historical_Posts 找互動 = 讚+留言+分享+存 最高的前 limit 筆、
 * 打 Graph API 拿最新 media_url / full_picture、寫回 G 欄
 *
 * @param {number} limit 預設 50（Top 10 + 緩衝）
 * @return {object} { ok, total, updated, failed, skipped, log }
 */
function refreshTopPostThumbs(limit) {
  limit = limit || 50;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Historical_Posts');
  if (!sh) return { ok: false, error: 'Historical_Posts 不存在' };

  // 讀 token
  const settings = ss.getSheetByName('Settings');
  let token = '';
  if (settings) {
    const sd = settings.getRange(1,1,settings.getLastRow(),2).getValues();
    for (const r of sd) {
      if (String(r[0]) === 'FB_PAGE_TOKEN') { token = String(r[1] || ''); break; }
    }
  }
  if (!token) return { ok: false, error: 'FB_PAGE_TOKEN 未設' };

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: true, total: 0, updated: 0, log: '無資料' };

  // 抓 A:R（18 欄）所有資料
  const data = sh.getRange(2, 1, lastRow - 1, 18).getValues();
  // 計算每列互動、排序找 Top
  const indexed = data.map(function(r, i){
    const eng = (Number(r[8])||0) + (Number(r[9])||0) + (Number(r[10])||0) + (Number(r[11])||0);
    return { rowIdx: i + 2, platform: r[1], pid: String(r[2]||''), thumb: String(r[6]||''), eng: eng };
  });
  indexed.sort(function(a,b){ return b.eng - a.eng; });
  const tops = indexed.slice(0, limit);

  let updated = 0, failed = 0, skipped = 0;
  const logs = [];
  const apiBase = 'https' + '://graph.facebook.com/v19.0/';

  for (let i = 0; i < tops.length; i++) {
    const t = tops[i];
    if (!t.pid) { skipped++; continue; }
    const pid = t.pid.replace(/^IG:/, '').replace(/^FB:/, '');

    let newThumb = '';
    try {
      if (t.platform === 'IG') {
        // IG: 拿 media_url + thumbnail_url（影片用 thumbnail_url、圖片用 media_url）
        const url = apiBase + pid + '?fields=' + encodeURIComponent('media_type,media_url,thumbnail_url') + '&access_token=' + token;
        const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        if (resp.getResponseCode() === 200) {
          const j = JSON.parse(resp.getContentText());
          // 影片優先 thumbnail_url（圖片才用 media_url）
          if (j.media_type === 'VIDEO' || j.media_type === 'REELS') {
            newThumb = j.thumbnail_url || j.media_url || '';
          } else {
            newThumb = j.media_url || j.thumbnail_url || '';
          }
        } else {
          logs.push('IG ' + pid + ' HTTP ' + resp.getResponseCode());
        }
      } else if (t.platform === 'FB') {
        // FB: full_picture
        const url = apiBase + pid + '?fields=full_picture&access_token=' + token;
        const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        if (resp.getResponseCode() === 200) {
          const j = JSON.parse(resp.getContentText());
          newThumb = j.full_picture || '';
        } else {
          logs.push('FB ' + pid + ' HTTP ' + resp.getResponseCode());
        }
      } else {
        skipped++;
        continue;
      }

      if (newThumb && newThumb !== t.thumb) {
        sh.getRange(t.rowIdx, 7).setValue(newThumb); // G 欄
        updated++;
      } else if (!newThumb) {
        failed++;
      } else {
        skipped++; // 沒變
      }
    } catch (e) {
      failed++;
      logs.push(t.platform + ' ' + pid + ' err: ' + String(e));
    }

    // API rate limit 緩衝
    if (i % 10 === 9) Utilities.sleep(500);
  }

  const result = {
    ok: true,
    total: tops.length,
    updated: updated,
    failed: failed,
    skipped: skipped,
    log: logs.slice(0, 10).join(' | ') || '完成'
  };
  Logger.log(JSON.stringify(result));
  return result;
}
