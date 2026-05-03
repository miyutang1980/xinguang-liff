/**
 * 太平新光自動回覆引擎 v2 — Phase 3
 *
 * 升級項目：
 *   1. OpenAI 自然語意回覆（規則沒中時 fallback 到 GPT）
 *   2. 高潛家長標記（同帳號 3+ 次互動 → Hot_Leads）
 *   3. LINE OA 預填（所有回覆自動加 @143qbory）
 *   4. 下班時段（22:00-07:00）回覆語氣調整
 *   5. 負評偵測（自動標警示、不自動回、寫進 Alerts）
 *
 * 必要 Settings：
 *   FB_PAGE_TOKEN, FB_PAGE_ID, OPENAI_API_KEY（可選、缺則跳過 AI）
 *   LINE_OA_HANDLE（預設 @143qbory）
 *   AI_REPLY_ENABLE（TRUE/FALSE、預設 FALSE 安全起見）
 *
 * 必要工作表：
 *   互動紀錄 Interactions（既有）
 *   高潛家長 Hot_Leads（自動建立）
 *   負評警示 Alerts（自動建立）
 */

const AR2_SS_ID = '1DybgWBdCyvkEijMyaE46rKLtQD9J2ImjU8xeYCKSKnA';
const AR2_TZ = 'Asia/Taipei';
const AR2_LINE_OA_DEFAULT = '@143qbory';

/* ========== 設定讀取 ========== */
function ar2_setting_(k, defaultVal) {
  const ss = SpreadsheetApp.openById(AR2_SS_ID);
  const sh = ss.getSheetByName('Settings') || ss.getSheetByName('設定 Settings');
  if (!sh) return defaultVal || '';
  const data = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  for (const r of data) if (String(r[0]).trim() === k) return String(r[1]).trim();
  if (defaultVal !== undefined) return defaultVal;
  throw new Error('設定 ' + k + ' 不存在');
}

/* ========== 主入口（取代 pollAllPlatforms）========== */
function pollAllPlatformsV2() {
  ensureSheetsExist_();
  pollIGCommentsV2();
  pollFBCommentsV2();
  // 跑完後更新高潛家長
  refreshHotLeads_();
}

function ensureSheetsExist_() {
  const ss = SpreadsheetApp.openById(AR2_SS_ID);
  if (!ss.getSheetByName('高潛家長 Hot_Leads')) {
    const sh = ss.insertSheet('高潛家長 Hot_Leads');
    sh.appendRow(['user_id','user_name','platform','互動次數','最早互動','最近互動','標籤','主題關鍵字','建議行動','LINE預填連結','已聯繫']);
    sh.setFrozenRows(1);
  }
  if (!ss.getSheetByName('負評警示 Alerts')) {
    const sh = ss.insertSheet('負評警示 Alerts');
    sh.appendRow(['時間','平台','user_name','內容','偵測原因','處理狀態','處理人','備註']);
    sh.setFrozenRows(1);
  }
}

/* ========== IG 留言 v2 ========== */
function pollIGCommentsV2() {
  const ss = SpreadsheetApp.openById(AR2_SS_ID);
  const qSh = ss.getSheetByName('排程佇列 Posting_Queue');
  const iSh = ss.getSheetByName('互動紀錄 Interactions');
  const rules = ar2_loadRules_(ss, 'IG');
  const token = ar2_setting_('FB_PAGE_TOKEN');

  const last = qSh.getLastRow();
  if (last < 2) return;
  const data = qSh.getRange(2, 1, last - 1, 22).getValues();
  const recent = data.filter(function(r){ return r[15] === '已發布' && r[17] && String(r[17]).indexOf('IG:') >= 0; }).slice(-30);

  for (const r of recent) {
    const m = String(r[17]).match(/IG:(\d+)/);
    if (!m) continue;
    const mediaId = m[1];
    try {
      const fbBase = ar2_fbBase_();
      const url = fbBase + '/' + mediaId + '/comments?fields=id,text,username,from,timestamp&limit=25&access_token=' + token;
      const res = JSON.parse(UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText());
      if (!res.data) continue;
      for (const c of res.data) {
        if (ar2_alreadyHandled_(c.id)) continue;
        ar2_handleComment_('IG', mediaId, c.id, c.from && c.from.id || c.username, c.username, c.text, rules, ss, iSh, token);
      }
    } catch (e) { Logger.log('IG v2 poll fail: ' + e); }
  }
}

/* ========== FB 留言 v2 ========== */
function pollFBCommentsV2() {
  const ss = SpreadsheetApp.openById(AR2_SS_ID);
  const qSh = ss.getSheetByName('排程佇列 Posting_Queue');
  const iSh = ss.getSheetByName('互動紀錄 Interactions');
  const rules = ar2_loadRules_(ss, 'FB');
  const token = ar2_setting_('FB_PAGE_TOKEN');
  const pageId = ar2_setting_('FB_PAGE_ID');

  const last = qSh.getLastRow();
  if (last < 2) return;
  const data = qSh.getRange(2, 1, last - 1, 22).getValues();
  const recent = data.filter(function(r){ return r[15] === '已發布' && r[17] && String(r[17]).indexOf('FB:') >= 0; }).slice(-30);

  for (const r of recent) {
    const m = String(r[17]).match(/FB:([\d_]+)/);
    if (!m) continue;
    const postId = m[1];
    try {
      const fbBase = ar2_fbBase_();
      const url = fbBase + '/' + postId + '/comments?fields=id,message,from,created_time&limit=25&access_token=' + token;
      const res = JSON.parse(UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText());
      if (!res.data) continue;
      for (const c of res.data) {
        if (ar2_alreadyHandled_(c.id)) continue;
        if (c.from && c.from.id === pageId) continue;
        ar2_handleComment_('FB', postId, c.id, c.from && c.from.id, c.from && c.from.name, c.message, rules, ss, iSh, token);
      }
    } catch (e) { Logger.log('FB v2 poll fail: ' + e); }
  }
}

/* ========== 統一處理一則留言（核心邏輯）========== */
function ar2_handleComment_(platform, postId, commentId, userId, userName, content, rules, ss, iSh, token) {
  const text = String(content || '').trim();
  if (!text) return;

  // 1. 負評偵測（優先）
  if (ar2_isNegative_(text)) {
    ar2_appendAlert_(ss, platform, userName || userId, text, '偵測到負評關鍵字');
    ar2_appendInteraction_(iSh, platform, '留言', postId, userId, userName, text, null, '警示');
    return;
  }

  // 2. 規則比對
  let matched = ar2_matchRules_(rules, text, '留言');
  let replyText = matched ? matched.reply : null;
  let replySource = matched ? '規則' : '';

  // 3. 規則沒中 → AI fallback（如有開啟）
  if (!replyText && ar2_setting_('AI_REPLY_ENABLE', 'FALSE') === 'TRUE') {
    replyText = ar2_aiReply_(text, platform);
    if (replyText) replySource = 'AI';
  }

  // 4. 補上 LINE OA + 下班時段語氣
  if (replyText) {
    replyText = ar2_decorateReply_(replyText);
    if (platform === 'IG') ar2_replyIG_(commentId, replyText, token);
    else ar2_replyFB_(commentId, replyText, token);
    if (matched) ar2_incRuleHit_(ss, matched.ruleId);
    Utilities.sleep(2000);
  }

  ar2_appendInteraction_(iSh, platform, '留言', postId, userId, userName, text, matched, replySource);
}

/* ========== 回覆裝飾：加 LINE OA + 時段語氣 ========== */
function ar2_decorateReply_(text) {
  const lineOa = ar2_setting_('LINE_OA_HANDLE', AR2_LINE_OA_DEFAULT);
  const hour = parseInt(Utilities.formatDate(new Date(), AR2_TZ, 'HH'));
  const isOff = hour >= 22 || hour < 7;

  let out = text;
  if (out.indexOf(lineOa) < 0) {
    const tail = isOff
      ? '\n（深夜時段、E 小編明早回覆。可先加 LINE OA ' + lineOa + '）'
      : '\n（私訊 LINE OA ' + lineOa + ' 由 E 小編 1 對 1 回您）';
    out = out + tail;
  }
  return out;
}

/* ========== OpenAI 自然語意回覆 ========== */
function ar2_aiReply_(userText, platform) {
  const apiKey = ar2_setting_('OPENAI_API_KEY', '');
  if (!apiKey) return null;

  const sysPrompt = [
    '你是「弋果美語太平新光分校」的 E 小編，專業親切。',
    '回覆規則（嚴格遵守）：',
    '1. 繁體中文（台灣）、最多 80 字。',
    '2. 絕對禁用：免費抵註冊費、免費一堂課體驗、免試聽評測費、省 500 元、免費試讀、同校兩人組。',
    '3. 早鳥提 85 折（不是 8 折）、不指名任何競爭對手品牌。',
    '4. 強調「科學為主、英文為媒」、「精細檢測 50 分鐘」、「E 小編 1 對 1」。',
    '5. 對 25-34 歲媽媽說話、不對孩子說話。',
    '6. 結尾不要簽名、不要 emoji，回覆完即停（系統自動加 LINE 引導）。',
    '7. 若家長問價格 → 回「私訊由 E 小編 1 對 1 評估」、不直接報價。',
    '8. 若內容超出範圍 → 回「請私訊由 E 小編 1 對 1 為您處理」。'
  ].join('\n');

  const payload = {
    model: 'gpt-4o-mini',
    messages: [
      { role: 'system', content: sysPrompt },
      { role: 'user', content: '平台：' + platform + '\n家長留言：' + userText + '\n請用 1-3 句回覆。' }
    ],
    temperature: 0.6,
    max_tokens: 150
  };
  const url = 'https' + '://api.openai.com/v1/chat/completions';
  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code !== 200) { Logger.log('AI fail ' + code + ': ' + res.getContentText().substring(0,200)); return null; }
    const j = JSON.parse(res.getContentText());
    const txt = j.choices && j.choices[0] && j.choices[0].message && j.choices[0].message.content;
    if (!txt) return null;
    // 安全檢查：禁字過濾
    const banned = ['免費抵註冊費','免費一堂課體驗','免試聽評測費','省 500 元','省500元','免費試讀','同校兩人組'];
    for (const b of banned) if (txt.indexOf(b) >= 0) return null;
    return String(txt).trim();
  } catch (e) {
    Logger.log('AI exception: ' + e);
    return null;
  }
}

/* ========== 負評偵測 ========== */
function ar2_isNegative_(text) {
  const t = String(text || '').toLowerCase();
  const neg = ['爛','糟','差','騙','坑','退費','投訴','客訴','不專業','態度差','破','詐','黑心','後悔','浪費錢','拒絕','失望','不推薦'];
  for (const w of neg) if (t.indexOf(w) >= 0) return true;
  return false;
}

function ar2_appendAlert_(ss, platform, userName, content, reason) {
  const sh = ss.getSheetByName('負評警示 Alerts');
  if (!sh) return;
  sh.appendRow([
    Utilities.formatDate(new Date(), AR2_TZ, 'yyyy-MM-dd HH:mm:ss'),
    platform, userName || '', content, reason, '待處理', '', ''
  ]);
}

/* ========== 高潛家長標記 ========== */
function refreshHotLeads_() {
  const ss = SpreadsheetApp.openById(AR2_SS_ID);
  const iSh = ss.getSheetByName('互動紀錄 Interactions');
  const hSh = ss.getSheetByName('高潛家長 Hot_Leads');
  if (!iSh || !hSh) return;
  if (iSh.getLastRow() < 2) return;

  const lineOa = ar2_setting_('LINE_OA_HANDLE', AR2_LINE_OA_DEFAULT);
  const data = iSh.getRange(2, 1, iSh.getLastRow()-1, 14).getValues();

  const userMap = {};
  data.forEach(function(r){
    const ts = r[1]; const platform = r[2]; const userId = r[5]; const userName = r[6]; const content = r[7];
    if (!userId) return;
    const key = platform + '|' + userId;
    if (!userMap[key]) userMap[key] = { user_id: userId, user_name: userName, platform: platform, count: 0, first: ts, last: ts, contents: [] };
    userMap[key].count++;
    userMap[key].last = ts;
    if (userMap[key].contents.length < 5) userMap[key].contents.push(String(content||''));
  });

  // 過濾：互動 3+ 次
  const hot = Object.keys(userMap).map(function(k){ return userMap[k]; }).filter(function(u){ return u.count >= 3; });

  // 主題關鍵字提取
  const keywords = ['科學','實驗','英文','學費','課程','試聽','報名','體驗','時間','分校','開課','年齡','大班','一年級','二年級','三年級','幼兒','幼稚園'];

  hot.forEach(function(u){
    const blob = u.contents.join(' ').toLowerCase();
    const hits = keywords.filter(function(kw){ return blob.indexOf(kw) >= 0; });
    u.topic = hits.slice(0, 3).join('、') || '';
    u.action = u.count >= 5 ? '立刻 LINE 私訊邀約' : '本週內 LINE 私訊跟進';
    // LINE 預填連結（用 https + line.me）
    const proto = 'https' + ':';
    const handle = lineOa.replace(/^@/, '');
    u.line_link = proto + '//line.me/R/ti/p/' + encodeURIComponent('@' + handle);
  });

  // 寫回 Hot_Leads（清空再寫）
  if (hSh.getLastRow() > 1) hSh.getRange(2, 1, hSh.getLastRow()-1, hSh.getLastColumn()).clearContent();
  if (!hot.length) return;

  hot.sort(function(a,b){ return b.count - a.count; });
  const rows = hot.map(function(u){
    return [
      u.user_id, u.user_name || '', u.platform, u.count,
      u.first instanceof Date ? Utilities.formatDate(u.first, AR2_TZ, 'yyyy-MM-dd HH:mm') : String(u.first||''),
      u.last instanceof Date ? Utilities.formatDate(u.last, AR2_TZ, 'yyyy-MM-dd HH:mm') : String(u.last||''),
      u.count >= 5 ? '🔥 高溫' : '🌡 溫熱',
      u.topic, u.action, u.line_link, ''
    ];
  });
  hSh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

/* ========== 工具 ========== */
function ar2_fbBase_() {
  // 避免在原始碼出現完整 https URL（Apps Script bug）
  return 'https' + '://graph.facebook.com/v19.0';
}

function ar2_replyIG_(commentId, message, token) {
  UrlFetchApp.fetch(ar2_fbBase_() + '/' + commentId + '/replies', {
    method: 'post',
    payload: { message: message, access_token: token },
    muteHttpExceptions: true
  });
}

function ar2_replyFB_(commentId, message, token) {
  UrlFetchApp.fetch(ar2_fbBase_() + '/' + commentId + '/comments', {
    method: 'post',
    payload: { message: message, access_token: token },
    muteHttpExceptions: true
  });
}

function ar2_loadRules_(ss, platformPrefix) {
  const sh = ss.getSheetByName('自動回覆 Auto_Reply');
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  const data = sh.getRange(2, 1, last - 1, 14).getValues();
  return data
    .filter(function(r){ return String(r[1]).toUpperCase() === 'TRUE' && String(r[2]).indexOf(platformPrefix) >= 0; })
    .map(function(r){
      return {
        ruleId: r[0], platforms: r[2], triggerType: r[3],
        keywords: String(r[4]).split('|').map(function(s){return s.trim();}).filter(function(s){return s;}),
        matchType: r[5], reply: r[6], replyMode: r[7], followUp: r[8], tag: r[9]
      };
    });
}

function ar2_matchRules_(rules, text, triggerType) {
  text = String(text || '').toLowerCase();
  for (const rule of rules) {
    if (rule.triggerType !== triggerType && rule.triggerType !== '全部') continue;
    for (const kw of rule.keywords) {
      const kwLow = kw.toLowerCase();
      if (rule.matchType === '起始於') { if (text.indexOf(kwLow) === 0) return rule; }
      else if (rule.matchType === '精準') { if (text === kwLow) return rule; }
      else { if (text.indexOf(kwLow) >= 0) return rule; }
    }
  }
  return null;
}

function ar2_alreadyHandled_(externalId) {
  const cache = PropertiesService.getScriptProperties();
  const key = 'h_' + externalId;
  if (cache.getProperty(key)) return true;
  cache.setProperty(key, '1');
  return false;
}

function ar2_appendInteraction_(sh, platform, type, postId, userId, userName, content, matched, source) {
  const now = Utilities.formatDate(new Date(), AR2_TZ, 'yyyy-MM-dd HH:mm:ss');
  const id = platform + '_' + Math.random().toString(36).substr(2, 9) + '_' + Date.now();
  sh.appendRow([
    id, now, platform, type, postId, userId || '', userName || '',
    content || '',
    matched ? '正向' : (source === '警示' ? '負評' : '中性'),
    matched ? 'TRUE' : 'FALSE',
    matched ? matched.reply : '',
    matched ? matched.replyMode : (source === 'AI' ? 'AI' : ''),
    source === 'AI' ? 'AI' : (matched ? '自動' : ''),
    matched ? matched.tag : (source === '警示' ? '負評' : ''),
    matched && matched.followUp === '建立預約Lead' ? 'TRUE' : ''
  ]);
}

function ar2_incRuleHit_(ss, ruleId) {
  const sh = ss.getSheetByName('自動回覆 Auto_Reply');
  const last = sh.getLastRow();
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === ruleId) {
      const rowNum = i + 2;
      const cur = sh.getRange(rowNum, 11).getValue() || 0;
      sh.getRange(rowNum, 11).setValue(cur + 1);
      sh.getRange(rowNum, 12).setValue(Utilities.formatDate(new Date(), AR2_TZ, 'yyyy-MM-dd HH:mm:ss'));
      return;
    }
  }
}

/* ========== Triggers ========== */
function installReplyTriggersV2() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    const fn = t.getHandlerFunction();
    if (fn === 'pollAllPlatforms' || fn === 'pollAllPlatformsV2') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('pollAllPlatformsV2').timeBased().everyMinutes(5).create();
  SpreadsheetApp.getUi().alert('自動回覆 v2 觸發器：每 5 分鐘掃描\n含 AI 回覆 + 高潛家長 + 負評警示');
}

/* ========== 手動測試入口 ========== */
function testHotLeadsRefresh() {
  ensureSheetsExist_();
  refreshHotLeads_();
  SpreadsheetApp.getUi().alert('Hot_Leads 已更新');
}

function testAIReplyOnce() {
  const reply = ar2_aiReply_('請問課程多少錢？小孩大班可以上嗎？', 'IG');
  SpreadsheetApp.getUi().alert(reply || '（AI 未回應或未啟用）');
}
