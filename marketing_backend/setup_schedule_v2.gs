/**
 * 太平新光 — 完整排程 v2
 * 依照 auto_posting_schedule_419_630.md 與 content_calendar_424_828.md 重灌 Posting_Queue
 *
 * 三階段：
 *   Phase 1 暖身：4/19-5/10（22 天，每天 1 篇 1x1，狀態=草稿）
 *   Phase 2 七月梯預告：5/11-6/8（用滿 D03-D31 共 29 天，每天 1 篇 1x1 + 週末加 9x16）
 *   Phase 3 八月梯預告：6/9-6/30 + 7/1-7/5 補播（用滿 D01-D28，每天 1 篇 1x1 + 週末加 9x16）
 *
 * 暑期當期 7/6-8/28：等實體課拍照後再灌
 *
 * 用法：
 *   1. 貼進 Apps Script
 *   2. 執行 setupScheduleV2()
 *   3. 提示「會清空目前 68 列、灌入 130+ 列」按確認
 */

const V2_SS_ID = '1DybgWBdCyvkEijMyaE46rKLtQD9J2ImjU8xeYCKSKnA';
const V2_QUEUE = '排程佇列 Posting_Queue';

// ============ Phase 1：4/19-5/10 暖身（22 篇）============
const P1_THEMES = [
  ['2026-04-19','日','開站公告','2026 夏季科學探索營正式上線'],
  ['2026-04-20','一','七月梯介紹','Science Through Discovery 超能科學派'],
  ['2026-04-21','二','八月梯介紹','Mystery Hunters 謎案追查隊'],
  ['2026-04-22','三','CLIL 概念','為什麼要用英文學科學？'],
  ['2026-04-23','四','校園環境','教室實體照'],
  ['2026-04-24','五','第一支 Reels','兩梯精彩預告'],
  ['2026-04-25','六','早鳥優惠倒數','5/31 前報名享 85 折'],
  ['2026-04-26','日','七月梯安親','課輔專線開跑'],
  ['2026-04-27','一','外師團隊','背影側影介紹'],
  ['2026-04-28','二','一天作息','課程一天作息表'],
  ['2026-04-29','三','寫作 W01','教材寫作 + Journaling W01'],
  ['2026-04-30','四','GEPT 路徑','全民英檢路徑介紹'],
  ['2026-05-01','五','FAQ 費用','家長常見問題：費用'],
  ['2026-05-02','六','FAQ 接送','家長常見問題：時段與接送'],
  ['2026-05-03','日','Reading W01','教材 Reading 每週主題'],
  ['2026-05-04','一','為什麼選弋果','夏令營理由 5 點'],
  ['2026-05-05','二','物理教具開箱','實驗教具：物理篇'],
  ['2026-05-06','三','科學 W02','浸潤班科學課程 W02'],
  ['2026-05-07','四','偵探教具開箱','實驗教具：偵探篇'],
  ['2026-05-08','五','Starters 介紹','劍橋 Starters 是什麼'],
  ['2026-05-09','六','家長回饋一','去年學員家長見證'],
  ['2026-05-10','日','兒美 W02','兒美班課程介紹 W02'],
];

// ============ Phase 2：5/11-6/8 七月梯預告（29 天 D03-D31）============
// 七月課程主題（D03-D31）
const JUL_TOPICS_V2 = {
  d03: '動量與碰撞', d04: '彈性與形變', d05: '波與震動', d06: '透鏡與光路',
  d07: '反射與成像', d08: '酸鹼指示劑', d09: '溶解與飽和', d10: '結晶與析出',
  d11: '氧化與還原', d12: '分離過濾',   d13: '蒸發凝結',  d14: '燃燒火焰',
  d15: '指紋證物',   d16: '骨骼關節',   d17: '感官反應',  d18: '心臟血液',
  d19: '昆蟲變態',   d20: '光合作用',   d21: 'DNA 遺傳',  d22: '地震板塊',
  d23: '唇紋筆跡',   d24: '足跡分析',   d25: '時間估算',  d26: '骨骼鑑定',
  d27: '粉末檢驗',   d28: '結案發表',   d29: '畢業典禮',  d30: '證物展',
  d31: '畢業典禮 II',
};

// ============ Phase 3：6/9-7/5 八月梯預告（28 天 D01-D28）============
const AUG_TOPICS_V2 = {
  d01: '拉力與張力', d02: '扭力與旋轉', d03: '動量碰撞',  d04: '彈性形變',
  d05: '波與震動',   d06: '透鏡光路',   d07: '反射成像',  d08: '酸鹼指示',
  d09: '溶解飽和',   d10: '結晶析出',   d11: '氧化還原',  d12: '分離過濾',
  d13: '蒸發凝結',   d14: '燃燒火焰',   d15: '指紋證物',  d16: '骨骼關節',
  d17: '感官反應',   d18: '心臟血液',   d19: '昆蟲變態',  d20: '光合作用',
  d21: 'DNA 遺傳',   d22: '地震板塊',   d23: '唇紋筆跡',  d24: '足跡分析',
  d25: '時間估算',   d26: '骨骼鑑定',   d27: '粉末檢驗',  d28: '結案發表',
};

// ============ 主流程 ============
function setupScheduleV2() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(V2_SS_ID);
  const sh = ss.getSheetByName(V2_QUEUE);
  if (!sh) throw new Error('找不到分頁：' + V2_QUEUE);

  const last = sh.getLastRow();
  const existing = Math.max(0, last - 1);

  const ans = ui.alert(
    '⚠️ 完整重灌 Posting_Queue',
    `將清空目前 ${existing} 列、灌入新 ${22 + 29 * 2 + 28 * 2} 列：\n\n` +
    `Phase 1 暖身：22 列（4/19-5/10）\n` +
    `Phase 2 七月預告：58 列（5/11-6/8 D03-D31，1x1+9x16）\n` +
    `Phase 3 八月預告：56 列（6/9-7/5 D01-D28，1x1+9x16）\n\n` +
    `合計約 136 列\n\n按 OK 確認、Cancel 取消`,
    ui.ButtonSet.OK_CANCEL
  );
  if (ans !== ui.Button.OK) return;

  // 清空舊資料（保留 row 1 表頭）
  if (last >= 2) sh.getRange(2, 1, last - 1, 22).clearContent();

  // 讀取素材對照（從 schedule_v2_assets.json 內嵌進來）
  const ASSETS = getAssets_();

  const rows = [];
  const now = new Date();
  let qSeq = 1;

  // ===== Phase 1 (22 篇 1x1) =====
  P1_THEMES.forEach((row) => {
    const [date, , topic, sub] = row;
    rows.push(buildRow_({
      qid: 'P1_' + String(qSeq).padStart(3, '0'),
      date: date,
      time: '12:30',
      platform: 'IG+FB',
      ratio: '1x1',
      asset_id: 'P1_' + topic.replace(/\s+/g, '_'),
      img_url: '', // Phase 1 沒對應圖、留空待補
      thumb_url: '',
      topic: topic,
      sub: sub,
      now: now,
      stage: 'Phase 1 暖身',
    }));
    qSeq++;
  });

  // ===== Phase 2 七月預告：5/11-6/8（29 天，D03-D31）=====
  // 每天 1x1 + 週六/週日 加發 9x16
  let day = new Date(2026, 4, 11); // 5/11
  const julKeys = Object.keys(JUL_TOPICS_V2).sort(); // d03-d31
  julKeys.forEach((dKey) => {
    const dateStr = fmtDate_(day);
    const wd = day.getDay(); // 0=Sun 6=Sat
    const topic = JUL_TOPICS_V2[dKey];

    // 1x1 必發
    const a1 = ASSETS.jul[dKey + '_1x1'];
    rows.push(buildRow_({
      qid: 'JUL_' + dKey.toUpperCase() + '_1x1',
      date: dateStr,
      time: '12:30',
      platform: 'IG+FB',
      ratio: '1x1',
      asset_id: 'JUL_' + dKey.toUpperCase() + '_1x1',
      img_url: a1 ? `https://drive.google.com/file/d/${a1.file_id}/view?usp=drivesdk` : '',
      thumb_url: a1 ? `https://drive.google.com/thumbnail?id=${a1.file_id}&sz=w400` : '',
      topic: '七月預告 ' + dKey.toUpperCase() + ' ' + topic,
      sub: '七月夏令營預告',
      now: now,
      stage: 'Phase 2 七月預告',
    }));

    // 週末多發 9x16（Reels）
    if (wd === 0 || wd === 6) {
      const a916 = ASSETS.jul[dKey + '_9x16'];
      rows.push(buildRow_({
        qid: 'JUL_' + dKey.toUpperCase() + '_9x16',
        date: dateStr,
        time: '17:00',
        platform: 'IG Reels+Stories',
        ratio: '9x16',
        asset_id: 'JUL_' + dKey.toUpperCase() + '_9x16',
        img_url: a916 ? `https://drive.google.com/file/d/${a916.file_id}/view?usp=drivesdk` : '',
        thumb_url: a916 ? `https://drive.google.com/thumbnail?id=${a916.file_id}&sz=w400` : '',
        topic: '七月預告 ' + dKey.toUpperCase() + ' ' + topic,
        sub: '七月夏令營預告 Reels',
        now: now,
        stage: 'Phase 2 七月預告',
      }));
    }
    day.setDate(day.getDate() + 1);
  });

  // ===== Phase 3 八月預告：6/9-7/6（D01-D28，每天 1x1 + 週末 9x16）=====
  day = new Date(2026, 5, 9); // 6/9
  const augKeys = Object.keys(AUG_TOPICS_V2).sort(); // d01-d28
  augKeys.forEach((dKey) => {
    const dateStr = fmtDate_(day);
    const wd = day.getDay();
    const topic = AUG_TOPICS_V2[dKey];

    const a1 = ASSETS.aug[dKey + '_1x1'];
    rows.push(buildRow_({
      qid: 'AUG_' + dKey.toUpperCase() + '_1x1',
      date: dateStr,
      time: '12:30',
      platform: 'IG+FB',
      ratio: '1x1',
      asset_id: 'AUG_' + dKey.toUpperCase() + '_1x1',
      img_url: a1 ? `https://drive.google.com/file/d/${a1.file_id}/view?usp=drivesdk` : '',
      thumb_url: a1 ? `https://drive.google.com/thumbnail?id=${a1.file_id}&sz=w400` : '',
      topic: '八月預告 ' + dKey.toUpperCase() + ' ' + topic,
      sub: '八月夏令營預告',
      now: now,
      stage: 'Phase 3 八月預告',
    }));

    if (wd === 0 || wd === 6) {
      const a916 = ASSETS.aug[dKey + '_9x16'];
      rows.push(buildRow_({
        qid: 'AUG_' + dKey.toUpperCase() + '_9x16',
        date: dateStr,
        time: '17:00',
        platform: 'IG Reels+Stories',
        ratio: '9x16',
        asset_id: 'AUG_' + dKey.toUpperCase() + '_9x16',
        img_url: a916 ? `https://drive.google.com/file/d/${a916.file_id}/view?usp=drivesdk` : '',
        thumb_url: a916 ? `https://drive.google.com/thumbnail?id=${a916.file_id}&sz=w400` : '',
        topic: '八月預告 ' + dKey.toUpperCase() + ' ' + topic,
        sub: '八月夏令營預告 Reels',
        now: now,
        stage: 'Phase 3 八月預告',
      }));
    }
    day.setDate(day.getDate() + 1);
  });

  // 一次寫入
  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, 22).setValues(rows);
  }

  ui.alert(
    '✅ 重灌完成',
    `已灌入 ${rows.length} 列：\n\n` +
    `Phase 1 暖身：22 列（無圖、待補素材）\n` +
    `Phase 2 七月預告：${rows.filter(r => r[0].startsWith('JUL_')).length} 列\n` +
    `Phase 3 八月預告：${rows.filter(r => r[0].startsWith('AUG_')).length} 列\n\n` +
    `下一步：執行 generateCopyV2() 產生所有文案`,
    ui.ButtonSet.OK
  );
}

// ============ 建單列 22 欄 ============
function buildRow_(o) {
  return [
    o.qid,                  // A queue_id
    o.date,                 // B 排程日期
    o.time,                 // C 排程時間
    o.platform,             // D 平台
    o.ratio,                // E 比例
    o.asset_id,             // F asset_id
    o.img_url,              // G 圖片URL
    o.thumb_url,            // H 縮圖URL
    o.topic,                // I 主題
    '', '', '', '',         // J-M 文案_主標/內文/Hashtags/CTA（generate 後填）
    '待審',                 // N 圖片審核
    '待審',                 // O 文案審核
    '草稿',                 // P 排程狀態
    '', '', '', '',         // Q 發文時間 R post_id S post_url T 錯誤
    o.now,                  // U 建立時間
    o.stage,                // V 備註（記錄階段）
  ];
}

function fmtDate_(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

// ============ 素材對照（從 schedule_v2_assets.json 取出後 inline）============
function getAssets_() {
  // 因為 Apps Script 不能直接讀 workspace 檔，這裡 inline 對照表（精簡版只放 file_id）
  // 完整版會在下一個檔 ASSETS_V2_DATA 提供
  return ASSETS_V2_DATA;
}
