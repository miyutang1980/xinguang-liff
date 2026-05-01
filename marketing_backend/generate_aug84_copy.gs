/**
 * 太平新光八月 84 篇文案批次生成器
 * 
 * 用法：
 *   1. 貼進 Apps Script 任何一個 .gs 檔
 *   2. 確認 Settings 分頁有 OPENAI_API_KEY
 *   3. 執行 generateAug84Copy()
 *   4. 自動跑 84 列、寫回 Posting_Queue
 * 
 * SOP：
 *   ✅ 精細檢測 50 分鐘、Sarah 老師 1 對 1、早鳥 85 折、剩 X 位
 *   ❌ 免費試聽、免費抵註冊費、免費評測、省 500 元、同校兩人組
 *   ✅ 不指名對手品牌、繁體中文（台灣）、無 emoji 除 CTA 引導
 *   ✅ 科學為主，英文為媒
 */

const COPY_GEN_SS_ID = '1DybgWBdCyvkEijMyaE46rKLtQD9J2ImjU8xeYCKSKnA';
const COPY_GEN_QUEUE = '排程佇列 Posting_Queue';

// 八月 28 天主題（D01-D28）
const AUG_TOPICS = {
  d01: { topic: '拉力與張力',     mainQ: '繩子拉得緊有多大力？',         mainEn: 'Tension Forces',     vocab: 'pull / tight / tension',        week: '升級偵探：物理進階' },
  d02: { topic: '扭力與旋轉',     mainQ: '為什麼螺絲要轉才會進去？',     mainEn: 'Torque',             vocab: 'twist / rotate / torque',       week: '升級偵探：物理進階' },
  d03: { topic: '動量與碰撞',     mainQ: '為什麼火車比腳踏車更難停下？', mainEn: 'Momentum & Collision', vocab: 'mass / momentum / crash',     week: '升級偵探：物理進階' },
  d04: { topic: '彈性與形變',     mainQ: '橡皮筋拉太長為什麼會斷？',     mainEn: 'Elasticity',         vocab: 'stretch / bend / break',        week: '升級偵探：物理進階' },
  d05: { topic: '波與震動',       mainQ: '聲音怎麼傳到耳朵裡？',         mainEn: 'Waves & Vibration',  vocab: 'wave / sound / vibrate',        week: '升級偵探：物理進階' },
  d06: { topic: '透鏡與光路',     mainQ: '放大鏡為什麼能看清楚指紋？',   mainEn: 'Lenses & Light',     vocab: 'lens / focus / light',          week: '升級偵探：物理進階' },
  d07: { topic: '反射與成像',     mainQ: '鏡子裡的字為什麼是反的？',     mainEn: 'Reflection',         vocab: 'mirror / reflect / image',      week: '升級偵探：物理進階' },
  d08: { topic: '酸鹼與指示劑',   mainQ: '紫色高麗菜汁怎麼變色？',       mainEn: 'Acid & Base',        vocab: 'acid / base / indicator',       week: '化學偵探：分子實驗' },
  d09: { topic: '溶解與飽和',     mainQ: '糖加多少會溶不掉？',           mainEn: 'Solubility',         vocab: 'dissolve / saturate / solution', week: '化學偵探：分子實驗' },
  d10: { topic: '結晶與析出',     mainQ: '冰糖是怎麼長出來的？',         mainEn: 'Crystallization',    vocab: 'crystal / form / grow',         week: '化學偵探：分子實驗' },
  d11: { topic: '氧化與還原',     mainQ: '蘋果切開為什麼會變黑？',       mainEn: 'Oxidation',          vocab: 'oxidize / rust / brown',        week: '化學偵探：分子實驗' },
  d12: { topic: '分離與過濾',     mainQ: '泥水怎麼變成清水？',           mainEn: 'Filtration',         vocab: 'filter / pure / separate',      week: '化學偵探：分子實驗' },
  d13: { topic: '蒸發與凝結',     mainQ: '水珠為什麼會在杯子外面？',     mainEn: 'Evaporation',        vocab: 'evaporate / condense / vapor',  week: '化學偵探：分子實驗' },
  d14: { topic: '燃燒與火焰',     mainQ: '蠟燭為什麼蓋住會熄？',         mainEn: 'Combustion',         vocab: 'burn / flame / oxygen',         week: '化學偵探：分子實驗' },
  d15: { topic: '指紋與證物',     mainQ: '每個人的指紋真的不一樣嗎？',   mainEn: 'Fingerprints',       vocab: 'unique / pattern / print',      week: '生物偵探：身體與生命' },
  d16: { topic: '骨骼與關節',     mainQ: '為什麼手腕能轉一圈？',         mainEn: 'Bones & Joints',     vocab: 'bone / joint / move',           week: '生物偵探：身體與生命' },
  d17: { topic: '感官與反應',     mainQ: '為什麼吃辣會流眼淚？',         mainEn: 'Senses',             vocab: 'taste / smell / spicy',         week: '生物偵探：身體與生命' },
  d18: { topic: '心臟與血液',     mainQ: '心臟一天跳幾下？',             mainEn: 'Heart & Blood',      vocab: 'heart / pulse / pump',          week: '生物偵探：身體與生命' },
  d19: { topic: '昆蟲與變態',     mainQ: '毛毛蟲怎麼變蝴蝶？',           mainEn: 'Metamorphosis',      vocab: 'caterpillar / butterfly / change', week: '生物偵探：身體與生命' },
  d20: { topic: '植物與光合',     mainQ: '葉子為什麼是綠色的？',         mainEn: 'Photosynthesis',     vocab: 'leaf / sun / energy',           week: '生物偵探：身體與生命' },
  d21: { topic: 'DNA 與遺傳',     mainQ: '為什麼小孩長得像爸媽？',       mainEn: 'DNA',                vocab: 'gene / inherit / DNA',          week: '生物偵探：身體與生命' },
  d22: { topic: '地震與板塊',     mainQ: '地震為什麼會搖？',             mainEn: 'Earthquakes',        vocab: 'shake / plate / fault',         week: '鑑識實戰：證物與結案' },
  d23: { topic: '唇紋與筆跡',     mainQ: '簽名能模仿嗎？',               mainEn: 'Signatures',         vocab: 'signature / unique / trace',    week: '鑑識實戰：證物與結案' },
  d24: { topic: '足跡分析',       mainQ: '腳印能說出多少秘密？',         mainEn: 'Footprints',         vocab: 'footprint / depth / size',      week: '鑑識實戰：證物與結案' },
  d25: { topic: '時間估算',       mainQ: '冰塊融多久能算時間？',         mainEn: 'Time Estimation',    vocab: 'estimate / melt / clock',       week: '鑑識實戰：證物與結案' },
  d26: { topic: '骨骼鑑定',       mainQ: '從骨頭能知道幾歲嗎？',         mainEn: 'Bone Forensics',     vocab: 'skeleton / age / forensic',     week: '鑑識實戰：證物與結案' },
  d27: { topic: '粉末檢驗',       mainQ: '白色粉末是糖還是鹽？',         mainEn: 'Powder Analysis',    vocab: 'powder / test / sample',        week: '鑑識實戰：證物與結案' },
  d28: { topic: '結案發表畢業',   mainQ: '你是夏令營小偵探了嗎？',       mainEn: 'Case Closed',        vocab: 'solve / present / detective',   week: '鑑識實戰：證物與結案' },
};


function generateAug84Copy() {
  const ss = SpreadsheetApp.openById(COPY_GEN_SS_ID);
  const sh = ss.getSheetByName(COPY_GEN_QUEUE);
  if (!sh) throw new Error('找不到分頁：' + COPY_GEN_QUEUE);

  const settingsSh = ss.getSheetByName('Settings') || ss.getSheetByName('設定 Settings');
  const apiKey = getSettingValue_(settingsSh, 'OPENAI_API_KEY');
  if (!apiKey) throw new Error('Settings 沒有 OPENAI_API_KEY');

  const last = sh.getLastRow();
  if (last < 2) throw new Error('Posting_Queue 是空的');

  const data = sh.getRange(2, 1, last - 1, 22).getValues();
  // 欄位 index：0=queue_id 5=asset_id 8=主題 9=文案_主標 10=文案_內文 11=Hashtags 12=CTA連結

  let processed = 0;
  let skipped = 0;
  let errors = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowNum = i + 2;
    const assetId = String(row[5] || '');
    const topic = String(row[8] || '');
    const platform = String(row[3] || '');
    const ratio = String(row[4] || '');
    const existingBody = String(row[10] || '');

    // 已有內文則跳過
    if (existingBody.trim().length > 30) { skipped++; continue; }

    // 從 asset_id 解析 Dxx
    const m = assetId.match(/D(\d{2})/i);
    if (!m) { errors.push(`列${rowNum}: 無法解析 D 編號 (${assetId})`); continue; }
    const dKey = 'd' + m[1];
    const meta = AUG_TOPICS[dKey];
    if (!meta) { errors.push(`列${rowNum}: ${dKey} 不在主題表`); continue; }

    try {
      const copy = callOpenAI_(apiKey, meta, platform, ratio);
      // 寫回欄 J/K/L/M（10/11/12/13）
      sh.getRange(rowNum, 10).setValue(copy.headline);
      sh.getRange(rowNum, 11).setValue(copy.body);
      sh.getRange(rowNum, 12).setValue(copy.hashtags);
      sh.getRange(rowNum, 13).setValue(copy.cta);
      processed++;
      Utilities.sleep(800); // 避免 rate limit
      if (processed % 10 === 0) {
        Logger.log(`進度：${processed}/${data.length}`);
        SpreadsheetApp.flush();
      }
    } catch (e) {
      errors.push(`列${rowNum}: ${e.message}`);
    }
  }

  Logger.log('====== 完成 ======');
  Logger.log(`已生成：${processed} 篇`);
  Logger.log(`跳過（已有內文）：${skipped} 篇`);
  Logger.log(`錯誤：${errors.length} 筆`);
  if (errors.length) errors.slice(0, 10).forEach(e => Logger.log(e));

  SpreadsheetApp.getUi().alert(
    `八月 84 篇文案生成完成\n\n` +
    `已生成：${processed} 篇\n` +
    `跳過（已有）：${skipped} 篇\n` +
    `錯誤：${errors.length} 筆\n\n` +
    `（如有錯誤詳見 Apps Script 執行記錄）`
  );
}


function callOpenAI_(apiKey, meta, platform, ratio) {
  const isReel = ratio === '9x16';
  const isFB16x9 = ratio === '16x9';
  const platformHint = isReel ? 'IG Reels + Stories（短影音/直式）'
                      : isFB16x9 ? 'FB 動態（橫式）'
                      : 'IG + FB 動態（方形）';

  const sys = [
    '你是台灣兒童美語+科學夏令營「弋果美語太平新光分校」的社群行銷文案師。',
    '品牌口吻：溫暖、專業、家長感、清晰可讀。',
    '永久 SOP：科學為主，英文為媒。寫作給家長看。',
    '',
    '✅ 必用：精細檢測 50 分鐘、Sarah 老師 1 對 1、早鳥 85 折、剩 X 位、繁體中文（台灣）',
    '❌ 嚴禁：免費試聽、免費抵註冊費、免費評測、省 500 元、同校兩人組、emoji（除 CTA 引導符號 →）',
    '❌ 不指名任何對手品牌',
    '',
    '輸出格式：嚴格 JSON，不要 markdown 不要前後文。',
    '{"headline":"...","body":"...","hashtags":"...","cta":"..."}'
  ].join('\n');

  const userPrompt = [
    `主題：八月 ${meta.topic}（${meta.mainEn}）`,
    `主問句：${meta.mainQ}`,
    `週主題：${meta.week}`,
    `英文詞彙：${meta.vocab}`,
    `平台：${platformHint}`,
    '',
    '請生成 4 個欄位：',
    '1. headline（主標）：12-18 字，吸睛、含主題或主問句、繁體中文',
    `2. body（內文）：${isReel ? '60-80' : '90-130'} 字，第二人稱對家長說話，包含：`,
    '   - 引入問題（用主問句）',
    '   - 一句精細檢測 50 分鐘 + Sarah 老師 1 對 1 的價值',
    '   - 一句早鳥 85 折 + 剩 X 位（X 自己編 3-8）的引導',
    '3. hashtags（5 個）：#弋果美語 #太平新光 #科學夏令營 + 2 個跟主題有關的，空格分隔',
    '4. cta：1 行 25 字內，含「→ 私訊預約」或「→ 加 LINE @143qbory」',
    '',
    '⚠️ 不要寫「免費」「抵註冊費」「省」「試聽」「評測」「同校」'
  ].join('\n');

  const payload = {
    model: 'gpt-4o-mini',
    messages: [
      { role: 'system', content: sys },
      { role: 'user', content: userPrompt }
    ],
    temperature: 0.7,
    response_format: { type: 'json_object' }
  };

  const res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  if (code !== 200) throw new Error('OpenAI HTTP ' + code + ': ' + res.getContentText().substring(0, 200));

  const json = JSON.parse(res.getContentText());
  const content = json.choices[0].message.content;
  const parsed = JSON.parse(content);

  // 安全檢查
  const allText = (parsed.headline + parsed.body + parsed.hashtags + parsed.cta);
  const banned = ['免費', '抵註冊', '省 500', '省500', '試聽', '評測', '同校兩人'];
  for (const b of banned) {
    if (allText.indexOf(b) >= 0) throw new Error('違規詞觸發：' + b);
  }
  return parsed;
}


function getSettingValue_(sh, key) {
  if (!sh) return '';
  const last = sh.getLastRow();
  if (last < 2) return '';
  const vals = sh.getRange(1, 1, last, 2).getValues();
  for (const row of vals) {
    if (String(row[0]).trim() === key) return String(row[1]).trim();
  }
  return '';
}


/**
 * 測試單列（避免一次跑爆）
 */
function testGenerateOneRow() {
  const ss = SpreadsheetApp.openById(COPY_GEN_SS_ID);
  const sh = ss.getSheetByName(COPY_GEN_QUEUE);
  const settingsSh = ss.getSheetByName('Settings') || ss.getSheetByName('設定 Settings');
  const apiKey = getSettingValue_(settingsSh, 'OPENAI_API_KEY');
  if (!apiKey) throw new Error('Settings 沒有 OPENAI_API_KEY');

  // 跑第 2 列
  const row = sh.getRange(2, 1, 1, 22).getValues()[0];
  const assetId = String(row[5]);
  const platform = String(row[3]);
  const ratio = String(row[4]);
  const m = assetId.match(/D(\d{2})/i);
  const meta = AUG_TOPICS['d' + m[1]];

  const copy = callOpenAI_(apiKey, meta, platform, ratio);
  Logger.log('Headline: ' + copy.headline);
  Logger.log('Body: ' + copy.body);
  Logger.log('Hashtags: ' + copy.hashtags);
  Logger.log('CTA: ' + copy.cta);

  // 寫回第 2 列
  sh.getRange(2, 10).setValue(copy.headline);
  sh.getRange(2, 11).setValue(copy.body);
  sh.getRange(2, 12).setValue(copy.hashtags);
  sh.getRange(2, 13).setValue(copy.cta);
  SpreadsheetApp.getUi().alert('測試列 2 已寫入，請到 Sheet 看效果。OK 後再執行 generateAug84Copy()');
}
