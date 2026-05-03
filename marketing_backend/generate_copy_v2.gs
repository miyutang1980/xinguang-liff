/**
 * 太平新光 — 文案生成 v2
 * 依照 V 欄「階段」自動切換 4 種口吻：
 *   Phase 1 暖身 → 介紹型（不催促）
 *   Phase 2 七月預告 → 「下個月開始上 → 預告」
 *   Phase 3 八月預告 → 「再過一個月 → 早鳥剩 X 天」
 *
 * SOP（不變）：
 *   ✅ 精細檢測 50 分鐘、E 小編 1 對 1、早鳥 85 折、剩 X 位
 *   ❌ 免費試聽、免費抵註冊費、免費評測、省 500 元、同校兩人組
 *   ✅ 不指名對手品牌、繁體中文（台灣）、無 emoji 除 CTA →
 *   ✅ 科學為主，英文為媒
 *   LINE OA：@143qbory
 */

const COPYV2_SS_ID = '1DybgWBdCyvkEijMyaE46rKLtQD9J2ImjU8xeYCKSKnA';
const COPYV2_QUEUE = '排程佇列 Posting_Queue';

function generateCopyV2() {
  const ss = SpreadsheetApp.openById(COPYV2_SS_ID);
  const sh = ss.getSheetByName(COPYV2_QUEUE);
  if (!sh) throw new Error('找不到分頁：' + COPYV2_QUEUE);

  const settingsSh = ss.getSheetByName('Settings') || ss.getSheetByName('設定 Settings');
  const apiKey = getSettingValueV2_(settingsSh, 'OPENAI_API_KEY');
  if (!apiKey) throw new Error('Settings 沒有 OPENAI_API_KEY');

  const last = sh.getLastRow();
  if (last < 2) throw new Error('Posting_Queue 是空的');

  const data = sh.getRange(2, 1, last - 1, 22).getValues();

  let processed = 0, skipped = 0;
  const errors = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowNum = i + 2;
    const platform = String(row[3] || '');
    const ratio = String(row[4] || '');
    const topic = String(row[8] || '');
    const existingBody = String(row[10] || '');
    const stage = String(row[21] || ''); // V 欄

    if (existingBody.trim().length > 30) { skipped++; continue; }

    try {
      const copy = callOpenAIV2_(apiKey, { topic, platform, ratio, stage });
      sh.getRange(rowNum, 10).setValue(copy.headline);
      sh.getRange(rowNum, 11).setValue(copy.body);
      sh.getRange(rowNum, 12).setValue(copy.hashtags);
      sh.getRange(rowNum, 13).setValue(copy.cta);
      processed++;
      Utilities.sleep(800);
      if (processed % 10 === 0) {
        Logger.log(`進度：${processed}/${data.length}`);
        SpreadsheetApp.flush();
      }
    } catch (e) {
      errors.push(`列${rowNum}: ${e.message}`);
    }
  }

  Logger.log('====== 完成 ======');
  Logger.log(`已生成：${processed}`);
  Logger.log(`跳過（已有內文）：${skipped}`);
  Logger.log(`錯誤：${errors.length}`);
  if (errors.length) errors.slice(0, 10).forEach(e => Logger.log(e));

  SpreadsheetApp.getUi().alert(
    `文案 v2 生成完成\n\n已生成：${processed} 篇\n跳過：${skipped} 篇\n錯誤：${errors.length} 筆`
  );
}


function callOpenAIV2_(apiKey, ctx) {
  const { topic, platform, ratio, stage } = ctx;
  const isReel = ratio === '9x16';
  const isFB16x9 = ratio === '16x9';
  const platformHint = isReel ? 'IG Reels + Stories（短影音/直式）'
                      : isFB16x9 ? 'FB 動態（橫式）'
                      : 'IG + FB 動態（方形）';

  // 依階段切換口吻
  let stageInstruction = '';
  if (stage.indexOf('Phase 1') >= 0) {
    stageInstruction = `階段：暖身介紹（4-5 月初）。
口吻：溫和介紹、不催促、不打折扣。重點是「讓家長認識弋果美語太平新光分校」。
寫作目標：建立品牌印象，引發好奇。`;
  } else if (stage.indexOf('七月預告') >= 0) {
    stageInstruction = `階段：七月夏令營預告（5/11-6/8 發、7/1 開課）。
口吻：「下個月就要上這堂課」的預告感，營造期待。
必含：「七月夏令營 × 偵探科學派」、實際課程主題、E 小編 1 對 1、精細檢測 50 分鐘
不要寫：早鳥催促（這不是早鳥階段）。`;
  } else if (stage.indexOf('八月預告') >= 0) {
    stageInstruction = `階段：八月夏令營預告（6/9-7/5 發、8/1 開課）。
口吻：「七月梯已開、八月梯名額快滿」的緊迫感。
必含：「八月夏令營 × 鑑識科學派」、實際課程主題、E 小編 1 對 1、早鳥 85 折剩 X 位（X 自編 3-8）。
要寫：催促家長把握早鳥。`;
  } else {
    stageInstruction = `階段：一般夏令營課程內容。口吻溫和、實用、聚焦課程內容。`;
  }

  const sys = [
    '你是台灣兒童美語+科學夏令營「弋果美語太平新光分校」的社群行銷文案師。',
    '品牌口吻：溫暖、專業、家長感、清晰可讀。',
    '永久 SOP：科學為主，英文為媒。寫作給家長看。',
    '',
    stageInstruction,
    '',
    '✅ 必用：精細檢測 50 分鐘、E 小編 1 對 1、繁體中文（台灣）',
    '❌ 嚴禁：免費試聽、免費抵註冊費、免費評測、省 500 元、同校兩人組、emoji（除 CTA 引導符號 →）',
    '❌ 不指名任何對手品牌',
    '',
    '輸出格式：嚴格 JSON，不要 markdown 不要前後文。',
    '{"headline":"...","body":"...","hashtags":"...","cta":"..."}',
  ].join('\n');

  const userPrompt = [
    `主題：${topic}`,
    `平台：${platformHint}`,
    `階段：${stage}`,
    '',
    '請生成 4 個欄位：',
    '1. headline（主標）：12-18 字，吸睛、繁體中文',
    `2. body（內文）：${isReel ? '60-80' : '90-130'} 字，第二人稱對家長說話，帶課程主題與弋果價值`,
    '3. hashtags（5 個）：#弋果美語 #太平新光 + 3 個跟主題相關的，空格分隔',
    '4. cta：1 行 25 字內，含「→ 私訊預約」或「→ 加 LINE @143qbory」',
    '',
    '⚠️ 不要寫「免費」「抵註冊費」「省」「試聽」「評測」「同校」',
  ].join('\n');

  const payload = {
    model: 'gpt-4o-mini',
    messages: [
      { role: 'system', content: sys },
      { role: 'user', content: userPrompt },
    ],
    temperature: 0.7,
    response_format: { type: 'json_object' },
  };

  const res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const code = res.getResponseCode();
  if (code !== 200) throw new Error('OpenAI HTTP ' + code + ': ' + res.getContentText().substring(0, 200));
  const json = JSON.parse(res.getContentText());
  const content = json.choices[0].message.content;
  const parsed = JSON.parse(content);

  const all = (parsed.headline + parsed.body + parsed.hashtags + parsed.cta);
  const banned = ['免費', '抵註冊', '省 500', '省500', '試聽', '評測', '同校兩人'];
  for (const b of banned) {
    if (all.indexOf(b) >= 0) throw new Error('違規詞：' + b);
  }
  return parsed;
}


function getSettingValueV2_(sh, key) {
  if (!sh) return '';
  const last = sh.getLastRow();
  if (last < 2) return '';
  const vals = sh.getRange(1, 1, last, 2).getValues();
  for (const row of vals) {
    if (String(row[0]).trim() === key) return String(row[1]).trim();
  }
  return '';
}
