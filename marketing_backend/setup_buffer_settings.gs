/**
 * Buffer API 設定一次性寫入
 * 執行：在 Apps Script 編輯器選 setupBufferSettings → 執行
 * 完成後可以刪這支或留著當文件
 */
function setupBufferSettings() {
  const SS_ID = '1DybgWBdCyvkEijMyaE46rKLtQD9J2ImjU8xeYCKSKnA';
  const ss = SpreadsheetApp.openById(SS_ID);
  let sh = ss.getSheetByName('Settings');
  if (!sh) sh = ss.insertSheet('Settings');

  const settings = [
    ['BUFFER_API_KEY',          'SKXfOWEKJf2COC6dLmIXCaliUGftNygKJeVslxODqqV'],
    ['BUFFER_ORG_ID',           '69f835eac8b1c4e0dbc9f94d'],
    ['BUFFER_IG_CHANNEL_ID',    '69fabd065c4c051afa157aa7'],
    ['BUFFER_FB_CHANNEL_ID',    '69fabda75c4c051afa157fc5'],
    ['BUFFER_TIKTOK_CHANNEL_ID','69fac1a45c4c051afa1590da']
  ];

  // 找出已存在的 key、決定 update 或 append
  const last = sh.getLastRow();
  const existing = last > 0 ? sh.getRange(1, 1, last, 2).getValues() : [];
  const keyToRow = {};
  existing.forEach((r, i) => { if (r[0]) keyToRow[r[0]] = i + 1; });

  settings.forEach(([k, v]) => {
    if (keyToRow[k]) {
      sh.getRange(keyToRow[k], 2).setValue(v);
    } else {
      sh.appendRow([k, v]);
    }
  });

  Logger.log('Buffer 設定寫入完成。請執行 testBufferConnection 驗證。');
  return '完成';
}

/** 驗證 token + channels 可以打通 */
function testBufferConnection() {
  const apiKey = pe_getSetting_('BUFFER_API_KEY');
  const orgId  = pe_getSetting_('BUFFER_ORG_ID');
  if (!apiKey || !orgId) {
    Logger.log('❌ 缺 BUFFER_API_KEY 或 BUFFER_ORG_ID、請先跑 setupBufferSettings');
    return;
  }
  const res = UrlFetchApp.fetch('https://api.buffer.com/graphql', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify({
      query: 'query($input: ChannelsInput!){ channels(input:$input){ id name service isDisconnected } }',
      variables: { input: { organizationId: orgId } }
    }),
    muteHttpExceptions: true
  });
  Logger.log(res.getContentText());
}
