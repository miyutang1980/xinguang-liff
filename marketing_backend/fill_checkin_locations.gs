/**
 * 太平新光 — 打卡地點輪換填入 V 欄
 *
 * 策略：
 *   - 30 個地點分 4 區：新光 7、軍功 7、廍子 7、太平中心 9
 *   - 95 篇貼文按列順序輪流，平均每個地點被輪到 3 次
 *   - V 欄會被覆蓋成「[L01 新高國小] 階段名」這種格式
 *   - 你發文時照 V 欄括號裡的地點名搜尋並打卡
 *
 * 用法：
 *   1. 貼進 Apps Script
 *   2. 執行 fillCheckinLocations()
 *   3. 等 5 秒、跳出 ✅
 */

const CHECKIN_SS_ID = '1DybgWBdCyvkEijMyaE46rKLtQD9J2ImjU8xeYCKSKnA';
const CHECKIN_QUEUE = '排程佇列 Posting_Queue';

const CHECKIN_LOCS = [
  { id: 'L01', area: '新光',         name: '新高國小' },
  { id: 'L02', area: '新光',         name: '新光國中' },
  { id: 'L03', area: '新光',         name: '新光黃昏市場' },
  { id: 'L04', area: '新光',         name: '全家便利商店 太平宜佳店' },
  { id: 'L05', area: '新光',         name: '新光重劃區' },
  { id: 'L06', area: '新光',         name: '新光公園' },
  { id: 'L07', area: '新光',         name: '新光國小幼兒園' },
  { id: 'L08', area: '軍功',         name: '軍功國小' },
  { id: 'L09', area: '軍功',         name: '麥當勞 台中軍功店' },
  { id: 'L10', area: '軍功',         name: '全聯福利中心 軍功門市' },
  { id: 'L11', area: '軍功',         name: '7-11 環東門市' },
  { id: 'L12', area: '軍功',         name: '星巴克 北屯軍功店' },
  { id: 'L13', area: '軍功',         name: '軍功公園' },
  { id: 'L14', area: '軍功',         name: '軍功國中' },
  { id: 'L15', area: '廍子',         name: '廍子國小' },
  { id: 'L16', area: '廍子',         name: '廍子公園' },
  { id: 'L17', area: '廍子',         name: '新都生態公園' },
  { id: 'L18', area: '廍子',         name: '廍子重劃區' },
  { id: 'L19', area: '廍子',         name: '中臺科技大學' },
  { id: 'L20', area: '廍子',         name: '藍天白雲橋' },
  { id: 'L21', area: '廍子',         name: '浪漫情人橋' },
  { id: 'L22', area: '太平中心',     name: '太平市公所' },
  { id: 'L23', area: '太平中心',     name: '麥當勞 太平中興店' },
  { id: 'L24', area: '太平中心',     name: '太平運動公園' },
  { id: 'L25', area: '太平中心',     name: '太平買菜市場' },
  { id: 'L26', area: '太平中心',     name: '豪泰百貨 太平店' },
  { id: 'L27', area: '太平中心',     name: '全聯福利中心 太平太平店' },
  { id: 'L28', area: '太平中心',     name: '德興公園' },
  { id: 'L29', area: '太平中心',     name: '太平圖書館' },
  { id: 'L30', area: '太平中心',     name: '太平警分局' },
];

function fillCheckinLocations() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(CHECKIN_SS_ID);
  const sh = ss.getSheetByName(CHECKIN_QUEUE);
  if (!sh) throw new Error('找不到分頁');

  const last = sh.getLastRow();
  if (last < 2) throw new Error('Posting_Queue 是空的');

  const totalRows = last - 1;
  const ans = ui.alert(
    '打卡地點輪換填入',
    `將為 ${totalRows} 列填入 30 個太平東邊打卡地點：\n\n` +
    `三學區（新光/軍功/廍子）+ 太平中心商圈\n` +
    `每個地點平均被輪到 ${Math.ceil(totalRows / 30)} 次\n\n` +
    `會覆蓋 V 欄（備註）內容、保留原階段名\n\n` +
    `按 OK 確認`,
    ui.ButtonSet.OK_CANCEL
  );
  if (ans !== ui.Button.OK) return;

  // 讀現有 V 欄階段名
  const stages = sh.getRange(2, 22, totalRows, 1).getValues();

  const updated = [];
  for (let i = 0; i < totalRows; i++) {
    const loc = CHECKIN_LOCS[i % 30];
    const oldStage = String(stages[i][0] || '');
    const newVal = `[${loc.id} ${loc.name}・${loc.area}] ${oldStage}`;
    updated.push([newVal]);
  }

  sh.getRange(2, 22, totalRows, 1).setValues(updated);

  ui.alert(
    '✅ 完成',
    `已為 ${totalRows} 列填入打卡地點。\n\n` +
    `打開 Posting_Queue → 看 V 欄（備註）：\n` +
    `[L01 新高國小・新光] Phase 1 暖身\n` +
    `[L02 新光國中・新光] Phase 1 暖身\n` +
    `...\n\n` +
    `發文時照括號裡的地點搜尋並打卡。`
  );
}


/**
 * 統計各地點被分配幾次
 */
function checkinSummary() {
  const ss = SpreadsheetApp.openById(CHECKIN_SS_ID);
  const sh = ss.getSheetByName(CHECKIN_QUEUE);
  const last = sh.getLastRow();
  const data = sh.getRange(2, 22, last - 1, 1).getValues();
  const counts = {};
  data.forEach(r => {
    const v = String(r[0] || '');
    const m = v.match(/\[(L\d{2})/);
    if (m) counts[m[1]] = (counts[m[1]] || 0) + 1;
  });
  let msg = '打卡地點分配統計：\n\n';
  CHECKIN_LOCS.forEach(loc => {
    msg += `${loc.id} ${loc.name}：${counts[loc.id] || 0} 次\n`;
  });
  SpreadsheetApp.getUi().alert(msg);
}
