# 上線 SOP v3 — 5 月全面調整版

本次升級三大重點：
1. 後台「歷史分析」加中文洞察 + 戰術建議
2. 後台「排程看板」改成未來 14 天即時讀 Posting_Queue
3. 數據儀表加 7/30/90 天切換 + 每日趨勢圖
4. 自動回覆引擎 v2：AI 自然語意 + 高潛家長 + 負評警示 + LINE OA 預填

---

## Step 1：更新 Apps Script 兩支主檔（5 分鐘）

### 1-1. 開 Apps Script 編輯器
打開試算表 → Extensions → Apps Script

### 1-2. 換掉 marketing_dashboard.gs
- 左邊找「marketing_dashboard」這支檔
- 全選原內容刪掉
- 從 GitHub 貼新版：
  - 點開 raw URL（GitHub 已最新）並全選複製：
  - `marketing_backend/marketing_dashboard.gs`
- 貼回去 → 按 💾 儲存

### 1-3. 換掉 dashboard.html
- 左邊找「dashboard」這支 HTML 檔
- 全選原內容刪掉
- 從 GitHub 貼新版：`marketing_backend/dashboard.html`
- 貼回去 → 按 💾 儲存

### 1-4. 新增 auto_reply_engine_v2.gs（Phase 3）
- 左邊「+」→「Script」
- 命名：`auto_reply_engine_v2`
- 貼上 GitHub `marketing_backend/auto_reply_engine_v2.gs`
- 儲存

---

## Step 2：補 Settings 兩個設定（2 分鐘）

打開試算表的「Settings」分頁，加兩列：

| Key | Value | 說明 |
|---|---|---|
| `LINE_OA_HANDLE` | `@143qbory` | LINE OA 帳號（已預設、可確認） |
| `OPENAI_API_KEY` | `（空著）` | 暫時空、Phase 3.5 再開 |
| `AI_REPLY_ENABLE` | `FALSE` | 預設 FALSE 安全、要開時改 TRUE |

---

## Step 3：重新部署 Web App（3 分鐘）

1. Apps Script 編輯器右上角 → 部署 → 管理部署
2. 找「dashboard」這個現有部署 → 點 ✏️ 編輯
3. 版本選「新版本」→ 描述填「v3 中文洞察+14天排程+7/30/90+高潛家長」
4. 部署
5. 開啟 Web App URL（之前那個 AKfycbw_xFWnKL... 的網址）

---

## Step 4：第一次測試（5 分鐘）

打開 Web App 後依序檢查：

### 4-1. 數據儀表頁
- 點上方「7 天 / 30 天 / 90 天」按鈕、KPI 數字應該會切換
- 下方「每日互動趨勢」應該有 IG 粉色 + FB 藍色直條圖

### 4-2. 排程看板頁
- 應該看到未來 14 天 IG/FB 排程卡片
- 每張卡片：日期、平台、時間、主題、圖文審核狀態
- 若顯示「未來 14 天無排程」表示 Posting_Queue 沒未來資料、是正常的

### 4-3. 歷史分析頁
- 頂端橘金色區塊應該有「📊 今日中文洞察 + 戰術建議」
- 至少看到 5-7 條洞察卡片：紅色（critical）、橘色（warning）、綠色（opportunity）
- 每條都有「戰術建議」虛線框

### 4-4. 高潛家長頁（新分頁）
- 第一次會空、因為 Hot_Leads 表還沒生成
- 等 Step 5 跑完才有資料

---

## Step 5：啟動 Phase 3 自動回覆 v2（3 分鐘）

### 5-1. 手動跑一次測試
Apps Script 編輯器 → 選 `auto_reply_engine_v2.gs` → 上方下拉選 `pollAllPlatformsV2` → 執行 ▶️

第一次會跳授權、按允許

### 5-2. 確認自動建立兩張新表
打開試算表，應該多出來：
- `高潛家長 Hot_Leads`
- `負評警示 Alerts`

### 5-3. 改成新觸發器（取代舊版）
Apps Script → 選 `auto_reply_engine_v2.gs` → 函式 `installReplyTriggersV2` → 執行
跳出對話框「自動回覆 v2 觸發器：每 5 分鐘掃描」就成功

### 5-4. 重新整理 Web App 高潛家長頁
- 等 5-10 分鐘讓觸發器跑一次（或手動執行 `refreshHotLeads_`）
- 重新整理 Web App → 高潛家長頁應該有資料

---

## Step 6（選配）：開啟 AI 自然語意回覆

只在你想要 AI 接手「規則沒中」的留言時才開：

1. 申請 OpenAI API Key（platform.openai.com）、月用量上限設 NT 300（用 gpt-4o-mini 約 1500 則）
2. Settings 分頁：
   - `OPENAI_API_KEY` 填上 sk-xxx
   - `AI_REPLY_ENABLE` 改成 `TRUE`
3. Apps Script 跑 `testAIReplyOnce` 測試一次
4. 若回覆 OK 就會在下次掃描時自動接手

**重要**：AI 回覆有禁字過濾、若 AI 不小心說了違規字、系統會自動丟掉不發。

---

## 常見問題

### Q1: 中文洞察區塊顯示「尚無可用洞察」
A: Historical_Posts 沒資料。先到 Apps Script 跑 `backfillIGHistorical` 跟 `backfillFBHistorical`。

### Q2: 排程看板顯示空白
A: Posting_Queue 沒未來日期的列、或都已標「已發」「失敗」。再到 Posting_Queue 加新貼文即可。

### Q3: 高潛家長一直是 0
A: 互動紀錄 Interactions 沒累積到 3 次同一帳號。等自動回覆觸發器跑一段時間後就會有。

### Q4: AI 回覆給的內容怪怪的
A: 在 Settings 把 `AI_REPLY_ENABLE` 改 FALSE 暫停。然後跟 E 小編說、調整 system prompt（在 `auto_reply_engine_v2.gs` line 175-185）。

### Q5: 想看哪些貼文是被 AI 自動回的
A: 互動紀錄 Interactions 第 13 欄「處理方式」會寫 `AI`，跟 `自動`（規則）區分。

---

## 後續維運（每週做一次）

1. **負評警示 Alerts** 巡一次、處理「待處理」狀態
2. **高潛家長** 看 🔥 高溫的、優先 LINE 邀約
3. **互動紀錄 Interactions** 看 AI 回覆的內容、不對勁的截圖回報
4. **Settings 的 OPENAI_API_KEY** 用量看一下（避免超額）

---

## GitHub 路徑速查

所有檔案都在：`https://github.com/miyutang1980/xinguang-liff/tree/main/marketing_backend`

關鍵檔：
- `marketing_dashboard.gs`（後台主邏輯）
- `dashboard.html`（前台介面）
- `auto_reply_engine_v2.gs`（Phase 3 自動回覆）
- `auto_reply_engine.gs`（舊版、可保留也可刪）

Web App URL：之前那組 AKfycbw_xFWnKL... 不變、改用「新版本」部署即可。
