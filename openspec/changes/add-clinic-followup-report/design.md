## Context

目前儀表板有三個報表視圖。總覽分析與區間比較共用同一份 `analysisResult`（由 `analyzeData` 從主試算表的 `sheetData` 產生）；競賽排行則走獨立資料流（`selectedCompetitionSpreadsheetIds` + `fetchMultipleSpreadsheets` + `aggregateCompetitionRankings`），因為它的資料範圍跨多份試算表。

診所相關邏輯在 `analyzeData` 內已有雛形：`analysis.ts:618-672` 的 `clinicSummary` 以 `customer['診所名稱']` 分組，統計 `total / potential / dealt / conversionRate / totalAmount`，輸出成 `clinicAnalysis`，唯一的消費端是 `AnalysisCharts.tsx` 的 Top 8 長條圖。

這次要做的回訪報告與該雛形的差異在於粒度：Top 8 圖只需要每家診所一個總數，回訪報告需要單一診所的月度趨勢、分布與逐筆客戶明細。此外「回訪」這個概念在現有程式碼中完全不存在，需要全新設計。

## Goals / Non-Goals

**Goals:**
- 列出分析區間內所有出現過的診所名稱，並可搜尋與排序
- 針對選定的單一診所產出可直接面對醫生的成效回訪報告
- 報告可匯出成 PDF，且列印版面乾淨、明細表跨頁可讀
- 沿用主畫面既有的分析區間與 PTA 閾值，不新增控制項語意
- 不改動總覽分析、區間比較、競賽排行的既有計算與顯示
- 測試先行，涵蓋新聚合邏輯與既有報表的回歸

**Non-Goals:**
- 不做多試算表複選（本次資料來源僅限目前主試算表）
- 不新增 CSV / Excel 匯出與相關依賴
- 不新增後端 API、資料庫或快取層
- 不修正既有 `clinicSummary` / Top 8 圖表的行為
- 不做跨區間的診所成效比較（那是區間比較的守備範圍）

## Decisions

### 1. 回訪報告資料掛進既有 `analyzeData`，而非比照競賽排行另做獨立資料流

新增 `clinicFollowUpAnalysis` 欄位到 `AnalysisResult`，在 `analyzeData` 內部用既有的 `customersArray` 產生。

替代方案是比照 `aggregateCompetitionRankings` 抽成獨立匯出的 pure function，再配上獨立 state 與 loading / error。該方案被否決，因為競賽排行之所以需要獨立資料流，唯一理由是它的資料範圍跨多份試算表、與主畫面不同；診所回訪的資料範圍與總覽分析完全相同（同一份主試算表、同一段區間、同一個 PTA 閾值）。硬套獨立資料流會白白多出一組 state、一次重複的表頭解析與日期解析，以及與主畫面不同步的更新時機，卻換不到任何隔離上的好處。

掛進 `analyzeData` 的另一個好處是更新時機自動與總覽分析一致：按「生成報告」或切換試算表時一起重算，使用者不會看到兩個報表對同一段區間給出不同數字。

### 2. 另建 `clinicFollowUpMap`，不改動既有 `clinicSummary`

新的明細收集在同一個 `customersArray.forEach` 迴圈內以獨立的累積器完成，既有 `clinicSummary` 的計算一行不動。

替代方案是擴充 `clinicSummary` 讓兩者共用一個累積器。該方案被否決：`clinicAnalysis` 的輸出形狀被 `AnalysisCharts.tsx` 直接消費，改動它就要同時驗證 Top 8 圖表沒有回歸，風險與收益不成比例。少量的迴圈內重複換取既有報表零風險是划算的。

### 3. 診所名稱改用模糊比對並過濾無效值

新增 `findClinicNameHeader()`，比照既有 `findStoreNameHeader()` 的優先序比對（`診所名稱` → `轉介診所` → `診所` → `醫院`），並重用 `INVALID_STORE_NAMES` 過濾 `#N/A` / `#REF!`。

既有 `clinicSummary` 用的是精確 key `customer['診所名稱']`，表頭若帶空白或叫「轉介診所」就會整批漏掉，且 `#N/A` 會被當成一間診所。新程式碼不沿用這兩個缺陷。與 `findStoreNameHeader` 的差別是：找不到時回傳空字串，不做 `headers[11]` 這種位置 fallback——診所欄位不存在時應該誠實顯示空狀態，而不是拿不相干的第 12 欄硬湊出一份錯誤報告。

### 4. 成交判定統一走 `checkIsDealt()`，日期統一走 `parseSheetDate()`

既有診所區塊用的是手寫的 inline 成交條件（`analysis.ts:649-651`），與其他區塊使用的 `checkIsDealt()` 涵蓋範圍不一致；`analyzeData` 內也有多處繞過 `parseSheetDate()` 直接 `new Date()`。新程式碼一律走既有的共用 helper，避免出現「進得了日期篩選卻算不出月份」或「同一筆資料在兩個報表成交狀態不同」的落差。

### 5. `daysSinceLastReferral` 在元件端計算，不放進分析結果

分析結果只保留 `lastReferralDate` 字串，「距今幾天」由元件端計算。

替代方案是在 `analyzeData` 內算好。該方案被否決，因為那會讓分析結果隨執行當日日期改變，測試需要凍結時間才能穩定斷言，徒增複雜度；而這個數字本來就只是顯示用途。

### 6. 匯出沿用 `window.print()`，不新增依賴

專案既有的唯一匯出機制就是 `window.print()` 搭配 `index.css` 的 `@media print` 規則，`package.json` 沒有任何 CSV / Excel / PDF 函式庫。使用者本次的需求是「帶著報告去跟醫生談」，列印版 PDF 完全滿足。

替代方案是引入 `xlsx` 或 `jspdf`。該方案被否決：會為單一功能新增依賴與新的匯出模式，且 PDF 的排版控制力反而不如瀏覽器列印。列印路徑需要額外處理的是明細表跨頁——加上 `thead { display: table-header-group }` 讓表頭每頁重複，並讓 `tr` 不被切斷。

### 7. 自動產生「回訪重點」文案

報告內以純函式 `buildTalkingPoints()` 從統計數字產生 3–5 條中文敘述（轉介量、聽損比例、配戴成效、轉介高峰月份、最近轉介距今天數）。

這是整份報告對使用者最直接的價值：面談時要講的話已經寫好，不需要業務人員自己從圖表推導。超過 60 天沒有新轉介時改為提醒文案，讓沉寂的診所被凸顯出來。

### 8. 測試分成聚合邏輯與元件互動兩層

分析層以手刻 `string[][]` 樣本驗證分組、排序、聽損與配戴計數、月度 bucket、缺欄位時回傳空陣列，以及既有 `clinicAnalysis` / `competitionRankingAnalysis` 不受影響。元件層 mock 掉 `react-chartjs-2` 後驗證預設選取、搜尋、切換診所、空狀態與姓名遮罩。

## Risks

- **`AnalysisResult` 新增欄位會讓既有測試的 mock 物件編譯失敗**：`src/App.test.tsx` 的 `analysisResult` mock 是完整型別物件，必須同步補上 `clinicFollowUpAnalysis: []`，否則 TS 編譯直接掛掉。
- **來源資料若無「診所名稱」欄位，整個報表為空**：以明確的空狀態文案處理，並指出應檢查來源工作表欄位，而非顯示空白畫面。
- **診所名稱在資料中可能有前後空白或全形差異造成分裂**：本次以 `trim()` 處理空白；更進一步的正規化（全半形、別名對應）不在本次範圍，若實務上出現再另案處理。
- **明細筆數多時列印頁數會很長**：屬預期行為（醫生要看的就是名單），以列印分頁樣式確保可讀性，不做截斷。
