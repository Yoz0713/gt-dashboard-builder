## 1. Test Baseline

- [x] 1.1 為競賽排行新增多試算表樣本資料測試，覆蓋跨目標聽力中心的總轉介人次加總與 PTA 聽損 / 正常拆分（僅看聽力、不考慮成交）
- [x] 1.2 新增排序測試，驗證競賽排行依總轉介人次降冪排序，且同分時以門市名稱字典序穩定排序
- [x] 1.3 為 `App` / `Controls` 補上回歸測試，鎖定新增競賽複選後總覽分析與區間比較仍維持原本行為
- [x] 1.4 新增測試：缺少服務日期或轉介門市欄位的試算表應被跳過且不影響其他試算表的聚合結果
- [x] 1.5 新增測試：PTA 欄位缺失或無法解析時，該筆資料仍計入總轉介人次且歸入正常客

## 2. Competition Data Flow

- [x] 2.1 在 `src/types.ts` 與競賽排行相關型別中改為支援 `totalReferrals`、`hearingLossCustomers`、`normalCustomers` 等整合欄位
- [x] 2.2 在 `src/services/sheetsApi.ts` 新增 `fetchMultipleSpreadsheets()` 工具函式，支援以 `Promise.all` 平行抓取多份已儲存試算表
- [x] 2.3 在 `src/hooks/useGoogleSheetData.ts` 新增競賽專用 state（`selectedCompetitionSpreadsheetIds`、`competitionSheetData`、loading / error），並支援主試算表存在於 saved list 時的預設選取
- [x] 2.4 在分析層抽出或新增競賽專用聚合函式（pure function），將多份 `來客紀錄` 合併後依轉介門市計算總轉介人次（僅限轉介門市欄位有效值的資料列），並依 PTA 閾值拆分聽損客 / 正常客（僅看聽力）；PTA 缺失時直接計入正常客，並包含以 `header.includes()` 做欄位模糊比對的正規化邏輯

## 3. UI Integration

- [x] 3.1 調整競賽專用試算表複選 UI，讓它只在競賽排行區塊顯示，並提供清楚的專用提示文案與選取上限 10 份
- [x] 3.2 更新 `src/App.tsx`，讓競賽排行讀取獨立的整合結果，而總覽分析與區間比較維持既有單一試算表資料流
- [x] 3.3 更新 `src/components/CompetitionRanking.tsx`，改為顯示整合後的總轉介人次、聽損客、正常客與必要的資料範圍說明；對缺少必要欄位的試算表顯示錯誤提示

## 4. Verification

- [x] 4.1 執行自動化測試，確認多試算表整合與既有報表回歸測試通過
- [x] 4.2 以至少兩份已儲存試算表手動驗證同一轉介門市跨中心加總後的名次正確
- [x] 4.3 驗證更動競賽複選範圍時，總覽分析與區間比較的畫面與數字不會被覆蓋
- [x] 4.4 驗證初始進入競賽排行時主試算表已預設選中且排行正常顯示
