## 0. Test Infrastructure

- [x] 0.1 安裝 `Jest` 與 React Testing Library 相關依賴（如 `@testing-library/react`、`@testing-library/jest-dom`），沿用 CRA 相容的測試路徑
- [x] 0.2 配置 TypeScript 測試環境（tsconfig、test script），確保可執行基本測試

## 1. Test Baseline

- [x] 1.1 為 `src/utils/analysis.ts` 建立門市轉介排行的樣本資料測試，先覆蓋門市轉介篩選、潛力客計數與非潛力客計數
- [x] 1.2 先寫排序規則測試，驗證潛力客數優先、非潛力客數次之，以及同條件時順序穩定（第三排序條件為門市名稱字典序）
- [x] 1.3 為既有總覽分析與區間比較補上基本回歸測試，鎖定新增 tab 前後的核心顯示行為

## 2. Analysis Model

- [x] 2.1 在 `src/types.ts` 新增 `CompetitionRankingEntry` 型別（含 `storeName`、`potentialCustomers`、`nonPotentialCustomers`、`rank`），並在 `src/utils/analysis.ts` 的 `AnalysisResult` 介面新增 `competitionRankingAnalysis` 欄位
- [x] 2.2 於既有門市轉介資料處理流程中計算每個轉介門市的潛力客數與非潛力客數
- [x] 2.3 在分析層套用競賽排序規則：先依潛力客數降冪，再依非潛力客數降冪，並以門市名稱字典序作為穩定第三排序條件

## 3. UI Integration

- [x] 3.1 新增競賽排行元件 `src/components/CompetitionRanking.tsx`，顯示名次、轉介門市、潛力客數與非潛力客數
- [x] 3.2 更新 `src/App.tsx`：將 `activeTab` 類型從 `'overview' | 'comparison'` 擴展為 `'overview' | 'comparison' | 'competition'`，加入「競賽排行」tab 按鈕並串接 `competitionRankingAnalysis` 資料
- [x] 3.3 確保競賽排行沿用既有日期區間、PTA 閾值與載入狀態呈現
- [x] 3.4 讓 UI 測試覆蓋三個報表 tab 的切換，確認總覽分析與區間比較未受影響

## 4. Verification

- [x] 4.1 執行自動化測試，確認競賽排行邏輯與既有報表回歸測試全部通過
- [x] 4.2 以包含門市轉介資料的試算表手動驗證排名與潛力客/非潛力客統計是否正確
- [x] 4.3 驗證潛力客數相同時是否依非潛力客數正確排序，且同條件資料顯示順序穩定
- [x] 4.4 執行前端 build 驗證，確認新增報表不造成型別或編譯錯誤
