## 1. Test Baseline

- [x] 1.1 在 `src/utils/analysis.test.ts` 新增診所回訪聚合測試，覆蓋依「診所名稱」分組、依總轉介人次降冪排序、同分時以診所名稱字典序穩定排序
- [x] 1.2 新增測試：依 PTA 閾值計算聽損個案數，並依成交判定計算配戴數與成交率
- [x] 1.3 新增測試：月度轉介趨勢 bucket、首次 / 最近轉介日期、轉介高峰月份
- [x] 1.4 新增測試：來源資料缺少診所名稱欄位時回傳空陣列；診所名稱為 `#N/A` / `#REF!` / 空值時該列不納入
- [x] 1.5 新增回歸測試：新增 `clinicFollowUpAnalysis` 後，既有 `clinicAnalysis`（Top 8 來源）與 `competitionRankingAnalysis` 輸出不變
- [x] 1.6 新增 `src/components/ClinicFollowUp.test.tsx`，覆蓋預設選中人次最高診所、搜尋過濾、切換診所後數字更新、空狀態文案、姓名遮罩切換
- [x] 1.7 更新 `src/App.test.tsx`，補上第四個 tab 的切換斷言與 `analysisResult` mock 的新欄位

## 2. Data Flow

- [x] 2.1 在 `src/types.ts` 新增 `ClinicPatientRecord`、`ClinicMonthlyPoint`、`ClinicFollowUpReport` 型別
- [x] 2.2 在 `src/utils/analysis.ts` 新增 `findClinicNameHeader()`，以優先序模糊比對定位診所欄位，找不到時回傳空字串
- [x] 2.3 在 `analyzeData` 內新增 `clinicFollowUpMap` 累積器，重用 `checkIsDealt` / `parseSheetDate` / `getEarPTA` / `getHearingLossDegree` / `parseAge` / `parseAmount` / `INVALID_STORE_NAMES`，且不改動既有 `clinicSummary`
- [x] 2.4 在 `AnalysisResult` 與其回傳物件補上 `clinicFollowUpAnalysis`

## 3. UI Integration

- [x] 3.1 新增 `src/components/ClinicFollowUp.tsx`：診所選取面板（搜尋 + 排序 + 清單）與空狀態
- [x] 3.2 完成報告本體：列印用報告抬頭、KPI 卡片、回訪重點、月度趨勢圖、聽損程度與年齡分布圖、客戶明細表與姓名遮罩
- [x] 3.3 加上「匯出此診所回訪報告 (PDF)」按鈕，沿用 `window.print()`
- [x] 3.4 更新 `src/App.tsx`：`activeTab` 加入 `clinic`、新增切換按鈕、將巢狀三元改為明確分支
- [x] 3.5 在 `src/index.css` 的 `@media print` 補上明細表跨頁樣式（重複表頭、單列不切斷）

## 4. Verification

- [x] 4.1 執行 `npm test -- --watchAll=false`，確認新測試與既有測試全部通過
- [x] 4.2 執行 `npm run build`，確認 TypeScript 編譯無錯誤
- [ ] 4.3 手動驗證診所清單列出區間內所有診所、預設選取、搜尋與排序、切換診所後 KPI 與圖表同步更新
- [ ] 4.4 手動驗證總覽分析、區間比較、競賽排行的數字與改動前一致
- [ ] 4.5 以瀏覽器列印預覽驗證 PDF 版面：選取面板不出現、報告抬頭含診所名稱與區間、明細表跨頁重複表頭且單列不被切斷

## 5. 轉介費用與列印修正

- [x] 5.1 新增 `calculateReferralFee()` 與預設 10% 回饋比例常數，並補上單元測試
- [x] 5.2 新增獨立的「轉介費用」區塊：KPI 摘要、可調整的回饋比例、逐筆已配戴個案的費用表與合計
- [x] 5.3 已配戴但未填成交金額的個案以 NT$0 計算並顯示補件提示
- [x] 5.4 客戶明細表加上成交金額與轉介費用欄位
- [x] 5.5 修正列印時 canvas 被頁面邊界裁切：`index.css` 加上等比縮放規則，並在列印事件觸發時主動 resize 圖表
- [x] 5.6 補上轉介費用區塊與費用計算的測試
- [ ] 5.7 以瀏覽器列印預覽確認直條圖不再跑版、不被邊緣切割

## 6. 匯出版面調整

- [x] 6.1 匯出報告標題改為「大樹聽力中心診所轉介名單分析」，移除「轉介成效回訪報告」字樣
- [x] 6.2 報告表頭與兩張表格的表頭改為藍底白字（`rgb(0, 140, 215)`），並確保列印時保留背景色
- [x] 6.3 回訪重點加上 `print:hidden`，不列入匯出報告
- [x] 6.4 無需給付轉介費用時，費用區塊不列入匯出報告（畫面仍顯示說明）
- [x] 6.5 為 `App.tsx` 的錯誤橫幅加上 `print:hidden`，避免原始錯誤訊息被印在報告最上方
- [x] 6.6 補上匯出版面相關測試
- [ ] 6.7 以瀏覽器列印預覽確認表頭配色、標題與隱藏區塊皆正確

## 7. 匯出版面精修

- [x] 7.1 匯出表頭移除分析區間
- [x] 7.2 新增 `formatStoreName()`，服務門市去掉工作表名「來客紀錄」只留門市，並補上單元測試
- [x] 7.3 匯出時轉介費用區塊隱藏四張統計卡片，只留明細表
- [x] 7.4 匯出時明細表合計列隱藏成交金額合計，保留轉介費用合計
- [x] 7.5 補上對應測試
- [ ] 7.6 以瀏覽器列印預覽確認表頭欄位與費用區塊內容符合預期

## 8. 列印樣式串接修正

- [x] 8.1 修正 `.print\:hidden` 被 `.grid-cols-1 { display: block !important }` 蓋掉的問題：兩者同權重且同為 `!important`，由後者勝出，因此把隱藏規則移到 `@media print` 區塊最後
- [x] 8.2 新增 `src/printStyles.test.ts`，直接解析 `src/index.css` 鎖住「隱藏必須勝過版面」的順序不變式（jsdom 不套用樣式表，元件測試抓不到這類問題）
- [x] 8.3 以還原舊順序的方式驗證新測試確實會失敗
- [ ] 8.4 以瀏覽器列印預覽確認轉介費用的四張統計卡確實不再出現
