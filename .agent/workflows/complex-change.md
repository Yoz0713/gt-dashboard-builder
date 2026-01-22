---
description: 複雜架構變更的標準作業流程
related_skill: complex-change-methodology
---

> 📚 **配套 Skill**: 此 Workflow 有配套的 [complex-change-methodology](./../skills/complex-change-methodology.md) Skill，提供變更評估的思考框架。

## 階段 1：RAG 找入口

1. 確認變更需求（例如：修改聽力圖顯示邏輯）
2. 使用語意搜尋找出相關檔案
3. 記錄入口檔案（Primary Target）

## 階段 2：工具找半徑

// turbo
4. 執行影響分析：`npx tsx scripts/scan_impact.ts <target>`
5. 記錄相依性 (Dependencies) 與受影響檔案 (Dependents)
6. 根據依賴關係，決定 Context 需要包含哪些檔案

## 階段 3：TDD 開發

7. 在 `test/` 目錄查找是否有現有測試，若無則建立
8. 撰寫失敗的測試案例（Red）來定義預期行為
9. 實作功能使測試通過（Green）
10. 若有大量依賴，執行相關依賴模組的測試以確保無回歸
// turbo
11. 執行完整測試：`npm run test:run`

## 階段 4：Agentic 驗證 (Browser Check)

12. 啟動開發伺服器：`npm run dev`
13. 呼叫 `browser_subagent` 進行視覺驗證
    - 任務提示： "Navigate to localhost:5173, use 'Dev Fill' to identify login fields (do not submit if creds are placeholder), and verify [Feature X]"
14. 截圖確認最終狀態
