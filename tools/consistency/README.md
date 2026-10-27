# 跨模型數字一致性檢查（API 版，2026-10-27 起）

只產出檢查結果，不改模型、公司檔案或網站；修正由各模型的維護串依工作單處理。

- **A. Tokenomics 快照落差**（純程式）：各資料夾 `tokenomics_snapshot_v*.json` 對最新快照的基準情境數值差異；格式不同者標「無法比對」。
- **B. 公司檔案 vs 模型儲存格**（API，Batch，Sonnet）：`markets/invest_map/profiles/*.md` 每句的出處標記 `[本專案：<檔名>.xlsx <工作表>!<儲存格>]`，程式讀出儲存格值（優先原引用版本；工作樹沒有就從 git 歷史取出），再請模型逐句判定一致／不一致／無法判定。
- **引用版本落後**：公司檔案引用的模型檔已有新版者，另頁列出，供 AI 投資地圖更新時參考。

執行：push 到分支 `claude/consistency`（工作流程 `consistency-check`）可測試；合併到 main 後每月 3 日自動執行，結果寫到分支 `consistency-results`。結果在 `tools/consistency/out/<日期>.xlsx／.md`。費用：首次約 US$1.2（Sonnet 5.5 Batch；單價 [Verified] platform.claude.com 價目頁 2026-10-27：一般 2／10、Batch 1／5 美元每百萬 token）。
