# CHANGELOG

每次同步記錄：版本、工作包、commit／PR 與變動摘要。

## v0.2-M1 — 流程補齊＋Tokenomics 官方快照取數、同步至 v19.0（2026-10-22；進行中，詳見 `docs/reports/20261022_minimax_M1.md`）
- 新增 `CLAUDE.md`、`CHANGELOG.md`、進度檔、`data/tokenomics_names.txt`、快照工具、CI（`.github/workflows/minimax-build.yml`）。PR #75。

## v0.1 — 首版（2026-10-08；PR #25）
- `dist/20261008_MiniMax收支模型_v0_1.xlsx／.html`：541 個公式，重算錯誤 0，Checks 合計 0；SRC_MM 77 列、Inputs 29 項。
- Tokenomics 數值：v5.24（master `bdb0de7`）12 格，**經 OpenAI v0.6 TK_Link 頁手抄**寫死於 `build_xlsx.py`（二手；無快照檔、無重現檢查）。
- 2026-10-08 PR #36（Tokenomics 連接審視）：只審視不改模型；以 v5.29（`a5061d9`）重取 12 格全部相同（差異 0）；列出「改用官方快照」等建議，成為 v0.2 M1 的依據。
