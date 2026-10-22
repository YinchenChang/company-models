# MiniMax 收支模型（company-models/minimax）

MiniMax（00100.HK，上海稀宇科技）收支模型 v0.1（2026-10-08）。比照 `openai/`（v0.6）的分層與 MW 主軸建立；MiniMax 為上市公司，另加 Funding 與 Valuation。

命題：MiniMax 每 MW 算力的營收能否覆蓋每 MW 全成本；若不能，缺口何時出現、由誰、以什麼條件補。FY2025–FY2030，曆年制。

- Excel 為唯一計算引擎：`build_xlsx.py` 只產生結構與公式（藍字輸入、黑字公式），數值由 LibreOffice 重算。
- 一鍵建置：`scripts/build_all.sh [日期] [版本]` → `dist/<日期>_MiniMax收支模型_<版本>.xlsx／.html`（需 `libreoffice-calc-nogui`、openpyxl）。
- 敏感度：`scripts/sensitivity.py`（逐項改寫 Inputs、重算，寫入 Sensitivity 頁快照）。
- 原始數據與來源：`data/research/financials.md`、`operations.md`（研究代理 2026-10-08 蒐集；一手／二手已標註）。
- Tokenomics 取數：`scripts/tk_snapshot.py`（讀 `model/CURRENT` 具名範圍）、名稱檔 `data/tokenomics_names.txt`、指標檔 `data/tokenomics_snapshot_CURRENT`；歸因 `scripts/tk_attribution.py`；驗收 `scripts/ci_check.py`；CI `.github/workflows/minimax-build.yml`。
- 工作規範：`CLAUDE.md`；變更紀錄：`CHANGELOG.md`；進度：`docs/reports/20261022_minimax_進度.md`。
- 報告：`docs/reports/20261008_minimax_v0.1.md`（v0.1）、`20261022_minimax_M1.md`（v0.2 M1）。

| 頁 | 內容 |
|---|---|
| README | 命題、驅動→推導對照、圖例、已套用的預設 |
| SRC_MM | 公司財務與營運原始數據（77 列；標記、一手／二手、來源） |
| TK_Link | Tokenomics 官方快照（`data/tokenomics_snapshot_*.json`；現用 v19.0，含檔名、提交、SHA-256、讀取日）；12 格基準值＋IF_HoldEcon 全 15 欄 |
| Inputs | 假設 29 項（值、低、高、標記、說明） |
| Demand | API 實收單價、計費 token |
| Revenue | 企業端（API）、消費端（產品與收費方式按 9M25 結構示意） |
| Compute | 機隊組合、每 MW 產能、η 校準（2025、1H26）、推論／訓練 MW、VR 等值 |
| PnL | 損益（算力／人事／其他）、每 MW 營收 vs 每 MW 全成本 |
| Funding | 現金路徑、需新融資、稀釋、已投入資金 |
| Valuation | 市值、EV、同業、反向檢驗 |
| Checks | 8 項檢查＋錯誤值計數（皆為 0） |
| Sensitivity | 25 個單項情境（快照值） |
