# tools/target_summary — AI半導體 目標價總表產生器

`python3 tools/target_summary/build_summary.py <輸出.xlsx> [--prev <前版總表.xlsx>]`，再以 LibreOffice headless 重算。成品放 Drive `03-AI半導體/選股與資料/`（只留最新版）。

| 檔案 | 內容 |
|---|---|
| `build_summary.py` | 讀各模型 `dist/*.xlsx` 的 Val 頁（標籤同 `markets/invest_map/extract_models.py`）與 README 頁重點觀察 |
| `universe.csv` | 54 家：資料夾、公司、代號、次產業、市場、本公司代號（`own_ticker`）、v0.2.3 對照值 |
| `peers.csv` | 各模型 Comparable 同業（抄自 Val 頁同業表，2026-10-27） |

工作表：目標價總表（v0.2.4 新增「與前版相比」與「穩健度：共識 EPS」欄）、投資建議統計、市場別分布、相對倍數試算（未套用）、說明。

相對倍數調整（`tools/relative_multiple.py`、`tools/data/rel_mult_hist.csv`）：2026-10-27 以年底落後 P/E 試算後**不套用**；待 Bloomberg 前瞻倍數（抓取單 `Bloomberg抓取單_前瞻倍數_全部_20261027.xlsx`）取得後重算。
