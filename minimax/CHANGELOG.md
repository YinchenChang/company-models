# CHANGELOG

每次同步記錄：版本、工作包、commit／PR 與變動摘要。

## v0.2 — 模型更新、P×Q 回本比值、成品（2026-10-22；PR #75；報告 `docs/reports/20261022_minimax_v0.2.md`）
- 合併 M2（資料更新）進本分支；SRC_MM 新增 S078–S112 共 35 筆（市場、共識、合約負債、FY2024–25 現金流量、匯率、M3.1 Flash 牌價、中國雲租價、消費端付費人數與 ARPPU）；S052、S062、S077 標「已更新」不刪。Valuation 改用 10-21 推算收盤價 HK$171.9（市值 $9.8B→$7.7B）。
- 新頁 `PxQ`（比照 OpenAI v0.6.1）：回本比值＝營收 ÷（年均上線 MW × 每 MW 年持有成本）；FY2025→2030：0.40→0.87。比值與合約價 I02 成正比（X26–X28）。
- Compute 第七節 D2：中國雲租價證據 ÷ TK 持有成本（I02 維持 12）；第八節 η 對照（TK 無 IF_ 落差分解名稱）；第九節自建集群（找不到規模→不建，註明）。Revenue 第五節消費端付費人數 × ARPPU 對照（僅 9M25）；Revenue 加 Bloomberg 共識等對照；Funding 加 FY2025 經營現金流對照。
- D8：M3.1 Flash（2026-10-19，較 M3 −50%）只更新 2026：新增 I30、I31（Assumed 0.20）、Demand D09；η 仍以降價前 I07 校準；I31＝0 時與 M1 逐 ID 相同（Valuation 除外）。2030 覆蓋率 0.5088→0.5064，累計需新融資 3,159→3,249。
- 新增 Inputs I30、I31、I33；Checks K09–K13；Sensitivity 31 情境（加回本比值欄、M3.1、I02 證據價）。
- HTML：P×Q 區、D2 與 M3.1 表、資料與版本區；每個數字標記並由 `scripts/check_html.py` 對照 Excel（納入 `build_all.sh` 與 CI）；無外部引用。
- 成品 `dist/20261022_MiniMax收支模型_v0_2.xlsx／.html`；v0.1 移 `dist/_待刪除/`。新增 `docs/handoff/MiniMax_handoff.md`、`docs/plan/20261022_MiniMax_命題與驅動對照表.md`、`scripts/xlvals.py`、`scripts/make_v02_compare.py`。

## v0.2-M1 — 流程補齊＋Tokenomics 官方快照取數、同步至 v19.0（2026-10-22；PR #75；報告 `docs/reports/20261022_minimax_M1.md`）
- 新增 `CLAUDE.md`、`CHANGELOG.md`、進度檔、`data/tokenomics_names.txt`、`scripts/tk_snapshot.py`、快照 v5.24／v19.0、`scripts/tk_attribution.py`、`scripts/ci_check.py`、CI `.github/workflows/minimax-build.yml`。`build_xlsx.py` 的 TK 改讀快照；TK_Link 加檔名／提交／SHA-256／讀取日與 IF_HoldEcon 全 15 欄。
- 以 v5.24 快照建置與 v0.1 逐格相同（2,396 格僅 TK_Link!A2 說明文字不同）。
- 同步 Tokenomics v19.0（`1916a06`，SHA-256 `d1728705…`）：IF_TokGW_* +34%～+57%、η 2025 3.42→2.25、1H26 0.80→0.52；覆蓋率 2030 0.5085→0.5088、累計需新融資 3,167→3,159 $M。`dist/` 未覆蓋。

## v0.1 — 首版（2026-10-08；PR #25）
- `dist/20261008_MiniMax收支模型_v0_1.xlsx／.html`：541 個公式，重算錯誤 0，Checks 合計 0；SRC_MM 77 列、Inputs 29 項。
- Tokenomics 數值：v5.24（master `bdb0de7`）12 格，**經 OpenAI v0.6 TK_Link 頁手抄**寫死於 `build_xlsx.py`（二手；無快照檔、無重現檢查）。
- 2026-10-08 PR #36（Tokenomics 連接審視）：只審視不改模型；以 v5.29（`a5061d9`）重取 12 格全部相同（差異 0）；列出「改用官方快照」等建議，成為 v0.2 M1 的依據。
