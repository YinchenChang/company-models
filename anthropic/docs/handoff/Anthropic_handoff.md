# Anthropic 收支模型 交接檔（權威版；放 repo）

> 新對話串：先讀本檔、規格 `docs/plan/20261008_Anthropic_v0.1_規格.md` 與進度檔 `docs/reports/20261008_anthropic_進度.md`，再以 git 確認狀態。

## 1. 命題與範圍
- 命題（與 OpenAI v0.6 同一句）：Anthropic 每 VR 等值 GW 的年營收能否覆蓋每 GW 年全成本；若不能，缺口要多少外部資金、由誰以什麼條件提供。FY2025–FY2030，曆年制，2025 為實際校準年。
- 來由：Andy 2026-10-08「請比照 OpenAI，做 Anthropic」；依 2026-10-07 授權，Claude 套用預設（規格 D1–D20）不先確認，v0.1 完成時彙總審查。

## 2. 位置
- repo `YinchenChang/company-models`，資料夾 `anthropic/`；建置分支 `claude/anthropic-v0.1`（單一 PR）。
- 成品：`dist/`；Drive `02-算力收支/Anthropic/`（chat 端放置）。

## 3. 模型狀態
- A0 完成（2026-10-08）：骨架（engine、tools、tests 基礎沿用 OpenAI v0.6）、規格 r0、共同規則、A1–A5 工作單。
- Tokenomics 快照：v5.26（master `3dd1216`），與 OpenAI v0.6 相同。
- A2 完成（2026-10-08，v0.1-A2）：`model/20261008_Anthropic_v0.1.xlsx`（README、SRC_ANT、TK_Link、OAI_Link、Inputs、Demand、Revenue、Checks）；2025 營收 4.60（校準）、2026 48.73（半校準）、2030 總額 214.9／淨額 196.8（$B）。報告 `docs/reports/20261008_v0.1-A2.md`（含「給 A3 的交接」）。
- A3 完成（2026-10-08，v0.1-A3）：Compute、Cost（命題表）。每 VR 等值 GW 差額 2026–2030 −6.5／−5.2／−2.9／＋5.1／＋11.2；覆蓋率 0.87 → 1.78。報告 `docs/reports/20261008_v0.1-A3.md`。
- **A4 完成＝v0.1 模型完成（2026-10-08）**：Funding、Reverse、Cost 第十二–十五節（逐家對帳 WARN、錨定招股書、機房租金、D22 量化）。命題 2：2026–2030 累計外部資金需求 0（2026 已交割 $100B 股權覆蓋 2026–2028 缺口 −40.1；現金谷底 2028 年 80.2）；OpenAI v0.6 為 442.8。前提：只計已簽合約（2029 起供給遞減）；兩個需求驅動同時取下緣才翻轉（2030 差額 −1.7、外部資金 12.7）。反向（管理層舊目標）累計 25.2。完成報告 `docs/reports/20261008_v0.1.md`（④A1–A4 預設彙總為 Andy 審查主體，☆ 5 項）。Checks：CHK_Errors＝0、CHK_Warnings＝1（C89 逐家合約，預期）。Tokenomics CURRENT 仍 v5.26（master 已到 `70d859e`）。
- 下一步：A5 成品（HTML 一頁摘要含 OpenAI 並排、命題與驅動對照表定稿、dist）；Andy 審查 v0.1 預設彙總。建議後續首項：新增合約機制（需 Andy 決定）。
