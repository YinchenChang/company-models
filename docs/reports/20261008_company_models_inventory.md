# company-models 盤點（2026-10-08，chat 端）

Andy 2026-10-08：「先盤點一下現在 company models 到底做了幾個」。以 master `f373885` 為準；Tokenomics 連接狀態取自同日 `docs/reports/20261008_tokenomics_link_audit.md`（對照 Tokenomics v5.27）。

## 一、八個模型的現況

| 公司 | 版本（成品） | 成品檔（dist） | 工作單數 | 取數方式 | 釘住 Tokenomics | 主要缺口 |
|---|---|---|---|---|---|---|
| CoreWeave（CRWV） | v4.6；v4.7 建置中（PR #27 W4 未合併） | 20261008 v4_6 HTML／Excel | 5 | 官方快照工具 | v5.26 | W4「每 MW 收入以 Tokenomics 為錨」待審查合併；快照落後 v5.27→v5.29 |
| Nebius（NBIS） | v0.2 | 20261007 v0_2 | 6 | 手工 JSON（permw） | v5.24 | 每 MW 收入仍為成本加成與市場價 50/50 平均；JSON 直接引用 Inputs!／DC_Cost! |
| Oracle（ORCL） | v0.2（延誤與租賃槓桿） | 20261008 v0_2 | 5 | 手工 JSON（Nebius 副本） | v5.24 | 同 Nebius；容量與單價情境同向綁定 |
| WhiteFiber（WYFI） | v0.1 | 20261008 v0_1 | 4 | 手工 JSON（Oracle 副本） | v5.24 | 同 Nebius |
| OpenAI | v0.6（S6 成品） | 20261008 v0_6 | 8 | 無快照；builder 直接讀 Tokenomics 名稱 | 未釘版本 | `IF_Alloc*` 六格在 v5.29 改變、四層瀑布未取；需重取數 |
| Anthropic | v0.1（A5 成品） | 20261008 v0_1 | 6 | 無快照；沿用 OpenAI 架構 | 未釘版本 | 同 OpenAI |
| MiniMax（00100.HK） | v0.1 | 20261008 v0_1 | 0（無 workorders 目錄） | 無快照；build_xlsx 直接引用 | README 稱 TK v5.24 | 無工作單紀錄；取數路徑不透明 |
| 智譜（02513.HK） | v0.1 r2（Z5 成品） | 20261008 v0_1 | 6 | 無快照；沿用 OpenAI 架構 | 未釘版本 | 同 OpenAI |

PR：company-models 共 27 個，26 個已合併，1 個開啟（#27 CoreWeave W4）。

## 二、共同問題（由審視工具得出）

1. **版本全部落後**：沒有任何模型釘在 v5.27；四家 neocloud／雲平台釘 v5.24，CoreWeave v5.26，模型商未釘版本。v5.29 將改變 `IF_Alloc*` 六格並新增 `IFW_`／`IFC_`，全部模型都要重取。
2. **取數方式三種並存**：只有 CoreWeave 用 `tools/tokenomics/import_tokenomics.py` 的官方快照（可 `--check` 重現）；Nebius／Oracle／WhiteFiber 用手工 JSON（含 `Inputs!`、`DC_Cost!` 的直接引用，契約第 2 條第 1 項違規）；四家模型商 builder 直接讀名稱、無版本紀錄。
3. **沒有模型取四層瀑布或用途欄**（v5.29 才有，預期中）。
4. **每 MW 收入**：Nebius／Oracle／WhiteFiber 仍為 50/50 平均（Copilot MW 報告指出的問題）；CoreWeave W4 已改為「Tokenomics 錨 × 公司因素 k」，與下游契約第 4 條一致，可作為其他三家的範本。
5. MiniMax 沒有工作單目錄，流程紀錄不完整。

## 三、建議順序

1. 合併 Tokenomics v5.29 → 全部模型執行 `docs/workorders/20261008_tokenomics_link_review.md`（審視，只回報）。
2. 審查 CoreWeave W4（PR #27）並合併；以 W4 為範本開 Nebius／Oracle／WhiteFiber 的「每 MW 收入改接 Tokenomics 錨」工作單（含停用 50/50、分離容量與單價情境軸）。
3. 四家模型商改用官方快照工具（names.txt + snapshot JSON + `--check` 進 CI），重取 v5.29 的 `IF_Alloc*` 與四層瀑布。
4. MiniMax 補 `docs/workorders/` 與進度檔。
