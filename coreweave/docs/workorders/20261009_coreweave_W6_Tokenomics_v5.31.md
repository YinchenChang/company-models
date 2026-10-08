# 工作單 CoreWeave W6｜接 Tokenomics v5.31（GB300 機架價格、IT 維護機齡兩段）→ v4.8（2026-10-09，r0，chat 端）

- 來源：W5 報告「建議 Tokenomics 修訂」→ Tokenomics v5.31（PR #36 已合併，master `f16f161`）；Andy 2026-10-08「全部同意」「不反對，請繼續」。
- 前置：v4.7 已合併（main）。必讀：共同規則、`coreweave/CLAUDE.md`（已決定事項 12–15）、W5 報告 `docs/reports/20261008_coreweave_v4.7_公司實況驗證.md`、Tokenomics `docs/reports/20261008_v5.31.md`。
- 分支：從最新 main 開 `claude/coreweave-w6-v4.8`；PR 標題「CoreWeave v4.8：接 Tokenomics v5.31」。

## 步驟（每步：verify → 進度檔（新增 W6 段落）→ commit → push）
0. 開分支、draft PR。更新 Tokenomics 唯讀副本到 master（v5.31）。
1. **重抓快照** `data/tokenomics_snapshot_v5.31.json`（移除舊快照），名稱清單加 `IF_MaintITWarr`、`IF_MaintITPost`、`IF_WarrantyYrs`；更新 `tokenomics` 區段；報告列出所有引用名稱的前後值。
2. **IT 維護改依機齡**：每期 IT 維護＝Σ_世代 [保固期內在役 MW × `IF_MaintITWarr` ＋ 保固期滿在役 MW × `IF_MaintITPost`]。機齡由 W2 的世代組合與各期新增 MW 推得（6/30 期初機隊的年份分層沿用 W1：2024 年底前 360 MW、2025 年新增 490 MW、2026 上半年新增 650 MW）；保固年限讀 `IF_WarrantyYrs`。Excel＋JS 雙引擎、cmp31 新增比對項。舊的「等值費率」`IF_MaintIT` 留為對照列。
3. **公司實況驗證頁重跑**（已決定事項 15）：各參數 Tokenomics 新值對 CRWV 實際的差距、是否仍 > 10%、處理（規則 1–4）。特別列出：每 MW 資本支出（Tokenomics 新值對 1H26 實際 24.8）、營運成本（對 Q2 1.36）。
4. **MW 口徑再查一次**（只查，不在沒有原文定義時改）：CoreWeave 10-K（FY2025）全文中「critical IT」「active power」「contracted power」的原文段落（EDGAR 全文或文件頁）。找到明確定義就依其設定 `meta.mwBasis` 並列入變動拆解；找不到則維持 IT，報告寫明試過的來源。
5. **升版 v4.8**：VLOG、`dist/` 換 v4.8、交接檔、README；對 v4.7 成品 `--vs-dist`（expect 只含成本相關列、Tokenomics 取數頁與其下游）；legacy 回歸維持對 v4.6 0 差異。變動拆解：① GB300 機架價格（資本支出、折舊、稅險）→ ② IT 維護機齡兩段 → ③ MW 口徑（若有改）；三情境。
6. **對照報告**：`docs/reports/20261009_coreweave_v4.8_Tokenomics_v5.31.xlsx`＋md：v4.7 對 v4.8 每 MW 收入、各成本項、EBITDA 率、資本支出（FY26–FY30、三情境）、FY26 對 Q2 實際對帳、三情境目標價與融資缺口、變動拆解、敏感度（k_長約 0.89、新約漲價 25%、Tokenomics 低／高成本）。結論先行。
7. verify 三組全過；PR 回報；「需要 Andy 做的事」：合併前用真正的 Excel 開啟 v4.8。

## 不做
收入口徑（k、隨需比例）不動；續約價格衰退、L 係數不做。
