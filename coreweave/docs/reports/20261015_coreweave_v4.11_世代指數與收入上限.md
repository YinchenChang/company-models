# CoreWeave v4.11：VR200 世代指數（q × p）與收入上限四層（W9；v4.10 → v4.11）

2026-10-15｜工作單 W9（`docs/workorders/20261015_coreweave_W9_世代指數與收入上限.md`）｜PR #55｜成品：`dist/20261015_CoreWeave收支模型_v4_11.{html,xlsx}`（尚未產生）

## 目前狀態
- 完成：步驟 0（快照換 v5.36、新增 8 個名稱；原 29 名 365 個數值差異 0；`--check` 通過；Excel 取數頁與具名範圍）。
- 下一步：步驟 1（收入上限四層，比較層 `_Prod`）→ 步驟 2（q × p 結構）→ 升版 v4.11 → verify 三組。
- 需要 Andy 做的事：合併前用真正的 Excel 開啟 v4.11（尚未產生）。

## 輪次紀錄
### 步驟 0：換快照 v5.36（工程類）
- `data/tokenomics_names.txt` 新增 8 個名稱；以 `import_tokenomics.py` 產生 `data/tokenomics_snapshot_v5.36.json`（Tokenomics master `63db13a779cf511504ee7b13bc752cd4623ddda1`，`model/CURRENT`＝`20261015_Tokenomics_v5.36.xlsx`）；`git rm` v5.35 快照。
- 逐值比對原 29 名：365 個數值、文字欄位差異 0。`IFW_RevGWFleet_Util` 與 `IF_RevGWFleet` 逐值相同（所以 `_Util` 層＝現用層）；`_Prod` 與 `_Life` 在 v5.36 數值相同。
- `company.json` → `tokenomics`：snapshotFile／version／commit／names／`_note` 更新。
- Excel「Tokenomics_取數」：新名稱建 `TK_` 具名範圍（IFW_ 前綴去掉，例：`TK_RevGWFleet_Prod_GB300`、`TK_TokMW_Gen_ratio_VR200`）；`IFC_Use` 用途文字顯示在 IFW_ 四層列的「Tokenomics 位置」欄旁；`check_tokenomics_tab.py` 同步支援 IFW_ 與 by_name。
