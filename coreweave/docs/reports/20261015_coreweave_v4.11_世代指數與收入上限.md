# CoreWeave v4.11：VR200 世代指數（q × p）與收入上限四層（W9；v4.10 → v4.11）

2026-10-15｜工作單 W9（`docs/workorders/20261015_coreweave_W9_世代指數與收入上限.md`）｜PR #55｜成品：`dist/20261015_CoreWeave收支模型_v4_11.{html,xlsx}`（尚未產生）

## 目前狀態
- 完成：步驟 0（快照換 v5.36、新增 8 個名稱；原 29 名 365 個數值差異 0；`--check` 通過；Excel 取數頁與具名範圍）；步驟 1（上限檢查四層，比較層 `_Prod`；快速建置 cmp31 基準情境全部 OK，完整 verify 於最後一併跑）。
- 下一步：步驟 2（q × p 結構，預設須 0 差異）→ 升版 v4.11 → verify 三組。
- 需要 Andy 做的事：合併前用真正的 Excel 開啟 v4.11（尚未產生）。

## 輪次紀錄
### 步驟 0：換快照 v5.36（工程類）
- `data/tokenomics_names.txt` 新增 8 個名稱；以 `import_tokenomics.py` 產生 `data/tokenomics_snapshot_v5.36.json`（Tokenomics master `63db13a779cf511504ee7b13bc752cd4623ddda1`，`model/CURRENT`＝`20261015_Tokenomics_v5.36.xlsx`）；`git rm` v5.35 快照。
- 逐值比對原 29 名：365 個數值、文字欄位差異 0。`IFW_RevGWFleet_Util` 與 `IF_RevGWFleet` 逐值相同（所以 `_Util` 層＝現用層）；`_Prod` 與 `_Life` 在 v5.36 數值相同。
- `company.json` → `tokenomics`：snapshotFile／version／commit／names／`_note` 更新。
- Excel「Tokenomics_取數」：新名稱建 `TK_` 具名範圍（IFW_ 前綴去掉，例：`TK_RevGWFleet_Prod_GB300`、`TK_TokMW_Gen_ratio_VR200`）；`IFC_Use` 用途文字顯示在 IFW_ 四層列的「Tokenomics 位置」欄旁；`check_tokenomics_tab.py` 同步支援 IFW_ 與 by_name。

### 步驟 1：收入上限檢查改四層（比較層 `_Prod`）
- 「每MW經濟性」W4 區新增 4 列「客戶每 MW 付費 token 營收｜_100／_Util／_Prod／_Life」（`IFW_RevGWFleet_*`，在役世代加權）與 4 列「CRWV 每 MW 計費收入 ÷ _層」；原「IF_RevGWFleet」列保留（數值與 `_Util` 層逐值相同）。原「上限檢查｜CRWV 每 MW 計費收入 ÷ 客戶付費 token 營收」列改取比較層。
- 比較層放 `company.json` → `methodology.checks.revCapLayer`＝`"Prod"`（省略＝`Util`，即 W4 原做法）；Excel、HTML 引擎、README 欄位表都讀此欄，程式不寫死。
- 「檢查_連動」收入上限列標籤加「比較層 _Prod」，B 欄＝比較層各期最高；HTML 連動檢查同一則（title／detail 帶層名，只是把 JS 原有同式計算換成比較層名稱，沒有新增計算邏輯）。
- 上限比例（快速建置，基準情境）：FY26 61.9%（警示）、FY27 49.8%、FY28 41.0%、FY29 36.1%、FY30 32.5%。完整四層表見最終版。
