# CoreWeave v4.12：Tokenomics 快照換 v18.0、世代門檻取 Tokenomics 名稱（W10；v4.11 → v4.12）

2026-10-21｜工作單 W10（`docs/workorders/20261021_coreweave_W10_Tokenomics_v18.0.md`）｜PR #58

## 目前狀態
- 完成：步驟 1（快照換 v18.0：原 37 名中 26 名數值不變、11 名改變，與工作單相符；新增 6 個門檻名稱；`import_tokenomics.py --check` 通過）；步驟 2 程式（判讀列改取 Tokenomics 門檻、上市間隔改推導、核對列與檢查頁警示、敏感度重算；VLOG v4.12；升版清單）。
- 下一步：verify 三組（進行中）→ `dist/` 換 v4.12、交接檔、README → 報告補完。
- 需要 Andy 做的事：合併前用真正的 Excel 開啟 v4.12。

## 輪次紀錄
### 步驟 1：換快照 v18.0
- `data/tokenomics_names.txt` 新增 6 名（`L1_GenPxThr_{Rev,Pay}_VR200` 與 `_Lo`／`_Hi`）；產生 `data/tokenomics_snapshot_v18.0.json`（Tokenomics master `868e2cc`、`20261021_Tokenomics_v18.0.xlsx`），`git rm` v5.36 快照。
- 逐值比對原 37 名：26 名差異 0；變動的 11 名＝`IF_RevGWFleet`、`IFW_RevGWFleet_100／_Util／_Prod／_Life`、`L1_TokMW_Gen_ratio_VR200` 與 `_Lo／_Hi`、`L1_RevGW_Fleet_VR200`、`L1_GPUhr_GB300_vsBE`（外部對照值）、`IFC_Use`（只有來源範圍列號 S6:S315 → S6:S344，用途文字不變）。成本、資本支出、折舊、維護、電力、`IF_HoldEcon`、`IF_GPUhrEcon`、`IF_Util`、`IF_WarrantyYrs` 全部不變。

### 步驟 2：世代門檻改取 Tokenomics、上市間隔改為推導值（程式完成，verify 進行中）
- `company.json` → `pricing.genIndex`：刪除 `gapYears`（固定值 1.18），新增 `thrRevName`／`thrPayName`（Tokenomics 名稱）；`methodology.checks.genThrGapMax`＝0.005（0.5 個百分點）。README 欄位表同步（`fields_doc.py`）。
- Excel：輸入頁上市間隔改活公式 `=LN(1/q)/LN(1+收入持平門檻)`；「每MW經濟性」判讀區改讀 `TK_GenPxThr_*`（基準＋低端＋高端各三列）、W9 自算公式保留為『核對｜』列、新增『核對差距』列；『檢查_連動』新增『世代門檻核對』列。HTML（segA.js）同式，快照缺門檻名稱時間隔與門檻顯示「不適用」、敏感度不產生；`scripts/permw_sens.py` 同步。
