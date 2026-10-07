# 工作單 CoreWeave W1｜Tokenomics 取數層（2026-10-07，r0，chat 端）

- 前置：W0 已合併（`coreweave/` 為 crwv-model `01b13ad`／v4.5 的原樣副本，verify 21 項全過）。
- 必讀：`20261007_coreweave_共同規則.md`。
- 分支：`claude/coreweave-w1-tokenomics`；PR 標題「CoreWeave W1：Tokenomics 取數層（數字不變）」。
- **本張不改任何模型數字、不改 `dist/`、不升版。** 驗收：`coreweave/scripts/verify.sh --vs-dist` 全過（新增的「Tokenomics_取數」分頁列為新增，不算差異）。

## 目標
1. 建立可重複的取數機制：Tokenomics 的 `IF_`／`L1_` 值 → 版本固定的快照檔 → `company.json` 的 `tokenomics` 區段指向它 → Excel「Tokenomics_取數」分頁顯示。
2. 產出 CRWV 欄位 ↔ Tokenomics 名稱對照表，給 W2 直接使用。
3. 查清 CRWV MW 的口徑（IT 關鍵電力或設施電力），以及世代組合與 GPU 小時價格的資料，給 W2 當輸入。

## 步驟（每步完成：進度檔＋commit＋push）
0. 開分支、draft PR、進度檔新增「W1」段落。唯讀 clone Tokenomics（共同規則第 8 節），記下 `model/CURRENT` 與 HEAD SHA。

1. **取數工具 `tools/tokenomics/import_tokenomics.py`（repo 根目錄，供各公司共用）**
   - 參數：`--tokenomics <clone 路徑>`、`--names <名稱清單檔>`、`--out <快照 json>`。
   - 讀 `model/CURRENT` 指向的 xlsx；只接受 `IF_`（非 `IF_Hdr*`）與 `L1_` 名稱，其他名稱報錯。讀取快取值（`data_only=True`）；若快取值缺漏，先以 LibreOffice headless 重算副本再讀。
   - 世代 × 成本情境範圍：依 `IF_HdrGen`、`IF_HdrCost` 表頭拆成 `{世代: {低成本, 基準, 高成本}}`；單格名稱存單值。
   - 快照格式：`{"source": {"file", "current", "commit", "extractedAt"}, "items": {名稱: {"label", "unit", "ref": "Interface!C9:Q9", "values": …}}}`。label、unit 取 Interface 該列 A、B 欄。
   - 加 `--check <快照>`：重新讀取並比對，任何值差異即非零結束（相對誤差 1e-9）。
   - 附 README 段落（工具用途、用法、名稱規則），寫在 repo 根目錄 README「規則」之下。

2. **CoreWeave 名稱清單與快照**
   - `coreweave/data/tokenomics_names.txt`：至少含 `IF_GPUsPerGW`、`IF_FacilityGW`、`IF_CapexIT`、`IF_CapexFacility`、`IF_CapexTotal`、`IF_HoldAcct`、`IF_HoldEcon`、`IF_GPUhrEcon`、`IF_PowerCost`、`L1_FacCapexMW`、`L1_GPUhr_GB200_vsCW`、`L1_GPUhr_GB300_vsBE`、`L1_RevGW_Fleet_VR200`。
   - 產出 `coreweave/data/tokenomics_snapshot_<版本>.json`（例：`tokenomics_snapshot_v5.24.json`）。
   - `company.json` 新增 `tokenomics` 區段：`snapshotFile`、`version`、`commit`、`names`（清單）、`_note`（說明：只引用 IF_／L1_；主值取基準成本情境）。`scripts/fields_doc.py --write` 同步 README 欄位表。

3. **Tokenomics v5.25 新名稱（chat 端另開工作單，同時進行）**
   Tokenomics v5.25 預計在 Interface 新增 DC_Cost 構件的名稱（每 GW、依世代 × 成本情境）：`IF_DeprLifeIT`（IT 折舊年限）、`IF_DeprIT`、`IF_DeprFac`、`IF_AvgDraw`（平均用電 ÷ 配電設計功率）、`IF_PowerPrice`（電價 $/kWh）、`IF_MaintIT`、`IF_MaintFac`、`IF_StaffSW`（人員、軟體、水與耗材）、`IF_TaxIns`（財產稅與保險）、`IF_OpexGW`（營運費用小計，不含折舊）。
   - 把這些名稱加進清單；取數工具遇到「名稱不存在」時，若在 `--optional` 清單內，就在快照記為 `{"missing": true}` 並警告，不中斷。
   - 若你結束前 Tokenomics `master` 的 `model/CURRENT` 已是 v5.25，改以 v5.25 產生快照（檔名改 v5.25），刪除 v5.24 快照；否則保留 v5.24 快照，在進度檔「下一步」寫明 W2 第 0 步要重抓。

4. **Excel「Tokenomics_取數」分頁（`build_xlsx.py`）**
   - 放在「來源」相關分頁附近；欄：名稱｜中文標籤｜單位｜世代｜低成本｜基準｜高成本｜Tokenomics 位置｜版本與 commit。值為藍字輸入格（取自快照），每個名稱的「基準」值格同時建 Excel 具名範圍 `TK_<名稱去掉 IF_／L1_>_<世代代碼>`（世代代碼：H100、GB200、GB300、VR200、RU；單值名稱不加世代），供 W2 公式引用。
   - 本張**不得**讓任何既有公式引用這些格。HTML 本張不顯示。
   - `verify.sh` 新增一步：`python3 ../tools/tokenomics/import_tokenomics.py --check`（Tokenomics clone 不存在時印警告並略過，不算失敗），以及「快照值＝Excel 分頁值」檢查（必做）。

5. **對照表（給 W2）**：`coreweave/docs/plan/20261007_coreweave_Tokenomics對照表.xlsx`（Andy 讀 Excel）＋同內容 md。欄：CRWV 欄位（company.json 路徑與 Excel 列名）｜現值（各期）｜單位與 MW 口徑｜現值依據（引用 README／交接檔原文）｜Tokenomics 名稱｜Tokenomics 基準值（換成 CRWV 單位，例如 $M/MW）｜差距 %｜W2 處理（引用／推導／對照／不動）。至少涵蓋：`costMW`、`gpuLife`、`m.power`、`m.pue`、`m.maint`、`m.revMW`、`ebStart`／`ebSteady`、`m.util`、站點每 MW 租金（`methodology.checks.rentVsBenchMin` 所用的市場基準）、`methodology.checks.capexPerMwBand`。

6. **資料蒐集（給 W2；全部依共同規則第 4 節記錄，存 `coreweave/data/permw_inputs_20261007.json`＋報告段落）**
   a. **MW 口徑**：CoreWeave「active power」「contracted power」的定義（10-K／10-Q／S-1／法說原文）：IT 關鍵負載或設施電力。找不到明確定義時，寫明試過的來源，預設視為 IT 關鍵電力 [Assumed]，並在報告列出若改為設施口徑（÷ IF_FacilityGW 基準 1.2）對每 MW 數字的影響。
   b. **世代組合**：6/30 在役機隊的世代結構（Hopper／GB200／GB300），與 2H26–FY27 新增產能的世代（公司說法、NVIDIA 出貨、第三方報導）。期初組合為 [Analogy]／[Assumed] 時給區間。
   c. **GPU 小時價格（每世代）**：H100、GB200、GB300、（若有）VR200 的長約與隨需價格，至少兩個獨立來源（SemiAnalysis 不可單獨引用）；CoreWeave 自家牌價可用但標 [Interested-party]。記錄合約期間、是否含網路與儲存。
   d. **公司營運費用實際數**（10-Q 損益表）：Q2 營收成本、技術與基礎設施、銷售行銷、一般管理，以及其中的 D&A、SBC、租金（若揭露），供 W2 計算管銷費用率與核對由下而上成本。
   e. **站點電價**（若有揭露，例如 PPA）只作對照，不作輸入（電價以 Tokenomics 為準）。

7. **整體 verify 與回報**：`verify.sh --vs-dist` 全過；PR 回報：取數工具用法、快照版本、對照表摘要（差距最大的 5 項）、MW 口徑結論、世代組合與 GPU 小時價格的資料是否足以支撐 W2 的「GPU 小時價格路線」（不足就寫明，W2 會改用備案）。

## 驗收（chat 端逐項核）
- verify 全過，既有數字 0 差異。
- 快照可由 `--check` 重現；Excel 分頁值＝快照值。
- 對照表每列都有 Tokenomics 名稱或寫明「Tokenomics 無對應（公司專屬）」。
- 第 6 步每筆資料有來源、日期、標記；沒有以公司 ARR 或 CapEx 指引反推任何參數。
