# OpenAI 模型｜Tokenomics 連接審視（v5.29）

- 分支：`claude/openai-tk-link-review`
- 底稿：company-models `main` @ `75327ef`；本報告的提交 SHA 見 PR 頁面（提交後才產生）
- 報告檔：`openai/docs/reports/20261008_openai_tokenomics_link_review.md`
- 依據：共用工作單 `docs/workorders/20261008_tokenomics_link_review.md`、`openai/docs/workorders/20261008_openai_tokenomics_link_review.md`；Tokenomics 下游資料契約 v0.1；Tokenomics master `a5061d9`（`model/CURRENT`＝`20261008_Tokenomics_v5.29.xlsx`，前提成立）
- **本輪只審視、不改模型**：company.json／builder／Excel／data 都沒有動。v5.29 快照只產生在暫存路徑 `/tmp/openai_v529_snapshot.json`，未提交。

## 〇、一句話結論

OpenAI 模型的取數方式是「重建時由 builder 直接讀 Tokenomics 現行檔的名稱，把值與版本寫進 Excel 的 TK_Link 頁」；現行 Excel 記錄的是 **v5.26（master `3dd1216`）**。對 v5.29 重取時，63 個名稱中 **只有 `IF_Alloc*` 六格變動**，而這六格在本模型**只用於對照列與檢查頁，不進入任何營收、成本、融資或 HTML 數字**——所以 v5.29 對命題結論的直接影響為 0。真正要處理的是契約新增的兩項要求：**取四層瀑布（`IFW_`）**與**用自身揭露校準服務端與研發端規模（`L1_ScaleServe`、`L1_Gap*`、`L1_ScaleRD`）**；本模型目前以一個「η＝0.28」的總係數吸收了服務端落差，尚未拆成四項口徑差。

---

## 一、取數點清單（共用工作單第 1 節）

### 1-1 取數方式與版本

| 項目 | 現況 |
|---|---|
| 取數程式 | `openai/builder/tk_link.py`（`read_snapshot`），由 `openai/builder/build.py --tk-dir <Tokenomics>` 呼叫 |
| 讀哪個檔 | `--tk-dir` 指向的 Tokenomics checkout 的 `model/CURRENT`（建置當下是哪一版就讀哪一版；builder 本身**不釘版本**） |
| 版本紀錄 | 有：Excel `TK_Link` 頁 B3–B6（具名範圍 `TK_File`、`TK_Version`、`TK_Commit`、`TK_ReadDate`），每列也記版本與 SHA；Checks C129／C130 顯示 |
| 現行 Excel 所記版本 | `20261007_Tokenomics_v5.26.xlsx`，master `3dd1216`，讀取日 2026-10-08（**落後 v5.29**） |
| 名稱選取方式 | 依樣式比對（`IF_TokGW_(Luna|Sol|Astra)`、`IF_Price…`、`IF_FullCost_…` 等）＋固定清單（`SRC_DEM_004–013`、`IF_Alloc*` 七名、`L1_Ans3`＋低高、`IF_CapexTotal`） |
| 名稱數 | 63 個（IF_ 48、SRC_DEM 12、L1_ 3）；另讀 Interface 第 4、5 列表頭作為世代與成本情境的鍵（`TK_HdrGen`、`TK_HdrCost`） |

> 盤點報告（`docs/reports/20261008_tokenomics_link_audit.md`）寫「取數方式：無；釘住 —」——**不精確**：本模型沒有用共用快照工具，但 Excel 內有版本紀錄（v5.26）。真正的缺點是 builder 不釘版本、也無「與記錄版本一致」的重現檢查（`tools/check_tk_snapshot.py` 只對「Tokenomics 現行版」比對，DIFF 只報警告）。

### 1-2 取數點清單表（只列進入公式或檢查的名稱；其餘只存放在 TK_Link 供參考，見 1-3）

| 本模型的格或參數 | Tokenomics 名稱 | 合規性 | 四層瀑布層 | 用途是否符合 `IFC_Use` | 版本 |
|---|---|---|---|---|---|
| Compute G03–G05（D10:I12）每 GW 總產出；→ G22–G35 每 GW 年產能 → G38–G49 推論 GW | `IF_TokGW_Astra`／`_Sol`／`_Luna` | 合規名稱；但**違反契約 2-2**（只取 100% 層，自行 × IF_Util） | 100%（自乘 IF_Util 後＝Util 層） | `IFC_Use`＝「情境、相對比較、反轉門檻；**上限，不得作預測**」。本模型不拿它當收入預測，而是把 token 換成 GW，再以 η（2025 校準）縮放——用途屬「換算」，基準情境下 2025 由支出校準覆蓋，**大致可接受**；但 η 回升情境（INP_258＝1，η→1）等於把 Util 層直接當生產產能，**不符**契約 2-5 | v5.26 |
| Compute G22–G35（每 GW 年產能）、Cost K92／K93 | `IF_Util` | 合規 | Util | 符合（情境用；契約第 3 節允許下游覆寫，本模型**未設自有覆寫格**） | v5.26 |
| Compute G06、G81、G136；Cost K92／K93 | `IF_TokGW_Sol`（VR 等值係數用） | 合規 | 100%（只取比值，層別不影響） | 符合（相對比較） | v5.26 |
| Cost K05–K07（D12:I14）每 GW 年持有成本 → 供應商持有成本、雲端毛利、經濟口徑算力成本 | `IF_HoldEcon` | 合規 | 不適用（成本） | 符合 | v5.26 |
| Cost K08；Compute G136 → 自建 GW（進命題分母） | `IF_CapexTotal` | 合規 | 不適用 | 符合 | v5.26 |
| Cost K92 TK 參考全成本（只作對照） | `IF_FullCostDefault_Sol` × `IF_TokGW_Sol` × `IF_Util` | 合規 | Util | 符合（對照） | v5.26 |
| Cost K93 TK 參考全成本（只作對照） | `IF_FullCost_Sol` × 同上 | 合規 | Util | 符合（對照） | v5.26 |
| Cost K69–K71（Q3 對照）；Checks C82 | `L1_Ans3` | 合規 | 不適用 | 符合（只列差距） | v5.26 |
| Demand D43、D45、D46（`DEM_TkRatio2025`）；Checks C37、C55 | `IF_AllocDemand` | 合規 | 不適用 | 符合（只列差距） | v5.26（**v5.29 變**） |
| Checks C54、C55 | `IF_AllocServeGW` | 合規 | Util（`IFC_Layer`＝IF_Util） | 符合（只並列） | v5.26（**v5.29 變**） |
| Checks C58 | `IF_AllocRDGW` | 合規 | 不適用 | 符合（只並列） | v5.26（v5.29 不變） |
| Checks C59 | `IF_AllocImpliedNk` | 合規 | Util | 符合（只並列） | v5.26（**v5.29 變**） |
| Compute G01、G02 的鍵；全部 SUMIFS；Checks C42 | Interface 第 4、5 列（＝`IF_HdrGen`、`IF_HdrCost` 的內容） | **形式上不符**契約 2-1（`IF_Hdr*` 禁止連結；且以工作表位置直接讀）；實質只當查表鍵，共用快照工具也同樣讀表頭 | 不適用 | 不適用 | v5.26 |

### 1-3 只存放、沒有公式引用的名稱（42 個）

`IF_Price{Fresh,Cached,Think,Out,Ref}_{Astra,Sol,Luna}`（15）、`IF_FrontRef_*`（3）、`IF_ProgGWyr_*`（3）、`IF_RDMult`、`IF_TrainGenDefault`、`IF_TrainCost_*`（3）、`IF_RevGW_*`（3；`IFC_Use` 標「上限，不得作預測」，本模型**沒有**拿來當收入，合規）、`IF_FullCost_Astra／_Luna`、`IF_FullCostDefault_Astra／_Luna`、`IF_AllocQ1`、`IF_AllocQ1_R2`、`IF_AllocQ2`、`SRC_DEM_004`–`013`（含 013 低／高，共 12）。

- `SRC_DEM_*`：契約允許「SRC_ Active 紀錄」；v5.29 中 12 筆**全部 Active、數值不變**，合規。但共用快照工具只收 `IF_`／`L1_`，改用官方工具時這 12 筆要另行處理（本模型公式不用它們，可直接移除，或由本模型自己的 SRC_OAI 登錄同源數字——SRC_OAI_047 已是 8.4）。
- `IF_AllocQ1`、`Q1_R2`、`Q2`：**沒有任何公式引用**，只在 TK_Link 顯示。

### 1-4 審視工具候選清單的人工確認

| 工具報告 | 人工確認 |
|---|---|
| 違規引用 4 筆：`Inputs!`（build.py、p2.py、p3.py、p4.py） | **全部誤判**：都是本模型自己的 Inputs 頁（例 Compute G86 `=Inputs!$F$222`＝合約價低端） |
| 「不存在於現行 Tokenomics」4 名：`IF_Alloc`、`IF_FullCost`、`IF_Price`、`IF_TokGW` | **全部誤判**：是 tk_link.py 樣式字串的前綴片段，實際讀的都是完整名稱 |
| 使用名稱 49 個（含 `IF_RevGW_*_Prod`／`_Life`、`IF_FullCost_*_Prod`） | **不準**：工具把樣式展開成 Tokenomics 所有同前綴名稱；實際讀 63 個（見 1-1），**沒有**讀任何 `_Prod`／`_Life` 名稱 |
| 四層瀑布／用途欄：否／否 | 正確 |

**人工確認後的違規引用：0 筆**（工具 8 個候選全為誤判）。另有契約不符 2 項：①只取 `IF_TokGW` 100% 層、未取 `IFW_` 四層（契約 2-2）；②世代／成本情境表頭直接讀 Interface 第 4、5 列（契約 2-1 的 `IF_Hdr*`，形式上）。

---

## 二、口徑問題（共用工作單第 2 節）

### 問 1：收入怎麼計價？利用率與簽約率各用哪一個？有沒有重複扣減？

- **按 token 計價（模型商）**：營收＝訂閱人數 × 方案價格、API 計費 token × 單價、廣告、其他（Demand、Revenue 頁）。**營收公式完全不經過利用率或簽約率**——它是由需求端（人數、任務數、每任務 token、API 營收倒推）直接算出，不是「產能 × 利用率 × 單價」。
- 利用率只出現在**算力換算**：推論 GW＝token ÷（`IF_TokGW` × `IF_Util`），再 ÷ η。簽約率：不用（正確，契約第 3 節：按 token 計價只用技術利用率）。
- **重複扣減：基準情境沒有**。理由：η 以 2025 年「token 換算 GW ÷ 支出換算 GW」校準，`IF_Util` 同時出現在分子分母而相消；基準 η 路徑「沿用」（INP_258＝0）時，之後每年的有效推論 GW 也與 `IF_Util` 無關。容量上限另有「閒置保留 10%」「訓練最低占比 30%」兩個供給端扣減，屬於可用容量而非負載率，與利用率性質不同；且 2026–2030 都未觸頂（2030 餘裕 46%），對營收沒有作用。
- **需要注意的情境**：η 回升情境（INP_258＝1，η 線性回升到 2030＝1）下，`IF_Util` 60% 不再被抵銷，而且同時有閒置 10% 扣減；Tokenomics v5.29 的 `L1_GapUtil`（1.457）顯示實際利用率約 35%。此情境等於假設落差全部消失、且以 60% 負載＋10% 閒置為準，**偏樂觀**，見第五節。

### 問 2：神雲／雲平台的每 MW 收入

**不適用**：OpenAI 是模型商，沒有「每 MW 收入」與成本加成／市場價平均。唯一相關的是「每 GW 年合約價」（INP_229，12 $B/GW/年，區間 8–20，Analogy），只用來把 OpenAI 的支出換算成 GW（η 校準、研發 GW、未揭露 GW 的合約供給），不是收入輸入。容量情境（合約、自建）與合約價情境各自獨立（敏感度 A 組只動合約價）。

### 問 3：`IF_Alloc*` 六格取用處與 v5.29 影響（詳細）

**取用處**（全部是對照或檢查，沒有進入命題輸出）：

| 名稱 | 本模型位置 | 進入的計算 | v5.26 值 | v5.29 值 | 變動 |
|---|---|---|---|---|---|
| `IF_AllocQ1` | 只在 TK_Link | 無 | 0.636 | 0.573 | −9.9% |
| `IF_AllocQ1_R2` | 只在 TK_Link | 無 | 0.667 | 0.629 | −5.7% |
| `IF_AllocQ2` | 只在 TK_Link | 無 | 0.664 | 0.603 | −9.2% |
| `IF_AllocServeGW` | Checks C54、C55 | 只並列 | 0.0519 GW | 0.0675 GW | +30.1% |
| `IF_AllocDemand` | Demand D43（→D45、D46 `DEM_TkRatio2025`）；Checks C37、C55 | 只列差距 | 4,190 T tok | 6,015 T tok | +43.6% |
| `IF_AllocImpliedNk` | Checks C59 | 只並列 | 1.524 | 1.983 | +30.1% |
| `IF_AllocRDGW`（不在六格內） | Checks C58 | 只並列 | 0.0907 | 0.0907 | 0 |

**重取後受影響的本模型格**（全部是對照或資訊列；Checks 狀態都是 info，不會轉 ERR）：

| 位置 | 內容 | 現值（v5.26） | 重取後（v5.29） |
|---|---|---|---|
| Demand D43 | Tokenomics 需求 D（2025） | 4,190 T tok | 6,015 T tok |
| Demand D45 | 差距：本模型 − Tokenomics | −652 T | −2,477 T |
| Demand D46／Checks C37 | 比值 `DEM_TkRatio2025` | 0.844 | **0.588** |
| Checks C54 | TK 服務 GW | 0.0519 | 0.0675 |
| Checks C55 | TK 隱含每 GW 年產能 | 8.08×10¹⁰ M tok/GW/年 | 8.91×10¹⁰（+10.3%） |
| Checks C59 | TK 隱含 N × k | 1.524 | 1.983 |
| TK_Link B3–B6、Checks C129／C130、README、HTML 第 ⑧ 節 | 版本字樣 | v5.26／3dd1216 | v5.29／a5061d9 |

**不受影響**：營收、Microsoft 分成、token 量、η（0.2795）、推論 GW、研發 GW、供給 GW、容量上限、算力成本、雲端毛利、Q3、每 VR 等值 GW 差額、自由現金流、外部資金需求、反向模式、HTML 374 個數字中除版本字樣外全部。

**按 token 計價時利用率用哪一個、有無重複扣減**：用 `IF_Util`（Tokenomics 基準 60%），本模型沒有自己的利用率輸入；重複扣減的判斷見問 1（基準無、η 回升情境需注意）。

**是否把 `IF_TokGW` 的 100% 層當生產產能**：**形式上是**——Compute G03–G05 取 100% 層，× `IF_Util` 後當作「每 GW 年產能」，沒有用 `CTL_ProdDerate`（Prod 層）或 L × m（Life 層），也沒有 ISL 混合修正。**實質上**由 η（2025 年 0.2795，即落差 3.58 倍）一次吸收，所以基準輸出沒有高估產能；但 η 是一個總係數，沒有對應到契約 2-5 要求的四項口徑差，η 回升情境則會回到「Util 層＝生產產能」。

### 問 4：公司特有變數是否只在本模型覆寫、沒有回寫 Tokenomics？

**是**。合約價（INP_229）、η 路徑（INP_258–260）、世代組合（INP_205、211、217、223、256）、自研產能係數（INP_204）、閒置與訓練最低占比（INP_232、233）、自建投產落後（INP_262）、人數與成本、融資條件全部在本模型 Inputs／SRC_OAI。builder 只讀 Tokenomics（openpyxl 唯讀），沒有任何寫回路徑。本模型不使用 WACC（命題以現金與經濟口徑比較，不折現）。唯一的缺口：契約允許下游覆寫利用率，本模型直接用 `TK_IF_Util`，沒有自有覆寫格。

---

## 三、缺口與版本（共用工作單第 3 節）

### 3-1 Tokenomics 缺的產業級數據（只列，不建表；建議走 G2 進 DB_Evidence）

| 名稱 | 本模型用途 | 目前以什麼代替 |
|---|---|---|
| 每 GW 年算力租用（合約）價格，按長約計價 | 支出 → GW 換算、η 校準、未揭露 GW 合約的供給 | INP_229＝12（8–20），Analogy；Tokenomics 有持有成本（`L1_HoldEconGW_*`）與 IREN 長約 9.70 $M/MW 對照，但沒有「租用價」時間序列 |
| 自研／其他晶片（Trainium、Cerebras、OpenAI 自研）相對 VR200 的產能係數 | 世代組合第 6 欄的每 GW 產能、VR 等值 | INP_204＝0.6（0.4–1.0），Assumed；Tokenomics 只有 Nvidia 五個世代 |
| 推論支出的計價口徑（Azure 帳單 vs 持有成本 vs json 一致集）對應的單價 | η 的分母 | 本模型用 SRC_OAI_047＝8.4（json 口徑，與 `SRC_DEM_004` 同源）÷ 合約價 12；Tokenomics v5.29 另有 `SRC_DEM_018`＝12.6（Azure 帳單），本模型未取 |
| 實際請求的 ISL／OSL 混合分布（每任務 token 的組成） | Demand 每任務 token（INP_112–120、INP_031） | 本模型 Assumed（Free 2K、Plus 4K、Pro 10K、API 10K tok/任務）；Tokenomics v5.29 每則提示 token 數改為 4,000（SRC_DEM_014／015，OpenRouter 6,400）——兩邊口徑接近但未對齊 |

### 3-2 改用官方快照工具的名稱清單草稿（`data/tokenomics_names.txt`；只附在報告，未建檔、未執行正式快照）

```
# openai — Tokenomics 取數名稱清單草稿（2026-10-08；只列 IF_／IFW_／IFC_／L1_）
# A. 已進入公式（沿用）
IF_TokGW_Astra
IF_TokGW_Sol
IF_TokGW_Luna
IF_Util
IF_HoldEcon
IF_CapexTotal
IF_FullCost_Sol
IF_FullCostDefault_Sol
L1_Ans3
L1_Ans3_Lo
L1_Ans3_Hi
# B. Alloc（目前只作對照）
IF_AllocQ1
IF_AllocQ1_R2
IF_AllocQ2
IF_AllocServeGW
IF_AllocRDGW
IF_AllocDemand
IF_AllocImpliedNk
# C. 目前只存放、沒有公式引用（保留作對照；可精簡，由 chat 端決定）
IF_PriceFresh_Astra
IF_PriceFresh_Sol
IF_PriceFresh_Luna
IF_PriceCached_Astra
IF_PriceCached_Sol
IF_PriceCached_Luna
IF_PriceThink_Astra
IF_PriceThink_Sol
IF_PriceThink_Luna
IF_PriceOut_Astra
IF_PriceOut_Sol
IF_PriceOut_Luna
IF_PriceRef_Astra
IF_PriceRef_Sol
IF_PriceRef_Luna
IF_FrontRef_Astra
IF_FrontRef_Sol
IF_FrontRef_Luna
IF_ProgGWyr_Astra
IF_ProgGWyr_Sol
IF_ProgGWyr_Luna
IF_RDMult
IF_TrainGenDefault
IF_TrainCost_Astra
IF_TrainCost_Sol
IF_TrainCost_Luna
IF_FullCost_Astra
IF_FullCost_Luna
IF_FullCostDefault_Astra
IF_FullCostDefault_Luna
IF_RevGW_Astra
IF_RevGW_Sol
IF_RevGW_Luna
# D. v5.29 新增、契約要求（建議新增）
IFW_TokGW_Astra_100
IFW_TokGW_Astra_Util
IFW_TokGW_Astra_Prod
IFW_TokGW_Astra_Life
IFW_TokGW_Sol_100
IFW_TokGW_Sol_Util
IFW_TokGW_Sol_Prod
IFW_TokGW_Sol_Life
IFW_TokGW_Luna_100
IFW_TokGW_Luna_Util
IFW_TokGW_Luna_Prod
IFW_TokGW_Luna_Life
IFW_RevGW_Astra_100
IFW_RevGW_Astra_Util
IFW_RevGW_Astra_Prod
IFW_RevGW_Astra_Life
IFW_RevGW_Sol_100
IFW_RevGW_Sol_Util
IFW_RevGW_Sol_Prod
IFW_RevGW_Sol_Life
IFW_RevGW_Luna_100
IFW_RevGW_Luna_Util
IFW_RevGW_Luna_Prod
IFW_RevGW_Luna_Life
L1_ScaleRD
L1_ScaleRD_Lo
L1_ScaleRD_Hi
L1_ScaleServe
L1_ScaleServe_Lo
L1_ScaleServe_Hi
L1_GapPrompt
L1_GapSpendBasis
L1_GapISL
L1_GapUtil
L1_GapProduct
L1_ExtServeGW
# E. 輸出契約欄（Interface R:U 整欄範圍，依列取用；共用工具目前不支援）
IFC_Use
IFC_Layer
IFC_Conf
IFC_Updated
```

說明：`SRC_DEM_004`–`013` 不列入（工具不收 `SRC_`；本模型公式不用）。`IF_Hdr*` 不列入（契約禁止；共用工具會以 `IF_HdrGen`／`IF_HdrCost` 拆世代 × 成本情境，本模型改用工具後不需自己讀表頭）。

### 3-3 共用工具的限制（影響第 4 節做法）

- `tools/tokenomics/import_tokenomics.py` 只接受 `IF_`、`L1_`；`IFW_`（契約要求）、`IFC_`、`SRC_` 都會報錯中止（實測：`名稱不合規則：IFW_TokGW_Astra_100`）。`IFC_*` 是 305 列的整欄範圍，也不符合工具「單格或世代 × 成本情境列」的兩種形狀。**需要先擴充工具**，本模型才能完全改用官方快照。

---

## 四、v5.29 重取影響（共用工作單第 4 節）

- 做法：以草稿中 `IF_`／`L1_` 的 63 個名稱執行 `import_tokenomics.py --tokenomics /home/claude/tk-master --names <草稿子集> --out /tmp/openai_v529_snapshot.json`（v5.29、commit `a5061d9`、missing 0），`--check` 通過；`IFW_`、`IFC_`、`SRC_DEM` 以 openpyxl 直接讀 v5.29 具名範圍補足。再與本模型 TK_Link 的 63 個現用值逐格比較（相對誤差 1e-9）；另以本模型自己的 `tools/check_tk_snapshot.py` 交叉驗證，結果一致（OK 57、DIFF 6）。

### 4-1 現用 63 名的變動

| 結果 | 名稱數 | 內容 |
|---|---|---|
| 不變 | 57 | 所有 `IF_TokGW_*`、`IF_Util`、`IF_HoldEcon`、`IF_CapexTotal`、`IF_FullCost*`、`IF_Price*`、`IF_FrontRef_*`、`IF_ProgGWyr_*`、`IF_RDMult`、`IF_TrainGenDefault`、`IF_TrainCost_*`、`IF_RevGW_*`、`L1_Ans3`（＋低高）、`IF_AllocRDGW`、`SRC_DEM_004`–`013`（皆 Active） |
| 變動 | 6 | `IF_Alloc*` 六格（幅度見第二節問 3：Q1 −9.9%、Q1_R2 −5.7%、Q2 −9.2%、ServeGW +30.1%、Demand +43.6%、ImpliedNk +30.1%） |
| 消失 | 0 | — |

### 4-2 `IFW_` 四層相對本模型目前所用層的差異（VR200、基準成本情境）

本模型目前所用＝100% 層 × `IF_Util`（＝Util 層）。

| 名稱 | 100% | Util（本模型現用） | Prod | Life | Prod ÷ 現用 |
|---|---|---|---|---|---|
| `IFW_TokGW_Sol_*`（M tok/GW/年） | 2.373×10¹¹ | 1.424×10¹¹ | 1.916×10¹¹ | 1.916×10¹¹ | **1.35** |
| `IFW_TokGW_Astra_*` | 5.117×10¹⁰ | 3.070×10¹⁰ | 3.919×10¹⁰ | 3.919×10¹⁰ | **1.28** |
| `IFW_TokGW_Luna_*` | 7.728×10¹¹ | 4.637×10¹¹ | 6.082×10¹¹ | 6.082×10¹¹ | **1.31** |
| `IFW_RevGW_Sol_*`（$B/GW/年；本模型未用於收入） | 381.3 | 228.8 | 184.8 | 228.8 | 0.81 |

**觀察（交 chat 端／Tokenomics 端）**：Tokenomics 的 `IFW_TokGW_*_Prod` 讀 `Interface_Prod` 同列（＝100% × 生產折減 0.85 等，**不含** `IF_Util`），所以 Prod 層**大於** Util 層；而 `IFW_RevGW_*_Prod` 是「100% × IF_Util × 生產折減」，比 Util 層小；`IFW_RevGW_*_Life` 又等於 Util 層。四層並不是逐層相乘的同一條瀑布。下游若照字面「改取 Prod 層當產能」，產能會比現用**高 28–35%**（推論 GW 相應少 22–26%），方向與契約「不得高估產能」的本意相反。本模型改接前，需要先確認 TokGW 的 Prod 層是否應再乘 `IF_Util`。

### 4-3 v5.29 新增的規模校準量與本模型 η 的粗略對照（只供判斷，不是計算結果）

| 項目 | Tokenomics v5.29 | 本模型對應 |
|---|---|---|
| 服務端總落差 | `L1_ScaleServe` 16.50（12.9–21.5）；四項乘積 `L1_GapProduct` 13.44 | 1 ÷ η＝**3.58**（2025） |
| ① 每則提示 token 數 | `L1_GapPrompt` 2.0 | 不適用：本模型 token 量由人數 × 任務 × 每任務 token 自行建立（2025 合計 3,538 T，為 Tokenomics D 的 0.59 倍） |
| ② 支出計價口徑 | `L1_GapSpendBasis` 2.0（Azure 計價 ÷ 持有成本） | 本模型以「合約價 12」換算，已含供應商利潤（2025 世代加權持有成本 9.73，比值約 1.23）；但支出用 8.4（json 口徑），若改用 `SRC_DEM_018` 12.6，支出換算 GW 由 0.70 → 1.05，η 由 0.280 → 0.186 |
| ③ ISL 混合 | `L1_GapISL` 2.31（低 0.52） | 未處理（本模型用參考請求的 `IF_TokGW`） |
| ④ 利用率 | `L1_GapUtil` 1.46（0.6 × 0.85 ÷ 0.35） | 本模型用 0.6、不乘生產折減；若實際 0.35，對應 1.71 |

粗算：③ × ④（僅利用率）＝ 2.31 × 1.71 ≈ 3.95，與本模型 3.58 相差約 10%；若支出改用 12.6，本模型落差變 5.37，③ × ④ 只能解釋約 74%。這表示**本模型的 η 大致可以拆成 Tokenomics 的口徑差，但支出口徑的選擇（8.4 或 12.6）會改變結論**——依契約 2-5，校準後的差額要回報 Tokenomics。

- 研發端：`L1_ScaleRD` 10.37（OpenAI 訓練支出 12 ÷ `L1_Ans3` 1.157）。本模型 Cost K71 已列「研發算力 ÷ TK L1_Ans3」＝10.37（2025），**數字與 `L1_ScaleRD` 相同**，只是沒有引用該名稱。

---

## 五、建議改動（交 chat 端另開工作單；本輪未執行）

依重要性排序：

1. **把 η 拆成契約 2-5 的四項口徑差，並取 `L1_ScaleServe`／`L1_Gap*`／`L1_ScaleRD`**：在 Compute 第六節旁新增「η 分解」（提示 token、支出口徑、ISL、利用率 → 本模型各自的值），η＝四項之積 × 殘差；殘差回報 Tokenomics。同時決定 η 的分母用 8.4（json）還是 12.6（Azure 帳單，`SRC_DEM_018`），與 Tokenomics r2 的分工（比例用 004、服務 GW 用 018）對齊。研發端 Cost K71 改引用 `L1_ScaleRD`。
2. **取 `IFW_TokGW_*` 四層並明定產能用哪一層；新增自有利用率覆寫格**：Compute G03–G05 改讀 `IFW_TokGW_*_100`，Util／Prod／Life 並列；先由 Tokenomics 確認 TokGW 的 Prod 層是否應含 `IF_Util`（見 4-2）。另在 Inputs 新增「本模型利用率」（預設＝`IF_Util`，區間含 0.35），並把 η 回升情境的終點與它綁在一起，避免「η→1 且用 60% 利用率＋10% 閒置」的樂觀組合。
3. **改用共用快照工具並釘版本**：擴充 `tools/tokenomics/import_tokenomics.py` 收 `IFW_`（同 IF_ 的世代 × 成本列形狀）與 `IFC_`（依列取 R–U），之後本模型改為 `company.json`／`data/tokenomics_snapshot_v5.29.json` → builder 只讀快照（世代鍵由快照提供，不再直接讀 Interface 第 4、5 列）；CI 的 `tk-snapshot` 改為「與快照記錄的同一個 xlsx 比對」（`--check`），Tokenomics 新版只報「有新版可取」。同時把只存放、沒有引用的 42 個名稱精簡或標明用途，`SRC_DEM_*` 移出 TK_Link。

（v5.29 重取本身——`IF_Alloc*` 六格——可隨第 3 項一起做，只影響對照列與版本字樣，不影響命題結論。）

---

## 待 chat 端判斷

1. **η 的支出分母**：維持 SRC_OAI_047＝8.4（json 一致集），或改用 Tokenomics 新增的 `SRC_DEM_018`＝12.6（Azure 帳單）？前者 η＝0.280，後者 η＝0.186（有效推論 GW 2025 起放大約 1.5 倍，研發 GW 殘差相應縮小、容量餘裕下降）。
2. **TokGW 四層的語意**：`IFW_TokGW_*_Prod` 不含 `IF_Util`、比 Util 層大 28–35%，與 `IFW_RevGW_*_Prod` 的定義不一致——是 Tokenomics 端要修正，還是下游要自己乘 `IF_Util`？（本模型改接 IFW 之前必須先定）
3. **η 回升情境是否保留**：若保留，終點是否改為「四項口徑差中只有可改善的項目（例如利用率）回升」，而不是整個 η→1。
4. **利用率覆寫**：本模型是否新增自有利用率（例如以 `AL_UtilActual` 0.35 為基準或低端），以及與「閒置保留 10%」是否算重複扣減。
5. **每任務 token 與 Tokenomics 每則提示 4,000 的關係**：本模型 Free 2K、Plus 4K tok/任務是否要對照 `SRC_DEM_014`／`015` 調整（屬公司模型的需求假設，但證據相同）。
6. **共用工具擴充**（`IFW_`、`IFC_`）由哪一方做、何時做；在工具擴充前，本模型是否先只做 v5.29 修補版重取（僅版本與 `IF_Alloc*` 對照列改變）。
7. 盤點報告把本模型記為「取數方式：無、未釘版本」，建議更正為「builder 直讀、Excel 記錄 v5.26」。
