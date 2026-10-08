# WhiteFiber｜Tokenomics 連接審視（v5.29）報告（2026-10-08）

- 分支：`claude/whitefiber-tk-link-review`（自 main `75327ef` 開出）
- 最新提交 SHA：見 PR 的最新提交（本檔第一次提交後，PR 留言會註明 SHA）
- 報告路徑：`whitefiber/docs/reports/20261008_whitefiber_tokenomics_link_review.md`
- 依據：共用工作單 `docs/workorders/20261008_tokenomics_link_review.md`、`whitefiber/docs/workorders/20261008_whitefiber_tokenomics_link_review.md`；Tokenomics 下游資料契約 v0.1（`docs/plan/Tokenomics_downstream_contract.md`）。
- 對照的 Tokenomics：master `a5061d9`，`model/CURRENT`＝`20261008_Tokenomics_v5.29.xlsx`（`IFW_` 44 個、`IFC_` 4 個名稱存在 → **前提成立**，第 1–4 節全做）。
- **本單只審視、不改模型**：company.json、builder、Excel、data 都沒有改；v5.29 快照只產生在暫存路徑 `/tmp/whitefiber_v529_snapshot.json`，未提交。

## 一、結論（先看這裡）

1. **數字面：v5.29 對 WhiteFiber 沒有任何影響。** WhiteFiber 用到的 Tokenomics 數值（GB300／VR200 的每 GW GPU 數、PUE、IT／廠房／合計資本支出、會計與經濟持有成本、每 GPU 小時成本、電費，以及 H100／GB200 的每 GW GPU 數）在 v5.24 → v5.29 逐格相同（`Interface!C7:Q15` 135 格差異 0，`IF_Util` 60% 不變）。v5.29 改變的 `IF_Alloc*` 六格是給模型商（OpenAI 等）用的，WhiteFiber 沒有引用。
2. **方法面：每 MW 收入仍是「成本加成與市場價 50/50 平均」（保守 11.62／基準 17.40／積極 24.20 US$m／MW-IT·年），契約第 4 條已停用此做法。** 而且容量情境和單價情境由同一個情境選擇器同向綁定（保守＝少容量＋低價、積極＝多容量＋高價），契約要求分離。這是最主要的待改項目，應比照 Nebius v0.2a／CoreWeave W4 改為「Tokenomics 持有成本錨 × 定價倍數 k」。
3. **取數方式：仍是手工 JSON（Oracle／Nebius 的副本），不是官方快照。** 有 3 處數據欄位直接引用禁止的工作表（`Inputs!D23:F23` IT 壽命、`Inputs!D28:F28` WACC、`DC_Cost!I48:N48` 營運費用），另 2 處說明文字提到 `DC_Cost!`；這 5 處都有合規的替代名稱或可刪除（見第三節表 1）。company.json 內的 11.62／17.40／24.20、37.45／37.59、每顆 GPU kW 等是手抄的數字（寫死），沒有快照可重現。
4. **可能的重複扣減**：每 MW 收入的市場價半邊已乘 85% 可計費利用率，模型又在其上乘「在役／已連網收斂比例」（穩態 90%），兩者都是商業面的打折，穩態合計約 76.5%。是否屬重複扣減需 chat 端判斷（第二節 Q1）。
5. **版本紀錄不一致**：`permw_tokenomics_20261008.json` 的 `tokenomics.file` 欄仍寫 v5.24（`098873a`），但同檔 whitefiber 註記與交接檔寫 v5.26（`70d859e`）；交接檔也沒有寫明「按 MW-year 計價、採簽約率不採 IF_Util」（契約第 3、5 條要求）。

## 二、第 1 節　取數點清單

工具 `python3 tools/tokenomics/link_audit.py --tokenomics /home/claude/tk-master` 的 whitefiber 結果：手工 JSON（`permw_tokenomics_20261008.json`）、釘 v5.24（落後）、使用 10 個名稱、不存在於現行 0、違規引用 2 類（`DC_Cost!`、`Inputs!`，皆在 permw 檔）、未取 `IFW_`／`IFC_`。人工逐筆確認後：**工具沒有誤判**（WhiteFiber 自己的 Excel 工作表名稱是中文，例如「輸入與假設」，沒有與 `Inputs`、`DC_Cost` 同名的頁）；工具未抓到的另有 `DC_Cost!J22`、`DC_Cost!52` 兩處說明文字，以及 company.json 內的手抄數字。

「版本」欄：permw 檔記錄 v5.24（`098873a`）；建置代理 2026-10-08 以 v5.26（`70d859e`）核對同一範圍數值相同；本次以 v5.29（`a5061d9`）核對仍相同。下表簡寫「v5.24 記錄／v5.26 核對」。

「四層」欄：契約的四層瀑布（100%／IF_Util／Prod／Life）只適用於每 GW 產能與理論營收；成本、資本支出類名稱不分層，記「不適用」。「IFC_Use」欄取 v5.29 `Interface` R–U 欄。

| # | 本模型的格或參數 | Tokenomics 名稱 | 合規／違規 | 四層 | 用途是否符合 IFC_Use | 版本 |
|---|---|---|---|---|---|---|
| 1 | `scenarios.revMW` 三情境 11.62／17.40／24.20（每 MW 年收入，模型主驅動） | `IF_HoldEcon`（GB300 基準 12.72）經成本加成；`IF_GPUsPerGW`（487 顆/MW）經市場價 | **違規**：名稱合規，但 (a) 數字手抄寫死、(b) 成本加成當收入輸入且 50/50 平均（契約第 4 條停用） | 不適用（成本） | IFC_Use「情境、相對比較、反轉門檻」：未用到「上限」列 → 符合；但違反契約第 4 條 | v5.24 記錄／v5.26 核對 |
| 2 | `scenarios.capexTemplate.costMW` 37.45／37.59（每 MW GPU 資本支出，雲端 CapEx 與汰換） | `IF_CapexIT`（GB300、VR200 基準） | 名稱合規；**寫死數字**（手抄、四捨五入至 2 位） | 不適用 | 符合（基準情境值） | v5.24／v5.26 |
| 3 | `defaults.refreshSteady` 汰換 CapEx（`IF_CapexIT` ÷ `gpuLife` 5 年） | `IF_CapexIT` | 合規（經 #2） | 不適用 | 符合 | 同上 |
| 4 | `methodology.checks.capexPerMwBand` [28, 51]（CapEx 強度檢查區間） | `IF_CapexIT` GB300 低／高（28.55／50.95） | 名稱合規；**寫死數字** | 不適用 | 符合（區間檢查） | 同上 |
| 5 | `scenarios.mwPath`（合約 GPU 數 → MW-IT：14.574、16.548 等）、`kwPerGpuIT`（H100 1.393、GB200 1.993、GB300 2.053、VR200 3.428 kW） | `IF_GPUsPerGW` 倒數（H100、GB200、GB300、VR200 基準） | 名稱合規；**寫死數字**（換算結果手寫進 company.json） | 不適用 | 符合 | v5.26 |
| 6 | 巴黎站 1.84 MW-IT（年額 >32 ÷ 17.40） | 間接：#1 的基準 17.40 | 隨 #1 | 不適用 | 隨 #1 | 同上 |
| 7 | `billableRatio` 期初可計費 MW（46.02 ÷ 各情境每 MW 收入 → 3.96／2.65／1.90） | 間接：#1 | 隨 #1 | 不適用 | 隨 #1 | 同上 |
| 8 | permw 檔 `tokenomics.rows`：PUE、IT／合計資本支出、會計持有成本、電費（對照、MW 口徑換算） | `IF_FacilityGW`、`IF_CapexIT`、`IF_CapexTotal`、`IF_HoldAcct`、`IF_PowerCost` | 合規（以名稱註記儲存格）；手抄 | 不適用 | 符合 | v5.24／v5.26 |
| 9 | permw 檔 `gpuHrEcon100`、VR200 價格換算（`Interface!M14 ÷ J14`＝1.674） | `IF_GPUhrEcon` | 合規；但以儲存格位址而非名稱表達換算 | 100% 時數（IFC_Layer「—」） | 符合（相對比較） | 同上 |
| 10 | permw 檔 `dcOpexPerMWit`（隱含 EBITDA 率檢核） | **`DC_Cost!I48:N48`** | **違規**（禁止工作表）；v5.26 起有 `IF_OpexGW`，數值相同（GB300 基準 2.63 對 2.625） | 不適用 | — | 同上 |
| 11 | permw 檔 `itLifeYears`（6／6／4，僅供參考） | **`Inputs!D23:F23`** | **違規**；v5.26 起有 `IF_DeprLifeIT`，數值相同 | 不適用 | — | 同上 |
| 12 | permw 檔 `wacc`（8%／10%／13%，僅供參考；已內含於 `IF_HoldEcon`） | **`Inputs!D28:F28`** | **違規**；Tokenomics 沒有對應的 `IF_` 名稱（見第三節缺口 6） | 不適用 | — | 同上 |
| 13 | permw 檔說明文字「儲存與管理 $500/GPU（`DC_Cost!J22`）」「資本回收係數 `DC_Cost!52`」 | — | **違規（說明文字）**；Nebius 推導沿用下來，不影響數值，可刪或改寫 | — | — | — |
| 14 | permw 檔 `IF_Util` 60%（明文**不沿用**，只作說明） | `IF_Util` | 合規（只列不用） | IF_Util 層 | 符合（未進收入） | 同上 |
| 15 | `data/whitefiber_facts_20261008.json` `build.tk`（廠房資本支出 10.21／12.67／16.19，託管建置成本區間的參考之一） | `IF_CapexFacility` | 名稱合規；手抄 | 不適用 | 符合（IFC_Use「基準輸入」） | v5.26 |

**違規小計（人工確認後）**：禁止工作表引用 5 處（數據欄位 3 處：#10–#12；說明文字 2 處：#13）；寫死數字 4 組（#1、#2、#4、#5；#6、#7 隨 #1 連動）；方法違反契約第 4 條 1 項（#1 的 50/50 平均與成本加成當收入）。未引用 `IFW_`、`IFC_`、`L1_`、`SRC_`；未引用任何 v5.29 不存在的名稱。

## 三、第 2 節　口徑問題（四答）

**Q1　收入按什麼計價？利用率與簽約率用哪個？有無重複扣減？**
- 按 **MW-year**（雲端：在役 MW × 每 MW 年收入；託管：已簽約每 MW 租金排程）。對應契約第 3 條第二列：收入應取決於**商業簽約率**，不應再乘技術利用率 `IF_Util`。
- `IF_Util`：**沒有進收入**（permw 檔明文「不沿用」；模型 `defaults.util` 為 100%）。這點合規。
- 但有**兩層商業面打折疊加**的疑慮：(a) 每 MW 收入的市場價半邊（路徑 B）已乘 85% 可計費利用率（保守 take-or-pay 不乘、積極 90%）；(b) 模型又把可計費 MW 設為「期初校準值＋(已連網 − 期初)× 收斂比例」，收斂比例 70%→85%→90%→90%→90%，穩態仍只有 90%。若 (b) 的本意是「新容量上線到起算的時間差」，穩態應趨近 100%；若本意是「簽約率」，則與 (a) 的 85% 重複。穩態合計約 0.85 × 0.90＝76.5%（只作用在 50/50 的路徑 B 半邊）。成本加成半邊（路徑 A）沒有打折，再乘 (b) 一次。→ **列為待 chat 端判斷**。
- 交接檔沒有寫明採哪一種計價（契約第 3 條要求寫明）。

**Q2　每 MW 收入是否仍為 50/50？容量與單價是否同一選擇器同向綁定？**
- **是，仍為 50/50。** 保守＝(路徑 A 保守 13.54＋路徑 B IREN–Microsoft 長約 9.70)÷2＝11.62；基準＝(15.91＋Verda GB300 spot 5.21 × 85%＝18.89)÷2＝17.40；積極＝(18.38＋24 個月預留 7.82 × 90%＝30.03)÷2＝24.20。三情境全期持平，沒有世代組合、也沒有價格隨時間變動。
- **是，同向綁定。** 一個情境選擇器（`scenarios` 的 low／base／high；HTML `SCENARIOS`、Excel「輸入與假設」情境選擇）同時切換：合約 MW 上限（`mwPath.contracted`）、每 MW 年收入（`revMW`）、建設延誤、NC-1 專案貸款、託管站點（可轉債上限也按情境取值，但三情境目前同為 75.9）。保守＝少容量＋低價＋延誤 6 月；積極＝多容量＋高價＋不延誤，因此目標價區間（$10.0–$34.7）被兩條軸疊加放大。
- 一個抵銷效果：期初可計費 MW＝Q2 實際營收年化 46.02 ÷ 該情境每 MW 收入，所以單價愈低、期初可計費 MW 愈高（保守 3.96／積極 1.90）。首期營收因此被拉回實際值附近，但隨收斂比例往已連網 MW 靠攏後，單價差距就完全顯現。
- 契約第 4 條（X14 (j)）：成本加成只作 ROIC 與資金缺口檢查、50/50 停用；中性情境定義為 q × p ≈ 1。CoreWeave W4、Nebius v0.2a 已改為「`IF_HoldEcon` 錨 × 定價倍數 k」並把容量軸與價格軸分開。

**Q3　`IF_Alloc*` 六格取用處與受影響輸出**
- **不適用。** WhiteFiber 是神雲（GPU 雲＋託管），沒有引用 `IF_Alloc`、`IF_AllocQ1`、`IF_AllocQ2`、`IF_AllocServeGW`、`IF_AllocDemand`、`IF_AllocImpliedNk` 任何一格；v5.29 這六格的變動（每則提示 token 4,000、推論支出改 SRC_DEM_018）對 WhiteFiber 的輸出影響為 0。

**Q4　公司特有變數是否只在本模型覆寫、沒有回寫 Tokenomics？**
- **是。** WACC（CAPM，β 2.5、kd 9.5%）、合約價（各 GPU 合約 TCV）、可計費比例、交付進度（`mwPath`、延誤月數）、託管租金與年調、GPU 經濟壽命（5 年）都只在 `company.json`；WhiteFiber 沒有任何寫入 Tokenomics 的路徑。
- 兩點口徑提醒（不是違規）：(a) 路徑 A 的「k＝0＝只賺 WACC」指的是 Tokenomics 內含的 10% 經濟 WACC，與 WhiteFiber 自己的 CAPM WACC 不同；(b) `IF_HoldEcon` 以 IT 壽命 6 年計資本回收，WhiteFiber 的 GPU 折舊與汰換用 5 年。改用錨定法時須在報告寫明兩者並存的理由。

## 四、第 3 節　缺口與名稱清單草稿

### 表 1　禁止引用的替代
| 現行引用 | 合規替代 | 數值差異 |
|---|---|---|
| `DC_Cost!I48:N48`（DC 營運費用） | `IF_OpexGW`（v5.26 起） | 0（四捨五入範圍內） |
| `Inputs!D23:F23`（IT 壽命） | `IF_DeprLifeIT`（v5.26 起） | 0 |
| `Inputs!D28:F28`（WACC） | 無 `IF_` 名稱 → 刪除該欄或改寫為文字說明，並提缺口 6 | — |
| `DC_Cost!J22`、`DC_Cost!52`（說明文字） | 刪除或改寫為「已內含於 `IF_CapexIT`／`IF_HoldEcon`」 | — |

### 表 2　Tokenomics 缺的產業級數據（不自建表；交 chat 端走 G2 進 DB_Evidence）
| # | 名稱 | 用途 | 目前以什麼代替 |
|---|---|---|---|
| 1 | GPU 租價分層時間序列（GB300／B300／VR200／H100；隨需／預留／長約） | 契約第 4 條的 p；定價倍數 k 的證據 | permw 檔自蒐的 Verda 牌價（隨需 10.42、24 月預留 7.82、spot 5.21）、GetDeploying 彙整（Oracle、CoreWeave、AWS、Azure）。Tokenomics `SRC_Price` 只有 CoreWeave GB200 牌價（SRC_PRC_001）與 Cape Fear 損益兩平租金（SRC_PRC_002）兩筆算力租金，v5.29 新增的分層欄（X–AB）仍空白；DB_Evidence E246 只有摘要級指數 |
| 2 | IREN–Microsoft 5 年 GB300 合約（9.70 US$m/MW-IT·年） | 長約端 k 的證據（契約第 4 條已引用 0.76，但無 SRC 紀錄） | permw 檔自行登錄（IREN 新聞稿 2025-11-03，[Interested-party]） |
| 3 | HGX（氣冷 8 卡）B300／B200 每 GPU IT kW | WhiteFiber 多為 HGX 機型，GPU 數 → MW 換算 | 以 NVL72 欄（GB300／GB200）近似，[Analogy]，可能高估 MW 約 0–12%（使合約隱含單價略低估） |
| 4 | H200 欄 | H200 機隊的成本與 kW | 以 Hopper H100 欄近似 |
| 5 | 託管（colocation）每 MW-IT 年租金市場區間 | 託管分部未簽約擴建的租金 | WhiteFiber 自蒐可比 1.45／1.85／2.35（[Analogy]）；屬產業價格，是否進 Tokenomics 由 chat 端判斷 |
| 6 | Tokenomics 經濟 WACC 的 `IF_` 名稱 | 說明 `IF_HoldEcon` 的資本回收口徑 | 只能引 `Inputs!D28:F28`（禁止） |

工具面缺口（非數據）：`import_tokenomics.py` 目前只接受 `IF_`、`L1_`，`IFW_`／`IFC_` 與 `SRC_` Active 紀錄會被拒絕；四層瀑布與用途欄要進快照需先擴充工具（本次以 openpyxl 直接讀 v5.29 替代，見第五節）。

### 名稱清單草稿 `whitefiber/data/tokenomics_names.txt`（只附在報告，不建立、不執行）
```
# WhiteFiber 引用的 Tokenomics 名稱（草稿；工具：tools/tokenomics/import_tokenomics.py）
# 成本、資本支出、MW 換算（現行已用）
IF_GPUsPerGW            # 每 GW GPU 數（GPU 數 → MW-IT；kW/GPU）
IF_FacilityGW           # PUE（MW-IT ↔ 設施 MW）
IF_CapexIT              # 每 MW GPU（IT）資本支出：雲端 CapEx、汰換、CapEx 強度檢查
IF_CapexFacility        # 每 MW 廠房資本支出：託管建置成本參考
IF_CapexTotal           # 每 MW 合計資本支出（對照）
IF_HoldAcct             # 會計持有成本（對照）
IF_HoldEcon             # 經濟持有成本（錨定法的錨；成本加成只作 ROIC 檢查）
IF_GPUhrEcon            # 每 GPU 小時經濟成本（k＝市場價 ÷ 此值；VR200 價格換算）
IF_PowerCost            # 電費
IF_DeprLifeIT           # 取代 Inputs!D23:F23
IF_OpexGW               # 取代 DC_Cost!I48:N48（隱含 EBITDA 率；成本端評估）
IF_Util                 # 只作對照：按 MW-year 計價，不進收入（契約第 3 條）
# v5.29 下游每 MW 介面（契約第 3、4 條）
L1_HoldEconMW_Hopper
L1_HoldEconMW_GB200
L1_HoldEconMW_GB300
L1_HoldEconMW_VR200
L1_TokMW_Gen_ratio_GB200
L1_TokMW_Gen_ratio_GB300
L1_TokMW_Gen_ratio_VR200
L1_GPUhr_GB300
L1_GPUhr_VR200
L1_GPUhr_GB300_vsBE     # 對照：新雲損益兩平租金
# 上限檢查（四層全取；IFC_Use 標「上限，不得作預測」者只作檢查）— 需工具先支援 IFW_／IFC_
IFW_RevGWFleet_100
IFW_RevGWFleet_Util
IFW_RevGWFleet_Prod
IFW_RevGWFleet_Life
IFC_Use
IFC_Layer
```
（本次實際以工具跑的是前兩段 22 個名稱，另加 `IF_RevGWFleet` 共 23 個；`IFW_`／`IFC_` 以 openpyxl 直接讀。）

## 五、第 4 節　v5.29 重取影響

**執行方式**：
- `python3 tools/tokenomics/import_tokenomics.py --tokenomics /home/claude/tk-master --names <草稿前兩段＋IF_RevGWFleet> --out /tmp/whitefiber_v529_snapshot.json` → 成功（`20261008_Tokenomics_v5.29.xlsx`，commit `a5061d9`；23 個名稱，missing 0；快取值完整，未觸發 LibreOffice 重算）。
- `--check /tmp/whitefiber_v529_snapshot.json` → 通過（相對誤差 ≤ 1e-9）。
- `IFW_`／`IFC_`：工具不接受，以 openpyxl 直接讀 v5.29 `Interface` 第 248–251 列與 R–U 欄。
- 現用值比較：permw 檔 `tokenomics.rows`（6 列 × 10 欄）與 `kwPerGpuIT`、company.json `costMW` 對 v5.29 快照；另以 openpyxl 逐格比較 v5.24（`model/archive/20261007_Tokenomics_v5.24.xlsx`）與 v5.29 的 `Interface!C7:Q15`。

**結果**：
| 項目 | 現用值 | v5.29 | 變動 |
|---|---|---|---|
| `Interface!C7:Q15`（9 列 × 15 欄，v5.24 對 v5.29） | — | — | **0 格** |
| permw 檔 60 個數值（GB300／VR200 × 低／基準／高 × 10 項） | 2 位小數 | 同 | 0（最大相對差 1.2% 為四捨五入，例如電費 0.40 對 0.405） |
| `dcOpexPerMWit`（原 `DC_Cost!`）對 `IF_OpexGW` | 2.63（GB300 基準） | 2.625 | 0（四捨五入） |
| `itLifeYears`（原 `Inputs!`）對 `IF_DeprLifeIT` | 6／6／4 | 6／6／4 | 0 |
| kW/GPU（H100／GB200／GB300／VR200） | 1.393／1.993／2.053／3.428 | 同 | 0 |
| `costMW`（GB300／VR200） | 37.45／37.59 | 37.446／37.586 | 0（四捨五入） |
| `IF_Util` | 60%（不用） | 60% | 0 |
| `IF_Alloc*` 六格 | 未引用 | 已變（v5.29 第 (k)(l) 項） | 對本模型無影響 |
| 每 MW 收入 11.62／17.40／24.20 | — | 輸入未變 → 若沿用現法數字不變 | 0 |

**v5.29 新提供、本模型尚未取用的名稱**（改錨定法時會用到）：
- `L1_HoldEconMW_*`（每 MW 年經濟持有成本）：Hopper 10.10、GB200 9.17、GB300 12.72、VR200 12.76（基準；低／高 GB300 8.71／23.87）。**注意**：L1 頁單位欄寫「$M/MW/年」，但儲存格數值是 0.01272（＝每 GW $B ÷ 1,000，實為 $B/MW），差 1,000 倍；取用時須換算，並回報 Tokenomics（見第六節）。
- `L1_TokMW_Gen_ratio_*`（相鄰世代每 MW 產出比）：GB200÷Hopper 4.85、GB300÷GB200 1.46、VR200÷GB300 1.57（區間 0.92–2.05）。
- 四層瀑布 `IFW_RevGWFleet_*`（OpenAI 有效單價下 1 GW 參考機隊付費營收，理想上限）。本模型不拿它當收入（IFC_Use：上限，不得作預測），但可作 Nebius v0.2a 第 8 步那種「上限檢查」。以 GB300 基準欄換算，WhiteFiber 現行每 MW 收入占比：

| 層 | GB300 每 MW 上限（US$m/年） | 基準 17.40 占比 | 積極 24.20 占比 |
|---|---|---|---|
| `_100`（滿載） | 43.78 | 40% | 55% |
| `_Util`（× IF_Util 60%） | 26.27 | 66% | 92% |
| `_Prod`（× 生產折減 CTL_ProdDerate） | 19.23 | 90% | **126%** |
| `_Life`（× L × m） | 26.27 | 66% | 92% |

  意思是：在積極情境下，WhiteFiber 每 MW 向客戶收的租金，已超過「客戶把同一 MW 拿去以 OpenAI 價格賣 token、扣生產折減後」能賺到的營收；基準也達 66%–90%。Nebius v0.2a 的警示門檻是 50%。這支持把積極情境從「高價」改為「多容量、基準價」。（VR200 欄上限較高：基準占比 36%／46%（Util／Prod）。）
- 一個 Tokenomics 端的小不一致：`IF_RevGWFleet`（第 143 列）的 IFC_Use 標「上限，不得作預測」，但數值相同的 `IFW_RevGWFleet_Util`（第 249 列）及其他三層的 IFC_Use 只寫「情境、相對比較、反轉門檻」，沒有「上限」字樣。

## 六、建議改動（交 chat 端開工作單；本單不做）

1. **每 MW 收入改以 Tokenomics 為錨（比照 Nebius v0.2a）**：每 MW 年收入＝Σ 在役世代占比 × `L1_HoldEconMW_*`（或 `IF_HoldEcon`）× 定價倍數 k；停用 50/50；成本加成移到 ROIC／資金缺口檢查；容量軸（現行三情境）與價格軸（k 低／基準／高）分離，三個容量情境預設用基準 k，加 3 × 3 目標價矩陣與上限檢查（門檻放 `methodology.checks`）。參考刻度（GB300 錨 12.72）：k＝0.76（IREN 長約）→ 9.67；k＝1 → 12.72；k＝1.76（現貨）→ 22.4；WhiteFiber 2026 新簽 B300／VR200 長約隱含 14.8–20.3（平均 18.2，約 1.43 倍錨）只作對照與驗證，**不得用來反推 k**。WhiteFiber 機隊以 HGX B300 為主，世代占比與 kW 換算的近似要寫明。
2. **取數改官方快照**：建立 `whitefiber/data/tokenomics_names.txt`（第四節草稿）與 `tokenomics_snapshot_v5.29.json`，Excel 加「Tokenomics_取數」分頁與 `TK_` 名稱、verify.sh 加 `--check`；permw 檔的 `Inputs!`／`DC_Cost!` 五處改名或刪除，舊檔標「已被快照取代」；company.json 的 37.45／37.59、[28, 51]、kW/GPU 改由快照產生。本步數字應不變（第五節已證）。
3. **釐清可計費比例與 85% 的關係並補交接檔口徑**：決定「收斂比例」是上線時間差（穩態→100%）還是簽約率（則路徑 B 的 85% 應移除），避免兩層商業打折；交接檔寫明「按 MW-year 計價、採簽約率、`IF_Util` 不進收入」與 Tokenomics 版本（`model/CURRENT` 檔名＋合併雜湊），並修正 permw 檔 v5.24／v5.26 版本欄不一致。

其他（次要）：工具 `import_tokenomics.py` 擴充接受 `IFW_`／`IFC_`／`SRC_` Active（各公司共用）；第四節表 2 缺口 1–3 走 G2。

## 七、待 chat 端判斷

1. **可計費收斂比例（穩態 90%）與路徑 B 的 85% 可計費利用率是否重複扣減？** 建議：收斂比例定義為上線時間差、穩態 100%；簽約率另設一個明確參數（三情境共用、可作敏感度）。改錨定法後路徑 B 不再直接進收入，此問題大部分自然消失，但收斂比例的定義仍需定。
2. **WhiteFiber 的 k 證據來源**：依 Nebius v0.2a 規則，不得用公司自己的合約、營收、ACV 反推 k。但 WhiteFiber 的雲端合約（Baseten、Prime Intellect 等）是逐筆揭露的第三方成交，是否可作為「可比 neocloud 長約」的證據之一？建議：只作驗證（模型 k 與合約隱含 k 的差距列表說明），k 仍以獨立來源（IREN–Microsoft、現貨指數、SRC_Price 新分層紀錄）為準。
3. **積極情境的上限**：積極 24.20 已達 GB300 `IFW_RevGWFleet_Prod` 的 126%。改版時積極情境是否直接改為「多容量＋基準 k」（建議是），或保留高 k 但加上限檢查警示。
4. **GPU 壽命口徑**：`IF_HoldEcon` 以 6 年計資本回收，WhiteFiber GPU 折舊與汰換用 5 年。錨定法下是否接受兩者並存（建議並存並寫明：錨是產業打平租金，5 年是公司會計）。
5. **Tokenomics 端回報**（不在本 repo 改）：(a) `L1_HoldEconMW_*` 單位欄寫「$M/MW/年」但數值為 $B/MW（差 1,000 倍）；(b) `IFW_RevGWFleet_*` 的 IFC_Use 未標「上限，不得作預測」，與數值相同的 `IF_RevGWFleet` 不一致；(c) 缺口表 2 的 1–3 與 6 是否走 G2。
6. **託管租金是否屬產業級數據**（缺口 5）：若屬，進 Tokenomics `SRC_`；若屬公司層，留在 WhiteFiber `data/`。建議後者（目前只有 WhiteFiber、Oracle 有託管分部）。
