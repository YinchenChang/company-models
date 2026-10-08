# Nebius｜Tokenomics 連接審視（v5.29；2026-10-08）

- 分支：`claude/nebius-tk-link-review`
- 最新提交 SHA：見 PR 頁面（本檔所在提交；以 `git log -1 -- nebius/docs/reports/20261008_nebius_tokenomics_link_review.md` 查得）
- 報告路徑：`nebius/docs/reports/20261008_nebius_tokenomics_link_review.md`
- 依據：共用工作單 `docs/workorders/20261008_tokenomics_link_review.md`、`nebius/docs/workorders/20261008_nebius_tokenomics_link_review.md`；Tokenomics 下游資料契約 v0.1（`docs/plan/Tokenomics_downstream_contract.md`）。
- 審視對象：company-models main `75327ef` 的 Nebius 模型（成品 v0.2，`dist/20261007_Nebius收支模型_v0_2.*`）；另評估進行中的 v0.2a 分支（`claude/nebius-v0.2a-tkanchor` `43b1f9c`，PR #30，draft、未合併）在 v5.29 下需要的調整。
- Tokenomics：master `a5061d9`，`model/CURRENT`＝`20261008_Tokenomics_v5.29.xlsx`（v5.29 已合併，`IFW_`、`IFC_` 名稱存在）→ 前提成立，第 1–4 節全做。
- **本單不改模型**：company.json、builder、Excel、data 均未動；v5.29 快照只產生在暫存路徑 `/tmp/nebius_v529_snapshot.json`，未提交。
- **與 PR #30 的檔案衝突**：PR #30 也新增了同名檔（v5.27 口徑、v5.29 未合併前寫的版本）。兩個 PR 先合併者生效，後者會衝突；建議以本檔（v5.29）為準，PR #30 合併前刪除或改名其同名檔（見「待 chat 端判斷」第 6 項）。

## 結論（先讀這裡）

1. **main 上的 Nebius（v0.2）還沒有照契約接上 Tokenomics**：取數仍是手工 JSON、釘在 v5.24（`098873a`）；有 4 處禁止的工作表直接引用（`DC_Cost!` 2 處、`Inputs!` 2 處），另有 3 組寫死的 Tokenomics 數字（每 MW 收入三情境、每 MW 建置成本、PUE）。好消息是：**Nebius 用到的 Tokenomics 數值從 v5.24 到 v5.29 全部沒變**（差異只在手工 JSON 取兩位小數的四捨五入），所以重取不會改變任何目標價。
2. **每 MW 收入仍是「成本加成與市場價 50/50 平均」，且容量與單價同一個情境選擇器同向綁定**（保守 11.62／基準 17.40／積極 24.20 跟著容量情境走）——兩點都違反契約第 4 條。進行中的 v0.2a 已改為「Tokenomics 持有成本錨 × 定價倍數 k」並把兩條情境軸分開，方向正確。
3. **v0.2a 在 v5.29 下數字不必動**（26 個名稱 320 格逐格比較 0 差異），但要補三件契約新要求：(a) 版本改釘 v5.29（`a5061d9`）；(b) 上限檢查要同時取四層瀑布 `IFW_RevGWFleet_100／_Util／_Prod／_Life`、並遵守 `IFC_Use`（`IF_RevGWFleet` 標「上限，不得作預測」，v0.2a 只拿來做上限檢查，合規）；(c) 契約第 3–4 條新給的「世代產出比」`L1_TokMW_Gen_ratio_*` 尚未使用——v0.2a 的錨隱含「VR200 每 MW 收入≈GB300」（持有成本比 1.003），而 Tokenomics 的每 MW 產出比是 1.565（0.915–2.051），兩者差在「每 token 價格下降多少」，需要 chat 端判斷是否把 q × p 列為價格軸的對照。

## 1. 取數點清單

### 1a. 工具結果與人工確認

`python3 tools/tokenomics/link_audit.py --tokenomics /home/claude/tk-master`（對照 v5.29、`a5061d9`）在 Nebius 節的結果：取數方式＝手工 JSON（`permw_tokenomics_20261007.json`）；釘住 `model/20261007_Tokenomics_v5.24.xlsx`（`098873a`，落後）；使用名稱 10 個；不存在於現行 0；違規引用 3；值漂移 0；四層瀑布／用途欄＝否／否。

人工逐筆確認工具列出的 3 筆違規（以「檔案 × 前綴」計）：

| 工具列出 | 人工確認 | 實際內容 | 是否驅動模型 |
|---|---|---|---|
| `DC_Cost!`（permw 檔） | **確實違規**（不是 Nebius 自己的工作表；Nebius Excel 沒有 `DC_Cost`、`Inputs` 分頁） | `DC_Cost!I48:N48`（DC 營運費用，用於隱含 EBITDA 率檢核）；`DC_Cost!J22`（每 GPU 儲存與管理 $500，只在缺口分解文字中說明） | 否（只作檢核與說明） |
| `Inputs!`（permw 檔） | **確實違規** | `Inputs!D23:F23`（IT 折舊年限 6／6／4）、`Inputs!D28:F28`（WACC 8／10／13%），標「僅供參考」 | 否（Nebius 折舊用公司會計 5 年、WACC 用自己的 11%） |
| `Inputs!`（facts 檔） | **確實違規** | `tok.itLife`、`tok.wacc` 兩筆來源欄寫 `Inputs!D23:F23`、`Inputs!D28:F28` | 否（對應的 `company.json → nebius.perMw.gpuLifeYears` 6 年只記錄、未映射；`defaults.gpuLife`＝5） |

工具沒有抓到、人工補列的不合規：

- **寫死數字（3 組）**：`scenarios.revMW` 與 `defaults.m.revMW`（0.01162／0.0174／0.0242 US$bn/MW）；`scenarios.capexTemplate.costMW`（50.12／50.26）；`defaults.m.pue`（1.2）。另 HTML 與 Excel 的說明文字（`segA`、`segD`、`build_xlsx.py`）寫死「50.12／50.26」「11.62／17.40／24.20」等數字。
- **以儲存格位址記來源**：permw 與 facts 檔的 `Interface!J7`、`Interface!C13:Q13` 等 25 處。`Interface` 是 IF_ 名稱所在頁，數值本身合規，但契約要求以名稱存取；v5.29 起 Interface 列號已因新增 J 節而有位移風險（這些格在 v5.29 仍在原位，未受影響）。

**違規引用數（人工確認後）**：工作表直接引用 4 處（`DC_Cost!` 2、`Inputs!` 2，分布在 2 個檔、工具計 3 筆）＋寫死數字 3 組；無誤判。所有違規都不經公式連動，只要重取並改記名稱即可消除。

### 1b. 取數點清單表（main，v0.2）

| 本模型的格或參數 | Tokenomics 名稱 | 合規或違規 | 四層瀑布哪一層 | 用途是否符合 `IFC_Use` | 所用版本 |
|---|---|---|---|---|---|
| 每 MW 年收入三情境 `scenarios.revMW`、`defaults.m.revMW`（0.01162／0.0174／0.0242） | `IF_HoldEcon`（路徑 A 成本加成的底）＋ `IF_GPUsPerGW`（路徑 B 市價換算） | **違規**：寫死數字；方法為 50/50 平均（契約第 4 條停用） | 不適用（持有成本＝100% 時數口徑）；路徑 B 另乘自訂 85% 可計費利用率 | 部分不符：`IF_HoldEcon` 的用途是「情境、相對比較、反轉門檻」，作成本加成收入輸入違反契約第 4 條（成本加成只作 ROIC／缺口檢查） | v5.24 `098873a` |
| 每 MW 建置成本 `scenarios.capexTemplate.costMW`（50.12／50.26） | `IF_CapexTotal`（GB300／VR200 基準） | **違規**：寫死數字（值與 v5.29 相同） | 不適用 | 符合（「情境、相對比較、反轉門檻」；作 CapEx 情境輸入） | v5.24 |
| PUE `defaults.m.pue`（1.2） | `IF_FacilityGW`（基準） | **違規**：寫死數字（值相同） | 不適用 | 符合（「基準輸入」） | v5.24 |
| GPU 壽命對照 `nebius.perMw.gpuLifeYears`（6，只記錄） | 原 `Inputs!D23:F23` → 應為 `IF_DeprLifeIT` | **違規**：工作表直接引用 | 不適用 | 符合（只作對照） | v5.24 |
| WACC 對照（facts `tok.wacc`，只記錄） | 原 `Inputs!D28:F28`；**無對應 IF_ 名稱** | **違規**：工作表直接引用 | 不適用 | — | v5.24 |
| 隱含 EBITDA 率檢核（permw `dcOpexPerMWit`） | 原 `DC_Cost!I48:N48` → 應為 `IF_OpexGW`（v5.26 起，數值相同 2.63） | **違規**：工作表直接引用 | 不適用 | 符合（只作檢核） | v5.24 |
| 每 GPU 儲存與管理（缺口分解文字） | 原 `DC_Cost!J22`；無 IF_ 名稱 | **違規**：工作表直接引用（只在說明文字） | 不適用 | — | v5.24 |
| 路徑 B 價格換算、VR200 類比（permw） | `IF_GPUsPerGW`、`IF_GPUhrEcon` | 合規名稱，但以 `Interface!` 位址記錄 | 不適用（`IF_GPUhrEcon` 為 100% 利用率） | 符合（相對比較） | v5.24 |
| 對照用（permw，不驅動） | `IF_CapexIT`、`IF_CapexFacility`、`IF_HoldAcct`、`IF_PowerCost` | 合規名稱，以位址記錄 | 不適用 | 符合 | v5.24 |
| 利用率對照（permw 註記：不沿用） | `IF_Util`（60%） | 合規；明文不進收入 | IF_Util 層（未使用） | 符合 | v5.24 |
| 上限檢查 | 無（main 未做） | — | — | — | — |
| 四層瀑布 `IFW_*`、用途欄 `IFC_*` | 未取 | **缺**（契約第 2 條第 2 項） | — | — | — |

### 1c. 進行中 v0.2a（PR #30）的取數現況（供對照）

v0.2a 已改用官方快照 `data/tokenomics_snapshot_v5.27.json`（26 名稱，`IF_`／`L1_`），Excel 新增「Tokenomics_取數」分頁與 110 個 `TK_` 名稱；permw 與 facts 檔的工作表引用已改寫為名稱、檔頭標「已被快照取代」；每 MW 收入改為 `IF_HoldEcon` × 世代占比 × k，`IF_RevGWFleet` 只作上限檢查。仍未處理：`costMW` 寫死（該分支自己列為待辦）、`IFW_`／`IFC_` 未取、版本釘 v5.27。

## 2. 口徑問題（四答）

1. **計價方式：按 MW-year**（神雲）。可計費 MW＝已連網 MW × 在役比例（0.75／0.85／0.90…）× 爬坡係數（60％／80％／100％）；利用率欄 100%；`IF_Util` 不進收入。契約第 3 條：MW-year 收入取決於商業簽約率（Nebius 新產能簽約率固定 100%），**沒有 `IF_Util` 的重複扣減**。但 v0.2 的每 MW 收入在路徑 B 內含一個自訂「可計費利用率」85%（保守 80%、積極 90%），與在役比例、爬坡係數同時作用，**有部分重疊扣減的疑慮**（不是 `IF_Util`，但性質都是「容量沒賣滿」）；v0.2a 停用路徑 B 後此疑慮消失。handoff 目前未寫明「採 MW-year 計價」（v0.2a 已補）。
2. **50/50 平均：是，仍在用**（保守＝(13.54＋9.70)÷2＝11.62；基準＝(15.91＋18.89)÷2＝17.40；積極＝(18.38＋30.03)÷2＝24.20，permw 檔 `scenarios`）。**容量與單價同向綁定：是**——`scenarios.revMW.low／base／high` 與容量路徑同一個情境選擇器，保守容量配低價、積極容量配高價。兩者都違反契約第 4 條（應分離；CoreWeave W4 為範本）。v0.2a 已改：容量軸三情境預設都用基準 k，價格軸（k 低／基準／高）另列 3 × 3 矩陣。
3. **`IF_Alloc*` 六格：不適用**。Nebius 不是模型商，main 與 v0.2a 都沒有取用任何 `IF_Alloc*`（全資料夾搜尋 0 筆）；v5.29 這六格的變動（Q1 0.636→0.573、ServeGW 0.0519→0.0675、隱含 N×k 1.524→1.983 等）**不影響 Nebius 任何輸出**。
4. **公司特有變數只在本模型覆寫、沒有回寫**：WACC 11%、GPU 壽命 5 年（Tokenomics 6 年只作對照）、在役比例、爬坡係數、新產能簽約率、合約 MW 上限、併網速度、預付比例都只在 `company.json`；Nebius 沒有任何寫入 Tokenomics 的程式或檔案。v0.2a 新增的 k、長約占比、世代組合也只在 `company.json → pricing／fleet`。合規。

## 3. 缺口與名稱清單草稿

### 3a. Tokenomics 缺的產業級數據（不自行建表；交 chat 端走 G2 進 DB_Evidence）

| 名稱 | 用途 | 目前以什麼代替 |
|---|---|---|
| GPU 租價分層時間序列（隨需／預留／長約；GB200、GB300、B300、VR200） | 契約第 4 條的 p；Nebius k 的證據 | v5.29 已在 SRC_Price 開 X–AB 五個分層欄，但**全部留空**。Nebius 以自蒐證據代替：IREN–Microsoft GB300 五年約（9.70 US$M/MW-IT/年）、Silicon Data H100 現貨指數 2.82、B300 指數、Verda GB300 隨需 10.42／預留 7.82／spot 5.21、Nebius 官網預留折扣最多 35% |
| 第二筆獨立 neocloud 長約（含年期與 GPU 數） | k_長約 的第二個證據點 | 只有 IREN–Microsoft；Nscale–Microsoft 未揭露年期、Lambda–Microsoft 未揭露金額（v0.2a 證據表列「找不到」） |
| Hopper H200 獨立欄 | 期初機隊世代組合的錨 | 以 Hopper H100 欄代表 |
| 每 GPU 儲存與管理成本的具名範圍 | 缺口分解（非 GPU 服務） | 原引用 `DC_Cost!J22`（禁止）；無 IF_ 名稱 |

建議送 G2 的候選：IREN–Microsoft 長約、Silicon Data H100／B300 指數、Verda GB300 三層價、Nebius 預留折扣（填 SRC_Price 長約／預留／現貨層）。

### 3b. 名稱清單草稿（`data/tokenomics_names.txt`；只列 IF_／IFW_／IFC_／L1_；**本單不執行、不提交**）

```
# 沿用 v0.2a 清單（26）
IF_RacksPerGW
IF_GPUsPerGW
IF_FacilityGW
IF_CapexIT
IF_CapexFacility
IF_CapexTotal          # 取代寫死的 costMW 50.12／50.26
IF_HoldAcct
IF_HoldEcon            # 每 MW 收入錨（與 L1_HoldEconMW_* 同值 ×1000）
IF_GPUhrEcon           # k 證據的分母
IF_PowerCost
IF_Util                # 只作對照；MW-year 計價不進收入
L1_FacCapexMW
L1_GPUhr_GB200_vsCW
L1_GPUhr_GB300_vsBE
L1_RevGW_Fleet_VR200
IF_DeprLifeIT          # 取代 Inputs!D23:F23
IF_DeprIT
IF_DeprFac
IF_AvgDraw
IF_PowerPrice
IF_MaintIT
IF_MaintFac
IF_StaffSW
IF_TaxIns
IF_OpexGW              # 取代 DC_Cost!I48:N48
IF_RevGWFleet          # IFC_Use：上限，不得作預測 → 只作上限檢查
# v5.29 新增（契約第 3 條：MW-year 下游的兩個介面）
L1_HoldEconMW_Hopper
L1_HoldEconMW_GB200
L1_HoldEconMW_GB300
L1_HoldEconMW_VR200
L1_TokMW_Gen_ratio_GB200
L1_TokMW_Gen_ratio_GB300
L1_TokMW_Gen_ratio_VR200
L1_FleetBreakeven      # 只作對照
L1_FleetMargin         # 只作對照（含折減）
# v5.29 四層瀑布與用途欄（現行 import_tokenomics.py 不接受 IFW_／IFC_，見待判斷第 4 項）
IFW_RevGWFleet_100
IFW_RevGWFleet_Util
IFW_RevGWFleet_Prod
IFW_RevGWFleet_Life
IFC_Use                # 只讀上列各名稱所在列的用途文字
IFC_Conf
IFC_Layer
IFC_Updated
```

## 4. v5.29 重取影響

### 4a. 執行方式

- `python3 tools/tokenomics/import_tokenomics.py --tokenomics /home/claude/tk-master --names <草稿前 35 個 IF_／L1_ 名稱> --out /tmp/nebius_v529_snapshot.json` → 成功（`20261008_Tokenomics_v5.29.xlsx`，commit `a5061d9`，35 個名稱，missing 0，未重算、直接讀快取值）。
- 工具只接受 `IF_`／`L1_`，`IFW_`、`IFC_` 名稱以 openpyxl 直接讀 v5.29 的具名範圍代替（Interface 第 143、248–251、268–271 列；R–U 欄）。
- `--check` 需要 main 上已有快照，Nebius main 沒有（仍是手工 JSON），故改為兩組比較：(1) main 手工 JSON（v5.24）各值對 v5.29；(2) v0.2a 快照（v5.27）對 v5.29。

### 4b. 比較結果

| 比較 | 格數 | 差異 | 說明 |
|---|---|---|---|
| main 手工 JSON（v5.24，GB300／VR200 × 三成本情境 × 10 欄）對 v5.29 | 60 | **0（超出四捨五入者）** | 手工 JSON 取兩位小數；例：`IF_HoldEcon` GB300 基準 12.72 對 12.7247、VR200 12.76 對 12.7620；`IF_GPUsPerGW` 487.0 對 487.008 |
| v0.2a 快照（v5.27，26 名稱）對 v5.29 | 320 | **0** | 契約說的 v5.29 只有 `IF_Alloc*` 六格改變，Nebius 未用 |
| `IF_Alloc*` | — | 不適用 | Nebius 未取用 |
| `IFW_RevGWFleet_*` 相對 v0.2a 所用層 | — | 新增 | v0.2a 上限檢查用 `IF_RevGWFleet`＝`_Util` 層；四層數值見下表 |

**結論：v5.29 重取對 Nebius 三情境目標價、每 MW 收入、CapEx 的影響都是 0**；要做的是版本、名稱與契約欄位的補齊。

### 4c. 新介面對 Nebius 的意義

**(1) 上限檢查換層會改變結論**（每 GW 客戶付費 token 營收，US$B/GW/年；v5.29 基準欄）：

| 世代 | `_100` | `_Util`（＝`IF_RevGWFleet`） | `_Prod` | `_Life` | 持有成本 `IF_HoldEcon` ÷ `_Util` | ÷ `_Prod` | v0.2 基準 17.40 ÷ `_Util` | ÷ `_Prod` |
|---|---|---|---|---|---|---|---|---|
| Hopper | 6.68 | 4.00 | 2.24 | 4.00 | 2.52 | 4.50 | 4.34 | 7.76 |
| GB200 | 33.62 | 20.17 | 15.29 | 20.17 | 0.45 | 0.60 | 0.86 | 1.14 |
| GB300 | 43.78 | 26.27 | 19.23 | 26.27 | 0.48 | 0.66 | 0.66 | 0.90 |
| VR200 | 81.21 | 48.72 | 38.12 | 48.72 | 0.26 | 0.33 | 0.36 | 0.46 |

- v0.2a 的警示門檻是 50%（`methodology.checks.revCapShareMax`）。用 `_Util` 層時，GB300 只收持有成本（k＝1）就占客戶 token 營收 48%，貼著門檻；改用 `_Prod` 層則 66%，必定警示。Hopper 在任何層都 > 100%（Tokenomics 認為 Hopper 跑前沿模型的 token 營收低於持有成本）。**用哪一層當分母是判斷題**（見待判斷第 2 項）。
- **四層不單調**：`_Life`（4.00／20.17／26.27／48.72）等於 `_Util`、且大於 `_Prod`。依契約字面，層次應為 100% ≥ Util ≥ Prod ≥ Life。可能是 L × m 套在 Util 層而非 Prod 層、或 L、m 基準為 1 的設計；屬 Tokenomics 端問題，只列不修。
- `IFC_Use` 只有 `IF_RevGWFleet`（Interface 第 143 列）標「上限，不得作預測」；四層瀑布的 `IFW_RevGWFleet_*` 列只標「情境、相對比較、反轉門檻」、沒有「上限」字樣，與同值的 `IF_RevGWFleet` 不一致（屬 Tokenomics 端，只列）。`IFW_RevGWFleetFront_Prod／_Life` 為「—」，用途欄空白。
- Nebius 用到的成本列 `IFC_Conf`：`IF_HoldEcon`、`IF_GPUhrEcon`、`IF_CapexTotal`、`IF_OpexGW` 為 B；`IF_FacilityGW`、`IF_PowerCost` 為 A；`IFC_Updated` 最弱輸入審查日 2026-10-05。可在 Nebius 的「Tokenomics_取數」分頁並列。

**(2) 契約第 3 條給 MW-year 下游的兩個新名稱**：

- `L1_HoldEconMW_*`（每 MW 年經濟持有成本）：Hopper 0.01010、GB200 0.00917、GB300 0.01272（0.00871–0.02387）、VR200 0.01276（0.00984–0.02068）。**與 `IF_HoldEcon ÷ 1000` 完全相同**，v0.2a 的錨可原地改指這組名稱，數字不變。注意：**L1 G 欄單位寫「$M/MW/年」，但數值 0.0127 實為 $B/MW/年（＝12.72 $M/MW/年）**，屬 Tokenomics 標籤錯誤，引用時須自行 ×1000，只列不修。
- `L1_TokMW_Gen_ratio_*`（相鄰世代每 MW 產出比，Sol、100%）：GB200÷Hopper 4.845；GB300÷GB200 1.456；**VR200÷GB300 1.565（0.915–2.051）**。v0.2a 的錨以各世代持有成本加權，VR200 對 GB300 的錨比只有 1.003——等於隱含「每 MW token 產出 ×1.565、每 token 價格 ×0.64」，正好是契約第 4 條的「中性情境 q × p ≈ 1」。v0.2a 尚未明示這個隱含假設，也沒有用 q 做價格軸的對照（見待判斷第 1 項）。

## 待 chat 端判斷

建議改動（依優先順序；全部由 chat 端另開工作單，本單未動）：

1. **（高）v0.2a 合併前或合併後立即：版本改釘 v5.29（`a5061d9`）並補四層瀑布與用途欄**。重取快照（數字 0 變動，可用 `--check` 證明）、handoff 改記 v5.29；上限檢查改為四層並列，並決定以哪一層當警示分母——建議 `_Util`（與 MW-year 下客戶自己的實際負載同口徑），`_Prod` 並列為保守對照；Hopper 分項另註「Tokenomics 判斷 Hopper token 營收低於持有成本，上限檢查不適用」。
2. **（高）判斷：錨的世代進步要不要接 q × p**。v0.2a 以持有成本為錨，等於假設每 MW 價格跟著成本走（VR200 ≈ GB300），隱含 q × p ≈ 1（中性）。選項：(a) 維持，於 handoff 明示「錨＝中性情境 q × p ≈ 1，對應每 token 價格每代 −36%」；(b) 價格軸加一條 q × p 對照（q＝`L1_TokMW_Gen_ratio_VR200` 0.915–2.051，p＝SRC_Price 每 token 價格下降，待 G2 補資料）。建議 (a)＋在敏感度列 (b) 的區間，不改基準。
3. **（中）main 若在 v0.2a 合併前另有發版**：手工 JSON 的 4 處禁止引用與 3 組寫死數字一併改為快照名稱（v0.2a 已做大半；`costMW`、`pue` 改引 `TK_CapexTotal_*`、`TK_FacilityGW_*` 兩項 v0.2a 仍未做，建議併入 v5.29 重取工作單）。
4. **（中）工具（repo 層，非 Nebius）**：`import_tokenomics.py` 只接受 `IF_`／`L1_`，無法取 `IFW_`／`IFC_`，與契約第 2 條第 1 項不一致；建議擴充（IFC_ 依名稱列號讀 R–U 欄）。`link_audit.py` 對 Nebius 本次沒有誤判，但不抓寫死數字與 `Interface!` 位址記錄，建議加這兩項提示。
5. **（低）回報 Tokenomics 的三個疑點**（只列，Nebius 端不修）：(a) `L1_HoldEconMW_*` 單位標籤 $M/MW/年 與數值 $B/MW/年 不符；(b) `IFW_RevGWFleet_Life`＝`_Util` > `_Prod`，四層不單調；(c) `IF_RevGWFleet` 的 `IFC_Use` 標「上限」，同值的 `IFW_RevGWFleet_Util` 卻沒有。
6. **（流程）同名報告檔衝突**：PR #30 也新增 `nebius/docs/reports/20261008_nebius_tokenomics_link_review.md`（v5.27 口徑、第 4 節「待 v5.29」）。建議以本 PR 的 v5.29 版為準，PR #30 合併前移除其同名檔，或改名為 `20261008_nebius_tokenomics_link_review_v527.md` 保留歷史。

需要判斷的口徑問題（摘要）：上限檢查的分母層（第 1 項）；錨的世代進步是否接 q × p（第 2 項）；路徑 B 舊有的 85% 可計費利用率與在役比例、爬坡的重疊（v0.2a 停用路徑 B 後自然消失，main 若再發版需確認，第 2 節第 1 答）。
