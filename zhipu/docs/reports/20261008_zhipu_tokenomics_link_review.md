# 智譜：Tokenomics 連接審視（v5.29）＋ CI 失敗診斷

- 分支：`claude/zhipu-tk-link-review`（底稿 main `75327ef`）
- 最新提交 SHA：報告內容提交 `dafad92`（其後只補記本行與文末提交紀錄）
- 報告檔：`zhipu/docs/reports/20261008_zhipu_tokenomics_link_review.md`
- 依據：`docs/workorders/20261008_tokenomics_link_review.md`（共用）、`zhipu/docs/workorders/20261008_zhipu_tokenomics_link_review.md`；Tokenomics 下游契約 v0.1（`docs/plan/Tokenomics_downstream_contract.md`）。
- 對照對象：Tokenomics master `a5061d9`，`model/CURRENT`＝`20261008_Tokenomics_v5.29.xlsx`（v5.29 已合併，第 4 節前提成立）。
- **本輪只審視、不改模型**：company.json（本模型無此檔）、builder、Excel、data 皆未改；只新增本報告。

## 零、結論先行

1. **取數方式與版本**：本模型沒有快照 JSON，但也不是「無」——builder（`builder/tk_link.py`）在建置時直接開 Tokenomics 的 `model/CURRENT` 檔，逐名讀值寫入活頁簿的 `TK_Link` 頁（74 個具名範圍＋7 列「讀表」），公式只引用 `TK_` 名稱。現行 Excel 釘在 **`20261007_Tokenomics_v5.26.xlsx`（master `70d859e`）**，落後 v5.29 三個版本。工具盤點寫「無／未釘版本」是誤判（工具只認 `data/tokenomics_snapshot_*.json`）。
2. **實際被公式用到的只有 16 個 Tokenomics 名稱**（另 58 個是沿用 OpenAI v0.6 一起搬進來、沒有任何公式引用的「死名稱」）。
3. **違規引用：7 項確定、1 項灰色**。確定違規＝7 列以列標籤從 Tokenomics 內部頁 `Cap_In`（2 列）、`Price_Frontier`（5 列）讀表（契約第 2 條第 1 項）；其中 4 列被 Revenue 對照列以儲存格位址引用。灰色＝把 `IF_HdrGen`／`IF_HdrCost` 表頭複製成 `TK_HdrGen`／`TK_HdrCost` 供 SUMIFS 選欄（官方快照工具內部也用這兩個表頭拆欄，建議改用工具的拆欄結構）。工具報的 2 項 `Inputs!` 是本模型自己的 Inputs 頁，**誤判**。
4. **v5.29 重取：命題輸出不變**。81 個值中 75 個與 v5.29 完全相同；只有 `IF_Alloc*` 六格變動（Q1 −9.9%、Q1_R2 −5.7%、Q2 −9.2%、ServeGW +30.1%、Demand +43.6%、ImpliedNk +30.1%），**這六格沒有任何公式引用**；唯一被引用的 `IF_AllocRDGW`（研發 GW 對照列）不變。
5. **真正會動數字的是「換層」而不是「換版」**：本模型的 VR 等值換算用的是 100% 層產能（`IF_TokGW_Sol`）。若依契約第 2 條第 2 項改看生產層（`IFW_TokGW_Sol_Prod`），Hopper ÷ VR200 的產能比由 0.0906 降為 0.0819（−9.6%），**每 VR 等值 GW 差額（命題 1）的絕對值約放大 10.6%**（2030 約 −2,275 → −2,516 RMB 億；覆蓋率不變）。這是判斷類，列在「待 chat 端判斷」。
6. **CI 失敗原因**：main 上的智譜資料夾是「半合併」狀態——資料檔 `data/zhipu_src.yaml` 已有 1,054 列（chat 端提交 `767f56a` 擷取招股章程 SRC_ZP_611–1054），但 Excel、ID 登記檔與測試期望值仍停在 610 列；把三者一起更新的 r2 P1–V9 共 7 個提交留在 `claude/zhipu-v0.1` 分支、從未合併。不是模型算錯，也不是期望值寫錯，而是 main 缺了 7 個提交。

## 一、取數點清單（共用工作單第 1 節）

### 1.1 實際進入公式的名稱（16 個＋表頭 2 個）

「層」依 Tokenomics v5.29 的 `IFC_Layer`；「用途」依 `IFC_Use`。版本欄：全部為 `20261007_Tokenomics_v5.26.xlsx`、master `70d859e`（Excel `TK_Link` B3／B5）。

| 本模型的格（頁!列） | Tokenomics 名稱 | 合規？ | 瀑布層 | 用途是否符合 `IFC_Use` | 對 v5.29 |
|---|---|---|---|---|---|
| Compute G04、G06（第 11、13 列）；Cost K63、K64 | `IF_TokGW_Sol` | 合規（IF_） | 100% | 符合：只作「產能比」（VR 等值係數，屬相對比較）與 TK 參考列；未當收入預測。注意此名 `IFC_Use` 含「上限，不得作預測」 | 不變 |
| Compute G05（第 12 列） | `IF_TokGW_Luna` | 合規 | 100% | 同上（舊法對照列） | 不變 |
| Compute G53–G56、G57–G63（舊法）；Cost K63、K64 | `IF_Util` | 合規 | IF_Util | 符合：只用在「token 換算所需 GW」（成本端），不扣收入 | 不變 |
| Compute G42、G164–G166 | `IF_HoldEcon` | 合規 | —（成本） | 符合：作 token→GW 換算的分母與供應商持有成本（租價下限） | 不變 |
| Compute G36–G41（第 46–51 列） | `IF_CostPre_Sol／Luna`、`IF_CostCache_Sol／Luna`、`IF_CostDec_Sol／Luna`（6 名） | 合規 | 100%（經濟口徑、100% 利用率） | 字面符合（用於算「需要多少 GW」，非收入）；但 `IFC_Use` 標「上限，不得作預測」，以 100% 產能換算的 GW 是**下界**，落差由 η 吸收（見第二節問題 3） | 不變 |
| Compute G99（第 114 列） | `IF_CapexTotal` | 合規 | — | 符合（自有 GW 每 GW 資本支出 × 國產比例） | 不變 |
| Compute G172（第 193 列） | `IF_OpexGW` | 合規 | — | 符合（自有算力營運費用） | 不變 |
| Compute G174（第 195 列） | `IF_DeprLifeIT` | 合規 | — | 符合（攤提年限） | 不變 |
| Cost K63（第 77 列，TK 參考） | `IF_FullCostDefault_Sol` | 合規 | IF_Util | 符合（只列對照，不進命題） | 不變 |
| Cost K64（第 78 列，TK 參考） | `IF_FullCost_Sol` | 合規 | IF_Util | 符合（同上） | 不變 |
| Compute G115（第 131 列，對照） | `IF_AllocRDGW` | 合規 | — | 符合（只列對照；標明「OpenAI 規模」，與決策 G16 一致） | 不變 |
| Compute 第 11–195 列多處 SUMIFS 的選欄 | `IF_HdrGen`、`IF_HdrCost`（複製為 `TK_HdrGen`／`TK_HdrCost`） | **灰色**（契約禁止連結 `IF_Hdr*`；此處是複製表頭文字供選欄，不是取值） | 不適用 | 不適用 | 表頭相同 |

### 1.2 讀進 TK_Link、但沒有任何公式引用的名稱（58 個）

`IF_TokGW_Astra`、`IF_FullCost_Luna／Astra`、`IF_Price{Fresh,Cached,Think,Out,Ref}_*`（15）、`IF_FrontRef_*`（3）、`IF_ProgGWyr_*`（3）、`IF_RDMult`、`IF_TrainGenDefault`、`SRC_DEM_004–013`＋`013_Lo／Hi`（12）、`IF_AllocQ1`、`IF_AllocQ1_R2`、`IF_AllocQ2`、`IF_AllocServeGW`、`IF_AllocDemand`、`IF_AllocImpliedNk`、`IF_TrainCost_*`（3）、`IF_RevGW_*`（3）、`IF_FullCostDefault_Luna／Astra`、`L1_Ans3`、`_Lo`、`_Hi`、`IF_Cost{Pre,Cache,Dec}_Astra`（3）。

全部合規（IF_／L1_／SRC_ Active），但都是從 OpenAI v0.6 整批搬來的。留著不違規，只是讓「這個模型到底依賴 Tokenomics 的哪些數」不易看清；建議在改用官方快照時刪掉（見第三節草稿）。

### 1.3 違規引用

| # | 內容 | 位置 | 判定 |
|---|---|---|---|
| 1–2 | `Cap_In` F 表 GLM-5.3、GLM-5.3-Flash 兩列（以列標籤找列，讀 E 起 5 欄） | `builder/tk_link.py` `TABLE_ROWS`；TK_Link 第 84–85 列；Revenue R78、R79 以 `TK_Link!$J$84`、`$L$84` 位址引用 | **違規**（Tokenomics 內部頁直接讀取） |
| 3–7 | `Price_Frontier` A 表 GLM 兩列、B 表中國合格前緣三列 | 同上；TK_Link 第 86–90 列；Revenue R80、R81 以 `TK_Link!$J$88`、`$K$86` 引用 | **違規**（同上） |
| 8 | `IF_HdrGen`／`IF_HdrCost` 表頭複製 | TK_Link 第 7–8 列 → `TK_HdrGen`／`TK_HdrCost` | 灰色（見 1.1） |
| — | `Inputs!`（`builder/build.py`、`builder/z2.py`） | 本模型自己的 Inputs 頁 | **工具誤判，不算** |

影響範圍：違規的 7 列只進 Revenue 第 R78–R81 的「對照列」，沒有任何計算頁再引用這 4 格，**不影響命題輸出**。

### 1.4 工具盤點的其他誤判（`tools/tokenomics/link_audit.py`）

- 「取數方式：無；釘住 —」：工具只認快照 JSON；本模型版本記在 Excel `TK_Link` 頁（v5.26、`70d859e`）。
- 「使用名稱 52、不存在 5」：工具把程式裡的正規表示式（如 `IF_Cost(Pre|Cache|Dec)_…`、`IF_Price(Fresh|…)`）截成 `IF_Cost`、`IF_Price` 等假名稱，又把 `IF_TokGW_`、`IF_FullCost_`、`IF_RevGW_` 字首展開成 Tokenomics 全部同字首名稱（因此列出了本模型根本沒讀的 `_Prod`、`_Life`）。實際讀入 74 名、全部存在於 v5.29。

## 二、口徑問題（共用工作單第 2 節）

**問題 1：收入按什麼計價？利用率與簽約率各用哪一個？有沒有重複扣減？**
- 收入**按 token 計價**（API 按量、Coding Plan 訂閱換算 token、智譜清言），單價與用量全部來自公司資料（SRC_ZP）與本模型假設（Inputs），**不取 Tokenomics 的價格或每 GW 營收**。
- 契約第 3 節：按 token 計價應用「技術利用率」。本模型用 `IF_Util`（60%）——但**只用在成本端**（把 token 換算成需要的 GW），收入端沒有任何利用率或簽約率扣減。**收入端無重複扣減。**
- 值得注意的一點（成本端）：2H26 起的供給 GW＝有效推論 GW ÷［(1−閒置保留 10%，INP_108)×(1−研發占比)］。`IF_Util` 60% 已含尖離峰空檔，再加 10% 閒置保留，等於實際負載只有約 54%。這不是契約說的「收入重複扣減」，但兩者可能部分重疊（列入待判斷第 4 項）。
- handoff 目前沒有寫明「採技術利用率、不用簽約率」（契約第 3 節要求）；建議下一版補一句。

**問題 2：神雲／雲平台的每 MW 收入（50/50 平均、選擇器綁定）**
- 不適用：智譜是模型商，不出租算力。
- 相關但不同的一點：本模型向雲廠商**租用**算力的每 GW 價格（成本），2H26 起取「觀察租價」與「供應商持有成本 ×（1＋最低毛利）」的較大者（V1，Compute G161–G171）。這是**成本下限**，不是 50/50 平均，也不是收入輸入，與契約第 4 條不衝突。容量（供給 GW）由需求推、單價由租價推，兩者是不同的 Inputs，**沒有同一選擇器同向綁定**。

**問題 3：`IF_Alloc*` 六格的取用處與受影響輸出**
- 六格（Q1、Q1_R2、Q2、ServeGW、Demand、ImpliedNk）**都讀進 TK_Link，但沒有任何公式引用**。研發 GW 是本模型自己的殘差（供給 − 推論），不是用 Tokenomics 的 OpenAI 配置比例——這正符合決策 G16（Alloc 以 OpenAI 為代理，其他實驗室須自行覆寫）。
- 唯一引用的 `IF_AllocRDGW`（Compute G115，對照列）在 v5.29 不變（0.0907 GW·年）。
- **受影響的輸出：無。**
- 延伸（非 Alloc 但同源）：v5.29 新增的服務端落差分解 `L1_Gap*`（每則提示 token 數 2.0、支出計價口徑 2.0、ISL 2.31、利用率 1.46，乘積 `L1_GapProduct` 13.44）解釋的正是「用 Tokenomics 產能換算的 GW」與「用支出換算的 GW」之間的落差。本模型的 η（main 版 2025 為 12.44、1H26 為 9.46）就是同一個落差在智譜身上的量測值。現在 η 是一個黑箱倍數；若對照 `L1_Gap*` 四項逐項解釋，可看出哪些是口徑差、哪些是智譜真實效率（列入待判斷第 2 項）。另：`claude/zhipu-v0.1` 分支上未合併的 r2 P1 已把推論支出口徑改為「營業成本計算服務費＋研發項下算力費」，η 降為 1.13／3.91，顯示 main 版的大 η 主要來自支出口徑，與 `L1_GapSpendBasis` 的方向一致。

**問題 4：公司特有變數是否只在本模型覆寫、沒有回寫 Tokenomics？**
- 是。builder 只以唯讀方式開 Tokenomics 檔，輸出只寫到本模型的 `model/`；沒有任何寫回 Tokenomics 的程式。
- 本模型的公司變數：牌價與方案價（SRC_ZP）、各來源 token 量與組成（Inputs）、晶片組合、租價、研發占比、閒置、自有資本支出時程（INP_111–113）、匯率等，全在 Inputs／SRC_ZP。本模型沒有 WACC、價值捕獲率、簽約率（不適用）。
- 需要提的反向問題：有幾個**產業級物理數據**目前放在本模型的 Inputs（H20 與國產晶片的產出比、持有成本比、每卡功耗），依契約第 2 條第 4 項應走 G2 進 Tokenomics（見第三節缺口）。

## 三、缺口與版本（共用工作單第 3 節）

### 3.1 Tokenomics 缺的產業級數據（不自行建表；請 chat 端走 G2 進 DB_Evidence）

| 名稱 | 本模型用途 | 目前以什麼代替 |
|---|---|---|
| H20 每 GW 產出（相對 Hopper） | 晶片族產能、VR 等值係數 | INP_082 0.34（0.26–0.42，Analogy，DatabaseMart 估計） |
| H20 每 GW 持有成本（相對 Hopper）、每卡 IT 功耗 | 供應商持有成本（租價下限）、每 GW 卡數 | INP_114 0.6（0.4–0.8）、INP_097 1.09 kW |
| 國產晶片（昇騰 910B／910C）每 GW 產出、持有成本與資本支出比、每卡功耗 | 同上；自有 GW 資本支出、營運費用 | INP_083 0.6（0.4–0.9）、INP_115 0.6（0.4–0.9）、INP_098 1.0 kW |
| 中國市場 GPU-hour 租價（H800／A800、H20、昇騰） | 每 GW 年算力價格（成本） | SRC_ZP_462、474、478、479（公司資料層）；屬產業價格點，宜進 SRC_Price |
| GLM 系列 API 牌價、中國合格前緣混合單價 | Revenue 對照列 R78–R81 | 目前以列標籤讀 `Cap_In`／`Price_Frontier`（違規）；需 Tokenomics 提供 `IF_`／`L1_` 具名範圍 |
| 逐 token 類型單位成本的四層瀑布（`IF_Cost{Pre,Cache,Dec}_*` 的 `_Prod`／`_Life`） | token 換算推論 GW | 只有 100% 層；v5.29 的 `IFW_` 只涵蓋 TokGW／RevGW |
| 程式碼代理的 token 組成（快取命中、輸出占比） | Coding Plan 換算 GW | INP_122 0.9、INP_123 0.03（Analogy）；Tokenomics 有 `IF_TaskTok*`（Coding agent），可評估是否改用（判斷類） |

### 3.2 官方快照工具名稱清單草稿（`data/tokenomics_names.txt`；只附在報告，未建立檔案）

```
# 智譜 取數名稱（草稿 2026-10-08）
# 一、計算鏈實際引用
IF_TokGW_Sol
IF_TokGW_Luna
IF_Util
IF_HoldEcon
IF_CostPre_Sol
IF_CostCache_Sol
IF_CostDec_Sol
IF_CostPre_Luna
IF_CostCache_Luna
IF_CostDec_Luna
IF_CapexTotal
IF_OpexGW
IF_DeprLifeIT
# 二、對照列（不進命題）
IF_FullCost_Sol
IF_FullCostDefault_Sol
IF_AllocRDGW
# 三、建議新增：四層瀑布（契約第 2 條第 2 項）
IFW_TokGW_Sol_100
IFW_TokGW_Sol_Util
IFW_TokGW_Sol_Prod
IFW_TokGW_Sol_Life
IFW_TokGW_Luna_100
IFW_TokGW_Luna_Util
IFW_TokGW_Luna_Prod
IFW_TokGW_Luna_Life
# 四、建議新增：服務端落差分解（契約第 2 條第 5 項；η 的對照）
L1_GapPrompt
L1_GapSpendBasis
L1_GapISL
L1_GapUtil
L1_GapProduct
```

- 刪掉的 58 個死名稱（1.2 節）不列；`SRC_DEM_*` 未列（本模型沒用，且工具不收 `SRC_`）。
- 讀表 7 列不列；待 Tokenomics 提供具名範圍後再加。
- **工具限制**：`import_tokenomics.py` 目前只收 `IF_` 與 `L1_`，遇到 `IFW_` 直接報錯（`名稱不合規則：IFW_TokGW_Sol_100`）；契約第 2 條第 1 項允許的 `IFW_`、`IFC_`、`SRC_` 都還不能用工具取。`IFC_` 是整欄範圍（Interface R6:R310），也不適合逐名取，需要工具改成「取某名稱同列的用途／層」。請 chat 端排工具工作單。

## 四、v5.29 重取影響（共用工作單第 4 節）

### 4.1 執行方式
- 以去掉 `IFW_` 的草稿（21 名）執行 `import_tokenomics.py --tokenomics /home/claude/tk-master --names <草稿> --out /tmp/zhipu_v529_snapshot.json`：成功（v5.29、`a5061d9`，21 名，missing 0）；`--check` 通過。快照放暫存路徑，**未提交**。
- `IFW_` 8 名與表頭以 openpyxl 直接讀 v5.29 具名範圍替代。
- 本模型沒有快照 JSON，`--check` 無法直接對本模型現用值；改用本模型自帶的 `tools/check_tk_snapshot.py --tk-dir /home/claude/tk-master`（逐名比對 TK_Link 與 v5.29），並以 openpyxl 另寫一段逐格比對（含 7 列讀表）交叉確認，兩者結果一致。

### 4.2 現用值（v5.26）對 v5.29

| 項目 | 結果 |
|---|---|
| 具名範圍 74 名＋讀表 7 列＝81 項 | **75 項相同、6 項變動、0 項消失**；世代與成本情境表頭相同 |
| 變動 6 項（全部是 `IF_Alloc*`，全部無公式引用） | Q1 0.636 → 0.573（−9.9%）；Q1_R2 0.667 → 0.629（−5.7%）；Q2 0.664 → 0.603（−9.2%）；ServeGW 0.0519 → 0.0675（+30.1%）；Demand 4.19e9 → 6.02e9 M tok（+43.6%）；ImpliedNk 1.524 → 1.983（+30.1%） |
| `IF_AllocRDGW`（唯一被引用的 Alloc 名稱） | 0.0907 不變 |
| 讀表 7 列（Cap_In、Price_Frontier） | 不變 |
| **命題輸出** | **不變**（重取後重建只會改 TK_Link 的版本欄與 6 個死格） |

### 4.3 四層瀑布：目前所用層 vs `IFW_*_Prod`／`_Life`（基準成本情境）

| 名稱 | 本模型目前所用 | `_100` | `_Util` | `_Prod` | `_Life` |
|---|---|---|---|---|---|
| TokGW_Sol：Hopper（M tok/GW/年） | 2.149e10（100% 層） | 2.149e10 | 1.289e10（×0.60） | 1.569e10（×0.730） | ＝Prod |
| TokGW_Sol：VR200 | 2.373e11（100% 層） | 2.373e11 | 1.424e11（×0.60） | 1.916e11（×0.808） | ＝Prod |
| **Hopper ÷ VR200（VR 等值係數）** | **0.0906** | 0.0906 | 0.0906 | **0.0819（−9.6%）** | 0.0819 |
| TokGW_Luna：Hopper ÷ VR200 | 0.0958 | 0.0958 | 0.0958 | 0.0835（−12.9%） | 0.0835 |

讀法：
- `_Util` 層只是全世代乘同一個 60%，比值不變；`_Prod` 層（生產環境折減，經 Perf_Prod 鏈）對 Hopper 折得比 VR200 多（0.73 對 0.81），所以**舊晶片換算成 VR 等值時會更少**。注意 Prod 層是在 100% 上做生產折減、未再乘利用率，它與 Util 層是並列的兩條，不是逐層相乘。
- 對本模型的影響（估算，未重建）：VR 等值係數（H20、國產都是 Hopper × 比例，同比例變動）× 0.904 → 供給 VR 等值 GW（命題分母）× 0.904 → **每 VR 等值 GW 差額的絕對值 × 約 1.106**；main 版 2030 約 −2,275 → 約 −2,516 RMB 億（分支 r2 版 −2,394 → 約 −2,648）；覆蓋率（營收 ÷ 全成本）不受影響。
- token 換算推論 GW 若改用生產層：所需 GW × 約 1/0.73；但 2025、1H26 的 η 是以同一換算校準的，基準路徑（η 沿用 1H26）下會相互抵銷；在「η 收斂至 1」情境或分支 r2 版的 η 上調機制下不會完全抵銷。
- `IF_Cost{Pre,Cache,Dec}` 沒有對應的 `IFW_`，目前無法取生產層（見 3.1）。

## 五、CI 失敗診斷（只診斷、未修）

### 5.1 失敗的是什麼
- workflow：`.github/workflows/zhipu-parity.yml`。main 上 run 37782720760（`75327ef`）、run 37766713633（`148a38e`）、PR run 37766730222 都失敗；最後一次成功是分支 `claude/zhipu-v0.1` 的 run 37753142646（`a3fdb3c`）。
- 失敗 job：**`structure-recalc`**（結構性測試）；**`parity`** 只是彙總門檻，因 structure-recalc 失敗而失敗。scenarios 兩片、scenario-guard、engine-cache、governance、rebuild、tk-snapshot 全部通過——**數值一致性沒有問題**。
- 失敗的兩個測試（`zhipu/tests/parity/test_builder.py`；`gh run view --log-failed` 下載記錄被代理擋（403），改讀 check-run annotations 取得）：
  1. `test_src_ids_stable`（第 36 行）：`assert [yaml 的 SRC_ID] == [SRC_ZP_001 … SRC_ZP_610]` 失敗——「Left contains 444 more items, first extra item: 'SRC_ZP_611'」。
  2. `test_src_and_inputs_values_equal_yaml`（第 68 行）：讀 Excel 具名範圍 `SRC_ZP_611` 時 `KeyError: 'SRC_ZP_611'`——Excel 沒有這個名稱。

### 5.2 為什麼
| 檔案（main） | SRC_ZP 列數 |
|---|---|
| `zhipu/data/zhipu_src.yaml` | **1,054** |
| `zhipu/builder/id_registry.json` | 610 |
| `zhipu/model/20261008_Zhipu_v0.1.xlsx`（SRC_ZP 頁） | 610 |
| `zhipu/tests/parity/scenarios.yaml` 的 `src_rows` | 610 |

- 2026-10-08 06:28Z chat 端提交 `767f56a`「智譜：招股章程擷取 SRC_ZP_611–1054；Z5 工作單追加 P1–P5」只改了 yaml（資料），尚未重建 Excel。
- PR #24（`claude/zhipu-v0.1` → main）在 07:28Z 合併，帶進了 `767f56a`。
- 之後建置代理在**同一分支**再推了 7 個提交（`3928712` r2 P1：「SRC_ZP 寫入 444 列（611–1054）」、Inputs INP_137–140、Checks C94–C95、parity ＋2；以及 P2、P3–P4、V4–V7＋P5、V9、最後 `a3fdb3c`「全套 pytest 120 passed」）。這 7 個提交把 registry、Excel、期望值一起更新到 1,054 列，在分支上 CI 全綠；但 PR #24 已經合併關閉，**這 7 個提交從未進 main**（GitHub compare：分支領先 main 7 個提交）。
- 結果：main 上「資料 1,054 列、Excel 與登記檔 610 列」，測試正確地抓到不一致。
- 判定：**不是模型錯，也不是期望值過時寫錯；是 main 少了 7 個已通過 CI 的提交（半合併）。** 另外 main 上 coreweave 等其他資料夾的 PR 會因 workflow 只看 `zhipu/**` 而不受影響；但凡碰到 `zhipu/**` 的提交（包括本 PR）都會繼續紅燈。

### 5.3 建議修法（未執行）
1. **首選**：從 `claude/zhipu-v0.1`（`a3fdb3c`）開一個新 PR 到 main，把 7 個提交補合併。與 main 的差異只在 `zhipu/` 內，main 這段期間在 `zhipu/` 只新增了一份工作單檔（本審視工作單），預期無衝突。合併後 main 的智譜模型會變成 r2 P1–V9 版（η 2025 1.13／1H26 3.91；2030 每 VR 等值 GW 差額 −2,394 RMB 億、覆蓋率 0.71、累計外部資金需求 43.6 億）——**命題數字會變**，chat 端需知悉。
2. 次選（若 r2 P1 尚不想進 main）：在 main 上把 `zhipu_src.yaml` 回退到 610 列（還原 `767f56a` 的 yaml 部分），等 r2 一起合併。不建議放寬或刪測試。
3. 流程建議：PR 合併後若分支繼續推，應另開 PR；或在智譜 CLAUDE.md 加一句「PR 合併後的後續提交須另開 PR」。

### 5.4 本 PR 的 CI
本 PR 只新增一份報告檔（`zhipu/docs/…`），但路徑在 `zhipu/**` 內，會觸發 zhipu-parity，**預期同樣在 structure-recalc 失敗（原因同上，與本 PR 無關）**。

## 六、建議改動（交 chat 端開工作單）

1. **先修 CI**：合併 `claude/zhipu-v0.1` 的 7 個提交（5.3 第 1 項），之後再做任何 Tokenomics 改接，避免在半合併狀態上疊改。
2. **改用官方快照工具**：建立 `data/tokenomics_names.txt`（3.2 草稿）與 `data/tokenomics_snapshot_v5.29.json`，TK_Link 改由快照產生、刪掉 58 個死名稱，CI 加 `--check`；同時移除 7 列讀表（Revenue R78–R81 對照列暫時改「待 Tokenomics 提供」），並請 Tokenomics 提供 GLM 牌價與中國合格前緣的具名範圍。前提：工具先支援 `IFW_`（及 `SRC_`、`IFC_` 同列讀取）。
3. **VR 等值係數改用生產層並列**：Compute 第十節加一組以 `IFW_TokGW_Sol_Prod` 計算的 VR 等值係數與命題 1，與現行 100% 層並列，由 chat 端決定主值（命題 1 絕對值約 +10.6%）。

其餘：handoff 補寫「按 token 計價、採技術利用率 `IF_Util`、不用簽約率」與 Tokenomics 版本／雜湊（契約第 3、5 節）；η 對照 `L1_Gap*` 四項逐項拆解。

## 七、待 chat 端判斷

1. **VR 等值換算用哪一層**：100% 層（現行，比值 0.0906）或生產層（0.0819）？兩者都屬 `IFC_Use` 允許的「相對比較」；契約第 2 條第 2 項要求同時取四層。改用生產層會使命題 1 絕對值放大約 10.6%。
2. **η 與 `L1_Gap*` 的關係**：是否在 Compute 第七節加一段「η ÷ `L1_GapProduct`」對照（並逐項列每則提示 token 數、支出口徑、ISL、利用率），把 η 從單一黑箱倍數拆成「口徑差」與「智譜真實效率」兩部分？
3. **token 換算 GW 的層**：現行以 100% 層逐類型單位成本換算（`IFC_Use`「上限，不得作預測」），得到的是所需 GW 的下界；是否請 Tokenomics 補 `IF_Cost*` 的生產層，或在本模型以 `CTL_ProdDerate` 類的折減並列？
4. **閒置保留 10% 與 `IF_Util` 60% 是否重疊**：兩者相乘使實際負載約 54%；是否維持（沿用 OpenAI v0.6 INP_232）或改為只用其一？
5. **晶片缺口走 G2**：H20、昇騰的產出比／持有成本比／每卡功耗與中國 GPU-hour 租價，是否由 chat 端整理成 DB_Evidence 候選送 Tokenomics？送入前本模型維持 Inputs 現值。
6. **`TK_HdrGen`／`TK_HdrCost` 表頭複製**是否視為違反「`IF_Hdr*` 禁止連結」？若是，改接官方快照時一併改用工具的世代／情境拆欄結構。
7. **CI 修法**：採 5.3 第 1 項（補合併分支，命題數字會更新為 r2 版）或第 2 項（回退 yaml）？
8. **Coding Plan token 組成**是否改取 Tokenomics `IF_TaskTok*`（Coding agent 任務）取代 INP_122／123？

## 提交紀錄
- 內容提交：`dafad92`（只新增本報告）；SHA 補記提交隨後；未改 company.json（本模型無此檔）、builder、Excel、data。
