# MiniMax｜Tokenomics 連接審視（v5.29）

- 分支：`claude/minimax-tk-link-review`
- 底稿：company-models main `75327ef`；報告提交 SHA：`5c4074d`（初稿；本行更新於下一個提交）
- 報告路徑：`minimax/docs/reports/20261008_minimax_tokenomics_link_review.md`
- 依據：共用工作單 `docs/workorders/20261008_tokenomics_link_review.md`、本公司工作單 `minimax/docs/workorders/20261008_minimax_tokenomics_link_review.md`；Tokenomics 下游資料契約 v0.1；Tokenomics master `a5061d9`（`model/CURRENT`＝`20261008_Tokenomics_v5.29.xlsx`，前提成立）。
- 本輪**只審視、不改模型**：`build_xlsx.py`、`build_html.py`、`data/`、`dist/` 的 Excel 都沒有動。v5.29 快照只產生在暫存路徑（`/tmp/minimax_v529_snapshot.json`），未提交。

## 摘要（給 Andy）

1. **MiniMax 怎麼拿 Tokenomics 的數字**：沒有用官方取數工具，也沒有直接讀 Tokenomics。12 個數字是**手抄寫死**在 `build_xlsx.py` 的 `TK` 清單裡，來源是「OpenAI 模型 v0.6 的 TK_Link 頁」（二手），標註版本 v5.24（master `bdb0de7`）。沒有快照檔、沒有重現檢查，也沒有 CI。
2. **拿了什麼**：每 GW 產能（Luna、Sol；Hopper、GB200、VR200）5 格、基準利用率 60% 1 格、每 GW 持有成本 2 格、每百萬 token 全成本 4 格。真正進入公式的只有 7 格（產能 5、利用率 1、Hopper 持有成本 1——後者只作對照）；另 5 格只擺在 TK_Link 頁供參考。
3. **v5.29 重取的結果**：這 12 個數字在 v5.29 **全部沒變**（差異 0）。v5.29 改變的 `IF_Alloc*` 六格 MiniMax 沒有使用，所以**本模型所有輸出都不受 v5.29 影響**。要改的是「取法」，不是數字。
4. **口徑**：收入按 token 計價（企業端 API）＋消費端訂閱／廣告；算力成本以「每 MW 年合約價 $12M」換算。利用率只用 Tokenomics 的 60%，沒有另扣簽約率，**沒有重複扣減**。而且因為 η（有效產出係數）以 2026 上半年實際支出校準，Tokenomics 產能的絕對水準在 2027 年後會被抵銷，模型只用到「不同世代之間的產能比例」。
5. **最重要的發現（Tokenomics 端）**：v5.29 新增的 `L1_HoldEconMW_*`（每 MW 持有成本）單位標示 `$M/MW/年`，但數值是 0.0101（＝每 GW 的 $B ÷ 1,000，實際是 $B/MW），差 1,000 倍。契約第 3 節把它列為神雲／雲平台的成本底線，下游照單位直接用會出錯。列入「待 chat 端判斷」。

## 一、取數點清單（共用工作單第 1 節）

`link_audit.py` 的候選清單（minimax 節）：使用名稱 6 個（`IF_FullCost_Luna`、`IF_FullCost_Sol`、`IF_HoldEcon`、`IF_TokGW_Luna`、`IF_TokGW_Sol`、`IF_Util`）；違規引用 1 筆 `Inputs!`（build_xlsx.py）。人工確認後：

- **`Inputs!` 為誤判**：出現在 `build_xlsx.py` 的檢查 K08（「期末現金不得低於最低水位」），引用的是 **MiniMax 自己 Excel 的 Inputs 頁**（最低現金水位 I24），不是 Tokenomics 的 Inputs 頁。不違規。
- 工具看不到的真正問題：**數字寫死**。下表 12 列全部是手抄值（`build_xlsx.py` 第 110–125 行），不是從具名範圍讀出。

| 本模型的格 | Tokenomics 名稱（世代） | 合規判定 | 四層瀑布的哪一層 | 用途是否符合 `IFC_Use` | 所用版本 |
|---|---|---|---|---|---|
| TK_Link TK01 | `IF_TokGW_Luna`（Hopper） | 名稱合規；**寫死數字（違規）** | 100%（公式內再乘 TK06 → 等同 IF_Util 層） | 是。v5.29 此列標「上限，不得作預測」；本模型只用來把 token 換算成 MW（需求面推 MW），不推收入；且乘了利用率、再經 η 以實際支出校準 | v5.24（`bdb0de7`，經 OpenAI v0.6 TK_Link 轉抄） |
| TK_Link TK02 | `IF_TokGW_Luna`（GB200） | 同上 | 同上 | 同上 | 同上 |
| TK_Link TK03 | `IF_TokGW_Sol`（Hopper） | 同上 | 同上；另用於 VR 等值係數（C23，比例，100% 層） | 同上；VR 等值為「相對比較」，符合 | 同上 |
| TK_Link TK04 | `IF_TokGW_Sol`（GB200） | 同上 | 同上 | 同上 | 同上 |
| TK_Link TK05 | `IF_TokGW_Sol`（VR200） | 同上 | 100%（只用於 VR 等值係數的分母） | 是（相對比較） | 同上 |
| TK_Link TK06 | `IF_Util` | 同上 | IF_Util（60%） | 是（情境用；本模型未覆寫） | 同上 |
| TK_Link TK07 | `IF_HoldEcon`（Hopper） | 同上 | 不適用（成本） | 是（C29 只作對照：合約價 − 持有成本＝雲端毛利） | 同上 |
| TK_Link TK08 | `IF_HoldEcon`（GB200） | 同上 | 不適用 | 未進公式（只顯示） | 同上 |
| TK_Link TK09 | `IF_FullCost_Luna`（Hopper） | 同上 | IF_Util | 未進公式（C30 註解說「可比較」，但沒有公式連結） | 同上 |
| TK_Link TK10 | `IF_FullCost_Sol`（Hopper） | 同上 | IF_Util | 未進公式 | 同上 |
| TK_Link TK11 | `IF_FullCost_Luna`（GB200） | 同上 | IF_Util | 未進公式 | 同上 |
| TK_Link TK12 | `IF_FullCost_Sol`（GB200） | 同上 | IF_Util | 未進公式 | 同上 |

文字中寫死的 Tokenomics 數字（不進公式，但會過期）：
- Inputs I02 說明「Tokenomics Hopper 持有成本 10.1（TK07）」。
- Inputs I04 說明「Tokenomics Luna（$0.08）與 Sol（$1.6）」＝`IF_PriceRef_Luna` 0.0804、`IF_PriceRef_Sol` 1.607（v5.29 未變）。
- 版本字串寫死 3 處：`build_xlsx.py` 的 `TKSRC`、README 頁說明、`build_html.py` 副標「Tokenomics v5.24」；`minimax/README.md` 的 TK_Link 列。

**合計**：寫死數值 12 個（進公式 7 個、只顯示 5 個）＋說明文字內 3 個數字＋4 處版本字串。違規引用（其他前綴、直接引用工作表）0 筆。

**取數流程重建**（MiniMax 無工作單、無 CLAUDE.md；git 只有 `f373885`（隨 PR #26 一次併入）與 `148a38e` 兩筆）：MiniMax v0.1 於 2026-10-08 一次建成，建置者打開 `openai/` v0.6 的 TK_Link 頁（OpenAI 由 builder 從 Tokenomics v5.24 讀出），把需要的 12 格數值複製進 `build_xlsx.py`。之後沒有任何重新取數的機制。

## 二、口徑問題（共用工作單第 2 節）

1. **計價方式與利用率**：企業端（API）按 token 計價（token 量 × 實收單價）；消費端是訂閱、儲值、廣告，算力以支出換算，不走 token。對應契約第 3 節「按 token（模型商）」，用 Tokenomics 的技術利用率 `IF_Util` 60%（未覆寫），**沒有使用簽約率**，所以沒有重複扣減。補充：模型另有 η（2025＝3.42、2026 上半年＝0.80，2026 下半年起 × 1.3＝1.04），是把「Tokenomics 推算產能」對到「實際算力支出」的總校準係數；它吸收了實際利用率、提示長度、計價口徑等差異。由於 2027 年起的 API MW＝token ÷（產能 × 效率）÷ η，而 η 又是用 2026 年同一組產能校準，**Tokenomics 產能和利用率的絕對水準會互相抵銷**，只剩「Hopper 與 Blackwell 的產能比」隨機隊組合變化而起作用。
2. **神雲／雲平台的 50/50 平均**：不適用（MiniMax 是模型商）。MiniMax 的每 MW 營收是「實際營收 ÷ 推得的 MW」，沒有成本加成與市價平均。相關的只有成本端：每 MW 年合約價 I02＝$12M（Analogy，抄自 OpenAI 模型 INP_229），是單一數字、不隨情境切換，也沒有容量與單價綁在同一選擇器的問題。
3. **`IF_Alloc*` 六格**：**MiniMax 完全沒有使用**（Q1、Q1_R2、Q2、服務 GW、需求 D、隱含 N×k、研發 GW 都沒有）。MiniMax 的研發／服務分配是自己由財報推得：訓練 MW＝研發算力支出 ÷ 合約價（Compute C17–C19），訓練占總 MW 2025 約 79%、2026 約 66%（C22）。這等於已依 Decisions G16「其他實驗室自行覆寫」的精神處理。v5.29 這六格的變動（例如 Q1 0.636 → 0.573、隱含 N×k 1.524 → 1.983）**不影響任何 MiniMax 輸出**；可列為對照：MiniMax 2026 訓練占比 66% 高於 Tokenomics 的 OpenAI 代理 Q1 57%。
4. **公司特有變數是否回寫 Tokenomics**：沒有。合約價（I02）、層級組合（I04）、機隊組合（I05）、模型效率（I06）、推論效率改善（I29）、最低現金、折價、折現率等都在 MiniMax 自己的 Inputs 頁；程式沒有任何寫入 Tokenomics 的路徑。符合契約第 2 條第 3 項。

## 三、缺口與版本（共用工作單第 3 節）

### Tokenomics 缺的產業級數據（只列出，不自行建表；由 chat 端走 G2 進 DB_Evidence）

| 缺口 | 用途 | 目前以什麼代替 |
|---|---|---|
| 中國可用 GPU 世代（H800、H20、華為昇騰 910C 等）的每 GW 產能與持有成本 | MiniMax 機隊以 Hopper 級為主（2025 為 100%），推 MW | 以 Hopper H100 代替 H800／H20；「Blackwell 級」含國產 ASIC 等值（I05 說明） |
| 中國雲端 GPU 租價（每 GPU-hr 或每 MW 年） | 支出 → MW 的換算（I02），決定所有 MW 與每 MW 指標 | 沿用 OpenAI 模型的 $12M/MW/年（Analogy，區間 8–20） |
| 小型 MoE 層級（啟用參數約 10–25B）的每 GW 產能 | API token → MW | Luna 與 Sol 以 60／40 混合（I04） |
| 影片生成的每 GW 產出（每秒影片的算力） | 消費端（海螺）占消費營收約 40%，目前只能用支出換算 | 消費端算力成本 ÷ 營收比例（I13） |
| 中國 API 價格時間序列（DeepSeek、Qwen、MiniMax 牌價與實收） | 價格下降路徑（I08） | 放在 MiniMax 自己的 SRC_MM（S056–S058）與假設 I07／I08 |

### names.txt 草稿（`minimax/data/tokenomics_names.txt`；只列出，不執行、不提交）

```
# MiniMax Tokenomics 取數名稱草稿（v5.29）
# 產能：四層瀑布（契約第 2 條第 2 項）；IF_TokGW_* 只作 100% 層對照
IFW_TokGW_Luna_100
IFW_TokGW_Luna_Util
IFW_TokGW_Luna_Prod
IFW_TokGW_Luna_Life
IFW_TokGW_Sol_100
IFW_TokGW_Sol_Util
IFW_TokGW_Sol_Prod
IFW_TokGW_Sol_Life
IF_TokGW_Luna
IF_TokGW_Sol
IF_Util
# 成本（只作對照）
IF_HoldEcon
IF_FullCost_Luna
IF_FullCost_Sol
IF_FullCost_Luna_Prod
IF_FullCost_Sol_Prod
# 世代比與服務端落差分解（η 拆解用）
L1_TokMW_Gen_ratio_GB200
L1_TokMW_Gen_ratio_VR200
L1_GapPrompt
L1_GapSpendBasis
L1_GapISL
L1_GapUtil
L1_GapProduct
# 價格參考（I04 說明用）
IF_PriceRef_Luna
IF_PriceRef_Sol
# 用途欄（逐列對應 Interface R–U）
IFC_Use
IFC_Conf
```

注意：共用工具 `import_tokenomics.py` 目前**只接受 `IF_` 與 `L1_`**，`IFW_`、`IFC_` 會被拒絕（見「待 chat 端判斷」2）。`L1_HoldEconMW_*` 暫不列入（單位問題，見摘要第 5 點）。

## 四、v5.29 重取影響（共用工作單第 4 節）

做法：以草稿中工具可接受的 17 個名稱（IF_／L1_）執行
`python3 tools/tokenomics/import_tokenomics.py --tokenomics /home/claude/tk-master --names <草稿> --out /tmp/minimax_v529_snapshot.json`
→ 成功（`20261008_Tokenomics_v5.29.xlsx`，commit `a5061d9`，17 個名稱，缺漏 0，未重算＝快取值完整）。MiniMax 沒有既有快照檔，`--check` 無法對本模型使用（沒有可比對的快照），改為把 v5.29 快照值與 `build_xlsx.py` 的 12 個寫死值逐一比對；`IFW_`／`IFC_` 以 openpyxl 直接讀 v5.29 Interface 頁。

### 現用 12 格：v5.24（經 OpenAI 轉抄）→ v5.29

| 格 | 名稱（世代，基準） | 現用值 | v5.29 | 變動 |
|---|---|---|---|---|
| TK01 | IF_TokGW_Luna（Hopper） | 74,067,251,305 M tok/GW/年 | 同 | 0 |
| TK02 | IF_TokGW_Luna（GB200） | 362,306,894,086 | 同 | 0 |
| TK03 | IF_TokGW_Sol（Hopper） | 21,490,429,685 | 同 | 0 |
| TK04 | IF_TokGW_Sol（GB200） | 104,125,401,081 | 同 | 0 |
| TK05 | IF_TokGW_Sol（VR200） | 237,287,338,735 | 同 | 0 |
| TK06 | IF_Util | 0.60 | 同 | 0 |
| TK07 | IF_HoldEcon（Hopper） | 10.0958 $B/GW/年 | 同 | 0 |
| TK08 | IF_HoldEcon（GB200） | 9.1710 | 同 | 0 |
| TK09 | IF_FullCost_Luna（Hopper） | 0.19229 $/M | 同 | 0 |
| TK10 | IF_FullCost_Sol（Hopper） | 0.76714 | 同 | 0 |
| TK11 | IF_FullCost_Luna（GB200） | 0.03722 | 同 | 0 |
| TK12 | IF_FullCost_Sol（GB200） | 0.13890 | 同 | 0 |

結論：**數值變動 0**；版本標記應由 v5.24 改為 v5.29（`a5061d9`）。`IF_Alloc*` 未使用，無影響。

### 四層瀑布：本模型目前所用層 vs `IFW_*_Prod`／`_Life`

本模型用「100% 產能 × IF_Util」，數值上等於 `IFW_TokGW_*_Util`（完全相同）。若改用 `_Prod` 層（× CTL_ProdDerate；v5.29 中 `_Life`＝`_Prod`，因 token 產能無壽命折減）：

| 世代（基準） | Util 層 | Prod 層 | Prod ÷ Util |
|---|---|---|---|
| Luna Hopper | 44,440,350,783 | 50,766,524,970 | 1.142 |
| Luna GB200 | 217,384,136,451 | 283,232,558,337 | 1.303 |
| Sol Hopper | 12,894,257,811 | 15,694,514,740 | 1.217 |
| Sol GB200 | 62,475,240,648 | 83,122,186,024 | 1.330 |

因為 η 以 2026 年校準，絕對水準抵銷，只有 Blackwell 對 Hopper 的相對倍數改變會傳到輸出。試算（不改模型，只算比例）：改用 Prod 層後，API 推論 MW 2027 −3.2%、2028 −4.7%、2029 −5.4%、2030 −5.7%（2025、2026 不變，兩年由實際支出決定）。API 推論 MW 只占總 MW 約 12–22%，對總 MW 與總成本的影響約 −0.5% 以內；覆蓋率（每 MW 營收 ÷ 每 MW 全成本）幾乎不變。

`IFC_Use`（v5.29 Interface S 欄）：`IF_TokGW_*` 與 `IFW_*_100` 標「情境、相對比較、反轉門檻；上限，不得作預測」；`_Util`、`_Prod`、`_Life` 與 `IF_FullCost_*`、`IF_HoldEcon` 標「情境、相對比較、反轉門檻」；信心度（`IFC_Conf`）皆為 B。本模型沒有把上限列當收入預測使用。

## 五、MiniMax 補齊工作流程的建議

目前 MiniMax 是各公司中唯一沒有 CLAUDE.md、工作單紀錄、CHANGELOG、CI 的資料夾（本輪工作單目錄是 chat 端建立的）。建議（由 chat 端另開工作單決定）：

1. **`minimax/CLAUDE.md`**：沿用 `openai/CLAUDE.md` 的守則，寫明 Excel 是唯一計算引擎（已是現況）、判斷類／工程類分工、Tokenomics 只經官方快照取數、每輪報告格式。
2. **`minimax/docs/workorders/`**：已存在（本單）；之後每一版（v0.2 起）先有工作單再建置。
3. **`minimax/CHANGELOG.md`**：補記 v0.1（2026-10-08，含 Tokenomics v5.24 經 OpenAI 轉抄的事實）。
4. **進度檔** `minimax/docs/reports/YYYYMMDD_minimax_進度.md`：比照 `openai/docs/reports/20261007_openai_進度.md`。
5. **Tokenomics 取數**：`data/tokenomics_names.txt`＋`data/tokenomics_snapshot_v5.29.json`；`build_xlsx.py` 的 TK 清單改為讀快照（數值不變）；TK_Link 頁加檔名、commit、`IFC_Use`／`IFC_Conf` 欄。
6. **CI**：新增 `.github/workflows/minimax-*.yml`（只在 `minimax/` 變動時執行）：建置、LibreOffice 重算、Checks 全 0、錯誤值 0、快照 `--check`。

## 六、建議改動（優先順序）

1. **改用官方快照取數**（工程類，數值不變）：建 names.txt 與 v5.29 快照，`build_xlsx.py` 不再寫死 12 個數字；版本由「v5.24 經 OpenAI 轉抄」改為「v5.29，`a5061d9`」。
2. **改取四層瀑布並決定採哪一層**（判斷類）：把「`IF_TokGW` × `IF_Util`」換成 `IFW_TokGW_*_Util`（數值相同、零影響），或改 `_Prod`（2027–30 API MW −3% 至 −6%）。TK_Link 頁同時顯示四層與用途欄。
3. **把 η 與合約價對到 Tokenomics 的落差分解**（判斷類）：η 目前是一個總係數（2026＝0.80），應拆成 `L1_GapPrompt`、`L1_GapSpendBasis`、`L1_GapISL`、`L1_GapUtil` 對照；並以 `L1_GapSpendBasis`（雲端計價 ÷ 持有成本＝2.0）檢查 I02 合約價 $12M（只比 Hopper 持有成本 10.1 高 19%，低於 2 倍）。

## 待 chat 端判斷

1. **Tokenomics `L1_HoldEconMW_*` 單位**：值為 0.0101（Hopper）等，公式＝`L1_HoldEconGW_* / 1000`，每 GW 單位為 $B，因此實際是 $B/MW/年，但 G 欄標示 `$M/MW/年`（v5.29 報告第 55 行同樣寫 $M）。正確的 $M/MW/年應為 10.1。這是 Tokenomics 端的單位不一致，依規則只回報，請 chat 端在 Tokenomics 開工作單（改標示為 $B/MW，或公式改為 × 1,000 ÷ 1,000＝與每 GW 同數值的 $M/MW）。
2. **共用工具不收 `IFW_`／`IFC_`**：`import_tokenomics.py` 只接受 `IF_`、`L1_`，契約第 2 條第 2 項卻要求取四層瀑布與用途欄。需決定共用工具是否擴充（影響八家）。
3. **層別選擇**：MiniMax 用 Util 層還是 Prod 層（見第四節）；或兩者都放、主值用哪一個。
4. **η 的合理性**：MiniMax 2026 η＝0.80（Tokenomics 產能只高估 1.25 倍），而 Tokenomics 對 OpenAI 的服務端落差乘積 `L1_GapProduct`＝13.4（約 16.5 倍觀測）。兩者差距很大，可能來自 MiniMax 模型較小（啟用 10B／23B）、合約價口徑（I02）、或 2026 上半年支出的推得方式（銷售成本 × 90% − 消費端）。是否需要逐項拆解，或接受 η 為總校準係數。
5. **合約價 I02＝$12M/MW/年**：目前抄自 OpenAI 模型；中國雲租價與 H800／H20 的實際持有成本 Tokenomics 都沒有。是否走 G2 補中國租價證據，或維持 Analogy 區間 8–20。
6. **中國晶片世代**：是否請 Tokenomics 新增 H800／H20／昇騰等世代，或維持以 H100 代理並在 MiniMax 標示。
7. **流程補齊**（第五節 1–6）是否一併開成 MiniMax v0.2 的工作單。
