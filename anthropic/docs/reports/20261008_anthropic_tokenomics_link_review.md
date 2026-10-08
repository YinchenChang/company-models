# Anthropic｜Tokenomics 連接審視（v5.29；只審視回報，不改模型）

- 分支：`claude/anthropic-tk-link-review`（base `main` @ `75327ef`）
- 最新提交 SHA：見文末「提交紀錄」
- 報告路徑：`anthropic/docs/reports/20261008_anthropic_tokenomics_link_review.md`
- 依據：共用工作單 `docs/workorders/20261008_tokenomics_link_review.md`、本公司工作單 `anthropic/docs/workorders/20261008_anthropic_tokenomics_link_review.md`；Tokenomics 下游契約 v0.1；Tokenomics v5.29 報告與 CHANGELOG。
- 對照的 Tokenomics：唯讀 checkout master `a5061d9`，`model/CURRENT`＝`20261008_Tokenomics_v5.29.xlsx`（前提成立，`IFW_`、`IFC_` 名稱存在）。
- **本輪沒有改動模型**（company.json 不適用；builder、Excel、data 皆未動）。v5.29 快照只產生在暫存路徑 `/tmp/anthropic_v529_snapshot.json`，未提交；名稱清單草稿附在第三節，未寫入 `data/`。

## 一、結論（先讀這段）

1. **取數方式**：Anthropic 模型沒有用共用的快照工具，而是 builder（`builder/tk_link.py`）建置時直接打開 Tokenomics 的 xlsx 讀名稱，把值寫進 Excel 的 `TK_Link` 頁。釘住的版本是 **v5.26**（`20261007_Tokenomics_v5.26.xlsx`，讀取時 master `70d859e`）。共用審視工具報「無、未釘版本」是因為它不讀 xlsx，實際上有釘版本，只是記在 Excel 裡。
2. **TK_Link 共 89 列**（68 個具名範圍＋NonNV 表 18 格＋PUE 3 格），**真正被公式用到的只有 13 個具名範圍＋NonNV 18 格＋PUE 3 格**；其餘 55 個名稱（OpenAI 代理單價、訓練成本、理論營收、`IF_Alloc*` 五格、`SRC_DEM_*` 等）是從 OpenAI v0.6 架構帶過來、沒有任何公式引用的「擺著」的值。
3. **違規引用（人工確認後）**：實質違規 **2 項**——讀 Tokenomics 的 `NonNV` 頁（非具名，18 格）與 Tokenomics `Inputs!` 頁的 PUE（非具名，3 格）。另有 **1 項形式違規**：`IF_HdrTask`（任務名稱表頭）被當成查表鍵。工具報的 6 筆中，4 筆是同一個 `IF_HdrTask` 用法，2 筆 `Inputs!` 是本模型自己的 Inputs 頁（誤判）。工具沒抓到的 NonNV 與 PUE 讀表是人工補出。PUE 有合規替代：Tokenomics 已有 `IF_FacilityGW`（每 IT GW 的設施電力，基準 1.2，與現用值相同）。
4. **v5.29 對本模型的影響：零**。v5.26 → v5.29 之間，本模型讀的 89 列中只有 `IF_Alloc*` 六格變動，而這六格**沒有任何一格**被本模型公式引用（唯一引用的 `IF_AllocRDGW` 不變）。以 engine 把六格換成 v5.29 值重算，所有關鍵輸出（每 VR 等值 GW 差額、覆蓋率、營收、累計外部資金需求、現金）逐位不變。
5. **真正值得處理的是四層瀑布**：本模型的每 GW 產能＝`IF_TokGW`（100% 層）× `IF_Util`（60%），等於瀑布的 `_Util` 層。改用 `_Prod` 層時，因為 η 以 2025 實際支出校準，2025 不變，但各世代的相對產能改變（Hopper 的生產折減比新世代深），**2030 有效推論 GW 4.12 → 2.93（−29%）、研發 GW（殘差）5.22 → 6.42**；每 VR 等值 GW 差額與資金結論幾乎不變（2030 ＋11.20 → ＋11.21）。

## 二、取數點清單（共用工作單第 1 節）

所用 Tokenomics 版本：`20261007_Tokenomics_v5.26.xlsx`，讀取時 master `70d859e`（handoff 註記 CURRENT 與 `3dd1216` 相同）。版本與雜湊寫在 `TK_Link` 第 3–5 列與每列 G、H 欄。

### 2.1 被公式引用的取數點

| 本模型的格或參數 | Tokenomics 名稱 | 合規判定 | 瀑布層 | 用途是否符合 `IFC_Use` | 版本 |
|---|---|---|---|---|---|
| Compute C06–C17（D13:D24）每 GW 總產出（Hopper／GB200／GB300／VR200 × 三層級）→ Compute 八節每 GW 年產能（×IF_Util）→ 推論 GW；Compute C21／C25…VR 等值係數（Sol 產能比） | `IF_TokGW_Astra`、`IF_TokGW_Sol`、`IF_TokGW_Luna` | 合規（IF_） | 讀 100%，公式乘 `IF_Util` → 實際用 **IF_Util 層** | 符合：`IFC_Use`「情境、相對比較、反轉門檻；上限，不得作預測」。本模型只用來把 token 換算成 GW（並以 η 校準 2025），不當收入預測 | v5.26 |
| Compute 八節（第 120–123、130–133 列）、Cost K111／K112 | `IF_Util`（0.6） | 合規 | IF_Util | 符合（情境、相對比較） | v5.26 |
| Cost K03–K29 每 GW 年持有成本（各族 × 三情境）→ 供應商持有成本、雲端毛利、經濟口徑 | `IF_HoldEcon` | 合規 | 不適用（成本） | 符合；只作參考與經濟口徑，算力現金成本＝合約實付 | v5.26 |
| Compute C76 每 GW 資本支出（VR200 基準）→ 自建 GW | `IF_CapexTotal` | 合規 | 不適用 | 符合 | v5.26 |
| Cost K111 TK 參考全成本（下游預設） | `IF_FullCostDefault_Sol` | 合規 | IF_Util（`IFC_Layer`＝IF_Util；公式同乘 IF_Util，一致） | 符合；只作參考列 | v5.26 |
| Cost K112 TK 參考全成本（並列） | `IF_FullCost_Sol` | 合規 | IF_Util | 符合；只作參考列 | v5.26 |
| Compute C148 研發 GW 對照 | `IF_AllocRDGW` | 合規 | 不適用 | 符合；只並列對照，不驅動 | v5.26 |
| Demand D23–D25 每任務 token（對話、程式代理、其他代理）；Checks C51–C52 | `IF_TaskTokFresh`、`IF_TaskTokCached`、`IF_TaskTokDec` | 合規 | 不適用 | 符合 | v5.26 |
| 同上（SUMIFS 的查表鍵）；Inputs INP_102–104 存任務名稱 | `IF_HdrTask` | **形式違規**（契約第 2 條第 1 項禁止 `IF_Hdr*`）；實質是表頭文字，`IFC_Conf` 標「（表頭）」 | 不適用 | — | v5.26 |
| 所有 SUMIFS 的世代、成本情境鍵（`TK_HdrGen`、`TK_HdrCost`） | Interface 第 4、5 列表頭（`IF_HdrGen`、`IF_HdrCost` 的位置，依儲存格位置讀，不依名稱） | 同上屬表頭；共用快照工具本身也用這兩個表頭拆世代，建議視為允許 | 不適用 | — | v5.26 |
| Compute C34–C53 各族產出比、持有比；Cost 持有成本；敏感度 A／B／C；Checks C17 | Tokenomics `NonNV` 頁 B5:G7（TPU v7、Trainium3、MI455X 的產出比與持有比低／基準／高；**非具名**，以列標籤＋欄標題讀表） | **實質違規**（工作表直接引用） | 不適用（比例） | Tokenomics 無對應 IF_／L1_ 名稱，無 `IFC_Use` | v5.26 |
| Compute C60 PUE → 揭露 GW 換 IT GW；Checks C53 | Tokenomics `Inputs!` 第 9 列 PUE（D／E／F；**非具名**） | **實質違規**；合規替代 `IF_FacilityGW`（1.1／1.2／1.3，與現用值相同） | 不適用 | — | v5.26 |

### 2.2 在 TK_Link 但沒有任何公式引用的名稱（55 個）

`IF_FullCost_Astra／Luna`、`IF_FullCostDefault_Astra／Luna`、`IF_Price{Fresh,Cached,Think,Out,Ref}_*`（15）、`IF_FrontRef_*`（3）、`IF_ProgGWyr_*`（3）、`IF_RDMult`、`IF_TrainGenDefault`、`IF_TrainCost_*`（3）、`IF_RevGW_*`（3）、`L1_Ans3`（含低高 3）、`IF_AllocQ1`、`IF_AllocQ1_R2`、`IF_AllocQ2`、`IF_AllocServeGW`、`IF_AllocDemand`、`IF_AllocImpliedNk`、`IF_TaskLen`、`SRC_DEM_004`–`013`（含 013 低高，12）。

- 這些都是 OpenAI v0.6 的 63 名原樣帶過來（`tk_link.py` 的 PATTERNS／BLOCK6／EXTRA），Anthropic 的公式不讀。其中 `IF_Price*`、`IF_FrontRef*`、`IF_RevGW*` 是「OpenAI 有效單價」與「理想上限」，`SRC_DEM_*` 是 OpenAI 的需求與支出紀錄（v5.29 全部仍為 Active，契約允許，但共用快照工具不收 `SRC_`）。
- 它們不影響任何數字，但會讓讀者以為本模型用了 OpenAI 的價格或配置；也讓「v5.29 變了 6 格」看起來像會影響本模型。建議下一版移除（見第六節建議 2）。

### 2.3 審視工具候選清單的逐筆確認

| 工具報告 | 人工確認 |
|---|---|
| 使用名稱 52 個，含 `IF_*_Prod`／`_Life` | 誤判：工具把程式中的前綴樣式（如 `IF_RevGW_` ＋層級）展開成所有同前綴名稱；本模型沒有讀任何 `_Prod`／`_Life` 名稱 |
| 不存在於現行：`IF_Alloc`、`IF_Price`、`IF_TaskTok` | 誤判：是程式裡的正規式片段；實際讀的名稱在 v5.29 全部存在 |
| 違規 `IF_HdrTask` × 4 檔 | 同一個用法（builder、tk_link、測試情境、報表工具）；形式違規 1 項 |
| 違規 `Inputs!`（builder/a3.py、build.py） | 誤判：是本模型自己的 Inputs 頁 |
| （工具未抓到） | Tokenomics `NonNV` 頁讀表、Tokenomics `Inputs` 頁 PUE 讀表：實質違規 2 項（寫法是 `("Inputs", "PUE")`、`NONNV_SHEET`，不含「頁名!」字樣） |
| 取數方式「無」、釘住「—」 | 實際釘在 v5.26（記在 Excel `TK_Link`）；工具不讀 xlsx |

## 三、口徑問題（共用工作單第 2 節）

**問 1：收入怎麼計價？利用率與簽約率各用哪一個？有沒有在收入端重複扣減？**
- 按 token 計價（API：Anthropic 自己的牌價 × token 數 × 快取與折扣係數）＋按席位訂閱（Pro／Max／Team／Enterprise 月費）。屬契約第 3 節「按 token（模型商）」那一列，**只用技術利用率 `IF_Util`，沒有簽約率**。
- `IF_Util` 只出現在算力端（token → GW 換算），**收入端不扣利用率**：營收＝需求 × 價格；唯一連到算力的是容量上限係數（`CMP_CapFactor`），2025–2026 固定 1，2027–2030 計算值也都是 1（可用 GW 大於需求）。**收入端沒有重複扣減。**
- 算力端有三個折減疊在一起：`IF_Util` 60%、閒置保留 10%（INP_115）、訓練最低占比 30%（INP_116，只用於上限），再除以 η。但 η 是用 2025 實際推論支出校準的，`IF_Util` 的水準會被 η 抵銷：engine 測試把 `IF_TokGW` 換成 `IFW_*_Util` 並把 `IF_Util` 設 1，所有輸出逐位相同。所以 60% 這個數字對結論沒有影響，影響的是「各世代之間的相對產能」。
- handoff（`docs/handoff/Anthropic_handoff.md`）**沒有寫明採哪一種**，契約第 3 節要求寫明，建議補一句（第六節建議 5）。

**問 2：神雲／雲平台的每 MW 收入 50/50 平均？**
- 不適用（模型商）。附記：本模型的算力成本＝逐合約實付，每 GW 年合約價用 Inputs 的 12 $B/GW/年（INP_114，Analogy，同 OpenAI v0.6）× Tokenomics NonNV 持有比；Tokenomics `IF_HoldEcon` 只作供應商成本參考與雲端毛利推算，沒有「成本加成與市場價平均」。Tokenomics 成本情境（Compute D12「基準」）與合約價（INP_114）是兩個分開的輸入，沒有綁在同一個選擇器上。

**問 3：`IF_Alloc*` 六格的取用處；Anthropic 直接用 OpenAI 代理值是否合理；依 G16 是否應自行覆寫？**
- 取用處：六格（Q1、Q1_R2、Q2、服務 GW、需求 D、隱含 N×k）都在 TK_Link，但**沒有任何公式引用**。唯一被引用的 Block 6 名稱是 `IF_AllocRDGW`（Compute C148，只並列對照），v5.29 不變。
- **受影響的輸出：無**。以 engine 把六格換成 v5.29 值重算，每 VR 等值 GW 差額（2027 −5.25、2030 ＋11.20）、2030 覆蓋率 1.78、2030 營收淨額 196.8、累計外部資金需求 0、現金谷底 80.2、2030 年底現金 218.4，全部不變。
- 是否合理：本模型的研發 GW 是**自己的殘差**（供給 ×（1 − 閒置）− 有效推論 GW，規格 D12），沒有套用 OpenAI 的研發／服務占比，**實質上已符合 G16「其他實驗室須自行覆寫」的精神**。不過 Tokenomics 的 `IF_AllocRDGW`（0.091 GW·年）是以 OpenAI 為代表性實驗室的「單一前沿實驗室物理下限」，標籤目前寫「Tokenomics 計畫當量」，沒有說明是 OpenAI 代理；它和本模型 2025 研發 GW 0.44 差約 4.8 倍，讀者容易誤讀。建議改標籤，並把其餘五格移出 TK_Link。
- 若 chat 端希望研發 GW 有外部錨點，正確做法是請 Tokenomics 以 Anthropic 的需求參數另跑一組 Alloc_In（G16 的「自行覆寫」），而不是在本模型重算；列入待判斷。

**問 4：公司特有變數是否只在本模型覆寫、沒有回寫 Tokenomics？**
- 是。本模型的公司特有變數：每 GW 年合約價（INP_114）、合約爬坡與「最高可達」計入比例、閒置與訓練最低占比、η 路徑、每任務 token 相對 Tokenomics 任務的倍數（INP_105–107）、雲端平台抽成、最低現金月數與融資順序（本模型沒有 WACC 或價值捕獲率參數）。全部在 `Inputs`／`SRC_ANT`。
- `tk_link.py` 只以 openpyxl 讀取 Tokenomics、沒有任何存檔動作；本模型沒有回寫 Tokenomics 的路徑。

## 四、缺口與名稱清單草稿（共用工作單第 3 節）

### 4.1 Tokenomics 缺的產業級數據（不自行建表；交 chat 端走 G2 進 DB_Evidence）

| 名稱 | 用途 | 目前以什麼代替 |
|---|---|---|
| 非 NVIDIA 加速器產出比／持有比的具名範圍（TPU v7、Trainium3、MI455X；低／基準／高） | Compute 各族產能、持有成本、VR 等值、敏感度 | 讀 Tokenomics `NonNV` 頁（非具名）。數據 Tokenomics 已有，只缺 `L1_`／`IF_` 名稱 |
| 舊世代非 NVIDIA 比例（Google TPU v6e、AWS Trainium2，相對 Hopper） | 既有雲端（2025 前簽）的組合產能與合約價 | 以新世代比例（TPU v7、Trainium3 相對 VR200）套在 Hopper 上代理（Compute C34–C37、C42–C45） |
| 雲端算力租用價（每 GW 年，NVIDIA 基準；隨需／預留／長約分層） | 金額型合約換算 GW、既有雲端加權合約價 | Inputs INP_114＝12 $B/GW/年（Analogy，同 OpenAI v0.6；8–20）。Tokenomics 有成本底線 `L1_HoldEconMW_*` 與 SRC_Price 的 GPU-hr 價，但沒有每 GW 年租用價的具名輸出 |
| Anthropic 自己的研發／服務配置（G16 覆寫） | 研發 GW 外部錨點（目前只有殘差） | 殘差法；`IF_AllocRDGW`（OpenAI 代理）並列 |

PUE 不是缺口：`IF_FacilityGW`（Interface 第 8 列，「每 IT GW 對應設施電力 GW」，1.1／1.2／1.3）就是 PUE，可直接替換。

### 4.2 名稱清單草稿（擬放 `anthropic/data/tokenomics_names.txt`；未執行、未寫入）

```text
# Anthropic：Tokenomics 取數清單（草稿 2026-10-08；只列 IF_／IFW_／IFC_／L1_）
# 產能：四層瀑布（契約第 2 條第 2 項；取代 IF_TokGW_* ＋ IF_Util 的組合）
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
IF_Util
# 成本
IF_HoldEcon
IF_CapexTotal
IF_FacilityGW          # 取代 Tokenomics Inputs!PUE 讀表
IF_FullCost_Sol        # TK 參考列
IF_FullCostDefault_Sol # TK 參考列
# 需求：任務別每次嘗試 token（表頭 IF_HdrTask 需工具以表頭方式處理，見待判斷 3）
IF_TaskTokFresh
IF_TaskTokCached
IF_TaskTokDec
IF_TaskLen
# 研發對照（OpenAI 代理；只並列）
IF_AllocRDGW
# 用途欄（每個取數點附 IFC_Use／IFC_Conf／IFC_Layer／IFC_Updated）
IFC_Use
IFC_Conf
IFC_Layer
IFC_Updated
# 規模校準對照（可選）
L1_ScaleRD optional
L1_ScaleServe optional
# 尚無具名範圍（缺口，待 Tokenomics 提供）：NonNV 產出比／持有比（TPU v7、Trainium3、MI455X）
```

- **共用工具目前跑不了這份清單**：`import_tokenomics.py` 的名稱規則只收 `IF_`（不含 `IF_Hdr*`）與 `L1_`，`IFW_TokGW_Sol_Prod` 實測回「名稱不合規則」；`IFC_*` 是整欄（305 列）的範圍，工具也沒有「依取數點列帶出用途欄」的功能；任務別名稱需要 `IF_HdrTask` 當表頭才能拆。這是共用工具的缺口，非本模型能改。
- 因此第五節的快照只用工具可執行的 `IF_／L1_` 部分（55 名，即 TK_Link 的全部具名 IF_／L1_ 名稱），`IFW_`、`IFC_` 與 NonNV、PUE 以 openpyxl 直接讀 v5.29。

## 五、v5.29 重取影響（共用工作單第 4 節）

### 5.1 快照與比較方法
- `python3 tools/tokenomics/import_tokenomics.py --tokenomics /home/claude/tk-master --names <草稿 IF_／L1_ 55 名> --out /tmp/anthropic_v529_snapshot.json` → 成功（v5.29，commit `a5061d9`，55 名、missing 0）；`--check` 通過。
- 與本模型 `TK_Link`（v5.26）逐值比較（相對誤差 > 1e-9 才算變動）；另以本模型自己的 `tools/check_tk_snapshot.py --tk-dir /home/claude/tk-master` 交叉核對：**OK 83、DIFF 6**（89 列；NonNV 18 格、PUE 3 格、`SRC_DEM_*` 12 格皆不變）。

### 5.2 變動的名稱與幅度（全部 6 個，皆為 `IF_Alloc*`，皆未被本模型公式引用）

| 名稱 | v5.26（本模型現用） | v5.29 | 變動 |
|---|---|---|---|
| `IF_AllocQ1` 研發算力占比 | 0.6361 | 0.5733 | −9.9% |
| `IF_AllocQ1_R2` | 0.6672 | 0.6294 | −5.7% |
| `IF_AllocQ2` 研發算力成本占比 | 0.6636 | 0.6025 | −9.2% |
| `IF_AllocServeGW` 服務 GW | 0.05188 | 0.06751 | +30.1% |
| `IF_AllocDemand` 需求 D（M tok/年） | 4,190,200,000 | 6,015,200,000 | +43.6% |
| `IF_AllocImpliedNk` 隱含 N × k | 1.524 | 1.983 | +30.1% |
| `IF_AllocRDGW`（本模型唯一引用的 Alloc 名稱） | 0.09068 | 0.09068 | 不變 |

（v5.29 報告第一節寫 `IF_AllocImpliedNk` 1.524 → 1.322，那是第 1–2 輪；r2 把支出比改回 `SRC_DEM_004` 後，合併檔實際值是 1.983，以上表為準。）

其餘 49 個 IF_／L1_ 名稱、NonNV、PUE、`SRC_DEM_*`：**全部不變**。

### 5.3 `IFW_` 各層對本模型現用層（IF_Util 層）的差異

本模型現用＝`IF_TokGW`（100%）× `IF_Util` 0.6＝`IFW_TokGW_*_Util`（實測逐位相同）。v5.29 基準成本情境下，`_Prod` 層 ÷ `_Util` 層：

| 層級 | Hopper | GB200 | GB300 | VR200 |
|---|---|---|---|---|
| 頂層 Astra（Opus） | 0.81 | 1.22 | 1.16 | 1.28 |
| 中層 Sol（Sonnet） | 1.22 | 1.33 | 1.33 | 1.35 |
| 低層 Luna（Haiku） | 1.14 | 1.30 | 1.30 | 1.31 |

- `_Life` 層：產能（TokGW）的 `_Life`＝`_Prod` × 1（Tokenomics 附註「無 H 節對應」），與 `_Prod` 相同。
- 注意：產能瀑布的 `_Prod` 層（＝100% × 生產折減，不含 `IF_Util`）**大於** `_Util` 層，不是逐層遞減的瀑布；理論營收瀑布中 `IFW_RevGW_Sol_Life`（VR200 基準 228.8）又等於 `_Util` 層、高於 `_Prod`（184.8）。這是 Tokenomics 端的定義問題，列入待判斷，請 chat 端向 Tokenomics 確認各層的意義再決定下游取哪一層。
- engine 試算（暫存，不提交；把三個 `TK_IF_TokGW_*` 換成對應層、`TK_IF_Util` 設 1）：

| 輸出 | 現用（Util 層） | 改 `_Util` 層 | 改 `_Prod` 層 |
|---|---|---|---|
| η（2025 校準） | 0.888 | 0.888 | 0.971 |
| 有效推論 GW 2026／2028／2030 | 0.72／1.87／4.12 | 相同 | 0.52／1.34／2.93 |
| 研發 GW（殘差）2030 | 5.22 | 相同 | 6.42 |
| 每 VR 等值 GW 差額 2027／2030 | −5.25／＋11.20 | 相同 | −5.26／＋11.21 |
| 2030 覆蓋率、營收淨額、累計外部資金需求、現金谷底 | 1.78、196.8、0、80.2 | 相同 | 相同 |

讀法：改用 `_Prod` 層不影響「錢夠不夠」的結論（供給是合約定的，成本是合約實付），但會改變「推論用了多少、研發用了多少」的拆分：新世代相對 Hopper 的生產效率較高，同樣的 token 需要的 GW 少約三成，殘差歸到研發。

## 六、建議改動（交 chat 端另開工作單）

1. **改用 `IFW_TokGW_*` 四層瀑布**，基準取 `_Util` 層（數字不變），`_Prod`／`_Life` 列為敏感度並在 Compute 顯示三層（推論／研發 GW 拆分差約三成）；同時把 `IFC_Use`、`IFC_Layer` 帶進 TK_Link 每列。需先擴充共用快照工具（接受 `IFW_`／`IFC_`、以表頭處理 `IF_HdrTask`）。
2. **改用官方快照工具並清理 TK_Link**：以 `data/tokenomics_names.txt`（第四節草稿）產生 `data/tokenomics_snapshot_v5.29.json`，TK_Link 只保留被引用的名稱；移除 55 個未引用名稱（OpenAI 有效單價、理論營收上限、`IF_Alloc*` 五格、`SRC_DEM_*` 等）。版本欄改記 `model/CURRENT` 檔名與 master 合併雜湊（目前記的是讀取時的 master HEAD）。
3. **消除兩項實質違規**：PUE 改讀 `IF_FacilityGW`（值相同 1.1／1.2／1.3）；NonNV 比例請 Tokenomics 提供具名範圍（`L1_` 或 `IF_`），提供前維持讀表並在報告持續列為缺口。
4. **`IF_AllocRDGW` 對照列改標籤**為「OpenAI 代理（G16）；單一前沿實驗室物理下限」，避免與本模型殘差研發 GW 直接比較。
5. **handoff 補一句計價口徑**：「按 token＋席位計價；只用技術利用率 `IF_Util`（算力端），不用簽約率；收入端不扣利用率」（契約第 3 節）。

## 七、待 chat 端判斷

1. **四層瀑布取哪一層當基準**：本模型的 η 以 2025 實際支出校準，`IF_Util` 水準被抵銷，但 `_Prod` 層改變世代間相對產能（2030 推論 GW −29%）。且產能瀑布 `_Prod` > `_Util`、營收瀑布 `_Life`＝`_Util`，層的意義需先向 Tokenomics 確認。建議基準 `_Util`、`_Prod` 作敏感度，請 chat 端決定。
2. **`IF_Hdr*` 當查表鍵是否允許**：契約第 2 條第 1 項字面禁止 `IF_Hdr*`；本模型用 `IF_HdrTask`（任務名稱）與 Interface 的世代／成本表頭當 SUMIFS 的鍵，共用快照工具自己也用 `IF_HdrGen`／`IF_HdrCost` 拆世代。建議契約明訂「表頭可作鍵、不可作數值」，並讓工具支援 `IF_HdrTask`。
3. **共用快照工具擴充**：接受 `IFW_`、`IFC_`，並能依取數點帶出該列的用途欄；否則各公司都無法依契約第 2 條第 2 項取四層。
4. **研發 GW 是否需要 Anthropic 版的 Alloc 錨點（G16 自行覆寫）**：目前用殘差，與 OpenAI 一致；若要外部錨點，應由 Tokenomics 以 Anthropic 參數跑 Alloc_In，而非本模型重算。
5. **對話每任務 token 是否對齊 v5.29**：Tokenomics v5.29 把 ChatGPT「每則提示 token 數」從 2,000 改為 4,000（2,000–6,400，只用在 Block 6）。本模型對話任務用 Tokenomics「一般聊天」每次嘗試 3,000 token × 1.0（INP_105，0.5–2）。兩者口徑不同（每則提示 vs 每次嘗試），且 η 校準會吸收大部分水準差異；是否調整 INP_105 由 chat 端判斷。
6. **非 NVIDIA 舊世代比例與每 GW 年租用價**：TPU v6e／Trainium2 以新世代比例代理、每 GW 年合約價用 OpenAI 的 Analogy 12 $B——是否走 G2 補進 Tokenomics。

## 提交紀錄
- 本報告提交：見 PR（分支 `claude/anthropic-tk-link-review`）。
