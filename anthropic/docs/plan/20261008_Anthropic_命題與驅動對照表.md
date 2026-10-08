# Anthropic 收支模型 v0.1：命題與驅動對照表（2026-10-08，A5）

- 模型：`model/20261008_Anthropic_v0.1.xlsx`（`model/CURRENT`）；Tokenomics 快照 v5.26（`20261007_Tokenomics_v5.26.xlsx`；建置時 master `70d859e`，CURRENT 與 `3dd1216` 相同）；OpenAI v0.6 快照（`OAI_Link`，SHA-256 `34fc34a7…`）。
- 用途：說明「命題由哪些頁、哪些輸入、經什麼因果方向算出」，供 Andy 與投資處同事審查與追溯。數值以 Excel 為準（一頁摘要 `dist/20261008_Anthropic收支模型_v0_1.html`）；本表只列結構。
- 讀法：「→」表示因果方向（左邊決定右邊）；ID 為各頁 A 欄的列 ID，具名範圍以 `等寬字` 表示。具名範圍與 OpenAI v0.6 同名（`DEM_`、`REV_`、`CMP_`、`COST_`、`FND_`、`RVS_`），兩家可直接並排。

## 1. 一句話命題
**Anthropic 每 VR 等值 GW 的年營收，能否覆蓋每 GW 的年全成本？若不能，缺口要多少外部資金、由誰以什麼條件提供？**（與 OpenAI v0.6 同一句；Andy 2026-10-08「請比照 OpenAI，做 Anthropic」；FY2025–FY2030，曆年制，2025 為實際校準年。）

命題的兩個核心輸出：
1. 每 VR 等值 GW 差額＝營收淨額 ÷ 供給 VR 等值 GW − 全成本 ÷ 供給 VR 等值 GW（`COST_PropGap_VR`；Cost K102）；覆蓋率 `COST_Coverage`（K104）。並列口徑：經濟口徑 `COST_PropGapEcon_VR`（K109）、每實體 GW `COST_PropGap_GW`（K107）。
2. 累計外部資金需求（`FND_ExtNeedCum`；Funding F33），及首次缺口年 `FND_FirstGapYear`（F41）、只靠已到位融資的現金谷底 `FND_CashNoExtTrough`（F44）。

## 2. 主因果鏈（正向；反向只作對照）

```
Tokenomics（TK_Link：每 GW token 產能、利用率、持有成本、每 GW 資本支出、TPU／Trainium／AMD 相對 NVIDIA 的產出比與持有比）
        │
SRC_ANT（公司財務原始數據）＋ Inputs（假設，附標記與區間）
        │
        ▼
Demand：個人付費人數、企業席位 × 任務 × 每任務 token；API：任務數 × 每任務 token ×（價格反應）──► token 量（DEM_Tok_*）
        │                                                │
        ▼                                                ▼
Revenue：牌價 × 計費 token（API）、月費 × 人數（訂閱）    Compute：逐合約供給 GW（金額 ÷ 合約價，或揭露 GW × 爬坡）＋自建
  ──► 營收總額（報導口徑）− 雲端平台抽成＝淨額            token ÷ 每 GW 產能 ÷ η ──► 有效推論 GW
        │                                                研發 GW＝供給 ×（1−閒置）− 推論（殘差）
        │        ▲ 容量上限係數（推論需求 > 推論可用 GW 時截頂）    │
        └────────┴─────────────────────────────────────┘
        ▼
Cost：算力成本（逐合約實付＋自建資本支出）＋非算力成本（人數 × 每人成本）＋股權報酬
  ──► 每 VR 等值 GW：營收淨額 vs 全成本（命題 1）
        ▼
Funding：自由現金流＝營收淨額 − 算力成本 − 非算力成本（股權報酬加回）
  ──► ①期初現金 → ②2026 已交割股權 → ③外部資金需求（補足至最低現金；命題 2）
       條件式（Amazon、Google、NVIDIA、AMD）與 IPO 只列情境；循環信貸、Broadcom 租賃融資、SPV 只列或有

Reverse（管理層營收目標 → 所需倍數與反向資金）：只被 Checks 引用，不回饋任何一頁。
OAI_Link（OpenAI v0.6 命題輸出快照）：只被 Checks 與 HTML 引用，不進任何計算頁。
Checks：讀取全部各頁，CHK_Errors 必須為 0。
```

## 3. 各頁的輸入、衍生量、因果方向

| 頁 | 角色 | 主要輸入（來自） | 主要衍生量（具名範圍／列 ID） | 因果方向（誰決定誰） | 繼承 Tokenomics？ |
|---|---|---|---|---|---|
| **SRC_ANT** | 第 0 層：公司財務原始數據（379 列） | 招股書草稿轉述（Reuters、Fortune）、公司新聞稿、8-K／10-Q、媒體（來源、日期、標記、利害方四欄） | `SRC_ANT_nnn`（值、低、高）；Derived 列為公式 | 只被引用，不計算（Derived 除外） | 否（公司專屬） |
| **TK_Link** | 第 0 層：Tokenomics 快照（68 名＋NonNV 18 格＋PUE 3 格） | Tokenomics master `model/CURRENT`（v5.26），builder 寫入 | `TK_IF_TokGW_*`、`TK_IF_Util`、`TK_IF_HoldEcon`、`TK_IF_CapexTotal`、`TK_IF_FullCost*`、`TK_IF_TaskTok*`、`TK_NNV_*_Out／Hold` | 只被引用；本模型不重算算力物理 | **是（全部）** |
| **OAI_Link** | 跨模型對照：OpenAI v0.6 命題輸出（17 名） | OpenAI v0.6 活頁簿快取值，builder 寫入 | `OAI_COST_PropGap_VR`、`OAI_FND_ExtNeedCum` 等 | 只被 Checks、HTML 引用（測試掃描） | 間接（同一 TK 快照） |
| **Inputs** | 假設（175 列；Assumed／Analogy／Decision；退役編號不重用） | A2–A4 工作單預設、規格 D1–D23 | `INP_nnn`（值、低、高、標記） | 只被引用；情境與敏感度改這裡 | 否 |
| **Demand** | 需求：個人人數、企業席位、任務、token | SRC_ANT（2025 訂閱營收、月活）、Inputs（人數成長、任務數與每任務 token 成長、彈性、層級組合）、TK（`IF_TaskTok*` 任務別每次嘗試 token） | `DEM_Users_*`（D04–D08）、`DEM_Seats_*`（D12–D15）、`DEM_ApiTaskGrowth`（D46）、`DEM_Tok_API`（D45）、`DEM_Tok_Paid`／`Free`／`Total`（D56–D58） | 人數 × 任務 × 每任務 token → token；API：任務數 × 價格反應 → 計費 token；2025 由實際營收倒推 | 部分：每任務 token 取 TK |
| **Revenue** | 營收：牌價（價格事件加權）、訂閱方案別、API 層級別、其他、抽成、淨額 | SRC_ANT（牌價與上市日、月費、2025 認列營收、2026 Q1／Q2、run-rate、雲端通路費）、Inputs、Demand token | `REV_SubConsumer`（R41）、`REV_SubSeats`（R42）、`REV_API`（R45）、`REV_Gross`（R52）、`REV_PartnerShare`（R56）、`REV_Net`（R57）；截頂後 `REV_GrossCapped`／`REV_NetCapped`（R65、R67）；對照 `REV_RREst2026`、`REV_RRGap2026`（R74–R75） | token × 有效單價＋人數 × 月費 → 總額 → 扣雲端平台抽成 → 淨額；Compute 的容量上限係數回乘（R64）；2025 校準、2026 半校準（D14、D15 r1） | 否（公司專屬） |
| **Compute** | 算力：加速器族 9 族、逐合約供給、自建、組合、η、研發 GW、容量上限、VR 等值、敏感度 | TK（每 GW 產能、利用率、NonNV 比例、每 GW 資本支出）、SRC_ANT（合約金額、GW、期間、2025 推論支出、年底 GW 報導）、Inputs（合約價 12、爬坡、期間、計入比例、閒置） | `CMP_Supply_*`（C59–C72）、`CMP_Supply_Owned`（C77）、`CMP_SupplyGW`（C79）、`CMP_InfGW_Eff`（C137）、`CMP_InfAvailGW`（C140）、`CMP_RDGW`（C142）、`CMP_CapFactor`（C144）、`CMP_Supply_VReq`（C150） | 合約付款 ÷（合約價 × 持有比）或揭露 GW × 爬坡 → 供給；token ÷ 每 GW 產能 ÷ η（2025 以推論支出校準）→ 有效推論 GW；**研發 GW＝殘差**；供給 × 機隊 VR 等值係數 → 命題分母 | **是**：產能、利用率、NonNV 比例、資本支出；**否**：合約、組合、η、研發殘差 |
| **Cost** | 成本、命題表、對帳與敏感度 | Compute（供給、自建、VR 等值）、SRC_ANT（2025 算力 7.33、營業費用、人數、招股書逐家承諾）、TK（`IF_HoldEcon`、`IF_FullCost*`）、Inputs（人數成長、每人成本、股權報酬） | `COST_Compute`（K45）、`COST_NonCompExSBC`（K78）、`COST_SBC`（K77）、`COST_FullCash`（K93）、`COST_GapCash`（K95）、`COST_PropRev_VR`（K96）、`COST_PropFull_VR`（K100）、`COST_PropGap_VR`（K102）、`COST_Coverage`（K104）、`COST_CloudGM`（K64）；逐家對帳 K143–K149；S4 錨定 `COST_ComputeAnchor`（K156）；S5 租金 `COST_LeaseRent`（K170）；D22 `COST_D22Under`（K176） | 合約實付＋自建資本支出 → 算力成本；人數 × 每人成本 → 非算力；÷ 供給 VR 等值 GW → 命題輸出 | **是**：持有成本與 TK 單位成本只作參考列；算力成本本身＝合約實付（公司專屬） |
| **Funding** | 融資 | Cost（`COST_GapCash`、`COST_SBC`）、Revenue（`REV_NetCapped`）、SRC_ANT（2025 年底現金 20.28、Series G／H、Amazon 特別股、條件式、或有負債）、Inputs（最低現金月數、開關、到位年） | `FND_FCF`（F07）、`FND_Committed`（F18）、`FND_MinCash`（F28）、`FND_ExtNeed`（F31）、`FND_CashEnd`（F32）、`FND_ExtNeedCum`（F33）、`FND_Headroom`（F35）、`FND_CashNoExtTrough`（F44）、情境 S1–S7（F46–F80）、`FND_Contingent`（F84）、回流對照 `FND_RoundTrip`（F87–F94） | 自由現金流 → ①期初現金 → ②已交割股權 → ③外部資金（補足至最低現金） | 否（公司專屬） |
| **Reverse** | 反向模式（D19；只作對照） | SRC_ANT（管理層營收目標各版）、正向各營收線 | `RVS_Target`（X01）、`RVS_MultAPI`／`MultSub`／`MultProp`（X10–X12）、`RVS_ExtNeedCum`（X21） | 目標 − 正向 → 所需倍數；目標營收＋同一支出 → 反向外部資金 | 否；**不回饋基準**（`test_reverse_not_fed_back`、Checks C88） |
| **Checks** | 治理：C01–C105 | 全部頁 | `CHK_Errors`（必須＝0）、`CHK_Warnings`（基準 1：C89 逐家合約 AMD） | 只讀 | 對照 TK 快照版本 |

## 4. 繼承 Tokenomics vs 公司專屬

| 類別 | 內容 | 位置 |
|---|---|---|
| 繼承 Tokenomics（本模型不重算） | 各世代、各層級每 GW 年 token 產能（`IF_TokGW_Astra／Sol／Luna`）；基準利用率（`IF_Util`）；TPU v7、Trainium3、AMD MI455X 相對 VR200 的產出比與持有比（NonNV 表，讀表）；每 GW 年持有成本（`IF_HoldEcon`）；每 GW 資本支出（`IF_CapexTotal`）；TK 單位全成本參考；任務別每次嘗試 token（`IF_TaskTok*`）；研發算力參考（`IF_AllocRDGW`）；PUE（讀表） | TK_Link → Demand、Compute、Cost、Checks |
| 公司專屬（SRC_ANT） | 2025 認列營收與訂閱／用量拆分、2026 Q1／Q2、run-rate 里程碑、牌價與上市日、方案月費、月活、雲端通路費、逐合約金額／GW／期間（Google、Broadcom、Amazon、Azure、AMD、SpaceX、Nscale、Lambda、Volta、Akamai、Fluidstack）、招股書逐家承諾、2025 算力與營業費用、人數、年底現金、各輪融資、條件式、或有負債、管理層目標 | SRC_ANT → 各頁 |
| 公司專屬（Inputs 假設） | 個人人數成長、方案組合、API 任務數與每任務 token 成長（與 OpenAI v0.6 同值）、價格彈性、牌價年變動、層級組合、有效價組成、雲端平台抽成率、合約價 12（Analogy 8–20）、爬坡、合約期間、「最高可達」計入比例、自建均攤、人數成長、每人成本、股權報酬、最低現金月數、融資開關與到位年 | Inputs → 各頁 |
| 本模型自算（介於兩者） | η（2025 以推論支出校準，D8）、研發 GW 殘差（D12）、加速器組合（逐合約 GW 加權，D7）、VR 等值係數、容量上限係數 | Compute |

## 5. 與 OpenAI v0.6 模型的關係
- **程式碼**：engine、tests、tools（`build_html.py`、`screenshot_html.py`、`check_tk_snapshot.py` 等）自 `openai/` 複製改寫；builder 依 OpenAI 的頁序與命名重寫為 Anthropic 版（`builder/a2.py`–`a4.py`）。未改動 `openai/`。
- **共用**：同一句命題、同一 Tokenomics 快照（v5.26）、同一每 GW 年合約價 12、同一 API 任務數成長與彈性（A2-12、A2-13）、同名具名範圍、同一 Excel 唯一引擎與 E6 規則、同一成品型式（HTML 一頁摘要＋Excel）。
- **不同處**：①營收線：Anthropic 以企業與 API 為主（2030 企業占 93%），無廣告（D3）、無 Microsoft 分成，改扣雲端平台抽成（D5 r1）；②算力：多家非 NVIDIA 加速器（TPU、Trainium、AMD）以 NonNV 比例換算（D6），供給逐合約加總（D10）；③融資：2026 已交割 $100B 使基準不需外部資金，OpenAI 需 $442.8B；④最低現金＝次年非算力 6 個月（OpenAI 固定 10）。
- **並排**：`OAI_Link` 只被 Checks 與 HTML ⑦ 引用，不進任何計算頁。

## 6. 三個關鍵驅動（v0.1 完成報告③；數值見 HTML 第 ⑥ 節）
| 驅動 | Inputs／TK | 作用路徑 |
|---|---|---|
| 1. API 需求成長（任務數、每任務 token） | INP_092–095（任務數 2027–30）、INP_087–091（每任務 token 2026–30） | Demand D45–D46 → Revenue R45 → 淨額 → 命題與現金流（容量上限 C144 回乘） |
| 2. 價格反應（彈性、牌價年變動） | INP_096（彈性 −0.7；−1.05／−0.35）、INP_098（牌價年變動 −15%；−35%／0） | Demand 計費 token → Revenue 有效單價 × token；牌價大降時 token 暴增觸發容量上限（C144） |
| 3. 算力合約條款與產能 | INP_114（合約價 12；8–20）、INP_122（爬坡 2 年）、INP_138（AMD 期間 5 年）、`TK_NNV_*_Out`（NonNV 產出比） | Compute 供給 GW 與 VR 等值分母 → Cost 算力成本 → 命題與現金流 |
| 結構前提（非參數）：只計已簽合約 | （無；新增合約機制需 Andy 決定） | 供給 GW 2028 → 2030 遞減 → 2029–2030 轉正；餘裕上限見 `COST_GapCash`（K95）、`FND_Headroom`（F35） |
