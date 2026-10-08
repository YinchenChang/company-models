# Anthropic 收支模型 v0.1 規格（r1，2026-10-08，chat 端）

- 修訂 r0 → r1（A1 資料審查後）：D5、D6、D10、D15、D16 改寫；新增 D21–D23。改動處標「r1」。

- 依據：Andy 2026-10-08「請比照 OpenAI，做 Anthropic」；OpenAI v0.6（`openai/`，分支 `claude/openai-s6-release` @ `e91df57`）的命題、頁結構、具名範圍與工具鏈；Tokenomics master `3dd1216`（CURRENT `20261007_Tokenomics_v5.26.xlsx`，與 OpenAI v0.6 同一快照，兩家可直接並排）。
- 授權：Andy 2026-10-07「改吧！就讓我們試試看」——新公司模型由 Claude 套用預設、不先確認；命題與驅動對照表隨成品交付。本規格第 3 節 D1–D20 即本版預設，全部在 v0.1 完成報告彙總給 Andy 審查。
- 範圍：做到 OpenAI v0.6 的水準（完整、具代表性），但不搬 OpenAI 的 v0.5 遷移層（Map_v05、Derived_V9 不建）。OpenAI v0.6 待辦（v0.7 候選）在本模型也先不做。

---

## 1. 一句話命題（與 OpenAI 同一句，換公司）

**Anthropic 每 VR 等值 GW 的年營收，能否覆蓋每 GW 的年全成本？若不能，缺口要多少外部資金、由誰以什麼條件提供？**（FY2025–FY2030，曆年制，2025 為實際校準年。）

命題的兩個核心輸出（具名範圍與 OpenAI 同名）：
1. 每 VR 等值 GW 差額 `COST_PropGap_VR`＝營收淨額 ÷ 供給 VR 等值 GW − 全成本 ÷ 供給 VR 等值 GW；覆蓋率 `COST_Coverage`。
2. 累計外部資金需求 `FND_ExtNeedCum`，及其峰值、首次缺口年。

附帶的比較問題（只在 HTML 與 Checks 並排，不改計算）：Anthropic 以企業與 API 為主、免費用戶負擔小的營收結構，是否使每 GW 經濟優於 OpenAI？——由 `OAI_Link`（OpenAI v0.6 命題輸出快照）並排回答。

## 2. 頁結構與因果鏈

```
Tokenomics（TK_Link：各世代 × 層級每 GW token 產能、利用率、持有成本、每 GW 資本支出；NonNV 產出比與持有比）
        │
SRC_ANT（公司財務原始數據）＋ Inputs（假設，附標記與區間）
        │
        ▼
Demand：個人用戶（方案別）× 任務 × 每任務 token；企業席位；API 任務 × 每任務 token（2025 由 API 營收倒推）──► token 量（DEM_Tok_*：層級 × 付費／免費）
        │                                                    │
        ▼                                                    ▼
Revenue：訂閱（方案別，個人／企業）＋ API（層級 × 輸入／輸出 × 有效價；直接／雲端通路）     Compute：token ÷ 每 GW 產能 × η ──► 有效推論 GW
  ──► 營收總額 − 雲端夥伴分成＝淨額；容量上限係數回乘                                       合約（Google TPU、AWS Trainium、Azure NVIDIA…）＋自建 ──► 供給 GW、VR 等值 GW
        ▼                                                                                     研發 GW＝供給×(1−閒置) − 推論（殘差）
Cost：算力成本（合約實付＋自有資本支出）＋非算力成本（人數 × 每人成本）＋股權報酬 ──► 每 VR 等值 GW：營收 vs 全成本（命題 1）
        ▼
Funding：自由現金流 → 期初現金 → 已到位融資 → 外部資金需求（命題 2）；策略投資人 vs 同一方的算力合約（回流對照）
Reverse：管理層營收目標 → 所需倍數與反向資金（只作對照，不回饋）
OAI_Link：OpenAI v0.6 命題輸出（只被 Checks 與 HTML 引用）
Checks：讀全部頁；CHK_Errors＝0
```

| 頁 | 內容 | 主要具名範圍（與 OpenAI 同名者以 * 標） |
|---|---|---|
| README | 版本、命題、頁說明、Tokenomics 與 OpenAI 快照版本 | — |
| SRC_ANT | 公司原始數據（每列：ID、項目、值、低、高、單位、期間、來源、網址、文件日期、擷取日、標記、利害方、來源等級／立場／立場說明／一手二手四欄、備註） | `SRC_ANT_nnn`（＋`_Lo`／`_Hi`） |
| TK_Link | Tokenomics 快照：OpenAI v0.6 的 63 名＋NonNV 表（見 D6） | `TK_*` |
| OAI_Link | OpenAI v0.6 命題輸出快照（逐年） | `OAI_*` |
| Inputs | 假設（值、低、高、標記、依據、區間理由） | `INP_nnn` |
| Demand | 用戶、席位、任務、token | `DEM_Users_*`、`DEM_Seats_*`、`DEM_Tok_*`* |
| Revenue | 訂閱方案別、API 層級別與通路別、其他、總額、分成、淨額、截頂 | `REV_Sub`*、`REV_API`*、`REV_Gross`*、`REV_PartnerShare`、`REV_Net`*、`REV_GrossCapped`*、`REV_NetCapped`*、`REV_Consumer`、`REV_Enterprise` |
| Compute | 加速器組合、每 GW 產能、推論 GW、η、研發 GW、供給、容量上限、VR 等值 | `CMP_InfGW_Eff`*、`CMP_RDGW`*、`CMP_DemandGW`*、`CMP_SupplyGW`*、`CMP_InfAvailGW`*、`CMP_CapFactor`*、`CMP_Supply_VReq`*、`CMP_Eta` |
| Cost | 算力成本（逐合約＋自有）、供應商持有成本與雲端毛利、非算力成本、股權報酬、命題表、經濟口徑 | `COST_Compute`*、`COST_NonCompExSBC`*、`COST_SBC`*、`COST_FullCash`*、`COST_PropRev_VR`*、`COST_PropFull_VR`*、`COST_PropGap_VR`*、`COST_Coverage`*、`COST_CloudGM`* |
| Funding | 自由現金流、來源順序、外部資金需求、條件式融資、或有負債、回流對照 | `FND_FCF`*、`FND_Committed`*、`FND_ExtNeed`*、`FND_ExtNeedCum`*、`FND_CashEnd`*、`FND_Contingent`* |
| Reverse | 管理層目標與反解 | `RVS_Target`*、`RVS_Mult*`*、`RVS_ExtNeedCum`* |
| Checks | C01 起；`CHK_Errors` 必須 0 | `CHK_Errors`* |
| _Defaults | 隱藏；Excel 優先機制 | — |

## 3. 本版預設 D1–D20（Claude 依 Andy 2026-10-07 授權套用；v0.1 審查時 Andy 可逐項改）

| ID | 預設 | 替代 | 理由 |
|---|---|---|---|
| D1 | 命題、期間、2025 校準年、具名範圍與 OpenAI v0.6 相同 | 另立 Anthropic 專屬命題 | 兩家可並排；同一 Tokenomics 快照 |
| D2 | 營收線：①訂閱（方案別：Free／Pro／Max 5x／Max 20x 個人；Team 標準／Premium、Enterprise 企業席位）②API（模型層級 Haiku／Sonnet／Opus × 輸入／輸出）③其他（政府、授權等）。另以個人 vs 企業（訂閱企業席位＋API）兩個彙總列並列 | 只分訂閱／API | Andy：訂閱一律依方案拆；營收拆消費者 vs 企業 |
| D3 | 廣告＝0（Decision；Anthropic 公開承諾 Claude 不放廣告），不設驅動 | 比照 OpenAI 設慢成長驅動 | 公司公開立場；若反轉再加 |
| D4 | Claude Code 不另立營收線（避免與訂閱、API 重複）；在需求面以「程式代理任務」類別拆出 token，另列 Claude Code 年化營收 SRC 對照列 | 獨立營收線 | 其營收實際落在 Max／Team Premium 訂閱與 API 用量 |
| D5 | API 通路：直接 vs 雲端夥伴（AWS Bedrock、Google Vertex、Microsoft Foundry）；營收總額與淨額並列（類比 OpenAI 的 Microsoft 分成），夥伴分成率為 Inputs（Analogy，區間）；命題與現金流用淨額。若查得 Anthropic 營收報導口徑為淨額，則總額＝淨額、分成列只作揭露，並在報告寫明 | 只算淨額 | 與 OpenAI 同一口徑才可比 |
| D6 | 每 GW 產能：NVIDIA 世代直接取 TK `IF_TokGW_*`；TPU v7 Ironwood、Trainium3 ＝ VR200 的 `IF_TokGW_*` × TK NonNV「產出比」（基準，低／高作區間）；同代以前的 TPU v6e／Trainium2 ＝同年份 NVIDIA 世代（Hopper）的 `IF_TokGW_*` × 同一 NonNV 產出比。NonNV 表不是具名範圍：`tk_link` 以列標籤讀表快照，狀態欄寫「讀表（非具名）」，列入 Tokenomics 缺口（建議 Tokenomics 把 NonNV 比例列進 Interface） | 自建 TPU／Trainium 產能 | 算力物理一律取自 Tokenomics，不在本模型重算 |
| D7 | 加速器組合時變（同 OpenAI V1）：由逐合約供給 GW 加權（TPU／Trainium／NVIDIA 各代），2025 以 Trainium2＋TPU v6e＋Hopper 為主 | 固定組合 | 合約揭露即組合 |
| D8 | η（有效產出係數）＝2025 token 換算推論 GW ÷ 2025 推論支出換算 GW（同 OpenAI V2）；基準沿用 2025 值，情境線性回升至 2030＝1 | 不校準 | 同 OpenAI |
| D9 | 每 GW 年合約價＝基準 12 $B/GW/年（Analogy 8–20，同 OpenAI V8）× 該加速器族的 TK NonNV「持有比」（NVIDIA＝1） | 單一價格 | TPU／Trainium 持有成本較低，合約價應同比例 |
| D10 | 算力成本＝合約實付＋自有資本支出（現金口徑；同 OpenAI V3）；只揭露金額不揭露 GW 的合約：GW＝金額 ÷ 期間 ÷ 合約價；只揭露 GW 不揭露金額者：金額＝GW × 合約價。供應商持有成本＝GW × TK `IF_HoldEcon` × 持有比；差額＝雲端毛利 | — | 同 OpenAI |
| D11 | 自建（Fluidstack 等 Anthropic 出資的資料中心）＝自有資本支出（公告金額與時程；無時程者按 Inputs 均攤），自投產年起計入供給與命題分母；每 GW 資本支出對照 TK `IF_CapexTotal` | 視為租賃 | 成本與分母同口徑 |
| D12 | 研發 GW＝供給 ×（1−閒置）− 有效推論 GW（殘差，同 OpenAI）；並列 TK `IF_AllocRDGW` 對照 | 研發 GW 外部錨點 | 同 OpenAI（OpenAI 待審預設 #1，一起審） |
| D13 | 非算力成本：2025 由公開報導的營業費用（或燒錢）減算力支出得出；之後人數 × 每人成本（人數成長 Inputs）；股權報酬單列，命題「全成本」含股權報酬（較保守），另並列不含者。搜不到 2025 費用：以人數 × 同業每人成本（Analogy，OpenAI v0.6 每人成本）暫代並標明 | — | 同 OpenAI V4 |
| D14 | 2025 校準：營收以 2025 認列營收實際值（SRC）校準；年化營收里程碑（run-rate）只作對照列（年度營收≈當年各月 run-rate 平均），不反推參數 | 以 run-rate 校準 | Andy：不以公司數字反推 |
| D15 | 2026 已過三季：2026 需求與價格仍由驅動推導；2026 已知 run-rate 里程碑在 Checks 對照（差距 >25% 立 WARN 並在報告說明，不自動校準） | 2026 也校準 | 同 D14 |
| D16 | 融資來源順序（同 OpenAI／CRWV S7）：期初現金 → 已到位股權（已宣布且無條件者）→ 外部資金需求（補足至最低現金）。「最高可達（up to）」的策略投資與條件式承諾不計入基準，列情境；循環信貸額度列或有 | 計入 | 保守 |
| D17 | 最低現金＝Inputs（Assumed；基準為次年非算力成本的 6 個月，區間 3–12 個月） | OpenAI 固定 10 | Anthropic 規模不同，固定值不可比 |
| D18 | 回流對照：策略投資人（Amazon、Google、Microsoft、NVIDIA）各自的投資額 vs Anthropic 對同一方的算力合約金額，Funding 頁只列對照、不沖銷 | 沖銷 | 檢驗「誰出錢」的命題第二問 |
| D19 | 反向模式：管理層營收目標（各年基準／樂觀）與現金流轉正年 → 所需倍數與反向資金（同一支出）；只作對照，不回饋（同 OpenAI） | — | Andy：不以公司數字反推 |
| D20 | OAI_Link：讀 OpenAI v0.6 Excel（`openai/model/CURRENT` 所指檔；建置時以分支 `claude/openai-s6-release` 或其合併後的 `main`），快照其命題輸出逐年值與檔案 SHA-256；只被 Checks 與 HTML 引用 | 不並排 | 比較問題 |


### r1 修訂（A1 資料審查後，chat 端；取代上表同 ID 各列）

| ID | r1 預設 | 替代 | 依據（SRC_ANT 見 `data/anthropic_src.yaml`） |
|---|---|---|---|
| D5 r1 | Anthropic 對雲端市集（Bedrock、Vertex）銷售以**總額**認列，平台抽成記為行銷費用（草擬公開說明書，經 Reuters 轉述）。本模型：`REV_Gross`＝報導口徑（總額）；`REV_PartnerShare`＝通路營收 × 平台抽成率（Inputs：基準取說明書隱含約 16%，區間取 Reuters 隱含值與 BofA 估計推得的高值）；`REV_Net`＝總額 − 抽成。命題與現金流用淨額（與 OpenAI 扣 Microsoft 分成同口徑）；抽成不再列入非算力費用，避免重複 | 用總額 | 2025 經雲端通路占 47% |
| D6 r1 | 加速器族增加 AMD（MI450／MI455X Helios，取 TK NonNV「AMD MI455X」列）；NVIDIA 經 SpaceX／xAI、Nscale、Lambda、Volta 等取得者歸 NVIDIA 族（世代依揭露：Vera Rubin＝VR200，其餘 GB200／GB300） | — | 2026 新合約 |
| D10 r1 | 供給 GW 只由「算力合約」（含晶片或算力服務者：Google／Broadcom TPU、AWS、Azure、AMD、SpaceX／xAI、Nscale、Lambda、Volta、CoreWeave、Akamai、Fluidstack 自建）加總；**機房租約**（TeraWulf、Riot、Hut 8、Nexus 等只含建物與電力）不另計 GW，只登 SRC 並在 Compute 對照列標明「容量已含於 TPU／AMD 合約」（不重複）。說明書總承諾約 $518B（約 80% 不可取消）作為逐合約加總的對帳列 | 機房租約另計 | 避免重複計算 |
| D15 r1 | 2026 為**半校準年**：2026 營收基準＝2026 Q1＋Q2 實際（SRC，說明書）＋下半年＝最新觀測年化營收 ÷ 2（持平，保守；Decision）；2026 需求驅動（API 任務成長等）校準到該總額（同 D8 的「以實際值校準」，非以公司目標反推）。2027 起由驅動正向推導。2026 年底投資人預期年化 $100–120B 只作對照列 | 2026 純驅動推導（將嚴重偏離已觀測實際值） | 2026 已過三季且有季度實際值 |
| D16 r1 | 已到位：2025 年底現金與短期投資（SRC）；Series H $65B（2026-05，含 $15B 早先超大雲端業者承諾，避免與各家「已到位」重複計）；Microsoft $5B（已全額）。條件式（情境）：Amazon 依算力交付里程碑的 $15B、Google 依績效的最高 $30B、NVIDIA 最高 $10B 未揭露到位部分、**IPO 募資（報導約 $100B，2026-11）**。或有：$15B 循環信貸、表外 TPU 融資（只有單一來源者標未證實） | IPO 計入基準 | 保守；IPO 尚未發生 |
| D21（新增） | 2025 淨損 $42B 中約 $34B 為非現金會計費用：Funding 一律用現金口徑；Cost 的命題表不含該非現金項 | — | 說明書 |
| D22（新增） | 2026 Q2 首次調整後營業利益（說明書／報導）：列 Checks 對照（模型 2026 營業口徑 vs 報導），不校準 | — | — |
| D23（新增） | 年化營收里程碑（含 YipitData 等第三方估計，Analogy）全部只作對照；HTML 以里程碑折線並列模型年營收 | — | — |

## 4. A1 資料需求（SRC_ANT；每筆依共同規則第 4 節的欄位）

A. 營收：2024、2025 認列營收；年化營收里程碑（2024-12 起到最新，每筆日期＋來源）；營收結構（API vs 訂閱 vs 其他；個人 vs 企業；直接 vs 雲端通路）；Claude Code 年化營收里程碑；前幾大客戶集中度（如 Cursor、GitHub Copilot）；2025 毛利率（實際與預測）；2025 推論成本、訓練成本。
B. 價格：API 牌價逐模型（Haiku／Sonnet／Opus 各版本、上市日、輸入／輸出、快取寫入／讀取、批次折扣、長上下文加價）；價格事件（如 Opus 4.5 降價）；訂閱方案價格（Pro 月付／年付、Max 5x／20x、Team 標準／Premium、Enterprise 席位價的報導值）。
C. 用戶：Claude 月活或週活（公司或第三方估計）；個人付費訂閱數；企業客戶數；大客戶數（年化 >$100k）；企業席位大單（如 Deloitte、Cognizant 等）；token 處理量（若有揭露；OpenRouter 份額只作 Analogy）。
D. 算力：每份合約（對象、日期、金額、GW 或晶片數、晶片型號、期間、起始與爬坡、IT 或設施口徑）：Google TPU（2025-10 及其後擴大）、AWS Project Rainier（Trainium2 數量、Indiana 園區 GW、Trainium3）、Microsoft Azure $30B（2025-11）與 NVIDIA 最高 1 GW、Fluidstack $50B 自建（德州、紐約，時程）、2026 年新增的任何合約；公司或媒體的算力支出預測（逐年）；總 GW 目標。
E. 成本：人數時間序列；股權報酬；營業費用、銷售與管理；2025 燒錢與 2026 以後預測。
F. 融資：每一輪（日期、金額、投前／投後估值、領投、結構）；策略投資（Amazon 累計、Google 累計、Microsoft 最高 $5B、NVIDIA 最高 $10B，條件與到位情形）；員工股份收購；循環信貸；2025 年底現金（若有）；債務；2026 年各輪；IPO 報導。
G. 管理層目標／預測：各年營收（基準、樂觀）、現金流轉正年、算力支出預測、毛利率目標。

「找不到」與「不存在」分開寫，列出試過的來源；公司自己的說法標 Interested-party（利害方 Anthropic，誘因：募資與招募）；外流的投資人簡報同標 Interested-party。

## 5. 分段

| 段 | 內容 | 版本 |
|---|---|---|
| A0 | 骨架、規格、共同規則、工作單、進度檔（chat 端） | — |
| A1 | 資料蒐集 → `data/anthropic_src.yaml`、研究筆記、資料報告 | — |
| A2 | builder 與 README、SRC_ANT、TK_Link（含 NonNV 讀表）、OAI_Link、Inputs、Demand、Revenue、Checks；CI | v0.1-A2 |
| A3 | Compute、Cost（命題表） | v0.1-A3 |
| A4 | Funding、Reverse、Checks 補齊 → v0.1 完成報告（預設彙總） | v0.1 |
| A5 | HTML 一頁摘要（含 OpenAI 並排）、命題與驅動對照表定稿、交接檔、dist | v0.1 成品 |
