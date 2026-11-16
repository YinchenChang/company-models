# Salesforce 研究包（原型 S：按席位計價）

擷取日 2026-11-16。會計年度 2 月制：FY2026＝2025-02-01～2026-01-31；FY2027＝2026-02-01～2027-01-31。最新已公布季＝FY2027 Q2（2026-05-01～2026-07-31，財報 2026-08-26）；FY2027 Q3（截至 2026-10-31）預計 2026-12 初公布，尚未公布。
標記：Verified／Interested-party／Analogy／Assumed／Derived；逐列紀錄見 `facts_salesforce.csv`。金額 US$M，除非註明。

## (1) 公司一句話
Salesforce（NYSE: CRM）是以訂閱制銷售 CRM 與企業雲端軟體（銷售、服務、行銷與電商、Slack、資料與整合）的公司，10-K 稱「一個營運分部」；FY2026 營收 41,525M，2025-11 以約 9.6B 美元併購 Informatica，並把產品線重新包裝為 Agentforce 系列。（出處：10-K FY2026 附註；Q4 FY26 新聞稿）

## (2) 官方報告分部與固定揭露產品線
- **官方報告分部：只有一個**（"The Company operates as one operating segment"；10-K FY2026 附註，經第三方鏡像讀取）。因此沒有分部營業利益。
- **營收拆分（法定附註 disaggregated revenue）有兩種口徑，需並列：**
  - **舊五分類（FY2024–FY2026 與 Q3、Q4 FY26 仍用；FY2026 Q3 起名稱冠上 Agentforce，營收歸屬不變）：** Sales、Service、Platform and Other（FY26 起稱 Agentforce 360 Platform, Slack and Other；含 MuleSoft 以外的平台、Slack、Informatica）、Marketing and Commerce、Integration and Analytics（MuleSoft、Tableau 等）。
  - **新二分類（FY2027 Q1 起；Q1 FY26 起回溯重分類，訂閱與支援合計不變）：** (a) Agentforce Apps＝Agentforce Sales、Service、Marketing、Commerce、Apps Flex Credits、Slack；(b) Data 360, Headless Platform, and Other＝Data 360、Data 360 與 Platform Flex Credits、Headless Platform、Informatica、Agentforce MuleSoft、Agentforce Tableau、其他。另有 Hedging gains 單列。（出處：Q2 FY27 新聞稿 Revenue by Service Offering 註記）
- **營收型態：** 訂閱與支援（subscription and support）、專業服務與其他（professional services and other）。
- **區域（法定揭露）：** Americas、Europe、Asia Pacific。
- 公司**未**回溯公布 FY2026 全年與 Q3、Q4 FY26 的新二分類值（Q1 FY27 新聞稿只回溯 Q1 FY26；Q2 FY27 回溯 Q2 FY26）。10-Q 是否含更多回溯，本次未讀到原檔（見 (9)）。

## (3) 分部營收（US$M）

### 3a. 營收型態：最近 5 個完整會計年度
| 項目 | FY2022 | FY2023 | FY2024 | FY2025 | FY2026 |
|---|---:|---:|---:|---:|---:|
| 訂閱與支援 | 24,657 | 29,021 | 32,537 | 35,679 | 39,388 |
| 專業服務與其他 | 1,835 | 2,331 | 2,320 | 2,216 | 2,137 |
| 合計 | 26,492 | 31,352 | 34,857 | 37,895 | 41,525 |
| 公司揭露合併營收 | 26,492 | 31,352 | 34,857 | 37,895 | 41,525 |
| **核對列（合計－公司值）** | **0** | **0** | **0** | **0** | **0** |

### 3b. 舊五分類：訂閱與支援（5 個完整年度）
| 分類 | FY2022 | FY2023 | FY2024 | FY2025 | FY2026 |
|---|---:|---:|---:|---:|---:|
| Sales | 5,989 | 6,831 | 7,580 | 8,322 | 9,028 |
| Service | 6,474 | 7,369 | 8,245 | 9,054 | 9,818 |
| Platform and Other | 4,509 | 5,967 | 6,611 | 7,247 | 8,882 |
| Marketing and Commerce | 3,902 | 4,516 | 4,912 | 5,281 | 5,428 |
| Integration and Analytics | 3,783 | 4,338 | 5,189 | 5,775 | 6,232 |
| 五分類合計 | 24,657 | 29,021 | 32,537 | 35,679 | 39,388 |
| 訂閱與支援（公司值） | 24,657 | 29,021 | 32,537 | 35,679 | 39,388 |
| **核對列** | **0** | **0** | **0** | **0** | **0** |

來源：FY2024–FY2026＝10-K FY2026 MD&A（第三方鏡像讀取）與 10-K FY2025 Item 7；FY2022–FY2023 分雲值取自 Stock Analysis on Net（第二來源，稱取自 10-K；加總與 10-K 訂閱與支援總額吻合）。FY2023 分雲值可能在後續 10-K 重編過，**待回 10-K FY2023／FY2024 原檔核對**。

### 3c. 最近 4 個已公布季（FY2026 Q3、Q4；FY2027 Q1、Q2）
| 項目 | Q3 FY26 | Q4 FY26 | Q1 FY27 | Q2 FY27 |
|---|---:|---:|---:|---:|
| 訂閱與支援 | 9,726 | 10,675 | 10,593 | 10,820 |
| 專業服務與其他 | 533 | 526 | 540 | 525 |
| 合計 | 10,259 | 11,201 | 11,133 | 11,345 |
| 公司揭露合併營收 | 10,259 | 11,201 | 11,133 | 11,345 |
| **核對列** | **0** | **0** | **0** | **0** |
| 舊五分類：Sales | 2,297 | 2,333 | — | — |
| Service | 2,495 | 2,531 | — | — |
| Platform, Slack and Other | 2,180 | 2,655 | — | — |
| Marketing and Commerce | 1,361 | 1,377 | — | — |
| Integration and Analytics | 1,393 | 1,779 | — | — |
| 舊五分類合計－訂閱與支援（核對） | 0 | 0 | — | — |
| 新二分類：Agentforce Apps | 未找到（未回溯） | 未找到（未回溯） | 6,910 | 7,193 |
| Data 360, Headless Platform, and Other | 未找到（未回溯） | 未找到（未回溯） | 3,683 | 3,618 |
| Hedging gains | — | — | 0 | 9 |
| 新二分類合計－訂閱與支援（核對） | — | — | 0 | 0 |

對照（回溯後的去年同期，新二分類）：Q1 FY26＝Agentforce Apps 6,345、Data 360 等 2,952（合計 9,297）；Q2 FY26＝6,682、3,008（合計 9,690）。Q2 FY27 年增：Agentforce Apps 名目 +7.6%（固定匯率 +8%）、Data 360 等 名目 +20.3%（固定匯率 +20%）。
區域（Q2 FY27）：Americas 7,404、Europe 2,764、Asia Pacific 1,168、Hedging 9，合計 11,345。

## (4) 本原型關鍵營運指標歷史
**RPO／cRPO（US$B，期末）**
| 期末 | cRPO | RPO |
|---|---:|---:|
| 2022-01-31 | 22.0 | 43.7 |
| 2023-01-31 | 24.6 | 48.6 |
| 2024-01-31 | 27.6 | 56.9 |
| 2025-01-31 | 30.2 | 63.4 |
| 2025-04-30 | 29.6 | 60.9 |
| 2025-07-31 | 29.4 | 59.9 |
| 2025-10-31 | 29.4 | 59.5 |
| 2026-01-31 | 35.1（含 Informatica；名目 +16%、固定匯率 +13%，其中約 4 點來自 Informatica） | 72.4（含 Informatica 約 2.2） |
| 2026-04-30 | 33.6 | 67.9 |
| 2026-07-31 | 33.5（固定匯率 +14%） | 66.3（+11%） |
（來源：各季新聞稿；標記 Verified。Q3 FY27 cRPO 指引：固定匯率約 +14%，不含 Contentful 與 Fin，標 Interested-party。）

**營收成長分解（Derived）：** FY2023 +18.3%、FY2024 +11.2%、FY2025 +8.7%、FY2026 +9.6%（其中 Informatica 約貢獻 399M，約 1 點）；Q2 FY27 名目 +10.8%，扣 Informatica 456M 後有機約 +6.3%（未扣匯率）；訂閱與支援扣 Informatica 440M 後約 +7.1%。10-K FY2026：「Pricing was not a significant driver of the increase in revenues」，成長主要來自量。管理層對 FY2027 指引的拆解：上調 200M＝約 100M 有機＋200M 併購（Contentful、Fin，預計 Q3 FY27 交割）－100M 匯率（Q2 FY27 新聞稿）。

**客戶員工與席位相關揭露：** 公司**未揭露**總席位數、每席位平均收入、淨留存率或客戶人力編制。僅有口述：Q2 FY27 法說（2026-08-26）Benioff「Agentforce sales, service, and Slack all saw year-over-year growth」（席位）；財務長「durable growth in seats」；「attrition was near its lowest level ever」；Milano「Only 5% of the knowledge workers that use sales and service have agreed to the higher end editions」。皆標 Interested-party。

**Informatica：** 2025-11 收購，對價約 9.6B 美元（大部分現金）；收購日 RPO 約 2.2B、遞延收入 651M；FY2026 貢獻訂閱與支援 388M、營收 399M（歸入 Platform, Slack and Other）；Q2 FY27 貢獻訂閱與支援 440M、營收 456M；Informatica Cloud ARR 1.1B（Q4 FY26、Q1 FY27）。
**其他併購：** Contentful 與 Fin，預計 Q3 FY27 交割，FY2027 營收指引含約 200M（財務長法說）。

## (5) AI 相關揭露
| 時點 | Agentforce ARR | Agentforce＋Data 360 ARR | 交易數 | 其他 |
|---|---|---|---|---|
| Q3 FY26（2025-10） | >0.5B（+330%） | ~1.4B（+114%） | paid deals 逾 9,500（季增 50%；累計成交逾 18,500） | — |
| Q4 FY26（2026-01） | 0.8B（+169%） | >2.9B（+200%以上；含 Informatica Cloud 1.1B） | Agentforce deals 逾 29,000（累計；季增 50%；來源未註明是否付費） | AWU 累計 2.4B；token 近 20 兆 |
| Q1 FY27（2026-04） | 1.2B（+205%） | ~3.4B（含 Informatica Cloud 1.1B） | — | token 28.6 兆（季增 152%）；AWU 當季 3.8B（Futurum 轉述） |
| Q2 FY27（2026-07） | >1.5B（+240%以上；**自本季起納入 Slackbot 與 Headless 360，口徑擴大**） | ~3.9B（+210%以上） | 新增 2,000 家付費客戶投入生產（季增 70%，口述） | AWU 累計 7.0B、當季 3.2B（季增 97%）；Data 360 當季攝入 104 兆筆（+355%） |
- Data 360 ARR 無單獨揭露（僅與 Agentforce 合計）。AI 營收占比：Agentforce ARR 1.5B ÷ 年化營收 45.4B（Q2 FY27 營收×4）≈ 3.3%（Derived）。
- 以上 ARR、交易數、AWU 皆為公司自揭 KPI，標 Interested-party，誘因＝公司向市場證明 AI 變現。

**定價方案變更時間線**
| 日期 | 事項 | 出處 |
|---|---|---|
| 2023-08 | list price 平均調漲約 10% | The Register 2025-06-17（第二來源） |
| 2024-10 | Agentforce 推出：每次對話 $2 | Salesforce 新聞稿 2025-05-15 提及「earlier conversational pricing was $2 per conversation」 |
| 2025-05-15 | Flex Credits：每 100,000 點 $500、每動作 20 點＝$0.10；Flex Agreement 可在席位與點數間轉換 | Business Wire／Salesforce 新聞稿 |
| 2025-06-17 公布、2025-08-01 生效 | Enterprise／Unlimited 版（Sales、Service、Field Service、部分 Industries）list price 平均 +6%；Agentforce 加購起價 $125／人／月；Agentforce 1 Editions 起價 $550／人／月（含每組織每年 100 萬 Flex Credits） | Salesforce 官方 pricing update 2025 |
| 2025-08 | Pay-as-you-go 開放；pre-commit 有限釋出 | MarTech（轉述公司） |
| 2025-10（Dreamforce） | AELA（Agentic Enterprise License Agreement）：定額不限量，涵蓋 Agentforce、Data 360、MuleSoft、Slack，2–3 年期；費率不公開 | 第二來源（Beri 部落格）；一手文件**未找到** |
| 2026-08（Q2 FY27 法說） | 高階版（A1E 等）溢價 60–80%，採用率僅約 5%；Agentforce 訂單約 50% 來自 Flex Credits 續購 | 法說逐字稿（Interested-party） |

## (6) 毛利率、營業利益率、員工數、股票薪酬
| 項目 | FY2022 | FY2023 | FY2024 | FY2025 | FY2026 | Q2 FY27 |
|---|---:|---:|---:|---:|---:|---:|
| 毛利率（Derived） | 73.5% | 73.3% | 75.5% | 77.2% | 77.7% | 76.7% |
| GAAP 營業利益率 | 2.1% | 3.3% | 14.4% | 19.0% | 20.1% | 20.5% |
| non-GAAP 營業利益率（公司定義） | 18.7% | 22.5% | 30.5% | 33.0% | 34.1% | 34.1% |
| 股票薪酬（US$M） | 2,779 | 3,279 | 2,787 | 3,183 | 3,480（排除重組 29） | 904 |
| 股票薪酬占營收（Derived） | 10.5% | 10.5% | 8.0% | 8.4% | 8.4% | 8.0% |
| 員工數（期末） | 未找到 | 未找到 | 未找到 | 76,453 | 83,334 | 未找到 |
FY2026 每員工營收約 498K（Derived）。FY2027 指引：GAAP 營業利益率 20.1%、non-GAAP 34.3%（Interested-party）。Q2 FY27 GAAP 營業利益 2,331 與去年同期持平（Q2 FY26 2,332）；Q2 淨利 3,526 含策略投資利得 2,613（非營業）。非流動債務 39,288（2026-01-31 為 10,439）；$25B 加速回購（ASR）預計 2026-10 結算，管理層稱至少回購 14% 流通股，迄今均價約 176 美元（口述）。

## (7) 市場共識
- **FY2027 營收：** 財報前共識約 46.11B（Futurum 轉述，賣方共識，Q2 財報前；公司已上調指引至 46.1–46.4B）。Interested-party。
- **FY2027 EPS：** 16.7553（247wallst 2026-09-29 引述；口徑未載明，與公司 non-GAAP 指引 16.67–16.71 一致，推測為 non-GAAP；含策略投資利得）。
- **FY2028 EPS：** 16.0067（同上；該文稱低於 FY2027 係因本年一次性投資收益墊高）。
- **FY2028 營收共識：** **未找到**（stockanalysis.com 頁面為 2025-04 舊資料，不採用）。
- **賣方平均目標價：** 247wallst（2026-09-29）：56 則評等中 40 買進／強力買進、14 持有、2 強力賣出，股價 $234.08；該文未載明平均目標價（其模型內的「分析師共識」輸入為 $281.08，非公布的目標價）。TIKR 引 Guggenheim 2026-07-01 升評買進時稱 Street 平均目標價約 $246（單點，來自媒體）。**一致、同日期、可引用的平均目標價與家數：未找到。**

## (8) 爭議（多方、空方各 3 點）
**多方**
1. 有機成長與 AI 變現加速，Agentforce 與 Data 360 ARR 近 3.9B、年增逾 210%；cRPO 年增 14%（固定匯率）並指引 Q3 維持約 14%。（Q2 FY27 新聞稿；Interested-party）
2. 席位仍成長、高階版溢價大但採用僅 5%，有升級空間；Agentforce 訂單約一半來自點數續購，顯示已實際消耗。（Q2 FY27 法說，Milano／Benioff；Interested-party）
3. 評價與資本回報：Guggenheim（DiFucci，2026-07-01 升評買進）稱市場已把「AI 末日情境」價入一家仍雙位數成長且產生現金的公司，約 3 倍營收、約 12 倍 NTM 盈餘；$25B ASR 約 14% 股數。（TIKR 引述；賣方，Interested-party）

**空方**
1. 席位壓縮：投資人擔心 AI 代理減少所需人工使用者席位，侵蝕按席位計價核心；2026 年 2 月 SaaS 股在 48 小時內蒸發約 2,850 億美元市值；股價 2026 年累計約 −30%。（The Next Web／作者 Cristian Dina；媒體）
2. AI 規模仍小且成長未傳導到合併指標：Agentforce 1.2B ARR（Q1）相對約 46B 營收規模很小；使用量成長快於 cRPO 等訂閱指標，營收連結未證實；Q2 FY27 扣 Informatica 後有機成長約 6–7%（Derived）。（TIKR 引 Guggenheim 對空方的轉述；Futurum Q1 FY27）
3. 利潤品質與資產負債表：Q2 FY27 GAAP 營業利益 2,331M 與去年持平；EPS 大幅超預期係來自策略投資利得 2.61B；非流動債務由 10.4B 升至 39.3B；重組費用升至 94M。（247wallst 2026-09-29，引 Q2 數字；Q2 FY27 新聞稿）
另：消費計價（$0.10／動作、$2／對話）的大企業規模經濟不明；SAP、ServiceNow、Google、Microsoft 同時推代理平台（The Next Web，作者 Dina 的分析）。

## (9) 未找到清單
- 10-K 原檔頁碼：10-K FY2026 本次經第三方鏡像（boardroomalpha）讀取；SEC 原檔頁碼與 10-Q（Q1、Q2 FY27）原檔未能由工具取得，數字取自公司 IR 新聞稿（8-K 附件）與 10-K 鏡像，**待回 SEC 原檔核對**。
- FY2022、FY2023 分雲值：只有第二來源（Stock Analysis on Net），待 10-K FY2023 核對（可能重編）。
- Q3 FY26、Q4 FY26 與 FY2026 全年之**新二分類**（Agentforce Apps／Data 360 等）回溯值：未找到；故無法計算新二分類於 Q3、Q4 的年增。
- 員工數 FY2022–FY2024、以及各季員工數：未找到（只有 FY2025、FY2026）。
- 席位總數、每席位年費、淨留存率、客戶人力替代率：公司未揭露。
- Data 360 ARR 單獨值：公司未揭露。
- AELA 官方一手文件與費率：未找到（只有第二來源）。
- FY2028 營收共識、同日期的賣方平均目標價與家數、FY2029 共識：未找到。
- Q3 FY27 實績（2026-12 公布）：尚未公布。
