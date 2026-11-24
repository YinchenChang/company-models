---
ticker: CRWD
name: CrowdStrike
name_full: CrowdStrike Holdings, Inc.
listing: Nasdaq｜CRWD
hq: 美國德州奧斯汀 [外部・未經本專案檢核：CrowdStrike 10-K FY2026 封面，2026-03]
founded: 2011 [外部・未經本專案檢核：CrowdStrike 10-K FY2026 Item 1，2026-03]
employees: 10,698（2026-01-31）；11,706（2026-07-31）[本專案：crowdstrike/data/facts_crowdstrike.csv 全職員工數]
fiscal_year_end: 1 月底
asof: 2026-11-23
---

## 一句話定位
位在企業軟體的「資安」層：以 Falcon 雲端平台按防護範圍（端點、身分、模組數）收年費訂閱，爭點是 AI 代理（非人類身分）會不會成為新的計價單位而帶來增量，還是 AI 平台業者自建防護、把每模組價格壓低 [本專案：crowdstrike/README.md 命題；data/research_crowdstrike.md (8)]。

## 公司在做什麼
- 賣雲端原生的資安平台 Falcon：從電腦與伺服器上的端點防護（偵測並阻擋惡意程式與入侵）起家，再擴到身分防護、雲端與資料等模組 [本專案：data/research_crowdstrike.md (1)]。
- 客戶按受保護的端點、身分數量和採用的模組數計費；公司也提供 Falcon Flex（先買一筆額度，再自由挑模組使用的平台授權） [本專案：data/research_crowdstrike.md (1)(4)]。
- 10-K 列出 32 個雲端模組（年底），申報日為 33 個；新增的 AI 相關產品有 Charlotte AI（生成式 AI 分析助理）、AIDR（偵測並回應 AI 代理風險）與 Falcon Shield（含非人類身分） [本專案：data/research_crowdstrike.md (2)]。
- AI 為什麼影響它：AI 代理會自行登入與操作系統，需要被辨識與管控，攻擊者也用 AI 放大攻擊面；但公司不揭露 AI 專屬營收 [本專案：data/research_crowdstrike.md (5)(9)]。

## 主要產品與營收占比
| 產品／部門 | 內容（白話） | 營收占比 | 期間 | 標記 |
|---|---|---|---|---|
| 訂閱 | Falcon 平台各模組的年度授權，依合約期間按比例認列 | 94.9%（US$4,564.7M） | FY2026（2025-02～2026-01） | [本專案：data/research_crowdstrike.md (3)；占比＝4,564.7÷4,812.0，換算] |
| 專業服務 | 事件應變、顧問與導入服務 | 5.1%（US$247.3M） | FY2026 | [本專案：data/research_crowdstrike.md (3)；換算] |
| 合計 | 單一應報導分部 | 100%（US$4,812.0M） | FY2026 | [本專案：data/research_crowdstrike.md (2)(3)] |

- 公司不揭露各模組營收，也不揭露 AI 專屬營收 [本專案：data/research_crowdstrike.md (2)(9)]。
- 期末 ARR（年化訂閱合約值）：FY2026 US$5.25B，Q2 FY2027（2026-07-31）US$5.84B，年增 25% [本專案：data/research_crowdstrike.md (4)]。
- 毛利率：GAAP 總毛利率 74.7%（FY2026）、訂閱 78%；GAAP 營業利益率 −6.1%，non-GAAP 22% [本專案：data/research_crowdstrike.md (6)]。
- 地區（FY2026）：美國 66.8%、EMEA 16.3%、亞太 10.3%、其他 6.6% [本專案：data/facts_crowdstrike.csv 營收地區四列；占比換算]。

## 市場定位與產業位置
- 本專案產業地位分數 93.8（資安次產業，與 Palo Alto Networks 並列最高；其餘 Cloudflare、Fortinet、Okta、Zscaler 為 77）[本專案：map/dist/data.json 資安次產業 companies]。
- 平台整併是主軸：採用 6 個以上模組的客戶占 51%，8 個以上占 26%（Q2 FY2027）；Falcon Flex 帳戶 ARR US$2.29B，約占 ARR 39% [本專案：data/research_crowdstrike.md (4)]。
- 客戶黏性：美元淨留存率 FY2026 為 115%（FY2025 為 112%）；RPO（已簽約未認列營收）US$10.7B [本專案：data/research_crowdstrike.md (4)]。
- 定價單位：本專案模型用「受保護單位 × 每單位模組數 × 每模組年費」推 ARR，三者都是 Assumed，是承重輸入 [本專案：crowdstrike/README.md 命題與已知限制]。
- 客戶集中度：公司稱業務不依賴任何單一終端客戶 [外部・未經本專案檢核：CrowdStrike 10-K FY2026 Customers 節，2026-03]。

## 管理層與經營風格
| 職位 | 姓名 | 任職起 | 背景 |
|---|---|---|---|
| 共同創辦人、執行長 | George Kurtz | 創辦 2011 | 10-K 載明為董事兼執行長（55 歲，截至 2026-03-04）[外部・未經本專案檢核] |
| 總裁 | Michael Sentonas | 2023-02 升任總裁 | 由原職位升任總裁 [外部・未經本專案檢核：Business Wire，2023-02-08] |
| 財務長 | Burt Podbere | 2015 加入公司 | 曾任 OpenDNS、Net Optics 財務長 [外部・未經本專案檢核：公司官網高階主管頁] |

- 經營風格：本專案尚未研究。
- 本專案記錄到的資本配置：FY2027 上半年回購 US$175.6M，回購授權尚餘 US$1.3B（總額 1.5B）[本專案：data/facts_crowdstrike.csv 股票回購、回購授權剩餘]。
- 併購持續：FY2026 併購現金支出 US$382.3M，FY2027 上半年 US$881.4M；標的偏 AI 資料管線與身分（Onum、Pangea、Seraphic、SGNL.AI；XM Cyber 技術資產 2026-07 簽約、未交割）[本專案：data/facts_crowdstrike.csv 併購列；data/research_crowdstrike.md (5)]。
- 股票薪酬約占營收 23%，是 GAAP 仍虧損的主因 [本專案：data/research_crowdstrike.md (6)(8)]。
- 治理與近期人事：2026-11-02 聘任 Roland Stritt 主管歐洲聯盟夥伴成長 [外部新聞：CrowdStrike 新聞稿（IR）]。

## 同業與競爭態勢
| 同業 | 重疊領域 | 市占或規模 | 標記 |
|---|---|---|---|
| Palo Alto Networks | 資安平台（防火牆、雲端、SOC）；同走平台整併，併購 CyberArk | 本專案未建立（地位分數同為 93.8） | [本專案：data/research_crowdstrike.md (8)；map/dist/data.json] |
| SentinelOne | 端點防護 | 本專案未建立（僅作同業 P/S 比較） | [本專案：crowdstrike/README.md v0.2.1] |
| Zscaler、Okta、Cloudflare、Fortinet | 零信任網路、身分、網路安全 | 本專案未建立（地位分數 77） | [本專案：map/dist/data.json] |
| Microsoft | 端點與身分防護，隨作業系統與辦公套件打包 | 本專案未建立 | [外部・未經本專案檢核] |
| AI 平台業者（Anthropic 等） | 以模型直接做程式碼與弱點掃描安全功能 | 本專案未建立 | [本專案：data/research_crowdstrike.md (5)(8)] |

- 2026-02-20 Anthropic 發表 Claude Code Security 後，CRWD 兩日跌近 20%，市場擔心 AI 平台能自建防護 [本專案：data/research_crowdstrike.md (8)]。
- 另一面，公司把 Claude、OpenAI 都當通路與合作夥伴 [外部新聞：CrowdStrike 新聞稿（IR）]；本專案研究包記錄公司稱 AI 擴大攻擊面是需求來源 [本專案：data/research_crowdstrike.md (8)]。
- 現價 P/S 約為同業中位數的 4.4 倍，代表市場給的是平台溢價 [本專案：crowdstrike/dist/20261124_CrowdStrike個股模型_v0.2.2.xlsx Val!B19]。

## 主要客戶
- 財報揭露的客戶集中度：10-K 稱「業務不依賴任何特定終端客戶」，未揭露占營收 10% 以上的客戶；客戶逾 88,000 家（含託管資安服務商〔MSSP〕的終端客戶）[外部・未經本專案檢核：CrowdStrike 10-K FY2026 Customers 節，2026-03]。
| 客戶 | 關係 | 依據 | 標記 |
|---|---|---|---|
| EY（安永） | 以 Falcon 平台建立其前沿 AI 與資安韌性業務，並承諾培訓人才 | 公司新聞稿 | [外部・未經本專案檢核：CrowdStrike 新聞稿，2026-11-03] |
| Delta Air Lines | 7 月 19 日事件相關訴訟的對手，案件仍在證據開示 | 本專案研究包 | [本專案：data/research_crowdstrike.md (8)] |
- 未找到：各客戶個別占營收比例（10-K 未揭露）。查過：10-K FY2026 Customers 節、公司官網新聞稿。

## 前十大股東
| 排名 | 股東 | 持股比例 | 標記 |
|---|---|---|---|
| 1 | Vanguard Capital Management | 7.27% | [外部・未經本專案檢核：CrowdStrike 2026 Proxy Statement（DEF 14A），資料日期 2026-04-03] |
| 2 | BlackRock, Inc. | 6.70% | [外部・未經本專案檢核：CrowdStrike 2026 Proxy Statement（DEF 14A），資料日期 2026-04-03] |
| 3 | State Street Corporation | 1.13% | [外部・未經本專案檢核：Yahoo Finance 持股頁（13F 彙整），資料日期 2026-06-30] |
| 4 | Invesco Ltd. | 0.82% | [外部・未經本專案檢核：Yahoo Finance 持股頁（13F 彙整），資料日期 2026-06-30] |
| 5 | JPMorgan Chase & Co. | 0.80% | [外部・未經本專案檢核：Yahoo Finance 持股頁（13F 彙整），資料日期 2026-06-30] |
| 6 | FMR, LLC | 0.76% | [外部・未經本專案檢核：Yahoo Finance 持股頁（13F 彙整），資料日期 2026-06-30] |
| 7 | Morgan Stanley | 0.63% | [外部・未經本專案檢核：Yahoo Finance 持股頁（13F 彙整），資料日期 2026-06-30] |
| 8 | Geode Capital Management | 0.59% | [外部・未經本專案檢核：Yahoo Finance 持股頁（13F 彙整），資料日期 2026-06-30] |
| 9 | Vanguard Portfolio Management | 0.58% | [外部・未經本專案檢核：Yahoo Finance 持股頁（13F 彙整），資料日期 2026-06-30] |
| 10 | T. Rowe Price Associates | 0.42% | [外部・未經本專案檢核：Yahoo Finance 持股頁（13F 彙整），資料日期 2026-06-30] |

- 口徑：第 1–2 名取自 2026 年委託書的 5% 以上持有人表（拆股前股數，總股數 254,580,533 股，資料日 2026-04-03）；第 3–10 名取自 Yahoo Finance 的 13F 彙整（拆股後，資料日 2026-06-30）。兩表日期與口徑不同，且 Vanguard 自 2026-01 起由旗下子公司分別申報，排名僅供參考，未逐筆核對 [外部・未經本專案檢核：CrowdStrike 2026 DEF 14A；Vanguard Schedule 13G/A，2026-03-26；Yahoo Finance，2026-06-30]。
- 創辦人 Kurtz 持 2,384,082 股，低於 1%；全體董事與高階主管 11 人合計 1.67% [外部・未經本專案檢核：CrowdStrike 2026 DEF 14A，資料日期 2026-04-03]。

## 近期重大事件
- 2026-11-03｜宣布收購 Dreadnode（AI 資安自動化測試）與 Fiddler AI（AI 代理可觀測性），以現金加股票支付，預計 FY2027 第 4 季交割；同日推出 AI Kill Switch、Falcon Guardian 對接 Claude Cowork、AWS 原生 AI 安全功能。重要性：持續補 AI 代理防護的產品線 [外部新聞：RTTNews；CrowdStrike 新聞稿（IR）]
- 2026-11-02｜與 CoreWeave 合作保護 AI 工作負載；10-21 與 NVIDIA 擴大 AI 代理安全合作。重要性：綁定 AI 基礎設施通路 [外部新聞：CrowdStrike 新聞稿（IR）]
- 2026-09-29｜Falcon 上架 OpenAI Marketplace，OpenAI 企業客戶可用既有額度購買。重要性：模型商成為通路 [外部新聞：CrowdStrike 新聞稿]
- 2026-09-02｜Falcon 上架 Anthropic Claude Marketplace；擴大與 OpenAI 合作；發表 Agentic Identity Provider（AI 代理身分）。重要性：同一天把兩家模型商納為夥伴 [外部新聞：CrowdStrike 新聞稿（IR）]
- 2026-08-26｜Q2 FY2027 財報：營收 US$1,470.9M、期末 ARR US$5.84B、淨新增 ARR US$332.8M；FY2027 淨新增 ARR 指引上調至約 US$1.35B；首度說明 AIDR 含 token 計價 [本專案：data/research_crowdstrike.md (3)(4)(5)(8)]
- 2026-07-02｜4 比 1 拆股生效，此後每股數字為拆股後 [外部新聞：Motley Fool；本專案：data/research_crowdstrike.md 開頭說明]
- 2026-07｜簽約取得 XM Cyber 技術資產（約 US$145M，未交割） [本專案：data/facts_crowdstrike.csv 併購：XM Cyber；data/research_crowdstrike.md (5)]

## 需要觀察的事項
- 價格保留與 AI 代理單位數｜本專案兩項承重 Assumed（ero、pf 與代理單位比），決定 AI 淨效果是正是負 [本專案：crowdstrike/README.md 承重 Assumed 2；Thesis 頁]｜公司揭露 AIDR 絕對 ARR 時
- 美元淨留存率與 Flex 轉換增幅｜用來檢驗定價是否受壓；FY2026 為 115%，Flex 轉換平均 ARR 增逾 40% [本專案：data/research_crowdstrike.md (4)(5)]｜Q3 FY2027 財報（預計 2026 年 12 月初）
- ARR 軌跡｜公司 FY2027 期末 ARR 指引 6,603.0M–6,611.9M，模型為 6,685M [本專案：data/facts_crowdstrike.csv FY2027 期末 ARR 指引；crowdstrike/README.md 承重 Assumed 1]｜Q3 FY2027 財報
- 推論成本｜模型情境 B 的增量推論成本約占 non-GAAP 營業成本 6.9%，不是主要風險 [本專案：crowdstrike/dist/20261124_CrowdStrike個股模型_v0.2.2.xlsx README!A7]｜Tokenomics 快照更新時
- SBC 與 GAAP 利益率｜SBC 約占營收 23%，GAAP 營業利益率仍為負，是 DCF 與 non-GAAP 口徑差距主因 [本專案：data/research_crowdstrike.md (6)；README!A8]｜Q3 FY2027 財報
- 法律與資料品質｜DOJ／SEC 就營收認列與 ARR 報告的資訊請求未有結論；Delta 案在證據開示 [本專案：crowdstrike/README.md 風險登錄]｜10-Q／8-K 披露

## 來源
1. 本專案：`software_ai/crowdstrike/data/research_crowdstrike.md`；`data/facts_crowdstrike.csv`；`README.md`；`dist/20261124_CrowdStrike個股模型_v0.2.2.xlsx`（README、Val）；`software_ai/map/dist/data.json`
2. [CrowdStrike Form 10-K FY2026](https://www.sec.gov/Archives/edgar/data/1535527/000153552726000010/crwd-20260131.htm)（2026-03）
3. [CrowdStrike 2026 Proxy Statement（DEF 14A）](https://www.sec.gov/Archives/edgar/data/1535527/000110465926055599/tm2532333-4_def14a.htm)（2026-04）
4. [Vanguard Group Schedule 13G/A](https://www.sec.gov/Archives/edgar/data/0001535527/000010290926001050/primary_doc.xml)（2026-03-26）
5. [Yahoo Finance：CRWD 持股頁](https://finance.yahoo.com/quote/CRWD/holders)（擷取 2026-11-23，資料日 2026-06-30）
6. [CrowdStrike 高階主管頁](https://www.crowdstrike.com/about-us/executive-team/)（2026-11）；[Burt Podbere](https://crowdstrike.com/en-us/about-us/executive-team/burt-podbere)
7. [CrowdStrike Promotes Michael Sentonas to President](https://www.businesswire.com/news/home/20230208005405/en/CrowdStrike-Promotes-Michael-Sentonas-to-President)（2023-02-08）
8. [CrowdStrike 新聞稿列表（IR）](https://ir.crowdstrike.com/press-releases)（2026-09 至 2026-11）
9. [CrowdStrike to Buy Dreadnode, Fiddler AI（RTTNews）](https://www.rttnews.com/3707615/crowdstrike-to-buy-dreadnode-fiddler-ai-to-boost-ai-trust.aspx?type=bn)（2026-11-03）
10. [CrowdStrike's 4-for-1 stock split is coming July 2（Motley Fool）](https://www.fool.com/investing/2026/06/17/crowdstrikes-4-for-1-stock-split-is-coming-july-2/)（2026-06-17）
11. [CrowdStrike Brings the Falcon Platform to the OpenAI Marketplace](https://www.crowdstrike.com/en-us/press-releases/crowdstrike-brings-the-falcon-platform-to-the-openai-marketplace/)（2026-09-29）
