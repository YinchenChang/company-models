---
ticker: PLTR
name: Palantir
name_full: Palantir Technologies Inc.
listing: Nasdaq｜PLTR
hq: 美國佛羅里達州邁阿密（2026 年 2 月自丹佛遷出）[外部・未經本專案檢核]
founded: 2003 [外部・未經本專案檢核]
employees: 4,429（2025-12-31；其中 28% 在美國以外）[本專案：facts_palantir.csv 全職員工數]
fiscal_year_end: 12 月底
asof: 2026-11-23
---

## 一句話定位
站在企業與政府資料層之上的「決策平台」：把客戶散在各系統的資料接起來，再讓 AI 代理在上面做事；靠軟體訂閱與按用量、按部署計價賺錢，AI 的爭點是工作量增加帶來的營收，會不會被單價下滑抵銷 [本專案：palantir/README.md 命題；20261117_Palantir個股模型_v0.2.3.xlsx README!A3]。

## 公司在做什麼
- 資料整合：把客戶各系統的資料串成一個共同的「物件模型」（Ontology，用業務語言標出客戶、零件、訂單等物件與關係），再在上面跑分析與應用 [外部・未經本專案檢核]。
- 產品線有 Gotham（偏國防與情報）、Foundry（偏企業資料整合）、Apollo（跨雲端與封閉網路的部署工具）、AIP（把大型語言模型接到客戶資料與流程上）。公司不按產品線揭露營收 [本專案：research_palantir.md 第 2 節]。
- 客戶是美國與盟國政府、大型企業；官方只分 Government 與 Commercial 兩個分部 [本專案：research_palantir.md 第 1、2 節]。
- 導入方式偏重人力：公司派工程師進駐客戶現場，與客戶一起把資料與應用做起來（公司稱 forward deployed engineering） [外部・未經本專案檢核]。
- AI 的角色：模型負責理解與生成，Palantir 提供模型能讀到的資料、權限與可執行的動作。公司沒有量化 AI 營收占比、token 用量或推論成本 [本專案：research_palantir.md 第 5 節]。

## 主要產品與營收占比
| 產品／部門 | 內容（白話） | 營收占比 | 期間 | 標記 |
|---|---|---|---|---|
| Government | 美國與外國政府：國防、情報、民用機關 | 53.7%（US$2,402.3M） | FY2025 | [本專案：research_palantir.md 第 3 節；占比由分部營收÷合併營收換算] |
| Commercial | 企業客戶（美國與國際） | 46.3%（US$2,073.2M） | FY2025 | [本專案：research_palantir.md 第 3 節；同上換算] |
| Government | 同上 | 50.9%（US$2,994.0M） | 2026 前三季 | [本專案：research_palantir.md 第 3 節；同上換算] |
| Commercial | 同上 | 49.1%（US$2,887.5M） | 2026 前三季 | [本專案：research_palantir.md 第 3 節；同上換算] |
| 其中：美國商業 | Commercial 的美國部分 | 42.5%（US$984M） | 2026Q3 | [本專案：facts_palantir.csv 美國商業營收；占比÷合併營收 2,313.4 換算] |
| 其中：國際商業 | Commercial 減美國商業 | 7.9%（US$183.9M） | 2026Q3 | [本專案：facts_palantir.csv 國際商業營收；換算] |

- 2026Q3 營收 US$2,313.4M，較 2025Q3 的 US$1,181.1M 約成長 96% [本專案：facts_palantir.csv 合併營收 2026Q3、2025Q3 分部合計；換算]。美國商業 US$984M 首度高於美國政府 US$953M [本專案：facts_palantir.csv 美國商業營收、美國政府營收 2026Q3]。
- 毛利率 GAAP 為 82.4%（FY2025）、84.8%（2026 前三季）；營業利益率 31.6% 升到 46.8% [本專案：research_palantir.md 第 6 節]。
- 產品線、AI 專屬營收：本專案未建立（公司未揭露）。

## 市場定位與產業位置
- 本專案把它歸在「資料平台、資料庫與可觀測性」次產業，原型為 U（按用量計價）[本專案：palantir/README.md；map/dist/data.json 次產業欄]。
- 護城河來自兩處：長年累積的政府與國防關係，以及客戶資料與流程已經建在其物件模型上，換掉成本高 [外部・未經本專案檢核]。
- 客戶變多、也變大：近 12 個月客戶 1,090 家（2025 年底 954 家）；前 20 大客戶平均營收 US$146.0M（FY2025 為 US$93.9M）；淨美元留存率 165% [本專案：research_palantir.md 第 4 節；公司定義，Interested-party]。
- 政府集中度高：FY2025 前三大客戶占營收 16%；政府約占 2026 前三季營收 51% [本專案：research_palantir.md 第 4、8 節]。
- 訂單指標比營收慢：2026Q3 總合約價值 US$2,918M，年增 6% [本專案：research_palantir.md 第 4 節]。公司未給 FY2027 指引 [本專案：research_palantir.md 第 9 節]。

## 管理層與經營風格
| 職位 | 姓名 | 任職起 | 背景 |
|---|---|---|---|
| 共同創辦人、執行長 | Alexander Karp | 2005（執行長；董事自 2003） | 公司共同創辦人 [外部・未經本專案檢核：Palantir 2026 委託書 DEF 14A] |
| 共同創辦人、總裁 | Stephen Cohen | 未確認（董事自 2005） | 共同創辦人 [外部・未經本專案檢核：Palantir 投資人關係網站，委託書] |
| 共同創辦人、董事長 | Peter Thiel | 董事自 2003 | 共同創辦人 [外部・未經本專案檢核：Palantir 2026 委託書 DEF 14A] |
| 技術長 | Shyam Sankar | 未確認 | 執行副總裁 [外部・未經本專案檢核：Palantir 投資人關係網站] |
| 財務長 | David Glazer | 未確認 | 兼任 Treasurer [外部・未經本專案檢核：Palantir 投資人關係網站] |
| 營收長兼法務長 | Ryan Taylor | 未確認 | 一人兼兩職 [外部・未經本專案檢核：Palantir 投資人關係網站] |

- 經營風格：本專案尚未研究。
- 本專案記錄到的資本配置：2026-09-30 無借款，現金及約當現金 US$3,619.4M、短期美國公債 US$6,942.1M；資本支出很小（2026 前三季 US$30.0M） [本專案：facts_palantir.csv 對應各列]。
- 股票薪酬占營收 13.4%（2026 前三季），FY2025 為 15.3% [本專案：research_palantir.md 第 6 節]。
- 股權結構：創辦人透過 Class F 股票與投票信託，可掌握最高 49.999999% 的表決權 [外部・未經本專案檢核：Palantir 2026 委託書 DEF 14A]。
- 近期人事：2026-07 起未查到高階主管異動 [外部・未經本專案檢核：Q3 財報新聞稿未揭露異動]。

## 同業與競爭態勢
| 同業 | 重疊領域 | 市占或規模 | 標記 |
|---|---|---|---|
| Snowflake、Datadog、MongoDB、Cloudflare | 資料平台與雲端用量計價軟體（本專案同業集合） | 本專案未建立 | [本專案：palantir/README.md 方法；research_palantir.md 第 8 節] |
| ServiceNow | 企業工作流程與 AI 代理平台 | 本專案未建立 | [本專案：research_palantir.md 第 8 節同業清單] |
| Booz Allen Hamilton、Leidos、CACI | 政府資訊服務與國防軟體整合 | 本專案未建立 | [本專案：research_palantir.md 第 8 節同業清單] |
| Databricks、Anduril | 資料與 AI 平台；國防軟體與系統 | 本專案未建立 | [外部・未經本專案檢核] |

- 本專案挑同業是為了估值對照，彼此業務並不完全相同；政府同業（IT 服務型）的商業模式與 Palantir 差異大，模型只把它們當對照 [本專案：20261117_Palantir個股模型_v0.2.3.xlsx README!A23]。
- 威脅來自兩邊：模型廠商自己做企業代理；客戶改用開源權重模型自建，單位價格被壓低 [本專案：20261117_Palantir個股模型_v0.2.3.xlsx README!A44]。
- 國防端的新進者與傳統承包商也在爭同一批預算 [外部・未經本專案檢核]。

## 主要客戶
- 財報揭露的集中度：前三大客戶占 FY2025 營收 16%（FY2024 為 17%），平均合作約十年；公司不具名 [外部・未經本專案檢核：Palantir FY2025 10-K（SEC），2026；本專案研究包有同一數字]。
- 地區：FY2025 營收 74% 來自美國客戶、26% 來自海外 [外部・未經本專案檢核：Palantir FY2025 10-K（SEC），2026]。
- 公開已知的客戶或合作對象（皆非 10-K 揭露的集中度客戶）：
| 客戶 | 關係 | 依據 | 標記 |
|---|---|---|---|
| 美國陸軍 | 軍方客戶；2026-09 起生產 8 套 TITAN 地面目標站（Palantir 為主承包商） | 新聞報導，未載金額 | [外部・未經本專案檢核：The Defense Post，2026-09-03] |
| 英國國防部 | 續約三年資料分析授權與支援，金額 £240.6M | 新聞報導，採免競標程序 | [外部・未經本專案檢核：TechRadar，2026-01-29] |
| NVIDIA | 用 Palantir 平台管理自家供應鏈（約 130 萬個零件／每座 Vera Rubin 機櫃） | 雙方公告 | [外部・未經本專案檢核：Stocktwits 轉述，2026-09-10] |
| L3Harris、McCarthy Building、Tampa General Hospital | 用 AIP 搭配開源權重模型做自有模型 | 公司新聞稿 | [外部・未經本專案檢核：Business Wire 經 Finviz，2026-10-21] |
- 未找到：各客戶個別營收占比（10-K 摘錄只給前三大合計）；查過 FY2025 10-K 前 10 萬字元、Q3 財報新聞稿（未點名客戶）。

## 前十大股東
| 排名 | 股東 | 持股比例 | 標記 |
|---|---|---|---|
| 1 | BlackRock Inc. | 8.70% | [外部・未經本專案檢核：Yahoo Finance 機構持股，2026-06-30] |
| 2 | Vanguard Capital Management LLC | 6.70% | [外部・未經本專案檢核：Yahoo Finance 機構持股，2026-06-30] |
| 3 | State Street Corporation | 4.79% | [外部・未經本專案檢核：Yahoo Finance 機構持股，2026-06-30] |
| 4 | Invesco Ltd. | 3.23% | [外部・未經本專案檢核：Yahoo Finance 機構持股，2026-06-30] |
| 5 | Geode Capital Management, LLC | 2.54% | [外部・未經本專案檢核：Yahoo Finance 機構持股，2026-06-30] |
| 6 | Vanguard Portfolio Management LLC | 2.31% | [外部・未經本專案檢核：Yahoo Finance 機構持股，2026-06-30] |
| 7 | Morgan Stanley | 1.56% | [外部・未經本專案檢核：Yahoo Finance 機構持股，2026-06-30] |
| 8 | Norges Bank | 1.43% | [外部・未經本專案檢核：Yahoo Finance 機構持股，2026-06-30] |
| 9 | JPMorgan Chase & Co | 1.35% | [外部・未經本專案檢核：Yahoo Finance 機構持股，2026-06-30] |
| 10 | Goldman Sachs Group Inc | 1.04% | [外部・未經本專案檢核：Yahoo Finance 機構持股，2026-06-30] |

- 口徑：13F 機構持股（占流通股比例，同一集團下不同實體分列）；資料日期 2026-06-30。此表只含機構，不含個人。
- 補充：共同創辦人 Peter Thiel 持股約 3.00%（TIKR 頁面 2026-06-01 更新，未載資料日）[外部・未經本專案檢核：TIKR，2026-06-01]。
- 2026-10-26 流通股：Class A 約 2,306.7M 股、Class B 約 101.8M 股、Class F 約 1.0M 股 [本專案：facts_palantir.csv 各類流通股]。

## 近期重大事件
- 2026-11-03｜2026Q3 財報：營收 US$2,313.4M，美國商業首度超過美國政府；FY2026 營收指引上修到 US$8,530–8,534M，2026Q4 指引 US$2,649–2,653M；訂單指標（總合約價值年增 6%）成長慢於營收 [本專案：research_palantir.md 第 3、4 節；facts_palantir.csv 指引列]。同日股價收跌約 1.6% [外部新聞：Webull]
- 2026-10-21｜公司宣布一批「主權 AI」客戶（L3Harris、McCarthy、Tampa General 等）用 Foundry、AIP 加開源模型自建 AI，是模型單價下滑風險的實例，也是用量增加的來源 [外部新聞：Business Wire 經 Finviz]
- 2026-10-19｜與 Knox Systems 合作，借其聯邦受管雲端，加快政府客戶取得 FedRAMP 與 IL5 授權的部署 [外部新聞：Business Wire 經 Finviz]
- 2026-09-10｜與 NVIDIA 合作，結合 Nemotron 開源模型與 Palantir 平台；同日與富士通續約並擴大合作（富士通成為全球部署工程夥伴） [外部新聞：Stocktwits 轉述；Finviz]
- 2026-09-03｜美國陸軍授予 8 套 TITAN 地面目標站生產訂單，使該計畫由原型轉入實際部署 [外部新聞：The Defense Post]
- 2026-08-03｜2026Q2 財報：營收 US$1.94B，年增 93%；FY2026 營收指引上修到 US$8.150–8.158B，股價盤後上漲約 13% [外部新聞：Fortune]

## 需要觀察的事項
- 美國商業每客戶營收與留存率｜模型最承重的假設是每客戶工作量成長（FY2027 三情境 42%／32%／20%）；公司揭露客戶層級資料即可驗證 [本專案：20261117_Palantir個股模型_v0.2.3.xlsx README!A58]｜2026Q4 財報（日期未公告，預期 2027 年初）
- 單價下滑｜單價變動假設 FY2027 為 -5%／-8%／-13%；情境 B 下 FY2028 AI 淨效果約 US$1.2B，工作負載成長低於 7.3% 就轉負 [本專案：20261117_Palantir個股模型_v0.2.3.xlsx README!A15、A58、A63]｜新簽合約單位價格與 Commercial 毛利
- 推論成本｜模型推得約占 Commercial 營收 2.2%，但用量為判斷值，公司未揭露 [本專案：20261117_Palantir個股模型_v0.2.3.xlsx README!A15]｜公司若揭露 AIP 用量或推論成本
- 政府營收與續約｜政府約占 FY2026 營收一半，模型把美國政府營收一次性減少只放在敏感度 [本專案：20261117_Palantir個股模型_v0.2.3.xlsx README!A43]｜政府分部合約價值與續約
- 訂單領先指標｜總剩餘交易價值較 Q2 僅增 2%、剩餘履約義務與上季持平的說法來自賣方 [本專案：research_palantir.md 第 8 節；Interested-party]｜2026Q4 財報
- 股票薪酬與稀釋｜季度股票薪酬由 US$172.3M 升到 US$318.9M；模型股數固定於評價日 [本專案：facts_palantir.csv 股票薪酬（合計）；README!A51]｜2026Q4 10-K

## 來源
1. 本專案：`software_ai/palantir/data/research_palantir.md`；`software_ai/palantir/data/facts_palantir.csv`；`software_ai/palantir/README.md`；`software_ai/palantir/dist/20261117_Palantir個股模型_v0.2.3.xlsx`（README 頁）；`software_ai/map/dist/data.json`
2. [Palantir 2026Q3 財報新聞稿（8-K 附件 99.1）](https://www.sec.gov/Archives/edgar/data/0001321655/000132165526000045/a2026q3ex991pressrelease.htm)（2026-11-03）
3. [Palantir FY2025 Form 10-K](https://investors.palantir.com/files/2025%20FY%20PLTR%2010-K.pdf)（2026-02）
4. [Palantir 2026 委託書 DEF 14A](https://www.sec.gov/Archives/edgar/data/0001321655/000132165526000019/pltr-20260423.htm)（2026-04）
5. [Palantir 投資人關係：Executive Management](https://investors.palantir.com/management.html)（2026-11 查閱）
6. [Yahoo Finance：PLTR 機構持股](https://finance.yahoo.com/quote/PLTR/holders)（資料日 2026-06-30）
7. [TIKR：Who Owns Palantir?](https://tikr.com/blog/who-owns-palantir-top-shareholders-and-recent-insider-trades)（2026-06-01）
8. [Fortune：Palantir 2Q26 財報](https://fortune.com/2026/08/03/palantir-earnings-guidance-beat-revenue-profit-ai-demand/)（2026-08-03）
9. [Webull：Palantir Q3 後股價](https://www.webull.com/news/15862517031134208)（2026-11）
10. [Business Wire 經 Finviz：Palantir accelerates delivery of sovereign AI](https://finviz.com/news/405850/palantir-accelerates-delivery-of-sovereign-ai-to-new-wave-of-customers)（2026-10-21）
11. [Business Wire 經 Finviz：Knox and Palantir partner](https://finviz.com/news/403839/knox-and-palantir-partner-to-further-accelerate-delivery-of-software-through-largest-federal-cloud)（2026-10-19）
12. [Stocktwits：Palantir builds out AI ecosystem with NVIDIA, Fujitsu, Method Security](https://stocktwits.com/news-articles/markets/equity/palantir-builds-out-ai-ecosystem-with-nvidia-fujitsu-method-security/cZtaIH2RJCV)（2026-09-10）
13. [Finviz：Palantir and Fujitsu renew AI partnership](https://finviz.com/news/390878/palantir-and-fujitsu-renew-ai-partnership-and-expand-forward-deployed-engineering-role)（2026-09-10）
14. [The Defense Post：Palantir to build eight TITAN targeting stations](https://thedefensepost.com/2026/09/03/palantir-targeting-stations-us-army/)（2026-09-03）
15. [TechRadar：Palantir UK MoD 三年續約](https://www.techradar.com/pro/security/palantir-to-continue-uk-ministry-of-defense-work-in-new-three-year-deal-gbp240-million-data-analytics-contract-awarded-without-procurement)（2026-01-29）
16. [Colorado Sun：Palantir 遷往邁阿密](https://coloradosun.com/2026/02/17/palantir-leaving-denver-moving-miami/)（2026-02-17）
