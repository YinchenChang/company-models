---
ticker: SNOW
name: Snowflake
name_full: Snowflake Inc.
listing: NYSE｜SNOW
hq: 美國加州門洛帕克（10-K 封面主要辦公室地址）[外部・未經本專案檢核]
founded: 2012 [外部・未經本專案檢核]
employees: 9,060（FY2026 期末）[本專案：facts_snowflake.csv 員工數（期末）FY2026]
fiscal_year_end: 1 月底
asof: 2026-11-23
---

## 一句話定位
資料層的雲端平台：把企業資料集中放在它的資料倉儲上，依實際用掉的運算、儲存與傳輸量計價；AI 功能掛在同一個計價單位上，爭點是 AI 帶來的新用量，能否大於推論成本對毛利的侵蝕，以及模型直接讀資料對核心消耗的替代 [本專案：research_snowflake.md 第 1 節、第 8 節；snowflake/README.md 命題]。

## 公司在做什麼
- 出租雲端資料平台（資料倉儲：把各部門的資料集中存放、用 SQL 查詢分析的系統）。客戶不買席位，用多少付多少 [本專案：research_snowflake.md 第 1 節]。
- 平台跑在 AWS、Azure、GCP 等第三方雲端上，自己不蓋資料中心；第三方雲端費用占產品營業成本 71%（FY2026） [本專案：research_snowflake.md 第 4.1 節]。
- AI 產品包括 Cortex（在資料旁直接呼叫大型語言模型）、CoCo（Cortex Code，寫程式的 AI 助手）、CoWork，以及 Snowflake Intelligence。計價改用獨立的「AI Credit」，與一般運算點數分開 [本專案：research_snowflake.md 第 5 節]。
- 2026-02 收購 Observe（可觀測性：蒐集與分析系統運作紀錄的工具），2026-06 收購 Natoma（讓 AI 代理安全連到各種系統的 MCP 平台） [本專案：research_snowflake.md 第 5 節]。
- 對讀者的一般說明：資料倉儲是 AI 代理取得企業資料的來源，代理多跑一次查詢，就多一筆消耗 [外部・未經本專案檢核]。

## 主要產品與營收占比
| 產品／部門 | 內容（白話） | 營收占比 | 期間 | 標記 |
|---|---|---|---|---|
| 產品營收 Product | 客戶使用平台的運算、儲存、資料傳輸，依消耗量認列；AI 功能營收也在此，不另列 | 95.5%（US$4,472.3M） | FY2026 | [本專案：research_snowflake.md 第 3.1 節；占比＝4,472.317÷4,683.946，換算] |
| 專業服務及其他 | 導入、顧問與訓練服務 | 4.5%（US$211.6M） | FY2026 | [本專案：research_snowflake.md 第 3.1 節；換算] |
| 合計 | 單一營運與應報導分部 | US$4,683.9M | FY2026 | [本專案：research_snowflake.md 第 2 節、第 3.1 節] |

- 地區別（FY2026）：美國 75.2%、EMEA 16.3%、亞太及日本 5.8%、其他美洲 2.7% [本專案：research_snowflake.md 第 3.3 節；占比由金額換算]。
- 產品營收年增：FY2026 29.2%；FY2027 Q2 37%，連續三季加速（30%→34%→37%） [本專案：research_snowflake.md 第 3.1 至 3.2 節]。
- AI 營收只在 FY2026 Q3 揭露過一次：年化 run-rate 約 US$100M，口徑未公開；之後未再給金額，AI 占營收比例與 AI 毛利率本專案未建立 [本專案：research_snowflake.md 第 5 節、第 9 節]。
- 毛利率（GAAP）：FY2026 產品 71.8%、合併 67.2%；FY2027 Q2 產品 70.9%（去年同期 72.3%） [本專案：research_snowflake.md 第 6 節、第 8 節]。

## 市場定位與產業位置
- 本專案把它放在「資料平台、資料庫與可觀測性」次產業，地位分數 77 [本專案：map/dist/data.json SNOW score]。全球資料倉儲市占：本專案未建立。
- 規模與黏著度（Q2 FY27）：淨營收留存率 126%；過去 12 個月產品營收逾 US$1M 的客戶 828 家，占產品營收 68%；Forbes Global 2000 客戶 829 家；容量合約客戶占營收 97% [本專案：research_snowflake.md 第 4.2 節]。
- 訂單能見度：RPO（已簽約但尚未認列的營收）US$9,004.3M，年增 30%，低於 Q4 FY26 的 42% [本專案：research_snowflake.md 第 4.2 節]。
- 對上游依賴：與 AWS 的最低消費承諾累計 US$60 億，至 2031-03；10-Q 點名 Anthropic、OpenAI 對價格與條款有議價力，並造成 AI 產品毛利壓縮 [本專案：research_snowflake.md 第 4.3 節]。
- 定價權：有，但 AI 部分較弱。管理層稱 AI 工作負載貢獻毛利較低，FY2027 非 GAAP 產品毛利率指引由 75.0% 降到 74.0% [本專案：research_snowflake.md 第 5 節]。

## 管理層與經營風格
| 職位 | 姓名 | 任職起 | 背景 |
|---|---|---|---|
| 執行長 | Sridhar Ramaswamy | 2024-02 | 原 Neeva 執行長，2023 年隨 Neeva 被收購加入；曾任職 Google 廣告與工程部門 [外部・未經本專案檢核：Snowflake 2026 Annual Report and Proxy Statement] |
| 財務長 | Brian Robins | 2025-09 | 接替 Michael Scarpelli；曾任 GitLab 財務長 [外部・未經本專案檢核：Snowflake 2026 Annual Report and Proxy Statement] |
| 董事長 | Frank Slootman | 董事長 2019-12；2019-04 至 2024-02 任執行長 | [外部・未經本專案檢核：Snowflake 2026 Annual Report and Proxy Statement] |
| 共同創辦人、董事 | Benoit Dageville | 董事自 2011 | MarketScreener 列為董事 [外部・未經本專案檢核：MarketScreener，2026-09 快照] |

- 經營風格：本專案尚未研究。
- 本專案記錄到：不配息（期間股利 0）；有 2,300 百萬美元可轉債，轉換價 157.50 美元，並有 capped call [本專案：20261124_Snowflake個股模型_v0.2.3.xlsx Val!B35；snowflake/README.md v0.2 修正內容第 4 點]。
- 股票薪酬偏重：占營收 34.1%（FY2026），FY2027 Q2 降到 27.4% [本專案：research_snowflake.md 第 6 節]。
- 管理層稱 AI 降低對人力成長的依賴：今年上半年新增員工 334 人（含 Observe 173 人），去年同期 935 人 [本專案：research_snowflake.md 第 6 節、第 8 節]。

## 同業與競爭態勢
| 同業 | 重疊領域 | 市占或規模 | 標記 |
|---|---|---|---|
| Databricks | 資料與 AI 平台，爭取 AI 工作負載 | 本專案未建立 | [本專案：research_snowflake.md 第 8 節空方論述提及] |
| AWS、Azure、GCP 自家資料服務 | 資料倉儲；同時也是 Snowflake 的底層雲端 | 本專案未建立 | [外部・未經本專案檢核] |
| MongoDB、Datadog、Elastic 等 | 資料庫與可觀測性（Snowflake 收購 Observe 後與 Datadog 重疊） | 本專案未建立 | [本專案：20261124_Snowflake個股模型_v0.2.3.xlsx Val!A64:A70 同業清單] |

- 空方論點是 AI 代理可能繞過資料倉儲、Databricks 追趕 AI 工作負載；多方論點是代理工作負載直接變成消耗，不必重新談約 [本專案：research_snowflake.md 第 8 節]。
- 本專案模型把「核心消耗被模型直讀資料或開放格式（如 Iceberg）替代」列為要建模的風險 [本專案：20261124_Snowflake個股模型_v0.2.3.xlsx README!A25]。
- 同業現價 P/S 中位數 9.36 倍，Snowflake 自身 18.87 倍 [本專案：20261124_Snowflake個股模型_v0.2.3.xlsx Val!B84、Val!B85]。

## 主要客戶
- 客戶集中度：10-K FY2026 是否有「單一客戶占營收 10% 以上」的揭露：未找到（查過 10-K FY2026 前 76 萬字元的業務、MD&A、附註 7 至 12，未見相關句）。 [外部・未經本專案檢核：SEC 10-K FY2026，2026-03]
- 本專案記錄的分散度指標：總客戶 14,554（CEO 口述）；逾 US$1M 客戶 828 家 [本專案：research_snowflake.md 第 4.2 節]。
| 客戶 | 關係 | 依據 | 標記 |
|---|---|---|---|
| 具名主要客戶 | 未找到 | 公司財報不具名；本次查過 10-K 與 2026 年報前段 | [外部・未經本專案檢核：Snowflake 10-K FY2026、2026 Annual Report，2026-03] |

- Anthropic、OpenAI 是 Snowflake 的供應商（採購模型），不是營收來源：CFO 稱 Anthropic 的 2 億美元為買方 [本專案：research_snowflake.md 第 4.3 節]。

## 前十大股東
| 排名 | 股東 | 持股比例 | 標記 |
|---|---|---|---|
| 1 | The Vanguard Group | 8.46% | [外部・未經本專案檢核：Simply Wall St，資料日期未載（頁面更新約 2025-12）] |
| 2 | BlackRock | 6.86% | 同上 |
| 3 | FMR LLC | 3.06% | 同上 |
| 4 | Jennison Associates | 2.68% | 同上 |
| 5 | State Street Global Advisors | 2.09% | 同上 |
| 6 | JP Morgan Asset Management | 1.95% | 同上 |
| 7 | TIAA-CREF | 1.52% | 同上 |
| 8 | Morgan Stanley Investment Management | 1.46% | 同上 |
| 9 | Geode Capital Management | 1.35% | 同上 |
| 10 | Benoit Dageville（共同創辦人） | 1.35% | 同上 |

- 口徑：資料庫彙整的機構與內部人持股，頁面未標明資料日期，已偏舊；2026 年委託書的受益所有權表本次未能取得，故不是最新申報值 [外部・未經本專案檢核：Simply Wall St 持股頁，2026-11 讀取]。

## 近期重大事件
- 2026-11-03｜Expedition 2026 大會發表 CoCo Cloud Agents 與 CoCo Automations：AI 代理自動處理資料管線的例行修復，工程師審核後上線｜是 AI 功能擴大變現的產品路線 [外部新聞：Snowflake 新聞稿]
- 2026-11-02｜公司宣布 FY2027 Q3 財報定於 2026-12-02 美股收盤後｜研究包寫發布日未找到，此為新資訊 [外部新聞：Snowflake 新聞稿]
- 2026-09-03｜FY2027 Q2 財報（2026-09-02 發布）：營收 US$1,546.8M、產品營收年增 37%，FY2027 產品營收指引上修到 US$6.07B，非 GAAP 產品毛利率指引由 75.0% 降到 74.0%；股價盤中一度大漲逾兩成｜成長加速，但 AI 稀釋毛利 [本專案：research_snowflake.md 第 3.2 節、第 7 節；外部新聞：The Motley Fool，2026-09-03]

## 需要觀察的事項
- 產品營收是否達 Q3 FY27 指引 US$1,588 至 1,593M｜低於 US$1,588M 為模型的失效條件 [本專案：20261124_Snowflake個股模型_v0.2.3.xlsx README!A22；research_snowflake.md 第 7 節]｜2026-12-02 財報 [外部新聞：Snowflake 新聞稿]
- 產品毛利率與推論成本｜非 GAAP 產品毛利率指引 74.0%，跌破即為失效；模型情境 B 的產品毛利率從 FY2027 的 75.3% 降到 FY2032 的 70.7% [本專案：20261124_Snowflake個股模型_v0.2.3.xlsx README!A8、A22]｜2026-12-02 財報
- RPO 年增與淨營收留存率｜RPO 年增由 42% 降到 30%，低於 25% 為失效；留存率低於 120% 為失效，目前 126% [本專案：20261124_Snowflake個股模型_v0.2.3.xlsx README!A22；research_snowflake.md 第 4.2 節]｜同上
- AI 功能的消耗占比｜AI 營收金額已不揭露，模型以 AI 功能 credit 占比停滯為失效條件 [本專案：20261124_Snowflake個股模型_v0.2.3.xlsx README!A22；research_snowflake.md 第 9 節]｜同上
- 股票薪酬與稀釋｜FY2027 Q2 SBC 為 US$423.6M；模型 DCF 不加回 SBC，估值對此敏感 [本專案：research_snowflake.md 第 6 節；snowflake/README.md 合理性關卡診斷]｜同上
- AI Credit 計價切換的實際生效日｜各來源日期互相矛盾（2026-04、2026-07、2026-11-12 版價目表），影響 AI 營收認列口徑 [本專案：research_snowflake.md 第 9 節]｜待公司說明

## 來源
1. 本專案：`software_ai/snowflake/data/research_snowflake.md`；`software_ai/snowflake/data/facts_snowflake.csv`；`software_ai/snowflake/README.md`；`software_ai/snowflake/dist/20261124_Snowflake個股模型_v0.2.3.xlsx`（README、Val）；`software_ai/map/dist/data.json`
2. [Snowflake 2026 Annual Report and Proxy Statement](https://s26.q4cdn.com/463892824/files/doc_financials/2025/ar/Snowflake-2026-Annual-Report-and-Proxy-Statement.pdf)（2026）
3. [Snowflake 10-K FY2026](https://www.sec.gov/Archives/edgar/data/1640147/000164014726000008/snow-20260131.htm)（2026-03）
4. [MarketScreener：Snowflake company governance](https://www.marketscreener.com/quote/stock/SNOWFLAKE-INC-112440376/company-governance/)（2026-09 快照）
5. [Simply Wall St：Snowflake ownership](https://simplywall.st/stocks/us/software/nyse-snow/snowflake/ownership)（資料日期未載）
6. [Snowflake to Announce Financial Results for the Third Quarter of Fiscal 2027 on December 2, 2026](https://www.snowflake.com/en/news/press-releases/fy27-q3-earnings-date-announcement/)（2026-11-02）
7. [Snowflake Expedition 2026：autonomous data engineering](https://www.snowflake.com/en/news/press-releases/expedition-2026-autonomous-data-engineering/)（2026-11-03）
8. [Why Snowflake Stock Rocketed Higher Today](https://www.fool.com/investing/2026/09/03/why-snowflake-stock-rocketed-higher-today/)（2026-09-03）
