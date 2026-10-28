---
ticker: ALAB
name: Astera Labs
name_full: Astera Labs, Inc.
listing: Nasdaq｜ALAB [外部・未經本專案檢核]
hq: 美國加州聖塔克拉拉 [外部・未經本專案檢核]
founded: 2017 [外部・未經本專案檢核]
employees: 本專案未建立
fiscal_year_end: 12 月底（會計年度＝日曆年）[本專案：alab/README.md]
asof: 2026-10-25
---

## 一句話定位
AI 伺服器內部連接晶片商：PCIe 重定時器起家，成長重心正轉向給非 NVIDIA 加速器用的 scale-up 交換晶片（Scorpio X）[本專案：alab/README.md 開頭；universe_ai_semis_v0.2.2.csv Astera Labs 列]。

## 公司在做什麼
- Aries 重定時器：PCIe 訊號在主機板上跑遠了會衰減，重定時器把訊號「整理乾淨再送出去」，讓 GPU、CPU、網卡之間連得穩 [外部・未經本專案檢核；本專案：alab/README.md 開頭]。
- Scorpio 交換晶片：P 系列做 PCIe 交換；X 系列做 scale-up 交換（把同一機櫃內多顆加速器連成一台大電腦），主要給 AMD、Trainium 與其他 ASIC 平台 [本專案：alab/README.md 開頭；Astera Labs模型_v0.1.xlsx In!B5]。
- 另有 Taurus AEC 模組（主動電纜用）與 Leo CXL（記憶體擴充）[本專案：alab/README.md 開頭]。
- 3Q26 起 Scorpio 成為最大產品線，X 超過 P；公司說已有逾 10 家客戶 [本專案：Astera Labs模型_v0.1.xlsx Sources!A6]。

## 主要產品與營收占比
| 產品／部門 | 內容（白話） | 營收占比 | 期間 | 標記 |
|---|---|---|---|---|
| Scorpio X | scale-up 交換晶片 | 20.2%（US$0.38B） | FY2026 錨點 | [本專案：Astera Labs模型_v0.1.xlsx Q!H18、Q!H20；Assumed，公司未揭露] |
| PCIe 與訊號調節 | Aries、Scorpio P、Taurus、Leo | 79.8%（US$1.50B） | FY2026 錨點 | [本專案：Astera Labs模型_v0.1.xlsx Q!H19 ÷ Q!H4；殘差] |
| 合計 | — | 100%（US$1.88B） | FY2026 錨點 | [本專案：Astera Labs模型_v0.1.xlsx Q!H4] |

- 公司不揭露產品線營收；FY2026 錨點＝1Q、2Q 實績＋3Q 指引中點＋4Q 假設 [本專案：alab/README.md 驅動因子]。
- 營收 FY2024 US$0.40B → FY2025 0.85B → FY2026 錨點 1.88B（+121%）[本專案：Astera Labs模型_v0.1.xlsx Q!B4、Q!C4、Q!H4、Q!B26]。
- non-GAAP 毛利率 FY2025 75.8% → FY2026 錨點 72.9%，公司長期目標 70%；淨利率約 39% [本專案：Astera Labs模型_v0.1.xlsx Q!C6、Q!H6、Recon!B13、Q!H15]。

## 市場定位與產業位置
- 本專案推算 Scorpio X 在非 NVIDIA、非 TPU scale-up 交換晶片的有效份額 FY2026 18.8%；若分母含 NVIDIA NVLink，只占 2.1% [本專案：Astera Labs模型_v0.1.xlsx Units!C7、Units!C8]。
- 非 NVIDIA scale-up 交換晶片市場 CY2026 US$2.0B → CY2027 7.7B → CY2029 10.9B（網通模型）；公司估 2030 年商用 scale-up 交換 200 億美元（口徑不同）[本專案：Astera Labs模型_v0.1.xlsx Units!C5、Units!D5、Units!F5、In!E115]。
- PCIe 等的有效份額由 FY2026 39.6% 假設逐年下滑（相對值 0.80／0.75／0.70）[本專案：Astera Labs模型_v0.1.xlsx Units!C13；alab/README.md 驅動因子]。
- 客戶集中：FY2025 單一終端客戶逾 70%、前三大約 86%，公司未具名 [本專案：Astera Labs模型_v0.1.xlsx In!H42]。Amazon 持有認股權證（媒體轉述）[本專案：Astera Labs模型_v0.1.xlsx In!H71]。

## 管理層與經營風格
| 職位 | 姓名 | 任職起 | 背景 |
|---|---|---|---|
| 共同創辦人暨執行長 | Jitendra Mohan | 2017 | [外部・未經本專案檢核] |
| 財務長 | Desmond Lynch | 2026-03-02 | 接替退休的 Mike Tate [外部・未經本專案檢核] |
| 共同創辦人 | Sanjay Gajendra | 2017 | [外部・未經本專案檢核] |

- 經營風格：本專案尚未研究。
- 本專案記錄到：1H26 股票薪酬 1.13 億美元，稀釋股數由 179.6 百萬股增到 183.3 百萬股；買回計畫未找到；現金與有價證券 12.5 億美元（2026-06-30）[本專案：Astera Labs模型_v0.1.xlsx In!H102、In!I102、In!H98]。

## 同業與競爭態勢
| 同業 | 重疊領域 | 市占或規模 | 標記 |
|---|---|---|---|
| NVIDIA（NVLink、NVLink Fusion） | scale-up 交換 | scale-up 交換晶片市場（含 NVLink）CY2026 US$18.3B | [本專案：Astera Labs模型_v0.1.xlsx Units!C4] |
| Credo | AEC、重定時器 | 本專案未建立對比份額 | [本專案：Credo模型_v0.2.xlsx In!A40] |
| Broadcom 等乙太網路 scale-up | scale-up 交換 | 本專案未建立 | — |

- NVLink Fusion 若開放第三方交換晶片，會改變 Scorpio X 的分母；乙太網路 scale-up 可能搶 UALink／PCIe 架構 [本專案：alab/README.md 最大限制；Astera Labs模型_v0.1.xlsx In!I103]。
- 公司稱採用其平台的設計案，每顆加速器內容逾 1,000 美元；本專案以全部非 NVIDIA 加速器為分母，FY2028 只約 195 美元 [本專案：Astera Labs模型_v0.1.xlsx Recon!B15、Recon!C15]。

## 主要客戶
- 財報集中度原文口徑（占營收 ≥10% 的客戶；10-Q Q2 2026）：Customer A 29%（Q2 2026）、27%（Q2 2025）、29%（上半年 2026）、20%（上半年 2025）；Customer B 25%（Q2 2026）、<10%（Q2 2025）、21%（上半年 2026）；Customer C 15%、13%、17%（上半年 2026）、19%（上半年 2025）；Customer D 13%、24%、13%（上半年 2026）、23%（上半年 2025）；Customer E <10%（Q2 2026）、11%（Q2 2025）；Customer F <10%（Q2 2026）、12%（Q2 2025）、<10%（上半年 2026）、15%（上半年 2025）。公司註明：列示客戶中有些是為其終端客戶代購的製造夥伴，終端客戶可能在製造夥伴間調整出貨量，故個別直接客戶的集中度可能波動。 [外部・未經本專案檢核：Astera Labs Form 10-Q，期間至 2026-06-30，2026-08-05]
- FY2025 10-K（占營收 ≥10%）：Customer A 20%、B 20%（FY2024 11%）、C 17%、D 16%（FY2024 24%、FY2023 37%）、E 11%；FY2024：Customer F 36%、D 24%、B 11%；FY2023：D 37%、F 24%、G 18%。應收帳款集中度 2025/12/31：Customer C 14%、D 27%、E 28%。 [外部・未經本專案檢核：Astera Labs Form 10-K，FY2025，2026-02-20]
- 所有客戶皆未具名（Customer A–G）。 [外部・未經本專案檢核]
- 公開已知具名客戶：未找到可引用來源（10-K 僅稱「shipping millions of devices across leading hyperscalers」；本次因搜尋額度用盡，無法查證法說會／新聞稿）。外界常見說法為 Amazon/AWS（曾發行認股權證予客戶，S-1 時期）、Microsoft、Google、Meta 及 NVIDIA 平台生態，但未經查證，不可視為已證實。 [外部・未經本專案檢核：Astera Labs Form 10-Q，期間至 2026-06-30，2026-08-05；Astera Labs Form 10-K，FY2025，2026-02-20]

## 前十大股東
| 排名 | 股東 | 持股比例 | 標記 |
|---|---|---|---|
| 1 | FMR LLC | 14.32% | [外部・未經本專案檢核：資料日 2026/6/30；InsiderSet ALAB Institutional Ownership（13F 2026-06-30），擷取日 2026-10-27] |
| 2 | BlackRock, Inc. | 7.89% | [外部・未經本專案檢核：資料日 2026/6/30；InsiderSet ALAB Institutional Ownership（13F 2026-06-30），擷取日 2026-10-27] |
| 3 | Vanguard Portfolio Management LLC | 4.55% | [外部・未經本專案檢核：資料日 2026/6/30；InsiderSet ALAB Institutional Ownership（13F 2026-06-30），擷取日 2026-10-27] |
| 4 | Vanguard Capital Management LLC | 3.88% | [外部・未經本專案檢核：資料日 2026/6/30；InsiderSet ALAB Institutional Ownership（13F 2026-06-30），擷取日 2026-10-27] |
| 5 | Invesco Ltd. | 2.83% | [外部・未經本專案檢核：資料日 2026/6/30；InsiderSet ALAB Institutional Ownership（13F 2026-06-30），擷取日 2026-10-27] |
| 6 | State Street Corp | 2.12% | [外部・未經本專案檢核：資料日 2026/6/30；InsiderSet ALAB Institutional Ownership（13F 2026-06-30），擷取日 2026-10-27] |
| 7 | Geode Capital Management, LLC | 1.61% | [外部・未經本專案檢核：資料日 2026/6/30；InsiderSet ALAB Institutional Ownership（13F 2026-06-30），擷取日 2026-10-27] |
| 8 | T. Rowe Price Associates Inc | 1.39% | [外部・未經本專案檢核：資料日 2026/6/30；InsiderSet ALAB Institutional Ownership（13F 2026-06-30），擷取日 2026-10-27] |
| 9 | Van Eck Associates Corp | 1.29% | [外部・未經本專案檢核：資料日 2026/6/30；InsiderSet ALAB Institutional Ownership（13F 2026-06-30），擷取日 2026-10-27] |
| 10 | 未找到（BCV 一列數據異常已剔除） | — | — [外部・未經本專案檢核] |

- 口徑：13F 機構持股（資料日 2026/6/30，InsiderSet 彙整）。InsiderSet 該頁不提供占流通股比例（其「Portfolio %」是該機構投資組合中配置 ALAB 的比重，非占 ALAB 股本），故下表「持股比例」為本檔以 2026/6/30 流通股 173,485 千股（10-Q 資產負債表）自行計算，標註為估算。Yahoo Finance 的 ALAB 持股頁無法載入（Nasdaq 頁亦無資料）。Banque Cantonale Vaudoise 一列數據（5,539 股卻報 26.75 億美元）明顯有誤，已剔除，因此僅列 9 家。對照 2026 proxy（受益所有權，資料日 2026/3/1，流通股 171,246,523）：FMR LLC 關係企業 23,498,437 股（13.7%）、Vanguard 12,049,223 股（7.0%）、BlackRock 10,507,515 股（6.1%）；董事與高管合計 17,751,334 股（10.4%），其中 Jitendra Mohan 7,789,805 股（4.5%）、Sanjay Gajendra 7,433,347 股（4.3%）、Manuel Alba 1,923,361 股（1.1%）。 [外部・未經本專案檢核]

## 近期重大事件
- 2026-09-21｜半導體類股反彈，股價單日漲約 12%；執行長在花旗會議稱 Scorpio X 已對主要客戶量產出貨 [外部新聞：TIKR]
- 2026-09-10｜d-Matrix 採用 NVIDIA NVLink Fusion，Astera Labs 為連接夥伴，系統預計 2027 年底進 NVIDIA MGX 機架 [外部新聞：TIKR]
- 2026-08-17｜執行長為繳納股票歸屬稅款賣股；同日 Northland 調升評等 [外部新聞：MarketBeat]
- 2026-08-04｜2Q26 財報：營收 US$392M、non-GAAP EPS 0.80；3Q26 指引營收 540–560M、EPS 1.16–1.21；Scorpio 將提前一季成為最大產品線 [本專案：Astera Labs模型_v0.1.xlsx In!E54、In!E60、In!D61、In!F61、In!D64、In!F64、Sources!A6]

## 需要觀察的事項
- 顆數情境 C｜FY2028 EPS −56%，最大變數；營業費用彈性高，營收下滑時淨利率由 39% 降到 29% [本專案：alab/README.md 結果]｜3Q26 財報 2026-11-03 盤後
- Scorpio X 份額｜份額低／高使 FY2028 EPS −25%／+30% [本專案：alab/README.md 結果]｜同上，看 Scorpio 是否如期成為最大產品線
- PCIe 等份額｜低／高 −14%／+18%；最大客戶部分轉單 −11% [本專案：alab/README.md 結果]｜同上
- 非 NVIDIA scale-up 市場｜×0.8 使 FY2028 EPS −11%；NVLink Fusion 是主要不確定性 [本專案：alab/README.md 結果、開頭]｜AMD、AWS Trainium 機架出貨
- 毛利率｜模型假設由 FY2026 72.9% 降到 70.0%／69.5%／69.0%（公司長期目標 70%）；毛利率低 −10% [本專案：alab/README.md 驅動因子、結果]｜3Q26 毛利率指引約 72% [本專案：Astera Labs模型_v0.1.xlsx In!E62]

## 來源
1. 本專案：`alab/README.md`；`alab/dist/20261025_Astera Labs模型_v0.1.xlsx`（Q、Units、In、Recon、Sources）；`crdo/dist/20261025_Credo模型_v0.2.xlsx`（In）；`markets/invest_map/data/universe_ai_semis_v0.2.2.csv`
2. [Astera Labs Stock Jumped 12% Yesterday（TIKR）](https://www.tikr.com/blog/astera-labs-stock-jumped-12-yesterday-the-street-sees-14-more-upside-heres-why)（2026-09-22）
3. [Astera Labs Stock Price Down 6.9%（MarketBeat）](https://www.marketbeat.com/instant-alerts/price-astera-labs-nasdaq-alab-stock-price-down-69-should-you-sell-2026-09-08/)（2026-09-08）
4. [Astera Labs management（Simply Wall St）](https://simplywall.st/stocks/mx/semiconductors/bmv-alab/astera-labs-shares/management)（2026-10）
