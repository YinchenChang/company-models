---
ticker: NVDA
name: NVIDIA
name_full: NVIDIA Corporation
listing: Nasdaq｜NVDA
hq: 美國加州聖塔克拉拉 [外部・未經本專案檢核]
founded: 1993 [外部・未經本專案檢核]
employees: 本專案未建立
fiscal_year_end: 1 月底（FY(t)＝t−1 年 2 月～t 年 1 月）[本專案：nvidia/README.md 開頭]
asof: 2026-10-25
---

## 一句話定位
AI 加速器（GPU）與整櫃系統的最大供應商，靠「加速器顆數 × 每顆售價＋網通附掛」賺錢 [本專案：nvidia/README.md 開頭]。

## 公司在做什麼
- 賣 AI 運算晶片（GPU，圖形處理器，現在主要用來訓練與執行 AI 模型）與整櫃伺服器系統，客戶是雲端業者、AI 實驗室與企業 [外部・未經本專案檢核]。
- 產品世代：Blackwell（B200／B300）、Rubin、Rubin Ultra；本專案以各世代顆數建模 [本專案：NVIDIA模型_v0.1.1.xlsx Units!A6:A9]。
- 網通（把成千上萬顆 GPU 連起來的交換器、網卡、NVLink）約占資料中心營收兩成：FY2027 Q1 運算 US$60.4B、網通 US$14.8B [本專案：先驗比對_NVDA_20261023.md 資料表、N3]。
- 公司自 FY2027 Q1 起只揭露 Data Center 與 Edge Computing 兩個平台 [本專案：NVIDIA模型_v0.1.1.xlsx Q!M6]。

## 主要產品與營收占比
| 產品／部門 | 內容（白話） | 營收占比 | 期間 | 標記 |
|---|---|---|---|---|
| Data Center | GPU、整櫃系統、網通 | 89.7%（US$193.7B ÷ 215.9B） | FY2026 | [本專案：NVIDIA模型_v0.1.1.xlsx Q!F5、Q!F4，占比為換算] |
| Edge Computing | 舊 Gaming＋ProViz＋Auto＋OEM 合計 | 10.3%（US$22.2B） | FY2026 | [本專案：NVIDIA模型_v0.1.1.xlsx Q!F6] |
| 合計 | — | 100% | FY2026 | — |

- FY2027 Q2 單季 Data Center 占比升到 92.5%（89.0 ÷ 96.2）[本專案：NVIDIA模型_v0.1.1.xlsx Q!H5、Q!H4，換算]。
- 毛利率：FY2026 71.1%；FY2027 錨點 73.7% [本專案：NVIDIA模型_v0.1.1.xlsx Q!F9、Q!K9]。
- 資料中心內運算與網通的全年拆分：本專案未建立（README 記為未找到）。

## 市場定位與產業位置
- 本專案推算 CY2027 加速器**金額份額** NVIDIA 75.1%，但**顆數份額**只有 40.9%：ASIC 每顆便宜很多 [本專案：NVIDIA模型_v0.1.1.xlsx Units!D32、Units!D26]。
- 同一題 Broadcom 模型算出 71.3%，兩個模型 ASIC 單價假設不同，並列 [本專案：Broadcom模型_v0.1.xlsx Recon!C24]；先驗比對推導區間 75–82% [本專案：先驗比對_NVDA_20261023.md N5a]。
- 上游依賴：台積電（晶圓、CoWoS 封裝）與 HBM（高頻寬記憶體）；本專案台積電模型判斷 2027 年 HBM 比 CoWoS 先卡 [本專案：tsmc/README.md 主要結果]。
- 中國：公司 FY2027 Q3 指引假設中國資料中心運算營收為 0 [本專案：nvidia/README.md 驅動因子]。

## 管理層與經營風格
| 職位 | 姓名 | 任職起 | 背景 |
|---|---|---|---|
| 創辦人、執行長 | 黃仁勳（Jensen Huang） | 1993 | 共同創辦人 [外部・未經本專案檢核] |
| 財務長 | Colette Kress | 2013 | 2026 年仍在任，未見異動 [外部・未經本專案檢核] |

- 經營風格：本專案尚未研究。
- 本專案只記錄到：公司供給承諾約 US$279B（二手，原文待複核）[本專案：nvidia/README.md 資料狀態]。

## 同業與競爭態勢
| 同業 | 重疊領域 | 市占或規模 | 標記 |
|---|---|---|---|
| Google TPU（Broadcom 設計） | 自研 AI 晶片 | CY2027 金額份額 9.8% | [本專案：NVIDIA模型_v0.1.1.xlsx Units!D34] |
| AMD Instinct | GPU | CY2027 金額份額 6.7% | [本專案：NVIDIA模型_v0.1.1.xlsx Units!D33] |
| AWS Trainium | 自研 AI 晶片 | CY2027 金額份額 3.0% | [本專案：NVIDIA模型_v0.1.1.xlsx Units!D35] |
| 其他 ASIC（MTIA、Maia、昇騰等） | 自研 AI 晶片 | CY2027 金額份額 5.4% | [本專案：NVIDIA模型_v0.1.1.xlsx Units!D36] |

- ASIC 顆數 2027 年已過半（JPMorgan 54%），但金額只占約 14–19% [本專案：先驗比對_NVDA_20261023.md N5b]。
- 本專案市場層 NVIDIA 顆數 CY2027 到 CY2028 持平（9.7M→9.5M）[本專案：先驗比對_NVDA_20261023.md 資料表]。

## 主要客戶
- 直接客戶（direct customers）集中度，10-Q 原文口徑（占總營收）：
- Q2 FY27：「one direct customer represented 16% of total revenue, which was primarily attributable to the Compute & Networking segment.」FY27 上半年：「three direct customers represented 16%, 15%, and 13% of total revenue」，均主要屬 Compute & Networking。去年同期 Q2 FY26 為兩家 23% 與 16%；FY26 上半年 20% 與 15%。 [外部・未經本專案檢核：NVIDIA Form 10-Q，期間至 2026-07-26，2026 年 8 月下旬]
- Q1 FY27：三家直接客戶分別占 21%、17%、16%（Q1 FY26 為 16%、14%）。 [外部・未經本專案檢核：NVIDIA Form 10-Q，期間至 2026-04-26，2026 年 5 月]
- 應收帳款集中度：2026/7/26 五家直接客戶占應收 22%、14%、13%、11%、10%（2026/4/26 為三家 30%、18%、16%）。 [外部・未經本專案檢核：NVIDIA Form 10-Q，期間至 2026-07-26，2026 年 8 月下旬；NVIDIA Form 10-Q，期間至 2026-04-26，2026 年 5 月]
- 間接客戶：10-Q 稱營收集中於少數直接與間接客戶，部分間接客戶個別占營收 10% 以上；並估計「one AI research and deployment company」透過向 NVIDIA 客戶購買雲端服務，貢獻了「meaningful amount」營收。直接客戶多為系統整合商與經銷商，間接客戶含 CSP、AI clouds、AI model makers、企業與公部門。 [外部・未經本專案檢核：NVIDIA Form 10-Q，期間至 2026-07-26，2026 年 8 月下旬]
- 結構性說明：Q1 FY27 CFO 稱 Hyperscaler 約占 Data Center 營收 50%；Q2 FY27 Hyperscale 48,710 ÷ Data Center 89,023 ≈ 55%（本檔計算）。 [外部・未經本專案檢核：NVIDIA Form 10-Q，期間至 2026-04-26，2026 年 5 月；NVIDIA Q2 FY27 CFO Commentary（8-K），2026-08-26]
- 10-Q 未具名直接客戶（Customer A/B…）；外部媒體曾推測為鴻海/Foxconn、Quanta、Wistron、Dell、Supermicro 等 ODM/OEM（推測，未獲 NVIDIA 證實，本次未取得可引用的原文來源）。 [外部・未經本專案檢核]
- 公開具名客戶／合作對象（Q2 FY27 法說會逐字稿，2026/8/26） [外部・未經本專案檢核：NVIDIA Q2 FY2027 Earnings Call 逐字稿，2026-08-26]：
- AWS：宣布本季起額外部署 200 萬顆 GPU 與 Vera CPU。 [外部・未經本專案檢核]
- Oracle（OCI）：Vera CPU 首批出貨的 lead partner。 [外部・未經本專案檢核]
- SpaceXAI：與 OCI 並列 Vera CPU lead partner。 [外部・未經本專案檢核]
- OpenAI：現有與規劃承諾約 12 GW NVIDIA Compute，含 4.25 GW Portsmouth 站點。 [外部・未經本專案檢核]
- CoreWeave、Nebius、Nscale：neocloud，NVIDIA 稱協助其建立基礎設施業務；Nebius 預計為 Groq 3 LPX 系統量產首批客戶。 [外部・未經本專案檢核]
- 區域 neocloud：Firebird、Cassava Technologies、GMI Cloud、Yotta、Neysa、Firmus、YTL AI Cloud。 [外部・未經本專案檢核]
- 企業/主權：Bristol Myers Squibb（Vera Rubin AI factory）、Roche、Lilly、Samsung Electronics（cuLitho）、Hudson River Trading、Jane Street、Noetra（日本）、LG、Hyundai Motor Group。 [外部・未經本專案檢核]
- Anthropic、Meta、Google（Gemini）、xAI(Grok)：逐字稿僅列為在 NVIDIA 上運行的模型，未稱為購買方。Microsoft 未出現在該逐字稿。 [外部・未經本專案檢核]
- 結論：未找到 NVIDIA 官方揭露各客戶占營收具名比例。 [外部・未經本專案檢核]

## 前十大股東
| 排名 | 股東 | 持股比例 | 標記 |
|---|---|---|---|
| 1 | BlackRock Inc. | 8.06% | [外部・未經本專案檢核：資料日 2026/6/30；Yahoo Finance NVDA Holders，資料截至 2026-06-30，擷取日 2026-10-27；InsiderSet NVDA Institutional Ownership，13F 2026-06-30，擷取日 2026-10-27] |
| 2 | Vanguard Capital Management LLC | 6.39% | [外部・未經本專案檢核：資料日 2026/6/30；Yahoo Finance NVDA Holders，資料截至 2026-06-30，擷取日 2026-10-27；InsiderSet NVDA Institutional Ownership，13F 2026-06-30，擷取日 2026-10-27] |
| 3 | FMR, LLC（Fidelity） | 4.26% | [外部・未經本專案檢核：資料日 2026/6/30；Yahoo Finance NVDA Holders，資料截至 2026-06-30，擷取日 2026-10-27；InsiderSet NVDA Institutional Ownership，13F 2026-06-30，擷取日 2026-10-27] |
| 4 | State Street Corporation | 4.19% | [外部・未經本專案檢核：資料日 2026/6/30；Yahoo Finance NVDA Holders，資料截至 2026-06-30，擷取日 2026-10-27；InsiderSet NVDA Institutional Ownership，13F 2026-06-30，擷取日 2026-10-27] |
| 5 | Geode Capital Management, LLC | 2.52% | [外部・未經本專案檢核：資料日 2026/6/30；Yahoo Finance NVDA Holders，資料截至 2026-06-30，擷取日 2026-10-27；InsiderSet NVDA Institutional Ownership，13F 2026-06-30，擷取日 2026-10-27] |
| 6 | Vanguard Portfolio Management LLC | 2.15% | [外部・未經本專案檢核：資料日 2026/6/30；Yahoo Finance NVDA Holders，資料截至 2026-06-30，擷取日 2026-10-27；InsiderSet NVDA Institutional Ownership，13F 2026-06-30，擷取日 2026-10-27] |
| 7 | JPMorgan Chase & Co | 1.86% | [外部・未經本專案檢核：資料日 2026/6/30；Yahoo Finance NVDA Holders，資料截至 2026-06-30，擷取日 2026-10-27；InsiderSet NVDA Institutional Ownership，13F 2026-06-30，擷取日 2026-10-27] |
| 8 | T. Rowe Price Associates Inc | 1.53% | [外部・未經本專案檢核：資料日 2026/6/30；Yahoo Finance NVDA Holders，資料截至 2026-06-30，擷取日 2026-10-27；InsiderSet NVDA Institutional Ownership，13F 2026-06-30，擷取日 2026-10-27] |
| 9 | Morgan Stanley | 1.48% | [外部・未經本專案檢核：資料日 2026/6/30；Yahoo Finance NVDA Holders，資料截至 2026-06-30，擷取日 2026-10-27；InsiderSet NVDA Institutional Ownership，13F 2026-06-30，擷取日 2026-10-27] |
| 10 | Invesco Ltd. | 1.37% | [外部・未經本專案檢核：資料日 2026/6/30；Yahoo Finance NVDA Holders，資料截至 2026-06-30，擷取日 2026-10-27；InsiderSet NVDA Institutional Ownership，13F 2026-06-30，擷取日 2026-10-27] |

- 口徑：13F 機構持股（資料日期 2026/6/30，Yahoo Finance 整理，% Out = 占流通在外股數）。同一集團下 Vanguard 分為兩個申報主體，已照原表列示。Yahoo 另列：內部人持股 4.01%、機構持股 71.41%、機構家數 7,703。創辦人 Jensen Huang 為內部人，未列在該 13F 表，最新 proxy beneficial ownership 比例本次未查。 [外部・未經本專案檢核]

## 近期重大事件
- 2026-10-19｜在德國投資高峰會與德國政府簽署「實體 AI」合作備忘錄，並宣布在歐洲投資新創與研究機構｜歐洲主權 AI 需求 [外部新聞：The Next Web]
- 2026-10-14｜IDC 報告指 NVIDIA 在中國 AI 晶片的份額於 2026 上半年明顯下滑，本土晶片（華為等）替代｜中國營收難回升 [外部新聞：GuruFocus 引 IDC]
- 2026-10-10｜FT 報導 NVIDIA 正探討與 AI 新創 Reflection AI 的交易（含收購）｜往模型端延伸 [外部新聞：Bloomberg 轉述 FT]
- 2026-08-27｜首批 H200 依美國新許可制度出貨中國，金額很小｜Q3 指引仍假設中國資料中心運算為 0 [外部新聞：SCMP]
- 2026-08-26｜FY2027 Q2 財報：營收 US$96.2B、Data Center 89.0B；Q3 指引 108B；FY2028 營收指引年增約 70%、毛利率約 72–73% [本專案：NVIDIA模型_v0.1.1.xlsx Q!H4、Q!H5、Q!I4、Recon!B6、Recon!B8]；Vera Rubin 全面量產 [外部新聞：StorageNewsletter]

## 需要觀察的事項
- 營收水準係數 k_rec（0.828）｜FY2027 由下而上比實績高 20.7%，差額來源未驗證；Q3 實績可檢驗是時點還是長期水準｜FY2027 Q3 財報（2026-11 中旬，公司尚未公告日期）[本專案：nvidia/README.md v0.1.1 第 1 點、Recon B2]
- FY2028 毛利率｜模型 68.7% vs 公司 72–73%，差在 HBM 成本吸收比例（假設 0.5）[本專案：NVIDIA模型_v0.1.1.xlsx Recon!C8、Recon!B8]｜同上
- ASIC 金額份額｜±5 個百分點使 FY2028 EPS ±6.6% [本專案：nvidia/README.md 敏感度]｜Broadcom 2026-12 財報
- 2028 後顆數持平｜CY2029 顆數持平時 FY2030 EPS 只成長 2.1% [本專案：nvidia/README.md 核心問題的答案]｜CoWoS、HBM 2027–28 擴產消息
- 中國恢復｜恢復 10／20／40（US$B）使 FY2028 EPS +1.5%／+3.0%／+6.0% [本專案：nvidia/README.md 敏感度]｜美國許可與中國採購政策

## 來源
1. 本專案：`nvidia/README.md`；`nvidia/dist/20261023_NVIDIA模型_v0.1.1.xlsx`（Q、Units、Recon）；`avgo/dist/20261023_Broadcom模型_v0.1.xlsx`（Recon）；AI半導體 `先驗比對_NVDA_20261023.md`；`tsmc/README.md`（2026-10-23）
2. [Germany signs AI partnership with Nvidia at Berlin investment summit](https://thenextweb.com/news/germany-nvidia-ai-partnership-berlin-summit)（2026-10-20）
3. [Nvidia's Market Share in China AI Chips Drops Amid U.S. Export Restrictions](https://www.gurufocus.com/news/9122830/nvidias-market-share-in-china-ai-chips-drops-amid-us-export-restrictions-nvda)（2026-10-14）
4. [Nvidia Explores Deal Options With Reflection AI, Financial Times Says](https://www.bloomberg.com/news/articles/2026-10-10/nvidia-is-in-talks-to-acquire-reflection-ai-the-ft-reports)（2026-10-10，僅標題）
5. [Nvidia ships first H200 chips to China, but forecasts no data-centre computing revenue](https://www.scmp.com/tech/big-tech/article/3365383/nvidia-ships-first-h200s-china-forecasts-no-data-centre-computing-revenue)（2026-08-27）
6. [NVIDIA fiscal 2Q27 financial results](https://www.storagenewsletter.com/2026/08/27/nvidia-fiscal-2q27-financial-results/)（2026-08-27）
7. [Nvidia Earnings Date: Q3 FY2027](https://www.cmelitegroup.com/knowledge-hub/nvidia-earnings-date/)（2026-10-21，日期未確認）
