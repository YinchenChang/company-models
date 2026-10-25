---
ticker: AMD
name: AMD
name_full: Advanced Micro Devices, Inc.
listing: Nasdaq｜AMD
hq: 美國加州聖塔克拉拉 [外部・未經本專案檢核]
founded: 1969 [外部・未經本專案檢核]
employees: 本專案未建立
fiscal_year_end: 12 月最後一個週六（FY(t)＝CY(t)，不平移）[本專案：amd/README.md 開頭]
asof: 2026-10-25
---

## 一句話定位
AI 加速器（GPU）市場中 NVIDIA 以外的第二大 GPU 供應商，同時賣伺服器 CPU；資料中心 GPU 營收＝加速器市場金額 × AMD 金額份額 [本專案：AMD模型_v0.2.1.xlsx README!A3]。

## 公司在做什麼
- 賣資料中心晶片：Instinct GPU（MI355X、MI450／MI455X，用來訓練與執行 AI 模型）、EPYC 伺服器 CPU、DPU、FPGA；另有個人電腦（Client）、遊戲主機與顯示卡（Gaming）、嵌入式（Embedded）[本專案：AMD模型_v0.2.1.xlsx Q!A5:A10]。
- Helios 是 AMD 的整櫃 AI 系統（MI455X GPU＋EPYC CPU＋網通），把 GPU 賣成整櫃，對標 NVIDIA 的 NVL72 [外部・未經本專案檢核]。
- 大客戶以「GW（十億瓦）」為單位簽約：OpenAI 6 GW、Meta 6 GW（兩案合計 12 GW）、Anthropic 最多 2 GW、Oracle 首批 5 萬顆 MI450 [本專案：AMD模型_v0.2.1.xlsx In!E122、In!E112、In!E124、In!E125]。
- 公司不揭露資料中心內 GPU 與 CPU 的拆分；本專案的 GPU 金額是假設拆分（Assumed）[本專案：AMD模型_v0.2.1.xlsx Q!M6]。

## 主要產品與營收占比
| 產品／部門 | 內容（白話） | 營收占比 | 期間 | 標記 |
|---|---|---|---|---|
| Data Center | Instinct GPU、EPYC CPU、DPU、FPGA | 48.0%（US$16.6B ÷ 34.6B） | FY2025 | [本專案：AMD模型_v0.2.1.xlsx Q!F5、Q!F4，占比為換算] |
| Client | 個人電腦 CPU | 30.7%（US$10.6B） | FY2025 | [本專案：AMD模型_v0.2.1.xlsx Q!F8，換算] |
| Gaming | 遊戲主機晶片、顯示卡 | 11.3%（US$3.9B） | FY2025 | [本專案：AMD模型_v0.2.1.xlsx Q!F9，換算] |
| Embedded | 工業、車用、通訊用嵌入式晶片 | 10.0%（US$3.5B） | FY2025 | [本專案：AMD模型_v0.2.1.xlsx Q!F10，換算] |
| 合計 | — | 100% | FY2025 | — |

- FY2026 錨點（Q1、Q2 實績＋Q3 指引＋Q4 假設）營收 US$49.6B，其中 Data Center 30.9B（62.1%）[本專案：AMD模型_v0.2.1.xlsx Q!K4、Q!K5，占比為換算]。
- 資料中心 GPU：FY2025 約 US$6.5B、FY2026 錨點約 15.2B，皆為假設拆分 [本專案：AMD模型_v0.2.1.xlsx Q!F6、Q!K6]；模型 FY2027 為 32.8B [本專案：AMD模型_v0.2.1.xlsx PL!E8]。
- non-GAAP 毛利率：FY2025 52.4%、FY2026 錨點 55.8%、模型 FY2027 49.7%（HBM 2027 漲價、吸收比例 0.5 壓 GPU 毛利率）[本專案：AMD模型_v0.2.1.xlsx Q!F13、Q!K13、PL!E12、Recon!J15]。

## 市場定位與產業位置
- 本專案推算 CY2027 AI 加速器**金額份額** AMD 6.7%、**顆數份額** 6.3%；NVIDIA 金額份額 75.1% [本專案：AMD模型_v0.2.1.xlsx Recon!C27、Recon!C28、Recon!C29]。
- 同一題 AI半導體 先驗比對的推導值為 3–6%（候選母體表同此寫法），比模型的 6.7% 低，兩者並列 [本專案：先驗比對_NVDA_20261023.md N5c；universe_ai_semis_v0.2.2.csv AMD 列]。
- 每顆售價水準：FY2026 隱含每顆認列售價約 US$20.2K，HSBC 估 MI355X 約 23K [本專案：AMD模型_v0.2.1.xlsx Recon!C18、Recon!B18]。
- 大單附帶認股權證：OpenAI、Meta 合計上限 3.2 億股，約為 FY2026 股數的 19%，依出貨 GW 分批歸屬 [本專案：AMD模型_v0.2.1.xlsx In!E111、Recon!J25]。
- 上游依賴 HBM（高頻寬記憶體）：HBM 是 AMD 的成本，漲價時 AMD 吸收多少由模型假設決定 [本專案：amd/README.md 驅動因子]。

## 管理層與經營風格
| 職位 | 姓名 | 任職起 | 背景 |
|---|---|---|---|
| 董事長兼執行長 | 蘇姿丰（Lisa Su） | 2014 | [外部・未經本專案檢核] |
| 財務長 | Jean Hu | 2022 | 2026-09 仍以財務長身分公開發言 [外部・未經本專案檢核] |

- 經營風格：本專案尚未研究。
- 本專案記錄到的資本配置事實：2026 上半年買回僅 US$0.22B [本專案：AMD模型_v0.2.1.xlsx In!H110]。

## 同業與競爭態勢
| 同業 | 重疊領域 | 市占或規模 | 標記 |
|---|---|---|---|
| NVIDIA | AI GPU、整櫃系統 | CY2027 加速器金額份額 75.1% | [本專案：NVIDIA模型_v0.2.xlsx Units!D32] |
| Google TPU（Broadcom、聯發科設計） | 自研 AI 晶片 | CY2027 金額份額 9.8% | [本專案：NVIDIA模型_v0.2.xlsx Units!D34] |
| AWS Trainium（世芯、Marvell 設計） | 自研 AI 晶片 | CY2027 金額份額 3.0% | [本專案：NVIDIA模型_v0.2.xlsx Units!D35] |
| 其他 ASIC（MTIA、Maia、昇騰等） | 自研 AI 晶片 | CY2027 金額份額 5.4% | [本專案：NVIDIA模型_v0.2.xlsx Units!D36] |

- AMD 與 NVIDIA 模型用同一份顆數與每顆價，所以份額可互相加總 [本專案：AMD模型_v0.2.1.xlsx Recon!J29]。
- 伺服器 CPU（與 Intel、Arm 架構自研 CPU 競爭）的份額：本專案未建立。

## 近期重大事件
- 2026-10-07｜MI455 加速器與 Helios 整櫃（使用 HBM4）開始出貨｜MI450 世代的售價是模型最敏感的變數之一 [外部新聞：Cryptopolitan]
- 2026-09-28｜宣布以全股票約 US$8.2B 收購 AI 研究實驗室 World Labs（李飛飛共同創辦），預計 2026 年底前完成｜股數稀釋，模型未納入 [外部新聞：Northeast Times]
- 2026-09-28｜Advancing AI 2026 大會展示 Helios 整櫃 [外部新聞：Stocktwits]
- 2026-08-04｜Q2 2026 財報：營收 US$11.5B、Data Center 6.7B；Q3 營收指引 13.0B [本專案：AMD模型_v0.2.1.xlsx Q!H4、Q!H5、Q!I4]；法說稱 2027 資料中心「逾一倍」（二手）[本專案：AMD模型_v0.2.1.xlsx Recon!I7]

## 需要觀察的事項
- 每顆售價（k_rec 0.673 是否延續）｜改用市場層每顆價時 FY2027 EPS +14.0% [本專案：AMD模型_v0.2.1.xlsx Recon!C17；amd/README.md 敏感度]｜Q3 2026 財報 2026-11-03 [本專案：amd/README.md 下次更新]
- HBM 漲價吸收比例｜0 或 1 使 FY2027 EPS ±10.9%；FY2028 起情境 A／B／C 取 0.2／0.5／0.7 [本專案：amd/README.md 敏感度、v0.2.1 節]｜HBM4 合約價消息
- 2027 資料中心「逾一倍」可達性｜模型 FY2027 資料中心 US$56.4B，比公司說法 61.7B 低 8.6% [本專案：AMD模型_v0.2.1.xlsx Recon!C7、Recon!B7、Recon!D7]｜同上
- 模型 EPS 與共識的落差｜FY2027 模型 9.14 vs 共識 15.76（stockanalysis，二手）差 −42% [本專案：AMD模型_v0.2.1.xlsx Recon!C10、Recon!B10、Recon!D10]｜同上
- 認股權證歸屬與 World Labs 增發｜權證占 FY2029 股數 7.2%（A）[本專案：AMD模型_v0.2.1.xlsx Recon!C25]；World Labs 本專案未建立｜收購完成（預計 2026 年底）

## 來源
1. 本專案：`amd/README.md`；`amd/dist/20261025_AMD模型_v0.2.1.xlsx`（README、In、Q、PL、Recon）；`nvidia/dist/20261025_NVIDIA模型_v0.2.xlsx`（Units）；AI半導體 `先驗比對_NVDA_20261023.md`；`markets/invest_map/data/universe_ai_semis_v0.2.2.csv`
2. [AMD to Buy AI Research Lab World Labs in $8.2 Billion Stock Deal](https://northeasttimes.com/2026/09/28/amd-to-buy-ai-research-lab-world-labs-in-8-2-billion-stock-deal)（2026-09-28）
3. AMD MI455／Helios 開始出貨（Cryptopolitan，2026-10-07；經 [autobound 新聞彙整](https://www.autobound.ai/companies/amd/news) 取得，僅標題）
4. [AMD Helios in production（igor'sLAB 轉述 Reuters）](https://www.igorslab.de/en/amd-helios-in-production-mi455x-epyc-venice-openai-challenge-nvidia-ai-racks/)（日期未標示）
5. [AMD Reports Second Quarter 2026 Financial Results（8-K）](https://www.sec.gov/Archives/edgar/data/0000002488/000000248826000121/q22026991.htm)（2026-08）
6. [AMD CFO Jean Hu（247wallst）](https://247wallst.com/investing/2026/09/09/amd-rises-3-as-cfo-lifts-ai-chip-market-forecast-toward-3-trillion-intel-ticks-up/)（2026-09-09）
