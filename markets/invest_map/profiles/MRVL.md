---
ticker: MRVL
name: Marvell（邁威爾）
name_full: Marvell Technology, Inc.
listing: Nasdaq｜MRVL
hq: 美國加州聖塔克拉拉（營運總部）[外部・未經本專案檢核]
founded: 1995 [外部・未經本專案檢核]
employees: 本專案未建立
fiscal_year_end: 1 月底前後（FY2027＝2026-02-01～2027-01-30；FY(t) 對應 CY(t−1)）[本專案：mrvl/README.md 開頭]
asof: 2026-10-25
---

## 一句話定位
資料中心網通與客製 AI 晶片的設計公司：靠光互連晶片（光模組裡的 DSP）、客製 XPU（ASIC）與交換器晶片賺錢；資料中心營收＝客製 XPU＋光互連＋交換器與其他 [本專案：mrvl/README.md 開頭]。

## 公司在做什麼
- 光互連：光模組裡負責訊號處理的 PAM4 DSP、TIA／驅動器、相干 DSP，用來把 AI 伺服器之間的電訊號轉成光訊號傳遞 [本專案：邁威爾模型_v0.2.xlsx Q!J13]。
- 客製 XPU（ASIC，為單一雲端客戶量身設計的 AI 晶片）與 XPU 周邊晶片（attach）；主要平台是 AWS Trainium 部分世代與其他雲端自研晶片 [本專案：邁威爾模型_v0.2.xlsx Units!A5:A16；In!H45]。
- 交換器晶片、AEC（主動式電纜）與 retimer（訊號中繼晶片）、儲存控制晶片 [本專案：邁威爾模型_v0.2.xlsx Q!A10:A12]。
- 非資料中心：企業網通、電信、消費、車用／工業（合稱「通訊與其他」）[本專案：邁威爾模型_v0.2.xlsx Q!A6]。

## 主要產品與營收占比
| 產品／部門 | 內容（白話） | 營收占比 | 期間 | 標記 |
|---|---|---|---|---|
| 資料中心：光互連 | 光模組 DSP 等（資料中心差額推得） | 42.1%（US$3.45B ÷ 8.19B） | FY2026 | [本專案：邁威爾模型_v0.2.xlsx Q!B13、Q!B4，Derived；占比為換算] |
| 資料中心：客製 | XPU＋XPU attach | 18.3%（US$1.5B） | FY2026 | [本專案：邁威爾模型_v0.2.xlsx Q!B7，法說二手；換算] |
| 資料中心：儲存與其他 | 儲存控制晶片等 | 7.9%（US$0.65B） | FY2026 | [本專案：邁威爾模型_v0.2.xlsx Q!B12，Assumed；換算] |
| 資料中心：交換器 | 乙太網路交換器晶片 | 3.7%（US$0.3B） | FY2026 | [本專案：邁威爾模型_v0.2.xlsx Q!B10，二手；換算] |
| 資料中心：AEC 與 retimer | 主動式電纜、訊號中繼 | 2.4%（US$0.2B） | FY2026 | [本專案：邁威爾模型_v0.2.xlsx Q!B11，二手；換算] |
| 通訊與其他 | 企業網通、電信、消費、車用 | 25.6%（US$2.09B） | FY2026 | [本專案：邁威爾模型_v0.2.xlsx Q!B6；換算] |
| 合計 | — | 100% | FY2026 | — |

- 資料中心合計占 74.4%（US$6.10B）；FY2027 錨點升到 81.3%（9.76B ÷ 12.0B）[本專案：邁威爾模型_v0.2.xlsx Q!B5、Q!G5、Q!G4，換算]。
- 成長最快的線：光互連 FY2027 錨點 US$6.10B，年增 76.9%（公司說法「超過 70%」）[本專案：邁威爾模型_v0.2.xlsx Q!G13、Q!G14、Q!J14]。
- non-GAAP 毛利率：FY2026 59.5%、FY2027 錨點 58.4%；模型 FY2030 降到 48.4%（客製含代購 HBM 稀釋）[本專案：邁威爾模型_v0.2.xlsx Q!B16、Q!G16、PL!G11]。

## 市場定位與產業位置
- 本專案假設 Marvell 占 AWS Trainium 顆數 10%／10%／8%、其他 ASIC 5%／6%／15%、Google TPU 0／0／5%（CY2026／27／28；皆 Assumed）[本專案：邁威爾模型_v0.2.xlsx Units!C5:E5、Units!C9:E9、Units!C13:E13]。
- 外部對 Marvell ASIC 份額的估計：2027 約 8%（Counterpoint）／10–12%（JPMorgan）[本專案：先驗比對_AVGO_20261023.md 資料表]。
- Trainium 3 主供應商為世芯，Marvell 做 Trainium 2.5（二手）[本專案：邁威爾模型_v0.2.xlsx In!H45]。
- 光晶片組：模型的 Marvell 光互連約等於本專案光晶片組市場的 110%（CY2026），代表光互連包含 AI 後端可插拔光模組以外的產品，或市場層光學偏低 [本專案：邁威爾模型_v0.2.xlsx Units!C27、Recon!J24]。
- 客製晶片若由 Marvell 代客戶採購 HBM，毛利率會被拉低；客戶是否自購 HBM 本專案未找到 [本專案：mrvl/README.md 對帳節]。

## 管理層與經營風格
| 職位 | 姓名 | 任職起 | 背景 |
|---|---|---|---|
| 董事長兼執行長 | Matt Murphy | 2016 | [外部・未經本專案檢核] |
| 財務長 | Dan Durn | 2026-06-15 | 前 Adobe、Applied Materials、NXP 財務長；接替 Willem Meintjes [外部・未經本專案檢核] |

- 經營風格：本專案尚未研究。
- 近期人事：財務長於 2026-06 更換（見上表）[外部・未經本專案檢核]。

## 同業與競爭態勢
| 同業 | 重疊領域 | 市占或規模 | 標記 |
|---|---|---|---|
| Broadcom | 客製 ASIC、交換器 | 占 Google TPU 顆數 80%、其他 ASIC 55%（CY2027 假設） | [本專案：Broadcom模型_v0.2.xlsx Units!D5、Units!D11] |
| 世芯-KY | Trainium 設計服務 | 占 Trainium 顆數 55%（CY2027，情境 A 假設） | [本專案：世芯-KY模型_v0.1.xlsx Units!D6] |
| 聯發科 | Google TPU | 占 TPU 顆數 15%（CY2027 假設） | [本專案：聯發科模型_v0.1.xlsx Units!D6] |
| 中際旭創等光模組廠 | 光互連（Marvell 為其晶片供應商） | 本專案未建立 | — |

- 本專案以「份額加總不超過 100%」檢查各家：CY2028 Google TPU 為 Broadcom 75%＋Marvell 5%＋聯發科 18%＝98% [本專案：20261025_Broadcom模型_v0.2.xlsx Units!E5；聯發科模型_v0.1.xlsx Units!E13]。
- 線性直驅（LPO，不用 DSP 的光模組）是光互連的替代技術；比例 +10pt 使 FY2028 EPS −6.4% [本專案：mrvl/README.md 敏感度]。

## 近期重大事件
- 2026-10-06｜投資人日把 FY2028 營收展望上修到 US$20B、資料中心 18B；FY2029 客製逾 12B [本專案：邁威爾模型_v0.2.xlsx Recon!B6、Recon!B7、Recon!B9]；公司展望一年內由 13B 上修五次 [本專案：邁威爾模型_v0.2.xlsx Recon!J6]｜投資人日目標高於市場預期 [外部新聞：KSL]
- 2026-08-27｜FY2027 Q2 財報：營收 US$2.74B、資料中心 2.17B；Q3 指引 3.15B、毛利率 58% [本專案：邁威爾模型_v0.2.xlsx Q!D4、Q!D5、Q!E4、Q!E16]
- 2026-08-18｜授予 Google 認股權證：最多 5,897 萬股、行使價 US$206.58，多數依每 US$5 億客製營收分批歸屬，至 FY2033；涵蓋 TPU 相關客製晶片 [外部新聞：Verdict]｜模型暫以 CY2028 TPU 份額 5% 反映 [本專案：mrvl/README.md 最大限制]
- 2026-06-11｜宣布財務長交接（Dan Durn 6-15 上任）[外部新聞：Storage Newsletter]

## 需要觀察的事項
- FY2028 營收 US$20B 的可達性｜模型 18.5B，差 7.5%；要補上需資料中心 ×1.09 或客製 ×1.32 [本專案：邁威爾模型_v0.2.xlsx Recon!C6、Recon!D6、Recon!C12、Recon!C14]｜FY2027 Q3 財報（約 2026-11-30～12 月初）[本專案：mrvl/README.md 最大限制]
- 光互連規模｜光互連 ×1.2 使 FY2028 EPS +12.0%；每模組晶片組 ×0.8 為 −12.0% [本專案：mrvl/README.md 敏感度]｜同上
- 客製代購 HBM 比例與毛利率｜模型 FY2030 毛利率 48.4% vs 公司 FY2031 目標中點 57.5% [本專案：邁威爾模型_v0.2.xlsx Recon!C32、Recon!B32]｜同上
- Google TPU 份額（認股權證案）｜TPU 份額 +5pt 使 FY2028 EPS +8.4% [本專案：mrvl/README.md 敏感度]｜Google 下一代 TPU 供應商消息
- 校準係數｜k_x 0.80、k_o 1.44 為顆數、份額、單價的合併係數，無法分開識別 [本專案：邁威爾模型_v0.2.xlsx Units!C30、Units!C31]｜下一版模型

## 來源
1. 本專案：`mrvl/README.md`；`mrvl/dist/20261025_邁威爾模型_v0.2.xlsx`（In、Q、Units、PL、Recon）；`avgo/dist/20261025_Broadcom模型_v0.2.xlsx`（Units）；`alchip/dist/20261025_世芯-KY模型_v0.1.xlsx`（Units）；`mediatek/dist/20261025_聯發科模型_v0.1.xlsx`（Units）；AI半導體 `先驗比對_AVGO_20261023.md`
2. [Marvell raises 2028 revenue forecast on strong AI data center demand](https://www.ksl.com/article/51632987/marvell-raises-2028-revenue-forecast-on-strong-ai-data-center-demand)（2026-10）
3. [Marvell gives Google right to buy shares](https://www.verdict.co.uk/marvell-google-to-buy-shares/)（2026-08-20）
4. [Marvell FY2027 Q2 8-K](https://www.sec.gov/Archives/edgar/data/0001835632/000183563226000022/q227_8kx812026ex-991.htm)（2026-08-27）
5. [Marvell Announces CFO Transition](https://investor.marvell.com/news-events/press-releases/detail/1025/marvell-announces-cfo-transition)（2026-06-11）
