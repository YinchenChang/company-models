---
ticker: VRT
name: Vertiv
name_full: Vertiv Holdings Co
listing: 紐約證券交易所｜VRT [外部・未經本專案檢核]
hq: 美國俄亥俄州 [外部・未經本專案檢核]
founded: 本專案未建立
employees: 本專案未建立
fiscal_year_end: 12 月底
asof: 2026-11-17
---

## 一句話定位
AI 資料中心「機櫃外」的電力與散熱設備商：賣不斷電系統（UPS）、配電、冷卻分配單元（CDU）與冷水機，再加上安裝與維運服務；營收隨全球資料中心新增容量與 Vertiv 的份額變動 [本專案：vrt/README.md 開頭；Units!A2:A6]。

## 公司在做什麼
- 把電網的電送進機房並確保不斷電：UPS（停電時的備援電源）、配電櫃（把電分到各機櫃）[本專案：vrt/dist/20261117_Vertiv個股模型_v0.1.0.xlsx Units!A4:A5]。
- 把機房的熱帶走：冷卻分配單元（CDU，液冷系統的泵與熱交換）、冷水機、機房空調（CRAH）[本專案：同上 Units!A6]。
- 邊界：機櫃內的電源、冷板、備援電池屬 IT 設備端（本專案由台達電模型負責），不在 Vertiv 的可及市場內 [本專案：同上 Units!A2]。
- 另有服務與備品（安裝、維修、備品），公司 10-Q 把營收分為產品與服務備品兩類 [本專案：同上 Sources!A4 對應的 3Q26 10-Q 分類]。
- 地區分部：美洲、亞太、歐洲中東非洲 [Verified：Vertiv 新聞稿；本專案：同上 PL!A167:A169]。

## 主要產品與營收占比
| 產品／部門 | 內容（白話） | 營收占比 | 期間 | 標記 |
|---|---|---|---|---|
| 美洲 | 北美為主的機房與商用設備 | 約 62% | FY2025 | [Derived：6.386 ÷ 10.230；本專案：同上 PL!E167、Q!B4] |
| 亞太 | 中國與亞洲其他地區 | 約 20% | FY2025 | [Derived：2.019 ÷ 10.230；本專案：同上 PL!E168、Q!B4] |
| 歐洲中東非洲 | 歐洲、中東與非洲 | 約 18% | FY2025 | [Derived：1.824 ÷ 10.230；本專案：同上 PL!E169、Q!B4] |

- 產品線（電力、熱管理、服務）營收公司未揭露：本專案未建立官方占比；模型以「資料中心設施產品」占營收約 72%（FY2026）推導，屬 [Assumed：本專案 In!E35]。
- 營收由 FY2025 約 10.2 十億美元升到 FY2026 錨點約 14.6 十億美元，年增約 43% [本專案：同上 Q!B4、Q!G4]。
- 毛利率 FY2025 約 36.3%、FY2026 錨點約 38.1%；GAAP 營業利益率約 17.9%、21.1% [本專案：同上 Q!B6、G6、B9、G9]。

## 市場定位與產業位置
- 本專案把 Vertiv 營收讀成「Wattonomics 機房設施市場（配電、UPS、冷卻）× Vertiv 份額」；FY2026 份額約 10.5%（FY2025 約 15.2%，差異主因是市場以裝機年度計支出、Vertiv 於出貨認列）[本專案：同上 Units!C12:D12，Derived]。
- 可及市場 FY2026 約 100.3 十億美元，FY2027 成長約 32%，FY2029 成長僅約 0.9%（Wattonomics 基準）[本專案：同上 Units!C8、D10、F10，Derived]。
- 訂單支持：FY2025 底 backlog 約 150 億美元 [Verified：Vertiv 4Q25 新聞稿；本專案：同上 Sources!A6、Units 份額備註]；3Q26 backlog 本專案未找到。
- 客戶集中度：本專案未建立（見「主要客戶」）。

## 管理層與經營風格
| 職位 | 姓名 | 任職起 | 背景 |
|---|---|---|---|
| 執行長 | Giordano Albertazzi | 本專案未建立 | 3Q26 新聞稿以其名義發言 [外部・未經本專案檢核：Vertiv 3Q26 新聞稿，2026-10-28] |

- 資本配置、債務與股利細節：本專案未在此檔彙整，見模型 DCF 與 Sources 頁（3Q26 10-Q）。
- 併購為主要成長補充：2026 年完成 ThermoKey，並於 2026-10-26 交割 UtilityInnovation [Verified：Vertiv 3Q26 10-Q；外部新聞：unite.ai 2026-09-02]。
- 經營風格細節（激勵、治理）本專案尚未研究。

## 同業與競爭態勢
| 同業 | 重疊領域 | 市占或規模 | 標記 |
|---|---|---|---|
| Schneider Electric | 配電、UPS、冷卻 | 本專案未建立 | [本專案：同上 Val 頁同業區塊（ETN、SU.PA、2308.TW）] |
| Eaton | 配電、UPS | 本專案未建立 | [本專案：同上 Val 頁同業區塊（ETN、SU.PA、2308.TW）] |
| 台達電 | 機櫃內電源與液冷；部分設施電力 | 本專案未建立 | [本專案：delta/README.md；同上 Val 頁 0.2 同業表] |

- 本專案的同業前瞻本益比中位數約 23.3 倍 [本專案：同上 Val!B41，Interested-party]。
- 競爭重點在 800V 直流與液冷的產品時程；機櫃內外邊界移動會改變可及市場，本專案以 Wattonomics 的邊界定義為準 [本專案：同上 Units!A2；判斷]。

## 主要客戶
- 財報揭露的客戶集中度：未找到。已讀 3Q26 10-Q 的損益、資產負債表、債務與期後事項，未見客戶集中度揭露；建議核對 FY2025 年報第 1 章與合併財報附註 [外部・未經本專案檢核：Vertiv 3Q26 10-Q，2026-10-28]。
| 客戶 | 關係 | 依據 | 標記 |
|---|---|---|---|
| 未具名超大型雲端業者與營運商 | 公司稱資料中心需求為主 | 新聞稿未列客戶名 | [外部・未經本專案檢核：Vertiv 3Q26 新聞稿，2026-10-28] |

## 前十大股東
| 排名 | 股東 | 持股比例 | 標記 |
|---|---|---|---|
| 1 | BlackRock | 8.10% | [外部・未經本專案檢核：Yahoo Finance 持股頁，資料日 2026-09-30] |
| 2 | Vanguard Capital Management | 6.54% | [外部・未經本專案檢核：Yahoo Finance 持股頁，資料日 2026-09-30] |
| 3 | Vanguard Portfolio Management | 4.74% | [外部・未經本專案檢核：Yahoo Finance 持股頁，資料日 2026-09-30] |
| 4 | State Street | 4.31% | [外部・未經本專案檢核：Yahoo Finance 持股頁，資料日 2026-09-30] |
| 5 | JPMorgan Chase | 2.85% | [外部・未經本專案檢核：Yahoo Finance 持股頁，資料日 2026-09-30] |
| 6 | T. Rowe Price | 2.37% | [外部・未經本專案檢核：Yahoo Finance 持股頁，資料日 2026-06-30] |
| 7 | Geode Capital Management | 2.37% | [外部・未經本專案檢核：Yahoo Finance 持股頁，資料日 2026-09-30] |
| 8 | Bank of America | 2.15% | [外部・未經本專案檢核：Yahoo Finance 持股頁，資料日 2026-06-30] |
| 9 | Morgan Stanley | 2.13% | [外部・未經本專案檢核：Yahoo Finance 持股頁，資料日 2026-09-30] |
| 10 | Norges Bank | 1.26% | [外部・未經本專案檢核：Yahoo Finance 持股頁，資料日 2026-06-30] |

- 口徑：13F 機構持股（占流通股比例），各家申報日不同（2026-06-30 或 2026-09-30）；機構間可能重複計入同一集團的不同帳戶 [外部・未經本專案檢核：Yahoo Finance 持股頁]。

## 近期重大事件
- 2026-10-28｜3Q26 財報：營收 4,312.8 百萬美元，上調全年指引 [本專案：同上 Q!E4；外部新聞：Vertiv 新聞稿（SEC 8-K 附件 99.1）]｜本專案錨點採用 4Q26 指引中位 [本專案：同上 Q!F4]。
- 2026-10-26｜UtilityInnovation 收購交割：現金約 14.5 億美元，另有最多 11.5 億美元 EBITDA 對價 [外部新聞：unite.ai 2026-09-02 轉述公告]｜屬微電網與電力設備，本專案基準視為價值中性 [本專案：同上 In!E103 與 Sens]。
- 2026-09-02｜公告收購 UtilityInnovation Group [外部新聞：unite.ai 2026-09-02]｜補強電網受限地區的資料中心供電。
- 2026-07-29｜2Q26 財報：營收 3,274.3 百萬美元 [本專案：同上 Q!D4；外部新聞：Vertiv 新聞稿（SEC 8-K 附件 99.1）]｜與 3Q26 合計構成模型 FY2026 錨點。
- 2026-06-11｜ThermoKey S.p.A. 併購完成 [Verified：Vertiv 3Q26 10-Q]｜強化熱管理產品線。

## 需要觀察的事項
- 市場延續性｜Wattonomics 基準下設施市場在 FY2029 後持平（FY2029 年增約 0.9%），本專案 FY2030 起用收斂法，是目標價最大不確定來源 [本專案：同上 Units!F10、Sens 業務主因]｜Wattonomics 版本更新時。
- 份額與認列時差｜FY2026 份額較 FY2025 低，若 backlog 轉營收加速，份額相對值上調 [本專案：同上 In!E36（rel_27）]｜3Q26 後每季財報。
- 折現率與估值落差｜WACC 約 13.3%，DCF 約 81 美元，同業倍數法調整後約 202 美元，現價 202.07 美元 [本專案：同上 Val!B47、B48、In!E68，Derived]｜Andy 對 β 與 ERP 的判斷。
- 4Q26 實績對指引｜本專案 4Q26 營收假設約 4,413 百萬美元（取管理層指引中位） [本專案：同上 In!E51；Interested-party：管理層指引]｜約 2027-02（日期未公布）。
- UtilityInnovation 整合與 EBITDA 對價｜最多 11.5 億美元對價與其貢獻營收，本專案未建模 [本專案：同上 Sens 業務主因表（收購現金）]｜2027 年起各季。

## 來源
1. [Vertiv 3Q26 新聞稿（SEC 8-K 附件 99.1）](https://www.sec.gov/Archives/edgar/data/0001674101/000162828026067707/q32026exhibit991vrt10282026.htm)（2026-10-28）
2. [Vertiv 3Q26 10-Q](https://www.sec.gov/Archives/edgar/data/1674101/000162828026067893/vrt-20260930.htm)（2026-09-30 季）
3. [Vertiv 2Q26 新聞稿（SEC 8-K 附件 99.1）](https://www.sec.gov/Archives/edgar/data/0001674101/000162828026050323/q22026exhibit991vrt07292026.htm)（2026-07-29）
4. [Vertiv 4Q25／FY2025 新聞稿](https://s205.q4cdn.com/554782763/files/doc_financials/2025/q4/Vertiv_Fourth-Quarter-2025-Earnings-Release.pdf)（2026-02）
5. [Yahoo Finance：VRT 持股](https://finance.yahoo.com/quote/vrt/holders)（資料日 2026-09-30）
6. [unite.ai：Vertiv 收購 UtilityInnovation Group](https://www.unite.ai/vertiv-strikes-1-45b-deal-for-microgrid-firm-utilityinnovation-group/)（2026-09-02）
