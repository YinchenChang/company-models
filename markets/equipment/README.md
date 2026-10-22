# markets/equipment — 半導體設備市場模型 v0.1

AI 半導體 Project 步驟 5（設備與材料）的設備部分。**設備需求＝為滿足記憶體、代工擴產所需的資本支出 × 設備占比。** 回答核心問題 1（循環在 2027 還是 2028 見頂）；核心問題 2、3 只做瓶頸表；公司份額與毛利留給公司層。

- 建置：`python3 markets/equipment/build_equipment.py raw.xlsx`，再以 LibreOffice headless 重算（`soffice --headless --calc --convert-to xlsx --outdir dist raw.xlsx`），錯誤格須為 0。成品：`dist/20261022_半導體設備市場模型_v0.1.xlsx`。CI：`.github/workflows/ai-semis-build.yml`（重建、錯誤格 0、逐格比對 dist）。
- 頁面：README、In（輸入、標記、兩個來源）、Mkt（SEMI 分項、記憶體、代工邏輯、測試封裝、三情境、摘要）、Recon、Sens、Sources。
- 取數：skhynix、micron、samsung（只記憶體）、tsmc 的 `dist/*.xlsx`「資本支出」列（三情境相同，建置時檢查）；不修改這些模型。

## 結果（十億美元；v0.1 建置時數值）

| 情境 | 2025（SEMI 實績） | 2026 | 2027 | 2028 | 2029 | 見頂年度 |
|---|---|---|---|---|---|---|
| A 續增（由下而上＋SEMI） | 134.7 | 176.3 | 223.4 | 251.9 | 265.8 | 2029（區間末，未見頂） |
| B 2028 持平 | 134.7 | 176.3 | 223.4 | 223.4 | 212.2 | 2028（與 2027 同水準） |
| C 2028 下滑約兩成 | 134.7 | 176.3 | 223.4 | 178.7 | 178.7 | 2027 |

年增：A 2027 +26.7%、2028 +12.8%、2029 +5.5%；B 2028 0%、2029 −5.0%；C 2028 −20%、2029 0%。2026 年增（+30.9%）含口徑轉換（2025＝SEMI 實績，2026 起由下而上），不可與後續年度直接比。

AI 相關設備（先進邏輯＋DRAM／HBM＋先進封裝）占總增量（相對 2026 累計）：A 2027 65%、2028 63%、2029 60%；B 65%、65%、68%；C 2027 65%，2028 起總增量接近 0，比值失真（見 Mkt 註記）。

**怎麼讀：** 見頂年度由情境係數決定，係數是 Assumed（B 的 2028 係數 0.95–1.05、C 的 0.75–0.85、2029 係數無實證），這張表是「若 B／C 則…」，不是預測。本模型本身不判定 2027 或 2028；它顯示：在資本支出沿用記憶體與台積電模型（A）時，設備總額到 2029 都還在增（增速降到 +5.5%，2029 外推為 Assumed）。

## 對帳（Recon；不為對上而調參）

| 項目 | 來源值 | 模型 A | 差距 |
|---|---|---|---|
| 2026 設備合計（SEMI） | 165.9 | 176.3 | +6.3% |
| 2026 晶圓前段（SEMI） | 143.9 | 154.3 | +7.2% |
| 2026 代工邏輯（SEMI） | 78.0 | 78.0 | 0%（結構性相等） |
| 2026 DRAM（SEMI） | 38.8 | 46.5 | **+19.8%** |
| 2026 NAND（SEMI） | 13.9 | 16.7 | **+19.8%** |
| 2027 晶圓前段年增（SEMI） | +21.8% | +28.3% | +6.5 個百分點 |
| 2027 晶圓前段（SEMI 連乘） | 175.3 | 197.9 | +12.9% |
| 2028 設備合計（SEMI） | 229.5 | 251.9 | +9.8% |
| 2026 晶圓前段（Lam 低 150B） | 150 | 154.3 | +2.9% |
| 2026 晶圓前段（KLA 低 150B，含先進封裝） | 150 | 154.3 | +2.9% |
| 2027 晶圓前段（KLA 約 190） | 190 | 197.9 | +4.2% |
| 2027 晶圓前段（Morgan Stanley） | 202 | 197.9 | −2.0% |
| 2028 晶圓前段（Morgan Stanley） | 227 | 222.5 | −2.0% |
| 2026 晶圓前段（Goldman Sachs） | 150 | 154.3 | +2.9% |
| 2027 晶圓前段（Goldman Sachs） | 218 | 197.9 | −9.2% |
| 2028 晶圓前段（Goldman Sachs） | 281 | 222.5 | **−20.8%** |

B、C 情境在 2028 對 MS／GS 的差距（B −12.8%／−29.6%，C −30.2%／−43.6%）是情境設計的結果。額外列示：Lam 2027「>200B」下限 −1.0%。

### 差距大於 15% 的項目與可能原因

1. **2026 DRAM、NAND 各 +19.8%。** 三家資本支出 98.6B × 64% ＝ 63.1B，已超過 SEMI 的 DRAM＋NAND 52.7B；「其他記憶體廠」殘差 −10.4B，採下限 0（所以記憶體總額是三家由下而上，是下限）。可能原因：（a）美光明言 FY27 資本支出多數為建築（無塵室），三家資本支出含大量非設備支出——隱含設備占比 48–53%，低於業界區間 58–70%；（b）三星記憶體資本支出為 Assumed（全公司 × 約 75%）；（c）資本支出是現金支出、SEMI 是設備商出貨營收，時點不同；（d）美光為財年（FY≈日曆年，未調整）；（e）SEMI 為 7 月預測，可能偏低。
2. **Goldman Sachs 2028 −20.8%。** GS 2028 年增 +29%（218→281），高於 SEMI +14.1%、Morgan Stanley 較早版本 +12%（191→215）與本模型 A +12.4%；可能納入不同口徑（先進封裝、廠務設備）或更激進的 DRAM／先進代工假設。
3. 未超過 15% 但值得注意：2027 前段 197.9 比 SEMI 推得的 175.3 高 12.9%，卻落在 KLA、Lam、Morgan Stanley 的 −2% 到 +4%、GS 的 −9% 之內，顯示 SEMI 7 月預測可能已落後之後的上修。
4. 本模型 A 在殘差為正的線上等於 SEMI（結構結果，非獨立驗證）；獨立檢驗只來自 Lam、KLA、Morgan Stanley、Goldman Sachs。
5. **Morgan Stanley 202／227 為提供的數字，搜尋只查到較早版本 149／191／215**；Lam 2026-10 法說另稱 2026「mid-150B」。

## 瓶頸表（核心問題 2、3；v0.1 只做表，不入模）

| 瓶頸 | 現況（找到的證據） | 來源 | 證據強度 |
|---|---|---|---|
| 無塵室空間 | 美光：FY27 資本支出增量一半以上、且多數資本支出為建築，用於提前取得無塵室；KLA：客戶受廠房就緒度與廠房外殼限制；SK 證券：記憶體供給成長受無塵室限制至 2027 | [Micron FQ3 FY26 準備稿](https://s25.q4cdn.com/621799436/files/doc_financials/2026/q3/Q3-FY26-Prepared-Remarks.pdf)；[Futurum](https://futurumgroup.com/insights/kla-corporation-q2-fy-2026-results-beat-on-ai-led-demand/)；[Cleanroom Technology](https://cleanroomtechnology.com/samsung-speeds-up-construction-of-p5-cleanroom-for) | 強（美光一手）／中／弱 |
| 長交期零組件 | 設備零組件（閥件管路、陶瓷、RF 電源、氣體箱）交期據報倍增、最長 40 個月；沉積設備零組件約 4→10 個月。應用材料、Lam 的一手說法：未找到 | [TrendForce 2026-09-18](https://www.trendforce.com/news/2026/09/18/news-chip-equipment-component-lead-times-reportedly-double-stretching-to-40-months-as-price-pressures-mount/)；[BigGo](https://finance.biggo.com/news/f46a8b91-046a-4708-87db-63b34e0862b2) | 中（「據報」）／弱 |
| EUV | 2026 低 NA 產能約 65 台、2027 再增約 30%；2027–28 需求超過供給（二手）；2027 台數各來源不一致（56／80／約 85） | [ASML 2026 Q2 6-K](https://www.sec.gov/Archives/edgar/data/0000937966/000162828026048235/pressreleasefinancialresul.htm)；[Pomegra](https://pomegra.io/news/asml-q3-2026-euv-demand-outruns-2027-capacity) | 強（產能，一手）／弱（供不應求，二手） |
| 量測 | KLA 2026 先進封裝製程控制營收約 $1.1B；交期 18–24 個月僅見社群貼文轉述；「量測產能是瓶頸」本身的證據：**未找到** | [X（KLA 法說摘要）](https://x.com/TheValueist/status/2082235400190640364)；Zacks／TradingView | 弱；瓶頸證據未找到 |
| HBM 鍵合機 | Hanmi 占 HBM TC 鍵合機營收 71.2%（TechInsights，2025 前三季）；SK 海力士 HBM4 約 50 組 TC 鍵合機、約半數 ASMPT；預期約 100 台訂單。供給短缺或交期：**未找到** | [TrendForce 2025-12-12](https://www.trendforce.com/news/2025/12/12/news-sk-hynix-reportedly-places-hbm4-tc-bonder-order-with-asmpt-as-hanmi-hanwha-patent-clash-heats-up/) | 中（集中度）；短缺未找到 |

## 限制

1. 三家記憶體資本支出 × 業界設備占比超過 SEMI 的 DRAM＋NAND 設備（見對帳 1）。設備占比只有業界合計值（SEMI WFE 116.9 ÷ Semiconductor Intelligence 資本支出 166 ＝ 70%；Rebound Capital 約 58%；基準取中點 64%，未找到記憶體／代工分項值）。
2. 2025 為 SEMI 實績：SK 海力士、三星記憶體 2025 資本支出未找到，其他模型亦未含。
3. 測試與封裝以 SEMI 為輸入，不由下而上：缺少每顆加速器的測試時間（最終測試、burn-in、SLT）與每條 CoWoS／HBM 堆疊產線的鍵合機台數資料。SEMI 2027 測試、封裝值未找到，以 2026、2028 幾何內插（Derived）。
4. SEMI 2029 無預測；2029 以 g29（Assumed，0–12%，基準 8%）外推。
5. 台積電資本支出含先進封裝與光罩（公司稱 10–20%），與 SEMI 後段分類可能重疊；未另扣。
6. 取數第二來源：SEMI 新聞稿的轉載不是獨立來源；SEMI 網頁本身在本環境無法直接開啟，分項年增取自搜尋摘要。各輸入的兩個來源見 In 頁，找不到者寫「未找到」。SemiAnalysis 未使用。
