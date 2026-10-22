# markets/equipment — 半導體設備市場模型 v0.1

AI 半導體 Project 步驟 5（設備與材料）的設備部分。**設備需求＝為滿足記憶體、代工擴產所需的資本支出 × 設備占比。** 回答核心問題 1（循環在 2027 還是 2028 見頂）；核心問題 2、3 只做瓶頸表；公司份額與毛利留給公司層。

- 建置：`python3 markets/equipment/build_equipment.py raw.xlsx`，再以 LibreOffice headless 重算（`soffice --headless --calc --convert-to xlsx --outdir dist raw.xlsx`），錯誤格須為 0。成品：`dist/20261022_半導體設備市場模型_v0.1.xlsx`。CI：`.github/workflows/ai-semis-build.yml`（重建、錯誤格 0、逐格比對 dist）。
- 頁面：README、In（輸入、標記、兩個來源）、Mkt（SEMI 分項、記憶體、代工邏輯、測試封裝、三情境、摘要）、Recon、Sens、Sources。
- 取數：skhynix、micron、samsung（只記憶體）、tsmc 的 `dist/*.xlsx`「資本支出」列（三情境相同，建置時檢查）；不修改這些模型。

## 結果（十億美元；v0.1 建置時數值）

| 情境 | 2025（SEMI 實績） | 2026 | 2027 | 2028 | 2029 | 見頂年度 |
|---|---|---|---|---|---|---|
| A 續增 | 134.7 | 165.9 | 200.7 | 229.4 | 247.7 | 2029（區間末，未見頂） |
| B 2028 持平 | 134.7 | 165.9 | 200.7 | 200.7 | 190.7 | 2028（與 2027 同水準） |
| C 2028 下滑約兩成 | 134.7 | 165.9 | 200.7 | 160.6 | 160.6 | 2027 |

年增：A 2026 +23.2%、2027 +21.0%、2028 +14.3%、2029 +8.0%（2029 為 g29，Assumed）；B 2028 0%、2029 −5.0%；C 2028 −20%、2029 0%。

**A 路徑的總額由 SEMI 決定**：「其他記憶體廠」「其他代工邏輯」定義為 SEMI 減已知廠的殘差，所以記憶體與代工邏輯合計≡SEMI。公司資本支出 × 設備占比只決定「三家 vs 其他」的分配與 AI 分類，不決定總額。

AI 相關設備（先進邏輯＋DRAM／HBM＋先進封裝）占總增量（相對 2026 累計）：A 2027 63%、2028 62%、2029 59%；B 63%／63%／65%；C 2027 63%，2028 起總增量接近 0，比值失真（見 Mkt 註記）。

**怎麼讀：** 見頂年度由情境係數決定，係數是 Assumed（B 2028 係數 0.95–1.05、C 0.75–0.85、2029 係數無實證），這張表是「若 B／C 則…」，不是預測。

## 設備占資本支出比例（按公司）

| 公司 | 低／基準／高 | 依據 | 標記 |
|---|---|---|---|
| 美光 FY26 | 47.1%／52.1%／57.1% | 公司揭露推導（見下） | Derived |
| 美光 FY27–FY29 | 38.5%／46.0%／50.0% | 公司揭露推導；高＝50% 為指定上限 | Derived |
| SK 海力士 | 58%／64%／70% | **未找到公司揭露**；業界值 | Analogy |
| 三星（記憶體） | 58%／64%／70% | **未找到公司揭露**；業界值 | Analogy |
| 台積電 | 58%／64%／70% | **未找到公司揭露**；業界值 | Analogy |

業界值：SEMI 2025 WFE 116.9 ÷ Semiconductor Intelligence 2025 資本支出 166 ＝ 70%；Rebound Capital 約 58%；基準取中點。

**美光推導與一項更正：** 公司原文是資本支出「**增量**」多數為建築，不是「FY27 資本支出多數為建築」：FQ3 FY26 稿「more than half the increase year over year in fiscal 2027 from construction capex as we pull in cleanroom capacity…」；FQ4 法說「A majority of the increase is for construction capex, most of which is to help accelerate clean room space availability in late calendar 2028 and beyond」。（引文取自搜尋摘要，未能開啟原文件。）文稿因此不直接給出設備占比。推導：FY25 淨資本支出 13.80（8-K）× (1 − 64%) 為建築基數；其後各年增量中建築占 c＝50%／60%／70%（50% 為「多於一半」的上限情境；60%、70% 為 Assumed）。c＝50% 時 FY27 設備占比約 53%，指定的 50% 上限比它更嚴。早先版本的 README 與 PR 說明寫成「FY27 資本支出多數為建築」，已更正。

SK 海力士、三星：查到的只有定性資訊（SK 海力士董事會 2026-02 核准 21.6 兆韓元完成龍仁第一晶圓廠結構工程與二至六期無塵室；三星 P5 無塵室先行），沒有年度建築／設備比例。

## 其他記憶體廠殘差（發現）

「其他記憶體廠」＝SEMI DRAM＋NAND − 三家設備，照實顯示、不取下限：2026–2029＝−7.2／−12.8／−11.7／−7.2 十億美元。三家資本支出 × 設備占比（加權 57–61%）已超過 SEMI 的 DRAM＋NAND 總額，這本身是一項發現。可能原因：資本支出含建築、資本支出與設備認列的時差、SEMI 預測偏低、口徑差異。Mkt 的「SEMI 隱含設備占比」（2026–2029：53.4%／47.9%／49.6%／52.7%）只是診斷值。

## 對帳（Recon；不為對上而調參）

| 項目 | 來源值 | 模型 A | 差距 |
|---|---|---|---|
| 2026 設備合計／前段／代工邏輯／DRAM／NAND（SEMI） | 165.9／143.9／78.0／38.8／13.9 | 同左 | 0%（結構性相等） |
| 2027 晶圓前段年增（SEMI） | +21.8% | +21.8% | 0 |
| 2028 設備合計（SEMI） | 229.5 | 229.4 | −0.1%（捨入） |
| 2026 晶圓前段（Lam 低 150B） | 150 | 143.9 | −4.1% |
| 2026 晶圓前段（KLA 低 150B，含先進封裝） | 150 | 143.9 | −4.1% |
| 2027 晶圓前段（KLA 約 190） | 190 | 175.3 | −7.8% |
| 2027 晶圓前段（Lam >200B，額外） | 200 | 175.3 | −12.4% |
| 2027 晶圓前段（Morgan Stanley） | 202 | 175.3 | −13.2% |
| 2028 晶圓前段（Morgan Stanley） | 227 | 200.0 | −11.9% |
| 2026 晶圓前段（Goldman Sachs） | 150 | 143.9 | −4.1% |
| 2027 晶圓前段（Goldman Sachs） | 218 | 175.3 | **−19.6%** |
| 2028 晶圓前段（Goldman Sachs） | 281 | 200.0 | **−28.8%** |

Morgan Stanley 202／227：[itiger 二手轉述（2026-07-06；連結由使用者提供，本環境無法開啟，未直接核對頁面）](https://www.itiger.com/news/2649678863)，內容為 2027 由 191 上修到 202、2028 由 215 上修到 227；上修前 191／215 與 [Investing.com 報導](https://www.investing.com/news/stock-market-news/semicap-stocks-morgan-stanley-raises-wfe-outlook-and-shifts-ratings-4696580)一致。標 Interested-party、二手轉述。

### 差距大於 15% 的項目與可能原因

1. **Goldman Sachs 2027 −19.6%、2028 −28.8%。** GS 2027 年增約 +45%（150→218）、2028 約 +29%，高於 SEMI 的 +21.8%、+14.1%；GS 是最高的一組，可能納入不同口徑（KLA 口徑含先進封裝；廠務設備）、更激進的 DRAM／先進代工假設，且發布較晚（2026-08，SEMI 為 2026-07）。
2. 未超過 15% 但方向一致：2027、2028 均低於 KLA、Lam、Morgan Stanley（−8% 到 −13%），顯示 SEMI 7 月預測（A 路徑天花板）可能落後之後的上修。
3. A 在殘差定義的線上等於 SEMI（結構結果，非獨立驗證）；獨立檢驗只來自 Lam、KLA、Morgan Stanley、Goldman Sachs。B、C 在 2028 的差距是情境設計的結果。

## 瓶頸表（核心問題 2、3；v0.1 只做表，不入模）

| 瓶頸 | 現況（找到的證據） | 來源 | 證據強度 |
|---|---|---|---|
| 無塵室空間 | 美光：FY27 資本支出『增量』一半以上為建築（FQ3 稿）、『增量』多數為建築（FQ4 法說）；KLA：客戶受廠房就緒度與廠房外殼限制；SK 證券：記憶體供給成長受無塵室限制至 2027 | [Micron FQ3 FY26 準備稿](https://s25.q4cdn.com/621799436/files/doc_financials/2026/q3/Q3-FY26-Prepared-Remarks.pdf)；[Futurum](https://futurumgroup.com/insights/kla-corporation-q2-fy-2026-results-beat-on-ai-led-demand/)；[Cleanroom Technology](https://cleanroomtechnology.com/samsung-speeds-up-construction-of-p5-cleanroom-for) | 強（美光一手）／中／弱 |
| 長交期零組件 | 設備零組件（閥件管路、陶瓷、RF 電源、氣體箱）交期據報倍增、最長 40 個月；沉積設備零組件約 4→10 個月。應用材料、Lam 的一手說法：未找到 | [TrendForce 2026-09-18](https://www.trendforce.com/news/2026/09/18/news-chip-equipment-component-lead-times-reportedly-double-stretching-to-40-months-as-price-pressures-mount/)；[BigGo](https://finance.biggo.com/news/f46a8b91-046a-4708-87db-63b34e0862b2) | 中（「據報」）／弱 |
| EUV | 2026 低 NA 產能約 65 台、2027 再增約 30%；2027–28 需求超過供給（二手）；2027 台數各來源不一致（56／80／約 85） | [ASML 2026 Q2 6-K](https://www.sec.gov/Archives/edgar/data/0000937966/000162828026048235/pressreleasefinancialresul.htm)；[Pomegra](https://pomegra.io/news/asml-q3-2026-euv-demand-outruns-2027-capacity) | 強（產能，一手）／弱（供不應求，二手） |
| 量測 | KLA 2026 先進封裝製程控制營收約 $1.1B；交期 18–24 個月僅見社群貼文轉述；「量測產能是瓶頸」本身的證據：**未找到** | [X（KLA 法說摘要）](https://x.com/TheValueist/status/2082235400190640364)；Zacks／TradingView | 弱；瓶頸證據未找到 |
| HBM 鍵合機 | Hanmi 占 HBM TC 鍵合機營收 71.2%（TechInsights，2025 前三季）；SK 海力士 HBM4 約 50 組 TC 鍵合機、約半數 ASMPT；預期約 100 台訂單。供給短缺或交期：**未找到** | [TrendForce 2025-12-12](https://www.trendforce.com/news/2025/12/12/news-sk-hynix-reportedly-places-hbm4-tc-bonder-order-with-asmpt-as-hanmi-hanwha-patent-clash-heats-up/) | 中（集中度）；短缺未找到 |

## 限制

1. 其他記憶體廠殘差為負，照實顯示（見上）。設備占比除美光外只有業界合計值（70% 與 58%，基準 64%）。
2. 2025 為 SEMI 實績：SK 海力士、三星記憶體 2025 資本支出未找到，其他模型亦未含；2026 年增率含口徑轉換。
3. 測試與封裝以 SEMI 為輸入，不由下而上：缺少每顆加速器的測試時間（最終測試、burn-in、SLT）與每條 CoWoS／HBM 堆疊產線的鍵合機台數資料。SEMI 2027 測試、封裝值未找到，以 2026、2028 幾何內插（Derived）。
4. SEMI 2029 無預測；2029 以 g29（Assumed，0–12%，基準 8%）外推。
5. 台積電資本支出含先進封裝與光罩（公司稱 10–20%），與 SEMI 後段分類可能重疊；未另扣。
6. **時差：資本支出是現金支出，設備營收是出貨認列，兩者約差一到兩季；v0.1 沒有建模（年度金額直接相乘）。**
7. SEMI 新聞稿的轉載不是獨立來源；SEMI 網頁本身在本環境無法直接開啟，分項年增取自搜尋摘要。各輸入的兩個來源見 In 頁，找不到者寫「未找到」。SemiAnalysis 未使用。
