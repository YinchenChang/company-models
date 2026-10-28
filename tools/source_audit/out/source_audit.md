# 來源複核結果

- 應核對 2380 列，已核對 836 列；累計花費約 40.03 美元（上限 40.0）；搜尋 754 次、擷取 722 次。

- 不符：7
- 部分相符：148
- 來源打不開：88
- 找不到來源：97
- 相符：496

## 不符與部分相符

| 判定 | 模型 | 鍵 | 項目 | 基準 | 來源上的數字 | 說明 |
|---|---|---|---|---|---|---|
| 不符 | advantest | mk_soc26 | SoC 測試機市場 CY2026（公司上修後） | 11 US$B | SoC テスタ市場 CY2026 予測レンジ 8.7–9.5 billion USD(1月時點は 8.5–9.5) | 已開啟公司原文:4/27 簡報寫 8.7–9.5,非 10.5–11.5。10.5–11.5 只見於 X 上的二手貼文,似為 7 月之後的上修,未經公司原文核實。 |
| 不符 | ibiden | pe_fwd_sa | 前瞻 P/E（stockanalysis） | 53.97 x | Forward PE 46.82(價格 ¥8,512,2026-10-23 收盤;頁面更新 2026-10-24) | 頁面前瞻 P/E 為 46.82x,模型 53.97x 高約 15%。此值隨股價每日變動,差異可能來自取值日不同。 |
| 不符 | kinsus | sh_cpu_27 | 景碩占 NVIDIA CPU（Grace／Vera）載板 2027 | 0.7 % | 經濟日報 2026-09-16：「2027年Vera CPU市占率將由Grace的50%提升為80%至100%」；TechNews 2026-08-04：目標奪得 Vera CPU 載板 100% 市占率 | 兩來源都高於模型基準0.7：經濟日報為80–100%，TechNews為100%目標。區間0.5～0.9只涵蓋80–90%。其他投顧僅估超過50%，不在標示來源內。 |
| 不符 | siliconomics | ts_mk_soc26 | SoC 測試機市場 CY2026（公司上修後） | 11 US$B | CY2026 SoC tester market size forecast range 8.7–9.5 billion USD(1月預測 8.5–9.5,下限上修 0.2) | 公司原文p.19寫SoC測試機CY2026為87～95億美元,非10.5～11.5;基準11也不符(中點約9.1)。 |
| 不符 | skhynix | nb_1 | NAND 位元成長 第 1 年 | 0.2 % | calendar 2027 and 2028, industry NAND bit shipments to grow approximately in the mid-20s percentage range | Micron 寫 2027 為中 20% 段（約 25%），模型 0.20 偏低；0.20 較接近 2026 的 low-20s。標 Verified 不當。 |
| 不符 | skhynix | nb_2 | NAND 位元成長 第 2 年 | 0.2 % | calendar 2027 and 2028, industry NAND bit shipments to grow approximately in the mid-20s percentage range | Micron 對 2028 同為中 20% 段（約 25%），模型 0.20 偏低；區間上緣 0.25 才接近來源。標 Verified 不當。 |
| 不符 | skhynix | q3_nb | Q3 NAND 位元季增 | 0.02 % | NAND bit shipments declined slightly from the previous quarter | 公司Q3簡報稱NAND位元出貨小幅下滑，模型+2%（區間0~4%）方向相反；公司原文未打開，為簡報摘要 |
| 部分相符 | accton | nap25s | FY2025 網路應用產品（AI 加速卡等）÷ 營收 | 0.63 % | 網路應用設備 營收佔比 63%，年增 361%（2025 前三季） | 63% 與模型值相符，但來源是前三季佔比，模型標的是 FY2025 全年。fetch 工具拒絕開啟該網址，內容取自搜尋結果摘錄。 |
| 部分相符 | accton | sw25s | FY2025 網路交換器 ÷ 營收 | 0.34 % | 網路交換器 營收佔比 34%，年增 46%（2025 前三季） | 34% 與 0.34 一致，但來源是前三季佔比，非 FY2025 全年；口徑不同。 |
| 部分相符 | alchip | q3_gm | 3Q26 毛利率（公司：低至中 20%） | 0.21 % | for the third quarter and the fourth quarter, the blended gross margin will most likely be about 20 something percent | 逐字稿只有「20幾%、低20%」的定性區間，0.21落在區間內，但來源沒有點值 |
| 部分相符 | alchip | q3_opx | 3Q26 營業費用（公司：與 2Q26 相近） | 1.05 NT$B | for the third quarter and the fourth quarter, the operating expense will be pretty similar with the second quarter | 「與2Q26相近」屬實，但逐字稿以美元報價且未給營業費用金額，NT$1.05B無法直接核對 |
| 部分相符 | amazon | dps | 每股股利 | 0 US$ | We have never declared or paid cash dividends on our common stock. | S02未定義;FY2017等舊10-K明寫從未派息;FY2025 10-K股利段被截斷未讀到 |
| 部分相符 | amazon | kd_new | 新增借款利率（2026-07 新發 2066 票據有效利率；只用於利息費用） | 0.0633 % | 6.250% notes due 2066；定價殖利率 6.297%（99.316% 發行價） | 只找到票面 6.25% 與殖利率 6.297%，未見 6.33% 有效利率原文；量級一致。 |
| 部分相符 | amazon | p_list | AWS p5.48xlarge 隨需（每 GPU 小時） | 6.88 US$/GPU-hr | p5.48xlarge on-demand $55.0400/hr(us-east-1),÷8=$6.88/GPU-hr | 多個價格彙整站皆為$55.04/hr,換算6.88吻合。AWS官方公告只列降幅44%,未見美元價,官方價目頁未能開啟。 |
| 部分相符 | amazon | p_spot | H100 現貨價（Silicon Data 指數） | 2.82 US$/GPU-hr | As of Oct 26, 2026: 2.89 USD / GPU-hour(Neo-Cloud SDH100RT) | 官方頁現值2.89,差約2.4%。2.82為10/10讀數(第三方轉述)。頁內FAQ另寫2.53,前後不一致。 |
| 部分相符 | amkr | az_inc | 政府補助與投資抵減 ÷ 投資（降低折舊基礎） | 0.35 % | CHIPS 補助 $400 million + 35% investment tax credit,合計 $2.8 billion;Arizona 投資 $7 billion | 逐字稿確有 35% 抵減與 $2.8B／$7B,但合計比率為 2.8÷7=0.40,基準 0.35 只等於抵減率,0.40 落在區間上限。 |
| 部分相符 | avc | q_gm | 3Q、4Q26 毛利率 | 0.33 % | 2Q26 合併毛利率 32.57%，季增 2.8 個百分點 | 2Q26 32.57% 已證實。找不到工商時報 8/6「3Q 朝 33%」原文，經濟日報同日報導只寫下半年毛利率提升，區間也無來源。 |
| 部分相符 | avc | srv_25 | 2025 年伺服器與網通 ÷ 營收 | 0.52 % | 伺服器(52%)與網通(7%)；圖3 2025年伺服器營收比重 52 | 52%僅為伺服器，網通另列7%（合計約59%）；核對的是8/5報告，未找到8/12版，1H25 48%未見 |
| 部分相符 | besi | cons_eps26 | 共識 EPS FY2026 | 4.334 € | EPS 2026: 4,23 EUR | 量級一致，來源4.23比4.334低約2.4%。表格取自搜尋摘要，快照日期未能確認為10/9。 |
| 部分相符 | besi | cons_eps28 | 共識 EPS FY2028 | 8.483 € | EPS 2028: 8,80 EUR | 方向一致，8.80比8.483高約3.7%，超出2%。快照日期未能確認為10/9。 |
| 部分相符 | besi | cons_rev26 | 共識營收 FY2026 | 1002 €M | Net sales 2026：1B（EUR，頁面四捨五入顯示）；EV 11.61B ÷ EV/Sales 2026 11.6x ≈ 1.00B | 量級一致，但頁面只顯示約1B，無法確認1002；頁面為10/26快照，非10/09；財務頁數字被截斷。 |
| 部分相符 | besi | cons_rev27 | 共識營收 FY2027 | 1345 €M | Net sales 2027*: 1.34B EUR（四捨五入顯示） | 1.34B 與 1345 差約 0.4%。但這是 10/27 左右（Q3 後、股價 163.05）的快照，查不到 10/9 財報前版本。 |
| 部分相符 | besi | tax | 有效稅率 | 0.135 % | Effective tax rate 13.1%（Q2-26） | Q2-26 的 13.1% 已在新聞稿證實；Q3-26 的 13.3% 因新聞稿 PDF 無權限讀取而未核實。模型值 0.135 略高於兩者。 |
| 部分相符 | dell | g_q3rev | 3Q FY27 營收指引（參考） | 49 US$B | Q3FY27 Revenue $49.0 (billions), 81% Y/Y | 基準49十億美元與8-K一致；來源只給單一點估計，未提供48.5～49.5區間，區間無從核對。 |
| 部分相符 | delta | ai_26 | 2026 年 AI 相關營收 ÷ 營收（讀為不含液冷） | 0.26 % | AI-Related Revenue Share: Over 25% this year (not disclosed last year), target 28%. | 法說原稿PDF打不開；BigGo只寫>25%、目標28%，無26%，0.26落在區間內，未見是否含液冷。 |
| 部分相符 | delta | lc_26 | 2026 年液冷營收 ÷ 營收 | 0.125 % | Liquid-cooling products accounted for approximately 10% of total revenue last year and will definitely exceed 12% this year. | CFO只說2026年液冷占比「超過12%」（下限）。基準12.5%與區間上限13%是模型自行假設，來源未提。 |
| 部分相符 | disco | mix_eq | 精密加工裝置占營收 FY2025 | 0.62 % | 精密加工装置合計 62%（Full Year 構成比，出荷額ベース） | p.29 通期62%與0.62一致，子項31/28/3、21、16亦同；但來源為出荷額（出貨額）基礎，非營收，口徑不同。 |
| 部分相符 | disco | mix_tl | 精密加工ツール（消耗品：刀片、砂輪）占營收 FY2025 | 0.21 % | 精密加工ツール 構成比 21%（Full Year，出荷額ベース） | 第29頁全年構成比21%與模型0.21一致，但來源為出荷額（出貨額）基礎，非營收占比，口徑不同。 |
| 部分相符 | foxconn | ais_26 | 2026 年 AI 伺服器 ÷ 雲端網路產品 | 0.6 % | AI 伺服器營收已占鴻海第一季雲端網路產品營收超過 50% | 來源只給 1Q26「超過50%」，未給2026全年0.6；方向一致，但口徑為單季下限值。 |
| 部分相符 | foxconn | cl_26 | 2026 年雲端網路產品 ÷ 營收 | 0.53 % | 2Q26 雲端網路產品佔比 51%（首次突破五成）；3Q26 預期季對季、年對年強勁成長 | 來源只有 2Q26 51% 與 3Q 成長展望，沒有 2026 全年佔比。0.53 與走勢相符，但屬推算值，無法直接核對。 |
| 部分相符 | foxconn | pc_26 | 2026 年電腦終端產品 ÷ 營收 | 0.12 % | 電腦終端產品：營收佔比 15%（2Q26）；3Q26 預期較 2Q 略為衰退 | 來源只有 2Q26 單季佔比 15%，沒有 2026 全年佔比，也沒找到 1Q26 的 14%。模型 0.12 低於單季值，屬推估。 |
| 部分相符 | gce | net26 | 2026 年網通營收 ÷ 營收 | 0.15 % | 網通約占15%比重 | 文章日期 2026/8/28，數字 15% 與 0.15 一致；但原文是「現階段」占比，不是 2026 全年口徑。 |
| 部分相符 | gce | q1_rev | 1Q26 營收（實績） | 19.313 NT$B | 19,312,774 仟元(約 193.13 億元) | 標示來源頁只顯示年度/TTM,無1Q26季資料;財報狗與鉅亨網報導皆為193.13億,與19.313相符。 |
| 部分相符 | gce | srv26 | 2026 年伺服器營收 ÷ 營收 | 0.78 % | 伺服器營收占比達78% | 文章日期 2026/8/28，數字 78% 與 0.78 一致；但原文是「現階段」占比，不是 2026 全年口徑。 |
| 部分相符 | guc | att26 | CY2026 CPU 對 TPU 配比（CPU 顆數 ÷ TPU 顆數；校準年） | 0.325 x | 預估該CPU出貨量將由2026年125萬顆大幅擴增至2027年約400萬顆；CPU與TPU配置比從過往一比四提升至一比二 | 只證實 CPU 125 萬顆；TPU 3.85 百萬顆不在文中。125÷385≈0.325 是模型自算，文中配比僅為 1:4、1:2。 |
| 部分相符 | guc | q3_gm | 3Q26 毛利率（公司：較 2Q26 再下滑） | 0.195 % | 第三季毛利率將較第二季進一步下滑，但全年仍可維持20%以上 | 來源只給方向（Q3低於Q2），沒有Q3數字；基準0.195低於Q2的21.5%，方向一致，但0.195與區間無從核對。 |
| 部分相符 | jentech | asp_bw | Blackwell（B200／B300）每顆均熱片＋加強環 | 35 US$/顆 | Blackwell Ultra 約30美元 | 未找到理財周刊8/5原文；MoneyDJ 7/29引美系外資寫約30美元。基準35偏高約17%，30僅在區間下緣。 |
| 部分相符 | jentech | th_26 | 2026 年散熱產品 ÷ 營收 | 0.75 % | 第二季散熱相關產品營收占比已提高至約77%（2025全年72.6%） | 來源是2Q26單季約77%，模型是2026全年0.75，期間不同，且差距約2.6%。來源沒有給全年數字。 |
| 部分相符 | kinsus | ai_abf | FY2026 AI（加速器、NVIDIA CPU、網通 ASIC）÷ ABF 營收 | 0.25 % | AI 相關應用將在 2026 年貢獻景碩約 10% 的總營收 | 來源只給 AI 占總營收約10%，沒有 AI÷ABF 營收，也沒有 ABF 占比，0.25 無法直接核對。 |
| 部分相符 | kinsus | cap_27 | ABF 產能指數 2027（FY2026 平均＝1） | 1.2 x | 每月ABF產能預計將從目前4,000萬片提升至2027年的5,000萬片 | 5,000/4,000=1.25，高於基準1.2（差約4%）但在區間內；來源未給2026全年平均，指數屬推算。 |
| 部分相符 | kinsus | sh_cpu_26 | 景碩占 NVIDIA CPU（Grace／Vera）載板 2026 | 0.35 % | 美系AI客戶CPU供貨率由10%攀升至50%，帶動ABF稼動率由86%提升至95% | 來源為50%，未點名NVIDIA、未給35%；模型基準0.35低於來源，50%僅是區間上限。 |
| 部分相符 | klac | fl26 | CY2026 半導體製程控制系統營收中代工邏輯占比 | 0.63 % | somewhere close to 65-ish%, so I'll call it low to mid-60s as I look at 2026 | 法說只給「low to mid-60s」、約65%，0.63在區間內但非點值；未找到MarketBeat頁，改以Investing.com逐字稿核對 |
| 部分相符 | klac | g_oth26 | SPTS＋PCB 兩事業 CY2026 年增 | 0.25 % | With these combined products expected to grow over 25% in calendar 2026 | 25%以上與0.25一致但屬下限；來源涵蓋特殊製程、PCB與元件檢測三塊。未找到MarketBeat版，改核對同場Investing.com逐字稿 |
| 部分相符 | kyec | ai_25 | 2025 年 AI 加速器測試營收 ÷ 營收 | 0.275 % | AI相關產品在2025年已貢獻京元電子25%至30%的營收 | 區間0.25～0.3與基準0.275（中點）相符；但來源是「AI相關產品」，口徑較廣，且為媒體轉述大摩報告，非原報告。 |
| 部分相符 | liteon | ai_25 | 2025 年 AI 相關營收 ÷ 營收 | 0.21 % | AI相關業務營收占比已突破20%（另有報導稱2025年約20%） | 指定uanalyze網址無法開啟，改核對股東會報導；只有「逾20%／約20%」，無21%，落在區間內 |
| 部分相符 | mediatek | fx26 | 3Q、4Q26 匯率（公司 3Q26 指引假設；FY2026 ASIC 換算用） | 32 NT$/US$ | forecasted exchange rate of 32 NT dollars to 1 US dollar（3Q26 指引假設） | 3Q26 基準 32 與講稿相符；講稿未給 4Q26 匯率，也沒有 30.5～33 區間。 |
| 部分相符 | micron | cb_1 | 一般 DRAM 位元成長 FY27 | 0.12 % | industry DRAM bit shipments grow in the low-20% range (CY2027); HBM grows faster than conventional DRAM | 來源只給含HBM的產業low-20%，未給一般DRAM數字；0.12方向合理但無法直接核對，且為CY非FY |
| 部分相符 | micron | cb_2 | 一般 DRAM 位元成長 FY28 | 0.15 % | industry DRAM bit shipments grow in the low-20% range (CY2028); HBM grows faster than conventional DRAM through CY2028 | 來源只給含HBM的產業low-20%，無一般DRAM單獨數字；0.15未見於來源，且為CY非FY |
| 部分相符 | micron | nb_1 | NAND 位元成長 FY27 | 0.18 % | For calendar 2027 and 2028, we expect industry NAND bit shipments to grow approximately in the mid-20s percentage range | 來源只給產業 CY27 中 20% 段（約 25%）。「低於產業」僅指 CY26，沒有美光自身數字。模型 0.18 是推算值，不是來源數字。 |
| 部分相符 | micron | nb_2 | NAND 位元成長 FY28 | 0.2 % | For calendar 2027 and 2028, we expect industry NAND bit shipments to grow approximately in the mid-20s percentage range | 來源給產業 CY28 中 20% 段（約 25%）。美光 CY27/28 供給成長未揭露，QA 明說不評論。模型 0.20 不是來源數字，標記 Verified 不成立。 |
| 部分相符 | micron | sh26 | HBM 營收市占 2026 | 0.19 % | Micron HBM 營收市占：Q1 2026 21%、Q2 2026 18%（Counterpoint） | Q1 21%、Q2 18%與來源欄一致；來源無2026全年值，0.19 為推算，介於兩季間且僅涵蓋上半年。 |
| 部分相符 | micron | sh27 | HBM 營收市占 2027 | 0.2 % | Micron accounting for the remaining 20%(UBS 2027 HBM bit market share,三星與 SK 海力士各約 40%) | 20% 一致,但來源是位元(出貨)市占,模型列為營收市占,口徑不同。引用的是 Investing.com 對 UBS 報告的轉述,未見 UBS 原文。 |
| 部分相符 | mrvl | g_cus27 | 客製營收年增 FY2027 | 0.25 % | Custom revenue remains on track to grow more than 20% year-over-year in fiscal 2027 | Q4/Q1 法說只給「>20%」下限。Q2 新聞稿稱下半年客製顯著加速但沒給數字。基準0.25與上限0.35找不到來源。 |
| 部分相符 | nanya_pcb | ai_abf | FY2026 AI 與資料中心 ÷ ABF 營收 | 0.65 % | AI and data center (800G switches) applications account for about 60% of ABF revenue | 來源約60%，模型0.65差約8%。60%在區間內；來源為7/27週報，非7/20報告，也非FY2026預估 |
| 部分相符 | nvidia | g28_gr | 公司 FY2028 營收年增 | 0.7 % | approximately 70%（FY2028 營收年增，受供給限制） | 官方新聞稿沒有FY28數字，逐字稿也未取得；只在二手報導看到「約70%」，與0.7一致。Quartz原文打不開。 |
| 部分相符 | nvidia | tx_f | 有效稅率（FY2028 起） | 0.17 % | For the full year fiscal 2027, NVIDIA expects GAAP and non-GAAP tax rates to be between 16.0% and 18.0% | 區間 16–18% 與中點 17% 吻合，但來源只涵蓋 FY2027，未提 FY2028 起，屬沿用外推。 |
| 部分相符 | oracle | lease_q | 未起租租約起租季數（FY27Q2 起平均） | 11 季 | expected to commence between the second quarter of fiscal 2027 and fiscal 2029(附註 6;未起租承諾 $288 billion,期限 15–19 年) | 10-Q 只寫起租區間 FY27Q2–FY2029,約 11 季。但沒有「平均季數」,11 季是區間總長,不是平均值。 |
| 部分相符 | oracle | p_spot | 現貨價格證據：H100 Silicon Data 指數 | 2.82 US$／GPU-hr | 官方頁面：2.89 USD/GPU-hour（截至 2026-10-26）；二手報導：2.82（2026-10-10） | 官網只顯示10/26的2.89，查不到10/06歷史值（需登入）。2.82只見於二手文章且日期為10/10，故無法確認。 |
| 部分相符 | oracle | pace_A | 情境 A：併網速度 | 3333 MW-IT／年 | expects to add nearly 1 GW of capacity in fiscal Q1 2027（Q4 FY26 電話會議，The Register 轉述）；Q1 實際交付 850 MW | 公司只預告 Q1 FY27 約 1 GW，未說每季都是 1 GW。×4 與 ÷1.2 是模型推算，來源沒有；Q1 實際 850 MW，Q4 約 400 MW。 |
| 部分相符 | samsung | q4_dp | Q4 DRAM ASP 季增 | 0.1 % | Conventional DRAM contract prices projected to grow 10–15% QoQ in 4Q26 | 基準0.10等於來源區間下限（中點12.5%）；來源指傳統DRAM合約價，非三星ASP；模型下限0.05低於來源區間。 |
| 部分相符 | samsung | sh25 | HBM 營收市占 2025 | 0.17 % | Samsung HBM 營收市占：Q2 2025 15%、Q3 2025 23%、Q4 2025 22% | 區間端點 15%、22% 與來源相符，但 Counterpoint 現已把 Q3 修訂為 23%。基準 0.17 來源未載，也無全年數字。 |
| 部分相符 | samsung | sh26 | HBM 營收市占 2026 | 0.29 % | Counterpoint:三星 HBM 營收市占 1Q26 21%、2Q26 33%(SK hynix 50%、Micron 18%);HBM4 出貨預期下半年明顯放量、三星市占逐步上升 | 季度數字相符,但來源沒有2026全年市占,0.29與區間0.25~0.33查無出處。Counterpoint頁面內文未顯示,21%由 sedaily 轉述。 |
| 部分相符 | samsung | sh27 | HBM 營收市占 2027 | 0.38 % | UBS 2027 HBM 位元出貨占比：三星 41%、SK 海力士 39%、美光 20% | 來源是位元占比 41%，模型是營收占比 0.38，口徑不同，差約 7%。來源 SK 為 39%，與說明欄的 42% 不同。 |
| 部分相符 | siliconomics | abfcap_ks_27 | ABF 產能指數 2027（FY2026 平均＝1） | 1.2 x | ABF 月產能由目前 4,000 萬片，提升至 2027 年 5,000 萬片（5,000/4,000＝1.25x） | 來源只有 4,000→5,000 萬（1.25x），沒有 FY2026 平均值。模型基準 1.2x 落在區間內，但與 1.25x 差約 4%，且此為法人推估。 |
| 部分相符 | siliconomics | asm_wyai_26 | 2026 年緯穎 AI 伺服器 ÷ 緯穎營收（CF09 單一值；緯穎、緯創模型同讀） | 0.48 % | AI 伺服器營收佔比將略高於 50%（2026-02 法說摘要）；2026-09-10 公司未設全年 ASIC 過半目標 | 公司說略高於50%，模型0.48略低；區間含0.50；未取得公司原稿，且曾有「略降」說法 |
| 部分相符 | siliconomics | asm_wyam | 緯穎 2026 年 ASIC 占 AI 伺服器營收 | 0.9 % | AI 伺服器中以 ASIC 伺服器為最大宗，約占 9 成（指 2025 年） | 約九成是 2025 年實績。2026 年只說 ASIC 仍為主，下半年 GPU 放量，未給 0.9。我只找到媒體轉述，未找到法說原稿。 |
| 部分相符 | siliconomics | cpo_A_26 | 情境 A：CPO 占 scale-out 光學金額 2026 | 0.01 % | scale-out CPO 交換器 2026 年 5,000 units(Citi) | 只找到 Citi 的 2026 台數,沒有金額占比 1%。Citi 原報告未公開,僅見媒體轉述。 |
| 部分相符 | siliconomics | cpo_A_27 | 情境 A：CPO 占 scale-out 光學金額 2027 | 0.05 % | scale-out demand ... 40,000 units(2027) | 2027 年 4 萬台與來源相符。5% 占比和 $2–3B 是模型自行換算,來源未載。報告為轉述。 |
| 部分相符 | siliconomics | cpo_A_28 | 情境 A：CPO 占 scale-out 光學金額 2028 | 0.1 % | 2028 scale-out 約 100,000 units(Citi) | 2028 年約 10 萬台與來源相符。10% 占比和 $6–8B 是模型換算,來源未載。報告為轉述。 |
| 部分相符 | siliconomics | cpo_B_26 | 情境 B：CPO 占 scale-out 光學金額 2026 | 0.01 % | CPO penetration in AI data center optical modules ~0.5% in 2026 (TrendForce) | TrendForce 為0.5%，模型1%約2倍；口徑是AI資料中心模組滲透率，非scale-out光學金額。LightCounting與MS原文未取得。 |
| 部分相符 | siliconomics | cpo_B_27 | 情境 B：CPO 占 scale-out 光學金額 2027 | 0.03 % | Volume production begins in 2026, with large-scale procurement expected in 2027–2028 (TrendForce) | TrendForce 只給2027–28放量的定性時程，沒有2027年數字；3%無法直接核對。報告全文付費。 |
| 部分相符 | siliconomics | cpo_B_28 | 情境 B：CPO 占 scale-out 光學金額 2028 | 0.05 % | 2027-2028 as the critical window for large-scale deployment (TrendForce) | 時程方向一致，但「2028年4–7%」與5%均無來源數字；LightCounting 2030約10%與MS原文均未取得。 |
| 部分相符 | siliconomics | cqA_10 | 一般 DRAM 合約價季增 2027Q3（A） | -0.02 % | 2027 年「持平或溫和修正」;2027 全年均價仍高於 2026 | 微降方向一致,但來源無 −2% 逐季數字;屬模型自行拆分 |
| 部分相符 | siliconomics | cqA_11 | 一般 DRAM 合約價季增 2027Q4（A） | -0.03 % | 2027 年「持平或溫和修正」;2027 全年均價仍高於 2026 | 溫和修正方向一致,但無 2027Q4 −3% 的逐季原文 |
| 部分相符 | siliconomics | cqA_12 | 一般 DRAM 合約價季增 2028Q1（A） | -0.03 % | 2028 DRAM 均價較 2027 約 −5%(整體 DRAM+NAND 約 −10%);8 月版為 DRAM −8% | 2028 為溫和修正,但來源只給年度降幅,無逐季數字 |
| 部分相符 | siliconomics | cqA_13 | 一般 DRAM 合約價季增 2028Q2（A） | -0.03 % | 2028 DRAM 均價較 2027 約 −5%(整體約 −10%) | 逐季 −3% 推得年均約 −10%,高於 DRAM −5%,口徑近整體 |
| 部分相符 | siliconomics | cqA_14 | 一般 DRAM 合約價季增 2028Q3（A） | -0.02 % | 2028 DRAM 均價較 2027 約 −5%(整體約 −10%) | 溫和修正方向一致;來源無 2028Q3 逐季數字 |
| 部分相符 | siliconomics | cqA_15 | 一般 DRAM 合約價季增 2028Q4（A） | -0.02 % | 2028 DRAM 均價較 2027 約 −5%(整體約 −10%) | 溫和修正方向一致;來源無 2028Q4 逐季數字,年降幅口徑不同 |
| 部分相符 | siliconomics | cqA_8 | 一般 DRAM 合約價季增 2027Q1（A） | 0.03 % | 2026Q4 DRAM +8% QoQ,之後 2027 各季「持平或溫和修正」;CSP 已同意 2027Q1 價格高於 2026Q4 | 方向相符(2027Q1 高於 2026Q4),但 BofA 公開摘要未給 +3% 的逐季數字 |
| 部分相符 | siliconomics | cqA_9 | 一般 DRAM 合約價季增 2027Q2（A） | 0.0 % | 2027 年「持平或溫和修正」(stable or mild correction),無逐季數字 | 0% 與「持平」方向一致,但找不到 2027Q2 的逐季原文數字 |
| 部分相符 | siliconomics | cqB_10 | 一般 DRAM 合約價季增 2027Q3（B） | -0.08 % | Citi：2H27 DRAM 價格 half-over-half 下降 3% | Citi 只給半年 −3%；模型 Q3/Q4 合計隱含約 −14%，方向同但幅度大得多。 |
| 部分相符 | siliconomics | cqB_11 | 一般 DRAM 合約價季增 2027Q4（B） | -0.15 % | Citi：2H27 DRAM 價格 half-over-half 下降 3% | Citi 無單季數字；模型 −15% 使 2H27 隱含約 −14%，遠比 Citi 陡。 |
| 部分相符 | siliconomics | cqB_12 | 一般 DRAM 合約價季增 2028Q1（B） | -0.25 % | Bernstein：DRAM 每 GB 價格 2028 年下降 52.9%（$2.23→$1.05） | 來源只有 2028 全年 −52.9%，無單季數字；模型路徑 2028 約 −47%，方向一致。 |
| 部分相符 | siliconomics | cqB_13 | 一般 DRAM 合約價季增 2028Q2（B） | -0.15 % | Bernstein：DRAM 每 GB 價格 2028 年下降 52.9%（$2.23→$1.05） | 來源僅年度數字，無 2028Q2 單季值；模型路徑 2028 約 −47%，非單季比對。 |
| 部分相符 | siliconomics | cqB_14 | 一般 DRAM 合約價季增 2028Q3（B） | -0.1 % | Bernstein：DRAM 每 GB 價格 2028 年下降 52.9%（$2.23→$1.05） | 只找到 2028 全年跌幅，無 Q3 單季值；模型年均約 −47% 與 −52.9% 量級一致。 |
| 部分相符 | siliconomics | cqB_15 | 一般 DRAM 合約價季增 2028Q4（B） | -0.05 % | Bernstein：DRAM 每 GB 價格 2028 年下降 52.9%（$2.23→$1.05） | 來源無 2028Q4 單季數字，僅全年 −52.9%；單季值為模型自設，無法逐季核對。 |
| 部分相符 | siliconomics | cqB_8 | 一般 DRAM 合約價季增 2027Q1（B） | 0.05 % | Citi：DRAM QoQ 2%（Micron 財年 2 月季，blended） | Citi 為 +2%（Micron 財季口徑），模型 +5%；放緩方向一致，數字與口徑不同。 |
| 部分相符 | siliconomics | cqB_9 | 一般 DRAM 合約價季增 2027Q2（B） | 0.02 % | Citi：5 月季 DRAM QoQ 0%，價格 2Q27 見頂 | 見頂時點一致，但 Citi 為 0%，模型 +2%；Citi 為財季 blended 口徑。 |
| 部分相符 | siliconomics | cq_4 | 一般 DRAM 合約價季增 2026Q1 | 0.95 % | conventional DRAM prices expected to surge 90–95% QoQ(1Q26) | 原稿只寫 1Q26 預測 90~95%(中點 92.5%)。93~98% 實際值僅見轉載報導。 |
| 部分相符 | siliconomics | eq_ind_cxmt | 長鑫 2025 資本支出（高值年度未核對） | 6.9 $B | CXMT capex estimated at about $6 billion to $7 billion in 2023 and 2024, with an increase of 5% expected in 2025 | 基準6.9與6–7B加5%（約6.3–7.35）相符；區間高值12在來源中找不到。文章日期為2025年10月，非年初。 |
| 部分相符 | siliconomics | eq_ind_kioxia | 鎧俠 FY2025 資本支出（FY 至 2026-03） | 283.7 ¥B | FY2026–28 年均約 ¥470 billion,較 FY2025 增加約 66% | TrendForce 未直接寫 283.7;470÷1.66≈283.1,差約0.2%,量級一致。FY26 ¥450B 也未見於該頁。 |
| 部分相符 | siliconomics | eq_rc_lam27 | Lam Research 2027 WFE（>200B，下限；額外對照） | 200 $B | preliminary 2027 WFE outlook is for more than $200 billion | FQ4-26稿未給2027數字（CFO稱不給）；>$200B出自10月FQ1-27法說，Yahoo文未能開啟 |
| 部分相符 | siliconomics | g_amd_27 | AMD Instinct 每顆 HBM 2027 | 360 GB | MI355X 288 GB HBM3E；MI455X 432 GB HBM4 | 區間 288～432 與上下限規格相符；基準 360 為 Assumed，來源未給。 |
| 部分相符 | siliconomics | g_bw_26 | NVIDIA Blackwell（B200／B300） 每顆 HBM 2026 | 260 GB | B200 192GB HBM3e（部分廠商寫 180GB）；B300 up to 288GB HBM3e | B300 288 相符；B200 公開為 180／192GB，非 186，組合權重來源未載，無法驗 260。 |
| 部分相符 | siliconomics | g_bw_27 | NVIDIA Blackwell（B200／B300） 每顆 HBM 2027 | 280 GB | B200 192GB HBM3e（部分廠商寫 180GB）；B300 up to 288GB HBM3e | B300 288 相符；B200 公開為 180／192GB，非 186，組合權重來源未載，無法驗 280。 |
| 部分相符 | siliconomics | g_hp_26 | NVIDIA Hopper（H200 等） 每顆 HBM 2026 | 141 GB | 141 GB HBM3e（H200 規格） | H200 規格 141GB 與模型值一致，但 TrendForce 發布未提 HBM 容量，且此頁為第三方。 |
| 部分相符 | siliconomics | g_hp_27 | NVIDIA Hopper（H200 等） 每顆 HBM 2027 | 141 GB | 141 GB HBM3e（H200 規格） | H200 規格 141GB 與模型值一致，但 TrendForce 無 2027 容量資料，僅為固定規格。 |
| 部分相符 | siliconomics | g_hp_28 | NVIDIA Hopper（H200 等） 每顆 HBM 2028 | 141 GB | 141 GB HBM3e（H200 規格） | H200 規格 141GB 與模型值一致，但 TrendForce 無 2028 容量資料，僅為固定規格。 |
| 部分相符 | siliconomics | g_rb_28 | NVIDIA Rubin 每顆 HBM 2028 | 288 GB | Up to 288 GB of HBM4 per GPU | 只能核對Rubin現行規格；2028年配置來源未載，Rubin Ultra規格可能不同。 |
| 部分相符 | siliconomics | g_ru_26 | NVIDIA Rubin Ultra 每顆 HBM 2026 | 1024 GB | 1024 GB (1 TB) of HBM4E memory per package | 只找到轉載 NVIDIA 路線圖的媒體，沒找到官方頁；Rubin Ultra 預定 2027 下半年出貨，2026 無此產品。 |
| 部分相符 | siliconomics | g_ru_27 | NVIDIA Rubin Ultra 每顆 HBM 2027 | 1024 GB | 1TB of HBM4e per four reticle-sized GPUs | 媒體轉載 GTC 2025 路線圖，每封裝 1TB 一致，但沒有官方頁；另有報導稱可能降為 192GB。 |
| 部分相符 | siliconomics | g_ru_28 | NVIDIA Rubin Ultra 每顆 HBM 2028 | 1024 GB | 1 TB of HBM4e per GPU (NVIDIA roadmap disclosures) | 數字是 1TB，但只有二手來源；報導稱 Kyber 機櫃延到 2028，另有 192GB 說法。 |
| 部分相符 | siliconomics | g_tpu_27 | Google TPU 每顆 HBM 2027 | 250 GB | TPU 8t HBM ~250 GB (est.) | 只有第三方網站標示約 250 GB 的估計值，Google 未公布 8t 的 HBM。 |
| 部分相符 | siliconomics | g_tpu_28 | Google TPU 每顆 HBM 2028 | 288 GB | TPU 8i HBM 288 GB | 第三方網站寫 8i 為 288 GB（2027 年底預覽）。沒有 2028 年的專屬數字，只有此網站為證。 |
| 部分相符 | siliconomics | hbm_p27 | HBM 單價 2027 變動（三情境相同；年約已議定） | 0.8 % | TrendForce：Blended ASP is now projected to jump 121% YoY；Micron：vast majority of calendar 2027 HBM bit supply 已簽、significant price increases YoY | 上限121%與Micron年約已簽屬實；UBS 79%僅見轉載、BofA 38%與JPM 54%查無來源，基準0.8無法確認 |
| 部分相符 | siliconomics | hbm_p28_A | HBM 單價 2028 變動（A） | 0.1 % | HBM平均售價2027年年增54%，2028年持續上漲25%（單位Gb價格達3.8美元） | JPM 2028 +25% 與區間上限 0.25 相符；基準 0.1 並非 JPM 數字。來源為媒體轉述 9 月研報，未見原報告。 |
| 部分相符 | siliconomics | hbm_sup27 | 2027 HBM 位元供給成長（三情境相同：產能已定） | 0.55 % | TrendForce:HBM bit shipments projected to grow by 50–60% YoY in 2027;J.P. Morgan(二手轉述):2026-28 bit-demand CAGR 63% for HBM | 基準0.55是TrendForce 50–60%的中點。上限0.63來自JPM的2026–28需求CAGR,不是2027供給,JPM原報告也未直接核對。 |
| 部分相符 | siliconomics | hbmsh_mu_26 | HBM 營收份額：美光 2026 | 0.19 % | Micron HBM revenue share: Q1 2026 21%, Q2 2026 18% | 來源只有Q1 21%、Q2 18%,無全年2026值;0.19介於兩者之間,上半年均值約19.5%,方向量級一致。 |
| 部分相符 | siliconomics | hbmsh_mu_27 | HBM 營收份額：美光 2027 | 0.2 % | Micron accounting for the remaining 20% (HBM bit market share, 2027) | UBS 2027 美光 20% 數字相符，但來源是位元出貨份額，模型鍵標為營收份額，口徑不同；7月更新版仍為20%（二手報導） |
| 部分相符 | siliconomics | hbmsh_skh_26 | HBM 營收份額：SK 海力士 2026 | 0.52 % | IDC Q1 2026：56.4%；Counterpoint Q1 58%、Q2 50% | IDC 56.4% 見 SK hynix F-1。Counterpoint 原文未開啟，Q1 58%、Q2 50% 僅見轉載報導。皆為單季值，無 2026 全年數字；0.52 介於兩者之間。 |
| 部分相符 | siliconomics | hbmsh_skh_27 | HBM 營收份額：SK 海力士 2027 | 0.42 % | UBS 2027 年 HBM 位元出貨份額：SK 海力士 39%（三星 41%、美光 20%） | 來源是位元份額 0.39，模型是營收份額 0.42，口徑不同、差約 7.7%。0.39 落在區間內。只核對到媒體轉述，未取得 UBS 原報告。 |
| 部分相符 | siliconomics | hbmsh_ss_25 | HBM 營收份額：三星 2025 | 0.17 % | Samsung HBM 營收份額：Q2 2025 15%、Q3 2025 22%（Counterpoint 現行表已修訂為 Q3 23%、Q4 22%） | 來源只有季度數字（15%、22%），沒有 2025 全年值；0.17 落在區間內但屬推算。Counterpoint 頁面僅由搜尋摘要讀到表格。 |
| 部分相符 | siliconomics | hbmsh_ss_26 | HBM 營收份額：三星 2026 | 0.29 % | Samsung HBM revenue share: 1Q26 21%, 2Q26 33%（Counterpoint） | 來源只有單季數字（21%、33%），沒有2026全年值。0.29落在兩者之間，方向一致；Counterpoint預期下半年HBM4放量、份額續增。 |
| 部分相符 | siliconomics | hbmsh_ss_27 | HBM 營收份額：三星 2027 | 0.38 % | Samsung 41%、SK hynix 39%、Micron 20%(UBS 2027 HBM 位元出貨份額) | UBS 為位元份額 41%,模型為營收份額 0.38,口徑與數字不同;0.41 在區間上緣。僅核對媒體報導,未取得 UBS 原報告。 |
| 部分相符 | siliconomics | lc_tf_27 | TrendForce：AI 晶片液冷滲透率 2027（顆數） | 0.6 % | 2027年有望逼近60% | 來源只寫「逼近60%」，沒有精確數字。模型值 0.6 方向與量級一致，但可能略高於實際預估。 |
| 部分相符 | siliconomics | nd_sh_essd25 | 2025 企業級 SSD 占位元需求 | 0.3 % | 企業級 SSD 占全球 NAND 出貨位元：2Q26 48%，一年前（2Q25）26% | 來源只給 2Q25 單季 26%，沒有 2025 全年值。模型 0.30 是全年口徑，與 26% 差約 15%。26% 剛好是區間下限 0.26。 |
| 部分相符 | siliconomics | nqA_10 | NAND 合約價季增 2027Q3（A） | -0.05 % | NAND is projected to enter a looser supply environment in 2H27, increasing downward pressure on prices (no figure given) | 方向相符，2H27 價格有下修壓力。來源沒給 -5% 這個數字。 |
| 部分相符 | siliconomics | nqA_11 | NAND 合約價季增 2027Q4（A） | -0.07 % | 2H27 供給轉寬鬆、價格下修壓力增加（無具體 %） | 方向相符，來源沒給 2027Q4 的 -7% 數字。 |
| 部分相符 | siliconomics | nqB_10 | NAND 合約價季增 2027Q3（B） | -0.1 % | NAND 2027 下半年較上半年 -5%（half-over-half） | Citi 只給半年 -5%，無單季數；方向一致，模型單季 -10% 較陡 |
| 部分相符 | siliconomics | nqB_11 | NAND 合約價季增 2027Q4（B） | -0.2 % | NAND 2027 下半年較上半年 -5%（half-over-half） | Citi 無 2027Q4 單季數；模型 -20% 比 Citi 半年 -5% 陡許多 |
| 部分相符 | siliconomics | nqB_12 | NAND 合約價季增 2028Q1（B） | -0.3 % | NAND $/GB 2028 下降 68.8%（$0.32→$0.10） | Bernstein 只給全年；模型 2028 四季複利約 -62.2%，量級一致但不同 |
| 部分相符 | siliconomics | nqB_13 | NAND 合約價季增 2028Q2（B） | -0.25 % | NAND $/GB 2028 下降 68.8%（$0.32→$0.10） | 無單季數；模型四季複利約 -62.2% 對 -68.8%，方向一致 |
| 部分相符 | siliconomics | nqB_14 | NAND 合約價季增 2028Q3（B） | -0.2 % | NAND $/GB 2028 下降 68.8%（$0.32→$0.10） | 無單季數；模型四季複利約 -62.2% 對 -68.8%，方向一致 |
| 部分相符 | siliconomics | nqB_15 | NAND 合約價季增 2028Q4（B） | -0.1 % | NAND $/GB 2028 下降 68.8%（$0.32→$0.10） | 無單季數；模型四季複利約 -62.2% 對 -68.8%，方向一致 |
| 部分相符 | siliconomics | nqB_8 | NAND 合約價季增 2027Q1（B） | 0.05 % | Citi NAND QoQ：Feb 季 0%（8月/11月/2月/5月季為 29%、7%、0%、-1%） | Citi 為 Micron 會計季、僅轉載報導；Feb 季 0% 對模型 5%，減速方向一致 |
| 部分相符 | siliconomics | nqB_9 | NAND 合約價季增 2027Q2（B） | 0.02 % | Citi NAND QoQ：May 季 -1%；2027 年 Q2 見頂 | 見頂時點一致，但 Citi 為 -1%、模型 +2%；來源為轉載報導 |
| 部分相符 | siliconomics | nq_4 | NAND 合約價季增 2026Q1 | 0.875 % | NextMSC：1Q26 NAND +85–90%（引述為最新預測修正）；TrendForce 2/2 新聞稿：+55–60% | 85–90% 僅見於轉述來源；TrendForce 2/2 稿寫 55–60%，未找到 85–90% 的原始發布。 |
| 部分相符 | siliconomics | nq_6 | NAND 合約價季增 2026Q3 | 0.15 % | NAND Flash contract prices expected to increase 10–15% QoQ in 3Q26 | 模型 0.15 為區間上限，中點為 0.125；來源欄所稱 MS 實際約 +20% 未核對。 |
| 部分相符 | siliconomics | nv_dc26 | NVIDIA 資料中心營收 2026（約 FY27） | 374 $B | Q1 資料中心 $75.2 billion；Q2 資料中心 $89.0 billion；Q3 總營收指引 $108.0 billion, plus or minus 2% | Q1、Q2 與 Q3 總營收指引屬實，加總 374.2 與 374 相符；但 Q3 資料中心 100、Q4 110 為假設，區間 340～400 無來源。 |
| 部分相符 | siliconomics | pa_amd | AMD Instinct 每顆 | 30 US$K | HSBC：MI355 ASP 由 USD25k 下修至 USD23k；MI455X 無官方價（部落格估 $15–30K+） | HSBC $23K 自 $25K 下修屬實，MI455X 確無官方價；但基準 30K 高於 MI355X 的 20–25K，僅落在 MI455X 估計區間內。 |
| 部分相符 | siliconomics | suopt_26 | scale-up 互連中光學占比 2026 | 0.0 % |  | 文章只有定性描述（現行 NVL72 銅互連），沒有 2026 光學占比數字；方向與 0% 一致。 |
| 部分相符 | siliconomics | suopt_27 | scale-up 互連中光學占比 2027 | 0.0 % |  | Kyber NVL144 未提光學，但 Rubin NVL576 第二層用光學（可插拔），嚴格 0% 略有出入，且無百分比。 |
| 部分相符 | siliconomics | suopt_28 | scale-up 互連中光學占比 2028 | 0.0 % |  | Feynman 預計 2028 年中至下半年出貨，可選銅或 CPO；方向支持低占比，但無 0–5% 數字。 |
| 部分相符 | siliconomics | suopt_29 | scale-up 互連中光學占比 2029 | 0.2 % |  | 文章稱 CPO 在設計中將更重要，只有方向，沒有 2029 年 10–30% 的依據；屬假設值。 |
| 部分相符 | siliconomics | u_bw_26 | NVIDIA Blackwell（B200／B300） 顆數 2026 | 5.4 百萬顆 | 70–80k NVL72 racks × 72 ≈ 5.0–5.8M packages（第三方推算） | 找不到 MS 原文 5.4M；只見第三方由 MS 機櫃數推算，且僅含 NVL72。BigGo 轉載頁無法開啟。 |
| 部分相符 | siliconomics | u_hp_26 | NVIDIA Hopper（H200 等） 顆數 2026 | 0.5 百萬顆 | Hopper series' share is forecast to drop from 10% to 7% | TrendForce 只給 2026 Hopper 占比 7%，沒給顆數，0.5 百萬顆無法直接核對。 |
| 部分相符 | siliconomics | u_rb_26 | NVIDIA Rubin 顆數 2026 | 1.7 百萬顆 | Rubin's share of NVIDIA's high-end GPU shipments is expected to decline from 29% to 22% | 22%占比相符；7.6M總量新聞稿未載（僅年增約26%）。9月報告稱Rubin占比再降。 |
| 部分相符 | siliconomics | u_rb_27 | NVIDIA Rubin 顆數 2027 | 7.0 百萬顆 | Rubin and Rubin Ultra chips to approach 7 million units by 2027 | 大摩原報告未公開，僅二手轉述；口徑為Rubin加Rubin Ultra，且是「接近」7M。 |
| 部分相符 | siliconomics | u_tpu_26 | Google TPU 顆數 2026 | 3.85 百萬顆 | MS 3.2 million（2026）；GF 4.504 million（2026） | MS 3.2、GF 4.5 兩端點屬實。3.85 是兩者中間值，來源沒有這個數字，屬模型取值。 |
| 部分相符 | siliconomics | u_tpu_27 | Google TPU 顆數 2027 | 5.5 百萬顆 | MS 5 million（2027）；GF 8.84 million（2027） | MS 5.0、GF 8.84 兩端點屬實。5.5 落在區間內且偏 MS，屬模型取值，來源沒有此數。 |
| 部分相符 | siliconomics | u_tpu_28 | Google TPU 顆數 2028 | 7.0 百萬顆 | MS 7 million（2028，2025/12 報告） | 來源欄標 2028 為 Assumed，並無對應來源。MS 舊報告的 7 百萬顆恰與 7.0 一致。 |
| 部分相符 | skhynix | q3_np | Q3 NAND ASP 季增 | 0.125 % | NAND Flash contract prices are expected to increase 10–15% QoQ (3Q26) | TrendForce 10–15%中點即0.125，相符；MS實際約+20%未找到原始發布，僅核到一半 |
| 部分相符 | skhynix | sh25 | HBM 營收市占 2025 | 0.58 % | 57% of revenue as of Q3(出貨量:62% as of Q2 2025) | 來源為Q3單季57%,模型0.58是2025全年口徑;Counterpoint現表已改為Q3 56%、Q4 57%,且此頁為SK海力士自家轉述。 |
| 部分相符 | smci | ais_25 | FY2025 AI GPU 平台 ÷ 營收 | 0.7 % | GPU AI platforms ... represented over 70% of Q4 revenues（FY25 第四季法說，CFO） | 法說只說 FY25 Q4 逾70%，未給全年占比；口徑為季非年，0.70 僅是下限。逐字稿取自第三方轉載，非公司原檔。 |
| 部分相符 | tsmc | cap_end25 | CoWoS 月產能 2025 年底 | 77.5 千片／月 | 台积电CoWoS月产能于年底接近翻倍，达7.5万至8万片 | TrendForce 2025年底7.5–8萬片，中點77.5吻合；但只找到2024-11《科創板日報》轉述，沒找到工商時報2025-12原文。 |
| 部分相符 | tsmc | cap_end26 | CoWoS 月產能 2026 年底 | 130 千片／月 | TSMC’s monthly CoWoS capacity could reach a record 120,000 to 140,000 wafers in 2026 | TrendForce 轉述經濟日報 12–14 萬，130 為區間中點；MS 原報告未找到，僅見轉載站(12–13 萬)，不採用 |
| 部分相符 | unimicron | ai_pcb | FY2026 HDI＋PCB 的 AI 占比 | 0.65 % | 上半年AI相關應用產品營收占比已逾60％，預期下半年可望提升至65～70％（PCB 下半年 65～70%） | 來源是PCB上半年逾60%、下半年65–70%，模型是全年0.65，口徑不同；未取得工商時報本站原文，改用同集團時報資訊同日報導 |
| 部分相符 | unimicron | mx_abf | FY2026 ABF 載板 ÷ 營收 | 0.53 % | ABF載板營收比重從上季的49%拉高至52% | 來源只有2Q26的52%（1Q 49%），無FY2026全年值；0.53在區間內但口徑不同。工商時報法說原文未找到，改用自由財經。 |
| 部分相符 | unimicron | mx_bt | FY2026 BT 載板 ÷ 營收 | 0.09 % | 占9％的BT載板季減3％、年減4％（2Q26 產品營收結構） | 2Q26 BT 占 9% 與 0.09 一致，但這是單季數字，模型為 FY2026；工商時報 7/29 文章無此數，數字出自中時新聞網；自由時報寫 10%。 |
| 部分相符 | wistron | q4g | 4Q26 合併營收季增 | 0.05 % | FY2026 營收 Avg 4.1T（High 4.2T／Low 4.0T，精確值 4.07T），No. Analysts 7；前三季累計 2兆8,908億、3Q26 約 1.15兆 | 共識與家數相符。以4.1T推得4Q約+5%，用4.07T則約+2.7%。來源未直接列4Q季增。 |
| 部分相符 | wistron | wai_26 | 2026 年緯穎 AI 伺服器 ÷ 緯穎營收 | 0.48 % | AI 伺服器營收佔比將略高於 50%（2026 法說摘要）；凱基 46% 未找到 | 法說摘要為二手整理，原始法說 PDF 與 siliconomics、凱基 46% 皆未取得。模型 0.48 介於 46% 與略高於 50% 之間。 |
| 部分相符 | wistron | wgm | 緯穎毛利率（FY2027 起） | 0.085 % | FY2025 毛利率 8.3%（富果）；1H26 毛利率 8.41%、2Q26 毛利率 9.26%（優分析） | 來源只列歷史毛利率，與 0.085 同量級（對 1H26 8.41% 差約1%）。FY2027 起的數字來源未載，屬模型假設。 |
| 部分相符 | wiwynn | ais_26 | 2026 年 AI 伺服器 ÷ 營收 | 0.48 % | AI伺服器佔比今年預期將超過一半（MoneyDJ 2026-02-26 法說報導）；另一則財訊報導稱今年佔比會略降於去年的 50% | 公司說「超過50%」，沒有給確切數字。模型 48% 略低於此。優分析原文找不到，3Q26 各半一事由 9/10 MoneyDJ 證實。 |
| 部分相符 | wiwynn | asic_mix | 2026 年 ASIC 占緯穎 AI 伺服器營收 | 0.9 % | 去年 AI 伺服器中 ASIC 伺服器為最大宗，約占 9 成 | 約九成是2025年數字，2026只說ASIC仍為主；下半年GPU新平台出貨，2026占比未揭露。優分析原文未找到，用法說報導核對。 |
| 部分相符 | wiwynn | q_gm | 3Q、4Q26 毛利率 | 0.083 % | 第二季毛利率9.26%；預期下半年ASIC比重提升，即便產品組合改變，毛利率變化不大，不過還需觀察NRE貢獻程度 | 來源只有定性說法「下半年毛利率變化不大」，沒有3Q、4Q數字。1H26 8.41%、2Q26 9.26%相符。0.083屬模型推估，無法直接驗證。 |
