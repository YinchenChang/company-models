# 來源複核結果

- 應核對 2192 列，已核對 1604 列；累計花費約 86.07 美元（上限 84.0）；搜尋 1546 次、擷取 1764 次。

- 不符：12
- 部分相符：232
- 來源打不開：238
- 找不到來源：148
- 相符：974

## 不符與部分相符

| 判定 | 模型 | 鍵 | 項目 | 基準 | 來源上的數字 | 說明 |
|---|---|---|---|---|---|---|
| 不符 | advantest | g_eps | 指引 EPS FY2026（基本） | 911.64 ¥ | 通期 基本的1株当たり当期利益 1,027.08円(2027年3月期、2026/10/28 第2四半期決算短信) | 模型值911.64是7月第1季短信的舊指引,10/28上修後最新為1,027.08,差約11%。 |
| 不符 | advantest | mk_soc26 | SoC 測試機市場 CY2026（公司上修後） | 11 US$B | SoC テスタ市場 CY2026 予測レンジ 8.7–9.5 billion USD(1月時點は 8.5–9.5) | 已開啟公司原文:4/27 簡報寫 8.7–9.5,非 10.5–11.5。10.5–11.5 只見於 X 上的二手貼文,似為 7 月之後的上修,未經公司原文核實。 |
| 不符 | amkr | cons_tp | 市場共識目標價（賣方平均） | 76.4 US$ | 平均 $74.5;中位數 $73.5;區間 $46–$92(11 位分析師,S&P Global,更新於 Oct 23, 2026) | 頁面平均為74.5(非76.4)、低標46(非65);高標92與中位數73.5相符。76.4/65似為舊版數字。 |
| 不符 | auras | q4g | 4Q26 營收季增 | 0.14 % | 預估雙鴻第3季、第4季營收分別季增17%、16% | fetch 工具報錯未能直開，內文取自搜尋結果的同一 udn 頁面。花旗 4Q 季增 16%，模型 14%，差 12.5%，超出容許。 |
| 不符 | ibiden | pe_fwd_sa | 前瞻 P/E（stockanalysis） | 53.97 x | Forward PE 46.82(價格 ¥8,512,2026-10-23 收盤;頁面更新 2026-10-24) | 頁面前瞻 P/E 為 46.82x,模型 53.97x 高約 15%。此值隨股價每日變動,差異可能來自取值日不同。 |
| 不符 | samsung | g_tpu_27 | Google TPU 每顆 HBM 2027 | 250 GB | TPU 8t HBM Capacity 216 GB;TPU 8i 288 GB | Google 官方 8t 為 216 GB、8i 為 288 GB,無 250 GB;250 僅見二手網站估計。 |
| 不符 | samsung | nq_4 | NAND 合約價季增 2026Q1 | 0.875 % | NAND Flash to +55–60% QoQ(1/5 原預測 33–38%,2/2 上修) | TrendForce 1Q26 為 55–60%(中點 0.575),找不到 85–90%。模型 0.875 不符。 |
| 不符 | siliconomics | ts_mk_soc26 | SoC 測試機市場 CY2026（公司上修後） | 11 US$B | CY2026 SoC tester market size forecast range 8.7–9.5 billion USD(1月預測 8.5–9.5,下限上修 0.2) | 公司原文p.19寫SoC測試機CY2026為87～95億美元,非10.5～11.5;基準11也不符(中點約9.1)。 |
| 不符 | skhynix | dps_2 | 每股股利 2028 | 5249 KRW | Dividend per Share 2028: 3,458 (KRW)；同表 2027: 5,158 | 2028 為 3,458，非 5,249；2027 的 5,158 近模型值，疑年份錯位。表格取自搜尋摘要，擷取日不明。 |
| 不符 | skhynix | nq_4 | NAND 合約價季增 2026Q1 | 0.875 % | NAND Flash +55–60% QoQ（由33–38%上調） | TrendForce 1Q26 NAND為55–60%，非85–90%（後者是DRAM的90–95%量級）；模型0.875不符。 |
| 不符 | unimicron | cons_tp | 市場共識目標價（賣方平均） | 1317.37 NT$ | Price Target 1,347.89 (+24.23%);19 位分析師(2026-10-22 頁面) | 同日同為19位分析師,頁面為1,347.89,模型1,317.37低約2.3%,超出容許。10/14統計頁為1,277.37,也不吻合。 |
| 不符 | victory_giant | ai_25 | 2025 年 AI 相關營收（AI 算力、資料中心、HPC）÷ 營收 | 0.45 % | AI及高性能计算对总收入的贡献率…2025年的43.2%（营收83.41亿元／192.92亿元） | 來源寫43.2%，模型0.45，相對差約4.2%，超過2%。此為摩根大通報告的媒體摘要，原報告未公開。 |
| 部分相符 | accton | nap25s | FY2025 網路應用產品（AI 加速卡等）÷ 營收 | 0.63 % | 網路應用設備 營收佔比 63%，年增 361%（2025 前三季） | 63% 與模型值相符，但來源是前三季佔比，模型標的是 FY2025 全年。fetch 工具拒絕開啟該網址，內容取自搜尋結果摘錄。 |
| 部分相符 | accton | pe_ttm | 近 12 月 P/E（鉅亨，10-22） | 27.88 x | 本益比 28.51（10/28 收盤，近四季EPS 63.65，股價1815） | 鉅亨只顯示最新值28.51，查不到10-22當日值。口徑與量級一致，27.88約對應股價1775，無法確認日期。 |
| 部分相符 | accton | sw25s | FY2025 網路交換器 ÷ 營收 | 0.34 % | 網路交換器 營收佔比 34%，年增 46%（2025 前三季） | 34% 與 0.34 一致，但來源是前三季佔比，非 FY2025 全年；口徑不同。 |
| 部分相符 | ajinomoto | px | 股價（2026-10-23 收盤） | 5064 ¥ | 前日終値: 5,064 (10/23)(株探 10/26 頁面) | 松井頁面僅顯示10/28最新值，時系列表空白，查不到10/23；株探佐證收盤5,064，前日5,046未驗證 |
| 部分相符 | alchip | fpe | 前瞻 P/E（stockanalysis） | 17.15 x | forward PE ratio is 16.74(頁面快照 2026-09-11;10/1 快照為 17.26) | 同量級,但為隨股價變動的比率。模型 17.15 與 16.74 差約 2.4%,略超容許,且快照日不同。 |
| 部分相符 | alchip | pe_ttm | 近 12 月 P/E（stockanalysis） | 50.74 x | The trailing PE ratio is 52.99(頁面快照 2026-09-11 收盤 3,650;10/1 快照為 54.66) | 同量級,但頁面是股價日更比率,快照日不同。模型 50.74 與 52.99 差約 4.2%,無法驗證 10/28 當日值。 |
| 部分相符 | alchip | q3_gm | 3Q26 毛利率（公司：低至中 20%） | 0.21 % | for the third quarter and the fourth quarter, the blended gross margin will most likely be about 20 something percent | 逐字稿只有「20幾%、低20%」的定性區間，0.21落在區間內，但來源沒有點值 |
| 部分相符 | alchip | q3_opx | 3Q26 營業費用（公司：與 2Q26 相近） | 1.05 NT$B | for the third quarter and the fourth quarter, the operating expense will be pretty similar with the second quarter | 「與2Q26相近」屬實，但逐字稿以美元報價且未給營業費用金額，NT$1.05B無法直接核對 |
| 部分相符 | amazon | dps | 每股股利 | 0 US$ | We have never declared or paid cash dividends on our common stock. | S02未定義;FY2017等舊10-K明寫從未派息;FY2025 10-K股利段被截斷未讀到 |
| 部分相符 | amazon | kd_new | 新增借款利率（2026-07 新發 2066 票據有效利率；只用於利息費用） | 0.0633 % | 6.250% notes due 2066；定價殖利率 6.297%（99.316% 發行價） | 只找到票面 6.25% 與殖利率 6.297%，未見 6.33% 有效利率原文；量級一致。 |
| 部分相符 | amazon | p_list | AWS p5.48xlarge 隨需（每 GPU 小時） | 6.88 US$/GPU-hr | p5.48xlarge on-demand $55.0400/hr(us-east-1),÷8=$6.88/GPU-hr | 多個價格彙整站皆為$55.04/hr,換算6.88吻合。AWS官方公告只列降幅44%,未見美元價,官方價目頁未能開啟。 |
| 部分相符 | amazon | p_spot | H100 現貨價（Silicon Data 指數） | 2.82 US$/GPU-hr | As of Oct 26, 2026: 2.89 USD / GPU-hour(Neo-Cloud SDH100RT) | 官方頁現值2.89,差約2.4%。2.82為10/10讀數(第三方轉述)。頁內FAQ另寫2.53,前後不一致。 |
| 部分相符 | amkr | az_inc | 政府補助與投資抵減 ÷ 投資（降低折舊基礎） | 0.35 % | CHIPS 補助 $400 million + 35% investment tax credit,合計 $2.8 billion;Arizona 投資 $7 billion | 逐字稿確有 35% 抵減與 $2.8B／$7B,但合計比率為 2.8÷7=0.40,基準 0.35 只等於抵減率,0.40 落在區間上限。 |
| 部分相符 | amkr | fpe | 前瞻 P/E（stockanalysis，10-23） | 14.02 x | forward PE ratio is 13.30(頁面更新日 Oct 26, 2026,股價 37.34) | 頁面只有 10-26 盤中值 13.30,查不到 10-23 的數字。前收盤 39.30 換算約 14.0,與 14.02 相近,但差距超過 2%。 |
| 部分相符 | amkr | g3_cm | 3Q26 通訊季增（公司說法） | -0.08 % | For the third quarter, we expect communications revenue to decline in the high single digits sequentially | 財務長原話只說高個位數季減，沒給確切數字；模型 -8%（區間 -10%～-6%）是量化推估，方向與量級一致 |
| 部分相符 | amkr | pe_ttm | 近 12 月 P/E（stockanalysis，10-23） | 17.64 x | The trailing PE ratio is 16.76(頁面更新日 Oct 26, 2026,股價 37.34) | 頁面只有 10-26 盤中值 16.76,查不到 10-23 的數字。前收盤 39.30 除以 EPS 2.23 約 17.6,與 17.64 相近,但差距超過 2%。 |
| 部分相符 | auras | px | 股價（2026-10-22 收盤） | 1360 NT$ | 成交價 1360、漲跌 -15、漲跌幅 -1.09%(10/22) | Yahoo 頁面未載有10/22收盤價;數字只在Goodinfo找到且一致,未能在Yahoo本身核實 |
| 部分相符 | avc | q_gm | 3Q、4Q26 毛利率 | 0.33 % | 2Q26 合併毛利率 32.57%，季增 2.8 個百分點 | 2Q26 32.57% 已證實。找不到工商時報 8/6「3Q 朝 33%」原文，經濟日報同日報導只寫下半年毛利率提升，區間也無來源。 |
| 部分相符 | avc | srv_25 | 2025 年伺服器與網通 ÷ 營收 | 0.52 % | 伺服器(52%)與網通(7%)；圖3 2025年伺服器營收比重 52 | 52%僅為伺服器，網通另列7%（合計約59%）；核對的是8/5報告，未找到8/12版，1H25 48%未見 |
| 部分相符 | besi | cons_eps26 | 共識 EPS FY2026 | 4.334 € | EPS 2026: 4,23 EUR | 量級一致，來源4.23比4.334低約2.4%。表格取自搜尋摘要，快照日期未能確認為10/9。 |
| 部分相符 | besi | cons_eps28 | 共識 EPS FY2028 | 8.483 € | EPS 2028: 8,80 EUR | 方向一致，8.80比8.483高約3.7%，超出2%。快照日期未能確認為10/9。 |
| 部分相符 | besi | cons_rev26 | 共識營收 FY2026 | 1002 €M | Net sales 2026：1B（EUR，頁面四捨五入顯示）；EV 11.61B ÷ EV/Sales 2026 11.6x ≈ 1.00B | 量級一致，但頁面只顯示約1B，無法確認1002；頁面為10/26快照，非10/09；財務頁數字被截斷。 |
| 部分相符 | besi | cons_rev27 | 共識營收 FY2027 | 1345 €M | Net sales 2027*: 1.34B EUR（四捨五入顯示） | 1.34B 與 1345 差約 0.4%。但這是 10/27 左右（Q3 後、股價 163.05）的快照，查不到 10/9 財報前版本。 |
| 部分相符 | besi | tax | 有效稅率 | 0.135 % | Effective tax rate 13.1%（Q2-26） | Q2-26 的 13.1% 已在新聞稿證實；Q3-26 的 13.3% 因新聞稿 PDF 無權限讀取而未核實。模型值 0.135 略高於兩者。 |
| 部分相符 | cohr | cons_tp | 市場共識目標價（賣方平均） | 394.79 US$ | The average price target for Coherent is $396.00(Analyst Count 26;Last updated: Oct 27, 2026) | 差0.3%在容許內,但來源是26位、396.00、10/27更新,非25位、394.79、10/24。394.79見於financhill。 |
| 部分相符 | cohr | px | 股價（2026-10-23 收盤） | 209.17 US$ | $221.66 +12.37 (+5.91%)(收盤,頁面更新於 10/26/2026) | 頁面只顯示10/26收盤221.66;扣漲幅得10/23收盤約209.29,與209.17差0.06%。頁面未直接列出209.17及+0.65%。 |
| 部分相符 | dell | fpe | 前瞻 P/E（stockanalysis，10-23） | 16.22 x | Forward PE 17.23(頁面更新 Oct 27, 2026) | 頁面僅顯示10-27即時值17.23,無10-23歷史值;差約6%,與pe_ttm差幅相同,應為股價時點差異。 |
| 部分相符 | dell | g_q3rev | 3Q FY27 營收指引（參考） | 49 US$B | Q3FY27 Revenue $49.0 (billions), 81% Y/Y | 基準49十億美元與8-K一致；來源只給單一點估計，未提供48.5～49.5區間，區間無從核對。 |
| 部分相符 | dell | pe_ttm | 近 12 月 P/E（GAAP，stockanalysis） | 26.63 x | PE Ratio 28.30(頁面更新 Oct 27, 2026;EPS $17.14) | 量級一致但數字不同:頁面為10-27即時值28.30,差約6%;TTM EPS同為17.14,差異應屬股價時點。 |
| 部分相符 | delta | ai_26 | 2026 年 AI 相關營收 ÷ 營收（讀為不含液冷） | 0.26 % | AI-Related Revenue Share: Over 25% this year (not disclosed last year), target 28%. | 法說原稿PDF打不開；BigGo只寫>25%、目標28%，無26%，0.26落在區間內，未見是否含液冷。 |
| 部分相符 | delta | lc_26 | 2026 年液冷營收 ÷ 營收 | 0.125 % | Liquid-cooling products accounted for approximately 10% of total revenue last year and will definitely exceed 12% this year. | CFO只說2026年液冷占比「超過12%」（下限）。基準12.5%與區間上限13%是模型自行假設，來源未提。 |
| 部分相符 | disco | mix_eq | 精密加工裝置占營收 FY2025 | 0.62 % | 精密加工装置合計 62%（Full Year 構成比，出荷額ベース） | p.29 通期62%與0.62一致，子項31/28/3、21、16亦同；但來源為出荷額（出貨額）基礎，非營收，口徑不同。 |
| 部分相符 | disco | mix_tl | 精密加工ツール（消耗品：刀片、砂輪）占營收 FY2025 | 0.21 % | 精密加工ツール 構成比 21%（Full Year，出荷額ベース） | 第29頁全年構成比21%與模型0.21一致，但來源為出荷額（出貨額）基礎，非營收占比，口徑不同。 |
| 部分相符 | eoptolink | px | 股價（2026-10-23 收盤，A 股） | 261 RMB | 前收 258.01 元（10-22 收盘）；中财网页面目前只显示 10-29 行情：228.40，昨收 256.79 | 指定网址无法直接打开。同站行情页已无10-23收盘价。10-22收盘258.01与说明一致，+1.16%推得261.00，但261.00本身未证实。 |
| 部分相符 | fabrinet | a26_op_in | FY2026 non-GAAP 營業利益 | 0.500063 % | Operating profit (Non-GAAP) $500,063(千美元),占營收 10.8% | 10-K 只讀到前段,未見此數;Q4 新聞稿有 500,063 千美元,模型單位標 % 不當 |
| 部分相符 | fabrinet | hpc_q1 | FY2026 Q1 HPC（高效能運算）營收 | 0.0154 US$B | high-performance computing revenue of $15 million | 找不到 Pulse 2.0 原文；法說只揭露約1,500萬美元（取整），0.0154需精確值，僅見部落格 |
| 部分相符 | foxconn | ais_26 | 2026 年 AI 伺服器 ÷ 雲端網路產品 | 0.6 % | AI 伺服器營收已占鴻海第一季雲端網路產品營收超過 50% | 來源只給 1Q26「超過50%」，未給2026全年0.6；方向一致，但口徑為單季下限值。 |
| 部分相符 | foxconn | cl_26 | 2026 年雲端網路產品 ÷ 營收 | 0.53 % | 2Q26 雲端網路產品佔比 51%（首次突破五成）；3Q26 預期季對季、年對年強勁成長 | 來源只有 2Q26 51% 與 3Q 成長展望，沒有 2026 全年佔比。0.53 與走勢相符，但屬推算值，無法直接核對。 |
| 部分相符 | foxconn | pc_26 | 2026 年電腦終端產品 ÷ 營收 | 0.12 % | 電腦終端產品：營收佔比 15%（2Q26）；3Q26 預期較 2Q 略為衰退 | 來源只有 2Q26 單季佔比 15%，沒有 2026 全年佔比，也沒找到 1Q26 的 14%。模型 0.12 低於單季值，屬推估。 |
| 部分相符 | foxconn | pe_ttm | 近 12 月 P/E（Yahoo，10-16） | 16.08 x | 16.15(本益比,Yahoo股市 2317 頁面,2026/10/27 更新) | 找不到10/16當日本益比。10/27為16.15,與16.08差約0.4%,量級一致但日期不同。 |
| 部分相符 | foxconn | q4g | 4Q26 營收季增 | 0.15 % | 可望帶動整體營運表現符合目前市場預期 | 公告只有定性展望（AI成長、ICT旺季、符合市場預期），無4Q季增數字。0.15無法驗證，只有成長方向一致。 |
| 部分相符 | foxconn | q_gm | 3Q、4Q26 毛利率 | 0.06 % | 1Q26 6.18%;2Q26 6.12%;4Q25 5.88% | 三個歷史毛利率與來源一致，基準0.06與區間0.057~0.063是預測，來源無3Q/4Q26數字。2Q26僅見媒體轉述財報。 |
| 部分相符 | gce | net26 | 2026 年網通營收 ÷ 營收 | 0.15 % | 網通約占15%比重 | 文章日期 2026/8/28，數字 15% 與 0.15 一致；但原文是「現階段」占比，不是 2026 全年口徑。 |
| 部分相符 | gce | q1_rev | 1Q26 營收（實績） | 19.313 NT$B | 19,312,774 仟元(約 193.13 億元) | 標示來源頁只顯示年度/TTM,無1Q26季資料;財報狗與鉅亨網報導皆為193.13億,與19.313相符。 |
| 部分相符 | gce | srv26 | 2026 年伺服器營收 ÷ 營收 | 0.78 % | 伺服器營收占比達78% | 文章日期 2026/8/28，數字 78% 與 0.78 一致；但原文是「現階段」占比，不是 2026 全年口徑。 |
| 部分相符 | guc | att26 | CY2026 CPU 對 TPU 配比（CPU 顆數 ÷ TPU 顆數；校準年） | 0.325 x | 預估該CPU出貨量將由2026年125萬顆大幅擴增至2027年約400萬顆；CPU與TPU配置比從過往一比四提升至一比二 | 只證實 CPU 125 萬顆；TPU 3.85 百萬顆不在文中。125÷385≈0.325 是模型自算，文中配比僅為 1:4、1:2。 |
| 部分相符 | guc | fpe | 前瞻 P/E（stockanalysis，2026-10-08；股價 8,360） | 102.85 x | Forward PE 99.75(股價 7,900,Oct 2);搜尋摘要另見 99.59(股價 8,095,Oct 13) | 找不到 10/8、股價 8,360 的快照。模型值與頁面隱含前瞻 EPS 約 81.28 相符,但頁面現值 99.75 與 99.59 皆差約 3%。 |
| 部分相符 | guc | q3_gm | 3Q26 毛利率（公司：較 2Q26 再下滑） | 0.195 % | 第三季毛利率將較第二季進一步下滑，但全年仍可維持20%以上 | 來源只給方向（Q3低於Q2），沒有Q3數字；基準0.195低於Q2的21.5%，方向一致，但0.195與區間無從核對。 |
| 部分相符 | guc | q3_opx | 3Q26 營業費用（公司：與 2Q26 約持平） | 1.35 NT$B | 第三季營業費用與前季大致持平 | 來源只寫與2Q26大致持平，沒有金額。1.35 NT$B 無法直接核對。 |
| 部分相符 | hanmi | hsk_25 | 韓美占 SK 海力士 TC 鍵合機採購 2025 | 0.5 % | SK hynix is using about fifty sets of TC bonders for HBM4 production, with roughly half supplied by ASMPT; Hanwha Semitech's bonders are not currently in use | 來源未直接給韓美占比；僅 ASMPT 約半數、韓華未使用，可推得韓美約 0.5，但口徑是 HBM4 在用台數，非 2025 採購占比。 |
| 部分相符 | hanmi | ms_mu_28 | 美光 HBM 營收市占 2028 | 0.2 % | 2028 份額合計 0.40 + 0.40 + 0.20 = 1.00(美光 0.20);高值合計 0.45 + 0.45 + 0.23 | xlsx 讀不到,只在 PR 核對留言見美光 2028 基準 0.20;該列原無外部來源,Interested-party 標記存疑 |
| 部分相符 | innolight | cons_tp | 市場共識目標價（賣方平均） | 1321 RMB | 14 analysts, average ¥1,321; median ¥1,400; low ¥430.00, high ¥1,800 | 預測頁數字全吻合,但頁面時點為9/2;10/19概覽頁已是1,392(17位),差5%,非10/25資料 |
| 部分相符 | jentech | asp_bw | Blackwell（B200／B300）每顆均熱片＋加強環 | 35 US$/顆 | Blackwell Ultra 約30美元 | 未找到理財周刊8/5原文；MoneyDJ 7/29引美系外資寫約30美元。基準35偏高約17%，30僅在區間下緣。 |
| 部分相符 | jentech | th_26 | 2026 年散熱產品 ÷ 營收 | 0.75 % | 第二季散熱相關產品營收占比已提高至約77%（2025全年72.6%） | 來源是2Q26單季約77%，模型是2026全年0.75，期間不同，且差距約2.6%。來源沒有給全年數字。 |
| 部分相符 | kinsus | ai_abf | FY2026 AI（加速器、NVIDIA CPU、網通 ASIC）÷ ABF 營收 | 0.25 % | AI 相關應用將在 2026 年貢獻景碩約 10% 的總營收 | 來源只給 AI 占總營收約10%，沒有 AI÷ABF 營收，也沒有 ABF 占比，0.25 無法直接核對。 |
| 部分相符 | kinsus | cap_27 | ABF 產能指數 2027（FY2026 平均＝1） | 1.2 x | 每月ABF產能預計將從目前4,000萬片提升至2027年的5,000萬片 | 5,000/4,000=1.25，高於基準1.2（差約4%）但在區間內；來源未給2026全年平均，指數屬推算。 |
| 部分相符 | kinsus | sh_cpu_26 | 景碩占 NVIDIA CPU（Grace／Vera）載板 2026 | 0.35 % | 美系AI客戶CPU供貨率由10%攀升至50%，帶動ABF稼動率由86%提升至95% | 來源為50%，未點名NVIDIA、未給35%；模型基準0.35低於來源，50%僅是區間上限。 |
| 部分相符 | kioxia | cap_2026 | 資本支出 FY2026 | 450 ¥B | FY26 Capex: 450 B¥ (+60% YoY) | TrendForce 該文未列 FY2026 ¥450B，只寫三年均值約¥470B。450 見於 Kioxia 投資人日簡報（來自搜尋摘要，未開全文）。來源標示需改。 |
| 部分相符 | kioxia | cap_2027 | 資本支出 FY2027 | 470 ¥B | average annual capital expenditure of around ¥470 billion from FY2026 to FY2028 | 來源只給 FY26–28 年均約¥470B，沒有 FY2027 單年數字。模型把年均當單年值，口徑不同。 |
| 部分相符 | kioxia | cap_2028 | 資本支出 FY2028 | 470 ¥B | average annual capital expenditure of around ¥470 billion from FY2026 to FY2028 | 來源只給三年年均約¥470B，沒有 FY2028 單年數字。450＋470＋470 的均值約463，與年均470略有差距。 |
| 部分相符 | kioxia | cons_tp | 市場共識目標價（賣方平均；低／高＝最低／最高） | 36929.23 JPY | Average target price 36,233.40 JPY(16位,MarketScreener 個股頁,收盤日10-28) | 共識頁(10-16)讀到但目標價數字被截斷;同站10-28為36,233,與36,929差約1.9%,日期不同。高低值16,333/66,667未能核對。 |
| 部分相符 | kioxia | q1_gp | FY2026 第一季毛利 | 1380.1 ¥B | TTM(Jun '26) 毛利 2,321,791；FY2026(Mar '26) 毛利 1,012,904（百萬日圓） | 頁面只渲染年度與TTM，讀不到季度欄。用TTM推算，與1,380,066相容（隱含上年同季約71,179），但未直接見到該數字。 |
| 部分相符 | klac | fl26 | CY2026 半導體製程控制系統營收中代工邏輯占比 | 0.63 % | somewhere close to 65-ish%, so I'll call it low to mid-60s as I look at 2026 | 法說只給「low to mid-60s」、約65%，0.63在區間內但非點值；未找到MarketBeat頁，改以Investing.com逐字稿核對 |
| 部分相符 | klac | g_oth26 | SPTS＋PCB 兩事業 CY2026 年增 | 0.25 % | With these combined products expected to grow over 25% in calendar 2026 | 25%以上與0.25一致但屬下限；來源涵蓋特殊製程、PCB與元件檢測三塊。未找到MarketBeat版，改核對同場Investing.com逐字稿 |
| 部分相符 | kyec | ai_25 | 2025 年 AI 加速器測試營收 ÷ 營收 | 0.275 % | AI相關產品在2025年已貢獻京元電子25%至30%的營收 | 區間0.25～0.3與基準0.275（中點）相符；但來源是「AI相關產品」，口徑較廣，且為媒體轉述大摩報告，非原報告。 |
| 部分相符 | kyec | pe_ttm | 近 12 月 P/E（鉅亨網，10-22） | 36.13 x | 本益比 38.19(10/28 收盤價 277.50;近四季EPS 7.27) | 頁面為10/28即時值38.19,與模型10-22的36.13差約5.7%;無10-22歷史值,無法逐日核對。 |
| 部分相符 | liteon | ai_25 | 2025 年 AI 相關營收 ÷ 營收 | 0.21 % | AI相關業務營收占比已突破20%（另有報導稱2025年約20%） | 指定uanalyze網址無法開啟，改核對股東會報導；只有「逾20%／約20%」，無21%，落在區間內 |
| 部分相符 | mediatek | fx26 | 3Q、4Q26 匯率（公司 3Q26 指引假設；FY2026 ASIC 換算用） | 32 NT$/US$ | forecasted exchange rate of 32 NT dollars to 1 US dollar（3Q26 指引假設） | 3Q26 基準 32 與講稿相符；講稿未給 4Q26 匯率，也沒有 30.5～33 區間。 |
| 部分相符 | mediatek | ss_asic28 | 賣方 FY2028 AI ASIC 營收（凱基） | 50 US$B | 2028年AI ASIC營收達500億美元的預估，占公司整體營收比重約67% | 數字與占比吻合，但報導只寫「法人」未點名凱基，也未找到凱基原報告；基準值50可對上，歸屬待確認 |
| 部分相符 | micron | cb_1 | 一般 DRAM 位元成長 FY27 | 0.12 % | industry DRAM bit shipments grow in the low-20% range (CY2027); HBM grows faster than conventional DRAM | 來源只給含HBM的產業low-20%，未給一般DRAM數字；0.12方向合理但無法直接核對，且為CY非FY |
| 部分相符 | micron | cb_2 | 一般 DRAM 位元成長 FY28 | 0.15 % | industry DRAM bit shipments grow in the low-20% range (CY2028); HBM grows faster than conventional DRAM through CY2028 | 來源只給含HBM的產業low-20%，無一般DRAM單獨數字；0.15未見於來源，且為CY非FY |
| 部分相符 | micron | sh26 | HBM 營收市占 2026 | 0.19 % | Micron HBM 營收市占：Q1 2026 21%、Q2 2026 18%（Counterpoint） | Q1 21%、Q2 18%與來源欄一致；來源無2026全年值，0.19 為推算，介於兩季間且僅涵蓋上半年。 |
| 部分相符 | micron | sh27 | HBM 營收市占 2027 | 0.2 % | Micron accounting for the remaining 20%(UBS 2027 HBM bit market share,三星與 SK 海力士各約 40%) | 20% 一致,但來源是位元(出貨)市占,模型列為營收市占,口徑不同。引用的是 Investing.com 對 UBS 報告的轉述,未見 UBS 原文。 |
| 部分相符 | microsoft | cons_cx29 | 共識資本支出 FY2029（MarketScreener，推定現金口徑） | 251.173 US$B | CAPEX 2029: 240,675(百萬美元) | 240.675B 對 251.173B,差約4.4%,超出容許範圍。量級一致,但快照日可能不同。 |
| 部分相符 | microsoft | cons_ebit27 | 共識營業利益 FY2027（MarketScreener） | 182.305 US$B | EBIT 2027: 182,411（百萬美元）；日曆頁 182,444 | 搜尋摘要約 182.4B，與 182.305 差約 0.06%。抓取被截斷，無法確認 10/08 快照。 |
| 部分相符 | microsoft | cons_ebit28 | 共識營業利益 FY2028（MarketScreener） | 217.245 US$B | EBIT 2028: 217,571（百萬美元）；日曆頁 217,504 | 搜尋摘要約 217.5B，與 217.245 差約 0.15%。抓取被截斷，無法確認 10/08 快照。 |
| 部分相符 | microsoft | cons_ebit29 | 共識營業利益 FY2029（MarketScreener） | 262.608 US$B | EBIT 2029: 262,864（百萬美元）；日曆頁 264,024 | 搜尋摘要約 262.9–264.0B，與 262.608 差 0.1–0.5%。抓取被截斷，無法確認 10/08 快照。 |
| 部分相符 | microsoft | cons_rev28 | 共識營收 FY2028（MarketScreener） | 467.955 US$B | Net sales 2028: 468,283（百萬美元）；日曆頁 468,554 | 搜尋摘要約 468.3B，與 467.955 差約 0.1%。網頁抓取被截斷，快照日期無法確認 10/08。 |
| 部分相符 | microsoft | cons_rev29 | 共識營收 FY2029（MarketScreener） | 569.502 US$B | Net sales 2029: 569,978（百萬美元）；日曆頁 570,057 | 搜尋摘要約 570.0B，與 569.502 差約 0.1%。抓取被截斷，無法確認 10/08 快照。 |
| 部分相符 | microsoft | g_q1_az | FY27Q1 Azure 年增指引 | 0.445 % | we expect revenue growth of approximately 45% in constant currency | 官方法說只給約45%固定匯率;0.445與0.45差1.1%。0.44~0.45區間僅見第三方稱9/2 8-K修正,未能核實。總帳檔本身查無公開版本。 |
| 部分相符 | mrvl | g_cus27 | 客製營收年增 FY2027 | 0.25 % | Custom revenue remains on track to grow more than 20% year-over-year in fiscal 2027 | Q4/Q1 法說只給「>20%」下限。Q2 新聞稿稱下半年客製顯著加速但沒給數字。基準0.25與上限0.35找不到來源。 |
| 部分相符 | mrvl | g_sw28 | 交換器年增 FY2028 | 0.8 % | scale-out switching on track to "more than double this year"（FY27）；FY28 僅稱 switching「biasing higher」；各業務朝 $1B 年化營收 | 法說未給FY2028交換器年增數字，80%無從核對；僅方向（FY28上調）與$1B run rate相符。 |
| 部分相符 | mrvl | q3_cog | Q3 FY27 通訊與其他年增（指引說法） | -0.135 % | For the third fiscal quarter, we expect revenue to decline in the low- to mid-teens percentage range both sequentially and year-over-year | 逐字稿只有「低至中段十幾個百分點」的年減與季減，沒有確切數字。-13.5%與區間-15%～-11%落在該說法內，但屬模型換算。未打開 Zacks。 |
| 部分相符 | mrvl | q4_gm | Q4 FY27 non-GAAP 毛利率 | 0.58 % | Q3 non-GAAP gross margin 57.5% to 58.5%;「We currently expect to maintain gross margin in this range in the fiscal fourth quarter」 | Q4 與 Q3 同區間屬實。基準 0.58 是區間中點，但來源區間為 0.575～0.585，模型 0.57～0.59 較寬；Zacks 未核對。 |
| 部分相符 | nanya_pcb | ai_abf | FY2026 AI 與資料中心 ÷ ABF 營收 | 0.65 % | AI and data center (800G switches) applications account for about 60% of ABF revenue | 來源約60%，模型0.65差約8%。60%在區間內；來源為7/27週報，非7/20報告，也非FY2026預估 |
| 部分相符 | nvidia | g28_gr | 公司 FY2028 營收年增 | 0.7 % | approximately 70%（FY2028 營收年增，受供給限制） | 官方新聞稿沒有FY28數字，逐字稿也未取得；只在二手報導看到「約70%」，與0.7一致。Quartz原文打不開。 |
| 部分相符 | nvidia | pa_amd | AMD Instinct 每顆 | 30 US$K | HSBC：MI355 ASP 由 USD25k 下修至 USD23k；MI455X：AMD 未公布官方價；MikeLongTerm 估 $45,000-$50,000 | HSBC $23K、MI455X 無官方價屬實；來源無30K基準，且部落格區間不全為20–45K |
| 部分相符 | nvidia | th_sh_hi | 命題：2027 年 NVIDIA 金額份額（上緣） | 0.8 % | ~80–81%（2025E）；~75%（2026E）；60–75% through 2028 | 80%是2025年值非2027；來源2027隱含≤75%，低於上緣0.8 |
| 部分相符 | nvidia | th_sh_lo | 命題：2027 年 NVIDIA 金額份額（下緣） | 0.75 % | NVIDIA retains 60–75% through 2028（Silicon Analysts，2026/4/13） | 無2027專數；60–75%是Silicon Analysts非Kaiso，上緣75%與0.75相近 |
| 部分相符 | nvidia | tx_f | 有效稅率（FY2028 起） | 0.17 % | For the full year fiscal 2027, NVIDIA expects GAAP and non-GAAP tax rates to be between 16.0% and 18.0% | 區間 16–18% 與中點 17% 吻合，但來源只涵蓋 FY2027，未提 FY2028 起，屬沿用外推。 |
| 部分相符 | oracle | lease_q | 未起租租約起租季數（FY27Q2 起平均） | 11 季 | expected to commence between the second quarter of fiscal 2027 and fiscal 2029(附註 6;未起租承諾 $288 billion,期限 15–19 年) | 10-Q 只寫起租區間 FY27Q2–FY2029,約 11 季。但沒有「平均季數」,11 季是區間總長,不是平均值。 |
| 部分相符 | oracle | p_spot | 現貨價格證據：H100 Silicon Data 指數 | 2.82 US$／GPU-hr | 官方頁面：2.89 USD/GPU-hour（截至 2026-10-26）；二手報導：2.82（2026-10-10） | 官網只顯示10/26的2.89，查不到10/06歷史值（需登入）。2.82只見於二手文章且日期為10/10，故無法確認。 |
| 部分相符 | oracle | pace_A | 情境 A：併網速度 | 3333 MW-IT／年 | expects to add nearly 1 GW of capacity in fiscal Q1 2027（Q4 FY26 電話會議，The Register 轉述）；Q1 實際交付 850 MW | 公司只預告 Q1 FY27 約 1 GW，未說每季都是 1 GW。×4 與 ÷1.2 是模型推算，來源沒有；Q1 實際 850 MW，Q4 約 400 MW。 |
| 部分相符 | quanta | nbsh_26 | 2026 年筆電營收 ÷ 營收 | 0.11 % | 1H26實際營收結構…筆電只剩約13%（估計）；公司：筆電占比降至低雙位數百分比 | 找不到財報狗10/8原文及凱基2026F 11%；僅見1H26約13%、公司稱低雙位數，凱基舊報告為16% |
| 部分相符 | quanta | pe_ttm | 近 12 月 P/E（10-22） | 14.31 x | 14.31 (25.52) 本益比 (同業平均) | 14.31 只出現在搜尋摘要（約7天前）。實開頁面顯示 14.90（9/23）與 14.03（10/27），無法確認 10-22 當日值。 |
| 部分相符 | quanta | px | 股價（2026-10-22 收盤） | 322 NT$ | 324.5（2026/10/22 09:51 盤中，漲跌 4.00、1.22%）；未找到 10/22 收盤價 | 盤中 324.5、4.00/1.22% 與前收 328.5 相容，但找不到 10/22 收盤 322，無法直接確認。 |
| 部分相符 | quanta | q_gm | 3Q、4Q26 毛利率 | 0.05 % | 毛利率為 5.02%(2Q26);預估全年利潤率將持穩於上半年水準 | 2Q26 為 5.02%,與 0.05 一致。來源只說全年持穩於上半年水準,沒給 3Q、4Q 數字。1H 平均約 4.9%(Yahoo 報導)。 |
| 部分相符 | samsung | cq_4 | 一般 DRAM 合約價季增 2026Q1 | 0.95 % | Contract prices rose approximately 93 to 98 percent compared to the previous quarter | 數字與模型相符(中點95.5%≈0.95),但只見轉載稿引述TrendForce,未找到原始新聞稿。 |
| 部分相符 | samsung | div_0 | 2026 下半年股利 | 32.5 兆韓元 | 3Q26 cash dividend: KRW 4,601 per share, totaling KRW 30 trillion（Q2 regular dividend 約 2.45 兆韓元，來自 sedaily 報導） | 三星官方確認3Q 30兆（含定期）。4Q 2.45兆官方尚未公布，僅可由2Q約2.45兆推算，30+2.45≈32.5。 |
| 部分相符 | samsung | dps_1 | 每股股利 2027 | 1500 KRW | KRW 9.8 T annually (regular), FY2024-2026;流通股 6,566,563,106 股 | 9.8兆÷6,567M≈1,492,與1,500差0.5%;但政策僅涵蓋2024–2026,2027年無公告,不宜標Verified |
| 部分相符 | samsung | dps_2 | 每股股利 2028 | 1500 KRW | KRW 9.8 T annually (regular), FY2024-2026;流通股 6,566,563,106 股 | 算式吻合(≈1,492),但2028年超出現行政策期;新政策預計2027年1月才決定,尚無來源 |
| 部分相符 | samsung | g_amd_27 | AMD Instinct 每顆 HBM 2027 | 360 GB | MI355X 288 GB；MI455X 432 GB | 區間上下限288/432有AMD官方支持；基準360為Assumed中間值，來源無此數 |
| 部分相符 | samsung | g_bw_26 | NVIDIA Blackwell（B200／B300） 每顆 HBM 2026 | 260 GB | B200: 192 GB HBM3E；B300: 288 GB HBM3E | B300 288 相符，但 B200 官方為 192 非 186；260 的混合權重未載明，無法驗證。 |
| 部分相符 | samsung | g_bw_27 | NVIDIA Blackwell（B200／B300） 每顆 HBM 2027 | 280 GB | B200: 192 GB HBM3E；B300: 288 GB HBM3E | B300 288 相符，但 B200 官方為 192 非 186；280 的混合權重未載明，無法驗證。 |
| 部分相符 | samsung | g_rb_27 | NVIDIA Rubin 每顆 HBM 2027 | 288 GB | Up to 288 GB of HBM4 per GPU | 官方只給 Rubin 規格「最高 288GB」，未分年度。2027 出貨量含 Rubin Ultra，其規格不在此頁。 |
| 部分相符 | samsung | g_rb_28 | NVIDIA Rubin 每顆 HBM 2028 | 288 GB | Up to 288 GB of HBM4 per GPU | 官方未提供 2028 年規格，只能確認 Rubin 現行規格為最高 288GB。 |
| 部分相符 | samsung | g_ru_26 | NVIDIA Rubin Ultra 每顆 HBM 2026 | 1024 GB | 1 TB HBM4E | 1TB=1024GB規格相符，但Rubin Ultra預定2027下半年，2026無此產品；僅找到轉載報導，未取得NVIDIA原文 |
| 部分相符 | samsung | g_ru_28 | NVIDIA Rubin Ultra 每顆 HBM 2028 | 1024 GB | 1 TB HBM4E | 單封裝1TB規格相符，但路線圖未列2028年Rubin Ultra（2028為Feynman），僅有延後傳聞 |
| 部分相符 | samsung | g_tpu_28 | Google TPU 每顆 HBM 2028 | 288 GB | TPU 8i HBM Capacity 288 GB;TPU 8t 216 GB | 288 GB 等於官方 8i 規格,但 8t 為 216 GB;2028 年各型號組合無來源。 |
| 部分相符 | samsung | hbm_mkt25 | HBM 市場 2025 | 35 $B | approximately $35 billion in 2025 | Micron 法說稿原文為 2025 HBM TAM 約 $35B，與基準 35 相符；Yole 約 $34B 的原始頁面兩次打不開，區間下緣未能核對 |
| 部分相符 | samsung | hbm_p26 | HBM 混合單價 2026 | 14.6 $/GB | HBM ASP is expected to rise 30% to 14.6 USD per GB in 2026 | BofA 2026 HBM 均價 14.6 美元/GB 與基準相符。但只找到媒體轉述，沒有 BofA 原報告。DigiTimes 的 12–12.8 與 16 也只見於二手整理。 |
| 部分相符 | samsung | hbm_p27 | HBM 單價 2027 變動（三情境相同；年約已議定） | 0.8 % | TrendForce：Blended ASP 2027 +121% YoY（原廠新聞稿）；UBS 約 +79%（前值 67%）；BofA HBM ASP 2027 +38%；JPM +54%（均為媒體轉述） | TrendForce 121% 已核對原文；UBS 79%、BofA 38%、JPM 54% 僅見媒體轉述；Micron 簽約說法僅見 JPM 轉述；基準 0.8 與 UBS 79% 差約 1%。 |
| 部分相符 | samsung | hbm_p28_A | HBM 單價 2028 變動（A） | 0.1 % | a further 25% increase in 2028(HBM blended ASP,JPMorgan) | 僅在 BigGo 轉述報導找到 JPM 2028 年 +25%,未取得 JPM 原報告。另一轉述的 2.8→3.4 美元/GB 約 +21%。基準 0.1 非 JPM 數字。 |
| 部分相符 | samsung | nq_6 | NAND 合約價季增 2026Q3 | 0.15 % | NAND Flash contract prices ... increase 10–15% QoQ(3Q26) | 3Q26 區間 10–15%,中點 0.125。模型 0.15 是區間上緣,與其他列用中點不一致。 |
| 部分相符 | samsung | nv_dc26 | NVIDIA 資料中心營收 2026（約 FY27） | 374 $B | Q1 FY27 資料中心 $75.2 billion；Q2 FY27 資料中心 $89.0 billion；Q3 總營收指引 $108.0 billion ±2% | Q1 75.2、Q2 89.0、Q3 總營收指引 108 均與 8-K 相符。Q3 資料中心 100 與 Q4 110 無來源，合計 374.2 只有約 164 為實績，不應標 Verified。 |
| 部分相符 | samsung | q4_dp | Q4 DRAM ASP 季增 | 0.1 % | Conventional DRAM contract prices projected to grow 10–15% QoQ in 4Q26 | 基準0.10等於來源區間下限（中點12.5%）；來源指傳統DRAM合約價，非三星ASP；模型下限0.05低於來源區間。 |
| 部分相符 | samsung | sh25 | HBM 營收市占 2025 | 0.17 % | Samsung HBM 營收市占：Q2 2025 15%、Q3 2025 23%、Q4 2025 22% | 區間端點 15%、22% 與來源相符，但 Counterpoint 現已把 Q3 修訂為 23%。基準 0.17 來源未載，也無全年數字。 |
| 部分相符 | samsung | sh26 | HBM 營收市占 2026 | 0.29 % | Counterpoint:三星 HBM 營收市占 1Q26 21%、2Q26 33%(SK hynix 50%、Micron 18%);HBM4 出貨預期下半年明顯放量、三星市占逐步上升 | 季度數字相符,但來源沒有2026全年市占,0.29與區間0.25~0.33查無出處。Counterpoint頁面內文未顯示,21%由 sedaily 轉述。 |
| 部分相符 | samsung | sh27 | HBM 營收市占 2027 | 0.38 % | UBS 2027 HBM 位元出貨占比：三星 41%、SK 海力士 39%、美光 20% | 來源是位元占比 41%，模型是營收占比 0.38，口徑不同，差約 7%。來源 SK 為 39%，與說明欄的 42% 不同。 |
| 部分相符 | samsung | u_amd_26 | AMD Instinct 顆數 2026 | 0.75 百萬顆 | about 600,000 AI GPUs in 2026（Morgan Stanley，2026/10/15 X 貼文轉述） | 只對上區間下限0.6M；基準0.75M為Assumed、無來源；MS原報告未公開，僅見轉述 |
| 部分相符 | samsung | u_amd_27 | AMD Instinct 顆數 2027 | 1.5 百萬顆 | 1 million MI455 units and 500,000 MI450 units in 2027 | 合計1.5M與基準相符，但僅見TechFlow轉述；MS 10月另有2027年1.3M的轉述（X貼文） |
| 部分相符 | samsung | u_bw_27 | NVIDIA Blackwell（B200／B300） 顆數 2027 | 2.4 百萬顆 | Net-new shipments in year 2027: ~2.4M | 只找到 AMCompute 預估 2027 新增約 2.4M，數字吻合；但「AC Research」是否即此機構無法確認。 |
| 部分相符 | samsung | u_hp_26 | NVIDIA Hopper（H200 等） 顆數 2026 | 0.5 百萬顆 | Hopper series’ share is forecast to drop from 10% to 7% | 只找到 Hopper 占 2026 高階 GPU 出貨 7%，沒有顆數。0.5 百萬顆無法從來源驗證。 |
| 部分相符 | samsung | u_rb_26 | NVIDIA Rubin 顆數 2026 | 1.7 百萬顆 | Rubin's share of NVIDIA's high-end GPU shipments is expected to decline from 29% to 22% | 22% 屬實，但新聞稿未給總量約 7.6M，1.7M 無法由來源直接驗證。KeyBanc 估 1.5M。 |
| 部分相符 | samsung | u_rb_27 | NVIDIA Rubin 顆數 2027 | 7 百萬顆 | shipments of Rubin and Rubin Ultra chips to approach 7 million units by 2027 | 只找到轉述，沒找到 MS 原報告。量級相符，但轉述口徑含 Rubin Ultra 且為「接近 7M」。 |
| 部分相符 | samsung | u_tpu_26 | Google TPU 顆數 2026 | 3.85 百萬顆 | MS: 3.2 million units (2026); GF: 4.504 million units (2026) | MS 3.2、GF 4.5 兩端點都找到;3.85 是兩者正中間,屬模型自行混合值,非來源數字。 |
| 部分相符 | samsung | u_tpu_27 | Google TPU 顆數 2027 | 5.5 百萬顆 | MS: 5 million units (2027); GF: 8.84 million units (2027) | MS 5.0、GF 8.84 兩端點相符;5.5 介於兩者之間且偏 MS,是模型混合值,非來源原文。 |
| 部分相符 | siliconomics | abfcap_ks_27 | ABF 產能指數 2027（FY2026 平均＝1） | 1.2 x | ABF 月產能由目前 4,000 萬片，提升至 2027 年 5,000 萬片（5,000/4,000＝1.25x） | 來源只有 4,000→5,000 萬（1.25x），沒有 FY2026 平均值。模型基準 1.2x 落在區間內，但與 1.25x 差約 4%，且此為法人推估。 |
| 部分相符 | siliconomics | asm_wyai_26 | 2026 年緯穎 AI 伺服器 ÷ 緯穎營收（CF09 單一值；緯穎、緯創模型同讀） | 0.48 % | AI 伺服器營收佔比將略高於 50%（2026-02 法說摘要）；2026-09-10 公司未設全年 ASIC 過半目標 | 公司說略高於50%，模型0.48略低；區間含0.50；未取得公司原稿，且曾有「略降」說法 |
| 部分相符 | siliconomics | asm_wyam | 緯穎 2026 年 ASIC 占 AI 伺服器營收 | 0.9 % | AI 伺服器中以 ASIC 伺服器為最大宗，約占 9 成（指 2025 年） | 約九成是 2025 年實績。2026 年只說 ASIC 仍為主，下半年 GPU 放量，未給 0.9。我只找到媒體轉述，未找到法說原稿。 |
| 部分相符 | siliconomics | cpo_A_26 | 情境 A：CPO 占 scale-out 光學金額 2026 | 0.01 % | scale-out CPO 交換器 2026 年 5,000 units(Citi) | 只找到 Citi 的 2026 台數,沒有金額占比 1%。Citi 原報告未公開,僅見媒體轉述。 |
| 部分相符 | siliconomics | cpo_A_27 | 情境 A：CPO 占 scale-out 光學金額 2027 | 0.05 % | scale-out demand ... 40,000 units(2027) | 2027 年 4 萬台與來源相符。5% 占比和 $2–3B 是模型自行換算,來源未載。報告為轉述。 |
| 部分相符 | siliconomics | cpo_A_28 | 情境 A：CPO 占 scale-out 光學金額 2028 | 0.1 % | 2028 scale-out 約 100,000 units(Citi) | 2028 年約 10 萬台與來源相符。10% 占比和 $6–8B 是模型換算,來源未載。報告為轉述。 |
| 部分相符 | siliconomics | cpo_B_26 | 情境 B：CPO 占 scale-out 光學金額 2026 | 0.01 % | CPO penetration in AI data center optical modules ~0.5% in 2026 (TrendForce) | TrendForce 為0.5%，模型1%約2倍；口徑是AI資料中心模組滲透率，非scale-out光學金額。LightCounting與MS原文未取得。 |
| 部分相符 | siliconomics | cpo_B_27 | 情境 B：CPO 占 scale-out 光學金額 2027 | 0.03 % | Volume production begins in 2026, with large-scale procurement expected in 2027–2028 (TrendForce) | TrendForce 只給2027–28放量的定性時程，沒有2027年數字；3%無法直接核對。報告全文付費。 |
| 部分相符 | siliconomics | cpo_B_28 | 情境 B：CPO 占 scale-out 光學金額 2028 | 0.05 % | 2027-2028 as the critical window for large-scale deployment (TrendForce) | 時程方向一致，但「2028年4–7%」與5%均無來源數字；LightCounting 2030約10%與MS原文均未取得。 |
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
| 部分相符 | skhynix | cap_0 | 資本支出 基準年 2026 | 48.7 兆韓元 | 2026년 투자 규모는 40조원 후반 수준으로 예상 (2026年投資規模預期在40兆韓元後段) | 2026Q2法說證實指引為「40兆後段」，與48.7及45～50區間量級一致；但共識48.7未找到來源，且Q3法說未重述2026年數字。 |
| 部分相符 | skhynix | cons_eps27 | 共識 EPS 2027 | 475721 KRW | EPS 2027: 469,819（valuation 頁）；476,971（盧森堡上市版）；464,639（印度版） | 不同時點的快照約 46.5–47.7 萬，模型 475,721 在範圍內。10-23 版本表格讀不到，無法逐點比對。 |
| 部分相符 | skhynix | cons_eps28 | 共識 EPS 2028 | 537981 KRW | EPS 2028: 505,883（valuation 頁）；533,797（盧森堡上市版）；511,356（印度版） | 快照差異大（50.6–53.4 萬），模型 537,981 比最高值略高。10-23 版本讀不到，無法確認。 |
| 部分相符 | skhynix | cons_tp | 市場共識目標價（賣方平均；低／高＝最低／最高） | 3096288 KRW | Number of Analysts 39；Average target price 3,096,287.92 KRW | 39位與均值3,096,288相符；最低1,535,000與最高4,700,000在頁面讀不到，無法核對（另一版本快照低標為1,575,000） |
| 部分相符 | skhynix | extra_eq | 基準年其餘增資 | 0 兆韓元 | 177.9 million ADS × US$149 = 約 US$26.5 billion（2026-07-09 定價）；3Q26 負債比 5%、借款 15.8 兆韓元 | ADS 發行於 7 月（Q3 內），邏輯上已含在 9-30 權益；但財報稿未列權益數字，也未見其餘增資為 0 的明文。 |
| 部分相符 | skhynix | extra_sh | 基準年其餘新股 | 0 百萬股 | 17,790,000 common shares represented by ADSs; 728,865,500 shares outstanding after the offering | ADS 新股 1779 萬股於 7/14 交割，早於 9-30，邏輯上已含。但未找到 9-30 股數原文，無法直接確認。 |
| 部分相符 | skhynix | g_amd_27 | AMD Instinct 每顆 HBM 2027 | 360 GB | MI355X 288 GB HBM3E；MI455X 432 GB HBM4 | 區間 288–432 的兩端有 AMD 官網佐證。基準 360 是兩端中點，屬 Assumed，來源上沒有。 |
| 部分相符 | skhynix | g_bw_26 | NVIDIA Blackwell（B200／B300） 每顆 HBM 2026 | 260 GB | 372 GB per Grace Blackwell Superchip（約 186 GB/GPU）；B300 288 GB HBM3e | 兩端點 186、288 可查證；260 GB 的 B200/B300 組合比重屬推導，來源上無此數。 |
| 部分相符 | skhynix | g_bw_27 | NVIDIA Blackwell（B200／B300） 每顆 HBM 2027 | 280 GB | 372 GB per Grace Blackwell Superchip（約 186 GB/GPU）；B300 288 GB HBM3e | 兩端點 186、288 可查證；280 GB 的組合比重屬推導，來源上無此數。 |
| 部分相符 | skhynix | g_rb_28 | NVIDIA Rubin 每顆 HBM 2028 | 288 GB | HBM4 capacity 288 GB (per Rubin GPU) | Rubin 現行規格 288GB 相符,但沒有 2028 年的來源,屬假設。 |
| 部分相符 | skhynix | g_ru_26 | NVIDIA Rubin Ultra 每顆 HBM 2026 | 1024 GB | 1TB of HBM4E memory (Rubin Ultra, due 2027) | 1TB來自媒體轉述，NVIDIA官方文章未載；Rubin Ultra 2027才出貨，2026年欄位不適用。 |
| 部分相符 | skhynix | g_ru_27 | NVIDIA Rubin Ultra 每顆 HBM 2027 | 1024 GB | 1TB of HBM4E memory (Rubin Ultra, due 2027) | 1TB與GTC 2026發布規格一致，但只有媒體來源；8月起有降至192/256GB的報導。 |
| 部分相符 | skhynix | g_ru_28 | NVIDIA Rubin Ultra 每顆 HBM 2028 | 1024 GB | 1TB of HBM4E memory (Rubin Ultra, due 2027) | 2028年規格無官方數字；1TB僅為2027款發布值，且有降規與延後至2028的報導。 |
| 部分相符 | skhynix | g_tpu_27 | Google TPU 每顆 HBM 2027 | 250 GB | TPU 8t: ~250 GB (est.); Codelabs: TPU 8t 216 GB, TPU 8i 288 GB | 250 GB 僅見第三方估計值；另一份 Google Codelabs 文件寫 8t 為 216 GB、8i 為 288 GB |
| 部分相符 | skhynix | hbm_sup27 | 2027 HBM 位元供給成長（三情境相同：產能已定） | 0.55 % | HBM bit shipments are projected to grow by 50–60% YoY in 2027 | TrendForce 50–60%（出貨）涵蓋0.55；上限0.63的JPM是2026–28位元需求CAGR，非供給，且僅見轉載 |
| 部分相符 | skhynix | nm_op_q | 非記憶體事業每季營業利益（基準年推估季） | 0 — | Semiconductor segment 131,895,033 (millions of Won), 100.0%; 3Q26 other products 1% of revenue | 6-K列單一半導體部門100%，CIS已轉AI記憶體；但3Q26有1%其他產品，未揭露非記憶體營益，0無法直接證實 |
| 部分相符 | skhynix | nq_5 | NAND 合約價季增 2026Q2 | 0.725 % | 70–75% QoQ（二手轉述） | 未取得TrendForce 2Q26原文，只見二手轉述70–75%，中點0.725與模型一致。 |
| 部分相符 | skhynix | nq_6 | NAND 合約價季增 2026Q3 | 0.15 % | NAND Flash contract prices expected to increase 10–15% QoQ | TrendForce為10–15%（中點0.125）；模型0.15取區間上緣，落在區間內但相對差約20%。 |
| 部分相符 | skhynix | nv_dc26 | NVIDIA 資料中心營收 2026（約 FY27） | 374 $B | Q1 FY27 資料中心 $75.2 billion;Q2 FY27 資料中心 $89.0 billion;Q3 總營收指引 $108.0 billion ±2% | Q1、Q2 與 Q3 總營收指引屬實,加上假設的 100 與 110 合計 374.2,與 374 吻合。Q3 資料中心 100 和 Q4 110 無來源,整列不宜標 Verified。 |
| 部分相符 | skhynix | other_y | 其他營收（每年） | 1.6 兆韓元 | 2025 營收 W97,147 十億、DRAM W74,904 十億、NAND W20,690 十億，推算其他約 W1,553 十億 | 來源欄未列出處。以 F-1 2025 年數字推算其他約 1.55 兆，與 1.6 差約 3%，且僅為 2025 單年。 |
| 部分相符 | skhynix | pb_2 | P/B 歷史中位 | 1.6 x | 2010–2025年底P/B：最低0.81（2022）、最高3.86（2025）；2013年2.00；其餘多在1.02–1.93 | 我算的中位數約1.605，與1.6相符、下限0.81也對；但納入2025年後上限是3.86，不是2.00，2.00只適用2010–2024。 |
| 部分相符 | skhynix | pub26_hi | 已發布 2026 HBM 市場（BofA） | 77.4 $B | HBM market size to reach 77.4 billion USD in 2026, a year-over-year increase of 124% | 數字77.4十億美元一致，但只在媒體轉述（BofA 8/7週報）找到；BofA原報告未公開，無法核對原文。 |
| 部分相符 | skhynix | pub27 | 已發布 2027 HBM 市場（BofA） | 152.8 $B | BofA expects HBM market to reach 152.8 billion USD in 2027, a growth of 98% | 數字與模型一致，但只在報導 BofA 8/7 週報的媒體文章找到，BofA 原始報告未公開取得。 |
| 部分相符 | skhynix | sh25 | HBM 營收市占 2025 | 0.58 % | 57% of revenue as of Q3(出貨量:62% as of Q2 2025) | 來源為Q3單季57%,模型0.58是2025全年口徑;Counterpoint現表已改為Q3 56%、Q4 57%,且此頁為SK海力士自家轉述。 |
| 部分相符 | skhynix | sh26 | HBM 營收市占 2026 | 0.52 % | Counterpoint：Q1 2026 58%、Q2 2026 50%；IDC（SK hynix F-1）：Q1 2026 56.4% | 來源只有單季數字（58/50/56.4%），沒有2026全年值。0.52落在Q1與Q2之間，量級一致但無法直接對上。Q1的58%出自二手報導。IDC數字來自SK hynix自己的F-1。 |
| 部分相符 | skhynix | sh27 | HBM 營收市占 2027 | 0.42 % | In 2027, Samsung is projected to take a 41% share, narrowly ahead of SK Hynix at 39%. Micron's share was forecast at 20%. | UBS 的 39%/41%/20% 與來源欄相符，但口徑是位元出貨。模型 0.42 是營收占比，差約 8%。 |
| 部分相符 | skhynix | tax | 有效稅率 | 0.23 % | Q1 2026 effective tax rate 21.8%;Q2 2026 稅前利益122,708,355百萬韓元、淨利93,922,593百萬韓元,推得約23.5% | Q1 21.8%有明文;Q2 23.5%需由稅前與淨利推算,未見明文。模型0.23介於兩者,是假設而非來源數字。 |
| 部分相符 | skhynix | u_hp_26 | NVIDIA Hopper（H200 等） 顆數 2026 | 0.5 百萬顆 | Hopper series’ share is forecast to drop from 10% to 7% | 只確認 2026 Hopper 占比 7%；無絕對顆數，0.5 百萬顆無法直接核對。 |
| 部分相符 | skhynix | u_rb_26 | NVIDIA Rubin 顆數 2026 | 1.7 百萬顆 | Rubin's share of NVIDIA's high-end GPU shipments is expected to decline from 29% to 22% | 只核到 22% 占比。新聞稿沒有 7.6M 總量,所以 1.7M(約 1.67M)無法完整驗證。 |
| 部分相符 | skhynix | u_rb_27 | NVIDIA Rubin 顆數 2027 | 7 百萬顆 | Rubin and Rubin Ultra chips to approach 7 million units by 2027 | 只有二手轉述。原報告未公開,口徑為 Rubin 加 Rubin Ultra,與模型僅計 Rubin 不同。 |
| 部分相符 | skhynix | u_tpu_26 | Google TPU 顆數 2026 | 3.85 百萬顆 | MS: 3.2 million units (2026); GF: 4.504 million units (2026) | MS 3.2、GF 4.5 兩端點屬實；3.85 是正中點而非「偏MS」，為自行混合值，非單一來源數字 |
| 部分相符 | skhynix | u_tpu_27 | Google TPU 顆數 2027 | 5.5 百萬顆 | MS: 5 million units (2027); GF: 8.84 million units (2027) | MS 5.0、GF 8.84 兩端點屬實；5.5 落在區間內且偏MS，為混合值。GF 3月曾寫 7.9 百萬 |
| 部分相符 | smci | ais_25 | FY2025 AI GPU 平台 ÷ 營收 | 0.7 % | GPU AI platforms ... represented over 70% of Q4 revenues（FY25 第四季法說，CFO） | 法說只說 FY25 Q4 逾70%，未給全年占比；口徑為季非年，0.70 僅是下限。逐字稿取自第三方轉載，非公司原檔。 |
| 部分相符 | smci | fpe | 前瞻 P/E（stockanalysis，10-23） | 7.06 x | Forward PE 7.29(頁面更新日 Oct 27, 2026) | 頁面現為10-27值,差3.3%;以10-23收盤32.94回推約7.05,接近7.06,但10-23當日頁面值無法直接核對 |
| 部分相符 | smci | pe_ttm | 近 12 月 P/E（stockanalysis，10-23） | 10.12 x | PE Ratio 10.46(頁面更新日 Oct 27, 2026) | 頁面現為10-27值,差3.4%;10-23收盤32.94÷EPS 3.26≈10.10,與10.12吻合,但10-23當日頁面值無法直接核對 |
| 部分相符 | tsmc | cap_end25 | CoWoS 月產能 2025 年底 | 77.5 千片／月 | 台积电CoWoS月产能于年底接近翻倍，达7.5万至8万片 | TrendForce 2025年底7.5–8萬片，中點77.5吻合；但只找到2024-11《科創板日報》轉述，沒找到工商時報2025-12原文。 |
| 部分相符 | tsmc | cap_end26 | CoWoS 月產能 2026 年底 | 130 千片／月 | TSMC’s monthly CoWoS capacity could reach a record 120,000 to 140,000 wafers in 2026 | TrendForce 轉述經濟日報 12–14 萬，130 為區間中點；MS 原報告未找到，僅見轉載站(12–13 萬)，不採用 |
| 部分相符 | unimicron | ai_pcb | FY2026 HDI＋PCB 的 AI 占比 | 0.65 % | 上半年AI相關應用產品營收占比已逾60％，預期下半年可望提升至65～70％（PCB 下半年 65～70%） | 來源是PCB上半年逾60%、下半年65–70%，模型是全年0.65，口徑不同；未取得工商時報本站原文，改用同集團時報資訊同日報導 |
| 部分相符 | unimicron | bk_eps27 | 券商 EPS FY2027（2026-10 兩家平均：美系 37、本土 41.10） | 39.05 NT$ | 美系券商 2026-2028年EPS 21/37/64.8元(2027年=37元) | 經濟日報只證實美系2027 EPS 37元；本土41.10與平均39.05未見於該文，無法核對。 |
| 部分相符 | unimicron | mx_abf | FY2026 ABF 載板 ÷ 營收 | 0.53 % | ABF載板營收比重從上季的49%拉高至52% | 來源只有2Q26的52%（1Q 49%），無FY2026全年值；0.53在區間內但口徑不同。工商時報法說原文未找到，改用自由財經。 |
| 部分相符 | unimicron | mx_bt | FY2026 BT 載板 ÷ 營收 | 0.09 % | 占9％的BT載板季減3％、年減4％（2Q26 產品營收結構） | 2Q26 BT 占 9% 與 0.09 一致，但這是單季數字，模型為 FY2026；工商時報 7/29 文章無此數，數字出自中時新聞網；自由時報寫 10%。 |
| 部分相符 | victory_giant | jpm_sh | 摩根大通：勝宏 AI／HPC PCB 份額 1H25（Prismark） | 0.138 % | 2025年上半年…13.8%,全球第一(图注来源:弗若斯特沙利文) | 13.8%數字相符,但AI/HPC份額圖註來源是弗若斯特沙利文,非Prismark。 |
| 部分相符 | wistron | b_eps27h | 法人 FY2027 EPS（樂觀說法中點） | 24.85 NT$ | 富邦 2027 EPS 24.03 元；統一 25.3 元；凱基 26.4 元 | 找不到8/10原文，只核對旺得富8/5文。上限26.4相符，但下限23.3未出現（最低為24.03），中點24.85為推算。 |
| 部分相符 | wistron | k_racks | 凱基：2026 年 GB 與 VR 出貨櫃數 | 15 千櫃 | AI伺服器系統L10出貨與營收年增1倍至1.5萬櫃 | 只找到4/14媒體轉述的匿名本土投顧報告，量級相符；未找到凱基5/8原報告，也未提VR。 |
| 部分相符 | wistron | px | 股價（2026-10-22 收盤） | 190 NT$ | 成交價 190、昨收 195、漲跌 -5（-2.56%），日期 10/22（Goodinfo）；富邦 K 線頁同為 190.00 元 ▼5 (-2.56%) | 財報狗頁已更新為10/28收盤194.5，查不到10/22；Goodinfo同日收盤190、-2.56%相符 |
| 部分相符 | wistron | q4g | 4Q26 合併營收季增 | 0.05 % | FY2026 營收 Avg 4.1T（High 4.2T／Low 4.0T，精確值 4.07T），No. Analysts 7；前三季累計 2兆8,908億、3Q26 約 1.15兆 | 共識與家數相符。以4.1T推得4Q約+5%，用4.07T則約+2.7%。來源未直接列4Q季增。 |
| 部分相符 | wistron | q_wgm | 3Q、4Q26 緯穎毛利率 | 0.088 % | 1H26 毛利率 8.41%；2Q26 毛利率 9.26% | 來源只有1H26、2Q26實績，且Q2受記憶體代採購與NRE墊高；未見3Q、4Q指引，0.088為估計值，無法直接比對。 |
| 部分相符 | wistron | q_wopm | 3Q、4Q26 緯穎營業利益率 | 0.068 % | 1H26 營益率 6.79%;2Q26 營益率 7.27% | 1H26 6.79%、2Q26 7.27% 屬實,0.068 約等於 1H26。來源未提供 3Q、4Q26 的數字或指引,無法直接核對。 |
| 部分相符 | wistron | wai_26 | 2026 年緯穎 AI 伺服器 ÷ 緯穎營收 | 0.48 % | AI 伺服器營收佔比將略高於 50%（2026 法說摘要）；凱基 46% 未找到 | 法說摘要為二手整理，原始法說 PDF 與 siliconomics、凱基 46% 皆未取得。模型 0.48 介於 46% 與略高於 50% 之間。 |
| 部分相符 | wistron | wgm | 緯穎毛利率（FY2027 起） | 0.085 % | FY2025 毛利率 8.3%（富果）；1H26 毛利率 8.41%、2Q26 毛利率 9.26%（優分析） | 來源只列歷史毛利率，與 0.085 同量級（對 1H26 8.41% 差約1%）。FY2027 起的數字來源未載，屬模型假設。 |
| 部分相符 | wiwynn | ais_26 | 2026 年 AI 伺服器 ÷ 營收 | 0.48 % | AI伺服器佔比今年預期將超過一半（MoneyDJ 2026-02-26 法說報導）；另一則財訊報導稱今年佔比會略降於去年的 50% | 公司說「超過50%」，沒有給確切數字。模型 48% 略低於此。優分析原文找不到，3Q26 各半一事由 9/10 MoneyDJ 證實。 |
| 部分相符 | wiwynn | asic_mix | 2026 年 ASIC 占緯穎 AI 伺服器營收 | 0.9 % | 去年 AI 伺服器中 ASIC 伺服器為最大宗，約占 9 成 | 約九成是2025年數字，2026只說ASIC仍為主；下半年GPU新平台出貨，2026占比未揭露。優分析原文未找到，用法說報導核對。 |
| 部分相符 | wiwynn | q_gm | 3Q、4Q26 毛利率 | 0.083 % | 第二季毛利率9.26%；預期下半年ASIC比重提升，即便產品組合改變，毛利率變化不大，不過還需觀察NRE貢獻程度 | 來源只有定性說法「下半年毛利率變化不大」，沒有3Q、4Q數字。1H26 8.41%、2Q26 9.26%相符。0.083屬模型推估，無法直接驗證。 |
