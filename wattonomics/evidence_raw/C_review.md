# C 段：獨立核對（A_us_cloud、B_labs_china_totals、B_delay、B_summary）｜2026-10-28

核對者：未參與 A／B 製作的代理。方法：先抓原文，再對照 CSV。WebFetch 只用 CSV 或摘要中出現過的網址，沒有自行拼網址。被擋的頁面記為「無法核對」，沒有繞過。未跑 pytest／LibreOffice，除本檔外未改任何檔案。

列識別：`A#n`＝A_us_cloud.csv 第 n 行（第 1 行是表頭）；`B#n`＝B_labs_china_totals.csv 第 n 行；`D#n`＝B_delay.csv 第 n 行；`A.md`／`B.md`＝兩份摘要。

## 0. 統計

- 抽查 **84 筆**（承重列優先；同一來源的連續列算一筆）。
- 判定：**屬實 62／小誤 14／錯誤 1／無法核對 7**。
- 被擋、未讀到原文的網址（未繞過）：
  - microsoft.com FY25Q4 IR 頁
  - paperjam.lu（Google 用電）
  - spglobal.com（Google 用電）
  - boardroomalpha（Broadcom 8-K）
  - SEC：Amazon Q3'25 Ex.99.1、CoreWeave 10-K、Oracle Q1FY27 Ex.99.1
  - aol.com（Amazon Q4'25 法說）
  - anthropic.com/news/higher-limits-spacex
  - insurancejournal.com（Jupiter）
  - invezz.com（Morgan Stanley）
  - investors.coreweave.com（Q2 新聞稿）
  - epoch.ai CSV 直接下載（curl 被代理擋；改用 WebFetch 讀到）

## 1. 錯誤（1 筆）

| 列 | 判定 | 正確值 | 說明 |
|---|---|---|---|
| A#33 Alphabet「SpaceX 租用 Google 容量」153–226 MW-IT | **錯誤** | 營運者＝SpaceX／xAI，使用者＝Google | 方向寫反了。DCD〈Google to pay $920 million to SpaceX monthly〉：Google 每月付 SpaceX $920M，租約 110k 顆 GPU，期間 2026-10 至 2029-06；SynMax 寫明地點在 Colossus 2（Southaven）。B#49 的方向是對的。這一列不能計入 Google 的營運者 GW，否則和 xAI 1.4 GW 重複計算。顆數換 MW 的算法（110k ÷ 487k／718k 顆每 GW ≈ 226／153）是對的。 |

## 2. 小誤（14 筆，承重的標 ★）

| 列 | 判定 | 正確值 | 說明 |
|---|---|---|---|
| ★A#19 Anthropic 於 Azure | 小誤 | 上限 1 GW（不是「1 GW＋選擇權再 1 GW」） | computeprices 原文：$30B Azure 承諾，加上「contract up to 1 GW more」。全案上限就是 1 GW。B#33「至多 1GW」是對的。 |
| ★A#45 Anthropic–AWS 5 GW | 小誤（標記） | 標 [Interested-party] | 這一列的 value 是 5,000 MW，GW 數來自 Anthropic 自述和 The Information，不是 10-Q。整列標 [Verified] 不恰當。數值本身由 AI Weekly（The Information）確認 AWS 5 GW。 |
| ★A#21 Microsoft 曆年 2026 資本支出 175B | 小誤（說明） | 175B 屬實；FY27Q1 >$50B 屬實（MarketBeat） | CFO Dive 寫的差額原因有兩項：資料中心耐用年限由 15 年延長為 25 年，以及租賃改列營業租賃。CSV 只寫了後者。Hood 原話是 "Outside of this useful life impact…remain unchanged"。 |
| ★B#3 OpenAI 10 GW | 小誤 | 「近 90 天簽了 3 GW（含 AWS 2 GW）」 | The Decoder 原文是 "Three of those 10 gigawatts"，不是「>3GW」。「DCD 判斷為簽約」這句的原頁被擋，無法核對。 |
| ★B#47 Colossus 2 | 小誤 | 946 MW-IT（2026-06-15）屬實 | Epoch 頁面沒有「1,531MW IT 2027」，只引 S-1 下一期「>400 MW、≥22 萬顆 GB300」。1,531 的出處要補。 |
| ★A.md／B.md D1 依據「26%＋10%≈36%」 | 小誤（方法） | 26% 是**容量**比例；10% 是**專案數**比例 | Sightline 新聞稿：26% of expected capacity delayed，另有 10% of projects 延後 COD。兩者單位不同，不能直接相加。D1 取 35% 仍在 2026 預測 30–50% 區間內，結論可保留，但依據要改寫成「26%（容量）～30–50%（2026 預測）取中下」。 |
| ★A#116 Epoch「Other」2,074 MW | 小誤（說明） | 數值屬實（1,813.74＋260.23） | CSV 寫「不含中國」，沒有出處。Epoch 擁有者表只有 Microsoft、Google、Amazon、Meta、Oracle、CoreWeave、Colossus、Other 八類，沒有中國類，所以中國買家的 NVIDIA／AMD 晶片很可能就在 Other 裡。這影響第 5 節的加總。 |
| A#16 Fairwater Wisconsin 369 | 小誤 | as_of 2026-05-15 | 2026-10-05 是頁面日期，不是估計日期。Epoch：369 MW-IT，衛星影像 2026-05-14。 |
| A#58 Hyperion 5 GW | 小誤 | Entergy 為 7 座複循環燃氣電廠 | Quartz 寫 seven，CSV 寫「7–10 座」；「租金 2029-06-01 起付」也不在 Quartz 原文。Meta 發言人的說法（2030 年前 2 GW、5 GW 沒有完成日）屬實。 |
| B#16 Ports Pike | 小誤 | 全園 8 IT-GW；NVIDIA 首期 4.25 IT-GW，另 3.75 IT-GW 為選擇權；2028 起分期 | NVIDIA 新聞稿沒有寫「首期 800MW」。 |
| B#36 TeraWulf×Anthropic | 小誤 | 首期交付 2027H2，2028 初全部交付 | TeraWulf Q2 新聞稿有寫交付時程，CSV 寫「未載」不對。A#114 寫的 2027H2 是對的。 |
| D#12 電力變壓器 | 小誤 | 低值依據不成立 | 「2020 年前 24–30 個月」和 Reuters／Kitco 不符：原文寫 2020–21 年約 1 年。中值 30（128 週）屬實。 |
| D#7 GE Vernova | 小誤（措辭） | "in many projects" | CEO 原話是「許多案子」渦輪不是卡關項，不是「多數」。其他數字都屬實：新單約 3 年、2029–30 年只剩約 10 GW 檔期、8-K 列 Q2 積壓 53 加檔期協議 63，共 116 GW。 |
| A#30 vs B#30 Broadcom 8-K 3.5 GW | 小誤（一致性） | — | 同一份 8-K，A 標 [Interested-party]，B 標 [Verified]，要統一。原文被擋，數值無法核對。 |

## 3. 屬實（62 筆）

| 列 | 判定 | 說明 |
|---|---|---|
| A#10–12 Microsoft Epoch 晶片 861.5／2,015.7／1,955.22 | 屬實 | Epoch CSV：801.05＋60.44；1,866.22＋149.52；2026Q1 只有 NVIDIA 一項，標 Incomplete。 |
| A#23–24 Google 1,298.3／2,953.2 | 屬實 | 991＋307.27；2,223＋730.19。 |
| A#38–39 Amazon 616.0／1,672.5 | 屬實 | 321.25＋294.7；832.48＋840。 |
| A#51–53 Meta 523.5／1,338.7／1,135.79 | 屬實 | 428.1＋95.37；1,079.17＋259.55。 |
| A#67–68 Oracle 208.9／650.2 | 屬實 | 184.68＋24.17；586.29＋63.88。 |
| A#77–79 CoreWeave 交叉 279／799 | 屬實 | Epoch 168.82、484.3 × 1.65。 |
| A#25 Google 4,873（4,134–5,611） | 屬實（Derived） | 2,953.19 × 1.4／1.65／1.9，算對。 |
| A.md 各家換算 | 屬實（Derived） | 各家 × 1.65 都算對：MSFT 1.42／3.33（2.82–3.83）、Meta 0.86／2.21（1.87–2.54）、Oracle 0.34／1.07、Amazon 1.02／2.76、Amazon NVIDIA 1.37。Trainium 比值 1,220 ÷ 840＝1.45。「晶片功耗×1.65→IT、×1.2→設施」寫得清楚。 |
| A#3–5 Microsoft FY26Q2–Q4 每季約 1 GW | 屬實 | DCD；口徑未說明，CSV 標 unknown 是對的。 |
| A#6 兩年翻倍、FY26 88 座 | 屬實 | MarketBeat："roughly double its overall capacity in two years"。 |
| A#7–8 Microsoft 12 GW→>38 GW（2032），約 1/3 AI；不含 neocloud | 屬實 | AI Weekly 轉述 Bloomberg，標為記者報導，處理正確。 |
| A#15 Fairwater Atlanta 636 MW-IT（2026-06-14） | 屬實 | Epoch。 |
| A#17 Fairwater Wisconsin 全站 2,263 MW-IT，2028Q2（2028-05-13） | 屬實 | Epoch。 |
| A#18 Abilene 鄰近 900 MW | 屬實 | W.Media：900 MW，含現場電廠，全站 2.1 GW，2027 年中首棟。Walsh 的引述裡沒有 900 這個數字，是報導寫的。 |
| A#29／B#29 Anthropic×Google TPU | 屬實 | 原文："up to one million TPUs"、"well over a gigawatt…in 2026"。 |
| A#35 Amazon 3.8 GW（2024-10 至 2025-10） | 屬實 | DCD：Jassy 說。8-K 原頁被擋。 |
| A#37 Amazon 2027 年底電力容量翻倍 | 屬實 | Fortune 是轉述，沒有逐字引語；「本人說」宜改成「報導轉述」。 |
| A#41 New Carlisle 398／626／910／1,037、2028 1,925 | 屬實 | Epoch 時序完全一致。1,037 是 rule-of-thumb 估計。 |
| A#42 New Carlisle 2.2 GW、30 棟 | 屬實 | 原文 "up to 2.2 GW of electricity"，屬設施（用電）口徑。報導沒有指名是誰說的。 |
| A#46／B#22 OpenAI–AWS Trainium 2 GW，Trainium4 2027 | 屬實 | SiliconANGLE。 |
| A#50 Amazon 2026 資本支出 220B（原 200B，記憶體漲價） | 屬實 | Fortune 以自己的話寫原因，不是 Jassy 原話。 |
| A#54 Meta 2024 資料中心用電 18,061,781 MWh→2.06 GW | 屬實 | 2023 年 14,975,435 MWh，換算正確。 |
| A#55 Prometheus 471 MW-IT（2026-08-31）、676／868／1,022 | 屬實 | 現場燃氣 138.8／200 MW。 |
| A#57 Hyperion 0 MW、一期 1,632 MW-IT（2028-01-01） | 屬實 | Epoch。 |
| A#59 El Paso | 屬實 | 原文 "1 gigawatt of compute capacity"、BlackRock 80%、2028 起、約 $14B。 |
| A#63 Meta 2026 資本支出 135–145B；Q3 34.91B | 屬實 | 8-K Ex.99.1。 |
| A#64–65 Oracle Q1FY27 交付 850 MW、>30 萬顆 | 屬實 | MarketBeat（Ex.99.1 被擋），原文寫 "AI capacity"。 |
| A#69／B#7 Abilene 618 MW＝6／8 棟＝75%、Q1 交付 13.1 萬顆 | 屬實 | MarketBeat。Epoch 設施 1.2 GW、2026Q4 完工，也屬實。 |
| A#72–75 Stargate 四站設施 2.0／2.2／1.3／1.4 GW、2028Q4 | 屬實 | Epoch〈where the US sites stand〉明寫是設施電力；Vantage 自報 Shackelford 1.4 GW IT。注意：這是 2026-04 的資料；B#10 引 Epoch 較新時程，Shackelford 首棟已延到 2027-08，A 列寫「首棟 2026 年底」已經過時。 |
| A#76 Oracle FY27 資本支出 90–95B、淨額 ≤70B | 屬實 | MarketBeat。 |
| A#81–82、84–85、94 CoreWeave 1.5 GW、>1.85（原 >1.7）、3.7→4.2 GW、35–39B | 屬實 | MarketBeat。4.2 GW 是 CEO 在法說會上說的。 |
| A#87–88 Helios 133 MW critical IT（約 200 MW gross）、全案 526 | 屬實 | Galaxy；Phase II 260 MW 於 2027H1。 |
| A#95 Nebius active 約 170 MW＝IT | 屬實 | 股東信："power being consumed by IT equipment"。 |
| A#104–105 IREN 4×50 MW IT、累計 0.3／0.8 GW IT | 屬實 | 8-K。 |
| A#114 TeraWulf 102 MW critical IT；Anthropic 401 MW、20 年、2027H2 起 | 屬實 | 新聞稿。 |
| A#117 五大超大雲 71%（H100e，2025Q4） | 屬實 | Epoch。 |
| B#2 OpenAI 2025 年底 1.9 GW（2023 年 0.2、2024 年 0.6） | 屬實 | DCD；口徑未說明。標 [Interested-party] 是對的。 |
| B#4 OpenAI 2027 "low double-digit" GW | 屬實 | Epoch 原文有。「2030 年 30 GW」不在 Epoch 頁，未核到。 |
| B#6 第三方估 OpenAI、Anthropic 各 5–6 GW | 屬實 | Epoch。 |
| B#23 Nvidia Vera Rubin 5 GW | 屬實 | SiliconANGLE；未給時程。 |
| B#24 AMD 6 GW＝處理器功耗、2026H2 起 | 屬實 | DCD：原文 "power consumption of 6GW"。 |
| B#33 Azure 至多 1 GW、$30B | 屬實 | — |
| B#44 Anthropic ≥14.8 GW、$517B | 屬實 | AI Weekly：只有 AWS 5 GW 和 Google／Broadcom 5 GW 給了 GW，另外 4.8 GW 沒有歸屬。 |
| B#45 Anthropic 2026 年底約 5 GW、2027 年底約 10 GW | 屬實 | Newsquawk 轉述 NYT（匿名、投資人溝通）。 |
| B#46 xAI 0.4→1→1.4 GW、年底 >2 GW | 屬實 | DCD；Musk 原話 "closer to 10GW…than 5GW"、"I don't think we're going to achieve 20GW"。 |
| B#48／D#25 Colossus 1 約 198 MW 自發＋300 MW 電網；Colossus 2 約 1.3 GW 表後、無電網 | 屬實 | Cleanview：約 59 台臨時機組；Doosan 1.9 GW。300 MW 電網是 Cleanview 的估計。 |
| B#49 Google 租 SpaceX：$920M／月、約 11 萬顆、2026-10 至 2029-06、90 天可退 | 屬實 | DCD；SynMax 稱這是補自建延遲的權宜。 |
| B#51 Rystad 中國 32→40→60 GW、AI／HPC 2026 年 39% | 屬實 | 口徑未說明。15.6 GW 是 [Derived]，算對。 |
| B#51 第二來源 SemiAnalysis 中國 >24 GW、字節約 1/5、幾乎全租用 | 屬實 | TNW。 |
| B#58 Humain 1.9 GW（2030）、6.6 GW（2034） | 屬實 | DCD。 |
| B#63 Epoch 全球 AI 約 30 GW（2025Q4，設施，晶片×約 2.5） | 屬實 | Epoch 原文 "~2.5x multiplier"。頁面沒有說明是否含中國，也沒有列出晶片廠清單。 |
| B#64 約 1,600 萬 H100e 運行、累計售出約 2,000 萬、三大實驗室占 20–30% | 屬實 | Epoch。 |
| B#65 IEA 485→950 TWh、AI 465 TWh；+17%／+50% | 屬實 | Enlit 沒有標明 485 是哪一年；415（2024）× 1.17 ≈ 485，與 2025 一致。換算約 55 GW 正確。 |
| B#67 McKinsey 156 GW（2030）、增量 125（受限 78、加速 205） | 屬實 | 2025 年約 31 GW 由 156−125 推得，算對。這是需求基期，不是實際裝機。 |
| B#70 TrendForce 155 GW、年增 29% | 屬實 | 2025 年約 120 GW 算對。 |
| B#72／D#15 Sightline 16 GW、約 140 案、5 GW 施工、11 GW 公告、30–50% 恐延、26% 容量延遲、10% 專案延後 | 屬實 | Utility Dive 新聞稿。「2025 年上線近 6 GW」出自 Axios，屬實。 |
| D#2–3 併網：北維州約 7 年（84 個月）、全美約 4 年（48 個月） | 屬實 | Construction Owners 引 JLL 2026 Outlook。144 個月出自別的來源。 |
| D#5 LBNL：約 4.3 年（52 個月）、>5 年（60）、PJM >7 年（84） | 屬實 | 轉述來源。 |
| D#6 歐洲：荷蘭至 10 年、德國與北維州相近、都柏林暫停至 2028、20% 專案有延遲風險 | 屬實 | Latitude 轉述 IEA。 |
| D#8 Siemens 積壓 69 GW、季接單 15 GW、季出貨 6 GW、交期 ≥3 年 | 屬實 | 2.9 年 [Derived] 算對。 |
| D#11 升壓變壓器 >160 週（2024 年 143 週） | 屬實 | 33／37 個月換算正確。 |
| D#13 斷路器 125 週（2023 年 77 週） | 屬實 | 換算 29／18 個月。 |
| D#14 建物工期 12–18（Sightline）、18–24 個月（Construction Owners） | 屬實 | — |
| D#16 Oracle 部分站 2027→2028 | 屬實 | 原因為勞工與材料短缺；Oracle 否認，未指明哪些站。 |
| D#17 Jupiter 首次供電 2027→2029、合約 2026-08-15 不可行 | 屬實 | SemiAnalysis；只作佐證。 |
| D#19 Nebius 新澤西目標 4 個月、實際 10–11 個月 | 屬實 | SemiAnalysis。 |
| D#20 SynMax 約 4.5 GW 延到 2027 年底 | 屬實 | — |
| D#21 Stargate UAE 200 MW 沒有試運轉公告、Mubadala 目標 Q3 2026、Reuters 報導改為分散多站 | 屬實 | The Gulf Tape。 |
| D#22 Nadella "chips sitting in inventory"、"warm shells" | 屬實 | TechRadar（2025-11-05）。引述未超過 15 字。 |
| D#23 Jassy "single biggest constraint"、"several quarters" | 屬實 | Latitude。 |
| D#24 Fairwater Wisconsin 4 月稱上線、5 月對地方報紙稱未上線 | 屬實 | AI Weekly 補充：Guardian（8 月）估 Microsoft AI 容量約 1.2 GW，對照公司稱的 5 GW，Microsoft 回應稱估計不準。這對 A 段 Microsoft 2025 年 3.3 GW-IT 是一個反向證據，應該收進來。 |

## 4. 無法核對（7 筆）

| 列 | 說明 |
|---|---|
| A#2 Microsoft FY25 >2 GW | IR 頁被擋。DCD 提到的 2 GW 沒有說明年度口徑。 |
| A#27 Google 2025 年平均負載 4,840 MW | Paperjam、S&P 都被擋。42.4 ÷ 8,760＝4.84 算對。但 CSV 寫 2024 年是 30.6 TWh，與標題「增 37%」不符：42.4 ÷ 1.37 ≈ 30.9。記憶中 Google 2024 年資料中心用電約 30.8 TWh，未核。若 42.4 只含資料中心，則 4.84 GW 成立。 |
| A#30／B#30 Broadcom 8-K 3.5 GW | 原頁被擋。The Information 稱 Google／Broadcom 合計 5 GW。 |
| A#36 Amazon 2025 全年 3.9 GW | AOL 被擋。DCD 只確認 Q4 >1.2 GW，而且出自 AWS 永續報告，不是法說會。 |
| A#79 等 CoreWeave 10-K 850 MW／3.1 GW | SEC 被擋。 |
| B#34 Anthropic×SpaceX >300 MW | Anthropic 頁被擋。Cleanview 估 Colossus 1 供電約 500 MW，量級相符。 |
| D#18 Denton | DCD 原頁沒有「原指引 2025 年底 250 MW」，也沒有寫 MW；月數 [Derived] 無法驗證。 |

## 5. 兩個已知矛盾

### 5.1 Google：2025 年平均負載 4.84 GW（設施，全部負載）vs AI 年底 IT 4.9 GW

**判斷：** 兩者不是真正互斥。差異主要來自四點，其中 (a) 和 (b) 影響最大。

- **(a) 年底與年平均的落差。**
  - Epoch 晶片功耗從 1,298 MW（2024Q4）升到 2,953 MW（2025Q4）。
  - 2025 年平均約 2,125 MW，× 1.65 ≈ 3.5 GW-IT。
  - 拿年底存量去比全年平均，本來就會高估約 1.4 GW。
- **(b) 「售出」到「通電」有落差。**
  - Epoch 的數字是按售出累計的。
  - 它自己在 gradient update 中說，部署量大約落後一季：售出約 2,000 萬 H100e，運行約 1,600 萬。
- **(c) 實際用電低於銘牌。**
  - 平均負載通常只有 IT 銘牌的 0.7–0.8。
- **(d) TPU 系統的 IT ÷ 晶片 比值可能低於 1.65。**
  - repo 也把下限設為 1.4。
- **另一種可能：** Google 環境報告的用電可能不含第三方機房（colo 或租用），但這點未核。

**反推檢查：**
- 年增 37%，所以 2025Q4 的用電 run-rate 約為 4.84 × 1.15 ≈ 5.6 GW 設施。
- 扣掉非 AI 負載 1.5–2.0 GW，AI 用電約 3.6–4.1 GW 設施。
- ÷ PUE 1.09 ÷ 負載率 0.8，得 AI IT 銘牌約 4.1–4.7 GW。

**建議：**
- 2025 年底 Google AI IT 取 **基準 4.3、區間 3.9–4.9 GW**。
- 基準的算法：TPU 取 × 1.45，NVIDIA 取 × 1.65，得 2,223 × 1.45 ＋ 730 × 1.65 ≈ 4.4，再酌減部署落差。
- 4.9 只作上緣。

### 5.2 Amazon NVIDIA：Epoch 推 1.4 GW-IT vs repo 約 0.6

**判斷：** 真值較可能在兩者之間，偏 Epoch 的下半段。

- **repo 偏低的原因：**
  - repo 的 NVIDIA 只由 Omdia 2024 年 Hopper 19.6 萬顆（273 MW）加上 [Assumed] 的 Blackwell 構成。
  - 沒有計入 2023 年以前的 A100／V100 存量（P4d 等）。
  - 也沒有計入 2025H2 開始交付給 OpenAI 的 GB200／GB300（$38B，報導稱 2026 年底前全部部署）。
- **Epoch 偏高的原因：**
  - Epoch 的 2024Q4 NVIDIA 是 321 MW 晶片，約合 46 萬顆 H100 等值功耗，大約是 Omdia 2024 年採購量的 2 倍以上。
  - 這表示 Epoch 用營收占比去分配晶片歸屬，而且累計值沒有扣除舊卡退役。
- **Epoch 的 Amazon 自身一致性良好：**
  - Trainium 840 MW 對 repo 1,220 MW-IT，比值 1.45。
  - 所以問題集中在 NVIDIA 的歸屬與退役，不是換算係數。

**建議：**
- 2025 年底 Amazon NVIDIA 取 **基準 1.0、區間 0.7–1.4 GW-IT**。
- Amazon 2025 年 AI 合計取 **基準 2.2、區間 1.9–2.6**。這與 A.md 的 2.2 一致，只把上緣由 2.8 降到 2.6。

## 6. 加總檢查（2025 年底，換到設施）

### 6.1 A 段各玩家 2025 年底 IT 估計（取 A.md 的值，區間取低／高）

| 玩家 | 基準 | 低 | 高 |
|---|---|---|---|
| Microsoft | 3.3 | 2.8 | 3.8 |
| Google | 4.9 | 4.1 | 5.6 |
| Amazon | 2.2 | 1.8 | 2.8 |
| Meta | 2.2 | 1.9 | 2.5 |
| Oracle | 1.1 | 1.0 | 1.5 |
| CoreWeave | 0.85 | 0.85 | 0.85 |
| Nebius | 0.17 | 0.17 | 0.17 |
| 其他新雲 | 0.5 | 0.3 | 0.8 |
| **合計** | **15.2** | **12.9** | **18.0** |

× PUE 1.2，得設施 **18.3（15.5–21.6）GW**。

### 6.2 B 段各項（設施）

| 項目 | 設施 GW | 依據 |
|---|---|---|
| xAI | 0.6–0.8 | Epoch Colossus 308 MW × 1.65 ≈ 0.5 GW-IT；SpaceX 自報 2025 年中 0.4、2026Q1 1.0 |
| 主權、企業、學研（Epoch Other 扣新雲） | 1.8–3.0 | [Assumed] |
| 中國 2025 年底 AI | 8–12，基準 10 | Rystad 32 GW × AI 占比約 0.3–0.35；2025 年占比 Rystad 未給，[Assumed] |

### 6.3 比較

- 由下而上合計：**基準約 31.4 GW**（區間 25.9–37.4）。
- Epoch 約 30 GW：基準差 **＋1.4 GW（＋5%）**，區間是 −4.1 到 ＋7.4。

### 6.4 這個「吻合」是巧合，要拆開看

**(1) 換算倍數不同。**
- Epoch 是晶片 × 2.5 到設施。
- A 段是晶片 × 1.65 × 1.2 ＝ × 1.98。
- 把 Epoch 八類擁有者的晶片加總，2025Q4 共 **11.50 GW**：2,015.74＋2,953.19＋1,672.48＋1,338.72＋650.17＋484.3＋308＋2,073.97。
  - 用 A 段倍數，只得 **22.8 GW 設施**；
  - 用 Epoch 的 2.5 倍，得 **28.7 GW**，與 30 GW 相符。
- 所以 A 段的美國玩家換到 Epoch 口徑，應該再乘約 1.26。
- 佐證：Epoch 自己的 Stargate 站點，設施 ÷ IT 約為 1.26–1.44（2,000／1,400、2,200／1,750、1,400／988、1,300／902），也高於 1.2。

**(2) 中國可能重複計算。**
- Epoch 的 Other 很可能已經包含中國買家的 NVIDIA／AMD 晶片。
- 在 Epoch 口徑下再加上 Rystad 的中國 8–12 GW，會變成 37–41 GW，明顯超過 30。
- 這表示要嘛 Epoch 的 30 GW 不含華為等中國自製晶片，要嘛 Rystad 的「AI／HPC」範圍比 AI 晶片廣。

**建議：**
- C0 主比對分兩層：
  - 「Epoch 可比口徑」：晶片 × 2.5，不加 Rystad 中國；
  - 「全球含中國自製晶片」：另加中國非 NVIDIA 的部分，昇騰約 1–1.3 GW 晶片層。
- PUE 1.2 改列為敏感度，範圍 1.2–1.5。
- 在倍數統一之前，不要用「31 對 30」去宣稱由下而上與 Epoch 吻合。
