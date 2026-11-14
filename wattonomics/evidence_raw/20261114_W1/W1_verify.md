# W-1 證據表獨立核對（Wattonomics v0.3.2 用）

核對日 2026-11-14｜核對代理：未參與 W1_evidence 製作；只讀與查證，未改任何 repo 或原檔
對象：`W1_evidence.md/.csv`（68 列）；工作單 company-models origin/main `wattonomics/docs/workorders/20261114_Wattonomics_v0.3.2.md`
限制：本環境 curl 被 proxy 擋，改用 WebFetch／WebSearch 開原文；SEC 6-K、CoreWeave IR、Galaxy、Hut 8、Applied Digital、IREN、Nscale、Crusoe、Dell 的原文可開。標「(既有知識)」者是核對代理對 2025 年公告的已知內容，本輪沒有重新打開原文，入基準前建議再補連結。

## 一、總計

| 判定 | 列數 |
|---|---|
| 屬實 | 40 |
| 小誤（已更正） | 12 |
| 錯誤 | 0 |
| 無法核對 | 16 |
| 合計 | 68 |

## 二、逐列判定

基準欄的意思：「基準」＝實績或已簽約，可入基準；「基準-去重」＝已簽約，但屬對七家合約，只用作 X2 去重下限，不算七家以外；「基準-X4」＝出租方的 MW，以承租方或最終使用者計，不另加；「Sens」＝只入敏感度；「參照」＝交叉檢核或參照列。

| # | 判定 | 基準 | 核對結果（出處） |
|---|---|---|---|
| 1 | 屬實 | 基準 | CoreWeave IR Q3'26 新聞稿：「Added more than 200 MWs of active power to reach over 1.7 GW」。可升級為 Verified（investors.coreweave.com）。口徑未明 |
| 2 | 屬實 | Sens | IR：「continue to expect more than 1.85 gigawatts of active power at year-end」（指引） |
| 3 | 屬實 | 基準 | IR：「Grew total contracted power more than 900 MW to over 4.6 GW」；其中國際 >1.3 GW。升級為 Verified |
| 4 | 屬實 | 基準 | 公司口徑 14.5B 出自二手（stocktitan 引 10-Q）；IR 現金流量表購置 PP&E 9,441M 屬實。基準取現金流 9.44B 比較穩健 |
| 5 | 屬實 | Sens | IR：「full year 2026 capital expenditures of $37 billion to $39 billion」 |
| 6 | 屬實 | 基準 | IR：backlog 約 132.5B（RPO 加其他承諾）；RPO 132.3B 為二手 |
| 7 | 屬實 | 基準-去重 | 22.4B＝11.9B（2025-03）＋4B（2025-05）＋6.5B（2025-09-25）。CoreWeave 2025-09-25 公告（既有知識）；簽約日應補 2025-09-25 |
| 8 | 屬實 | 基準-去重 | 14.2B（2025-09-30 8-K，至 2031；既有知識）；21B 於 2026-04-09 簽約（CoreWeave IR 標題與 Bloomberg 2026-04-09），至 2032 |
| 9 | 屬實 | 參照 | 二手數字與 IR 一致的部分無矛盾，負債 46.9B 未在 IR 頁另核 |
| 10 | 屬實 | 基準 | SEC 6-K 股東信：「more than 4 GW contracted」；年底 5 GW 是目標（Sens）。沒有揭露在役 GW |
| 11 | 屬實 | 基準 | 6-K：Q3 capex「approximately $7.3 billion」 |
| 12 | 屬實 | Sens | 6-K：約 27B，高於原 20–25B |
| 13 | 屬實 | Sens | 6-K：2027 capex 40–45B，「more than 1 GW of new capacity」；營收 12–13B；融資覆蓋 75–85% |
| 14 | 屬實 | Sens | 6-K：2026 已承諾預付 10B；2027 為 17–20B |
| 15 | 屬實 | 基準 | 6-K：可轉債 5.75B、ATM 約 1.0B、ABS 775M、Citi 3.5B；現金 7.4B；RPO 38.7B |
| 16 | 小誤 | 基準-去重 | 簽約日是 2025-09-08（列中 09-10 是文章日）。17.4B 最高可到 19.4B 屬實（既有知識）。合約本身沒有寫 300 MW；Vineland 站點 300 MW 是站點規模，不是合約 MW。6-K：NJ 過半容量年底交付，其餘 2Q27，屬實 |
| 17 | 屬實 | 基準-去重 | Nebius Q3'25 股東信（2025-11-11）：Meta 3B、5 年（既有知識） |
| 18 | 屬實 | 基準-去重 | 2025-11-03 公告（列中 11-04 是文章日）：9.7B、5 年、預付 20%、Childress Horizons 1–4 共 200 MW 關鍵 IT、向 Dell 採購約 5.8B（既有知識）。IREN Q1 FY27 新聞稿：Horizon 1 已對 Microsoft 投入服務 |
| 19 | 無法核對 | Sens | 原文（GlobeNewswire 2026-11-05）只寫剩餘 FY27 capex 15B，其中 9B 已有來源、6B「targeted, subject to market conditions」；「由 25–30B 下修至約 20B」原文沒有出現 |
| 20 | 屬實 | 基準 | 原文：backlog 20.3B（QoQ +22%，約三分之二為投資級對手方）；>4B ARR 目標延到 2027-01、可能 02。「每 MW 約 22M」原文沒有出現 |
| 21 | 小誤 | 基準-X4 | 原文日期是 2026-10-07。250 MW 是「fully operational critical IT load」（2026-10-01 RFS），口徑應改為 IT，不是推定。季末為 175 MW |
| 22 | 屬實 | 基準-X4 | PF1 對 CoreWeave 400 MW、約 11B（2025-08 加碼 150 MW；既有知識）。MW 已含在 CoreWeave 的 4.6 GW 內，不可重複計 |
| 23 | 小誤 | 基準-去重 | Delta Forge 2 的 210 MW、15 年、約 5.2B 屬實，但簽約公告日是 2026-06-08（10-07 是財報日）。承租方是「tier-one investment grade hyperscaler」，與七家重疊極高。另見第四節補漏 |
| 24 | 屬實 | 基準 | TeraWulf Q3'26 原文：「102 MW of revenue-generating critical IT capacity」；CB-4／CB-5 共 336 MW 建設中。本文只點名 Anthropic；Fluidstack、Core42 是 2025 年舊租約（既有知識） |
| 25 | 屬實 | 基準（X3：Anthropic） | 原文：「20-year lease with Anthropic for approximately 401 MW of critical IT」，約 19B。注意：Google 擔保是「Anticipated Project Financing」，融資尚未完成；建造成本指引每 MW 關鍵 IT 8–10M 屬實 |
| 26 | 屬實 | 基準 | 原文：購置與存出保證金 2,230.8M；另有資產收購 234.7M；股權淨募 1,199.8M；現金含受限 2.36B |
| 27 | 屬實 | 基準（X3） | Cipher Q3'26 原文：Barber Lake 延長到 20 年，總合約約 9.1B，增量約 5.2B；承租方只寫「a leading AI lab」，MW 未揭露。原租約是 2025-09-25 對 Fluidstack 的 168 MW 關鍵 IT、10 年、約 3B、Google 擔保 1.4B（新聞稿標題已核到） |
| 28 | 屬實 | 基準-去重 | 2025-11-03：AWS 15 年、約 5.5B、300 MW（forklog 等多家一致），分兩期 2026-07 到 Q4 交付。MW 口徑未明。可移出「待核」 |
| 29 | 屬實 | Sens | 原文：「development pipeline today totals approximately 5.9 GW」；「電網 3.4 GW＋燃氣」的拆分原文沒有看到 |
| 30 | 小誤 | 基準-X4 | Galaxy 官方新聞稿日期是 2026-07-06（列中 07-07 是 The Block）。133 MW 關鍵 IT、CoreWeave 承諾 526 MW 關鍵 IT、800 MW「gross power」、Phase II 2027H1 起交付、年均租金 >1B 皆屬實；可升級為 Verified。Phase II＝260 MW 原文沒有寫出 |
| 31 | 小誤 | 基準（X3：Anthropic） | Hut 8 原文日期是 2025-12-17（02-25 是文章日）：「245 megawatts (MW) of IT capacity」、15 年、7.0B（含三次 5 年續約約 17.7B）；Google 擔保租金；首個 data hall 預計 2027Q2；最高 85% LTC（JPM、GS）。可升級為 Verified |
| 32 | 屬實 | 基準（gross，去重） | Crusoe 原文：「6GW+ of gross contracted capacity」、「1 GW of gross capacity delivered and operational」，口徑明寫 gross。Abilene 為 OpenAI 站點，去重比例需判斷 |
| 33 | 小誤 | 參照 | 原文是「initial closing of its anticipated $3.9 billion Series F」，不是 3.9B 全數到位；估值 30.9B post-money；co-lead 為 Atreides、Mubadala Capital、Valor |
| 34 | 屬實 | 基準-去重（無 MW） | 2025-11-03 BusinessWire 公告「multibillion-dollar」、數萬顆 GPU 含 GB300；金額與 MW 未揭露（既有知識） |
| 35 | 無法核對 | 參照 | 本輪未開原文 |
| 36 | 屬實（二手） | Sens | dealroom 2026-09-28：「nearly 1GW of power」、「investor memo projects … 1.3GW … more than 10 sites in 2026」。這是投資人備忘錄的預測，不是簽約 |
| 37 | 屬實 | 參照 | 募資 1.5B、估值 18B+ 屬實；結案約 2026-09-03（aiweekly）。660M 營收是備忘錄預測 |
| 38 | 小誤 | 基準-去重 | Series C 2B 是 2026-03-09 宣布（Bloomberg、Computer Weekly），不是 04-18。Narvik 擴約的原文是 Nscale 2026-04-14 新聞稿：230 MW 園區、>30,000 Rubin、2027 交付，MW 口徑未明。新聞稿沒有提到 OpenAI 或 Stargate Norway；估值 14.6B 未核 |
| 39 | 小誤（需更新） | 基準-去重 | 約 200,000 顆 GB300 於 2025-10-15 屬實（Nscale 新聞稿）；14B 是媒體數字。**2026-11-05（TNW 引 FT）**：Microsoft 放棄 Texas Barstow 第二期 700 MW 選擇權（原定 2027 年底起），並已退出英國 Essex 站點；Texas 既有合約（240 MW 園區、約 104,000 顆 GB300，2026Q3 起）維持。200k 應視為已縮減 |
| 40 | 小誤 | Sens | 原文（middleeastainews 2026-09-01）只說在 LEAP 2026「signed a strategic partnership」、250 MW 資料中心，沒有寫是約束性合約還是 MOU；120,000 顆晶片與營收分成都沒有核到。依工作單只入 Sens |
| 41 | 無法核對 | 參照 | 研調部落格轉述，沒有一手來源 |
| 42 | 屬實 | 不入（七家：OpenAI） | Epoch 頁（最後更新 2026-07-22）：「CURRENT IT POWER 0MW」；200 MW 預估 2026-12、7.6B；1 GW 中位 2028-01、37.9B；晶片可能為 G42 所有、OpenAI 使用。站點已簽約在建，但**尚未投運**，使用者是 OpenAI，屬七家，應全額去重 |
| 43 | 屬實（標題） | 不入（同上） | 多家（Gulf News、Reuters）報導第一期 200 MW 在 2026 年完工；2026-11 的現況本輪沒有找到 |
| 44 | 無法核對 | Sens | 目標，不是簽約 |
| 45 | 無法核對 | Sens | 洽談中、非約束性 |
| 46 | 屬實 | Sens（且屬七家 xAI） | HUMAIN 新聞稿 2025-11-19 是「framework agreement」，500 MW+，沒有時程、沒有寫歸屬。不是確定站點合約 |
| 47 | 小誤 | Sens（待判斷） | AGM 2026-06-19：「will commission the first 120 MW … by end 2026」，是指引，尚未投運；首批 GB300 相當於 >75,000 顆 H100（推論口徑），不是列中的 >200k。屬自建，沒有簽約相對方 |
| 48 | 屬實 | Sens | 110B 承諾有多處引用，是多年承諾 |
| 49 | 無法核對 | Sens | — |
| 50 | 小誤 | Sens | 計畫是 2026-06-29／30 宣布（DCD、Asiae、edaily），09-02 是 aiweekly 轉載日。屬計畫，不是簽約 |
| 51 | 無法核對 | 參照 | 本輪未開 NVIDIA 原文 |
| 52 | 無法核對 | Sens | — |
| 53 | 小誤 | 基準（實績，低重疊） | Jiji 2026-01-23：KDDI 堺 AI 資料中心已開始營運（購地 100 億日圓；4 月起主要供製藥業使用）。48 MW 只見 mingtiandi 標題，Jiji 沒有寫 MW；日期應補 2026-01-22／23。屬商用電信業者，歸「主權」需判斷 |
| 54 | 無法核對 | Sens | — |
| 55 | 無法核對 | Sens | 本輪未開委員會頁 |
| 56 | 屬實 | 參照（OEM） | Dell 部落格 2026-09-02：AI 伺服器營收 16.4B |
| 57 | 屬實 | 參照 | AI 訂單 60.9B、backlog 95B、AI 客戶 5,000→6,500 |
| 58 | 屬實 | 參照 | FY27 AI 伺服器「raised by $14 billion to $74 billion」；Q3 AI 伺服器 19B |
| 59 | 無法核對 | 參照 | — |
| 60 | 屬實 | 參照 | 多家一致：19.95B |
| 61 | 屬實 | 參照 | 多家一致：FY27 72–82B |
| 62 | 無法核對 | 參照 | — |
| 63 | 無法核對 | 參照 | — |
| 64 | 無法核對 | 參照 | — |
| 65 | 無法核對 | 參照 | — |
| 66 | 屬實 | 參照 | 48,710M／40,313M，合計 89,023M。註：本季有一家客戶由 ACIE 改列 Hyperscale，前期已重編 |
| 67 | 無法核對 | 參照 | 轉載文沒有管理層「roughly half」的直接引語；約 45% 是作者自算 |
| 68 | 屬實 | 參照 | 2025-08 法說，Kress 稱 FY26 主權 AI 營收 >20B（既有知識） |

## 三、可入基準的列

- **七家以外、可入基準（實績或已簽約）**
  - CoreWeave：1、3、4（現金流 9.44B）、6
  - Nebius：10、11、15
  - IREN：20
  - TeraWulf：24、25（Anthropic 401 MW IT，依 X3）、26
  - Cipher：27（依 X3，但 MW 待補）
  - Hut 8：31（Anthropic 經 Fluidstack，245 MW IT，依 X3）
  - Crusoe：32（gross 口徑，需換算並去重）
  - KDDI：53
  - CoreWeave、Nebius、IREN、Crusoe 的總量仍要依 X2 扣除對七家的部分
- **只作 X2 去重下限（已簽約、對七家）**
  - 7（OpenAI 22.4B）、8（Meta 14.2B＋21B）
  - 16（Microsoft 17.4–19.4B）、17（Meta 3B）
  - 18（Microsoft 9.7B、200 MW IT）
  - 23（210 MW 投資級 hyperscaler）
  - 28（AWS 300 MW）
  - 34（Lambda 對 Microsoft，無 MW）
  - 38、39（Nscale 對 Microsoft，39 需依縮約更新）
  - 42、43（Stargate UAE，OpenAI）
- **X4（出租方 MW 歸承租方，不另加）**
  - 21、22（Applied Digital 對 CoreWeave）
  - 30（Galaxy 對 CoreWeave）
  - Fluidstack 列 36 的約 1 GW 與 25、27、31 重疊
- **只入 Sens**
  - 2、5、12、13、14、19、29、36、40、44–50、52、54、55
- **參照或 OEM 交叉檢核**：其餘各列

## 四、主要更正與補漏

1. **Nscale 對 Microsoft 已縮約**（2026-11-05 FT 報導）：Microsoft 放棄 Barstow 第二期 700 MW 選擇權並退出 Essex。列 39 的 200k GB300 不宜原樣作為去重下限，Texas 約 104k 顆仍然有效。
2. **Stargate UAE 尚未投運**：Epoch 2026-07 顯示 0 MW，200 MW 預估在 2026-12 才上線。使用者是 OpenAI，屬七家，應全額去重，不構成七家以外的主權基準。
3. **「主權已簽約」的實際狀態**
   - HUMAIN×xAI 只是 framework agreement，而且 xAI 屬七家。
   - Together×HUMAIN 是 strategic partnership。
   - Reliance 120 MW 是年底投運的指引。
   - 南韓、EU、HUMAIN 1–6 GW 都是計畫。
   - 主權線目前真正已投運、而且屬七家以外的只有 KDDI（MW 待核）。
4. **口徑更正**
   - Applied Digital 250 MW 確定為關鍵 IT。
   - Hut 8 245 MW 為 IT capacity。
   - Crusoe 6 GW／1 GW 原文明寫 gross，基準需依 X5 換算（Galaxy 的 0.66 比例可作參照）。
5. **日期更正**（列中多為文章日，應改為原始公告日）
   - 列 16：2025-09-08
   - 列 18：2025-11-03
   - 列 21：2026-10-07
   - 列 23：2026-06-08
   - 列 30：2026-07-06
   - 列 31：2025-12-17
   - 列 38：2026-03-09（Series C）、2026-04-14（Narvik）
   - 列 50：2026-06-29／30
   - 列 53：2026-01-22／23
6. **其他小誤**
   - Crusoe 3.9B 只是 initial closing。
   - TeraWulf 的 Google 擔保針對「預期中」的專案融資，尚未完成。
   - IREN「capex 25–30→20B」原文沒有出現，原文是剩餘 15B，其中 6B 未籌。
7. **證據表漏列（X3、X4 去重需要）**
   - **Cipher–Fluidstack Barber Lake**：168 MW 關鍵 IT、10 年、約 3B、Google 擔保（2025-09-25）。這是列 27 的 MW 基礎，原表寫「MW 未揭露」。
   - **TeraWulf–Fluidstack Lake Mariner**：CB-4／CB-5 約 200 MW＋（2025-08，Google 擔保）。
   - **TeraWulf–Fluidstack Abernathy TX 合資**：168 MW 關鍵 IT、25 年、約 9.5B（2025-10）。
   - 這三批容量都屬 Fluidstack／Anthropic 鏈，與列 36 重疊；本輪只核到新聞標題（既有知識），入基準前要開原文。
   - **Applied Digital 租約總量**：五個園區共約 1.41 GW 關鍵 IT、約 36B（Q1 FY27 新聞稿）。除對 CoreWeave 的 400 MW 外，其餘承租方都是未具名「investment-grade hyperscaler」，即 PF2、Delta Forge 1、Delta Forge 2，與七家重疊極高。
   - **IREN**：Q1 FY27 後與未具名「Leading Frontier AI Lab」擴約，0.3 GW（IT）的計畫支撐 >4B ARR；客戶是不是七家未揭露，按 X2 處理。
8. **建議升級為 Verified 的列**（已開到公司或主管機關原文）：1、3、5、6、10–15、20、21、24–27、30、31、32、56–58。
