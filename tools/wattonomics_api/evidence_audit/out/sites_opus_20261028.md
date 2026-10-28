# Wattonomics 證據逐列回查｜sites_opus_20261028

- 資料夾：wattonomics/evidence_candidates；列數 409；來源網址 322 個（可自動取得 259）
- 判定：不符 25、口徑不同 45、小誤 180、原文找不到 27、無法取得 88、屬實 44
- 用量：input 1,437,155／output 147,035 tokens；估計 US$4.34（claude-opus-5-5，Batch 五折）

| 檔案 | 不符 | 口徑不同 | 小誤 | 原文找不到 | 無法取得 | 未判定 | 屬實 |
|---|---|---|---|---|---|---|---|
| S_sites_20261028.csv | 25 | 45 | 180 | 27 | 88 | 0 | 44 |

## 需處理（不符／口徑不同，前 60 列）

| 判定 | 檔案:列 | 項目 | 值 | 理由 |
|---|---|---|---|---|
| 不符 | S_sites_20261028.csv:5 | Microsoft Fairwater Atlanta (Fayettevill | 636 MW | 636 MW IT 屬實，但全園區 2026-06 才達 636 MW，2025-10 僅 321 MW。屬模型估算，標 |
| 不符 | S_sites_20261028.csv:7 | Microsoft Fairwater Atlanta (Fayettevill |  MW | 2026-05 試運轉屬實，但 2026-06-14 原文稱 Building 4 已運轉，7 月仍列在建不符。原文未提 |
| 不符 | S_sites_20261028.csv:12 | Microsoft Fairwater Wisconsin｜Fairwater  | 369 MW | 369 MW IT 屬實，但原文為 2026-04-16 上線，非 2026-06；450 MW 報導不在原文；Epoc |
| 不符 | S_sites_20261028.csv:14 | Microsoft Fairwater Wisconsin｜第二座資料中心（相鄰 |  MW | 第二座相鄰在建、預計 2028 完工屬實，原文確無 MW。但原文是 2024–2028 威州總計約 47 億美元，非此期 |
| 不符 | S_sites_20261028.csv:23 | Stargate Abilene Campus 1｜已運轉 B1–B4（Epoc | 421 MW | 421MW IT與Q4預測843MW相符；但頁面更新日為2026-07-17而非10-19。421MW時間軸點為2026 |
| 不符 | S_sites_20261028.csv:28 | Google Mesa AI Campus (Project Redhawk)｜ | 183 MW | 原文 Building 1 於 2025-07-01 上線時 IT 為 90MW；183MW 是 2026-08-15  |
| 不符 | S_sites_20261028.csv:35 | AWS Ridgeland Campus｜Building 5（累計） | 284 MW | 284 MW 累計與「缺冷卻」相符，但原文預估 2027-02-09 上線，非 2026。屬 Epoch 預測，標 [V |
| 不符 | S_sites_20261028.csv:41 | AWS Madison Mega Site, Canton｜全園區 | 1000 MW | 原文日期為2025-07-30，非2026-09-02。1 GW是州政府準備吸引的園區級別，並非AWS公布的規劃MW。2 |
| 不符 | S_sites_20261028.csv:64 | Stargate Lighthouse（Vantage Port Washing | 1300 MW | 1.3GW為Vantage總裁Dana Adams所言，屬實；但原文完工年為2028非2027。原文無Epoch AI說 |
| 不符 | S_sites_20261028.csv:85 | Nebius Vineland Campus｜Phase 1（已運作部分） | 50 MW | 50MW IT、57k H100e、7月新冷卻機未連線、使用者可能為Microsoft皆屬實；但原文稱2025年12月仍 |
| 不符 | S_sites_20261028.csv:130 | Nexus Hubbard Campus｜Phase 1（首期） | 500 MW | 500MW與年內上線相符，但原文日期為2026-03-30，非推估的10-06，差約半年。TDLR工期與Eagle Po |
| 不符 | S_sites_20261028.csv:133 | Nexus Hubbard Campus｜全園區遠期擴建 | 7700 MW | 7.7GW與原文相符，但文章日期為2026-03-30而非10-06。7.2GW申請量與Eagle Point達FID原 |
| 不符 | S_sites_20261028.csv:151 | Campus AI Fouju（MGX–Mistral 巴黎區 1.4GW AI | 1400 MW | 1.4GW為「至終」容量且未說口徑，相符；但原文的2028初營運指Mistral設施，非全園區。RTE與環評說法不在本文 |
| 不符 | S_sites_20261028.csv:155 | Campus AI Fouju（MGX–Mistral 巴黎區 1.4GW AI |  MW | 原文第二期為2030–2033，非2029；文章日期8/13；約9座及DC1/2/4許可原文無；標題誤填 |
| 不符 | S_sites_20261028.csv:178 | Meta Los Lunas Campus｜全園區 | 87 MW | 87MW／33億美元是2026/3/1至7/28的估值；原文最新為2026/10/26的118MW、45億美元。原文無2 |
| 不符 | S_sites_20261028.csv:195 | Google Council Bluffs Southlands Campus｜ | 237 MW | 237 MW IT 與更新日 2026-09-03 相符，但 2024-11 僅 Building 1 上線（79 MW |
| 不符 | S_sites_20261028.csv:202 | Google Project Zodiac (Fort Wayne)｜全園區 |  MW | 原文園區面積為 892 英畝（併入 728 英畝），不是約 858 英畝。12 棟、第三方 1,200／1,035 MW |
| 不符 | S_sites_20261028.csv:234 | Core Scientific Pecos AI Campus｜全園區 | 1500 MW | 1.5GW 總電力、1GW 可出租、逾200英畝均符合。原文未提 AMD，僅提 CoreWeave 為公司整體主要租戶。 |
| 不符 | S_sites_20261028.csv:252 | Crusoe–Tallgrass Cheyenne AI Campus（Proj | 10000 MW | 10GW與2.7GW屬實，但原文主體是Crusoe，未提Google或Jupiter Star。原文也無2031，僅說2 |
| 不符 | S_sites_20261028.csv:269 | DataVolt Oxagon AI Campus｜全園區 | 1500 MW | 1.5GW 規劃屬實，口徑未說明。但 2028 僅指首批 100MW 可用，全園區上線時點原文未揭露，expected_ |
| 不符 | S_sites_20261028.csv:270 | DataVolt Oxagon AI Campus｜第一期（Phase 1） | 360 MW | 360MW 首期動工屬實。原文 2028 僅指首批 100MW 可用，未說 360MW 何時上線；300MW 早期報導不 |
| 不符 | S_sites_20261028.csv:317 | Envision Galaxy Base (遠景烏蘭察布星河基地)｜全園區 | 2000 MW | 超過2GW為設計可擴展規模，未給完工時點；8/6投運的是園區，不是2GW全數上線，expected_online填202 |
| 不符 | S_sites_20261028.csv:357 | Racks Central Johor AI Campus (RCJM1–4)｜ | 510 MW | 510 MW 全園區合計與原文相符，原文未說明 IT 或設施口徑。但 2027 Q1 RFS 僅指 RCJM1 的 70 |
| 不符 | S_sites_20261028.csv:372 | DayOne Johor Campus (Nusajaya / Kempas)｜ | 240 MW | 原文目前（2026-10-02 更新）IT 電力為 473MW。240MW 是 2026 年 1～4 月的時點值。此為  |
| 不符 | S_sites_20261028.csv:408 | Microsoft–G42 Olkaria AI Campus｜全園區（遠期擴建 | 1000 MW | 1GW與三分之一電網原文有。但文章日期為2026-05-07，非09-07；原文未見G42稱可擴至1GW，1GW為報導及 |
| 口徑不同 | S_sites_20261028.csv:8 | QTS Fayetteville（Microsoft 承租）｜QTS Fayet | 300 MW | 逾300MW是兩棟新建建物的容量，原文未說已全數上線，且說園區尚未完全運轉；原文也未指明出自QTS發債資料，屬NordS |
| 口徑不同 | S_sites_20261028.csv:13 | Microsoft Fairwater Wisconsin｜Fairwater  | 450 MW | 原文450MW是第一期預估需求，非已上線容量。原文無2026-06上線日；係記者未註明出處的敘述，非利害關係人說法；so |
| 口徑不同 | S_sites_20261028.csv:18 | Microsoft Fairwater Wisconsin｜全園區（WEC 預期 | 2600 MW | 2.6GW 與原文相符，但這是 WEC 投資人簡報的預期，屬利害關係人說法，不應標為[Verified]。範圍是 I-9 |
| 口徑不同 | S_sites_20261028.csv:19 | Stargate Abilene Campus 1｜全園區 | 1200 MW | 1.2GW 是全園區滿載容量，非在建量；已交付 6/8 棟、618MW。原文無 2027Q1 延期說法，完工日寫「未說明 |
| 口徑不同 | S_sites_20261028.csv:24 | Crusoe Abilene Campus 2 (Microsoft)｜全園區 | 900 MW | 900MW 是園區及自備電廠規模，原文另載兩棟各 336MW IT，計 672MW，basis 應註明。700MW／DC |
| 口徑不同 | S_sites_20261028.csv:27 | Google Mesa AI Campus (Project Redhawk)｜ | 400 MW | 400MW規劃、SRP州內風光供電均符合原文；但2019是宣布日期（Announced 2019-01），一期於2025 |
| 口徑不同 | S_sites_20261028.csv:31 | AWS Project Rainier (New Carlisle)｜全園區 | 2200 MW | 2.2GW、30棟是記者敘述，非Amazon本人引述。2025-10僅7棟上線，全園區無上線時點。Measured AI |
| 口徑不同 | S_sites_20261028.csv:47 | Meta Prometheus (New Albany)｜Epoch 預測 20 | 854 MW | 854 MW IT 確為 Epoch 對 2027 Q1 的累計預測，但 metric 寫「在建」。在建部分應約為 85 |
| 口徑不同 | S_sites_20261028.csv:56 | Stargate Frontier (Vantage Frontier Camp | 1400 MW | 1.4GW、10棟、370萬平方呎、2026下半年首棟皆屬實。但1.4GW是全園區規劃，原文僅說已動工；高層稱GPU運算 |
| 口徑不同 | S_sites_20261028.csv:61 | Project Jupiter｜全園區（微電網運轉上限） | 1600 MW | 1,600MW約65%屬實，但係YGI修訂申請自述，宜標[Interested-party]；刊登日為10/27非10/ |
| 口徑不同 | S_sites_20261028.csv:91 | IREN Childress Campus｜全園區 | 750 MW | 750MW是完整建成的規劃容量，不是在建量。2026年只交付H1-4（4×50MW IT）。來源是DCD報導，不是公司本 |
| 口徑不同 | S_sites_20261028.csv:104 | Nscale–Aker Narvik Campus (ex-Stargate N | 230 MW | 原文 230MW 指園區規模，未稱在建或初期；2027、逾3萬 Rubin、Nscale 單獨管理屬實。Aker 簡報、 |
| 口徑不同 | S_sites_20261028.csv:131 | Nexus Hubbard Campus｜Phase 1 自備燃氣電廠許可 | 612.93 MW | 612.93MW 是2025-06-18核證的發電許可量，原文稱燃氣廠尚未在建，不是在建；口徑應為發電而非未說明。202 |
| 口徑不同 | S_sites_20261028.csv:132 | Nexus Hubbard Campus｜全園區（Anthropic四筆租約＋1 | 1600 MW | 1.6GW是場內燃氣電廠的發電容量，不是租約簽約量。原文只說融資在洽談中，未提已募得；也未提WSJ。 |
| 口徑不同 | S_sites_20261028.csv:141 | xAI Colossus 1｜全園區 | 300 MW | 原文300MW以上、22萬片GPU屬實，但為簽約後當月內可用，非已上線；2024-07、150MW、247.2MW出自其 |
| 口徑不同 | S_sites_20261028.csv:149 | PORTS-Pike Technology Campus｜初期 4.25 IT- | 4250 MW | 4.25 IT-GW 是 NVIDIA 殘值擔保範圍，簽約租賃量是約 8 IT-GW，metric 寫「簽約」易被誤讀。 |
| 口徑不同 | S_sites_20261028.csv:152 | Campus AI Fouju（MGX–Mistral 巴黎區 1.4GW AI | 700 MW | 700MW是規劃中臨時變電站併網容量，非在建亦非IT；來源為保證人報告(2025-12-22)非公司本人；RTE分段數字 |
| 口徑不同 | S_sites_20261028.csv:182 | Google Armstrong County (Goodnight) Camp | 265.5 MW | 265.5MW與Docket 58881核准、30分鐘削減相符，但原文是核准負載而非在建；180+85分段與2026-0 |
| 口徑不同 | S_sites_20261028.csv:183 | Google Armstrong County (Goodnight) Camp | 260 MW | 260MW、Crusoe Two、Ensign購址、合計525.5MW與須全額削減均有據；但本件為申請方聯合提出的擬議命 |
| 口徑不同 | S_sites_20261028.csv:185 | Google Armstrong County (Goodnight) Camp | 530 MW | 530MW屬Aterio估算，非已證實在建，tag與metric偏高。原文日期為2026-07-30，未提Google、 |
| 口徑不同 | S_sites_20261028.csv:196 | Google Council Bluffs Southlands Campus｜ | 696 MW | 696 MW 屬實，但原文明示以 power capacity（電網取電，即設施口徑）排名，並非未說明。原文更新日為 2 |
| 口徑不同 | S_sites_20261028.csv:200 | Google Project Zodiac (Fort Wayne)｜Phase |  MW | 核准日9/16與五棟屬實，但原文僅為填濕地許可，未證實已動工，列為「在建」不妥；原文稱總容量未知，無MW數值 |
| 口徑不同 | S_sites_20261028.csv:207 | AWS Susquehanna Campus, Salem Township｜擴 |  MW | 3-0 通過約1,600英畝重劃屬實；15棟十年內是 Amazon 官員說法，應標 [Interested-party] |
| 口徑不同 | S_sites_20261028.csv:218 | Amazon Pecos County Campus (GW Ranch)｜發電 | 1000 MW | 1GW與2027上半年首電屬實，但原文為planned並僅取得許可，非在建；文章日期為1/27；原文未提Amazon |
| 口徑不同 | S_sites_20261028.csv:226 | Poolside Project Horizon｜第一期（原 CoreWeave | 250 MW | 250MW、15年租約、2027 Q1完工皆符；但原文僅稱power capacity，未提「gross power」， |
| 口徑不同 | S_sites_20261028.csv:229 | Crusoe Victoria AI Factory｜全園區 | 1500 MW | 1.5GW 是 Blue Energy 核電廠上限，不是 Crusoe 園區規劃 MW；原文只說 gigawatt-sc |
| 口徑不同 | S_sites_20261028.csv:237 | Stratos AI Data Center Campus｜全園區 | 7500 MW | 7.5GW是發電設施容量，不是資料中心負載。原文無2028及面積縮減說法，且稱園區約1萬英畝。來源是KSL報導。 |
| 口徑不同 | S_sites_20261028.csv:246 | Applied Digital Polaris Forge 2｜全園區 | 280 MW | 280MW為設計容量且9月才動工，8/18非在建；口徑未說明屬實。200MW IT、B. Riley 300MW、202 |
| 口徑不同 | S_sites_20261028.csv:251 | Crusoe–Tallgrass Cheyenne AI Campus（Proj | 2700 MW | 2.7GW 是記者描述，Google 申請文件未載容量；2031、四棟屬實。文中無 Tallgrass 電力中心容量、2 |
| 口徑不同 | S_sites_20261028.csv:255 | Nscale Monarch Compute Campus｜Phase 1 | 1350 MW | 1.35GW與2028年初交付屬實，但原文是設計容量，未說在建或9月動土；Anthropic、Microsoft及頁面日 |
| 口徑不同 | S_sites_20261028.csv:282 | EcoDataCenter Borlänge（Kvarnsveden 超大園區） | 250 MW | 原文指啟用時可取用 250 MW 電力，並非首座資料中心在建容量；2027 年初完工一致。原文未提 Mistral AI |
| 口徑不同 | S_sites_20261028.csv:320 | DeepSeek Ulanqab AI Campus｜全園區 | 1000 MW | 約1GW、部分自建部分租賃、2027年底至2028年初上線均屬實；但原文明說1GW為供電和設施口徑，非「未說明」。原文未 |
| 口徑不同 | S_sites_20261028.csv:324 | Huawei Cloud Gui'an Data Center｜全園區 | 400 MW | 400MW與營運狀態相符，但原文為Guiyang的Guizhou AI中心。第三方追蹤站數字不宜標[Verified]， |
| 口徑不同 | S_sites_20261028.csv:332 | CtrlS Chandanvelly AI Campus｜Phase 1 | 250 MW | 原文 250MW 是第一期已核准電力，案子只是宣布規劃，未說在建，metric 應非在建。原文無 2027 或 FY20 |
| 口徑不同 | S_sites_20261028.csv:339 | SK Group–AWS Ulsan AI Data Center｜一期 | 41 MW | 41MW、2027-11 與原文相符。但原文稱 2025 年 8 月才計畫動工，as_of 當時屬規劃中，不是在建。8/ |

判定由模型產生，只作人工複核的優先順序；「無法取得」依規則不繞過，須人工開原文。詳見同名 csv。
