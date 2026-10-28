# T2 園區管線第二輪合併提案（以非美國為主）摘要

- 來源：M2.csv 共 345 列證據，涵蓋 126 個園區；其中 248 列有 MW，逐列放進 proposal.csv。另 97 列沒有 MW，無法構成增量，不列入。
- 去重比對：existing_T2.csv（60 列）與 M1_round1.csv（第一輪證據）。
- 本檔只是提案，沒有改動 /home/claude/company-models。
- 依專案規則 5，交付前須由未參與製作的代理獨立核對數字，這一步**尚未做**。

## 1. 筆數

| 項目 | 列數 |
|---|---|
| 提案列 | 248 |
| 納入=1 | 6 |
| 納入=0 | 242 |
| 新 | 240 |
| 重疊:cp_sov_hum2 | 6 |
| 重疊:cp_sov_hum1 | 1 |
| 重疊:R1（Bosquel Business Park，第一輪證據，未入模型） | 1 |
| 補充 | 0 |

營運者分布：oth 180、cn 20、sov 15、neo 13、nbis 4、crwv 4、amzn 4、googl 3、meta 3、msft 1、oai 1。

## 2. 納入=1 的 MW，依年份與營運者

| 鍵 | 園區／分期 | op | MW | 口徑 | 上線 | 已確保電力 | 證據 |
|---|---|---|---|---|---|---|---|
| cp_neo_gsdcsyd1 | GreenSquareDC SYD1 Stage 1（Sharon AI） | neo | 15 | 未說明 | 2026-10（已上線） | 1 | T#226 小誤（未觸及 MW 與日期） |
| cp_nbis_bshemesh1 | Mega Or Beit Shemesh 一期（Nebius） | nbis | 58 | 未說明 | 2027-03 | 0 | T#18 小誤（只觸及 quote） |
| cp_crwv_dv3 | DataVita DV3（CoreWeave） | crwv | 40 | 未說明 | 2027-09 | 0 | T#39 屬實 |
| cp_googl_kron1 | Google Kronstorf 一期 | googl | 150 | 未說明 | 2027-12 | 0 | T#83 小誤（只觸及延遲說明與 quote） |
| cp_crwv_taloja1 | AdaniConneX Taloja 三棟（CoreWeave） | crwv | 240 | 未說明 | 2028-12 | 0 | T#161 小誤（只觸及 tag、metric、source_title） |
| cp_meta_sturg1 | Meta Sturgeon County 第一期 | meta | 970 | 未說明 | 2029-12 | 0 | T#293 小誤（只觸及 tag） |

| 年 | neo | nbis | crwv | googl | meta | 合計 |
|---|---|---|---|---|---|---|
| 2026 | 15 | – | – | – | – | 15 |
| 2027 | – | 58 | 40 | 150 | – | 248 |
| 2028 | – | – | 240 | – | – | 240 |
| 2029 | – | – | – | – | 970 | 970 |
| **合計** | 15 | 58 | 280 | 150 | 970 | **1,473** |

日期換算說明：
- DV3：原文寫「2027 夏季」，取第三季末，即 9 月。
- Kronstorf：只有年份，取 12 月。
- Taloja：原文寫「約 2028 年中」，但 expected_online 只填 2028，依年份規則保守取 12 月。
- Sturgeon：原文為 2028–2029 區間，取後端 2029-12。
- Beit Shemesh：2026Q3 起分批交付，至 2027Q1 完成，取完成月 2027-03。

## 3. 納入=0，依原因分類

一列可能同時沒過好幾個條件，下表只按主要原因歸類。MW 加總含全園區與分期的重複計算，只能看量級。

| 類別 | 列數 | MW 加總（含重複） | 說明 |
|---|---|---|---|
| B 其他業者（oth） | 165 | 83,602 | 託管、企業、電信、開發商，沒有具名超大規模或實驗室租戶；SoftBank 自用歸在這類 |
| A 中國（cn） | 20 | 9,405 | 一律只列不入 |
| C 非站點容量 | 16 | 10,379 | 電廠、BESS、PPA、電網連接合約、晶片叢集、供電總能量（條件 5） |
| D 核對未過 | 19 | 8,726 | API 判為口徑不同、不符、原文找不到或無法取得（條件 1） |
| H 重疊既有 | 7 | 1,550 | Humain 各站併入 cp_sov_hum1／hum2 的公司層級路徑（條件 3） |
| E 小誤觸及 MW、口徑或日期 | 5 | 880 | 例：T#151 的 2030 是 GPU 擴充時點、T#170 的 Batam 2027Q1 不在原文、T#85、T#155、T#63 |
| G 無日期或在窗口外 | 4 | 1,130 | 例：T#16 Khazna Ankara、T#106 AWS Zaragoza 850、T#159 Yotta NM2、T#158 OpenAI–HyperVault |
| F 累計或願景 | 1 | 2,000 | T#160 Yotta 2 GW 上限 |
| I 保守判斷 | 5 | 760 | T#40 DV1（增量無法確定）、T#61／62 西班牙 Gigafactory（歐盟候選案）、T#198 True IDC（Microsoft 承租未證實）、T#304 BUZZ HPC（320 為園區 utility capacity，2027H2 只是開始上線） |

## 4. 與既有列的矛盾

| 既有鍵 | 既有內容 | 候選證據 | T# | API 判定 | 判讀 |
|---|---|---|---|---|---|
| cp_sov_hum1 | 200 MW @2026-12（Riyadh、Dammam 各 100） | MIS 承建 Humain Riyadh：工單 1 對應 50 MW，2026-10-01 才收到；工單 2–4 各 16 個月工期（約到 2028-02）；第三方稱 Riyadh 首期原訂 2026Q2，到 9/27 仍無營運確認 | T#5、T#6、T#7、T#9 | 小誤、無法取得、小誤、口徑不同 | 如果 MIS 案就是 hum1 的 Riyadh 那 100 MW，2026-12 上線有風險，可能延到 2027–2028。MIS 案與 hum1 的對應關係，來源都沒有說明 |
| cp_sov_hum1 | 200 MW @2026-12 | Humain 執行長稱 2026 年底前上線 250 MW（公司層級） | T#9 | 口徑不同 | 既有說明已記「約 250、保留 200」，不算新矛盾，只是再次確認 |
| cp_sov_hum2 | 1,500 MW @2027-12（公司層級到 1.7 GW） | AWS–Humain AI Zone 首期最多 50 MW「2028 前」；xAI–Humain 500 MW 沒有時程；Humain–AirTrunk 250 MW 只到招標、未決標；MIS 案 250 MW | T#10、T#11、T#14、T#3、T#4、T#8 | 無法取得、口徑不同、小誤、小誤、口徑不同、小誤 | 單站證據合計遠低於 2027 年 1.5 GW 增量，而且多數沒有日期。xAI 500 MW 是否落在 1.9 GW 路徑之內，來源沒有交代 |
| （第一輪證據，非模型列）Bosquel Business Park | 第一輪記為 Fluidstack 退出後接手 | SoftBank–Sesterce Bosquel 1 GW；2031 完工是省長說法，原文找不到 | T#29 | 原文找不到 | 歸 oth、只列；與第一輪同一站 |

除上表外，其餘 59 個既有鍵（DataVolt Oxagon、SK–AWS 蔚山、Stargate UAE、Mistral／MGX、Khazna Ajman、Nscale Monarch 等）在本輪 M2 都沒有對應證據，沒有矛盾。T#69 的 Nscale–Aker Korgen 是挪威 Korgen 站，和第一輪的 Narvik 以及 cp_neo_monarch1 不是同一站。

## 5. 待決事項

1. **Sturgeon 970 MW**（本輪 inc=1 的 66%）：這個數字是第一期的電網取電量，屬設施級用電，不是 IT；只有省部長單方說法；上線是 2028–2029 區間。請 Andy 決定三件事：要不要入，取 2029-12 還是 2028-12，要不要拆成兩段。
2. **Taloja 240 MW**：原文寫約 2028 年中，提案依 expected_online 年份規則取 2028-12。如果改用原文月份，就是 2028-06。
3. **Mega Or 58 MW**：交付分批（2026Q3 至 2027Q1），提案取完成月。也可拆成 2026-12 與 2027-03 兩段。
4. **Kronstorf 150 MW**：數字是電力需求上限，偏設施口徑，而且不是專為 AI 的園區；口徑保留未說明。
5. **True IDC EEC 250 MW**（T#198 屬實）：如果能從 DCD 原文證實 Microsoft 承租，可改列 msft，取「最高 250」或首期，2027-12。
6. **BUZZ HPC 320 MW**（T#304 屬實，neo）：等公司公布分期量。
7. **Firmus Batam 360 MW**（T#170）：需要找到 2027Q1 上線的原始出處。
8. **DataVita DV1**（T#40）：需要確認擴建前的既有 MW 和承租方，才能算出增量（候選 16 MW @2026-12）。
9. **西班牙 AI Gigafactory**（T#61、T#62）：歐盟預計 2027 初決標，選定後再評估。
10. **GreenSquareDC SYD1**：比原訂 2026Q3 晚約 3 週，不到 1 個月，已知延遲月留空。
11. 本輪 345 列只有 37 列屬實、145 列小誤；中東、亞洲多數主權案或超大規模案停在「無法取得」或「原文找不到」（例：SDAIA Hexagon 480、PPC–AWS 300、IREN Bundey 800、Compute Nordic–Cerebras）。建議補抓原文後重跑 API 核對。
12. 依規則 5，納入=1 的 6 列在交付前要由獨立代理核對數字。
