# A 段：美國雲與新雲玩家 AI 容量證據（原始證據摘要）｜2026-10-28

配套檔：`A_us_cloud.csv`（117 筆）。查證範圍是規劃書 `docs/plan/20261028_全市場GW帳_規劃.md` 第 5 節 v0.1 的玩家表。這份是原始證據，還沒有交給未參與製作的代理核對。

## 讀法與共同口徑

- **單位：** 本檔的「最佳估計」一律是 **AI 在役 IT GW（年底）**。換成 TK GW（設施）時乘 PUE 約 1.2（[Assumed]，與 repo 個股模型一致）。
- **Epoch〈AI Chip Owners〉：** 本檔大量用到這個資料集（2026-05-12 版）。
  - 它的「Power in MW」是晶片功耗加總（repo `alphabet_facts` 標為 TDP），不是 IT，也不是設施電力。
  - 換成 IT 乘 1.4–1.9，基準 1.65（repo `mw.conv.itPerTdp`，[Analogy]）。
  - 它的數字是累計值，沒有扣除退役。2026Q1 那一列資料集自己標了不完整。
  - 它按「晶片擁有者」歸屬。所以 CoreWeave 擁有、租給 Microsoft 的晶片算在 CoreWeave，不算 Microsoft。用它加總不會重複，但要另外做「營運者 × 使用者」矩陣。
- **公司口徑：** 公司自己揭露容量時，絕大多數沒有說明是 IT 還是設施（Microsoft、Amazon、Oracle、CoreWeave、Nebius connected／contracted 都是如此）。明確寫 IT 的只有這幾家：
  - Nebius active；
  - 房東端的 critical IT（Galaxy、Applied Digital、IREN、TeraWulf、Core Scientific）。
- **時效：** 截至今天（10-28），Microsoft、Alphabet 的 Q3 財報今天盤後才發布，Amazon 明天（10-29）發布，所以本檔沒有收錄。下一輪要補兩項：新增 GW 和資本支出。

## Microsoft

- **年底 AI 在役 IT GW 最佳估計：**

  | 年底 | 估計 | 區間 | 依據 |
  |---|---|---|---|
  | 2024 | 1.4 | 1.2–1.6 | Epoch 晶片 861 MW × 1.65 |
  | 2025 | 3.3 | 2.8–3.8 | Epoch 2,016 MW × 1.65 |
  | 2026E | 5.8 | 4.8–6.8 | repo 期初 2026-06 自有 4.5（3.5–5.5）＋H2 約 2 GW 設施 × AI 80% ÷ 1.2，[Derived] |

  以上都不含向 neocloud 租用的部分（2026-06 為 0.8–1.5 GW-IT，repo `nc.mwLeased.open`）。
- **口徑：**
  - 公司說法：FY25 新增 >2 GW，FY26Q2–Q4 每季約 1 GW。未說明 IT 或設施，也沒有分 AI 與非 AI。
  - Bloomberg 報導：全公司自有＋租賃約 12 GW、2032 年目標 >38 GW。這是記者引用知情人士，口徑未說明。
- **大園區：**
  - Fairwater Atlanta 636 MW-IT，已上線。
  - Fairwater Wisconsin 一期 369 MW-IT，2026-06 完工；全園區 2,263 MW-IT，預計 2028Q2。公司說單站可擴至 2 GW。
  - 以上 MW 都是 Epoch 的估計（[Analogy]）。
  - Abilene 鄰近園區：向 Crusoe 承租 900 MW，2027 年中首棟上線。但另有報導寫 700 MW，待核。
- **重複計算：**
  - 向 CoreWeave（500–900 MW-IT）、Nebius（Vineland 約 205）、IREN（200）、Nscale（約 411）、Lambda（未揭露）租用的容量，同時出現在營運者那一列。
  - 2026-10-19 有報導指出 Microsoft 退出 Nscale Loughton 的 50 MW。
- **repo 既有值：** company.json `mwYearEnd` 640／1,630／3,250 是「對外計費」AI MW，不是在役總量，不能直接用作 GW 帳的值。
- **資本支出：**
  - 曆年 2026 約 175B（記者整理）。原為 190B，差額是租賃改列營業租賃造成的，本人說計畫不變。
  - FY27Q1 >$50B。
- **已搜尋未找到：**
  - FY26Q1 新增 GW。搜尋詞：「Microsoft FY26 Q1 capacity gigawatt」，repo 亦標缺。
  - OpenAI 在 Azure 的 GW。搜尋詞：「OpenAI Azure gigawatt 250 billion」。
  - Anthropic Azure 1 GW 的一手原文（anthropic.com 頁抓取被擋）。

## Alphabet / Google

- **年底 AI 在役 IT GW 最佳估計：**

  | 年底 | 估計 | 區間 | 依據 |
  |---|---|---|---|
  | 2024 | 2.1 | 1.8–2.5 | Epoch TPU 991＋NVIDIA 307 MW |
  | 2025 | 4.9 | 4.1–5.6 | repo `mw.ai.it.epoch.2025q4` |
  | 2026E | 7.6 | 6.0–9.4 | repo 2026-06 為 6.3（5.0–7.9），再加兩季平均新增，[Derived] |

- **矛盾待解：**
  - 2025 年全公司資料中心平均負載只有 4.84 GW（設施口徑，含非 AI；42.4 TWh ÷ 8,760）。
  - 這個數字低於 AI 年底 IT 估計 4.9 GW。
  - 可能原因：
    - Epoch 累計值沒有扣退役；
    - TPU 的換算係數偏高（取 1.4 則為 4.1）；
    - 年平均低於年底值；
    - 平均負載低於裝置容量。
  - 建議 v0.1 的 Google 取區間下半段，並把 TPU 換算係數列為敏感度。
- **口徑：** 公司沒有揭露任何 MW 或 GW 的容量。只有用電量和潔淨能源合約 GW（發電容量，不是負載）。
- **大園區：**
  - Mesa, AZ 183 MW-IT（Epoch）。
  - Google–Blackstone 合資 500 MW，2027 年上線（二手來源）。
- **Anthropic 相關：**
  - Google Cloud TPU：2026 年上線「well over 1 GW」，Anthropic 自述，有第二來源 DCD。
  - Broadcom 次世代 TPU：3.5 GW，2027 年起。可能是 TPU 系統銷售，裝在第三方機房。
- **重複計算：**
  - Google 為 Fluidstack 擔保 TeraWulf、Cipher、Hut 8 等站點，那些容量可能是 Anthropic 在用。
  - TPU 系統銷售給客戶自建機房，不屬 Google 自有容量。
  - 這兩類都不能計入 Google 營運者 GW。
- **資本支出：** 2026 年 195–205B（Q2 法說，CFO）。Q3 財報今天盤後才出。
- **已搜尋未找到：** Google 自報總 GW。搜尋詞：「Google data center gigawatts capacity total 2025」。

## Amazon / AWS

- **年底 AI 在役 IT GW 最佳估計：**

  | 年底 | 估計 | 區間 | 依據 |
  |---|---|---|---|
  | 2024 | 0.8 | 0.4–1.0 | repo 對外 360 ÷ 0.8＝0.45；Epoch 616 MW × 1.65＝1.0 |
  | 2025 | 2.2 | 1.8–2.8 | repo 顆數法 1.8；Epoch 1,672 MW × 1.65＝2.8 |
  | 2026E | 3.5 | 3.0–4.5 | repo 2026-06 為 2.6，加上 Anthropic「2026 年底近 1 GW」新容量的剩餘部分，以及 OpenAI GB300 |

- **矛盾待解（NVIDIA 部分）：**
  - Epoch 估 Amazon 擁有的 NVIDIA 晶片功耗為 832 MW，換算約 1.4 GW-IT。
  - repo 只估約 0.6 GW-IT。
  - Trainium 兩邊一致：Epoch 840 MW 對 repo 1.22 GW-IT，比值約 1.45。
- **全公司（全部負載）：**
  - 2025 年新增 3.9 GW，其中 Q4 >1.2 GW。這是本人說，口徑不明；同期 Q3 新聞稿稱近 12 個月加 3.8 GW。
  - 總容量公司沒有揭露。
  - 目標：2027 年底為 2025 年底的 2 倍（本人說）。
  - Morgan Stanley 估 2026 年新增 6 GW、2027 年新增 8 GW（賣方）。依此隱含 2025 年底約 14 GW（設施）。這是 [Derived]，只作對照。
- **大園區：**
  - New Carlisle（Rainier，Anthropic）：目前 1,037 MW-IT（Epoch，2026-10-19）。公司說全園區 2.2 GW、30 棟；Epoch 預測 2028 年達 1,925 MW-IT。
  - Ridgeland 228、Madison 284 MW-IT（Epoch）。
  - Cipher Black Pearl：216 MW-IT，2026-10 起租。
- **合約：**
  - Anthropic 最多 5 GW。金額 >$100B／10 年，見 10-Q，[Verified]；GW 數字為 Anthropic 自述。
  - OpenAI 約 2 GW Trainium，2027 年起爬坡；另有 $38B NVIDIA 合約，MW 未揭露。
- **資本支出：** 2026 年約 220B（本人說，原約 200B，主因記憶體漲價）。Q3 財報 10-29 發布。
- **已搜尋未找到：** 2026H1 新增 GW。搜尋詞：「AWS added gigawatts second quarter 2026」。

## Meta

- **年底 AI 在役 IT GW 最佳估計：**

  | 年底 | 估計 | 區間 | 依據 |
  |---|---|---|---|
  | 2024 | 0.9 | 0.7–1.0 | Epoch 524 MW × 1.65 |
  | 2025 | 2.2 | 1.9–2.5 | Epoch 1,339 MW × 1.65 |
  | 2026E | 3.2 | 2.8–3.6 | Prometheus 由 0.47 增至 0.68，加上 2026 晶片新增，[Derived／Assumed] |

  以上不含 MTIA，也不含租用的 CoreWeave、Nebius、Google、Crusoe 容量。
- **交叉驗證：** 2024 年全公司資料中心平均用電 2.06 GW（設施，含非 AI；18.06 TWh，Meta 環境資料，[Derived]）。AI 0.9 GW-IT 約 1.0 GW 設施，與全公司用電量吻合。2025 年用電數字已搜尋未找到。
- **口徑：** Meta 自報的「1 GW」「5 GW」是 compute capacity，沒有說明 IT 或設施。Epoch 的值是 IT。
- **大園區：**
  - **Prometheus（New Albany, OH）：**
    - Epoch 估 471 MW-IT（2026-08），預測 2028Q3 達 1,022 MW-IT。
    - 本人說：2026 年上線（Zuckerberg，2025-07）。2026-10 已運轉。
  - **Hyperion（Richland Parish, LA）：**
    - 目前 0 MW。
    - Meta 發言人：2030 年前達 2 GW；5 GW 沒有完成日。
    - 媒體報導：租金 2029-06 起付。
    - Epoch 預測：2028-01 一期 1,632 MW-IT。
  - **El Paso：** 1 GW，2028 年起上線，BlackRock 持 80%。
- **合約：**
  - CoreWeave $14.2B＋$21B，MW 未揭露。
  - Nebius $12B（另有最多 $15B 選擇權），2027 年初起交付。
  - Crusoe 約 1.6 GW（記者報導）。
- **表外：** Hyperion 與 El Paso 都是 80/20 合資，資本支出不在 Meta 帳上。C1 對帳時要分開處理。
- **資本支出：** 2026 年 135–145B，含融資租賃本金（Q3 新聞稿，2026-10-21）。
- **已搜尋未找到：**
  - Meta 自報的總 GW 或在役 GW。搜尋詞：「Meta total data center capacity gigawatts」。
  - Q3 法說會的 GW 說法（逐字稿未取得）。

## Oracle

- **年底 AI 在役 IT GW 最佳估計（曆年）：**

  | 年底 | 估計 | 區間 | 依據 |
  |---|---|---|---|
  | 2024 | 0.35 | 0.3–0.45 | Epoch 209 MW × 1.65 |
  | 2025 | 1.1 | 1.0–1.5 | Epoch 650 MW × 1.65；repo FY26 年底推估 1.84 減 FY26H2 交付 |
  | 2026E | 3.2 | 2.8–3.6 | repo 2026-08 為 2.54，加上 FY27Q2–Q3 交付 0.6–1.0 GW ÷ 1.2 |

- **口徑：**
  - 公司說：FY26 增 1.2 GW、FY27Q1 增 850 MW，屬於「Incremental DC capacity」，未說明口徑。
  - Abilene：618 MW＝75%（本人說）。Epoch 估 1.2 GW 設施。
- **大園區：**
  - Abilene，OpenAI 使用：6／8 棟已上線。擴建已取消（2026-03 Bloomberg），改由 Microsoft 承租 900 MW。
  - 其餘四站都是 Epoch 的設施估計：Shackelford 2.0 GW、Jupiter 2.2 GW、Port Washington 1.3 GW、Saline 1.4 GW。預計 2028Q4，Shackelford 首棟 2026 年底。
  - Jupiter：2026-09-24 Bloomberg 報導 Oracle 發不可抗力通知（燃氣管線被拒）。Oracle 本人說仍照時程。
- **合約：** OpenAI 新增 4.5 GW（OpenAI 自述）。
- **資本支出：**
  - FY27（2026-06～2027-05）90–95B，淨額 ≤70B。
  - 曆年 2026 指引不存在（財年錯位）。
- **已搜尋未找到：** 全公司在役 GW。FY25 以前的累計沒有揭露（repo 已標）。

## CoreWeave

- **年底在役：**

  | 年底 | active power | 來源標記 |
  |---|---|---|
  | 2024 | >360 MW | S-1，[Verified] |
  | 2025 | >850 MW | 10-K，[Verified] |
  | 2026E | >1.85 GW | 目標，[Interested-party] |

  2026-06 為 1.5 GW。區間：2026E 1.75–2.0（Q2 加了約 500 MW；房東端 Helios Phase II 要到 2027H1）。
- **contracted：** 3.1 GW（2025 年底）→ 3.7 GW（2026-06）→ 4.2 GW（2026-08）。2030 年目標 8 GW。
- **口徑：** 公司從未定義。repo 預設為 IT（[Assumed]），依據是房東一律以 critical IT 計，且 Epoch 晶片功耗 × 1.65 的量級一致：2024 年 279 MW、2025 年 799 MW。
- **房東端（[Interested-party]，多有第二來源）：**
  - Galaxy Helios：133 MW IT 已交付，全案 526。
  - Applied Digital Polaris Forge 1：250 MW IT（2026-10），全案 400。
  - Core Scientific：約 590 MW，2026-07 全客戶計費 437 MW。
- **重複計算：**
  - 房東列（Applied Digital、Core Scientific、Galaxy）和 CoreWeave 列不能相加。
  - 使用者主要是 Microsoft（repo 估 500–900 MW）、OpenAI、Meta，在使用者端也會再出現一次。
- **資本支出：** 2026 年 35–39B。

## Nebius

- **年底在役：**

  | 年底 | 估計 | 依據 |
  |---|---|---|
  | 2024 | 未找到 | 推估 <0.05 GW，[Assumed] |
  | 2025 | 0.17 GW-IT | active，公司定義為 IT，[Interested-party] |
  | 2026E | 0.65 GW-IT（0.5–0.83） | 目標 connected 0.8–1.0 GW，口徑不明；連網後要數月才轉 active，管理層說 2027H1 才全部 active。以 ÷1.2 與 active 落後折算，[Derived] |

- **contracted：** >3.5 GW（2026-05）。年底目標 5 GW，2–3.5 年內上線。
- **站點：** Lappeenranta 310 MW、Missouri 1.2 GW、Pennsylvania 1.2 GW，都預計 2027 年。
- **Vineland：** Microsoft 專用約 205 MW-IT（以 GPU 數換算，[Derived]）。有報導指燃料電池變更審議延會。
- **重複計算：** Vineland 也計入 Microsoft 的租用容量；Meta 合約在 2027 年起交付。
- **資本支出：** 2026 年 20–25B。
- **已搜尋未找到：** 2024 年底 active MW，以及 2026Q1、Q2 季末 active MW。搜尋詞：「Nebius end of 2024 capacity MW」「Nebius active power Q2 2026」。

## 其他新雲合計

- **最佳估計（不含 CoreWeave、Nebius；只算非重複的營運者自有雲端容量）：**

  | 年底 | 估計 | 依據 |
  |---|---|---|
  | 2024 | 0.2 GW-IT（0.1–0.4） | [Assumed] |
  | 2025 | 0.5 GW-IT（0.3–0.8） | [Assumed] |
  | 2026E | 1.3 GW-IT（0.8–1.8） | IREN 0.3 目標＋Lambda、Nscale、Crusoe cloud、TeraWulf／Fluidstack 等 |

  對照：Epoch「Other」擁有者 2025Q4 晶片功耗 2.07 GW，約 3.4 GW-IT。但這個類別也包含主權、企業和學研，不能直接當成新雲合計。
- **逐家：**
  - **IREN：** Childress 4 × 50 MW IT 給 Microsoft，H3–4 於 2026Q4。累計交付目標 2026 年 0.3、2027 年 0.8 GW IT。營運 MW 未揭露。
  - **Applied Digital：** 租約 1.41 GW IT。北達科他到 2026 年底交付 300 MW，房東屬性。
  - **Core Scientific：** 租約約 1.1 GW；AMD 530 MW。
  - **Crusoe：** 簽約近 5 GW（含 Abilene，與 Oracle 列重複）。Wyoming 1.8 GW 已暫停。營運 MW 未揭露。
  - **Lambda：** Kansas City 24→100+ MW；Anthropic 約 350 MW（記者報導）。營運 MW 已搜尋未找到。
  - **Nscale：** active＋contracted 1.37 GW（IPO 申報，二手）。約 2.5 萬顆 active GPU，約 51 MW-IT（[Derived]）。Loughton 因電網延到 2030 年代，Microsoft 已退出。
  - **TeraWulf：** 102 MW IT 在線（Fluidstack，Google 擔保）。Anthropic 401 MW，2027H2 起。
  - **Together：** HUMAIN 250 MW（沙烏地）。自有 MW 未找到。
- **重複計算：** 這一層的重複最嚴重。
  - Applied Digital、Core Scientific、Galaxy 屬房東，容量歸 CoreWeave 或超大雲。
  - Crusoe Abilene 屬 Oracle／OpenAI 和 Microsoft。
  - TeraWulf、Cipher、Hut 8 屬 Fluidstack／Google／Anthropic。

## 重複計算總表（v0.1 營運者 × 使用者矩陣要先處理）

| 容量 | 營運者列 | 使用者列 | 規模（MW） |
|---|---|---|---|
| CoreWeave 租給 Microsoft／OpenAI／Meta | CoreWeave | Microsoft、OpenAI、Meta | 1,500（2026-06）全部 |
| 房東交付 CoreWeave（Applied Digital、Core Scientific、Galaxy） | 房東 | CoreWeave | 約 800 在線 |
| Oracle Stargate | Oracle（開發商 Crusoe／Vantage／STACK／Related） | OpenAI | 618 在線、約 8,100 規劃（設施） |
| Abilene 鄰近 900 MW | Crusoe | Microsoft | 2027 起 |
| Nebius、IREN、Nscale、Lambda 租給 Microsoft | 各新雲 | Microsoft | 約 800–1,500 |
| Anthropic 在 AWS（Rainier 等）、Google TPU、Azure | Amazon、Google、Microsoft | Anthropic | 約 1.5 在線；規劃 5＋5＋1 GW |
| Google 擔保的 Fluidstack 站點 | TeraWulf、Cipher、Hut 8 | Fluidstack／Anthropic | 約 1,353 規劃 |

## 存取與限制

- **被擋的頁面：** 兩頁因權限請求未回應而沒有發出請求，依規則不繞過：
  - `anthropic.com/news/microsoft-nvidia-anthropic-announce-strategic-partnerships`
  - `anthropic.com/news/anthropic-amazon-compute`
- **抓取失敗：** client_error，均已改用其他來源：
  - DCD Microsoft 38 GW 頁；
  - dutchitchannel Prometheus 頁；
  - letsdatascience Meta Q3 頁；
  - Business Wire Core Scientific 頁（內容為空）。
- **超長網址：** Tom's Hardware Hyperion 頁的網址超過長度上限，被代理拒絕。
- **SemiAnalysis：** 只用了一筆，而且只作市場對照：2027 年全球新增 50 GW。
