# E 段：2026Q3 財報季站點／上線 GW／延遲證據（2026-11-05；查證代理產出，未經未參與代理核對）

用途：供下一版 Wattonomics 判斷是否寫入基準（In T2 M 欄已確認延遲月、或新增 T2 列）。依 README「證據」規則與「已確認延遲門檻」（工作單 20261102 E12），本清單**不直接入 In**；任何一列要入基準，須先經未參與代理開原文核對（`evidence_raw/D_review_<日期>.md`）並判「屬實」。

- 範圍：(1) AMD 2026Q3 財報與法說（2026-11-03 美東）；(2) 2026-10-28 以後 T2 已列園區的上線、延期或擴建公告；(3) IREN 季報。
- 對照基準：`wattonomics/model/20261105_Wattonomics_v0.3.xlsx` In 頁 T2 與 Delay 頁。
- 引文經擷取工具摘錄，皆不超過 15 個英文字。「二手」＝新聞或彙整頁轉述，公司原文未開。
- 等級：Verified（公司財報數字或監理文件）／Interested-party（公司或開發商自述）／Analogy（第三方估計、衛星判讀、媒體）。
- E12：公司本身揭露，或兩家以上獨立可信媒體報導且公司未否認；公司否認者只追蹤、不入基準。轉載同一篇原始報導的多家網站**不算獨立**。
- 以下網址都是本次搜尋結果中出現、且 WebFetch 成功開啟者；開不了的頁面（SEC 8-K、CNBC、Fool、Benzinga、DCD 等）未計入來源。

## 證據表

| # | 日期 | 玩家 | 站點（T2 鍵） | 內容（口徑） | 上線或延遲年月 | 來源 1（網址＋引文） | 來源 2（網址＋引文） | 等級 | 達到 E12 | 建議 |
|---|---|---|---|---|---|---|---|---|---|---|
| EQ#1 | 2026-11-03 | AMD → OpenAI | 新（站點未揭露） | OpenAI 6 GW 協議的首批 1 GW MI450（口徑未說明，晶片部署量） | 2026H2 開始（原公告）；Q3 法說未重申、未改期 | https://ir.amd.com/news-events/press-releases/detail/1260/amd-and-openai-announce-strategic-partnership-to-deploy-6-gigawatts-of-amd-gpus ："first 1 gigawatt deployment … set to begin in the second half of 2026" | https://stockanalysis.com/stocks/amd/transcripts/724028-q3-2026/ ："We have a big step-up in Q4." | Interested-party（AMD） | 不適用（沒有延遲） | 只追蹤 |
| EQ#2 | 2026-11-03 | AMD（Helios 首批客戶：Anthropic、Meta、Microsoft、OpenAI、Oracle） | 新（站點未揭露） | 首批 Helios 機架已上線；Q4 部署大增、2027 持續放量；MW 找不到 | 2026Q3 首批出貨；2026Q4 放量 | https://ir.amd.com/news-events/press-releases/detail/1303/amd-reports-third-quarter-2026-financial-results ："running models and workloads on the first" Helios | https://singjupost.com/amd-advanced-micro-devices-q3-2026-earnings-call-transcript/ ："expect Helios deployments to increase substantially in the fourth quarter" | Interested-party（AMD） | 不適用 | 只追蹤 |
| EQ#3 | 2026-11-03 | AMD（客戶未指名） | 新 | 分析師稱約 3 GW AMD 部署規劃至 2027；管理層未確認數字（口徑未說明） | 2027 | https://stockanalysis.com/stocks/amd/transcripts/724028-q3-2026/ ："roughly 3 GW or so of planned AMD deployments into 2027" | https://singjupost.com/amd-advanced-micro-devices-q3-2026-earnings-call-transcript/ ："we know where the GPUs are going" | Analogy（分析師提問）；管理層答覆 Interested-party | 不適用 | 只追蹤 |
| EQ#4 | 2026-11-03 | EuroHPC（LUMI-AI，AMD MI430X）；Vultr（HPE Helios） | 新 | LUMI-AI 2027H2 上線；Vultr 部署首批 HPE Helios；MW 皆找不到 | LUMI-AI 2027H2；Vultr 未說明 | https://ir.amd.com/news-events/press-releases/detail/1303/amd-reports-third-quarter-2026-financial-results ："Vultr will deploy the first AMD Helios by HPE systems" | https://stockanalysis.com/stocks/amd/transcripts/724028-q3-2026/ ："planned to come online in the second half of 2027" | Interested-party（AMD） | 不適用 | 新站候選（MW 找不到，暫不入） |
| EQ#5 | 2026-10-28 | Microsoft（第三方供應商，未指名） | 不能從現有資料分離（候選：cp_neo_iren、cp_neo_sines 等 Microsoft 使用的新雲列） | 第三方容量交付延遲：FY27Q1（2026-07～09）預期交付延遲，部分 Q2 供給移到 H2；MW 未揭露 | FY27Q1／Q2 → FY27H2（2027-01～06） | https://www.microsoft.com/en-us/investor/events/fy-2027/earnings-fy-2027-q1 ："some third-party capacity deliveries we expected in Q1 were delayed" | https://www.datastudios.org/post/microsoft-q1-fy27-results-azure-48-percent-33-2-billion-capex-45-2-billion ："moved into the second half of the fiscal year" | Interested-party（Microsoft 財務長，公司原文） | 是（公司自揭）；但未指站點、未給 MW | 只追蹤（無法寫入特定 T2 列 M 欄） |
| EQ#6 | 2026-10-28 | Microsoft | 不能從現有資料分離（全公司） | 本季新增 24 座資料中心；dock-to-live 時間季減約 20%；MW 未揭露 | 2026Q3（FY27Q1） | https://www.microsoft.com/en-us/investor/events/fy-2027/earnings-fy-2027-q1 ："adding 24 new data centers across five continents this quarter" | https://www.datastudios.org/post/microsoft-q1-fy27-results-azure-48-percent-33-2-billion-capex-45-2-billion ："brought 24 new data centers online across five continents" | Interested-party（Microsoft） | 不適用 | 只追蹤 |
| EQ#7 | 2026-11-04 | IREN → Microsoft | cp_neo_iren（Childress H3–4，100 MW，2026-12） | Aterio 衛星判讀：Horizon 4 尚缺備援發電機與冷卻設備，年底交付有風險；IREN 不予置評（Horizon 各 50 MW，口徑未說明） | 原 2026 年內；可能滑出 2026，月數未給 | https://www.tftc.io/satellite-images-microsoft-ai-capex-shortfall-iren-nebius ："Business Insider first reported"（Aterio 影像至 2026-10-21） | https://phemex.com/news/article/satellite-imagery-reveals-potential-delays-at-microsoft-ai-data-center-partners-iren-and-nebius-102427 ："still lacks backup generators and cooling equipment" | Analogy（Aterio 衛星判讀；兩來源皆轉述同一篇 Business Insider） | 否（實為單一原始報導；非公司揭露） | 只追蹤（Delay 頁加 cp_neo_iren 追蹤，M＝0） |
| EQ#8 | 2026-11-04 | Nebius／DataOne → Microsoft | 新（T2 無列；對應 Delay 頁 D#19） | Vineland（新澤西）約 140 MW 可能滑到 2027（口徑未說明）；DataOne 不予置評 | 原 2026；可能 2027 | https://runtimewire.com/article/satellite-images-iren-nebius-microsoft-deliveries ："roughly 140 MW at Nebius's New Jersey project could slip into 2027" | https://panews.io/articles/01a1a8a0-3eba-7357-9979-d7ce14c28ac9 ："may be delayed to 2027" | Analogy（Aterio；同一篇 Business Insider） | 否 | 只追蹤（更新 D#19：MW 由 50 首期改記約 140，仍未達 E12） |
| EQ#9 | 2026-11-05 查 | IREN | cp_neo_iren、cp_neo_iren27 | Q1 FY27（2026-09-30 季；工作單所稱「FY26Q1」）：**未公布**。發布日 2026-11-05（美東），法說 17:00 ET；本檔撰寫時（台北 11-05）尚未發布 | — | https://iren.com/investors/news ："IREN to Release Q1 FY27 Results on November 5, 2026" | https://www.stocktitan.net/news/IREN/iren-to-release-q1-fy27-results-on-november-5-5kte3kvsf6am.html （法說 5:00 p.m. ET） | Interested-party（IREN） | 不適用 | 只追蹤（季報出來後再補） |
| EQ#10 | 2026-11-02 | SpaceX／xAI | 不能從現有資料分離（可能屬 cpx_xai_c2 Colossus 2 範圍，或新站） | Southaven 兩座新設施（「Gun Range」「Spam」）預計 2026 年底運作；Whitehaven／Southaven 共規劃六座；MW 找不到 | 2026-12（地方媒體轉述） | https://www.fox13memphis.com/news/xai-expansion-brings-more-data-centers-growing-concerns-to-southaven/article_49846aba-f811-46ea-8acc-f3a8d4d931b2.html ："expected to become operational by the end of 2026" | https://www.fox13memphis.com/news/spacexai-building-fifth-mid-south-data-center/article_f7c6b339-03a3-478e-8fa8-777835f84a5b.html （2026-09-02，第五座 660,000 平方呎；同一媒體，非獨立） | Analogy（地方媒體；Musk 說法轉述＝Interested-party） | 不適用 | 只追蹤（MW 找不到；不能寫新站） |
| EQ#11 | 2026-10-28 | Oracle | 全公司（T2 Oracle 列合計） | Investor Day：未來兩年容量約增 7 GW；FY28 年底總容量逾 10 GW（口徑未說明；是否含非 AI 舊區域找不到） | FY28 年底（2028-05） | https://aiweekly.co/alerts/oracle-projects-30b-fcf-by-fy2029-caps-fundraising-at-20b ："grow by about seven gigawatts in the next couple of years"（＝E#1） | https://book.st-hakky.com/en/news/oracle-targets-30b-free-cash-flow-by-2029 ："total data center capacity to exceed 10 gigawatts"（＝E#2） | Interested-party（Oracle 財務長；二手） | 不適用 | 只追蹤（公司層級目標只作對照） |

## 每列說明

**EQ#1 AMD–OpenAI 首批 1 GW。** 原協議（2025-10-06）寫明首批 1 GW MI450 於 2026 下半年開始部署，第一批認股權證隨首批 1 GW 部署而歸屬。2026-11-03 的新聞稿與法說（兩份逐字稿：stockanalysis、singjupost；investing.com 版內容一致）都沒有提到 OpenAI 1 GW 的進度、站點或改期；財務長只以「客戶協議」泛稱認股權證會計，未說是否已歸屬。可從法說確認的只有：Q3 首次出貨 MI455X／Helios 機架、Q4 大幅成長、2027 持續放量。依 README「晶片／算力合約不算站點容量」，本列不能入 T2；OpenAI 的 AMD 機架落在哪個站點（Oracle 或其他）不能從現有資料分離。

**EQ#2 Helios 首批客戶上線。** 新聞稿與法說都點名 Anthropic、Meta、Microsoft、OpenAI、Oracle 已讓首批 Helios 機架上線。這是晶片層上線，沒有 MW 或站點，只能作 2026Q4 放量的旁證（與 Siliconomics 的 AMD 出貨對帳較相關）。

**EQ#3 約 3 GW 至 2027。** 3 GW 是分析師 Vivek Arya 提問時的說法，蘇姿丰未確認數字，只說 2027 能見度良好、知道 GPU 會進哪些資料中心。等級 Analogy，不可當作公司數字。

**EQ#4 LUMI-AI、Vultr。** 兩個具名部署都找不到 MW。LUMI-AI 屬歐洲 HPC（主權／學研），若日後找到 MW 可列「其他」或「主權」新站候選。

**EQ#5 Microsoft 第三方容量延遲。** 財務長 Amy Hood 在 FY27Q1 法說（2026-10-28）自揭：Q1 預期的第三方容量交付延遲、部分原計 Q2 的供給移到下半財年（2027-01～06），並以此解釋資本支出低於指引、曆年資本支出約 1,700 億。公司自揭，形式上已達 E12；但 Microsoft 未指名供應商、站點或 MW（stocktwits 報導亦稱未指名），所以無法對應到 T2 任何一列的 M 欄。T2 中 Microsoft 為使用者的新雲列（cp_neo_iren 100 MW 2026-12、cp_neo_sines 200 MW 2027-09）是可能對象，但只有 EQ#7 有站點層級的旁證。建議只追蹤，並在 Delay 頁「無 T2 列的候選追蹤表」新增一筆（全公司層級、公司自揭、站點未指名）。

**EQ#6 Microsoft 新增 24 座資料中心。** 只有座數與 dock-to-live 時間縮短，沒有 MW；Fairwater、Atlanta、Wisconsin 在本季法說中都沒有出現。T2 的 Fairwater 列無需變動。

**EQ#7 IREN Childress Horizon 4。** Business Insider 引 Aterio 衛星影像（至 2026-10-21）稱 Horizon 4 尚缺備援發電機與冷卻設備，年底前交付有風險；IREN 與 DataOne 不予置評，Microsoft 稱基礎設施策略保持彈性。tftc 與 phemex 兩頁都是轉述同一篇 Business Insider，不是兩家獨立報導；Business Insider 原文本次未開到。公司沒有否認，也沒有自揭，未達 E12。T2 cp_neo_iren 為 H3–4 合計 100 MW、2026-12；若延遲屬實，最可能的影響是約 50 MW（Horizon 4）由 2026-12 滑到 2027，但月數找不到。Horizon 1 已於 2026-08-13 交付（IREN 新聞頁標題）。

**EQ#8 Nebius／DataOne Vineland。** 同一篇 Business Insider／Aterio：約 140 MW 可能滑到 2027。T2 沒有此站；Delay 頁 D#19（SemiAnalysis：Nebius／DataOne 新澤西首期 50 MW 延約 6–7 個月）是同站的舊候選。本列增加了一個非 SemiAnalysis 的來源，但仍是單一原始報導、公司未表態，未達 E12。

**EQ#9 IREN 季報。** IREN 自稱 Q1 FY27（截至 2026-09-30）；發布日 2026-11-05 美東、法說 17:00 ET。撰寫本檔時（台北 11-05）尚未發布，MW 上線與時程寫「未公布」。IREN 新聞頁同時列出 Childress 750 MW、Sweetwater 2,000 MW（建設中）與 10-21 西班牙 300 MW 園區公告，這三項都不是本季財報數字，本檔不列。

**EQ#10 xAI Southaven。** FOX13（2026-11-02）稱 Southaven 兩座新設施預計 2026 年底運作，Musk 確認其中一座名為 Gun Range。沒有 MW；第二來源只找到同一家媒體 09-02 的報導（第五座設施，同樣沒有 MW）。T2 xAI 只有 cpx_xai_c2（Colossus 2 增量 585 MW-IT 2027-03）；新設施是否屬 Colossus 2 範圍，不能從現有資料分離。

**EQ#11 Oracle Investor Day。** 已收錄於 `E_oracle_20261104.md` E#1、E#2，本列只是放進財報季清單以便對照。「約 7 GW」與「逾 10 GW」都是公司層級目標，依 README 只作對照。比較：Delay 頁 Oracle 公告上線 GW（T2 有效納入、上線年 ≥ 2026）＝5.366 GW IT；另觀測值 obs_orcl＝2,543 MW-IT（FY27Q1 末）。兩者相加約 7.9 GW IT，與「FY28 年底逾 10 GW」（口徑未說明）的落差，要看 10 GW 是設施還是 IT（設施 ÷ 1.2 ≈ 8.3 GW IT），本檔不判斷。

## 找不到或不適用（2026-10-28 以後）

- **Oracle Stargate 站點**（Abilene 餘 2 棟、Shackelford、Doña Ana／Jupiter、Port Washington、Saline、Red Oak）：10-28 以後找不到站點層級的上線、延期或擴建公告。Jupiter 不可抗力通知（2026-09-24）與以卡車運天然氣（2026-10 上旬）都在期間之前，已見 E#12；SynMax 對 Shackelford 的質疑為 2026-04（Oracle、OpenAI 否認），也在期間之前。
- **Microsoft Fairwater**：FY27Q1 法說沒有提到 Fairwater 的 MW 或時程（見 EQ#6）。
- **Meta Hyperion／Prometheus**：Meta Q3 財報為 2026-10-21，在期間之前。Meta 工程部落格（2026-10-19，期間之前）稱 Prometheus 第一期「already complete」，但沒有 MW。10-28 以後找不到新公告。
- **Amazon Rainier（New Carlisle）**：Q3 財報（2026-10-29）新聞稿與逐字稿（investing.com、tradingkey、datastudios）都沒有 GW、Rainier、New Carlisle 或站點時程；Jassy 只說大客戶有時一次預訂「up to a gigawatt」。
- **Google**：Q3 法說（2026-10-28）CEO 講稿沒有 MW、GW 或新站點，只說正在加快上線容量。
- **CoreWeave**：Q3 財報訂於 2026-11-10 17:00 ET（公司 10-27 公告），**未公布**。
- **AMD–OpenAI 站點**：找不到。

## 與 T2 現值差異（小結）

| T2 鍵 | 現值（MW、上線年月、M） | 本次證據 | 差異 | 是否建議寫入基準 |
|---|---|---|---|---|
| cp_neo_iren | 100 MW、2026-12、M＝0 | EQ#7：Horizon 4（約 50 MW）年底交付有風險（Analogy，單一原始報導） | 可能約 50 MW 延到 2027，月數找不到 | 否；Delay 頁加追蹤 |
| （無 T2 列）Nebius／DataOne Vineland | —（Delay D#19：首期 50 MW） | EQ#8：約 140 MW 可能延到 2027 | 追蹤 MW 由 50 改記約 140 | 否；更新 D#19 |
| （無 T2 列）Microsoft 第三方容量 | — | EQ#5：公司自揭延遲 1～2 季，站點與 MW 未揭露 | 無法對應站點 | 否；Delay 頁加全公司層級候選 |
| cpx_xai_c2 | 585 MW、2027-03 | EQ#10：Southaven 兩座新設施 2026 年底運作，MW 找不到 | 不能分離 | 否 |
| Oracle 各列（合計 5.366 GW） | 不變 | EQ#11：公司層級 7 GW／10 GW 目標 | 只作對照 | 否 |
| 其他 T2 列（Fairwater、Prometheus、Hyperion、New Carlisle、Google、CoreWeave Helios） | 不變 | 10-28 以後找不到新公告 | 無 | 否 |

**結論**：本次 11 列中，只有 EQ#5（Microsoft 第三方容量延遲）在形式上達到 E12（公司自揭），但沒有站點與 MW，不能寫入任何 T2 列的 M 欄。**本次沒有建議寫入基準的變動**；建議下一版只在 Delay 頁新增或更新三筆追蹤（cp_neo_iren、D#19 Vineland、Microsoft 全公司第三方延遲）。待 IREN（11-05 美東）與 CoreWeave（11-10）季報公布後再補查。
