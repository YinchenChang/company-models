# 園區容量獨立核對（I_review_sites）

- 核對日：2026-10-28；核對者：獨立代理（未參與 M.csv／提案製作）
- 方法：只開 M.csv 列出的網址，或 WebSearch 回傳的網址；M.csv 的 `verdict`／`note` 一律不採信，判定只依原文。
- 技術限制：M.csv 網址若沒有出現在搜尋結果中，WebFetch 會拒絕抓取（PROVENANCE）。這類網址記為「無法取得」，改找同一事實的其他來源。
- 口徑代號：IT＝關鍵 IT 負載；設施＝設施／電網受電／電廠；未說明＝來源沒講。
- 「說話者」分三類：公司（利害關係人）、記者（含轉述）、第三方估算（Epoch 等）。

---

## A. 提案新增（11 列）

| 項目鍵 | 主張 | 來源所載（含引文） | 來源日期 | 口徑 | 說話者 | 判定 | 建議值 | 網址 |
|---|---|---|---|---|---|---|---|---|
| cp_amzn_rdg_b614 | 556 IT @2028-09（840−284） | Epoch：2027-02-09 預測「Building 5 is operational」284 MW；2028-09-29 預測「Buildings 6-14 are operational (estimate)」840 MW。假設 6–14 號樓 2026-07 動工、工期約 2.25 年。可能使用者「Anthropic (speculative)」 | 頁面更新 2026-07-22 | IT | 第三方估算（Epoch 預測） | 確認 | 556 IT @2028-09，維持。另註：228→284 的 5 號樓增量 56 MW（2027-02）目前不在 T2 | https://epoch.ai/data/ai-data-centers/directory/amazon-ridgeland |
| cp_meta_hyp_5g | 3,000 未說明 @2032-12（5000−2000） | AI Business：「expected that 2 GW will be available by 2030, with 5 GW reached by 2032」，這段沒有標出說話者。ODSC 轉述 CNBC：Meta「has not provided a completion date for the full 5GW buildout」。CNBC 原文抓不到 | 2026-07-13／14 | 未說明（「compute」） | 公司（Meta）；2032 年出處不明 | 部分確認 | MW 確認（5 GW−2 GW）。2032 年只有一篇沒標說話者的報導支持，另一篇說 Meta 沒給完成日。建議標 [Assumed] 或「時點未定」 | https://aibusiness.com/generative-ai/cost-to-build-meta-louisiana-ai-supercluster-hits-50-billion ；https://opendatascience.com/meta-expands-louisiana-ai-data-center-investment-beyond-50-billion/ |
| cp_neo_bp_p1 | 352 IT @2027-09 | Hut 8 新聞稿：「352 MW IT Lease」「high-investment-grade tenant」，沒寫時程。TNW：「Initial data-hall delivery is expected in the third quarter of 2027」 | 2026-05-06 | IT | 公司；時程經記者轉述 | 部分確認 | 352 IT 確認。2027Q3 只是「首個資料廳」交付，不是 352 MW 全數到位。建議改成首批（量未揭露）@2027-09，其餘延後；或整筆移到 2027Q4 之後並標 [Assumed] | https://hut8.com/news-insights/press-releases/hut-8-commercializes-first-phase-of-1-gw-beacon-point-ai-data-center-campus-with-15-year-352-mw ；https://thenextweb.com/news/hut-8-beacon-point-9-8-billion-ai-lease-texas |
| cp_neo_barber | 207 IT @2027-03（168＋39） | remio（轉述 DCD）：「207 megawatts of critical IT load across two phases」；「halls are expected to arrive between the fourth quarter of 2026 and the first quarter of 2027」。300 MW 為 gross。Cipher 2026-02 簡報：「300 MW Gross Capacity」 | 2026-10-10 | IT（gross 300） | 公司（經記者轉述） | 確認 | 207 IT @2027-03（最後一批資料廳），維持 | https://www.remio.ai/post/cipher-barber-lake-lease-reaches-20-years-but-construction-risk-moves-forward ；https://baxtel.com/news/cipher-digital-extends-barber-lake-lease-to-20-years-revenue-tops-9-billion |
| cp_neo_pf2 | 200 IT @2027-12 | DCD：「200MW of critical IT load」「phased across two buildings」；2026 年開始上線，2027 年達到 200 MW 全量；15 年約 50 億美元；租戶為「unnamed US-based investment-grade hyperscaler」 | 2025-10（「this week」） | IT | 公司（經記者轉述） | 確認 | 200 IT @2027-12，維持 | https://datacenterdynamics.com/en/news/applied-digital-secures-5bn-15-year-lease-with-hyperscaler-at-north-dakota-data-center-campus/ |
| cp_neo_monarch_p1 | 1,350 未說明 @2028-03 | Nscale：「Phase 1 is engineered to host 1.35GW of AI computing capacity, with delivery targeted by early 2028.」 | 頁面無日期（2026-10 抓取） | 未說明 | 公司 | 確認（僅公司目標） | 1,350 @2028-03，可維持。注意：已簽約部分未驗證（Anthropic 460 MW 的來源抓不到），建議標 [Interested-party] | https://www.nscale.com/ai-infrastructure/monarch |
| cp_sov_ajman_p1 | 30 IT @2026-12 | Blackridge 頁面無法取得。其他來源：MEED「expected capacity of 100MW」「completed within 15 months」（2024-10）；S&P「100-MW datacenter…operational by 2026」。都沒有一期 30 MW 的數字（w.media 的 30 MW 是 Mafraq 站） | 2024-10-18；2025-10-27 | 未說明 | 記者／第三方 | 無法核對 | 30 MW 與 2026-12 都沒有可取得的來源。建議暫不納入 | https://www.blackridgeresearch.com/…（無法取得）；https://www.meed.com/construction-begins-on-khazna-data-centre-in-ajman |
| cp_nbis_oxagon | 50 未說明 @2027-12 | Baxtel：「preliminary framework agreement with Nebius for up to 50MW at Oxagon」；「expected to become operational in phases from the fourth quarter of 2027」 | 2026-10-22 | 未說明 | 公司（NEOM 公告，經媒體轉述） | 部分確認 | MW 是「最多 50」，而且只是初步框架協議、不具約束力。2027Q4 是分期「開始」，不是全數到位。建議 50 @2028（[Assumed]），並註明「初步框架」 | https://baxtel.com/news/nebius-signs-preliminary-framework-for-50mw-at-neom-s-oxagon-campus |
| cp_oai_tomakomai | 50 未說明 @2027-03 | DCD：「launch with an initial 50MW in 2026」，可擴充到「more than 300MW」。SoftBank：「受電容量が300MW超まで拡大する見込み」「2026年度の開業を目指しています」（2025-04-15 動工）。日經原文無法取得 | 2025-04 | 受電容量（偏設施） | 公司（經記者轉述） | 確認 | 50 @2027-03（日本 2026 會計年度底），維持。口徑建議改標「設施（受電）」 | https://www.datacenterdynamics.com/en/news/softbank-breaks-ground-on-data-center-campus-in-hokkaido-japan/ ；https://www.softbank.jp/sbnews/entry/20250421_01 |
| cp_amzn_ulsan1 | 41 未說明 @2027-11 | DCD 轉述朝鮮經濟：「41MW of data center capacity built by November 2027」。Korea Times：SK 會長說「initially launched as a 100-megawatt facility」，預定明年（2027）下半年營運 | 2025-06；2026-09-13 | 未說明 | 記者（轉述業界／SK） | 確認 | 41 @2027-11，維持 | https://www.datacenterdynamics.com/en/news/sk-group-and-aws-to-build-ai-data-center-in-ulsan-south-korea/ ；https://www.koreatimes.co.kr/business/tech-science/20260913/sk-chairman-says-ulsan-ai-data-center-expansion-nearing-900-megawatts |
| cp_amzn_ulsan2 | 62 未說明 @2029-02（103−41） | DCD：「103MW of capacity at completion」；「the entire facility available for use by February 2029」（轉述 Chosun Biz） | 2025-08／09 | 未說明 | 記者 | 確認 | 62 @2029-02，維持（[Derived]） | https://www.datacenterdynamics.com/en/news/sk-group-and-aws-break-ground-for-data-center-in-ulsan-korea/ |

---

## B. 候選證據與既有列的差異（21 列）

| 項目鍵 | 主張（既有 vs 候選） | 來源所載（含引文） | 來源日期 | 口徑 | 說話者 | 判定 | 建議值 | 網址 |
|---|---|---|---|---|---|---|---|---|
| cp_msft_abl | 既有 700 未說明 @2027-06 整筆；候選 2×336＝672 IT，首棟 2027 年中 | Crusoe：「Each of the two new buildings is designed to deliver 336 MW of critical IT load」；「first building expected to be energized in mid-2027」；「The 900MW on-site power plant」。DCD 轉述 Bloomberg：Microsoft 承租「around 700MW of capacity」，口徑未說明 | 2026-03-27（Bloomberg 03-24） | 672 是 IT；900 是電廠；700 未說明 | 公司（Crusoe）；700 為記者引述知情人士 | 支持候選 | 拆成 336 IT @2027-06 與 336 IT（第二棟時點未揭露，[Assumed] 例如 2027-12） | https://www.crusoe.ai/resources/newsroom/crusoe-announces-new-900-mw-ai-factory-campus-in-abilene-texas-to-support-microsoft-ai-infrastructure ；https://datacenterdynamics.com/en/news/microsoft-agrees-to-lease-700mw-at-crusoes-data-center-in-abilene-texas-report/ |
| cp_orcl_shk1／shk2 | 既有 2,000 設施（200@2027-08＋1,800@2028-11）；候選 Vantage 1.4GW、首棟 2026 下半年 | Epoch 目錄頁：「1,400 MW of IT power」，B1 2027-08-15 140 MW，B10 2028-11-11 1,400 MW。Epoch 部落格：Shackelford「2 GW projected」是總設施電力，「Vantage reports 1.4 GW IT」。Vantage：「1.4GW of GPU compute capacity」，首棟「scheduled for delivery in the second half of 2026」 | Epoch 2026-07-31／04-17；Vantage 2025-08-19 | 2,000 設施 ≈ 1,400 IT | 第三方估算（Epoch）；公司（Vantage） | 支持既有 | 維持 200 設施 @2027-08＋1,800 @2028-11（等同 140＋1,260 IT）。Vantage 說 2026 下半年是開發商目標，衛星資料不支持；Oracle 說 2027 上半年的那篇（stargate.how）無法取得 | https://epoch.ai/data/ai-data-centers/directory/openai-stargate-shackelford ；https://epoch.ai/blog/openai-stargate-where-the-us-sites-stand ；https://www.businesswire.com/news/home/20250819333311/en/… |
| cp_oai_pp1–pp3 | 既有 4,250 IT 均分到 2028／29／30 年 12 月；候選 2028 首批 800 MW，2028–2032 分期 | OpenAI：「The first 800 megawatts are expected to become available in 2028」；「six-year buildout through 2032」；「initial 4.25 IT-GW」。S-1/A：「Buildings 1-9 | 4,248」Last RFS 2031；「Buildings 10-17 | 3,776」Last RFS 2032 | OpenAI 2026-08-17；S-1/A 2026-09-21 | IT | 公司（OpenAI、SB Energy 申報） | 兩者都不精確 | 800 IT @2028-12；其餘 3,448 IT 分到 2029–2031（例如每年約 1,149，[Assumed]）。既有列最後一段在 2030，比 S-1/A 的 2031 早一年 | https://openai.com/index/openai-joins-ports-pike-project ；https://www.sec.gov/Archives/edgar/data/0002133037/000162828026062846/sbenergy-sx1a2.htm |
| cp_sov_hum2 | 既有 1,700 @2030-12；候選 Al Sa'ad 一期 1GW 2027 交付 | Middle East AI News：「first phase, covering two of six 1GW plots, will deliver 1GW of capacity by or in 2027」。EnterpriseAM 引述 Humain 執行長 Tareq Amin：「targeting 1.7 GW of data center capacity in 2027」「roughly 3 GW by 2030 and 6 GW by 2034」 | 2026-09-04；2026-09-30／10-01 | 未說明 | 公司（Humain 執行長） | 兩者都不精確 | 執行長說全國 2027 年達 1.7 GW。若 hum1 是 200～250，hum2 約 1,450～1,500 @2027-12（[Interested-party]）；另有 2030 年約 3 GW 的目標可另列 | https://www.middleeastainews.com/p/humain-seeks-25-billion-for-data ；https://enterpriseam.com/ksa/2026/10/01/humain-scales-up-2027-data-center-target-as-it-aims-for-a-self-funding-buildout/ ；https://enterpriseam.com/ai/2026/09/30/humain-plans-a-1-7-gw-saudi-datacenter-build-next-year/ |
| cp_meta_pro1 | 既有 471 IT 至 2026-08；候選 562 IT（2026-09-02，Epoch 轉述） | Epoch 原始頁：「Current IT power 471MW」，時間軸 2026-08-31 為 471 MW。頁面上沒有 562 | 頁面更新 2026-10-02 | IT | 第三方估算（Epoch） | 支持既有 | 471 @2026-08，維持。562 只出現在二手轉述（brightlio 無法取得） | https://epoch.ai/data/ai-data-centers/directory/meta-prometheus |
| cp_meta_pro2／pro3 | 既有 676 @2026-11、868 @2027-07；候選 854 IT @2027Q1 | Epoch：2026-11-30 預測 676 MW；2027-01-15 740；2027-04-15 804；2027-07-14 868。頁面上沒有 854 | 2026-10-02 | IT | 第三方估算（Epoch） | 支持既有 | 維持（205 @2026-11、192 @2027-07） | 同上 |
| cp_meta_pro1–4 | 既有合計 1,022 至 2028-09；候選 Meta 稱 1GW 叢集 2026 年上線 | AP（Spectrum News）：「a 1-gigawatt cluster spanning across multiple data center buildings」「anticipated to come online this year」，是記者敘述，文中沒有 Meta 直接引語。Epoch：2028-07-01 預測 1,022 MW | 2026-01-09；2026-10-02 | 未說明 vs IT | 記者轉述公司 vs 第三方估算 | 支持既有 | 維持。小修：pro4 月份依 Epoch 可由 2028-09 改為 2028-07 | https://spectrumnews1.com/oh/columbus/news/2026/01/09/meta-lines-up-massive-supply-of-nuclear-power-to-energize-ai-data-centers ；https://epoch.ai/data/ai-data-centers/directory/meta-prometheus |
| cp_orcl_sal | 既有 1,400 設施 @2028-12 整筆；候選 MPSC 1,383、Oracle 2H27、Related 2027Q1 | MPSC：「1,383- megawatt (MW) data center in Saline Township」，客戶為 Oracle 子公司 Green Chile Ventures。Epoch：「Buildings 1-3 operational」2028-12-31 預測；988 MW IT；「reports 1,383 MW as the site capacity」 | MPSC 2025-12-18；Epoch 2026-05-22 | 1,383 為站點合約負載（設施） | 主管機關；第三方估算 | 支持既有 | 1,383 設施 @2028-12（數值微調）。Oracle 2H27、Related 2027Q1 的說法無法取得原文 | https://www.michigan.gov/mpsc/commission/news-releases/2025/12/18/mpsc-approves-dte-electric-energy-contracts-for-data-center ；https://epoch.ai/data/data-centers/satellite-explorer/OpenAIStargateMichigan |
| cp_orcl_jup | 既有已知延遲 12 月；候選 Oracle 稱進度不變、Phase 1 2028Q3 | Oracle（X）：「Project Jupiter remains on our planned schedule」。天然氣管線完工日由 2026-08-15 延到「February 1, 2027」。Bloomberg 報導 Oracle 發出不可抗力通知，若園區沒有「come online in 2028 as planned」就延後付款 | 2026-08-14；2026-09-25 | — | 公司；記者引述知情人士 | 兩者都不精確 | 沒有來源支持「12 個月」。官方立場是沒有延遲，但管線晚 6 個月，而且有不可抗力通知。建議延遲欄改為 0（官方），在風險欄註明 6 個月管線延後與不可抗力通知；12 個月只當情境用。americanbuildout 那篇無法取得 | https://stocktwits.com/news-articles/markets/equity/orcl-stock-drops-after-6-month-delay-to-gas-pipeline-project-supporting-new-mexico-ai-data-center/cZotDceRJKT ；https://analyticsindiamag.com/ai-news/oracle-moves-to-shield-itself-from-delays-at-new-mexico-ai-data-centre-report |
| cp_msft_atl | 既有 636 IT 全放 2025-11；候選 321 @2025-10＋315 @2026-06 | Epoch：2025-10-14「Building 1 and 2 are operational」321 MW；2026-06-14「Building 3 and 4 are operational」636 MW。Microsoft 2025-11-12 宣布上線，沒有給 MW | Epoch 2026-07-17；SiliconANGLE 2025-11-12 | IT | 第三方估算；公司 | 支持候選 | 321 IT @2025-10（或依公司宣布日用 2025-11）＋315 IT @2026-06 | https://epoch.ai/data/ai-data-centers/directory/microsoft-fairwater-atlanta ；https://siliconangle.com/2025/11/12/microsoft-brings-atlanta-based-ai-super-factory-online/ |
| cp_msft_wi1 | 既有 369 IT @2026-06；候選 2026-04-16 已上線 | Epoch：2026-04-16「Building 1 is operational」369 MW，引用 Microsoft 宣布。另預測 2027-02-20 B2 達 737 MW、2028-05-13 達 2,263 MW | 2026-07-25 | IT | 第三方估算（引用公司） | 支持候選 | 369 IT @2026-04。附帶：wi2 的 1,894 可拆成 368 @2027-02＋1,526 @2028-05 | https://epoch.ai/data/ai-data-centers/directory/microsoft-fairwater-wisconsin |
| cp_sov_hum1 | 既有 200 @2026-06；候選至 2026-09 沒有公開的運作確認 | DCD：「initial capacity of 100MW each」「expected to go live in Q2 2026」（DCD 自述）。EnterpriseAM（2026-10-01）：250 MW 的建案「still due online by the end of this year」 | 2026-03；2026-10-01 | 未說明 | 記者；公司（執行長） | 兩者都不精確 | 沒有找到 Riyadh／Dammam 已運作的證據。執行長最新說法是年底前上線 250 MW。建議 200（或 250）@2026-12 | https://www.datacenterdynamics.com/en/news/humain-secures-211-plots-of-land-in-saudi-arabia-for-data-centers/ ；https://enterpriseam.com/ksa/2026/10/01/humain-scales-up-2027-data-center-target-as-it-aims-for-a-self-funding-buildout/ |
| cpx_sov_uae | 既有 200 @2026-09；候選 2026Q4 | Epoch：2026-12-01 預測「Building 1 and 2 are operational」200 MW（IT）；2028-01-01 達 1,000 MW。G42 發言人：「first 200 megawatts…completed by 2026」。EnterpriseAM 那篇（Khazna 稱 Q4）無法取得 | Epoch 2026-07-22；Gulf News 2025-10-16 | IT（Epoch） | 第三方估算；公司 | 支持候選 | 200 @2026-12 | https://epoch.ai/data/ai-data-centers/directory/openai-stargate-uae ；https://gulfnews.com/business/markets/stargate-uae-first-phase-of-200mw-to-complete-by-2026-1.500310364 |
| cpx_xai_c2 | 既有 585 @2027-12，註「1,531 無出處」；候選 1,531 IT @2027Q1 | Epoch 原始頁：「CURRENT IT POWER 946MW」（2026-06-15）；「PROJECTED IT POWER 1,531MW」，時間軸 2027-02-25 | 2026-07-30 | IT | 第三方估算（Epoch） | 支持候選 | 1,531 有一手出處。增量 585 IT @2027-03 | https://epoch.ai/data/ai-data-centers/directory/colossus-2 |
| cpx_sov_mis | 既有 1,400 @2028-12；候選 2028 初只有 Mistral 96（可擴到 200），Phase 2 為 2030–2033 | L'Essentiel de l'Éco：「il accueillera 96 mégawatts sécurisés par Mistral」「montée en charge contractuelle jusqu'à 200 mégawatts」「ouverture en 2028」。官方時程：第一期「Début 2028」，含前 3 棟資料中心。全站 1.4 GW 沒有時點；Phase 2 的 2030–2033（dcmag）無法取得 | 2026-07-02；時程文件 2026-07 | 未說明 | 記者轉述公司；開發方時程 | 支持候選 | 96 @2028-03 納入；其餘 1,304 不入 | https://lessentieldeleco.fr/7371-mistral-a-20-milliards-champion-europeen-ou-mirage/ ；https://www.briedesrivieresetchateaux.fr/wp-content/uploads/2026/07/Calendrier-du-projet-Campus-IA.pdf |
| cpx_nbis_lpr | 既有 310 @2027-12；候選 310 為完工容量，2027 只有首期 | Nebius：「with capacity of up to 310 MW」「The first capacity…expected to be available to customers in 2027」，沒有完工日 | 2026-03-31 | 未說明 | 公司 | 支持候選 | 2027 只放首期（量未揭露），310 全量時點未定 | https://nebius.com/newsroom/nebius-to-construct-310-mw-ai-factory-in-finland |
| cpx_nbis_mo | 既有 1,200 @2027-12；候選首棟 200MW，時程 2026H2–2029+ | ARCO：「488,000-square-foot facility with 200-megawatt computing capacity」。Nebius：「Q2 2026 to 2029+: Site Construction & Capacity Online」「five to seven years」「up to 1.2 GW」 | ARCO 2026-05-18；Nebius 頁無日期 | 未說明 | 承包商；公司 | 支持候選 | 首棟 200（時點未揭露，最早 2027）；其餘 1,000 在 2029 年以後 | https://www.arconational.com/arco-national-construction-celebrates-groundbreaking-of-nebius-ai-factory-campus-in-independence-missouri/ ；https://nebius.com/independencemo |
| cpx_crwv_lan | 既有 300 @2027-12；候選 300 為擴充潛力、沒有時程，初期 100MW | Turner：「initial 100 megawatt (MW) data center」「with potential to expand to 300 MW」。WITF：第二期「could be online by mid- to late-2027, according to the developer」，沒有寫 MW | 2025-07-21；2025-10-19 | 未說明 | 公司；開發商（經記者轉述） | 支持候選 | 100 為首期；300 沒有時程。第二期 2027 年上線是開發商說法，但量未揭露 | https://www.turnerconstruction.com/insights/coreweave-announces-multi-billion-dollar-commitment-to-ai-infrastructure-in-pennsylvania ；https://www.witf.org/2025/10/19/data-centers-second-phase-expected-to-cost-up-to-2-5-billion-in-lancaster-county/ |
| cpx_oai_milam | 既有 857 IT @2026-10；候選 1、2 號樓 800MW（2027），全站 1.2GW | S-1/A：「Milam County 1 | 308」「Milam County 2 | 445」，兩棟 Last RFS 都是 2028（合計 753 MW-IT，不是 800）。SB Energy：「1.2 GW data center site」，初期設施「expected to enter service starting in 2026」。Epoch 部落格：首棟「delivered by October」 | S-1/A 2026-09-21；SB Energy 2026-01-09 | IT（S-1/A） | 公司（申報文件） | 兩者都不精確 | 753 IT。308 @2026Q4–2027（[Assumed]）＋445 @2028（最後 RFS） | https://www.sec.gov/Archives/edgar/data/0002133037/000162828026062846/sbenergy-sx1a2.htm ；https://sbenergy.com/openai-and-softbank-group-partner-with-sb-energy/ |
| cpx_neo_nexus | 既有 500 @2026-12；候選 TDLR 工期到 2027-10，首棟 2027 啟用 | TDLR「Nexus - Apex DFW HB1」：「Estimated completion date: 10/29/2027」。w.media：首棟「scheduled for completion in October 2027」，「up to 600 MW」。DCD 轉述 FT：「500MW of capacity later this year」 | TDLR；2026-08-20；FT 約 2026-03 | 未說明 | 主管機關登記；記者 | 支持候選 | 首棟 @2027-10／11；MW 未揭露（500 是 FT 轉述的舊目標） | https://www.tdlr.texas.gov/TABS/Projects/TABS2026022673 ；https://w.media/?p=42704 |
| cpx_amzn_cip | 既有 216 IT @2026-10；候選約 300 gross，2026-08 首批提前交付，全站 2026Q4 | Cipher 2026-02 簡報（SEC 8-K）：「300 MW Gross Capacity」「2026 Oct Lease Start Date」。Zacks：「delivered initial Black Pearl HPC capacity two months early, with rent starting in August」「remaining…deadlines remain unchanged」。216 IT 沒有在來源中找到 | 2026-02-24；2026-08-07 | gross（設施） | 公司 | 支持候選 | 300 設施（gross）；首批 2026-08，其餘依原合約（2026-10 起）。216 IT 只能標 [Derived] | https://www.sec.gov/Archives/edgar/data/1819989/000181998926000007/cipherdigital_businessup.htm ；https://www.zacks.com/stock/news/2971431/cipher-digital-ifr-q2-earnings-call-highlights-early-hpc-delivery |

（B 段共 21 列：cp_meta_pro2／pro3 算 1 列、cp_orcl_shk1／shk2 算 1 列、cp_oai_pp1–pp3 算 1 列，其餘各 1 列。）

---

## C. 建議升級的既有 納入=0 列（3 列）

| 項目鍵 | 主張 | 來源所載（含引文） | 來源日期 | 口徑 | 說話者 | 判定 | 建議值 | 網址 |
|---|---|---|---|---|---|---|---|---|
| cpx_sov_uae | 200 未說明 @2026-12，納入 1 | Epoch：2026-12-01 預測「Building 1 and 2 are operational」200 MW。G42：「first 200 megawatts…completed by 2026」 | 2026-07-22；2025-10-16 | Epoch 標 IT | 第三方估算；公司 | 確認 | 200 @2026-12，納入。口徑可依 Epoch 改標 IT | https://epoch.ai/data/ai-data-centers/directory/openai-stargate-uae |
| cpx_xai_c2 | 585 IT @2027-03，納入 1 | Epoch：946 MW（2026-06-15）→ 1,531 MW（2027-02-25 預測）；「enough power infrastructure installed for a total of 1.5 GW of IT Power」 | 2026-07-30 | IT | 第三方估算 | 確認 | 585 IT @2027-03，納入（標預測） | https://epoch.ai/data/ai-data-centers/directory/colossus-2 |
| cpx_sov_mis | 96 未說明 @2028-03，納入 1（其餘 1,304 不入） | 「96 mégawatts sécurisés par Mistral」「ouverture en 2028」；第一期「Début 2028」 | 2026-07-02；2026-07 | 未說明 | 記者轉述公司；開發方 | 確認 | 96 @2028-03，納入 | https://lessentieldeleco.fr/7371-mistral-a-20-milliards-champion-europeen-ou-mirage/ |

---

## 判定統計

- A（11 列）：確認 7（rdg_b614、barber、pf2、monarch_p1、tomakomai、ulsan1、ulsan2）；部分確認 3（hyp_5g、bp_p1、oxagon）；無法核對 1（ajman_p1）
- B（21 列）：支持既有 5（meta_pro1、pro2／3、pro1–4、orcl_shk、orcl_sal）；支持候選 11（msft_atl、msft_wi1、msft_abl、sov_uae、xai_c2、sov_mis、nbis_lpr、nbis_mo、crwv_lan、neo_nexus、amzn_cip）；兩者都不精確 5（orcl_jup、oai_pp1–3、sov_hum1、sov_hum2、oai_milam）
- C（3 列）：確認 3

---

## 無法取得的網址

PROVENANCE 一欄指網址不在搜尋結果裡、WebFetch 拒絕抓取；client_error 指網站本身回錯。依規定都沒有改走快取、鏡像或封存站。

| 網址 | 情形 | 替代來源 |
|---|---|---|
| https://brightlio.com/largest-ai-data-centers-in-the-world/ | PROVENANCE | Epoch 原始頁 |
| https://datacentercensus.com/dc/meta-prometheus-new-albany-oh | PROVENANCE | Epoch 原始頁 |
| https://www.cbsnews.com/news/meta-nuclear-power-deals-ai-data-centers/ | PROVENANCE | AP（Spectrum News）同文 |
| https://www.cnbc.com/2026/07/13/meta-louisiana-data-center-investment-reaches-50-billion-amid-ai-push.html | PROVENANCE | AI Business、ODSC（兩者說法不一致） |
| https://stargate.how/ | PROVENANCE | Epoch 目錄頁、部落格 |
| https://www.michigan.gov/mpsc/commission/news-releases/2026/08/27/mpsc-approves-siting-settlement-on-ingham-county-solar-project | PROVENANCE | MPSC 2025-12-18 新聞稿 |
| https://americanbuildout.com/project-jupiter/ | PROVENANCE | Stocktwits、AIM |
| https://www.prnewswire.com/news-releases/hut-8-commercializes-first-phase-…302763484.html | 未試，以同文替代 | hut8.com 原稿、TNW |
| https://www.sec.gov/Archives/edgar/data/1964789/000110465926084862/tm2620835d1_ex99-1.htm | PROVENANCE | — |
| https://www.datacenterdynamics.com/en/news/cipher-extends-barber-lake-data-center-lease-commitments-to-20-years/ | client_error | remio、Baxtel |
| https://www.sec.gov/…/ciphermining_businessupd.htm；https://investors.cipherdigital.com/…additional-56-mw… | PROVENANCE | Cipher 2026-02 簡報 |
| https://www.sec.gov/Archives/edgar/data/0001819989/000181998926000028/cifr-20260331.htm；…/q226_earningsxprxdraftxvf.htm | PROVENANCE | Cipher 2026-02 簡報、Zacks |
| https://capital.com/en-gb/news/cipher-digital-expands-barber-lake-lease-term-to-20-years | client_error | remio |
| https://blockspace.media/short/cipher-digitals-black-pearl-site-is-earning-revenue-ahead-of-schedule-q2-earnings/ | client_error | Zacks |
| https://ir.applieddigital.com/news-events/press-releases/detail/132/…；https://www.sec.gov/…/apld_invxfinalpresentati.htm | PROVENANCE | DCD |
| https://cryptobriefing.com/nscale-103-billion-data-center-contracts-ipo/；https://w.media/nscale-shops-second-anthropic-deal-at-monarch-data-center-campus-report/ | PROVENANCE | — |
| https://www.blackridgeresearch.com/project-profiles/ajman-data-center-project-khazna-… | PROVENANCE（搜尋也找不到） | MEED、S&P（都沒有 30 MW） |
| https://www.datacenterdynamics.com/en/news/nebius-signs-on-for-50mw-at-neoms-oxagon-data-center-in-saudi-arabia/ | client_error | Baxtel |
| https://www.neom.com/en-us/newsroom/HUMAIN-and-DataVolt-to-develop-AI-data-center-in-NEOM | PROVENANCE | — |
| https://www.nikkei.com/article/DGXZQOUC08AKQ0Y5A400C2000000/ | PROVENANCE（付費牆） | DCD、SoftBank |
| https://enterpriseam.com/ksa/issues/no-blinking-in-the-buildout/ | PROVENANCE（搜尋也找不到） | Epoch、Gulf News |
| https://presse.bpifrance.fr/mistral-securise-jusqua-200-mw-…/?lang=fra | PROVENANCE | L'Essentiel de l'Éco |
| https://dcmag.fr/le-prefet-de-seine-et-marne-a-valide-le-projet-campus-ia-de-fouju-14-gw-et-50-milliards-e/ | PROVENANCE | 開發方時程 PDF（只有第一期） |
| https://www.insidedeeptech.com/xai-colossus-2-…；https://www.ventureatlas.org/company/xai | PROVENANCE | Epoch 原始頁 |
| https://marketwise.com/investing/sb-energy-ipo-openai-warrants-ai-bubble/；https://www.zabalist.com/projects/TABS2026017746 | PROVENANCE | S-1/A、Epoch 部落格 |
| https://everydatacenter.com/nexus-hubbard | PROVENANCE | TDLR、w.media |
| https://www.arconational.com/nebius-independence-mo/ | PROVENANCE | ARCO 動工新聞 |
| https://www.hochtief.com/…/coreweave-announces-multi-billion-dollar-commitment-… | PROVENANCE | Turner 同文 |
| https://www.cnbc.com/2026/01/09/openai-and-softbank-group-announce-1-billion-investment-in-sb-energy-.html | PROVENANCE | SB Energy 新聞稿 |
| https://letsdatascience.com/…；https://www.mobileworldlive.com/ai-cloud/oracle-seeks-protection-against-stargate-delays/ | client_error | AIM |
| https://www.tomshardware.com/…meta-expands-colossal-hyperion… | 代理拒絕（網址過長） | AI Business、ODSC |
