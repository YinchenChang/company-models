# B 段：AI 實驗室（使用者）、中國、主權、第三方總量、基建延遲｜原始證據摘要｜2026-10-28

配套檔：`B_labs_china_totals.csv`（72 筆）、`B_delay.csv`（24 筆）。範圍依規劃書 `docs/plan/20261028_全市場GW帳_規劃.md`。Microsoft、Google、Amazon、Meta、Oracle、CoreWeave、Nebius、新雲由 A 段（`A_us_cloud.*`）負責，本段只從「使用者」角度列出它們與實驗室的合約。本檔是原始證據，**還沒有交給未參與製作的代理核對**。

標記：[Verified] 正式文件／財報／政府或獨立機構數據；[Interested-party] 利害關係方（寫明是誰）；[Analogy] 第三方模型或類比；[Derived] 本段換算；[Assumed] 假設。「本人說」與「記者推論」在 csv 的 quote 欄分開寫。網頁多由抓取工具摘述，只有引號內是原文。

## 0. 先看結論

1. **實驗室「簽約 GW」比「上線 GW」大一個數量級，而且簽約數大多是上限。** OpenAI 稱已簽約 10GW（美國，2026-04），另有 AMD 6GW、Broadcom 10GW、Nvidia 5GW（Vera Rubin）、AWS 2GW、Cerebras 0.75GW、SB Energy Ohio 4.25 IT-GW 等，彼此大量重疊（晶片合約與機房合約是同一批容量的兩面）。上線量只有 Friar 的 2025 年底 1.9GW（[Interested-party]），2026 年沒有任何自報上線數。Anthropic 簽約 ≥14.8GW（The Information 彙整），自己計畫 2026 年底約 5GW（NYT 報導）。
2. **「營運者 × 使用者」矩陣必須用「晶片合約不算容量」的規則。** AMD／Broadcom／Nvidia／Cerebras（除 Cerebras 自營雲外）的 GW 是晶片功耗，落地要靠 Oracle、Microsoft、CoreWeave、SB Energy 等的機房。建議：矩陣只放「機房或雲容量合約」；晶片合約另列一欄做 C2（顆數）對帳，不進 C0 加總。
3. **xAI 現在是營運者兼出租方。** SpaceX 季報稱 2026 年中 1.4GW nameplate compute（[Interested-party]），客戶有 Anthropic（Colossus 1 全部＋Colossus 2 部分，>300MW）、Google（約 11 萬顆 GPU）、Reflection AI。Epoch 估 Colossus 2 為 946MW IT（2026-06）。xAI 自用量＝總量減出租，未揭露，屬不可分離。
4. **第三方總量口徑差很大，不能直接比。** Epoch：2025 Q4 AI 晶片推得約 30GW（設施，晶片功率 × 約 2.5）；McKinsey：2025 AI 約 31GW（推）、2030 156GW；IEA：2025 全部資料中心用電 485TWh（≈55GW 平均負載）；TrendForce：全部資料中心已裝 155GW（2026）。由下而上加總時，主比對對象應是 Epoch（AI 專用、設施口徑、可拆季）。
5. **延遲證據：設備交期（30–60 個月）與併網等待（36–84 個月）都長於建物工期（12–24 個月）。** 已排定專案的實際延遲多數是 3–12 個月（Denton 約 6、Nebius 新澤西約 6–7、Oracle 部分站約 12），極端案 24 個月（新墨西哥 Jupiter）。Sightline：2025 年 26% 容量延遲，2026 年 30–50% 恐延到 2027 以後；SemiAnalysis（只佐證）反駁分母錯誤，認為大型自建幾乎沒動。

## 1. A 區塊：實驗室作為使用者

### 1.1 OpenAI

| 項目 | 營運者 | GW／MW | 口徑 | 時程 | 標記 |
|---|---|---|---|---|---|
| 全體上線 2025 年底 | 多家 | 1.9GW | 未說明（Epoch 假設 IT） | — | Interested-party（Friar），單一來源 |
| 全體簽約（美國） | 多家 | 10GW | 未說明 | 原目標 2029 | Interested-party；DCD 判斷為簽約 |
| 2027 目標 | 多家 | low double-digit GW | 未說明 | 2027 | Interested-party（投資人備忘錄，二手） |
| Abilene | Oracle（Crusoe 建） | 618MW（8 棟中 6 棟，2026-08 底） | 未說明 | 全站 2026 Q4 | Oracle 法說；第二來源 Epoch（421MW IT 於 2026-05） |
| Oracle 雲合約 | Oracle | 4.5GW（$300B 報導） | 未說明 | 2027 起 5 年 | Interested-party；金額為報導 |
| Shackelford／Doña Ana／Michigan／Wisconsin | Oracle | 1,400／1,750／988／902 MW IT | IT（Epoch） | 2027-08 起／2028 Q4 | Analogy（Epoch） |
| Milam／Lordstown | SoftBank／SB Energy | 857／214 MW IT | IT | 首棟約 2026-10／2027-03 | Analogy（Epoch） |
| Ports Pike, Ohio | SB Energy（20 年租約） | 4.25 IT-GW（全園 8） | IT | 2028 起，首期 800MW | Interested-party（Nvidia） |
| Microsoft Azure | Microsoft | 未揭露（$250B） | — | 未揭露 | Verified（8-K） |
| Fairwater Atlanta／Wisconsin | Microsoft | 636／369 MW IT（OpenAI 占比未揭露） | IT | Wisconsin 全站 2,263MW 於 2028 Q2 | Analogy（Epoch）；不可分離 |
| CoreWeave | CoreWeave | 未揭露（$22.4B） | — | 至 2029–31 | Verified（8-K） |
| AWS | AWS | $38B（GB200/GB300，MW 未揭露）＋ Trainium 2GW | 未說明 | Trainium4 2027 起 | Interested-party |
| Nvidia Vera Rubin | 未指定 | 5GW（原 10GW LOI 被改為 $30B 股權，報導） | 未說明 | 未揭露 | Interested-party |
| AMD MI450 | 未揭露 | 6GW（首 1GW 2026 下半年） | 晶片功耗 | AMD：Q3 開始爬坡 | Interested-party |
| Broadcom 客製 | 未揭露 | 10GW（Broadcom：2027 1.3GW、2028 >5GW） | 未說明 | 至 2029 | Interested-party |
| Cerebras | Cerebras（自營雲） | 750MW（＋1.25GW 選擇權） | CS-3 系統 IT | 至 2028／2030 | Verified（S-1）＋Interested-party |
| Stargate UAE 第一期 | Khazna（G42） | 200MW | 未說明 | 2026（Q3 目標未見上線） | Interested-party |
| Stargate UK、Norway | Nscale 等 | — | — | 2026-04 暫停 | 報導 |

TK 已有的 OpenAI 2026 上線估計（`20261024_OpenAI2026GW_A段`）：平均 2.5–3.8GW、基準 2.8–3.2（[Derived]，輸入 [Assumed]），本段沒有找到能取代它的新直接證據。

### 1.2 Anthropic（repo `anthropic_src.yaml` 已收大部分，本段只整理成矩陣列）

| 營運者 | GW／MW | 口徑 | 時程 | 標記 |
|---|---|---|---|---|
| Google（TPU，至多 100 萬顆） | >1GW | 未說明 | 2026 | Interested-party |
| Google／Broadcom（下一代 TPU） | 3.5GW（8-K）；Anthropic 稱 5GW | 未說明 | 2027 起 | Verified（8-K）；差 1.5GW 未解 |
| AWS（Rainier New Carlisle） | 1,037MW IT（Epoch 2026-10）；全站 1,925MW IT 2028 | IT | — | Analogy（Epoch）；AWS 稱至多 2.2GW 用電 |
| AWS（Trn2–Trn4） | 至多 5GW；2026 年底近 1GW 新增 | 未說明 | 2026–2036 | Interested-party |
| Microsoft Azure | 至多 1GW（$30B） | 未說明 | 2026-11 至 2033-05 | Interested-party |
| SpaceX／xAI（Colossus 1＋2） | >300MW | 未說明 | 2026-05 起，90 天可退 | Verified（S-1，Reuters） |
| Fluidstack／Hut 8 | 245MW（至多 2.295GW） | 未說明 | 2027 Q2 | Interested-party |
| TeraWulf | 401MW | critical IT | 未載 | Interested-party |
| Riot、Nscale、Lambda、Volta | 191／460／350／133 MW | 未說明 | Nscale 2027 年底 | Interested-party（報導） |
| Nexus Hubbard TX | 500MW（遠期 7.7GW） | 未說明 | 2026 | Interested-party（FT） |
| AMD MI450 | 至多 2GW | 晶片 | 首 GW 2027 上半年 | Interested-party |
| CoreWeave | 未揭露 | — | 2026 下半年起 | Interested-party |

全體：簽約 ≥14.8GW（The Information）；2026 年底計畫約 5GW、2027 年底約 10GW（NYT 報導）；第三方逐站 2025 1.4GW、2026 5.0GW（Measured AI）。S-1 全文沒有 GW 或 MW（TK `20261026_Anthropic_S1`）。

### 1.3 xAI／SpaceX

- 全體 nameplate compute：0.4GW（2025 年中）→ 1GW（Q1 2026）→ 1.4GW（Q2 2026）；目標 2026 年底 >2GW、2027 年底「接近 10GW」（Musk）。DCD 註明 Musk 自己說 20GW（電力散熱層）不會達成。[Interested-party：SpaceX 剛上市]
- Colossus 2：946MW IT（Epoch，2026-06-15），預估 2027 年 1,531MW IT；約 1.3GW 表後燃氣發電、沒有電網連接（Cleanview）。
- Colossus 1：Anthropic 稱 >300MW 且包下全部；Cleanview：約 500MW 供電（198MW 自發＋300MW 電網）。
- 出租：Anthropic $1.25B／月（至 2029-05）；Google $920M／月（2026-10 至 2029-06，約 11 萬顆 GPU）；Reflection AI。

### 1.4 營運者 × 使用者關係清單（給 C0 矩陣）

列＝營運者，欄＝使用者；「—」表示未揭露 MW。**粗體**是可直接放進矩陣的機房／雲容量；其餘是晶片合約，只放 C2。

| 營運者 \ 使用者 | OpenAI | Anthropic | xAI | Google | Microsoft 自用 |
|---|---|---|---|---|---|
| Oracle | **Abilene ~0.82GW 全站（已上線 0.618）；另 4.5GW 合約（含 Shackelford、Doña Ana、Michigan、Wisconsin 約 5.0GW IT）** | — | — | — | — |
| Microsoft | **Azure $250B（—）；Fairwater 兩站共用（占比不可分離）** | **Azure 至多 1GW** | — | — | Fairwater 共用 |
| CoreWeave | **$22.4B（—）** | **多年約（—）** | — | — | （A 段） |
| AWS | **$38B（—）、Trainium 2GW** | **Rainier 1.04GW IT 上線；至多 5GW** | — | — | — |
| Google Cloud | 2025 有報導的雲合約（GW 未找到） | **TPU >1GW（2026）；Broadcom 3.5–5GW（2027 起）** | — | 自用 | — |
| SpaceX／xAI | — | **Colossus >300MW** | 自用（總量 1.4GW 減出租） | **約 11 萬顆 GPU** | — |
| SB Energy／SoftBank | **Ports Pike 4.25 IT-GW（2028 起）；Milam、Lordstown** | — | — | — | — |
| Cerebras | **750MW** | — | — | — | — |
| 新雲（Fluidstack、Nscale、Lambda、Volta、TeraWulf、Riot、Nexus） | Nscale（UK，已暫停） | **合計約 2.3GW 已揭露 MW** | — | — | Nscale 1.35GW（同園區） |
| Khazna（G42） | **Stargate UAE 200MW→1GW** | — | — | — | — |
| 晶片合約（只放 C2） | AMD 6、Broadcom 10、Nvidia 5 GW | AMD 2 GW、Broadcom TPU（同上） | — | — | — |

重複計算提醒：Rainier（AWS）、Colossus（xAI）、Fairwater（Microsoft）、Abilene（Oracle）在 A 段會以營運者身分出現，本段只能放在「使用者」欄。

## 2. B 區塊：中國

| 項目 | 數值 | 標記 |
|---|---|---|
| 全國智算 | 2,185 EFLOPS（2026-06 底，年增 177%）；2030 目標 9,800 EFLOPS | Verified（MIIT，二手轉述）；精度口徑未明，不能換成 GW |
| 全國資料中心容量 | 32GW（2025 年底）→ 40GW（2026 年底）→ >60GW（2030）；AI/HPC 占 39%（2026）→ 2026 AI 約 15.6GW [Derived] | Analogy（Rystad） |
| 同上（第二來源） | 已交付 >24GW；字節約占 1/5、幾乎全租用 | Analogy（SemiAnalysis，只佐證） |
| 字節 2026 資本支出 | 1,600 億人民幣（FT 2025-12）→ >2,000 億（SCMP 2026-05）；Bloomberg：討論至多 $70B | Interested-party（匿名人士），三個數字範圍不同 |
| 阿里 | 單季 677 億人民幣（2026 年 4–6 月）；三年 3,800 億承諾，報導考慮提高到 4,800 億 | Verified（財報）；提高為匿名報導 |
| 騰訊 | 單季 528 億人民幣（2026 Q2，+176%） | Verified（財報） |
| 百度 | 未找到 2026 單季資本支出（財報逐字稿頁面被擋） | — |
| 四大合計 | 約 $84B（至 2027，Goldman） | Analogy |
| 華為昇騰 | 2026 至多 160 萬 die（910C 約 60 萬顆）→ 約 0.95–1.3GW [Derived，用 TK 每 GW 約 63.2 萬顆] | Interested-party（匿名，Bloomberg 2025-09） |

讀法：中國 AI 容量最可用的總量是 Rystad 的「容量 × AI 占比」（約 15–16GW，2026）。但它與 SemiAnalysis 的 24GW 總量差 8GW，口徑（設施或 IT、已交付或已建）都未說明。建議 v0.1 中國列用 12–16GW（2026 年底，AI，設施）為區間，[Derived]，並在 C1 用四大資本支出交叉檢查。

## 3. C 區塊：主權與其他（只列大項）

| 玩家 | 數值 | 狀態 | 標記 |
|---|---|---|---|
| Stargate UAE（Khazna／G42） | 200MW 第一期 ⊂ 1GW ⊂ 5GW 園區（巢狀不可加） | 2026-09 尚無試運轉公告；Reuters：設計改成分散多站 | Interested-party |
| Humain（沙烏地） | 2030 1.9GW、2034 6.6GW；Riyadh、Dammam 各 100MW 首期 2026 Q2；AMD 500MW、xAI 500MW | GB300 交付需美方許可 | Interested-party |
| 歐盟 AI Gigafactories | 至多 7 座、每座 ≥10 萬顆 GPU、>€30B | 2026-11 截止、2027 初決標、2027 動工 → 2028 前不貢獻上線 | Verified（執委會，轉述） |
| Mistral／MGX（法國） | 1.4GW 園區，2028 營運；Bruyères-le-Châtel 13,800 顆 GPU 2026 Q2 | — | Interested-party |
| 印度 Reliance JV（Visakhapatnam） | 1GW、$11B、5 年 | — | Interested-party |
| 日本 | 只找到 SoftBank 堺市 150MW（舊 Sharp 廠）報導，沒有總量 | — | — |

## 4. D 區塊：第三方市場總量比較

| 來源 | 2024 | 2025 | 2026E | 之後 | 口徑 | 標記 |
|---|---|---|---|---|---|---|
| Epoch（AI 晶片推） | — | 約 30GW（Q4） | 未發布 | — | **設施**（晶片功率 × 約 2.5）；只含 AI 晶片 | Analogy |
| Epoch（算力） | — | 運行約 1,600 萬 H100e；>10GW（晶片層） | 約 3 倍成長 | — | 算力；10GW 口徑未說明 | Analogy |
| McKinsey | — | 約 31GW（由 156−125 推） | — | 2030 156GW（受限 109、加速 236 [Derived]） | AI 需求；口徑未說明 | Analogy |
| IEA | — | 全部 DC 485TWh（≈55GW 平均）；AI +50% | — | 2030 全部 950TWh、AI 465TWh（≈53GW 平均） | **用電量**，全部 DC 與 AI 分列 | Verified |
| LBNL（美國） | 176TWh（2023） | — | — | 2028 325–580TWh | 用電量，美國全部 DC | Verified |
| TrendForce | — | 約 120GW [Derived] | 155GW | — | 全部 DC 已裝容量 | Analogy |
| SemiAnalysis（只佐證） | — | 新增 15.6GW | 新增 25GW | 2027 新增 51.2GW | 全部 DC，口徑未說明 | Analogy |
| Sightline | — | 近 6GW 上線（2025）；26% 延遲 | 排定 16GW（>50MW），30–50% 恐延 | — | 大型專案，用電 | Analogy |
| BNEF（美國） | 約 25GW（Bloom Energy） | — | — | 2035 194GW（2025-12 版 106GW） | 需求 | Analogy |
| 451 Research（美國） | — | 61.8GW | 75.8GW | 2030 134.4GW | 電網需求（全部 DC） | Analogy |
| Grid Strategies（美國） | — | — | — | 2030 新增 65GW（公用事業稱 90） | 新增負載 | Analogy |
| Schneider | — | — | — | 本輪未找到 GW 預測（只轉引 Goldman：2030 較 2023 +175%） | — | — |

比對建議：(1) C0 的主比對用 Epoch 的設施口徑 AI GW（30GW，2025 Q4），因為它與 TK GW 同為設施電力、且只含 AI；(2) McKinsey 2025 約 31GW 與 Epoch 幾乎相同，可作第二來源（但 McKinsey 是 2025-04 的需求預測）；(3) IEA 與 LBNL 是用電量，換成容量要再除以負載率（[Assumed]），只作量級檢查。

## 5. E 區塊：基建延遲與 D1／D2 依據

### 5.1 瓶頸月數（細節見 `B_delay.csv`）

| 瓶頸 | 低 | 中 | 高 | 主要來源 |
|---|---|---|---|---|
| 大負載併網（北維州） | 48 | 84 | 144 | JLL（轉述）；IEA |
| 大負載併網（美國平均） | 36 | 48 | 84 | JLL（轉述） |
| 大負載併網（ERCOT） | 36 | 42 | 48 | 整理站（低品質）；ERCOT 佇列 410GW（官方） |
| 發電併網（美國中位數） | 52 | 60 | 84 | LBNL Queued Up 2026（轉述） |
| 燃氣渦輪（GE Vernova／Siemens／MHI 新單） | 24–36 | 36–42 | 48 | 廠商法說 |
| 燃氣渦輪（全產業等待，EPRI） | 36 | 60 | 84 | EPRI（轉述） |
| 升壓變壓器 | 33 | 37 | 60 | Wood Mackenzie（轉述） |
| 電力變壓器 | 24 | 30 | 60 | Wood Mackenzie |
| 高壓斷路器 | 18 | 29 | 29 | Wood Mackenzie |
| 建物工期（對照） | 12 | 18 | 24 | Sightline；Construction Owners |
| 已排定專案的實際延遲（案例） | 3 | 6 | 12（極端 24） | Denton、Nebius NJ、Oracle 部分站、Jupiter |

### 5.2 「晶片等電」發言

- Nadella（2025-11，BG2 podcast，本人說）：有晶片在庫存，缺的是 warm shells。FY26 Q4（2026-07-29）Hood：需求仍超過可用容量，未給緩解時點（記者轉述）。
- Jassy（2025-08，本人說）：電力是 AWS 最大限制，要 several quarters 解決。
- Fairwater Wisconsin：Nadella 4 月稱上線，Guardian 報導 5 月仍未上線；Microsoft 否認 Guardian 的估計。
- 量化的「已出貨未上線」存量：**已搜尋，未找到**任何公司揭露。

### 5.3 延遲情景建議（交 Project 判斷，皆 [Assumed]，依據 [Derived]）

**D1 基準延遲（建議）：** 每年「排定上線」的新 GW 中，**35%** 延後，延後量的**中位 6 個月**（區間 3–12）。
- 比例依據：Sightline 2025 年實績 26% 容量延遲＋10% 專案延後商轉日（合計約 36%）；2026 年預測 30–50%，取下緣，因為 SemiAnalysis 指出大型自建（占新增大宗）延遲很少。
- 月數依據：已排定且已施工的案例多為 3–12 個月（Denton 約 6、Nebius NJ 約 6–7、Oracle 部分站 12）。
- 等效：全年平均上線 GW 相當於整體延後約 2 個月（0.35 × 6）[Derived]。

**D2 嚴重延遲（建議）：** 在 D1 之上，**沒有已確保電力（已簽併網協議或已有表後發電）的新 GW，再延 12 個月**；初值取 2027–2028 新增 GW 的 **40%**（區間 30–50%）。
- 依據：併網等待中位 48–60 個月、重型燃氣渦輪新單 36 個月以上、升壓變壓器 37 個月，都長於 12–24 個月的建物工期，所以 2026 年才開始找電的 2028 專案無法準時；Sightline 2026 年 11GW（約 69%）仍只在公告階段；GE Vernova 2029–2030 檔期只剩約 10GW。
- 反向證據：表後發電（Colossus 2 約 1.3GW、無電網）可把併網延遲從「年」壓到「月」；GE Vernova CEO 說多數案子渦輪不是卡關項；SemiAnalysis 說北美自建預測半年只動 1%。這些支持 D2 的比例不宜高於 50%。

## 6. 已搜尋，未找到

| 項目 | 搜尋詞（節錄） |
|---|---|
| OpenAI 2026 任何時點的自報上線 GW | "OpenAI compute capacity October 2026 online gigawatts Altman Friar"；"OpenAI says compute capacity gigawatts September 2026" |
| Microsoft、AWS、CoreWeave 對 OpenAI 專屬交付的 MW | "OpenAI Microsoft Fairwater share OpenAI capacity Azure"；"OpenAI AWS $38 billion … deployed before end of 2026" |
| OpenAI 與 Google Cloud 合約的 GW | "OpenAI Google Cloud TPU capacity deal OpenAI 2026 gigawatt" |
| xAI 自用與出租的拆分 | "SpaceX S-1 xAI Colossus 2 megawatts compute capacity" |
| 百度 2026 單季資本支出 | "Baidu capex Q2 2026 AI infrastructure Kunlun"（Benzinga、247wallst 頁面被擋） |
| 中國 AI 資料中心 GW 的官方數 | "China computing power total EFLOPS 2026 MIIT … gigawatts"；"China data center total power capacity GW 2025 CAICT racks" |
| 昇騰 2026 實際出貨 | "Huawei Ascend 2026 shipments units 950"（只有 2025-09 的產量目標） |
| Schneider Electric 的 GW 預測 | "Schneider Electric AI data center power demand 2025 2030 GW forecast" |
| Epoch 2026 年全球 AI GW | "Epoch AI global AI data center power capacity GW 2025 2026 2027 projection" |
| 日本 AI DC 總量 | "Japan AI data center capacity 2026 SoftBank Sakai GW" |
| 「已出貨未上線」量化存量 | "Nadella chips sitting in inventory"；"Jassy power biggest constraint … chips waiting" |
| 延遲月數的系統性中位數（非案例） | "data center projects delayed average months 2026 analysis"（nPlan、SynMax 都沒給月數） |

## 7. 被擋或未讀到原文的網址（未繞過）

- DCD〈OpenAI to use 2GW of Trainium on AWS…〉：權限請求逾時。
- DCD〈OpenAI claims to have secured 10GW…〉、DCD〈Musk: SpaceX/xAI is actively seeking more AI compute customers…〉、DCD〈Google and Anthropic confirm 1GW…〉：網站 client_error。
- SCMP〈China targets fourfold boost in AI computing capacity…〉、fourweekmba〈SemiAnalysis puts China's datacenter fleet at over 24GW〉、cryptobriefing〈Huawei's AI chip output…〉、CNBC 2026-08-17、electricchoice、techjacksolutions、blockspace、benzinga、247wallst、mobileworldlive：client_error。
- folio3（Humain）：跨網域轉址，未追。

## 8. 限制

- 網頁多由工具摘述，不是逐字全文；Epoch 的 MW 是衛星加模型推估，Epoch 自稱誤差約 1.4 倍。
- 公司自揭 GW 多數只有同源轉述，csv 的 second_source_url 欄寫了「同源」或留空時，視為單一來源。
- 未跑 pytest、LibreOffice，未改 repo 其他檔案。
