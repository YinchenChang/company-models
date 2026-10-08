# Z1-R 研究筆記：價格、用戶與用量、算力（規格第 4 節 B、C、D）

- 擷取日：2026-10-08；檔案：`data/z1_raw/revenue_compute.yaml`（251 筆 items、19 筆 not_found）
- 分組筆數：B 價格 156、C 用戶與用量 43、D 算力 52；標記：Verified 160、Interested-party 56、Analogy 27、Derived 8
- 與 Z1-F（`financials.yaml`）分工：算力服務費的財報科目（`cogs_compute_*`、`rd_compute_*`）、算力應付與預付款、資本支出由 Z1-F 登錄，本檔不重複；本檔只登錄**媒體轉述的算力服務費合計**（`cmp_fee_total_*`），供 Z1-F 對帳。key 已逐一比對，與 financials／capital_market 無重複。

---

## 1. 來源與取得方式

| 類別 | 主要來源 | 備註 |
|---|---|---|
| 一手（HKEXnews） | 2026 中期業績公告 `2026083101539.pdf`（2026-08-31，未經審計）；2025 年度業績公告 `2026033101549.pdf`（2026-03-31）；招股章程 `2025123000017.pdf`（2025-12-30） | 直接讀 PDF，note 附頁碼與英文原文 |
| 一手（官方定價） | docs.z.ai `guides/overview/pricing.md`（國際站美元價，2026-10-08 讀取） | 直接讀官方頁 |
| 官方價的第三方轉錄 | 奇連 AI（轉錄 bigmodel.cn/pricing，資料截至 2026-08-09，09-28 補 GLM-5.3）；CodePick、17173（Coding Plan 人民幣價）；Digital Applied、UsagePricing、aipricing.guru（Coding Plan 美元價） | bigmodel.cn 定價頁需 JavaScript，工具只讀到空殼；z.ai 訂閱頁讀取需授權逾時。轉錄價以至少兩源交叉核對（GLM-5.2 的 8／28 元與華為雲上限一致；Pro 149 元月付與 402 元季付兩源一致），故標 Verified 並在 note 寫「第三方轉錄」 |
| 媒體 | 財聯社、同花順（網易科技、證券日報轉載）、虎嗅、36氪、華爾街見聞、每經、鈦媒體、IT之家、AIbase | 公司數字的轉述標 Interested-party |
| 第三方量測（Analogy） | 上海有色網 SMM 算力日報（卡時、整機月租）；AtomGit／CSDN 租價彙整；OpenRouter 週榜（21 世紀、老虎）；QuestMobile（每經） | 一律給區間 |
| 晶片效能 | Tom's Hardware、The Register、Unite.AI、GIGAZINE、華為 CloudMatrix-Infer 論文（arXiv 2506.12708）、Lenovo Press、DatabaseMart | SemiAnalysis 來源另行標示 |

## 2. 口徑說明（重要）

### 2.1 「API 平均售價上升約 101%」
- 原文（中期公告第 14 頁）：「the average API selling price increased by approximately 101%」，基期「from the beginning of 2026」。
- 口徑判斷：**實現平均售價**（開放平台及 API 收入 ÷ 計費 token），不是牌價；與 token 量「較年初 40 倍以上」同一句，期間為 2026 年初至報告期（或公告日）。
- 與其他漲價數字的關係（全部登錄，不互相取代）：
  - 年度業績公告（2026-03）：「83% increase in API call pricing compared to the end of last year」→ 2025 年底至 2026-03 的定價漲幅。
  - 虎嗅：2026-02～04「連續三次上調API價格，累計漲幅超過83%」（與年報 83% 疑為同一數）。
  - AIbase：2026-02-12 海外 API「提升了 67%-100%」；牌價自證：GLM-4.7 $0.6／$2.2 → GLM-5 $1.0／$3.2（+67%／+45%）；人民幣 GLM-4.7 2／8 元 → GLM-5 4／18 元（+100%／+125%）。
  - 2026-04 GLM-5.1：再漲約 10%（騰訊新聞、SCMP），較 GLM-5-Turbo 上調 8%～17%。
  - 101% ＞ 83% 的差額，可能來自 1H26 內組合轉向較貴的新型號（GLM-5.1／5.2 的 1M 全量價 8／28 元），以及免費 Flash 流量比重變化；公司未拆分。

### 2.2 用戶數
- 「740 萬」是**公告日（2026-08-31）**的「enterprise and developer users」，不是 6 月底；6 月底為「超過 580 萬」（第 3 頁）。模型若取 1H26 期末，應用 580 萬。
- 「較年初 +144%」以 740 萬計，反推年初約 303 萬；年度公告稱 2026-03 時「超過 400 萬 registered users」。三點可排成時間序列：約 303 萬（2026 初，Derived）→ 400 萬（2026-03）→ 580 萬（06-30）→ 740 萬（08-31）。
- 付費日活只有增幅（+603%），無絕對數；Coding Plan 付費開發者 24.2 萬是 2026-03 的數字（年度公告）。

### 2.3 token 量
- 唯一的公司揭露絕對值：招股章程「2025-11 日均 4.2 兆 token」。鈦媒體轉述 PHIP 稱「2025 上半年 4.6 萬億」，兩者矛盾（期間不同、或口徑一為月均一為期間平均），全部登錄（`tok_daily_nov25`、`tok_daily_1h25_b`）。
- 2026 年只有倍數：MaaS 較年初 40 倍以上、Coding Plan 23 倍以上、前十大付費客戶日均 98 倍；基期都是「2026 年初」，**不是** 2025-11。若要推 2026 絕對量，需假設年初日均值（Assumed，建議以 4.2T 為上限附近的區間），由 Z2 決定。
- Ox-Alpha（GLM-5.3-Flash 上市前匿名測試）6 日累計 62 兆 token，含 OpenRouter／OpenCode 免費流量，不可視為付費量。

### 2.4 算力
- 「10 萬級國產晶片」：原文「a cluster of over 100,000 domestic chips」，用於**推論**；未揭露廠牌、型號、自有或租用。媒體報導的 Day-0 適配平台（昇騰、寒武紀、海光、崑崙芯、壁仞、摩爾線程）只代表軟體相容，不等於採購對象。
- 「單位 token 推論成本較年初降 80%」「同硬體端到端效能提升 3 倍」：皆為公司自述，基期為年初或初始部署，未說明是否含硬體成本變化。
- 「算力乘數」：定義為每 1 元算力投入（含訓練與推論）產生的開放平台及 API 收入，同比約 14 倍；公司未揭露絕對值。
- 1GW 國產算力資料中心（2026-07-21 財聞）：**電力口徑未明**（IT 或設施），也未說是自有或合作方持有，標 Interested-party，basis 寫「GW 口徑未明」。
- 實體清單：招股章程記載 2025-01-16 公司與 9 家子公司列入；公司稱往績期間未依賴 EAR 物項。

### 2.5 卡時租價（Analogy）
- SMM 的報價多以「unit／月」表示，未定義 unit。H20 在同一篇同時給「卡時 7.47 元」與「整機月租 43,000 元」，7.47 × 8 卡 × 730 小時 ≈ 43,600，故推定 SMM 的整機＝8 卡；H800、A800、H100 據此折算卡時價（Derived：H800 13.36、A800 6.16、H100 13.70、H20 7.36 元／卡／時）。
- 昇騰 910B 系列 unit 定義不明：若為 8 卡整機，16,000 元／月 ≈ 2.74 元／卡／時；AtomGit 列 910B 八卡 25,000 元／月 ≈ 4.28 元／卡／時。若 unit 是單卡則為 21.9 元／卡／時，高於 H800，不合理，故採 8 卡推定，區間 2.74–4.28。
- 阿里雲 H800 單卡時租 18.75 元、圖靈小鎮 11.83 元（AtomGit 2026-05），與 SMM 折算 13.36 元同一量級；雲平台牌價高於批發長約，屬正常。
- **910C 租價找不到**（見 not_found）。

### 2.6 晶片效能與功耗
- 910C 推論約 H100 的 60%（GIGAZINE 引 AGI Hunt／DeepSeek 測試；aiwiki 另引 Reuters 同數；分析師另估 BF16 吞吐約 80%）→ 區間 0.6–0.8。
- 910C 單卡功耗 310W（Unite.AI）；aiwiki 列 310–350W（引 SemiAnalysis）→ 區間 310–350W，晶片 TDP 口徑。
- **CloudMatrix 384 系統功耗 559kW／約 600kW 及「為 GB200 NVL72 的 3.9 倍」**：Tom's Hardware、The Register、少數派、aiwiki 全部追溯到 SemiAnalysis，找不到華為官方或其他獨立量測。依共同規則「SemiAnalysis 不可單獨引用」，這兩筆已登錄但 note 標明「暫不入模型」，並列入 not_found。
- 華為論文（Interested-party）：DeepSeek-R1 在 CloudMatrix384 每 NPU 解碼 1,943 token/s、prefill 6,688；每 TFLOPS 效率 prefill 4.45 vs H800 3.96，decode 約領先 H800 10%。
- H20 對 H100：規格 BF16 148 vs 989.5 TFLOPS（H100 稠密值由 Lenovo 1,979 含稀疏 ÷ 2 推得）＝ 0.15；頻寬 H20 4.0 TB/s 反而高於 H100 SXM 3.35 TB/s，所以 decode 推論比遠高於 0.15。DatabaseMart 工程估計 Llama2-70B 吞吐比 0.26–0.42。SemiAnalysis 稱「H20 在 LLM 推論比 H100 快 20% 以上」（Asia Times 轉述）與前者矛盾且為單一源，已登錄並標明不可單獨使用。規格 D6 所述「H20 算力約 H100 的 15%」與本次查得的 0.15 一致。

## 3. 矛盾值處理（全部登錄，key 加 _b、_c）

| 項目 | 主值 | 矛盾值 | 處理建議 |
|---|---|---|---|
| 日均 token | 4.2T（2025-11，招股章程） | 4.6T（1H25，鈦媒體轉述） | 以招股章程為準；期間不同 |
| 2024 算力服務費 | 15.52 億（華爾街見聞） | 15.53 億（每經） | 四捨五入差，以 Z1-F 財報科目為準 |
| 海外 Coding Plan 漲幅 | 30–60%（AIbase，2026-02） | 80–150%（虎嗅） | 虎嗅疑為含 2026-07 再調的累計 |
| GLM-4.5 發布日 | 2025-07-28（Z.ai blog，經 Wikipedia） | 2025-07-25（彙整站） | 用 07-28 |
| GLM-5 發布日 | 2026-02-12（上證報） | 2026-02-11（Reuters／彙整站） | 時區差；用 02-12 |
| GLM-5.2 發布日 | 2026-06-16 晚（36氪） | 2026-06-17（華為雲） | 用 06-16 |
| 清言會員價 | App 內購：連續包月 ¥79、包季 ¥219、包年 ¥859、SVIP ¥229 | App 描述：連續包月 ¥19、包季 ¥79、包年 ¥299 | 描述文字疑為舊版；以內購清單為現行 |
| OpenRouter 份額 | GLM-5.2 5.5%（2.58T ÷ 46.7T，Derived） | 「75% 市占」（GIGAZINE，分母不明） | 用 5.5%，75% 不建議使用 |
| H20／H100 推論比 | 0.26–0.42（DatabaseMart 估計） | 「快 20% 以上」（SemiAnalysis 單一源） | 用 0.26–0.42；SemiAnalysis 不得單獨引用 |

## 4. GLM 各模型對應 Tokenomics 層級（Sol／Luna）的依據（規格 D24）

Tokenomics 的層級是以**物理代表架構**定義（Arch 頁）：Luna＝DeepSeek V4-Flash 類（總參數 284B、啟用 13B）、Sol＝DeepSeek V4-Pro 類（1.6T、啟用 49B）、Astra＝封閉前沿代理（4T、啟用 180B）。對應 GLM 時用三個角度，三者一致才定案：

| GLM 型號 | 對應層級 | 依據 |
|---|---|---|
| GLM-5.3、GLM-5.2、GLM-5.1、GLM-5、GLM-5-Turbo | **Sol** | ①規模：GLM-5.3 總參數 744B、啟用 40B（DataLearner），啟用參數接近 Sol 代表的 49B，遠高於 Luna 的 13B。②能力：Artificial Analysis 指數 v4.3.2 GLM-5.3＝45，介於 GPT-6 Sol 48 與 GPT-6 Luna 37 之間、較接近 Sol（Tokenomics S56）。③價格：國際站 $1.0–1.4／$3.2–4.4，屬中國廠商的中層價帶，遠低於 Astra 級封閉前沿。 |
| GLM-4.7、GLM-4.6、GLM-4.5（含 -X 高速版） | **Sol**（當代旗艦） | 2025 年的主力旗艦；GLM-4.5 為 355B 總參數、32B 啟用的 MoE（DataLearner 型號頁命名 `glm-4_5_moe-355b-a32b`），啟用參數接近 Sol 代表。-X／-AirX 是同一模型的高速版，只是價格較高，層級不變。 |
| GLM-4.5-Air（含 -AirX）、GLM-4.7-FlashX、GLM-5.3-FlashX、GLM-5.3-Flash、GLM-4.7-Flash／4.5-Flash（免費） | **Luna** | 價格比旗艦低一個數量級（Air $0.2／$1.1、5.3-Flash $0.15／$0.50、4.7-FlashX $0.07／$0.40；人民幣 Air 0.8／2–8 元、5.3-Flash 0.8／2.8 元）；產品定位為輕量或高速版。**注意**：GLM-5.3-Flash 的 AA 指數 42（v4.3）高於 GPT-6 Luna 37，以能力論接近 Luna 與 Sol 的中點（42.5），但規模未揭露、價格是 Luna 級，故仍歸 Luna；建議 Z2 在報告「已套用的預設」列出此點。 |
| 無 | Astra | 無 GLM 型號的價格、規模或能力指數達到 Astra 級（Astra 代表啟用 180B；AA 指數 53）。 |

- 免費型號（GLM-4.7-Flash、GLM-4.5-Flash）歸 Luna，但**單價為 0**：會吃算力、不產生營收，建議 Z2 在 Demand 頁把免費 token 單列（不是只用加權平均單價吸收）。
- 參考 OpenRouter 份額（Analogy）：GLM-5.2 單一型號在 2026-06-29～07-05 週占全平台 token 約 5.5%（2.58T／46.7T），為 Sol 級；Luna 級 GLM 在 OpenRouter 的份額本次未查得。

## 5. 對模型的使用提醒

1. 牌價（B 組）是定價上限；實現單價應以「開放平台及 API 收入 ÷ token」校準（D14），101% 是實現售價口徑。
2. 公司自述的倍數（40 倍、23 倍、98 倍、80%、3 倍、14 倍）全部是 Interested-party，只作對照列，不反推參數（共同規則第 4 節）。
3. 1GW 資料中心電力口徑未明；若要用，須以 Inputs 的 PUE 換算並標 Assumed。
4. 卡時租價為批發報價，不含智譜可能取得的國資或合作方優惠；CloudMatrix 384 功耗暫不入模型。
5. 年末 ARR 指引 30 億美元屬公司目標，只入反向模式對照（D22）。
