# OpenAI 模型｜P×Q 與「實際 m×L ÷ 所需 m×L」比值序列｜交接（2026-10-22，由 Tokenomics 串交給 OpenAI 串）

來源串：https://claude.ai/code/session_01KzzCGL9iYEEGheAnXYLyVy（Tokenomics v18.0、token 帳閉合 A 段、本規格）。Tokenomics 權威交接：Project `claude/Tokenomics_m研究_交接_20261017.md` 第 0h、0i、0j 節。

## 1. 要做什麼（一句話）
在 company-models 的 OpenAI 模型新增一頁：用 Tokenomics 的每 GW 參數，加上 OpenAI 自己的機隊規模與營收，算出 2025、2026 的「實際 m×L ÷ 回本所需 m×L」比值，看趨勢。

## 2. Andy 的規格（原話，權威；2026-10-21～22）
| 原話 | 參數化 |
|---|---|
| 「m 要定義成『付費服務 token 的有效牌價（1 − 折扣 20%）變現率』」「m 單純只看收費業務」「跟營收有關，跟成本無關」「20% 是預設折扣，不要寫死」 | m＝實收單價 ÷ 有效牌價；有效牌價＝表列 ×（1 − d），d＝`IF_Disc` |
| 「P 其實就是等於 M 去乘上 L，再乘上一減掉預設折扣」「L 其實就是牌價」 | P＝表列牌價 ×（1 − d）× L × m。**L 讀法待 Andy 確認**：預設＝壽命期價格係數（Andy 10-18 定義），表列牌價另為一因子 |
| 「Q 就是真正拿去用作服務的算力，去乘上每單位算力可以產生的 token 數」 | Q＝平均上線 GW × 服務占比 ×（1 − 免費占比）× Σ（世代占比 × 每 GW 付費產能）× 實際利用率 × 生產折減 |
| 「看的並不是回本的絕對數字，而是看實際除以所需這個比例的趨勢」 | 主輸出＝比值的年度序列 |
| 「總算力支出是從持有的機隊成本去推算出來的……2026 年……用 3 G 瓦的這個機隊……所需要的持有成本」 | 年度算力成本＝平均上線 GW × Σ（世代占比 × 每 GW 年持有成本） |
| 「計算上一定要先從 tokenomics 開始……真的沒有，再從外面找，放進去之後……依據那個規則去把它擺進 tokenomics」 | 每 GW 參數只取 Tokenomics Interface；外部數字先登錄（OpenAI 模型自己的 SRC／Evidence），只作對照 |
| 「三 G 瓦其實應該是擺在 OpenAI 他們自己裡面……只是去取用 Tokenomics 裡面一 G 瓦的相關的訊息」 | 公司層級（GW、組合、營收、比值）放 OpenAI 模型；Tokenomics 只露出每 GW |
| 「二零二五年……兩百零四億，很明顯是高估了」 | 推論＋訓練支出 204 億（SRC_DEM_004＋006）不作算力成本，只作對照 |
| 「當然是用 Tokenomics……隨便亂算一下的」（撤回 1.0 GW） | 2025 平均 GW 用 Tokenomics 0.6→1.9 內插：指數 1.13、線性 1.25 並列 |

## 3. 從 Tokenomics 取的名稱（v18.0 已有＋v19.0 新增；v19.0 在 PR 中，合併後才可連）
- 已有：`IF_HoldEcon`（每 GW 年持有成本，C:Q＝5 世代 × 3 成本情境；基準欄 Hopper 第 2、GB200 第 5、GB300 第 8、VR200 第 11 欄）、`IF_TokGW_Luna／_Sol／_Astra`（每 GW 年產出，100% 利用）、`IF_PriceRef_*`（層級有效牌價）、`IF_Util`、`IF_ServeShare`（0.41）、`IF_FreeShare`（0.46，支出口徑）。
- v19.0 新增：`IF_Disc`、`IF_MixPaid_Luna／_Sol／_Astra`（0.40／0.45／0.15）、`IF_MixFree_*`、`IF_UtilActual`（0.41；0.30–0.65）、`IF_ProdDerate`（0.85）、`IF_FleetProdRatio_VR200`、`IF_PxEffMix`（2026 有效牌價 1.961）、`IF_PxEffMix25`（2025 有效牌價 1.359）、`IF_LTotalDisc`、`IF_LTotalFleetVDisc`、`IF_TokGWPaidMix`（付費組合每 GW 年產能，C:Q）。
- 不得連 `B4_*`、`AL_*`、`L1_*`、`CTL_*`（Tokenomics README 規定）。

## 4. OpenAI 模型自己要登錄的輸入（公司層級）
| 輸入 | 值 | 標記／出處 |
|---|---|---|
| 2025 平均上線 GW | 1.13（指數）／1.25（線性） | Derived；端點 0.6（2024 底）、1.9（2025 底），OpenAI 官方（Interested-party），Tokenomics SRC_DEM_008／009 |
| 2025 世代組合（GW） | Hopper 0.6／GB200 0.4 | Assumed（Tokenomics Alloc_In 同值；建議在 OpenAI 模型自有一份） |
| 2026 平均上線 GW | 2.5／3.0／3.5 情境 | Assumed（官方未公布；Andy「約 3 GW」） |
| 2026 世代組合 | Hopper 0.25／GB200 0.45／GB300 0.25／VR200 0.05 | Assumed（Tokenomics Cap_In D211:F214 同值） |
| 2025 營收 | 130.7 億 | Verified（SRC_DEM_007） |
| 2026 營收 | 360 億 | 市場傳聞（Andy 提供，Interested-party）；出處待補 |
| 服務占比 | 0.41（Tokenomics）；對照 0.50（Epoch「研發推論各半」） | Assumed／第三方 |

## 5. 計算式（逐年）
1. 每 GW 持有成本（年）＝Σ 世代占比 × `IF_HoldEcon` 基準欄。
2. 算力成本 C＝平均 GW × (1)。
3. Q＝平均 GW × 服務占比 ×（1 − `IF_FreeShare`）× Σ（世代占比 × `IF_TokGWPaidMix` 基準欄）× `IF_UtilActual` × `IF_ProdDerate`。
4. 有效牌價 P_eff：2025 用 `IF_PxEffMix25`，2026 用 `IF_PxEffMix`。
5. 實際 m×L＝營收 ÷（Q × P_eff）；所需 m×L＝C ÷（Q × P_eff）；**比值＝實際 ÷ 所需＝營收 ÷ C**（Q 與牌價約掉；Q 只決定 m×L 的水準）。
6. 2026 持平所需平均 GW＝2026 營收 ÷（2026 每 GW 持有成本 × 2025 比值）。

## 6. chat 端預覽值（用 Tokenomics v18.0；建好後應對得上）
| | 2025（1.13 GW） | 2025（1.25 GW） | 2026（2.5／3.0／3.5 GW） |
|---|---|---|---|
| 每 GW 持有成本 | 93.5 億 | 93.5 億 | 97.3 億 |
| 算力成本 | 106 億 | 117 億 | 243／292／341 億 |
| 比值（營收 ÷ 成本） | 1.24 | 1.12 | 1.48／1.23／1.06 |
| 2026 持平所需 GW | 2.99 | 3.31 | — |
讀法：2026 約 3 GW 時與 2025 大致持平、略升；平均 GW 低於約 3.0–3.3 才算上升。
另：服務占比 0.41 時 2025 Q 約 6.4×10¹⁵ token／年（1.25 GW），對 Tokenomics 需求面付費 token 3.8×10¹⁵ 約 1.7 倍（token 帳未全閉合）；實際 m×L 2025 約 1.24–1.64、2026 約 0.49–0.83（隨 GW 與服務占比）。

## 7. 對照列（不進主序列）
- 外部報導算力支出：2025 雲端 160 億（Epoch）、2026 500 億（Brockman 證詞）、2025 推論＋訓練 204 億（Tokenomics SRC）。三者都高於「GW × 持有成本」（約 1.4–2.2 倍）；與計價倍數題同源（Tokenomics token 帳 A 段 Q1），趨勢中固定倍數約掉，但必須逐年用同一口徑。

## 8. 待 Andy 判斷（附預設）
1. L 讀法——預設壽命期價格係數。
2. 服務占比——預設 0.41（Tokenomics），0.50 為對照。
3. 360 億出處——預設標「傳聞」照用。

## 9. 規則提醒
- 主版或下游數值變動須合併後獨立查核；給 Andy 閱讀的檔由未參與代理核數字（Project Instructions 第 5 條）。
- 合併、寄信只在 Andy 指示時做；派代理寫明「不得在任何對外請求中使用使用者 email」。
- CoreWeave 已由 Andy 跑完（2026-10-22 05:23），Tokenomics 下游同步不再擱置（v5.33、v5.36、v15.2、v15.3 累積的名稱變動，見 Tokenomics 索引「下游同步」）。
