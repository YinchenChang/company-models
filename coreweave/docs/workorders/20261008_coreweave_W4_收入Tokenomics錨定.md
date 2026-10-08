# 工作單 CoreWeave W4｜每 MW 收入改以 Tokenomics 為錨＋公司因素（2026-10-08，r1，chat 端）

- r1 修訂（取代 r0「Q2 實績校準／GPU 小時長約價」）：Andy 2026-10-08「每 MW 收入要跟 Tokenomics 連動並校準；CoreWeave 可以有自己的調整，但要基於 Tokenomics，再加上公司層面的因素」「你直接做掉」。r0 的選項 A（以 Q2 實績調參數）屬「擬合公司數字」，不採用；選項 B（`gpuHr`）的價格不以 Tokenomics 為錨，降為對照。見 `coreweave/CLAUDE.md` 已決定事項 14。
- 前置：W3（v4.6）已合併（main `6410860`）。必讀：共同規則、W2／W3 PR 回報、W3 報告第 5 節（Q2 對帳）、`coreweave/CLAUDE.md` 已決定事項 12、13、14。
- 分支：從最新 main 開 `claude/coreweave-w4-revenue`；PR 標題「CoreWeave v4.7：每 MW 收入以 Tokenomics 為錨」。

## 目標（一句話）
每 MW 年收入＝**Tokenomics 錨**（各世代每 MW 經濟持有成本，即 WACC 10% 下的打平租金）×**公司因素**（定價倍數 k、計費利用率），再加非算力服務收入；Q2 實績只作驗證，不作校準目標。

## 公式（新方法 `methodology.perMw.revenue=tkAnchor`，設為預設）
1. **錨**：`anchor_t ＝ Σ_g 平均在役世代占比_{t,g} × IF_HoldEcon_g（基準成本情境）÷ 1000`（US$bn/MW/年）。世代占比沿用 W2 的 `fleet.openMix`／`newMix` 與汰換邏輯。`IF_HoldEcon`＝`IF_GPUhrEcon × 每 MW GPU 數 × 8,760`，報告附一行核對（快照值）。
2. **公司因素一：定價倍數 k**（`company.json` → `pricing.anchorMultiple`）：
   - 依合約組合加權：`k_t ＝ 長約占比_t × k_長約 ＋（1 − 長約占比_t）× k_現貨`。
   - `k_長約`、`k_現貨` 以「市場價格 ÷ `IF_GPUhrEcon`（同世代）」為證據，存證據表（每筆：價格、世代、合約型態與期間、來源、日期、標記、倍數）。chat 端已整理的起點（W1 資料，[Derived]）：IREN–Microsoft GB300 五年約 0.76；H100 現貨指數 1.76；B300 現貨指數（類比 GB300）2.29；CoreWeave GB200 牌價 5.0（[Interested-party]，牌價非成交價，只列不用）。
   - **預設**：`k_長約` 基準 0.76、區間 0.70–1.00（[Analogy]：目前只有一筆可比長約；步驟 1 補到第二筆獨立來源時更新）；`k_現貨` 基準 1.76、區間 1.5–2.3（[Analogy]）；長約占比以 RPO 涵蓋的產能 ÷ 在役產能估計（[Derived]，寫明算式），區間列出。
   - **不得**用 Q2 營收、ARR、RPO 金額反推 k。
   - 合約期內 k 固定；續約價格衰退與 Tokenomics 壽命期係數 L 不做（已決定事項：先不要）。
3. **公司因素二：計費利用率**：沿用 `m.util`；`Billable MW ÷ Accepted MW`（爬坡）沿用現有機制。兩者都要在每 MW 彙總表中顯示，避免利用率在收入端被重複乘（W2 已處理：revMW 定義為 100% 計費時數的值）。
4. **每 MW 年收入（100% 計費時數）＝ anchor_t × k_t**，取代 `m.revMW` 進入引擎；非算力服務收入不變。
5. **上限檢查（新）**：CRWV 每 MW 年收入 ÷ `IF_RevGWFleet`（客戶每 GW 付費 token 營收，換算每 MW；依世代）＝「neocloud 拿走客戶營收的比例」；> 50% 時在「檢查」頁警示（門檻放 `methodology.checks`）。
6. **對照列（保留，不驅動）**：v4.6 舊輸入 `m.revMW`；`gpuHr` 路線（若有價格）；期末 ARR ÷ MW；Q2 實際年化 ÷ 平均在役 MW 與 ÷ 計費 MW。

## 範圍界線
- 只動收入口徑。營運成本依已決定事項 13 維持 Tokenomics 滿載值；MW 口徑、折舊年限、資本支出、租金路徑都不動。
- 舊方法 `legacy` 必須仍可精確重現 v4.6（verify 以 legacy 跑 `--vs-dist` 對 v4.6 0 差異）。

## 步驟（每步：verify → 進度檔 → commit → push）
0. 開分支、draft PR、進度檔新增「W4」段落。確認 Tokenomics 快照為 v5.26（或 master 最新；若有新版，重抓並在報告列出變化）。
1. **證據補充**（共同規則第 4 節）：GB200／GB300／VR200 的長約或多年期合約每 GPU 小時（或每 MW）價格，至少再找一筆獨立於 IREN–Microsoft 的長約（例如其他 neocloud 與 hyperscaler／實驗室的揭露合約、房東租約換算不算）；CRWV 合約組合（長約與隨需比例、平均合約年期）。找不到與不存在分開寫。SemiAnalysis 不可單獨引用。
2. **實作** `tkAnchor`（Excel 公式＋JS 雙引擎，經 `TK_` 具名範圍引用，不寫死），`pricing.anchorMultiple` 欄位、證據表、上限檢查、對照列；`fields_doc.py --write`。
3. **Q2 驗證（不是校準）**：FY26 模型每 MW 收入對 Q2 實際年化，差距拆成 (i) 爬坡分母（計費 MW 對在役 MW）、(ii) 利用率、(iii) 定價倍數 k（模型 k 對 Q2 隱含 k＝Q2 年化 ÷ 錨）、(iv) 世代組合、(v) 其他；各項加總＝總差距（誤差 < 0.01）。差距 > 5% 時依已決定事項 2 附原因與類型，但**不得回頭改 k 去貼近 Q2**。
4. **敏感度**：k_長約 0.70／1.00、k_現貨 1.5／2.3、長約占比區間、Tokenomics 低／高成本（錨與成本同時動）、Rubin Ultra 世代版；對三情境目標價與融資缺口（建置時快照機制）。
5. **升版 v4.7**：VLOG（`vlog.py`、`tail.js`）、`dist/` 換 v4.7、交接檔、README；對 v4.6 成品跑 `--vs-dist`，`--expect` 只含收入相關列與其下游；變動拆解依已決定事項 12，(d) 方法變更再拆：① 錨取代舊輸入（k＝1）→ ② 套用 k → ③ 上限檢查（應為 0）。
6. **對照報告**：`docs/reports/20261008_coreweave_v4.7_收入錨定.xlsx`＋md：每 MW 收入 v4.6 vs v4.7（FY26–FY30、三情境）、錨與 k 的逐年值、證據表、Q2 驗證拆解、敏感度、三情境目標價與融資缺口；結論先行三句內回答「CRWV 每 MW 收入是 Tokenomics 打平租金的幾倍、這代表什麼」。
7. **整體 verify 與回報**：`verify.sh` 全過（含離線開啟）；PR 回報貼摘要、已套用的預設、資料缺口；「需要 Andy 做的事」：合併前以真正的 Excel 開啟 v4.7 檢查。

## 驗收（chat 端逐項核）
- 每 MW 收入可追溯為 `IF_HoldEcon`（`TK_` 名稱）× 世代占比 × k；k 的每個值都有證據表來源，沒有一筆來自 Q2 營收、ARR 或 RPO 金額。
- 營運成本相關格對 v4.6 0 差異；legacy 對 v4.6 0 差異。
- Q2 驗證拆解加總＝總差距；目標價變動拆解加總＝總變動（三情境，誤差 < 0.01 美元／股）。
