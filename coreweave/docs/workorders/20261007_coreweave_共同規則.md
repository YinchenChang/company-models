# CoreWeave × Tokenomics 改造：共同規則（2026-10-07，chat 端）

適用於工作單 W1、W2、W3。每張開工前必讀：本檔、該張工作單、repo 根目錄 `README.md`、`coreweave/CLAUDE.md`、`coreweave/README.md`、`coreweave/docs/handoff/` 交接檔。
衝突時的優先順序：**該張工作單 > 本檔 > repo 根目錄 README > coreweave/CLAUDE.md**。CLAUDE.md 中「判斷口徑停下來問 Andy」一條由本檔第 5 節取代（Andy 2026-10-07：判斷類一律套用預設、不逐項詢問）。

## 0. 目的（一句話）
讓 CoreWeave 模型中所有「算力相關的產業與物理層資料」（每 MW GPU 數、IT 資本支出、折舊年限、電力與 PUE、營運費用、持有成本）改為引用 Tokenomics（`YinchenChang/Tokenomics` 的 `IF_`／`L1_` 具名範圍），公司專屬資料（價格、合約、MW 路徑、債務、租約、世代組合）留在 `company.json`；最後以 v4.6 成品與前後對照報告回答「每 MW 收入、每 MW 成本改變了多少、為什麼」。

## 1. 分段與進度檔（防止額度用完時遺失）
- 每完成一個步驟：更新進度檔 → commit → push，三件事一起做。
- 進度檔：`coreweave/docs/reports/20261007_coreweave_tokenomics_進度.md`（三張共用；各張新增自己的段落，不刪他人段落）。最上方「目前狀態」每次覆寫：**已完成**（含 commit）／**下一步**（具體到下一個代理可直接動手）／**未解問題**。
- 察覺額度或時間快用完：停止新工作，先 commit、寫清楚進度檔，push 後結束。

## 2. 分支、PR、合併
- 每張工作單一個分支、一個 PR（base `main`），分支名寫在工作單。開工第一步：從最新 `main` 開分支、push、開 draft PR（`gh api repos/YinchenChang/company-models/pulls -f title=... -f head=... -f base=main -F draft=true -f body=...`；GraphQL 不可用）。
- **代理不得合併、不得 force push、不得改 `main`、不得刪除分支或歷史。** 合併由 chat 端審查後執行。
- 推送前 `git fetch origin main`；遇 HTTP 413 先 fetch 再推。
- 只改 `coreweave/` 與本工作單指定的根目錄路徑（`tools/tokenomics/`）；不得動 `nebius/`。

## 3. Tokenomics 引用規則
- 只引用 `IF_`（不含 `IF_Hdr*`）與 `L1_` 名稱（Tokenomics README「具名範圍與下游連結」）。不得引用 Inputs、DC_Cost 等內部頁的儲存格。
- 版本固定：快照檔記錄 Tokenomics 檔名、`model/CURRENT`、commit SHA、擷取日；每個值記名稱、世代、成本情境、單位、工作表!儲存格。
- 世代欄：Hopper H100、GB200 NVL72、GB300 NVL72、VR200 NVL72、Rubin Ultra；成本情境：低成本／基準／高成本。模型主值一律取「基準」，低／高只作敏感度。
- 單位：Tokenomics 每 GW＝IT 關鍵電力。CRWV 的 MW 若口徑不同，一律在 CRWV 端換算並標明，不改 Tokenomics。
- 缺少需要的名稱時：不得自行改連內部頁；記入進度檔「未解問題」，由 chat 端在 Tokenomics 開工作單補上（v5.25 已在處理第 W1 步驟 3 列出的名稱）。

## 4. 資料紀律
- 每筆數字：數值、單位、期間、來源（文件名稱＋網址）、文件日期、擷取日、標記（[Verified]／[Interested-party]／[Analogy]／[Assumed]／[Derived]）。Analogy／Assumed 一律給區間。
- 不得杜撰數字；「找不到」與「不存在」分開寫，並寫明試過哪些來源。
- SemiAnalysis（含 InferenceX）類資料不可單獨引用，須有第二個獨立來源；同一底層來源的多個轉載不算獨立。
- **不以公司自己的數字或指引反推參數**（例如以 ARR ÷ MW 回推單價、以 CapEx 指引回推每 MW 建置成本）。公司數字只作對照列。

## 5. 判斷類事項
- 依工作單寫明的預設處理，不停下來問。工作單沒涵蓋的判斷：選最保守且最常見的口徑，並在進度檔與 PR 回報的「已套用的預設」逐項列出：問題｜採用的預設｜替代選項｜對結果的影響方向（能量化就量化）。
- 範圍：CRWV 原待辦（5a 精簡、5b 只讀檢視器、折舊修正、機率加權目標價、GPU 批次與續約價格衰退、Tokenomics 壽命期價格係數 L）**不做**。

## 6. 停止條件（只有碰到才停；其餘照預設做完）
1. `YinchenChang/company-models` 無法推送，或 `YinchenChang/Tokenomics` 無法讀取。
2. 同一個 verify 失敗，三次實質不同的修正仍不過，且原因是模板結構限制。
3. 必須做第 5 節列為「不做」的範圍才能成立。
4. 需要改 `main`、force push、刪分支或歷史。
停下時：commit＋push、進度檔寫明卡點與選項建議、PR 留言（第一行 `[CRWV 回報] Wx｜停止｜YYYY-MM-DD`），然後結束。

## 7. 回報
- 完成時 PR 留言（第一行 `[CRWV 回報] Wx｜完成｜YYYY-MM-DD`）：結論先行、步驟與 commit、verify.sh 摘要、已套用的預設、資料缺口、下一張需要知道的事、「需要 Andy 做的事」（沒有寫「無」）。
- PR 改為 ready：`gh api -X POST repos/YinchenChang/company-models/pulls/<n>/ccr/ready_for_review`（不可用時在留言註明）。
- 文件與回報一律繁體中文、非工程師看得懂。

## 8. 環境
- 工具：soffice、openpyxl、playwright、Chromium（見 `coreweave/CLAUDE.md` 第 0 步）。W0（chat 端，2026-10-07）在本環境跑 `coreweave/scripts/verify.sh --vs-dist` 21 項全過，可作為環境正常的判準。
- Tokenomics 唯讀 clone：`GIT_LFS_SKIP_SMUDGE=1 git clone --depth 1 https://github.com/YinchenChang/tokenomics <路徑>`；現行檔名見 `model/CURRENT`。
