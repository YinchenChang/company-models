# WhiteFiber v0.1 共同規則（2026-10-08，r0，chat 端）

適用於工作單 v0.1a、v0.1b、v0.1c。每張工作單開工前必讀本檔、該張工作單、`whitefiber/docs/plan/20261008_WhiteFiber_命題與驅動對照表_r1.md`、repo 根目錄 `README.md`、`whitefiber/CLAUDE.md`、`whitefiber/README.md`。
衝突時的優先順序：**該張工作單 > 本檔 > repo 根目錄 README > whitefiber/CLAUDE.md**（CLAUDE.md 是 CRWV／Oracle 時期的守則，其中「判斷口徑停下來問 Andy」「共識數據只能由 Andy 提供」兩條，本次由工作單取代；回報第一行改為 `[WhiteFiber 回報]`）。

`whitefiber/` 是分支 `claude/oracle-v0.2` @ `e4540c3` 的 `oracle/` 副本（Oracle v0.2：含建設延誤模組與租賃調整後槓桿；尚未合併到 main）：引擎已含 MW × 每 MW 收入主軸（OCI）、首期校準（converge 模式）、傳統事業多線區塊、預付款與重大財務組成、股利、現金稅、未起租租約排程、融資瀑布（投資級債 → 股權 → 高息債，可轉債步驟）、強制轉換特別股、CAPM、分部 EV/EBITDA、穩態汰換、EBITDAR 率、延誤模組。`company.json`、`dist/`、`docs/handoff/` 目前仍是 Oracle v0.2，v0.1b／v0.1c 會全部換成 WhiteFiber。參考：`oracle/docs/`、`nebius/docs/` 的工作單與進度檔（同一引擎的改寫路徑與踩過的坑；Oracle v0.2 的報告與進度檔在分支 `claude/oracle-v0.2` 的 `oracle/docs/reports/`，用 `git show origin/claude/oracle-v0.2:<路徑>` 讀）。

## 1. 分段與並行
- 每完成一個步驟就 commit 並 push，同時覆寫進度檔；任何時候中斷，下一個代理都能從進度檔接手。
- 同一 repo 有其他代理同時在其他資料夾工作：**只改 `whitefiber/` 內的檔案**（v0.1c 另可改 repo 根目錄 README 表格中 WhiteFiber 那一列）。不得動其他資料夾。

## 2. 分支、PR、合併
- 三張工作單**共用一個分支** `claude/whitefiber-v0.1`（已由 chat 端建立並推送）與**一個 PR**（base `main`，標題「WhiteFiber v0.1：收支與評價模型」；chat 端已開 draft，編號見進度檔）。各張依序在同一分支接續 commit，各張完成時在同一 PR 留言回報。
- PR 操作用 GitHub REST API（`gh api ...`）；GraphQL 在此環境不可用。
- **代理不得合併、不得 force push、不得改 `main`、不得刪除分支或歷史。**
- 推送前先 `git fetch origin claude/whitefiber-v0.1` 並在其上接續（不要從 `main` 重開）；淺層 clone 推送若遇 HTTP 413，先 fetch 再推。

## 3. 進度檔（接手用，必做）
- 路徑：`whitefiber/docs/reports/20261008_whitefiber_進度.md`（三張共用；每張開工時新增自己的段落，不刪其他段落）。
- 最上方「目前狀態」每次 push 前覆寫：**已完成**（含 commit）／**下一步**（具體到下一個代理可以直接動手）／**未解問題**。
- 下方各段「步驟紀錄」表：步驟｜狀態｜commit｜備註。
- 每步完成 → 更新進度檔 → commit → push，三件事一起做。
- 察覺額度或時間快用完：立即停止新工作，先 commit、更新進度檔（寫清楚做到哪、下一步），push 後結束。

## 4. 資料紀律
- 每一筆數字都要有：數值、單位、期間、來源（文件名稱＋網址）、文件日期、擷取日、標記。
- 標記：[Verified] 可公開查證；[Interested-party] 來源有利害關係（寫明是誰與動機）；[Analogy] 由可比對象推估（寫明對象與差異）；[Assumed] 無實證；[Derived] 由其他數字換算（寫明公式）。Analogy／Assumed 一律給區間。
- WhiteFiber 自己的新聞稿、法說、投影片＝[Interested-party]；經查核的 10-K（FY2025）財務報表可標 [Verified]；10-Q（核閱、未查核）標 [Interested-party] 並註明「10-Q 核閱數」。法說逐字稿多為二手（Motley Fool 等），註明「二手逐字稿」。
- **不得杜撰數字。** 找不到就標「不適用」或列入資料缺口，寫明試過哪些來源。「找不到」與「不存在（公司未揭露）」分開寫。
- **SemiAnalysis（含 InferenceX）類資料不可單獨引用**：須有第二個獨立來源佐證，否則只列參考、不作輸入。
- **不以公司自己的預測反推參數**（例如以「年化 > $200M」目標 ÷ MW 回推每 MW 單價）。雲端每 MW 單價由 Tokenomics 正向推導；公司數字與合約隱含單價只作對照列與敏感度。例外：(a) 以「已實現實際數」（Q2 經常性雲端營收）校準首期期初狀態；(b) **已簽約的託管合約排程**（NC-1：金額、MW、期間、年調）屬合約事實，可作託管營收輸入（對照表 r1 第 4 節第 4 條），須寫明。
- 同一底層來源的多個呈現（例如 StockAnalysis 與 MarketScreener 都引用 S&P Capital IQ）不算獨立交叉驗證，要寫明。
- WhiteFiber 上市未滿兩年（2025-08 IPO）、分析師家數少：共識與 beta 可能只有少數來源，照實記錄家數與期間。

## 5. 判斷類事項
- 照對照表 r1 第 4 節的預設處理，不停下來問；沒涵蓋的判斷，選最保守且最常見的口徑，在進度檔與最終報告「已套用的預設」逐項列出：問題｜採用的預設｜替代選項｜對結果的影響方向（能量化就量化）。
- 範圍：做到 CRWV v4.5＋Nebius v0.1＋Oracle v0.2 水準；CRWV 未完成的待辦（5a 精簡、5b 只讀檢視器、折舊修正、機率加權目標價、GPU 批次與續約價格衰退）不做。

## 6. 停止條件（只有碰到這些才停；其餘一律照預設做完）
1. 必要的 repo（`YinchenChang/company-models`、`YinchenChang/Tokenomics`）無法讀取或無法推送。
2. 主要一手來源全部無法取得（SEC EDGAR 與 WhiteFiber 投資人關係網站都拿不到 2026 Q2 10-Q）。
3. 同一個 verify 失敗，經三次實質不同的修正仍無法通過，且原因是模板結構限制而非本次改寫的錯誤。
4. 要讓模型成立，必須做第 5 節列為「不做」的範圍。
5. 需要改 `main`、force push、刪除分支或歷史、或改 `whitefiber/` 以外的資料夾。

停下時：commit 並 push 現況、進度檔「未解問題」寫明卡在哪、選項與建議，在 PR 留言（第一行 `[WhiteFiber 回報] v0.1x｜停止｜YYYY-MM-DD`），然後結束。

## 7. 回報
- 完成時在 PR 留一則留言（第一行 `[WhiteFiber 回報] v0.1x｜完成｜YYYY-MM-DD`）：結論、步驟與 commit、verify.sh 結果摘要、已套用的預設、資料缺口、下一張工作單需要知道的事。非工程師看得懂的繁體中文，結論先行。
- 你的最後一則訊息（回給派工的 chat 端）：同上內容的精簡版，加上分支最新 commit。

## 8. 環境
- 工具檢查與補裝見 `whitefiber/CLAUDE.md`「環境」第 0 步（soffice、openpyxl、playwright、Chromium）。`whitefiber/scripts/verify.sh` 基線應全過（Oracle v0.2 為 19 項）。
- 產物在 `whitefiber/out/`（不納入版控）。
- SEC EDGAR 需 User-Agent（例如 `User-Agent: company-models research admin@example.com`，不得留空）；WhiteFiber 的 CIK 自 EDGAR 公司搜尋取得並記入總帳。shell 若無法連 sec.gov，改用 WebFetch 讀同一份文件並註明。
