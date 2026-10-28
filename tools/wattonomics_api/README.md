# Wattonomics API 工具（2026-10-28 起）

服務 Wattonomics（全市場 GW 帳，舊名 gwledger）的三個 Claude API 工具，在 GitHub Actions 上執行，用 secret `ANTHROPIC_API_KEY`（Max 方案每月 API credit，2026-11-10 到期）。
**三者都只產出候選或核對結果，不改證據檔、不寫 Excel、不寫 main**；判讀與修正由 chat 端與 Andy 決定。結果參考用，不阻擋合併。

慣例沿用 Tokenomics `tools/evidence_audit` 與本 repo `mag/extract`：Batch API（五折）；批次逾 6 小時未完成就取消，未完成的請求改走一般 API（約 2 倍價）；
輪詢逾時則把批次編號寫回請求檔，下次觸發續接、不重送；結果 commit 回結果分支；分支有開啟中的 PR 時留言摘要。

| 工具 | 程式 | 工作流程 | 結果分支 |
|---|---|---|---|
| 1. 證據逐列回查 | `evidence_audit/audit.py` | `wattonomics-evidence-audit` | `claude/wattonomics-audit` |
| 2. 全球 AI 園區候選表 | `site_pipeline/sites.py` | `wattonomics-site-pipeline` | `claude/wattonomics-sites` |
| 3. 每週變動監測 | `weekly_monitor/monitor.py` | `wattonomics-weekly-monitor` | `wattonomics-monitor` |

共用：`common.py`（API、Batch／一般 API 切換、用量、預算）、`.github/scripts/wattonomics_commit.sh`（commit 到結果分支）。

## 觸發方式
- **請求檔（主要）**：在結果分支（工具 1、2）或 `claude/wattonomics-monitor-test`（工具 3）的 `tools/wattonomics_api/<工具>/requests/` 放 JSON 並 push。處理完移到 `requests/done/`。
- **手動**：Actions 頁 → 該工作流程 → Run workflow（`workflow_dispatch` 須本檔已在預設分支，即 PR 合併後）。輸入的預設值就是冒煙測試規模。
- **排程**：只有工具 3，每週一 01:00 UTC＝09:00 台北，在預設分支執行。

### 1. evidence_audit（證據逐列回查）
逐列讀 `evidence_raw/*.csv`（A_us_cloud、B_labs_china_totals、B_delay、F_funding、P_facility_power…），抓每列 `url` 原文，模型比對 value／unit／as_of／basis／tag 等欄與原文。
判定：屬實／小誤／不符／口徑不同／原文找不到；抓不到原文記「無法取得」。**不重試、不繞過**（403、需登入、付費牆、需 JavaScript 皆同）；SEC 一樣不重試（GitHub 雲端 IP 實測被 SEC 擋，F_funding 的 sec.gov 列預期大多「無法取得」）。
同一網址的列合併成一則請求（原文只送一次）；原文過長時保留開頭與該列數字附近段落，上限 `AUDIT_MAX_SRC_CHARS`（預設 120,000 字元）。
```json
{"evidence_dir": null, "files": [], "max_rows": null, "cost_cap_usd": 15}
```
`evidence_dir` 預設 `wattonomics/evidence_raw`，找不到則用 `gwledger/evidence_raw`。輸出 `out/<名>.csv`（依嚴重度排序）、`out/<名>.md`、`out/<名>_usage_<時間>.json`。

### 2. site_pipeline（全球 AI 園區候選表）
1. 種子：evidence_raw 各 csv 的園區欄（`site_or_scope`／`site_or_contract`，程式粗篩掉「全公司、合約、晶片」等）與 `B_summary.md` 表格第一欄。
2. 擴充（`expand: true`）：一則一般 API 請求附 web search，去重並擴充到 `target_sites`（100–150）個全球 AI 園區。可改用 `sites` 直接給清單。
3. 逐園區查核（Batch，每園區一則，附 web search）：MW（IT／設施／未說明）、狀態、營運者、使用者、地點、公告上線時點、分期、延遲報導、來源 URL 與日期、標記（[Verified]／[Interested-party]／[Derived]／[Assumed]）、說話者。
```json
{"max_sites": 150, "expand": true, "target_sites": 130, "searches_per_site": 5, "expand_searches": 25, "cost_cap_usd": 25}
```
輸出 `out/<名>_sites.csv`（園區清單）、`out/<名>_candidates.csv／.md`（候選表）、usage json。

### 3. weekly_monitor（每週變動清單）
玩家：Microsoft、Alphabet、Amazon、Meta、Oracle、CoreWeave、Nebius、OpenAI、Anthropic、xAI、中國大廠（阿里、騰訊、字節、百度、華為），每組一則 Batch 請求附 web search，
找上一週（上週一至週日，台北）的新園區、延期、取消、發債、合資、股權募資。輸出 `out/<週一>_week.md／.csv` 到分支 `wattonomics-monitor`。
**2026-11-10（台北）之後自動停止**：workflow 第一步與 `monitor.py` 都判斷日期，過期不呼叫 API。實際只會排程執行 11-02、11-09 兩次。
週排程要當天出結果：批次 4.5 小時未完成就取消改一般 API。

## 上限（費用與次數）
單價：claude-sonnet-5-5 一般 API 2／10、Batch 1／5 美元每百萬 token（input／output；沿用 `tools/consistency/README.md` [Verified] 2026-10-27）；web search US$10／千次（Batch 同價；platform.claude.com web search tool 文件 2026-10-28）。

| 項目 | evidence_audit | site_pipeline | weekly_monitor |
|---|---|---|---|
| web search 次數上限 | 0（不用搜尋） | 每園區 `searches_per_site` 預設 5、硬上限 8；擴充段預設 25、硬上限 40；園區數硬上限 150 → 全量最多 1,240 次 | 每玩家預設 6、硬上限 8 → 每週最多 88 次 |
| 每則 max_tokens（輸出） | 8,000 | 6,000（擴充段 32,000） | 6,000 |
| 輸入 token 控制 | 原文 ≤120,000 字元／網址 | 搜尋結果由 max_uses 限制 | 同左 |
| 每次預設費用上限 `cost_cap_usd` | 15 | 25 | 4 |
| 估計費用（全量） | US$5–12（約 510 列／約 300 網址；SEC 列不送模型） | US$18–26（150 園區 × 約 US$0.12–0.17＋擴充 1–2） | 每週 US$1.5–2.5；兩週合計 ≤5 |
| 冒煙測試估計 | 5 列 ≈ US$0.05–0.3 | 3 園區 ≈ US$0.3–0.5 | 1 玩家 ≈ US$0.2 |

預算機制（`common.py`）：
- **本批上限 US$140**（`BATCH_BUDGET_USD`；10-28 17:19 Andy 同意跑 A 美國以外園區擴充、B Opus 交叉回查、C Q3 財報萃取後由 85 放寬）：每次執行前加總所有遠端分支上 `tools/wattonomics_api/*/out/**/usage*.json` 的 `est_usd`（同 blob 只算一次），本次可用＝min(請求上限, 140－已用, 190－5－本批以外已用－已用)。
- **累計停止線 US$190**（`HARD_STOP_USD`，含 5 美元緩衝；本期額度 200）：本批以外已用的 credit 用 repo variable `API_PRIOR_SPEND_USD` 設定；未設時 workflow 預設 80（Andy 10-28 17:17：剩約 100；本批當時已用約 20 → 本批以外約 80）。代理無法經代理伺服器設定 repo variable，改值時直接改 workflow 的預設。
- 送出前按事前估計截量：evidence_audit 從最後的網址往前砍；site_pipeline／weekly_monitor 依「每則估計 × 則數」減少園區或玩家。被砍的在摘要標出。
- 估計是事前推估，實際以 usage.json（API 回傳的 token 與搜尋次數）為準；Batch 中途不能停，單次實際費用可能略高於上限。
- 三個工作流程共用 concurrency group `wattonomics-api-budget`，一次只跑一個，預算不會重複計算。

## 用量檔 usage.json
每次執行寫 `out/<名>_usage_<UTC 時間>.json`：`input_tokens`、`output_tokens`、cache tokens、`web_search_requests`、Batch／一般 API 則數、`est_usd`、狀態。失敗或處理中也寫。

## 隱私與對外請求
- 任何請求都不帶使用者 email。抓網頁 User-Agent：`company-models wattonomics tool (github.com/YinchenChang/company-models)`。
- SEC（sec.gov）一律用 `company-models research tool YinchenChang@users.noreply.github.com`（Andy 2026-10-27 同意用 GitHub 免回覆地址）。

## 冒煙測試（PR 合併前不要跑全量）
合併前（workflow_dispatch 尚不可用）用 push 觸發；每個約 US$0.05–0.5：
```bash
# 1. evidence_audit 5 列
git checkout -B claude/wattonomics-audit origin/claude/wattonomics-api
echo '{"max_rows": 5, "cost_cap_usd": 1}' > tools/wattonomics_api/evidence_audit/requests/smoke_5rows.json
git add -A && git commit -m "smoke: evidence_audit 5 列" && git push -u origin claude/wattonomics-audit
# 2. site_pipeline 3 園區（不擴充）
git checkout -B claude/wattonomics-sites origin/claude/wattonomics-api
echo '{"max_sites": 3, "expand": false, "searches_per_site": 3, "cost_cap_usd": 1}' > tools/wattonomics_api/site_pipeline/requests/smoke_3sites.json
git add -A && git commit -m "smoke: site_pipeline 3 園區" && git push -u origin claude/wattonomics-sites
# 3. weekly_monitor 1 玩家
git checkout -B claude/wattonomics-monitor-test origin/claude/wattonomics-api
echo '{"players": ["Oracle"], "searches_per_player": 4, "cost_cap_usd": 0.5}' > tools/wattonomics_api/weekly_monitor/requests/smoke_oracle.json
git add -A && git commit -m "smoke: weekly_monitor 1 玩家" && git push -u origin claude/wattonomics-monitor-test
```
合併後也可在 Actions 頁 Run workflow，用預設輸入（audit 5 列／US$1；sites 3 園區、不擴充／US$1；monitor 填 players=Oracle）。
檢查：結果分支上 `out/*.md` 摘要、`out/*_usage_*.json` 的 `est_usd` 與搜尋次數、`out/_run.log`。

全量（冒煙通過、Andy 同意後）：audit `{"cost_cap_usd": 15}`；sites `{"max_sites": 150, "expand": true, "target_sites": 130, "cost_cap_usd": 25}`；monitor 交給每週排程。
