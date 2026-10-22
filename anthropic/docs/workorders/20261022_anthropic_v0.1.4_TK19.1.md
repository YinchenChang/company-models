# 工作單 Anthropic v0.1.4｜Tokenomics v19.1 同步＋η 診斷兩項補充（2026-10-22，chat 端）

- 來由：(1) Tokenomics v19.1 已合併（master `c2ecb71`，CURRENT `20261022_Tokenomics_v19.1.xlsx`；Luna CSA top-k 修正）；下游同步通知見 company-models PR #80 的 `docs/workorders/20261022_TK_v19.1_下游同步通知.md`；下游同步擱置已於 2026-10-22 解除。chat 端已跑 `check_tk_snapshot`：本模型讀取的 126 值中 DIFF 39、OK 87（`IF_TokGW_Luna`、`IF_FullCost*`、`IF_FullCostDefault_*`、`IF_ProgGWyr_Luna`、`IF_Alloc*`、`IF_TrainCost_Luna`、`IF_RevGW_Luna`、`L1_Ans3*`、`L1_RevGW_PQ_VR200*`、`L1_TokGW_Paid*`、`L1_TokLedger_2026Q2`、`L1_Gap*`）。(2) Andy 2026-10-22 對 v0.1.3 B 段兩題「照預設」：η 拆解**不加**「TK 絕對產能水準」因子；利用率因子**計入** TK 生產折減。(3) `dist/_待刪除/` 依預設刪除。
- 分支：沿用 `claude/anthropic-v0.1.3`（PR #69 未合併；v0.1.4 疊在同一 PR，PR 標題改「Anthropic v0.1.3＋v0.1.4」）。**不得合併。**
- 必讀：`anthropic/CLAUDE.md`（含第 8 節 S-1 替換程序）、`anthropic/docs/workorders/20261008_anthropic_共同規則.md`、`20261022_anthropic_v0.1.3_共同.md`、v0.1.3 報告 `docs/reports/20261022_v0.1.3.md`、進度檔。規則同 v0.1.3 共同說明（新列加在尾端、公式不含常數、每步 commit＋push＋進度檔、pytest 依檔分開跑、長命令背景執行但必須等完、不得合併／force push／改 main、不得在對外請求中使用使用者 email）。
- Tokenomics：`GIT_LFS_SKIP_SMUDGE=1 git clone --depth 1 https://github.com/YinchenChang/tokenomics /tmp/tk`（驗證 master `c2ecb71`、CURRENT v19.1；若 master 已更新，採最新 CURRENT 並在報告寫明）。

## 步驟（每步 commit＋push＋進度檔）
1. **快照 v19.1**：TK_Link 重建；`model/20261022_Anthropic_v0.1.4.xlsx`（v0.1.3 檔 `git mv` 到 archive、更新 CURRENT）。逐一列出 39 個 DIFF 名稱的新舊值與本模型是否實際讀進計算（v0.1.2 報告的語意核對表為底）。
2. **η 診斷補充**（EtaDiag 頁尾端，η 值不得改變）：(a) 利用率因子乘入 TK 生產折減（讀 TK 具名範圍；若無具名範圍，讀表並標「讀表（非具名）」，列 Tokenomics 缺口）；殘差隨之重算，報告列前後值；(b) 新增對照列：本模型未解釋殘差 vs TK `L1_GapUnexplained`（OpenAI 口徑，Analogy）與比值；(c) 小表「各 TK 版本的 η 與殘差」加一欄 v19.1。
3. **刪除** `dist/_待刪除/`（git rm；Drive 上仍有副本）。
4. 重算（LibreOffice 錯誤 0、`CHK_Errors`＝0）；parity 期望值有意識地更新；pytest 依檔分開全過。
5. HTML 與 dist：`dist/20261022_Anthropic收支模型_v0_1_4.html`／`.xlsx`（v0.1.3 dist 檔直接 git rm，不再放 `_待刪除/`）；「v0.1.3 更新了什麼」節保留，前面新增一句 v0.1.4 摘要；「資料與版本」改 TK v19.1；以 Playwright 逐節讀畫面（桌機與手機寬度）。
6. 報告 `docs/reports/20261022_v0.1.4.md`＋`_對照.xlsx`：①TK 名稱新舊值與是否進計算；②**TK 變動造成的差異拆解**：每 VR 等值 GW 差額 2025–2030、覆蓋率、η、供給 VR 等值 GW、資金三層、擴張三檔，v0.1.3 → v0.1.4，逐群拆解（`IF_TokGW_Luna`、其他）；③η 診斷補充前後；④已套用的預設。CHANGELOG、交接檔、進度檔；PR 留言 `[Anthropic 回報] v0.1.4｜完成｜日期`。
