# 工作單 Anthropic v0.1.2｜Tokenomics 快照 v5.26 → v18.0（2026-10-21，chat 端）

- 來由：Andy 2026-10-21 23:00「我發現你用的 TK 是舊版，請依最新版更新」。下游同步（Project 索引「Tokenomics 下游同步」條）；Andy 明確指示本模型先做。
- 分支：沿用 `claude/anthropic-v0.1.1`（PR #56 未合併；v0.1.2 疊在 v0.1.1 上，同一個 PR；PR 標題改為「Anthropic v0.1.1＋v0.1.2」）。**不得合併。**
- Tokenomics：master `868e2cc`，CURRENT `20261021_Tokenomics_v18.0.xlsx`（唯讀 clone 在 `/home/claude/tk`，已 checkout 該提交；先驗證 CURRENT 與 SHA）。
- chat 端已跑 `check_tk_snapshot`：本模型讀取的 89 個值中 OK 59、DIFF 30、MISSING 0（名稱全在）。DIFF：`IF_TokGW_*`、`IF_HoldEcon`、`IF_FullCost_*`、`IF_FullCostDefault_*`、`IF_ProgGWyr_*`、`IF_TrainCost_*`、`IF_RevGW_*`、`IF_Alloc*`（7）、`L1_Ans3`（含低高）、`IF_CapexTotal`。NonNV 表與 PUE 讀表值不變。
- 必讀：`anthropic/CLAUDE.md`、共同規則、v0.1.1 工作單與報告、Tokenomics `docs/plan/Tokenomics_downstream_contract.md`（**v6.0 起各版契約變更條文**：v6.0 `_Prod`／`_Life` 定義、v9.0 快取命中進產能、v10.2 SLO、v11–v15 批次倍數與 GB200 η_d、v18.0 新名稱）。
- 不得在任何對外請求（含 HTTP User-Agent）中使用使用者 email。超過 2 分鐘的命令背景執行，但等它跑完再回報。

## 步驟（每步 commit＋push＋進度檔）
1. **快照更新**：builder 以 v18.0 重建 TK_Link（同一份名稱清單；記錄 master SHA、CURRENT 檔名、SHA-256、讀取日）；`model/20261021_Anthropic_v0.1.2.xlsx`（v0.1.1 檔 `git mv` 到 archive、更新 CURRENT）。
2. **語意核對**：逐一列出本模型讀取的每個 DIFF 名稱，在契約條文中找出造成變動的版本與原因，判定是「定義改變」還是「上游輸入改變」；對每個定義改變的名稱，檢查本模型公式的用法是否仍正確（特別是：契約第 19 行「不得把 `IF_TokGW_*` 直接當成生產環境產能」——本模型以 η 在 2025 校準，確認 η 是否吸收了 `IF_TokGW_*` 的變動；`IF_FullCost*` 只作參考列）。需要改公式才正確者：照改並列入報告（工程類）；屬判斷者：套用最保守口徑並列入「待 Andy」。
3. **新名稱（可選）**：v18.0 新增的 `L1_RevGW_PQ_VR200`（含 `_Life`、`_FY`）與本模型的每 VR 等值 GW 營收是同一個量的 Tokenomics 版本：在 Checks 或 Cost 尾端加**對照列**（不回饋計算），並註明契約所說的 token 帳未閉合（`L1_TokLedger_2026Q2` 約 6.26 倍）。
4. 重算（LibreOffice 錯誤 0、`CHK_Errors`＝0）；parity 期望值有意識地更新；pytest 分檔執行全過（全套單次執行會因記憶體被終止）。
5. HTML 與 dist：`dist/20261021_Anthropic收支模型_v0_1_2.html`／`.xlsx`（v0.1.1 dist 檔 `git mv` 到 `dist/_待刪除/`）；HTML「資料與版本」改 TK v18.0；「①之二」改為「v0.1.2 更新了什麼」，同時保留 v0.1.1 的變動摘要一句。
6. 報告 `docs/reports/20261021_v0.1.2.md`＋`_對照.xlsx`：①TK 名稱新舊值表（30 個 DIFF）與每個的變動原因；②**TK 變動造成的差異拆解**：命題輸出（每 VR 等值 GW 差額 2025–2030、覆蓋率、η、供給 VR 等值 GW、累計外部資金需求、現金谷底）v0.1.1 → v0.1.2，逐群拆解（TokGW 群、HoldEcon、CapexTotal、其他）；③語意核對結論與已套用的預設；④對照列（`L1_RevGW_PQ_VR200` vs 本模型）。結論先行、繁體中文。
7. CHANGELOG、交接檔第 3 節（TK 版本）、進度檔；PR 留言第一行 `[Anthropic 回報] v0.1.2｜完成｜2026-10-21`。
