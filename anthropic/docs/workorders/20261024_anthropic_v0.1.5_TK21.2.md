# 工作單 Anthropic v0.1.5｜Tokenomics v21.2 同步（2026-10-24，chat 端）

- 來由：Tokenomics v21.2 已合併（master `45ce1ca`，CURRENT `20261023_Tokenomics_v21.2.xlsx`；計價倍數 `AL_SpendBasis` 0.85 → 1.8）。下游同步通知：`docs/workorders/20261024_TK_v21.2_下游同步通知.md`。chat 端已跑 `check_tk_snapshot`（快照 v19.1 對 CURRENT v21.2）：本模型讀取的 127 值中 OK 117、**DIFF 10**：`L1_TokLedger_2026Q2`、`L1_GapSpendBasis`（＋`_Lo`／`_Hi`）、`L1_GapProduct`（＋`_Lo`／`_Hi`）、`L1_GapUnexplained`（＋`_Lo`／`_Hi`）。v20.0–v21.1 期間本模型讀取的值沒有變動。
- 預期：這 10 個名稱只進對照列（a4 Q 對照 `pq_led`、`pq_q_ratio` 註記；a5 η 對照頁；HTML η 節），**本模型自身的每 VR 等值 GW 差額、覆蓋率、η、供給、資金、擴張數值應全部不變**。若有任何非對照數值改變，停下列入「待 Project 判斷」，不自行處理。
- 分支：`claude/anthropic-v0.1.5`（從 main 開）。開 PR 到 main，**不得合併**。
- 必讀：`anthropic/CLAUDE.md`（含第 8 節 S-1 替換程序）、`anthropic/docs/workorders/20261008_anthropic_共同規則.md`、`20261022_anthropic_v0.1.3_共同.md`、上一張 `20261022_anthropic_v0.1.4_TK19.1.md` 與報告 `docs/reports/20261022_v0.1.4.md`、進度檔。規則同前（新列加在尾端、公式不含常數、每步 commit＋push＋進度檔、pytest 依檔分開跑、長命令背景執行但必須等完、不得合併／force push／改 main、不得在任何對外請求中使用使用者 email）。
- Tokenomics：`GIT_LFS_SKIP_SMUDGE=1 git clone --depth 1 https://github.com/YinchenChang/tokenomics /tmp/tk`（驗證 master `45ce1ca`、CURRENT v21.2）。

## 步驟（每步 commit＋push＋進度檔）
1. **快照 v21.2**：TK_Link 重建；`model/20261024_Anthropic_v0.1.5.xlsx`（v0.1.4 檔 `git mv` 到 archive、更新 CURRENT）。列出 10 個 DIFF 名稱新舊值與在本模型的使用位置。
2. **對照文字更新**：a4 第 562 列附近 `c2_tkled` 說明「契約 v18.0：約 6.26」等寫死的舊倍數描述改為不寫死數值（或改讀 TK 名稱），並註明「v21.2 起 TK 計價倍數 1.8，token 帳約 3.0」；η 對照頁「各 TK 版本的 η 與殘差」小表加一欄 v21.2。η 本身不得改變。
3. 重算（LibreOffice 錯誤 0、`CHK_Errors`＝0）；parity 期望值有意識地更新（只限上述 10 名稱與其對照列）；pytest 依檔分開全過；逐格比較 v0.1.4 → v0.1.5，差異只能在 TK_Link 的 10 個值、對照列、說明文字與版本字串。
4. HTML 與 dist：`dist/20261024_Anthropic收支模型_v0_1_5.html`／`.xlsx`（v0.1.4 dist 檔 git rm）；前面加一句 v0.1.5 摘要；「資料與版本」改 TK v21.2；Playwright 逐節讀畫面（桌機與手機寬度）。
5. 報告 `docs/reports/20261024_v0.1.5.md`：①TK 名稱新舊值與使用位置；②**TK 變動造成的差異拆解**（預期：本模型主要輸出差 0；對照列逐項列新舊值）；③逐格比較表 2（範圍外差異，應為 0）。CHANGELOG、交接檔、進度檔；PR 留言 `[Anthropic 回報] v0.1.5｜完成｜日期`。

🤖 Generated with [Claude Code](https://claude.com/claude-code)
