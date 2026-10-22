# Anthropic v0.1.3：共同說明（2026-10-22，chat 端）

- 來由：Andy 2026-10-22 交來外部驗收報告 `docs/reports/20261022_外部驗收報告_v0.1.2.md`（判定「有條件通過」）；chat 端評估後採納其中 5 項（見下表），另加 v0.1.1 時 Andy 同意預設的 U2 改法，並把 Tokenomics 快照版本標記換成 v19.0（chat 端已比對：本模型讀取的 103 個值與 v18.0 完全相同）。
- 執行者：Claude Code 雲端工作階段（Andy 貼指令啟動）。分兩段（A、B），各約 1.5–2 小時；之後一段獨立核對（C）。
- 分支 `claude/anthropic-v0.1.3`（自 main `1079d5c`）；A、B 依序在同一分支；完成後開一個 PR、轉 ready；**不得合併**（Project 規則：合併只在 Andy 明確指示時做）。
- 必讀：`anthropic/CLAUDE.md`、`anthropic/docs/workorders/20261008_anthropic_共同規則.md`、規格 `anthropic/docs/plan/20261008_Anthropic_v0.1_規格.md`、交接檔 `anthropic/docs/handoff/Anthropic_handoff.md`、v0.1.1／v0.1.2 報告、本檔、該段工作單。
- 規則摘要：Excel 是唯一計算引擎；新增列一律加在各頁**尾端**（HTML 產生器寫死既有列編號）；公式不含常數；穩定 ID 只增不重用；不以公司目標反推參數；每步 commit＋push＋進度檔；pytest 依檔分開執行（單次全套會因記憶體被終止）；超過 2 分鐘的命令背景執行，但等它跑完再回報；不得在任何對外請求（含 HTTP User-Agent）中使用使用者 email；不得合併、force push、改 main。
- Tokenomics：`GIT_LFS_SKIP_SMUDGE=1 git clone --depth 1 https://github.com/YinchenChang/tokenomics /tmp/tk`（唯讀；master `1916a06`，CURRENT `20261022_Tokenomics_v19.0.xlsx`）。OpenAI v0.6：`git show origin/main:openai/model/20261008_OpenAI_v0.6.xlsx`（確認 `openai/model/CURRENT`）。

## 外部驗收報告：採納與不採納
| 報告項 | 處理 | 段 |
|---|---|---|
| Priority 4：資金命題拆三層（營運自由現金流能否自給／已到位股權是否足夠／含條件式與談判中） | 採納 | A |
| Priority 3：持續擴張情境（2029–2030 每年追加算力，分「無需求／部分變現／完全變現」） | 採納（只作情境，基準仍只計已簽合約） | A |
| 結論措辭改條件式（「按目前已到位融資與基準支出路徑，不出現最低現金缺口」，不寫「不需要再融資」） | 採納 | A |
| Priority 2：拆分 η | 部分採納：以可得資料拆成「計價口徑（合約價 ÷ TK 持有成本）」「利用率」「工作負載組合（層級與輸入輸出）」與**未解釋殘差**，並列 TK `L1_Gap*`（OpenAI 口徑）作 Analogy；不得因拆分改變 η 的值（只作診斷） | B |
| Priority 1：需求 → token → GW → 成本橋接頁 | 部分採納：新增唯讀 Bridge 頁，逐層列公式、單位、基準／低／高、標記與本模型 vs TK 對照差額；**不以「與 TK 逐項閉合」為驗收條件**（6.26 倍 token 帳是 Tokenomics 端的待辦） | B |
| Priority 5：公開 S-1 後的替換程序 | 採納：寫進 CLAUDE.md（新增列、舊列標「已被取代」並保留、差異比例與受影響輸出入報告） | B |
| 情境機率 50／30／20 | 不採納（無依據的 Assumed；模型不輸出機率） | — |
| 以 TK 閉合作為「通過」門檻 | 不採納（屬 Tokenomics 範圍） | — |
