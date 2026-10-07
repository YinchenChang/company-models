# CoreWeave × Tokenomics 改造進度（接手用；W1–W3 共用）

## 目前狀態（每次 push 前覆寫）
- 已完成：W0 遷移（crwv-model `01b13ad`／v4.5 原樣複製到 `coreweave/`；本環境 `scripts/verify.sh --vs-dist` 21 項全過、與 v4.5 成品 0 差異）；工作單 W1–W3＋共同規則（chat 端）。
- 下一步：W1——從最新 main 開分支 `claude/coreweave-w1-tokenomics`，依 `coreweave/docs/workorders/20261007_coreweave_W1_Tokenomics取數層.md` 步驟 0 開始。
- 未解問題：Tokenomics v5.25（新增 DC_Cost 構件 IF_ 名稱）由 chat 端另開工作單同時進行。

## 工作單總覽
| 工作單 | 分支 | PR | 狀態 |
|---|---|---|---|
| W0 遷移 | `claude/coreweave-w0-migrate` | | chat 端完成 |
| W1 Tokenomics 取數層 | `claude/coreweave-w1-tokenomics` | | 未開始 |
| W2 每 MW 改寫 | `claude/coreweave-w2-permw` | | 未開始 |
| W3 v4.6 成品與對照 | `claude/coreweave-w3-v4.6` | | 未開始 |

<!-- 各工作單在下方新增自己的段落：「## Wx」＋步驟紀錄表（步驟｜狀態｜commit｜備註） -->
