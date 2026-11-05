# 獨立核對：E_earnings_20261105（EQ#1–11）

- 日期：2026-11-05
- 核對者：未參與製作的代理
- 分支：`claude/wattonomics-evidence-20261105`，head 9a7930a（已用 ls-remote 確認）
- 方法：逐列重新開啟來源網址，核對內容、引文（≤15 字）、日期、等級與 E12 門檻。
- 不提交。
- 遇到需要授權、而授權沒有回應的頁面，一律標「無法核對」，之後不再重試。

## 判定：**有條件通過**

- 每一列的核心事實都屬實。
- 「不建議寫入基準」的結論成立。
- 需要修正的有 4 處：EQ#1 日期欄、EQ#7 與 EQ#8 的未核實細節、EQ#11 無法核對（說明見下方必要修正）。

## 逐列結果

| # | 結果 | 說明 |
|---|---|---|
| EQ#1 | 屬實；需修正 | 來源 1 是 2025-10-06 的協議新聞稿，原文為 "set to begin in the second half of 2026"。日期欄寫 2026-11-03，應改為「2025-10-06（協議）／2026-11-03（法說）」。來源 2 的 "big step-up in Q4" 屬實，是蘇姿丰回答 Arya 時說的。**否定性陳述「沒有提 OpenAI 1 GW 時程」：屬實。** stockanalysis 逐字稿提到 OpenAI 3 次：Dots、Turin 主機 CPU、Helios 首批上線，都沒有 1 GW、時程或改期；全文也沒有出現 "1 gigawatt"。singjupost 版本相同。財務長 Jean Hu 只泛談認股權證「條件達成後」才計入，沒有點名客戶。另用網路搜尋相關報導，結果多為 2025 年舊文，沒有找到相反的報導。等級 Interested-party 恰當 |
| EQ#2 | 屬實 | 新聞稿（2026-11-03）原文為 "running models and workloads on the first"，列出 Anthropic、Meta、Microsoft、OpenAI、OCI。法說原文為 "expect Helios deployments to increase substantially in the fourth quarter"（蘇姿丰） |
| EQ#3 | 屬實 | "roughly 3 GW…" 是 BofA 分析師 Vivek Arya 提問時說的，蘇姿丰沒有確認數字；"we know where the GPUs are going" 屬實。等級標註恰當 |
| EQ#4 | 屬實 | 新聞稿有 Vultr、HPE Helios 的句子；"second half of 2027" 指 LUMI-AI（MI430X）。兩者都沒有 MW |
| EQ#5 | 屬實；**無法對應站點** | Microsoft 投資人頁的逐字稿（2026-10-28），財務長 Amy Hood 原話："some third-party capacity deliveries we expected in Q1 were delayed"，另有 "3P GPU as a service contracts that have been delayed to H2"。她沒有指名供應商、站點或 MW。來源 2（datastudios，2026-10-31）是媒體的轉述文字，不是 Microsoft 原話，應標「二手轉述」。判斷：公司自揭，達到 E12。但 Hood 的話與 IREN、Nebius 的連結只來自 Business Insider 的推論（tftc 把兩者並列），不能據此寫入 cp_neo_iren 或任何 T2 列的 M 欄。建議只在 Delay 頁列為全公司層級的候選 |
| EQ#6 | 屬實 | 原話（納德拉）："adding 24 new data centers across five continents"，沒有 MW。dock-to-live 縮短約 20% 這一點本次沒有逐字核到（未核） |
| EQ#7 | 部分屬實；**單一來源** | tftc（2026-11-04）原文為 "Business Insider first reported"，tftc 本身沒有寫備援發電機或冷卻設備。phemex（2026-11-04）原文 "still lacks backup generators and cooling equipment" 屬實，但它的出處是 PANews 轉 Business Insider。runtimewire 也把 Business Insider 列為主要來源。**結論：兩個來源都源自同一篇 Business Insider（Aterio 判讀，飛越日 10-21），應視為單一來源。** 「IREN 不予置評」「Microsoft 稱策略保持彈性」「每座 Horizon 50 MW」三點，所開頁面都看不到，應標「未核」。Horizon 1 在 2026-08-13 交付，經 IREN 新聞頁確認屬實。等級 Analogy、未達 E12，判斷正確 |
| EQ#8 | 屬實；**單一來源** | runtimewire（2026-11-04）原文："roughly 140 MW at Nebius's New Jersey project could slip into 2027"；PANews 原文 "may be delayed to 2027"。兩者都明確標示出處為 Business Insider，**是同一篇的轉載**。所開來源只寫 "New Jersey project"，沒有出現 "Vineland"，對應到 Vineland 屬推論，需要註明。「DataOne 不予置評」也未核到。未達 E12，判斷正確 |
| EQ#9 | 屬實 | IREN 新聞頁 10-22 公告 11-05 發布季報；stocktitan 確認法說為 11-05 17:00 ET、季度截至 2026-09-30。核對時還沒有季報。Sweetwater 2 GW（09-08）、西班牙 300 MW（10-21）屬實；Childress 750 MW 只見於站點選單 |
| EQ#10 | 屬實 | FOX13 2026-11-02 原文："expected to become operational by the end of 2026"；共規劃 6 座，沒有 MW。Musk 確認的名稱是第五座的 "Gun Range"；"Spam" 只是依舊址推測的名稱，不是 Musk 確認的，建議改寫。第二來源（09-02，660,000 平方呎）屬實，同屬 FOX13，不構成獨立來源，原檔已註明 |
| EQ#11 | **無法核對** | aiweekly 與 st-hakky 兩頁需要授權，未獲回應。本列與 E_oracle_20261104 E#1、E#2 重複，該檔另有核對紀錄 |

## 重點結論

1. **EQ#5**：Microsoft 財務長的原話屬實，公司自揭，達到 E12；但沒有供應商、站點或 MW，**不能對應任何 T2 站點**。
2. **IREN Childress 與 Nebius New Jersey**：兩列的四個來源全部源自同一篇 Business Insider（Aterio 衛星判讀）。**每一列都應視為單一來源**，公司沒有自揭，未達 E12。
3. **AMD 否定性陳述**：兩份逐字稿都確認法說沒有提 OpenAI 1 GW 的時程，屬實。

## 必要修正

- EQ#1：日期欄改為協議日 2025-10-06／法說日 2026-11-03。
- EQ#5：來源 2 標「二手轉述」，不要當作原話。
- EQ#7：「IREN 不予置評」「Microsoft 保持彈性」「每座 50 MW」標「未核」；來源欄註明「同一 Business Insider，單一來源」。
- EQ#8：Vineland 標為推論；「DataOne 不予置評」標「未核」。
- EQ#10：「Spam」不是 Musk 確認的名稱。
- EQ#11：標「無法核對（本次）」。

## 可寫進 Delay 頁追蹤的列（不入基準、M＝0）

- **EQ#5**：Microsoft 全公司第三方容量延遲。公司自揭，站點未指名。
- **EQ#7**：cp_neo_iren 的 Horizon 4，單一來源 Analogy。
- **EQ#8**：更新 D#19，MW 約 140，單一來源 Analogy。

另外，EQ#9（IREN 11-05 季報）與 CoreWeave（11-10）季報出來後，需要再補查。
