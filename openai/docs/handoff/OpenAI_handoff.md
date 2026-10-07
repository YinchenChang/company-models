# OpenAI 收支模型 交接檔（權威版；2026-10-07 起放 repo）

> 取代 Project「算力收支」的 `20261004_OpenAI_v0.6_handoff.md`（已刪除）。新對話串：先讀本檔與進度檔 `docs/reports/20261007_openai_進度.md`，再以 git 確認狀態。

## 1. 命題與範圍
- 命題（Andy 2026-09-30）：OpenAI 每 VR 等值 GW 的營收能否覆蓋每 GW 全成本；若不能，缺口由誰、以什麼條件融資。
- 期間 FY2025–FY2030，曆年制；2025 為實際校準年。上層目的：在下游 AI 變現尚未驗證時，檢驗重度、舉債支撐的中上游投資能否維持 OpenAI 成長。
- 分層：算力物理與產業資料取自 Tokenomics（`YinchenChang/Tokenomics` master `model/CURRENT`，以 TK_Link 快照）；公司財務原始數據在 `SRC_OAI`；假設在 `Inputs`。

## 2. 位置
- repo `YinchenChang/company-models`，資料夾 `openai/`（2026-10-07 由 `YinchenChang/openai-model` 搬入，來源 `main@13c16f0`；舊 repo 停用，Andy 可封存）。
- 規格：`docs/workorders/20261004_openai_v0.6.md`（r6；第 2 節 V1–V16 為全部已定決定，第 5–9 節為 P1 審查紀錄）。
- 分段工作單：`docs/workorders/20261007_openai_*`（共同規則＋S1–S6）。
- 成品：`dist/`（v0.6：HTML 一頁摘要＋xlsx）；Drive `02-算力收支/OpenAI/`（chat 端放置）。

## 3. 流程（2026-10-07 起）
- 自主建置：chat 端寫工作單 → 建置代理執行（每步 commit＋push、進度檔）→ chat 端審查並合併；判斷類套用工作單預設，彙總給 Andy 在里程碑（S5＝v0.6）審查。
- 停止條件見共同規則第 6 節。

## 4. 模型狀態：**v0.6 完成**（2026-10-08；S1–S6）
| 項目 | 現況 |
|---|---|
| Excel | `model/20261008_OpenAI_v0.6.xlsx`（`model/CURRENT`）；成品副本 `dist/20261008_OpenAI收支模型_v0_6.xlsx` |
| HTML 一頁摘要 | `dist/20261008_OpenAI收支模型_v0_6.html`（產生器 `tools/build_html.py`；敏感度情境 `tools/html_scenarios.yaml`；測試 `tests/test_html.py`） |
| 命題與驅動對照表 | `docs/plan/20261008_OpenAI_命題與驅動對照表.md` |
| Tokenomics | v5.26，master `3dd1216688e3872fcb80b28c279d8976a07df29b`，CURRENT `20261007_Tokenomics_v5.26.xlsx`（SHA-256 `abb8592c6921738d20d1c870bfb44dbc0afcb6a412dbb429f4bd24013224ed0d`）；63 名，check_tk_snapshot OK 63 |
| 規模 | SRC_OAI 131；Inputs 263；具名範圍 646；公式格 3,063；Checks C01–C130（CHK_Errors＝0）；parity 107 項 |
| 頁 | README、SRC_OAI、TK_Link、Inputs、Derived_V9、Demand、Revenue、Compute、Cost、Funding、Reverse、Checks、Map_v05、_Defaults |

**主要結果（基準）**：每 VR 等值 GW 差額（含股權報酬、現金口徑）2025–2030 −66.4／−42.1／−22.3／−16.7／−12.7／−10.1，每年皆為負；2030 營收 8.5 對全成本 18.5 $B/GW/年（覆蓋 46%）。已到位融資 87 撐到 2026 年底；2027–2030 累計外部資金需求 442.8（峰值 2029 年 137.6）；Amazon 條件式計入 407.8；反向模式（管理層目標達成、同一支出）92.7。三個關鍵驅動：API 需求成長（2030 淨營收 92.6–188.3）、自有資本支出（累計 217.8–555.3）、員工人數成長（399.4–509.4）。

**PR 列表**（請依序以 merge commit 合併；main 已含 #7、#8、#10、#14）
| PR | 段 | 分支 | 狀態 |
|---|---|---|---|
| #7 | S0 搬遷 | `claude/openai-s0-move` | 已合併 |
| #8 | S1 TK 快照 v5.24 | `claude/openai-s1-tk-snapshot` | 已合併 |
| #10 | S2 需求與營收 | `claude/openai-s2-p2` | 已合併 |
| #14 | S3 算力 | `claude/openai-s3-p3` | 已合併 |
| #15 | S4 成本（v0.6-P4） | `claude/openai-s4-p4` | 待合併 |
| #19 | S5 融資 → v0.6 | `claude/openai-s5-p5` | 待合併（依賴 #15） |
| #20 | S6 成品 | `claude/openai-s6-release` | 待合併（依賴 #15、#19） |

**歷程**：v0.6-P1 `13c16f0`（舊 repo）→ S1 TK v5.24（`bdb0de7`）→ S2 Demand／Revenue → S3 Compute → S4 Cost（TK v5.25 `97e7b20`）→ S5 Funding／Reverse（TK v5.26 `3dd1216`）＝v0.6 → S6 成品。各段報告 `docs/reports/20261007_v0.6-P1.1.md`、`-P2`、`-P3`、`-P4`、`20261008_v0.6-P5.md`（v0.6 完成報告，S1–S5 預設彙總 99 項）、`20261008_v0.6_成品.md`。
- v5.23 X10 已知影響（Tokenomics 串通知）與 v5.24 訓練輸出上修：已反映於快照；詳見 S1、S4 報告。

## 5. 已定決定摘要（詳見 r6 第 2 節）
V1 時變世代組合；V2 η 以 2025 推論支出校準；V3 算力成本＝合約實付（TK 持有成本只作參考，差額＝雲端毛利）；V4 非算力研發＝研發總額 − 2025 訓練支出，之後人數 × 每人成本，股權報酬單列；V5 公司原始數據在 SRC_OAI；V6 訓練支出 12.0＝10.59＋1.41（0–2.91）；V7 複合標記拆開；V8 合約價單一參數 12（8–20）；V9 跨年拆開；V10／V14 來源四欄規則；V11 付費／免費產能本模型自算；V12／V15 pro 占比公式；V13 Cerebras／Azure 期間；V16 Azure 起點 2025（2025–2026）。

## 6. 2026-10-07 預設（Andy 授權 Claude 依建議）
- P2 價格事件時點：公告／生效日為基準，±1 季區間，天數加權。
- P3 η 路徑：基準沿用 2025 值；情境線性回升到 2030＝1。
- P4 人數等：先搜公開來源；搜不到暫用 v0.5 營收占比並標明。
- S5（v0.6）：最低現金沿用 v0.5（10）；Amazon 條件式 $35B 不計入基準、列情境；Nvidia $30B 現金比例 1；反向模式只作對照、不回饋基準（反向資金用同一支出）。
- 其餘見各段報告「已套用的預設」；S1–S5 全部彙總見 `docs/reports/20261008_v0.6-P5.md`。

## 7. Andy 待審預設（v0.6 里程碑；全表 99 項見 `docs/reports/20261008_v0.6-P5.md` 與 `_對照.xlsx` 第 ④ 頁）
1. 研發 GW 為殘差（S3 #15）：2027–2030 占總需求約 6–9 成；替代為 TK `IF_AllocRDGW` 或計畫算力。
2. 自有資本支出 20→70／年（v0.5 Assumed）與自建 GW 計入分母（S4 #2）：現金面最大驅動（歸零時累計 217.8）。
3. API 需求成長與價格彈性；2026 成長率 134% 不重校準（S2 #13）：2030 淨營收 92.6–188.3。
4. Microsoft 分成在現金流扣除（S5 #5）、2025 算力成本用合約實付 14.76 而非實際 20.4（S4 #7）。
5. 員工人數成長路徑（S4 #16）與股權報酬自費用扣出單列（S4 #13）。

## 8. 待辦（v0.7 候選；皆不在 v0.6 範圍）
- 營收上行與算力容量連動（有利情境 2030 容量上限係數 0.50，營收被截住）：「營收翻倍需要多少新增 GW 與成本」情境。
- 融資條件：外部資金分股權／債務、利息與稀釋；或有負債轉實際負債的觸發條件；2026-03 輪實際到位時程。
- 反向模式加「達成目標所需推論 GW」對照供給。
- 研發 GW 外部錨點（請 Tokenomics 端提供）；v5.26 DC_Cost 新名稱可否取代 `IF_CapexTotal`；自研加速器持有成本 IF 名稱（Tokenomics 缺口，可選）。
- S-1 公開後補：現金、承諾、股權報酬、人數、Microsoft 分成在財報中的位置。
- 競爭（Anthropic 等）與價格戰情境（未建）。
- 已修正（chat 端 2026-10-08）：P5 報告「有利組合」2030 營收淨額由 142.1（未截頂 263.9、係數 0.54）更正為 Excel 值 138.1（273.95、0.504）。
- 成品放 Drive `02-算力收支/OpenAI/`（chat 端）。

