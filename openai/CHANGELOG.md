# CHANGELOG

每次同步記錄：Excel 版本、工作包、commit 與變動摘要。

## 20261007_OpenAI_v0.6.xlsx — v0.6-P1.1（S1：TK_Link 快照更新至 Tokenomics v5.24）

- 快照來源：Tokenomics master `bdb0de7`（CURRENT＝`20261007_Tokenomics_v5.24.xlsx`）。check_tk_snapshot OK 37→62、MISSING 11→0、DIFF 0。原待合併 11 名全部取得；新增 IF_TrainCost_*、IF_RevGW_*、IF_FullCostDefault_*、L1_Ans3（＋低高）、SRC_DEM_013_Lo／Hi。原 37 名中 10 名值變動（TokGW×3、HoldEcon、FullCost×3、ProgGWyr×3）；第 5 世代口徑 Rubin Ultra NVL576→MGX NVL 單架 72 GPU。
- 只動 TK_Link、Checks（C07、C08 標籤與期望值；ID 不變）、README 兩格；具名範圍 401→426；SRC_OAI、Inputs、Derived_V9、Map_v05 逐格不變。舊檔移至 `model/archive/`。
- 工具：`tools/diff_tk_snapshots.py` 新增；`builder/tk_link.py`、`build.py` 沿用舊 PR #2（`479e94b`）改動。parity 32 項全過。報告：`docs/reports/20261007_v0.6-P1.1.md`、`_對照.xlsx`。PR #8。

## S0 搬遷（2026-10-07，Excel 不變）

- 由 `YinchenChang/openai-model` `main@13c16f0` 搬入 `company-models/openai/`；CI 移至根目錄 `.github/workflows/openai-parity.yml`（working-directory `openai`）。新增共同規則、S1–S6 工作單、進度檔、交接檔（`docs/handoff/OpenAI_handoff.md`，取代 Project 內舊交接檔）。parity 32 項全過，Excel 逐位元組相同。

## 20261004_OpenAI_v0.6.xlsx — v0.6-P1（repo、Source 與 Tokenomics 連結）

- 報告：`docs/reports/20261004_v0.6-P1.md`。Commit：見本 PR 的合併提交（合併後補上雜湊）。
- 新增：repo 骨架（builder／engine／tests／tools／CI）、CLAUDE.md；Excel 8 頁：README、SRC_OAI（106 列）、TK_Link（Tokenomics v5.14 快照，37 個名稱有值、11 個待 v5.15 合併）、Inputs（233 列）、Derived_V9、Checks（26 項，編號凍結）、Map_v05（849 個 v0.5 葉節點去處）、_Defaults（隱藏）。穩定 ID registry（`builder/id_registry.json`）。
- 測試：parity 24 項全過；v0.5 每個葉節點有去處；`check_tk_snapshot`：OK 37、MISSING 11（WARN）。
- 依工作單 r2：V6（SRC 10.59、訓練支出改公式、其他雲端 1.41／0–2.91）、V8（合約價單一列 12／8–20）、V9（跨年拆開）、V10（四欄初評）、V11；E4（PR）、E5（環境安全）。詳見報告第 2 節。
- 依工作單 r3（內容依 chat 訊息摘要，r3 原文未在 main）：E6（V9 公式不含常數、換算係數進 Inputs）、E7（帳目四項）、V12（pro 改正）、V13（Cerebras／Azure 終點與總額 Assumed＋區間）。V14、E8 待 r3 原文。詳見報告。
- 依工作單 r3 補做與 V15：V14／F2（來源初評規則）、E8 a–j（待判斷 A 類全數處理）、F1（Checks 編號 registry）、V15（pro 2027–2030 改公式）。詳見報告。
- 審查補充（G1–G8、V16）：合約列立場初評（未評 11→5）、Checks C27 WARN、定義常數註記、Azure 起點區間 2025–2026。詳見 `docs/reports/20261004_v0.6-P1_supplement.md`。
- 審查補充二（H1–H4）：SRC_OAI_105–108 補出處（OpenAI 與 NVIDIA／Broadcom／AMD 聯合公告）、「同上」出處追溯、一手／二手圖例、SRC_OAI_082 立場說明。V10：立場未評 1（082）、等級未評 0、一手／二手未評 0。
