# CHANGELOG（Anthropic 收支模型）

每次同步記錄：Excel 版本、段別、commit 與變動摘要。

## A0 骨架（2026-10-08，Excel 尚無）
- 新增 `anthropic/`：CLAUDE.md、規格 r0（D1–D20）、共同規則、A1–A5 工作單、進度檔、交接檔；engine、tools、tests 基礎檔沿用 OpenAI v0.6（`claude/openai-s6-release` @ `e91df57`），待 A2 改寫。

## v0.1-A2 需求與營收（2026-10-08；Excel `model/20261008_Anthropic_v0.1.xlsx`）
- commit：`aa06cb2`（步驟 1 骨架）、`52d8dd5`（2 NonNV）、`7609b7c`（3 Inputs）、`691b46f`（4–6 Demand／Revenue／Checks）、`d9c6284`（7 測試與 CI）、`fd5db5b`（每任務 token 改取 Tokenomics）；報告提交見 PR #22 留言。
- 新建活頁簿 9 頁：README、SRC_ANT（379 列；新增 377–379；Derived 12 列：11 公式、1 保留報導值）、TK_Link（Tokenomics v5.26：63＋5 名＋NonNV 18 格）、OAI_Link（OpenAI v0.6 命題輸出 17 名）、Inputs（105 列；INP_067–069 退役）、Demand、Revenue、Checks（C01–C48）、_Defaults。
- 結果：2025 營收 4.60（校準差 0）；2026 48.73（半校準）；2030 總額 214.9、雲端平台抽成 18.1、淨額 196.8（$B）。
- 工具：`builder/build.py`（分段模組 a2／a3／a4）、`builder/a2.py`、`builder/oai_link.py`；`tools/check_tk_snapshot.py` 支援 NonNV 讀表；`tools/make_report_tables.py` 改寫；`tools/export_csv.py` 改 SRC_ANT；`engine/core.py` 檔名規則改 Anthropic。
- 測試：parity 19 情境、test_builder（E6、穩定 ID、OAI_Link 隔離、Derived 公式化）；54 項全過。CI：`.github/workflows/anthropic-parity.yml`（GitHub 帳號付款／用量上限，job 未啟動）。

## v0.1-A3 算力與成本（2026-10-08；Excel `model/20261008_Anthropic_v0.1.xlsx`，同檔名覆寫）
- 新增 `builder/a3.py`：Compute（加速器族 × 世代 9 族、逐合約供給 GW 10 份＋既有雲端＋自建、η、推論／研發 GW、容量上限、VR 等值、敏感度 A／B／C、對照列與機房租約）、Cost（逐合約實付、自建資本支出、供應商持有成本與雲端毛利、經濟口徑、非算力成本與股權報酬、推論／研發拆分、命題表、TK 單位成本參考、D22 對照、$518B 對帳）。
- Revenue `REV_CapFactor` 改接 Compute `CMP_CapFactor`（2025、2026 固定 1）；Inputs 佔位 INP_016 退役。Inputs 160 列（新增 INP_109–164）。TK_Link 加讀 Tokenomics Inputs 頁 PUE（`TK_PUE`、`_Lo`、`_Hi`；讀表，非具名）；`tools/check_tk_snapshot.py` 通用化讀表比對（OK 89）。
- Checks C49–C82；CHK_Errors＝0、CHK_Warnings＝0。parity 情境 27 個（A3 新增 10 個：合約價、η、產出比、人數成長、自建資本支出、PUE、Broadcom 5 GW＋爬坡、容量截頂、兩個 ERR 情境；移除 A2 的容量上限佔位情境 2 個）。
- 結果（$B；每 VR 等值 GW 為 $B/GW/年）：2030 供給 10.4 GW（VR 等值 7.7）、算力成本 90.8、全成本 110.3、營收淨額 196.8；每 VR 等值 GW 差額 2026 −6.5、2027 −5.2、2028 −2.9、2029 ＋5.1、2030 ＋11.2；覆蓋率 2030 1.79（OpenAI v0.6：−10.1、0.46）。
