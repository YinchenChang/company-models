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
