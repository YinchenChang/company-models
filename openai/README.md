# OpenAI 收支模型（company-models/openai）

OpenAI 收支模型（v0.6，建置中）。2026-10-07 由 `YinchenChang/openai-model`（`main@13c16f0`，停用）搬入本資料夾；以下指令都在 `openai/` 內執行。交接檔：`docs/handoff/OpenAI_handoff.md`；進度：`docs/reports/20261007_openai_進度.md`。命題：OpenAI 每 VR 等值 GW 的營收能否覆蓋每 GW 全成本；若不能，缺口由誰、以什麼條件融資。FY2025–FY2030，曆年制。

- Excel 為唯一計算引擎：`model/CURRENT` 指向現行檔。
- 工作規範見 `CLAUDE.md`；工作單見 `docs/workorders/`；每包報告見 `docs/reports/`；版本紀錄見 `CHANGELOG.md`。
- 本地建置：`pip install -r requirements-dev.txt`；`python3 builder/build.py --tk-dir <Tokenomics> --out model/<檔名>.xlsx --base model/<現行>.xlsx`；測試 `python3 -m pytest tests/parity`。
- 需要 LibreOffice Calc（Ubuntu：`apt-get install -y --no-install-recommends libreoffice-calc-nogui`）。

## v0.6.1：P×Q 與回本比值
- 新增 `PxQ` 頁與 `TK_PxQ`（Tokenomics v19.0 第二快照，具名範圍 `TKQ_`）；`TK_Link` 仍為 v5.26（63 名不變）。建置需兩份 Tokenomics：`python3 builder/build.py --tk-dir <v5.26 checkout> --tk-pxq-dir <v19.0 checkout> --out model/20261022_OpenAI_v0.6.1.xlsx --base model/archive/20261008_OpenAI_v0.6.xlsx --date 2026-10-22 --tk-date 2026-10-08`。
- 主輸出：比值＝實際 m×L ÷ 所需 m×L＝營收 ÷ 算力成本，2025（指數／線性內插）與 2026（2.5／3.0／3.5 GW）。規格（Andy 原話）見 `docs/handoff/OpenAI_handoff.md` 第 9 節；測試把關：`tests/parity/test_builder.py` 的 PxQ 引用白名單與 E6。
- 逐格比對工具：`python3 tools/diff_workbooks.py <舊版> <新版>`；對照 Excel：`tools/make_pxq_report.py`。
