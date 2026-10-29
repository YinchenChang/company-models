# OpenAI 收支模型（company-models/openai）

OpenAI 收支模型（v0.6，建置中）。2026-10-07 由 `YinchenChang/openai-model`（`main@13c16f0`，停用）搬入本資料夾；以下指令都在 `openai/` 內執行。交接檔：`docs/handoff/OpenAI_handoff.md`；進度：`docs/reports/20261007_openai_進度.md`。命題：OpenAI 每 VR 等值 GW 的營收能否覆蓋每 GW 全成本；若不能，缺口由誰、以什麼條件融資。FY2025–FY2030，曆年制。

- Excel 為唯一計算引擎：`model/CURRENT` 指向現行檔。
- 工作規範見 `CLAUDE.md`；工作單見 `docs/workorders/`；每包報告見 `docs/reports/`；版本紀錄見 `CHANGELOG.md`。
- 本地建置：`pip install -r requirements-dev.txt`；`python3 builder/build.py --tk-dir <Tokenomics> --out model/<檔名>.xlsx --base model/<現行>.xlsx`；測試 `python3 -m pytest tests/parity`。
- 需要 LibreOffice Calc（Ubuntu：`apt-get install -y --no-install-recommends libreoffice-calc-nogui`）。

## v0.6.4：口徑一 Q2 GW 採 B／C 段、D2 推導式、衍生值移 Inputs（Tokenomics v21.7）
- 口徑一 Q2 季平均 GW 2.26（2.17–2.35）；`PxQ` 新增第 8.4 節（D2 推導式，候選因子）與第九節（Microsoft 10-Q 對照）；SRC_OAI_141–143 退役、改為 Inputs 公式列；`TK_Link` 改取 v21.7。建置：`python3 builder/build.py --tk-dir <Tokenomics checkout@8309fdb> --out model/20261029_OpenAI_v0.6.4.xlsx --base model/archive/20261025_OpenAI_v0.6.3.xlsx --date 2026-10-29 --tk-date 2026-10-29`。範圍外變動比較：`python3 tools/diff_v064.py <v0.6.3> <v0.6.4>`。報告 `docs/reports/20261029_v0.6.4.md`。

## v0.6.3：PxQ 比值 2026 三口徑（Tokenomics v21.4）
- `PxQ` 頁 11 欄：2025 指數／線性 + 2026 口徑一（Q2 年化營收÷Q2 季平均 GW，主）／口徑二（FY2026 營收÷全年平均 GW）／口徑三（9 月底 run-rate÷年底 GW）× 低／基準／高；新增 Fairwater 六情境（第七節）與 GW 路線 token 帳（第八節）。建置：`python3 builder/build.py --tk-dir <Tokenomics checkout> --out model/20261025_OpenAI_v0.6.3.xlsx --base model/archive/20261022_OpenAI_v0.6.2.xlsx --date 2026-10-25 --tk-date 2026-10-25`。報告 `docs/reports/20261025_v0.6.3.md`。

## v0.6.2：Tokenomics 同步 v19.0（單一快照）
- `TK_Link` 改取 Tokenomics master `model/CURRENT`（v19.0）；`TK_PxQ`／`TKQ_` 退役。建置只需一份 Tokenomics：`python3 builder/build.py --tk-dir <Tokenomics checkout> --out model/20261022_OpenAI_v0.6.2.xlsx --base model/archive/20261022_OpenAI_v0.6.1.xlsx --date 2026-10-22`。報告 `docs/reports/20261022_v0.6.2_TK同步.md`；歸因 `tools/tk_attribution.py`。（以下 v0.6.1 的 `TK_PxQ`／`--tk-pxq-dir` 說明為歷史紀錄。）

## v0.6.1：P×Q 與回本比值
- 新增 `PxQ` 頁與 `TK_PxQ`（Tokenomics v19.0 第二快照，具名範圍 `TKQ_`）；`TK_Link` 仍為 v5.26（63 名不變）。建置需兩份 Tokenomics：`python3 builder/build.py --tk-dir <v5.26 checkout> --tk-pxq-dir <v19.0 checkout> --out model/20261022_OpenAI_v0.6.1.xlsx --base model/archive/20261008_OpenAI_v0.6.xlsx --date 2026-10-22 --tk-date 2026-10-08`。
- 主輸出：比值＝實際 m×L ÷ 所需 m×L＝營收 ÷ 算力成本，2025（指數／線性內插）與 2026（2.5／3.0／3.5 GW）。規格（Andy 原話）見 `docs/handoff/OpenAI_handoff.md` 第 9 節；測試把關：`tests/parity/test_builder.py` 的 PxQ 引用白名單與 E6。
- 逐格比對工具：`python3 tools/diff_workbooks.py <舊版> <新版>`；對照 Excel：`tools/make_pxq_report.py`。
