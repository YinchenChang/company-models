# OpenAI 收支模型（company-models/openai）

OpenAI 收支模型（v0.6，建置中）。2026-10-07 由 `YinchenChang/openai-model`（`main@13c16f0`，停用）搬入本資料夾；以下指令都在 `openai/` 內執行。交接檔：`docs/handoff/OpenAI_handoff.md`；進度：`docs/reports/20261007_openai_進度.md`。命題：OpenAI 每 VR 等值 GW 的營收能否覆蓋每 GW 全成本；若不能，缺口由誰、以什麼條件融資。FY2025–FY2030，曆年制。

- Excel 為唯一計算引擎：`model/CURRENT` 指向現行檔。
- 工作規範見 `CLAUDE.md`；工作單見 `docs/workorders/`；每包報告見 `docs/reports/`；版本紀錄見 `CHANGELOG.md`。
- 本地建置：`pip install -r requirements-dev.txt`；`python3 builder/build.py --tk-dir <Tokenomics> --out model/<檔名>.xlsx --base model/<現行>.xlsx`；測試 `python3 -m pytest tests/parity`。
- 需要 LibreOffice Calc（Ubuntu：`apt-get install -y --no-install-recommends libreoffice-calc-nogui`）。
