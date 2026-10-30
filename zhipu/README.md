# 智譜（智譜華章，02513.HK）個股模型

智譜每 VR 等值 GW 的年營收能否覆蓋每 GW 年全成本；若不能，缺口由誰、以什麼條件融資（FY2025–FY2030）。架構比照 `openai/`（v0.6），算力物理取自 Tokenomics；金額單位人民幣億元，命題另列美元口徑與 OpenAI 並排。

- 現行版本：**v0.1.2**（`model/20261030_Zhipu_v0.1.2.xlsx`；見 `model/CURRENT`）；成品 `dist/20261030_智譜個股模型_v0_1_2.html／.xlsx`（依 `docs/plan/個股模型原則_命題加營收分部.md` 第 1 節改稱「個股模型」）。
- 官方報告分部（部署方式）：本地化部署（②企業級智能體＋③企業級通用大模型＋④技術服務及其他）、雲端部署（①開放平台及 API）。宣告檔 `segments.json`；對帳在 Revenue 第七節（FY2025 對年報、1H2026 對中期公告子項加總，容差 Inputs `seg_recon_tol`），Checks C100–C103；`python3 tools/segment_check/check_segments.py zhipu`（repo 根目錄執行）須為「符合」。
- 規格：`docs/plan/20261008_Zhipu_v0.1_規格.md`
- 交接：`docs/handoff/Zhipu_handoff.md`；進度：`docs/reports/20261008_zhipu_進度.md`（v0.1.2：`docs/reports/20261030_v0.1.2_進度.md`）
- 工作規範：`CLAUDE.md`
