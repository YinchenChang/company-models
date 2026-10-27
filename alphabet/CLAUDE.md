# Alphabet 個股模型 — Claude Code 工作守則（v0.3 起）

- 架構：命題模組（Thesis）＋官方營收分部（Seg）→ 損益表（PL）→ CF → Val（固定規格 `docs/plan/個股模型原則_命題加營收分部.md`）。說明見 `README.md`，本版報告見 `docs/reports/20261027_alphabet_v0.3_分部架構.md`。
- 唯一來源：`build_alphabet.py`（全部活公式）；成品以 `bash alphabet/make_v03.sh` 產生（建置 → LibreOffice 重算 → `check_v03.py` → segment_check → `gen_html_json.py` → HTML 檢核）。改了任何輸入或公式，用 `SENS=1` 重跑敏感度，否則 Sens 頁「基準差」會不為 0、`check_v03.py` 失敗。
- `html.json` 由 `gen_html_json.py` 產生，不要手改；`dist/` 只放一個 xlsx 與同名 html，舊版移 `dist/_待刪除/`。
- 舊 MAG 共用引擎（v0.2.2）在 `legacy_v0.2/`（含其 CLAUDE.md），只供參照，不再建置。
- 資料：`data/alphabet_facts_20261008.json`（SEC 一手）為實績來源；SEC 網站在雲端會被擋，新資料用公司 IR 新聞稿並寫來源與日期。
- 下一步：2026Q3 財報（2026-10-28）後以 v0.3.1 滾動（Q 頁加季、FY2026 改 3Q 實績＋Q4 模型、評價日 2026-09-30）。
