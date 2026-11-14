# Alphabet 個股模型 — Claude Code 工作守則（v0.3 起）

- 架構：命題模組（Thesis）＋官方營收分部（Seg）→ 損益表（PL）→ CF → Val（固定規格 `docs/plan/個股模型原則_命題加營收分部.md`）。說明見 `README.md`，v0.3.5 報告見 `docs/reports/20261104_alphabet_v0.3.5_10Q補入.md`（v0.3.9 說明見 README「v0.3.9」節）（季度滾動 `docs/reports/20261029_alphabet_v0.3.1_季度滾動.md`）（v0.3 架構報告 `docs/reports/20261027_alphabet_v0.3_分部架構.md`）。
- 唯一來源：`build_alphabet.py`（全部活公式）；成品以 `bash alphabet/make_v03.sh` 產生（建置 → LibreOffice 重算 → `check_v03.py` → segment_check → `gen_html_json.py` → HTML 檢核）。改了任何輸入或公式，用 `SENS=1` 重跑敏感度，否則 Sens 頁「基準差」會不為 0、`check_v03.py` 失敗。
- `html.json` 由 `gen_html_json.py` 產生，不要手改；`dist/` 只放一個 xlsx 與同名 html，舊版移 `dist/_待刪除/`。
- 舊 MAG 共用引擎（v0.2.2）在 `legacy_v0.2/`（含其 CLAUDE.md），只供參照，不再建置。
- 資料：`data/alphabet_facts_20261008.json`（SEC 一手）＋`data/alphabet_facts_2026q3.json`（2026Q3 新聞稿＋v0.3.5 起 10-Q 摘錄，`facts_2026q3.py` 兩讀核對）為實績來源；SEC 網站在雲端會被擋，新資料用公司 IR 新聞稿並寫來源與日期。
- 改了資本支出相關輸入：`CALIB=1` 重解 C10 速度（`calib_pace_v031.py`）。
- v0.3.4（2026-11-04）Tokenomics 快照 v21.9；v0.3.5（2026-11-04）已用 2026Q3 10-Q 補入債務面額、在建工程、租賃、權益證券（SpaceX）、罰款、RPO（報告 `docs/reports/20261104_alphabet_v0.3.5_10Q補入.md`；拆解鏈舊值 `data/q10_v035_old.json`）。v0.3.6（2026-11-04）Tokenomics 快照 v21.10，基準數值不變，只有讀 `IF_NonNV_Out_TPU` 的 Sens 案例改變（報告 `docs/reports/20261104_alphabet_v0.3.6_TK快照v21.10.md`）。
- v0.3.9（2026-11-14）版面 v1.3＋RevDCF 頁＋Coord 頁（三層協調；追公式先於 C0 比較，詳見 README v0.3.9 節）；後處理在 `v039_post.py`（只新增、不改既有格），HTML 在 `gen_html_json.py` 的 `_v13()`。
- 下一步：2026Q4 財報（2027-02）以同一方式滾動（Q 頁加季、FY2026 全年實績、評價日 2026-12-31；2027 資本支出指引 → C10）；同時把 Q4 26 起融資租賃本金換成最新 10-Q／10-K 到期表（v0.3.5 報告第 5 節第 2 題）。滾動時 `q10|*` 拆解鏈與 `q10_v035_old.json` 可移除或換成新一版的舊值。
