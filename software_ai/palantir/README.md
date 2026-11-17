# Palantir（PLTR）個股模型

資料與 AI 決策平台；原型 U（按用量計價）。官方分部：Government、Commercial（Commercial 之下子列：美國商業、國際商業）。命題：AI 對美國商業營收的淨效果＝AI 增量（客戶數 × 工作負載成長）減單價侵蝕減推論成本；推論成本只讀 Tokenomics 快照 v22.0。

## 版本

<!-- version-table: dist/_待刪除/20261117_Palantir個股模型_v0.2.1.xlsx｜dist/20261117_Palantir個股模型_v0.2.2.xlsx -->
| 項目 | v0.2 | v0.2.1 |
|---|---|---|
| 現價 | 167.99 | 167.99 |
| DCF 目標價 | 39.52 | 39.52 |
| Comparable 目標價 | 174.76 | 126.42 |
| 加權目標價（上檔空間） | 107.14（-36.2%） | 82.97（-50.6%） |
| 投資建議 | 放空 | 強力放空 |
| 收支關卡 A／B／C（FY2026–FY2030） | 52,436.8／43,897.7／34,925.1 | 52,436.8／43,897.7／34,925.1 |
<!-- /version-table -->

## 檔案

| 路徑 | 內容 |
|---|---|
| `data/research_palantir.md`、`data/facts_palantir.csv`、`data/facts_palantir.py` | 研究包與事實總帳（逐筆：指標、值、單位、期間、口徑、來源、擷取日、標記、來源誘因） |
| `data/tokenomics_snapshot_v22.0.json`、`data/tk_names.txt` | Tokenomics 快照與讀取名稱清單 |
| `build_palantir.py` | 建置腳本（openpyxl 活公式）；`segments.json`、`html.json`、`display_check.json` 為檢查與展示設定 |
| `sens_pltr.py`、`reverse_pltr.py`、`data/sens_v02_*.json`、`data/reverse_solve.json` | 敏感度（一次改一個輸入整本重算）與主導變數反解 |
| `make_v02.sh`、`check_pltr.py` | 一次跑完建置與全部核對 |
| `dist/` | 成品 xlsx 與 html |

## 方法

三層：事實（Q）→ 命題模組與官方分部（Seg）→ 損益（PL）→ 現金流（CF）→ DCF 與 Comparable（Val）。DCF 至 FY2032，FCFF 期中折現，終值 RONIC＝WACC；Comparable＝8 家同業合併中位 P/S（peer_ps_20261113.json all 群組，現價日 2026-11-13 口徑；非 11-13 者標逾期；資料庫最近一季中位列 Sens 案）× 整體 geo_base × 合併目標時點 TTM 營收（v0.2.1：整體單一口徑，因分組讀數把政府分部按 Palantir 自己的 P/S 計價、政府同業與其業務不同類，比值失真，故分組讀數只列對照、不進加權；G6：與 PS／PSG 同口徑，EV／營收與盈餘型倍數只作對照）（software_ai/data/relmult_hist/out/palantir.json 的 geo_base，Derived；θ 成長調整只留作 Sens）；另並列 DCF（股票薪酬加回）對照讀數，不進加權目標價；目標價＝二者加權（預設各半），情境 A／B／C 權重 25%／50%／25%。五級評等與收支關卡照原則第 24、63 條。情境 A／B／C 的差別在 AI 淨效果（增量、侵蝕、推論成本）與政府成長，依據見 README 頁「情境分化依據」。

命題模組變數（客戶數、工作負載、單價、推論成本）暫建在 In 頁，標「v0.2 改讀 unit_economics」。

## 不適用與限制

- 不適用：Coord、三層基準協調、Wattonomics、自用算力（輕資產，資本支出占營收不到 1%）。
- 先驗未做（工作單 G 節：先建模型，先驗後補）。
- 公司未揭露 token 用量、AIP 任務數、推論成本（未找到）；每任務 token 以快照單代理倍數為類比（Analogy），一旦公司揭露即改用揭露推得值（G2-1）。
- 同業倍數非同一估值日（2026-10-26 至 2026-11-15）；賣方共識為 Q3 財報前（標逾期）。
- 公開 repo：不含任何個人部位或基金情境。

## 重跑

`bash software_ai/palantir/make_v02.sh`；`SENS=1 REV=1` 重跑敏感度與反解；需 python3、openpyxl、bs4、LibreOffice。
