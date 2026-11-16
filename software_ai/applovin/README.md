# AppLovin（APP）個股模型

數位廣告科技，原型 A（按廣告計價）。命題模組＝安裝量 × 單次安裝淨收入的單位經濟，驅動單一官方分部 Advertising，再到損益表、現金流、DCF（FCFF 到 FY2032）與同業前瞻本益比，目標價＝兩腿各半加權，三情境 A／B／C 權重 25／50／25。

**版本表（由建置成品產生，不手寫；原則第 10 節第 51 條）：**

<!-- version-table: dist/20261117_AppLovin個股模型_v0.1.xlsx -->
<!-- /version-table -->

**v0.1（2026-11-17）：** 工作單 `docs/workorders/20261116_第一批_研究包與單位經濟.md`（G、G2、G3）。建置 `bash software_ai/applovin/make_v01.sh`（`KREQ=1 SENS=1` 重跑反解與敏感度）。

- **本版未做先驗**（工作單 G：2026-11-16 指示先建模、先驗略過）；先驗到位後改版。
- **不適用：** Meta 式自建算力、Tokenomics 自用 AI 讀法 2、Wattonomics 對照（APP 運算以雲端租用為主，推論價格只讀 Tokenomics 快照）；三層協調（Coord 頁、C0、兩讀數）：不適用：非 Siliconomics 覆蓋。
- **單位經濟模型：** `software_ai/unit_economics/` 尚未完成，變數先在 In 頁自建，標「v0.2 改讀 unit_economics」。
- **研究包：** `data/research_applovin.md`、`data/facts_applovin.csv`（228 筆逐列紀錄）、`segments.json`。
- **Tokenomics：** `data/tokenomics_snapshot.json`（v22.0，commit 79f6b30，只讀）；公司未揭露 token 量與任務量（未找到），每任務 token 以快照倍數推得並標 Assumed；token 價格層級標 Analogy。
- **檔案：** `build_applovin.py`（建置）、`sens_v01.py`、`kreq_v01.py`（整本重算的敏感度與反解）、`check_v01.py`（成品核對）、`gen_html_json.py`（產生 `html.json`）、`display_check.json`（展示層一致性設定）、`dist/`（成品 xlsx 與 html）。
- **公開 repo：** 不含任何個人部位、基金或投資可行性資訊；評等為模型自身的估值評等。
