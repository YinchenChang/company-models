# AppLovin（APP）個股模型

數位廣告科技，原型 A（按廣告計價）。命題模組＝安裝量 × 單次安裝淨收入的單位經濟，驅動單一官方分部 Advertising，再到損益表、現金流、DCF（FCFF 到 FY2032）與同業 P/S 法（G6 起；前瞻本益比只作對照列、不進目標價），目標價＝DCF 與 Comparable 兩腿各半加權，三情境 A／B／C 權重 25／50／25。

**版本表（由建置成品產生，不手寫；原則第 10 節第 51 條）：**

<!-- version-table: dist/_待刪除/20261117_AppLovin個股模型_v0.1.xlsx｜dist/_待刪除/20261117_AppLovin個股模型_v0.2.xlsx｜dist/20261117_AppLovin個股模型_v0.2.1.xlsx -->
| 項目 | v0.1 | v0.2 | v0.2.1 |
|---|---|---|---|
| 現價 | 409.10 | 409.10 | 409.10 |
| DCF 目標價 | 263.13 | 261.31 | 261.31 |
| Comparable 目標價 | 438.12 | 948.67 | 542.52 |
| 加權目標價（上檔空間） | 350.62（-14.3%） | 604.99（+47.9%） | 401.92（-1.8%） |
| 投資建議 | 中立 | 買進 | 中立 |
| 情境目標價 A／B／C | 426.65／353.92／268.01 | 723.59／611.10／474.17 | 461.94／404.45／336.83 |
| 收支關卡 A／B／C（FY2027–FY2031，US$M） | 59,184.2／48,677.8／36,342.7 | 59,184.2／48,677.8／36,342.7 | 59,184.2／48,677.8／36,342.7 |
<!-- /version-table -->

**v0.2.1（2026-11-17）：** 工作單 G6：Comparable 主腿由前瞻本益比改為 **P/S 口徑**，以與 PS／PSG 相對值同基礎（v0.2 把已含利潤率差異的 PS／PSG 相對值套到本益比，重複計入而高估）。每股價值＝同業中位 P/S（`software_ai/data/relmult_hist/` 資料庫最近一季，季末日與逾期標示見 Val 頁；副本 `data/relmult_ps_latest.csv`）× 相對倍數 geo_base × 目標時點 TTM 營收 ÷ 評價日稀釋股數。**盈餘類倍數（前瞻本益比）只作對照列，不進加權目標價**（Val 頁「前瞻本益比對照」：相對倍數 1.00 與本公司現行相對值兩讀）。Sens 的 sens.low、sens.high、ps_only、psg_only 四案皆落在新 P/S 腿。

**v0.2（2026-11-17）：** 工作單 G5：相對倍數改讀 `software_ai/data/relmult_hist/out/applovin.json` 的 geo_base（副本 `data/relmult_applovin.json`；PS 與 PSG 讀數差異大，成長比 2.60，原因見 README 頁與 HTML）、rf 統一 5.10%（FRED DGS10，2026-11-12）、DCF 並列股票薪酬加回讀數（只作對照）。v0.1 為工作單 G、G2、G3 首版；v0.1 與 v0.2 成品皆移 `dist/_待刪除/`。建置 `bash software_ai/applovin/make_v02.sh`（`KREQ=1 SENS=1` 重跑反解與敏感度）。

- **本版未做先驗**（工作單 G：2026-11-16 指示先建模、先驗略過）；先驗到位後改版。
- **不適用：** Meta 式自建算力、Tokenomics 自用 AI 讀法 2、Wattonomics 對照（APP 運算以雲端租用為主，推論價格只讀 Tokenomics 快照）；三層協調（Coord 頁、C0、兩讀數）：不適用：非 Siliconomics 覆蓋。
- **單位經濟模型：** `software_ai/unit_economics/` 尚未完成，變數先在 In 頁自建，標「v0.2 改讀 unit_economics」。
- **研究包：** `data/research_applovin.md`、`data/facts_applovin.csv`（228 筆逐列紀錄）、`segments.json`。
- **Tokenomics：** `data/tokenomics_snapshot.json`（v22.0，commit 79f6b30，只讀）；公司未揭露 token 量與任務量（未找到），每任務 token 以快照倍數推得並標 Assumed；token 價格層級標 Analogy。
- **檔案：** `build_applovin.py`（建置）、`sens_v02.py`、`kreq_v02.py`（整本重算的敏感度與反解）、`check_v02.py`（成品核對）、`gen_html_json.py`（產生 `html.json`）、`display_check.json`（展示層一致性設定）、`dist/`（成品 xlsx 與 html）。
- **公開 repo：** 不含任何個人部位、基金或投資可行性資訊；評等為模型自身的估值評等。
