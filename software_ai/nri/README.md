# software_ai/nri/ — 野村總合研究所 NRI（東證 4307）個股模型 v0.2（原型 L：按人力計價）

**v0.2（2026-11-17，G5）：** 相對倍數改取 `software_ai/data/relmult_hist/out/nri.json` 的 geo_base（Derived）；現價更新為 2026-11-13 收盤；無風險利率統一為 2.77%（2026-11-11）；Recon 與 HTML 估值章新增 DCF（股票薪酬加回）對照。**本版未做先驗。** 工作單：`software_ai/docs/workorders/20261116_第一批_研究包與單位經濟.md` G、G2、G3 節。

**建置：** `bash software_ai/nri/make_nri.sh`（`SENS=1` 重跑敏感度與反解）。成品 `dist/20261117_野村總研個股模型_v0.2.xlsx`／`.html`。投資建議、目標價與各數字見成品 README 頁與 HTML（由公式組字，本檔不手寫）。

**命題：** AI 進入 IT 服務後，人力計價營收＝人數 × 利用率 × 每人年工時 × 實現小時單價；AI 淨效果＝AI 新專案增量 − 價格侵蝕（客戶分享生產力）− 推論成本，營業利益再加省下的人力成本。推論成本只讀 Tokenomics v22.0 快照（`data/tokenomics_snapshot_v22.0.json`，市場價）。變數名稱沿用 `software_ai/unit_economics`。

**結構：** README、In、Hist、Q、Thesis、PL、CF、Val、RevDCF、Recon、Sens、Sources。官方報告分部取 2026-04 新口徑（Consulting／Financial IT Solutions／Industrial IT Solutions／IT Platform Services／Others）；產品線取公司固定揭露的產業別外部營收。

**檔案：** `build_nri.py`、`facts_nri.py`（事實總帳，產出 `data/facts_nri.csv`）、`sens_nri.py`、`kreq_nri.py`、`gen_html_json.py`、`segments.json`、`html.json`、`display_check.json`、`make_nri.sh`、`data/`（`research_nri.md`、事實總帳、快照、敏感度與反解結果）、`dist/`。

## 不適用與未做
- 先驗：未做。三層基準協調、Wattonomics 對照：不適用（非 Siliconomics 覆蓋、不自建算力）。
- 分部營業利益新口徑 FY2026 重編值、勞務費拆分、數值利用率、token 與代理任務量、財報後共識、歷史相對倍數：未找到（見研究包第 9 節）。

## 已知限制
- 財報數字經 WebFetch 摘要模型轉錄，非逐字人工核對；現價二來源互證（Traders Web、StockAnalysis；2026-11-13）；同業前瞻 P/E 日期不一致且逾期；rf 單一來源。
- 每任務 token 與價格層級為 Analogy，額外推論倍數與利用率、單價、價格保留率、自動化比例為 Assumed。

<!-- version-table: dist/_待刪除/20261116_野村總研個股模型_v0.1.xlsx｜dist/20261117_野村總研個股模型_v0.2.xlsx -->
| 項目 | v0.1 | v0.2 |
|---|---|---|
| 現價 | 6,056.00 | 6,012.00 |
| DCF 目標價 | 5,636.44 | 5,639.96 |
| Comparable 目標價 | 4,142.88 | 8,553.69 |
| 加權目標價（上檔空間） | 4,889.66（-19.3%） | 7,096.82（+18.0%） |
| 投資建議 | 中立 | 中立 |
| 收支關卡 A／B／C | 361.5／335.8／311.4 | 361.5／335.8／311.4 |
<!-- /version-table -->
