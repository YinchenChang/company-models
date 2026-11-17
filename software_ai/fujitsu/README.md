# software_ai/fujitsu/ — 富士通（6702）個股模型 v0.2（原型 L：按人力計價，套用於 Service Solutions）

**本版未做先驗**（Andy 2026-11-16 23:27 指示）。工作單：`software_ai/docs/workorders/20261116_第一批_研究包與單位經濟.md` C 節與 G、G2、G3 節。

**建置：** `bash software_ai/fujitsu/make_v02.sh`（`SENS=1` 重跑敏感度與反解）。成品 `dist/20261117_Fujitsu個股模型_v0.2.xlsx`／`.html`。投資建議、目標價與各數字見成品 README 頁與 HTML（由公式組字，本檔不手寫）。

**版本表（由建置成品產生，不手寫）：**

<!-- version-table: dist/20261117_Fujitsu個股模型_v0.2.xlsx -->
| 項目 | v0.2 |
|---|---|
| 現價 | 3,455.00 |
| DCF 目標價 | 3,218.96 |
| Comparable（分部加總）目標價 | 3,575.74 |
| 加權目標價（上檔空間） | 3,397.35（-1.7%） |
| 投資建議 | 中立 |
| 收支關卡 A／B／C（FY2026–FY2030，¥百萬） | 1,294,356.2／1,152,885.3／858,042.1 |
<!-- /version-table -->

**命題：** AI 進入按人力計價的 IT 服務後，每單位營收＝AI 新專案增量 − 定價侵蝕 − 推論成本。Service 的 AI 前營收 R0＝交付人力 × 利用率 × 可用工時 × 時薪；AI 後營收＝R0 −（被自動化工時 × 時薪 ×（1−價格保留率））＋ R0 × 新專案占比；推論成本＝被自動化工時 ÷ 任務長度 × 每成功任務成本，只讀 Tokenomics v22.0 快照（`data/tokenomics_snapshot_v22.0.json`，任務 5）。變數名稱沿用 `software_ai/unit_economics`（L_util、L_hrs、L_rate、a、r、n、tier），v0.2 改讀 unit_economics／待 AI供應鏈 登錄。Hardware 與 Ubiquitous 為不按人力計價的獨立分部，以年增率預測。

**結構：** README、In、Q（決算短信歷史與半年度）、Seg、PL、DCF（含收支關卡）、Val、RevDCF、Recon（含整本重算反解）、Sens、Sources。官方報告分部四列（Service Solutions、Hardware Solutions、Ubiquitous Solutions、Elimination/Corporate）；分部營收與分部調整後營業利益合計與合併值差 0。

**情境：** A 樂觀／B 基準／C 保守，在 AI 淨效果上分化（自動化比例 `a32`、價格保留率 `r`、新專案占比 `n32`），依據與標記見 In 頁。

**檔案：** `build_fujitsu.py`、`sens_fujitsu.py`、`gen_html_json.py`、`segments.json`、`html.json`、`display_check.json`、`make_v02.sh`、`data/`、`dist/`。

## v0.2 改版（G5，2026-11-17）
- Comparable 相對倍數（rel_mult／cm_adj）改讀 `software_ai/data/relmult_hist/out/fujitsu.json` 的 `geo_base`（Derived；副本 `data/relmult_hist_fujitsu.json`）；Sens 加 `ps_only`／`psg_only`（缺者寫明）與 `sens.low`／`sens.high`（即 cm_adj 低高端）。倍數為 EV／EBIT，以 PS／PSG 相對值近似套用。
- 無風險利率統一：日本 10 年期 2.77%（Trading Economics，2026-11-11）。
- 股票薪酬加回 DCF 讀數：Val 與 Recon 並列，只作對照，不進加權目標價。
- 舊成品移 `dist/_待刪除/`；`data/_待刪除/` 放 v0.1 敏感度與反解結果。

## 不適用與未做
- Meta 的算力、Tokenomics 自用 AI 讀法 2、Wattonomics 對照：不適用：富士通不以自建算力為命題。
- 三層協調（Coord 頁、C0、兩讀數）：不適用：非 Siliconomics 覆蓋。
- 先驗：未做。
- 分部員工、利用率、計費工時、時薪、token 量與代理任務量：公司不揭露，未找到；時薪由 Service 營收反推（Derived），每任務 token 取 Tokenomics 任務 5（類比）。
- Blume β 第二來源、債券有效利率、退休金淨負債、稀釋證券明細、匯率來源、歷史營運資金比率、決算後賣方目標價與共識：未找到。
- 本公司歷史相對倍數（G3 第 4 條、G5 第 1 條）：已由共用資料庫取得，只用 PS（PSG 有效點 6 少於 12，參考讀數不進基準）；psg_only 案缺。
- 股票報酬費用金額：未找到，加回 DCF 讀數無法計算。
- G3 第 1 條：費用率（毛利率、折舊、資本支出）持平於 LTM，因公司以分部利益率持平加 AI 路徑建模，未發現可用的歷史規模經濟證據；Sens 涵蓋毛利率區間。

## 合理性關卡診斷（工作單 G3 第 3 條）
結論見成品 README 頁 A19（公式生成）。本版未觸發；FY2026 調整後營業利益率對公司指引的相對差接近 10% 門檻（模型分部利益率持平 LTM，公司指引隱含下半年回升）。現價 2026-11-13（¥3,455）較 2026-10-23（¥4,477）低約 23%，單一來源，待核對。

## 已知限制
- 同業 EV／EBIT 取自 StockAnalysis（4 家；日期 2026-10-05 至 11-13），準則與租賃處理未逐家核對；BIPROGY 為日本會計準則。
- 2026 下半年實績預計 2027-04 公布；FY2026 DCF 下半年現金流為全年 FCFF 乘計入比例。
