# software_ai/accenture/ — Accenture（ACN）個股模型 v0.2.3（原型 L：按人力計價）

**本版未做事前驗證**（專案負責人 2026-11-16 23:27 指示）。工作單：`software_ai/docs/workorders/20261116_第一批_研究包與單位經濟.md` C 節與 G、G2、G3、G4、G5 節。

**建置：** `bash software_ai/accenture/make_v023.sh`（`SENS=1` 重跑敏感度與反解）。成品 `dist/20261117_Accenture個股模型_v0.2.3.xlsx`／`.html`。投資建議、目標價與各數字見成品 README 頁與 HTML（由公式組字，本檔不手寫）。

**版本表（由建置成品產生，不手寫）：**

<!-- version-table: dist/_待刪除/20261117_Accenture個股模型_v0.2.2.xlsx｜dist/20261117_Accenture個股模型_v0.2.3.xlsx -->
| 項目 | v0.2.2 | v0.2.3 |
|---|---|---|
| 現價 | 250.49 | 250.49 |
| DCF 目標價 | 172.21 | 172.21 |
| Comparable（P/S 主腿）目標價 | 264.60 | 261.83 |
| 加權目標價（上檔空間） | 218.41（-12.8%） | 217.02（-13.4%） |
| 投資建議 | 中立 | 中立 |
| 收支關卡 A／B／C（FY2027–FY2031，US$M） | 21,454.5／17,012.0／7,718.5 | 21,454.5／17,012.0／7,718.5 |
<!-- /version-table -->

**命題：** AI 進入按人力計價的 IT 服務後，每單位營收＝AI 新專案增量 − 定價侵蝕 − 推論成本。AI 前營收 R0＝交付人力 × 利用率 × 可用工時 × 時薪；AI 後營收＝R0 −（被自動化工時 × 時薪 ×（1−價格保留率））＋ R0 × 新專案占比；推論成本＝被自動化工時 ÷ 任務長度 × 每成功任務成本，只讀 Tokenomics v22.0 快照（`data/tokenomics_snapshot_v22.0.json`，任務 5）。變數名稱沿用 `software_ai/unit_economics`（L_util、L_hrs、L_rate、a、r、n），該模型 v0.1 完成前在本模型 In 頁自建（v0.2 改讀 unit_economics）。

**結構：** README、In、Q（10-K 歷史）、Seg（命題與三分部營收）、PL（標準損益表、分部營收與營業利益、三情境）、DCF（含收支關卡）、Val、RevDCF、Recon（含整本重算反解）、Sens、Sources。官方報告分部三個（Americas、EMEA、Asia Pacific；10-K FY2026 Note 16）；分部營收與營業利益合計與合併值差 0。

**情境：** A 樂觀／B 基準／C 保守，在 AI 淨效果上分化（自動化比例 `a32`、價格保留率 `r`、新專案占比 `n32`），依據與標記見 In 頁。

**檔案：** `build_accenture.py`、`sens_accenture.py`（敏感度與反解，整本重算）、`gen_html_json.py`、`segments.json`、`html.json`、`display_check.json`、`make_v022.sh`、`data/`（研究包、事實總帳、Tokenomics 快照、敏感度與反解結果）、`dist/`。

## v0.2 改版（G5，2026-11-17）
- Comparable 相對倍數（rel_mult／cm_adj）改讀 `software_ai/data/relmult_hist/out/accenture.json` 的 `geo_base`（Derived；副本 `data/relmult_hist_accenture.json`）；Sens 加 `ps_only`／`psg_only`（缺者寫明）與 `sens.low`／`sens.high`（即 cm_adj 低高端）。倍數為 EV／EBIT，以 PS／PSG 相對值近似套用。
- 無風險利率統一：美國 10 年期 5.10%（FRED DGS10，2026-11-12）。
- 股票薪酬加回 DCF 讀數：Val 與 Recon 並列，只作對照，不進加權目標價。
- 舊成品移 `dist/_待刪除/`；`data/_待刪除/` 放 v0.1 敏感度與反解結果。

## v0.2.1（G6，2026-11-17）
- Comparable 主腿改 P/S 口徑：同業中位 P/S（relmult_hist 最近一季 2026-09-30，2 家，已標逾期）× rel_mult（geo_base）× 目標時點 TTM 營收 ÷ 股數。原因：PS／PSG 相對值已含利潤率差，套到 EV／EBIT 會重複計入。
- EV／EBIT 分部加總改對照列（rel_mult＝1.00 與本公司現行相對值兩讀數），不進加權目標價。Sens 的 ps_only、psg_only、sens.low、sens.high 套在 P/S 主腿。

## v0.2.2（G7，2026-11-17）
- 同業中位 P/S 改讀 `software_ai/data/relmult_hist/out/peer_ps_20261113.json`（現價日 2026-11-13；非當日者逐家標日期與逾期）；副本 `data/peer_ps_20261113.json`。rel_mult 仍讀 geo_base。
- Sens 加「同業 P/S＝資料庫最近一季中位」對照案（v0.2.1 的值）。P/S 主腿為股權口徑，不加淨現金。

## 不適用與未做
- Meta 的算力、Tokenomics 自用 AI 讀法 2、Wattonomics 對照：不適用：Accenture 不自建算力，推論成本為向模型商購買。
- 三層協調（Coord 頁、C0、兩讀數）：不適用：非 Siliconomics 覆蓋。
- 事前驗證：未做。
- 公司不揭露計費工時、時薪、價量拆分、token 量與代理任務量：未找到；時薪由營收反推（Derived），每任務 token 取 Tokenomics 任務 5（類比）。
- Blume β 第二來源、公司債有效利率、融資租賃餘額、同業 NTM 倍數、平均賣方目標價與家數：未找到。
- 同業倍數三家、日期 2026-10-21 至 10-27，已標逾期，使用前須更新。
- 本公司歷史相對倍數（G3 第 4 條、G5 第 1 條）：已由共用資料庫取得（PS 與 PSG 幾何平均）。

## 已知限制
- 同業 EV／EBIT 取自 StockAnalysis，準則與租賃處理未逐家核對。
- 人力成長含併購併入，對應 FCFF 內併購現金；匯率未建模（美元口徑對公司當地幣指引）。
- 展示層與治理檢查見 make 腳本末段。

## v0.2.3（G8，2026-11-17）
- 修正 P/S 主腿股權價值：Val!C52:E52 由「P/S × 目標日 TTM 營收」改為再扣非控制權益帳面值（In!E31＝US$1,689.86M，扣一次、不乘倍數）後除以評價股數；仍不加淨現金、不扣淨債務。
- 去除公開檔中的個人情境用語；舊成品與舊敏感度／反解檔移 `dist/_待刪除/`、`data/_待刪除/`；建置腳本改 `make_v023.sh`。
