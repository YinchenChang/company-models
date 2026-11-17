# software_ai/paloalto/ — Palo Alto Networks（PANW）個股模型 v0.2（原型 P：按防護範圍計價）

**本版未做先驗**（工作單 G 節：先驗跳過）。工作單：`software_ai/docs/workorders/20261116_第一批_研究包與單位經濟.md`（共同規則、研究包格式、G、G2、G3）。

**建置：** `bash software_ai/paloalto/make_paloalto.sh`（約 1.5 分鐘；含敏感度整本重算、html.json、HTML 與全部核對，包含 `check_rating_gate.py --strict`）。成品 `dist/20261117_Palo Alto Networks個股模型_v0.2.xlsx`／`.html`。投資建議、目標價與各數字見成品 README 頁與 HTML（由公式組字，本檔不手寫）。

**命題：** AI 讓資安的受保護範圍從人與端點擴大到代理（非人類身分）。NGS ARR ＝ 受保護單位 × 每單位營收；每單位營收 ＝ R0 ×（1 − 價格侵蝕 ＋ 模組加購 ＋ 代理身分數 × 代理單價比）；推論成本 ＝ 分析次數 × 每次 token × 市場價，只讀 Tokenomics v22.0 快照（`software_ai/unit_economics/data/tk_snapshot.json`）。變數名稱沿用 `software_ai/unit_economics`（P_mod、P_px、P_rho、P_n26／n32、P_m26／m32、P_e26／e32、P_k26／k32、P_tier、g_px、g_tok、ovh）；該模型 v0.1 完成前在本模型 In 頁自建，v0.2 改讀 unit_economics。

**結構：** README、In、Q、Hist、TK、PL（情境 B）、PL_A、PL_C、Val、RevDCF、Recon、Sens、Sources。官方分部只有一個（10-K 附註 19）；營收列取 Product 與 Subscription and support。估值＝FCFF DCF（至 FY2032，期中折現，終值 NOPAT×(1−g/RONIC)/(WACC−g)）與同業前瞻本益比各半，情境 25／50／25。會計年度 7 月底；FY2026 含 CyberArk（2026-02-11 完成）。

**檔案：** `build_paloalto.py`（Excel 產生器，全部活公式）、`gen_html_json.py`（依列標籤產生 `html.json`）、`segments.json`、`make_paloalto.sh`、`data/`（研究包 `research_paloalto.md`、`facts_paloalto.csv`、`make_research.py`、`sens_v02_20261117.json`）、`dist/`。

## 不適用與未做
- 不適用：Wattonomics 對照與自用算力：PANW 不自建算力，推論成本為向模型商購買的市場價。
- 不適用：Tokenomics 每 MW 持有成本（IF_HoldEcon）群組與 he_all 案。
- 不適用：三層協調（Coord 頁、C0、兩讀數）：非 Siliconomics 覆蓋。
- 不適用：分部加總：單一應報導分部。
- 不適用：收支模組與融資瀑布：淨現金為正。
- 未做：先驗。
- 未找到：公司揭露的 token 量或代理任務量（G2 第 1 條無揭露可校準，用 Tokenomics 快照的任務 token，選層級屬類比）；受保護單位數、每單位年價、模組數；單一經銷商占比；CyberArk 商譽稅務抵減金額與年限；同業 NTM 成長率與利益率；PANW 歷史前瞻本益比序列、Cloudflare 歷史市銷率（macrotrends 禁止擷取）。

## v0.2 變更（2026-11-17；G5）
- rel_adj 基準改讀共用庫 `software_ai/data/relmult_hist/out/paloalto.json` 的 geo_base（PS 與 PSG 兩讀數幾何平均，Derived），取代 v0.1.1 自算的 macrotrends 序列；`RelHist` 頁為庫內讀值。Sens 加 ps_only、psg_only、sens.low、sens.high 四案。以 PS／PSG 相對值近似套用於前瞻本益比。兩讀數差異未超過 30%。
- Comparable 主腿改為同業 P/S（G6）：同業中位 P/S（Fortinet 16.64、Zscaler 7.29、Check Point 4.95，共用庫最近一季，逾期標示）× 歷史相對倍數 × 目標時點 TTM 營收 ÷ 稀釋股數；這樣與 PS／PSG 相對值同口徑。前瞻本益比 Comparable 改為對照列（相對值 1.00 與本公司現行相對值兩讀數），不進加權目標價。
- 無風險利率統一為美國 10 年期 5.10%（FRED DGS10，2026-11-12，現價日前 1 個交易日，未逾期，Verified）。
- Recon 第 6 節與 HTML 估值章並列「DCF（股票薪酬加回）」每股價值與目標價，只作對照，不進加權目標價。
- 合理性關卡重跑並更新診斷（README 頁「合理性關卡診斷」列）。
- v0.1.1 成品移到 `dist/_待刪除/`。

## 已知限制
- 現價 322.24（2026-11-13）單一來源。
- 可轉債以公允價值計為債務（深度價內，不再計股數），未採「本金＋股數」；營業租賃不計。股數 818 百萬加 RSU 17 百萬，PSU 只在 Sens。
- Comparable 為同業前瞻本益比中位數 × 模型 FY2028 Non-GAAP EPS，同業為 US GAAP 資料商前瞻口徑，未取得 NTM 對齊。
- 本模型成本不隨價格侵蝕下降；FY2027 營收高於指引上緣、NGS ARR 低於指引下緣（單位成長不以指引校準，見 Recon）。

## G3 合理性關卡診斷
觸發：DCF 目標價低於現價 25%、評等強力放空。診斷：費用率已依歷史規模經濟延伸（Non-GAAP 營業利益率每年 +0.3pt，SBC 占營收每年 −0.5pt，標 Derived，全期持平案為 Sens 下緣）；將費用率改善、折現率下調、終值替代法同時取最樂觀，DCF 目標價仍遠低於現價；現價隱含核心營業利益率超過 100%。結論：經濟判斷（現價依賴 FY2032 之後長期高成長，而終值只給 RONIC＝WACC），不是建模錯誤。數字見成品 README 頁「合理性關卡診斷」列（公式組字）。

## 承重 Assumed 與可證偽指標
見成品 README 頁最末列：受保護代理身分數與加購（n32、m32）、價格侵蝕（e32）、費用率延伸路徑（nm_up、sbc_dn）。
