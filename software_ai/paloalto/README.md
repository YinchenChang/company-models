# software_ai/paloalto/ — Palo Alto Networks（PANW）個股模型 v0.1.1（原型 P：按防護範圍計價）

**本版未做先驗**（工作單 G 節：先驗跳過）。工作單：`software_ai/docs/workorders/20261116_第一批_研究包與單位經濟.md`（共同規則、研究包格式、G、G2、G3）。

**建置：** `bash software_ai/paloalto/make_paloalto.sh`（約 1.5 分鐘；含敏感度整本重算、html.json、HTML 與全部核對，包含 `check_rating_gate.py --strict`）。成品 `dist/20261117_Palo Alto Networks個股模型_v0.1.1.xlsx`／`.html`。投資建議、目標價與各數字見成品 README 頁與 HTML（由公式組字，本檔不手寫）。

**命題：** AI 讓資安的受保護範圍從人與端點擴大到代理（非人類身分）。NGS ARR ＝ 受保護單位 × 每單位營收；每單位營收 ＝ R0 ×（1 − 價格侵蝕 ＋ 模組加購 ＋ 代理身分數 × 代理單價比）；推論成本 ＝ 分析次數 × 每次 token × 市場價，只讀 Tokenomics v22.0 快照（`software_ai/unit_economics/data/tk_snapshot.json`）。變數名稱沿用 `software_ai/unit_economics`（P_mod、P_px、P_rho、P_n26／n32、P_m26／m32、P_e26／e32、P_k26／k32、P_tier、g_px、g_tok、ovh）；該模型 v0.1 完成前在本模型 In 頁自建，v0.2 改讀 unit_economics。

**結構：** README、In、Q、Hist、TK、PL（情境 B）、PL_A、PL_C、Val、RevDCF、Recon、Sens、Sources。官方分部只有一個（10-K 附註 19）；營收列取 Product 與 Subscription and support。估值＝FCFF DCF（至 FY2032，期中折現，終值 NOPAT×(1−g/RONIC)/(WACC−g)）與同業前瞻本益比各半，情境 25／50／25。會計年度 7 月底；FY2026 含 CyberArk（2026-02-11 完成）。

**檔案：** `build_paloalto.py`（Excel 產生器，全部活公式）、`gen_html_json.py`（依列標籤產生 `html.json`）、`segments.json`、`make_paloalto.sh`、`data/`（研究包 `research_paloalto.md`、`facts_paloalto.csv`、`make_research.py`、`sens_v011_20261117.json`）、`dist/`。

## 不適用與未做
- 不適用：Wattonomics 對照與自用算力：PANW 不自建算力，推論成本為向模型商購買的市場價。
- 不適用：Tokenomics 每 MW 持有成本（IF_HoldEcon）群組與 he_all 案。
- 不適用：三層協調（Coord 頁、C0、兩讀數）：非 Siliconomics 覆蓋。
- 不適用：分部加總：單一應報導分部。
- 不適用：收支模組與融資瀑布：淨現金為正。
- 未做：先驗。
- 未找到：公司揭露的 token 量或代理任務量（G2 第 1 條無揭露可校準，用 Tokenomics 快照的任務 token，選層級屬類比）；受保護單位數、每單位年價、模組數；單一經銷商占比；CyberArk 商譽稅務抵減金額與年限；同業 NTM 成長率與利益率；PANW 歷史前瞻本益比序列、Cloudflare 歷史市銷率（macrotrends 禁止擷取）。

## v0.1.1 變更（2026-11-17）
- rel_adj 基準改為歷史相對倍數中位數（Derived）：PANW 近 5 年 21 個財季末 TTM 市銷率 ÷ 同業（Zscaler、Fortinet、Check Point）中位數，見 `RelHist` 頁；取不到歷史前瞻本益比，以市銷率代理。Sens 區間取歷史最低／最高。
- 上檔空間三段歸因（DCF、同業中位倍數、歷史相對倍數調整）列在 README 橫幅與 Val 頁；README 與 HTML 寫明 DCF（RONIC＝WACC）與 Comparable 對競爭優勢的假設相反。
- 反向 DCF 隱含營業利益率超過 100% 時改為「無解」結論，並反解所需 RONIC（無解時寫明）。反解算法已驗算：把核心率設為反解值，EV 與現價隱含 EV 相符。

## 已知限制
- 無風險利率 4.99% 取 Meta 模型共用值（日期 2026-10-27），與現價日相差逾 2 個交易日，標逾期。現價 322.24（2026-11-13）單一來源。
- 可轉債以公允價值計為債務（深度價內，不再計股數），未採「本金＋股數」；營業租賃不計。股數 818 百萬加 RSU 17 百萬，PSU 只在 Sens。
- Comparable 為同業前瞻本益比中位數 × 模型 FY2028 Non-GAAP EPS，同業為 US GAAP 資料商前瞻口徑，未取得 NTM 對齊。
- 本模型成本不隨價格侵蝕下降；FY2027 營收高於指引上緣、NGS ARR 低於指引下緣（單位成長不以指引校準，見 Recon）。

## G3 合理性關卡診斷
觸發：DCF 目標價低於現價 25%、評等強力放空。診斷：費用率已依歷史規模經濟延伸（Non-GAAP 營業利益率每年 +0.3pt，SBC 占營收每年 −0.5pt，標 Derived，全期持平案為 Sens 下緣）；將費用率改善、折現率下調、終值替代法同時取最樂觀，DCF 目標價仍遠低於現價；現價隱含核心營業利益率超過 100%。結論：經濟判斷（現價依賴 FY2032 之後長期高成長，而終值只給 RONIC＝WACC），不是建模錯誤。數字見成品 README 頁「合理性關卡診斷」列（公式組字）。

## 承重 Assumed 與可證偽指標
見成品 README 頁最末列：受保護代理身分數與加購（n32、m32）、價格侵蝕（e32）、費用率延伸路徑（nm_up、sbc_dn）。
