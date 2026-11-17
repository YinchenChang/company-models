# software_ai/crowdstrike/ — CrowdStrike（CRWD）個股模型 v0.2.1（原型 P：按防護範圍計價）

**本版未做先驗**（Andy 2026-11-16 23:27 指示）。工作單：`software_ai/docs/workorders/20261116_第一批_研究包與單位經濟.md` F、G、G2、G3、G5、G6、G7 節。公開 repo：只含公司與市場事實，無任何個人情境。

**建置：**
1. `python3 software_ai/crowdstrike/build_crowdstrike.py`（全部活公式；含整本重算敏感度，約 2 分鐘；`--quick` 沿用 `data/sens_v01.json`；`--reverse` 重算 `data/reverse_v01.json` 整本重算反解）。成品 `dist/20261117_CrowdStrike個股模型_v0.2.1.xlsx`（LibreOffice 重算後存檔）。
2. `python3 software_ai/crowdstrike/gen_html_json.py <xlsx> <layout.json> software_ai/crowdstrike/html.json`（layout.json 由步驟 1 寫到暫存目錄）。
3. `python3 tools/company_html/build_company_html.py software_ai/crowdstrike` → `dist/…_v0.2.1.html`。

投資建議、目標價與各數字見成品 README 頁與 HTML（由公式組字，本檔不手寫）。

**命題：** ARR＝受保護單位 × 每單位模組數 × 每模組年費；AI 代理（非人類身分）另計單位與價格保留率；AI 淨效果＝增量營收 − 定價侵蝕 − 增量推論成本。推論成本＝任務數 × 每任務 token × 市場價，只讀 Tokenomics v22.0 快照（`data/tk_snapshot_v22.0.json`），基期推論成本已含在實績毛利，只扣增量。三情境 A／B／C 在 AI 淨效果上分化：AI 代理單位數與價格保留率、每模組年費變動、任務數。

## 版本
| 版本 | 日期 | 變更 |
|---|---|---|
| v0.1 | 2026-11-16 | 首版；G2-3 把費用率、SBC 率、資本支出率全期持平，DCF 約 1、加權目標價約 35，強力放空（舊檔在 `dist/_待刪除/`） |
| v0.1.1 | 2026-11-17 | G3：(1) 規模經濟改以 FY2023–FY2026 回歸的固定＋變動成本延伸（non-GAAP 成本固定部分、SBC 固定部分取 0、資本支出固定部分），全期持平改為 Sens 下緣；(2) 營運資金符號核對，折舊攤銷改由前一年資本支出推得（`da_k`）；(3) `rel_mult` 改用最近 5 個會計年度末自算平均（`data/relmult_history.csv`）；(4) 合理性關卡診斷寫入成品 README 頁；新增事件尾端 Sens（`ev_hc`、`ev_delta`） |
| v0.1.2 | 2026-11-17 | 批次獨立核對修正：(1) 合理性關卡補第三項診斷：情境 A 的 DCF 企業價值為負；(2) 揭露三情境 DCF 排序倒掛（終值現值皆為負、營收越大負終值越大，比照原則第 60 條）；(3) EV 小於或等於零時終值占 EV 顯示 N/M，Val 頁新增 DCF 有效性檢查（情境 A 不通過），DCF 失效條件加「EV 小於或等於零」；(4) 收支關卡 B16／B26 統一為三情境最小值（CF 頁 B3） |
| v0.2 | 2026-11-17 | G5：(1) `rel_mult` 改讀 `software_ai/data/relmult_hist/out/crowdstrike.json` 的 `geo_base`（PS 與 PSG 兩讀數幾何平均，Derived），取代自算的年度 EV／營收值；以 PS／PSG 相對值近似套用於本模型的 EV／營收（TTM）倍數；Sens 低／高＝`sens.low`／`sens.high`，另列 `ps_only`、`psg_only` 兩案；`divergence_gt_30pct` 為假，無需另寫成長比原因；(2) 無風險利率改 FRED DGS10 5.10%（2026-11-12）；(3) Recon 與 HTML 估值章並列 DCF（股票薪酬加回）讀數，不進加權目標價；(4) 合理性關卡重跑，診斷不變（情境 A EV 為負、DCF 偏低、強力放空） |
| v0.2.1 | 2026-11-17 | G6／G7：Comparable 主腿改 P/S 股權口徑：同業中位 P/S 讀 `software_ai/data/relmult_hist/out/peer_ps_20261113.json`（價格日 2026-11-13；SentinelOne 為 2026-10-26，標逾期）× `geo_base` × 目標時點 TTM 營收（相鄰兩個會計年度依日數加權）÷ 稀釋股數，不加淨現金；主腿用 P/S 是為了與 PS／PSG 相對值同口徑；原 EV／營收改為 Val 頁對照列，不進加權目標價；Sens 加「同業 P/S＝資料庫最近一季中位」案（`ps_med` 低值）；`ps_only`、`psg_only`、`sens.low`、`sens.high` 四案套在新主腿 |

## 合理性關卡診斷（G3-3，v0.1.2 更新）
觸發三項：(a) DCF 目標價低於現價的 25%；(b) 情境 A 的 DCF 企業價值為負（情境 B、C 為正）；(c) 評等為強力放空。FY2027 營收與 non-GAAP 營業利益和指引的差距未超過 10%。
診斷：以上都不是公式或符號錯誤，而是口徑結果：SBC 費用化，歷史 SBC 占營收（含薪資稅約四分之一）沒有下降趨勢，GAAP 營業利益率在預測期與終值期都在零附近（略為負），NOPAT 與終值因而為負或近零。數字見成品 README 頁（公式組字）。
**情境排序倒掛（比照原則第 60 條）：** 三情境的 DCF 終值現值皆為負，DCF 每股價值依情境 A、B、C 遞增，方向與 AI 淨效果（A 大於 B 大於 C）相反。成因是 GAAP 營業利益率長期在零附近，營收越大，負的 NOPAT 與負終值越大，營收較高的情境 DCF 反而較低。因此三情境的 DCF 排序不代表經營好壞排序，EV 為負的情境不作為估值依據，加權目標價主要由 Comparable 決定。
**DCF 失效條件：** 任一情境 DCF 企業價值小於或等於零（原則第 56 條）；此時終值占 EV 等比率顯示 N/M，Val 頁「DCF 有效性檢查」判不通過。

**結構：** README、In、Hist、Q、Thesis、PL（標準損益表、分部營收與對帳、三情境）、CF（FCFF DCF，含 FY2033–FY2037 延伸期）、Val、RevDCF、Recon、Sens、Sources。官方分部只有一個，營收線為訂閱與專業服務。

**檔案：** `build_crowdstrike.py`、`gen_html_json.py`、`segments.json`（根目錄與 `data/` 各一份，內容相同）、`html.json`、`data/`（`research_crowdstrike.md`、`facts_crowdstrike.csv`、Tokenomics 快照、`sens_v01.json`、`reverse_v01.json`）、`dist/`。

## G2 共通必修（逐條）
1. 每任務 token：**未做到（公司未揭露 token 量或代理任務量，無從校準）**；沿用 Tokenomics 任務 3 的 token 數（Assumed，任務類別選擇為 Assumed），價格層級選擇標 Analogy。
2. 聯動 Sens：做到。token 價格高案標上界（只動成本一邊）；另列價格與成本聯動、模型組合路由、低價聯動三案。
3. 比率基準持平：v0.1 做到；v0.1.1 依 G3-1 改為歷史回歸的固定＋變動成本延伸，全期持平案放 Sens 下緣（`fix_ng`、`fix_sbc`、`fix_cx` 低值 0）；延伸期利益率仍持平（`ext_d` 基準 0）。
4. 股數：部分做到。以公司稀釋加權平均近似（含 RSU 與選擇權庫藏股法影響），RSU 預期沒收與庫藏股法明細未找到；ASR、可轉債未見；法定封面流通股數作對照。
5. 淨債務：做到。只計優先票據帳面，營業租賃不計；融資租賃未另列（未找到）。
6. Comparable 口徑：部分做到。EV（無槓桿）對營收（TTM），乘目標日 TTM 營收；同業各家會計準則與租賃處理未核對，逐家標明；揭露現行倍數 ÷ 同業中位數與反解係數。
7. 收支關卡：做到。Val 頁有估值評等、入不敷出判定、投資建議三列；`check_rating_gate.py --strict` 通過。
8. 文字數字公式組字：做到。README 日期、指引、落差歸因由公式組字；DCF 與 Comparable 落差以公式分解（股票薪酬、槓桿、時點、倍數殘差；股票薪酬項連結 Sens 重算值）。
9. 反解：做到。主導變數（相對倍數係數）與人類端單位年增整本重算反解，寫在 RevDCF 頁。
10. 逾期標示：做到。SentinelOne 同業倍數（2026-10-26 收盤）標逾期；rf 與現價同為 2026-11-13。

## 不適用與未做
- 算力 MW、Wattonomics 對照、Tokenomics 自用 AI 讀法 2：不適用（CrowdStrike 向模型商購買推論）。三層協調：不適用。
- 先驗：未做。展示層數字一致性檢查（`display_check.json`，原則第 51 條）：未做。現金稅率橋、事件型尾端 Sens（DOJ／SEC 資訊請求、Delta 案）：未做。
- 未找到：AI 專屬營收、Charlotte AI 用量與 ARR、AIDR 絕對 ARR、受保護端點總數、FY2022–FY2023 年度淨新增 ARR 與淨留存、FY2028 以後共識明細、RSU 與選擇權明細、現金稅率（見 `data/research_crowdstrike.md` 第 9 節與 Recon 頁）。

## 已知限制
- 現價與同業倍數只有單一資料商（StockAnalysis）；賣方「次一年」營收共識的會計年度待確認；FY2022–FY2023 員工數為第二來源。
- 合理性關卡（G3-3）：DCF 目標價低於現價 25% 與評等為強力放空兩項觸發，診斷結論寫在成品 README 頁；成因是 SBC（含薪資稅）占營收約四分之一且歷史無下降趨勢，GAAP 營業利益率在預測期仍接近零，結果由『SBC 視為費用、無證據不假設下降』口徑決定，不是對公司價值的經濟定論；加回 SBC、SBC 規模經濟加大案、延伸期利益率上移案在 Sens。
- 命題子列的人類端單位與模組路徑（gu、gm）、價格保留率（ero、pf）、相對倍數係數（rel_mult）為承重輸入。
