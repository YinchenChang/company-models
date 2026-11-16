# software_ai/crowdstrike/ — CrowdStrike（CRWD）個股模型 v0.1（原型 P：按防護範圍計價）

**本版未做先驗**（Andy 2026-11-16 23:27 指示）。工作單：`software_ai/docs/workorders/20261116_第一批_研究包與單位經濟.md` F、G、G2 節。公開 repo：只含公司與市場事實，無任何個人情境。

**建置：**
1. `python3 software_ai/crowdstrike/build_crowdstrike.py`（全部活公式；含整本重算敏感度，約 2 分鐘；`--quick` 沿用 `data/sens_v01.json`；`--reverse` 重算 `data/reverse_v01.json` 整本重算反解）。成品 `dist/20261116_CrowdStrike個股模型_v0.1.xlsx`（LibreOffice 重算後存檔）。
2. `python3 software_ai/crowdstrike/gen_html_json.py <xlsx> <layout.json> software_ai/crowdstrike/html.json`（layout.json 由步驟 1 寫到暫存目錄）。
3. `python3 tools/company_html/build_company_html.py software_ai/crowdstrike` → `dist/…_v0.1.html`。

投資建議、目標價與各數字見成品 README 頁與 HTML（由公式組字，本檔不手寫）。

**命題：** ARR＝受保護單位 × 每單位模組數 × 每模組年費；AI 代理（非人類身分）另計單位與價格保留率；AI 淨效果＝增量營收 − 定價侵蝕 − 增量推論成本。推論成本＝任務數 × 每任務 token × 市場價，只讀 Tokenomics v22.0 快照（`data/tk_snapshot_v22.0.json`），基期推論成本已含在實績毛利，只扣增量。三情境 A／B／C 在 AI 淨效果上分化：AI 代理單位數與價格保留率、每模組年費變動、任務數。

**結構：** README、In、Hist、Q、Thesis、PL（標準損益表、分部營收與對帳、三情境）、CF（FCFF DCF，含 FY2033–FY2037 延伸期）、Val、RevDCF、Recon、Sens、Sources。官方分部只有一個，營收線為訂閱與專業服務。

**檔案：** `build_crowdstrike.py`、`gen_html_json.py`、`segments.json`（根目錄與 `data/` 各一份，內容相同）、`html.json`、`data/`（`research_crowdstrike.md`、`facts_crowdstrike.csv`、Tokenomics 快照、`sens_v01.json`、`reverse_v01.json`）、`dist/`。

## G2 共通必修（逐條）
1. 每任務 token：**未做到（公司未揭露 token 量或代理任務量，無從校準）**；沿用 Tokenomics 任務 3 的 token 數（Assumed，任務類別選擇為 Assumed），價格層級選擇標 Analogy。
2. 聯動 Sens：做到。token 價格高案標上界（只動成本一邊）；另列價格與成本聯動、模型組合路由、低價聯動三案。
3. 比率基準持平：做到。non-GAAP 營業費用率、SBC 率、資本支出率、延伸期利益率基準全期持平於近四季，收斂只放 Sens。
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
- 基準全期持平下，GAAP 營業利益在整個預測期偏低，DCF 單法接近零；費用率收斂案只在 Sens（opx_32、sbc_32、ext_d 低／高值）。
- 命題子列的人類端單位與模組路徑（gu、gm）、價格保留率（ero、pf）、相對倍數係數（rel_mult）為承重 Assumed。
