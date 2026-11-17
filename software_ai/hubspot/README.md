# software_ai/hubspot/ — HubSpot（HUBS）個股模型 v0.2（原型 S：按席位計價）

**本版未做先驗。** 工作單：`software_ai/docs/workorders/20261116_第一批_研究包與單位經濟.md` G、G2、G3、G5 節。v0.2（2026-11-17）依 G5 改版；v0.1 成品在 `dist/_待刪除/`。

**建置：** `bash software_ai/hubspot/make_v02.sh`（`SENS=1` 重跑敏感度與反解）。成品 `dist/20261117_HubSpot個股模型_v0.2.xlsx`／`.html`。投資建議、目標價與各數字見成品 README 頁與 HTML（由公式組字，本檔不手寫）。

**版本表（由建置成品產生，不手寫）：**

<!-- version-table: dist/_待刪除/20261116_HubSpot個股模型_v0.1.xlsx｜dist/20261117_HubSpot個股模型_v0.2.xlsx -->
<!-- /version-table -->

**命題：** AI 進入 CRM 後，每客戶訂閱營收受席位被 AI 取代而侵蝕；AI 代理（Customer Agent）按解決任務計價為增量；推論成本只讀 Tokenomics v22.0 快照（`data/tokenomics_snapshot_v22.0.json`，0.1 小時任務向量，類比）。三情境在席位取代率、agent 採用率與單價變動上分化。

**結構：** README、In、Hist、Q4E、Thesis、PL、CF、Val、RevDCF、Recon、Sens、Sources。官方分部單一；營收線訂閱、專業服務與其他。

**檔案：** `build_hubspot.py`、`sens_hubspot.py`、`facts_hubspot.py`、`gen_research.py`、`hs_inputs.py`、`gen_html_json.py`、`segments.json`、`html.json`、`display_check.json`、`make_v02.sh`、`data/`、`dist/`。

## 不適用與未做
- 算力（Compute）：不適用：HubSpot 不自建算力，推論成本向模型商購買。
- Tokenomics 自用 AI 讀法 2：不適用：同上。
- Wattonomics：不適用：不自建算力。
- 三層協調（Coord 頁、C0、兩讀數）：不適用：非 Siliconomics 覆蓋。
- 先驗：未做先驗。
- 未找到：每任務 token 數公司未揭露、選擇權履約價、融資租賃餘額、Blume 第二來源以外的 beta 來源等，見 `data/research_hubspot.md` 第 9 節與 HTML 未找到清單。

## G5 v0.2 五條逐條
1. 相對倍數：`rel_mult` 讀 `software_ai/data/relmult_hist/out/hubspot.json` 的 `geo_base`＝1.2715（Derived；base_source：PS 與 PSG 兩讀數幾何平均；窗口 2021-09-30 至 2026-06-30，同業 11 家）。低／高取 `sens.low`／`sens.high`＝0.6465／1.5797。Sens 四案齊備：`ps_only`（1.2592）、`psg_only`（1.2839）、`sens.low`、`sens.high`（後兩者即 rel_mult 低／高案）。本模型 Comparable 為前瞻 P/E，**以 PS／PSG 相對值近似套用於前瞻 P/E**。`divergence_gt_30pct` 為假（PS 與 PSG 兩讀數相差 2.0%），不需寫分歧原因；HUBS 與同業 TTM 營收成長比 1.52。提醒：HUBS 現行前瞻 P/E（13.2）僅為同業中位數（15.4）的 0.86 倍，歷史 PS／PSG 相對值 1.27 高於現行 P/E 相對值，Comparable 因此較 v0.1 高。
2. 無風險利率：美國 10 年期 5.10%（FRED DGS10，2026-11-12，現價日前 1 個交易日；Verified），取代 v0.1 的 5.15%（Trading Economics）。
3. 股票薪酬加回 DCF 讀數：Val 頁、Recon 第 7 節與 HTML 估值章並列「DCF（股票薪酬加回）每股價值與目標價」，公式組字；＝費用化讀數加稅後 SBC 現值（含終值），股數不變；只作對照，不進加權目標價。
4. 版本 v0.2；v0.1 成品移 `dist/_待刪除/`；全部檢查與 SENS=1 重跑；合理性關卡重跑，診斷結論見下。
5. 回報：新舊目標價、rel_mult 新舊值、SBC 加回讀數與檢查結果見回報訊息；數字以成品 README 頁為準。

**v0.2 合理性關卡診斷：** 四關皆過（DCF 目標價為現價的 42%、EV 為正、評等中立、FY2026 對指引最大差距 2.5%），未觸發。Comparable 與 DCF 的落差擴大（相對倍數係數提高），股票薪酬約占落差 44%，其餘為同業 P/E 隱含的成長與風險溢酬；加回股票薪酬的 DCF 目標價接近現價，只作對照。

## G2 共通必修十條逐條
1. 每任務 token 以公司揭露校準：公司只揭露每解決任務單價與 Customer Agent 營收，token 數未找到，取 Tokenomics 0.1 小時任務（類比），標記於 In 頁。做到（以類比，已註明）。
2. Sens 案例連結：每個低／高不等於基準的輸入都有案例，`check_sens_coverage --require-keycol` 通過。做到。
3. 比率參數除非有證據否則持平：有歷史規模經濟證據者改為固定加變動拆分（見 G3.1）。做到。
4. 股數：RSU 扣除沒收、選擇權庫藏股法。RSU 淨沒收做到；選擇權履約價未找到，以稀釋股數代理並列為未做。部分做到。
5. 淨負債＝本金加融資租賃：公司無借款，融資租賃餘額未找到。部分做到（債務為零，已註明）。
6. Comparable 同一口徑：同業前瞻 P/E 乘 HUBS 目標時點 NTM 非 GAAP EPS；Recon 列出同業準則差異與 HUBS 現行 P/E 相對同業。做到。
7. 收支關卡列：Val 三列齊備，`check_rating_gate --strict` 通過。做到。
8. 公式組字與 DCF 對 Comparable 差距分解：股票薪酬、時點、槓桿、其餘。做到。
9. 反解主導變數：Recon 第 3 節對增量營業利益率 inc_m 反解（現價所需值、評等翻轉值）。做到。
10. 逾期標記：資料日距股價基準日超過兩個交易日者標逾期，Recon 第 4 節列出。做到。

## G3 四條逐條與診斷結論
1. 有證據的規模經濟不凍結：非 GAAP 營運成本與股票薪酬改用歷史固定加變動拆分的增量斜率（`inc_m`、`sbc_inc`，Derived），全期持平案放在 Sens 作下緣。做到。
2. 營運資金符號與比率：`nwc_k` 由歷史遞延收入與應收帳款關係推導，營運資金為現金流入。做到。
3. 發布前合理性關卡（v0.2 重跑，仍四關皆過）：DCF 目標價占現價比例、EV 符號、評等是否為強力級、FY2026 與指引差距，四關皆通過（Val 合理性關卡列）。
4. rel_mult：v0.2 改讀 relmult_hist 的幾何平均基準（見下節 G5.1），不再是 1.00。做到。

**G3.3 診斷結論：** 四關皆通過，未觸發診斷。DCF 目標價顯著低於 Comparable，主因是 DCF 將股票薪酬費用化（約占落差一半以上）與 WACC；是經濟判斷，不是建模錯誤。評等中立。

**承重 Assumed 輸入：** 增量非 GAAP 營業利益率（`inc_m`）、席位取代（`repl_*`、`seat_sh`）、客戶淨增與每客戶年增（`na_27`／`na_32`、`g_up`）；`rel_mult` 自 v0.2 起為 Derived（歷史相對值），但仍是 Comparable 的最大單一槓桿。

## 已知限制
- 同業前瞻 P/E 取自第三方，準則與租賃處理未逐家核對。
- 價格年降標「待 AI供應鏈 登錄」，只作敏感度。
