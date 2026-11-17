# software_ai/salesforce/ — Salesforce（CRM）個股模型 v0.2.2（原型 S：按席位計價）

**本版未做先驗**（Andy 2026-11-16 23:27 指示）。工作單：`software_ai/docs/workorders/20261116_第一批_研究包與單位經濟.md` B 節與 G 節。

**建置：** `bash software_ai/salesforce/make_v022.sh`（`SENS=1` 重跑敏感度）。成品 `dist/20261117_Salesforce個股模型_v0.2.2.xlsx`／`.html`。投資建議、目標價與各數字見成品 README 頁與 HTML（由公式組字，本檔不手寫）。

**命題：** AI 進入席位型軟體後，每單位營收＝代理營收增量 − 席位被替代的侵蝕 − 推論成本。Agentforce Apps ＝ 席位型營收（席位 × 每席位年費，扣客戶人力替代）＋ 代理營收（任務數 × 每任務單價）；推論成本＝任務數 × 每任務 token × 市場價，只讀 Tokenomics v22.0 快照（`data/tokenomics_snapshot_v22.0.json`）。變數名稱沿用 `software_ai/unit_economics`，該模型 v0.1 完成前在本模型 In 頁自建（v0.2 改讀 unit_economics）。

**結構：** README、In、Q、Seg（命題與營收推導）、PL（標準損益表、分部營收與對帳、三情境）、DCF、Val、RevDCF、Recon、Sens、Sources。官方報告分部只有一個（10-K）；營收列取公司固定揭露的營收型態與新二分類產品線（Agentforce Apps；Data 360, Headless Platform, and Other）。

**情境：** A 樂觀／B 基準／C 保守，在 AI 淨效果上分化（席位替代率 `rho`、每任務動作量成長 `gt0`、每任務單價變動 `dp`），依據與標記見 In 頁。

**檔案：** `build_salesforce.py`（Excel 產生器，全部活公式）、`sens_salesforce.py`（敏感度，整本重算）、`gen_html_json.py`、`segments.json`、`kreq_salesforce.py`（主導變數整本重算反解）、`html.json`、`make_v022.sh`（含 check_rating_gate）、`data/`（研究包 `research_salesforce.md`、`facts_salesforce.csv`、Tokenomics 快照、`sens_v022_20261117.json`、`kreq_v022.json`）、`dist/`。

## 不適用與未做
- Meta 的算力、Tokenomics 自用 AI 讀法 2、Wattonomics 對照：不適用：Salesforce 不自建算力，推論成本為向模型商購買。
- 三層協調（Coord 頁、C0、兩讀數）：不適用：非 Siliconomics 覆蓋。
- 先驗：未做。
- 分部營業利益：無（單一營運分部）。Non-GAAP 預測：未建，只在 Recon 對照公司指引。
- 新二分類在 FY2026 全年與 Q3、Q4 FY26 的回溯值、Blume β 的第二來源、FY2028 營收共識、同日期賣方平均目標價：未找到。
- 展示層數字一致性檢查（`check_display.py`、`display_check.json`、版本表）：本版未建，下次改版補。

## 已知限制
- v0.2 變更（工作單 G5，2026-11-17）：(1) Comparable 相對倍數 rel_mult 改讀 `software_ai/data/relmult_hist/out/salesforce.json` 的 geo_base（PS 與 PSG 兩讀數幾何平均，Derived），本公司 Comparable 倍數為 P／FCF，README 頁寫明「以 PS／PSG 相對值近似套用於 P／FCF」；Sens 加 ps_only、psg_only、sens.low、sens.high 四案（JSON 皆有）；兩讀數分歧 19%（divergence_gt_30pct 為假），無須另寫原因，成長比與分歧度列在 In 頁。(2) 無風險利率改 5.10%（FRED DGS10，2026-11-12，已讀原檔）。(3) Val、Recon、HTML 估值章並列 DCF（股票薪酬加回）讀數，只作對照、不進加權目標價。(4) 舊成品與舊 Sens／反解檔移至 `_待刪除/`。
- v0.2.2 變更（工作單 G7）：同業中位 P／S 改讀 `software_ai/data/relmult_hist/out/peer_ps_20261113.json`（現價日 2026-11-13；4.605，6 家；WDAY、HUBS 為 2026-10-26，逾期），取代資料庫季末值 4.065；Sens 加「同業 P／S＝資料庫最近一季中位」對照案（即 v0.2.1 的值）。P／S 主腿為股權口徑（不加淨現金、無非控制權益）；分母為目標時點 TTM 營收（FY2027E 與 FY2028E 按日數加權）。
- v0.2.1 變更（工作單 G6）：PS／PSG 相對值已含利潤率差，不能套在 P／FCF；Comparable 主腿改 P／S 口徑＝同業中位 P／S（relmult_hist 最近一季，季末 2026-06-30，相對現價日逾期）× rel_mult（geo_base 0.725）× 目標日 TTM 營收 ÷ 評價股數。P／FCF Comparable 改對照列（rel_mult＝1.00 與本公司現行相對值兩讀數），不進加權。Sens 四案（ps_only、psg_only、sens.low、sens.high）套在 P／S 主腿。
- 合理性關卡（G3 第 3 條）診斷：加權目標價 158.68、上檔空間 -27.2%，評等放空（僅略低於 -25% 門檻；Sens 的 psg_only 與 sens.high 案為中立）。DCF 目標價 126.54 約為現價的 58%（高於 25%），EV 為正，評等非強力放空，FY2027 營收與指引差距不到 1%，營業利益率高於指引約 0.7 個百分點，GAAP EPS 低於指引約 10%（股數與業外淨額為 Assumed）。P／S 腿 190.8 高於 DCF，但兩者皆低於現價；評等落點靠近門檻，結論對 rel_mult 與同業 P／S 敏感。無須阻擋發布。
- v0.1.1 變更：每任務 token 改公司揭露累計 token÷累計 AWU（推得，Derived／Interested-party）；ptok 標類比；Astra 案為上界（只動成本），另有價格與成本聯動與模型組合路由案；股數加 RSU 與選擇權庫藏股法；Comparable 改有槓桿 P／FCF（TTM）；營業費用與銷貨成本占比基準持平、收斂只放 Sens；淨債務只計融資租賃；Val 加收支關卡三列；Recon 加反解。
- 同業倍數取自 StockAnalysis（改以 P／FCF 推導），日期為 2026-10-26 至 2026-11-13、家數 6，未取得 NTM 口徑；無風險利率與 5 家同業倍數日期逾期；不做迴歸，以調整係數作敏感度。
- Q2 10-Q 封面股數與 RSU、選擇權、融資租賃（來自協調者核對，未讀原檔）、向模型商購買的價格層級、token 價格路徑：Assumed／類比，待 AI供應鏈 登錄。
- 10-K FY2026 經第三方鏡像讀取、FY2022–FY2023 分雲值為第二來源，待回 SEC 原檔核對。
