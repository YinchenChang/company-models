# software_ai/salesforce/ — Salesforce（CRM）個股模型 v0.1（原型 S：按席位計價）

**本版未做先驗**（Andy 2026-11-16 23:27 指示）。工作單：`software_ai/docs/workorders/20261116_第一批_研究包與單位經濟.md` B 節與 G 節。

**建置：** `bash software_ai/salesforce/make_v01.sh`（`SENS=1` 重跑敏感度）。成品 `dist/20261116_Salesforce個股模型_v0.1.xlsx`／`.html`。投資建議、目標價與各數字見成品 README 頁與 HTML（由公式組字，本檔不手寫）。

**命題：** AI 進入席位型軟體後，每單位營收＝代理營收增量 − 席位被替代的侵蝕 − 推論成本。Agentforce Apps ＝ 席位型營收（席位 × 每席位年費，扣客戶人力替代）＋ 代理營收（任務數 × 每任務單價）；推論成本＝任務數 × 每任務 token × 市場價，只讀 Tokenomics v22.0 快照（`data/tokenomics_snapshot_v22.0.json`）。變數名稱沿用 `software_ai/unit_economics`，該模型 v0.1 完成前在本模型 In 頁自建（v0.2 改讀 unit_economics）。

**結構：** README、In、Q、Seg（命題與營收推導）、PL（標準損益表、分部營收與對帳、三情境）、DCF、Val、RevDCF、Recon、Sens、Sources。官方報告分部只有一個（10-K）；營收列取公司固定揭露的營收型態與新二分類產品線（Agentforce Apps；Data 360, Headless Platform, and Other）。

**情境：** A 樂觀／B 基準／C 保守，在 AI 淨效果上分化（席位替代率 `rho`、每任務動作量成長 `gt0`、每任務單價變動 `dp`），依據與標記見 In 頁。

**檔案：** `build_salesforce.py`（Excel 產生器，全部活公式）、`sens_salesforce.py`（敏感度，整本重算）、`gen_html_json.py`、`segments.json`、`html.json`、`make_v01.sh`、`data/`（研究包 `research_salesforce.md`、`facts_salesforce.csv`、Tokenomics 快照、敏感度結果）、`dist/`。

## 不適用與未做
- Meta 的算力、Tokenomics 自用 AI 讀法 2、Wattonomics 對照：不適用：Salesforce 不自建算力，推論成本為向模型商購買。
- 三層協調（Coord 頁、C0、兩讀數）：不適用：非 Siliconomics 覆蓋。
- 先驗：未做。
- 分部營業利益：無（單一營運分部）。Non-GAAP 預測：未建，只在 Recon 對照公司指引。
- 新二分類在 FY2026 全年與 Q3、Q4 FY26 的回溯值、Blume β 的第二來源、FY2028 營收共識、同日期賣方平均目標價：未找到。
- 展示層數字一致性檢查（`check_display.py`、`display_check.json`、版本表）：本版未建，下次改版補。

## 已知限制
- 同業倍數取自 StockAnalysis 的 EV／FCF，日期為 2026-10-26 至 2026-11-13、家數 6，未取得 NTM 口徑；不做迴歸，以調整係數作敏感度。
- 每任務 token 的聊天基數、Tokenomics 任務倍數對應哪一項、向模型商購買的價格層級、token 價格路徑：Assumed／類比，待 AI供應鏈 登錄。
- 10-K FY2026 經第三方鏡像讀取、FY2022–FY2023 分雲值為第二來源，待回 SEC 原檔核對。
