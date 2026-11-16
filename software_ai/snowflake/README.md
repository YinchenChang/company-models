# software_ai/snowflake/ — Snowflake（SNOW）個股模型 v0.1（原型 U：按用量計價）

**本版未做先驗**（Andy 2026-11-16 23:27 指示）。工作單：`software_ai/docs/workorders/20261116_第一批_研究包與單位經濟.md` G 節與 G2 節。

**建置：** `bash software_ai/snowflake/make_v01.sh`（`SENS=1` 重跑敏感度與反解）。成品 `dist/20261116_Snowflake個股模型_v0.1.xlsx`／`.html`。投資建議、目標價與各數字見成品 README 頁與 HTML（由公式組字，本檔不手寫）。

**命題：** AI 進入按用量計價的資料平台後，產品營收＝核心平台消耗＋AI 工作負載；AI 淨效果＝AI 營收增量 − 核心被替代的侵蝕 − 推論成本。AI 營收＝基準消耗 B ×（代理工作負載乘數 w ＋ 呼叫數 j × 每次推論成本 ×（1＋轉售加成 U_mk））；推論成本只讀 Tokenomics v22.0 快照（`data/tokenomics_snapshot_v22.0.json`，任務 2「推理聊天」，市場價），變數名稱沿用 `software_ai/unit_economics`（`U_gm0`、`U_w26`／`U_w32`、`U_e26`／`U_e32`、`U_j26`／`U_j32`、`U_mk`、`U_tier`、`g_px`、`g_tok`、`ovh`）。命題模組在本模型 In 頁與 Seg 頁自建，v0.2 改讀 unit_economics。

**結構：** README、In、Q（法定申報歷史）、Seg（命題與營收推導）、PL（標準損益表、分部營收與對帳、三情境）、DCF、Val、RevDCF、Recon、Sens、Sources。損益表為非 GAAP 口徑（公司與賣方共識口徑），DCF 的 EBIT 扣除股票薪酬（SBC 不加回）；GAAP 為參考列。

**情境：** A 樂觀／B 基準／C 保守，在 AI 淨效果分化：`U_w32`（AI 工作負載）、`U_e32`（核心被替代）、`U_mk`（AI 轉售加成＝AI 毛利）、`sbc32`（GAAP 調整路徑：股票薪酬占營收）。依據與標記見 In 頁。

**檔案：** `build_snowflake.py`（Excel 產生器，全部活公式）、`sens_snowflake.py`（敏感度與反解，整本重算）、`gen_html_json.py`、`segments.json`、`html.json`、`make_v01.sh`、`data/`（研究包 `research_snowflake.md`、`facts_snowflake.csv`、Tokenomics 快照、`sens_v01_20261116.json`、`kreq_v01.json`）、`dist/`。

## 不適用與未做
- Meta 的算力、Tokenomics 自用 AI 讀法 2、Wattonomics 對照：不適用：Snowflake 不自建算力，推論成本為向模型商購買。
- 三層基準協調（Coord 頁、C0、兩讀數）：不適用：非 Siliconomics 覆蓋。
- 先驗：未做。
- 分部營業利益：無（單一營運分部）。
- 展示層數字一致性檢查（`check_display.py`、`display_check.json`、版本表）：本版未建，下次改版補。
- 以 `tools/fcff_dcf.py` 的寫表函式產生 DCF：本版未用，DCF 以與 Salesforce v0.1 相同的自含公式實作（FCFF、期中折現、永續價值含 RONIC），欄位與標準規格一致；下次改版可改接。
- 相對倍數調整（`tools/relative_multiple.py`）：未用：`tools/data/fwd_mult_hist.csv` 無 SNOW 與同業，調整係數為 1.00（未調整：缺歷史前瞻倍數），以 `cm_adj` 作敏感度。

## 已知限制
- Q3 FY27 實績尚未公布；FY2027 Q3、Q4 為假設（`q3g`、`q4g`）。
- 費用率與毛利率依規格持平於最近四季實績，不含公司預期的槓桿，故 FY2027 非 GAAP 營業利益率低於公司指引；收斂只放 Sens（`opx_dr`、`gm_dr`）。
- 10-Q Note 12 的在外 RSU、選擇權與沒收率、可轉債轉換價、融資租賃：未找到；稀釋加項 `sh_dil` 為由公司指引推得的代理值。
- AI 營收占比、AI 毛利率、每任務 token 數、呼叫數：公司未揭露，為 Assumed／Analogy；token 價格路徑待 AI供應鏈 登錄。
- 同業前瞻 P/S 取自 StockAnalysis，期間不一，部分逾期；SNOW 與 Blume β 只有單一來源。
