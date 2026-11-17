# gfs/：格芯（GlobalFoundries，NASDAQ: GFS）個股模型 v0.1.0

AI 半導體 Project 代工層：成熟與特殊製程晶圓代工（美國、新加坡、德國廠）。樣板 `umc/`（成熟製程代工，量 × 價 × 稼動率）與 `vis/`。美元報表、美元股價，匯率恆為 1；單位 US$B，晶圓 kWSE（千片，300 mm 約當）。會計準則 IFRS（公司 20-F）。

| 檔案 | 內容 |
|---|---|
| `build_gfs.py` | 產生器（上半部為公司設定 C，下半部為與 `umc/build_umc.py` 相同的共用建置程式，格芯改動處見下） |
| `gfs_post.py` | 後處理（只新增頁與列）：Hist 頁、PL 頁末標準損益表與分部營收推導、RevDCF 頁、Val 0.7 並列四個目標價、Sens 頁末加權目標價相對基準、Coord 頁 C0（公司實績）、README v1.3 白話列、Sources 新增來源 |
| `alt_run_gfs.py` | 暫代 `tools/coord/alt_run.py`（逐行相同，只把建檔指令寫死）；`.github/workflows/ai-semis-build.yml` 尚無 `gfs` target，**待主串加入 workflow target 後改用 `python3 tools/coord/alt_run.py gfs`** |
| `coord_alt.json` | 另一側讀數整本重建快照（Siliconomics v0.20；Wattonomics v0.3.6） |
| `gen_html_json.py` → `html_src.json`／`segments.json` → `html.json` | HTML 設定、分部宣告、轉換結果（勿手改 `html.json`） |
| `dist/20261117_格芯個股模型_v0.1.0.xlsx`、`.html` | 成品（LibreOffice 重算，錯誤格 0） |

工作表：README、In、Q、Units、PL、Val、DCF、RevDCF、Recon、Sens、Sources、Coord、Hist。Sens 頁是頁內逐案整本重算（與 `umc/` 同做法），所以沒有另外的 `sens_gfs.py` 快照腳本。

## 結論（模型輸出，2026-11-13 收盤價 41.47 美元）

加權目標價 30.9 美元，上檔空間 −25.4%，評等放空（DCF 15.57、Comparable 相對倍數調整後 46.27；DCF 權重 50%；情境權重 A／B／C＝0.3／0.5／0.2）。另一側讀數加權目標價 29.4（放空）。**上檔空間 −25.4% 剛好跨過 −25% 線**，稼動率或售價小幅調整即可轉為中立，見待判斷。

## 模型結構

- 製造服務營收＝產能 × 稼動率 × 美元平均售價（3Q26：出貨 647 千片、稼動率公司只說「高 80%」，取 88%）。
- 矽光子營收（公司 3Q26 法說：2026 年約 5 億美元、2027 年底年化 10 億美元）：FY2028 起乘 Siliconomics 顆數情境係數（`um28`／`um29`）。
- 技術服務營收（MIPS、ARC IP 與設計服務；公司稱約占營收 12%）：絕對金額情境。
- FY2026 錨點＝1Q–3Q26 實績＋4Q26 公司指引（營收 20 億美元、IFRS 毛利率 30.7%）；FY2027–29 三情境明確期；FY2030–32 收斂。
- 3Q26 IFRS 淨損來自 4.41 億美元非現金股份發行（美國政府）；營業利益 2.30 億美元為正。FY2027 起完全稀釋股數 6.00 億股（公司指引）。

## 讀 Siliconomics 變數（只讀現行 dist 基準 v0.20）

`SIF_UnitsCoef_Base`（um28／um29，顆數情境係數）、`SIF_WTAnchor_Base`（wta_27／28／29，Sens 錨定案例）、`SIF_RecoFlag_<年>`、`SIF_DeployGW_AltRatio_<年>`（兩讀數縮放）、`SIF_UnitsPlat_Total_Base`（Coord 層間參考）、設備層 `eq_g_oth_fd`（Coord C0 參考，群組口徑）。匯率不適用。

## 建議補進 Siliconomics（缺口；本模型以 Assumed／Analogy 自建）

1. 成熟與特殊製程代工行業層：稼動率、美元平均售價年增、產能年增、市場規模與格芯份額（本模型 `ut27`–`ut29`、`asp27`–`asp29`、`capg27`–`capg29` 為公司自有 Assumed）。
2. 矽光子（資料中心光互連）營收起點與年增、晶圓代工端的顆數映射（本模型 `ai_26`、`aig27`、`aig_L`）。
3. 美國／歐洲在地製造的政府補貼與客戶預付款口徑（影響資本支出與營運資金）。
4. 代工資本支出強度（占營收；本模型 `fc_cx`＝25%／35%／45%）。

## 待判斷（預設值、標籤、區間；不需停下等人）

| 項目 | 預設 | 標籤與區間 | 影響（加權目標價） |
|---|---|---|---|
| 評等邊界 | 上檔空間 −25.4%＝放空 | 規則固定；−25% 以上為中立 | 售價水準 +0.2% 或稼動率 +0.3 個百分點即跨線 |
| 稼動率（公司只說高 80%） | 3Q26 取 88%；FY2027–29 B 取 89／89／90% | Assumed；A 92／93／93%、C 84／82／82%；Sens ±3 個百分點 | ±6% |
| 製造服務售價年增 | FY2027 B +5%（公司：混合 ASP 與組合約 +10%，矽光子另列） | Assumed；A +8%、C +2%；Sens ±3% | ±14% |
| 矽光子營收 | FY2026 5 億美元；FY2027 年增 B 60% | Interested-party（起點）／Assumed（年增 40–80%）；Sens ±30% | ±7% |
| 技術服務營收 | FY2027–29 B 1.02／1.08／1.14 十億美元 | Assumed（情境 A 1.10–1.30、C 0.92 持平） | 見 Sens 新業務邊際貢獻率 ±3% |
| 資本支出強度 | 明確期 35% 營收（公司指引 2027 年約 40 億美元；含補貼前總額未揭露） | Assumed；25%／45% | 約 +13%／−12%（DCF 最大單一變數） |
| 完全稀釋股數 | FY2027 起 6.00 億股 | Interested-party；0.57／0.62 十億股 | 約 +5%／−3% |
| β | Yahoo 單一來源 1.85 去槓桿後重新加槓桿約 1.59 | Interested-party；±0.3 | +4.3%／−3.3% |
| ERP／國家風險溢酬 | 5%／0 | Assumed；4%／6% | 同 β |
| 無風險利率 | 5.00%（2026-10-23；11-13 殖利率頁無法讀取，沿用） | Verified（日期較舊）；稅前債務成本 rf＋1% | WACC ±1pt：+3.1%／−2.6% |
| 稅率 | FY2027 起 17% | Assumed；13%／21% | ±5.4% |
| DCF 與 Comparable 差距 | DCF 15.57 vs Comparable 46.27 | Comparable 對 IFRS EPS 套非 GAAP 同業倍數（口徑偏高）；DCF 終值隱含 EV／EBITDA 偏低（重資本期 FCFF 薄） | DCF 權重 0.3／0.7：+19.9%／−19.9% |
| 相對倍數調整 | 不調整（係數 1.00） | 格芯 Bloomberg 前瞻 P/E 只有五年，不足七年門檻 | — |
| 同業組（Comparable） | 台積電、聯電、高塔半導體、德州儀器、中芯國際（`tools/target_summary` 尚無 gfs 條目，以建置腳本內名單代替，不改 tools） | Analogy | — |
| 持股與庫藏股 | Mubadala 約七成持股、庫藏股後續授權不建模；股數只用公司指引 6.00 億股（含美國政府約 7%） | Interested-party（持股比例見 profile，第三方）；Sources 頁已標待補 | 稀釋區間見股數列 |
| 歷史損益 | FY2022–FY2025 取 stockanalysis 對 20-F 的轉錄值，交叉核對公司 6-K | Verified（第三方轉錄） | — |

## 檢查結果（2026-11-17）

| 檢查 | 結果 |
|---|---|
| LibreOffice 重算錯誤格（#REF!／#VALUE!／#DIV/0! 等） | 0 |
| `check_company_html gfs` | 全部通過 |
| `check_layout_v13`（含 `--xlsx`） | 通過；全頁參數代號警告 3 項（fcff_dcf、g_cap 等，與 `umc/` 相同，不判失敗） |
| `check_sens_coverage --require-keycol` | 通過（20 個有區間的鍵皆有案例；為此新增資本支出 25%／45% 與股數 0.57／0.62 兩組案例） |
| `segment_check/check_segments.py gfs` | 符合 |
| `baseline_coordination/check.py` | 0 項不符、8 項提醒（全為 Siliconomics 層級既有提醒） |
| `check_rating_gate`、`check_display` | 不適用（無收支關卡列、無 `display_check.json`） |
| `model_governance/check_governance.py` | **失敗 9 項**：Sens 頁無 build_id／模型狀態列、In／Q 頁無「來源誘因」欄、Q 頁 5 格標記為複合標記。這是 `umc/` 樣板系列同樣的既有狀態（對 `umc/dist` 跑同樣失敗），該檢查針對較新的 SpaceX 系列版型；若要通過須改共用建置程式或 `tools/`，本分支不動，請主串決定 |
| `markets/invest_map/check_profiles.py` | 通過（主文無殘留出處標記） |

## 資料標籤與出處

輸入以 [Verified]／[Interested-party]／[Analogy]／[Assumed]／[Derived] 標在 In 頁，出處在 In 頁與 Sources 頁。財報原文：公司 6-K 3Q26 新聞稿（2026-11-04）、3Q26 法說逐字稿、1Q26／2Q26 新聞稿、20-F。價格只取 `markets/invest_map/data/closes_20261020_20261113.csv` 與 `closes_extra_20261113.csv`（2026-11-13 收盤）；Yahoo 快照（共識、β、EPS）只作交叉檢查，不一致處見 Recon 頁。

## 工具與整合備註

- `tools/target_summary/universe.csv`／`peers.csv` 尚無 gfs：建置腳本在記憶體內補入（`RMX._universe`／`RMX._peers`），未改 tools。
- `siliconomics/builder/seed/coord_meta_v*.json` 的 `hardware_companies` 尚未列入 gfs：**待主串加入 coord_meta**（並加 workflow target）；在此之前，Coord 頁與 `coord_alt.json` 由 `alt_run_gfs.py` 產生。
- 地圖抽取程式、`prices.csv`、`model_status.csv` 未改（主串整合）。
