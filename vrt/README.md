# vrt — Vertiv（VRT）個股模型 v0.1.1（2026-11-21）

AI 供應鏈：電源與散熱（機房設施端）。命題＝Vertiv 資料中心設施營收＝Wattonomics 機房設施市場（機櫃外配電、UPS、冷卻）× Vertiv 份額。單位：十億美元、EPS 美元（GAAP 稀釋）；會計年度＝日曆年。

## 檔案
- `build_vrt.py`、`build_part2.py`、`build_part3.py`：建置（`python3 vrt/build_vrt.py <raw.xlsx>`，再以 LibreOffice 重算；part2／part3 由 build_vrt.py 載入）。
- `sens_vrt.py`、`sens.json`：業務主因敏感度快照（整本重建；In 頁改動後重跑，否則 Sens 頁模型狀態為 STALE）。
- `rowmap.json`、`gen_html_json.py`、`html.json`：HTML 設定（html.json 由產生器寫，勿手改）。
- `segments.json`：分部宣告（官方分部＝地區；產品線拆分為本模型推導子列）。
- `dist/`：xlsx（LibreOffice 重算）與 html，各一個（檔名日期 20261121）。

## 產業變數
機房設施市場讀 Wattonomics v0.3.6 快照（`wattonomics/data/wattonomics_snapshot_v0.3.6.json`，names `WIF_Spend_fac_dist／ups／cool／gen／sum_<年>`，sha256 acb92d6f…）。邊界：機櫃內（PSU、電源櫃、冷板、機櫃內歧管、BBU）屬 IT 端（台達電），不計入；機櫃外（UPS、配電、CDU、冷水機、CRAH、發電機組）屬設施端。Vertiv 可及市場＝配電＋UPS＋冷卻（CDU 歸冷卻；發電機組不計）。Vertiv 份額為 Derived（資料中心設施產品營收 ÷ 可及市場），區間進 Sens。不讀 Siliconomics（其電源與散熱為機櫃內口徑）。

## 來源與等級
- 公司原文（Verified）：3Q26 新聞稿與 10-Q、2Q26 與 4Q25 新聞稿、4Q23 新聞稿（見 xlsx Sources 頁）。1Q26 為 1H26 減 2Q26。
- Interested-party：4Q26 與 FY2026 指引（管理層）、共識（Yahoo 2026-11-17）、同業前瞻 EPS（沿用 delta 模型 Val 頁同業表）。
- Assumed：資料中心設施產品占營收比例、情境倍數、份額相對值、毛利率與費用假設等（In 頁每列標單一標記；Interested-party 另寫來源誘因）。
- 價格：`markets/invest_map/data/closes_20261109_20261120.csv` 2026-11-20 收盤（VRT 166.24 美元；v0.1.0 為 2026-11-13 收盤 202.07），未上網抓價。

## Yahoo 與公司原文差異（Yahoo 只作交叉核對；完整表見 xlsx Recon 頁）
營業利益 FY2023 公司 872.2／Yahoo 916.8、FY2024 1,367.4／1,382.0、FY2025 1,829.7／1,896.2（百萬美元）；3Q26 983.8／991.1；1Q26 440.1（1H26 減 2Q26）／433.6。營收與淨利吻合。Yahoo 總債務 3,356.9 含租賃，本模型債務與租賃分列。

## 待判斷（已決定並標範圍；列於 xlsx README 頁與 HTML 資料狀態章）
1. 資料中心設施產品占營收：FY2026 取 0.72（範圍 0.65–0.80）。
2. 情境倍數：設施市場 2028／2029 倍數 A 1.10／1.15、B 1.00／1.00、C 0.80／0.65（個股自訂；建議 Wattonomics 提供官方上下情境）。
3. 份額相對值：FY2027／28／29＝1.03／1.02／1.00（範圍 0.95–1.10）。FY2025 份額 15.2% 高於 FY2026 的 10.5%，視為認列落後市場支出。另一種解釋是 Vertiv 實際上在流失份額；基準不改，Sens 組合案例「份額隨認列落後回升到 FY2025 水準（FY2028–29 約 15%；rel 1.20／1.45／1.45）」的加權目標價約 192 美元（基準 142，+35.8%），兩種解釋待 Andy 判斷。
4. FY2030–FY2032 不讀 Wattonomics（新增容量下降）：營收年增收斂到 3%（2%–4%）、營業利益率收斂係數 0.85（0.75–0.95）。
5. UtilityInnovation 收購（2026-10-26，現金 14.5 億美元）：基準價值中性；Sens 另列純流出下緣。
6. β 單一來源（Yahoo 2.11，Blume 1.74），區間 1.45–2.05；WACC 約 13.3%，終值隱含 EV／EBITDA 約 5 倍，是 DCF（81 美元）遠低於同業倍數法（202 美元）的主因。
7. EPS 口徑：本模型 GAAP，共識與同業前瞻 EPS 為調整後或混合口徑，未調整。
8. 稀釋股數取 3Q26 稀釋加權 392.7M（流通 385.1M）。
9. 3Q26 backlog 與 book-to-bill：未找到（只有 FY2025 底 150 億美元、4Q25 約 2.9）。
10. 同業僅 Eaton、Schneider、台達電三家（peers.csv 尚無 vrt 條目，由主串加入）。

## 本版要點
- 現價與同業報價改 2026-11-20 收盤（工作單 px1120）：現價 202.07 → 166.24；同業前瞻 P/E 中位數 23.26 → 18.29（Eaton 352.97 → 311.33、Schneider 243.45 → 226.55 歐元、台達電 1,240 → 900 新台幣；前瞻 EPS 不變）；Comparable（調整後）202.25 → 159.07；加權目標價 141.57 → 120.21；上檔空間 −29.9% → −27.7%；評等 放空 → 放空（公式重算）。共識、rf、β、ERP、營運假設、情境權重不變。未換價同業：無。Sens 快照已重跑。
- 2026-11-17 獨立核對後修正：profile 出處儲存格錯位更正（Units!D8／E10／G10、Val!B19）；反向 DCF 現價隱含 2.86 倍改為「現價隱含值、非模型值；差距來自 WACC 13.3%、終值 EV／EBITDA 約 5.2 倍與 FY2028 後成長」，新增 RevDCF 列 13（現價隱含終值 EV／EBITDA 約 24.9 倍）；份額加註實際流失份額解釋並新增 Sens 組合案例。

## 建議補進 Siliconomics／Wattonomics
Wattonomics：設施市場官方上下情境；2030–2032 設施端更換與服務需求；機櫃內外邊界與每 GW 內容對照；支出與認列的時差。

## 主串處理（2026-11-17）
- Coord／alt 讀數：不適用（電力與機房範圍定案 2026-11-17）：機房設施供應商只讀 Wattonomics 廠房拆分，該數字只有 Wattonomics 一個權威層、沒有 Siliconomics 側的另一讀數，三層協調 C4 兩讀數不適用；與 Wattonomics 的差異走對帳單。 CI 建檔目標已由主串加入（2026-11-17）。
- 相對倍數調整以 `RMX.add_block` 直接接入（own＝VRT、同業 ETN／SU.PA／2308.TW）；universe.csv 補齊後可改用 `RMX.apply`。

## 檢查結果（2026-11-21，v0.1.1）
check_company_html、check_layout_v13 --xlsx、check_sens_coverage --require-keycol、check_governance --html、baseline_coordination（0 項不符）、check_segments、check_cached_values 皆通過；xlsx 無錯誤值。

## 檢查結果（2026-11-17，v0.1.0）
check_company_html（全部通過）、check_layout_v13 --xlsx（ok；全頁參數代號警告 2 筆來自共用模組列標籤）、check_sens_coverage --require-keycol（30 鍵皆涵蓋）、check_governance --html（通過）、baseline_coordination（0 項不符）、check_segments（符合）。check_display 不適用（無 display_check.json）。xlsx 無 #REF!／#VALUE!／#DIV/0!。
