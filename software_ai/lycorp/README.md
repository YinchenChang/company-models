# LY Corporation（東證 4689）個股模型 v0.2.0

資料截止日 2026-11-16（FY2026 第 2 季，4 月制）。本版沒有做先驗（未預先設定任何投資觀點，結論全由 Assumed 輸入與公開資料推出）。

## 命題
原型 A（按廣告計價）單位經濟：搜尋廣告 = 查詢量 x 廣告負載 x CTR x CPC。LY 不揭露查詢量與 ARPU，故以「查詢分流占比」與「AI 轉換率帶動的價格提升」兩個 Assumed 區間取代，另加官方帳號 AI Mode 增量與 Agent i 訂閱／廣告。推論成本只讀 Tokenomics 快照（`data/tokenomics_snapshot_v22.0.json`，只用 IF_／L1_／SRC_ 名稱），不自行估價。

## 情境分化依據
A／B／C 只在四個 AI 驅動分化：查詢分流占比（10%／20%／35%）、價格提升（15%／8%／2%）、AI Mode 增量（12%／6%／0%）、Agent i 付費占比（6%／3%／1%）。依據：公司只揭露 AI 功能使用者購前轉換 1.7 倍（非隨機對照）與 Agent i 計費起點，沒有量化 AI 營收；分化幅度為判斷，不是觀測值。FY2032E AI 淨效果（營業利益，十億日圓）A／B／C = +52.3／+10.4／-33.4。其他分部成長三情境相同。權重 25／50／25（Assumed）。

## 不適用項目
- 不適用：非 Siliconomics 覆蓋。三層協調（Coord 頁、C0、兩個讀數）不做；對帳中兩讀數目標價亦不適用。
- 不適用：LY 不自建算力，只向模型商購買推論。Tokenomics 自用 AI 讀法 2 與 Wattonomics 對照不做。

## 與標準規格的偏離
1. 永續成長率 2%（標準 3%）：JGB 10 年期 2.89%，3% 偏高。
2. beta 0.9 為 Assumed，未取兩來源 Blume 調整中位數。
3. 資產負債表取 2026-09-30，現價取 2026-11-16。
4. 非控制權益以帳面 x 1.5（Assumed）。
5. 同業倍數（Analogy）：Peers 頁 11 列，多數資料日逾期，會計準則與租賃處理多數未取得；本公司 5 年歷史相對倍數自算（財報年度值，非月資料）。
6. PayPay 銀行未模擬存放款，營運資金只以營收增量比率估計（歷史比率未取得）。
7. 股數：封面發行股數 − 庫藏 + 股票交付信託股（視同 RSU）；選擇權、可轉債、ASR 未取得，記 0。
8. 淨債務含全部 IFRS 16 租賃負債（與 FCFF 資本支出含使用權資產一致），租賃負債金額未單獨取得。

## G2／G3 共通必修落實（v0.1.1）
G2：(1) 公司未揭露 token 量與代理任務量，無法校準，維持 Assumed 並標明，token 價格層級標 Analogy；(2) ptok 高案標上界，另有價格與成本聯動、模型路由兩個組合案例；(3) 營業利益率與營業成本率基準持平於最近四季實績（15.1%），收斂只在 Sens；(4) 股數以封面為主來源；(5) 淨債務口徑明寫；(6) Comparable 為同業 TTM 倍數 x 目標日 TTM 分部 EBITDA，列同業準則與逾期，揭露現行倍數 / 同業中位數；(7) Val 頁三列（估值評等、入不敷出判定用、投資建議），make 腳本跑 check_rating_gate --strict；(8) README 日期、指引、落差歸因由公式組字，DCF 與 Comparable 落差以公式分解；(9) 反解寫在 Recon（支持現價所需營業利益率 13.6%）；(10) 逾期標示在 Peers 與 Val。
G3：(1) 費用率歷史無單調規模經濟證據（營業費用率 FY2023 至最近四季約 56 至 60% 先降後回），不延伸，Sens 保留高案；(2) 營運資金正值為流出，FCFF 以減項計；歷史比率未取得；(3) 合理性關卡：DCF/現價 1.32、EV 為正、評等中立、營收差距 -1.4% 皆未觸發；FY2026E EBITDA 對公司調整後指引 -14.8% 觸發，診斷為基準持平規則（H1 實績 17.8% 對 15.1%，約 62 十億日圓）加口徑差，是經濟判斷非建模錯誤（README 與 README 工作表 A42）；(4) rel_mult 自算：Media 0.74、Commerce 0.59、Strategic 0.83。

## v0.2 改動（G5）
1. 相對倍數改讀 `software_ai/data/relmult_hist/out/lycorp.json` 的 Media／Commerce／Strategic 分組 `geo_base`（PS 與 PSG 幾何平均，Derived；base_source 寫在 In 頁）：Media 0.58、Commerce 0.51、Strategic 1.31（v0.1.1 自算為 0.74、0.59、0.83）。Sens 加 ps_only、psg_only、sens.low、sens.high 四個組合案例（三分部同時改）。Comparable 倍數為 EV／EBITDA，以 PS／PSG 相對值近似套用。Strategic 兩讀數差異 119% 大於 30%：本公司 TTM 營收成長 8.1% 對同業中位 5.5%，成長比 1.46 使 PSG 讀數 1.94 遠高於 PS 讀數 0.89；同業僅 3 家，結論偏不穩。
2. 無風險利率統一取日本 10 年期 2.77%（2026-11-11，Trading Economics）；距現價基準日 3 個交易日，Val 頁標逾期。
3. Recon 與 HTML 並列 DCF（股票薪酬加回）對照讀數，不進加權目標價。LY 不揭露股票薪酬費用，以權益變動表股份基礎給付代理（營收約 0.05%），加回影響僅約 +3.4 日圓。
4. 合理性關卡重跑：FY2026E EBITDA 對指引 -14.8% 仍觸發，診斷同 v0.1.1（基準持平規則與口徑差異，經濟判斷非建模錯誤）。

## 版本註記
v0.2 起改讀 AI 軟體專案的 unit_economics 模型輸出，取代本版內建的單位經濟推導。

## 目錄與建置
- `build_lycorp.py` 建模型；`sens_lycorp.py` 逐案重算敏感度；`gen_html_json.py` 產 html.json；`check_lycorp.py` LY 專屬核對；`segments.json` 分部對帳設定。
- `data/`：研究包（research_lycorp.md、facts_lycorp.csv、segments.json）、Tokenomics 快照、敏感度結果。
- 建置：`SENS=1 bash software_ai/lycorp/make_lycorp.sh`（不加 SENS=1 沿用既有敏感度 JSON）。成品在 `dist/`（xlsx 與 html）。公開 repo，不含個人情境。
