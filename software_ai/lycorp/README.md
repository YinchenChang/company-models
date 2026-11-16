# LY Corporation（東證 4689）個股模型 v0.1.0

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
5. 未蒐集同業倍數：分部加總倍數 Media 9／Commerce 7／Strategic 11 為 Assumed。
6. PayPay 銀行未模擬存放款，營運資金只以營收增量比率估計。
7. 股數取流通股含庫藏 x (1+稀釋係數)，未做庫藏股法。

## 版本註記
v0.2 起改讀 AI 軟體專案的 unit_economics 模型輸出，取代本版內建的單位經濟推導。

## 目錄與建置
- `build_lycorp.py` 建模型；`sens_lycorp.py` 逐案重算敏感度；`gen_html_json.py` 產 html.json；`check_lycorp.py` LY 專屬核對；`segments.json` 分部對帳設定。
- `data/`：研究包（research_lycorp.md、facts_lycorp.csv、segments.json）、Tokenomics 快照、敏感度結果。
- 建置：`SENS=1 bash software_ai/lycorp/make_lycorp.sh`（不加 SENS=1 沿用既有敏感度 JSON）。成品在 `dist/`（xlsx 與 html）。公開 repo，不含個人情境。
