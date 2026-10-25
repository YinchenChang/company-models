# markets/invest_map — AI 投資地圖

把算力收支（Tokenomics、雲端／neocloud／模型商模型）與 AI半導體（各公司模型、候選母體）的研究，放回 AI 供應鏈，呈現三個視角：

1. **供應鏈地圖**：每 1 GW 的資金流（Tokenomics 每 GW 資本支出、持有成本、P×Q 營收、回本比值）＋各層次產業與公司。
2. **產業地位**：研究覆蓋公司的地位分數排序。
3. **可投資性**：模型隱含價 ÷ 同日收盤價，分「機會／觀察／偏貴」。

## 檔案

| 檔案 | 內容 |
|---|---|
| `extract_models.py` | 從各公司 `dist/*.xlsx` 抽估值欄位，每值記錄檔案與儲存格 → `data/models_extract.json` |
| `data/tokenomics_snapshot_v*.json` | Tokenomics 官方取數快照（`tools/tokenomics/import_tokenomics.py`；名稱清單 `data/tokenomics_names.txt`） |
| `data/universe_ai_semis_v*.csv`、`data/segments_ai_semis_v*.csv` | AI半導體 候選母體與次產業排序的抄錄（權威在 AI半導體 Project Drive） |
| `data/prices.csv`、`data/prices_*_sources.md` | 同日收盤價與來源 |
| `build_xlsx.py` | 產生資料活頁簿 `dist/<日期>_AI投資地圖_資料_v*.xlsx`（**事實與判斷的權威位置**；黃底＝判斷輸入，重建時讀回） |
| `build_html.py`、`template.html` | LibreOffice 重算活頁簿 → 產生 `dist/index.html`（網站只呈現值，不計算）與 `dist/data.json` |

## 重建

```
python3 markets/invest_map/extract_models.py [--avgo <AVGO 成品 xlsx>]
python3 markets/invest_map/build_xlsx.py
python3 markets/invest_map/build_html.py
```

## 固定規格（2026-10-25，Andy 要求）

- 三個視角：供應鏈（依 AI 產業供應鏈）、個股產業地位排序、個股可投資性（考慮評價、估值與現價）。
- 風格：專業投資機構研究報告風格，不用休閒網站風格。
- 依據 Tokenomics：每 GW 金額只讀 Tokenomics 快照；公司估值只讀各公司模型成品。
- 可投資性一律用同一日收盤價；模型價日與比較日差距超過 20% 時標示「模型尚未反映期間變化」。
- 網站不做計算；判斷（分級門檻、權重、地位分項）只在活頁簿黃底欄修改。
- **數字只用本專案檢核過的資料**（Andy 2026-10-25 10:25）：網站上的數字只能出自算力收支或 AI半導體 的模型與比對檔；網路只用於新聞（近期重大事件）、管理層身分與即時股價。本專案沒有的數字寫「本專案未建立」。公司檔案規格見 `profiles/00_格式規格.md`。
- **公司檔案**（Andy 2026-10-25 10:22）：每家含管理層、市場定位與產業位置、在做什麼、主要產品與營收占比、同業與競爭、近期重大事件、需要觀察的事項；以右側面板分頁呈現，不整片攤開。
- **發布**（Andy 2026-10-25 10:24）：放 GitHub Pages（`gh-pages` 分支，`publish_pages.sh`），公司內可直接開啟；不再用 Streamlit。

## 預設判斷（Andy 未回覆時照此；可改）

- 中心價：半導體 P/E 類取「十年倍數規則值」與「Andy 先驗合理 P/E」兩者較低者；規則值不適用者只用 Andy 合理 P/E（標替代口徑）；記憶體 P/B 規則值為主、P/E 為對照；雲端與 neocloud 用模型加權目標價；MiniMax 用反向檢驗折現價值。
- 分級：中心 ≥ +20% 且下檔 ≥ −25% → 機會；中心 < −10% → 偏貴；其餘 → 觀察。
- 地位分數＝0.40 × 不可替代性＋0.35 × 市場位置＋0.25 × 價值鏈分量。

## 版本

- v0.3（2026-10-25）：公司檔案新增「我們的模型」分頁（Andy 10:45 要求）：嵌入各公司模型成品 HTML（與抽數 xlsx 同一版次），附「在新分頁開啟」；`publish_pages.sh` 把 22 份模型 HTML 複製到 gh-pages 的 `models/<資料夾>/`。AVGO 改讀已合併的 `avgo/dist/`（PR #107）；MAG 三家改讀 v0.2.2（數值不變）。頁首加 `<meta charset="utf-8">`（不依賴伺服器宣告編碼）。

- v0.2（2026-10-25）：22 家公司檔案（profiles/）、公司檔案面板與總覽分頁、GitHub Pages 發布腳本；台達電改用模型 10-23 收盤。

- v0.1（2026-10-25）：22 檔上市公司估值、22 檔地位、71 個地圖節點；Tokenomics v21.4；股價 2026-10-23（台積電、廣達 10-22）。
