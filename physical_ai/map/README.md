# physical_ai/map — 實體 AI 地圖

把 AI實體應用 專案的研究（六條線模型、個股模型、產業地位）放成一張地圖，比照 `markets/invest_map/`（AI 算力地圖）的方法與版面：

1. **六條線地圖**：1a Robotaxi、1b 智慧車、2a 工業機器人（含協作）、2b 人形機器人（硬體）、3 倉儲物流自動化（AMR、AS/RS）、6 共用零組件。每條線列線模型的重點數字（低／基準／高），下方列次產業與公司。
2. **產業地位**：100 列個股的地位分數與排名（可依線篩選）。
3. **可投資性**：有個股模型者的目標價、收盤、上檔空間、評等與收支關卡；Waymo、Tesla Robotaxi 以未上市分部列價值區間。

## 檔案

| 檔案 | 內容 |
|---|---|
| `extract_position.py` | 抄錄產業地位活頁簿（權威在 AI實體應用 專案 Drive）→ `data/segments.csv`、`positions.csv`、`watchlist.csv`、`params.csv`；移除工作情境欄、由「市場」推得「亞太掛牌」 |
| `data/prices.csv` | 同一日收盤（Yahoo Finance 日線）與來源網址 |
| `build_xlsx.py` | 產生資料活頁簿 `dist/<日期>_實體AI地圖_資料_v*.xlsx`（**事實與判斷的權威位置**；黃底＝判斷輸入：評等門檻、地位權重） |
| `build_html.py`、`template.html` | LibreOffice 重算活頁簿 → `dist/index.html`（資料內嵌、可單檔開啟）與 `dist/data.json`；版面 CSS 讀 `markets/invest_map/template.html` |
| `check_map.py` | 核對（見下） |

## 重建

```
python3 physical_ai/map/extract_position.py <產業地位 xlsx>   # 產業地位改版時
python3 physical_ai/map/build_xlsx.py
python3 physical_ai/map/build_html.py
python3 physical_ai/map/check_map.py                          # 全部通過才交付
```

## 固定規格

- 與 AI 算力地圖同方法：三個視角；專業研究報告風格；網站不做計算；本專案沒有的數字寫「本專案未建立」或「未估值」；主文不夾出處（虛線底數字，滑鼠停留顯示出處）。
- 地位分數＝0.40 × 不可替代性＋0.35 × 市場位置＋0.25 × 價值鏈分量；價值鏈分量的分母是人形機器人每台 BOM（零組件依 BOM 占比中值分級）。
- 數字來源：線指標只讀 `physical_ai/lines/*/dist` 成品（Robotaxi 讀 v0.2 下游取數）；個股目標價與收支關卡只讀個股模型成品 Val 頁；未上市分部讀 Waymo「估值」頁 `eq_C／eq_w／eq_A`、Tesla Robotaxi「給Tesla模型」頁 `RT_VAL_LO／BASE／HI`。
- **跨圖**（2026-11-16 定）：與 AI 算力地圖重複的公司照列、計入排名、標「跨圖」；主圖＝持有個股模型的專案所屬地圖，目標價與評等只讀主圖模型（`markets/invest_map/dist/data.json` 所引模型儲存格）。
- 上檔空間＝最新一份模型目標價 ÷ 最新收盤 − 1；收盤一律同一日。五級評等：> +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、< −50% 強力放空；不用「賣出」。另有收支不足、未估值、未上市。
- 收支關卡：評價為買進以上且模型 FY2026–FY2030 自由現金流扣股利（情境 B 合計）< 0 者改列「收支不足」；放空、強力放空、中立維持。
- 台股只顯示上檔空間與評等，不顯示目標價（各台股模型的規格）。
- 亞太掛牌標的標「亞太掛牌」；公開 repo 不放任何個人或工作情境。
- 不加造訪計數器。發布（GitHub Pages）屬站點維護方流程，未經明確指示不發布。

## 核對（`check_map.py`）

1. data.json 的線指標、目標價、收支關卡、分部價值逐一重新開出處檔比對；收盤＝`data/prices.csv`。
2. 上檔空間＝目標價 ÷ 收盤 − 1；評等依五級門檻重算一致；最終分級依收支關卡規則一致；台股不顯示目標價。
3. 地位分數與抄錄 csv 一致；分級只用允許集合。
4. 跨圖列都有主圖模型（AI 算力地圖）或分部估值。
5. data、dist、README、程式無個人或工作情境字樣。

## 未做

- 公司檔案：第一版只有估值摘要（地位依據、目標價、收盤、評等、收支關卡、模型檔）；管理層、主要產品與營收占比、同業與競爭、近期重大事件、客戶與股東未做。
- 其餘 92 列公司未建個股模型（顯示「未估值」或「未上市」）；第二批候選：Mobileye、地平線、小馬、優必選、宇樹、極智嘉、Symbotic、耐世特、THK、綠的諧波、SMC、Keyence。
- 跨圖公司的收盤日（2026-11-16）與 AI 算力地圖（2026-11-13）不同，評等可能不同（本版 Teradyne：本圖買進、算力地圖中立）。
- 「我們的模型」連結：未嵌入模型 HTML（physical_ai 個股模型尚無 HTML 一頁摘要）。
- 未發布到 GitHub Pages；入口與站點整合待站點維護方回覆。

## 版本

- v0.1（2026-11-17）：首版。線指標 29 項；個股 100 列（跨圖 7 列）；估值 12 家（本專案 6 家：Tesla v0.3、哈默納科、Nabtesco、上銀、Fanuc、安川；跨圖 6 家：NVIDIA、台積電、Amazon、Teradyne、台達電、鴻海）；未上市分部 2（Waymo v0.2、Tesla Robotaxi v0.3）。收盤 2026-11-16。評等（去重後）：強力買進 2（NVIDIA、鴻海）、買進 1（Teradyne）、中立 7、強力放空 2（Tesla、哈默納科）。
