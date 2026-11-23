# physical_ai/map — 實體 AI 地圖

把 AI實體應用 專案的研究（六條線模型、個股模型、產業地位）放成一張地圖，比照 `markets/invest_map/`（AI 算力地圖）的方法與版面：

頁首有站際切換「算力｜實體 AI｜軟體 AI」（`<nav class="site-switch">`，與 AI 軟體地圖同一標記）。網址 https://yinchenchang.github.io/company-models/physical_ai/（v0.1 的 `physical/` 改為轉址頁）。

1. **六條線地圖**（版面同 AI 軟體地圖「產業地圖」：每列一條線，左欄線別、關鍵問題、覆蓋家數；右側上方為線模型重點數字，下方每格一個次產業與公司）：線 1 Robotaxi、2 智慧車、3 工業機器人（含協作）、4 人形機器人、5 倉儲物流自動化為應用端；線 6 共用零組件依功能分 6A 傳動與致動、6B 感知、6C 末端執行、6D 運算、電池與線束（分組在 `data/map_groups.csv`，判斷，只影響呈現）。
2. **產業地位**：100 列個股的地位分數與排名（可依 1–5、6A–6D 篩選）。
3. **可投資性**（v0.3 起比照 AI 算力地圖）：先列總結（各級家數與公司、收支關卡、模型現價日）→ 散布圖（橫軸地位分數、縱軸上檔空間）→ 個股表，可依「全部／強力買進／買進／中立／放空／強力放空／收支不足／未上市」篩選；Waymo、Tesla Robotaxi 以未上市分部列價值區間。
4. **公司檔案**：格式同 AI 算力地圖（`profiles/<代號>.md`，規格 `profiles/00_格式規格.md`）；跨圖公司的檔案在 AI 算力地圖。v0.3 起有「個股模型」分頁，嵌入模型成品的 HTML 版（`physical_ai/<公司>/dist/*.html`，由 `tools/company_html` 產生）；跨圖公司嵌入 AI 算力地圖的模型頁。

## 檔案

| 檔案 | 內容 |
|---|---|
| `extract_position.py` | 抄錄產業地位活頁簿（權威在 AI實體應用 專案 Drive）→ `data/segments.csv`、`positions.csv`、`watchlist.csv`、`params.csv`；移除工作情境欄、由「市場」推得「亞太掛牌」 |
| `data/prices.csv` | 同一日收盤（Yahoo Finance 日線）與來源網址；原始擷取檔在 `physical_ai/data/closes_*.csv`（GitHub Actions `px-closes-fetch` 以 yfinance 擷取，雲端工作環境連不到 Yahoo） |
| `build_xlsx.py` | 產生資料活頁簿 `dist/<日期>_實體AI地圖_資料_v*.xlsx`（**事實與判斷的權威位置**；黃底＝判斷輸入：評等門檻、地位權重） |
| `build_html.py`、`template.html` | LibreOffice 重算活頁簿 → `dist/index.html`（資料內嵌、可單檔開啟）與 `dist/data.json`；版面 CSS 讀 `markets/invest_map/template.html` |
| `profiles/*.md`、`profiles_md.py` | 公司檔案與解析程式（`profiles_md.py` 複製自 `markets/invest_map/`，同一解析規則） |
| `data/map_groups.csv` | 地圖分組：線 1–5、6A–6D 與次產業代碼 |
| `publish_physical_map.sh` | 發布到 gh-pages `physical_ai/`（追加 commit、不 force；`physical/` 改轉址）；v0.3 起一併把本專案個股模型的 HTML 與同版 xlsx 複製到 `physical_ai/models/<資料夾>/` |
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
- 不加造訪計數器。發布需明確指示（2026-11-17 08:10 已授權）；只動 `physical_ai/`、`physical/`。
- 線號（2026-11-17 起）：1–5 為應用端，6 為共用零組件；產業地位活頁簿「所屬線」原文仍為 1a／1b／2a／2b／3，地圖以分組表對照。

## 核對（`check_map.py`）

1. data.json 的線指標、目標價、收支關卡、分部價值逐一重新開出處檔比對；收盤＝`data/prices.csv`。
2. 上檔空間＝目標價 ÷ 收盤 − 1；評等依五級門檻重算一致；最終分級依收支關卡規則一致；台股不顯示目標價。
3. 地位分數與抄錄 csv 一致；分級只用允許集合。
4. 跨圖列都有主圖模型（AI 算力地圖）或分部估值。
5. 分組恰好涵蓋全部次產業、線號 1–6、站際切換正確、公司檔案章節齊全。
5a. （v0.3）本專案 8 份個股模型都有 HTML 且與 xlsx 同版；跨圖模型連結與 AI 算力地圖 data.json 一致；本專案模型現價與 `prices.csv` 同日同值；頁面含散布圖、分級篩選與「個股模型」分頁。
6. data、dist、profiles、README、程式無個人或工作情境字樣（私人字樣由環境變數 `PRIVATE_TERMS` 提供，不寫進程式）。

## 未做

- 公司檔案：見公司檔案頁的完成數；未完成者顯示「撰寫中」。
- 其餘公司未建個股模型（v0.4 時 66 列）（顯示「未估值」或「未上市」）；第二批候選：Mobileye、地平線、小馬、優必選、宇樹、極智嘉、Symbotic、耐世特、THK、綠的諧波、SMC、Keyence。

## 版本

- v0.1（2026-11-17）：首版。線指標 29 項；個股 100 列（跨圖 7 列）；估值 12 家（本專案 6 家：Tesla v0.3、哈默納科、Nabtesco、上銀、Fanuc、安川；跨圖 6 家：NVIDIA、台積電、Amazon、Teradyne、台達電、鴻海）；未上市分部 2（Waymo v0.2、Tesla Robotaxi v0.3）。收盤 2026-11-16。評等（去重後）：強力買進 2（NVIDIA、鴻海）、買進 1（Teradyne）、中立 7、強力放空 2（Tesla、哈默納科）。
- v0.2（2026-11-17）：版面改成與 AI 軟體地圖「產業地圖」同一排法（每列一條線、左欄線別、右側次產業格）；線號改 1–6，第 6 線共用零組件拆 6A 傳動與致動、6B 感知、6C 末端執行、6D 運算、電池與線束（`data/map_groups.csv`）；頁首加站際切換「算力｜實體 AI｜軟體 AI」；新增視角 04 公司檔案（89 家，格式同 AI 算力地圖；跨圖 7 家見 AI 算力地圖），經兩位未參與撰寫的代理核對本專案數字；發布位置改為 `physical_ai/`（`physical/` 轉址）。待補：部分公司近期事件、管理層與前十大股東（撰寫時網路搜尋額度用盡，檔內標「待補」）。
- v0.3（2026-11-23）：(1) **可投資性頁比照 AI 算力地圖**：總結、地位 × 上檔空間散布圖、分級篩選（全部／強力買進／買進／中立／放空／強力放空／收支不足／未上市）。(2) **公司檔案加「個股模型」分頁**：本專案 8 份模型（Tesla、哈默納科、Nabtesco、上銀、Fanuc、安川、Waymo、Tesla Robotaxi）新增 HTML 展示層並嵌入；跨圖 6 家嵌入 AI 算力地圖的模型頁。(3) **收盤改 2026-11-20**：六家上市公司模型的現價同步換成 2026-11-20 收盤並重算（各模型 patch 升版），與 AI 算力地圖同一收盤日。(4) 站際切換：AI 算力地圖、AI 軟體地圖原連 `physical/`（轉址頁，換頁時多跳一頁），已改直連 `physical_ai/`（PR #713，跨專案）。
- v0.4（2026-11-23）：接入第二批 12 份個股模型 v0.1（Mobileye、地平線、小馬智行、優必選、宇樹、極智嘉、Symbotic、耐世特、THK、綠的諧波、SMC、Keyence；現價 2026-11-20 收盤）。有估值 24 家（本專案 18、跨圖 6）。宇樹收盤為單一來源（Yahoo 收盤檔無此代號），待以交易所資料複核。

