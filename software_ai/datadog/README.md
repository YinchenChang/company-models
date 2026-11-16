# Datadog（DDOG）個股模型 v0.1

資料截止日 2026-11-16（Q3 2026 財報後；會計年度＝日曆年）。**本版沒有做先驗**（未預先設定任何投資觀點，結論全由 Assumed 輸入與公開資料推出）。

## 命題
原型 U（按用量計價）單位經濟：營收＝客戶用量 × 單價。平台營收拆成三條命題子列（非官方分部；官方只有單一應報導部門「Datadog 平台」）：非 AI 原生客戶用量、AI 原生客戶用量（含單價讓利與最大客戶事件型尾端）、AI 產品 credits。AI 淨效果＝AI 營收增量 − AI 原生客戶價格讓利 − 推論成本（Tokenomics 快照 v22.0，價格只讀不估）。單位成本隨用量，價格讓利直接壓低毛利率。

## 情境分化依據
A／B／C 只在四個 AI 驅動分化：非 AI 客戶 FY2027 用量成長（27%／24%／18%，依近五季非 AI 成長 high 20s 的延續或減速）、AI 原生客戶用量成長（45%／30%／5%，依代理工作負載放大 vs 最大客戶停留最低承諾）、AI 原生客戶單價變動（−5%／−10%／−15%，依提前續約讓利）、AI credits FY2027 營收（180／110／60 百萬美元，公司未揭露，Assumed）。權重 25／50／25。

## 估值
DCF 為標準 FCFF 到 FY2032（股票薪酬視為費用、期中折現、RONIC＝WACC），Comparable 為同業 Forward P/S 與 P/E 中位數（Analogy）× 目標時點 NTM；兩法各 50%。結果數字見下方版本表與 xlsx／HTML。**DCF 與 Comparable 落差極大**（落差分解在 Val 頁 0.4：股票薪酬與倍數成長殘差），評等對 DCF 口徑高度敏感；發布前合理性關卡（工作單 G3）的診斷結論寫在 xlsx README 頁與 Recon 第 7 節。

## 共通必修 G2／G3 落實狀況
1. 每任務 token 由公司揭露校準（credits 價格 × 耗用 credits × 成本占比 ÷ 校準價格）；聊天基數 × IF_TaskTokMult 只列 Recon 第 5 節對照；token 價格層級標 Analogy。
2. token 價格高案標「上界」；另列價格與成本聯動、模型路由兩個組合案（Sens）。
3. 比率基準持平於近四季；歷史有證據者（股票薪酬占營收：FY2023 22.7% → 近四季 20.1%）依斜率延伸並設下限，Sens 低案為全期持平。non-GAAP 費用率三項合計歷史無一致下降（56.9%／58.6%／56.2%），故持平。
4. 股數取 10-Q 封面；含 RSU 與選擇權庫藏股法；2029 可轉換票據預設視為債務（現價低於轉換價），if-converted 為 Sens。
5. 淨債務只計票據本金與融資租賃；營業租賃不計。
6. Comparable 倍數標分母期間（資料商未載明）與會計準則（US GAAP；租賃處理未找到）；揭露本公司現行倍數 ÷ 同業中位數與隱含重估。同業歷史相對倍數取不到，rel_m＝1.00 並列為承重輸入。
7. Val 頁有「估值評等」「入不敷出判定用」「投資建議」三列；`check_rating_gate.py --strict` 已跑。
8. README 頁文字數字由公式組字；DCF 與 Comparable 落差以公式分解。
9. 主導變數整本重算反解寫在 Recon 第 6 節（rel_m：Comparable 需要的相對倍數；DCF 一腿無單一營運變數可支持現價）。
10. rf、同業倍數與共識日期逾期標示在 Val 0.6（MDB、DT、ESTC 倍數為 2026-11-06，逾期）。
G3：營運資金比率取歷史三期平均（遞延收入為現金流入）；合理性關卡七項判定在 Recon 第 7 節；DCF 與公司自由現金流對帳在 Recon 7b。

## 不適用項目
- 不適用：Meta 的算力、Tokenomics 自用 AI 讀法 2、Wattonomics 對照——Datadog 不自建 GPU 算力，只向模型商購買推論。
- 不適用：三層協調（Coord 頁、C0、兩讀數、Siliconomics）——非 Siliconomics 覆蓋；對帳中兩讀數目標價亦不適用。

## 版本註記
v0.2 起改讀 AI 軟體專案的 unit_economics 模型輸出，取代本版在 In 頁自建的單位經濟。

## 目錄與建置
- `build_datadog.py`（＋`build_datadog_p2/p3/p4.py` 為同一建置的分段）建模型；`facts_datadog.py` 事實總帳；`sens_datadog.py` 逐案重算敏感度與反解；`gen_html_json.py` 產 html.json；`check_datadog.py` 專屬核對；`segments.json`、`display_check.json` 設定。
- `data/`：研究包（research_datadog.md、facts_datadog.csv、segments.json 於上層）、Tokenomics 快照、敏感度結果、診斷結論。
- 建置：`SENS=1 bash software_ai/datadog/make_datadog.sh`（不加 SENS=1 沿用既有敏感度 JSON）。成品在 `dist/`（xlsx 與 html）。公開 repo，不含個人情境。
- 研究包的 `segments.json` 位於本資料夾（與 `data/` 並列）。

## 版本表（由成品產生，不手寫）
<!-- version-table: dist/20261116_Datadog個股模型_v0.1.xlsx -->
<!-- /version-table -->
