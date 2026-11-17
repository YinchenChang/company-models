# Naver（韓國 035420）個股模型 v0.2.1

資料截止日 2026-11-16（最新財報 3Q26）。成品日 2026-11-17。本版沒有做先驗（未預先設定任何投資觀點，結論全由資料與 Assumed 輸入推出）。金額十億韓元，K-IFRS，曆年制。

## v0.2.1 變更（2026-11-17，工作單 G6）
- 口徑一致：歷史 PS 與 PSG 相對倍數已含利潤率差，不能再套前瞻本益比。Comparable 主腿改為 P/S 口徑：同業中位 P/S（`data/ps_peers_naver.json`，取自 `software_ai/data/relmult_hist` 各家最近一季，季末日逐家標示，逾期者旗標，Derived）x 相對倍數（geo_base）x 目標日 TTM 營收（FY2026E 與 FY2027E 依已過天數加權）÷ 稀釋股數，三情境加權。
- 前瞻本益比 Comparable（分部 NTM NOPAT x 同業前瞻本益比）改為對照列：相對倍數 1.00 與本公司現行比值兩個讀數，不進加權目標價。
- Sens 保留只用 PS、只用 PSG、低案、高案四案，套用在 P/S 主腿；落差分解與 RevDCF 改為 P/S 口徑。
- 舊版 v0.2 成品移至 `dist/_待刪除/`。

## v0.2 變更（2026-11-17，工作單 G4、G5）
- 相對倍數 rel_mult 改讀 `data/relmult_naver.json`（來源 `software_ai/data/relmult_hist/out/naver.json` 的 geo_base，Derived）：歷史 PS 與 PSG 兩讀數幾何平均；Sens 加只用 PS、只用 PSG、sens.low、sens.high 四案。Comparable 用前瞻本益比，以 PS／PSG 相對值近似套用。兩讀數差異大於 30%，原因（本公司 TTM 營收成長低於同業中位，成長比 0.72）寫在模型 README 與 HTML。
- 無風險利率統一：韓國 10 年期國債取 2026-11-13 收盤（Trading Economics，頁面載 11-16 為 3.986%、單日 −0.001，回推 3.99%），日期寫入 In 頁。
- Recon 與 HTML 估值章並列 DCF（股票薪酬加回）讀數，只作對照、不進加權目標價；Naver 股票薪酬金額未找到，占營收比 Assumed 0，故讀數與標準 DCF 相同。
- 舊版 v0.1 成品移至 `dist/_待刪除/`。

## 命題
原型 A（按廣告計價）單位經濟：AI 暴露廣告 ＝ 查詢量 x 廣告負載 x 點擊率 x 每次點擊價格。Naver 不揭露查詢量與 ARPU，查詢量為 Assumed（寬區間），每次點擊價格由校準年營收反推。AI 淨效果 ＝ 增量（Q・s・(R−1)・u）− 侵蝕（−λ・Q・u）− 推論成本；損益平衡 R* ＝ 1 + 推論成本 ÷ 單位營收；PL 頁每情境有恆等式檢查列。Naver 自建 HyperCLOVA X，其自用推論成本放在命題模組（只涵蓋搜尋 AI 回答，其他自用未找到），只讀 Tokenomics 快照（`data/tokenomics_snapshot_v22.0.json`，Tokenomics v22.0，commit 79f6b30）的 IF_ 全成本名稱，Terra 級對應 HyperCLOVA X 屬類比，待 AI供應鏈 登錄；不自行估 token 價格。

## 情境分化依據
A／B／C 只在三個 AI 驅動分化：FY2032 AI 回答占查詢比例、AI 回答查詢的每查詢營收相對值 R、外部 AI 分流占比（數值見 In 頁情境驅動對照）。依據：公司只揭露 AI 對廣告成長的貢獻（1Q26 逾 50%、2Q26 逾 60%）與 AI 簡報廣告點擊率 +30%（Interested-party，非隨機對照），無查詢量與營收拆分；分化幅度為判斷，不是觀測值。AI 淨效果（營收層）FY2032E 三情境為正、近零、負，詳 Val 頁。其他分部成長三情境相同。權重 25／50／25（Assumed）。

## 不適用項目
- 不適用：算力、Wattonomics 與 Tokenomics 自用 AI 讀法 2 屬 Meta 範本項目；Naver 的自用 AI 成本在命題模組內。
- 不適用：非 Siliconomics 覆蓋。三層協調（Coord 頁、C0、兩個讀數）不做；對帳中兩讀數目標價亦不適用。

## 與標準規格的偏離與限制
1. 股數取第二來源，法定申報封面股數未取得；RSU、選擇權未找到，以稀釋係數 Assumed 取代庫藏股法。
2. 非控制權益拆分未找到，暫為 0；EPS 以合計淨利計。
3. 比率參數（Naver Platform、Financial Platform、Global 利益率與營業成本率）基準持平於近四季；歷史營業利益率無規模經濟趨勢可延伸（G3 第 1 條），收斂只放 Sens。
4. 折舊攤銷與資本支出歷史來自第二來源，公司未揭露年度資本支出。
5. Comparable 只有 5 家同業，不做迴歸；相對倍數為歷史 PS／PSG 幾何平均（Derived）。
6. AI Factory 資產由 Asset Co 持有，其資本支出與負債未入模型；NAVER Financial 與 Dunamu 合併未入模型。
7. 分部新口徑只有 FY2025 起，FY2023A–FY2024A 分部列空白（舊口徑見 Hist 頁）。

## 目錄與建置
- `build_naver.py` 建模型；`facts_naver.py` 事實總帳（寫出 `data/facts_naver.csv`）；`sens_naver.py` 逐案重算敏感度並反解；`gen_html_json.py` 產 html.json；`check_naver.py` Naver 專屬核對；`segments.json` 分部對帳設定。
- `data/`：研究包（research_naver.md、facts_naver.csv、segments.json）、Tokenomics 快照、敏感度結果。
- 建置：`SENS=1 bash software_ai/naver/make_naver.sh`（不加 SENS=1 沿用既有敏感度 JSON）。成品在 `dist/`（xlsx 與 html）。公開 repo，不含個人情境。
