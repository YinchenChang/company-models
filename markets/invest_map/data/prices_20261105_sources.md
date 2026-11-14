# 2026-11-05 收盤價來源說明（地圖 v0.6.30 起）

- AI半導體 60 家：`prices.csv` 取各公司最新個股模型（main `bd2b4979` 起，全部改用 2026-11-05 收盤）的現價儲存格，`source` 欄記錄模型檔與儲存格。地圖現價與模型現價同一數字，上檔空間＝模型加權目標價 ÷ 現價 − 1，與模型上檔空間相同。
- 算力收支公司（CoreWeave、Nebius、WhiteFiber、Oracle、Microsoft、Amazon、Alphabet、Meta、SpaceX、MiniMax、智譜）：本版不更新，沿用 2026-11-04 收盤（`prices_20261104_sources.md`）。

## 地圖 v0.6.38（2026-11-11 重建）補記

- 本次重建未取得 11-05 之後的收盤（本次工作環境無法連到 README 規定的收盤來源網站），收盤日維持：AI半導體 61 家 2026-11-05、算力收支公司 2026-11-04。
- **Oracle 改 2026-11-05 收盤 115.16**（原 11-04 123.73）：取自 `oracle/dist/20261115_Oracle個股模型_v0.5.2.xlsx` In!E5（模型出處：Yahoo Finance 日線 115.16，2026-11-06 擷取；StockAnalysis 115.19；CNBC 11-05「Oracle 跌近 7%」），與 AI半導體 同一收盤日。
- 來源欄改指現行模型檔（收盤值不變，儲存格值已逐一核對相同）：AMD v0.4.4 → v0.4.6（Val!B67）、Astera Labs v0.3.5 → v0.3.7（Val!B63）、南電 v0.3.3 → v0.3.4（Val!B53）、美超微 v0.3.5 → v0.3.6（Val!B53）、Arm v0.3.3 → v0.3.4（Val!B60 → Val!B64，版面移列）。
- Microsoft、Amazon、Alphabet、Meta、SpaceX 現行模型的現價亦為 2026-11-05 收盤，本版未改用（維持 11-04），待下次全面更新收盤時一併處理。

## 地圖 v0.6.41（2026-11-15 重建）補記

- **Nebius 改 2026-11-12 收盤 128.64**（原 11-04 141.52）、**WhiteFiber 改 2026-11-10 收盤 9.21**（原 11-04 9.15）：取自新模型 `nebius/dist/20261115_Nebius個股模型_v0.6.0.xlsx`、`whitefiber/dist/20261115_WhiteFiber個股模型_v0.4.0.xlsx` 的 Val!C4（In 頁 px），比 prices.csv 的 11-04 收盤新；作法與 v0.6.40 對 Meta、CoreWeave 相同（模型所載收盤較新時改用模型現價）。
- 其餘公司收盤不變；來源欄模型檔名改指現行版（儲存格值逐一核對相同；Lumentum 指 v0.5.6）。
