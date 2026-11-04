# 2026-11-03（週二）收盤價來源說明（地圖 v0.6.25）

`prices.csv` 每列的 `source` 欄記錄來源網站，`sources_count` 為取得同日收盤的獨立來源數。三組代理分別查美股、台股、其他市場（2026-11-04）。

## 取價原則
- 只採用頁面明示 11-03 日期的收盤；未標日期或日期較舊者不用。本日 stockanalysis 美股歷史頁停在 10-14／10-30，不可用；台股頁加 `&p=daily` 才取得新資料。
- 美股：massive.com 分列「At close」與盤後，以正規盤收盤為準；marketbeat 似記錄最後成交而非官方收盤（11-02 對照亦偏差數美分），兩者不一致時取 massive。
- 「cnyes(11/04頁昨收)」：cnyes 已換到 11-04 盤中，取其昨收；只作第二來源。
- AMD、Arista 於 11-03 盤後公布財報：用正規盤收盤（盤後 AMD 526.80、ANET 186.83 未用）。
- **日股 11-03 為文化之日休市：六家日股沿用 11-02 收盤。**

## 單一來源（sources_count＝1）
Oracle、NVIDIA、Arista、Coherent、Fabrinet、KLA、Dell（massive；marketbeat 差數美分）、Marvell、SanDisk、Credo（只有 marketbeat）、MiniMax、智譜、ASML、BESI、韓美半導體、中際旭創、新易盛、勝宏科技。

## 待查
- MiniMax：stockanalysis 155.40（−8.05%）；etnet 11-04 盤中漲跌回推昨收 155.70。本版取 155.40，差 0.2%。
- ASML、BESI：荷蘭財經部落格報價列（標示為收盤或來源時間），另一來源只給漲跌幅。
- 韓美半導體 180,000 為 KRX 正規盤（另有 KRX＋Nextrade 合併價 183,300）。
- Fabrinet 11-02 收盤：v0.6.22 取 316.89（stockscan），massive 與 Motley Fool 前收為 314.70，疑舊值為盤中快照（只影響已過去的版本）。
