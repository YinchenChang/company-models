# 2026-11-04（週三）收盤價來源說明（地圖 v0.6.27 起）

`prices.csv` 每列的 `source` 欄記錄來源網站，`sources_count` 為取得同日收盤的獨立來源數。本版由單一代理查詢（2026-11-05 台北時間清晨，美股收盤後約 1–2 小時、亞洲 11-05 開盤前）。69 檔中 68 檔取得 11-04 收盤：雙來源 37 檔、單一來源 31 檔；1 檔沿用舊值（TOWA）。

## 取價原則
- 只採用頁面明示 11-04 日期的收盤，或頁面時間／昨收足以確認為 11-04 那一場的收盤；未標日期或日期較舊者不用。
- 本次環境只能讀取搜尋結果中出現過的網址（另有 massive、marketbeat 可直接讀），stockanalysis、cnyes、Yahoo 台股等頁面多為舊快取（停在 10 月），不可用；台股改用財報狗（statementdog，頁面標「11/04 收盤價」）為主，nstock、histock、cmoney 為第二來源。
- 美股：massive.com 分列「At close」與盤後，以正規盤收盤為準；marketbeat 標「Closing price 04:00 PM」者與 massive 一致，標「03:59 PM」者為最後成交、與 massive 差數美分至數十美分，此時取 massive、計單一來源。marketbeat 頁首未寫日期，以「頁首漲跌＋表中 11/03 收盤」換算確認為 11-04。
- ARM、Coherent 於 11-04 盤後公布財報：只用正規盤收盤（Coherent 盤後 211.20、ARM 盤後約 189–191 未用）。
- 台股第二來源：nstock 個股頁（更新時間 2026-11-04 13:30）；histock（13:30 已收盤，頁面時間 11-05）；cmoney（頁面昨收＝11-03 收盤，換算一致）；nstock「做什麼」頁（2026/11/05 06:53 發布，只列現價、未標價格時間）。後兩類未直接寫 11-04，屬推定一致。
- 韓股：KRX 正規盤收盤與 KRX＋Nextrade 合併價不同，只取正規盤。SK 海力士 alphasquare 1,422,000、mydailybyte 頁首 1,403,000 疑為含 NXT 盤後價，未用；取 mydailybyte 摘要句「11월 4일 종가…1,397,000원에 마감」。
- 11-04 日股、韓股、陸股、港股、歐股均有交易（無休市）。

## 單一來源（sources_count＝1）
- 美股（massive；marketbeat 為 03:59 PM 最後成交）：Oracle、Amazon、Broadcom、Arista、Coherent、Lumentum、Fabrinet、Dell。
- 港股：MiniMax、智譜（etnet，04/11/2026 17:59）。
- 日股：Ibiden、Disco、Kioxia（kabutan「11/4 本日」）。
- 台股：京元電子、欣興、鴻海、緯創、緯穎、奇鋐、光寶科、聯電、台燿（只有 statementdog）；金像電（只有 nstock）；世界先進（只有 cmoney，未標日期）。
- 歐股：ASML、BESI（beursgenoten，4 november 2026；昨收與 11-03 值一致）。
- 韓股：SK 海力士（mydailybyte）、三星電子（alphasquare 정규장）。
- 陸股：中際旭創、新易盛、勝宏科技（stoxray「Nov 04」；漲幅與 11-03 收盤換算一致）。

## 待查
- TOWA（6315）：未取得 11-04 收盤（kabutan、stockanalysis 為舊快取；松井證券只有 11-04 13:16 盤中 1,949）。沿用 11-02 收盤 2,133。
- 世界先進（5347）：cmoney 137（+3，+2.24%）未標日期，只以昨收 134 推定為 11-04；建議下次補第二來源。
- SK 海力士（000660）：1,397,000 只有一句摘要支持；頁首 1,403,000、alphasquare 1,422,000 疑含 NXT，待以 KRX 正規盤確認。三星電子 221,500 亦待確認是否為正規盤（alphasquare 標 정규장）。
- MiniMax、智譜：etnet 漲跌回推昨收 155.70、442.00，與 11-03 表列 155.40、441.80 差約 0.2 港元（同 11-03 版待查），本版以 etnet 11-04 收市值為準。
- 創意（3443）：statementdog 頁首 7,115（+320）與 nstock 一致，但同頁內文另有 5,995 元舊值，已以頁首為準。
