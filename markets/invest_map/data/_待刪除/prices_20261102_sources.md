# 2026-11-02（週一）收盤價來源說明（地圖 v0.6.22）

`prices.csv` 每列的 `source` 欄記錄來源網站，`sources_count` 為取得同日收盤的獨立來源數。三組代理分別查美股、台股、其他市場（2026-11-03）。

## 取價原則
- 只採用頁面明示 11-02 日期的收盤；未標日期或日期較舊者不用（stockanalysis 部分頁面停在 8–10 月、Investing、Google Finance、Yahoo 美股頁多為舊快取）。
- 「cnyes(11/03頁昨收)」：cnyes 頁面已換到 11-03 盤中，取其「昨收」（＝11-02 收盤），頁面未印日期；這些代號皆另有一個標日期的來源且數值相同。
- 「(% implied)」：來源只給 11-02 漲跌幅，以 10-30 收盤換算；只當第二來源。
- futunn 與 moomoo 為同一資料源；fool.com 與其 Yahoo 轉載為同一來源。
- 五個市場（港、東京、阿姆斯特丹、韓、深圳）11-02 皆有交易（日本假日為 11-03）。

## 未取得 11-02 收盤（沿用 10-30 收盤）
- Micron（MU）：各頁皆為舊值、盤前或盤中；有標題稱 11-02 下跌 5.8%（約 833），無確切收盤。
- Coherent（COHR）：唯一來源自相矛盾（200.26／200.00，−4.94%）。

## 單一來源（sources_count＝1）
WhiteFiber、Fabrinet、Credo（以上只有 stockscan，曾出現盤中快照，信度較低）、AMD、Amkor、AMAT、LRCX、KLAC、Teradyne、智邦、智譜、Disco、ASML、BESI、韓美半導體、勝宏科技。

## 來源之間不一致（CSV 取值在前）
- Oracle 119.22（massive，−8.72%，與 stocktitan −8.76% 吻合）；fool 119.10、stockscan 119.12。
- KLA 153.11（massive）對 stockscan 153.51。Marvell 171.41（stocktitan、massive）對 fool 170.80。AMD 466.84 對 466.74。
- Credo 140.02（stockscan）對 massive 圖表「last」147.04；Amkor 36.96（stockanalysis）對 massive 39.80（判為不可靠）。
- 韓美半導體 191,700（中央 11-03 文「지난 종가」，與 −4.07% 換算一致）；同文另列 192,100、mydailybyte 192,200，推測為 KRX＋Nextrade 盤後合併價（推論，未證實）。
- 已剔除：stockscan GOOGL 282.92（與其 GOOG 頁矛盾）、stockscan LRCX 241.33（成交量過低，疑為盤中）。

## 期間重大事件（價格大幅變動，供判讀）
- 11-02 Marvell 宣布以全股票約 1,150 億美元收購 Lumentum：LITE −18.6%、MRVL 約 −12.7%（報導轉述，未經本專案檢核）。
- CoreWeave −8.6%（債務疑慮報導）；Oracle −8.7%。
