# 萃取摘要｜20261027_backtest_oracle_FY27Q1_r1

- 公司 oracle；期末 2026-08-31；模式 backtest；模型 claude-opus-5-5、claude-sonnet-5-5
- 文件：_drive_20261027_backtest_oracle_FY27Q1_r1_0.pdf（）
- 欄位 169：兩模型皆有值 48（一致 47、不一致 1）；只一個有值 12；皆無 109
- 估計費用約 US$0.95
- **回測**：可比對 57 項，與 company.json 現值一致 50 項（88%）
  - latestQuarter：30/32
  - ytdActual：15/16
  - ROLL_FIELDS（首期一次性金額）：2/3
  - ROLL_FIELDS（期初餘額）：3/6
  - 註：ROLL_FIELDS 現值多為推導或排程，不一致不代表萃取錯；latestQuarter／ytdActual 才是直接可比的部分。

| 欄位 | 現值 | 萃取值 |
|---|---|---|
| latestQuarter.debtPrincipal | 125.904 | 125.0 |
| latestQuarter.ppe | 127.845 | 153.596 |
| ytdActual.prepay | 11.363 | 15.36 |
| rpo.bucketWeights[0] | 0.0975 | 0.13 |
| valuation.shares | 3.057736 | 3.024 |
| defaults.ppeOpen | 75.74 | 127.845 |
| defaults.prepay.openBalance | 15.955 | 30.789 |

候選值不是模型數值；填入 company.json 前須逐項對原文核對（見 xlsx「候選值」頁，紅＝兩模型不同、黃＝只一個找到）。
