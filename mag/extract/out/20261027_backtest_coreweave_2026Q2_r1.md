# 萃取摘要｜20261027_backtest_coreweave_2026Q2_r1

- 公司 coreweave；期末 2026-06-30；模式 backtest；模型 claude-opus-5-5、claude-sonnet-5-5
- 文件：_drive_20261027_backtest_coreweave_2026Q2_r1_0.pdf（）
- 欄位 166：兩模型皆有值 86（一致 84、不一致 2）；只一個有值 7；皆無 73
- 估計費用約 US$1.54
- **回測**：可比對 91 項，與 company.json 現值一致 87 項（96%）
  - latestQuarter：64/66
  - ytdActual：17/17
  - callFacts：1/1
  - ROLL_FIELDS（首期一次性金額）：3/3
  - ROLL_FIELDS（期初餘額）：2/4
  - 註：ROLL_FIELDS 現值多為推導或排程，不一致不代表萃取錯；latestQuarter／ytdActual 才是直接可比的部分。

| 欄位 | 現值 | 萃取值 |
|---|---|---|
| latestQuarter.ppe | 46.736 | 52.622 |
| latestQuarter.vieExposure | 0.108 | 1.7 |
| valuation.shares | 0.587 | 0.551 |
| defaults.eqCapShares | 0.5865 | 0.551 |

候選值不是模型數值；填入 company.json 前須逐項對原文核對（見 xlsx「候選值」頁，紅＝兩模型不同、黃＝只一個找到）。
