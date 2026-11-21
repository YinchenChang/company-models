# 2026-11-20 收盤價來源說明（地圖 v0.6.57 起）

- **77 檔上市公司＋Anthropic 改用 2026-11-20（週五）收盤**：Yahoo Finance 日線（未還原股利；交易所當地日期），2026-11-21 以 GitHub Actions 工作流程 `px-closes-fetch`（`tools/px_closes/fetch_closes.py`，yfinance）擷取。雲端工作環境連不到 Yahoo，所以改在 GitHub Actions 抓。每日收盤全檔在 `closes_20261109_20261120.csv`（215 個代號，含同業、子公司、匯率）；Anthropic 同業補抓在 `closes_extra_20261120.csv`。
- **核對：** 新檔的 2026-11-13 收盤與上一份 `closes_20261020_20261113.csv`、`closes_extra_20261113.csv` 逐檔相同（SanDisk 除外，見下）。
- **12 個代號日線尚無 11-20 一列**（ABBN.SW、ASM.AS、ASML.AS、ATS.VI、BESI.AS、ENR.DE、LR.PA、NESN.SW、SIE.DE、SNDK、SU.PA、TPRO.MI）：取 Yahoo 報價中繼資料的最近成交價（時間為當地 11-20），清單在 `closes_20261109_20261120.csv.meta.txt`。地圖個股中屬於這一類的是 ASML、BESI、SanDisk、ABB。日線補上後應再對一次。
- **SanDisk：** 2026-11-20 起 8:1 分割，Yahoo 全期間為分割後口徑（11-13＝110.626251＝885.01 ÷ 8；11-20＝79.24）。模型 v0.5.7 已改分割後口徑。
- **Anthropic：** 2026-11-19 於 Nasdaq 掛牌（代號 ANT，IPO 價 83）；11-19 收盤 97.17、11-20 收盤 88.91。地圖自本版起用收盤。
- 代號對應同 `prices_20261113_sources.md`。
