# 趨勢科技 Trend Micro（4704）個股模型 v0.1

- 原型 P（按防護範圍計價）；12 月制；官方報告分部＝所在地別四分部（Japan、Americas、Europe、Asia Pacific），Vision One 為各分部之下的命題子列。
- 本版未做先驗。不適用：Meta 的算力、自用 AI 讀法 2、Wattonomics、三層協調（非 Siliconomics 覆蓋）。
- 建置：`bash software_ai/trendmicro/make_v02.sh`（`SENS=1` 重跑敏感度與反解）；研究包 `data/research_trendmicro.md`、`data/facts_trendmicro.csv`；Tokenomics 快照 `data/tk_snapshot.json`（v22.0）。
- 私有化查證（2026-11-16）：2025-02 Reuters 傳聞（Bain、Advent、EQT、KKR）未見成交，公司仍上市，繼續建模。
- 成品：`dist/20261116_趨勢科技個股模型_v0.1.xlsx`、`.html`。
