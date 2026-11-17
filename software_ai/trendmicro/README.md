# 趨勢科技 Trend Micro（4704）個股模型 v0.2.1

- 原型 P（按防護範圍計價）；12 月制；官方報告分部＝所在地別四分部（Japan、Americas、Europe、Asia Pacific），Vision One 為各分部之下的命題子列。
- 本版未做先驗。不適用：Meta 的算力、自用 AI 讀法 2、Wattonomics、三層協調（非 Siliconomics 覆蓋）。
- 建置：`bash software_ai/trendmicro/make_v02.sh`（`SENS=1` 重跑敏感度與反解）；研究包 `data/research_trendmicro.md`、`data/facts_trendmicro.csv`；Tokenomics 快照 `data/tk_snapshot.json`（v22.0）。
- 私有化查證（2026-11-16）：2025-02 Reuters 傳聞（Bain、Advent、EQT、KKR）未見成交，公司仍上市，繼續建模。
- 成品：`dist/20261117_趨勢科技個股模型_v0.2.1.xlsx`、`.html`。
- v0.2（2026-11-17）：相對倍數讀 `software_ai/data/relmult_hist/out/trendmicro.json` 的 geo_base（0.35）；Comparable 主腿改 P/S 口徑（同業 P/S 中位 × rel_mult × FY2027E 營收），EV/EBIT 改對照列；Sens 加 ps_only、psg_only、sens.low、sens.high；DCF（股票薪酬加回）對照列；舊成品在 `dist/_待刪除/`。PS 與 PSG 讀數分歧原因（成長比 0.47）見模型 README 工作表與 HTML。
- v0.2.1（2026-11-17，G7）：同業 P/S 改讀 `peer_ps_20261113.json`（11-13 價，GEN、QLYS 逾期）；Comparable 分母改目標時點 TTM 營收（日數加權）、股權口徑；Sens 加「同業 P/S＝資料庫最近一季中位」案。
