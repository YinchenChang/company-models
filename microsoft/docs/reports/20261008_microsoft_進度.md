# Microsoft v0.1 進度檔（接手用）

## 目前狀態（每次 push 前覆寫）
- **已完成**：v0.1a 全部步驟（0–9）。總帳 722 筆、Tokenomics v5.27 快照（41 名、missing 3）、k 證據、共識檔、`company.json` → `mag` 區段（21 子區）、資料報告 `docs/reports/20261008_microsoft_v0.1a_資料.md`；verify.sh 19 項全過（引擎未動）。PR #33 已留言回報。
- **下一步（v0.1b′ 起點）**：等 Amazon v0.1b 產出 MAG 共用引擎後移植（可讀 `amazon/` 複製）。起點：(1) `meta.consensusFile` → `data/consensus_msft_20261008.json`，刪 `consensus_orcl_20261007.json`、`consensus_crwv_20260925.json` 與 `company.json` → `oracle` 區段；(2) `calendar.fiscalYearEndMonth`＝6、最新已申報 FY26Q4、首期 FY27 全年（若 FY27Q1 已於 2026-10-28 公布，改 Q1 實際＋0.75 年，並更新總帳）；(3) 依 `mag.*.mapTo` 搬欄位：分部用 FY27 新分部（Agents and Infra／Devices and Consumer）＋新產品別，雲端拆分對象＝Azure（新定義，FY26Q4 年化 117.7）；(4) AI MW 期初 4.5 GW-IT（3.5–5.5）、對外 60%、k 1.06（0.76–1.31）；(5) neocloud 租金入營業成本；(6) 租賃改分類（FY27 起更多營業租賃）與未起租 329.1 的處理需設預設。注意事項全文見資料報告第 10 節。
- **未解問題**：無（停止條件未觸發）。待 chat 端知悉：拆分基準 36.5 可能偏高 20–40%（公司 AI run-rate 對照）、上限檢查 59% > 50% 示警。

## v0.1a 資料蒐集（2026-10-08）
工作單：`mag/docs/workorders/20261008_mag_v0.1a_資料蒐集.md`；分支 `claude/microsoft-v0.1`；PR #33。

| 步驟 | 狀態 | commit | 備註 |
|---|---|---|---|
| 0 基線、PR、進度檔 | 完成 | 4a1df12 | verify.sh 19 項全過；SEC EDGAR shell 可直連；draft PR #33 |
| 1 SEC 一手文件 | 完成 | 42a80ac／本次 | 10-K FY26、三份 10-Q、8 季新聞稿、2026-09-02 8-K（Azure 絕對值）；本次補回年度損益 45 筆（建置腳本漏寫） |
| 2 AI 容量（MW） | 完成 | 42a80ac | 公司增量＋Bloomberg、Morgan Stanley、Epoch、Omdia；期初 3.5／4.5／5.5 GW-IT |
| 3 k 證據 | 完成 | 42a80ac | Azure 牌價 11 筆、長約 5 筆；建議 k_長約 0.76、k_現貨 1.76、加權 1.06 |
| 4 租用算力與租賃 | 完成 | 42a80ac | Nebius、IREN、Nscale、Lambda、CoreWeave；未起租 329.1 |
| 5 關聯方 | 完成 | 42a80ac | OpenAI 25%、24.1、250、45% RPO、分成 20%／上限 38；Anthropic 30／≤5 |
| 6 AI／非 AI 拆分 | 完成 | 本次 | 對外 AI 36.5（17.3–67.7）、非 AI 81.2（50.0–100.3） |
| 7 評價輸入 | 完成 | 42a80ac | beta 1.10（兩源）、rf 5.28%、kd 6.26%（Aaa 類比）、同業 11 家 |
| 8 市場共識 | 完成 | 42a80ac | 現價 529.76；PT 587.63（S&P）／578.73（MarketBeat） |
| 9 mag 區段、verify、報告、PR 回報 | 完成 | 本次 | fields_doc 補 mag 21 子區說明並 --write；verify.sh 19 項全過 |
