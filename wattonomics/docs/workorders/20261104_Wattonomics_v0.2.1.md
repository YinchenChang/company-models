# 工作單：Wattonomics v0.2.1＝非 NVIDIA 占比殘差法＋汰換年份攤提＋Oracle 三園區 Epoch IT 容量＋Anthropic v0.1.8＋TK 快照確認

- 日期：2026-11-04（台北）；chat 端：session_01V11BcwWqQyh9suCNhAYQKQ
- 現行：main（v0.2，PR #314，merge dcc43db）；分支 `claude/wattonomics-v0.2.1`
- 授權：Andy 11-04 12:52「我現在人在外面……所有需要我決定的，就先依你的建議去處理，然後你就依照還要做的事情的優先順序繼續做下去。」以下判斷題均依 chat 端建議預設處理。
- 依據：v0.2 獨立核對 `docs/verify/20261104_v0.2_independent_verify.md`（非 NVIDIA 占比疊加、汰換年份人為結果）；v0.2 建置報告待決 1–3。

## 1. 改動
| # | 改動 |
|---|---|
| W1 | **非 NVIDIA 占比改殘差法**：有逐家值的玩家維持現值（Google 0.753、Amazon 0.67、中國 1〔NVIDIA 對帳點口徑〕）；其餘玩家共用殘差占比＝（市場非 NVIDIA GW 總量 − 有逐家值玩家非 NVIDIA GW）÷ 其餘玩家 GW。市場總量依現有口徑；0.27 若為 $ 占比，先換成 GW 占比並在報告寫換算。殘差 < 0 時設 0 並 WARN。每年驗證加權總占比＝市場值。Decisions 註明：AI半導體 若提供逐買方份額，改讀之。報告列 Compare (8)、(12) 新舊值並重算 2027 未校準與校準後比值。 |
| W2 | **汰換年份攤提**：2025 年新納入存量 10.08 GW（含中國 8.33）預設平均攤到 2020–2025 六個年份 → 2026–2031 各年汰換 1／6；站點證據或 In 有年份資料者以資料為準。解決後移除 2031–32「暫定」。毛額與 R4 不變（check.py 驗證）。報告列 2026–2032 汰換與淨新增新舊。 |
| W3 | **Oracle 三園區改 Epoch IT 容量**：Port Washington 902、Saline 988、Jupiter 1,750 MW IT；上線年月沿用；Epoch 有分期則照其分期。標記 [Analogy]（第三方衛星與許可推估），原值保留為區間另一端；來源網址寫進 In（無法重開原頁則沿用 `evidence_raw/E_oracle_20261104.md` 與 v0.2 建置時的查證紀錄並標出處）。報告列 Oracle 年底 2026–28、站點 vs 資本支出推得差異新舊（上次 8.61 vs 4.87）。 |
| W4 | **Anthropic 取數改 v0.1.8**：`tools/import_gap_models.py` MODEL_OVERRIDE 改指 `anthropic/dist/20261103_Anthropic收支模型_v0_1_8.xlsx`；驗證 F07 2026–30＝7.034／−42.289／−49.563／−5.631／42.833、F03 皆 0；報告 Gap 頁 Anthropic 列新舊值。 |
| W5 | **TK 快照**：v0.2 若仍讀 v21.8 則換 v21.9（master 7276259）並報告影響；已是 v21.9 則註明不需動。 |

## 2. 不做（W6）
- 內容成長佔位值：等 R7（AI半導體，11-12）。
- 供給上限檢查：等 Siliconomics v0.13（約 11-17）；依 Andy 11-04 定案三層校準原則第 2 條「基準 GW＝需求側錨點，但不得超過供給產能」，排 v0.3。
- `siliconomics/tools/r4_bridge.py` 依賴列序：屬 AI半導體 工具，不改（已在交流告知）。

## 3. 建置與驗收
- build → soffice → export_snapshot → export_explorer → check.py → diff_snapshot；`r4_attrib.py --base-compare --prev <v0.2>` 出 R4 r4。
- 報告 `docs/verify/20261104_v0.2.1_build.md`。不開 PR、不動 gh-pages。
