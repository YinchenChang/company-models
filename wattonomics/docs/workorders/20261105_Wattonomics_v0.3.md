# 工作單：Wattonomics v0.3＝R7 分世代內容＋逐買方 NVIDIA 份額＋供給上限可比口徑＋讀取來源換新

- 日期：2026-11-05（台北）；分支 `claude/wattonomics-v0.3`（由 main `2c93b32`）
- 授權：Andy 11-04 12:52「我現在人在外面……所有需要我決定的，就先依你的建議去處理，然後你就依照還要做的事情的優先順序繼續做下去。」；11-05 08:11「請繼續」。判斷類採預設，寫進 Decisions，可撤回。
- 依據：AI半導體 回執 `20261104_AI半導體-to-算力收支_回執_R7分世代NVIDIA內容、逐買方NVIDIA份額、產能上限（Wattonomics v0.3前置）`（Siliconomics v0.13.1，PR #324）；Tokenomics v21.10（company-models #328）。

| # | 改動 |
|---|---|
| Y1 | SIL 快照讀 v0.13.1；TK 快照讀 v21.10（報告列影響；預期只影響非 NVIDIA 相關列） |
| Y2 | Compare 每 GW NVIDIA 內容改讀 `SIF_NVContentGW_{H100,GB200,GB300,VR200_26,VR200,RU}`（含低、高），依各年世代組合加權，取代佔位值；重算所需內容成長與校準後比值，列 2025–2027 新舊 |
| Y3 | 改讀 `SIF_nvsh_<買方>_<yy>`（含 _lo／_hi），取代 Google 0.753、Amazon 0.67、中國 1 與殘差法。金額份額換 GW 份額：NVIDIA GW＝買方 GW × s ÷（s＋(1−s)×c），c＝非 NVIDIA ÷ NVIDIA 每 GW 加速器成本（TK）；「資本支出 × 加速器占比」作敏感度並列。買方鍵對應 goog、amzn、msft、meta、orcl、xai、oai、neo、cn、sov、anth（Anthropic 經 Google 雲的 TPU 算 Google）。報告列全市場加權 NVIDIA 占比 vs 由上而下 82%／80%、2025–27 比值新舊 |
| Y4 | 供給上限主檢查改讀 `SIF_CapOnlineGW_Base`（可支撐新增上線 IT GW，毛額）對 Wattonomics 毛額年新增；原 G 節成本等值保留作第二口徑；只檢查不截斷 |
| Y5 | r4_attrib 出新版；README、Interface 更新；check.py：SIL 的 nvsh、NVContent 名稱存在；世代內容 低 ≤ 基準 ≤ 高 |

不做：W1 殘差法套到每 GW 成本（Y3 取代後不需要，Decisions 註明）；汰換權重與 C2–C4 排 v0.4。
建置：build → soffice → export_snapshot → export_explorer → check.py → diff_snapshot → r4_attrib。不開 PR、不動 gh-pages。
