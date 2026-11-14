# 工作單：Wattonomics v0.3.2＝W-1 七家以外買方資本支出路徑（取代 neo、sov 站點路徑）

- 日期：2026-11-14（台北）；分支 `claude/wattonomics-v0.3.2`，**由 v0.3.1 合併後的 main 建立**（v0.3.1 未合併前不得開工）
- 授權：Andy 2026-11-14 10:53「其他的照預設進行」、「我授權你合併之後的 PR」。判斷類照下列預設，寫進 Decisions（接續編號），可撤回。
- 依據：三層基準協調規則 v1.1（本版完成後 Wattonomics 覆蓋視為補齊，`cov_done`＝1 的前提）；證據 Drive `02-AI供應鏈/GW帳/20261114_W1_七家以外買方證據.{md,csv}`（68 列；建置代理拿不到 Drive 時由 chat 端複製到 `wattonomics/evidence_raw/20261114_W1/`）；v0.3.1 財報季證據。
- 證據紀律：W-1 證據多為二手轉述；**入基準的列須經未參與代理核對**（chat 端另派），核對前標「待核對」只入 Sens。

| # | 改動 |
|---|---|
| X1 | **三條線建路徑（2026–2029，年新增 IT GW 與資本支出）：** (1) 新雲由下而上：CoreWeave、Nebius、IREN 等以實績＋已簽約 MW＋公司指引（指引只作參照、不作基準）；非上市新雲（Crusoe、Lambda、Fluidstack、Nscale、Together）以已簽約 MW 為準；(2) 主權：只有已簽約站點入基準（E12），目標與 MOU 只入 Sens；(3) 伺服器 OEM（Dell、HPE、Supermicro、Lenovo）AI 伺服器營收與在手訂單作交叉檢核列，不直接入基準 |
| X2 | **與七家去重（判斷題預設，Andy 10:53 照預設）：** 先扣已公開的對七家合約（重疊下限，例：CoreWeave 對 OpenAI、Meta；Nebius 對 Microsoft；IREN 對 Microsoft；Cipher 對 AWS）；其餘未揭露客戶部分列 [Assumed]，基準取 50% 屬七家以外、區間 0–100%，並列敏感度 |
| X3 | **Anthropic 經 Fluidstack 等租用的容量（TeraWulf、Cipher、Hut 8；Google 擔保）算 Anthropic（七家以外）**；Google 擔保只是信用支持（判斷題預設） |
| X4 | 出租方與承租方同一批 MW 只算一次（例：Applied Digital、Galaxy 租給 CoreWeave）：以最終使用者計 |
| X5 | 口徑：MW 未標口徑者基準取 IT 讀法（既有規則）；Galaxy 526 MW IT ÷ 800 MW 毛電力＝0.66 作換算參照與敏感度 |
| X6 | 2028–29：無公司指引；只有實績與已簽約撐住的部分入基準，其餘 [Assumed]＋區間（既有判斷題預設）；外部預測只入 Sens |
| X7 | 新路徑取代 neo、sov 站點路徑（既有預設）；更新 `WIF_MktGW_*`、`WIF_Evidence_*`（分母的覆蓋缺口 ④ 改由本路徑實績取代）；Compare 與 Siliconomics 非七家 GW 比對；「NVIDIA 稱約占一半」只作參照列 |
| X8 | README、Interface、Decisions、check.py；報告 `wattonomics/docs/verify/20261114_v0.3.2_build.md`（表 1、表 2、待判斷；與 v0.3.1 逐格比，`WIF_MktGW_*` 變動逐年說明） |

建置：同 v0.3.1。完成後開 PR；**未參與代理核對通過後由 chat 端合併**（Andy 10:53 授權）。不動 gh-pages。
