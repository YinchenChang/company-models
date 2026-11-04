# 工作單：Wattonomics v0.2.2＝改讀 Siliconomics v0.13 基準＋供給上限檢查

- 日期：2026-11-04（台北）；分支 `claude/wattonomics-v0.2.2`（由 main `2ce77fb`）
- 依據：Andy 11-04 14:10「AI 半導體的請求，先處理」；AI半導體 回執 `20261103_AI半導體-to-算力收支_回執_Siliconomics v0.12只留基準、R4 r3、NVIDIA對帳點`（請 Wattonomics 改讀 `*_Base`；Compare 的 Siliconomics 側改讀出貨年值）；Siliconomics v0.13（main）新增 Cap 頁與 `SIF_CapGW_Base`、`SIF_CapHead_Base`、`SIF_CoWoSUtil_Base`、`SIF_HBMUtil_Base`；三層校準原則（工作方式卡 v1.8 第 8 節）第 2 條「基準 GW＝需求側錨點，但不得超過供給產能」。

| # | 改動 |
|---|---|
| X1 | SIL 快照讀 v0.13（CURRENT_SNAPSHOT）；所有 `SIF_*_{A,B,C}` 改讀 `*_Base`（SIL_In、Compare、Spend 的 `SIF_AIITCapex`、`SIF_DeployGW`）；build.py 不再以 SCEN 迴圈讀 A／B／C；check.py：SIL_In 頁無 A／B／C 名稱 |
| X2 | Deploy 2028–29 拿掉 `SIF_UnitsCoef_A`（＝1，數值不變）；不改乘 `_Base`（地位鏈：Wattonomics 在上游）；Decisions 附 Andy 11-02 16:22 原話；驗收：市場年新增、毛額、R4 不變 |
| X3 | Compare Siliconomics 側改讀出貨年值：NVIDIA GW＝`SIF_DeployGW_Base` −（Siliconomics 口徑非 NVIDIA）；不再上線年往返換算；列 2026、2027 比值並對照回執 1.017／1.044；Wattonomics 側不變 |
| X4 | 供給上限檢查：Checks 逐年 Wattonomics GW（依 `SIF_CapGW_Base` 口徑對齊）≤ `SIF_CapGW_Base`，超過 WARN；Compare 加「供給上限可支撐 GW」與 headroom；只檢查不截斷（判斷題，預設不截斷） |
| X5 | README、Interface 註明 SIL 讀 v0.13 `*_Base` |

建置：build → soffice → export_snapshot → export_explorer → check.py → diff_snapshot → r4_attrib。不開 PR、不動 gh-pages。
