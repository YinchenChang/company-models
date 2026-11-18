# 工作單：三層對帳 analysis-R（R1、R2）

- 指派日：2026-11-18（chat 端指派，analysis-R 執行；只讀不寫 repo，不開 PR）
- 依據：/home/claude/briefs/COMMON.md、/home/claude/briefs/建議書_r2.md 第 3 節與工作單 R1、R2
- 資料來源：company-models main d04f591；Tokenomics 79f6b30
- 輸出：/home/claude/out/analysis-R/

## 指派內容（摘要原文要點）

背景：2027–2029 Siliconomics 基準部署 GW 比 Wattonomics 換算值每年高約 14–15 GW（39.16／42.23／42.68 對 24.49／28.18／28.48，成本等值、出貨年）。初步粗算以營收反推，2027 差距八到九成在 NVIDIA（Siliconomics 約 27.9–29.6 實體 GW，Wattonomics 約 15.3 GW）。任務：用顆數與功耗重做。

R1 平台別拆解（2026–2029 逐年）
- Siliconomics：Units、Deploy、Coord、PGW 與快照 SIF_UnitsPlat_*，換成實體 GW 與成本等值 GW，寫明每 GW 顆數與來源；平台加總回到 SIF_DeployGWPhys_Base、SIF_DeployGW_Base（列殘差）。
- Wattonomics：找出 NVIDIA／非 NVIDIA（或更細）的拆法；拆不到的層級明說是模型結構限制，不用比例硬拆。
- 兩層並排表與差距；差距集中在哪些平台與年；NVIDIA 占差距比例（區間與來源）；非 NVIDIA 是否相近。
- 往上追：買方清單／覆蓋、NVIDIA 份額假設、每 GW 顆數或內容、年度歸屬，各能解釋多少 GW；不能唯一歸因的留在「未解釋」。

R2 每 GW 顆數共同口徑：同一晶片世代在 Tokenomics（Interface）、Siliconomics（TK_In、PGW、chip_bridge）、Wattonomics（TK_In）的顆／GW 對照、相對差、功耗與 PUE／利用率口徑；量化 chip_bridge 讀 v21.7 的差異。通過條件：換算差 ≤ 1% 且口徑明示。

樣本外：NVIDIA FY2027 Q3（107.6）不得用於校準或選參數；只可在報告最後「樣本外檢驗」引用一次並標明未入模型。

輸出：(1) 白話報告 .md；(2) 同名 .xlsx（公式連輸入頁）；(3) Python 重算自我核對。
