# OpenAI 收支模型 交接檔（權威版；2026-10-07 起放 repo）

> 取代 Project「算力收支」的 `20261004_OpenAI_v0.6_handoff.md`（已刪除）。新對話串：先讀本檔與進度檔 `docs/reports/20261007_openai_進度.md`，再以 git 確認狀態。

## 1. 命題與範圍
- 命題（Andy 2026-09-30）：OpenAI 每 VR 等值 GW 的營收能否覆蓋每 GW 全成本；若不能，缺口由誰、以什麼條件融資。
- 期間 FY2025–FY2030，曆年制；2025 為實際校準年。上層目的：在下游 AI 變現尚未驗證時，檢驗重度、舉債支撐的中上游投資能否維持 OpenAI 成長。
- 分層：算力物理與產業資料取自 Tokenomics（`YinchenChang/Tokenomics` master `model/CURRENT`，以 TK_Link 快照）；公司財務原始數據在 `SRC_OAI`；假設在 `Inputs`。

## 2. 位置
- repo `YinchenChang/company-models`，資料夾 `openai/`（2026-10-07 由 `YinchenChang/openai-model` 搬入，來源 `main@13c16f0`；舊 repo 停用，Andy 可封存）。
- 規格：`docs/workorders/20261004_openai_v0.6.md`（r6；第 2 節 V1–V16 為全部已定決定，第 5–9 節為 P1 審查紀錄）。
- 分段工作單：`docs/workorders/20261007_openai_*`（共同規則＋S1–S6）。
- 成品：`dist/`；Drive `02-算力收支/OpenAI/`（chat 端放置）。

## 3. 流程（2026-10-07 起）
- 自主建置：chat 端寫工作單 → 建置代理執行（每步 commit＋push、進度檔）→ chat 端審查並合併；判斷類套用工作單預設，彙總給 Andy 在里程碑（S5＝v0.6）審查。
- 停止條件見共同規則第 6 節。

## 4. 模型狀態（v0.6-P1，`13c16f0`）
- SRC_OAI 106 列；Inputs 233 列（Assumed 182、Decision 35、Analogy 16；編號 1–245 無空洞）；v0.5 葉節點 849 全有去處；Checks C1–C27（C27 WARN：listPriceFYAvg 2026 區間待 P2）。
- TK_Link：v5.14 快照（OK 37、MISSING 11）→ S1 更新至 v5.24。
- v5.23 X10 已知影響（Tokenomics 串通知）：`IF_TrainCost_Luna／Sol／Astra` +14.34%／+10.82%／+9.21%、`IF_ProgGWyr_*` 同比例、`IF_AllocQ1` 0.609→0.636、`IF_AllocQ2` 0.637→0.664、`IF_AllocRDGW` +12.4%、`L1_Ans3` +12.4%；`IF_FullCost_*` VR200 +0.20%／+0.55%／+2.87%；`IF_RevGW_*`、`IF_TrainGenDefault` 不變。v5.24：訓練輸出上修（Tokenomics 交接第 0.1 節）。

## 5. 已定決定摘要（詳見 r6 第 2 節）
V1 時變世代組合；V2 η 以 2025 推論支出校準；V3 算力成本＝合約實付（TK 持有成本只作參考，差額＝雲端毛利）；V4 非算力研發＝研發總額 − 2025 訓練支出，之後人數 × 每人成本，股權報酬單列；V5 公司原始數據在 SRC_OAI；V6 訓練支出 12.0＝10.59＋1.41（0–2.91）；V7 複合標記拆開；V8 合約價單一參數 12（8–20）；V9 跨年拆開；V10／V14 來源四欄規則；V11 付費／免費產能本模型自算；V12／V15 pro 占比公式；V13 Cerebras／Azure 期間；V16 Azure 起點 2025（2025–2026）。

## 6. 2026-10-07 預設（Andy 授權 Claude 依建議）
- P2 價格事件時點：公告／生效日為基準，±1 季區間，天數加權。
- P3 η 路徑：基準沿用 2025 值；情境線性回升到 2030＝1。
- P4 人數等：先搜公開來源；搜不到暫用 v0.5 營收占比並標明。
- 其餘見各段報告「已套用的預設」。
