# 工作單 Anthropic v0.1.3 B 段｜Bridge 頁、η 拆解、S-1 替換程序、成品（2026-10-22，chat 端）

先讀 `20261022_anthropic_v0.1.3_共同.md` 與進度檔 v0.1.3 節（A 段須已完成；若未完成，先完成 A 段的剩餘步驟）。分支 `claude/anthropic-v0.1.3`。

## 步驟（每步 commit＋push＋進度檔）
1. **Bridge 頁（新工作表，放在 Cost 之後；唯讀，不被任何計算頁引用，測試確認）**：由上而下逐層：用戶與席位 → 任務 → 每任務 token（TK `IF_TaskTok*`）→ 付費／免費 × 層級 token → 有效單價 → 營收；token → 每 GW 理論產能（TK `IF_TokGW_*` × 世代與加速器組合）→ ÷ η → 有效推論 GW → ＋研發 GW 殘差 → 供給 GW → 合約付款與自建 → 成本 → 每 VR 等值 GW 差額。每列：公式（引用既有具名範圍）、單位、2025 與 2030 值、低／高（若 Inputs 有）、標記。最後一節「與 Tokenomics 對照」：本模型 vs TK 同口徑量（付費 token、每 GW 營收 `L1_RevGW_PQ_VR200`、`L1_TokLedger_2026Q2`）與差額，註明 TK token 帳未閉合屬 Tokenomics 待辦。
2. **η 拆解（診斷，不改 η 值）**：η（2025）＝ 計價口徑因子 × 利用率因子 × 工作負載組合因子 × 未解釋殘差。可得資料：計價口徑＝每 GW 合約實付 ÷ TK 持有成本（加權）；利用率＝TK `IF_Util` 對本模型隱含利用率（若無，以 Inputs Analogy 區間）；工作負載＝本模型層級與輸入輸出組合下的每 GW token ÷ TK 參考請求組合下的每 GW token（若無法同口徑，標「無法分離」）；殘差＝η ÷ 前三者乘積。並列 TK `L1_GapPrompt`、`L1_GapSpendBasis`、`L1_GapISL`、`L1_GapUtil`、`L1_GapUnexplained`（OpenAI 口徑，Analogy）。另做一張小表：TK v5.26 → v18.0 時各因子與殘差的變化（從 archive 的 v0.1.1 模型取舊值），回答「版本變動是否被殘差吸收」。
3. **S-1 替換程序**寫入 `anthropic/CLAUDE.md` 新節：公開 S-1 或經審計資料出現時，新增 SRC 列（不覆寫）、舊列 note 標「已被 SRC_ANT_nnn 取代（日期）」並保留、模型引用改指新列；報告列每項舊值／新值／差異比例／受影響輸出／結論是否翻轉。
4. **成品**：`tools/build_html.py` 加「資金三層」「持續擴張情境」「η 拆解」三小節（精簡；每數字附 Excel 位置）與「v0.1.3 更新了什麼」；`dist/20261022_Anthropic收支模型_v0_1_3.html`／`.xlsx`（v0.1.2 dist 檔 `git mv` 到 `dist/_待刪除/`）；以 Playwright 打開 HTML，逐節讀畫面數字並截圖（`docs/reports/img/`，桌機與手機寬度，無水平捲動）。
5. 報告 `docs/reports/20261022_v0.1.3.md`＋`_對照.xlsx`（①新增與變動列、②前後對照與逐項拆解〔TK 標記、U2、其餘只作情境或顯示〕、③η 拆解表、④外部驗收報告逐項處理表（採納／部分／不採納與理由，照共同說明）、⑤已套用的預設）；CHANGELOG、交接檔（版本、待 Andy、待辦）、進度檔；PR 留言 `[Anthropic 回報] v0.1.3｜完成｜日期`；PR 轉 ready（`gh api -X POST repos/YinchenChang/company-models/pulls/<n>/ccr/ready_for_review`）。不合併。
