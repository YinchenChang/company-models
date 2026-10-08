# 工作單 Anthropic A2｜builder、需求與營收（v0.1-A2；2026-10-08，chat 端）

- 分支 `claude/anthropic-v0.1`（已存在；A0、A1 已在其上）。開工：`git fetch origin && git checkout claude/anthropic-v0.1 && git pull`；若 PR 尚未開，開 draft PR（base `main`）。
- 規格：第 2 節頁結構、第 3 節 D1–D5、D14、D15、D20；資料：`data/anthropic_src.yaml`。

## 步驟（每步 commit＋push＋進度檔）
1. **builder 骨架**：仿 OpenAI v0.6 `builder/build.py`，但資料來源改為 `data/anthropic_src.yaml`→SRC_ANT、`data/anthropic_inputs.yaml`→Inputs（不建 Map_v05、Derived_V9）。沿用 `common.py`、`preserve.py`（Excel 優先）、`source_rules.py`（四欄初評，改為讀 yaml 的 tag／party）、`tk_link.py`（同 OpenAI v0.6 的 63 名）。新增 `oai_link.py`（D20）。`id_registry.json` 發出 SRC_ANT_nnn、INP_nnn、Cnn。README 頁。
2. **TK_Link 擴充（D6）**：Tokenomics `NonNV` 頁以列標籤讀出 TPU v7 Ironwood、AWS Trainium3（另可帶 AMD MI455X 供將來）各 6 值（產出比低／基準／高、持有比低／基準／高），名稱 `TK_NNV_<族>_Out`、`_OutLo`、`_OutHi`、`_Hold`、`_HoldLo`、`_HoldHi`；狀態欄「讀表（非具名）」；`check_tk_snapshot.py` 同步比對這些列。
3. **Inputs 草稿**：`data/anthropic_inputs.yaml`：本段需要的全部假設（每筆值、低、高、標記、依據、區間理由）。可參考 OpenAI v0.6 Inputs 的同類參數作 Analogy（寫明「OpenAI v0.6 INP_nnn」與差異）。
4. **Demand**：個人用戶（Free／Pro／Max 5x／Max 20x；2025 校準）、企業席位（Team 標準／Premium、Enterprise）、任務類別（對話、程式代理、其他代理）× 每任務 token × 層級組合（Haiku→Luna、Sonnet→Sol、Opus→Astra）；API：2025 token 由 2025 API 營收 ÷ 有效單價倒推（D14），之後任務成長 × 每任務 token 成長 × 價格彈性；`DEM_Tok_*`（層級 × 付費／免費）。
5. **Revenue**：訂閱方案別；API 層級 × 輸入／輸出 × 有效價（牌價 × 快取與批次折扣組合；價格事件以公告日天數加權）；通路（直接／雲端夥伴，D5）；其他；廣告＝0（D3）；總額、夥伴分成、淨額；個人 vs 企業彙總；2025 對 SRC 實際差距＝0（校準）並列未校準值；run-rate 里程碑對照列（D14、D15）。容量上限係數先以 1 佔位（A3 接 Compute）。
6. **Checks**（C01 起）：CHK_Errors、SRC／Inputs 筆數、TK 快照版本、OAI_Link 未被計算頁引用、2025 營收校準差距、2026 run-rate 對照（>25% WARN）、E6。
7. **測試與 CI**：仿 OpenAI 的 `tests/parity`（情境 ≥8 個，涵蓋價格彈性、通路分成、方案組合、任務成長）、`test_builder.py`（E6、ID、OAI_Link 隔離）；`.github/workflows/anthropic-parity.yml`（由 `openai-parity.yml` 改寫；環境變數改 `ANTHROPIC_ENGINE_CACHE` 或沿用並在 `engine/core.py` 相容）。
8. 報告 `docs/reports/YYYYMMDD_v0.1-A2.md`＋`_對照.xlsx`、CHANGELOG、PR 留言、進度檔。

## 判斷類預設（未列者依規格 D 表）
| 問題 | 預設 | 替代 |
|---|---|---|
| 方案價格的年付折扣 | 年付占比 Inputs（Assumed 0.3，0.1–0.5） | 全月付 |
| 個人付費人數無官方值 | 2025 由個人訂閱營收（若有）÷ 方案組合加權價倒推（D14 校準）；兩者皆無：第三方估計（Analogy）並給區間 | — |
| 企業席位單價 | Team 牌價；Enterprise 以報導值（無則 Team Premium 的 Analogy） | — |
| 層級組合（Haiku／Sonnet／Opus 占 token） | Inputs（Assumed；依 OpenRouter 等公開份額作 Analogy 區間） | — |
| 價格逐年變動 | 每代新模型同價或降價：Inputs 年變動（Assumed −15%，−35%～0%）；已發生的價格事件進 SRC | 不降價 |
| 2025 API 營收（若只有比例） | API 占比 × 2025 營收（Derived，比例為 Interested-party 或 Analogy） | — |
