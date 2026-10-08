# 工作單 Anthropic A3｜算力與成本（v0.1-A3；2026-10-08，chat 端）

- 分支 `claude/anthropic-v0.1`；A2 經 chat 端審查後開工。規格：D6–D13；參考 OpenAI v0.6 `builder/p3.py`、`p4.py` 與其報告 `20261007_v0.6-P3.md`、`-P4.md`。

## 步驟（每步 commit＋push＋進度檔）
1. **Compute**：加速器族 × 世代（NVIDIA Hopper／GB200／GB300／VR200；TPU v6e／v7；Trainium2／3）每 GW 年產能（付費／免費、層級；D6）；逐合約供給 GW（D10：金額或 GW 互推；設施口徑 ÷ PUE）；時變組合（D7）；token GW、η（D8，2025 推論支出校準）、有效推論 GW；研發 GW 殘差（D12，並列 TK `IF_AllocRDGW`）；總需求 GW；容量上限係數（接回 Revenue，取代 A2 佔位）；VR 等值 GW（各族 × 世代 Sol 產能比）；敏感度（合約價、η 路徑、TPU／Trainium 產出比低高）。
2. **Cost**：算力成本＝逐合約實付＋自有資本支出（D10、D11）；供應商持有成本與雲端毛利；經濟口徑並列；非算力成本（D13）、股權報酬；**命題表**（每 VR 等值 GW 的營收淨額、算力成本、非算力成本、全成本（含／不含股權報酬）、差額、覆蓋率）；TK `IF_FullCostDefault_*` 單位成本參考列。
3. Checks 補 Compute／Cost 項（供給 ≥ 推論、研發 GW 非負旗標、2025 算力支出對帳、η 合理範圍 WARN）。
4. parity 情境補 ≥6 個（合約價、η、產出比、人數成長、自有資本支出、PUE）；報告 `v0.1-A3`＋對照 Excel（第②頁必含各年命題表）、CHANGELOG、PR 留言、進度檔。

## 判斷類預設
| 問題 | 預設 | 替代 |
|---|---|---|
| 合約起始與爬坡無揭露 | 公告後次季起，線性爬坡至公告規模（期間 Inputs，Assumed 2 年，1–3 年） | 立即全量 |
| 「最高可達」GW（如 Google 1M TPU、Azure 1 GW） | 基準計入公告值的 Inputs 比例（Assumed 0.8，0.5–1.0） | 全計 |
| PUE（設施→IT） | Inputs（Analogy 1.25，1.1–1.4；TK 若有同名則改引用） | — |
| 閒置比例、訓練最低占比 | 沿用 OpenAI v0.6 同名 Inputs 的值與區間（Analogy，寫明 INP 編號） | — |
| 自研（Anthropic 無自研晶片） | 不設 genFactor.custom | — |
