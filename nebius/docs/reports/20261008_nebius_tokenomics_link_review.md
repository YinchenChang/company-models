# Nebius｜Tokenomics 連接審視（2026-10-08）

- 分支：`claude/nebius-v0.2a-tkanchor`（PR #30；本審視併在工作單 v0.2a 步驟 1 完成，見 `docs/workorders/20261008_nebius_v0.2a_每MW收入Tokenomics錨定.md` 依據第 4 項）。
- 共用工作單：`docs/workorders/20261008_tokenomics_link_review.md`（PR #28 分支 `claude/tokenomics-link-review`，開工時未合併，從該分支讀取）。
- 對照：Tokenomics master `ca78a8f`，`model/CURRENT`＝`20261008_Tokenomics_v5.27.xlsx`（v5.27 合併雜湊 `19d667f`；`19d667f..ca78a8f` 之間 `model/` 無變動）。下游資料契約 `docs/plan/Tokenomics_downstream_contract.md` v0.1。
- 前提：v5.29 **尚未合併**（`IFW_`、`IFC_` 名稱不存在）→ 依共用工作單只做第 1–3 節，第 4 節標「待 v5.29」。

## 結論（三句）
1. 取數方式已由手工 JSON（v5.24，含 `DC_Cost!`、`Inputs!`、`Interface!` 直接引用）改為官方快照 `data/tokenomics_snapshot_v5.27.json`（26 個 `IF_`／`L1_` 名稱，`--check` 可重現），Excel 新增「Tokenomics_取數」分頁與 110 個 `TK_` 具名範圍；v5.24 → v5.27 本模型用到的值逐格不變，三情境目標價不變。
2. Nebius 收入按 **MW-year** 計價（契約第 3 條）：收入端只用「可計費 MW ÷ 已連網 MW」（在役比例 × 爬坡）與簽約率（MW 驅動固定 100%），`IF_Util` 不進收入、利用率欄維持 100%，沒有重複扣減。
3. v0.2 的每 MW 收入仍是「成本加成與市場價 50/50 平均」且容量與單價同一選擇器同向綁定（違反契約第 4 條）；本工作單 v0.2a 步驟 3 改為 `IF_HoldEcon` 錨 × 定價倍數 k、兩條情境軸分離，50/50 平均停用。

## 1. 取數點清單

`link_audit.py`（PR #28 分支版本）在本分支的結果：取數方式＝官方快照、釘住 `20261008_Tokenomics_v5.27.xlsx`、使用名稱 26、不存在於現行 0、值漂移 0、四層瀑布／用途欄＝否／否。工具列出的 9 筆「違規引用」全部是 `IF_Hdr` 字樣出現在說明文字「（不含 IF_Hdr*）」（`build_xlsx.py` 註解、`company.json` 的 `_note`、`data/tokenomics_names.txt` 註解，另 6 筆是 `out/` 暫存副本，不入版控）——屬工具誤判，沒有任何公式或程式引用 `IF_Hdr*`。

| 本模型的格或參數 | Tokenomics 名稱 | 合規 | 四層瀑布 | 用途是否符合 `IFC_Use` | 版本 |
|---|---|---|---|---|---|
| Excel「Tokenomics_取數」全部 330 個值（26 名稱 × 世代 × 成本情境） | 26 個（見 `data/tokenomics_names.txt`） | 合規（`IF_`／`L1_`，經官方快照） | 不適用（成本與物理量，非產能／營收） | 待 v5.29（`IFC_` 尚無） | v5.27（`19d667f`；快照 commit `ca78a8f`） |
| 每 MW 年收入的錨（v0.2a 步驟 3 起） | `IF_HoldEcon`（`TK_HoldEcon_<世代>`，基準成本情境） | 合規 | 不適用（持有成本＝打平租金，100% 計費時數口徑；利用率不另扣） | 待 v5.29；成本底線用途，非上限 | v5.27 |
| 定價倍數 k 的證據分母（步驟 3） | `IF_GPUhrEcon`（證據表；每 GPU 小時持有成本，100% 利用率） | 合規 | 不適用 | 同上 | v5.27 |
| 收入上限檢查（步驟 3） | `IF_RevGWFleet`（`TK_RevGWFleet_<世代>`） | 合規 | 屬「理想上限」 | 只作上限檢查、不作收入預測（符合「上限，不得作預測」） | v5.27 |
| 成本端評估並列（步驟 5；不切換） | `IF_OpexGW`、`IF_MaintIT`、`IF_StaffSW`、`IF_TaxIns`、`IF_PowerCost` | 合規 | 不適用 | 只作對照 | v5.27 |
| v0.2 舊每 MW 收入 11.62／17.40／24.20（對照列） | 原手工 JSON（v5.24 `IF_HoldEcon`、`IF_GPUhrEcon`、`IF_GPUsPerGW`、`IF_FacilityGW`） | 已改：檔案保留作歷史、標「已被快照取代」，工作表直接引用改寫為名稱 | — | — | v5.24（數值與 v5.27 相同） |
| 每 MW 建置成本 `scenarios.capexTemplate.costMW` 50.12／50.26 | `IF_CapexTotal`（GB300／VR200 基準） | 數值相符、仍為寫死常數（v0.2 起即如此） | 不適用 | 符合 | v5.24＝v5.27 |
| GPU 經濟壽命 5 年 | 非 Tokenomics（公司會計政策）；Tokenomics `IF_DeprLifeIT` 6 年只作對照 | — | — | — | — |

手工 JSON 的處理：`data/permw_tokenomics_20261007.json` 與 `data/nebius_facts_20261007.json` 開頭加 `_supersededBy`；`Interface!…` 改寫為 `TK_<名稱>_<世代>（成本情境）`，`Inputs!D23:F23` → `IF_DeprLifeIT`，`DC_Cost!I48:N48` → `IF_OpexGW`，`Inputs!D28:F28`（WACC）→ 註明「無 IF_ 名稱、內含於 IF_HoldEcon、不直接引用」，`DC_Cost!J22` → 註明「內部計算格，只作說明」。數值未改。改寫後兩檔 `DC_Cost!`、`Inputs!`、`Interface!` 字樣 0 筆。

## 2. 口徑問題
1. **計價方式**：MW-year（神雲）。簽約率：MW 驅動下新產能簽約率固定 100%（`defaults.revenueDriver`＝mw）；可計費 MW＝已連網 × 在役比例（0.75／0.85／0.90…）× 爬坡係數（60%／80%／100%）；利用率欄 100%。`IF_Util`（60%）不進收入（v0.1a 路徑 B 曾乘 80–90% 可計費利用率，已隨 50/50 平均一併停用）。**無重複扣減**；交接檔寫明「按 MW-year 計價」。
2. **50/50 平均與同向綁定**：v0.2 是（三情境 11.62／17.40／24.20 隨容量情境同向）。v0.2a 改為錨 × k，容量軸（保守／基準／積極）與價格軸（k 低／基準／高）分離，三個容量情境預設都用基準 k；成本加成路徑只留作 ROIC 對照。
3. **模型商 `IF_Alloc*`**：不適用（Nebius 不是模型商，未取用）。
4. **公司特有變數**：WACC 11%、定價倍數 k、長約占比、簽約率、MW 交付進度都只在本模型 `company.json` 設定，沒有回寫 Tokenomics。

## 3. 缺口與版本
- Tokenomics 缺的產業級數據（不自行建表，交 chat 端走 G2）：(a) GB300／VR200 **長約**每 GPU 小時成交價時間序列（目前只有 IREN–Microsoft 一筆可推算；契約第 4 條要求隨需／預留／長約三層分開，`SRC_Price` 尚無長約層）；(b) Hopper H200 獨立欄（以 Hopper H100 欄代表）。
- 名稱清單：`data/tokenomics_names.txt`（26 個，已執行，不是草稿；與 CoreWeave W4 相同清單）。

## 4. v5.29 重取影響
待 v5.29（`IFW_` 四層瀑布、`IFC_Use`、`L1_HoldEconMW_*`、`L1_TokMW_Gen_ratio_*` 尚不存在）。已列入交接檔與 `nebius/CLAUDE.md` 待辦：v5.29 合併後另開工作單重取，並以 `--check` 對現用值比較。

## 待 chat 端判斷
1. `link_audit.py` 把說明文字中的「IF_Hdr*」判為違規（CoreWeave 同樣 3 筆）；建議工具只掃描程式碼與值、或排除「不含 IF_Hdr」字樣。
2. `scenarios.capexTemplate.costMW` 仍為寫死的 Tokenomics 值（數值相符）；改為引用 `TK_CapexTotal_*` 屬成本端，本單範圍外，建議併入 v5.29 重取工作單。
