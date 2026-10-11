# Microsoft（MSFT）MAG v0.2 進度檔（接 Tokenomics v5.31＋四層瀑布上限檢查）

分支 `claude/microsoft-v0.2`；工作單 `mag/docs/workorders/20261011_mag_v0.2_Tokenomics_v5.31.md`。Tokenomics 唯讀 clone `/home/claude/tk`：master `f16f161`（`model/CURRENT`＝`20261008_Tokenomics_v5.31.xlsx`），2026-10-11 確認仍為最新。引擎來源：`origin/claude/amazon-v0.2` @ `c07a167`（進度檔「v0.2 引擎完成 @41f9a54」；c07a167 只改進度檔，引擎與 41f9a54 相同）。

## 目前狀態（每次 push 前覆寫）
- **已完成：v0.2 全部步驟**。引擎與 amazon@c07a167（v0.2 引擎 41f9a54）逐檔相同；快照 v5.31（含四層瀑布）；C10 重解 2,511.7；升版 v0.2（VLOG、dist `20261011_Microsoft收支模型_v0_2`、交接檔 v0.2、README、根目錄 README）；報告 `docs/reports/20261011_microsoft_v0.2.md`；截圖 17 張（溢出 0、錯誤 0）；verify.sh --vs-dist 對 v0.1（EXPECT）與對 v0.2 皆 26 項全過。
- 結果：讀法 1（D2）385.91／355.08／356.77；讀法 2 409.44／425.10／424.43；FY29 對外 AI ROIC 9.81% vs WACC 10.90%、打平 k 1.092。
- **下一步**：FY27Q1 季度更新（2026-10-28 後），升版 v0.3。
- **未解問題（待 Andy）**：讀法主值；基準速度高於積極情境（新增）；對外比例；長約 k；未起租營業租金。

## 步驟
| 步驟 | 內容 | 結果 | commit |
|---|---|---|---|
| 0–1 | 名稱清單加 `IFW_RevGWFleet_100／_Util／_Prod／_Life`、`IFC_Use`；重取 `data/tokenomics_snapshot_v5.31.json`（40 名、missing 3：IF_NonNVRatio、IF_NonNVCostRatio、IF_NonNV_Maia；`--check` 通過）；舊快照 v5.27 刪除（v0.1 成品的重現以 main `97a8bf4` 為準） | — | （本次） |
| 2 | 引擎同步 amazon@c07a167：segA.js、build_xlsx.py、cmp31.js、check_tokenomics_tab.py、verify.sh、xl_diff.py、fields_doc.py（mag 說明保留本家）；`company.json` 以 `scripts/port_company_json.py` 重產（只多 capUtilWarn、capProdMax 與 tokenomics 指 v5.31）。**不變性**：v0.2 引擎＋v5.27 快照＋舊速度，對 v0.1 成品值 0 差異，只有 9 列改名（上限檢查改「利用率層」、取數頁來源列） | — | （本次） |
| 3 | C10 重解：基準速度 2,325.9 → **2,511.7 MW-IT／年**（對帳落差 0） | 上限檢查：利用率層最高 43.9%、生產折減層最高 58.7%（皆通過）；verify.sh 23 項全過 | （本次） |
| 4–7 | 拆解（序列 (a) IF_HoldEcon →(b) 其餘 Tokenomics 名稱與 neocloud 淨租金改讀 v5.31 營運成本 →(c) C10 →(d)＝0）；評價口徑敏感度與翻轉點重算；升版 v0.2；`scripts/expect/v0_2_mag_vs_v0_1.txt`（14 列改名＋1,506 格）；port 腳本營運成本改讀快照、meta.updateDate 2026-10-11；報告、截圖 | 基準 353.55 → 355.08（+1.53＝+0.20＋3.89−2.56＋0）；讀法 2 422.48 → 425.10 | （本次） |
