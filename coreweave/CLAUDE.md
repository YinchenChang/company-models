# CoreWeave 個股模型 — Claude Code 工作守則（v5.0 起；Oracle 收支型樣板）

舊守則（v4.x JS 雙引擎）保留在 `legacy_v4/CLAUDE.md`；其中第 8 條（雙引擎、company.json 唯一來源）、第 16(1) 條 DCF 權重 45%、第 12(1)(2) 條目標價時點、評等「賣出」門檻已被取代（工作單 `docs/workorders/20261112_CoreWeave_v5.0_Oracle樣板重建_Q3.md` 第 6 節）。

## 結構
- `build_coreweave.py`：唯一計算來源（Python 產生活公式 xlsx，LibreOffice 重算）；事實總帳 `data/coreweave_facts_*.json`；Tokenomics 快照 `data/tokenomics_snapshot_*.json`。
- HTML 由共用產生器 `tools/company_html/` 讀 xlsx（`scripts/make_html_json.py` 產生 `html.json`）；不自帶計算。
- 建置與核對一律 `bash coreweave/make_v5.sh`（改公式或輸入時加 `SENS=1 REV=1 DEC=1`）；全部通過才可提交。

## 仍有效的已決定事項（承接 v4.x）
1. **可移植**：新雲（Nebius、WhiteFiber）照本樣板重建；公司特有內容放 In 頁與事實總帳，程式只寫通用邏輯。
2. **差異要能解釋**：與共識、指引的差距寫在 Recon，標觀點／拆法／已知限制；指引只作 Recon 與 Sens，不作界值（原則 59）。
3. **精簡**：主畫面只放結論與關鍵數字，明細摺疊（HTML 版面規格 v1.2）。
10. **每季滾動**：財報後 FY 首期改為「年初至今實績＋剩餘季度模型」，升版並拆解目標價變動。
11. **HTML 單一檔案、離線開啟**：`check_company_html.py` 通過；1280／390 寬無水平捲動（`layout_check.py`）。
12. **目標價變動拆解**：每次升版拆段相加＝總變動（`scripts/decomp_v50.py`）。
13. **營運成本**以 Tokenomics 滿載穩態為主值。
14. **每 MW 收入以 Tokenomics `IF_HoldEcon` × 公司因素 k 為錨**；不以營收、ARR、RPO、指引或公司宣稱價格反推 k（v5.0 J4：公司宣稱只進 Recon、Sens、反解）。
15. **公司實況驗證**：差距 >10% 找有證據的機制，找不到則 Tokenomics 為基準、公司實際為 Sens（Recon C 節）。
16(3). **單一 k** 套用所有世代，待有 VR200 長約價再分世代。

## 個股模型共通規格
`docs/plan/個股模型原則_命題加營收分部.md`（§10 第 1–62 條）、`docs/plan/個股DCF標準規格.md`、`docs/plan/個股HTML版面規格.md`、`docs/plan/個股模型驗收_不需驗收項目.md`、`docs/plan/自算產業β_20261105.md`。

## 溝通
繁體中文；結論先行；判斷題列選項與預設，由 Andy 決定。合併 PR 由 chat 端安排未參與代理核對後進行。
