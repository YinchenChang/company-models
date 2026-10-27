# Siliconomics 變更紀錄

## v0.1（2026-10-27；`model/20261027_Siliconomics_v0.1.xlsx`；Tokenomics v21.7）

- 新建 AI 半導體產業基準庫（比照 Tokenomics）：README、Scen、In、Path、Mkt、TK_In、Conflicts、Interface、Checks、Decisions、DB_Evidence。
- `markets/memory/market.py` 的全部輸入（DRAM／NAND 季價、加速器顆數與每顆 HBM、情境係數、HBM 市場、對帳參考，共 161 列）逐字搬入 In 頁；`market.py` 改為只讀 `data/` 快照，保留原 Python 名稱。驗證：新舊 `market.py` 的全部公開常數與輸入列 `repr` 相同；`ai-semis-build` 58 個目標重建後與 dist 逐格相同（錯誤格 0）。
- Tokenomics v21.7 鏡像 22 個名稱（`TK_In`、`SIF_TK_*`）；Checks K3：下游（chip_bridge v20.1 快照）與現行 TK 不同 2 格（VR200 低成本欄 `IF_CapexIT`、`IF_CapexTotal`）。
- Conflicts 頁登錄 22 項跨模型衝突（含算力收支 2026-10-25 所列 15 項；盤點見 docs/plan/20261027_行業變數盤點.md）。
- CI：`ai-semis-build` 新增本庫重建比對與快照一致檢查。
