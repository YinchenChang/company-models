# Siliconomics 變更紀錄

## v0.2（2026-10-27；`model/20261027_Siliconomics_v0.2.xlsx`；Tokenomics v21.7）

- In 頁輸入改動（工作單 `docs/workorders/20261027_v0.2_inputs.json`，以 `tools/apply_inputs.py` 寫入 Excel 後重建）：Assumed 列全部補區間（C 情境 DRAM／NAND 季價 2027 ±5pt、2028–29 ±10pt；Rubin Ultra、AMD、Trainium、其他 ASIC 顆數與每顆 HBM）；Rubin Ultra 每顆 HBM 標記改 Interested-party、情境 A 係數改 Derived；2029 HBM 單價兩列補來源。**基準值全部不變。**
- 新增 M5 節 2029 平台別顆數成長（`g29_*`，Blackwell −50%、Hopper 0、其餘 +15%，區間 ±10pt）與 Units 頁；Interface 新增 `SIF_g29_*`、`SIF_UnitsPlat_*`、`SIF_UnitsGrowth29`（13.77%）。
- Checks：G2（缺來源）、G3（Assumed 缺區間）升為 ERROR（皆 0）；新增 K4（平台別合計＝HBM 頁合計）、K5（成長率在區間內）。
- 下游：`market.py` 新增 `G29`、`G_U29`，replay 只放 M1–M4；tsmc、network 的 2029 成長改讀本庫（值不變）；NAND 的 2029 顆數年增 10% → 13.77%（CF18 已併入）。重建 57 個目標：數值變動只有 NAND、SanDisk、Kioxia；skhynix、micron、samsung、tsmc、network 只有 In 頁區間與來源文字；第二輪重建全部相同。

## v0.1（2026-10-27；`model/20261027_Siliconomics_v0.1.xlsx`；Tokenomics v21.7）

- 新建 AI 半導體產業基準庫（比照 Tokenomics）：README、Scen、In、Path、Mkt、TK_In、Conflicts、Interface、Checks、Decisions、DB_Evidence。
- `markets/memory/market.py` 的全部輸入（DRAM／NAND 季價、加速器顆數與每顆 HBM、情境係數、HBM 市場、對帳參考，共 161 列）逐字搬入 In 頁；`market.py` 改為只讀 `data/` 快照，保留原 Python 名稱。驗證：新舊 `market.py` 的全部公開常數與輸入列 `repr` 相同；`ai-semis-build` 58 個目標重建後與 dist 逐格相同（錯誤格 0）。
- Tokenomics v21.7 鏡像 22 個名稱（`TK_In`、`SIF_TK_*`）；Checks K3：下游（chip_bridge v20.1 快照）與現行 TK 不同 2 格（VR200 低成本欄 `IF_CapexIT`、`IF_CapexTotal`）。
- Conflicts 頁登錄 22 項跨模型衝突（含算力收支 2026-10-25 所列 15 項；盤點見 docs/plan/20261027_行業變數盤點.md）。
- CI：`ai-semis-build` 新增本庫重建比對與快照一致檢查。
