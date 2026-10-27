# Siliconomics 下游資料契約 v0.4.1（2026-10-27；對應 Siliconomics v0.4.1、Tokenomics v21.7）

適用對象：`company-models` 內 AI 半導體 Project 的市場層（`markets/*`）與公司模型（每家公司一個資料夾），以及其他專案取用晶片產業資料時（經 Drive `00-共用情境/交流/`）。

## 1. 三層邊界

| 層 | 內容 | 持有者 |
|---|---|---|
| 算力層（第 0 層） | 每 GW 顆數、機架數、機架價、功率、每 GW 資本支出、NonNV 產出比與持有成本比、產業年度部署路徑（R4） | Tokenomics／算力收支（本庫只鏡像於 `TK_In`） |
| 產業層（本庫） | 加速器顆數與每顆內容、HBM 與記憶體價格路徑、各市場規模、技術替代滲透率（CPO、銅轉光、液冷、混合鍵合）、同一市場各供應商份額表、供需缺口、情境 A／B／C 定義 | Siliconomics |
| 公司層 | 會計年度錨點、毛利率與費用、資本支出、股數、估值參數、公司特有的份額相對值 | 各公司模型 |

## 2. 取數規則

1. 只取 `SIF_` 具名範圍（Interface 頁）；`SGOV_Errors` 須為 0 的版本才可匯出快照。不得連結 In、Path、Mkt、Checks 等內部頁。
2. 經快照取數：`data/CURRENT_SNAPSHOT` 指向的 JSON（`tools/export_snapshot.py` 產生，記錄 xlsx 檔名、版本、sha256）。公司模型在來源欄寫明快照版本。
3. 低／高（`SIF_<鍵>_lo`、`_hi`）只作敏感度；主值一律取基準。
4. 情境：公司模型的 A／B／C 一律讀本庫同字母（Andy 2026-10-25 13:13）。行業變數不得在公司模型內另訂一套。
5. Tokenomics 的量：經本庫 `SIF_TK_*` 取（值＝Tokenomics 現行快照，版本見 Interface 第 2 列），或直接讀 Tokenomics 快照；兩者不得混用不同版本。
6. 公司模型發現行業變數需要修正：改本庫 In 頁（同一 repo 的 PR），不在公司模型覆寫。

## 3. 版本與引用

| 版本 | 日期 | 契約名稱 | 說明 |
|---|---|---|---|
| v0.1 | 2026-10-27 | `SIF_` 輸入 161 個（＋ `_lo`／`_hi`）、推導 20 列、`SIF_TK_*` 鏡像；`SGOV_Errors`、`SGOV_Warn` | 由 `markets/memory/market.py` 逐字搬入（加速器顆數、每顆 HBM、情境係數、DRAM／NAND 季價、HBM 市場、對帳參考）；`market.py` 改讀快照，38 個下游建檔程式不需修改、重建後逐格相同 |
| v0.2 | 2026-10-27 | 新增輸入 `SIF_g29_*`（8 個，2029 平台別顆數成長）；新增推導 `SIF_UnitsPlat_*`（平台別、合計、A／B／C 合計，2026–2029）、`SIF_UnitsGrowth29`；既有名稱的基準值不變，Assumed 列補 `_lo`／`_hi` | `market.py` 新增 `G29`、`G_U29`；tsmc、network 讀 `G29`，nand 讀 `G_U29` |
| v0.3 | 2026-10-27 | 新增輸入 `SIF_pa_*`、`SIF_core_*`、`SIF_f_hbm`、`SIF_sh_*`、`SIF_msh_*`、`SIF_mtk_*`、`SIF_al*`（35 個）；推導 `SIF_ASICShareSum_*`（5 列） | 新增 `siliconomics/reader.py`（`row8`／`row9`／`V`／`NAMES`）；nvidia、avgo、mrvl、mediatek、alchip 改讀；取數規則第 2 條：公司建檔程式可直接 import reader，不必經 market.py |
| v0.4 | 2026-10-27 | 新增輸入 `SIF_c_*`（54）、`SIF_it_*`（9）、`SIF_suopt_*`、`SIF_su_cu`、`SIF_cpo_*`（12）、`SIF_vr26`；推導 `SIF_AIITCapex_*`、`SIF_ITCapexPerGW`、`SIF_DeployGW_*` | 新增 `reader.vals()`；markets/chip_bridge 改讀 `siliconomics/data/tokenomics_snapshot_v21.7.json`（本專案唯一一份 TK 快照） |
| v0.4.1 | 2026-10-27 | 新增輸入 `SIF_cposh_*`（15 個，CPO 光學金額份額 2027–2029）；推導 `SIF_CPOShareSum`（四家模組與雷射廠合計，不含 Fabrinet） | lite、cohr、innolight、eoptolink、fabrinet 改讀 `SIL.row9("cposh_<公司>_<年>")` |

版本號：只改證據、標記、文字為修補版（v0.1.x）；輸入值或公式改變為次版（v0.x）；Interface 名稱刪除或定義改變為主版（v1.0 前只新增名稱以次版處理）。
