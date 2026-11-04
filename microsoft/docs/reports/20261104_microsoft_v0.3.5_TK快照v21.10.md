# Microsoft v0.3.4 → v0.3.5：Tokenomics 快照換 v21.10（基準與 Sens 都不變）

**結論：** 數值完全不變（加權目標價 384.10、放空）。v21.10 只改 `IF_NonNV_Out_TPU`／`IF_NonNV_Out_Trn`，本模型的快照不含這兩個名稱（Maia 讀 `IF_NonNV_*_Maia`，不變），所以差異只有版本與快照說明文字。整本重算的 Sens 案例重跑，逐一相同。

## 做了什麼
- 用 `tools/tokenomics/import_tokenomics.py` 產生新快照 `data/tokenomics_snapshot_v21.10.json`。
  - 來源：`20261104_Tokenomics_v21.10.xlsx`，master eded8e6（`model/CURRENT` 已核對），2026-11-04 取數。
  - 名稱清單與舊快照相同（`data/tokenomics_names.txt`，42 個，optional 3 個）。
  - `--check` 通過。
  - 與 v21.9 快照逐值比對：值全部相同（只有檔名、版本、commit 不同）。
- 舊檔移到 `_待刪除/`（沿用本資料夾慣例）：
  - `data/tokenomics_snapshot_v21.9.json`
  - `data/sens_v034_20261104.json`
  - `dist/20261104_Microsoft個股模型_v0.3.4.xlsx`、`.html`
  - 例外：v21.4、v21.7 快照保留在 `data/`，因為 `attrib_v031.py` 以 `--snap` 重現 v0.3 → v0.3.1 拆解（同 #308）。
- 改了 `build_microsoft.py`：`--snap` 預設改 v21.10；In A1、In 頁 C 節標題加上 v0.3.5 的快照說明；敏感度檔名改 `sens_v035_20261104.json`。沒有改任何模型輸入值或判斷。
- `sens_v033.py`、`make_v03.sh`：預設輸出改 `data/sens_v035_20261104.json`，成品名 v0.3.5。
- 重跑敏感度：`sens_v033.py` 整本重算 53 筆，與 v0.3.4 逐一相同。活公式的 Sens 案例在同一本 xlsx 內重算，也逐格相同。變動拆解 `attrib_v031.py --solve` 不受快照影響（固定讀 v21.4／v21.7），沒有重跑。
- 成品 `dist/20261104_Microsoft個股模型_v0.3.5.xlsx`／`.html`。

## 基準數值比較（逐格比對，openpyxl data_only，容差 1e-9）
| 項目 | v0.3.4 | v0.3.5 |
|---|---|---|
| 加權目標價 | 384.1046 | 384.1046 |
| 投資建議（上檔） | 放空（−27.7%） | 放空（−27.7%） |
| DCF／分部加總 | 324.37／443.84 | 324.37／443.84 |
| 情境 A／B／C | 423.25／386.05／341.07 | 423.25／386.05／341.07 |
| 讀法 1 | 310.88（放空） | 310.88（放空） |
| 收支關卡 A／B／C | 363.2／259.6／137.9 | 363.2／259.6／137.9 |

全部差異格共 24 格，都是文字：
- **In 頁**：A1、A95（快照說明）、H96:H111、H140:H141（來源文字中的快照版本與 commit）。
- **Sens 頁**：F36、F37，`he_all` 說明中的 Tokenomics 版本（數值不變：GB300 10.89 → 8.43／16.69）。
- **Sources 頁**：B6、D6，快照檔名與 commit。

數值格 0 差異。

## Sens 變動
無。整本重算 53 筆與活公式案例都與 v0.3.4 逐一相同，沒有評等級距變動。

## 檢查
`make_v03.sh` 全部通過：
- `facts_fy27q1.py` 兩讀核對 67 筆，不一致 2 筆皆已解釋（口徑不同）；沒有改動已提交的 JSON。
- `check_v03.py`：公式錯誤 0、核對列 26、加權目標價 384.10。
- `mag/check_in_layout.py`（警告 Val!C38:D38 讀敏感度低／高欄，與 v0.3.4 相同，不判失敗）。
- `check_sens_coverage.py --require-keycol`：50 鍵，豁免 3。
- segment_check。
- HTML 建置與核對。

## 待 Project 判斷
無。
