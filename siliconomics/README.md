# Siliconomics — AI 半導體產業基準庫

AI 半導體 Project 的「唯一事實」Excel：加速器顆數、記憶體與 HBM、網通、CoWoS、載板、伺服器組裝、電源散熱、設備等**行業變數只有一個值**，各公司模型與市場層只讀這裡。規格比照 `YinchenChang/Tokenomics`；與 Tokenomics 相關的數據一律讀 Tokenomics、不自行重估。

| 項目 | 內容 |
|---|---|
| 現行版本 | 見 `model/CURRENT` |
| 下游快照 | `data/CURRENT_SNAPSHOT` 指向的 `data/siliconomics_snapshot_v*.json`（由 `tools/export_snapshot.py` 產生） |
| Tokenomics 快照 | `data/tokenomics_snapshot_v21.7.json`（名稱清單 `data/tokenomics_names.txt`） |
| 規劃書 | `docs/plan/Siliconomics_規劃書.md` |
| 下游契約 | `docs/Siliconomics_downstream_contract.md` |
| 變更紀錄 | `CHANGELOG.md` |

## 建置（每個 Excel 新版本）

```bash
# 1) 產生（之後各版加 --base model/<現行>.xlsx，In 頁藍字讀回現行 Excel）
python3 siliconomics/builder/build.py --out build/raw/<YYYYMMDD>_Siliconomics_vX.Y.xlsx
# 2) LibreOffice headless 重算存檔（錯誤格須為 0、Checks 的 SGOV_Errors 須為 0）
soffice --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir siliconomics/model build/raw/<檔名>.xlsx
# 3) 匯出快照並指向它
python3 siliconomics/tools/export_snapshot.py siliconomics/model/<檔名>.xlsx siliconomics/data/siliconomics_snapshot_vX.Y.json
echo siliconomics_snapshot_vX.Y.json > siliconomics/data/CURRENT_SNAPSHOT
# 4) 下游全部重建（CI ai-semis-build）：數值變動須全部落在本版預期範圍內
```

## 原則（比照 Tokenomics CLAUDE.md）

- Excel 是唯一事實來源；`builder/` 只產生公式頁與格式，不得在程式內新增或省略計算機制。
- 輸入（In 頁藍字）由 Excel 擁有：要改值直接改 Excel；builder 重建時以 `--base` 讀回。`builder/seed/` 只用於 v0.1 首建。
- 下游只讀 Interface 的 `SIF_` 名稱（經快照）；不得連結 In、Mkt 等內部頁。
- Tokenomics 權威的量（每 GW 顆數、機架數、機架價、功率、每 GW 資本支出、NonNV 比例、產業部署路徑）只在 `TK_In` 頁鏡像；Tokenomics 升版時重抽快照、重建，Checks K3 列出下游仍用舊版的格數。
- 本 repo 為公開 repo：不得放 Andy 的個人情境、先驗或校準紀錄。
