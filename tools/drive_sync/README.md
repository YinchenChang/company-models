# Drive 同步清單（`tools/drive_sync/`）

repo 內的成品（`dist/` 的 xlsx、html，及 `siliconomics/model/`、`wattonomics/model/` 的現行活頁簿）與 Google Drive 閱讀資料夾的對照。共用區，由 AI供應鏈 維護（見 `docs/OWNERSHIP.md`）。

## 檔案

| 檔案 | 用途 |
|---|---|
| `manifest.csv` | 對照表：`repo_dist_glob,drive_folder,category,project`。一列一個成品資料夾（檔名 glob）；`software_ai/*/dist/*`、`physical_ai/*/dist/*` 為萬用列，各該專案新增公司時不必改本表 |
| `manifest.py` | 依對照表列出每個資料夾的**現行版本檔**與目標 Drive 資料夾（CSV 或 JSON）；`--check` 檢查每個有 `dist/` 的資料夾都已列入 |

## Drive 資料夾（Andy 2026-11-17）

| 類別 | Drive 資料夾 | 範圍 |
|---|---|---|
| 雲端大廠 | `02-AI供應鏈/個股模型/雲端大廠/` | Alphabet、Amazon、Microsoft、Meta、Oracle |
| 模型商 | `02-AI供應鏈/個股模型/模型商/` | OpenAI、Anthropic、MiniMax、智譜、SpaceX |
| 算力平台 | `02-AI供應鏈/個股模型/算力平台/` | CoreWeave、Nebius、WhiteFiber |
| 硬體供應鏈 | `02-AI供應鏈/個股模型/硬體供應鏈/` | 其餘單一公司資料夾（半導體、硬體、網通、電源等） |
| AI算力地圖 | `02-AI供應鏈/AI算力地圖/` | `markets/invest_map/dist/`（資料活頁簿與 `index.html`） |
| Siliconomics | `02-AI供應鏈/Siliconomics/` | `siliconomics/model/` 現行活頁簿；`markets/` 的市場模型（晶片內容橋接表、網通、設備、NAND） |
| GW帳 | `02-AI供應鏈/GW帳/` | `wattonomics/model/` 現行活頁簿 |
| software_ai | `04-AI軟體應用/個股模型/` | `software_ai/` 各公司與單位經濟模型 |
| physical_ai | `03-AI實體應用/個股模型/` | `physical_ai/` 各公司（併入 main 後生效） |

各類別資料夾內扁平排列，不開公司資料夾。

## 現行版本的判定

- 檔名 `<YYYYMMDD>_<名稱>_v<版本>.xlsx|.html`：同一資料夾取版本號最大者（同版本取日期最新），同版本的 xlsx 與 html 全列。版本號的 `.` 與 `_` 視為相同分隔（`v0_6_6`＝`v0.6.6`）。
- 資料夾內有 `CURRENT` 檔（`siliconomics/model/`、`wattonomics/model/`）時，以 `CURRENT` 指定的檔為準。
- 檔名沒有版本號的 xlsx／html（`markets/invest_map/dist/index.html`）一律列為現行。
- `_待刪除/` 子資料夾與 `.json` 等其他檔不列。

## 同步流程（由 Claude 對話執行）

1. 在 main 的最新 checkout 執行 `python3 tools/drive_sync/manifest.py`（或 `--format json`），取得「repo 檔 → Drive 資料夾」清單。
2. Claude 對話經裝置橋接（`device_commit_files`）把清單上的檔案複製到 Andy 裝置上對應的 Drive 資料夾。
3. 同一資料夾內被新版取代的舊檔（同公司、較舊版本），移到該 Drive 資料夾的 `_待刪除/`，並告知 Andy；不直接刪除。
4. **Drive 只放現行版本，歷史版本由 git 保留**（需要舊版時從 git 歷史取出）。

## 檢查

```
python3 tools/drive_sync/manifest.py --check
```

`ownership-check` CI 每個 PR 都會執行；新增有 `dist/` 的公司資料夾而未列入 `manifest.csv`（且不在萬用列範圍內）時失敗。
