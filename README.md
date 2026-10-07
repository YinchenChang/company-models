# company-models

個股收支與評價模型（算力收支 Project）。每家公司一個資料夾，各自完整可建置（CRWV 模板的副本＋該公司的 company.json 與資料）。

| 資料夾 | 公司 | 來源模板 | 狀態 |
|---|---|---|---|
| `nebius/` | Nebius（NBIS） | `YinchenChang/crwv-model` @ `01b13ad`（v4.5） | 建置中 |
| `coreweave/` | CoreWeave（CRWV） | `YinchenChang/crwv-model` @ `01b13ad`（v4.5；2026-10-07 起在此維護，改接 Tokenomics 中，見 `coreweave/docs/workorders/`） | 改造中 |

規則：
- 事實來源與分層：算力相關的產業與物理層資料取自 Tokenomics（`YinchenChang/Tokenomics` 的 Interface／L1／SRC，X8）；公司專屬資料放各公司的 `company.json` 與 `data/`。
- 範圍：做到 CRWV 目前水準（v4.5）；CRWV 尚未完成的待辦（5a 精簡、5b 只讀檢視器、折舊修正、機率加權目標價、GPU 批次與續約價格衰退），各公司也先不做（Andy 2026-10-07）。
- 流程：chat 端（Claude）寫工作單（`<公司>/docs/workorders/`）→ 建置代理執行 → chat 端審查並合併；成品（HTML、Excel）由 chat 端放到 Andy 的 Google Drive 對應資料夾。
- 各公司資料夾內的 `CLAUDE.md` 沿用 CRWV 工作守則；與本 README 衝突時，以本 README 的流程與範圍為準。

## 共用工具：Tokenomics 取數（`tools/tokenomics/import_tokenomics.py`）
- **用途**：把 Tokenomics 的 Interface（`IF_`）與 L1（`L1_`）具名範圍讀成版本固定的快照 JSON（記錄檔名、`model/CURRENT`、commit、擷取日；每個值記名稱、世代、成本情境、單位、工作表!儲存格），供各公司 `company.json` → `tokenomics.snapshotFile` 引用。公司模型只讀快照，不直接讀 Tokenomics。
- **產生快照**：`python3 tools/tokenomics/import_tokenomics.py --tokenomics <Tokenomics 唯讀 clone> --names <公司>/data/tokenomics_names.txt --out <公司>/data/tokenomics_snapshot_v<版本>.json`
  （唯讀 clone：`GIT_LFS_SKIP_SMUDGE=1 git clone --depth 1 https://github.com/YinchenChang/tokenomics <路徑>`；讀 `model/CURRENT` 指向的 xlsx。）
- **重現檢查**：`python3 tools/tokenomics/import_tokenomics.py --check <快照.json> [--tokenomics <clone>]`：重新讀取快照記錄的同一個 xlsx，任何值差異（相對誤差 > 1e-9）即失敗；找不到 clone 時警告並略過（`TOKENOMICS_DIR` 環境變數可指定路徑）。
- **名稱規則**：只接受 `IF_`（不含 `IF_Hdr*`）與 `L1_`，其他名稱報錯；不得引用 Inputs、DC_Cost 等內部頁。名稱清單一行一個，行尾 `optional`（或 `--optional`）＝Tokenomics 尚未提供時記為 `{"missing": true}` 並警告。世代 × 成本情境範圍依 `IF_HdrGen`、`IF_HdrCost` 拆成 `{世代: {低成本, 基準, 高成本}}`；模型主值一律取「基準」，低／高只作敏感度。Tokenomics 的「每 GW」＝IT 關鍵電力；公司 MW 口徑不同時在公司端換算。
- 快取值缺漏時自動以 LibreOffice headless 重算暫存副本再讀（`--force-recalc` 可強制；v5.24 實測重算值＝快取值）。
