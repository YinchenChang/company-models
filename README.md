# company-models

個股收支與評價模型（算力收支 Project）。每家公司一個資料夾，各自完整可建置（CRWV 模板的副本＋該公司的 company.json 與資料）。

| 資料夾 | 公司 | 來源模板 | 狀態 |
|---|---|---|---|
| `nebius/` | Nebius（NBIS） | `YinchenChang/crwv-model` @ `01b13ad`（v4.5） | v0.1 |
| `coreweave/` | CoreWeave（CRWV） | `YinchenChang/crwv-model` @ `01b13ad`（v4.5；2026-10-07 起在此維護，改接 Tokenomics 中，見 `coreweave/docs/workorders/`） | 改造中 |
| `openai/` | OpenAI（未上市） | `YinchenChang/openai-model` @ `13c16f0`（v0.6-P1；原 repo 停用，歷史留在原處） | 建置中（v0.6） |
| `oracle/` | Oracle（ORCL） | `nebius/` @ `642d144`（CRWV v4.5＋Nebius v0.1b） | v0.1 |
| `minimax/` | MiniMax（00100.HK） | 比照 `openai/` 結構（MW 主軸、P×Q 回本比值、TK v19.0 官方快照）；獨立的 Python→Excel 產生器 | v0.2 |
| `zhipu/` | 智譜（智譜華章，02513.HK） | `openai/` v0.6（`claude/openai-s6-release` @ `e91df57`）的架構；骨架經 Anthropic A0 | v0.1（r2） |
| `anthropic/` | Anthropic（未上市） | `openai/`（v0.6，`claude/openai-s6-release` @ `e91df57`）的架構與工具鏈 | v0.1.4（成品：`anthropic/dist/20261022_Anthropic收支模型_v0_1_4.html`／`.xlsx`） |
| `whitefiber/` | WhiteFiber（WYFI） | `oracle/` @ `e4540c3`（Oracle v0.2 分支） | v0.1 |
| `amazon/` | Amazon（AMZN） | `oracle/`（v0.2）→ MAG 共用引擎（Amazon v0.1b；共用文件見 `mag/`） | v0.2（Tokenomics v5.31） |
| `alphabet/` | Alphabet（GOOGL） | `oracle/`（v0.2）→ 套用 Amazon 引擎（amazon@`41f9a54`，v0.2） | v0.2 |
| `microsoft/` | Microsoft（MSFT） | `oracle/`（v0.2）→ 套用 Amazon 引擎（amazon@`41f9a54`，v0.2） | v0.2 |
| `markets/` | 市場層（AI 半導體）：每 GW 晶片內容橋接表等，見 `markets/README.md` | 新建（AI 半導體 Project） | chip_bridge v0.2.1、network v0.2.2、equipment v0.1.3、overview v0.6（Tokenomics v20.1；memory v0.4） |
| `skhynix/` | SK 海力士（000660.KS） | 記憶體樣板 `tools/memory_model`（AI 半導體 Project） | v0.4.1（HBM 位元隨情境；Val 頁含十年倍數與規則值） |
| `micron/` | 美光（MU） | 記憶體樣板（複製測試） | v0.3.1 |
| `samsung/` | 三星電子（005930.KS） | 記憶體樣板（集團公司：只建模記憶體，其餘事業為輸入） | v0.2.1 |
| `tsmc/` | 台積電（2330.TW） | 獨立產生器（CoWoS 供需上限＋損益估值；AI 半導體 Project） | v0.2.1 |
| `markets/equipment/` | 半導體設備市場模型：記憶體、代工邏輯資本支出 → 設備需求；三情境、對帳、瓶頸表 | 新建（AI 半導體 Project 步驟 5） | v0.1.3（成品：`markets/equipment/dist/20261023_半導體設備市場模型_v0.1.3.xlsx`；台積電設備占比查證，數值不變） |
| `advantest/` | Advantest（6857.T）：SoC／記憶體測試機收支與 P/E 估值；三情境、對帳、敏感度；見 `advantest/README.md` | 獨立產生器（比照 `tsmc/`；AI 半導體 Project 步驟 5 第一家公司） | v0.2.2（成品：`advantest/dist/20261023_Advantest模型_v0.2.2.xlsx`、`.html`；設備股共用循環位置規則；HTML 展示層） |
| `disco/` | Disco（6146.T）：研削／切割裝置與消耗品收支、P/E 估值；HBM 工序與 CoWoS 片數驅動、三情境、混合鍵合開關、對帳、敏感度；見 `disco/README.md` | 獨立產生器（比照 `advantest/`；AI 半導體 Project 步驟 5 第二家公司） | v0.1.3（成品：`disco/dist/20261023_Disco模型_v0.1.3.xlsx`；FY2026 錨定公司資料、設備股共用循環位置規則） |
| `nvidia/` | NVIDIA（NVDA）：運算顆數 × 每顆售價＋網通附掛的收支與 P/E 估值；三情境、FY2027 錨點、對帳、敏感度、金額份額可重用函式；見 `nvidia/README.md` | 獨立產生器（比照 `advantest/`；AI 半導體 Project 運算層第一個樣板，之後 AVGO、AMD 複製） | v0.1.1（成品：`nvidia/dist/20261023_NVIDIA模型_v0.1.1.xlsx`、`.html`；認列顆數係數 k_rec 使 FY2027 與 FY2028 基準一致；10 年落後 P/E 規則值不適用；循環位置用共用規則的 lag 版） |
| `arista/` | Arista Networks（ANET）：AI 後端乙太網路交換器市場 × Arista 份額＋非 AI 的收支與 P/E 估值；三情境、FY2026 錨點、對帳、敏感度；Arista 份額為本專案權威；見 `arista/README.md` | 獨立產生器（比照 `nvidia/`、`avgo/`；AI 半導體 Project 網通系統端樣板） | v0.1（成品：`arista/dist/20261024_Arista模型_v0.1.xlsx`、`.html`） |
| `delta/` | 台達電（2308）：部署 GW ×（AI 電源每 kW 內容 × 台達電份額＋液冷每 GW 營收）＋非 AI 的收支與 P/E 估值；三情境、FY2026 錨點、對帳、敏感度；AI 電源內容、份額、液冷每 GW 為本專案權威；見 `delta/README.md` | 獨立產生器（比照 `arista/`；AI 半導體 Project 電源與散熱樣板） | v0.1（成品：`delta/dist/20261024_台達電模型_v0.1.xlsx`、`.html`） |
| `quanta/` | 廣達（2382）：部署 GW × 每 GW ODM 整櫃價值 × 廣達有效份額 ×（1 − 客供比例）＋非 AI 的收支與 P/E 估值；三情境、FY2026 錨點、對帳、敏感度；組裝市場、份額、ODM 每櫃價格、客供比例為本專案權威；見 `quanta/README.md` | 獨立產生器（比照 `delta/`；AI 半導體 Project 伺服器組裝樣板） | v0.1（成品：`quanta/dist/20261024_廣達模型_v0.1.xlsx`、`.html`） |
| `markets/overview/` | AI 半導體總覽儀表板 | 彙整橋接表、記憶體、網通、台積電模型；估值規則；記憶體三家 P/E 對照、美光買回加回 P/B；設備分頁（設備市場、Advantest、Disco；設備股共用循環位置規則） | v0.6 |

規則：
- 事實來源與分層：算力相關的產業與物理層資料取自 Tokenomics（`YinchenChang/Tokenomics` 的 Interface／L1／SRC，X8）；公司專屬資料放各公司的 `company.json` 與 `data/`。
- 範圍：做到 CRWV 目前水準（v4.5）；CRWV 尚未完成的待辦（5a 精簡、5b 只讀檢視器、折舊修正、機率加權目標價、GPU 批次與續約價格衰退），各公司也先不做（Andy 2026-10-07）。
- 流程：chat 端（Claude）寫工作單（`<公司>/docs/workorders/`）→ 建置代理執行 → chat 端審查並合併；成品（HTML、Excel）由 chat 端放到 Andy 的 Google Drive 對應資料夾。
- 各公司資料夾內的 `CLAUDE.md` 沿用各自來源模板的工作守則；與本 README 衝突時，以本 README 的流程與範圍為準。
- CI：各公司一個 workflow，只在該資料夾有變動時執行（`.github/workflows/<公司>-*.yml`）。
- 成品的 Drive 位置：`02-算力收支/<公司>/`（資料夾 2026-10-07 由錯字「02-算利收支」改名）。

## 共用工具：公司模型 HTML 展示層（`tools/company_html/`）
- Excel 仍是計算事實的來源；`<公司>/html.json` ＋ 共用產生器 → `<公司>/dist/<日期>_<公司>模型_v<版本>.html`（單頁自含、每個數字標註來源儲存格）。檢核（重現性、抽數、無手寫數字）在 `ai-semis-build` CI。說明見 `tools/company_html/README.md`。目前：Advantest、Disco、SK 海力士、美光、三星、台積電、NVIDIA、Arista、台達電、廣達。風格比照 AI 半導體價值鏈儀表板：白底、青綠主色，不做深色模式。

## 共用工具：Tokenomics 取數（`tools/tokenomics/import_tokenomics.py`）
- **用途**：把 Tokenomics 的 Interface（`IF_`、四層瀑布 `IFW_`、用途欄 `IFC_`）、L1（`L1_`）與來源（`SRC_`，限 Active）具名範圍讀成版本固定的快照 JSON（記錄檔名、`model/CURRENT`、commit、擷取日；每個值記名稱、世代、成本情境、單位、工作表!儲存格），供各公司 `company.json` → `tokenomics.snapshotFile` 引用。公司模型只讀快照，不直接讀 Tokenomics。
- **產生快照**：`python3 tools/tokenomics/import_tokenomics.py --tokenomics <Tokenomics 唯讀 clone> --names <公司>/data/tokenomics_names.txt --out <公司>/data/tokenomics_snapshot_v<版本>.json`
  （唯讀 clone：`GIT_LFS_SKIP_SMUDGE=1 git clone --depth 1 https://github.com/YinchenChang/tokenomics <路徑>`；讀 `model/CURRENT` 指向的 xlsx。）
- **重現檢查**：`python3 tools/tokenomics/import_tokenomics.py --check <快照.json> [--tokenomics <clone>]`：重新讀取快照記錄的同一個 xlsx，任何值差異（相對誤差 > 1e-9）即失敗；找不到 clone 時警告並略過（`TOKENOMICS_DIR` 環境變數可指定路徑）。
- **名稱規則**（Tokenomics 下游資料契約第 2 條第 1 項）：只接受 `IF_`（不含 `IF_Hdr*`）、`IFW_`、`IFC_`、`L1_`、`SRC_`，其他名稱報錯；`IFW_` 須四層（`_100`／`_Util`／`_Prod`／`_Life`）一起取（契約第 2 條第 2 項）；`IFC_Use` 等用途欄存成 {列上的 IF_／IFW_ 名稱: 值}，標「上限，不得作預測」者不得當收入預測；`SRC_` 只取狀態 Active 的紀錄，另存同列的低／高、口徑、日期、出處、來源等級；不得引用 Inputs、DC_Cost 等內部頁。名稱清單一行一個，行尾 `optional`（或 `--optional`）＝Tokenomics 尚未提供時記為 `{"missing": true}` 並警告。世代 × 成本情境範圍依 `IF_HdrGen`、`IF_HdrCost` 拆成 `{世代: {低成本, 基準, 高成本}}`；模型主值一律取「基準」，低／高只作敏感度。Tokenomics 的「每 GW」＝IT 關鍵電力；公司 MW 口徑不同時在公司端換算。
- 快取值缺漏時自動以 LibreOffice headless 重算暫存副本再讀（`--force-recalc` 可強制；v5.24 實測重算值＝快取值）。
