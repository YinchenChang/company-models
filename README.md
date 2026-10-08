# company-models

個股收支與評價模型（算力收支 Project）。每家公司一個資料夾，各自完整可建置（CRWV 模板的副本＋該公司的 company.json 與資料）。

| 資料夾 | 公司 | 來源模板 | 狀態 |
|---|---|---|---|
| `nebius/` | Nebius（NBIS） | `YinchenChang/crwv-model` @ `01b13ad`（v4.5） | v0.1 |
| `coreweave/` | CoreWeave（CRWV） | `YinchenChang/crwv-model` @ `01b13ad`（v4.5；2026-10-07 起在此維護，改接 Tokenomics 中，見 `coreweave/docs/workorders/`） | 改造中 |
| `openai/` | OpenAI（未上市） | `YinchenChang/openai-model` @ `13c16f0`（v0.6-P1；原 repo 停用，歷史留在原處） | 建置中（v0.6） |
| `oracle/` | Oracle（ORCL） | `nebius/` @ `642d144`（CRWV v4.5＋Nebius v0.1b） | v0.1 |
| `minimax/` | MiniMax（00100.HK） | 比照 `openai/` v0.6 結構（MW 主軸、TK v5.24）；獨立的 Python→Excel 產生器 | v0.1 |
| `zhipu/` | 智譜（智譜華章，02513.HK） | `openai/` v0.6（`claude/openai-s6-release` @ `e91df57`）的架構；骨架經 Anthropic A0 | v0.1（r2） |
| `anthropic/` | Anthropic（未上市） | `openai/`（v0.6，`claude/openai-s6-release` @ `e91df57`）的架構與工具鏈 | v0.1（成品：`anthropic/dist/20261008_Anthropic收支模型_v0_1.html`／`.xlsx`；PR #22 待合併） |
| `whitefiber/` | WhiteFiber（WYFI） | `oracle/` @ `e4540c3`（Oracle v0.2 分支） | v0.1 |

規則：
- 事實來源與分層：算力相關的產業與物理層資料取自 Tokenomics（`YinchenChang/Tokenomics` 的 Interface／L1／SRC，X8）；公司專屬資料放各公司的 `company.json` 與 `data/`。
- 範圍：做到 CRWV 目前水準（v4.5）；CRWV 尚未完成的待辦（5a 精簡、5b 只讀檢視器、折舊修正、機率加權目標價、GPU 批次與續約價格衰退），各公司也先不做（Andy 2026-10-07）。
- 流程：chat 端（Claude）寫工作單（`<公司>/docs/workorders/`）→ 建置代理執行 → chat 端審查並合併；成品（HTML、Excel）由 chat 端放到 Andy 的 Google Drive 對應資料夾。
- 各公司資料夾內的 `CLAUDE.md` 沿用各自來源模板的工作守則；與本 README 衝突時，以本 README 的流程與範圍為準。
- CI：各公司一個 workflow，只在該資料夾有變動時執行（`.github/workflows/<公司>-*.yml`）。
- 成品的 Drive 位置：`02-算力收支/<公司>/`（資料夾 2026-10-07 由錯字「02-算利收支」改名）。
