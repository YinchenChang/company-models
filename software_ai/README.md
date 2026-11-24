# software_ai/ — AI 軟體應用

本資料夾由 AI 軟體應用 專案擁有（歸屬見 `docs/OWNERSHIP.md`）。公開 repo：不放任何個人情境，檔案內不寫人名（決策者寫「專案負責人」）。

| 位置 | 內容 |
|---|---|
| `unit_economics/` | AI 對軟體單位經濟（五原型：S 按席位、U 按用量、L 按人力、A 按廣告、P 按防護範圍） |
| `<公司>/` | 個股模型（`build_<公司>.py`、`data/` 研究包與 facts、`dist/` 現行 xlsx＋html；舊版在 `dist/_待刪除/`） |
| `data/relmult_hist/` | 歷史相對倍數庫（P/S、PSG）與同業現價 P/S |
| `map/` | AI 軟體地圖（見 `map/README.md`）；`map/profiles/` 為公司檔案 |
| `docs/workorders/` | 工作單（規格與決策的權威紀錄） |

規格沿用 repo 共用文件：`docs/plan/個股模型原則_命題加營收分部.md`、`docs/plan/個股DCF標準規格.md`、`docs/plan/個股HTML版面規格.md`；本專案的補充規則在工作單 G2–G10。
推論單價、token 價格只讀 Tokenomics 快照，不自行重估。
