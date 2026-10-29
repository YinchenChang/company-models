# Microsoft v0.3.1 進度（季度滾動 FY27Q1；工作單 `mag/docs/workorders/20261029_MAG_v0.3.1_季度滾動_2026Q3.md`）

分支 `claude/microsoft-v0.3.1`。

| 段 | 內容 | 狀態 |
|---|---|---|
| 1 | 事實總帳 FY27Q1：`microsoft/facts_fy27q1.py` → `data/microsoft_facts_FY27Q1_20261029.json`（119 筆；雙讀核對 67 筆，不一致 2 筆皆為口徑不同、已註明）；Tokenomics v21.7 快照 `data/tokenomics_snapshot_v21.7.json`（與 v21.4 逐名稱比對：0 差異） | 完成 |
| 2 | Q 頁加 FY27Q1（八季 FY25Q2–FY27Q1）、FY2027＝Q1 實績＋Q2–Q4 模型、評價日 2026-09-30 | 完成 |
| 3 | Tokenomics 升級與 NonNV（Maia 改讀 `IF_NonNV_*`；開關 `nonnv`） | 完成 |
| 4 | 指引與 C10 重解（`attrib_v031.py --solve`：MW 速度 2,527.2 → 2,726.5；NonNV 舊口徑 2,645.1） | 完成 |
| 5 | 現價（2026-10-28 收盤 531.01：MarketBeat、247wallst 兩源）與共識（S&P 605.12／57 位、MarketBeat 596.07／49 位，皆 2026-10-27） | 資料已取 |
| 7 | `make_v03.sh`、`check_v03.py` 新增；重建＋核對全過（加權目標價 384.10、−27.7%、放空；v0.3 重現 385.56 差 0） | 完成 |
| 6、8、9 | 證偽與命題、報告 `20261029_microsoft_v0.3.1_季度滾動.md`、README（本家與根目錄本家列）、PR 留言與轉 ready | 完成 |

全部完成（2026-10-29）。下一步：chat 端審查與獨立代理核對；Andy 決定報告第 12 節各項。
