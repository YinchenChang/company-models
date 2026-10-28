# Microsoft v0.3.1 進度（季度滾動 FY27Q1；工作單 `mag/docs/workorders/20261029_MAG_v0.3.1_季度滾動_2026Q3.md`）

分支 `claude/microsoft-v0.3.1`。

| 段 | 內容 | 狀態 |
|---|---|---|
| 1 | 事實總帳 FY27Q1：`microsoft/facts_fy27q1.py` → `data/microsoft_facts_FY27Q1_20261029.json`（119 筆；雙讀核對 67 筆，不一致 2 筆皆為口徑不同、已註明）；Tokenomics v21.7 快照 `data/tokenomics_snapshot_v21.7.json`（與 v21.4 逐名稱比對：0 差異） | 完成 |
| 2 | Q 頁加 FY27Q1（八季 FY25Q2–FY27Q1）、FY2027＝Q1 實績＋Q2–Q4 模型、評價日 2026-09-30 | 進行中 |
| 3 | Tokenomics 升級與 NonNV（Maia 改讀 `IF_NonNV_*`） | 待辦 |
| 4 | 指引與 C10 重解 | 待辦 |
| 5 | 現價（2026-10-28 收盤 531.01：MarketBeat、247wallst 兩源）與共識（S&P 605.12／57 位、MarketBeat 596.07／49 位，皆 2026-10-27） | 資料已取 |
| 6–9 | 證偽、重建核對、報告、README、PR | 待辦 |

下一步：改 `build_microsoft.py`（Q 頁窗口、FY2027 實績＋剩餘、評價日、NonNV 開關、--set），寫 `make_v03.sh`／`check_v03.py`／`attrib_v031.py`。
