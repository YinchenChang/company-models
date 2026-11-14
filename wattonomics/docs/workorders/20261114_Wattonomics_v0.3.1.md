# 工作單：Wattonomics v0.3.1＝財報季證據＋④ 實績下限重算＋Compare 依三層基準協調規則＋證據比例改全市場口徑

- 日期：2026-11-14（台北）；分支 `claude/wattonomics-v0.3.1`（由 main `5ca29ac`）
- 授權：Andy 2026-11-14 10:12「本串接著做／其他照建議或預設」；10:35「請盡可能地平行處理」。判斷類採下列預設，寫進 Decisions（新編號接續 GL01-*），可撤回。工作單沒寫到的判斷不自行決定，列報告「待判斷」。
- 依據：三層基準協調規則 v1.1 `docs/plan/三層基準協調規則.md`（C3 證據定義、分母、已知限制）；步驟 3 橋接表 `siliconomics/tools/baseline_bridge.py`；Wattonomics v0.3 交接（以下 Z 項即其待辦第 1 項）。
- 現況：main `wattonomics/model/CURRENT`＝`20261105_Wattonomics_v0.3.xlsx`；Siliconomics main＝v0.18；TK 快照 v21.10。

| # | 改動 |
|---|---|
| Z1 | **財報季證據**（證據由並行查證代理寫在 `/mnt/user-data/outputs/w031/evidence.md`，完成時 chat 端通知；在那之前先做 Z2–Z7）：IREN Q1 FY27、CoreWeave 2026Q3（11-10）、Nebius 2026Q3（11-12）→ T2／Delay 列；EQ#5（Microsoft 第三方容量延至 FY27H2）、EQ#7（IREN Childress）、EQ#8（Nebius）追蹤列更新。新證據列須經未參與代理核對才入基準（chat 端另派；建置代理先以「待核對」標示、不入基準） |
| Z2 | **④ 覆蓋缺口在 v0.3 上重算（W-2 實績下限讀法）：** Compare 的 ④ 改為兩欄並列——原值 4.81（1.84／4.81／8.28）與實績下限讀法 3.03（0.06／3.03／7.36；(a) 七家多花由 1.84 改 10-Q 實績外推下限 0.06〔0–0.92〕）；扣 Oracle 重疊 0.77 的讀法另列；每項寫證據等級與升級觀察點；CoreWeave 項標「未取得」（不是 0），Z1 證據到後補。④ 仍不入基準（判斷題預設） |
| Z3 | **Compare 依協調規則 v1.1：** R4 統一讀 Siliconomics v0.18（SIL 快照升 v0.18；`SIL_In` 只讀 `*_Base`、Cap、R7 內容、逐買方份額）；刪「2028–29 部分已歸因；兩邊都不錨定，只 WARN」的文字與判定，改為逐年列協調狀態（Z4 證據比例、門檻 50%、對帳中標示）；並列 W1＋W2 調整後比值（④ 兩讀法兩欄）；W2 揭露；NVIDIA 對帳每 GW 營收改區間 低 19.67／基準 26.13（W-3），NVIDIA 指引只作參照列、不作校準目標，2026 實績校準保留；五家（Microsoft、Alphabet、Amazon、Meta、Oracle）資本支出與 Siliconomics 顆數假設並排（W-4） |
| Z4 | **證據比例改規則 v1.1 口徑（新名稱，舊 `WIF_Coverage_*` 保留不改義）：** `WIF_Evidence_<yyyy>`＝（站點證據年新增 GW＋七家已公布資本支出實績支撐的年新增 GW）÷（主路徑年新增 GW＋④ 覆蓋缺口〔實績下限讀法基準〕＋W2）。七家實績部分：2026 以已公布 10-Q／10-K 的 Q1–Q3 AI 資料中心資本支出 ÷ 該家全年路徑資本支出 × 該家年新增 GW（甲2 時差同主路徑）；尚無實績的年份＝0。輸出 2026–2030；分子各項並列在 Deploy 或 Compare 讓人看得到。證據等級對照規則 v1.1 表（實績／已簽約／已確認站點），等級 3 只撐到上線年 |
| Z5 | 查五家個股模型 2028–29 資本支出怎麼推得（`import_gap_models.py` 讀的檔）；只是向共識收斂者在 In 或 Decisions 標 Assumed，報告列明 |
| Z6 | R7 口徑：PR #340 已合併 → `SIL_In` 改讀 Siliconomics 最新快照（v0.18）。2028–29 診斷與 W-2 兩組檔若在 repo 外取不到，報告寫「未入 repo，待 Andy 電腦上線後由 chat 端放」，不重建 |
| Z7 | README、Interface、Decisions、check.py：`WIF_Evidence_*` 存在且 0–1；Compare 不再出現「兩邊都不錨定」；SIL 快照版本＝v0.18 |

不做：v0.3.2 W-1 七家以外買方路徑（另一張工作單；並行查證中）；任何基準 GW 值改變（本版預期 `WIF_MktGW_*` 不變，若變動須在報告逐項說明）；不動 gh-pages。
建置：build → soffice 重算 → export_snapshot → export_explorer → check.py → diff_snapshot（與 v0.3 逐格比，差異須都在本單範圍）→ r4_attrib。報告寫 `wattonomics/docs/verify/20261114_v0.3.1_build.md`（繁體中文、非工程語言；表 1 本單各項完成狀況、表 2 範圍外差異、待判斷）。開 PR（不合併）。
