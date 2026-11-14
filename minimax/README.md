# MiniMax 收支模型（company-models/minimax）

MiniMax（00100.HK，上海稀宇科技）收支模型，現行 **v0.2.2**（2026-11-15；v0.2 為 2026-10-22）。比照 `openai/`（v0.6／v0.6.1）的分層與 MW 主軸建立；MiniMax 為上市公司，另加 Funding 與 Valuation；v0.2 新增 P×Q 回本比值頁。

命題：MiniMax 每 MW 算力的營收能否覆蓋每 MW 全成本；若不能，缺口何時出現、由誰、以什麼條件補。FY2025–FY2030，曆年制。

- Excel 為唯一計算引擎：`build_xlsx.py` 只產生結構與公式（藍字輸入、黑字公式），數值由 LibreOffice 重算。
- 一鍵建置：`scripts/build_all.sh [日期] [版本]` → `dist/<日期>_MiniMax收支模型_<版本>.xlsx／.html`（需 `libreoffice-calc-nogui`、openpyxl）。
- 敏感度：`scripts/sensitivity.py`（逐項改寫 Inputs、重算，寫入 Sensitivity 頁快照）。
- 原始數據與來源：`data/research/financials.md`、`operations.md`（研究代理 2026-10-08 蒐集；一手／二手已標註）。
- Tokenomics 取數：`scripts/tk_snapshot.py`（讀 `model/CURRENT` 具名範圍）、名稱檔 `data/tokenomics_names.txt`、指標檔 `data/tokenomics_snapshot_CURRENT`；歸因 `scripts/tk_attribution.py`；驗收 `scripts/ci_check.py`；CI `.github/workflows/minimax-build.yml`。
- 工作規範：`CLAUDE.md`；變更紀錄：`CHANGELOG.md`；進度：`docs/reports/20261022_minimax_進度.md`。
- 報告：`docs/reports/20261008_minimax_v0.1.md`（v0.1）、`20261022_minimax_M1.md`、`_M2.md`、`20261022_minimax_v0.2.md`（v0.2 成品）＋ `_對照.xlsx`；命題與驅動對照表 `docs/plan/20261022_MiniMax_命題與驅動對照表.md`；交接檔 `docs/handoff/MiniMax_handoff.md`。
- HTML 驗收：`scripts/check_html.py`（每個數字對照 Excel、無外部引用；已納入 `build_all.sh` 與 CI）。

| 頁 | 內容 |
|---|---|
| README | 命題、驅動→推導對照、圖例、已套用的預設 |
| SRC_MM | 公司財務、市場與營運原始數據（112 列；標記、一手／二手、來源） |
| TK_Link | Tokenomics 官方快照（`data/tokenomics_snapshot_*.json`；現用 v19.0，含檔名、提交、SHA-256、讀取日）；12 格基準值＋IF_HoldEcon 全 15 欄 |
| Inputs | 假設 32 項（值、低、高、標記、說明） |
| Demand | API 實收單價、計費 token |
| Revenue | 企業端（API）、消費端（產品與收費方式按 9M25 結構示意） |
| Compute | 機隊組合、每 MW 產能、η 校準（2025、1H26）、推論／訓練 MW、VR 等值；D2 租價並列、η 對照、自建集群註記 |
| PnL | 損益（算力／人事／其他）、每 MW 營收 vs 每 MW 全成本 |
| PxQ | 回本比值＝營收 ÷（年均上線 MW × 每 MW 年持有成本）、實際／所需 m×L、分項比值、持平 MW、M3.1 用量倍數、比值對 I02 的結構分解 |
| Funding | 現金路徑、需新融資、稀釋、已投入資金 |
| Valuation | 市值、EV、同業、反向檢驗 |
| Checks | 13 項檢查＋錯誤值計數（皆為 0） |
| Sensitivity | 31 個單項情境（快照值；含 M3.1 降價與 I02 租價證據價） |
| PL | 標準損益表 FY2023–FY2030（v0.2.2；非 IFRS 經調整；版面 v1.3 D 節） |
| Coord | 三層協調：讀取層級、C0（實績）公司實績 vs 層級隱含值（v0.2.2；規則 v1.4.2） |

## v0.2.2（2026-11-15）：個股 HTML 版面 v1.3＋三層協調（讀取層級、C0 公司實績）

**依據：** `docs/workorders/20261114_第二批13家_v1.3與三層協調.md`（Andy 11-14 16:53、16:55、17:40、21:29）；`docs/plan/三層基準協調規則.md` v1.4.2 C0 第 2 點；`docs/plan/個股HTML版面規格.md` v1.3；`tools/coord/README.md`。樣板：Anthropic v0.2.2、OpenAI v0.6.5（PR #524）。

**做法：**
- `coord_pl.py`（由 `build_xlsx.py` 呼叫）新增 PL、Coord 兩頁，只列示、不被既有計算頁引用。
- 個股模型頁 `build_html_v13.py`：章節 命題 → 估值（不適用）→ 敏感度（對至 FY2030 累計需新融資）→ 損益（標準損益表）→ 驅動因子（分部營收推導：企業端＝token × 單價、消費端成長）→ 算力與每 MW 收支 → 資金 → 回本比值 → 對帳（C0 表）→ 資料狀態；除估值、對帳、資料狀態外每章有圖。數字全取自 xlsx（`scripts/check_html_v13.py` 逐一核對 data-ref）。
- **估值：本模型沒有目標價與評等**（Valuation 頁只有市值、EV 與反向檢驗）。MiniMax 已上市、有市價，但 G 節四個目標價、H 節加權目標價敏感度、反向 DCF 章無從呈現；估值章標「估值：不適用（本模型未建目標價）」並帶 `data-unlisted="1"`（沿用規格 J 節的檢查豁免；J 節原文只寫未上市，需主控／Andy 決定是否擴充到「上市但模型未建估值」，或另建 DCF／同業估值）。

**實績與出處（C0）：**
| 項目 | 期間 | 公司實績 | 出處 | 層級隱含 | 判定 |
|---|---|---|---|---|---|
| 營收（IFRS 總營收） | 2026-01～06（主）；2026 Q2（第二，Derived） | 116.573 $M；Q2 約 75.2 $M | 2026 中期業績公告（2026-08-26，港交所 2026082600680）；Q2 由 1H 與季增 81.8%（業績會轉述，InfoQ）推得 | 無（Wattonomics 無 MiniMax 玩家） | 無可比實績（層級無此玩家） |
| 資本支出（物業廠房設備＋預付款增加） | 1H2026 | 322.3 $M | FY2025 年度業績公告、2026 中期業績公告資產負債表 | 無 | 無可比實績（層級無此玩家） |
| 在役／簽約算力 MW | — | 公司未揭露 | — | 無 | 無可比實績（公司未揭露、層級無此玩家） |
| 算力承諾（阿里雲 2026 年度上限） | 2026 | 300 $M（人民幣 21 億） | 關連交易上限修訂（媒體轉述） | 無 | 無可比實績（層級無此玩家） |

**判定結果：** 可判定 0、對帳中 0、無可比實績 4；無同源列（Wattonomics 快照全文無 MiniMax，層級不讀本模型）。讀取層級：Tokenomics 快照 v19.0（現行 v22.0，本輪不換）。

**驗證：** 建置 → LibreOffice 重算（錯誤 0、Checks 合計 0）→ 敏感度 → HTML（收支頁 506 個數字、個股頁 734 個 data-ref 全對 xlsx）；`tools/xlsx_diff.py` 對 v0.2.1：只差版本字串與新增格；`check_layout_v13 --xlsx`：main 版因 G、H、反向 DCF 必修不過（本模型無估值），PR #524 版（J 節豁免）全過；`tools/baseline_coordination/check.py` 通過。`check_company_html` 不適用（收支模型無 html.json，同 anthropic／openai）。
