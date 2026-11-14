# 智譜（智譜華章，02513.HK）個股模型

智譜每 VR 等值 GW 的年營收能否覆蓋每 GW 年全成本；若不能，缺口由誰、以什麼條件融資（FY2025–FY2030）。架構比照 `openai/`（v0.6），算力物理取自 Tokenomics；金額單位人民幣億元，命題另列美元口徑與 OpenAI 並排。

- 現行版本：**v0.1.5**（`model/20261115_Zhipu_v0.1.5.xlsx`；成品 `dist/20261115_智譜個股模型_v0_1_5.html／.xlsx`、版面 v1.3 個股頁 `dist/20261115_智譜個股頁_v0_1_5.html`；見下方 v0.1.5 節）。前版 v0.1.3（`model/20261104_Zhipu_v0.1.3.xlsx`；見 `model/CURRENT`）；成品 `dist/20261104_智譜個股模型_v0_1_3.html／.xlsx`（依 `docs/plan/個股模型原則_命題加營收分部.md` 第 1 節改稱「個股模型」）。v0.1.3＝文件與核對小修：Excel README A1 改為「個股模型 v0.1.3（2026-11-04）」、`CLAUDE.md` 命名、parity 情境 `k2_rev_total_mismatch` 預期加 C101；數值逐格不變（報告 `docs/reports/20261104_v0.1.3.md`）。
- 官方報告分部（部署方式）：本地化部署（②企業級智能體＋③企業級通用大模型＋④技術服務及其他）、雲端部署（①開放平台及 API）。宣告檔 `segments.json`；對帳在 Revenue 第七節（FY2025 對年報、1H2026 對中期公告子項加總，容差 Inputs `seg_recon_tol`），Checks C100–C103；`python3 tools/segment_check/check_segments.py zhipu`（repo 根目錄執行）須為「符合」。
- 規格：`docs/plan/20261008_Zhipu_v0.1_規格.md`
- 交接：`docs/handoff/Zhipu_handoff.md`；進度：`docs/reports/20261008_zhipu_進度.md`（v0.1.2：`docs/reports/20261030_v0.1.2_進度.md`）
- 工作規範：`CLAUDE.md`


## v0.1.5（2026-11-15）：個股 HTML 版面 v1.3＋三層協調（讀取層級、C0 公司實績）

**依據：** `docs/workorders/20261114_第二批13家_v1.3與三層協調.md`；`docs/plan/三層基準協調規則.md` v1.4.2 C0 第 2 點；`docs/plan/個股HTML版面規格.md` v1.3；`tools/coord/README.md`。樣板：Anthropic v0.2.2、OpenAI v0.6.5（PR #524）。

**做法：** `builder/z5.py` 新增 PL（標準損益表）、Coord（讀取層級、C0、判定摘要、參考）兩頁，只列示；`tools/build_html_v13.py` 產生版面 v1.3 個股頁（命題 → 估值〔不適用〕→ 敏感度〔各情境對 FY2030 覆蓋率〕→ 損益 → 驅動因子〔分部營收推導〕→ 算力 → 資金 → 與 OpenAI 並排 → 對帳〔C0 表〕→ 資料狀態），數字與一頁摘要同一套取值（LibreOffice 重算、情境重算）與 engine 比對。
- **估值：本模型沒有目標價與評等**（Reverse 只作反向對照：市值隱含營收）。智譜已上市、有市價，但 G 節四個目標價、H 節加權目標價敏感度、反向 DCF 章無從呈現；估值章標「不適用（本模型未建目標價）」並帶 `data-unlisted="1"` 沿用規格 J 節豁免（J 節原文只寫未上市，需主控／Andy 決定是否擴充，或另建估值頁）。

**實績與出處（C0）：**
| 項目 | 期間 | 公司實績 | 出處 | 層級隱含 | 判定 |
|---|---|---|---|---|---|
| 營收總額 | 2026-01～06（主）；季營收未揭露（第二讀數無） | 9.5389 億元 | 2026 中期業績公告（2026-08-31，港交所 2026083101539），SRC_ZP_177 | 無（Wattonomics 無智譜玩家） | 無可比實績（層級無此玩家） |
| 資本開支 | 1H2026 | 3.906 億元（主要為收購北京紅鑽） | 同上，SRC_ZP_209 | 無 | 無可比實績（層級無此玩家） |
| 在役／簽約算力 GW | — | 公司未揭露（只有國產晶片十萬張級、1 GW 資料中心等說法） | SRC_ZP_452、458（Interested-party） | 無 | 無可比實績（公司未揭露、層級無此玩家） |
| 應付＋預付算力服務費 | 2026-06-30 | 26.71 億元（14.27＋12.44） | 中期公告資產負債表，SRC_ZP_218、219 | 無 | 無可比實績（層級無此玩家） |

**判定結果：** 可判定 0、對帳中 0、無可比實績 4；無同源列（Wattonomics 快照全文無智譜，層級不讀本模型）。讀取層級：Tokenomics 快照 v19.0（現行 v22.0，本輪不換）。

**驗證：** builder 重建（restore matched 3,585、unmatched 0；e6 0）；`tools/xlsx_diff.py` 對 v0.1.4：只差 README A1、Checks A1 版本字串與新增格（README「v0.1.5 變動」列、PL、Coord）；test_builder 14 通過；parity（結構、具名範圍、base 與兩個預期 ERR 情境）通過，完整情境以 CI 為準；test_html（含個股頁數字比對）；`check_layout_v13 --xlsx`：main 版因 G、H、反向 DCF 必修不過（本模型無估值），PR #524 版（J 節豁免）全過；`tools/baseline_coordination/check.py` 通過。
