# 智譜（智譜華章，02513.HK）個股模型

智譜每 VR 等值 GW 的年營收能否覆蓋每 GW 年全成本；若不能，缺口由誰、以什麼條件融資（FY2025–FY2030）。架構比照 `openai/`（v0.6），算力物理取自 Tokenomics；金額單位人民幣億元，命題另列美元口徑與 OpenAI 並排。

- 現行版本：**v0.1.5**（`model/20261115_Zhipu_v0.1.5.xlsx`；成品 `dist/20261115_智譜個股模型_v0_1_5.html／.xlsx`、版面 v1.3 個股頁 `dist/20261115_智譜個股頁_v0_1_5.html`；見下方 v0.1.5 節）。前版 v0.1.4（`model/archive/20261106_Zhipu_v0.1.4.xlsx`；成品 `dist/_待刪除/20261106_智譜個股模型_v0_1_4.html／.xlsx`）＝只換現價：2026-11-05 收盤 HK$476（SRC_ZP_1055），只影響 Reverse 反向對照與 Funding 可換股債券「現價 ÷ 換股價」（報告 `docs/reports/20261106_v0.1.4.md`）。再前版 v0.1.3＝文件與核對小修，依 `docs/plan/個股模型原則_命題加營收分部.md` 第 1 節改稱「個股模型」（報告 `docs/reports/20261104_v0.1.3.md`）。
- 官方報告分部（部署方式）：本地化部署（②企業級智能體＋③企業級通用大模型＋④技術服務及其他）、雲端部署（①開放平台及 API）。宣告檔 `segments.json`；對帳在 Revenue 第七節（FY2025 對年報、1H2026 對中期公告子項加總，容差 Inputs `seg_recon_tol`），Checks C100–C103；`python3 tools/segment_check/check_segments.py zhipu`（repo 根目錄執行）須為「符合」。
- 規格：`docs/plan/20261008_Zhipu_v0.1_規格.md`
- 交接：`docs/handoff/Zhipu_handoff.md`；進度：`docs/reports/20261008_zhipu_進度.md`（v0.1.2：`docs/reports/20261030_v0.1.2_進度.md`）
- 工作規範：`CLAUDE.md`


## v0.1.5（2026-11-15）：個股 HTML 版面 v1.3＋三層協調（讀取層級、C0 公司實績）

**依據：** `docs/workorders/20261114_第二批13家_v1.3與三層協調.md`；`docs/plan/三層基準協調規則.md` v1.4.2 C0 第 2 點；`docs/plan/個股HTML版面規格.md` v1.3；`tools/coord/README.md`。樣板：Anthropic v0.2.2、OpenAI v0.6.5（PR #524）。

**做法：** `builder/z5.py` 新增 PL（標準損益表）、Coord（讀取層級、C0、判定摘要、參考）兩頁，只列示；`tools/build_html_v13.py` 產生版面 v1.3 個股頁（命題 → 估值〔不適用〕→ 敏感度〔各情境對 FY2030 覆蓋率〕→ 損益 → 驅動因子〔分部營收推導〕→ 算力 → 資金 → 與 OpenAI 並排 → 對帳〔C0 表〕→ 資料狀態），數字與一頁摘要同一套取值（LibreOffice 重算、情境重算）與 engine 比對。
- **估值：本模型未建目標價與評等**（Reverse 只作反向對照：市值隱含營收）。智譜已上市、有市價，但本模型是收支模型，未建 DCF／同業估值；依規格 J 節 (b)（2026-11-15 擴充：已上市但本模型未建目標價），估值章標「估值：不適用（本模型未建目標價）」並帶 `data-novaluation="1"`，列現價、市值、EV 與反向檢驗（市值隱含營收），附一張反向檢驗圖（模型營收 vs 市值隱含年營收與分析師共識；數字取自 xlsx）。分析師共識只到 FY2028，FY2029–FY2030 為沿用值（FY2028 共識平均 ×（1＋反向模式年增率假設），基準年增率 0，故三年同值 324.7）。若要給目標價需另建估值頁（判斷類，待 Andy 決定）。

**實績與出處（C0）：**
| 項目 | 期間 | 公司實績 | 出處 | 層級隱含 | 判定 |
|---|---|---|---|---|---|
| 營收總額 | 2026-01～06（主）；季營收未揭露（第二讀數無） | 9.5389 億元 | 2026 中期業績公告（2026-08-31，港交所 2026083101539），SRC_ZP_177 | 無（Wattonomics 無智譜玩家） | 無可比實績（層級無此玩家） |
| 資本開支 | 1H2026 | 3.906 億元（主要為收購北京紅鑽） | 同上，SRC_ZP_209 | 無 | 無可比實績（層級無此玩家） |
| 在役／簽約算力 GW | — | 公司未揭露（只有國產晶片十萬張級、1 GW 資料中心等說法） | SRC_ZP_452、458（Interested-party） | 無 | 無可比實績（公司未揭露、層級無此玩家） |
| 應付＋預付算力服務費 | 2026-06-30 | 26.71 億元（14.27＋12.44） | 中期公告資產負債表，SRC_ZP_218、219 | 無 | 無可比實績（層級無此玩家） |

**判定結果：** 可判定 0、對帳中 0、無可比實績 4；無同源列（Wattonomics 快照全文無智譜，層級不讀本模型）。讀取層級：Tokenomics 快照 v19.0（現行 v22.0，本輪不換）。

**驗證：** builder 重建（restore matched 3,585、unmatched 0；e6 0）；`tools/xlsx_diff.py` 對 v0.1.4：只差 README A1、Checks A1 版本字串與新增格（README「v0.1.5 變動」列、PL、Coord）；test_builder 14 通過；parity（結構、具名範圍、base 與兩個預期 ERR 情境）通過，完整情境以 CI 為準；test_html（含個股頁數字比對）；`check_layout_v13 --xlsx`：本分支版（J 節 (b) `data-novaluation` 豁免）全過；`tools/baseline_coordination/check.py` 通過。

**修正（同版 v0.1.5 重建，2026-11-15）：**
- PL 核對列「營業利益（預測）− Cost 差額 − 股權報酬」原為 `ROUND(…, 9)`；pycel 的 ROUND 在 1 GW 資料中心情境（`w_dc1gw_on`、`aq_owned_opex_amort`）遇 numpy 浮點數無法轉 Decimal（InvalidOperation），parity 失敗 6 項。改為 `IF(ABS(x)<1E-9, 0, x)`（同義：誤差小於 1E-9 視為 0）；其他核對列（稅前 − 自由現金流、FY2025 校準）本來就不含 ROUND，全簿無其他 ROUND。數值逐格不變（xlsx_diff 對修正前只差 Coord!L15 說明文字，見下）。
- 規格 J 節擴為「無目標價的模型」：(a) 未上市 `data-unlisted`（OpenAI）；(b) 已上市但本模型未建目標價 `data-novaluation`（智譜、MiniMax）。`tools/company_html/check_layout_v13.py` 兩旗標同組豁免（不要求反向 DCF 章、略過 G、H），(b) 估值章仍須至少一張圖。智譜個股頁估值章改帶 `data-novaluation`、補反向檢驗圖、橫幅改「本模型未建目標價與評等」。
- Coord「Wattonomics」列說明補：Wattonomics Decisions D43（GL01-37，Andy 2026-10-31 已定）已將智譜移出 Gap（資金缺口收支只算 11 家代表玩家，不含中國）。
- 驗證：完整 parity（tests/parity，47 情境）全過——情境一致性 47／47、快取等價 47／47、結構與效能測試通過；`test_scenarios_actually_change_outputs`、`test_scenario_expected_results` 在本機單一行程記憶體不足（7 GB 上限，pycel 引擎未即時回收），改以同一斷言分段執行（每情境後回收記憶體）全過。test_builder、test_html 通過；`tools/baseline_coordination/check.py` 0 項不符。
- 檢查：main 版 `check_layout_v13` 對本頁不過（main 版只認 `data-unlisted`；本分支合併後即為新版）；新版全過。repo 內其他既有 HTML（77 份）main 版與新版輸出逐字相同。
