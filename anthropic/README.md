# Anthropic 收支模型

Anthropic 每 VR 等值 GW 的年營收能否覆蓋每 GW 年全成本；若不能，缺口由誰、以什麼條件融資（FY2025–FY2030）。架構比照 `openai/`（v0.6），算力物理取自 Tokenomics。

- 規格：`docs/plan/20261008_Anthropic_v0.1_規格.md`
- 交接：`docs/handoff/Anthropic_handoff.md`；進度：`docs/reports/20261008_anthropic_進度.md`
- 工作規範：`CLAUDE.md`

## v0.2.4：現價與同業報價改 2026-11-20 收盤（工作單 px1120）

- 現價與同業報價改 2026-11-20 收盤（工作單 px1120）。現價 74（IPO 區間中點）→ 88.91（ANT 2026-11-20 收盤；IPO 價 83，11-19 首日收盤 97.17）；同業 EV／遠期營收中位數 26.32 → 26.32（不變：只有 DDOG 有 11-20 收盤，其倍數 16.36 → 10.94，不是中位數那一家）；加權目標價 82.71 → 83.66；上檔空間 ＋11.8% → −5.9%；評等 中立 → 中立（模型公式重算，未手改）。
- 價格來源只用 `markets/invest_map/data/closes_20261109_20261120.csv`（ANT、DDOG 11-20）與 `closes_20261020_20261113.csv`（DDOG 10-30＝235.27）。DDOG 倍數＝16.36 × 157.37 ÷ 235.27；EV 沒有市值可用，以倍數 × 價格比近似（SRC_ANT_622–624，SRC_ANT_590 改為公式）。
- **同版補換（第二輪，版號不變）：** PLTR、CRWD、NET、SNOW 取 `closes_extra_20261120.csv`，EV／遠期營收由資料日 2026-11-06 換算到 11-20 收盤（倍數 × 11-20 收盤 ÷ 11-06 收盤；EV 無市值可用，以倍數 × 價格比近似，標 Derived；新增 SRC_ANT_625–636）。五家同業皆已換價，無未換價同業。同業 EV／遠期營收中位數 26.32 → 26.93；加權目標價 83.66 → 84.80；上檔空間 −5.9% → −4.6%；評等中立不變。Sens、Val 情境快照已重跑。
- 現價 `ipo_px` 同時是淨募資、私募股數、選擇權庫藏股的連動輸入，照既有公式隨現價重算（淨募資 80.8 → 91.8）；實際 IPO 價 83 未入模型；Sens 的 70／78 仍是招股說明書區間。共識、rf、β、ERP、營運假設、情境權重不變；Sens 與 Val 情境快照已重跑。
- Coord 維持讀 Wattonomics v0.3.5 快照（`CURRENT_SNAPSHOT` 已到 v0.3.9，不在本版範圍）。
- 重建：`python3 builder/build.py --tk-dir <Tokenomics@eded8e6> --oai-xlsx ../openai/model/archive/20261029_OpenAI_v0.6.4.xlsx --out model/20261121_Anthropic_v0.2.4.xlsx --base model/archive/20261115_Anthropic_v0.2.3.xlsx --date 2026-11-21`；再 `tools/val_sens.py`、`tools/val_snapshots.py` 後重建；`tools/build_html.py`、`tools/build_html_v2.py --html … --xlsx … --out dist/20261121_Anthropic個股模型_v0_2_4.html`。

## v0.2.3：Coord 改讀 Wattonomics v0.3.5（營業現金流對照差 0）

**依據：** Wattonomics v0.3.5（#618，main `2debcc7`）②⑥：Anthropic 在 Wattonomics 的營業現金流不加回自建資本支出（`gocf_ant`＝本模型 Funding F07；F03 恆為 0），且 Y9 重取改讀本模型 v0.2.2 dist（SRC_GL_320，commit `90ab689`）。建置：`python3 builder/build.py --tk-dir <Tokenomics checkout@eded8e6> --oai-xlsx ../openai/model/archive/20261029_OpenAI_v0.6.4.xlsx --out model/20261115_Anthropic_v0.2.3.xlsx --base model/archive/20261114_Anthropic_v0.2.2.xlsx --date 2026-11-15`（OAI_Link 仍為 OpenAI v0.6.4，與 v0.2.2 相同；Coord 頁隨 `../wattonomics/data/CURRENT_SNAPSHOT`＝`wattonomics_snapshot_v0.3.5.json`）。

- Coord B 節「營業現金流 2026–2030」：本模型 `FND_FCF`（Funding F07）10.32／−37.39／−42.21／4.64／56.26 對 Wattonomics `WIF_Gap_Player_ant_<年>_OCF` 同值，差 0（殘差 < 1e-6，來自快照取 6 位小數）。v0.2.2 時 Wattonomics 匯入停在 v0.1.8（7.03／−42.29／−49.56／−5.63／42.83），差 3.3 至 13.4（層級匯入落後），本版已消除；說明文字改由快照 `gocf_ant` 出處自動帶出。
- 自建 IT GW 年新增對照（Inputs vs `WIF_Gap_Player_ant_<年>_CapexAIGW`）改寫為對照 v0.3.5 快照，差 0。追溯重跑 `tools/coord/trace_wattonomics.py --player ant`（v0.3.5 xlsx）：自建 GW＝0、CapexAI 取 GW 法＝0、CapexAIModel＝0，結論不變。C 節 T1 配置（AWS 1.037、Google 1.0、2025 年底 0.626）與 v0.3.2 相同，C0 判定（營收「時差邊界（參考）」、對帳中 0）不變。
- 其餘工作表數值與 v0.2.2 逐格相同（`tools/xlsx_diff.py`：差異只在 README 列位移、TK_Link／OAI_Link 讀取日與建置提交、Coord、Checks 標題）；加權目標價 82.71、評等中立不變。

## v0.2.2 三層協調（Coord 頁）：模型商樣板（OpenAI 複製用）

**依據：** 三層基準協調規則 v1.4.2 C0 第 2 點（Andy 2026-11-14 17:40：「有意義的對帳要拿公司實績去對層級值」）；工作單 `docs/workorders/20261114_第二批13家_v1.3與三層協調.md`、`20261114_C0實績對帳補列_七家.md`。同輪完成版面 v1.3 收尾（檢查 G、H 位置規定，見 `tools/company_html/fundmodel_v13.py`）。版號：#469 的 v0.2.1 未合併，本 PR 直接收為 v0.2.2（v0.2.1 檔移 `model/archive/`、`dist/_待刪除/`）。

**Coord 頁結構（builder/a13.py；只列示，不被任何計算頁引用，Checks 除外；既有格逐格不變）：**
1. A 參數：差距門檻 ±15%、季化倍數 4（藍字）；月數取 Inputs 定義常數。
2. B 讀取層級（同源，不判定）：本模型實際讀的層級與版本——Tokenomics 快照 v21.10（Tokenomics master 已是 v22.0，本輪不換快照）；Wattonomics 只在 v0.1.7 把自建 GW 與廠房提前月數寫進 Inputs（逐年與 v0.3.2 快照對照，差 0）；Wattonomics Gap 頁匯入的 Anthropic 營業現金流停在 v0.1.8，與本模型現值差 3 至 13（層級匯入落後，非證據分歧，待 Wattonomics 重匯）。另有「追溯」列，寫明 Wattonomics 的 Anthropic GW 與資本支出路徑公式讀什麼。
3. C 層級隱含值：Wattonomics T1（營運者→使用者）配置給 Anthropic 的年底 GW（AWS 1.037、Google 1.0；2025 年底 0.626）取 2026 年平均 1.33 GW × Tokenomics L1_RevGW_PQ_VR200（49.5，低 36.2、高 78.5 十億美元／GW／年）＝CY2026 隱含營收 65.9（48.2–104.5）。
4. D C0（實績）：
   - 營收：S-1/A 2026 前九月 33.242（SRC_ANT_547）、Q3 16.956（SRC_ANT_488）。主讀數＝33.242 對 65.9 × 9/12＝49.4，差距 −32.8%；第二讀數＝Q3 × 4＝67.8 對 65.9，差距 ＋2.9%。只有主讀數超過 ±15%，判「時差邊界（參考）」，不開對帳單。
   - 自建資本支出：S-1/A 前九月購置不動產及設備 0.925（SRC_ANT_519）vs 層級自建 AI 資本支出 0（Wattonomics 認定無自建 GW）：口徑不同（辦公與小型設施 vs 自建 AI 算力），層級值為 0 無百分比，列示、不計入。
   - 算力承諾：S-1/A 以美元揭露、Wattonomics 以 GW：無可比實績（缺層級每 GW 合約價）。
5. E 層級之間（參考）：Wattonomics 內部 usr_ant 報導 GW（2026 年底 5，NYT 轉述的目標）對 T1 配置（2.037）；不計入 C0。

**設計決定（OpenAI 複製時照做，數字自行換）：**
- 層級側只用「不讀本模型輸出」的值：Tokenomics 每 GW 營收（以 OpenAI 營收與牌價校準，非 Anthropic 值）、Wattonomics T1 配置 GW。追溯 Wattonomics 公式後確認：T1 是輸入列（只進 Matrix）、玩家路徑自建＝0、Gap 頁只匯入本模型現金流（影響資金缺口，不影響 GW）；Google 1 GW 為公司公告的規劃容量（公司揭露，非本模型輸出）。OpenAI 複製時須重新追溯（OpenAI 在 Wattonomics 有玩家 oai 與 usr_oai，查是否有公司目標寫入 GW）。
- 本模型自層級讀入的值（TK_Link、Inputs 的 Wattonomics 值）只放 B 節，差距恆為 0 或屬匯入落後，不判定、不計入對帳中。
- 偏誤方向依期間中點說明（主讀數期間中點 5 月中，早於全年中點；Q3 × 4 是 8 月中的年化速率，晚於全年中點），不寫成「夾住真值」。
- 已知限制與敏感度（Coord D2 節、HTML 對帳章同表，不計入判定）：同一組實績只換層級側取法，判定都翻成「對帳中」：
  - 每 GW 營收 ÷ 2.0（Tokenomics 自註「誤差全在產能側時實際約為該值 ÷ 2.0–2.1」，OpenAI 2026 口徑）：主讀數 ＋34.5%、Q3×4 ＋105.8%。
  - 機隊以 Trainium／TPU 為主：每 GW 營收是 VR200 口徑，用 Tokenomics NonNV 產出比（Trainium3 0.28、TPU v7 0.66）換算，AWS 配置全當 Trainium3、Google 全當 TPU v7（下界估計），年平均約 0.56 GW 等值（C 節），兩讀數約 ＋59%、＋143%。
  - GW 改取 Wattonomics 報導值 usr_ant（2026 年底 5 GW，NYT 轉述目標）：兩讀數 −72%、−57%。
  判定對 GW 取法與每 GW 營收口徑敏感，所以營收列只作量級檢驗，維持「時差邊界（參考）」，不據以開對帳單。
- 實績為 GAAP 總額（含雲端市集銷售，未扣平台抽成），層級每 GW 營收的 m 以 OpenAI 營收口徑校準，兩者口徑不同，差距含口徑差（Coord 營收列說明已寫）。
- Coord 的 T1 配置 GW 由 builder 依鍵（al_amzn_ant、al_googl_ant、al_neo_ant；使用者＝ant、形式＝GW）從 wattonomics 快照 JSON 讀入，並以斷言確認鍵與欄位。
- HTML 對帳章新增「三層協調」表（`fundmodel_v13.py` 的 `coord_v13`），數字全由 xlsx 具名範圍 COORD_C0、COORD_C0Sum、COORD_C0Verdict 取值。
- **OpenAI 複製注意（同源）：** Tokenomics 的 m 由 OpenAI 營收 ÷（付費 token × 牌價）校準，OpenAI 的營收 C0 與層級每 GW 營收同源，OpenAI 營收列應標「參考（同源校準）」、結論「無可比實績（層級值同源）」，改比不同源的量（例如資本支出、在役 MW、算力承諾）。
- **建置前提：** 本模型建置須 Tokenomics checkout 在 commit `eded8e6`（v21.10；`--tk-dir` 指向該 checkout，`model/CURRENT` 為 `20261104_Tokenomics_v21.10.xlsx`），並以 `--base` 指向現行 xlsx；CI 重建只作比對，須固定此 ref，不可取 Tokenomics master（現為 v22.0，TK_Link 值會變）。
- 不做兩讀數目標價（模型商，非硬體股）。

**驗證：** 見 CHANGELOG v0.2.2。xlsx 逐格比對 v0.2.1：既有計算頁 0 差異（加權目標價 82.71、評等中立不變）；新增 Coord 頁與 Checks 2 列；README 列位移。
