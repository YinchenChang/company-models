# Anthropic 收支模型

Anthropic 每 VR 等值 GW 的年營收能否覆蓋每 GW 年全成本；若不能，缺口由誰、以什麼條件融資（FY2025–FY2030）。架構比照 `openai/`（v0.6），算力物理取自 Tokenomics。

- 規格：`docs/plan/20261008_Anthropic_v0.1_規格.md`
- 交接：`docs/handoff/Anthropic_handoff.md`；進度：`docs/reports/20261008_anthropic_進度.md`
- 工作規範：`CLAUDE.md`

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
- 已知限制（可見說明已寫）：Tokenomics 自註每 GW 營收若誤差全在產能側，實際約為該值 ÷ 2.0–2.1（OpenAI 口徑），此時第二讀數會轉為明顯正差距，判定會變；故營收列是量級檢驗，不是精確對帳。
- HTML 對帳章新增「三層協調」表（`fundmodel_v13.py` 的 `coord_v13`），數字全由 xlsx 具名範圍 COORD_C0、COORD_C0Sum、COORD_C0Verdict 取值。
- 不做兩讀數目標價（模型商，非硬體股）。

**驗證：** 見 CHANGELOG v0.2.2。xlsx 逐格比對 v0.2.1：既有計算頁 0 差異（加權目標價 82.71、評等中立不變）；新增 Coord 頁與 Checks 2 列；README 列位移。
