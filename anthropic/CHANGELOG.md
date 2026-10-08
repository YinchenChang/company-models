# CHANGELOG（Anthropic 收支模型）

每次同步記錄：Excel 版本、段別、commit 與變動摘要。

## A0 骨架（2026-10-08，Excel 尚無）
- 新增 `anthropic/`：CLAUDE.md、規格 r0（D1–D20）、共同規則、A1–A5 工作單、進度檔、交接檔；engine、tools、tests 基礎檔沿用 OpenAI v0.6（`claude/openai-s6-release` @ `e91df57`），待 A2 改寫。

## v0.1-A2 需求與營收（2026-10-08；Excel `model/20261008_Anthropic_v0.1.xlsx`）
- commit：`aa06cb2`（步驟 1 骨架）、`52d8dd5`（2 NonNV）、`7609b7c`（3 Inputs）、`691b46f`（4–6 Demand／Revenue／Checks）、`d9c6284`（7 測試與 CI）、`fd5db5b`（每任務 token 改取 Tokenomics）；報告提交見 PR #22 留言。
- 新建活頁簿 9 頁：README、SRC_ANT（379 列；新增 377–379；Derived 12 列：11 公式、1 保留報導值）、TK_Link（Tokenomics v5.26：63＋5 名＋NonNV 18 格）、OAI_Link（OpenAI v0.6 命題輸出 17 名）、Inputs（105 列；INP_067–069 退役）、Demand、Revenue、Checks（C01–C48）、_Defaults。
- 結果：2025 營收 4.60（校準差 0）；2026 48.73（半校準）；2030 總額 214.9、雲端平台抽成 18.1、淨額 196.8（$B）。
- 工具：`builder/build.py`（分段模組 a2／a3／a4）、`builder/a2.py`、`builder/oai_link.py`；`tools/check_tk_snapshot.py` 支援 NonNV 讀表；`tools/make_report_tables.py` 改寫；`tools/export_csv.py` 改 SRC_ANT；`engine/core.py` 檔名規則改 Anthropic。
- 測試：parity 19 情境、test_builder（E6、穩定 ID、OAI_Link 隔離、Derived 公式化）；54 項全過。CI：`.github/workflows/anthropic-parity.yml`（GitHub 帳號付款／用量上限，job 未啟動）。

## v0.1-A3 算力與成本（2026-10-08；Excel `model/20261008_Anthropic_v0.1.xlsx`，同檔名覆寫）
- 新增 `builder/a3.py`：Compute（加速器族 × 世代 9 族、逐合約供給 GW 10 份＋既有雲端＋自建、η、推論／研發 GW、容量上限、VR 等值、敏感度 A／B／C、對照列與機房租約）、Cost（逐合約實付、自建資本支出、供應商持有成本與雲端毛利、經濟口徑、非算力成本與股權報酬、推論／研發拆分、命題表、TK 單位成本參考、D22 對照、$518B 對帳）。
- Revenue `REV_CapFactor` 改接 Compute `CMP_CapFactor`（2025、2026 固定 1）；Inputs 佔位 INP_016 退役。Inputs 160 列（新增 INP_109–164）。TK_Link 加讀 Tokenomics Inputs 頁 PUE（`TK_PUE`、`_Lo`、`_Hi`；讀表，非具名）；`tools/check_tk_snapshot.py` 通用化讀表比對（OK 89）。
- Checks C49–C82；CHK_Errors＝0、CHK_Warnings＝0。parity 情境 27 個（A3 新增 10 個：合約價、η、產出比、人數成長、自建資本支出、PUE、Broadcom 5 GW＋爬坡、容量截頂、兩個 ERR 情境；移除 A2 的容量上限佔位情境 2 個）。
- 結果（$B；每 VR 等值 GW 為 $B/GW/年）：2030 供給 10.4 GW（VR 等值 7.7）、算力成本 90.8、全成本 110.3、營收淨額 196.8；每 VR 等值 GW 差額 2026 −6.5、2027 −5.2、2028 −2.9、2029 ＋5.1、2030 ＋11.2；覆蓋率 2030 1.79（OpenAI v0.6：−10.1、0.46）。

## v0.1 融資與反向（A4，2026-10-08；Excel `model/20261008_Anthropic_v0.1.xlsx`，同檔名覆寫）＝v0.1 模型完成
- commit：`07565be`（步驟 1–3）；完成報告與對照 Excel 的提交見 PR #22 留言。
- 新增 `builder/a4.py`：Funding（自由現金流、來源順序 D16 r1、已到位融資 $100B、條件式與 IPO、最低現金 D17、外部資金需求、情境 S1–S7、或有負債、回流對照 D18）、Reverse（管理層目標 → 倍數與反向資金，D19）；Cost 第十二–十五節（逐家合約總額對招股書、錨定招股書敏感度、機房租金敏感度、D22 量化）；`COST_PropGap_GW` 具名範圍。
- Inputs 175 列（新增 INP_165–179）。Checks C83–C105；CHK_Errors＝0、CHK_Warnings＝1（C89 逐家合約 WARN，AMD ×3.9，chat 端要求）。parity 情境 34 個（A4 新增 7）；`test_reverse_not_fed_back`、`test_named_outputs_v01`；基準 WARN 數改由 scenarios.yaml `base_warnings` 設定。
- 工具：`tools/sensitivity.py`（engine 重算 14 個驅動的區間端點）；`tools/make_report_tables.py --stage A4`（④A1–A4 預設彙總、⑤敏感度）；`builder/a2.py` 列鍵記號支援 F、X；`build.py` 加 Reverse 隔離掃描。
- 結果：累計外部資金需求 2026–2030＝0（無缺口年；只靠已到位融資的現金谷底 2028 年 $80.2B）；2030 年底現金 218.4；每 VR 等值 GW 差額不變（−6.5／−5.2／−2.9／＋5.1／＋11.2）。OpenAI v0.6：累計 442.8。反向（管理層舊目標、同一支出）累計 25.2。
- Tokenomics：CURRENT 仍為 v5.26（master 前進到 `70d859e`，只新增文件）；check_tk_snapshot OK 89。

## v0.1 成品（A5，2026-10-08；Excel `model/20261008_Anthropic_v0.1.xlsx` 未改動）
- commit：`5d08b70`（產生器、測試、dist、截圖、CI）、`210faf0`（命題與驅動對照表）、`82e88e7`（交接檔、根目錄 README）；成品報告提交見 PR #22 完成留言。
- `tools/build_html.py` 改寫為 Anthropic 版（9 節：結論、每 VR 等值 GW、營收結構與 run-rate 里程碑、算力與逐家合約對帳、現金與融資（含 S1–S7、回流對照）、三個關鍵驅動與「要成立必須」、與 OpenAI v0.6 並排、Andy 待審 5 項、資料與版本）；取位支援「頁!列 ID:欄」；`tools/html_scenarios.yaml` 20 個情境（含 TK NonNV 區間）；`tools/screenshot_html.py` 預設檔名改 v0.1。
- 成品：`dist/20261008_Anthropic收支模型_v0_1.html`（818 個可追溯數字、無外部資源）、`dist/20261008_Anthropic收支模型_v0_1.xlsx`（＝現行模型）；截圖 `docs/reports/img/20261008_v0.1_成品_*.png`。
- 測試：新增 `tests/test_html.py`（HTML＝engine 重算、無外部引用、dist xlsx＝現行模型、文字主張、單一驅動不產生外部資金需求）；`.github/workflows/anthropic-parity.yml` governance 加 HTML 測試一步。
- 文件：`docs/plan/20261008_Anthropic_命題與驅動對照表.md`；交接檔更新為 v0.1 完成；報告 `docs/reports/20261008_v0.1_成品.md`。
