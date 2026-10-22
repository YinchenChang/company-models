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

## v0.1.1 草擬 S-1 與 Q3 簡報更新（2026-10-21；Excel `model/20261021_Anthropic_v0.1.1.xlsx`；v0.1 移至 `model/archive/`）
- 依據：工作單 `docs/workorders/20261021_anthropic_v0.1.1_S1草稿更新.md`（patch 版，非新階段）。commit 見 PR 留言。
- 資料：SRC_ANT 409 列（新增 380–409，30 列；WSJ 草擬 S-1 轉述與 Q3 簡報，Interested-party）；`not_found` 新增 4 項（2026-06-30 現金、WSJ 逐字措辭、SemiAnalysis 第二來源、付款排程與股權報酬）。Inputs 182 列（新增 INP_180–186：u1_switch、q4_growth、u2_switch、spend_warn_threshold、lambda_end、neg_start_year、neg_term_years）。
- U1 2026 半年校準 D15 r2（H1 實際＋Q3 實際＋Q4＝Q3）：2026 營收 48.73→49.90。U2 Nscale＋Lambda 錨定 86（MW 比例分攤）。U3 逐家改對照 S-1、總額三口徑。U4 2026 算力支出對照（C116 WARN）。U5 D22 加 Q3 11%。U6 談判中 $150B 情境（Cost 第二十節、Funding S8）。U7 非現金 72 列示。U8 資料缺口。開關 INP_180／INP_182＝0 時輸出回到 v0.1。
- 結果：差額 2026–2030 −5.4／−4.9／−2.6／＋5.2／＋11.2；覆蓋率 2030 1.80；累計外部資金需求 0；現金谷底 84.2。CHK_Errors 0、CHK_Warnings 2（C89、C116）。
- 檔案：新增列皆附加在各頁尾，既有列編號不變。`engine/core.py` 檔名規則允許三段版號。`tools/build_html.py` 新增「①之二 v0.1.1 更新了什麼」，`tools/html_scenarios.yaml` 新增 3 情境（v010、only_u1、only_u2）。parity 新增 6 情境。報告 `docs/reports/20261021_v0.1.1.md` 與 `_對照.xlsx`。

## v0.1.2 Tokenomics 快照 v5.26 → v18.0（2026-10-21；Excel `model/20261021_Anthropic_v0.1.2.xlsx`；v0.1.1 移至 `model/archive/`）
- Tokenomics：master `868e2cc`，`20261021_Tokenomics_v18.0.xlsx`（SHA-256 `e654876e…`）；TK_Link 重讀同一份 68 名＋NonNV 18＋PUE 3，新增 14 個 v18.0 對照名稱（`L1_RevGW_PQ_VR200` 系列、`L1_TokGW_PaidPQ_VR200` 系列、`L1_TokLedger_2026Q2`、`L1_M`；只對照）。check_tk_snapshot：OK 103。
- 公式未改。η 0.888→0.527；供給 VR 等值 GW 2030 8.044→8.067；每 VR 等值 GW 差額 2026–2030 −5.446／−4.921／−2.648／＋5.185／＋11.172 → −5.365／−4.888／−2.633／＋5.166／＋11.139；覆蓋率 2030 1.804、累計外部資金需求 0、現金谷底 84.23 不變。
- 新增：Cost 第二十一節 K240–K255（TK P × Q 對照）、Checks C124–C126（info）、具名 `COST_TKPQ`／`COST_TKPQRatio`／`COST_TKTokLedger`；TK_Link C3 記錄 TK 檔案 SHA-256。說明文字更正 5 處（見報告③）。
- 測試：parity 情境 +2（ha、hb）；HTML 情境 +3（v011、only_tokgw、only_hold），v010／only_u1／only_u2 改為同時把 TK 回到 v5.26（spec `prev`）；期望值 formula_cells 5878、defined_names 1146、tk_values 82。
- 成品 `dist/20261021_Anthropic收支模型_v0_1_2.html／.xlsx`；v0.1.1 成品移至 `dist/_待刪除/`。報告 `docs/reports/20261021_v0.1.2.md`＋`_對照.xlsx`。

## v0.1.3 A 段（2026-10-22；Excel `model/20261022_Anthropic_v0.1.3.xlsx`；分支 `claude/anthropic-v0.1.3`，PR #69）
- 採納外部驗收報告（v0.1.2）：TK 快照改 v19.0（103 個值與 v18.0 相同）；U2 新法（chat 端更正公式：Nscale、Lambda 供給 GW 取揭露容量，付款路徑不變＝『最高 86』按 MW 比例的攤提；隱含每 GW 年付款 Nscale 19.3／Lambda 14.2／合計 17.1；Inputs `u2_method`＝INP_187）；資金三層（Funding 第十節）；持續擴張情境（Cost 第二十二、二十三節、Funding 第十一節；Inputs INP_188–194）；HTML 結論改條件式並新增「①之三」。
- 結果：每 VR 等值 GW 差額 2026–2030 −5.37／−4.91／−2.70／＋5.38／＋11.60（v0.1.2：−5.37／−4.89／−2.63／＋5.17／＋11.14）；覆蓋率 2030 1.80（不變）；營運自由現金流首次轉正 2029；三層累計外部資金需求 41.9／0／0；現金谷底 84.2（不變）。
- 公式格 6,543、具名範圍 1,243（A 段原記 6,550／1,240，B 段以 engine 實測更正）、Inputs 190、Checks 至 C142；CHK_Errors＝0、CHK_Warnings＝2。報告 `docs/reports/20261022_v0.1.3-A.md`。


## v0.1.3 B 段（2026-10-22；Excel `model/20261022_Anthropic_v0.1.3.xlsx` 同檔名重建；分支 `claude/anthropic-v0.1.3`，PR #69）
- 新增 `builder/a5.py`：**Bridge 頁**（Cost 之後；109 列；需求 → token → GW → 成本逐層；每列來源編號、2025–2030、標記；低／高＝2030 敏感度包絡或 2025 Inputs 區間；末節與 TK 對照）與 **EtaDiag 頁**（37 列；η（2025）＝計價口徑 f_P × 利用率 f_U × 工作負載組合 f_W × 未解釋殘差；並列 TK L1_Gap*）。兩頁唯讀，不被計算頁引用（Checks e1_iso；`test_bridge_not_fed_back`）；η 的值不變。
- TK_Link 尾端 +23 名（IF_UtilActual 含低高、L1_TokGW_PaidCap_VR200、B4_MixPaid、L1_Gap{Prompt,SpendBasis,ISL,Util,Product,Unexplained} 各含低高），只供 EtaDiag；`check_tk_snapshot`：OK 126。Checks C143–C152（5 個 OK 類、5 個 info）；具名範圍 +29（TK 23、ETA_ 6）。
- 既有頁（Inputs、Demand、Revenue、Funding、Reverse）對 v0.1.2 逐格比對 0 差異；Compute、Cost 的差異全為 U2 新法鏈（A 段）；`u2_method`＝0 時與 v0.1.2 逐格相同（僅 5 格說明文字）。公式格 6,543 → 7,231；具名範圍 1,243 → 1,272；Inputs 190（不變）。
- `anthropic/CLAUDE.md` 新增第 8 節：公開 S-1／經審計資料出現時的替換程序。
- 成品：`tools/build_html.py`「①之三」改名「v0.1.3 更新了什麼」並加 η 拆解、Bridge 逐層對照（HTML 1,134 個數字）；`tools/read_html_screen.py`（Playwright 逐節讀畫面並截圖，桌機 1280／手機 390，無水平捲動、無外部請求、畫面文字與 data-v 0 不符）；`dist/20261022_Anthropic收支模型_v0_1_3.html／.xlsx`。
- 報告 `docs/reports/20261022_v0.1.3.md`＋`_對照.xlsx`（`tools/make_report_v013b.py`）。

## v0.1.4 Tokenomics v19.1 同步、η 診斷補充（2026-10-22；Excel `model/20261022_Anthropic_v0.1.4.xlsx`；v0.1.3 移至 `model/archive/`；分支 `claude/anthropic-v0.1.3`，PR #69）
- 工作單 `docs/workorders/20261022_anthropic_v0.1.4_TK19.1.md`。commit 見 git log「Anthropic v0.1.4 步驟 1–6」。
- TK 快照 v19.0 → v19.1（master `c2ecb71`）：39 個名稱值變動（check_tk_snapshot 重建後 OK 126、DIFF 0）；僅 `IF_TokGW_Luna` 進計算（Compute）；其餘為參考列、診斷或僅快照。命題輸出與 v0.1.3 逐位相同；η 0.52730 → 0.52543，有效推論 GW 2030 ＋0.0043、研發 GW −0.0043。
- 新列（皆在尾端）：TK_Link ＋`IF_ProdDerate`；EtaDiag 第七節 E38–E47（`ETA_fU_Prod`、`ETA_Prod2`、`ETA_Resid2`）；Checks C153–C156。具名範圍 1,272 → 1,276、公式格 7,231 → 7,256。η 值不因診斷改變。
- 刪除 `dist/_待刪除/` 與 v0.1.3 dist 檔；HTML「①之三」前端加 v0.1.4 摘要、逐群拆解表、η 各版本小表；`tools/html_scenarios.yaml` 加 `tk19_*` 群組與 `v013`、`v14_only_luna`、`v14_only_rest` 情境（舊情境加 `tk19_back` 以保持當時數字）。
- 測試：`tests/parity/scenarios.yaml` 預期值（formula_cells 7,256、defined_names 1,276、tk_values 105）；`test_builder` 新增 `test_eta_prod_derate_section`，η 預期值改 0.525427732668394。報告 `docs/reports/20261022_v0.1.4.md`；工具 `tools/make_report_v014.py`。
