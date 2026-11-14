# OpenAI 收支模型（company-models/openai）

OpenAI 收支模型（v0.6，建置中）。2026-10-07 由 `YinchenChang/openai-model`（`main@13c16f0`，停用）搬入本資料夾；以下指令都在 `openai/` 內執行。交接檔：`docs/handoff/OpenAI_handoff.md`；進度：`docs/reports/20261007_openai_進度.md`。命題：OpenAI 每 VR 等值 GW 的營收能否覆蓋每 GW 全成本；若不能，缺口由誰、以什麼條件融資。FY2025–FY2030，曆年制。

- Excel 為唯一計算引擎：`model/CURRENT` 指向現行檔。
- 工作規範見 `CLAUDE.md`；工作單見 `docs/workorders/`；每包報告見 `docs/reports/`；版本紀錄見 `CHANGELOG.md`。
- 本地建置：`pip install -r requirements-dev.txt`；`python3 builder/build.py --tk-dir <Tokenomics> --out model/<檔名>.xlsx --base model/<現行>.xlsx`；測試 `python3 -m pytest tests/parity`。
- 需要 LibreOffice Calc（Ubuntu：`apt-get install -y --no-install-recommends libreoffice-calc-nogui`）。

## v0.6.5 三層協調（Coord 頁）：模型商（複製自 Anthropic v0.2.2）

**依據：** 三層基準協調規則 v1.4.2 C0 第 2 點（Andy 2026-11-14 17:40）；工作單 `docs/workorders/20261114_第二批13家_v1.3與三層協調.md`。數值計算頁逐格不變（與 v0.6.4 比對：既有計算頁 0 差異；只新增 Coord 頁、Checks 2 列、README 列位移）。TK_Link 仍為 Tokenomics v21.7（master 現為 v22.0，本輪不換快照）。建置：`python3 builder/build.py --tk-dir <Tokenomics checkout@8309fdb> --out model/20261115_OpenAI_v0.6.5.xlsx --base model/archive/20261029_OpenAI_v0.6.4.xlsx --date 2026-11-15 --tk-date 2026-10-29`。

**Coord 頁（builder/p7.py；只列示，不被計算頁引用）：** A 參數；B 讀取層級（同源，不判定：Tokenomics v21.7；Wattonomics 讀本模型的自有資本支出、營業現金流、2025 年底現金，與本模型現值差皆為 0）；C 層級隱含值（Wattonomics T1 配置給 OpenAI 的年底 GW：2025 年底 0、2026 年底 0.824〔Oracle Abilene〕，2026 年平均 0.412 GW × Tokenomics L1_RevGW_PQ_VR200 49.5（36.2–78.5）＝CY2026 隱含營收 20.4）；D C0；E 層級之間（參考）。

**追溯 Wattonomics oai、usr_oai（tools/coord/trace_wattonomics.py＋人工確認）：** T1 配置為輸入列（Oracle 公告）；自建 GW 2025–2026 為 0、自建 AI 資本支出 2026 為 0；CapexAIModel 讀本模型 Cost（只進對照欄）；Interface 的 CapexAI 取 `IF(G475<>"",G471,G460)`，G475 是空字串常數故取 GW 法＝0（巢狀 IF 已人工確認）；營業現金流、期初現金讀本模型（影響資金缺口，不影響 GW）；usr_oai 只在 Matrix 對照列被讀取，不進計算。

**C0 判定（對帳中 0）：**
- 營收：2026 Q1 5.7＋Q2 6.7＝H1 12.4（SRC_OAI_004／005，公司告知投資人，二手轉載）對層級 20.4 × 6/12＝10.2（＋21.6%）；Q2×4＝26.8 對 20.4（＋31.4%）。**同源：** Tokenomics 的 m 由 OpenAI 營收校準，標「參考（同源校準）」、結論「無可比實績（層級值同源）」，不計入；機械判定欄另列（若不計同源為「對帳中」，僅供參考）。
- 在役算力：公司 CFO 揭露 2025 年底約 1.9 GW（SRC_OAI_090）對層級 T1 配置 0（Wattonomics 只計有逐站證據的配置，空白＝未分配，層級值為下界）：無可比實績；層級值 0，百分比無定義。建議 Wattonomics 端以公司揭露補 2025 年底配置（證據回饋 C5，非本輪判定）。
- 資本支出：OpenAI 無已發生資本支出揭露（計畫值不算）；層級自建 AI 資本支出 0；Wattonomics 的 20 為本模型回寫（同源）：無可比實績。
- 算力承諾：Oracle 4.5 GW（SRC_OAI_061）與 Wattonomics ctr_orcl_oai 4.5 讀同一份公司公告：無可比實績（層級讀同一份公告）。
- 不做兩讀數目標價（模型商，非硬體股）。

**HTML：** `tools/company_html/fundmodel_v2.py` 新增 `coord_c0_generic`（有 `COORD_C0` 才輸出；其他版本逐位元組不變，已以 v0.6.4 重建驗證）；對帳章新增三層協調表，數字全取自 xlsx。

**版面 v1.3：未完成（待 Project 判斷）。** OpenAI 無 Val（估值）、PL（標準損益表）、RevDCF 頁：未上市、無現價與目標價，且歷史三年損益缺（只有 FY2024 營收、FY2025 外流財報一年）；補這三頁需要新的估值參數與損益口徑（判斷類），本輪未動。`check_layout_v13` 對現行個股頁仍有 A（缺 revdcf／pl／drivers 章）、F、G、H 必修項。

## v0.6.4：口徑一 Q2 GW 採 B／C 段、D2 推導式、衍生值移 Inputs（Tokenomics v21.7）
- 口徑一 Q2 季平均 GW 2.26（2.17–2.35）；`PxQ` 新增第 8.4 節（D2 推導式，候選因子）與第九節（Microsoft 10-Q 對照）；SRC_OAI_141–143 退役、改為 Inputs 公式列；`TK_Link` 改取 v21.7。建置：`python3 builder/build.py --tk-dir <Tokenomics checkout@8309fdb> --out model/20261029_OpenAI_v0.6.4.xlsx --base model/archive/20261025_OpenAI_v0.6.3.xlsx --date 2026-10-29 --tk-date 2026-10-29`。範圍外變動比較：`python3 tools/diff_v064.py <v0.6.3> <v0.6.4>`。報告 `docs/reports/20261029_v0.6.4.md`。

## v0.6.3：PxQ 比值 2026 三口徑（Tokenomics v21.4）
- `PxQ` 頁 11 欄：2025 指數／線性 + 2026 口徑一（Q2 年化營收÷Q2 季平均 GW，主）／口徑二（FY2026 營收÷全年平均 GW）／口徑三（9 月底 run-rate÷年底 GW）× 低／基準／高；新增 Fairwater 六情境（第七節）與 GW 路線 token 帳（第八節）。建置：`python3 builder/build.py --tk-dir <Tokenomics checkout> --out model/20261025_OpenAI_v0.6.3.xlsx --base model/archive/20261022_OpenAI_v0.6.2.xlsx --date 2026-10-25 --tk-date 2026-10-25`。報告 `docs/reports/20261025_v0.6.3.md`。

## v0.6.2：Tokenomics 同步 v19.0（單一快照）
- `TK_Link` 改取 Tokenomics master `model/CURRENT`（v19.0）；`TK_PxQ`／`TKQ_` 退役。建置只需一份 Tokenomics：`python3 builder/build.py --tk-dir <Tokenomics checkout> --out model/20261022_OpenAI_v0.6.2.xlsx --base model/archive/20261022_OpenAI_v0.6.1.xlsx --date 2026-10-22`。報告 `docs/reports/20261022_v0.6.2_TK同步.md`；歸因 `tools/tk_attribution.py`。（以下 v0.6.1 的 `TK_PxQ`／`--tk-pxq-dir` 說明為歷史紀錄。）

## v0.6.1：P×Q 與回本比值
- 新增 `PxQ` 頁與 `TK_PxQ`（Tokenomics v19.0 第二快照，具名範圍 `TKQ_`）；`TK_Link` 仍為 v5.26（63 名不變）。建置需兩份 Tokenomics：`python3 builder/build.py --tk-dir <v5.26 checkout> --tk-pxq-dir <v19.0 checkout> --out model/20261022_OpenAI_v0.6.1.xlsx --base model/archive/20261008_OpenAI_v0.6.xlsx --date 2026-10-22 --tk-date 2026-10-08`。
- 主輸出：比值＝實際 m×L ÷ 所需 m×L＝營收 ÷ 算力成本，2025（指數／線性內插）與 2026（2.5／3.0／3.5 GW）。規格（Andy 原話）見 `docs/handoff/OpenAI_handoff.md` 第 9 節；測試把關：`tests/parity/test_builder.py` 的 PxQ 引用白名單與 E6。
- 逐格比對工具：`python3 tools/diff_workbooks.py <舊版> <新版>`；對照 Excel：`tools/make_pxq_report.py`。
