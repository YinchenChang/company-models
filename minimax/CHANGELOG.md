# CHANGELOG

每次同步記錄：版本、工作包、commit／PR 與變動摘要。

## v0.2.3 — 估值頁 Val、反向 DCF 頁 RevDCF（2026-11-15；Andy 決定給目標價與評等）
- 新檔 `val_page.py`、`valcore.py`（與智譜共用）、`scripts/val_snapshots.py`、`data/val_scenarios.json`、`data/val_sens.json`；`build_xlsx.py` 附加 SRC S142–S155、Inputs I35–I61 並建 Val、RevDCF；`scripts/build_all.sh` 先跑快照、加 `check_rating_gate --strict`。
- 加權目標價 HK$147.87（DCF 60.92、Comparable 234.82）、上檔空間 −13.1%、估值評等中立、收支關卡不通過（−5,983.2），投資建議中立。
- 個股模型頁：估值章照 A–I 節（拿掉 data-novaluation）、新增反向 DCF 章、敏感度改加權目標價相對基準主圖。xlsx_diff 對 v0.2.2：既有格只差 README!A1。

## v0.2.2 同版修正（2026-11-15；只改標籤與旗標，數字不變）
- 個股模型頁估值章拿掉 `data-unlisted`、只留 `data-novaluation`（規格 J 節 (b)）；橫幅改「本模型未建目標價與評等」。
- 損益表 FY2026 欄標「1H 實績＋2H 預測」（`build_html_v13.py` 表頭與口徑說明；`coord_pl.py` PL!A1、PL!F5）。
- Valuation V07 標籤註明智譜市值日期：「智譜 市值 ÷ 1H26 年化營收（市值為 2026-10-02，未更新）」。
- 同版重建：xlsx 對修正前只差 Valuation!B14、PL!A1、PL!F5 三格文字；收支模型頁不變；check_html、check_html_v13、ci_check、check_layout_v13 --xlsx 全過。

## v0.2.2 — 版面 v1.3＋三層協調（讀取層級、C0 實績）（2026-11-15；工作單 `docs/workorders/20261114_第二批13家_v1.3與三層協調.md`）
- 新頁 `PL`（只列示）：標準損益表 FY2023–FY2030，非 IFRS 經調整口徑；FY2023–FY2025 取招股書與 FY2025 年度業績公告（SRC_MM 新增 S114–S141 共 28 筆一手歷史列），FY2026 起引用 PnL；所得稅歷史由 IFRS 各項推得為 0；EPS 預測＝經調整淨利 ÷（S051 ×（1＋F14）），歷史非 IFRS 每股未揭露。核對列合計 0。
- 新頁 `Coord`（只列示）：讀取層級（Tokenomics 快照 v19.0；Wattonomics 不讀、亦不讀本模型；Siliconomics 不適用）；C0（實績）四列（1H2026 營收 116.6、資本支出 322.3、算力 MW 未揭露、阿里雲上限 300）皆「無可比實績（層級無此玩家）」——Wattonomics v0.3.2 玩家名單無 MiniMax（中國只有合計 cn）。對帳中 0。具名範圍 COORD_C0、COORD_C0Sum、COORD_C0Verdict、PL_*。
- 新個股模型頁 `build_html_v13.py`（版面 v1.3；本模型無目標價，估值章標「不適用」並帶 `data-unlisted` 以沿用規格 J 節豁免）＋驗收 `scripts/check_html_v13.py`（每個 data-ref 數字對 xlsx）；`build_all.sh` 加入兩步（缺 bs4 時略過）。收支模型頁版本改讀 xlsx。
- 不變性：`tools/xlsx_diff.py` 對 v0.2.1 只差 README A1 版本字串與新增格（README A34–A37、SRC_MM S114–S141、PL、Coord）；Sensitivity 31 案逐格相同。
- 成品 `dist/20261115_MiniMax收支模型_v0_2_2.xlsx／.html`、`dist/20261115_MiniMax個股模型_v0_2_2.html`；v0.2.1 移 `dist/_待刪除/`。

## v0.2.1 — 只換現價（2026-11-06；報告 `docs/reports/20261106_minimax_v0.2.1_現價更新.md`）
- Andy 2026-11-06 07:32「請所有個股都換成十一月五號的收盤價」：SRC_MM 新增 S113（2026-11-05 收盤 HK$170.2，Verified；Yahoo Finance、StockAnalysis 兩來源一致）；Valuation V01／V02 改讀 S113（v0.2 為 S078 10-21 推算 171.9，S078 標【已更新，見 S113】）。
- 影響只在市值相關列：市值 7,716 → 7,640 $M；V13 折現價值 ÷ 目前市值（2030）0.882 → 0.888；V14 支持現價所需 2030 營收 2,207 → 2,192 $M；Funding 新股稀釋比例（F14，依現價發行）小幅上升；Sensitivity J 欄 31 案同步。營收、成本、MW、回本比值不變。
- 成品 `dist/20261106_MiniMax收支模型_v0_2_1.xlsx／.html`；v0.2 移 `dist/_待刪除/`。`build_all.sh 20261106 v0_2_1`：重算錯誤 0、Checks 合計 0、HTML 506 個數字不符 0。

## v0.2 — 模型更新、P×Q 回本比值、成品（2026-10-22；PR #75；報告 `docs/reports/20261022_minimax_v0.2.md`）
- 合併 M2（資料更新）進本分支；SRC_MM 新增 S078–S112 共 35 筆（市場、共識、合約負債、FY2024–25 現金流量、匯率、M3.1 Flash 牌價、中國雲租價、消費端付費人數與 ARPPU）；S052、S062、S077 標「已更新」不刪。Valuation 改用 10-21 推算收盤價 HK$171.9（市值 $9.8B→$7.7B）。
- 新頁 `PxQ`（比照 OpenAI v0.6.1）：回本比值＝營收 ÷（年均上線 MW × 每 MW 年持有成本）；FY2025→2030：0.40→0.87。比值與合約價 I02 成正比（X26–X28）。
- Compute 第七節 D2：中國雲租價證據 ÷ TK 持有成本（I02 維持 12）；第八節 η 對照（TK 無 IF_ 落差分解名稱）；第九節自建集群（找不到規模→不建，註明）。Revenue 第五節消費端付費人數 × ARPPU 對照（僅 9M25）；Revenue 加 Bloomberg 共識等對照；Funding 加 FY2025 經營現金流對照。
- D8：M3.1 Flash（2026-10-19，較 M3 −50%）只更新 2026：新增 I30、I31（Assumed 0.20）、Demand D09；η 仍以降價前 I07 校準；I31＝0 時與 M1 逐 ID 相同（Valuation 除外）。2030 覆蓋率 0.5088→0.5064，累計需新融資 3,159→3,249。
- 新增 Inputs I30、I31、I33；Checks K09–K13；Sensitivity 31 情境（加回本比值欄、M3.1、I02 證據價）。
- HTML：P×Q 區、D2 與 M3.1 表、資料與版本區；每個數字標記並由 `scripts/check_html.py` 對照 Excel（納入 `build_all.sh` 與 CI）；無外部引用。
- 成品 `dist/20261022_MiniMax收支模型_v0_2.xlsx／.html`；v0.1 移 `dist/_待刪除/`。新增 `docs/handoff/MiniMax_handoff.md`、`docs/plan/20261022_MiniMax_命題與驅動對照表.md`、`scripts/xlvals.py`、`scripts/make_v02_compare.py`。

## v0.2-M1 — 流程補齊＋Tokenomics 官方快照取數、同步至 v19.0（2026-10-22；PR #75；報告 `docs/reports/20261022_minimax_M1.md`）
- 新增 `CLAUDE.md`、`CHANGELOG.md`、進度檔、`data/tokenomics_names.txt`、`scripts/tk_snapshot.py`、快照 v5.24／v19.0、`scripts/tk_attribution.py`、`scripts/ci_check.py`、CI `.github/workflows/minimax-build.yml`。`build_xlsx.py` 的 TK 改讀快照；TK_Link 加檔名／提交／SHA-256／讀取日與 IF_HoldEcon 全 15 欄。
- 以 v5.24 快照建置與 v0.1 逐格相同（2,396 格僅 TK_Link!A2 說明文字不同）。
- 同步 Tokenomics v19.0（`1916a06`，SHA-256 `d1728705…`）：IF_TokGW_* +34%～+57%、η 2025 3.42→2.25、1H26 0.80→0.52；覆蓋率 2030 0.5085→0.5088、累計需新融資 3,167→3,159 $M。`dist/` 未覆蓋。

## v0.1 — 首版（2026-10-08；PR #25）
- `dist/20261008_MiniMax收支模型_v0_1.xlsx／.html`：541 個公式，重算錯誤 0，Checks 合計 0；SRC_MM 77 列、Inputs 29 項。
- Tokenomics 數值：v5.24（master `bdb0de7`）12 格，**經 OpenAI v0.6 TK_Link 頁手抄**寫死於 `build_xlsx.py`（二手；無快照檔、無重現檢查）。
- 2026-10-08 PR #36（Tokenomics 連接審視）：只審視不改模型；以 v5.29（`a5061d9`）重取 12 格全部相同（差異 0）；列出「改用官方快照」等建議，成為 v0.2 M1 的依據。
