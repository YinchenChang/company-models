# 既有模型改版工作單 v2：DCF＋Comparable 目標價、延伸到 FY2032、現價更新（AI半導體，2026-10-25）

你負責下列資料夾，在自己的 worktree 工作。開工前 `git fetch -q origin main && git reset -q --hard origin/main`（若被拒，確認 HEAD 已等於 origin/main 即可）。

**規格：** 完整照 `docs/plan/AI半導體_公司模型工作單.md` 的「範圍與禁止事項」「預估期間：三情境延伸到 FY2032」「目標價＝DCF 與 Comparable 加權」三節（列標籤逐字照抄）。今天上午加的「0. 目標價與上檔空間」（情境加權 FY2027 規則值）一節，改名為「舊法：規則值情境加權（參考）」保留，不再叫目標價。

**每家另外要做：**
1. **現價更新到 2026-10-23 收盤**（台股、日股若無 10-23 用 10-22）：先讀 `markets/invest_map/data/prices.csv`（兩來源核對過）；沒有的公司上網查兩個來源。只改現價輸入與日期，其餘既有輸入不動。
2. **記憶體三家（skhynix、samsung、micron）**：三情境對應改為 A 順利＝P/B 75 分位、B 平平＝中位數、C 不如預期＝25 分位，C 一定最低；Comparable 用同業前瞻 P/B。
3. 升小版號（例如 v0.1.1 → v0.2；新增估值與延伸期間屬結構改動，次版號 +1），dist 檔名日期 20261025，舊 dist 刪除，README 版本與估值段更新。
4. 若某家原本的情境三者在 FY2028 前完全相同（只差顆數係數），DCF 情境仍照算（差異來自 FY2028 之後），README 註明。

**驗證：** 每家錯誤格 0、重建一致（xlsx_diff 對新 dist 差異 0）、`check_company_html.py <資料夾>` 全過；與舊 dist 比，除現價、新增列、延伸年度、版本文字外，**FY2026–FY2029 的既有 EPS 不得改變**（列一張對照表確認）。下游依賴：amd 讀 nvidia、mrvl 讀 avgo、foxconn 讀 quanta、avc 讀 delta、多家讀 tsmc；同組內先改上游再改下游。

**不要改** `.github/`、`markets/`、其他組的資料夾。完成後一次 commit（「估值 v2：<資料夾清單>」＋兩行 attribution：`Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`、`Claude-Session: https://claude.ai/code/session_01CRVW4qsiiKVECHPubNVS8F`），不 push。

**回報（每家一列，表格）：** 新版號、現價（日期）、DCF 目標價、Comparable 目標價（同業中位數）、加權目標價與上檔空間、市場共識目標價、歷史分位 A／B／C；最後列任何無法照規格做的地方。
