# WhiteFiber 個股模型 — Claude Code 工作守則（v0.4 起；CoreWeave v5.0 新雲樣板，經 Nebius v0.6 分部版）

舊守則（v0.2.x JS 雙引擎、company.json 唯一來源）保留在 `legacy_v02/CLAUDE.md`，已被本檔取代；舊引擎原始碼整套在 `legacy_v02/`（保留不刪）。

## 結構
- `build_whitefiber.py`：唯一計算來源（Python 產生活公式 xlsx，LibreOffice 重算）；事實總帳 `data/whitefiber_facts_*.json`；Tokenomics 快照 `data/tokenomics_snapshot_*.json`；標準損益表歷史 `data/whitefiber_is_hist_*.json`；Wattonomics 追公式輸出 `data/wattonomics_trace_*.json`。
- HTML 由共用產生器 `tools/company_html/` 讀 xlsx（`scripts/make_html_json.py` 產生 `html.json`）；不自帶計算。
- 建置與核對一律 `bash whitefiber/make_v1.sh`（改公式或輸入時加 `SENS=1 REV=1 DEC=1`）；全部通過才可提交。

## 仍有效的已決定事項（承接 CoreWeave v5.0 樣板 CLAUDE.md）
1. 可移植：公司特有內容放 In 頁與事實總帳，程式只寫通用邏輯。
2. 差異要能解釋：與共識、指引的差距寫在 Recon；指引只作 Recon 與 Sens，不作界值（原則 59）。
3. 精簡：HTML 版面規格 v1.3（`tools/company_html/check_layout_v13.py` 必過）。
4. 每季滾動：財報後 FY 首期改為「年初至今實績＋剩餘季度模型」，升版並拆解目標價變動（`scripts/decomp_v50.py`）。
5. HTML 單一檔案、離線開啟（`check_company_html.py`）；可見文字不含網址（網址只在 Sources 頁）。
6. 營運成本以 Tokenomics 滿載穩態為主值；每 MW 收入＝Tokenomics `IF_HoldEcon` × 公司因素 k；不以營收、RPO、合約價、指引反推 k（公司宣稱只進 Recon、Sens、反解）。
7. 公司未揭露季末雲端在役 MW：2026 年雲端在役 MW 為明列假設（In mw_q2、mw_q3、mw_ye26），公司揭露後改為實績並升版。
8. 三層協調：Coord 頁依 `docs/plan/三層基準協調規則.md` 與 `tools/coord/README.md`；Wattonomics 層級值先用 `tools/coord/trace_wattonomics.py --player wyfi` 追公式再判定同源。
9. 金額單位百萬美元（US$M）、股數百萬股；託管（NC-1、MTL、代管服務）以 In 頁 H 節路徑進 Thesis 託管列（三情境相同），只計已簽合約與公司目標的打折值。

## 個股模型共通規格
`docs/plan/個股模型原則_命題加營收分部.md`、`docs/plan/個股DCF標準規格.md`、`docs/plan/個股HTML版面規格.md`、`docs/plan/自算產業β_20261105.md`。

## 溝通
繁體中文；結論先行；判斷題列選項與預設，由 Andy 決定。
