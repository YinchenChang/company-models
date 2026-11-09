# tools/model_governance — 個股模型治理核對

起因：2026-11-09 SpaceX v0.2.5 驗收報告（M1 Sens 失同步、M8 標記混用）。原則見 `docs/plan/個股模型原則_命題加營收分部.md` 第 10 節。

`python3 tools/model_governance/check_governance.py <重算後 xlsx> [--html <成品 html>]`

| 檢查 | 失敗條件 |
|---|---|
| build_id | Sens 頁無「build_id」列；README!A1 或 HTML 不含同一字串 |
| 模型狀態 | Sens 頁「模型狀態」不是 CURRENT |
| 單一標記 | In 頁 G 欄、Q 頁「標記」欄不是 Verified／Interested-party／Analogy／Assumed／Derived／— 之一；來源以「…；Derived）」或「…；Interested-party）」結尾卻標 Verified |
| 來源誘因 | 沒有「來源誘因」欄，或 Interested-party 列該欄空白 |

已套用：SpaceX v0.2.6。其他個股於下次改版時加入建置腳本（做法照 `spacex/build_spacex.py` 的 `BUILD_ID`、`norm_tag`、Sens 指紋三段）。
