# Wattonomics 互動頁（explorer）

- 線上版（Artifact，私人）：https://claude.ai/artifact/JxC86we7i8tKkd84xThNyu
- 資料：`wattonomics/data/explorer_v*.json` 經 `build_page.py` 壓成 data.json 後替換 `template.html` 的 `__DATA__`。
- 驗收（v0.1.1，2026-10-28）：未參與代理比對 971 個值與 Excel 0 差異；修正園區計入旗標判斷、2025 年增加含新納入存量的標示、資金比例部分期間註記。
- 規劃：AI 投資地圖 v0.6（PR #190）上線後，以 v0.7 分頁併入，讀同一份 explorer JSON（Andy 2026-10-28 同意）。
