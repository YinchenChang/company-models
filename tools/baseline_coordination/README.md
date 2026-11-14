# 三層基準協調檢查（草稿）

規則本文：`docs/plan/三層基準協調規則.md`（第 10 節列出六項檢查）。

- **現況：** 步驟 2 起草，尚未接 CI。狀態檔 `siliconomics/data/baseline_status.json` 在步驟 3（橋接表量測）後建立；在那之前程式只印「未啟用」並正常結束。
- **執行：** `python3 tools/baseline_coordination/check.py [--status 路徑] [--today YYYY-MM-DD]`；有不符就列出並以 1 結束。
- **狀態檔欄位：** `threshold`（證據門檻，現行 0.5）；`years.<年>`：`wattonomics.evidence_share`、`siliconomics.evidence_share`、`status`（甲／乙／丙）、丙狀態的 `sil_base`、`converted`、`supply_cap`，對帳中的 `reconciling`、`recon_doc`、`deadline`；`level_coefficients`（基準鏈上的水準係數與標籤）；`hardware_companies`；`depends_on_years`（個股 → 估值依賴的年份）。
- 欄位定義若在步驟 3 有調整，以本檔與程式同一個 PR 更新。
