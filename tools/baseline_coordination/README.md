# 三層基準協調檢查

規則本文：`docs/plan/三層基準協調規則.md`（第 10 節列出檢查項目，現行 v1.4.3 共 12 項）。

- **現況（v1.4.3）：** 已接 CI：`.github/workflows/coord-check.yml`（PR 與 push 到 main，路徑含 `wattonomics/**`、`siliconomics/**`、`tools/baseline_coordination/**`、`tools/coord/**`、規則本文）。同一 workflow 另跑 `wattonomics/tools/check.py`（讀活頁簿快取值，不需要 LibreOffice）。狀態檔 `siliconomics/data/baseline_status.json` 不存在時只印「未啟用」並正常結束。
- **執行：** `python3 tools/baseline_coordination/check.py [--status 路徑] [--today YYYY-MM-DD]`。
- **輸出與結束碼：** 「[n] …」＝不符，有任一項以 1 結束；「提醒 [n] …」＝未量測、結構無法判定或豁免中的事項，只列出、以 0 結束。最後一行「N 項不符、M 項提醒」。
- **狀態檔欄位：** `threshold`（證據門檻，現行 0.5）；`years.<年>`：`wattonomics.evidence_share`、`siliconomics.evidence_share`、`status`（甲／乙／丙／待定）、`label`、丙狀態的 `sil_base`、`converted`、`supply_cap`，對帳中的 `reconciling`、`recon_doc`、`deadline`；`level_coefficients`（基準鏈上的水準係數與標籤）；`hardware_companies`；`depends_on_years`（個股 → 估值依賴的年份）。
- **v1.1 新增欄位：** 頂層 `wattonomics_coverage_complete`（Wattonomics 覆蓋是否已補齊）；`years.<年>.lead`（乙狀態由哪一層定：`wattonomics`／`siliconomics`）、`timing_cum`（C2 ② 累計 GW）、各層 `prev_has_evidence`（上次判定是否有證據，供緩衝區：進入 ≥ 55%、退出 < 45%；缺值＝首次判定，以門檻為界）。
- **v1.4.3 新增欄位：**
  - 頂層 `wattonomics_version`（例 `"v0.3.6"`）：狀態檔換算值與 Wattonomics 證據比例取自的 Wattonomics 版本。缺欄位時程式改從 `_note` 擷取版本字串印提醒。
  - `years.<年>.c0_fail`（布林）：該年某層的 C0 公司實績對帳（純實績，只到最近已公布期間）超出 ±15% 容差。為 true 時須同時有 `c0_recon_doc`（對帳單位置）與 `c0_deadline`（YYYY-MM-DD，下一個財報季結束），`label` 在甲年須含「有條件」（例：「有條件一致」）；待定、乙、丙年份標示不變（規則 C0 第 8 點）。只認布林 true（字串 "false" 判不符）。用獨立欄名、不共用 `recon_doc`／`deadline`，因為 C0 對帳單可能與層間對帳單不同張，且「有條件一致」的年份不標 `reconciling`。若與層間對帳單是同一缺口，依規則 C0 第 10 點合併為一張，兩組欄位填同一份對帳單。
- **各項做法（v1.4.3 新增或改動）：**
  - [7] 年份缺 `timing_cum` → 「提醒 [7] 未量測」（原為靜默略過）。
  - [9] 讀 `siliconomics/data/CURRENT_RECON` 指向的 recon 快照。市場有 `status` 欄時才判定：`"未建"` 的市場列數與名稱作提醒；其他狀態（已建）的 `rows` 須有 `who` 含「殘差」「其他」或「未建模」的列，否則不符。**限制：** 現行 `siliconomics/tools/recon.py` 產生的快照沒有 `status`、市場值與殘差欄位（`soc` 只有兩家個股份額列，`r4`、`nv_dc` 是層間／實績對帳點，不是個股加總），結構無法判定 C0 恆等式，只印提醒；未建市場清單目前只在 Siliconomics Recon 頁 D 節，未進快照。要讓 [9] 真正把關，`recon.py` 須補每市場 `status`、市場值與殘差列。
  - [10] 讀 `wattonomics/data/CURRENT_SNAPSHOT` 的 `names`：對每個 `WIF_Evidence_<yyyy>`，若有 `WIF_Evidence_Lo_<yyyy>` 與 `WIF_Evidence_Hi_<yyyy>`，須 低 ≤ 基準 ≤ 高，否則不符；沒有這兩組名稱 → 提醒（Wattonomics v0.3.6 起新增）。
  - [11] 版本漂移：`wattonomics_version` 與 `wattonomics/model/CURRENT` 檔名中的版本不同 → 不符；欄位不存在 → 提醒（附 `_note` 擷取的版本）。
  - [12] `c0_fail` 為 true：甲年 `label` 須含「有條件」；所有年份須有 `c0_recon_doc`、`c0_deadline`（YYYY-MM-DD，格式錯誤判不符）；期限已過 → 不符（同 [5]，依 C6 第 3 點強制裁定）。協調年份以外的已公布年份（例 2025）記在頂層 `c0_actuals.<年>`（欄位同上，不需甲乙丙標籤；C0 第 9 點）。
  - [10] 只有 Lo 或只有 Hi → 不符；兩者皆缺 → 提醒，逐年列出。
  - [6] 只在 coord-check 被觸發時檢查；個股 HTML 變動也會觸發（paths 含 `*/dist/*.html`）。有條件一致的年份不要求兩讀數（C0 第 8 點例外於 C4）。
- 欄位定義若有調整，以本檔、程式與狀態檔匯出程式（`siliconomics/tools/export_baseline_status.py`）同一個 PR 更新。
