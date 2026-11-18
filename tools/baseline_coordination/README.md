# 三層基準協調檢查

規則本文：`docs/plan/三層基準協調規則.md`（第 10 節列出檢查項目，現行 v1.4.3 共 12 項；2026-11-18 起另有程式內的 [13]、[14]、[15] 三項，三層對帳工作單 A8、A7、W2-05B，尚未寫入規則本文第 10 節）。

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
  - [9] 讀 `siliconomics/data/CURRENT_RECON` 指向的 recon 快照。**Siliconomics v0.20 起**（`siliconomics/tools/recon.py`）每個市場有 `kind`、`status`、`status_by_year`：
    - `kind`＝`market`（由下往上市場）：`status`＝`"未建"` 的市場列名稱作提醒；已建者 `rows` 須有殘差列（`residual: true`，或 `who` 含「殘差」「其他」「未建模」），殘差值不得為空；有 `market`（市場值）時逐年檢查 Σ `rows`（含殘差）＝`market`（相對誤差 1e-6），否則不符。現況：`soc`（SoC 測試機，份額口徑，市場值＝100%）、`gpu_asic`（Recon D1 GPU／ASIC 晶片金額）已判定；CoWoS、晶圓前段設備、AI 高速 PCB／CCL、光學模組總量為「未建」。
    - `kind`＝`structural`（Recon C 節結構一致市場，恆等式由 Siliconomics Checks K6 約束）、`layer`（R4 層間對帳點，狀態讀 Coord 頁）、`actual`（NVIDIA 資料中心營收公司實績對帳點）：不是個股加總，不檢查恆等式，只計數提醒。
    - 無 `kind` 的舊快照（v0.19 以前）：市場有 `status` 時同上判定；無 `status` 而有 `rows`＝結構無法判定，只印提醒。
  - [10] 讀 `wattonomics/data/CURRENT_SNAPSHOT` 的 `names`：對每個 `WIF_Evidence_<yyyy>`，若有 `WIF_Evidence_Lo_<yyyy>` 與 `WIF_Evidence_Hi_<yyyy>`，須 低 ≤ 基準 ≤ 高，否則不符；沒有這兩組名稱 → 提醒（Wattonomics v0.3.6 起新增）。
  - [11] 版本漂移：`wattonomics_version` 與 `wattonomics/model/CURRENT` 檔名中的版本不同 → 不符；欄位不存在 → 提醒（附 `_note` 擷取的版本）。
  - [12] `c0_fail` 為 true：甲年 `label` 須含「有條件」；所有年份須有 `c0_recon_doc`、`c0_deadline`（YYYY-MM-DD，格式錯誤判不符）；期限已過 → 不符（同 [5]，依 C6 第 3 點強制裁定）。協調年份以外的已公布年份（例 2025）記在頂層 `c0_actuals.<年>`（欄位同上，不需甲乙丙標籤；C0 第 9 點）。
  - [10] 只有 Lo 或只有 Hi → 不符；兩者皆缺 → 提醒，逐年列出。
  - [6] 只在 coord-check 被觸發時檢查；個股 HTML 變動也會觸發（paths 含 `*/dist/*.html`）。有條件一致的年份不要求兩讀數（C0 第 8 點例外於 C4）。
- **Siliconomics v0.20 狀態檔（`export_baseline_status.py`）已輸出：** `wattonomics_version`（取自 `wattonomics/model/CURRENT`）、逐年 `c0_fail`（讀 `SIF_RecoC0Fail_<年>`）與 `c0_recon_doc`／`c0_deadline`／`c0_point`（`builder/seed/coord_meta_v0.20.json` 的 `c0`）、頂層 `c0_actuals.<年>`（值讀 Wattonomics 快照 `WIF_Recon_NV_*_<年>`，對帳單與期限讀 meta 的 `c0_actuals`）。
- **2026-11-18 新增（三層對帳工作單 A1、A7、A8；不改任何基準數值）：**
  - `years.<年>.timing_cum`（C2 ② 累計，**GW**）與 `timing_cum_ratio`（② 累計 ÷ 當年基準）：`export_baseline_status.py` 讀協調橋接檔 Bridge 頁 `g2c`、`g2p`（`siliconomics/docs/reports/*協調橋接_<Wattonomics 版本>*.xlsx` 最新一份）。[7] 因此四年皆有實際判定（現況 11.9／11.2／10.7／10.9%，皆 ≤ 15%）。找不到橋接檔時不寫欄位、[7] 仍印「未量測」。
  - 頂層 `method`、`generated`、`source_sha256`（Siliconomics 活頁簿）、`wattonomics_source_sha256`、`bridge`（`file`、`sha256`、`in_title`、`note`）。逐年 `legacy_s33`（舊 S33 取法，僅供對照；`current_s33` 保留一版作別名）。
  - [13] 來源雜湊一致：Siliconomics 與 Wattonomics 的 CURRENT 快照 `source.sha256` ＝ `model/` 內該檔實際 SHA-256；狀態檔 `source_sha256`、`wattonomics_source_sha256`、`bridge.sha256` 亦同；不一致＝不符。活頁簿重存（LibreOffice 重算）會改變位元組，故快照與狀態檔匯出須是定稿後的最後一步。
  - [14] 欄位語義（`SIF_DeployGW_*` 五組，語義見 `siliconomics/README.md`「欄位語義」）：硬體個股（`hardware_companies` 資料夾 `*.py`）不得讀 `Current`、`Demand`（唯一登錄用途：`nvidia/coord_post.py` 的「改前現行值」對照欄與註解，`SEM_ALLOW`）；AST 檢查「目標名稱含 base 的指派」右側不得含 `Current／Supply／Demand／Alt／AltRatio`（`Supply`、`Alt`、`AltRatio` 只可並列顯示或作情境）。已知例外登錄於 `SEM_BASE_EXC`（現況：`nvidia/coord_post.py` 的退回路徑，以提醒列出；例外已不需要時也以提醒列出）。靜態檢查，擋不住間接使用。
  - [15] 只顯示的觀測數列（W2-05B；依據：**規則 v1.5 草案 C3（待 Andy 核可）；本項只防止只顯示數列被讀取，不改任何判定**）。狀態檔 `years.<年>.observations`（`mid`＝i／ii／iii／lo／hi 中間估計、`decay`＝λ 0／0.5／1 衰減寫法、`display_only`、`rule`；不適用年份的子項為 null）只供 Siliconomics Coord 頁顯示。
    - (a) 狀態檔每年若有 `observations`，`display_only` 必須為 `true`，否則＝不符；沒有 `observations`（舊狀態檔或不適用年份）只列提醒。
    - (b) 掃描全 repo 的 `*.py`（排除 `siliconomics/`、`docs/`、`tools/baseline_coordination/`、`.git` 與隱藏資料夾），先挑出文字含 `SIF_`、`baseline_status` 或 `siliconomics`（不分大小寫）的腳本，再以 AST 比對：字串常數（含 f-string 片段與下標鍵）整串等於 `observations`、字串含整詞 `mid_i／mid_ii／mid_iii／mid_lo／mid_hi／dec0／dec05／dec1` 或前綴 `crd_lam_`、屬性 `.observations`。命中＝不符，印檔名與行號；無命中時印一則提醒說明掃描支數。
    - 已知限制：靜態檢查。(1) 註解與說明文字不進 AST，所以不會誤判（也不會擋註解）；(2) 整詞比對要求前後不緊接英數或底線，所以 `w_mid_lo` 之類不算命中，但名稱前面加了底線前綴的真實鍵（例：`SIF_x_mid_i`）也不會被抓；(3) 只看字串常數與屬性，變數名稱（例：變數 `dec1`）不算；(4) 以動態組字串（拼接、`getattr`）取得 `observations` 擋不住；(5) 只掃 `*.py`，不含 JS／HTML／Excel 公式；(6) 檔案文字完全不含 `SIF_`／`baseline_status`／`siliconomics` 者不掃。
- 欄位定義若有調整，以本檔、程式與狀態檔匯出程式（`siliconomics/tools/export_baseline_status.py`）同一個 PR 更新。
