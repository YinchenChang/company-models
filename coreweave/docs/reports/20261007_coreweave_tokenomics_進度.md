# CoreWeave × Tokenomics 改造進度（接手用；W1–W3 共用）

## 目前狀態（每次 push 前覆寫）
- 已完成：W0、W1（PR #6，待審）。**W2 進行中**（PR 見下表，疊加於 W1 分支 8d97234）：第 0 步完成。
- 下一步：W2 第 1 步（世代組合 `fleet`）。逐步紀錄見下方「W2 每 MW 改寫」。
- 未解問題：(1) Tokenomics v5.26（原 v5.25 名稱，IF_MaintIT／IF_StaffSW／IF_TaxIns／IF_DeprLifeIT 等）尚未建置，W2 以暫代值實作；(2) GB200／GB300／VR200 長約 GPU 小時價格不足兩個獨立來源 → 收入採備案 `revenue=legacy`；(3) CoreWeave「active power」口徑定義未找到（預設 IT）；(4) Tokenomics 無租金名稱（租金維持公司專屬）。

## 工作單總覽
| 工作單 | 分支 | PR | 狀態 |
|---|---|---|---|
| W0 遷移 | `claude/coreweave-w0-migrate` | | chat 端完成 |
| W1 Tokenomics 取數層 | `claude/coreweave-w1-tokenomics`（疊加於 W0 分支） | #6 | 完成，待審 |
| W2 每 MW 改寫 | `claude/coreweave-w2-permw`（疊加於 W1 分支） | #9 | 進行中 |
| W3 v4.6 成品與對照 | `claude/coreweave-w3-v4.6` | | 未開始 |

<!-- 各工作單在下方新增自己的段落：「## Wx」＋步驟紀錄表（步驟｜狀態｜commit｜備註） -->

## W1 Tokenomics 取數層

Tokenomics 版本：`model/CURRENT`＝`20261007_Tokenomics_v5.24.xlsx`，master HEAD `098873a3c6855d1ef5e3414d5a0f12e39b2644ed`（2026-10-07 讀取）。

| 步驟 | 狀態 | commit | 備註 |
|---|---|---|---|
| 0 開分支、draft PR #6、Tokenomics 唯讀副本 | 完成 | d4775c1 | W0 未合併：自 `origin/claude/coreweave-w0-migrate` a29db91 開分支，PR base＝W0 分支 |
| 1 取數工具＋根目錄 README 段落 | 完成 | （本 commit） | 測試：產生 v5.24 快照 25 名（missing 10）；`--check` 通過；`--force-recalc`（LibreOffice 重算副本）值＝快取值；IF_Hdr*／不存在名稱報錯；竄改快照 1e-6 時 `--check` 失敗；無 clone 時警告略過 |
| 2 名稱清單、快照、company.json `tokenomics`、README 欄位表 | 完成 | （本 commit） | `coreweave/data/tokenomics_names.txt`（15 個現有名稱＋10 個 optional；另加 `IF_RacksPerGW`、`IF_Util` 供對照）；`coreweave/data/tokenomics_snapshot_v5.24.json`；`fields_doc.py --write` |
| 3 v5.25 名稱列 optional | 完成（以 v5.24 產快照） | 同上 | Tokenomics master 仍為 v5.24（098873a）；10 名記為 `{"missing": true}`。**W2 第 0 步：若 master 已是 v5.25，重抓快照（檔名改 v5.25）並刪 v5.24 快照** |
| 4 Excel「Tokenomics_取數」分頁、TK_ 具名範圍、verify 新增 0c／5d | 完成 | （本 commit） | 分頁放在「來源」之後（最後一頁，灰色標籤）；55 個具名範圍（missing 名稱有列、不建名稱）；`xl_diff.py` 新增工作表列為「新增工作表」不計差異；`verify.sh --vs-dist` **23 項全過**（原 21＋0c＋5d），crawl 0 差異、xl_diff 值與公式 0 差異 |
| 5 對照表（給 W2） | 完成 | （本 commit） | `coreweave/docs/plan/20261007_coreweave_Tokenomics對照表.xlsx`＋`.md`，由 `coreweave/scripts/build_tk_map.py` 自 company.json＋快照產生；11 列（工作單 10 項＋GPU 數、GPU 小時持有成本兩列 W2 新增項）；比較世代預設 FY26–27＝GB300、FY28–30＝VR200 [Assumed] |
| 6 資料蒐集（MW 口徑、世代、GPU 小時價格、10-Q 費用、站點電價） | 完成 | （本 commit） | `coreweave/data/permw_inputs_20261007.json`；摘要見下方「W1 第 6 步資料摘要」 |

### W1 第 6 步資料摘要（完整逐筆在 `coreweave/data/permw_inputs_20261007.json`）

**6a MW 口徑：視為 IT 關鍵電力 [Assumed]（有三項佐證）**
- CoreWeave 自己的文件（10-K FY2025、10-Q Q2 2026、Q2 新聞稿、法說逐字稿）只給數字（年底 >850 MW、6/30 1.5 GW active；簽約 3.7 GW，法說日 4.2 GW），找不到「active／contracted power」是 IT 或設施口徑的定義。
- 房東以 critical IT 計算對 CoreWeave 的容量：Core Scientific（2025-02-26：590 MW，Denton「full critical IT load just over 260 MW」，並把 1.3 GW 電網合約電力與約 900 MW 可出租容量分開）、Galaxy Helios（2025-04-23：「393 MW of critical IT load」）[Interested-party]。
- 密度對照：2024 年底約 250,000 GPU ÷ 360 MW ≈ 694 顆/MW，接近 Tokenomics Hopper 每 IT MW 718 顆（−3%），遠於設施口徑的 598 顆（+16%）[Derived，只作口徑判斷]。
- 若改為設施口徑：IT MW＝設施 MW ÷ 1.2（6/30 1,500 → 1,250 IT MW）；CRWV 每 MW 數字換成每 IT MW × 1.2（每 MW 年收入 11.2 → 13.4、建置成本 34 → 40.8 US$m）；每 MW 利潤率不變。

**6b 世代組合（6/30 在役 1,500 MW）：Hopper 27%／GB200 29%／GB300 43% [Assumed]**，區間 Hopper 24–31%、GB200 20–42%、GB300 30–53%。方法：依年份分層——2024 年底前 360 MW＝Hopper（含 A100）；2025 年新增 490 MW＝GB200 70%／GB300 20%／H200 10%；2026 上半年新增 650 MW＝GB300 85%／GB200 15%。MW 錨點 [Verified]，各層占比 [Assumed]（公司未揭露機隊世代結構）。2H26–FY27 新增：GB300 為 2026 主力、Vera Rubin Q2 2026 才完成首套驗證、Meta $21bn 擴約（2026-04-09）含首批 Vera Rubin——與 W2 工作單 `newMix` 預設方向一致，無更精確資料。

**6c GPU 小時價格（US$/GPU-hr）：不足以支撐 W2 的 GPU 小時價格路線 → 建議 W2 採備案（revenue=legacy）**

| 世代 | 數值 | 口徑 | 來源（日期） | 標記 |
|---|---|---|---|---|
| H100 | 6.16 | 隨需牌價（8 GPU $49.24/hr；長約最多 −60%） | CoreWeave 價目（2026-10-07 擷取） | [Interested-party] |
| H100 | 2.82 | 市場租金指數 | Silicon Data（2026-10-06） | [Verified] |
| H100 | 3.85（10/1 起 4.50） | 隨需牌價 | Nebius 價目 | [Interested-party] |
| H100 | 3.99 | 隨需牌價 | Lambda 價目 | [Interested-party] |
| GB200 | 10.50 | 隨需牌價（4 GPU $42/hr；長約最多 −60%） | CoreWeave 價目 | [Interested-party] |
| GB200 | 8.00 起 | 隨需牌價 | GMI Cloud（2026-07-24） | [Interested-party] |
| GB200（以 B200 類比） | 5.87（區間 4.08–6.11） | 市場指數 | Silicon Data（2026-10-06）；Ornn 4.08（2026-04-13，經 WSJ 轉述） | [Analogy] |
| GB300（以 B300 類比） | 6.82（區間 6.82–7.85） | 市場指數 | Silicon Data（2026-10-06） | [Analogy] |
| GB300（以 B300 類比） | 7.85（10/1 起 9.50） | 隨需牌價 | Nebius 價目 | [Analogy] |
| GB300 | 9.70 US$m/MW/年（≈2.27/GPU-hr） | 同業五年合約（$9.7bn ÷ 5 年 ÷ 200 MW critical IT；GPU 數未揭露，換算用 Tokenomics 487 顆/MW） | IREN–Microsoft（2025-11-03） | [Derived] |
| VR200 | 找不到 | — | CoreWeave／Nebius／Lambda 價目未列；Meta 合約未揭露單價 | — |

- 牌價都不含網路與儲存（CoreWeave、Nebius 儲存另計）。所有合約期間：牌價＝隨需；指數＝市場現行租期；只有 IREN 為 5 年合約。
- 未使用 SemiAnalysis／InferenceX。GB200 的 CoreWeave 牌價與 Tokenomics `L1_GPUhr_GB200_vsCW` 外部值為同一底層來源，不算獨立。

**6d 公司營運費用（10-Q Q2 2026，US$m，三個月）**：營收 2,575；營收成本 879；技術與基礎設施 1,507；銷售行銷 60（其中 SBC 12）；一般管理 178（其中 SBC 84）；營業損益 −49；D&A 合計 1,393（未依列別揭露）；SBC 合計 165（技術與基礎設施 60、營收成本 9）；營業租賃成本 500、變動租賃 150、融資租賃 17。
- **管銷率＝（60 − 12 ＋ 178 − 84）÷ 2,575 ＝ 142 ÷ 2,575 ＝ 5.51% [Derived]**；假設：銷售行銷與一般管理內的 D&A＝0 [Assumed]。區間 5.51%（扣 SBC）～9.24%（GAAP 含 SBC）；1H26 扣 SBC 為 6.15%。

**6e 站點電價**：找不到（CoreWeave 與房東 Applied Digital、Core Scientific、Galaxy 公告皆未揭露電價或電費負擔方）；電價以 Tokenomics 為準。
| 7 整體 verify 與回報 | 完成 | （本 commit） | `verify.sh --vs-dist` 23 項全過；PR 留言「[CRWV 回報] W1｜完成｜2026-10-07」 |

### W1 已套用的預設（問題｜採用的預設｜替代選項｜對結果的影響）
| 問題 | 預設 | 替代 | 影響 |
|---|---|---|---|
| 分頁 A 欄 | A 欄放具名範圍鍵（`TK_CapexIT_GB300`），其後依工作單欄位 | A 欄放 Tokenomics 名稱（同名多世代會重複，違反 CLAUDE.md 勿改 7） | 無數字影響 |
| L1 名稱的低／高 | 取 L1 頁 E／F 欄（Tokenomics 自己的低／高） | 留白 | 只影響敏感度欄 |
| 名稱清單 | 工作單 13 名＋`IF_RacksPerGW`、`IF_Util`（對照用）＋10 個 v5.25 optional | 只列工作單 13 名 | 無 |
| 分頁位置 | 「來源」之後（最後一頁），導覽頁不加列 | 導覽加一列（會使導覽列位移） | 無 |
| missing 名稱 | 分頁有列、不建具名範圍 | 建名稱指向空格 | W2 若誤引用會在建置時就報錯，較安全 |
| 新工作表的比對 | `xl_diff.py`：新版多出的工作表列為「新增工作表」，不計差異 | `--ignore=Tokenomics_取數` | 無 |
| 對照表比較世代 | FY26–27＝GB300、FY28–30＝VR200 [Assumed] | 依 W2 newMix 加權 | 差距數字會隨世代組合改變（例：每 MW 年收入 vs 持有成本 FY26 −12.0%；改 GB200 為 +22%） |
| MW 口徑 | IT 關鍵電力 [Assumed] | 設施口徑 | 每 MW 數字差 1.2 倍（約 ±17–20%） |
| 6/30 世代組合 | 年份分層 27／29／43% | 區間見上 | W2 敏感度 |
| 管銷率 | 5.51%（扣 SBC；S&M／G&A 內 D&A 視為 0） | 9.24%（GAAP 含 SBC） | 改用 9.24% 時 EBITDA 率約 −3.7pt |
| 每 MW 年收入的 Tokenomics 對照 | `IF_HoldEcon`（經濟持有成本＝打平下限） | `L1_RevGW_Fleet_VR200`（token 層上限，不同層） | 無數字影響 |
| 快照版本 | v5.24（master 的 CURRENT 仍為 v5.24） | — | W2 第 0 步重抓 |

### W1 verify.sh --vs-dist 完整輸出（2026-10-07，最終）
<details><summary>展開</summary>

```
=== 0. 季度層檢查（季度加總＝年度、指引一致性、超過門檻的差距都有原因；v4.4）
一致性檢查 14 組
low：季度加總＝年度 10 項
low：年度差異原因 7 項、季度差異原因 3 項
base：季度加總＝年度 10 項
base：年度差異原因 6 項、季度差異原因 3 項
high：季度加總＝年度 10 項
high：年度差異原因 6 項、季度差異原因 3 項
check_quarterly：全部通過
[PASS] check_quarterly：季度加總＝年度、差異原因齊全

=== 0b. README 欄位說明與 company.json 一致（scripts/fields_doc.py --check；5a）
README 欄位說明與 company.json 一致
[PASS] README 欄位說明與 company.json 一致

=== 0c. Tokenomics 快照可重現（tools/tokenomics/import_tokenomics.py --check；W1）
--check 通過：25 個名稱（missing 10）與快照一致（20261007_Tokenomics_v5.24.xlsx，相對誤差 ≤ 1e-09）
[PASS] Tokenomics 快照 --check：25 個名稱（missing 10）與快照一致（20261007_Tokenomics_v5.24.xlsx，相對誤差 ≤ 1e-09）

=== 1. 建 HTML（v4.5 · 2026-09-26）
built 965548
[PASS] 建 HTML

=== 2. 建 Excel
saved /home/claude/company-models/coreweave/out/20260926_CoreWeave收支模型_v4_5.xlsx
[PASS] 建 Excel

=== 3. 重算（LibreOffice headless）
{"status": "success", "total_formulas": 1955, "total_errors": 0, "error_summary": {}}
[PASS] 重算：0 公式錯誤

=== 3b. 反向 DCF：以 Excel 求解（scripts/rv_solve.py；v4.5 取代 JS 快照）
反向 DCF（Excel 求解）：現價 90.13、R 1.3428、C 0.5970、Eb 0.8566、Rt 1.1801；矩陣 20/20 格有解
rv_snap.json 無變動
[PASS] 反向 DCF：Excel 求解，rv_snap.json 與 Excel 一致

=== 4. fix_outline
outline fixed
[PASS] fix_outline

=== 4b. fix_datatable（模擬運算表還原為 Excel 格式）
datatable fixed: xl/worksheets/sheet3.xml D171:E173 輸入格 C5
[PASS] fix_datatable

=== 5. verify_ooxml
OOXML OK
[PASS] verify_ooxml：OOXML OK

=== 5d. 快照值＝Excel「Tokenomics_取數」分頁值（scripts/check_tokenomics_tab.py；W1）
快照值＝Excel 分頁值：25 個名稱（missing 10）、165 個值、55 個具名範圍一致；其他工作表無公式引用（v5.24）
[PASS] 快照值＝Excel 分頁值：25 個名稱（missing 10）、165 個值、55 個具名範圍一致；其他工作表無公式引用（v5.24）

=== 5c. 離線開啟檢查（已決定事項 11：單一檔案、無網路請求、console 無錯誤、關鍵數字正常顯示）
(a) 網路請求：0 個
(b) console 錯誤與例外：0 個
(c) 關鍵數字：5 個分頁、15 項，失敗 0 項
離線開啟檢查：通過
[PASS] 離線開啟：新建 HTML
(a) 網路請求：0 個
(b) console 錯誤與例外：0 個
(c) 關鍵數字：5 個分頁、15 項，失敗 0 項
離線開啟檢查：通過
[PASS] 離線開啟：dist/ 成品

=== 6. 三情境 xlx＋cmp31（預設錨定）
ok 1 820
[PASS] cmp31 low：313 項全部 OK
ok 2 820
[PASS] cmp31 base：313 項全部 OK
ok 3 820
[PASS] cmp31 high：313 項全部 OK

=== 7. FY27 錨定 cmp31（基準）
ok 2 820
[PASS] cmp31 base_FY27：313 項全部 OK

=== 8. 季度層測試（暫存副本：假設 Q3 實際數、可移植性；v4.4）
[PASS] test_quarterly：假設實際數與可移植性測試通過
  畫面錯誤：無
=== 結果
test_quarterly：全部通過（測試 A 假設實際數、測試 B 可移植性）

=== 8b. 期間滾動測試（暫存副本：日曆推算 6 種情況、滾動後第一屏無過期日期與期間字樣；v4.5）
[PASS] test_rolling：日曆推算與滾動後第一屏
C 滾動檢查：只滾日曆、未更新 asOf → 建置失敗並列出 23 項（清單 23 項）
A 滾動 FY26Q3（評價日 2026-06-30 → 2026-09-30）：過期字樣 ['2026-06-30', 'Q2 2026', '2H26', '6/30', '1H26', '上半年', '下半年', 'H1', '1H', '2H']；第一屏命中 0 行
期間滾動測試：通過

=== 8c. 目標價變動拆解工具測試（scripts/attrib.py：(a)＝(1＋WACC)^(月數÷12)，WACC 讀 Excel、月數讀 calendar_q；四項相加＝總變動）
[PASS] test_attrib：拆解工具（月數、同版 0、滾動一季與 WACC 12%）
目標價變動拆解測試：通過

=== 9. 與 dist/ 成品比對
34 views; errors []
34 views; errors []
畫面：dist 34 個、新版 34 個頁面／分頁，約 129,242 字；差異 0 個；頁面錯誤 dist 0、新版 0
[PASS] crawl 畫面文字 0 差異
0 differences
新增工作表（不計為差異）：Tokenomics_取數
以列名稱配對的工作表：輸入與假設
新增列 6：
  輸入與假設 第 13 列 表外租金起算 MW
  輸入與假設 第 89 列 股權上限的股數基礎
  輸入與假設 第 93 列 CDS 傳入門檻
  輸入與假設 第 94 列 CDS 傳入比例
  輸入與假設 第 109 列 NOL 每年可抵用比例
  輸入與假設 第 110 列 營運資金占營收增量
[PASS] xl_diff --values --by-label 0 差異
0 differences
新增工作表（不計為差異）：Tokenomics_取數
以列名稱配對的工作表：輸入與假設
新增列 6：
  輸入與假設 第 13 列 表外租金起算 MW
  輸入與假設 第 89 列 股權上限的股數基礎
  輸入與假設 第 93 列 CDS 傳入門檻
  輸入與假設 第 94 列 CDS 傳入比例
  輸入與假設 第 109 列 NOL 每年可抵用比例
  輸入與假設 第 110 列 營運資金占營收增量
常數改為引用輸入格 30（引用換回常數後與舊版公式相同）：輸入與假設!C60、輸入與假設!D60、輸入與假設!E60、輸入與假設!F60、輸入與假設!G60、各期收支!C48、各期收支!D48、各期收支!E48、各期收支!F48、各期收支!G48、各期收支!C60、各期收支!D60、各期收支!E60、各期收支!F60、各期收支!G60、損益!C14、損益!D14、損益!E14、損益!F14、損益!G14、評價_DCF與目標價!C8、評價_DCF與目標價!D8、評價_DCF與目標價!E8、評價_DCF與目標價!F8、評價_DCF與目標價!G8、評價_DCF與目標價!C10、評價_DCF與目標價!D10、評價_DCF與目標價!E10、評價_DCF與目標價!F10、評價_DCF與目標價!G10
[PASS] xl_diff 公式 --by-label 0 差異

================ 結果 ================
PASS  check_quarterly：季度加總＝年度、差異原因齊全
PASS  README 欄位說明與 company.json 一致
PASS  Tokenomics 快照 --check：25 個名稱（missing 10）與快照一致（20261007_Tokenomics_v5.24.xlsx，相對誤差 ≤ 1e-09）
PASS  建 HTML
PASS  建 Excel
PASS  重算：0 公式錯誤
PASS  反向 DCF：Excel 求解，rv_snap.json 與 Excel 一致
PASS  fix_outline
PASS  fix_datatable
PASS  verify_ooxml：OOXML OK
PASS  快照值＝Excel 分頁值：25 個名稱（missing 10）、165 個值、55 個具名範圍一致；其他工作表無公式引用（v5.24）
PASS  離線開啟：新建 HTML
PASS  離線開啟：dist/ 成品
PASS  cmp31 low：313 項全部 OK
PASS  cmp31 base：313 項全部 OK
PASS  cmp31 high：313 項全部 OK
PASS  cmp31 base_FY27：313 項全部 OK
PASS  test_quarterly：假設實際數與可移植性測試通過
PASS  test_rolling：日曆推算與滾動後第一屏
PASS  test_attrib：拆解工具（月數、同版 0、滾動一季與 WACC 12%）
PASS  crawl 畫面文字 0 差異
PASS  xl_diff --values --by-label 0 差異
PASS  xl_diff 公式 --by-label 0 差異
verify.sh：全部通過（23 項）
EXIT 0
```
</details>


## W2 每 MW 改寫

chat 端審查 W1 後的決定（優先於工作單原文，2026-10-07）：收入採備案 `revenue=legacy`（GPU 小時路線實作、預設不啟用）；Tokenomics v5.26 未建置時成本暫代值（IT 維護＝`m.maint`、人員軟體與稅險＝0、`gpuLife`＝6）並在「檢查」頁警告；`meta.mwBasis=IT`（設施口徑換算實作並測試）；`fleet.openMix` 27／29／43%；`check_tokenomics_tab.py` 改為允許其他工作表引用；管銷率 5.51%（敏感度 9.24%）。

| 步驟 | 狀態 | commit | 備註 |
|---|---|---|---|
| 0 開分支、draft PR、進度檔 W2 段落、檢查 Tokenomics 版本 | 完成 | （本 commit） | 自 `origin/claude/coreweave-w1-tokenomics` 8d97234 開分支；Tokenomics master `bdb0de7`，`model/CURRENT`＝`20261007_Tokenomics_v5.24.xlsx`（無 v5.26）→ 快照維持 v5.24，v5.26 名稱以暫代值處理 |
