# CoreWeave × Tokenomics 改造進度（接手用；W1–W3 共用）

## 目前狀態（每次 push 前覆寫）
- 已完成：W0、W1（PR #6，待審）；**W2 全部完成**（PR #9，疊加於 W1 分支）：每 MW 資本支出（Tokenomics）、由下而上營運成本、收入備案（legacy）＋對照列、世代組合、每MW經濟性彙總表、敏感度快照；新方法 `verify.sh` 22 項全過；舊方法 `scripts/verify_legacy.sh`（--vs-dist）25 項全過、與 v4.5 成品 0 差異。
- 下一步：W3（`claude/coreweave-w3-v4.6`，疊加於 W2 分支，除非已合併）：v4.6 成品與前後對照報告；直接讀 Excel「每MW經濟性」彙總表；v4.5 側用 `scripts/verify_legacy.sh` 的副本（out/legacy_copy）或 dist/ 成品。
- 未解問題：(1) Tokenomics v5.26 尚未建置：`IF_MaintIT`、`IF_StaffSW`、`IF_TaxIns`、`IF_DeprLifeIT` 以暫代值（檢查頁警告）；建置後只需重抓快照、改 company.json `tokenomics`、verify；(2) GB200／GB300／VR200 長約 GPU 小時價格不足兩個獨立來源 → 收入採備案 legacy；(3) CoreWeave「active power」口徑定義未找到（預設 IT）；(4) 由下而上 EBITDA 率 70–74% 高於 Q2 實際 58.6%（收入每 MW 10.0 vs 8.2、租金每 MW 1.6 vs 2.1；人員軟體與稅險暫代 0），需 Andy 決定是否進 v4.6。

## 工作單總覽
| 工作單 | 分支 | PR | 狀態 |
|---|---|---|---|
| W0 遷移 | `claude/coreweave-w0-migrate` | | chat 端完成 |
| W1 Tokenomics 取數層 | `claude/coreweave-w1-tokenomics`（疊加於 W0 分支） | #6 | 完成，待審 |
| W2 每 MW 改寫 | `claude/coreweave-w2-permw`（疊加於 W1 分支） | #9 | 完成，待審 |
| W3 v4.6 成品與對照 | `claude/coreweave-w3-v4.6`（疊加於 W2 分支） | （開 PR 中） | 進行中 |

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
| 0 開分支、draft PR、進度檔 W2 段落、檢查 Tokenomics 版本 | 完成 | d047fa8 | 自 `origin/claude/coreweave-w1-tokenomics` 8d97234 開分支；Tokenomics master `bdb0de7`，`model/CURRENT`＝`20261007_Tokenomics_v5.24.xlsx`（無 v5.26）→ 快照維持 v5.24，v5.26 名稱以暫代值處理 |
| 1 世代組合 `fleet` | 完成 | 3b63e34 | openMix＝W1 年份分層精確值（27.27／29.37／43.37%；W1 回報 27／29／43% 為四捨五入）、newMix 工作單預設、newMixAlt＝FY29–30 Rubin Ultra；`fleet.openMix` 列入 ROLL_FIELDS／asOf；汰換由最舊世代先出 |
| 2 每 MW 資本支出（`capex=tokenomics`） | 完成 | 3b63e34 | 每 MW 建置成本＝Σ 新增世代占比 × IF_CapexIT：34.76、37.52、37.59、37.59、37.59（舊值 34 留為對照列）；不含廠房 |
| 3 折舊年限 | 完成（暫代） | 3b63e34 | IF_DeprLifeIT 缺 → `gpuLife` 維持 6（檢查頁警告）；有名稱時＝期初在役世代加權、取整數 |
| 4 由下而上營運成本（`cost=bottomUp`） | 完成（部分暫代） | 3b63e34 | 電費 IF_PowerCost；IT 維護暫代 m.maint；人員軟體、稅險暫代 0；管銷率 5.51%（Excel 由 10-Q 輸入格算出）；overlay 自動停用；最近一季實際對照列 |
| 5 每 MW 收入（備案 `revenue=legacy`） | 完成 | 3b63e34 | GPU 小時路線實作（pricing.gpuHr 空白＝「不適用」）；對照列：IF_HoldEcon、IF_GPUhrEcon 倍數、隱含每 GPU 小時價格、IREN 9.7、W1 各世代市場價格、v4.5 舊值、期末 ARR ÷ MW |
| 6 MW 口徑 | 完成 | （本 commit） | `meta.mwBasis=IT`；設施口徑換算以 `scripts/test_permw.py` 測試 A 驗證（÷ 1.2） |
| 7 每 MW 經濟性彙總表 | 完成 | 3b63e34 | Excel「每MW經濟性」頁最上方；HTML「資金模型 → 運營活動 → 每 MW 經濟性」（新方法時才顯示）；cmp31 逐列比對 |
| 8 敏感度 | 完成 | （本 commit） | `scripts/permw_sens.py`（Excel 求值快照，verify 步驟 3c）＋快照狀態格；HTML 即時計算、cmp31 比對；四組：Tokenomics 低／高成本、GPU 小時價格（不適用）、Rubin Ultra 版、管銷率 GAAP 9.24% |
| 9 整體 verify 與回報 | 完成 | （本 commit） | 新方法 `verify.sh` 22 項全過（cmp31 三情境各 426 項、FY27 錨定 405 項）；舊方法 `scripts/verify_legacy.sh`（--vs-dist）全過、畫面文字與 Excel 值／公式 0 差異 |

### W2 關鍵數字（v4.5 → W2 新方法；基準情境除另註）

| 情境 | v4.5 加權目標價（評等） | 新（評等） | 融資缺口 v4.5 → 新（US$bn） |
|---|---|---|---|
| 保守 4.2 GW | $69.74（賣出） | $106.92（中立） | 19.7 → 12.3 |
| 基準 5.6 GW | $45.50（賣出） | $66.70（賣出） | 61.7 → 59.1 |
| 積極 8 GW | $30.95（賣出） | $44.18（賣出） | 126.2 → 130.6 |

情境區間：$31.0–$69.7 → $44.2–$106.9；方法區間（5–6x）：$29.5–$45.5 → $47.4–$66.7。

| 每 MW（US$m／MW／年，基準情境；v4.5 → 新） | FY26（2H） | FY27 | FY28 | FY29 | FY30 |
|---|---|---|---|---|---|
| 年收入（算力＋服務） | 10.05 → 10.05 | 10.56 → 10.56 | 10.33 → 10.33 | 9.63 → 9.63 | 8.65 → 8.65 |
| 電費 | 0.67 → 0.67 | 0.67 → 0.67 | 0.67 → 0.67 | 0.67 → 0.67 | 0.67 → 0.67 |
| IT 維護（暫代） | 0.15 → 0.15 | 0.15 → 0.15 | 0.16 → 0.16 | 0.16 → 0.16 | 0.17 → 0.17 |
| 人員、軟體、水與耗材（暫代 0） | 0.00 → 0.00 | 0.00 → 0.00 | 0.00 → 0.00 | 0.00 → 0.00 | 0.00 → 0.00 |
| 財產稅與保險（暫代 0） | 0.00 → 0.00 | 0.00 → 0.00 | 0.00 → 0.00 | 0.00 → 0.00 | 0.00 → 0.00 |
| 公司管銷 | 0.55 → 0.55 | 0.58 → 0.58 | 0.57 → 0.57 | 0.53 → 0.53 | 0.48 → 0.48 |
| 租金 | 1.60 → 1.60 | 1.59 → 1.59 | 1.33 → 1.33 | 1.16 → 1.16 | 1.00 → 1.00 |
| 現金成本合計（含租金） | 4.12 → 2.98 | 4.17 → 3.00 | 3.93 → 2.73 | 3.51 → 2.53 | 3.03 → 2.32 |
| EBITDA | 5.93 → 7.07 | 6.39 → 7.56 | 6.40 → 7.60 | 6.11 → 7.10 | 5.62 → 6.33 |
| D&A | 5.69 → 5.73 | 5.84 → 6.03 | 5.66 → 5.96 | 5.64 → 5.99 | 5.59 → 5.97 |
| 利息 | 2.83 → 2.82 | 2.64 → 2.66 | 2.14 → 2.12 | 1.73 → 1.73 | 1.54 → 1.54 |
| 稅前 | -2.59 → -1.48 | -2.10 → -1.13 | -1.40 → -0.49 | -1.25 → -0.62 | -1.50 → -1.18 |
| 每 MW 資本支出（US$m／新增 MW） | 34.00 → 34.76 | 34.00 → 37.52 | 34.00 → 37.59 | 34.00 → 37.59 | 34.00 → 37.59 |
| EBITDA 率 | 59.0% → 70.4% | 60.5% → 71.6% | 62.0% → 73.5% | 63.5% → 73.7% | 65.0% → 73.2% |

v4.5 未拆項：v4.5 欄的電費、IT 維護、管銷是同一組由下而上對照值（不入 v4.5 損益），v4.5 的現金成本合計＝收入 − EBITDA（隱含，其中非租金部分 2.52、2.58、2.60、2.35、2.03）。新方法的非租金現金成本 1.38、1.41、1.40、1.37、1.32。

**租金只扣一次（FY27 基準）**：EBITDA 率 71.62%＝由下而上 EBITDAR 率 86.69% − 租金 3.939 ÷ 營收 26.127（15.08%）；資金端現金利潤率＝EBITDA 率＋租金 ÷ 營收＝86.69%，租金 3.939 只在支出端扣一次；現金 EBITDA 18.602＋信用調整 0.109＝損益 EBITDA 18.711。

**最近一季實際對照（不強制平衡）**：Q2 每 MW 現金營運成本（租金前、不含管銷）0.877 vs 模型 FY26 0.823（−0.054）；Q2 每 MW 年租金 2.08 vs 模型 1.60；Q2 每 MW 年收入 8.24（2.575 × 4 ÷ 1,250 MW）vs 模型 10.05。由下而上 EBITDA 率高於 Q2 實際 58.6%，主要來自收入（Q2 分母含爬坡中尚未計費的 MW）與租金（模型路徑每 MW 低於 Q2），營運成本本身接近。

**敏感度（基準情境加權目標價，建置時快照）**：Tokenomics 低成本 $120.57／高成本 $34.04；Rubin Ultra 版 $54.92；管銷率 GAAP 9.24% $53.21；GPU 小時價格低／高：不適用（收入備案）。

**收入對照（基準，FY26→FY30）**：隱含每 GPU 小時價格 $2.34→$3.11 對 IF_GPUhrEcon（GPU 加權）$2.31→$3.65，倍數 1.01→0.85；每 MW 年收入 ÷ IF_HoldEcon 1.01→0.85；同業 IREN–Microsoft 9.7 US$m/MW/年（模型 100% 計費時數 11.2→10.5）。

### W2 已套用的預設（問題｜採用的預設｜替代選項｜對結果的影響）
| 問題 | 預設 | 替代 | 影響 |
|---|---|---|---|
| 方法開關的實作 | 建置時依 `methodology.perMw` 產生公式（Excel 內不設即時切換格） | Excel 加方法選擇格、IF 切換 | 無數字影響；舊方法組合可證明 0 差異 |
| revMW 是否含利用率 | revMW＝100% 計費時數的每 MW 收入（引擎在收入端另乘 `m.util`）；隱含 GPU 小時價格＝revMW ÷（GPU 數 × 8,760），GPU 路線公式也不乘利用率 | 依工作單字面在兩處都乘／除利用率 | 避免利用率重複計算；隱含價格為「每計費小時」 |
| 世代結構起點 | 首期（2H26）自 6/30 在役 1,500 MW 起算；CapEx 仍依 YE25 850 MW 的全年新增 | 自 YE25 850 MW 起算 | 只影響世代加權（每 MW 成本、對照列） |
| 每 MW 成本的 MW 分母 | 平均在役 Accepted MW（公司口徑） | Billable MW | 改 Billable 時由下而上成本約 −8%、EBITDA 率約 +1pt |
| 汰換的世代 | 汰換批次由世代清單中最舊者先出、以當期新增世代補回 | 依實際年份批次的世代 | FY29–30 汰換 Hopper（100、260 MW） |
| 建置成本的世代 | 當期新增與次年預建（λ）都用當期 newMix | 預建部分用次期 newMix | FY27 預建差 <0.2% |
| FY26 每 MW 建置成本 | 全年新增 1,000 MW 都用 2H26 newMix（34.76） | 1H 650 MW 用 W1 的 2026 上半年組合 | FY26 CapEx 下限 35 未觸發，差 <1% |
| 穩態 EBITDA 率（bottomUp） | 預設＝由下而上 FY30（公式）；輸入數值視為 FY30 目標、差額線性分攤（反向 DCF Eb、敏感度 59%／70% 沿用） | 停用該格 | 反向 DCF 與敏感度仍可用；預設無差額 |
| 管銷率 | 5.51%（扣 SBC；S&M／G&A 內 D&A 視為 0），Excel 由 10-Q 輸入格算出 | 9.24%（GAAP） | 敏感度：基準目標價 $66.70 → $53.21 |
| 最近一季實際對照 | D&A 全數視為在營收成本與技術基礎設施；租金＝營業 0.500＋變動 0.150（不含融資租賃） | 含融資租賃 0.017 | 每 MW 現金成本 ±0.05 |
| 稅險 × IT 資本占比 | 各世代 IF_TaxIns × IF_CapexIT ÷ IF_CapexTotal，再依在役世代加權 | 單一占比 | v5.26 前暫代 0 |
| 持有成本加權 | IF_HoldEcon 依 MW 加權；IF_GPUhrEcon 依 GPU 數加權 | 都依 MW | 只影響對照列 |
| 設施口徑換算 | IF_FacilityGW 取基準（不隨成本情境變動） | 隨成本情境 | 只在 facility 口徑時有影響 |
| 敏感度第 4 組 | 加管銷率 GAAP 9.24%（chat 端決定 6）；GPU 小時價格低／高在收入備案下為「不適用」 | — | 見上 |
| 差異原因 | `varianceReasons` 加 `perMw` 條件；新方法下 FY26 EBITDA、FY26 CapEx 超過門檻，新增「觀點」原因 | — | 建置檢查通過（無「未歸類」） |
| 新工作表位置 | Excel「每MW經濟性」放在「運營_站點」之後；世代結構推導放在此頁（輸入在「輸入與假設」B 區後） | 推導也放輸入頁 | 無數字影響 |

### W2 可移植性
- 新增公司特有內容：`fleet`（世代組合）、`pricing`（GPU 小時價格、同業與市場對照）、`costs.sgaBasis`、`latestQuarter` 費用欄（sm、smSbc、ga、gaSbc、sbcCostTi、opLeaseCost、varLeaseCost、finLeaseCost）、`meta.mwBasis`、`methodology.perMw`、`varianceReasons` 的 `perMw` 條件。換公司時：填 fleet（或刪除＝每 MW 經濟性不適用，方法須為 legacy）、pricing、latestQuarter 費用。
- 沒有 fleet 或 Tokenomics 時：Excel 不建「每MW經濟性」頁，HTML 不顯示該分頁；新方法在建置時要求 fleet 與必要的 Tokenomics 名稱（缺少即報錯）。

### W2 下一張（W3）需要知道的事
- 前後對照：v4.5 側＝`scripts/verify_legacy.sh` 產生的副本（out/legacy_copy）或 dist/ 成品；新側＝目前 company.json。彙總表在 Excel「每MW經濟性」第 6–23 列（列名固定）。
- 目標價變動全屬 (d) 方法變更；拆解可依序切換 capex → cost（`methodology.perMw`）各重建一次。
- v5.26 名稱補齊後成本會上升（人員軟體、稅險目前為 0），由下而上 EBITDA 率會下降。

### W2 verify.sh 完整輸出（新方法，2026-10-07 最終）
<details><summary>展開</summary>

```
================ 結果 ================
PASS  check_quarterly：季度加總＝年度、差異原因齊全
PASS  README 欄位說明與 company.json 一致
PASS  Tokenomics 快照 --check：25 個名稱（missing 10）與快照一致（20261007_Tokenomics_v5.24.xlsx，相對誤差 ≤ 1e-09）
PASS  建 HTML
PASS  建 Excel
PASS  重算：0 公式錯誤
PASS  反向 DCF：Excel 求解，rv_snap.json 與 Excel 一致
PASS  每 MW 敏感度快照：Excel 求值，permw_sens.json 與 Excel 一致
PASS  fix_outline
PASS  fix_datatable
PASS  verify_ooxml：OOXML OK
PASS  快照值＝Excel 分頁值：25 個名稱（missing 10）、165 個值、55 個具名範圍一致；其他工作表 35 格公式引用 35 個 TK_ 名稱（v5.24）
PASS  離線開啟：新建 HTML
PASS  離線開啟：dist/ 成品
PASS  cmp31 low：426 項全部 OK
PASS  cmp31 base：426 項全部 OK
PASS  cmp31 high：426 項全部 OK
PASS  cmp31 base_FY27：405 項全部 OK
PASS  test_quarterly：假設實際數與可移植性測試通過
PASS  test_rolling：日曆推算與滾動後第一屏
PASS  test_attrib：拆解工具（月數、同版 0、滾動一季與 WACC 12%）
PASS  test_permw：設施口徑換算與 GPU 小時價格路線
verify.sh：全部通過（22 項）
EXIT 0
```
</details>


## W3 v4.6 成品與前後對照

chat 端追加（優先於工作單，2026-10-07）：第 0 步把 Tokenomics 快照換成 v5.26（master `4074684`，`model/CURRENT`＝`20261007_Tokenomics_v5.26.xlsx`），補齊 W2 暫代的 `IF_MaintIT`、`IF_StaffSW`、`IF_TaxIns`、`IF_DeprLifeIT`；報告須含 FY26 每 MW 對 Q2 2026 實際的逐項對帳。

| 步驟 | 狀態 | commit | 備註 |
|---|---|---|---|
| 0 開分支、draft PR、進度檔 W3 段落 | 完成 | （本 commit） | W0 #5、W1 #6、W2 #9 皆未合併：自 `origin/claude/coreweave-w2-permw` 9af51ca 開分支，PR base＝W2 分支 |
