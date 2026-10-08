# CoreWeave × Tokenomics 改造進度（接手用；W1–W3 共用）

## 目前狀態（每次 push 前覆寫）
- 已完成：W0–W5（v4.7，已合併）；**W6 進行中**（PR #48，分支 `claude/coreweave-w6-v4.8`）：第 0–4 步完成（快照 v5.31、IT 維護依機齡兩段、公司實況驗證重跑、MW 口徑查證＝無原文定義、維持 IT）。
- 下一步：W6 第 5 步——升版 v4.8（VLOG 兩處、dist、交接檔、README）；`scripts/expect/v4_8_rules.json` → `make_expect.py`（對 v4.7 成品）→ `EXPECT=… verify.sh --vs-dist`；`verify_legacy.sh`（對 v4.6）；`scripts/attrib_w6.py --v47 dist v4.7`；第 6 步 `scripts/compare_w6.py`＋報告 md。
- 未解問題（W6）：見 W6 段落。

## 工作單總覽
| 工作單 | 分支 | PR | 狀態 |
|---|---|---|---|
| W0 遷移 | `claude/coreweave-w0-migrate` | | chat 端完成 |
| W1 Tokenomics 取數層 | `claude/coreweave-w1-tokenomics`（疊加於 W0 分支） | #6 | 完成，待審 |
| W2 每 MW 改寫 | `claude/coreweave-w2-permw`（疊加於 W1 分支） | #9 | 完成，待審 |
| W3 v4.6 成品與對照 | `claude/coreweave-w3-v4.6`（疊加於 W2 分支） | #18 | 已合併 |
| W4 收入 Tokenomics 錨定（v4.7） | `claude/coreweave-w4-revenue`（自 main f373885） | #27 | 完成，待審 |
| W5 公司實況驗證（覆寫 v4.7） | 同 W4 分支 | #27 | 已合併 |
| W6 接 Tokenomics v5.31（v4.8） | `claude/coreweave-w6-v4.8`（自 main 8671a11） | （見下） | 進行中 |

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
| 7 整體 verify 與回報 | 完成 | fcf0865 | `verify.sh --vs-dist` 23 項全過；PR 留言「[CRWV 回報] W1｜完成｜2026-10-07」 |

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
| 0 開分支、draft PR #18、進度檔 W3 段落 | 完成 | df61c82 | W0 #5、W1 #6、W2 #9 皆未合併：自 `origin/claude/coreweave-w2-permw` 9af51ca 開分支，PR base＝W2 分支 |
| 0′ Tokenomics 快照換 v5.26（chat 端追加） | 完成 | b32828a | 唯讀副本 `git fetch origin master` → `4074684`；`import_tokenomics.py` 重抓 `data/tokenomics_snapshot_v5.26.json`（25 名、missing 0），刪 v5.24 快照；名稱清單移除 optional；company.json `tokenomics`（v5.26、commit 4074684、optional 空）；`fields_doc.py --write`。原 15 名數值與儲存格位置與 v5.24 完全相同。檢查頁「名稱缺漏」警告消失。新方法 `verify.sh` 22 項全過；舊方法 `verify_legacy.sh` 25 項全過（與 v4.5 成品 0 差異） |
| 1 升版 v4.6 | 完成 | b10381e | `vlog.py`、`tail.js` VLOG 新增 v4.6（10-08；含 (a)–(d) 拆解）；`dist/20261008_CoreWeave收支模型_v4_6.{html,xlsx}`，移除 v4.5；交接檔換成 `docs/handoff/20261008_CoreWeave收支模型_交接檔_v4_6.md`（新增 2i v4.6、2j Tokenomics 連結：快照版本、引用名稱、升版步驟）；README v4.6 段落與工具表 |
| 2 升版驗收 | 完成 | b10381e | `DATE=2026-10-08 EXPECT=scripts/expect/v4_6_vs_v4_5.txt scripts/verify.sh --vs-dist`（對 v4.5 成品）25 項全過：預期差異 687 格＋5 列改名（`scripts/make_expect.py` 依 `scripts/expect/v4_6_rules.json` 產生，未歸類 0），其餘 0 差異；畫面文字為升版預期差異（無頁面錯誤、無缺頁）。舊方法組合 `verify_legacy.sh` 對 v4.5 成品 25 項全過、0 差異（副本 dist/ 改取 git 歷史的 v4.5 成品） |
| 3 目標價變動拆解 | 完成 | b10381e | `scripts/attrib_permw.py`：三情境 (a)(b)(c)＝0，(d) ① 每 MW 資本支出／② 折舊年限／③ 營運成本／④ 收入／⑤ MW 口徑依序與單獨切換；各步相加＝總變動（誤差 < 0.01）；結果 `out/w3/attrib_permw.json` → 對照 Excel「變動拆解」 |
| 6 HTML／Excel 成品畫面 | 完成 | b10381e | 一頁摘要「結論」加一句每 MW（HTML `pmLineQ`＝Excel「摘要」→「結論｜每 MW 經濟性句」，cmp31 逐字比對；舊方法不顯示）；「每 MW 經濟性」分頁彙總表放次層（預設收合）。新方法 `verify.sh` 22 項全過（cmp31 三情境各 427 項、FY27 錨定 406 項） |

| 4 前後對照 Excel | 完成 | （本 commit） | `docs/reports/20261007_coreweave_v4.6_前後對照.xlsx`（摘要＋Q2 實際對帳、每MW_前後〔三情境 × FY26–30，IT 與設施口徑〕、參數對照、變動拆解、Tokenomics參考線、已知限制）；`scripts/compare_gather.py` 檢查舊方法副本與 v4.5 成品 5,439 個共有數值格相同；`scripts/build_compare.py`；LibreOffice 重算 601 個公式 0 錯誤；拆解檢查格三情境「通過」 |
| 5 報告 md | 完成 | （本 commit） | `docs/reports/20261007_coreweave_v4.6_前後對照.md`（結論三句、關鍵數字、拆解、Q2 對帳、W2 暫代 vs v5.26、已套用預設、未解問題、verify 輸出） |
| 7 整體 verify 與回報 | 完成 | fcf0865 | 新方法 `verify.sh` 22 項全過；`--vs-dist`（對 v4.5，EXPECT）25 項全過；`verify_legacy.sh` 25 項全過；PR 留言「[CRWV 回報] W3｜完成｜2026-10-08」、PR 改 ready |
### W3 第 0′ 步：v5.26 正式值取代 W2 暫代值（基準情境）

核對 Tokenomics 值（$B/GW/年＝US$m/MW/年，基準）：GB300 IT 維護 1.1234、人員軟體 0.325、稅險 0.2506；VR200 IT 維護 1.1276、稅險 0.2513（與 chat 端核對值相同）；Hopper 0.8253／0.325／0.2009、GB200 0.7205／0.325／0.1834；IF_DeprLifeIT 基準各世代 6 年（高成本 4 年）→ 加權取整 6，與暫代值相同。
稅險口徑：各世代 IF_TaxIns × IF_CapexIT ÷ IF_CapexTotal（只算 CRWV 擁有的 IT 部分；GB300 0.2506 × 37.45 ÷ 50.12＝0.187），人員軟體全額（0.325）——W2 公式已如此實作，未改。

| 每 MW（US$m／MW／年，基準） | FY26（2H） | FY27 | FY28 | FY29 | FY30 |
|---|---|---|---|---|---|
| 年收入 | 10.05 | 10.56 | 10.33 | 9.63 | 8.65 |
| 電費 | 0.67 | 0.67 | 0.67 | 0.67 | 0.67 |
| IT 維護（W2 暫代 → v5.26） | 0.15 → 0.94 | 0.15 → 0.99 | 0.16 → 1.03 | 0.16 → 1.05 | 0.17 → 1.08 |
| 人員軟體（0 → v5.26） | 0 → 0.33 | 0 → 0.33 | 0 → 0.33 | 0 → 0.33 | 0 → 0.33 |
| 稅險（IT 部分；0 → v5.26） | 0 → 0.16 | 0 → 0.17 | 0 → 0.17 | 0 → 0.18 | 0 → 0.18 |
| 管銷 | 0.55 | 0.58 | 0.57 | 0.53 | 0.48 |
| 租金 | 1.60 | 1.59 | 1.33 | 1.16 | 1.00 |
| 現金成本合計（含租金） | 2.98 → 4.24 | 3.00 → 4.33 | 2.73 → 4.10 | 2.53 → 3.92 | 2.32 → 3.73 |
| EBITDA | 7.07 → 5.80 | 7.56 → 6.23 | 7.60 → 6.23 | 7.10 → 5.71 | 6.33 → 4.92 |
| EBITDA 率 | 70.4% → 57.8% | 71.6% → 59.0% | 73.5% → 60.3% | 73.7% → 59.3% | 73.2% → 56.9% |

| 情境 | v4.5 加權目標價 | W2 暫代 | v5.26 正式 | 融資缺口 v4.5／W2／v5.26（US$bn） |
|---|---|---|---|---|
| 保守 4.2 GW | $69.74 | $106.92 | $43.99 | 19.7／12.3／32.7 |
| 基準 5.6 GW | $45.50 | $66.70 | $29.68 | 61.7／59.1／81.6 |
| 積極 8 GW | $30.95 | $44.18 | $19.95 | 126.2／130.6／157.9 |

讀法：v5.26 補上 IT 維護（約 1.0／MW，暫代 0.15 的 6–7 倍）、人員軟體 0.33、稅險 0.16–0.18 後，每 MW 現金成本增加約 1.3–1.4，EBITDA 率由 70–74% 降到 57–60%（Q2 實際 Adj. EBITDA 率 58.6%）；三情境目標價都低於 v4.5。

**verify 工具修正（不影響成品數字）**：(1) `xlx.py`：LibreOffice 把「輸入與假設」H 區模擬運算表轉成 MULTIPLE.OPERATIONS，批次重算非基準情境時部分格殘留運算表代入的中間值（v5.26 積極情境實測：D&A 車隊 FY27–28 取到基準情境值，cmp31 74 項不一致；以 UNO 逐格重算與移除運算表兩種方式確認 Excel 公式本身正確）。改為：未改錨定年度時，運算表輸出（三情境加權目標價與評等，與情境選擇無關）以來源檔快取值取代後再重算。(2) `cmp31.js`：DCF 失效（常態化 FCF ≤ 0）時 HTML 每股為 NaN（畫面顯示「失效」）、Excel 採用值為 0，比對時視為 0（v5.26 積極情境首次觸發；基準情境 DCF 0 截斷後也為 0）。

### W3 已套用的預設
見 `docs/reports/20261007_coreweave_v4.6_前後對照.md` 第 5 節（成品更新日 10-08、② 折舊年限獨立一步、(c)＝0、預期差異清單以規則產生、升版時畫面文字只查錯誤與缺頁、舊方法回歸改取 git 歷史 v4.5、LibreOffice 運算表凍結、DCF 失效比對、每 MW 句 FY27、持有成本倍數口徑、v4.5 每 MW 欄來源、Q2 對帳 D&A／SBC 口徑）。

### W3 關鍵數字（v4.5 → v4.6）
加權目標價：保守 $69.74 → $43.99、基準 $45.50 → $29.68、積極 $30.95 → $19.95（皆賣出）；融資缺口 19.7／61.7／126.2 → 32.7／81.6／157.9；基準累計新股 FY30 0.171 → 0.261bn 股。拆解（依序）：① −16.57／−7.62／−9.13、③ −9.18／−8.21／−1.88，②④⑤ 0。


## W4 每 MW 收入以 Tokenomics 為錨（v4.7）

工作單 r1（`docs/workorders/20261008_coreweave_W4_收入Tokenomics錨定.md`）。每 MW 年收入（100% 計費時數）＝Σ 平均在役占比 × `IF_HoldEcon`（TK_ 名稱，目前成本情境）× k；k＝長約占比 × k_長約＋（1 − 長約占比）× k_現貨；長約占比＝MIN(1, MAX(0, 排程 RPO ÷（平均計費 MW × 錨 × k_長約 × 利用率 × 期間長度）＋調整))。

| 步驟 | 狀態 | commit | 備註 |
|---|---|---|---|
| 0 開分支、draft PR #27、Tokenomics 版本 | 完成 | 0d00e96 | 自 origin/main f373885；Tokenomics master `862bdd4`，`model/CURRENT`＝`20261008_Tokenomics_v5.27.xlsx`（v5.26 → v5.27）。重抓快照 `data/tokenomics_snapshot_v5.27.json`：原 25 名數值與儲存格位置與 v5.26 **完全相同**；新增 `IF_RevGWFleet`（上限檢查）。刪 v5.26 快照（舊方法回歸自 git 歷史取回） |
| 1 證據補充 | 完成 | 686991d | 見下方「W4 第 1 步證據表」；`company.json` → `pricing.anchorMultiple`（evidence、contractMix、notFound） |
| 2 實作 tkAnchor | 完成（新方法 verify 22 項全過） | 686991d | Excel：輸入頁「定價倍數 k」三格＋TK 收入上限區塊；「每MW經濟性」新增「每 MW 收入：Tokenomics 錨 × 定價倍數 k（W4）」與「k 證據表」；B 區每 MW 年收入改引用；檢查頁「收入上限」；JS：`anchorRevQ`、perMwQ `tk`／`kev`、pmwSensQ 六組、檢查卡；cmp31 三情境各 472 項、FY27 錨定 433 項全 OK。另修 JS 既有缺陷：segB 無槓桿 NOL 虧損只加回 80%（Excel 為 100%；v4.6 前未觸發，W4 數字下 FY30 UFCF 差 0.41）。舊方法回歸 `scripts/verify_legacy.sh` 基準改為 v4.6（revenue=legacy、其餘 v4.6 設定、快照取 git 歷史 v5.26、dist 取 f373885）：25 項全過，畫面文字、Excel 值與公式對 v4.6 成品 **0 差異**（新增列 44） |
| 3 Q2 驗證（不校準） | 完成 | 686991d | `scripts/q2_check_w4.py`：模型首期 7.638 對 Q2 年化 8.240（−0.602，−7.3%）＝(i) 爬坡分母 +0.091、(ii) 利用率 0、(iii) 定價倍數 −0.737（模型 k 0.760 vs Q2 隱含 0.838）、(iv) 世代組合 +0.072、(v) 其他 −0.028；相加＝總差距。Excel「每MW經濟性」最近一季實際對照新增 5 列（Q2 每 MW 收入 ÷ 在役／÷ 計費 MW、Q2 錨、Q2 隱含 k 未調整 0.752） |
| 4 敏感度 | 完成 | fcf0865 | `permw_sens.json` 新增 k_長約 0.70／1.00、k_現貨 1.50／2.30、長約占比 −20pt、長約占比 100% 六組（HTML 即時、cmp31 逐格）；基準：$59.24／$236.10、$54.16／$188.00、$201.31、$5.25；Tokenomics 低／高 $0.21／$605.66；Rubin Ultra 版 $104.23 |
| 5 升版 v4.7 | 完成 | fcf0865 | VLOG 兩處、dist 換 v4.7（移除 v4.6）、交接檔 v4.7（2k、2j、第 5 節）、README v4.7 段落與工具表；`scripts/expect/v4_7_rules.json` → `make_expect.py`（1,022 格＋4 列改名，未歸類 0；營運成本列不在清單）→ `EXPECT=… verify.sh --vs-dist` 對 v4.6 成品 25 項全過；`scripts/attrib_w4.py`：(d) ① 錨（k＝1）+30.16／+15.46／+8.82、② 套用 k +17.09／+52.88／+103.09、③ 0，相加＝總變動 |
| 6 對照報告 | 完成 | fcf0865 | `scripts/compare_w4.py` → `docs/reports/20261008_coreweave_v4.7_收入錨定.xlsx`（摘要、每MW_前後、錨與k、證據表、Q2驗證、敏感度、變動拆解；LibreOffice 重算 255 公式 0 錯誤，檢查格皆「通過」）＋同名 md |
| 7 整體 verify 與回報 | 完成 | fcf0865 | 見報告 md 第 10 節；PR 留言「[CRWV 回報] W4｜完成｜2026-10-08」 |

### W4 r2（chat 端審查後修正，2026-10-08；覆寫 v4.7）
| 項目 | 狀態 | commit | 備註 |
|---|---|---|---|
| 工作單改 r2 | 完成 | 1716228 | r1 修訂說明下方加 r2 一段（隨需占比、k 隨成本情境重算、三筆中位數、DCF 失效限制） |
| 隨需占比取代長約占比 | 完成 | 1716228 | k＝隨需占比 × k_現貨＋（1 − 隨需占比）× k_長約；`pricing.anchorMultiple.onDemandShare` 0%（敏感度 10%／20%，[Assumed]）；RPO 覆蓋率改為對照列 |
| k 隨成本情境重算 | 完成 | 1716228 | 輸入值為基準情境；k_長約 × IF_HoldEcon｜GB300（基準）÷ 目前情境、k_現貨 × IF_GPUhrEcon｜H100 同理；低／高成本下收入大致不變（FY30 10.1／8.5 vs 9.4） |
| 敏感度 | 完成 | 1716228 | k_長約 0.70／0.89（三筆中位數）／1.00、隨需占比 10%／20%；移除長約占比與 k_現貨 兩組 |
| 核對工具 | 完成 | 1716228 | `xlx.py` 運算表一律逐情境另存重算（LibreOffice MULTIPLE.OPERATIONS 會使累計新股 FY29–30 等格殘留代入值，連 test_permw 的未修正檔也受影響）；`fix_datatable.py` 設 fullCalcOnLoad；`rv_solve.py` 無解寫 null |
| v4.7 覆寫與驗收 | 完成 | （本輪最後 commit，見 PR） | VLOG 兩處、dist、交接檔、README、對照 Excel／md；升版驗收對 v4.6 25 項全過（清單 978 格＋4 列改名）；新方法 --vs-dist 25 項全過；舊方法回歸 25 項全過、0 差異 |

**r2 關鍵數字（v4.6 → v4.7 r2）**：加權目標價 保守 $43.99 → $13.98、基準 $29.68 → $5.25、積極 $19.95 → $0.00（皆賣出）；融資缺口 32.7／81.6／157.9 → 54.8／104.8／183.5。EV/EBITDA 腿單獨 $25.42／$5.25／$0.00；DCF 腿 0 截斷／失效／失效。基準每 MW 年收入（100% 計費）8.42／8.79／9.06／9.21／9.35（k 各期 0.76）。拆解 (d) ① +30.16／+15.46／+8.82、② −60.17／−39.89／−28.77、③ 0。

（以下「W4 關鍵數字」為 r1 結果，已被 r2 取代，保留紀錄）

### W4 關鍵數字（v4.6 → v4.7）
加權目標價：保守 $43.99 → $91.24、基準 $29.68 → $98.02、積極 $19.95 → $131.86（皆由賣出轉中立）；融資缺口 32.7／81.6／157.9 → 30.3／67.5／118.1。基準每 MW 年收入（100% 計費）11.2／11.5／11.5／11.0／10.5 → 8.42／8.79／10.13／12.70／16.47；錨 11.07–12.31；k 0.76、0.76、0.85、1.05、1.34；長約占比 100%、100%、91%、71%、42%；上限檢查最高 42%（門檻 50%）。

### W4 已套用的預設
見 `docs/reports/20261008_coreweave_v4.7_收入錨定.md` 第 7 節（長約占比逐期公式、第二筆長約不改基準、證據倍數口徑、k 不隨成本情境、上限檢查分子、Q2 計費比例／利用率／服務／世代組合假設、Tokenomics v5.27、舊方法回歸基準、JS NOL 修正、摘要每 MW 句、差異原因）。

### W4 第 1 步證據表（倍數＝價格 ÷ Tokenomics v5.27 基準同世代持有成本；Excel 以 TK_ 名稱計算）
| 證據 | 世代 | 價格 | Tokenomics | 倍數 | 用途 | 合約／期間 | 來源（日期） | 標記 |
|---|---|---|---|---|---|---|---|---|
| IREN–Microsoft 五年約 | GB300 | 9.70 US$m/MW/年 | IF_HoldEcon 12.725 | 0.76 | k_長約基準 | 長約 5 年 | IREN 新聞稿（2025-11-03） | [Derived] |
| IREN–NVIDIA 氣冷 Blackwell 五年約 | GB300 類比 | 11.33 US$m/MW/年（$3.4bn ÷ 5 ÷ 60 MW） | IF_HoldEcon 12.725 | 0.89 | 支持區間 | 長約 5 年 | IREN 新聞稿（2026-05-07） | [Analogy] |
| Oracle–OpenAI（未經證實） | GB300 類比 | 13.33 US$m/MW/年（$300bn ÷ 5 ÷ 4.5 GW） | IF_HoldEcon 12.725 | 1.05 | 只列 | 約 5 年 | The Register 轉述 WSJ（2025-09-11） | [Analogy] |
| H100 Silicon Data 指數 | H100 | 2.82 US$/GPU-hr | IF_GPUhrEcon 1.605 | 1.76 | k_現貨基準 | 現貨 | Silicon Data（2026-10-06） | [Verified] |
| B300 Silicon Data 指數 | GB300 類比 | 6.82 | 2.983 | 2.29 | 支持區間（上緣） | 現貨 | Silicon Data（2026-10-06） | [Analogy] |
| B200 Ornn 成交指數 | GB200 類比 | 4.08 | 2.086 | 1.96 | 支持區間 | 現貨 | Ornn 經 WSJ（2026-04-13） | [Analogy] |
| B200 Silicon Data 指數 | GB200 類比 | 5.87 | 2.086 | 2.81 | 只列（高於區間） | 現貨 | Silicon Data（2026-10-06） | [Analogy] |
| B300 Nebius 隨需牌價 | GB300 類比 | 7.85 | 2.983 | 2.63 | 只列 | 牌價 | Nebius 價目（2026-10-07） | [Interested-party] |
| GB200 CoreWeave 隨需牌價 | GB200 | 10.50 | 2.086 | 5.03 | 只列 | 牌價 | CoreWeave 價目（2026-10-07） | [Interested-party] |
| H100 CoreWeave 隨需牌價 | H100 | 6.155 | 1.605 | 3.83 | 只列 | 牌價 | CoreWeave 價目（2026-10-07） | [Interested-party] |

- 最終預設：k_長約 0.76（0.70–1.00）、k_現貨 1.76（1.5–2.3），皆 [Analogy]；第二筆長約（IREN–NVIDIA 0.89）與第一筆同一供應商，依保守原則只支持區間、不改基準。
- CRWV 合約組合：已承諾合約加權平均年期約 5 年（10-K FY2025 [Verified]）；「絕大多數收入來自多年期已承諾合約」（S-1，未給百分比 [Interested-party]）；長約／隨需比例 **找不到**（試過 10-K FY2025、S-1、Q2 新聞稿與搜尋）。
- 找不到：VR200 長約價（Nebius–Meta $12bn Vera Rubin、Nebius–Microsoft $17.4–19.4bn 皆未揭露 MW 或 GPU 數）；GB200 長約價；獨立於 IREN 的第二家供應商長約（Nscale–Microsoft 未揭露金額，FT 估計不用）。未使用 SemiAnalysis。

## W5 公司實況驗證（覆寫 v4.7；工作單 `docs/workorders/20261008_coreweave_W5_公司實況驗證.md`）

| 步驟 | 狀態 | commit | 備註 |
|---|---|---|---|
| 0 讀工作單與前置 | 完成 | — | 自 origin/claude/coreweave-w4-revenue 0199e77 繼續；環境 verify --vs-dist（W4 狀態）全過，約 8 分鐘 |
| 1 逐項查證 | 完成 | 035eeb1 | 證據寫入 `company.json` → `companyAdjust.evidence`（15 筆）與 `notFound`（4 項）；k 證據表新增 CRWV Q3 短天期合約（$40m/MW，[Interested-party]，只列） |
| 2 實作 | 完成 | 035eeb1 | 見下方「W5 實作」；cmp31 三情境各 528 項、FY27 錨定 474 項 OK；test_rolling、test_permw 通過 |
| 3 Q2 逐項對帳 | 完成 | e155432 | 收入 8.22 對 8.24（+0.3%）、固定租金 1.60 對 1.60、管銷 0.45 對 0.45、營運成本 2.09 對 1.36（−35%，規則 3）、EBITDA 率 49.6% 對 58.6% |
| 4 v4.7 覆寫、報告、交接檔、README | 完成 | e155432 | VLOG 兩處、dist、`scripts/expect/v4_7_rules.json`（＋營運成本四列公式 × 倍數）→ 980 格＋4 列改名、未歸類 0；`attrib_w5.py`、`compare_w5.py`；報告 `20261008_coreweave_v4.7_公司實況驗證.{md,xlsx}` |
| 5 verify 三組與 PR 回報 | 完成 | （本輪最後 commit） | 見下方「W5 關鍵數字」與報告第 11 節 |

### W5 第 1 步證據（SEC EDGAR 原文優先；WebFetch 讀 10-Q／10-K 的 XBRL R 頁）
| 參數 | CRWV 實際 | 來源（日期） | 標記 | 結論 |
|---|---|---|---|---|
| 電費 | 未單獨揭露（不存在於列項）；10-K：營收成本含 rent、utilities including power；變動租賃＝共同區域維護、轉付房東的水電、保全；Core Scientific：「Client pays for capex, power, and utilities」；電力交換名目 $104m 認列於營收成本 | 10-K FY2025 R28、10-Q Q2 2026 附註 8、CORZ Q2 FY26 簡報（2026-07-28） | [Verified]／[Interested-party] | CRWV 自付電費；模型租金只含固定租金 → 無重複計算；單項無法量測 |
| IT 維護 | 未單獨揭露；保固條款**找不到**（10-K、10-Q、S-1）；10-K：技術與基礎設施含「維護運算基礎設施的人員」；Supermicro 標準保固 GPU 系統人工 3 年／零件 1 年（廠商通則） | 10-K R28；Supermicro 保固頁 | [Verified]／[Analogy] | 保固機制未證實 → 規則 3 |
| 人員軟體 | 未單獨揭露；技術與基礎設施含研發與基礎設施人員、軟體訂閱；Q2 研發 $117m；員工 2,189（2025 年底） | 10-K、10-Q | [Verified] | 模型管銷只含 S&M＋G&A → 無重複 |
| 稅險 | **找不到** | — | — | 併入合計 |
| 營運成本合計（租金前、不含管銷） | Q2：營收成本＋技術與基礎設施 − D&A − SBC − 營業租賃（**含變動租賃**）＝0.424bn／季 → 每 MW 年 1.357（W3／W4 的 0.88 把變動租賃＝水電誤扣） | 10-Q Q2 2026 | [Derived] | 對模型首期 2.09：−35%，找不到機制 → 規則 3 |
| 租金 | Q2 營業租賃成本 0.500 → 每 MW 年 1.60（固定）；變動租賃 0.150 歸營運成本 | 10-Q 附註 8 | [Verified] | 對模型 1.60：0% → 規則 1 |
| 每 MW 資本支出 | 1H26 資本支出 16.139 ÷ 新增 650 MW＝24.8；技術設備毛額 20.903 → 33.823 ÷ 650＝19.9 | 10-Q 附註 5；10-K／法說 MW | [Verified]／[Derived] | 對模型首期 34.8：−29%（−43%），找不到機制 → 規則 3 |
| 折舊年限 | 技術設備 6 年 | 10-K 會計政策（edgar.tools 轉載） | [Verified] | 一致 → 規則 1 |
| k 既有合約 | Q2 營收 2.575 → k_既有 0.828（扣服務、÷ 計費比例 90% × 利用率 95% × Q2 錨 10.96） | 10-Q | [Derived] | 對市場長約 0.76：+8.9%；依工作單第 4(a) 項套用於最新季末在役 1,500 MW |
| k 新合約 | 新長約每 MW 價格**找不到**；「7 月各 SKU 漲價約 25%」只有轉述（Trefis 2026-08-13）；Q3 已簽 3–6 個月短天期合約約 $40m/MW 年化（公司新聞稿 2026-09-16） | Trefis、CoreWeave 新聞稿 | [Interested-party] | 基準維持 0.76；+25% 與短天期列為敏感度 |
| 隨需占比 | 比例**找不到**；短天期合約存在（新聞稿） | 同上 | [Interested-party] | 維持 0%；敏感度 |

未使用 SemiAnalysis。Q2 法說逐字稿原文未取得（Seeking Alpha／Platform Aeronaut 需訂閱），25% 的說法只有轉述。

### W5 實作（Excel＋JS 雙引擎；`company.json` → `companyAdjust`）
- **既有合約 k**（工作單第 4(a) 項）：`k_t＝既有占比_t × k_既有＋（1 − 既有占比_t）× k_新約`；既有合約 MW＝最新季末在役 1,500 MW 逐期扣汰換（最舊世代先出），既有占比＝平均既有 MW ÷ 平均在役 MW；`k_既有＝（Q2 每在役 MW 年收入 − 服務）÷（計費比例 × 利用率 × Q2 錨）`，Q2 錨隨目前成本情境（價格是事實）；`k_新約＝隨需占比 × k_現貨＋（1 − 隨需占比）× k_長約 ×（1＋新約價格調整）`。輸入：既有合約 k 開關（1）、新約價格調整（0）。
- **營運成本倍數**（輸入 1；敏感度＝Q2 實際 ÷ Q2 季末世代 Tokenomics 合計 0.654）、**每 MW 建置成本倍數**（既有輸入；敏感度＝1H26 實際 ÷ 首期新增世代 Tokenomics 0.714）。
- 新頁「公司實況驗證」：參數驗證 12 列（Tokenomics 值｜CRWV 實際｜差距｜採用值｜規則；文字欄＝名稱、實際與標記、機制與證據、公司調整、證據）、敏感度輸入 5 列、Q2 逐項對帳 6 列、證據清單、找不到清單。HTML「每 MW 經濟性 → 公司實況驗證」同列（cmp31 比對）。
- 敏感度新增 6 組：既有合約 k 不套用（＝W4）、新約 +25%、隨需 10% × CRWV 短天期 k 3.14、營運成本＝Q2 實際比率、建置成本＝1H26 實際比率、兩者皆實際。
- 舊方法回歸：`verify_legacy.sh` 的 v4.6 副本另移除 `companyAdjust`（公司調整只在新方法使用）。滾動清單 `ROLL_FIELDS` 新增 `companyAdjust.capexActual`（期間標籤取日曆的年初至今標籤）。

### W5 關鍵數字
- 加權目標價（保守／基準／積極）：v4.6 $43.99／$29.68／$19.95 → v4.7 W4 $13.98／$5.25／$0.00 → **v4.7 W5 $18.44／$13.29／$0.00**（皆賣出）；融資缺口 32.7／81.6／157.9 → 54.8／104.8／183.5 → **51.0／100.9／179.6**。
- 基準每 MW 年收入（含服務，每平均在役 MW）FY26–FY30：v4.6 10.05／10.56／10.33／9.63／8.65；W4 7.64／8.16／8.40／8.30／7.86；**W5 8.22／8.58／8.67／8.50／8.00**；k 0.82／0.80／0.79／0.78／0.78。營運成本各項三版相同（電費 0.67、IT 維護 0.94–1.08、人員軟體 0.33、稅險 0.16–0.18、租金 1.60→1.00）。
- 拆解（對 v4.6，(d)）：① +30.16／+15.46／+8.82、② −60.17／−39.89／−28.77、③ 0、④-1 既有合約 k +4.46／+8.05／0.00、④-2–④-5 0；相加＝總變動 −25.55／−16.38／−19.95。
- 敏感度（基準）：既有 k 不套用 $5.25、新約 +25% $27.74、隨需 10% × 短天期 k 3.14 $32.28、營運成本 Q2 比率 $22.47、建置成本 1H26 比率 $26.00、兩者皆實際 $43.30、k_長約 0.89 $21.22。

## W6 接 Tokenomics v5.31（v4.8；工作單 `docs/workorders/20261009_coreweave_W6_Tokenomics_v5.31.md`）

Tokenomics：master `f16f161`（PR #36 合併），`model/CURRENT`＝`20261008_Tokenomics_v5.31.xlsx`（兩個唯讀副本皆已更新到 origin/master）。

| 步驟 | 狀態 | commit | 備註 |
|---|---|---|---|
| 0 開分支、draft PR、Tokenomics 副本更新 | 完成 | （本 commit） | 自 origin/main 8671a11；環境：soffice 24.2.7、openpyxl、playwright Chromium 可啟動（未補裝） |
| 1 快照 v5.31、名稱清單、`tokenomics` 區段 | 完成 | （本 commit） | `data/tokenomics_snapshot_v5.31.json`（29 名、missing 0；刪 v5.27 快照）；新增 `IF_MaintITWarr`、`IF_MaintITPost`、`IF_WarrantyYrs`；各名稱儲存格位置與 v5.27 相同。`verify.sh` 22 項全過（cmp31 三情境各 528 項、FY27 錨定 474 項）。此時 IT 維護仍為等值費率（IF_MaintIT v5.31），基準加權目標價 $13.29 → $15.37（中間值，只作紀錄） |

#### W6 第 1 步：引用名稱前後值（v5.27 → v5.31；US$m／MW／年＝$B/GW/年，基準欄；低／高另列於對照報告）
| 名稱 | 變動 |
|---|---|
| IF_CapexIT | GB300 37.45 → 32.24（−13.9%；高 50.95 → 39.28）；其他世代不變 |
| IF_CapexTotal | GB300 50.12 → 44.91（高 67.14 → 55.47） |
| IF_DeprIT | GB300 6.24 → 5.37（高 12.74 → 9.82） |
| IF_TaxIns | GB300 0.2506 → 0.2245（高 0.537 → 0.444） |
| IF_MaintIT（等值費率 3% → 1.57%） | Hopper 0.825 → 0.433、GB200 0.721 → 0.378、GB300 1.123 → 0.507、VR200 1.128 → 0.591、Rubin Ultra 1.503 → 0.788 |
| IF_MaintITWarr（新，保固期內） | Hopper 0.138、GB200 0.120、GB300 0.161、VR200 0.188、Rubin Ultra 0.250 |
| IF_MaintITPost（新，保固期滿） | Hopper 0.825、GB200 0.720、GB300 0.967、VR200 1.128、Rubin Ultra 1.503（＝IT 資本 × 3%；GB300 隨資本下降） |
| IF_WarrantyYrs（新） | 3 年 |
| IF_HoldEcon | Hopper 10.10 → 9.70、GB200 9.17 → 8.83、GB300 12.72 → 10.89、VR200 12.76 → 12.23、Rubin Ultra 16.07 → 15.35 |
| IF_HoldAcct | GB300 9.50 → 7.99；其他世代 −4% 到 −6% |
| IF_GPUhrEcon（US$/GPU-hr） | H100 1.605 → 1.543、GB200 2.087 → 2.009、GB300 2.983 → 2.552、VR200 4.994 → 4.784 |
| IF_OpexGW | GB300 2.63 → 1.98；其他世代 −17% 到 −23%（只對照） |
| L1_GPUhr_GB200_vsCW／L1_GPUhr_GB300_vsBE | 2.086 → 2.008／2.983 → 2.552 |
| 不變 | IF_RacksPerGW、IF_GPUsPerGW、IF_FacilityGW、IF_CapexFacility、IF_PowerCost、IF_Util、L1_FacCapexMW、L1_RevGW_Fleet_VR200、IF_DeprLifeIT、IF_DeprFac、IF_AvgDraw、IF_PowerPrice、IF_MaintFac、IF_StaffSW、IF_RevGWFleet |
| 2 IT 維護依機齡兩段（Excel＋JS） | 完成 | （本 commit） | `methodology.perMw.maint`＝age；`fleet.openMix.vintages`（2023 年 100／2024 年 260／2025 年 490／2026 上半年 650 MW，投入使用＝期間中點）；Excel「輸入與假設」期初機齡層、TK IT 維護兩段、保固年限（TK_WarrantyYrs）；「每MW經濟性」新增「機齡與保固」區（期間起訖、各層／各期新增的保固期內比例、各世代保固期內／期滿平均在役 MW）與 IT 維護兩部分、等值費率對照列；「公司實況驗證」IT 維護 Tokenomics 值改依機齡、Q2 季末營運成本合計改依機齡（新增「Q2 季末保固期內占比」）；JS `vintQ`／`maintAgeQ`／`q2WarrQ`；cmp31 三情境各 555 項（+27）、FY27 錨定 501 項；舊方法回歸 `verify_legacy.sh` 設 maint=flat |
| 3 公司實況驗證重跑 | 完成 | （本 commit） | 營運成本合計 Tokenomics 1.326 對 Q2 1.357：+2.3%（v4.7 −35%）→ 規則 1；每 MW 資本支出 30.59 對 1H26 24.83：−18.8%（v4.7 −28.6%）→ 仍 > 10%，規則 3；k_既有 0.911 對 k_長約 0.76（+19.9%，規則 4a；錨下降使 k_既有上升、既有合約收入不變）；租金、折舊年限一致。company.json 參數說明（maint、opexBundle、capex、q2Notes）改寫 |
| 4 MW 口徑查證 | 完成（未改） | — | EDGAR 10-K FY2025（crwv-20251231.htm）全文 curl 讀取＋WebFetch 第二次讀取：「active power」6 次、「contracted power」2 次，皆只有數字、無定義；「critical IT」只出現在資訊系統語境；10-Q Q2 2026 全文：「active power」0 次；EDGAR 全文檢索（efts）經代理 403；網路搜尋無公司原文定義 → 維持 `meta.mwBasis`＝IT。另發現：10-K 寫 2023 年底約 70 MW（模型 `defaults.mwYearEnd` 2023＝100），列未解問題 |

#### W6 第 2 步：IT 維護逐年值（US$m／平均在役 MW／年；v4.7＝等值費率 IF_MaintIT v5.27）
| 情境 | 項目 | FY26（2H） | FY27 | FY28 | FY29 | FY30 |
|---|---|---|---|---|---|---|
| 基準 | v4.7 IT 維護 | 0.94 | 0.99 | 1.03 | 1.05 | 1.08 |
| 基準 | v4.8 IT 維護（依機齡） | 0.185 | 0.216 | 0.275 | 0.388 | 0.503 |
| 基準 | 　保固期內部分 | 0.136 | 0.139 | 0.138 | 0.122 | 0.105 |
| 基準 | 　保固期滿部分 | 0.049 | 0.077 | 0.138 | 0.266 | 0.398 |
| 基準 | 保固期內占比 | 94.0% | 90.7% | 83.0% | 68.8% | 56.8% |
| 基準 | 對照：等值費率 IF_MaintIT v5.31 | 0.452 | 0.479 | 0.505 | 0.524 | 0.540 |
| 保守 | v4.8 IT 維護 | 0.185 | 0.216 | 0.278 | 0.409 | 0.578 |
| 積極 | v4.8 IT 維護 | 0.185 | 0.216 | 0.269 | 0.350 | 0.417 |
