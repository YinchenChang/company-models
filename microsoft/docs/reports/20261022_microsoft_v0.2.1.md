# Microsoft（MSFT）MAG v0.2.1 報告：Tokenomics 同步 v19.0（2026-10-22）

分支 `claude/mag-v0.2.1-tk19`（三家共用，PR [#64](https://github.com/YinchenChang/company-models/pull/64)）；工作單 `mag/docs/workorders/20261022_mag_v0.2.1_Tokenomics_v19同步.md`；進度檔 `docs/reports/20261022_microsoft_v0.2.1_進度.md`。成品 `dist/20261022_Microsoft收支模型_v0_2_1.html`／`.xlsx`；交接檔 `docs/handoff/20261022_Microsoft收支模型_交接檔_v0_2_1.md`。Tokenomics `model/CURRENT`＝`20261022_Tokenomics_v19.0.xlsx`，master `1916a06`（工作單寫 `ad33e28`，兩者之間 `model/` 無差異）。引擎自 `amazon/` v0.2.1 同步（第 8 節）。

## 1. 結論（三句）

1. **TK 同步後目標價**（讀法 1）：保守 **380.56**／基準 **359.06**／積極 **360.16**（v0.2：385.91／355.08／356.77），讀法 2：404.09／429.43／428.15；三情境、兩讀法皆「賣出」（賣出門檻 $450.30），評等不變。
2. **主要驅動是 k 隨成本重算，不是 Tokenomics 本身**：v5.31 → v19.0 所有驅動名稱的值完全不變，單純換快照使目標價變動 0.00；但 v0.2 的 k 是以舊 v5.27 成本固定的倍數，改成「證據價格 ÷ 目前成本」後 k_長約 0.76 → 0.891（IREN–Microsoft GB300 約 9.7 US$m/MW-年 ÷ GB300 持有成本，GB300 成本 12.725 → 10.886，−14.4%）、k_現貨 1.76 → 1.828，混合 k 1.06 → 1.172（+10.6%）；基準 +4.13、積極 +3.39、保守 −5.35。收入端上調 10% 但目標價只動 1%：Azure 首期營收固定，AI 營收增加只是從以 13.8× 估值的 Azure 非 AI 殘差移到以 6× 估值的 AI 雲端（與 v0.2 報告同一機制，方向相反）。
3. **主命題答案改變：由「差一點沒有」變為「有」**。FY29 基準對外 AI ROIC 9.81% → 13.55%（WACC 10.90%，差 −1.1pt → +2.7pt），打平 k 1.092 不變；目前混合 k 1.172 高於打平 k 約 7%。這個翻轉完全來自 k 的分母更新（價格不變），而且只有基準價格軸如此：k 低軸（0.877）下 ROIC 仍只有 3.7%（見第 2 節），結論取決於 k 落在打平 k 之上，已列入待 Project 判斷。

**補充**：若 k_長約維持 0.76，Microsoft FY29 對外 AI ROIC 為 10.49%，低於 WACC 10.90%（主命題答案取決於 k_長約 0.76 或 0.891）。

## 2. 前後對照（v0.2 → v0.2.1）

| 項目 | 保守 v0.2 | 保守 v0.2.1 | 基準 v0.2 | 基準 v0.2.1 | 積極 v0.2 | 積極 v0.2.1 |
|---|---|---|---|---|---|---|
| 目標價 讀法 1 | 385.91 | 380.56 | 355.08 | 359.06 | 356.77 | 360.16 |
| 目標價 讀法 2 | 409.44 | 404.09 | 425.10 | 429.43 | 424.43 | 428.15 |
| FY29 對外 AI ROIC | 1.75% | 4.77% | 9.81% | 13.55% | 9.55% | 13.22% |
| 打平 k | 1.394 | 1.397 | 1.092 | 1.093 | 1.101 | 1.102 |
| 採用 k（混合，基準價格軸） | 1.060 | 1.172 | 1.060 | 1.172 | 1.060 | 1.172 |

基準情境其他：資本支出 FY28／FY29 224.5／234.6（v0.2 224.9／235.4）；FCF FY27–FY31 1.2／37.2／80.5／88.3／23.3（v0.2 0.4／34.2／75.3／80.8／12.8）；新債五期合計 23.1（v0.2 48.5）、新股 0；回購（22.3／年）FY27、FY28 取消、FY31 減為 1.6；調整後槓桿最高 0.83×。3 × 3（讀法 1；列＝容量，欄＝k 低 0.877／基準 1.172／高 3.014）：保守 395.27／380.56／352.85；基準 348.89／359.06／532.46；積極 351.84／360.16／521.59。對外比例 40%／80%（基準）：$285.64／$391.20。評價口徑：WACC 9／10／11% → 419.06／383.70／356.52；ERP 4／6% → 390.91／334.42；非 AI 自身倍數 → 420.00；AI 15× → 408.67；組合 → 533.49（中立）；翻轉點 WACC 約 8.3%（v0.2 8.19%）。k 高軸（Azure 牌價 7.967）的 $532.46／$521.59 高於賣出門檻 $450.30，但那是牌價上限而非預設。

## 3. k 新舊值（證據價格 ÷ 目前快照同世代成本；證據價格本身未改）

| k | 規格（證據） | v0.2（固定） | v0.2.1 | 原因 |
|---|---|---|---|---|
| k_長約 低 | 0.45 × 12.762 ＝ 5.7429 US$m/MW-年 ÷ VR200 `IF_HoldEcon` | 0.450 | 0.470 | 分母 12.762 → 12.2255（−4.2%）|
| k_長約 基準／高 | IREN–Microsoft GB300 五年約 9.7 ÷ `IF_HoldEcon`｜GB300 | 0.760（＝9.7 ÷ 12.725 = 0.762 四捨五入）| 0.891 | 分母 12.725 → 10.8864（−14.4%）；另含 0.762 → 0.76 的四捨五入（0.3%）|
| k_現貨 低／基準 | Silicon Data H100 指數 2.82 ÷ `IF_GPUhrEcon`｜H100 | 1.76 | 1.828 | 分母 1.605 → 1.5426（−3.9%）|
| k_現貨 高 | Azure ND96isr_H100_v5 隨需 12.29 ÷ `IF_GPUhrEcon`｜H100 | 7.66 | 7.967 | 同上 |

k_現貨（基準）三家一致（1.828）；長約占比 70% 不變。k_長約基準的證據是 Microsoft 付給 neocloud 的買價（[Analogy]，代理賣價）；GB300 機架價格 −14% 使它的 k 受影響最大，是本家結論翻轉的來源。

## 4. 目標價變動拆解（Excel 重算；`scripts/tk_attrib.py`；順序固定，加總＝總變動，殘差 ≤ 0.0000）

順序：① 換快照（k、C10 速度固定為 v0.2 值）→ ② k 隨成本重算 → ③ C10 重解（基準 2,511.7 → 2,527.2；積極 2,400 固定）→ ④ 上限檢查層（新增第五列等）。

| 情境／讀法 | 舊版 | ① 換快照 | ② k 隨成本重算 | ③ C10 重解 | ④ 上限檢查層 | 新版 | 總變動 |
|---|---|---|---|---|---|---|---|
| 保守 讀法 1 | 385.91 | +0.00 | -5.35 | +0.00 | +0.00 | 380.56 | -5.35 |
| 保守 讀法 2 | 409.44 | +0.00 | -5.35 | +0.00 | +0.00 | 404.09 | -5.35 |
| 基準 讀法 1 | 355.08 | +0.00 | +4.13 | -0.15 | +0.00 | 359.06 | +3.98 |
| 基準 讀法 2 | 425.10 | +0.00 | +4.18 | +0.15 | +0.00 | 429.43 | +4.33 |
| 積極 讀法 1 | 356.77 | +0.00 | +3.39 | +0.00 | +0.00 | 360.16 | +3.39 |
| 積極 讀法 2 | 424.43 | +0.00 | +3.72 | +0.00 | +0.00 | 428.15 | +3.72 |

① 全為 0.00，證明換快照本身不影響目標價；④ 全為 0.00。保守情境 −5.35：容量受 Epoch 站點封頂，k 上調使 AI 營收升高，Azure 首期營收固定，非 AI 殘差下降（推測，未另做拆解）。

## 5. 上限檢查（五列；基準情境，FY27／FY28／FY29／FY30／FY31）

| 列 | FY27 | FY28 | FY29 | FY30 | FY31 | 門檻 |
|---|---|---|---|---|---|---|
| 每 MW 年收入 ÷ 利用率層 `IFW_RevGWFleet_Util` | 32.2% | 28.8% | 26.9% | 25.8% | 25.2% | 示警 50%（通過） |
| ÷ 生產折減層 `IFW_RevGWFleet_Prod` | 40.4% | 36.1% | 33.5% | 32.2% | 31.4% | 硬上限 100%（通過） |
| **第五列：÷ `L1_RevGW_PQ_VR200`（48.79 US$m/MW）** | 25.0% | 26.0% | 26.7% | 27.2% | 27.5% | 無門檻（對照） |
| 　以 Hi 端點（77.35）為分母 | 15.8% | 16.4% | 16.9% | 17.1% | 17.3% | |
| 　以 Lo 端點（35.70）為分母 | 34.2% | 35.6% | 36.5% | 37.1% | 37.6% | |

v19.0 把上限層上調 33%–124%，所以前四列比 v0.2（43.9%／58.7%）低；上調來自 Tokenomics 的層定義，與模型收入無關。第五列的 token 帳未閉合（模型付費 token 為需求面的 `L1_TokLedger_2026Q2` 約 6.26 倍），只作對照，不得與 `L1_RevGW_Fleet_VR200` 相減比較。

## 6. C10、C15

- C10：基準 2,511.7 → **2,527.2** MW-IT／年（`node scripts/calib_pace.js --write`；FY27 資本支出公式值 204.199 ＝ 指引換算 204.2）。積極 2,400 為固定值，不連動（基準仍高於積極，與 v0.2 相同）。
- C15（FY29，基準，對外口徑）：模型稅後 13.55% → 稅前 16.93% → k＝1 9.75% → 營運成本 9.75% → 移除汰換 9.75% → 穩態 13.71% → 乾淨稅前 ROIC **12.89%**（v0.2 12.9%）；通過（±5pt 內）。

## 7. 已套用的預設與待 Project 判斷

| 問題 | 採用 | 替代 | 影響 |
|---|---|---|---|
| 長約 k 基準的證據 | 直接引用 kevidence 的 IREN–Microsoft 買價 9.7 US$m/MW-年（lc.iren）÷ 目前 GB300 成本 | 以 v0.2 取整的 0.76 × 12.725 換成固定價格（k＝0.76 × 12.725 ÷ 10.886＝0.888，差異 0.3%）；或改取 Nscale／Anthropic 證據 | 見待判斷 |
| 長約 k 低（0.45） | v0.2 判斷倍數 × 當時 VR200 成本換成固定價格（5.7429），÷ 目前 VR200 | 改以 Anthropic–Azure 證據 | 低軸 0.470 |
| 自研晶片（Maia） | 錨與係數維持 v0.2 口徑（Maia 200 → GB300 × 1.0）；`IF_NonNVRatio`／`IF_NonNVCostRatio`／`IF_NonNV_Maia` 在 v19.0 仍不存在 | — | 無 |
| 積極情境速度 | 2,400 固定，不連動（v0.2 先例） | — | 0 |
| 第五列 | 無門檻，不影響任何目標價 | — | 0 |

**待 Project 判斷**：(1) 本家結論翻轉（ROIC 在 WACC 之上）是否採信：它取決於 k_長約基準以「買價 ÷ 目前成本」計（0.891）；買價含 neocloud 利潤、屬 [Analogy]，且 GB300 機架價格 −14% 全數轉進 k。若 k_長約仍取 0.76，混合 k＝0.70×0.76＋0.30×1.828＝1.08，略低於打平 k 1.09，答案回到「差一點沒有」。(2) 長約 k 規格（上表）。(3) `scripts/port_company_json.py` 內 pricing 區塊仍是 v0.2 固定倍數，重產 company.json 時須帶入現行 pricing（交接檔已註明）。

## 8. 引擎同步（自 `amazon/` v0.2.1，逐檔比對）

`segA.js`、`build_xlsx.py`、`cmp31.js`、`kspec.py`（新）、`scripts/tk_attrib.py`（新）、`scripts/gen_expect.py`（新）自 `amazon/` 複製；`scripts/fields_doc.py`（公司專屬說明檔）只套用 `pricing.kEvidence` 與 `pricing.kLong.low` 兩處說明；`tail.js`／`vlog.py` 只加 v0.2.1 一列。逐檔比對（`cmp`）：`app_pretty.js`、`build_html_portable.py`、`build_xlsx.py`、`calendar_q.py`、`cmp31.js`、`crawl.py`、`fix_*.py`、`kspec.py`、`load_engine.js`、`rename_ids.py`、`rm_andy.js`、`seg*.js`、`verify_ooxml.py`、`xl_diff.py`、`xlx.py`、`docs/template_v3_3.html` 與 `scripts/` 內全部引擎檔與 `amazon/` 逐字相同；`tail.js` 除 VLOG 外相同。

## 9. 驗證

`verify.sh --vs-dist`（對 v0.2 成品，`EXPECT=scripts/expect/v0_2_1_mag_vs_v0_2.txt`）：

```
================ 結果 ================
PASS  check_quarterly：季度加總＝年度、差異原因齊全
PASS  README 欄位說明與 company.json 一致
PASS  EBITDAR 校準：基準情境起點／穩態 EBITDA 率不變
PASS  併網速度校準：首期資本支出＝指引、對帳落差 0
PASS  Tokenomics 快照 --check：找不到 Tokenomics clone，略過（警告）
PASS  建 HTML
PASS  建 Excel
PASS  重算：0 公式錯誤
PASS  反向 DCF：Excel 求解，rv_snap.json 與 Excel 一致
PASS  fix_outline
PASS  fix_datatable
PASS  verify_ooxml：OOXML OK
PASS  快照值＝Excel 分頁值：43 個名稱（missing 3）、443 個值、139 個具名範圍一致；其他工作表 65 格公式引用 32 個 TK_ 名稱（v19.0）
PASS  離線開啟：新建 HTML
PASS  離線開啟：dist/ 成品
PASS  cmp31 low：544 項全部 OK
PASS  cmp31 base：563 項全部 OK
PASS  cmp31 high：544 項全部 OK
PASS  cmp31 base_FY27：544 項全部 OK
PASS  test_quarterly：假設實際數與可移植性測試通過
PASS  test_rolling：日曆推算與滾動後第一屏
PASS  test_mag_mechanisms：C8 (a)–(f)＋回購非零 一致
PASS  test_attrib：拆解工具（月數、同版 0、滾動一季與 WACC 12%）
PASS  crawl 畫面文字（升版：只檢查頁面錯誤與頁數） 0 差異
PASS  xl_diff --values --by-label 0 差異
PASS  xl_diff 公式 --by-label 0 差異
verify.sh：全部通過（26 項）
```

升版後（dist 換 v0.2.1）完整 `verify.sh`（`TOKENOMICS_DIR=/tmp/tk`，--check 實跑）：

```
================ 結果 ================
PASS  check_quarterly：季度加總＝年度、差異原因齊全
PASS  README 欄位說明與 company.json 一致
PASS  EBITDAR 校準：基準情境起點／穩態 EBITDA 率不變
PASS  併網速度校準：首期資本支出＝指引、對帳落差 0
PASS  Tokenomics 快照 --check：43 個名稱（missing 3）與快照一致（20261022_Tokenomics_v19.0.xlsx，相對誤差 ≤ 1e-09）
PASS  建 HTML
PASS  建 Excel
PASS  重算：0 公式錯誤
PASS  反向 DCF：Excel 求解，rv_snap.json 與 Excel 一致
PASS  fix_outline
PASS  fix_datatable
PASS  verify_ooxml：OOXML OK
PASS  快照值＝Excel 分頁值：43 個名稱（missing 3）、443 個值、139 個具名範圍一致；其他工作表 65 格公式引用 32 個 TK_ 名稱（v19.0）
PASS  離線開啟：新建 HTML
PASS  離線開啟：dist/ 成品
PASS  cmp31 low：544 項全部 OK
PASS  cmp31 base：563 項全部 OK
PASS  cmp31 high：544 項全部 OK
PASS  cmp31 base_FY27：544 項全部 OK
PASS  test_quarterly：假設實際數與可移植性測試通過
PASS  test_rolling：日曆推算與滾動後第一屏
PASS  test_mag_mechanisms：C8 (a)–(f)＋回購非零 一致
PASS  test_attrib：拆解工具（月數、同版 0、滾動一季與 WACC 12%）
verify.sh：全部通過（23 項）
```

`--check`：快照 42 個名稱（missing 3）可由 Tokenomics `20261022_Tokenomics_v19.0.xlsx` 重現（相對誤差 ≤ 1e-9）。

## 10. 補充（PR #64 追加；版本仍為 v0.2.1）

- **k 證據價格改為獨立輸入格**：Excel「輸入與假設」B0 新增「k 證據價格」區塊（每個證據價格一格、藍字，旁邊是目前同世代成本與倍數，旁註寫來源與性質）；k｜長約／k｜現貨 公式改引用這些格（`=C<列>/TK_…`，不再把價格寫在公式裡）。HTML 檢查頁新增「k 證據價格 ÷ 目前快照同世代成本」一項同步顯示。價格登錄於 `data/kevidence_*.json` 的 `kspecInputs`，`scripts/check_kevidence.py`（已納入 verify.sh）檢查 company.json 與其逐筆相同、且 kLong／kSpot 皆為 ref。
- **數字不變**：對本分支原 dist 的 `xl_diff --values` 為 0 差異，三情境目標價與所有數字逐格不變；公式差異僅為新增列與常數改引用輸入格（xl_diff 歸類為「常數改為引用輸入格」，`--expect` 清單為空）。
- **主命題取決於 k_長約**：若 k_長約維持 0.76，Microsoft FY29 對外 AI ROIC 為 10.49%，低於 WACC 10.90%（主命題答案取決於 k_長約 0.76 或 0.891）。 Microsoft 長約低軸的證據價格 5.7429 也是 v0.2 判斷倍數 0.45 × 12.762 的換算值，不是觀察價格；基準 9.7（IREN）是 neocloud 買價（[Analogy]）。
- `scripts/port_company_json.py` 的 pricing 區塊已改為現行 kspec 規格（證據價格＋世代），並新增 `scripts/check_port_pricing.py`：重產後的 pricing 與現行 company.json 完全相同（17 個欄位）。

`verify.sh --vs-dist`（對本分支原 dist，`EXPECT=scripts/expect/v0_2_1b_證據價格輸入格.txt`）：

```
================ 結果 ================
PASS  check_quarterly：季度加總＝年度、差異原因齊全
PASS  README 欄位說明與 company.json 一致
PASS  EBITDAR 校準：基準情境起點／穩態 EBITDA 率不變
PASS  併網速度校準：首期資本支出＝指引、對帳落差 0
PASS  Tokenomics 快照 --check：43 個名稱（missing 3）與快照一致（20261022_Tokenomics_v19.0.xlsx，相對誤差 ≤ 1e-09）
PASS  k 證據價格：k 證據價格 4 筆與 kevidence_microsoft_20261008.json 一致，kLong／kSpot 6 個規格皆為 ref
PASS  建 HTML
PASS  建 Excel
PASS  重算：0 公式錯誤
PASS  反向 DCF：Excel 求解，rv_snap.json 與 Excel 一致
PASS  fix_outline
PASS  fix_datatable
PASS  verify_ooxml：OOXML OK
PASS  快照值＝Excel 分頁值：43 個名稱（missing 3）、443 個值、139 個具名範圍一致；其他工作表 69 格公式引用 32 個 TK_ 名稱（v19.0）
PASS  離線開啟：新建 HTML
PASS  離線開啟：dist/ 成品
PASS  cmp31 low：544 項全部 OK
PASS  cmp31 base：563 項全部 OK
PASS  cmp31 high：544 項全部 OK
PASS  cmp31 base_FY27：544 項全部 OK
PASS  test_quarterly：假設實際數與可移植性測試通過
PASS  test_rolling：日曆推算與滾動後第一屏
PASS  test_mag_mechanisms：C8 (a)–(f)＋回購非零 一致
PASS  test_attrib：拆解工具（月數、同版 0、滾動一季與 WACC 12%）
PASS  crawl 畫面文字（升版：只檢查頁面錯誤與頁數） 0 差異
PASS  xl_diff --values --by-label 0 差異
PASS  xl_diff 公式 --by-label 0 差異
verify.sh：全部通過（27 項）
```

完整 `verify.sh`（dist 以同版號重建；`TOKENOMICS_DIR=/tmp/tk`，--check 實跑）：

```
================ 結果 ================
PASS  check_quarterly：季度加總＝年度、差異原因齊全
PASS  README 欄位說明與 company.json 一致
PASS  EBITDAR 校準：基準情境起點／穩態 EBITDA 率不變
PASS  併網速度校準：首期資本支出＝指引、對帳落差 0
PASS  Tokenomics 快照 --check：43 個名稱（missing 3）與快照一致（20261022_Tokenomics_v19.0.xlsx，相對誤差 ≤ 1e-09）
PASS  k 證據價格：k 證據價格 4 筆與 kevidence_microsoft_20261008.json 一致，kLong／kSpot 6 個規格皆為 ref
PASS  建 HTML
PASS  建 Excel
PASS  重算：0 公式錯誤
PASS  反向 DCF：Excel 求解，rv_snap.json 與 Excel 一致
PASS  fix_outline
PASS  fix_datatable
PASS  verify_ooxml：OOXML OK
PASS  快照值＝Excel 分頁值：43 個名稱（missing 3）、443 個值、139 個具名範圍一致；其他工作表 69 格公式引用 32 個 TK_ 名稱（v19.0）
PASS  離線開啟：新建 HTML
PASS  離線開啟：dist/ 成品
PASS  cmp31 low：544 項全部 OK
PASS  cmp31 base：563 項全部 OK
PASS  cmp31 high：544 項全部 OK
PASS  cmp31 base_FY27：544 項全部 OK
PASS  test_quarterly：假設實際數與可移植性測試通過
PASS  test_rolling：日曆推算與滾動後第一屏
PASS  test_mag_mechanisms：C8 (a)–(f)＋回購非零 一致
PASS  test_attrib：拆解工具（月數、同版 0、滾動一季與 WACC 12%）
verify.sh：全部通過（24 項）
```
