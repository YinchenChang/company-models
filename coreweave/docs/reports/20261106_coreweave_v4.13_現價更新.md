# CoreWeave v4.12 → v4.13：只換現價（2026-11-05 收盤）

## 目前狀態
完成。成品 `dist/20261106_CoreWeave收支模型_v4_13.html`／`.xlsx`；舊版移 `dist/_待刪除/`。待 Andy：無（共識檔 `data/consensus_crwv_20260925.json` 仍為 9/25，是否更新由 Andy 決定）。

## 改了什麼
- Andy 2026-11-06 07:32「最近市場變動比較大，請所有個股都換成十一月五號的收盤價。」
- `company.json`：`meta.priceDate` 2026-09-24 → 2026-11-05、`meta.updateDate` → 2026-11-06；`callFacts.priceLast` 90.13 → 43.66、`callFacts.priceDate` → 2026-11-05、新增 `callFacts.priceSource`（Excel 現價列來源改讀此欄，不再引用共識檔的 priceReference）；`defaults.eqPx`、`valuation.price` 90.13 → 43.66。
- 不變：Comps 的 CRWV 市值 `mktCapLast` 47.12（與同業同日 9/21，README 規則）、共識檔（只讀）。
- 來源：Yahoo Finance 日線 43.66；Motley Fool 2026-11-05 CRWV −9.83%（48.42 × (1 − 9.83%) = 43.66）；StockAnalysis 43.63（`docs/reports/20261105_收盤價_個股換價來源.md`）。

## 關鍵數字前後對照
| 項目 | v4.12 | v4.13 |
|---|---:|---:|
| 現價 | 90.13 | 43.66 |
| 基準（5.6 GW）加權目標價 | 32.0 | 25.5 |
| 保守情境 | 42.7 | 38.8 |
| 積極情境 | 14.7 | 8.2 |
| 賣出門檻（現價 × 85%） | 76.6 | 37.1 |
| 評等 | 三情境皆賣出 | 基準、積極賣出；保守 38.8 高於門檻 37.1 |
| 點位對現價 | −64% | −42% |

拆解：(a) 時間推移 0、(b) 實際數 0、(c) 假設變更（現價與發行價）基準 −6.5（保守 −3.9、積極 −6.5）、(d) 方法 0。機制：新股以「現價 × (1 − 折價 10%)」發行，現價下降 → 同一融資缺口需發更多股 → 每股價值下降。

## verify.sh（DATE=2026-11-06）
```
================ 結果 ================
PASS  check_quarterly：季度加總＝年度、差異原因齊全
PASS  README 欄位說明與 company.json 一致
PASS  Tokenomics 快照 --check：找不到 Tokenomics clone，略過（警告）
PASS  建 HTML
PASS  建 Excel
PASS  重算：0 公式錯誤
PASS  反向 DCF：Excel 求解，rv_snap.json 與 Excel 一致
PASS  每 MW 敏感度快照：Excel 求值，permw_sens.json 與 Excel 一致
PASS  fix_outline
PASS  fix_datatable
PASS  verify_ooxml：OOXML OK
PASS  快照值＝Excel 分頁值：43 個名稱（missing 0）、450 個值、150 個具名範圍一致；其他工作表 163 格公式引用 100 個 TK_ 名稱（v18.0）
PASS  離線開啟：新建 HTML
PASS  離線開啟：dist/ 成品
PASS  cmp31 low：595 項全部 OK
PASS  cmp31 base：595 項全部 OK
PASS  cmp31 high：595 項全部 OK
PASS  cmp31 base_FY27：526 項全部 OK
PASS  test_quarterly：假設實際數與可移植性測試通過
PASS  test_attrib：拆解工具（月數、同版 0、滾動一季與 WACC 12%）
PASS  test_permw：設施口徑換算與 GPU 小時價格路線
FAIL  test_rolling（見 /home/claude/company-models/coreweave/out/test_rolling.log）
verify.sh：失敗 1 項
```
test_rolling 失敗為既有問題（v4.12 原樣重跑同樣失敗：滾動測試 A 第一屏「1H26」過期字樣），與本次換價無關。
