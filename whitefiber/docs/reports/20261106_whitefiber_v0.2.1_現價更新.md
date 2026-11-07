# WhiteFiber v0.2 → v0.2.1：只換現價（2026-11-05 收盤）

## 目前狀態
完成。成品 `dist/20261106_WhiteFiber收支模型_v0_2_1.html`／`.xlsx`；舊版移 `dist/_待刪除/`。待 Andy：無。

## 改了什麼
- Andy 2026-11-06 07:32「請所有個股都換成十一月五號的收盤價。」
- `company.json`：`meta.priceDate` 2026-10-07 → 2026-11-05、`meta.updateDate` → 2026-11-06；`callFacts.priceLast` 16.83 → 7.80、`callFacts.priceDate`、新增 `callFacts.priceSource`（Excel 現價列來源改讀此欄）；`defaults.eqPx`、`valuation.price` → 7.80。Comps 市值、共識檔不變。
- 來源：Yahoo Finance 7.80；Nasdaq.com 報價（4:00 PM 收盤）7.80；StockAnalysis 7.79。

## 關鍵數字前後對照
| 項目 | v0.2 | v0.2.1 |
|---|---:|---:|
| 現價 | 16.83 | 7.80 |
| 基準加權目標價 | 15.2（中立） | 15.9（中立） |
| 保守（只有已簽約） | 11.5（賣出） | 13.0（中立） |
| 積極（NC-2／3 與 Krambu） | 25.4（買進） | 25.3（買進） |
| 賣出門檻（現價 × 85%） | 14.31 | 6.63 |

- 基準上檔 +103% 仍為中立：買進需「上檔 ≥ 25% 且終值占 EV < 90%」，本版終值占 EV 364%（company.json → methodology.rating.buyTvShareMax）。
- 拆解：(a) 0、(b) 0、(c) 假設變更（現價與發行價）基準 +0.7（保守 +1.5、積極 −0.1）、(d) 0。目標價方向與 CRWV／NBIS 相反（現價降、目標價小升）的原因本輪未逐項拆解（候選：可轉債價內判斷用 MIN(現價, 第一輪目標價)，現價降低後可轉債視為價外、不計轉換稀釋），待查。

## verify.sh（DATE=2026-11-06）
```
================ 結果 ================
PASS  check_quarterly：季度加總＝年度、差異原因齊全
PASS  README 欄位說明與 company.json 一致
PASS  EBITDAR 校準：基準情境起點／穩態 EBITDA 率不變
PASS  Tokenomics 快照 --check：找不到 Tokenomics clone，略過（警告）
PASS  建 HTML
PASS  tk_sens 快照：low 12.02／13.01／16.46；base 14.98／15.87／20.82；high 24.63／25.29／29.06（無變動）
PASS  建 Excel
PASS  重算：0 公式錯誤
PASS  反向 DCF：Excel 求解，rv_snap.json 與 Excel 一致
PASS  fix_outline
PASS  fix_datatable
PASS  verify_ooxml：OOXML OK
PASS  快照值＝Excel 分頁值：46 個名稱（missing 0）、432 個值、144 個具名範圍一致；其他工作表 64 格公式引用 41 個 TK_ 名稱（v19.0）
PASS  離線開啟：新建 HTML
PASS  離線開啟：dist/ 成品
PASS  cmp31 low：484 項全部 OK
PASS  cmp31 base：484 項全部 OK
PASS  cmp31 high：484 項全部 OK
PASS  cmp31 base_FY27：484 項全部 OK
PASS  test_quarterly：假設實際數與可移植性測試通過
PASS  test_rolling：日曆推算與滾動後第一屏
PASS  test_attrib：拆解工具（月數、同版 0、滾動一季與 WACC 12%）
verify.sh：全部通過（22 項）
```
