# Microsoft（MSFT）個股模型 v0.3 — 命題＋官方營收分部 → 損益表

更新日 2026-10-27（MAG v0.3 工作單 `mag/docs/workorders/20261027_MAG_v0.3_個股模型重建_分部架構.md`；固定規格 `docs/plan/個股模型原則_命題加營收分部.md`）。資料截至 FY26Q4（FY2026 10-K、2026-09-02 8-K 分部重編）；**FY27Q1 財報（2026-10-28）未納入**，之後以 v0.3.1 滾入。財年結束於 6/30（FY2027＝2026-07～2027-06）。金額十億美元（US$B）。

## 命題與核心問題

- **命題（一句話）**：AI 能否同時在 Azure（算力）與 Microsoft 365（Copilot 應用）變現，撐住翻倍的資本支出。
- **核心問題**：FY2029 的 AI 增量營業利益能否覆蓋 AI 資本支出的資金成本（AI 增量 ROIC 對 WACC），且自由現金流不需停止回購。

## 結果摘要（v0.3；數字皆由 Excel 重算）

| 項目 | 情境 A（樂觀） | 情境 B（基準） | 情境 C（保守） |
|---|---|---|---|
| 營收 FY2029 | 585.2 | 552.8 | 515.3 |
| 　Azure AI（命題子列） | 149.8 | 124.8 | 94.9 |
| 　Copilot（命題子列） | 31.5 | 24.0 | 16.5 |
| 營業利益率 FY2029（FY2026 46.8%） | 42.2% | 39.3% | 35.5% |
| EPS FY2029（FY2026 17.95） | 26.18 | 23.01 | 19.35 |
| AI 增量 ROIC FY2029（WACC 10.9%） | 6.5% | 1.1% | −5.2% |
| 收支關卡：自由現金流扣股利（FY2027–FY2031 合計） | 342.0 | 189.4 | 10.1 |
| 情境目標價（DCF 50%＋SOTP 50%） | 363.96 | 330.32 | 291.07 |

- **加權目標價 $328.92**（DCF 234.93、分部加總 422.91；情境 25／50／25），現價 $532.60（2026-10-26 收盤）→ 上檔空間 **−38.2%**，**投資建議：放空**。市場共識目標價 603.43。
- **命題答案（基準）**：對外 AI 雲端本身賺到資金成本（對外 AI ROIC FY2029 13.1%），但自用 AI MW（對外比例 60% 以外的 40%）只由 Copilot 收入覆蓋，AI 整體增量 ROIC 1.1%，遠低於 WACC。若自用 AI 以剛好回收持有成本計（讀法 2），AI ROIC 10.6%，約等於 WACC。
- **收支關卡**：基準五年合計 +189.4，但 FY2027 自由現金流只有 1.2（資本支出 204.2），需舉債 9.8、回購兩年減少。

## 檔案

| 檔案 | 內容 |
|---|---|
| `build_microsoft.py` | Python → Excel 產生器（openpyxl，全部活公式；藍字輸入、黑字公式）；同時寫 `html.json` |
| `dist/20261027_Microsoft個股模型_v0.3.xlsx` | 成品（LibreOffice headless 重算，公式錯誤 0） |
| `dist/20261027_Microsoft個股模型_v0.3.html` | HTML 展示層（`tools/company_html`；數字全由 xlsx 抽出） |
| `html.json` | HTML 設定（由產生器寫出，列號與 xlsx 一致） |
| `segments.json` | 官方分部宣告（`tools/segment_check/check_segments.py microsoft`） |
| `data/` | 事實總帳（SEC 一手）、k 證據、共識、同業倍數、Tokenomics v21.4 快照（沿用，未改） |
| `docs/reports/20261027_microsoft_v0.3_分部架構.md` | 本版報告 |
| `legacy_v0.2/` | 舊引擎（v0.2.2 以前：segA–E.js、build_xlsx.py、company.json、scripts/、CLAUDE.md、舊 README 等），保留不刪 |
| `dist/_待刪除/` | v0.2.2 成品（HTML、xlsx） |

建置：

```
python3 microsoft/build_microsoft.py raw.xlsx
soffice --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir out raw.xlsx
cp out/raw.xlsx "microsoft/dist/20261027_Microsoft個股模型_v0.3.xlsx"
python3 tools/company_html/build_company_html.py microsoft
python3 tools/company_html/check_company_html.py microsoft
python3 tools/segment_check/check_segments.py microsoft
```

## 結構（三層）

1. **命題模組（Thesis）**：在役對外 AI MW（沿用 v0.2.2：期初 3,250、每年 +2,527.2，由 FY2027 資本支出指引解出）× 每 MW 年收入（Tokenomics `IF_HoldEcon` 世代加權 × k）→ Azure AI；AI 資本支出（新增自有 AI MW × Tokenomics 每 MW 成本，FY2027 以公司指引推估 204.2 為下限）→ AI 折舊與 AI 營運成本 → Agents and Infra 營業利益；AI 增量 ROIC 對 WACC。
2. **營收分部層（Seg）**：官方分部 Agents and Infra、Devices and Consumer（8-K 2026-09-02 重編）→ 八條產品線：Azure、Microsoft 365 cloud、Productivity and server licensing、Industry solutions、Frontier and support services；Search and advertising、XBOX、Windows OEM and devices。命題子列：Azure ＝ Azure AI＋Azure 其他；Microsoft 365 cloud ＝ M365 基礎＋Copilot。其他產品線＝FY2026 年增線性收斂到長期值（FY2032）。
3. **損益（PL）→ CF → Val**：分部營業利益 → 合併 → 業外（持股評價損益預測期 0）→ 稅 20% → 淨利 → EPS；自由現金流、股利、回購、舉債；DCF＋分部加總。

## 驅動 → 推導對照表

| 頁 | 輸入（In 頁鍵） | 推導量 | 因果方向 |
|---|---|---|---|
| Thesis | `pace`、`mw_ext_26`、`ext_sh`、`rent_ext_f` | 在役對外 AI MW、自有 AI MW（含自用） | MW 越多 → Azure AI 營收↑、AI 資本支出與折舊↑ |
| Thesis | `tk_hold_*`、`p_long`、`p_spot`、`ls_A/B/C` | k、每 MW 年收入 | k 越高 → Azure AI 營收與 EBITDA↑（三情境主要差異） |
| Thesis | `tk_cxit_*`、`tk_cxfac`、`self_build`、`capex_floor27`、`life_it` | AI 資本支出、AI 折舊（含 GPU 汰換） | 資本支出 → 折舊 → 分部營業利益↓；→ 自由現金流↓ |
| Seg | `cop_s26`、`cop_add_A/B/C`、`cop_px` | Copilot 營收 | 座席 × 月費 → Microsoft 365 cloud↑ |
| Seg | `glt_*`、`n_conv` | 非命題產品線營收 | 近一年年增收斂到長期值 |
| PL | FY2026 校準、`dm_seg` | Agents and Infra 核心利益率、分部營業利益 | 核心利益率 ×（分部 − Azure AI）＋ Azure AI − AI 營運成本 − AI 折舊 |
| PL | `tax_f`、`oie_f` | 淨利、EPS | — |
| CF | `sbc_r`、`wc27`、`cx_*`、`fl_pay_*`、`bb_target`、`dps_*` | 自由現金流、回購、舉債、股數 | 自由現金流扣股利 → 回購（不足先減）→ 舉債（最低現金 50） |
| Val | `rf`、`beta`、`erp`、`kd`、`g_tv`、`m_*`、`al_*`、`w_*` | DCF、SOTP、加權目標價、投資建議 | WACC↑ → DCF↓；倍數↑ → SOTP↑ |

## 工作表

README → In（輸入一鍵一列、標記與區間）→ Q（FY25Q1–FY26Q4 與 FY2025、FY2026 實績，核對列 0）→ Thesis → Seg → PL（最上方四列三情境核對列）→ CF（最上方收支關卡表）→ Val（0. 目標價；列 70 起各情境明細）→ Recon（共識、指引、v0.3 vs v0.2.2 拆解）→ Sens（26 個案例的龍捲風＋判斷預設替代；列 70 起各案例完整引擎區塊）→ Sources。

地圖抽數用的列標籤：Val「目標價（DCF 與 Comparable 加權）」「DCF 目標價（情境加權）」「Comparable 目標價（市場主流同業法）」（＝分部加總 SOTP）「市場共識目標價」「現價（2026-10-26 收盤）」「上檔空間：加權目標價」「上檔空間級距」「投資建議」「DCF 情境價值推到 12 個月後（A／B／C）」；CF「收支關卡：自由現金流扣股利（五年合計）」（B／C／D 欄＝A／B／C 情境）。

## 與 v0.2.2 的差異

- 架構：v0.2.2 沿用 CoreWeave 架構（對外 AI MW × 每 MW 收入＋其餘業務成長率線、融資瀑布、站點、租賃承諾）；v0.3 改為官方兩分部、八條產品線，命題只驅動 Azure AI 與 Copilot 兩條子列。
- AI 資本支出經折舊進分部營業利益與 EPS（v0.2.2 只在 EBITDA 與現金）。
- 評價：DCF 50%＋SOTP 50%（產品線 × 同業倍數）、情境 25／50／25（v0.2.2：DCF 45%＋分部 EV/EBITDA 55%，以基準情境為點位）。
- 目標價：v0.2.2 基準 $359.06 → v0.3 加權 $328.92（權重 −3.18、DCF 腿 −43.54、倍數腿 +17.97、情境加權 −1.40；Recon 第 5 節）。
- 沿用不變：Tokenomics v21.4、k 規格（k_長約 0.891、k_現貨 1.828、混合 1.172）、C10 速度 2,527.2、對外比例 60%、自建比例 60%、WACC 的 CAPM 輸入、OpenAI 持股 25% × 852 ×（1 − 20%）。
