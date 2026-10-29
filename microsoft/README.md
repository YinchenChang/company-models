# Microsoft（MSFT）個股模型 v0.3.2 — 命題＋官方營收分部 → 損益表

更新日 2026-10-29（v0.3.1 季度滾動：工作單 `mag/docs/workorders/20261029_MAG_v0.3.1_季度滾動_2026Q3.md`；架構工作單 `mag/docs/workorders/20261027_MAG_v0.3_個股模型重建_分部架構.md`；固定規格 `docs/plan/個股模型原則_命題加營收分部.md`）。資料截至 **FY27Q1**（2026-09-30；財報 2026-10-28：新聞稿、IR 財務報表、10-Q、FY27Q2 展望簡報、法說）。評價日 2026-09-30、目標價時點 2027-10-29。財年結束於 6/30（FY2027＝2026-07～2027-06）。金額十億美元（US$B）。

## 命題與核心問題

- **命題（一句話）**：AI 能否同時在 Azure（算力）與 Microsoft 365（Copilot 應用）變現，撐住翻倍的資本支出。
- **核心問題**：FY2029 的 AI 增量營業利益能否覆蓋 AI 資本支出的資金成本（AI 增量 ROIC 對 WACC），且自由現金流不需停止回購。

## 結果摘要（v0.3.1；數字皆由 Excel 重算）

| 項目 | 情境 A（樂觀） | 情境 B（基準） | 情境 C（保守） |
|---|---|---|---|
| 營收 FY2029 | 587.3 | 555.7 | 519.3 |
| 　Azure AI（命題子列） | 153.8 | 128.2 | 97.4 |
| 　Copilot（命題子列） | 31.5 | 24.0 | 16.5 |
| 營業利益率 FY2029（財報口徑；FY2026 46.8%） | 41.9% | 39.1% | 35.3% |
| EPS FY2029：財報口徑／評價口徑（FY2026 17.95） | 26.01／30.04 | 22.89／27.37 | 19.24／24.17 |
| AI 增量 ROIC FY2029：讀法 2／讀法 1（WACC 10.29%） | 13.2%／6.4% | 8.6%／1.0% | 3.1%／−5.3% |
| 收支關卡：自由現金流扣股利（FY2027–FY2031 合計） | 363.2 | 259.6 | 137.9 |
| 情境目標價（DCF 50%＋SOTP 50%；讀法 2） | 423.25 | 386.05 | 341.07 |

- **加權目標價 $384.10**（讀法 2 主值；v0.3 $385.56，−1.46；DCF 324.37、分部加總 443.84；情境 25／50／25），現價 $531.01（2026-10-28 收盤）→ 上檔空間 **−27.7%**，**投資建議：放空**（中立需 > 398.26）。讀法 1（自用 AI 只計成本）$310.88（−41.5%，放空；v0.3 315.27）。市場共識目標價 605.12（S&P 57 位）／596.07（MarketBeat 49 位），皆財報前 2026-10-27。
- **變動拆解（Recon 第 5 節，加總＝總變動）**：(a) 時間推移 +0.27、(b) FY27Q1 實績 +0.42、(c) 指引與 C10 +1.79、(d) Tokenomics v21.7（0）＋NonNV（Maia 改讀 IF_NonNV_*）−2.99、(e) 現價 +0.01、(f) OpenAI 持股 25% → 24% −0.96。
- **FY27Q1 實績對 v0.3 模型季值**：營收 92.28 vs 93.49（−1.3%；Azure +4.1%、Windows −21.4%、授權 −7.6%）；營業利益 45.92 vs 43.22（+6.2%）；公司口徑資本支出 45.2（第三方產能延後；v0.3 全年 ÷ 4＝51.1）。
- **C10**：曆年 2026 資本支出由約 175 調為約 170（第三方產能延到 FY27 下半年）；FY2027 推估 204.2（上半年 97.1＋下半年 107.1）不變，MW 速度依 v0.3 引擎重解 2,527.2 → 2,726.5 MW／年（v0.3 有 8.3 的資本支出沒有對應 MW；Maia 每 MW 資本支出較低）。
- **收支關卡**：基準五年合計 +259.6（v0.3 270.9）；FY2027 自由現金流 11.3，需舉債 8.1、回購兩年減少（共少 44.3）——與公司「FY27 自由現金流為正、適當時機取得額外資金」一致。
- 報告：`docs/reports/20261029_microsoft_v0.3.1_季度滾動.md`（v0.3 的架構說明與修正一至三見 `docs/reports/20261027_microsoft_v0.3_分部架構.md`）。

## 檔案

| 檔案 | 內容 |
|---|---|
| `build_microsoft.py` | Python → Excel 產生器（openpyxl，全部活公式；藍字輸入、黑字公式）；同時寫 `html.json`。`--set`／`--legacy-q1`／`--snap` 供變動拆解 |
| `make_v03.sh` | 一次跑完：事實總帳雙讀 → 建置 → LibreOffice 重算 → `check_v03.py` → segment_check → HTML（`SENS=1` 先重跑 C10 重解與變動拆解） |
| `check_v03.py` | 成品核對：公式錯誤 0、核對列 0、C10 對帳落差 0、固定列標籤、拆解 JSON 未過期 |
| `facts_fy27q1.py` | FY27Q1 事實總帳：新聞稿 docx 與 IR 財務報表 xlsx（必要時 10-Q）兩次讀取核對 → `data/microsoft_facts_FY27Q1_20261029.json` |
| `attrib_v031.py` | C10 重解（割線法，每次整本重算）與 v0.3 → v0.3.1 變動拆解 → `data/attrib_v031_20261029.json` |
| `dist/20261029_Microsoft個股模型_v0.3.2.xlsx` | 成品（v0.3.2＝v0.3.1 數值＋In 頁情境驅動對照）（LibreOffice headless 重算，公式錯誤 0） |
| `dist/20261029_Microsoft個股模型_v0.3.2.html` | HTML 展示層（`tools/company_html`；數字全由 xlsx 抽出） |
| `html.json` | HTML 設定（由產生器寫出，列號與 xlsx 一致） |
| `segments.json` | 官方分部宣告（`tools/segment_check/check_segments.py microsoft`） |
| `data/` | 事實總帳（SEC 一手 v0.1a 起＋FY27Q1）、k 證據、共識、同業倍數、Tokenomics v21.7 快照（v21.4 快照保留供拆解重現） |
| `docs/reports/20261029_microsoft_v0.3.1_季度滾動.md` | 本版報告 |
| `legacy_v0.2/` | 舊引擎（v0.2.2 以前），保留不刪 |
| `dist/_待刪除/` | v0.3.1、v0.3、v0.2.2 成品 |

建置：

```
bash microsoft/make_v03.sh          # SENS=1 時先重跑 C10 重解與變動拆解
```

## 結構（三層）

1. **命題模組（Thesis）**：在役對外 AI MW（期初 3,250；v0.3.1 C10 重解每年 +2,726.5，使 FY2027 模型資本支出＝公司指引推估 204.2）× 每 MW 年收入（Tokenomics `IF_HoldEcon` 世代加權 × k；Maia 200 以 `IF_NonNV_Out_Maia` × VR200）→ Azure AI；AI 資本支出（新增自有 AI MW × Tokenomics 每 MW 成本；Maia 以 `IF_NonNV_Hold_Maia` × VR200）→ AI 折舊與 AI 營運成本 → Agents and Infra 營業利益；AI 增量 ROIC 對 WACC。
2. **營收分部層（Seg）**：官方分部 Agents and Infra、Devices and Consumer（8-K 2026-09-02 重編）→ 八條產品線：Azure、Microsoft 365 cloud、Productivity and server licensing、Industry solutions、Frontier and support services；Search and advertising、XBOX、Windows OEM and devices。命題子列：Azure ＝ Azure AI＋Azure 其他；Microsoft 365 cloud ＝ M365 基礎＋Copilot。其他產品線＝FY2026 年增線性收斂到長期值（FY2032）。FY2027＝Q1 實績＋第二至第四季模型（FY2026 同期 ×（1＋年增）；Q1 實績與模型的差額歸 Azure 其他、M365 基礎等殘差列，不回推 k）。
3. **損益（PL）→ CF → Val**：分部營業利益 → 合併 → 業外（持股評價損益預測期 0）→ 稅 20% → 淨利 → EPS；自由現金流、股利、回購、舉債；DCF＋分部加總。

## 驅動 → 推導對照表

| 頁 | 輸入（In 頁鍵） | 推導量 | 因果方向 |
|---|---|---|---|
| Thesis | `pace`、`mw_ext_26`、`ext_sh`、`rent_ext_f` | 在役對外 AI MW、自有 AI MW（含自用） | MW 越多 → Azure AI 營收↑、AI 資本支出與折舊↑ |
| Thesis | `tk_hold_*`、`p_long`、`p_spot`、`ls_A/B/C` | k、每 MW 年收入 | k 越高 → Azure AI 營收與 EBITDA↑（三情境主要差異） |
| Thesis | `tk_cxit_*`、`tk_cxfac`、`self_build`、`capex_floor27`、`life_it` | AI 資本支出、AI 折舊（含 GPU 汰換） | 資本支出 → 折舊 → 分部營業利益↓；→ 自由現金流↓ |
| Seg | `cop_s26`、`cop_add_A/B/C`、`cop_px` | Copilot 營收 | 座席 × 月費 → Microsoft 365 cloud↑ |
| Seg | `glt_*`、`n_conv` | 非命題產品線營收 | 近一年年增收斂到長期值 |
| PL | FY2026 校準、`dm_seg` | Agents and Infra 核心利益率、分部營業利益 | 核心利益率 ×（分部 − Azure AI）＋ Azure AI − AI 營運成本 − AI 折舊 − FY2027 補足差額的折舊 |
| PL | `tax_f`、`oie_f` | 淨利、EPS | — |
| CF | `sbc_r`、`wc27`、`cx_*`、`fl_pay_*`、`bb_target`、`dps_*` | 自由現金流、回購、舉債、股數 | 自由現金流扣股利 → 回購（不足先減）→ 舉債（最低現金 50） |
| Val | `rf`、`beta`、`erp`、`kd`、`g_tv`、`m_*`、`al_*`、`w_*` | DCF、SOTP、加權目標價、投資建議 | WACC↑ → DCF↓；倍數↑ → SOTP↑ |

## 工作表

README → In（輸入一鍵一列、標記與區間）→ Q（FY25Q2–FY27Q1 八季、FY2026 四季合計、FY2027 年初至今與 FY2025、FY2026 年度值，核對列 0）→ Thesis → Seg → PL（最上方四列三情境核對列）→ CF（最上方收支關卡表與核對列）→ Val（0. 目標價；列 70 起各情境明細）→ Recon（共識、指引、FY27Q1 實績對 v0.3、v0.3 → v0.3.1 拆解、v0.3.1 vs v0.2.2）→ Sens（30 個案例的龍捲風＋判斷預設替代；列 80 起各案例完整引擎區塊）→ Sources。

地圖抽數用的列標籤：Val「目標價（DCF 與 Comparable 加權）」「DCF 目標價（情境加權）」「Comparable 目標價（市場主流同業法）」（＝分部加總 SOTP）「市場共識目標價」「現價（2026-10-28 收盤）」「上檔空間：加權目標價」「上檔空間級距」「投資建議」「DCF 情境價值推到 12 個月後（A／B／C）」；CF「收支關卡：自由現金流扣股利（五年合計）」（B／C／D 欄＝A／B／C 情境）。

## v0.3.2（2026-10-29，呈現修正，數值不變）
- In 頁 D／F 欄名改「敏感度低」「敏感度高」（只用於 Sens 頁，不進目標價）；頁尾新增「J. 情境驅動對照」：長約占比、k、Copilot 淨增、情境權重，情境 A／B／C 三欄並排（只引用上方各列）。
- HTML「主要輸入」拆成「情境驅動」與「單一參數區間」兩張表。規格見 `mag/README.md`；核對 `mag/check_in_layout.py`。
- 與 v0.3.1 逐格比對：只有 In 頁 A1、D3、F3 與新增一節不同，其餘 10 頁全部相同（加權目標價 384.10 不變）。v0.3.1 成品移 `dist/_待刪除/`。

## 與 v0.3 的差異（v0.3.1）

- 資料：FY27Q1 實績滾入（Q 頁八季窗口、FY2027＝Q1 實績＋剩餘三季模型、評價日 2026-09-30 的資產負債）。
- Tokenomics v21.4 → v21.7（契約名稱值 0 差異）；Maia 改讀 `IF_NonNV_*`（`nonnv`＝1；舊口徑列 Sens 判斷替代）。
- C10 重解、業外改用 FY27Q2 指引（−1.2／年）、OpenAI 持股 24%（10-Q）、現價與共識更新。

## 與 v0.2.2 的差異

- 架構：v0.2.2 沿用 CoreWeave 架構（對外 AI MW × 每 MW 收入＋其餘業務成長率線、融資瀑布、站點、租賃承諾）；v0.3 改為官方兩分部、八條產品線，命題只驅動 Azure AI 與 Copilot 兩條子列。
- AI 資本支出經折舊進分部營業利益與 EPS（v0.2.2 只在 EBITDA 與現金）。
- 評價：DCF 50%＋SOTP 50%（產品線 × 同業倍數）、情境 25／50／25（v0.2.2：DCF 45%＋分部 EV/EBITDA 55%，以基準情境為點位）。
- 目標價：v0.2.2 基準 $359.06 → v0.3 修正三加權 $385.56（權重 −3.18、DCF 腿 +2.68、倍數腿 +28.95、情境加權 −1.95；Recon 第 5 節）。
- 沿用不變（v0.3 時；v0.3.1 的變動見上節）：Tokenomics v21.4、k 規格（k_長約 0.891、k_現貨 1.828、混合 1.172）、C10 速度 2,527.2（FY2029 前）、對外比例 60%、自建比例 60%、ERP、OpenAI 持股 25% × 852 ×（1 − 20%）。修正一改變：讀法 2 為主值（Andy 2026-10-27 15:03 決定）、WACC 10.90% → 10.29%（rf 5.00%〔2026-10-26〕、Blume β、10 年期 kd、負債含融資租賃）、SOTP 的 AI 腿改模組 DCF。
