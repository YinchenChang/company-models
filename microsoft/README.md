# Microsoft（MSFT）個股模型 v0.3.8 — 命題＋官方營收分部 → 損益表（v0.3.8 只改展示層並新增頁〔HTML 版面規格 v1.3、Coord 三層協調〕，數值不變；v0.3.7 只換現價；v0.3.6 只改展示層〔個股 HTML 版面規格 v1.0〕；v0.3.4：Tokenomics 快照 v21.9；v0.3.5：換 v21.10；基準數值同 v0.3.1）

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

- **加權目標價 $384.10**（讀法 2 主值；v0.3 $385.56，−1.46；DCF 324.37、分部加總 443.84；情境 25／50／25），現價 $531.01（2026-10-28 收盤）→ 上檔空間 **−27.7%**（v0.3.7：目標價 384.07、現價 535.93〔11-05〕→ −28.3%），**投資建議：放空**（中立需 > 398.26）。讀法 1（自用 AI 只計成本）$310.88（−41.5%，放空；v0.3 315.27）。市場共識目標價 605.12（S&P 57 位）／596.07（MarketBeat 49 位），皆財報前 2026-10-27。
- **變動拆解（Recon 第 5 節，加總＝總變動）**：(a) 時間推移 +0.27、(b) FY27Q1 實績 +0.42、(c) 指引與 C10 +1.79、(d) Tokenomics v21.7（0）＋NonNV（Maia 改讀 IF_NonNV_*）−2.99、(e) 現價 +0.01、(f) OpenAI 持股 25% → 24% −0.96。
- **FY27Q1 實績對 v0.3 模型季值**：營收 92.28 vs 93.49（−1.3%；Azure +4.1%、Windows −21.4%、授權 −7.6%）；營業利益 45.92 vs 43.22（+6.2%）；公司口徑資本支出 45.2（第三方產能延後；v0.3 全年 ÷ 4＝51.1）。
- **C10**：曆年 2026 資本支出由約 175 調為約 170（第三方產能延到 FY27 下半年）；FY2027 推估 204.2（上半年 97.1＋下半年 107.1）不變，MW 速度依 v0.3 引擎重解 2,527.2 → 2,726.5 MW／年（v0.3 有 8.3 的資本支出沒有對應 MW；Maia 每 MW 資本支出較低）。
- **收支關卡**：基準五年合計 +259.6（v0.3 270.9）；FY2027 自由現金流 11.3，需舉債 8.1、回購兩年減少（共少 44.3）——與公司「FY27 自由現金流為正、適當時機取得額外資金」一致。
- 報告：`docs/reports/20261029_microsoft_v0.3.1_季度滾動.md`（v0.3 的架構說明與修正一至三見 `docs/reports/20261027_microsoft_v0.3_分部架構.md`）。

## 檔案

| 檔案 | 內容 |
|---|---|
| `build_microsoft.py` | Python → Excel 產生器（openpyxl，全部活公式；藍字輸入、黑字公式）；同時寫 `html.json`。`--set`／`--legacy-q1`／`--snap` 供變動拆解 |
| `make_v03.sh` | 一次跑完：事實總帳雙讀 → 建置 → LibreOffice 重算 → `check_v03.py` → In 版面 → 敏感度覆蓋（`--require-keycol`）→ segment_check → 版面轉換（`tools/company_html/relayout.py`，v0.3.6 起）→ HTML（`SENS=1` 先重跑 C10 重解、變動拆解與 Sens 整本重算案例） |
| `sens_v033.py` | v0.3.3：Sens 整本重算案例（In 頁有低／高、活公式案例未涵蓋的鍵；含 he_all）→ `data/sens_v036_20261105.json`（v0.3.5 結果，v0.3.6 改名沿用；v0.3.3 結果移 `data/_待刪除/`） |
| `check_v03.py` | 成品核對：公式錯誤 0、核對列 0、C10 對帳落差 0、固定列標籤、拆解 JSON 與 Sens 整本重算案例未過期 |
| `facts_fy27q1.py` | FY27Q1 事實總帳：新聞稿 docx 與 IR 財務報表 xlsx（必要時 10-Q）兩次讀取核對 → `data/microsoft_facts_FY27Q1_20261029.json` |
| `attrib_v031.py` | C10 重解（割線法，每次整本重算）與 v0.3 → v0.3.1 變動拆解 → `data/attrib_v031_20261029.json` |
| `dist/20261105_Microsoft個股模型_v0.3.6.xlsx` | 成品（v0.3.6＝v0.3.5 數值＋README 展示層；v0.3.5＝v0.3.1 數值＋In 頁情境驅動對照＋敏感度覆蓋＋Tokenomics 快照 v21.10）（LibreOffice headless 重算，公式錯誤 0） |
| `dist/20261105_Microsoft個股模型_v0.3.6.html` | HTML 展示層（`tools/company_html`；數字全由 xlsx 抽出；版面依 `docs/plan/個股HTML版面規格.md`） |
| `html.json` | HTML 設定（由產生器寫出、再由 `relayout.py` 依 `layout_spec.py` 轉成新版面；列號與 xlsx 一致；勿手改） |
| `layout_spec.py` | v0.3.6：HTML 版面設定（In 參數依業務分四塊、估值與對帳的主要列） |
| `segments.json` | 官方分部宣告（`tools/segment_check/check_segments.py microsoft`） |
| `data/` | 事實總帳（SEC 一手 v0.1a 起＋FY27Q1）、k 證據、共識、同業倍數、Tokenomics v21.7 快照（v21.4 快照保留供拆解重現） |
| `docs/reports/20261029_microsoft_v0.3.1_季度滾動.md` | 本版報告 |
| `legacy_v0.2/` | 舊引擎（v0.2.2 以前），保留不刪 |
| `dist/_待刪除/` | v0.3.5、v0.3.4、v0.3.3、v0.3.2、v0.3.1、v0.3、v0.2.2 成品 |

建置：

```
bash microsoft/make_v03.sh          # SENS=1 時先重跑 C10 重解、變動拆解與 Sens 整本重算案例
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

## v0.3.8（2026-11-15，HTML 版面規格 v1.3＋三層協調 C0，數值不變）
- 依工作單 `docs/workorders/20261114_第二批13家_v1.3與三層協調.md`（樣板 Alphabet v0.3.9）；既有數值格逐格比對 0 差異（只有 README!A1 版本字樣不同），加權目標價 384.07、評等放空、讀法 1 目標價均不變。
- `microsoft/v038_post.py`（由 `build_microsoft.py` 在 README 之後、存檔前呼叫；`--set` 變體與非 v21.10 快照不跑）新增：PL 頁末標準損益表（GAAP；FY2024A–FY2026A 實績、FY2027E–FY2032E 情境 B、最右欄 FY27Q1A；營業成本與費用在預測期以「非折舊成本率＋折舊」與殘差推導，營業利益至 EPS 全取自 PL 情境 B；核對列全為 0）、RevDCF 頁（反向 DCF：現價隱含的營業利益率、營收規模、WACC〔二分法 40 次，核對列＝0〕與分部加總隱含倍數；無頁引用本頁）、Sens 頁末「加權目標價相對基準」低／高對照（龍捲風圖用，白話標籤）、Coord 頁、README 頁末 v1.3 白話列（橫幅、投資建議、重點觀察 5 點、分部營收推導、本版要點）。
- `microsoft/gen_html_v13.py` 在 `relayout.py` 之後把 `html.json` 轉成 v1.3（章節：命題→估值→敏感度→反向 DCF→損益→驅動因子→對帳→資料狀態；估值章首為四個目標價與現價圖，敏感度章首為兩張相對基準龍捲風圖）。`make_v03.sh` 加此步驟、`ROWMAP` 與版面檢查。
- Coord 頁（三層協調）：F1 讀取層級（Tokenomics v21.10 十七列，Σ|採用−層級值|＝0；Wattonomics v0.3.2 只作對帳列）；F1b 追公式（`tools/coord/trace_wattonomics.py --player msft`：788 格追蹤、207 格生效、148 格有旗標；結論見下）；F2 C0（實績）。
- C0 結果：資本支出（公司口徑，含融資租賃）CY2026 已公布三季（FY26Q3 31.9、FY26Q4 41.0、FY27Q1 45.2；公司 IR 財務報表 CapEx 工作表，Verified）合計 118.1；層級（Wattonomics CY2026 AI 119.96＋非 AI 32.97＝152.93）× 9/12＝114.70，**主讀數 +3.0%**；最近一季 × 4＝180.8 對 152.93，**第二讀數 +18.2%**。但 Wattonomics 2026 年 Microsoft 資本支出與新增 GW 由公司資本支出（In!K169 FY2026 實績 145.3、L169 個股模型寫回 FY2027 204.2、K168／L168 非 AI）換算，屬同源，故判定「參考（同源校準）／無可比實績（層級值同源）」，對帳中 0 項；2025 年底 GW 來自 Epoch（不同源）；`act26_msft` 只被 Wattonomics Compare 頁讀。現金 PP&E（FY27Q1 10-Q 39.367，Verified）只作口徑參考；公司指引約 170 對層級 152.93、Wattonomics 輸入 174.75 對輸出 152.93（−12.5%）只列參考。AI 營收（公司 AI 年化 >37 含 Copilot，FY27Q1 法說未更新年化數字）不揭露 Azure AI 單獨營收，無可比實績；在役容量存量公司不揭露，無可比實績（存量）。
- 新增容量 GW（PR #535 核對必修 1 更正：原寫「公司不揭露 GW」不實）：Microsoft 法說逐季口頭揭露新增容量——FY26Q2 約 1 GW、FY26Q3「another gigawatt」、FY26Q4 31 座資料中心約 1 GW；FY27Q1（2026-10-28）只說新增 24 座資料中心、未給 GW。Wattonomics 新增 GW＝公司資本支出 × AI 占比 ÷ 每 GW 成本（部分同源），依規則改比 GW 物理量：主讀數 CY2026Q1＋Q2 合計 2.0 GW 對層級 2.632 × 6/12＝1.316，+52%；第二讀數 FY26Q4 × 4＝4.0 對 2.632，+52%。但公司為整體容量（含非 AI）、未說明電力或 IT 口徑、可能含租用；以 Tokenomics 設施電力 1.2 與 AI 占比 0.85 橋接後 +7.7%，但這是 Assumed 參數下的讀數，**不得用來說差距已解**：（a）同一組參數套到 CY2025 參考列（公司 2.0 GW 對層級 2.135）得 −34%；未換算時 CY2025 與 CY2026 分別為 −6% 與 +52%，公司對層級比值由 0.94 升到 1.52，口徑因素（整體對 AI、設施對 IT、自有對租用）大致固定，解釋不了跳升，真正差異在 2026 年增速——層級 2025→2026 新增 GW 只增 23%，公司口頭約翻倍（最近揭露一季年化對 2025 全年 +100%）；可能原因為租用 neocloud 容量增加、層級每 GW 成本或上線時差。（b）設施電力係數 1.1–1.3 與 AI 占比 0.75–0.95 一併取區間時主讀數 −12% 至 +31%（Coord 公式算出），且租用容量未扣。結論隨 Assumed 參數跨過門檻，判定「參考（無可比實績：口徑無法對齊，缺 AI／IT 口徑新增容量揭露）」（無可比實績列數 3 → 4），不計對帳中；Coord 另加證據回饋註記「2025→2026 公司口頭容量增速遠高於層級，待公司揭露口徑」。每 GW 資本支出（參考）：公司 CY2026 上半年 72.9 ÷ 2.0 GW＝36.5 對 Wattonomics CapexGW!G22＋G25 每 IT GW 41.1。Coord 新增「公司口頭揭露新增容量」各列，HTML 對帳章加一張表。
- 追公式位址更正（PR #535 核對必修 2）：CapexAI ← Interface!F596 ← Gap!G62 ← G45＝G42＋G43＋G44；CapexAIModel ← Interface!F601 ← Gap!G56（＝G55 × 6/12＋H55 × 6/12）← G55 ← In!K167（空白）、H55 ← In!L167（168.45）。Coord 層級 CY2026 玩家資本支出列（現第 306 列）出處網址改為 Wattonomics 檔與快照路徑。
- PR #535 第二輪複核必修（數值不變）：新增容量 GW 判定改為規則類別「參考（無可比實績：…）」，判定說明、對帳摘要、README!A70 補 CY2025 同參數橋接（−34%）與兩參數區間（−12%～+31%）；Coord F1 加讀取列 WIF_Gap_Player_msft_2025_AddGW（2.135，CY2025 參考列 E 欄改連此列，Coord 第 29 列起列號 +1）、加「同參數橋接套到 CY2025」列與證據回饋註記列；HTML 新增容量表加 CY2025 橋接列、表下註記加證據回饋。
- 資料：`data/wattonomics_c0_v0.3.2.json`、`data/wattonomics_trace_v0.3.2.json`（`extract_wattonomics_c0.py` 由 Wattonomics v0.3.2 快照與主分支追公式工具產生）。Tokenomics 仍讀 v21.10（工作單所列 v22.0 未換快照）。
- 驗證：`check_v03.py`、`mag/check_in_layout.py`、敏感度覆蓋、segment_check、`check_company_html.py` 全過；`check_layout_v13.py --xlsx` 全過（只有全頁參數代號警告）；`tools/baseline_coordination/check.py` 0 項不符。v0.3.7 成品移 `dist/_待刪除/`。

## v0.3.7（2026-11-06，只換現價）
- Andy 2026-11-06 07:32「請所有個股都換成十一月五號的收盤價」：In 頁 `px` 531.01（10-28）→ 535.93（11-05 收盤；Yahoo Finance、Motley Fool、MarketAlleys，`docs/reports/20261105_收盤價_個股換價來源.md`）。
- 本模型 WACC 權益權重與回購價用現價（v0.3 起的設計，未改），故加權目標價 384.10 → 384.07（−0.03，Recon 第 5 節新增 (g) 列；拆解 JSON `attrib_v031_20261029.json` 保留 v0.3.1 終點，`check_v03.py` 改為核對「v0.3.1 終點＋(g)＝本檔」）。上檔空間 −27.7% → −28.3%，放空不變。
- 敏感度整本重跑（`sens_v037_20261106.json`，53 案目標價都隨 WACC 小幅變動）；評等改變 1 案：`beta` 低 中立 → 放空。Recon 第 7 節標籤 v0.3.1 → v0.3.7（公式本來就讀本檔 Val）。v0.3.6 成品移 `dist/_待刪除/`。

## v0.3.6（2026-11-05，只改展示層，數值格 0 差異）
- 依 `docs/plan/個股HTML版面規格.md` v1.0：README 頁第 42 列起新增「展示層」（命題摘要、命題 1–3、白話重點 1–3）、「命題回覆」（三題，格式「問題？答：…」）與「本版要點」三段，全部公式帶數字；A1 改版本字串，A2–A40 原列保留給下游。
- HTML 由 `tools/company_html/relayout.py microsoft <成品>` 轉成新版面（`make_v03.sh` 在 LibreOffice 重算後執行；設定 `layout_spec.py`）：投資建議橫幅、命題分點、命題回覆摺疊；In 參數分「Azure 與 AI 基礎設施／生產力與企業軟體／Windows、搜尋與 XBOX／全公司」四塊（H 共識與指引、J 情境驅動對照不歸組）；估值、對帳只列主要項目（對帳改每節一張表）；龍捲風圖只列變動 ≥ 10% 的案例。
- `build_microsoft.py` 的 html.json 另加：驅動 → 推導對照表標題（移附錄用）、DCF 彙總前的折現率表、Thesis!A2 說明移到驅動章。共用產生器新增可選功能：日期格以 YYYY-MM-DD 文字顯示（In 頁五個日期鍵；其他公司沒有日期格，輸出不變）。
- 與 v0.3.5 比對：README 以外 10 頁逐格 0 差異；加權目標價 384.10、放空不變。敏感度不重跑：`data/sens_v035_20261104.json` 改名 `data/sens_v036_20261105.json`；v0.3.5 成品移 `dist/_待刪除/`。

## v0.3.5（2026-11-04，Tokenomics 快照 v21.9 → v21.10，基準數值不變）
- 快照 `data/tokenomics_snapshot_v21.10.json`（`20261104_Tokenomics_v21.10.xlsx`，master eded8e6；名稱清單 `data/tokenomics_names.txt` 不變；`--check` 通過）；`build_microsoft.py --snap` 預設改 v21.10。v21.9 快照、`sens_v034_20261104.json`、v0.3.4 成品移 `_待刪除/`；v21.4、v21.7 快照仍保留（`attrib_v031.py` 以 `--snap` 重現拆解）。
- v21.10 只改 `IF_NonNV_Out_TPU`／`IF_NonNV_Out_Trn`，本模型的快照不含這兩個名稱（讀 `IF_NonNV_*_Maia`，不變），新舊快照的值逐一相同。
- 與 v0.3.4 逐格比對 24 格不同，全是版本與快照說明文字（In A1、A95、H 欄來源、Sens F36:F37 說明、Sources 第 6 列）；數值格 0 差異，加權目標價 384.10、放空不變。敏感度整本重算 53 筆重跑，逐一相同（`data/sens_v035_20261104.json`）。報告 `docs/reports/20261104_microsoft_v0.3.5_TK快照v21.10.md`。

## v0.3.4（2026-11-04，Tokenomics 快照 v21.7 → v21.9，基準數值不變）
- 快照 `data/tokenomics_snapshot_v21.9.json`（`20261103_Tokenomics_v21.9.xlsx`，master 7276259；名稱清單 `data/tokenomics_names.txt` 不變）；`build_microsoft.py --snap` 預設改 v21.9。`data/` 內 v21.4、v21.7 快照**保留**（`attrib_v031.py` 以 `--snap v21.4／v21.7` 重現 v0.3 → v0.3.1 拆解）。
- 契約名稱值的變動只在高成本欄（v21.8）；本模型讀的 `IF_NonNV_*_Maia` 不變。
- 與 v0.3.3 逐格比對：只有 In A1、A95（快照說明）、F96:F99（`tk_hold_*` 敏感度高）、H 欄來源文字、Sources 第 6 列、Sens 第 36–37 列不同；活公式 Sens 案例與其餘各頁全部相同（加權目標價 384.10、放空不變）。
- Sens 變動：`he_all` 高 462.76 → 440.13（上檔 −12.9% → −17.1%，仍跨到中立）。敏感度結果 `data/sens_v034_20261104.json`（其餘整本重算案例逐一相同）。報告 `docs/reports/20261104_microsoft_v0.3.4_TK快照v21.9.md`。

## v0.3.3（2026-10-30，敏感度覆蓋，基準數值不變）
- 依個股模型原則第 9 節：Sens 頁加 G 欄「鍵」（原有 32 個活公式案例與 19 個判斷替代逐列填鍵）；In 頁第 96–99 列 `tk_hold_*`（IF_HoldEcon 四世代）D／F 補 Tokenomics v21.7 快照的低成本／高成本欄（Interface!13 列；基準不變）；In 頁 50 個有低／高區間的鍵全部有案例或合格豁免（`tools/sens_coverage/check_sens_coverage.py --require-keycol` 通過，`make_v03.sh` 每次執行）。
- 新增 27 個鍵的整本重算案例（`sens_v033.py` → `data/sens_v033_20261030.json`，52 列），含 **IF_HoldEcon 四世代同時（鍵 `he_all`）**：低 352.04（−8.3%，上檔 −33.7%，放空）、高 462.76（+20.5%，上檔 −12.9%，**跨到中立**）。活公式案例照舊；完整引擎區塊由列 80 移到列 200（數值逐格相同）。
- 13 個對照列（`cons_tp`、`cons_tp_mb`、`cons_rev27`、`cons_eps27`、`g_q1_*` 4 個、`g_q2_rev／ai／dc／az`、`old_tp`）說明欄寫「只作對照，不進模型」（核對程式追引用鏈通過）；`g_q2_win` 由判斷替代「Windows 套用指引」一案涵蓋。豁免 3 個（Sens!A2，含理由與重算值）：`ai_net_open`、`m_ai`（重算變動 0）、`bb_g`（< 0.001%）。HTML 龍捲風圖略過變動 < 0.01% 的整本重算案例（展開表照列）。
- 與 v0.3.2 逐格比對：In 頁只有 A1（版本字串）、第 96–99 列 D／F／I、13 個對照列 I 欄不同；Sens 頁不同；其餘 9 頁全部相同（加權目標價 384.10、放空不變）。報告 `docs/reports/20261030_MAG_Oracle_敏感度補齊.md`；v0.3.2 成品移 `dist/_待刪除/`。

## v0.3.2（2026-10-29，呈現修正，數值不變）
- In 頁 D／F 欄名改「敏感度低」「敏感度高」（只用於 Sens 頁，不進目標價）；頁尾新增「J. 情境驅動對照」：長約占比、k、Copilot 淨增、情境權重，情境 A／B／C 三欄並排（只引用上方各列）。
- HTML「主要輸入」拆成「情境驅動」與「單一參數區間」兩張表。規格見 `mag/README.md`；核對 `mag/check_in_layout.py`。
- `tools/in_layout` 第 5 項警告 2 格（Val!C38、D38）：市場共識目標價區間 `cons_tp_lo`／`cons_tp_hi`，非情境值，保留。
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
