# meta/ — Meta Platforms（META）個股模型 v0.2.0（命題＋官方營收分部 → 損益表；2026Q3；v0.1.6 驗收採用＋原則第 10 節第 1–55 條＋DCF 標準規格）

**建置：** `bash meta/make_v02.sh`（`KREQ=1 SENS=1` 重跑 k 梯度反解與敏感度）。成品 `dist/20261111_Meta個股模型_v0.2.0.xlsx`／`.html`；v0.1.6 成品移 `dist/_待刪除/`。報告 `docs/reports/20261111_meta_v0.2.0_驗收採用.md`（目標價逐項歸因、原則逐條清單、Wattonomics 對照）。

**版本表（由建置成品產生，不手寫；原則第 10 節第 51 條）：**

<!-- version-table: dist/_待刪除/20261106_Meta個股模型_v0.1.6.xlsx｜dist/20261111_Meta個股模型_v0.2.0.xlsx -->
| 項目 | v0.1.6 | v0.2.0 |
|---|---|---|
| 現價 | 605.08 | 605.08 |
| DCF 目標價 | 792.59 | 694.51 |
| Comparable（分部加總）目標價 | 895.60 | 865.26 |
| 加權目標價（上檔空間） | 844.10（+39.5%） | 779.88（+28.9%） |
| 投資建議 | 買進 | 買進 |
| 情境目標價 A／B／C | 933.27／844.10／754.93 | 864.71／779.87／695.07 |
| 讀法 1 目標價 | 428.70 | 431.44 |
| 收支關卡 A／B／C（FY2026–FY2030） | -19.6／15.0／49.0 | 7.6／40.5／72.7 |
<!-- /version-table -->

**v0.2.0（2026-11-11）：** 依 Meta v0.1.6 驗收報告採用規劃書（Andy 2026-11-11 22:09 確認，J1–J3 選 (a)）與工作單 `docs/workorders/20261111_Meta_v0.2.0_驗收採用.md`：
- 讀法 2 仍為主值（不需驗收清單 X1）；新增 k 梯度、支持現價所需 k、評等翻轉 k（`kreq_v020.py`，Recon 第 6 節）、隱含 AI 歸因比例（6c）、兩法依賴讀法 2 的占比（Val）與 HTML 橫幅一行（版面規格 v1.1）。
- 在建工程對應 MW 自 Q4 26 起分 4 季線性轉入在役（J2；`cip_q`，Sens 2／8 季、不另增 MW）；容量口徑分「在役 MW」與「支出口徑 MW（含在建）」（原則第 53 條）。
- 目標時點＝`EDATE(px_date,12)`；README 首列分列資料截止日、成品日、股價基準日；Val 並列折現期間與持有期間。
- 評價股數計入未歸屬 RSU（DCF 標準規格第 3 節；Sens `dil_inc`＝0 為 v0.1.6 口徑）；評等邊界改為 > +50%、> +25%、≥ −25%、≥ −50%。
- 融資可行性檢查（CF；`lev_max`、`icr_min`）、融資受限 Sens（`bd_delay`）、期間流動性並列收支關卡（原則第 55 條）；經濟汰換壓力 Sens（`ec_press`，待 Tokenomics 登錄）。
- Wattonomics 對照（Recon 第 7 節；個股不改 Wattonomics）；同業子群中位數（Recon 第 8 節，只作展示；同業換到現價日未完成）。
- Tokenomics 快照 v21.10 → v22.0（本模型讀的名稱值不變）；build_id、輸入指紋與模型狀態（原則第 1 條）；In、Q 頁單一標記與來源誘因欄。
- 文字中的數字由 xlsx 公式組字；本檔不再手寫目標價（見上方版本表與成品 README 頁）。

**v0.1.6（2026-11-06）：** 只換現價（Andy 2026-11-06 07:32「請所有個股都換成十一月五號的收盤價」）：In 頁 `px` 612.49（10-28）→ 605.08（11-05 收盤）。WACC 用評價日股價 `px_val`（9-30 收盤 725.18），不受影響；加權目標價 844.10 不變，上檔空間 +37.8% → +39.5%，買進不變。敏感度整本重跑 `sens_v016_20261106.json`（目標價逐案相同；評等在工作表內依現價即時判斷）；v0.1.5 成品移 `dist/_待刪除/`。

**v0.1.5（2026-11-05）：** 只改展示層，數值格 0 差異。README 頁新增「展示層」「命題回覆」兩段（公式帶數字），HTML 套用個股 HTML 版面規格 v1.0（`docs/plan/個股HTML版面規格.md`；`tools/company_html/relayout.py`，參數分「廣告本業／Reality Labs／自用 AI 算力／全公司」四塊）。敏感度沿用 v0.1.4 結果改名 `sens_v015_20261105.json`；v0.1.4 成品移 `dist/_待刪除/`。

**v0.1.4（2026-11-04）：** Tokenomics 快照 v21.9 → v21.10（`data/tokenomics_snapshot_v21.10.json`，`20261104_Tokenomics_v21.10.xlsx`，master eded8e6；名稱清單同 Alphabet；`--check` 通過）。契約名稱值的變動只在 `IF_NonNV_Out_TPU`（本模型不讀）。與 v0.1.3 逐格比對 21 格不同，全是版本與快照說明文字；數值格 0 差異，加權目標價 844.10、買進不變；敏感度 84 案重跑，逐一相同（`data/sens_v014_20261104.json`）。v0.1.3 成品、`sens_v013` 與 v21.9 快照移 `_待刪除/`。報告 `docs/reports/20261104_meta_v0.1.4_TK快照v21.10.md`。

**v0.1.3（2026-11-04）：** 基準數值不變（加權目標價 844.10，買進）。
- **協議應計驗證：** Recon 5 新增驗證列。協議應計可能大於 11.89，因為協議總付款含約 0.5 的隱私案。若如此，legal_oth 最多高估 0.885（名目上限 12.775 − 11.89）。這個情況只列入 Sens（In `ag_acc_x`），結果 844.45（+0.35）。
- **受限現金口徑：** 與 Alphabet、Amazon 一致，三家的淨現金都不含受限現金。把 10.80 採購保證金計入淨現金的情況只列入 Sens（In `esc_in`），結果 848.40（+4.30）。
- **檔案：** 敏感度結果在 `data/sens_v013_20261104.json`；v0.1.2 成品與 `sens_v012` 已移到 `_待刪除/`。報告見 `docs/reports/20261104_meta_v0.1.3_小修.md`。

**v0.1.2（2026-11-04）：** Tokenomics 快照 v21.7 → v21.9（`data/tokenomics_snapshot_v21.9.json`，master 7276259；名稱清單同 Alphabet）。契約名稱值的變動只在高成本欄（v21.8）與 `IF_NonNV_Out_TPU`（本模型不讀）。與 v0.1.1 逐格比對：只有版本與快照說明文字、In 頁 F90:F93（`he_*` 敏感度高）、Sens 第 28–29 列不同；加權目標價 844.10、買進與其他數值全部相同。Sens `he_all` 高 1,054.79 → 1,000.38（上檔 +63.3%，強力買進，與舊版同級）；`he_all` 標籤改由快照取值。v21.7 快照、`sens_v011_20261030.json`、v0.1.1 成品移 `_待刪除/`。報告 `docs/reports/20261104_meta_v0.1.2_TK快照v21.9.md`。

**命題（一句話）：** Meta 的 AI 資本支出能否由廣告現金流支撐，以及自用算力（GPU＋自研 MTIA）能否帶來足以回本的廣告變現提升。投資建議、目標價與各項數字見成品 README 頁（公式組字）與上方版本表。

## 結構（三層）
| 層 | 內容 | 頁 |
|---|---|---|
| 命題模組 | 自用 AI 容量：支出口徑 MW（含在建）與在役 MW 分開命名。期初機隊＝FY2023–2026 年 9M 的 AI 資本支出，按當年錨世代的每 MW 資本支出（Tokenomics IF_CapexIT＋自建比例 × IF_CapexFacility）換算。評價日支出口徑 MW（含在建）與評價日在役 MW 分開命名（原則第 53 條；數值見成品 Thesis 頁）。<br>Q4 26 新增 MW＝(2026 指引中點 − 9M 實績 − 非 AI) ÷ 每 MW（C10 公式解）。FY2027＝(情境 FY2027 資本支出〔M6〕− 非 AI) ÷ 每 MW。FY2028 同速，FY2029 起收斂。<br>MW 再推出新增 AI 折舊、自用 AI 增量營運成本（相對評價日在役機隊）與讀法 2 影子效益（× `k_ai`，主值 1）。<br>**v0.2.0（J2）**：在建工程對應 MW 自 Q4 26 起分 `cip_q`（4）季線性轉入在役；Sens 2／8 季與不另增 MW（`cip_ai`＝0）。Alphabet、Amazon、Microsoft 下次改版同規則（原則第 54 條） | Thesis |
| 營收分部 | **Family of Apps**（子列 **Advertising**、**Other revenue**）＋**Reality Labs**。<br>Advertising＝曝光數 × 每則平均價格。兩者的近四季年增各自在 FY2030 收斂到長期值；起始價格年增由廣告金額年增反推。<br>**情境軸＝AI 對每則價格的提升**：A +10%、C −10%，相對 B 路徑，FY2030 到位 | Seg |
| 損益表 | FoA EBITDA＝(FoA 營收 − 價格情境差額) × 9M26 EBITDA 率（55.1%，排除一次性）＋差額 × 85% − 自用 AI 增量營運成本。<br>減去折舊（基礎 28.6／年＋新增 AI＋新增非 AI）得 FoA 營業利益。RL＝指引與法說路徑（2026 −18.5、2027 −15、FY2032 −5）。<br>由此推到合併營業利益、業外、稅（16%）、EPS（財報口徑／排除一次性／評價口徑） | PL |

之後依序為：
- **CF**：自由現金流＝營運現金 − 資本支出（含融資租賃本金）；協議付款 FY2027 起每年 1.27（10-Q 未載各年金額，等額；第一年前高後低 ag_y1 只進 Sens）；強制轉換特別股（淨額 12.46）、遠期出售；現金 → 回購 → 新債。
- **Val**：DCF 50%＋分部加總 50%；情境 25／50／25。
- **Recon**、**Sens**、**Sources**。

## 驅動 → 推導對照表
| 頁 | 輸入（In 頁鍵） | 推導量 | 因果方向 |
|---|---|---|---|
| Thesis | capex_g、cx27_A／B／C、cx27_k、cx_na、selfb、a28、cip_0、cip_ai、cip_q、bd_delay、ec_press、k_ai、Tokenomics IF_CapexIT／Facility／OpexGW／HoldEcon（he_*，區間取 Tokenomics 低／高成本欄）／DeprIT | 在役 MW、支出口徑 MW（含在建）、AI 資本支出、新增 AI 折舊、增量營運成本、影子效益 | 資本支出 → MW（C10）→ 折舊與營運成本 → FoA 營業利益；MW → 影子效益 |
| Seg | 近四季年增（Q 頁）、gl_impr、gl_price、pA／pC、gl_other、gl_rl、conv_n、g0_q | Advertising（曝光 × 價格）、Other revenue、FoA、RL；非 AI 資本支出與折舊 | 近四季年增 → 長期值；Advertising＝B 路徑 ×(1＋價格情境調整) |
| PL | m_foa（Q 頁）、inc_m、rl_loss26／27／32、rl_da_sh、tax、ot_*、phi_q4 | 分部 EBITDA 與營業利益、合併營業利益、三種 EPS | 分部營收 → EBITDA − 折舊 → 營業利益 → EPS |
| CF | sbc_r、wc_k、ag_pay、ag_y1、ag_rest、dps、bb_t、min_cash、mcps、fwd_cash、fwd_on、r_cash、r_debt、dps_g、lev_max、icr_min | 自由現金流、股利、回購、新債、淨現金、收支關卡 | 淨利＋折舊 − 協議付款 − 資本支出 → 自由現金流 → 股利 → 回購／舉債 |
| Val | rf、beta、erp、kd_sp、g_tv、ronic_sp、同業倍數、ads_mode、rl_mode、tp_w*、w_dcf、legal_oth、ag_acc_x、contg、fin_lease、esc_in、rsu_unv、opt_n、opt_k、dil_inc、px_date、tp_date | DCF、分部加總、目標價、上檔空間、評等 | 無槓桿 FCF 折現＋分部 EBITDA × 倍數 → 情境加權 → 50／50 |
| Sens | In 頁敏感度低／高 | 加權目標價、DCF 目標價、投資建議、收支關卡 B、FY2027 資本支出 B（含 IF_HoldEcon 四世代同時取低／高） | 一次改一個輸入、整本重算 |

**情境驅動對照（In 頁 J 節）：**
- 每則價格調整：A +10%／B 0／C −10%。
- FY2027 資本支出：A 236.6／B 205.7／C 174.9。
- 權重：25／50／25。

In 頁 D／F 欄＝敏感度低／高，只用於 Sens 頁（原則第 8 節）。Recon 頁 3 節的「租用算力年租金滿載值」讀 In 頁 F 欄，屬對照用途，不進主模型。

## 日期與口徑註記
- **法說日期**：逐字稿沒有載明，推定為 **2026-10-21**。依據有三：新聞稿日期 2026-10-21；逐字稿說「today's earnings press release」；8-K 載明 Class C 由 10-21 的重述章程設立。法說早於 10-26 的股權融資。
- **現價**：v0.1.6 起為 2026-11-05 收盤 605.08（Yahoo Finance、Motley Fool、MarketAlleys 一致；`docs/reports/20261105_收盤價_個股換價來源.md`）。v0.1.5 以前為 2026-10-28 收盤 612.49（StockScan、Digrin、TipRanks 相符）。工作單 M8 要 10-29 收盤，且需兩個來源一致。2026-10-30 只找到 ChartRow 一個來源（624.55，+1.97%），故不採用。
- **評價日股價**：2026-09-30 收盤 725.18（TipRanks；StockScan 9/30 −1.84% 相符）。用於 WACC 權益權重與回購價。
- **同業倍數（M7）**：`data/peers_ads_20261030.json` 的六家中位數 9.41×。其中 PINS、TTD、APP、RDDT、SNAP 為 **2026-10-07 基準，未更新**；GOOGL 以 repo 內 Alphabet 資料推算，同基準日。META 自身 8.99×（2026-10-28）只作對照。
- **β**：只取得兩個公開頁面，未能取得 StockAnalysis（讀取被擋）。兩個來源為 ChartRow 1.24（期間未載）與 CoinDataFlow 5Y Monthly 1.185，Blume 調整後中位數 1.142（In 頁 beta＝1.142375）。
- **rf**：沿用 Alphabet v0.3.1 讀取的 2026-10-27 10 年美債 4.99%，因本次重讀財政部網頁被擋。
- **共識**：MarketScreener 年度共識的頁面日期為 2026-10-08，在 Q3 財報之前。

## 檔案
| 檔案 | 內容 |
|---|---|
| `build_meta.py` | Python → Excel 產生器（openpyxl，全部活公式；藍字輸入、黑字公式、黃底判斷題） |
| `facts_2026q3.py` | 四季新聞稿兩讀核對 → `data/meta_facts_2026q3.json`（648 筆、729 項核對一致）。v0.1.1 加第 8 節：完整 10-Q 新事實 22 筆（程式自 `mag/extract/docs/2026Q3/text/meta_2026Q3_10-Q.txt` 擷取數字作第二讀，記 PDF 頁碼）、v0.1 的 14 筆 10-Q 事實以完整版重核、424B5 承銷折扣、follow-up call 13 筆短句 |
| `check_v02.py` | 成品核對（v0.2.0 加：橫幅列、k 梯度快照、模型狀態、README 容量口徑字樣）。須全部為 0：公式錯誤、Q 頁核對列、Thesis 的 C10 與 M6 對帳、Seg 核對、PL 分部對帳（三情境 × 全部年度）、CF 核對。另查敏感度未過期、固定列標籤存在 |
| `sens_v02.py` | 敏感度：一次改一個 In 頁輸入、整本 LibreOffice 重算，結果寫入 `data/sens_v020_20261111.json`（含 DCF 目標價、評等與基準檔輸入指紋）|
| `kreq_v020.py` | k 梯度與支持現價、評等翻轉的 k（整本重算割線法）→ `data/kreq_v020.json` |
| `gen_html_json.py` | 由列號對照產生 `html.json`（勿手改） |
| `make_v02.sh` | 一次跑完（版號 `VER` 預設 0.2.0、日期 `VDATE` 預設 20261111）：事實總帳 →（k 反解、敏感度）→ 建置 → 重算 → check_v02 → check_in_layout → check_sens_coverage → check_segments → HTML → 版本表 → check_governance（含展示層一致性） |
| `display_check.json` | 展示層數字一致性設定（原則第 51 條） |
| `segments.json` | 官方分部宣告（`tools/segment_check/check_segments.py meta` 通過） |
| `dist/20261111_Meta個股模型_v0.2.0.xlsx`／`.html` | 成品（舊版移 `dist/_待刪除/`） |
| `data/` | 事實總帳與兩讀轉錄（`reads/`）、Tokenomics v22.0 快照、共識、同業、敏感度結果 |
| `docs/` | 工作單（v0.1、v0.1.1）、A 段事實總帳報告、v0.1 報告、本版報告 `docs/reports/20261030_meta_v0.1.1.md`、進度檔 |

## 建置
```
bash meta/make_v02.sh                        # 沿用已提交的 k 反解與敏感度結果
KREQ=1 SENS=1 JOBS=2 bash meta/make_v02.sh   # 改了任何輸入或公式後重跑（否則模型狀態 STALE、check_v02 失敗）
VER=0.2.1 VDATE=YYYYMMDD bash meta/make_v02.sh   # 下一版：成品與敏感度檔名隨版號與日期（先把舊成品移 dist/_待刪除/）
```
法說與 follow-up call 逐字稿不入庫（repo 公開）。`facts_2026q3.py` 會讀本機 `/tmp/claude-0/meta_src/Transcript_META_3Q26.txt` 與 `Transcript_follow up call_META_3Q26.txt`（可用環境變數 `META_TR`、`META_TR2` 指定；v0.1.1 起為 PDF 轉文字版，比對不分大小寫），用來核對短句位置。檔案不在時，短句標「未核對」，數值與核對結果不變，但事實總帳中的位置欄文字會不同。完整 10-Q（SEC 公開）入庫於 `mag/extract/docs/2026Q3/meta_2026Q3_10-Q.pdf` 與 `text/meta_2026Q3_10-Q.txt`。

**工作表順序**：README → In → Q → Thesis → Seg → PL → CF → Val → Recon → Sens → Sources。
- 金額單位十億美元（US$B）；日曆年＝會計年度。
- FY2026＝9M 實績＋Q4 模型（欄 E＝9M26A、F＝Q4 26E）。
- 明確期到 FY2032。評價日 2026-09-30；目標時點＝股價基準日＋12 個月（`tp_date`＝`EDATE(px_date,12)`，v0.2.0 起）。
