# alphabet/ — Alphabet（GOOGL）個股模型 v0.3.1（命題＋營收分部；2026Q3 季度滾動）

**命題（一句話）：** AI 改變搜尋之後，Google Search 的變現能否守住，同時 Google Cloud（含 TPU）能否把翻倍的資本支出變成營收；加計 AI 折舊與 2026 年股權發行後，每股價值是否高於現價。

**核心問題：** Search & other 在 C 情境（AI 助理分流、每次查詢變現下降）下 FY2028 營收與利益率會掉多少，Google Cloud 的 AI 營收能否抵銷；加計 AI 折舊與股權發行後的每股價值。

**投資建議：中立**（v0.3.1，2026Q3 季度滾動；依五級固定規則：加權目標價 US$258.21，比現價 324.89〔2026-10-28 收盤〕低 20.5%）。主值＝自用 AI 讀法 2（Andy 2026-10-27 15:03 決定）；讀法 1 目標價 188.31（−42.0%，放空）只作敏感度。放空門檻（−25%）＝243.67：單獨改用讀法 1 或 DCF 權重 70% 會跌破；改用廣告同業中位數倍數、ERP 6%、WACC＋1pt 會逼近。收支關卡（FY2026–FY2030 自由現金流扣股利合計）：A 77.7／B −0.6／C −145.8（v0.3：261.2／180.1／29.8）。B ≈ 0，在 2026 指引區間內正負不定（195：+72.3；205：−74.9；C10 下半年視窗：+113.1）；地圖是否變號待 Andy 決定（預設：容忍帶 |B| < 50 不變號）。v0.3 → v0.3.1 變動拆解（−0.49）：時間推移 +2.48、新季實績 −6.86、指引與 C10 +1.22、Tokenomics／NonNV +2.79、現價 0、其他 −0.12（報告 `docs/reports/20261029_alphabet_v0.3.1_季度滾動.md`）。

## 重點觀察（數字取自成品 Excel）
1. 目標價 258＝DCF 189 與分部加總 327 各半；情境 A／B／C 為 282／262／226；市場共識目標價 429（StockAnalysis，2026-10-26，財報前）。2026Q3：Google Cloud 30.7（年增 103%，比 v0.3 模型值高 17.8%，營業利益率 36.2%），Google Search & other 63.4（年增 12%，比模型低 4.1%）。
2. C 情境 FY2028 Search & other 284.2（比 B 少 14.5%），Services 營業利益少 33.5；同年 Google Cloud 營業利益比 FY2025 增加 71.9（214%）——Cloud 抵得過搜尋侵蝕，且比 v0.3（141%）更明確。
3. 2026 資本支出指引維持 195–205 → C10 速度重解 4,159.7 → 6,228.3 MW／年；資本支出 FY2027 290、FY2028 317（共識 308／352，財報前）；TPU 係數改讀 Tokenomics `IF_NonNV_Hold_TPU` 0.70。情境 B 合併營業利益率由 FY2025 32.0% 降到 FY2028 22.4%（財報口徑；AI 折舊與自用 AI 成本），評價口徑 EPS FY2028 15.45 對財報口徑 10.89。

## 結構（三層）
| 層 | 內容 | 頁 |
|---|---|---|
| 命題模組 | AI 總 MW（評價日 2026-09-30 7,894、每年 +6,228，C10 重解）× 對外比例 30% → **Cloud AI**＝Σ 世代 MW × Tokenomics `IF_HoldEcon` × k（混合 1.201）× TPU 調整（`IF_NonNV_Hold_TPU` 0.70）；同一套 MW → AI 資本支出（IT＋自建機房＋6 年汰換；FY2026＝9M 實績＋Q4 公式＝指引 200）→ 新增 AI 折舊；自用 AI 增量成本（營運＋SpaceX 租金＋折舊）60% 進 Services、40% 進 Alphabet 層級 | Thesis |
| 營收分部 | **Google Services**（Search & other、YouTube ads、Google Network、Subscriptions, platforms, and devices）＋**Google Cloud**（Cloud AI＋TPU 系統銷售＋Cloud 其他）＋**Other Bets**＋避險損益；各線 Q4 26 年增＝近四季年增，FY2030 收斂到長期值；**情境軸＝Search**（A 查詢回流 5%＋變現 +10%；B 收斂；C 分流 20%＋變現 −10%，FY2030 到位） | Seg |
| 損益表 | 分部 EBITDA（Services 9M26 EBITDA 率 48.7%＋Search 差額 × 增量利潤率 70%；Cloud＝Cloud AI EBITDA＋TPU × 25%＋Cloud 其他 × 33.3%）− 折舊（基礎 33.5／年〔2026Q3 × 4〕按營收分攤＋新增 AI＋新增非 AI）→ 分部營業利益 → 合併 → 利息、稅 → EPS | PL |

之後：CF（自由現金流＝營運現金 − 資本支出 − 融資租賃本金；2026 年 9M 股權發行 61.9、發債 85.2 與暫停回購〔C4〕；現金 → 減回購 → 新債）、Val（DCF 50%＋分部加總 50%；情境 25／50／25）、Recon、Sens、Sources。

## 驅動 → 推導對照表
| 頁 | 輸入（In 頁鍵） | 推導量 | 因果方向 |
|---|---|---|---|
| Thesis | pace、ext_sh、mw_tot0、k（kl_px、ks_px、long_sh）、Tokenomics IF_HoldEcon／OpexGW／CapexIT | Cloud AI 營收與 EBITDA、AI 資本支出、新增 AI 折舊、自用 AI 增量成本 | MW → 收入；MW → 資本支出 → 折舊 → 分部營業利益 |
| Seg | 近四季年增（Q 頁）、gl_*、conv_n、dA／mA／dC／mC、tpu_*、ob_rev_* | 四條產品線、三個分部、合併營收；非 AI 資本支出與折舊 | 近四季年增 → 長期值；Search＝B 路徑 ×(1 − 分流)×(1＋變現變動) |
| PL | m_svc、inc_m、m_co、tpu_m、abc_rate、svc_sh、da_base、tax | 分部 EBITDA、分部營業利益、合併營業利益、EPS | 分部營收 → EBITDA − 折舊 → 營業利益 → EPS |
| CF | sbc_r、wc_*、dps、bb_t、min_cash、mat_*、anth_* | 自由現金流、股利、回購、新債、收支關卡 | 淨利＋折舊 − 資本支出 → 自由現金流 → 股利 → 回購／舉債 |
| Val | rf、beta、erp、kd、g_tv、ronic、同業倍數、持股、tp_w*、w_dcf、rd2 | DCF、分部加總、目標價、上檔空間、投資建議 | 無槓桿 FCF 折現＋分部 EBITDA × 倍數 → 情境加權 → 50／50 |
| Sens | In 頁低／高值 | 加權目標價變動（42 案，整本重算） | 一次改一個輸入 |

## 與 v0.2.2（收支模型）的差異
- 營收主軸由「AI 雲端 MW × 每 MW 收入＋其餘成長率線」改為官方分部；Cloud AI 只是 Google Cloud 的子列（子列合計＝分部，核對列 0）。
- 情境軸由 AI 容量（保守／基準／積極）改為 Search；AI 容量三情境相同（C10 基準速度），改列 Sens。
- AI 資本支出接到折舊與分部營業利益（不只接現金）。
- 評價由 DCF 45%＋EV/EBITDA 55%（錨定 2029）改為 DCF 50%＋分部加總 50%（各分部同業倍數）；評等改五級（上檔空間）。
- 融資瀑布、站點、租賃承諾、RPO 等 CRWV 頁面移除；舊引擎保留在 `legacy_v0.2/`（未刪），舊成品在 `dist/_待刪除/`。
- 目標價：v0.2.2 讀法 1 加權 147.54 → v0.3 初版 135.28 → 修正一 251.68 → 修正二 253.30 → 修正三 258.70（讀法 1 → 2 +64.3；見報告「修正三」）。

## v0.3.1（2026-10-29）相對 v0.3 的差異
- Q 頁 2024Q4–2026Q3；FY2026＝9M 實績＋Q4 模型；評價日 2026-09-30、目標日 2027-10-29；2026Q3 事實總帳 `data/alphabet_facts_2026q3.json`（`facts_2026q3.py`，兩讀核對 67 項一致）。
- Tokenomics v21.4 → v21.7（只有 `IF_NonNV_Hold_TPU` 改變）；TPU 係數改讀 Tokenomics（舊 C2 口徑列 Sens）。
- C10 速度依 2026 指引重解（`calib_pace_v031.py` → `data/pace_v031.json`）；WACC 權重與回購價改用評價日股價 344.08。
- 變動拆解 (a)–(f) 在 Recon 頁第 3 節（重算鏈 `sens_v03.py` chain|S1–S4；反事實 `q3cf_v03.py` → `data/q3cf_v03_2026q3.json`）。
- v0.2.2 → v0.3 的差異見 `docs/reports/20261027_alphabet_v0.3_分部架構.md`（v0.3 成品在 `dist/_待刪除/`）。

## 檔案
| 檔案 | 內容 |
|---|---|
| `build_alphabet.py` | Python → Excel 產生器（openpyxl，全部活公式；藍字輸入、黑字公式、黃底判斷題） |
| `facts_2026q3.py` | 2026Q3 新聞稿與簡報 → `data/alphabet_facts_2026q3.json`（兩讀核對，不一致即失敗） |
| `calib_pace_v031.py` | C10：解 AI MW 速度使 FY2026 資本支出＝指引中點 → `data/pace_v031.json` |
| `q3cf_v03.py` | 變動拆解用：v0.3 對 2026Q3 的模型值與隱含輸入 → `data/q3cf_v03_2026q3.json` |
| `sens_v03.py` | 敏感度與拆解鏈：一次改一個 In 頁輸入、整本 LibreOffice 重算，結果寫 `data/sens_v031_20261029.json` |
| `check_v03.py` | 成品核對：公式錯誤 0、Q／Seg／PL／CF／Val 核對列 0、敏感度未過期、固定列標籤存在 |
| `gen_html_json.py` | 由列號對照產生 `html.json`（勿手改 html.json） |
| `make_v03.sh` | 一次跑完：事實總帳 → （C10）→（敏感度）→ 建置 → 重算 → 核對 → segment_check → HTML |
| `segments.json` | 官方分部宣告（`tools/segment_check/check_segments.py alphabet` 通過） |
| `dist/20261029_Alphabet個股模型_v0.3.1.xlsx`／`.html` | 成品（LibreOffice 重算，公式錯誤 0；HTML 檢核通過） |
| `data/` | 事實總帳（SEC 一手＋2026Q3 新聞稿）、Tokenomics v21.7 快照、k 證據、共識、同業倍數、C10 解、敏感度結果；舊版在 `data/_待刪除/` |
| `docs/reports/20261029_alphabet_v0.3.1_季度滾動.md` | 本版報告（結論、變動拆解、新季 vs 模型、指引與 C10、Tokenomics、收支關卡、預設、資料缺口、待決） |
| `legacy_v0.2/` | v0.2.2 MAG 共用引擎（保留不刪） |

## 建置
```
bash alphabet/make_v03.sh                         # 沿用已提交的敏感度結果
SENS=1 JOBS=6 bash alphabet/make_v03.sh           # 重跑敏感度與拆解鏈（約 56 次整本重算）
CALIB=1 SENS=1 JOBS=6 bash alphabet/make_v03.sh   # 改了資本支出相關輸入時，先重解 C10 速度
```
`tools/segment_check/check_segments.py` 尚未在 main 時，以 `SEGCHECK=<路徑>` 指定（來源分支 `claude/geguxing-model-principle`）。

**工作表**：README → In → Q → Thesis → Seg → PL → CF → Val → Recon → Sens → Sources。金額十億美元（US$B）；日曆年＝會計年度；FY2026＝9M 實績＋Q4 模型（欄 E＝9M26A、F＝Q4 26E）；明確期到 FY2032。
