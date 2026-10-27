# alphabet/ — Alphabet（GOOGL）個股模型 v0.3（命題＋營收分部）

**命題（一句話）：** AI 改變搜尋之後，Google Search 的變現能否守住，同時 Google Cloud（含 TPU）能否把翻倍的資本支出變成營收；加計 AI 折舊與 2026 年股權發行後，每股價值是否高於現價。

**核心問題：** Search & other 在 C 情境（AI 助理分流、每次查詢變現下降）下 FY2028 營收與利益率會掉多少，Google Cloud 的 AI 營收能否抵銷；加計 AI 折舊與股權發行後的每股價值。

**投資建議：強力放空**（依五級固定規則：加權目標價 US$135.28，比現價 316.80〔2026-10-26 收盤〕低 57.3%）。**這個結論取決於一個待 Andy 決定的口徑**（對照表 r1 C16）：自用 AI 增量投資若以「價值中性」計（讀法 2），目標價 211.73（−33.2%，放空）。收支關卡（FY2026–FY2030 自由現金流扣股利合計）：A 137.4／B 56.2／C −93.9（C 需舉債 44.8 並減少回購）。

## 重點觀察（數字取自成品 Excel）
1. 目標價 135＝DCF 57（情境加權）與分部加總 214 各半；情境 A／B／C 為 157／139／106；市場共識目標價 429（StockAnalysis，2026-10-26）。DCF 偏低的原因是 AI 資本支出到 FY2032 仍每年 238–416 十億美元、明確期無槓桿自由現金流幾乎全為負，價值幾乎全在終值。
2. C 情境 FY2028 Search & other 290.1（比 B 少 14.5%），Google Services 營業利益率 32.1%（B 35.8%）；同年 Google Cloud 營業利益 62.1，比 FY2025 增加 48.2，是 C 情境 Services 營業利益損失 34.2 的 141%——Cloud 的成長足以抵銷搜尋被侵蝕的利潤。
3. AI 折舊進損益：情境 B 新增 AI 折舊 FY2028 62.0（FY2032 205），合併營業利益率由 FY2025 32.0% 降到 FY2028 24.2%；FY2028 營業利益比 MarketScreener 共識低 36%（營收低 7%）。差距主要來自自用 AI（占 AI MW 70%）的折舊與營運成本沒有對應的額外營收。

## 結構（三層）
| 層 | 內容 | 頁 |
|---|---|---|
| 命題模組 | AI 總 MW（評價日 6,337、每年 +4,160）× 對外比例 30% → **Cloud AI**＝Σ 世代 MW × Tokenomics `IF_HoldEcon` × k（混合 1.201）；同一套 MW → AI 資本支出（IT＋自建機房＋6 年汰換；2H26 以指引 200 為下限）→ 新增 AI 折舊；自用 AI 增量成本（營運＋SpaceX 租金＋折舊）60% 進 Services、40% 進 Alphabet 層級 | Thesis |
| 營收分部 | **Google Services**（Search & other、YouTube ads、Google Network、Subscriptions, platforms, and devices）＋**Google Cloud**（Cloud AI＋TPU 系統銷售＋Cloud 其他）＋**Other Bets**＋避險損益；各線 2H26 年增＝近四季年增，FY2030 收斂到長期值；**情境軸＝Search**（A 查詢回流 5%＋變現 +10%；B 收斂；C 分流 20%＋變現 −10%，FY2030 到位） | Seg |
| 損益表 | 分部 EBITDA（Services 1H26 EBITDA 率 49.4%＋Search 差額 × 增量利潤率 70%；Cloud＝Cloud AI EBITDA＋TPU × 25%＋Cloud 其他 × 28.2%）− 折舊（基礎 28.4／年按營收分攤＋新增 AI＋新增非 AI）→ 分部營業利益 → 合併 → 利息、稅 → EPS | PL |

之後：CF（自由現金流＝營運現金 − 資本支出 − 融資租賃本金；2026 股權發行 49.6 與暫停回購〔C4〕；現金 → 減回購 → 新債）、Val（DCF 50%＋分部加總 50%；情境 25／50／25）、Recon、Sens、Sources。

## 驅動 → 推導對照表
| 頁 | 輸入（In 頁鍵） | 推導量 | 因果方向 |
|---|---|---|---|
| Thesis | pace、ext_sh、mw_tot0、k（kl_px、ks_px、long_sh）、Tokenomics IF_HoldEcon／OpexGW／CapexIT | Cloud AI 營收與 EBITDA、AI 資本支出、新增 AI 折舊、自用 AI 增量成本 | MW → 收入；MW → 資本支出 → 折舊 → 分部營業利益 |
| Seg | 近四季年增（Q 頁）、gl_*、conv_n、dA／mA／dC／mC、tpu_*、ob_rev_* | 四條產品線、三個分部、合併營收；非 AI 資本支出與折舊 | 近四季年增 → 長期值；Search＝B 路徑 ×(1 − 分流)×(1＋變現變動) |
| PL | m_svc、inc_m、m_co、tpu_m、abc_rate、svc_sh、da_base、tax | 分部 EBITDA、分部營業利益、合併營業利益、EPS | 分部營收 → EBITDA − 折舊 → 營業利益 → EPS |
| CF | sbc_r、wc_*、dps、bb_t、min_cash、mat_*、anth_* | 自由現金流、股利、回購、新債、收支關卡 | 淨利＋折舊 − 資本支出 → 自由現金流 → 股利 → 回購／舉債 |
| Val | rf、beta、erp、kd、g_tv、ronic、同業倍數、持股、tp_w*、w_dcf、rd2 | DCF、分部加總、目標價、上檔空間、投資建議 | 無槓桿 FCF 折現＋分部 EBITDA × 倍數 → 情境加權 → 50／50 |
| Sens | In 頁低／高值 | 加權目標價變動（41 案，整本重算） | 一次改一個輸入 |

## 與 v0.2.2（收支模型）的差異
- 營收主軸由「AI 雲端 MW × 每 MW 收入＋其餘成長率線」改為官方分部；Cloud AI 只是 Google Cloud 的子列（子列合計＝分部，核對列 0）。
- 情境軸由 AI 容量（保守／基準／積極）改為 Search；AI 容量三情境相同（C10 基準速度），改列 Sens。
- AI 資本支出接到折舊與分部營業利益（不只接現金）。
- 評價由 DCF 45%＋EV/EBITDA 55%（錨定 2029）改為 DCF 50%＋分部加總 50%（各分部同業倍數）；評等改五級（上檔空間）。
- 融資瀑布、站點、租賃承諾、RPO 等 CRWV 頁面移除；舊引擎保留在 `legacy_v0.2/`（未刪），舊成品在 `dist/_待刪除/`。
- 目標價：v0.2.2 讀法 1 加權 147.54 → v0.3 135.28（−12.3：評價權重 −7.8、現價 +0.1、其餘結構與口徑 −4.5；Recon 頁第 3 節）。

## 檔案
| 檔案 | 內容 |
|---|---|
| `build_alphabet.py` | Python → Excel 產生器（openpyxl，全部活公式；藍字輸入、黑字公式、黃底判斷題） |
| `sens_v03.py` | 敏感度：一次改一個 In 頁輸入、整本 LibreOffice 重算，結果寫 `data/sens_v03_20261027.json`（建置時寫入 Sens 頁與 Recon 拆解） |
| `check_v03.py` | 成品核對：公式錯誤 0、Q／Seg／PL 核對列 0、敏感度未過期、固定列標籤存在 |
| `gen_html_json.py` | 由列號對照產生 `html.json`（勿手改 html.json） |
| `make_v03.sh` | 一次跑完：建置 → 重算 → 核對 → segment_check → HTML（`SENS=1` 重跑敏感度） |
| `segments.json` | 官方分部宣告（`tools/segment_check/check_segments.py alphabet` 通過） |
| `html.json` | HTML 展示層設定（`tools/company_html/`） |
| `dist/20261027_Alphabet個股模型_v0.3.xlsx`／`.html` | 成品（LibreOffice 重算，公式錯誤 0；HTML 檢核通過） |
| `data/` | 事實總帳（SEC 一手資料）、Tokenomics v21.4 快照、k 證據、共識、同業倍數、敏感度結果 |
| `docs/reports/20261027_alphabet_v0.3_分部架構.md` | 本版報告（結論、預設、差異拆解、資料缺口、檢查結果） |
| `legacy_v0.2/` | v0.2.2 MAG 共用引擎（segA–E.js、build_xlsx.py、company.json、scripts/ 等；保留不刪） |

## 建置
```
bash alphabet/make_v03.sh                 # 沿用已提交的敏感度結果
SENS=1 JOBS=6 bash alphabet/make_v03.sh   # 重跑敏感度（約 40 次整本重算）
```
`tools/segment_check/check_segments.py` 尚未在 main 時，以 `SEGCHECK=<路徑>` 指定（來源分支 `claude/geguxing-model-principle`）。

**工作表**：README → In → Q → Thesis → Seg → PL → CF → Val → Recon → Sens → Sources。金額十億美元（US$B）；日曆年＝會計年度；FY2026＝1H 實績＋2H 模型；明確期到 FY2032。
