# 公司模型 v0.1 建置工作單（AI半導體 Project，共用）

你負責建一家公司的收支與估值模型 v0.1，放在 repo `company-models` 的新資料夾 `<slug>/`。你在自己的 git worktree 裡工作（目前目錄就是 repo 根目錄）。今天是 2026-10-25（台北）。

## 範圍與禁止事項

- **暫存檔：** 自己的暫存腳本與資料一律放 `/tmp/claude-0/-home-claude/2c6c69f9-cd8c-5be7-8f8a-fd54291db1bc/scratchpad/work_<slug>/`（自建），**不得**直接寫進 scratchpad 根目錄，也不得用與 Python 標準模組同名的檔名（inspect.py、json.py 等）。

- **只新增 `<slug>/` 一個資料夾**（`build_<slug>.py`、`README.md`、`html.json`、`dist/`）。**不得修改**其他資料夾、`tools/`、`markets/`、`.github/`（CI 設定由主串統一加）。
- 不做 `git push`、不開 PR。完成後在 worktree 內 `git add <slug>/ && git commit -m "<公司> 收支與估值模型 v0.1"`（commit 訊息結尾加兩行：`Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>` 與 `Claude-Session: https://claude.ai/code/session_01CRVW4qsiiKVECHPubNVS8F`）。
- repo 是公開的：**不得放任何 Andy 個人情境**（部位、可否投資、基金、個人看法）。
- **本批不做先驗**（Andy 2026-10-25 暫停先驗測驗）：Val 頁**不要**放「Andy 先驗：合理 P/E」輸入與「對照：以 Andy 合理 P/E」那一節；README「估值」段寫一句「本版未做先驗，無 Andy 合理 P/E 對照」。
- 所有檔案文字用**繁體中文（台灣用語）**，風格與樣板相同。

## 樣板

先完整讀樣板資料夾（README、`build_*.py`、`html.json`），照它的工作表結構（README、In、Q、Units、PL、Val、Recon、Sens、Sources）、三情境（A／B／C）、列標籤寫法與 `html.json` 七章結構複製，再換成這家公司的驅動因子。**Val 頁的列標籤沿用樣板原文**（例如「規則值：…」開頭的列、「現價 ÷ FY2027 EPS（模型）」「現價 ÷ FY2028 EPS（模型）」），因為 `markets/invest_map/extract_models.py` 以列標籤抽數；讀一下該檔確認它抓哪些標籤。

上游只讀、不改：`markets/network/dist`、`markets/equipment/dist`、`markets/chip_bridge/dist`、`tsmc/dist`（情境顆數係數）、`tools/cycle_position.py`、`tools/valuation_rule.py`。算力假設（部署 GW、每 GW capex 等）只能從這些成品讀，不自行重估。設備股用統一循環規則（`tools/cycle_position.py` 讀 `markets/equipment` 見頂年）；非設備股比照樣板的循環位置判定。

## 預估期間：三情境延伸到 FY2032（必做；Andy 2026-10-25 11:18、11:20）

- PL 頁三情境（A＝順利、B＝平平、C＝不如預期；若原本是「顆數情境」就在 README 註明對應）延伸到 **FY2032**。上游市場層（`markets/*`、`tsmc` 顆數）只到 2029／2030，**不得修改上游**；上游結束後用「收斂法」：
  - 營收年增率由上游最後一年的年增，線性收斂到 FY2032 的長期成長：A＝`tp_g`＋2pt、B＝`tp_g`、C＝`tp_g`−2pt。
  - 淨利率由最後一年線性收斂到中段利潤率：A＝最後一年、B＝最後一年與十年平均的中點、C＝十年平均（找不到十年平均用最後一年 × 0.8，標 Assumed）。
  - 收斂年度在 PL 頁換底色，README 寫明「FY2030 起為收斂法，非驅動因子推導」。

## 目標價＝DCF 與 Comparable 加權（必做；Andy 2026-10-25 11:28 取代先前的目標價寫法）

Andy：「目標價要包括 DCF 跟 comparable 兩個的加權平均……顯示一個是 DCF 再加 comparable 的目標價，另外一個就顯示只有 comparable 的目標價，comparable 要用市場主流的 comparable……最主要的還是要看 DCF 與 comparable 加權平均的目標價。」

Val 頁最前面加「0. 目標價」一節，列標籤**逐字照抄**；三情境欄位 B／C／D＝A／B／C，單一值放 B 欄：

**0.1 DCF（股東權益現金流，FY2027–FY2032 明確期＋FY2032 後永續）**
| 列標籤 | 內容 |
|---|---|
| `情境 EPS（FY2027–FY2032）` | 每情境一列或一個小表，引用 PL 頁 |
| `股東現金流＝EPS ×（1 − tp_g ÷ tp_roe）` | 再投資率＝g／ROE；記憶體、設備等資本密集公司若模型有資本支出與折舊，改用（淨利＋折舊－資本支出），並在方法欄註明 |
| `DCF 情境每股價值（折現率 tp_r）` | ＝ Σ 明確期現金流 ÷ (1+r)^t ＋ 永續價值 ÷ (1+r)^6；永續價值＝FY2032 現金流 ×(1+tp_g)÷(tp_r−tp_g)；基準日＝2026 年底，t＝1 對 FY2027 |
| `DCF 目標價（情境加權）` | ＝ Σ tp_wA／wB／wC × 情境價值 ×(1+tp_r)（推到約 12 個月後） |

**0.2 Comparable（市場主流同業前瞻 P/E；記憶體用 P/B）**
| 列標籤 | 內容 |
|---|---|
| `同業與前瞻 P/E` | 3–6 家市場（券商報告）最常拿來比的同業，各列名稱、代號、前瞻 12 月 P/E、日期、來源（放 Val 頁小表；Interested-party 標記） 。**A 股與其他市場分開（Andy 11:53）：非 A 股公司的同業不放 A 股；A 股公司的同業只用 A 股。** 若想參考另一市場，可在小表下方「並列參考」另列，但不進目標價 |
| `同業前瞻 P/E 中位數` | 公式 |
| `12 個月後的前瞻 EPS（情境加權）` | ＝ Σ 權重 ×（2/12 × FY2027 EPS＋10/12 × FY2028 EPS）（會計年度非日曆年時依月份調整並註明） |
| `Comparable 目標價` | ＝ 同業中位數 × 上列 EPS |

**0.3 結論（HTML 估值章第一個區塊，兩個目標價並列）**
| 列標籤 | 內容 |
|---|---|
| `目標價（DCF 與 Comparable 加權）` | ＝ tp_w_dcf × DCF 目標價 ＋ (1−tp_w_dcf) × Comparable 目標價 |
| `Comparable 目標價（市場主流同業法）` | 引用 0.2 |
| `市場共識目標價` | 賣方平均目標價、家數、日期、來源；找不到填「未找到」 |
| `上檔空間：加權目標價` | 公式 |
| `上檔空間：Comparable 目標價` | 公式 |
| `與市場共識目標價的差距：EPS 因素` | ＝（我們的前瞻 EPS ÷ 共識前瞻 EPS − 1）；共識 EPS 取 Recon |
| `與市場共識目標價的差距：倍數因素` | ＝ 市場共識目標價 ÷ 共識前瞻 EPS 隱含的倍數 vs 同業中位數 |
| `價格帶（DCF 不如預期～順利）` | C 情境價值～A 情境價值（推到 12 個月後） |
| `歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）` | 現價 ÷ 各情境 FY2028 EPS 在 2016–2025 年底 P/E 分布的分位（PERCENTRANK）；只作檢查，不進目標價 |
| `目標價方法` | 一句：DCF 參數、同業名單、主要偏誤方向 |

In 頁最後新增黃底輸入（Assumed，「Andy 可改」）：`tp_wA`=0.3、`tp_wB`=0.5、`tp_wC`=0.2、`tp_r`=0.10（美、台、日、韓同一值，若公司有明顯不同風險可在方法欄說明但不改值）、`tp_g`=0.04、`tp_roe`＝近三年平均 ROE（Derived；找不到 0.20）、`tp_w_dcf`=0.5（Andy 11:53 定 50／50）。若 tp_roe ≤ tp_g，再投資率上限 0.8。

- **DCF 下限（Andy 12:22 定）：** 任一情境的 DCF 每股價值 < 0 時取 0（股東有限責任），**仍照 tp_w_dcf 加權**，不改用 100% Comparable、不標不適用；Comparable 目標價另列供比較。
- Sens 頁加三列：tp_r ±1pt、tp_g ±1pt、tp_w_dcf 0.3／0.7 對加權目標價的影響。
- 舊的「規則值」節保留不刪（invest_map 抽數用），但不再叫目標價。
- 回報加：DCF 三情境價值、DCF 目標價、同業名單與中位數、Comparable 目標價、加權目標價、市場共識目標價與差距拆解、歷史分位。

## 研究紀律

- 用 WebSearch／WebFetch 取最新實績（最近一季／半年財報，截至 2026-10-25 已公布者）、公司指引、共識、現價（註明日期）、股數、十年年底 P/E（2016–2025，至少一個來源，盡量兩個）。
- 每個輸入加標記：Verified（附出處）、Interested-party（公司、賣方、媒體——寫明立場）、Analogy、Assumed、Derived。Analogy／Assumed 不當點估計用：放情境或 Sens 區間。
- 找不到就寫「未找到」，不要編。共識、券商數字只放 Recon，不作校準。
- 關鍵數字（營收、毛利率、份額）盡量兩個獨立來源；只有一個來源要註明。

## 建置與驗證（全部通過才算完成）

1. `python3 <slug>/build_<slug>.py /tmp/<slug>_raw.xlsx`
2. `soffice --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir <slug>/dist /tmp/<slug>_raw.xlsx`，再把檔名改成 `<slug>/dist/20261025_<公司中文簡稱>模型_v0.1.xlsx`（dist 只能有一個 xlsx）。
3. 錯誤格 0、重建一致：重跑步驟 1–2 到另一個暫存位置，`python3 tools/xlsx_diff.py <新重算檔> <dist 檔> --tol 1e-9 --limit 20` 必須通過。
4. `python3 tools/company_html/build_company_html.py <slug>`，再 `python3 tools/company_html/check_company_html.py <slug>` 必須全過。
5. 自己抽查：PL 三情境的營收、毛利率、EPS 是否合理；FY2025 實績是否與財報一致；EPS 量級是否與共識同級（差距超過 30% 要在 README 說明原因）。

## 回報（給主串，300 字以內，不要貼檔案內容）

- 分支名稱與 commit hash；dist 兩個檔名
- 一行命題與核心問題
- 結果表：FY2026（錨點）、FY2027、FY2028 的 EPS（A／B／C）、共識 EPS、現價（日期）、現價 ÷ FY2027 EPS、規則值隱含價與相對現價
- 最敏感的 2–3 個變數
- 主要 Assumed 輸入與「未找到」項目
- 下次應更新的事件（財報日期）
