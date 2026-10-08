# CoreWeave 收支模型 v4.7 原始碼包

HTML 與 Excel 都由這裡的原始碼產生。所有路徑皆相對於 repo 根目錄；建置產物寫到 `out/`（不納入版控），交付成品放 `dist/`。HTML 的函式庫模板為 `docs/template_v3_3.html`。

交接檔（業務背景、目前結果、待辦）在 `docs/handoff/`，資料夾內只保留最新版（升版時新增新版、移除舊版）。各任務的完整回報（改了什麼、verify.sh 輸出、關鍵數字對照、待決事項）存於 `docs/reports/`。

**一鍵建置與核對**：`scripts/verify.sh`（純結構修改加 `--vs-dist`，另與 `dist/` 成品比對畫面文字與 Excel 數值、公式；v4.2 起 Excel 以列名稱配對，見「v4.2 核對方式」）。

**換公司**：先讀下方「company.json 欄位說明（換公司填表指引）」。任何一項失敗即以非零代碼結束。

## 檔案
| 檔案 | 用途 |
|---|---|
| segA.js | 資金引擎：情境、產能與收入、CapEx／折舊、租金、期前融資瀑布、連動檢查、敏感性、反向 DCF（rvQ） |
| segB.js | 評價引擎：損益、DCF（0 截斷／選擇權、失效條件）、EV/EBITDA、加權目標價、目標價區間（`targetRange`，v4.1）、評等規則 `rateCall`（門檻讀 `methodology.rating`，v4.2）；同業資料 `lM` 由 `company.json` → `peers` 產生（v4.2）；市場共識對照 `consensusView`、共識逐筆清單 `consensusItems`、來源頁引用句 `consSourceTxtQ`（v4.3）；季度層 `quarterlyView` 與差異原因工具 `findRsnQ`／`fillTokQ`／`reasonSumQ`（v4.4） |
| segC.js | 「損益與評價」頁 UI（損益簡表、Comps、DCF、目標價、反向 DCF；v4.3「市場共識」分頁＝segE 的 `ConsTabQ`） |
| segE.js | 「總結」簡報頁 UI（v3.4 新增）：10 張 1280×720 投影片（v4.3 第 1 張為一頁摘要 `onePageQ`，其餘 9 張為附錄）、簡報模式、列印 CSS；「市場共識」分頁 `ConsTabQ`（v4.3）；「資金模型 → 各期收支 → 季度追蹤」分頁 `QuarterTabQ`（v4.4）；數字即時取自 runFunding／runValuation／reverseDcf／sensitivities |
| segD.js | 頁首與頁籤（總結｜資金模型｜損益與評價）、「資金模型」頁 UI（標題摘要、左欄分類選單、各模組分頁、新債與新股、版本紀錄） |
| tail.js | 共用元件：Excel 式表格 BM、浮動說明 tipQ、分類選單 accQ、兩層選單 navQ、FMODS／VMODS、VLOG |
| app_pretty.js | 原 Oracle 模板的美化程式碼（取其中三段共用 UI） |
| rm_andy.js | Andy 修改過的開關元件（說明改為浮動提示） |
| build_html_portable.py | 組裝 HTML：`python3 build_html_portable.py 4.0 out/輸出.html 2026-09-24 [模板 HTML]`（模板省略時用 `docs/template_v3_3.html`；組裝順序：segA、模板段、segB、模板段、segC、模板段、segD、tail、segE） |
| build_xlsx.py | 產生 Excel：`python3 build_xlsx.py [輸出.xlsx]`（預設 `out/20260926_CoreWeave收支模型_v4_4.xlsx`；版本號、更新日、市價日寫在導覽 A1（v4.3 起讀 vlog.py 與 company.json）；自同目錄讀取 vlog.py、rv_snap.json、company.json 與 `meta.consensusFile` 指定的共識資料檔；群組資訊寫到 `out/outline.json`） |
| fix_outline.py | LibreOffice 重算後，補回 Excel 群組按鈕位置與收合狀態：`python3 fix_outline.py 檔案.xlsx [outline.json]`（預設讀 `out/outline.json`；v3.4 修正：outlinePr 依 schema 放在 tabColor 之後） |
| fix_datatable.py | v4.1：LibreOffice 重算後，把模擬運算表（情境區間）由 `TABLE()` 一般公式還原為 Excel 的 `dataTable` 公式，保留算出的值：`python3 fix_datatable.py 檔案.xlsx` |
| verify_ooxml.py | Excel 嚴格結構檢查（XML 格式、工作表與 sheetPr 子元素順序；v4.1 起另檢查模擬運算表存在且無殘留 `TABLE()`）；v3.4 新增 |
| calendar_q.py | v4.5（5a-1）：期間滾動。由 `company.json` → `calendar`（財年結束月份、最新已申報／已公布季度、模型年數、目標價時點）推算期間標籤、期間長度、評價日、各期期末日與折現年數（以月計）、目標價時點與 EV/EBITDA 折回調整；`build_xlsx.py`、`build_html_portable.py`、`load_engine.js` 都讀同一份推算結果（`COMPANY_DATA.cal`；Excel「輸入與假設」F 區「期間與日期」）。`periods`、`periodYears`、`valuation.optT` 不再存於 company.json。`python3 calendar_q.py` 印出推算結果 | 說明文字中的期間字樣以佔位符書寫（«YTD»、«STUB»、«VMD»、«P0»…，`tokens`），build_xlsx 存檔前、cmp31 查列名稱時換成目前日曆的字樣；目前日曆下換出的文字與 v4.4 相同。期間特有的敘述句與數字放 `company.json` → `texts`、`ytdActual.notes`，隨每季資料更新 || vlog.py | 版本紀錄資料，Excel 的唯一來源（v3.4 起 build_xlsx.py 直接 import，不再內嵌副本）；HTML 的 VLOG 在 tail.js 末尾，需同步 |
| rv_snap.json | 反向 DCF 快照（Excel「評價_反向DCF」與「摘要」頁使用）。v4.5 起由 `scripts/rv_solve.py` 以 Excel 求解產生（verify.sh 步驟 3b；有變動時自動重建 Excel），不再用 JS 引擎（`scripts/rv_snap.js` 已移除） |
| scripts/rv_solve.py | v4.5：反向 DCF 求解。以 LibreOffice（UNO，`scripts/uno_q.py`）開啟建好的 Excel，改「每 MW 年收入倍數」「每 MW 建置成本倍數」「穩態 EBITDA 率」並重算，依與 HTML `reverseDcf` 相同的規則（同一上下界、22 次對半）二分法反解，寫入 rv_snap.json；結果有變動時代碼 3 |
| scripts/uno_q.py | v4.5：以 LibreOffice UNO 驅動 Excel 的共用模組（依 A 欄列名稱找格、改值、讀重算後的值；不存檔） |
| data/consensus_crwv_20260925.json | 市場共識資料檔（v4.3；Andy 查證後提供，只讀，不得增補或修改數字）；路徑寫在 `company.json` → `meta.consensusFile`，HTML 建置時併入 `COMPANY_DATA.consensus`，Excel 直接讀取 |
| xl_diff.py | Excel 比對：`python3 xl_diff.py 舊.xlsx 新.xlsx [--values] [--by-label] [--ignore=工作表,…] [--expect=清單檔]`（W1：新版多出的工作表列為「新增工作表」、不計為差異；v4.5 `--expect`：預期差異清單，每行「工作表!儲存格 原因」，列出但不計入差異數）；（v4.2 `--by-label`：新增列造成位移時以「區段＋列名稱」配對，見「v4.2 核對方式」；v4.3 `--ignore`：兩檔都略過指定工作表） |
| scripts/fields_doc.py | v4.2：產生下方「company.json 欄位說明」的表格（`python3 scripts/fields_doc.py > out/fields.md`）；company.json 有欄位沒寫說明即報錯 |
| xlx.py／cmp31.js | 一致性核對：xlx.py 依情境重算 Excel 並擷取數值（暫存檔在 `out/`，以 `scripts/recalc.py` 重算）；cmp31.js 以 HTML 引擎逐列比對（三情境各 313 項，含 v4.1 目標價區間與判斷句、v4.3 市場共識逐筆與一頁摘要的差距、隱含倍數與句子逐字比對、v4.4 季度層 6 季數字／差距／差異原因／驗證點句與年度差異原因；讀取目前目錄的 `xl17_*.json`） |
| scripts/recalc.py | LibreOffice headless 重算：`python3 scripts/recalc.py 檔案.xlsx [逾時秒數]`＝開啟、全部重算、原地存檔，輸出 JSON 並回報公式錯誤數（#REF!、#DIV/0!、#VALUE!、#NAME?、#N/A）；有錯誤或失敗時代碼 1 |
| scripts/verify.sh | 一次跑完建置與核對（見下節） |
| scripts/first_screen.py | v4.5：擷取「第一屏」文字——HTML 各頁面／分頁在 1440×900 視窗、捲動到頂端時可見的文字；Excel 各工作表前 40 個可見列（A–J 欄，重算後的值）。`python3 scripts/first_screen.py 檔案.html｜檔案.xlsx [輸出.json]` |
| scripts/attrib.py | 目標價變動拆解：輸入前一版與各步驟（.xlsx、company.json 或 repo 目錄），以 Excel 求值，輸出三情境兩條腿與加權目標價的 (a)(b)(c)(d)，檢查四項相加＝總變動 |
| scripts/test_attrib.py | 拆解工具測試：月數推算、同版對同版全為 0、滾動一季且 WACC 改為 12% 時 (a) 係數＝(1.12)^(3/12)。verify.sh 步驟 8c |
| scripts/test_rolling.py | v4.5：期間滾動測試（暫存副本）。B：日曆推算 6 種情況（12 月財年 Q1–Q4 已申報、5 月財年、無已申報季度）× 2 種目標價時點；C：只滾日曆、未更新 `asOf` 時建置必須失敗並逐項列出；A：把副本滾動到下一個已申報季度（只改日曆與標籤、不改數字），建 HTML 與 Excel，第一屏不得出現舊日曆特有的字樣（評價日、年初至今／首期標籤、已申報季度）。例外：版本紀錄與來源頁、整段落在 company.json／共識檔資料字串內的字樣。verify.sh 步驟 8b |
| scripts/check_offline.py | 離線開啟檢查（已決定事項 11）：`python3 scripts/check_offline.py 檔案.html 同版.xlsx`。HTML 單獨複製到空資料夾，Playwright 阻斷網路、以 file:// 開啟；(a) 除 HTML 本身外的請求數＝0（網路與旁邊的本機檔案都算）、(b) console 無錯誤與例外、(c) 5 個分頁 10 項關鍵數字（依 A 欄列名稱取自同版 Excel）出現在畫面且無 NaN／undefined／Infinity；v4.5 另比對第一屏（視窗內）的資料更新日、模型期間、現價日與評價日（`CHECKS_FIRST`；列名稱中的期間佔位符以任意文字比對，新舊版 Excel 皆可）。verify.sh 步驟 5c 對新建 HTML 與 dist/ 成品各跑一次 |
| scripts/check_tokenomics_tab.py | W1：「快照值＝Excel 分頁值」檢查：`python3 scripts/check_tokenomics_tab.py 檔案.xlsx`。Excel「Tokenomics_取數」分頁每個名稱 × 世代的低成本／基準／高成本＝`company.json` → `tokenomics.snapshotFile` 快照值；每個「基準」格有具名範圍 `TK_<名稱去掉 IF_／L1_>_<世代代碼>`（H100、GB200、GB300、VR200、RU；單值名稱不加世代）；W2 起其他工作表可引用 TK_ 名稱（只核對引用的名稱都存在），本分頁值仍須＝快照值。verify.sh 步驟 5d；步驟 0c 另以 `../tools/tokenomics/import_tokenomics.py --check` 確認快照可由 Tokenomics 重現（找不到 clone 時警告略過；`TOKENOMICS_DIR` 可指定） |
| scripts/permw_sens.py | W2：每 MW 敏感度的建置時快照。以 LibreOffice（UNO）開啟建好的 Excel，依序切換情境選擇與敏感度輸入（Tokenomics 低／高成本、GPU 小時價格低／高、世代組合 Rubin Ultra 版、管銷率 GAAP），讀加權目標價與融資缺口，寫入 `permw_sens.json`（有變動時代碼 3，verify.sh 步驟 3c 重建 Excel）；Excel「每MW經濟性」頁「敏感度」區讀此檔，並以「快照狀態」格比對目前輸入（不一致＝快照已過期） |
| scripts/test_permw.py | W2：暫存副本測試（verify.sh 步驟 8d）——A：`meta.mwBasis`＝facility 時 Tokenomics 每 MW 值 ÷ IF_FacilityGW；B：`revenue`＝gpuHr（虛構價格，只在副本）時每 MW 年收入＝Σ 占比 × GPU 數 × 價格 × 8,760，Python 獨立計算一致；兩者 cmp31 基準全部一致 |
| scripts/verify_legacy.sh | 舊方法回歸驗收：副本改 `methodology.perMw` 後跑 `verify.sh --vs-dist`，畫面文字、Excel 值與公式須與前一版成品 0 差異（新增列、新增工作表不計）；結果在 `out/verify_legacy.log`。`LEGACY_BASE`＝v4.6（預設，W4 起：只把收入改回 legacy、Tokenomics 快照取 git 歷史 v5.26，對 v4.6 成品，`LEGACY_DIST_REF` 預設 f373885）｜v4.5（全部舊方法，對 v4.5 成品，預設 9af51ca） |
| scripts/attrib_w4.py | W4：v4.6 → v4.7 目標價變動拆解（(d) ① 錨取代舊輸入〔k＝1〕→ ② 套用 k → ③ 上限檢查；另列融資缺口；檢查相加＝總變動） |
| scripts/q2_check_w4.py | W4：Q2 驗證拆解（模型首期每 MW 收入對 Q2 年化：爬坡分母、利用率、定價倍數、世代組合、其他；相加＝總差距；只作驗證、不校準 k） |
| scripts/attrib_w5.py | W5：v4.6 → v4.7（W4＋W5）變動拆解：W4 的 ①②③ 後加 ④ 公司調整（④-1 既有合約 k；④-2–④-5 驗證後未調整＝0）；檢查相加＝總變動 |
| scripts/compare_w5.py | W5：v4.6／v4.7 W4／v4.7 W5 三版對照 Excel（摘要、每MW_三版、公司實況驗證、敏感度、變動拆解） |
| scripts/compare_w4.py | W4：v4.6 → v4.7 收入錨定對照 Excel（摘要、每MW_前後、錨與k、證據表、Q2驗證、敏感度、變動拆解） |
| scripts/attrib_permw.py | W3：v4.5 → v4.6 目標價變動拆解（(d) 方法變更逐項依序／單獨切換；數值取自 Excel；三情境；檢查相加＝總變動） |
| scripts/make_expect.py | W3：升版預期差異清單產生器（規則檔 `scripts/expect/*_rules.json`；未歸類的差異即失敗） |
| scripts/compare_gather.py、scripts/build_compare.py | W3：前後對照取數（三檔 × 三情境，xlx.py）與對照 Excel 產生 |
| scripts/check_quarterly.js | v4.4：季度加總＝年度（三情境）、`quarterly.consistency` 一致性、超過門檻的差距都有原因；`build_html_portable.py` 建置時呼叫，verify.sh 步驟 0 |
| scripts/test_quarterly.py | v4.4：暫存副本測試——(A) 假設 Q3 實際數，Python 獨立計算差距並與 HTML、Excel 比對；(B) 可移植性（無 MW、無季度指引與共識）。verify.sh 步驟 8 |
| scripts/cloud_setup.sh | 雲端環境 setup script（只裝 Python 套件；LibreOffice Calc 於工作階段內補裝，見 CLAUDE.md） |

## Excel 產生流程（scripts/verify.sh 步驟 2–7 即為此流程）
1. `python3 build_xlsx.py`（輸出到 `out/`）
2. `python3 scripts/recalc.py out/<xlsx> 120`（必須回報 `"total_errors": 0`）
3. `python3 fix_outline.py out/<xlsx>`，再 `python3 fix_datatable.py out/<xlsx>`（v4.1：還原模擬運算表；未執行時 verify_ooxml 會失敗）
4. Excel 結構核對：`python3 verify_ooxml.py out/<xlsx>` 必須輸出 `OOXML OK`（LibreOffice／openpyxl 能開不代表 Excel 能開）
5. 數值核對：`for s in 1 2 3; do python3 xlx.py out/<xlsx> $s out/xl17_$s.json; done`，再於 `out/` 內 `node ../cmp31.js low|base|high`（輸出 OK 行數）

## 執行環境需求
- Python 3.11 以上＋`openpyxl`、`playwright==1.56.0`（`crawl.py`；對應雲端環境預裝的 Chromium build 1194，見 `scripts/cloud_setup.sh`）、`pypdf`；Node 18 以上（`structuredClone`）。
- LibreOffice **含 Calc 模組**（Debian／Ubuntu：`libreoffice-calc-nogui`）。只有 `libreoffice-core` 時 `soffice --version` 正常，但開啟 xlsx 會停住；`scripts/recalc.py` 逾時會提示此原因。
- 未設 locale 的環境（`LANG` 為空）下 LibreOffice 會把中文檔名轉成 `??`：`scripts/recalc.py` 已固定 `LC_ALL=C.UTF-8` 並以 file URL 傳遞路徑。
- `crawl.py`：Playwright 預設的 Chromium 無法啟動時（例如 playwright 未固定版本、與預裝 Chromium 不符），改用環境變數 `CHROMIUM_PATH`，未設時用 `/opt/pw-browsers/chromium`。

## 注意
- 新增到 tail.js 或 seg 檔的全域名稱，必須避開模板函式庫已用的名稱（例如 nf、XS、Hd、Tp 都會衝突），慣例是加 Q 字尾。
- LibreOffice 不支援 NORM.S.DIST，Excel 公式請用 NORMSDIST。

- 列印：總結頁的列印 CSS 依賴 segD 的 `sumQ-np`（頁首、頁籤）與 `sumQ-main`（主容器）兩個 class，修改 segD 時請保留。
- 目視檢查可用 Playwright：開啟 HTML、`page.pdf(prefer_css_page_size=True)` 應得 10 頁（v4.3：一頁摘要＋附錄 9 頁）。

## v3.5 核對方式（EV/EBITDA 錨定年度）
- 預設三情境：`python3 xlx.py <xlsx> 1|2|3 xl17_<n>.json`，再 `node cmp31.js low|base|high`（目前各 313 項）。
- 其他錨定年度：`python3 xlx.py <xlsx> 2 out/xl17_2_a3.json 3` 後於 `out/` 內 `node ../cmp31.js base 3`（第 4 個參數＝錨定年度 1–4）。
- 矩陣：HTML `evGrid()`（segB）與 Excel「評價_DCF與目標價」的「錨定×倍數｜…」列逐格比對（40 格）。

## v4.1 目標價區間
- **點位**：目前輸入的加權目標價（評等仍依點位，規則不變；規則抽成 segB 的 `rateCall`，`blendCall` 共用）。
- **情境區間（A）**：保守與積極情境的加權目標價，取 min／max（保守情境擴張最少、目標價最高）。HTML 以 `scnQ`（v4.1 自 segE 移到 segB）換情境、保留其餘手動調整；Excel 以「輸入與假設」底部 H 區的模擬運算表計算（輸入格＝情境選擇；Excel 規定輸入格與運算表須在同一工作表）。
- **方法區間（B）**：目前輸入，EV/EBITDA 倍數換成 `company.json` → `methodology.rangeMultiples`（預設 5、6）；Excel 輸入在 F 區「方法區間倍數：下端／上端」。
- **判斷句**：情境區間上緣與現價比較，並依三情境實際評等說明賣出是否在三個情境下都成立；任一情境與賣出門檻（現價 × 85%）的距離小於現價 5% 時揭露該距離。
- **DCF 權重說明**：只在 DCF 被 0 截斷時顯示；DCF 失效時改為一句；DCF 為正或選擇權模式時不顯示。
- 字串在 `targetRange()` 產生，HTML 各頁共用；Excel 以文字公式（`TEXT(x,"0.0")` 等）產生同一字串，cmp31 逐字比對。
- 模擬運算表的建置鏈：build_xlsx 以 openpyxl 寫入 `t="dataTable"` → LibreOffice 能正確計算，但存檔時改寫成 `TABLE(公式列,輸入格,輸入值)` → `fix_datatable.py` 還原為單一 `dataTable` 公式（只放在區塊左上格）。本環境沒有 Microsoft Excel，結構依 ECMA-376；Excel 實機開檔由 Andy 於合併前確認（v4.2：2026-09-25 確認正常）。

## v4.2 設定集中（同業 Comps、評等門檻）
- **同業 Comps**：數字、名稱、備註與 Comps 分頁的說明文字都在 `company.json` → `peers`；segB 的 `lM` 由 `peers.list` 產生（EV＝市值＋淨負債、EBITDA＝營業損益＋D&A 現算），build_xlsx「可比公司」頁讀同一份。名稱與備註暫時保留 HTML、Excel 兩種寫法（`labelHtml／labelXlsx`、`noteHtml／noteXlsx`），統一另開任務。Excel CRWV 列改讀 `callFacts`／`latestQuarter`。
- **評等門檻**：`company.json` → `methodology.rating`（7 項，見欄位說明）。HTML：`rateCall`、`blendCall`、`targetRange` 與含門檻的畫面文字（`pctQ`／`multTxt` 組字）。Excel：「輸入與假設」F 區尾端「▸ 評等門檻」7 格；評等代碼、賣出門檻價、判斷句的餘裕揭露、「檢查」頁的終值與股權條件都引用這些格；評等規則說明句與檢查頁的標準欄改為文字公式，改門檻時跟著變。cmp31 以標籤為鍵的「賣出門檻價（現價 × (1 − 15%)）」列，標籤由門檻值產生。
- **不移入的技術容忍值**：現價下限 0.01（避免除以 0）、股權需求 ≤0.01bn 視為不需新股。

### v4.2 核對方式（新增輸入列造成的位移）
- `scripts/verify.sh --vs-dist` 的 Excel 比對改用 `xl_diff.py --by-label`：兩檔列數或 A 欄不同的工作表，以「區段＋A 欄列名稱」配對（區段＝上方最近的第一層區段標題列，淺藍底），其餘工作表仍逐格比對。
- 配對鍵重複、或 A 欄空白的列有內容時直接報錯並列出，不自行配對。新版多出的列另列「新增列」清單，不計為差異。
- 公式比對時，舊版公式的列號（含跨工作表引用）依列對照換算後再比。另有兩類變更列清單、不計為差異：「常數改為引用輸入格」（把引用換回該格常數後與舊版公式相同）、「文字改為活公式」（重算後的值與舊版文字相同）。
- 以不同門檻測試過連動：把 `methodology.rating` 改成其他值後重建，cmp31 三情境＋FY27 錨定仍全部一致（HTML 與 Excel 同步改變）。

## v4.3 一頁摘要與市場共識
- **共識資料**：只讀 `data/consensus_crwv_20260925.json`（路徑在 `company.json` → `meta.consensusFile`）；不增補、不修改數字，資料檔沒有的欄位顯示「未列」。建置時檢查兩件事，不符即停止：共識檔的 Q3 營收指引與 `callFacts.nextQRevLo／nextQRevHi`（下一季營收指引） 相同（以 company.json 為準）；`actual1H.adjEbitda`＝Q1＋Q2。
- **逐筆顯示**：`consensusItems()`（segB）與 build_xlsx 的 `cons_items()` 產生同一組列（標籤以「共識｜」開頭）：HTML「損益與評價 → 市場共識」分頁、Excel「輸入與假設」I 區（C–E 數值、F 擷取日期、G 標記、I 來源、J 備註）。另揭露 `sourceIndependence`（兩網站的數字都來自 S&P，不是獨立交叉驗證）與「未取得」清單。
- **與市場的差異**（`consensusView()`）：模型用目前情境（預設＝基準）；FY26 為全年口徑：營收＝1H 實際＋2H 模型、調整後 EBITDA＝1H 實際調整後 EBITDA 2.667（Q1 1.157＋Q2 1.510，報導轉述 [Verified]，只用於此比較）＋2H 模型、CapEx＝1H 實際毛額 16.139＋2H 模型毛額。差距＝模型 ÷ 共識 − 1；EBITDA 率以百分點列；淨負債只列不判斷（共識口徑未揭露）。
- **判斷句**：營收、EBITDA、CapEx 任一項差距絕對值超過 `methodology.consensusGapTol`（5%）即為分歧；句子寫出分歧起始年與之後各年超過門檻的項目。
- **隱含倍數**：（價格 × 模型股數 0.587bn＋共識 FY28 淨負債）÷ 共識 FY28 調整後 EBITDA，價格分別用共識平均目標價與現價；與方法區間上緣（6x）並列。
- **一頁摘要**：HTML 總結頁第 1 頁（`onePageQ`；原 9 頁順延為附錄，列印 10 頁）；Excel「摘要」分頁（「導覽」之後，全部活公式）。四個句子（結論句、共識判斷句、隱含倍數句、Q3 句）由 HTML 函式與 Excel 文字公式各自產生，cmp31 逐字比對。反向 DCF 在 Excel 引用快照（`rv_snap.json`）。
- **現價**：$90.13（2026-09-24 收盤）；股權發行參考價同步。Comps 表的 CRWV 市值維持 9/21（與同業同一天）；股權／可轉債金額維持 6.2bn（Andy 決定）。
- **核對**：B＋C（現價不變）先與 v4.2 比對：Excel 用 `xl_diff.py --by-label --ignore=來源,摘要`（值與公式 0 差異，新增列 64）；畫面文字差異限於新頁、Street 列刪除、來源句、附錄前綴與頁碼。再做 A（改現價）並列出數字對照（見 `docs/reports/20260925_一頁摘要與市場共識.md`）。

## v4.4 季度層與差異原因
- **季度層**（`company.json` → `quarterly`；HTML `quarterlyView`〔segB〕、畫面 `QuarterTabQ`〔segE〕；Excel「季度追蹤」分頁＋「輸入與假設」J 區）：2026 Q3–2027 Q4 共 6 季，由年度模型拆分，只用於追蹤，不回寫年度數字、不影響目標價。
  - 營收：2H26 從 Q2 實際營收 2.575 逐季線性爬升（`revenueSplit` anchor），FY27 依平均在役 MW（Billable，年底之間線性內插）。
  - EBITDA 率：6 季一條直線，同時符合 2H26 與 FY27 的年度 EBITDA。調整後營業利益＝調整後 EBITDA − 車隊折舊（依期內平均 PP&E，與年度公式相同；未扣 SBC）。
  - CapEx（毛額）：有季度指引的季取指引中點（`capexSplit`＝guidanceAnchor；Q3 12.5），其餘季分配期間餘數（依季內新增 Accepted MW）。MW：季末 Accepted（6/30 起點＝公司公布 1.5 GW）。
  - 沒有 MW 資料（`driver.type`＝none）、沒有指引或共識時，對應欄位顯示「不適用」（`scripts/test_quarterly.py` 測試 B）。
- **第一屏**：焦點季（`quarterly.focus`；Excel J 區「焦點季」）各指標的模型／共識／指引／實際與差距，加一行差異原因；6 季路徑、期間合計核對、拆分方法放次層（預設收合）。
- **差距（第 4 輪，Andy 決定）**：比例差只用於營收、CapEx；利潤類（調整後 EBITDA、調整後營業利益）列金額差＋利潤率百分點；EBITDA 率為百分點；MW 與淨負債為差額。指引以中點並標位置（高於上緣／區間上半部／下半部／低於下緣）。年度共識對照同一規則（一頁摘要、市場共識分頁、Excel 摘要；判斷門檻仍以比例計）。文字型指引（`guidanceText`，例如 Q4 調整後營業利益率 low teens）只列不計差距。
- **差異原因（已決定事項 2）**：需附原因的條件——營收、CapEx、淨負債、MW：差距超過比較基準的 `methodology.consensusGapTol`（5%）；利潤類（調整後 EBITDA、調整後營業利益，年度 EBITDA 亦同）：利潤率差超過 `methodology.profitMarginTolPt`（2pt；第 5 輪）；另外模型落在公司指引區間外也須附原因。分歧判斷句仍用 5% 比例門檻，不受影響。
- **拆法／觀點的判定（第 5 輪）**：有期間合計（同期各季共識合計、或全年指引中點 − 1H 實際）時拆解季度差距：水準＝(模型期間 − 比較期間) × 比較的季占比；分配＝(模型季占比 − 比較季占比) × 模型期間，兩項相加＝季度差距。分配較大→「拆法」、水準較大→「觀點」（有 varianceReasons 設定時用其類型並附其文字）；原因文字寫出兩項金額（例：Q3 營收 vs 指引 水準 +$0.115bn、分配 +$0.023bn → 觀點）。沒有期間合計時讀 varianceReasons。原因文字的 `{路徑:格式}` 可用加減（例如 `{c.ebitda-g.adjOpInc:2}`），格式 `s2`＝帶正負號。「拆法」由程式判定（單季超過門檻、所屬期間合計在門檻內）；「觀點」「已知限制」寫在 `varianceReasons`（原因文字的 `{路徑:格式}` 由模型數字帶入，Excel 以文字公式產生同一字串）；找不到即「未歸類」，`check_quarterly.js` 讓建置失敗。年度共識對照（一頁摘要、市場共識分頁、Excel 摘要）同一機制。
- **一頁摘要**：驗證點改為焦點季 3 個指標（`quarterly.keyMetrics`：營收、調整後營業利益、CapEx）；「與市場的差異」加一句差異原因摘要。依精簡原則，方法註腳與最新分析師動作移出一頁摘要（仍在「市場共識」分頁與 Excel I 區）。
- **核對**：`check_quarterly.js`（季度加總＝年度，誤差 ≤1e-9）；cmp31 每次 313 項；`test_quarterly.py` 以暫存副本測試假設實際數與可移植性，repo 內實際數維持空白。

## v4.5 期間滾動＋反向 DCF（5a-1）
- **日曆**：`company.json` → `calendar`（財年結束月份、最新已申報／已公布季度、`modelYears`＝4、`targetHorizon`）；`calendar_q.py` 推算期間標籤、期間長度、評價日、各期期末日、折現年數（以月計）、目標價時點與 EV/EBITDA 折回調整（`evOffset`）。Excel「輸入與假設」F 區「期間與日期」列出推算結果；HTML 讀 `COMPANY_DATA.cal`（`CALQ`）。首期長度支援 0.25／0.5／0.75／1 年；5 月財年（Oracle 型）與沒有已申報季度（OpenAI 型）皆可推算（`scripts/test_rolling.py` B）。
- **折現**：DCF 的 UFCF 折到評價日至各期期末（首期 0.5、FY27 1.5…）、終值與選擇權期限＝評價日至模型期末、新股 PV＝評價日至各期期初；EV/EBITDA 腿自錨定期期末以 WACC 折回目標價時點（錨定期 k 的年數＝k − `evOffset`）。
- **評價日**（Andy 決定）：最新**已申報**（10-Q／10-K）季度的季末（目前 2026-06-30）。DCF 現值、淨負債與股數的基準日。
- **目標價時點**（2026-09-26 Andy 決定）：`valuationPlus12m`＝評價日＋12 個月（目前 2027-06-30）；`firstFullYearEnd`＝首期之後第一個完整財年末，即 v4.4 的「FY27 末」。EV/EBITDA 腿自錨定期期末以 WACC 折回此時點；DCF 腿（0 截斷與選擇權皆同）以 WACC 自評價日推到此時點（`perShareT`；Excel「DCF 每股（推到目標價時點）」），兩條腿同一時點。「DCF 每股」仍是評價日現值，反向 DCF 以它對現價求解。依設計，任何日曆下錨定期的折回年數都不為負，不需內插。
- **與賣方 12 個月目標價的差異**：賣方以股價日起算（2026-09-24 → 2027-09-24），本模型以評價日起算，目前早約 2.8 個月。若以股價日起算，兩條腿約 × 1.025：基準 $45.5 → 約 $46.6（[Derived]）。差距在 10-Q 申報後最小（約 1.5–3 個月）、下一次申報前最大（約 4–5 個月）。
- **升版驗收**：對前一版成品做 `EXPECT=scripts/expect/v4_5_vs_v4_4.txt scripts/verify.sh --vs-dist`。清單列出預期變動的格（每格附原因），以及「區段 舊 => 新」「列 舊 => 新」改名對照；Excel 其餘須 0 差異。畫面文字差異另外分類說明。
- **期間字樣**：說明文字以佔位符書寫，建置時換成目前日曆的字樣；期間特有的敘述與數字放 `texts`、`ytdActual.notes`。滾動後第一屏不得出現舊日曆字樣（`test_rolling.py` A；離線檢查另比對第一屏的資料更新日、模型期間、現價日與評價日）。
- **反向 DCF**：`scripts/rv_solve.py` 以 Excel（LibreOffice UNO）求解，取代 JS 快照（verify.sh 步驟 3b）。
- **目標價變動拆解**（已決定事項 12）：每次升版把目標價變動拆成 (a) 時間推移、(b) 實際數更新、(c) 假設變更、(d) 方法變更，依序重算，四項相加＝總變動，並列於版本紀錄。v4.5 全部屬 (d)。
  - (a) 時間推移是**定義值**（純時間價值，以 WACC 累積；2026-09-26 Andy 決定）：前一版兩條腿（0 截斷前）×（1＋WACC）^（評價日後移月數 ÷ 12），再做 0 截斷與加權。WACC 讀前一版 Excel 的 WACC 輸入格，月數由 `calendar_q.py` 推算的評價日計算（`calendar_q.months_between`），不寫死。等同「離開模型期的季度以模型值當作暫定實際數、淨負債與股數推到新評價日、其餘路徑不變」。不可用「只滾動日曆、不改其他資料」重算：那樣會刪掉該季流入、保留流出，淨負債與股數停在舊評價日，結果偏低。
  - (b) 實際數更新含兩個來源：該季實際與模型的差異；以及融資時點殘差（(a) 以 WACC 累積淨負債，實際以債務成本累積；該季實際發股）。
  - 只拆解目標價採用的 0 截斷口徑；選擇權模式的 DCF 腿只列前後值（Andy 決定）。
  - 工具：`scripts/attrib.py`（數值取自 Excel；三情境以「情境選擇」切換）。測試：`scripts/test_attrib.py`（verify.sh 步驟 8c）。
- **每季滾動流程**（已決定事項 10）：(1) 財報當天（新聞稿）：填 `quarterly.actuals`、`calendar.latestQuarterReported` 加一季、焦點季改為下一季；年度首期不動。(2) 10-Q／10-K 申報後：`calendar.latestQuarterFiled` 加一季；更新 `ytdActual`（含 `throughQuarter`、`months`、`label`、`notes`、`jvSplit`）、`historicalPL` 最後一格、`latestQuarter`、RPO 桶、債務與租賃到期表、`texts`（頁首來源一行、現金稅、首期股權組成）與各期清單（首期縮短一季）；更新首期一次性金額與期初餘額，並把 `asOf` 對應季度改為新的 `latestQuarterFiled`（滾動檢查，清單見下方「asOf」欄位說明）；建置時 `calendar_q.py` 檢查截止季與月數、`historicalPL` 與 `ytdActual.label` 的標籤是否一致，以及 `asOf` 各項的所屬季度。(3) 財年最後一季（10-K）後：首期為下一財年全年模型，模型期延長一年，需補新一年的假設（Andy 決定）。

## W2 每 MW 經濟性（資本支出、營運成本、收入改接 Tokenomics）
- **方法開關**（`company.json` → `methodology.perMw`，建置時決定公式）：`capex`＝tokenomics（每 MW 建置成本＝Σ 新增世代占比 × `IF_CapexIT`，不含廠房）｜legacy；`cost`＝bottomUp（電費 `IF_PowerCost`、IT 維護 `IF_MaintIT`、人員軟體 `IF_StaffSW`、稅險 `IF_TaxIns` × IT 資本占比，依平均在役世代加權 × 平均在役 MW；公司管銷＝營收 × 管銷率）｜ebitdaPct；`revenue`＝gpuHr（`pricing.gpuHr` × 每 MW GPU 數 × 8,760）｜legacy。全為舊方法時畫面與 Excel 與 v4.5 完全相同（`scripts/verify_legacy.sh`）。目前預設：tokenomics／bottomUp／legacy（收入採備案：W1 查不到 GB200／GB300／VR200 兩個獨立長約價來源）。
- **世代組合**（`fleet`）：期初在役（最新季末）MW 與世代占比、各期新增 MW 的世代占比（`newMix`；`newMixAlt`＝Rubin Ultra 版敏感度）。汰換批次（各年新增 MW 在第 `gpuLife` 年）由最舊世代先出、以當期新增世代補回。Excel「輸入與假設」B 區後「世代組合」子區（含 Tokenomics 引用值：列＝世代、欄＝項目，經 `TK_` 具名範圍，`OFFSET` 依「Tokenomics 成本情境」取低／基準／高）；期初／期末／平均在役結構在「每MW經濟性」頁。
- **EBITDA 率（bottomUp）**：由下而上 EBITDAR 率＝1 −（現金營運成本 ÷ 營收）；EBITDA 率＝EBITDAR 率 − 租金 ÷ 營收（租金只扣一次）。C 區「穩態 EBITDA 率」預設＝由下而上 FY30（公式），改成數值即為 FY30 目標、差額線性分攤到各期（反向 DCF、敏感度 59%／70% 沿用此格）。電力／維護 overlay 自動停用。
- **Tokenomics 名稱缺漏**（v5.26 前；v4.6 起快照為 v5.26，已無缺漏，機制保留）：`IF_MaintIT`→C 區「維護成本」、`IF_StaffSW`／`IF_TaxIns`→0、`IF_DeprLifeIT`→`defaults.gpuLife`；「檢查_連動」與 HTML 連動檢查顯示「Tokenomics 名稱缺漏 N 項，成本為暫代值」，重抓含這些名稱的快照後自動改用正式值並消失（不需改程式）。
- **MW 口徑**（`meta.mwBasis`）：IT（預設）｜facility（Tokenomics 每 MW 值 ÷ `IF_FacilityGW`，基準；`scripts/test_permw.py` 測試 A）。
- **對照列**（不入損益）：每 MW 經濟持有成本（`IF_HoldEcon`，在役世代加權；含廠房資本回收）、隱含每 GPU 小時價格（＝每 MW 年收入 ÷（每 MW GPU 數 × 8,760）；revMW 為 100% 計費時數的值，利用率在收入端另乘）與 `IF_GPUhrEcon` 倍數、同業每 MW 收入（`pricing.peerRevPerMw`）、各世代市場價格（`pricing.marketRefs`）、v4.5 舊值、期末 ARR ÷ MW、最近一季實際每 MW 現金營運成本。
- **每 MW 經濟性彙總**（Excel「每MW經濟性」頁最上方；HTML「資金模型 → 運營活動 → 每 MW 經濟性」，只在使用新方法時顯示）：每平均在役 MW、年化（US$m／MW／年）的收入、電費、IT 維護、人員軟體、稅險、管銷、租金、現金成本、EBITDA、D&A、利息、稅前，另列 IT 與設施兩種 MW 口徑。cmp31 逐列比對；敏感度為建置時快照（`scripts/permw_sens.py`）。
- **差異原因**：`varianceReasons.list` 可加 `perMw` 條件，只在方法相符時適用（排在前面者優先）。

## v4.6 成品與前後對照（W3；2026-10-08）
- **成品**：`dist/20261008_CoreWeave收支模型_v4_6.{html,xlsx}`；Tokenomics 快照 v5.26（`data/tokenomics_snapshot_v5.26.json`，commit 4074684）。每 MW 方法：tokenomics／bottomUp／legacy、MW 口徑 IT。
- **畫面**：一頁摘要「1｜結論」加一句每 MW（首個完整財年：年收入、現金成本〔含租金〕、EBITDA、與 Tokenomics 經濟持有成本的倍數；HTML `pmLineQ`＝Excel「摘要」→「結論｜每 MW 經濟性句」，cmp31 逐字比對；舊方法組合不顯示）；「每 MW 經濟性」分頁的彙總表放次層（預設收合）。
- **升版驗收**：`DATE=… EXPECT=scripts/expect/v4_6_vs_v4_5.txt scripts/verify.sh --vs-dist`（對 v4.5 成品）。預期差異清單由 `scripts/make_expect.py` 依 `scripts/expect/v4_6_rules.json` 產生：下游工作表整張允許變動；「輸入與假設」只允許每 MW 方法直接改動的列；同業頁只允許 CRWV 列；其他工作表與列須 0 差異（未歸類即失敗）。有 `EXPECT` 時畫面文字（crawl）只要求無頁面錯誤、無缺頁，差異明細存 `out/crawl_upgrade_diff.txt`。舊方法組合對 v4.5 成品仍須 0 差異（`scripts/verify_legacy.sh`；v4.6 起副本的 dist/ 換成 git 歷史中的 v4.5 成品）。
- **目標價變動拆解**：`scripts/attrib_permw.py --v45 v4.5.xlsx`：(a)(b)(c)＝0、(d) 依 ① 每 MW 資本支出 → ② 折舊年限 → ③ 營運成本由下而上 → ④ 收入 → ⑤ MW 口徑逐項切換，另列單獨切換；檢查各步相加＝總變動（< 0.01 美元／股）。
- **前後對照報告**：`scripts/compare_gather.py`（三檔 × 三情境取數；檢查舊方法副本與 v4.5 成品 5,439 個共有數值格相同）→ `scripts/build_compare.py` → `docs/reports/20261007_coreweave_v4.6_前後對照.xlsx`（摘要含 Q2 2026 實際對帳、每MW_前後、參數對照、變動拆解、Tokenomics參考線、已知限制）與同名 `.md`。
- **verify 工具修正**：`xlx.py` 切換情境前把模擬運算表輸出換成來源檔的快取值（LibreOffice 批次重算時運算表會使非基準情境的部分格殘留中間值；改錨定年度時保留運算表）；`cmp31.js` DCF 失效時 HTML 的 NaN 與 Excel 的 0 視為一致。
- **Tokenomics 升版時**：見交接檔「Tokenomics 連結」（重抓快照 → verify → 升版）。

## v4.7 每 MW 收入以 Tokenomics 為錨（W4 r2；2026-10-08；已決定事項 14）
- **方法**：`methodology.perMw.revenue`＝tkAnchor（預設）。每 MW 年收入（100% 計費時數）＝Σ 平均在役占比 × `IF_HoldEcon`（`TK_HoldEcon_*`，目前 Tokenomics 成本情境）× 定價倍數 k；**k＝隨需占比 × k_現貨＋（1 − 隨需占比）× k_長約**（r2：CRWV 先簽多年期合約再建產能，RPO 未覆蓋的新增產能也按長約價；隨需占比基準 0%、敏感度 10%／20%）。B 區「每 MW 年收入」改引用「每MW經濟性」頁「Tokenomics 錨 × k」列 ÷ 1000；利用率與爬坡分母（Billable ÷ Accepted）照舊在收入端另乘。
- **k 隨成本情境重算**（r2）：輸入的 k_長約、k_現貨 是基準成本情境的值；其他成本情境以基準證據世代（`long.refEvidence`／`spot.refEvidence`）的 Tokenomics 成本比例重算（k＝市場價格 ÷ 同情境持有成本），所以 Tokenomics 低／高成本敏感度下收入大致不變、成本改變。
- **公司因素**（`pricing.anchorMultiple`）：k_長約 0.76（0.70–1.00；三筆長約中位數 0.89 列為敏感度）、k_現貨 1.76（1.5–2.3），以「市場價格 ÷ Tokenomics 同世代持有成本（IF_GPUhrEcon 或 IF_HoldEcon）」為證據（`evidence`；倍數在 Excel 以 TK_ 名稱計算，「每MW經濟性」k 證據表）；`onDemandShare`＝隨需占比。不以公司營收、ARR、RPO 金額或 Q2 隱含 k 反推 k 或隨需占比；Q2 只作驗證（`scripts/q2_check_w4.py`）。RPO 覆蓋率（排程 RPO ÷ 長約價產能收入）只作對照列。
- **輸入頁**：B 區後「世代組合」子區新增「定價倍數 k」三格（k_長約、k_現貨、隨需占比）與「TK 收入上限」區塊（`IF_RevGWFleet`）。
- **上限檢查**：CRWV 每 MW 計費收入 ÷ `IF_RevGWFleet`（Tokenomics 客戶端每 GW 付費 token 營收，在役世代加權）＝neocloud 拿走客戶營收的比例；> `methodology.checks.revCapShareMax`（50%）時「檢查_連動」與 HTML 連動檢查警示。快照沒有 `IF_RevGWFleet` 時顯示「不適用」。
- **對照列（不驅動）**：v4.6 舊輸入 `m.revMW`、RPO 覆蓋率、GPU 小時路線（`pricing.gpuHr` 空白＝不適用）、期末 ARR ÷ MW、Q2 每 MW 年收入 ÷ 平均在役 MW 與 ÷ 計費 MW、Q2 隱含 k。
- **敏感度**（`permw_sens.json` 建置時快照＋HTML 即時）：既有四組外新增 k_長約 0.70／0.89（三筆中位數）／1.00、隨需占比 10%／20%。
- **反向 DCF**：`rv_solve.py` 在區間內無解時寫 null，Excel 顯示「無解」（v4.7 穩態 EBITDA 率在 99% 內無解）。
- **Tokenomics 快照**：v5.27（`data/tokenomics_snapshot_v5.27.json`，commit 862bdd4；引用的 25 個名稱數值與位置與 v5.26 相同，新增 `IF_RevGWFleet`）。
- **升版驗收**：`DATE=2026-10-08 EXPECT=scripts/expect/v4_7_vs_v4_6.txt scripts/verify.sh --vs-dist`（對 v4.6 成品；清單由 `scripts/make_expect.py` 依 `scripts/expect/v4_7_rules.json` 產生，規則檔新增 `renames`〔指定列改名〕；營運成本列〔電費、IT 維護、人員軟體、稅險、租金〕不在清單＝須 0 差異）。舊方法組合（revenue=legacy）對 v4.6 成品 0 差異：`scripts/verify_legacy.sh`。
- **JS 修正**：segB 無槓桿 NOL 的虧損改為全額加回（原只加回 80%，與 Excel 不一致；v4.6 前未觸發）。
- **報告**：`docs/reports/20261008_coreweave_v4.7_收入錨定.{xlsx,md}`。

## v4.7 公司實況驗證（W5；2026-10-08；已決定事項 15；覆寫 v4.7）
- **規則**：每個取自 Tokenomics 的參數對 CRWV 已申報實際數逐項對照（`company.json` → `companyAdjust`）：差距 ≤ `gapTol`（10%）用 Tokenomics 值；> 10% 且有證據的機制 → 公司調整；> 10% 找不到機制 → Tokenomics 為基準、公司實際為敏感度。不以營收或 EBITDA 總數倒推係數。
- **公司調整（唯一一項）**：既有合約 k——`k_t＝既有占比_t × k_既有＋（1 − 既有占比_t）× k_新約`；既有合約 MW＝最新季末在役（`fleet.openMix.activeMW`）逐期扣汰換（最舊世代先出）；`k_既有＝（Q2 每在役 MW 年收入 − 服務）÷（計費比例 × 利用率 × Q2 錨）`，Q2 錨隨目前成本情境；`k_新約＝隨需占比 × k_現貨＋（1 − 隨需占比）× k_長約 ×（1＋新約價格調整）`。輸入頁「定價倍數 k」區新增「既有合約 k 開關」「新約價格調整」；C 區新增「由下而上營運成本倍數」（基準 1）。
- **新頁「公司實況驗證」**：參數驗證（Tokenomics 值｜CRWV 實際｜差距｜採用值｜規則＋機制、證據、調整說明）、敏感度輸入、Q2 逐項對帳、證據與找不到清單；HTML「每 MW 經濟性 → 公司實況驗證」同列（cmp31 比對）。Q2 營運成本改為含變動租賃（轉付房東的水電），固定租金對固定租金。
- **敏感度**新增 6 組：既有合約 k 不套用（＝W4）、新約 +25%、隨需 10% × CRWV 短天期 k、營運成本＝Q2 實際比率、每 MW 建置成本＝年初至今實際比率、兩者皆實際（`scripts/permw_sens.py` 讀「公司實況驗證」頁敏感度輸入列）。
- **工具**：`scripts/attrib_w5.py`（W4 的 ①②③ 後加 ④ 公司調整逐項）、`scripts/compare_w5.py`（三版對照 Excel）；`verify_legacy.sh` 的 v4.6 副本移除 `companyAdjust`；`calendar_q.ROLL_FIELDS` 新增 `companyAdjust.capexActual`（選用區段，沒有時不檢查）。
- **報告**：`docs/reports/20261008_coreweave_v4.7_公司實況驗證.{xlsx,md}`。

## company.json 欄位說明（換公司填表指引）
換成 Nebius、Oracle、OpenAI 等公司時，照這一節逐欄填寫 `company.json`；HTML 與 Excel 都從這個檔讀資料，改完執行 `scripts/verify.sh`。

**填表慣例**
- 金額單位是**十億美元（US$bn）**，例如 4.653 代表 46.53 億美元；另有標示的例外：每股（US$）、每 MW 建置成本（百萬美元／MW，US$m/MW）、股數（十億股，bn）。
- 「比例」寫成小數（0.25＝25%）；標示「%」的欄位寫成百分點（25＝25%）。兩種寫法沿用既有程式，不可混用。
- 「清單」依模型期順序填：FY26 下半年、FY27、FY28、FY29、FY30，共 5 格（除非另有說明）。
- 文字中的來源標記沿用 [Verified]（已公開可查）、[Interested-party]（利害關係人說法）、[Derived]（由其他數字換算）、[Assumed]（判斷值）。
- 「換公司」欄：**必改**＝公司特有的資料；**檢查**＝判斷值，要依新公司重新評估；**可沿用**＝口徑或方法，通常不必改。
- 下表的「目前數值」是 CoreWeave v4.7 的值（版本號讀 `vlog.py`、期間讀 `calendar_q.py`，由本檔自動帶入）；過長的文字只顯示開頭。表格由 `scripts/fields_doc.py` 產生，新增欄位時先在該檔補說明，再重新產生。

### `meta`：基本資料

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `meta.company` | 公司名稱 | 文字 | CoreWeave | 必改 |
| `meta.ticker` | 股票代號 | 文字 | CRWV | 必改 |
| `meta.updateDate` | 資料更新日 | 日期 | 2026-09-25 | 必改 |
| `meta.priceDate` | 股價日期（現價的收盤日；畫面與 Excel 的現價日期都讀這格） | 日期 | 2026-09-24 | 必改 |
| `meta.consensusFile` | 市場共識資料檔路徑（v4.3；只讀，由使用者查證後提供；建置時併入 HTML、Excel 讀同一檔） | 路徑 | data/consensus_crwv_202609… | 必改 |
| `meta.sourceOrderNote` | 資料來源的先後與衝突時的取捨原則（畫面說明文字） | 文字 | 時序先法說（8/11）、後 10-Q（8/12）、再… | 必改 |
| `meta.mwBasis` | MW 口徑（W2）：IT＝IT 關鍵電力（與 Tokenomics 每 GW 相同）；facility＝設施電力（Tokenomics 每 MW 值 ÷ IF_FacilityGW） | 代碼 | IT | 必改 |
| `meta.mwBasisNote` | MW 口徑的依據說明（W2） | 文字 | CoreWeave 的 MW（active／cont… | 必改 |

### `calendar`：期間與日期（v4.5）

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `calendar.fiscalYearEndMonth` | 財年結束月份（v4.5；CoreWeave 12、Oracle 5） | 月 | 12 | 必改 |
| `calendar.latestQuarterFiled` | 最新已申報（10-Q／10-K）的財季，格式 FYyyQn；驅動年度首期滾動與評價日（v4.5） | 文字 | FY26Q2 | 必改 |
| `calendar.latestQuarterReported` | 最新已公布（財報新聞稿）的財季；驅動季度層，年度首期不受影響（v4.5） | 文字 | FY26Q2 | 必改 |
| `calendar.firstModelFY` | 沒有已申報季度時（例如未上市公司）的首個模型財年，例如 FY26；有已申報季度時填 null（v4.5） | 文字 | None | 檢查 |
| `calendar.modelYears` | 首期之後的完整財年數（固定期數滾動；Excel 欄位固定 5 期，所以為 4）（v4.5） | 年 | 4 | 可沿用 |
| `calendar.targetHorizon` | 目標價時點：firstFullYearEnd＝首期之後第一個完整財年末；valuationPlus12m＝評價日＋12 個月（v4.5） | 代碼 | valuationPlus12m | 可沿用 |

### `asOf`：滾動檢查（首期一次性金額與期初餘額的所屬季度）

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `asOf` | 滾動檢查：首期一次性金額與期初餘額所屬的已申報季度（鍵＝欄位路徑，清單定義在 calendar_q.py → ROLL_FIELDS）。每季 10-Q 後逐項更新數值，並把季度改為 calendar.latestQuarterFiled；缺漏或季度不符即建置失敗 | 物件（季度） | 物件（_note、defaults.capexFloorFY0、leases.onBalanceCash[0]、leases.operatingPayments[0]、leases.financePayments[0]、debt.amortization[0]、defaults.jvCommit[0]、scenarios.capexTemplate.div[0]、defaults.intCal、defaults.services[0]、defaults.atm、scenarios.leaseHighPath[0]、rpo.bucketWeights[0]、defaults.cash、debt.instruments、debt.convertible、valuation.netDebt、valuation.shares、defaults.ppeOpen、defaults.billableOpen、defaults.rpoOpen、defaults.rpoPendingAdd、defaults.eqCapShares、defaults.mwYearEnd、fleet.openMix、companyAdjust.capexActual） | 必改 |

### `ytdActual`：年初至今實際數（10-Q；v4.5 前為 actual1H）

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `ytdActual.throughQuarter` | 年初至今實際數的截止財季，須等於 calendar.latestQuarterFiled（v4.5，原 actual1H） | 文字 | FY26Q2 | 必改 |
| `ytdActual.months` | 年初至今的月數，須等於日曆推算值（v4.5） | 月 | 6 | 必改 |
| `ytdActual.label` | 「年初至今實際數」的標題 | 文字 | 1H26 實際（10-Q） | 必改 |
| `ytdActual.jvSplit` | JV 出資與策略投資的拆分（v4.5；說明文字與「JV 已付」讀此） | 物件（US$bn） | 物件（jv、strategic） | 必改 |
| `ytdActual.notes` | 各欄位的逐列說明（來源、口徑、拆分；Excel「輸入與假設」G 區；v4.5 起隨資料一起更新） | 物件（文字） | 物件（cash1231、cfo、cashCapex、capex、jv、borrow、debtRepaid、cappedCall、equity、interest、leasePaid、revenue、opInc、ni、prepay、da、sbc、eps、ngEps） | 必改 |
| `ytdActual.cash1231` | 上一年底現金 | US$bn | 3.127 | 必改 |
| `ytdActual.revenue` | 上半年營收 | US$bn | 4.653 | 必改 |
| `ytdActual.capex` | 上半年資本支出（認列口徑，含設備商融資） | US$bn | 16.139 | 必改 |
| `ytdActual.cashCapex` | 上半年現金購置固定資產 | US$bn | 14.117 | 必改 |
| `ytdActual.interest` | 上半年利息費用 | US$bn | 1.176 | 必改 |
| `ytdActual.leasePaid` | 上半年租賃現金支付 | US$bn | 0.748 | 必改 |
| `ytdActual.debtRepaid` | 上半年還款 | US$bn | 5.219 | 必改 |
| `ytdActual.borrow` | 上半年借款 | US$bn | 16.747 | 必改 |
| `ytdActual.equity` | 上半年股權募資 | US$bn | 2.982 | 必改 |
| `ytdActual.cappedCall` | 上半年 capped call（可轉債配套避險）支出 | US$bn | 0.492 | 必改 |
| `ytdActual.jv` | 上半年合資（JV）與策略投資出資 | US$bn | 0.688 | 必改 |
| `ytdActual.cfo` | 上半年營運現金流 | US$bn | 3.663 | 必改 |
| `ytdActual.prepay` | 上半年客戶預付（遞延收入淨流入） | US$bn | 1.365 | 必改 |
| `ytdActual.da` | 上半年折舊攤銷 | US$bn | 2.54 | 必改 |
| `ytdActual.sbc` | 上半年股份基礎薪酬 | US$bn | 0.318 | 必改 |
| `ytdActual.opInc` | 上半年 GAAP 營業損益 | US$bn | -0.193 | 必改 |
| `ytdActual.ni` | 上半年淨損益 | US$bn | -1.366 | 必改 |
| `ytdActual.eps` | 上半年 GAAP 每股盈餘 | US$ | -2.53 | 必改 |
| `ytdActual.ngEps` | 上半年每股盈餘（加回股份基礎薪酬） | US$ | -1.94 | 必改 |
| `ytdActual.adjEbitda` | 上半年調整後 EBITDA（v4.3；只用於與市場共識比較 FY26，不進模型損益與評價） | US$bn | 2.667 | 必改 |
| `ytdActual.adjEbitdaMeta` | 上述數字的 Q1／Q2 拆分、來源、標記與備註（建置時檢查 Q1＋Q2＝合計） | 物件 | 物件（q1、q2、tag、sources、crossCheck、note、usage） | 必改 |

### `historicalPL`：歷年損益

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `historicalPL` | 歷年損益，一年一列（損益頁的歷史欄）。每列欄位：year 年度、revenue 營收、opInc GAAP 營業損益、ni 淨損益（以上 US$bn）、eps GAAP 每股盈餘、ngEps 加回股份基礎薪酬的每股盈餘（US$）、shares 加權流通股數（bn）、tag 來源標記 | 清單 | 4 筆 | 必改 |

### `rpo`：已簽約未認列營收（RPO）

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `rpo.scheduledShare` | 剩餘履約義務（RPO，已簽約未認列的營收）預計在模型期內認列的比例 | 比例 | 0.84 | 必改 |
| `rpo.bucketWeights` | RPO 在五期（首期模型部分＋4 個完整財年；目前為 FY26 下半年、FY27、FY28、FY29、FY30）各期認列的比例，合計等於上一欄 | 比例清單 | 0.1025、0.205、0.2、0.195、0.1375 | 必改 |

### `leases`：租約

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `leases.onBalanceCash` | 已入帳租約在五期（首期模型部分＋4 個完整財年；目前為 FY26 下半年、FY27、FY28、FY29、FY30）各期的現金租金 | US$bn 清單 | 1.04、2.339、2.306、2.373、2.289 | 必改 |
| `leases.afterFY30` | 已入帳租約在模型期之後還要付的租金合計 | US$bn | 19.023 | 必改 |
| `leases.facts.onBal` | 已入帳租約未折現付款合計 | US$bn | 29.37 | 必改 |
| `leases.facts.notCommenced` | 已簽約但尚未起租的租約（表外） | US$bn | 35.5 | 必改 |
| `leases.facts.singleCap` | 單一大型站點的租金上限（10-Q 揭露） | US$bn | 14.7 | 必改 |
| `leases.facts.share` | 第三方租賃占機房取得的比例（用於租金基準檢驗） | 比例 | 0.85 | 檢查 |
| `leases.operatingPayments` | 營業租賃到期表：五期各期，最後一格為之後合計（Excel 租賃頁） | US$bn 清單 | 1.028、2.116、2.306、2.373、2.289、19.023 | 必改 |
| `leases.financePayments` | 融資租賃到期表：同上格式 | US$bn 清單 | 0.012、0.223、0、0、0、0 | 必改 |

### `debt`：既有債務

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `debt.amortization` | 既有債務在五期（首期模型部分＋4 個完整財年；目前為 FY26 下半年、FY27、FY28、FY29、FY30）各期的排程還本 | US$bn 清單 | 4.413、6.184、4.416、2.421、3.221 | 必改 |
| `debt.amortAfterFY30` | 模型期之後的還本合計 | US$bn | 14.896 | 必改 |
| `debt.instruments` | 既有債務逐筆明細，一筆一列：[名稱, 追索／非追索, 到期, 有效利率（比例）, 本金（US$bn）, 備註]。加總本金須等於 10-Q 本金合計 | 清單 | 16 筆 | 必改 |
| `debt.convertible.principal` | 期後新發行可轉債本金（不在五期還本表內，只計利息） | US$bn | 3.7 | 必改 |
| `debt.convertible.coupon` | 該可轉債票面利率 | 比例 | 0.02875 | 必改 |

### `latestQuarter`：最新一季財報數字（10-Q）

換公司：全部必改（公司特有資料）。「程式使用」為否的欄位只是存檔備查，畫面與 Excel 的說明文字目前另外寫在程式裡。

| 欄位 | 意義 | 單位 | 目前數值 | 程式使用 |
|---|---|---|---|---|
| `filed` | 申報日 | 日期 | 2026-08-12 | 否 |
| `periodEnd` | 季末日 | 日期 | 2026-06-30 | 否 |
| `revenue` | 當季營收 | US$bn | 2.575 | 是 |
| `yoy` | 當季營收年增率 | 比例 | 1.12 | 是 |
| `h1Revenue` | 上半年營收 | US$bn | 4.653 | 是 |
| `costRev` | 當季營收成本 | US$bn | 0.879 | 否 |
| `techInfra` | 當季技術與基礎設施費用 | US$bn | 1.507 | 否 |
| `opInc` | 當季 GAAP 營業損益 | US$bn | -0.049 | 否 |
| `interest` | 當季利息費用 | US$bn | 0.64 | 否 |
| `h1Interest` | 上半年利息費用 | US$bn | 1.176 | 否 |
| `ni` | 當季淨損益 | US$bn | -0.626 | 否 |
| `h1Ni` | 上半年淨損益 | US$bn | -1.366 | 否 |
| `epsDiluted` | 當季稀釋每股盈餘 | US$ | -1.14 | 否 |
| `sbc` | 當季股份基礎薪酬 | US$bn | 0.165 | 否 |
| `da` | 當季折舊攤銷 | US$bn | 1.393 | 否 |
| `adjEbitda` | 當季調整後 EBITDA | US$bn | 1.51 | 否 |
| `adjOpInc` | 當季調整後營業利益 | US$bn | 0.128 | 否 |
| `rpo` | 季末 RPO | US$bn | 103.7 | 是 |
| `backlog` | 季末 backlog（含其他） | US$bn | 104.2 | 否 |
| `rpo24m` | RPO 於 24 個月內認列比例 | 比例 | 0.41 | 否 |
| `rpo25to48` | RPO 於 25–48 個月認列比例 | 比例 | 0.39 | 否 |
| `rpo49to78` | RPO 於 49–78 個月認列比例 | 比例 | 0.2 | 否 |
| `cash` | 季末現金 | US$bn | 5.524 | 是 |
| `restricted` | 受限現金 | US$bn | 1.38 | 否 |
| `marketable` | 有價證券 | US$bn | 0.015 | 是 |
| `availability` | 未動用信用額度 | US$bn | 10.014 | 是 |
| `debtPrincipal` | 債務本金合計 | US$bn | 35.551 | 是 |
| `recourseNet` | 追索債務淨額 | US$bn | 31.405 | 否 |
| `nonRecourseNet` | 非追索債務淨額 | US$bn | 3.663 | 否 |
| `ddtlOut` | DDTL 未償餘額 | US$bn | 13.6 | 否 |
| `notesOut` | 票據與可轉債未償餘額 | US$bn | 16.6 | 否 |
| `sharesA` | A 股流通股數 | bn 股 | 0.458872 | 否 |
| `sharesB` | B 股流通股數 | bn 股 | 0.0926649 | 否 |
| `sharesOut` | 流通股數合計 | bn 股 | 0.551537 | 是 |
| `basicWaso` | 加權平均流通股數 | bn 股 | 0.551 | 否 |
| `capexQ2` | 當季資本支出 | US$bn | 9.352 | 是 |
| `capexH1` | 上半年資本支出 | US$bn | 16.139 | 是 |
| `cashCapexH1` | 上半年現金購置固定資產 | US$bn | 14.117 | 否 |
| `cfoH1` | 上半年營運現金流 | US$bn | 3.663 | 否 |
| `cashInterestH1` | 上半年現金利息 | US$bn | 0.806 | 否 |
| `capInterestH1` | 上半年資本化利息 | US$bn | 0.176 | 否 |
| `deferredTotal` | 遞延收入合計 | US$bn | 9.692 | 是 |
| `deferredIn` | 上半年遞延收入淨流入 | US$bn | 1.365 | 是 |
| `ppe` | 固定資產毛額 | US$bn | 46.736 | 否 |
| `cip` | 在建工程 | US$bn | 11.9 | 否 |
| `rouOp` | 營業租賃使用權資產 | US$bn | 16.595 | 否 |
| `opLeaseLiab` | 營業租賃負債 | US$bn | 16.319 | 是 |
| `finLeaseLiab` | 融資租賃負債 | US$bn | 0.221 | 否 |
| `onBalanceUndiscounted` | 已入帳租約未折現付款 | US$bn | 29.37 | 是 |
| `offBalanceLease` | 未起租租約（表外） | US$bn | 35.5 | 是 |
| `singleSiteCap` | 單一站點租金上限 | US$bn | 14.7 | 是 |
| `singleSiteMw` | 該站點未交付 MW | MW | 393 | 否 |
| `constructionCostMw` | 按造價計租的未交付 MW | MW | 355 | 否 |
| `equipCommitLo` | 設備採購承諾下緣 | US$bn | 0.5 | 否 |
| `equipCommitHi` | 設備採購承諾上緣 | US$bn | 1.2 | 否 |
| `jvCommit` | JV 承諾出資上限 | US$bn | 1.7 | 否 |
| `jvPaidH1` | 上半年已付 JV 出資 | US$bn | 0.55 | 否 |
| `vieExposure` | 可變利益實體（VIE）最大曝險 | US$bn | 0.108 | 否 |
| `rouObtainedH1` | 上半年新取得使用權資產 | US$bn | 8.359 | 否 |
| `leaseCashH1` | 上半年租賃現金支付 | US$bn | 0.748 | 否 |
| `custA` | 最大客戶營收占比 | 比例 | 0.36 | 否 |
| `custB` | 第二大客戶營收占比 | 比例 | 0.26 | 否 |
| `custC` | 第三大客戶營收占比 | 比例 | 0.1 | 否 |
| `sm` | 當季銷售行銷費用 | US$bn | 0.06 | 否 |
| `smSbc` | 其中 SBC（銷售行銷） | US$bn | 0.012 | 否 |
| `ga` | 當季一般管理費用 | US$bn | 0.178 | 否 |
| `gaSbc` | 其中 SBC（一般管理） | US$bn | 0.084 | 否 |
| `sbcCostTi` | 當季 SBC（營收成本＋技術與基礎設施） | US$bn | 0.069 | 否 |
| `opLeaseCost` | 當季營業租賃成本 | US$bn | 0.5 | 否 |
| `varLeaseCost` | 當季變動租賃成本 | US$bn | 0.15 | 否 |
| `finLeaseCost` | 當季融資租賃成本 | US$bn | 0.017 | 否 |
| `debtIssuedH1` | 上半年借款 | US$bn | 16.747 | 否 |
| `debtRepaidH1` | 上半年還款 | US$bn | 5.219 | 否 |
| `equityH1` | 上半年股權募資 | US$bn | 2.982 | 否 |

### `callFacts`：法說會與期後事項

換公司：全部必改（公司特有資料）。「程式使用」為否的欄位只是存檔備查，畫面與 Excel 的說明文字目前另外寫在程式裡。

| 欄位 | 意義 | 單位 | 目前數值 | 程式使用 |
|---|---|---|---|---|
| `activeGw` | 主動電力 | GW | 1.5 | 否 |
| `activeAddQ2` | 當季新增主動電力 | MW | 500 | 是 |
| `juneAddMw` | 季末月單月新增 | MW | 300 | 否 |
| `contractedGw` | 簽約電力（法說日） | GW | 4.2 | 是 |
| `contractedGwQ2` | 簽約電力（季末） | GW | 3.7 | 否 |
| `poweredLandGw` | 已取得電力的土地／選擇權／意向書 | GW | 1.5 | 否 |
| `yeActiveGw` | 年底主動電力指引 | GW | 1.85 | 是 |
| `gw2030` | 2030 年電力目標 | GW | 8 | 否 |
| `dataCenters` | 資料中心數 | 座 | 51 | 否 |
| `capexLo` | 全年資本支出指引下緣 | US$bn | 35 | 是 |
| `capexHi` | 全年資本支出指引上緣 | US$bn | 39 | 是 |
| `nextQCapexLo` | 下一季資本支出指引下緣 | US$bn | 11.5 | 否 |
| `nextQCapexHi` | 下一季資本支出指引上緣 | US$bn | 13.5 | 否 |
| `nextQIntLo` | 下一季利息指引下緣 | US$bn | 0.86 | 否 |
| `nextQIntHi` | 下一季利息指引上緣 | US$bn | 0.94 | 否 |
| `revLo` | 全年營收指引下緣 | US$bn | 12.4 | 否 |
| `revHi` | 全年營收指引上緣 | US$bn | 13.2 | 否 |
| `nextQRevLo` | 下一季營收指引下緣 | US$bn | 3.45 | 是 |
| `nextQRevHi` | 下一季營收指引上緣 | US$bn | 3.6 | 是 |
| `adjOpLo` | 全年調整後營業利益指引下緣 | US$bn | 0.96 | 否 |
| `adjOpHi` | 全年調整後營業利益指引上緣 | US$bn | 1.15 | 否 |
| `nextQAdjOpLo` | 下一季調整後營業利益指引下緣 | US$bn | 0.2 | 否 |
| `nextQAdjOpHi` | 下一季調整後營業利益指引上緣 | US$bn | 0.26 | 否 |
| `arrLo` | 期末年化經常性收入（ARR）指引下緣 | US$bn | 18.5 | 是 |
| `arrHi` | 期末 ARR 指引上緣 | US$bn | 19.5 | 是 |
| `postQNewCommit` | 季末後新增承諾 | US$bn | 25 | 否 |
| `priceUp` | 新約漲價幅度 | 比例 | 0.25 | 否 |
| `marginStep` | 新約貢獻率提升 | 文字 | 5–10 pt | 否 |
| `backlogStarted` | backlog 中已開始交付的比例（法說「>50%」） | 比例 | 0.5 | 否 |
| `inferenceArr` | 推論服務 ARR（年底） | US$bn | 0.25 | 否 |
| `otherArr` | 其他服務 ARR | US$bn | 0.4 | 否 |
| `convert` | 期後可轉債發行額 | US$bn | 3.7 | 否 |
| `convertNet` | 期後可轉債淨額 | US$bn | 3.6 | 否 |
| `convertCoupon` | 期後可轉債票面利率 | 比例 | 0.02875 | 否 |
| `convertPx` | 轉換價 | US$ | 97.85 | 否 |
| `atmShares` | 股權分銷（ATM）上限 | m 股 | 35 | 否 |
| `price0917` | 9/17 收盤價 | US$ | 79.88 | 是 |
| `tier34Active` | 主動電力位於 Tier 3/4 市場比例 | 比例 | 0.25 | 否 |
| `tier34Contracted` | 簽約電力位於 Tier 3/4 市場比例 | 比例 | 0.74 | 否 |
| `cdsMid` | CDS 中價 | bps | 720 | 否 |
| `cdsBidAsk` | CDS 買／賣價 | 文字 | 690 / 750 | 否 |
| `cdsRange` | CDS 近期區間 | 文字 | 680–855 | 否 |
| `cdsPeak` | CDS 峰值 | bps | 855 | 否 |
| `cdsPeakDate` | CDS 峰值日期 | 日期 | 2026-07-28 | 否 |
| `cdsLow` | CDS 低點 | bps | 452 | 否 |
| `cdsLowDate` | CDS 低點日期 | 日期 | 2026-06 | 否 |
| `availability` | 未動用信用額度 | US$bn | 10.014 | 否 |
| `price0918` | 9/18 收盤價 | US$ | 80.2 | 否 |
| `priceLast` | 最新收盤價 | US$ | 90.13 | 是 |
| `priceDate` | 最新收盤價日期 | 日期 | 2026-09-24 | 是 |
| `postQShortDated` | 期後公告摘要 | 文字 | 2026-09-16 公告：6/30 後持續以更高價… | 否 |
| `ttmRev` | 近十二個月營收 | US$bn | 7.59 | 是 |
| `ttmOpInc` | 近十二個月 GAAP 營業損益 | US$bn | -0.231 | 是 |
| `ttmDa` | 近十二個月折舊攤銷 | US$bn | 3.991 | 是 |
| `ttmOpLease` | 近十二個月營業租賃成本 | US$bn | 1.377 | 否 |
| `mktCapLast` | CRWV 市值（Comps 用；與同業同一收盤日 peers.priceDate，v4.3 起不隨現價更新） | US$bn | 47.12 | 是 |
| `nextEarn` | 下次財報時間 | 文字 | 2026-11 (10-Q 期限 11/16) | 否 |

### `scenarios`：三個擴張情境

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `scenarios.labels.low` | 保守情境名稱（空格前的文字會當作情境簡稱） | 文字 | 保守 4.2 GW | 必改 |
| `scenarios.labels.base` | 基準情境名稱 | 文字 | 基準 5.6 GW | 必改 |
| `scenarios.labels.high` | 積極情境名稱 | 文字 | 積極 8 GW | 必改 |
| `scenarios.accepted.low` | 保守情境：五期（首期模型部分＋4 個完整財年；目前為 FY26 下半年、FY27、FY28、FY29、FY30）各期末已驗收的主動電力 | MW 清單 | 1850、3100、3800、4200、4200 | 必改 |
| `scenarios.accepted.base` | 基準情境：同上 | MW 清單 | 1850、3100、4000、4800、5600 | 必改 |
| `scenarios.accepted.high` | 積極情境：同上 | MW 清單 | 1850、3100、4600、6300、8000 | 必改 |
| `scenarios.billableRatio.billable` | 計算「可計費 ÷ 已驗收」爬坡比例用的可計費 MW（與下一欄逐期相除） | MW 清單 | 1700、2900、4300、5900、7500 | 檢查 |
| `scenarios.billableRatio.accepted` | 同上，分母的已驗收 MW | MW 清單 | 1850、3100、4600、6300、8000 | 檢查 |
| `scenarios.leaseHighPath` | 積極情境下，尚未起租租約的新增年租金路徑；其他情境依 MW 比例縮放 | US$bn 清單 | 0.3、1.6、3、4、4.6 | 檢查 |
| `scenarios.leaseRampFloorMw` | 低於此電力時不產生新增表外租金（縮放公式的起點） | MW | 1500 | 檢查 |
| `scenarios.mw31.low` | 保守情境：模型期後一年（FY31）新增的 MW，用於 FY30 的預建支出 | MW | 0 | 檢查 |
| `scenarios.mw31.base` | 基準情境：同上 | MW | 500 | 檢查 |
| `scenarios.mw31.high` | 積極情境：同上 | MW | 1000 | 檢查 |
| `scenarios.capexTemplate.costMW` | 每 MW 建置成本（GPU＋網路＋機房內裝），各期 | US$m/MW 清單 | 34、34、34、34、34 | 檢查 |
| `scenarios.capexTemplate.customerFund` | 客戶預付占資本支出的比例，各期 | 比例清單 | 0.08、0.08、0.08、0.08、0.08 | 檢查 |
| `scenarios.capexTemplate.div` | JV 後續增資與策略投資，各期 | US$bn 清單 | 0.1、0.4、0.3、0.2、0.2 | 檢查 |

### `legacy`：舊版對照值

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `legacy.capexV14` | 舊版（v1.4）手動資本支出，只供畫面對照 | US$bn 清單 | 21、45、50、50、45 | 可沿用 |
| `legacy.interestV14` | 舊版（v1.4）手動利息，只供畫面對照 | US$bn 清單 | 1.8、2.7、2.3、2、1.8 | 可沿用 |

### `defaults`：預設假設（畫面上可調的輸入）

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `defaults.scenario` | 開啟時的預設情境（low／base／high） | 文字 | base | 檢查 |
| `defaults.lambda` | 提前支出比例：次年才上線的 MW，其建置支出落在前一年的比例 | 比例 | 0.35 | 檢查 |
| `defaults.mwYearEnd` | 各年底主動電力，以年份為鍵（例如 "2025": 850）：首期期初取首期前一財年末；GPU 汰換批次＝各年新增 MW（5a；列入滾動檢查） | MW 物件 | 物件（2023、2024、2025） | 必改 |
| `defaults.mw31` | 模型期後一年新增 MW 的預設值（情境切換時改用 scenarios.mw31） | MW | 500 | 檢查 |
| `defaults.capexFloorFY0` | 首期所屬財年的全年資本支出下限（已下單的承諾，取公司指引下緣；首期＝下限 − 年初至今實際認列） | US$bn | 35 | 必改 |
| `defaults.gpuLife` | GPU 經濟壽命（決定汰換時點與折舊） | 年 | 6 | 檢查 |
| `defaults.revScale` | 每 MW 年收入整體倍數（反向 DCF 與壓力測試用，1＝不調整） | 倍 | 1 | 可沿用 |
| `defaults.capexScale` | 每 MW 建置成本整體倍數（1＝不調整） | 倍 | 1 | 可沿用 |
| `defaults.ebStart` | 第一期 EBITDA 率 | 比例 | 0.59 | 必改 |
| `defaults.ebSteady` | 最後一期（穩態）EBITDA 率；中間各期線性內插 | 比例 | 0.65 | 檢查 |
| `defaults.services` | 非算力服務營收（軟體、儲存等），各期 | US$bn 清單 | 0.3、1、1.8、2.8、3.8 | 檢查 |
| `defaults.debtBacklog` | 新債上限：總債務不超過 backlog 的倍數 | 倍 | 1 | 檢查 |
| `defaults.ctrTerm` | 新簽合約的平均年期（決定 backlog 補入量） | 年 | 5 | 檢查 |
| `defaults.minCash` | 最低現金：每期融資後期末現金不低於此值 | US$bn | 2 | 檢查 |
| `defaults.eqPx` | 新股發行參考價（預設＝現價） | US$ | 90.13 | 必改 |
| `defaults.eqDisc` | 新股發行折價 | 比例 | 0.1 | 檢查 |
| `defaults.eqCapPct` | 每年股權募資上限（占現市值）；輸入 9 以上視為無上限 | 比例 | 0.2 | 檢查 |
| `defaults.eqCapShares` | 股權年上限的股數基礎：現市值＝發行參考價 × 此股數（5a；評價日時點，列入滾動檢查） | bn | 0.5865 | 必改 |
| `defaults.junkRate` | 股權上限用完後的高息債利率 | 比例 | 0.12 | 檢查 |
| `defaults.ppeOpen` | 最新季末固定資產毛額 | US$bn | 46.736 | 必改 |
| `defaults.jvCommit` | 已承諾的 JV 出資餘額，各期 | US$bn 清單 | 1.15、0、0、0、0 | 必改 |
| `defaults.intCal` | 第一期利息校準值（讓模型利息對上公司季度指引） | US$bn | 0.35 | 必改 |
| `defaults.rpoOpen` | 最新季末 RPO | US$bn | 103.7 | 必改 |
| `defaults.rpoPendingAdd` | 季末後新簽、尚未進 RPO 的承諾 | US$bn | 25 | 必改 |
| `defaults.rp` | RPO 在模型期內認列的比例（百分點；應與 rpo.scheduledShare 一致） | % | 84 | 必改 |
| `defaults.cash` | 最新季末現金（第一期期初現金） | US$bn | 5.5 | 必改 |
| `defaults.includeDebt` | 是否依到期表攤還既有債務（true＝是；false＝假設全數再融資） | 是／否 | 是 | 可沿用 |
| `defaults.includeAtm` | 是否計入期後股權／可轉債募資 | 是／否 | 是 | 可沿用 |
| `defaults.atm` | 期後股權／可轉債募資淨額（記在第一期） | US$bn | 6.2 | 必改 |
| `defaults.overlay` | 電力／維護成本另計（預設關；EBITDA 率已含電費，開啟會重複扣除） | 是／否 | 否 | 可沿用 |
| `defaults.cdsLink` | CDS 利差是否傳入新債利率（預設關） | 是／否 | 否 | 可沿用 |
| `defaults.cdsBaseBp` | CDS 傳入新債利率的門檻：超過此值的部分才傳入（5a） | bps | 450 | 檢查 |
| `defaults.cdsPassThrough` | CDS 超額傳入新債利率的比例（每 1bp 傳入的 bp；5a） | 比例 | 0.4 | 檢查 |
| `defaults.linkSites` | 具名站點的 MW 是否連動第一期產能下限 | 是／否 | 是 | 可沿用 |
| `defaults.linkLeaseTail` | 舊模板遺留開關，目前程式未使用 | 是／否 | 是 | 可沿用 |
| `defaults.useAvgMw` | 收入以平均在役 MW 計（true）或期末存量計（false） | 是／否 | 是 | 可沿用 |
| `defaults.billableOpen` | 最新季末可計費 MW（第一期期初） | MW | 1350 | 必改 |
| `defaults.cds` | 信用違約交換（CDS）中價 | bps | 720 | 必改 |
| `defaults.cdsBid` | CDS 買價 | bps | 690 | 必改 |
| `defaults.cdsAsk` | CDS 賣價 | bps | 750 | 必改 |
| `defaults.cdsLo` | CDS 近期區間下緣 | bps | 680 | 必改 |
| `defaults.cdsHi` | CDS 近期區間上緣 | bps | 855 | 必改 |
| `defaults.cdsDate` | CDS 報價日期與來源標記 | 文字 | 2026-09（User-provided 報價） | 必改 |
| `defaults.useFacility` | 融資時是否先動用未動用信用額度 | 是／否 | 是 | 可沿用 |
| `defaults.facility` | 未動用信用額度 | US$bn | 10 | 必改 |
| `defaults.m.accepted` | 已驗收 MW 的預設路徑（實際依所選情境覆寫） | MW 清單 | 1850、3100、4000、4800、5600 | 檢查 |
| `defaults.m.billable` | 可計費 MW 的預設路徑（實際依情境與爬坡比例覆寫） | MW 清單 | 1700、2900、3739、4495、5250 | 檢查 |
| `defaults.m.util` | 利用率，各期 | % 清單 | 95、95、94、93、92 | 檢查 |
| `defaults.m.revMW` | 每 MW 年收入，各期 | US$bn/MW 清單 | 0.0112、0.0115、0.0115、0.011、0.0105 | 必改 |
| `defaults.m.aiShare` | AI 占比，各期（目前只做範圍檢查，未參與計算） | % 清單 | 100、100、100、100、100 | 可沿用 |
| `defaults.m.fill` | 新產能簽約率：未被既有 RPO 占用的產能能賣出的比例 | % 清單 | 100、100、90、85、80 | 檢查 |
| `defaults.m.power` | 電價（overlay 開啟時才用） | $/MWh 清單 | 60、62、64、66、68 | 檢查 |
| `defaults.m.pue` | 電力使用效率 PUE（overlay 用） | 倍 清單 | 1.25、1.24、1.22、1.2、1.2 | 檢查 |
| `defaults.m.maint` | 維護成本（overlay 用） | US$m/MW 清單 | 0.15、0.15、0.16、0.16、0.17 | 檢查 |
| `defaults.m.defaultP` | 客戶違約率 | % 清單 | 0.5、1、1.5、2、2.5 | 檢查 |
| `defaults.m.recovery` | 違約回收率 | % 清單 | 50、50、50、50、50 | 檢查 |
| `defaults.m.rate` | 新債利率 | % 清單 | 9、9、9、9、9 | 檢查 |
| `defaults.terminal.residual` | 模型期末 GPU 殘值率 | % | 25 | 檢查 |
| `defaults.terminal.rerent` | 期末設備再出租率 | % | 75 | 檢查 |
| `defaults.terminal.margin` | 模型期後剩餘 RPO 的利潤率 | % | 40 | 檢查 |
| `defaults.terminal.residualLeaseYears` | 模型期後租約剩餘年數 | 年 | 12 | 檢查 |
| `defaults.sites` | 具名資料中心站點，一站一列。欄位：id 代碼、name 名稱、operator 房東／合作方、planned 契約 MW、energized 已通電 MW、accepted 已驗收 MW、billable 可計費 MW、contract 合約總值（US$bn，可無）、years 合約年期（可無）、status 狀態說明、next 下一里程碑、date 預計時間、confidence 信心（高／中／低） | 清單 | 6 筆 | 必改 |

### `valuation`：評價參數

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `valuation.price` | 現價 | US$ | 90.13 | 必改 |
| `valuation.shares` | 評價股數（含期後股權發行上限） | bn 股 | 0.587 | 必改 |
| `valuation.netDebt` | 淨負債（最新季末本金 − 現金） | US$bn | 30 | 必改 |
| `valuation.tax` | 稅率 | 比例 | 0.21 | 檢查 |
| `valuation.nol` | 期初可扣抵虧損（NOL） | US$bn | 4 | 必改 |
| `valuation.wacc` | 加權平均資金成本 WACC | 比例 | 0.11 | 檢查 |
| `valuation.nolUsePct` | NOL 每年可抵用上限占應稅所得的比例（美國 80%；依公司稅籍調整；5a） | 比例 | 0.8 | 檢查 |
| `valuation.wcPctOfRevGrowth` | 營運資金變動占營收增量的比例（DCF 自由現金流；5a） | 比例 | 0.02 | 檢查 |
| `valuation.g` | 永續成長率 | 比例 | 0.03 | 檢查 |
| `valuation.sbc` | 年度股份基礎薪酬 | US$bn | 0.7 | 必改 |
| `valuation.maintRatio` | 終值的維持性資本支出占折舊比例 | 比例 | 0.8 | 檢查 |
| `valuation.evEbitda` | EV/EBITDA 倍數 | 倍 | 6 | 檢查 |
| `valuation.evYear` | EV/EBITDA 錨定年度（1＝FY27 … 4＝FY30） | 年度代碼 | 3 | 檢查 |
| `valuation.dcfMode` | DCF 股權為負時的處理：zero＝0 截斷、option＝選擇權法 | 文字 | zero | 可沿用 |
| `valuation.sigma` | 企業價值波動率（選擇權法用） | 比例 | 0.5 | 檢查 |
| `valuation.rf` | 無風險利率（選擇權法用） | 比例 | 0.04 | 檢查 |

### `methodology`：評價方法與評等門檻

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `methodology.blendWeights.dcf` | 加權目標價中 DCF 的權重 | 比例 | 0.45 | 可沿用 |
| `methodology.blendWeights.pe` | 加權目標價中 EV/EBITDA 的權重（兩者合計 1） | 比例 | 0.55 | 可沿用 |
| `methodology.rangeMultiples` | 方法區間的 EV/EBITDA 倍數下端與上端 | 倍 清單 | 5、6 | 檢查 |
| `methodology.consensusGapTol` | 與市場共識比較的判斷句門檻：營收、EBITDA、CapEx 差距絕對值超過此比例即視為分歧（v4.3） | 比例 | 0.05 | 可沿用 |
| `methodology.profitMarginTolPt` | 利潤類（調整後 EBITDA、調整後營業利益）須附差異原因的門檻：利潤率差超過此百分點（寫成比例，0.02＝2pt）；營收、CapEx、淨負債、MW 仍用 consensusGapTol 的比例門檻；不影響分歧判斷句（v4.4） | 比例 | 0.02 | 可沿用 |
| `methodology.rating.buyUpsideMin` | 買進：空間（加權目標價 ÷ 現價 − 1）至少要達到的值 | 比例 | 0.25 | 可沿用 |
| `methodology.rating.buyTvShareMax` | 買進：終值占企業價值須低於此值 | 比例 | 0.9 | 可沿用 |
| `methodology.rating.sellUpsideMax` | 賣出：空間小於或等於此值即賣出；賣出門檻價＝現價 ×（1＋此值） | 比例（負數） | -0.15 | 可沿用 |
| `methodology.rating.equityRaiseMaxMult` | 五期股權募資超過「現市值 × 此倍數」即禁止買進 | 倍 | 1.5 | 可沿用 |
| `methodology.rating.sellUpsideMaxIfEquityOver` | 股權募資超標時，空間小於或等於此值即賣出 | 比例（負數） | -0.08 | 可沿用 |
| `methodology.rating.sellMarginAlert` | 任一情境與賣出門檻的距離小於「現價 × 此值」時，判斷句另外揭露 | 比例 | 0.05 | 可沿用 |
| `methodology.rating.tvShareWarn` | 終值占企業價值超過此值時提出警示（融資說明、檢查頁） | 比例 | 0.85 | 可沿用 |
| `methodology.checks.capexPerMwBand` | 檢查頁：模型期 CapEx 強度（每 MW 百萬美元）的合理區間下端與上端（5a） | US$m/MW 清單 | 20、45 | 檢查 |
| `methodology.checks.leaseVsCommitMin` | 檢查頁：表外租金路徑 ÷ 已承諾租約至少要達到的倍數（5a） | 倍 | 0.8 | 檢查 |
| `methodology.checks.revCapShareMax` | 檢查頁：CRWV 每 MW 計費收入 ÷ Tokenomics 客戶付費 token 營收（IF_RevGWFleet）的上限，超過即警示（W4） | 比例 | 0.5 | 檢查 |
| `methodology.checks.unsignedRevShareMax` | 檢查頁：後段年度依賴未簽約收入的比例上限（5a） | 比例 | 0.5 | 檢查 |
| `methodology.checks.siteRentGapMax` | 檢查頁：站點租賃五期租金可能低估的金額上限（5a） | US$bn | 10 | 檢查 |
| `methodology.checks.rentVsBenchMin` | 檢查頁：模型每 MW 年租金至少要達到「市場基準 × 第三方占比」的比例（5a） | 比例 | 0.8 | 檢查 |
| `methodology.perMw._note` | 每 MW 方法開關的說明（不進程式；W2） | 文字 | 每 MW 方法開關（W2）：capex＝tokeno… | 可沿用 |
| `methodology.perMw.capex` | 每 MW 資本支出方法（W2）：tokenomics＝Σ 新增世代占比 × IF_CapexIT；legacy＝scenarios.capexTemplate.costMW | 代碼 | tokenomics | 檢查 |
| `methodology.perMw.cost` | 營運成本方法（W2）：bottomUp＝Tokenomics 電費、IT 維護、人員軟體、稅險 × 平均在役 MW＋管銷率；ebitdaPct＝起始→穩態 EBITDA 率線性 | 代碼 | bottomUp | 檢查 |
| `methodology.perMw.revenue` | 每 MW 收入方法（W2／W4）：tkAnchor＝Σ 平均在役占比 × IF_HoldEcon × 定價倍數 k（pricing.anchorMultiple；W4 預設）；gpuHr＝pricing.gpuHr × 每 MW GPU 數 × 8,760；legacy＝defaults.m.revMW（對照） | 代碼 | tkAnchor | 檢查 |

### `peers`：同業比較（Comps）

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `peers.priceDate` | 同業市值的收盤日 | 日期 | 2026-09-21 | 必改 |
| `peers.priceSource` | 同業市值來源 | 文字 | S&P Global via stockanalys… | 必改 |
| `peers.list` | 同業，一家一列（HTML Comps 分頁與 Excel「可比公司」頁共用）。欄位：ticker 代號、name 公司名、labelHtml／labelXlsx 兩邊顯示的名稱、roleHtml 定位說明（HTML）、mkt 市值、netDebt 淨負債、rev 近十二個月營收、opl 營業租賃負債、opInc GAAP 營業損益、da 折舊攤銷（以上 US$bn）、asOf 資料期、noteHtml／noteXlsx 兩邊的備註。EV＝市值＋淨負債、EBITDA＝營業損益＋折舊攤銷，由程式計算 | 清單 | 4 筆 | 必改 |
| `peers.textHtml.headerTip` | HTML Comps 表標題的浮動說明 | 文字 | 市值為 2026-09-21 收盤（S&P Glob… | 必改 |
| `peers.textHtml.readingTip` | HTML「讀法」段落的浮動說明（折價來源） | 文字 | 這個折價有三個來源，不能全部解讀為便宜：(1) 分母… | 必改 |
| `peers.textHtml.caveat` | HTML Comps 表下方的「口徑與限制」 | 文字 | 口徑與限制：市值為 2026-09-21 收盤，負債… | 必改 |
| `peers.textXlsx.subtitle` | Excel「可比公司」頁第 2 列說明 | 文字 | EV＝市值＋（有息負債−非受限現金）；含租賃＝EV＋… | 必改 |
| `peers.textXlsx.notes` | Excel「可比公司」頁表下方的讀法說明（每句一列） | 文字清單 | 讀法：CRWV 的 TTM EV/Sales 約為同…、(2) CRWV 用租約取得機房，加入營業租賃負債後…、前瞻：NBIS 前瞻 EV/Sales 約 9.3x… | 必改 |

### `quarterly`：季度層（v4.4）

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `quarterly._note` | 季度層的說明文字（不進程式） | 文字 | 季度層（v4.4）：由年度模型拆分，只用於追蹤，不影… | 可沿用 |
| `quarterly.quarters` | 追蹤的季度，一季一列：key 季別代碼（與共識檔 quarterlyEstimates 的鍵相同，例如 2026Q3）、label 顯示名稱、period 所屬模型期（0＝第 1 期…）、reportNote 財報日說明（選填）。季度加總必須等於所屬模型期的年度數字 | 清單 | 6 筆 | 必改 |
| `quarterly.periodNames` | 各所屬模型期的顯示名稱（依期別順序，例如 2H26、FY27） | 文字清單 | 2H26、FY27 | 必改 |
| `quarterly.focus` | 焦點季（「季度追蹤」第一屏與一頁摘要驗證點顯示的季度；財報後改為下一季） | 季別代碼 | 2026Q3 | 必改 |
| `quarterly.keyMetrics` | 一頁摘要「驗證點」列出的指標（2–3 項；metrics 的 key） | 文字清單 | revenue、adjOpInc、capex | 檢查 |
| `quarterly.driver` | 拆分依據：type＝mw 時依 MW 內插（endStart＝第 1 期期初 Accepted MW、endStartNote 來源）；type＝none（沒有 MW 資料的公司）時營收依 revenueSplit、CapEx 平分，MW 列顯示「不適用」 | 物件 | 物件（type、endStart、endStartNote） | 檢查 |
| `quarterly.revenueSplit` | 各所屬模型期的營收拆法：anchor＝從 revenueAnchor 逐季線性爬升、driverAvg＝依平均在役 MW、equal＝平分 | 文字清單 | anchor、driverAvg | 檢查 |
| `quarterly.revenueAnchor` | 營收拆法 anchor 的起點（最新一季實際營收）：value、label、tag | 物件 | 物件（value、label、tag） | 必改 |
| `quarterly.metrics` | 追蹤的指標，一項一列：key（revenue、adjEbitda、ebitdaMargin、adjOpInc、capex、mw 之一）、label、unit（US$bn／%／MW）、gap 差距寫法（ratio＝比例，用於營收、CapEx；diff＝金額差；pt＝百分點）、marginOf 利潤類另列利潤率百分點時的分母指標（revenue）、consensus 共識檔季度欄位名（derived＝由共識 EBITDA ÷ 營收換算）、consensusSecondary／consensusSecondaryLabel 只列不比較的共識欄位、actualLabel 實際數列名稱 | 清單 | 6 筆 | 檢查 |
| `quarterly.capexSplit` | CapEx 拆法：guidanceAnchor＝有季度指引的季取指引中點、其餘季分配期間餘數（依新增 MW）；driverAdds（或不填）＝全部依新增 MW | 文字 | guidanceAnchor | 檢查 |
| `quarterly.guidanceText` | 文字型指引（例如「調整後營業利益率 low teens」），{季別: {指標: {text, source, tag}}}：只列不計差距；完整支援（年增率、利潤率區間）列入待辦 5 | 物件 | 物件（2026Q4） | 必改 |
| `quarterly.guidance` | 季度指引（公司預估），{季別: {指標: [低, 高]}}；沒有指引的季度或公司不填，顯示「不適用」 | 物件 | 物件（2026Q3） | 必改 |
| `quarterly.guidanceMeta` | 季度指引的來源與標記 | 物件 | 物件（source、tag） | 必改 |
| `quarterly.periodGuidance` | 期間隱含指引，{模型期: {指標: {range: [全年低, 全年高], less: 已實現部分}}}：只用於判斷季度差距是否為「拆法」 | 物件 | 物件（0） | 必改 |
| `quarterly.periodGuidanceNote` | 上一欄的說明文字 | 文字 | 全年指引 − 1H 實際＝下半年隱含指引 [Deri… | 必改 |
| `quarterly.consistency` | 建置檢查：[路徑 A, 路徑 B, 倍數（選填）]，兩者（B × 倍數）必須相同，否則建置失敗（防止同一數字在兩處不一致） | 清單 | 14 筆 | 檢查 |
| `quarterly.actuals` | 季度實際數（公司公布後填入；預設空白＝待公布），{季別: {各指標, source, date, tag}}；Excel 對應「輸入與假設」J 區藍字格 | 物件 | 物件（2026Q3、2026Q4、2027Q1、2027Q2、2027Q3、2027Q4） | 必改 |

### `varianceReasons`：差異原因（v4.4）

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `varianceReasons._note` | 差異原因的說明文字（不進程式） | 文字 | 差異原因（已決定事項 2）：差距超過 methodo… | 可沿用 |
| `varianceReasons.list` | 差異原因（已決定事項 2），一筆一列：scope（annual 年度共識對照／quarter 季度）、period（FY27、2026Q3 或 *）、metric（年度：rev、ebitda、capex、nd；季度：metrics 的 key）、vs（consensus、guidance、actual 或 *）、type（觀點／已知限制）、text 一句原因，{路徑:格式} 由模型數字帶入。「拆法」由程式判定，不需填。差距超過 methodology.consensusGapTol 卻沒有原因時建置失敗；perMw（W2，選填）＝只在 methodology.perMw 相符時適用的條件，排在前面者優先 | 清單 | 11 筆 | 檢查 |

### `texts`：公司特有的說明文字（v4.5；隨資料更新）

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `texts.fy0EquityNote` | 首期股權／可轉債的組成說明（含年初至今與首期模型的金額；v4.5；可用期間佔位符 «P0»、«YTD»、«STUB»） | 文字 | «P0»＝«YTD» 私募股權 2.982＋«STU… | 必改 |
| `texts.sourceLine` | 頁首的資料來源一行（例如最新 10-Q、法說、期後 8-K；v4.5） | 文字 | Q2 2026 10-Q + 法說 + 9/17 8… | 必改 |
| `texts.cashTaxNote` | 年初至今現金稅的說明（損益與評價頁；v4.5） | 文字 | H1 現金稅約 0.1bn | 必改 |
| `texts.mwYearEndNotes` | 各年底主動電力的來源說明，以年份為鍵（Excel「輸入與假設」說明欄；5a） | 物件（文字） | 物件（2023、2024、2025） | 必改 |

### `tokenomics`：Tokenomics 取數層（W1；快照檔、版本與引用名稱）

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `tokenomics._note` | Tokenomics 取數層的說明（不進程式；W1） | 文字 | 算力相關的產業與物理層資料改引用 Tokenomic… | 可沿用 |
| `tokenomics.snapshotFile` | Tokenomics 快照檔路徑（tools/tokenomics/import_tokenomics.py 產生；Excel「Tokenomics_取數」分頁讀此檔；W1） | 路徑 | data/tokenomics_snapshot_v… | 可沿用 |
| `tokenomics.version` | 快照的 Tokenomics 版本（model/CURRENT 的版本號） | 文字 | v5.31 | 可沿用 |
| `tokenomics.commit` | 快照的 Tokenomics commit SHA | 文字 | f16f161a28b5d43e00e89aa019… | 可沿用 |
| `tokenomics.names` | 引用的 Tokenomics 名稱（只限 IF_、L1_；清單檔 data/tokenomics_names.txt） | 清單 | IF_RacksPerGW、IF_GPUsPerGW、IF_FacilityGW、IF_CapexIT、IF_CapexFacility、IF_CapexTotal、IF_HoldAcct、IF_HoldEcon、IF_GPUhrEcon、IF_PowerCost、IF_Util、L1_FacCapexMW、L1_GPUhr_GB200_vsCW、L1_GPUhr_GB300_vsBE、L1_RevGW_Fleet_VR200、IF_DeprLifeIT、IF_DeprIT、IF_DeprFac、IF_AvgDraw、IF_PowerPrice、IF_MaintIT、IF_MaintFac、IF_StaffSW、IF_TaxIns、IF_OpexGW、IF_RevGWFleet、IF_MaintITWarr、IF_MaintITPost、IF_WarrantyYrs | 檢查 |
| `tokenomics.optional` | 其中 Tokenomics 尚未提供時記為 missing 的名稱（v5.25 預計新增） | 清單 |  | 檢查 |

### `fleet`：世代組合（W2；公司專屬）

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `fleet._note` | 世代組合的說明（不進程式；W2） | 文字 | 世代組合（W2；公司專屬）：世代名稱與 Tokeno… | 可沿用 |
| `fleet.generations` | 世代清單（與 Tokenomics 快照世代同名；順序＝由舊到新，汰換由最舊世代先出） | 文字清單 | Hopper H100、GB200 NVL72、GB300 NVL72、VR200 NVL72、Rubin Ultra（MGX NVL 單架，72 … | 檢查 |
| `fleet.openMix.asOf` | 期初在役機隊的日期（最新已申報季末） | 日期 | 2026-06-30 | 必改 |
| `fleet.openMix.activeMW` | 期初在役主動電力（世代組合起點；列入滾動檢查） | MW | 1500 | 必改 |
| `fleet.openMix.mix` | 期初在役 MW 的世代占比，{世代: 比例}，合計 1 | 物件（比例） | 物件（Hopper H100、GB200 NVL72、GB300 NVL72） | 必改 |
| `fleet.openMix.tag` | 期初世代占比的來源標記 | 文字 | [Assumed] | 必改 |
| `fleet.openMix.source` | 期初世代占比的來源與方法 | 文字 | W1 第 6b 步（coreweave/data/p… | 必改 |
| `fleet.newMix` | 五期（首期模型部分＋4 個完整財年；目前為 FY26 下半年、FY27、FY28、FY29、FY30）各期新增 MW（含汰換補回）的世代占比，每期 {世代: 比例}，合計 1 | 清單 | 5 筆 | 檢查 |
| `fleet.newMixNote` | 新增世代占比的依據 | 文字 | 工作單 W2 預設 [Assumed]：2H26 G… | 檢查 |
| `fleet.newMixAlt.label` | 世代組合敏感度（替代路徑）的名稱 | 文字 | Rubin Ultra 版（FY29–FY30 新增… | 檢查 |
| `fleet.newMixAlt.mix` | 世代組合敏感度的各期新增世代占比（格式同 newMix；只作敏感度） | 清單 | 5 筆 | 檢查 |
| `fleet.newMixAlt.tag` | 替代路徑的標記 | 文字 | [Assumed]（只作敏感度） | 檢查 |

### `pricing`：GPU 小時價格與對照價格（W2）

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `pricing._note` | GPU 小時價格與對照價格的說明（不進程式；W2） | 文字 | GPU 小時價格（W2，公司專屬）：gpuHr＝合約… | 可沿用 |
| `pricing.gpuHr` | GPU 小時合約價，{世代: {base, low, high, source, date, tag}}（US$/GPU-hr）；空白＝不適用（revenue=gpuHr 時必填所有在役世代） | 物件 | 物件（） | 必改 |
| `pricing.anchorMultiple._note` | Tokenomics 錨的公司因素說明（不進程式；W4） | 文字 | W4 r2（已決定事項 14）：每 MW 年收入（1… | 可沿用 |
| `pricing.anchorMultiple.long.base` | 定價倍數 k_長約 基準（市場長約價 ÷ Tokenomics 同世代持有成本；W4） | 倍 | 0.76 | 檢查 |
| `pricing.anchorMultiple.long.low` | k_長約 區間下緣（敏感度） | 倍 | 0.7 | 檢查 |
| `pricing.anchorMultiple.long.high` | k_長約 區間上緣（敏感度） | 倍 | 1 | 檢查 |
| `pricing.anchorMultiple.long.tag` | k_長約 的資料標記 | 文字 | [Analogy] | 可沿用 |
| `pricing.anchorMultiple.long.note` | k_長約 基準與區間的依據 | 文字 | 基準＝IREN–Microsoft GB300 五年… | 可沿用 |
| `pricing.anchorMultiple.spot.base` | 定價倍數 k_現貨 基準（市場現貨價 ÷ Tokenomics 同世代持有成本；W4） | 倍 | 1.76 | 檢查 |
| `pricing.anchorMultiple.spot.low` | k_現貨 區間下緣（敏感度） | 倍 | 1.5 | 檢查 |
| `pricing.anchorMultiple.spot.high` | k_現貨 區間上緣（敏感度） | 倍 | 2.3 | 檢查 |
| `pricing.anchorMultiple.spot.tag` | k_現貨 的資料標記 | 文字 | [Analogy] | 可沿用 |
| `pricing.anchorMultiple.spot.note` | k_現貨 基準與區間的依據 | 文字 | 基準＝H100 Silicon Data 現貨指數 … | 可沿用 |
| `pricing.anchorMultiple.longShare.method` | RPO 覆蓋率對照列的算法（rpoCover＝RPO 涵蓋的產能 ÷ 在役計費產能；r2 起只作對照、不驅動） | 代碼 | rpoCover | 檢查 |
| `pricing.anchorMultiple.longShare.tag` | RPO 覆蓋率對照列的資料標記 | 文字 | [Derived] | 可沿用 |
| `pricing.anchorMultiple.longShare.formula` | RPO 覆蓋率對照列算式說明（不進程式） | 文字 | 對照列（不驅動）：RPO 涵蓋的產能 ÷ 在役計費產… | 可沿用 |
| `pricing.anchorMultiple.long.sensMedian` | k_長約 三筆長約中位數（敏感度；W4 r2） | 倍 | 0.89 | 檢查 |
| `pricing.anchorMultiple.long.refEvidence` | k_長約 的基準證據（evidence.label；成本情境重算時用其世代的 Tokenomics 成本比例） | 文字 | IREN–Microsoft GB300 五年約 | 檢查 |
| `pricing.anchorMultiple.spot.refEvidence` | k_現貨 的基準證據（evidence.label；成本情境重算時用其世代的 Tokenomics 成本比例） | 文字 | H100 Silicon Data 現貨指數 | 檢查 |
| `pricing.anchorMultiple.onDemandShare.base` | 隨需（現貨）占在役計費產能比例，基準（W4 r2；k＝隨需占比 × k_現貨＋（1 − 隨需占比）× k_長約） | 比例 | 0 | 檢查 |
| `pricing.anchorMultiple.onDemandShare.sens` | 隨需占比敏感度 | 清單 | 0.1、0.2 | 檢查 |
| `pricing.anchorMultiple.onDemandShare.tag` | 隨需占比的資料標記 | 文字 | [Assumed] | 可沿用 |
| `pricing.anchorMultiple.onDemandShare.note` | 隨需占比依據 | 文字 | 隨需（現貨）占在役計費產能的比例；公司未揭露（S-1… | 可沿用 |
| `pricing.anchorMultiple.contractMix` | 公司合約組合事實（label、value、unit、tag、source、url、date、retrieved；只列，不入公式） | 清單 | 3 筆 | 必改 |
| `pricing.anchorMultiple.evidence` | k 證據表：label、gen、price、unit、tkName（IF_GPUhrEcon／IF_HoldEcon）、contract、term、use（long／spot／range／list）、tag、source、url、date、retrieved、note；倍數在 Excel 以 TK_ 名稱計算 | 清單 | 11 筆 | 必改 |
| `pricing.anchorMultiple.notFound` | 找不到的資料（試過的來源；W4） | 清單 | VR200 NVL72 長約或多年期合約的每 GPU…、GB200 NVL72 長約價：找不到可換算的揭露（…、其他 neocloud 對 hyperscaler／… | 可沿用 |
| `pricing.peerRevPerMw` | 同業每 MW 年收入對照列（label、value、unit、tag、source、date、url、note；不入損益） | 清單 | 1 筆 | 必改 |
| `pricing.marketRefs` | 各世代市場 GPU 小時價格對照列（gen、label、value、basis、tag、source、date、url；不入損益） | 清單 | 6 筆 | 必改 |

### `companyAdjust`：公司實況驗證與公司調整（W5）

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `companyAdjust._note` | 公司實況驗證與公司調整的說明（不進程式；W5，已決定事項 15） | 文字 | W5（2026-10-08，已決定事項 15）：每個… | 可沿用 |
| `companyAdjust.gapTol` | 驗證門檻：差距 ≤ 此值直接用 Tokenomics 值（規則 1）；超過時須有證據的機制才調整（W5） | 比例 | 0.1 | 可沿用 |
| `companyAdjust.existingK.on` | 既有合約 k 開關（1＝最新季末在役 MW 按 Q2 實現單價 ÷ Q2 錨；0＝全部按 k_新約） | 代碼 | 1 | 檢查 |
| `companyAdjust.existingK.tag` | 既有合約 k 的資料標記 | 文字 | [Derived] | 可沿用 |
| `companyAdjust.existingK.formula` | 既有合約 k 的算式說明（不進程式） | 文字 | k_既有＝（Q2 營收 × 4 ÷ Q2 平均在役 … | 可沿用 |
| `companyAdjust.existingK.appliesTo` | 既有合約 k 適用的 MW 說明（不進程式） | 文字 | 6/30 在役的 1,500 MW（既有合約）；汰換… | 可沿用 |
| `companyAdjust.existingK.reverts` | 何時回到 k_新約（不進程式） | 文字 | 該批 MW 汰換時回到 k_新約（市場長約證據）；不… | 可沿用 |
| `companyAdjust.existingK.assumptions` | Q2 計費比例、利用率、服務收入的假設說明（不進程式） | 文字 | Q2 計費比例＝季末 Billable ÷ 季末在役… | 可沿用 |
| `companyAdjust.existingK.evidence` | 既有合約 k 的證據 id（companyAdjust.evidence） | 清單 | crwvQ2Rev、crwvContractDur | 必改 |
| `companyAdjust.newK.adjBase` | 新約價格調整基準（相對 k_長約；只作用於長約部分） | 比例 | 0 | 檢查 |
| `companyAdjust.newK.adjSens` | 新約價格調整敏感度（公司說法） | 比例 | 0.25 | 必改 |
| `companyAdjust.newK.tag` | 新約價格調整的資料標記 | 文字 | [Interested-party] | 可沿用 |
| `companyAdjust.newK.note` | 新約 k 的說明（不進程式） | 文字 | 新合約 k＝隨需占比 × k_現貨＋（1 − 隨需占… | 必改 |
| `companyAdjust.newK.evidence` | 新約 k 的證據 id | 清單 | crwvPriceUp、crwvShortDated | 必改 |
| `companyAdjust.spotCw.evidenceLabel` | 公司短天期合約證據（pricing.anchorMultiple.evidence 的 label；隨需敏感度用其倍數） | 文字 | CoreWeave Q3 短天期合約（3–6 個月，… | 必改 |
| `companyAdjust.spotCw.od` | 公司短天期敏感度的隨需占比 | 比例 | 0.1 | 檢查 |
| `companyAdjust.spotCw.tag` | 公司短天期證據的標記 | 文字 | [Interested-party] | 可沿用 |
| `companyAdjust.opexScale.base` | 由下而上營運成本倍數基準（1＝Tokenomics 值；敏感度改為 Q2 實際比率） | 倍 | 1 | 可沿用 |
| `companyAdjust.opexScale.tag` | 營運成本倍數的資料標記 | 文字 | [Assumed] | 可沿用 |
| `companyAdjust.opexScale.note` | 營運成本倍數的說明 | 文字 | 由下而上營運成本（電費、IT 維護、人員軟體、稅險；… | 可沿用 |
| `companyAdjust.capexActual.mwStart` | 年初至今期初主動電力（前一財年末；期間標籤取日曆的年初至今標籤） | MW | 850 | 必改 |
| `companyAdjust.capexActual.mwEnd` | 年初至今期末主動電力（最新已申報季末） | MW | 1500 | 必改 |
| `companyAdjust.capexActual.techEquip` | 技術設備 PP&E 毛額 [期初, 期末] | US$bn | 20.903、33.823 | 必改 |
| `companyAdjust.capexActual.dcEquip` | 資料中心設備與租賃改良 PP&E 毛額 [期初, 期末]（只列） | US$bn | 2.842、5.997 | 必改 |
| `companyAdjust.capexActual.cip` | 在建工程 [期初, 期末]（只列） | US$bn | 9.376、11.918 | 必改 |
| `companyAdjust.capexActual.tag` | 資本支出驗證資料的標記 | 文字 | [Verified] | 可沿用 |
| `companyAdjust.capexActual.note` | 資本支出驗證資料的來源說明 | 文字 | 10-Q Q2 2026 附註 5（PP&E 毛額：… | 必改 |
| `companyAdjust.capexActual.evidence` | 資本支出驗證的證據 id | 清單 | crwvPPE、crwvCapex | 必改 |
| `companyAdjust.params` | 驗證表逐列文字：key、label、unit、tkName、actual、tag、mechanism、adjust、rule（1／3／4a）、evidence（數值列由 Excel「公司實況驗證」頁與 HTML 以同一算式產生） | 清單 | 11 筆 | 必改 |
| `companyAdjust.evidence` | 證據清單：id、text、source、url、date、retrieved、tag | 清單 | 15 筆 | 必改 |
| `companyAdjust.notFound` | 找不到的資料（試過的來源；W5） | 清單 | CRWV 伺服器／GPU 保固或維護合約條款：10-…、電費、維護、財產稅、保險的金額：10-Q／10-K …、CRWV 新長約的每 MW 價格或合約金額 ÷ MW…、主動電力（active power）的 IT／設施口… | 必改 |
| `companyAdjust.q2Notes` | Q2 逐項對帳各列的差異說明（鍵＝列名稱） | 物件（文字） | 物件（Q2 對帳｜每 MW 年收入（算力＋服務）、Q2 對帳｜每 MW 營運成本（租金前、不含管銷；Q2 含變動租賃）、Q2 對帳｜每 MW 固定租金、Q2 對帳｜每 MW 管銷（扣 SBC）、Q2 對帳｜每 MW EBITDA（Q2＝調整後 EBITDA）、Q2 對帳｜EBITDA 率（差距＝百分點）） | 必改 |

### `costs`：由下而上營運成本口徑（W2）

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `costs._note` | 由下而上營運成本的公司口徑說明（不進程式；W2） | 文字 | 由下而上營運成本的公司口徑（W2）：管銷率＝（銷售行… | 可沿用 |
| `costs.sgaBasis` | 管銷率口徑（W2）：exSbc＝扣 SBC（預設）；gaap＝GAAP 含 SBC（敏感度） | 代碼 | exSbc | 檢查 |

## v4.0 架構：公司資料單一來源
- **company.json**：所有公司原始輸入（HTML 引擎與 Excel 共用）。換公司時先改這個檔；衍生值（情境 Billable 比率、Q3 新增 RPO 權重、債務合計與平均利率）留在 segA 開頭由程式推導。
- HTML：`build_html_portable.py` 把 company.json 注入為 `COMPANY_DATA`，segA 開頭讀取。Excel：`build_xlsx.py` 開頭讀同一檔（75 項輸入，百分點欄位以 `PCT_()` 轉成比例）。
- 頂層鍵：

| 鍵 | 型別 |
|---|---|
| `meta` | dict |
| `periods` | list |
| `periodYears` | list |
| `actual1H` | dict |
| `historicalPL` | list |
| `rpo` | dict |
| `leases` | dict |
| `debt` | dict |
| `latestQuarter` | dict |
| `callFacts` | dict |
| `scenarios` | dict |
| `legacy` | dict |
| `defaults` | dict |
| `valuation` | dict |
| `methodology` | dict |
| `peers` | dict（v4.2） |

### 建置時的兩個陷阱（勿改）
1. 模板函式庫在 segA、segB 的插入點前都是**未結束的 var 宣告鏈**，所以注入的第一句必須是 `COMPANY_DATA = {…};`，segB 第一行必須是 `PERIOD_LABELS = …,`（兩者都不可加 `var`）。
2. 模板函式庫本身已宣告 `CO`，且其片段仍以 `Qk` 引用預設值：因此資料變數名為 `COMPANY_DATA`，並保留 `Qk = DEFAULTS` 別名。

### 識別字改名對照（rename_ids.py；只改程式碼，不改畫面文字）
| 舊名 | 新名 |
|---|---|
| `zk` | `PERIODS` |
| `Lq` | `PERIOD_YEARS` |
| `Bk` | `UPDATE_DATE` |
| `Vk` | `SOURCE_ORDER_NOTE` |
| `H1` | `ACTUAL_1H` |
| `Uk` | `RPO_SCHEDULED_SHARE` |
| `Wk` | `RPO_BUCKET_W` |
| `Gk` | `RPO_Q3ADD_W` |
| `Kk` | `LEASE_CASH_ON_BAL` |
| `qk` | `DEBT_AMORT` |
| `Jk` | `LEASE_AFTER_FY30` |
| `Yk` | `LATEST_Q` |
| `Xk` | `CALL_FACTS` |
| `Zk` | `SCENARIOS` |
| `Qk` | `DEFAULTS` |
| `DBT` | `DEBT_TOOLS` |
| `LSE` | `LEASE_FACTS` |
| `cM` | `VAL_DEFAULTS` |
| `oM` | `PERIOD_LABELS` |
| `sM` | `HIST_PL` |
| `dA` | `runFunding` |
| `TM` | `runValuation` |
| `rvQ` | `reverseDcf` |
| `pA` | `sensitivities` |
| `oA` | `rpoBridge` |
| `evGrid` | `evAnchorGrid` |
| `gM` | `dcfValue` |
| `vM` | `evEbitdaLeg` |
| `xM` | `blendCall` |
| `hM` | `forwardPL` |
| `pM` | `shareCount` |
| `SM` | `dcfGridWaccG` |
| `yM` | `BLEND_W` |

尚未改名（用途未逐一確認，留待後續）：segA 的 eA、tA、nA、rA、iA、aA、sA、cA、lA、uA、fA；segB 的 uM、dM、mM、mR、wM、EM；segC 的 AM；segD 的 zM；tail 的 BM、VM；格式化函式 Y、mA、hA（模板片段也使用，不可改）。

### 尚未移入 company.json
「來源」頁與站點說明、法說對照表等長文字（同業市值與賣方目標價的來源引用句已於 v4.3 改讀 company.json 與共識檔），以及 `latestQuarter`／`callFacts` 中「程式使用：否」的數字在畫面上的文字寫法。換公司時仍需手動改這些地方。同業 Comps 與評等門檻已於 v4.2 移入（`peers`、`methodology.rating`）。

### v4.0 核對工具
- `load_engine.js`：在 node 載入 company.json＋segA＋segB（cmp31.js 使用）。
- `xl_diff.py a.xlsx b.xlsx [--values]`：Excel 逐格比對（公式或重算後的值）。
- `crawl.py <html> <out.json>`：逐頁逐分頁擷取畫面文字，用於比對兩版 HTML 的顯示是否相同。2026-09-25 修正：原本只擷取第一個 `<main>`（資金模型頁的一個區塊），總結頁與損益與評價頁的內容、以及多數次級分頁實際上沒有被比對；現改為擷取整頁可見文字並逐一展開兩層選單（v4.0 為 32 個視圖、約 111,000 字）。
- 結構性修改的驗收標準：cmp31 三情境全數一致；crawl 比對除版本紀錄外 0 差異；xl_diff 除版本紀錄外 0 差異（v4.2 起有新增列時用 `--by-label`）。
