# Oracle 收支模型 v0.2 原始碼包

Oracle（ORCL）收支與評價模型；由 `nebius/` @ `642d144`（CRWV v4.5＋Nebius v0.1b）複製建立。下方各節為模板沿革與技術說明（以 CoreWeave 為例），引擎與工具仍適用；Oracle 新增的結構（OCI 以 MW × 每 MW 收入、首期以 Q1 OCI 校準可計費 MW、傳統事業四線、客戶出資與預付重大財務組成、股利、投資級融資瀑布、強制轉換特別股、未起租租賃排程、CAPM WACC、分部 EV/EBITDA；v0.1c：EBITDAR 率 − 固定租金、MW 觸頂後穩態 GPU 汰換、終值以末期 UFCF 為基準）見交接檔 `docs/handoff/20261008_Oracle收支模型_交接檔_v0_2.md` 與下方「company.json 欄位說明」；v0.2 新增建設延誤模組（計費 MW 平移、GPU 資本支出照原時程與閒置資本、租約起租連動、延誤罰則；Excel「運營_產能與收入」（E）區）與租賃負債／租賃調整後槓桿（投資級上限 ≤ 4.5×；「各期收支」租賃負債區、「資產負債_新債與新股」槓桿列）。成品名稱＝`更新日_<meta.company>收支模型_v版本`（目前 `dist/20261008_Oracle收支模型_v0_2.html`／`.xlsx`）。升版驗收的預期差異清單在 `scripts/expect/`（v0.2：`v0_2_vs_v0_1.txt`；延誤 0 不變性：`v0_2_delay0_vs_v0_1.txt`）。`scripts/calib_ebitdar.js`：EBITDAR 率校準（`defaults.ebitdarAdj`；verify.sh 第 0c 項）。

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
| build_xlsx.py | 產生 Excel：`python3 build_xlsx.py [輸出.xlsx]`（預設 `out/<更新日>_<meta.company>收支模型_v<版本>.xlsx`；版本號、更新日、市價日寫在導覽 A1（v4.3 起讀 vlog.py 與 company.json）；自同目錄讀取 vlog.py、rv_snap.json、company.json 與 `meta.consensusFile` 指定的共識資料檔；群組資訊寫到 `out/outline.json`） |
| fix_outline.py | LibreOffice 重算後，補回 Excel 群組按鈕位置與收合狀態：`python3 fix_outline.py 檔案.xlsx [outline.json]`（預設讀 `out/outline.json`；v3.4 修正：outlinePr 依 schema 放在 tabColor 之後） |
| fix_datatable.py | v4.1：LibreOffice 重算後，把模擬運算表（情境區間）由 `TABLE()` 一般公式還原為 Excel 的 `dataTable` 公式，保留算出的值：`python3 fix_datatable.py 檔案.xlsx` |
| verify_ooxml.py | Excel 嚴格結構檢查（XML 格式、工作表與 sheetPr 子元素順序；v4.1 起另檢查模擬運算表存在且無殘留 `TABLE()`）；v3.4 新增 |
| calendar_q.py | v4.5（5a-1）：期間滾動。由 `company.json` → `calendar`（財年結束月份、最新已申報／已公布季度、模型年數、目標價時點）推算期間標籤、期間長度、評價日、各期期末日與折現年數（以月計）、目標價時點與 EV/EBITDA 折回調整；`build_xlsx.py`、`build_html_portable.py`、`load_engine.js` 都讀同一份推算結果（`COMPANY_DATA.cal`；Excel「輸入與假設」F 區「期間與日期」）。`periods`、`periodYears`、`valuation.optT` 不再存於 company.json。`python3 calendar_q.py` 印出推算結果 | 說明文字中的期間字樣以佔位符書寫（«YTD»、«STUB»、«VMD»、«P0»…，`tokens`），build_xlsx 存檔前、cmp31 查列名稱時換成目前日曆的字樣；目前日曆下換出的文字與 v4.4 相同。期間特有的敘述句與數字放 `company.json` → `texts`、`ytdActual.notes`，隨每季資料更新 || vlog.py | 版本紀錄資料，Excel 的唯一來源（v3.4 起 build_xlsx.py 直接 import，不再內嵌副本）；HTML 的 VLOG 在 tail.js 末尾，需同步 |
| rv_snap.json | 反向 DCF 快照（Excel「評價_反向DCF」與「摘要」頁使用）。v4.5 起由 `scripts/rv_solve.py` 以 Excel 求解產生（verify.sh 步驟 3b；有變動時自動重建 Excel），不再用 JS 引擎（`scripts/rv_snap.js` 已移除） |
| scripts/rv_solve.py | v4.5：反向 DCF 求解。以 LibreOffice（UNO，`scripts/uno_q.py`）開啟建好的 Excel，改「每 MW 年收入倍數」「每 MW 建置成本倍數」「穩態 EBITDA 率」並重算，依與 HTML `reverseDcf` 相同的規則（同一上下界、22 次對半）二分法反解，寫入 rv_snap.json；結果有變動時代碼 3 |
| scripts/uno_q.py | v4.5：以 LibreOffice UNO 驅動 Excel 的共用模組（依 A 欄列名稱找格、改值、讀重算後的值；不存檔） |
| data/consensus_crwv_20260925.json | 市場共識資料檔（v4.3；Andy 查證後提供，只讀，不得增補或修改數字）；路徑寫在 `company.json` → `meta.consensusFile`，HTML 建置時併入 `COMPANY_DATA.consensus`，Excel 直接讀取 |
| xl_diff.py | Excel 比對：`python3 xl_diff.py 舊.xlsx 新.xlsx [--values] [--by-label] [--ignore=工作表,…] [--expect=清單檔]`（v4.5 `--expect`：預期差異清單，每行「工作表!儲存格 原因」，列出但不計入差異數）；（v4.2 `--by-label`：新增列造成位移時以「區段＋列名稱」配對，見「v4.2 核對方式」；v4.3 `--ignore`：兩檔都略過指定工作表） |
| scripts/fields_doc.py | v4.2：產生下方「company.json 欄位說明」的表格（`python3 scripts/fields_doc.py > out/fields.md`）；company.json 有欄位沒寫說明即報錯 |
| xlx.py／cmp31.js | 一致性核對：xlx.py 依情境重算 Excel 並擷取數值（暫存檔在 `out/`，以 `scripts/recalc.py` 重算）；cmp31.js 以 HTML 引擎逐列比對（三情境各 313 項，含 v4.1 目標價區間與判斷句、v4.3 市場共識逐筆與一頁摘要的差距、隱含倍數與句子逐字比對、v4.4 季度層 6 季數字／差距／差異原因／驗證點句與年度差異原因；讀取目前目錄的 `xl17_*.json`） |
| scripts/recalc.py | LibreOffice headless 重算：`python3 scripts/recalc.py 檔案.xlsx [逾時秒數]`＝開啟、全部重算、原地存檔，輸出 JSON 並回報公式錯誤數（#REF!、#DIV/0!、#VALUE!、#NAME?、#N/A）；有錯誤或失敗時代碼 1 |
| scripts/verify.sh | 一次跑完建置與核對（見下節） |
| scripts/first_screen.py | v4.5：擷取「第一屏」文字——HTML 各頁面／分頁在 1440×900 視窗、捲動到頂端時可見的文字；Excel 各工作表前 40 個可見列（A–J 欄，重算後的值）。`python3 scripts/first_screen.py 檔案.html｜檔案.xlsx [輸出.json]` |
| scripts/attrib.py | 目標價變動拆解：輸入前一版與各步驟（.xlsx、company.json 或 repo 目錄），以 Excel 求值，輸出三情境兩條腿與加權目標價的 (a)(b)(c)(d)，檢查四項相加＝總變動 |
| scripts/test_attrib.py | 拆解工具測試：月數推算、同版對同版全為 0、滾動一季且 WACC 改為 12% 時 (a) 係數＝(1.12)^(3/12)。verify.sh 步驟 8c |
| scripts/test_rolling.py | v4.5：期間滾動測試（暫存副本）。B：日曆推算 6 種情況（12 月財年 Q1–Q4 已申報、5 月財年、無已申報季度）× 2 種目標價時點；C：只滾日曆、未更新 `asOf` 時建置必須失敗並逐項列出；A：把副本滾動到下一個已申報季度（只改日曆與標籤、不改數字），建 HTML 與 Excel，第一屏不得出現舊日曆特有的字樣（評價日、年初至今／首期標籤、已申報季度）。例外：版本紀錄與來源頁、整段落在 company.json／共識檔資料字串內的字樣。verify.sh 步驟 8b |
| scripts/check_offline.py | 離線開啟檢查（已決定事項 11）：`python3 scripts/check_offline.py 檔案.html 同版.xlsx`。HTML 單獨複製到空資料夾，Playwright 阻斷網路、以 file:// 開啟；(a) 除 HTML 本身外的請求數＝0（網路與旁邊的本機檔案都算）、(b) console 無錯誤與例外、(c) 5 個分頁 10 項關鍵數字（依 A 欄列名稱取自同版 Excel）出現在畫面且無 NaN／undefined／Infinity；v4.5 另比對第一屏（視窗內）的資料更新日、模型期間、現價日與評價日（`CHECKS_FIRST`；列名稱中的期間佔位符以任意文字比對，新舊版 Excel 皆可）。verify.sh 步驟 5c 對新建 HTML 與 dist/ 成品各跑一次 |
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

## company.json 欄位說明（換公司填表指引）
換成 Nebius、Oracle、OpenAI 等公司時，照這一節逐欄填寫 `company.json`；HTML 與 Excel 都從這個檔讀資料，改完執行 `scripts/verify.sh`。

**填表慣例**
- 金額單位是**十億美元（US$bn）**，例如 4.653 代表 46.53 億美元；另有標示的例外：每股（US$）、每 MW 建置成本（百萬美元／MW，US$m/MW）、股數（十億股，bn）。
- 「比例」寫成小數（0.25＝25%）；標示「%」的欄位寫成百分點（25＝25%）。兩種寫法沿用既有程式，不可混用。
- 「清單」依模型期順序填：FY27 第 2–4 季、FY28、FY29、FY30、FY31，共 5 格（除非另有說明）。
- 文字中的來源標記沿用 [Verified]（已公開可查）、[Interested-party]（利害關係人說法）、[Derived]（由其他數字換算）、[Assumed]（判斷值）。
- 「換公司」欄：**必改**＝公司特有的資料；**檢查**＝判斷值，要依新公司重新評估；**可沿用**＝口徑或方法，通常不必改。
- 下表的「目前數值」是 Oracle v0.2 的值（版本號讀 `vlog.py`、期間讀 `calendar_q.py`，由本檔自動帶入）；過長的文字只顯示開頭。表格由 `scripts/fields_doc.py` 產生，新增欄位時先在該檔補說明，再重新產生。

### `meta`：基本資料

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `meta.company` | 公司名稱 | 文字 | Oracle | 必改 |
| `meta.ticker` | 股票代號 | 文字 | ORCL | 必改 |
| `meta.updateDate` | 資料更新日 | 日期 | 2026-10-08 | 必改 |
| `meta.priceDate` | 股價日期（現價的收盤日；畫面與 Excel 的現價日期都讀這格） | 日期 | 2026-10-06 | 必改 |
| `meta.consensusFile` | 市場共識資料檔路徑（v4.3；只讀，由使用者查證後提供；建置時併入 HTML、Excel 讀同一檔） | 路徑 | data/consensus_orcl_202610… | 必改 |
| `meta.sourceOrderNote` | 資料來源的先後與衝突時的取捨原則（畫面說明文字） | 文字 | 時序先 FY27Q1 財報新聞稿與法說（2026-0… | 必改 |

### `calendar`：期間與日期（v4.5）

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `calendar.fiscalYearEndMonth` | 財年結束月份（v4.5；CoreWeave 12、Oracle 5） | 月 | 5 | 必改 |
| `calendar.latestQuarterFiled` | 最新已申報（10-Q／10-K）的財季，格式 FYyyQn；驅動年度首期滾動與評價日（v4.5） | 文字 | FY27Q1 | 必改 |
| `calendar.latestQuarterReported` | 最新已公布（財報新聞稿）的財季；驅動季度層，年度首期不受影響（v4.5） | 文字 | FY27Q1 | 必改 |
| `calendar.firstModelFY` | 沒有已申報季度時（例如未上市公司）的首個模型財年，例如 FY26；有已申報季度時填 null（v4.5） | 文字 | None | 檢查 |
| `calendar.modelYears` | 首期之後的完整財年數（固定期數滾動；Excel 欄位固定 5 期，所以為 4）（v4.5） | 年 | 4 | 可沿用 |
| `calendar.targetHorizon` | 目標價時點：firstFullYearEnd＝首期之後第一個完整財年末；valuationPlus12m＝評價日＋12 個月（v4.5） | 代碼 | valuationPlus12m | 可沿用 |

### `asOf`：滾動檢查（首期一次性金額與期初餘額的所屬季度）

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `asOf` | 滾動檢查：首期一次性金額與期初餘額所屬的已申報季度（鍵＝欄位路徑，清單定義在 calendar_q.py → ROLL_FIELDS）。每季 10-Q 後逐項更新數值，並把季度改為 calendar.latestQuarterFiled；缺漏或季度不符即建置失敗 | 物件（季度） | 物件（_note、defaults.capexFloorFY0、leases.onBalanceCash[0]、leases.operatingPayments[0]、leases.financePayments[0]、debt.amortization[0]、defaults.jvCommit[0]、scenarios.capexTemplate.div[0]、defaults.intCal、defaults.services[0]、defaults.atm、leases.uncommenced、rpo.bucketWeights[0]、defaults.cash、debt.instruments、debt.convertible、valuation.netDebt、valuation.shares、defaults.ppeOpen、defaults.billableOpen、defaults.rpoOpen、defaults.rpoPendingAdd、defaults.eqCapShares、defaults.mwYearEnd、defaults.prepay.openBalance、debt.convertibles、defaults.otherEbitda[0]、valuation.holdings、valuation.debtLike、defaults.legacyBiz、defaults.dividend.preferred[0]） | 必改 |

### `ytdActual`：年初至今實際數（10-Q；v4.5 前為 actual1H）

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `ytdActual.throughQuarter` | 年初至今實際數的截止財季，須等於 calendar.latestQuarterFiled（v4.5，原 actual1H） | 文字 | FY27Q1 | 必改 |
| `ytdActual.months` | 年初至今的月數，須等於日曆推算值（v4.5） | 月 | 3 | 必改 |
| `ytdActual.label` | 「年初至今實際數」的標題 | 文字 | 1Q27 實際（10-Q） | 必改 |
| `ytdActual.jvSplit` | JV 出資與策略投資的拆分（v4.5；說明文字與「JV 已付」讀此） | 物件（US$bn） | 物件（jv、strategic） | 必改 |
| `ytdActual.notes` | 各欄位的逐列說明（來源、口徑、拆分；Excel「輸入與假設」G 區；v4.5 起隨資料一起更新） | 物件（文字） | 物件（cash1231、cfo、cashCapex、capex、jv、borrow、debtRepaid、cappedCall、equity、interest、leasePaid、revenue、opInc、ni、prepay、da、sbc、eps、ngEps、dividends） | 必改 |
| `ytdActual.cash1231` | 上一年底現金 | US$bn | 31.289 | 必改 |
| `ytdActual.revenue` | 上半年營收 | US$bn | 19.345 | 必改 |
| `ytdActual.capex` | 上半年資本支出（認列口徑，含設備商融資） | US$bn | 28.499 | 必改 |
| `ytdActual.cashCapex` | 上半年現金購置固定資產 | US$bn | 28.499 | 必改 |
| `ytdActual.interest` | 上半年利息費用 | US$bn | 1.428 | 必改 |
| `ytdActual.leasePaid` | 上半年租賃現金支付 | US$bn | 1.136 | 必改 |
| `ytdActual.debtRepaid` | 上半年還款 | US$bn | 4.202 | 必改 |
| `ytdActual.borrow` | 上半年借款 | US$bn | 0 | 必改 |
| `ytdActual.equity` | 上半年股權募資 | US$bn | 19.909 | 必改 |
| `ytdActual.cappedCall` | 上半年 capped call（可轉債配套避險）支出 | US$bn | 0 | 必改 |
| `ytdActual.jv` | 上半年合資（JV）與策略投資出資 | US$bn | 0 | 必改 |
| `ytdActual.cfo` | 上半年營運現金流 | US$bn | 23.103 | 必改 |
| `ytdActual.prepay` | 上半年客戶預付（遞延收入淨流入） | US$bn | 11.363 | 必改 |
| `ytdActual.da` | 上半年折舊攤銷 | US$bn | 3.358 | 必改 |
| `ytdActual.sbc` | 上半年股份基礎薪酬 | US$bn | 1.127 | 必改 |
| `ytdActual.opInc` | 上半年 GAAP 營業損益 | US$bn | 6.728 | 必改 |
| `ytdActual.ni` | 上半年淨損益 | US$bn | 4.76 | 必改 |
| `ytdActual.eps` | 上半年 GAAP 每股盈餘 | US$ | 1.56 | 必改 |
| `ytdActual.ngEps` | 上半年每股盈餘（加回股份基礎薪酬） | US$ | 1.96 | 必改 |
| `ytdActual.dividends` | 年初至今股利支付（普通股＋特別股；融資活動；v0.1b） | US$bn | 1.565 | 必改 |
| `ytdActual.adjEbitda` | 上半年調整後 EBITDA（v4.3；只用於與市場共識比較 FY26，不進模型損益與評價） | US$bn | 11.307 | 必改 |
| `ytdActual.adjEbitdaMeta` | 上述數字的 Q1／Q2 拆分、來源、標記與備註（建置時檢查 Q1＋Q2＝合計） | 物件 | 物件（q1、tag、sources、crossCheck、note、usage） | 必改 |

### `historicalPL`：歷年損益

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `historicalPL` | 歷年損益，一年一列（損益頁的歷史欄）。每列欄位：year 年度、revenue 營收、opInc GAAP 營業損益、ni 淨損益（以上 US$bn）、eps GAAP 每股盈餘、ngEps 加回股份基礎薪酬的每股盈餘（US$）、shares 加權流通股數（bn）、tag 來源標記 | 清單 | 4 筆 | 必改 |

### `rpo`：已簽約未認列營收（RPO）

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `rpo.scheduledShare` | 剩餘履約義務（RPO，已簽約未認列的營收）預計在模型期內認列的比例 | 比例 | 0.7975 | 必改 |
| `rpo.bucketLabels` | RPO 季報桶的名稱（與 rpo.split 對應；v0.1b） | 文字清單 | ≤12 個月、13–36 個月、37–60 個月、其後 | 必改 |
| `rpo.split` | RPO 季報桶的比例（季報揭露；v0.1b） | 比例清單 | 0.13、0.37、0.34、0.16 | 必改 |
| `rpo.within36m` | RPO 預計 36 個月內認列的比例（季報／法說；只作對照列；v0.1b） | 比例 | 0.5 | 必改 |
| `rpo.bucketWeights` | RPO 在五期（首期模型部分＋4 個完整財年；目前為 FY27 第 2–4 季、FY28、FY29、FY30、FY31）各期認列的比例，合計等於上一欄 | 比例清單 | 0.0975、0.17125、0.185、0.17375、0.17 | 必改 |

### `leases`：租約

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `leases.onBalanceCash` | 已入帳租約在五期（首期模型部分＋4 個完整財年；目前為 FY27 第 2–4 季、FY28、FY29、FY30、FY31）各期的現金租金 | US$bn 清單 | 3.839、4.987、5.02、5.056、5.06 | 必改 |
| `leases.afterFY30` | 已入帳租約在模型期之後還要付的租金合計 | US$bn | 39.283 | 必改 |
| `leases.facts.onBal` | 已入帳租約未折現付款合計 | US$bn | 63.245 | 必改 |
| `leases.facts.notCommenced` | 已簽約但尚未起租的租約（表外） | US$bn | 288 | 必改 |
| `leases.liability.discRate` | 租賃負債折現率（10-K 加權平均；v0.2） | 比例 | 0.057 | 必改 |
| `leases.liability.tailYears` | 在帳到期表模型期後尾端的平均分攤年數（租賃負債用；v0.2） | 年 | 12 | 檢查 |
| `leases.liability.note` | 租賃負債口徑說明（v0.2） | 文字 | v0.2：租賃負債（每期末）＝剩餘租金現值。折現率＝… | 必改 |
| `leases.uncommenced.startQ` | 未起租租賃自評價日後第幾季開始起租（0＝首期第一季） | 季 | 0 | 必改 |
| `leases.uncommenced.quarters` | 未起租租賃平均分攤起租的季數 | 季 | 11 | 必改 |
| `leases.uncommenced.termYears` | 每筆未起租租賃的租期（直線付租） | 年 | 17 | 檢查 |
| `leases.uncommenced.termSens` | 租期敏感度（報告用） | 年 清單 | 15、19 | 檢查 |
| `leases.uncommenced.note` | 起租排程的來源與假設說明 | 文字 | 10-Q FY27Q1 附註 6：未起租租賃 288… | 必改 |
| `leases.uncommenced.delayLink` | 未起租租約起租隨建設延誤後移的比例（0–1；其餘照原時程；v0.2） | 比例 | 0.5 | 檢查 |
| `leases.uncommenced.delayLinkNote` | delayLink 的依據說明（v0.2） | 文字 | v0.2：未起租 288B 中有此比例的起租時點隨建… | 必改 |
| `leases.facts.singleCap` | 單一大型站點的租金上限（10-Q 揭露） | US$bn | 0 | 必改 |
| `leases.facts.share` | 第三方租賃占機房取得的比例（用於租金基準檢驗） | 比例 | 1 | 檢查 |
| `leases.operatingPayments` | 營業租賃到期表：五期各期，最後一格為之後合計（Excel 租賃頁） | US$bn 清單 | 3.219、4.135、4.097、4.109、4.089、28.401 | 必改 |
| `leases.financePayments` | 融資租賃到期表：同上格式 | US$bn 清單 | 0.62、0.852、0.923、0.947、0.971、10.882 | 必改 |

### `debt`：既有債務

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `debt.amortization` | 既有債務在五期（首期模型部分＋4 個完整財年；目前為 FY27 第 2–4 季、FY28、FY29、FY30、FY31）各期的排程還本 | US$bn 清單 | 3.007、10.147、5.5、7.25、9.75 | 必改 |
| `debt.amortAfterFY30` | 模型期之後的還本合計 | US$bn | 90.25 | 必改 |
| `debt.instruments` | 既有債務逐筆明細，一筆一列：[名稱, 追索／非追索, 到期, 有效利率（比例）, 本金（US$bn）, 備註]。加總本金須等於 10-Q 本金合計 | 清單 | 57 筆 | 必改 |
| `debt.convertible.principal` | 期後新發行可轉債本金（不在五期還本表內，只計利息） | US$bn | 0 | 必改 |
| `debt.convertible.coupon` | 該可轉債票面利率 | 比例 | 0 | 必改 |
| `debt.convertibles` | 可轉債逐檔，一檔一列：[名稱, 原始本金（US$bn）, 票息（比例）, 到期 YYYY-MM, 到期累積倍數, 轉換價（US$）, 備註, 強制轉換（選填；true＝一律轉股、不計利息與還本、不列債務本金；Oracle v0.1b）]。有效轉換價＝轉換價 × 累積倍數；低於判斷價視為轉股（若轉換法），否則以到期累積本金計債務並付現金票息（v0.1b） | 清單 | 1 筆 | 必改 |
| `debt.convertibleBridge.exchangedAccreted` | 評價日後以股換債註銷的舊債到期本金（季報本金與逐檔清單的調節項） | US$bn | 0 | 必改 |
| `debt.convertibleBridge.newIssuesAccreted` | 評價日後新發可轉債的到期本金（季報本金與逐檔清單的調節項） | US$bn | 0 | 必改 |
| `debt.convertibleBridge.note` | 調節說明與來源 | 文字 | 無評價日後新發可轉債、無以股換債；債務明細＝55 檔… | 必改 |

### `latestQuarter`：最新一季財報數字（10-Q）

換公司：全部必改（公司特有資料）。「程式使用」為否的欄位只是存檔備查，畫面與 Excel 的說明文字目前另外寫在程式裡。

| 欄位 | 意義 | 單位 | 目前數值 | 程式使用 |
|---|---|---|---|---|
| `filed` | 申報日 | 日期 | 2026-09-11 | 否 |
| `periodEnd` | 季末日 | 日期 | 2026-08-31 | 否 |
| `revenue` | 當季營收 | US$bn | 19.345 | 是 |
| `yoy` | 當季營收年增率 | 比例 | 0.2961 | 是 |
| `h1Revenue` | 上半年營收 | US$bn | 19.345 | 是 |
| `costRev` | 當季營收成本 | US$bn | None | 否 |
| `techInfra` | 當季技術與基礎設施費用 | US$bn | None | 否 |
| `opInc` | 當季 GAAP 營業損益 | US$bn | 6.728 | 否 |
| `interest` | 當季利息費用 | US$bn | 1.428 | 否 |
| `h1Interest` | 上半年利息費用 | US$bn | 1.428 | 否 |
| `ni` | 當季淨損益 | US$bn | 4.76 | 否 |
| `h1Ni` | 上半年淨損益 | US$bn | 4.76 | 否 |
| `epsDiluted` | 當季稀釋每股盈餘 | US$ | 1.56 | 否 |
| `sbc` | 當季股份基礎薪酬 | US$bn | 1.127 | 否 |
| `da` | 當季折舊攤銷 | US$bn | 3.358 | 是 |
| `adjEbitda` | 當季調整後 EBITDA | US$bn | 11.307 | 否 |
| `adjOpInc` | 當季調整後營業利益 | US$bn | 8.151 | 否 |
| `rpo` | 季末 RPO | US$bn | 664 | 是 |
| `backlog` | 季末 backlog（含其他） | US$bn | None | 否 |
| `rpo24m` | RPO 於 24 個月內認列比例 | 比例 | None | 否 |
| `rpo25to48` | RPO 於 25–48 個月認列比例 | 比例 | None | 否 |
| `rpo49to78` | RPO 於 49–78 個月認列比例 | 比例 | None | 否 |
| `cash` | 季末現金 | US$bn | 36.369 | 是 |
| `restricted` | 受限現金 | US$bn | 2.6 | 是 |
| `marketable` | 有價證券 | US$bn | 0.708 | 是 |
| `availability` | 未動用信用額度 | US$bn | 10 | 否 |
| `debtPrincipal` | 債務本金合計 | US$bn | 125.904 | 是 |
| `recourseNet` | 追索債務淨額 | US$bn | None | 否 |
| `nonRecourseNet` | 非追索債務淨額 | US$bn | None | 否 |
| `ddtlOut` | DDTL 未償餘額 | US$bn | 0 | 否 |
| `notesOut` | 票據與可轉債未償餘額 | US$bn | 120.5 | 否 |
| `sharesA` | A 股流通股數 | bn 股 | None | 否 |
| `sharesB` | B 股流通股數 | bn 股 | None | 否 |
| `sharesOut` | 流通股數合計 | bn 股 | 3.02374 | 是 |
| `basicWaso` | 加權平均流通股數 | bn 股 | 2.966 | 否 |
| `capexQ2` | 當季資本支出 | US$bn | 28.499 | 否 |
| `capexH1` | 上半年資本支出 | US$bn | 28.499 | 是 |
| `cashCapexH1` | 上半年現金購置固定資產 | US$bn | 28.499 | 否 |
| `cfoH1` | 上半年營運現金流 | US$bn | 23.103 | 否 |
| `cashInterestH1` | 上半年現金利息 | US$bn | None | 否 |
| `capInterestH1` | 上半年資本化利息 | US$bn | None | 否 |
| `deferredTotal` | 遞延收入合計 | US$bn | 30.789 | 是 |
| `deferredIn` | 上半年遞延收入淨流入 | US$bn | 15.394 | 是 |
| `ppe` | 固定資產毛額 | US$bn | 127.845 | 是 |
| `cip` | 在建工程 | US$bn | 48.546 | 是 |
| `rouOp` | 營業租賃使用權資產 | US$bn | None | 否 |
| `opLeaseLiab` | 營業租賃負債 | US$bn | 34.621 | 是 |
| `finLeaseLiab` | 融資租賃負債 | US$bn | 9.185 | 是 |
| `onBalanceUndiscounted` | 已入帳租約未折現付款 | US$bn | 63.245 | 是 |
| `offBalanceLease` | 未起租租約（表外） | US$bn | 288 | 是 |
| `singleSiteCap` | 單一站點租金上限 | US$bn | 0 | 是 |
| `singleSiteMw` | 該站點未交付 MW | MW | None | 否 |
| `constructionCostMw` | 按造價計租的未交付 MW | MW | None | 否 |
| `equipCommitLo` | 設備採購承諾下緣 | US$bn | None | 否 |
| `equipCommitHi` | 設備採購承諾上緣 | US$bn | None | 否 |
| `jvCommit` | JV 承諾出資上限 | US$bn | 0 | 否 |
| `jvPaidH1` | 上半年已付 JV 出資 | US$bn | 0 | 否 |
| `vieExposure` | 可變利益實體（VIE）最大曝險 | US$bn | None | 否 |
| `rouObtainedH1` | 上半年新取得使用權資產 | US$bn | None | 否 |
| `leaseCashH1` | 上半年租賃現金支付 | US$bn | 1.136 | 否 |
| `custA` | 最大客戶營收占比 | 比例 | 0.5 | 否 |
| `custB` | 第二大客戶營收占比 | 比例 | 0 | 否 |
| `custC` | 第三大客戶營收占比 | 比例 | 0 | 否 |
| `debtIssuedH1` | 上半年借款 | US$bn | 0 | 否 |
| `debtRepaidH1` | 上半年還款 | US$bn | 4.202 | 否 |
| `equityH1` | 上半年股權募資 | US$bn | 19.909 | 否 |

### `callFacts`：法說會與期後事項

換公司：全部必改（公司特有資料）。「程式使用」為否的欄位只是存檔備查，畫面與 Excel 的說明文字目前另外寫在程式裡。

| 欄位 | 意義 | 單位 | 目前數值 | 程式使用 |
|---|---|---|---|---|
| `activeGw` | 主動電力 | GW | None | 否 |
| `activeAddQ2` | 當季新增主動電力 | MW | 0.85 | 否 |
| `juneAddMw` | 季末月單月新增 | MW | None | 否 |
| `contractedGw` | 簽約電力（法說日） | GW | 4.5 | 否 |
| `contractedGwQ2` | 簽約電力（季末） | GW | None | 否 |
| `poweredLandGw` | 已取得電力的土地／選擇權／意向書 | GW | None | 否 |
| `yeActiveGw` | 年底主動電力指引 | GW | None | 否 |
| `gw2030` | 2030 年電力目標 | GW | None | 否 |
| `dataCenters` | 資料中心數 | 座 | None | 否 |
| `capexLo` | 全年資本支出指引下緣 | US$bn | 90 | 是 |
| `capexHi` | 全年資本支出指引上緣 | US$bn | 95 | 是 |
| `nextQCapexLo` | 下一季資本支出指引下緣 | US$bn | None | 否 |
| `nextQCapexHi` | 下一季資本支出指引上緣 | US$bn | None | 否 |
| `nextQIntLo` | 下一季利息指引下緣 | US$bn | None | 是 |
| `nextQIntHi` | 下一季利息指引上緣 | US$bn | None | 是 |
| `revLo` | 全年營收指引下緣 | US$bn | 90 | 是 |
| `revHi` | 全年營收指引上緣 | US$bn | None | 是 |
| `nextQRevLo` | 下一季營收指引下緣 | US$bn | None | 是 |
| `nextQRevHi` | 下一季營收指引上緣 | US$bn | None | 是 |
| `adjOpLo` | 全年調整後營業利益指引下緣 | US$bn | None | 否 |
| `adjOpHi` | 全年調整後營業利益指引上緣 | US$bn | None | 否 |
| `nextQAdjOpLo` | 下一季調整後營業利益指引下緣 | US$bn | None | 否 |
| `nextQAdjOpHi` | 下一季調整後營業利益指引上緣 | US$bn | None | 否 |
| `arrLo` | 期末年化經常性收入（ARR）指引下緣 | US$bn | None | 否 |
| `arrHi` | 期末 ARR 指引上緣 | US$bn | None | 否 |
| `postQNewCommit` | 季末後新增承諾 | US$bn | None | 否 |
| `priceUp` | 新約漲價幅度 | 比例 | 0.2 | 否 |
| `marginStep` | 新約貢獻率提升 | 文字 | None | 否 |
| `backlogStarted` | backlog 中已開始交付的比例（法說「>50%」） | 比例 | None | 否 |
| `inferenceArr` | 推論服務 ARR（年底） | US$bn | None | 否 |
| `otherArr` | 其他服務 ARR | US$bn | None | 否 |
| `convert` | 期後可轉債發行額 | US$bn | 0 | 否 |
| `convertNet` | 期後可轉債淨額 | US$bn | 0 | 否 |
| `convertCoupon` | 期後可轉債票面利率 | 比例 | None | 否 |
| `convertPx` | 轉換價 | US$ | None | 否 |
| `atmShares` | 股權分銷（ATM）上限 | m 股 | None | 否 |
| `price0917` | 9/17 收盤價 | US$ | None | 否 |
| `tier34Active` | 主動電力位於 Tier 3/4 市場比例 | 比例 | None | 否 |
| `tier34Contracted` | 簽約電力位於 Tier 3/4 市場比例 | 比例 | None | 否 |
| `cdsMid` | CDS 中價 | bps | None | 否 |
| `cdsBidAsk` | CDS 買／賣價 | 文字 | None | 否 |
| `cdsRange` | CDS 近期區間 | 文字 | None | 否 |
| `cdsPeak` | CDS 峰值 | bps | None | 否 |
| `cdsPeakDate` | CDS 峰值日期 | 日期 | None | 否 |
| `cdsLow` | CDS 低點 | bps | None | 否 |
| `cdsLowDate` | CDS 低點日期 | 日期 | None | 否 |
| `availability` | 未動用信用額度 | US$bn | 10 | 否 |
| `price0918` | 9/18 收盤價 | US$ | None | 否 |
| `priceLast` | 最新收盤價 | US$ | 144.77 | 是 |
| `priceDate` | 最新收盤價日期 | 日期 | 2026-10-06 | 是 |
| `postQShortDated` | 期後公告摘要 | 文字 | 2026-09-24 Oracle 對 Projec… | 是 |
| `ttmRev` | 近十二個月營收 | US$bn | 71.777 | 是 |
| `ttmOpInc` | 近十二個月 GAAP 營業損益 | US$bn | 23.056 | 是 |
| `ttmDa` | 近十二個月折舊攤銷 | US$bn | 10.881 | 是 |
| `ttmOpLease` | 近十二個月營業租賃成本 | US$bn | 34.621 | 否 |
| `mktCapLast` | 本公司市值（Comps 用；與同業同一收盤日 peers.priceDate，v4.3 起不隨現價更新） | US$bn | 437.75 | 是 |
| `nextEarn` | 下次財報時間 | 文字 | 2026-12（Q2 FY27 財報） | 是 |

### `scenarios`：三個擴張情境

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `scenarios.labels.low` | 保守情境名稱（空格前的文字會當作情境簡稱） | 文字 | 保守 OpenAI 相關 60% | 必改 |
| `scenarios.labels.base` | 基準情境名稱 | 文字 | 基準 已簽約站點依 Q1 速度 | 必改 |
| `scenarios.labels.high` | 積極情境名稱 | 文字 | 積極 依公司時程 | 必改 |
| `scenarios.mwPath.connectedStart` | 首期期末已連網 MW（三情境共用；v0.1b） | MW | 4668 | 必改 |
| `scenarios.mwPath.contracted.low` | 保守情境：五期（首期模型部分＋4 個完整財年；目前為 FY27 第 2–4 季、FY28、FY29、FY30、FY31）各期末的合約 MW 上限（已連網不得超過；v0.1b） | MW 清單 | 6284、6284、6284、6284、6284 | 必改 |
| `scenarios.mwPath.contracted.base` | 基準情境：同上 | MW 清單 | 8778、8778、8778、8778、8778 | 必改 |
| `scenarios.mwPath.contracted.high` | 積極情境：同上 | MW 清單 | 8778、8778、8778、8778、8778 | 必改 |
| `scenarios.mwPath.pace.low` | 保守情境：併網速度（每年新增已連網 MW；已連網＝MIN(合約上限, 前期＋速度×期間長度)；v0.1b） | MW／年 | 1000 | 檢查 |
| `scenarios.mwPath.pace.base` | 基準情境：同上 | MW／年 | 2833 | 檢查 |
| `scenarios.mwPath.pace.high` | 積極情境：同上 | MW／年 | 3333 | 檢查 |
| `scenarios.mwPath.note` | 已連網 MW 路徑的說明（來源與口徑） | 文字 | 已連網 MW-IT（期末）＝MIN(合約 MW-IT… | 必改 |
| `scenarios.revMW.low` | 保守情境：每 MW 年收入，各期（Tokenomics 正向推導；不得用公司 ACV；v0.1b） | US$bn/MW 清單 | 0.01162、0.01162、0.01162、0.01162、0.01162 | 必改 |
| `scenarios.revMW.base` | 基準情境：同上 | US$bn/MW 清單 | 0.0174、0.0174、0.0174、0.0174、0.0174 | 必改 |
| `scenarios.revMW.high` | 積極情境：同上 | US$bn/MW 清單 | 0.0242、0.0242、0.0242、0.0242、0.0242 | 必改 |
| `scenarios.descriptions` | 三情境的一句說明（Excel A 區；v0.1b） | 物件（文字） | 物件（low、base、high） | 必改 |
| `scenarios.billableRatio.mode` | 可計費 MW 的算法：ratio＝已連網 × 比例；converge＝期初以最新季實際營收年化 ÷ 每 MW 年收入校準，之後向已連網收斂（v0.1b） | 代碼 | converge | 檢查 |
| `scenarios.billableRatio.openAnnualRevenue` | converge 模式的校準營收：最新季實際營收（OCI 等 MW 驅動部分）× 4（v0.1b） | US$bn | 29.552 | 必改 |
| `scenarios.billableRatio.note` | 可計費 MW 校準與收斂比例的說明（v0.1b） | 文字 | 期初可計費 MW＝Q1 FY27 實際 OCI 營收… | 必改 |
| `scenarios.billableRatio.ratio` | 各期比例：ratio 模式＝在役 ÷ 已連網；converge 模式＝期初校準值向已連網收斂的比例（三情境共用；v0.1b） | 比例清單 | 0.5、0.8、0.85、0.9、0.9 | 檢查 |
| `scenarios.mw31.low` | 保守情境：模型期後一年（FY31）新增的 MW，用於 FY30 的預建支出 | MW | 0 | 檢查 |
| `scenarios.mw31.base` | 基準情境：同上 | MW | 0 | 檢查 |
| `scenarios.mw31.high` | 積極情境：同上 | MW | 0 | 檢查 |
| `scenarios.convCap.low` | 保守情境：融資瀑布可轉債步驟每年新發行上限（0＝不新發；v0.1b） | US$bn／年 | 0 | 檢查 |
| `scenarios.convCap.base` | 基準情境：同上 | US$bn／年 | 0 | 檢查 |
| `scenarios.convCap.high` | 積極情境：同上 | US$bn／年 | 0 | 檢查 |
| `scenarios.delayMonths.low` | 保守情境：建設延誤月數——計費 MW＝原可計費路徑往後平移此月數（期間長度線性內插）；GPU 資本支出與客戶出資照原時程，折舊自投入使用起算（v0.2） | 月 | 6 | 檢查 |
| `scenarios.delayMonths.base` | 基準情境：同上 | 月 | 3 | 檢查 |
| `scenarios.delayMonths.high` | 積極情境：同上 | 月 | 0 | 檢查 |
| `scenarios.delayMonths.note` | 建設延誤月數的依據與說明（v0.2） | 文字 | 建設延誤月數（v0.2）：計費／營收的 MW＝原可計… | 必改 |
| `scenarios.capexTemplate.costMW` | 每 MW 建置成本（GPU＋網路＋機房內裝），各期 | US$m/MW 清單 | 37.45、37.59、37.59、37.59、37.59 | 檢查 |
| `scenarios.capexTemplate.div` | JV 後續增資與策略投資，各期 | US$bn 清單 | 0、0、0、0、0 | 檢查 |

### `legacy`：舊版對照值

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `legacy.capexV14` | 舊版（v1.4）手動資本支出，只供畫面對照 | US$bn 清單 | 0、0、0、0、0 | 可沿用 |
| `legacy.interestV14` | 舊版（v1.4）手動利息，只供畫面對照 | US$bn 清單 | 0、0、0、0、0 | 可沿用 |

### `defaults`：預設假設（畫面上可調的輸入）

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `defaults.scenario` | 開啟時的預設情境（low／base／high） | 文字 | base | 檢查 |
| `defaults.revenueDriver` | 營收驅動：mw＝平均在役 MW × 每 MW 年收入 × 利用率（新產能簽約率固定 100%，RPO 只作對照）；rpo＝CRWV 模板的 RPO 排程＋新簽約（v0.1b） | 代碼 | mw | 檢查 |
| `defaults.lambda` | 提前支出比例：次年才上線的 MW，其建置支出落在前一年的比例 | 比例 | 0.35 | 檢查 |
| `defaults.mwYearEnd` | 各年底主動電力，以年份為鍵（例如 "2025": 850）：首期期初取首期前一財年末；GPU 汰換批次＝各年新增 MW（5a；列入滾動檢查） | MW 物件 | 物件（2026） | 必改 |
| `defaults.mw31` | 模型期後一年新增 MW 的預設值（情境切換時改用 scenarios.mw31） | MW | 0 | 檢查 |
| `defaults.capexFloorFY0` | 首期所屬財年的全年資本支出下限（已下單的承諾，取公司指引下緣；首期＝下限 − 年初至今實際認列） | US$bn | 90 | 必改 |
| `defaults.gpuLife` | GPU 經濟壽命（決定汰換時點與折舊） | 年 | 6 | 檢查 |
| `defaults.refreshSteady` | 穩態汰換：true＝已連網 MW 不再增加的期間（觸頂後）及終值年，汰換 CapEx＝平均已連網 MW × 每 MW 建置成本 ÷ GPU 壽命 × 期間長度；false＝只有批次汰換（Oracle v0.1c） | 是／否 | 是 | 檢查 |
| `defaults.refreshNote` | 穩態汰換的說明文字 | 文字 | v0.1c：已連網 MW 不再增加的期間（觸頂後）及… | 可沿用 |
| `defaults.revScale` | 每 MW 年收入整體倍數（反向 DCF 與壓力測試用，1＝不調整） | 倍 | 1 | 可沿用 |
| `defaults.capexScale` | 每 MW 建置成本整體倍數（1＝不調整） | 倍 | 1 | 可沿用 |
| `defaults.ebStart` | 第一期 EBITDA 率 | 比例 | 0.652 | 必改 |
| `defaults.ebSteady` | 最後一期（穩態）EBITDA 率；中間各期線性內插 | 比例 | 0.47 | 檢查 |
| `defaults.ebitdaBasis` | EBITDA 口徑：ebitdar＝EBITDA＝EBITDAR 率 × 營收 − 租金（租金為固定成本，EBITDAR 率三情境共用）；其他值＝三情境共用 EBITDA 率（模板）（Oracle v0.1c） | 代碼 | ebitdar | 檢查 |
| `defaults.ebitdarAdj` | EBITDAR 率校準：基準情境 [首期, 末期] 租金 ÷ OCI 營收；EBITDAR 率＝ebStart／ebSteady＋此值（scripts/calib_ebitdar.js --write 產生，verify.sh 檢查） | 比例清單 | 0.193075、0.156683 | 必改 |
| `defaults.ebitdarNote` | EBITDAR 口徑的說明文字 | 文字 | v0.1c：OCI EBITDA＝EBITDAR 率… | 可沿用 |
| `defaults.services` | 非算力服務營收（軟體、儲存等），各期 | US$bn 清單 | 0、0、0、0、0 | 檢查 |
| `defaults.otherEbitda` | 其他事業 EBITDA（負值＝燒錢），各期；同時進入損益 EBITDA 與營運來源（v0.1b） | US$bn 清單 | 0、0、0、0、0 | 必改 |
| `defaults.legacyBiz.lines` | 傳統事業各線，一線一列：key、label、fyBase 上一財年實際營收（US$bn）、ytd 年初至今實際營收、g0 起始年增率、gLT 長期年增率（自首期線性收斂到末期）（Oracle v0.1b） | 清單 | 4 筆 | 必改 |
| `defaults.legacyBiz.ebitdaMargin` | 傳統事業 EBITDA 率，各期（Oracle v0.1b） | 比例清單 | 0.5426、0.5426、0.5426、0.5426、0.5426 | 檢查 |
| `defaults.legacyBiz.note` | 傳統事業輸入的來源與推導說明 | 文字 | 傳統事業（OCI 以外四線）：fyBase＝FY20… | 必改 |
| `defaults.cashTaxRate` | 類現金流量的現金稅率：稅 ＝ 稅率 × MAX(0, 損益 EBITDA − 車隊 D&A − 存量利息)（v0.1b；虧損或 NOL 公司填 0） | 比例 | 0.151 | 檢查 |
| `defaults.delayPenalty` | 延誤罰則／服務抵減：延誤期間應計費而未計費營收的比例，列為營業費用（預設 0＝未揭露；v0.2） | 比例 | 0 | 檢查 |
| `defaults.delayPenaltyNote` | delayPenalty 的依據說明（v0.2） | 文字 | v0.2：延誤罰則或服務抵減＝延誤期間「應計費而未計… | 必改 |
| `defaults.debtCapBasis` | 瀑布新債的上限基準：leaseAdj＝(總債務＋租賃負債) ≤ 倍數 ×(EBITDA＋租金)（年化；S&P 口徑近似的投資級上限，v0.2）；ebitda＝總債務 ≤ 倍數 × 當期 EBITDA（v0.1b）；backlog＝模板的債務／backlog | 代碼 | leaseAdj | 檢查 |
| `defaults.debtEbitdaMax` | 投資級上限倍數：leaseAdj＝調整後槓桿上限（S&P BBB- 降評門檻 4.5×，v0.2）；ebitda＝總債務 ÷ 當期 EBITDA 上限 | 倍 | 4.5 | 檢查 |
| `defaults.dividend.perShareQ` | 普通股每股每季股利（Oracle v0.1b；不發股利的公司刪除 dividend 區段） | US$ | 0.5 | 必改 |
| `defaults.dividend.sharesBase` | 股利的基礎股數（最新流通股；另加前期累計瀑布新股與已強制轉換特別股） | bn 股 | 3.02374 | 必改 |
| `defaults.dividend.preferred` | 特別股股利，各期 | US$bn 清單 | 0.244、0.325、0.244、0、0 | 必改 |
| `defaults.dividend.note` | 股利的來源與推導說明 | 文字 | 普通股股利＝每股每季 $0.50 × 4 × 期間長… | 必改 |
| `defaults.otherEbitdaNote` | 其他事業 EBITDA 的推導與來源說明 | 文字 | 不適用（Oracle 無需另列的非核心事業燒錢；傳統… | 必改 |
| `defaults.debtBacklog` | 新債上限：總債務不超過 backlog 的倍數 | 倍 | 0.5 | 檢查 |
| `defaults.ctrTerm` | 新簽合約的平均年期（決定 backlog 補入量） | 年 | 3 | 檢查 |
| `defaults.minCash` | 最低現金：每期融資後期末現金不低於此值 | US$bn | 10 | 檢查 |
| `defaults.eqPx` | 新股發行參考價（預設＝現價） | US$ | 144.77 | 必改 |
| `defaults.eqDisc` | 新股發行折價 | 比例 | 0.1 | 檢查 |
| `defaults.eqCapPct` | 每年股權募資上限（占現市值）；輸入 9 以上視為無上限 | 比例 | 0.05 | 檢查 |
| `defaults.eqCapShares` | 股權年上限的股數基礎：現市值＝發行參考價 × 此股數（5a；評價日時點，列入滾動檢查） | bn | 3.02374 | 必改 |
| `defaults.junkRate` | 股權上限用完後的高息債利率 | 比例 | 0.12 | 檢查 |
| `defaults.ppeOpen` | 最新季末固定資產毛額 | US$bn | 75.74 | 必改 |
| `defaults.jvCommit` | 已承諾的 JV 出資餘額，各期 | US$bn 清單 | 0、0、0、0、0 | 必改 |
| `defaults.intCal` | 第一期利息校準值（讓模型利息對上公司季度指引） | US$bn | 0 | 必改 |
| `defaults.rpoOpen` | 最新季末 RPO | US$bn | 664 | 必改 |
| `defaults.rpoPendingAdd` | 季末後新簽、尚未進 RPO 的承諾 | US$bn | 0 | 必改 |
| `defaults.rp` | RPO 在模型期內認列的比例（百分點；應與 rpo.scheduledShare 一致） | % | 79.75 | 必改 |
| `defaults.cash` | 最新季末現金（第一期期初現金） | US$bn | 36.369 | 必改 |
| `defaults.includeDebt` | 是否依到期表攤還既有債務（true＝是；false＝假設全數再融資） | 是／否 | 是 | 可沿用 |
| `defaults.includeAtm` | 是否計入期後股權／可轉債募資 | 是／否 | 是 | 可沿用 |
| `defaults.atm` | 期後股權／可轉債募資淨額（記在第一期） | US$bn | 0 | 必改 |
| `defaults.prepay.shareOfDeals` | 有預付的合約比例（預付流入＝成長型 CapEx × 此比例 × 下一欄；覆蓋比已是整體口徑時填 1；v0.1b） | 比例 | 1 | 檢查 |
| `defaults.prepay.capexCover` | 預付（客戶出資）占相關資本支出的比例 | 比例 | 0.243 | 檢查 |
| `defaults.prepay.recogYears` | 預付在合約期內的認列年數：依(期初合約負債＋本期累積利息)直線認列為營收（非現金） | 年 | 5 | 檢查 |
| `defaults.prepay.coverRefresh` | 客戶出資覆蓋比是否也適用 GPU 汰換 CapEx（true＝預付流入＝(成長型＋汰換)× 覆蓋比；Oracle v0.1c） | 是／否 | 是 | 檢查 |
| `defaults.prepay.financingRate` | 預付重大財務組成的隱含利率：合約負債以此利率累積非現金利息（期初餘額＋本期流入一半），認列時轉營收；0＝不計財務組成（Oracle v0.1b） | 比例 | 0.0811 | 檢查 |
| `defaults.prepay.openBalance` | 期初合約負債（客戶預付餘額；列入滾動檢查） | US$bn | 15.955 | 必改 |
| `defaults.prepay.note` | 預付款區塊的說明（來源與口徑） | 文字 | 客戶出資（Oracle v0.1b 步驟 4）：預付… | 必改 |
| `defaults.convIssue.coupon` | 瀑布新發可轉債的票息（v0.1b） | 比例 | 0 | 檢查 |
| `defaults.convIssue.premium` | 瀑布新發可轉債的轉換溢價（轉換價＝發行參考價 ×（1＋此值）；只用於潛在股數揭露） | 比例 | 0 | 檢查 |
| `defaults.convIssue.note` | 可轉債步驟的說明（來源與口徑） | 文字 | 不適用（Oracle 融資瀑布不設可轉債步驟；sce… | 必改 |
| `defaults.overlay` | 電力／維護成本另計（預設關；EBITDA 率已含電費，開啟會重複扣除） | 是／否 | 否 | 可沿用 |
| `defaults.cdsLink` | CDS 利差是否傳入新債利率（預設關） | 是／否 | 否 | 可沿用 |
| `defaults.cdsBaseBp` | CDS 傳入新債利率的門檻：超過此值的部分才傳入（5a） | bps | 450 | 檢查 |
| `defaults.cdsPassThrough` | CDS 超額傳入新債利率的比例（每 1bp 傳入的 bp；5a） | 比例 | 0.4 | 檢查 |
| `defaults.linkSites` | 具名站點的 MW 是否連動第一期產能下限 | 是／否 | 是 | 可沿用 |
| `defaults.linkLeaseTail` | 舊模板遺留開關，目前程式未使用 | 是／否 | 是 | 可沿用 |
| `defaults.useAvgMw` | 收入以平均在役 MW 計（true）或期末存量計（false） | 是／否 | 是 | 可沿用 |
| `defaults.billableOpen` | 最新季末可計費 MW（第一期期初） | MW | 1698 | 必改 |
| `defaults.cds` | 信用違約交換（CDS）中價 | bps | None | 必改 |
| `defaults.cdsBid` | CDS 買價 | bps | None | 必改 |
| `defaults.cdsAsk` | CDS 賣價 | bps | None | 必改 |
| `defaults.cdsLo` | CDS 近期區間下緣 | bps | None | 必改 |
| `defaults.cdsHi` | CDS 近期區間上緣 | bps | None | 必改 |
| `defaults.cdsDate` | CDS 報價日期與來源標記 | 文字 | 不適用（ORCL CDS 無可引用的一手報價） | 必改 |
| `defaults.useFacility` | 融資時是否先動用未動用信用額度 | 是／否 | 否 | 可沿用 |
| `defaults.facility` | 未動用信用額度 | US$bn | 10 | 必改 |
| `defaults.m.accepted` | 已驗收 MW 的預設路徑（實際依所選情境覆寫） | MW 清單 | 4668、7501、8778、8778、8778 | 檢查 |
| `defaults.m.billable` | 可計費 MW 的預設路徑（實際依情境與爬坡比例覆寫） | MW 清單 | 3183、6340、7716、8070、8070 | 檢查 |
| `defaults.m.util` | 利用率，各期 | % 清單 | 100、100、100、100、100 | 檢查 |
| `defaults.m.revMW` | 每 MW 年收入，各期 | US$bn/MW 清單 | 0.0174、0.0174、0.0174、0.0174、0.0174 | 必改 |
| `defaults.m.aiShare` | AI 占比，各期（目前只做範圍檢查，未參與計算） | % 清單 | 100、100、100、100、100 | 可沿用 |
| `defaults.m.fill` | 新產能簽約率：未被既有 RPO 占用的產能能賣出的比例 | % 清單 | 100、100、100、100、100 | 檢查 |
| `defaults.m.power` | 電價（overlay 開啟時才用） | $/MWh 清單 | 60、62、64、66、68 | 檢查 |
| `defaults.m.pue` | 電力使用效率 PUE（overlay 用） | 倍 清單 | 1.2、1.2、1.2、1.2、1.2 | 檢查 |
| `defaults.m.maint` | 維護成本（overlay 用） | US$m/MW 清單 | 0.15、0.15、0.16、0.16、0.17 | 檢查 |
| `defaults.m.defaultP` | 客戶違約率 | % 清單 | 0.5、1、1.5、2、2.5 | 檢查 |
| `defaults.m.recovery` | 違約回收率 | % 清單 | 50、50、50、50、50 | 檢查 |
| `defaults.m.rate` | 新債利率 | % 清單 | 5.8、5.8、5.8、5.8、5.8 | 檢查 |
| `defaults.terminal.residual` | 模型期末 GPU 殘值率 | % | 25 | 檢查 |
| `defaults.terminal.rerent` | 期末設備再出租率 | % | 75 | 檢查 |
| `defaults.terminal.margin` | 模型期後剩餘 RPO 的利潤率 | % | 40 | 檢查 |
| `defaults.terminal.residualLeaseYears` | 模型期後租約剩餘年數 | 年 | 12 | 檢查 |
| `defaults.sites` | 具名資料中心站點，一站一列。欄位：id 代碼、name 名稱、operator 房東／合作方、planned 契約 MW、energized 已通電 MW、accepted 已驗收 MW、billable 可計費 MW、contract 合約總值（US$bn，可無）、years 合約年期（可無）、status 狀態說明、next 下一里程碑、date 預計時間、confidence 信心（高／中／低） | 清單 | 5 筆 | 必改 |

### `valuation`：評價參數

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `valuation.price` | 現價 | US$ | 144.77 | 必改 |
| `valuation.shares` | 評價股數（含期後股權發行上限） | bn 股 | 3.05774 | 必改 |
| `valuation.atmSharesInValuation` | 評價股數中「期後股權發行上限」的股數：期後股權／可轉債開關關閉時由評價股數扣回（v0.1b；CRWV 0.035、無此項的公司填 0） | bn 股 | 0 | 必改 |
| `valuation.netDebt` | 淨負債（不含可轉債：其他借款 − 現金，含期後已入帳的股權／可轉債募得淨額；可轉債依 debt.convertibles 另計） | US$bn | 88.827 | 必改 |
| `valuation.holdings` | 持股清單，一筆一列：[名稱, 估值（100%，US$bn）, 持股比例, 備註]；價值＝估值 × 持股比例 ×（1 − holdingsDiscount），自淨負債扣除 | 清單 |  | 必改 |
| `valuation.holdingsDiscount` | 持股折價（流動性、少數股權） | 比例 | 0.3 | 必改 |
| `valuation.debtLike` | 類債項目，一筆一列：[名稱, 金額（US$bn）, 備註]；加入淨負債 | 清單 |  | 必改 |
| `valuation.tax` | 稅率 | 比例 | 0.151 | 檢查 |
| `valuation.nol` | 期初可扣抵虧損（NOL） | US$bn | 0 | 必改 |
| `valuation.wacc` | 加權平均資金成本 WACC 手動覆蓋（null＝採 CAPM） | 比例或 null | None | 檢查 |
| `valuation.capm.beta` | CAPM β | 倍 | 1.77 | 檢查 |
| `valuation.capm.erp` | CAPM 股權風險溢酬 | 比例 | 0.05 | 檢查 |
| `valuation.capm.kdPretax` | 稅前債務成本（市場邊際） | 比例 | 0.0811 | 檢查 |
| `valuation.capm.betaSens` | β 敏感度（報告用） | 倍 清單 | 1.2、1.5 | 檢查 |
| `valuation.capm.note` | WACC 公式與來源說明 | 文字 | WACC＝E/(D+E)×(rf＋β×ERP)＋D/… | 必改 |
| `valuation.nolUsePct` | NOL 每年可抵用上限占應稅所得的比例（美國 80%；依公司稅籍調整；5a） | 比例 | 0.8 | 檢查 |
| `valuation.wcPctOfRevGrowth` | 營運資金變動占營收增量的比例（DCF 自由現金流；5a） | 比例 | 0.02 | 檢查 |
| `valuation.g` | 永續成長率 | 比例 | 0.03 | 檢查 |
| `valuation.sbc` | 年度股份基礎薪酬 | US$bn | 4.508 | 必改 |
| `valuation.maintRatio` | 終值的維持性資本支出占折舊比例 | 比例 | 0.8 | 檢查 |
| `valuation.tvBasis` | 終值基準：ufcf＝末期 UFCF（含穩態汰換 CapEx）；其他值＝常態化 FCF（EBIT ×(1−稅)＋D&A × (1 − 維持比率)）（Oracle v0.1c） | 代碼 | ufcf | 檢查 |
| `valuation.evEbitda` | EV/EBITDA 倍數（分部加總的 OCI／算力部分） | 倍 | 6 | 檢查 |
| `valuation.evYear` | EV/EBITDA 錨定年度（1＝模型第 2 期 … 4＝第 5 期） | 年度代碼 | 2 | 檢查 |
| `valuation.legacyEvEbitda` | 傳統事業 EV/EBITDA 手動覆蓋（null＝軟體同業中位數） | 倍或 null | None | 檢查 |
| `valuation.legacyEvEbitdaNote` | 分部加總說明 | 文字 | 傳統事業（雲端應用＋授權＋硬體＋服務）EV/EBIT… | 必改 |
| `valuation.dcfMode` | DCF 股權為負時的處理：zero＝0 截斷、option＝選擇權法 | 文字 | zero | 可沿用 |
| `valuation.sigma` | 企業價值波動率（選擇權法用） | 比例 | 0.5 | 檢查 |
| `valuation.rf` | 無風險利率（CAPM 與選擇權法用） | 比例 | 0.0527 | 檢查 |

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
| `methodology.checks.capexPerMwBand` | 檢查頁：模型期 CapEx 強度（每 MW 百萬美元）的合理區間下端與上端（5a） | US$m/MW 清單 | 28、51 | 檢查 |
| `methodology.checks.leaseVsCommitMin` | 檢查頁：表外租金路徑 ÷ 已承諾租約至少要達到的倍數（5a） | 倍 | 0.8 | 檢查 |
| `methodology.checks.unsignedRevShareMax` | 檢查頁：後段年度依賴未簽約收入的比例上限（5a） | 比例 | 0.5 | 檢查 |
| `methodology.checks.siteRentGapMax` | 檢查頁：站點租賃五期租金可能低估的金額上限（5a） | US$bn | 10 | 檢查 |
| `methodology.checks.rentVsBenchMin` | 檢查頁：模型每 MW 年租金至少要達到「市場基準 × 第三方占比」的比例（5a） | 比例 | 0.8 | 檢查 |

### `peers`：同業比較（Comps）

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `peers.priceDate` | 同業市值的收盤日 | 日期 | 2026-09-21 | 必改 |
| `peers.priceSource` | 同業市值來源 | 文字 | S&P Global via stockanalys… | 必改 |
| `peers.list` | 同業，一家一列（HTML Comps 分頁與 Excel「可比公司」頁共用）。欄位：ticker 代號、name 公司名、labelHtml／labelXlsx 兩邊顯示的名稱、roleHtml 定位說明（HTML）、mkt 市值、netDebt 淨負債、rev 近十二個月營收、opl 營業租賃負債、opInc GAAP 營業損益、da 折舊攤銷（以上 US$bn）、asOf 資料期、noteHtml／noteXlsx 兩邊的備註。EV＝市值＋淨負債、EBITDA＝營業損益＋折舊攤銷，由程式計算 | 清單 | 5 筆 | 必改 |
| `peers.software` | 軟體同業 NTM EV/EBITDA，一家一列（ticker、name、ntmEvEbitda、ref＝事實總帳 id）；中位數為傳統事業倍數 | 清單 | 5 筆 | 必改 |
| `peers.softwareNote` | 軟體同業倍數的來源說明 | 文字 | 軟體同業 NTM EV/EBITDA（EV：Stoc… | 必改 |
| `peers.textHtml.headerTip` | HTML Comps 表標題的浮動說明 | 文字 | AI 基礎設施同業（Neocloud 與房東）。市值… | 必改 |
| `peers.textHtml.readingTip` | HTML「讀法」段落的浮動說明（折價來源） | 文字 | ORCL 列為整體公司（傳統軟體＋OCI），倍數不可… | 必改 |
| `peers.textHtml.caveat` | HTML Comps 表下方的「口徑與限制」 | 文字 | 口徑與限制：市值日期不一致（見上），負債與現金為各公… | 必改 |
| `peers.textXlsx.subtitle` | Excel「可比公司」頁第 2 列說明 | 文字 | EV＝市值＋（有息負債−非受限現金）；含租賃＝EV＋… | 必改 |
| `peers.textXlsx.notes` | Excel「可比公司」頁表下方的讀法說明（每句一列） | 文字清單 | 讀法：ORCL 為整體公司（傳統軟體＋OCI），倍數…、CRWV 以租約取得機房，加入營業租賃負債後倍數上升…、前瞻：見下方 ORCL 模型列。此表為市場比較用，不… | 必改 |

### `quarterly`：季度層（v4.4）

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `quarterly._note` | 季度層的說明文字（不進程式） | 文字 | 季度層（v4.4）：由年度模型拆分，只用於追蹤，不影… | 可沿用 |
| `quarterly.quarters` | 追蹤的季度，一季一列：key 季別代碼（與共識檔 quarterlyEstimates 的鍵相同，例如 2026Q3）、label 顯示名稱、period 所屬模型期（0＝第 1 期…）、reportNote 財報日說明（選填）。季度加總必須等於所屬模型期的年度數字 | 清單 | 6 筆 | 必改 |
| `quarterly.periodNames` | 各所屬模型期的顯示名稱（依期別順序，例如 2H26、FY27） | 文字清單 | 2Q–4Q27、FY28 | 必改 |
| `quarterly.focus` | 焦點季（「季度追蹤」第一屏與一頁摘要驗證點顯示的季度；財報後改為下一季） | 季別代碼 | FY27Q2 | 必改 |
| `quarterly.keyMetrics` | 一頁摘要「驗證點」列出的指標（2–3 項；metrics 的 key） | 文字清單 | revenue、adjEbitda、capex | 檢查 |
| `quarterly.driver` | 拆分依據：type＝mw 時依 MW 內插（endStart＝第 1 期期初 Accepted MW、endStartNote 來源）；type＝none（沒有 MW 資料的公司）時營收依 revenueSplit、CapEx 平分，MW 列顯示「不適用」 | 物件 | 物件（type、endStart、endStartNote） | 檢查 |
| `quarterly.revenueSplit` | 各所屬模型期的營收拆法：anchor＝從 revenueAnchor 逐季線性爬升、driverAvg＝依平均在役 MW、equal＝平分 | 文字清單 | anchor、driverAvg | 檢查 |
| `quarterly.revenueAnchor` | 營收拆法 anchor 的起點（最新一季實際營收）：value、label、tag | 物件 | 物件（value、label、tag） | 必改 |
| `quarterly.metrics` | 追蹤的指標，一項一列：key（revenue、adjEbitda、ebitdaMargin、adjOpInc、capex、mw 之一）、label、unit（US$bn／%／MW）、gap 差距寫法（ratio＝比例，用於營收、CapEx；diff＝金額差；pt＝百分點）、marginOf 利潤類另列利潤率百分點時的分母指標（revenue）、consensus 共識檔季度欄位名（derived＝由共識 EBITDA ÷ 營收換算）、consensusSecondary／consensusSecondaryLabel 只列不比較的共識欄位、actualLabel 實際數列名稱 | 清單 | 6 筆 | 檢查 |
| `quarterly.capexSplit` | CapEx 拆法：guidanceAnchor＝有季度指引的季取指引中點、其餘季分配期間餘數（依新增 MW）；driverAdds（或不填）＝全部依新增 MW | 文字 | guidanceAnchor | 檢查 |
| `quarterly.guidanceText` | 文字型指引（例如「調整後營業利益率 low teens」），{季別: {指標: {text, source, tag}}}：只列不計差距；完整支援（年增率、利潤率區間）列入待辦 5 | 物件 | 物件（） | 必改 |
| `quarterly.guidance` | 季度指引（公司預估），{季別: {指標: [低, 高]}}；沒有指引的季度或公司不填，顯示「不適用」 | 物件 | 物件（FY27Q2） | 必改 |
| `quarterly.guidanceMeta` | 季度指引的來源與標記 | 物件 | 物件（source、tag） | 必改 |
| `quarterly.periodGuidance` | 期間隱含指引，{模型期: {指標: {range: [全年低, 全年高], less: 已實現部分}}}：只用於判斷季度差距是否為「拆法」 | 物件 | 物件（0） | 必改 |
| `quarterly.periodGuidanceNote` | 上一欄的說明文字 | 文字 | 全年指引 − 1Q 實際＝2Q–4Q 隱含指引 [D… | 必改 |
| `quarterly.consistency` | 建置檢查：[路徑 A, 路徑 B, 倍數（選填）]，兩者（B × 倍數）必須相同，否則建置失敗（防止同一數字在兩處不一致） | 清單 | 6 筆 | 檢查 |
| `quarterly.actuals` | 季度實際數（公司公布後填入；預設空白＝待公布），{季別: {各指標, source, date, tag}}；Excel 對應「輸入與假設」J 區藍字格 | 物件 | 物件（FY27Q2、FY27Q3、FY27Q4、FY28Q1、FY28Q2、FY28Q3） | 必改 |

### `varianceReasons`：差異原因（v4.4）

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `varianceReasons._note` | 差異原因的說明文字（不進程式） | 文字 | 差異原因（已決定事項 2）：差距超過 methodo… | 可沿用 |
| `varianceReasons.list` | 差異原因（已決定事項 2），一筆一列：scope（annual 年度共識對照／quarter 季度）、period（FY27、2026Q3 或 *）、metric（年度：rev、ebitda、capex、nd；季度：metrics 的 key）、vs（consensus、guidance、actual 或 *）、type（觀點／已知限制）、text 一句原因，{路徑:格式} 由模型數字帶入。「拆法」由程式判定，不需填。差距超過 methodology.consensusGapTol 卻沒有原因時建置失敗 | 清單 | 8 筆 | 檢查 |

### `texts`：公司特有的說明文字（v4.5；隨資料更新）

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `texts.fy0EquityNote` | 首期股權／可轉債的組成說明（含年初至今與首期模型的金額；v4.5；可用期間佔位符 «P0»、«YTD»、«STUB»） | 文字 | «P0»＝«YTD» ATM 淨收入 19.909（… | 必改 |
| `texts.sourceLine` | 頁首的資料來源一行（例如最新 10-Q、法說、期後 8-K；v4.5） | 文字 | FY27Q1 10-Q（2026-09-11）＋財報… | 必改 |
| `texts.cashTaxNote` | 年初至今現金稅的說明（損益與評價頁；v4.5） | 文字 | «YTD» 現金稅未單獨揭露；10-Q 有效稅率 1… | 必改 |
| `texts.thesis` | 模型命題一句話（HTML 標題列與檔案說明；v0.1b） | 文字 | 客戶預付與投資級債能否把 $664B RPO 變成現… | 必改 |
| `texts.guideLine` | 公司年度指引的一句摘要（損益頁、檢查與來源頁；v0.1b） | 文字 | FY27 公司指引：營收 ≥$90B、非 GAAP … | 必改 |
| `texts.taxNote` | 稅率的來源說明（v0.1b） | 文字 | FY27Q1 有效稅率（10-Q）；FY26 為 1… | 必改 |
| `texts.revMwCompare` | 每 MW 年收入的公司對照值說明（只作對照；v0.1b） | 文字 | Oracle 對照值：OCI 年化 ÷ 新增 MW … | 必改 |
| `texts.rpoNote` | RPO 桶的說明（桶界、後段假設、口徑；v0.1b） | 文字 | 12 個月內 13%、36 個月內合計 50%；其餘… | 必改 |
| `texts.leaseNote` | 租賃結構的一句說明（自有或租賃為主；v0.1b） | 文字 | Oracle 資料中心以租賃為主（開發商建機房、Or… | 必改 |
| `texts.offBalanceLeaseTerm` | 未起租租賃的起租時程與期限（季報揭露；v0.1b） | 文字 | FY27Q2–FY2029 起租、期限 15–19 年 | 必改 |
| `texts.leaseLiabNote` | 租賃負債的折現率與期限說明（v0.1b） | 文字 | 10-Q 附註 6；FY26 加權平均折現率 5.7% | 必改 |
| `texts.capexGuideSource` | 資本支出指引的來源與可信度（v0.1b） | 文字 | 法說會二手逐字稿；新聞稿與 10-Q 未列 [Int… | 必改 |
| `texts.prepayCheck` | 客戶預付的公司說法與推得覆蓋比（檢查頁；v0.1b） | 文字 | 公司：Q1 新簽 >$30B AI 合約以客戶預付或… | 必改 |
| `texts.mwFacts` | 產能（MW）相關的公司揭露摘要（檢查頁；v0.1b） | 文字 | Q4 FY26 簡報與 Q1 FY27 新聞稿：FY… | 必改 |
| `texts.ebStartSource` | 起始 EBITDA 率的來源或推導方式（v0.1b） | 文字 | FY27Q1 合併 EBITDA 扣除傳統事業後拆推… | 必改 |
| `texts.costMwNote` | 每 MW 建置成本的來源與口徑（v0.1b） | 文字 | Tokenomics IF_CapexIT（GB30… | 必改 |
| `texts.ppeOpenNote` | 期初 PP&E 基礎的來源或校準方式（v0.1b） | 文字 | 以 FY27Q1 實際折舊 3.156 × 4 × … | 必改 |
| `texts.facilityName` | 未動用額度的名稱與條件（v0.1b） | 文字 | 循環信用額度，未動用 | 必改 |
| `texts.newDebtRateNote` | 瀑布新債利率的來源（v0.1b） | 文字 | 瀑布新債利率取 2026-02 一次性投資級無擔保債… | 必改 |
| `texts.concentration` | 客戶集中的標題與說明（title、detail；檢查頁與信用分頁；v0.1b） | 物件（文字） | 物件（title、detail） | 必改 |
| `texts.sharesNote` | 評價股數在季末流通股之外的組成（v0.1b） | 文字 | RSU 與選擇權（稀釋與基本加權股數差 0.034bn） | 必改 |
| `texts.netDebtNote` | 淨負債輸入的口徑說明（v0.1b） | 文字 | 不含租賃與強制轉換特別股（債務面額 − 現金 − 有… | 必改 |
| `texts.otherRevNote` | 非算力服務／其他事業營收的說明（v0.1b） | 文字 | 非算力服務預設 0（傳統事業四線另列，v0.1b 步… | 必改 |
| `texts.rvRevNote` | 反向 DCF 每 MW 年收入列的對照說明（v0.1b） | 文字 | Oracle 對照：OCI 年化 ÷ 新增 MW 1… | 必改 |
| `texts.rvCostNote` | 反向 DCF 每 MW 建置成本列的對照說明（v0.1b） | 文字 | FY27 指引 90–95 ÷ 估計新增 MW-IT… | 必改 |
| `texts.priceNote` | 定價（續約價、新約價）的公司說法與讀法（v0.1b） | 文字 | 公司稱到期 GPU 續約價 +20%（4 年以上 G… | 必改 |
| `texts.prepayCoverNote` | 客戶預付覆蓋比的來源（v0.1b） | 文字 | Oracle：FY27 指引隱含客戶出資覆蓋比 24… | 必改 |
| `texts.prepayOpenNote` | 期初合約負債的口徑（v0.1b） | 文字 | 客戶預付（重大財務組成）累計現金 15.955，非全… | 必改 |
| `texts.ytdEquityNote` | 年初至今股權募資的組成（v0.1b） | 文字 | ATM 淨收入，141m 股、均價約 $141.84 | 必改 |
| `texts.prepayRiskNote` | 客戶預付的風險與會計說明（v0.1b） | 文字 | Oracle 10-Q 稱預付含重大財務組成（預付不… | 必改 |
| `texts.consistencyTitle` | 「公司說法 vs 季報」對照表的標題（v0.1b） | 文字 | 新聞稿／法說 vs 季報 vs 期後公告：一致、未量… | 必改 |
| `texts.callVsFiling` | 來源頁「法說／季報／本模型取捨」對照表，一列一項：[項目, 前瞻說法, 季報, 本模型取捨]（v0.1b） | 清單 | 15 筆 | 必改 |
| `texts.commitments` | 來源頁「表外與契約性支出總表」，一列一項：[項目, 金額, 時程, 入表？, 模型位置]（v0.1b） | 清單 | 6 筆 | 必改 |
| `texts.consistency` | 來源頁「公司說法 vs 季報」一致性表，一列一項：[項目, 公司說法, 季報, 判定]（v0.1b） | 清單 | 8 筆 | 必改 |
| `texts.sources` | 來源頁的來源清單，一列一筆：[標記, 內容]（HTML 與 Excel「來源」頁共用；v0.1b） | 清單 | 9 筆 | 必改 |
| `texts.headerCards` | 資金模型頁首的公司揭露卡片，一列一張：[標題, 內容]（v0.1b） | 清單 | 3 筆 | 必改 |
| `texts.legacyMarginNote` | 傳統事業 EBITDA 率的推導說明（v0.1b） | 文字 | FY26 合併 EBITDA 率＝非 GAAP 營業… | 必改 |
| `texts.mwYearEndNotes` | 各年底主動電力的來源說明，以年份為鍵（Excel「輸入與假設」說明欄；5a） | 物件（文字） | 物件（2026） | 必改 |

### `oracle`：Oracle 資料草稿（v0.1a；v0.1b 逐步搬入，只留後續步驟用或只作對照的欄位）

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `oracle._readme` | Oracle 資料草稿區段的說明（v0.1a 產出；引擎尚未讀取，v0.1b 依 mapTo 搬入既有欄位或新增） | 文字 | Oracle 資料草稿（v0.1a 產出）。v0.1… | 必改 |
| `oracle.files` | Oracle 事實總帳、每 MW 推導、共識資料檔的路徑 | 物件（路徑） | 物件（facts、perMw、consensus） | 必改 |
| `oracle.debt` | Oracle 債務草稿（逐檔票券、定期貸款、商業本票、到期梯、循環額度、強制轉換特別股、信評）；格式同上 | 物件 | 物件（ratings） | 必改 |
| `oracle.equity` | Oracle ATM、FY27 融資計畫、股利、買回授權、股價草稿；格式同上 | 物件 | 物件（atmDone、fundingPlanFY27、dividendPerShareQ、buybackAuth、eqPx） | 必改 |
| `oracle.prepay` | Oracle 客戶出資覆蓋比、預付累計、預付＋自帶硬體合約額草稿；格式同上 | 物件 | 物件（capexCover、cashCum、prepayAndByohContract、q1NewContracts） | 必改 |
| `oracle.mw` | Oracle 已交付 MW、站點、合約容量、利用率、續約溢價草稿（口徑逐欄註明）；格式同上 | 物件 | 物件（deliveredFY26、deliveredFY27Q1、deliveredCumSinceFY26、abileneDelivered、openaiContract、oracleStargateSitesEpoch、gpusQ1、utilization、renewalPremium、pue） | 必改 |
| `oracle.legacy` | Oracle 傳統事業四線（SaaS、軟體、硬體、服務）近四季營收、成長率、分部利潤率草稿；格式同上 | 物件 | 物件（revenueLTM、growthLTM、segMarginFY26、consolidatedNonGaapOpMarginFY26） | 必改 |
| `oracle.guidance` | Oracle 公司指引草稿（Q2 FY27、FY27、OCI 路徑與 FY30 目標只作對照）；格式同上 | 物件 | 物件（q2Revenue、q2Cloud、q2EpsNonGaap、fy27Revenue、fy27EpsNonGaap、fy27Capex、fy27NetCashCapexMax、ociPathFY26PR、fy30Targets） | 必改 |
| `oracle.valuation` | Oracle 評價輸入草稿（beta、無風險利率、ERP、債務成本、稅率、同業倍數、淨負債、持股）；格式同上 | 物件 | 物件（beta、rf、erp、costOfDebtPretax、taxRate、softwarePeerNtmEvEbitda、ociEvEbitda、netDebtExLeases、tiktokStake） | 必改 |
| `oracle.perMw` | Oracle 每 MW 年收入三情境（沿用 Nebius 的 Tokenomics 推導）、EBITDA 率、伺服器壽命草稿；格式同上 | 物件 | 物件（revPerMWit、ebitdaSteady、ebitdaStartCompanyClaim、serverLife） | 必改 |
| `oracle.openai` | Oracle–OpenAI 合約年額、隱含每 MW、OpenAI 計畫算力支出（只作對照）；格式同上 | 物件 | 物件（contractAnnual、contractPerMw、computePlan2026to2030） | 必改 |
| `oracle.events` | Oracle 評價日後事件（Project Jupiter 不可抗力通知）；格式同上 | 物件 | 物件（jupiterForceMajeure） | 必改 |

### `mag`：MAG（Alphabet）資料草稿（MAG v0.1a；引擎尚未讀取，v0.1b′ 依 mapTo 搬入）

| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |
|---|---|---|---|---|
| `mag._readme` | MAG（Alphabet）資料草稿區段的說明（v0.1a 產出；引擎尚未讀取，v0.1b′ 依 mapTo 搬入 MAG 共用引擎欄位） | 文字 | Alphabet MAG 資料草稿（v0.1a 產出… | 必改 |
| `mag.files` | MAG 事實總帳、k 證據表、Tokenomics 快照與名稱清單、市場共識資料檔的路徑 | 物件（路徑） | 物件（facts、kEvidence、tokenomicsSnapshot、tokenomicsNames、consensus） | 必改 |
| `mag.calendar` | MAG 期間草稿（最新已申報季、年初至今月數）；每欄 value＋unit＋ref（總帳 id）＋mapTo | 物件 | 物件（latestQuarterFiled、ytdMonths） | 必改 |
| `mag.segments` | MAG 各營收線近四季營收、TTM、年增（搜尋、YouTube、聯播網、訂閱平台裝置、雲端、Other Bets）；格式同上 | 物件 | 物件（search、youtube、network、subsPlatformsDevices、cloud、otherBets） | 必改 |
| `mag.segmentOpInc` | MAG 各分部營業利益近四季與近三年（含 Alphabet 層級費用）；格式同上 | 物件 | 物件（googleServices、googleCloud、otherBets、alphabetLevel、total、annual） | 必改 |
| `mag.pl` | MAG 損益草稿（D&A、SBC、利息、稅率、權益證券一次性利益）；格式同上 | 物件 | 物件（da、sbc、interestExpense、taxRate、equityGainsH1） | 必改 |
| `mag.cashflow` | MAG 現金流草稿（營運現金流、資本支出、FCF、回購、股利、特別股股利）；格式同上 | 物件 | 物件（cfo、capex、fcf、buybacks、dividends、prefDividend） | 必改 |
| `mag.capexGuidance` | MAG 資本支出指引、組成與共識（D3 首期與次期）；格式同上 | 物件 | 物件（fy2026、fy2027、mix、consensus） | 必改 |
| `mag.balance` | MAG 資產負債草稿（現金、債務、租賃與未起租承諾、RPO、採購承諾、擔保、評等）；格式同上 | 物件 | 物件（cashAndSecurities、debtFace、debtPostQ、commercialPaper、leases、rpo、purchaseCommitments、backstops、rating） | 必改 |
| `mag.shares` | MAG 股數、強制轉換特別股、回購授權、ATM、2026 股權募資；格式同上 | 物件 | 物件（outstanding、dilutedWeightedQ2、mcps、buybackAuthRemaining、atm、equityRaised2026） | 必改 |
| `mag.ai` | MAG AI 容量草稿（期初在役 MW、對外比例、新增速度、已公開管線、TPU 出貨、用電交叉檢查；口徑 MW-IT）；格式同上 | 物件 | 物件（mwInService、extShare、mwExternal、addPerYear、publicPipeline、companyTarget、tpuShipments、envCheck） | 必改 |
| `mag.perMw` | MAG 每 MW 年收入錨（IF_HoldEcon 世代組合）與 k 建議值；格式同上 | 物件 | 物件（anchor、k、tokenomics） | 必改 |
| `mag.split` | MAG AI／非 AI 雲端拆分試算起點（只作 v0.1b 校準起點）；格式同上 | 物件 | 物件（cloudAnnualized、aiCloudBase、nonAiBase） | 必改 |
| `mag.relatedParties` | MAG 關聯方（Anthropic 持股、承諾、雲端承諾、支出路徑；SpaceX 持股）；格式同上 | 物件 | 物件（anthropicStake、anthropicPostMoney、anthropicCommitment、anthropicCloudCommit、anthropicSpendPath、spacexHolding） | 必改 |
| `mag.rentals` | MAG 租用算力（SpaceX 雲端服務合約）；格式同上 | 物件 | 物件（spacex） | 必改 |
| `mag.valuation` | MAG 評價輸入草稿（beta、無風險利率、債務成本、ERP、現價、同業倍數、共識目標價）；格式同上 | 物件 | 物件（beta、rf、kdPretax、erp、price、peerMultiples、consensusPT） | 必改 |

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
