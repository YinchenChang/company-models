# Nebius 收支模型 — Claude Code 工作守則

（本檔不寫版本號；目前版本見 `docs/handoff/` 內交接檔檔名與 `vlog.py`。）

## 專案是什麼
Nebius（NBIS）資金與評價模型，核心命題「預付款是否讓 backlog 真的變成現金」（CRWV「Backlog 不是現金」的對照組）。由 CRWV 模型 v4.5（`YinchenChang/crwv-model` @ `01b13ad`）複製而來，本資料夾各自完整可建置。同一套引擎輸出兩個成品，數字必須完全一致：
- HTML（互動版，`build_html_portable.py` 產生）
- Excel（全部活公式，`build_xlsx.py` 產生）
業務背景、目前結果與待辦見 `docs/handoff/` 內的交接檔（資料夾內只保留最新版）；技術細節見 `README.md`。工作流程與範圍以 repo 根目錄 `README.md` 與 `docs/workorders/` 的工作單為準（衝突時優先於本檔）。

## 目錄
- `company.json`：公司原始輸入的**唯一來源**（HTML 與 Excel 共用）。換公司先改這裡。
- `segA.js`～`segE.js`、`tail.js`：HTML 引擎與畫面；`app_pretty.js`：模板函式庫片段來源。
- `build_xlsx.py`、`fix_outline.py`、`verify_ooxml.py`：Excel 建置與結構檢查。
- `cmp31.js`＋`xlx.py`（HTML vs Excel 數值）、`crawl.py`（畫面文字）、`xl_diff.py`（Excel 逐格）：核對工具。
- `dist/`：交付成品（HTML、xlsx）。`docs/handoff/`：交接檔（只保留最新版；升版時新增新版並移除舊版，供 Project 的 GitHub 同步勾選整個資料夾）；`docs/`：建置用模板 `template_v3_3.html`；`docs/reports/`：各任務的完整回報。

## 環境（CRWV 時期的第一個任務「環境移植」已完成；第 0 步的工具檢查每次開工仍適用）
0. 先檢查工具：`soffice --version`、`python3 -c "import openpyxl, playwright"`、Playwright 的 Chromium 能否啟動。雲端環境的 setup script 有 5 分鐘上限，可能沒裝完（紀錄在 `/tmp/crwv_setup.log`；v3 起只預裝 Python 套件）；缺什麼就在工作階段內補裝：`apt-get install -y --no-install-recommends libreoffice-calc-nogui`、`pip install --break-system-packages openpyxl pypdf playwright`、`python3 -m playwright install chromium`；Chromium 若因缺系統函式庫無法啟動，執行 `python3 -m playwright install-deps chromium`。回報時列出實際補裝了哪些。
1. 把 `/home/claude/...` 與 `/mnt/...` 的寫死路徑改為相對於 repo 根目錄（`build_xlsx.py` 的輸出、`fix_outline.py` 的 outline.json、`xlx.py` 的暫存檔、`build_html_portable.py` 的模板參數）。
2. 原本的 Excel 重算用沙盒內建的 `recalc.py`（此處沒有）：用 LibreOffice headless 自寫 `scripts/recalc.py`，功能＝開啟、全部重算、存檔，並回報公式錯誤數（#REF!、#DIV/0!、#VALUE!、#NAME?、#N/A）。
3. 新增 `scripts/verify.sh`，一次跑完：建 HTML → 建 Excel → 重算 → fix_outline → verify_ooxml → 三情境 xlx＋cmp31 → FY27 錨定 cmp31。任何一項失敗即非零結束。
4. 驗收：以 repo 內原始碼重建的 HTML／Excel，與 `dist/` 內 v4.0 成品比對：`crawl.py` 畫面文字 0 差異、`xl_diff.py --values` 0 差異、cmp31 三情境各 104 項全部 OK。

## 每次修改的鐵則
- 先讀交接檔與 README，再動手；純結構修改的驗收標準是「數字與畫面與前一版完全相同」。
- 每次修改後必跑 `scripts/verify.sh`，全數通過才可提交。
- 交付的 HTML 必須是「單一檔案、離線、本機雙擊即可開啟」（已決定事項 11）：所有 JS、CSS、字型、圖片、資料在建置時內嵌；不得有任何網路請求（CDN、Google Fonts、外部圖片、fetch／XHR、外部 import），也不得讀取旁邊的檔案。`verify.sh` 的「離線開啟檢查」（`scripts/check_offline.py`）必須通過。改到數字的修改，要在回報中列出受影響的關鍵數字（三情境目標價等）前後對照。
- 版本號：**只有 `dist/` 成品（HTML 或 Excel）改變時才升版**；純工具修改（建置／核對腳本、文件、setup script 等，`dist/` 不變）不升版、不新增 VLOG。升版時 `vlog.py`（Excel 唯一來源）與 `tail.js` 的 VLOG（HTML）**兩處都要**新增一列；成品命名 `更新日_<company.json → meta.company>收支模型_v版本號`（目前 `20261007_Nebius收支模型_v0_1`），放 `dist/`，並移除舊版成品；交接檔同樣以新版取代舊版（`docs/handoff/` 只留一個檔）。
- 同時更新交接檔（`docs/handoff/`，只保留最新版）與 README。
- 一個任務一個分支、一個 PR；PR 說明寫：改了什麼、驗收結果、需 Andy 決定的事項（任務完成時與報告檔「目前狀態」同步更新）。
- 任務回報：依下節「每輪回報規則」。報告檔另須保存 `scripts/verify.sh` 完整輸出、關鍵數字前後對照（三情境目標價等；未改數字時註明「無變動」並列出現值）。

## 每輪回報規則
Andy 不熟程式，他透過 Claude 聊天端（讀 Gmail 裡的 GitHub 通知）掌握進度，所以每一輪都必須留下書面回報。
1. 任務一開始就開 draft PR（不必等完成），之後每輪的回報都掛在這個 PR 上。
2. 每一輪回覆 Andy 之前，都要做完以下三件事：
   a. 更新 `docs/reports/YYYYMMDD_任務名稱.md`：
      - 檔案最上方是「目前狀態」：完成到哪裡、下一步、需要 Andy 做的事；每輪覆寫。
      - 下方是「輪次紀錄」：依時間追加，寫本輪做了什麼、驗證結果、提交編號。
   b. 在 PR 留一則留言，內容就是本輪紀錄。第一行固定為：
      `[Nebius 回報] 任務名稱｜第 N 輪｜YYYY-MM-DD`（分段工作單用 `[Nebius 回報] v0.1x｜完成／停止｜YYYY-MM-DD`）
   c. commit 並 push。
3. 寫法：用非工程師看得懂的中文，先寫結論，技術細節放後面。
   PR 留言只貼 verify.sh 的結果摘要（例如「12 項全部通過」），完整輸出放在報告檔。
4. 每則回報的最後必須有兩段：
   - 「下一步」
   - 「需要 Andy 做的事」：逐項列出、可以直接照做；沒有就寫「無」。
5. 計畫階段（還沒改任何檔案）也算一輪，同樣要回報。
6. PR 留言必須完整列出「需要 Andy 決定的事項」：逐項寫出問題與你的建議，不可只寫「見回報檔」。Andy 的聊天端透過 Gmail 讀留言，看不到 PR 分支上的檔案。

## 勿改（踩過的坑）
1. 模板函式庫在 segA、segB 的插入點前是**未結束的 var 宣告鏈**：注入的第一句必須是 `COMPANY_DATA = {…};`，segB 第一行必須是 `PERIOD_LABELS = …,`，兩者都**不可加 var**。
2. 模板已宣告 `CO`，且其片段仍以 `Qk` 引用預設值：資料變數名為 `COMPANY_DATA`，並保留 `Qk = DEFAULTS` 別名。
3. 改名一律用 `rename_ids.py`（詞法分析，只改程式碼、不改畫面文字，含 `...name` 展開語法）；新名稱須先確認不存在於模板函式庫。
4. Excel `<sheetPr>` 子元素順序必須是 tabColor → outlinePr → pageSetUpPr，否則 Excel 回報「檔案損毀」；LibreOffice 與 openpyxl 不會發現，所以 `verify_ooxml.py` 必須通過。
5. 總結頁列印依賴 segD 的 `sumQ-np`、`sumQ-main` class。
6. Excel 情境區間是模擬運算表（「輸入與假設」H 區）：Excel 規定輸入格（情境選擇）與運算表須在同一工作表；LibreOffice 重算後會改寫成 `TABLE()`，必須接著執行 `fix_datatable.py`（`verify_ooxml.py` 會檢查）。

7. Excel 同一工作表、同一區段內的 A 欄列名稱不可重複：`xl_diff.py --by-label`（`verify.sh --vs-dist` 使用）以「區段＋列名稱」配對新增列造成的位移，重複時直接報錯、不自行配對。
8. 同業 Comps 與評等門檻只放在 `company.json`（`peers`、`methodology.rating`）；segB／segC／build_xlsx 不再寫任何同業數字或門檻數字。含門檻的畫面文字以門檻值組字（HTML `pctQ`／`multTxt`，Excel `_PC`／`_MT`），cmp31 以標籤為鍵的列（例如「賣出門檻價（現價 × (1 − 15%)）」）標籤也由門檻值產生。

## 待辦（Nebius；依序，每項一張工作單、一個 PR）
CRWV 時期的待辦（環境移植、區間、設定集中、季度層、期間滾動已完成；5a 精簡、5b 只讀檢視器、折舊修正、短名稱改名未做）不在 Nebius 範圍內：依 repo 根目錄 README，CRWV 未完成的待辦（5a 精簡、5b 只讀檢視器、折舊修正、機率加權目標價、GPU 批次與續約價格衰退）Nebius 也先不做。
1. **Q3 2026 季度更新**（6-K 預計 2026-11；步驟見交接檔「待辦與季度更新」）：填 `quarterly.actuals`、滾動評價日至 2026-09-30、更新 `ytdActual`、`latestQuarter`、`asOf` 全部欄位；期初可計費 MW 依新一季營收 × 4 重新校準（`defaults.billableOpen` 隨之更新）；升版 v0.2。
2. 若 Q3 揭露季末 active／connected MW：改以實際 MW 取代內插與校準，重估「實現單價 vs 正向推導單價」差距與爬坡係數。
3. 市場共識與同業 Comps 更新（共識檔放 `data/`，改 `meta.consensusFile`；同業市值與淨負債仍為 2026-09-21）。
4. 另估 Nebius 的 WACC 與 EV/EBITDA 倍數（目前沿用 CRWV 模板 11%／6x；Nebius 淨現金、槓桿較低）。
5. ClickHouse 持股比例一手揭露後更新 `valuation.holdings`。
6. CRWV 未完成待辦（範圍外，待 Andy 另開）：續約價格衰退（預付款優勢可能高估）、GPU 批次、折舊修正、機率加權目標價、5a 精簡、5b 只讀檢視器。

## 已決定事項（不再列為待決）
前三條是工作原則，適用所有後續工作與回報（2026-09-26，Andy 決定）。
1. **可移植**：本 repo 的主要目的是可用於其他個股的框架（下一步 Nebius、Oracle、OpenAI），不只是 CoreWeave 模型（本資料夾即其 Nebius 版）。公司特有的內容（數字、指標清單、拆分依據、指引項目、說明文字）一律放 `company.json`，程式只寫通用邏輯；公司沒有某類資料（例如沒有 MW、沒有指引或共識）時，對應欄位顯示「不適用」，不可出錯。每輪回報加一段「可移植性」：這次新增了哪些公司特有內容、放在哪裡、換公司時要改什麼；只適用 CoreWeave、無法設定的部分列為待決事項。
2. **差異要能解釋**：與共識或公司指引不同沒有問題，但每一個差異都必須能解釋。差距超過 5% 的項目附「差異原因」並標明類型：觀點（刻意的不同假設，寫出依據）／拆法（季度分配方法造成，年度或 2H 合計不受影響）／已知限制（模型已知的偏差）。類型與原因放 `company.json` 或依條件產生，不寫死；無法歸類的差異列為待決事項。
3. **精簡**：讀者是人，吸收量有限；呈現的內容必須是核心、關鍵、必要的，在不減損模型能力的前提下盡量精簡。主畫面（一頁摘要、各分頁第一屏）只放結論與關鍵數字；方法、推導、明細放附錄或次層（預設收合）。新增畫面內容前先問「讀者少了它會不會做錯判斷」，不會就放次層或不放。說明以一句為原則，數字優先於文字。回報與 PR 留言同樣適用：結論先行，Andy 要做的事放最前面。
4. LibreOffice Calc 維持在工作階段內補裝（`apt-get install -y --no-install-recommends libreoffice-calc-nogui`），不改由環境預裝。
5. playwright 固定 1.56.0；只在雲端環境更新預裝 Chromium 時才調整，現在不需處理。
6. 技術容忍值（現價下限 0.01、股權需求 ≤0.01bn 視為不需新股）不是評等口徑，不移入 company.json（v4.2）。
7. 同業名稱與備註暫時保留 HTML、Excel 兩種寫法（`peers.list` 的 `labelHtml／labelXlsx`、`noteHtml／noteXlsx`），統一另開任務（v4.2）。
8. **Excel 為唯一計算引擎**（2026-09-26，Andy 決定）：HTML 改為只讀檢視器，不再自帶計算邏輯（5b 完成）。新增計算一律做在 Excel 公式，或建置時的 Python（結果寫入 Excel）；不再新增 JS 計算。具名範圍是 Excel → JSON → HTML 的介面：凡 HTML 要顯示的輸入與輸出都加名稱（不限數量），另建一張名稱清單表。
9. 待辦 5 的口徑（2026-09-26，Andy 決定）：精簡預設「下移到次層（收合）」，只有重複內容才刪除；分頁合併先提候選清單（一次最多 3 組）逐組核可；公司特有長文字放 `company.json` 的 `texts` 區；非區間型指引——年增率型有去年同季實際數才換算金額算差距、否則只列，利潤率區間以百分點比較（2pt），文字型預設只列不算；換公司測試用暫存的虛構 company.json（不入 repo、不用任何真實數字）。第 2 輪：HTML 即時輸入依盤點建議取捨（保留三情境 × 四錨定預算組合與敏感度表，其餘移除）；分頁合併三組全部同意（運營_站點＋資產負債_租賃承諾、評價_可比公司併入評價_DCF與目標價、來源＋檢查_版本紀錄）；精簡重點同意，但「客戶集中」保留一句在一頁摘要或檢查頁的第一屏，其餘觀察移到次層；版本號在合併時取下一號（例如 Q3 更新先合併為 v4.5，則 5a 為 v4.6）。第 3、4 輪：敏感度採折衷方案 C（評價的 WACC × g、錨定 × 倍數為活的；敏感性排名與資金面兩張矩陣為建置時快照，附「快照已過期」檢查）；模型期長度固定為「首期＋4 個完整財年」（每年最後一季後延長一年，由 Andy 補新一年假設）；目標價時點改為「評價日＋12 個月」（理由：賣出門檻與共識目標價都是 12 個月口徑；分兩段驗收，先現行口徑數字不變，再切換並列三情境點位、區間、評等、距賣出門檻前後對照；內插口徑有疑義時列選項由 Andy 決定，不得自行退回舊口徑）；Comps 先做 LTM／NTM 機制，缺資料顯示「不適用」、現有數字標「未日曆化」；Q3 採路徑 2（5a-1 先合併，季度層財報當天更新，年度首期 10-Q 後滾動）；滾動後至 5b 前，一頁摘要與各分頁第一屏的日期與期間文字須自動產生、不得出現過期的「6/30」，其餘次層文字由 5b 取代；「客戶集中」一句放檢查頁第一屏。
10. **每季滾動**（2026-09-26，Andy 決定）：每次財報後，首期改為「年初至今實際＋剩餘季度模型」，目標價隨之更新並升版。期間滾動的設計基準是 CoreWeave、Nebius（財年 12 月）、Oracle（財年結束於 5/31，FY27＝2026/6–2027/5，首期「Q1 實際＋Q2–Q4 模型」）、OpenAI（未上市、日曆年、首期可能只有全年模型）。要求：(1) company.json 以「財年結束月份」與「最新已公布季度」自動推算期間，首期長度支援 0.25／0.5／0.75／1 年，期間與季度標籤依財年產生，`actual1H` 改為通用的「年初至今實際」；(2) DCF 折現與 FY27 末錨定改用實際日期（評價日到各期期末），不再假設日曆年；(3) Comps 改用 LTM／NTM 或日曆化，不同財年不直接並列；(4) 共識檔年度標籤依公司財年對應。
11. **HTML 單一檔案、離線開啟**（2026-09-26，Andy 決定；硬性要求，公司網路會阻擋許多外部資源）：(1) `dist/` 的 HTML 為單一檔案：所有 JS、CSS、字型、圖片、資料（含 5b 由 Excel 匯出的 JSON）都在建置時內嵌；不得依賴旁邊的任何檔案（file:// 開啟時瀏覽器禁止讀取其他本機檔案）。(2) 不得有任何網路請求：不用 CDN、Google Fonts、外部圖片、fetch／XHR、ES module 外部 import；字型只用系統字型或內嵌字型。(3) `verify.sh` 的「離線開啟檢查」（`scripts/check_offline.py`）：把 HTML 單獨複製到空資料夾，以 Playwright 在阻斷網路的狀態下用 file:// 開啟，確認 (a) 網路請求數＝0、(b) console 無錯誤、(c) 一頁摘要與主要分頁的關鍵數字（取自同版 Excel）正常顯示；任一項失敗即不通過。(4) 5b 驗收納入以上三項。
12. **評價口徑、目標價變動拆解與升版驗收**（2026-09-26，Andy 決定，5a-1 第 1、2 輪）：(1) DCF 腿以 WACC 推到目標價時點（0 截斷與選擇權皆同），兩條腿同一時點；(2) 評價日維持「最新已申報季末」（與賣方以股價日起算 12 個月的差異寫明於 README 與交接檔）；(3) 期間滾動或任何升版後的目標價變動須拆成四項並列於版本紀錄：(a) 時間推移（評價日後移造成的折現變化）、(b) 實際數更新、(c) 假設變更、(d) 方法變更（評價口徑、倍數、權重、錨定年度等），(a)＋(b)＋(c)＋(d)＝總變動（v4.5 的變動全部歸入 (d)）；(4) 升版驗收對前一版成品跑 --vs-dist，以 --expect 清單列出預期變動的格，其餘 0 差異。(5) **(a) 時間推移為定義值**（2026-09-26，Andy 決定；PR #10）：純時間價值，以 WACC 累積＝前一版兩條腿（0 截斷前）×（1＋WACC）^（評價日後移月數 ÷ 12），再做 0 截斷與加權；WACC 讀 Excel 的 WACC 輸入格、月數讀 `calendar_q.py` 推算結果，不得寫死；不可用「只滾動日曆、不改其他資料」重算（會刪掉該季流入、保留流出，淨負債與股數停在舊評價日）。(b) 含該季實際與模型的差異，以及融資時點殘差（淨負債以債務成本累積、該季發股）。只拆解 0 截斷口徑；選擇權模式的 DCF 腿只列前後值。工具 `scripts/attrib.py`。(6) **滾動檢查**：首期一次性金額與期初餘額（清單 `calendar_q.py` → `ROLL_FIELDS`）以 `company.json` → `asOf` 標記所屬已申報季度，滾動後未逐項更新即建置失敗；新增此類輸入時同步加入清單。

## 溝通
以繁體中文回報；專業、客觀、以事實為準。遇到需要判斷口徑（評價方法、預設值、評等門檻）的事項，停下來列出選項與影響，由 Andy 決定。
