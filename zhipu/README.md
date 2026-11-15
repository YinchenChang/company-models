# 智譜（智譜華章，02513.HK）個股模型

智譜每 VR 等值 GW 的年營收能否覆蓋每 GW 年全成本；若不能，缺口由誰、以什麼條件融資（FY2025–FY2030）。架構比照 `openai/`（v0.6），算力物理取自 Tokenomics；金額單位人民幣億元，命題另列美元口徑與 OpenAI 並排。

- 現行版本：**v0.1.6**（`model/20261115_Zhipu_v0.1.6.xlsx`；成品 `dist/20261115_智譜個股模型_v0_1_6.html`〔版面 v1.3〕、`dist/20261115_智譜收支模型_v0_1_6.html`〔一頁摘要〕、`dist/20261115_智譜個股模型_v0_1_6.xlsx`；見下方 v0.1.6 節：估值頁、目標價與評等）。前版 v0.1.5（`model/archive/20261115_Zhipu_v0.1.5.xlsx`；成品 `dist/_待刪除/`）＝版面 v1.3＋三層協調；v0.1.4＝只換現價 2026-11-05 收盤 HK$476（SRC_ZP_1055）。
- 官方報告分部（部署方式）：本地化部署（②企業級智能體＋③企業級通用大模型＋④技術服務及其他）、雲端部署（①開放平台及 API）。宣告檔 `segments.json`；對帳在 Revenue 第七節（FY2025 對年報、1H2026 對中期公告子項加總，容差 Inputs `seg_recon_tol`），Checks C100–C103；`python3 tools/segment_check/check_segments.py zhipu`（repo 根目錄執行）須為「符合」。
- 規格：`docs/plan/20261008_Zhipu_v0.1_規格.md`
- 交接：`docs/handoff/Zhipu_handoff.md`；進度：`docs/reports/20261008_zhipu_進度.md`（v0.1.2：`docs/reports/20261030_v0.1.2_進度.md`）
- 工作規範：`CLAUDE.md`


## v0.1.6（2026-11-15）：估值頁（目標價、評等）與反向 DCF；成品命名對齊 MiniMax

**依據：** Andy 2026-11-15 決定智譜與 MiniMax 都給目標價與評等（取代 v0.1.5「本模型未建目標價」與 DCF 標準規格「AI 收支模型不加 DCF 腿」的例外）。樣板：Anthropic v0.2.0 估值頁（`anthropic/builder/a11.py`）、反向 DCF（`a12.py`）；`docs/plan/個股DCF標準規格.md`；個股模型原則第 24、36、56、63 條；版面規格 v1.3 A–I 節。

**做法：** `builder/z6.py`（參照）＋`builder/valcore.py`（Val、RevDCF 排版與公式；與 `minimax/valcore.py` 同一份程式）；新增 SRC_ZP_1056–1064（同業 MiniMax 2026-11-05 收盤、股數、淨現金組成、共識營收）、Inputs 26 列（`val_*`）。`tools/val_snapshots.py scen|sens <xlsx>` 改 Inputs、LibreOffice 重算，產生 `data/val_scenarios.yaml`（情境 A／C）與 `data/val_sens.yaml`（龍捲風），再以 builder 重建寫入 Val。兩頁不被任何既有計算頁或 Checks 引用。
- **成品命名（主控 2026-11-15 指示，對齊 MiniMax）：** 版面 v1.3 頁 `dist/20261115_智譜個股模型_v0_1_6.html`（v0.1.5 稱「個股頁」）；一頁摘要改 `dist/20261115_智譜收支模型_v0_1_6.html`；xlsx 仍為 `dist/20261115_智譜個股模型_v0_1_6.xlsx`。AI 投資地圖 `pick_html()` 取檔名含「個股模型」者 → 嵌 v1.3 頁。
- 個股頁估值章拿掉 `data-novaluation`，照 A–I 節：G＝四個目標價與現價圖在估值章最前；H＝加權目標價相對基準低案／高案兩張主圖在敏感度章最前；新增反向 DCF 章；原反向檢驗（市值隱含營收）與收支關卡敏感度收進摺疊。

**結果（2026-11-05 收盤 HK$476）：** DCF 114.99、Comparable 198.88（三情境加權；199.74 為情境 B）、**加權目標價 HK$156.94**、市場共識 1,567.3（Investing.com 24 位，2026-10-03）；上檔空間 −67.0%，**估值評等「強力放空」**；收支關卡 FY2026–FY2030 自由現金流（Funding F09，與 AI 投資地圖同口徑）合計 −281.2 億元（不通過），投資建議＝強力放空；距最近邊界 17.0pt（升一級需 HK$238）。
- DCF：現金流沿用模型既有路徑（Funding F01 營收、F06 差額〔含股權報酬〕＝EBIT、F10 非算力資本支出、F12 併購；2H26＝Funding L 欄）。基準 2H26–FY2030 各年 FCFF 皆負、延伸期維持 FY2030 利益率（−36%），三情境 EV 皆 ≤0，DCF 觸及淨現金下限（評價日 pro forma 淨現金 417.9 億元 ÷ 4.876 億股 × (1＋ke)^1.35 ÷ HKD/CNY）。
- Comparable：同業＝MiniMax，EV ÷ 未來 12 個月共識營收 5.65 倍（2026-11-05 同日市值）；智譜自身 9.74 倍（同業的 1.73 倍）。三情境 NTM 營收 84.9／74.2／60.9 億元 → Comparable 214／200／182 港元。
- WACC 10.49%（ke 11.10%）；反向 DCF：支持現價需 EV 1,312 億元（模型 −1,032）；只改終值期利益率需 94.0%；延伸期利益率加碼需 +115pt（2033 年約 79%）；同業倍數法下現價隱含同業的 3.76 倍（所需 NTM 營收 279 億元 vs 模型 74）。
- 敏感度（相對基準）：DCF 權重 25%／75% → +13.4%／−13.4%；依 FY2030 利益率調整倍數 → −11.6%；延伸期利益率 +50pt → +4.0%；任務數高／低端 → +2.3%／−2.8%；WACC、g、稅率、營運資金、算力價格、研發算力占比為 0（DCF 在淨現金下限）。

**判斷類參數（依樣板慣例預設，列為待 Andy 確認）：**
| 參數 | 預設 | 理由 | 替代 |
|---|---|---|---|
| rf | 4.86%（美債 10 年 CMT，2026-11-06） | 港元聯繫匯率，沿用 Anthropic 同一來源；現金流為人民幣，未另換人民幣無風險利率 | 港府或中國 10 年債 |
| β | 1.248（Anthropic 同業 5 年 β 中位數 1.37 經 Blume） | 智譜 2026-01 上市，自身 β 期間不足 | 1.0／1.6（Sens） |
| ERP | 5%，未加中國國家風險溢酬 | 個股DCF標準規格預設 | 6%（Sens） |
| 稅率 | 25%（扣評價日起自身虧損） | 中國企業所得稅 | 15%（高新技術企業；Sens） |
| g、RONIC | 3%、RONIC＝WACC | 同規格 | 2%／4% |
| 延伸期 | 營收年增收斂到 10%、利益率維持 FY2030 | Anthropic 樣板 | 利益率加碼（Sens +50pt） |
| 同業 | 只有 MiniMax（2026-11-05 同日市值，與 MiniMax 模型 v0.2.3 同一組數字） | 其餘中國與海外上市 AI 模型／應用公司（商湯 0020.HK、科大訊飛、第四範式、SoundHound、C3.ai、Palantir 等）本輪查不到 11-05 同日可查證市值（網路搜尋與行情頁快取皆過期），未納入、不編造 | 補資料後加入同業 |
| 倍數口徑 | EV ÷ NTM 共識營收（FY2026 剩餘 56/365 權重）；不做相對倍數調整 | 同業同為虧損 AI 模型公司 | 依 FY2030 利益率調整（Sens） |
| 權重 | DCF／Comparable 50／50；情境 25／50／25 | 個股模型原則、DCF 標準規格 | 25／75（Sens） |
| 評價日與股數 | 2026-06-30 pro forma（計入 07、09 配售與可換股債券，扣銀行借款、可換股債券本金〔換股價 892.5 價外〕、租賃負債〔FY2025〕）；股數 4.876 億（2026-09 配售後） | 中期報告為最近資產負債表 | — |
| 情境 A／C | 2H26 API 任務數成長率取高／低端（INP_038） | Anthropic 樣板以 API 任務數定情境 | 研發算力占比、算力價格 |

**資料缺口：** 受限制股份有揭露（SRC_ZP_876，招股章程附註 29(b)：1H25 尚未行使 6,620 千份〔拆股前〕，1 拆 10 後約 6,620 萬股，行使價 0.10–2.00 元），未計入評價股數；若全部為增量約占 13.6%，加權目標價約降至 138–140 港元，評等不變（是否已含在已發行股數內未查證）；1H26 租賃負債未揭露（沿用 FY2025）；共識營收為 2026-10-06 彙整、共識目標價為 2026-10-03，早於現價日（其間股價大跌）。

**驗證：** builder 重建（restore matched 3,585、unmatched 0；e6 0；oai_refs 0）；`tools/xlsx_diff.py` 對 v0.1.5：既有格只差 README!A1／B7／B10 與 Checks!A1 版本與列數字串、Checks C5／D5、C9／D9 的 SRC／Inputs 列數（1055→1064、140→166），隱藏 `_Defaults` 因新增列位移，其餘皆為新增格；test_builder 14 通過；test_html 7 通過（含 v1.3 頁全部 data-ref 以 engine 比對）；parity：結構、具名範圍（含對 LibreOffice）、效能，與 base、n_task_growth_high、w_dc1gw_on、af_cb_convert、k2_rev_total_mismatch 五個情境通過（其餘情境以 CI 為準；Val 不被 Checks 引用，`test_scenario_expected_results` 不受影響）；`check_layout_v13 --xlsx` 全過；`check_rating_gate --strict` 符合；`tools/baseline_coordination/check.py`（全樹）0 項不符。重建輸入：Tokenomics v19.0（1916a06）、OpenAI v0.6.2、`--base model/archive/20261115_Zhipu_v0.1.5.xlsx --date 2026-10-22`。

**獨立核對文字修正（同版 v0.1.6 重建，2026-11-15，PR #588；數值格不變）：**
- 「Comparable 199.74」改為三情境加權 198.88（199.74 為情境 B）。
- Val!K40 與本節資料缺口：原寫「選擇權、受限制股份未揭露」與 SRC_ZP_876 矛盾，改為「有揭露、未計入評價股數」；若全部為增量約 13.6%，加權目標價約降至 138–140 港元，評等不變（是否已含在已發行股數內未查證）。v1.3 頁估值章可見處加「稀釋未計」說明。
- v1.3 頁估值章加註：β 借自 Anthropic 估值頁美國軟體同業（Palantir、CrowdStrike、Cloudflare、Snowflake、Datadog）5 年 β 中位數；DCF 淨現金下限未扣評價日後燒錢（若改用模型目標時點淨現金內插，加權目標價約 142，評等不變）；智譜與 MiniMax 互為唯一同業，兩家 Comparable 結論本質上是同一組相對價差。
- 估值章摺疊「原收支模型的反向檢驗」加口徑說明：淨現金／EV 417.9／1,576（Val，6-30 pro forma）vs 329／1,665（Reverse!D54，另扣 7 月配售款至 8-31 已動用部分）；同業倍數 39.8 倍（MiniMax 10-07 盤中、年化營收）vs 正文 5.65 倍（11-05 收盤、NTM 共識營收）。
- 合併 origin/main 後，Coord 的 Wattonomics 對照快照隨 main 由 v0.3.2 換為 v0.3.3（只改 Coord!C15、L15 說明文字；本模型不讀 Wattonomics）。
- xlsx_diff 對修正前：只差 Val!K40、Coord!C15、L15 三格文字；一頁摘要 HTML 逐位元組不變。

## v0.1.5（2026-11-15）：個股 HTML 版面 v1.3＋三層協調（讀取層級、C0 公司實績）

**依據：** `docs/workorders/20261114_第二批13家_v1.3與三層協調.md`；`docs/plan/三層基準協調規則.md` v1.4.2 C0 第 2 點；`docs/plan/個股HTML版面規格.md` v1.3；`tools/coord/README.md`。樣板：Anthropic v0.2.2、OpenAI v0.6.5（PR #524）。

**做法：** `builder/z5.py` 新增 PL（標準損益表）、Coord（讀取層級、C0、判定摘要、參考）兩頁，只列示；`tools/build_html_v13.py` 產生版面 v1.3 個股頁（命題 → 估值〔不適用〕→ 敏感度〔各情境對 FY2030 覆蓋率〕→ 損益 → 驅動因子〔分部營收推導〕→ 算力 → 資金 → 與 OpenAI 並排 → 對帳〔C0 表〕→ 資料狀態），數字與一頁摘要同一套取值（LibreOffice 重算、情境重算）與 engine 比對。
- **估值：本模型未建目標價與評等**（Reverse 只作反向對照：市值隱含營收）。智譜已上市、有市價，但本模型是收支模型，未建 DCF／同業估值；依規格 J 節 (b)（2026-11-15 擴充：已上市但本模型未建目標價），估值章標「估值：不適用（本模型未建目標價）」並帶 `data-novaluation="1"`，列現價、市值、EV 與反向檢驗（市值隱含營收），附一張反向檢驗圖（模型營收 vs 市值隱含年營收與分析師共識；數字取自 xlsx）。分析師共識只到 FY2028，FY2029–FY2030 為沿用值（FY2028 共識平均 ×（1＋反向模式年增率假設），基準年增率 0，故三年同值 324.7）。若要給目標價需另建估值頁（判斷類，待 Andy 決定）。

**實績與出處（C0）：**
| 項目 | 期間 | 公司實績 | 出處 | 層級隱含 | 判定 |
|---|---|---|---|---|---|
| 營收總額 | 2026-01～06（主）；季營收未揭露（第二讀數無） | 9.5389 億元 | 2026 中期業績公告（2026-08-31，港交所 2026083101539），SRC_ZP_177 | 無（Wattonomics 無智譜玩家） | 無可比實績（層級無此玩家） |
| 資本開支 | 1H2026 | 3.906 億元（主要為收購北京紅鑽） | 同上，SRC_ZP_209 | 無 | 無可比實績（層級無此玩家） |
| 在役／簽約算力 GW | — | 公司未揭露（只有國產晶片十萬張級、1 GW 資料中心等說法） | SRC_ZP_452、458（Interested-party） | 無 | 無可比實績（公司未揭露、層級無此玩家） |
| 應付＋預付算力服務費 | 2026-06-30 | 26.71 億元（14.27＋12.44） | 中期公告資產負債表，SRC_ZP_218、219 | 無 | 無可比實績（層級無此玩家） |

**判定結果：** 可判定 0、對帳中 0、無可比實績 4；無同源列（Wattonomics 快照全文無智譜，層級不讀本模型）。讀取層級：Tokenomics 快照 v19.0（現行 v22.0，本輪不換）。

**驗證：** builder 重建（restore matched 3,585、unmatched 0；e6 0）；`tools/xlsx_diff.py` 對 v0.1.4：只差 README A1、Checks A1 版本字串與新增格（README「v0.1.5 變動」列、PL、Coord）；test_builder 14 通過；parity（結構、具名範圍、base 與兩個預期 ERR 情境）通過，完整情境以 CI 為準；test_html（含個股頁數字比對）；`check_layout_v13 --xlsx`：本分支版（J 節 (b) `data-novaluation` 豁免）全過；`tools/baseline_coordination/check.py` 通過。

**修正（同版 v0.1.5 重建，2026-11-15）：**
- PL 核對列「營業利益（預測）− Cost 差額 − 股權報酬」原為 `ROUND(…, 9)`；pycel 的 ROUND 在 1 GW 資料中心情境（`w_dc1gw_on`、`aq_owned_opex_amort`）遇 numpy 浮點數無法轉 Decimal（InvalidOperation），parity 失敗 6 項。改為 `IF(ABS(x)<1E-9, 0, x)`（同義：誤差小於 1E-9 視為 0）；其他核對列（稅前 − 自由現金流、FY2025 校準）本來就不含 ROUND，全簿無其他 ROUND。數值逐格不變（xlsx_diff 對修正前只差 Coord!L15 說明文字，見下）。
- 規格 J 節擴為「無目標價的模型」：(a) 未上市 `data-unlisted`（OpenAI）；(b) 已上市但本模型未建目標價 `data-novaluation`（智譜、MiniMax）。`tools/company_html/check_layout_v13.py` 兩旗標同組豁免（不要求反向 DCF 章、略過 G、H），(b) 估值章仍須至少一張圖。智譜個股頁估值章改帶 `data-novaluation`、補反向檢驗圖、橫幅改「本模型未建目標價與評等」。
- Coord「Wattonomics」列說明補：Wattonomics Decisions D43（GL01-37，Andy 2026-10-31 已定）已將智譜移出 Gap（資金缺口收支只算 11 家代表玩家，不含中國）。
- 驗證：完整 parity（tests/parity，47 情境）全過——情境一致性 47／47、快取等價 47／47、結構與效能測試通過；`test_scenarios_actually_change_outputs`、`test_scenario_expected_results` 在本機單一行程記憶體不足（7 GB 上限，pycel 引擎未即時回收），改以同一斷言分段執行（每情境後回收記憶體）全過。test_builder、test_html 通過；`tools/baseline_coordination/check.py` 0 項不符。
- 檢查：main 版 `check_layout_v13` 對本頁不過（main 版只認 `data-unlisted`；本分支合併後即為新版）；新版全過。repo 內其他既有 HTML（77 份）main 版與新版輸出逐字相同。
- **同業倍數日期加註（2026-11-15，PR #573 合併後；只加註，數值不變）：** 個股頁估值章「同業（MiniMax）企業價值 ÷ 年化營收」列與一頁摘要同業倍數說明，註明 MiniMax 市值取 2026-10-07 盤中、與智譜股價 2026-11-05 收盤不同日。只改 `tools/build_html_v13.py`、`tools/build_html.py` 文字，同版 v0.1.5 重建 HTML（xlsx 位元組相同）；test_html、test_builder、parity base 情境通過；`check_layout_v13 --xlsx` 全過。
