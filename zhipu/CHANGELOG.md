# CHANGELOG（智譜收支模型）

## Z0（2026-10-08，chat 端）
- 建立 `zhipu/`：骨架（builder／engine／tools／tests 基礎自 OpenAI v0.6 經 Anthropic A0 複製，名稱待 Z2 改寫）、規格 r0（D1–D26）、共同規則、Z1–Z5 工作單、進度檔、交接檔。

## Z1（2026-10-08，chat 端＋三個研究代理）
- `data/zhipu_src.yaml`：SRC_ZP_001–604（財報 240、價格用量算力 251、融資市值 113），找不到 41 項；規格 r1（第 6 節：D2r、D8r、D9r、D10r、D11r、D14r、D16r、D21r、D22r、D24r）。

## v0.1-Z2（2026-10-08，建置代理）— Excel `model/20261008_Zhipu_v0.1.xlsx`
- commit：e425dce（步驟 1–3：builder 骨架、TK_Link 讀表 7 列、OAI_Link、Inputs 草稿）、6a177f6（步驟 4–6：Demand、Revenue、Checks）、569d362（步驟 7：parity 20 情境、test_builder、`zhipu-parity.yml`）、本節報告 commit。
- 新頁：README、SRC_ZP（604 列）、TK_Link（63 名＋7 讀表列）、OAI_Link（OpenAI v0.6 命題輸出 14 列）、Inputs（73 列）、Demand、Revenue、Checks（C01–C42）。
- 營收（RMB 億）：2025 7.24（校準差距 0）、2026 37.14（1H 實際 9.54＋2H 27.60）、2027 69.55、2030 226.66；雲端 2H26 年化 40.4 對 MaaS ARR 107.8 → WARN −63%。
- 驗收：pytest 59 passed；CHK_Errors＝0；TK 快照 OK 70。報告 `docs/reports/20261008_v0.1-Z2.md`＋`_對照.xlsx`。
- 工具：engine／tools 名稱改智譜；`check_tk_snapshot.py` 支援讀表列；`make_report_tables.py` 改為產生本模型對照 Excel。

## v0.1-Z3（2026-10-08，建置代理）— Excel `model/20261008_Zhipu_v0.1.xlsx`（Z2 版移 `model/archive/20261008_Zhipu_v0.1-Z2.xlsx`）
- commit：aada6aa（第 0 步：chat 端審查改動 R1 INP_038 1.3→2.24、R2 INP_069 1→0）、34451f8（步驟 1–3：Compute、Cost、Checks）、fc9e969（步驟 4：parity、報告）。
- 新頁：Compute（G01–G130：晶片族 Hopper／H20／國產每 GW 產能與組合、token 換算 GW、算力服務費→供給 GW、η 2025 49.05／1H26 38.35、研發 GW 殘差、容量上限、VR 等值、10 萬國產晶片對照、R4 對帳、敏感度）、Cost（K01–K65：算力成本、供應商持有成本與雲端毛利、本地化交付成本、非算力成本、股權報酬、命題表 RMB 億與 $B、TK 單位成本參考）。
- Inputs ＋46（INP_074–120）；INP_072（容量佔位）退役；Revenue 容量上限列改引用 Compute；Checks C43–C67（C42 退役）。
- 命題：每 VR 等值 GW 差額（含股權報酬）−17,949／−6,310／−2,447／+599／+2,961／+4,859 RMB 億（−266／−94／−36／+9／+44／+72 $B）；2025 全成本對財報差距 0。
- parity 情境 33（＋12）；報告 `docs/reports/20261008_v0.1-Z3.md`＋`_對照.xlsx`；`tools/make_report_tables.py` 加 `--seg z3`。

## v0.1-Z3b（2026-10-08，建置代理；η≫1 處理）— Excel `model/20261008_Zhipu_v0.1.xlsx`（Z3 版移 `model/archive/20261008_Zhipu_v0.1-Z3.xlsx`）
- commit：1a9e3f3（F1–F3）、bf2e702（F5）、本節報告 commit。
- F1：token 換算推論 GW 改依 token 類型（新鮮 prefill／快取命中／decode）× TK `IF_CostPre／Cache／Dec_*` ÷ `IF_HoldEcon` ÷ `IF_Util`；TK_Link ＋9 名；Inputs ＋3（INP_121–123）。F2：η 2025 49.05→12.44、1H26 38.35→9.46，η>1 基準沿用 1H26、收斂至 1 列情境。F3：舊法保留對照列。F4：供給機制列為 chat 端確認預設。F5：CI pytest 加 `-p no:warnings`。
- 命題：每 VR 等值 GW 差額 −17,949／−6,503／−3,439／−1,618／−747／−231 RMB 億（−266／−96／−51／−24／−11／−3.4 $B）；2030 前無轉正年；覆蓋率 2030 0.96。
- parity 35 情境；pytest 88 passed；CHK_Errors＝0；TK 快照 OK 79。報告 `docs/reports/20261008_v0.1-Z3b.md`＋`_對照.xlsx`。

## v0.1（Z4，2026-10-08，建置代理）— Excel `model/20261008_Zhipu_v0.1.xlsx`（Z3b 版移 `model/archive/20261008_Zhipu_v0.1-Z3b.xlsx`）
- commit：1c020e5（Funding、Reverse、Checks C68–C92、parity ＋8、`tools/sensitivity_z4.py`）、3a60b39（Cost 第九節每實體 GW、Funding 第十節美元口徑、除以 0 防護）、本節報告 commit。
- 新頁：Funding（F01–F74：自由現金流、投資活動、已到位融資（IPO 42.07、2026-07 配售 269.58、2026-09 配售 134.59、可換股債券 202.92 RMB 億）、最低現金與外部資金需求、可轉債轉股情境、或有、配售用途對照、回流對照、2025／1H26 現金對帳、美元口徑）、Reverse（X01–X62：管理層目標不存在、共識 3 筆平均、所需倍數、反向資金、市值隱含營收）。
- Inputs ＋12（INP_124–135）；SRC_ZP ＋6（SRC_ZP_605–610，由既有 note 拆出數值）；OAI_Link ＋3 列；Cost ＋第九節（`COST_PropGap_Phys(_USD)`）。
- 命題 2：累計外部資金需求 2026–2030 全為 0；年底現金 624.2／382.7／352.4／330.8／321.2 RMB 億（可轉債轉股 624.2／584.1／553.8／532.2／522.6）；命題 1 不變（2030 −231 RMB 億／VR 等值 GW；每實體 GW −12.96，−$0.19B）。
- 反向：共識 2028 平均 324.7 億＝正向 2.47 倍；市值隱含營收 62.7 億（MiniMax 39.8 倍），正向 2027 達到。
- 驗收：pytest（見報告）；CHK_Errors＝0；TK 快照 OK 79（Tokenomics master 70d859e，CURRENT 仍 v5.26）。報告 `docs/reports/20261008_v0.1.md`＋`_對照.xlsx`＋`_敏感度.json`。

## v0.1 成品（Z5，2026-10-08，建置代理）— Excel 未改（`model/20261008_Zhipu_v0.1.xlsx`）
- commit：9c60ffa（HTML 一頁摘要、dist、test_html、截圖）、e391320（命題與驅動對照表）、本節交接檔／README／報告 commit。
- `dist/20261008_智譜收支模型_v0_1.html`：11 節（命題與兩種讀法＋四個翻轉點、每 GW、營收與共識、算力、現金與外部資金、敏感度、主要風險／與實際觀察的落差、市值對照、與 OpenAI v0.6 並排（美元）、最該審的 5 項預設、資料與版本）；686 個數字皆讀自 LibreOffice 重算後的 Excel（19 個情境），附 Excel 位置；單一檔案、無外部資源、圖表內嵌 SVG。
- `dist/20261008_智譜收支模型_v0_1.xlsx`＝model 複本；`tools/build_html.py` 改寫為智譜（Inputs 欄位 F／G／H、OAI_Link F–K 欄位移、半年欄 K／L、翻轉點取自敏感度 JSON）；`tools/html_scenarios.yaml`；`tests/test_html.py`（6 項）。
- `docs/plan/20261008_Zhipu_命題與驅動對照表.md`；交接檔改為 v0.1 完成狀態（未解問題、v0.2 建議）；根目錄 README 智譜列「v0.1」。

## v0.1 r2（Z5b，2026-10-08，建置代理；獨立查核後修正 V1–V9）— Excel `model/20261008_Zhipu_v0.1.xlsx`（Z4 版移 `model/archive/20261008_Zhipu_v0.1-Z4.xlsx`）
- commit：d10b940（V1–V3 Excel）、8c0ea1e（V4–V9 HTML、報告、成品）、本節 commit（parity 期望值、CHANGELOG、README）。
- V1：2H26 起每 GW 租價＝MAX（觀察租價, 供應商持有成本 ×（1＋最低毛利 INP_136，Assumed 0，0–0.2））；Compute 第十四節 G161–G171；Cost K15／K16 改名「供應商推算毛利（率）」；Checks C93。
- V2：TK_Link ＋IF_OpexGW、IF_DeprLifeIT；自有算力投產後營運費用計入算力成本（Compute G172–G173、Cost K74）；攤提口徑並列（Compute G174–G175、Cost K75–K79：`COST_FullAmort`、`COST_PropGap_VR_Amort`、`COST_Coverage_Amort`）。
- V3：Reverse X42 淨現金扣 2026-07 配售款已動用（Funding F52）。
- 命題 1：2030 每 VR 等值 GW 差額 −231 → −2,275 RMB 億（−$33.8B），覆蓋率 0.96 → 0.72，2028 年後不再收斂；命題 2：累計外部資金需求 2030 0 → 18.9 億（首次缺口年 2030）。翻轉點只剩研發占比 ≤0.572。
- V4–V9：翻轉點重算（`tools/sensitivity_r2.yaml` → `docs/reports/20261008_v0.1-r2_敏感度.json`）；HTML 措辭（命題 2 可信度低、ARR 與支出落差對稱）、晶片同口徑、OpenAI 並排加註；`docs/reports/20261008_v0.1.md` 標頭與口徑欄；成品重出；報告 `docs/reports/20261008_v0.1-r2_查核修正.md`。
- parity：44 情境（＋ap_min_gm_02、aq_owned_opex_amort）；workbook_expectations 更新（公式格 4,810、具名範圍 1,111、Inputs 135、TK 74）。

## v0.1 r2（招股章程 P1–P5＋Z5b V1–V9 重驗，2026-10-08，建置代理）— Excel `model/20261008_Zhipu_v0.1.xlsx`（Z5b 版移 `model/archive/20261008_Zhipu_v0.1-Z5b.xlsx`）
- commit：3928712（P1）、1f83c31（P2）、be9ef59（P3–P4）、e1466f4（V4–V7＋P5：翻轉點、HTML、成品、截圖）、本節 commit（報告、對照 Excel、交接檔、對照表、進度檔）；另 ea6a2ca 合併遠端（chat 端已合併 main）。
- SRC_ZP 寫入招股章程 444 列（SRC_ZP_611–1054；共 1,054 列）；Checks C01／C02 範圍改 A5:A2000。
- P1（取代規格 D8r）：
  - 總算力費＝營業成本計算服務費（API 銷售成本 × INP_102）＋研發開支 × INP_137（0.718）；供給 GW＝總算力費 ÷ 每 GW 價格；推論 GW＝token 換算 GW ÷ η，η 基準 1（INP_138），校準期研發殘差 < 訓練最低占比時上調並 WARN（C94）；研發 GW＝殘差。
  - 舊法以 INP_139 開關與 Compute 第十五節保留對照。
  - η 1H26 9.46→3.91、2025 12.44→1.13；研發占比起點 0.70→0.30；2030 供給 0.829→0.856 GW。
- P2：客戶 A 改 SRC_ZP_785／786；資本開支權責 vs 現金對照（Funding 第十一節）；媒體口徑 note 更正；一手列取代。
- P3：租賃（使用權資產、租賃負債 5.45 億 `FND_LeaseLiab`、到期、4.75%、售後租回）只列；租賃付款不另加。
- P4：政府補助沖減開支，只列；D20 說明改寫。
- P5：HTML ⑦ 並列招股章程「無計劃自建資料中心」與 2026-07「已落地 1GW」。
- 命題 1：2030 每 VR 等值 GW 差額 −2,275 → −2,394 RMB 億（−$35.5B）、覆蓋率 0.72 → 0.71；每實體 GW −$1.99B。
- 命題 2：累計外部資金需求 2030 18.9 → 43.6 億（首次缺口年 2030；轉股 0；1 GW 915.8、2026）。
- 翻轉點：命題 1 只剩研發項下算力費占比 ≤0.316；命題 2：該占比 ≤0.661、INP_102 ≤0.805、最低毛利 ≤−4.8%。
- parity：46 情境（＋ar_eta_old_method、as_rd_fee_share_low）；aa_rd_fee_too_high 改用 INP_137；workbook_expectations：公式格 6,214、具名範圍 1,559、SRC_ZP 1,054、Inputs 139。
- HTML 782 個數字（基準＋22 情境）；`tests/test_html.py` 定性主張更新。
- 報告 `docs/reports/20261008_v0.1-r2_查核修正.md`＋`_對照.xlsx`＋`_r2_敏感度.json`。
