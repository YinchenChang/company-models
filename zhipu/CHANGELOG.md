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
