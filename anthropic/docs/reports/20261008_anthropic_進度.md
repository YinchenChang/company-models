# Anthropic v0.1 分段建置進度（接手用；A1–A5 共用）

## 目前狀態（每次 push 前覆寫）
- 已完成：A0（`aa925dd`）；A1（`8d955cf`）；A2（`fd5db5b`＋報告 `e903023`）；A3（`c9d2478`＋報告 `a1b7b8c`）；A4＝v0.1（`07565be`、`7842e3f`）；**A5 成品完成**：步驟 1 產生器／測試／dist／截圖（`5d08b70`）、步驟 2 對照表（`210faf0`）、步驟 3 交接檔與 README（`82e88e7`）、步驟 4 成品報告、CHANGELOG、進度檔（見 git log「A5 步驟 4」）。報告 `docs/reports/20261008_v0.1_成品.md`。
- 下一步：chat 端審查 PR #22 並合併（代理不合併）；Andy 審查 v0.1 預設彙總（HTML ⑧、完成報告 ④）；Drive 放置 dist 兩檔。v0.2 候選見交接檔第 6 節（新增合約機制最重要）。
- 未解問題：(1) GitHub Actions 因帳號付款／用量上限未啟動（本地 pytest 全過）；(2) 基準 CHK_Warnings＝1（C89：AMD 模型 77.8 vs 招股書 20，待決定以金額或 GW 為準）；(3) D22：2026 非算力成本隱含低估約 $7.0B（待公開版 S-1）；(4) 合約價 12 對 VR 世代偏低；(5) 「只計已簽合約」使 2029–2030 轉正，新增合約機制需 Andy 決定。

## 段落總覽
| 段 | 工作單 | 狀態 |
|---|---|---|
| A0 骨架 | （chat 端） | 完成 |
| A1 資料蒐集 | A1_資料蒐集 | 完成（報告 `20261008_v0.1-A1_資料.md`） |
| A2 需求與營收（v0.1-A2） | A2_需求與營收 | 完成（報告 `20261008_v0.1-A2.md`） |
| A3 算力與成本（v0.1-A3） | A3_算力與成本 | 完成（報告 `20261008_v0.1-A3.md`） |
| A4 融資與反向（v0.1） | A4_融資與反向 | 完成（報告 `20261008_v0.1.md`） |
| A5 成品 | A5_成品 | 完成（報告 `20261008_v0.1_成品.md`） |

<!-- 各段在下方新增「## An」＋步驟紀錄表（步驟｜狀態｜commit｜備註）＋已套用的預設表 -->

## A2 需求與營收（v0.1-A2）
| 步驟 | 狀態 | commit | 備註 |
|---|---|---|---|
| 1 builder 骨架（README、SRC_ANT、TK_Link、OAI_Link、Inputs、Checks；registry） | 完成 | `aa06cb2` | SRC_ANT 379 列（新增 377–379 供 Derived 公式化）；Derived 12 列：11 公式、1 保留報導值；TK 63 名＋NonNV 18 格；OAI_Link 17 名；CHK_Errors＝0 |
| 2 TK_Link NonNV（D6）＋ check_tk_snapshot | 完成 | `52d8dd5` | TK_NNV_{TPUv7,Trn3,MI455X}_{Out,OutLo,OutHi,Hold,HoldLo,HoldHi} 18 格；check_tk_snapshot：OK 81（63＋18） |
| 3 Inputs 草稿 | 完成 | `7609b7c` | 101 列：定義常數 14、開關 4、訂閱 11、人數成長 15、使用量 29、API 23、通路與其他 3（Analogy 多引 OpenAI v0.6 INP 編號） |
| 4 Demand | 完成 | `691b46f` | 兩頁互相引用，4–6 同一 commit（builder/a2.py 的 build／checks） |
| 5 Revenue | 完成 | 同上 | 2025 總額 4.60（校準差 0）；2026 48.73（半校準）；2030 總額 214.9、淨額 196.8 |
| 6 Checks | 完成 | 同上 | C01–C46；CHK_Errors＝0、CHK_Warnings＝0（2026 對里程碑推估 −17.4%，未超 25%） |
| 7 測試與 CI | 完成 | `d9c6284`；修正見下一列 | pytest 52 項全過（parity 情境 18 個 × 2＋結構 16）；`.github/workflows/anthropic-parity.yml`（ANTHROPIC_ENGINE_CACHE）；修正：日期列改數值格式、類別檢查改相對差（浮點雜訊） |
| 7a 修正：每任務 token 改取 Tokenomics（IF_TaskTok*；分層規則） | 完成 | `fd5db5b` | TK_Link＋5 名（68）；Inputs k_chat／k_code／k_other（INP_067–069）退役，改 Tokenomics 任務對應（INP_102–104）× 倍數（INP_105–107）；pytest 54 項全過 |
| 8 報告、CHANGELOG、PR 留言 | 完成 | 見 git log「A2 步驟 8」 | 報告＋對照 Excel（3 頁）；PR #22 留言 |

已套用的預設（A2）：共 24 項，全表見報告 `docs/reports/20261008_v0.1-A2.md`「已套用的預設」與對照 Excel 第③頁。影響最大者：API 任務數成長（2030 淨額 +236／−71）、價格彈性（+64／−48）、2026 半校準基準（+30）、抽成率上緣（−16）。

## A3 算力與成本（v0.1-A3）
| 步驟 | 狀態 | commit | 備註 |
|---|---|---|---|
| 1 Compute（加速器族 9 族、逐合約供給、自建、組合、η、研發 GW、容量上限、VR 等值、敏感度） | 完成 | `c9d2478` | `builder/a3.py`；REV_CapFactor＝CMP_CapFactor（2025、2026 固定 1，INP_016 退役）；TK_PUE 讀表 |
| 2 Cost（逐合約實付、自建、供應商持有成本、經濟口徑、非算力、股權報酬、命題表、TK 參考、D22、$518B 對帳） | 完成 | 同上 | 2025 算力成本＝7.33（差 0）；2030 每 VR 等值 GW 差額 ＋11.2、覆蓋率 1.78 |
| 3 Checks C49–C82 | 完成 | 同上 | CHK_Errors＝0、CHK_Warnings＝0；年底 GW ＋20%／＋1%；$518B −3.1% |
| 4 parity 情境、報告、對照 Excel、CHANGELOG、PR 留言 | 完成 | 見 git log「A3 步驟 4」 | 情境 27 個（A3 新增 10）；pytest 71 項全過 |

已套用的預設（A3）：共 30 項，全表見報告 `docs/reports/20261008_v0.1-A3.md`「已套用的預設」與對照 Excel 第③頁。對 2030 差額影響最大的幾項：
- 牌價年降 35%（容量截頂，−8.8）；
- AMD 期間 3 年（＋4.5）；
- NonNV 產出比低／高（＋3.4／−2.9）；
- 口徑未明 GW 視為設施（＋1.8）；
- 最高可達計入比例（＋1.5／−0.9）；
- Broadcom 5 GW（−1.5）。


## A4 融資與反向（v0.1）
| 步驟 | 狀態 | commit | 備註 |
|---|---|---|---|
| 0 check_tk_snapshot | 完成 | — | CURRENT 仍 v5.26（master `70d859e`）；OK 89 |
| 1 Funding（自由現金流、來源順序、最低現金、外部資金、條件式、或有、回流對照）＋Cost 第十二–十五節 | 完成 | `07565be` | 已到位 $100B；累計外部資金需求 0；谷底 80.2（2028） |
| 2 Reverse＋test_reverse_not_fed_back | 完成 | 同上 | 反向累計 25.2 |
| 3 Checks C83–C105、parity 情境 ＋7（共 34）、敏感度 | 完成 | 同上（含 `tools/sensitivity.py`） | pytest 87 項全過；CHK_Warnings＝1（C89 預期） |
| 4 v0.1 完成報告＋對照 Excel（①–⑤） | 完成 | 見 git log「A4 步驟 4–5」 | ④A1–A4 預設彙總 |
| 5 CHANGELOG、進度檔、交接檔、PR 留言 | 完成 | 同上 | — |

已套用的預設（A4）：共 21 項，全表見報告 `docs/reports/20261008_v0.1.md`「已套用的預設」與對照 Excel 第③頁；A1–A4 彙總（規格 D1–D23＋A2 24＋A3 30＋A4 21）見第④頁。影響最大者：Series G $30B 計入（谷底 −30）、D22 補足（谷底 80.2 → 52.3）、逐家錨定招股書（谷底 52.3）。


## A5 成品
| 步驟 | 狀態 | commit | 備註 |
|---|---|---|---|
| 1 HTML 產生器、html_scenarios（20 情境）、test_html、dist HTML＋xlsx、截圖、CI | 完成 | `5d08b70` | 818 個數字；test_html 6 項全過；截圖外部請求 0、水平捲動 0 |
| 2 命題與驅動對照表 | 完成 | `210faf0` | `docs/plan/20261008_Anthropic_命題與驅動對照表.md` |
| 3 交接檔 v0.1 完成、根目錄 README Anthropic 列 | 完成 | `82e88e7` | — |
| 4 成品報告、CHANGELOG、進度檔、PR 留言、轉 ready | 完成 | 見 git log「A5 步驟 4」 | pytest 93 項全過（本地）；不合併 |

已套用的預設（A5）：11 項，全表見報告 `docs/reports/20261008_v0.1_成品.md`「已套用的預設」（皆為呈現層，無數值影響）。
