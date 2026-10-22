# Anthropic v0.1 分段建置進度（接手用；A1–A5 共用）

## 目前狀態（每次 push 前覆寫）
- **v0.1.3 A 段進行中**（分支 `claude/anthropic-v0.1.3`，PR #69 draft）：步驟 1–7 完成（報告 `docs/reports/20261022_v0.1.3-A.md`）。**B 段起點**：η 診斷拆分（不得改 η 值）、Bridge 頁（唯讀）、S-1 替換程序寫入 CLAUDE.md；B 段結束時重產 HTML 與 dist（目前 dist 為 A 段草稿，v0.1.2 成品在 `dist/_待刪除/`）。留意：HTML ⑥ combo1 已不產生外部資金需求（見報告 ⑤）。B 段尚未開始。

- 已完成：v0.1（A0–A5）；v0.1.1（PR #56）；**v0.1.2（2026-10-21）全部步驟**：TK v5.26 → v18.0（master `868e2cc`）、語意核對、對照列、重算、pytest 109 項全過、HTML／dist、報告＋對照 Excel、CHANGELOG、交接檔。報告 `docs/reports/20261021_v0.1.2.md`。
- 下一步：chat 端審查 PR #56（v0.1.1＋v0.1.2）並合併（代理不合併）；Andy 確認後刪除 `dist/_待刪除/`；OpenAI 模型同步 TK v18.0。
- 未解問題：(1) GitHub Actions 因帳號付款／用量上限未啟動（本地 pytest 為準）；(2) CHK_Warnings＝2（C89、C116）；(3) D22 低估 5–10；(4) 2026-06-30 現金等待公開版 S-1；(5) 只計已簽合約；(6) OpenAI v0.6 仍用 TK v5.26；(7) 待 Andy：η 是否回升、是否改 `_Prod` 層。

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


## v0.1.3 A 段（TK v19.0、U2 改法、資金三層、持續擴張、條件式措辭）
| 步驟 | 狀態 | commit | 備註 |
|---|---|---|---|
| 1 TK 標記改 v19.0 | 完成 | 見 git log「v0.1.3 A1」 | check_tk_snapshot：快照 v18.0 vs CURRENT v19.0 → OK 103、DIFF 0、MISSING 0；重建後 27,934 格逐格比對，322 格差異全為版本／日期／SHA 文字，數值差異 0。TK 提交 `1916a06`，SHA-256 `d1728705…`。OpenAI 快照維持 v0.6（`34fc34a7…`；main 現行為 v0.6.1，不在本版範圍） |
| 2 U2 新法（Nscale、Lambda 揭露容量＋86 當上限） | 完成 | 見 git log「v0.1.3 A2–A4」 | 新 Inputs `u2_method`（INP_187；1＝新法，0＝v0.1.1）；Compute 新列 `sup_*_disc`（揭露 GW）、Cost 第二十二節（年付款上限、約束年數、檢查）；Checks C127–C139。2030 差額 11.139 → 12.133 |
| 3 資金三層 | 完成 | 同上 | Funding 第十節；新 Inputs `l1_incl_cash`（第 1 層是否計 2025 年底現金，基準 0）；具名範圍 FND_L1/2/3_*；Checks 另加 d1_l2_eq（第 2 層＝基準）、d1_l_order |
| 4 持續擴張情境 | 完成 | 同上 | Cost 第二十三節、Funding 第十一節；新 Inputs exp_pct（0.3，0–0.6）、exp_first_year／exp_last_year、exp_mon_none／half／full（0／0.5／1）；基準不受影響 |
| 5 HTML 條件式措辭、①之三 | 完成 | 見 git log | `tools/build_html.py`、`tools/html_scenarios.yaml`（新增 `u2_method_old`、`v012`；v011 等情境加關閉 U2 新法）；`tests/test_html.py` combo1 斷言改為 0、combo3 為正 |
| 6 重建與測試 | 完成 | 見 git log | test_builder 10、test_html 6 全過；parity 分 3 片（結果見 PR 留言）；`scenarios.yaml` 預期值 formula_cells 6550、defined_names 1240、inp_rows 190；新增情境 `ha_u2_method_old`、`hb_u2_cap_binds` |
| 7 報告、對照 Excel、CHANGELOG、PR 留言 | 完成 | 見 git log | `tools/make_report_v013.py` 產生對照 Excel |
