# Anthropic v0.1 分段建置進度（接手用；A1–A5 共用）

## 目前狀態（每次 push 前覆寫）
- **v0.1.7 完成**（2026-10-31；工作單 `docs/workorders/20261031_anthropic_v0.1.7_S1全面套用.md`；分支 `claude/anthropic-v0.1.7`；**不合併、未開 PR**）。步驟 1 模型（W1–W6）、2 測試／HTML／dist、3 報告與文件完成；報告 `docs/reports/20261031_v0.1.7.md`。**下一步**：chat 端審查；未參與代理獨立核對瀑布與首節數字；Andy 決定未歸屬餘額付款期等替代讀法。**未解問題**：S-1 付款時程留白（未歸屬餘額付款期為 Assumed）；η 0.468 低於 WARN 下緣；sec.gov 一手章節仍未直讀。
- **v0.1.6 完成**（2026-10-31；工作單 `docs/workorders/20261031_anthropic_v0.1.6_S1正式版.md`；分支 `claude/anthropic-v0.1.6`；**不合併**）。步驟 1 模型（W1–W3）、2 測試、3 HTML／dist／對照 Excel、4 報告與文件完成；報告 `docs/reports/20261031_v0.1.6.md`。**下一步**：chat 端審查 PR；以瀏覽器核對 S-1 二手列；決定 6-30 後新簽合約是否進基準、OAI_Link 與 TK 是否同步。**未解問題**：sec.gov 一手章節（季度、Note 8、承諾）未能直讀。
- **v0.1.5 完成**（2026-10-24；工作單 `docs/workorders/20261024_anthropic_v0.1.5_TK21.2.md`；分支 `claude/anthropic-v0.1.5`；**不合併**）。步驟 1–5 全部完成，報告 `docs/reports/20261024_v0.1.5.md`；pytest 117 項全過（builder 14、parity 97、html 6）；主要輸出與 v0.1.4 逐位相同。**下一步**：chat 端審查 PR（v0.1.5）；待決僅 U2 付款分攤比例（A 段遺留）。**未解問題**：無新增。
- **v0.1.4 完成**（2026-10-22；工作單 `docs/workorders/20261022_anthropic_v0.1.4_TK19.1.md`；同分支 `claude/anthropic-v0.1.3`、PR #69，標題「Anthropic v0.1.3＋v0.1.4」；**不合併**）。步驟 1–6 全部完成，報告 `docs/reports/20261022_v0.1.4.md`；pytest 117 項全過；命題輸出與 v0.1.3 逐位相同。**下一步**：chat 端審查 PR #69；待決僅 U2 付款分攤比例（A 段遺留）。**未解問題**：無新增。
- **v0.1.3 A 段進行中**（分支 `claude/anthropic-v0.1.3`，PR #69 draft）：步驟 1–7 完成（報告 `docs/reports/20261022_v0.1.3-A.md`）。**B 段起點**：η 診斷拆分（不得改 η 值）、Bridge 頁（唯讀）、S-1 替換程序寫入 CLAUDE.md；B 段結束時重產 HTML 與 dist（目前 dist 為 A 段草稿，v0.1.2 成品在 `dist/_待刪除/`）。chat 端已更正 U2 公式（付款路徑不變、只改供給 GW）；待確認：工作單「公告金額比例」與「v0.1.2 路徑不變」的衝突（報告 ⑥）。B 段完成（2026-10-22；見下方「v0.1.3 B 段」表；PR #69 ready，待 chat 端審查合併）。

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
| 2 U2 新法（Nscale、Lambda 揭露容量＋86 當上限） | 完成 | 見 git log「v0.1.3 A2–A4」 | 新 Inputs `u2_method`（INP_187；1＝新法，0＝v0.1.1）；Compute 新列 `sup_*_disc`（揭露 GW）、Cost 第二十二節（年付款列、隱含每 GW 年付款、檢查）；Checks C127、C129–C142。2030 差額 11.139 → 11.602（chat 端更正後：供給取揭露容量，付款路徑不變） |
| 3 資金三層 | 完成 | 同上 | Funding 第十節；新 Inputs `l1_incl_cash`（第 1 層是否計 2025 年底現金，基準 0）；具名範圍 FND_L1/2/3_*；Checks 另加 d1_l2_eq（第 2 層＝基準）、d1_l_order |
| 4 持續擴張情境 | 完成 | 同上 | Cost 第二十三節、Funding 第十一節；新 Inputs exp_pct（0.3，0–0.6）、exp_first_year／exp_last_year、exp_mon_none／half／full（0／0.5／1）；基準不受影響 |
| 5 HTML 條件式措辭、①之三 | 完成 | 見 git log | `tools/build_html.py`、`tools/html_scenarios.yaml`（新增 `u2_method_old`、`v012`；v011 等情境加關閉 U2 新法）；`tests/test_html.py` 不變 |
| 6 重建與測試 | 完成 | 見 git log | test_builder 10、test_html 6 全過；pytest 113 項全過（builder 10、html 6、parity 情境 88、其餘 parity 9）；`scenarios.yaml` 預期值 formula_cells 6550、defined_names 1240、inp_rows 190；新增情境 `ha_u2_method_old`、`hb_u2_nscale_mw_up` |
| 7 報告、對照 Excel、CHANGELOG、PR 留言 | 完成 | 見 git log | `tools/make_report_v013.py` 產生對照 Excel |


## v0.1.3 B 段（Bridge 頁、η 拆解、S-1 替換程序、成品）
A 段已確認完成（進度檔 A 段步驟 1–7 全部「完成」，分支上 12 個 commit，模型 CHK_Errors＝0）。B 段起點模型：`model/20261022_Anthropic_v0.1.3.xlsx`（A 段版）。
| 步驟 | 狀態 | commit | 備註 |
|---|---|---|---|
| 1 Bridge 頁（Cost 之後，唯讀） | 完成 | 見 git log「v0.1.3 B1–B2」 | `builder/a5.py`；109 列；每列來源編號、2025–2030、標記；低／高＝2030 敏感度包絡（合約價 8／20、η 沿用／回升、產出比低／高）或 Inputs 區間（2025）；最後一節與 TK 對照（P × Q、Q、token 帳）。新增 Checks C143–C152；builder 掃描 `bridge_refs`＝0 |
| 2 η 拆解（EtaDiag 頁，唯讀） | 完成 | 同上 | η（2025）0.5273 ＝ f_P 1.237 × f_U 0.683 × f_W 1.399（機隊加權）× 殘差 0.446；f_P、f_U 有區間，f_W 無區間，輸入輸出組合「無法分離」；並列 TK L1_Gap*（OpenAI 口徑）。TK_Link 尾端 +23 名（IF_UtilActual、L1_TokGW_PaidCap_VR200、B4_MixPaid、L1_Gap* 各含低／高）。重建後既有頁（SRC_ANT、Inputs、Demand、Revenue、Compute、Cost、Funding、Reverse）逐格比對：公式與值差異 0 |
| 3 S-1 替換程序 | 完成 | 見 git log「v0.1.3 B3」 | `anthropic/CLAUDE.md` 第 8 節（新增 SRC 列、舊列標已被取代並保留、模型引用改指新列、報告七項欄位） |
| 4 成品（HTML、dist） | 完成 | 見 git log「v0.1.3 B4」 | `tools/build_html.py` 「①之三」加 η 拆解與 Bridge 逐層對照；HTML 1,134 個數字；`tools/read_html_screen.py`：12 節 × 桌機／手機，畫面文字與 data-v 0 不符、無水平捲動、外部請求 0；截圖 `docs/reports/img/20261022_v0.1.3_*`；dist 重產（v0.1.2 成品已在 `dist/_待刪除/`） |
| 5 報告、對照 Excel、CHANGELOG、交接檔、PR 留言、轉 ready | 完成 | 見 git log「v0.1.3 B4–B5」 | `docs/reports/20261022_v0.1.3.md`＋`_對照.xlsx`；PR #69 留言；PR 轉 ready（不合併） |

B 段測試（依檔分開）：test_builder 13、test_html 6、test_parity 97，合計 116 項全過；CHK_Errors＝0、CHK_Warnings＝2。B 段**完成**；PR #69 轉 ready，不合併。


## v0.1.4（TK v19.1 同步、η 診斷補充、刪除 dist/_待刪除）
Tokenomics：master `c2ecb71`，CURRENT `20261022_Tokenomics_v19.1.xlsx`（Luna CSA top-k 1024 → 512）。Andy 2026-10-22 預設：η 拆解不加「TK 絕對產能水準」因子；利用率因子計入 TK 生產折減；`dist/_待刪除/` 刪除。
| 步驟 | 狀態 | commit | 備註 |
|---|---|---|---|
| 1 快照 v19.1 | 完成 | 見 git log「v0.1.4 步驟 1」 | 重建 24 秒；check_tk_snapshot：OK 126、DIFF 0；重建前 v19.0 快照 vs v19.1 為 DIFF 39（與 chat 端一致）；v0.1.3 檔 git mv 到 archive |
| 2 η 診斷補充（EtaDiag 第七節） | 完成 | 見 git log「v0.1.4 步驟 2」 | f_U' 0.581（＝0.683×0.85）；殘差 0.440 → 0.517（×1.176）；殘差倒數 1.933 ÷ TK L1_GapUnexplained 5.616 ＝ 0.344；η（2025）＝0.5254（TK v19.1 使 η 由 0.5273 小降，診斷不改 η）；版本小表加 v19.1 欄在報告／HTML（步驟 5–6） |
| 3 刪除 dist/_待刪除 | 完成 | 見 git log「v0.1.4 步驟 3」 | git rm 6 檔 |
| 4 重算與測試 | 完成 | 見 git log「v0.1.4 步驟 4–5」 | LO 錯誤 0、CHK_Errors＝0、Warnings＝2；pytest 依檔分開：test_builder 14、test_html 6、test_parity 97（分 7 片），共 117 項全過；預期值：formula_cells 7,256、defined_names 1,276、tk_values 105 |
| 5 HTML 與 dist | 完成 | 同上 | dist v0.1.4；HTML 1,232 個數字；Playwright 12 節 × 桌機／手機 0 不符、無水平捲動、外部請求 0 |
| 6 報告、CHANGELOG、交接檔、PR 留言 | 完成 | 見 git log「v0.1.4 步驟 6」 | `docs/reports/20261022_v0.1.4.md`＋`_對照.xlsx`；PR #69 留言 `[Anthropic 回報] v0.1.4｜完成｜2026-10-22` |


## v0.1.7 正式 S-1 全面套用（2026-10-31）
| 步驟 | 狀態 | commit | 備註 |
|---|---|---|---|
| 1 模型（W1 承諾進基準、W2 自建 GW 法、W3 2025 校準、W4 非算力校準、W5 TK v21.7／OpenAI v0.6.4、W6 Amazon 核對） | 完成 | `c1cf693` | builder/a7.py；INP_197–214；開關 0＋TK v21.2 → 與 v0.1.6 逐格相同 |
| 2 測試、HTML「①之五」、dist、html 情境、報告產生器 | 完成 | `2fe88cc` | parity +9 情境；test_builder +4；test_html 敘述主張改寫 |
| 3 報告、對照 Excel、CHANGELOG、交接檔、進度檔 | 完成 | 見 git log「步驟 3」 | |

已套用的預設（v0.1.7）：全表見報告 ⑤ 與對照 Excel ④（未歸屬餘額 2027–2033、上限型沿用既有規則、Google 額外 TPU 不進基準、自建 GW 0、二手列維持 Interested-party、W4 照推得值入基準、AMD 不扣、Broadcom 首個爬坡年 2027 等）。
