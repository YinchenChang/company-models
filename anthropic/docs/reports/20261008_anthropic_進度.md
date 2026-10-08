# Anthropic v0.1 分段建置進度（接手用；A1–A5 共用）

## 目前狀態（每次 push 前覆寫）
- 已完成：A0（`aa925dd`）；A1（`8d955cf`）；A2（`fd5db5b`＋報告 `e903023`）；**A3 完成（v0.1-A3）**：步驟 1–3（`c9d2478`）、步驟 4 報告與對照 Excel（本 commit）。報告 `docs/reports/20261008_v0.1-A3.md`。
- 下一步：等 chat 端審查 A3 後開工 A4（工作單 A4）：新增 `builder/a4.py`（Funding、Reverse），build.py 自動載入；見 A3 報告「給 A4 的交接」6 點（COST_GapCash 起點、股權報酬加回、D21 非現金、自建資本支出、回流對照用 COST_Pay_*、2029–2030 正差額的前提）。
- 未解問題：(1) GitHub Actions 因帳號付款／用量上限未啟動（本地 71 項全過）；(2) D22 對照：模型 2026 調整後營業利益率 29% vs 報導 Q2 約 5%（模型 2026 非算力成本可能偏低約 $4B；依 D22 未校準，待 chat 端判斷）；(3) 基準合約價 12 低於 TK VR200 持有成本，雲端毛利為負（待 chat 端判斷是否改合約價口徑）。

## 段落總覽
| 段 | 工作單 | 狀態 |
|---|---|---|
| A0 骨架 | （chat 端） | 完成 |
| A1 資料蒐集 | A1_資料蒐集 | 完成（報告 `20261008_v0.1-A1_資料.md`） |
| A2 需求與營收（v0.1-A2） | A2_需求與營收 | 完成（報告 `20261008_v0.1-A2.md`） |
| A3 算力與成本（v0.1-A3） | A3_算力與成本 | 完成（報告 `20261008_v0.1-A3.md`） |
| A4 融資與反向（v0.1） | A4_融資與反向 | 未開始 |
| A5 成品 | A5_成品 | 未開始 |

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

