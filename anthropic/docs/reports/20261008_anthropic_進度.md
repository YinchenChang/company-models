# Anthropic v0.1 分段建置進度（接手用；A1–A5 共用）

## 目前狀態（每次 push 前覆寫）
- 已完成：A0（`aa925dd`）；A1（`8d955cf`）；A2 步驟 1（`aa06cb2`）、2（`52d8dd5`）、3（`7609b7c`）、4–6（`691b46f`）、7 測試與 CI（`d9c6284`）、修正：任務類別每任務 token 改取 Tokenomics（本 commit）。
- 下一步：A2 步驟 8：報告 `docs/reports/20261008_v0.1-A2.md`＋`_對照.xlsx`（3 頁）、CHANGELOG、PR #22 留言。
- 未解問題：無。

## 段落總覽
| 段 | 工作單 | 狀態 |
|---|---|---|
| A0 骨架 | （chat 端） | 完成 |
| A1 資料蒐集 | A1_資料蒐集 | 完成（報告 `20261008_v0.1-A1_資料.md`） |
| A2 需求與營收（v0.1-A2） | A2_需求與營收 | 進行中 |
| A3 算力與成本（v0.1-A3） | A3_算力與成本 | 未開始 |
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
| 7a 修正：每任務 token 改取 Tokenomics（IF_TaskTok*；分層規則） | 完成 | 見 git log「A2 修正」 | TK_Link＋5 名（68）；Inputs k_chat／k_code／k_other（INP_067–069）退役，改 Tokenomics 任務對應（INP_102–104）× 倍數（INP_105–107）；pytest 54 項全過 |
| 8 報告、CHANGELOG、PR 留言 | 未開始 | | |

已套用的預設（A2）：見報告 `docs/reports/20261008_v0.1-A2.md`（完成時）；進行中先列於此：
| 問題 | 預設 | 替代 | 影響 |
|---|---|---|---|
| Derived 列中的常數（成長率 1,088%、10%、倍增） | 由原文另立 SRC_ANT_377–379（只增不重用），Derived 列以公式引用 | 移入 Inputs | 無數值影響（值不變，改為可追溯） |
| SRC_ANT_031 通路費 $351M（標 Derived） | 保留報導值（Reuters 依招股書計算後公布），SRC_ANT_032 抽成率改公式＝031 ÷（4.59 × 47%）＝0.1627 | 031 也改公式（會與 032 循環） | 抽成率 0.16→0.1627（只影響對照列） |
