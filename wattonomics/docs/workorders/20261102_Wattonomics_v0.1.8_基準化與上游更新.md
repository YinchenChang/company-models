# 工作單：Wattonomics v0.1.8＝基準化（三層只給基準）＋延遲逐站追蹤＋R4 上游更新

- 日期：2026-11-02（台北）；chat 端：算力收支 Wattonomics 串（session_01V11BcwWqQyh9suCNhAYQKQ）
- 現行：main `63ad400`，`model/20261031_Wattonomics_v0.1.7.xlsx`（r5），Checks 177＝OK 167／WARN 8／ERROR 0
- 分支：`claude/wattonomics-v0.1.8-base`（舊分支 `claude/wattonomics-v0.1.8`＝已關閉 PR #290，只作參考：Gap 逐家寫法與 export 結構）
- 產出：`model/20261102_Wattonomics_v0.1.8.xlsx`、快照 `data/wattonomics_snapshot_v0.1.8.json`、`data/explorer_v0.1.8.json`、互動頁、README、`docs/verify/20261102_v0.1.8_build.md`、R4 r2 歸因

## 0. Andy 原話（依據）

- 11-02 15:58：「Tokenomics(1GW單位基準)，Wattonomics(tokoenomics的全市場版)及Siliconomics(Wattonomics的零部件版)都只要提供基準版的數據就好。情境只留在下游取用的個股。個股才根據自身狀況作情境分析與預估。Siliconomics對帳的時候，只對個股的基準情境加總。Tokenomics及Wattonomics也是照這個原則處理。」
- 11-02 16:22：「3-nomics與個股的地位鏈應該是Tokenomics->Wattonomics->Siliconomics->覆蓋個股；對帳的時候則反過來: 覆蓋個股->Siliconomics->Wattonomics->Tokenomics。若對帳時差很大……則先用自己的數字，在較上游趕快更新後彼此對帳。」
- 11-02 15:13：逐家表要有「新增 GW」欄，合計與逐年數字一致。
- R4 歸因報告 `docs/verify/20261031_R4_Wattonomics側歸因_r1.md`「排入 Wattonomics v0.1.8」節。

## 1. chat 端已定（工程決定，照做；理由附後）

| # | 決定 | 理由 |
|---|---|---|
| E1 | **基準＝需求 A × 延遲「D0＋已知延遲」**。Players／Pipeline 的 D0 上線時點改為「公告上線年月＋T2 已知延遲月」（T2 M 欄），標籤改「基準（公告＋已確認延遲）」。D1、D2 欄與 B、C 欄保留公式、標籤加「〔停用，v0.2 刪〕」，不再進任何收支或頁面。 | 已確認延遲是事實，屬基準；D1 對已知延遲園區會重複延，但已停用，於標籤註明即可 |
| E2 | **不新增 `_Base_` 別名**：`*_A_D0_*` 名稱即基準名稱；其餘 `_B_`、`_C_`、`_D1`、`_D2`、`WIF_ShippedNotOnline_*`、Gap 頁 r(s,D) 衍生名稱在 Interface 說明欄標「停用（v0.2 刪）」，數值照算。 | 下游無人取用情境（11-02 掃描）；避免先加別名後又改名兩次 |
| E3 | 收支（Gap）與頁面只用基準。Gap 頁的 A／B／C × D0–D2 區塊保留但標停用；「11 家當一間公司」主表、判讀、累計只讀基準欄。 | Andy 15:58 |
| E4 | 新增名稱：`WIF_Gap_Player_<p>_<yyyy>_AddGW`（11 家 × 2024–2030，擁有者口徑年新增 IT GW，即 Gap 頁計 AI 資本支出所用的 GW）、`WIF_Gap_Add11_<yyyy>`（＝Σ11 家）；Checks 新增「Σ逐家 AddGW＝Add11（0.001 GW）」；check.py 同步。 | 15:13 |
| E5 | GL01-19 改：主路徑七家自 **2026** 起一律甲2（含 Meta、CoreWeave 2027）。GL01-23 改：甲2 需要的 t−1、t−2 資本支出缺值時，缺 t−2 用 t−1，兩者皆缺用 t；代入處在 CapexGW 標「代入」。 | R4 報告：跳過整年的偏差（2026 約 −8.7 GW）遠大於代入誤差 |
| E6 | 甲2 由付款年轉上線年：In S 新增兩個 Wattonomics 自有參數 `lag_on`（付款到上線延遲，4.5 個月；3–6；Assumed）、`inv_on`（未上線庫存 ÷ 年出貨，0.10；0.05–0.20；Assumed），公式與 `tools/r4_attrib.py` d2 相同（延遲平移＋庫存增量扣除，合計約 5.7 個月效果）。GL01-08 撤回「IT 付款年＝上線年」。值與 Siliconomics ②③ 相同但屬 Wattonomics 權威。 | R4 報告 Independent Analysis 第 1 點 |
| E7 | CapexGW 七家資本支出改讀 In Y9（個股模型基準），Y9 缺年份（Microsoft、Oracle、Nebius FY2025 等）用 In Y2；In Y2 保留作 C1／Compare 對照。先重跑 `tools/import_gap_models.py` 取最新模型：Microsoft v0.3.3、Alphabet v0.3.3、Amazon v0.3.1、Meta v0.1.1、Oracle v0.4.2、Nebius v0.5.2、CoreWeave、WhiteFiber、SpaceX v0.2.1、OpenAI v0.6.4、**Anthropic v0.1.7**（自建＝0；`WIF_Gap_Player_ant_*_CapexAIModel` 隨之換新值）。 | R4 第 3 項；Anthropic v0.1.7 已合併（#289） |
| E8 | xAI：有 SpaceX 模型 AI 資本支出的年份，Players xai 年新增＝模型 AI 資本支出 ÷ xAI 每 GW 成本（Gap 頁同一組成本、同 E6 平移）；2026 年依 R4 d4「不平移、持平」處理（Y9 無 2025）。Anthropic 自建維持 0。OpenAI／Stargate 維持 GL01-31。 | R4 第 4 項 |
| E9 | 「其他（企業、學研）」2026 年起隨市場（不含「其他」）年底 GW 成長率成長，公式同 r4_attrib d5；標 Analogy，低＝持平、高＝成長率 × 1.5。 | R4 第 5 項 |
| E10 | Deploy 年新增標籤註明「毛額，未扣汰換（汰換列 v0.2）」。 | R4 第 6 項 |
| E11 | **延遲逐站追蹤**：新頁 `Delay`：逐家（11 家＋其餘）列「公告上線 GW（T2 納入＝1、公告上線年 ≥ 2026）、已延遲 GW（M>0）、MW 加權平均延遲月（對已延遲站）、延遲占比」，另列明細（園區、MW、原時程、延遲月、證據 ID、確認等級）。新名稱 `WIF_Delay_<p>_{AnnGW,DelayedGW,AvgMo}`、`WIF_Delay_Mkt_*`。互動頁加一區塊。 | 交接檔第 4 節 1(b) |
| E12 | **「已確認延遲」門檻（寫入基準）**：公司本身揭露，或兩家以上獨立可信媒體報導且公司未否認。公司否認者（例：D#16 Oracle，Bloomberg 報導、Oracle 否認）只列追蹤表「報導、公司否認」，T2 M 欄＝0。此門檻屬判斷題，交 Andy（預設如上）。 | 公告時程本身屬 Interested-party；否認的報導不當事實 |
| E13 | 延遲回填候選：B_delay D#16（Oracle）、D#19（Nebius NJ 6–7 個月）、D#24（Fairwater Wisconsin）、D#21（Stargate UAE），及 `wattonomics-monitor` 分支 11-02 輸出中的延期／取消消息。每筆須由**未參與代理開原文核對**後才寫進 T2 M 欄（Q 段同規則：`evidence_raw/D_review_20261102.md`）。 | 證據規則 |

## 2. 建置步驟

1. 由 main 開 worktree 與分支 `claude/wattonomics-v0.1.8-base`。
2. 延遲證據：派查證子代理整理候選（D#16、#19、#21、#24＋monitor 11-02），再派另一個未參與代理開原文逐筆判定（屬實／錯誤／無法核對＋延遲月數＋確認等級）；判定結果寫 `evidence_raw/D_review_20261102.md`，屬實且達 E12 門檻者寫 T2 M 欄（經 `--base` 讀回或 apply_in_edits）。
3. `import_gap_models.py` 重取 Y9（E7）；核對各家標籤與版本。
4. 改 `builder/build.py`：E1–E11；Decisions 新增 GL01-39 起（引 Andy 原話與本工作單），GL01-08、-19、-23 改狀態並註明撤回日期；README 頁版本紀錄。
5. 建置 → soffice 重算 → export_snapshot、export_explorer → check.py → `tools/diff_snapshot.py`（v0.1.7 → v0.1.8，列所有 `WIF_` 變動的前 30 大與分類）。
6. **重現檢查**：在 v0.1.8 Excel 中把 Y9 換回 Y2、known delay 歸 0、E9 歸持平時（只在暫存複本），2026–2029 七家甲2 上線年新增應重現 R4 d1＋d2 的值（2026 +8.67−2.65 等）；不符要說明。
7. 互動頁：拿掉需求與延遲兩組按鈕；Gap 區塊只基準；「需要：11 家 AI 資料中心」改名「AI 資本支出（11 家）」並加「11 家年新增 IT GW」列；逐家表加「新增 GW」欄、合計列＝逐年表；延遲追蹤區塊；README 與標頭標「基準」。`build_page.py` 產生本機頁面（**不發布 gh-pages、不更新 artifact**——合併後由 chat 端做）。
8. R4：`tools/r4_attrib.py` 加「單一基準比較」模式（v0.1.8 基準 vs Siliconomics r1 橋接），輸出 `docs/verify/20261102_R4_Wattonomics側歸因_r2.{md,xlsx}`：逐年殘差、比值、判定（一致／時差／待診斷），並對照 r1 預測的調整後值（19.42／26.01／29.36／27.91）說明差異來源（Y9 更新、已知延遲、xAI）。
9. 報告 `docs/verify/20261102_v0.1.8_build.md`（STATUS 格式：結論、做了什麼、驗了什麼／沒驗什麼、待 Andy 決定）；README 固定規格節改寫：「三層只給基準」「地位鏈與對帳方向」原話、「已確認延遲」門檻、需求情景／延遲情景節改為「停用，v0.2 刪」。
10. commit（不 push 到 main；push 分支即可），回報 chat 端。

## 3. 驗收

- 錯誤格 0；Checks ERROR 0（WARN 數與原因列出）；check.py 通過；快照與 Excel 一致。
- Σ逐家 AddGW＝Add11；頁面逐家表合計＝逐年表。
- 第 6 步重現檢查通過或差異已解釋。
- 未參與代理獨立核對（chat 端另派）：重算 2026 七家甲2、抽 5 個 `WIF_` 變動對公式追到輸入、延遲寫入的每一筆對原文、頁面數字對快照。

## 4. 不做

- 不刪任何停用欄位或名稱（v0.2）；不加汰換列（v0.2）；不改 TK、Siliconomics 權威的量；不碰 OpenAI／Stargate 擁有者規則。
