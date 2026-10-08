# CoreWeave｜Tokenomics 連接審視（v5.29；只審視回報，不改模型）

- 分支：`claude/coreweave-tk-link-review`
- 最新提交 SHA：見 PR 最後一則留言（本檔的內容提交為底稿 `75327ef` 之後的第一個提交）
- 報告檔：`coreweave/docs/reports/20261008_coreweave_tokenomics_link_review.md`
- 依據：共用工作單 `docs/workorders/20261008_tokenomics_link_review.md`；本公司工作單 `coreweave/docs/workorders/20261008_coreweave_tokenomics_link_review.md`；Tokenomics 下游資料契約 v0.1（`docs/plan/Tokenomics_downstream_contract.md`）。
- 審視對象：(1) main 上的現行成品 **v4.6**（`dist/20261008_CoreWeave收支模型_v4_6.*`）；(2) 未合併的 **PR #27**（分支 `claude/coreweave-w4-revenue` @ `a3405e2`，W4 收入錨定＋W5 公司實況驗證，成品 v4.7）只作背景閱讀，未合併。
- Tokenomics：master `a5061d9`，`model/CURRENT`＝`20261008_Tokenomics_v5.29.xlsx`（前提成立，第 4 節照做）。
- 本輪沒有改任何模型檔（company.json、builder、Excel、data 皆未動）；v5.29 快照只產生在暫存路徑 `/tmp/coreweave_v529_snapshot.json`，未提交。

---

## 一、結論（先看這裡）

1. **取數方式合規**：CoreWeave 是八個模型中唯一用官方快照工具取數的（`data/tokenomics_names.txt` → `data/tokenomics_snapshot_v5.26.json` → Excel「Tokenomics_取數」分頁的 `TK_` 具名範圍）。所有引用都是 `IF_`／`L1_` 名稱；**人工確認後違規引用 0 筆**（工具標出的 3 筆 `IF_Hdr` 都是說明文字「不含 IF_Hdr*」的誤判）。
2. **版本落後但數字不受影響**：main 釘在 v5.26（`4074684`），PR #27 釘在 v5.27（`862bdd4`）。以 v5.29 重取後，**CoreWeave 引用的 25 個名稱（PR #27 為 26 個）數值與儲存格位置全部相同，差異 0 格**。v5.29 改變的 `IF_Alloc*` 六格與 L1_Ans1／Ans2 等，CoreWeave 都沒有引用。所以升到 v5.29 只是「換釘選版本」，目標價不會變。
3. **主要缺口在新規則**：v5.29 新增的四層瀑布（`IFW_`）與用途欄（`IFC_`）都還沒取。對 CoreWeave 實際有影響的只有一處：PR #27 的「收入上限檢查」用 `IF_RevGWFleet`，它是「× 60% 利用率」那一層，契約要求四層同取；換成「× 生產降額」層（`IFW_RevGWFleet_Prod`）時，FY26 上限比例約由 42% 升到約 57%，會超過 50% 門檻、檢查由「通過」變「警告」。用哪一層要 chat 端判斷。
4. **口徑大致一致**：CoreWeave 按 MW-year／GPU 小時計價，用公司自己的「售出比例」（模型裡叫「利用率」95%→92%），沒有用 Tokenomics 的技術利用率 `IF_Util`，與契約第 3 節一致；Tokenomics 的持有成本是 100% 時數口徑，不會與公司的售出比例重複扣減。但模型內部「利用率」與「未簽約部分填補率（fill）」兩者可能對同一塊未售出產能扣兩次，需要判斷。
5. **官方工具不支援新前綴**：`import_tokenomics.py` 目前只收 `IF_`／`L1_`，`IFW_`、`IFC_` 會直接報錯；要先升級共用工具，CoreWeave 才能照契約取四層瀑布。

---

## 二、取數點清單（共用工作單第 1 節）

先跑 `python3 tools/tokenomics/link_audit.py --tokenomics /home/claude/tk-master`，CoreWeave 節的結果：官方快照、釘 v5.26、使用名稱 25 個、違規候選 3 筆（全部是 `IF_Hdr`）。逐筆人工確認如下。

「所用版本」欄：main＝v4.6 成品，快照 `20261007_Tokenomics_v5.26.xlsx`（Tokenomics `4074684`）；PR #27＝v4.7，快照 `20261008_Tokenomics_v5.27.xlsx`（`862bdd4`）。兩者數值相同。

「IFC_Use」欄取自 v5.29 Interface 的 S 欄：A 級＝「基準輸入」；B 級＝「情境、相對比較、反轉門檻」（可作主值，但要附低／高敏感度）；另加註「上限，不得作預測」者不得當收入預測。

### 2a. 模型實際採用（進入損益、資金或評價）

| 本模型的格或參數 | Tokenomics 名稱 | 合規／違規 | 四層瀑布的哪一層 | 是否符合 IFC_Use | 所用版本 |
|---|---|---|---|---|---|
| 「輸入與假設」TK 資本區「IT 設備」→ D 區「每 MW 建置成本」＝Σ 新增世代占比 × 本值；驅動 CapEx、汰換、車隊折舊 | `IF_CapexIT` | 合規（IF_） | 不適用（成本） | 符合：B 級，主值取基準，低／高成本進敏感度（「Tokenomics 成本情境」選擇格） | v5.26（PR #27：v5.27） |
| TK 資本區「合計」→ 稅險的 IT 占比分母；每MW經濟性「含廠房合計（對照）」 | `IF_CapexTotal` | 合規 | 不適用 | 符合（B） | 同上 |
| TK 資本區「設施／IT」→ MW 口徑換算（預設 IT 口徑，換算係數 1） | `IF_FacilityGW` | 合規 | 不適用 | 符合（A 基準輸入） | 同上 |
| D 區「GPU 經濟壽命」＝依期初在役世代加權取整（6 年） | `IF_DeprLifeIT` | 合規 | 不適用 | 符合（B） | 同上 |
| 由下而上營運成本「電費」 | `IF_PowerCost` | 合規 | 不適用 | 符合（A） | 同上 |
| 由下而上營運成本「IT 維護」 | `IF_MaintIT` | 合規 | 不適用 | 符合（B） | 同上 |
| 由下而上營運成本「人員、軟體、水與耗材」 | `IF_StaffSW` | 合規 | 不適用 | 符合（A） | 同上 |
| 由下而上營運成本「財產稅與保險 × IT 占比」 | `IF_TaxIns` | 合規 | 不適用 | 符合（B） | 同上 |
| TK 收入參考「GPU／MW」→ 每 MW GPU 數（GPU 小時路線、GPU 數加權） | `IF_GPUsPerGW` | 合規 | 不適用（100% 物理量） | 符合（B） | 同上 |
| **PR #27 才有**：每 MW 年收入＝Σ 在役占比 × 本值 × 定價倍數 k（main 上只作「不賠錢下限」對照列與一頁摘要每 MW 句） | `IF_HoldEcon` | 合規 | 不適用（100% 時數的打平租金，未扣利用率） | 符合：B 級，用法是「市場價 ÷ 持有成本」的相對比較，不是上限列；低／高成本情境下 k 會重算、價格不變 | main：對照；PR #27：收入主值 |
| **PR #27 才有**：k_現貨證據的分母；main 上為每MW經濟性「每 GPU 小時經濟持有成本」對照 | `IF_GPUhrEcon` | 合規 | 不適用（100% 時數） | 符合（B） | 同上 |
| **PR #27 才有**：「每MW經濟性」與檢查頁「收入上限檢查」（CRWV 計費收入 ÷ 客戶每 MW 付費 token 營收，門檻 50%） | `IF_RevGWFleet` | 名稱合規；**未依契約第 2 條第 2 項同取四層** | **IF_Util 層（× 60%）**，只取這一層 | 用途符合（只作上限檢查，未當收入預測；該列 IFC_Use＝「…；上限，不得作預測」） | PR #27：v5.27 |

### 2b. 只作對照或說明（不進任何計算）

| 本模型的格或參數 | Tokenomics 名稱 | 合規／違規 | 層 | IFC_Use | 版本 |
|---|---|---|---|---|---|
| W1 對照表（`docs/plan/20261007_coreweave_Tokenomics對照表`）「利用率」列，註明「口徑不同，W2 不引用」 | `IF_Util` | 合規 | IF_Util | 符合（只對照） | v5.26 |
| W1 對照表：租金 ÷ 廠房造價 | `L1_FacCapexMW` | 合規 | 不適用 | 符合 | v5.26 |
| W1 對照表：GB200 持有成本對 CoreWeave 隨需牌價 | `L1_GPUhr_GB200_vsCW` | 合規 | 不適用 | 符合 | v5.26 |
| W1 對照表：GB300 持有成本對新雲損益兩平租金 | `L1_GPUhr_GB300_vsBE` | 合規 | 不適用 | 符合 | v5.26 |
| W1 對照表：token 層理想上限（VR200） | `L1_RevGW_Fleet_VR200` | 合規；只有單一層（基準利用率） | IF_Util | 符合（只作上限對照） | v5.26 |

### 2c. 有取數、但沒有任何公式引用（只顯示在「Tokenomics_取數」分頁）

`IF_RacksPerGW`、`IF_HoldAcct`、`IF_CapexFacility`、`IF_DeprIT`、`IF_DeprFac`、`IF_AvgDraw`、`IF_PowerPrice`、`IF_MaintFac`、`IF_OpexGW`（9 個；全部合規）。交接檔 2j 把 `IF_HoldAcct`、`IF_RacksPerGW` 列為「對照與參考線」，但程式中找不到引用，實際只顯示在取數分頁——交接檔的歸類小誤，不影響數字。

### 2d. 工具候選的人工確認

| 工具標出的候選 | 位置 | 判定 |
|---|---|---|
| `IF_Hdr` | `build_xlsx.py` 第 3328 列 | 誤判：說明文字「只引用 IF_（不含 IF_Hdr*）」 |
| `IF_Hdr` | `company.json` → `tokenomics._note` | 誤判：同上 |
| `IF_Hdr` | `data/tokenomics_names.txt` 第 1 列註解 | 誤判：同上 |

另以前綴搜尋（`DRV_`、`CAL_`、`B4_`、`B5_`、`CST_`、`IDX_`、`SRC_`、`Inputs!`、`DC_Cost!`、`Cap_In!`、`CTL_`）複查：`CAL_` 是 CoreWeave 自己的期間變數（`CAL_FYE` 等），`SRC_PRC_*`／`SRC_DC_003` 只出現在快照工具自動帶出的 L1 外部對照欄與 `permw_inputs` 的說明文字。**違規引用 0 筆；沒有直接引用 Tokenomics 內部工作表。**

**寫死數字（半違規，提醒）**：PR #27 的定價倍數基準值 `k_長約 0.76`（＝IREN–Microsoft 9.70 ÷ GB300 `IF_HoldEcon` 12.725）與 `k_現貨 1.76`（＝H100 指數 2.82 ÷ Hopper `IF_GPUhrEcon` 1.605）是用 Tokenomics 數字手算後打進 company.json。v5.29 這兩個分母沒變，所以目前無誤差；但 Tokenomics 下次改持有成本時，k 不會自動跟著變（證據表會重算，基準輸入不會），價格會被成本拖著走。建議加一條「k 基準＝證據表即時比值」的檢查（見第七節）。

---

## 三、口徑問題（共用工作單第 2 節）

**問 1：收入按什麼計價？利用率與簽約率各用哪一個？有無重複扣減？**
- 計價：按 MW-year／GPU 小時（神雲）。main v4.6：每 MW 年收入是輸入值（舊方法）；PR #27：Tokenomics 持有成本 `IF_HoldEcon` × 定價倍數 k。
- 簽約率：用公司自己的設定——「可計費 MW ÷ 已驗收 MW」（情境表）、「利用率」95%→92%（W1 已註明其實是長約售出比例，不是技術利用率）、「未被 RPO 覆蓋收入的填補率 fill」100%→80%。
- Tokenomics 端：沒有用 `IF_Util`（只在 W1 對照表並列並註明口徑不同）；`IF_HoldEcon`、`IF_GPUhrEcon` 都是 100% 時數口徑，所以**與 Tokenomics 之間沒有重複扣減**，符合契約第 3 節。
- 要判斷的一點：模型內部的「利用率」與「fill」都在處理「產能沒賣出去」，未被 RPO 覆蓋的那部分收入同時被乘了兩次（收入算式：可計費 MW × 每 MW 收入 × 利用率，超出 RPO 的部分再 × fill）。FY28–FY30 fill 為 90%／85%／80%，兩者若意義重疊即為重複扣減。這是 v4.5 以來既有設計，不是 Tokenomics 連接造成，但契約要求在交接檔寫明「採哪一種」，現行交接檔沒有這段。

**問 2：每 MW 收入是否仍用「成本加成與市場價 50/50 平均」？容量情境與單價情境是否綁定？**
- CoreWeave 從未用 50/50 平均（main：輸入值；PR #27：Tokenomics 錨 × k），**符合契約第 4 條「停用 50/50」**。
- 選擇器分離：情境選擇（保守 4.2／基準 5.6／積極 8 GW）只改產能路徑、計費比例與租賃；單價由 k 與隨需占比決定，有各自的敏感度，不跟產能情境同向移動。Tokenomics 成本情境（低／基準／高）切換時，k 會依成本比例重算，使市場價格不變。**符合契約第 4 條**，可作 Nebius／Oracle／WhiteFiber 的範本。
- 與契約公式的差別（需判斷）：契約寫的是「每 MW 收入＝q（每 MW token 產出，跨世代用 `L1_TokMW_Gen_ratio_*`）× p（每 token 價格，SRC_Price）× c × 簽約率」；PR #27 是「持有成本 × k」，世代之間的升級幅度跟著**成本**走，不是跟著**產出**走。以 v5.29 數字：VR200 對 GB300 的每 MW 持有成本比為 1.003，產出比為 1.565。換句話說，PR #27 隱含「VR200 每 token 價格比 GB300 低約 36%、每 MW 收入持平」，正好等於契約定義的「中性情境 q × p ≈ 1」。結論：基準一致，但樂觀（q × p > 1）與保守（q × p < 1）情境目前沒有對應的設定。

**問 3：模型商的 `IF_Alloc*` 六格**
- 不適用：CoreWeave 沒有引用任何 `IF_Alloc*`，也沒有引用每則提示 token 數或推論支出相關名稱。v5.29 這六格的變動對 CoreWeave 輸出沒有影響。
- 間接相關的 `IF_RevGWFleet`（PR #27 上限檢查用）與 `L1_RevGW_Fleet_VR200`（W1 對照）在 v5.29 也沒有變。

**問 4：公司特有變數是否只在本模型覆寫、沒有回寫 Tokenomics？**
- 是。WACC（CoreWeave 11%；`IF_HoldEcon` 內含 Tokenomics 的 10%）、合約價與 k、隨需占比、計費比例、利用率、交付進度（MW 路徑）、世代組合都在 company.json；模型只讀版本固定的快照，沒有任何程式寫入 Tokenomics。W1 對照表也只輸出到 CoreWeave 自己的 `docs/plan/`。
- 注意：CoreWeave 評價用 11%，收入錨的打平租金用 Tokenomics 的 10%，兩者不同是刻意的（錨代表市場打平價，不是 CoreWeave 自己的資金成本），建議在交接檔寫一句。

---

## 四、缺口與名稱清單草稿（共用工作單第 3 節）

### 4a. Tokenomics 缺的產業級數據（不自行建表；走 G2 進 DB_Evidence）

| 缺什麼 | 用途 | 目前以什麼代替 |
|---|---|---|
| 神雲長約價（每 MW 或每 GPU 小時；分世代） | 收入錨的 k_長約 | PR #27 company.json 證據表：IREN–Microsoft GB300（9.70，契約已引用）、IREN–NVIDIA（11.33）、Oracle–OpenAI（13.33，未證實）。v5.29 新增的 SRC_Price 分層欄（隨需／預留／長約）仍空白 |
| 市場現貨租金指數（H100、B200、B300） | k_現貨與區間 | company.json：Silicon Data 指數（2.82／5.87／6.82）、Ornn B200（4.08） |
| 神雲短天期合約價 | 隨需敏感度（W5） | CoreWeave 新聞稿：3–6 個月合約約 $40m／MW／年（公司自述） |
| VR200、Rubin Ultra 的任何合約價 | 新世代收入 | 找不到；以成本比例外推（k 不變） |
| 保固期與爬坡期的營運成本曲線 | Q2 實際每 MW 營運成本 0.88 對模型 2.09 的差距 | Tokenomics 只有滿載穩態；W5 以公司實際比例作敏感度 |
| 續約價格衰退（GPU 租賃，不是 token） | 舊合約到期後的價格 | 未做（交接檔已列為範圍外）；Tokenomics 的 L 係數是 token 價格口徑，不能直接套 |
| Rubin Ultra 的 `L1_TokMW_Gen_ratio` | 若改用 q × p 的世代升級 | 無（目前只到 VR200） |

### 4b. 建議的 `data/tokenomics_names.txt` 草稿（附在此，不執行）

```
# CoreWeave 引用的 Tokenomics 名稱（契約 v0.1：只限 IF_／IFW_／IFC_／L1_）
# ── 模型採用 ──
IF_CapexIT
IF_CapexTotal
IF_FacilityGW
IF_DeprLifeIT
IF_PowerCost
IF_MaintIT
IF_StaffSW
IF_TaxIns
IF_GPUsPerGW
IF_HoldEcon                  # 收入錨（PR #27）
IF_GPUhrEcon                 # k_現貨分母（PR #27）
# ── 上限檢查：取代 IF_RevGWFleet，四層同取（需工具支援 IFW_）──
IFW_RevGWFleet_100
IFW_RevGWFleet_Util
IFW_RevGWFleet_Prod
IFW_RevGWFleet_Life
# ── 契約第 3、4 節給神雲的介面（對照：世代產出比 q 與每 MW 成本底線）──
L1_TokMW_Gen_ratio_GB200
L1_TokMW_Gen_ratio_GB300
L1_TokMW_Gen_ratio_VR200
L1_HoldEconMW_GB300          # 與 IF_HoldEcon 同值 ÷ 1000（單位標示見第七節第 6 項）
# ── 用途欄（需工具支援 IFC_）──
IFC_Use
IFC_Conf
# ── 對照（保留）──
IF_Util
L1_FacCapexMW
L1_GPUhr_GB200_vsCW
L1_GPUhr_GB300_vsBE
L1_RevGW_Fleet_VR200
# ── 目前沒有公式引用，保留只為取數分頁透明；可刪 ──
IF_RacksPerGW
IF_HoldAcct
IF_CapexFacility
IF_DeprIT
IF_DeprFac
IF_AvgDraw
IF_PowerPrice
IF_MaintFac
IF_OpexGW
```

---

## 五、v5.29 重取影響（共用工作單第 4 節）

執行方式：
- `python3 tools/tokenomics/import_tokenomics.py --tokenomics /home/claude/tk-master --names <PR #27 的 names.txt（25＋IF_RevGWFleet）> --out /tmp/coreweave_v529_snapshot.json` → 成功，26 個名稱、缺漏 0，讀快取值（不需重算），約 11 秒。
- `--check /tmp/coreweave_v529_snapshot.json`：通過（26 名稱，相對誤差 ≤ 1e-9）。
- `--check coreweave/data/tokenomics_snapshot_v5.26.json`：通過（用 Tokenomics archive 裡同名的 v5.26 檔比對，25 名稱）。
- 新舊快照逐值比較（每個名稱 × 世代 × 低／基準／高成本、L1 的區間與外部對照欄、儲存格位置）：

| 比較 | 名稱數 | 數值不同 | 位置改變 | 新增／消失 |
|---|---|---|---|---|
| main v5.26 → v5.29 | 25 | **0** | 0 | v5.29 多了 PR #27 才用的 `IF_RevGWFleet` |
| PR #27 v5.27 → v5.29 | 26 | **0** | 0 | 0（整個項目內容逐字相同） |

- `IF_Alloc*`：CoreWeave 未引用，不適用。
- `IFW_*_Prod`／`_Life` 相對目前所用層：CoreWeave 唯一相關的是 PR #27 的上限檢查（目前用 `IF_RevGWFleet`＝`IFW_RevGWFleet_Util`）。v5.29 四層（US$m／IT MW／年，基準）：

| 世代 | _100 | _Util（現用） | _Prod | _Life | _Prod ÷ 現用 |
|---|---|---|---|---|---|
| Hopper H100 | 6.68 | 4.00 | 2.24 | 4.00 | 0.56 |
| GB200 NVL72 | 33.62 | 20.17 | 15.29 | 20.17 | 0.76 |
| GB300 NVL72 | 43.78 | 26.27 | 19.23 | 26.27 | 0.73 |
| VR200 NVL72 | 81.21 | 48.72 | 38.12 | 48.72 | 0.78 |
| Rubin Ultra | 114.88 | 68.93 | 54.42 | 68.93 | 0.79 |

  影響（以 PR #27 報告的上限比例 FY26–FY30＝42%／35%／29%／25%／23% 推算）：改用 `_Prod` 時分母縮小約 22–44%，FY26 約升到 57%（以期初世代組合近似），**超過 50% 門檻**；改用 `_100` 並乘 CoreWeave 自己的利用率時約降到 26%。`_Life` 在基準下與 `_Util` 相同（Tokenomics 的 L × m 套在 Util 層、基準值為 1），換用不改數字。
- **目標價影響：0**（所有進入計算的值都沒變；上限檢查只是檢查列，不影響數字）。

---

## 六、本輪做了什麼

1. 讀 repo README、`coreweave/CLAUDE.md`、`README`、交接檔 v4.6、W1–W5 工作單（W5 在 PR #27 分支）、兩份 10-08 盤點報告與 Tokenomics 契約、v5.29 報告。
2. 跑 `link_audit.py`，逐筆確認候選；另以前綴搜尋複查。
3. 逐一追蹤 25 個名稱在 `build_xlsx.py`、`segA.js`、`scripts/` 中的實際用途（採用／對照／未引用）；讀 PR #27 的 company.json 與 builder 了解 W4／W5 的新用法。
4. 以 openpyxl 讀 v5.29 Interface 的 `IFC_Use`／`IFC_Conf`／`IFC_Layer` 欄與 `IFW_`、`L1_HoldEconMW_*`、`L1_TokMW_Gen_ratio_*` 值（官方工具不支援這兩個前綴）。
5. 產生 v5.29 暫存快照並與 v5.26、v5.27 快照逐值比較。
6. 未改任何模型檔；未合併 PR #27。

---

## 七、待 chat 端判斷

依優先順序：

1. **【高】PR #27 的收入上限檢查改取四層瀑布，並決定比較哪一層。** 選項：(a) 維持 `_Util`（60% 利用率層，現況 42%，通過）；(b) 改 `_Prod`（生產降額層，FY26 約 57%，警告）；(c) 分母用 `_100` × CoreWeave 自己的利用率（約 26%）。分子是「CRWV 每 MW 計費收入（已 × 利用率 95%）」，與分母口徑最接近的是 (c)；最保守的是 (b)。建議：四層都取、並列顯示，門檻判斷用 (b) 或 (c) 由 chat 端決定。可併入 PR #27 的 r3 或合併後另開工作單。
2. **【高】先升級共用工具 `import_tokenomics.py` 支援 `IFW_`（四層）與 `IFC_`（用途欄，整欄文字）。** 不升級，任何公司都無法照契約第 2 條第 2 項取數。屬共用工具，建議由 chat 端另開共用工作單。
3. **【中】換釘 v5.29。** 數字 0 變動，只改 company.json 的 snapshotFile／version／commit 與快照檔；建議在 PR #27 合併後、與第 1 項一起做（避免 PR #27 再重建一次）。
4. **【中】「利用率」與「fill」是否重複扣減未售出產能**（第三節問 1）。這是 v4.5 既有設計；若判定重疊，FY28–FY30 收入會上修。另請決定在交接檔寫明「採商業簽約率口徑、不用 IF_Util」（契約第 3 節要求）。
5. **【中】世代升級跟成本走還是跟產出走**（第三節問 2）。PR #27 的成本錨等於契約的中性情境 q × p ≈ 1；是否加入 q × p > 1／< 1 的樂觀／保守敏感度（用 `L1_TokMW_Gen_ratio_*`），由 chat 端決定。
6. **【低】回報 Tokenomics 端的兩個觀察（只列，不在本模型處理）：**
   - `L1_HoldEconMW_*` 的值是 0.0127（＝每 GW ÷ 1,000），單位欄卻寫「$M/MW/年」；照數值應為「$B/MW/年」（或值應為 12.72，與契約第 4 節「GB300 12.72」一致）。下游若照單位讀會差 1,000 倍。
   - 營收列的四層瀑布不是逐層相乘：`_Life`＝`_Util` × L × m，不是 `_Prod` × L × m，所以基準下 `_Life`（4.00）高於 `_Prod`（2.24）。若本意如此，建議在契約寫明各層的基底。
7. **【低】k 基準值的一致性檢查**：在 PR #27 加一條檢查「company.json 的 k_長約／k_現貨基準＝證據表以目前快照算出的比值（容差 ±0.01）」，避免 Tokenomics 改持有成本後價格被拖動。
8. **【低】names.txt 整理**：9 個取數但未引用的名稱（第二節 2c）保留或刪除；交接檔 2j 的歸類小誤（`IF_HoldAcct`、`IF_RacksPerGW` 實際未引用）一併修正。
9. **【低】G2 證據提交**：第四節 4a 的長約價、現貨指數、短天期價格，建議由 chat 端整理進 Tokenomics DB_Evidence／SRC_Price 分層欄，之後 CoreWeave 改從 `SRC_` Active 紀錄讀，減少 company.json 自建證據表。
