# Oracle｜Tokenomics 連接審視（v5.29；2026-10-08）

- 分支：`claude/oracle-tk-link-review`（自 main `75327ef` 開）
- 最新提交 SHA：見 PR 說明（本檔內容提交後才產生雜湊）
- 報告檔：`oracle/docs/reports/20261008_oracle_tokenomics_link_review.md`
- 依據：共用工作單 `docs/workorders/20261008_tokenomics_link_review.md`、`oracle/docs/workorders/20261008_oracle_tokenomics_link_review.md`；Tokenomics 下游資料契約 v0.1（`docs/plan/Tokenomics_downstream_contract.md`）。
- 對照的 Tokenomics：master `a5061d9`，`model/CURRENT`＝`20261008_Tokenomics_v5.29.xlsx`（前提成立，第 4 節已做）。
- **本輪只審視、只回報**：沒有改 `company.json`、builder、Excel、`data/`；v5.29 快照只產生在暫存路徑，沒有提交。

## 結論（先看這段）
1. **Oracle 目前用 Tokenomics v5.24（`098873a`）的數字，全部是手抄進 JSON 的，沒有快照、沒有重現檢查。** 真正驅動模型輸出的只有四組：每 MW 年收入（11.62／17.40／24.20）、每 MW IT 建置成本（37.45／37.59）、PUE 1.2（MW 口徑換算）、每 MW GPU 數（487）。
2. **改用 v5.29 重取，這四組數字一個都不會變**（60 格逐格比對，差異都在兩位小數四捨五入內）；三情境目標價不受影響。v5.29 改變的 `IF_Alloc*` 六格 Oracle 沒有用到。
3. **方法面有三個與契約不符的地方**：(a) 每 MW 年收入仍是「成本加成與市場價 50/50 平均」，契約第 4 節已停用；(b) 容量（合約上限、併網速度）、建設延誤月數、每 MW 單價綁在同一個情境選擇器上同向移動，契約要求分開；(c) 收入端可能部分重複扣減（市場價路徑內含 85% 可計費率，模型又另外用「可計費 MW 收斂比例」）。
4. **新發現的警訊**：用 v5.29 新增的四層瀑布做上限檢查，Oracle 基準每 MW 收入 17.40 等於 GB300 機隊理論 token 營收（× 基準利用率層）的 66%、生產折減層的 90%；積極情境 24.20 超過生產折減層（126%）。以 VR200 計則為 36%／46%。也就是說，在 GB300 世代下，Oracle 收的租金接近「客戶用這些 GPU 能賺到的全部 token 營收」。
5. 違規引用（人工確認後）：內部頁或禁用前綴 **6 處**（`Inputs!` 2、`DC_Cost!` 3、`Cap_In!`／`CTL_PriceLife` 1），**全部不驅動輸出**（只作說明、參考或對照）；工具列出的 19 個名稱中，9 個 `IF_RevGW_*` 是誤判（只出現在一句說明文字）。

## 1｜取數點清單

「所用版本」除註明者外皆為 Tokenomics `model/20261007_Tokenomics_v5.24.xlsx`（master `098873a`，手工 JSON `oracle/data/permw_tokenomics_20261007.json`，2026-10-07 以 `bdb0de7` 確認未更新）。

| # | 本模型的格或參數 | Tokenomics 名稱 | 合規／違規 | 四層瀑布 | 是否符合 `IFC_Use` | 是否驅動輸出 |
|---|---|---|---|---|---|---|
| 1 | `scenarios.revMW`（0.01162／0.0174／0.0242 US$bn/MW；Excel「每 MW 年收入｜三情境」） | 路徑 A：`IF_HoldEcon`（GB300 基準 12.72）；路徑 B：`IF_GPUsPerGW`（487）；兩路平均 | 名稱合規；**取法違規**（手抄寫死、無快照）；**方法違規**（成本加成當收入、50/50 平均，契約第 4 節） | 路徑 A 為成本，無層；路徑 B 自乘 80／85／90% 可計費率（非 Tokenomics 的層） | `IF_HoldEcon` 用途「情境、相對比較、反轉門檻」，字面不衝突；但契約第 4 節限定成本加成只作 ROIC 檢查 | 是（OCI 營收、期初可計費 MW 校準） |
| 2 | `scenarios.capexTemplate.costMW`（37.45、37.59…；Excel「每 MW 建置成本」）；觸頂後汰換資本支出 | `IF_CapexIT`（GB300／VR200 基準） | 名稱合規；取法違規（手抄） | 不適用 | 符合（情境） | 是（成長與汰換資本支出） |
| 3 | MW 口徑換算 ÷1.2（`mwPath.connectedStart`、`contracted`、`pace`、站點表、`billableRatio` 說明） | `IF_FacilityGW`（GB300 基準 1.2） | 名稱合規；取法違規（手抄，散在多處說明與推導） | 不適用 | 符合（「基準輸入」） | 是（已連網 MW 路徑） |
| 4 | 路徑 B 每 MW GPU 數 487.0（VR200 291.7） | `IF_GPUsPerGW` | 名稱合規；取法違規（手抄） | 不適用 | 符合 | 是（經 #1） |
| 5 | 路徑 B VR200 換價比 1.674（寫成 `Interface!M14 ÷ J14`） | `IF_GPUhrEcon` | 以儲存格位址記錄（應記名稱）；取法違規 | 100%（每 GPU 小時持有成本） | 符合 | 否（三情境只用 GB300） |
| 6 | permw 檔 `holdAcct`、`capexTotal`、`powerCost` 欄 | `IF_HoldAcct`、`IF_CapexTotal`、`IF_PowerCost` | 名稱合規；手抄 | 不適用 | 符合 | 否（記錄） |
| 7 | permw 檔 `dcOpexPerMWit`（隱含 EBITDA 率檢核） | 記為 `DC_Cost!I48:N48` | **違規**（內部頁）；應改 `IF_OpexGW`（數值相同） | 不適用 | — | 否（檢核欄） |
| 8 | permw 檔 `itLifeYears`（6／6／4） | 記為 `Inputs!D23:F23` | **違規**；應改 `IF_DeprLifeIT`（數值相同） | 不適用 | — | 否（參考；Oracle 用自己的會計年限 6 年） |
| 9 | permw 檔 `wacc`（8／10／13%） | 記為 `Inputs!D28:F28` | **違規**；Tokenomics 沒有對應的 `IF_` 名稱（缺口，見第 3 節） | 不適用 | — | 否（參考；Oracle WACC 以 CAPM 自算） |
| 10 | permw 檔說明「資本回收係數 `DC_Cost!52`」 | — | **違規**（說明文字引用內部頁） | — | — | 否 |
| 11 | permw 檔 `gapDecomposition.e` 說明「`DC_Cost!J22` 儲存與管理 $500/GPU」 | — | **違規**（說明文字引用內部頁） | — | — | 否 |
| 12 | `oracle_facts` `pm.renewal.vs.tk`（續約價對照） | `CTL_PriceLife`（`Cap_In!C63`） | **違規**（`CTL_` 不在可取前綴；Tokenomics 也沒有 `IF_` 版本） | Life 層參數 | — | 否（只作對照） |
| 13 | permw 檔 `gapDecomposition.b`（世代成本對照，Nebius 用） | `IF_HoldEcon`（以 `Interface!D13、G13、J13、M13` 記錄） | 以儲存格位址記錄；手抄 | 不適用 | 符合（相對比較） | 否 |
| 14 | permw 檔 `utilization` 說明「`IF_Util` 60% 不沿用」 | `IF_Util` | 合規（只作對照、明文不進收入） | IF_Util | 符合 | 否 |
| — | 工具列出的 `IF_RevGW_Astra／Luna／Sol`（含 `_Prod`、`_Life`，共 9 個） | — | **工具誤判**：只出現在 `pm.renewal.vs.tk` 說明「`IF_RevGW_*` 不受影響」一句，模型沒有取用 | — | — | — |

- 人工確認後：工具列的 19 個名稱，實際被記錄或使用的是 10 個 `IF_`（`IF_GPUsPerGW`、`IF_FacilityGW`、`IF_CapexIT`、`IF_CapexFacility`、`IF_CapexTotal`、`IF_HoldAcct`、`IF_HoldEcon`、`IF_GPUhrEcon`、`IF_PowerCost`、`IF_Util`），其中驅動輸出的 4 個（#1–#4）；9 個 `IF_RevGW_*` 為誤判。
- 工具列的 3 筆違規（`Cap_In!`、`DC_Cost!`、`Inputs!`，各以檔案計）全部屬實；逐處展開為 6 處（#7–#12）。沒有「公司自己的同名工作表」造成的誤判（Oracle 的 Excel 分頁名稱為中文，未與 `Inputs!` 等重名）。
- 沒有取四層瀑布（`IFW_`）或用途欄（`IFC_`）：預期中（v5.24 沒有）。

## 2｜口徑問題

1. **計價方式：按 MW-year（容量）計價。** OCI 營收＝平均可計費 MW × 每 MW 年收入 × 期間長度；利用率欄 `defaults.m.util` 為 100%（不扣），`IF_Util` 不進收入（契約第 3 節要求，符合）。商業面的扣減有兩處：(i) 可計費 MW 從期初校準值向已連網 MW 收斂（50%→80%→85%→90%→90%）；(ii) 每 MW 年收入的市場價路徑內含 85%（80／90%）可計費利用率。**對以 take-or-pay 容量合約為主的 Oracle（OpenAI 等），(ii) 與 (i) 有部分重複扣減的疑慮**：若路徑 B 不乘 85%，基準每 MW 年收入由 17.40 升為 19.07（+9.6%），積極 24.20 → 25.87；保守（IREN 長約，本來就不乘）不變。需判斷（見末節）。
2. **每 MW 收入仍為 50/50 平均；容量交付與單價同向綁定。** 每 MW 年收入＝(成本加成 × 1.00／1.15／1.30 ÷ (1−6／8／10%)＋市場價) ÷ 2，契約第 4 節（X14 (j)）已停用此法。同一個情境選擇器同時決定：合約 MW 上限（6,284／8,778／8,778）、併網速度（1,000／2,833／3,333 MW-IT／年）、**v0.2 新增的延誤月數（6／3／0）**、每 MW 年收入（11.62／17.40／24.20）、可轉債上限，三者同向（保守＝少、慢、延誤長、單價低）。契約要求容量與價格分成兩條軸（CoreWeave W4、Nebius v0.2a 為範本）。補充：期初可計費 MW 以「Q1 OCI 年化 ÷ 每 MW 年收入」校準（2,543／1,698／1,221），所以單價情境在期初會被反向抵銷一部分（單價低 → 校準出的可計費 MW 多）；拆軸時這個校準應跟著價格軸走。延誤月數建議歸容量軸。
3. **`IF_Alloc*`：不適用**（Oracle 是雲平台，不是模型商）。模型沒有引用 `IF_Alloc*`；v5.29 這六格的改變（Q1 0.636→0.573、服務 GW 0.0519→0.0675 等）不影響 Oracle 任何輸出。Oracle–OpenAI 合約隱含每 MW 13.33 只作對照，也不受影響。
4. **公司特有變數只在本模型覆寫、沒有回寫。** WACC（CAPM 自算，β 1.77）、GPU 會計年限 6 年、合約 MW 上限、併網速度、延誤月數、未起租租約連動、可計費收斂比例、穩態 EBITDA 率 47% 全在 `company.json`；模型只讀 JSON，沒有任何寫回 Tokenomics 的路徑。附註：`IF_HoldEcon` 內含 Tokenomics 的經濟 WACC 10% 與 IT 壽命 6 年，與 Oracle 自己的 WACC 不同；若改以 `IF_HoldEcon` 為收入錨，這是「錨本身的口徑」，不需回寫，但 handoff 應寫明。

## 3｜缺口與名稱清單草稿

### Tokenomics 缺的產業級數據（只列出，不自行建表；由 chat 端走 G2 進 DB_Evidence）

| 名稱 | 用途 | 目前以什麼代替 |
|---|---|---|
| GB300／GB200 GPU 租價（隨需、預留、長約三層，含時間序列） | 每 MW 收入的價格 p、定價倍數 k 的證據 | permw 檔手蒐：Verda spot 5.21、24 個月預留 7.82、隨需 10.42；超大型雲牌價（Oracle 18.0、AWS 10.58、Azure 11.9、CoreWeave 10.5）。v5.29 已在 `SRC_Price` 新增五個分層欄，但目前留空 |
| 長約成交價（每 MW-IT·年） | k_長約、契約第 4 節的對帳 | IREN–Microsoft 9.70（2025-11，[Interested-party]） |
| VR200 租價 | VR200 世代每 MW 收入 | 以 `IF_GPUhrEcon` 比（1.674）由 GB300 換算（[Analogy]） |
| 非 GPU 服務（儲存、網路、軟體）營收占每 MW 收入比 | 每 MW 收入是否只算 GPU | 無（列缺輸入） |
| 壽命期價格係數 `CTL_PriceLife` 的 `IF_`／`L1_` 版本 | 續約價對照（Oracle 稱 +20%） | 直接引用 `Cap_In!C63`（違規，#12） |
| 經濟 WACC（`IF_HoldEcon` 內含 8／10／13%）的 `IF_` 版本 | 記錄錨的口徑 | 直接引用 `Inputs!D28:F28`（違規，#9） |

### `oracle/data/tokenomics_names.txt` 草稿（只列 IF_／L1_；未建檔、未執行於模型）

```
# Oracle 引用的 Tokenomics 名稱（只限 IF_（不含 IF_Hdr*）與 L1_；工具：tools/tokenomics/import_tokenomics.py）
IF_GPUsPerGW        # 每 GW GPU 數（路徑 B、k 換算）
IF_FacilityGW       # 每 IT GW 設施電力＝PUE（MW 口徑換算 ÷1.2）
IF_CapexIT          # 每 GW IT 資本支出（capexTemplate.costMW）
IF_CapexFacility    # 每 GW 廠房資本支出（對照：機房以租賃取得）
IF_CapexTotal       # 每 GW 資本支出合計
IF_HoldAcct         # 每 GW 年持有成本—會計
IF_HoldEcon         # 每 GW 年持有成本—經濟（收入錨）
IF_GPUhrEcon        # 每 GPU 小時經濟持有成本（k 的分母）
IF_PowerCost        # 每 GW 年電費
IF_Util             # 基準利用率（只作對照；按 MW-year 計價不進收入）
IF_DeprLifeIT       # 取代 Inputs!D23:F23
IF_OpexGW           # 取代 DC_Cost!I48:N48
IF_DeprIT
IF_DeprFac
IF_PowerPrice
IF_MaintIT
IF_MaintFac
IF_StaffSW
IF_TaxIns
IF_AvgDraw
L1_HoldEconMW_GB300        # 成本底線（契約第 3 節）
L1_HoldEconMW_VR200
L1_TokMW_Gen_ratio_VR200   # 世代產出比 q（契約第 4 節）
L1_GPUhr_GB300_vsBE        # GB300 每 GPU 小時持有成本 vs 新雲損益兩平租金
L1_RevGW_Fleet_VR200       # 1 GW 參考機隊付費營收（理想上限；只作上限檢查）
L1_FacCapexMW              # 每 MW 廠房資本支出（租賃對照）
```

另需（取數工具目前只收 `IF_`／`L1_`，`IFW_`／`IFC_` 要等工具擴充）：`IFW_RevGWFleet_100／_Util／_Prod／_Life`（上限檢查，四層同取）；上列 `IF_` 名稱對應列的 `IFC_Use`、`IFC_Conf`（記錄用途限制與信心等級）。

## 4｜v5.29 重取影響

- 以 `python3 tools/tokenomics/import_tokenomics.py --tokenomics /home/claude/tk-master --names <上列草稿> --out /tmp/oracle_v529_snapshot.json` 產生暫存快照：26 個名稱、缺漏 0；`--check` 重現檢查通過（相對誤差 ≤ 1e-9）。快照沒有提交。
- `--check` 只能比對「快照 vs 同一個 xlsx」；Oracle 沒有快照，所以改以 openpyxl 讀快照值，逐格對 permw 檔 v5.24 手抄值比較。

| 比較範圍 | 結果 |
|---|---|
| permw 檔 GB300／VR200 × 低／基準／高成本，10 個欄位（GPU 數、PUE、IT 資本支出、資本支出合計、會計與經濟持有成本、每 GPU 小時成本、電費、營運費用、IT 年限），共 60 格 | **0 格變動**（全部在手抄的兩位小數四捨五入內；最大相對差 1.2% 為電費 0.40 對 0.4047 的四捨五入） |
| `capexTemplate.costMW` 37.45／37.59 | 不變（v5.29：37.4455／37.5864） |
| PUE 1.2 | 不變 |
| `IF_Util` | 不變（60%；Oracle 不用） |
| 以 v5.29 重算三情境 50/50 每 MW 年收入 | 11.62／17.40／24.20，**與現用值相同** → 三情境目標價不變 |
| `IF_Alloc*` | Oracle 未使用，不受影響 |
| `IFW_*_Prod`／`_Life` 相對目前所用層 | Oracle 目前不取任何 token 營收列（收入不是由 Tokenomics 理論營收推得），沒有「目前所用層」；下表為新增的上限檢查 |

**新增的上限檢查（v5.29 四層瀑布，`IFW_RevGWFleet_*`，US$bn/GW＝US$m/MW；三種成本情境同值）**

| Oracle 每 MW 年收入 | ÷ GB300 _100（43.78） | ÷ GB300 _Util（26.27） | ÷ GB300 _Prod（19.23） | ÷ VR200 _Util（48.72） | ÷ VR200 _Prod（38.11） | ÷ `IF_HoldEcon` GB300（＝隱含 k） |
|---|---|---|---|---|---|---|
| 保守 11.62 | 27% | 44% | 60% | 24% | 30% | 0.91 |
| 基準 17.40 | 40% | **66%** | **90%** | 36% | 46% | 1.37 |
| 積極 24.20 | 55% | **92%** | **126%** | 50% | 64% | 1.90 |

- 讀法：`IFW_RevGWFleet_Util` 是「1 GW 機隊以 OpenAI 有效單價、基準利用率 60% 能收到的付費 token 營收」（理想上限）。Oracle 每 MW 租金占它的比例，就是 Oracle 從最終客戶 token 營收中拿走的份額。Nebius v0.2a 工作單訂的警示門檻是 50%：**GB300 世代下，基準與積極都超過**；VR200 世代下都在 50% 以內（積極剛好 50%）。Oracle FY27–FY28 在役以 GB200／GB300 為主，所以這個警訊是實質的。
- 對照：Oracle–OpenAI 合約隱含 13.33（MW 口徑不明）→ k≈1.05；IREN–Microsoft 9.70 → k≈0.76（契約第 4 節的對帳提醒）。

## 5｜建議改動（交 chat 端另開工作單；本單未做）
1. **每 MW 收入改以 Tokenomics 為錨並拆軸**（比照 CoreWeave W4、Nebius v0.2a）：每 MW 年收入＝在役世代加權 `IF_HoldEcon` × 定價倍數 k（證據＝市場價或合約價 ÷ `IF_GPUhrEcon`），停用 50/50 平均，成本加成只留作 ROIC／資金缺口檢查；容量軸（合約上限、併網速度、延誤月數）與價格軸（k 的低／基準／高）分開，加 3 × 3 目標價矩陣；期初可計費 MW 校準隨價格軸重算。
2. **取數改官方快照**：建立 `oracle/data/tokenomics_names.txt`（第 3 節草稿）與 `tokenomics_snapshot_v5.29.json`，`verify.sh` 加 `--check`；Excel 加「Tokenomics_取數」分頁與 `TK_` 名稱；`costMW`、PUE 1.2、GPU 數改讀快照；移除 `Inputs!`、`DC_Cost!`、`Cap_In!` 引用（舊檔標「已被快照取代」保留）。handoff 記錄 v5.29／`a5061d9`。
3. **加上限檢查與釐清利用率扣減**：每 MW 年收入 ÷ `IFW_RevGWFleet_Util`（並列 `_Prod`），按在役世代加權，超過門檻在檢查頁警示；決定路徑 B 的 85% 可計費率與「可計費 MW 收斂比例」是否重複扣減，handoff 寫明「按 MW-year 計價、用簽約率、不乘 `IF_Util`」。

## 待 chat 端判斷
1. **Oracle 的 k 證據組合**：Oracle 是超大型雲（客戶以 OpenAI 長約為主），不是 neocloud。是否沿用 W4 的 k_長約 0.76／k_現貨 1.76，或加入 Oracle–OpenAI 合約隱含（k≈1.05，MW 口徑不明）與超大型雲牌價（Oracle GB300 隨需 18.0 $/GPU-hr）作為 Oracle 專屬證據？
2. **延誤月數歸哪條軸**：建議歸容量軸（與合約上限、併網速度一起），價格軸只動 k；或延誤另成第三條軸？
3. **利用率重複扣減**：路徑 B 的 85% 可計費率在 take-or-pay 容量合約下是否應為 100%？影響：基準每 MW 年收入 17.40 → 19.07（+9.6%）。若改採 k 錨定法，此問題轉為「k 的證據是否已含可計費率」。
4. **上限檢查的門檻與世代**：GB300 世代下基準已占理論 token 營收（基準利用率層）66%、生產折減層 90%。門檻沿用 Nebius 的 50%？用 `_Util` 還是 `_Prod` 層？這是否應反映為下修基準單價的理由？
5. **回報 Tokenomics 端（不由本 repo 修正）**：
   - `L1_HoldEconMW_*` 的單位標籤寫「$M/MW/年」，但數值是 0.0127（實為 US$bn/MW，＝12.72 US$m/MW），差 1,000 倍，下游易誤用。
   - `IFW_RevGWFleet_*` 的標籤寫「理想上限」，但 `IFC_Use` 為「情境、相對比較、反轉門檻」、未標「上限」（X／Y 旗標皆 0），與契約第 2 條「上限不得作預測」的判讀不一致。
   - 共用取數工具只收 `IF_`／`L1_`，契約要求的 `IFW_`／`IFC_` 無法進快照，需擴充工具。
   - 缺 `CTL_PriceLife` 與經濟 WACC 的 `IF_`／`L1_` 版本（第 3 節）。
6. **MW 口徑換算**：Oracle 對外揭露的 MW 口徑不明，現以 ÷ PUE 1.2 換成 IT。改讀快照後是否仍取 `IF_FacilityGW` 基準 1.2，或改為區間（1.1–1.3）做敏感度？
