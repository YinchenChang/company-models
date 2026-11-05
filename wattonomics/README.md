# Wattonomics（瓦帳）— 全市場 AI 資料中心 GW 唯一基準事實

**Wattonomics（瓦帳）＝全市場 GW 唯一基準事實，與 Tokenomics（token）、Siliconomics（矽）成系列；舊名 GW 帳／gwledger。**

> **只給基準（v0.2 起）**：三層只給基準（Andy 2026-11-02 15:58）。基準＝上線時點「公司公告時程＋已確認延遲」；v0.2 起需求 B、C 與延遲 D1、D2 已全部刪除。新名稱 `WIF_MktGW_<YE|Avg|Add>_<yyyy>`（無情境字樣）；舊名 `*_A_D0_*` 同值保留（A_D0 為舊名，v0.3 刪）。
>
> **下游取用 2027 年後數值前看覆蓋率**（`WIF_Coverage_<yyyy>`＝有證據支撐的新增 GW ÷ 主路徑新增 GW；< 0.5 時 Checks WARN）。2026 起主路徑七家（MSFT、GOOGL、AMZN、META、ORCL、CRWV、NBIS）由資本支出推得（甲2，讀 In Y9 個股模型基準，付款年轉上線年），xAI 由 SpaceX 個股模型 AI 資本支出推得，「其他（企業、學研）」隨市場成長；純站點證據版另出 `_SiteOnly`。

算力收支 R4「產業部署路徑」的實作：市場總 GW、各玩家 GW（營運者／使用者）、部署管線與基建延遲情景的**唯一事實** Excel。每 GW 係數只讀 Tokenomics（TK）快照（v21.9 起快照已含 `IF_FacilityGW`；舊版 `builder/seed/tk_extra_v21.7.json` 只在快照缺名稱時使用）、加速器顆數與情景係數只讀 Siliconomics 快照，不重估兩者權威的量。

| 項目 | 內容 |
|---|---|
| 現行版本 | 見 `model/CURRENT` |
| 下游快照 | `data/CURRENT_SNAPSHOT` 指向的 `data/wattonomics_snapshot_v*.json`（`tools/export_snapshot.py` 產生；名稱前綴 `WIF_`） |
| TK 快照 | `wattonomics/data/tokenomics_snapshot_v21.10.json`（v0.3；先前 v21.9）`wattonomics/data/tokenomics_snapshot_v21.9.json`（v0.2 起自抽：`tools/tokenomics/import_tokenomics.py`，名稱清單＝Siliconomics 清單＋`IF_FacilityGW`；Siliconomics 仍為 v21.7） |
| Siliconomics 快照 | `siliconomics/data/CURRENT_SNAPSHOT` 指向者（v0.3 起 v0.13.1，另讀 R7 `SIF_NVContentGW_*`、逐買方 `SIF_nvsh_*`、`SIF_CapOnlineGW_Base`；只讀基準名稱 `*_Base`：`SIF_DeployGW_Base`、`SIF_AIITCapex_Base`、`SIF_UnitsCoef_Base`、`SIF_CapGW_Base` 等，不讀 A／B／C） |
| 規劃書 | `docs/plan/20261028_全市場GW帳_規劃.md` |
| 工作單 | `wattonomics/docs/workorders/` |
| 證據（查證代理產出，chat 端複核後才入 In） | `wattonomics/evidence_raw/`；`tools/import_evidence.py` → `builder/seed/evidence_v0.1.json`（DB_Evidence，附 C 段判定）；`tools/import_funding.py` → `builder/seed/funding_v<版>.json`（FundRaw、Facility，附 G／H／Q 段判定與更正） |
| 公開網頁（公司內可開） | https://yinchenchang.github.io/company-models/wattonomics/ （gh-pages 分支 `wattonomics/index.html`；由 `markets/invest_map/publish_pages.sh` 隨地圖一併發布，取 `model/CURRENT` 對應的 explorer）；私人版 claude.ai artifact 同內容 |
| 互動頁取數 | `data/explorer_v*.json`（`tools/export_explorer.py`；扁平紀錄，含標記與來源 ID） |
| AI半導體 次產業比例 | `markets/invest_map/data/segments_ai_semis_v0.2.3.csv`（SegIn 頁鏡像；Spend 拆分） |

## 建置（每個 Excel 新版本）

```bash
# 0) 證據有更新時：python3 wattonomics/tools/import_evidence.py; python3 wattonomics/tools/import_funding.py
# 0b) 個股模型更新且 In Y9 要重取時（v0.1.7 起；平常 In 為權威）：python3 wattonomics/tools/import_gap_models.py（逐家列位寫死並核對標籤，不符即停；models_extract.json 未更新的個股在工具內 MODEL_OVERRIDE 覆寫路徑），
#     之後 build.py 加 --refresh-y9（In Y9 改讀 builder/seed/gap_models_v<版>.json，其餘 In 仍讀 --base）
# 0c) In 頁輸入更正（例如 T2 已確認延遲 M 欄、延遲證據 P／Q 欄）：python3 wattonomics/tools/apply_in_edits.py <現行.xlsx> wattonomics/edits/v<版>_In_edits.csv <base.xlsx>，再以該檔作 --base
# 1) 產生（之後各版加 --base wattonomics/model/<現行>.xlsx，In 頁藍字讀回現行 Excel）
python3 wattonomics/builder/build.py --base <base.xlsx> --version vX.Y --date <YYYYMMDD> --out build/raw/<YYYYMMDD>_Wattonomics_vX.Y.xlsx   # --version 一律明寫
# 2) LibreOffice headless 重算存檔（錯誤格須為 0）
soffice --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir wattonomics/model build/raw/<檔名>.xlsx
# 3) 匯出快照並指向它
python3 wattonomics/tools/export_snapshot.py wattonomics/model/<檔名>.xlsx wattonomics/data/wattonomics_snapshot_vX.Y.json
echo wattonomics_snapshot_vX.Y.json > wattonomics/data/CURRENT_SNAPSHOT; echo <檔名> > wattonomics/model/CURRENT
# 3b) 互動頁取數
python3 wattonomics/tools/export_explorer.py wattonomics/model/<檔名>.xlsx wattonomics/data/explorer_vX.Y.json
# 4) 核對（錯誤格 0、Checks 全過〔WGOV_Errors＝0〕、In 每列有標記與來源、快照一致、本 README 含固定規格）
python3 wattonomics/tools/check.py wattonomics/model/<檔名>.xlsx wattonomics/data/wattonomics_snapshot_vX.Y.json
```

Gap 頁敏感度（lag_bldg、ai_sh 低／高；只在暫存目錄改輸入格並以 LibreOffice 重算）：`python3 wattonomics/tools/gap_sensitivity.py wattonomics/model/<檔名>.xlsx`

## 原則

- Excel 是唯一事實來源；`builder/` 只產生公式頁與格式，計算一律為 Excel 公式。`builder/seed/` 只用於 v0.1 首建。
- 輸入（In 頁藍字）由 Excel 擁有；每列有標記（Verified／Interested-party／Analogy／Assumed／Derived／待證據）、來源 ID（Sources 頁）、口徑（IT／設施／未說明）。Assumed 與 Analogy 純量一律給低／基準／高。
- 單位（規劃書 r1，Requests R1 裁定）：主口徑 **IT GW＝TK GW（IT 關鍵電力）**；設施電力＝IT × TK `IF_FacilityGW`（基準 1.2，1.1–1.3）。公司標設施者 ÷ `IF_FacilityGW`；口徑未說明者基準取 IT 讀法，Players 另列設施讀法；Epoch 晶片功耗 × `k_chip2it`（[Analogy] 1.65／1.85／2.08）。
- 季報滾動（Q 段，v0.1.3 起）：新一季財報列附在 F_funding.csv 末尾，交未參與代理開原文核對（`evidence_raw/Q_review_<季>.md`）後才列入 `Q_TRUE`；同期間較舊或已含在新期間現金流量表內的列（例如上半年各列、期內個別發債）列入 `Q_SUPERSEDED`（採用＝0，保留作事件紀錄）；`BUCKET` 的 2026 期間標籤隨之改（例如 2026H1 → 9M2026）。
- 證據：只用 `evidence_raw/C_review.md` 判「屬實」或已更正的列；「錯誤」「無法核對」「未抽查」的站點列在 In T2 但納入＝0。
- 站點時程補證據（J 段，v0.1.5 起）：Epoch 站點時間線的增量只取 2027 年以後的點（相對該站 2026 年底最後一點；2026 年底仍＝觀測＋T2），每點須經未參與代理逐頁重開核對（`evidence_raw/J_review_2027.md`）；發電容量、供電合約、晶片／算力合約不算站點容量；公司層級目標只作對照。覆蓋率報告同時列「七家逐家以主路徑為上限」的覆蓋率（`docs/verify/20261031_v0.1.5_coverage.md` 第 1 節）。
- 電力證據路徑（K 段，v0.1.6 起）：T2「納入」＝2 的列為監理機關已核准、指定單一園區的供電合約（只計確定 MW），不入 Players／Pipeline，只進 `WIF_Coverage_Power_<yyyy>`（參考用，Checks 不以此判定）；未核准合約列納入＝0。K4 乙（Andy 2026-11-02 18:36；v0.2 起）：SiteOnly 2030–2032 年新增直接加總站點證據（不再用前年 × g），主路徑 2030–2032 仍＝前一年年底 × g；覆蓋率 2030–2032 因此＝站點證據 ÷ 主路徑。
- 資金缺口（Gap 頁，規劃 `docs/plan/20261031_主要玩家資金缺口_規劃.md`；Andy 2026-10-31 規格）：成本＝AI 資料中心資本支出＝新增 GW × 每 GW 建置成本（土地、廠房、設備；TK `IF_CapexTotal`，與 Spend 頁供應商營收同一套；Andy 2026-10-31 10:22），每 GW 只在擁有者計一次，大玩家之間的租賃不重複計算；資金來源＝全公司營業現金流（本業＋AI）；不扣股票買回與股利（Andy 2026-10-31 04:20：資金不夠時可能減少，寬容處理；另列備忘）；輸出自籌比、缺口與所需融資（股權、債務、租賃與表外合資），並與已觀測融資對比；判讀「OpenAI 型」（持續需要大規模融資）或「Anthropic 型」（數年後轉正）。互動頁以「整個市場當一間公司」畫逐年「需要多少錢 vs 錢從哪來」與尚需新融資（10:22）；只用基準，逐家表含「新增 GW」欄，合計＝逐年數字（Andy 11-02 15:13）。v0.2 起承租玩家只算 IT＋自付／融資租賃廠房（見固定規格「承租玩家只算 IT」）。
  - Andy 原話（2026-10-31）：04:13「他們大玩家之間的租賃什麼的，就不要重複計算到了」「營收賺的不只是 AI 帶來的營收，還有…他們的本業」；04:20「那個股票買回跟股利就先不要算好了，寬容一點。畢竟如果他們真的資金不夠的時候，有可能會減少股票買回跟股利」；10:22「依照我們…追蹤到這個市場總共要蓋幾 G 瓦的 data center，然後每座 data center 要花多少錢來判斷」「以建設 data center 所需要的土地、廠房、設備，所需要花的錢作為依據，這樣才可以跟對得上那些供應商的營收」「把整個市場當做一間公司，他要蓋那麼多的 data center，需要多少錢？然後目前錢準備夠了嗎？」
  - chat 端定案（v0.1.7）：(a) 已簽未起租租賃與表外合資承諾屬市場內部移轉，只列備忘，不作資金來源；主讀法「累計仍未安排」＝累計尚需 − 期初現金。(b) 玩家有營業現金流、但無 GW 證據致 AI 資本支出為 0 的年份，改用個股模型 AI 資本支出並標示。(c) 並列讀法 `*_ModelCapex`：七家 2027–2030 改用個股模型 AI 資本支出。(d) 模型無非 AI 拆分時，非 AI＝In Y2 資本支出（曆年化）×（1 − AI 占比），使七家需要量合計等於其資本支出。(e) 不扣買回與股利。(f) 玩家以自己的世代與晶片組合計每 GW 成本（與 CapexGW 由資本支出換算 GW 同一組成本，避免換算來回不一致）；r5（Andy 12:58）起市場其餘不進收支，只留參考列。
- 資金缺口的範圍（Andy 2026-10-31 12:58）：「我們的那個主要玩家其實就是不要放…中國公司」「就拿我們那十一家當做代表來計算就好了，這樣是比較單純一點」「但是呢，我們還是要保留中國還有其他廠商，因為他們的需求會影響到我們下游，例如 Siliconomics…還有更下面的個股的營收」「不然上面需求有寫那個其他，可是算現金流的時候又沒有，那其實也是很奇怪」。規則：Gap 收支只算 **11 家代表玩家（本專案覆蓋，不含中國）**——Microsoft、Alphabet、Amazon、Meta、Oracle、CoreWeave、Nebius、WhiteFiber、SpaceX／xAI、OpenAI、Anthropic；中國與其他廠商（其他新雲、主權 AI、OpenAI／Stargate SPV、企業學研）的需求保留在全市場 GW、Deploy、Spend 與 Gap 參考列（`WIF_Gap_MktRestMemo_<yyyy>`），供 Siliconomics 與個股讀取，不納入收支；v0.2 起只剩基準，情景比例 r 恆＝1（列保留以維持舊名）。互動頁（13:04）：區塊內只留一行 11 家占全市場比例。
- TK 或 Siliconomics 權威的量需要改值 → 停，寫進 Requests 頁。
- 三層對帳與更新原則（Andy 2026-10-31 11:18、12:43，於 AI半導體；算力收支 10-31 13:20 定稿採用）：原話節錄「整套体系从 Tokenomics 到 Wattonomics 到 Siliconomics 都用这样的做法…个别更新，然后交叉比对，然后再 sync」。三層分工：Tokenomics＝每 GW 成本與營收；Wattonomics＝全市場 GW 的收入與成本（基準層）；Siliconomics＝其上的零組件營收（原則上跟隨 Wattonomics）。規則：(1) 各層各自更新；(2) 每次升版對上一層與下一層跑對帳，逐年標「一致（容差內，預設 ±15%＝In `tolband`）／時差（對方尚未納入我方已送證據，附送件日期）／待診斷」；(3)「時差」必須對應一則已寄出的 `交流/` 證據送件，否則為「待診斷」；(4) 升版時在 `交流/` 註明已納入哪些對方送件；(5) 兩邊都更新後仍有落差 → 診斷錯在哪一方，錯的一方修正，涉及 Andy 觀點者交 Andy；(6) 公司公布實績優先於第三方估計，依賴方向不變（Tokenomics → Wattonomics → Siliconomics → 個股），反方向只送證據。Wattonomics 端：對 Tokenomics 是快照版本與係數一致檢查（Wattonomics 不重估每 GW 係數）；對 Siliconomics 用 Compare 頁與 `siliconomics/tools/r4_bridge.py` 的結果；Wattonomics 側殘差歸因用 `tools/r4_attrib.py`（讀 `model/CURRENT` 與 Siliconomics 側 r1 橋接活頁簿；輸出 `docs/verify/<日期>_R4_Wattonomics側歸因_r<n>.xlsx`，LibreOffice 重算；不改模型）。
- 下游只讀 Interface 的 `WIF_` 名稱（經快照）。
- 本 repo 為公開 repo：不放 Andy 的個人情境或先驗。

## 固定規格（Andy 2026-10-28；規劃書第 3、4 節；2026-11-02 增訂）

### 三層只給基準（Andy 2026-11-02）

- 原話（11-02 15:58）：「Tokenomics(1GW單位基準)，Wattonomics(tokoenomics的全市場版)及Siliconomics(Wattonomics的零部件版)都只要提供基準版的數據就好。情境只留在下游取用的個股。個股才根據自身狀況作情境分析與預估。Siliconomics對帳的時候，只對個股的基準情境加總。Tokenomics及Wattonomics也是照這個原則處理。」
- Wattonomics 基準＝公告時程＋已確認延遲（Pipeline 基準時點＝公告上線年月＋In T2 M 欄已確認延遲月）。v0.2 起需求 B、C 與延遲 D1、D2 的欄位、公式、Checks、`WIF_` 名稱（含 `WIF_ShippedNotOnline_*`、Gap 頁非基準衍生名稱、`WIF_Gap_PlayerCumGap_A_D1_*`）與 explorer 紀錄全部刪除。
- 名稱：新增無情境字樣別名 `WIF_MktGW_<YE|Avg|Add>_<yyyy>`；舊名 `*_A_D0_*` 同值保留（A_D0 為舊名，v0.3 刪；下游在讀）。Excel 標籤、README、Decisions 不寫需求代號，改「基準（公告時程＋已確認延遲）」（Andy 11-04 07:44）。保留三列名：Deploy「年新增 IT GW（毛額，未扣汰換〔汰換列 v0.2〕；不含新納入存量）」「年底上線 IT GW」、Compare「瓦帳 年新增 IT GW（淨；SiteOnly）」（AI半導體 `r4_bridge.py` 依列名讀取；B 欄改「基準」）。
- 資金缺口（Gap）、互動頁、判讀、累計只讀基準。

### 地位鏈與對帳方向（Andy 2026-11-02）

- 原話（11-02 16:22）：「3-nomics與個股的地位鏈應該是Tokenomics->Wattonomics->Siliconomics->覆蓋個股；對帳的時候則反過來: 覆蓋個股->Siliconomics->Wattonomics->Tokenomics。若對帳時差很大……則先用自己的數字，在較上游趕快更新後彼此對帳。」
- 地位鏈：Tokenomics → Wattonomics → Siliconomics → 覆蓋個股；對帳方向：覆蓋個股 → Siliconomics → Wattonomics → Tokenomics，只對基準情境加總。落差很大時各層先用自己的數字，較上游儘快更新後再對帳。

### 已確認延遲門檻（工作單 20261102 E12；判斷題，交 Andy，預設如下）

- 寫入基準（In T2 M 欄）的條件：**公司本身揭露，或兩家以上獨立可信媒體報導且公司未否認**；每筆須由未參與代理開原文核對（`evidence_raw/D_review_<日期>.md`），判「屬實」才寫入，並在 T2 P 欄填證據 ID、Q 欄填確認等級（Checks D-K2）。公司否認者（例：D#16 Oracle，Bloomberg 報導、Oracle 否認）只列 Delay 頁追蹤「報導、公司否認」，M＝0。公司公告時程本身屬 Interested-party。
- 逐站追蹤：Delay 頁與 `WIF_Delay_<玩家|Rest|Mkt>_<AnnGW|DelayedGW|AvgMo|Share>`；互動頁「延遲逐站追蹤」區塊。

### NVIDIA 資料中心營收對帳點（工作單 20261102 v0.1.9 F4–F8；Andy 2026-11-02 16:22）

- Compare 頁 F 節，2025–2027 逐年：(1) 上線年新增 IT GW（基準）→ (2) 出貨年 GW（上線年反推：出貨_t＝(1−L)·上線_t＋L·上線_t+1＋inv·(上線_t+1 − 上線_t)，L＝`lag_on` ÷ 12、inv＝`inv_on`）→ (3) 非 NVIDIA 占比（Players `nnv_mkt`，出貨年）→ (4) NVIDIA 出貨 GW → (5) 每 IT GW NVIDIA 內容（世代加權）→ (6) 隱含 NVIDIA 資料中心營收 → (7) NVIDIA 實績／指引（低、基準、高）→ (8) 比值與 `tolband` 判定 → (9) 反解所需每 GW 內容；另列 Siliconomics 欄（v0.2.2 起直接讀出貨年 `SIF_DeployGW_Base`，不再由 r1 上線值往返換算）。
- **只並列、不校準**：基準數值不因對帳點改變；帶外標「待診斷」，各層先用自己的數字，上游（Tokenomics 每 GW 內容）先更新再對帳（Andy 11-02 16:22）。NVIDIA 公司指引（FY2028 年增約 70%）屬 Interested-party，只作對帳點，不入基準。
- (5) 的推導：Siliconomics `nvc_gb_tk`（20.84）＝TK `IF_GPUsPerGW` GB300 × Siliconomics `nv_content`（每顆 NVIDIA 內容，GB300）；H100、GB200 以同一 NVIDIA 占機架價比 × TK 各自機架價（Analogy）；VR200、Rubin Ultra 用 GB300 值〔Assumed；TK 無權威分世代值〕；markets/chip_bridge v0.2.2 Bridge 第 44 列（AI半導體產品，非 TK）的 VR200 拆分 23.47（2026 HBM 價）／21.85（2027 HBM 價、NVIDIA 吸收）只列備忘（`WIF_Recon_NV_*Memo*`）；Requests R7 請 Tokenomics 採用權威分世代值。
- 名稱 `WIF_Recon_NV_<ShipGW|NVGW|ContentPerGW|ImpliedRev|ActualRev|Ratio|ReqContent>_<2025–2027>`；Checks N-K1（INFO，不計入 ERROR／WARN）。互動頁不加此節。
- v0.2 方法修正（工作單 20261104 V4；2027 時差來源診斷「現在可做」）：(2) 七家與 xAI 直接用付款年 GW（CapexGW），其餘玩家逐家以上線年鏡像式反推；(3)(4) NVIDIA GW 逐家（各家非 NVIDIA 占比），中國 NVIDIA 占比 0（2025–2027 一致）；v0.3 Y2／Y3：每 GW 內容改讀 R7 `SIF_NVContentGW_*`；非 NVIDIA 占比改讀逐買方 `SIF_nvsh_*`（金額份額 s 換 GW 份額＝s ÷（s＋（1−s）÷ c），c＝非 NVIDIA ÷ NVIDIA 每 GW IT capex），取代殘差法。（以下為 v0.2.1 紀錄）v0.2.1 W1：非 NVIDIA 占比改殘差法——逐家值（In S5 覆寫、中國 1）維持，其餘玩家共用殘差＝（市場 GW 占比 × Σ 出貨 − 逐家值玩家非 NVIDIA 出貨）÷ 其餘出貨，市場 $ 占比先換 GW 占比，殘差 < 0 設 0 並 WARN；只用於對帳點，AI半導體 提供逐買方份額後改讀；新增 2026 校準係數 λ、所需內容成長（2027÷2026）、模型內容成長、校準後比值，**判定以成長比較**（`WIF_Recon_NV_<Calib_2026|ReqGrowth_2027|ContentGrowth_2027|CalibRatio_2027>`）；營收口徑內容（GB300 24.88、VR200 26.13，不含 scale-out）與黃仁勳每 GW（**晶片口徑，Andy 11-04**；÷ `k_chip2it` 換成每 IT GW）只列備忘。Requests R7 收件人改 AI半導體。

### 對帳鏈（每一環都是一列「左值、右值、比、容許帶、狀態」）

瓦帳的單位一律用 **TK GW＝IT 關鍵電力**，另列設施電力（× TK `IF_FacilityGW`，基準 1.2）；公司公布值逐筆標記 IT／設施／未說明（規劃書 r1）。

| 環 | 恆等式或比對 | 左側 | 右側 | 主要口徑陷阱 | 版本 |
|---|---|---|---|---|---|
| C0 | 市場總 GW ＝ Σ 主要玩家 GW ＋ 其他 | 由下而上加總（兩讀：含中國／不含中國） | 第三方總量（Epoch 設施 ÷ IF_FacilityGW；McKinsey 第二列；只作佐證） | 重複計算：OpenAI 在 Oracle／Microsoft／CoreWeave 的容量、Anthropic 在 Google／AWS 的容量。以「營運者 × 使用者」矩陣處理，列和＝欄和＝總量 | v0.1 |
| C1 | 玩家資本支出 ＝ Σ 世代 新增 GW × TK `IF_CapexTotal`（NonNV 乘持有比）＋ 非 AI 資本支出 ＋ 時間差 | 瓦帳推算 | 個股模型 | 建物比晶片早 12–24 個月入帳；融資租賃與表外 SPV；租用容量不在租用方資本支出 | v0.1 |
| C2 | 加速器顆數 ＝ Σ 平台 新增 GW × 每 GW 顆數 | 瓦帳推算 | Siliconomics Units | ASIC 每 GW 顆數 TK 尚無；出貨年 ≠ 上線年 | v0.2 |
| C3 | HBM 位元 ＝ 顆數 × 每顆 HBM | 由 C2 帶出 | markets/memory 供給 | 位元供給先於需求兩年鎖定 | v0.2 |
| C4 | 電廠 GW ＝ 設施 GW ＋ 備援；併網 ＋ 自建發電 ≥ 上線 GW | 瓦帳推算 | 電力公司大負載合約、併網佇列、自建電廠公告 | 併網是延遲的主要來源 | v0.3 |
| C5 | 發電機組 ＝ 自建發電 GW 與備援需求 | 瓦帳推算 | 燃氣渦輪、往復式機組、燃料電池、柴油備援廠商訂單與產能 | 訂單積壓 ≠ 年交付量 | v0.3 |
| C6 | 電源與散熱 ＝ 新增 GW × 每 GW 內容（Siliconomics） | 瓦帳推算 | 台達電、光寶科、奇鋐、雙鴻、Vertiv 等營收與指引 | 份額與單價 | v0.4 |
| C7 | AI 投資地圖個股：個股模型的 AI 子列 ≈ 瓦帳切片 × 份額 | 瓦帳切片 | 各個股模型 `segments.json` 的 AI 子列 | AI 只是官方分部下的子列 | v0.4 |

- 容許帶預設 ±15%（[Assumed]，In `tolband`）；雙方皆 [Verified] 才判「對得上／偏差」，否則標「證據不足」（附帶內／帶外）。超出容許帶時列出可能原因，不自動調參數。
- 推導方向：公司公告 GW 屬 [Interested-party]。瓦帳同時算兩條：(甲) 公告路徑；(乙) 實體可支撐路徑＝MIN（CoWoS→顆數→GW、HBM、併網、發電機組、變壓器）。主輸出是兩者差距與最先綁定的瓶頸，不以公告值反推參數。v0.1 只有 (甲)。

### 承租玩家只算 IT（Andy 2026-11-04；工作單 20261104 V1、V1b）

- 原話（11-04 08:06）：「要，這才是Oracle真正付出的資本開支，其他的改租賃->營業現金流降低這條線」；（11-04 08:21）融資租賃「處理上應該就要走資本開支那條線」。
- 承租玩家（Oracle、CoreWeave、Nebius）的甲2 每 GW 成本與 Gap AI 需要＝IT＋（1 − 承租比例 ×（1 − 融資租賃占比））× 廠房（In S9：`lease_<p>`、`finl_<p>`）。**融資租賃走資本支出線**：融資租賃廠房算承租方資本支出（會計擁有者），其金額為資金來源③（同 Funding「新增融資租賃」類，2026 起；`WIF_Gap_FinLeaseSrc_<yyyy>`）；營業租賃廠房歸**出租方**列（`WIF_Gap_LessorFacility_<yyyy>`；含在市場其餘參考列與 Spend 備忘，Spend 合計不變），不入 11 家收支。
- 營業現金流：個股模型已扣租金（Oracle v0.4.2 Thesis 第 57 列 → CF!55），不另扣；融資租賃部分加回租金中的本金＝Σ（2026 起融資租賃廠房上線年金額）÷ `lease_term`（直線法，Assumed 15–20 年；`WIF_Gap_FinLeasePrincipalAddback_<yyyy>`），利息留在營業現金流。融資租賃本金償還屬籌資活動、不在營業現金流；日後若改用財報實際營業現金流，須另扣融資租賃本金。

### 供給上限（v0.2.2；校準原則第 2 條）

- v0.3 起主檢查：Deploy 毛額年新增 ≤ `SIF_CapOnlineGW_Base`（可支撐新增上線 IT GW，毛額；Checks X4-K1）；以下成本等值比較改為第二口徑（X4-K2）。
- 基準 GW＝需求側錨點，但不得超過供給產能（Andy 11-04）。Compare G 節：Wattonomics 出貨年 GW（全市場一律以上線年新增鏡像反推；F2 對七家與 xAI 則用付款年，若 G 節也改用付款年，2026 headroom 為 6.4%）× 成本等值係數（Wattonomics 每 IT GW IT capex ÷ `SIF_ITCapexPerGW`）對 `SIF_CapGW_Base`，逐年 headroom；超過時 Checks X4-K1 WARN。只檢查，不自動截斷（判斷題，預設不截斷）。限制：2026 年上限等於 Siliconomics 自己的需求（CapHead 0、HBM 利用率 1.0），2026 年只是與 Siliconomics 需求對帳，不是獨立的供給檢驗；成本等值係數以 $ 占比當 GW 權重、市場含中國而上限多半不含，兩項都使 Wattonomics 偏高（保守方向），各讀法下 2026 headroom 介於 6%–18%。

### 汰換（v0.2；工作單 20261104 V7）

- Deploy 主路徑汰換 IT GW＝t − `life_ret` 年的裝機量（年新增毛額）＋新納入存量攤提（v0.2.1 W2：2023 年底存量與 2024–2025 新納入存量平均攤到認列年與前 5 年共 6 個年份，即攤到認列年前後到期；站點或 In 有年份資料者以資料為準）；壽命 6 年〔TK 折舊年限〕，Assumed 5–7；世代別依該年 Players 世代組合與非 NVIDIA 占比。淨新增＝毛額 − 汰換（`WIF_MktGW_Retire_<yyyy>`、`WIF_MktGW_AddNet_<yyyy>`）。既有毛額列不改值、列名不變；R4 比較讀淨額或毛額由 AI半導體 決定。

### 已刪的情境（v0.2）

- 需求情景 A／B／C 與基建延遲情景 D0／D1／D2 的定義紀錄見 v0.1.9 版 README（git 歷史）；v0.2 起只剩基準（GL01-67）。延遲只以「已確認延遲」進基準（上節門檻），逐站追蹤見 Delay 頁。

## 頁面

| 頁 | 內容 |
|---|---|
| README | 命題、驅動 → 推導圖、頁面、版本 |
| In | S 純量（單位換算、容許帶、lag_bldg、g_lt、延遲 D1／D2、玩家最新觀測、合約與租賃、NonNV 覆寫、財年月）；Y 年度（玩家年底 MW、個股資本支出、非 AI 資本支出、使用者總量、第三方總量、世代組合）；T1 營運者→使用者分攤；T2 園區管線；v0.1.8：S7 Wattonomics 自有參數 `lag_on`（4.5 月；3–6）、`inv_on`（0.10；0.05–0.20）、`oth_gmult`（1；0–1.5）；T2 M 欄改稱「已確認延遲月」（進基準），新增 P 欄（延遲證據 ID）、Q 欄（確認等級／追蹤狀態）；v0.1.9：頁尾 S8（NVIDIA 資料中心營收 `nv_dc_cy25`、`nv_dc_cy26`、FY2027 共識 `nv_fy27_tot`、FY2028 指引年增 `nv_fy28_g`、資料中心占比 `nv_dc_sh`；Siliconomics r1 上線值 `sil_r1_on`），既有列位不動 |
| TK_In、SIL_In | 鏡像（綠字） |
| Players | 玩家 × 年 IT GW（甲：公告路徑；輸入→晶片換算→觀測→內插→園區延伸），低／高、年平均、年新增、新納入存量、設施 GW；規則見 Decisions GL01-03 |
| Matrix | 營運者 × 使用者（2025–2027 年底；In T1 逐筆分配），無揭露者在未分配欄 |
| Pipeline | 園區＋未分園區殘差；基準時點＝公告＋已確認延遲（In T2 M 欄）；v0.2 刪除 D1／D2 欄位與矩陣（欄位重排：N 基準上線年、O 有效納入、P 已確保電力） |
| Delay（v0.1.8） | 延遲逐站追蹤：逐家（11 家＋其餘＋全市場）公告上線 GW（T2 有效納入、公告上線年 ≥ 2026）、已確認延遲 GW、平均延遲月（IT GW 加權）、延遲占比；明細（T2 有延遲證據或 M＞0 的園區）；無 T2 列的候選追蹤表（D#16、D#19）；`WIF_Delay_*` |
| CapexGW | 甲2：七家資本支出（個股模型 annualEstimates／指引）× AI 占比（`ai_sh` 0.75／0.85／0.95，Amazon 0.68）÷ TK `IF_CapexTotal`（世代加權，NonNV × 持有比；建物提前 `lag_bldg`）＝年新增 IT GW；敏感度列；v0.1.8：七家改讀 In Y9（缺年用 Y2，標代入），t−1、t−2 缺值代入（缺 t−2 用 t−1、兩者皆缺用當年），付款年經 `lag_on`、`inv_on` 轉上線年，2026 起進主路徑；另列 xAI（SpaceX 模型 AI 資本支出 ÷ Gap 同式每 GW；E8）與其他（隨市場成長；E9），主路徑調整 2026–2029 |
| Deploy | 基準（主路徑、SiteOnly）；主路徑 2030–32 年底 GW 年增率由 2029 線性收斂到 `g_lt`（0.10，0.05–0.15）；SiteOnly 2030–32 直接加總站點證據（K4 乙）；v0.2 汰換列（世代別）與淨新增列；覆蓋率自 2026 起計算 |
| C0、C1、Compare | 對帳與並列；C1 七家 2027 起標「同源（不對帳）」，C1b 廣義口徑（資本支出含融資租賃＋尚未起租變化；Microsoft 改列營業租賃問題）；Compare 含 xAI 宣稱（不入主路徑）、AI 占比、lag 敏感度；v0.1.8 起「同源」自 2026 起（含 xAI）；v0.1.9：F 節 NVIDIA 資料中心營收對帳點（見固定規格） |
| Funding、FundRaw | 玩家 × 期間（2024、2025、2026 最新）營業現金流、資本支出（現金／含融資租賃）、自由現金流、債（公司債、定期貸款、GPU 擔保債、私募信貸）、淨負債變化、新增融資租賃、尚未起租、表外合資、股權、供應商融資、買回＋股利；資金來源比例兩種口徑：【支付順序】內部→租賃→股權→債→其餘（外部資金排在營業現金流之後，營業現金流足以支付時外部比例為 0）；【總額】期間外部資金全額（發債＋定期貸款＋GPU 債＋私募信貸＋新增融資租賃＋表外合資＋股權）÷ 資本支出，與營業現金流 ÷ 資本支出並列（可合計 > 100%）；互動頁預設總額口徑；循環融資 17 條。只用 G 段判屬實或已更正列 |
| Gap（v0.1.7） | 資金缺口（v0.2：承租玩家只算 IT＋自付／融資租賃廠房、出租方列、來源③融資租賃、營業現金流本金加回）：每 IT GW 建置成本（IT＝Spend 同式、廠房＝CapexGW 同式；土地、建築、配電、冷卻、發電機、IT 等項目展開）；玩家層 11 家代表玩家（本專案覆蓋，不含中國）× 2024–2030：AI 資料中心資本支出（新增 GW × 每 GW，擁有者口徑，廠房提前 `lag_bldg`）、非 AI 資本支出、營業現金流（個股模型基準，In Y9）、自籌比、缺口、買回＋股利備忘、已觀測融資與推算尚需、判讀；「11 家代表玩家當一間公司」（基準；v0.2 起情境已刪）：需要兩層（11 家 AI、11 家非 AI）、來源兩層（營業現金流、已觀測外部融資 2024–2025）、尚需新融資、累計尚需與累計仍未安排（扣期初現金；承諾只列備忘）；市場其餘只留一列參考；無 GW 證據的玩家年份改用個股模型 AI 資本支出；並列讀法「七家依個股模型資本支出」、判讀；敏感度 M1（每 GW 用市場平均）。判斷題 G1–G6 見 Decisions GL01-25 至 GL01-38 |
| Facility | 每 IT GW 廠房 12.67（TK）拆分：配電、備援發電機（1.6 × 1.0）、UPS、冷卻、建築、人工設計、其他專業服務、網路消防、土地、殘差；自建發電另列（1.44 GW × 1.6 $M/MW，× 自建比例 0.10／0.20／0.35）；主變 2.0 GVA 為判斷值；市場層＝主路徑年新增 × 每 GW |
| SegIn、Spend | AI半導體 次產業「GB300 對照 ÷ IT」比例鏡像；年 × 項目市場金額（IT 8 項＋廠房＋自建發電） |
| Interface | v0.2：`WIF_MktGW_<YE|Avg|Add>_<yyyy>`（新別名）、`WIF_MktGW_<Retire|AddNet>_<yyyy>`、`WIF_Gap_<LessorFacility|FinLeaseSrc|FinLeasePrincipalAddback>_<yyyy>`、`WIF_Recon_NV_<Calib|ReqGrowth|ContentGrowth|CalibRatio>_<yyyy>`；非基準名稱已刪。以下為 v0.1.x 紀錄：`WIF_Coverage_<yyyy>`；`WIF_Spend_<項目>_<yyyy>`、`WIF_Facility_<項目>_<yyyy>`、`WIF_Funding_<玩家>_<期間>_<項目>`；`WIF_MktGW_YE／Avg／Add_<A|B|C>_<D0|D1|D2>_<yyyy>`、`WIF_ShippedNotOnline_<A|B|C>_<D1|D2>_<yyyy>`、`WIF_PlayerGW_YE_<player>_<yyyy>`（主路徑），同名加 `_SiteOnly` 後綴＝只有 甲；2024–2032；v0.1.7 起 `WIF_Gap_*`（2024–2030）：`WIF_Gap_PerGW_<yyyy>`、`WIF_Gap_Item_<項目>_<yyyy>`、`WIF_Gap_Need／NeedOwn／Short／MktAdd／Ratio／Share11_<A|B|C>_<D0|D1|D2>_<yyyy>`、`WIF_Gap_CumNeed／CumUnarranged_<s>_<D>_<yyyy>`（2026–2030；仍未安排＝累計尚需 − 期初現金）、`*_MemoInclCommitted`（再扣承諾，只作備忘）、並列讀法 `*_ModelCapex_*`（七家 2027–2030 依個股模型 AI 資本支出）、`WIF_Gap_Fallback(Cells)_<yyyy>`（無 GW 證據的模型值替代）、`WIF_Gap_NonAI／OCF／Buyback／OCFMissing_<yyyy>`、`WIF_Gap_ExtObs_<yyyy>`（2024–2025）、`WIF_Gap_OpenCash`、`WIF_Gap_Committed`、`WIF_Gap_ExtObs2026Memo`、`WIF_Gap_BufferCum`、`WIF_Gap_Verdict(_ModelCapex)_<s>_<D>`（文字）、`WIF_Gap_MktRestMemo_<yyyy>`、`WIF_Gap_PlayerCumGap_A_D1_<player>`、`WIF_Gap_CumUnarranged_M1_<s>_<D>_<yyyy>`、`WIF_Gap_Player_<player>_<yyyy>_<OCF|CapexAI|CapexNonAI|Gap|SelfFund|Buyback|CapexAIModel|CapexAIGW|CapexAIMC>`、`WIF_Gap_PlayerCash／PlayerCommitted／PlayerVerdict_<player>`；v0.1.8：`WIF_Gap_Player_<player>_<yyyy>_AddGW`、`WIF_Gap_Add11_<yyyy>`（＝Σ逐家）、`WIF_Delay_<玩家|Rest|Mkt>_<AnnGW|DelayedGW|AvgMo|Share>`；v0.1.9：`WIF_Recon_NV_<項>_<2025–2027>`（21 名）與備忘 `WIF_Recon_NV_<ContentMemoCB26|RatioMemoCB26|ContentMemoCB27|RatioMemoCB27>_<yyyy>`（12 名）；非基準名稱說明欄標〔停用，v0.2 刪〕 |
| Checks | 錯誤格 0、Matrix 列和＝欄和＝總量、Pipeline D0＝Players（0.01 GW）、殘差恆等、延遲單調、In 標記與來源、`WGOV_Errors`＝0；v0.1.8：K17b（主路徑上線年新增 < 0，WARN）、G-K2 改（2026 起、含 xAI）、G-K11（Σ逐家 AddGW＝Add11）、D-K1–D-K3（Delay 恆等、M＞0 須有屬實證據、基準時點）；v0.1.9：N-K1（INFO 類，NVIDIA 對帳點 2026 比值是否在 ±tolband 內；不計入 ERROR／WARN）；v0.2：刪 K3、K6 延遲參數、K9、K18、K19、G-K3c；新增 C-K1（七家園區年底存量 vs 甲2 年底，WARN）、R-K1（汰換恆等式）；K8 改 OK |
| Decisions、DB_Evidence、Sources、Requests | 決策（GL01-01 起）、證據、來源 ID、對 TK／Siliconomics 的請求；v0.1.9：GL01-55 至 GL01-61、DB_Evidence N#1、SRC_GL_331–333、Requests R7 |

## 版本紀錄（摘要；完整見活頁簿 README 頁）

- 2026-11-05（不改版）：`tools/import_gap_models.py` 的 Anthropic `MODEL_OVERRIDE` 改指 `anthropic/dist/20261104_Anthropic收支模型_v0_1_9.xlsx`（#327 已移除 v0_1_8）；重跑 `--only ant` 與 `builder/seed/gap_models_v0.2.1.json` 逐項相同（Funding F07／F03 2026–2030 與 v0.1.8 逐位相同，只有 file／commit 欄不同），種子與成品不重建。

- v0.3（2026-11-05）：Siliconomics v0.13.1、Tokenomics v21.10（Y1）；R7 分世代內容（Y2）；逐買方 NVIDIA 份額（Y3）；供給上限主檢查改可支撐上線 GW（Y4）。基準 GW 不變。報告 `docs/verify/20261105_v0.3_build.md`。

- v0.2.2（2026-11-04）：Siliconomics 改讀 v0.13 `*_Base`（X1）；Deploy 不乘 Siliconomics 顆數係數（X2，地位鏈）；Compare F3 改讀 Siliconomics 出貨年（X3）；供給上限檢查 Compare G 節與 Checks X4-K1（X4，只檢查不截斷）。基準 GW 不變。報告 `docs/verify/20261104_v0.2.2_build.md`。

- v0.2.1（2026-11-04）：W1 對帳點非 NVIDIA 殘差法、W2 新納入存量汰換攤提、W3 Oracle 三園區 Epoch IT 容量、W4 Anthropic v0.1.8；基準 GW 不變。工作單 `docs/workorders/20261104_Wattonomics_v0.2.1.md`；報告 `docs/verify/20261104_v0.2.1_build.md`；R4 r4。

- v0.2（2026-11-04）：承租玩家只算 IT＋融資租賃走資本支出線（V1、V1b）、Oracle AI 占比讀個股模型拆分（V2）、Oracle 2028 園區分期與 Shackelford 1,400 MW IT（V3）、NVIDIA 對帳點方法修正（V4）、刪停用情境（V5）、K4 乙（V6）、汰換列（V7）、Tokenomics 快照 v21.9（V8）。工作單 `docs/workorders/20261104_Wattonomics_v0.2.md`；報告 `docs/verify/20261104_v0.2_build.md`；R4 r3 `docs/verify/20261104_R4_Wattonomics側歸因_r3.md`。

- v0.1.9（2026-11-02）：NVIDIA 資料中心營收對帳點與小修（工作單 `20261102_Wattonomics_v0.1.9_NVIDIA對帳點與小修.md`）——Compare 新節 F（2025–2027；只並列、不校準）與 Siliconomics r1 欄（r1 淨額、Wattonomics 毛額）；In 頁尾 S8；SIL_In 加 `nv_content`；新名稱 `WIF_Recon_NV_*` 33 名（含備忘 12 名）；Checks N-K1（INFO）；Requests R7（Tokenomics：每 IT GW NVIDIA 內容分世代）；核對意見 D1–D3 修正（GL01-19 文字、`tools/export_explorer.py` 資本支出說明與 `capex_gw` 的 `basis: "付款年"`、Pipeline D1／D2 輔助矩陣標題〔停用，v0.2 刪〕）。對帳點比值 2025／2026／2027＝0.911／0.939／0.653（2027 帶外，待診斷）。既有 `WIF_` 變動 0。報告 `docs/verify/20261102_v0.1.9_build.md`。
- v0.1.8（2026-11-02）：基準化與上游更新（工作單 `20261102_Wattonomics_v0.1.8_基準化與上游更新.md`）——三層只給基準：基準＝公告時程＋已確認延遲，其餘情境標停用（v0.2 已刪）；主路徑 2026 起七家一律甲2（讀 In Y9，缺年 Y2 代入；付款年經 `lag_on`、`inv_on` 轉上線年；t−1、t−2 缺值代入），xAI 2026 起以 SpaceX 模型 AI 資本支出推得，「其他（企業、學研）」2026 起隨市場成長；Y9 重取（Anthropic v0.1.7 自建＝0）；新頁 Delay 與 `WIF_Delay_*`；新名稱 `WIF_Gap_Player_<p>_<yyyy>_AddGW`、`WIF_Gap_Add11_<yyyy>`；已確認延遲本版 0 筆（待未參與代理核對）。基準年新增 2026–2030＝19.06／25.67／29.13／27.73／27.88 GW（v0.1.7：12.52／23.96／26.92／27.66／28.07）。既有 `WIF_` 1,613 名變動、新增 136 名。報告 `docs/verify/20261102_v0.1.8_build.md`；R4 r2 `docs/verify/20261102_R4_Wattonomics側歸因_r2.md`。
- v0.1.7（2026-10-31）：新增 Gap 頁（資金缺口）與互動頁「整個市場當一間公司」區塊（需求、延遲選擇器正下方）；In 新節 Y9（12 家個股模型基準值，`tools/import_gap_models.py`）；`WIF_Gap_*` 1,823 個新名稱（r5：11 家代表玩家）；Checks G-K1–G-K10；Gap 未改動任何既有 `WIF_` 名稱。另合併 Amazon、Meta 2026Q3 10-Q 證據（F#258–275，Q_review 第五節；F#77 由 F#258 取代），因此既有名稱只有 `WIF_Funding_amzn_2026_lnc` 變動（137.214 → 164.922）。報告 `docs/verify/20261031_v0.1.7_gap.md`。
- v0.1.6（2026-10-31）：K 段 2028H2–2032 證據、`WIF_Coverage_Power_<yyyy>`。
