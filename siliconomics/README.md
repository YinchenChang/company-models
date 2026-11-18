# Siliconomics — AI 半導體產業基準庫

AI 半導體 Project 的「唯一事實」Excel：加速器顆數、記憶體與 HBM、網通、CoWoS、載板、伺服器組裝、電源散熱、設備等**行業變數只有一個值**，各公司模型與市場層只讀這裡。規格比照 `YinchenChang/Tokenomics`；與 Tokenomics 相關的數據一律讀 Tokenomics、不自行重估。

| 項目 | 內容 |
|---|---|
| 現行版本 | 見 `model/CURRENT`（本分支：v0.21-trial2，2026-11-18，會改值的草案〔Rubin Ultra 重複計算修正〕；main：v0.20 r1；Wattonomics v0.3.6、Tokenomics v22.0、規則 v1.4.3） |
| 下游快照 | `data/CURRENT_SNAPSHOT` 指向的 `data/siliconomics_snapshot_v*.json`（由 `tools/export_snapshot.py` 產生） |
| Tokenomics 快照 | `data/tokenomics_snapshot_v22.0.json`（v0.20 起；名稱清單 `data/tokenomics_names.txt`；下游 chip_bridge 仍讀 `v21.7`，見 Checks K3） |
| 協調狀態檔 | `data/baseline_status.json`（`tools/export_baseline_status.py`；逐年狀態的唯一權威，規則 v1.4.3 C3；2026-11-18 起含 `method`、`generated`、`source_sha256`、`wattonomics_source_sha256`、`bridge`〔協調橋接檔名與雜湊〕與逐年 `timing_cum`／`timing_cum_ratio`） |
| Recon 快照 | `data/CURRENT_RECON` 指向的 `data/recon_snapshot_v*.json`（`tools/recon.py`；每市場 kind／status／殘差列，見檔頭） |
| 規劃書 | `docs/plan/Siliconomics_規劃書.md` |
| 下游契約 | `docs/Siliconomics_downstream_contract.md` |
| 變更紀錄 | `CHANGELOG.md` |

## 建置（每個 Excel 新版本）

```bash
# 1) 產生（之後各版加 --base model/<現行>.xlsx，In 頁藍字讀回現行 Excel）
python3 siliconomics/builder/build.py --out build/raw/<YYYYMMDD>_Siliconomics_vX.Y.xlsx
# 2) LibreOffice headless 重算存檔（錯誤格須為 0、Checks 的 SGOV_Errors 須為 0）
soffice --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir siliconomics/model build/raw/<檔名>.xlsx
# 3) 匯出快照並指向它
python3 siliconomics/tools/export_snapshot.py siliconomics/model/<檔名>.xlsx siliconomics/data/siliconomics_snapshot_vX.Y.json
echo siliconomics_snapshot_vX.Y.json > siliconomics/data/CURRENT_SNAPSHOT
#    快照記錄的 source.sha256 必須等於 model/ 內該檔的實際 SHA-256：匯出須是「定稿後的最後一步」——之後任何 LibreOffice 重算存檔（含 cascade 的 K1 自舉重建）都會改變檔案位元組，
#    須重跑本步驟（值不變時只有 sha256 一格不同）。`export_snapshot.py --check` 與 `tools/baseline_coordination/check.py` 第 [13] 項（CI）會擋住不一致（v0.20 曾發生：快照記 cf29d2bb…，檔案 ad8dfaaf…，原因＝快照匯出後 cascade 又重存活頁簿）
# 4) 對帳（v0.11 起每次升版都跑）：R4 橋接 → Recon 快照 → 以快照重建本版（Recon 頁）
python3 siliconomics/tools/r4_bridge.py build/r4_raw.xlsx && soffice --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir build build/r4_raw.xlsx
cp build/r4_raw.xlsx siliconomics/docs/reports/r4_bridge/<YYYYMMDD>_R4對帳橋接_Siliconomics側_r<N>.xlsx
python3 siliconomics/tools/recon.py siliconomics/docs/reports/r4_bridge/<該檔> siliconomics/data/recon_snapshot_vX.Y.json   # 第三參數預設 model/CURRENT（讀 Recon D1、Coord 狀態）
echo recon_snapshot_vX.Y.json > siliconomics/data/CURRENT_RECON   # 之後回到步驟 1 以 --base 重建一次
# 5) 狀態檔與協調檢查（v0.19 起）：python3 siliconomics/tools/export_baseline_status.py && python3 tools/baseline_coordination/check.py（須 0 項不符）
# 6) 下游全部重建（CI ai-semis-build；本機 tools/cascade.sh，並行時設 LO_PROFILE）：數值變動須全部落在本版預期範圍內。
#    K1（本庫 HBM 市場 vs SK 海力士成品）須先以暫用快照重建 SK 海力士再重建本庫（K1 自舉）
```

**顆數單位（v0.20 r1 加註）：** In 頁 M2 顆數（`u_<平台>_<年>`）、Units 頁、Interface 的 `SIF_u_*`、`SIF_UnitsPlat_*`、`SIF_Units_Total` 一律是**封裝（package）數，非晶粒數**（依據 `docs/reports/20261118_三層對帳/unit-basis/20261118_顆數來源口徑查核.md`：兩層顆數都是封裝數；Tokenomics `IF_GPUsPerGW` 為每 GW 封裝數）。引用 NVIDIA 2025 年「Blackwell 顆數」（GTC 2025、GTC DC，晶粒口徑）等外部數字前須先換算。In 頁顆數列的單位欄仍寫「百萬顆」（記憶體模型逐字重播 M2 節，改單位欄會改動下游），口徑寫在 In 頁 K 欄。

**只顯示的列（v0.20 r1；規則 v1.5 草案，待 Andy 逐條核可）：** Coord 頁 D 節「觀測數列」（中間估計 (i)(ii)(iii) 與顯示區間、衰減寫法 λ＝0／0.5／1；標「權重未校準」）、Recon 頁 G 節「季節權重讀數」（C0 第 11 點第三讀數；權重＝NVIDIA 資料中心營收 FY2024／FY2025／FY2026 上半年占比 30.7%／42.4%／41.4%，低＝最小、基準＝中位數、高＝最大）與其輸入 In 頁 M31。**不輸出 Interface 名稱、不進基準鏈與 `crd_k`、不進 Recon E 節彙總與 Coord `c0f`／`c0_fail`、不進任何個股計算**；`builder/build.py` 的 `_display_only_guard` 在建置時掃描全活頁簿公式與名稱，任何引用即中止。狀態檔逐年 `observations`（`display_only: true`）只供顯示。核對工具：`tools/seasonal_h1_reading.py`（重算季節讀數並與活頁簿逐格核對）。

對帳狀態（Recon 頁）：|比值 − 1| ≤ 容差＝一致；超出且有已寄出的 `交流/` 證據送件（登錄在 `builder/seed/recon_evidence_v*.json`）＝時差；其餘＝待診斷（Checks K9）。只加總覆蓋個股的**基準情境**（工作方式 v1.7）。
v0.20 起（規則 v1.4.3）：B 節（R4）狀態讀 Coord 頁、不自判；D 節由下往上市場（D1 GPU／ASIC 晶片金額）＝已建模個股加總＋殘差＝市場值，公司實績以主讀數（已公布期間累計 ÷ 層級同期間）與第二讀數（最近一季年化）判定，兩者同號超出 ±15% 才不一致、一內一外為時差邊界；F 節 NVIDIA 資料中心營收主讀數＝純實績，含指引的比值只列參考。甲年任一實績對帳點不過 → Coord 標「有條件一致」（`SIF_RecoC0Fail_<年>`）。

## 欄位語義（`SIF_DeployGW_*`；三層對帳工作單 A7，2026-11-18）

Interface E 節有五組部署 GW 欄位。**每一組只有一種語義**；寫新程式、新個股模型時只能照下表用，不得用「名稱聽起來像」來猜。

| 名稱 | 語義（唯一） | 2029 現值（GW） | 誰可以讀、怎麼用 |
|---|---|---|---|
| `SIF_DeployGW_Base`（向量，欄 E–H＝2026–2029） | **依規則的基準**（規則 v1.4.3 C3：甲／乙／丙／待定／連續法決定）。成本等值、出貨年 | 42.680 | 硬體個股**計算基準的唯一來源**；個股不裁決 |
| `SIF_DeployGW_Supply_<年>` | 供給側讀數＝原始顆數部署 GW（不乘任何係數；情境 A） | 60.139 | 只並列顯示與對照（個股 Coord 頁、對帳）；**不得進基準計算** |
| `SIF_DeployGW_Demand_<年>` | 需求側讀數＝Wattonomics 換算值（平滑法 x2，由 `tools/baseline_bridge.py` 重算） | 28.482 | 只用於 Coord 頁判定與狀態檔；個股**不得讀** |
| `SIF_DeployGW_Alt_<年>`／`SIF_DeployGW_AltRatio_<年>` | 另一側讀數（未被採為基準的那一條）及其 ÷ 基準 | 28.482／0.667 | 只用於對帳中年份的並列顯示與**兩讀數目標價（情境）**；不得取代基準 |
| `SIF_DeployGW_LegacyS33_<年>`（v0.20 r1 新增；Interface E2 節） | **舊取法**＝原始顆數 × 情境 B 顆數係數（舊 S33 係數；v0.18 基準）。**不是基準**；Coord 頁 A 節 `cur` 列「舊取法（原始 × 情境 B 顆數係數；非基準）」 | 51.118 | 只供「改前現行值」對照與中間估計 (ii)；**任何程式不得拿它算基準** |
| `SIF_DeployGW_Current_<年>` | **相容別名，下一版刪除**：值同 `SIF_DeployGW_LegacyS33_<年>`（舊名「Current」易被誤讀為現行基準）；2029 為 51.118，≠ Base 42.680 | 51.118 | 舊讀取者（`nvidia/coord_post.py`）改讀 `LegacyS33` 後刪除；新程式一律讀 `LegacyS33` |

**誰在讀（2026-11-18 全 repo 盤點；逐支用途見 `docs/reports/20261118_三層對帳_s-gov_A1A4A7A8.md` 第 2 節）：**

- `Current`：`nvidia/coord_post.py`（Coord 頁「改前現行值」對照欄；另有一條「讀不到 Base 就退回讀 Current」的路徑，已另開 PR 改成報錯）、`siliconomics/tools/export_baseline_status.py`（狀態檔 `legacy_s33`／`current_s33`，無下游讀取者）。
- `Demand`：`nvidia/coord_post.py`（註解與摘錄清單）、`export_baseline_status.py`（狀態檔 `converted`）。
- `Supply`：`foxconn`、`quanta`、`wistron`、`wiwynn` 的 `*_post.py`（Coord 頁「情境 A 供給側讀數」並列列）、`nvidia/coord_post.py`（並列與一致性核對）、`supply_evidence_ratio.py`（分母核對）、`export_baseline_status.py`。
- `Alt`／`AltRatio`：約 60 支個股 `*_post.py`／`c13.py`（並列顯示與兩讀數目標價）、`siliconomics/reader.py`、`tools/coord/coord.py`。
- **以上沒有任何一支用 `Current`、`Supply`、`Demand`、`Alt` 計算基準。**

**v0.20 r1（2026-11-18）已套用：** Interface 新增 `SIF_DeployGW_LegacyS33_<年>`（放在 Interface 最後的 E2 節，不移動既有名稱列位），`Current` 標籤改為「相容別名，下一版刪除」；狀態檔 `legacy_s33` 改讀新名。採「原地修補 v0.20、不升版」（理由與驗證見 CHANGELOG「v0.20 r1」）。

**改名方案（選風險較低者）：「新增新名＋舊名保留一版作別名」，不直接改名。** 理由：(1) 直接改名＝刪除 Interface 名稱，依下游契約屬主版變動，且需同步所有讀取端；(2) 讀取端雖只有 2 支，但舊快照（v0.19／v0.19-trial／`nvidia/coord_inputs.json` 的 trial 摘錄）仍含舊名；(3) 新增名稱屬次版、不影響既有讀取者。
**為何本 PR 不改活頁簿：** 新增名稱必須升版（連帶 SIL.VERSION 文字改變）→ 約 65 個下游模型須整批重建（`tools/cascade.sh` 兩輪），超出本單範圍與只動 `siliconomics/`、`tools/baseline_coordination/` 的邊界。已驗證的 builder 修補與驗證結果寫在報告第 3 節，等下一個 Siliconomics 版本一併套用。本 PR 做的是：語義寫入 README／契約、狀態檔讀取端先支援新名（有 `LegacyS33` 就讀它）、`tools/baseline_coordination/check.py` 第 [14] 項在 CI 擋住誤用。

## 原則（比照 Tokenomics CLAUDE.md）

- Excel 是唯一事實來源；`builder/` 只產生公式頁與格式，不得在程式內新增或省略計算機制。
- 輸入（In 頁藍字）由 Excel 擁有：要改值直接改 Excel；builder 重建時以 `--base` 讀回。`builder/seed/` 只用於 v0.1 首建。
- 下游只讀 Interface 的 `SIF_` 名稱（經快照）；不得連結 In、Mkt 等內部頁。
- Tokenomics 權威的量（每 GW 顆數、機架數、機架價、功率、每 GW 資本支出、NonNV 比例、產業部署路徑）只在 `TK_In` 頁鏡像；Tokenomics 升版時重抽快照、重建，Checks K3 列出下游仍用舊版的格數。
- 本 repo 為公開 repo：不得放 Andy 的個人情境、先驗或校準紀錄。

## 網頁版（展示層）

- 線上版：https://yinchenchang.github.io/company-models/siliconomics/ （GitHub Pages，公司內可直接開啟；Andy 2026-10-30 要求，比照 AI 算力地圖〔原名 AI 投資地圖，Andy 2026-11-16 改稱〕）。
- 產生：`python3 siliconomics/tools/build_html.py <輸出.html>`（讀 `model/CURRENT`）；`markets/invest_map/publish_pages.sh` 發布時自動產生，不在 repo 另存 HTML。數字全部取自重算後的 xlsx，程式不做計算；每格 title 標註來源儲存格。
- Siliconomics 升版合併後，重跑 `publish_pages.sh` 即更新線上版。

