# GW 帳（GW Ledger）— 全市場 AI 資料中心 GW 唯一基準事實

算力收支 R4「產業部署路徑」的實作：市場總 GW、各玩家 GW（營運者／使用者）、部署管線與基建延遲情景的**唯一事實** Excel。每 GW 係數只讀 Tokenomics（TK）快照、加速器顆數與情景係數只讀 Siliconomics 快照，不重估兩者權威的量。

| 項目 | 內容 |
|---|---|
| 現行版本 | 見 `model/CURRENT` |
| 下游快照 | `data/CURRENT_SNAPSHOT` 指向的 `data/gwledger_snapshot_v*.json`（`tools/export_snapshot.py` 產生；名稱前綴 `GIF_`） |
| TK 快照 | `siliconomics/data/tokenomics_snapshot_v21.7.json`（與 Siliconomics 同一份，不另抽） |
| Siliconomics 快照 | `siliconomics/data/CURRENT_SNAPSHOT` 指向者 |
| 規劃書 | `docs/plan/20261028_全市場GW帳_規劃.md` |
| 工作單 | `gwledger/docs/workorders/` |
| 證據（查證代理產出，chat 端複核後才入 In） | `gwledger/evidence_raw/` |

## 建置（每個 Excel 新版本）

```bash
# 1) 產生（之後各版加 --base gwledger/model/<現行>.xlsx，In 頁藍字讀回現行 Excel）
python3 gwledger/builder/build.py --out build/raw/<YYYYMMDD>_GWLedger_vX.Y.xlsx
# 2) LibreOffice headless 重算存檔（錯誤格須為 0）
soffice --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir gwledger/model build/raw/<檔名>.xlsx
# 3) 匯出快照並指向它
python3 gwledger/tools/export_snapshot.py gwledger/model/<檔名>.xlsx gwledger/data/gwledger_snapshot_vX.Y.json
echo gwledger_snapshot_vX.Y.json > gwledger/data/CURRENT_SNAPSHOT; echo <檔名> > gwledger/model/CURRENT
# 4) 核對（錯誤格 0、Checks 全過〔GGOV_Errors＝0〕、In 每列有標記與來源、快照一致、本 README 含固定規格）
python3 gwledger/tools/check.py gwledger/model/<檔名>.xlsx gwledger/data/gwledger_snapshot_vX.Y.json
```

## 原則

- Excel 是唯一事實來源；`builder/` 只產生公式頁與格式，計算一律為 Excel 公式。`builder/seed/` 只用於 v0.1 首建。
- 輸入（In 頁藍字）由 Excel 擁有；每列有標記（Verified／Interested-party／Analogy／Assumed／Derived／待證據）、來源 ID（Sources 頁）、口徑（IT／設施／未說明）。Assumed 與 Analogy 純量一律給低／基準／高。
- 單位：主口徑 **TK GW＝設施電力**（VR200 每 TK GW 含 919.8 MW IT）；IT 讀法 ÷ g。口徑未說明者基準取 IT 讀法，Players 另列設施讀法。
- TK 或 Siliconomics 權威的量需要改值 → 停，寫進 Requests 頁。
- 下游只讀 Interface 的 `GIF_` 名稱（經快照）。
- 本 repo 為公開 repo：不放 Andy 的個人情境或先驗。

## 固定規格（Andy 2026-10-28；規劃書第 3、4 節）

### 對帳鏈（每一環都是一列「左值、右值、比、容許帶、狀態」）

GW 帳的單位一律用 **TK GW（設施電力；VR200 每 GW 含 919.8 MW IT，其他世代的 IT 占比 v0.1 查證）**，另列 IT GW；公司公布值逐筆標記 IT／設施／未說明。

| 環 | 恆等式或比對 | 左側 | 右側 | 主要口徑陷阱 | 版本 |
|---|---|---|---|---|---|
| C0 | 市場總 GW ＝ Σ 主要玩家 GW ＋ 其他 | 由下而上加總 | 第三方總量（Epoch、IEA、SemiAnalysis 等，只作佐證） | 重複計算：OpenAI 在 Oracle／Microsoft／CoreWeave 的容量、Anthropic 在 Google／AWS 的容量。以「營運者 × 使用者」矩陣處理，列和＝欄和＝總量 | v0.1 |
| C1 | 玩家資本支出 ＝ Σ 世代 新增 GW × TK `IF_CapexTotal`（NonNV 乘持有比）＋ 非 AI 資本支出 ＋ 時間差 | GW 帳推算 | 個股模型 | 建物比晶片早 12–24 個月入帳；融資租賃與表外 SPV；租用容量不在租用方資本支出 | v0.1 |
| C2 | 加速器顆數 ＝ Σ 平台 新增 GW × 每 GW 顆數 | GW 帳推算 | Siliconomics Units | ASIC 每 GW 顆數 TK 尚無；出貨年 ≠ 上線年 | v0.2 |
| C3 | HBM 位元 ＝ 顆數 × 每顆 HBM | 由 C2 帶出 | markets/memory 供給 | 位元供給先於需求兩年鎖定 | v0.2 |
| C4 | 電廠 GW ＝ 設施 GW ＋ 備援；併網 ＋ 自建發電 ≥ 上線 GW | GW 帳推算 | 電力公司大負載合約、併網佇列、自建電廠公告 | 併網是延遲的主要來源 | v0.3 |
| C5 | 發電機組 ＝ 自建發電 GW 與備援需求 | GW 帳推算 | 燃氣渦輪、往復式機組、燃料電池、柴油備援廠商訂單與產能 | 訂單積壓 ≠ 年交付量 | v0.3 |
| C6 | 電源與散熱 ＝ 新增 GW × 每 GW 內容（Siliconomics） | GW 帳推算 | 台達電、光寶科、奇鋐、雙鴻、Vertiv 等營收與指引 | 份額與單價 | v0.4 |
| C7 | AI 投資地圖個股：個股模型的 AI 子列 ≈ GW 帳切片 × 份額 | GW 帳切片 | 各個股模型 `segments.json` 的 AI 子列 | AI 只是官方分部下的子列 | v0.4 |

- 容許帶預設 ±15%（[Assumed]，In `tolband`）；雙方皆 [Verified] 才判「對得上／偏差」，否則標「證據不足」（附帶內／帶外）。超出容許帶時列出可能原因，不自動調參數。
- 推導方向：公司公告 GW 屬 [Interested-party]。GW 帳同時算兩條：(甲) 公告路徑；(乙) 實體可支撐路徑＝MIN（CoWoS→顆數→GW、HBM、併網、發電機組、變壓器）。主輸出是兩者差距與最先綁定的瓶頸，不以公告值反推參數。v0.1 只有 (甲)。

### 基建延遲情景（供給側、與需求情景 A／B／C 正交的第二軸）

| 延遲情景 | 定義（初值 [Assumed]，依證據校準） | v0.1 實作 |
|---|---|---|
| D0 準時 | 公告上線時程照做 | 延遲 0 |
| D1 基準延遲 | 每一管線階段（動工→完工→併網→上線）延遲取各瓶頸的歷史中位數 | 評價時點後的上線量延後 MAX(已知延遲, `dly_d1_m`)；佔位 3／6／9 月 |
| D2 嚴重延遲 | 併網與發電機組綁定：一定比例的新 GW 再延 12 個月 | D1 之上，`dly_d2_bind` 比例再延 `dly_d2_m`；佔位 0.2／0.35／0.5、12 月 |

延遲的傳導：晶片等電（出貨不變、上線延後 → 「已出貨未上線」GW，`GIF_ShippedNotOnline_*`）；TK 端（閒置期持有成本照付、產能為零，由公司模型讀 GW 帳上線路徑計算）；AI半導體 端（延遲持續兩季以上 → 後續拉貨減少）。輸出：需求 × 延遲 3×3；對外發布 A-D0、B-D1、C-D2 三條與基準 A-D1。

## 頁面

| 頁 | 內容 |
|---|---|
| README | 命題、驅動 → 推導圖、頁面、版本 |
| In | S 純量（單位換算、容許帶、lag_bldg、g_lt、延遲 D1／D2、玩家最新觀測、合約與租賃、NonNV 覆寫、財年月）；Y 年度（玩家年底 MW、個股資本支出、非 AI 資本支出、使用者總量、第三方總量、世代組合）；T1 營運者→使用者分攤；T2 園區管線 |
| TK_In、SIL_In | 鏡像（綠字） |
| Players | 玩家 × 年 TK GW（甲：公告路徑），年平均、年新增、新納入存量；補值規則見 Decisions GL01-03 |
| Matrix | 營運者 × 使用者（2025–2027 年底），含未分配欄 |
| Pipeline | 園區＋未分園區殘差 × D0／D1／D2 |
| Deploy | 需求 A／B／C × 延遲 D0／D1／D2 |
| C0、C1、Compare | 對帳與並列 |
| Interface | `GIF_MktGW_YE／Avg／Add_<A|B|C>_<D0|D1|D2>_<yyyy>`、`GIF_ShippedNotOnline_<A|B|C>_<D1|D2>_<yyyy>`、`GIF_PlayerGW_YE_<player>_<yyyy>`，2024–2032 |
| Checks | 錯誤格 0、Matrix 列和＝欄和＝總量、Pipeline D0＝Players（0.01 GW）、殘差恆等、延遲單調、In 標記與來源、`GGOV_Errors`＝0 |
| Decisions、DB_Evidence、Sources、Requests | 決策（GL01-01 起）、證據、來源 ID、對 TK／Siliconomics 的請求 |
