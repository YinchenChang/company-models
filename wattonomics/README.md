# Wattonomics（瓦帳）— 全市場 AI 資料中心 GW 唯一基準事實

**Wattonomics（瓦帳）＝全市場 GW 唯一基準事實，與 Tokenomics（token）、Siliconomics（矽）成系列；舊名 GW 帳／gwledger。**

> **下游取用 2027 年後數值前看覆蓋率**（`WIF_Coverage_<yyyy>`＝有證據支撐的新增 GW ÷ 主路徑新增 GW；< 0.5 時 Checks WARN）。2027 起主路徑七家（MSFT、GOOGL、AMZN、META、ORCL、CRWV、NBIS）由資本支出推得（甲2），純站點證據版另出 `_SiteOnly`。

算力收支 R4「產業部署路徑」的實作：市場總 GW、各玩家 GW（營運者／使用者）、部署管線與基建延遲情景的**唯一事實** Excel。每 GW 係數只讀 Tokenomics（TK）快照（快照未含的 `IF_FacilityGW` 直接讀 TK 活頁簿 Interface 第 8 列，`builder/seed/tk_extra_v21.7.json`）、加速器顆數與情景係數只讀 Siliconomics 快照，不重估兩者權威的量。

| 項目 | 內容 |
|---|---|
| 現行版本 | 見 `model/CURRENT` |
| 下游快照 | `data/CURRENT_SNAPSHOT` 指向的 `data/wattonomics_snapshot_v*.json`（`tools/export_snapshot.py` 產生；名稱前綴 `WIF_`） |
| TK 快照 | `siliconomics/data/tokenomics_snapshot_v21.7.json`（與 Siliconomics 同一份，不另抽） |
| Siliconomics 快照 | `siliconomics/data/CURRENT_SNAPSHOT` 指向者 |
| 規劃書 | `docs/plan/20261028_全市場GW帳_規劃.md` |
| 工作單 | `wattonomics/docs/workorders/` |
| 證據（查證代理產出，chat 端複核後才入 In） | `wattonomics/evidence_raw/`；`tools/import_evidence.py` → `builder/seed/evidence_v0.1.json`（DB_Evidence，附 C 段判定）；`tools/import_funding.py` → `builder/seed/funding_v0.1.1.json`（FundRaw、Facility，附 G 段判定與更正） |
| 互動頁取數 | `data/explorer_v*.json`（`tools/export_explorer.py`；扁平紀錄，含標記與來源 ID） |
| AI半導體 次產業比例 | `markets/invest_map/data/segments_ai_semis_v0.2.3.csv`（SegIn 頁鏡像；Spend 拆分） |

## 建置（每個 Excel 新版本）

```bash
# 0) 證據有更新時：python3 wattonomics/tools/import_evidence.py; python3 wattonomics/tools/import_funding.py
# 1) 產生（之後各版加 --base wattonomics/model/<現行>.xlsx，In 頁藍字讀回現行 Excel）
python3 wattonomics/builder/build.py --out build/raw/<YYYYMMDD>_Wattonomics_vX.Y.xlsx
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

## 原則

- Excel 是唯一事實來源；`builder/` 只產生公式頁與格式，計算一律為 Excel 公式。`builder/seed/` 只用於 v0.1 首建。
- 輸入（In 頁藍字）由 Excel 擁有；每列有標記（Verified／Interested-party／Analogy／Assumed／Derived／待證據）、來源 ID（Sources 頁）、口徑（IT／設施／未說明）。Assumed 與 Analogy 純量一律給低／基準／高。
- 單位（規劃書 r1，Requests R1 裁定）：主口徑 **IT GW＝TK GW（IT 關鍵電力）**；設施電力＝IT × TK `IF_FacilityGW`（基準 1.2，1.1–1.3）。公司標設施者 ÷ `IF_FacilityGW`；口徑未說明者基準取 IT 讀法，Players 另列設施讀法；Epoch 晶片功耗 × `k_chip2it`（[Analogy] 1.65／1.85／2.08）。
- 證據：只用 `evidence_raw/C_review.md` 判「屬實」或已更正的列；「錯誤」「無法核對」「未抽查」的站點列在 In T2 但納入＝0。
- TK 或 Siliconomics 權威的量需要改值 → 停，寫進 Requests 頁。
- 下游只讀 Interface 的 `WIF_` 名稱（經快照）。
- 本 repo 為公開 repo：不放 Andy 的個人情境或先驗。

## 固定規格（Andy 2026-10-28；規劃書第 3、4 節）

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

### 基建延遲情景（供給側、與需求情景 A／B／C 正交的第二軸）

| 延遲情景 | 定義（初值 [Assumed]，依證據校準） | v0.1 實作 |
|---|---|---|
| D0 準時 | 公告上線時程照做 | 延遲 0 |
| D1 基準延遲 | 每一管線階段（動工→完工→併網→上線）延遲取各瓶頸的歷史中位數 | 評價時點後，每年排定新 GW 的 `dly_d1_s` 35%（25–45%）延 `dly_d1_m` 6 個月（3–12）；園區有已知延遲者全額延已知月數。[Derived]：B_delay D#15（Sightline 26% 為容量、10% 為專案數，不可相加）、D#16／D#19 |
| D2 嚴重延遲 | 併網與發電機組綁定：一定比例的新 GW 再延 12 個月 | D1 之上，原排定 2027–28 且未確保電力（In T2 O 欄≠1）的新 GW 中 `dly_d2_bind` 40%（30–50%）再延 12 個月。[Derived]：D#2、D#3、D#5、D#7、D#8、D#11；反證 D#25 |

延遲的傳導：晶片等電（出貨不變、上線延後 → 「已出貨未上線」GW，`WIF_ShippedNotOnline_*`）；TK 端（閒置期持有成本照付、產能為零，由公司模型讀瓦帳上線路徑計算）；AI半導體 端（延遲持續兩季以上 → 後續拉貨減少）。輸出：需求 × 延遲 3×3；對外發布 A-D0、B-D1、C-D2 三條與基準 A-D1。

## 頁面

| 頁 | 內容 |
|---|---|
| README | 命題、驅動 → 推導圖、頁面、版本 |
| In | S 純量（單位換算、容許帶、lag_bldg、g_lt、延遲 D1／D2、玩家最新觀測、合約與租賃、NonNV 覆寫、財年月）；Y 年度（玩家年底 MW、個股資本支出、非 AI 資本支出、使用者總量、第三方總量、世代組合）；T1 營運者→使用者分攤；T2 園區管線 |
| TK_In、SIL_In | 鏡像（綠字） |
| Players | 玩家 × 年 IT GW（甲：公告路徑；輸入→晶片換算→觀測→內插→園區延伸），低／高、年平均、年新增、新納入存量、設施 GW；規則見 Decisions GL01-03 |
| Matrix | 營運者 × 使用者（2025–2027 年底；In T1 逐筆分配），無揭露者在未分配欄 |
| Pipeline | 園區＋未分園區殘差 × D0／D1／D2 |
| CapexGW | 甲2：七家資本支出（個股模型 annualEstimates／指引）× AI 占比（`ai_sh` 0.75／0.85／0.95，Amazon 0.68）÷ TK `IF_CapexTotal`（世代加權，NonNV × 持有比；建物提前 `lag_bldg`）＝年新增 IT GW；敏感度列 |
| Deploy | 需求 A／B／C × 延遲 D0／D1／D2；2030–32 年底 GW 年增率由 2029 線性收斂到 `g_lt`（0.10，0.05–0.15）|
| C0、C1、Compare | 對帳與並列；C1 七家 2027 起標「同源（不對帳）」，C1b 廣義口徑（資本支出含融資租賃＋尚未起租變化；Microsoft 改列營業租賃問題）；Compare 含 xAI 宣稱（不入主路徑）、AI 占比、lag 敏感度 |
| Funding、FundRaw | 玩家 × 期間（2024、2025、2026 最新）營業現金流、資本支出（現金／含融資租賃）、自由現金流、債（公司債、定期貸款、GPU 擔保債、私募信貸）、淨負債變化、新增融資租賃、尚未起租、表外合資、股權、供應商融資、買回＋股利；資金來源比例兩種口徑：【支付順序】內部→租賃→股權→債→其餘（外部資金排在營業現金流之後，營業現金流足以支付時外部比例為 0）；【總額】期間外部資金全額（發債＋定期貸款＋GPU 債＋私募信貸＋新增融資租賃＋表外合資＋股權）÷ 資本支出，與營業現金流 ÷ 資本支出並列（可合計 > 100%）；互動頁預設總額口徑；循環融資 17 條。只用 G 段判屬實或已更正列 |
| Facility | 每 IT GW 廠房 12.67（TK）拆分：配電、備援發電機（1.6 × 1.0）、UPS、冷卻、建築、人工設計、其他專業服務、網路消防、土地、殘差；自建發電另列（1.44 GW × 1.6 $M/MW，× 自建比例 0.10／0.20／0.35）；主變 2.0 GVA 為判斷值；市場層＝主路徑年新增 × 每 GW |
| SegIn、Spend | AI半導體 次產業「GB300 對照 ÷ IT」比例鏡像；年 × 項目市場金額（IT 8 項＋廠房＋自建發電） |
| Interface | `WIF_Coverage_<yyyy>`；`WIF_Spend_<項目>_<yyyy>`、`WIF_Facility_<項目>_<yyyy>`、`WIF_Funding_<玩家>_<期間>_<項目>`；`WIF_MktGW_YE／Avg／Add_<A|B|C>_<D0|D1|D2>_<yyyy>`、`WIF_ShippedNotOnline_<A|B|C>_<D1|D2>_<yyyy>`、`WIF_PlayerGW_YE_<player>_<yyyy>`（主路徑），同名加 `_SiteOnly` 後綴＝只有 甲；2024–2032 |
| Checks | 錯誤格 0、Matrix 列和＝欄和＝總量、Pipeline D0＝Players（0.01 GW）、殘差恆等、延遲單調、In 標記與來源、`WGOV_Errors`＝0 |
| Decisions、DB_Evidence、Sources、Requests | 決策（GL01-01 起）、證據、來源 ID、對 TK／Siliconomics 的請求 |
