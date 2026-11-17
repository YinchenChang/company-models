# 工作單：新增五家個股模型與地圖內容（2026-11-17）

## 1. 範圍
AI 算力地圖上狀態為「待建」或「另一串進行中」、目前沒有內容的五家：

| 地圖代號 | 公司 | 資料夾 | 地圖分層 | 母體重要性 | 樣板（同次產業既有模型） | 幣別 |
|---|---|---|---|---|---|---|
| INTC | Intel | `intc/` | 代工與封測 | 3 | `tsmc/`、`umc/`（代工）；`amd/`（產品線） | USD |
| GFS | GlobalFoundries | `gfs/` | 代工與封測 | 3 | `umc/`、`vis/` | USD |
| 8035 | Tokyo Electron（TEL） | `tel/` | 設備與材料 | 1 | `amat/`、`lrcx/`（WFE）；`accretech/`、`advantest/`（日股格式） | JPY |
| VRT | Vertiv | `vrt/` | 電源與散熱 | 2 | `delta/`（電源散熱） | USD |
| ABBN | ABB | `abb/` | 電源與散熱 | 3 | `delta/`；日、歐股處理參 `accretech/` | 報表 USD、股價 CHF |

先前「Vertiv、ABB 不納入（等電力與機房範圍定案）」「Intel、GlobalFoundries 暫不做」的預設，已由 2026-11-17 的新指示取代。

## 2. 規格（權威文件，照做）
- `docs/plan/個股DCF標準規格.md`：標準版 FCFF DCF、加權目標價與評等邊界（>+50% 強力買進、>+25% 買進、≥−25% 中立、≥−50% 放空，其餘為強力放空）。
- `docs/plan/個股HTML版面規格.md`：版面 v1.3，章節順序為命題 → 估值 → 敏感度 → 反向DCF → 損益 → 驅動 → 個別模組 → 對帳 → 資料狀態。
- `docs/plan/個股模型原則_命題加營收分部.md`：原則 1–51。
- `docs/plan/三層基準協調規則.md`：Coord 頁與 coord_alt。
- `docs/plan/個股模型驗收_不需驗收項目.md`。
- 預測期到 FY2031–32。Excel（事實基礎）與 HTML（展示層）兩者都要。至少三個重點觀察；上檔超過 +25% 或低於 −25% 時，須有一點說明主因。
- **資料標籤：** 每個輸入標 [Verified]、[Interested-party]、[Analogy]、[Assumed] 或 [Derived]。Analogy 與 Assumed 給區間並放進 Sens。
- **行業變數：** 只讀 Siliconomics 基準（`siliconomics/dist` 現行版），情境只留在個股。比照樣板讀哪些變數：代工讀晶圓／先進製程需求，設備讀 WFE，電源散熱比照 `delta/` 的讀法。Siliconomics 沒有的變數，在個股以 [Assumed] 或 [Analogy] 建立，並在 README 列為「建議補進 Siliconomics」，不得改 Siliconomics。

## 3. 資料
- **價格：** 11-13 收盤。來源只能是 `markets/invest_map/data/closes_20261020_20261113.csv` 與 `markets/invest_map/data/closes_extra_20261113.csv`（含 ABBN.SW 69.96 CHF；CHF=X 0.83593＝每美元 CHF；EUR=X 0.90496；JPY=X 159.875）。
- **財報：** 以公司原文為準（10-Q／10-K、IR 新聞稿、決算短信、季報），可用 WebSearch／WebFetch 取得。SEC 原站常回 403，可改用公司 IR 頁或 stockanalysis 等轉錄，並在 README 註明來源等級。
- **共識與同業倍數：** `docs/workorders/20261117_新增五家_Yahoo基本面快照.md` 為第三方彙整，只作交叉檢查與 Comparable 同業 EPS，標 [Interested-party]／第三方。同業 EPS 可用 stockanalysis 的前瞻 EPS（repo 既有做法）。
- 最近一季必須滾入：INTC、GFS、VRT、ABB 為 2026-09 季；TEL 為 FY3/27 上半年（2026-09）與公司通期預想。

## 4. 各家重點
- **INTC：**
  - 分部：Intel Products（CCG、DCAI）與 Intel Foundry（含外部代工 18A／14A、先進封裝 EMIB／Foveros）。
  - 估值主軸：代工轉虧為盈的路徑與資本支出強度。
  - 少數股權明顯上升（Arizona Fab 與 Brookfield、Apollo 的 Ireland 合資；Mobileye、Altera 持股）；美國政府持股與 SoftBank、NVIDIA 入股使股數增加。以上皆以原文為準，計入 EV 到股權的橋接。
- **GFS：**
  - 成熟與特殊製程（FDX、矽光子、RF）。2026-09 季淨利為負，須讀原因（一次性項目？）。
  - Mubadala 持股、庫藏股。
- **TEL：**
  - 分部：塗布顯影（市占約 91%，公司自估，[Interested-party]）、蝕刻、沉積、清洗、測試。
  - 讀 WFE 與 HBM／先進邏輯需求。日股格式（3 月年度、日圓）比照 `accretech/`、`advantest/`。
- **VRT：**
  - 機房電力（UPS、配電）、熱管理（含液冷 CDU）、服務。訂單與 backlog、book-to-bill。
  - AI 資料中心連動：讀 Siliconomics 中對應的機房／功率變數。若無，比照 `delta/` 以 [Assumed] 建立，並列入 README 建議。
- **ABB：**
  - 分部：Electrification、Motion、Process Automation（Robotics 已出售給 SoftBank，依原文處理為停業單位或已完成出售）。
  - AI 相關只占 Electrification 的資料中心部分，須拆出比重，並在命題中寫明「AI 暴露度有限」的含義。
  - 股價為 CHF，報表為 USD：每股價值換算用 CHF=X。

## 5. 地圖內容（每家同一分支一併做）
- `markets/invest_map/profiles/<代號>.md`（INTC、GFS、VRT、ABBN、8035）：依 `markets/invest_map/profiles/00_格式規格.md` 撰寫。出處標記引用本模型現行 dist 檔名與儲存格，或 README 段落。不得有無出處的數字；外部事實標 [外部・未經本專案檢核]。
- 地圖抽取程式（`extract_models.py`、`build_xlsx.py`）、prices.csv、model_status.csv 由主串整合時處理，建置代理不要改。

## 6. 交付
- 分支 `claude/new5-<資料夾>`，由 `claude/new5-base` 開出。
- 每完成一個段落就 commit 並 push：資料與建置腳本 → dist → HTML → profile。
- 版號 v0.1.0。dist 只留一個 xlsx 與一個 HTML。
- 檢查：
  - `tools/company_html/check_company_html.py`、`check_layout_v13`。
  - `tools/sens_coverage/check_sens_coverage.py <xlsx>`。
  - `tools/model_governance/check_governance.py`、`check_display.py`（若適用）。
  - `tools/baseline_coordination/check.py`。
  - 公司自帶檢查。
- 由未參與建置的代理獨立核對數字後，才併入 main。
