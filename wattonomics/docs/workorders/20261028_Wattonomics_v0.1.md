# 工作單｜瓦帳（Wattonomics；舊名 GW 帳）v0.1（R4）｜2026-10-28 r1（第 0 節口徑更正；v0.1.1 改名）

規劃：`docs/plan/20261028_全市場GW帳_規劃.md`（PR #201 已合併）。本版只做 C0、C1、延遲管線與 WIF_ 輸出；C2–C7 留待 v0.2 起。

## 0. 原則（比照 `siliconomics/README.md`）
- Excel 是唯一事實來源；`wattonomics/builder/build.py` 只產生公式頁與格式，計算一律為 Excel 公式。輸入（In 頁藍字）由 Excel 擁有；之後各版以 `--base` 讀回。
- TK 的量只在 `TK_In` 頁鏡像（讀 `siliconomics/data/tokenomics_snapshot_v21.7.json`，同一份快照，不另抽）；Siliconomics 的量只在 `SIL_In` 頁鏡像（讀 `siliconomics/data/` CURRENT 快照）。不得重估 TK 或 Siliconomics 權威的量。
- 每個輸入列有標記（[Verified]／[Interested-party]／[Analogy]／[Assumed]／[Derived]）、來源 ID、口徑（IT／設施／未說明）；Analogy 與 Assumed 一律給低／基準／高。
- 單位（r1 更正，2026-10-28；Requests R1 裁定）：主口徑 **TK GW＝IT 關鍵電力**（TK Interface 第 1 列「每 GW＝IT 關鍵電力」）；**設施電力＝IT × `IF_FacilityGW`**（TK Interface 第 8 列，基準 1.2，1.1–1.3；v21.7 快照未含，直接讀 Tokenomics 活頁簿）。`g_ITtoFac`＝1（IT→IT）。公司標設施者 ÷ `IF_FacilityGW`；口徑未說明的公司值以「IT、設施」兩讀並列，基準取 IT 讀法 [Assumed]（CoreWeave 模型 mwBasis 先例）。r0 原文「TK GW＝設施電力、VR200 每 TK GW 含 919.8 MW IT」作廢（919.8 MW 是 VR200 每 IT GW 中機架本身的 IT 功率）。
- 本 repo 為公開 repo：不放 Andy 個人情境或先驗。

## 1. 頁面
| 頁 | 內容 |
|---|---|
| README | 命題（一句）、驅動 → 推導圖、頁面說明、版本 |
| In | 輸入：玩家年底上線 MW（2024–2026 實際或最新，2027–2032 公告路徑）、口徑、標記、來源；園區管線（園區、營運者、使用者、MW、公告上線年月、已知延遲）；延遲參數 D0–D2；世代組合；容許帶 0.15 |
| Players | 玩家 × 年（2024–2032）年底上線 TK GW（甲：公告路徑），年平均＝(期初＋期末)/2；分「營運者」口徑 |
| Matrix | 營運者 × 使用者（2025 年底、2026 年底、2027 年底）：列＝營運者、欄＝使用者（自用、OpenAI、Anthropic、xAI、Meta、其他）；Check：列和＝Players 營運者合計；欄和＝使用者合計 |
| Pipeline | 園區管線 × 延遲情景：上線年月＝公告年月＋延遲（D0 0；D1 各瓶頸中位數；D2 併網與發電機組綁定比例 × 12 個月）；按年彙總成 D0／D1／D2 的上線 GW |
| Deploy | 市場總 GW（年底、年平均、年新增）× 需求情景 A／B／C × 延遲 D0／D1／D2。需求情景係數讀 `SIF_UnitsCoef_*`（2028–29）；2030–2032 先用「2029 年增率線性收斂到長期成長率」共用規則（長期成長率 In 頁 [Assumed]）|
| C0 | 由下而上總 GW vs 第三方總量（每來源一列：口徑換算後、比值、容許帶、狀態）；第三方只作佐證 |
| C1 | 每玩家：瓦帳推算資本支出＝Σ 世代 新增 GW × TK `IF_CapexTotal`（NonNV 部分 × `IF_NonNV_Hold_*`），建物提前 `lag_bldg` 月入帳（[Assumed] 12–24）＋非 AI 資本支出（In）vs 個股模型資本支出（讀各公司 `company.json` annualEstimates／companyGuidance；MAG 若 PR #190 未合併則讀現行檔並標版本）。輸出比值、容許帶、狀態、可能原因欄 |
| Compare | 瓦帳部署 GW vs Siliconomics `SIF_DeployGW_A/B/C`（成本等值 GW）並列，差額與原因說明 |
| TK_In、SIL_In | 鏡像（綠字） |
| Interface | `WIF_` 名稱：`WIF_MktGW_YE_<scen>_<delay>_<yyyy>`、`WIF_MktGW_Add_*`、`WIF_ShippedNotOnline_*`（D1、D2 相對 D0 的未上線 GW）、`WIF_PlayerGW_YE_<player>_<yyyy>`；2024–2032 |
| Checks | 錯誤格 0；Matrix 列和＝欄和＝總量；Pipeline D0 彙總＝Players 公告路徑（容差 0.01 GW）；每個 In 列有標記與來源；`WGOV_Errors`＝0 |
| Decisions、DB_Evidence、Sources | 決策紀錄（GL01-01 起）、證據表（讀 `wattonomics/evidence_raw/*.csv` 整理後的版本）、來源 |

## 2. 輸入來源
- 證據：`wattonomics/evidence_raw/A_us_cloud.*`、`B_labs_china_totals.*`、`B_delay.csv`（查證代理產出，chat 端複核後才可入 In）。
- 公司模型：microsoft、alphabet、amazon、oracle、coreweave、nebius、whitefiber、openai、anthropic 的 company.json／data。
- 世代組合：2024 Hopper；2025 GB200／GB300；2026 GB300／VR200（vr26 讀 Siliconomics）；2027 起 VR200，Rubin Ultra 比例讀 Siliconomics Units 比例；NonNV 比例＝Siliconomics Units 中非 NVIDIA 平台 IT capex 占比（[Derived]）。

## 3. 建置與檢查
```bash
python3 wattonomics/builder/build.py --out build/raw/<YYYYMMDD>_Wattonomics_v0.1.xlsx
soffice --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir wattonomics/model build/raw/<檔名>.xlsx
python3 wattonomics/tools/export_snapshot.py wattonomics/model/<檔名>.xlsx wattonomics/data/wattonomics_snapshot_v0.1.json
python3 wattonomics/tools/check.py wattonomics/model/<檔名>.xlsx   # 錯誤格 0、Checks 全過、標記齊全
```
`wattonomics/README.md` 寫入規劃書第 3、4 節的對帳鏈與延遲情景規格（Andy 10-28 固定規格）。

## 4. 停止條件
- 任何 TK 或 Siliconomics 權威的量需要改值 → 停，寫進 Requests 頁。
- 證據互相矛盾超過 2 倍且無法判定口徑 → 列兩讀並標 [Assumed]，不自行挑一個。
- Checks 不過 → 不交。

## 5. 交付
分支 `claude/gwledger-v0.1`（v0.1 已合併；v0.1.1 起為 `claude/wattonomics-v0.1.1`）；PR 說明列：玩家表摘要、2025／2026／2027 年底市場 GW（D0／D1／D2）、C0 與 C1 對帳狀態表、與 Siliconomics 成本等值 GW 的差。
