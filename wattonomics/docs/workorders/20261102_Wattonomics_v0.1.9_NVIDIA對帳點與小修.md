# 工作單：Wattonomics v0.1.9＝NVIDIA 資料中心營收對帳點＋核對意見 D1–D3

- 日期：2026-11-02（台北）；chat 端：session_01V11BcwWqQyh9suCNhAYQKQ
- 現行：main `07afb21`，`model/20261102_Wattonomics_v0.1.8.xlsx`；Checks 182＝OK 172／WARN 8／ERROR 0
- 分支：`claude/wattonomics-v0.1.9`
- 依據：v0.1.8 獨立核對 `docs/verify/20261102_v0.1.8_independent_verify.md` D1–D3；AI半導體 11-02 證據送件「NVIDIA FY2028 營收指引」（交流 `已處理/20261102_AI半導體-to-算力收支_回執_R4歸因與2027證據送件.md` 第 3 節）；算力收支回執 `20261102_算力收支-to-AI半導體_回執_NVIDIA證據送件、K4乙排入v0.2、r4_bridge列名變更.md` 第 2 節。
- 原則：**不改基準任何數值**（三層只給基準；對帳點只並列，不校準）。v0.1.8 → v0.1.9 的 `WIF_` 既有名稱值變動須為 0（diff_snapshot 驗證）。

## 1. 工程決定（chat 端已定）

| # | 決定 |
|---|---|
| F1 | D1：GL01-19 決議文刪去「最早年之前取最早值」，改寫為 v0.1.8 實際規則（缺 t−2 用 t−1、兩者皆缺用 t；不往前回填）。 |
| F2 | D2：`tools/export_explorer.py` 資本支出說明改「In Y9 優先、缺年 In Y2」；`capex_gw` 紀錄加欄位 `basis: "付款年"`（頁面不顯示，不改值）。 |
| F3 | D3：Pipeline 兩個延遲情景輔助表標題（v0.1.8 的 A207、A219 附近）加「〔停用，v0.2 刪〕」。 |
| F4 | **Compare 新節「NVIDIA 資料中心營收對帳點」（2025–2027）**，逐年列：(1) Wattonomics 上線年新增 IT GW（基準）；(2) 換成出貨年 GW＝上線年反推（用 `lag_on`、`inv_on`，與甲2 平移同一套參數的反向；公式寫明）；(3) 非 NVIDIA 占比（Players 市場列，Derived）；(4) NVIDIA 出貨 GW＝(2)×(1−(3))；(5) 每 IT GW NVIDIA 內容（$B/GW，世代加權）；(6) 隱含 NVIDIA 資料中心營收＝(4)×(5)；(7) NVIDIA 實績／指引：CY2025 193.7（Verified）、CY2026 375.5（Q1–Q2 Verified、Q3 指引 Interested-party、Q4 Assumed）、CY2027 約 620–645（FY2028 指引年增約 70% × FY2027 × 資料中心占比 92.5%；Derived／Interested-party；低、基準、高三值）；(8) 比值＝(6)÷(7) 與容許帶 `tolband` 判定；(9) 反解：要對上 (7) 所需的每 GW NVIDIA 內容。另列同口徑的 Siliconomics r1 欄（讀 `siliconomics/docs/reports/r4_bridge/` r1 活頁簿的淨新增上線，只作並列）。 |
| F5 | (5) 的來源：先追 Siliconomics R4 橋接 In 頁 `nvc_gb_tk`（20.8）的推導方式（預期為 TK `IF_RacksPerGW` × `IF_RackPrice` × 某 NVIDIA 占比），照同法逐世代（GB200、GB300、VR200）由 TK 快照算出，用 Wattonomics 各年 NVIDIA 世代組合加權；若 TK 快照無 VR200 架價或無法拆出 NVIDIA 占比，(5) 用 GB300 值並標「VR200 未拆分」，同時在 Requests 頁新增一列請 Tokenomics 提供「每 IT GW NVIDIA 內容（分世代）」。 |
| F6 | Checks 新增 1 項 INFO 類（不影響 ERROR／WARN 計數）：對帳點 2026 比值是否在 ±15% 內。新增 `WIF_Recon_NV_<項>_<yyyy>`（ShipGW、NVGW、ContentPerGW、ImpliedRev、ActualRev、Ratio、ReqContent；2025–2027）。 |
| F7 | 證據登錄：AI半導體 11-02 送件以 Sources 新 ID 登錄（NVIDIA FY27 Q2 法說 2026-08-26，CFO 指引 FY2028 年增約 70%），標 Interested-party；DB_Evidence 判定「只作對帳點，不入基準」。 |
| F8 | Decisions 新增 GL01-55 起：F4–F7 的理由；引 Andy 11-02 16:22 原話（差距大時各用自己的數字、上游先更新再對帳）。README 版本紀錄與頁面節同步；互動頁不加此節（v0.2 再議）。 |

## 2. 驗收
- 錯誤格 0；Checks ERROR 0、WARN 數同 v0.1.8（8）；check.py 通過。
- `tools/diff_snapshot.py` v0.1.8 → v0.1.9：既有名稱值變動 0；只有新增 `WIF_Recon_NV_*`。
- 對帳點 2027 值與 chat 端回執估算（Wattonomics 出貨年約 27.3、NVIDIA 約 20.8 GW、所需每 GW 約 300–310 億）差異需說明。
- 報告 `docs/verify/20261102_v0.1.9_build.md`（STATUS 格式）。未參與代理獨立核對由 chat 端另派。

## 3. 不做
- 不改基準；不發布 gh-pages；不刪停用欄位（v0.2）；不做 K4（v0.2）。
