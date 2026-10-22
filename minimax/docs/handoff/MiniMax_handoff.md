# MiniMax 收支模型 交接檔（2026-10-22；v0.2）

## 1. 版本與位置
- 現行：`dist/20261022_MiniMax收支模型_v0_2.xlsx／.html`（PR #75，分支 `claude/minimax-v0.2`，待獨立查核、待 Andy 指示合併）。v0.1 在 `dist/_待刪除/`（確認後由 Andy 或 chat 端刪除）。
- Tokenomics：官方快照 `data/tokenomics_snapshot_v19.0.json`（`data/tokenomics_snapshot_CURRENT` 指向；提交 `1916a06`；SHA-256 `d1728705…970b9d8`）。重取需在 `tokenomics_names.txt` 加名稱後跑 `scripts/tk_snapshot.py`，再 `scripts/tk_attribution.py`。
- 報告：`docs/reports/20261022_minimax_v0.2.md`（結論、差異拆解、預設、缺口）＋ `_對照.xlsx`；`20261022_minimax_M1.md`、`_M2.md`；命題與驅動對照表 `docs/plan/20261022_MiniMax_命題與驅動對照表.md`。
- 建置：`scripts/build_all.sh <日期> <版本> [dist]`（含敏感度、HTML、`check_html.py`、`ci_check.py`）；約 1.5 分鐘。CI：`.github/workflows/minimax-build.yml`。

## 2. 決定（已定案）
- D1–D8 預設全部沿用（見規格第 3 節與報告第 5 節）；D2 維持 12、D8 只更新 2026。
- 回本比值的 MW 口徑＝模型自身年均上線 MW；h＝Hopper 與 GB200 持有成本依 I05 加權；P＝MiniMax 實收單價。
- 資料分層：市場與公司資料入 `SRC_MM`（S078–S112），假設入 `Inputs`（新增 I30、I31、I33），Tokenomics 只經快照。

## 3. 目前預設（會改變輸出的）
I02＝12（Analogy 8–20）；I31＝0.20（Assumed 0–1，M3.1 新價流量占 2H26）；I30＝0.50（觀測）；I33＝8（SMM 租價卡數口徑，Assumed 1–8）；市值＝10-21 推算收盤價 HK$171.9；1H26 平均匯率 6.90（Assumed）。

## 4. 已知限制與待辦
1. **回本比值水準∝I02**（報告第 2.1 節）：任何與其他公司（OpenAI、智譜）並列的比較，要同時標示 I02 假設。
2. 1H26 現金流量表、自建集群規模、2026 消費端付費人數 × ARPPU、M3 實收 0.22 的一手出處：找不到（可修）；需在可連 HKEX 的環境補。
3. Tokenomics 缺口：中國晶片世代／租價、GPU／MW 官方值、服務端落差分解 `IF_` 名稱、中國 API 有效牌價；快照擴充 `IF_ProdDerate`、`IF_UtilActual` 後補 D1 的 Prod 層敏感度。
4. 營運燒錢仍以經調整淨利代營運現金流（FY2025 差 −22.7）；有 1H26 現金流量表後可校準。
5. D5：A 股（科創板）情境、多情境機率加權未做；現價低於 IPO 價，融資條件假設（I25 折價 10%、可轉債現金贖回）需重審。
6. 智譜同業對照的市值為 10-02（未更新）。
7. 獨立查核（未參與建置的代理核 Excel 與 HTML 數字）待排。
