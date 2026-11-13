# CoreWeave（CRWV）個股模型 v5.0.1（Oracle 收支型樣板）

**v5.0.1（2026-11-13）：投資建議加上收支關卡（原則第 10 節第 63 條）。** Andy 2026-11-13「个股模型还是要遵守这个收之不过就不能买进的这个规则……个股模型这边还是要用相同的原则来评断它的評等」。Val 頁新增「估值評等（只看上檔空間）」「入不敷出判定」兩組列；「投資建議」＝收支關卡不通過（情境 B 的 FY2026–FY2030 自由現金流合計 < 0）且估值評等為買進以上時改列「收支不足」，其餘同估值評等。本版結果：估值評等買進（邊界）、收支合計 −95.7 → **投資建議：收支不足**。加權目標價與其他數值逐格不變（76.20）。另更正 xlsx README「已知限制」過時的同業家數、In 頁隨需占比說明（v4.13 基準 0%，v5.0 改 10% 未列入採用規劃書，待 Andy 決定）。

**v5.0（2026-11-12）：** Andy 2026-11-12「Coreweave 已经公布财报了，所以要更新 Coreweave 的财报。再将 html 的版面套用 Oracle 最新的版本。」＋確認 v4.13 驗收採用規劃書（J1–J5 照預設）。工作單 `docs/workorders/20261112_CoreWeave_v5.0_Oracle樣板重建_Q3.md`；報告 `docs/reports/20261112_coreweave_v5.0_Oracle樣板重建與Q3.md`。舊 JS 雙引擎（v4.13，HTML＋Excel、`company.json`）整套移到 `legacy_v4/`（保留不刪）；v4.13 成品移 `dist/_待刪除/`。

數字以 `dist/` 的 xlsx 為準；下表由建置成品產生（`python3 tools/model_governance/version_table.py coreweave --update`），不手寫。

<!-- version-table: dist/_待刪除/20261112_CoreWeave個股模型_v5.0.xlsx｜dist/20261113_CoreWeave個股模型_v5.0.1.xlsx -->
| 項目 | v5.0 | v5.0.1 |
|---|---|---|
| 加權目標價（上檔空間） | 76.20（+25.4%） | 76.20（+25.4%） |
| DCF 目標價（情境加權） | 0.00 | 0.00 |
| Comparable 目標價 | 152.41 | 152.41 |
| 投資建議 | 買進 | 收支不足 |
| 現價 | 60.77 | 60.77 |
| 收支關卡 A／B／C（FY2026–FY2030） | -175.2／-95.7／-56.9 | -175.2／-95.7／-56.9 |
<!-- /version-table -->

**評等標示：** 投資建議為「收支不足」；估值評等的加權上檔空間緊貼「買進」下緣，xlsx README 與 HTML 橫幅以「收支不足〔估值評等：買進（邊界，±距離）；收支合計〕」標示，並由公式列出三個脆弱點：同業中位數剛好是 Oracle、Nebius 淨負債用估值日後一天公布的 9-30 申報、多個 Sens 單一案例會降回中立（Val 頁「橫幅：評等脆弱點」列）。

**命題：** 已簽約的 backlog 能否以高於資金成本的報酬轉成營收與現金，同時不讓資金缺口失控。**核心問題：** 每增加 1 GW，是否創造高於資金成本的價值；缺口靠多少新債與新股補、股東被稀釋多少。

## 建置

```
bash coreweave/make_v5.sh                        # 沿用已提交的敏感度、反解、拆解快照
SENS=1 REV=1 DEC=1 bash coreweave/make_v5.sh     # 改了公式或輸入時（約 2 分鐘）
```
步驟：建置（`build_coreweave.py`，openpyxl 活公式）→ LibreOffice 重算 → 公式錯誤 0 → 檢查列與收支恆等式（`scripts/check_v50.py`，含 Sens／RevDCF 快照）→ In 版面 → 敏感度覆蓋（`--require-keycol`）→ 分部 → HTML（`scripts/make_html_json.py` → 共用產生器）→ HTML 檢核 → 治理（build_id、模型狀態、單一標記、來源誘因、展示層一致性 `display_check.json`）。

## 檔案

| 檔案 | 內容 |
|---|---|
| `build_coreweave.py` | Python → Excel 產生器（17 頁：README → In → Q → Thesis → Seg → PL → CF → Fund → RPO → Sites → Debt → Lease → Val → RevDCF → Recon → Sens → Sources） |
| `scripts/runner_v50.py` | 整本重建＋LibreOffice 重算的共用工具 |
| `scripts/sens_v50.py`、`reverse_v50.py`、`decomp_v50.py` | 敏感度、反向 DCF（反解 k_adj）、目標價變動拆解 → `data/*_v50.json` |
| `data/coreweave_facts_20261112.json` | 事實總帳（Q3 新聞稿逐筆、Q2 10-Q 沿用項、來源 URL） |
| `data/tokenomics_snapshot_v22.0.json` | Tokenomics 快照（v18.0 保留供拆解） |
| `segments.json`、`display_check.json`、`html.json` | 分部宣告、展示層一致性設定、HTML 設定（由腳本產生） |
| `legacy_v4/` | v4.13 舊引擎（JS＋company.json＋驗證腳本），保留不刪 |

## 方法要點

- **命題模組（Thesis）：** 在役 MW（情境 A 8 GW／B 5.6 GW／C 已簽約 4.6 GW，2030 年底）× 每 MW 年收入（Tokenomics `IF_HoldEcon` 在役世代加權 × 單一定價倍數 k；k＝IREN–Microsoft 長約〔Analogy〕與 H100 現貨加權）；營運費用＝Tokenomics `IF_OpexGW`；租金＝每 MW 年租金；資本支出＝新增 MW × `IF_CapexIT`（λ 提前付款）＋自建機房＋汰換。不以營收、ARR、RPO、指引反推 k（CLAUDE.md 第 14 條）。
- **分部：** 只有一個應報導部門（AI 雲端平台）。
- **收支模組（CF、Fund、Debt、Lease、RPO、Sites）：** 現金 → 投資級債（借款 ÷ EBITDA ≤ 4.5）→ 新股（≤ 市值 20%）→ 高息債；原則 9 的淨負債 ÷ EBITDA、利息保障倍數只 WARN；原則 55 期間流動性、原則 62 新股占市值與累計稀釋。
- **評價（Val）：** DCF（FCFF、WACC、RONIC＝WACC、評價日股數）50%＋Comparable（新雲同業前瞻 EV/EBITDA 中位數 × FY2028 調整後 EBITDA，錨定 FY2027 年底）50%；情境 25／50／25；評等五級。DCF 企業價值 ≤0 時判失效、每股取 0（有限責任）。
- **對帳（Recon）：** 公司指引、Q3 實況驗證、J4 公司宣稱新約價格、RPO、IREN 類比、Wattonomics、舊法（v4.13 融資調整 DCF，不進結論）。

## 已知限制

Q3 10-Q 在建置環境無法開啟：債務逐檔、到期表、租賃、在建工程、期末股數、RSU 與選擇權沿用 Q2 10-Q 或標待補；Q3 公司口徑資本支出、RPO 分桶、加權債務成本為法說數。Comparable 同業 3 家（Nebius、Oracle、IREN；MarketScreener 共識，估值日 2026-11-11），淨負債口徑影響大（Val 並列兩口徑）。β 以 US 伺服器／EMS 組代理。kd 為發行票面利率。利息以期初餘額計（新增借款當年不計息，FY2027 約少算 1.1，Recon 標示）。
