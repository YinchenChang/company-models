# CoreWeave（CRWV）個股模型 v5.0.3（Oracle 收支型樣板）

**v5.0.3（2026-11-14）：版面 v1.3 補缺項＋三層協調（Coord 頁；新雲樣板，Nebius 複製用）。** 依據：工作單 `docs/workorders/20261114_第二批13家_v1.3與三層協調.md` 第 2 節；規則 `docs/plan/三層基準協調規則.md` v1.4.2 C0 第 2 點；`tools/coord/README.md`；Andy 2026-11-14 12:47（G、H 位置）、17:40（C0 用公司實績）。open PR #450（v4.14 草稿）不動。
- **版面 v1.3 補缺項（`check_layout_v13 --xlsx` 原缺 G、H）：** 估值章第一個區塊改為四個目標價與現價圖、敏感度章前兩個區塊改為低案、高案主圖，章節說明（原 intro）移到圖下（`scripts/make_html_json.py` 末段；只改區塊順序）。
- **Coord 頁設計（算力買賣方／新雲樣板；不做兩讀數目標價，只有兩節，Nebius 照抄結構、換成自己的玩家與實績）：**
  - **G1 讀取層級（同源，不判定）：** 列本模型讀的 Tokenomics 名稱與版本（v22.0，commit 4a13ba4；In 頁 he_／cx_／ox_ 列直連，差距恆 0）、Wattonomics 對照值（In 頁 wat26、wat30 為手填，來源 v0.3 mw_crwv，已核對 v0.3.2 同值；層級值是公司 2026 年底、2030 年目標上緣，同源）。差距與判定標「讀取層級（同源，不判定）」，不計入對帳中；Recon 頁 wat26、wat30 兩列說明文字同步改標（數值不動）。本輪不換快照。
  - **G2 C0（實績）：** 個股側只用已公布期間的公司實績（2026 年前三季，Q3 新聞稿與 10-Q），層級側用同口徑隱含值，每個量給兩讀數。主讀數＝前三季累計 vs 層級 CY2026 × 9 ÷ 12；第二讀數＝Q3 × 4 vs 層級 CY2026（在役電力是時點量：主讀數＝2026-09-30 實績 vs YE2025→YE2026 線性內插，第二讀數＝實績 vs YE2026）。兩者都超過 ±15% 且同號才「對帳中」，只一個超過為「時差邊界（參考）」。門檻是輸入（藍字）。層級輸入（Wattonomics 名稱與值）以藍字寫入 Coord 頁並標快照版本；Tokenomics 每 MW 持有成本連到 Thesis 頁加權錨列。核對列：營收與資本支出實績對 Thesis 9M 實績差為 0。
  - **同源處理：** 層級側不含本模型的定價倍數 k 與計費比例（公司因素）；資本支出讀 `WIF_Gap_Player_crwv_2026_CapexAIGW`（GW 路徑 × Tokenomics 每 GW 資本支出），不讀 `CapexAIModel`（個股模型回寫值，同源）；不用 `_SiteOnly` 路徑（公司站點目標）。Coord 頁與 HTML 都寫明這些。
  - **HTML：** 對帳章最前新增「實績對帳」（摘要、C0 表、口徑與出處摺疊表、參考列、同源與口徑說明）與摺疊的「讀取層級」表；數字與文字全取自 xlsx（`scripts/make_html_json.py` 的 `R["COORD"]` 列號）。
- **實績與出處：** 營收 2026 年前三季 US$8.314B、Q3 US$3.661B（Q3 新聞稿 8-K 附件 99.1，2026-11-10；10-Q 0001769628-26-000465，Verified）；現金購置 PP&E 前三季 US$23.558B、Q3 US$9.441B（同新聞稿現金流量表，Verified；現金口徑，不含應付設備款與 OEM 融資）；在役電力 over 1.7 GW、簽約電力 over 4.6 GW（同新聞稿，營運指標，Interested-party）。公司口徑資本支出前三季 30.65（上半年 10-Q 16.139＋Q3 法說 14.511）只列參考。
- **層級隱含與判定（US$B；Wattonomics v0.3.2、Tokenomics v22.0）：** 營收層級 CY2026＝（YE2025 0.85＋YE2026 1.70 GW）÷ 2 × 加權每 MW 年持有成本 10.03＝12.78；主讀數 8.31 vs 9.59＝−13.3%、第二讀數 14.64 vs 12.78＝＋14.6%，都在門檻內 → **一致**（第二讀數貼近門檻）。資本支出層級 29.97（WIF_Gap_Player_crwv_2026_CapexAIGW）：主讀數 23.56 vs 22.48＝＋4.8%、第二讀數 37.76 vs 29.97＝＋26.0%，只一個超過 → **時差邊界（參考）**；改用公司口徑 30.65 則 ＋36.3%（參考列，等級混合，不判定）。在役電力：1,700 vs 內插 1,488＝＋14.3%、vs YE2026 1,700＝0% → **一致**。簽約電力 4,600 MW：**無可比實績**（Wattonomics 沒有簽約電力序列）。**對帳中項目 0 個。** 參考：實績隱含每平均在役 MW 年收入 8.69 vs Tokenomics 加權 10.03（−13.3%，與營收主讀數同差，因 Wattonomics 平均在役 GW 與實績平均 MW 相同）；J4 公司宣稱新約每 MW 年收入 20–25（Interested-party）與 Wattonomics 2030 年底 8 GW（公司目標上緣）對情境 B、C 的差距，見 Drive 對帳單 `20261113_對帳單_CoreWeave_J4新約價格與Wattonomics上線GW.md`，只列參考。
- **驗證：** 既有 xlsx 逐格比對只有 README A1 版號與 build_id、README 新增版本列（A67、A68 順移）、Recon I35／I36 說明改標、Sens D56 build_id、新增 Coord 頁；加權目標價 76.20、評等「收支不足」、各情境數值不變。`check_company_html`、`check_layout_v13 --xlsx`、`tools/baseline_coordination/check.py`、`make_v5.sh` 全過（Coord 名稱如 IF_HoldEcon 出現在 HTML 對帳章屬必要的層級名稱，版面檢查僅警告）。

**v5.0.2（2026-11-14，試作）：HTML 套用個股 HTML 版面規格 v1.3。** Andy 2026-11-14 08:11：章節順序「命題與核心問題 → 估值 → 敏感度 → 反向 DCF → 損益 → 驅動因子 → 個別模組（CoreWeave＝資金缺口與融資，收支模組）→ 對帳 → 資料狀態」；「html 是給我念的，所以可讀性是很重要的」。本版固定規格（`scripts/make_html_json.py`，檢查 `tools/company_html/check_layout_v13.py`）：
- 橫幅說明 ≤200 字、不含參數代號（README 頁「展示層 v1.3」hero13 列，公式組字）；原橫幅的級距邊界、評等脆弱點、方法註記移到估值章摺疊。
- 重點觀察 ≤5 點、合計 ≤750 字（README ob1–ob5 白話列；原 rec、o1–o6 保留並摺疊）。
- 每章至少一張圖（反向 DCF、對帳、資料狀態除外）：命題＝收支關卡逐年自由現金流 × 三情境；損益＝三情境營收與營業利益；驅動＝分部營收推導（長約／隨需堆疊＋計費 MW）、每 MW 年收入、backlog 轉營收；資金＝融資來源；估值＝目標價長條；敏感度＝龍捲風圖。
- 損益章＝標準損益表（**GAAP**；FY2023–FY2025 10-K 實績、FY2026＝9M 10-Q 實績＋第四季模型、FY2027–FY2030 情境 B，另列 9M 2026 比較欄）；必備列：營收、營收年增率、營業成本（不含折舊）、毛利、毛利率、營業費用、營業利益、營業利益率、利息費用淨額、其他收入（淨額）、稅前利益、所得稅、稅後淨利、淨利率、EPS（完全稀釋）；表尾比率區另含淨利年增率。歷史數字來源：SEC XBRL（`data/coreweave_is_hist_20261114.json`、Sources 頁 SEC_XBRL）。
- 驅動因子章必含分部營收推導：AI 雲端平台營收＝計費 MW ×（長約每 MW 年收入 ×（1 − 隨需占比）＋隨需每 MW 年收入 × 隨需占比），Seg 頁「分部營收推導」區塊，核對列＝0。
- 可見文字不出現參數代號（In 參數表、Thesis 推導表收進摺疊；龍捲風圖用 Sens 頁 O 欄白話標籤）。
- 版面規格 v1.3 G 節（Andy 2026-11-14 09:03）：橫幅現價標明日期（README px13 列）；估值章開頭並排四個目標價（表＋長條圖＋現價線；Val 頁末「估值章四個目標價」區塊）：DCF（三情境失效取 0 並註明）、Comparable 情境加權（另列未調整值）、加權目標價、市場共識目標價（TIKR 平均 US$132，2026-11-11，Q3 財報後；Sources TIKR_CONS；只作對照）。
- xlsx 只在各頁最後附加列（PL、Seg、CF、Val、Sens O 欄、Sources），以「新增列以外逐格比對 0 差異」驗證（README、版本字串、成品日與其衍生的輸入指紋除外）；目標價 76.20、投資建議「收支不足」不變。

**v5.0.1（2026-11-13）：投資建議加上收支關卡（原則第 10 節第 63 條）。** Andy 2026-11-13「个股模型还是要遵守这个收之不过就不能买进的这个规则……个股模型这边还是要用相同的原则来评断它的評等」。Val 頁新增「估值評等（只看上檔空間）」「入不敷出判定」兩組列；「投資建議」＝收支關卡不通過（情境 B 的 FY2026–FY2030 自由現金流合計 < 0）且估值評等為買進以上時改列「收支不足」，其餘同估值評等。本版結果：估值評等買進（邊界）、收支合計 −95.7 → **投資建議：收支不足**。加權目標價與其他數值逐格不變（76.20）。另更正 xlsx README「已知限制」過時的同業家數、In 頁隨需占比說明（v4.13 基準 0%，v5.0 改 10% 未列入採用規劃書，待 Andy 決定）。

**v5.0（2026-11-12）：** Andy 2026-11-12「Coreweave 已经公布财报了，所以要更新 Coreweave 的财报。再将 html 的版面套用 Oracle 最新的版本。」＋確認 v4.13 驗收採用規劃書（J1–J5 照預設）。工作單 `docs/workorders/20261112_CoreWeave_v5.0_Oracle樣板重建_Q3.md`；報告 `docs/reports/20261112_coreweave_v5.0_Oracle樣板重建與Q3.md`。舊 JS 雙引擎（v4.13，HTML＋Excel、`company.json`）整套移到 `legacy_v4/`（保留不刪）；v4.13 成品移 `dist/_待刪除/`。

數字以 `dist/` 的 xlsx 為準；下表由建置成品產生（`python3 tools/model_governance/version_table.py coreweave --update`），不手寫。

<!-- version-table: dist/_待刪除/20261114_CoreWeave個股模型_v5.0.2.xlsx｜dist/20261114_CoreWeave個股模型_v5.0.3.xlsx -->
| 項目 | v5.0.2 | v5.0.3 |
|---|---|---|
| 加權目標價（上檔空間） | 76.20（+25.4%） | 76.20（+25.4%） |
| DCF 目標價（情境加權） | 0.00 | 0.00 |
| Comparable 目標價 | 152.41 | 152.41 |
| 投資建議 | 收支不足 | 收支不足 |
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
| `build_coreweave.py` | Python → Excel 產生器（18 頁，v5.0.3 起末尾加 Coord 頁；v5.0.2 起 PL、Seg 末尾另有標準損益表與分部營收推導：README → In → Q → Thesis → Seg → PL → CF → Fund → RPO → Sites → Debt → Lease → Val → RevDCF → Recon → Sens → Sources） |
| `scripts/runner_v50.py` | 整本重建＋LibreOffice 重算的共用工具 |
| `scripts/sens_v50.py`、`reverse_v50.py`、`decomp_v50.py` | 敏感度、反向 DCF（反解 k_adj）、目標價變動拆解 → `data/*_v50.json` |
| `data/coreweave_facts_20261112.json` | 事實總帳（Q3 新聞稿逐筆、Q2 10-Q 沿用項、來源 URL） |
| `data/coreweave_is_hist_20261114.json` | 標準損益表的 FY2023–FY2025 與 9M 2026 實績（SEC XBRL，10-K／10-Q） |
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
