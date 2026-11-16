# Nebius（NBIS）個股模型 v0.6.3（現價與同業報價改 2026-11-13 收盤；v0.6.2：Wattonomics 改讀 v0.3.5；v0.6.1：每 MW 收入比照 CoreWeave；v0.6.0：CoreWeave v5.0.3 新雲樣板重建＋2026Q3 滾入＋三層協調）

**v0.6.3（2026-11-16）：現價與同業報價改 2026-11-13 收盤（工作單 px1113）。** 價格取自 `markets/invest_map/data/closes_20261020_20261113.csv`（Yahoo Finance 日線收盤）。
- 現價 128.64（2026-11-12）→ **115.88**；CoreWeave 60.77 → 53.13、Oracle 128.38 → 125.35（EV＝市值＋淨負債，公式重算）。**未換價同業：IREN**（CSV 無此代號，維持 2026-11-11 收盤 20.84，Val 估值日欄標「未換價」）。
- 同業前瞻 EV/EBITDA 中位數 5.62 → **5.62**（中位數仍是 IREN；CoreWeave 5.01 → 4.73、Oracle 7.86 → 7.71）；Comparable 目標價 60.62 → **58.14**；加權目標價 30.31 → **29.07**；上檔空間 −76.4% → **−74.9%**；評等 強力放空 → **強力放空**。
- rf、共識、淨負債、股數、營運假設、情境權重、三層基準讀數不變。Sens、RevDCF 快照已重跑，拆解沿用。
- 驗證：`SENS=1 REV=1 JOBS=1 bash nebius/make_v1.sh`。

**v0.6.2（2026-11-15）：Wattonomics 快照改讀 v0.3.5（原 v0.3.2）。**
依據：主控指示（main 已合併 Wattonomics v0.3.5，#618）；規則 `docs/plan/三層基準協調規則.md` v1.4.2；`wattonomics/docs/verify/20261115_v0.3.5_build.md`。
- **改了什麼：** `build_nebius.py` 讀 `wattonomics/data/wattonomics_snapshot_v0.3.5.json`（In 頁 wat26、wat30，Recon 原則 54 列，Coord 頁）；追公式重跑 `tools/coord/trace_wattonomics.py wattonomics/model/20261115_Wattonomics_v0.3.5.xlsx nebius/data/wattonomics_trace_v0.3.5.json --player nbis`（無法解析的公式 0 格；舊 v0.3.2 追公式檔保留）。
- **層級讀數新舊值（v0.3.2 → v0.3.5）：** In 頁 gcap_nbis 2026 由 Nebius v0.5.2 回寫 26.046 重匯為 **v0.6.1 Thesis!E66 的 24.185**；`WIF_Gap_Player_nbis_2026_CapexAI` 13.17 → **20.56**（Gap 主讀法改付款年實際花費＝24.185 × AI 占比 0.85；舊值是上線年 GW 回推）；`WIF_Gap_Player_nbis_2026_CapexAIModel` 26.046 → 24.185；`WIF_PlayerGW_YE_nbis_2026` 0.4365 → **0.4214**（CapexGW 頁 G227＝F227＋G226，原 G221／G220）；`_2030` 2.364 → **2.482**。
- **Coord 頁同源說明更新：** 層級 nbis 的 2026 年資本支出現在就是本公司模型自己的值（同源），只列參考；C0 結論不變——營收、資本支出「參考（同源校準）：無可比實績（層級值同源）」，在役電力「無可比實績」，對帳中 0 個、有判定 0 個。讀數（參考）：資本支出主讀數 前三季 15.431 vs 20.557 × 9 ÷ 12＝15.418（+0.1%），第二讀數 Q3 × 4＝29.20 vs 20.56（+42.0%）；營收主讀數 1.938 vs 2.222（−12.7%），第二讀數 3.895 vs 2.962（+31.5%）。G1 的 CapexAIModel 列差距 0（本版基準輸入與 v0.6.1 相同）。v0.6.0 節「交主控重匯回寫值」一項已由 Wattonomics v0.3.5 完成。
- **目標價與評等：** Wattonomics 不進估值，基準輸入未變：加權目標價 **30.31 → 30.31**、投資建議 **強力放空 → 強力放空**；收支關卡、反向 DCF 19.70 不變。只有 In 頁 wat26（436.5 → 421.4 MW）、wat30（2,363.7 → 2,481.9 MW）與 Recon 對照列、Coord 頁數值與文字變動。
- **驗證：** `SENS=1 REV=1 bash nebius/make_v1.sh` 全過（公式錯誤 0；檢查列與收支恆等式 157 項；In 版面；敏感度覆蓋 41 鍵；分部；HTML；治理；收支關卡評等；敏感度與反向 DCF 快照已重跑，拆解快照沿用）；`check_company_html nebius`、`check_layout_v13 --xlsx`、`tools/baseline_coordination/check.py` 見 PR 說明。

**v0.6.1（2026-11-15）：每 MW 收入比照 CoreWeave 辦理（Andy 決定，2026-11-15）。**
依據：Andy 2026-11-15「每 MW 收入比照 CoreWeave 辦理」；CoreWeave v5.0.3 做法（`coreweave/README.md` 方法要點）：每 MW 年收入＝Tokenomics `IF_HoldEcon` 在役世代加權 × 單一定價倍數 k；k＝IREN–Microsoft 長約〔Analogy〕與 H100 現貨加權；不以營收、ARR、RPO、ACV、指引反推 k（本資料夾 CLAUDE.md 第 6 條）；公司宣稱的新約每 MW 年收入只列參考、作敏感度，並開 Drive 對帳單（比照 CoreWeave J4 對帳單 `20261113_對帳單_CoreWeave_J4新約價格與Wattonomics上線GW.md`；Nebius 的對帳單由 Andy 開立）。
- **（1）k 的構成比對（逐項對 `coreweave/build_coreweave.py` 與 v5.0.3 xlsx In 頁）：** Tokenomics 快照同檔（`data/tokenomics_snapshot_v22.0.json`，sha256 ad0decd3… 兩資料夾相同）；錨＝在役世代加權 `IF_HoldEcon`，Thesis 錨列公式（期初＋期末存量平均加權、FY2026 欄 9M／Q4 0.75／0.25）逐字相同；長約價格證據 p_long＝9.70（IREN–Microsoft，Analogy；低／高 8.9073／12.7247）、現貨 p_spot＝2.82（H100 Silicon Data）、gh_h100＝1.5426、he_gb300＝10.8864，k_長約＝9.70 ÷ 10.8864＝0.891、k_現貨＝1.828，k＝隨需占比 × k_現貨＋（1 − 隨需占比）× k_長約，k 調整係數基準 1（×0.9／×1.2）——**同法、同參數**。唯一不同：隨需占比 Nebius 0%（CoreWeave 10%，該值本身在 CoreWeave 也待 Andy 決定），依 Andy 指示屬 Nebius 公司特有值保留，所以 k＝0.891（CoreWeave 0.985）。世代占比、計費比例（0.92，兩家同值）、在役 MW 路徑亦屬公司事實，保留。**結論：不需改值，基準 k、每 MW 年收入（FY2027 情境 B 9.63、FY2030 10.48）與目標價不變。** 參考：若隨需占比改成 CoreWeave 的 10%，加權目標價 39.59（評等仍強力放空；整本重算一次，未入版）。
- **（2）公司宣稱與敏感度：** 公司稱 Q3 新約 ACV 每 MW 超過 US$25m（Q3 股東信，Interested-party；分母口徑未說明）只列 Recon D 節與 Sens。新增 In 鍵 `k_claim`（0＝IREN 長約類比；1＝長約價改用公司宣稱 25，即 k_長約＝25 ÷ 10.8864＝2.296，作用於全部在役 MW）與 Sens 案例「定價倍數改以公司說法」：加權目標價 409.12（+1,250%；估值評等強力買進），但情境 B 收支合計 −22.9，仍入不敷出，投資建議依收支關卡不得列為買進。原有案例「新增 MW 改用公司宣稱 ACV」（p_new_ovr＝25，只作用於 2026-09-30 後新增 MW）251.57（+730%；收支合計 −41.4）。橫幅與命題章（新增說明列 README ksens13）寫明：**評等對每 MW 收入高度敏感；要支撐現價需 2030 年每 MW 年收入 ≥ 19.70 百萬美元**（反向 DCF 反解 k 調整 1.881 × 模型 10.48；低於公司宣稱 25）。
- **（3）新舊目標價與評等：** v0.6.0 加權目標價 **30.31**（上檔 −76.4%）、投資建議**強力放空** → v0.6.1 **30.31**、**強力放空**，差異 0。原因：比對後 k 的構成已與 CoreWeave 同法同參數，本版只加敏感度案例與文字，基準輸入不變（`tools/xlsx_diff.py` 只有版號、build_id、新增列與說明文字差異）。
- **驗證：** `SENS=1 bash nebius/make_v1.sh` 全過（公式錯誤 0；檢查列與收支恆等式 157 項，含新增的 k 構成對 CoreWeave 現行 xlsx 逐鍵比對 10 項；In 版面；敏感度覆蓋 41 鍵；分部；HTML；治理；收支關卡評等）；`check_company_html nebius` 全過；`check_layout_v13 --xlsx` 通過（僅警告：估值章 CONS_NBIS、對帳章層級名稱，同 v0.6.0）；`tools/baseline_coordination/check.py` 0 項不符（全樹）。`tools/xlsx_diff.py` 對 v0.6.0：數值差異只來自 In 頁插入 k_claim 一列造成的列位移與 README 新增列（橫幅、命題章說明、版本列）、Sens 新增一案；加權目標價 30.31、評等強力放空、收支關卡 A／B／C、反向 DCF 19.70 不變。反向 DCF 與拆解快照沿用（基準輸入未變）。

**v0.6.0（2026-11-15）：依 CoreWeave 新雲樣板重建為 xlsx 事實基礎＋HTML 展示層，滾入 2026Q3 實績，版面 v1.3＋Coord 頁（三層協調、C0 實績）。**
依據：工作單 `docs/workorders/20261114_第二批13家_v1.3與三層協調.md`（第二波：Nebius ← CoreWeave）；規則 `docs/plan/三層基準協調規則.md` v1.4.2 C0 第 2 點；`tools/coord/README.md`；主控指示（預設 a）：舊 JS 引擎改依 CoreWeave 樣板（`coreweave/`，v5.0.3）重建並滾入 3Q26 實績，數字可改變，列出新舊目標價與評等差異及原因。

- **做法：** 舊 JS 雙引擎（v0.5.3：segA–E.js、company.json、build_xlsx.py 等）整套移到 `legacy_v05/`（保留不刪，含舊 README 與 CLAUDE.md），v0.5.3 成品移 `dist/_待刪除/`。新建 `build_nebius.py`（自 `coreweave/build_coreweave.py` 複製，換成 Nebius 的 In 頁、Q 頁、站點、債務、同業與 Coord），`scripts/`（runner、sens、reverse、decomp、check、make_html_json，自 CoreWeave 複製）、`make_v1.sh`、`segments.json`、`display_check.json`。事實總帳 `data/nebius_facts_20261115.json`、標準損益表歷史 `data/nebius_is_hist_20261115.json`、追公式輸出 `data/wattonomics_trace_v0.3.2.json`；Tokenomics 快照改讀 v22.0（與 CoreWeave 樣板同檔，舊版讀 v19.0）。
- **Nebius 與 CoreWeave 樣板的結構差異（工程決定，均在 In 頁標記與說明）：**（1）公司未揭露 2026 年各季在役 MW，2026-06-30 240、2026-09-30 330（區間 280–400）、2026 年底 550 MW 為明列假設（Assumed；推估依據寫在 In 頁）；（2）分部兩列：Nebius AI cloud（命題模組驅動）＋其他事業（Avride、TripleTen、沖銷；Q3 年化、零毛利）；（3）客戶預付：2026-09-30 遞延收入 9.409 自 Q4 起 3 年直線認列（非現金營收），Q4 預付流入 1.5、FY2027 起預付＝資本支出 × 40%（公司 2026／2027 預付對資本支出比的約數，Interested-party）；（4）可轉債列負債，NVIDIA 預付認股權證 21.07 百萬股計入股數；（5）非上市權益證券投資 1.604（ClickHouse 等，帳面）加入 DCF 與 Comparable 股權價值；（6）會計折舊年限 5 年（公司 Q3 由 4 年改 5 年）；荷蘭稅率 25.8%，NOL 未取得、基準 0；（7）同業：CoreWeave、Oracle、IREN（同業估值日 2026-11-11，比本模型現價日 2026-11-12 早一個交易日；CoreWeave 共識為其財報前 2026-09-25）；（8）營業費用率 20%（Q3 實績 30.3%，含 Avride 等，Assumed 規模效應）。
- **實績與出處（Q3 2026 6-K，2026-11-12；SEC EDGAR accession 0001104659-26-126008）：** 附件 99.1（財報）：合併營收 Q3 0.9808、9M 1.9621；營業利益 Q3 0.0504、9M −0.2535；調整後 EBITDA Q3 0.4658、9M 0.8315（非 GAAP）；續營淨利 9M 0.3933（含權益證券重估利益 0.7806）；折舊攤銷 9M 0.8264；股票薪酬 9M 0.1965；利息費用 9M 0.2953、利息收入 0.0685；營運現金流 9M 4.2311；購置 PP&E 與無形資產 9M 15.4307（Q1 2.4729、Q2 5.6574、Q3 7.3004）；現金 7.4256、受限現金 1.2155、PP&E 淨額 19.3419、借款 14.1908、遞延收入 9.4092、非上市權益證券 1.6039；流通股 292,902,405。附件 99.2（股東信）：AI Cloud 營收 Q3 0.9738（含非經常性 asset-light 約 0.325）；RPO 38.7；ARR 6.4；簽約電力超過 4 GW（年底目標 5 GW）；ACV 每 MW 超過 US$25m；指引 2026 營收 3.0–3.4 下緣、調整後 EBITDA 率約 40%、資本支出約 27；2027 營收 12–13、資本支出 40–45；2026 預付約 10、2027 預付 17–20；季末在役 MW 未揭露。FY2023–FY2025 損益取 20-F 合併損益表（續營部門）。現價 128.64（2026-11-12 收盤，MarketScreener 5 日表，經 CoreWeave In 登錄）。
- **新舊目標價與評等（主控要求列示）：** v0.5.3（2026-11-06，舊 JS 引擎、現價 120.37）加權目標價 **10.30、評等放空** → v0.6.0（現價 128.64）加權目標價 **30.31（上檔 −76.4%）、投資建議強力放空**（估值評等強力放空；收支關卡情境 B −104.4 不通過，但評等已低於買進，不改列）。拆解（`scripts/decomp_v50.py`，兩段相加＝總變動 +20.01）：① 新樣板＋Q3 實績＋Tokenomics v22.0（以舊現價 120.37 計）+19.20 → 29.50；② 現價 120.37 → 128.64 +0.81。原因：舊引擎目標價以 EV/EBITDA 6x（FY29 錨定，權重較高）＋融資調整 DCF 為主、且評等規則為三級（空間 ≤ −15% 放空）；新樣板為 DCF 50%＋Comparable 50%（新雲同業中位數 5.62 倍 × FY2028 調整後 EBITDA，錨定 FY2027 年底）、評等五級（上檔空間 < −50% 為強力放空）。DCF 三情境皆失效（企業價值 ≤0，每股取 0），目標價全部來自同業倍數；評等由放空變強力放空是因為新樣板五級門檻，而非價值上調（目標價反而上升）。舊引擎與新樣板結構不同，① 未再細拆。
- **C0（實績）判定（Coord 頁 G2；追公式 `tools/coord/trace_wattonomics.py --player nbis`，main 版工具）：** Wattonomics v0.3.2 的 Nebius 2026 年底上線 GW（0.4365）＝2025 年底 0.17（公司 20-F）＋2026 年新增 0.2665，新增量由 2026 年資本支出換算，而資本支出讀 In 頁 gcap_nbis＝**Nebius v0.5.2 個股模型回寫值 26.046**（生效分支；公司指引 cap_nbis 22.5 分支未生效）；玩家資本支出 CapexAI 2026＝13.167 也由此推得；Nebius 沒有 Q3 實績下限（公司不揭露在役 MW）。依同源規則：
  - 營收（AI cloud）：主讀數 前三季 1.938 vs 持有成本隱含 CY2026 3.038 × 9 ÷ 12＝2.278（−14.9%）；第二讀數 Q3 × 4＝3.895 vs 3.038（+28.2%）→ **參考（同源校準）：無可比實績（層級值同源）**，不判定。
  - 資本支出（公司口徑＝現金購置，公司未另定應計口徑）：主讀數 前三季 15.431 vs 層級 CapexAI 13.167 × 9 ÷ 12＝9.875（+56.3%）；第二讀數 Q3 × 4＝29.202 vs 13.167（+121.8%）→ **參考（同源校準）**；參考列：對 v0.5.2 回寫值 26.046 為 −21.0%／+12.1%。
  - 在役電力：**無可比實績（公司未揭露季末在役 MW）**；簽約電力超過 4 GW vs Wattonomics ctr_nbis 3.5 GW（公司 Q1 揭露，同源）：無可比實績。
  - **對帳中項目 0 個、有判定 0 個。** 交主控：Wattonomics nbis 的 2026 年量讀舊版 v0.5.2 回寫值，本版合併後重匯回寫值（gcap_nbis 等）時一併評估；前三季現金資本支出年化速率高於層級 CapexAI。
  - G1 讀取層級（同源，不判定）：Tokenomics v22.0 IF_HoldEcon、IF_CapexIT、IF_OpexGW（差距 0）；Wattonomics 2026／2030 年底上線 MW、2026 年 CapexAIModel、ctr_nbis（皆與本公司模型或公司揭露同源）。Recon 頁 wat26、wat30 說明改標同源。
- **其他結果：** 情境目標價 A／B／C 36.86／30.37／23.64；收支關卡 −143.1／−104.4／−71.3；情境 B 2030 年底累計稀釋 79.7%、淨負債 ÷ EBITDA 最差 6.30 倍（WARN）。敏感度最大：新增產能改用公司宣稱 ACV 25（+730%，估值評等強力買進）、長約價格證據高案（+92%）、同業淨負債改共識預估（+90%）。反向 DCF：支持現價需 FY2030 每 MW 年收入 19.70（模型 10.48，k 調整 1.88 倍），仍低於公司宣稱 ACV 25。FY2026 營收 2.88 低於指引下緣 3.9%；FY2027 營收 8.00 對指引 12–13。
- **驗證：** `make_v1.sh` 全過（公式錯誤 0、檢查列與收支恆等式 147 項、In 版面、敏感度覆蓋 40 鍵、分部核對、HTML、治理、收支關卡評等）；`check_company_html nebius`、`check_layout_v13 --xlsx` 通過（僅警告：對帳章 Coord 必要的層級名稱）；`tools/baseline_coordination/check.py` 0 項不符。`tools/xlsx_diff.py` 對 v0.5.3 不適用（重建，工作表結構全換；主控已允許數字改變）。

數字以 `dist/` 的 xlsx 為準；下表由建置成品產生（`python3 tools/model_governance/version_table.py nebius --update`），不手寫（舊版 v0.5.3 檔名非點號版號、結構不同，未入表；其值見上）。

<!-- version-table: dist/_待刪除/20261115_Nebius個股模型_v0.6.2.xlsx｜dist/20261116_Nebius個股模型_v0.6.3.xlsx -->
| 項目 | v0.6.2 | v0.6.3 |
|---|---|---|
| 加權目標價（上檔空間） | 30.31（-76.4%） | 29.07（-74.9%） |
| DCF 目標價（情境加權） | 0.00 | 0.00 |
| Comparable 目標價 | 60.62 | 58.14 |
| 投資建議 | 強力放空 | 強力放空 |
| 現價 | 128.64 | 115.88 |
| 收支關卡 A／B／C（FY2026–FY2030） | -143.1／-104.4／-71.3 | -144.1／-104.8／-71.4 |
<!-- /version-table -->

## 建置

```
bash nebius/make_v1.sh                        # 沿用已提交的敏感度、反解、拆解快照
SENS=1 REV=1 DEC=1 bash nebius/make_v1.sh     # 改了公式或輸入時（約 5 分鐘）
```

## 檔案

| 檔案 | 內容 |
|---|---|
| `build_nebius.py` | Python → Excel 產生器（README → In → Q → Thesis → Seg → PL → CF → Fund → RPO → Sites → Debt → Lease → Val → RevDCF → Recon → Sens → Sources → Coord） |
| `scripts/` | runner、敏感度、反向 DCF、拆解、檢查、HTML 設定檔產生器（自 CoreWeave v5.0 樣板） |
| `data/nebius_facts_20261115.json` | 事實總帳（Q3 6-K 附件 99.1、99.2 逐筆；來源網址） |
| `data/nebius_is_hist_20261115.json` | 標準損益表 FY2023–FY2025（20-F）與 9M 2026（6-K） |
| `data/wattonomics_trace_v0.3.5.json` | Wattonomics v0.3.5 玩家 nbis 追公式輸出（Coord 頁同源判定依據；v0.3.2 版保留供 v0.6.0–v0.6.1 對照） |
| `data/tokenomics_snapshot_v22.0.json` | Tokenomics 快照（v19.0 保留供舊引擎） |
| `legacy_v05/` | v0.5.3 舊 JS 雙引擎（含舊 README、CLAUDE.md），保留不刪 |

## 已知限制

公司未揭露 2026 年各季在役 MW，2026 年在役 MW 為假設，公司實況驗證只檢驗 MW 假設與 k 的合併結果；Q3 附註（尚未啟用資產、可轉債逐檔變動、RPO 分桶、RSU）未取得，沿用 Q2 6-K 或標假設；同業 3 家、CoreWeave 共識為其財報前；市場共識目標價為財報前（2026-10-07）；β 以 US 伺服器／EMS 組代理；kd 為擔保融資類比。
