# Nebius（NBIS）個股模型 v0.6.0（CoreWeave v5.0.3 新雲樣板重建＋2026Q3 滾入＋三層協調）

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

<!-- version-table: dist/20261115_Nebius個股模型_v0.6.0.xlsx -->
| 項目 | v0.6.0 |
|---|---|
| 加權目標價（上檔空間） | 30.31（-76.4%） |
| DCF 目標價（情境加權） | 0.00 |
| Comparable 目標價 | 60.62 |
| 投資建議 | 強力放空 |
| 現價 | 128.64 |
| 收支關卡 A／B／C（FY2026–FY2030） | -143.1／-104.4／-71.3 |
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
| `data/wattonomics_trace_v0.3.2.json` | Wattonomics 玩家 nbis 追公式輸出（Coord 頁同源判定依據） |
| `data/tokenomics_snapshot_v22.0.json` | Tokenomics 快照（v19.0 保留供舊引擎） |
| `legacy_v05/` | v0.5.3 舊 JS 雙引擎（含舊 README、CLAUDE.md），保留不刪 |

## 已知限制

公司未揭露 2026 年各季在役 MW，2026 年在役 MW 為假設，公司實況驗證只檢驗 MW 假設與 k 的合併結果；Q3 附註（尚未啟用資產、可轉債逐檔變動、RPO 分桶、RSU）未取得，沿用 Q2 6-K 或標假設；同業 3 家、CoreWeave 共識為其財報前；市場共識目標價為財報前（2026-10-07）；β 以 US 伺服器／EMS 組代理；kd 為擔保融資類比。
