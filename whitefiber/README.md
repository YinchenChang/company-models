# WhiteFiber（WYFI）個股模型 v0.4.0（CoreWeave v5.0.3 新雲樣板重建＋2026Q3 滾入＋三層協調）

**v0.4.0（2026-11-15）：依 CoreWeave 新雲樣板（經 Nebius v0.6.0 分部版）重建為 xlsx 事實基礎＋HTML 展示層，滾入 2026Q3 實績，版面 v1.3＋Coord 頁（三層協調、C0 實績）。**
依據：工作單 `docs/workorders/20261114_第二批13家_v1.3與三層協調.md`；規則 `docs/plan/三層基準協調規則.md` v1.4.2 C0 第 2 點；`tools/coord/README.md`；主控指示（預設 a）：舊 JS 引擎改依 CoreWeave 樣板重建並滾入 3Q26 實績，數字可改變，列出新舊目標價與評等差異及原因。版號：v0.3 由 open PR #451（`claude/whitefiber-2026q3`，舊格式 Q3 滾入草稿）占用，本版用 v0.4.0；#451 不動。

- **做法：** 舊 JS 雙引擎（v0.2.1）整套移到 `legacy_v02/`（含舊 README、CLAUDE.md），v0.2.1 成品移 `dist/_待刪除/`。新建 `build_whitefiber.py`（自 `nebius/build_nebius.py`〔CoreWeave 樣板的分部版〕複製，換成 WhiteFiber 的輸入、站點、債務、同業、Coord），`scripts/`、`make_v1.sh`、`segments.json`、`display_check.json`。事實總帳 `data/whitefiber_facts_20261115.json`、損益歷史 `data/whitefiber_is_hist_20261115.json`、追公式 `data/wattonomics_trace_v0.3.2.json`；Tokenomics 改讀 v22.0（舊版 v19.0）。**事實數字轉錄自 #451 分支事實總帳**（原文 10-Q 2026 Q3〔SEC accession 0001213900-26-117728〕、10-K FY2025、法說二手逐字稿；本版只轉錄、未改數字）。
- **與 CoreWeave 樣板的結構差異（工程決定，均在 In 頁標記）：**（1）金額百萬美元、股數百萬股；（2）雲端在役 MW 公司未揭露：2026-09-30 6.5（區間 5.5–8）、2026 年底 14.5 MW 為假設（以已簽合約 GPU 數經 Tokenomics 功耗換算）；情境 B 2027 年底 35 MW（已簽約約 37 MW，對應公司 2027 年中雲端年化 365 目標）、其後每年 +8；A 40／+15；C 只計已具名合約 25；（3）雲端機房全向第三方承租（自建 0、每 MW 年租金 1.9，以公司自身託管租金類比）；（4）託管與其他（NC-1 Nscale 10 年 865 直線每年 86.5、MTL-1／MTL-3、代管服務打折 30／50）走 Thesis 託管列與 Seg 第二分部，託管 EBITDA 率 65%、託管資本支出 Q4 70（NC-2／NC-3 收購）、2027 年 15、其後 5；NC-2／NC-3 開發與 NC-1 擴充未簽約、不計入；（5）營業費用率 20%（Q3 實績約 59%，Assumed 規模效應）；（6）在建工程 312.3 依 20 年折舊（多為 NC-1、MTL-2 建物；Sens 5 年）；（7）客戶預付覆蓋比 40%、遞延收入 306.0 依 5 年認列；（8）Comparable 改用 4 家新雲同業單一倍數（CoreWeave、Nebius、IREN、Oracle；中位數 6.74 倍），不再分部加總（舊版託管 20.75 倍＋雲端 6 倍）。
- **實績與出處（10-Q 2026 Q3）：** 營收 Q3 24.017／9M 74.779（雲端 Q3 18.106 含終止費 3.4、9M 58.679 含終止費 15.7；託管 Q3 5.617）；營業損失 9M −32.678；調整後 EBITDA 9M 8.512（非 GAAP）；淨損 9M −128.453（含誘導轉換費用 80.767）；折舊攤銷 9M 21.085；營運現金流 9M 211.312（含客戶預付 226.524）；購置 PP&E 與設備訂金 9M 549.804；現金 138.544；借款帳面 402.633（可轉債 330.088）；遞延收入 306.034；在建工程 312.326；流通股 45.281 百萬股；雲端 RPO 716.3（年度分布已揭露）；期後：DDTL 追加 70、NC-2／NC-3 收購 60、NC-1 信用協議 300（未動用）。法說（Interested-party）：已簽 GPU 約 20,000 顆、Q4 雲端營收約持平 18.1、2027 年中雲端年化 365、代管服務年化 50、預付約占資本支出 40%。現價 9.21（2026-11-10，Investing.com／Stockscan）；共識目標價 40.45（MarketScreener，11 家）。
- **新舊目標價與評等（主控要求列示）：** v0.2.1（2026-11-06，舊 JS 引擎、現價 7.80）加權目標價 **15.87、評等中立**（上檔 +103%，因股權募資規則列中立）→ v0.4.0（現價 9.21）加權目標價 **5.39（上檔 −41.5%）、投資建議放空**（收支關卡情境 B −1,513 不通過；估值評等已是放空，不改列）。拆解（`scripts/decomp_v50.py`，兩段相加＝−10.48）：① 新樣板＋Q3 實績＋Tokenomics v22.0（以舊現價 7.80 計）−10.63 → 5.24；② 現價 7.80 → 9.21 +0.14。原因：舊版以分部加總（託管 EBITDA × 20.75 倍＋雲端 6 倍）與 β 2.5 的 DCF，新樣板為 DCF 50%（三情境企業價值 ≤0，取 0）＋Comparable 50%（新雲中位數 6.74 倍 × FY2028 調整後 EBITDA 197.5，扣 FY2027 年底淨負債 676），淨負債與融資缺口大、DCF 失效使目標價大幅下降；評等規則改五級（−25% 至 −50% 為放空）。託管若按資料中心同業約 20 倍評價，Comparable 會明顯提高（判斷題，見下）。
- **C0（實績）判定（Coord 頁 G2；`tools/coord/trace_wattonomics.py --player wyfi`，main 版工具）：** Wattonomics v0.3.2 的 WhiteFiber（雲端）年底上線 MW 直接讀 In 頁 mw_wyfi（2025 年 5.5、2026 年 14.574、2027 年 16.548），註記來源為 **whitefiber/company.json scenarios.mwPath（舊版 v0.2 個股模型）**；玩家資本支出 CapexAI 由該路徑 × Tokenomics 每 GW 資本支出推得；cap_wyfi（公司實績與指引）空白 → 同源。
  - 雲端營收：主讀數 前三季 58.68 vs 持有成本隱含 CY2026 98.54 × 9 ÷ 12＝73.91（−20.6%）；第二讀數 Q3 × 4＝72.42 vs 98.54（−26.5%）→ **參考（同源校準）：無可比實績（層級值同源）**。
  - 資本支出：主讀數 前三季 549.80 vs 層級 CapexAI 315.70 × 9 ÷ 12＝236.78（+132.2%）；第二讀數 Q3 × 4＝820.42 vs 315.70（+159.9%）→ **參考（同源校準）**；口徑另不同（公司含 NC-1 託管建置與設備訂金，層級只含雲端）。參考列：對舊版回寫值 CapexAIModel 473.24 為 +54.9%／+73.4%。
  - 雲端在役電力：**無可比實績（公司未揭露季末在役 MW）**；已簽約雲端容量（約 37 MW，GPU 換算）：無可比實績（層級無簽約序列）。
  - **對帳中項目 0 個、有判定 0 個。** 交主控：Wattonomics wyfi 的 MW 路徑與資本支出讀舊版個股模型，2028 年起持平 16.5 MW，明顯低於本版已簽約約 37 MW；本版合併後重匯 mw_wyfi、gcap_wyfi 時一併評估，並考慮改由公司揭露（已簽 GPU 數）推得。
  - G1 讀取層級（同源，不判定）：Tokenomics v22.0 IF_HoldEcon、IF_CapexIT、IF_OpexGW（差距 0）；Wattonomics 2026／2030 年底上線 MW、2026 年 CapexAIModel。
- **其他結果：** 情境目標價 A／B／C 6.21／5.49／4.36；收支關卡 −2,113／−1,513／−762（百萬美元）；情境 B 2030 年底累計稀釋 74%、淨負債 ÷ EBITDA 最差 6.25 倍（WARN）。敏感度最大：新增產能改用合約換算每 MW 收入 17.5（+255%，強力買進）、長約價格證據高案（+124%，買進）、營業費用率 40%（−100%）。反向 DCF：支持現價需 FY2030 每 MW 年收入 11.98（模型 10.11，k 調整 1.18 倍），低於合約換算中點 17.5。對照：FY2026 營收 118.3 對 114.7（＋3.1%）、Q4 雲端 21.7 對公司預期 18.1、2027 年底雲端年化 349.5 對 2027 年中目標 365。
- **驗證：** `make_v1.sh` 全過（公式錯誤 0、檢查列與收支恆等式 144 項、In 版面、敏感度覆蓋 44 鍵、分部、HTML、治理、收支關卡評等）；`check_company_html whitefiber`、`check_layout_v13 --xlsx` 通過（僅警告：對帳章 Coord 必要的層級名稱）；`tools/baseline_coordination/check.py` 0 項不符；version_table 一致。`tools/xlsx_diff.py` 對 v0.2.1 不適用（重建）。
- **需 Andy 決定：**（1）Comparable 是否改回分部加總（託管按資料中心同業約 20 倍）——預設：維持新雲單一倍數（樣板一致），託管倍數列判斷題；（2）營業費用率 20%、託管 EBITDA 率 65%、代管服務打折值為假設；（3）零履約價買權 5.9 百萬股是否自評價股數扣除——預設不扣。

數字以 `dist/` 的 xlsx 為準；下表由建置成品產生（`python3 tools/model_governance/version_table.py whitefiber --update`），不手寫（舊版 v0.2.1 檔名非點號版號、結構不同，未入表；其值見上）。

<!-- version-table: dist/20261115_WhiteFiber個股模型_v0.4.0.xlsx -->
| 項目 | v0.4.0 |
|---|---|
| 加權目標價（上檔空間） | 5.39（-41.5%） |
| DCF 目標價（情境加權） | 0.00 |
| Comparable 目標價 | 10.77 |
| 投資建議 | 放空 |
| 現價 | 9.21 |
| 收支關卡 A／B／C（FY2026–FY2030） | -2,113.5／-1,513.3／-762.1 |
<!-- /version-table -->

## 建置

```
bash whitefiber/make_v1.sh                        # 沿用已提交的敏感度、反解、拆解快照
SENS=1 REV=1 DEC=1 bash whitefiber/make_v1.sh     # 改了公式或輸入時（約 5 分鐘）
```

## 檔案

| 檔案 | 內容 |
|---|---|
| `build_whitefiber.py` | Python → Excel 產生器（README → In → Q → Thesis → Seg → PL → CF → Fund → RPO → Sites → Debt → Lease → Val → RevDCF → Recon → Sens → Sources → Coord） |
| `scripts/` | runner、敏感度、反向 DCF、拆解、檢查、HTML 設定檔產生器 |
| `data/whitefiber_facts_20261115.json` | 事實總帳（轉錄自 #451 分支；10-Q、10-K、法說） |
| `data/whitefiber_is_hist_20261115.json` | 標準損益表 FY2024–FY2025（10-K）與 9M 2026（10-Q） |
| `data/wattonomics_trace_v0.3.2.json` | Wattonomics 玩家 wyfi 追公式輸出 |
| `legacy_v02/` | v0.2.1 舊 JS 雙引擎（保留不刪） |

## 已知限制

雲端在役 MW 公司未揭露（2026 年為假設）；GPU 伺服器租賃視為自購；託管不分情境、NC-2／NC-3 未計入；同業 4 家、CoreWeave 共識為其財報前、WhiteFiber EBITDA 共識找不到；FY2023 損益找不到；β 以 US 伺服器／EMS 組代理。
