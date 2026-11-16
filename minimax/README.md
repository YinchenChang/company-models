# MiniMax 收支模型（company-models/minimax）

MiniMax（00100.HK，上海稀宇科技）收支模型，現行 **v0.2.4**（2026-11-16：現價與同業報價改 2026-11-13 收盤；v0.2.3 為 2026-11-15 估值頁；v0.2 為 2026-10-22）。比照 `openai/`（v0.6／v0.6.1）的分層與 MW 主軸建立；MiniMax 為上市公司，另加 Funding 與 Valuation；v0.2 新增 P×Q 回本比值頁。

命題：MiniMax 每 MW 算力的營收能否覆蓋每 MW 全成本；若不能，缺口何時出現、由誰、以什麼條件補。FY2025–FY2030，曆年制。

- Excel 為唯一計算引擎：`build_xlsx.py` 只產生結構與公式（藍字輸入、黑字公式），數值由 LibreOffice 重算。
- 一鍵建置：`scripts/build_all.sh [日期] [版本]` → `dist/<日期>_MiniMax收支模型_<版本>.xlsx／.html`（需 `libreoffice-calc-nogui`、openpyxl）。
- 敏感度：`scripts/sensitivity.py`（逐項改寫 Inputs、重算，寫入 Sensitivity 頁快照）。
- 原始數據與來源：`data/research/financials.md`、`operations.md`（研究代理 2026-10-08 蒐集；一手／二手已標註）。
- Tokenomics 取數：`scripts/tk_snapshot.py`（讀 `model/CURRENT` 具名範圍）、名稱檔 `data/tokenomics_names.txt`、指標檔 `data/tokenomics_snapshot_CURRENT`；歸因 `scripts/tk_attribution.py`；驗收 `scripts/ci_check.py`；CI `.github/workflows/minimax-build.yml`。
- 工作規範：`CLAUDE.md`；變更紀錄：`CHANGELOG.md`；進度：`docs/reports/20261022_minimax_進度.md`。
- 報告：`docs/reports/20261008_minimax_v0.1.md`（v0.1）、`20261022_minimax_M1.md`、`_M2.md`、`20261022_minimax_v0.2.md`（v0.2 成品）＋ `_對照.xlsx`；命題與驅動對照表 `docs/plan/20261022_MiniMax_命題與驅動對照表.md`；交接檔 `docs/handoff/MiniMax_handoff.md`。
- HTML 驗收：`scripts/check_html.py`（每個數字對照 Excel、無外部引用；已納入 `build_all.sh` 與 CI）。

| 頁 | 內容 |
|---|---|
| README | 命題、驅動→推導對照、圖例、已套用的預設 |
| SRC_MM | 公司財務、市場與營運原始數據（112 列；標記、一手／二手、來源） |
| TK_Link | Tokenomics 官方快照（`data/tokenomics_snapshot_*.json`；現用 v19.0，含檔名、提交、SHA-256、讀取日）；12 格基準值＋IF_HoldEcon 全 15 欄 |
| Inputs | 假設 32 項（值、低、高、標記、說明） |
| Demand | API 實收單價、計費 token |
| Revenue | 企業端（API）、消費端（產品與收費方式按 9M25 結構示意） |
| Compute | 機隊組合、每 MW 產能、η 校準（2025、1H26）、推論／訓練 MW、VR 等值；D2 租價並列、η 對照、自建集群註記 |
| PnL | 損益（算力／人事／其他）、每 MW 營收 vs 每 MW 全成本 |
| PxQ | 回本比值＝營收 ÷（年均上線 MW × 每 MW 年持有成本）、實際／所需 m×L、分項比值、持平 MW、M3.1 用量倍數、比值對 I02 的結構分解 |
| Funding | 現金路徑、需新融資、稀釋、已投入資金 |
| Valuation | 市值、EV、同業、反向檢驗 |
| Checks | 13 項檢查＋錯誤值計數（皆為 0） |
| Sensitivity | 31 個單項情境（快照值；含 M3.1 降價與 I02 租價證據價） |
| PL | 標準損益表 FY2023–FY2030（v0.2.2；非 IFRS 經調整；版面 v1.3 D 節） |
| Coord | 三層協調：讀取層級、C0（實績）公司實績 vs 層級隱含值（v0.2.2；規則 v1.4.2） |
| Val | 估值頁：DCF（FCFF）＋同業 EV／NTM 營收 → 加權目標價、評等、收支關卡（v0.2.3） |
| RevDCF | 反向 DCF：現價隱含什麼（v0.2.3） |

> 標籤修正批 labels（2026-11-17，同版同檔名原地重建，數值不變）：Val!A1 標題版號 v0.2.3 → v0.2.4（現行版）；Val 快照列日期戳記隨重跑更新為 2026-11-17；Coord 頁 Wattonomics 快照版本字樣隨 repo 現行檔 v0.3.5 → v0.3.6（只讀取標籤）。

## v0.2.4（2026-11-16）：現價與同業報價改 2026-11-13 收盤（工作單 px1113）

價格取自 `markets/invest_map/data/closes_20261020_20261113.csv`（Yahoo Finance 日線收盤）。

| 項目 | v0.2.3 | v0.2.4 |
|---|---|---|
| MiniMax 現價（HK$） | 170.2（11-05） | 152.4（11-13） |
| 智譜現價（HK$，唯一同業） | 476.0（11-05） | 408.4（11-13） |
| 同業 EV／NTM 營收中位數 | 9.74 | 7.99 |
| Comparable 目標價（HK$） | 234.82 | 202.13 |
| DCF 目標價（HK$） | 60.92 | 60.92 |
| 加權目標價（HK$）與上檔空間 | 147.87（−13.1%） | 131.52（−13.7%） |
| 估值評等／投資建議 | 中立／中立 | 中立／中立 |

- 未換價同業：無（唯一同業智譜已換 11-13）。rf、β、ERP、共識、匯率、營運假設、情境權重不變；收支關卡不變（−5,983.2，不通過）。
- Val 的目標時點 I44（1.35 年）、NTM 權重 I49（0.153425）是依 2026-11-05 推出的日期衍生參數，工作單只換價格，未重算（影響小）。
- 驗證：`scripts/build_all.sh 20261116 v0_2_4 dist`：重算錯誤 0、Checks 合計 0、check_html 506 個不符 0、check_html_v13 910 個 data-ref 不符 0、`check_rating_gate --strict` 符合。

## v0.2.3（2026-11-15）：估值頁（目標價、評等）與反向 DCF

**依據：** Andy 2026-11-15 決定 MiniMax 與智譜都給目標價與評等（取代規格 J 節 (b)「本模型未建目標價」與 DCF 標準規格「AI 收支模型不加 DCF 腿」的例外）。樣板：Anthropic v0.2.0 估值頁（`anthropic/builder/a11.py`）、反向 DCF（`a12.py`）；`docs/plan/個股DCF標準規格.md`；個股模型原則第 24、36、56、63 條；版面規格 v1.3 A–I 節。

**做法：** `val_page.py`（參照、新增 SRC S142–S155、Inputs I35–I61，皆附加在尾端）＋`valcore.py`（Val、RevDCF 排版與公式；與 `zhipu/builder/valcore.py` 同一份程式）；`scripts/val_snapshots.py scen|sens` 以 `build_xlsx.py --set` 改 Inputs、LibreOffice 重算，產生 `data/val_scenarios.json`（情境 A／C）與 `data/val_sens.json`（龍捲風），`scripts/build_all.sh` 依序執行。個股模型頁（`build_html_v13.py`）估值章拿掉 `data-novaluation`，照 A–I 節：G＝四個目標價與現價圖在估值章最前、H＝加權目標價相對基準低案／高案兩張主圖在敏感度章最前、新增反向 DCF 章；原「累計需新融資」敏感度與原反向檢驗收進摺疊。

**結果（2026-11-05 收盤 HK$170.2）：** DCF 60.92、Comparable 234.82、**加權目標價 HK$147.87**、市場共識 588.79（Investing.com 23 位，2026-09-30）；上檔空間 −13.1%，**估值評等「中立」**；收支關卡 FY2026–FY2030 現金收支合計 −5,983.2 百萬美元（不通過），投資建議＝中立（第 63 條只擋買進以上）；距最近邊界 11.9pt。
- DCF：現金流沿用模型既有路徑（營收 R09、EBIT＝PnL P16 營業虧損含股份支付、資本支出 I22；2H26＝FY2026 − 1H26 實績）。基準 2H26–FY2030 各年 FCFF 皆負、延伸期維持 FY2030 利益率（−97%），三情境 EV 皆 ≤0（DCF 失效旗標 3），DCF 觸及淨現金下限（評價日 pro forma 淨現金 2,394.5 百萬美元 ÷ 352.48M 股 × (1＋ke)^1.35）。
- Comparable：同業 EV ÷ 未來 12 個月共識營收；同業＝智譜 9.74 倍（2026-11-05 同日市值）；MiniMax 自身 5.65 倍（同業的 0.58 倍）。三情境 NTM 營收 1,216／805／558 → Comparable 314／226／173 港元。
- WACC 10.44%（ke 11.10%）；反向 DCF：支持現價需 EV 4,295 百萬美元（模型 −17,832）；只改終值期利益率需 111.6%；延伸期利益率加碼需 +185pt（2033 年約 87%）；同業倍數法下現價只隱含同業的 0.68 倍（所需 NTM 營收 546 vs 模型 805）。
- 敏感度（相對基準）：DCF 權重 25%／75% → +29%／−29%；延伸期利益率加碼 +110pt → +74%；依 FY2030 利益率調整倍數 → −39%；API 單價持平 → +35%、每年 −50% → −5%；token 量高／低端 → +15%／−9%；WACC、g、稅率、營運資金、訓練 MW 為 0（DCF 在淨現金下限，這些輸入不影響目標價）。

**判斷類參數（依樣板慣例預設，列為待 Andy 確認）：**
| 參數 | 預設 | 理由 | 替代 |
|---|---|---|---|
| rf | 4.86%（美債 10 年 CMT，2026-11-06） | 港元聯繫匯率，沿用 Anthropic 同一來源 | 港府 10 年債 |
| β | 1.248（Anthropic 同業 5 年 β 中位數 1.37 經 Blume） | MiniMax 2026-01 上市，自身 β 期間不足 | 1.0／1.6（Sens） |
| ERP | 5%，未加中國國家風險溢酬 | 個股DCF標準規格預設 | 6%（Sens） |
| g、RONIC | 3%、RONIC＝WACC | 同規格 | 2%／4% |
| 延伸期 | 營收年增收斂到 10%、利益率維持 FY2030 | Anthropic 樣板 ext_g33、ext_m_add＝0 | 利益率加碼（Sens +110pt） |
| 同業 | 只有智譜（2026-11-05 同日市值） | 其餘中國與海外上市 AI 模型／應用公司（商湯 0020.HK、科大訊飛、第四範式、SoundHound、C3.ai、Palantir 等）本輪查不到 11-05 同日可查證市值（網路搜尋與行情頁快取皆過期），未納入、不編造 | 補資料後加入同業 |
| 倍數口徑 | EV ÷ NTM 共識營收（FY2026 剩餘 56/365 權重）；不做相對倍數調整 | 同業同為虧損 AI 模型公司 | 依 FY2030 利益率調整（Sens） |
| 權重 | DCF／Comparable 50／50；情境 25／50／25 | 個股模型原則、DCF 標準規格 | 25／75（Sens） |
| 評價日與股數 | 2026-06-30 pro forma（計入 07 配售與可轉債、扣銀行借款與可轉債本金）；股數 352.48M（2026-09-30 FF301） | 中期報告為最近資產負債表 | — |
| 情境 A／C | API token 量年增高端／低端（I09） | 與 Sensitivity 頁同兩案；Anthropic 樣板以需求量定情境 | — |

**資料缺口：** 選擇權與受限制股份稀釋未計（模型未登錄）；租賃負債未登錄；共識營收為 9-21（FY2026）、10-22（FY2027）彙整，早於現價日。既有 Valuation 頁 V02 仍用 S051（349.2M 股），不改。

**驗證：** `scripts/build_all.sh 20261115 v0_2_3 dist`：重算錯誤 0、Checks 合計 0；check_html（506）、check_html_v13（910 個 data-ref）、ci_check、`check_rating_gate --strict` 符合；`check_layout_v13 --xlsx` 全過；`tools/xlsx_diff.py` 對 v0.2.2：既有格只差 README!A1 版本字串，其餘皆為新增格（SRC、Inputs 尾端列、README 尾端三列、Val、RevDCF）。

**獨立核對文字修正（同版 v0.2.3，2026-11-15；只改 `build_html_v13.py` 文字並重產個股模型頁，xlsx 與一頁摘要逐位元組不變）：**估值章加「與智譜互為唯一同業」說明（兩家 Comparable 結論本質上是同一組相對價差）；註明 DCF 淨現金下限以評價日淨現金推到目標時點、未扣其後燒錢（若改用模型 2027-11 內插淨現金約 1,298 百萬美元 → 加權目標價約 131.7 港元，仍中立但距放空邊界約 2.4pt）；β 借用來源（Anthropic 估值頁美國軟體同業 Palantir、CrowdStrike、Cloudflare、Snowflake、Datadog 5 年 β 中位數，Blume 調整）寫進 HTML；原反向檢驗摺疊註明股數基準 349.2M（S051，2026-07-14 配售後），與正文評價股數 352.5M 不同。檢查：check_html_v13（910 個 data-ref 不符 0）、check_html、ci_check、check_rating_gate --strict、check_layout_v13 --xlsx、segment_check、baseline_coordination/check.py 全過。

## v0.2.2（2026-11-15）：個股 HTML 版面 v1.3＋三層協調（讀取層級、C0 公司實績）

**依據：** `docs/workorders/20261114_第二批13家_v1.3與三層協調.md`（Andy 11-14 16:53、16:55、17:40、21:29）；`docs/plan/三層基準協調規則.md` v1.4.2 C0 第 2 點；`docs/plan/個股HTML版面規格.md` v1.3；`tools/coord/README.md`。樣板：Anthropic v0.2.2、OpenAI v0.6.5（PR #524）。

**做法：**
- `coord_pl.py`（由 `build_xlsx.py` 呼叫）新增 PL、Coord 兩頁，只列示、不被既有計算頁引用。
- 個股模型頁 `build_html_v13.py`：章節 命題 → 估值（不適用）→ 敏感度（對至 FY2030 累計需新融資）→ 損益（標準損益表）→ 驅動因子（分部營收推導：企業端＝token × 單價、消費端成長）→ 算力與每 MW 收支 → 資金 → 回本比值 → 對帳（C0 表）→ 資料狀態；除估值、對帳、資料狀態外每章有圖。數字全取自 xlsx（`scripts/check_html_v13.py` 逐一核對 data-ref）。
- **估值：本模型沒有目標價與評等**（Valuation 頁只有市值、EV 與反向檢驗）。MiniMax 已上市、有市價，但 G 節四個目標價、H 節加權目標價敏感度、反向 DCF 章無從呈現；估值章標「估值：不適用（本模型未建目標價）」並帶 `data-novaluation="1"`（規格 J 節 (b)：已上市但本模型未建目標價，2026-11-15 擴充），橫幅寫「本模型未建目標價與評等」；估值章列市值、EV 與反向檢驗，附一張反向檢驗圖。若要給目標價需另建 DCF／同業估值頁（判斷類，待 Andy 決定）。

**同版修正（2026-11-15，只改標籤與旗標，數字不變）：** 估值章拿掉 `data-unlisted`、只留 `data-novaluation`；橫幅改規格字樣「本模型未建目標價與評等」；損益表 FY2026 欄標「1H 實績＋2H 預測」（個股頁表頭與 PL 頁第 5 列）；Valuation 頁 V07 標籤改「智譜 市值 ÷ 1H26 年化營收（市值為 2026-10-02，未更新）」。同版 v0.2.2 重建：`tools/xlsx_diff.py` 對修正前只差上述 3 格文字（Valuation!B14、PL!A1、PL!F5），數值全同；check_html（506）、check_html_v13（734）、ci_check、`check_layout_v13 --xlsx` 全過。

**實績與出處（C0）：**
| 項目 | 期間 | 公司實績 | 出處 | 層級隱含 | 判定 |
|---|---|---|---|---|---|
| 營收（IFRS 總營收） | 2026-01～06（主）；2026 Q2（第二，Derived） | 116.573 $M；Q2 約 75.2 $M | 2026 中期業績公告（2026-08-26，港交所 2026082600680）；Q2 由 1H 與季增 81.8%（業績會轉述，InfoQ）推得 | 無（Wattonomics 無 MiniMax 玩家） | 無可比實績（層級無此玩家） |
| 資本支出（物業廠房設備＋預付款增加） | 1H2026 | 322.3 $M | FY2025 年度業績公告、2026 中期業績公告資產負債表 | 無 | 無可比實績（層級無此玩家） |
| 在役／簽約算力 MW | — | 公司未揭露 | — | 無 | 無可比實績（公司未揭露、層級無此玩家） |
| 算力承諾（阿里雲 2026 年度上限） | 2026 | 300 $M（人民幣 21 億） | 關連交易上限修訂（媒體轉述） | 無 | 無可比實績（層級無此玩家） |

**判定結果：** 可判定 0、對帳中 0、無可比實績 4；無同源列（Wattonomics 快照全文無 MiniMax，層級不讀本模型）。讀取層級：Tokenomics 快照 v19.0（現行 v22.0，本輪不換）。

**驗證：** 建置 → LibreOffice 重算（錯誤 0、Checks 合計 0）→ 敏感度 → HTML（收支頁 506 個數字、個股頁 734 個 data-ref 全對 xlsx）；`tools/xlsx_diff.py` 對 v0.2.1：只差版本字串與新增格；`check_layout_v13 --xlsx`（main 版，J 節 (b) `data-novaluation`）全過；`tools/baseline_coordination/check.py` 通過。`check_company_html` 不適用（收支模型無 html.json，同 anthropic／openai）。
