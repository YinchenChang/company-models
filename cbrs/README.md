# cbrs/ — Cerebras Systems（CBRS）收支與估值模型 v0.1.5

AI 半導體 Project 運算層的新個股（Andy 2026-11-08 要求新增，格式比照 `nvidia/` v0.3.4）。獨立的 Python→Excel 產生器（openpyxl，LibreOffice 重算），HTML 展示層用共用產生器 `tools/company_html/`。

## v0.1.5（2026-11-21）：現價與同業報價改 2026-11-20 收盤（工作單 px1120）

依據：`docs/workorders/20261121_px1120_現價與同業報價換11-20收盤.md`；價格只取 `markets/invest_map/data/closes_20261109_20261120.csv`。base＝v0.1.4。

- **現價與同業報價改 2026-11-20 收盤（工作單 px1120）：** 現價 108.72 → 82.67（In 頁 px，Yahoo Finance 日線收盤，Verified）；Comparable 同業前瞻 P/S（＝共識前瞻 P/S × 11-20 收盤 ÷ 頁面價格）：NVIDIA 7.30→5.63、AMD 10.42→8.28、Broadcom 10.53→9.24、Marvell 8.77→6.84；**同業中位數 9.59 → 7.56**。**加權目標價 97.21 → 80.10；上檔空間 −10.6% → −3.1%；評等 中立 → 中立。**未換價同業：無。
- **不變：** 共識、rf、β、ERP、股數、情境權重與營運假設；以市值計的 WACC 權重隨現價自動重算。本模型沒有 Sens／Coord 另一側讀數快照。
- **檢查：** 錯誤格 0；`check_company_html` 通過；`check_layout_v13` 通過（警告只剩既有參數代號）；`baseline_coordination/check.py` 0 項不符；`check_sens_coverage` 未通過（27／36 鍵無案例或豁免，v0.1.2 起既有，本版未改）。
- 新增 `px1120_post.py`（只改版號與說明文字）；建置同 v0.1.4。

## v0.1.4（2026-11-16）：現價與同業報價改 2026-11-13 收盤（工作單 px1113）

依據：`docs/workorders/20261116_px1113_現價與同業報價換11-13收盤.md`；價格只取 `markets/invest_map/data/closes_20261020_20261113.csv`。base＝v0.1.3。

- **現價與同業報價改 2026-11-13 收盤（工作單 px1113）：** 現價 81.71 → 108.72（In 頁 px；來源改 Yahoo Finance 日線收盤，標記由 Derived 改 Verified，保留 v0.1.3 的 equibles 推算 81.71，兩者 11-05 收盤相同）；Comparable 同業前瞻 P/S（＝共識前瞻 P/S × 11-13 收盤 ÷ 頁面價格，共識與頁面價格不變）：NVIDIA 8.14→7.30、AMD 7.89→10.42、Broadcom 8.84→10.53、Marvell 7.47→8.77；**同業中位數 8.02 → 9.59**。**加權目標價 83.94 → 97.21；上檔空間 +2.7% → −10.6%；評等 中立 → 中立。**未換價同業：無。
- **不變：** 共識（stockanalysis 10-05）、rf、β、ERP、股數（2026-11-05 封面）、情境權重與營運假設；以市值計的 WACC 權重與 D/E 隨現價自動重算。本模型沒有「資料基準日」型輸入，也沒有 Sens／Coord 另一側讀數快照。
- **文字與沿革：** 「2026-11-05 收盤」標籤與來源欄改 11-13；命題 4 文字改「11-13 收盤 108.72 美元」；README 頁本版要點列改寫為 v0.1.4，前一版（v0.1.3）本版要點移到 README 頁末；既有沿革列保留原文。新增 `px1113_post.py`（只改版號與說明文字）。
- 建置：同 v0.1.3（`python3 cbrs/build_cbrs.py raw.xlsx --html` → LibreOffice 重算 → `dist/` → `gen_html_json.py <xlsx>` → `build_company_html.py cbrs`）。

## v0.1.3（2026-11-15）：滾入 3Q26（2026-11-12 財報）

**依據：** Cerebras 8-K 2026-11-12（3Q26 財報新聞稿，https://www.sec.gov/Archives/edgar/data/0002021728/000162828026073366/cbrsannouncesfinancialresu.htm）、10-Q（季度結束 2026-09-30，https://www.sec.gov/Archives/edgar/data/0002021728/000162828026073494/cbrs-20260930.htm）。只換實績、資產負債表日與指引；情境假設（計費 MW、每 MW 年營收、毛利率、費用彈性、股票薪酬率等）與公式架構不變。

**改了哪些輸入（Q 頁、In 頁 C 節與 DCF 橋接）：**

| 項目 | v0.1.2 | v0.1.3 | 出處／等級 |
|---|---|---|---|
| Q 頁 3Q26 欄 | 指引中點（核心營收 0.215） | 實績：GAAP 營收 0.2108（硬體 0.0689、雲端 0.1419）、GAAP 毛利 0.0421、GAAP 營業損益 −0.2241、淨利 −0.1612；核心營收 0.2309（硬體 0.0860、雲端 0.1449）、核心毛利 0.0946（41.0%）、核心營業費用 0.1277、核心營業損益 −0.0331；股票薪酬 0.1551、折舊攤銷 0.0298、其他收益 0.0639 | 8-K 2026-11-12［Verified］ |
| 4Q26 欄 | FY 指引中點 − 前三季 | 4Q26 指引中點：核心營收 0.32（0.31–0.33）、毛利率 48%（47%–49%）、營業利益率 0%（(1%)–1%）；硬體 0.085、折舊 0.035、資本支出 0.5（0.1／1.1）、其他收益 0.06 為 Assumed（沿用 v0.1.2 下半年每季值） | 8-K 2026-11-12［Interested-party］／Assumed |
| FY2026 指引（只作 Recon） | 8.80–8.90 億、41%–43%、(19%)–(17%) | 9.40–9.60 億、42%–44%、(10%)–(8%) | 8-K 2026-11-12［Interested-party］ |
| 資產負債表日 | 2026-06-30 | 2026-09-30：現金 3.0748、短期投資 4.2860、受限現金 0.4440、OpenAI 貸款 0.8291、租賃 0.6211、股東權益 9.1589 | 8-K 2026-11-12［Verified］ |
| 現金流（累計） | 1H26：營業 −0.0475、資本支出 0.5489 | 9M26：營業 −0.3889、資本支出 1.0461（3Q26 資本支出 0.4972＝差額，Derived） | 8-K 2026-11-12［Verified］ |
| 評價股數 | 3.2361 億（2.3756＋0.7601＋0.1003） | 3.1533 億（已發行 2.4300（2026-11-05）＋潛在稀釋 0.7233（2026-09-30）；7 月 N 類認股權證已含在內，不重複） | 10-Q［Verified／Derived］ |
| Stub 期 | 2026-06-30 → 12-31（下半年 FCFF −1.263） | 2026-09-30 → 12-31（4Q26 FCFF −0.625） | Derived |
| RPO、認列時點 | 254 億；24 個月內 22%、25–48 個月 43% | 261 億；33%、44% | 8-K／10-Q［Verified］ |
| 公司容量與 2027 說法 | 逾 600 MW 上線或簽約；財務長逾三倍 | 700 gross MW 已簽約、2027 年底前交付；執行長「2027 年營收三倍」 | 8-K［Interested-party］ |

FY2026 錨點（核心）：營收 0.952、毛利率 44.4%、營業利益率 (7.4%)、資本支出 1.546。Recon RPO 對照窗口改為 4Q26–3Q28／4Q28–3Q30（跟著 10-Q 起算日）。Coord D 節 C0 主讀數改為 9M26（雲端核心營收 0.3524、GAAP 營收 0.5843、資本支出 1.0461；第二讀數 3Q26 × 4），判定仍為「無可比實績」，對帳中 0 項。

**目標價與評等：**

| 項目 | v0.1.2 | v0.1.3 |
|---|---|---|
| DCF 情境價值（評價日；A／B／C） | 60.27／15.13／0 | 65.06／17.96／0 |
| DCF 目標價（情境加權） | 29.46 | 32.74 |
| Comparable 目標價 | 131.67 | 135.15 |
| 加權目標價（現價 81.71，11-05 收盤） | 80.56（−1.4%） | 83.94（+2.7%） |
| 投資建議 | 中立 | 中立 |

**未做：** 現價未更新（仍為 11-05 收盤；財報後股價變動未納入）；共識（stockanalysis 10-05）與同業 P/S 未更新；股票薪酬 3Q26 實績 1.55 億／季高於 sbc_27 假設的推導基礎（2Q26 經常性 1.03 億），未改情境假設；公司全年指引與 4Q26 指引不一致（前三季實績＋4Q26 中點合成毛利率 44.4%、營業利益率 (7.4%)，皆優於全年指引上緣），本版以 4Q26 指引為錨點，Recon A 節並列；FY2023 GAAP 損益仍未取得。`check_sens_coverage --require-keycol` 未通過（27／36 鍵無案例或豁免，v0.1.2 起既有，本版只換鍵名 q3_*/h2_* → q4_*）。

## v0.1.2（2026-11-15）：個股 HTML 版面 v1.3＋三層協調（讀取層級、兩讀數判定、C0 實績）

**依據：** `docs/workorders/20261114_第二批13家_v1.3與三層協調.md`（第 2 節）、`docs/workorders/20261114_C0實績對帳補列_七家.md`、`docs/plan/三層基準協調規則.md` v1.4.2（C0 第 2 點、C4）、`docs/plan/個股HTML版面規格.md` v1.3。樣板：AMD v0.4.7（`amd/v047_post.py`）、NVIDIA v0.3.10。

**做法：** `build_cbrs.py` 存檔前呼叫 `v012_post.apply()`（只新增頁與列）：PL 頁末標準損益表（FY2024A–FY2025A GAAP 實績＋FY2026E–FY2032E 情境 B，GAAP 近似口徑）、RevDCF 頁（反向 DCF；Comparable 部分為前瞻 P/S）、Val 頁 0.7（四個目標價）、Sens 頁末「加權目標價相對基準」（營運案例只套用情境 B：基準加權目標價＋情境 B 權重 ×（案例的情境 B 加權價值 − 基準））、Coord 頁、README 頁末 v1.3 白話列。`build_cbrs.py --html` 改寫 `html_src.json`，`gen_html_json.py <xlsx>` 轉成 v1.3 章節順序的 `html.json`。

**兩讀數（C4）：不適用。** 本模型是唯一不讀 Siliconomics 的硬體個股（量＝平均計費 MW，由公司合約與揭露推得）；Siliconomics 對帳中年份（CY2027–CY2029）沒有傳導路徑，另一側讀數（只換產業部署 GW）與現行逐格相同。Coord B 節並列兩欄（另一側欄引用現行欄）並寫明理由；不建 `co_flag`、不跑 `alt_run.py`、橫幅不加另一側讀數格；`coord_meta` 的 `hardware_companies` 不需加入 cbrs。

**讀取層級（Coord A 節）：** Tokenomics v22.0 快照（Cerebras 產出比、持有成本比、VR200 每 GW 持有成本與 IT 資本支出；只作 Recon 對照與每 MW 資本支出敏感度高值）；Siliconomics v0.19、Wattonomics v0.3.2 皆未讀（Wattonomics 的 Cerebras（OpenAI）750 MW 園區列不納入加總、無 Cerebras 玩家）。

**C0（實績）與出處（Coord D 節）：**

| 實績（2026 已公布期間） | 值（十億美元） | 出處 |
|---|---|---|
| 營收合計（GAAP）1Q26＋2Q26 | 0.1934＋0.1801＝0.3735 | Cerebras 10-Q（2026-03-31）、8-K 2026-08-12（Q 頁第 7 列） |
| 核心雲端營收 1Q26＋2Q26 | 0.0798＋0.1277＝0.2075（2Q26 × 4＝0.511） | 8-K（1Q26、2Q26 財報新聞稿；Q 頁第 19 列） |
| 資本支出 1H26 | 0.549 | 10-Q（2026-06-30）現金流量表（Q 頁第 48 列） |

**判定：** 三列皆「無可比實績」——Siliconomics 無晶圓級加速器（Cerebras）顆數或份額；Wattonomics 無 Cerebras 玩家；Tokenomics 只有每 GW 收支，換成營收需在役／計費 MW（公司未揭露）。對帳中項目 0。Recon C 節「合約每 MW 年營收 vs Tokenomics 持有成本」屬合約與層級假設之比較，只列參考。（v0.1.2 時 3Q26 未取得，主讀數只到 1H26；v0.1.3 已滾入 9M26。）

**不變性與驗證：** 既有數值、目標價（加權 80.56、DCF 29.46、Comparable 131.67）、評等（中立）不變；`tools/xlsx_diff.py` 新版 vs v0.1.1 只差 README A1 版號與日期、新增頁（RevDCF、Coord）與新增列。錯誤格 0；`check_layout_v13 --xlsx` 通過（警告只剩既有估值與對帳說明中的參數代號）、`check_company_html cbrs` 通過；`tools/baseline_coordination/check.py` 的不符項皆為其他公司（本機 sparse 工作樹缺其 HTML），與本公司無關。歷史只有兩年（FY2023 GAAP 損益本版未取得原文，待補）。

## v0.1.1（2026-11-08）：Comparable 只用晶片同業

Andy 2026-11-08 21:21：Comparable 同業只用晶片同業（NVIDIA、AMD、Broadcom、Marvell），移除 CoreWeave、Nebius；其他判斷項照預設。同業前瞻 P/S 中位數 7.68 → 8.02。

| 項目 | 數值 |
|---|---|
| 現價（2026-11-05 收盤） | 81.71 |
| DCF 目標價（情境加權） | 29.46 |
| Comparable 目標價（同業前瞻 P/S） | 131.67（上檔 +61.1%） |
| 加權目標價（DCF 50%） | 80.56（上檔 −1.4%） |
| 投資建議 | 中立 |

- 成品：`dist/20261108_Cerebras個股模型_v0.1.1.xlsx`、`.html`。v0.1（2026-11-08 新建）：加權目標價 77.80（同業含 CoreWeave、Nebius）。驗證：錯誤格 0、重建後逐格差異 0、DCF 檢查列 0、check_company_html、check_in_layout、check_segments 通過。

## 模型結構

營收＝推論雲端（平均計費 MW × 每 MW 年營收）＋系統（硬體）銷售。會計年度＝日曆年；金額十億美元；每 MW 年營收與每 MW 資本支出為百萬美元／MW（＝十億美元／GW）。

| 驅動 | 來源 | 推導量 |
|---|---|---|
| 平均計費 MW（三情境 FY2027–FY2032） | Assumed；公司「2027 年底前逾 600 MW 上線或簽約」、OpenAI 750 MW 承諾＋1.25 GW 選擇權、RPO 254 億與認列時點 | 雲端營收、新增 MW |
| 每 MW 年營收 | OpenAI 主合作協議：逾 200 億美元 ÷ 0.75 GW ÷ 3.5 年＝7.62（二手；低／高＝5 年／3 年） | 雲端營收 |
| 雲端現金毛利率（不含折舊） | FY2026 錨點（指引推得約 65%）→ FY2029 情境值（72%／66%／58%） | 雲端毛利＝營收 × 現金毛利率 − 雲端折舊 |
| 每 MW 資本支出 | 9（FY2026 折舊與不動產廠房設備推估）；高＝Tokenomics 隱含 22.6 | 資本支出、折舊（5 年）、FCFF |
| 營業費用 | FY2026 錨點；年增＝彈性 0.6 × 營收年增（5%–60%） | EBIT |
| 股票薪酬 | 營收 14%（FY2027）→ 6%（FY2032） | EBIT（標準版 DCF 視為費用） |

- FY2026 錨點＝1Q26–3Q26 實績＋4Q26 指引中點（Q 頁；v0.1.3 起）。4Q26 財報公布後改為全年實績。
- EBIT 口徑：核心營業損益 − 經常性股票薪酬；不扣客戶認股權證攤銷（認股權證已計入評價股數 3.153 億股）、代墊資料中心收入與 IPO 一次性費用。
- 估值：標準版 DCF（`tools/fcff_dcf.py` v2；自算產業 β 1.884 重新加槓桿；評價日 2026-12-31；stub 期 FCFF＝4Q26，Q 頁 H33）與 Comparable（同業前瞻 P/S 中位數 × FY2028 營收 ÷ 稀釋股數；同業 NVDA、AMD、AVGO、MRVL（只用晶片同業））各半；情境權重 0.3／0.5／0.2。上市未滿一年，相對倍數調整係數＝1。
- Tokenomics：`data/tokenomics_snapshot_v22.0.json`（IF_NonNV_Out_CS、IF_NonNV_Hold_CS、IF_HoldEcon、IF_CapexIT 等），只作 Recon 與 capex_mw 高值。

## 已知衝突（Recon C 節）

OpenAI 合約隱含每 MW 年營收 7.62，低於 Tokenomics 隱含的 Cerebras 每 GW 年持有成本 12.23（VR200 12.23 × 持有成本比 1.0）。公司雲端毛利率實績 40%–50% 與「單價低於持有成本」矛盾，較可能是 Tokenomics 的持有成本比（Assumed）偏高，或 750 MW 為設施電力口徑。待整理證據送算力收支。

## 建置

```
python3 cbrs/build_cbrs.py raw.xlsx --html        # --html 另寫 cbrs/html_src.json（v0.1.2 起）
soffice --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir out raw.xlsx
cp out/raw.xlsx cbrs/dist/<日期>_Cerebras個股模型_v<版本>.xlsx
python3 cbrs/gen_html_json.py cbrs/dist/<日期>_Cerebras個股模型_v<版本>.xlsx   # 版面 v1.3 → cbrs/html.json
python3 tools/company_html/build_company_html.py cbrs
```
