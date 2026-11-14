# cbrs/ — Cerebras Systems（CBRS）收支與估值模型 v0.1.2

AI 半導體 Project 運算層的新個股（Andy 2026-11-08 要求新增，格式比照 `nvidia/` v0.3.4）。獨立的 Python→Excel 產生器（openpyxl，LibreOffice 重算），HTML 展示層用共用產生器 `tools/company_html/`。

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

**判定：** 三列皆「無可比實績」——Siliconomics 無晶圓級加速器（Cerebras）顆數或份額；Wattonomics 無 Cerebras 玩家；Tokenomics 只有每 GW 收支，換成營收需在役／計費 MW（公司未揭露）。對帳中項目 0。Recon C 節「合約每 MW 年營收 vs Tokenomics 持有成本」屬合約與層級假設之比較，只列參考。3Q26（2026-11-12 公布）本版未取得原文，主讀數只到 1H26（6 個月）。

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

- FY2026 錨點＝1Q26、2Q26 實績＋3Q26 指引中點＋全年指引中點推得 4Q26（Q 頁）。3Q26 財報 2026-11-12 公布後更新。
- EBIT 口徑：核心營業損益 − 經常性股票薪酬；不扣客戶認股權證攤銷（認股權證已計入評價股數 3.236 億股）、代墊資料中心收入與 IPO 一次性費用。
- 估值：標準版 DCF（`tools/fcff_dcf.py` v2；自算產業 β 1.884 重新加槓桿；評價日 2026-12-31；stub 期 FCFF＝2026 下半年，Q 頁 H33）與 Comparable（同業前瞻 P/S 中位數 × FY2028 營收 ÷ 稀釋股數；同業 NVDA、AMD、AVGO、MRVL（只用晶片同業））各半；情境權重 0.3／0.5／0.2。上市未滿一年，相對倍數調整係數＝1。
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
