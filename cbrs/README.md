# cbrs/ — Cerebras Systems（CBRS）收支與估值模型 v0.1

AI 半導體 Project 運算層的新個股（Andy 2026-11-08 要求新增，格式比照 `nvidia/` v0.3.4）。獨立的 Python→Excel 產生器（openpyxl，LibreOffice 重算），HTML 展示層用共用產生器 `tools/company_html/`。

## v0.1（2026-11-08）：新建

| 項目 | 數值 |
|---|---|
| 現價（2026-11-05 收盤） | 81.71 |
| DCF 目標價（情境加權） | 29.46 |
| Comparable 目標價（同業前瞻 P/S） | 126.15（上檔 +54.4%） |
| 加權目標價（DCF 50%） | 77.80（上檔 −4.8%） |
| 投資建議 | 中立 |

- 成品：`dist/20261108_Cerebras個股模型_v0.1.xlsx`、`.html`。驗證：錯誤格 0、重建後逐格差異 0、DCF 檢查列 0、check_company_html、check_in_layout、check_segments 通過。

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
- 估值：標準版 DCF（`tools/fcff_dcf.py` v2；自算產業 β 1.884 重新加槓桿；評價日 2026-12-31；stub 期 FCFF＝2026 下半年，Q 頁 H33）與 Comparable（同業前瞻 P/S 中位數 × FY2028 營收 ÷ 稀釋股數；同業 NVDA、AMD、AVGO、MRVL、CRWV、NBIS）各半；情境權重 0.3／0.5／0.2。上市未滿一年，相對倍數調整係數＝1。
- Tokenomics：`data/tokenomics_snapshot_v22.0.json`（IF_NonNV_Out_CS、IF_NonNV_Hold_CS、IF_HoldEcon、IF_CapexIT 等），只作 Recon 與 capex_mw 高值。

## 已知衝突（Recon C 節）

OpenAI 合約隱含每 MW 年營收 7.62，低於 Tokenomics 隱含的 Cerebras 每 GW 年持有成本 12.23（VR200 12.23 × 持有成本比 1.0）。公司雲端毛利率實績 40%–50% 與「單價低於持有成本」矛盾，較可能是 Tokenomics 的持有成本比（Assumed）偏高，或 750 MW 為設施電力口徑。待整理證據送算力收支。

## 建置

```
python3 cbrs/build_cbrs.py raw.xlsx --html        # --html 另寫 cbrs/html.json
soffice --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir out raw.xlsx
cp out/raw.xlsx cbrs/dist/<日期>_Cerebras個股模型_v<版本>.xlsx
python3 tools/company_html/build_company_html.py cbrs
```
