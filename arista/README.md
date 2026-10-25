# arista/ — Arista Networks（ANET）收支與估值模型 v0.1.1

Arista 的 AI 營收＝AI 後端 scale-out 交換器市場（`markets/network`）× 乙太網路占比 × Arista 份額（含前端與 scale-across 附掛）；非 AI（雲端前端、企業、園區）另依成長率推。股價約 37 倍前瞻 EPS：在 NVIDIA（Spectrum-X）與白牌夾擊下，Arista 的份額路徑能否支撐這個倍數？

**核心問題：2027–28 年 Arista 在 AI 乙太網路的份額每掉一截，EPS 少多少；現價隱含的份額路徑合不合理。**

AI 半導體 Project **網通系統端的樣板**（AVGO 為網通晶片端）。**Arista 在 AI 後端乙太網路的份額**在 Units 頁建立，是本專案的權威資料（跨專案規則第 13 條）。會計年度＝日曆年，不平移。金額以十億美元為主；口徑為公司 non-GAAP；十年 P/E 規則值用 GAAP 換算 EPS（non-GAAP × 1H26 GAAP÷non-GAAP 0.926）。

| 檔案 | 內容 |
|---|---|
| `build_arista.py` | 產生器（參數與公式；唯讀取用上游成品） |
| `dist/20261025_Arista模型_v0.1.1.xlsx` | 成品（LibreOffice 重算，錯誤 0 格；CI 逐格比對） |
| `html.json` | HTML 展示層設定（七章；格式見 `tools/company_html/README.md`） |
| `dist/20261025_Arista模型_v0.1.1.html` | HTML 展示層（單頁自含；數字全由 xlsx 重算值抽出；CI 檢核） |

建置：`python3 arista/build_arista.py raw.xlsx`，再 `soffice --headless --calc --convert-to xlsx --outdir dist raw.xlsx`。工作表：README、In、Q、Units、PL、Val、Recon、Sens、Sources。唯讀取用 `markets/network/dist`（scale-out 交換器市場、AI IT 資本支出）、`tsmc/dist`（情境顆數係數）、`markets/equipment/dist`（經 `tools/cycle_position.py`，只用規則函式）；**不修改它們**。不依賴 #93（`avgo/`）。

## 驅動因子

| 驅動 | 來源 | 推導量 |
|---|---|---|
| AI 後端 scale-out 交換器市場 × 情境顆數係數 | `markets/network` Mkt 頁；`tsmc` In 頁 um28／um29 | 市場 |
| 乙太網路占比（2026 75% → 2029 85%） | In 頁 B 節（Assumed；Dell'Oro 2025 超過三分之二） | 乙太網路市場 |
| Arista 有效份額（FY2026 以公司 AI 目標 3.5 校準＝14.6%）× 相對值（0.95／0.90／0.88） | In 頁 B 節（Assumed）；Units 頁 | AI 營收 |
| 非 AI 成長（+10%／+8%／+6%） | In 頁 C 節（Assumed） | 非 AI 營收 |
| AI 毛利率低 6 個百分點；非 AI 毛利率為 FY2026 殘差 | In 頁 D 節 | 毛利率 |
| 營業費用彈性 0.55（FY2026 實測 0.59）、其他收益 0.55、稅率 21%、股數持平 | Q 頁＋In 頁 F 節 | EPS |

- **FY2026 錨點**＝1Q、2Q 實績（8-K）＋3Q 指引（3.3、毛利率 63%、營益率 48.5%）＋全年指引（12.6、63%、48.5%）：EPS 4.09（共識 4.12）；3Q26 EPS 1.071（指引中點 1.07）。
- 共識、單一分析師、研調份額只在 Recon，不作校準。

## 結果（US$B；EPS 為美元；non-GAAP）

| 情境 | 項目 | FY2025A | FY2026（錨點） | FY2027 | FY2028 | FY2029 |
|---|---|---|---|---|---|---|
| A | 營收 | 9.01 | 12.60 | 16.77 | 18.31 | 20.00 |
| A | AI | 1.63 | 3.50 | 6.76 | 7.50 | 8.54 |
| A | 毛利率 | 64.6% | 63.0% | 62.2% | 62.2% | 62.1% |
| A | EPS | 2.98 | 4.09 | 5.47 | 5.99 | 6.56 |
| B | EPS | 2.98 | 4.09 | 5.47 | 5.51 | 6.14 |
| C | EPS | 2.98 | 4.09 | 5.47 | 4.97 | 5.39 |

- **FY2027 比市場高一截：** EPS 5.47 vs Erste 4.79（+14%）、vs 前瞻 P/E 隱含未來 12 月 EPS 4.65（+18%，口徑不同）。差距來自網通模型 2027 年 scale-out 交換器 +90%（Rubin 每顆內容倍增）；公司稱供給吃緊到 2028 年，模型未設供給上限（Sens：市場 ×0.8 → FY2027 EPS −7.8%）。
- **份額是最大變數：** 份額相對值 0.80／0.70 時 FY2028 EPS −8.8%；情境 C（顆數 ×0.6）−17.1%。

## 估值（Val）

- **目標價（12 個月，情境加權 A 30%／B 50%／C 20%）：US$174.08，上檔空間 +1.0%**（情境 A／B／C 隱含價 US$181.31／US$181.31／US$145.13；Val 頁「0. 目標價與上檔空間」，放在 Val 既有內容之後；權重為 In 頁 tp_wA／tp_wB／tp_wC，Assumed，Andy 可改）。方法：P/E：各情境 FY2027 GAAP 換算 EPS（non-GAAP EPS × g2n）× 十年（2016–2025 年底）GAAP P/E 依循環位置取分位（以本模型營收最高年判定：A、B 中段＝中位數，C 見頂＝25 分位；companiesmarketcap 與 Macrotrends 兩來源平均）。十年倍數含 2021 年與 2024–2025 年高成長期的高倍數，偏誤方向為高估；FY2027 三情境 EPS 相同，A、B 隱含價相同、C 只差在倍數；隱含價約為 FY2027 年底的價格、未折現，當作 12 個月目標價時偏高。

現價 $172.37（2026-10-23）÷ FY2027 EPS＝31.5x、÷ FY2028 EPS：A 28.8x、B 31.3x、C 34.7x。十年 GAAP P/E（companiesmarketcap、Macrotrends 兩來源）：25 分位 28.7x、中位數 35.8x、75 分位 43.0x。循環位置以本模型營收最高年判定（A、B 到 FY2029 仍成長＝中段；C 見頂於 FY2027）。規則值隱含價（約 FY2027 年底）：A、B $181（+5%）、C $145（−16%）。以 Andy 先驗合理 25x × FY2027 EPS 計 $137（−21%）。

## 最大限制

- 份額相對值與乙太網路占比皆為 Assumed；FY2026 有效份額含前端與 scale-across，不是純後端份額。
- FY2026 AI 用公司目標「至少 3.5」；各季 AI 營收未揭露。
- 2027 共識未找到（付費）；FY2025 non-GAAP 淨利為原報值（公司 2026 年起改稅額方法，重編全年值未找到）。
- 十年 P/E 含高成長期溢價；前瞻 P/E 歷史待 Bloomberg。R4 產業部署路徑未建。

## v0.1.1（2026-10-25）：補目標價

Val 頁新增「0. 目標價與上檔空間」（列標籤依工作單；為不位移既有儲存格，放在 Val 既有內容之後），In 頁最後新增目標價情境權重 tp_wA／tp_wB／tp_wC（0.3／0.5／0.2，Assumed）；HTML 估值章第一個區塊為本節。除新增列、README 頁標題與引用上游成品的檔名文字外，原有儲存格的值與 v0.1 相同（`tools/xlsx_diff.py` 比對）。

- **目標價（12 個月，情境加權 A 30%／B 50%／C 20%）：US$174.08，上檔空間 +1.0%**（情境 A／B／C 隱含價 US$181.31／US$181.31／US$145.13；Val 頁「0. 目標價與上檔空間」，放在 Val 既有內容之後；權重為 In 頁 tp_wA／tp_wB／tp_wC，Assumed，Andy 可改）。方法：P/E：各情境 FY2027 GAAP 換算 EPS（non-GAAP EPS × g2n）× 十年（2016–2025 年底）GAAP P/E 依循環位置取分位（以本模型營收最高年判定：A、B 中段＝中位數，C 見頂＝25 分位；companiesmarketcap 與 Macrotrends 兩來源平均）。十年倍數含 2021 年與 2024–2025 年高成長期的高倍數，偏誤方向為高估；FY2027 三情境 EPS 相同，A、B 隱含價相同、C 只差在倍數；隱含價約為 FY2027 年底的價格、未折現，當作 12 個月目標價時偏高。
