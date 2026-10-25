# avgo/ — Broadcom（AVGO）收支與估值模型 v0.1.1

Broadcom 的 AI 營收＝各 XPU 平台的顆數 × Broadcom 份額 × 每顆售價（核心＋代購 HBM）＋網通附掛。2027–28 年的增量集中在 Anthropic、OpenAI 兩家靠融資部署的客戶。股價約 17 倍 FY2027 EPS：這個折價是在定價客戶信用風險，還是低估了營收？

**核心問題：以市場層顆數往前推，Broadcom 的 AI 營收與 EPS 能否支撐公司展望（FY2027 約 115B、FY2028 約 230B）；差距要靠多少額外顆數補上。**

AI 半導體 Project **運算層的第二個樣板**（ASIC＋網通），複製 `nvidia/` 的結構，匯入 `nvidia/accel_share.py` 的份額函式（本模型是 ASIC 側）。**ASIC 每顆規格**（Broadcom 在各平台的份額、核心售價、HBM 代購比例、每顆 Broadcom 營收與總價）在 Units 頁建立，是本專案的權威資料（跨專案規則第 13 條）。

會計年度 FY(t) 於 t 年 10 月底／11 月初結束（FY2026：2025-11-03～2026-11-01），10/12 落在日曆年 t，所以 **FY(t) 直接對照 CY(t)，不做平移**（NVIDIA 需平移一年）。金額以十億美元為主；口徑為公司 non-GAAP（排除股票薪酬、攤銷、重組）。

| 檔案 | 內容 |
|---|---|
| `build_avgo.py` | 產生器（參數與公式；唯讀取用上游成品） |
| `dist/20261025_Broadcom模型_v0.1.1.xlsx` | 成品（LibreOffice 重算，錯誤 0 格；CI 逐格比對） |
| `html.json` | HTML 展示層設定（七章；格式見 `tools/company_html/README.md`） |
| `dist/20261025_Broadcom模型_v0.1.1.html` | HTML 展示層（單頁自含；數字全由 xlsx 重算值抽出；CI 檢核） |

建置：`python3 avgo/build_avgo.py raw.xlsx`，再 `soffice --headless --calc --convert-to xlsx --outdir dist raw.xlsx`。工作表：README、In、Q、Units、PL、Val、Recon、Sens、Sources。唯讀取用 `markets/memory/market.py`、`markets/network/dist`、`skhynix/dist`、`tsmc/dist`、`markets/equipment/dist`、`nvidia/dist`；**不修改它們**。

**依賴：** 本模型讀 `nvidia/accel_share.py` 與 `nvidia/dist`（PR #92）。PR #92 須先合併。

## 驅動因子（全部唯讀上游）

| 驅動 | 來源 | 推導量 |
|---|---|---|
| Google TPU、「其他 ASIC」顆數 × 情境係數 | `markets/memory` UNITS、SCEN_UNITS；與台積電模型同值（建置時檢查） | 市場顆數 |
| Broadcom 份額（TPU 2026 90%、2027 80%、2028 75%；其他 ASIC 35%／55%／55%） | In 頁 B 節（Assumed） | Broadcom 顆數 |
| 核心售價（TPU 8K、其他 6K，不含 HBM）＋ 代購比例 0.6 × 每顆 HBM 金額 | In 頁 B 節（Assumed）；HBM 容量取 UNITS、單價取 `skhynix` 成品情境別路徑 | 每顆 Broadcom 營收 |
| k_x（XPU 校準係數，Derived） | FY2026 錨點 XPU（AI 57.6 − 網通 17.6）÷ 未校準由下而上 | XPU 營收水準 |
| 網通歸屬 × 情境係數 × k_n | `markets/network` 的「Broadcom 網通（本模型歸屬）」；k_n＝rep_avgo26 ÷ 歸屬 | AI 網通營收 |
| 非 AI 半導體、軟體 | FY2026 錨點 × Assumed 成長 | 其他營收 |
| 各段毛利率（網通 72%、非 AI 65%、軟體 93%、代購 HBM 10%）；XPU 核心毛利率 m_core 為 FY2026 殘差 | In 頁 D、E 節；Units 頁 | 毛利率 |
| 營業費用彈性 0.10、利息、稅率 16%、股數 | Q 頁實績＋ In 頁 G 節 | EPS |

- **FY2026 錨點**＝Q1–Q3 實績（8-K）＋Q4 指引（營收 34.8、AI 21.7、營業利益率 66%、毛利率約 73%）：營收 105.9（共識 105.98）、EPS 11.62（共識 11.66）。校準未使用共識。
- 公司 FY2027、FY2028 AI 營收展望（115、230）與共識只放在 Recon，不作校準。

## 結果（US$B；EPS 為美元；non-GAAP）

| 情境 | 項目 | FY2025A | FY2026（錨點） | FY2027 | FY2028 | FY2029 |
|---|---|---|---|---|---|---|
| A | 營收 | 63.9 | 105.9 | 173.5 | 207.8 | 233.0 |
| A | AI 半導體 | 20.3 | 57.6 | 120.9 | 151.9 | 173.5 |
| A | 毛利率 | 78.6% | 75.1% | 67.8% | 64.7% | 65.0% |
| A | EPS | 6.82 | 11.62 | 18.08 | 21.03 | 24.01 |
| B | EPS | 6.82 | 11.62 | 18.08 | 17.84 | 21.20 |
| C | EPS | 6.82 | 11.62 | 18.08 | 14.82 | 16.83 |

## 對帳（Recon，重點）

| 項目 | 對照值 | 模型 A | 差距 |
|---|---|---|---|
| FY2027 AI 營收 vs 公司展望 | 115 | 120.9 | +5.2% |
| FY2027 AI 營收 vs JPMorgan | 150 | 120.9 | −19.4% |
| FY2028 AI 營收 vs 公司展望 | 230 | 151.9 | −34.0% |
| FY2028 要對上公司展望所需的 Broadcom XPU 顆數倍數 | — | 1.73 倍（A）；B 2.59、C 3.85 | — |
| FY2027 營收 vs 共識 | 173.97 | 173.5 | −0.3% |
| FY2027 EPS vs 共識 | 19.40 | 18.08 | −6.8% |
| FY2027 毛利率 vs 共識隱含 | 72.4% | 67.8% | −4.6 pt |
| 2027 年 ASIC 金額份額（模組層級） | 先驗比對 14–19% | 22.4% | 口徑：本模型 NVIDIA 只計模組價 |
| 2027 年 ASIC 顆數份額 vs JPMorgan | 53.6% | 52.7% | 非獨立（市場層即取自 JPMorgan） |
| 2027 年 Broadcom ÷ ASIC 金額 vs Counterpoint | 60% | 61.5% | — |

- **FY2027 對得上、FY2028 差三成四。** 市場層顆數推得 FY2027 AI 121B，與公司 115B、共識隱含約 121B 一致；FY2028 市場層只給 152B，公司說 230B（含 Anthropic 2028 年再加約 10 GW、OpenAI 5 GW 以上）。差距要靠 Broadcom XPU 顆數多 73%（約 13.8M 顆 vs 8.0M）。
- **毛利率比共識低約 4.6 個百分點**：主要來自 HBM 代購比例（Assumed 0.6）。代購 30% 時 FY2027 毛利率 70.3%，EPS 幾乎不變（Sens）——代購只改組合，不太改利潤。
- **客戶融資**（Recon F）：Anthropic 放款上限 42B＋OpenAI 融資安排 50B ≈ 模型 FY2027 AI 營收的 76%、淨利的 1.0 倍；不進基準 EPS。

## 估值（Val）

- **目標價（12 個月，情境加權 A 30%／B 50%／C 20%）：US$647.37，上檔空間 +100.3%**（情境 A／B／C 隱含價 US$647.37／US$647.37／US$647.37；Val 頁「0. 目標價與上檔空間」，放在 Val 既有內容之後；權重為 In 頁 tp_wA／tp_wB／tp_wC，Assumed，Andy 可改）。方法：P/E：十年（2016–2025，10 月底會計年度年底）GAAP P/E 中位數（Macrotrends，第 5 節；2016 GAAP EPS 為負、2025 未找到，取有值的 8 年）× 各情境 FY2027 non-GAAP EPS（FY2027＝2026-11～2027-10）。本模型無規則值（前瞻 P/E 歷史待補），依工作單改用此備案。偏誤方向為高估：GAAP EPS 受併購無形資產攤銷壓低，GAAP P/E 高於 non-GAAP 口徑（2024 年併購 VMware 後約 137x），乘在 non-GAAP EPS 上偏高；FY2027 三情境 EPS 相同（情境差從 FY2028 起），故三情境隱含價相同；隱含價約為 FY2027 年底（2027-10）的價格、未折現。

現價 $323.24（2026-10-22）÷ FY2027 EPS＝**17.9x**（模型）、16.7x（共識）；÷ FY2028 EPS：A 15.4x、B 18.1x、C 21.8x。近 12 月 non-GAAP P/E 33.1x、GAAP 41.3x。以 Andy 先驗的合理 25x 計，FY2027 EPS 隱含 $452（+40%）。

**十年 P/E 規則值：本版不適用。** GAAP 落後 P/E 受無形資產攤銷扭曲（2016 年 EPS 為負、2018 年 6.5x、2024 年 137x）；前瞻 P/E 十年歷史待 Bloomberg 抓取單（10-27）。循環位置（共用規則，不平移）只作紀錄：FY2027 A、B 中段、C 見頂。

## 敏感度（Sens，情境 A，FY2028 EPS）

核心售價 ×0.8／×1.2：−8.3%／+8.3%；TPU 份額 −15pt：−6.4%；網通 ×0.8：−5.2%；XPU 顆數 ×0.9：−4.5%；營業費用率 +2pt：−3.4%；HBM 代購比例 0.3／1.0：−1.0%／+1.1%（毛利率 68.4%／61.3%）；情境 B／C：−15.2%／−29.5%。

## 最大限制

- k_x 是顆數、份額、核心單價的合併係數，無法分開識別；FY2026 k_x＝1.108。
- Broadcom 各平台份額、核心售價、HBM 代購比例皆為 Assumed；「其他 ASIC」含昇騰、Maia 等非 Broadcom 晶片，份額為粗估。
- XPU 占 AI 營收（FY2026 錨點 69%）依 `markets/network` 的網通實際推估 17.6；Q3 的 73% 只有一個二手來源。
- 客戶融資曝險不進基準 EPS；十年倍數待補；R4 產業部署路徑未建（Recon 預留列）。

## v0.1.1（2026-10-25）：補目標價

Val 頁新增「0. 目標價與上檔空間」（列標籤依工作單；為不位移既有儲存格，放在 Val 既有內容之後），In 頁最後新增目標價情境權重 tp_wA／tp_wB／tp_wC（0.3／0.5／0.2，Assumed）；HTML 估值章第一個區塊為本節。除新增列、README 頁標題與引用上游成品的檔名文字外，原有儲存格的值與 v0.1 相同（`tools/xlsx_diff.py` 比對）。

- **目標價（12 個月，情境加權 A 30%／B 50%／C 20%）：US$647.37，上檔空間 +100.3%**（情境 A／B／C 隱含價 US$647.37／US$647.37／US$647.37；Val 頁「0. 目標價與上檔空間」，放在 Val 既有內容之後；權重為 In 頁 tp_wA／tp_wB／tp_wC，Assumed，Andy 可改）。方法：P/E：十年（2016–2025，10 月底會計年度年底）GAAP P/E 中位數（Macrotrends，第 5 節；2016 GAAP EPS 為負、2025 未找到，取有值的 8 年）× 各情境 FY2027 non-GAAP EPS（FY2027＝2026-11～2027-10）。本模型無規則值（前瞻 P/E 歷史待補），依工作單改用此備案。偏誤方向為高估：GAAP EPS 受併購無形資產攤銷壓低，GAAP P/E 高於 non-GAAP 口徑（2024 年併購 VMware 後約 137x），乘在 non-GAAP EPS 上偏高；FY2027 三情境 EPS 相同（情境差從 FY2028 起），故三情境隱含價相同；隱含價約為 FY2027 年底（2027-10）的價格、未折現。
