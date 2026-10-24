# ibiden/ — Ibiden（4062）收支與估值模型 v0.1

Ibiden 的 AI 伺服器載板營收＝AI 加速器顆數（`markets/network`）× Ibiden 份額 × 每顆載板金額（面積、層數逐年增加）× 短缺溢價；電子事業的非 AI 部分、陶瓷與其他事業另依成長率推。股價約 54–65 倍前瞻 EPS：2028 年產能擴到 2.5 倍、競爭者同時擴產之後，短缺溢價能否守住電子事業約 29% 的營益率？

**核心問題：短缺溢價回落多少、何時回落，FY2028 EPS 會少多少；現價隱含的溢價路徑合不合理。**

AI 半導體 Project **載板與 PCB 次產業的樣板**。**Ibiden 在 AI 加速器載板的有效金額（每顆 NVIDIA 加速器的 Ibiden 營收、每顆載板金額）**在 Units 頁建立，是本專案的權威資料（跨專案規則第 13 條）。會計年度 4 月至隔年 3 月，FY(t)＝t/4–(t+1)/3，對照 CY(t) 顆數（重疊 9 個月，不平移）。金額為十億日圓；EPS 為日圓（2026-10-01 一拆二後）。

| 檔案 | 內容 |
|---|---|
| `build_ibiden.py` | 產生器（參數與公式；唯讀取用上游成品） |
| `dist/20261024_Ibiden模型_v0.1.xlsx` | 成品（LibreOffice 重算，錯誤 0 格；CI 逐格比對） |
| `html.json` | HTML 展示層設定（七章；格式見 `tools/company_html/README.md`） |
| `dist/20261024_Ibiden模型_v0.1.html` | HTML 展示層（單頁自含；數字全由 xlsx 重算值抽出；CI 檢核） |

建置：`python3 ibiden/build_ibiden.py raw.xlsx`，再 `soffice --headless --calc --convert-to xlsx --outdir dist raw.xlsx`。工作表：README、In、Q、Units、PL、Val、Recon、Sens、Sources。唯讀取用 `markets/network/dist`（加速器顆數、AI IT 資本支出）、`tsmc/dist`（情境顆數係數）、`tools/cycle_position.py`（規則函式）；**不修改它們**。

## 驅動因子

| 驅動 | 來源 | 推導量 |
|---|---|---|
| 加速器顆數：NVIDIA ＋ 0.4 × 非 NVIDIA（Assumed 權重） | `markets/network` In 頁；`tsmc` In 頁 um28／um29 | 有效顆數 |
| FY2026 AI 伺服器載板＝電子事業 × 65%（Assumed；公司未揭露） | In 頁 B 節 | 每顆有效金額（校準） |
| 每顆載板金額 +18%／年（Analogy：Ibiden 路線圖面積與層數） | In 頁 B 節 | 金額指數 |
| 份額相對值 0.95／0.88／0.85（Assumed） | In 頁 B 節 | AI 載板營收 |
| 短缺溢價指數 1.00／0.95／0.90（Assumed） | In 頁 C 節 | 營收與 AI 載板營益率（成本不變） |
| 非 AI 部分營益率 15%、陶瓷與其他事業營益率沿用 FY2026 預測、淨利 ÷ 營業利益 0.661 | In 頁 C、D 節 | EPS |

- **FY2026 錨點**＝公司 2026-08-04 上修後預測（營收 5,500 億、營業利益 1,270 億、淨利 840 億、EPS 149.68）；Q2 決算後更新。
- 十年 P/E 規則值待 Bloomberg（本版不計）；估值以現價 ÷ 模型 EPS 與 what must be true 呈現。
- 公司目標、前瞻 P/E 隱含 EPS 只在 Recon，不作校準。
