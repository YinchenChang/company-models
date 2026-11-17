# Tesla（TSLA）個股模型 v0.2（AI實體應用）

- 成品：`physical_ai/tesla/dist/20261116_Tesla個股模型_v0.2.xlsx`（v0.1 在 `dist/_待刪除/`）（LibreOffice 重算後的成品，含公式與快取值；0 錯誤值）
- 建置：`python3 physical_ai/tesla/build_tesla.py [--base <原檔>] [--nosens]`（約 5 分鐘；Sens 為整本重算快照，4 個 LibreOffice 平行）
- 底稿：以 `20261028_TeslaTSLA.US_post_3Q26.xlsx`（post-3Q26 版）為底改版，原檔全部工作表保留；預設路徑 `/mnt/user-data/uploads/10-AI實體應用/`
- 資料截止日 2026-11-16；股價基準日 2026-11-13（收盤 390.22）；財報 3Q26 10-Q（2026-09-30）
- 目標價、評等以成品 Val 頁為準；本檔數字為建置當時快照（build_id 見 README 頁 A1、Sens 頁底部）

## 命題

Tesla 的價值有多少來自「AI 進入實體世界」（FSD、Robotaxi、Optimus），而不是汽車與能源本業？3Q26 營業利益率 2.6%，市值主要押在三個尚未有分部揭露的業務。

## 分部（個股模型原則第 2 節）

| 列 | 口徑 | 驅動 |
|---|---|---|
| 汽車分部（官方） | 汽車營收＋服務及其他 | 子列：車輛銷售（原檔量×價）、FSD（原檔活躍數×月費 × `fsd_k`）、Robotaxi（FY2027 起讀分部 v0.2 `RT_REV_<情境>_<年>` × `rt_k`；FY2026 以前原檔第 131 列）、碳權、租賃、服務及其他 |
| 能源分部（官方） | 能源發電與儲存 | 原檔儲能 MWh × ASP |
| Optimus（尚未列報） | 模型新增，歷史為 0 | Optimus 頁：交付量 × 售價；BOM 學習曲線（讀線框架 2b） |

`segments.json`＋`tools/segment_check` 通過；PL 頁「檢查：分部合計 − 合併營收」全期為 0；Hist 頁四季與 FY2024／FY2025 分部對帳為 0。

## 估值（DCF 50%＋分部加總 50%；三情境 25／50／25）

- DCF：標準 FCFF，明確期 FY2027–FY2035（原檔引擎期間）＋4Q26 stub；WACC 約 12.7%（rf 5.00%、β 兩來源 Blume 中位數 1.556、ERP 5%）；g 3%、RONIC＝WACC；淨現金與投資含 SpaceX 持股 2,865、比特幣 962（百萬美元）。
- 分部加總：沿用原檔「長期評價分析」2035 年倍數法（車輛 P/E 12、非汽車業務 P/E 15、成長分部 P/S＝年增率×100×0.8，折現 15%、10 年）；Robotaxi 讀 `tesla_robotaxi` 分部 v0.2 的 `RT_VAL_A/B/C`（Derived）：133.2／−1.0／−12.0 十億美元（v0.1 暫定 1,200／627／150 已取代）。
- DCF 腿的 Robotaxi 營收 FY2027–FY2035 同樣讀分部 v0.2（`RT_REV_<情境>_<年>`，百萬美元），取代「原檔第 131 列 × `rt_k_*`」；兩條 Robotaxi 路徑不再並存。
- 建置快照（build_id v0.2-daa4a8f80d）：DCF 32.70（A／B／C 89.65／20.57／0）、分部加總 126.92（180.70／118.11／90.77），加權目標價 **79.81**，上檔空間 −79.5%，估值評等**強力放空**（距放空邊界 29.5pt）；收支關卡情境 B −46,157 百萬美元＝不通過；投資建議強力放空。
- Robotaxi 在加權目標價的占比降到約 6%（v0.1 約 55%）；評等不再依附 Robotaxi 輸入，承重變數轉為 DCF 腿與 SOTP 非 Robotaxi 部分（Optimus、FSD、車輛）。
- 與對帳 1 預估（91.4）的差：對帳 1 只換 SOTP 腿、DCF 腿維持 55.8；v0.2 DCF 腿改讀分部營收後降為 32.7（−23.1，對加權 −11.5）。SOTP 腿與預估一致。情境 C 的 DCF 每股價值為 0（股權價值為負，原檔 MAX(0,·) 限責處理）。

## 主要假設（承重者附可證偽指標）

| 鍵 | 值 | 標記 | 可證偽指標 |
|---|---|---|---|
| `rt_val_A/B/C` | 133.2／−1.0／−12.0 十億美元 | Derived（分部 v0.2；其輸入多為 Assumed） | 分部與線 1a：每車每日付費英里、價格減全成本；德州、加州車籍 |
| `rt_rev_<年>_<情境>` | 分部 v0.2 營收（B FY2030／FY2035：2,550／28,749 百萬美元）；`rt_k` 單一 Sens 乘數 1.0 | Derived／Assumed | 同上 |
| `opt_u_*`、`opt_k_*` | 基準 FY27 2,000 台 → FY35 130 萬台；A ×1.6、C ×0.25 | Assumed | 法說揭露的對外交付台數（公司稱 2027 年底起） |
| `opt_bom0`、`opt_lr` | 5.5 萬美元、15% | Interested-party（MS）、Analogy | Tesla 成本說法、供應商報價 |
| `opt_asp27→35` | 4.0 → 2.5 萬美元 | Assumed | 首批工業客戶售價 |
| `cx_r27→cx_ss` | 核心資本支出／營收 25% → 8% | Assumed／Analogy | 2026Q4 法說 2027 資本支出指引 |
| `fsd_k` | 1.0（0.6–1.4） | Assumed | EU 2026-12 表決、FSD 活躍數季增 |
| `w_dcf` | 0.5 | Assumed | 方法判斷（見判斷題） |

## Robotaxi 兩條路徑的待對帳項

DCF 腿的 Robotaxi 毛利仍用原檔邏輯：分部營收與原檔路徑之差 × `rt_gm`（55%）。分部 v0.2 的每英里經濟在情境 B 近零毛利（EBIT 率約 0–8%），兩者不同口徑，屬待對帳項；方向上 DCF 腿對 Robotaxi 偏寬鬆。

## 判斷題

DCF 到 FY2035 截斷、終值價值中性；FY2035 時 Robotaxi 年增仍約 30%。分部加總以 2035 年成長率給 P/S。兩法對 2035 年後成長的隱含假設相反，加權結果貼近評等邊界，`w_dcf`（Sens 0.3／0.7：246.80／137.64）是最大的方法判斷。

## 相對原檔改了什麼

1. **新增工作表**：README、In、Hist、Optimus、PL、DCF、SOTP、Val、Sens、Recon、對帳、風險登錄、改版紀錄_v0.2；原檔各頁保留（`長期評價分析` 加註「舊法，不進結論」）。
2. **原檔年度欄公式修正**（只影響年度顯示欄）：Non-GAAP `CI34`、`CI51`、`CI52`、`CI147`、`CI149` 誤取 AA:AD（應為 FY2025 的 AD:AG）；`CJ78`–`CM78` 業外總額年度範圍錯位。
3. **滾入 3Q26 後新資訊**：Cybertruck 2027 年中停產（Musk 2026-11-06 X 貼文，[Interested-party]）→ `AN86`、`AO86` 由 0 改 −1；第 86、124 列預估期年增率包 IFERROR。
4. **分部加總**：非汽車業務 2035 年價值改用 FY2035 營收（原檔 `V5＝N30+N31` 誤用 FY2030）；新增 Optimus；Robotaxi 改讀輸入。
5. **新增 DCF 腿、資產負債與股數**（SEC XBRL 10-Q 2026-09-30）、2025 CEO 獎酬（12 批市值閘門，扣每股抵付 334.09）。
6. **未改（保留原檔判斷）**：量價、毛利率路徑（3Q26 經常性 16.5% 起每季 +0.1pt）、FSD 滲透率、Robotaxi 車隊路徑、研發費用公式、稅率 25%、原檔「待決定」五項的預設值。新資料點（德州 Robotaxi 車籍 >1,200 台、其中 Cybercab 約 570；Fort Worth 11-02、San Antonio 11-09 開通）記在對帳頁，車隊路徑未改，待 tesla_robotaxi 對帳。
7. **發現未修**：原檔 FY2024 汽車營收比 10-K 多 100（引擎 77,170 vs 10-K 77,070；Hist 頁以 10-K 為準，「對原檔引擎差」欄列出）；`Non-GAAP_legacy` 隱藏頁的既有錯誤值未動。

## 對帳（線層級變數）

集中在「對帳」頁：FSD 活躍數與月費（1b）、Robotaxi 車隊、票價、每英里成本、分部價值（1a）、人形 BOM、學習率、交付時點、工人時薪、生產力（2b）、訓練算力（只讀，未取得）。各列寫明本模型使用值、口徑、出處、標記與影響的儲存格。

## 限制與未取得資料

- 依規格略過：Siliconomics／Tokenomics／Wattonomics 協調（算力線）、自用 AI 讀法（本模型無自用算力影子效益）。
- 未做：同業倍數（原則 7、30）、HTML 一頁摘要、自算產業 β、反向 DCF 整本反解（Recon 只有 Robotaxi 價值的線性反解）。
- 未取得：Tesla 自建訓練算力支出（AI供應鏈只讀）、2027 資本支出指引、市場共識 EPS、Optimus 任何財務數字、Robotaxi 每英里成本、FSD 毛利率。
- rf 取 2026-10-23（逾期，與其他個股模型一致）。
- 分部加總沿用原檔「不加淨現金」口徑。

## 核對（建置後已跑）

`tools/in_layout/check_in_layout.py`、`tools/sens_coverage/check_sens_coverage.py --require-keycol`、`tools/model_governance/check_rating_gate.py --strict`、`tools/model_governance/check_governance.py`、`tools/segment_check/check_segments.py physical_ai/tesla` 全部通過。數字尚未經未參與製作的代理獨立核對（Project 規則第 10 條）。

## 改版紀錄

| 版本 | 內容 |
|---|---|
| v0.2（2026-11-16） | Robotaxi 分部價值與 FY2027–FY2035 營收改讀分部 v0.2（對帳 1）；`rt_k_*` 移除、`rt_k` 成為 Sens 乘數；Val／SOTP／README／Recon／對帳頁文字更新；加權目標價 192.2 → 79.8，評等強力放空不變 |
| v0.1 | 首版 |

## 公開與底稿（2026-11-17）

- Andy 2026-11-17 08:25 同意本模型放進公開 repo（「可以」）。
- 底稿原檔 `20261028_TeslaTSLA.US_post_3Q26.xlsx` 不放 repo，權威位置在 Andy 私人 Drive（AI實體應用 `個股模型/_原始資料/`）；原檔全部工作表已保留在 `dist/` 成品內，可直接查閱。重建時以 `--base <原檔路徑>` 指定。
- 口徑提醒：DCF 腿情境 C 每股價值為 0（股權價值為負，沿用原檔 MAX(0,·) 限責處理）；分部加總腿的 Robotaxi 情境 B、C 為負值且**不歸零**。兩條腿的處理不同，是刻意的，不是錯誤。
- 待下一版處理（需重建）：成品中 17 處「研究者先驗」改為「原檔假設（非外部來源）」；DCF 腿 Robotaxi 毛利（rt_gm 55%）與分部 B 情境近零毛利的對帳。
