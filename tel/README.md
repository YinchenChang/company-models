# tel/ — 東京威力科創（Tokyo Electron，TEL，8035.T）收支與估值模型 v0.1.0

## v0.1.0（2026-11-17）：首版（工作單 new5）

- 範圍：單一報導分部（營收）；新設備依終端市場（DRAM、NAND、非記憶體＝先進邏輯＋其他代工邏輯＋其他前段）；Field Solutions 與其他。預測到 FY2031–32（模型 FY2025＝公司 2026年3月期；公司稱 FY2026）。
- 財務數字取自公司原文（2027年3月期 第2四半期決算短信、說明資料、Q&A、逐字稿，2026-10-30；網頁擷取工具讀取，為轉錄，數字以算術互核，版面與頁碼待人工複核）。歷史長序列取 stockanalysis.com、finboard.net（第三方轉錄，Verified 以公司原文互核者為限）。Yahoo Finance 只作共識與目標價對照，不進損益輸入。
- 價格只取 `markets/invest_map/data/closes_20261020_20261113.csv` 與 `closes_extra_20261113.csv`（2026-11-13 收盤）。
- 結果：現價 9,461；DCF 目標價（情境加權）4,073；Comparable 10,999（未調整）／9,360（相對倍數調整後）；加權目標價 6,716（DCF 50%＋Comparable 50%）；上檔空間 −29.0%；評等 放空（帶：>+50% 強力買進、>+25% 買進、≥−25% 中立、≥−50% 放空、其餘強力放空）。
- 另一側讀數（C4，Wattonomics 錨定）：加權 5,635（−40.4%），評等 放空。
- 上檔空間超過 ±25% 的主因：DCF（4,073）遠低於現價。WACC 12.3%（β_u 1.61、無負債；rf 2.82%、ERP 5%、CRP 0.91%）使終值折現重；現價隱含的前瞻 P/E 25.2 倍與同業中位數 24.3 倍相近，Comparable 因此較高，但相對倍數調整係數 0.85 再向下。見 Excel Val、DCF、RevDCF 頁。

### 引擎

- 五個設備市場（先進邏輯、其他代工邏輯、其他前段、DRAM、NAND）讀 `markets/equipment` Mkt 5A／5B／5C（同字母情境）。FY 指數＝3/4·CY(t)＋1/4·CY(t+1)。
- 新設備＝FY2026 錨（公司指引）× 終端市場 FY 指數 × 超額成長因子 × 中國因子；Field Solutions 及其他＝上年 ×（1＋gfso）。
- 營業利益＝FY2026 錨＋（邊際營業利益率 mm＋dm）× 營收變化（情境 C 自 FY2028 另加 madjC × 營收）；淨利＝營業利益 × nioi；股數每年減 bb。
- FY2026 錨定公司指引：營收 3,700、營業利益 1,169、淨利 865、EPS 381.69（¥B，EPS 為円）。
- DCF 為標準 FCFF（`tools/fcff_dcf.py`；資產負債表日 2026-09-30、評價日 2026-12-31）。
- 相對倍數（`tools/relative_multiple.py`）：8035 不在 `target_summary`，以建置時替換 `_universe`／`_peers` 的方式套用；同業前瞻 P/E（2026-11-13）AMAT 21.02、LRCX 21.98、KLAC 27.88、ASML 23.53、Disco 25.06、Lasertec 32.31，中位數 24.295。十年參考 P/E 取 Bloomberg 前瞻 P/E 各年中位數（`tools/data/fwd_mult_hist.csv`），因年底收盤價無第二來源。

### 與 Yahoo Finance 的差異

- Yahoo 共識 FY2027（2027年3月期）EPS 517.6、營收 4,487；本模型情境 B EPS 457.5（A 490.4、C 408.0）。差異來源：本模型錨定公司指引並把市場層設備規模（CY2026 160.0B 美元，與公司說的約 160B 一致）往後推，CY2027 為 192.8B，公司說 >210B，共識傾向接近公司說法。
- Yahoo 目標價只作對照；賣方立場見 In 頁「來源誘因」欄。

### 待判斷（已選預設值，範圍如下；請人工複核）

| 項目 | 預設（低／基準／高） | 說明 |
|---|---|---|
| 新設備超額成長 exc（A／B／C） | A 0/2%/4%；B −2%/0/2%；C −4%/−2%/0 | 份額＋組合；Assumed，進 Sens |
| 中國新設備營收水準 chh（B／C） | B −10%/−5%/0；C −30%/−20%/−10% | Assumed |
| Field Solutions 年增 gfso（A／B／C） | A 8/12/16%；B 4/8/12%；C 0/3/6% | Analogy（參 LRCX、AMAT 服務與備品列），進 Sens |
| 邊際營業利益率調整 dm、情境 C 回吐 madjC | dm −3/0/+3pt；madjC −3/−1.5/0pt | Assumed |
| 淨買回 bb | 0／0.8%／1.5%／年 | Assumed |
| 目標時點每股股利 dps_tgt | 237.4 円 | Assumed（以公司配當方針推算） |
| FY2030 設備市場相對 CY2029 g_tail | −10%/0/+10% | 市場層只到 CY2029，Assumed |
| 情境權重、DCF 權重 | A25/B50/C25；DCF 50% | Assumed；DCF 權重 30%／70% 在 Sens |
| 租賃負債、非營業投資 | 0、0 | 未取得明細，Assumed；Sens 有「淨現金不含權益投資」 |
| 十年參考 P/E | Bloomberg 前瞻 P/E 年中位數 | 無年底收盤第二來源 |
| 價格 | 11-13 收盤 | 若主串另有換價工作單，以其為準 |

### Sens 範圍結果（加權目標價相對基準 6,716）

- DCF 權重 0.3／0.7 為最大幅度（−25.0%／+25.9%）；其餘案例（超額成長、中國、Field Solutions、尾端、dm、madj、nioi、bb、錨定）約在 ±6.4% 內。
- 已知現象：FY2030 後延伸期的營業利益＝淨利 ÷ nioi，故 Sens 的 nioi 案例中 DCF 與 nioi 反向移動。

### 建議補進 Siliconomics

- TEL 實際用到的變數：Siliconomics v0.20 的 SIF_WTAnchor_Base（wta_27／28／29＝0.734／0.649／0.629）、SIF_RecoFlag、SIF_DeployGW_AltRatio；設備市場 markets/equipment 的 la／lo／of／dr／nd。
- 建議補進：TEL 塗佈顯影（Coater/Developer）全球份額與其對應設備市場、蝕刻與成膜各自的市場份額與年份序列（目前只有五市場合計）、中國設備需求占比的基準路徑（本模型以 chh 在個股層處理）、NAND／DRAM 先進製程轉換的設備密集度。

### 檢查

- 重算錯誤格 0（#REF!／#VALUE!／#DIV/0!）；check_governance 通過（build_id 見 Excel README A1 與 Sens 頁）。
- check_sens_coverage：同 accretech、amat，找不到 Sens 表頭而未過（全 repo 既有基線，非本模型獨有）；未修改 tools/。
- Coord：`coord_alt.json` 由 `tel/alt_run_tel.py` 產生（accretech 版 `tools/coord/alt_run.py` 的複本，工作流查找改為固定指令）。待主串加入 `coord_meta`、workflow 與 `target_summary` 項目。

### 檔案

- `build_tel.py`（建置）、`tel_post.py`（後處理）、`sens_tel.py`（敏感度快照，寫 `sens.json`）、`alt_run_tel.py`（兩讀數快照，寫 `coord_alt.json`）、`segments.json`、`html.json`。
- 成品：`dist/20261117_東京威力科創個股模型_v0.1.0.xlsx`、`.html`。
- 重建順序：`alt_run_tel.py`（`LO_PROFILE=file:///tmp/lo_<名>`）→ `sens_tel.py` → `build_tel.py <輸出.xlsx>` → LibreOffice 重算到 dist。改 In 頁後須重跑前兩步，否則 Sens 頁「模型狀態」為 STALE。
