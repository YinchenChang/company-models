# 工作單｜Wattonomics v0.1.7：資金缺口頁 `Gap`＋互動頁「整個市場當一間公司」圖｜2026-10-31 r0（chat 端）

**規劃：** `wattonomics/docs/plan/20261031_主要玩家資金缺口_規劃.md` r2（以下稱「規劃書」）。本工作單只把規劃書落成可建置的規格；規劃書與本單衝突時以規劃書為準，並在報告中列出。
**基底：** main `7dcdd32`，現行 `wattonomics/model/20261031_Wattonomics_v0.1.6.xlsx`。
**分支：** `claude/wattonomics-v0.1.7`。
**不在本版：** Amazon 2026Q3 10-Q 項目（Andy 尚未下載，留待 v0.1.8）；K4；C2–C7。

## 0. Andy 原話（固定規格；README 已寫，本版不得偏離）

| Andy 原話（2026-10-31） | 參數 |
|---|---|
| 「依照我們…追蹤到這個市場總共要蓋幾 G 瓦的 data center，然後每座 data center 要花多少錢來判斷」（10:22） | 需要的錢＝新增 GW × 每 GW 建置成本；不用公司公告的資本支出金額作基準 |
| 「以建設 data center 所需要的土地、廠房、設備，所需要花的錢作為依據，這樣才可以跟對得上那些供應商的營收」（10:22） | 每 GW 成本＝TK `IF_CapexTotal`（IT＋廠房），與 Spend／Facility 同一套數字；可展開成項目 |
| 「他們大玩家之間的租賃什麼的，就不要重複計算到了」（04:13） | 每 GW 只在擁有者（營運者）計一次 |
| 「營收賺的不只是 AI 帶來的營收，還有…他們的本業」（04:13） | 資金來源＝全公司營業現金流 |
| 「那個股票買回跟股利就先不要算好了，寬容一點…」（04:20） | 不扣買回與股利；另列備忘「可削減緩衝」 |
| 「把整個市場當做一間公司，他要蓋那麼多的 data center，需要多少錢？然後目前錢準備夠了嗎？」（10:22） | 互動頁在需求、延遲選擇器下方加圖：需要 vs 來源、尚需新融資、累計缺口 |

## 1. 判斷題（chat 端預設；照預設建置，Andy 可撤回）

| 題 | 預設 | 理由 |
|---|---|---|
| G1 規劃書 Q1 成本口徑 | 主口徑＝資本支出（現金）；年化持有成本（TK `AL_HoldGen` 或 `IF_HoldGen` 世代加權）× 年平均 GW 只作 Gap 頁對照列 | 規劃書預設 |
| G2 規劃書 Q3 非 AI 資本支出 | 計入「需要」，但在圖上與報表都獨立成一層（「12 家非 AI 資本支出」），使用者可一眼扣除。來源＝各個股模型的非 AI 資本支出列；模型沒有拆分的公司標「無模型值」並以 0 計（報告列出） | 12 家營業現金流是全公司口徑，用途側也要全公司才對稱 |
| G3 規劃書 Q4 判讀門檻 | 規劃書預設（至 2030 市場合計自由現金流由負轉正＝Anthropic 型；未轉正或缺口 > 資本支出 30% 連 3 年＝OpenAI 型）。**「市場合計」只含 12 家＋市場其餘的資本支出**：市場其餘沒有來源，所以另出「只看 12 家」的判讀並列 | 避免市場其餘無資料把判讀一律推成 OpenAI 型 |
| G4 前瞻年份的已到位外部資金 | 2024–2025 用 Funding 頁已觀測融資（總額口徑）；2026 起「已到位」只算可驗證的存量：2025 年底 12 家現金及約當現金（期初現金，一次性）＋ Funding 頁「尚未起租」租賃與表外合資承諾（存量，一次性）。兩者只進**累計**缺口，不分攤到各年；年度圖的 2026 起只有營業現金流 | 承諾的入帳時點無揭露；分攤到各年會是 Assumed 時程 |
| G5 營業現金流不隨需求情景變動 | 是（個股模型基準情境只有一條）；頁面與互動頁註明「來源固定為個股模型基準，不隨 A／B／C 變」 | Wattonomics 不重估個股模型權威的量 |
| G6 年份 | 2024–2030；2028 起標「情景（覆蓋率 < 0.35）」；個股模型沒有的年份標「無模型值」、不外推 | 規劃書 §2、§4 |

## 2. 頁面 `Gap`（新頁，置於 Funding 之後；計算一律 Excel 公式）

### 2a. 每 GW 建置成本（逐年，US$B／IT GW）
- 直接沿用 CapexGW／C1 既有的世代加權每 GW 成本（`IF_CapexIT`＋`IF_CapexFacility`，NonNV × 持有比 × VR200；廠房依 `lag_bldg` 提前入帳，權重 w0／w1／w2）。**不得另寫一套。** 若現有公式只在 CapexGW 以「資金 → GW」方向出現，在 Gap 頁以同一組儲存格引用重組成「GW → 資金」方向，並加 Checks 列驗證：CapexGW 的「當年上線 GW 的資金」÷ 每 GW 成本＝甲2 年新增 GW（同源一致性）。
- 項目展開（土地、建築、配電、冷卻、發電機、IT 等）讀 Facility 頁既有拆分比例；項目合計＝每 GW 成本（Checks）。

### 2b. 市場層「需要多少錢」（需求 A／B／C × 延遲 D0／D1／D2，2024–2030）
- AI 資料中心資本支出＝IT 部分：當年新增 GW（Deploy 頁 `na<s><D>`，淨新增；若淨新增含汰換抵減，改用毛新增並在報告說明）× 當年 IT 每 GW；廠房部分：w0 × 當年＋w1 × 次年＋w2 × 後年 新增 GW × 廠房每 GW（建物先付）。2030 年廠房要用到 2031、2032 的新增 GW（Deploy 已有到 2032）。
- 拆成「12 家擁有」與「市場其餘」：12 家擁有＝Players 頁各營運者年新增 IT GW（主路徑；七家甲2、其餘甲）× 同一每 GW 成本；市場其餘＝市場合計 − 12 家（只在 A／D0 主路徑上有此拆分；其他情景市場其餘＝市場合計 − 12 家 A／D0 值，並在註記寫明 12 家不隨情景變）。**若 Players 只有 A 情景，照此辦理；若 Players 有情景維度，照實際維度。** 負值不截斷，標「12 家 > 市場」，列入 Checks WARN。
- 年化持有成本對照列（G1）。

### 2c. 玩家層（12 家 × 2024–2030）
- 列：AI 資本支出（新增 GW × 每 GW；擁有者口徑）｜非 AI 資本支出（G2）｜財報／模型資本支出（對照，C1）｜營業現金流｜自籌比＝營業現金流 ÷（AI＋非 AI 資本支出）｜缺口＝（AI＋非 AI）− 營業現金流（正＝需外部資金）｜備忘：買回＋股利（可削減緩衝，不進缺口）｜已觀測融資（2024–2025，Funding 總額口徑，拆股權、債、租賃與表外）｜推算尚需＝缺口 − 已觀測（2024–2025）｜判讀列。
- 12 家：Microsoft、Alphabet、Amazon、Meta、Oracle、CoreWeave、Nebius、WhiteFiber、SpaceX／xAI、OpenAI、Anthropic、智譜（智譜只列，不進合計；報告寫明）。所以合計實為 11 家，頁面標題用「主要玩家（11 家合計＋智譜只列）」。
- 財年：照 In 財年月與既有 C1 換算方式轉曆年（Microsoft 6 月、Oracle 5 月）。

### 2d. 「整個市場當一間公司」（每個需求 × 延遲情景一組，2024–2030；互動頁的資料源）
年度：
1. 需要：AI 資料中心資本支出（12 家擁有）｜AI 資料中心資本支出（市場其餘，「資金來源未知」）｜12 家非 AI 資本支出。
2. 來源：12 家營業現金流（不扣買回股利）｜已觀測外部融資（只 2024–2025）。
3. 尚需新融資＝需要（不含市場其餘）− 來源；另列「含市場其餘」版本。
累計（自 2026 起）：累計尚需新融資 − 期初現金 − 已簽未起租承諾（G4）＝累計仍未安排；兩個版本（不含／含市場其餘）。
判讀：G3 兩種讀法並列（含市場其餘／只看 12 家），輸出「OpenAI 型」「Anthropic 型」或「未定（資料不足）」。

### 2e. 前瞻營業現金流與非 AI 資本支出的取數（新工具）
- 新增 `wattonomics/tools/import_gap_models.py`：依 `markets/invest_map/data/models_extract.json` 的 `model` 路徑開各個股模型（main 上的現行版），讀基準情境的年度營業現金流、AI／非 AI 資本支出、總資本支出、買回＋股利、2025 年底現金（能取到才取），寫 `wattonomics/builder/seed/gap_models_v0.1.7.json`；每個值附 `file`、`sheet!cell`、標籤原文、財年。
- 三種模型家族：MAG 型（`CF` 頁「營業現金流」列；MSFT 已確認 CF!A23），CoreWeave／Nebius／WhiteFiber 型（`各期收支`、`輸入與假設` 的 CFO 列），OpenAI／Anthropic／智譜 型（`Funding` 頁）。逐家人工確認列位，**不要靠關鍵字自動抓第一個命中**；把每家選用的列與理由寫進報告表。
- 找不到的值寫 null 並標「無模型值」，不外推、不補估。
- build.py 讀此 seed，寫入 In 頁新節「Y9. 資金缺口：個股模型基準值」（藍字；標記：OpenAI／Anthropic／xAI 內容屬 [Interested-party] 或 [Derived]，依該模型自己的標記；來源 ID 新增 SRC_GL_3xx 對應各模型檔）。之後各版 `--base` 讀回時 In 為權威，seed 只用於首建本節。

## 3. Interface 名稱（新增；`WIF_Gap_` 前綴）
- `WIF_Gap_PerGW_<yyyy>`（US$B／IT GW）
- `WIF_Gap_Need_<A|B|C>_<D0|D1|D2>_<yyyy>`（12 家擁有＋市場其餘＋12 家非 AI）、`WIF_Gap_NeedOwn_*`、`WIF_Gap_NeedRest_*`、`WIF_Gap_NonAI_<yyyy>`
- `WIF_Gap_OCF_<yyyy>`、`WIF_Gap_ExtObs_<yyyy>`（2024–2025）
- `WIF_Gap_Short_<s>_<D>_<yyyy>`、`WIF_Gap_ShortInclRest_<s>_<D>_<yyyy>`、`WIF_Gap_CumShort_<s>_<D>_<yyyy>`、`WIF_Gap_CumShortInclRest_*`
- `WIF_Gap_OpenCash`、`WIF_Gap_Committed`
- `WIF_Gap_Player_<player>_<yyyy>_<OCF|CapexAI|CapexNonAI|Gap|SelfFund|Buyback>`
- `WIF_Gap_Verdict_<s>_<D>`、`WIF_Gap_Verdict12_<s>_<D>`（文字）
快照 `export_snapshot.py` 照既有方式全收；名稱數變動在報告列出。

## 4. Checks（新增，WGOV_Errors 計入 ERROR 者）
- G-K1（ERROR）：每 GW 成本項目合計＝每 GW 成本（1E-6）。
- G-K2（ERROR）：同源一致：CapexGW「當年上線 GW 的資金」÷ 每 GW 成本＝甲2 年新增 GW（七家、2027 起；1E-6）。
- G-K3（ERROR）：市場層 A／D0 的 12 家擁有＋市場其餘＝市場合計。
- G-K4（ERROR）：每情景 尚需新融資＝需要 − 來源；累計＝逐年加總 − 期初現金 − 承諾。
- G-K5（ERROR）：缺口與自籌比的分子分母不含買回股利（以公式文字檢查：Gap 頁缺口列不得引用買回欄）。
- G-K6（WARN）：市場其餘 < 0 的年份；無模型值的公司 × 年份數。
- G-K7（WARN）：Gap 頁 AI 資本支出（12 家，2024–2026）與財報／模型資本支出 × AI 占比的比值超出 `tolband`（C1 對照；只警告）。

## 5. 互動頁（`explorer/template.html`、`tools/export_explorer.py`）
- `export_explorer.py` 新增 kind `gap`（每情景年度與累計、每 GW 成本與項目、玩家層表）。
- 頁面：在需求 A／B／C 與延遲 D0／D1／D2 選擇器**正下方**新增區塊「整個市場當一間公司：要蓋這麼多資料中心，錢夠嗎？」
  - 圖 1（年度，2024–2030）：左柱堆疊「需要」三層（12 家 AI 資料中心、市場其餘〔另色、斜線或淡色，圖例「資金來源未知」〕、12 家非 AI）；右柱堆疊「來源」兩層（營業現金流、已觀測外部融資）；折線或標記「尚需新融資」。跟著選擇器重畫。
  - 圖 2（累計，2026–2030）：累計尚需新融資、扣期初現金與承諾後的「仍未安排」；不含／含市場其餘兩條。
  - 一行判讀文字（G3 兩讀）＋註記：「來源固定為個股模型基準，不隨情景變」「2028 起覆蓋率 < 0.35，屬情景」「不扣買回股利（備忘：累計可削減緩衝 X）」。
  - 可展開表：每 GW 成本項目（土地、建築、配電、冷卻、發電機、IT…）；玩家層表。
- 先讀 repo 根目錄或 `dataviz` 相關規範（若 repo 有圖表規範檔）；顏色沿用 template 既有色票。純 SVG 或 template 既有圖表方式，不加外部 CDN。
- `build_page.py` 產頁後依 `markets/invest_map/publish_pages.sh` Wattonomics 段補 doctype 骨架；**本工作單不發布 gh-pages、不更新 artifact**（chat 端在合併後做）。

## 6. 建置順序與交付
1. `import_gap_models.py` → seed。
2. `build.py --base wattonomics/model/20261031_Wattonomics_v0.1.6.xlsx --version 0.1.7 --date 20261031 --out build/raw/20261031_Wattonomics_v0.1.7.xlsx`（參數照 build.py 實際介面）。
3. soffice 重算 → `wattonomics/model/`；錯誤格 0。
4. `export_snapshot.py`、`export_explorer.py`；CURRENT 兩檔更新；舊 v0.1.6 xlsx 照慣例移 `model/archive/`。
5. `check.py`：既有全過＋新 Checks；**既有 `WIF_` 名稱數值 0 變動**（逐名比對 v0.1.6 快照，差異清單必須為空；新增名稱另列）。
6. README：頁面表加 Gap 列、Interface 列加 `WIF_Gap_*`、版本紀錄加 v0.1.7；Decisions 加 GL01-xx（G1–G6 預設）。
7. 報告 `wattonomics/docs/verify/20261031_v0.1.7_gap.md`：
   - 一句結論（A／D0 與 B／D1 的判讀）；
   - 每家取數表（檔、儲存格、標籤原文、財年、值）；
   - 市場層 A／D0、B／D1、C／D2 三組年度表（需要三層、來源兩層、尚需、累計）；
   - **敏感度先於基準**：`lag_bldg` 低／高、`ai_sh` 低／高（只影響 C1 對照與 12 家甲2 GW）、延遲情景，對 2027–2030 累計缺口的影響；
   - 「無模型值」清單；限制（2028 起情景；OpenAI／Anthropic 營業現金流屬 Interested-party 或模型推估）。
8. 開 PR（草稿亦可），標題「Wattonomics v0.1.7：資金缺口頁與『整個市場當一間公司』圖」。**不要合併。**

## 7. 停止條件（遇到即停、寫報告、交回 chat 端）
- 需要改 TK、Siliconomics 或任何個股模型的值 → 停，寫 Requests 頁。
- 既有 `WIF_` 名稱任一數值變動 → 停（目的：本版只加頁，不得改動既有輸出；若是 LibreOffice 浮點雜訊 < 1E-9 可列表說明後繼續）。
- 某家模型找不到營業現金流列 → 不停，標無模型值並列出。
- 不得在任何對外請求中使用使用者 email；不得改核對代理的腳本邏輯；不得 force push；不得動 gh-pages。
