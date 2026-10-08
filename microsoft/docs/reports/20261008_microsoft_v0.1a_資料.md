# Microsoft v0.1a 資料蒐集報告（2026-10-08）

工作單：`mag/docs/workorders/20261008_mag_v0.1a_資料蒐集.md`；分支 `claude/microsoft-v0.1`；PR #33。本報告為人讀版本，機器可讀的唯一來源是 `data/microsoft_facts_20261008.json`（事實總帳）、`data/kevidence_microsoft_20261008.json`（k 證據）、`data/consensus_msft_20261008.json`（市場共識）、`data/tokenomics_snapshot_v5.27.json`（Tokenomics 快照）與 `company.json` → `mag` 區段。金額單位除另註明外為十億美元（US$bn；總帳為百萬美元）；Microsoft 財年結束於 6/30（FY26＝2025/7–2026/6）。

## 1. 結論（先讀這段）

1. **一手資料齊全，未觸發停止條件。** SEC EDGAR 直接取得 FY26 10-K（經查核，2026-07-29）、FY26 三份 10-Q、FY25–FY26 八季財報新聞稿（8-K 附件 99.1）、2026-09-02 8-K（FY27 分部重編與歷史重述，**首次揭露 Azure 絕對金額**）、XBRL companyfacts；法說內容以 WebFetch 讀 Microsoft IR 網站。總帳 **722 筆**：[Verified] 324、[Interested-party] 347、[Derived] 39、[Analogy] 6、[Assumed] 6。FY27Q1（季末 2026-09-30）財報預計 2026-10-28，尚未公布。
2. **FY26 核心數字**（10-K）：營收 331.8（+18%）、營業利益 155.2；舊三分部 PBP 140.0／83.9、IC 137.8／57.0、MPC 54.1／14.4（營收／營業利益）。FY27 新分部（重述）：Agents and Infra 268.1／136.4、Devices and Consumer 63.7／18.9。**Azure（新定義）FY26 101.9、Q4 29.4（年增 42%），年化 117.7**。營運現金流 182.9；現金資本支出 115.9、**含融資租賃 145.3**（融資租賃 28.1）；FCF 67.0；現金＋短期投資 76.8；債券面額 46.1（無商業本票）；融資租賃負債 66.6、營業租賃負債 21.9；**尚未起租租賃 329.1**（FY27–FY33 起租，1–20 年；一年前的 10-Q 序列 106→155→197→329）；回購（現金）22.3、股利 26.4；商用 RPO 678（12 個月內約 30%）。
3. **指引**：FY27Q1 營收 89.85–90.95、資本支出 >50（含租賃改分類影響）、Azure +44–45%（cc）；FY27 全年資本支出「年增」（無金額）、營業利益率下降 <1pt；**曆年 2026 資本支出約 175**（原約 190；差額為更多資料中心租約改列營業租賃）。FY27 起資料中心與辦公建物耐用年限 15→25 年。
4. **AI 容量（MW）沒有總量揭露，只有增量**：FY25 新增 >2 GW、FY26 Q2–Q4 各約 1 GW（口徑不明）。第三方：Bloomberg 稱總容量約 12 GW（2026-09）、2032 年目標約 38 GW；Morgan Stanley 估 FY24 約 5 GW；Epoch 兩座 Fairwater 在役 1.0 GW-IT。三法交叉的**期初 AI 在役 3.5／4.5／5.5 GW-IT**（低／基準／高；未用營收）。另向 neocloud 租用約 0.8–1.5 GW-IT（CoreWeave、Nebius、IREN 等；租金屬營業成本）。
5. **k**：Azure 公開牌價換算 k 很高（隨需 6.6–13.0、預留 1–3 年 3.3–8.3），但大客戶成交價未揭露，只作上限。長約證據：IREN–Microsoft 0.76（一手 MW）、Nscale 約 0.88（金額二手）、Anthropic–Azure 0.39–0.71（年期假設）、Nebius ≤1.33。**建議 k_長約 0.76（0.45–1.00）、k_現貨 1.76（1.5–2.3）、長約占比 70% → 加權 k 1.06（0.76–1.31）**，與 nebius／coreweave 基準相同。
6. **AI／非 AI 雲端拆分（試算）**：對外計費 AI 約 3.25 GW-IT × 加權 IF_HoldEcon 10.60 × k 1.06 ＝ **對外 AI 雲端約 36.5／年（17.3–67.7）**，Azure 年化 117.7 減去得**非 AI 殘差約 81.2（50.0–100.3）**。驗證：公司 AI run-rate 37（FY26Q3，含自有 Copilot）扣 Copilot 後隱含對外 AI 約 26–30 → **基準可能偏高 20–40%**（不回頭改 k）；上限檢查 59% > 50% 示警。
7. **關聯方**：OpenAI 持股約 **25%**（10-K；重組時 27%）× 投後 852 → 213（未折價）；FY26 來自 OpenAI 營收 **24.1**（占 7.3%）；OpenAI 增量 Azure 承諾 **250**（年期、GW 未揭露）；FY26Q2 約 **45%** 商用 RPO 來自 OpenAI；OpenAI 付 Microsoft 營收分成 20% 至 2030、總額上限約 38（二手）。Anthropic：Azure 承諾 **30**（up to 1 GW）、Microsoft 投資上限 5、FY26Q4 認列利得 3.2，持股比例未揭露。
8. **市場共識**：收盤 **529.76**（2026-10-07）；目標價 S&P 平均 587.63（中位 586、440–870、56 家，Strong Buy）、MarketBeat 578.73（48 家）；FY27–FY29 營收 391.3／468.0／569.5、資本支出（推定現金口徑）192.8／216.8／251.2（MarketScreener）。
9. **verify.sh 19 項全過**（引擎未動，數字與基線相同；摘要見第 10 節）。

## 2. 核心數字表（chat 端抽查用）

| 項目 | 數值 | 期間 | 標記 | 總帳 id | 來源 |
|---|---|---|---|---|---|
| 營收／營業利益 | 331.839／155.237 | FY2026 | [Verified] | is.rev.fy2026／is.oi.fy2026 | [10-K](https://www.sec.gov/Archives/edgar/data/789019/000119312526323660/msft-20260630.htm) |
| 分部營收：PBP／IC／MPC | 139.996／137.791／54.052 | FY2026 | [Verified] | seg.*.rev.fy2026 | 10-K 附註 18 |
| 分部營業利益：PBP／IC／MPC | 83.879／56.972／14.386 | FY2026 | [Verified] | seg.*.oi.fy2026 | 10-K 附註 18 |
| 新分部營收／營業利益：A&I | 268.127／136.365 | FY2026（重述） | [Interested-party] | seg27.ai.* | [8-K 2026-09-02](https://www.sec.gov/Archives/edgar/data/789019/000119312526380280/d291965dex991.htm) |
| 新分部營收／營業利益：D&C | 63.712／18.872 | FY2026（重述） | [Interested-party] | seg27.dc.* | 同上 |
| 最新季營收／營業利益 | 90.007／40.603 | FY26Q4 | [Interested-party] | q.rev／q.oi.fy26q4 | [新聞稿](https://www.sec.gov/Archives/edgar/data/789019/000119312526323632/msft-ex99_1.htm) |
| 最新季分部：PBP／IC／MPC 營收 | 37.847／39.306／12.854 | FY26Q4 | [Interested-party] | q.seg.*.rev.fy26q4 | 同上 |
| 最新季分部：PBP／IC／MPC 營業利益 | 21.900／15.955／2.748 | FY26Q4 | [Interested-party] | q.seg.*.oi.fy26q4 | 同上 |
| 最新季新分部：A&I／D&C 營收 | 74.576／15.431 | FY26Q4 | [Interested-party] | q.seg27.*.rev.fy26q4 | 8-K 2026-09-02 |
| Azure（新定義） | FY26 101.938；Q4 29.417（年化 117.668）；FY25 72.610 | FY2026 | [Interested-party] | prod27.azure.* | 8-K 2026-09-02 |
| Azure 成長（舊／新口徑） | Q1–Q4：40／39／40／43%；新口徑 40／39／40／42%；FY 41%／40% | FY2026 | [Interested-party]／[Verified] | q.azure*.growth.* | 新聞稿、8-K、10-K |
| 其他產品線 FY26（新口徑） | M365 cloud 100.3；授權 37.3；搜尋與廣告 24.8；XBOX 21.8；Industry 20.3；Windows 17.1；Frontier 8.3 | FY2026 | [Interested-party] | prod27.*.fy2026 | 8-K 2026-09-02 |
| D&A（PP&E 折舊＋無形攤銷） | 34.3＋4.7＝39.0（現金流量表「折舊攤銷及其他」38.534） | FY2026 | [Verified]／[Derived] | pl.depr／pl.amort／cf.da | 10-K 附註 6、9 |
| SBC／利息費用／稅率 | 12.405／3.051（含融資租賃利息 2.547）／19% | FY2026 | [Verified] | cf.sbc／is.intexp／is.etr | 10-K |
| 營運現金流／現金資本支出／FCF | 182.935／115.948／66.987 | FY2026 | [Verified]／[Derived] | cf.ocf／cf.capex／cf.fcf | 10-K |
| 資本支出含融資租賃（公司口徑） | 34.9／37.5／31.9／41.0；合計 145.3；其中融資租賃 11.1／6.7／4.7／5.6 | FY26 各季 | [Interested-party] | capexFL.* | [IR 法說](https://www.microsoft.com/en-us/investor/events/fy-2026/earnings-fy-2026-q4) |
| 資本支出指引 | Q1 FY27 >50；FY27 年增；CY2026 約 175（原 190） | — | [Interested-party] | guid.* | IR 法說、8-K |
| 現金＋短期投資 | 76.843 | 2026-06-30 | [Verified] | bs.cashSti.fy2026 | 10-K |
| 債券面額／帳面／公允價值 | 46.136／40.294／36.5；FY27 到期 9.25 | 2026-06-30 | [Verified] | debt.* | 10-K 附註 10 |
| 融資／營業租賃負債 | 66.594／21.925 | 2026-06-30 | [Verified] | lease.finLiab／opLiab | 10-K 附註 13 |
| 尚未起租租賃 | 329.1（FY27–FY33，1–20 年） | 2026-06-30 | [Verified] | lease.notCommenced.fy2026 | 10-K 附註 13 |
| 建設承諾／採購承諾 | 34.566／194.060 | 2026-06-30 | [Verified] | commit.* | 10-K MD&A |
| 商用 RPO（12 個月內） | 678（約 30%；加權平均 2.3 年） | 2026-06-30 | [Verified] | rpo.commercial.fy2026 | 10-K 附註 12 |
| 流通股（封面） | 7.4255bn（2026-07-23）；期末 7.427bn | — | [Verified] | sh.out.* | 10-K |
| 股利 | FY26 每股 3.64（宣告 27.0）；FY27 起每季 0.98（+8%） | — | [Verified]／[Interested-party] | sh.dps*／sh.dpsQ.fy27q1 | 10-K、2026-09-15 新聞稿 |
| 回購 | 計畫內 16.719（36m 股）＋扣稅買回 5.6；授權餘額 40.6 | FY2026 | [Verified] | sh.* | 10-K 附註 15 |
| OpenAI 持股／營收／承諾 | 約 25%／24.1／13.0（已出資 11.9） | FY2026 | [Verified] | oai.* | 10-K 附註 1 |
| OpenAI Azure 增量承諾 | 250 | 2025-10-28 | [Interested-party] | oai.azureCommit | 10-Q FY26Q1 |
| Anthropic Azure 承諾／投資 | 30（up to 1 GW）／≤5 | 2025-11-18 | [Interested-party] | ant.* | [Anthropic](https://www.anthropic.com/news/microsoft-nvidia-anthropic-announce-strategic-partnerships) |
| 現價 | 529.76 | 2026-10-07 收盤 | [Verified] | val.price | StockAnalysis、Yahoo |

## 3. MW 估計比較

| 來源 | 數值 | 口徑 | 獨立性 |
|---|---|---|---|
| 公司：FY25 新增 | >2 GW | 不明（推定設施；含租賃） | 一手（利害關係） |
| 公司：FY26 Q2／Q3／Q4 新增 | 約 1／1／1 GW（Q1 未揭露） | 同上 | 一手 |
| 公司：FY26 AI 總產能增幅、兩年翻倍 | >80%；約 2× | 目標 | 一手 |
| Bloomberg（經 Reuters／Runtime Wire） | 總容量約 12 GW（2026-09）；2032 約 38 GW | 不明 | 獨立第三方（知情人士） |
| Morgan Stanley（經 Yahoo） | FY24 約 5 GW → FY28 約 20 GW | 不明 | 賣方 |
| Epoch AI：Fairwater Wisconsin／Atlanta | 369／636 MW-IT 在役（2026-10-05） | IT、AI | 獨立第三方 |
| Omdia（經 FT／DCD） | 2024 年 Hopper 485k 顆 ≈ 0.68 GW-IT | GPU 數 | 獨立第三方 |
| **本報告期初 AI 在役（自有＋融資租賃）** | **3.5／4.5／5.5 GW-IT** | IT、AI | [Derived]（三法交叉，見 mw.ai.open.est） |
| neocloud 租用（在役） | 0.8／1.1／1.5 GW-IT | IT、AI | [Derived]／[Assumed] |

三法：(1) 增量法：FY25＋FY26 新增約 5.5–6 GW（設施）× AI 70–90% ÷ PUE 1.2＋FY24 前 Hopper 約 0.9 GW-IT → 4.2–5.2；(2) 資本支出法：FY25–FY26 短期資產資本支出約 130 ÷ Tokenomics `IF_CapexIT`（GB200 24.0、GB300 37.4 $B/GW）→ 3.5–5.4；(3) 總量法：FY24 約 5 GW＋新增 ≈ 10.7 GW，與 Bloomberg 12 GW（2026-09）相容。對外／自用比例公司未揭露，基準 60%（50–75%，[Assumed]）。

## 4. k 的證據

| id | 世代 | 計價層 | 價格 | IF_GPUhrEcon／IF_HoldEcon | k |
|---|---|---|---|---|---|
| az.h100.od／r1／r3／r5／spot | H100（Hopper 欄） | 隨需／1 年／3 年／5 年／Spot | 12.29／7.87／5.40／4.92／2.27 $/GPU-hr | 1.605 | 7.66／4.90／3.36／3.06／1.42 |
| az.h200.od／r1／r3 | H200（Hopper 欄） | 隨需／1 年／3 年 | 10.60／5.82／5.28 | 1.605 | 6.60／3.62／3.29 |
| az.gb200.od／r1／r3 | GB200 | 隨需／1 年／3 年 | 27.04／17.31／11.90 | 2.086 | 12.96／8.29／5.70 |
| lc.iren | GB300 | 長約 5 年（Microsoft 買） | 9.70 $M/MW-年（9,700÷5÷200） | 12.72 | **0.76** |
| lc.nscale | GB300 | 長約（年期假設 5） | 11.20（23,000 二手 ÷ 5 ÷ 411 MW） | 12.72 | 0.88（4–6 年：0.73–1.10） |
| lc.nebius | GB300 | 長約 5 年 | ≤16.95（17,400÷5÷≥205 MW） | 12.72 | ≤1.33 |
| lc.anthropic | VR200 | 長約（年期未揭露） | 5.0–9.0（30,000 ÷ 4–6 年 ÷ 0.83–1 GW） | 12.76 | 0.39–0.71 |
| lc.openai | — | 長約 | 250 承諾，無年期、GW | — | 無法計算 |

Azure 牌價為公開 List（Azure Retail Prices API），大客戶以 EA／MACC 折扣成交，實際價格未揭露 → 只作上限。建議值見第 1 節第 5 點；與 nebius／coreweave（0.76／1.76）基準相同，長約下緣放寬至 0.45（Anthropic 隱含）。未用 Azure／Intelligent Cloud 營收、RPO 或 AI run-rate 反推 k。

## 5. AI／非 AI 雲端拆分試算（US$bn／年）

| | 低 | 基準 | 高 |
|---|---|---|---|
| 期初 AI 在役（自有）MW-IT | 3,500 | 4,500 | 5,500 |
| × 對外占比 | 50% | 60% | 75% |
| ＋ neocloud 租用 × 50% | 400 | 550 | 750 |
| ＝ 對外計費 MW-IT | 2,150 | 3,250 | 4,875 |
| 加權 IF_HoldEcon（Hopper 20%／GB200 45%／GB300 35%） | 10.60 | 10.60 | 10.60 |
| k | 0.76 | 1.06 | 1.31 |
| **對外 AI 雲端收入** | **17.3** | **36.5** | **67.7** |
| Azure 年化（FY26Q4 × 4） | 117.7 | 117.7 | 117.7 |
| **非 AI 殘差** | **100.3** | **81.2** | **50.0** |

3 × 3（MW × k）：基準 MW 下 k 低／高為 26.2／45.1；基準 k 下 MW 低／高為 24.2／54.8。驗證（只對照）：AI run-rate 37（FY26Q3，含 Copilot 約 11）→ 對外 AI 約 26–30；OpenAI FY26 營收 24.1（含分成）→ OpenAI Azure 消費約 19–21。**示警**：(1) 基準可能偏高 20–40%（若公司 run-rate 口徑可比）；(2) 上限檢查：每 MW 收入 11.24 ÷ 加權 `IF_RevGWFleet` 19.07 ＝ 59% > 50%（Hopper 的 RevGWFleet 4.0 偏低拉高）。非 AI 殘差 81.2 對 FY25Q4 Azure 年化 82.8（當時含 AI）→ 非 AI 年增約 10–25%，在合理區間。依規不回頭改 k。影子收入（自用 1.8 GW-IT × 10.60 × k＝1）約 19.1，只對照。

## 6. 租用算力（neocloud）與租賃

| 對手方 | 金額 | 年期 | 容量 | 來源 |
|---|---|---|---|---|
| Nebius | TCV 17.4（可擴至 19.4） | 5 年（2025-09 起） | >100,000 GB300（≈205 MW-IT） | Nebius 6-K；GPU 數 Bloomberg |
| IREN | 9.7（20% 預付） | 5 年 | 200 MW-IT（Childress, TX）、GB300 | IREN 新聞稿 |
| Nscale | 未揭露（二手約 23） | 多年期 | 約 200,000 GB300（≈411 MW-IT）；Texas 2026Q3 起 | Nscale 新聞稿 |
| Lambda | multibillion | 多年期 | tens of thousands（含 GB300 NVL72） | Lambda 新聞稿 |
| CoreWeave | 2025 年支付約 3.44（CRWV 營收 67%） | — | 推估 0.5–0.9 GW-IT | CoreWeave 10-K |
| 合計（報導） | >60（2025-11） | — | — | Bloomberg（經 Introl） |

租金屬營業成本，不計入資本支出；Microsoft 未揭露 neocloud 支出總額（可能含在採購承諾 194.1 內）。融資租賃：FY26 新增 24.6、負債 66.6、加權 13 年、4.5%；尚未起租 329.1（未拆營業／融資）。FY27 起更多資料中心租約將列營業租賃（不計入資本支出）。

## 7. 矛盾與處理

| # | 矛盾 | 處理 |
|---|---|---|
| 1 | Motley Fool 逐字稿摘要稱「FY27 資本支出約 175」；Microsoft IR 與 CFO Dive 為「曆年 2026 約 175」 | 以 IR（一手）為準：CY2026；FY27 只有「年增」 |
| 2 | OpenAI 持股：10-Q 27%（重組時、as-converted diluted）vs 10-K 約 25%（as-converted） | 取 10-K 25%（最新、經查核）；27% 列對照 |
| 3 | 公司口徑資本支出 ≠ 現金 PP&E＋融資租賃（Q3：30.9＋4.7＝35.6 vs 31.9） | 只取公司公布值；不自行加總 |
| 4 | 季度「折舊攤銷及其他」原始 Q1 值與後續季報重編不同（FY26Q1 13.061 vs 8.147） | 用重編後（Q2 年初至今減 Q2）；四季合計＝10-K |
| 5 | DCD 正文 Nebius 19.4 與網址 17.4 | Nebius 6-K：17.4，可擴至 19.4 |
| 6 | Introl 稱 IREN 站點在澳洲 | IREN 新聞稿：Childress, Texas |
| 7 | Nscale 標題 200,000 顆 vs 站點合計約 191,600 | 用 200,000（約數），站點明細列 note |
| 8 | MarketScreener FY29 淨負債 +99.6 與 FCF 89 方向不一致 | 不採用該欄 |
| 9 | Yahoo 與 S&P 的 FY27 營收高低值完全相同 | 視為獨立性存疑，不算強交叉 |
| 10 | 拆分基準 36.5 vs 公司 AI run-rate 隱含 26–30 | 只示警，不改 k（工作單） |

## 8. 資料缺口（「找不到」與「不存在」分開）

| 項目 | 狀態 | 試過的來源 |
|---|---|---|
| 累計 AI 在役 MW、對外／自用比例 | 不存在（公司只揭露增量） | 10-K、10-Q、四季法說 |
| FY26 Q1 新增 GW | 找不到 | Q1 法說、DCD |
| Maia 200 部署量 | 不存在（公司未揭露）；DIGITIMES 付費無數字 | 法說、DIGITIMES、TrendForce 搜尋 |
| Epoch AI Chip Owners 的 Microsoft 公司別 H100e | 找不到（互動圖表，shell 被擋） | epoch.ai 頁面、CSV（403） |
| OpenAI 250 承諾的年期與 GW；Anthropic 30 的年期 | 不存在 | 10-Q、OpenAI／Anthropic 公告 |
| OpenAI、Anthropic 帳列價值；Anthropic 持股比例 | 不存在 | 10-K 附註 4 |
| OpenAI 占 RPO（FY26Q4 最新比例） | 找不到（只有 FY26Q2 約 45%） | Q3、Q4 法說 |
| neocloud 支出總額、Nscale／Lambda 金額 | 不存在（公司未揭露）；Nscale 只有二手 | 10-K、對手方新聞稿 |
| Azure 大客戶實際成交 GPU 價格 | 不存在 | Azure 價目表只有牌價 |
| FY27 全年資本支出金額 | 不存在（只給方向） | Q4 法說 |
| 分部 D&A；未起租租賃逐年時程與營業／融資拆分 | 不存在 | 10-K 附註 13、18 |
| MSFT 個別長債殖利率 | 找不到（以 Moody's Aaa 指數類比） | FSMOne、搜尋 |
| 遊戲、PC 同業 NTM EV/EBITDA | 找不到（用 TTM） | StockAnalysis（只有 TTM） |
| Tokenomics NonNV（Maia 換 VR 等值）、PUE 具名範圍 | Tokenomics 未提供（快照標 missing） | Tokenomics v5.27 Interface／L1 |

## 9. 已套用的預設

| 問題 | 採用的預設 | 替代選項 | 對結果的影響方向 |
|---|---|---|---|
| 首期 | FY27 全年模型（1.0 年），期間 FY27–FY31 | FY27Q1 公布後改 Q1 實際＋0.75 年 | FY27Q1 10-28 公布，v0.1b′ 若在其後須滾動 |
| 模型分部口徑 | FY27 新分部（A&I、D&C）＋新產品別；舊三分部只作對照 | 沿用舊三分部 | 與 FY27 起季報一致；歷史可比（8-K 已重述 8 季） |
| 雲端拆分對象 | Azure（新定義，8-K 揭露絕對值） | Intelligent Cloud 或舊「Azure 及其他雲端」 | 新定義排除 GitHub、Security Copilot、醫療雲 → 基數較小、AI 占比較高 |
| 期初 AI MW | 4.5 GW-IT（3.5–5.5） | 只用公司增量 | 每 ±1 GW-IT 對外 AI 收入約 ±6.7 |
| 對外占比 | 60%（50–75%）[Assumed] | 50/50 | 每 ±10pt 約 ±5.1 |
| neocloud 用途 | 50/50（對照表 5.7） | 全對外／全自用 | ±6.2 |
| 世代組合 | Hopper 20／GB200 45／GB300 35 | GB300 占比更高 | GB300 每 +10pt 加權錨 +0.36 $M/MW |
| k | 長約 0.76、現貨 1.76、長約 70% → 1.06 | 長約 100%（0.76） | k 0.76 時對外 AI −10.3 |
| 資本支出口徑 | 公司口徑（含融資租賃）作首期／次期實際 | 現金口徑 | FY26：145.3 vs 115.9 |
| 稅前債務成本 | Moody's Aaa 6.26% | 30Y UST 5.67% | WACC 影響 <0.05pt（債務占比極低） |
| 企業軟體同業 | ORCL、SAP、CRM、ADBE、NOW 中位 10.8×（不含 MSFT） | oracle 總帳 5 家中位 15.02×（含 MSFT） | 非 AI 分部倍數較低 → 評價偏保守 |
| 遊戲／PC 倍數 | TTM 中位 12.63×／14.47×（EA 下市、RBLX 負 EBITDA 排除） | NTM（找不到） | 不確定 |
| OpenAI 持股 | 25% × 852 ＝ 213（v0.1b′ 再 × 0.8） | 27% | 27% 時 +17 |
| Anthropic 持股 | 不估值（比例未揭露） | 以 5 投入成本或 1.4% × 965 | 少計約 5–14 |
| 季度新聞稿數字標記 | [Interested-party]（年度以 10-K 為 [Verified]） | — | — |
| 舊資料檔 | 刪除 `permw_tokenomics_20261007`、`oracle_facts_20261007`；`consensus_orcl_20261007` 保留（`meta.consensusFile` 仍引用） | 一併刪除 | v0.1b′ 改 `meta.consensusFile` 後刪除 |

## 10. 給 v0.1b′ 的注意事項

1. **FY27Q1 財報 2026-10-28**：屆時以新分部報導；若 v0.1b′ 在其後執行，首期改「Q1 實際＋Q2–Q4 模型」（0.75 年），並更新尚未起租租賃、RPO、資本支出。
2. **租賃改分類**：FY27 起更多資料中心租約列營業租賃 → 不計入公司口徑資本支出、改為租金；329.1 未起租要拆成「起租後的租金或融資租賃」兩種處理，公司未揭露比例（需設 [Assumed]）。曆年 2026 指引 190 → 175 的差額即此效果。
3. **耐用年限 15→25 年**（建物）：機房折舊下降；伺服器 2–6 年不變（Tokenomics `IF_DeprLifeIT` 6 年）。
4. **產品線多於引擎 6 線**：Azure 以外 7 條非雲端線，建議合併 Industry＋Frontier、Windows＋XBOX 或依倍數群組（軟體／遊戲／PC）。
5. **neocloud 租金是營業成本**（對照表 5.7），MW 依用途分自用／對外；Nscale Texas 2026Q3 起、IREN 分階段到 2026 年底，需排起租時點。
6. **上限檢查 59%**：可能需把 Hopper 占比調低或接受示警；Tokenomics 缺 NonNV（Maia）名稱，暫不列 Maia。
7. `meta.consensusFile` 改指 `data/consensus_msft_20261008.json` 後刪除 `consensus_orcl_20261007.json`、`consensus_crwv_20260925.json`（後者也是 oracle 副本留下）；`oracle` 區段與 `oracle.files` 指向已刪檔，v0.1b′ 換成 MAG 引擎時一併移除。
8. OpenAI 營收分成至 2030、上限約 38（二手）：只列對照；OpenAI 模型（v0.6，未合併）Azure 路徑 2030 約 5.95 GW ≈ 71／年，可與「從 OpenAI 收到的雲端收入」並排。

## 11. verify.sh（2026-10-08，引擎未動；另 `tools/tokenomics/import_tokenomics.py --check` 通過：41 名、missing 3）

```
PASS  check_quarterly：季度加總＝年度、差異原因齊全
PASS  README 欄位說明與 company.json 一致
PASS  EBITDAR 校準：基準情境起點／穩態 EBITDA 率不變
PASS  建 HTML
PASS  建 Excel
PASS  重算：0 公式錯誤
PASS  反向 DCF：Excel 求解，rv_snap.json 與 Excel 一致
PASS  fix_outline
PASS  fix_datatable
PASS  verify_ooxml：OOXML OK
PASS  離線開啟：新建 HTML
PASS  離線開啟：dist/ 成品
PASS  cmp31 low：398 項全部 OK
PASS  cmp31 base：398 項全部 OK
PASS  cmp31 high：398 項全部 OK
PASS  cmp31 base_FY27：398 項全部 OK
PASS  test_quarterly：假設實際數與可移植性測試通過
PASS  test_rolling：日曆推算與滾動後第一屏
PASS  test_attrib：拆解工具（月數、同版 0、滾動一季與 WACC 12%）
verify.sh：全部通過（19 項）
```
