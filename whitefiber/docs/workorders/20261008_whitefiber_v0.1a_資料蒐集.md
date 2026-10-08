# 工作單 WhiteFiber v0.1a｜資料蒐集（2026-10-08，r0，chat 端）

- 必讀：`20261008_whitefiber_共同規則.md`、對照表 r1。參考：Oracle 同名工作單（`oracle/docs/workorders/20261007_oracle_v0.1a_資料蒐集.md`）與 `oracle/docs/reports/20261007_oracle_v0.1a_資料.md`（格式範本）。
- 分支與 PR：共用 `claude/whitefiber-v0.1` 與同一個 PR。
- **本張不改模型引擎**（seg*.js、tail.js、build_xlsx.py 等不動），`dist/` 不動，不升版。

## 目標
把 v0.1b 需要的所有 WhiteFiber 專屬事實、每 MW 推導與託管租金區間整理成可追溯的資料檔，寫入 `whitefiber/company.json` 新區段 `whitefiber`（引擎尚未讀取）；`verify.sh` 仍須全過。

## 產出
1. `whitefiber/data/whitefiber_facts_20261008.json`：事實總帳，每筆 `id`、`item`、`value`、`unit`、`period`、`source`、`url`、`docDate`、`retrieved`、`tag`、`note`。
2. `whitefiber/data/consensus_wyfi_20261008.json`：市場共識，格式比照 `whitefiber/data/consensus_orcl_20261007.json`；日曆財年。
3. `whitefiber/data/permw_tokenomics_20261008.json`：自 `permw_tokenomics_20261007.json` 複製；確認 Tokenomics master `model/CURRENT` 是否仍為 v5.24（若已更新，重算三情境並列差異，否則 `_readme` 註明「沿用 Nebius v0.1a 推導，跨模型一致」）；另加 `whitefiber` 對照區段（步驟 5）。
4. `whitefiber/company.json` 新增頂層區段 `whitefiber`：v0.1b 會用到的欄位草稿，每欄以 `ref` 指回總帳 `id`，附 `mapTo`（應搬到的既有欄位路徑或「新欄位」）。`scripts/fields_doc.py --check` 失敗時補 `whitefiber` 子區說明後 `--write`。
5. `whitefiber/docs/reports/20261008_whitefiber_v0.1a_資料.md`：資料報告（結論先行；核心數字表；矛盾與處理；每 MW 與託管租金對照；資料缺口；verify 輸出；總帳表格版）。
6. 進度檔新增「v0.1a」段落並持續覆寫。

## 步驟（每步完成即 commit＋push＋更新進度檔）
0. 跑基線 `whitefiber/scripts/verify.sh`（預期全過；Oracle v0.2 基線 19 項）、建立進度檔（記下 PR 編號）。若基線失敗，先判斷是否因複製遺漏（Oracle v0.1a 曾漏 data 檔），補齊後再跑。
1. **SEC 一手文件**：FY2025 10-K、2026 Q1 與 Q2 10-Q、各季財報新聞稿 8-K、IPO S-1／424B（2025-08）、2026-08 可轉債 8-K 與發行說明；XBRL companyfacts。取：
   - 損益：上市以來各季與 FY2024–FY2025 年度：營收分線（雲端、託管、其他）、毛利（不含 D&A）、營業費用（G&A、SBC）、D&A、減損、利息（第三方、關係人）、稅、淨損；調整後 EBITDA（公司定義）；**一次性項目**（Q2 客戶終止 $12.3M 收入、相關 $4M 成本與 $2.2M 呆帳）分開列。
   - 現金流量：營運現金流、資本支出（PP&E、GPU、建物分開若有）、GPU 租賃或設備融資的現金流。
   - 資產負債：現金（含受限）、債務明細（RBC 聯貸：額度、已動用、利率、到期、契約條件；Bit Digital 延遲提款貸款；設備與過橋融資約 $83.2M；可轉債 $310M：利率、到期、轉換價、上限買權、可贖回條件）、**租賃**（GPU 租賃、站點土地與建物租賃：到期表、未起租承諾）、PP&E 依類別（建物、電力設備、GPU）、**遞延營收／客戶預付**（約 $143M：託管與雲端分開若有）、RPO（託管約 $932.9M 與雲端，及認列時程）。
   - 股數：基本、稀釋、季末流通；Bit Digital 持股比例與表決權（雙重股權若有）；可轉債轉換股數。
   - 指引或目標（全部 [Interested-party]）：雲端合約全部到位後年化 > $200M、NC-1 時程、任何 2026／2027 營收或 EBITDA 目標。
2. **站點與 MW**：逐站點（NC-1、MTL-1、MTL-2、MTL-3、冰島、巴黎、安大略第三方機房、候選 60 MW 站點、Krambu 100 MW）記錄：**毛電力（gross／設施）vs IT 負載**、已通電 vs 已計費 vs 合約 vs 規劃、自有 vs 承租、客戶、時程。逐筆註明口徑，不明者標「口徑不明」。NC-1：Duke 99 MW 協議、下一批 45 MW 毛容量、長期約 300 MW；PUE（若有揭露，否則以同業或設計值 [Analogy]）。
3. **合約**：NC-1 Nscale 託管約（$865M、10 年、40 MW IT 兩期、年調 3%、轉嫁項目、起租日、預付或保證金）；雲端合約逐筆（Baseten、Prime Intellect、冰島、巴黎、Cerebras〔MTL-3〕、終止的舊客戶與替代合約）：TCV、期間、GPU 型號與數量、站點、起始日、是否預付；換算每顆 GPU 每小時價格與**每 MW-IT 年營收**（[Derived]：GPU 數 × 每顆 IT kW，kW 取 Tokenomics 或 NVIDIA 規格並註明）。**合約隱含單價只作對照**。客戶集中度（10-Q 前幾大客戶占比）。對手方：Nscale 的融資與主要合約公開資訊（只作文字對照）。
4. **託管租金區間** [Analogy]：可觀察的批發／HPC 託管合約每 MW-IT 年租金至少 4 筆（例如 Applied Digital–CoreWeave、Galaxy Helios–CoreWeave、IREN、Core Scientific–CoreWeave、Cipher–AWS／Fluidstack 等；逐筆來源、期間、是否含轉嫁）；每 MW 建置成本（建物＋電力＋冷卻，不含 GPU）至少 3 個來源（同業揭露、WhiteFiber 自己的 NC-1 資本支出若有）；建物折舊年限（10-K 會計政策）。
5. **每 MW（雲端）**：沿用產出 3；另列 WhiteFiber 對照：(a) Q2 經常性雲端營收（扣一次性）年化 ÷ 可辨識的在役雲端 MW（[Derived]，只作對照）；(b) 各雲端合約隱含每 MW 年營收 vs 11.62／17.40／24.20；(c) 毛利率 59%（含一次性）與扣除一次性後毛利率 vs 可觀察 neocloud EBITDA 率（IREN 約 35%、CRWV 約 59%，取自既有資料檔並註明）。
6. **評價輸入**：beta（至少兩個來源，註明期間長度）、10 年期美債殖利率（日期）、債務成本（可轉債有效利率、Bit Digital 貸款 9.5%、RBC 利率）、託管／資料中心同業 NTM EV/EBITDA（DLR、EQIX、APLD、IREN、CORZ、CIFR、GLXY 中至少 4 家，逐筆來源）、neocloud 同業（CRWV、NBIS；可沿用既有資料並更新日期）。
7. **市場共識**：共識目標價（平均、中位、高低、家數）、FY2026–FY2028 營收與 EBITDA／EPS 預估、評等分布；現價（最新收盤，標日期）；市值。至少兩個呈現並寫明底層來源。家數少照實記錄。
8. **寫入 company.json 的 `whitefiber` 區段**、跑 `verify.sh`（全過）、資料報告、PR 回報。

## 驗收
- 總帳每筆有網址與標記；chat 端抽查 10 筆，須與原文件一致。
- 核心欄位齊全：2026 Q2 與 FY2025 各分線營收、毛利、調整後 EBITDA、D&A、營運現金流、資本支出、現金、債務明細（含可轉債條款）、租賃、股數與 Bit Digital 持股、RPO、遞延營收、逐站點 MW、逐筆合約；缺漏列資料缺口並寫試過的來源。
- 雲端每 MW 三情境沒有以公司目標或合約隱含單價作為輸入。
- `verify.sh` 全過（引擎未動，數字應與基線相同）。
- 進度檔「下一步」寫好 v0.1b 的起點。

## 停止條件
依共同規則第 6 節。補充：若 2026 Q2 10-Q 的主要數字（營收、現金、債務、遞延營收）無法從任何一手來源取得，屬第 2 條。
