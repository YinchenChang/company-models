# 工作單 CoreWeave W2｜每 MW 資本支出、成本、收入改寫（2026-10-07，r0，chat 端）

- 前置：W1 已合併（快照、`TK_` 具名範圍、對照表、`data/permw_inputs_20261007.json`）。
- 必讀：共同規則、W1 PR 回報與進度檔、對照表、`coreweave/README.md`（引擎結構）。
- 分支：`claude/coreweave-w2-permw`；PR 標題「CoreWeave W2：每 MW 經濟性改接 Tokenomics」。
- 本張改引擎與 `company.json`，**不改 `dist/`、不升版**（v4.6 成品在 W3）。

## 目標
每 MW 的資本支出、營運成本、收入改由「Tokenomics 物理與產業數據 × 公司專屬輸入」正向推導；舊方法保留為可切換的對照（W3 前後對照與變動拆解要用）。

## 實作原則（已套用的預設）
1. **雙引擎同步**：5b（只讀檢視器）未做，新邏輯 Excel 公式與 HTML 引擎（segA–E）兩邊都實作，以 Excel 為準；cmp31 新增比對項涵蓋每個新區塊。
2. **方法開關**（`company.json` → `methodology.perMw`，皆為 `company` 可設定）：`capex`：`tokenomics`｜`legacy`；`cost`：`bottomUp`｜`ebitdaPct`；`revenue`：`gpuHr`｜`legacy`。預設全為新方法；`legacy` 必須精確重現 v4.5 數字（verify 以 legacy 組合跑 `--vs-dist` 須 0 差異，作為不破壞舊邏輯的證明）。
3. 公司專屬內容放 `company.json`；新欄位同步 `fields_doc.py --write`；首期一次性金額或期初餘額加入 `calendar_q.py` 的 `ROLL_FIELDS`；改名用 `rename_ids.py`；遵守 CLAUDE.md「勿改」8 條。
4. Tokenomics 值一律經 W1 的「Tokenomics_取數」分頁具名範圍引用，不得在公式或 JS 寫死數字；JS 端讀同一份快照（建置時內嵌）。
5. 世代主值取「基準」成本情境；低／高成本只進敏感度表。

## 步驟（每步：verify 全過 → 進度檔 → commit → push；verify 未過不得進下一步）
0. 開分支、draft PR、進度檔新增「W2」段落。若 Tokenomics `master` 已有 v5.25（W1 第 3 步的新名稱），重抓快照並更新 `tokenomics` 區段；若仍沒有，用到 v5.25 名稱的成本項以 W1 對照表的 CRWV 現值暫代並標「待 v5.25」，進度檔列入未解問題（不停工）。

1. **世代組合（新，公司專屬）**：`fleet.generations`（與 Tokenomics 世代同名）、`fleet.openMix`（6/30 在役 MW 依世代占比，取 W1 第 6b 步）、`fleet.newMix`（各期新增 MW 的世代占比）。`newMix` 預設：2H26 GB300 80%／GB200 20%；FY27 GB300 50%／VR200 50%；FY28–FY30 VR200 100%（Rubin Ultra 只作敏感度：FY29–FY30 改 100% Rubin Ultra）。推導各期期末與平均在役 MW 的世代結構（汰換依 `gpuLife`；模型期內 6 年壽命不會汰換，但公式要通用）。Excel 放「輸入與假設」B 區之後的新子區「世代組合」；HTML 次層顯示。

2. **每 MW 資本支出**（`capex=tokenomics`）：各期每 MW 建置成本＝Σ 新增 MW 世代占比 × `IF_CapexIT`（基準，換成 US$m/MW）。CRWV 機房以租約取得，**不含廠房**（廠房成本體現在租金）。`customerFund`、`capexScale`、`lambda` 機制不變。舊值 34 留為對照列。`methodology.checks.capexPerMwBand` 檢查照常。

3. **折舊年限**：`gpuLife` 改為引用 `IF_DeprLifeIT`（v5.25；基準 6 年，與現值相同，數字應不變）。

4. **每 MW 營運成本，由下而上**（`cost=bottomUp`）
   以「租金前」口徑計算（模型的租金與利息已在支出頁另列；EBITDAR＝EBITDA＋租金，見 build_xlsx 說明文字）：
   - 電費＝平均在役 IT MW × `IF_PowerCost`（每 IT GW 年電費，已含平均用電比與 PUE）÷ 1000。
   - IT 維護＝平均在役 IT MW × Σ 在役世代占比 × `IF_MaintIT`。
   - 人員、軟體、水與耗材＝MW × `IF_StaffSW`；財產稅與保險＝MW × `IF_TaxIns` × 「IT 資本占總資本比」（只算 CRWV 擁有的 IT 部分；廠房屬房東）。
   - 廠房維護與廠房折舊**不計**（房東負擔，含在租金）。
   - 公司管銷與其他（銷售行銷＋一般管理，扣除 D&A 與 SBC）＝營收 × 管銷率；管銷率取 W1 第 6d 步最近一季實際（報告列出計算式），各期持平 [Derived]。
   - 推導：EBITDAR 率＝1 −（以上合計 ÷ 營收）；EBITDA 率＝EBITDAR 率 − 租金 ÷ 營收。取代原 `ebStart`／`ebSteady` 的線性路徑；原路徑保留為 `ebitdaPct` 對照。
   - 對照列（不強制平衡）：Q2 實際（營收成本＋技術與基礎設施，扣 D&A、SBC、租金）÷ 平均在役 MW，與模型由下而上的每 MW 現金成本並列，差距與可能原因寫入報告（例如爬坡期閒置產能、未揭露的成本項）。
   - 原「電力／維護 overlay」區：bottomUp 模式下自動停用（避免重複扣），legacy 模式維持原行為。

5. **每 MW 收入**（`revenue=gpuHr`）
   - 每 MW 年收入＝Σ 平均在役世代占比 × 每 MW GPU 數（`IF_GPUsPerGW` ÷ 1000）× 該世代 GPU 小時合約價 × 8,760 × 計費利用率（沿用 `m.util`）÷ 10⁹（US$bn/MW）。
   - GPU 小時價格放 `company.json` → `pricing.gpuHr`（每世代：基準與低／高、來源、日期、標記），取 W1 第 6c 步。合約價在合約期內固定（續約價格衰退不做）。
   - **不得**以公司 ARR、營收指引或 RPO 反推價格。舊 `m.revMW` 與「期末 ARR ÷ MW」留為對照列。
   - 另列兩條 Tokenomics 參考線（只作對照，不入損益）：每 MW 經濟持有成本（`IF_HoldEcon` 依世代加權，＝「不賠錢」下限，但含廠房資本回收，報告寫明與 CRWV 租金口徑的差異）、每 GPU 小時經濟持有成本（`IF_GPUhrEcon`）對合約價的倍數。
   - **備案**：若 W1 回報 GPU 小時價格不足（任一主要世代少於兩個獨立來源），改為 `revenue=legacy`（保留 `m.revMW` 為輸入），但仍新增上述兩條參考線與「隱含每 GPU 小時價格」＝revMW ÷（每 MW GPU 數 × 8,760 × 利用率）的反算對照列（這是對照，不是輸入）；報告寫明採用備案。

6. **MW 口徑**：依 W1 第 6a 步結論設定 `meta.mwBasis`（`IT`｜`facility`）。若為設施口徑，所有 Tokenomics 每 MW 值在引用處乘換算係數（IT MW ＝ 設施 MW ÷ `IF_FacilityGW`），並在 Excel 與 HTML 的每 MW 表頭標明口徑。

7. **每 MW 經濟性彙總表（新，Excel＋HTML 次層，預設收合）**：各期一欄，列：每 MW 年收入、電費、IT 維護、人員軟體、稅險、管銷、租金、現金成本合計、EBITDA、D&A（＝每 MW 資本支出 ÷ 壽命，在役加權）、利息（依模型債務 ÷ 平均在役 MW）、稅前；另並列「IT MW」與「設施 MW（× `IF_FacilityGW`）」兩種口徑。W3 的對照報告直接讀這張表。

8. **敏感度**：Tokenomics 成本情境（低／高成本）、GPU 小時價格低／高、世代組合（Rubin Ultra 版）三組，對三情境目標價與融資缺口的影響（建置時快照，沿用 CRWV 敏感度快照機制，附「快照已過期」檢查）。

9. **整體 verify 與回報**：`verify.sh` 全過；另以 legacy 組合跑 `--vs-dist` 0 差異。PR 回報列：三情境目標價與區間（新 vs v4.5）、FY27／FY30 每 MW 收入與每 MW 成本（新 vs v4.5）、已套用的預設、資料缺口、是否採用收入備案。

## 驗收（chat 端逐項核）
- 新方法下：資本支出、成本、收入的公式都可追溯到 `TK_` 名稱與 `company.json` 公司專屬欄位；Excel 與 JS 皆無寫死的 Tokenomics 數字。
- legacy 組合 0 差異；新方法 cmp31 三情境＋FY27 錨定全部 OK。
- bottomUp 模式下 overlay 不重複扣費；租金只扣一次（報告附一行算式證明）。
- 未以公司 ARR／CapEx 指引／RPO 反推任何參數。
