# CoreWeave v4.7：每 MW 收入以 Tokenomics 為錨（v4.6 → v4.7；W4 r2）

2026-10-08｜工作單 W4 r2（`docs/workorders/20261008_coreweave_W4_收入Tokenomics錨定.md`）｜PR #27｜Excel 版：`docs/reports/20261008_coreweave_v4.7_收入錨定.xlsx`

## 結論
1. **CRWV 每 MW 收入（100% 計費時數）是 Tokenomics 打平租金（`IF_HoldEcon`，WACC 10%）的 0.76 倍（各期）**：k 取唯一已證實的可比長約（IREN–Microsoft GB300 五年約）；CRWV 先簽多年期合約再建產能，新增產能也按長約價，隨需占比 0%。
2. **這代表**：以市場長約價計，CRWV 收不回含廠房資本回收的完整持有成本——每 MW 年收入 8.8–9.4 對錨 11.6–12.3，每 MW 稅前各期約 −4 至 −5；擴張愈快、虧損與融資需求愈大。
3. **加權目標價：保守 $43.99 → $13.98、基準 $29.68 → $5.25、積極 $19.95 → $0.00，三者皆賣出**；基準與積極 DCF 失效，結論全部來自 EV/EBITDA 腿。k_長約 改用三筆長約中位數 0.89 時基準 $25.46。

## 需要 Andy 做的事
1. **決定 k_長約 基準**：0.76（唯一已證實的可比長約，目前採用）或 0.89（三筆長約中位數，含未證實的 Oracle–OpenAI）。基準目標價 $5.25 對 $25.46。建議：維持 0.76，待第二家供應商的已證實長約。
2. 合併前以真正的 Excel 開啟 `dist/20261008_CoreWeave收支模型_v4_7.xlsx` 檢查（成品已設「開檔時全部重算」；運算表與活公式只以 LibreOffice 驗證過）。

## r2 修訂（chat 端審查後，2026-10-08）
- **隨需占比取代長約占比**：r1 以「RPO 涵蓋的產能 ÷ 在役產能」為長約占比，把 RPO 未覆蓋的新增產能當成現貨（k 1.76），基準一度為 $98.02；這是工作單 r1 的設計錯誤。r2 改為 k＝隨需占比 × k_現貨＋（1 − 隨需占比）× k_長約，隨需占比 0%（敏感度 10%／20%，[Assumed]），RPO 覆蓋率只作對照列。
- **k 隨 Tokenomics 成本情境重算**：證據是市場價格（事實），k＝價格 ÷ 同成本情境的持有成本；輸入值為基準情境，其他情境以基準證據世代（GB300 的 IF_HoldEcon、H100 的 IF_GPUhrEcon）的成本比例換算。結果：Tokenomics 低／高成本敏感度下收入大致不變（基準 FY30 每 MW 10.1／8.5，對基準 9.4；差異來自在役世代組合與證據世代不同）、成本隨情境變，**高成本使目標價下降**（基準 $0.00）。

## 1. 關鍵數字（v4.6 → v4.7）
| 情境 | 加權目標價 v4.6 | v4.7 | 評等 v4.6 → v4.7 | 融資缺口 v4.6 → v4.7（US$bn） | v4.7 EV/EBITDA 腿單獨 | v4.7 DCF 腿 |
|---|---|---|---|---|---|---|
| 保守 4.2 GW | $43.99 | $13.98 | 賣出 → 賣出 | 32.7 → 54.8 | $25.42 | $0（0 截斷） |
| 基準 5.6 GW | $29.68 | $5.25 | 賣出 → 賣出 | 81.6 → 104.8 | $5.25 | 失效 |
| 積極 8 GW | $19.95 | $0.00 | 賣出 → 賣出 | 157.9 → 183.5 | $0.00 | 失效 |

讀法：目標價＝DCF 腿 45%＋EV/EBITDA 腿 55%；DCF 失效時權重歸零、目標價＝EV/EBITDA 腿。v4.7 基準與積極 DCF 失效（常態化 FCF ≤ 0），保守 DCF 股權為負被 0 截斷，所以**三情境結論都由 EV/EBITDA 腿決定**。

## 2. 每 MW 年收入（US$m／MW／年）
### 2a. 100% 計費時數（B 區「每 MW 年收入」）：v4.6 輸入 vs v4.7 錨 × k
| 情境 | 版本 | FY26（2H） | FY27 | FY28 | FY29 | FY30 |
|---|---|---|---|---|---|---|
| 保守 4.2 GW | v4.6 | 11.20 | 11.50 | 11.50 | 11.00 | 10.50 |
| 保守 4.2 GW | v4.7 | 8.42 | 8.79 | 9.04 | 9.16 | 9.27 |
| 基準 5.6 GW | v4.6 | 11.20 | 11.50 | 11.50 | 11.00 | 10.50 |
| 基準 5.6 GW | v4.7 | 8.42 | 8.79 | 9.06 | 9.21 | 9.35 |
| 積極 8 GW | v4.6 | 11.20 | 11.50 | 11.50 | 11.00 | 10.50 |
| 積極 8 GW | v4.7 | 8.42 | 8.79 | 9.11 | 9.30 | 9.45 |

### 2b. 每平均在役 MW（含服務、含爬坡分母與利用率；彙總表）
| 情境 | 版本 | FY26（2H） | FY27 | FY28 | FY29 | FY30 |
|---|---|---|---|---|---|---|
| 保守 4.2 GW | v4.6 | 10.05 | 10.56 | 10.37 | 9.78 | 8.99 |
| 保守 4.2 GW | v4.7 | 7.64 | 8.16 | 8.42 | 8.41 | 8.14 |
| 基準 5.6 GW | v4.6 | 10.05 | 10.56 | 10.33 | 9.63 | 8.65 |
| 基準 5.6 GW | v4.7 | 7.64 | 8.16 | 8.40 | 8.30 | 7.86 |
| 積極 8 GW | v4.6 | 10.05 | 10.56 | 10.23 | 9.34 | 8.27 |
| 積極 8 GW | v4.7 | 7.64 | 8.16 | 8.34 | 8.09 | 7.54 |

### 2c. 錨與 k 的逐年值（基準；三情境見 Excel「錨與k」）
| 項目 | FY26（2H） | FY27 | FY28 | FY29 | FY30 |
|---|---|---|---|---|---|
| 錨（Σ 在役占比 × IF_HoldEcon） | 11.074 | 11.562 | 11.922 | 12.115 | 12.307 |
| k_長約（目前成本情境） | 0.760 | 0.760 | 0.760 | 0.760 | 0.760 |
| 隨需占比 | 0.000 | 0.000 | 0.000 | 0.000 | 0.000 |
| 定價倍數 k | 0.760 | 0.760 | 0.760 | 0.760 | 0.760 |
| 每 MW 年收入＝錨 × k | 8.416 | 8.787 | 9.061 | 9.207 | 9.353 |
| 相對 v4.6 舊輸入 | -0.249 | -0.236 | -0.212 | -0.163 | -0.109 |
| 對照：RPO 覆蓋率（不驅動） | 2.164 | 1.374 | 0.910 | 0.712 | 0.422 |
| 上限：客戶每 MW 付費 token 營收（IF_RevGWFleet） | 19.10 | 24.17 | 29.63 | 33.82 | 37.66 |
| 上限：CRWV 計費收入 ÷ 客戶 token 營收 | 0.419 | 0.345 | 0.287 | 0.253 | 0.228 |

**上限檢查**：三情境各期最高 42%（FY26），皆低於門檻 50%，檢查頁「通過」。

**營運成本**（電費、IT 維護、人員軟體、稅險、租金）與 v4.6 **0 差異**（升版驗收清單不含這些列）；管銷＝營收 × 5.51%，隨營收改變。

## 3. 定價倍數 k 的證據表（倍數＝價格 ÷ Tokenomics v5.27 基準同世代持有成本）
| 證據 | 世代 | 價格 | Tokenomics | 倍數 | 用途 | 合約／期間 | 來源（文件日期） | 標記 |
|---|---|---|---|---|---|---|---|---|
| IREN–Microsoft GB300 五年約 | GB300 NVL72 | 9.7 US$m/MW/年 | IF_HoldEcon 12.725 | 0.76 | long | 長約（hyperscaler）｜5 年 | IREN Secures $9.7bn AI Cloud Contract with Microsoft（IREN 新聞稿）（2025-11-03） | [Derived] |
| IREN–NVIDIA 氣冷 Blackwell 五年約（GB300 類比） | GB300 NVL72 | 11.333 US$m/MW/年 | IF_HoldEcon 12.725 | 0.89 | range | 長約（晶片商客戶）｜5 年 | IREN Secures $3.4bn AI Cloud Contract with NVIDIA（IREN 新聞稿）（2026-05-07） | [Analogy] |
| Oracle–OpenAI 約 $300bn／4.5 GW（未經證實） | GB300 NVL72 | 13.333 US$m/MW/年 | IF_HoldEcon 12.725 | 1.05 | list | 長約（hyperscaler 對實驗室）｜約 5 年（2027 起） | The Register 轉述 WSJ（2025-09-11）：OpenAI 五年支付 Oracle $300bn；4.5 GW 為 2025-07 Stargate 擴充（2025-09-11） | [Analogy] |
| H100 Silicon Data 現貨指數 | Hopper H100 | 2.82 US$/GPU-hr | IF_GPUhrEcon 1.605 | 1.76 | spot | 現貨（市場租金指數）｜隨需／短期 | Silicon Data H100 Index 頁面（2026-10-06） | [Verified] |
| B300 Silicon Data 現貨指數（GB300 類比） | GB300 NVL72 | 6.82 US$/GPU-hr | IF_GPUhrEcon 2.983 | 2.29 | range | 現貨（市場租金指數）｜隨需／短期 | Silicon Data H100 Index 頁面（含 B200、B300 指數）（2026-10-06） | [Analogy] |
| B200 Ornn 成交指數（GB200 類比） | GB200 NVL72 | 4.08 US$/GPU-hr | IF_GPUhrEcon 2.086 | 1.96 | range | 現貨（市場成交指數）｜隨需／短期 | Ornn（經 WSJ 轉述；W1 第 6c 步）（2026-04-13） | [Analogy] |
| B200 Silicon Data 現貨指數（GB200 類比） | GB200 NVL72 | 5.87 US$/GPU-hr | IF_GPUhrEcon 2.086 | 2.81 | list | 現貨（市場租金指數）｜隨需／短期 | Silicon Data H100 Index 頁面（含 B200、B300 指數）（2026-10-06） | [Analogy] |
| B300 Nebius 隨需牌價（GB300 類比） | GB300 NVL72 | 7.85 US$/GPU-hr | IF_GPUhrEcon 2.983 | 2.63 | list | 隨需牌價｜隨需 | Nebius NVIDIA GPU Pricing（官網價目）（2026-10-07） | [Interested-party] |
| GB200 CoreWeave 隨需牌價 | GB200 NVL72 | 10.5 US$/GPU-hr | IF_GPUhrEcon 2.086 | 5.03 | list | 隨需牌價｜隨需 | CoreWeave Pricing（官網價目）（2026-10-07） | [Interested-party] |
| H100 CoreWeave 隨需牌價 | Hopper H100 | 6.155 US$/GPU-hr | IF_GPUhrEcon 1.605 | 3.83 | list | 隨需牌價｜隨需 | CoreWeave Pricing（官網價目）（2026-10-07） | [Interested-party] |

- **最終值**：k_長約 0.76（區間 0.7–1.0；**三筆長約中位數 0.89**＝IREN–Microsoft 0.76、IREN–NVIDIA 0.89、Oracle–OpenAI 1.05〔未證實〕，列為敏感度）、k_現貨 1.76（1.5–2.3），皆 [Analogy]；隨需占比 0%（10%／20% 敏感度）[Assumed]。沒有一個 k 值或隨需占比來自 CRWV 的營收、ARR、RPO 金額或 Q2 隱含 k。
- 用途：long＝k_長約基準、spot＝k_現貨基準、range＝支持區間、list＝只列不用。
- **CRWV 合約組合**：已承諾合約加權平均年期約 5 年（10-K FY2025 [Verified]）；「絕大多數收入來自多年期已承諾合約」（S-1，未給百分比 [Interested-party]）；長約／隨需比例：**找不到**（試過 10-K FY2025、S-1、Q2 新聞稿與網路搜尋）。
- **找不到**：VR200 NVL72 長約或多年期合約的每 GPU 小時／每 MW 價格：Nebius–Meta（2026-03-16，$12bn 專屬產能、Vera Rubin、五年）與 Nebius–Microsoft（2025-09-08，$17.4–19.4bn、五年）都未揭露 MW 或 GPU 數，無法換算（找不到）
- **找不到**：GB200 NVL72 長約價：找不到可換算的揭露（Nscale–Microsoft 約 20 萬顆 GB300 未揭露金額，FT 估約 $14bn 屬估計，不用）
- **找不到**：其他 neocloud 對 hyperscaler／實驗室的長約（獨立於 IREN 的供應商）：上列 Nebius、Nscale 皆缺 MW 或金額；Oracle–OpenAI 金額未經證實、只列
- 未使用 SemiAnalysis／InferenceX。

## 4. Q2 驗證（不是校準；`scripts/q2_check_w4.py`）
模型首期（FY26 下半年）每平均在役 MW 年收入 7.638 對 Q2 年化 8.240（營收 2.575 × 4 ÷ 平均在役 1,250 MW）：差距 -0.602（-7.3%）。

| 項目 | 模型首期 | Q2 | 貢獻 |
|---|---|---|---|
| (i) 爬坡分母（計費 MW ÷ 在役 MW） | 0.910 | 0.900 | +0.091 |
| (ii) 利用率 | 0.950 | 0.950 | +0.000 |
| (iii) 定價倍數 k（模型 k − Q2 隱含 k） | 0.760 | 0.838 | -0.737 |
| (iv) 世代組合（錨：模型首期平均 vs Q2 季末） | 11.074 | 10.964 | +0.072 |
| (v) 其他（非算力服務等） | 0.358 | 0.386 | -0.028 |
| **合計＝總差距** | 7.638 | 8.240 | **-0.602** |

- Q2 假設：計費比例＝季末 Billable 1,350 ÷ 季末在役 1,500＝90% [Assumed]；Q2 利用率＝模型 95%（未揭露）[Assumed]；服務收入占比同模型 [Assumed]；Q2 錨＝季末世代組合 × IF_HoldEcon。Q2 隱含 k：未調整 0.752、調整後 0.838。
- 差距 7.3% 超過 5%：**類型＝觀點**。模型 k 0.76 取市場長約證據；Q2 營收含較早簽約的舊世代合約與隨需，隱含 k 約 0.84。依已決定事項 14，不回頭改 k 或隨需占比去貼近 Q2。

## 5. 敏感度（建置時快照 `permw_sens.json`；加權目標價 US$／融資缺口 US$bn）
| 設定 | 保守 | 基準 | 積極 |
|---|---|---|---|
| 基準（目前輸入） | $13.98／54.8 | $5.25／104.8 | $0.00／183.5 |
| Tokenomics 低成本 | $64.67／22.3 | $39.99／63.3 | $26.61／125.7 |
| Tokenomics 高成本 | $0.00／107.0 | $0.00／167.5 | $0.00／265.8 |
| GPU 小時價格 低 | 不適用 | 不適用 | 不適用 |
| GPU 小時價格 高 | 不適用 | 不適用 | 不適用 |
| 世代組合 Rubin Ultra 版 | $14.20／60.9 | $0.00／123.8 | $0.00／219.4 |
| 管銷率 GAAP（含 SBC） | $15.38／59.4 | $0.00／109.7 | $0.00／189.3 |
| k_長約 0.70 | $8.99／62.6 | $0.00／113.2 | $0.00／193.6 |
| k_長約 0.89（三筆長約中位數） | $35.49／38.2 | $25.46／86.5 | $6.54／161.5 |
| k_長約 1.00 | $74.15／24.1 | $45.14／71.0 | $28.77／143.0 |
| 隨需占比 10%（k_現貨 1.76） | $30.49／42.0 | $20.16／90.7 | $0.98／166.6 |
| 隨需占比 20%（k_現貨 1.76） | $55.52／29.2 | $36.59／76.6 | $20.30／149.7 |

讀法：(1) k_長約 是最大槓桿（基準 0.70／0.89／1.00：$0.00／$25.46／$45.14）；(2) Tokenomics 低／高成本：k 依同情境成本重算，收入大致不變、成本變，**低成本 → 目標價高、高成本 → 低**（基準 $39.99／$0.00）；(3) 隨需占比 10%／20%：$20.16／$36.59；(4) Rubin Ultra 版 $0.00（每 MW 建置成本較高）。

## 6. 目標價變動拆解（已決定事項 12；0 截斷口徑，US$／股；`scripts/attrib_w4.py`）
(a) 時間推移 0（評價日 2026-06-30 不變）、(b) 實際數更新 0、(c) 假設變更 0（Tokenomics v5.26 → v5.27：引用的 25 個名稱數值與位置相同；新增 IF_RevGWFleet 只用於上限檢查）；全部屬 (d) 方法變更，依序：

| (d) 逐項 | 保守 | 基準 | 積極 |
|---|---|---|---|
| ① 錨取代舊輸入（k＝1：每 MW 年收入＝在役世代 IF_HoldEcon） | +30.16 | +15.46 | +8.82 |
| ② 套用定價倍數 k（隨需占比 0%：k＝k_長約 0.76） | -60.17 | -39.89 | -28.77 |
| ③ 上限檢查（只新增檢查列） | +0.00 | +0.00 | +0.00 |
| 合計（＝總變動） | -30.01 | -24.43 | -19.95 |
| 融資缺口變動（US$bn） | +22.2 | +23.2 | +25.5 |

檢查：三情境依序各步相加＝v4.7 − v4.6（誤差 < 0.01 美元／股）；③＝0。① k＝1 時每 MW 年收入＝錨（11.1–12.3，略高於 v4.6 輸入 10.5–11.5）；② 套用 k＝0.76 後每 MW 年收入降為 8.4–9.4。

## 7. 已套用的預設（問題｜採用的預設｜替代選項｜對結果的影響）
| 問題 | 預設 | 替代 | 影響 |
|---|---|---|---|
| k_長約 基準 | 0.76（唯一已證實的可比長約；r2 維持） | 0.89（三筆長約中位數） | 基準 $5.25 → $25.46；**待 Andy 決定** |
| 隨需占比 | 0%（公司未揭露；S-1「絕大多數為多年期已承諾合約」）[Assumed] | 10%／20% | 基準 $20.16／$36.59 |
| k 隨成本情境重算的換算 | 輸入 × 基準證據世代成本（基準）÷ 同名稱目前情境（k_長約：IF_HoldEcon｜GB300；k_現貨：IF_GPUhrEcon｜H100） | 逐世代證據重算（各世代缺合約價） | 收入在低／高成本下仍隨在役世代組合小幅變動（FY30 10.1／8.5 vs 9.4） |
| k 的證據倍數口徑 | 每 MW 報價 ÷ IF_HoldEcon、每 GPU 小時報價 ÷ IF_GPUhrEcon（同世代、基準成本情境） | 都換成每 GPU 小時 | 兩者等價；只影響呈現 |
| 上限檢查的分子 | 每 MW 計費收入（100% 計費 × 利用率） | 100% 計費時數收入 | 比例 ÷ 0.92–0.95 |
| Q2 計費比例、利用率、服務、世代組合 | 90%（季末）、同模型、同模型、季末組合 | 另估 | (i)(ii)(iv)(v) 各 ±0.1 以內 |
| Tokenomics 版本 | v5.27（引用值與 v5.26 相同） | 維持 v5.26 | 無數字影響 |
| 舊方法回歸的基準 | v4.6 成品（revenue=legacy、其餘 v4.6 設定、快照 v5.26 取自 git 歷史）；仍可 LEGACY_BASE=v4.5 | — | 無 |
| JS 無槓桿 NOL | 虧損全額加回（與 Excel 同；原只加回 80%） | — | v4.6 前未觸發 |
| LibreOffice 運算表 | 核對時逐情境另存重算（不用 LibreOffice 的 MULTIPLE.OPERATIONS）；成品設 fullCalcOnLoad（Excel 開檔全部重算） | — | 無成品數字影響；成品檔內 LibreOffice 快取值的少數格（例如積極情境 FY30 累計新股）可能殘留，Excel 開檔即重算 |
| 反向 DCF 無解 | rv_solve 寫 null，畫面顯示「無解」（穩態 EBITDA 率 99% 內無法使 DCF＝現價） | — | 無 |
| 差異原因 | 營收（tkAnchor）觀點原因（FY26／FY27 營收低於共識超過 5%） | — | 建置檢查通過 |

## 8. 已知限制
1. **DCF 失效與終值占比**：基準與積極情境常態化 FCF ≤ 0，DCF 失效、權重歸零；保守情境 DCF 股權為負、0 截斷。三情境目標價完全來自 EV/EBITDA 腿（FY29 錨定 × 6x，扣錨定年末淨負債；上表「EV/EBITDA 腿單獨」）。r1 的終值占 EV 144% 已不適用（DCF 失效時終值不入目標價）；DCF 精度低的限制仍在。
2. **k_長約 證據薄**：只有 IREN 一家供應商的兩筆合約已證實；第二家供應商、VR200 的長約價找不到。
3. **隨需占比未揭露**：基準 0% 為假設；Q2 隱含 k 0.84 高於 0.76，可能反映部分隨需或舊約高價，但不得用來反推。
4. 舊未解：CoreWeave「active power」口徑（預設 IT）；Q2 實際營運成本 0.88 對模型 2.09（已決定事項 13）；Excel 實機開啟待 Andy。

## 9. 可移植性
- 新增公司特有內容：`pricing.anchorMultiple`（k_長約、k_現貨、基準證據、隨需占比、證據表、合約組合、找不到清單）、`methodology.checks.revCapShareMax`、`varianceReasons` 的 `perMw: {revenue: tkAnchor}` 原因。換公司時：填 anchorMultiple（證據的 tkName＋gen 對到 Tokenomics 世代、`refEvidence` 指向基準證據）、隨需占比；沒有 anchorMultiple 時不建 W4 列、收入方法不得為 tkAnchor（建置即報錯）；快照沒有 IF_RevGWFleet 時上限檢查「不適用」。RPO 覆蓋率對照列需要 RPO 桶資料。
- 通用工具：`scripts/attrib_w4.py`、`scripts/q2_check_w4.py`、`scripts/compare_w4.py`；`scripts/make_expect.py` 規則檔 `renames`；`scripts/verify_legacy.sh` `LEGACY_BASE`；`xlx.py` 運算表逐情境重算。

## 10. verify 輸出

- 新方法 `scripts/verify.sh --vs-dist`（dist＝v4.7 r2；成品＝重建，0 差異）：25 項全過；cmp31 三情境各 474 項、FY27 錨定 438 項全 OK。
- 升版驗收 `DATE=2026-10-08 EXPECT=scripts/expect/v4_7_vs_v4_6.txt scripts/verify.sh --vs-dist`（dist＝v4.6）：25 項全過（清單 978 格、列改名 4；其餘 0 差異）。之後只改了「評價_反向DCF」一格說明文字（下游工作表，在清單規則內）。
- 舊方法回歸 `scripts/verify_legacy.sh`（LEGACY_BASE＝v4.6）：25 項全過，畫面文字、Excel 值與公式對 v4.6 成品 0 差異。

<details><summary>新方法 --vs-dist（對 v4.7 成品）</summary>

```
================ 結果 ================
PASS  check_quarterly：季度加總＝年度、差異原因齊全
PASS  README 欄位說明與 company.json 一致
PASS  Tokenomics 快照 --check：26 個名稱（missing 0）與快照一致（20261008_Tokenomics_v5.27.xlsx，相對誤差 ≤ 1e-09）
PASS  建 HTML
PASS  建 Excel
PASS  重算：0 公式錯誤
PASS  反向 DCF：Excel 求解，rv_snap.json 與 Excel 一致
PASS  每 MW 敏感度快照：Excel 求值，permw_sens.json 與 Excel 一致
PASS  fix_outline
PASS  fix_datatable
PASS  verify_ooxml：OOXML OK
PASS  快照值＝Excel 分頁值：26 個名稱（missing 0）、330 個值、110 個具名範圍一致；其他工作表 80 格公式引用 60 個 TK_ 名稱（v5.27）
PASS  離線開啟：新建 HTML
PASS  離線開啟：dist/ 成品
PASS  cmp31 low：474 項全部 OK
PASS  cmp31 base：474 項全部 OK
PASS  cmp31 high：474 項全部 OK
PASS  cmp31 base_FY27：438 項全部 OK
PASS  test_quarterly：假設實際數與可移植性測試通過
PASS  test_rolling：日曆推算與滾動後第一屏
PASS  test_attrib：拆解工具（月數、同版 0、滾動一季與 WACC 12%）
PASS  test_permw：設施口徑換算與 GPU 小時價格路線
PASS  crawl 畫面文字 0 差異
PASS  xl_diff --values --by-label 0 差異
PASS  xl_diff 公式 --by-label 0 差異
verify.sh：全部通過（25 項）
EXIT 0
```
</details>

<details><summary>升版驗收（對 v4.6 成品，EXPECT）</summary>

```
================ 結果 ================
PASS  check_quarterly：季度加總＝年度、差異原因齊全
PASS  README 欄位說明與 company.json 一致
PASS  Tokenomics 快照 --check：26 個名稱（missing 0）與快照一致（20261008_Tokenomics_v5.27.xlsx，相對誤差 ≤ 1e-09）
PASS  建 HTML
PASS  建 Excel
PASS  重算：0 公式錯誤
PASS  反向 DCF：Excel 求解，rv_snap.json 與 Excel 一致
PASS  每 MW 敏感度快照：Excel 求值，permw_sens.json 與 Excel 一致
PASS  fix_outline
PASS  fix_datatable
PASS  verify_ooxml：OOXML OK
PASS  快照值＝Excel 分頁值：26 個名稱（missing 0）、330 個值、110 個具名範圍一致；其他工作表 80 格公式引用 60 個 TK_ 名稱（v5.27）
PASS  離線開啟：新建 HTML
PASS  離線開啟：dist/ 成品
PASS  cmp31 low：474 項全部 OK
PASS  cmp31 base：474 項全部 OK
PASS  cmp31 high：474 項全部 OK
PASS  cmp31 base_FY27：438 項全部 OK
PASS  test_quarterly：假設實際數與可移植性測試通過
PASS  test_rolling：日曆推算與滾動後第一屏
PASS  test_attrib：拆解工具（月數、同版 0、滾動一季與 WACC 12%）
PASS  test_permw：設施口徑換算與 GPU 小時價格路線
PASS  crawl 畫面文字：升版預期差異（頁面無錯誤、無缺頁；明細見 out/crawl_upgrade_diff.txt）
PASS  xl_diff --values --by-label 0 差異
PASS  xl_diff 公式 --by-label 0 差異
verify.sh：全部通過（25 項）
EXIT 0
```
</details>

<details><summary>舊方法回歸（對 v4.6 成品）</summary>

```
================ 結果 ================
PASS  check_quarterly：季度加總＝年度、差異原因齊全
PASS  README 欄位說明與 company.json 一致
PASS  Tokenomics 快照 --check：25 個名稱（missing 0）與快照一致（20261007_Tokenomics_v5.26.xlsx，相對誤差 ≤ 1e-09）
PASS  建 HTML
PASS  建 Excel
PASS  重算：0 公式錯誤
PASS  反向 DCF：Excel 求解，rv_snap.json 與 Excel 一致
PASS  每 MW 敏感度快照：Excel 求值，permw_sens.json 與 Excel 一致
PASS  fix_outline
PASS  fix_datatable
PASS  verify_ooxml：OOXML OK
PASS  快照值＝Excel 分頁值：25 個名稱（missing 0）、315 個值、105 個具名範圍一致；其他工作表 75 格公式引用 55 個 TK_ 名稱（v5.26）
PASS  離線開啟：新建 HTML
PASS  離線開啟：dist/ 成品
PASS  cmp31 low：450 項全部 OK
PASS  cmp31 base：450 項全部 OK
PASS  cmp31 high：450 項全部 OK
PASS  cmp31 base_FY27：429 項全部 OK
PASS  test_quarterly：假設實際數與可移植性測試通過
PASS  test_rolling：日曆推算與滾動後第一屏
PASS  test_attrib：拆解工具（月數、同版 0、滾動一季與 WACC 12%）
PASS  test_permw：設施口徑換算與 GPU 小時價格路線
PASS  crawl 畫面文字 0 差異
PASS  xl_diff --values --by-label 0 差異
PASS  xl_diff 公式 --by-label 0 差異
verify.sh：全部通過（25 項）
```
</details>

<details><summary>新方法 verify.sh --vs-dist 完整輸出</summary>

```
=== 0. 季度層檢查（季度加總＝年度、指引一致性、超過門檻的差距都有原因；v4.4）
一致性檢查 14 組
low：季度加總＝年度 10 項
low：年度差異原因 11 項、季度差異原因 6 項
base：季度加總＝年度 10 項
base：年度差異原因 10 項、季度差異原因 6 項
high：季度加總＝年度 10 項
high：年度差異原因 11 項、季度差異原因 6 項
check_quarterly：全部通過
[PASS] check_quarterly：季度加總＝年度、差異原因齊全

=== 0b. README 欄位說明與 company.json 一致（scripts/fields_doc.py --check；5a）
README 欄位說明與 company.json 一致
[PASS] README 欄位說明與 company.json 一致

=== 0c. Tokenomics 快照可重現（tools/tokenomics/import_tokenomics.py --check；W1）
--check 通過：26 個名稱（missing 0）與快照一致（20261008_Tokenomics_v5.27.xlsx，相對誤差 ≤ 1e-09）
[PASS] Tokenomics 快照 --check：26 個名稱（missing 0）與快照一致（20261008_Tokenomics_v5.27.xlsx，相對誤差 ≤ 1e-09）

=== 1. 建 HTML（v4.7 · 2026-10-08）
built 1028730
[PASS] 建 HTML

=== 2. 建 Excel
saved /home/claude/company-models/coreweave/out/20261008_CoreWeave收支模型_v4_7.xlsx
[PASS] 建 Excel

=== 3. 重算（LibreOffice headless）
{"status": "success", "total_formulas": 2573, "total_errors": 0, "error_summary": {}}
[PASS] 重算：0 公式錯誤

=== 3b. 反向 DCF：以 Excel 求解（scripts/rv_solve.py；v4.5 取代 JS 快照）
反向 DCF（Excel 求解）：現價 90.13、R 1.9736、C 0.1856、Eb 無解、Rt 1.7080；矩陣 20/20 格有解
rv_snap.json 無變動
[PASS] 反向 DCF：Excel 求解，rv_snap.json 與 Excel 一致

=== 3c. 每 MW 敏感度快照：以 Excel 求值（scripts/permw_sens.py；W2）
每 MW 敏感度快照（Excel 求值）：基準情境加權目標價 5.25、Tokenomics 低／高 39.99／0.00、Rubin Ultra 版 0.00
permw_sens.json 無變動
[PASS] 每 MW 敏感度快照：Excel 求值，permw_sens.json 與 Excel 一致

=== 4. fix_outline
outline fixed
[PASS] fix_outline

=== 4b. fix_datatable（模擬運算表還原為 Excel 格式）
datatable fixed: xl/worksheets/sheet3.xml D245:E247 輸入格 C5
[PASS] fix_datatable

=== 5. verify_ooxml
OOXML OK
[PASS] verify_ooxml：OOXML OK

=== 5d. 快照值＝Excel「Tokenomics_取數」分頁值（scripts/check_tokenomics_tab.py；W1）
快照值＝Excel 分頁值：26 個名稱（missing 0）、330 個值、110 個具名範圍一致；其他工作表 80 格公式引用 60 個 TK_ 名稱（v5.27）
[PASS] 快照值＝Excel 分頁值：26 個名稱（missing 0）、330 個值、110 個具名範圍一致；其他工作表 80 格公式引用 60 個 TK_ 名稱（v5.27）

=== 5c. 離線開啟檢查（已決定事項 11：單一檔案、無網路請求、console 無錯誤、關鍵數字正常顯示）
(a) 網路請求：0 個
(b) console 錯誤與例外：0 個
(c) 關鍵數字：5 個分頁、15 項，失敗 0 項
離線開啟檢查：通過
[PASS] 離線開啟：新建 HTML
(a) 網路請求：0 個
(b) console 錯誤與例外：0 個
(c) 關鍵數字：5 個分頁、15 項，失敗 0 項
離線開啟檢查：通過
[PASS] 離線開啟：dist/ 成品

=== 6. 三情境 xlx＋cmp31（預設錨定）
ok 1 1098
[PASS] cmp31 low：474 項全部 OK
ok 2 1098
[PASS] cmp31 base：474 項全部 OK
ok 3 1098
[PASS] cmp31 high：474 項全部 OK

=== 7. FY27 錨定 cmp31（基準）
ok 2 1098
[PASS] cmp31 base_FY27：438 項全部 OK

=== 8. 季度層測試（暫存副本：假設 Q3 實際數、可移植性；v4.4）
[PASS] test_quarterly：假設實際數與可移植性測試通過
  畫面錯誤：無
=== 結果
test_quarterly：全部通過（測試 A 假設實際數、測試 B 可移植性）

=== 8b. 期間滾動測試（暫存副本：日曆推算 6 種情況、滾動後第一屏無過期日期與期間字樣；v4.5）
[PASS] test_rolling：日曆推算與滾動後第一屏
C 滾動檢查：只滾日曆、未更新 asOf → 建置失敗並列出 24 項（清單 24 項）
A 滾動 FY26Q3（評價日 2026-06-30 → 2026-09-30）：過期字樣 ['2026-06-30', 'Q2 2026', '6/30', '1H26', '2H26', '下半年', '上半年', '1H', 'H1', '2H']；第一屏命中 0 行
期間滾動測試：通過

=== 8c. 目標價變動拆解工具測試（scripts/attrib.py：(a)＝(1＋WACC)^(月數÷12)，WACC 讀 Excel、月數讀 calendar_q；四項相加＝總變動）
[PASS] test_attrib：拆解工具（月數、同版 0、滾動一季與 WACC 12%）
目標價變動拆解測試：通過

=== 8d. 每 MW 改寫測試（暫存副本：MW 設施口徑換算、GPU 小時價格路線；W2）
[PASS] test_permw：設施口徑換算與 GPU 小時價格路線
test_permw：全部通過（測試 A 設施口徑、測試 B GPU 小時價格路線）

=== 9. 與 dist/ 成品比對
35 views; errors []
35 views; errors []
畫面：dist 35 個、新版 35 個頁面／分頁，約 135,067 字；差異 0 個；頁面錯誤 dist 0、新版 0
[PASS] crawl 畫面文字 0 差異
0 differences
以列名稱配對的工作表：無（全部逐格）
新增列 0：
[PASS] xl_diff --values --by-label 0 差異
0 differences
以列名稱配對的工作表：無（全部逐格）
新增列 0：
[PASS] xl_diff 公式 --by-label 0 差異

================ 結果 ================
PASS  check_quarterly：季度加總＝年度、差異原因齊全
PASS  README 欄位說明與 company.json 一致
PASS  Tokenomics 快照 --check：26 個名稱（missing 0）與快照一致（20261008_Tokenomics_v5.27.xlsx，相對誤差 ≤ 1e-09）
PASS  建 HTML
PASS  建 Excel
PASS  重算：0 公式錯誤
PASS  反向 DCF：Excel 求解，rv_snap.json 與 Excel 一致
PASS  每 MW 敏感度快照：Excel 求值，permw_sens.json 與 Excel 一致
PASS  fix_outline
PASS  fix_datatable
PASS  verify_ooxml：OOXML OK
PASS  快照值＝Excel 分頁值：26 個名稱（missing 0）、330 個值、110 個具名範圍一致；其他工作表 80 格公式引用 60 個 TK_ 名稱（v5.27）
PASS  離線開啟：新建 HTML
PASS  離線開啟：dist/ 成品
PASS  cmp31 low：474 項全部 OK
PASS  cmp31 base：474 項全部 OK
PASS  cmp31 high：474 項全部 OK
PASS  cmp31 base_FY27：438 項全部 OK
PASS  test_quarterly：假設實際數與可移植性測試通過
PASS  test_rolling：日曆推算與滾動後第一屏
PASS  test_attrib：拆解工具（月數、同版 0、滾動一季與 WACC 12%）
PASS  test_permw：設施口徑換算與 GPU 小時價格路線
PASS  crawl 畫面文字 0 差異
PASS  xl_diff --values --by-label 0 差異
PASS  xl_diff 公式 --by-label 0 差異
verify.sh：全部通過（25 項）
EXIT 0
```
</details>
