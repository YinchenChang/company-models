# CoreWeave v4.7：每 MW 收入以 Tokenomics 為錨（v4.6 → v4.7）

2026-10-08｜工作單 W4 r1（`docs/workorders/20261008_coreweave_W4_收入Tokenomics錨定.md`）｜PR #27｜Excel 版：`docs/reports/20261008_coreweave_v4.7_收入錨定.xlsx`

## 結論
1. **CRWV 每 MW 收入（100% 計費時數）是 Tokenomics 打平租金（`IF_HoldEcon`，WACC 10%）的 0.76 倍（FY26–FY27）→ 1.34 倍（FY30，基準）**：近兩年 RPO 涵蓋全部在役產能，全以長約倍數 0.76 計價；FY28 起新產能超出 RPO，未涵蓋部分以現貨倍數 1.76 計價。
2. **這代表**：以市場長約價計，CRWV 的長約單價低於「含廠房資本回收、WACC 10% 的打平線」——每 MW 年收入 8.4–8.8 對錨 11.1–11.6，FY26–FY27 每 MW 稅前約 −5.2；價值取決於 FY28 後新產能能否以高於打平線的價格（現貨／新約）售出。
3. **結果由長約占比主導**：基準加權目標價 $29.68 → $98.02（三情境皆由賣出轉為中立）；若新簽約全視為長約（k＝0.76），基準只有 $5.25；長約占比 −20pt 則 $201.31。需 Andy 決定長約占比口徑。

## 需要 Andy 做的事
1. **決定長約占比口徑**（見第 8 節第 1 項；建議：新簽約視為長約〔與 10-K 加權平均約 5 年、模型新合約 5 年一致〕，另設明確的隨需占比）。
2. **決定 Tokenomics 成本情境下 k 是否重算**（第 8 節第 3 項；建議重算，讓「成本高 → 價值低」）。
3. 合併前以真正的 Excel 開啟 `dist/20261008_CoreWeave收支模型_v4_7.xlsx` 檢查（運算表與活公式只以 LibreOffice 驗證過）。

## 1. 關鍵數字（v4.6 → v4.7）
| 情境 | 加權目標價 v4.6 | v4.7 | 評等 v4.6 → v4.7 | 融資缺口 v4.6 → v4.7（US$bn） |
|---|---|---|---|---|
| 保守 4.2 GW | $43.99 | $91.24 | 賣出 → 中立 | 32.7 → 30.3 |
| 基準 5.6 GW | $29.68 | $98.02 | 賣出 → 中立 | 81.6 → 67.5 |
| 積極 8 GW | $19.95 | $131.86 | 賣出 → 中立 | 157.9 → 118.1 |

情境區間 $19.9–$44.0 → $91.2–$131.9；方法區間（基準，5–6x）$15.2–$29.7 → $80.5–$98.0；錨定 FY27 對照（基準）$13.4 → $55.1。終值占 EV（基準）144%。

## 2. 每 MW 年收入（US$m／MW／年）
### 2a. 100% 計費時數（B 區「每 MW 年收入」）：v4.6 輸入 vs v4.7 錨 × k
| 情境 | 版本 | FY26（2H） | FY27 | FY28 | FY29 | FY30 |
|---|---|---|---|---|---|---|
| 保守 4.2 GW | v4.6 | 11.20 | 11.50 | 11.50 | 11.00 | 10.50 |
| 保守 4.2 GW | v4.7 | 8.42 | 8.79 | 9.77 | 11.72 | 15.04 |
| 基準 5.6 GW | v4.6 | 11.20 | 11.50 | 11.50 | 11.00 | 10.50 |
| 基準 5.6 GW | v4.7 | 8.42 | 8.79 | 10.13 | 12.70 | 16.47 |
| 積極 8 GW | v4.6 | 11.20 | 11.50 | 11.50 | 11.00 | 10.50 |
| 積極 8 GW | v4.7 | 8.42 | 8.79 | 11.09 | 14.58 | 18.10 |

### 2b. 每平均在役 MW（含服務、含爬坡分母與利用率；彙總表）
| 情境 | 版本 | FY26（2H） | FY27 | FY28 | FY29 | FY30 |
|---|---|---|---|---|---|---|
| 保守 4.2 GW | v4.6 | 10.05 | 10.56 | 10.37 | 9.78 | 8.99 |
| 保守 4.2 GW | v4.7 | 7.64 | 8.16 | 9.00 | 10.31 | 12.12 |
| 基準 5.6 GW | v4.6 | 10.05 | 10.56 | 10.33 | 9.63 | 8.65 |
| 基準 5.6 GW | v4.7 | 7.64 | 8.16 | 9.24 | 10.88 | 12.77 |
| 積極 8 GW | v4.6 | 10.05 | 10.56 | 10.23 | 9.34 | 8.27 |
| 積極 8 GW | v4.7 | 7.64 | 8.16 | 9.91 | 11.99 | 13.51 |

### 2c. 錨與 k 的逐年值（基準；三情境見 Excel「錨與k」）
| 項目 | FY26（2H） | FY27 | FY28 | FY29 | FY30 |
|---|---|---|---|---|---|
| 錨（Σ 在役占比 × IF_HoldEcon） | 11.074 | 11.562 | 11.922 | 12.115 | 12.307 |
| 排程 RPO（US$bn） | 13.19 | 26.38 | 25.74 | 25.10 | 17.70 |
| 長約價產能收入（US$bn） | 6.10 | 19.20 | 28.27 | 35.25 | 41.93 |
| RPO 涵蓋 ÷ 在役計費產能 | 2.164 | 1.374 | 0.910 | 0.712 | 0.422 |
| 長約占比 | 1.000 | 1.000 | 0.910 | 0.712 | 0.422 |
| 定價倍數 k | 0.760 | 0.760 | 0.850 | 1.048 | 1.338 |
| 每 MW 年收入＝錨 × k | 8.416 | 8.787 | 10.129 | 12.697 | 16.465 |
| 相對 v4.6 舊輸入 | -0.249 | -0.236 | -0.119 | 0.154 | 0.568 |
| 上限：客戶每 MW 付費 token 營收（IF_RevGWFleet） | 19.10 | 24.17 | 29.63 | 33.82 | 37.66 |
| 上限：CRWV 計費收入 ÷ 客戶 token 營收 | 0.419 | 0.345 | 0.321 | 0.349 | 0.402 |

**上限檢查**：三情境各期 CRWV 每 MW 計費收入 ÷ 客戶付費 token 營收最高 42%（FY26），皆低於門檻 50%（`methodology.checks.revCapShareMax`），檢查頁「通過」。

**營運成本**（電費、IT 維護、人員軟體、稅險、租金）與 v4.6 **0 差異**（升版驗收清單不含這些列）；管銷＝營收 × 5.51%，隨營收改變。基準 FY27：現金成本（含租金）4.33 → 4.20、EBITDA 6.23 → 3.97；FY30：3.73 → 3.96、4.92 → 8.81。

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

- **最終值**：k_長約 0.76（0.7–1.0）、k_現貨 1.76（1.5–2.3），皆 [Analogy]；長約占比＝RPO 涵蓋的產能 ÷ 在役計費產能 [Derived]（基準 FY26–FY30：100%、100%、91%、71%、42%）。沒有一個 k 值來自 CRWV 的營收、ARR 或 RPO 金額；RPO 只用來估長約占比。
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
- 差距 7.3% 超過 5%：**類型＝觀點**。模型 k 0.76 取市場長約證據（IREN–Microsoft GB300 五年約）；Q2 營收含較早簽約的舊世代合約與隨需，隱含 k 約 0.84。依已決定事項 14，不回頭改 k 去貼近 Q2。

## 5. 敏感度（建置時快照 `permw_sens.json`；加權目標價 US$／融資缺口 US$bn）
| 設定 | 保守 | 基準 | 積極 |
|---|---|---|---|
| 基準（目前輸入） | $91.24／30.3 | $98.02／67.5 | $131.86／118.1 |
| Tokenomics 低成本 | $4.54／48.8 | $0.21／85.8 | $10.91／136.1 |
| Tokenomics 高成本 | $541.46／0.0 | $605.66／0.0 | $774.81／0.0 |
| GPU 小時價格 低 | 不適用 | 不適用 | 不適用 |
| GPU 小時價格 高 | 不適用 | 不適用 | 不適用 |
| 世代組合 Rubin Ultra 版 | $98.04／34.4 | $104.23／82.5 | $142.07／146.1 |
| 管銷率 GAAP（含 SBC） | $72.31／35.8 | $75.05／73.9 | $103.34／126.5 |
| k_長約 0.70 | $53.46／42.5 | $59.24／80.3 | $96.12／130.9 |
| k_長約 1.00 | $226.82／0.0 | $236.10／29.5 | $283.37／80.1 |
| k_現貨 1.50 | $62.41／36.7 | $54.16／77.2 | $63.17／135.1 |
| k_現貨 2.30 | $150.35／17.0 | $188.00／47.4 | $274.72／82.8 |
| 長約占比 -20pt | $178.62／10.9 | $201.31／45.6 | $257.23／90.6 |
| 長約占比 100%（新簽約全視為長約） | $13.98／54.8 | $5.25／104.8 | $0.00／183.5 |

讀法：(1) 長約占比與 k 是最大槓桿（基準 $5.25–$236.10）；(2) Tokenomics 低／高成本同時移動錨（收入）與成本，因錨含資本回收、收入增幅大於成本增幅，**高成本 → 目標價高**（基準 $605.66），方向與 v4.6 相反；(3) Rubin Ultra 版 $104.23（錨較高）。

## 6. 目標價變動拆解（已決定事項 12；0 截斷口徑，US$／股；`scripts/attrib_w4.py`）
(a) 時間推移 0（評價日 2026-06-30 不變）、(b) 實際數更新 0、(c) 假設變更 0（Tokenomics v5.26 → v5.27：引用的 25 個名稱數值與位置相同；新增 IF_RevGWFleet 只用於上限檢查）；全部屬 (d) 方法變更，依序：

| (d) 逐項 | 保守 | 基準 | 積極 |
|---|---|---|---|
| ① 錨取代舊輸入（k＝1：每 MW 年收入＝在役世代 IF_HoldEcon） | +30.16 | +15.46 | +8.82 |
| ② 套用定價倍數 k（長約／現貨依 RPO 涵蓋加權） | +17.09 | +52.88 | +103.09 |
| ③ 上限檢查（只新增檢查列） | +0.00 | +0.00 | +0.00 |
| 合計（＝總變動） | +47.25 | +68.34 | +111.91 |
| 融資缺口變動（US$bn） | -2.4 | -14.0 | -39.8 |

檢查：三情境依序各步相加＝v4.7 − v4.6（誤差 < 0.01 美元／股）；③＝0。① k＝1 時每 MW 年收入＝錨（11.1–12.3，與 v4.6 輸入 10.5–11.5 接近）；② 套用 k 後近年下降（0.76）、遠年上升（最高 1.34），DCF 腿由 0 變為正值。

## 7. 已套用的預設（問題｜採用的預設｜替代選項｜對結果的影響）
| 問題 | 預設 | 替代 | 影響 |
|---|---|---|---|
| 長約占比（工作單公式）的時間口徑 | 逐期：排程 RPO ÷（平均計費 MW × 錨 × k_長約 × 利用率 × 期間長度），截斷 0–100%；未涵蓋產能（含模型的新簽約）以 k_現貨 計價 | 新簽約視為長約（k＝0.76） | 基準 $98.02 → $5.25（敏感度「長約占比 100%」）；**最大判斷項，待 Andy** |
| 第二筆長約（IREN–NVIDIA 0.89）是否改基準 | 不改：與第一筆同一供應商、世代與 MW 口徑為類比，只支持區間 | k_長約 取兩筆平均 0.83 | k_長約 1.00 時基準 $236.10；0.83 約在 $98–$236 之間 |
| k 的證據倍數用哪個 Tokenomics 值 | 每 MW 報價 ÷ IF_HoldEcon、每 GPU 小時報價 ÷ IF_GPUhrEcon（同世代、基準成本情境，不隨成本情境變） | 都換成每 GPU 小時 | 兩者等價（HoldEcon＝GPUhrEcon × GPU 數 × 8,760）；只影響呈現 |
| k 是否隨 Tokenomics 成本情境變 | 固定（錨與成本同時動） | k 隨情境重算（＝市場價格 ÷ 新持有成本） | 低／高成本敏感度方向：目前高成本 → 高目標價 |
| 上限檢查的分子 | 每 MW 計費收入（100% 計費 × 利用率） | 100% 計費時數收入 | 比例 ÷ 0.92–0.95 |
| Q2 計費比例 | 季末 Billable 1,350 ÷ 季末在役 1,500＝90% | 平均在役對平均計費（未揭露） | (i) 爬坡分母 ±0.1 |
| Q2 利用率、服務收入 | 同模型（未揭露） | 另估 | (ii)＝0；(v) −0.03 |
| Q2 錨的世代組合 | 季末（6/30）世代組合 | Q2 平均（含 3/31） | (iv) ±0.1 |
| Tokenomics 版本 | v5.27（master 最新；引用值與 v5.26 相同） | 維持 v5.26 | 無數字影響 |
| 舊方法回歸的基準 | v4.6 成品（revenue=legacy、其餘 v4.6 設定、快照 v5.26 取自 git 歷史）；仍可 LEGACY_BASE=v4.5 | — | 無 |
| JS 無槓桿 NOL | 改為虧損全額加回（與 Excel 同；原只加回 80%） | — | v4.6 前未觸發；v4.7 下 FY30 UFCF 差 0.41，修正後 HTML＝Excel |
| 一頁摘要每 MW 句 | 沿用 W3 句型（倍數＝每 MW 年收入 ÷ 經濟持有成本，tkAnchor 下＝k） | 改寫成「定價倍數 k」 | 無數字影響 |
| 差異原因 | 新增「營收（tkAnchor）」觀點原因（FY26／FY27 營收低於共識超過 5%） | — | 建置檢查通過 |

## 8. 未解問題與資料缺口
1. **長約占比口徑（需 Andy 決定）**：工作單公式下 FY28 起 RPO 不足以涵蓋新產能，新簽約以現貨倍數 1.76 計價（基準 FY30 k 1.34、每 MW 年收入 16.5）。但模型的新簽約本身是 5 年合約，10-K 已承諾合約加權平均約 5 年。選項：(A) 維持（基準 $98.02）；(B) 新簽約視為長約、k＝0.76（基準 $5.25）；(C) 新簽約視為長約，另設隨需占比（例如 5–10%，需資料）。**建議 (C)**，在找到隨需比例前以 (B) 為主值、(A) 為敏感度。
2. **第二家供應商的長約價**：找不到（Nebius–Meta／Microsoft、Nscale–Microsoft 未揭露 MW 或金額）；IREN–NVIDIA 為同一供應商。VR200 長約價找不到。
3. **Tokenomics 成本情境與收入同向**：tkAnchor 下高成本使錨上升 → 收入上升，敏感度方向與直覺相反；若要「市場價格固定」的讀法，k 應隨情境重算（需 Andy 決定）。
4. 終值占 EV 144%（基準）：DCF 結論由 FY30 後的單價決定。
5. 舊未解：CoreWeave「active power」口徑（預設 IT）；Q2 實際營運成本 0.88 對模型 2.09（已決定事項 13）；Excel 實機開啟待 Andy。

## 9. 可移植性
- 新增公司特有內容：`pricing.anchorMultiple`（k_長約、k_現貨、長約占比方法與調整、證據表、合約組合、找不到清單）、`methodology.checks.revCapShareMax`、`varianceReasons` 的 `perMw: {revenue: tkAnchor}` 原因。換公司時：填 anchorMultiple（證據的 tkName＋gen 對到 Tokenomics 世代）；沒有 anchorMultiple 時不建 W4 列、收入方法不得為 tkAnchor（建置即報錯）；快照沒有 IF_RevGWFleet 時上限檢查顯示「不適用」。
- 通用工具：`scripts/attrib_w4.py`、`scripts/q2_check_w4.py`、`scripts/compare_w4.py`；`scripts/make_expect.py` 規則檔新增 `renames`；`scripts/verify_legacy.sh` 新增 `LEGACY_BASE`。
- 只適用 CoreWeave 的部分：長約占比以 RPO 排程估計（需公司有 RPO 與桶分攤）；Nebius 等沒有 RPO 桶的公司需另定長約占比來源（待決）。

## 10. verify 輸出

- 新方法 `scripts/verify.sh --vs-dist`（dist＝v4.7；成品＝重建，0 差異）：25 項全過。
- 升版驗收 `DATE=2026-10-08 EXPECT=scripts/expect/v4_7_vs_v4_6.txt scripts/verify.sh --vs-dist`（dist＝v4.6）：25 項全過（預期差異 1,022 格值＋209 格公式、列改名 4；其餘 0 差異）。
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
PASS  快照值＝Excel 分頁值：26 個名稱（missing 0）、330 個值、110 個具名範圍一致；其他工作表 70 格公式引用 60 個 TK_ 名稱（v5.27）
PASS  離線開啟：新建 HTML
PASS  離線開啟：dist/ 成品
PASS  cmp31 low：477 項全部 OK
PASS  cmp31 base：477 項全部 OK
PASS  cmp31 high：477 項全部 OK
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
PASS  快照值＝Excel 分頁值：26 個名稱（missing 0）、330 個值、110 個具名範圍一致；其他工作表 70 格公式引用 60 個 TK_ 名稱（v5.27）
PASS  離線開啟：新建 HTML
PASS  離線開啟：dist/ 成品
PASS  cmp31 low：477 項全部 OK
PASS  cmp31 base：477 項全部 OK
PASS  cmp31 high：477 項全部 OK
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
PASS  快照值＝Excel 分頁值：25 個名稱（missing 0）、315 個值、105 個具名範圍一致；其他工作表 65 格公式引用 55 個 TK_ 名稱（v5.26）
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
built 1028368
[PASS] 建 HTML

=== 2. 建 Excel
saved /home/claude/company-models/coreweave/out/20261008_CoreWeave收支模型_v4_7.xlsx
[PASS] 建 Excel

=== 3. 重算（LibreOffice headless）
{"status": "success", "total_formulas": 2573, "total_errors": 0, "error_summary": {}}
[PASS] 重算：0 公式錯誤

=== 3b. 反向 DCF：以 Excel 求解（scripts/rv_solve.py；v4.5 取代 JS 快照）
反向 DCF（Excel 求解）：現價 90.13、R 0.9572、C 1.0557、Eb 0.6612、Rt 0.9800；矩陣 20/20 格有解
rv_snap.json 無變動
[PASS] 反向 DCF：Excel 求解，rv_snap.json 與 Excel 一致

=== 3c. 每 MW 敏感度快照：以 Excel 求值（scripts/permw_sens.py；W2）
每 MW 敏感度快照（Excel 求值）：基準情境加權目標價 98.02、Tokenomics 低／高 0.21／605.66、Rubin Ultra 版 104.23
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
快照值＝Excel 分頁值：26 個名稱（missing 0）、330 個值、110 個具名範圍一致；其他工作表 70 格公式引用 60 個 TK_ 名稱（v5.27）
[PASS] 快照值＝Excel 分頁值：26 個名稱（missing 0）、330 個值、110 個具名範圍一致；其他工作表 70 格公式引用 60 個 TK_ 名稱（v5.27）

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
ok 1 1102
[PASS] cmp31 low：477 項全部 OK
ok 2 1102
[PASS] cmp31 base：477 項全部 OK
ok 3 1102
[PASS] cmp31 high：477 項全部 OK

=== 7. FY27 錨定 cmp31（基準）
ok 2 1102
[PASS] cmp31 base_FY27：438 項全部 OK

=== 8. 季度層測試（暫存副本：假設 Q3 實際數、可移植性；v4.4）
[PASS] test_quarterly：假設實際數與可移植性測試通過
  畫面錯誤：無
=== 結果
test_quarterly：全部通過（測試 A 假設實際數、測試 B 可移植性）

=== 8b. 期間滾動測試（暫存副本：日曆推算 6 種情況、滾動後第一屏無過期日期與期間字樣；v4.5）
[PASS] test_rolling：日曆推算與滾動後第一屏
C 滾動檢查：只滾日曆、未更新 asOf → 建置失敗並列出 24 項（清單 24 項）
A 滾動 FY26Q3（評價日 2026-06-30 → 2026-09-30）：過期字樣 ['2026-06-30', 'Q2 2026', '2H26', '1H26', '6/30', '上半年', '下半年', 'H1', '2H', '1H']；第一屏命中 0 行
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
畫面：dist 35 個、新版 35 個頁面／分頁，約 135,347 字；差異 0 個；頁面錯誤 dist 0、新版 0
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
PASS  快照值＝Excel 分頁值：26 個名稱（missing 0）、330 個值、110 個具名範圍一致；其他工作表 70 格公式引用 60 個 TK_ 名稱（v5.27）
PASS  離線開啟：新建 HTML
PASS  離線開啟：dist/ 成品
PASS  cmp31 low：477 項全部 OK
PASS  cmp31 base：477 項全部 OK
PASS  cmp31 high：477 項全部 OK
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
