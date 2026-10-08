# Microsoft v0.1 進度檔（接手用）

## 目前狀態（每次 push 前覆寫）
- **已完成：v0.1 交付前修訂（查核後，不升版）**。引擎與 amazon@`d07bfbe` 逐檔相同；ownMultiple 15.67×、customCapexFactor；(a) neocloud 租約續約延續租金（−$2.41）、(b) Maia 200 → GB300（−$1.29）、(c) 主值措辭、(d) 長約 k 標 [Analogy]、(e) 共識 FY29 淨負債標資料可疑、(f) 一頁摘要第 3 塊不再裁切；dist 重建；verify.sh --vs-dist 26 項全過。讀法 1（D2）382.25／353.55／351.96；讀法 2 407.13／422.48／422.39；主值待 Andy 決定。
- **下一步**：FY27Q1 季度更新（2026-10-28 後），升版 v0.2。
- **未解問題（待 Andy）**：讀法主值；對外比例；長約 k 證據；未起租營業部分租金。

## v0.1 交付前修訂（2026-10-09；查核後，不升版）
| 項 | 內容 | 結果 |
|---|---|---|
| 引擎同步 | amazon@d07bfbe：segA、segB、segE、build_xlsx.py、cmp31.js、rv_solve.py、mag_sens.js、test_mag_mechanisms.py、fields_doc.py（mag 說明保留本家） | 逐檔相同（公司資料、文件、VLOG 除外） |
| 評價口徑 | ownMultiple 15.67×（val.msft.ownNtm）；翻轉點 WACC 8.15%、非 AI 18.05×、AI 26.2×；終值占比 94.5% 封鎖買進 | 報告第 6 節 |
| (a)(b) | 租金 years 10；Maia 200 → GB300；基準速度重解 2,325.9 | 基準 357.24 → 353.55 |
| (c)(d)(e)(f) | 措辭；kevidence [Analogy]；varianceReasons FY29 nd「資料可疑」；截圖 | — |
| verify | --vs-dist 26 項全過（cmp31 基準 554 項） | — |

## v0.1c（2026-10-09）
| 步驟 | 內容 | commit |
|---|---|---|
| 1–4 | VLOG 重設 v0.1；dist/ 移除 Oracle v0.2、放入 Microsoft v0.1；--vs-dist 26 項全過；crawl：描述本公司的 Oracle／OCI 為 0（只剩版本沿革、同業 ORCL）；截圖 17 張（溢出 0、錯誤 0）；交接檔 v0.1 取代 Oracle v0.2；README／CLAUDE.md 開頭；根目錄 README Microsoft 列 | a34d1f4 |
| 5–6 | 最終報告、重點 Excel（6 頁）、PR 回報並轉 ready | （本次） |

## v0.1b r3 同步（2026-10-09；對照表 r1 第 9 節 C16–C23）
| 步驟 | 內容 | commit |
|---|---|---|
| 引擎同步 | 自 amazon@6ac9745 複製 r3 引擎檔（segA、segB、segD、segE、build_xlsx.py、cmp31.js、rv_solve.py、mag_sens.js、test_mag_mechanisms.py、fields_doc.py〔mag 說明換回本家〕、tail.js、vlog.py）；`data/peers_ads_20261009.json` 只讀複製（C18） | （本次） |
| company.json | C20 `capexModel.rentedExt`（期初 550、FY27 起 756、每 MW 淨租金 6.55）；C21 opShare 0.25 → 0.75（統一為營業租賃比例）、cashShareLabel null；rentedCompute 改每筆＝自用 50% 全額＋對外 50% 淨額；C17 nonAiLife 10；daReconTol；kdNote；gpuLifeNote；股利年增 8%；廣告同業 6 家；基準速度重解 2,376 MW-IT／年 | （本次） |
| 結果 | 讀法 1：385.73／357.24／356.78；讀法 2：410.44／426.15／426.13；FY29 對外 AI ROIC 10.1% vs WACC 10.9%、打平 k 1.084；verify.sh 23 項全過 | — |

## v0.1b′ 移植（2026-10-09）
工作單：`mag/docs/workorders/20261008_mag_v0.1b2_移植.md`（對照表 r1 第 7、8 節 C1–C15 優先）；引擎來源：`origin/claude/amazon-v0.1` @ `59f4d5f`（Amazon 進度檔已寫「v0.1b r2 完成」，C10–C15 已含在內）。

| 步驟 | 狀態 | commit | 備註 |
|---|---|---|---|
| 0 進度檔段落 | 完成 | （步驟 1–2 同 commit） | — |
| 1 複製引擎 | 完成 | （本次） | Amazon 移植說明第 6 節清單＋r2 改檔（`scripts/calib_pace.js` 等）原樣複製；保留本家 `data/`、`docs/`、`dist/`（Oracle v0.2 成品不動）、`CLAUDE.md`；刪除 `data/consensus_orcl_20261007.json`、`consensus_crwv_20260925.json`；Tokenomics 快照改用 Amazon 引擎的名稱清單重產（35 名、missing 3：IF_NonNVRatio、IF_NonNVCostRatio、IF_NonNV_Maia；`--check` 通過）；`data/peers_cloud_20261008.json` 自 amazon/ 只讀複製（C12） |
| 2 company.json | 完成 | （本次） | 以 Amazon r2 `company.json` 為骨架（產生腳本 `scripts/port_company_json.py`；Amazon 引擎欄位再變時可重跑）；財年 6 月、FY26Q4 已申報、首期 FY27 全年 1.0 年、期間 FY27–FY31、錨定 FY29（evYear＝2、roicYear＝2）；6 線＋AI 雲端；C10 基準速度 2,294.1 MW-IT／年（calib_pace 解；首期資本支出 204.2、對帳落差 0）；verify.sh 23 項全過 |
| 3 引擎差異 | 完成（隨步驟 1–2） | （本次） | 見下「引擎差異」 |
| 4 敏感度、報告、PR 留言 | 完成 | （本次） | `node scripts/mag_sens.js`＋暫存副本變體（opShare、起租季數、自建比例、期初 MW 不含租用、股利調升）；報告第 1–11 節；收尾 fetch amazon：無新 commit |

### 引擎差異（相對 Amazon @ 59f4d5f；全部在 microsoft/，chat 端決定是否回寫）
1. `build_xlsx.py`、`segB.js`：共識檔沒有 `recentActions`／`recentActionsMeta`（Microsoft v0.1a 共識檔未蒐集最新分析師動作）時略過，不報錯。
2. `build_xlsx.py`：普通股股利列備註改讀 `defaults.dividend.note`（原寫死 Oracle「每季 $0.50；不回購」）。
3. `build_xlsx.py`：未起租「自現金扣除比例」列名可由 `leases.uncommenced.cashShareLabel` 指定（本家＝營業租賃部分；見已套用預設 opShare），備註尾改為「本格＝1 − leases.uncommenced.opShare」。
4. `scripts/test_rolling.py`、`scripts/test_attrib.py`：Q4 已申報（年初至今無標籤）時，滾動一季的 ytdActual.label 改寫為新標籤；test_rolling 的「過期字樣」排除在新日曆仍有效者（新的上一財年末日期、期間標籤、「全年」）——Microsoft 型日曆（評價日＝上一財年末）才會碰到。
5. `scripts/mag_sens.js`：對外比例敏感度列名依 `capexModel.extShare` 產生（原寫死 80% 口徑）；有租用算力時加兩列（租金全額、租金 0）。
6. `scripts/fields_doc.py`：mag 區段說明換回 Microsoft v0.1a 版；新增 `leases.uncommenced.cashShareLabel`、`defaults.legacyBiz.split.*`（Microsoft 營業利益率分配）、`mag.mw.externalIT` 的說明。
7. `historicalPL` 第 4 格為年初至今（引擎以索引 3 讀取）：Microsoft 首期為全年，加一格「FY27 年初至今」全 0（company.json 表達，非引擎改動）。

## v0.1a 資料蒐集（2026-10-08）
工作單：`mag/docs/workorders/20261008_mag_v0.1a_資料蒐集.md`；分支 `claude/microsoft-v0.1`；PR #33。

| 步驟 | 狀態 | commit | 備註 |
|---|---|---|---|
| 0 基線、PR、進度檔 | 完成 | 4a1df12 | verify.sh 19 項全過；SEC EDGAR shell 可直連；draft PR #33 |
| 1 SEC 一手文件 | 完成 | 42a80ac／本次 | 10-K FY26、三份 10-Q、8 季新聞稿、2026-09-02 8-K（Azure 絕對值）；本次補回年度損益 45 筆（建置腳本漏寫） |
| 2 AI 容量（MW） | 完成 | 42a80ac | 公司增量＋Bloomberg、Morgan Stanley、Epoch、Omdia；期初 3.5／4.5／5.5 GW-IT |
| 3 k 證據 | 完成 | 42a80ac | Azure 牌價 11 筆、長約 5 筆；建議 k_長約 0.76、k_現貨 1.76、加權 1.06 |
| 4 租用算力與租賃 | 完成 | 42a80ac | Nebius、IREN、Nscale、Lambda、CoreWeave；未起租 329.1 |
| 5 關聯方 | 完成 | 42a80ac | OpenAI 25%、24.1、250、45% RPO、分成 20%／上限 38；Anthropic 30／≤5 |
| 6 AI／非 AI 拆分 | 完成 | 本次 | 對外 AI 36.5（17.3–67.7）、非 AI 81.2（50.0–100.3） |
| 7 評價輸入 | 完成 | 42a80ac | beta 1.10（兩源）、rf 5.28%、kd 6.26%（Aaa 類比）、同業 11 家 |
| 8 市場共識 | 完成 | 42a80ac | 現價 529.76；PT 587.63（S&P）／578.73（MarketBeat） |
| 9 mag 區段、verify、報告、PR 回報 | 完成 | 本次 | fields_doc 補 mag 21 子區說明並 --write；verify.sh 19 項全過 |
