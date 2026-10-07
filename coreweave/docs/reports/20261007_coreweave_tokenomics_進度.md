# CoreWeave × Tokenomics 改造進度（接手用；W1–W3 共用）

## 目前狀態（每次 push 前覆寫）
- 已完成：W0 遷移（chat 端）；W1 步驟 0–6（資料蒐集 `coreweave/data/permw_inputs_20261007.json`；對照表 xlsx＋md；Excel「Tokenomics_取數」分頁＋55 個 TK_ 具名範圍、verify 23 項全過；快照 v5.24、company.json `tokenomics` 區段；取數工具 `tools/tokenomics/import_tokenomics.py`；分支 `claude/coreweave-w1-tokenomics` 自 `claude/coreweave-w0-migrate` a29db91 開出、draft PR、Tokenomics 唯讀副本 master 098873a／v5.24）。
- 下一步：W1 步驟 7——整體 verify（--vs-dist）與 PR 回報、PR 改 ready。
- 未解問題：Tokenomics v5.25（IF_DeprLifeIT 等 10 個名稱）尚未合併到 Tokenomics master；本張以 v5.24 產生快照，10 個名稱列為 optional（記為 missing）。

## 工作單總覽
| 工作單 | 分支 | PR | 狀態 |
|---|---|---|---|
| W0 遷移 | `claude/coreweave-w0-migrate` | | chat 端完成 |
| W1 Tokenomics 取數層 | `claude/coreweave-w1-tokenomics`（疊加於 W0 分支） | #6 | 進行中 |
| W2 每 MW 改寫 | `claude/coreweave-w2-permw` | | 未開始 |
| W3 v4.6 成品與對照 | `claude/coreweave-w3-v4.6` | | 未開始 |

<!-- 各工作單在下方新增自己的段落：「## Wx」＋步驟紀錄表（步驟｜狀態｜commit｜備註） -->

## W1 Tokenomics 取數層

Tokenomics 版本：`model/CURRENT`＝`20261007_Tokenomics_v5.24.xlsx`，master HEAD `098873a3c6855d1ef5e3414d5a0f12e39b2644ed`（2026-10-07 讀取）。

| 步驟 | 狀態 | commit | 備註 |
|---|---|---|---|
| 0 開分支、draft PR #6、Tokenomics 唯讀副本 | 完成 | d4775c1 | W0 未合併：自 `origin/claude/coreweave-w0-migrate` a29db91 開分支，PR base＝W0 分支 |
| 1 取數工具＋根目錄 README 段落 | 完成 | （本 commit） | 測試：產生 v5.24 快照 25 名（missing 10）；`--check` 通過；`--force-recalc`（LibreOffice 重算副本）值＝快取值；IF_Hdr*／不存在名稱報錯；竄改快照 1e-6 時 `--check` 失敗；無 clone 時警告略過 |
| 2 名稱清單、快照、company.json `tokenomics`、README 欄位表 | 完成 | （本 commit） | `coreweave/data/tokenomics_names.txt`（15 個現有名稱＋10 個 optional；另加 `IF_RacksPerGW`、`IF_Util` 供對照）；`coreweave/data/tokenomics_snapshot_v5.24.json`；`fields_doc.py --write` |
| 3 v5.25 名稱列 optional | 完成（以 v5.24 產快照） | 同上 | Tokenomics master 仍為 v5.24（098873a）；10 名記為 `{"missing": true}`。**W2 第 0 步：若 master 已是 v5.25，重抓快照（檔名改 v5.25）並刪 v5.24 快照** |
| 4 Excel「Tokenomics_取數」分頁、TK_ 具名範圍、verify 新增 0c／5d | 完成 | （本 commit） | 分頁放在「來源」之後（最後一頁，灰色標籤）；55 個具名範圍（missing 名稱有列、不建名稱）；`xl_diff.py` 新增工作表列為「新增工作表」不計差異；`verify.sh --vs-dist` **23 項全過**（原 21＋0c＋5d），crawl 0 差異、xl_diff 值與公式 0 差異 |
| 5 對照表（給 W2） | 完成 | （本 commit） | `coreweave/docs/plan/20261007_coreweave_Tokenomics對照表.xlsx`＋`.md`，由 `coreweave/scripts/build_tk_map.py` 自 company.json＋快照產生；11 列（工作單 10 項＋GPU 數、GPU 小時持有成本兩列 W2 新增項）；比較世代預設 FY26–27＝GB300、FY28–30＝VR200 [Assumed] |
| 6 資料蒐集（MW 口徑、世代、GPU 小時價格、10-Q 費用、站點電價） | 完成 | （本 commit） | `coreweave/data/permw_inputs_20261007.json`；摘要見下方「W1 第 6 步資料摘要」 |

### W1 第 6 步資料摘要（完整逐筆在 `coreweave/data/permw_inputs_20261007.json`）

**6a MW 口徑：視為 IT 關鍵電力 [Assumed]（有三項佐證）**
- CoreWeave 自己的文件（10-K FY2025、10-Q Q2 2026、Q2 新聞稿、法說逐字稿）只給數字（年底 >850 MW、6/30 1.5 GW active；簽約 3.7 GW，法說日 4.2 GW），找不到「active／contracted power」是 IT 或設施口徑的定義。
- 房東以 critical IT 計算對 CoreWeave 的容量：Core Scientific（2025-02-26：590 MW，Denton「full critical IT load just over 260 MW」，並把 1.3 GW 電網合約電力與約 900 MW 可出租容量分開）、Galaxy Helios（2025-04-23：「393 MW of critical IT load」）[Interested-party]。
- 密度對照：2024 年底約 250,000 GPU ÷ 360 MW ≈ 694 顆/MW，接近 Tokenomics Hopper 每 IT MW 718 顆（−3%），遠於設施口徑的 598 顆（+16%）[Derived，只作口徑判斷]。
- 若改為設施口徑：IT MW＝設施 MW ÷ 1.2（6/30 1,500 → 1,250 IT MW）；CRWV 每 MW 數字換成每 IT MW × 1.2（每 MW 年收入 11.2 → 13.4、建置成本 34 → 40.8 US$m）；每 MW 利潤率不變。

**6b 世代組合（6/30 在役 1,500 MW）：Hopper 27%／GB200 29%／GB300 43% [Assumed]**，區間 Hopper 24–31%、GB200 20–42%、GB300 30–53%。方法：依年份分層——2024 年底前 360 MW＝Hopper（含 A100）；2025 年新增 490 MW＝GB200 70%／GB300 20%／H200 10%；2026 上半年新增 650 MW＝GB300 85%／GB200 15%。MW 錨點 [Verified]，各層占比 [Assumed]（公司未揭露機隊世代結構）。2H26–FY27 新增：GB300 為 2026 主力、Vera Rubin Q2 2026 才完成首套驗證、Meta $21bn 擴約（2026-04-09）含首批 Vera Rubin——與 W2 工作單 `newMix` 預設方向一致，無更精確資料。

**6c GPU 小時價格（US$/GPU-hr）：不足以支撐 W2 的 GPU 小時價格路線 → 建議 W2 採備案（revenue=legacy）**

| 世代 | 數值 | 口徑 | 來源（日期） | 標記 |
|---|---|---|---|---|
| H100 | 6.16 | 隨需牌價（8 GPU $49.24/hr；長約最多 −60%） | CoreWeave 價目（2026-10-07 擷取） | [Interested-party] |
| H100 | 2.82 | 市場租金指數 | Silicon Data（2026-10-06） | [Verified] |
| H100 | 3.85（10/1 起 4.50） | 隨需牌價 | Nebius 價目 | [Interested-party] |
| H100 | 3.99 | 隨需牌價 | Lambda 價目 | [Interested-party] |
| GB200 | 10.50 | 隨需牌價（4 GPU $42/hr；長約最多 −60%） | CoreWeave 價目 | [Interested-party] |
| GB200 | 8.00 起 | 隨需牌價 | GMI Cloud（2026-07-24） | [Interested-party] |
| GB200（以 B200 類比） | 5.87（區間 4.08–6.11） | 市場指數 | Silicon Data（2026-10-06）；Ornn 4.08（2026-04-13，經 WSJ 轉述） | [Analogy] |
| GB300（以 B300 類比） | 6.82（區間 6.82–7.85） | 市場指數 | Silicon Data（2026-10-06） | [Analogy] |
| GB300（以 B300 類比） | 7.85（10/1 起 9.50） | 隨需牌價 | Nebius 價目 | [Analogy] |
| GB300 | 9.70 US$m/MW/年（≈2.27/GPU-hr） | 同業五年合約（$9.7bn ÷ 5 年 ÷ 200 MW critical IT；GPU 數未揭露，換算用 Tokenomics 487 顆/MW） | IREN–Microsoft（2025-11-03） | [Derived] |
| VR200 | 找不到 | — | CoreWeave／Nebius／Lambda 價目未列；Meta 合約未揭露單價 | — |

- 牌價都不含網路與儲存（CoreWeave、Nebius 儲存另計）。所有合約期間：牌價＝隨需；指數＝市場現行租期；只有 IREN 為 5 年合約。
- 未使用 SemiAnalysis／InferenceX。GB200 的 CoreWeave 牌價與 Tokenomics `L1_GPUhr_GB200_vsCW` 外部值為同一底層來源，不算獨立。

**6d 公司營運費用（10-Q Q2 2026，US$m，三個月）**：營收 2,575；營收成本 879；技術與基礎設施 1,507；銷售行銷 60（其中 SBC 12）；一般管理 178（其中 SBC 84）；營業損益 −49；D&A 合計 1,393（未依列別揭露）；SBC 合計 165（技術與基礎設施 60、營收成本 9）；營業租賃成本 500、變動租賃 150、融資租賃 17。
- **管銷率＝（60 − 12 ＋ 178 − 84）÷ 2,575 ＝ 142 ÷ 2,575 ＝ 5.51% [Derived]**；假設：銷售行銷與一般管理內的 D&A＝0 [Assumed]。區間 5.51%（扣 SBC）～9.24%（GAAP 含 SBC）；1H26 扣 SBC 為 6.15%。

**6e 站點電價**：找不到（CoreWeave 與房東 Applied Digital、Core Scientific、Galaxy 公告皆未揭露電價或電費負擔方）；電價以 Tokenomics 為準。
