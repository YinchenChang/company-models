# 工作單 WhiteFiber v0.2｜每 MW 收入改以 Tokenomics 為錨＋同步 Tokenomics v19.0（2026-10-22，r0，chat 端；給 Claude Code 雲端工作階段執行）

- 來源：Tokenomics 下游同步（2026-10-22 解除擱置）；WhiteFiber v0.1 仍用 TK v5.24 手抄的每 MW 單價 11.62／17.40／24.20（成本加成與市場價 50/50 平均，契約第 4 條已停用），容量與單價同向綁定，取數無快照（審視報告 `whitefiber/docs/reports/20261008_whitefiber_tokenomics_link_review.md`）。
- Tokenomics 現行：`model/CURRENT`＝`20261022_Tokenomics_v19.0.xlsx`，master `ad33e28`（若 master 已更新，以最新版取數並在報告列出差異）。
- **範本（最重要，照抄做法）**：Oracle v0.3（main，PR #60）——同一個引擎譜系（WhiteFiber v0.1 由 Oracle v0.2 `e4540c3` 複製）。讀 `oracle/docs/workorders/20261022_oracle_v0.3_Tokenomics錨定與v19同步.md`、`oracle/docs/reports/20261022_oracle_v0.3.md`、`oracle/docs/reports/20261007_oracle_進度.md` 的 v0.3 段落，以及 `oracle/` 的實作（`kspec.py`、`tkanchor.py`、`tk_sens*`、`segA.js`／`segB.js` 的 tkAnchor 與 kSpec、`build_xlsx.py` 的「Tokenomics_取數」與「每MW收入_錨定」分頁、`scripts/verify.sh` 與 `verify_legacy.sh`、`scripts/calib_ebitdar.js`）。**只讀，不改 `oracle/`、`nebius/`、`coreweave/`。**
- 共同規則：`whitefiber/docs/workorders/20261008_whitefiber_共同規則.md`，另加算力收支 Project 規則：分 A、B 兩段、每段 ≤2 小時；每步完成就更新進度檔、commit、push；需判斷的事照本單預設做，未涵蓋者列「待 Project 判斷」；**不得合併、不得 force push、不得改 main**；不得在任何對外請求（含 HTTP User-Agent）中使用使用者 email；只改 `whitefiber/`。
- 分支：`claude/whitefiber-v0.2-tk19`（chat 端已從 main `1079d5c` 開並推送本工作單）。PR 標題「WhiteFiber v0.2：每 MW 收入以 Tokenomics 為錨＋同步 v19.0」，base `main`。成品升 **v0.2**。

## 已套用的預設（判斷類；報告逐項列出替代與影響）
1. **雲端分部每 MW 年收入（100% 計費）**＝Σ 在役世代占比 × `TK_HoldEcon_<世代>` × k；k＝長約占比 × k_長約＋(1 − 長約占比) × k_現貨；k＝證據價格 ÷ 目前快照同世代成本（Oracle v0.3 `kspec` 同一組證據：k_長約＝IREN–Microsoft 9.70 ÷ `TK_HoldEcon_GB300`≈0.891，k_現貨≈1.83）。三家（CoreWeave、Nebius、Oracle）GB300 長約每 MW 年收入 9.70，WhiteFiber 必須同價。不得以 WhiteFiber 營收、RPO、合約金額反推 k。
2. **長約占比**：由 WhiteFiber 事實總帳（`data/whitefiber_facts_20261008.json`）已揭露的多年期雲端合約估計（寫明算式）；無足夠揭露時預設 70%（[Assumed]，區間 50–90%）。WhiteFiber 自身合約隱含的 k 只作敏感度。
3. **託管（colocation）分部不動**：其收入是合約租金（公司層資料），不接 Tokenomics 錨；只把它引用的 PUE、每 GPU kW 等換算改為 `TK_` 名稱。
4. **世代組合**：在役 H100／GB200／GB300 與新增世代占比取事實總帳；無揭露給區間並標 [Assumed]。
5. **不重複扣減**：收入端只用「可計費 MW ÷ 已連網 MW」收斂比例；`IF_Util` 不進收入；舊市場價路徑的 85% 隨舊方法停用。
6. **情境軸分離**：容量軸＝現行三情境；價格軸＝k 低／基準／高，三個容量情境預設都用基準 k；期初可計費 MW 校準跟價格軸重算；舊做法保留為 `legacy`，可精確重現 v0.1。新增 3 × 3 容量 × 價格目標價矩陣。
7. **資本支出**：雲端 GPU 資本支出與汰換引用 `TK_CapexIT_<世代>`；WhiteFiber 自建或購入資料中心（NC-1 等）的廠房資本支出維持公司資料，`TK_CapexFacility` 只作對照列；汰換年限引用 `TK_DeprLifeIT`；`methodology.checks.capexPerMwBand` 改引用 TK 低／高。
8. **EBITDAR 重校準**（若 WhiteFiber 有 Oracle v0.1c 的 EBITDAR 機制；沒有則維持現行 EBITDA 率路徑，報告寫明）。
9. **出租方 L 證據（Tokenomics E4）**：合約期 L≈1、殘值期才用市場租金；續約衰退本版不做（CRWV 未完成待辦），交接檔寫明。
10. **收入上限五列**（只檢查）：`IFW_RevGWFleet_100／_Util／_Prod／_Life`（比較層 `_Prod`，> 50% 警示）＋ `L1_RevGW_PQ_VR200`（模型商每 GW 實收 P×Q）對照列。

## A 段（數字變動）
0. 開 draft PR；進度檔 `whitefiber/docs/reports/20261008_whitefiber_進度.md` 新增「v0.2」段落；工具檢查（`whitefiber/CLAUDE.md` 環境第 0 步，缺什麼就補裝）；基線 `whitefiber/scripts/verify.sh`。
1. 取得 Tokenomics：`GIT_LFS_SKIP_SMUDGE=1 git clone --depth 1 https://github.com/YinchenChang/Tokenomics /tmp/tk`（或任一非 repo 路徑），確認 `model/CURRENT`。用 `tools/tokenomics/import_tokenomics.py` 產生 `whitefiber/data/tokenomics_names.txt`（以 Oracle v0.3 名稱清單為底，加審視報告第三節列出的名稱）與 `tokenomics_snapshot_v19.0.json`；Excel「Tokenomics_取數」分頁與 `TK_` 具名範圍；verify 加 `--check`。permw 檔與事實總帳中 5 處 `Inputs!`／`DC_Cost!` 引用改指名稱或加註「已被快照取代」（保留原值）。
2. 錨定與 k（移植 Oracle 的 `kspec.py`／`tkanchor.py`／`kSpecApplyQ`），Excel 公式與 HTML 引擎雙邊，cmp31 加比對項，`fields_doc.py --write`。
3. 情境軸分離＋期初可計費 MW 隨價格軸重算；`legacy` 重現 v0.1（`--vs-dist` 對 v0.1 成品 0 差異）。
4. 資本支出、汰換年限、PUE、每 GPU kW 改引用 TK；（若適用）EBITDAR 重校準。A 段末跑一次完整 verify 並等它結束，進度檔寫三情境目標價 v0.1 → 目前值。

## B 段
5. 收入上限五列；Q2 驗證（不是校準）：模型每 MW 收入 vs Q2 經常性雲端營收年化 ÷ 已連網 MW，差距拆成爬坡分母、k、世代組合、其他，加總＝總差距；> 5% 附原因，不得回頭改 k。
6. 敏感度：k_長約 0.80／1.00、自身合約 k、長約占比區間、Tokenomics 低／高成本、收斂比例；3 × 3 矩陣；三情境目標價、融資缺口、NC-1 專案貸款相關覆蓋指標。
7. 升版 v0.2：VLOG 兩處、`dist/` 換 v0.2（移除 v0.1）、交接檔（記錄 TK v19.0 與 commit、取數方式、按 MW-year 計價、錨與 k）、README、CLAUDE.md 待辦（FY26Q3 季度更新改為 v0.3，Q3 財報約 11 月中）；`--vs-dist` 對 v0.1 成品以 `--expect` 清單驗收；目標價變動依已決定事項 12 拆解（① 快照與錨取代舊推導（k＝1）→ ② 套用 k → ③ 情境軸分離 → ④ 資本支出改引用 TK → ⑤ EBITDAR 重校準（若有）→ ⑥ 上限檢查，應為 0），三情境加總＝總變動（誤差 < 0.01）。
8. 對照 Excel `whitefiber/docs/reports/20261022_whitefiber_v0.2_TK同步.xlsx`（格式比照 Oracle v0.3 對照 Excel）。
9. 報告 `whitefiber/docs/reports/20261022_whitefiber_v0.2.md`（結論三句：每 MW 收入是 TK 打平租金幾倍、與三家同價確認、主命題答案是否改變）；最後跑完整 `verify.sh` 與 `verify.sh --vs-dist`（對新 dist）並等跑完；PR 留言（第一行 `[WhiteFiber 回報] v0.2｜完成｜日期`；列出待 Project 判斷事項）、轉 ready。**不合併。**

## 驗收（chat 端）
- 快照可由 v19.0 重現；無禁用引用；GB300 長約每 MW 9.70；k 由證據 ÷ 成本算出。
- legacy 重現 v0.1；拆解加總正確；verify 全過（含離線開啟）。
- chat 端另派未參與製作的代理獨立核對後才交付 Andy。

## 停止條件
依共同規則；另：v19.0 名稱改名或移除導致讀取失敗；移植與託管分部或 NC-1 專案貸款邏輯衝突且無法不改其數字邏輯而解決——停下、push、PR 留言說明。
