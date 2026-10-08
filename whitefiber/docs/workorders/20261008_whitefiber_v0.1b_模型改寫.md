# 工作單 WhiteFiber v0.1b｜模型改寫（2026-10-08，r0，chat 端）

- 前置：v0.1a 已完成並經 chat 端審查（審查留言在 PR，與本工作單同等效力）。
- 必讀：`20261008_whitefiber_共同規則.md`、v0.1a 資料報告、對照表 r1 第 3、4 節、`whitefiber/README.md`（引擎結構與欄位說明）。參考：`oracle/docs/reports/20261007_oracle_進度.md` 的 v0.1b／v0.1c 步驟紀錄（同一引擎的改寫路徑：傳統事業區塊、預付、租賃、分部 EV/EBITDA、穩態汰換、EBITDAR）與 Oracle v0.2 報告（延誤模組、debtCapBasis）。
- 本張改引擎與 `company.json`，**不改 `dist/`、不升版**。

## 目標
把引擎改成 WhiteFiber 的因果結構（對照表 r1 第 3 節），HTML 與 Excel 數字一致，`verify.sh` 全過。

## 實作原則
1. **雙引擎同步**：新邏輯 Excel 公式與 HTML 引擎兩邊都實作，以 Excel 為準，cmp31 新增比對項涵蓋每個新區塊。
2. **不刪通用程式**：WhiteFiber 沒有的機制（強制轉換特別股、股利、傳統事業、投資級口徑等）輸入設 0、空清單或改選項，畫面顯示「不適用」，不刪除通用程式碼。
3. 公司特有內容一律放 `company.json`；新增欄位同步 `scripts/fields_doc.py --write`。新增首期一次性金額或期初餘額時，加入 `calendar_q.py` 的 `ROLL_FIELDS` 並在 `asOf` 標記季度。
4. 改名用 `rename_ids.py`；遵守 `whitefiber/CLAUDE.md`「勿改」8 條（注意 Oracle CLAUDE.md 待辦 2b：「輸入與假設」D 區重複表頭，新列避免放該頁 A 欄同區段重名）。

## 步驟（每步完成即跑 verify.sh、commit＋push、更新進度檔；verify 未過不得進下一步；每步記錄三情境目標價）
0. 進度檔新增「v0.1b」段落；先處理 v0.1a 審查留言交代的事項（若有）。
1. **欄位搬移與換名**：`whitefiber` 區段資料搬入既有欄位（`meta`〔company＝WhiteFiber、ticker WYFI〕、`calendar`〔財年結束月 12、latestQuarterFiled 2026Q2〕、`asOf`、`ytdActual`〔1H26〕、`historicalPL`、`latestQuarter`、`callFacts`、`debt`、`leases`、`valuation`〔現價取 v0.1a 共識檔〕、`quarterly`〔焦點季 2026Q3〕、`texts`、`meta.consensusFile`）。刪除 `oracle` 區段與 Oracle 資料檔（`oracle_facts_*`、`consensus_orcl_*`；`consensus_crwv_*` 若仍被同業引用則保留）。`peers.list`：CRWV、NBIS＋託管同業（v0.1a）。畫面文字中描述本公司的 Oracle／CoreWeave／Nebius 全部換成 WhiteFiber。確認首期長度＝0.5 年、期間標籤 FY26–FY30（模型期＝首期＋4 個完整財年）。Oracle 專屬機制（傳統事業四線、股利、強制轉換特別股、OpenAI 實現率）設為不適用（輸入 0／空）。
2. **雲端營收主軸**：`scenarios.mwPath`（期初雲端 IT MW 取 v0.1a；三情境依對照表 r1 第 4 節第 1 條，雲端合約 GPU 數換算 MW 與起始時點）、`scenarios.revMW`（11.62／17.40／24.20 或 v0.1a 更新值）、可計費收斂。**首期校準**：以 Q2 經常性雲端營收（扣除 $12.3M 一次性）年化校準期初可計費 MW（converge 模式），寫明是對齊已實現實際數。首期實際 1H26 營收照 10-Q（含一次性）。合約隱含單價列對照（運營頁與檢查頁）。
3. **託管分部（新）**：建 MW 驅動的託管區塊（可擴充傳統事業多線區塊或另建，擇一並說明）：逐站點／逐批（NC-1 第一批 40 MW、MTL 站點、三情境的擴建批次）輸入：IT MW、起租時點、每 MW 年租金（已簽約者照合約，含年調 3%；未簽約者取 v0.1a 託管租金區間：保守低／基準中／積極中）、EBITDA 率（v0.1a 推得或 [Analogy]）、每 MW 建置成本與支出時點、建物折舊年限。產出託管營收、EBITDA、建置 CapEx（進類現金流量與 PP&E）、建物 D&A（另列，不與 GPU 壽命混用）。轉嫁電費預設淨額（見對照表）。託管 RPO 對照列（合約排程 vs 10-Q RPO $932.9M）。
4. **GPU 取得與租賃**：依 v0.1a 自購／租賃比例，雲端新增 MW 的 GPU：自購部分走既有 CapEx 與 GPU 折舊；租賃部分以未起租租約模組（或 GPU 租金列）轉為固定租金，並以 EBITDAR 機制處理（租金扣 EBITDA）。既有租賃到期表照 10-Q。`node scripts/calib_ebitdar.js --write` 依新口徑校準並在報告說明。
5. **客戶出資（預付）**：`defaults.prepay` 改為 WhiteFiber：期初合約負債＝10-Q 遞延營收（託管與雲端分開若有）、新合約預付覆蓋比（v0.1a 推得，否則 [Assumed] 區間）、認列年數；重大財務組成：10-Q 有揭露則照揭露，否則利率 0 並列敏感度。
6. **融資與既有證券**：債務清單（RBC 聯貸、Bit Digital 貸款〔可用額度計入〕、設備與過橋融資、NC-1 專案融資〔基準與積極情境假設完成：金額依 v0.1a 或建置成本 × 貸款比率 [Assumed] 60–70%；保守情境不完成〕）；2026-08 可轉債以可轉債模組承接（轉換價、利率、到期；上限買權若有，在報告說明對稀釋的影響，模型以轉換價計稀釋）。瀑布：預付 → 現金（最低現金 $30M）→ 資產層／專案債（`defaults.debtCapBasis`＝`ebitda`、上限 5.0×，利率取 v0.1a）→ 可轉債（現市值 10%／年）→ ATM 股權（現市值 10%／年）→ 高息債溢出。一頁摘要的「需失去投資級才能融資的金額」一句改為通用文字（例如「超出可融資上限、需高息債的金額」），文字放 `company.json`。不發股利（股利輸入 0）。
7. **建設延誤**：`scenarios.delayMonths` 保守 6／基準 3／積極 0；`delayLink` 依租約性質（預設 0.5）；罰則 0。
8. **評價**：WACC 依 v0.1a CAPM 輸入計算並寫入 `valuation`（列出計算式）；EV/EBITDA 腿分部加總：託管 EBITDA × 託管同業 NTM 中位數＋雲端 EBITDA × 6×（錨定 FY29）；DCF、反向 DCF（穩態 EBITDA 率格線維持雲端口徑並說明）、季度層、期間滾動、共識對照、一頁摘要全部改用 WhiteFiber 欄位。差距超過 5% 的共識項目附差異原因。
9. **整體 verify 與回報**：`verify.sh` 全過；報告 `whitefiber/docs/reports/20261008_whitefiber_v0.1b_模型.md` 列：三情境目標價與區間、雲端每 MW 3×3 敏感度、合約隱含單價取代 Tokenomics 單價時的目標價、託管租金與建置成本敏感度、Nscale 違約（託管營收歸零）目標價、NC-1 專案融資未完成時的缺口與稀釋、債務上限 4×／6× 敏感度、已套用的預設、未解問題。

## 驗收
- `verify.sh` 全過（HTML 與 Excel 一致、重算錯誤 0、verify_ooxml、離線開啟）。
- 雲端營收可追溯為 MW × 每 MW × 可計費比；每 MW 來源是 Tokenomics 推導檔，不是公司數字。託管營收可追溯到合約排程或 [Analogy] 區間。
- FY26 合計營收＝1H 實際＋2H 模型；與共識差距都有原因。
- 畫面文字（`crawl.py`）描述本公司的 Oracle／CoreWeave／Nebius 為 0（同業、對照、來源、版本紀錄除外，逐項列出）。
- 進度檔「下一步」寫好 v0.1c 起點。

## 停止條件
依共同規則第 6 節。
