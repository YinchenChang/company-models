# CHANGELOG

每次同步記錄：Excel 版本、工作包、commit 與變動摘要。

## 20261007_OpenAI_v0.6.xlsx — v0.6-P3（S3：算力需求與供給）

- 新增 Compute 頁（G01–G133）：V1 世代組合（Blackwell 拆 GB200／GB300；自研＝VR200 × genFactor.custom）、付費／免費每 GW 年產能（TK `IF_TokGW_*` × `IF_Util`，Tokenomics Alloc 同式）、推論 GW（token 換算）、V2 η（2025 推論支出 ÷ 合約價校準；基準沿用，情境線性回升至 2030＝1）、有效推論 GW、三角對照、研發 GW（v0.5 殘差法）、總需求 GW、供給 GW（逐合約；v0.5 S2、N4(b)）、容量上限（V1a）、VR 等值換算、敏感度（合約價、η 路徑、GB300 占比、自研產能）。Revenue 新增第八節 R46–R52（容量上限係數、截頂後總額／分成／淨額；未截頂值保留）。
- Inputs 243→249（INP_256–261）；Checks 新增 C40–C66（C05 期望 243→249）；具名範圍 516→562（`CMP_*`、`REV_*Capped`）；公式格 1,260→1,989。SRC_OAI 不變。
- 主要結果：η（2025）0.2795；有效推論 GW 2025 0.70 → 2030 5.60；合約供給 1.23 → 13.01 GW；2026–2030 皆未截頂（2030 餘裕 46%）；2025 支出換算總需求 1.70 GW 對年均 1.25（+0.45）；三角對照 16.32 $B/GW/年（+36%）。
- 測試：parity 66 項全過（新增情境 p–w；E6 掃描擴及 Compute）；範圍外變動 0。C46 以最大絕對值寫法避開 pycel `SUMPRODUCT(ABS())` 型別差異。S2 版改名移至 `model/archive/20261007_OpenAI_v0.6-P2.xlsx`。報告：`docs/reports/20261007_v0.6-P3.md`、`_對照.xlsx`。PR #14。

## 20261007_OpenAI_v0.6.xlsx — v0.6-P2（S2：需求與營收）

- 新增 Demand（D01–D48）、Revenue（R01–R45＋第七節價格事件時點表）兩頁；機制沿用 v0.5「訂閱與廣告」「API」「營收彙總」頁。訂閱（T4 遷移、pro 依 V12／V15）、任務 × 每任務 token、API（2025 由營收倒推；2026 牌價依價格事件時點天數加權；2027 起年變動與價格彈性）、廣告、其他、總額、Microsoft 分成、淨額；token 量（層級 × 付費／免費）具名範圍 `DEM_Tok_*` 供 P3。
- SRC_OAI 106→122（價格事件數值與公告日，SRC_OAI_113–128）；Inputs 233→243（E8c 事件時點偏移 INP_246–249、E6 定義常數 INP_250–255）；Checks C27 WARN 撤除（退役），新增 C28–C39；具名範圍 426→516；公式格 652→1,260。
- 主要結果：2025 營收對 SRC 實際差距 0（校準），未校準 +1.03B；2026 總額 34.20B；2030 總額 126.38B、淨額 126.38B（分成上限 2029 觸頂）；2025 token 3,538T＝TK `IF_AllocDemand` 的 0.84 倍。
- 測試：parity 49 項全過（新增情境 h–o；E6 掃描擴及 Demand、Revenue）；範圍外變動 0。S1 版改名移至 `model/archive/20261007_OpenAI_v0.6-P1.1.xlsx`。報告：`docs/reports/20261007_v0.6-P2.md`、`_對照.xlsx`。PR #10。

## 20261007_OpenAI_v0.6.xlsx — v0.6-P1.1（S1：TK_Link 快照更新至 Tokenomics v5.24）

- 快照來源：Tokenomics master `bdb0de7`（CURRENT＝`20261007_Tokenomics_v5.24.xlsx`）。check_tk_snapshot OK 37→62、MISSING 11→0、DIFF 0。原待合併 11 名全部取得；新增 IF_TrainCost_*、IF_RevGW_*、IF_FullCostDefault_*、L1_Ans3（＋低高）、SRC_DEM_013_Lo／Hi。原 37 名中 10 名值變動（TokGW×3、HoldEcon、FullCost×3、ProgGWyr×3）；第 5 世代口徑 Rubin Ultra NVL576→MGX NVL 單架 72 GPU。
- 只動 TK_Link、Checks（C07、C08 標籤與期望值；ID 不變）、README 兩格；具名範圍 401→426；SRC_OAI、Inputs、Derived_V9、Map_v05 逐格不變。舊檔移至 `model/archive/`。
- 工具：`tools/diff_tk_snapshots.py` 新增；`builder/tk_link.py`、`build.py` 沿用舊 PR #2（`479e94b`）改動。parity 32 項全過。報告：`docs/reports/20261007_v0.6-P1.1.md`、`_對照.xlsx`。PR #8。

## S0 搬遷（2026-10-07，Excel 不變）

- 由 `YinchenChang/openai-model` `main@13c16f0` 搬入 `company-models/openai/`；CI 移至根目錄 `.github/workflows/openai-parity.yml`（working-directory `openai`）。新增共同規則、S1–S6 工作單、進度檔、交接檔（`docs/handoff/OpenAI_handoff.md`，取代 Project 內舊交接檔）。parity 32 項全過，Excel 逐位元組相同。

## 20261004_OpenAI_v0.6.xlsx — v0.6-P1（repo、Source 與 Tokenomics 連結）

- 報告：`docs/reports/20261004_v0.6-P1.md`。Commit：見本 PR 的合併提交（合併後補上雜湊）。
- 新增：repo 骨架（builder／engine／tests／tools／CI）、CLAUDE.md；Excel 8 頁：README、SRC_OAI（106 列）、TK_Link（Tokenomics v5.14 快照，37 個名稱有值、11 個待 v5.15 合併）、Inputs（233 列）、Derived_V9、Checks（26 項，編號凍結）、Map_v05（849 個 v0.5 葉節點去處）、_Defaults（隱藏）。穩定 ID registry（`builder/id_registry.json`）。
- 測試：parity 24 項全過；v0.5 每個葉節點有去處；`check_tk_snapshot`：OK 37、MISSING 11（WARN）。
- 依工作單 r2：V6（SRC 10.59、訓練支出改公式、其他雲端 1.41／0–2.91）、V8（合約價單一列 12／8–20）、V9（跨年拆開）、V10（四欄初評）、V11；E4（PR）、E5（環境安全）。詳見報告第 2 節。
- 依工作單 r3（內容依 chat 訊息摘要，r3 原文未在 main）：E6（V9 公式不含常數、換算係數進 Inputs）、E7（帳目四項）、V12（pro 改正）、V13（Cerebras／Azure 終點與總額 Assumed＋區間）。V14、E8 待 r3 原文。詳見報告。
- 依工作單 r3 補做與 V15：V14／F2（來源初評規則）、E8 a–j（待判斷 A 類全數處理）、F1（Checks 編號 registry）、V15（pro 2027–2030 改公式）。詳見報告。
- 審查補充（G1–G8、V16）：合約列立場初評（未評 11→5）、Checks C27 WARN、定義常數註記、Azure 起點區間 2025–2026。詳見 `docs/reports/20261004_v0.6-P1_supplement.md`。
- 審查補充二（H1–H4）：SRC_OAI_105–108 補出處（OpenAI 與 NVIDIA／Broadcom／AMD 聯合公告）、「同上」出處追溯、一手／二手圖例、SRC_OAI_082 立場說明。V10：立場未評 1（082）、等級未評 0、一手／二手未評 0。
