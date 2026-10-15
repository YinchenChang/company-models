# 工作單 CoreWeave W9｜收入改 q × p 結構（VR200 世代指數）＋收入上限檢查改四層 → v4.11（2026-10-15，r0，chat 端）

- 來源：CoreWeave 對話串交接檔（2026-10-15）待辦 B2、B3；Tokenomics 回覆 C 題（Project 文件 `claude/TK回覆CoreWeave_C題_20261015.md`；Tokenomics 報告 `docs/reports/20261015_世代價格A_查證.md`，PR #49）；Andy 2026-10-15「依建議」（選項甲：v4.11 先做 B3＋B2 結構，門檻相關的上市間隔先寫固定值；v4.12 等 Tokenomics B 段合併後改為取 `L1_GenPxThr_*` 名稱）。
- 類別：方法變更，但**基準數值設計為 0 變動**（理由見 B2）；唯一預期的檢查結果變動是 B3 的上限檢查比例。預估 1.5–2.5 小時，一段完成（兩個步驟各自 commit）。代理模型：Sonnet。
- 前置：v4.10 已合併（main `85e6b85`）。必讀：`coreweave/docs/workorders/20261007_coreweave_共同規則.md`、`coreweave/CLAUDE.md`、本單、W8 工作單（換快照前例）、W4 工作單（收入錨與上限檢查的原始做法）、Tokenomics 下游資料契約 `docs/plan/Tokenomics_downstream_contract.md`（第 2、3、4 節）。
- 分支：`claude/coreweave-w9-v4.11`（chat 端已建立並放入本單）；PR 標題「CoreWeave v4.11：VR200 世代指數（q × p）與收入上限四層」。

## 事實（chat 端已查證，2026-10-15）
- Tokenomics master `63db13a`（PR #49 合併；只新增報告），`model/CURRENT`＝`20261015_Tokenomics_v5.36.xlsx`。v5.36 對 CRWV 現有 29 個名稱數值不變（chat 端 15:30 判定；本單第 0 步由代理以工具逐值確認）。
- `import_tokenomics.py` 已支援 `IFW_`、`IFC_`（PR #40 時尚未支援，之後已升級）。
- v5.36 Interface（基準成本情境；US$B/GW/年）：

| 名稱 | H100 | GB200 | GB300 | VR200 | Rubin Ultra |
|---|---|---|---|---|---|
| `IF_HoldEcon` | 9.703 | 8.828 | 10.886 | 12.225 | 15.354 |
| `IFW_RevGWFleet_100` | 6.675 | 33.624 | 43.775 | 81.206 | 114.879 |
| `IFW_RevGWFleet_Util`（＝現用 `IF_RevGWFleet`） | 4.005 | 20.175 | 26.265 | 48.724 | 68.928 |
| `IFW_RevGWFleet_Prod` | 2.243 | 15.285 | 19.232 | 38.115 | 54.418 |
| `IFW_RevGWFleet_Life` | 2.243 | 15.285 | 19.232 | 38.115 | 54.418 |

- `IFC_Use` 上述四列：「情境、相對比較、反轉門檻；上限，不得作預測」→ CRWV 只用作檢查，不作收入輸入（符合）。
- `L1_TokMW_Gen_ratio_VR200`（VR200 對 GB300 每 MW token 產出比 q）：基準 1.565、低 0.915、高 2.051（`_Lo`／`_Hi`）。`L1_TokMW_Gen_ratio_GB300`（對 GB200）1.456。
- Tokenomics 回覆的門檻（固定模型口徑）：GB300→VR200 上市間隔 1.18 年（1.11–1.24，[Analogy]）；收入持平門檻 −32%／年；回本門檻 −25%／年。chat 端以上表驗算：收入持平 ＝ (1 ÷ 1.565)^(1/1.18) − 1 ＝ −31.6%；回本 ＝ (12.225/10.886 ÷ 1.565)^(1/1.18) − 1 ＝ −24.5%，與 Tokenomics 一致。
- 現行收入式：每 MW 年收入｜世代 g ＝ `IF_HoldEcon`｜g × k（k＝k_長約 0.891，隨需占比 0%）。所以 VR200 每 MW 收入 ÷ GB300 ＝ 12.225 ÷ 10.886 ＝ **1.123**（「跟著成本走」），代表的年價格變化恰好等於回本門檻（定義上相等）。

## 步驟 0：換快照 v5.36 並新增名稱（工程類）
1. `coreweave/data/tokenomics_names.txt` 新增：`IFW_RevGWFleet_100`、`IFW_RevGWFleet_Util`、`IFW_RevGWFleet_Prod`、`IFW_RevGWFleet_Life`、`IFC_Use`、`L1_TokMW_Gen_ratio_VR200`、`L1_TokMW_Gen_ratio_VR200_Lo`、`L1_TokMW_Gen_ratio_VR200_Hi`（行尾註解寫用途）。`IF_RevGWFleet` 保留（＝`_Util` 層，供對照；若 `_Util` 與之逐值相同，報告註明）。
2. 以 v5.36 產生 `coreweave/data/tokenomics_snapshot_v5.36.json`；`git rm` v5.35 快照。逐值比對原 29 名：**差異必須為 0**，否則停下回報（共同規則第 6 節）。
3. `company.json` Tokenomics 區：`snapshotFile`、`version`、`commit`（完整 SHA）；`_note` 追加「W9：快照換 v5.36（原 29 名數值與 v5.35 相同）；新增四層瀑布、IFC_Use、VR200 產出比 q」。
4. 「Tokenomics_取數」頁新增對應 `TK_` 具名範圍；`IFC_Use` 的用途文字顯示在四層列旁。

## 步驟 1：B3 收入上限檢查改四層（判斷已定：比較層＝`_Prod`）
1. 「每MW經濟性」上限檢查區：四層（`_100`、`_Util`、`_Prod`、`_Life`）各一列「客戶每 MW 付費 token 營收（在役世代加權）」，並各一列「CRWV 每 MW 計費收入 ÷ 該層」。
2. **門檻比較層改為 `_Prod`**（契約第 2 條：不得只取單一層；`_Prod` 已扣 SLO 下的產能折減，是四層中最接近可實際售出 token 的一層）；門檻維持 `methodology.checks.revCapShareMax`＝50%。比較層放 `company.json`（新欄位，例如 `methodology.checks.revCapLayer`＝`"Prod"`，README 欄位表同步），程式不寫死。
3. 「檢查_連動」的上限檢查列改讀 `_Prod` 比例；HTML 連動檢查同步（只顯示 Excel／建置時結果，不新增 JS 計算，已決定事項 8）。
4. 預期：FY26 比例由 45.3%（`_Util`）升到約 60–65%（`_Prod`）→ **警示**；FY27 約 50% 上下。這是檢查結果，不改任何收入、成本、目標價。報告列出四層 × 各期比例表。
5. 已套用的預設（報告逐項列）：比較層 `_Prod`｜替代 (a) `_Util`（原做法，全期通過）、(c) `_100` × CRWV 計費利用率｜影響：只影響警示，不影響數字。

## 步驟 2：B2 收入改 q × p 結構（基準 0 變動）
### 公式
- 參考價 p_ref ＝ k × `IF_HoldEcon`｜GB300（目前成本情境）。基準＝0.891 × 10.886 ＝ 9.70（＝IREN 長約價；k 的定義使兩者相等）。
- 世代指數 idx_g：每 MW 簽約收入相對 GB300。
  - **預設（空白）＝跟成本走**：idx_g ＝ `IF_HoldEcon`｜g ÷ `IF_HoldEcon`｜GB300（同成本情境）。
  - VR200 可覆寫：`company.json` 新欄位（例如 `pricing.genIndex.VR200`），基準留空（＝跟成本走 1.123）；有覆寫值時用覆寫值。H100、GB200、GB300、Rubin Ultra 只用預設，不開覆寫（GB300 恆為 1）。
- 每 MW 年收入｜g ＝ p_ref × idx_g。在預設下與現行式恆等，所以**全部數值 0 差異**；若不是 0，代表公式寫錯，停下修正。
- 已決定事項 16(3) 不變：日後有 VR200 長約價時，覆寫值＝VR200 每 MW 長約價 ÷ p_ref（證據入 `pricing.anchorMultiple` 證據表）。

### 判讀列（「每MW經濟性」新增，次層／預設收合）
1. q（`L1_TokMW_Gen_ratio_VR200`，基準／低／高）。
2. 上市間隔 GB300→VR200：`company.json` 新欄位（例如 `pricing.genIndex.gapYears`＝{base 1.18, low 1.11, high 1.24, tag "[Analogy]", source "Tokenomics 世代價格 A 段查證（docs/reports/20261015_世代價格A_查證.md，PR #49）", note "v4.12 改取 Tokenomics 名稱"}）。**這是本版唯一的固定值**，備註寫明待 Tokenomics B 段合併後換成名稱。
3. VR200 指數隱含的年價格變化 ＝ (idx ÷ q)^(1 ÷ 間隔) − 1。
4. 收入持平門檻 ＝ (1 ÷ q)^(1 ÷ 間隔) − 1；回本門檻 ＝ (`IF_HoldEcon`｜VR200 ÷ `IF_HoldEcon`｜GB300 ÷ q)^(1 ÷ 間隔) − 1。兩者以活公式算，不寫死 −32%／−25%；報告中與 Tokenomics 回覆值對照（差距 ≤ 1 個百分點即視為一致）。
5. 判讀文字（依條件產生，不寫死）：隱含年價格變化 ≥ 回本門檻 →「VR200 回本餘裕不低於 GB300」；介於兩門檻之間 →「每 MW 收入高於 GB300，但回本餘裕較低」；< 收入持平門檻 →「每 MW 收入低於 GB300」。基準下隱含值＝回本門檻，報告註明「定義上相等」。

### 敏感度（建置時快照，沿用現有每 MW 敏感度的做法與「快照已過期」檢查）
VR200 指數 5 列 → 三情境目標價（保守、基準、積極）、加權目標價、基準融資缺口：
| 列 | VR200 指數 | 意義 |
|---|---|---|
| 1 | q 低 × (1 + 回本門檻)^間隔（約 0.66） | 產出比取低端，價格同基準 |
| 2 | 1.00 | 收入持平（契約「持平情境」） |
| 3 | 基準（跟成本走，約 1.12） | ＝回本門檻 |
| 4 | q × (1 − 10%)^間隔（約 1.38） | 世代間只按壽命內降價 g −10%／年 |
| 5 | q 高 × (1 + 回本門檻)^間隔（約 1.47） | 產出比取高端，價格同基準 |
指數值以建置時 Python 依快照與 company.json 算出，不寫死；表內同時列每列的隱含年價格變化。

## 不做
- B4（`IFC_Use` 用途標註的精度提醒）：Tokenomics 尚未說明。
- `L1_GenPxThr_*`、上市間隔名稱：Tokenomics B 段尚未建置（v4.12）。
- 不改 k、隨需占比、世代組合、成本、資本支出。

## 驗收
- verify 三組全過：(1) `scripts/verify.sh`；(2) 升版對 v4.10 成品 `--vs-dist`，EXPECT 只含：版本文字、快照版本與 commit、新增名稱與列（by-label 配對）、VLOG 新列、上限檢查比較層與比例（`_Util` → `_Prod`）及其結果文字；**其他所有數值 0 差異**（含三情境目標價、融資缺口、每 MW 收入各期值）；(3) legacy 回歸對 v4.6 維持 0 差異（同 W6–W8）。
- `import_tokenomics.py --check coreweave/data/tokenomics_snapshot_v5.36.json` 通過（須設 `TOKENOMICS_DIR`）。
- 三情境目標價不變：基準 $31.99、保守 $42.70、積極 $14.72。
- 收入持平與回本門檻活公式結果與 Tokenomics 回覆差距 ≤ 1 個百分點。

## 升版
v4.11：`vlog.py` 與 `tail.js` 的 VLOG 各加一列（目標價欄 $32.0；變動欄「三情境目標價不變；拆解 (a)(b)(c)(d) 皆 0；收入改 q × p 結構（VR200 指數預設跟成本走）、上限檢查改 _Prod 層」）；`dist/` 換 v4.11 並移除 v4.10；交接檔以 v4.11 取代 v4.10（寫入 q × p 結構、上市間隔待換名稱、比較層 `_Prod`、Tokenomics 版本 v5.36 與 master SHA）；README 同步。

## 回報
- 報告 `coreweave/docs/reports/20261015_coreweave_v4.11_世代指數與收入上限.md`：結論先行；四層 × 各期上限比例表；VR200 判讀列各期值；敏感度表；verify 摘要；已套用的預設；可移植性（新增的 company.json 欄位、換公司時要改什麼）。
- PR 留言第一行 `[CRWV 回報] W9｜完成｜2026-10-15`。
- 「需要 Andy 做的事」：合併前用真正的 Excel 開啟 v4.11。
- 代理不得合併；不得在任何 HTTP 請求放入使用者 email；長命令背景執行、不輪詢。
