# 工作單 CoreWeave W10｜Tokenomics 快照換 v18.0、世代門檻改取 Tokenomics 名稱 → v4.12（2026-10-21，r0，chat 端）

- 來源：Andy 2026-10-21「請依據最新版TK進行更新」；Tokenomics 下游同步規則（Project 索引）——擱置中的下游同步由 CoreWeave 先跑；W9 交接「v4.12 改取 `L1_GenPxThr_*` 與上市間隔名稱」。
- 類別：下游同步（換快照＋工程類取數改寫）。預估 1–1.5 小時，一段完成。代理模型：Sonnet。
- 前置：v4.11 已合併（main `b4b04dc`）。必讀：`coreweave/docs/workorders/20261007_coreweave_共同規則.md`、`coreweave/CLAUDE.md`、本單、W9 工作單（`20261015_coreweave_W9_世代指數與收入上限.md`）、Tokenomics 下游資料契約 v0.20（`docs/plan/Tokenomics_downstream_contract.md` 第 2 節與「世代價格」段）。
- 分支：`claude/coreweave-w10-v4.12`（chat 端已建立並放入本單）；PR 標題「CoreWeave v4.12：Tokenomics 快照換 v18.0、世代門檻取 Tokenomics 名稱」。

## 事實（chat 端已查證，2026-10-21）
- Tokenomics master `868e2cc`（PR #140），`model/CURRENT`＝`20261021_Tokenomics_v18.0.xlsx`。
- 以現行名稱清單（37 名）產生 v18.0 快照，與 v5.36 快照逐值比較：**26 個名稱數值不變**，包括全部成本、資本支出、折舊、維護、電力、`IF_HoldEcon`、`IF_GPUhrEcon`、`IF_Util`、`IF_WarrantyYrs`。→ 收入（跟成本走）、成本、目標價、融資缺口**預期不變**。
- 數值改變的 11 個名稱：

| 名稱 | v5.36 → v18.0（基準） | 在 CRWV 的用途 |
|---|---|---|
| `IF_RevGWFleet`、`IFW_RevGWFleet_100／_Util／_Prod／_Life` | 各世代上升，例：GB200 `_Prod` 15.29 → 26.28、H100 `_Prod` 2.24 → 5.02（US$B/GW/年） | 只用於收入上限檢查（不入損益） |
| `L1_TokMW_Gen_ratio_VR200`（q）＋`_Lo`／`_Hi` | 1.565 → 1.502；低 0.915 → 0.823；高 2.051 → 2.072 | 世代指數判讀列與 VR200 指數敏感度 |
| `L1_RevGW_Fleet_VR200` | 48.72 → 65.96 | 只在取數頁顯示 |
| `L1_GPUhr_GB300_vsBE` | 外部對照值 3.57 → 3.0345（基準值不變） | 只在取數頁顯示 |

- v18.0 已有世代門檻名稱（Tokenomics 世代價格 B 段，v7.0 起）：`L1_GenPxThr_Rev_VR200` −29.17%（低 −44.42%、高 +19.14%）、`L1_GenPxThr_Pay_VR200` −21.85%（低 −38.97%、高 +32.26%）。上市間隔在 `CAP_GenGap_VR200`（1.18；1.11–1.24），但 **`CAP_` 不在契約第 2 條第 1 項允許的前綴內**（只限 `IF_`／`IFW_`／`IFC_`／`L1_`／`SRC_`），所以 CRWV **不直接取 `CAP_GenGap_VR200`**。
- chat 端驗算：由 q 與收入門檻反推間隔 ＝ ln(1 ÷ 1.5023) ÷ ln(1 − 0.2917) ＝ 1.180，與 `CAP_GenGap_VR200` 相同；CRWV 以自身公式算回本門檻 ＝ (12.225 ÷ 10.886 ÷ 1.5023)^(1/1.18) − 1 ＝ −21.8%，與 `L1_GenPxThr_Pay_VR200` 一致。

## 步驟 1：換快照 v18.0（下游同步）
1. `coreweave/data/tokenomics_names.txt` 新增 `L1_GenPxThr_Rev_VR200`、`L1_GenPxThr_Rev_VR200_Lo`、`L1_GenPxThr_Rev_VR200_Hi`、`L1_GenPxThr_Pay_VR200`、`L1_GenPxThr_Pay_VR200_Lo`、`L1_GenPxThr_Pay_VR200_Hi`（行尾註解寫用途）。
2. 產生 `coreweave/data/tokenomics_snapshot_v18.0.json`（`TOKENOMICS_DIR` 指向 Tokenomics clone）；`git rm` v5.36 快照。逐值比對原 37 名：**只有上表 11 個名稱可以改變**，其餘 26 名差異必須為 0；不符即停下回報（共同規則第 6 節）。
3. `company.json` Tokenomics 區：`snapshotFile`、`version`（v18.0）、`commit`（完整 SHA）；`_note` 追加「W10：快照換 v18.0（26 名數值不變；四層收入上限與 VR200 q 改變）；新增 VR200 收入與回本門檻名稱」。
4. 「Tokenomics_取數」頁新增 6 個 `TK_` 具名範圍。

## 步驟 2：世代門檻改取 Tokenomics、上市間隔改為推導（工程類）
1. 判讀列「收入持平門檻」「回本門檻」改讀 `TK_GenPxThr_Rev_VR200`、`TK_GenPxThr_Pay_VR200`，並各加一列低／高端（Tokenomics 的 `_Lo`／`_Hi`）。
2. W9 的 CRWV 自算公式保留為「核對」列（同一列名加「核對｜」前綴）；兩者差距 > 0.005（0.5 個百分點）時「檢查_連動」警示。
3. **上市間隔改為推導值**：間隔 ＝ ln(1 ÷ q) ÷ ln(1 + 收入持平門檻)（兩者都是 L1_ 名稱，基準欄），標 [Derived]，備註「等於 Tokenomics `CAP_GenGap_VR200`；CAP_ 不在契約允許前綴，故由 L1_ 名稱反推」。刪除 `company.json` → `pricing.genIndex.gapYears` 的固定值 1.18（README 欄位表、輸入頁、`scripts/permw_sens.py`、segA.js 同步改讀推導值）。快照沒有門檻名稱時（例如 legacy 回歸的舊快照）間隔與門檻顯示「不適用」，不可出錯。
4. 隱含年價格變化、判讀文字（W9 步驟 2 第 5 點）改用 Tokenomics 門檻比較；基準下隱含值仍＝回本門檻（定義上相等，報告註明）。
5. 敏感度（VR200 指數 5 列）以新 q 與推導間隔重算，預期約：0.62／1.00／1.12（基準）／1.33／1.55。

## 預期變動（其他一律 0 差異）
- 版本文字、快照版本與 commit、新增名稱與列（by-label 配對）、VLOG 新列。
- 取數頁 11 個名稱的值。
- 收入上限檢查：四層比例全部下降（預期 FY26 `_Prod` 由 61.9% 降到約 35–40%，結果文字可能由「警示」變「通過」）。
- 世代指數判讀列：q、間隔（數值仍 1.18，來源改為推導）、隱含年價格變化（約 −24.5% → −21.9%）、兩個門檻、低高端列。
- VR200 指數敏感度表（指數值、隱含價格、三情境目標價、融資缺口）。基準列（跟成本走）必須仍等於 $42.70／$31.99／$14.72。
- **三情境目標價、融資缺口、每 MW 收入與成本各期值：0 差異。**

## 驗收
- verify 三組全過：(1) `scripts/verify.sh`；(2) 升版對 v4.11 成品 `--vs-dist`，EXPECT 只含上節所列；(3) legacy 回歸對 v4.6 維持 0 差異。
- `import_tokenomics.py --check coreweave/data/tokenomics_snapshot_v18.0.json` 通過。
- 報告必須有一節「Tokenomics 變動造成的差異」（下游同步規則）：逐項列出 11 個名稱新舊值 → 影響到的 CRWV 列與新舊值；結論句寫明目標價是否改變。

## 升版
v4.12：`vlog.py` 與 `tail.js` 的 VLOG 各加一列（目標價欄 $32.0；變動欄「三情境目標價不變；拆解 (a)(b)(c)(d) 皆 0；Tokenomics 快照換 v18.0（上限檢查與 VR200 q 改變）、世代門檻改取 Tokenomics」）；`dist/` 換 v4.12 並移除 v4.11；交接檔以 v4.12 取代 v4.11（寫明 Tokenomics v18.0 與 master SHA、門檻取數方式、間隔推導、CAP_ 契約缺口）；README 同步。

## 回報
- 報告 `coreweave/docs/reports/20261021_coreweave_v4.12_Tokenomics_v18.0.md`：結論先行；「Tokenomics 變動造成的差異」表；四層 × 各期上限比例新舊對照；判讀列新舊；敏感度表新舊；verify 摘要；已套用的預設；可移植性。
- 「待 Project 判斷」至少列一項：回報 Tokenomics——契約第 2 條第 1 項的允許前綴不含 `CAP_`，但契約又要求下游讀 `CAP_GenGap_*` 且不得覆寫；建議 Tokenomics 新增 `L1_GenGap_*` 別名或修訂第 2 條。
- PR 留言第一行 `[CRWV 回報] W10｜完成｜2026-10-21`；PR 改 ready。
- 「需要 Andy 做的事」：合併前用真正的 Excel 開啟 v4.12。
- 代理不得合併；不得在任何對外請求（含 HTTP User-Agent）中使用使用者 email；verify 等長命令背景執行、不輪詢，**必須等長命令跑完再回報**（代理結束時背景程序會被終止）。
