# OpenAI 收支模型 — Claude Code 工作規範（v0.6；以 Tokenomics CLAUDE.md 修訂四為底改寫）

> **2026-10-07 搬遷與流程更新（優先於本檔其餘各節）**
> - 本模型已由 `YinchenChang/openai-model`（停用）搬入 `YinchenChang/company-models` 的 `openai/` 資料夾（來源 `main@13c16f0`）。本檔所有路徑都相對於 `openai/`；指令一律在 `openai/` 內執行。預設分支 `main`；CI＝repo 根目錄 `.github/workflows/openai-parity.yml`（合併門檻 job 名稱 `parity`）。
> - 流程改為自主建置：chat 端寫工作單（`docs/workorders/`）→ 建置代理執行 → chat 端審查並合併；Andy 只在里程碑審查。開工前必讀 `docs/workorders/20261007_openai_共同規則.md`。
> - **判斷類**：不再停下來等 Project 端。依序套用：該張工作單寫明的預設 → 工作單 r6 第 2 節 V1–V16 → 共同規則第 5 節（最保守且最常見的口徑），並在報告「已套用的預設」逐項列出。第 1 節「判斷類只能在 Project 端決定」「分不清時不要動手」兩句，以此取代。
> - 停止條件只有共同規則第 6 節所列各項。

> 工作單說底稿為「修訂五」；開工時 Tokenomics master 的 CLAUDE.md 為「修訂四（2026-10-01）」，本檔以它為底。若修訂五另有內容，請 Project 端指出差異。

## 1. 角色與事實來源

- **Excel 活頁簿是唯一事實來源**：`model/YYYYMMDD_OpenAI_vN.xlsx`（現行檔名見 `model/CURRENT`）。所有計算機制與數值以 Excel 為準。
- Python（builder、engine、tools）只能**產生結構**或**執行／呈現** Excel；不得自行新增、修改或省略任何計算機制與參數，不得在 Python 硬編碼價格、規格、比例等參數。
- **修改分兩類**：
  - **判斷類**：只能在 Project 端（Andy 與 Claude chat）決定。範圍：輸入值、機制或公式邏輯、新參數、來源與標記、區間、分類（SRC_OAI／Inputs／公式／不遷入）。工作單沒寫清楚者，列入報告「待 Project 判斷」，能做的先做、不能做的保留現狀。
  - **工程類**：CC 可自行決定並在報告說明。範圍：具名範圍、標籤與單位文字、格式、僅供顯示的列、工作表順序、不改數值的 builder 重構、測試與 CI。
  - **分不清屬哪一類時，一律視為判斷類**：在報告中提出，不要動手。
- **Excel 優先**：藍字輸入（SRC_OAI 的值／低／高、Inputs 的值／低／高）由 Excel 擁有。builder 重建時只在「Excel 值仍等於上次 builder 寫入的預設值」才改寫（隱藏頁 `_Defaults` 記錄）；已被改過的保留並記入 `restore_log.txt`。TK_Link 的值例外：由 builder 從 Tokenomics 快照重寫（更新快照＝本模型的修補版）。
- **（E6）公式不含常數**：跨年混合（V9、V12）等公式內不得有常數（算式中的恆等式除外：`1-比例`、年數含頭尾的 `+1`）；換算係數與比例全進 Inputs（標記＋區間），`test_builder.py::test_e6_no_constants_in_v9_formulas` 檢查。
- **穩定 ID**：SRC_OAI_nnn、INP_nnn 一經發出即固定（`builder/id_registry.json`）；新列取下一個號碼，被移除的號碼退役、不重用（Excel 優先的保留機制與審查都以 ID 為鍵）。
- **資料分層**（Tokenomics 治理 G1、G8、G9）：公司財務原始數據→`SRC_OAI`；AI 技術與算力的原始數據與推導→不在本模型重算，一律取自 Tokenomics 名稱（`TK_Link`）；假設→`Inputs`（藍字，附標記與區間）。**分層優先於標記**。
- 本模型公式只引用 `TK_` 具名範圍，不直接寫 Tokenomics 數值；不使用 Excel 外部連結。

## 2. 目錄

```
model/            現行 xlsx（唯一一份）；model/archive/ 放舊版；model/CURRENT 一行記錄現行檔名
builder/          產生 Excel 的程式（v05_map.py 遷移對照、source_rules.py 來源四欄初評規則、tk_link.py 快照讀取、build.py 組裝、preserve.py Excel 優先、id_registry.json 穩定 ID）
engine/           以 pycel 載入 xlsx 並重算（不手抄公式）
tests/parity/     Excel 與 engine 的一致性測試與情境檔
tools/            check_tk_snapshot.py、export_csv.py、build_engine_cache.py、ci_summary.py
reference/v0.5/   v0.5 JSON 與 Excel（Andy 上傳；唯讀參考）
docs/workorders/  工作單；docs/reports/  每個工作包的報告
```

### engine 規則

- 以 pycel 載入並重算；每次建構都強制全簿重算。選型標準：parity 全過，且增量重算與全簿重算各少於 2 秒（硬性）。
- 遇到引擎不支援的函數：改用 pycel 支援的等價函數寫在 Excel（工程類），並在報告說明；不得在 Python 另寫旁路計算。（P1 例：COUNTA 不支援，改為 `SUMPRODUCT(--(LEN(範圍)>0))`。）
- 輸入與輸出一律透過 Excel 具名範圍存取。

## 3. 一致性測試（必須通過才可合併）

1. LibreOffice headless 重算 xlsx，取得期望值（環境需有 `libreoffice-calc-nogui`；只有 `soffice --version` 能執行不代表 Calc 已安裝）。
2. engine 計算同一組情境（`tests/parity/scenarios.yaml`），全部公式格相對誤差 ≤1e-9、字串完全一致、任何情境不得出現錯誤值。
3. 結構性：工作表、具名範圍、`model/CURRENT`、SRC_OAI／Inputs 筆數與報告一致；v0.5 每個葉節點有去處（`test_builder.py`）。
4. GitHub Actions 每次 push 與 PR 執行（`parity` 為合併門檻）。`tk-snapshot` 為 WARN，不擋合併。

## 4. 建置與同步（每個工作包）

1. `python3 builder/build.py --tk-dir <Tokenomics checkout> --out model/<新檔名>.xlsx --base model/<現行>.xlsx`；舊版以 `git mv` 移入 `model/archive/`；更新 `model/CURRENT`。
2. 執行 parity 與 `tools/check_tk_snapshot.py`；更新 `tests/parity/scenarios.yaml` 的期望值（有意識地）。
3. 在 `CHANGELOG.md` 記錄 Excel 版本、工作包（`v0.6-P1`…`v0.6-P5`）、commit 與變動摘要。
4. 發現 Excel 本身的錯誤（`#REF!`、`#DIV/0!`、循環參照、單位不一致）：列在報告，不在 Python 端修補。

## 5. 每輪報告（必須；Andy 不熟悉程式，用繁體中文、非工程語言）

- 寫入 `docs/reports/YYYYMMDD_v0.6-Pn.md`，同一份內容貼成 PR comment。開頭列分支、最新 SHA、報告路徑。
- 內容：表 1（工作單每一項→完成與否、位置、數值摘要）、**預期變動 vs 實際變動** 兩張表（表 2＝範圍外變動，應為 0）、待 Project 判斷清單、TK 快照比對結果、重算秒數、parity 項數、發現的 Excel 問題。
- 每個工作包一個 PR、一份報告；前一包經 chat 端審查後才開下一包。CI 只是合併門檻，不是等待點。
- **（E4）分支與 PR**：本 repo 的預設分支是 `main`（工作單提到的 master 一律指 Tokenomics）。每個工作包開工即開 PR（可先標 draft），報告全文以 PR comment 張貼，同一份存 `docs/reports/`；chat 端從 Gmail 的 PR 通知讀取報告，只推分支不開 PR 時 chat 端讀不到。

## 6. 禁止事項

- 不得在 Python 硬編碼任何價格、規格、效率或比例參數。
- 不得刪除或改寫 Excel 的來源、標記、查核狀態欄；不得自行更新任何外部數據。
- 不得從未合併的 Tokenomics 分支取值。
- **（E5）環境安全**：任何情況不得嘗試 `dangerouslyDisableSandbox` 或其他繞過系統權限的做法。環境錯誤先診斷（缺套件等；先查本檔與 `crwv-model` 的 CLAUDE.md），無法解決即在報告回報。

## 7. 慣例

- **電力口徑**：`GW` 以 IT 關鍵電力為基準；外部揭露的 GW 數字逐一標註口徑（IT、設施或未明）。
- **來源標記**：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 一律以區間呈現（v0.5 無區間者於備註說明，不自行補值）。
- **曆年制**：FY2025–FY2030，2025 為實際校準年。
- **版本命名**：`YYYYMMDD_OpenAI_v0.6.xlsx`；各工作包完成時的檔案記為 `v0.6-P1`…`v0.6-P5`，P5 完成即 v0.6。
- **Excel 語言**：工作表名稱與具名範圍用英文；標籤與註解用繁體中文。
