# Anthropic 收支模型 — 建置代理工作規範（v0.1；以 `openai/CLAUDE.md`（v0.6）改寫）

> 本檔所有路徑相對於 `anthropic/`；指令一律在 `anthropic/` 內執行。預設分支 `main`；本版建置分支 `claude/anthropic-v0.1`（單一分支、單一 PR，分段 commit）。
> 流程＝自主建置：chat 端寫工作單（`docs/workorders/`）→ 建置代理執行 → chat 端審查並合併；Andy 只在里程碑（v0.1 完成）審查。開工前必讀 `docs/workorders/20261008_anthropic_共同規則.md` 與規格 `docs/plan/20261008_Anthropic_v0.1_規格.md`。
> 判斷類不停下來等：依序套用該張工作單的預設 → 規格第 3 節 D1–D20 → 共同規則第 5 節（最保守且最常見的口徑），並在報告「已套用的預設」逐項列出。停止條件只有共同規則第 6 節所列。

## 1. 角色與事實來源

- **Excel 活頁簿是唯一事實來源**：`model/YYYYMMDD_Anthropic_v0.1.xlsx`（現行檔名見 `model/CURRENT`）。所有計算機制與數值以 Excel 為準。
- Python（builder、engine、tools）只能**產生結構**或**執行／呈現** Excel；不得在 Python 另寫旁路計算，不得在 Python 硬編碼價格、規格、比例等參數——參數一律來自 `data/*.yaml`（經 builder 寫入 SRC_ANT／Inputs）或 Tokenomics 快照。
- **資料分層**（Tokenomics 治理 G1、G8、G9；分層優先於標記）：
  - 公司財務原始數據（營收、方案價格、用戶、合約、現金、融資輪、人事、管理層目標）→ `SRC_ANT`（來源 `data/anthropic_src.yaml`）。
  - AI 技術與算力的原始數據與推導 → 不在本模型重算，一律取自 Tokenomics（`TK_Link` 快照）。Tokenomics 沒有、但屬算力相關者 → 列「Tokenomics 缺口」，不在本模型自建。
  - 假設 → `Inputs`（藍字，附標記與區間；來源 `data/anthropic_inputs.yaml`）。
  - 跨模型對照（OpenAI v0.6 的命題輸出）→ `OAI_Link` 快照，只作並排對照，**不得被任何計算頁引用**（Checks 除外）。
- **Excel 優先**：藍字輸入（SRC_ANT、Inputs 的值／低／高）由 Excel 擁有；builder 重建時只在「Excel 值仍等於上次 builder 寫入的預設值」才改寫（隱藏頁 `_Defaults`；`builder/preserve.py`）。TK_Link、OAI_Link 例外：由 builder 從快照重寫。
- **（E6）公式不含常數**：算式中的恆等式除外（`1-比例`、年數 `+1`、單位換算 1000／1e6 須寫成 Inputs 的定義常數）；測試掃描 Demand、Revenue、Compute、Cost、Funding、Reverse。
- **穩定 ID**：SRC_ANT_nnn、INP_nnn、Checks Cnn 一經發出即固定（`builder/id_registry.json`）；新列取下一個號碼，被移除的號碼退役、不重用。
- 本模型公式只引用 `TK_`／`SRC_ANT_`／`INP_` 與本模型具名範圍；不使用 Excel 外部連結。

## 2. 目錄

```
model/            現行 xlsx（唯一一份）；model/archive/ 放舊版；model/CURRENT 一行記錄現行檔名
data/             anthropic_src.yaml（SRC_ANT 來源）、anthropic_inputs.yaml（Inputs 來源）、A1 研究筆記
builder/          產生 Excel 的程式（build.py 組裝、各頁模組、tk_link.py、oai_link.py、preserve.py、source_rules.py、id_registry.json）
engine/           以 pycel 載入 xlsx 並重算（沿用 openai/engine）
tests/            parity（Excel 與 engine 一致）、builder 結構、HTML 一致性
tools/            check_tk_snapshot.py、export_csv.py、build_engine_cache.py、build_html.py 等（沿用 openai/tools 改寫）
docs/plan/        規格與命題對照表；docs/workorders/ 工作單；docs/reports/ 每段報告與進度檔；docs/handoff/ 交接檔
dist/             成品（HTML 一頁摘要＋xlsx）
```

## 3. 一致性測試（必須通過才可合併）

1. LibreOffice headless 重算 xlsx 取得期望值（需 `libreoffice-calc`；只有 `soffice --version` 能執行不代表 Calc 已安裝）。
2. engine 計算同一組情境（`tests/parity/scenarios.yaml`），全部公式格相對誤差 ≤1e-9、字串完全一致、任何情境不得出現錯誤值；每個情境至少改變一格輸出（防空轉）。
3. 結構性：工作表、具名範圍、`model/CURRENT`、SRC_ANT／Inputs 筆數與報告一致；E6 掃描；OAI_Link 不被計算頁引用。
4. CI：`.github/workflows/anthropic-parity.yml`（只在 `anthropic/**` 變動時執行；合併門檻 job 名稱 `parity`）；`tk-snapshot` 為 WARN。
5. 增量重算與全簿重算各 < 2 秒（硬性）。遇 pycel 不支援的函數，改用等價函數寫在 Excel（工程類），不得在 Python 旁路。

## 4. 建置與同步（每段）

1. `python3 builder/build.py --tk-dir <Tokenomics checkout> --oai-xlsx <OpenAI v0.6 xlsx> --out model/<新檔名>.xlsx [--base model/<現行>.xlsx]`；舊版以 `git mv` 移入 `model/archive/`；更新 `model/CURRENT`。
2. 執行 parity 與 `tools/check_tk_snapshot.py`；有意識地更新 `tests/parity/scenarios.yaml`。
3. `CHANGELOG.md` 記錄 Excel 版本、段別（`v0.1-A2`…）、commit 與變動摘要。
4. 發現 Excel 本身的錯誤（`#REF!`、`#DIV/0!`、循環參照、單位不一致）：修正公式（工程類）並在報告列出；不在 Python 端修補數值。

## 5. 每段報告（必須；Andy 不熟悉程式，用繁體中文、非工程語言、結論先行）

- `docs/reports/YYYYMMDD_v0.1-An.md`，同一份內容貼成 PR comment（第一行 `[Anthropic 回報] An｜完成｜YYYY-MM-DD`）。開頭列分支、最新 SHA、報告路徑。
- 內容：表 1（工作單每一項→完成與否、位置、數值摘要）、預期變動 vs 實際變動、範圍外變動（應為 0）、**已套用的預設**（問題｜預設｜替代｜對結果影響方向，能量化就量化）、資料缺口（「找不到」與「不存在」分開）、Tokenomics 缺口、TK 快照比對、重算秒數、parity 項數、建議後續。
- **對照 Excel** `docs/reports/YYYYMMDD_v0.1-An_對照.xlsx`：①本段新增或變動的每一列（頁、列名、ID、值、來源／標記）；②本段關鍵輸出 FY2025–2030；③已套用的預設。

## 6. 禁止事項

- 不得在 Python 硬編碼任何價格、規格、效率或比例參數；不得杜撰數字。
- 不得以公司自己的目標或預測反推參數（反向模式是獨立對照區塊，不回饋基準）。
- 不得刪除或改寫 Excel 的來源、標記、查核狀態欄；SemiAnalysis（含 InferenceX）不可單獨引用。
- 不得從未合併的 Tokenomics 分支取值；不得改動 Tokenomics、`openai/`、`nebius/`、`oracle/` 或其他資料夾（讀取可以）。
- 不得合併 PR、force push、改 `main`、刪除分支或歷史。
- **環境安全**：任何情況不得嘗試 `dangerouslyDisableSandbox` 或其他繞過系統權限的做法。

## 7. 慣例

- **電力口徑**：`GW` 以 IT 關鍵電力為基準；外部揭露的 GW 數字逐一標註口徑（IT、設施或未明），設施口徑以 Inputs 的 PUE 換算。
- **來源標記**：Verified／Interested-party／Analogy／Assumed／Derived／Decision；Analogy、Assumed 一律給區間。
- **曆年制**：FY2025–FY2030，2025 為實際校準年。金額單位 $B（十億美元），token 單位兆（T），電力 GW。
- **版本命名**：`YYYYMMDD_Anthropic_v0.1.xlsx`；各段完成時記為 `v0.1-A2`…`v0.1-A4`，A4 完成即 v0.1；成品 `dist/YYYYMMDD_Anthropic收支模型_v0_1.html／.xlsx`。
- **Excel 語言**：工作表名稱與具名範圍用英文；標籤與註解用繁體中文。
- **與 OpenAI v0.6 同名**：命題輸出與主要衍生量的具名範圍沿用 OpenAI 的前綴與名稱（`DEM_`、`REV_`、`CMP_`、`COST_`、`FND_`、`RVS_`；例如 `COST_PropGap_VR`、`COST_Coverage`、`FND_ExtNeedCum`），方便兩家並排。
