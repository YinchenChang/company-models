# MiniMax 收支模型 — Claude Code 工作規範（v0.2；沿用 `openai/CLAUDE.md` 的守則，改寫為本模型）

優先順序：該張工作單 > `docs/workorders/*_共同規則.md` > 規格（`docs/plan/`）> repo 根目錄 README > 本檔。所有路徑相對於 `minimax/`；指令在 `minimax/` 內執行。預設分支 `main`（工作單提到的 master 一律指 Tokenomics）。

## 1. 角色與事實來源
- **Excel 是唯一計算引擎**：`dist/` 與 `out/` 的 xlsx 由 `build_xlsx.py` 產生結構與公式（藍字輸入、黑字公式、綠字跨頁引用），數值由 LibreOffice 重算。Python 不寫入任何計算結果，不在 Python 硬編碼價格、規格、效率、比例；`build_xlsx.py` 的 `SRC`、`INPUTS` 是結構化的輸入表（含標記與區間），Tokenomics 數值一律來自快照。
- **判斷類**：輸入值、機制或公式邏輯、新參數、來源與標記、區間、分類。不停下來等人：依序套用 ①該張工作單的預設 ②規格第 3 節 D1–D8 ③最保守且最常見的口徑；每項在報告「已套用的預設」列出（問題｜預設｜替代｜影響方向）。分不清屬哪一類時視為判斷類並列入報告。
- **工程類**：具名範圍、標籤與單位文字、格式、僅供顯示的列、不改數值的重構、測試與 CI。可自行決定，報告說明。
- **資料分層**：公司財務原始數據→`SRC_MM`；算力物理與產業資料→只取 Tokenomics 官方快照（`TK_Link`），本模型不重算；假設→`Inputs`（藍字，附標記與區間）；Tokenomics 沒有的算力資料→報告「Tokenomics 缺口」。分層優先於標記。
- **不以公司自己的目標或預測反推參數**；公司數字只作對照列（η 校準與 2025 營收校準除外）。

## 2. Tokenomics 取數
- 只經官方快照：`scripts/tk_snapshot.py --tk-dir <Tokenomics checkout> --names data/tokenomics_names.txt --out data/tokenomics_snapshot_<版本>.json`，讀 `model/CURRENT` 的具名範圍；快照含版本、提交、檔名、SHA-256、讀取日。`build_xlsx.py` 讀 `data/tokenomics_snapshot_current.json` 指向的快照（或 `--snapshot` 指定）。
- 唯讀：不改 Tokenomics、不從未合併分支取值、不使用 Excel 外部連結；公司特有變數不回寫 Tokenomics。
- 更新快照＝本模型的修補版：重新產生快照 → 建置 → 重算 → 記 CHANGELOG。`scripts/check_tk_snapshot.py` 比對快照與 Tokenomics 現行 `CURRENT`（DIFF／MISSING 只報 WARN）。
- Tokenomics 新版若改了名稱或口徑，以 Tokenomics 為準並在報告標明。

## 3. 建置與驗收
- `scripts/build_all.sh <日期> <版本> [dist]`：建置 → 重算 → 敏感度 → 再重算 → HTML；預設只寫 `out/`，第三參數 `dist` 才複製到 `dist/`。依賴 `libreoffice-calc-nogui`、openpyxl（只有 `soffice --version` 能執行不代表 Calc 已安裝）。
- 驗收：重算錯誤值 0；Checks 合計 0；M1 起 CI（`.github/workflows/minimax-build.yml`）通過。CI 只是事後對照，驗收以本機為準。
- 發現 Excel 本身的錯誤（`#REF!`、`#DIV/0!`、循環參照、單位不一致）：列在報告，不在 Python 端修補。

## 4. 每輪報告（繁體中文、非工程師看得懂、結論先行）
- `docs/reports/YYYYMMDD_minimax_<段>.md`；建置類另附 `_對照.xlsx`（①變動列 ②關鍵輸出前後值 ③已套用的預設）；進度檔 `docs/reports/20261022_minimax_進度.md` 每步更新。
- 內容：結論、步驟與 commit、已套用的預設、資料缺口、Tokenomics 缺口、給下一段的事、TK 快照比對結果。
- 每個步驟：更新進度檔 → commit → push。PR 完成留言第一行 `[MiniMax 回報] <段>｜完成｜YYYY-MM-DD`；轉 ready 與合併只在工作單或使用者指示時做。

## 5. 禁止事項
- 不得合併任何 PR、force push、改 `main`、刪除分支或歷史；只改 `minimax/` 與 `.github/workflows/minimax-*.yml`（及 repo 根 README 的 MiniMax 一列）。
- 不得杜撰數字；「找不到」與「不存在」分開寫，寫明試過哪些來源。
- **不得在任何對外請求（含 HTTP User-Agent）中使用使用者 email**；需要 User-Agent 時用通用字串（例如 `research-bot`）。
- **環境安全**：任何情況不得嘗試 `dangerouslyDisableSandbox` 或其他繞過系統權限的做法。環境錯誤先診斷，無法解決即在報告回報。
- SemiAnalysis（含 InferenceX）不可單獨引用；同一來源的多個轉載不算獨立交叉驗證。

## 6. 慣例
- 電力口徑：MW 以 IT 關鍵電力為基準。來源標記：Verified／Interested-party／Analogy／Assumed／Derived／Decision；Analogy、Assumed 一律附區間。
- 曆年制 FY2025–FY2030（2025 為實際校準年）；金額 US$M（人民幣揭露以期間匯率換算並登錄來源）。
- 版本命名：`YYYYMMDD_MiniMax收支模型_vX_Y.xlsx／.html`；Excel 語言：工作表名稱英文，標籤與註解繁體中文。
