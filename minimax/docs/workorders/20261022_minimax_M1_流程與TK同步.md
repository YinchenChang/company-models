# 工作單 MiniMax M1｜流程補齊＋Tokenomics 官方快照取數、同步至現行版（2026-10-22，chat 端，r0）

- 分支 `claude/minimax-v0.2`（已含規格與工作單；PR base `main`，標題「MiniMax v0.2」——M1 開 draft PR，M3 沿用同一個 PR）。先讀共同規則與規格。
- 參考實作（唯讀）：`openai/tools/check_tk_snapshot.py`、`openai/builder/tk_link.py`、`openai/tools/tk_attribution.py`、`openai/CLAUDE.md`、`openai/.github` 無（CI 在 repo 根 `.github/workflows/openai-parity.yml`）。

## 步驟
0. 開 draft PR；建立進度檔 `minimax/docs/reports/20261022_minimax_進度.md`（目前狀態／步驟紀錄／已套用的預設）。
1. **流程文件**：`minimax/CLAUDE.md`（沿用 `openai/CLAUDE.md` 的守則，改寫為本模型：Excel 為唯一計算引擎、`build_xlsx.py` 只產生結構與公式、判斷類依工作單預設、Tokenomics 只經官方快照取數、每輪報告格式、禁止事項含 email 與 sandbox 條款）；`minimax/CHANGELOG.md`（補記 v0.1：2026-10-08，PR #25，Tokenomics v5.24 經 OpenAI v0.6 手抄 12 格；PR #36 連接審視）。
2. **官方快照取數（數值先不變）**：新增 `minimax/data/tokenomics_names.txt`（現用 12 名＋規格 D6 P×Q 需要的 `IF_HoldEcon` 全 15 欄；可參考審視報告第三節草稿）與快照工具（照抄改寫 `openai/builder/tk_link.py` 的讀取邏輯，或寫 `minimax/scripts/tk_snapshot.py`）：讀 Tokenomics checkout 的 `model/CURRENT` 具名範圍 → `minimax/data/tokenomics_snapshot_<版本>.json`（含版本、提交、檔名、SHA-256、讀取日）。`build_xlsx.py` 的 `TK` 清單改為讀快照，不再寫死數字；`TK_Link` 頁加檔名、提交、SHA-256、讀取日。**先用 v5.24（Tokenomics 提交 `bdb0de7`）產生快照並建置，證明與 v0.1 逐格相同**（工程類改動零數值影響）。
3. **同步至現行版**：取 Tokenomics master HEAD 的 `model/CURRENT`（2026-10-22 為 v19.0，`1916a06`／`ad33e28` 同檔，SHA-256 `d1728705…`；開工時若有更新版合併，用當時 HEAD）。重新產生快照、建置、重算（`scripts/build_all.sh <日期> v0_2_m1` 產物放 `out/`，不覆蓋 `dist/`）。
4. **差異拆解**：列 v0.1 → M1 主要輸出前後值（各年 API MW、η 2025 與 1H26、總 MW、每 MW 營收、每 MW 全成本、覆蓋率、累計需新融資、Valuation 反向檢驗）；歸因：一次只換一組名稱（① `IF_TokGW_*` ② `IF_Util` ③ 持有成本與全成本類），以 LibreOffice 重算完成。審視報告指出 η 校準會吸收產能絕對水準、只剩世代比例起作用——報告驗證這一點。
5. **CI**：`.github/workflows/minimax-build.yml`（只在 `minimax/**` 或本檔變動時執行）：安裝 LibreOffice Calc 與 openpyxl → `build_xlsx.py` → `scripts/recalc.py` → 檢查錯誤值 0 與 Checks 合計 0 → 快照與 `tokenomics_names.txt` 一致（可選：比對 Tokenomics master 的 WARN，不擋合併）。CI 只是事後對照，驗收以本機為準。
6. **報告** `minimax/docs/reports/20261022_minimax_M1.md`＋`_對照.xlsx`；README 更新（頁表與 TK 來源）；PR 留言 `[MiniMax 回報] M1｜完成｜YYYY-MM-DD`；**不轉 ready**（M3 完成後才轉）。

## 判斷類預設
- 產能層照規格 D1（維持 `IF_TokGW × IF_Util`）。Tokenomics 新版若改了名稱或口徑（例如 v9.0 起快取命中進產能），以 Tokenomics 為準並在報告標明。
- v5.29 審視指出 `L1_HoldEconMW_*` 單位疑義：本模型不引用 `L1_HoldEconMW_*`，改用 `IF_HoldEcon`（$B/GW/年；本模型換算 $M/MW/年時數值相同，在 TK_Link 頁寫明換算）。
