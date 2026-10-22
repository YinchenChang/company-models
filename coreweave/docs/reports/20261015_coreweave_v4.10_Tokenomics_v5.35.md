# CoreWeave v4.10：Tokenomics 快照換 v5.35（W8；v4.9 → v4.10）

2026-10-15｜工作單 W8（`docs/workorders/20261015_coreweave_W8_Tokenomics_v5.35快照.md`）｜PR #54｜成品：`dist/20261015_CoreWeave收支模型_v4_10.{html,xlsx}`｜工程類，數值 0 差異

## 目前狀態
- 完成：快照換 v5.35、company.json 更新、升版 v4.10、verify 三組全過、`import_tokenomics.py --check` 通過。
- 下一步：chat 端審查 PR #54 後合併。B2（q × p 收入）、B3（收入上限層）、B4（用途標註）不在本單，等 Tokenomics 回覆。
- 需要 Andy 做的事：合併前用真正的 Excel 開啟 `dist/20261015_CoreWeave收支模型_v4_10.xlsx`（看「Tokenomics_取數」第 2 列與 J 欄顯示 v5.35 · 9dff7f5、「檢查_版本紀錄」有 v4.10、開檔重算）。

## 結論
1. Tokenomics 快照由 v5.31 換成 v5.35（master `9dff7f5`，`model/CURRENT`＝`20261014_Tokenomics_v5.35.xlsx`）；29 個名稱、366 個數值與 v5.31 逐值相同，所以 k、收入、成本、資本支出、折舊、融資缺口、敏感度全部不變。
2. 三情境目標價無變動：保守 $42.70、基準 $31.99、積極 $14.72（皆賣出）。變動拆解 (a)(b)(c)(d) 全為 0。
3. 與 v4.9 成品的 Excel 差異只有：「Tokenomics_取數」版本與 commit 字樣（J 欄 121 格＋第 2 列來源說明）、「檢查_版本紀錄」新增 v4.10 列；數值差異 0。

## 1. 前後對照（無變動）
| 情境 | v4.9 | v4.10 | 融資前缺口（US$bn） |
|---|---|---|---|
| 保守 4.2 GW | $42.70 | **$42.70** | 32.1 |
| 基準 5.6 GW | $31.99 | **$31.99** | 79.3 |
| 積極 8 GW | $14.72 | **$14.72** | 151.8 |

敏感度、錨定 × 倍數矩陣、反向 DCF、每 MW 收入與成本：無變動。

## 2. verify 三組
| 組 | 指令 | 結果 |
|---|---|---|
| (1) 升版 | `EXPECT=scripts/expect/v4_10_vs_v4_9.txt scripts/verify.sh --vs-dist`（對 v4.9 成品） | 全部通過 25 項；xl_diff 值與公式 0 差異（預期差異 122 格＋新增版本紀錄列）；crawl 只有「版本紀錄」1 頁有差異（VLOG 新列）；cmp31 三情境各 555 項、FY27 錨定 501 項 OK |
| (2) 新方法 | `scripts/verify.sh --vs-dist`（對 v4.10 成品） | 全部通過 25 項；crawl、xl_diff 值與公式皆 0 差異 |
| (3) 舊方法回歸 | `scripts/verify_legacy.sh`（對 v4.6） | 全部通過 25 項，0 差異 |
| 取數重現 | `import_tokenomics.py --check data/tokenomics_snapshot_v5.35.json` | 通過：29 個名稱（missing 0），相對誤差 ≤ 1e-9 |

## 3. 做了什麼
- `data/tokenomics_snapshot_v5.35.json` 由 `tools/tokenomics/import_tokenomics.py` 依 `data/tokenomics_names.txt` 產生；舊 v5.31 快照 `git rm`。
- `company.json` → `tokenomics`：snapshotFile、version（v5.35）、commit（`9dff7f548c8557e46179dd9c9aa945c28c92e959`）；`_note` 末尾加「W8：快照換 v5.35（29 名數值與 v5.31 相同）」。
- 描述歷史的 v5.31 字樣（k 證據表「v5.27 0.762、v5.31 0.891」、「v5.31 起提供」「Tokenomics v5.31 已把保固期…」、README W6 小節）一律保留；k 數值不變。
- 交接檔以 v4.10 取代 v4.9（新增 2o 節，並加「收入採商業簽約率口徑（MW-year），不使用技術利用率 `IF_Util`」）；README、VLOG（`vlog.py`、`tail.js` 各加 v4.10 列，目標價 $32.0）同步；`dist/` 換 v4.10、移除 v4.9。
- 升版清單：`scripts/expect/v4_10_rules.json` → `v4_10_vs_v4_9.txt`（逐格列出，非整張工作表放行）。

## 4. 已套用的預設（問題｜預設｜替代｜影響）
| 問題 | 預設 | 替代 | 影響 |
|---|---|---|---|
| 版本正規化只認一位數次版號（`v\d\.\d`），v4.10 會被誤讀成「vX0」使全站出現假差異 | 把 `xl_diff.py`、`crawl.py`、`scripts/make_expect.py` 的正規化改為 `v\d+\.\d+` | 另寫 EXPECT 逐格列出 | 只影響比對工具；舊版比對結果不變 |
| 「Tokenomics v5.31 壽命期等值費率」（`build_xlsx.py` 說明文字）是否改 v5.35 | 保留（描述 v5.31 起的定義，且改了會多一個 Excel 文字差異） | 改為 v5.35 | 無數字影響 |
| `import --check` 預設找不到 Tokenomics clone 會略過 | 驗收時設 `TOKENOMICS_DIR=/home/claude/tokenomics` 實際檢查 | — | 無 |

## 5. 可移植性
無新增公司特有內容。換公司沿用 `tokenomics` 區段與 `import_tokenomics.py` 的做法。
