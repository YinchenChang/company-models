# 工作單 CoreWeave W8｜Tokenomics 快照換 v5.35 → v4.10（2026-10-15，r0，chat 端）

- 來源：CoreWeave 對話串交接檔（2026-10-15）待辦 B1；Andy 2026-10-15「都OK」。
- 類別：工程類（數值 0 差異）。預估 1 小時以內，不拆段（Tokenomics CLAUDE.md 第 4b 節第 3 項）。代理模型：Sonnet。
- 前置：v4.9 已合併（main `63d7e2b`）。必讀：`coreweave/docs/workorders/20261007_coreweave_共同規則.md`、`coreweave/CLAUDE.md`、本單、W6 工作單（`20261009_coreweave_W6_Tokenomics_v5.31.md`，換快照的前例）。
- 分支：`claude/coreweave-w8-v4.10`（chat 端已建立並放入本單）；PR 標題「CoreWeave v4.10：Tokenomics 快照換 v5.35」。

## 事實（chat 端已查證，2026-10-15）
- Tokenomics master `9dff7f5`，`model/CURRENT`＝`20261014_Tokenomics_v5.35.xlsx`。
- 以 `tools/tokenomics/import_tokenomics.py` 依 `coreweave/data/tokenomics_names.txt` 產生 v5.35 快照，與 v5.31 快照逐值比較：29 個名稱、366 個數值，差異 0。

## 改動
1. 產生 `coreweave/data/tokenomics_snapshot_v5.35.json`（同上工具與名稱清單）；`git rm` 舊的 v5.31 快照（目前資料夾只保留現行快照）。
2. `company.json` Tokenomics 區：`snapshotFile`、`version`（v5.35）、`commit`（Tokenomics 完整 SHA）；`_note` 末尾追加一句「W8：快照換 v5.35（29 名數值與 v5.31 相同）」。
3. k 隨錨重算：因錨不變，k 數值不變；證據表與 note 中「v5.31」作為現行錨的字樣，只在指「現行快照」時改為 v5.35；描述歷史（例如「v5.31 起提供」「v5.27 0.762、v5.31 0.891」）一律保留原字樣。
4. 交接檔加一句（下游資料契約第 3 節要求寫明）：「收入採商業簽約率口徑（MW-year），不使用技術利用率 `IF_Util`。」
5. 升版 v4.10：`vlog.py` 與 `tail.js` 的 VLOG 各加一列（目標價欄 $32.0，變動欄「三情境目標價不變；拆解 (a)(b)(c)(d) 皆 0」）；`dist/` 換 v4.10 成品並移除 v4.9；交接檔以 v4.10 取代 v4.9；README 有版本處同步。

## 驗收
- verify 三組全過：(1) `scripts/verify.sh`；(2) 升版對 v4.9 成品 `--vs-dist`，EXPECT 只含版本文字、快照版本與 commit、VLOG 新列；**所有數值 0 差異**；(3) legacy 回歸對 v4.6 維持 0 差異（同 W6／W7 做法）。
- `import_tokenomics.py --check coreweave/data/tokenomics_snapshot_v5.35.json` 通過。
- 三情境目標價、融資缺口、敏感度全部不變（基準 $31.99、保守 $42.70、積極 $14.72）。

## 回報
- 短報告 `coreweave/docs/reports/20261015_coreweave_v4.10_Tokenomics_v5.35.md`（結論先行、verify 摘要、前後對照「無變動」並列現值）；PR 留言第一行 `[CRWV 回報] W8｜完成｜2026-10-15`。
- 「需要 Andy 做的事」：合併前用真正的 Excel 開啟 v4.10。
- 不做：B2（q × p 收入）、B3（收入上限層）、B4（用途標註），等 Tokenomics 回覆。
