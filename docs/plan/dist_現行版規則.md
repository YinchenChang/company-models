# dist/ 現行版規則

- dist/ 只留現行版；舊版在 git 歷史（`git log -- <路徑>`），Drive 只留現行閱讀副本（Andy 2026-11-17）

（本行待 `claude/repo-governance` 分支的根 README 規則清單改完後併入；為避免衝突先放這裡。）

## 適用範圍與判定

- 範圍：repo 內所有 `*/dist/`（含 `dist/_待刪除/`），不含 `software_ai/**`、`physical_ai/**`（其他專案自管）。
- 產出族：檔名去掉開頭日期（`YYYYMMDD_`）與版號（`v0.3.10`、`v0_6_5` 等）後相同者為同一族（副檔名不同視為不同族）。同族只留版號最大者（版號同時比日期）；無版號的檔（`index.html`、`data.json` 等）一律保留。
- 舊版若在 dist/ 以外的程式、資料或設定被引用（`*.py`、`*.sh`、`*.json`、`*.csv`、`*.yml`、`*.yaml`、`*.txt`、`CURRENT`、`markets/invest_map/data/*`、`markets/invest_map/profiles/*`），先保留；只在 `docs/reports/`、`docs/workorders/`、README、CHANGELOG、交接檔等敘述中出現者不擋。
- 取回舊版：`git log --oneline -- <公司>/dist/_待刪除/<檔名>` 找到刪除前的 commit，再 `git show <commit>^:<路徑> > <檔名>`。`markets/invest_map/audit_profile_cites.py --old-dir <資料夾>` 可吃取回的舊檔。
- CI 不需要舊版留在樹上：`ai-semis-build` 只比對 dist/ 頂層唯一的 xlsx；`.github/scripts/xlsx_pr_diff.py`、`tools/consistency/check.py`、`tools/segment_check` 本來就略過 `_待刪除`；`tools/source_audit` 與 `api_pr_review.py` 只讀 `dist/*.xlsx`／`dist/*.html` 頂層。

## 2026-11-17 首次清理後仍留在 `_待刪除/` 的檔

### 因被程式／資料引用而保留（65 個）

其中真正會開檔的只有 `alphabet/q3cf_v03.py`、`amazon/q3cf_v03.py`（讀 v0.3 xlsx 產生 Q3 現金流拆解）與 `tools/model_governance/version_table.py`（手動給路徑）；其餘是快照、seed、公司檔案中的出處字串。改寫這些引用（或改用 git 歷史路徑）後可再移除。

| 檔案 | 引用來源 |
|---|---|
| `advantest/dist/_待刪除/20261031_Advantest個股模型_v0.6.xlsx` | siliconomics 快照／seed（出處字串） |
| `advantest/dist/_待刪除/20261106_Advantest個股模型_v0.7.4.xlsx` | siliconomics 快照／seed（出處字串） |
| `alab/dist/_待刪除/20261104_Astera Labs個股模型_v0.2.xlsx` | invest_map 公司檔案／資料 |
| `alchip/dist/_待刪除/20261115_世芯-KY個股模型_v0.2.5.xlsx` | siliconomics 快照／seed（出處字串） |
| `alphabet/dist/_待刪除/20261027_Alphabet個股模型_v0.3.xlsx` | alphabet/build_alphabet.py；alphabet/data/q3cf_v03_2026q3.json；alphabet/q3cf_v03.py |
| `alphabet/dist/_待刪除/20261030_Alphabet個股模型_v0.3.3.xlsx` | alphabet/data/wattonomics_c0_v0.3.2.json；wattonomics 快照／seed（出處字串） |
| `alphabet/dist/_待刪除/20261104_Alphabet個股模型_v0.3.6.xlsx` | invest_map 公司檔案／資料 |
| `alphabet/dist/_待刪除/20261114_Alphabet個股模型_v0.3.9.xlsx` | wattonomics 快照／seed（出處字串） |
| `amazon/dist/_待刪除/20261027_Amazon個股模型_v0.3.xlsx` | amazon/build_amazon.py；amazon/q3cf_v03.py |
| `amazon/dist/_待刪除/20261031_Amazon個股模型_v0.3.1.xlsx` | amazon/data/wattonomics_c0_v0.3.2.json；wattonomics 快照／seed（出處字串） |
| `amazon/dist/_待刪除/20261104_Amazon個股模型_v0.3.2.xlsx` | amazon/build_amazon.py |
| `amazon/dist/_待刪除/20261104_Amazon個股模型_v0.3.4.xlsx` | invest_map 公司檔案／資料 |
| `amazon/dist/_待刪除/20261106_Amazon個股模型_v0.3.6.xlsx` | amazon/build_amazon.py |
| `amazon/dist/_待刪除/20261115_Amazon個股模型_v0.3.7.xlsx` | wattonomics 快照／seed（出處字串） |
| `amd/dist/_待刪除/20261115_AMD模型_v0.4.7.xlsx` | siliconomics 快照／seed（出處字串） |
| `anthropic/dist/_待刪除/20261114_Anthropic收支模型_v0_2_2.xlsx` | wattonomics 快照／seed（出處字串） |
| `arista/dist/_待刪除/20261104_Arista個股模型_v0.4.xlsx` | invest_map 公司檔案／資料 |
| `asml/dist/_待刪除/20261105_ASML個股模型_v0.2.4.xlsx` | asml/build_asml.py |
| `avgo/dist/_待刪除/20261114_Broadcom個股模型_v0.5.7.xlsx` | siliconomics 快照／seed（出處字串） |
| `coreweave/dist/_待刪除/20261021_CoreWeave收支模型_v4_12.xlsx` | wattonomics 快照／seed（出處字串） |
| `coreweave/dist/_待刪除/20261114_CoreWeave個股模型_v5.0.3.xlsx` | wattonomics 快照／seed（出處字串） |
| `crdo/dist/_待刪除/20261030_Credo個股模型_v0.2.4.xlsx` | invest_map 公司檔案／資料 |
| `delta/dist/_待刪除/20261102_台達電個股模型_v0.4.xlsx` | invest_map 公司檔案／資料 |
| `disco/dist/_待刪除/20261031_Disco模型_v0.2.3.xlsx` | invest_map 公司檔案／資料 |
| `guc/dist/_待刪除/20261115_創意個股模型_v0.3.5.xlsx` | siliconomics 快照／seed（出處字串） |
| `ibiden/dist/_待刪除/20261102_Ibiden個股模型_v0.4.xlsx` | invest_map 公司檔案／資料 |
| `lite/dist/_待刪除/20261104_Lumentum個股模型_v0.4.xlsx` | invest_map 公司檔案／資料 |
| `liteon/dist/_待刪除/20261031_光寶個股模型_v0.3.xlsx` | invest_map 公司檔案／資料 |
| `mediatek/dist/_待刪除/20261115_聯發科個股模型_v0.3.5.xlsx` | siliconomics 快照／seed（出處字串） |
| `meta/dist/_待刪除/20261030_Meta個股模型_v0.1.1.xlsx` | meta/data/wattonomics_c0_v0.3.2.json；wattonomics 快照／seed（出處字串） |
| `meta/dist/_待刪除/20261115_Meta個股模型_v0.2.1.xlsx` | software_ai/palantir/data/facts_palantir.csv；software_ai/palantir/data/facts_palantir.py；wattonomics 快照／seed（出處字串） |
| `microsoft/dist/_待刪除/20261027_Microsoft個股模型_v0.3.xlsx` | microsoft/attrib_v031.py；microsoft/build_microsoft.py |
| `microsoft/dist/_待刪除/20261030_Microsoft個股模型_v0.3.3.xlsx` | microsoft/data/wattonomics_c0_v0.3.2.json；wattonomics 快照／seed（出處字串） |
| `microsoft/dist/_待刪除/20261104_Microsoft個股模型_v0.3.5.xlsx` | invest_map 公司檔案／資料 |
| `microsoft/dist/_待刪除/20261115_Microsoft個股模型_v0.3.8.xlsx` | wattonomics 快照／seed（出處字串） |
| `minimax/dist/_待刪除/20261022_MiniMax收支模型_v0_2.xlsx` | invest_map 公司檔案／資料 |
| `minimax/dist/_待刪除/20261115_MiniMax收支模型_v0_2_3.xlsx` | invest_map 公司檔案／資料 |
| `mrvl/dist/_待刪除/20261115_邁威爾個股模型_v0.3.5.xlsx` | siliconomics 快照／seed（出處字串） |
| `nanya_pcb/dist/_待刪除/20261030_南電個股模型_v0.2.1.xlsx` | invest_map 公司檔案／資料 |
| `nebius/dist/_待刪除/20261030_Nebius個股模型_v0_5_2.xlsx` | wattonomics 快照／seed（出處字串） |
| `nebius/dist/_待刪除/20261106_Nebius個股模型_v0_5_3.xlsx` | nebius/build_nebius.py |
| `nebius/dist/_待刪除/20261115_Nebius個股模型_v0.6.0.xlsx` | invest_map 公司檔案／資料 |
| `nebius/dist/_待刪除/20261115_Nebius個股模型_v0.6.1.xlsx` | wattonomics 快照／seed（出處字串） |
| `nvidia/dist/_待刪除/20261106_NVIDIA個股模型_v0.3.4.xlsx` | tools/model_governance/version_table.py |
| `nvidia/dist/_待刪除/20261111_NVIDIA個股模型_v0.3.5.xlsx` | tools/model_governance/version_table.py |
| `nvidia/dist/_待刪除/20261114_NVIDIA個股模型_v0.3.10.xlsx` | siliconomics 快照／seed（出處字串） |
| `openai/dist/_待刪除/20261008_OpenAI收支模型_v0_6.xlsx` | amazon/data/amazon_facts_20261008.json |
| `openai/dist/_待刪除/20261029_OpenAI收支模型_v0_6_4.xlsx` | invest_map 公司檔案／資料；wattonomics 快照／seed（出處字串） |
| `openai/dist/_待刪除/20261115_OpenAI收支模型_v0_6_5.xlsx` | wattonomics 快照／seed（出處字串） |
| `oracle/dist/_待刪除/20261030_Oracle個股模型_v0.4.2.xlsx` | oracle/data/wattonomics_c0_v0.3.2.json；wattonomics 快照／seed（出處字串） |
| `oracle/dist/_待刪除/20261104_Oracle個股模型_v0.4.5.xlsx` | invest_map 公司檔案／資料 |
| `oracle/dist/_待刪除/20261115_Oracle個股模型_v0.5.2.xlsx` | invest_map 公司檔案／資料；wattonomics 快照／seed（出處字串） |
| `skhynix/dist/_待刪除/20261030_SK海力士個股模型_v0.6.1.xlsx` | invest_map 公司檔案／資料 |
| `smci/dist/_待刪除/20261106_美超微個股模型_v0.3.5.xlsx` | smci/build_smci.py |
| `spacex/dist/_待刪除/20261030_SpaceX個股模型_v0.2.1.xlsx` | spacex/data/_待刪除/wattonomics_c0_v0.3.2.json；spacex/data/wattonomics_c0_v0.3.3.json；wattonomics 快照／seed（出處字串） |
| `spacex/dist/_待刪除/20261115_SpaceX個股模型_v0.2.9.xlsx` | wattonomics 快照／seed（出處字串） |
| `ter/dist/_待刪除/20261031_泰瑞達個股模型_v0.3.xlsx` | siliconomics 快照／seed（出處字串） |
| `ter/dist/_待刪除/20261106_泰瑞達個股模型_v0.4.4.xlsx` | siliconomics 快照／seed（出處字串） |
| `unimicron/dist/_待刪除/20261102_欣興個股模型_v0.3.xlsx` | invest_map 公司檔案／資料 |
| `whitefiber/dist/_待刪除/20261022_WhiteFiber收支模型_v0_2.xlsx` | wattonomics 快照／seed（出處字串） |
| `whitefiber/dist/_待刪除/20261115_WhiteFiber個股模型_v0.4.0.xlsx` | invest_map 公司檔案／資料；whitefiber/build_whitefiber.py；whitefiber/scripts/decomp_v50.py |
| `whitefiber/dist/_待刪除/20261115_WhiteFiber個股模型_v0.4.1.xlsx` | wattonomics 快照／seed（出處字串） |
| `whitefiber/dist/_待刪除/20261115_WhiteFiber個股模型_v0.4.2.xlsx` | wattonomics 快照／seed（出處字串） |
| `zhipu/dist/_待刪除/20261104_智譜個股模型_v0_1_3.xlsx` | invest_map 公司檔案／資料 |
| `zhipu/dist/_待刪除/20261115_智譜個股模型_v0_1_6.xlsx` | invest_map 公司檔案／資料 |

### 改名後的舊產出族（17 個，依「每族留最新」暫留）

這些族已改名（收支模型→個股模型、AI投資地圖→AI算力地圖、智譜個股頁→智譜個股模型），在 `_待刪除/` 中是該舊族唯一一份，依規則暫不移除，待 Andy 決定。

- `alphabet/dist/_待刪除/20261025_Alphabet收支模型_v0_2_2.html`
- `alphabet/dist/_待刪除/20261025_Alphabet收支模型_v0_2_2.xlsx`
- `amazon/dist/_待刪除/20261025_Amazon收支模型_v0_2_2.html`
- `amazon/dist/_待刪除/20261025_Amazon收支模型_v0_2_2.xlsx`
- `coreweave/dist/_待刪除/20261106_CoreWeave收支模型_v4_13.html`
- `coreweave/dist/_待刪除/20261106_CoreWeave收支模型_v4_13.xlsx`
- `markets/invest_map/dist/_待刪除/20261116_AI投資地圖_資料_v0.6.49.xlsx`
- `microsoft/dist/_待刪除/20261025_Microsoft收支模型_v0_2_2.html`
- `microsoft/dist/_待刪除/20261025_Microsoft收支模型_v0_2_2.xlsx`
- `nebius/dist/_待刪除/20261022_Nebius收支模型_v0_5_1.html`
- `nebius/dist/_待刪除/20261022_Nebius收支模型_v0_5_1.xlsx`
- `oracle/dist/_待刪除/20261022_Oracle收支模型_v0_3_1.html`
- `oracle/dist/_待刪除/20261022_Oracle收支模型_v0_3_1.xlsx`
- `whitefiber/dist/_待刪除/20261106_WhiteFiber收支模型_v0_2_1.html`
- `whitefiber/dist/_待刪除/20261106_WhiteFiber收支模型_v0_2_1.xlsx`
- `zhipu/dist/_待刪除/20261022_智譜收支模型_v0_1_1.xlsx`
- `zhipu/dist/_待刪除/20261115_智譜個股頁_v0_1_5.html`
