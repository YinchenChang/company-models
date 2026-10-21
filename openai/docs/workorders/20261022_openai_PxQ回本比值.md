# 工作單 OpenAI v0.6.1｜P×Q 與「實際 m×L ÷ 所需 m×L」比值序列（2026-10-22，chat 端，r0）

- 依據：`openai/docs/plan/20261022_OpenAI_PxQ回本比值_交接.md`（Tokenomics 串交接，第 2 節 Andy 原話為權威規格）；Andy 2026-10-22 07:25 指示。
- 分支 `claude/openai-pxq`（base `main`）；一個 PR；單段（預估 ≤2 小時，超過 4 小時停下交 chat 端）。
- 共同規則：`openai/docs/workorders/20261007_openai_共同規則.md`（資料紀律、穩定 ID、E6 公式不含常數、Excel 優先、進度檔每步 commit＋push）；另依 Project 索引「分段與派代理」條：超過 2 分鐘的命令背景執行並**等它跑完再回報**；不得在任何對外請求（含 HTTP User-Agent）中使用使用者 email；不得合併、不得 force push、不得改 main。
- 版本：`model/20261022_OpenAI_v0.6.1.xlsx`（v0.6 原有數值一格不變；新增一頁＋輸入＋檢查）。

## 1. 範圍與不變條件（工程類，chat 端定案）
1. **既有 `TK_Link`（63 名）維持 v5.26 快照不動**：builder 以 `--tk-dir /home/claude/yinchenchang/tk-v526`（Tokenomics `3dd1216688e3872fcb80b28c279d8976a07df29b`，CURRENT v5.26）重建，`check` 既有 63 名值不變。全量同步另案（chat 端列清單給 Andy）。
2. **新增獨立快照頁 `TK_PxQ`**：釘選 Tokenomics v19.0（master `ad33e28`，`/home/claude/yinchenchang/tokenomics-s1`，檔案 `20261022_Tokenomics_v19.0.xlsx`；開工時以 `git log -1` 與 `sha256sum` 驗證並記入頁首）。具名範圍前綴 `TKQ_`（例：`TKQ_IF_HoldEcon`）。名稱清單：`IF_HoldEcon`、`IF_TokGWPaidMix`、`IF_Disc`、`IF_UtilActual`（含 `_Lo`／`_Hi`）、`IF_ProdDerate`、`IF_FleetProdRatio_VR200`、`IF_PxEffMix`、`IF_PxEffMix25`（含 `_Lo`／`_Hi`）、`IF_ServeShare`、`IF_FreeShare`、`IF_MixPaid_Luna／_Sol／_Astra`、`IF_LTotalDisc`（含 `_Lo`／`_Hi`）、`IF_LTotalFleetVDisc`（含 `_Lo`／`_Hi`）、`SRC_DEM_004`、`SRC_DEM_006`、`SRC_DEM_007`、`SRC_DEM_008`、`SRC_DEM_009`。**不得連 `B4_*`、`AL_*`、`L1_*`、`CTL_*`**。builder 新增讀取第二個 Tokenomics 目錄的參數（例如 `--tk-pxq-dir`），兩份快照各自記版本、提交、檔案 SHA-256、讀取日。
3. 新頁 `PxQ`（英文頁名；標籤繁體中文）；公式只引用 `TKQ_`、`INP_`、`SRC_OAI_`、既有 `CMP_`／`REV_` 具名範圍；不含常數（E6）。

## 2. 公司層級輸入（Inputs，新 INP 號；SRC_OAI 用既有列）
| 項目 | 值 | 標記 |
|---|---|---|
| 2025 平均上線 GW（指數內插） | 公式＝(端點差)÷ln(端點比)，端點 `TKQ_SRC_DEM_008` 0.6、`TKQ_SRC_DEM_009` 1.9 → 1.128 | Derived |
| 2025 平均上線 GW（線性內插） | 公式＝端點平均 → 1.25 | Derived |
| 2025 GW 主值選擇 | 開關：指數（基準）／線性 | Decision（chat 端預設，判斷題 5） |
| 2026 平均上線 GW | 3.0（2.5–3.5），三欄情境 2.5／3.0／3.5 | Assumed（Andy「約 3 GW」，官方未公布） |
| 世代組合 | **基準＝OpenAI v0.6 自有 V1 組合**（`CMP_Share_*` 的 2025、2026 欄；自研晶片持有成本沿用 v0.6 S4 做法＝VR200 × `genFactor.custom`）；對照＝交接檔組合 2026 Hopper 0.25／GB200 0.45／GB300 0.25／VR200 0.05（Assumed，Tokenomics Cap_In 同值，新 INP 列） | 判斷題 4 |
| 2025 營收 | `SRC_OAI_002` 13.07（＝`TKQ_SRC_DEM_007`，Checks 驗相等） | 既有 |
| 2026 營收 | 基準 `SRC_OAI_043` 36（管理層目標，FT 2026-09-18 七月投資人簡報；**改標說明：利害關係方的管理層目標，非市場傳聞**）；對照列＝模型正向 `REV_GrossCapped` 2026（34.2） | 判斷題 3 |
| 服務占比 | 基準 `TKQ_IF_ServeShare` 0.41；對照 0.50（Epoch「研發推論各半」，新 SRC_OAI 列，附出處網址；查不到出處則列『已搜尋，未找到』並保留 0.50 為 Assumed 對照值） | 判斷題 2 |
| 生產折減 | 依 Tokenomics README v19.0 讀法：VR200 用 `TKQ_IF_FleetProdRatio_VR200`，其他世代用 `TKQ_IF_ProdDerate`（交接檔一律 0.85 的讀法列對照） | 判斷題 6 |

## 3. 計算（逐年 2025、2026；交接檔第 5 節）
1. 每 GW 年持有成本 h＝Σ 世代占比 × `TKQ_IF_HoldEcon` 基準欄（Hopper 第 2、GB200 第 5、GB300 第 8、VR200 第 11 欄；Rubin Ultra 第 14 欄若占比 >0）。
2. 算力成本 C＝平均 GW × h。
3. Q（M tok／年）＝平均 GW × 服務占比 ×（1 − `TKQ_IF_FreeShare`）× Σ（世代占比 × `TKQ_IF_TokGWPaidMix` 基準欄 × 該世代生產折減）× `TKQ_IF_UtilActual`。
4. 有效牌價 P_eff：2025＝`TKQ_IF_PxEffMix25`，2026＝`TKQ_IF_PxEffMix`；另列表列牌價＝P_eff ÷（1 − `TKQ_IF_Disc`）（顯示 d 未寫死）。
5. 實際 m×L＝營收 ÷（Q × P_eff）；所需 m×L＝C ÷（Q × P_eff）；**比值＝實際 ÷ 所需＝營收 ÷ C**（另列恆等式檢查）。
6. 判斷題 1（L 讀法，預設壽命期價格係數）：另列顯示「隱含 m＝實際 m×L ÷ L」，L 取 `TKQ_IF_LTotalFleetVDisc`（2025 機隊）與 `TKQ_IF_LTotalDisc`（新世代），只作顯示。
7. 2026 持平所需平均 GW＝2026 營收 ÷（2026 h × 2025 比值）；2025 指數／線性兩欄各算。
8. 主輸出：比值的年度序列（2025 指數／線性 × 2026 三情境），一張小表＋讀法一句（Excel 文字公式或固定文字，數字不得手寫）。
9. 敏感度列：服務占比 0.50 時的 Q、m×L（比值不變，標示）；2026 營收用模型正向 34.2 時的比值；世代組合用交接檔組合時的 h 與比值；2026 以 v0.6 合約供給 GW（`CMP_Supply_Total` 或等義名稱的 2026 值）代入時的比值。

## 4. 對照列（不進主序列）
外部報導算力支出：2025 雲端 160 億（Epoch）、2026 500 億（Brockman 證詞）、2025 推論＋訓練 204 億（`TKQ_SRC_DEM_004`＋`TKQ_SRC_DEM_006`）。前兩者登錄新 SRC_OAI 列（出處網址、日期、標記；依 V10 規則初評），各算「÷（GW × h）」倍數。找不到出處者列『已搜尋，未找到』，不得杜撰。

## 5. 驗收
1. **預覽值對帳**（交接檔第 6 節，以交接檔組合與一律 0.85 的對照欄計）：每 GW 持有成本 2025 93.5、2026 97.3 億；算力成本 2025 106／117、2026 243／292／341 億；比值 1.24／1.12、1.48／1.23／1.06；持平 GW 2.99／3.31；2025 Q（1.25 GW）約 6.4×10¹⁵ token。誤差 >0.5% 須說明。基準欄（v0.6 自有組合）與預覽的差異逐項說明。
2. Checks 新增（取下一號並入 registry）：比值恆等式、營收 13.07 兩來源相等、`PxQ` 頁不引用 `B4_`／`AL_`／`L1_`／`CTL_`、TK_Link 63 名值與 v5.26 相同、`TK_PxQ` 版本＝v19.0。
3. **v0.6 原有數值不變**：除新頁、新 Inputs／SRC_OAI 列、Checks、README 外，所有既有公式格逐格與 `20261008_OpenAI_v0.6.xlsx` 相同（表 2＝0）。
4. parity：新增情境（2026 GW、2025 主值開關、組合開關、服務占比），全套分片跑（記憶體約 8GB：`PARITY_SHARD` 分 4 片＋結構測試分開），全過；LibreOffice 重算錯誤 0；`CHK_Errors`＝0。
5. 新增測試：`PxQ` 頁引用白名單（只許 `TKQ_`、`INP_`、`SRC_OAI_`、`CMP_`、`REV_`、本頁名稱）。

## 6. 交付
- 報告 `docs/reports/20261022_v0.6.1_PxQ.md`（結論先行：比值序列與讀法；預覽對帳表；已套用的預設；資料缺口）＋ `_對照.xlsx`（①新增列 ②比值與中間量全表 ③預設）。
- 規格寫入（Project Instructions 第 2 條）：`docs/handoff/OpenAI_handoff.md` 新增一節「P×Q 回本比值規格（Andy 2026-10-21～22 原話）」；`openai/README.md` 加一段；規格中的定義由第 5 節測試把關。
- CHANGELOG、進度檔、PR（內文結尾 `🤖 Generated with [Claude Code](https://claude.com/claude-code)`）、完成留言（第一行 `[OpenAI 回報] v0.6.1｜完成｜YYYY-MM-DD`）、轉 ready。commit 結尾兩行 `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`、`Claude-Session: https://claude.ai/code/session_019uT3zcHojpvDeDuvC6LMXX`。
