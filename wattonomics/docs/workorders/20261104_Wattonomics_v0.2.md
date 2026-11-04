# 工作單：Wattonomics v0.2＝承租玩家只算 IT＋對帳點修正＋刪停用情境＋K4 乙＋汰換列

- 日期：2026-11-04（台北）；chat 端：session_01V11BcwWqQyh9suCNhAYQKQ
- 現行：main（v0.1.9＋#301、#307 頁面）；分支 `claude/wattonomics-v0.2`
- 依據：Andy 11-04 08:06（承租玩家 IT-only）、11-04 07:44（不寫需求 A）、11-04 07:47（黃仁勳每 GW＝晶片口徑）、11-02 18:36（K4 乙）、11-02 15:58／16:22（三層只給基準）；診斷報告 PR #305 `wattonomics/docs/verify/20261103_2027時差來源診斷.md`（v0.2 建議「現在可做」）、PR #310 `wattonomics/docs/verify/20261104_Oracle園區與資本支出差異診斷.md`（若尚未合併，讀分支 `claude/wattonomics-2027-diagnosis`、`claude/wattonomics-oracle-gap`）；AI半導體 11-02 要求保留三列名。
- 原則：基準數值**會**變（這是升版）；每項改動在建置報告列出 v0.1.9 → v0.2 的影響（市場 GW、11 家 Add11、Gap 判讀、R4 比值、NVIDIA 對帳點）。

## 1. 改動（工程決定，chat 端已定）
| # | 改動 |
|---|---|
| V1 | **承租玩家只算 IT（Oracle、CoreWeave、Nebius）**：In 新增逐家「承租比例」（Oracle 1.0〔CFO：不含土地建物電力〕；CoreWeave、Nebius 依各自個股模型／10-Q 自建 vs 承租 MW，查不到給 Assumed 區間）；CapexGW 甲2 每 GW 成本＝IT＋（1−承租比例）×廠房；Gap 頁 AI 需要同一成本；新增「市場其餘——出租方廠房資本支出」列（承租 GW × 廠房成本，廠房提前 lag_bldg；只入 Spend 與市場參考，不入 11 家收支）；營業現金流不另扣租金（個股模型已扣，Oracle CF!55）；註明融資租賃本金不在營業現金流。Decisions：GL01-21 改為 Andy 11-04 決定，引原話。 |
| V2 | Oracle AI 占比改用 Oracle 個股模型 AI／非 AI 資本支出拆分（Y9）。 |
| V3 | T2 四座 2028 Oracle 園區（Shackelford 其餘、Jupiter、Port Washington、Saline 其餘）分期：有棟別時程用之，否則 18 個月等量分期（終點＝公告完工）；Shackelford 規模查證（Vantage 1.4 GW vs Epoch 2.0）：查得者改值並標來源，查不到維持並標 Assumed 區間。Checks 的 Oracle 站點 vs 甲2 比較改看年底存量（WARN 門檻沿用容許帶）。 |
| V4 | NVIDIA 對帳點（Compare）：七家與 xAI 直接用付款年 GW（不由上線年反推）；NVIDIA GW 逐家計算（各家非 NVIDIA 占比）、中國 NVIDIA 占比 0（2025–2027 一致）；新增「2026 校準係數」「所需內容成長（2027÷2026）」兩列，判定以成長比較；營收口徑內容（26.13）備忘列；黃仁勳每 GW 備忘註明「晶片口徑（Andy 11-04）」。R7 收件人改 AI半導體（Requests 頁）。 |
| V5 | **刪停用情境**：需求 B、C 與延遲 D1、D2 的欄位、公式、Checks、`WIF_` 名稱、explorer json 紀錄全部刪除。**保留三列名**：Deploy「年新增 IT GW（毛額…）」「年底上線 IT GW」、Compare「瓦帳 年新增 IT GW（淨；SiteOnly）」（標籤可去掉「A、D0」欄位文字，但列名開頭不變）。基準名稱：`*_A_D0_*` 保留（下游在讀），另新增無情境字樣的別名 `WIF_MktGW_{YE,Avg,Add}_<yyyy>`、`WIF_PlayerGW…` 已無情境字樣者不動；Interface 說明「A_D0 為舊名、v0.3 刪」。Excel 標籤、README、Decisions 文字一律不寫「需求 A」，改「基準（公告時程＋已確認延遲）」。 |
| V6 | **K4 乙**：SiteOnly 2030–2032 改直接加總站點證據（不再用前年 × g）；主路徑 2030–32 不變。Coverage 名稱定義隨之更新。 |
| V7 | **汰換列**：Deploy 新增汰換 GW（壽命 6 年〔TK 折舊年限〕，Assumed 5–7；世代別依 Players 世代組合）與淨新增列；既有「毛額」列不改值；R4 比較讀淨額或毛額由 AI半導體 決定（通知時列出）。 |
| V8 | 雜項：In `cap_g_ext` 說明半句（v0.1.9 核對）；TK 快照換 master 最新（v21.8；若 v21.9 已合併則 v21.9），只影響高成本欄與 NonNV 產出比 → 報告列影響；`mw_fy26` 單位標示（Oracle 診斷核對小註）。 |

## 2. 建置與驗收
- build.py `--base <現行> --version v0.2 --date <YYYYMMDD>`；soffice；export_snapshot、export_explorer；check.py；diff_snapshot（v0.1.9 → v0.2：列出刪除名稱數、新增名稱、既有名稱變動前 30 大與分類）。
- 重跑 r4_attrib（單一基準比較）→ R4 r3；重算 NVIDIA 對帳點。
- 頁面：build_page 本機產生並以無頭瀏覽器驗證（無「需求 A」「情景」字樣、JS 錯誤 0），不發布。
- 報告 `docs/verify/<日期>_v0.2_build.md`（STATUS 格式），含 v0.1.9 → v0.2 影響表。chat 端另派未參與代理核對。
- 不做：不改 TK、Siliconomics 檔；不發布；不開 PR、不合併。

## 3. 追加（chat 端 2026-11-04；Andy 08:21「融資租賃『處理上應該就要走資本開支那條線』」）
| # | 改動 |
|---|---|
| V1b | **融資租賃走資本支出線、營業租賃走營業現金流**：In 逐家（Oracle、CoreWeave、Nebius）新增「融資租賃占比」（承租廠房中屬融資租賃者；Oracle 預設＝FY27Q1 帳上融資租賃負債 ÷ 租賃負債合計＝9.185 ÷（34.621＋9.185）≈ 0.21，Derived，區間 0–0.40 Assumed〔2,880 億未起租的分類未揭露〕；CoreWeave、Nebius 依最新 10-Q／20-F 或個股模型，否則 Assumed 區間）。資本支出線（CapexGW 甲2 每 GW 成本與 Gap AI 需要）：IT ＋（1 − 承租比例 ×（1 − 融資租賃占比））× 廠房——融資租賃廠房算承租方資本支出（會計擁有者），營業租賃廠房歸出租方列。資金來源：融資租賃廠房金額為融資來源（同 Funding 頁「新增融資租賃」類），不是營業現金流。營業現金流一致性：個股模型（Oracle v0.4.2 Thesis 租金 → CF!55）把全部未起租租金列為營業；融資租賃部分把租金中的本金加回營業現金流（利息留在營業現金流），寫明公式，避免廠房成本重複計入（資本支出需要一次、租金經營業現金流一次）。出租方列只含營業租賃部分（不與承租方重複）。影響表單列 V1b 效果；Decisions 引 Andy 原話。 |
