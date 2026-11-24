你在 repo YinchenChang/company-models 的工作樹根目錄（GitHub Actions 執行環境，不是 /home/claude）。任務：依工作單 `physical_ai/docs/20261123_工作單_新建個股模型第二批起.md`（先完整讀它，並照它列的順序讀原則與規格文件），為下面這一家公司從零建一份個股模型 v0.1 與 HTML 展示層。

公司：{NAME}
資料夾：`physical_ai/{FOLDER}/`
收盤檔代號：{YAHOO}（`physical_ai/data/closes2_20261109_20261120.csv`）
所屬線：{LINE}（線模型成品在 `physical_ai/lines/` 下對應資料夾的 `dist/`）
公司檔案（背景，數字要自己查證）：`physical_ai/map/profiles/{PROFILE}.md`

與工作單不同或補充之處：
1. 工作樹路徑以目前目錄為準；LibreOffice 用私有設定檔（`-env:UserInstallation=file:///tmp/lo_{FOLDER}`），`export HOME=/tmp/h_{FOLDER}`。
2. 成品檔名日期用 20261124：`physical_ai/{FOLDER}/dist/20261124_<公司名>個股模型_v0.1.xlsx` 與同名 `.html`。
3. 只能新增或修改 `physical_ai/{FOLDER}/`。不得 git commit、push、開 PR（外層流程會處理）。
4. 最接近的已完成範本除了工作單列的，還有第二批：日本零組件看 `physical_ai/thk/`、`physical_ai/smc/`；A 股看 `physical_ai/leaderdrive/`；美股虧損公司看 `physical_ai/pony/`、`physical_ai/symbotic/`；港股看 `physical_ai/geekplus/`。整個資料夾照抄結構再改內容，是最省成本的做法。
5. **同業倍數：所有同業報價一律用收盤檔 2026-11-20 收盤**（與標的現價同一天）；收盤檔沒有的同業不要納入中位數，另列「報價日未取得」。這一條是第二批獨立核對的教訓，務必遵守。
6. 2026-11-16 至 11-20 全球 AI 與半導體股大跌屬實（NVIDIA 當週約 −23%、日經 225 當週約 −7.7%），收盤檔的下跌不是資料錯誤。
7. 成本控制：網路搜尋 25 次以內；先把可重建的最小完整模型做出來並通過核對，再補細節。預算有上限，接近上限時優先確保 `dist/` 的 xlsx 與 html 存在、README 寫明未完成項目。
8. 做不完或遇到無法判斷的事：不要硬做。把已完成的部分留在可重建的狀態，在 `physical_ai/{FOLDER}/README.md` 開頭寫清楚卡在哪裡，然後結束。
9. 文字用台灣繁體中文，直述、不用行話。公開 repo：不得出現任何個人或工作情境字樣，也不得出現 `/tmp`、`/home`、`/mnt` 路徑字樣於成品與 README。

完成後照工作單「回報」一節輸出（不超過 45 行，先一行結論）。
