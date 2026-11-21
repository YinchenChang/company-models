你在 repo YinchenChang/company-models 的工作樹根目錄。任務：把公司檔案 `markets/invest_map/profiles/{PROFILE}` 裡仍引用舊版個股模型檔、而數值已改變的出處標記，改寫到現行模型。

背景：2026-11-21 全部個股模型換了 2026-11-20 收盤並升版（工作單 `docs/workorders/20261121_px1120_現價與同業報價換11-20收盤.md`）。稽核清單在 `markets/invest_map/data/profile_cites_audit.csv`（欄位 profile、tag、old、cur、status、detail）：只處理 profile＝`{PROFILE}` 的列。status＝DIFF 的 detail 寫了每個儲存格「舊值→新值」；NOCELL 表示現行檔找不到該格。

做法：
1. 讀清單中屬於本檔的每一列，在公司檔案裡找到帶該出處標記的句子或表格列。
2. 用 python3＋openpyxl（data_only=True）打開現行模型檔（cur 欄的檔名，在對應公司資料夾的 dist/ 下；用 `git ls-files | grep` 找路徑），讀標記列出的每個儲存格現值，並看該列 A、B 欄標籤確認還是同一個項目。
3. 把句子裡由這些儲存格支撐的數字改成現值（沿用原句的單位、四捨五入位數與寫法；百分比、倍數、金額照原格式），標記裡的檔名改成現行檔名（保留原本的路徑寫法與儲存格位址；儲存格若在新版移位，改成新位址）。同一句裡由這些數字推出來的描述（例如「上檔空間有限」「評等為放空」「高於現價」）若與新值不符，一併改成與新值相符的寫法。
4. 現行模型已不再支撐的內容（儲存格不存在、項目已移除、或句子講的是舊價位事件）：整句刪除，不要改成沒有出處的句子，也不要自己上網查或估算。
5. 只改 `markets/invest_map/profiles/{PROFILE}` 這一個檔。不得改其他檔、不得 git commit 或 push。不改任何 `[外部…]` 標記的內容。不新增內容。文字用台灣繁體中文，直述、不用行話；出處標記格式照原樣（規則見 `markets/invest_map/README.md` 的「公司檔案」各節）。
6. 改完跑 `python3 markets/invest_map/check_profiles.py`，確認本檔沒有新增的問題；再跑 `python3 markets/invest_map/audit_profile_cites.py` 是不需要的（外層會跑）。

完成後輸出：逐列「標記 → 處理方式（改值／改位址／刪句）＋改後的關鍵數字」，以及刪掉的句子原文。不超過 40 行。
