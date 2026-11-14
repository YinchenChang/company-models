# tools/coord — 三層協調 C4：對帳中標示與兩讀數目標價（共用工具）

規則：`docs/plan/三層基準協調規則.md` v1.4 C4（個股不裁決；估值依賴對帳中年份的硬體個股並列兩讀數目標價）。樣板 NVIDIA v0.3.8（`nvidia/coord_post.py`，專用）；Broadcom v0.5.5 起用本工具。

## 每家要做的事
1. **主引擎掛鉤（每家一處，依該模型的 Wattonomics 錨定參數名）：** `SIL.COORD_ALT`（環境變數 `SIL_COORD_ALT=1`）時，情境 B、C 於 CY2027 起產業量 × `wta_27`／`wta_28`／`wta_29`；情境 A 不縮放。該模式下 `siliconomics/reader.py` 的 `WTA` 改為另一側讀數比值（只在對帳中年份，其他年份＝1）。
2. `python3 tools/coord/alt_run.py <資料夾>` → `<資料夾>/coord_alt.json`（建置兩次、LibreOffice 重算、讀 Val 頁；檢查另一側建置的 wta 列確實＝比值、錯誤格 0）。
3. 建檔程式存檔前：`COORD.add_sheet(wb, <資料夾>, SIL)`（新增 Coord 頁；回傳非 None 時 `COORD.stamp(...)` 升版號）。
4. `gen_html_json.py` 最後：`COORD.inject_html(cfg, xlsx)`（橫幅一格＋揭露句、估值章開頭兩讀數表與長條圖）。
5. `siliconomics/builder/seed/coord_meta_v*.json` 的 `hardware_companies` 加入該資料夾，重匯 `baseline_status.json`；`tools/baseline_coordination/check.py` 第 6 項須通過。

## 重跑時機
Siliconomics 升版、該模型輸入改變後：先 `alt_run.py`，再正式建置；否則 Coord 頁「另一側讀數快照狀態」顯示 STALE。
