# tools/in_layout：In 頁版面核對

規格：`docs/plan/個股模型原則_命題加營收分部.md` 第 8 節（「敏感度低／高」與「情境 A／B／C」分開；In 頁最後一節「情境驅動對照」；HTML 輸入表拆兩張）。

- `python3 tools/in_layout/check_in_layout.py <成品 xlsx>`：單檔核對；不適用的模型（無 In 頁三欄或無情境專屬列）回報「不適用」並通過；適用而不符者失敗。建置腳本必跑。
- `python3 tools/in_layout/check_in_layout.py --all [--md 報告.md]`：掃描全部 `*/dist/*.xlsx`，只報告。
- 檢查：D3／F3 欄名；「情境驅動對照」恰好一節、表頭 D／E／F＝情境 A／B／C；每格為 `=E<列>` 且引用 H 欄「引用 In 頁 k1／k2／k3」所列鍵（情境 B 可為常數 0）；上方所有情境專屬列都被引用（判定：鍵名同族兄弟或項目文字「情境 A：」等，見程式開頭註解）；`scn_w` 合計＝1；本節未被任何公式或定義名稱引用。
