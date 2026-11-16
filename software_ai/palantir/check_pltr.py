# Palantir 成品核對：公式錯誤 0、各頁「檢查：」列為 0、Val 必備列存在、Sens 模型狀態 CURRENT 且基準差為 0、README 無手寫舊數字（由 display_check 另查）。
# 用法：python3 software_ai/palantir/check_pltr.py <重算後的 xlsx>；失敗時非零結束。
import sys
import openpyxl

wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
fail = []
ERR = ("#REF", "#DIV", "#VALUE", "#NAME", "#N/A", "#NUM", "Err:")
n_err = 0
for ws in wb:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith(ERR):
                n_err += 1
                if n_err <= 10: fail.append(f"公式錯誤 {ws.title}!{c.coordinate} {c.value}")
print(f"公式錯誤：{n_err} 格")
def nums(ws, r): return [c.value for c in ws[r][1:] if isinstance(c.value, (int, float)) and not isinstance(c.value, bool)]
n_chk = {}
for sh in ("Q", "Seg", "PL", "CF", "Recon"):
    ws = wb[sh]; n_chk[sh] = 0
    for r in range(1, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if not (isinstance(a, str) and a.startswith("檢查：")): continue
        if sh == "Recon" and "反解後目標價" in a: continue      # 反解快照列另查
        vs = nums(ws, r); n_chk[sh] += 1
        if not vs: fail.append(f"{sh}!A{r}「{a}」沒有數值")
        elif any(abs(v) > 1e-6 for v in vs): fail.append(f"{sh}!A{r}「{a}」有非零值 {[round(v, 6) for v in vs if abs(v) > 1e-6][:5]}")
print("核對列：" + "、".join(f"{k} {v} 列" for k, v in n_chk.items()))
for sh, need in (("Q", 2), ("Seg", 3), ("PL", 9), ("CF", 3), ("Recon", 1)):
    if n_chk[sh] < need: fail.append(f"{sh} 核對列少於 {need}（{n_chk[sh]}）")
V = wb["Val"]; labs = {str(V.cell(r, 1).value): r for r in range(1, V.max_row + 1)}
for n in ("目標價（DCF 與 Comparable 加權）", "上檔空間：加權目標價", "投資建議", "DCF 目標價（情境加權）", "Comparable 目標價（情境加權）", "估值評等（依上檔空間五級）", "入不敷出判定用：FY2026–30 自由現金流扣股利合計（B）", "現價（US$）"):
    if n not in labs: fail.append(f"Val 缺列「{n}」")
se = wb["Sens"]
for r in range(1, se.max_row + 1):
    a = str(se.cell(r, 1).value or "")
    if a.startswith("敏感度重算時的基準值"):
        d = se.cell(r, 5).value
        if not isinstance(d, (int, float)) or abs(d) > 1e-6: fail.append(f"Sens 敏感度已過期或缺：基準差 {d}")
    if a.startswith("模型狀態") and se.cell(r, 4).value != "CURRENT": fail.append(f"Sens 模型狀態 {se.cell(r, 4).value}")
# Scenario A/B/C 的 AI 淨效果必須有差異（工作單 G）
S = wb["Seg"]; nets = {}
for r in range(1, S.max_row + 1):
    if str(S.cell(r, 1).value or "").startswith("AI 淨效果") and isinstance(S.cell(r, 8).value, (int, float)): nets[len(nets)] = S.cell(r, 8).value
if len(nets) != 3 or len(set(round(v, 3) for v in nets.values())) != 3: fail.append(f"三情境 FY2028 AI 淨效果未分化：{nets}")
tp = V.cell(labs["目標價（DCF 與 Comparable 加權）"], 2).value; up = V.cell(labs["上檔空間：加權目標價"], 2).value; rec = V.cell(labs["投資建議"], 2).value
print(f"加權目標價 {tp:.2f}；上檔空間 {up:+.1%}；投資建議 {rec}；三情境 FY2028 AI 淨效果 {[round(v, 1) for v in nets.values()]}")
if fail:
    print("失敗：")
    for f in fail: print("  -", f)
    sys.exit(1)
print("全部通過")
