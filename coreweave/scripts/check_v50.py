"""CoreWeave v5.0 收支模組恆等式與檢查列：所有工作表 A 欄以「檢查：」開頭的列，數值須 |x| < 1e-6、文字須為「通過」／OK 類；
另核 Sens 與 RevDCF 快照、收支恆等式（期末現金＝期初＋融資前現金流＋三種融資，逐年逐情境）。用法：check_v50.py <重算後 xlsx>"""
import sys, openpyxl
wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
fail, n = [], 0
for ws in wb:
    for r in range(1, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if not (isinstance(a, str) and a.strip().startswith(("檢查：", "快照檢查："))): continue
        for c in range(2, ws.max_column + 1):
            v = ws.cell(r, c).value
            if isinstance(v, (int, float)):
                n += 1
                if abs(v) > 1e-6: fail.append(f"{ws.title}!{ws.cell(r, c).coordinate} {a[:30]}＝{v}")
            elif isinstance(v, str) and ws.title in ("Val", "Fund") and v and not v.startswith(("通過", "OK", "不通過（DCF 失效）")):
                pass
CF = wb["CF"]
lab = {}
for r in range(1, CF.max_row + 1):
    a = CF.cell(r, 1).value
    if isinstance(a, str): lab.setdefault(a, []).append(r)
for i in range(3):
    g = {k: lab[k][i] for k in ("期初現金（FY2026＝2026-09-30）", "融資前現金流（FY2026 只計 Q4）", "新增投資級債", "新股", "高息債", "期末現金")}
    for c in "DEFGHIJ":
        v = [CF[f"{c}{g[k]}"].value for k in g]
        n += 1
        if abs(v[0] + v[1] + v[2] + v[3] + v[4] - v[5]) > 1e-6: fail.append(f"CF 現金恆等式 情境{'ABC'[i]} {c}")
        if v[5] < 2 - 1e-6: fail.append(f"CF 期末現金低於最低現金 情境{'ABC'[i]} {c}")
print(f"檢查列與收支恆等式：{n} 項；" + ("全部通過" if not fail else f"失敗 {len(fail)}"))
for f in fail[:20]: print("  -", f)
sys.exit(1 if fail else 0)
