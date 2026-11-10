#!/usr/bin/env python3
"""Oracle v0.5.0 收支模組核對：恆等式（應為 0 的檢查列）與建置快照未過期。
用法：python3 oracle/scripts/check_funding_v05.py 成品.xlsx rowmap.json
"""
import json, sys
import openpyxl

wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
F = json.load(open(sys.argv[2], encoding="utf-8"))["FUND"]
bad = []; n = 0


def zero(sheet, cell, what, tol=1e-6):
    global n
    v = wb[sheet][cell].value; n += 1
    if not isinstance(v, (int, float)) or abs(v) > tol: bad.append(f"{sheet}!{cell} {what}＝{v}")


fu = F["Fund"]
for c in "CDE": zero("Fund", f"{c}{fu['SUM']['chk']}", "現金橋差異合計")
for s, k in fu["R"].items():
    for c in "DEFGHI":
        zero("Fund", f"{c}{k['chk1']}", f"情境 {s} 融資前現金流對 CF 頁"); zero("Fund", f"{c}{k['chk2']}", f"情境 {s} 現金橋")
if fu["STRESS"]:
    zero("Fund", f"D{fu['STRESS']['_chk']}", "壓力表快照（瀑布融資）"); zero("Fund", f"D{fu['STRESS']['_chk2']}", "壓力表快照（目標價）")
else: bad.append("Fund 壓力表不存在（STRESS=1 bash oracle/make_v05.sh）")
d = F["Debt"]
for c in "DEFGH": zero("Debt", f"{c}{d['L']['chk']}", "逐檔到期 − 10-Q 申報到期表", 5e-4)
zero("Debt", f"E{d['n1'] + 2}", "新票據合計 − ex_nt_par")
zero("Lease", "D9", "在帳租賃付款總額 − 申報", 5e-4)
v = wb["RPO"]["D7"].value; n += 1
if not isinstance(v, (int, float)) or abs(v - 0.7975) > 1e-9: bad.append(f"RPO!D7 排程權重合計＝{v}（應為 0.7975）")
r = F["RevDCF"]["R"]
if "_chk" in r: zero("RevDCF", f"D{r['_chk']}", "反向 DCF 快照")
else: bad.append("RevDCF 未求解（REV=1 bash oracle/make_v05.sh）")
print(f"收支模組核對：{n} 項，失敗 {len(bad)} 項")
if bad: sys.exit("失敗：\n" + "\n".join(bad))
