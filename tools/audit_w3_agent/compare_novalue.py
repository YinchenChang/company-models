#!/usr/bin/env python3
"""外部稽核優化第三波（不改值）驗收用：重建成品與 main 成品逐頁逐格比對，差異格分類。

用法：python3 tools/audit_w3_agent/compare_novalue.py <名稱> <新.xlsx> <main.xlsx> [--json out.json]
兩檔都須含快取值（LibreOffice 重算後）。數值容差 1e-9（相對）。In 頁列號位移時按 A 欄鍵比對。

允許的差異類別：
  a  PL 頁標準損益表延伸期的營業費用格、核對列、隱含費用率揭露列（數值）
  b  標籤與說明文字（兩邊都是文字）
  c  build_id 與由腳本雜湊衍生的格（含建置當下的 git 提交雜湊）
  d  企業價值 ≤ 0 時終值占比由數字變「不適用」（會回頭確認 main 檔同欄企業價值 ≤ 0）
  e  新增的提示列與欄（main 為空格；含並列的調整後 Comparable 數值）
  X  其他任何數值格差異（不允許；結束碼 1）
"""
import json, re, sys
import openpyxl
from openpyxl.utils import get_column_letter

TOL = 1e-9
A_PAT = re.compile(r"營業費用|核對|差額|揭露|隱含")
C_PAT = re.compile(r"build_id|-v\d+(\.\d+)+-[0-9a-f]{10}")


def num(v): return isinstance(v, (int, float)) and not isinstance(v, bool)


def same(a, b):
    if num(a) and num(b): return abs(a - b) <= TOL * max(1.0, abs(a), abs(b))
    return a == b


def grid(ws):
    g = {}
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None: g[(c.row, c.column)] = c.value
    return g


def ev_nonpos(gref, r, c):
    """d 類確認：往上 20 列內找 A 欄以「企業價值」或「EV」開頭的列，同欄數值 ≤ 0。"""
    for rr in range(r - 1, max(0, r - 21), -1):
        lab = str(gref.get((rr, 1), "") or "")
        if lab.startswith(("企業價值", "EV（", "EV ", "EV＝")) and "占" not in lab:
            v = gref.get((rr, c))
            if num(v): return v <= 0
    return False


def classify(sheet, r, c, new, old, gnew, gref, lab_new, lab_old):
    lab = f"{lab_new or ''}｜{lab_old or ''}"
    s = f"{new if isinstance(new, str) else ''}{old if isinstance(old, str) else ''}"
    if C_PAT.search(s): return "c"
    if all(isinstance(v, str) and re.fullmatch(r"[0-9a-f]{40}", v) for v in (new, old)): return "c"   # 建置當下的 git 提交（Anthropic OAI_Link）
    if num(new) and num(old):
        return "a" if (sheet == "PL" and A_PAT.search(lab)) else "X"
    if num(old):   # 數字變成文字或消失
        if isinstance(new, str) and new.startswith("不適用") and "終值" in lab and "占" in lab and ev_nonpos(gref, r, c): return "d"
        return "X"
    if num(new):   # main 是空格或文字
        if old is None: return "a" if (sheet == "PL" and A_PAT.search(lab)) else "e"
        return "X"
    if old is None: return "e"
    if isinstance(new, str) and new.startswith("不適用") and isinstance(old, str): return "d" if "終值" in lab else "b"
    return "b"


def main():
    name, fn, fr = sys.argv[1:4]
    wn = openpyxl.load_workbook(fn, data_only=True); wr = openpyxl.load_workbook(fr, data_only=True)
    cnt = dict.fromkeys("abcdeX", 0); bad = []; e_num = []; notes = []
    if wn.sheetnames != wr.sheetnames:
        notes.append(f"工作表名單不同：{wn.sheetnames} ≠ {wr.sheetnames}"); cnt["X"] += 1
    cells = 0
    for sh in wn.sheetnames:
        if sh not in wr.sheetnames: continue
        gn, gr = grid(wn[sh]), grid(wr[sh]); cells += len(gn)
        keyed = False
        if sh == "In":
            an = [gn.get((r, 1)) for r in range(1, wn[sh].max_row + 1)]; ar = [gr.get((r, 1)) for r in range(1, wr[sh].max_row + 1)]
            if len(an) != len(ar) and [x for x in an if x is not None] != [x for x in ar if x is not None]:
                keyed = True; notes.append("In 頁列號位移，按 A 欄鍵比對")
                def bykey(g, a):
                    seen = {}; out = {}
                    for r, k in enumerate(a, 1):
                        if k is None: continue
                        seen[k] = seen.get(k, 0) + 1
                        for (rr, cc), v in g.items():
                            if rr == r: out[((k, seen[k]), cc)] = v
                    return out
                gn, gr = bykey(gn, an), bykey(gr, ar)
        for k in sorted(set(gn) | set(gr), key=str):
            new, old = gn.get(k), gr.get(k)
            if same(new, old): continue
            r, c = k
            if keyed: ln = lo = r[0]; cat = classify(sh, 0, c, new, old, gn, gr, ln, lo); ref = f"{sh}!{get_column_letter(c)}[{r[0]}]"
            else:
                ln, lo = gn.get((r, 1)), gr.get((r, 1)); cat = classify(sh, r, c, new, old, gn, gr, ln, lo); ref = f"{sh}!{get_column_letter(c)}{r}"
            cnt[cat] += 1
            if cat == "X": bad.append({"cell": ref, "new": new, "old": old, "label": str(ln or lo)[:60]})
            if cat == "e" and num(new): e_num.append({"cell": ref, "new": new, "left": str(gn.get((r, c - 1)) if not keyed else "")[:50], "hdr": str(ln)[:40]})
    out = {"name": name, "cells": cells, "counts": cnt, "bad": bad, "e_num": e_num, "notes": notes}
    if "--json" in sys.argv: json.dump(out, open(sys.argv[sys.argv.index("--json") + 1], "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(name, " ".join(f"{k}={v}" for k, v in cnt.items()), f"e內數值格={len(e_num)}", "；".join(notes))
    for b in bad[:30]: print("   X", b)
    sys.exit(1 if cnt["X"] else 0)


if __name__ == "__main__":
    main()
