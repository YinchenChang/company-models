#!/usr/bin/env python3
"""Naver 個股模型 v0.2 專屬核對：無 #錯誤、分部對帳為零、AI 恆等式為零、三情境 AI 淨效果分化、權重合計 1、落差分解合計為零、Tokenomics 快照版本一致、合理性關卡診斷存在。"""
import json, sys, os
import openpyxl
x = sys.argv[1]
HERE = os.path.dirname(os.path.abspath(__file__))
wb = openpyxl.load_workbook(x, data_only=True)
fail = []
for ws in wb:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and (c.value.startswith("#") or c.value.startswith("Err:")):
                fail.append(f"{ws.title}!{c.coordinate} {c.value}")
PL, V = wb["PL"], wb["Val"]
def find(ws, text, col=1):
    for r in range(1, ws.max_row + 1):
        if str(ws.cell(r, col).value or "").startswith(text): return r
for s in "ABC":
    for lab, tol in ((f"情境 {s}：分部合計 − 合併營收", 0.05), (f"情境 {s}：AI 恆等式", 1e-6), (f"情境 {s}：子線合計 − 分部", 0.15)):
        v = PL.cell(find(PL, lab), 3).value
        if v is None or v > tol: fail.append(f"{lab} 差 {v}")
r = find(V, "AI 淨效果（營收層")
a, b, c = (V.cell(r, k).value for k in (2, 3, 4))
if not (a > b > c): fail.append(f"AI 淨效果三情境未分化 {a} {b} {c}")
if abs(V.cell(find(V, "情境權重合計"), 2).value - 1) > 1e-9: fail.append("情境權重合計不為 1")
if abs(V.cell(find(V, "檢查：分項合計 − 落差"), 2).value) > 1e-6: fail.append("DCF 與 Comparable 落差分解不為零")
snap = json.load(open(os.path.join(HERE, "data", "tokenomics_snapshot_v22.0.json"), encoding="utf-8"))["source"]["version"]
if snap not in str(wb["README"]["A1"].value): fail.append("README 未記 Tokenomics 快照版本")
if "合理性關卡診斷" not in str(wb["README"]["A30"].value): fail.append("README 缺合理性關卡診斷")
print("Naver 專屬核對：" + ("全部通過" if not fail else f"失敗 {len(fail)} 項"))
for f in fail[:20]: print("  -", f)
sys.exit(1 if fail else 0)
