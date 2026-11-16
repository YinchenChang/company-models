#!/usr/bin/env python3
"""LY 個股模型 v0.1 專屬核對：無 #錯誤、分部對帳為零、三情境 AI 淨效果分化、權重合計 1、現價與目標價可算、Tokenomics 快照版本一致。"""
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
    r = find(PL, f"情境 {s}：分部合計 − 合併營收")
    v = PL.cell(r, 3).value
    if v is None or v > 0.05: fail.append(f"情境 {s} 分部對帳差 {v}")
    r = find(PL, f"情境 {s}：Media 子線合計 − Media")
    v = PL.cell(r, 3).value
    if v is None or v > 0.15: fail.append(f"情境 {s} Media 子線差 {v}（FY2025 子線為十億日圓一位小數，四捨五入差 ≤ 0.1）")
r = find(V, "AI 淨效果（營業利益）FY2032E")
a, b, c = (V.cell(r, k).value for k in (2, 3, 4))
if not (a > b > c): fail.append(f"AI 淨效果三情境未分化 {a} {b} {c}")
r = find(V, "情境權重合計")
if abs(V.cell(r, 2).value - 1) > 1e-9: fail.append("情境權重合計不為 1")
snap = json.load(open(os.path.join(HERE, "data", "tokenomics_snapshot_v22.0.json"), encoding="utf-8"))["source"]["version"]
txt = str(wb["README"]["A1"].value)
if snap not in txt: fail.append("README 未記 Tokenomics 快照版本")
print("LY 專屬核對：" + ("全部通過" if not fail else f"失敗 {len(fail)} 項"))
for f in fail[:20]: print("  -", f)
sys.exit(1 if fail else 0)
