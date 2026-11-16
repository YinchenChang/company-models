#!/usr/bin/env python3
"""Datadog 個股模型 v0.1 專屬核對：無 # 錯誤、分部對帳為零、三情境 AI 淨效果分化、權重合計 1、FCF 對帳為零、快照版本一致。"""
import json, os, sys
import openpyxl
x = sys.argv[1]; HERE = os.path.dirname(os.path.abspath(__file__))
wb = openpyxl.load_workbook(x, data_only=True); fail = []
for ws in wb:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and (c.value.startswith("#") or c.value.startswith("Err:")): fail.append(f"{ws.title}!{c.coordinate} {c.value}")
PL, V, RC = wb["PL"], wb["Val"], wb["Recon"]
def find(ws, text, col=1):
    for r in range(1, ws.max_row + 1):
        if str(ws.cell(r, col).value or "").startswith(text): return r
for lab in ("核對：三情境 × 全部年度 分部合計", "核對：三情境 × 全部年度 子列合計", "核對：標準損益表 營業利益"):
    r = find(PL, lab)
    vals = [c.value for c in PL[r] if isinstance(c.value, (int, float))]
    if not vals or max(abs(v) for v in vals) > 0.05: fail.append(f"{lab} 非零或缺值")
r = find(V, "AI 淨效果（FY2032E")
a, b, c = (V.cell(r, k).value for k in (2, 3, 4))
if not (a > b > c): fail.append(f"AI 淨效果三情境未分化 {a} {b} {c}")
if abs(V.cell(find(V, "情境權重合計"), 2).value - 1) > 1e-9: fail.append("情境權重合計不為 1")
if abs(V.cell(find(V, "檢查：股票薪酬"), 2).value) > 1e-6: fail.append("落差分解不閉合")
if abs(RC.cell(find(RC, "檢查：公司定義自由現金流"), 3).value) > 1e-6: fail.append("DCF 與公司自由現金流對帳不為零")
snap = json.load(open(os.path.join(HERE, "data", "tokenomics_snapshot_v22.0.json"), encoding="utf-8"))["source"]["version"]
if snap not in str(wb["README"]["A1"].value): fail.append("README 未記 Tokenomics 快照版本")
print("Datadog 專屬核對：" + ("全部通過" if not fail else f"失敗 {len(fail)} 項"))
for f in fail[:20]: print("  -", f)
sys.exit(1 if fail else 0)
