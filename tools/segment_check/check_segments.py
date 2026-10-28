#!/usr/bin/env python3
"""個股模型分部檢查（docs/plan/個股模型原則_命題加營收分部.md 第 3 條）。

用法：
  python3 tools/segment_check/check_segments.py <公司資料夾>      # 檢查一家，失敗時非零結束
  python3 tools/segment_check/check_segments.py --all [--md 報告.md]  # 掃描全部有 dist/*.xlsx 的資料夾，只報告不中斷

每家公司在資料夾放 segments.json：
{
  "company": "Microsoft", "ticker": "MSFT",
  "model_sheet": "PL",                       # 分部列所在工作表
  "official_segments": [                      # 官方報告分部（10-K／年報）
    {"name": "Agents and Infra", "label": "Agents and Infra", "period": "FY2026",
     "value": 268.127, "unit": "US$B", "source": "8-K 2026-09-02 附件 99.1"}
  ],
  "sublines": [ {"parent": "Agents and Infra", "labels": ["Azure", "Microsoft 365 cloud", ...]} ],
  "check_labels": ["檢查：分部合計 − 合併營收"]   # 核對列標籤；該列所有數值須 |x| < tol
}
檢查：(a) 每個官方分部的 label 在 model_sheet 的 A／B 欄出現；(b) check_labels 的列存在且數值皆近 0；
(c) sublines 的每個 label 都存在（子列合計＝母分部由模型內的 check 列負責）。
"""
import glob
import json
import os
import sys

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
TOL = 0.05


def latest_xlsx(folder):
    fs = sorted(f for f in glob.glob(os.path.join(folder, "dist", "*.xlsx")) if "_待刪除" not in f)
    return fs[-1] if fs else None


def labels_in(ws, maxcol=3):
    out = {}
    for r in ws.iter_rows(min_row=1, max_col=maxcol):
        for c in r:
            if isinstance(c.value, str):
                out.setdefault(c.value.strip(), c.row)
    return out


def check(folder):
    """回傳 (狀態, 訊息清單)；狀態：符合／不符／未宣告。"""
    msgs = []
    spec_f = os.path.join(folder, "segments.json")
    x = latest_xlsx(folder)
    if not x:
        return "無成品", ["dist/ 沒有 xlsx"]
    if not os.path.exists(spec_f):
        return "未宣告", ["沒有 segments.json（依原則第 3 條須宣告官方分部與對應列）"]
    spec = json.load(open(spec_f, encoding="utf-8"))
    wb = openpyxl.load_workbook(x, data_only=True, read_only=False)
    sh = spec.get("model_sheet", "PL")
    if sh not in wb.sheetnames:
        return "不符", [f"找不到工作表 {sh}"]
    ws = wb[sh]
    lab = labels_in(ws)
    ok = True
    for s in spec.get("official_segments", []):
        if s["label"] not in lab:
            ok = False
            msgs.append(f"官方分部「{s['name']}」在 {sh} 頁沒有列（標籤「{s['label']}」）")
    for g in spec.get("sublines", []):
        for l in g["labels"]:
            if l not in lab:
                ok = False
                msgs.append(f"子列「{l}」（屬 {g['parent']}）在 {sh} 頁沒有列")
    for cl in spec.get("check_labels", []):
        if cl not in lab:
            ok = False
            msgs.append(f"核對列「{cl}」不存在")
            continue
        r = lab[cl]
        vals = [c.value for c in ws[r] if isinstance(c.value, (int, float))]
        bad = [v for v in vals if abs(v) >= TOL]
        if not vals:
            ok = False
            msgs.append(f"核對列「{cl}」沒有數值")
        elif bad:
            ok = False
            msgs.append(f"核對列「{cl}」有非零值：{bad[:5]}")
    if ok:
        msgs.append(f"{len(spec.get('official_segments', []))} 個官方分部、{sum(len(g['labels']) for g in spec.get('sublines', []))} 條子列、{len(spec.get('check_labels', []))} 個核對列皆通過（{os.path.basename(x)}）")
    return ("符合" if ok else "不符"), msgs


def main():
    if len(sys.argv) >= 2 and sys.argv[1] == "--all":
        md = sys.argv[sys.argv.index("--md") + 1] if "--md" in sys.argv else None
        rows = []
        for d in sorted(os.listdir(ROOT)):
            p = os.path.join(ROOT, d)
            if os.path.isdir(os.path.join(p, "dist")) and latest_xlsx(p):
                st, m = check(p)
                rows.append((d, st, "；".join(m)))
        lines = ["| 資料夾 | 狀態 | 說明 |", "|---|---|---|"] + [f"| {a} | {b} | {c} |" for a, b, c in rows]
        txt = "\n".join(lines)
        print(txt)
        if md:
            open(md, "w", encoding="utf-8").write(txt + "\n")
        return 0
    folder = sys.argv[1]
    st, m = check(os.path.abspath(folder))
    print(f"[{st}] {folder}")
    for x in m:
        print("  -", x)
    return 0 if st == "符合" else 1


if __name__ == "__main__":
    sys.exit(main())
