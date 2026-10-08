#!/usr/bin/env python3
"""TK_Link 快照 vs Tokenomics master 的 CURRENT 逐名比對（工作單第 1 節第 4 點；補充 1 第 1 點）。

用法：python3 tools/check_tk_snapshot.py --tk-dir <Tokenomics checkout> [--model model/xxx.xlsx] [--md out.md]
輸出：每個名稱 OK／DIFF／MISSING／PENDING。DIFF 與 MISSING 只報 WARN（結束碼 0，不擋合併）。
- MISSING：快照有值但 Tokenomics 現行版已無此名稱，或 Tokenomics 端尚未提供（狀態非 OK 者：Tokenomics 仍無此名稱報 MISSING，已提供報 NOW_AVAILABLE）。
- 不從未合併分支取值。
- Anthropic A2（規格 D6）：狀態「讀表（非具名）」的列（NonNV；A3 起加 Inputs 頁 PUE），以列標籤＋欄標題在 Tokenomics 該頁找值比對（名稱格式 <頁>!<列標籤>!<欄標題>）；
  找不到報 MISSING（Tokenomics 改了表的列標籤或欄標題）。
"""
import argparse
import re
import sys
from pathlib import Path

import openpyxl

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
from engine import current_model_path  # noqa: E402
sys.path.insert(0, str(REPO / "builder"))
from tk_link import NONNV_STATUS, table_lookup  # noqa: E402


def tk_values(tk_dir: Path):
    cur = (tk_dir / "model" / "CURRENT").read_text(encoding="utf-8").strip()
    wb = openpyxl.load_workbook(tk_dir / "model" / cur, data_only=True)
    out = {}
    for n, d in wb.defined_names.items():
        m = re.match(r"^(?:'([^']+)'|([^!]+))!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$", d.attr_text)
        if not m:
            continue
        ws = wb[m.group(1) or m.group(2)]
        from openpyxl.utils import column_index_from_string as ci
        c1, r1 = ci(m.group(3)), int(m.group(4))
        c2 = ci(m.group(5)) if m.group(5) else c1
        r2 = int(m.group(6)) if m.group(6) else r1
        out[n] = [ws.cell(r, c).value for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]
    return cur, out, wb


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tk-dir", required=True, type=Path)
    ap.add_argument("--model", type=Path)
    ap.add_argument("--md", type=Path)
    a = ap.parse_args()
    cur, tk, tkwb = tk_values(a.tk_dir)
    wb = openpyxl.load_workbook(a.model or current_model_path(), data_only=False)
    ws = wb["TK_Link"]
    snap_file, snap_ver = ws["B3"].value, ws["B4"].value
    rows, bad = [], 0
    for r in range(10, ws.max_row + 1):
        name = ws.cell(r, 1).value
        if not name:
            continue
        status, n = ws.cell(r, 6).value, ws.cell(r, 5).value or 0
        snap = [ws.cell(r, 10 + k).value for k in range(n)]
        if status == NONNV_STATUS:                       # 讀表列：NonNV!<列標籤>!<欄標題>
            sheet, label, hdr = name.split("!")
            v, addr = table_lookup(tkwb, sheet, label, hdr)
            if addr is None:
                rows.append((name, "MISSING", f"Tokenomics {sheet} 表找不到此列標籤或欄標題"))
                bad += 1
                continue
            tk[name] = [v]
        elif status != "OK":
            rows.append((name, "MISSING" if name not in tk else "NOW_AVAILABLE", status))
            if name in tk:
                bad += 1
            continue
        if name not in tk:
            rows.append((name, "MISSING", "Tokenomics 現行版無此名稱"))
            bad += 1
            continue
        same = len(tk[name]) == len(snap) and all(
            (x == y) or (isinstance(x, (int, float)) and isinstance(y, (int, float)) and abs(x - y) <= 1e-9 * max(abs(x), abs(y), 1e-300))
            for x, y in zip(tk[name], snap))
        rows.append((name, "OK" if same else "DIFF", ""))
        bad += 0 if same else 1
    from collections import Counter
    cnt = Counter(s for _, s, _ in rows)
    head = f"快照 {snap_file}（{snap_ver}） vs Tokenomics CURRENT {cur}：{dict(cnt)}"
    print(head)
    for n, s, why in rows:
        if s != "OK":
            print(f"::warning::TK 快照 {s}：{n} {why}")
    if a.md:
        a.md.write_text("\n".join([f"# {head}", "", "| 名稱 | 結果 | 說明 |", "|---|---|---|"] + [f"| {n} | {s} | {w} |" for n, s, w in rows]) + "\n", encoding="utf-8")
    sys.exit(0)        # WARN 不擋合併


if __name__ == "__main__":
    main()
