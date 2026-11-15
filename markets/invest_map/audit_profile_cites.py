"""公司檔案出處稽核：找出 profiles/*.md 中引用「已不是現行版」模型 xlsx 的出處標記，逐格比對舊檔與現行檔。

用法：
  python3 markets/invest_map/audit_profile_cites.py            # 只稽核，寫出 data/profile_cites_audit.csv
  python3 markets/invest_map/audit_profile_cites.py --apply    # 另把「SAME」的出處改指現行檔名
  --old-dir <資料夾>   另外提供已從 repo 刪除的舊檔（例如由 git 歷史取回），檔名須與引用相同

判定（每個出處標記一列）：
  SAME      引用的每一格：舊檔值非空且＝現行檔值，且該列 A、B 欄標籤相同 → 可只換檔名（--apply 會換）
  DIFF      至少一格值或標籤不同 → 需逐項對新版數值改寫文字（人工）
  NOOLD     舊檔已不在 repo（含 _待刪除）→ 需人工
  NOCUR     找不到同公司現行檔 → 需人工
  NOCELL    標記中解析不到儲存格 → 需人工
值比較：數值相對差 ≤1e-9 視為相同；文字完全相同。
"""
import csv, glob, os, re, sys
from collections import defaultdict
import openpyxl
from openpyxl.utils import range_boundaries

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
TAG = re.compile(r"\[本專案：([^\]]*?)\]")
FILE = re.compile(r"(\d{8}_[^\s\]：、，（）]+?\.xlsx)")
CELL = re.compile(r"(?:([^\s!，、；：:（）()\[\]÷×+=]+)!)?\$?([A-Z]{1,3})\$?(\d+)(?:\s*[:–~-]\s*\$?([A-Z]{1,3})\$?(\d+))?")

def stem(fn):  # 20261030_SK海力士個股模型_v0.6.1.xlsx → SK海力士個股模型
    s = re.sub(r"^\d{8}_", "", fn)
    return re.sub(r"_v[\d._]+[a-z]?\.xlsx$", "", s)

def index_files():
    cur, old = {}, defaultdict(list)
    for p in glob.glob(os.path.join(ROOT, "**", "*.xlsx"), recursive=True):
        rel = os.path.relpath(p, ROOT)
        if "/node_modules/" in rel:
            continue
        parts = rel.split(os.sep)
        old[os.path.basename(p)].append(p)
        if len(parts) >= 3 and parts[-2] == "dist" and parts[0] != "markets":
            cur.setdefault(parts[0], []).append(p)
    return cur, old

def current_for(oldpath, cur):
    folder = os.path.relpath(oldpath, ROOT).split(os.sep)[0]
    cands = cur.get(folder, [])
    st = stem(os.path.basename(oldpath))
    same = [c for c in cands if stem(os.path.basename(c)) == st]
    if not same:
        same = [c for cs in cur.values() for c in cs if stem(os.path.basename(c)) == st]
    pool = same or [c for c in cands if "個股模型" in c or "收支模型" in c or "模型" in os.path.basename(c)]
    return sorted(pool)[-1] if pool else None

_wb = {}
def wb(p):
    if p not in _wb:
        _wb[p] = openpyxl.load_workbook(p, data_only=True, read_only=False)
    return _wb[p]

def cells_of(text):
    out, sheet = [], None
    for m in CELL.finditer(text):
        sh, c1, r1, c2, r2 = m.groups()
        if sh:
            sheet = sh
        if not sheet:
            continue
        if c2:
            mc, mr, xc, xr = range_boundaries(f"{c1}{r1}:{c2}{r2}")
            if (xc - mc + 1) * (xr - mr + 1) > 60:
                continue
            for r in range(mr, xr + 1):
                for c in range(mc, xc + 1):
                    out.append((sheet, r, c))
        else:
            mc, mr, _, _ = range_boundaries(f"{c1}{r1}")
            out.append((sheet, int(r1), mc))
    return out

def same_val(a, b):
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return a == b or abs(a - b) <= 1e-9 * max(abs(a), abs(b), 1)
    return a == b

def compare(oldp, newp, cells):
    diffs = []
    wo, wn = wb(oldp), wb(newp)
    for sh, r, c in cells:
        if sh not in wo.sheetnames or sh not in wn.sheetnames:
            diffs.append(f"{sh}!缺工作表"); continue
        so, sn = wo[sh], wn[sh]
        vo, vn = so.cell(r, c).value, sn.cell(r, c).value
        lo = (so.cell(r, 1).value, so.cell(r, 2).value)
        ln = (sn.cell(r, 1).value, sn.cell(r, 2).value)
        coord = f"{sh}!{openpyxl.utils.get_column_letter(c)}{r}"
        if vo is None:
            diffs.append(f"{coord} 舊值空白，無法確認")
        elif not same_val(vo, vn):
            diffs.append(f"{coord} 值 {vo!r}→{vn!r}")
        elif lo != ln:
            diffs.append(f"{coord} 標籤 {lo}→{ln}")
    return diffs

def main(apply=False, old_dir=None):
    cur, old = index_files()
    if old_dir:
        for p in glob.glob(os.path.join(old_dir, "*.xlsx")):
            old[os.path.basename(p)].append(p)
    curnames = {os.path.basename(p) for ps in cur.values() for p in ps}
    rows, edits = [], defaultdict(list)
    for prof in sorted(glob.glob(os.path.join(HERE, "profiles", "*.md"))):
        txt = open(prof, encoding="utf-8").read()
        for m in TAG.finditer(txt):
            inner = m.group(1)
            fms = list(FILE.finditer(inner))
            stale = [f for f in fms if f.group(1) not in curnames]
            if not stale:
                continue
            row = {"profile": os.path.basename(prof), "tag": m.group(0), "old": "；".join(f.group(1) for f in stale),
                   "cur": "", "status": "", "detail": ""}
            st, curs, dets, new_tag = "SAME", [], [], m.group(0)
            for i, f in enumerate(fms):
                fn = f.group(1)
                if fn in curnames:
                    continue
                seg = inner[f.end(): fms[i + 1].start() if i + 1 < len(fms) else len(inner)]
                if fn not in old:
                    st = "NOOLD"; break
                oldp = sorted(old[fn], key=len)[0]
                newp = current_for(oldp, cur)
                if not newp:
                    st = "NOCUR"; break
                curs.append(os.path.basename(newp))
                cells = cells_of(seg)
                if not cells:
                    st = "NOCELL"; break
                d = compare(oldp, newp, cells)
                if d:
                    st = "DIFF"; dets += d
                new_tag = new_tag.replace(fn, os.path.basename(newp))
            row["status"], row["cur"], row["detail"] = st, "；".join(curs), "；".join(dets[:6])
            rows.append(row)
            if st == "SAME":
                edits[prof].append((m.group(0), new_tag))
    out = os.path.join(HERE, "data", "profile_cites_audit.csv")
    with open(out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()) if rows else ["profile"])
        w.writeheader(); w.writerows(rows)
    from collections import Counter
    print("舊檔出處標記：", len(rows), dict(Counter(r["status"] for r in rows)))
    print("寫出", os.path.relpath(out, ROOT))
    if apply:
        n = 0
        for prof, pairs in edits.items():
            s = open(prof, encoding="utf-8").read()
            for a, b in pairs:
                if a in s:
                    s = s.replace(a, b); n += 1
            open(prof, "w", encoding="utf-8").write(s)
        print("已改指現行檔名（SAME）：", n, "處標記")

if __name__ == "__main__":
    od = sys.argv[sys.argv.index("--old-dir") + 1] if "--old-dir" in sys.argv else None
    main(apply="--apply" in sys.argv, old_dir=od)
