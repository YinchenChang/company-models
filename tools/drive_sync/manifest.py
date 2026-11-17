#!/usr/bin/env python3
"""Drive 同步清單：列出每個成品資料夾的現行版本檔與目標 Drive 資料夾（說明見本目錄 README.md）。

用法：
  python3 tools/drive_sync/manifest.py [--root <repo>] [--format csv|json]
  python3 tools/drive_sync/manifest.py --check     # 有 dist/ 的資料夾未列入 manifest.csv 時失敗
標準函式庫即可執行。

現行版本：同一資料夾內檔名為「<YYYYMMDD>_<名稱>_v<版本>.xlsx|.html」者，取版本號最大（同版本取日期最新）
的那一組（xlsx 與同版本的 html 全列）；資料夾內有 CURRENT 檔（如 siliconomics/model/）時以 CURRENT 指定的檔為準；
檔名沒有版本號的 .xlsx／.html（如 markets/invest_map/dist/index.html）一律視為現行。
`_待刪除/` 子資料夾與其他副檔名（.json 等）不列。
"""
import argparse
import csv
import fnmatch
import glob
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
VER = re.compile(r"^(\d{8})_.*_v(\d+(?:[._]\d+)*)\.(xlsx|html)$")
EXTS = (".xlsx", ".html")


def load_manifest(path):
    with open(path, encoding="utf-8", newline="") as f:
        return list(csv.DictReader(f))


def vkey(name):
    m = VER.match(name)
    if not m:
        return None
    return tuple(int(x) for x in re.split(r"[._]", m.group(2))), m.group(1)


def current_files(folder, pattern):
    """folder 內符合 pattern（檔名 glob）的現行版本檔，回傳檔名清單。"""
    names = sorted(n for n in os.listdir(folder)
                   if os.path.isfile(os.path.join(folder, n)) and fnmatch.fnmatch(n, pattern) and n.endswith(EXTS))
    unversioned = [n for n in names if vkey(n) is None]
    versioned = [n for n in names if vkey(n) is not None]
    cur = os.path.join(folder, "CURRENT")
    if os.path.isfile(cur):
        target = open(cur, encoding="utf-8").read().split("\n")[0].strip()
        if target not in names:
            raise SystemExit(f"停止：{cur} 指向的 {target} 不存在")
        best = vkey(target)
    else:
        best = max((vkey(n) for n in versioned), default=None)
    return [n for n in versioned if vkey(n) == best] + unversioned


def resolve(root, rows):
    out = []
    for r in rows:
        g = r["repo_dist_glob"]
        dir_glob, pattern = os.path.split(g)
        for d in sorted(glob.glob(os.path.join(root, dir_glob))):
            if not os.path.isdir(d):
                continue
            rel = os.path.relpath(d, root)
            files = current_files(d, pattern)
            out.append({"folder": rel, "files": [f"{rel}/{n}" for n in files],
                        "drive_folder": r["drive_folder"], "category": r["category"], "project": r["project"],
                        "manifest_glob": g})
    return out


def dist_dirs(root):
    found = []
    for d, subs, _ in os.walk(root):
        subs[:] = [s for s in subs if s not in (".git", "_待刪除", "node_modules")]
        if os.path.basename(d) == "dist":
            found.append(os.path.relpath(d, root))
    return sorted(found)


def check(root, rows):
    dir_globs = [os.path.split(r["repo_dist_glob"])[0] for r in rows]
    missing = [d for d in dist_dirs(root) if not any(fnmatch.fnmatch(d, g) for g in dir_globs)]
    empty = []
    for r in rows:
        dg = os.path.split(r["repo_dist_glob"])[0]
        if "*" in dg:
            continue  # 萬用列（software_ai、physical_ai）可以暫時沒有資料夾
        if not glob.glob(os.path.join(root, dg)):
            empty.append(r["repo_dist_glob"])
    for e in empty:
        print(f"::warning::manifest.csv 的 {e} 在 repo 內找不到資料夾（公司已移除？請更新清單）")
    if missing:
        print("::error::下列有 dist/ 的資料夾未列入 tools/drive_sync/manifest.csv（請補一列：repo_dist_glob,drive_folder,category,project）：")
        for m in missing:
            print(f"  - {m}")
        return 1
    # 每個資料夾都要解得出現行檔
    res = resolve(root, rows)
    nofile = [x["folder"] for x in res if not x["files"]]
    if nofile:
        print("::error::下列資料夾找不到現行版本的 xlsx／html：" + "、".join(nofile))
        return 1
    print(f"通過：{len(dist_dirs(root))} 個 dist/ 資料夾全部列入 manifest.csv；解析出 {len(res)} 個成品資料夾。")
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="Drive 同步清單")
    ap.add_argument("--root", default=os.path.abspath(os.path.join(HERE, "..", "..")), help="repo 根目錄")
    ap.add_argument("--manifest", default=os.path.join(HERE, "manifest.csv"))
    ap.add_argument("--format", choices=("csv", "json"), default="csv")
    ap.add_argument("--check", action="store_true", help="檢查有 dist/ 的資料夾是否都列入清單")
    a = ap.parse_args(argv)
    rows = load_manifest(a.manifest)
    if a.check:
        return check(a.root, rows)
    res = resolve(a.root, rows)
    if a.format == "json":
        json.dump(res, sys.stdout, ensure_ascii=False, indent=1)
        print()
    else:
        w = csv.writer(sys.stdout)
        w.writerow(["repo_file", "drive_folder", "category", "project"])
        for x in res:
            for f in x["files"]:
                w.writerow([f, x["drive_folder"], x["category"], x["project"]])
    return 0


if __name__ == "__main__":
    sys.exit(main())
