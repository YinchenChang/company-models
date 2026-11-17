#!/usr/bin/env python3
"""PR 路徑歸屬檢查（規則見 docs/OWNERSHIP.md）。

把變更檔分到歸屬群組（AI供應鏈、AI軟體應用、AI實體應用、共用）；
跨兩個以上群組而 PR 標題不以「[跨專案]」開頭時失敗。

用法：
  python3 tools/ownership/check_ownership.py --base <基準 commit> [--head HEAD] --title "<PR 標題>"
  python3 tools/ownership/check_ownership.py --files a/b.py software_ai/x.py --title "<PR 標題>"
標準函式庫即可執行。
"""
import argparse
import subprocess
import sys

TAG = "[跨專案]"
SHARED = "共用(AI供應鏈)"
# 由上而下比對，第一個符合者決定歸屬；以 "/" 結尾者為前綴，其餘為完整路徑
RULES = [
    ("software_ai/", "AI軟體應用"),
    ("physical_ai/", "AI實體應用"),
    ("tools/", SHARED),
    ("markets/invest_map/", SHARED),
    ("docs/plan/", SHARED),
    (".github/", SHARED),
    ("README.md", SHARED),
    ("docs/OWNERSHIP.md", SHARED),
]
DEFAULT = "AI供應鏈"


def owner(path):
    for pat, grp in RULES:
        if (pat.endswith("/") and path.startswith(pat)) or path == pat:
            return grp
    return DEFAULT


def changed_files(base, head):
    out = subprocess.run(["git", "diff", "--name-only", "-z", f"{base}...{head}"],
                         check=True, capture_output=True).stdout.decode("utf-8")
    return [f for f in out.split("\0") if f]


def main(argv=None):
    ap = argparse.ArgumentParser(description="PR 路徑歸屬檢查（docs/OWNERSHIP.md）")
    ap.add_argument("--files", nargs="*", help="直接指定變更檔（本機測試用）")
    ap.add_argument("--base", help="基準 commit（與 --head 比較，三點 diff）")
    ap.add_argument("--head", default="HEAD")
    ap.add_argument("--title", default="", help="PR 標題")
    a = ap.parse_args(argv)
    if a.files is None and not a.base:
        ap.error("需要 --files 或 --base")
    files = a.files if a.files is not None else changed_files(a.base, a.head)

    groups = {}
    for f in files:
        groups.setdefault(owner(f), []).append(f)
    print(f"變更檔 {len(files)} 個，歸屬群組 {len(groups)} 個：")
    for g, fs in sorted(groups.items()):
        print(f"  {g}：{len(fs)} 個（例：{', '.join(fs[:3])}）")

    if len(groups) <= 1:
        print("通過：只動到單一群組。")
        return 0
    if a.title.strip().startswith(TAG):
        print(f"通過：跨 {len(groups)} 個群組，PR 標題已標「{TAG}」（請確認 PR 說明寫明交流訊息或理由）。")
        return 0
    print(f"::error::本 PR 同時改到 {len(groups)} 個歸屬群組（{'、'.join(sorted(groups))}），"
          f"但 PR 標題未以「{TAG}」開頭。請拆成各專案各自的 PR，或在標題前加「{TAG}」"
          f"並於 PR 說明寫明交流訊息或理由（規則見 docs/OWNERSHIP.md）。")
    return 1


if __name__ == "__main__":
    sys.exit(main())
