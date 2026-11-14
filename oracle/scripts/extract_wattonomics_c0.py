#!/usr/bin/env python3
"""由 Wattonomics 快照（wattonomics/data/wattonomics_snapshot_v0.3.2.json）摘錄 Oracle C0 對帳用的最小資料：
WIF_ 名稱中含 orcl 者，加上 In 頁 id 含 orcl 的列（不含園區列 cp_*）與共用參數 ai_sh、k_chip2it、tolband。
寫入 oracle/data/wattonomics_c0_v0.3.2.json（不進模型計算；Coord 頁只讀此檔）。
用法：python3 oracle/scripts/extract_wattonomics_c0.py
"""
import json, os

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
S = json.load(open(os.path.join(ROOT, "wattonomics", "data", "wattonomics_snapshot_v0.3.2.json"), encoding="utf-8"))
names = {k: v for k, v in S["names"].items() if "orcl" in k}
keep = {"ai_sh", "k_chip2it", "tolband"}
rows = {}
for r in S["rows"]:
    i = r.get("id")
    if i and ((("orcl" in i) and not i.startswith("cp_")) or i in keep):
        rows[i] = {k: r.get(k) for k in ("label", "basis", "unit", "tag", "src", "note", "values")}
out = {"source": {"file": S["source"]["file"], "version": S["source"]["version"], "sha256": S["source"]["sha256"],
                  "snapshot": "wattonomics/data/wattonomics_snapshot_v0.3.2.json（main，2026-11-14）",
                  "note": "Oracle C0 對帳用的最小摘錄（只讀 WIF_ 名稱值與 In 列；不進模型計算）；In 列 values 的第 0 欄＝In 頁 H 欄（2023），第 3 欄＝K 欄（2026）"},
       "names": names, "rows": rows}
json.dump(out, open(os.path.join(ROOT, "oracle", "data", "wattonomics_c0_v0.3.2.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(len(names), "names;", len(rows), "rows")
