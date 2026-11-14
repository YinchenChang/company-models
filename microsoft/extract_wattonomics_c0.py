#!/usr/bin/env python3
"""Microsoft C0 對帳用的 Wattonomics 摘錄（只讀；不進模型計算）。
用法：python3 microsoft/extract_wattonomics_c0.py <追公式輸出 json> [--out-dir microsoft/data]
  追公式輸出 json：python3 tools/coord/trace_wattonomics.py wattonomics/model/20261114_Wattonomics_v0.3.2.xlsx <json> --player msft
輸出：
  microsoft/data/wattonomics_c0_v0.3.2.json    快照名稱值與 In 頁相關列（msft）
  microsoft/data/wattonomics_trace_v0.3.2.json 追公式表：只留「生效」格（任一輸出根生效）與非園區旗標格；園區旗標格只計數
"""
import argparse
import json
import os

import openpyxl

ap = argparse.ArgumentParser()
ap.add_argument("trace")
ap.add_argument("--out-dir", default="microsoft/data")
a = ap.parse_args()
SNAP = "wattonomics/data/wattonomics_snapshot_v0.3.2.json"
XLSX = "wattonomics/model/20261114_Wattonomics_v0.3.2.xlsx"
s = json.load(open(SNAP, encoding="utf-8"))
names = {k: v for k, v in s["names"].items() if "_msft_" in k or k.endswith("_msft")}
wb = openpyxl.load_workbook(XLSX, data_only=True)
w = wb["In"]
hdr_row = None
rows = {}
for r in range(1, w.max_row + 1):
    k = w.cell(r, 1).value
    if k in ("act26_msft", "cap_msft", "gcap_msft", "gnai_msft", "gcapai_msft", "al_msft_self", "chip_msft", "obs_msft", "ai_sh", "k_chip2it", "fye_msft", "lag_on", "inv_on"):
        rows[k] = dict(label=w.cell(r, 2).value, unit=w.cell(r, 4).value, tag=w.cell(r, 5).value, src=w.cell(r, 6).value, note=w.cell(r, 7).value,
                       values=[w.cell(r, c).value for c in range(8, 18)], row=r)
out = dict(source=dict(file=os.path.basename(XLSX), version="v0.3.2", sha256=s["source"]["sha256"], snapshot="wattonomics/data/wattonomics_snapshot_v0.3.2.json（main，2026-11-14）",
                       note="Microsoft C0 對帳用的最小摘錄（只讀 WIF_ 名稱值與 In 列；不進模型計算）；In 列 values 欄＝In 頁 H–Q 欄（資本支出類：I＝FY2024、J＝FY2025、K＝FY2026、L＝FY2027…）"),
           names=names, rows=rows)
json.dump(out, open(os.path.join(a.out_dir, "wattonomics_c0_v0.3.2.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
t = json.load(open(a.trace, encoding="utf-8"))
CAMPUS = ("園區列", "園區輸入列", "公司公告園區")
keep, n_campus = [], 0
for x in t["rows"]:
    eff = "生效" in x["roots"].values()
    campus = any(x["flag"].startswith(c) for c in CAMPUS) if x["flag"] else False
    if campus:
        n_campus += 1
        continue
    if eff or x["flag"]:
        keep.append(dict(depth=x["depth"], cell=x["cell"], formula=x["formula"], value=x["value"], label=x["label"], flag=x["flag"],
                         status="生效" if eff else x["status"], roots=[k.split("_")[-1] if "AddGW" not in k else "AddGW" for k, v in x["roots"].items() if v == "生效"], leaf=x["leaf"]))
tr = dict(source=t["source"], player=t["player"], roots=t["roots"], rows=keep,
          act26_readers=t["act26_readers"], flagged=len(t["flagged"]), n_total=len(t["rows"]), n_campus_flag=n_campus, n_effective=sum(1 for x in keep if x["status"] == "生效"))
json.dump(tr, open(os.path.join(a.out_dir, "wattonomics_trace_v0.3.2.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(len(names), len(rows), len(keep), n_campus)
