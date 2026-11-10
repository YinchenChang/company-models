#!/usr/bin/env python3
"""Oracle v0.5.0 資金缺口壓力表（預先算好的組合）：情境 B 的建設延誤月數 × 客戶出資覆蓋比，整本重算，
記錄情境 B 的融資前累積現金谷底、瀑布另需外部融資、調整後槓桿峰值與加權目標價。寫入 oracle/data/stress_v05.json，建置時寫進 Fund 頁。
用法：python3 oracle/scripts/stress_v05.py [--jobs 4]
"""
import json, os, sys, tempfile, shutil, datetime, subprocess
from concurrent.futures import ThreadPoolExecutor
import openpyxl
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sens_v04 import run, HERE   # noqa: E402

DELAY = [0, 3, 6, 9, 12]          # delay_B（月）；基準 3
COV = [0.0, 0.243, 0.37]          # prepay_cov；基準 0.243、In 頁敏感度高 0.37、0＝客戶完全不出資


def one(ovr, tag, tmp):
    res = run(ovr, tag, tmp)
    d = os.path.join(tmp, tag)
    FU = json.load(open(os.path.join(d, "rowmap.json"), encoding="utf-8"))["FUND"]["Fund"]; R = FU["SUM"]; K = FU["R"]["B"]
    ws = openpyxl.load_workbook(os.path.join(d, "out", "raw.xlsx"), data_only=True)["Fund"]
    eqhy = sum(ws[f"{c}{K[k]}"].value or 0 for k in ("eq", "hy") for c in "DEFGH")
    return {"tp": res["tp"], "trough": ws[f"D{R['trough']}"].value, "wf": ws[f"D{R['wf']}"].value, "eqhy": eqhy, "lev": ws[f"D{R['lev']}"].value}


def main():
    jobs = int(sys.argv[sys.argv.index("--jobs") + 1]) if "--jobs" in sys.argv else 4
    tmp = tempfile.mkdtemp(prefix="orcl_st5_")
    try:
        tasks = [("base", {})] + [(f"d{d}_c{i}", {"delay_B": d, "prepay_cov": c}) for d in DELAY for i, c in enumerate(COV)]
        with ThreadPoolExecutor(jobs) as ex:
            res = dict(zip([t for t, _ in tasks], ex.map(lambda t: one(t[1], t[0], tmp), tasks)))
        out = {"date": datetime.date.today().isoformat(), "delay": DELAY, "cov": COV, "base": res["base"],
               "grid": [[res[f"d{d}_c{i}"] for i in range(len(COV))] for d in DELAY]}
        json.dump(out, open(os.path.join(HERE, "data", "stress_v05.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        for d, row in zip(DELAY, out["grid"]): print(d, [(round(x["wf"], 1), round(x["trough"], 1), round(x["lev"], 2), round(x["tp"], 1)) for x in row])
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
