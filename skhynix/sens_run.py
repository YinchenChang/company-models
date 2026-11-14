#!/usr/bin/env python3
"""SK 海力士 v0.7.6：加權目標價敏感度的整本重建快照（Sens 頁「加權目標價相對基準」後段；個股 HTML 版面規格 v1.3 H 節）。

用法：python3 skhynix/sens_run.py   → 寫 skhynix/sens_snapshot.json
做法：建置一份不含 v13 後處理與 Coord 頁的底稿（V13_SKIP=1、COORD_RUN=1），LibreOffice 重算取基準加權目標價；
每個案例把 In 頁對應輸入的基準欄（E）改成敏感度低（D）／高（F），其餘不變，存檔、LibreOffice 重算、讀 Val 頁加權目標價。
LibreOffice 私有設定檔用環境變數 LO_PROFILE（預設 file:///tmp/lo_skhynix）。
"""
import glob
import json
import os
import shutil
import subprocess
import sys

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
LO = os.environ.get("LO_PROFILE", "file:///tmp/lo_skhynix")
WORK = os.environ.get("SENS_WORK", "/tmp/sk_sens")

CASES = [  # (白話標籤, [In 鍵])
    ("一般 DRAM 位元成長（2027–2029 三年一起）", ["cb_1", "cb_2", "cb_3"]),
    ("NAND 位元成長（2027–2029 三年一起）", ["nb_1", "nb_2", "nb_3"]),
    ("HBM 營收市占（2026–2029 四年一起）", ["sh26", "sh27", "sh28", "sh29"]),
    ("HBM 2027 年單價變動", ["hbm_p27"]),
    ("HBM 2028 年單價變動（三情境一起）", ["hbm_p28_A", "hbm_p28_B", "hbm_p28_C"]),
    ("HBM 2029 年單價變動（三情境一起）", ["hbm_p29_A", "hbm_p29_B", "hbm_p29_C"]),
    ("HBM 位元供給成長（2027–2029 三年一起）", ["hbm_sup27", "hbm_sup28", "hbm_sup29"]),
    ("HBM 2026 年混合單價", ["hbm_p26"]),
    ("HBM 2026 年位元供給", ["hbm_sup26"]),
    ("情境 B 基準顆數係數 K1（2028、2029 一起）", ["um28_B", "um29_B"]),
    ("第四季一般 DRAM 價格季增", ["q4_dp"]),
    ("第四季一般 DRAM 位元季增", ["q4_db"]),
    ("第四季 NAND 價格季增", ["q4_np"]),
    ("2027–2028 一般 DRAM 合約價路徑（保守情境）", [f"cqC_{i}" for i in range(8, 16)]),
    ("2029 一般 DRAM 合約價季增（三情境一起）", [f"cq{s}_{i}" for s in "ABC" for i in range(16, 20)]),
    ("2029 NAND 合約價季增（三情境一起）", [f"nq{s}_{i}" for s in "ABC" for i in range(16, 20)]),
    ("單位成本年降", ["costdn_y"]),
    ("單位成本季降", ["costdn_q"]),
    ("營業費用變動率", ["opex_var"]),
    ("有效稅率", ["tax"]),
    ("資本支出（2026–2029 四年一起）", ["cap_0", "cap_1", "cap_2", "cap_3"]),
    ("新增資本支出轉折舊的年數", ["da_life"]),
    ("股東回饋占自由現金流比例", ["payout"]),
    ("美元兌韓元匯率", ["fx"]),
    ("基準年第四季資本支出", ["capex_rem"]),
    ("下半年每季業外淨收益", ["nonop_q"]),
]


def run(cmd, env=None, **kw):
    e = dict(os.environ, PYTHONDONTWRITEBYTECODE="1")
    e.update(env or {})
    subprocess.run(cmd, check=True, env=e, cwd=ROOT, stdout=subprocess.DEVNULL, **kw)


def recalc(files, outdir):
    os.makedirs(outdir, exist_ok=True)
    for i in range(0, len(files), 12):
        run(["soffice", f"-env:UserInstallation={LO}", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", outdir] + files[i:i + 12],
            stderr=subprocess.DEVNULL)


def weighted_tp(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    V = wb["Val"]
    for r in range(1, V.max_row + 1):
        if str(V.cell(r, 1).value or "").startswith("目標價（DCF 與 Comparable 加權）"):
            return V.cell(r, 2).value
    raise SystemExit("找不到加權目標價列")


def main():
    shutil.rmtree(WORK, ignore_errors=True)
    os.makedirs(os.path.join(WORK, "raw"))
    base_raw = os.path.join(WORK, "base.xlsx")
    run([sys.executable, "tools/memory_model/build.py", "skhynix/company.py", base_raw], env={"V13_SKIP": "1", "COORD_RUN": "1"})
    recalc([base_raw], os.path.join(WORK, "out"))
    base_val = openpyxl.load_workbook(os.path.join(WORK, "out", "base.xlsx"), data_only=True)["In"]
    base_tp = weighted_tp(os.path.join(WORK, "out", "base.xlsx"))
    rows = {str(base_val.cell(r, 1).value): r for r in range(4, base_val.max_row + 1) if base_val.cell(r, 1).value}
    wbf = openpyxl.load_workbook(base_raw)
    files, meta = [], []
    for ci, (label, keys) in enumerate(CASES):
        for side, col in (("lo", 4), ("hi", 6)):
            vals = {k: base_val.cell(rows[k], col).value for k in keys}
            if all(v is None for v in vals.values()):
                continue
            wb = openpyxl.load_workbook(base_raw)
            ws = wb["In"]
            for k, v in vals.items():
                if v is not None:
                    ws.cell(rows[k], 5).value = v
            f = os.path.join(WORK, "raw", f"c{ci:02d}_{side}.xlsx")
            wb.save(f)
            files.append(f)
            meta.append((ci, side))
    recalc(files, os.path.join(WORK, "out"))
    res = {}
    for (ci, side), f in zip(meta, files):
        res[(ci, side)] = weighted_tp(os.path.join(WORK, "out", os.path.basename(f)))
    cases = []
    for ci, (label, keys) in enumerate(CASES):
        lo, hi = res.get((ci, "lo")), res.get((ci, "hi"))
        cases.append({"label": label, "keys": "、".join(keys), "lo": lo, "hi": hi})
    cases.sort(key=lambda c: -max(abs((c[s] or base_tp) / base_tp - 1) for s in ("lo", "hi")))
    snap = {"snapshot": "Siliconomics v0.19（現行基準）", "base_tp": base_tp, "cases": cases}
    json.dump(snap, open(os.path.join(HERE, "sens_snapshot.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    for c in cases:
        print(f"{c['label']:<40} lo {((c['lo'] or base_tp) / base_tp - 1):+.1%}  hi {((c['hi'] or base_tp) / base_tp - 1):+.1%}")


if __name__ == "__main__":
    main()
