# 支持現價所需的 k（v0.2.6；驗收報告 2026-11-09 A6「what must be true」）：整本 LibreOffice 重算反解，只作對照、不進基準。
#   flat：三情境 k 由 Q4 26 起維持同一值到 FY2032（k0_ovr＝x、conv_k 極大），解 x 使加權目標價＝現價。
#   conv：k0_ovr＝x，FY2030 情境值（k30_A／B／C）與收斂年數照基準，解 x 使加權目標價＝現價；搜尋上限內無解時記錄上限與該點目標價。
# 用法：python3 spacex/kreq_v02.py <工作資料夾> <輸出 json>（割線法；每次試算一次整本重算，約 5 秒）
import json, os, subprocess, sys, shutil, tempfile
import openpyxl
HERE = os.path.dirname(os.path.abspath(__file__))
WD, OUT = sys.argv[1], sys.argv[2]
os.makedirs(WD, exist_ok=True)
N = [0]


def tp_at(sets):
    N[0] += 1
    raw = os.path.join(WD, f"k{N[0]}.xlsx"); rmf = raw + ".rowmap.json"
    cmd = [sys.executable, os.path.join(HERE, "build_spacex.py"), raw] + sum((["--set", f"{k}={v}"] for k, v in sets), [])
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, env={**os.environ, "ROWMAP": rmf})
    out = os.path.join(WD, "out"); prof = tempfile.mkdtemp(prefix="lo_kr_", dir=WD)
    subprocess.run(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", out, raw],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env={**os.environ, "LC_ALL": "C.UTF-8"}, timeout=900)
    shutil.rmtree(prof, ignore_errors=True)
    rm = json.load(open(rmf, encoding="utf-8")); wb = openpyxl.load_workbook(os.path.join(out, os.path.basename(raw)), data_only=True)
    return wb["Val"].cell(rm["VR"]["tp"], 2).value, wb["Val"].cell(rm["VR"]["px"], 2).value


def solve(extra, a, b, hi_cap):
    fa, px = tp_at([("k0_ovr", a)] + extra); fa -= px
    fb_, _ = tp_at([("k0_ovr", b)] + extra); fb_ -= px
    while fb_ < 0 and b < hi_cap:           # 往上擴張區間
        a, fa = b, fb_; b = min(hi_cap, b * 1.6)
        fb_, _ = tp_at([("k0_ovr", b)] + extra); fb_ -= px
    if fb_ < 0:
        return {"solved": False, "k_cap": b, "tp_at_cap": fb_ + px, "px": px}
    for _ in range(25):
        c = b - fb_ * (b - a) / (fb_ - fa)
        fc, _ = tp_at([("k0_ovr", c)] + extra); fc -= px
        print(round(c, 4), round(fc, 3), flush=True)
        a, fa, b, fb_ = b, fb_, c, fc
        if abs(fc) < 0.005: break
    return {"solved": True, "k": b, "tp": fb_ + px, "px": px}


res = {"flat": solve([("conv_k", 1000000)], 1.8666, 4.0, 30.0),
       "conv": solve([], 4.8, 8.0, 30.0),
       "note": "flat＝三情境 k 由 Q4 26 起維持同一值（conv_k 極大）；conv＝只改 k0、FY2030 情境值與收斂年數照基準。整本 LibreOffice 重算反解；只作對照、不進基準"}
json.dump(res, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps(res, ensure_ascii=False))
