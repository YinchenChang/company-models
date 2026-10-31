# v0.3 對 2026Q3 的模型值（變動拆解 (a)／(b) 與報告「新季實績 vs v0.3 模型值」用）。
# v0.3 只有 2H26 一欄：營收線、分部營收與分部營業利益依 2025 年季節性拆成 Q3（Q3 25 ÷（Q3 25＋Q4 25）），其餘流量取 2H26 的一半；
# 9/30 現金＝6/30 現金＋2H 自由現金流 × 0.5 − 2H 策略投資＋2H 已發行債券＋2H 借款 × 0.5；債務＝6/30 面額＋已發行債券＋2H 借款 × 0.5。情境 B。
# v0.3 的列號由 git 上 v0.3 的 build_amazon.py（main 1eec35e9）重建取得；數值讀 v0.3 成品（已重算）。
# 用法：python3 amazon/q3cf_v03.py  → 寫 amazon/data/q3cf_v03_2026q3.json（facts：{事實 id: 值（USD M）}；另含 v0.3 隱含 Q4 26 營收與對照表）
import json, os, shutil, subprocess, sys, tempfile
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
V03 = os.path.join(HERE, "dist", "_待刪除", "20261027_Amazon個股模型_v0.3.xlsx")
if not os.path.exists(V03):
    V03 = os.path.join(HERE, "dist", "20261027_Amazon個股模型_v0.3.xlsx")
TD = tempfile.mkdtemp(prefix="v03_")
os.makedirs(os.path.join(TD, "amazon", "data"))
src = subprocess.run(["git", "-C", ROOT, "show", "1eec35e9:amazon/build_amazon.py"], check=True, capture_output=True, text=True).stdout
open(os.path.join(TD, "amazon", "build_amazon.py"), "w", encoding="utf-8").write(src)
for f in ("amazon_facts_20261008.json", "consensus_amzn_20261008.json"):
    shutil.copy(os.path.join(HERE, "data", f), os.path.join(TD, "amazon", "data", f))
shutil.copy(os.path.join(HERE, "data", "_待刪除", "tokenomics_snapshot_v21.4.json"), os.path.join(TD, "amazon", "data", "tokenomics_snapshot_v21.4.json"))
rmf = os.path.join(TD, "rm.json")
subprocess.run([sys.executable, os.path.join(TD, "amazon", "build_amazon.py"), os.path.join(TD, "raw.xlsx")], check=True, stdout=subprocess.DEVNULL, env=dict(os.environ, AMZN_ROWMAP=rmf))
REG = {tuple(k.split("|")): v for k, v in json.load(open(rmf, encoding="utf-8"))["REG"].items()}
wb = openpyxl.load_workbook(V03, data_only=True)


def g(blk, key, col="F"):
    sh, r = REG[(blk, key)]
    return wb[sh][f"{col}{r}"].value


OLD = json.load(open(os.path.join(HERE, "data", "amazon_facts_20261008.json"), encoding="utf-8"))
F0 = {f["id"]: f["value"] for gg in OLD["groups"].values() for f in gg}
NEW = {f["id"]: f["value"] for f in json.load(open(os.path.join(HERE, "data", "amazon_facts_2026q3.json"), encoding="utf-8"))["items"]}
sh_q3 = lambda base: F0[f"{base}.2025q3"] / (F0[f"{base}.2025q3"] + F0[f"{base}.2025q4"])
M = 1000.0
out = {}
LINES = (("os", "is.cat.online"), ("ps", "is.cat.physical"), ("sp", "is.cat.seller3p"), ("ad", "is.cat.ads"), ("sub", "is.cat.subs"), ("ot", "is.cat.other"))
rt = 0.0
for k, fid in LINES:
    v = g("BS", "rev_" + k) * sh_q3(fid) * M
    out[f"{fid}.2026q3"] = v; rt += v
na_sh = g("BS", "na") / g("BS", "rt")
out["is.NA.rev.2026q3"] = rt * na_sh
out["is.INTL.rev.2026q3"] = rt * (1 - na_sh)
out["is.AWS.rev.2026q3"] = g("BS", "aws") * sh_q3("is.AWS.rev") * M
out["is.TOT.rev.2026q3"] = rt + out["is.AWS.rev.2026q3"]
for seg, key in (("NA", "oi_na"), ("INTL", "oi_in"), ("AWS", "oi_aws")):
    out[f"is.{seg}.oi.2026q3"] = g("BP", key) * sh_q3(f"is.{seg}.oi") * M
out["is.TOT.oi.2026q3"] = sum(out[f"is.{s}.oi.2026q3"] for s in ("NA", "INTL", "AWS"))
intx = g("BP", "intx") / 2 * M; inti = g("BP", "inti") / 2 * M
out["is.intExp.2026q3"] = intx; out["is.intInc.2026q3"] = inti
pretax = out["is.TOT.oi.2026q3"] - intx + inti
taxr = g("BP", "taxx") / g("BP", "pretax")
out["is.pretax.2026q3"] = pretax; out["is.tax.2026q3"] = pretax * taxr; out["is.ni.2026q3"] = pretax * (1 - taxr)
out["sh.wDiluted.2026q3"] = g("BP", "sh") * M
out["sh.wDiluted.2026m9"] = (F0["sh.wDiluted.2026q1"] + F0["sh.wDiluted.2026q2"] + out["sh.wDiluted.2026q3"]) / 3
for fid, blk, key in (("cf.cfo", "BC", "cfo"), ("cf.sbc", "BC", "sbc")):
    out[f"{fid}.2026q3"] = g(blk, key) / 2 * M
cx = g("BC", "cx") / 2 * M
out["cf.capexGross.2026q3"] = cx; out["cf.capexNet.2026q3"] = cx
segda = g("BP", "da") / 2 * M
out["cf.da.2026q3"] = F0["cf.da.2026q2"] * segda / (F0["is.NA.da.2026q2"] + F0["is.INTL.da.2026q2"] + F0["is.AWS.da.2026q2"])
out["cf.finLeasePrin.2026m9"] = (g("BC", "lease", "E") + g("BC", "lease") / 2) * M
out["cf.acqInvest.2026m9"] = (g("BC", "inv", "E") + g("BC", "inv")) * M
cash_q2 = F0["bs.cashMs.2026q2"]
out["bs.cashMs.2026q3"] = cash_q2 + (0.5 * g("BC", "fcfd") - g("BC", "inv") + g("BC", "iss") + 0.5 * g("BC", "borrow")) * M
out["debt.face.20260930"] = F0["debt.face.2026q2"] + (g("BC", "iss") + 0.5 * g("BC", "borrow")) * M
out["sh.out.20260930"] = F0["sh.out.20260722"]
out["sh.awards.2026q3"] = F0["sh.awards.2026q2"]
for base in ("is.TOT.rev", "is.TOT.oi"):
    out[f"{base}.2026m9"] = F0[f"{base}.2026q1"] + F0[f"{base}.2026q2"] + out[f"{base}.2026q3"]
q4rev = g("BS", "tot") - out["is.TOT.rev.2026q3"] / M
CMP = [(k, out[k] / M, NEW[k] / M) for k in sorted(out) if k in NEW and isinstance(NEW[k], (int, float))]
res = {"_readme": "v0.3（2026-10-27，情境 B）對 2026Q3 的模型值；amazon/q3cf_v03.py 產生。facts 單位 USD M（與事實總帳同）；q4rev＝v0.3 隱含 Q4 26 合併營收（US$bn）",
       "facts": out, "q4rev": q4rev, "v03_2H26": {"tot": g("BS", "tot"), "aws": g("BS", "aws"), "cx": g("BC", "cx"), "cfo": g("BC", "cfo"), "pace_B": 3634.3},
       "compare": [{"id": k, "v03": a, "actual": b, "gap": b / a - 1 if a else None} for k, a, b in CMP]}
json.dump(res, open(os.path.join(HERE, "data", "q3cf_v03_2026q3.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
for k, a, b in CMP:
    print(f"{k:28s} v0.3 {a:9.3f}  實績 {b:9.3f}  {b / a - 1:+.1%}" if a else k)
print("v0.3 隱含 Q4 26 營收", round(q4rev, 3))
shutil.rmtree(TD, ignore_errors=True)
