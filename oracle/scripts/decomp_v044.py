#!/usr/bin/env python3
"""Oracle v0.4.4 增資拆解：v0.4.3 → v0.4.4 的加權目標價，逐步打開增資各部分，以完整模型重算（build_oracle.py ＋ LibreOffice headless）。
結果寫入 oracle/data/decomp_v044.json（報告第 3 節的拆解表）。soffice 一次一件（依序執行）。

步驟（每步在前一步之上）：
  S0  v0.4.3 重現：現價 126.06、增資全部設 0（應等於 v0.4.3 的 73.2244）
  S1  現價 126.06 → 126.19（瀑布照算）；之後 S2–S4 把瀑布三列（投資級債、新股、高息債）凍結在 S1 的值
  S2  股數稀釋：新股 0.0862 十億股（評價股數、FY2027 起股數、新股股利、WACC 股權權重）
  S3  現金增加：新股淨額 9.9＋票據淨額 10.0 入帳（評價淨借款、FY2027 現金、利息收入）
  S4  負債與利息：票據面額 10.0 進借款、票息 7.77% 進損益與現金流、WACC 負債權重
  S5  瀑布後續變化：解除凍結（＝v0.4.4 成品）
另跑：EQ_only（只有新股、瀑布凍結；票據效果＝S4 − EQ_only）；MULT_lo／hi（mult_oci 豁免驗證）；判斷題替代案（不進成品）ALT_kd＝新增投資級債利率與 kd 改用新票據加權票息 7.77%（ALT_kd_v043＝同設定套在 v0.4.3）。

用法：python3 oracle/scripts/decomp_v044.py
"""
import json, os, subprocess, sys, tempfile, shutil, datetime
import openpyxl

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BUILD = os.path.join(HERE, "build_oracle.py")
ZERO = {"ex_sh": 0, "ex_eq_net": 0, "ex_nt_net": 0, "ex_nt_par": 0}


def run(ovr, tag, tmp, freeze=None):
    d = os.path.join(tmp, tag); os.makedirs(d, exist_ok=True)
    raw = os.path.join(d, "raw.xlsx"); rm = os.path.join(d, "rowmap.json")
    env = dict(os.environ, SENS_OVR=json.dumps(ovr), ORCL_ROWMAP=rm, ORCL_WF_FREEZE=json.dumps(freeze or {}))
    subprocess.run([sys.executable, BUILD, raw], check=True, env=env, cwd=os.path.dirname(HERE))
    prof = "file://" + os.path.join(d, "lo_profile")
    for _ in range(3):
        subprocess.run(["soffice", f"-env:UserInstallation={prof}", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", os.path.join(d, "out"), raw],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300)
        if os.path.exists(os.path.join(d, "out", "raw.xlsx")): break
    R = json.load(open(rm, encoding="utf-8"))
    wb = openpyxl.load_workbook(os.path.join(d, "out", "raw.xlsx"), data_only=True)
    V, CF = wb["Val"], wb["CF"]
    CON, DC, SO = R["VAL"]["CON"], R["VAL"]["DC"], R["VAL"]["SO"]
    G = R["GATE"]
    out = {"tp": V[f"C{CON['tp']}"].value, "dcf": V[f"C{CON['dcf']}"].value, "comp": V[f"C{CON['comp']}"].value,
           "wacc": V["C9"].value, "nd0": V["C10"].value, "shfd": V["C11"].value,
           "dcf_s": [V.cell(DC["v12"], c).value for c in (3, 4, 5)], "sotp_s": [V.cell(SO["v12"], c).value for c in (3, 4, 5)],
           "nd28": [V.cell(SO["nd"], c).value for c in (3, 4, 5)], "sh28": [V.cell(SO["sh"], c).value for c in (3, 4, 5)],
           "gate": [CF.cell(G, c).value for c in (3, 4, 5)], "ig5": [CF.cell(G + 1, c).value for c in (3, 4, 5)],
           "eq5": [CF.cell(G + 2, c).value for c in (3, 4, 5)], "hy5": [CF.cell(G + 3, c).value for c in (3, 4, 5)],
           "levmax": [CF.cell(G + 4, c).value for c in (3, 4, 5)], "wf": {}}
    for s in "ABC":
        C = R["CR"][s]
        out["wf"][s] = {c: [CF[f"{c}{C[k]}"].value for k in ("ig", "eq", "hy")] for c in "DEFGHI"}
    assert all(isinstance(out[k], (int, float)) for k in ("tp", "dcf", "comp")), (tag, out)
    return out


def main():
    tmp = tempfile.mkdtemp(prefix="orcl_dec_")
    try:
        res = {}
        res["S0"] = run({**ZERO, "px": 126.06}, "S0", tmp)
        res["S1"] = run(dict(ZERO), "S1", tmp)
        frz = res["S1"]["wf"]
        res["S2"] = run({k: v for k, v in ZERO.items() if k != "ex_sh"}, "S2", tmp, frz)
        res["S3"] = run({"ex_nt_par": 0}, "S3", tmp, frz)
        res["S4"] = run({}, "S4", tmp, frz)
        res["S5"] = run({}, "S5", tmp)
        res["EQ_only"] = run({"ex_nt_par": 0, "ex_nt_net": 0}, "EQ_only", tmp, frz)   # 補充：只有新股（股數＋淨額），票據全 0，瀑布凍結
        res["MULT_lo"] = run({"mult_oci": 4.81}, "MULT_lo", tmp)                          # Sens 頁 mult_oci 豁免理由的重算值
        res["MULT_hi"] = run({"mult_oci": 7.88}, "MULT_hi", tmp)
        res["ALT_kd"] = run({"r_new": 0.077675, "kd": 0.077675}, "ALT_kd", tmp)
        res["ALT_kd_v043"] = run({**ZERO, "px": 126.06, "r_new": 0.077675, "kd": 0.077675}, "ALT_kd_v043", tmp)
        for v in res.values(): v.pop("wf", None)
        out = {"date": datetime.date.today().isoformat(), "steps": res}
        json.dump(out, open(os.path.join(HERE, "data", "decomp_v044.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        prev = None
        for k, v in res.items():
            print(f"{k}: tp {v['tp']:.4f} dcf {v['dcf']:.4f} comp {v['comp']:.4f} gateB {v['gate'][1]:.3f}" + (f"  Δtp {v['tp'] - prev:+.4f}" if prev is not None and k.startswith("S") else ""))
            if k.startswith("S"): prev = v["tp"]
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
