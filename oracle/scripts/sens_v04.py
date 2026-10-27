#!/usr/bin/env python3
"""Oracle v0.4 敏感度：一次動一個輸入（低／高），以完整模型重算（build_oracle.py ＋ LibreOffice headless），
結果寫入 oracle/data/sens_v04.json；之後重跑 build_oracle.py 時寫進 Sens 頁。

用法：python3 oracle/scripts/sens_v04.py [--jobs 4]
（需要 soffice；每個案例一個獨立的 LibreOffice 使用者設定檔，可平行）
"""
import json, os, subprocess, sys, tempfile, shutil, datetime
from concurrent.futures import ThreadPoolExecutor
import openpyxl

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BUILD = os.path.join(HERE, "build_oracle.py")
TK = json.load(open(os.path.join(HERE, "data", "tokenomics_snapshot_v19.0.json"), encoding="utf-8"))["items"]
GN = {"h100": "Hopper H100", "gb200": "GB200 NVL72", "gb300": "GB300 NVL72", "vr200": "VR200 NVL72"}
OX = {g: round(TK["IF_OpexGW"]["values"][n]["基準"], 4) for g, n in GN.items()}

# (標籤, 設定說明, 低覆寫, 高覆寫)
CASES = [
    ("定價倍數 k", "k ×0.9／×1.2（1.2 約當 Oracle 自身合約）", {"k_adj": 0.9}, {"k_adj": 1.2}),
    ("計費比例", "0.90／1.00", {"u_bill": 0.90}, {"u_bill": 1.00}),
    ("併網速度（三情境）", "×0.8／×1.2", {"pace_A": 3333 * 0.8, "pace_B": 2833 * 0.8, "pace_C": 1000 * 0.8}, {"pace_A": 3333 * 1.2, "pace_B": 2833 * 1.2, "pace_C": 1000 * 1.2}),
    ("合約容量上限（三情境）", "×0.9／×1.1", {"cap_A": 8778 * 0.9, "cap_B": 8778 * 0.9, "cap_C": 6284 * 0.9}, {"cap_A": 8778 * 1.1, "cap_B": 8778 * 1.1, "cap_C": 6284 * 1.1}),
    ("建設延誤（三情境）", "全部 0 月／各加 6 月", {"delay_A": 0, "delay_B": 0, "delay_C": 0}, {"delay_A": 6, "delay_B": 9, "delay_C": 12}),
    ("AI 折舊年限", "5／7 年", {"life_ai": 5}, {"life_ai": 7}),
    ("AI 每 MW 營運費用", "×0.8／×1.2", {f"ox_{g}": v * 0.8 for g, v in OX.items()}, {f"ox_{g}": v * 1.2 for g, v in OX.items()}),
    ("未起租租約期限", "15／19 年", {"lease_term": 15}, {"lease_term": 19}),
    ("OCI 其他（非 AI）FY2027 年增", "10%／30%", {"gna0": 0.10}, {"gna0": 0.30}),
    ("SaaS 長期年增", "5%／10%", {"gLT_saas": 0.05}, {"gLT_saas": 0.10}),
    ("Software support 長期年增", "−3%／+1%", {"gLT_sup": -0.03}, {"gLT_sup": 0.01}),
    ("「OCI AI 以外」利潤率 m_sw", "−3pt／+3pt", {"m_adj": -0.03}, {"m_adj": 0.03}),
    ("未分攤費用年增", "2%／6%", {"un_g": 0.02}, {"un_g": 0.06}),
    ("客戶出資覆蓋比", "22.2%／37%", {"prepay_cov": 0.222}, {"prepay_cov": 0.37}),
    ("Beta（WACC）", "1.5／2.0", {"beta": 1.5}, {"beta": 2.0}),
    ("永續成長率 g", "2%／4%", {"tp_g": 0.02}, {"tp_g": 0.04}),
    ("分部倍數（OCI 與其餘業務）", "×0.8／×1.2", {"mult_adj": 0.8}, {"mult_adj": 1.2}),
    ("DCF 權重", "30%／70%", {"w_dcf": 0.3}, {"w_dcf": 0.7}),
    ("非 AI 資本支出 ÷ 營收", "5%／12%", {"nonai_cx": 0.05}, {"nonai_cx": 0.12}),
    ("FY2027 資本支出（指引）", "90／95", {"capex_g27": 90}, {"capex_g27": 95}),
]


def run(ovr, tag, tmp):
    d = os.path.join(tmp, tag); os.makedirs(d, exist_ok=True)
    raw = os.path.join(d, "raw.xlsx"); rm = os.path.join(d, "rowmap.json")
    env = dict(os.environ, SENS_OVR=json.dumps(ovr), ORCL_ROWMAP=rm)
    subprocess.run([sys.executable, BUILD, raw], check=True, env=env, cwd=os.path.dirname(HERE))
    prof = "file://" + os.path.join(d, "lo_profile")
    subprocess.run(["soffice", f"-env:UserInstallation={prof}", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", os.path.join(d, "out"), raw],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300)
    R = json.load(open(rm, encoding="utf-8"))
    wb = openpyxl.load_workbook(os.path.join(d, "out", "raw.xlsx"), data_only=True)
    tp = wb["Val"][f"C{R['VAL']['CON']['tp']}"].value
    om = wb["Thesis"][f"G{R['TR']['B']['om']}"].value
    gate = wb["CF"][f"D{R['GATE']}"].value
    assert all(isinstance(x, (int, float)) for x in (tp, om, gate)), (tag, tp, om, gate)
    return {"tp": tp, "om29": om, "gate": gate}


def main():
    jobs = int(sys.argv[sys.argv.index("--jobs") + 1]) if "--jobs" in sys.argv else 4
    tmp = tempfile.mkdtemp(prefix="orcl_sens_")
    try:
        tasks = [("base", {})] + [(f"c{i}_{side}", ov) for i, (_, _, lo, hi) in enumerate(CASES) for side, ov in (("lo", lo), ("hi", hi))]
        with ThreadPoolExecutor(jobs) as ex:
            res = dict(zip([t for t, _ in tasks], ex.map(lambda t: run(t[1], t[0], tmp), tasks)))
        out = {"date": datetime.date.today().isoformat(), "base": res["base"], "cases": []}
        for i, (lab, setting, lo, hi) in enumerate(CASES):
            out["cases"].append({"label": lab, "setting": setting, "lo": res[f"c{i}_lo"], "hi": res[f"c{i}_hi"], "lo_ovr": lo, "hi_ovr": hi})
        json.dump(out, open(os.path.join(HERE, "data", "sens_v04.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        print("base tp", res["base"]["tp"])
        for c in out["cases"]:
            print(f"{c['label']}: {c['lo']['tp']:.2f} / {c['hi']['tp']:.2f}")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
