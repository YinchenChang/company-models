#!/usr/bin/env python3
"""Oracle v0.5.0 反向 DCF：一次只動一個驅動（其餘維持基準），以整本重算（build_oracle.py ＋ LibreOffice）
二分法反解 (a) 加權目標價＝現價、(b) 加權目標價＝現價 × 0.75（評等由放空翻為中立）。結果寫入 oracle/data/reverse_v05.json，
建置時寫進 RevDCF 頁。搜尋範圍是各驅動的物理或常識界線（比 In 頁敏感度區間寬）；界線內無解就照實記錄，不外推。
用法：python3 oracle/scripts/reverse_v05.py [--jobs 6]
"""
import json, os, sys, tempfile, shutil, datetime
from concurrent.futures import ThreadPoolExecutor
import openpyxl
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sens_v04 import run, HERE   # noqa: E402

TOL = 0.02
# id, 標籤, In 鍵（顯示用）, 單位, 型態（abs＝直接設值；scale＝多鍵同乘一個倍數）, 鍵清單, 搜尋下界, 上界, 格式, 說明
DRV = [
    ("p_long", "長約價格（每 MW-IT 年收入的價格證據）", "p_long", "US$M／MW·年", "abs", ["p_long"], 6.0, 25.0, "0.00", "主導變數：OCI AI 營收＝計費 MW × 持有成本 × k，k 由長約價格決定"),
    ("capex", "每 MW IT 資本支出（四世代同乘倍數；1＝Tokenomics 基準）", "cx_*", "倍數", "scale", ["cx_h100", "cx_gb200", "cx_gb300", "cx_vr200"], 0.3, 1.5, "0.00", "倍數 < 1＝設備比 Tokenomics 基準便宜"),
    ("opex", "每 MW 年營運費用（四世代同乘倍數）", "ox_*", "倍數", "scale", ["ox_h100", "ox_gb200", "ox_gb300", "ox_vr200"], 0.0, 1.5, "0.00", "下界 0＝營運費用全免"),
    ("u_bill", "計費比例（計費 MW ÷ 在役 MW）", "u_bill", "x", "abs", ["u_bill"], 0.5, 1.0, "0.00", "上界 1.0＝全部在役產能都計費"),
    ("prepay", "客戶出資覆蓋比（預付 ÷ AI 資本支出）", "prepay_cov", "x", "abs", ["prepay_cov"], 0.0, 1.0, "0.00", "上界 1.0＝AI 資本支出全由客戶預付"),
    ("m_adj", "「OCI AI 以外」利潤率調整", "m_adj", "百分點（小數）", "abs", ["m_adj"], -0.2, 0.3, "0.000", "基準 0；+0.10＝雲端與軟體其餘業務利潤率多 10 個百分點"),
    ("wacc", "WACC 調整", "wacc_adj", "百分點（小數）", "abs", ["wacc_adj"], -0.06, 0.03, "0.0000", "基準 0；−0.02＝WACC 少 2 個百分點"),
    ("mult", "分部倍數調整係數", "mult_adj", "x", "abs", ["mult_adj"], 0.5, 4.0, "0.00", "基準 1；分部加總中各業務的 EV／EBITDA 倍數同乘"),
]


def main():
    jobs = int(sys.argv[sys.argv.index("--jobs") + 1]) if "--jobs" in sys.argv else 6
    tmp = tempfile.mkdtemp(prefix="orcl_rev5_")
    try:
        base = run({}, "base", tmp)
        ws = openpyxl.load_workbook(os.path.join(tmp, "base", "out", "raw.xlsx"), data_only=True)["In"]
        IN = {ws.cell(r, 1).value: (ws.cell(r, 4).value, ws.cell(r, 5).value, ws.cell(r, 6).value) for r in range(4, ws.max_row + 1) if ws.cell(r, 1).value}
        px = IN["px"][1]

        def ovr(d, x):
            return {k: x for k in d[5]} if d[4] == "abs" else {k: IN[k][1] * x for k in d[5]}

        def solve(d, target, tag):
            f = lambda x, i: run(ovr(d, x), f"{tag}_{i}", tmp)["tp"] - target
            a, b = d[6], d[7]
            fa, fb = f(a, "a"), f(b, "b")
            if fa * fb > 0:
                return None, {"tp_lo": fa + target, "tp_hi": fb + target}
            for i in range(14):
                m = (a + b) / 2; fm = f(m, i)
                if abs(fm) < TOL:
                    return m, {}
                if fa * fm <= 0: b, fb = m, fm
                else: a, fa = m, fm
            return (a + b) / 2, {}

        tasks = [(d, t, f"{d[0]}_{n}") for d in DRV for n, t in (("eq", px), ("neu", px * 0.75))]
        with ThreadPoolExecutor(jobs) as ex:
            res = list(ex.map(lambda t: solve(*t), tasks))
        out = {"date": datetime.date.today().isoformat(), "px": px, "base_tp": base["tp"], "tol": TOL, "drivers": []}
        for j, d in enumerate(DRV):
            (eq, eqx), (neu, _) = res[2 * j], res[2 * j + 1]
            if d[4] == "abs": lo, b0, hi = IN[d[5][0]]
            else: b0 = 1.0; lo, hi = [None if IN[d[5][0]][i] in (None, "") else IN[d[5][0]][i] / IN[d[5][0]][1] for i in (0, 2)]
            lo = lo if isinstance(lo, (int, float)) else None; hi = hi if isinstance(hi, (int, float)) else None
            if eq is None:
                inr = f"搜尋範圍 {d[6]}–{d[7]} 內無解（兩端目標價 {eqx['tp_lo']:.1f}／{eqx['tp_hi']:.1f}）"
            elif lo is None or hi is None:
                inr = "In 頁未設敏感度區間"
            else:
                inr = "區間內" if min(lo, hi) <= eq <= max(lo, hi) else "區間外"
            out["drivers"].append({"id": d[0], "label": d[1], "key": d[2], "unit": d[3], "base": b0, "eq_px": None if eq is None else round(eq, 6),
                                   "neutral": None if neu is None else round(neu, 6), "lo": lo, "hi": hi, "inrange": inr, "fmt": d[8],
                                   "note": d[9] + f"；搜尋範圍 {d[6]}–{d[7]}"})
        json.dump(out, open(os.path.join(HERE, "data", "reverse_v05.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        for x in out["drivers"]: print(x["id"], x["base"], x["eq_px"], x["neutral"], x["inrange"])
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
