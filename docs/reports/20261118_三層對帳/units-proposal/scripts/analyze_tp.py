import json, sys, statistics as st
TH = [0.5, 0.25, -0.25, -0.5]
def rec(u):
    for t, l in zip(TH, ["強力買進", "買進", "中立", "放空"]):
        if u > t: return l
    return "強力放空"
def run(a, b, label):
    A = json.load(open(a)); B = json.load(open(b)); rows = []
    for k in A:
        x, y = A[k], B.get(k, {})
        if "tp" not in x or "tp" not in y: continue
        px = x["tp"] / (1 + x["up"]); u1 = y["tp"] / px - 1
        rows.append(dict(k=k, n=x["name"], tp0=x["tp"], tp1=y["tp"], ch=y["tp"]/x["tp"]-1, u0=x["up"], u1=u1, r0=x["rec"], r1=y["rec"],
                         dist=min(abs(u1-t) for t in TH), near=min(TH, key=lambda t: abs(u1-t))))
    ch = [r["ch"] for r in rows]
    print(f"== {label}: n={len(rows)} 變動非零={sum(abs(c)>1e-9 for c in ch)} min={min(ch):+.2%} max={max(ch):+.2%} 中位={st.median(ch):+.2%} 平均={st.mean(ch):+.2%} ≤−5%={sum(c<=-0.05 for c in ch)} ≥+5%={sum(c>=0.05 for c in ch)}")
    print(" 建議改變：")
    for r in sorted([r for r in rows if r["r0"] != r["r1"]], key=lambda r: r["dist"]):
        # 越過的門檻
        crossed = [t for t in TH if (r["u0"] - t) * (r["u1"] - t) < 0]
        tc = crossed[0] if crossed else None
        print(f"  {r['n']}({r['k']}) {r['r0']}→{r['r1']} 上檔 {r['u0']:+.2%}→{r['u1']:+.2%} 越過 {tc:+.0%} 超過門檻 {abs(r['u1']-tc)*100:.2f} 個百分點；目標價 {r['ch']:+.2%}")
    print(" 接近門檻未翻轉（≤1 個百分點）：")
    for r in sorted([r for r in rows if r["r0"] == r["r1"] and r["dist"] <= 0.01], key=lambda r: r["dist"]):
        print(f"  {r['n']}({r['k']}) {r['r1']} 上檔 {r['u1']:+.2%}，離 {r['near']:+.0%} {r['dist']*100:.2f} 個百分點")
    print(" 變動最大 10 家（絕對值）：")
    for r in sorted(rows, key=lambda r: -abs(r["ch"]))[:10]:
        print(f"  {r['n']}({r['k']}) {r['ch']:+.2%} 上檔 {r['u0']:+.2%}→{r['u1']:+.2%} {r['r0']}→{r['r1']}")
    return rows
for a, b, l in zip(sys.argv[1::3], sys.argv[2::3], sys.argv[3::3]):
    run(a, b, l)
