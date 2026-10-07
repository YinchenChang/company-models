# v4.5（5a-1）：反向 DCF 改由建置時的 Python 求解——以 LibreOffice（UNO）開啟建好的 Excel，改輸入、重算、二分法反解，
# 結果寫入 rv_snap.json（build_xlsx.py 讀此檔寫入「評價_反向DCF」與「摘要」）。取代 v4.3 的 scripts/rv_snap.js（JS 引擎）。
# 求解規則與 HTML 的 reverseDcf（segA）相同：同一組上下界、22 次對半、以同一方向判斷，所以兩者結果在數值上一致（核對見 verify.sh）。
#   R＝DCF 每股＝現價所需的每 MW 年收入倍數（0.5–4）；C＝所需的建置成本倍數（0.1–1.5）；Eb＝所需的穩態 EBITDA 率（0.3–0.99）；
#   Rt＝加權目標價＝現價所需的每 MW 年收入倍數；grid＝建置成本倍數 × 穩態 EBITDA 率 下，DCF＝現價所需的每 MW 年收入倍數（無解為 null）。
# 用法：python3 scripts/rv_solve.py 檔案.xlsx [輸出 json，預設 repo 根目錄 rv_snap.json]
#   結果與原檔相同時代碼 0；有變動時寫入並以代碼 3 結束（verify.sh 據此重建 Excel）。
import sys, os, json, math
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from uno_q import Workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
xlsx = sys.argv[1]; out = sys.argv[2] if len(sys.argv) > 2 else os.path.join(ROOT, 'rv_snap.json')
IN, VA = '輸入與假設', '評價_DCF與目標價'

with Workbook(xlsx) as wb:
    c_rev, c_cap, c_eb = wb.cell(IN, '每 MW 年收入倍數（整體）'), wb.cell(IN, '每 MW 建置成本倍數（整體）'), wb.cell(IN, '穩態 EBITDA 率（FY30）')
    c_px = wb.cell(IN, '現價（', prefix=True)
    o_dcf, o_tgt = wb.cell(VA, 'DCF 每股'), wb.cell(VA, '加權目標價')
    base_in = {k: wb.get(c) for k, c in (('rev', c_rev), ('cap', c_cap), ('eb', c_eb))}
    P = wb.get(c_px)

    def run(**chg):
        for k, c in (('rev', c_rev), ('cap', c_cap), ('eb', c_eb)):
            wb.set(c, chg.get(k, base_in[k]))
        return {'dcf': wb.get(o_dcf), 'tgt': wb.get(o_tgt)}

    def solve(f, lo, hi, inc, key='dcf'):  # 與 segA reverseDcf 的 solve 相同
        a, b = f(lo)[key] - P, f(hi)[key] - P
        if not (math.isfinite(a) and math.isfinite(b)) or a * b > 0: return None
        for _ in range(22):
            m = (lo + hi) / 2; c = f(m)[key] - P
            if (c > 0) == inc: hi = m
            else: lo = m
        return (lo + hi) / 2

    base = run()
    R = solve(lambda x: run(rev=x), .5, 4, True)
    C = solve(lambda x: run(cap=x), .1, 1.5, False)
    Eb = solve(lambda x: run(eb=x), .3, .99, True)
    Rt = solve(lambda x: run(rev=x), .5, 4, True, 'tgt')
    caps, ebs = [.7, .8, .9, 1, 1.1], [.59, .65, .7, .75]
    grid = [[solve(lambda x: run(cap=cp, eb=s, rev=x), .5, 4, True) for s in ebs] for cp in caps]
    run()  # 還原輸入（不存檔）
    rev30 = wb.get(wb.cell(IN, '每 MW 年收入', col='G')) * base_in['rev'] * 1e3
    util30 = wb.get(wb.cell(IN, '利用率', col='G'))
    cost30 = wb.get(wb.cell(IN, '每 MW 建置成本', col='G')) * base_in['cap']

res = {'price': P, 'baseDcf': base['dcf'], 'baseTgt': base['tgt'], 'R': R, 'C': C, 'Eb': Eb, 'Rt': Rt,
       'caps': caps, 'ebs': ebs, 'grid': grid, 'rev30': rev30, 'util30': util30, 'cost30': cost30, 'eb30': base_in['eb']}
num = lambda x: int(x) if isinstance(x, float) and x.is_integer() else x   # 與 JSON.stringify 相同：整數不帶 .0
res = {k: ([[num(y) for y in r] for r in v] if k == 'grid' else [num(y) for y in v] if isinstance(v, list) else num(v)) for k, v in res.items()}
txt = json.dumps(res, separators=(',', ':'), ensure_ascii=False)
old = open(out, encoding='utf-8').read() if os.path.exists(out) else None
print(f"反向 DCF（Excel 求解）：現價 {P}、R {R:.4f}、C {C:.4f}、Eb {Eb:.4f}、Rt {Rt:.4f}；矩陣 {sum(x is not None for r in grid for x in r)}/20 格有解")
if old == txt:
    print('rv_snap.json 無變動'); sys.exit(0)
open(out, 'w', encoding='utf-8').write(txt)
print(f'rv_snap.json 已更新（{out}）'); sys.exit(3)
