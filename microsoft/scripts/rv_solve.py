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
    c_rev, c_cap, c_eb = wb.cell(IN, '每 MW 年收入倍數（整體）'), wb.cell(IN, '每 MW 建置成本倍數（整體）'), wb.cell(IN, '穩態 EBITDA 率（', prefix=True)
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
    caps, ebs = [.7, .8, .9, 1, 1.1], [.35, .47, .59, .7]  # v0.1b：可觀察 neocloud 區間（與 segA reverseDcf 相同）
    grid = [[solve(lambda x: run(cap=cp, eb=s, rev=x), .5, 4, True) for s in ebs] for cp in caps]
    # MAG v0.1b：容量軸（情境選擇）× 價格軸（價格軸選擇）3 × 3 加權目標價（與 HTML grid33Q 相同；其餘輸入維持目前值）
    g33 = None
    try:
        c_sc, c_ax = wb.cell(IN, '情境選擇（', prefix=True), wb.cell(IN, '價格軸選擇（', prefix=True)
    except KeyError:
        c_sc = c_ax = None
    if c_ax:
        run(); sc0, ax0 = wb.get(c_sc), wb.get(c_ax)
        g33 = []
        for s_ in (1, 2, 3):
            wb.set(c_sc, s_); row = []
            for a_ in (1, 2, 3):
                wb.set(c_ax, a_); row.append(wb.get(o_tgt))
            g33.append(row)
        wb.set(c_sc, sc0); wb.set(c_ax, ax0)
    # MAG v0.1b r3（對照表 r1 C16、C23）：三情境讀法 2 目標價；對外比例 ±20pt 目標價（其餘輸入維持目前值；基準 MW 速度不重解，與 HTML 相同）
    r2s = xs = None
    try:
        o_t2 = wb.cell(VA, '讀法 2｜加權目標價'); c_xs = wb.cell(IN, '對外 AI MW 占 AI 總 MW 比例')
    except KeyError:
        o_t2 = c_xs = None
    if o_t2 and c_sc:
        run(); sc0 = wb.get(c_sc); r2s = []
        for s_ in (1, 2, 3):
            wb.set(c_sc, s_); r2s.append(wb.get(o_t2))
        wb.set(c_sc, sc0)
        x0 = wb.get(c_xs); xs = []
        for x_ in (x0 - 0.2, min(1, x0 + 0.2)):
            wb.set(c_xs, x_); xs.append(wb.get(o_tgt))
        wb.set(c_xs, x0)
    # v0.1 交付前修訂：評價口徑敏感度（與 HTML segB valSensQ 同一演算法）：WACC 以 ERP 調整（ERP 對 WACC 線性）、ERP、非 AI 改自身倍數、AI 雲端 15×、組合；評等翻轉點（30 次對半）
    vs = vflip = None
    try:
        c_erp, c_w = wb.cell(IN, 'CAPM：股權風險溢酬'), wb.cell(IN, 'WACC')
        c_lov, c_ai = wb.cell(IN, '非 AI 事業 EV/EBITDA 手動覆蓋（空白＝同業中位數）', need_value=False), wb.cell(IN, 'EV/EBITDA 倍數')
        o_code = wb.cell(VA, '目標價區間｜點位評等代碼（1＝買進、0＝中立、−1＝賣出）'); o_th = wb.cell(VA, '目標價區間｜賣出門檻價', prefix=True)
    except KeyError:
        c_erp = None
    own = (json.load(open(os.path.join(ROOT, 'company.json'), encoding='utf-8'))['valuation'].get('ownMultiple') or {}).get('value')
    if c_erp:
        run(); e0, ai0 = wb.get(c_erp), wb.get(c_ai)
        clr = lambda ref: ref[0].getCellRangeByName(ref[1]).setFormula('')
        def at(erp=None, ownm=None, ai=None):
            wb.set(c_erp, e0 if erp is None else erp)
            if ownm: wb.set(c_lov, ownm)
            else: clr(c_lov)
            wb.set(c_ai, ai0 if ai is None else ai)
            return wb.get(o_tgt), wb.get(o_code)
        wb.set(c_erp, e0); wa = wb.get(c_w)
        wb.set(c_erp, e0 + .01); wb_ = wb.get(c_w); wb.set(c_erp, e0)
        ew = lambda w: e0 + (w - wa) / (wb_ - wa) * .01
        specs = [dict(erp=ew(.09)), dict(erp=ew(.10)), dict(erp=ew(.11)), dict(erp=.04), dict(erp=.06)] + ([dict(ownm=own)] if own else []) + [dict(ai=15), dict(erp=ew(.09), ownm=own, ai=15)]
        vs = [list(at(**k)) for k in specs]
        th = wb.get(o_th); base_tp = at()[0]
        if base_tp < th:
            f = lambda w: at(erp=ew(w))[0] - th
            lo, hi = .05, wa
            if f(lo) >= 0:
                for _ in range(30):
                    m = (lo + hi) / 2
                    if f(m) >= 0: lo = m
                    else: hi = m
                vflip = [(lo + hi) / 2, ew((lo + hi) / 2)]
        at()
    run()  # 還原輸入（不存檔）
    rev30 = wb.get(wb.cell(IN, '每 MW 年收入', col='G')) * base_in['rev'] * 1e3
    util30 = wb.get(wb.cell(IN, '利用率', col='G'))
    cost30 = wb.get(wb.cell(IN, '每 MW 建置成本', col='G')) * base_in['cap']

res = {'price': P, 'baseDcf': base['dcf'], 'baseTgt': base['tgt'], 'R': R, 'C': C, 'Eb': Eb, 'Rt': Rt,
       'caps': caps, 'ebs': ebs, 'grid': grid, **({'g33': g33} if g33 else {}), **({'r2s': r2s, 'xs': xs} if r2s else {}), **({'vs': vs, 'vflip': vflip} if vs else {}), 'rev30': rev30, 'util30': util30, 'cost30': cost30, 'eb30': base_in['eb']}
num = lambda x: int(x) if isinstance(x, float) and x.is_integer() else x   # 與 JSON.stringify 相同：整數不帶 .0
res = {k: ([[num(y) for y in r] for r in v] if k in ('grid', 'g33', 'vs') else [num(y) for y in v] if isinstance(v, list) else num(v)) for k, v in res.items()}
txt = json.dumps(res, separators=(',', ':'), ensure_ascii=False)
old = open(out, encoding='utf-8').read() if os.path.exists(out) else None
_f = lambda x: "無解" if x is None else f"{x:.4f}"
print(f"反向 DCF（Excel 求解）：現價 {P}、R {_f(R)}、C {_f(C)}、Eb {_f(Eb)}、Rt {_f(Rt)}；矩陣 {sum(x is not None for r in grid for x in r)}/20 格有解")
if old == txt:
    print('rv_snap.json 無變動'); sys.exit(0)
open(out, 'w', encoding='utf-8').write(txt)
print(f'rv_snap.json 已更新（{out}）'); sys.exit(3)
