# MiniMax 收支模型 — 一頁摘要 HTML（呈現層；所有數字讀自已由 LibreOffice 重算的 Excel 與敏感度快照）
# 用法：python3 build_html.py 模型.xlsx out/sensitivity.json 輸出.html
import sys, json, html
from openpyxl import load_workbook

xl, sens_path, out = sys.argv[1:4]
wb = load_workbook(xl, data_only=True)
V = {}
for ws in wb.worksheets:
    if ws.title in ('Demand', 'Revenue', 'Compute', 'PnL', 'Funding', 'Valuation'):
        for r in range(6, ws.max_row + 1):
            k = ws.cell(r, 1).value
            if isinstance(k, str) and len(k) <= 4:
                V[k] = [ws.cell(r, c).value for c in range(4, 10)]
                V[k + '_label'] = ws.cell(r, 2).value
SENS = json.load(open(sens_path))
Y = [2025, 2026, 2027, 2028, 2029, 2030]


def g(k, y): return V[k][Y.index(y)]
def f0(x): return '—' if x is None else f'{x:,.0f}'
def f1(x): return '—' if x is None else f'{x:,.1f}'
def f2(x): return '—' if x is None else f'{x:,.2f}'
def pc(x): return '—' if x is None else f'{x*100:.0f}%'


base = SENS[0][2]
first_fund = g('F19', 2026)
first_fund_txt = '2030 年內不需要' if first_fund == 9999 else f'{first_fund:.0f} 年'


def line_chart(series, ymax, unit, ident):
    W, H, L, R, T, B = 640, 260, 48, 16, 16, 34
    xs = [L + i * (W - L - R) / (len(Y) - 1) for i in range(len(Y))]
    sy = lambda v: T + (H - T - B) * (1 - v / ymax)
    out = [f'<svg viewBox="0 0 {W} {H}" class="chart" role="img" aria-labelledby="{ident}-t">']
    step = ymax / 4
    for i in range(5):
        v = step * i; yy = sy(v)
        out.append(f'<line x1="{L}" x2="{W-R}" y1="{yy:.1f}" y2="{yy:.1f}" class="grid"/>'
                   f'<text x="{L-6}" y="{yy+4:.1f}" class="ax" text-anchor="end">{v:,.0f}</text>')
    for i, y in enumerate(Y):
        out.append(f'<text x="{xs[i]:.1f}" y="{H-12}" class="ax" text-anchor="middle">{y}{"A" if y == 2025 else "E"}</text>')
    for name, vals, var in series:
        pts = ' '.join(f'{xs[i]:.1f},{sy(v):.1f}' for i, v in enumerate(vals))
        out.append(f'<polyline points="{pts}" fill="none" stroke="var({var})" stroke-width="2" stroke-linejoin="round"/>')
        for i, v in enumerate(vals):
            out.append(f'<circle cx="{xs[i]:.1f}" cy="{sy(v):.1f}" r="4" fill="var({var})" stroke="var(--surface)" stroke-width="2" '
                       f'data-tip="{html.escape(name)}｜FY{Y[i]}：{v:,.2f} {unit}"/>')
        out.append(f'<text x="{xs[-1]+0:.1f}" y="{sy(vals[-1])-10:.1f}" class="dl" text-anchor="end">{name} {vals[-1]:,.1f}</text>')
    out.append('</svg>')
    return ''.join(out)


def bar_chart(vals, lo, hi, unit, minline):
    W, H, L, R, T, B = 640, 260, 56, 16, 16, 34
    n = len(vals); bw = (W - L - R) / n
    sy = lambda v: T + (H - T - B) * (hi - v) / (hi - lo)
    out = ['<svg viewBox="0 0 %d %d" class="chart" role="img">' % (W, H)]
    for v in range(int(lo), int(hi) + 1, 1000):
        yy = sy(v)
        out.append(f'<line x1="{L}" x2="{W-R}" y1="{yy:.1f}" y2="{yy:.1f}" class="{"zero" if v == 0 else "grid"}"/>'
                   f'<text x="{L-6}" y="{yy+4:.1f}" class="ax" text-anchor="end">{v:,}</text>')
    ym = sy(minline)
    out.append(f'<line x1="{L}" x2="{W-R}" y1="{ym:.1f}" y2="{ym:.1f}" class="minl"/>'
               f'<text x="{W-R}" y="{ym-5:.1f}" class="ax" text-anchor="end">最低現金 {minline:,.0f}</text>')
    for i, v in enumerate(vals):
        x = L + i * bw + bw * 0.25; w = bw * 0.5
        y0, y1 = sy(max(v, 0)), sy(min(v, 0))
        cls = 'var(--neg)' if v < minline else 'var(--s1)'
        out.append(f'<rect x="{x:.1f}" y="{y0:.1f}" width="{w:.1f}" height="{max(y1-y0,1):.1f}" rx="3" fill="{cls}" '
                   f'data-tip="FY{Y[i]} 期末現金（新融資前）：{v:,.0f} {unit}"/>')
        out.append(f'<text x="{x+w/2:.1f}" y="{(y0-6) if v>=0 else (y1+14):.1f}" class="dl" text-anchor="middle">{v:,.0f}</text>')
        out.append(f'<text x="{x+w/2:.1f}" y="{H-12}" class="ax" text-anchor="middle">{Y[i]}{"A" if Y[i]==2025 else "E"}</text>')
    out.append('</svg>')
    return ''.join(out)


rev_mw = [g('P21', y) for y in Y]; cost_mw = [g('P22', y) for y in Y]
c1 = line_chart([('每 MW 全成本', cost_mw, '--s2'), ('每 MW 營收', rev_mw, '--s1')], 20, '$M/MW/年', 'c1')
cash = [g('F09', y) for y in Y]
c2 = bar_chart(cash, -2000, 4000, '$M', 300)

rows_pl = [('R01', f1, '企業端營收（API）'), ('R02', f1, '消費端營收'), ('R09', f1, '總營收'), ('P07', pc, '毛利率'),
           ('P10', f1, '訓練算力'), ('P12', f1, '人事（含 SBC）'), ('P14', f1, '其他營業費用'),
           ('P18', f1, '經調整淨利'), ('C21', f1, '總 MW（年均）'), ('C22', pc, '訓練占 MW'),
           ('P21', f2, '每 MW 營收 $M'), ('P22', f2, '每 MW 全成本 $M'), ('P23', f2, '覆蓋率（x）'),
           ('F09', f0, '期末現金（新融資前）'), ('F11', f0, '累計需新融資')]
tbl = ['<table class="t"><thead><tr><th>$M（除註明）</th>' + ''.join(f'<th>FY{y}{"A" if y==2025 else "E"}</th>' for y in Y) + '</tr></thead><tbody>']
for k, fm, lab in rows_pl:
    strong = ' class="key"' if k in ('R09', 'P23', 'F11') else ''
    tbl.append(f'<tr{strong}><td>{lab}</td>' + ''.join(f'<td>{fm(v)}</td>' for v in V[k]) + '</tr>')
tbl.append('</tbody></table>')

sens_rows = ['<table class="t s"><thead><tr><th>情境（只改一項）</th><th>2030 營收</th><th>2030 覆蓋率</th><th>累計需新融資</th><th>首次融資年</th><th>折現價值÷市值</th></tr></thead><tbody>']
for name, _, v in SENS:
    fy = '不需要' if v[6] == 9999 else f'{v[6]:.0f}'
    hl = ' class="key"' if name == '基準' else ''
    sens_rows.append(f'<tr{hl}><td>{html.escape(name)}</td><td>{f0(v[1])}</td><td>{v[3]:.2f}x</td><td>{f0(v[5])}</td><td>{fy}</td><td>{v[7]:.2f}x</td></tr>')
sens_rows.append('</tbody></table>')

eta25, eta26 = g('C11', 2025), g('C11', 2026)
page = f'''<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>MiniMax 收支模型 v0.1</title>
<style>
:root{{--surface:#fcfcfb;--bg:#f4f4f2;--ink:#0b0b0b;--ink2:#52514e;--ink3:#8a8984;--line:#e4e3df;--s1:#2a78d6;--s2:#eb6834;--neg:#e34948;--key:#fff6d9}}
@media (prefers-color-scheme: dark){{:root:not([data-theme="light"]){{--surface:#1a1a19;--bg:#121211;--ink:#fff;--ink2:#c3c2b7;--ink3:#8f8e86;--line:#33322f;--s1:#3987e5;--s2:#d95926;--neg:#e66767;--key:#3a3320}}}}
:root[data-theme="dark"]{{--surface:#1a1a19;--bg:#121211;--ink:#fff;--ink2:#c3c2b7;--ink3:#8f8e86;--line:#33322f;--s1:#3987e5;--s2:#d95926;--neg:#e66767;--key:#3a3320}}
*{{box-sizing:border-box}}body{{margin:0;background:var(--bg);color:var(--ink);font:15px/1.6 "Noto Sans TC","PingFang TC","Microsoft JhengHei",system-ui,sans-serif}}
main{{max-width:1080px;margin:0 auto;padding:24px 16px 64px}}
h1{{font-size:26px;margin:0 0 4px}}h2{{font-size:18px;margin:32px 0 8px}}.sub{{color:var(--ink2);margin:0 0 16px}}
.card{{background:var(--surface);border:1px solid var(--line);border-radius:10px;padding:16px 18px;margin:12px 0}}
.ans{{font-size:17px}}.ans b{{color:var(--ink)}}
.tiles{{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px}}
.tile{{background:var(--surface);border:1px solid var(--line);border-radius:10px;padding:12px 14px}}
.tile .v{{font-size:26px;font-weight:700}}.tile .l{{color:var(--ink2);font-size:13px}}
.chart{{width:100%;height:auto;display:block}}.grid{{stroke:var(--line);stroke-width:1}}.zero{{stroke:var(--ink3);stroke-width:1}}
.minl{{stroke:var(--ink3);stroke-dasharray:4 4}}.ax{{fill:var(--ink3);font-size:11px}}.dl{{fill:var(--ink2);font-size:11px}}
.legend{{display:flex;gap:16px;font-size:13px;color:var(--ink2);margin:4px 0}}.sw{{display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:6px;vertical-align:middle}}
.scroll{{overflow-x:auto}}table.t{{border-collapse:collapse;width:100%;font-size:13px;font-variant-numeric:tabular-nums}}
.t th,.t td{{padding:5px 8px;border-bottom:1px solid var(--line);text-align:right;white-space:nowrap}}.t th:first-child,.t td:first-child{{text-align:left}}
.t thead th{{color:var(--ink2);font-weight:600}}.t tr.key td{{background:var(--key);font-weight:600}}
ul{{margin:6px 0;padding-left:20px}}li{{margin:3px 0}}.muted{{color:var(--ink2);font-size:13px}}
#tip{{position:fixed;pointer-events:none;background:var(--ink);color:var(--surface);font-size:12px;padding:4px 8px;border-radius:6px;opacity:0;transition:opacity .1s}}
.two{{display:grid;grid-template-columns:1fr 1fr;gap:12px}}@media (max-width:760px){{.two{{grid-template-columns:1fr}}}}
</style></head><body><main>
<h1>MiniMax（00100.HK）收支模型 v0.1</h1>
<p class="sub">上海稀宇科技｜比照 OpenAI 收支模型（MW 主軸、Tokenomics v5.24）｜FY2025 實際、FY2026 上半年實際＋下半年推估、FY2027–2030 推估｜2026-10-08｜單位 US$M</p>

<div class="card ans"><b>命題：</b>MiniMax 每 MW 算力的營收，能否覆蓋每 MW 全成本？若不能，缺口何時出現、由誰補？<br>
<b>答案（基準）：不能。</b>每 MW 營收由 FY2025 的 ${f2(g('P21',2025))}M 升到 FY2030 的 ${f2(g('P21',2030))}M，但每 MW 全成本仍約 ${f1(g('P22',2030))}M，覆蓋率只從 {f2(g('P23',2025))}x 升到 {f2(g('P23',2030))}x。
推論本身已經賺錢（API 毛利率 FY2030 {pc(g('P08',2030))}），虧損來自訓練：訓練占總 MW {pc(g('C22',2030))}，訓練＋人事＋其他費用是營收的 {f1(g('P29',2030))} 倍。
2026 年募得的約 ${f0(g('F18',2026)-g('F15',2025))}M（IPO、配售、可轉債）撐到 <b>{first_fund_txt}</b>，之後至 2030 年累計還需約 <b>${f0(g('F11',2030))}M</b>（相當目前市值 {pc(g('F14',2030))} 的新股）。</div>

<div class="tiles">
<div class="tile"><div class="v">{f2(g('P23',2030))}x</div><div class="l">FY2030 覆蓋率（每 MW 營收 ÷ 全成本）</div></div>
<div class="tile"><div class="v">{f0(g('R09',2026))}</div><div class="l">FY2026 營收（模型）；共識 {f0(g('R13',2026))}</div></div>
<div class="tile"><div class="v">{f0(g('C21',2026))} MW</div><div class="l">FY2026 年均算力（約 {f0(g('C25',2026)/1000)}K 張 Hopper 等值）</div></div>
<div class="tile"><div class="v">{first_fund_txt}</div><div class="l">首次需新融資</div></div>
<div class="tile"><div class="v">{f2(g('V13',2030))}x</div><div class="l">反向檢驗：模型價值 ÷ 目前市值</div></div>
</div>

<h2>每 MW 營收 vs 每 MW 全成本</h2>
<div class="card"><div class="legend"><span><i class="sw" style="background:var(--s1)"></i>每 MW 營收</span><span><i class="sw" style="background:var(--s2)"></i>每 MW 全成本（銷售成本＋營業費用，含 SBC）</span></div>{c1}
<p class="muted">每 MW 全成本中算力部分固定為合約價 $12M（雲端租用口徑），其餘為人事與其他費用。覆蓋率不受合約價影響：所有 MW 皆由支出 ÷ 合約價換算，合約價只改變 MW 規模。</p></div>

<h2>現金路徑（新融資前）</h2>
<div class="card">{c2}<p class="muted">2026 現金來源：期初 1,050、IPO 淨額 {f0(g('F02',2026))}、配售 {f0(g('F03',2026))}、可轉債 {f0(g('F04',2026))}；可轉債轉換價 HK$335 高於目前約 HK${f0(g('V01',2026))}，基準於 2027 年以現金贖回 {f0(-g('F08',2027))}。</p></div>

<h2>逐年數字</h2>
<div class="card scroll">{''.join(tbl)}</div>

<h2>敏感度：哪些假設會翻轉結論</h2>
<div class="card scroll">{''.join(sens_rows)}
<p class="muted">翻轉結論（覆蓋率 ≥1 或至 2030 不需新融資）只有三種：API 單價不再下跌、token 量年增遠高於基準、訓練 MW 2027 年起不再成長。合約價與層級組合對覆蓋率幾乎沒有影響（被 η 校準吸收）。</p></div>

<h2>主要驅動、校準與差距</h2>
<div class="card"><ul>
<li><b>API 價量：</b>2026 計費 token 由營收 ÷ 實收單價 $0.22/M（公司說法，未見一手）推得，約 {f0(g('D03',2026))}T，為 2025 的 {f1(g('D06',2026))} 倍（公司稱 7 月 token 為 1 月的 20 倍）。基準 2027 起量 +200%/+100%/+60%/+40%、價 −35%/−25%/−20%/−15%。</li>
<li><b>η（V2）：</b>2025 為 {f2(eta25)}、2026 上半年為 {f2(eta26)}；同一 MW 產出的計費 token 大幅下降，反映 Token Plan 低價 token（滿用約 $0.012/M）、長上下文與較大的 M3。下半年起按公司 M3.1 降本說法只採 1.3 倍改善。</li>
<li><b>訓練：</b>2026 由上半年研發費用扣人事推得約 ${f0(g('C17',2026))}M，全年約 ${f0(g('C18',2026))}M；2027 起 MW 年增 80%/50%/35%/25%。這是虧損的主因，也是最大的判斷項。</li>
<li><b>與市場：</b>FY2026 模型 {f0(g('R09',2026))} vs 共識 452；FY2028 模型 {f0(g('R09',2028))} vs Goldman 2,470；FY2030 模型 {f0(g('R09',2030))} vs Visible Alpha 5,800。差距來自單價：賣方隱含價格下跌遠小於本模型。</li>
<li><b>誰補缺口：</b>上市前創投累計約 $1.56B（阿里巴巴 Pre-B 領投 $654M）；IPO 基石 ADIA、阿里巴巴等；2026-07 配售與可轉債約 $2.0B。下一輪可能來源：A 股（科創板）上市（已聘中信證券，媒體）、再配售（7 月前例折價 9.9%）、阿里雲（股東兼供應商，採購上限 2026–28 年 $300/400/500M）的供應商融資、銀行借款（上半年由 35 增至 134）。</li>
</ul></div>

<h2>失效條件與追蹤</h2>
<div class="card"><ul>
<li>FY2026 全年營收若 &gt; $540M（2H 企業端年化 ≥ $750M），首次融資年延到 2029，反向檢驗升至約 0.9x。</li>
<li>中國 API 價格若續持平或上漲（DeepSeek 8 月已漲價），覆蓋率可在 2030 前超過 1：追蹤 MiniMax 牌價與 OpenRouter 實際成交價。</li>
<li>半年報毛利率：1H26 為 17.9%；若 2H26 未回升到 35% 以上，η 改善假設（I29）偏樂觀。</li>
<li>訓練支出：研發費用扣人事後的半年增速；若 2027 增速 &lt; 40%，基準的融資缺口明顯縮小。</li>
</ul>
<p class="muted">資料缺口：FY2025 後產品別與方案別營收、分部毛利率、GPU 數與 MW、自建集群規模皆未揭露；2026-10-08 股價未取得（用 10-02 收盤市值 HK$76.15B）。一手／二手與標記見 Excel 的 SRC_MM 與 Inputs。</p></div>
</main><div id="tip"></div>
<script>
const tip=document.getElementById('tip');
document.querySelectorAll('[data-tip]').forEach(el=>{{
 el.addEventListener('mouseenter',()=>{{tip.textContent=el.dataset.tip;tip.style.opacity=1}});
 el.addEventListener('mousemove',e=>{{tip.style.left=(e.clientX+12)+'px';tip.style.top=(e.clientY+12)+'px'}});
 el.addEventListener('mouseleave',()=>{{tip.style.opacity=0}});
}});
</script></body></html>'''
open(out, 'w').write(page)
print(out)
