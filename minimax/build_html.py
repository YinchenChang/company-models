# MiniMax 收支模型 — 一頁摘要 HTML（呈現層）。所有數字讀自已由 LibreOffice 重算的 Excel（含 Sensitivity 頁），
# 每個數字以 <span class="x" data-x="鍵|年|格式"> 標記，scripts/check_html.py 逐一回頭比對 Excel；無外部引用。
# 用法：python3 build_html.py 模型.xlsx [out/sensitivity.json（相容舊參數，不再使用）] 輸出.html
import sys, os, html
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'scripts'))
from xlvals import load, val, FM, Y

xl, out = sys.argv[1], sys.argv[-1]
D = load(xl)


def N(key, y=None, fm='f1'):
    """帶標記的數字（鍵：計算頁 ID／S／I 開頭／'s:情境名'／'tk'）。"""
    return f'<span class="x" data-x="{key}|{"" if y is None else y}|{fm}">{html.escape(FM[fm](val(D, key, y)))}</span>'


def g(k, y): return val(D, k, y)


def sn(name, idx, fm): return N('s:' + name, idx, fm)   # Sensitivity 頁：idx＝輸出欄序（0 2026 營收 … 9 2030 回本比值）


first_fund = g('F19', 2026)
TK = val(D, 'tk')

def line_chart(series, ymax, unit, ident, dec=0, ref=None):
    W, H, L, R, T, B = 640, 260, 48, 16, 16, 34
    xs = [L + i * (W - L - R) / (len(Y) - 1) for i in range(len(Y))]
    sy = lambda v: T + (H - T - B) * (1 - v / ymax)
    out = [f'<svg viewBox="0 0 {W} {H}" class="chart" role="img" aria-labelledby="{ident}-t">']
    step = ymax / 4
    for i in range(5):
        v = step * i; yy = sy(v)
        out.append(f'<line x1="{L}" x2="{W-R}" y1="{yy:.1f}" y2="{yy:.1f}" class="grid"/>'
                   f'<text x="{L-6}" y="{yy+4:.1f}" class="ax" text-anchor="end">{v:,.{dec}f}</text>')
    if ref is not None:
        out.append(f'<line x1="{L}" x2="{W-R}" y1="{sy(ref):.1f}" y2="{sy(ref):.1f}" class="minl"/><text x="{W-R}" y="{sy(ref)-5:.1f}" class="ax" text-anchor="end">＝1：營收剛好覆蓋持有成本</text>')
    for i, y in enumerate(Y):
        out.append(f'<text x="{xs[i]:.1f}" y="{H-12}" class="ax" text-anchor="middle">{y}{"A" if y == 2025 else "E"}</text>')
    for name, vals, var in series:
        pts = ' '.join(f'{xs[i]:.1f},{sy(v):.1f}' for i, v in enumerate(vals))
        out.append(f'<polyline points="{pts}" fill="none" stroke="var({var})" stroke-width="2" stroke-linejoin="round"/>')
        for i, v in enumerate(vals):
            out.append(f'<circle cx="{xs[i]:.1f}" cy="{sy(v):.1f}" r="4" fill="var({var})" stroke="var(--surface)" stroke-width="2" '
                       f'data-tip="{html.escape(name)}｜FY{Y[i]}：{v:,.{dec+2}f} {unit}"/>')
        out.append(f'<text x="{xs[-1]+0:.1f}" y="{sy(vals[-1])-10:.1f}" class="dl" text-anchor="end">{name} {vals[-1]:,.{dec+1 if dec else 1}f}</text>')
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
ratio = [g('X05', y) for y in Y]
c3 = line_chart([('回本比值', ratio, '--s1')], 1.2, '倍', 'c3', dec=1, ref=1.0)
cash = [g('F09', y) for y in Y]
c2 = bar_chart(cash, -2000, 4000, '$M', 300)


def table(rows, head='$M（除註明）'):
    t = ['<table class="t"><thead><tr><th>' + head + '</th>' + ''.join(f'<th>FY{y}{"A" if y==2025 else "E"}</th>' for y in Y) + '</tr></thead><tbody>']
    for k, fm, lab in rows:
        strong = ' class="key"' if k in ('R09', 'P23', 'F11', 'X05') else ''
        t.append(f'<tr{strong}><td>{lab}</td>' + ''.join(f'<td>{N(k, y, fm) if g(k, y) is not None else "—"}</td>' for y in Y) + '</tr>')
    t.append('</tbody></table>')
    return ''.join(t)


tbl = table([('R01', 'f1', '企業端營收（API）'), ('R02', 'f1', '消費端營收'), ('R09', 'f1', '總營收'), ('P07', 'pc', '毛利率'),
             ('P10', 'f1', '訓練算力'), ('P12', 'f1', '人事（含 SBC）'), ('P14', 'f1', '其他營業費用'),
             ('P18', 'f1', '經調整淨利'), ('C21', 'f1', '總 MW（年均）'), ('C22', 'pc', '訓練占 MW'),
             ('P21', 'f2', '每 MW 營收 $M'), ('P22', 'f2', '每 MW 全成本 $M'), ('P23', 'f2', '覆蓋率（x）'),
             ('F09', 'f0', '期末現金（新融資前）'), ('F11', 'f0', '累計需新融資')])
pxq = table([('X01', 'f1', '年均上線 MW'), ('X02', 'f2', '每 MW 年持有成本 h（$M）'), ('X03', 'f0', '算力持有成本 C（$M）'), ('X04', 'f0', '總營收（$M）'),
             ('X05', 'f2', '回本比值＝營收 ÷ C'), ('X06', 'f2', '比值相對 FY2025'), ('X11', 'f2', '實際 m×L'), ('X12', 'f2', '所需 m×L'),
             ('X14', 'f2', '企業端（API）比值'), ('X15', 'f2', '消費端比值'), ('X16', 'pc', '訓練 MW 占比'),
             ('X26', 'f2', '營收 ÷ 算力支出（合約價口徑）'), ('X28', 'f1', '比值＝1 所需合約價 I02（$M）')], 'P×Q（$M／MW／倍）')

SC = ['基準', '合約價 $8M/MW', '合約價 $20M/MW', '合約價 $5.5M（H20 證據價）', '合約價 $9.2M（H800 證據價）', '合約價 $13.4M（H100 證據價）']
rows_i02 = ''.join(f'<tr{" class=key" if n == "基準" else ""}><td>{html.escape(n)}</td><td>{sn(n, 8, "f2")}</td><td>{sn(n, 9, "f2")}</td><td>{sn(n, 3, "x2")}</td></tr>' for n in SC)
SD = ['M3.1 降價：不計入（v0.1 行為）', '基準', 'M3.1 降價：新價流量 50%', 'M3.1 降價：下半年全面 −50%']
rows_px = ''.join(f'<tr{" class=key" if n == "基準" else ""}><td>{html.escape(n)}</td><td>{sn(n, 3, "x2")}</td><td>{sn(n, 5, "f0")}</td><td>{sn(n, 9, "f2")}</td><td>{sn(n, 7, "x2")}</td></tr>' for n in SD)

sens_rows = ['<table class="t s"><thead><tr><th>情境（只改一項）</th><th>2030 營收</th><th>2030 覆蓋率</th><th>2030 回本比值</th><th>累計需新融資</th><th>首次融資年</th><th>折現價值÷市值</th></tr></thead><tbody>']
for name in D['SENS']:
    hl = ' class="key"' if name == '基準' else ''
    sens_rows.append(f'<tr{hl}><td>{html.escape(name)}</td><td>{sn(name, 1, "f0")}</td><td>{sn(name, 3, "x2")}</td><td>{sn(name, 9, "x2")}</td><td>{sn(name, 5, "f0")}</td><td>{sn(name, 6, "yr")}</td><td>{sn(name, 7, "x2")}</td></tr>')
sens_rows.append('</tbody></table>')

ev = ''.join(f'<tr><td>{lab}</td><td>{N(a, 2026, "f2")}</td><td>{N(b, 2026, "f1")}</td><td>{N(c, 2026, "x2")}</td></tr>'
             for lab, a, b, c in (('AutoDL H800 單卡（口徑明確）', 'C32', 'C35', 'C38'), ('SMM H20 96G 年約（8 卡整機口徑，樣本 2 台）', 'C33', 'C36', 'C39'), ('SMM H100 80G 現貨', 'C34', 'C37', 'C40')))

page = f'''<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>MiniMax 收支模型 v0.2</title>
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
.mono{{font-family:ui-monospace,Menlo,monospace;font-size:12px;word-break:break-all}}
</style></head><body><main>
<h1>MiniMax（00100.HK）收支模型 v0.2</h1>
<p class="sub">上海稀宇科技｜比照 OpenAI 收支模型（MW 主軸、P×Q 回本比值、Tokenomics 官方快照）｜FY2025 實際、FY2026 上半年實際＋下半年推估、FY2027–2030 推估｜2026-10-22｜單位 US$M</p>

<div class="card ans"><b>命題：</b>MiniMax 每 MW 算力的營收，能否覆蓋每 MW 全成本？若不能，缺口何時出現、由誰補？<br>
<b>答案（基準）：不能。</b>每 MW 營收由 FY2025 的 ${N('P21',2025,'f2')}M 升到 FY2030 的 ${N('P21',2030,'f2')}M，但每 MW 全成本仍約 ${N('P22',2030,'f1')}M，覆蓋率只從 {N('P23',2025,'f2')}x 升到 {N('P23',2030,'f2')}x。
推論本身已經賺錢（API 毛利率 FY2030 {N('P08',2030,'pc')}），虧損來自訓練：訓練占總 MW {N('C22',2030,'pc')}，訓練＋人事＋其他費用是營收的 {N('P29',2030,'f1')} 倍。
2026 年已募得 IPO ${N('F02',2026,'f0')}M、配售 ${N('F03',2026,'f0')}M、可轉債 ${N('F04',2026,'f0')}M，基準下 <b>{N('F19',2026,'yr')}{'' if first_fund == 9999 else ' 年'}</b>起需要新融資（9999＝2030 年內不需要），至 2030 年累計還需約 <b>${N('F11',2030,'f0')}M</b>（相當目前市值 {N('F14',2030,'pc')} 的新股）。</div>

<div class="tiles">
<div class="tile"><div class="v">{N('X05',2030,'f2')}</div><div class="l">FY2030 回本比值（營收 ÷ 機隊持有成本）；FY2025 為 {N('X05',2025,'f2')}</div></div>
<div class="tile"><div class="v">{N('P23',2030,'f2')}x</div><div class="l">FY2030 覆蓋率（每 MW 營收 ÷ 全成本）</div></div>
<div class="tile"><div class="v">{N('R09',2026,'f0')}</div><div class="l">FY2026 營收（模型）；共識 {N('R13',2026,'f0')}／Bloomberg {N('R15',2026,'f0')}</div></div>
<div class="tile"><div class="v">{N('C21',2026,'f0')} MW</div><div class="l">FY2026 年均算力（約 {N('C25',2026,'f0')} 張 Hopper 等值）</div></div>
<div class="tile"><div class="v">{N('F19',2026,'yr')}</div><div class="l">首次需新融資</div></div>
<div class="tile"><div class="v">{N('V13',2030,'f2')}x</div><div class="l">反向檢驗：模型價值 ÷ 目前市值（HK${N('V01',2026,'f1')}）</div></div>
</div>

<h2>P×Q 回本比值：營收 ÷（年均上線 MW × 每 MW 年持有成本）</h2>
<div class="card"><p class="ans"><b>{N('X07',2025,'txt')}</b></p>
<div class="legend"><span><i class="sw" style="background:var(--s1)"></i>回本比值（營收 ÷ 算力持有成本；實際 m×L ÷ 所需 m×L）</span></div>{c3}
<div class="scroll">{pxq}</div>
<p class="muted">定義（比照 OpenAI v0.6.1）：每 MW 年持有成本 h＝Σ 世代占比 × Tokenomics IF_HoldEcon 基準欄（Hopper 占比 + GB200，v0.1 機隊組合；中國晶片以 Hopper 代理）；算力持有成本 C＝MW × h；比值＝營收 ÷ C，Q 與牌價約掉。P 取 MiniMax 實收單價，Tokenomics 有效牌價（OpenAI 層級牌價）不適用。看的是年度趨勢，不是絕對數字。</p>
<p class="muted"><b>先看敏感度再看基準：</b>所有 MW 皆由算力支出 ÷ 合約價 I02 換算（API 經 η 校準亦同），因此比值＝（營收 ÷ 算力支出）×（I02 ÷ h），<b>水準與 I02 成正比</b>。基準 I02＝${N('I02',None,'f0')}M（Analogy，區間 {N('I02#lo',None,'f0')}–{N('I02#hi',None,'f0')}）下比值未達 1；比值＝1 需 I02 約 ${N('X28',2030,'f1')}M（FY2030）。趨勢（逐年上升）不受 I02 影響。</p>
<div class="scroll"><table class="t s"><thead><tr><th>合約價情境（D2：租價證據並列，基準不動）</th><th>2026 比值</th><th>2030 比值</th><th>2030 覆蓋率</th></tr></thead><tbody>{rows_i02}</tbody></table></div>
<p class="muted">D2 租價證據（中國雲 GPU 租價 ÷ Tokenomics Hopper 持有成本 ${N('TK07',None,'f2')}M，每 MW 以 {N('C31',2026,'f0')} 張計；I02 ÷ 持有成本＝{N('C41',2026,'x2')}）：</p>
<div class="scroll"><table class="t s"><thead><tr><th>證據</th><th>每卡時 US$</th><th>隱含每 MW 年租金 $M</th><th>÷ TK 持有成本</th></tr></thead><tbody>{ev}</tbody></table></div>
<p class="muted">落在 I02 區間 {N('I02#lo',None,'f0')}–{N('I02#hi',None,'f0')} 外的證據共 {N('C42',2026,'f0')} 項（H20 一項，樣本 2 台、口徑不明，若實為單卡則價格高 8 倍）；依 D2 規則基準維持 12，影響見上表。</p>
<p class="muted"><b>M3.1 Flash 降價（D8；2026-10-19，較 M3 牌價 −{N('X19',2026,'pc')}）：</b>只更新已觀測的 2026 年。降價後營收持平所需用量為 {N('X20',2026,'f1')} 倍（全面調降時）；基準假設下半年 {N('I31',None,'pc')} 流量採新價，所需用量 {N('X22',2026,'f2')} 倍、全年平均折減 {N('X23',2026,'pc1')}。</p>
<div class="scroll"><table class="t s"><thead><tr><th>M3.1 降價情境</th><th>2030 覆蓋率</th><th>累計需新融資</th><th>2030 回本比值</th><th>折現價值÷市值</th></tr></thead><tbody>{rows_px}</tbody></table></div></div>

<h2>每 MW 營收 vs 每 MW 全成本</h2>
<div class="card"><div class="legend"><span><i class="sw" style="background:var(--s1)"></i>每 MW 營收</span><span><i class="sw" style="background:var(--s2)"></i>每 MW 全成本（銷售成本＋營業費用，含 SBC）</span></div>{c1}
<p class="muted">每 MW 全成本中算力部分固定為合約價 ${N('I02',None,'f0')}M（雲端租用口徑），其餘為人事與其他費用。覆蓋率不受合約價影響：所有 MW 皆由支出 ÷ 合約價換算，合約價只改變 MW 規模與回本比值的水準。</p></div>

<h2>現金路徑（新融資前）</h2>
<div class="card">{c2}<p class="muted">2026 現金來源：期初 {N('S040',None,'f0')}、IPO 淨額 {N('F02',2026,'f0')}、配售 {N('F03',2026,'f0')}、可轉債 {N('F04',2026,'f0')}；可轉債轉換價 HK${N('S046',None,'f0')} 高於 10-21 推算收盤價 HK${N('V01',2026,'f1')}（已跌破 IPO 價 HK${N('S080',None,'f0')}），基準於 2027 年以現金贖回（流出 {N('F08',2027,'f0')}）。
公告 FY2025 經營活動現金流 {N('S089',None,'f1')}，與模型經調整淨利的差額 {N('F21',2025,'f1')}（負＝實際流出較多；營運資金與其他）；1H26 現金流量表未取得。</p></div>

<h2>逐年數字</h2>
<div class="card scroll">{tbl}</div>

<h2>敏感度：哪些假設會翻轉結論</h2>
<div class="card scroll">{''.join(sens_rows)}
<p class="muted">翻轉覆蓋率 ≥1 或至 2030 不需新融資的只有：API 單價不再下跌、token 量年增遠高於基準、訓練 MW 2027 年起不再成長。合約價與層級組合對覆蓋率沒有影響（被 η 校準吸收），但合約價決定回本比值的水準。</p></div>

<h2>主要驅動、校準與差距</h2>
<div class="card"><ul>
<li><b>API 價量：</b>2026 計費 token 由營收 ÷ 全年平均實收單價 ${N('D01',2026,'f3')}/M 推得（上半年校準價 ${N('I07',2026,'f2')}，公司說法，未見一手；下半年含 M3.1 Flash 已觀測降價 ${N('D09',2026,'f3')}），約 {N('D03',2026,'f0')}T，為 2025 的 {N('D06',2026,'f1')} 倍（公司稱 7 月 token 為 1 月的 {N('S055',None,'f0')} 倍）。2027 起量年增 {N('I09',2027,'pc')}／{N('I09',2028,'pc')}／{N('I09',2029,'pc')}／{N('I09',2030,'pc')}、價年變動 {N('I08',2027,'pc')}／{N('I08',2028,'pc')}／{N('I08',2029,'pc')}／{N('I08',2030,'pc')}。</li>
<li><b>η（V2）：</b>2025 為 {N('C11',2025,'f2')}、2026 上半年為 {N('C11',2026,'f2')}；同一 MW 產出的計費 token 大幅下降，反映 Token Plan 低價 token、長上下文與較大的 M3。下半年起按公司 M3.1 降本說法只採 {N('I29',None,'f1')} 倍改善。Tokenomics 服務端落差分解無 IF_ 等義名稱，只列對照（缺口）。</li>
<li><b>訓練：</b>2026 由上半年研發費用扣人事推得約 ${N('C17',2026,'f0')}M，全年約 ${N('C18',2026,'f0')}M；2027 起 MW 年增 {N('I17',2027,'pc')}／{N('I17',2028,'pc')}／{N('I17',2029,'pc')}／{N('I17',2030,'pc')}。這是虧損的主因，也是最大的判斷項。</li>
<li><b>與市場：</b>FY2026 模型 {N('R09',2026,'f0')} vs 共識 {N('R13',2026,'f0')}（Yahoo／stockanalysis，疑同源）與 {N('R15',2026,'f0')}（Bloomberg）；Bloomberg 共識隱含 2H26 為 1H26 的 {N('R21',2026,'f1')} 倍。FY2027 模型 {N('R09',2027,'f0')} vs 共識 {N('R17',2027,'f0')}／{N('R18',2027,'f0')}；FY2028 {N('R09',2028,'f0')} vs Goldman {N('R13',2028,'f0')}；FY2030 {N('R09',2030,'f0')} vs Visible Alpha {N('R13',2030,'f0')}。共識只作對照，不反推參數。</li>
<li><b>誰補缺口：</b>上市前創投累計約 ${N('F15',2025,'f0')}M；IPO 基石 ADIA、阿里巴巴等；2026-07 配售與可轉債約 ${N('F17',2026,'f0')}M。市值自 10-02 到 10-21 約降 21%，再配售與可轉債轉股條件更不利（轉換價為現價的 {N('S046',None,'f0')}÷{N('V01',2026,'f1')} 倍）；A 股（科創板）情境本版不做（D5）。阿里雲採購上限 2026–28 年 ${N('S059',None,'f0')}／{N('S060',None,'f0')}／{N('S061',None,'f0')}M。</li>
</ul></div>

<h2>失效條件與追蹤</h2>
<div class="card"><ul>
<li>2H26 企業端年化若達 $750M（FY2026 營收 {sn('2H26 企業端年化 750', 0, 'f0')}），首次融資年延到 {sn('2H26 企業端年化 750', 6, 'yr')}，反向檢驗升至 {sn('2H26 企業端年化 750', 7, 'x2')}。</li>
<li>中國 API 價格若持平（單價不再下跌），2030 覆蓋率 {sn('API 單價持平', 3, 'x2')}、2030 回本比值 {sn('API 單價持平', 9, 'x2')}：追蹤 MiniMax 牌價與 OpenRouter 實際成交價。</li>
<li>1H26 毛利率 {N('P31',2026,'pc1')}（公告）；若 2H26 未回升，η 改善假設（I29）偏樂觀（I29＝1.0 時 2030 覆蓋率 {sn('2H26 推論效率改善 1.0', 3, 'x2')}）。</li>
<li>合約價證據：若新的中國雲租價顯示多年合約價長期低於 $9M，回本比值水準下修（上表）。</li>
</ul>
<p class="muted">資料缺口：1H26 現金流量表、自建集群規模（GPU 數／MW）、2026 消費端付費人數與 ARPPU、FY2025 後產品別與方案別營收、GPU 數與 MW、大廠雲官方租價皆未取得；一手／二手與標記見 Excel 的 SRC_MM 與 Inputs。</p></div>

<h2>資料與版本</h2>
<div class="card"><ul>
<li><b>模型：</b>{N('readme',None,'txt')}。Excel 為唯一計算引擎，本頁數字逐一對照 Excel（scripts/check_html.py，納入 CI）。</li>
<li><b>Tokenomics 官方快照：</b><span class="mono">{N('tk',None,'txt')}</span></li>
<li><b>市場資料（M2，擷取 2026-10-22）：</b>10-21 推算收盤價 HK${N('S078',None,'f1')}（Derived）、盤中 {N('S079',None,'f1')}；市值 ${N('V02',2026,'f0')}M；M3.1 Flash 牌價 {N('S098',None,'f2')}／{N('S099',None,'f2')}（輸入／輸出 $/M）；合約負債 {N('S085',None,'f1')}（2026-06-30）。</li>
<li><b>財務一手：</b>1H26 中期業績公告 2026-08-26（銷售成本 {N('S020',None,'f1')}、營收 {N('S004',None,'f1')}）；FY2025 年度業績公告 2026-03-02；FY2025 現金流量為經濟通轉載年報（二手）。</li>
</ul></div>
</main><div id="tip"></div>
<script>
const tip=document.getElementById('tip');
document.querySelectorAll('[data-tip]').forEach(el=>{{
 el.addEventListener('mouseenter',()=>{{tip.textContent=el.dataset.tip;tip.style.opacity=1}});
 el.addEventListener('mousemove',e=>{{tip.style.left=(e.clientX+12)+'px';tip.style.top=(e.clientY+12)+'px'}});
 el.addEventListener('mouseleave',()=>{{tip.style.opacity=0}});
}});
</script></body></html>'''
open(out, 'w', encoding='utf-8').write(page)
print(out)
