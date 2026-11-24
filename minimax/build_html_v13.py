# MiniMax 個股模型頁（個股 HTML 版面規格 v1.3；v0.2.3 起有目標價與評等：Val、RevDCF 頁，照 A–I 節，估值章不再帶 data-novaluation）。
# 數字一律讀自 LibreOffice 重算後的 xlsx（openpyxl data_only），每個數字帶 data-ref（頁!格）與 data-v；本檔不做模型計算，只做圖的幾何。
# 共用：tools/company_html/fundmodel_v2.py（span、fmt、CSS）與 fundmodel_v13.py（長條圖、v1.3 CSS）。需 beautifulsoup4、lxml（共用模組匯入用）。
# 用法：python3 build_html_v13.py <重算後.xlsx> <輸出.html>
import html
import os
import re
import sys

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, 'tools', 'company_html'))
import fundmodel_v2 as M  # noqa: E402
from fundmodel_v13 import C_ACC, C_GOLD, C_NEG, C_PUR, C_TEAL, CSS_V13, bar_chart  # noqa: E402

YRS = [2025, 2026, 2027, 2028, 2029, 2030]
YF = YRS[1:]
sp = M.span


class X:
    def __init__(self, path):
        self.wb = openpyxl.load_workbook(path, data_only=True)
        self.rows = {}
        for sh in ('Demand', 'Revenue', 'Compute', 'PnL', 'PxQ', 'Funding', 'Valuation', 'SRC_MM', 'Inputs'):
            ws = self.wb[sh]
            for r in range(1, ws.max_row + 1):
                v = ws.cell(r, 1).value
                if isinstance(v, str) and re.fullmatch(r'[A-Z]\d{2,3}', v):
                    self.rows.setdefault(v, (sh, r))

    def at(self, sh, cell):
        return f'{sh}!{cell}', self.wb[sh][cell].value

    def y(self, rid, yr):
        sh, r = self.rows[rid]
        return self.at(sh, f'{"DEFGHI"[yr - 2025]}{r}')

    def ys(self, rid, yrs=YRS):
        return [self.y(rid, y) for y in yrs]

    def s(self, sid):
        sh, r = self.rows[sid]
        return self.at(sh, f'E{r}')

    def i(self, iid):
        sh, r = self.rows[iid]
        return self.at(sh, f'D{r}')

    def name(self, nm):
        dn = self.wb.defined_names[nm]
        sh, rng = next(iter(dn.destinations))
        rng = rng.replace('$', '')
        ws = self.wb[sh]
        if ':' in rng:
            return [[(f'{sh}!{c.coordinate}', c.value) for c in row] for row in ws[rng]]
        return [[(f'{sh}!{rng}', ws[rng].value)]]


def V(B, nm):
    """具名範圍（單格）→ (頁!格, 值)。"""
    return B.name(nm)[0][0]


def SR(B, prefix, col):
    """Sensitivity 頁依案例名稱（前綴）找列。"""
    ws = B.wb['Sensitivity']
    r = next(r for r in range(4, ws.max_row + 1) if str(ws.cell(r, 1).value or '').startswith(prefix))
    return B.at('Sensitivity', f'{col}{r}')


def T(B, sh, cell):
    """xlsx 文字格（標籤）。"""
    return sp(*B.at(sh, cell), 't')


def tbl(head, rows, cls='drv'):
    th = ''.join(f'<th>{h}</th>' for h in head)
    return f'<div class="tw"><table class="{cls}"><thead><tr>{th}</tr></thead><tbody>{"".join(rows)}</tbody></table></div>'


def yrow(B, lab, rid, f, cls='', yrs=YRS):
    return f'<tr class="{cls}"><td>{lab}</td>' + ''.join(f'<td>{sp(r, v, f) if v is not None else "—"}</td>' for r, v in B.ys(rid, yrs)) + '</tr>'


YH = [f'FY{y}' for y in YRS]


# ── 橫幅 ──
def hero(B):
    rec, vr = V(B, 'VAL_Rec'), V(B, 'VAL_Rating')
    cls = {'強力放空': 'hero-neg', '放空': 'hero-neg', '中立': 'hero-neu', '買進': 'hero-pos', '強力買進': 'hero-pos'}.get(str(rec[1]), 'hero-neg')
    w, uw, px = V(B, 'VAL_TP_W'), V(B, 'VAL_Up_W'), V(B, 'VAL_Px')
    tiles = (f'<div class="hero-t"><div class="hero-l">目標價（DCF 與同業倍數各半，港元）</div><div class="hero-n">{sp(*w, "n2")}</div>'
             f'<div class="hero-s">上檔空間 {sp(*uw, "sp1")}</div></div>'
             f'<div class="hero-t"><div class="hero-l">現價（2026-11-20 收盤，港元）</div><div class="hero-n">{sp(*px, "n2")}</div>'
             f'<div class="hero-s">市場共識目標價 {sp(*V(B, "VAL_TP_Cons"), "n2")}（2026-09-30）</div></div>'
             f'<div class="hero-t"><div class="hero-l">收支關卡：FY2026–FY2030 現金收支合計（百萬美元）</div><div class="hero-n">{sp(*V(B, "VAL_Gate"), "sn0")}</div>'
             f'<div class="hero-s">入不敷出；估值評等 {sp(*vr, "t")}</div></div>')
    floor = V(B, 'VAL_DCFFloorAll')[1] == 1
    note = (f'同業倍數法（以同日的智譜倍數計）給 {sp(*V(B, "VAL_TP_Comp"), "n0")} 港元；DCF '
            + ('因基準情境到 FY2030 仍大幅虧損，只剩淨現金價值 ' if floor else '給 ') + f'{sp(*V(B, "VAL_TP_DCF"), "n0")} 港元。'
            f'兩者各半得 {sp(*w, "n0")} 港元，上檔空間 {sp(*uw, "sp0")}，估值評等{sp(*vr, "t")}。'
            + ('收支關卡不通過（入不敷出），依規則不給買進。' if V(B, 'VAL_GateFail')[1] == 1 else ''))
    return (f'<div class="hero {cls}"><div class="hero-rec"><span class="hero-k">投資建議</span><span class="hero-v">{sp(*rec, "t")}</span></div>'
            f'<div class="hero-tiles">{tiles}</div><p class="hero-note">{note}</p></div>')


# ── 1 命題 ──
def thesis(B):
    items = ['每 MW 收支：MiniMax 每 MW 算力（推論＋訓練）的營收，能否覆蓋每 MW 全成本（算力、人事、其他），哪一年打平。',
             '錢夠不夠：IPO、配售與可轉債募得的現金，能否撐過虧損期；不夠時何時、要多少新融資，對現有股東稀釋多少。',
             '值多少：DCF（現金流沿用本收支模型）與同業倍數各半的目標價與評等；收支關卡不通過時不給買進。']
    rep = (f'<ol class="items"><li>每 MW 能否打平？答：基準情境預測期內都未打平，覆蓋率由 FY2026 的 {sp(*B.y("P23", 2026), "n2")} 倍升到 FY2030 的 {sp(*B.y("P23", 2030), "n2")} 倍。</li>'
           f'<li>錢夠不夠？答：2026 年募資後年底現金 {sp(*B.y("F09", 2026), "n0")} 百萬美元，{sp(*B.y("F19", 2026), "t")} 年起需新融資，至 FY2030 累計 {sp(*B.y("F11", 2030), "n0")}，'
           f'稀釋約 {sp(*B.y("F14", 2030), "p0")}（占目前股數）。</li>'
           f'<li>現價隱含什麼？答：FY2030 營收 {sp(*B.y("R09", 2030), "n0")} 百萬美元、退出倍數 {sp(*B.i("I27"), "n0")} 倍折回今天，價值為市值的 {sp(*B.y("V13", 2030), "n2")} 倍；'
           f'支持現價需 FY2030 營收 {sp(*B.y("V14", 2030), "n0")} 百萬美元。</li></ol>')
    obs = [
        f'<li><b>營收長得快，但成本長得更快：</b>FY2026 營收 {sp(*B.y("R09", 2026), "n0")} 百萬美元（年增 {sp(*B.y("R10", 2026), "p0")}），'
        f'FY2030 達 {sp(*B.y("R09", 2030), "n0")}；同期經調整淨虧損由 {sp(*B.y("P18", 2026), "n0")} 擴大到 {sp(*B.y("P18", 2030), "n0")}，主因訓練算力與人事。</li>',
        f'<li><b>推論本身賺錢、訓練拖累：</b>企業端（API）每推論 MW 營收是合約價的 {sp(*B.y("P28", 2030), "n2")} 倍（FY2030），'
        f'但訓練占總 MW {sp(*B.y("C22", 2030), "p0")}，整體回本比值（營收 ÷ 機隊持有成本）FY2030 只有 {sp(*B.y("X05", 2030), "n2")} 倍。</li>',
        f'<li><b>現金撐到 {sp(*B.y("F19", 2026), "t")} 年：</b>2026 年 IPO 淨募資 {sp(*B.y("F16", 2026), "n0")}、配售與可轉債 {sp(*B.y("F17", 2026), "n0")} 百萬美元，'
        f'但可轉債 FY2027 到期以現金贖回 {sp(*B.y("F08", 2027), "sn0")}；之後每年要新融資。</li>',
        f'<li><b>最敏感的是 token 量：</b>token 量成長高案下 FY2030 覆蓋率 {sp(*SR(B, "token 量成長高", "F"), "n2")} 倍、累計新融資 {sp(*SR(B, "token 量成長高", "H"), "n0")}；'
        f'低案累計新融資增加到 {sp(*SR(B, "token 量成長低", "H"), "n0")} 百萬美元（敏感度章）。</li>',
    ]
    ch1 = bar_chart(sp, YRS, [('每 MW 營收', C_TEAL, B.ys('P21')), ('每 MW 全成本', C_NEG, B.ys('P22'))], '百萬美元／MW／年',
                    '問題一：每 MW 營收 vs 每 MW 全成本', fmt='n1', note='全成本含推論與訓練算力（合約價口徑）、人事（含股份支付）與其他營業費用；FY2025 為實績校準年。')
    ch2 = bar_chart(sp, YRS, [('期末現金（新融資前）', 'auto', B.ys('F09'))], '百萬美元', '問題二：期末現金（新融資前；負＝缺口）', fmt='sn0',
                    note='低於最低現金水位的差額由新融資補足（資金章）。')
    return (f'<section id="thesis"><h2>命題與核心問題</h2>'
            f'<p class="thesis">命題：MiniMax 每 MW 算力的營收能否覆蓋每 MW 全成本；若不能，缺口何時出現、由誰、以什麼條件補。</p>'
            f'<p class="q"><strong>核心問題：</strong>每 MW 營收能否覆蓋全成本？現金能撐多久、要多少新融資？現價隱含多少營收？</p>'
            f'<h3>命題</h3><ol class="items">{"".join(f"<li>{html.escape(t)}</li>" for t in items)}</ol>'
            f'<details class="more"><summary>展開：命題回覆（依本模型計算的答案）</summary>{rep}</details>'
            f'<h3>回答核心問題的圖</h3><div class="charts2">{ch1}{ch2}</div>'
            f'<h3>重點觀察</h3><ol class="items">{"".join(obs)}</ol></section>')


# ── 2 估值 ──
def val_rows(B, keys):
    """Val 頁 A 欄標籤 → (標籤 span, 值 span)；keys＝[(標籤開頭, 欄, 格式)]。"""
    ws = B.wb['Val']
    lab = {}
    for r in range(1, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if isinstance(a, str):
            lab.setdefault(a, r)
    out = []
    for pre, col, f in keys:
        r = next(r for a, r in lab.items() if a.startswith(pre))
        out.append((sp(f'Val!A{r}', ws.cell(r, 1).value, 't'), sp(f'Val!{col}{r}', ws.cell(r, ord(col) - 64).value, f)))
    return out


def valuation(B):
    px = V(B, 'VAL_Px')
    four = [('DCF 目標價（單純 DCF，三情境加權）', 'VAL_TP_DCF', 'VAL_Up_DCF', '無槓桿自由現金流折現；基準情境到 FY2030 仍虧損，三情境皆觸及淨現金下限'),
            ('Comparable 目標價（同業 EV ÷ 未來 12 個月營收，三情境加權）', 'VAL_TP_Comp', 'VAL_Up_Comp', '同業只有智譜（同日 2026-11-20 市值）；未做相對倍數調整'),
            ('加權目標價（DCF 與 Comparable 各半；橫幅主值）', 'VAL_TP_W', 'VAL_Up_W', '評等依此判定'),
            ('市場共識目標價（Investing.com 23 位平均，2026-09-30）', 'VAL_TP_Cons', 'VAL_Up_Cons', '二手、早於現價日一個多月，只作對照')]
    ch = bar_chart(sp, ['DCF', 'Comparable', '加權', '市場共識', '現價'],
                   [('港元／股', C_ACC, [V(B, n) for _, n, _, _ in four] + [px])], '港元／股',
                   '四個目標價與現價（港元；現價 2026-11-20 收盤）', fmt='n1', h=260)
    t4 = tbl(['目標價', '值（港元）', '對現價上檔空間', '說明'],
             [f'<tr><td>{a}</td><td>{sp(*V(B, n), "n2")}</td><td>{sp(*V(B, u), "sp1")}</td><td style="text-align:left">{c}</td></tr>' for a, n, u, c in four], 'rd')
    rate = (f'<h3>評等規則與邊界揭露</h3><p class="note">{sp(*V(B, "VAL_RatingNote"), "t")}</p>'
            f'<p class="note">升一級（{sp(*V(B, "VAL_Rating"), "t")} → 上一級）所需加權目標價 {sp(*V(B, "VAL_UpFlip"), "n2")} 港元；'
            f'收支關卡 FY2026–FY2030 現金收支合計 {sp(*V(B, "VAL_Gate"), "sn0")} 百萬美元（負＝入不敷出：估值評等為買進以上時改為「收支不足」）。</p>')
    wacc = tbl(['折現率與股數', '值'], [f'<tr><td>{a}</td><td>{b}</td></tr>' for a, b in val_rows(B, [
        ('無風險利率', 'C', 'p2'), ('β（', 'C', 'n3'), ('股權風險溢酬', 'C', 'p1'), ('股東權益成本', 'C', 'p2'), ('WACC＝', 'C', 'p2'), ('永續成長率', 'C', 'p1'),
        ('評價股數', 'C', 'n2'), ('淨現金（評價日', 'C', 'n0'), ('股權市值 E', 'C', 'n0'), ('市場 EV', 'C', 'n0')])], 'rd')
    dcf = tbl(['DCF 彙總（百萬美元）', '情境 A', '情境 B', '情境 C'], [
        '<tr><td>' + a + '</td>' + ''.join(f'<td>{x}</td>' for x in (b, c, d)) + '</tr>' for a, b, c, d in
        [(x[0][0], x[0][1], x[1][1], x[2][1]) for x in zip(*[val_rows(B, [(k, col, f) for k, f in (
            ('明確期現值合計', 'n0'), ('終值現值', 'n0'), ('企業價值 EV', 'n0'), ('股權價值＝', 'n0'), ('DCF 每股價值（目標時點；港元）', 'n2'),
            ('Comparable 每股（港元）', 'n2'), ('情境目標價', 'n2'))]) for col in 'CDE'])]], 'rd')
    peer = tbl(['同業（2026-11-20 市值）', '值'], [f'<tr><td>{a}</td><td>{b}</td></tr>' for a, b in val_rows(B, [
        ('同業 EV／NTM 營收中位數', 'C', 'n2'), ('對照：本公司市場 EV ÷ 共識 NTM 營收', 'C', 'n2'), ('對照：本公司倍數 ÷ 同業中位數', 'C', 'n2'),
        ('本模型 NTM 營收', 'D', 'n0')])], 'rd')
    old = [f'<tr><td>{T(B, "Valuation", f"B{B.rows[k][1]}")}</td><td>{sp(*B.y(k, y), f)}</td></tr>' for k, y, f in
           (('V09', 2030, 'n0'), ('V12', 2030, 'n0'), ('V13', 2030, 'n2'), ('V14', 2030, 'n0'))]
    return ('<section id="valuation"><h2>估值</h2>'
            f'<div class="figblk">{ch}</div>'
            '<p class="note">Andy 2026-11-15 決定 MiniMax 給目標價與評等。估值頁（xlsx Val 頁）＝DCF（無槓桿自由現金流，現金流沿用本收支模型）50%＋同業倍數 50%，'
            '三情境（API token 量成長高／基準／低）25／50／25。折現率、同業選擇與權重為判斷類參數，依 Anthropic 估值頁樣板慣例預設，待確認。</p>'
            + t4 + rate + '<h3>折現率、股數與淨現金</h3>' + wacc
            + '<p class="note"><b>β 的來源：</b>MiniMax 上市未滿一年，自身 β 期間不足，借用 Anthropic 估值頁的美國軟體同業 5 年 β 中位數'
            '（Palantir、CrowdStrike、Cloudflare、Snowflake、Datadog；stockanalysis.com 2026-11-06），再經 Blume 調整；屬判斷類。</p>'
            '<p class="note"><b>DCF 淨現金下限：</b>三情境皆觸及下限，下限以評價日（2026-06-30 pro forma）淨現金按股東權益成本推到目標時點，'
            '未扣評價日之後的燒錢；若改用本模型 2027-11 的淨現金（年底期末值內插，約 1,298 百萬美元），加權目標價約 131.7 港元，評等仍為中立，但距放空邊界只剩約 2.4 個百分點。</p>'
            + '<h3>DCF 與同業法（三情境）</h3>' + dcf + '<h3>同業倍數</h3>' + peer
            + '<p class="note">MiniMax 與智譜互為唯一同業（智譜模型的同業也只有 MiniMax），兩家的 Comparable 結論本質上是同一組相對價差：一家偏貴即另一家偏便宜，不是兩個獨立的檢驗。</p>'
            + f'<details class="more"><summary>展開：估值方法（模型原文）</summary><p class="note">{sp(*V(B, "VAL_Method"), "t")}</p></details>'
            + '<details class="more"><summary>展開：原收支模型的反向檢驗（2030 年退出倍數折現；只作對照）</summary>' + tbl(['項目（FY2030）', '值'], old, 'rd')
            + '<p class="note">本摺疊沿用原收支模型 Valuation 頁的股數基準 349.2 百萬股（2026-07-14 配售後已發行股數，中期業績公告）；估值章正文的評價股數為 2026-09-30 已發行股份 352.5 百萬股，兩者不同。</p>'
            + '</details></section>')


# ── 3 敏感度 ──
def sens_val(B):
    ws = B.wb['Val']
    r0 = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == '案例')
    rows = []
    r = r0 + 1
    while isinstance(ws.cell(r, 2).value, (int, float)) and not str(ws.cell(r, 1).value).startswith('Sens 快照'):
        g = lambda c: (f'Val!{c}{r}', ws[f'{c}{r}'].value)  # noqa: E731
        rows.append(dict(name=g('H'), lo=g('I'), hi=g('J'), tlo=g('B'), thi=g('C'), plo=g('F'), phi=g('G')))
        r += 1
    rb = next(x for x in range(r, ws.max_row + 1) if str(ws.cell(x, 1).value or '').startswith('Sens 快照的基準加權目標價'))
    base = (f'Val!C{rb}', ws[f'C{rb}'].value)
    mx = max([abs(x[k][1]) for x in rows for k in ('plo', 'phi')] + [0.1])

    def line(x, k, lk):
        ref, v = x[k]
        w = abs(v) / mx * 50
        bar = f'<i class="p" style="left:50%;width:{w:.2f}%"></i>' if v >= 0 else f'<i class="n" style="left:{50 - w:.2f}%;width:{w:.2f}%"></i>'
        return f'<div class="trw"><span>{sp(*x["name"], "t")}（{sp(*x[lk], "t")}）</span><div class="ttr">{bar}</div><span class="tvv">{sp(ref, v, "sp1")}</span></div>'

    def fig(k, lk, cap):
        nz = sorted([x for x in rows if abs(x[k][1]) > 1e-12], key=lambda x: -abs(x[k][1]))
        big = [line(x, k, lk) for x in nz if abs(x[k][1]) >= 0.10]
        small = [line(x, k, lk) for x in nz if abs(x[k][1]) < 0.10]
        body = ''.join(big) if big else '<p class="note">沒有案例使加權目標價變動達 10%。</p>'
        more = f'<details class="more"><summary>展開：{cap}，變動未達 10% 的案例（{len(small)} 項）</summary>{"".join(small)}</details>' if small else ''
        return f'<div class="figblk"><h3>加權目標價相對基準：{cap}</h3>{body}{more}</div>'
    allt = tbl(['案例', '低值設定', '加權目標價', '相對基準', '高值設定', '加權目標價', '相對基準'],
               [f'<tr><td>{sp(*x["name"], "t")}</td><td>{sp(*x["lo"], "t")}</td><td>{sp(*x["tlo"], "n2")}</td><td>{sp(*x["plo"], "sp1")}</td>'
                f'<td>{sp(*x["hi"], "t")}</td><td>{sp(*x["thi"], "n2")}</td><td>{sp(*x["phi"], "sp1")}</td></tr>' for x in rows], 'rd')
    intro = (f'<p class="note">一次只動一個輸入（取低值或高值），整本模型重算，看加權目標價（基準 {sp(*base, "n2")} 港元）相對基準變動多少。'
             '長條依影響大小排序；只直接列出變動達 10% 的案例，其餘收在摺疊；變動為 0 的案例（DCF 觸及淨現金下限時，折現率、永續成長等不影響目標價）只列在全部案例表。</p>')
    return fig('plo', 'lo', '輸入取低值') + fig('phi', 'hi', '輸入取高值') + intro + f'<details class="more"><summary>展開：全部案例表（{len(rows)} 項）</summary>{allt}</details>'


# ── 3 敏感度（對累計需新融資） ──
def sens_fund(B):
    ws = B.wb['Sensitivity']
    rows = []
    for r in range(4, ws.max_row + 1):
        n, v = ws.cell(r, 1).value, ws.cell(r, 8).value
        if isinstance(n, str) and isinstance(v, (int, float)):
            rows.append((n, f'Sensitivity!H{r}', v, f'Sensitivity!A{r}'))
    base = rows[0]
    bv = base[2]
    cases = sorted(rows[1:], key=lambda x: -abs(x[2] - bv))
    mx = max([abs(x[2] - bv) for x in cases] + [1e-9])

    def line(x):
        d = x[2] - bv
        w = abs(d) / mx * 50
        bar = (f'<i class="n" style="left:50%;width:{w:.2f}%"></i>' if d >= 0 else f'<i class="p" style="left:{50 - w:.2f}%;width:{w:.2f}%"></i>')
        return f'<div class="trw"><span>{sp(x[3], x[0], "t")}</span><div class="ttr">{bar}</div><span class="tvv">{sp(x[1], x[2], "n0")}</span></div>'
    big = [line(x) for x in cases if abs(x[2] - bv) / bv >= 0.10]
    small = [line(x) for x in cases if abs(x[2] - bv) / bv < 0.10]
    full = [f'<tr><td>{sp(f"Sensitivity!A{r}", ws.cell(r, 1).value, "t")}</td>' + ''.join(
        f'<td>{sp(f"Sensitivity!{c}{r}", ws.cell(r, ci).value, f)}</td>' for c, ci, f in
        (('D', 4, 'n0'), ('F', 6, 'n2'), ('H', 8, 'n0'), ('I', 9, 't'), ('J', 10, 'n2'), ('L', 12, 'n2'))) + '</tr>'
        for r in range(4, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, str) and isinstance(ws.cell(r, 8).value, (int, float))]
    return ('<div>'
            f'<h3>至 FY2030 累計需新融資：各案例（基準 {sp(base[1], bv, "n0")} 百萬美元；中線＝基準，右紅＝需求增加，左青＝需求減少）</h3>'
            + ''.join(big) + (f'<details class="more"><summary>展開：變動未達 10% 的案例（{len(small)} 項）</summary>{"".join(small)}</details>' if small else '')
            + '</div><p class="note">收支關卡的敏感度：一次只動一個輸入，整本模型重算，看至 FY2030 累計需新融資變動多少；'
              '只直接列出變動達 10% 的案例。合約價改變只影響 MW 與回本比值的換算，不改變營收與現金（全部案例表可見）。</p>'
            + '<details class="more"><summary>展開：全部案例（營收、覆蓋率、累計需新融資、首次融資年、反向檢驗、回本比值）</summary>'
            + tbl(['案例', 'FY2030 營收', 'FY2030 覆蓋率', '累計需新融資', '首次融資年', '折現價值÷市值', 'FY2030 回本比值'], full, 'rd') + '</details></div>')


def sens(B):
    return ('<section id="sens"><h2>敏感度</h2>' + sens_val(B)
            + '<details class="more"><summary>展開：收支關卡的敏感度（至 FY2030 累計需新融資）</summary>' + sens_fund(B) + '</details></section>')


def revdcf(B):
    R = lambda n: V(B, n)  # noqa: E731

    def tr(lab, imp, base, fi, fb, note):
        return f'<tr><td>{lab}</td><td><b>{sp(*R(imp), fi)}</b></td><td>{sp(*R(base), fb) if base else "—"}</td><td style="text-align:left">{note}</td></tr>'
    main = tbl(['現價隱含什麼', '現價隱含', '模型基準', '讀法'], [
        tr('支持現價所需的企業價值（百萬美元）', 'RVD_EVReq', 'RVD_EVB', 'n0', 'n0', 'DCF 情境 B 路徑；模型 EV 為負＝基準現金流不支持任何正價值'),
        tr('(b) 隱含終值期營業利益率（只改 2033 年）', 'RVD_TermMargin', 'RVD_TermMargin0', 'p1', 'p1', '其餘不變，永續期營業利益率要到這個水準'),
        tr('(c) 隱含 2031–2033 營業利益率加碼（相對 FY2030）', 'RVD_ExtAdd', None, 'p0', 'p0', '對應 2033 年營業利益率 ' + sp(*R('RVD_ExtMargin33'), 'p0') + '（不計稅的上限讀法）'),
        tr('(d) 隱含 EV ÷ 未來 12 個月營收（本模型營收）', 'RVD_CompMult', 'RVD_PeerMult', 'n2', 'n2', '模型欄為同業（智譜）中位數'),
        tr('　隱含相對同業係數', 'RVD_CompCoef', None, 'n2', 'n2', '1＝與同業同價'),
        tr('　支持現價所需的未來 12 個月營收（同業倍數；百萬美元）', 'RVD_NTMReq', 'RVD_NTMB', 'n0', 'n0', '模型欄為本模型基準')], 'rd')
    read = ('<p class="note">讀法：以 DCF 看，基準情境到 FY2030 仍大幅虧損，要支持現價 ' + sp(*R('RVD_Px'), 'n2') + ' 港元，永續期營業利益率須達 '
            + sp(*R('RVD_TermMargin'), 'p0') + '；以同業倍數看，現價隱含的營收倍數是同業的 ' + sp(*R('RVD_CompCoef'), 'n2') + ' 倍。'
            'WACC 與營收等比例放大的反解不做：基準自由現金流各年皆負，降低折現率或放大營收只會使價值更低，無解。</p>')
    chk = (f'<p class="note">核對（應為 0）：(b) {sp(*R("RVD_TermChk"), "n6")}；(d) {sp(*R("RVD_CompChk"), "n6")}。計算在 xlsx 的 RevDCF 頁，該頁不被任何計算頁引用。</p>')
    return f'<section id="revdcf"><h2>反向 DCF：現價隱含什麼</h2>{main}{read}{chk}</section>'


# ── 5 損益 ──
def pl(B):
    names = [('PL_Rev', 'n1', 'b'), ('PL_RevYoY', 'p1', 'r'), ('PL_COGS', 'n1', ''), ('PL_GP', 'sn1', 'b'), ('PL_GPM', 'p1', 'r'),
             ('PL_Opex', 'n1', ''), ('PL_OI', 'sn1', 'b'), ('PL_OIM', 'p1', 'r'), ('PL_Oth', 'sn1', ''), ('PL_PBT', 'sn1', 'b'),
             ('PL_Tax', 'n1', ''), ('PL_NI', 'sn1', 'b'), ('PL_NIM', 'p1', 'r'), ('PL_EPS', 'sn2', 'b'), ('PL_NIYoY', 'p1', 'r')]
    yrs = list(range(2023, 2031))
    th = ''.join(f'<th class="{"fc" if y >= 2026 else ""}">FY{y}{"<br>1H 實績＋2H 預測" if y == 2026 else ("<br>預測" if y > 2026 else "")}</th>' for y in yrs)
    trs = []
    for nm, f, cls in names:
        g = B.name(nm)[0]
        r0 = int(re.search(r'(\d+)$', g[0][0]).group(1))
        lab = sp(*B.at('PL', f'A{r0}'), 't')
        cells = ''.join(f'<td class="{"fc" if y >= 2026 else ""}">{(sp(r, v, f) if isinstance(v, (int, float)) else (sp(r, v, "t") if v else "—"))}</td>'
                        for (r, v), y in zip(g, yrs))
        trs.append(f'<tr class="{cls}"><td>{lab}</td>{cells}</tr>')
    table = f'<div class="tw"><table class="pl"><thead><tr><th>百萬美元（Non-GAAP：非 IFRS 經調整）</th>{th}</tr></thead><tbody>{"".join(trs)}</tbody></table></div>'
    basis = ('<p class="basis"><strong>口徑：Non-GAAP（非 IFRS 經調整）。</strong>公司業績公告以經調整淨虧損為主要指標，市場亦以此口徑比較；'
             'IFRS 淨虧損受上市前優先股公允價值變動扭曲（FY2025 該項損失遠大於營運虧損），不適合比較。FY2023–FY2025 為招股書與年度業績公告實績；FY2026 欄為上半年實績校準＋下半年預測（本模型基準情境），FY2027 起為基準情境。</p>')
    g = lambda nm: B.name(nm)[0]  # noqa: E731
    ch = bar_chart(sp, list(range(2023, 2031)), [('營收', C_TEAL, g('PL_Rev')), ('營業利益（非 IFRS）', C_NEG, g('PL_OI'))], '百萬美元',
                   '營收與營業利益（FY2023–FY2025 實績；FY2026 起基準情境）', fmt='sn0', h=290)
    notes = ('<details class="more"><summary>展開：各列口徑</summary><ul class="note">'
             '<li>營業費用＝銷售及分銷＋行政＋研發，扣除股份支付與上市開支；預測期＝訓練算力＋人事（扣股份支付）＋其他營業費用。</li>'
             '<li>其他收入（淨額）：歷史＝其他收入及收益淨額扣財務成本與金融資產減值；預測＝利息與理財收益（本模型未建財務成本）。</li>'
             '<li>所得稅：歷史三年 IFRS 稅前與淨利相同，所得稅為零；預測期本模型不計稅。</li>'
             '<li>EPS：歷史非 IFRS 每股未揭露（公司只揭露 IFRS 每股虧損），不補編；預測＝經調整淨利 ÷ 稀釋後股數（配售後股數 × 一加累計新股稀釋）。</li>'
             f'<li>核對：PL 頁核對列合計 {sp(*B.name("PL_CheckSum")[0][0], "n3")}（應為零）。</li></ul></details>')
    return f'<section id="pl"><h2>損益（標準損益表）</h2>{basis}{ch}{table}{notes}</section>'


# ── 6 驅動因子 ──
def drivers(B):
    seg = tbl(['分部營收推導（百萬美元）'] + YH, [
        yrow(B, 'API 計費 token（兆／年）', 'D03', 'n1'), yrow(B, '× API 綜合實收單價（美元／百萬 token）', 'D01', 'n3'),
        yrow(B, '＝企業端營收（API；FY2025–FY2026 為實績與年化錨）', 'R01', 'n1', 'eq'),
        yrow(B, '消費端營收年增', 'R10', 'p0'),
        yrow(B, '＝消費端營收（AI 原生產品：Talkie／星野、海螺、Agent、Audio）', 'R02', 'n1', 'eq'),
        yrow(B, '＝總營收（企業端＋消費端）', 'R09', 'n1', 'eq'), yrow(B, '企業端占比', 'R11', 'p0')])
    ch1 = bar_chart(sp, YRS, [('企業端（API）', C_TEAL, B.ys('R01')), ('消費端', C_GOLD, B.ys('R02'))], '百萬美元',
                    '分部營收（堆疊；頂端數字＝總營收）', stacked=True, totals=B.ys('R09'), fmt='n0')
    ch2 = bar_chart(sp, YRS, [('API 計費 token', C_TEAL, B.ys('D03'))], '兆 token／年', '量：API 計費 token', fmt='n0')
    ch3 = bar_chart(sp, YRS, [('API 綜合實收單價', C_ACC, B.ys('D01'))], '美元／百萬 token', '價：API 綜合實收單價', fmt='n2')
    cost = tbl(['成本面（百萬美元）'] + YH, [yrow(B, 'API 推論算力', 'P02', 'n0'), yrow(B, '消費端推論算力', 'P03', 'n0'),
                                       yrow(B, '訓練算力（研發）', 'P10', 'n0'), yrow(B, '人事成本（含股份支付）', 'P12', 'n0'),
                                       yrow(B, '其他營業費用', 'P14', 'n0'), yrow(B, '經調整淨利', 'P18', 'sn0', 'eq')])
    return ('<section id="drivers"><h2>驅動因子 → 推導量</h2>'
            '<p class="note">企業端營收＝API 計費 token × 綜合實收單價（FY2025 實績、FY2026 上半年實績＋下半年年化錨；FY2027 起 token 量成長、單價逐年下降）；'
            '消費端營收依成長路徑延伸。成本面：推論算力由 token 換算 MW（Tokenomics 每 GW 產能）乘合約價，訓練與消費端算力由支出換算。</p>'
            f'<h3>分部營收怎麼算出來</h3>{ch1}{seg}<div class="charts2">{ch2}{ch3}</div><h3>成本面</h3>{cost}</section>')


# ── 7 個別模組 ──
def mod_unit(B):
    ch = bar_chart(sp, YRS, [('API 推論', C_TEAL, B.ys('C13')), ('消費端推論', C_GOLD, B.ys('C16')), ('訓練', C_PUR, B.ys('C19'))], 'MW（年均）',
                   '算力：推論與訓練 MW（堆疊；頂端＝總 MW）', stacked=True, totals=B.ys('C21'), fmt='n0')
    t = tbl(['每 MW 收支'] + YH, [yrow(B, '總 MW（年均）', 'C21', 'n1'), yrow(B, '每 MW 營收（百萬美元）', 'P21', 'n2'),
                                yrow(B, '每 MW 全成本（百萬美元）', 'P22', 'n2'), yrow(B, '覆蓋率（倍）', 'P23', 'n2', 'eq'),
                                yrow(B, 'API 每推論 MW 營收 ÷ 合約價（倍）', 'P28', 'n2'), yrow(B, '訓練占總 MW', 'C22', 'p0')])
    return f'<section id="unit"><h2>算力與每 MW 收支</h2>{ch}{t}</section>'


def mod_funding(B):
    ch = bar_chart(sp, YF, [('營運燒錢（經調整淨利）', C_NEG, B.ys('F06', YF)), ('需新融資', C_GOLD, B.ys('F10', YF))], '百萬美元',
                   '資金：每年營運燒錢與需新融資', fmt='sn0')
    t = tbl(['資金（百萬美元）'] + [f'FY{y}' for y in YF], [yrow(B, '期初現金', 'F01', 'n0', yrs=YF), yrow(B, '已完成融資：IPO', 'F02', 'n0', yrs=YF),
                                                        yrow(B, '已完成融資：配售', 'F03', 'n0', yrs=YF), yrow(B, '已完成融資：可轉債', 'F04', 'n0', yrs=YF),
                                                        yrow(B, '營運燒錢', 'F06', 'sn0', yrs=YF), yrow(B, '算力預付與自建資本支出', 'F07', 'sn0', yrs=YF),
                                                        yrow(B, '可轉債到期現金贖回', 'F08', 'sn0', yrs=YF), yrow(B, '期末現金（新融資前）', 'F09', 'sn0', 'eq', yrs=YF),
                                                        yrow(B, '累計需新融資', 'F11', 'n0', 'eq', yrs=YF), yrow(B, '累計新股稀釋', 'F14', 'p1', yrs=YF)])
    return f'<section id="funding"><h2>資金：現金、缺口與融資</h2>{ch}{t}</section>'


def mod_payback(B):
    ch = bar_chart(sp, YRS, [('回本比值（營收 ÷ 機隊持有成本）', 'auto', B.ys('X05'))], '倍', '回本比值（一倍＝營收剛好覆蓋 Tokenomics 機隊持有成本）', fmt='n2')
    t = tbl(['回本比值'] + YH, [yrow(B, '年均上線 MW', 'X01', 'n1'), yrow(B, '每 MW 年持有成本（百萬美元）', 'X02', 'n2'),
                              yrow(B, '回本比值', 'X05', 'n2', 'eq'), yrow(B, '企業端（API）比值', 'X14', 'n2'), yrow(B, '消費端比值', 'X15', 'n2')])
    return f'<section id="payback"><h2>回本比值（P×Q）</h2>{ch}{t}<p class="note">持有成本取 Tokenomics 快照（Hopper／GB200 世代組合）；看年度趨勢，不看絕對值。</p></section>'


# ── 8 對帳 ──
def recon(B):
    g = B.name('COORD_C0')
    trs = []
    for row in g:
        c = lambda k, f: (sp(*row[k], f) if isinstance(row[k][1], (int, float)) else (sp(*row[k], 't') if row[k][1] else '—'))  # noqa: E731
        trs.append(f'<tr><td>{c(0, "t")}</td><td>{c(1, "t")}</td><td>{c(2, "n1")}</td><td>{c(3, "n1")}</td><td>{c(4, "sp1")}</td>'
                   f'<td>{c(5, "n1")}</td><td>{c(6, "n1")}</td><td>{c(7, "sp1")}</td><td><b>{c(8, "t")}</b></td></tr>')
    sm = B.name('COORD_C0Sum')
    c0 = ('<h3>三層協調：C0（實績）公司實績 vs 層級隱含值</h3>'
          '<p class="note">個股側＝公司法定揭露（招股書、年度與中期業績公告）；層級側＝Tokenomics 每 GW 營收 × Wattonomics 該玩家 GW、Wattonomics 該玩家資本支出。'
          'Wattonomics 的玩家名單沒有 MiniMax（中國只有合計），所以層級側沒有值。兩讀數都超過門檻且同號才標「對帳中」。</p>'
          + tbl(['項目', '單位', '公司實績（主）', '層級隱含（主）', '差距', '公司實績（第二）', '層級隱含（第二）', '第二讀數差距', '判定'], trs, 'tp')
          + f'<p class="note">判定摘要：可判定 {sp(*sm[0][0], "n0")} 項、對帳中 {sp(*sm[1][0], "n0")} 項、時差邊界 {sp(*sm[2][0], "n0")} 項、'
            f'參考（同源校準） {sp(*sm[3][0], "n0")} 項、無可比實績 {sp(*sm[4][0], "n0")} 項。</p>'
          + f'<p class="note">{sp(*B.name("COORD_C0Verdict")[0][0], "t")}</p>')
    cw = B.wb['Coord']
    rl = [f'<tr><td>{sp(f"Coord!A{r}", cw.cell(r, 1).value, "t")}</td><td>{sp(f"Coord!C{r}", cw.cell(r, 3).value, "t" if isinstance(cw.cell(r, 3).value, str) else "n0")}</td></tr>'
          for r in range(1, cw.max_row + 1) if cw.cell(r, 1).value in ('Tokenomics（TK_Link 快照）', 'Wattonomics', 'Wattonomics 是否讀本模型', 'Siliconomics')]
    read = ('<h3>讀取層級（同源，不判定）</h3><p class="note">本模型只讀 Tokenomics 快照（算力物理與持有成本）；不讀 Wattonomics，Wattonomics 也不讀本模型；'
            'Siliconomics 不適用（算力買方）。自層級讀入的值只列在 Coord 頁讀取層級節，不進 C0。</p>'
            '<details class="more"><summary>展開：讀取層級明細（版本、檔名）</summary>' + tbl(['層級', '本模型讀取'], rl, 'rd') + '</details>')
    cons = tbl(['對照市場（只對照，不校準）'] + YH, [yrow(B, '本模型總營收', 'R09', 'n0'), yrow(B, '賣方預估營收（共識／Goldman／Visible Alpha）', 'R13', 'n0'),
                                                 yrow(B, '本模型 ÷ 賣方預估（倍）', 'R14', 'n2'), yrow(B, 'Bloomberg 共識', 'R15', 'n0'), yrow(B, '本模型 ÷ Bloomberg 共識（倍）', 'R16', 'n2')])
    return (f'<section id="recon"><h2>對帳</h2>{c0}{read}'
            f'<details class="more"><summary>展開：模型對市場預估（賣方共識、Goldman、Visible Alpha）</summary>{cons}</details></section>')


# ── 9 資料狀態 ──
def status(B, xlsx_name, ver):
    tk = B.at('TK_Link', 'A2')
    return ('<section id="status"><h2>資料狀態</h2>'
            f'<dl class="kv"><div><dt>版本</dt><dd>{html.escape(ver)}（{html.escape(xlsx_name)}）</dd></div>'
            f'<div><dt>Tokenomics 快照</dt><dd>{sp(*tk, "t")}</dd></div>'
            '<div><dt>公司資料</dt><dd>招股書（2025-12-31）、FY2025 年度業績公告（2026-03-02）、2026 中期業績公告（2026-08-26）；SRC_MM 頁逐筆列出處與一手／二手。</dd></div>'
            '<div><dt>口徑</dt><dd>美元；曆年制（公司財年＝曆年）；MW＝IT 關鍵電力。</dd></div></dl>'
            '<h3>附錄</h3><details class="hist"><summary>版本要點</summary><p class="note">'
            'v0.2.6：Tokenomics 快照 v19.0 → v22.0（外部稽核優化 W3-5）；算力與成本的中間列改變，加權目標價與評等不變。v0.2.5：現價與同業報價改 2026-11-20 收盤（工作單 px1120）；其餘數值不變。v0.2.4：現價與同業報價改 2026-11-13 收盤（工作單 px1113）。v0.2.3：估值頁 Val 與反向 DCF 頁 RevDCF（目標價、評等、收支關卡）；既有收支數值不變。v0.2.2：版面 v1.3（新增 PL 標準損益表、Coord 三層協調：讀取層級與 C0 實績）；既有數值不變。v0.2.1：現價改 2026-11-05 收盤。'
            'v0.2：P×Q 回本比值、M3.1 降價、合約價證據。v0.1：首版。</p></details>'
            '<details class="hist"><summary>已知限制</summary><p class="note">季度層未建；自建集群規模找不到；2026 消費端付費人數與 ARPPU 缺一手資料；'
            'Tokenomics 缺中國晶片世代、租價與中國 API 有效牌價；同業只有智譜一家（其餘候選無同日可查證市值）；選擇權與受限制股份稀釋未計。</p></details></section>')


def main(xlsx, out):
    B = X(xlsx)
    ver = re.search(r'v\d+\.\d+\.\d+', B.wb['README']['A1'].value).group(0)
    parts = [thesis(B), valuation(B), sens(B), revdcf(B), pl(B), drivers(B), mod_unit(B), mod_funding(B), mod_payback(B), recon(B),
             status(B, os.path.basename(xlsx), ver)]
    navs = [('thesis', '命題與核心問題'), ('valuation', '估值'), ('sens', '敏感度'), ('revdcf', '反向 DCF'), ('pl', '損益'), ('drivers', '驅動因子'),
            ('unit', '算力與每 MW 收支'), ('funding', '資金'), ('payback', '回本比值'), ('recon', '對帳'), ('status', '資料狀態')]
    nav = ''.join(f'<a href="#{a}">{t}</a>' for a, t in navs)
    css = M.CSS_EXTRA + CSS_V13 + ('table{border-collapse:collapse;width:100%;font-size:13.5px}th,td{border-bottom:1px solid var(--line);padding:4px 6px;text-align:right}'
                                   'th:first-child,td:first-child{text-align:left}.tw{overflow-x:auto}tr.eq td{font-weight:600}'
                                   'table.tp td:nth-child(9){text-align:left}')
    doc = (f'<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
           f'<title>MiniMax 個股模型 {ver}</title><style>{css}</style></head><body><div class="wrap">'
           f'<header class="top"><h1>MiniMax 個股模型</h1><div class="meta">00100.HK（上海稀宇科技）｜版本 {ver}｜目標價 DCF＋同業倍數各半</div>'
           f'<nav class="toc">{nav}</nav></header>{hero(B)}<main class="v2">{"".join(parts)}</main>'
           f'<footer class="v2">MiniMax 個股模型 {ver}｜版面 v1.3｜minimax/build_html_v13.py｜xlsx <code>{html.escape(os.path.basename(xlsx))}</code></footer>'
           f'</div></body></html>')
    open(out, 'w', encoding='utf-8').write(doc)
    print(out)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
