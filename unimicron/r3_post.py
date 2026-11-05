# 第三輪修正（合併前，2026-11-06）共用後處理：建置腳本在最後存檔前呼叫 apply(wb, R3)。
# 公司資料夾各放同一份（公司自有，不屬於 tools/；tools/fcff_dcf.py v2 不變）。項目編號對照第三輪修正清單：
#   23 情境名稱統一「情境 A（順利）／B（平平）／C（不如預期）」；15／39 情境驅動對照收斂法列名中性、% 與 pt 格式化
#   24 Val「上檔空間級距」列名；14 stub 期說明依實際口徑；16 g 受 rf 上限時 Sens 列名
#   22／40 README 舊「重點觀察」改為公式（引用展示層「重點 1–3」），其餘舊重點觀察刪除
#   35／26 現價與 Comparable 同業前瞻 P/E 統一為 2026-11-04 收盤（同業 P/E＝原 P/E × 11-04 收盤 ÷ 原報價日收盤；前瞻 EPS 視為不變）
#   30 工具字樣改「未取得」
import re

YH = "Yahoo Finance 日線收盤 2026-11-04"
SCN_REPL = (("情境 A（樂觀）", "情境 A（順利）"), ("情境 B（基準）", "情境 B（平平）"), ("情境 C（保守）", "情境 C（不如預期）"),
            ("情境 C（悲觀）", "情境 C（不如預期）"), ("情境 C（下檔）", "情境 C（不如預期）"),
            ("情境 A 隱含價（樂觀）", "情境 A 隱含價（順利）"), ("情境 B 隱含價（基準）", "情境 B 隱含價（平平）"), ("情境 C 隱含價（下檔）", "情境 C 隱含價（不如預期）"),
            ("情境權重 A（樂觀）", "情境權重 A（順利）"), ("情境權重 B（基準）", "情境權重 B（平平）"), ("情境權重 C（保守）", "情境權重 C（不如預期）"),
            ("情境權重 C（悲觀）", "情境權重 C（不如預期）"))
FADE_LAB = {"scn_fade_g": "收斂法：各情境 FY2032 營收年增目標（延伸期收斂用）",
            "scn_fade_m": "收斂法：各情境 FY2032 淨利率目標（延伸期收斂用）"}
BAD = ("本環境", "讀不到", "抓取單", "取不到", "已抓取", "建置環境", "內建瀏覽器", "/home/claude", "原樣取自", "無法開啟")
TOOL_REPL = (('投資建議維持"&', '投資建議為"&'), ("Damodaran 2026-07 更新檔為 xlsx，本環境讀不到，沿用 2026-01 表", "Damodaran 2026-07 更新檔未取得，沿用 2026-01 表"),)


def _rows(ws, prefix, col=1):
    return [r for r in range(1, ws.max_row + 1) if str(ws.cell(r, col).value or "").startswith(prefix)]


def _each_str(wb):
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str):
                    yield ws, c


def replace_all(wb, pairs, must=()):
    hit = {a: 0 for a, _ in pairs}
    for ws, c in _each_str(wb):
        v = c.value
        for a, b in pairs:
            if a in v:
                v = v.replace(a, b); hit[a] += 1
        c.value = v
    for a in must:
        assert hit.get(a), ("未命中", a[:60])
    return hit


def peers(wb, PX, OLD, log):
    V = wb["Val"]
    h = _rows(V, "0.2 Comparable")
    assert len(h) == 1, h
    out = []
    for r in range(h[0] + 2, h[0] + 16):
        t, pe, d = V.cell(r, 2).value, V.cell(r, 3).value, V.cell(r, 4).value
        if isinstance(t, str) and isinstance(pe, (int, float)) and isinstance(d, str) and re.match(r"\d{4}-\d{2}-\d{2}$", d):
            p0, p1 = OLD[(t, d)], PX[t]
            V.cell(r, 3).value = round(pe * p1 / p0, 2)
            V.cell(r, 4).value = "2026-11-04"
            V.cell(r, 5).value = (f"stockanalysis.com Forward PE {pe}（{d}，當日收盤 {p0:,}）× {YH} {p1:,} ÷ {p0:,}"
                                  "（第三輪修正：換算到 2026-11-04，分析師共識前瞻 EPS 視為不變）")
            out.append((t, pe, d, V.cell(r, 3).value))
    assert out
    log.append("同業：" + "；".join(f"{t} {a}（{d}）→ {b}" for t, a, d, b in out))
    return out


def px(wb, tk, PX, src2, log):
    IN, V = wb["In"], wb["Val"]
    r = [x for x in range(1, IN.max_row + 1) if IN.cell(x, 1).value == "px"]
    assert len(r) == 1
    r = r[0]
    old, lab = IN.cell(r, 5).value, str(IN.cell(r, 2).value)
    m = re.search(r"\d{4}-\d{2}-\d{2} 收盤", lab)
    assert m, lab
    old_date = m.group(0)[:10]
    IN.cell(r, 2).value = re.sub(r"（.*?）", "（2026-11-04 收盤" + ("，分割後" if "分割後" in lab else "") + "）", lab, count=1)
    IN.cell(r, 5).value = PX[tk]
    IN.cell(r, 8).value = f"{YH}：{PX[tk]:,}（{tk}）"
    IN.cell(r, 9).value = f"第三輪修正（2026-11-06）：現價統一為 2026-11-04 收盤；原 {old_date} 收盤 {old:,}（{src2}）"
    vr = [x for x in range(1, V.max_row + 1) if str(V.cell(x, 1).value or "").startswith("現價（")]
    old_lab = V.cell(vr[0], 1).value
    new_lab = re.sub(r"（.*?）", "（2026-11-04 收盤" + ("，分割後" if "分割後" in old_lab else "") + "）", old_lab, count=1)
    replace_all(wb, ((old_lab, new_lab),))
    pat = re.compile(r"px（\d{4}-\d{2}-\d{2} 收盤[^）]*）")
    for ws, c in _each_str(wb):
        if ws.title != "README" and not (ws.title == "In" and c.row == r) and pat.search(c.value):
            c.value = pat.sub(lambda m: "px（2026-11-04 收盤" + ("，分割後" if "分割後" in m.group(0) else "") + "）", c.value)
    log.append(f"現價 {old:,}（{old_date}）→ {PX[tk]:,}（2026-11-04）")
    return old, old_date


def own_fpe(wb, key, p0, d0, p1, log):
    """本公司前瞻 P/E 換算到 2026-11-04（前瞻 EPS 不變）；cons_feps 改為 px ÷ fpe（與原報價日價格 ÷ 原 P/E 相同）"""
    IN = wb["In"]
    k = {IN.cell(x, 1).value: x for x in range(1, IN.max_row + 1) if IN.cell(x, 1).value}
    r = k[key]; v = IN.cell(r, 5).value
    assert isinstance(v, (int, float)), v
    IN.cell(r, 5).value = round(v * p1 / p0, 2)
    IN.cell(r, 2).value = re.sub(r"（.*?）$", "", str(IN.cell(r, 2).value)) + "（stockanalysis；換算到 2026-11-04 收盤）"
    IN.cell(r, 9).value = f"第三輪修正：原 {v}（{d0}，當日價格 {p0:,}）× 2026-11-04 收盤 {p1:,} ÷ {p0:,}（共識前瞻 EPS 視為不變）"
    if "cons_feps" in k:
        c = IN.cell(k["cons_feps"], 5)
        if isinstance(c.value, str) and re.match(rf"^=([\d.]+|px)/{key}$", c.value):
            c.value = f"=px/{key}"
            IN.cell(k["cons_feps"], 2).value = "共識前瞻 12 月 EPS（現價 ÷ 前瞻 P/E；兩者同為 2026-11-04 口徑）"
    log.append(f"本公司前瞻 P/E {v} → {IN.cell(r, 5).value}")


def apply(wb, R3):
    log = []
    replace_all(wb, SCN_REPL + TOOL_REPL + tuple(R3.get("repl", ())), must=tuple(a for a, _ in R3.get("repl", ())))
    IN, V, SE, RD = wb["In"], wb["Val"], wb["Sens"], wb["README"]
    # 15／39：情境驅動對照
    s0 = [r for r in range(1, IN.max_row + 1) if str(IN.cell(r, 1).value or "").lstrip("J.Z. ").startswith("情境驅動對照")
          or re.match(r"^[A-Z]\d?\.\s*情境驅動對照", str(IN.cell(r, 1).value or ""))]
    assert len(s0) == 1, s0
    r = s0[0] + 2
    while r <= IN.max_row and str(IN.cell(r, 1).value or "").startswith("scn_"):
        k, u = IN.cell(r, 1).value, IN.cell(r, 3).value
        if k in FADE_LAB:
            IN.cell(r, 2).value = FADE_LAB[k]
        elif k.startswith("scn_fade") and isinstance(IN.cell(r, 2).value, str):
            IN.cell(r, 2).value = "收斂法：" + re.sub(r"（[^）]*(tp_g|十年|情境|FY2029)[^）]*）", "", IN.cell(r, 2).value) + "（延伸期收斂用）"
        if u in ("%", "pt", "%/年"):
            for cc in (4, 5, 6):
                IN.cell(r, cc).number_format = "0.0%"
        elif u == "x":
            for cc in (4, 5, 6):
                IN.cell(r, cc).number_format = "0.00"
        r += 1
    # 24
    for rr in _rows(V, "上檔空間級距"):
        if V.cell(rr, 2).value is not None:
            V.cell(rr, 1).value = "上檔空間級距（六級，供分組）"
            V.cell(rr, 5).value = "以「上檔空間：加權目標價」分六級（Andy 2026-10-25 14:00），只供分組；投資建議用下一列的五級固定規則"
    # 14：stub 口徑
    st = R3["stub"]   # 例 "含 stub 期的會計年度 FY2026 FCFF × stub_frac"
    rr = _rows(V, "＋Stub 期 FCFF")
    assert len(rr) == 1
    V.cell(rr[0], 5).value = '="資產負債表日 "&bs_date&"；stub_frac "&TEXT(stub_frac,"0.000")&" 年（' + st + '）"'
    k = [x for x in range(1, IN.max_row + 1) if IN.cell(x, 1).value == "stub_frac"]
    IN.cell(k[0], 9).value = "stub FCFF＝" + st
    D = wb["DCF"]
    for x in range(1, D.max_row + 1):
        if str(D.cell(x, 2).value or "").startswith("stub FCFF＝"):
            D.cell(x, 2).value = "stub FCFF＝" + st
    # 16
    e = {IN.cell(x, 1).value: IN.cell(x, 5).value for x in range(1, IN.max_row + 1)}
    if isinstance(e.get("rf"), (int, float)) and isinstance(e.get("g_cap"), (int, float)) and e["rf"] < e["g_cap"]:
        g = [x for x in range(1, SE.max_row + 1) if SE.cell(x, 1).value == "永續成長率 g ±1pt"]
        assert len(g) == 1, g
        SE.cell(g[0], 1).value = "永續成長率 g −1pt／上限 rf"
        log.append("g 上限列名")
    # 35／26
    if R3.get("px"):
        px(wb, R3["px"], R3["PX"], R3["px_src"], log)
        peers(wb, R3["PX"], R3["OLD"], log)
        if R3.get("fpe"):
            own_fpe(wb, *R3["fpe"], R3["PX"][R3["px"]], log)
    # 22：README 舊重點觀察 → 公式引用展示層「重點 1–3」
    zr = {}
    for x in range(1, RD.max_row + 1):
        v = RD.cell(x, 1).value
        m = re.match(r'^="?重點 (\d)：', str(v or ""))
        if m and int(m.group(1)) not in zr:
            zr[int(m.group(1))] = x
    assert all(i in zr for i in (1, 2, 3)), zr
    n = 0
    for x in range(1, RD.max_row + 1):
        v = str(RD.cell(x, 1).value or "")
        m = re.match(r'^(="?)?重點觀察 (\d)：', v)
        if m:
            i = int(m.group(2))
            RD.cell(x, 1).value = f'="重點觀察 {i}："&MID(A{zr[i]},6,2000)' if i in zr and i <= 3 else None
            n += 1
    log.append(f"README 舊重點觀察 {n} 列改公式／刪除")
    # 30：工具字樣
    for ws, c in _each_str(wb):
        for b in BAD:
            assert b not in c.value, (ws.title, c.coordinate, c.value[:120])
    for f in R3.get("extra", ()):
        f(wb, log)
    print("r3_post：" + "；".join(log))
    return log
