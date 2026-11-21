# 2026-11-21 現價與同業報價改 2026-11-20 收盤（工作單 docs/workorders/20261121_px1120_現價與同業報價換11-20收盤.md；前例 px1105_post.py）。
# 共用後處理：建置腳本在最後一個後處理之後、存檔前呼叫 apply(wb, P)（公司資料夾各放同一份，公司自有，不屬於 tools/）。
#   價格只取自 markets/invest_map/data/closes_20261109_20261120.csv（Yahoo Finance 日線收盤），不臆造；CSV 沒有的同業維持原值並列入 UNCHANGED。
#   In 頁 px＝2026-11-20 收盤；Val 0.2 Comparable 同業前瞻 P/E＝現行倍數 × 11-20 收盤 ÷ 該列報價日（Val D 欄）收盤（共識 EPS 不變）；
#   本公司前瞻 P/E 同法換算。rf、共識 EPS、預測用匯率、模型假設不改。
import csv, os, re

HERE = os.path.dirname(os.path.abspath(__file__))
DD = os.path.join(HERE, "..", "markets", "invest_map", "data")
CSVS = ["closes_20261109_20261120.csv", "closes_20261020_20261113.csv", "closes_extra_20261113.csv"]
NEW = "2026-11-20"
YH = "Yahoo Finance 日線收盤 2026-11-20"
SRCTAG = "工作單 px1120；`closes_20261109_20261120.csv`"
UNCHANGED = []
_C = {}


def close(sym, d):
    if not _C:
        for f in CSVS:   # 新檔優先；舊檔只補沒有的（代號, 日期）
            for r in csv.DictReader(open(os.path.join(DD, f), encoding="utf-8")):
                _C.setdefault(r["symbol"], {}).setdefault(r["date"], round(float(r["close"]), 4))
    return _C.get(sym, {}).get(d)


def num(x):
    return int(x) if float(x) == int(x) else x


def fmt(x):
    return f"{x:,}"


def _each_str(wb):
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str):
                    yield ws, c


def _is_hist(ws, c):
    v = c.value
    if ws.title == "README" and (c.coordinate == "A1" or re.match(r"^v0\.\d", v)):
        return True
    return "第三輪修正" in v or "第三輪（合併前）" in v


def px(wb, P, log, done):
    IN = wb["In"]
    tk = P["px"]
    r = [x for x in range(1, IN.max_row + 1) if IN.cell(x, 1).value == "px"]
    assert len(r) == 1
    r = r[0]
    old, lab = IN.cell(r, 5).value, str(IN.cell(r, 2).value)
    d0 = re.search(r"2026-\d\d-\d\d", lab).group(0)
    c0, c1 = close(tk, d0), close(tk, NEW)
    assert c0 is not None and c1 is not None, (tk, d0)
    assert isinstance(old, (int, float)) and abs(old - c0) < 0.006, (old, c0, tk, d0)
    IN.cell(r, 2).value = lab.replace(d0, NEW)
    IN.cell(r, 5).value = num(c1)
    IN.cell(r, 8).value = f"{YH}：{fmt(num(c1))}（{tk}）"
    IN.cell(r, 9).value = (f"2026-11-21：現價改 2026-11-20 收盤（{SRCTAG}）；{P['prev']} 為 {d0} 收盤 {fmt(old)}（Yahoo Finance 日線收盤 {d0}）；前一版備註：{IN.cell(r, 9).value}")
    for j in (2, 8, 9):
        done.add(("In", IN.cell(r, j).coordinate))
    log.append(f"現價 {fmt(old)} → {fmt(num(c1))}（{d0} → {NEW}）")
    P["_old_px"], P["_new_px"] = old, c1
    return old


def peers(wb, P, log, done):
    V = wb["Val"]
    h = [r for r in range(1, V.max_row + 1) if str(V.cell(r, 1).value or "").startswith("0.2 Comparable")]
    assert len(h) == 1, h
    out = []
    for r in range(h[0] + 2, h[0] + 20):
        t, pe, d = V.cell(r, 2).value, V.cell(r, 3).value, V.cell(r, 4).value
        if str(V.cell(r, 1).value or "").startswith("同業前瞻 P/E 中位數"):
            break
        if isinstance(t, str) and isinstance(pe, (int, float)) and isinstance(d, str) and re.fullmatch(r"2026-\d\d-\d\d", d):
            p0, p1 = close(t, d), close(t, NEW)
            if p0 is None or p1 is None:
                UNCHANGED.append(t); out.append((t, pe, pe, "未換價")); continue
            V.cell(r, 3).value = round(pe * p1 / p0, 2)
            V.cell(r, 4).value = NEW
            V.cell(r, 5).value = (f"現行前瞻 P/E {pe}（{d} 口徑）× {YH} {fmt(num(p1))} ÷ {d} 收盤 {fmt(num(p0))}"
                                  f"（共識前瞻 EPS 不變；{SRCTAG}）；前一版來源：{V.cell(r, 5).value}")
            for j in (3, 4, 5):
                done.add(("Val", V.cell(r, j).coordinate))
            out.append((t, pe, V.cell(r, 3).value))
    assert out
    log.append("同業：" + "；".join(f"{o[0]} {o[1]} → {o[2]}" for o in out))


def own_fpe(wb, P, log, done):
    IN = wb["In"]
    k = {IN.cell(x, 1).value: x for x in range(1, IN.max_row + 1) if IN.cell(x, 1).value}
    r = k[P["fpe"]]; v = IN.cell(r, 5).value
    assert isinstance(v, (int, float)), v
    p0, p1 = P["_old_px"], P["_new_px"]
    IN.cell(r, 5).value = round(v * p1 / p0, 2)
    IN.cell(r, 2).value = re.sub(P["d0"], NEW, str(IN.cell(r, 2).value))
    IN.cell(r, 9).value = f"2026-11-21：{v}（{P['d0']} 口徑）× 2026-11-20 收盤 {fmt(num(p1))} ÷ {fmt(p0)}（共識前瞻 EPS 不變；{SRCTAG}）；前一版：{IN.cell(r, 9).value}"
    if "cons_feps" in k:
        IN.cell(k["cons_feps"], 2).value = str(IN.cell(k["cons_feps"], 2).value).replace(P["d0"], NEW)
        done.add(("In", IN.cell(k["cons_feps"], 2).coordinate))
    for j in (2, 9):
        done.add(("In", IN.cell(r, j).coordinate))
    log.append(f"本公司前瞻 P/E {v} → {IN.cell(r, 5).value}")


def refs(wb, P, log, done):
    """Val 「並列參考」列（不進目標價的他國同業倍數，文字內 N.Nx）：P["refs"]＝[(代號, 報價日 d0 的倍數)]，依文字內出現順序；× 11-20 收盤 ÷ d0 收盤"""
    V = wb["Val"]
    r = next(x for x in range(1, V.max_row + 1) if str(V.cell(x, 1).value or "").startswith("並列參考"))
    old = V.cell(r, 1).value
    head = old.split(f"（{P['d0']}", 1)[0]
    nums = list(re.finditer(r"\d+\.\d+x", head))
    assert len(nums) == len(P["refs"]), (head, P["refs"])
    new_head, last, vs = "", 0, []
    for m, (t, v0) in zip(nums, P["refs"]):
        p0, p1 = close(t, P["d0"]), close(t, NEW)
        v1 = v0 * p1 / p0
        vs.append((t, v0, v1))
        new_head += head[last:m.start()] + f"{v1:.1f}x"; last = m.end()
    new_head += head[last:]
    V.cell(r, 1).value = (new_head + f"（{NEW}；" + "、".join(f"{t} {v0:g}x → {v1:.1f}x" for t, v0, v1 in vs)
                          + f"（{P['d0']} 口徑 × {NEW} 收盤 ÷ {P['d0']} 收盤；Yahoo Finance 日線收盤；{SRCTAG}）；前一版：{old}")
    done.add(("Val", V.cell(r, 1).coordinate))
    log.append("並列參考：" + "；".join(f"{t} {v0:g}x → {v1:.1f}x" for t, v0, v1 in vs))


def dates(wb, P, log, done):
    """非歷史說明格內的現價日期字樣 2026-11-05 收盤 → 2026-11-20 收盤（rf 等不在範圍的日期不動）"""
    d0 = P["d0"]; s0, s1 = d0[5:], NEW[5:]
    pats = ((f"{d0} 收盤", f"{NEW} 收盤"), (f"{s0} 收盤", f"{s1} 收盤"), (f"（{d0}）", f"（{NEW}）"),
            (f"統一為 {NEW} 收盤（2026-11-06 更新）", f"統一為 {NEW} 收盤（2026-11-21 更新）")) + tuple(P.get("dates_extra", ()))
    n = 0
    for ws, c in _each_str(wb):
        if (ws.title, c.coordinate) in done:
            continue
        if _is_hist(ws, c) or re.search(rf"(原|改|前一版[:：]?|前一版來源[:：]?) ?{d0}", c.value) or "公債" in c.value or str(ws.cell(c.row, 1).value) == "rf":
            continue
        v = c.value
        for a, b in pats:
            v = v.replace(a, b)
        if v != c.value:
            c.value = v; n += 1
    log.append(f"日期字樣 {n} 格")


def ver_row(wb, P, log, done):
    """README「v1.3 展示層」節的本版要點列（r13＋9，HTML 取用）：舊文複製到 README 末（沿革保留原文），本列改本版說明"""
    RD = wb["README"]
    r13 = [r for r in range(1, RD.max_row + 1) if str(RD.cell(r, 1).value or "").strip().startswith("v1.3 展示層")]
    if not r13:
        return   # 敏感度快照或 Coord 另一側建置不含 v1.3 後處理
    row = r13[-1] + 9
    v = RD.cell(row, 1).value
    assert isinstance(v, str) and v.startswith(P.get("vprev", P["prev"])), (row, str(v)[:60])
    RD.cell(RD.max_row + 1, 1, v)
    RD.cell(row, 1).value = (f"{P['ver']}（2026-11-21）：現價與同業報價改 2026-11-20 收盤（工作單 px1120）。只換價格：現價、Comparable 同業前瞻 P/E（現行倍數 × 11-20 收盤 ÷ 原報價日收盤）、"
                             f"本公司前瞻 P/E 隨之換算；rf、共識 EPS、預測用匯率、模型假設、公式與架構不變，加權目標價、上檔空間與評等隨公式重算（前後數字見 README 本版要點與成品）。承 {P['prev']}。")
    log.append(f"本版要點列改 {P['ver']}")


def version(wb, P, log):
    RD = wb["README"]
    t = f"{P['ver']}（2026-11-21）：現價與同業報價改 2026-11-20 收盤（工作單 px1120）"
    RD.cell(RD.max_row + 1, 1, t)
    a1 = RD["A1"].value
    a2 = re.sub(r"(模型 v)[\d.]+(（)", lambda m: m.group(1) + P["ver"][1:] + m.group(2), a1, count=1)
    assert a2 != a1, a1[:100]
    RD["A1"] = a2 + "；" + t


def leftovers(wb, P):
    return [(ws.title, c.coordinate, c.value) for ws, c in _each_str(wb)
            if (P["d0"][5:] in c.value) and re.search("收盤|現價|股價|報價", c.value) and not _is_hist(ws, c)]


def apply(wb, P):
    log = []; done = set()
    UNCHANGED.clear()
    if P.get("px"):
        px(wb, P, log, done)
    if P.get("peers", True):
        peers(wb, P, log, done)
    if P.get("fpe"):
        own_fpe(wb, P, log, done)
    if P.get("refs"):
        refs(wb, P, log, done)
    if P.get("vrow", True):
        ver_row(wb, P, log, done)
    for f in P.get("extra", ()):
        f(wb, log, done)
    dates(wb, P, log, done)
    version(wb, P, log)
    if UNCHANGED:
        log.append("未換價同業：" + "、".join(UNCHANGED))
    print("px1120_post：" + "；".join(log))
    if os.environ.get("PX1120_REPORT"):
        with open(os.environ["PX1120_REPORT"], "w") as fh:
            for x in leftovers(wb, P):
                fh.write("DATE\t%s\t%s\t%s\n" % (x[0], x[1], x[2][:300].replace("\n", " ")))
    return log
