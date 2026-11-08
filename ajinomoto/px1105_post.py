# 2026-11-06 現價與同業報價改 2026-11-05 收盤；情境名稱改回「情境 A（樂觀）／B（基準）／C（保守）」（Andy 2026-11-06 07:28、03:14）。
# 共用後處理：建置腳本在 r3_post.apply 之後、存檔前呼叫 apply(wb, P5)。公司資料夾各放同一份（公司自有，不屬於 tools/；tools/fcff_dcf.py 不變）。
#   價格：In 頁 px＝2026-11-05 收盤；Val 0.2 Comparable 同業前瞻 P/E＝現行倍數（2026-11-04 口徑）× 11-05 收盤 ÷ 11-04 收盤（共識 EPS 不變）；
#         本公司前瞻 P/E 同法換算。來源欄寫「Yahoo Finance 日線收盤 2026-11-05」。rf、共識 EPS、預測用匯率、模型假設不改。
#   情境名稱：取代第三輪修正第 23 點；「平平」作一般形容詞時不改（只換「B（平平）」「B 平平」「B＝平平」「B：平平」等情境用法）。
import re

YH5 = "Yahoo Finance 日線收盤 2026-11-05"
YH4 = "Yahoo Finance 日線收盤 2026-11-04"
D4, D5 = "2026-11-04", "2026-11-05"
SCN = (("（順利）", "（樂觀）"), ("（平平）", "（基準）"), ("（不如預期）", "（保守）"),
       ("A 順利", "A 樂觀"), ("B 平平", "B 基準"), ("C 不如預期", "C 保守"),
       ("A＝順利", "A＝樂觀"), ("B＝平平", "B＝基準"), ("C＝不如預期", "C＝保守"),
       ("A：順利", "A：樂觀"), ("B：平平", "B：基準"), ("C：不如預期", "C：保守"),
       ("不如預期～順利", "保守～樂觀"))


def _each_str(wb):
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str):
                    yield ws, c


def rename_text(s):
    for a, b in SCN:
        s = s.replace(a, b)
    return s


def scenario_names(wb, log):
    n = 0
    for ws, c in _each_str(wb):
        v = rename_text(c.value)
        if v != c.value:
            c.value = v; n += 1
    left = [(ws.title, c.coordinate, c.value) for ws, c in _each_str(wb) if re.search("順利|不如預期|平平", c.value)]
    log.append(f"情境名稱 {n} 格")
    return left


def _is_hist(ws, c):
    """版本沿革與歷史說明不改日期：README 頁 A1 與 v0. 開頭列、含「第三輪修正」的說明格"""
    v = c.value
    if ws.title == "README" and (c.coordinate == "A1" or re.match(r"^v0\.\d", v)):
        return True
    return "第三輪修正" in v or "第三輪（合併前）" in v


def px(wb, P, log):
    IN, V = wb["In"], wb["Val"]
    tk = P["px"]
    r = [x for x in range(1, IN.max_row + 1) if IN.cell(x, 1).value == "px"]
    assert len(r) == 1
    r = r[0]
    old, lab = IN.cell(r, 5).value, str(IN.cell(r, 2).value)
    assert isinstance(old, (int, float)) and abs(old - P["PX4"][tk]) < 1e-9, (old, P["PX4"][tk])
    assert D4 in lab or D5 in lab, lab
    IN.cell(r, 2).value = lab.replace(D4, D5)
    IN.cell(r, 5).value = P["PX5"][tk]
    IN.cell(r, 8).value = f"{YH5}：{P['PX5'][tk]:,}（{tk}）"
    IN.cell(r, 9).value = f"2026-11-06：現價改 2026-11-05 收盤；原 2026-11-04 收盤 {old:,}（{YH4}）"
    log.append(f"現價 {old:,} → {P['PX5'][tk]:,}")
    return old


def peers(wb, P, log):
    V = wb["Val"]
    h = [r for r in range(1, V.max_row + 1) if str(V.cell(r, 1).value or "").startswith("0.2 Comparable")]
    assert len(h) == 1, h
    out = []
    for r in range(h[0] + 2, h[0] + 16):
        t, pe, d = V.cell(r, 2).value, V.cell(r, 3).value, V.cell(r, 4).value
        if isinstance(t, str) and isinstance(pe, (int, float)) and d == D4:
            p4, p5 = P["PX4"][t], P["PX5"][t]
            V.cell(r, 3).value = round(pe * p5 / p4, 2)
            V.cell(r, 4).value = D5
            V.cell(r, 5).value = (f"現行前瞻 P/E {pe}（2026-11-04 口徑）× {YH5} {p5:,} ÷ 2026-11-04 收盤 {p4:,}"
                                  f"（共識前瞻 EPS 不變）；前一版來源：{V.cell(r, 5).value}")
            out.append((t, pe, V.cell(r, 3).value))
    assert out
    log.append("同業：" + "；".join(f"{t} {a} → {b}" for t, a, b in out))
    return out


def own_fpe(wb, P, log):
    IN = wb["In"]
    k = {IN.cell(x, 1).value: x for x in range(1, IN.max_row + 1) if IN.cell(x, 1).value}
    r = k[P["fpe"]]; v = IN.cell(r, 5).value
    assert isinstance(v, (int, float)), v
    p4, p5 = P["PX4"][P["px"]], P["PX5"][P["px"]]
    IN.cell(r, 5).value = round(v * p5 / p4, 2)
    IN.cell(r, 2).value = str(IN.cell(r, 2).value).replace(D4, D5)
    IN.cell(r, 9).value = f"2026-11-06：{v}（2026-11-04 口徑）× 2026-11-05 收盤 {p5:,} ÷ {p4:,}（共識前瞻 EPS 不變）；前一版：{IN.cell(r, 9).value}"
    if "cons_feps" in k:
        IN.cell(k["cons_feps"], 2).value = str(IN.cell(k["cons_feps"], 2).value).replace(D4, D5)
    log.append(f"本公司前瞻 P/E {v} → {IN.cell(r, 5).value}")


def dates(wb, P, log):
    """非歷史說明格內的現價日期字樣 2026-11-04 收盤 → 2026-11-05 收盤（rf 等不在範圍的日期不動）"""
    pats = (("2026-11-04 收盤", "2026-11-05 收盤"), ("11-04 收盤", "11-05 收盤"), ("（2026-11-04）", "（2026-11-05）"))
    force = (("報價統一為 2026-11-04 收盤（第三輪修正）", "報價統一為 2026-11-05 收盤（2026-11-06 更新）"),) + tuple(P.get("dates_force", ()))
    n = 0
    for ws, c in _each_str(wb):
        for a, b in force:
            if a in c.value:
                c.value = c.value.replace(a, b); n += 1
        if _is_hist(ws, c) or re.search(r"(為|原|前一版[:：]?) ?2026-11-04", c.value) or "公債" in c.value or "rf" == str(ws.cell(c.row, 1).value):
            continue
        v = c.value
        for a, b in pats + tuple(P.get("dates_extra", ())):
            v = v.replace(a, b)
        if v != c.value:
            c.value = v; n += 1
    log.append(f"日期字樣 {n} 格")


def version(wb, P, log):
    RD = wb["README"]
    t = f"{P['ver']}：現價與同業報價改 2026-11-05 收盤；情境名稱改回樂觀／基準／保守"
    RD.cell(RD.max_row + 1, 1, t)
    RD["A1"] = RD["A1"].value + "；" + t


def leftovers(wb):
    return [(ws.title, c.coordinate, c.value) for ws, c in _each_str(wb)
            if ("11-04" in c.value) and not _is_hist(ws, c)]


def apply(wb, P):
    log = []
    left = scenario_names(wb, log)
    dates(wb, P, log)   # 先換說明格日期，再寫入本輪的新說明（避免把「原 2026-11-04 收盤」也換掉）
    if P.get("px"):
        px(wb, P, log)
    if P.get("peers", True):
        peers(wb, P, log)
    if P.get("fpe"):
        own_fpe(wb, P, log)
    for f in P.get("extra", ()):
        f(wb, log)
    version(wb, P, log)
    print("px1105_post：" + "；".join(log))
    import os
    if os.environ.get("PX1105_REPORT"):
        with open(os.environ["PX1105_REPORT"], "w") as fh:
            for x in left:
                fh.write("SCN\t%s\t%s\t%s\n" % (x[0], x[1], x[2][:300].replace("\n", " ")))
            for x in leftovers(wb):
                fh.write("DATE\t%s\t%s\t%s\n" % (x[0], x[1], x[2][:300].replace("\n", " ")))
    return log
