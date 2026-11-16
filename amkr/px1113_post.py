# 工作單 px1113（2026-11-16）：現價與同業報價改 2026-11-13 收盤（價格只取自 markets/invest_map/data/closes_20261020_20261113.csv）。
# 建置腳本在存檔前呼叫 apply(wb, here, cfg)。各公司資料夾放同一份（公司自有，不屬於 tools/）；只改價格輸入與相關說明文字，不改任何公式結構。
#   1. In 頁 px（現價）改 11-13 收盤；來源欄保留上一版值。
#   2. Val 頁同業前瞻倍數：新倍數＝現行倍數 × 11-13 收盤 ÷ 現行倍數所用報價日收盤（共識 EPS 不變）。只改公式內的價格項；
#      CSV 沒有的報價日（10-20 以前）或沒有的代號：維持原值並回報（cfg 的 unchanged）。
#   3. cfg["extra_subs"]：其他寫死在公式裡的股價（子公司市值等）逐筆換價，每筆必須命中。
#   4. README 版本沿革中以公式帶出「本版」數字的舊版列凍結為上一版成品的顯示文字（hist_frozen_px1113.json）。
#   5. README A1 與「本版要點」加本版說明（固定文字）。
import csv, json, os, re

NEWD = "2026-11-13"
OLDD = "2026-11-05"
_CSV = None


def closes():
    global _CSV
    if _CSV is None:
        p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "markets", "invest_map", "data", "closes_20261020_20261113.csv")
        _CSV = {(r["symbol"], r["date"]): float(r["close"]) for r in csv.DictReader(open(p, encoding="utf-8"))}
    return _CSV


def px(sym, d=NEWD):
    return closes()[(sym, d)]


def fmt(x):
    return "%.10g" % x


def n(x):
    return f"{x:,.10g}"


def gen_frozen(old_xlsx, out_json):
    """由上一版成品（LibreOffice 重算值）產生凍結文字：README 以 =\"v… 開頭且含 & 的公式列。"""
    import openpyxl
    f = openpyxl.load_workbook(old_xlsx)["README"]
    d = openpyxl.load_workbook(old_xlsx, data_only=True)["README"]
    out = {}
    for r in range(1, f.max_row + 1):
        v = f.cell(r, 1).value
        if isinstance(v, str) and re.match(r'^="v\d', v) and "&" in v:
            k = v[2:62]
            assert k not in out, k
            out[k] = d.cell(r, 1).value
    json.dump(out, open(out_json, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return len(out)


def apply(wb, here, cfg):
    """cfg：ver、old_ver、tk（本公司 CSV 代號）、old_date（現價現行日，預設 2026-11-05）、peer_map（代號覆寫）、extra_subs[(old,new)]、sweep、note（README 本版說明）。"""
    log, unchanged = [], []
    old_date = cfg.get("old_date", OLDD)
    INw, Vw, RD = wb["In"], wb["Val"], wb["README"]
    # 1 現價
    r = next(r for r in range(1, INw.max_row + 1) if INw.cell(r, 1).value == "px")
    old = INw.cell(r, 5).value
    assert abs(old - px(cfg["tk"], old_date)) < 1e-6 * abs(old), (cfg["tk"], old, px(cfg["tk"], old_date))
    new = px(cfg["tk"])
    INw.cell(r, 5).value = new
    INw.cell(r, 2).value = re.sub(r"\d{4}-\d{2}-\d{2} 收盤", f"{NEWD} 收盤", str(INw.cell(r, 2).value))
    prevH, prevI = str(INw.cell(r, 8).value or ""), str(INw.cell(r, 9).value or "")
    INw.cell(r, 8).value = f"Yahoo Finance 日線收盤 {NEWD}：{n(new)}（{cfg['tk']}；markets/invest_map/data/closes_20261020_20261113.csv；工作單 px1113）"
    INw.cell(r, 9).value = f"{cfg['old_ver']} 為 {old_date} 收盤 {n(old)}；" + prevI
    log.append(f"px {n(old)} → {n(new)}")
    # 2 同業倍數
    pm = cfg.get("peer_map", {})
    npeer = 0
    for rr in range(1, Vw.max_row + 1):
        t, f, d = Vw.cell(rr, 2).value, Vw.cell(rr, 3).value, Vw.cell(rr, 4).value
        if not (isinstance(t, str) and isinstance(d, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", d) and f is not None and not (isinstance(f, str) and f.startswith("=MEDIAN"))):
            continue
        k = pm.get(t) or re.split(r"[／/]", t)[-1].strip()
        if d != old_date or (k, NEWD) not in closes():
            unchanged.append((k, d)); continue
        p0, p1 = px(k, d), px(k)
        if isinstance(f, str) and f.startswith("="):
            m = re.match(r"^=([\d.]+)\*([\d.]+)/([\d.]+)$", f)            # =PE*N/O
            m2 = re.match(r"^=([\d.]+)/\(([\d.]+)/([\d.]+)\)$", f)        # =N/(O/PE)
            m3 = re.match(r"^=\(2\*([\d.]+)/([\d.]+)\)/(.*)$", f)         # =(2*N/FX)/(…)  ADS 換算
            if m:
                assert abs(float(m.group(2)) - p0) < 1e-4 * p0, (k, f, p0)
                new_f = f"={m.group(1)}*{fmt(p1)}/{m.group(3)}"
            elif m2:
                assert abs(float(m.group(1) if m else m2.group(1)) - p0) < 1e-4 * p0, (k, f, p0)
                new_f = f"={fmt(p1)}/({m2.group(2)}/{m2.group(3)})"
            elif m3:
                assert abs(float(m3.group(1)) - p0) < 1e-4 * p0, (k, f, p0)
                new_f = f"=(2*{fmt(p1)}/{m3.group(2)})/{m3.group(3)}"
            else:
                raise AssertionError(("未知同業公式", k, f))
        else:
            new_f = f"={fmt(float(f))}*{fmt(p1)}/{fmt(p0)}"
        Vw.cell(rr, 3).value = new_f
        Vw.cell(rr, 4).value = NEWD
        Vw.cell(rr, 5).value = str(Vw.cell(rr, 5).value or "") + f"；工作單 px1113：換算到 {NEWD}＝現行倍數 × {NEWD} 收盤 {n(p1)} ÷ {old_date} 收盤 {n(p0)}（Yahoo Finance 日線收盤；markets/invest_map/data/closes_20261020_20261113.csv；共識前瞻 EPS 不變）"
        npeer += 1
        log.append(f"peer {k} {f} → {new_f}")
    # 3 其他寫死股價
    for a, b in cfg.get("extra_subs", ()):
        hit = 0
        for ws in wb.worksheets:
            if ws.title == "README":
                continue
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value, str) and a in c.value:
                        c.value = c.value.replace(a, b); hit += 1
        assert hit, a
        log.append(f"extra {a} → {b}（{hit} 格）")
    # 4 文字：現價日期標籤（不動 README 版本沿革、Val 同業列、In px 列）
    hist = re.compile(r'^(="?)?v\d+\.\d')
    swept = []
    peer_rows = {rr for rr in range(1, Vw.max_row + 1) if isinstance(Vw.cell(rr, 4).value, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", Vw.cell(rr, 4).value)}
    for ws in wb.worksheets:
        if ws.title in ("Sources",):
            continue
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not isinstance(v, str) or hist.match(v):
                    continue
                if ws.title == "In" and c.row == r:
                    continue
                if ws.title == "README" and not v.startswith("="):
                    continue
                if ws.title == "Val" and c.row in peer_rows:
                    continue
                w = v
                for a, b in cfg.get("sweep", (("%s 收盤" % old_date, "%s 收盤" % NEWD),)):
                    w = w.replace(a, b)
                if v.startswith("=") and re.search(r"\bpx\b", v):
                    w = w.replace(f"（{old_date}）", f"（{NEWD}）")
                if w != v:
                    c.value = w; swept.append((ws.title, c.coordinate, v[:40]))
    log.append(f"文字日期 {len(swept)} 格：" + "；".join(f"{a}!{b}" for a, b, _ in swept))
    # 5 凍結 README 舊版公式列
    fp = os.path.join(here, "hist_frozen_px1113.json")
    frozen = json.load(open(fp, encoding="utf-8")) if os.path.exists(fp) else {}
    m = 0
    for rr in range(1, RD.max_row + 1):
        v = RD.cell(rr, 1).value
        if isinstance(v, str) and re.match(r'^="v\d', v) and "&" in v:
            t = frozen.get(v[2:62])
            if t is not None:
                RD.cell(rr, 1).value = t; m += 1
    assert m == len(frozen), ("hist_frozen_px1113 未全部套用", m, len(frozen))
    log.append(f"凍結版本沿革 {m} 列")
    # 6 版本說明：A1 標題段版號、A2 版本列、「本版要點」列（以 old_ver（開頭的靜態列；cfg["note_in_post"] 時由該公司後處理自行放）
    a1 = RD["A1"].value
    head, sep, rest = str(a1).partition("；")
    if cfg["old_ver"] in head:
        head = head.replace(cfg["old_ver"], cfg["ver"], 1)
    RD["A1"].value = f"{head}{sep}{rest}；{cfg['ver']}（2026-11-16）：現價與同業報價改 2026-11-13 收盤（工作單 px1113）"
    a2 = RD["A2"].value
    if isinstance(a2, str) and a2.startswith(f"版本：{cfg['old_ver']}（"):
        RD["A2"].value = a2.replace(f"版本：{cfg['old_ver']}（", f"版本：{cfg['ver']}（2026-11-16：現價與同業報價改 2026-11-13 收盤；前版 {cfg['old_ver']}）；{cfg['old_ver']}（", 1)
        log.append("A2 版本列")
    if cfg.get("note") and not cfg.get("note_in_post"):
        rows = [rr for rr in range(1, RD.max_row + 1) if isinstance(RD.cell(rr, 1).value, str) and RD.cell(rr, 1).value.startswith(f"{cfg['old_ver']}（")]
        assert len(rows) <= 1, ("本版要點列", cfg["old_ver"], rows)   # 0 列：另一側讀數／敏感度快照建置（後處理未加 v1.3 列）
        if rows:
            RD.cell(rows[0], 1).value = cfg["note"] + RD.cell(rows[0], 1).value
        log.append(f"本版要點 README!A{rows[0] if rows else '（無）'}")
    log.append(f"未換價同業：{unchanged}")
    return "；".join(log)
