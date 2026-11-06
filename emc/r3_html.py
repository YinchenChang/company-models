# 第三輪修正（合併前，2026-11-06）HTML 設定後處理：gen_html_json.py 在寫出 html.json 前呼叫 post(cfg, wb, 幣別)。
# 公司資料夾各放同一份（公司自有）。對照第三輪修正清單：
#   32／44 情境驅動表一律依 In 頁「情境驅動對照」節標題找列範圍（不寫死列號），每列鍵值須以 scn_ 開頭且在該節內，表尾須有情境權重（自檢）
#   44 「情境定義與口徑」kv：各格內容須與標籤相符（情境 A／B／C 列首須為同字母情境；其他標籤須出現在內容中），不符時在同欄 ±3 列內找正確列
#   15／31 情境驅動表：% 與 pt 用百分比、x 用兩位小數
#   38 DCF 明細 row_fmt 依列首判斷，只對比例列（折舊攤銷 ÷、資本支出 ÷、稅率、終值現值占、評價日後占）套百分比
#   17／34 橫幅三格標示幣別
import re


def _txt(wb, ref):
    m = re.match(r"^@([^!]+)!([A-Z]+)(\d+)", str(ref))
    return str(wb[m.group(1)][f"{m.group(2)}{m.group(3)}"].value or "") if m else str(ref)


def _scn_block(wb):
    IN = wb["In"]
    s0 = [r for r in range(1, IN.max_row + 1) if "情境驅動對照" in str(IN.cell(r, 1).value or "")[:12]]
    assert len(s0) == 1, s0
    h = s0[0] + 1
    assert str(IN.cell(h, 1).value) == "鍵", IN.cell(h, 1).value
    rows, r = [], h + 1
    while r <= IN.max_row and str(IN.cell(r, 1).value or "").startswith("scn_"):
        u = IN.cell(r, 3).value
        f = "p1" if u in ("%", "pt", "%/年") else "n2" if u == "x" else "n3" if u == "十億股" else "n1" if isinstance(u, str) and ("B" in u or "億" in u) else "n2"
        vals = [IN.cell(r, c).value for c in (4, 5, 6)]
        rows.append([f"@In!B{r}"] + [f"@In!{c}{r}|{f}" if isinstance(v, (int, float)) else f"@In!{c}{r}" for c, v in zip("DEF", vals)] + [f"@In!C{r}"])
        r += 1
    assert rows, "情境驅動對照節沒有 scn_ 列"
    keys = [str(IN.cell(int(x[0][5:]), 1).value) for x in rows]
    assert all(k.startswith("scn_") for k in keys), keys
    assert any(k == "scn_w" or "權重" in str(IN.cell(int(x[0][5:]), 2).value) for k, x in zip(keys, rows)), ("表尾缺情境權重", keys)
    return [f"@In!B{h}", f"@In!D{h}", f"@In!E{h}", f"@In!F{h}", f"@In!C{h}"], rows, (s0[0], r)


def _kv_fix(wb, b, log):
    for it in b["items"]:
        lab, ref = it[0], it[1]
        if not (isinstance(ref, str) and ref.startswith("@")):
            continue
        m = re.match(r"^@([^!]+)!([A-Z]+)(\d+)(.*)$", ref)
        sh, col, row, tail = m.group(1), m.group(2), int(m.group(3)), m.group(4)
        lt = _txt(wb, lab) if str(lab).startswith("@") else str(lab)
        mm = re.match(r"^情境 ([ABC])", lt)
        if mm:
            ok = lambda t: t.startswith(f"情境 {mm.group(1)}")
        else:
            key = lt.replace("與口徑", "")
            ok = lambda t, key=key: key in t
        if ok(str(wb[sh][f"{col}{row}"].value or "")):
            continue
        cand = [row + d for d in (-1, 1, -2, 2, -3, 3) if row + d > 0 and ok(str(wb[sh][f"{col}{row + d}"].value or ""))]
        assert cand, ("情境定義 kv 對不上", lt, ref)
        it[1] = f"@{sh}!{col}{cand[0]}{tail}"
        log.append(f"kv「{lt}」{ref} → {it[1]}")


def _row_fmt(wb, b):
    rg = b.get("rowgen")
    if not rg or rg.get("sheet") != "DCF":
        return
    D = wb["DCF"]
    a, z = map(int, rg["rows"].split("-"))
    cols = rg["cols"][1:]
    rf = {}
    for r in range(a, z + 1):
        lab = str(D.cell(r, 1).value or "")
        if lab.startswith(("折舊攤銷 ÷", "資本支出 ÷", "稅率", "終值現值占", "評價日後占")):
            rf[str(r)] = {c: "p1" for c in cols}
        elif lab.startswith("折現因子"):
            rf[str(r)] = {c: "n3" for c in cols}
    rg["row_fmt"] = rf


def post(cfg, wb, ccy):
    log = []
    head, rows, rng = _scn_block(wb)
    n = 0
    for s in cfg["sections"]:
        for b in s["blocks"]:
            if b["type"] == "table" and str(b.get("caption", "")).startswith("情境驅動"):
                b["head"], b["rows"] = head, rows
                b.pop("rowgen", None); n += 1
            if b["type"] == "table":
                _row_fmt(wb, b)
    assert n == 1, n
    # 44：自檢（每列鍵值在節內）
    IN = wb["In"]
    for x in rows:
        r = int(x[0][5:])
        assert rng[0] < r < rng[1] and str(IN.cell(r, 1).value).startswith("scn_"), r
    log.append(f"情境驅動表 {len(rows)} 列（In 第 {rng[0]}–{rng[1] - 1} 列）")
    th = next(s for s in cfg["sections"] if s["id"] == "thesis")
    for b in th["blocks"]:
        if b["type"] == "kv":
            _kv_fix(wb, b, log)
    # 17／34：橫幅幣別
    tag = f"（{ccy}）"
    for t in cfg["hero"]["tiles"]:
        lab = t[0]
        flat = "".join(_txt(wb, x) if str(x).startswith("@") else str(x) for x in (lab if isinstance(lab, list) else [lab]))
        if ccy not in flat:
            t[0] = (lab + [tag]) if isinstance(lab, list) else ([lab, tag] if str(lab).startswith("@") else lab + tag)
    print("r3_html：" + "；".join(log))
    return log
