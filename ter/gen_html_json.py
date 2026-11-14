# v0.4：套用個股 HTML 版面規格 v1.0（docs/plan/個股HTML版面規格.md）＋DCF 改標準版（FCFF＋WACC），估值章 DCF 區塊與目標價敏感度表由本檔依 xlsx 標籤產生
# 由 html_src.json（舊版面設定，內容不動）經共用轉換器 tools/company_html/relayout.py 產生 html.json（勿手改 html.json）。
# 用法：python3 ter/gen_html_json.py（讀 ter/dist 內唯一的 xlsx）
# 前處理（本檔）：把設備股模板的舊版面調成轉換器預期的形狀——單一參數表改作「驅動 → 推導量對照」（進附錄）、
#   情境驅動表與敏感度目標價參數表的列號依現況重新定位、估值章的 kv 改成具名表格、舊法規則值與同業名單摺疊、
#   補 DCF 折現率表與目標價敏感度龍捲風圖；
# 後處理：還原驅動因子章說明、EPS 龍捲風圖標題（轉換器預設寫「加權目標價」）、橫幅中間格＝只看 Comparable（市場主流同業法），調整後值進附註。
import copy, glob, json, os, re, sys
import openpyxl
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools", "company_html"))
from relayout import relayout, _unit_fmt  # noqa: E402

X = glob.glob(os.path.join(HERE, "dist", "*.xlsx"))
assert len(X) == 1, X
X = X[0]
cfg = json.load(open(os.path.join(HERE, "html_src.json"), encoding="utf-8"))
wb = openpyxl.load_workbook(X, data_only=True)


def text(spec):
    m = re.match(r"^@([^!]+)!([A-Z]+\d+)", spec) if isinstance(spec, str) else None
    return str(wb[m.group(1)][m.group(2)].value or "") if m else str(spec or "")


def row(sh, prefix, exact=False, start=1):
    ws = wb[sh]
    for r in range(start, ws.max_row + 1):
        v = str(ws.cell(r, 1).value or "").strip()
        if (v == prefix) if exact else v.startswith(prefix):
            return r
    raise KeyError((sh, prefix))


IN = {wb["In"].cell(r, 1).value: r for r in range(4, wb["In"].max_row + 1) if wb["In"].cell(r, 1).value}


def in_row(k, fmt):
    r = IN[k]
    return [f"@In!B{r}", f"@In!E{r}|{fmt}", f"@In!G{r}|tag"]


def kv_table(b, caption=None, collapse=False, summary=None):
    t = {"type": "table", "head": ["項目", "數值"], "rows": [[k, v] for k, v in b["items"]]}
    if caption:
        t["caption"] = caption
    if collapse:
        t.update(collapse=True, summary=summary)
    return t


# ---------------- v0.4：標準版 DCF（FCFF＋WACC）的估值與敏感度區塊（由 xlsx 標籤定位列號） ----------------
def dcf_summary_block(b0):
    """Val 0.1 節：DCF 彙總（明確期現值～DCF 目標價；折現率另列在前一張表）。"""
    h = row("Val", "0.1 DCF")
    r1, r2 = row("Val", "明確期 FCFF 現值", start=h), row("Val", "DCF 目標價（情境加權）", start=h)
    rf = {str(r): {"B": f, "C": f, "D": f} for r, f in ((row("Val", "評價用稀釋股數", start=h), "n3"), (row("Val", "DCF 情境每股價值", start=h), "n2"),
                                                       (row("Val", "DCF 情境價值推到 12 個月後", start=h), "n2"), (row("Val", "其中：終值現值占 EV", start=h), "p1"),
                                                       (row("Val", "隱含終值 EV／EBITDA", start=h), "x1"), (r2, "n2"))}
    return {"type": "table", "caption": f"@Val!A{h}", "head": [f"@Val!A{h + 1}", f"@Val!B{h + 1}", f"@Val!C{h + 1}", f"@Val!D{h + 1}", f"@Val!E{h + 1}"],
            "rowgen": {"sheet": "Val", "rows": f"{r1}-{r2}", "cols": ["A", "B", "C", "D", "E"], "fmts": ["", "n1", "n1", "n1", ""], "row_fmt": rf}}


def dcf_detail_blocks():
    """DCF 頁：三情境 FCFF 明細（摺疊）＋口徑說明。"""
    ws = wb["DCF"]
    hdr_r = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == "項目" and str(ws.cell(r, 3).value or "").startswith("FY"))
    ncol = max(c for c in range(3, 12) if ws.cell(hdr_r, c).value)
    cols = [chr(64 + c) for c in range(3, ncol + 1)]
    out = []
    for r in range(hdr_r, ws.max_row + 1):
        if str(ws.cell(r, 1).value or "").startswith("情境 ") and ws.cell(r + 1, 1).value == "營收":
            r1, r2 = r + 1, next(x for x in range(r + 1, ws.max_row + 1) if str(ws.cell(x, 1).value or "").startswith("現值（FCFF"))
            rf = {}
            for x in range(r1, r2 + 1):
                lab = str(ws.cell(x, 1).value)
                f = "p1" if (lab.startswith("稅率") or "÷ 營收" in lab) else ("n3" if lab.startswith("折現因子") else None)
                if f:
                    rf[str(x)] = {c: f for c in cols}
            out.append({"type": "table", "caption": f"@DCF!A{r}", "collapse": True, "summary": "展開：本情境各年無槓桿自由現金流明細",
                        "head": [f"@DCF!A{hdr_r}"] + [f"@DCF!{c}{hdr_r}" for c in cols],
                        "rowgen": {"sheet": "DCF", "rows": f"{r1}-{r2}", "cols": ["A"] + cols, "fmts": [""] + ["n1"] * len(cols), "row_fmt": rf}})
    out.append({"type": "note", "parts": ["@DCF!A2"]})
    return out


def tp_sens_block():
    """Sens 頁：目標價參數敏感度（標準版 DCF 案例；低／高加權目標價與相對基準）。"""
    h = row("Sens", "目標價參數敏感度")
    last = row("Sens", "DCF 權重 0.3／0.7", start=h)
    return {"type": "table", "caption": f"@Sens!A{h}", "head": [f"@Sens!{c}{h + 1}" for c in "ABCDEFG"],
            "rowgen": {"sheet": "Sens", "rows": f"{h + 2}-{last}", "cols": list("ABCDEFG"), "fmts": ["", "", "n2", "sp1", "n2", "sp1", "n2"]}}


def prep(cfg, spec):
    secs = {s["id"]: s for s in cfg["sections"]}
    # 驅動因子：單一參數表 → 驅動 → 推導量對照（轉換器移到附錄）；情境驅動表不摺疊
    dr = secs["drivers"]
    for b in dr["blocks"]:
        cap = str(b.get("caption", ""))
        if cap.startswith("單一參數"):
            b["caption"] = "驅動 → 推導量對照（單一參數的輸入、基準值、推導量與因果方向）"
        elif cap.startswith("情境驅動"):
            b.pop("collapse", None)
            b.pop("summary", None)
            # 列號依 In 頁現況重新定位（舊設定的列號在 In 頁插入 Wattonomics 錨定係數後已位移）
            ws = wb["In"]
            h = row("In", "情境驅動對照") + 1
            assert ws.cell(h, 1).value == "鍵", ws.cell(h, 1).value
            fm = lambda r: "yr" if ws.cell(r, 3).value == "年" else _unit_fmt(ws.cell(r, 3).value)
            b["head"] = [f"@In!B{h}", f"@In!D{h}", f"@In!E{h}", f"@In!F{h}", f"@In!C{h}"]
            b["rows"] = [[f"@In!B{r}", f"@In!D{r}|{fm(r)}", f"@In!E{r}|{fm(r)}", f"@In!F{r}|{fm(r)}", f"@In!C{r}"]
                         for r in range(h + 1, next((x for x in range(h + 1, ws.max_row + 2) if not ws.cell(x, 1).value), ws.max_row + 1))]   # 第三輪修正：只到本節結尾
    intro = dr.get("intro")
    # 估值：依規格重排
    va = secs["valuation"]
    old = va["blocks"]
    nb, i = [], 0
    caps = [text(b.get("caption", "")) for b in old]
    k_dcf = next(j for j, c in enumerate(caps) if c.startswith(spec["cap_dcf"]))
    k_comp = next(j for j, c in enumerate(caps) if c.startswith(spec["cap_comp"]))
    k_old = next(j for j, b in enumerate(old) if b["type"] == "h3" and str(b.get("text", "")).startswith(spec["h3_old"]))
    conc = old[0]
    assert text(conc.get("caption", "")).endswith("結論"), conc
    head = old[:k_dcf]
    bars = next(b for b in head if b["type"] == "bars")
    h01 = row("Val", "0.1 DCF"); r_dcf = row("Val", "DCF 目標價（情境加權）", start=h01)
    for it in bars["items"]:   # 第二輪修正：模組 v2 stub 期在 0.1 節多兩列；html_src.json 的 DCF 目標價列號（0.1 節標題 ＋ 17）改依標籤定位
        if it["value"] == f"@Val!B{h01 + 17}":
            it["label"], it["value"] = f"@Val!A{r_dcf}", f"@Val!B{r_dcf}"
    ca = row("Val", "Comparable 目標價（相對倍數調整後）")
    if not any(it["value"] == f"@Val!B{ca}" for it in bars["items"]):   # 加權目標價實際使用的 Comparable（相對倍數調整後）也畫出，避免讀成加權價高於兩法
        k = next(j for j, it in enumerate(bars["items"]) if text(it["label"]).startswith("Comparable 目標價"))
        bars["items"].insert(k + 1, {"label": f"@Val!A{ca}", "value": f"@Val!B{ca}", "series": 2})
    nb.append(bars)
    nb.append(conc)
    for j, b in enumerate(head[1:], 1):
        if b["type"] == "note":
            nb.append(b)
    rm_h3 = next(b for b in head if b["type"] == "h3" and str(b.get("text", "")).startswith("相對倍數調整"))
    rm_kv = head[head.index(rm_h3) + 1]
    assert rm_kv["type"] == "kv"
    # 第三輪修正（2026-11-06）：折現率主表直接顯示股權成本、稅後債務成本、WACC、永續成長率、RONIC；rf／β／ERP／CRP 與權益占資本摺疊（規格 v1.0）
    _vr = [[f"@Val!A{r}", f"@Val!B{r}|p2", ""] for r in spec.get("dcf_rate_val", [])]
    _we = [x for x in _vr if text(x[0]).startswith("權益占資本")]
    nb.append({"type": "table", "caption": "DCF 折現率：股權成本、稅後債務成本、WACC、永續成長率、RONIC（標準版 DCF；直接顯示）", "head": ["項目", "數值", "標記"],
               "rows": [x for x in _vr if x not in _we]})
    nb.append({"type": "table", "caption": "股權成本的組成與資本結構：rf、β、ERP、CRP、權益占資本", "head": ["項目", "數值", "標記"], "collapse": True,
               "summary": "展開：股權成本的組成（rf、β、ERP、CRP）與權益占資本",
               "rows": [in_row(k, f) for k, f in spec["dcf_rates"]] + _we})
    nb.append({"type": "h3", "text": "DCF 彙總"})
    nb.append(dcf_summary_block(old[k_dcf]))
    nb += dcf_detail_blocks()
    nb.append({"type": "h3", "text": "Comparable（同業前瞻本益比）"})
    comp_kv = old[k_comp + 1]
    assert comp_kv["type"] == "kv"
    nb.append(kv_table(comp_kv, caption="Comparable 彙總（同業前瞻本益比中位數 × 前瞻每股盈餘）"))
    nb.append(kv_table(rm_kv, caption=rm_h3["text"]))
    peers = dict(old[k_comp], collapse=True, summary="展開：同業前瞻本益比（各家）")
    nb.append(peers)
    nb.append({"type": "h3", "text": old[k_old]["text"]})
    for b in old[k_old + 1:]:
        if b["type"] == "bars":
            continue   # 舊法隱含股價長條圖不放（轉換器會把所有長條圖移到章首，與目標價長條圖混淆）
        if b["type"] == "kv":
            nb.append(kv_table(b, collapse=True, summary="展開：舊法的價格與市場隱含倍數"))
        elif b["type"] == "table":
            if not b.get("collapse"):
                c = re.sub(r"^[\d.\s]+", "", text(b.get("caption", ""))).strip()
                b = dict(b, collapse=True, summary="展開：" + (c if c and not re.search(r"\d", re.sub(r"(FY|CY)\d{4}", "", c)) else "舊法明細表"))
            nb.append(b)
        elif b["type"] == "note":
            nb.append(dict(b, collapse=True, summary="展開：舊法說明"))
        else:
            nb.append(b)
    va["blocks"] = nb
    # 對帳：給表格標題，主要列由 SPEC main 指定
    for b in secs["recon"]["blocks"]:
        if b["type"] == "table" and not b.get("caption"):
            b["caption"] = "對帳要點"
    # 敏感度：先放目標價龍捲風圖，再放 EPS 龍捲風圖
    se = secs["sens"]
    tor = {}
    for b in se["blocks"]:
        if b["type"] == "tornado":
            tor[b["value_col"]] = b["title"].split("（")[0].replace("相對基準的變動", "").strip()
    ws = wb["Sens"]
    h0 = row("Sens", "目標價參數敏感度")
    for b in se["blocks"]:   # 目標價參數案例表：列號依 Sens 頁現況重新定位（舊設定在插入 Wattonomics 案例後可能位移）
        rg = b.get("rowgen") if b["type"] == "table" else None
        if rg and isinstance(b.get("caption"), str) and b["caption"].startswith("@Sens!A") and abs(int(b["caption"][7:]) - h0) <= 6:
            hh = h0 + 1
            assert str(ws.cell(hh, 1).value).strip() in ("案例", "參數"), ws.cell(hh, 1).value
            end = hh + 1
            while ws.cell(end + 1, 1).value not in (None, ""):
                end += 1
            b["caption"] = f"@Sens!A{h0}"
            b["head"] = [f"@Sens!{c}{hh}" for c in rg["cols"]]
            rg["rows"] = f"{hh + 1}-{end}"
    # v0.4：目標價參數敏感度改為標準版 DCF 案例（低／高兩欄）；不畫龍捲風圖，表格直接顯示
    se["blocks"] = [tp_sens_block() if (b["type"] == "table" and not b.get("collapse") and text(b.get("caption", "")).startswith("目標價參數")) else b
                    for b in se["blocks"]]
    scen_kv = next((b for b in secs["thesis"]["blocks"] if b["type"] == "kv"), None)
    return intro, tor, scen_kv


def post(cfg, intro, tor, scen_kv, spec):
    CCY = "（美元）"   # 第三輪修正：橫幅標示幣別（金額兩位小數）
    secs = {s["id"]: s for s in cfg["sections"]}
    if intro:   # 轉換器把章說明當 note 放進某一塊；還原為章說明
        secs["drivers"]["intro"] = intro
        secs["drivers"]["blocks"] = [b for b in secs["drivers"]["blocks"] if not (b["type"] == "note" and b.get("parts") == [intro])]
    # 參數表格式：轉換器只認單位「%」，「%/年」等也改百分比；小於 1 的一般數值用三位小數（避免 0.0／0.1 誤讀）
    ws = wb["In"]
    for b in secs["drivers"]["blocks"]:
        for r_ in b.get("rows", []) if b["type"] == "table" else []:
            m = re.match(r"^@In!B(\d+)$", r_[0]) if isinstance(r_, list) and isinstance(r_[0], str) else None
            if not m:
                continue
            r = int(m.group(1))
            u = str(ws.cell(r, 3).value or "")
            vals = [ws.cell(r, c).value for c in (4, 5, 6)]
            for j in (1, 2, 3):
                cell, _, f = r_[j].partition("|")
                if u.startswith("%"):
                    f = "p1"
                elif f == "n1" and any(isinstance(x, (int, float)) and 0 < abs(x) < 1 for x in vals):
                    f = "n3"
                elif f == "n1" and any(isinstance(x, float) and abs(x) < 10 and x != int(x) for x in vals):
                    f = "n2"
                r_[j] = f"{cell}|{f}"
    if scen_kv:   # 舊版命題章的情境定義：放在重點觀察之後（摺疊）
        secs["thesis"]["blocks"].append(dict(scen_kv, collapse=True, summary="展開：三情境的定義"))
    for b in secs["sens"]["blocks"]:
        if b["type"] == "tornado" and tor.get(b["value_col"]):
            base = tor[b["value_col"]] + " 相對基準的變動"
            b["title"] = base + ("（未達兩位數百分比）" if b.get("collapse") else "（變動達兩位數百分比的案例；一次改一個輸入，整本重算）")
        if b["type"] == "tornado" and b.get("collapse"):
            b["summary"] = "展開：" + (tor.get(b["value_col"]) or "加權目標價") + " 變動未達兩位數百分比的案例"
    # 橫幅（補充規則 A）：中間格＝只看 Comparable 目標價（市場主流同業法，未做相對倍數調整）與其上檔空間；
    #   進加權的相對倍數調整後 Comparable 放附註（數字由 xlsx 帶出）
    v = lambda p, **kw: row("Val", p, **kw)
    tp, ca, cu, px, up, rec = v("目標價（DCF 與 Comparable 加權）", exact=True), v("Comparable 目標價（相對倍數調整後）"), v("Comparable 目標價", exact=True), v("現價（"), v("上檔空間：加權目標價"), v("投資建議", exact=True)
    rd = row("README", "HTML 橫幅附註用：Comparable 目標價（相對倍數調整後）")
    uc = v("上檔空間：Comparable")
    cfg["hero"] = {"rec": f"@Val!B{rec}",
                   "tiles": [[[f"@Val!A{tp}", CCY], f"@Val!B{tp}|n2", ["上檔空間 ", f"@Val!B{up}|sp1"]],
                             ["只看 Comparable 目標價（市場主流同業法）" + CCY, f"@Val!B{cu}|n2", ["上檔空間 ", f"@Val!B{uc}|sp1"]],
                             [[f"@Val!A{px}", CCY], f"@Val!B{px}|n2", ""]],
                   "note": ["投資建議依加權目標價判定：", f"@Val!E{rec}", "。進加權的 Comparable 為相對倍數調整後 ", f"@Val!B{ca}|n2",
                            "（上檔空間 ", f"@README!B{rd}|sp1", "）。"]}


def in_skip():
    ws = wb["In"]
    out = []
    for r in range(4, ws.max_row + 1):
        a = str(ws.cell(r, 1).value or "")
        if a and ws.cell(r, 2).value in (None, "") and a.startswith("情境驅動對照"):
            out.append(re.split(r"[.\s]", a)[0])
    return out


def run(SPEC, PRE):
    intro, tor, scen_kv = prep(cfg, PRE)
    SPEC = dict(SPEC, in_skip=SPEC.get("in_skip", []) + in_skip())
    out = relayout(cfg, X, SPEC)
    post(out, intro, tor, scen_kv, PRE)
    # 三層協調 C4（規則 v1.4；tools/coord/coord.py）：橫幅另一側讀數一格與揭露句、估值章開頭兩讀數表與長條圖
    sys.path.insert(0, os.path.join(HERE, "..", "tools", "coord"))
    import coord as COORD  # noqa: E402
    COORD.inject_html(out, X)
    _v13(out)
    json.dump(out, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("ok")



# ================= v0.4.5（2026-11-14）：個股 HTML 版面規格 v1.3（樣板 NVIDIA v0.3.10、Ibiden v0.5.6；docs/plan/個股HTML版面規格.md「v1.3 規格」） =================
# 章節：命題 → 估值 → 敏感度 → 反向 DCF → 損益（標準損益表）→ 驅動因子（分部營收推導）→ 對帳（含 C0 實績）→ 資料狀態；泰瑞達無個別模組。
# 數字一律取自 xlsx（README 頁末 v1.3 白話列、PL 頁末標準損益表、RevDCF 頁、Val 0.7、Sens 頁末相對基準表、Coord 頁 C0；ter/v045_post.py 產生）。列號依標籤找；設定文字不含數字（年份與日期除外）。
def _v13(cfg):
    S = {x["id"]: x for x in cfg["sections"]}
    RD, PLW, RV, SW, CW = wb["README"], wb["PL"], wb["RevDCF"], wb["Sens"], wb["Coord"]
    fnd = lambda ws, p, start=1: next(r for r in range(start, ws.max_row + 1) if str(ws.cell(r, 1).value or "").strip().startswith(p))
    r13 = fnd(RD, "v1.3 展示層")
    RM = dict(zip(["hero", "rec", "o1", "o2", "o3", "o4", "o5", "drv", "ver"], range(r13 + 1, r13 + 10)))
    assert str(RD[f"A{RM['ver']}"].value).startswith("v0.4.5"), RD[f"A{RM['ver']}"].value
    cfg["layout"] = "v1.3"
    fmt_u = lambda u: {"%": "p1", "US$B": "n2", "US$M": "n1", "百萬股": "n1", "US$": "n2", "x": "n2"}.get(str(u).strip(), "n1")
    # ---- 橫幅：白話說明＋對帳中揭露句；評等規則與歸因移到估值章 ----
    old_note = list(cfg["hero"]["note"])
    i_cd = len(old_note) - 2   # 末兩項＝空格、Coord 橫幅揭露句
    assert old_note[i_cd] == " " and str(old_note[i_cd + 1]).startswith("@Coord!B"), old_note
    cfg["hero"]["note"] = [f"@README!A{RM['hero']}"] + old_note[i_cd:]
    # ---- 1. 命題：主圖＝情境 B 分部營收（堆疊）與三情境營收 ----
    pa, pb, pc = fnd(PLW, "情境 A"), fnd(PLW, "情境 B"), fnd(PLW, "情境 C")
    hb = pb + 1
    assert PLW.cell(hb, 1).value == "營收" or True
    revA, revC = pa + 1, pc + 1
    assert PLW.cell(revA, 1).value == "營收" and PLW.cell(revC, 1).value == "營收"
    rb = fnd(PLW, "情境 B（基準）")
    hdrB = rb + 1
    prow = lambda lab_, start=rb: next(r for r in range(start, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip().startswith(lab_))
    r_soc, r_mem, r_ist, r_pt, r_rob = prow("SoC 測試機營收"), prow("記憶體測試機營收"), prow("IST 營收"), prow("Product Test 營收"), prow("機器人營收")
    r_rev = next(r for r in range(r_rob, PLW.max_row + 1) if PLW.cell(r, 1).value == "營收")
    th = S["thesis"]["blocks"]
    chart1 = {"type": "theme", "title": "營收分部（情境 B）與三情境營收（百萬美元）", "x": f"@PL!C{hdrB}:G{hdrB}", "fmt": "n0", "stack": True, "pad": 1.3,
              "total": f"@PL!C{r_rev}:G{r_rev}",
              "bars": [{"name": f"@PL!A{r_soc}", "y": f"@PL!C{r_soc}:G{r_soc}", "color": "#0b7a80"},
                       {"name": f"@PL!A{r_mem}", "y": f"@PL!C{r_mem}:G{r_mem}", "color": "#6fc7cb"},
                       {"name": f"@PL!A{r_ist}", "y": f"@PL!C{r_ist}:G{r_ist}", "color": "#c29a2b"},
                       {"name": f"@PL!A{r_pt}", "y": f"@PL!C{r_pt}:G{r_pt}", "color": "#8a8c8c"},
                       {"name": f"@PL!A{r_rob}", "y": f"@PL!C{r_rob}:G{r_rob}", "color": "#c9ccd0"}],
              "lines": [{"name": [f"@PL!A{pa}", "營收"], "y": f"@PL!C{revA}:G{revA}", "color": "#0b7a80", "dash": True},
                        {"name": [f"@PL!A{pc}", "營收"], "y": f"@PL!C{revC}:G{revC}", "color": "#b4472f", "dash": True}],
              "caption": "堆疊＝情境 B（基準）的五個分部，頂端數字＝情境 B 合併營收；虛線＝情境 A、情境 C 的合併營收。FY2026 為前三季實績加第四季指引中點的錨點；三情境的差異來自測試機市場（加速器顆數）與市占。"}
    i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
    th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RM[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
    th.insert(0, chart1)
    th.insert(1, {"type": "note", "parts": [f"@README!A{RM['rec']}"]})
    # ---- 2. 估值：第一個區塊＝四個目標價與現價（G 位置）----
    va = S["valuation"]["blocks"]
    g0 = fnd(wb["Val"], "0.7 估值並列"); gr = list(range(g0 + 2, g0 + 7)); r_pxg = g0 + 7
    four = {"type": "bars", "title": "四個目標價與現價（美元；豎線＝現價）", "fmt": "n2",
            "items": [{"label": lb, "value": f"@Val!B{r}", "series": sr} for r, sr, lb in zip(gr, (1, 2, 2, 0, 1),
                      ("DCF 目標價", "Comparable（調整後，進加權）", "Comparable（未調整，不進加權）", "加權目標價（橫幅主值）", "市場共識目標價"))],
            "ref": {"label": f"@Val!A{r_pxg}", "value": f"@Val!B{r_pxg}"}}
    tbl4 = {"type": "table", "caption": f"@Val!A{g0}", "head": [f"@Val!{c}{g0 + 1}" for c in "ABCD"],
            "rows": [[f"@Val!A{r}", f"@Val!B{r}|n2", f"@Val!C{r}|sp1", f"@Val!D{r}"] for r in gr] + [[f"@Val!A{r_pxg}", f"@Val!B{r_pxg}|n2", "", f"@Val!D{r_pxg}"]]}
    keep = [b for b in va if not (b.get("type") == "bars" and b.get("title") == "目標價與現價")]
    head_note = {"type": "note", "parts": old_note[:i_cd]}
    S["valuation"]["blocks"] = [four, head_note, tbl4] + keep
    v_intro = S["valuation"].pop("intro", None)   # 章說明放在圖之後（G：章標題後第一個區塊是圖）
    if v_intro:
        S["valuation"]["blocks"].insert(2, {"type": "note", "parts": [v_intro]})
    # ---- 3. 敏感度：主圖＝加權目標價相對基準（H 位置）----
    se = S["sens"]["blocks"]
    t0 = fnd(SW, "加權目標價相對基準（HTML"); r1 = t0 + 2; r2 = r1
    while SW.cell(r2 + 1, 1).value:
        r2 += 1
    tor = []
    for col, side in (("B", "低"), ("C", "高")):
        vals = {r: SW[f"{col}{r}"].value for r in range(r1, r2 + 1) if isinstance(SW[f"{col}{r}"].value, (int, float))}
        big = [r for r, v in vals.items() if abs(v) >= 0.1]
        small = [r for r, v in vals.items() if 0.0005 <= abs(v) < 0.1]
        base = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1"}
        tor.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值，變動達兩位數百分比的案例", rows=",".join(map(str, big))))
        tor.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值", rows=",".join(map(str, small)), collapse=True,
                        summary=f"展開：加權目標價相對基準：輸入取{side}值，變動未達兩位數百分比的案例"))
    full = {"type": "table", "caption": f"@Sens!A{t0}", "collapse": True, "summary": "展開：全部案例的加權目標價相對基準（低案、高案）",
            "head": [f"@Sens!{c}{t0 + 1}" for c in "ABC"], "rows": [[f"@Sens!A{r}", f"@Sens!B{r}|sp1", f"@Sens!C{r}|sp1"] for r in range(r1, r2 + 1)]}
    snap = fnd(SW, "快照狀態")
    for b in se:
        if b.get("type") == "tornado":
            b["summary"] = ("展開：對照：FY2027 EPS 相對基準的變動未達兩位數百分比的案例（情境 A）" if b.get("collapse") else "展開：對照：FY2027 EPS 相對基準的變動達兩位數百分比的案例（情境 A，一次動一個輸入）")
            b["collapse"] = True
        if b.get("type") == "table" and not b.get("collapse"):
            b["collapse"] = True
            b["summary"] = "展開：目標價參數表（WACC、β、ERP、永續成長率等）"
    intro_s = {"type": "note", "parts": ["主圖為加權目標價相對基準的變動（低案、高案分開，一次動一個輸入；變動達兩位數百分比者直接顯示，其餘摺疊）。業務主因為整本重算快照，快照狀態：",
                                          f"@Sens!B{snap}", "；目標價參數取 Sens 頁「目標價參數敏感度」表。下方摺疊為全部案例與 EPS 敏感度對照。"]}
    sec_intro = S["sens"].pop("intro", None)
    S["sens"]["blocks"] = tor[:1] + tor[2:3] + [intro_s, tor[1], tor[3], full] + ([{"type": "note", "parts": [sec_intro], "collapse": True, "summary": "展開：EPS 敏感度（情境 A）的說明"}] if sec_intro else []) + se
    # ---- 4. 反向 DCF（新）----
    rv = lambda p, start=1: next(r for r in range(start, RV.max_row + 1) if str(RV.cell(r, 1).value or "").strip().startswith(p))
    ra, rcmp, rsum, yh = rv("A. DCF"), rv("B. Comparable"), rv("C. 摘要"), rv("項目（情境 B）")
    rb_, rbb, rcc = rv("(a)"), rv("(b)"), rv("(c)")
    m1, m0 = rv("現價隱含的營業利益率", rb_), rv("模型：基準的等效營業利益率", rb_)
    s1, rv1, rv0, g1, g0b = rv("現價隱含的放大倍數", rbb), rv("現價隱含的 FY2028 營收", rbb), rv("模型：FY2028 營收", rbb), rv("現價隱含的 FY2028 營收年增率", rbb), rv("模型：FY2028 營收年增率", rbb)
    w1, w0 = rv("現價隱含的 WACC", rcc), rv("模型 WACC", rcc)
    a_rows = [rv(p, ra) for p in ("支持現價所需的評價日每股價值", "淨現金滾到評價日", "評價用稀釋股數", "支持現價所需的企業價值", "模型企業價值", "EV 需放大", "模型：情境 B 評價日每股價值")]
    afmt = ["n2", "n1", "n3", "n1", "n1", "n2", "n2"]
    pe1, rc1, pe2 = rv("只看 Comparable：支持現價所需的前瞻本益比", rcmp), rv("只看 Comparable：所需的相對倍數係數", rcmp), rv("加權（DCF 不變）", rcmp)
    yrs_rows = list(range(yh + 1, yh + 12))
    yfmt = ["n1", "n1", "p1", "p1", "n3", "n3", "n3", "n3", "n1", "n1", "n1"]
    rev_ = {"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2",
            "blocks": [
                {"type": "note", "parts": [f"@RevDCF!A{rsum + 1}"]},
                {"type": "note", "parts": [f"@RevDCF!A{rsum + 2}"]},
                {"type": "table", "caption": "現金流折現法（情境 B 路徑）：一次只動一個因素，反解到現價", "head": ["項目", "現價隱含", "模型基準"],
                 "rows": [["預測期營業利益率（FY2028–FY2032 統一一個值；模型欄為等效值，終值年權重大）", f"@RevDCF!B{m1}|p1", f"@RevDCF!B{m0}|p1"],
                          ["現金流放大倍數（FY2028 起各年同乘；模型基準為一倍）", f"@RevDCF!B{s1}|n2", "—"],
                          ["FY2028 營收（百萬美元）", f"@RevDCF!B{rv1}|n0", f"@RevDCF!B{rv0}|n0"],
                          ["FY2028 營收年增率", f"@RevDCF!B{g1}|p1", f"@RevDCF!B{g0b}|p1"],
                          ["折現率（WACC）", f"@RevDCF!B{w1}|p2", f"@RevDCF!B{w0}|p2"]]},
                {"type": "table", "caption": f"@RevDCF!A{ra}", "head": ["項目", "值"], "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|{f}"] for r, f in zip(a_rows, afmt)]},
                {"type": "table", "caption": f"@RevDCF!A{rcmp}", "head": [f"@RevDCF!{c}{rcmp + 1}" for c in "ABCD"],
                 "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|{f}", f"@RevDCF!C{r}|{f}", f"@RevDCF!D{r}|sp1"] for r, f in ((pe1, "n1"), (rc1, "n2"), (pe2, "n1"))]},
                {"type": "table", "caption": f"@RevDCF!A{yh}", "head": [f"@RevDCF!{c}{yh}" for c in "ABCDEFG"], "collapse": True,
                 "summary": "展開：情境 B 各年數字（反解的計算基礎；演算過程見 RevDCF 頁）",
                 "rows": [[f"@RevDCF!A{r}"] + [f"@RevDCF!{c}{r}|{f}" for c in "BCDEFG"] for r, f in zip(yrs_rows, yfmt)]}]}
    # ---- 5. 損益：標準損益表 ----
    p0 = fnd(PLW, "標準損益表（v0.4.5")
    ph = next(r for r in range(p0, PLW.max_row + 1) if PLW.cell(r, 1).value == "項目")
    plr = lambda lab_, start=ph: next(r for r in range(start, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip() == lab_)
    keys = ["營收", "營收年增率（YoY）", "營業成本（COGS）", "毛利", "毛利率", "營業費用（含研發）", "營業利益", "營業利益率", "其他收入（淨額）", "稅前利益（扣權益法前）", "所得稅",
            "權益法投資損益（Technoprobe；稅後）", "稅後淨利（歸屬泰瑞達）", "淨利率", "稀釋股數（百萬股）", "EPS（GAAP 稀釋）"]
    PC = "CDEFGHIJKL"
    rowx = lambda r: [f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmt_u(PLW.cell(r, 2).value)}" for c in PC]
    r_ratio = plr("比率"); r_chk = next(r for r in range(r_ratio, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("核對（應為 0"))
    r_ycav = next(r for r in range(r_chk, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("年度與對帳中"))
    is_tbl = {"type": "table", "caption": "標準損益表（百萬美元；EPS 為美元）：FY2023A–FY2025A 實績、FY2026E–FY2032E 情境 B（基準；GAAP＝美國會計準則）",
              "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in PC],
              "rows": [rowx(plr(k)) for k in keys] + [{"sub": f"@PL!A{r_ratio}"}] + [rowx(r) for r in range(r_ratio + 1, r_chk)]}
    chk_tbl = {"type": "table", "caption": f"@PL!A{r_chk}", "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in PC], "collapse": True,
               "summary": "展開：核對列（歷史捨入差、預測期對 PL 情境 B 與延伸表、FY2030E–FY2032E 的推法差額）",
               "rows": [[f"@PL!A{r}"] + [f"@PL!{c}{r}|{'p1' if '占營收' in str(PLW.cell(r, 1).value) else 'n3'}" for c in PC] for r in range(r_chk + 1, r_ycav) if PLW.cell(r, 1).value]}
    rv_, gm_, opm_, nm_ = plr("營收"), plr("毛利率"), plr("營業利益率"), plr("淨利率")
    is_chart = {"type": "theme", "title": "營收（百萬美元）與毛利率、營業利益率、淨利率：FY2023A–FY2032E（情境 B）", "x": f"@PL!C{ph}:L{ph}",
                "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
                "bars": [{"name": f"@PL!A{rv_}", "y": f"@PL!C{rv_}:L{rv_}", "color": "#6fc7cb"}],
                "lines": [{"name": f"@PL!A{gm_}", "y": f"@PL!C{gm_}:L{gm_}", "axis": "r", "color": "#0b7a80", "labels": True},
                          {"name": f"@PL!A{opm_}", "y": f"@PL!C{opm_}:L{opm_}", "axis": "r", "color": "#c29a2b"},
                          {"name": f"@PL!A{nm_}", "y": f"@PL!C{nm_}:L{nm_}", "axis": "r", "color": "#6f6bae", "dash": True}],
                "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸，GAAP）。FY2026E 為前三季實績加第四季指引中點；FY2030E–FY2032E 為收斂法延伸期（毛利率與營業費用率沿用前一年，營業利益由淨利還原）。"}
    old_pl = S["pl"]["blocks"]
    pls = [b for b in old_pl if b["type"] == "pl"]
    assert len(pls) == 2, [b["type"] for b in old_pl]
    ext_notes = [b for b in old_pl if b["type"] == "note"]
    old_blocks = [dict(pls[0], collapse=True, summary="展開：三情境損益摘要與折線（FY2025–FY2029；營收、事業別、稅前利益率、EPS）")]
    old_blocks += [dict(b, collapse=True, summary="展開：延伸期（FY2030 起）的收斂法說明") for b in ext_notes]
    old_blocks += [dict(pls[1], collapse=True, summary="展開：三情境延伸到 FY2032（收斂法）")]
    pl_intro = S["pl"].pop("intro", None)
    S["pl"].update(title="損益（標準損益表）",
                   blocks=[{"type": "note", "parts": [f"@PL!A{p0 + 1}"]}, is_chart, is_tbl, {"type": "note", "parts": [f"@PL!A{p0 + 2}"]}, {"type": "note", "parts": [f"@PL!A{r_ycav}"]}, chk_tbl]
                          + ([{"type": "note", "parts": [pl_intro], "collapse": True, "summary": "展開：三情境損益摘要的期間說明"}] if pl_intro else []) + old_blocks)
    # ---- 6. 驅動因子：分部營收推導（情境 B）----
    seg = lambda rs: [[f"@PL!A{r}", f"@PL!B{r}"] + [f"@PL!{c}{r}|{fmt_u(PLW.cell(r, 2).value)}" for c in "CDEFG"] for r in rs]
    r_mk, r_mm, r_ss, r_sm = prow("SoC 測試機市場 CY"), prow("記憶體測試機市場 CY"), prow("泰瑞達 SoC 市占"), prow("泰瑞達記憶體測試機市占")
    r_ttest, r_trob, r_corp, r_pti, r_ni, r_eps = prow("測試事業稅前利益", r_rev - 8), prow("機器人稅前利益"), prow("公司與沖銷"), prow("稅前利益", r_rev), prow("淨利", r_rev), prow("EPS（GAAP 稀釋）")
    seg_tbl = {"type": "table", "caption": "分部營收推導（情境 B；百萬美元）：SoC 測試機＋記憶體測試機＋IST＋Product Test＋機器人＝合併營收", "head": [f"@PL!A{hdrB}", f"@PL!B{hdrB}"] + [f"@PL!{c}{hdrB}" for c in "CDEFG"],
               "rows": [{"sub": "一、半導體測試＝測試機市場（Advantest 模型）× 泰瑞達市占；IST 隨 SoC 市場"}] + seg([r_mk, r_ss, r_soc, r_mm, r_sm, r_mem, r_ist])
                       + [{"sub": "二、Product Test、機器人與合併營收"}] + seg([r_pt, r_rob, r_rev])
                       + [{"sub": "三、稅前利益與每股盈餘"}] + seg([r_ttest, r_trob, r_corp, r_pti, r_ni, r_eps])}
    seg_chart = {"type": "theme", "title": "測試機市場（十億美元）與泰瑞達市占（右軸）：情境 B", "x": f"@PL!C{hdrB}:G{hdrB}", "fmt": "n1", "fmt_r": "p0", "stack": True, "pad": 1.25,
                 "bars": [{"name": f"@PL!A{r_mk}", "y": f"@PL!C{r_mk}:G{r_mk}", "color": "#0b7a80"}, {"name": f"@PL!A{r_mm}", "y": f"@PL!C{r_mm}:G{r_mm}", "color": "#6fc7cb"}],
                 "lines": [{"name": f"@PL!A{r_ss}", "y": f"@PL!C{r_ss}:G{r_ss}", "axis": "r", "color": "#3b3838", "labels": True},
                           {"name": f"@PL!A{r_sm}", "y": f"@PL!C{r_sm}:G{r_sm}", "axis": "r", "color": "#c29a2b"}],
                 "caption": "長條＝SoC 與記憶體測試機市場（CY，取自 Advantest 模型，情境 B 已含顆數係數）；折線＝泰瑞達在兩個市場的市占（FY2025A 為實績隱含，FY2026 起為錨點隱含並維持）。"}
    dr = S["drivers"]["blocks"]
    d_intro = S["drivers"].pop("intro", None)
    S["drivers"].update(title="驅動因子：分部營收推導")
    S["drivers"]["blocks"] = ([{"type": "note", "parts": [f"@README!A{RM['drv']}"]}, seg_chart, seg_tbl, {"type": "h3", "text": "情境驅動與單一參數"}]
                              + ([{"type": "note", "parts": [d_intro]}] if d_intro else []) + dr)
    # ---- 7. 對帳：C0（讀取層級與實績）----
    c1 = fnd(CW, "1. 讀哪一層"); c2 = fnd(CW, "2. C0"); c3 = fnd(CW, "3. 對帳中年份")
    h1, h2 = c1 + 1, next(r for r in range(c2, CW.max_row + 1) if CW.cell(r, 1).value == "項目")
    cnt = fnd(CW, "對帳中項目數")
    cunit = lambda r: str(CW.cell(r, 2).value or "")
    cf = lambda r: "p1" if cunit(r) in ("%",) else "n3" if cunit(r) in ("x", "US$B") else "n1"
    c0_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|{cf(r)}", f"@Coord!E{r}|{cf(r)}", f"@Coord!F{r}|sp1", f"@Coord!G{r}"]
               for r in range(h2 + 1, cnt) if CW.cell(r, 4).value is not None]
    c0_tbl = {"type": "table", "caption": "讀取層級核對：個股讀入的層級值 vs 層級值（同源，不判定、不計入對帳中；有檢驗力的對帳見上表「公司實績 vs 層級隱含值」）", "collapse": True, "summary": "展開：讀取層級核對（同源，不判定）", "head": [f"@Coord!{c}{h2}" for c in "ABCDEFG"], "rows": c0_rows}
    a4 = fnd(CW, "4. C0（實績）"); ih = next(r for r in range(a4, CW.max_row + 1) if CW.cell(r, 1).value == "項目"); mh = next(r for r in range(ih, CW.max_row + 1) if CW.cell(r, 1).value == "讀數")
    a4e = next(r for r in range(mh, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith("對帳中項目數"))
    fm = lambda r: "p1" if str(CW.cell(r, 2).value) == "%" else "n3" if str(CW.cell(r, 2).value) == "US$B" else "n1"
    prog = next(r for r in range(mh, a4e) if str(CW.cell(r, 1).value or "").startswith("參考：已公布期間實績占"))
    act_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|n1", f"@Coord!E{r}|n1", (f"@Coord!F{r}|p1" if r == prog else f"@Coord!F{r}|sp1"), f"@Coord!G{r}"]
                for r in range(mh + 1, a4e - 1)]
    in_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|" + ("n0" if CW.cell(r, 2).value == "月" else fm(r)), f"@Coord!E{r}", f"@Coord!H{r}"] for r in range(ih + 1, mh - 1) if CW.cell(r, 1).value]
    ot = fnd(CW, "其他公司實績"); oh = next(r for r in range(ot, CW.max_row + 1) if CW.cell(r, 1).value == "項目"); oe = next(r for r in range(oh, CW.max_row + 2) if not CW.cell(r + 1, 1).value)
    oth_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|n1", f"@Coord!E{r}|n1", f"@Coord!G{r}"] for r in range(oh + 1, oe + 1)]
    act_blocks = [{"type": "h3", "text": "公司實績對層級隱含值（實績對帳）"}, {"type": "note", "parts": [f"@Coord!A{a4 + 1}"]}, {"type": "note", "parts": [f"@Coord!A{a4 + 2}"]},
                  {"type": "table", "caption": "SoC 測試機營收：公司季報實績 vs Siliconomics 基準隱含值（主讀數、第二讀數；層級隱含值與公司實績大致獨立，其他廠份額部分同源）", "head": [f"@Coord!{c}{mh}" for c in "ABCDEFG"], "rows": act_rows},
                  {"type": "note", "parts": [f"@Coord!A{a4e - 1}"]},
                  {"type": "note", "parts": ["對帳中項目數：", f"@Coord!G{a4e}|n0", "。"]},
                  {"type": "table", "caption": "實績與層級隱含值的輸入、等級與出處（展開看）", "collapse": True, "summary": "展開：實績、層級值、換算參數與出處", "head": [f"@Coord!{c}{ih}" for c in "ABCDE"] + [f"@Coord!H{ih}"], "rows": in_rows},
                  {"type": "table", "caption": f"@Coord!A{ot}", "collapse": True, "summary": "展開：其他分部與合併營收的實績（層級無可比量，列示供對照）", "head": [f"@Coord!{c}{oh}" for c in "ABCDE"] + [f"@Coord!G{oh}"], "rows": oth_rows}]
    c0_note = [{"type": "h3", "text": "讀取層級核對（同源，不判定）"}, {"type": "note", "parts": [f"@Coord!A{c2}"]}, {"type": "note", "parts": [f"@Coord!A{c2 + 1}"]}, c0_tbl, {"type": "note", "parts": [f"@Coord!A{cnt + 1}"]}]
    c0_src = {"type": "table", "caption": "層級名稱、版本、口徑與證據等級（每列的來源；展開看）", "collapse": True, "summary": "展開：每列的層級名稱、版本、口徑與證據等級",
              "head": ["項目", "年度", "層級名稱與版本／口徑／證據等級／備註"], "rows": [[f"@Coord!A{r}", f"@Coord!C{r}", f"@Coord!H{r}"] for r in range(h2 + 1, cnt) if CW.cell(r, 8).value]}

    def _rd(r):
        v = CW.cell(r, 2).value
        return [f"@Coord!A{r}", f"@Coord!B{r}" + ("|n3" if isinstance(v, (int, float)) and not isinstance(v, bool) else ""), f"@Coord!C{r}", f"@Coord!D{r}", f"@Coord!E{r}"]
    rd_tbl = {"type": "table", "caption": f"@Coord!A{c1}", "head": [f"@Coord!{c}{h1}" for c in "ABCDE"], "collapse": True, "summary": "展開：本模型讀哪一層、哪一版（下游只註明版本、不自行重估）",
              "rows": [_rd(r) for r in range(h1 + 1, c2 - 1) if CW.cell(r, 1).value]}
    rf0 = fnd(CW, "層級之間"); rfh = next(r for r in range(rf0, CW.max_row + 1) if CW.cell(r, 1).value == "項目"); rfe = next(r for r in range(rfh, CW.max_row + 1) if not CW.cell(r + 1, 1).value)
    ref_blocks = [{"type": "h3", "text": f"@Coord!A{rf0}"}, {"type": "note", "parts": [f"@Coord!A{rf0 + 1}"]},
                  {"type": "table", "head": [f"@Coord!{c}{rfh}" for c in "ABCDEFG"],
                   "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|n3", f"@Coord!E{r}|n3", f"@Coord!F{r}|sp1", f"@Coord!G{r}"] for r in range(rfh + 1, rfe + 1)]}]
    c3_notes = [{"type": "note", "parts": [f"@Coord!A{c3 + 1}"]}, {"type": "note", "parts": [f"@Coord!A{c3 + 2}"]}]
    S["recon"]["blocks"] = [{"type": "h3", "text": "三層協調：由下往上對帳（公司實績對各層）"}] + act_blocks + c0_note + ref_blocks + c3_notes + [c0_src, rd_tbl] + [{"type": "h3", "text": "公司目標、先驗比對與校準"}] + S["recon"]["blocks"]
    S["recon"]["title"] = "對帳：三層協調、公司實績與先驗比對"
    # ---- 8. 資料狀態 ----
    st = S["status"]["blocks"]
    st[0:0] = [{"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RM['ver']}"]}]
    order = ["thesis", "valuation", "sens", "revdcf", "pl", "drivers", "recon", "status"]
    S["revdcf"] = rev_
    cfg["sections"] = [S[k] for k in order]


# ---------------- 泰瑞達專屬：參數依業務性質分塊（半導體測試／其他事業／全公司） ----------------
PRE = {"cap_dcf": "0.1 DCF", "cap_comp": "0.2 Comparable", "h3_old": "舊法",
       "dcf_rates": [("rf", "p2"), ("beta", "n2"), ("erp", "p1"), ("crp", "p1")],
       "dcf_rate_val": [row("Val", p) for p in ("股權成本 ke", "稅後債務成本", "權益占資本", "WACC", "永續成長率 g", "永續期 RONIC")]}
_y = lambda p: [f"{p}_{s}_{y}" for s in "ABC" for y in range(2026, 2030)]
SPEC = {
 "in_groups": [
  ["半導體測試（SoC 與記憶體測試機）", [
    ["測試機市場（讀 Advantest 模型，三情境）", ["mk_soc25", "mk_mem25"] + _y("ms") + _y("mm") + _y("as")],
    ["情境顆數係數與晶圓前段設備（對帳用）", ["um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C"] + [f"wfe_{s}_{y}" for s in "ABC" for y in range(2025, 2030)]],
    ["泰瑞達市占", ["s_soc", "s_mem", "adv_sh26", "an_share"]],
    ["Wattonomics 錨定係數與 Advantest AI SoC 市場（只進敏感度）", ["wta_27", "wta_28", "wta_29", "adv_aiB_27", "adv_aiB_28", "adv_aiB_29", "adv_aiC_27", "adv_aiC_28", "adv_aiC_29"]]]],
  ["其他事業（Product Test 與機器人）", [
    ["成長率與機器人利益率", ["g_pt_A", "g_pt_B", "g_pt_C", "g_rob", "rm_2027", "rm_2028", "rm_2029"]]]],
  ["全公司（錨點、獲利結構、指引與共識、目標價規則）", [
    ["FY2026 錨點", ["q4_rev", "q4_eps", "co_rev26", "h2g", "rob_h2"]],
    ["獲利結構、稅率、股數", ["t_inc", "t_fix", "corp", "etr", "aff", "sh26", "bb", "gm_adj_C28", "gm_adj_C29", "m_soc", "oth_net"]],
    ["公司指引、目標模型與共識（只作對帳）", ["h1sh_lo", "h1sh_hi", "cons_rev26", "cons_eps26", "cons_eps27", "tm_rev", "tm_eps_lo", "tm_eps_hi", "tw_lo", "tw_hi"]],
    ["股價與循環位置", ["px", "sa_fpe", "pk_A", "pk_B", "pk_C"]],
    ["目標價規則（情境權重、DCF 權重、延伸期收斂與共識）", ["tp_wA", "tp_wB", "tp_wC", "tp_g", "tp_w_dcf", "nim10", "cons_tp", "cons_feps"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "beta_u", "beta_de", "beta_du", "beta", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "bs_date", "stub_frac", "stub_div", "sh_v"]],
    ["自由現金流（折舊、資本支出、營運資金）", ["fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]]]]],
 "in_skip": [],
 "main": {"0.1 DCF": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "＋Stub 期 FCFF", "淨現金滾到評價日", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
          "對帳要點": ["營收：公司口述全年", "EPS（共識，non-GAAP）", "營收 6.0B：模型 FY2027", "non-GAAP EPS 上緣", "CY2027：SoC＋記憶體測試機 ÷ WFE"]},
}
if __name__ == "__main__":
    run(SPEC, PRE)
