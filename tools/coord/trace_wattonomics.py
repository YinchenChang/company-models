#!/usr/bin/env python3
"""追公式（三層協調 C0 的必做步驟；MAG 類與其他算力買賣方共用）。

把 Wattonomics xlsx 中某玩家的輸出名稱沿公式一路追到輸入，傾印每格的公式、值、上層格（parent），
標出「讀到公司實績或目標」的輸入（同源旗標），並反向列出哪些公式讀 act26_<玩家>（公司已公布實績）。
Coord 頁追公式表貼這份 JSON 輸出，不手寫。

用法：
  python3 tools/coord/trace_wattonomics.py <Wattonomics xlsx> <輸出 json> --player <玩家碼，例 googl> [--year 2026]
         [--roots "名稱:說明,名稱:說明"] [--manual "鍵=標記,..."] [--md]
  --player  玩家碼（Wattonomics 名稱 WIF_*_<碼>_* 的碼；旗標鍵 gcap_<碼>、cap_<碼>、act26_<碼>、w1_<碼>、obs_<碼> 由此組出，不寫死）。
  --roots   省略時追預設六個輸出：<年>新增 GW、AI 與非 AI 資本支出、前一年與當年底上線 GW、個股模型 AI 資本支出回寫。
  --manual  關鍵字誤觸但人工判定的鍵，例 cap_g_ext=不同源（預設已含 cap_g_ext、ai_sh）。

輸出每列：root、depth、cell、parent（上層格）、formula、value、label、flag（同源旗標）、
  status：生效＝值沿此路徑傳到輸出；條件讀取＝只出現在 IF 條件、或 SUMIFS 的條件欄／不符其他條件的列，不傳值；分支未生效＝IF 沒取到的分支。
IF 的條件用快取值遞迴計算（支援 ref<>""、ref=""、常數比較如 2026>=2026、INT(名稱)=年、AND／OR／NOT／N／ISNUMBER 等簡單函數）；
算不出來（不支援的函數、範圍）時兩支都當生效，須人工確認。
SUMIFS／COUNTIFS／SUMIF 依條件欄逐列判斷：文字相等條件（如 Pipeline 營運者欄＝玩家碼、鍵欄＝園區鍵）不符的列整列不追；數值或比較條件（年度、納入）不符的列標條件讀取。
Pipeline 與 FundRaw（10-K／10-Q 實績）的格子標旗標（園區列、公司實績）；其他範圍參照在這兩頁只展開「列內有玩家碼」的列。
與玩家無關的 In 頁列（鍵含其他玩家碼、且不含本玩家碼）不追；鍵含本玩家碼的跨玩家列（ctr_msft_ant、xc_nbis_msft、al_orcl_oai）保留。
限制：不支援整欄參照（A:A）；未支援的函數視為不確定。上線 GW 根的生效判斷若出現「不確定」，須人工確認。
"""
import argparse
import json
import re

import openpyxl

WATCH_TEXT = ("個股模型", "指引", "法說", "act26", "公告")
DEFAULT_MANUAL = {"cap_g_ext": "不同源", "ai_sh": "不同源"}
RANK = {"生效": 2, "條件讀取": 1, "分支未生效": 0}
FILTER_SHEETS = {"Pipeline": "園區列", "FundRaw": "公司實績（10-K／10-Q，FundRaw）"}
TOK = re.compile(r"""\s*(?:
 (?P<str>"(?:[^"]|"")*")
 |(?P<num>\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)
 |(?P<ref>(?:(?:'[^']+'|[A-Za-z][A-Za-z0-9_]*)!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?(?![A-Za-z0-9_(]))
 |(?P<fn>[A-Za-z][A-Za-z0-9_.]*(?=\())
 |(?P<name>[A-Za-z_][A-Za-z0-9_.]*)
 |(?P<op><>|<=|>=|[=<>&+\-*/^(),%:])
)""", re.X)
REFP = re.compile(r"^(?:(?:'([^']+)'|([A-Za-z][A-Za-z0-9_]*))!)?\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$")


class Unknown(Exception):
    pass


def col2n(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n


def n2col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def tokenize(f):
    pos, out = 0, []
    while pos < len(f):
        m = TOK.match(f, pos)
        if not m or m.end() == pos:
            if f[pos:].strip() == "":
                break
            raise Unknown(f"token@{pos}")
        pos = m.end()
        k = m.lastgroup
        out.append((k, m.group(k)))
    return out


class Parser:
    def __init__(self, toks):
        self.t, self.i = toks, 0

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None)

    def eat(self, v=None):
        k, x = self.peek()
        if v is not None and x != v:
            raise Unknown(f"expect {v}")
        self.i += 1
        return k, x

    def parse(self):
        n = self.cmp()
        if self.i != len(self.t):
            raise Unknown("trailing")
        return n

    def cmp(self):
        n = self.cat()
        while self.peek()[1] in ("=", "<>", "<", ">", "<=", ">=") and self.peek()[0] == "op":
            op = self.eat()[1]
            n = ("bin", op, n, self.cat())
        return n

    def cat(self):
        n = self.add()
        while self.peek() == ("op", "&"):
            self.eat()
            n = ("bin", "&", n, self.add())
        return n

    def add(self):
        n = self.mul()
        while self.peek()[0] == "op" and self.peek()[1] in ("+", "-"):
            op = self.eat()[1]
            n = ("bin", op, n, self.mul())
        return n

    def mul(self):
        n = self.unary()
        while self.peek()[0] == "op" and self.peek()[1] in ("*", "/", "^"):
            op = self.eat()[1]
            n = ("bin", op, n, self.unary())
        return n

    def unary(self):
        if self.peek()[0] == "op" and self.peek()[1] in ("-", "+"):
            op = self.eat()[1]
            n = self.unary()
            return ("neg", n) if op == "-" else n
        n = self.prim()
        while self.peek() == ("op", "%"):
            self.eat(); n = ("bin", "/", n, ("num", 100.0))
        return n

    def prim(self):
        k, x = self.eat()
        if k == "num":
            return ("num", float(x))
        if k == "str":
            return ("str", x[1:-1].replace('""', '"'))
        if k == "ref":
            return ("ref", x)
        if k == "fn":
            self.eat("(")
            args = []
            if self.peek() != ("op", ")"):
                while True:
                    args.append(self.cmp() if self.peek() not in (("op", ","), ("op", ")")) else ("str", ""))
                    if self.peek() == ("op", ","):
                        self.eat(); continue
                    break
            self.eat(")")
            return ("fn", x.upper(), args)
        if k == "name":
            if x.upper() in ("TRUE", "FALSE"):
                return ("num", 1.0 if x.upper() == "TRUE" else 0.0)
            return ("name", x)
        if (k, x) == ("op", "("):
            n = self.cmp()
            self.eat(")")
            return n
        raise Unknown("prim")


def main(xlsx, out, player, year, roots, manual, md=False):
    wb = openpyxl.load_workbook(xlsx)
    wv = openpyxl.load_workbook(xlsx, data_only=True)
    names = {n: wb.defined_names[n].attr_text for n in wb.defined_names}
    keys_re = re.compile(rf"^(g(cap|capai|nai|ocf)_{player}|cap_{player}|act26_{player}|w1_{player}|obs_{player})$")
    own_re = re.compile(rf"(^|_){player}(_|$)")
    in_keys = [str(wb["In"].cell(r, 1).value or "") for r in range(1, wb["In"].max_row + 1)] if "In" in wb.sheetnames else []
    players = {m.group(1) for k in in_keys for m in [re.match(r"^cap_([a-z0-9]+)$", k)] if m}
    players.discard(player)

    def resolve(n):
        t = names.get(n)
        if not t or "!" not in t or ":" in t:
            return None
        sh, c = t.split("!")
        return sh.strip("'"), c.replace("$", "")

    def cached(sh, c):
        return wv[sh][c].value

    # ---------- 求值（快取值）----------
    def num(x):
        if x is None or x == "":
            return 0.0
        if isinstance(x, bool):
            return float(x)
        if isinstance(x, (int, float)):
            return float(x)
        raise Unknown("text-as-number")

    def compare(op, a, b):
        if a is None:
            a = "" if isinstance(b, str) else 0.0
        if b is None:
            b = "" if isinstance(a, str) else 0.0
        if isinstance(a, bool): a = float(a)
        if isinstance(b, bool): b = float(b)
        if isinstance(a, str) != isinstance(b, str):
            sa = isinstance(a, str)
            if op == "=": return False
            if op == "<>": return True
            return (op in (">", ">=")) == sa
        if isinstance(a, str):
            a, b = a.lower(), b.lower()
        return {"=": a == b, "<>": a != b, "<": a < b, ">": a > b, "<=": a <= b, ">=": a >= b}[op]

    def ev(n, sh):
        k = n[0]
        if k in ("num", "str"):
            return n[1]
        if k == "ref":
            m = REFP.match(n[1])
            if m.group(5):
                raise Unknown("range")
            s2 = m.group(1) or m.group(2) or sh
            return cached(s2, f"{m.group(3)}{m.group(4)}")
        if k == "name":
            t = resolve(n[1])
            if not t:
                raise Unknown("name")
            return cached(*t)
        if k == "neg":
            return -num(ev(n[1], sh))
        if k == "bin":
            op = n[1]
            a, b = ev(n[2], sh), ev(n[3], sh)
            if op in ("=", "<>", "<", ">", "<=", ">="):
                return compare(op, a, b)
            if op == "&":
                return ("" if a is None else str(a)) + ("" if b is None else str(b))
            x, y = num(a), num(b)
            return {"+": x + y, "-": x - y, "*": x * y, "/": (x / y if y else (_ for _ in ()).throw(Unknown("div0"))), "^": x ** y}[op]
        if k == "fn":
            f, a = n[1], n[2]
            if f == "IF":
                c = ev(a[0], sh)
                if isinstance(c, str):
                    raise Unknown("if-str")
                return ev(a[1], sh) if (c if not isinstance(c, (int, float)) else c != 0) else (ev(a[2], sh) if len(a) > 2 else False)
            if f in ("AND", "OR"):
                v = [bool(ev(x, sh)) for x in a]
                return all(v) if f == "AND" else any(v)
            if f == "NOT":
                return not ev(a[0], sh)
            if f == "INT":
                import math
                return float(math.floor(num(ev(a[0], sh))))
            if f == "N":
                v = ev(a[0], sh)
                return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0
            if f == "ISNUMBER":
                v = ev(a[0], sh)
                return isinstance(v, (int, float)) and not isinstance(v, bool)
            if f == "ABS":
                return abs(num(ev(a[0], sh)))
            if f in ("MAX", "MIN"):
                v = [num(ev(x, sh)) for x in a]
                return max(v) if f == "MAX" else min(v)
        raise Unknown(k)

    # ---------- 參照收集 ----------
    def crit_match(cv, crit):
        if isinstance(crit, (int, float)) and not isinstance(crit, bool):
            return isinstance(cv, (int, float)) and cv == crit
        crit = "" if crit is None else str(crit)
        m = re.match(r"^(<>|<=|>=|=|<|>)(.*)$", crit)
        op, rhs = (m.group(1), m.group(2)) if m else ("=", crit)
        try:
            rv = float(rhs)
        except ValueError:
            rv = rhs
        if isinstance(rv, float) and not isinstance(cv, (int, float)):
            return op == "<>"
        return compare(op, cv, rv)

    def rng_cells(sh, m):
        c1, r1, c2, r2 = m.group(3), int(m.group(4)), m.group(5), int(m.group(6))
        return [(cc, rr) for cc in range(col2n(c1), col2n(c2) + 1) for rr in range(r1, r2 + 1)]

    def row_has_player(s2, rr):
        for cc in range(1, 8):
            v = wb[s2].cell(rr, cc).value
            if isinstance(v, str) and (v.lower() == player or own_re.search(v.lower())):
                return True
        return False

    def collect(n, sh, status, out):
        k = n[0]
        low = lambda a, b: a if RANK[a] <= RANK[b] else b
        if k == "ref":
            m = REFP.match(n[1])
            s2 = m.group(1) or m.group(2) or sh
            if s2 not in wb.sheetnames:
                return
            if not m.group(5):
                out.append((s2, f"{m.group(3)}{m.group(4)}", status)); return
            cells = rng_cells(s2, m)
            if len(cells) > 3000:
                return
            for cc, rr in cells:
                if wb[s2].cell(rr, cc).value is None:
                    continue
                if s2 in FILTER_SHEETS and not row_has_player(s2, rr):
                    continue
                out.append((s2, f"{n2col(cc)}{rr}", status))
        elif k == "name":
            t = resolve(n[1])
            if t:
                out.append((t[0], t[1], status))
        elif k == "neg":
            collect(n[1], sh, status, out)
        elif k == "bin":
            collect(n[2], sh, status, out); collect(n[3], sh, status, out)
        elif k == "fn":
            f, a = n[1], n[2]
            if f == "IF":
                try:
                    c = ev(a[0], sh)
                    if isinstance(c, str):
                        raise Unknown("if-str")
                    cond = bool(c)
                except Unknown:
                    cond = None
                collect(a[0], sh, low("條件讀取", status), out)
                for idx, br in ((1, True), (2, False)):
                    if idx < len(a):
                        st = status if cond is None or cond == br else low("分支未生效", status)
                        collect(a[idx], sh, st, out)
            elif f in ("SUMIFS", "COUNTIFS", "SUMIF"):
                sumifs_collect(f, a, sh, status, out)
            else:
                for x in a:
                    collect(x, sh, status, out)

    def sumifs_collect(f, a, sh, status, out):
        if f == "SUMIF":
            pairs, sum_r = [(a[0], a[1])], (a[2] if len(a) > 2 else a[0])
        elif f == "COUNTIFS":
            pairs, sum_r = [(a[i], a[i + 1]) for i in range(0, len(a) - 1, 2)], None
        else:
            sum_r, pairs = a[0], [(a[i], a[i + 1]) for i in range(1, len(a) - 1, 2)]
        try:
            rs = []
            for crn, cn in pairs:
                m = REFP.match(crn[1]) if crn[0] == "ref" else None
                if not m or not m.group(5):
                    raise Unknown("crit-range")
                rs.append((m.group(1) or m.group(2) or sh, m, ev(cn, sh)))
            sm = None
            if sum_r is not None:
                sm = REFP.match(sum_r[1]) if sum_r[0] == "ref" else None
                if not sm or not sm.group(5):
                    raise Unknown("sum-range")
        except Unknown:
            for x in a:
                collect(x, sh, status, out)
            return
        base = rs[0][1]
        for cc, rr in rng_cells(rs[0][0], base):
            ok_player, ok_other = True, True
            for s2, m, crit in rs:
                cv = cached(s2, f"{m.group(3)}{rr - int(base.group(4)) + int(m.group(4))}")
                match = crit_match(cv, crit)
                if isinstance(crit, str) and crit.strip() and not re.match(r"^(<>|<=|>=|=|<|>)", crit.strip()):
                    ok_player = ok_player and match      # 文字等於條件（玩家碼、園區鍵）：不符的列與本玩家無關，整列不追
                else:
                    ok_other = ok_other and match
            if not ok_player:
                continue
            st = status if ok_other else ("條件讀取" if RANK["條件讀取"] < RANK[status] else status)
            if sm is not None:
                out.append((rs[0][0] if False else (sm.group(1) or sm.group(2) or sh), f"{sm.group(3)}{rr - int(base.group(4)) + int(sm.group(4))}", st))
            for s2, m, crit in rs:
                out.append((s2, f"{m.group(3)}{rr - int(base.group(4)) + int(m.group(4))}", "條件讀取" if RANK["條件讀取"] < RANK[status] else status))

    # ---------- 輸入列資訊與旗標 ----------
    def in_info(sh, c):
        if sh != "In":
            return "", "", False
        r = int(re.sub(r"\D", "", c))
        row = [wb["In"].cell(r, k).value for k in range(1, 12)]
        key = str(row[0] or "")
        txt = " ".join(str(x) for x in row if isinstance(x, str))
        other = any(key.endswith("_" + p) or f"_{p}_" in key for p in players) and not own_re.search(key)
        label = f"{key}｜{row[1]}"
        if keys_re.match(key):
            return label, "公司實績／指引／個股模型寫回", other
        if key in manual:
            return label, f"人工確認：{manual[key]}", other
        if key.startswith("cp_") and own_re.search(key) and any(w in txt for w in ("公告", "公司")):
            return label, "公司公告園區（人工確認）", other
        if any(w in txt for w in WATCH_TEXT):
            return label, "人工確認（列文字含個股模型、指引或公告字樣）", other
        return label, "", other

    best, rows, order = {}, [], {}

    def walk(sh, c, depth, root, parent, status):
        key = (sh, c)
        if key in best and RANK[best[key]] >= RANK[status]:
            return
        best[key] = status
        v = wb[sh][c].value
        val = cached(sh, c)
        isf = isinstance(v, str) and v.startswith("=")
        label, flag, other = in_info(sh, c)
        if other:
            return
        if sh in FILTER_SHEETS and not flag:
            flag = FILTER_SHEETS[sh]
        if sh == "In" and (parent or "").startswith("Pipeline!") and not flag and own_re.search(label.split("｜")[0]):
            flag = "園區輸入列（人工確認：多為公司公告）"
        rec = dict(root=root, depth=depth, cell=f"{sh}!{c}", parent=parent, formula=v if isf else None, value=val, label=label, flag=flag, status=status, leaf=(not isf) or sh == "In")
        if key in order:
            rows[order[key]] = rec
        else:
            order[key] = len(rows); rows.append(rec)
        if not isf or sh == "In" or depth >= 16:
            return
        me = f"{sh}!{c}"
        try:
            ast = Parser(tokenize(v[1:])).parse()
        except Unknown:
            rec["note"] = "公式無法解析，所有參照視為生效，須人工確認"
            for s2, c2 in {(s, x) for s, x, _ in fallback_refs(v, sh)}:
                walk(s2, c2, depth + 1, root, me, status)
            return
        refs = []
        collect(ast, sh, status, refs)
        for s2, c2, st in refs:
            walk(s2, c2, depth + 1, root, me, st)

    def fallback_refs(f, sh):
        t = re.sub(r'"[^"]*"', "", f)
        res = []
        for m in re.finditer(r"(?<![A-Za-z0-9_!$.'])(?:(?:'([^']+)'|([A-Za-z][A-Za-z0-9_]*))!)?\$?([A-Z]{1,3})\$?(\d+)(?![A-Za-z0-9_(:])", t):
            s2 = m.group(1) or m.group(2) or sh
            if s2 in wb.sheetnames:
                res.append((s2, f"{m.group(3)}{m.group(4)}", "生效"))
        for m in re.finditer(r"(?<![A-Za-z0-9_!$.])([A-Za-z][A-Za-z0-9_]*_[A-Za-z0-9_]+)(?![A-Za-z0-9_(!])", t):
            tt = resolve(m.group(1))
            if tt:
                res.append((tt[0], tt[1], "生效"))
        return res

    for nm, desc in roots:
        t = resolve(nm)
        assert t, nm
        walk(t[0], t[1], 0, nm, None, "生效")
    rev = []
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("=") and re.search(rf"act26_{player}(?![A-Za-z0-9_])", c.value):
                    rev.append(f"{ws.title}!{c.coordinate}")
    flagged = [r["cell"] for r in rows if r["flag"] and r["flag"] not in FILTER_SHEETS.values()]
    campus = sorted({int(re.sub(r"[A-Z!a-z]", "", r["cell"].split("!")[1])) for r in rows if r["cell"].startswith("Pipeline!")})
    res = dict(source=xlsx, player=player, roots=[dict(name=n, desc=d, cell="!".join(resolve(n))) for n, d in roots], rows=rows, act26_readers=rev, flagged=flagged, campus_rows=campus,
               unparsed=[r["cell"] for r in rows if r.get("note")])
    with open(out, "w", encoding="utf-8") as fh:
        json.dump(res, fh, ensure_ascii=False, indent=1, default=str)
    if md:
        for r in rows:
            print(f"{'  ' * r['depth']}{r['cell']} <- {r['parent']} = {r['formula'] or r['value']} => {r['value']} {r['label']} [{r['status']}] {r['flag']}")
        print("act26 readers:", rev)
    inact = sum(1 for r in rows if r["status"] != "生效")
    print(f"追公式：{len(rows)} 格（生效 {len(rows) - inact}、條件讀取或分支未生效 {inact}）；旗標輸入 {len(flagged)} 格；命中 Pipeline 列 {len(campus)} 列；"
          f"act26_{player} 被 {len(rev)} 格讀取（{sorted({x.split('!')[0] for x in rev})}）；無法解析的公式 {len(res['unparsed'])} 格")


def default_roots(p, y):
    return [(f"WIF_Gap_Player_{p}_{y}_AddGW", f"{y} 年新增 IT GW"), (f"WIF_Gap_Player_{p}_{y}_CapexAI", f"{y} 年 AI 資本支出"),
            (f"WIF_Gap_Player_{p}_{y}_CapexNonAI", f"{y} 年非 AI 資本支出"), (f"WIF_PlayerGW_YE_{p}_{y - 1}", f"{y - 1} 年底上線 IT GW"),
            (f"WIF_PlayerGW_YE_{p}_{y}", f"{y} 年底上線 IT GW"), (f"WIF_Gap_Player_{p}_{y}_CapexAIModel", f"個股模型 AI 資本支出回寫（{y}）")]


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx"); ap.add_argument("out"); ap.add_argument("--player", required=True); ap.add_argument("--year", type=int, default=2026)
    ap.add_argument("--roots", default=""); ap.add_argument("--manual", default=""); ap.add_argument("--md", action="store_true")
    a = ap.parse_args()
    roots = [tuple(x.split(":", 1)) for x in a.roots.split(",") if x] or default_roots(a.player, a.year)
    man = dict(DEFAULT_MANUAL)
    man.update(dict(x.split("=", 1) for x in a.manual.split(",") if x))
    main(a.xlsx, a.out, a.player, a.year, roots, man, a.md)
