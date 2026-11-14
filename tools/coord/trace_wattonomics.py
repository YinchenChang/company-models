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

輸出每列：root（根名稱）、depth、cell、parent（上層格）、formula、value、label（In 頁列鍵與名稱）、flag（同源旗標）、
  status（生效＝值沿此路徑傳到輸出；條件讀取＝只出現在 IF 條件、不傳值；分支未生效＝IF 未取到的分支）、leaf。
範圍參照（如 Pipeline!$L$5:$L$150）逐格展開；Pipeline 頁只展開本玩家的園區列（鍵含 _<碼>_），命中的園區列標「園區列」，
其 In 頁輸入列含「公告」字樣者標「公司公告園區（人工確認）」。與本玩家無關的 In 頁列（鍵尾碼為其他玩家）不追、不列。
"""
import argparse
import json
import re
import sys

import openpyxl

WATCH_TEXT = ("個股模型", "指引", "法說", "act26", "公告")
DEFAULT_MANUAL = {"cap_g_ext": "不同源", "ai_sh": "不同源"}
REF = re.compile(r"(?<![A-Za-z0-9_!$.'])(?:(?:'([^']+)'|([A-Za-z][A-Za-z0-9_]*))!)?\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?(?![A-Za-z0-9_(])")
NAME = re.compile(r"(?<![A-Za-z0-9_!$.])([A-Za-z][A-Za-z0-9_]*_[A-Za-z0-9_]+)(?![A-Za-z0-9_(!])")
RANK = {"生效": 2, "條件讀取": 1, "分支未生效": 0}


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


def split_args(s):
    out, d, cur, q = [], 0, "", False
    for ch in s:
        if ch == '"':
            q = not q
        if not q:
            if ch == "(":
                d += 1
            elif ch == ")":
                d -= 1
            elif ch == "," and d == 0:
                out.append(cur); cur = ""; continue
        cur += ch
    out.append(cur)
    return out


def main(xlsx, out, player, year, roots, manual, md=False):
    wb = openpyxl.load_workbook(xlsx)
    wv = openpyxl.load_workbook(xlsx, data_only=True)
    names = {n: wb.defined_names[n].attr_text for n in wb.defined_names}
    keys_re = re.compile(rf"^(g(cap|capai|nai|ocf)_{player}|cap_{player}|act26_{player}|w1_{player}|obs_{player})$")
    in_keys = [str(wb["In"].cell(r, 1).value or "") for r in range(1, wb["In"].max_row + 1)] if "In" in wb.sheetnames else []
    players = {m.group(1) for k in in_keys for m in [re.match(r"^cap_([a-z0-9]+)$", k)] if m}
    players.discard(player)

    def resolve(n):
        t = names.get(n)
        if not t or "!" not in t or ":" in t:
            return None
        sh, c = t.split("!")
        return sh.strip("'"), c.replace("$", "")

    def blank(v):
        return v is None or v == ""

    def in_info(sh, c):
        if sh != "In":
            return "", "", False
        r = int(re.sub(r"\D", "", c))
        row = [wb["In"].cell(r, k).value for k in range(1, 12)]
        key = str(row[0] or "")
        txt = " ".join(str(x) for x in row if isinstance(x, str))
        other = any(key.endswith("_" + p) or f"_{p}_" in key for p in players)
        label = f"{key}｜{row[1]}"
        if keys_re.match(key):
            return label, "公司實績／指引／個股模型寫回", other
        if key in manual:
            return label, f"人工確認：{manual[key]}", other
        if key.startswith("cp_") and any(w in txt for w in ("公告", "公司")) and f"_{player}_" in key:
            return label, "公司公告園區（人工確認）", other
        if any(w in txt for w in WATCH_TEXT):
            return label, "人工確認（列文字含個股模型、指引或公告字樣）", other
        return label, "", other

    best, rows, order = {}, [], []

    def branch_roles(sh, c, f, val):
        """回傳 {參照原文: status}：只處理頂層 IF(cond, a, b)，其餘整式視為生效。"""
        m = re.match(r"=IF\((.*)\)$", f, re.S)
        if not m:
            return None
        args = split_args(m.group(1))
        if len(args) != 3:
            return None
        cond, a, b = args

        def evalb(x):
            x = x.strip()
            if x in ('""', "0", "1") or re.fullmatch(r'"[^"]*"|\d+(\.\d+)?', x):
                v = None if x == '""' else (float(x) if re.fullmatch(r"\d+(\.\d+)?", x) else x.strip('"'))
                return "eq" if (blank(v) and blank(val)) or (v == val) else "ne"
            mm = REF.fullmatch(x)
            if mm and not (mm.group(5)):
                s2 = mm.group(1) or mm.group(2) or sh
                v = wv[s2][f"{mm.group(3)}{mm.group(4)}"].value
                return "eq" if (blank(v) and blank(val)) or (v == val) else "ne"
            return "?"
        ea, eb = evalb(a), evalb(b)
        st = {"cond": "條件讀取", "a": "生效", "b": "生效"}
        if ea == "ne" and eb in ("eq", "?"):
            st["a"] = "分支未生效"
        if eb == "ne" and ea in ("eq", "?"):
            st["b"] = "分支未生效"
        if ea == "?" and eb == "eq" and blank(val):
            st["a"] = "分支未生效"
        return {"cond": cond, "a": a, "b": b}, st

    def refs_of(text, sh):
        t = re.sub(r'"[^"]*"', "", text)
        res = []
        for m in REF.finditer(t):
            s2 = m.group(1) or m.group(2) or sh
            if s2 not in wb.sheetnames:
                continue
            c1, r1 = m.group(3), int(m.group(4))
            if m.group(5):
                c2, r2 = m.group(5), int(m.group(6))
                cells = [f"{n2col(cc)}{rr}" for cc in range(col2n(c1), col2n(c2) + 1) for rr in range(r1, r2 + 1)]
                if len(cells) > 3000:
                    continue
                for cell in cells:
                    v = wb[s2][cell].value
                    if v is None:
                        continue
                    if s2 == "Pipeline":
                        key = str(wb[s2][f"A{int(re.sub(r'[A-Z]', '', cell))}"].value or "")
                        if f"_{player}_" not in key:
                            continue
                    res.append((s2, cell))
            else:
                res.append((s2, f"{c1}{r1}"))
        for m in NAME.finditer(t):
            tt = resolve(m.group(1))
            if tt:
                res.append(tt)
        return res

    def walk(sh, c, depth, root, parent, status):
        key = (sh, c)
        if key in best and RANK[best[key]] >= RANK[status]:
            return
        best[key] = status
        v = wb[sh][c].value
        val = wv[sh][c].value
        isf = isinstance(v, str) and v.startswith("=")
        label, flag, other = in_info(sh, c)
        if other:
            best.pop(key, None); best[key] = status
            return
        if sh == "Pipeline" and not flag:
            flag = "園區列"
        if sh == "In" and (parent or "").startswith("Pipeline!") and not flag:
            flag = "園區輸入列（人工確認：多為公司公告）"
        rec = dict(root=root, depth=depth, cell=f"{sh}!{c}", parent=parent, formula=v if isf else None, value=val, label=label, flag=flag, status=status, leaf=(not isf) or sh == "In")
        if key in order:
            rows[order.index(key)] = rec
        else:
            order.append(key); rows.append(rec)
        if not isf or sh == "In" or depth >= 16:
            return
        br = branch_roles(sh, c, v, val)
        me = f"{sh}!{c}"
        if br is None:
            for s2, c2 in refs_of(v, sh):
                walk(s2, c2, depth + 1, root, me, status)
        else:
            parts, st = br
            for k in ("cond", "a", "b"):
                sub = st[k] if RANK[st[k]] <= RANK[status] else status
                for s2, c2 in refs_of(parts[k], sh):
                    walk(s2, c2, depth + 1, root, me, sub)

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
    flagged = [r["cell"] for r in rows if r["flag"] and r["flag"] != "園區列"]
    campus = sorted({r["cell"].split("!")[1] for r in rows if r["cell"].startswith("Pipeline!")} and {re.sub(r"[A-Z]", "", r["cell"].split("!")[1]) for r in rows if r["cell"].startswith("Pipeline!")}, key=int)
    res = dict(source=xlsx, player=player, roots=[dict(name=n, desc=d, cell="!".join(resolve(n))) for n, d in roots], rows=rows, act26_readers=rev, flagged=flagged, campus_rows=campus)
    json.dump(res, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
    if md:
        for r in rows:
            print(f"{'  ' * r['depth']}{r['cell']} <- {r['parent']} = {r['formula'] or r['value']} => {r['value']} {r['label']} [{r['status']}] {r['flag']}")
        print("act26 readers:", rev)
    inact = sum(1 for r in rows if r["status"] != "生效")
    print(f"追公式：{len(rows)} 格（生效 {len(rows) - inact}、條件讀取或分支未生效 {inact}）；旗標輸入 {len(flagged)} 格；命中 Pipeline 園區列 {len(campus)} 列；act26_{player} 被 {len(rev)} 格讀取（{sorted({x.split('!')[0] for x in rev})}）")


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
