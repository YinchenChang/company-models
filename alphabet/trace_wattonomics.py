#!/usr/bin/env python3
"""追公式（三層協調 C0 的必做步驟；MAG 類樣板，Microsoft、Meta、SpaceX、Oracle 複製時改 ROOTS 與 WATCH）。
把 Wattonomics xlsx 中某玩家的輸出名稱沿公式一路追到輸入（In 頁或常數），傾印每格的公式與值，
並標出「讀到公司實績或目標」的輸入（同源旗標）。結果寫成 JSON，v039_post.py 的 Coord 頁追公式表直接貼這份輸出，不手寫。
用法：python3 alphabet/trace_wattonomics.py <Wattonomics xlsx> <輸出 json> [--md]
  ROOTS：要追的輸出名稱（Interface 頁儲存格；WIF_* 具名範圍）。
  WATCH：視為「公司實績或目標」的 In 頁列特徵（列文字含這些字串，或列鍵符合 KEYS）。換公司時改這兩處。
另做反向查詢：哪些公式讀了 act26_<玩家>（公司已公布實績），列出工作表與儲存格。
"""
import json
import re
import sys
import openpyxl

PLAYER = "googl"
ROOTS = [("WIF_Gap_Player_googl_2026_AddGW", "2026 年新增 IT GW"), ("WIF_Gap_Player_googl_2026_CapexAI", "2026 年 AI 資本支出"),
         ("WIF_Gap_Player_googl_2026_CapexNonAI", "2026 年非 AI 資本支出"), ("WIF_PlayerGW_YE_googl_2025", "2025 年底上線 IT GW"),
         ("WIF_PlayerGW_YE_googl_2026", "2026 年底上線 IT GW"), ("WIF_Gap_Player_googl_2026_CapexAIModel", "個股模型 AI 資本支出回寫（2026）")]
WATCH_TEXT = ("個股模型", "指引", "法說", "act26")
KEYS = re.compile(r"^(g(cap|capai|nai|ocf)_googl|cap_googl|act26_googl|w1_googl|obs_googl)$")
REF = re.compile(r"(?:'([^']+)'|([A-Za-z][A-Za-z0-9_]*))!\$?([A-Z]{1,3})\$?(\d+)|(?<![A-Za-z0-9_!$])\$?([A-Z]{1,3})\$?(\d+)(?![A-Za-z0-9_(])")
NAME = re.compile(r"(?<![A-Za-z0-9_!$.])([A-Za-z][A-Za-z0-9_]*_[A-Za-z0-9_]+)(?![A-Za-z0-9_(!])")


def main(xlsx, out, md=False):
    wb = openpyxl.load_workbook(xlsx)
    wv = openpyxl.load_workbook(xlsx, data_only=True)
    names = {n: wb.defined_names[n].attr_text for n in wb.defined_names}

    def resolve(n):
        t = names.get(n)
        if not t or "!" not in t or ":" in t:
            return None
        sh, c = t.split("!")
        return sh.strip("'"), c.replace("$", "")

    def in_flag(sh, c):
        if sh != "In":
            return ""
        r = int(re.sub(r"\D", "", c))
        row = [wb["In"].cell(r, k).value for k in range(1, 12)]
        txt = " ".join(str(x) for x in row if isinstance(x, str))
        if KEYS.match(str(row[0] or "")):
            return "公司實績／指引／個股模型寫回"
        return "人工確認（列文字含個股模型或指引字樣）" if any(w in txt for w in WATCH_TEXT) else ""

    seen, rows = {}, []

    def walk(sh, c, depth, root, via):
        key = (sh, c)
        if key in seen:
            return
        v = wb[sh][c].value
        val = wv[sh][c].value
        isf = isinstance(v, str) and v.startswith("=")
        label = ""
        if sh == "In":
            r = int(re.sub(r"\D", "", c))
            label = f"{wb['In'].cell(r, 1).value}｜{wb['In'].cell(r, 2).value}"
        seen[key] = len(rows)
        rows.append(dict(root=root, depth=depth, cell=f"{sh}!{c}", formula=v if isf else None, value=val, label=label, flag=in_flag(sh, c), leaf=(not isf) or sh == "In"))
        if not isf or sh == "In" or depth >= 14:
            return
        f = re.sub(r'"[^"]*"', "", v)
        for m in REF.finditer(f):
            if m.group(3):
                walk((m.group(1) or m.group(2)), f"{m.group(3)}{m.group(4)}", depth + 1, root, c)
            else:
                walk(sh, f"{m.group(5)}{m.group(6)}", depth + 1, root, c)
        for m in NAME.finditer(f):
            t = resolve(m.group(1))
            if t:
                walk(t[0], t[1], depth + 1, root, c)

    for nm, desc in ROOTS:
        t = resolve(nm)
        assert t, nm
        walk(t[0], t[1], 0, nm, None)
    # 反向：哪些公式讀公司實績 act26_<玩家>
    rev = []
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("=") and f"act26_{PLAYER}" in c.value:
                    rev.append(f"{ws.title}!{c.coordinate}")
    flagged = [r for r in rows if r["flag"]]
    res = dict(source=xlsx, roots=[dict(name=n, desc=d, cell="!".join(resolve(n))) for n, d in ROOTS], rows=rows, act26_readers=rev, flagged=[r["cell"] for r in flagged])
    json.dump(res, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
    if md:
        for r in rows:
            print(f"{'  ' * r['depth']}{r['cell']} = {r['formula'] or r['value']} => {r['value']} {r['label']} {r['flag']}")
        print("act26 readers:", rev)
    print(f"追公式：{len(rows)} 格；讀公司實績或目標的輸入 {len(flagged)} 格；act26_{PLAYER} 被 {len(rev)} 格讀取（{sorted({x.split('!')[0] for x in rev})}）")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], "--md" in sys.argv)
