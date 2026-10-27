"""一次性工具（v0.6）：把研究檔（主要客戶、前十大股東）轉成公司檔案的兩個段落，插在「近期重大事件」之前。
用法：python3 markets/invest_map/tools_add_holders.py <研究檔資料夾> [代號 ...]
研究檔格式：## B. 主要客戶／## C. 前十大股東／## 來源（[n] 編號引用）。所有內容標 [外部・未經本專案檢核：來源，日期]。"""
import os, re, sys
HERE = os.path.dirname(os.path.abspath(__file__))
PROF = os.path.join(HERE, "profiles")
MAP = {"GOOGL": "GOOGL", "2802": None, "3189": None}

def sec(t, name):
    m = re.search(r"^## " + name + r".*?$(.*?)(?=^## |\Z)", t, re.S | re.M)
    return m.group(1).strip() if m else ""

def sources(t):
    out = {}
    for ln in sec(t, "來源").splitlines():
        m = re.match(r"\s*(\d+)\.\s*\[(.+?)\]\((.+?)\)\s*(?:（(.+?)）)?", ln)
        if m:
            out[m.group(1)] = (m.group(2), m.group(3), m.group(4) or "日期未載")
    return out

def tagify(s, src):
    def rep(m):
        ns = re.findall(r"\d+", m.group(0))
        parts = []
        for n in ns:
            if n in src:
                title, _, d = src[n]
                title = re.sub(r"[\[\]]", "", title)[:60]
                parts.append(f"{title}，{d}")
        return f" [外部・未經本專案檢核：{'；'.join(parts)}]" if parts else ""
    s = re.sub(r"(?:\s*\[\d+\](?:\s*\[\d+\])*)", rep, s)
    if "[外部" not in s and s.strip() and not s.rstrip().endswith(("：", ":")):
        s = s.rstrip() + " [外部・未經本專案檢核]"
    return s

def bullets(body, src):
    out = []
    for ln in body.splitlines():
        if not ln.strip() or ln.strip().startswith("|"):
            continue
        t = re.sub(r"^\s*[-*]\s*", "", ln).strip()
        t = t.replace("**", "")
        out.append("- " + tagify(t, src))
    return out

def holders(body, src):
    rows = [l for l in body.splitlines() if l.strip().startswith("|")]
    pre = [l for l in body.splitlines() if l.strip() and not l.strip().startswith("|")]
    out = []
    if len(rows) >= 3:
        hdr = [c.strip() for c in rows[0].strip("|").split("|")]
        def col(*ks):
            return next((i for i, h in enumerate(hdr) if any(k in h for k in ks)), None)
        ir, ih, ip, idt, isrc = col("排名"), col("股東"), col("比例", "%"), col("日期"), col("來源")
        out += ["| 排名 | 股東 | 持股比例 | 標記 |", "|---|---|---|---|"]
        for r in rows[2:]:
            c = [x.strip() for x in r.strip().strip("|").split("|")]
            c += [""] * (len(hdr) - len(c))
            g = lambda i: c[i] if i is not None and i < len(c) else ""
            refs = g(isrc)
            d = g(idt)
            t = tagify(refs, src).strip() if refs else "[外部・未經本專案檢核]"
            if d and "外部" in t and d not in t:
                t = t.replace("[外部・未經本專案檢核", f"[外部・未經本專案檢核：資料日 {d}；", 1).replace("；：", "；").replace("；]", "]")
            out.append(f"| {g(ir)} | {g(ih).replace('**','')} | {g(ip)} | {t} |")
        out.append("")
    else:
        out.append("- 前十大股東：未找到 [外部・未經本專案檢核]")
    out += bullets("\n".join(pre), src)
    return out

def build(code, rfile):
    t = open(rfile, encoding="utf-8").read()
    src = sources(t)
    cust = bullets(sec(t, "B"), src) or ["- 未找到 [外部・未經本專案檢核]"]
    hold = holders(sec(t, "C"), src)
    return "## 主要客戶\n" + "\n".join(cust) + "\n\n## 前十大股東\n" + "\n".join(hold) + "\n\n"

def apply(code, rdir):
    pf = os.path.join(PROF, f"{code}.md")
    rf = os.path.join(rdir, f"{code}.md")
    if not (os.path.exists(pf) and os.path.exists(rf)):
        return False
    p = open(pf, encoding="utf-8").read()
    p = re.sub(r"^## 主要客戶$.*?(?=^## 近期重大事件)", "", p, flags=re.S | re.M)
    new = build(code, rf)
    assert "## 近期重大事件" in p, code
    p = p.replace("## 近期重大事件", new + "## 近期重大事件", 1)
    open(pf, "w", encoding="utf-8").write(p)
    return True

if __name__ == "__main__":
    rdir = sys.argv[1]
    codes = sys.argv[2:] or [f[:-3] for f in os.listdir(PROF) if f.endswith(".md") and not f.startswith("00_")]
    done = [c for c in codes if apply(c, rdir)]
    print("已更新", len(done), "；未更新", sorted(set(codes) - set(done)))
