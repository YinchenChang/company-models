"""公司檔案 Markdown → 結構化資料（網站用）。

只支援規格內用到的 Markdown 子集：front matter、## 標題、表格、項目符號、段落、**粗體**、[連結](url)。
所有文字先跳脫 HTML，再套用粗體與連結，避免檔案內容被當成網頁程式碼執行。
"""
import glob
import html
import os
import re

SECTIONS = ["一句話定位", "公司在做什麼", "主要產品與營收占比", "市場定位與產業位置", "管理層與經營風格",
            "同業與競爭態勢", "近期重大事件", "需要觀察的事項", "來源"]
TAG = re.compile(r"\[(Verified|Interested-party|Derived|判斷|本專案)[^\]]*\]")


def inline(s):
    s = html.escape(s.strip(), quote=False)
    s = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", s)
    s = re.sub(r"\[([^\]]+)\]\((https?://[^)\s]+)\)",
               lambda m: f'<a href="{html.escape(m.group(2), quote=True)}" target="_blank" rel="noopener">{m.group(1)}</a>', s)
    s = TAG.sub(lambda m: f'<span class="tg">{m.group(0)}</span>', s)
    return s


def block(lines):
    out, i = [], 0
    while i < len(lines):
        ln = lines[i].rstrip()
        if not ln.strip():
            i += 1
            continue
        if ln.lstrip().startswith("|"):
            rows = []
            while i < len(lines) and lines[i].lstrip().startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                if not all(re.fullmatch(r":?-{2,}:?", c) for c in cells if c):
                    rows.append(cells)
                i += 1
            if rows:
                head = "".join(f"<th>{inline(c)}</th>" for c in rows[0])
                body = "".join("<tr>" + "".join(f"<td>{inline(c)}</td>" for c in r) + "</tr>" for r in rows[1:])
                out.append(f'<div class="ptbl"><table><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>')
            continue
        if re.match(r"^\s*([-*]|\d+\.)\s+", ln):
            ordered = bool(re.match(r"^\s*\d+\.", ln))
            items = []
            while i < len(lines) and re.match(r"^\s*([-*]|\d+\.)\s+", lines[i]):
                items.append(re.sub(r"^\s*([-*]|\d+\.)\s+", "", lines[i]))
                i += 1
            tagn = "ol" if ordered else "ul"
            out.append(f"<{tagn}>" + "".join(f"<li>{inline(x)}</li>" for x in items) + f"</{tagn}>")
            continue
        para = []
        while i < len(lines) and lines[i].strip() and not lines[i].lstrip().startswith(("|", "-", "*")) \
                and not re.match(r"^\s*\d+\.\s", lines[i]):
            para.append(lines[i].strip())
            i += 1
        out.append(f"<p>{inline(' '.join(para))}</p>")
    return "".join(out)


def parse(path):
    text = open(path, encoding="utf-8").read()
    meta = {}
    m = re.match(r"^---\n(.*?)\n---\n", text, re.S)
    if m:
        for ln in m.group(1).splitlines():
            if ":" in ln:
                k, v = ln.split(":", 1)
                meta[k.strip()] = v.strip()
        text = text[m.end():]
    parts = re.split(r"^## +(.+)$", text, flags=re.M)
    sec = {}
    for title, body in zip(parts[1::2], parts[2::2]):
        sec[title.strip()] = body
    missing = [s for s in SECTIONS if s not in sec]
    data = {"meta": {k: inline(v) for k, v in meta.items()},
            "oneliner": inline(" ".join(l.strip() for l in sec.get("一句話定位", "").splitlines() if l.strip())),
            "sections": {s: block(sec.get(s, "").splitlines()) for s in SECTIONS if s != "一句話定位"},
            "missing": missing}
    return data


def load_all(folder):
    out = {}
    for p in sorted(glob.glob(os.path.join(folder, "*.md"))):
        name = os.path.basename(p)[:-3]
        if name.startswith("00_"):
            continue
        out[name] = parse(p)
    return out
