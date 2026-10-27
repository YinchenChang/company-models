"""公司檔案 Markdown → 結構化資料（網站用）。

出處標記（[本專案：…]、[外部…] 等）不留在主文：數字改為懸浮看出處，並在各頁底部列出（Andy 10-25 11:36：可讀性最優先）。
只支援規格內用到的 Markdown 子集：front matter、## 標題、表格、項目符號、段落、**粗體**、[連結](url)。
所有文字先跳脫 HTML，再套用粗體與連結，避免檔案內容被當成網頁程式碼執行。
"""
import glob
import html
import os
import re

SECTIONS = ["一句話定位", "公司在做什麼", "主要產品與營收占比", "市場定位與產業位置", "管理層與經營風格",
            "同業與競爭態勢", "主要客戶", "前十大股東", "近期重大事件", "需要觀察的事項", "來源"]
# 「主要客戶」「前十大股東」（Andy 2026-10-27 要求）屬外部資料：一律標 [外部・未經本專案檢核：來源，日期]
TAGP = r"(?:本專案|外部|Verified|Interested-party|Derived|判斷|Analogy|Assumed|同上)"
# 一組出處：可帶全形括號，數個標記可用；、，或空白相連
TAGGRP = re.compile(r"(?:（\s*)?\[" + TAGP + r"[^\]]*\](?:\s*）)?(?:[\s；、，]*(?:（\s*)?\[" + TAGP + r"[^\]]*\](?:\s*）)?)*")
ONE = re.compile(r"\[(" + TAGP + r"[^\]]*)\]")
DATE = re.compile(r"\d{4}-\d{1,2}(?:-\d{1,2}|-xx)?|\d{1,2}/\d{1,2}")
NUM = re.compile(r"(?<![A-Za-z0-9.,/_-])[$＄]?\d[\d,]*(?:\.\d+)?(?:\s?(?:%|％|bps|bp|bn|tn|GWh|TWh|GW|MW|kW|GB|TB|nm|x|m|M|B|K|k|T|W|倍|億|兆|萬|美元|港元|元))?(?![A-Za-z0-9/_.,-])(?!\s?[年月日週季])")
YEAR = re.compile(r"^(?:19|20)\d\d$")
HTMLTAG = re.compile(r"(<[^>]+>)")
TIP_MARK = '<sup class="sm" tabindex="0" data-src="{a}" aria-label="出處">ⓘ</sup>'


def _src_text(grp):
    """標記群 → 純文字出處（多個以「；」相連）。"""
    return "；".join(m.group(1).strip() for m in ONE.finditer(html.unescape(grp)))


def _wrap_nums(seg, attr):
    """把段落中的數字包成可懸浮看出處的元素；HTML 標籤內的文字不動。回傳 (新段落, 是否有包到)。"""
    hit = False
    parts = HTMLTAG.split(seg)
    for k, part in enumerate(parts):
        if k % 2 == 1 or not part:
            continue
        masks = []
        def keep(m):
            masks.append(m.group(0)); return chr(0xE000 + len(masks) - 1)
        t = DATE.sub(keep, part)
        def rep(m):
            nonlocal hit
            x = m.group(0)
            if YEAR.match(x):
                return x
            hit = True
            return f'<span class="sv" tabindex="0" data-src="{attr}">{x}</span>'
        t = NUM.sub(rep, t)
        parts[k] = re.sub("[\uE000-\uEFFF]", lambda m: masks[ord(m.group(0)) - 0xE000], t)
    return "".join(parts), hit


def inline(s, notes=None):
    """文字 → HTML。出處標記不顯示在主文：改為數字上的懸浮提示（無數字時用小 ⓘ），並收進 notes 供頁底列出。"""
    s = html.escape(s.strip(), quote=False)
    s = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", s)
    s = re.sub(r"\[([^\]]+)\]\((https?://[^)\s]+)\)",
               lambda m: f'<a href="{html.escape(m.group(2), quote=True)}" target="_blank" rel="noopener">{m.group(1)}</a>', s)
    out, last = [], 0
    for m in TAGGRP.finditer(s):
        src = _src_text(m.group(0))
        if not src:
            continue
        before = s[last:m.start()]
        cut = before.rfind("。")
        head, seg = (before[:cut + 1], before[cut + 1:]) if cut >= 0 else ("", before)
        attr = html.escape(src, quote=True)
        seg2, hit = _wrap_nums(seg, attr)
        seg2 = seg2.rstrip()
        if not hit:
            seg2 += TIP_MARK.format(a=attr)
        out.append(head + seg2)
        if notes is not None and src not in notes:
            notes.append(src)
        last = m.end()
    out.append(s[last:])
    return "".join(out)


def block(lines, notes=None):
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
                rows = _fold_tag_col(rows)
                head = "".join(f"<th>{inline(c)}</th>" for c in rows[0])
                body = "".join("<tr>" + "".join(f"<td>{inline(c, notes)}</td>" for c in r) + "</tr>" for r in rows[1:])
                out.append(f'<div class="ptbl"><table><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>')
            continue
        if re.match(r"^\s*([-*]|\d+\.)\s+", ln):
            ordered = bool(re.match(r"^\s*\d+\.", ln))
            items = []
            while i < len(lines) and re.match(r"^\s*([-*]|\d+\.)\s+", lines[i]):
                items.append(re.sub(r"^\s*([-*]|\d+\.)\s+", "", lines[i]))
                i += 1
            tagn = "ol" if ordered else "ul"
            out.append(f"<{tagn}>" + "".join(f"<li>{inline(x, notes)}</li>" for x in items) + f"</{tagn}>")
            continue
        para = []
        while i < len(lines) and lines[i].strip() and not lines[i].lstrip().startswith(("|", "-", "*")) \
                and not re.match(r"^\s*\d+\.\s", lines[i]):
            para.append(lines[i].strip())
            i += 1
        out.append(f"<p>{inline(' '.join(para), notes)}</p>")
    return "".join(out)


def _fold_tag_col(rows):
    """表格的「標記」欄不顯示：把該欄的出處併到同列的數字欄（占比、市占、規模、金額；找不到就用前一欄）。"""
    hdr = rows[0]
    j = next((k for k, h in enumerate(hdr) if h.strip() == "標記"), None)
    if j is None:
        return rows
    tgt = next((k for k, h in enumerate(hdr) if re.search("占比|市占|規模|金額|數值|份額", h)), None)
    if tgt is None:
        tgt = max(0, j - 1)
    out = [hdr[:j] + hdr[j + 1:]]
    for r in rows[1:]:
        r = list(r) + [""] * (len(hdr) - len(r))
        if ONE.search(r[j]):
            r[tgt] = (r[tgt] + " " + r[j]).strip()
        out.append(r[:j] + r[j + 1:])
    return out


def parse(path):
    with open(path, encoding="utf-8") as fh:
        text = fh.read()
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
    notes = {}
    hn = notes.setdefault("基本資料", [])
    data = {"meta": {k: inline(v, hn) for k, v in meta.items()},
            "oneliner": inline(" ".join(l.strip() for l in sec.get("一句話定位", "").splitlines() if l.strip()), hn),
            "sections": {s: block(sec.get(s, "").splitlines(), notes.setdefault(s, [])) for s in SECTIONS if s != "一句話定位"},
            "notes": {k: v for k, v in notes.items() if v},
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
