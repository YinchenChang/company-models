"""把 PR 中 Excel 的變動轉成可讀文字，供 PR 自動核對使用（2026-10-28 起）。

為什麼：模型 PR 的核心是 xlsx，過去自動核對只看文字檔，xlsx 一律略過。
做法：
1. 找出 PR 新增或修改的 xlsx（略過 archive/、_待刪除、/legacy）。
2. 找它的「舊版」：修改→同路徑的 base 版；改名→原路徑；新增→base 中同資料夾或其 archive/ 內、同模型名
   （去掉日期與版本號）的最新版；若模型名也改了（例 Kioxia模型→Kioxia個股模型），同資料夾在 PR 中被移走的
   xlsx 只有一個時以它為舊版。找不到舊版則只報錯誤格與基本統計。
   列對齊：依列標籤（A–D 欄第一個文字格，兩版皆唯一時）；同一列號的標籤改名時，該列改依列號逐格比較。
   計算值讀檔內快取值：若公式格大量沒有快取值（未經 Excel／LibreOffice 重算），輸出會警示。
3. 逐格比較兩版的計算值（openpyxl data_only）與公式文字，輸出：工作表、儲存格、列標籤、舊值 → 新值；
   公式有變的另列。新版另報錯誤格（#REF! 等）。
只讀，不改任何檔案；git 取檔用 `git fetch --depth=1 origin <sha>` 與 `git show`。
"""
import io
import os
import re
import subprocess

ERR = ("#DIV/0!", "#NAME?", "#REF!", "#VALUE!", "#N/A", "#NUM!", "#NULL!", "Err:")
STEM = re.compile(r"^(?:\d{8}_)?(.+?)_v[\d._A-Za-z-]+\.xlsx$")
SKIP_DIR = ("/archive/", "_待刪除", "/legacy")
TOL = 1e-9


def git(*args, binary=False):
    r = subprocess.run(["git", "-c", "core.quotepath=off", *args], capture_output=True)
    if r.returncode != 0:
        return None
    return r.stdout if binary else r.stdout.decode("utf-8", "replace")


def fetch(sha):
    git("fetch", "-q", "--depth=1", "origin", sha)


def ls_tree(sha):
    out = git("ls-tree", "-r", "--name-only", sha) or ""
    return [l for l in out.splitlines() if l.endswith(".xlsx")]


def blob(sha, path):
    return git("show", f"{sha}:{path}", binary=True)


def ver_key(name):
    m = re.search(r"_v([\d._]+)", name)
    nums = tuple(int(x) for x in re.findall(r"\d+", m.group(1))) if m else ()
    d = re.match(r"^(\d{8})_", name)
    return nums, d.group(1) if d else ""


def load(data, values):
    from openpyxl import load_workbook
    return load_workbook(io.BytesIO(data), data_only=values, read_only=False)


def cellmap(wb):
    out = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None:
                    out[(ws.title, c.coordinate)] = c.value
    return out


def row_labels(wb):
    """每列的標籤：A–D 欄中第一個文字格（給差異加上下文）。"""
    lab = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows(max_col=4):
            for c in row:
                if isinstance(c.value, str) and c.value.strip() and not c.value.startswith("="):
                    lab[(ws.title, c.row)] = c.value.strip()
                    break
    return lab


def same(a, b):
    if isinstance(a, bool) or isinstance(b, bool):
        return a == b
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= TOL * max(1.0, abs(a), abs(b))
    return a == b


def fmt(v, digits=6):
    if isinstance(v, float):
        return f"{v:.{digits}g}"
    s = str(v)
    return s if len(s) <= 60 else s[:57] + "…"


def find_old(path, status, prev, base_files, removed):
    if status == "modified":
        return path
    if status == "renamed" and prev:
        return prev
    m = STEM.match(os.path.basename(path))
    d = os.path.dirname(path)
    if not m:
        return None
    def near(f):   # 同資料夾，或 <d>/archive（舊版常被移到 archive）
        fd = os.path.dirname(f)
        return fd == d or fd == d + "/archive"
    cands = [f for f in set(base_files) | set(removed)
             if (STEM.match(os.path.basename(f)) or [None, None])[1] == m.group(1) and near(f)]
    cands = [f for f in cands if f != path]
    if not cands:   # 改名（例「Kioxia模型」→「Kioxia個股模型」）：同資料夾在 PR 中被移走的 xlsx 只有一個時，以它為舊版
        same_dir = [r for r in removed if os.path.dirname(r) == d and r != path]
        if len(same_dir) == 1:
            return same_dir[0]
    return max(cands, key=lambda f: ver_key(os.path.basename(f))) if cands else None


def aligned(vals, forms, wb, dup=frozenset()):
    """把每格的鍵由 (工作表, 座標) 改成 (工作表, 列標籤, 欄)：列標籤在該表唯一時用標籤，否則用列號。
    這樣插入或刪除列時，其餘列不會被誤判為整片變動。回傳 ({鍵: 值}, {鍵: 公式}, {鍵: 顯示用座標})。"""
    lab = row_labels(wb)
    count = {}
    for (sh, r), t in lab.items():
        count[(sh, t)] = count.get((sh, t), 0) + 1
    def key(sh, coord):
        col = re.sub(r"\d", "", coord)
        r = int(re.sub(r"\D", "", coord))
        t = lab.get((sh, r))
        return (sh, f"［{t}］" if t and count[(sh, t)] == 1 and (sh, t) not in dup else f"列{r}", col)
    kv, kf, where = {}, {}, {}
    for (sh, c), v in vals.items():
        k = key(sh, c); kv[k] = v; where[k] = f"{sh}!{c}"
    for (sh, c), v in forms.items():
        k = key(sh, c); kf[k] = v; where.setdefault(k, f"{sh}!{c}")
    return kv, kf, where


def short(k, n=40):
    """顯示用：把「［標籤］」截成 n 字以內並保留成對括號。"""
    t = k[1:-1] if k.startswith("［") and k.endswith("］") else k
    return "［" + (t if len(t) <= n else t[: n - 1] + "…") + "］"


def diff_pair(new_b, old_b, max_lines):
    nv, ov = load(new_b, True), load(old_b, True)
    nf, of = load(new_b, False), load(old_b, False)
    nvc, ovc = cellmap(nv), cellmap(ov)
    errs = [(k, v) for k, v in nvc.items() if isinstance(v, str) and v.startswith(ERR)]
    def dups(wb):
        c = {}
        for (sh, r), t in row_labels(wb).items():
            c[(sh, t)] = c.get((sh, t), 0) + 1
        return {k for k, v in c.items() if v > 1}
    dup = frozenset(dups(nv) | dups(ov))   # 任一版重複的標籤，兩版都改用列號，避免一邊用標籤一邊用列號
    nk, nkf, nwhere = aligned(nvc, cellmap(nf), nv, dup)
    ok, okf, owhere = aligned(ovc, cellmap(of), ov, dup)
    sheets_new = [ws.title for ws in nv.worksheets]
    sheets_old = [ws.title for ws in ov.worksheets]
    lines = []
    added = [x for x in sheets_new if x not in sheets_old]
    gone = [x for x in sheets_old if x not in sheets_new]
    if added or gone:
        lines.append(f"工作表：新增 {added or '無'}；移除 {gone or '無'}")
    lines.append(f"新版錯誤格：{len(errs)}" + ("（" + "、".join(f"{a}!{c}={v}" for (a, c), v in errs[:10]) + "）" if errs else ""))
    def rownum(where, k):
        w = where.get(k, "")
        return int(re.sub(r"\D", "", w.split("!")[-1])) if "!" in w else None
    rows_new = {(k[0], k[1]): rownum(nwhere, k) for k in nk if k[1].startswith("［")}
    rows_old = {(k[0], k[1]): rownum(owhere, k) for k in ok if k[1].startswith("［")}
    add_rows = sorted(set(rows_new) - set(rows_old))
    del_rows = sorted(set(rows_old) - set(rows_new))
    # 同表、同列號、標籤不同 → 視為標籤改名，改依列號比較（不當成整列新增／移除）
    renamed = []
    del_by_pos = {(sh, rows_old[(sh, t)]): (sh, t) for sh, t in del_rows}
    for sh, t in list(add_rows):
        old = del_by_pos.get((sh, rows_new[(sh, t)]))
        if old:
            renamed.append((old, (sh, t), rows_new[(sh, t)]))
            add_rows.remove((sh, t))
            del_rows.remove(old)
    ren_old = {(o[0], o[1]): f"列{r}" for o, _, r in renamed}
    ren_new = {(n[0], n[1]): f"列{r}" for _, n, r in renamed}
    def rekey(d, m):
        return {((k[0], m[(k[0], k[1])], k[2]) if (k[0], k[1]) in m else k): v for k, v in d.items()}
    ok, okf, owhere = rekey(ok, ren_old), rekey(okf, ren_old), rekey(owhere, ren_old)
    nk, nkf, nwhere = rekey(nk, ren_new), rekey(nkf, ren_new), rekey(nwhere, ren_new)
    if renamed:
        lines.append(f"標籤改名或同位置替換的列（依列號逐格比較；若實為替換，下列數值變動是兩個不同項目的差，不是調整）：{len(renamed)}：" + "、".join(f"{o[0]}{short(o[1])}→{short(n[1])}" for o, n, _ in renamed[:20]))
    nofc = sum(1 for k, f in nkf.items() if isinstance(f, str) and f.startswith("=") and nk.get(k) is None)
    nform = sum(1 for f in nkf.values() if isinstance(f, str) and f.startswith("="))
    if nform and nofc / nform > 0.05:
        lines.append(f"**警示：新版 {nofc}/{nform} 個公式格沒有快取值（可能未經 Excel／LibreOffice 重算），計算結果的比較不可靠。**")
    if add_rows or del_rows:
        lines.append(f"新增的列（依列標籤）：{len(add_rows)}" + ("：" + "、".join(f"{a}{short(t)}" for a, t in add_rows[:30]) if add_rows else "")
                     + f"；移除的列：{len(del_rows)}" + ("：" + "、".join(f"{a}{short(t)}" for a, t in del_rows[:30]) if del_rows else ""))
    num_in, num_out, txt, form, by_sheet = [], [], [], [], {}
    skip_rows = set(add_rows) | set(del_rows)
    for k in set(nk) | set(ok):
        if (k[0], k[1]) in skip_rows:
            continue   # 整列新增／移除：上面已列，不再逐格列
        a, b = ok.get(k), nk.get(k)
        fa, fb = okf.get(k), nkf.get(k)
        if isinstance(fa, str) and fa.startswith("=") or isinstance(fb, str) and fb.startswith("="):
            if fa != fb:
                form.append((k, fa, fb))
        if same(a, b):
            continue
        by_sheet[k[0]] = by_sheet.get(k[0], 0) + 1
        is_formula = isinstance(fb, str) and fb.startswith("=")
        if isinstance(a, (int, float)) and isinstance(b, (int, float)) and not isinstance(a, bool):
            rank = abs(b - a) / max(abs(a), abs(b), 1e-12)          # 排序用（0–1）
            pct = f"{(b - a) / abs(a):+.1%}" if a != 0 else "原值 0"  # 顯示用：帶方向的變動率
            (num_out if is_formula else num_in).append((rank, k, a, b, pct))
        else:
            txt.append((k, a, b))
    lines.append(f"值有變的格：{sum(by_sheet.values())}（" + "、".join(f"{a} {n}" for a, n in by_sheet.items())
                 + f"）；其中數值輸入 {len(num_in)}、數值計算結果 {len(num_out)}、文字或型別 {len(txt)}；公式文字有變 {len(form)}")

    def where(k):
        return nwhere.get(k) or owhere.get(k) or f"{k[0]}!{k[2]}"
    budget = max_lines
    for title, rows in (("數值輸入變動（非公式格；依變動幅度排序）", num_in), ("數值計算結果變動（公式格；依變動幅度排序）", num_out)):
        if not rows:
            continue
        rows.sort(key=lambda x: -x[0])
        n = min(len(rows), budget)
        lines.append(f"\n#### {title}：{len(rows)} 格" + ("" if n == len(rows) else f"（只列幅度最大的 {n}）"))
        for rank, k, a, b, pct in rows[:n]:
            d = 6 if fmt(a) != fmt(b) else 12   # 6 位有效數字看不出差異時改顯示 12 位
            lines.append(f"- {where(k)}{(k[1][:40] + '］') if k[1].startswith('［') and len(k[1]) > 41 else (k[1] if k[1].startswith('［') else '')}：{fmt(a, d)} → {fmt(b, d)}（變動率 {pct}）")
        budget = max(budget - n, 40)
    if txt:
        lines.append(f"\n#### 文字或型別變動：{len(txt)} 格" + ("" if len(txt) <= 40 else "（只列前 40）"))
        for k, a, b in sorted(txt, key=lambda x: where(x[0]))[:40]:
            lines.append(f"- {where(k)}{(k[1][:40] + '］') if k[1].startswith('［') and len(k[1]) > 41 else (k[1] if k[1].startswith('［') else '')}：{fmt(a)} → {fmt(b)}")
    if form:
        lines.append(f"\n#### 公式文字變動：{len(form)} 格" + ("" if len(form) <= 50 else "（只列前 50）"))
        for k, fa, fb in sorted(form, key=lambda x: where(x[0]))[:50]:
            lines.append(f"- {where(k)}：{fmt(fa)} → {fmt(fb)}")
    return "\n".join(lines)


def build(files, base_sha, head_sha, max_chars=60000, max_lines=150):
    """files：GitHub PR files API 的清單。回傳 (給模型的文字, 已處理的 xlsx 路徑集合, 備註清單)。"""
    xl = [f for f in files if f["filename"].lower().endswith(".xlsx") and f["status"] != "removed"
          and not any(x in f["filename"] for x in SKIP_DIR)]
    if not xl:
        return "", set(), []
    fetch(base_sha)
    fetch(head_sha)
    base_files = ls_tree(base_sha)
    removed = [f["filename"] for f in files if f["filename"].endswith(".xlsx") and
               (f["status"] == "removed" or any(x in f["filename"] for x in SKIP_DIR))]
    removed += [f.get("previous_filename") for f in files if f.get("previous_filename", "").endswith(".xlsx")]
    removed = [r for r in removed if r and r in set(base_files)]   # 只收 base 上真的存在的舊路徑（排除 PR 內移到 archive 的新路徑）
    parts, done, notes, used = [], set(), [], 0
    for f in xl:
        path = f["filename"]
        new_b = blob(head_sha, path)
        if not new_b:
            notes.append(f"{path}（讀不到新版）")
            done.add(path)
            continue
        old = find_old(path, f["status"], f.get("previous_filename"), base_files, [r for r in removed if r])
        old_b = blob(base_sha, old) if old else None
        try:
            if old_b:
                body = diff_pair(new_b, old_b, max_lines)
                head = f"\n## Excel 變動：{path}\n比較對象（舊版）：{old}\n"
            else:
                nv = load(new_b, True)
                cs = cellmap(nv)
                errs = [k for k, v in cs.items() if isinstance(v, str) and v.startswith(ERR)]
                body = f"工作表 {len(nv.worksheets)} 張、非空格 {len(cs)}；錯誤格 {len(errs)}；找不到舊版可比較（新檔）。"
                head = f"\n## Excel（新檔）：{path}\n"
        except Exception as e:
            notes.append(f"{path}（讀檔失敗 {type(e).__name__}）")
            done.add(path)   # 已列入 notes，避免重複
            continue
        block = head + body + "\n"
        if used + len(block) > max_chars:
            tail = "\n…（Excel 變動超過上限，以下截斷）\n"
            room = max_chars - used
            block = (block[: room - len(tail)] + tail) if room > len(tail) else ""
            parts.append(block)
            done.add(path)
            notes.append(f"{path}（Excel 變動部分截斷）")
            rest = [g["filename"] for g in xl[xl.index(f) + 1:]]
            notes += [f"{r}（超過 Excel 變動上限，未比較）" for r in rest]
            done.update(rest)   # 已列入 notes，避免與原「二進位略過」重複
            break
        parts.append(block)
        used += len(block)
        done.add(path)
    return "".join(parts), done, notes
