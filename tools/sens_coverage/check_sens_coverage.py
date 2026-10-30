# 個股模型敏感度覆蓋核對（全 repo 共用；規則見 docs/plan/個股模型原則_命題加營收分部.md 第 9 節，本註解與第 9 節一致）。
#
# 起因：2026-10-30 Meta v0.1 與 SpaceX v0.1 的獨立核對都發現 In 頁有低／高區間的鍵沒有對應 Sens 案例；
#       其中 Tokenomics 每 MW 持有成本（IF_HoldEcon，he_h／he_b2／he_b3／he_vr）在讀法 2 下會讓評等跨線或目標價 −38%～+99%。
#
# 規則：
#   1. 待涵蓋的鍵＝In 頁（「情境驅動對照」一節以上）A 欄有鍵、且 D（敏感度低）或 F（敏感度高）有數值並與 E（基準）不同者。
#   2. Sens 頁表頭＝前 40 列中第一個至少三格非空、且有「…目標價」欄名（20 字以內）的列（找不到＝失敗）；「鍵」欄＝表頭文字以「鍵」開頭的欄。
#      案例列＝表頭以下（到「敏感度「…」」逐案明細區為止）在任一「…目標價」欄有數值、且 A 欄不是「讀法：」「註：」「說明：」「備註：」「檢查：」（開頭字後接冒號）的列；
#      其他列（讀法、註記、空白）只當註記。
#   3. 涵蓋＝某案例列的「鍵」欄含該鍵，或該列 A、B 欄（驅動、設定）以獨立識別字寫出該鍵（例如「cip_ai＝0」「同時改 tpu_cf」）。
#      A、B 欄以「，；。（）」與換行切子句，含「不變／不改／維持／另見／不重解」的子句裡的鍵不算（例如「WACC +1pt（beta 不變）」）。
#      說明欄等其他欄提到的鍵不算涵蓋（例如 WACC 案說明「beta 維持基準不變」）。
#   4. 群組：IF_HoldEcon（鍵 he_h／he_b2／he_b3／he_vr，亦認 he_h100 等、tk_hold_*，及項目文字含「IF_HoldEcon｜」）須有一個案例列
#      「鍵」欄寫 he_all 或寫出全部有區間的成員鍵（只認鍵欄；A、B 欄與說明欄文字一律不算）；
#      沒有鍵欄的舊 Sens 頁，只認 A、B 欄明寫識別字 he_all。該列不得含否定字。本群組不得豁免，缺案例一律失敗。
#      其他群組可用 --groups 規則檔（JSON）增加，格式見 GROUPS_DEFAULT。
#   5. 豁免：Sens 頁註記（表頭以上各列，及案例區內的非案例列）或案例列文字中，有一句（以「。」「；」或換行分句）同時含
#      (a) 該鍵、(b)「不列／豁免／免列／不另列」之一、(c) 該字之後的「：」後至少 4 個字的理由，且 (d) 不含否定字（待補／尚未／未定／未做／未列／不得／缺少／缺案例；「未來」「缺口」不算）。
#      含鍵與豁免字但不符 (c)(d) 者列為「豁免不合格」，仍判未涵蓋。
#   6. 不進模型：In 頁該列「說明欄」（G 欄以後最後一個文字格）含完整片語「只作對照，不進模型」者不列入待涵蓋，
#      但須追引用鏈：引用 In 格（D／E／F）的公式（含加引號工作表名、包住目標格的範圍、整欄／整列參照、指向它們的定義名稱與同名定義名稱）
#      只可在 Recon、README、Sens（其下游也須全在這三頁），或 Val 頁的「顯示用」格（本身沒有被任何公式或定義名稱引用）；
#      INDIRECT／OFFSET 一律視為引用；其他引用＝失敗（「寫不進模型但被 X 引用」）。
#   7. 警告（不判失敗）：In 頁 IF_HoldEcon 列（同第 4 條成員）有基準值但 D、F 都空白——該有區間卻沒填，本規則無法觸發。
#   8. --require-keycol：Sens 頁沒有「鍵」欄即失敗（新模型的建置腳本加此旗標）。
# 適用：In 頁 E3＝「基準」、D3 含「低」、F3 含「高」者（舊模板如 Nebius、CoreWeave、WhiteFiber、OpenAI 無 In 頁三欄＝不適用）。
#       適用但無 Sens 頁＝失敗。
# 「候選案例列」提示：只用項目文字前綴比對，準確度差，僅供人工參考、列號可能錯指，不影響判定。
# 用法：
#   python3 tools/sens_coverage/check_sens_coverage.py <成品 xlsx> [--require-keycol] [--groups 規則.json]   # 單檔；失敗＝結束碼 1
#   python3 tools/sens_coverage/check_sens_coverage.py --all [--md 報告.md] [--groups 規則.json]  # 掃描全部 */dist/*.xlsx（每資料夾取檔名排序最後一檔），只報告
import glob, json, os, re, sys
import openpyxl

GROUPS_DEFAULT = [
    {"name": "IF_HoldEcon", "key_re": r"^(he_(h|b2|b3|vr|h100|gb200|gb300|vr200)|tk_hold_\w+)$", "label_re": r"IF_HoldEcon｜",
     "aliases": ["he_all"], "strict": True},
]
EXEMPT_WORDS = ("不列", "豁免", "免列", "不另列")
NEG_WORDS = ("待補", "尚未", "未定", "未做", "未列", "不得", "缺少", "缺案例")
STAY_WORDS = ("不變", "不改", "維持", "另見", "不重解")
NOT_MODEL_RE = re.compile(r"只作對照，不進模型")
NOT_MODEL_OK_SHEETS = {"Recon", "README", "Sens"}  # 對照頁：可再往下引用，但整條鏈都須留在這些頁（Val 另論）
REF_HINT = re.compile(r"^(cons_|c_|g_q\d|old_tp|tp_v\d)")
REF_LABEL = re.compile(r"共識|v\d+\.\d[\d.]*\s*基準目標價")
IDENT = r"(?<![A-Za-z0-9_]){}(?![A-Za-z0-9_])"
HDR_ROWS = 40
NOTE_ROW = re.compile(r"^\s*(讀法|註|說明|備註|檢查)\s*[：:]")


def num(x):
    return isinstance(x, (int, float)) and not isinstance(x, bool)


def has_tok(text, key):
    return re.search(IDENT.format(re.escape(key)), text) is not None


def has_neg(text):
    return any(w in text for w in NEG_WORDS)


def in_rows(wb):
    """回傳 None（不適用）或 [dict(key, label, row, d, e, f, note)]（「情境驅動對照」以上、A 欄為鍵的所有列）"""
    if "In" not in wb.sheetnames:
        return None
    ws = wb["In"]
    d3, e3, f3 = (str(ws.cell(3, c).value or "") for c in (4, 5, 6))
    if e3 != "基準" or "低" not in d3 or "高" not in f3:
        return None
    out = []
    for r in range(4, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if "情境驅動對照" in str(a or ""):
            break
        if not a or not re.fullmatch(r"[A-Za-z0-9_]+", str(a)):
            continue
        texts = [ws.cell(r, c).value for c in range(7, ws.max_column + 1)]
        texts = [t for t in texts if isinstance(t, str) and t.strip()]
        out.append({"key": str(a), "label": str(ws.cell(r, 2).value or ""), "row": r,
                    "d": ws.cell(r, 4).value, "e": ws.cell(r, 5).value, "f": ws.cell(r, 6).value, "note": texts[-1] if texts else ""})
    return out


def ranged(x):
    lo = num(x["d"]) and not (num(x["e"]) and abs(x["d"] - x["e"]) <= 1e-12)
    hi = num(x["f"]) and not (num(x["e"]) and abs(x["f"] - x["e"]) <= 1e-12)
    return lo or hi


def active_clauses(text):
    """以「，；。（）」與換行切子句，去掉含「不變／不改／維持／另見／不重解」的子句（這些子句提到的鍵沒有被動到）"""
    parts = re.split(r"[，,；;。（）()\n]", text)
    return " ".join(p for p in parts if not any(w in p for w in STAY_WORDS))


def sens_region(ws):
    """回傳 dict(hdr, keycol, notes[文字], cases[(列號, 鍵欄文字, 全列文字, A／B 欄（驅動、設定）文字)])；找不到表頭時 hdr＝None"""
    maxc = min(ws.max_column, 30)
    hdr = None
    for r in range(1, min(ws.max_row, HDR_ROWS) + 1):
        vals = [str(ws.cell(r, c).value or "").strip() for c in range(1, maxc + 1)]
        if sum(bool(v) for v in vals) >= 3 and any("目標價" in v and len(v) <= 20 for v in vals):
            hdr = r
            break
    if hdr is None:
        return {"hdr": None, "keycol": None, "notes": [], "cases": []}
    hv = [str(ws.cell(hdr, c).value or "").strip() for c in range(1, maxc + 1)]
    keycol = next((i + 1 for i, v in enumerate(hv) if v.startswith("鍵")), None)
    tpcols = [i + 1 for i, v in enumerate(hv) if "目標價" in v and len(v) <= 20]
    notes = [str(ws.cell(r, c).value) for r in range(1, hdr) for c in range(1, maxc + 1) if isinstance(ws.cell(r, c).value, str)]
    cases = []
    for r in range(hdr + 1, min(ws.max_row, 2000) + 1):
        vals = [ws.cell(r, c).value for c in range(1, maxc + 1)]
        if str(vals[0] or "").startswith("敏感度「"):  # Microsoft 式逐案明細區
            break
        txt = " ".join(str(v) for v in vals if isinstance(v, str))
        if any(num(ws.cell(r, c).value) for c in tpcols) and not NOTE_ROW.match(str(vals[0] or "")):
            kt = str(ws.cell(r, keycol).value or "") if keycol else ""
            ab = " ".join(active_clauses(str(ws.cell(r, c).value)) for c in (1, 2) if isinstance(ws.cell(r, c).value, str))
            cases.append((r, kt, txt, ab))
        elif txt:
            notes.append(txt)
    return {"hdr": hdr, "keycol": keycol, "notes": notes, "cases": cases}


def load_groups(path):
    g = list(GROUPS_DEFAULT)
    if path:
        g += json.load(open(path, encoding="utf-8"))["groups"]
    return g


REF_RE = re.compile(
    r"(?<![A-Za-z0-9_.$'!:])(?:(?:'(?P<qs>[^']+)'|(?P<s>[A-Za-z_][A-Za-z0-9_.]*))!)?"
    r"(?:\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)(?::(?:(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+))?(?![A-Za-z0-9_(])"
    r"|\$?(?P<fc1>[A-Z]{1,3}):\$?(?P<fc2>[A-Z]{1,3})(?![A-Za-z0-9_(])"
    r"|\$?(?P<fr1>\d+):\$?(?P<fr2>\d+)(?![\d]))")
DYN_RE = re.compile(r"\b(INDIRECT|OFFSET)\s*\(", re.I)


def col_idx(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n


def formula_hits(f, cur, targets):
    """公式 f（位於工作表 cur）是否引用 targets＝{(工作表, 欄號, 列號)} 中任一格。
    支援加引號的工作表名、範圍（含目標格即算）、整欄／整列參照；INDIRECT、OFFSET 一律視為引用（保守）。"""
    f2 = re.sub(r'"[^"]*"', "", f)
    if DYN_RE.search(f2):
        return True
    for m in REF_RE.finditer(f2):
        st = m.group("qs") or m.group("s") or cur
        for ts, tc, tr in targets:
            if st != ts:
                continue
            if m.group("fc1"):
                if col_idx(m.group("fc1")) <= tc <= col_idx(m.group("fc2")):
                    return True
            elif m.group("fr1"):
                if int(m.group("fr1")) <= tr <= int(m.group("fr2")):
                    return True
            else:
                c1, r1 = col_idx(m.group("c1")), int(m.group("r1"))
                c2, r2 = (col_idx(m.group("c2")), int(m.group("r2"))) if m.group("c2") else (c1, r1)
                if min(c1, c2) <= tc <= max(c1, c2) and min(r1, r2) <= tr <= max(r1, r2):
                    return True
    return False


def names_on(wbf, targets):
    """指向 targets 中任一格（含範圍包住）的定義名稱"""
    out = set()
    for n, dn in wbf.defined_names.items():
        try:
            dests = list(dn.destinations)
        except Exception:  # 名稱是運算式或引用其他名稱，openpyxl 無法解析：保守視為引用
            out.add(n)
            continue
        txt = str(dn.attr_text or "")
        if not dests and txt:  # 運算式名稱（例：In!$E$40*1）：以文字判斷，無法判斷的保守視為引用
            if formula_hits("=" + txt, "", targets) or not re.search(r"![$A-Z]", txt):
                out.add(n)
            continue
        for t, ref in dests:
            if formula_hits("=" + ref, t, targets):
                out.add(n)
    return out


def referencing_cells(wbf, targets, extra_names=()):
    """回傳引用 targets（直接、範圍、或經指向它們的定義名稱）的公式格 [(工作表, 座標)]"""
    names = names_on(wbf, targets) | set(extra_names)
    pats = [re.compile(IDENT.format(re.escape(n)), re.I) for n in names]
    hits = []
    for sh in wbf.worksheets:
        for row in sh.iter_rows():
            for c in row:
                v = c.value
                if not (isinstance(v, str) and v.startswith("=")):
                    continue
                if (sh.title, c.column, c.row) in targets:
                    continue
                v2 = re.sub(r'"[^"]*"', "", v)
                if formula_hits(v, sh.title, targets) or any(p.search(v2) for p in pats):
                    hits.append((sh.title, c.coordinate, c.column, c.row))
    return hits


def trace_not_model(wbf, targets, key):
    """追「不進模型」列的引用鏈，回傳違規格清單。
    Recon／README／Sens 的引用格可再被引用，但下游整條鏈都須留在這三頁；
    Val 的引用格須是顯示用（本身沒有被任何公式或定義名稱引用）；其他工作表（含 In 本身）一律違規。"""
    bad, seen, frontier, first = [], set(), set(targets), True
    for _ in range(50):
        if not frontier:
            break
        nxt = set()
        for sh, co, cc, rr in referencing_cells(wbf, frontier, [key] if first else []):
            if (sh, cc, rr) in seen:
                continue
            seen.add((sh, cc, rr))
            if sh in NOT_MODEL_OK_SHEETS:
                nxt.add((sh, cc, rr))
            elif sh == "Val":
                if referencing_cells(wbf, {(sh, cc, rr)}):
                    bad.append(f"{sh}!{co}（Val 格另被引用，非顯示用）")
            else:
                bad.append(f"{sh}!{co}")
        frontier, first = nxt, False
    else:
        bad.append("引用鏈超過 50 層（未追完）")
    return bad


def check(path, groups=None, require_keycol=False):
    """回傳 dict：applicable, keys（待涵蓋數）, missing [(鍵, 項目, 低, 基準, 高, 原因, 提示)], exempt [(鍵, 理由句)],
    group_cov {群組: 案例列}, not_model [鍵], warn [文字], keycol（Sens 頁有無鍵欄）, error"""
    groups = groups if groups is not None else load_groups(None)
    wb = openpyxl.load_workbook(path, data_only=True)
    allrows = in_rows(wb)
    res = {"applicable": allrows is not None, "keys": 0, "missing": [], "exempt": [], "group_cov": {}, "not_model": [],
           "warn": [], "keycol": None, "error": None}
    if allrows is None:
        return res

    def member(g, x):
        return bool(re.search(g["key_re"], x["key"])) or bool(g.get("label_re") and re.search(g["label_re"], x["label"]))

    # 第 7 條警告：IF_HoldEcon 有基準、無區間
    for x in allrows:
        for g in groups:
            if g.get("strict") and member(g, x) and num(x["e"]) and x["d"] is None and x["f"] is None:
                res["warn"].append(f"{g['name']} 列 {x['key']}（In 第 {x['row']} 列）有基準 {fmt(x['e'])} 但 D／F 空白：該有區間卻沒填，本規則無法觸發")
    ks = [x for x in allrows if ranged(x)]
    # 第 6 條：不進模型
    wbf = None
    keep = []
    for x in ks:
        if NOT_MODEL_RE.search(x["note"]):
            wbf = wbf or openpyxl.load_workbook(path, data_only=False)
            bad = trace_not_model(wbf, {("In", c, x["row"]) for c in (4, 5, 6)}, x["key"])
            if bad:
                res["missing"].append((x["key"], x["label"], x["d"], x["e"], x["f"],
                                       "寫「只作對照，不進模型」但被計算鏈引用：" + "、".join(bad[:5]), ""))
            else:
                res["not_model"].append(x["key"])
            continue
        keep.append(x)
    ks = keep
    res["keys"] = len(ks) + len(res["missing"])
    if "Sens" not in wb.sheetnames:
        res["error"] = "適用（In 頁有低／高欄）但無 Sens 頁"
        res["missing"] += [(x["key"], x["label"], x["d"], x["e"], x["f"], "無 Sens 頁", "") for x in ks]
        return res
    sr = sens_region(wb["Sens"])
    res["keycol"] = bool(sr["keycol"]) if sr["hdr"] else None
    if sr["hdr"] is None:
        res["error"] = f"Sens 頁前 {HDR_ROWS} 列找不到表頭（含「目標價」欄名的列）"
        res["missing"] += [(x["key"], x["label"], x["d"], x["e"], x["f"], "Sens 頁找不到表頭", "") for x in ks]
        return res
    if require_keycol and not sr["keycol"]:
        res["error"] = "Sens 頁沒有「鍵」欄（--require-keycol）"
    cases = sr["cases"]
    sents = []
    for t in sr["notes"] + [c[2] for c in cases]:
        sents += [s.strip() for s in re.split(r"[。；;\n]", t) if s.strip()]

    def row_keys(kt):
        return set(re.findall(r"[A-Za-z0-9_]+", kt))

    def covered(k):
        for r, kt, txt, ab in cases:
            if k in row_keys(kt) or has_tok(ab, k):
                return r
        return None

    def exemption(k):
        """回傳 ('ok', 句) / ('bad', 句) / None"""
        bad = None
        for s in sents:
            if not has_tok(s, k):
                continue
            pos = [s.find(w) for w in EXEMPT_WORDS if w in s]
            if not pos:
                continue
            after = s[min(pos):]
            m = re.search(r"[：:](.*)$", after)
            reason = re.sub(r"\s", "", m.group(1)) if m else ""
            if len(reason) >= 4 and not has_neg(s):
                return ("ok", s)
            bad = bad or s
        return ("bad", bad) if bad else None

    for g in groups:
        mem = [x["key"] for x in ks if member(g, x)]
        if not mem:
            continue
        for r, kt, txt, ab in cases:
            if has_neg(kt + " " + txt):
                continue
            ks_ = row_keys(kt)
            if any(a in ks_ for a in g.get("aliases", [])) or all(m in ks_ for m in mem) \
                    or (not sr["keycol"] and any(has_tok(ab, a) for a in g.get("aliases", []))):
                res["group_cov"][g["name"]] = r
                break
    for x in ks:
        k = x["key"]
        g = next((g for g in groups if member(g, x)), None)
        if g is not None:
            if g["name"] in res["group_cov"]:
                continue
            if not g.get("strict"):
                ex = exemption(k)
                if ex and ex[0] == "ok":
                    res["exempt"].append((k, ex[1][:160]))
                    continue
            single = [r for r, kt, txt, ab in cases if k in row_keys(kt) or has_tok(ab, k)]
            why = f"{g['name']} 群組缺「整組同時」案例（鍵欄寫 he_all 或全部成員鍵、且無否定字；A、B 欄文字不算）" + \
                  ("；只有單一成員案例：第 " + "、".join(map(str, single)) + " 列" if single else "") + ("；群組不得豁免" if g.get("strict") else "")
            res["missing"].append((k, x["label"], x["d"], x["e"], x["f"], why, ""))
            continue
        if covered(k):
            continue
        ex = exemption(k)
        if ex and ex[0] == "ok":
            res["exempt"].append((k, ex[1][:160]))
            continue
        why = "無案例、無豁免" if not ex else f"豁免不合格（缺「：理由」或含否定字）：{ex[1][:60]}"
        stem = re.split(r"[（(｜]", x["label"])[0].strip()
        cand = [r for r, kt, txt, ab in cases if len(stem) >= 3 and stem in txt]
        hs = ("候選案例列（僅供參考，項目文字前綴比對，列號可能錯指）：第 " + "、".join(map(str, cand[:4])) + " 列") if cand else ""
        if REF_HINT.search(k) or REF_LABEL.search(x["label"]):
            hs = (hs + "；" if hs else "") + "疑似對照用（共識／指引／舊版目標價）：若確實不進模型，可在 In 頁說明欄寫「只作對照，不進模型」"
        res["missing"].append((k, x["label"], x["d"], x["e"], x["f"], why, hs))
    return res


def fmt(v):
    return "—" if v is None else (f"{v:.4g}" if num(v) else str(v))


def failed(r):
    return bool(r["missing"] or r["error"])


def main(argv):
    gpath = argv[argv.index("--groups") + 1] if "--groups" in argv else None
    groups = load_groups(gpath)
    if argv and argv[0] == "--all":
        md = argv[argv.index("--md") + 1] if "--md" in argv else None
        root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
        lines = ["| 資料夾 | 成品 | 待涵蓋鍵 | 未涵蓋 | 其中有候選案例列（僅供參考） | 其中疑似對照用 | IF_HoldEcon | 豁免 | 已排除鍵（不進模型） | 警告 | Sens 鍵欄 | 未涵蓋清單 |",
                 "|---|---|---|---|---|---|---|---|---|---|---|---|"]
        na, n_ok, n_bad = [], 0, 0
        for d in sorted(glob.glob(os.path.join(root, "*", "dist"))):
            xs = sorted(glob.glob(os.path.join(d, "*.xlsx")))
            if not xs:
                continue
            x, co = xs[-1], os.path.basename(os.path.dirname(d))
            try:
                r = check(x, groups)
            except Exception as e:
                lines.append(f"| `{co}` | {os.path.basename(x)} | — | 讀檔失敗 | — | — | — | — | — | — | — | {e} |")
                n_bad += 1
                continue
            if not r["applicable"]:
                na.append(co)
                continue
            he = [m for m in r["missing"] if m[5].startswith("IF_HoldEcon")]
            hes = "缺（失敗）" if he else ("有（第 %d 列）" % r["group_cov"]["IF_HoldEcon"] if "IF_HoldEcon" in r["group_cov"] else "無此群組")
            miss = "、".join(m[0] for m in r["missing"]) + (f"（{r['error']}）" if r["error"] else "")
            n_like = sum("候選案例列" in m[6] for m in r["missing"])
            n_ref = sum("疑似對照用" in m[6] for m in r["missing"])
            kc = "有" if r["keycol"] else ("無" if r["keycol"] is False else "—")
            lines.append(f"| `{co}` | {os.path.basename(x)} | {r['keys']} | {len(r['missing'])} | {n_like} | {n_ref} | {hes} | {len(r['exempt'])} | "
                         f"{'、'.join(r['not_model']) or '—'} | {len(r['warn'])} | {kc} | {miss or '—'} |")
            n_ok += not failed(r)
            n_bad += failed(r)
        print("\n".join(lines))
        tail = f"\n適用 {n_ok + n_bad} 檔：全涵蓋 {n_ok}、有缺口 {n_bad}；不適用 {len(na)} 檔：" + "、".join(na)
        print(tail)
        if md:
            open(md, "w", encoding="utf-8").write("\n".join(lines) + "\n" + tail + "\n")
        return 0
    r = check(argv[0], groups, require_keycol="--require-keycol" in argv)
    if not r["applicable"]:
        print("敏感度覆蓋核對：不適用（無 In 頁低／基準／高三欄）")
        return 0
    for w in r["warn"]:
        print("警告（不判失敗）：" + w)
    for k, s in r["exempt"]:
        print(f"豁免 {k}：{s}")
    if r["not_model"]:
        print("In 頁說明欄寫「只作對照，不進模型」且未被計算鏈引用（不列入）：" + "、".join(r["not_model"]))
    if failed(r):
        print(f"敏感度覆蓋核對失敗：{len(r['missing'])}／{r['keys']} 個有低／高區間的鍵沒有 Sens 案例或合格豁免")
        if r["error"]:
            print("  -", r["error"])
        for k, lab, d, e, f, why, hint in r["missing"]:
            print(f"  - {k}（{lab[:30]}；{fmt(d)}／{fmt(e)}／{fmt(f)}）：{why}" + (f"；{hint}" if hint else ""))
        return 1
    g = "；IF_HoldEcon 以第 %d 列一案涵蓋" % r["group_cov"]["IF_HoldEcon"] if "IF_HoldEcon" in r["group_cov"] else ""
    print(f"敏感度覆蓋核對通過（{r['keys']} 個有低／高區間的鍵皆有案例或豁免；豁免 {len(r['exempt'])}{g}）")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
