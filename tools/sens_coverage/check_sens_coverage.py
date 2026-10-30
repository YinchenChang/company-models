# 個股模型敏感度覆蓋核對（全 repo 共用；規則見 docs/plan/個股模型原則_命題加營收分部.md 第 9 節）。
#
# 起因：2026-10-30 Meta v0.1 與 SpaceX v0.1 的獨立核對都發現 In 頁有低／高區間的鍵沒有對應 Sens 案例；
#       其中 Tokenomics 每 MW 持有成本（IF_HoldEcon，he_h／he_b2／he_b3／he_vr）在讀法 2 下會讓評等跨線或目標價 −38%～+99%。
#
# 規則：
#   1. 待涵蓋的鍵＝In 頁（「情境驅動對照」一節以上）A 欄有鍵、且 D（敏感度低）或 F（敏感度高）有數值並與 E（基準）不同者。
#      In 頁該列任一文字格註明「不進模型」（低／高只作對照，例如市場共識目標價的最低／最高）者不列入。
#   2. 涵蓋＝Sens 頁案例區（表頭列以下，到「敏感度「…」」明細區或連續 10 列空白為止）任一列：
#      (a) 「鍵」欄（表頭文字為「鍵」的欄）含該鍵；或 (b) 任一文字格以獨立識別字形式提到該鍵（例如說明寫「cip_ai＝0」）。
#   3. 群組：同一群組的鍵以一案同時動整組即可。預設群組 IF_HoldEcon（鍵 he_h／he_b2／he_b3／he_vr（Oracle 寫法 he_h100 等），或項目文字含「IF_HoldEcon｜」）：
#      案例列含別名 he_all、或文字同時含「IF_HoldEcon」與「四世代／同時」、或同一列含全部成員鍵，才算涵蓋。
#      群組可用 --groups 規則檔（JSON）增加，格式見 GROUPS_DEFAULT。
#   4. 豁免：Sens 頁表頭以上註記（A1／A2 等）或案例區任一文字格中，有一句（以「。」「；」或換行分句）同時含該鍵與
#      「不列」「豁免」「免列」「不另列」之一，視為涵蓋（理由寫在同一句）。strict 群組（IF_HoldEcon）不得豁免：缺案例一律失敗。
# 適用：In 頁 E3＝「基準」、D3 含「低」、F3 含「高」者（舊模板如 Nebius、CoreWeave、WhiteFiber、OpenAI 無 In 頁三欄＝不適用）。
#       適用但無 Sens 頁＝失敗。
# 用法：
#   python3 tools/sens_coverage/check_sens_coverage.py <成品 xlsx> [--groups 規則.json]   # 單檔；有未涵蓋即失敗（結束碼 1）
#   python3 tools/sens_coverage/check_sens_coverage.py --all [--md 報告.md] [--groups 規則.json]  # 掃描全部 */dist/*.xlsx（每資料夾取檔名排序最後一檔），只報告（結束碼 0）
import glob, json, os, re, sys
import openpyxl

GROUPS_DEFAULT = [
    {"name": "IF_HoldEcon", "key_re": r"^he_(h|b2|b3|vr|h100|gb200|gb300|vr200)$", "label_re": r"IF_HoldEcon｜",
     "aliases": ["he_all"], "text_all": [["IF_HoldEcon", "四世代"], ["IF_HoldEcon", "同時"]], "strict": True},
]
EXEMPT_WORDS = ("不列", "豁免", "免列", "不另列")
NOT_MODEL_WORDS = ("不進模型",)
REF_HINT = re.compile(r"^(cons_|c_|g_q\d|old_tp|tp_v\d|p_spot|p_long)")
IDENT = r"(?<![A-Za-z0-9_]){}(?![A-Za-z0-9_])"


def num(x):
    return isinstance(x, (int, float)) and not isinstance(x, bool)


def has_tok(text, key):
    return re.search(IDENT.format(re.escape(key)), text) is not None


def in_keys(wb):
    """回傳 None（不適用）或 [(鍵, 項目, 低, 基準, 高, 不進模型)]"""
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
        d, e, f = (ws.cell(r, c).value for c in (4, 5, 6))
        lo = num(d) and not (num(e) and abs(d - e) <= 1e-12)
        hi = num(f) and not (num(e) and abs(f - e) <= 1e-12)
        if not (lo or hi):
            continue
        texts = " ".join(str(ws.cell(r, c).value) for c in range(2, ws.max_column + 1) if isinstance(ws.cell(r, c).value, str))
        out.append((str(a), str(ws.cell(r, 2).value or ""), d, e, f, any(w in texts for w in NOT_MODEL_WORDS)))
    return out


def sens_region(ws):
    """回傳 (表頭以上註記文字列, 案例列清單[(列號, 鍵欄文字, 全列文字)])"""
    maxc = min(ws.max_column, 30)
    hdr, keycol = None, None
    for r in range(1, min(ws.max_row, 15) + 1):
        vals = [str(ws.cell(r, c).value or "").strip() for c in range(1, maxc + 1)]
        if "鍵" in vals:
            hdr, keycol = r, vals.index("鍵") + 1
            break
    if hdr is None:  # 無「鍵」欄：第一個「至少三格非空、A 欄為短文字」的列當表頭
        for r in range(1, min(ws.max_row, 15) + 1):
            vals = [ws.cell(r, c).value for c in range(1, maxc + 1)]
            if sum(v is not None for v in vals) >= 3 and isinstance(vals[0], str) and len(vals[0]) <= 12:
                hdr = r
                break
    if hdr is None:
        hdr = 1
    notes = [str(ws.cell(r, c).value) for r in range(1, hdr) for c in range(1, maxc + 1) if isinstance(ws.cell(r, c).value, str)]
    rows, blank = [], 0
    for r in range(hdr + 1, min(ws.max_row, 2000) + 1):
        vals = [ws.cell(r, c).value for c in range(1, maxc + 1)]
        if all(v is None for v in vals):
            blank += 1
            if blank >= 10:
                break
            continue
        blank = 0
        if str(vals[0] or "").startswith("敏感度「"):  # Microsoft 式逐案明細區
            break
        kt = str(ws.cell(r, keycol).value or "") if keycol else ""
        rows.append((r, kt, " ".join(str(v) for v in vals if isinstance(v, str))))
    return notes, rows, keycol


def load_groups(path):
    g = list(GROUPS_DEFAULT)
    if path:
        g += json.load(open(path, encoding="utf-8"))["groups"]
    return g


def check(path, groups=None):
    """回傳 dict：applicable, keys（待涵蓋數）, missing [(鍵, 項目, 低, 基準, 高, 原因, 提示)], exempt [(鍵, 理由句)],
    group_cov {群組: 案例列}, not_model [鍵], keycol（Sens 頁有無鍵欄）, error"""
    groups = groups if groups is not None else load_groups(None)
    wb = openpyxl.load_workbook(path, data_only=True)
    ks = in_keys(wb)
    res = {"applicable": ks is not None, "keys": 0, "missing": [], "exempt": [], "group_cov": {}, "not_model": [], "keycol": None, "error": None}
    if ks is None:
        return res
    res["not_model"] = [k[0] for k in ks if k[5]]
    ks = [k for k in ks if not k[5]]
    res["keys"] = len(ks)
    if "Sens" not in wb.sheetnames:
        res["error"] = "適用（In 頁有低／高欄）但無 Sens 頁"
        res["missing"] = [(k, lab, d, e, f, "無 Sens 頁", "") for k, lab, d, e, f, _ in ks]
        return res
    notes, rows, keycol = sens_region(wb["Sens"])
    res["keycol"] = bool(keycol)
    sents = []
    for t in notes + [x[2] for x in rows]:
        sents += [s for s in re.split(r"[。；;\n]", t) if s.strip()]
    keyset = {r[1] for r in rows}

    def member(g, k, lab):
        return bool(re.search(g["key_re"], k)) or bool(g.get("label_re") and re.search(g["label_re"], lab))

    gmembers = {g["name"]: [k for k, lab, *_ in ks if member(g, k, lab)] for g in groups}
    for g in groups:
        mem = gmembers[g["name"]]
        if not mem:
            continue
        for r, kt, txt in rows:
            t = kt + " " + txt
            if any(has_tok(t, a) for a in g.get("aliases", [])) or any(all(w in t for w in ws) for ws in g.get("text_all", [])) \
                    or all(has_tok(t, m) for m in mem):
                res["group_cov"][g["name"]] = r
                break
    for k, lab, d, e, f, _ in ks:
        g = next((g for g in groups if member(g, k, lab)), None)
        if g is not None:
            if g["name"] in res["group_cov"]:
                continue
            if not g.get("strict"):
                ex = [s for s in sents if has_tok(s, k) and any(w in s for w in EXEMPT_WORDS)]
                if ex:
                    res["exempt"].append((k, ex[0].strip()[:160]))
                    continue
            single = [r for r, kt, txt in rows if has_tok(kt + " " + txt, k)]
            why = f"{g['name']} 群組缺「整組同時」案例" + ("（只有單一成員案例，第 " + "、".join(map(str, single)) + " 列）" if single else "") + \
                  ("；群組不得豁免" if g.get("strict") else "")
            res["missing"].append((k, lab, d, e, f, why, ""))
            continue
        if any(has_tok(kt, k) or has_tok(txt, k) for r, kt, txt in rows):
            continue
        ex = [s for s in sents if has_tok(s, k) and any(w in s for w in EXEMPT_WORDS)]
        if ex:
            res["exempt"].append((k, ex[0].strip()[:160]))
            continue
        stem = re.split(r"[（(｜]", lab)[0].strip()
        hint = [r for r, kt, txt in rows if len(stem) >= 3 and stem in txt]
        hs = ("疑似案例列（項目文字相符，但未寫鍵）：第 " + "、".join(map(str, hint[:4])) + " 列") if hint else ""
        if REF_HINT.search(k) or re.search(r"共識|證據|v\d+\.\d[\d.]*\s*基準目標價", lab):
            hs = (hs + "；" if hs else "") + "疑似對照用（共識／指引／舊版目標價／證據）：可在 In 頁註明「不進模型」或 Sens 頁寫豁免"
        res["missing"].append((k, lab, d, e, f, "無案例、無豁免", hs))
    return res


def fmt(v):
    return "—" if v is None else (f"{v:.4g}" if num(v) else str(v))


def main(argv):
    gpath = argv[argv.index("--groups") + 1] if "--groups" in argv else None
    groups = load_groups(gpath)
    if argv and argv[0] == "--all":
        md = argv[argv.index("--md") + 1] if "--md" in argv else None
        root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
        lines = ["| 資料夾 | 成品 | 待涵蓋鍵 | 未涵蓋 | 其中疑似有案例（文字相符、未寫鍵） | 其中疑似對照用 | IF_HoldEcon | 豁免 | Sens 鍵欄 | 未涵蓋清單 |",
                 "|---|---|---|---|---|---|---|---|---|---|"]
        na, n_ok, n_bad = [], 0, 0
        for d in sorted(glob.glob(os.path.join(root, "*", "dist"))):
            xs = sorted(glob.glob(os.path.join(d, "*.xlsx")))
            if not xs:
                continue
            x, co = xs[-1], os.path.basename(os.path.dirname(d))
            try:
                r = check(x, groups)
            except Exception as e:
                lines.append(f"| `{co}` | {os.path.basename(x)} | — | 讀檔失敗 | — | — | — | — | — | {e} |")
                n_bad += 1
                continue
            if not r["applicable"]:
                na.append(co)
                continue
            he = [m for m in r["missing"] if m[5].startswith("IF_HoldEcon")]
            hes = "缺（失敗）" if he else ("有（第 %d 列）" % r["group_cov"]["IF_HoldEcon"] if "IF_HoldEcon" in r["group_cov"] else "無此群組")
            miss = "、".join(m[0] for m in r["missing"])
            n_like = sum("疑似案例列" in m[6] for m in r["missing"])
            n_ref = sum("疑似對照用" in m[6] for m in r["missing"])
            lines.append(f"| `{co}` | {os.path.basename(x)} | {r['keys']} | {len(r['missing'])} | {n_like} | {n_ref} | {hes} | {len(r['exempt'])} | "
                         f"{'有' if r['keycol'] else ('無' if r['keycol'] is False else '—')} | {miss or '—'} |")
            n_ok += not r["missing"]
            n_bad += bool(r["missing"])
        print("\n".join(lines))
        tail = f"\n適用 {n_ok + n_bad} 檔：全涵蓋 {n_ok}、有缺口 {n_bad}；不適用 {len(na)} 檔：" + "、".join(na)
        print(tail)
        if md:
            open(md, "w", encoding="utf-8").write("\n".join(lines) + "\n" + tail + "\n")
        return 0
    r = check(argv[0], groups)
    if not r["applicable"]:
        print("敏感度覆蓋核對：不適用（無 In 頁低／基準／高三欄）")
        return 0
    for k, s in r["exempt"]:
        print(f"豁免 {k}：{s}")
    if r["not_model"]:
        print("In 頁註明不進模型（不列入）：" + "、".join(r["not_model"]))
    if r["missing"]:
        print(f"敏感度覆蓋核對失敗：{len(r['missing'])}／{r['keys']} 個有低／高區間的鍵沒有 Sens 案例或豁免理由")
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
