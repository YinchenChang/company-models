"""三層基準協調檢查（規則見 docs/plan/三層基準協調規則.md 第 10 節；第 13 項來源雜湊（A8）、第 14 項欄位語義（A7）、第 15 項只顯示數列（W2-05B）為三層對帳工作單新增；v1.4.3 起接 CI：.github/workflows/coord-check.yml）。

狀態檔 siliconomics/data/baseline_status.json 不存在時只印「未啟用」並以 0 結束。
輸出兩類：「不符」（以 1 結束）與「提醒 [n]」（不影響結束碼）。

用法：python3 tools/baseline_coordination/check.py [--status 路徑] [--today YYYY-MM-DD]
"""
import argparse, ast, datetime, hashlib, json, pathlib, re, sys

ROOT = pathlib.Path(__file__).resolve().parents[2]
DEFAULT_STATUS = ROOT / "siliconomics/data/baseline_status.json"
VALID = {"甲", "乙", "丙", "待定"}  # 待定：規則 v1.3 區間跨線
TOL = 0.005  # 第 2 項：換算值比對容許的相對誤差（四捨五入）
NOTES = []   # 提醒（不算不符）


def note(msg):
    NOTES.append(msg)


ENTER, EXIT = 0.55, 0.45  # v1.1 緩衝區


def has_evidence(share, thr, prev):
    """v1.1：首次判定（prev 為 None）以門檻為界；已有證據者 < EXIT 才退出；原無證據者 ≥ ENTER 才進入。"""
    if prev is None:
        return share >= thr
    return share >= EXIT if prev else share >= ENTER


def _st(w, s):
    return "甲" if (w and s) else ("乙" if (w or s) else "丙")


def _three(x):
    """證據比例可為點值或 {lo, base, hi}；回傳 (lo, base, hi)。"""
    if isinstance(x, dict):
        return x.get("lo", x["base"]), x["base"], x.get("hi", x["base"])
    return x, x, x


def expected_status(w_share, s_share, thr, prev_w=None, prev_s=None):
    """C3：兩層證據比例 → 甲／乙／丙（含 v1.1 緩衝區）；v1.3：區間端點與基準值判定不同 → 待定。"""
    wl, wb, wh = _three(w_share); sl, sb, sh = _three(s_share)
    ev = lambda v, p: has_evidence(v, thr, p)
    base = _st(ev(wb, prev_w), ev(sb, prev_s))
    ends = {_st(ev(a, prev_w), ev(b, prev_s)) for a in (wl, wb, wh) for b in (sl, sb, sh)}
    return "待定" if len(ends) > 1 else base


def _check_c0(yr, y, today, require_label):
    """12. C0 實績對帳不過（c0_fail is True）的登錄檢查；c0_fail 只認布林 True。"""
    errs = []
    flag = y.get("c0_fail")
    if flag is not None and not isinstance(flag, bool):
        return [f"[12] {yr} c0_fail 須為布林值，現為 {flag!r}"]
    if flag is not True:
        return errs
    if require_label and "有條件" not in str(y.get("label") or ""):
        errs.append(f"[12] {yr} C0 實績對帳不過（c0_fail），甲年標示「{y.get('label')}」未含「有條件」")
    if not y.get("c0_recon_doc"):
        errs.append(f"[12] {yr} C0 實績對帳不過但未登錄對帳單（c0_recon_doc）")
    dl = y.get("c0_deadline")
    if not dl:
        errs.append(f"[12] {yr} C0 實績對帳不過但未寫期限（c0_deadline）")
    else:
        try:
            d = datetime.date.fromisoformat(dl)
        except (TypeError, ValueError):
            errs.append(f"[12] {yr} c0_deadline 格式錯誤：{dl!r}（應為 YYYY-MM-DD）")
        else:
            if d < today:
                errs.append(f"[12] {yr} C0 對帳單已逾期（{dl}）：依 C6 第 3 點強制裁定")
    return errs


def check_status(st, today):
    errs = []
    thr = st["threshold"]
    for yr, y in sorted(st["years"].items()):
        tag = y.get("status")
        # 1. 有標籤且與證據比例一致
        if tag not in VALID:
            errs.append(f"[1] {yr} 缺狀態標籤或標籤無效：{tag!r}")
        else:
            exp = expected_status(y["wattonomics"]["evidence_share"], y["siliconomics"]["evidence_share"], thr,
                                  y["wattonomics"].get("prev_has_evidence"), y["siliconomics"].get("prev_has_evidence"))
            if tag != exp:
                errs.append(f"[1] {yr} 標「{tag}」，依證據比例與門檻 {thr:.0%} 應為「{exp}」")
        # 8. 丙、或上游有證據的乙：須 Wattonomics 覆蓋已補齊，否則須標對帳中（v1.1）
        upstream_led = tag == "丙" or (tag == "乙" and y.get("lead") == "wattonomics")
        if upstream_led and not st.get("wattonomics_coverage_complete") and not y.get("reconciling"):
            errs.append(f"[8] {yr} 由上游定基準但 Wattonomics 覆蓋未補齊，應標對帳中")
        # 7. ② 累計 ÷ 當年基準 ≤ 15%（v1.1）；缺欄位＝未量測，列提醒（v1.4.3）
        if "timing_cum" not in y:
            note(f"提醒 [7] {yr} 未量測（狀態檔缺 timing_cum，C2 ② 累計上限未檢查）")
        elif abs(y["timing_cum"]) > 0.15 * abs(y["sil_base"]):
            errs.append(f"[7] {yr} ② 時差與庫存累計 {y['timing_cum']} 超過當年基準 15%")
        # 12. C0 實績對帳不過時（c0_fail）：甲年標示須含「有條件」；所有年份須登錄對帳單與期限（v1.4.3）
        errs += _check_c0(yr, y, today, require_label=(tag == "甲"))
        # 2. 丙狀態：Siliconomics 基準＝換算值（以供給上限為頂）
        if tag == "丙" and st.get("wattonomics_coverage_complete"):
            conv = min(y["converted"], y.get("supply_cap", float("inf")))
            if y.get("continuity"):  # v1.4 連續法：前一年基準 × 換算值成長率
                py = st["years"].get(str(int(yr) - 1), {})
                if py.get("sil_base") and py.get("converted"):
                    conv = py["sil_base"] * y["converted"] / py["converted"]
            if abs(y["sil_base"] - conv) > TOL * abs(conv):
                errs.append(f"[2] {yr} 丙狀態基準 {y['sil_base']} ≠ 換算值 {conv}")
        # 5. 對帳中：須有對帳單且未逾期
        if y.get("reconciling"):
            if not y.get("recon_doc"):
                errs.append(f"[5] {yr} 標對帳中但未登錄對帳單")
            dl = y.get("deadline")
            if not dl:
                errs.append(f"[5] {yr} 對帳單未寫診斷期限")
            elif datetime.date.fromisoformat(dl) < today:
                errs.append(f"[5] {yr} 對帳單已逾期（{dl}）：依 C6 第 3 點強制裁定")
    # 3. 基準鏈上不得有 [Assumed] 水準係數
    for c in st.get("level_coefficients", []):
        if c.get("tag") == "Assumed":
            y = st.get("years", {}).get(str(c.get("year")), {})
            dl = y.get("deadline")
            exempt = (y.get("status") == "待定" and y.get("reconciling") and y.get("recon_doc")
                      and dl and datetime.date.fromisoformat(dl) >= today)   # v1.4.1：待定維持現行的年份豁免
            if exempt:
                note(f"提醒 [3] {c.get('year')} 待定維持現行，基準鏈上暫有 [Assumed] 係數：{c.get('name')}（對帳單期限 {dl}）")
            else:
                errs.append(f"[3] 基準鏈上有 [Assumed] 水準係數：{c.get('name')}")
    return errs


def check_company_reads(st):
    """4. 硬體個股只讀 *_Base；6. 依賴對帳中年份者，成品有標示與兩讀數目標價。"""
    errs = []
    pat = re.compile(r"\bSIF_\w+?_(A|C)\b")  # 已停用的情境名稱
    for co in st.get("hardware_companies", []):
        for f in (ROOT / co).glob("*.py"):
            for m in pat.finditer(f.read_text(encoding="utf-8", errors="ignore")):
                errs.append(f"[4] {co}/{f.name} 讀非基準名稱 {m.group(0)}")
    recon_years = {str(yr) for yr, y in st["years"].items() if y.get("reconciling")}
    for co, years in st.get("depends_on_years", {}).items():
        if not recon_years & {str(x) for x in years}:  # 年份鍵一律轉字串比對（整數、字串皆可）
            continue
        htmls = sorted((ROOT / co / "dist").glob("*.html"))
        if not htmls:
            errs.append(f"[6] {co} 依賴對帳中年份但無 HTML 成品")
            continue
        txt = htmls[-1].read_text(encoding="utf-8", errors="ignore")
        for need in ("對帳中", "兩讀數目標價"):
            if need not in txt:
                errs.append(f"[6] {co} {htmls[-1].name} 缺「{need}」")
    return errs


def _ver(s):
    m = re.search(r"v\d+(?:\.\d+)+", s or "")
    return m.group(0) if m else None


def check_version_drift(st):
    """11. 狀態檔記錄的 Wattonomics 版本＝wattonomics/model/CURRENT 的版本（v1.4.3）。"""
    errs = []
    cur_file = ROOT / "wattonomics/model/CURRENT"
    cur = _ver(cur_file.read_text(encoding="utf-8").strip()) if cur_file.exists() else None
    if cur is None:
        note("提醒 [11] 找不到 wattonomics/model/CURRENT 或其中無版本字串，版本漂移未檢查")
        return errs
    rec = st.get("wattonomics_version")
    if rec is None:
        m = re.search(r"Wattonomics\s*(v\d+(?:\.\d+)+)", st.get("_note", ""))
        nv = m.group(1) if m else "（_note 亦無版本字串）"
        note(f"提醒 [11] 狀態檔未記錄 Wattonomics 版本（_note 寫 {nv}；wattonomics/model/CURRENT＝{cur}）"
             + ("：兩者不同，下次匯出狀態檔時補 wattonomics_version 並重算" if m and nv != cur else ""))
    elif _ver(str(rec)) != cur:
        errs.append(f"[11] 狀態檔 wattonomics_version＝{rec}，wattonomics/model/CURRENT＝{cur}：版本漂移")
    return errs


def check_recon_residual():
    """9. C0 恆等式：讀 CURRENT_RECON 指向的 recon 快照（v1.4.3；Siliconomics v0.20 起快照有 kind／status／market／殘差列）。
    kind＝market：status「未建」列提醒；其他狀態（已建）須有殘差列（who 含「殘差」「其他」「未建模」或 residual: true），
    且有 market 值時逐年檢查 Σ rows（含殘差）＝market（相對誤差 1e-6），否則不符。
    kind＝structural（結構一致，由 Siliconomics Checks K6 約束）、layer（層間對帳點，R4）、actual（公司實績對帳點）不是個股加總，不檢查恆等式，只計數提醒。
    無 kind 欄的舊快照：市場有 status 欄時同上判定；無 status 而有 rows＝結構無法判定，只列提醒。"""
    errs = []
    ptr = ROOT / "siliconomics/data/CURRENT_RECON"
    if not ptr.exists():
        note("提醒 [9] 找不到 siliconomics/data/CURRENT_RECON，C0 恆等式未檢查")
        return errs
    fn = ptr.read_text(encoding="utf-8").strip()
    p = ROOT / "siliconomics/data" / fn
    if not p.exists():
        errs.append(f"[9] CURRENT_RECON 指向不存在的檔案 {fn}")
        return errs
    mk = json.loads(p.read_text(encoding="utf-8")).get("markets", {})
    unbuilt = [k for k, m in mk.items() if isinstance(m, dict) and m.get("status") == "未建"]
    if unbuilt:
        note(f"提醒 [9] {fn}：狀態「未建」的市場 {len(unbuilt)} 個：{'、'.join(unbuilt)}")
    other = [k for k, m in mk.items() if isinstance(m, dict) and m.get("kind") in ("structural", "layer", "actual")]
    if other:
        note(f"提醒 [9] {fn}：非個股加總的對帳點 {len(other)} 個不檢查恆等式（結構一致／層間／公司實績）：{'、'.join(other)}")
    undecidable, built = [], []
    for k, m in mk.items():
        if not isinstance(m, dict) or m.get("status") == "未建" or k in other:
            continue
        rows = m.get("rows")
        if "status" not in m:
            if rows is not None:
                undecidable.append(k)
            continue
        built.append(k)
        rws = rows or []
        res = [r for r in rws if r.get("residual") or any(t in str(r.get("who", "")) for t in ("殘差", "其他", "未建模"))]
        if not res:
            errs.append(f"[9] {fn} 市場 {k} 已建但缺殘差／其他列")
            continue
        if any(v is None for r in res for v in (r.get("vals") or [None])):
            errs.append(f"[9] {fn} 市場 {k} 殘差列有空值（殘差不得為空）")
        mv = m.get("market")
        if mv is None:
            note(f"提醒 [9] {fn} 市場 {k} 無 market 值，恆等式只檢查殘差列存在")
            continue
        for i, yv in enumerate(mv):
            if not isinstance(yv, (int, float)):
                continue
            tot = sum(r["vals"][i] for r in rws if isinstance(r.get("vals", [None] * (i + 1))[i], (int, float)))
            if abs(tot - yv) > 1e-6 * max(1.0, abs(yv)):
                yr = (m.get("years") or list(range(len(mv))))[i]
                errs.append(f"[9] {fn} 市場 {k} {yr}：已建模加總＋殘差 {tot:.6g} ≠ 市場值 {yv:.6g}")
    if built:
        note(f"提醒 [9] {fn}：已判定恆等式的市場 {len(built)} 個：{'、'.join(built)}（狀態：{'、'.join(str(mk[k].get('status')) for k in built)}）")
    if undecidable:
        note(f"提醒 [9] {fn}：市場 {'、'.join(undecidable)} 只有個股列、無 status 與殘差欄位，結構無法判定 C0 恆等式（未建市場清單亦未記錄於快照）")
    return errs


def check_c0_actuals(st, today):
    """12（續）. 協調年份以外的已公布年份：頂層 c0_actuals.<年>（C0 第 9 點；v1.4.3）。"""
    errs = []
    for yr, y in sorted((st.get("c0_actuals") or {}).items()):
        errs += _check_c0(yr, y, today, require_label=False)
    return errs


def check_evidence_range():
    """10. Wattonomics 證據比例區間：WIF_Evidence_Lo_<yyyy> ≤ WIF_Evidence_<yyyy> ≤ WIF_Evidence_Hi_<yyyy>（v1.4.3）。"""
    errs = []
    ptr = ROOT / "wattonomics/data/CURRENT_SNAPSHOT"
    if not ptr.exists():
        note("提醒 [10] 找不到 wattonomics/data/CURRENT_SNAPSHOT，證據比例區間未檢查")
        return errs
    fn = ptr.read_text(encoding="utf-8").strip()
    p = ROOT / "wattonomics/data" / fn
    if not p.exists():
        errs.append(f"[10] CURRENT_SNAPSHOT 指向不存在的檔案 {fn}")
        return errs
    names = json.loads(p.read_text(encoding="utf-8")).get("names", {})
    years = sorted(k.rsplit("_", 1)[1] for k in names if re.fullmatch(r"WIF_Evidence_\d{4}", k))
    missing = []
    for yr in years:
        b, lo, hi = names[f"WIF_Evidence_{yr}"], names.get(f"WIF_Evidence_Lo_{yr}"), names.get(f"WIF_Evidence_Hi_{yr}")
        if lo is None and hi is None:
            missing.append(yr)
            continue
        if lo is None or hi is None:
            errs.append(f"[10] {fn} {yr} 只有 {'Hi' if lo is None else 'Lo'}，證據比例區間不完整")
            continue
        if not (lo <= b <= hi):
            errs.append(f"[10] {fn} {yr} 證據比例 低 {lo} ≤ 基準 {b} ≤ 高 {hi} 不成立")
    if missing:
        note(f"提醒 [10] {fn} 無 WIF_Evidence_Lo_／Hi_ 名稱（{'、'.join(missing)}），證據比例區間未檢查")
    return errs


def _sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()


def check_source_hashes(st):
    """13. 來源雜湊一致（三層對帳工作單 A8）：快照記錄的 source.sha256＝repo 內該檔實際 SHA-256。
    檢查對象：Siliconomics 與 Wattonomics 的 CURRENT 快照（對應 model/<source.file>）、狀態檔的 source_sha256（對 Siliconomics 活頁簿）與
    wattonomics_source_sha256、bridge.sha256（對協調橋接檔）。活頁簿重存（LibreOffice 重算）會改變位元組，故匯出須在定稿後的最後一步。"""
    errs = []
    for layer in ("siliconomics", "wattonomics"):
        ptr = ROOT / layer / "data" / "CURRENT_SNAPSHOT"
        if not ptr.exists():
            note(f"提醒 [13] 找不到 {layer}/data/CURRENT_SNAPSHOT，來源雜湊未檢查")
            continue
        fn = ptr.read_text(encoding="utf-8").strip()
        sp = ROOT / layer / "data" / fn
        if not sp.exists():
            errs.append(f"[13] {layer} CURRENT_SNAPSHOT 指向不存在的檔案 {fn}")
            continue
        src = json.loads(sp.read_text(encoding="utf-8")).get("source", {})
        mp = ROOT / layer / "model" / str(src.get("file", ""))
        if not src.get("sha256"):
            errs.append(f"[13] {fn} 未記錄 source.sha256")
        elif not mp.is_file():
            errs.append(f"[13] {fn} 記錄的來源檔不存在：{src.get('file')}")
        elif src["sha256"] != _sha(mp):
            errs.append(f"[13] {fn} 記錄 source.sha256＝{src['sha256'][:8]}…，{layer}/model/{mp.name} 實際為 {_sha(mp)[:8]}…（匯出後活頁簿被重存？重跑匯出）")
    # 狀態檔自身
    sil_ptr = ROOT / "siliconomics/model/CURRENT"
    wat_ptr = ROOT / "wattonomics/model/CURRENT"
    pairs = []
    if "source_sha256" in st and sil_ptr.exists():
        pairs.append(("source_sha256", ROOT / "siliconomics/model" / st.get("source", ""), st["source_sha256"]))
    elif "source_sha256" not in st:
        note("提醒 [13] 狀態檔未記錄 source_sha256（下次匯出狀態檔時補上）")
    if "wattonomics_source_sha256" in st and wat_ptr.exists():
        pairs.append(("wattonomics_source_sha256", ROOT / "wattonomics/model" / wat_ptr.read_text(encoding="utf-8").strip(), st["wattonomics_source_sha256"]))
    br = st.get("bridge")
    if br:
        pairs.append(("bridge.sha256", ROOT / "siliconomics/docs/reports" / br.get("file", ""), br.get("sha256")))
    for label, fp, rec in pairs:
        if not fp.is_file():
            errs.append(f"[13] 狀態檔 {label} 對應的檔案不存在：{fp.relative_to(ROOT)}")
        elif rec != _sha(fp):
            errs.append(f"[13] 狀態檔 {label}＝{str(rec)[:8]}…，{fp.relative_to(ROOT)} 實際為 {_sha(fp)[:8]}…")
    return errs


# 14. 欄位語義（三層對帳工作單 A7）：Siliconomics Interface E 節 SIF_DeployGW_* 五組欄位各只有一種語義（見 siliconomics/README.md「欄位語義」）：
#     Base＝依規則的基準（硬體個股計算基準的唯一來源）；Supply＝供給側讀數（原始顆數，情境 A，只並列顯示）；Demand＝需求側讀數（Wattonomics 換算值，只並列顯示）；
#     Alt／AltRatio＝另一側讀數與比值（只用於並列顯示與兩讀數目標價〔情境〕）；Current＝舊 S33 係數取法（v0.18 基準，僅供對照，不得進任何計算）。
# 這是靜態檢查（字串與 AST），只能擋「名稱出現在基準指派」與「未登錄用途的讀取」，不能證明個股沒有在別處間接使用；新增讀取者須先登錄於下表。
SEM_STEMS = ("Current", "Supply", "Demand", "AltRatio", "Alt")
SEM_PAT = re.compile(r"SIF_DeployGW_(Current|Supply|Demand|AltRatio|Alt)(?![A-Za-z])")
# 非 Base 欄位允許出現的用途（檔案 → {欄位: 用途}）；其餘硬體個股腳本只可讀 Alt／AltRatio／Supply（並列顯示或情境），讀 Current、Demand 一律不符。
SEM_ALLOW = {
    "nvidia/coord_post.py": {
        "Current": "並列顯示：Coord 頁「改前現行值」對照欄（C['prev']），不進任何計算",
        "Demand": "註解與 extract() 摘錄清單，不計算",
    },
}
# 例外：目前已知違規、待修（修完後這裡可刪；例外已不再需要時只列提醒，不算不符）
SEM_BASE_EXC = {
}


def _sem_base_assign_hits(tree):
    """AST：目標名稱含 base 的指派，其右側若含非 Base 欄位的字串（含 f-string 片段）→ 回傳 (行號, 欄位)。"""
    hits = []
    for node in ast.walk(tree):
        tg = []
        if isinstance(node, ast.Assign):
            tg, val = node.targets, node.value
        elif isinstance(node, (ast.AnnAssign, ast.NamedExpr)) and node.value is not None:
            tg, val = [node.target], node.value
        else:
            continue
        if not any("base" in (n.id if isinstance(n, ast.Name) else getattr(n, "attr", "")).lower()
                   for t in tg for n in ast.walk(t) if isinstance(n, (ast.Name, ast.Attribute))):
            continue
        for c in ast.walk(val):
            if isinstance(c, ast.Constant) and isinstance(c.value, str):
                for m in SEM_PAT.finditer(c.value):
                    hits.append((node.lineno, m.group(1)))
    return hits


def check_field_semantics(st):
    errs = []
    used_exc = set()
    for co in st.get("hardware_companies", []):
        for f in sorted((ROOT / co).glob("*.py")):
            rel = f"{co}/{f.name}"
            txt = f.read_text(encoding="utf-8", errors="ignore")
            if "SIF_DeployGW_" not in txt:
                continue
            allow = SEM_ALLOW.get(rel, {})
            for ln, line in enumerate(txt.splitlines(), 1):
                for m in SEM_PAT.finditer(line):
                    stem = m.group(1)
                    if stem in ("Current", "Demand") and stem not in allow:
                        errs.append(f"[14] {rel}:{ln} 讀 SIF_DeployGW_{stem}_（{'舊 S33 取法，僅供對照' if stem == 'Current' else '需求側讀數，只供並列顯示'}），硬體個股不得讀；未登錄於 check.py SEM_ALLOW")
            try:
                tree = ast.parse(txt)
            except SyntaxError:
                note(f"提醒 [14] {rel} 無法解析（語法錯誤），基準指派未檢查")
                continue
            for ln, stem in _sem_base_assign_hits(tree):
                if rel in SEM_BASE_EXC:
                    used_exc.add(rel)
                    note(f"提醒 [14] 例外：{rel}:{ln} 基準指派含 SIF_DeployGW_{stem}_（{SEM_BASE_EXC[rel]}）")
                else:
                    errs.append(f"[14] {rel}:{ln} 計算基準的指派含 SIF_DeployGW_{stem}_：硬體個股計算基準只能讀 SIF_DeployGW_Base（{stem} 只可並列顯示或作情境）")
    for rel in SEM_BASE_EXC:
        if rel not in used_exc:
            note(f"提醒 [14] 例外已不需要，可從 check.py SEM_BASE_EXC 刪除：{rel}")
    return errs


# 15. 只顯示的觀測數列（三層對帳 W2-05B）。依據：規則 v1.5 草案 C3（待 Andy 核可）；本項只防止只顯示數列被讀取，不改任何判定。
#     (a) 狀態檔每年若有 observations，display_only 必須為 true；沒有 observations（舊狀態檔或不適用年份）只列提醒。
#     (b) 讀取 Siliconomics 名稱或狀態檔的腳本（全 repo *.py；排除 siliconomics/、docs/、tools/baseline_coordination/、.git 與隱藏資料夾）不得出現讀取
#         observations 或其欄位的程式碼。靜態檢查（AST）：只看字串常數、屬性名稱，不看註解與說明文字（註解不會進 AST）。
OBS_SKIP = ("siliconomics/", "docs/", "tools/baseline_coordination/")
OBS_FILE_PAT = re.compile(r"SIF_|baseline_status|siliconomics", re.I)  # 先挑出「讀 Siliconomics 名稱或狀態檔」的腳本
OBS_KEY_PAT = re.compile(r"^observations$")  # 整個字串＝鍵名
# 欄位名稱：整詞比對（前後不得緊接英數或底線，避免 w_mid_lo、dec10 之類誤判）；crd_lam_ 為前綴（crd_lam_0／_05／_1）
OBS_FIELD_PAT = re.compile(r"(?<![A-Za-z0-9_])(?:mid_(?:iii|ii|i|lo|hi)|dec(?:05|0|1)|crd_lam_[A-Za-z0-9_]*)(?![A-Za-z0-9_])")


def _obs_hits(tree):
    """AST：回傳 (行號, 說明)。字串常數（含 f-string 片段、下標鍵）、屬性名稱命中 observations 或其欄位名；變數名稱、註解不算。"""
    hits = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Constant) and isinstance(node.value, str):
            v = node.value
            if OBS_KEY_PAT.match(v.strip()):
                hits.append((node.lineno, "字串 'observations'"))
            for m in OBS_FIELD_PAT.finditer(v):
                hits.append((node.lineno, f"字串含 {m.group(0)}"))
        elif isinstance(node, ast.Attribute) and node.attr == "observations":
            hits.append((node.lineno, "屬性 .observations"))
    return hits


def check_observations(st):
    errs = []
    missing = []
    for y, v in sorted((st.get("years") or {}).items()):
        if "observations" not in v:
            missing.append(y)
            continue
        ob = v["observations"]
        if not isinstance(ob, dict) or ob.get("display_only") is not True:
            got = ob.get("display_only") if isinstance(ob, dict) else type(ob).__name__
            errs.append(f"[15] 狀態檔 years.{y}.observations.display_only 必須為 true（只顯示數列；規則 v1.5 草案 C3，待 Andy 核可），實際為 {got}")
    if missing:
        note(f"提醒 [15] 狀態檔 {'、'.join(missing)} 年沒有 observations 欄位（舊狀態檔或不適用），display_only 未檢查")
    n, bad = 0, 0
    for f in sorted(ROOT.rglob("*.py")):
        rel = f.relative_to(ROOT).as_posix()
        if rel.startswith(OBS_SKIP) or any(part.startswith(".") for part in rel.split("/")):
            continue
        txt = f.read_text(encoding="utf-8", errors="ignore")
        if not OBS_FILE_PAT.search(txt):
            continue
        n += 1
        try:
            tree = ast.parse(txt)
        except SyntaxError:
            note(f"提醒 [15] {rel} 無法解析（語法錯誤），observations 讀取未檢查")
            continue
        for ln, what in _obs_hits(tree):
            bad += 1
            errs.append(f"[15] {rel}:{ln} 讀取只顯示的觀測數列（{what}）：observations 與其欄位（mid_*／dec*／crd_lam_*）只供 Siliconomics Coord 頁顯示，個股與下游腳本不得讀取（規則 v1.5 草案 C3，待 Andy 核可）")
    if not bad:
        note(f"提醒 [15] 已掃描 {n} 支讀取 Siliconomics 名稱或狀態檔的腳本，無一讀取 observations（規則 v1.5 草案 C3，待 Andy 核可）")
    return errs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--status", default=str(DEFAULT_STATUS))
    ap.add_argument("--today", default=datetime.date.today().isoformat())
    a = ap.parse_args()
    p = pathlib.Path(a.status)
    if not p.exists():
        print(f"未啟用：找不到狀態檔 {p}（步驟 3 量測後建立）")
        return 0
    st = json.loads(p.read_text(encoding="utf-8"))
    today = datetime.date.fromisoformat(a.today)
    errs = (check_status(st, today) + check_c0_actuals(st, today) + check_company_reads(st)
            + check_version_drift(st) + check_recon_residual() + check_evidence_range() + check_source_hashes(st) + check_field_semantics(st) + check_observations(st))
    for n in NOTES:
        print(n)
    for e in errs:
        print(e)
    print(f"三層基準協調檢查：{len(errs)} 項不符、{len(NOTES)} 項提醒")
    return 1 if errs else 0


if __name__ == "__main__":
    sys.exit(main())
