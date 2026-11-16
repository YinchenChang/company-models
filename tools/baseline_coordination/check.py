"""三層基準協調檢查（規則見 docs/plan/三層基準協調規則.md 第 10 節；v1.4.3 起接 CI：.github/workflows/coord-check.yml）。

狀態檔 siliconomics/data/baseline_status.json 不存在時只印「未啟用」並以 0 結束。
輸出兩類：「不符」（以 1 結束）與「提醒 [n]」（不影響結束碼）。

用法：python3 tools/baseline_coordination/check.py [--status 路徑] [--today YYYY-MM-DD]
"""
import argparse, datetime, json, pathlib, re, sys

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
            + check_version_drift(st) + check_recon_residual() + check_evidence_range())
    for n in NOTES:
        print(n)
    for e in errs:
        print(e)
    print(f"三層基準協調檢查：{len(errs)} 項不符、{len(NOTES)} 項提醒")
    return 1 if errs else 0


if __name__ == "__main__":
    sys.exit(main())
