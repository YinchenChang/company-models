"""三層基準協調檢查（草稿；規則見 docs/plan/三層基準協調規則.md 第 10 節）。

步驟 2 起草、步驟 5 啟用。狀態檔 siliconomics/data/baseline_status.json
尚未建立時（步驟 3 量測後才建），只印「未啟用」並以 0 結束，不擋任何 PR。

用法：python3 tools/baseline_coordination/check.py [--status 路徑] [--today YYYY-MM-DD]
"""
import argparse, datetime, json, pathlib, re, sys

ROOT = pathlib.Path(__file__).resolve().parents[2]
DEFAULT_STATUS = ROOT / "siliconomics/data/baseline_status.json"
VALID = {"甲", "乙", "丙"}
TOL = 0.005  # 第 2 項：換算值比對容許的相對誤差（四捨五入）


def expected_status(w_share, s_share, thr):
    """C3：兩層證據比例對門檻 → 甲／乙／丙。"""
    w, s = w_share >= thr, s_share >= thr
    return "甲" if (w and s) else ("乙" if (w or s) else "丙")


def check_status(st, today):
    errs = []
    thr = st["threshold"]
    for yr, y in sorted(st["years"].items()):
        tag = y.get("status")
        # 1. 有標籤且與證據比例一致
        if tag not in VALID:
            errs.append(f"[1] {yr} 缺狀態標籤或標籤無效：{tag!r}")
        else:
            exp = expected_status(y["wattonomics"]["evidence_share"],
                                  y["siliconomics"]["evidence_share"], thr)
            if tag != exp:
                errs.append(f"[1] {yr} 標「{tag}」，依證據比例與門檻 {thr:.0%} 應為「{exp}」")
        # 2. 丙狀態：Siliconomics 基準＝換算值（以供給上限為頂）
        if tag == "丙":
            conv = min(y["converted"], y.get("supply_cap", float("inf")))
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
    recon_years = {yr for yr, y in st["years"].items() if y.get("reconciling")}
    for co, years in st.get("depends_on_years", {}).items():
        if not recon_years & set(years):
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
    errs = check_status(st, datetime.date.fromisoformat(a.today)) + check_company_reads(st)
    for e in errs:
        print(e)
    print(f"三層基準協調檢查：{len(errs)} 項不符")
    return 1 if errs else 0


if __name__ == "__main__":
    sys.exit(main())
