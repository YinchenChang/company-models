# 核對列與錯誤值掃描（2026-11-20；起因：外部稽核指出台積電、智邦、Astera Labs 的「應為 0」核對列非零，
# 而現有檢查都沒有掃這類列）。只讀快取值，不改任何檔案。
# 用法：
#   python3 tools/model_governance/check_zero_rows.py <成品.xlsx> [...] [--tol 0.005] [--tol-round 0.2]
#       掃指定活頁簿；有非零核對列或錯誤值時以非 0 結束。
#   python3 tools/model_governance/check_zero_rows.py --all [--baseline tools/model_governance/zero_rows_baseline.json]
#       掃 repo 內全部 dist／model 現行成品（略過 _待刪除、archive）。有 --baseline 時，只有「不在基準清單內」的
#       項目才算失敗（基準清單＝已知待修項目；修好一項就從清單刪一項，清單只減不增）。
#   python3 tools/model_governance/check_zero_rows.py --all --write-baseline <路徑>
#       把目前的失敗項目寫成基準清單。
# 判定：
#   核對列＝前三欄任一格的文字含「應為 0」（容許全形／空白）且同列有數值；或緊接在「應為 0」節標題
#       （同列無數值）之下、以「核對：」「檢查：」開頭的連續列。
#       不算的列：文字以「差額」開頭（刻意揭露兩種推法差距）；
#       　　外部稽核優化 W3-2（2026-11-23）：標準損益表延伸期的「差額」列已改為「核對：延伸期毛利 − 營業費用 − 營業利益（…應為 0）」，
#       　　延伸期營業費用＝毛利 − 營業利益，該列現在會被本工具掃到；其下「揭露（不是核對列）：隱含營業費用率」不是核對列。
#       「應接近 0」（逆向 DCF 的搜尋殘差，落在搜尋下界時不為 0）；Sens 頁（建置快照，重複 PL 頁的核對列）。
#       只有部分欄或部分情境應為 0、「A vs B」並列原值、明寫「非應為 0」的列也不判（無法逐列自動判定）。
#       列內任一數值的絕對值 > tol 即失敗。標籤寫「捨入」或「四捨五入」的列用 --tol-round（來源數字四捨五入）；
#       標籤寫明容差數字（例：「四捨五入差 < 0.01」）時用標籤的數字。
#   錯誤值＝任何工作表、任何儲存格的快取值為 #REF!、#NAME?、#VALUE!、#DIV/0!、#N/A、#NUM!、#NULL! 或以 Err: 開頭。
import argparse, glob, json, os, re, sys
import openpyxl

LABEL = re.compile(r"應為\s*[0０]")
# 無法逐列自動判定的寫法：只有部分欄或部分情境應為 0、「A vs B」並列原值與差、明寫「非應為 0」
PARTIAL = re.compile(r"非應為|[A-Z]\s*[–-]\s*[A-Z]\s*應為|情境\s*[ABC]\s*應為|\bvs\b")
# 來源數字四捨五入造成的差：「捨入」或「四捨五入」；後面接「差 < 數字」時，該數字就是這一列的容差（W2-3）
ROUNDING = re.compile(r"捨入|四捨五入")
DECLARED = re.compile(r"(?:捨入|四捨五入)差?\s*[<＜≤]\s*([0-9]*\.?[0-9]+)")
PREFIX = ("核對：", "檢查：", "核對:", "檢查:")
ERRS = ("#REF!", "#NAME?", "#VALUE!", "#DIV/0!", "#N/A", "#NUM!", "#NULL!")
SKIP_DIR = ("_待刪除", "archive", "legacy")


def declared_tol(label, tol, tol_round):
    """標籤自己寫明容差時用標籤的值（例：「四捨五入差 < 0.01」→ 0.01）；只寫「捨入」或「四捨五入」而沒有數字時用 tol_round；其餘用 tol。"""
    m = DECLARED.search(label)
    if m:
        return float(m.group(1))
    return tol_round if ROUNDING.search(label) else tol


def is_check_label(s):
    s = s.strip()
    if s.startswith("差額"):
        return False
    return bool(LABEL.search(s)) or s.startswith(PREFIX)


def scan(path, tol, tol_round):
    """回傳 [(類別, 工作表, 列號, 標籤, 值)]；類別為 zero 或 error。"""
    out = []
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    for ws in wb.worksheets:
        block = False  # 是否在「應為 0」節標題之下
        for r, row in enumerate(ws.iter_rows(values_only=True), start=1):
            for v in row:
                if isinstance(v, str) and (v in ERRS or v.startswith("Err:")):
                    out.append(("error", ws.title, r, v, v))
            if ws.title == "Sens":
                continue
            first = next((v.strip() for v in row[:3] if isinstance(v, str) and v.strip()), None)
            nums = [v for v in row if isinstance(v, (int, float)) and not isinstance(v, bool)]
            if first is None:
                block = False
                continue
            has0 = bool(LABEL.search(first)) and not first.startswith("差額") and not PARTIAL.search(first)
            if has0 and not nums:
                block = True
                continue
            in_block = block and first.startswith(PREFIX)
            if not first.startswith(PREFIX) and not has0:
                block = False
            if not (has0 or in_block) or not nums:
                continue
            t = declared_tol(first, tol, tol_round)
            worst = max(nums, key=abs)
            if abs(worst) > t:
                out.append(("zero", ws.title, r, first[:60], worst))
    wb.close()
    return out


def current_files(root):
    files = []
    for pat in ("**/dist/*.xlsx", "*/model/*.xlsx"):
        for p in glob.glob(os.path.join(root, pat), recursive=True):
            rel = os.path.relpath(p, root)
            if any(part in SKIP_DIR or part.startswith("legacy") for part in rel.split(os.sep)):
                continue
            # */model/ 只掃 CURRENT 指到的檔
            d = os.path.dirname(p)
            cur = os.path.join(d, "CURRENT")
            if os.path.basename(d) == "model" and os.path.exists(cur):
                if open(cur, encoding="utf-8").read().strip() != os.path.basename(p):
                    continue
            files.append(rel.replace(os.sep, "/"))
    return sorted(set(files))


def key(path, rec):
    # 基準清單以「資料夾｜工作表｜標籤」為鍵，不含檔名版本與列號，升版或插列後仍對得上
    folder = path.split("/dist/")[0].split("/model/")[0]
    return f"{folder}|{rec[1]}|{rec[3]}|{rec[0]}"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="*")
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--tol", type=float, default=0.005)
    ap.add_argument("--tol-round", type=float, default=0.2)
    ap.add_argument("--baseline")
    ap.add_argument("--write-baseline")
    a = ap.parse_args()
    root = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
    files = current_files(root) if a.all else [f.replace(os.sep, "/") for f in a.files]
    if not files:
        ap.error("沒有要掃的檔案")
    known = set()
    if a.baseline and os.path.exists(a.baseline):
        known = set(json.load(open(a.baseline, encoding="utf-8"))["known"])
    found, seen = {}, set()
    for f in files:
        recs = scan(os.path.join(root, f) if a.all else f, a.tol, a.tol_round)
        if recs:
            found[f] = recs
        for rec in recs:
            k = key(f, rec)
            seen.add(k)
            flag = "已知" if k in known else "新增"
            print(f"[{flag}] {f} {rec[1]}!R{rec[2]} {rec[3]} = {rec[4]}")
    if a.write_baseline:
        json.dump({"說明": "check_zero_rows.py 的已知待修項目；只減不增。鍵＝資料夾|工作表|標籤|類別",
                   "tol": a.tol, "tol_round": a.tol_round, "known": sorted(seen)},
                  open(a.write_baseline, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    new = len(seen - known)
    stale = sorted(known - seen) if a.all else []
    for k in stale:
        print(f"[已修好，請從基準清單刪除] {k}")
    print(f"掃描 {len(files)} 檔；有問題 {len(found)} 檔；項目 {len(seen)}（新增 {new}、已知 {len(seen) - new}）；基準清單可刪 {len(stale)}")
    sys.exit(1 if new or stale else 0)


if __name__ == "__main__":
    main()
