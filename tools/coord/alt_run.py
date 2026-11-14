# 三層協調 C4：另一側讀數整本重建 → <資料夾>/coord_alt.json（tools/coord/coord.py 說明流程）。
# 用法：python3 tools/coord/alt_run.py <資料夾> [暫存資料夾，預設 /tmp/work_coord/<資料夾>]
# 建檔指令取自 .github/workflows/ai-semis-build.yml 的 targets（與 CI 相同）。建置兩次：現行（COORD_RUN＝1）與另一側（另加 SIL_COORD_ALT＝1），
# LibreOffice 重算後讀 Val 頁加權目標價、DCF、Comparable（市場主流同業法）、上檔空間；錯誤格須為 0。
import glob, json, os, re, shutil, subprocess, sys
import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import coord as CO  # noqa: E402

folder = sys.argv[1].rstrip("/")
TMP = sys.argv[2] if len(sys.argv) > 2 else f"/tmp/work_coord/{folder}"
wf = open(os.path.join(ROOT, ".github", "workflows", "ai-semis-build.yml"), encoding="utf-8").read()
m = re.search(r'"' + re.escape(folder) + r'\|(python3 [^|"]+)\|', wf)
assert m, f"workflow targets 找不到 {folder}"
cmd = m.group(1).split()
shutil.rmtree(TMP, ignore_errors=True); os.makedirs(os.path.join(TMP, "raw")); os.makedirs(os.path.join(TMP, "rc"))
env0 = {k: v for k, v in os.environ.items() if k not in ("SIL_COORD_ALT",)}
env0["COORD_RUN"] = "1"; env0["PYTHONDONTWRITEBYTECODE"] = "1"
for nm, extra in (("base", {}), ("alt", {"SIL_COORD_ALT": "1"})):
    subprocess.run(cmd + [os.path.join(TMP, "raw", f"{nm}.xlsx")], check=True, cwd=ROOT, env={**env0, **extra}, stdout=subprocess.DEVNULL)
subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", os.path.join(TMP, "rc")]
               + sorted(glob.glob(os.path.join(TMP, "raw", "*.xlsx"))), check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
ERR = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!", "Err:502")


def read(nm):
    wb = openpyxl.load_workbook(os.path.join(TMP, "rc", f"{nm}.xlsx"), data_only=True)
    V = CO.val_rows(wb["Val"])
    out = {k: wb["Val"].cell(V[k], 2).value for k in CO.KEYS}
    out["rec"] = CO.rec(out["up"])
    out["errors"] = sum(1 for ws in wb for row in ws.iter_rows() for x in row if isinstance(x.value, str) and x.value in ERR)
    w = {y: wb.defined_names.get(f"wta_{y % 100}") for y in (2027, 2028, 2029)}
    out["wta"] = {}
    for y, d in w.items():
        if d:
            s, c = next(iter(d.destinations)); out["wta"][str(y)] = wb[s][c.replace("$", "")].value
    return out


B, A = read("base"), read("alt")
assert B["errors"] == 0 and A["errors"] == 0, (B["errors"], A["errors"])
sys.path.insert(0, os.path.join(ROOT, "siliconomics"))
import reader as SIL  # noqa: E402
exp = {str(y): (SIL.ALT_RATIO[y] if SIL.RECO_FLAG[y] == 1 else 1) for y in (2027, 2028, 2029)}
assert A["wta"] and all(abs(A["wta"][y] - exp[y]) < 1e-12 for y in exp), (A["wta"], exp)   # 另一側建置確實讀到另一側比值
out = {"_說明": "三層協調 C4 兩讀數整本重建快照（tools/coord/alt_run.py）；建檔程式以 tools/coord/coord.py add_sheet 寫入 Coord 頁",
       "folder": folder, "snapshot": f"Siliconomics {SIL.VERSION}（{SIL.FILE}）", "ratio": exp,
       "base": {k: B[k] for k in (*CO.KEYS, "rec")}, "alt": {k: A[k] for k in (*CO.KEYS, "rec")}}
json.dump(out, open(os.path.join(ROOT, folder, "coord_alt.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
for k in (*CO.KEYS, "rec"):
    print(f"{k:5s} 現行 {B[k]!s:>22}  另一側 {A[k]!s:>22}")
