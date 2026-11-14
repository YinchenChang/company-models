# NVIDIA 三層協調試作（2026-11-14）：另一側讀數整本重建快照 → nvidia/coord_alt.json（正式建置的 Coord 頁讀入）。
# 做法：同一個建置程式，NV_OVERRIDE="_coord=1"——三情境 CY2027 起運算顆數與網通營收 × co_r_<yy>（In 頁 Y 節：另一側 ÷ 現行基準），
# 其他輸入不變；LibreOffice 重算後讀加權目標價、DCF、Comparable（調整後）、前瞻 EPS、三情境 EPS、情境 B 營收與顆數。基準檔同時重建一次作核對。
# In 頁 Y 節（coord_inputs.json 或 Siliconomics 快照）改動後先重跑本檔，再正式建置（否則 Coord 頁快照狀態顯示 STALE）。
# 用法：python3 nvidia/coord_alt.py [暫存資料夾，預設 /tmp/work_nvidia/coord]
import glob, json, os, shutil, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
TMP = sys.argv[1] if len(sys.argv) > 1 else "/tmp/work_nvidia/coord"
import coord_post as CO  # noqa: E402

env0 = {k: v for k, v in os.environ.items() if k not in ("NV_OVERRIDE", "NV_V36_PARTS", "NV_ROWMAP")}
raw, rc = os.path.join(TMP, "raw"), os.path.join(TMP, "rc")
shutil.rmtree(TMP, ignore_errors=True); os.makedirs(raw); os.makedirs(rc)
for nm, ov in (("base", "_noop=1"), ("alt", "_coord=1")):   # 兩者都設覆寫（不跑 v037 以後的展示層後處理），只差 _coord
    subprocess.run([sys.executable, os.path.join(HERE, "build_nvidia.py"), os.path.join(raw, f"{nm}.xlsx")], check=True,
                   env={**env0, "NV_OVERRIDE": ov}, stdout=subprocess.DEVNULL)
subprocess.run(["soffice"] + ([f"-env:UserInstallation={os.environ['LO_PROFILE']}"] if os.environ.get("LO_PROFILE") else [])
               + ["--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", rc] + sorted(glob.glob(os.path.join(raw, "*.xlsx"))),
               check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def read(nm):
    wb = openpyxl.load_workbook(os.path.join(rc, f"{nm}.xlsx"), data_only=True)
    L = CO.locate(wb)
    out = {k: wb[s][c].value for k, (s, c) in L.items()}
    s, c = next(iter(wb.defined_names["tp_comp_adj"].destinations)); out["comp"] = wb[s][c.replace("$", "")].value
    ratio = {}
    for y in CO.APPLY:
        s, c = next(iter(wb.defined_names[f"co_r_{y % 100}"].destinations)); ratio[str(y)] = wb[s][c.replace("$", "")].value
    errs = sum(1 for ws in wb if ws.title != "README" for row in ws.iter_rows() for x in row if isinstance(x.value, str) and x.value.startswith("#") and x.value.rstrip("!?0/").upper() in ("#REF", "#DIV", "#VALUE", "#NAME", "#N/A", "#NUM", "ERR:502", "#NULL"))
    return out, ratio, errs   # README 頁在覆寫建置有一格依賴 Sens 快照的文字公式（#VALUE!），不影響數值；不計


b, rb, eb = read("base"); a, ra, ea = read("alt")
assert eb == 0 and ea == 0, (eb, ea)
assert rb == ra
out = {"_說明": "另一側讀數整本重建快照（nvidia/coord_alt.py）；Coord 頁 B 節讀入", "ratio": ra, "base": b, "alt": a}
json.dump(out, open(os.path.join(HERE, "coord_alt.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
for k in ("tp", "dcf", "comp", "fwd", "up", "eps_B_2028", "eps_B_2029", "eps_B_2030", "rev_B_2028", "nvu_B_2028"):
    print(f"{k:12s} 現行 {b[k]:10.4f}  另一側 {a[k]:10.4f}  {a[k] / b[k] - 1:+.2%}")
