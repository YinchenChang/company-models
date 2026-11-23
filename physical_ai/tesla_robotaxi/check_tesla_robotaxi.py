# Tesla Robotaxi 分部模型核對：LibreOffice 重算 → 0 錯誤值、公式值＝Python 引擎值、FCFF 檢查列＝0、三支 repo 核對程式。
# 用法：python3 physical_ai/tesla_robotaxi/check_tesla_robotaxi.py
import glob, json, os, subprocess, sys, tempfile
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
src = sorted(glob.glob(os.path.join(HERE, "dist", "*.xlsx")))[-1]
tmp = tempfile.mkdtemp()
subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", tmp, src], check=True, capture_output=True)
rc = os.path.join(tmp, os.path.basename(src))
wb = openpyxl.load_workbook(rc, data_only=True)
fail = []
errs = []
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("#") and c.value.rstrip("!?0/").upper() in ("#DIV", "#VALUE", "#REF", "#NAME", "#N/A", "#NUM", "#NULL", "#ERR") or (isinstance(c.value, str) and c.value[:4] in ("#DIV", "#VAL", "#REF", "#NAM", "#N/A", "#NUM", "Err:")):
                errs.append(f"{ws.title}!{c.coordinate}={c.value}")
if errs:
    fail.append(f"錯誤值 {len(errs)} 格：" + "、".join(errs[:10]))
snap = json.load(open(os.path.join(HERE, "engine_snapshot.json")))
V = wb["估值"]; r = snap["ev_row"]
for col, key in (("C", "A"), ("D", "B"), ("E", "C")):
    x = V[f"{col}{r}"].value
    if not isinstance(x, (int, float)) or abs(x - snap["scn"][key]) > 1e-6 * max(1, abs(x)):
        fail.append(f"估值 情境 {key}：xlsx {x} ≠ 引擎 {snap['scn'][key]}")
x = V[f"F{r}"].value
if abs(x - snap["weighted"]) > 1e-6 * abs(x):
    fail.append(f"加權：xlsx {x} ≠ 引擎 {snap['weighted']}")
# FCFF 檢查列
for row in V.iter_rows():
    if str(row[0].value or "").startswith("檢查：各列加總"):
        for c in row[3:]:
            if isinstance(c.value, (int, float)) and abs(c.value) > 1e-9:
                fail.append(f"FCFF 檢查列非 0：{c.coordinate}")
T = wb["給Tesla模型"]
def named(nm):
    dest = list(wb.defined_names[nm].destinations)[0]
    return wb[dest[0]][dest[1].replace("$", "")].value
for nm in ("RT_VAL_LO", "RT_VAL_BASE", "RT_VAL_HI", "RT_VAL_A", "RT_VAL_B", "RT_VAL_C"):
    if nm not in wb.defined_names:
        fail.append(f"缺具名儲存格 {nm}")
for s_ in "ABC":
    x = named(f"RT_VAL_{s_}")
    if abs(x - snap["scn"][s_]) > 1e-6 * max(1, abs(x)):
        fail.append(f"RT_VAL_{s_}：xlsx {x} ≠ 引擎 {snap['scn'][s_]}")
    for y in range(2027, 2036):
        nm = f"RT_REV_{s_}_{y}"
        if nm not in wb.defined_names:
            fail.append(f"缺具名儲存格 {nm}"); continue
        x = named(nm); e = snap["rev"][s_][str(y)] * 1000
        if abs(x - e) > 1e-6 * max(1, abs(e)):
            fail.append(f"{nm}：xlsx {x} ≠ 引擎 {e}")
# 與對帳 2 預設變體對照（市場、滲透、份額讀線；每趟英里丙；誘發相對 2026、ε 讀線）：加權 38.98、A 164.11、B −0.63、C −6.91（容許 ±0.5）
for key, ref in (("A", 164.11), ("B", -0.63), ("C", -6.91)):
    if abs(snap["scn"][key] - ref) > 0.5:
        fail.append(f"對帳 2 預設變體：情境 {key} {snap['scn'][key]:.2f} 與 {ref} 差超過 0.5")
if abs(snap["weighted"] - 38.98) > 0.5:
    fail.append(f"對帳 2 預設變體：加權 {snap['weighted']:.2f} 與 38.98 差超過 0.5")
# In 頁線讀入值＝線成品 xlsx 的值（獨立重算：不經建置程式的 load_line）
import importlib.util
spec = importlib.util.spec_from_file_location("trb", os.path.join(HERE, "build_tesla_robotaxi.py"))
trb = importlib.util.module_from_spec(spec); spec.loader.exec_module(trb)
Iw = wb["In"]; inval = {Iw.cell(r, 1).value: Iw.cell(r, 5).value for r in range(4, Iw.max_row + 1) if Iw.cell(r, 1).value}
nline = 0
for kind, key in (("price", "lp"), ("opx", "lx"), ("pmd", "lm")):
    for s_ in "ABC":
        for y in range(2026, 2036):
            v = inval.get(f"{key}_{y}_{s_}"); nline += 1
            if v is None or abs(v - trb.LINE[kind][s_][y]) > 1e-9:
                fail.append(f"In {key}_{y}_{s_}：{v} ≠ 線讀值 {trb.LINE[kind][s_][y]}")
for kind, key in (("trips", "lt"), ("pen", "lpen"), ("tsh", "lts")):
    for s_ in "ABC":
        for y in range(2026, 2036):
            v = inval.get(f"{key}_{y}_{s_}"); nline += 1
            if v is None or abs(v - trb.LINE["mkt"][(kind, s_, y)]) > 1e-9:
                fail.append(f"In {key}_{y}_{s_}：{v} ≠ 線讀值 {trb.LINE['mkt'][(kind, s_, y)]}")
# 份額加總（對帳 2 議題 1）：線的 Waymo＋Tesla＋其他＝100%，且 Tesla 讀值＝線 Tesla 份額
lw = openpyxl.load_workbook(sorted(glob.glob(trb.LINE_GLOB))[-1], data_only=True)["情境"]
for s_, b in trb.MKT_BLK.items():
    for i in range(10):
        tot = sum(lw.cell(b + o, 3 + i).value for o in (27, 28, 29))
        if abs(tot - 1) > 1e-9:
            fail.append(f"線份額加總 {s_} {2026 + i}：{tot}")
for s_ in "ABC":   # 錨點年：價格需落在 Model Y 與 Cybercab 價格之間（加權權重 0–1 之外即失敗）
    for y in trb.LINE_ANCH:
        w = trb.LINE["w"][(s_, y)]
        if not (-1e-9 <= w <= 1 + 1e-9):
            fail.append(f"車隊權重 w_MY {s_} {y} 超出 0–1：{w}")
if abs(inval.get("cf", 0) - 1.0) > 1e-12:
    fail.append(f"cf 基準應為 1.0：{inval.get('cf')}")
lo, base, hi = T["C4"].value, T["C5"].value, T["C6"].value
if not (lo <= base <= hi):
    fail.append(f"低／基準／高順序不符：{lo} {base} {hi}")
if wb["Sens"]["D4"].value != "CURRENT":
    fail.append(f"Sens 模型狀態 {wb['Sens']['D4'].value}")
print("重算檔：", rc)
print("分部價值（十億美元）低／基準／高：", round(lo, 2), round(base, 2), round(hi, 2))
print("情境 A／B／C：", {k: round(v, 2) for k, v in snap["scn"].items()}, "引擎加權", round(snap["weighted"], 2), f"；線讀入 {nline} 格與線成品一致")
for tool, args in (("tools/in_layout/check_in_layout.py", []), ("tools/sens_coverage/check_sens_coverage.py", ["--require-keycol"]),
                   ("tools/model_governance/check_governance.py", [])):
    p = subprocess.run([sys.executable, os.path.join(ROOT, tool), rc] + args, capture_output=True, text=True)
    print(f"== {tool}：結束碼 {p.returncode}\n{p.stdout[-1500:]}{p.stderr[-800:]}")
    if p.returncode != 0:
        fail.append(f"{tool} 失敗")
print("核對：" + ("全部通過" if not fail else "失敗\n- " + "\n- ".join(fail)))
sys.exit(1 if fail else 0)
