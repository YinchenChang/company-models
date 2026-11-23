# 線 2a 工業機器人 v0.1 核對程式
# 用法：python3 physical_ai/lines/industrial_robot/check_industrial_robot.py
# 1) LibreOffice headless 重算 dist 內 xlsx（覆寫為含快取值版本）2) 0 錯誤值 3) 對帳與合理性檢查 4) 印出關鍵下游取數
import os, sys, glob, shutil, subprocess, tempfile
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
f = sorted(glob.glob(os.path.join(HERE, "dist", "*_工業機器人單位經濟_v*.xlsx")))[-1]
tmp = tempfile.mkdtemp()
subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", tmp, f],
               check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
out = os.path.join(tmp, os.path.basename(f))
shutil.copy(out, f)

wbv = openpyxl.load_workbook(f, data_only=True)
wbf = openpyxl.load_workbook(f, data_only=False)
ERR = ("#VALUE!", "#REF!", "#NAME?", "#DIV/0!", "#N/A", "#NUM!", "#NULL!", "Err:")
fails = []
nform = 0
for ws in wbf.worksheets:
    wv = wbv[ws.title]
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                nform += 1
                v = wv[c.coordinate].value
                if v is None:
                    fails.append(f"{ws.title}!{c.coordinate} 無快取值")
                elif isinstance(v, str) and v.startswith(ERR):
                    fails.append(f"{ws.title}!{c.coordinate} {v}")

def find(ws, col, text):
    for row in ws.iter_rows():
        for c in row:
            if c.column_letter == col and c.value == text:
                return c.row
    return None

T = wbv["敏感度"]
for row in T.iter_rows():
    for c in row:
        if c.value in ("一致", "不一致") and c.value == "不一致":
            fails.append(f"敏感度!{c.coordinate} 基準與情境／計算頁不一致")
C = wbv["計算"]
def cval(key):
    r = find(C, "A", key); return C[f"C{r}"].value
if cval("MIX_OTHER") < 0: fails.append("機型結構：其他占比 < 0")
for k in ("RC_RV_RATIO", "RC_HD_RATIO"):
    v = cval(k)
    if not (0.3 <= v <= 1.0): fails.append(f"{k}={v:.2f} 不在 0.3–1.0")
O = wbv["下游取數"]
ids = {}
for row in O.iter_rows(min_row=5):
    if row[0].value:
        if row[0].value in ids: fails.append(f"下游 ID 重複 {row[0].value}")
        ids[row[0].value] = row[4].value
need = ["IR.inst.global.2030.base", "IR.inst.china.2030.base", "IR.inst.cobot.2030.base", "IR.red.total.mkt.2030.base",
        "IR.unit.N_RV_ART.base", "IR.unit.N_HD_COBOT.base", "IR.cycle.jara.orders.latest", "IR.cycle.jara.b2b.latest"]
for n in need:
    if n not in ids: fails.append(f"缺下游 ID {n}")
# 單調：低 ≤ 基準 ≤ 高
for y in range(2026, 2033):
    for s in ("inst.global", "red.total.mkt", "inst.cobot"):
        lo, b, hi = (ids[f"IR.{s}.{y}.{x}"] for x in ("lo", "base", "hi"))
        if not (lo <= b <= hi): fails.append(f"{s} {y} 情境非單調 {lo},{b},{hi}")

print(f"檔案：{f}\n公式數：{nform}")
for n in ["IR.inst.global.2025.base", "IR.inst.global.2030.lo", "IR.inst.global.2030.base", "IR.inst.global.2030.hi",
          "IR.inst.china.2030.lo", "IR.inst.china.2030.base", "IR.inst.china.2030.hi",
          "IR.inst.cobot.2030.lo", "IR.inst.cobot.2030.base", "IR.inst.cobot.2030.hi",
          "IR.red.total.mkt.2025.base", "IR.red.total.mkt.2030.lo", "IR.red.total.mkt.2030.base", "IR.red.total.mkt.2030.hi",
          "IR.red.rv.units.2025.base", "IR.red.hd.units.2025.base", "IR.red.rv.mkt.2025.base", "IR.red.hd.mkt.2025.base",
          "IR.unit.U_RV_TRAD", "IR.unit.U_HD_TRAD", "IR.payback.cn6ax", "IR.payback.for6ax", "IR.payback.cncobot",
          "IR.cycle.jara.b2b.latest", "IR.cycle.jara.orders.ltm"]:
    print(f"  {n} = {ids.get(n)}")
for k in ("BOM_TD", "BOM_BU", "BOM_RATIO", "RC_RV_M", "RC_NAB_IMP", "RC_RV_RATIO", "RC_HD_M", "RC_HDS_IMP", "RC_HD_RATIO", "RC_RED_SH", "RC_NBS_IFR", "CY_G26_GAP"):
    print(f"  計算 {k} = {cval(k)}")
print("敏感度：")
for row in T.iter_rows(min_row=4):
    if isinstance(row[0].value, int) and row[0].value <= 5:
        print(f"  #{row[0].value} {row[2].value}: 低 {row[5].value:,.2f} 高 {row[6].value:,.2f} 擺幅÷基準 {row[8].value:.0%}")
    elif isinstance(row[0].value, str) and row[0].value.startswith(("輸出", "基準")):
        print(" ", row[0].value, row[3].value if row[0].value == "基準值" else "")
if fails:
    print("失敗：", *fails[:50], sep="\n  "); sys.exit(1)
print("全部通過：0 錯誤值")
