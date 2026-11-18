# 用兩層活頁簿自己的公式做單因子重算（LibreOffice UNO；只在 /tmp 開複本，不存檔）。
# 輸出 /home/claude/out/sens/scripts/uno_results.json
import uno, subprocess, time, sys, os, json, shutil
from com.sun.star.beans import PropertyValue
W = "/home/claude/wt/sens/wattonomics/model/20261116_Wattonomics_v0.3.6.xlsx"
S = "/home/claude/wt/sens/siliconomics/model/20261116_Siliconomics_v0.20.xlsx"
for p in (W, S): shutil.copy(p, "/tmp/sens/" + os.path.basename(p))
W = "/tmp/sens/" + os.path.basename(W); S = "/tmp/sens/" + os.path.basename(S)
port = 2300 + os.getpid() % 300
proc = subprocess.Popen(['soffice', '-env:UserInstallation=file:///tmp/sens/lo_prof', '--headless', '--invisible', '--norestore',
                         f'--accept=socket,host=localhost,port={port};urp;'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
ctx = None
for i in range(90):
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
        ctx = resolver.resolve(f'uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext'); break
    except Exception:
        time.sleep(1)
desktop = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
def pv(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p
def cells(doc, name):
    return doc.NamedRanges.getByName(name).getReferredCells()
def get(doc, name):
    return [x for row in cells(doc, name).getDataArray() for x in row]
def setv(doc, name, v):
    c = cells(doc, name); c.getCellByPosition(0, 0).setValue(v)
def run(path, scen, outs):
    doc = desktop.loadComponentFromURL('file://' + path, '_blank', 0, (pv('Hidden', True),))
    res = {}
    for sname, sets in scen.items():
        orig = {n: get(doc, n)[0] for n in sets}
        for n, v in sets.items(): setv(doc, n, v)
        doc.calculateAll()
        res[sname] = {o: get(doc, o) for o in outs}
        for n, v in orig.items(): setv(doc, n, v)
    doc.close(True)
    return res
WOUT = ["WIF_MktGW_Add_2027", "WIF_MktGW_Add_2028", "WIF_Recon_NV_ShipGW_2027", "WIF_Recon_NV_NVGW_2027"]
wsc = {"base": {}, "ai_hi": {"ai_sh": 0.95, "ai_amzn": 0.75}, "ai_lo": {"ai_sh": 0.75, "ai_amzn": 0.6},
       "lag_on_lo": {"lag_on": 3}, "lag_on_hi": {"lag_on": 6}, "inv_lo": {"inv_on": 0.05}, "inv_hi": {"inv_on": 0.2},
       "bldg_lo": {"lag_bldg": 12}, "bldg_hi": {"lag_bldg": 24}, "oth_lo": {"oth_gmult": 0}, "oth_hi": {"oth_gmult": 1.5},
       # TK 功率低端（VR200 188 kW、RU 300 kW、GB300 132 kW）：每 GW 顆數與 IT capex 依機架數等比放大（近似）
       "pow_lo": {"TK_CapexIT_VR200": 37.586352 * 227 / 188, "TK_CapexIT_RU": 50.08578 * 340 / 300, "TK_CapexIT_GB300": 32.237224 * 136 / 132},
       "cost_lo": {"TK_CapexIT_VR200": 31.9970232, "TK_CapexIT_RU": 35.845578, "TK_CapexIT_GB300": 28.5549024,
                   "TK_CapexFacility_VR200": 10.212, "TK_CapexFacility_RU": 10.212, "TK_CapexFacility_GB300": 10.212}}
SOUT = ["SIF_DeployGW_Base", "SIF_DeployGWPhys_Base"]
ssc = {"base": {}, "rb_67": {"u_rb_27": 6.7 * 0.9404}, "nv_x0884": {"u_rb_27": 6.582652 * 0.884, "u_bw_27": 2.256909 * 0.884, "u_ru_27": 0.282114 * 0.884},
       "pow_lo": {"TK_CapexIT_VR200": 37.586352 * 227 / 188, "TK_GPUsPerGW_VR200": 291744 * 227 / 188, "TK_CapexIT_RU": 50.08578 * 340 / 300,
                  "TK_GPUsPerGW_RU": 194760 * 340 / 300, "TK_CapexIT_GB300": 32.237224 * 136 / 132, "TK_GPUsPerGW_GB300": 487008 * 136 / 132},
       "cost_lo": {"TK_CapexIT_VR200": 31.9970232, "TK_CapexIT_RU": 35.845578, "TK_CapexIT_GB300": 28.5549024},
       "it_tpu_hi": {"it_TPU": 35}}
t0 = time.time()
out = {"wat": run(W, wsc, WOUT), "sil": run(S, ssc, SOUT), "sec": None}
out["sec"] = time.time() - t0
json.dump(out, open("/home/claude/out/sens/scripts/uno_results.json", "w"), ensure_ascii=False, indent=1)
try: desktop.terminate()
except Exception: pass
proc.wait(timeout=60)
print(json.dumps(out, ensure_ascii=False)[:3000])
