"""CoreWeave v5.0 整本重算工具：以 SENS_OVR／CRWV_TKV 建置多個變體 xlsx，分批交給 LibreOffice headless 重算（每批一個 soffice 行程、獨立使用者設定檔），再讀關鍵輸出。
sens_v50.py、reverse_v50.py、decomp_v50.py 共用。"""
import json, os, subprocess, sys, tempfile, shutil, concurrent.futures as cf
import openpyxl

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BUILD = os.path.join(HERE, "build_coreweave.py")


def _find(ws, prefix, col=1):
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.startswith(prefix): return r
    raise KeyError((ws.title, prefix))


def read(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    V, F, S, T = wb["Val"], wb["Fund"], wb["Sens"], wb["Thesis"]
    g = lambda ws, lab, c=3: ws.cell(_find(ws, lab), c).value
    tb = _find(T, "情境 B（基準")
    pm_r = next(r for r in range(tb, T.max_row + 1) if str(T.cell(r, 1).value or "").startswith("每 MW 年收入（錨"))
    return {"tp": g(V, "目標價（DCF 與 Comparable 加權）"), "dcf": g(V, "DCF 目標價（情境加權）"), "comp": g(V, "Comparable 目標價（情境加權）"),
            "rec": g(V, "投資建議"), "gate": g(F, "收支關卡：", 4), "fp": g(S, "輸入指紋（本檔即時重算", 4), "pm30": T.cell(pm_r, 9).value}


def run(cases, jobs=2):
    """cases：[(名稱, ovr dict, env dict)] → {名稱: 輸出}"""
    wd = tempfile.mkdtemp(prefix="crwv_run_")
    files = {}
    for i, (name, ovr, env) in enumerate(cases):
        p = os.path.join(wd, f"c{i:03d}.xlsx")
        e = dict(os.environ); e.pop("CRWV_ROWMAP", None); e["SENS_OVR"] = json.dumps(ovr or {"_": 0}); e.update(env or {})
        subprocess.run([sys.executable, BUILD, p], env=e, check=True, stdout=subprocess.DEVNULL)
        files[name] = p
    names = list(files)
    chunks = [names[k::jobs] for k in range(jobs)]
    def conv(k):
        out = os.path.join(wd, "out"); os.makedirs(out, exist_ok=True)
        if not chunks[k]: return
        subprocess.run(["soffice", f"-env:UserInstallation=file:///tmp/lo_crwv_run{k}", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", out] + [files[n] for n in chunks[k]],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    with cf.ThreadPoolExecutor(jobs) as ex: list(ex.map(conv, range(jobs)))
    res = {n: read(os.path.join(wd, "out", os.path.basename(files[n]))) for n in names}
    shutil.rmtree(wd, ignore_errors=True)
    return res


def in_ranges():
    """回傳 In 頁每個鍵的（低, 基準, 高）（基準檔，公式字串原樣）。"""
    wd = tempfile.mkdtemp(prefix="crwv_in_")
    p = os.path.join(wd, "b.xlsx"); e = dict(os.environ); e.pop("CRWV_ROWMAP", None); e["SENS_OVR"] = '{"_": 0}'
    subprocess.run([sys.executable, BUILD, p], env=e, check=True, stdout=subprocess.DEVNULL)
    ws = openpyxl.load_workbook(p)["In"]
    out = {}
    for r in range(4, ws.max_row + 1):
        k = ws.cell(r, 1).value
        if k and ws.cell(r, 2).value: out[k] = (ws.cell(r, 4).value, ws.cell(r, 5).value, ws.cell(r, 6).value)
    shutil.rmtree(wd, ignore_errors=True)
    return out
