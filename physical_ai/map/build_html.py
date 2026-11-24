"""實體 AI 地圖 — 由資料活頁簿（LibreOffice 重算後的值）產生 dist/index.html（資料內嵌、可單檔開啟）與 dist/data.json。
網站只呈現活頁簿的值，不做計算（排序、篩選、繪圖除外）。
用法：python3 physical_ai/map/build_html.py
"""
import glob, json, os, re, shutil, subprocess, tempfile
import openpyxl
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from profiles_md import load_all

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
CSS = os.path.join(ROOT, "markets", "invest_map", "template.html")   # 版面樣式沿用 AI 算力地圖（同一套 CSS 變數與元件）
BUILT = "2026-11-23"
CLOSE_DATE = "2026-11-20"
ALIAS = {"TSLA-EV": "TSLA"}   # 同一公司多列：讀同一份模型
PKEY = {"TSLA-EV": "TSLA", "TSLA-RT": "TSLA", "1810-EV": "1810", "300124-M": "300124"}   # 同一公司多列：讀同一份公司檔案


# 個股模型 HTML（tools/company_html 產生）：本專案模型發布到 gh-pages physical_ai/models/<資料夾>/；跨圖公司連 AI 算力地圖的 models/
MODEL_DIR = {"TSLA": "tesla", "6324": "harmonic_drive", "6268": "nabtesco", "2049": "hiwin", "6954": "fanuc", "6506": "yaskawa",
             "WAYMO": "waymo", "TSLA-RT": "tesla_robotaxi",
             "MBLY": "mobileye", "9660": "horizon", "PONY": "pony", "9880": "ubtech", "UNITREE": "unitree", "2590": "geekplus", "SYM": "symbotic", "1316": "nexteer", "6481": "thk", "688017": "leaderdrive", "6273": "smc", "6861": "keyence",
             "6383": "daifuku", "003021": "zhaowei", "WRD": "weride", "6471": "nsk", "NOVT": "novanta", "AUTO": "autostore", "KGX": "kion", "300124": "inovance", "601100": "hengli", "RRX": "regal_rexnord", "603501": "omnivision", "2382.HK": "sunny", "603667": "wuzhou", "603728": "moons", "002050": "sanhua", "601689": "tuopu", "002747": "estun", "2498": "robosense", "HSAI": "hesai"}
CROSS = {"NVDA-C": "NVDA", "2330": "2330", "AMZN": "AMZN", "TER": "TER", "2308": "2308", "2317": "2317"}


def model_files():
    M = {}
    for tk, folder in MODEL_DIR.items():
        hs = sorted(glob.glob(os.path.join(ROOT, "physical_ai", folder, "dist", "*.html")))
        if not hs:
            continue
        h = os.path.basename(hs[-1]); m = re.match(r"(\d{4})(\d{2})(\d{2})_(.+)_(v[\d.]+)\.html$", h)
        x = h[:-5] + ".xlsx"
        if not os.path.exists(os.path.join(ROOT, "physical_ai", folder, "dist", x)):
            raise SystemExit(f"{folder}/dist/{h} 沒有同名 xlsx")
        M[tk] = dict(folder=folder, file=h, xlsx=x, title=m.group(4), date=f"{m.group(1)}-{m.group(2)}-{m.group(3)}", ver=m.group(5),
                     url=f"models/{folder}/{h}", local=f"../../{folder}/dist/{h}", home="實體 AI 地圖")
    inv = json.load(open(os.path.join(ROOT, "markets", "invest_map", "dist", "data.json"), encoding="utf-8"))
    for ptk, itk in CROSS.items():
        mm = (inv.get("models") or {}).get(itk)
        if mm:
            M[ptk] = dict(folder=mm["folder"], file=mm["file"], xlsx=mm["xlsx"], title=mm["title"], date=mm["date"], ver=mm["ver"],
                          url=f"../models/{mm['folder']}/{mm['file']}", local=f"../../../{mm['folder']}/dist/{mm['file']}", home="AI 算力地圖")
    return M


def recalc(path):
    tmp = tempfile.mkdtemp()
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", tmp, path], check=True, capture_output=True)
    return os.path.join(tmp, os.path.basename(path)), tmp


def table(ws):
    rows = list(ws.iter_rows(values_only=True))
    hdr = rows[0]
    return [dict(zip(hdr, r)) for r in rows[1:] if r[0] is not None]


def main():
    book = sorted(glob.glob(os.path.join(HERE, "dist", "*_實體AI地圖_資料_*.xlsx")))[-1]
    calc, tmp = recalc(book)
    wb = openpyxl.load_workbook(calc, data_only=True)
    shutil.rmtree(tmp, ignore_errors=True)
    errs = [(ws.title, v) for ws in wb for r in ws.iter_rows(values_only=True) for v in r
            if isinstance(v, str) and (v.startswith("#") or v.startswith("Err:"))]
    if errs:
        raise SystemExit(f"活頁簿有錯誤值：{errs[:5]}")
    ver = re.search(r"_(v[\d.]+)\.xlsx$", book).group(1)
    lines = {}
    for r in table(wb["線"]):
        L = lines.setdefault(r["線"], dict(code=r["線"], name=r["線名"], q=r["關鍵問題"], metrics=[]))
        L["metrics"].append(dict(lab=r["指標"], lo=r["低"], base=r["基準"], hi=r["高"], unit=r["單位"], fmt=r["格式"], src=r["出處"]))
    segs = [dict(code=r["代碼"], layer=r["層級"], name=r["次產業"], q=r["關鍵問題"], share=r["占比（文字）"], irr=r["不可替代性（1–3）"],
                 vc=r["價值鏈分量分"], basis=r["依據與標記"], line=r["所屬線"]) for r in table(wb["次產業"])]
    vals = {}
    for r in table(wb["估值"]):
        tw = r["台股不公開目標價"] == "是"
        vals[r["代號"]] = dict(map=r["主圖"], ccy=r["幣別"], tp=None if tw else r["目標價"], tp_src=None if tw else r["目標價出處"],
                              close=r["收盤"], close_date=str(r["收盤日"])[:10], close_src=r["收盤出處"], up=r["上檔空間"],
                              vtier=r["估值評等"], gate=r["收支關卡（情境 B 合計）"], gate_src=r["收支出處"], tier=r["最終分級"],
                              mpx=r["模型現價"], mpd=r["模型現價日"], mfile=r["模型檔"], tw=tw,
                              up_src=f"資料活頁簿 {os.path.basename(book)}｜估值頁：目標價 ÷ 收盤 − 1（目標價出處：{r['目標價出處']}）")
    segv = {r["代號"]: dict(lo=r["低（情境 C）"], base=r["基準（加權）"], hi=r["高（情境 A）"], unit=r["單位"], note=r["說明"],
                           src=f"{r['出處低']}；{r['出處基準']}；{r['出處高']}", mfile=r["模型檔"]) for r in table(wb["分部"])}
    pos = []
    for r in table(wb["個股地位"]):
        tk = r["代號"]; k = ALIAS.get(tk, tk)
        v = vals.get(k); s = segv.get(tk)
        cross = "跨圖" in (r["地圖歸屬"] or "")
        tier = v["tier"] if v else ("未上市" if (s or str(r["市場"]).startswith("未上市")) else "未估值")
        pos.append(dict(tk=tk, name=r["公司"], mkt=r["市場"], asia=r["亞太掛牌"] == "是", layer=r["層級"], seg=r["次產業代碼"], segname=r["次產業"],
                        line=r["所屬線"], sub=r["子分類"], imp=r["公司重要性"], posn=r["市場位置"], irr_s=r["不可替代性分"], pos_s=r["市場位置分"],
                        vc_s=r["價值鏈分量分"], score=r["地位分數"], rank=r["排名"], mapnote=r["地圖歸屬"], cross=cross,
                        basis=r["地位依據"], tag=r["標記"], exp=r["實體 AI 營收曝險"], tier=tier, val_key=k if v else None, seg_key=tk if s else None))
    groups = [dict(code=r["組別"], name=r["組名"], q=r["組關鍵問題"], line=str(r["線"]), subs=str(r["次產業代碼"]).split()) for r in table(wb["地圖分組"])]
    g_of = {sc: g["code"] for g in groups for sc in g["subs"]}
    for p in pos:
        p["grp"] = g_of[p["seg"]]
    profiles = load_all(os.path.join(HERE, "profiles"))
    need = {PKEY.get(p["tk"], p["tk"]) for p in pos if not p["cross"]}
    extra = sorted(set(profiles) - need)
    if extra:
        raise SystemExit(f"公司檔案多出不在地圖的代號：{extra}")
    miss_sec = {k: v["missing"] for k, v in profiles.items() if v["missing"]}
    if miss_sec:
        raise SystemExit(f"公司檔案缺章節：{miss_sec}")
    watch = table(wb["觀察名單"])
    tiers = {}
    seen = set()
    for p in pos:
        key = p["val_key"] or p["seg_key"] or p["tk"]
        if key in seen:
            continue
        seen.add(key); tiers[p["tier"]] = tiers.get(p["tier"], 0) + 1
    data = dict(meta=dict(version=ver, book=os.path.basename(book), built=BUILT, close_date=CLOSE_DATE,
                          price_note=f"上市公司一律用 {CLOSE_DATE} 收盤（Yahoo Finance 日線，data/prices.csv）；上檔空間＝最新一份模型目標價 ÷ 最新收盤 − 1。台股只顯示上檔空間與評等，不顯示目標價。",
                          tiers=tiers, n_pos=len(pos), n_cross=sum(p["cross"] for p in pos), n_prof=len(profiles), n_prof_need=len(need)),
                lines=list(lines.values()), models=model_files(), groups=groups, pkey=PKEY, profiles=profiles, segments=segs, positions=pos, valuation=vals, segvals=segv, watch=watch)
    js = json.dumps(data, ensure_ascii=False, default=str)
    json.dump(data, open(os.path.join(HERE, "dist", "data.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
    raw = open(CSS, encoding="utf-8").read()
    css = re.sub(r"/\*.*?\*/", "", raw[raw.index("<style>"):raw.index("</style>")], flags=re.S)  # 只取樣式（去註解：原檔註解含規格來源）
    tpl = open(os.path.join(HERE, "template.html"), encoding="utf-8").read()
    html = tpl.replace("/*CSS*/", css).replace("/*DATA*/", js.replace("</", "<\\/"))
    open(os.path.join(HERE, "dist", "index.html"), "w", encoding="utf-8").write(html)
    print("index.html", len(html), "tiers", tiers, "跨圖", data["meta"]["n_cross"], "公司檔案", len(profiles), "/", len(need))


if __name__ == "__main__":
    main()
