"""AI 軟體地圖：由地位活頁簿（LibreOffice 重算後的值）＋ data/models_extract.json 產生 dist/index.html 與 dist/data.json。

網頁只呈現值、不做計算（排序、篩選除外）；數字全部由抽數檔與活頁簿帶入。
用法：
  python3 software_ai/map/build_position.py       # 產生地位活頁簿（有改地位輸入時）
  python3 software_ai/map/extract_models.py       # 抽各模型成品
  python3 software_ai/map/build_html.py           # 產生 dist/index.html、dist/data.json，並把各模型最新 HTML 複製到 dist/models/（本機預覽用，不進版控）
"""
import glob
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
SW = os.path.join(ROOT, "software_ai")
sys.path.insert(0, os.path.join(ROOT, "markets", "invest_map"))   # 公司檔案解析程式沿用 AI 算力地圖的 profiles_md.py（唯讀引用，不複製）
from profiles_md import load_all  # noqa: E402
BUILT = "2026-11-23"          # 頁首日期集中在此維護
VERSION = "v0.2"

RATING_BANDS = [   # 個股模型原則第 24 條（各模型 Val 頁同文）；這是規則常數，不是模型輸出
    {"name": "強力買進", "rule": "加權目標價上檔空間 > +50%"},
    {"name": "買進", "rule": "上檔空間 > +25%"},
    {"name": "中立", "rule": "上檔空間 ≥ −25%"},
    {"name": "放空", "rule": "上檔空間 ≥ −50%"},
    {"name": "強力放空", "rule": "其餘（上檔空間 < −50%）"},
    {"name": "收支不足", "rule": "收支關卡不通過（基準情境期間自由現金流扣股利合計為負），且估值評等為買進以上時改判；先看收支再看估值"},
]
DCF_NOTES = [
    "標準 DCF 把股票薪酬當作費用：營業利益已扣股票薪酬、評價股數含未歸屬 RSU 與選擇權淨股數。高股票薪酬公司的自由現金流因此遠低於公司口徑（non-GAAP 或「營運現金流減資本支出」加回股票薪酬）。",
    "Comparable 使用同業倍數：P/E 的分子盈餘多為 non-GAAP EPS、P/FCF 的現金流含股票薪酬加回、PS 與 EV／營收根本不反映股票薪酬。所以 DCF 與 Comparable 的落差，在高股票薪酬公司主要來自這個口徑差異加上倍數隱含的成長，不是公式錯誤。",
    "表中「DCF（股票薪酬加回）」是對照讀數：把股票薪酬加回現金流與永續價值，只作與標準 DCF 對照，不進加權目標價、不影響評等；加回讀數未計入加回後的稀釋，會高估。v0.2 以前的成品沒有此欄，顯示「—」。",
    "落差大時看各模型 Val 頁的落差分解（股票薪酬、時點、槓桿、倍數殘差）與上檔空間三段歸因（DCF 貢獻、未調整 Comparable 貢獻、相對倍數調整貢獻）。相對倍數調整來自歷史相對倍數（PS 與 PSG），不是對 DCF 的修正。",
]
DCF_SPEC = [
    "DCF 為標準 FCFF 版：明確期現金流以 WACC 折現（期中慣例），永續價值＝NOPAT ×（1 − g ÷ RONIC）÷（WACC − g）；RONIC 預設等於 WACC 加模型設定的利差；股權成本用 Blume 調整 β。",
    "目標價為股價基準日起 12 個月：DCF 的評價日價值以股權成本推到目標時點並扣期間股利；Comparable 用目標時點的 NTM 或 TTM 指標乘同業倍數。上檔空間是持有期間價差，不是年化報酬。",
    "加權目標價＝DCF 與 Comparable 依各模型設定的權重加權；兩腿各依三情境（A 順利、B 基準、C 不如預期）權重加權。權重與同業家數見各模型 Val 頁。",
    "同業少於 8 家不做迴歸，改用中位數；同業資料日期落後現價日者標「逾期」，使用前須更新。",
    "收支關卡先於估值：基準情境期間自由現金流（扣股利）合計為負即入不敷出，買進以上改判「收支不足」。各公司的關卡口徑（期間、是否取三情境最小值）見表中游標提示。",
]


def recalc(path, tmp):
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", tmp, path], check=True, capture_output=True)
    return os.path.join(tmp, os.path.basename(path))


def book_key(p):
    m = re.search(r"_v(\d+(?:\.\d+)*)\.xlsx$", p)
    return tuple(int(x) for x in m.group(1).split(".")) if m else ()


def rows(ws):
    hdr = [c.value for c in ws[1]]
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] is None or str(r[0]).strip() == "":
            continue
        yield dict(zip(hdr, r))


def num(x):
    return x if isinstance(x, (int, float)) and not isinstance(x, bool) else None


def short(name):
    return re.split(r"[（(]", name)[0].strip()


def main():
    books = sorted(glob.glob(os.path.join(HERE, "dist", "AI軟體地圖_產業地位_v*.xlsx")), key=book_key)
    if not books:
        raise SystemExit("找不到 dist/AI軟體地圖_產業地位_v*.xlsx；先跑 build_position.py")
    src = books[-1]
    tmp = tempfile.mkdtemp()
    try:
        wb = openpyxl.load_workbook(recalc(src, tmp), data_only=True)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    errs = [(ws.title, c.coordinate, c.value) for ws in wb for r in ws.iter_rows() for c in r
            if isinstance(c.value, str) and (c.value.startswith("#") or c.value.startswith("Err:"))]
    if errs:
        raise SystemExit(f"地位活頁簿有錯誤值：{errs[:5]}")

    ext = json.load(open(os.path.join(HERE, "data", "models_extract.json"), encoding="utf-8"))
    params = {r[0]: r[1] for r in wb["參數"].iter_rows(min_row=2, values_only=True) if r[0]}

    # ---- 次產業、個股 ----
    sub_rows = list(rows(wb["層級與次產業"]))
    sub_rows = [r for r in sub_rows if r.get("次產業")]               # 排除合計列
    pos_rows = list(rows(wb["個股地位"]))
    position = []
    for d in pos_rows:
        position.append({"tk": str(d["代號"]), "name": d["公司"], "short": short(d["公司"]), "market": d["市場"], "apac": d["亞太掛牌"],
                         "layer": d["層級"], "sub": d["次產業代碼"], "sub_name": d["次產業"], "imp": d["公司重要性（1–3）"],
                         "posn": d["市場位置（1–4）"], "basis": d["地位依據"], "tag": d["標記"], "ai": d["AI 方向（公司，判斷，不進分數）"],
                         "score": num(d["地位分數"]), "rank_all": num(d["全表排名"]), "rank_sub": num(d["次產業內排名"]),
                         "map_owner": d["地圖歸屬"]})
    position.sort(key=lambda p: (p["rank_all"] if p["rank_all"] is not None else 9999, p["name"]))

    ue = ext.get("unit_econ")
    ue_map = {}
    if ue:
        ue_map = {s["code"]: s for s in ue["subs"]}
    arch_name = {a["code"]: f'{a["code"]}｜{a["name"]}' for a in (ue["archetypes"] if ue else [])}

    subs = []
    for r in sub_rows:
        code = r["代碼"]
        comps = sorted((p for p in position if p["sub"] == code), key=lambda p: (p["rank_sub"] if p["rank_sub"] is not None else 99, p["name"]))
        subs.append({"code": code, "layer": r["層級"], "name": r["次產業"], "question": r["關鍵問題"], "pooltext": r["營收池（文字）"],
                     "pool": num(r["營收池中值（十億美元）"]), "irr": r["不可替代性（1–3）"], "vc": r["價值鏈分量分"], "ai": r["AI 方向（判斷，不進分數）"],
                     "pricing": r["主要定價模式"], "basis": r["依據與標記"], "n": len(comps), "n_apac": sum(1 for p in comps if str(p["apac"]).startswith("是")),
                     "n_sheet": num(r["覆蓋公司數"]), "n_apac_sheet": num(r["亞太標的數（是）"]),
                     "archetype": arch_name.get(ue_map[code]["archetype"]) if code in ue_map else None,
                     "companies": [{"tk": p["tk"], "name": p["name"], "short": p["short"], "map_owner": p["map_owner"], "score": p["score"]} for p in comps]})
    layers = []
    for s in subs:
        code = s["layer"].split()[0]
        if not layers or layers[-1]["code"] != code:
            layers.append({"code": code, "title": s["layer"][len(code):].strip(), "subs": []})
        layers[-1]["subs"].append(s["code"])

    # ---- 個股模型 ----
    pos_by_tk = {}
    for p in position:                                   # 同一公司多次產業時取排名最前者
        pos_by_tk.setdefault(p["tk"], p)
    models, order = {}, []
    os.makedirs(os.path.join(HERE, "dist", "models"), exist_ok=True)
    for f in glob.glob(os.path.join(HERE, "dist", "models", "*.html")):
        os.remove(f)
    for s in subs:
        for c in s["companies"]:
            if c["tk"] in ext["models"] and c["tk"] not in order:
                order.append(c["tk"])
    for tk, m in ext["models"].items():
        if tk not in pos_by_tk:
            raise SystemExit(f"個股模型 {tk} 不在地位活頁簿")
        p = pos_by_tk[tk]
        rm = m.get("relmult")
        rec = {"tk": tk, "score": p["score"], "rank_all": p["rank_all"], "folder": m["folder"], "name": p["short"], "sub": p["sub"], "sub_name": p["sub_name"], "layer": p["layer"],
               "status": m["status"], "ccy": rm["ccy"] if rm else None}
        if m["status"] == "ok":
            for k in ("price", "dcf", "comp", "tp", "upside", "rating", "rec", "scen", "sbc_dcf", "sbc_abc", "ai_net"):
                rec[k] = m.get(k)
            rec.update({"version": m["version"], "date": m["date"], "xlsx": os.path.basename(m["xlsx"]), "sbc_note": m.get("sbc_note"), "missing": m.get("missing", [])})
            g = dict(m["gate"])
            v = (g.get("verdict") or {}).get("value")
            b = (g.get("basis") or {}).get("value")
            if v:
                g["status"] = "通過" if v.startswith("通過") else "入不敷出"
                g["derived"] = False
            elif b is not None:
                g["status"] = "通過" if b >= 0 else "入不敷出"      # 模型沒有判定文字列：依基準值正負判定（標派生）
                g["derived"] = True
            else:
                g["status"] = None
            rec["gate"] = g
            h = m.get("html")
            if h:
                shutil.copy(os.path.join(ROOT, h["file"]), os.path.join(HERE, "dist", "models", m["folder"] + ".html"))
                rec["link"] = f"models/{m['folder']}.html"
                if not m.get("html_matches"):
                    rec["html_note"] = f"網頁版為 {h['version']}（{h['date'][:4]}-{h['date'][4:6]}-{h['date'][6:]}），xlsx 已為 {m['version']}；網頁版改版後同步"
            else:
                rec["link"] = None
                rec["html_note"] = "尚無 HTML 成品"
        if m["status"] == "ok":
            rec["tier"] = (rec.get("rec") or {}).get("value") or "未估值"      # 分級＝模型的投資建議（已含收支關卡）
        rec["relmult"] = ({k: rm[k] for k in ("file", "generated", "ps_reading", "psg_reading", "geo_base", "divergence", "divergence_gt_30pct", "valid_points", "gaps", "peers")} if rm else None)
        models[tk] = rec

    # ---- 單位經濟 ----
    years = ["2026", "2030", "2032"]
    ue_out = None
    if ue:
        tk_src = json.load(open(os.path.join(SW, "unit_economics", "data", "tk_snapshot.json"), encoding="utf-8"))["source"]
        ue_out = {"file_name": os.path.basename(ue["file"]), "version": ue["version"], "date": ue["date"], "years": years,
                  "tokenomics": f'{tk_src["file"]}（{tk_src["version"]}）', "recalc_errors": ue["recalc_errors"],
                  "archetypes": [{"code": a["code"], "name": a["name"], "scope": a["scope"], "title": a["title"],
                                  **{k: a[k] for k in ("netr", "gm", "infsh", "incrr", "erosr", "infr") if k in a}} for a in ue["archetypes"]],
                  "subs": [{"code": s["code"], "name": next((x["name"] for x in subs if x["code"] == s["code"]), s["label"]),
                            "archetype": arch_name.get(s["archetype"]), "label": s["label"],
                            **{k: s[k] for k in ("base", "net", "xfer", "netsh") if k in s}} for s in ue["subs"]]}
    rel0 = next((m["relmult"] for m in ext["models"].values() if m.get("relmult")), None)
    vp = (rel0 or {}).get("valid_points") or {}

    # ---- 限制與未找到 ----
    ok = [m for m in models.values() if m["status"] == "ok"]
    limits = [
        "營收池為第三方研調與公司揭露的區間中值（Analogy、Interested-party），單位經濟完成的次產業改讀模型；其餘次產業的營收池只作分量分級，不作估值輸入。",
        "單位經濟模型的參數除 Tokenomics 快照與少數外部證據外多為 Assumed，區間寬；樣板公司校準欄標「待研究包」；未來 token 單價路徑尚待登錄。",
        "公司名單、重要性與市場位置是判斷，市場份額未逐家查證；上市狀態有疑者在活頁簿標「待核」。已下市公司仍列為競爭參照。",
        "地位分數衡量今天的市場地位，不衡量 AI 讓它變好或變壞；AI 方向欄為判斷，不進分數。",
    ]
    for tk, m in models.items():
        if m["status"] != "ok":
            limits.append(f'{m["name"]}（{tk}）：{m["status"]}；估值與評等欄顯示「—」，待新版成品放入 dist/ 後重跑抽數。')
            continue
        names = {"sbc_dcf": "股票薪酬加回 DCF", "ai_net": "AI 淨效果（命題模組）", "scen": "情境 A／B／C 目標價", "tp": "加權目標價", "dcf": "DCF", "comp": "Comparable",
                 "upside": "上檔空間", "rating": "估值評等", "rec": "投資建議", "price": "現價", "gate": "收支關卡"}
        miss = [names.get(k, k) for k in m["missing"]]
        if m.get("sbc_dcf") is None and m.get("sbc_abc"):
            miss = [x for x in miss if x != "股票薪酬加回 DCF"] + ["股票薪酬加回 DCF（只有情境 A／B／C，無情境加權值，表中顯示情境 B）"]
        if m.get("sbc_note"):
            miss.append("股票薪酬加回 DCF：模型標「" + m["sbc_note"] + "」")
        if miss:
            limits.append(f'{m["name"]}（{tk}，{m["version"]}）未找到：' + "、".join(miss))
        if m.get("html_note"):
            limits.append(f'{m["name"]}（{tk}）：{m["html_note"]}')
        r = m.get("relmult")
        if r and r.get("gaps"):
            limits.append(f'{m["name"]}（{tk}）相對倍數庫註記：' + "；".join(r["gaps"]))
        if r is None:
            limits.append(f'{m["name"]}（{tk}）：相對倍數庫沒有此公司，基準欄顯示「—」')
    pdates = sorted({m["price"]["date"] for m in ok if m.get("price") and m["price"].get("date")})
    mdates = sorted({m["date"] for m in ok})
    sources = [
        {"k": "地位活頁簿", "v": os.path.basename(src) + "（LibreOffice 重算後讀值）；公司名單與判斷輸入見該檔「說明」「參數」頁"},
        {"k": "單位經濟", "v": f'{os.path.basename(ue["file"])}（LibreOffice 重算後讀值；錯誤值 {ue["recalc_errors"]} 格）' if ue else "未找到"},
        {"k": "推論成本", "v": "只讀 Tokenomics 快照，不自行估價：" + (ue_out["tokenomics"] if ue_out else "—")},
        {"k": "個股模型", "v": "；".join(f'{m["name"]} {m["xlsx"]}' for m in ok) or "—"},
        {"k": "相對倍數庫", "v": "software_ai/data/relmult_hist/out/<公司>.json（" + (rel0["generated"] if rel0 else "—") + " 產生）"},
        {"k": "抽數", "v": f'software_ai/map/extract_models.py（{ext["generated"]}）；每個數字在 data/models_extract.json 記檔名與儲存格'},
    ]
    n_models_ok = len(ok)
    # ---- 公司檔案（v0.2）：software_ai/map/profiles/<代號>.md ----
    profiles = {k: v for k, v in load_all(os.path.join(HERE, "profiles")).items() if k in models}
    for k, v in profiles.items():
        if v["missing"]:
            raise SystemExit(f"公司檔案 {k} 缺章節：{v['missing']}")
    no_pf = [m["name"] for tk, m in models.items() if m["status"] == "ok" and tk not in profiles]
    if no_pf:
        limits.append("公司檔案尚未完成：" + "、".join(no_pf))
    sources.append({"k": "公司檔案", "v": f"software_ai/map/profiles/<代號>.md，{len(profiles)} 家；格式規格 markets/invest_map/profiles/00_格式規格.md；工作單 software_ai/docs/workorders/20261123_AI軟體地圖_v0.2_版面對齊與公司檔案.md"})
    data = {
        "meta": {"version": VERSION, "built": BUILT, "position_book": os.path.basename(src),
                 "ue_file": os.path.basename(ue["file"]) if ue else "—", "ue_version": ue["version"] if ue else "—", "ue_date": ue["date"] if ue else "—",
                 "models_note": f"{n_models_ok}／{len(models)} 家有現行成品（版本 " + "、".join(sorted({m['version'] for m in ok})) + "）",
                 "price_dates": "、".join(pdates) or "—", "relmult_date": (rel0 or {}).get("generated", "—"), "extract_date": ext["generated"]},
        "kpis": [{"k": "次產業", "v": str(len(subs)), "s": f"{len(layers)} 層"},
                 {"k": "覆蓋個股（地位表）", "v": str(len(position)), "s": f'亞太掛牌 {sum(1 for p in position if str(p["apac"]).startswith("是"))} 家'},
                 {"k": "已建模", "v": f"{n_models_ok}／{len(models)}", "s": "個股模型現行成品"},
                 {"k": "現價日", "v": pdates[-1] if pdates else "—", "s": "各模型股價基準日"}],
        "layers": [dict(l, n_cov=sum(next(s for s in subs if s["code"] == c)["n"] for c in l["subs"]),
                        n_model=sum(1 for c in l["subs"] for x in next(s for s in subs if s["code"] == c)["companies"] if x["tk"] in models))
                   for l in layers],
        "subs": subs, "position": position, "models": models, "profiles": profiles, "model_order": order, "ue": ue_out, "params": params,
        "rating_bands": RATING_BANDS, "dcf_notes": DCF_NOTES, "dcf_spec": DCF_SPEC, "limits": limits, "sources": sources,
        "relmult_window": vp.get("window_quarters"), "relmult_min": vp.get("min_required"),
    }
    tpl = open(os.path.join(HERE, "template.html"), encoding="utf-8").read()
    html = tpl.replace("/*__DATA__*/null", json.dumps(data, ensure_ascii=False))
    os.makedirs(os.path.join(HERE, "dist"), exist_ok=True)
    with open(os.path.join(HERE, "dist", "index.html"), "w", encoding="utf-8") as f:
        f.write(html)
    with open(os.path.join(HERE, "dist", "data.json"), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    print(f"寫出 dist/index.html：{len(subs)} 次產業、{len(position)} 檔地位、{n_models_ok}／{len(models)} 個股模型有現行成品、{len(profiles)} 份公司檔案")


if __name__ == "__main__":
    main()
