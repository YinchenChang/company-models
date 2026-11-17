"""AI 軟體地圖核對（交付前必跑）。
(1) 每家評等、投資建議、加權目標價等於模型成品值（重讀記錄的儲存格，並確認引用的是 dist/ 最新成品）；
(2) 次產業與公司數等於地位活頁簿（LibreOffice 重算後）；單位經濟數字等於重算後成品；
(3) 網頁無外部網址（模型連結為站內相對路徑）、無錯誤值字樣；
(4) 不含個人情境字樣；頁面骨架（lang、charset、viewport、站際切換、不用 .tab）。
用法：python3 software_ai/map/check_map.py   （全部通過結束碼 0）
"""
import glob
import zipfile
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
sys.path.insert(0, HERE)
FAIL = []


def ok(cond, msg):
    print(("OK   " if cond else "FAIL ") + msg)
    if not cond:
        FAIL.append(msg)


def recalc(path, tmp):
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", tmp, path], check=True, capture_output=True)
    return openpyxl.load_workbook(os.path.join(tmp, os.path.basename(path)), data_only=True)


def cellval(path, ref):
    sh, c = ref.split("!")
    return openpyxl.load_workbook(path, data_only=True)[sh][c].value


def main():
    html = open(os.path.join(HERE, "dist", "index.html"), encoding="utf-8").read()
    data = json.load(open(os.path.join(HERE, "dist", "data.json"), encoding="utf-8"))
    ext = json.load(open(os.path.join(HERE, "data", "models_extract.json"), encoding="utf-8"))
    m = re.search(r"const D = (.*?);\nconst DASH", html, re.S)
    ok(m is not None and json.loads(m.group(1)) == data, "index.html 內嵌資料等於 data.json")

    # ---- (1) 模型值 ----
    import extract_models as E
    for tk, mm in data["models"].items():
        x = ext["models"][tk]
        latest = E.latest(x["folder"], "xlsx")
        if mm["status"] != "ok":
            ok(latest is None, f"{tk} 無現行成品且 dist/ 確實沒有（狀態：{mm['status']}）")
            continue
        ok(latest is not None and E.rel(latest[0]) == x["xlsx"], f"{tk} 引用的是 dist/ 最新成品 {x['xlsx'].split('/')[-1]}")
        path = os.path.join(ROOT, x["xlsx"])
        if not os.path.exists(path):
            ok(False, f"{tk} 抽數時的成品已不在（模型剛改版？重跑 extract_models.py 與 build_html.py）")
            continue
        for key, lab in (("tp", "目標價"), ("rec", "投資建議"), ("rating", "評等")):
            o = mm.get(key)
            if o is None:
                ok(False, f"{tk} {key} 缺值")
                continue
            ws = openpyxl.load_workbook(path, data_only=True)[o["cell"].split("!")[0]]
            c = o["cell"].split("!")[1]
            v = ws[c].value
            lbl = str(ws.cell(ws[c].row, 1).value)
            same = (abs(v - o["value"]) < 1e-9) if isinstance(v, (int, float)) else v == o["value"]
            ok(same and lab in lbl, f"{tk} {key}：頁面 {o['value'] if not isinstance(o['value'], float) else round(o['value'], 4)} = 模型 {o['cell']}（{lbl[:22]}）")
        for key in ("dcf", "comp", "upside", "price"):
            o = mm.get(key)
            if o:
                v = cellval(path, o["cell"])
                ok(isinstance(v, (int, float)) and abs(v - o["value"]) < 1e-9, f"{tk} {key} = {o['cell']}")
        if mm.get("rec") and mm.get("rating"):
            ok(mm["rec"]["value"] in ("強力買進", "買進", "中立", "放空", "強力放空", "收支不足"), f"{tk} 投資建議在六級內")

    # ---- (2) 地位活頁簿與單位經濟 ----
    # 需要 LibreOffice 重算；CI 執行環境若沒有 soffice，本段略過並明示（本機建置與發布腳本皆會跑完整版）
    if not shutil.which("soffice"):
        print("SKIP 地位活頁簿與單位經濟重算比對：找不到 soffice（CI 環境）；本機與發布前必跑")
    tmp = tempfile.mkdtemp()
    try:
      if shutil.which("soffice"):
          books = sorted(glob.glob(os.path.join(HERE, "dist", "AI軟體地圖_產業地位_v*.xlsx")))
          ok(os.path.basename(books[-1]) == data["meta"]["position_book"], "地圖使用最新地位活頁簿")
          wb = recalc(books[-1], tmp)
          subs = [r for r in wb["層級與次產業"].iter_rows(min_row=2, values_only=True) if r[0] and r[2] and r[1]]
          ok(len(subs) == len(data["subs"]), f"次產業數 {len(data['subs'])} = 活頁簿 {len(subs)}")
          ok([r[0] for r in subs] == [s["code"] for s in data["subs"]] and [r[2] for r in subs] == [s["name"] for s in data["subs"]], "次產業代碼與名稱順序一致")
          stocks = [r for r in wb["個股地位"].iter_rows(min_row=2, values_only=True) if r[0]]
          ok(len(stocks) == len(data["position"]), f"公司數 {len(data['position'])} = 活頁簿 {len(stocks)}")
          for r, s in zip(subs, data["subs"]):
              n = sum(1 for t in stocks if t[5] == r[0])
              ok(n == s["n"] == len(s["companies"]) == r[11], f"{r[0]} 覆蓋公司數 頁面 {s['n']} = 個股列 {n} = 活頁簿欄 {r[11]}")
          ok(sum(s["n"] for s in data["subs"]) == len(stocks), "各次產業公司數加總 = 個股列數")
          sc = {(t[0], t[5]): t[16] for t in stocks}
          ok(all(abs(sc[(p["tk"], p["sub"])] - p["score"]) < 1e-9 for p in data["position"]), "每列地位分數 = 活頁簿")
          if data["ue"]:
              ueb = recalc(os.path.join(ROOT, "software_ai", "unit_economics", "dist", data["ue"]["file_name"]), tmp)
              bad = 0
              for a in data["ue"]["archetypes"]:
                  for k in ("netr", "gm", "infsh"):
                      for y, o in a.get(k, {}).items():
                          sh, c = o["cell"].split("!")
                          if abs(ueb[sh][c].value - o["value"]) > 1e-9:
                              bad += 1
              for s in data["ue"]["subs"]:
                  for k in ("net", "xfer", "base", "netsh"):
                      for y, o in s.get(k, {}).items():
                          sh, c = o["cell"].split("!")
                          if abs(ueb[sh][c].value - o["value"]) > 1e-9:
                              bad += 1
              ok(bad == 0, "單位經濟頁面數字 = 重算後成品（不符 %d 格）" % bad)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    # ---- (3) 外部網址與錯誤值 ----
    ok(not re.search(r"https?:|//[a-z0-9.-]+\.[a-z]{2,}", re.sub(r"<!--.*?-->", "", html), re.I), "網頁無外部網址（http／https／協定相對網址）")
    ok(not re.search(r"<(script|link|img|iframe)[^>]+(src|href)=[\"']?(https?:|//)", html, re.I), "無外部載入資源")
    hrefs = re.findall(r'href="([^"]+)"', html) + re.findall(r"href=\\\"([^\\\"]+)\\\"", html)
    hrefs = [h for h in hrefs if not h.startswith("${")]          # 模板運算式（值由 data.json 的 link 帶入，另於下一項檢查）
    ok(all(h.startswith(("models/", "../")) for h in hrefs), f"連結皆為站內相對路徑（{len(hrefs)} 個 href 於模板，含站際切換）")
    ok(all(not (mm.get("link") and not mm["link"].startswith("models/")) for mm in data["models"].values()), "模型連結皆為 models/<公司>.html")
    ok(not re.search(r"#(DIV/0!|N/A|REF!|VALUE!|NAME\?|NUM!|NULL!)|Err:\d{3}|\bNaN\b|undefined", html), "網頁無 # 錯誤值、NaN、undefined 字樣")
    # ---- (4) 個人情境與骨架 ----
    banned = ["玉山", "基金", "部位", "持倉", "持股", "Andy", "董事長", "投資處", "先驗", "可投資", "同事", "個人"]
    hits = [b for b in banned if b in html]
    ok(not hits, f"不含個人情境字樣（{'、'.join(banned)}）" + (f"：命中 {hits}" if hits else ""))
    # 個股網頁複本與地圖活頁簿：只查明確的個人情境字樣（「先驗」「持股」在個股文中為方法／會計用語，不列）
    personal = ["玉山", "基金", "部位", "持倉", "Andy", "董事長", "投資處", "可投資", "同事"]
    mdir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dist", "models")
    mh = {}
    for fn in sorted(os.listdir(mdir)) if os.path.isdir(mdir) else []:
        if fn.endswith(".html"):
            t = open(os.path.join(mdir, fn), encoding="utf-8").read()
            h = [b for b in personal if b in t]
            if h: mh[fn] = h
    ok(not mh, "個股網頁複本不含個人情境字樣" + (f"：{mh}" if mh else ""))

    xh = {}
    for xp in glob.glob(os.path.join(os.path.dirname(os.path.abspath(__file__)), "dist", "*.xlsx")):
        with zipfile.ZipFile(xp) as z:
            t = "".join(z.read(n).decode("utf-8", "ignore") for n in z.namelist() if n.endswith(".xml"))
        h = [b for b in personal if b in t]
        if h: xh[os.path.basename(xp)] = h
    ok(not xh, "地圖活頁簿不含個人情境字樣" + (f"：{xh}" if xh else ""))
    ok('<html lang="zh-Hant">' in html and '<meta charset="utf-8">' in html and 'name="viewport"' in html, "html lang、charset、viewport")
    sw = re.search(r'<nav class="site-switch".*?</nav>', html, re.S)
    ok(sw is not None and 'href="../"' in sw.group(0) and 'href="../physical_ai/"' in sw.group(0) and "軟體 AI" in sw.group(0) and 'class="tab' not in sw.group(0), "站際切換 算力｜實體 AI｜軟體 AI，未用 .tab")
    print("\n通過" if not FAIL else f"\n未通過：{len(FAIL)} 項")
    sys.exit(1 if FAIL else 0)


if __name__ == "__main__":
    main()
