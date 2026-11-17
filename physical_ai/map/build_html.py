"""實體 AI 地圖 — 由資料活頁簿（LibreOffice 重算後的值）產生 dist/index.html（資料內嵌、可單檔開啟）與 dist/data.json。
網站只呈現活頁簿的值，不做計算（排序、篩選、繪圖除外）。
用法：python3 physical_ai/map/build_html.py
"""
import glob, json, os, re, shutil, subprocess, tempfile
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
CSS = os.path.join(ROOT, "markets", "invest_map", "template.html")   # 版面樣式沿用 AI 算力地圖（同一套 CSS 變數與元件）
BUILT = "2026-11-17"
CLOSE_DATE = "2026-11-16"
ALIAS = {"TSLA-EV": "TSLA"}   # 同一公司多列：讀同一份模型


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
                          tiers=tiers, n_pos=len(pos), n_cross=sum(p["cross"] for p in pos)),
                lines=list(lines.values()), segments=segs, positions=pos, valuation=vals, segvals=segv, watch=watch)
    js = json.dumps(data, ensure_ascii=False, default=str)
    json.dump(data, open(os.path.join(HERE, "dist", "data.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
    css = re.sub(r"/\*.*?\*/", "", "\n".join(open(CSS, encoding="utf-8").read().split("\n")[6:255]), flags=re.S)  # 去註解（原檔註解含規格來源）
    tpl = open(os.path.join(HERE, "template.html"), encoding="utf-8").read()
    html = tpl.replace("/*CSS*/", css).replace("/*DATA*/", js.replace("</", "<\\/"))
    open(os.path.join(HERE, "dist", "index.html"), "w", encoding="utf-8").write(html)
    print("index.html", len(html), "tiers", tiers, "跨圖", data["meta"]["n_cross"])


if __name__ == "__main__":
    main()
