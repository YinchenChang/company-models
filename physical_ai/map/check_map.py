"""實體 AI 地圖核對：用法 python3 physical_ai/map/check_map.py ；失敗＝結束碼 1。
1. data.json 線指標、個股目標價、收支關卡、分部價值都能回溯到出處儲存格（重新開檔讀值比對）；收盤＝data/prices.csv。
2. 評等只用五級（另收支不足、未估值、未上市）；上檔空間＝目標價 ÷ 收盤 − 1、評等依門檻重算一致。
3. 地位分數與抄錄 csv 一致。
4. 跨圖列都有主圖模型目標價（AI 算力地圖 data.json）或為分部估值。
5. 公開 repo 無個人或工作情境字樣（data、dist、README、程式）。
"""
import csv, glob, json, os, re, sys
import openpyxl
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
D = json.load(open(os.path.join(HERE, "dist", "data.json"), encoding="utf-8"))
fail, cache = [], {}

def cell(path, ref):
    wb = cache.setdefault(path, openpyxl.load_workbook(os.path.join(ROOT, path), data_only=True))
    sh, a = ref.split("!")
    return wb[sh][a.split(":")[0]].value

def close(a, b, tol=1e-6):
    return isinstance(a, (int, float)) and isinstance(b, (int, float)) and abs(a - b) <= tol * max(1, abs(b))

# 1a 線指標
n = 0
for L in D["lines"]:
    for m in L["metrics"]:
        f, ref = m["src"].split("｜")
        sh, rng = ref.split("!")
        cols = re.findall(r"([A-Z]+)(\d+)", rng)
        r = cols[0][1]
        want = [cell(f, f"{sh}!{c}{r}") for c, _ in cols]
        if len(cols) == 1:
            ok = close(m["base"], want[0])
        else:
            letters = [c for c, _ in cols]
            full = [openpyxl.utils.get_column_letter(i) for i in range(openpyxl.utils.column_index_from_string(letters[0]), openpyxl.utils.column_index_from_string(letters[-1]) + 1)]
            vals = [cell(f, f"{sh}!{c}{r}") for c in full]
            ok = all((x is None and y in (None, "")) or close(x, y) for x, y in zip((m["lo"], m["base"], m["hi"]), vals))
        n += 1
        if not ok:
            fail.append(f"線指標 {m['lab']} 與 {m['src']} 不符")
# 1b 個股
prices = {r["tk"]: r for r in csv.DictReader(open(os.path.join(HERE, "data", "prices.csv"), encoding="utf-8"))}
inv = json.load(open(os.path.join(ROOT, "markets/invest_map/dist/data.json"), encoding="utf-8"))
ib = {v["tk"]: v for v in inv["valuation"]}
CROSS = {"NVDA-C": "NVDA", "2330": "2330", "AMZN": "AMZN", "TER": "TER", "2308": "2308", "2317": "2317"}
TIERS = {"強力買進", "買進", "中立", "放空", "強力放空"}
rec5 = lambda u: "強力買進" if u > 0.5 else "買進" if u > 0.25 else "中立" if u >= -0.25 else "放空" if u >= -0.5 else "強力放空"
book = openpyxl.load_workbook(sorted(glob.glob(os.path.join(HERE, "dist", "*_實體AI地圖_資料_*.xlsx")))[-1], data_only=False)
Vb = {r[0].value: [c.value for c in r] for r in book["估值"].iter_rows(min_row=2) if r[0].value}
for tk, v in D["valuation"].items():
    tp = Vb[tk][4]
    if tk in CROSS:
        ok = close(tp, ib[CROSS[tk]]["center"]); pk = CROSS[tk]
    else:
        f, ref = Vb[tk][5].split("｜"); ok = close(tp, cell(f, ref)); pk = tk
        f, ref = v["gate_src"].split("｜")
        if not close(v["gate"], cell(f, ref)):
            fail.append(f"{tk} 收支關卡與 {v['gate_src']} 不符")
    if not ok:
        fail.append(f"{tk} 目標價不符出處")
    if not close(v["close"], float(prices[pk]["close"])) or v["close_date"] != prices[pk]["date"]:
        fail.append(f"{tk} 收盤與 prices.csv 不符")
    if not close(v["up"], tp / v["close"] - 1, 1e-9):
        fail.append(f"{tk} 上檔空間不等於目標價 ÷ 收盤 − 1")
    if v["vtier"] not in TIERS or v["vtier"] != rec5(v["up"]):
        fail.append(f"{tk} 估值評等 {v['vtier']} 與五級規則不符")
    exp = "收支不足" if (v["gate"] is not None and v["gate"] < 0 and v["vtier"] in ("強力買進", "買進")) else v["vtier"]
    if v["tier"] != exp:
        fail.append(f"{tk} 最終分級 {v['tier']} ≠ {exp}")
    if v["tw"] and v["tp"] is not None:
        fail.append(f"{tk} 台股不得顯示目標價")
for tk, s in D["segvals"].items():
    for val, src in zip((s["lo"], s["base"], s["hi"]), s["src"].split("；")):
        f, ref = src.split("｜")
        if not close(val, cell(f, ref)):
            fail.append(f"{tk} 分部價值不符 {src}")
# 2/3 地位與評等集合
pos_csv = {r["代號"]: r for r in csv.DictReader(open(os.path.join(HERE, "data", "positions.csv"), encoding="utf-8"))}
ALLOWED = TIERS | {"收支不足", "未估值", "未上市"}
for p in D["positions"]:
    if p["tier"] not in ALLOWED:
        fail.append(f"{p['tk']} 分級 {p['tier']} 不在允許集合")
    if not close(p["score"], float(pos_csv[p["tk"]]["地位分數"]), 1e-9):
        fail.append(f"{p['tk']} 地位分數與抄錄不符")
    if p["cross"] and not (p["val_key"] or p["seg_key"]):
        fail.append(f"跨圖列 {p['tk']} 沒有主圖模型")
    if p["cross"] and p["val_key"] and "算力" not in D["valuation"][p["val_key"]]["map"]:
        fail.append(f"跨圖列 {p['tk']} 主圖不是 AI 算力地圖")
# 4 無個人字樣
# 通用字樣寫在這裡；人名、雇主等私人字樣不寫進公開程式碼，改由環境變數 PRIVATE_TERMS（逗號分隔）提供
_terms = ["部位", "下注"] + [t for t in os.environ.get("PRIVATE_TERMS", "").split(",") if t]
PAT = re.compile("|".join(re.escape(t) for t in _terms))
for f in glob.glob(os.path.join(HERE, "**", "*"), recursive=True):
    if os.path.isfile(f) and f.endswith((".csv", ".json", ".html", ".md", ".py")) and not f.endswith("check_map.py"):
        t = open(f, encoding="utf-8", errors="ignore").read()
        m = PAT.search(t)
        if m:
            fail.append(f"個人字樣「{m.group(0)}」：{os.path.relpath(f, ROOT)}")
for f in glob.glob(os.path.join(HERE, "dist", "*.xlsx")):
    wb = openpyxl.load_workbook(f)
    for ws in wb:
        for row in ws.iter_rows(values_only=True):
            for x in row:
                if isinstance(x, str) and PAT.search(x):
                    fail.append(f"個人字樣：{os.path.basename(f)} {ws.title}：{x[:40]}"); break
print(f"線指標 {n} 項、估值 {len(D['valuation'])} 家、分部 {len(D['segvals'])}、個股列 {len(D['positions'])}、跨圖 {sum(p['cross'] for p in D['positions'])}")
print("核對：" + ("全部通過" if not fail else "失敗\n- " + "\n- ".join(fail)))
sys.exit(1 if fail else 0)
