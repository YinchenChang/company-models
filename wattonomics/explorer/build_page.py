"""把 explorer JSON 壓成頁面資料並替換 template.html 的 __DATA__。用法：python3 build_page.py <explorer.json> <out.html>"""
import json, sys, collections, pathlib
src, out = sys.argv[1], sys.argv[2]
d = json.load(open(src, encoding="utf-8")); R = d["records"]
f = lambda x: None if x in (None, "None", "") else round(float(x), 3)
o = {"src": d["source"]["file"], "sha": d["source"]["sha256"][:12]}
mk = collections.defaultdict(dict)
for r in R:
    if r["kind"] == "market_gw": mk[f"{r['path']}|{r['scen']}|{r['delay']}"][r["year"]] = f(r["value"])
o["mkt"] = mk
pg = collections.defaultdict(lambda: collections.defaultdict(dict)); names = {}
for r in R:
    if r["kind"] == "player_gw": pg[r["path"]][r["player"]][r["year"]] = f(r["value"]); names[r["player"]] = r["player_name"]
o["pgw"] = pg; o["pname"] = names
o["cov"] = {r["year"]: f(r["value"]) for r in R if r["kind"] == "coverage"}
o["campus"] = [{k: r[k] for k in ("key","name_","operator","user","mw","basis","online_year","online_month","known_delay_m","power_secured","tag","included")} for r in R if r["kind"] == "campus"]
for k in ("capex_calendar", "capex_ai", "capex_gw"):
    t = collections.defaultdict(dict)
    for r in R:
        if r["kind"] == k: t[r["player"]][r["year"]] = f(r["value"])
    o[k] = t
sp = collections.defaultdict(dict)
for r in R:
    if r["kind"] == "spend": sp[r["item"]][r["year"]] = f(r["value"])
o["spend"] = sp
o["fund"] = [{"player": r["player"], "name": r["player_name"], "year": r["year"], "period": r["period"], "values": r["values"], "gross": r.get("gross"), "po": r.get("payment_order")} for r in R if r["kind"] == "funding"]
tpl = (pathlib.Path(__file__).parent / "template.html").read_text(encoding="utf-8")
pathlib.Path(out).write_text(tpl.replace("__DATA__", json.dumps(o, ensure_ascii=False, default=str)), encoding="utf-8")
