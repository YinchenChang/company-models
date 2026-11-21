# 把 relayout 後的 microsoft/html.json 轉成個股 HTML 版面規格 v1.3（v0.3.8；樣板 alphabet/gen_html_json.py 的 _v13()）。
# 用法：python3 microsoft/gen_html_v13.py <成品 xlsx> <rowmap.json> microsoft/html.json     （在 relayout.py 之後執行；列號來自 v038_post 回傳的 ROWMAP）
import json
import sys

import openpyxl

X, RM, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
rm = json.load(open(RM, encoding="utf-8"))
cfg = json.load(open(OUT, encoding="utf-8"))
assert cfg.get("layout") != "v1.3", "html.json 已是 v1.3（須先重跑建置與 relayout）"
wb = openpyxl.load_workbook(X, data_only=True)
V = rm["V38"]; PLA, RVA, CRD, RMA = V["PL"], V["RV"], V["CRD"], V["RM"]
VR = rm["VR"]


def find(sheet, prefix, start=1):
    ws = wb[sheet]
    for r in range(start, 40000):
        if str(ws.cell(r, 1).value or "").startswith(prefix):
            return r
    raise KeyError((sheet, prefix))


S = {x["id"]: x for x in cfg["sections"]}
PLW = wb["PL"]
assert str(wb["README"][f"A{RMA['ver']}"].value).startswith("v0.3.10")
cfg["layout"] = "v1.3"
old_note = cfg["hero"]["note"]
cfg["hero"]["note"] = [f"@README!A{RMA['hero']}"]
ph = PLA["hdr"]; ix = PLA["idx"]; sx = PLA["sidx"]; rid = PLA["ridx"]
# ---- 1. 命題 ----
th = S["thesis"]["blocks"]
i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RMA[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
th_chart = {"type": "theme", "title": "營收（十億美元，依分部堆疊）與營業利益率：FY2025A–FY2032E（情境 B）", "x": f"@PL!D{ph}:K{ph}",
            "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.3, "total": f"@PL!D{ix['rev']}:K{ix['rev']}",
            "bars": [{"name": [f"@PL!A{sx['s_ai']}"], "y": f"@PL!D{sx['s_ai']}:K{sx['s_ai']}", "color": "#0b7a80"},
                     {"name": [f"@PL!A{sx['s_dc']}"], "y": f"@PL!D{sx['s_dc']}:K{sx['s_dc']}", "color": "#c29a2b"}],
            "lines": [{"name": [f"@PL!A{ix['opm']}"], "y": f"@PL!D{ix['opm']}:K{ix['opm']}", "axis": "r", "color": "#3b3838", "labels": True}],
            "caption": "堆疊＝Agents and Infra（Azure、生產力與企業軟體等）與 Devices and Consumer 兩個分部；折線＝合併營業利益率（右軸，GAAP）。FY2027E 起為情境 B（基準）；AI 折舊上升使營業利益率逐年下降。"}
th[0:0] = [{"type": "note", "parts": [f"@README!A{RMA['rec']}"]}, th_chart]
# ---- 2. 估值：四個目標價與現價置頂 ----
va = S["valuation"]["blocks"]
i_bar = next(i for i, b in enumerate(va) if b["type"] == "bars")
va[i_bar] = {"type": "bars", "title": "四個目標價與現價（美元；豎線＝現價）", "fmt": "n2",
             "items": [{"label": f"@Val!A{VR[k]}", "value": f"@Val!B{VR[k]}", "series": sr} for k, sr in (("dcftp", 1), ("comp", 2), ("tp", 0), ("cons", 1))],
             "ref": {"label": f"@Val!A{VR['px']}", "value": f"@Val!B{VR['px']}"}}
bar = va.pop(i_bar)
top = [bar, {"type": "note", "parts": [S["valuation"].pop("intro")]}, {"type": "note", "parts": list(old_note)}]
S["valuation"]["blocks"] = top + va
# ---- 3. 敏感度：加權目標價相對基準龍捲風圖置頂 ----
SW = wb["Sens"]; t0 = find("Sens", "加權目標價相對基準（HTML"); r1_ = t0 + 2; r2_ = r1_
while SW.cell(r2_ + 1, 1).value:
    r2_ += 1
tors = []
for col, side in (("B", "低"), ("C", "高")):
    big = [r for r in range(r1_, r2_ + 1) if isinstance(SW[f"{col}{r}"].value, (int, float)) and abs(SW[f"{col}{r}"].value) >= 0.1]
    small = [r for r in range(r1_, r2_ + 1) if isinstance(SW[f"{col}{r}"].value, (int, float)) and 0.0005 <= abs(SW[f"{col}{r}"].value) < 0.1]
    base = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1"}
    assert big, (col, "無達兩位數的案例")
    tors.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值，變動達兩位數百分比的案例", rows=",".join(map(str, big))))
    if small:
        tors.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值", rows=",".join(map(str, small)), collapse=True,
                         summary=f"展開：加權目標價相對基準：輸入取{side}值，變動未達兩位數百分比的案例"))
tors.append({"type": "table", "caption": f"@Sens!A{t0}", "collapse": True, "summary": "展開：全部案例的加權目標價相對基準（低案、高案）",
             "head": [f"@Sens!{c}{t0 + 1}" for c in "ABC"], "rows": [[f"@Sens!A{r}", f"@Sens!B{r}|sp1", f"@Sens!C{r}|sp1"] for r in range(r1_, r2_ + 1)]})
sb = S["sens"]["blocks"]
for b in sb:
    if b["type"] == "tornado":
        b["collapse"] = True; b["summary"] = "展開：對照：加權目標價相對基準（原圖，標籤為模型輸入名稱）"
first2 = [t for t in tors if not t.get("collapse")]
S["sens"]["blocks"] = first2 + [{"type": "note", "parts": [S["sens"].pop("intro")]}] + [t for t in tors if t.get("collapse")] + sb
# ---- 4. 反向 DCF（新）----
A_ = RVA; yh = A_["yh"]
yr_fmt = {"y_rev": "n1", "y_ebit": "n1", "y_m": "p1", "y_df": "n3", "y_tvf": "n3", "y_k": "n3"}
rev_ = {"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2", "blocks": [
    {"type": "note", "parts": [f"@RevDCF!A{A_['sum_dcf']}"]}, {"type": "note", "parts": [f"@RevDCF!A{A_['sum_cmp']}"]},
    {"type": "table", "caption": "現金流折現法（情境 B 路徑）：現價隱含值與模型基準", "head": ["項目", "現價隱含", "模型基準"],
     "rows": [["預測期營業利益率（FY2028–FY2032 統一一個值；模型欄為等效值，終值年權重大）", f"@RevDCF!B{A_['m1']}|p1", f"@RevDCF!B{A_['m0']}|p1"],
              ["營收規模（FY2028 起各年同乘的倍數）", f"@RevDCF!B{A_['s1']}|n2", "—"],
              ["FY2028 營收（十億美元）", f"@RevDCF!B{A_['rv1']}|n0", f"@RevDCF!B{A_['rv0']}|n0"],
              ["FY2028 營收年增率", f"@RevDCF!B{A_['g1']}|p1", f"@RevDCF!B{A_['g0']}|p1"],
              ["折現率（WACC）", f"@RevDCF!B{A_['w1']}|p2", f"@RevDCF!B{A_['w0']}|p2"]]},
    {"type": "table", "caption": "分部加總法：現價隱含的倍數與模型", "head": ["項目", "現價隱含", "模型", "差距（現價隱含相對模型）"],
     "rows": [[f"@RevDCF!A{A_[k]}", f"@RevDCF!B{A_[k]}|{f}", f"@RevDCF!C{A_[k]}|{f}", f"@RevDCF!D{A_[k]}|sp1"] for k, f in (("sps", "n2"), ("sev", "n0"), ("smul", "n1"), ("scmp", "n2"))]},
    {"type": "table", "caption": "現金流折現法：現價隱含的企業價值與模型", "head": ["項目", "值"],
     "rows": [[f"@RevDCF!A{A_[k]}", f"@RevDCF!B{A_[k]}|{f}"] for k, f in (("px", "n2"), ("v0", "n2"), ("evr", "n0"), ("ev", "n0"), ("evx", "n2"), ("v_b", "n2"))]},
    {"type": "table", "caption": f"@RevDCF!A{yh}", "head": [f"@RevDCF!{c}{yh}" for c in "ABCDEFG"], "collapse": True,
     "summary": "展開：情境 B 各年數字與反解係數（封閉解的計算過程）",
     "rows": [[f"@RevDCF!A{A_[k]}"] + [f"@RevDCF!{c}{A_[k]}|{f}" for c in "BCDEFG"] for k, f in yr_fmt.items()]}]}
# ---- 5. 損益：標準損益表 ----
fmtu = lambda u: {"%": "p1", "US$": "n2", "十億股": "n2"}.get(u, "n1")
cols = "CDEFGHIJKL"
rowx = lambda r: [f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmtu(PLW.cell(r, 2).value)}" for c in cols]
head = [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in cols]
keys = ["rev", "yoy", "cogs", "gp", "gm", "opx", "oi", "opm", "oth", "pti", "tax", "ni", "nm", "sh", "eps"]
is_tbl = {"type": "table", "caption": "標準損益表（GAAP；十億美元，EPS 為美元）：FY2024A–FY2026A 實績、FY2027E–FY2032E 情境 B（基準）、最右欄為最新一季的單季實績", "head": head,
          "rows": [rowx(ix[k]) for k in keys] + [{"sub": f"@PL!A{PLA['ratio']}"}] + [rowx(rid["r_niy"])]}
chk_tbl = {"type": "table", "caption": "核對（應為零）", "collapse": True, "summary": "展開：標準損益表核對列", "head": head,
           "rows": [rowx(r) for r in range(PLA["chk"] + 1, PLA["chk"] + 7)]}
is_chart = {"type": "theme", "title": "營收（十億美元）與毛利率、營業利益率、淨利率：FY2024A–FY2032E（情境 B）", "x": f"@PL!C{ph}:K{ph}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
            "bars": [{"name": [f"@PL!A{ix['rev']}"], "y": f"@PL!C{ix['rev']}:K{ix['rev']}", "color": "#6fc7cb"}],
            "lines": [{"name": [f"@PL!A{ix['gm']}"], "y": f"@PL!C{ix['gm']}:K{ix['gm']}", "axis": "r", "color": "#0b7a80", "labels": True},
                      {"name": [f"@PL!A{ix['opm']}"], "y": f"@PL!C{ix['opm']}:K{ix['opm']}", "axis": "r", "color": "#c29a2b"},
                      {"name": [f"@PL!A{ix['nm']}"], "y": f"@PL!C{ix['nm']}:K{ix['nm']}", "axis": "r", "color": "#6f6bae", "dash": True}],
            "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸，GAAP）。FY2026A 淨利率含 OpenAI 等持股損益，預測期不估；毛利率逐年下降是 AI 折舊上升。"}
old = S["pl"]["blocks"]
for b in old:
    if b.get("type") == "pl":
        b["collapse"] = True; b["summary"] = "展開：三情境損益（A／B／C 並列，含折線）"
    elif b.get("type") == "table":
        b["collapse"] = True; b["summary"] = "展開：三情境損益頁的分部明細表"
pl_intro = S["pl"].pop("intro", None)
S["pl"].update(title="損益（標準損益表）", blocks=[{"type": "note", "parts": [f"@PL!A{PLA['note']}"]}, is_chart, is_tbl, {"type": "note", "parts": [f"@PL!A{PLA['note2']}"]}, chk_tbl]
               + ([{"type": "note", "parts": [pl_intro], "collapse": True, "summary": "展開：三情境損益頁說明"}] if pl_intro else []) + old)
# ---- 6. 驅動因子：分部營收推導 ----
seg_chart = {"type": "theme", "title": "分部營收堆疊（情境 B，十億美元）：Agents and Infra 與 Devices and Consumer", "x": f"@PL!D{ph}:K{ph}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.2,
             "total": f"@PL!D{ix['rev']}:K{ix['rev']}",
             "bars": [{"name": [f"@PL!A{sx['s_ai']}"], "y": f"@PL!D{sx['s_ai']}:K{sx['s_ai']}", "color": "#0b7a80"},
                      {"name": [f"@PL!A{sx['s_dc']}"], "y": f"@PL!D{sx['s_dc']}:K{sx['s_dc']}", "color": "#c29a2b"}],
             "caption": "堆疊＝兩個分部；頂端數字＝合併營收（FY2025A、FY2026A 為實績，FY2027E 起情境 B）；分部內各產品線的推導見下方。"}
seg_tbl = {"type": "table", "caption": "分部營收（十億美元；情境 B）：兩分部加總＝合併營收", "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in "DEFGHIJK"],
           "rows": [[f"@PL!A{sx[k]}"] + [f"@PL!{c}{sx[k]}|n1" for c in "DEFGHIJK"] for k in ("s_ai", "s_dc", "s_tot")]}
S["drivers"]["blocks"] = [{"type": "note", "parts": [f"@README!A{RMA['drv']}"]}, seg_chart, seg_tbl] + S["drivers"]["blocks"]
S["drivers"]["title"] = "驅動因子：分部營收推導"
# ---- 7. 對帳：三層協調 ----
CW = wb["Coord"]
cr = lambda p_: find("Coord", p_)
f1h, f1b = cr("F1. 本模型讀取"), cr("F1b. 追公式")
F1a, _ = CRD["F1"]; r_tr = CRD["rtr"]; wr = CRD["wrows"]
rd1, rd2, rj, rjr, rg, rl = CRD["rd1"], CRD["rd2"], CRD["rj"], CRD["rjr"], CRD["rg"], CRD["rl"]
rc9, rc3, rai, rcu, rmw = CRD["rc9"], CRD["rc3"], CRD["rai"], CRD["rcu"], CRD["rmw"]
ra = CRD["ra"]; rq = CRD["rq"]; rp = CRD["rp"]
capx = {"type": "table", "caption": "資本支出（公司已公布實績 vs Wattonomics 隱含值；US$B）", "head": ["讀數", "期間／口徑", "公司實績", "層級隱含", "差距", "狀態／判定"],
        "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!D{r}|n1", f"@Coord!E{r}|n1", f"@Coord!F{r}|sp1", f"@Coord!G{r}"] for r in (rd1, rd2, rc9, rc3, rg)]
                + [[f"@Coord!A{rj}", f"@Coord!B{rj}", "", "", "", f"@Coord!G{rj}"], [f"@Coord!A{rjr}", "", "", "", "", f"@Coord!G{rjr}"]]}
nocmp = {"type": "table", "caption": "沒有可比公司實績的項目（參考；不判定）", "head": ["項目", "期間／口徑", "公司實績", "層級隱含", "層級占公司實績", "狀態"],
         "rows": [[f"@Coord!A{rai}", f"@Coord!B{rai}", f"@Coord!D{rai}|n1", f"@Coord!E{rai}|n1", f"@Coord!F{rai}|p0", f"@Coord!G{rai}"],
                  [f"@Coord!A{rcu}", "", "", "", "", f"@Coord!G{rcu}"], [f"@Coord!A{rmw}", f"@Coord!B{rmw}", "", "", "", f"@Coord!G{rmw}"]]}
G = CRD["rgw"]
gw_t = {"type": "table", "caption": "新增容量（公司法說口頭揭露 vs Wattonomics 新增 IT GW；參考，口徑無法對齊）", "head": ["讀數", "期間／口徑", "公司", "層級隱含", "差距", "狀態／判定"],
        "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!D{r}|n2", f"@Coord!E{r}|n2", f"@Coord!F{r}|sp1", f"@Coord!G{r}"] for r in (G["d1"], G["d2"], G["b1"], G["b2"], G["y25"], G["b25"])]
                + [[f"@Coord!A{G['pg']}", f"@Coord!B{G['pg']}", f"@Coord!D{G['pg']}|n1", f"@Coord!E{G['pg']}|n1", f"@Coord!F{G['pg']}|sp1", f"@Coord!G{G['pg']}"],
                   [f"@Coord!A{G['j']}", f"@Coord!B{G['j']}", "", "", "", f"@Coord!G{G['j']}"], [f"@Coord!A{G['jr']}", "", "", "", "", f"@Coord!G{G['jr']}"]]}
gw_note = {"type": "note", "parts": [f"@Coord!H{G['j']}", "　", f"@Coord!A{CRD['rfb']}", "：", f"@Coord!B{CRD['rfb']}"]}
d3 = {"type": "table", "caption": "展開：實績對帳的輸入、口徑、出處與證據等級", "collapse": True, "summary": "展開：實績對帳的輸入、出處與證據等級",
      "head": ["項目", "期間／口徑", "公司實績口徑與出處", "公司實績等級", "層級口徑與出處", "層級等級"],
      "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!H{r}", f"@Coord!I{r}", f"@Coord!J{r}", f"@Coord!K{r}"] for r in list(rq) + list(rp) + [rd1, rd2, rg, rl] + list(ra) + [rai, rcu, rmw] + list(G["q"]) + [G["l"], G["d1"], G["d2"], G["fac"], G["ais"], G["pg"], G["y25"]]]}
trc = {"type": "table", "caption": f"@Coord!A{f1b}", "head": ["輸出名稱 ← 公式鏈", "同源判定", "讀到什麼"],
       "rows": [[f"@Coord!A{r}", f"@Coord!G{r}", f"@Coord!H{r}"] for r in range(f1b + 2, r_tr)]}
t1 = {"type": "table", "caption": f"@Coord!A{f1h}", "collapse": True, "summary": "展開：本模型讀取的 Tokenomics 名稱與 Wattonomics 對照名稱、值與版本",
      "head": ["名稱", "年度", "層級值", "本模型採用值", "單位", "層級與版本"],
      "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}|n3", f"@Coord!D{r}|n3", f"@Coord!E{r}", f"@Coord!F{r}"] for r in range(F1a + 1, wr["chip"] + 1)]}
d0, d1 = CRD["rdump"]
dump_t = {"type": "table", "caption": "展開：追公式傾印（腳本輸出：每格公式與值）", "collapse": True, "summary": "展開：追公式傾印（每格公式與值；腳本 trace_wattonomics.py 的輸出）",
          "head": ["儲存格（縮排＝依賴深度）", "值", "標記", "公式"],
          "rows": [[f"@Coord!A{r}", f"@Coord!D{r}|n3", f"@Coord!G{r}", f"@Coord!H{r}"] for r in range(d0, d1 + 1)]}
S["recon"]["blocks"].extend([
    {"type": "h3", "text": "三層協調：讀取層級與實績對帳"},
    {"type": "note", "parts": [f"@Coord!D{CRD['rs']}"]}, capx, nocmp, gw_t, gw_note, d3,
    {"type": "note", "parts": [f"@Coord!B{CRD['rr']}"]}, {"type": "note", "parts": [f"@Coord!B{CRD['rsame']}"]}, {"type": "note", "parts": [f"@Coord!B{CRD['rk']}"]},
    {"type": "h3", "text": "追公式：層級值是否讀公司實績或目標"}, trc, {"type": "note", "parts": [f"@Coord!A{r_tr}"]}, dump_t, t1])
# ---- 8. 資料狀態：本版要點 ----
S["status"]["blocks"][0:0] = [{"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RMA['ver']}"]}]
order = ["thesis", "valuation", "sens", "revdcf", "pl", "drivers", "recon", "status"]
S["revdcf"] = rev_
cfg["sections"] = [S[k] for k in order]
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", OUT)
