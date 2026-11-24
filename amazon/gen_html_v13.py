# 把 relayout 後的 amazon/html.json 轉成個股 HTML 版面規格 v1.3（v0.3.7；樣板 microsoft/gen_html_v13.py、alphabet/gen_html_json.py 的 _v13()）。
# 用法：python3 amazon/gen_html_v13.py <成品 xlsx> <rowmap.json> amazon/html.json     （在 relayout.py 之後執行；列號來自 v037_post 回傳、寫在 ROWMAP 的 V37）
# 數字與文字全取自 xlsx；設定文字除年份與日期外不含數字。
import json
import sys

import openpyxl

X, RM, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
rm = json.load(open(RM, encoding="utf-8"))
cfg = json.load(open(OUT, encoding="utf-8"))
assert cfg.get("layout") != "v1.3", "html.json 已是 v1.3（須先重跑建置與 relayout）"
wb = openpyxl.load_workbook(X, data_only=True)
V = rm["V37"]; PLA, RVA, CRD, RMA = V["PL"], V["RV"], V["CRD"], V["RM"]
VR = rm["VR"]


def find(sheet, prefix, start=1):
    ws = wb[sheet]
    for r in range(start, ws.max_row + 1):
        if str(ws.cell(r, 1).value or "").startswith(prefix):
            return r
    raise KeyError((sheet, prefix))


S = {x["id"]: x for x in cfg["sections"]}
PLW = wb["PL"]
assert str(wb["README"][f"A{RMA['ver']}"].value).startswith("v0.3.10")
cfg["layout"] = "v1.3"
old_note = cfg["hero"]["note"]
cfg["hero"]["note"] = [f"@README!A{RMA['hero']}"]
ph = PLA["hdr"]; ix = PLA["idx"]; sx = PLA["sidx"]
# ---- 1. 命題：核心問題對照圖、投資建議前提、重點觀察 5 點 ----
th = S["thesis"]["blocks"]
i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RMA[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
grp = lambda nm, r, fmt, ref=None: {"name": [nm], "fmt": fmt, "items": ([ref] if ref else []) + [{"name": [f"模型 情境 {s}"], "v": f"@Val!{c}{r}"} for s, c in (("A", "B"), ("B", "C"), ("C", "D"))]}
th_chart = {"type": "compare", "title": "回答核心問題：FY2028 對外 AI 報酬率對資金成本、AWS 營業利益率與自由現金流扣股利（三情境）",
            "groups": [grp("FY2028 對外 AI 報酬率", VR["roic28"], "p1", {"name": ["WACC（資金成本）"], "v": f"@Val!B{VR['wacc']}", "ref": True}),
                       grp("FY2028 AWS 營業利益率", VR["awsm28"], "p1"), grp("FY2028 自由現金流扣股利（十億美元）", VR["fcf28"], "n1")],
            "caption": "灰色＝WACC；青綠＝本模型三情境。報酬率低於 WACC 表示 AI 投資在該情境下不創造價值；自由現金流扣股利轉正表示零售與廣告的現金流撐得住資本支出。"}
th[0:0] = [th_chart, {"type": "note", "parts": [f"@README!A{RMA['rec']}"]}]
# ---- 2. 估值：四個目標價與現價置頂 ----
va = S["valuation"]["blocks"]
i_bar = next(i for i, b in enumerate(va) if b["type"] == "bars")
bar = {"type": "bars", "title": "四個目標價與現價（美元；豎線＝現價）", "fmt": "n2",
       "items": [{"label": f"@Val!A{VR[k]}", "value": f"@Val!B{VR[k]}", "series": sr} for k, sr in (("dcftp", 1), ("comptp", 2), ("tpw", 0), ("cons", 1))],
       "ref": {"label": f"@Val!A{VR['px']}", "value": f"@Val!B{VR['px']}"}}
va.pop(i_bar)
tbl4 = {"type": "table", "caption": "四個目標價並列（美元；相對現價）", "head": ["目標價", "美元", "相對現價"],
        "rows": [[f"@Val!A{VR[k]}", f"@Val!B{VR[k]}|n2", u] for k, u in (("dcftp", ""), ("comptp", f"@Val!B{VR['upc']}|sp1"), ("tpw", f"@Val!B{VR['upw']}|sp1"), ("cons", f"@Val!B{VR['upcons']}|sp1"))]
        + [[f"@Val!A{VR['px']}", f"@Val!B{VR['px']}|n2", ""]]}
top = [bar, tbl4, {"type": "note", "parts": [S["valuation"].pop("intro")]}, {"type": "note", "parts": list(old_note)}]
S["valuation"]["blocks"] = top + va
# ---- 3. 敏感度：加權目標價相對基準龍捲風圖置頂 ----
SW = wb["Sens"]; t0 = V["tor"]; r1_ = t0 + 2; r2_ = r1_
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
# ---- 4. 反向 DCF ----
A_ = RVA; yh = A_["yh"]
yr_fmt = {"y_rev": "n1", "y_ebit": "n1", "y_m": "p1", "y_df": "n3", "y_fcf": "n1", "y_tvf": "n3", "y_k": "n3"}
rev_ = {"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2", "blocks": [
    {"type": "note", "parts": [f"@RevDCF!A{A_['sum_dcf']}"]}, {"type": "note", "parts": [f"@RevDCF!A{A_['sum_cmp']}"]},
    {"type": "table", "caption": "現金流折現法（情境 B 路徑）：現價隱含值與模型基準", "head": ["項目", "現價隱含", "模型基準"],
     "rows": [["預測期營業利益率（FY2028–FY2032 統一一個值；模型欄為等效值，終值年權重大）", f"@RevDCF!B{A_['m1']}|p1", f"@RevDCF!B{A_['m0']}|p1"],
              ["營收規模（FY2028 起各年同乘的倍數）", f"@RevDCF!B{A_['s1']}|n2", "—"],
              ["FY2028 營收（十億美元）", f"@RevDCF!B{A_['rv1']}|n0", f"@RevDCF!B{A_['rv0']}|n0"],
              ["FY2028 營收年增率", f"@RevDCF!B{A_['g1']}|p1", f"@RevDCF!B{A_['g0']}|p1"],
              ["折現率（WACC）", f"@RevDCF!B{A_['w1']}|p2", f"@RevDCF!B{A_['w0']}|p2"],
              ["隱含 WACC 判讀", f"@RevDCF!B{A_['w_ok']}", ""]]},
    {"type": "table", "caption": "分部加總法：現價隱含的倍數與模型", "head": ["項目", "現價隱含", "模型", "差距（現價隱含相對模型）"],
     "rows": [[f"@RevDCF!A{A_[k]}", f"@RevDCF!B{A_[k]}|{f}", f"@RevDCF!C{A_[k]}|{f}", f"@RevDCF!D{A_[k]}|sp1"] for k, f in (("sps", "n2"), ("sev", "n0"), ("smul", "n1"), ("scmp", "n2"))]},
    {"type": "table", "caption": "現金流折現法：現價隱含的企業價值與模型", "head": ["項目", "值"],
     "rows": [[f"@RevDCF!A{A_[k]}", f"@RevDCF!B{A_[k]}|{f}"] for k, f in (("px", "n2"), ("v0", "n2"), ("evr", "n0"), ("ev", "n0"), ("evx", "n2"), ("v_b", "n2"))]},
    {"type": "table", "caption": f"@RevDCF!A{yh}", "head": [f"@RevDCF!{c}{yh}" for c in "ABCDEFG"], "collapse": True,
     "summary": "展開：情境 B 各年數字與反解係數（封閉解的計算過程）",
     "rows": [[f"@RevDCF!A{A_[k]}"] + [f"@RevDCF!{c}{A_[k]}|{f}" for c in "BCDEFG"] for k, f in yr_fmt.items()]}]}
# ---- 5. 損益：標準損益表 ----
fmtu = lambda u: {"%": "p1", "US$": "n2", "十億股": "n2"}.get(u, "n1")
cols = "CDEFGHIJKLM"
cellx = lambda r, c: f"@PL!{c}{r}|{fmtu(PLW.cell(r, 2).value)}" if isinstance(PLW[f"{c}{r}"].value, (int, float)) else f"@PL!{c}{r}"
rowx = lambda r: [f"@PL!A{r}"] + [cellx(r, c) for c in cols]
head = [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in cols]
keys = ["rev", "yoy", "cogs", "gp", "gm", "opx", "oi", "opm", "oth", "pti", "tax", "ni", "nm", "sh", "eps"]
is_tbl = {"type": "table", "caption": "標準損益表（GAAP；十億美元，EPS 為美元）：FY2023A–FY2025A 實績、FY2026E（九個月實績加第四季模型）–FY2032E 情境 B（基準）、最右欄為今年九個月實績", "head": head,
          "rows": [rowx(ix[k]) for k in keys] + [{"sub": f"@PL!A{PLA['ratio']}"}] + [rowx(ix[k]) for k in ("r_yoy", "r_opm", "r_nm", "r_niy")]}
chk_tbl = {"type": "table", "caption": "核對與差額", "collapse": True, "summary": "展開：標準損益表核對列", "head": head,
           "rows": [rowx(ix[k]) for k in ("c1", "c2", "c3")]}
is_chart = {"type": "theme", "title": "營收（十億美元）與營業利益率、淨利率：FY2023A–FY2032E（情境 B）", "x": f"@PL!C{ph}:L{ph}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
            "bars": [{"name": [f"@PL!A{ix['rev']}"], "y": f"@PL!C{ix['rev']}:L{ix['rev']}", "color": "#6fc7cb"}],
            "lines": [{"name": [f"@PL!A{ix['opm']}"], "y": f"@PL!C{ix['opm']}:L{ix['opm']}", "axis": "r", "color": "#c29a2b", "labels": True},
                      {"name": [f"@PL!A{ix['nm']}"], "y": f"@PL!C{ix['nm']}:L{ix['nm']}", "axis": "r", "color": "#6f6bae", "dash": True}],
            "caption": "長條＝營收；兩條線＝營業利益率、淨利率（右軸，GAAP）。FY2026E 淨利率含持股評價利益，預測期不估；營業成本與毛利模型未分列。"}
old = S["pl"]["blocks"]
for b in old:
    if b.get("type") == "pl":
        b["collapse"] = True; b["summary"] = "展開：三情境損益（A／B／C 並列，含折線）"
    elif b.get("type") == "table":
        b["collapse"] = True; b["summary"] = "展開：三情境損益頁的分部核對"
pl_intro = S["pl"].pop("intro", None)
S["pl"].update(title="損益（標準損益表）", blocks=[{"type": "note", "parts": [f"@PL!A{PLA['note']}"]}, is_tbl, is_chart, {"type": "note", "parts": [f"@PL!A{PLA['note2']}"]}, chk_tbl]
               + ([{"type": "note", "parts": [pl_intro], "collapse": True, "summary": "展開：三情境損益頁說明"}] if pl_intro else []) + old)
# ---- 6. 驅動因子：分部營收推導 ----
seg_chart = {"type": "theme", "title": "分部營收堆疊（情境 B，十億美元）：North America、International、AWS 其他、AWS AI", "x": f"@PL!D{ph}:L{ph}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.2,
             "total": f"@PL!D{sx['tot']}:L{sx['tot']}",
             "bars": [{"name": [f"@PL!A{sx[k]}"], "y": f"@PL!D{sx[k]}:L{sx[k]}", "color": c} for k, c in (("na", "#c29a2b"), ("intl", "#e6c87a"), ("aws_o", "#6fc7cb"), ("aws_ai", "#0b7a80"))],
             "caption": "堆疊＝四個營收塊（零售兩個地區分部、AWS 拆成非 AI 雲端與對外 AI 算力；AWS 的拆分為估計）；頂端數字＝合併營收（FY2024A、FY2025A 為實績，FY2026E 含九個月實績）。"}
seg_tbl = {"type": "table", "caption": "分部營收（十億美元；情境 B）：四塊加總＝合併營收", "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in "DEFGHIJKL"],
           "rows": [[f"@PL!A{sx[k]}"] + [f"@PL!{c}{sx[k]}|n1" for c in "DEFGHIJKL"] for k in ("na", "intl", "aws_o", "aws_ai", "tot")]}
S["drivers"]["blocks"] = [{"type": "note", "parts": [f"@README!A{RMA['drv']}"]}, seg_chart, seg_tbl] + S["drivers"]["blocks"]
S["drivers"]["title"] = "驅動因子：分部營收推導"
# ---- 8. 對帳：三層協調 ----
CW = wb["Coord"]
f1h = find("Coord", "F1. 本模型讀取"); f1b = find("Coord", "F1b. 追公式")
F1a, _ = CRD["F1"]; rtr0, rtrc = CRD["rtr"]
rd1, rd2, rj, rjr, rg, rl, rc9, rai, rmw = (CRD[k] for k in ("rd1", "rd2", "rj", "rjr", "rg", "rl", "rc9", "rai", "rmw"))
rq = CRD["rq"]; G = CRD["rgw"]
capx = {"type": "table", "caption": "資本支出（公司已公布實績 vs Wattonomics 隱含值；十億美元）", "head": ["讀數", "期間／口徑", "公司實績", "層級隱含", "差距", "狀態／判定"],
        "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!D{r}|n1", f"@Coord!E{r}|n1", f"@Coord!F{r}|sp1", f"@Coord!G{r}"] for r in (rd1, rd2, rc9, rg)]
                + [[f"@Coord!A{rj}", f"@Coord!B{rj}", "", "", "", f"@Coord!G{rj}"], [f"@Coord!A{rjr}", "", "", "", "", f"@Coord!G{rjr}"]]}
nocmp = {"type": "table", "caption": "沒有可比公司實績的項目（不判定）", "head": ["項目", "說明", "狀態"],
         "rows": [[f"@Coord!A{r}", f"@Coord!H{r}", f"@Coord!G{r}"] for r in (rai, rmw, G["q"], G["j"])]}
gw_t = {"type": "table", "caption": "新增電力容量（參考；口徑無法對齊）", "head": ["讀數", "期間／口徑", "公司", "層級隱含", "差距", "狀態"],
        "rows": [[f"@Coord!A{G['y25']}", f"@Coord!B{G['y25']}", f"@Coord!D{G['y25']}|n2", f"@Coord!E{G['y25']}|n2", f"@Coord!F{G['y25']}|sp1", f"@Coord!G{G['y25']}"],
                 [f"@Coord!A{G['l']}", f"@Coord!B{G['l']}", "", f"@Coord!E{G['l']}|n2", "", ""]]}
d3 = {"type": "table", "caption": "實績對帳的輸入、口徑、出處與證據等級", "collapse": True, "summary": "展開：實績對帳的輸入、出處與證據等級",
      "head": ["項目", "期間／口徑", "公司實績口徑與出處", "公司實績等級", "層級口徑與出處", "層級等級"],
      "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!H{r}", f"@Coord!I{r}", f"@Coord!J{r}", f"@Coord!K{r}"] for r in list(rq) + [CRD["rg9"], rl, rd1, rd2, rg, G["y25"], G["l"]]]}
trc = {"type": "table", "caption": f"@Coord!A{f1b}", "head": ["輸出名稱 ← 公式鏈", "同源判定", "讀到什麼"],
       "rows": [[f"@Coord!A{r}", f"@Coord!G{r}", f"@Coord!H{r}"] for r in range(rtr0, rtrc)]}
wr = CRD["wrows"]
t1 = {"type": "table", "caption": f"@Coord!A{f1h}", "collapse": True, "summary": "展開：本模型讀取的 Tokenomics 名稱與 Wattonomics 對照名稱、值與版本",
      "head": ["名稱", "年度", "層級值", "本模型採用值", "單位", "層級與版本"],
      "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}|n3", f"@Coord!D{r}|n3", f"@Coord!E{r}", f"@Coord!F{r}"] for r in range(F1a + 1, wr["act26"] + 1) if CW.cell(r, 3).value is not None]}
d0, d1 = CRD["rdump"]
dump_t = {"type": "table", "caption": "追公式傾印（腳本輸出：每格公式與值）", "collapse": True, "summary": "展開：追公式傾印（每格公式與值；腳本 trace_wattonomics.py 的輸出）",
          "head": ["儲存格（縮排＝依賴深度）", "值", "標記", "公式"],
          "rows": [[f"@Coord!A{r}", f"@Coord!D{r}|n3", f"@Coord!G{r}", f"@Coord!H{r}"] for r in range(d0, d1 + 1)]}
S["recon"]["blocks"].extend([
    {"type": "h3", "text": "三層協調：讀取層級與實績對帳"},
    {"type": "note", "parts": [f"@Coord!D{CRD['rs']}"]}, capx, nocmp, gw_t, d3,
    {"type": "note", "parts": [f"@Coord!B{CRD['rr']}"]}, {"type": "note", "parts": [f"@Coord!B{CRD['rsame']}"]},
    {"type": "h3", "text": "追公式：層級值是否讀公司實績或目標"}, trc, {"type": "note", "parts": [f"@Coord!A{rtrc}"]}, dump_t, t1])
# ---- 9. 資料狀態：本版要點 ----
S["status"]["blocks"][0:0] = [{"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RMA['ver']}"]}]
order = ["thesis", "valuation", "sens", "revdcf", "pl", "drivers", "recon", "status"]
S["revdcf"] = rev_
cfg["sections"] = [S[k] for k in order]
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", OUT)
