# 產生 physical_ai/tesla_robotaxi/html.json（HTML 展示層，版面規格 v1.3；未上市分部模型）。
# 列號依 A 欄標籤從成品 xlsx 找，不寫死；勿手改 html.json，改本程式後重跑。
# 用法：python3 physical_ai/tesla_robotaxi/gen_html_json.py
# 之後：python3 tools/company_html/build_company_html.py physical_ai/tesla_robotaxi
import glob
import json
import os

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
X = sorted(glob.glob(os.path.join(HERE, "dist", "*.xlsx")))[-1]
wb = openpyxl.load_workbook(X, data_only=True)

YC = "CDEFGHIJKL"          # 2026–2035
FC = "DEFGHIJKL"           # 2027–2035
F28 = "EFGHIJKL"           # 2028–2035


def find(sheet, label, start=1, exact=False):
    ws = wb[sheet]
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and (v == label if exact else v.startswith(label)):
            return r
    raise SystemExit(f"{sheet} 找不到標籤「{label}」（自第 {start} 列起）")


def rng(sheet, r, cols):
    return f"@{sheet}!{cols[0]}{r}:{cols[-1]}{r}"


# ───── 各頁列號 ─────
MK, PM, VA, SE, PE = "市場天花板與情境", "每英里經濟", "估值", "Sens", "滲透"
mk0 = {s: find(MK, f"情境 {s}（") for s in "ABC"}
pm0 = {s: find(PM, f"情境 {s}（") for s in "ABC"}
va0 = {s: find(VA, f"情境 {s}（") for s in "ABC"}
MKR = {s: {k: find(MK, lab, mk0[s]) for k, lab in (
    ("yr", "項目"), ("pen", "Robotaxi 滲透率"), ("price", "Tesla 實收每英里價格"), ("trips", "美國 Robotaxi 趟數"),
    ("conv", "自有車 VMT 轉換里程"), ("us", "美國 Robotaxi 總里程"), ("share", "Tesla 份額"), ("tus", "Tesla 美國里程"),
    ("glob", "Tesla 全球網路付費里程"), ("own", "Tesla 自有車隊付費里程"), ("host", "車主車輛付費里程"),
    ("rev", "Tesla Robotaxi 營收"), ("ratio", "Robotaxi 趟數 ÷ 叫車趟數"), ("vmt", "Robotaxi 里程 ÷ 小客車 VMT"))} for s in "ABC"}
PMR = {s: {k: find(PM, lab, pm0[s]) for k, lab in (
    ("yr", "項目"), ("price", "價格（實收每英里）"), ("opx", "營運成本（不含折舊"), ("mpv", "每車年付費英里"),
    ("dep", "車輛折舊每英里"), ("sga", "平台管銷分攤每英里"), ("full", "每英里全成本"), ("gap", "價格 − 全成本"),
    ("flag", "是否低於成本"), ("vs_own", "價格 ÷ 自有車每英里全成本"), ("fleet", "Tesla 自有車隊（年底）"),
    ("ebit", "分部 EBIT"), ("end", "EBIT 率"))} for s in "ABC"}
VAR = {s: {k: find(VA, lab, va0[s]) for k, lab in (("yr", "項目"), ("ebit", "EBIT"), ("end", "每股（÷ 隱含股數"))} for s in "ABC"}
v_sum = find(VA, "分部價值彙總")
v_ev = find(VA, "分部價值（EV）", v_sum)
v_share = find(VA, "分部價值 ÷ TSLA 市值", v_sum)
v_rate = find(VA, "評等", v_sum)
v_wm = find(VA, "對照：Waymo")
v_wm_end = find(VA, "讀法：Waymo")
mk_cmp = find(MK, "情境比較")
pm_obs = find(PM, "2026 觀察")
pm_obs_h, pm_obs_px, pm_obs_full = pm_obs + 1, find(PM, "Austin 現行牌價每英里", pm_obs), find(PM, "每英里全成本（2026", pm_obs)
pe_a, pe_ah = find(PE, "A. 價格"), find(PE, "日期")
pe_now = max(r for r in range(pe_ah + 1, find(PE, "實付觀察")) if wb[PE].cell(r, 2).value == "Austin")   # 本模型現行牌價（最後一筆 Austin 牌價）
pe_own = find(PE, "錨 2：自有車每英里全成本")
pe_vs_own = find(PE, "現行牌價 ÷ 錨 2")
pe_el = find(PE, "觀察到的價格彈性")
pe_b, pe_bh = find(PE, "B. 每城市現況"), find(PE, "城市", find(PE, "B. 每城市現況"))
pe_bt = find(PE, "合計（自駕里程）")
pe_c, pe_ch = find(PE, "C. 全國時間序列"), find(PE, "期末")
pe_ce = find(PE, "注意：累計付費里程") - 1
pe_wk = find(PE, "每週付費趟數（估）")
pe_d, pe_dh = find(PE, "D. 與 Waymo 並排"), find(PE, "指標", find(PE, "D. 與 Waymo 並排"))
pe_de = wb[PE].max_row
pe_cities = find(PE, "營運城市（無人）", pe_dh)
pe_tx = find(PE, "德州 AV 登記車數", pe_dh)
pe_wkd = find(PE, "每週付費趟數", pe_dh)
pe_mi = max(r for r in range(pe_ch + 1, pe_ce + 1) if isinstance(wb[PE].cell(r, 8).value, (int, float)))   # 最新的每日無人里程
s_hdr = find(SE, "驅動")
s_base = find(SE, "基準（情境加權）")
s_end = find(SE, "讀法") - 2
s_opx = find(SE, "情境 B 的線營運成本路徑")
s_stat = find(SE, "模型狀態")
s_bid = find(SE, "build_id")
T = "給Tesla模型"
t_lo, t_base, t_hi = (find(T, n, exact=True) for n in ("RT_VAL_LO", "RT_VAL_BASE", "RT_VAL_HI"))
IN = "In"
in_end = find(IN, "8. 線 1a") - 1
in_px, in_mcap = find(IN, "px", exact=True), find(IN, "mcap", exact=True)
in_w = {s: find(IN, f"w_{s}", exact=True) for s in "ABC"}

mb, pb = MKR["B"], PMR["B"]
ma, pa = MKR["A"], PMR["A"]
pc = PMR["C"]

# ───── 敏感度：低／高案相對基準（Sens F、G 欄）─────
S = wb[SE]
case_rows = list(range(s_base + 1, s_end + 1))
sens_main, sens_more = [], []
for side, col in (("低", "F"), ("高", "G")):
    rs = [r for r in case_rows if isinstance(S[f"{col}{r}"].value, (int, float)) and S[f"{col}{r}"].value != 0]
    big = [r for r in rs if abs(S[f"{col}{r}"].value) >= 0.1]
    small = [r for r in rs if abs(S[f"{col}{r}"].value) < 0.1]
    base = {"type": "tornado", "sheet": SE, "label_col": "A", "value_col": col, "fmt": "sp1"}
    sens_main.append(dict(base, title=f"加權分部價值相對基準：輸入取{side}值，變動達兩位數百分比的案例", rows=",".join(map(str, big or rs[:1]))))
    if small:
        sens_more.append(dict(base, title=f"加權分部價值相對基準：輸入取{side}值，其餘案例", rows=",".join(map(str, small)),
                              collapse=True, summary=f"展開：輸入取{side}值，變動未達兩位數百分比的案例"))


def pm_fmt(s):
    p = PMR[s]
    f = {}
    for k in ("price", "opx", "dep", "sga", "full", "gap"):
        f[str(p[k])] = {c: "n2" for c in YC}
    for k in ("mpv", "fleet"):
        f[str(p[k])] = {c: "n0" for c in YC}
    f[str(p["vs_own"])] = {c: "n2" for c in YC}
    f[str(p["vs_own"] - 1)] = {c: "p0" for c in YC}
    f[str(p["end"])] = {c: "p0" for c in YC}
    for r in range(p["fleet"] + 1, p["end"]):
        f[str(r)] = {c: "n2" for c in YC}
    f[str(p["mpv"] + 1)] = {c: "n0" for c in YC}      # 單車取得成本
    return f


def pm_table(s, collapse):
    p = PMR[s]
    b = {"type": "table", "grid": True, "caption": f"@{PM}!A{pm0[s]}",
         "head": [f"@{PM}!A{p['yr']}", f"@{PM}!B{p['yr']}"] + [f"@{PM}!{c}{p['yr']}|yr" for c in YC],
         "rowgen": {"sheet": PM, "rows": f"{p['price']}-{p['end']}", "cols": ["A", "B"] + list(YC),
                    "fmts": ["", ""] + ["n2"] * len(YC), "row_fmt": pm_fmt(s)}}
    if collapse:
        b.update(collapse=True, summary=f"展開：情境 {s} 每英里經濟與分部 EBIT 全表")
    return b


def mk_fmt(s):
    m = MKR[s]
    f = {}
    for k in ("pen", "share", "vmt"):
        f[str(m[k])] = {c: "p1" for c in YC}
    f[str(m["share"] + 4)] = {c: "p1" for c in YC}    # 車主車輛占比
    f[str(m["vmt"])] = {c: "p2" for c in YC}
    f[str(m["price"])] = {c: "n2" for c in YC}
    f[str(m["trips"] - 1)] = {c: "n2" for c in YC}    # 誘發倍數
    f[str(m["pen"] - 1)] = {c: "n2" for c in YC}      # S 曲線
    f[str(m["trips"] + 1)] = {c: "n2" for c in YC}    # 閘門
    f[str(m["share"] + 2)] = {c: "n2" for c in YC}    # 海外 ÷ 美國
    f[str(m["ratio"])] = {c: "n2" for c in YC}
    f[str(m["us"] - 1)] = {c: "n1" for c in YC}       # 每趟英里
    return f


def mk_table(s, collapse):
    m = MKR[s]
    b = {"type": "table", "grid": True, "caption": f"@{MK}!A{mk0[s]}",
         "head": [f"@{MK}!A{m['yr']}", f"@{MK}!B{m['yr']}"] + [f"@{MK}!{c}{m['yr']}|yr" for c in YC],
         "rowgen": {"sheet": MK, "rows": f"{m['yr'] + 1}-{m['vmt']}", "cols": ["A", "B"] + list(YC),
                    "fmts": ["", ""] + ["n2"] * len(YC), "row_fmt": mk_fmt(s)}}
    if collapse:
        b.update(collapse=True, summary=f"展開：情境 {s} 市場天花板推導全表")
    return b


def va_table(s):
    v = VAR[s]
    return {"type": "table", "grid": True, "caption": f"@{VA}!A{va0[s]}", "collapse": True, "summary": f"展開：情境 {s} 自由現金流與分部價值",
            "head": [f"@{VA}!A{v['yr']}", f"@{VA}!B{v['yr']}"] + [f"@{VA}!{c}{v['yr']}|yr" for c in FC],
            "rowgen": {"sheet": VA, "rows": f"{v['ebit']}-{v['end']}", "cols": ["A", "B", "C"] + list(FC),
                       "fmts": ["", "", "n2"] + ["n2"] * len(FC),
                       "row_fmt": {str(v["ebit"] + 7): {c: "n3" for c in FC}, str(v["ebit"] + 13): {"C": "p1"}}}}


def nourl(sheet, co):
    """含網址的出處格不顯示（自含規格禁止網址），改為指向 xlsx。"""
    v = wb[sheet][co].value
    return "出處含網址，見 xlsx" if isinstance(v, str) and "http" in v else f"@{sheet}!{co}"


YRS28 = [str(y) for y in range(2028, 2036)]
YRS27 = [str(y) for y in range(2027, 2036)]
YRS26 = [str(y) for y in range(2026, 2036)]

cfg = {
    "company": "Tesla Robotaxi 分部", "ticker": "Tesla Robotaxi（Tesla 分部，未上市）", "layout": "v1.3", "wide_text": True,
    "thesis": "@README!A2",
    "question": "每英里實收價格能不能壓到自有車每英里成本附近、何時壓到，以及 Tesla 在 Robotaxi 里程裡拿到多少；做不到時，分部只是在叫車市場裡分份額。",
    "scenarios": [{"key": "A", "name": "情境 A（高滲透、低成本、低價）"}, {"key": "B", "name": "情境 B（線基準）"}, {"key": "C", "name": "情境 C（低滲透、高成本、高價）"}],
    "hero": {
        "rec": f"@{VA}!C{v_rate}",
        "tiles": [[f"@{T}!B{t_lo}", f"@{T}!C{t_lo}|n1", "十億美元，評價日現值"],
                  [f"@{T}!B{t_base}", f"@{T}!C{t_base}|n1", ["十億美元；約占 TSLA 市值 ", f"@{VA}!F{v_share}|p1"]],
                  [f"@{T}!B{t_hi}", f"@{T}!C{t_hi}|n1", "十億美元，評價日現值"]],
        "note": ["未上市分部，不給目標價與評等，只給價值區間。基準情境價值接近零：", f"@{PM}!I{pb['yr']}|yr", " 年每英里價格 ",
                 f"@{PM}!I{pb['price']}|n2", " 美元，對每英里全成本 ", f"@{PM}!I{pb['full']}|n2",
                 " 美元，沒有利差；加權價值主要來自情境 A（價格壓到自有車成本以下、私人里程轉入）。"],
    },
    "sections": [],
}

thesis = {"id": "thesis", "title": "命題與核心問題", "blocks": [
    {"type": "theme", "title": "每英里實收價格 vs 每英里全成本（美元／英里）：2028–2035，三情境", "x": YRS28, "fmt": "n2",
     "lines": [{"name": "價格：情境 A", "y": rng(PM, pa["price"], F28), "color": "#0b7a80"},
               {"name": "價格：情境 B", "y": rng(PM, pb["price"], F28), "color": "#c29a2b", "labels": True},
               {"name": "價格：情境 C", "y": rng(PM, pc["price"], F28), "color": "#6f6bae"},
               {"name": "全成本：情境 A", "y": rng(PM, pa["full"], F28), "color": "#0b7a80", "dash": True},
               {"name": "全成本：情境 B", "y": rng(PM, pb["full"], F28), "color": "#c29a2b", "dash": True}],
     "caption": ["實線＝Tesla 實收每英里價格，虛線＝每英里全成本（營運＋車輛折舊＋管銷分攤）。情境 C 全成本遠高於圖上範圍（",
                 f"@{PM}!L{pc['yr']}|yr", " 年 ", f"@{PM}!L{pc['full']}|n2", " 美元），未畫。比較錨：自有車每英里全成本 ",
                 f"@{PE}!G{pe_own}|n2", " 美元。"]},
    {"type": "h3", "text": "命題（分點）"},
    {"type": "list", "ordered": True, "items": [
        ["價格是承重變數：Austin 現行牌價中程每英里 ", f"@{PE}!G{pe_now}|n2", " 美元，是自有車每英里全成本 ", f"@{PE}!G{pe_own}|n2",
         " 美元的 ", f"@{PE}!G{pe_vs_own}|n1", " 倍；要壓到自有車成本附近，私人開車里程才會轉入，天花板才會從叫車市場跳到小客車行駛里程。"],
        ["壓不到的話，市場只是叫車：情境 B 到 ", f"@{MK}!L{mb['yr']}|yr", " 年美國 Robotaxi 趟數為叫車趟數的 ", f"@{MK}!L{mb['ratio']}|n2",
         " 倍，Robotaxi 里程只占小客車行駛里程 ", f"@{MK}!L{mb['vmt']}|p1", "。"],
        ["成本決定有沒有利潤：情境 B ", f"@{PM}!I{pb['yr']}|yr", " 年每英里價格 ", f"@{PM}!I{pb['price']}|n2", " 美元、全成本 ",
         f"@{PM}!I{pb['full']}|n2", " 美元，量放大也不改變價值；情境 A 同年價格 ", f"@{PM}!I{pa['price']}|n2", " 美元、全成本 ",
         f"@{PM}!I{pa['full']}|n2", " 美元。"],
        ["價值落在誰手上：Tesla 份額（情境 B 到 ", f"@{MK}!L{mb['yr']}|yr", " 年 ", f"@{MK}!L{mb['share']}|p0",
         "）與車主車輛的抽成決定 Tesla 能留下多少；Waymo 等對手共用同一市場基準。"]]},
    {"type": "list", "collapse": True, "summary": "展開：命題回覆（依模型計算逐一回答）", "ordered": True, "items": [
        ["價格能不能壓到自有車成本附近？答：", f"@{PM}!I{pa['yr']}|yr", " 年價格 ÷ 自有車每英里全成本，情境 A ", f"@{PM}!I{pa['vs_own']}|n2",
         " 倍、情境 B ", f"@{PM}!I{pb['vs_own']}|n2", " 倍、情境 C ", f"@{PM}!I{pc['vs_own']}|n2", " 倍；只有情境 A 壓到自有車成本以下。"],
        ["何時跨過門檻？答：情境 A 自 ", f"@{PM}!E{pa['yr']}|yr", " 年起價格高於全成本（判定：", f"@{PM}!E{pa['flag']}", "），自有車轉換閘門在 ",
         f"@{MK}!H{ma['yr']}|yr", " 年全開（", f"@{MK}!H{ma['trips'] + 1}|n2", "）；情境 B 到 ", f"@{PM}!L{pb['yr']}|yr", " 年仍",
         f"@{PM}!L{pb['flag']}", "（每英里差 ", f"@{PM}!L{pb['gap']}|n3", " 美元）。"],
        ["Tesla 能拿多少？答：", f"@{MK}!L{mb['yr']}|yr", " 年全球網路付費里程，情境 A ", f"@{MK}!L{ma['glob']}|n0", "、情境 B ",
         f"@{MK}!L{mb['glob']}|n1", "、情境 C ", f"@{MK}!L{MKR['C']['glob']}|n2", " 十億英里；營收情境 A ", f"@{MK}!L{ma['rev']}|n1",
         "、情境 B ", f"@{MK}!L{mb['rev']}|n1", "、情境 C ", f"@{MK}!L{MKR['C']['rev']}|n2", " 十億美元。"],
        ["分部值多少？答：低（情境 C）", f"@{T}!C{t_lo}|n1", "、基準（加權）", f"@{T}!C{t_base}|n1", "、高（情境 A）", f"@{T}!C{t_hi}|n1",
         " 十億美元；基準約占 TSLA 市值 ", f"@{VA}!F{v_share}|p1", "。"]]},
    {"type": "h3", "text": "重點觀察"},
    {"type": "list", "ordered": True, "items": [
        ["基準情境沒有利潤：情境 B 分部價值 ", f"@{VA}!D{v_ev}|n1", " 十億美元，", f"@{PM}!I{pb['yr']}|yr", " 年後每英里價格與全成本幾乎相同。"],
        ["加權價值靠情境 A 的私人里程轉入：情境 A 價值 ", f"@{VA}!C{v_ev}|n1", " 十億美元；", f"@{MK}!L{ma['yr']}|yr",
         " 年美國 Robotaxi 里程 ", f"@{MK}!L{ma['us']}|n0", " 十億英里中，自有車轉換占 ", f"@{MK}!L{ma['conv']}|n0", " 十億英里。"],
        ["現況規模很小：無人營運城市 ", f"@{PE}!B{pe_cities}", " 個（Waymo ", f"@{PE}!C{pe_cities}", "），每日無人里程約 ",
         f"@{PE}!H{pe_mi}|n0", " 英里，推估每週付費趟數約 ", f"@{PE}!H{pe_wk}|n0", "；德州登記車數 ", f"@{PE}!B{pe_tx}|n0",
         " 台，但登記不等於營運。"],
        ["現在每英里還在大幅補貼：Austin 中程牌價每英里 ", f"@{PM}!D{pm_obs_px}|n2", " 美元，對 ", f"@{PM}!D{pm_obs_full}|n2",
         " 美元的全成本；稼動率低、安全員與遠端監看占營運成本大宗。"],
        ["最大風險在成本路徑：情境 B 改用高成本路徑，加權價值變為 ", f"@{SE}!D{s_opx}|n1", " 十億美元（相對基準 ",
         f"@{SE}!F{s_opx}|sp0", "）；改用低成本路徑為 ", f"@{SE}!E{s_opx}|n1", "。"]]},
]}

valuation = {"id": "valuation", "title": "估值：分部價值區間（未上市，不給目標價與評等）", "blocks": [
    {"type": "theme", "title": "分部價值（十億美元，評價日現值）：三情境與加權", "x": ["情境 A（高）", "情境 B", "情境 C（低）", "加權（基準）"],
     "fmt": "n1", "stack": True, "total": f"@{VA}!C{v_ev}:F{v_ev}",
     "bars": [{"name": f"@{VA}!A{v_ev}", "y": f"@{VA}!C{v_ev}:F{v_ev}", "color": "#0b7a80"}],
     "lines": [{"name": f"@{VA}!A{v_ev - 1}", "y": f"@{VA}!C{v_ev - 1}:F{v_ev - 1}", "color": "#c29a2b", "dash": True}],
     "caption": ["長條＝分部價值（企業價值，評價日現值），頂端為數值；虛線點＝其中終值現值（各情境獨立，連線只為辨識）。加權權重 情境 A／B／C＝", f"@In!E{in_w['A']}|p0", "／", f"@In!E{in_w['B']}|p0", "／", f"@In!E{in_w['C']}|p0",
                 "。本分部未上市、沒有市場價格，不設目標價與上檔空間。"]},
    {"type": "note", "parts": [f"@{VA}!A1"]},
    {"type": "note", "parts": [f"@{VA}!A2"]},
    {"type": "table", "grid": True, "caption": f"@{VA}!A{v_sum}",
     "head": [f"@{VA}!A{v_sum + 1}", f"@{VA}!C{v_sum + 1}", f"@{VA}!D{v_sum + 1}", f"@{VA}!E{v_sum + 1}", f"@{VA}!F{v_sum + 1}"],
     "rowgen": {"sheet": VA, "rows": f"{v_sum + 2}-{v_share}", "cols": ["A", "C", "D", "E", "F"], "fmts": ["", "n1", "n1", "n1", "n1"],
                "row_fmt": {str(v_ev + 1): {c: "p0" for c in "CDEF"}, str(v_share - 1): {c: "n2" for c in "CDEF"}, str(v_share): {c: "p1" for c in "CDEF"}}}},
    {"type": "kv", "items": [[f"@{VA}!A{v_rate}", f"@{VA}!C{v_rate}"]]},
    {"type": "h3", "text": "折現率"},
    {"type": "table", "grid": True, "head": ["項目", "值"],
     "rows": [[f"@{VA}!A5", f"@{VA}!C5|p2"], [f"@{VA}!A6", f"@{VA}!C6|n1"], [f"@{VA}!A7", f"@{VA}!C7|p1"], [f"@{VA}!A8", f"@{VA}!C8|p2"],
              [f"@{VA}!A9", f"@{VA}!C9|p1"], [f"@{VA}!A10", f"@{VA}!C10|p2"], [f"@{VA}!A11", f"@{VA}!C11|n3"]]},
    va_table("A"), va_table("B"), va_table("C"),
]}

sens = {"id": "sens", "title": "敏感度：加權分部價值相對基準", "blocks": sens_main + [
    {"type": "note", "parts": [f"@{SE}!A1"]},
    {"type": "note", "parts": [f"@{SE}!A2"], "collapse": True, "summary": "展開：敏感度方法說明（哪些參數沒有區間）"},
    {"type": "note", "parts": [f"@{SE}!A{s_end + 2}"]}] + sens_more + [
    {"type": "table", "collapse": True, "summary": "展開：全部案例（低案、高案的加權分部價值與變化）",
     "head": [f"@{SE}!{c}{s_hdr}" for c in "ABDEFG"],
     "rowgen": {"sheet": SE, "rows": f"{s_base}-{s_end}", "cols": list("ABDEFG"), "fmts": ["", "", "n1", "n1", "sp1", "sp1"]}},
    {"type": "kv", "items": [[f"@{SE}!B{s_stat}", f"@{SE}!C{s_stat}|n2"], [f"@{SE}!A{s_stat}", f"@{SE}!D{s_stat}"], [f"@{SE}!A{s_bid}", f"@{SE}!D{s_bid}"]],
     "collapse": True, "summary": "展開：敏感度頁模型狀態與建置識別"},
]}

revdcf = {"id": "revdcf", "title": "反向檢驗：TSLA 市值與對手估值（分部未上市，無反向 DCF）", "blocks": [
    {"type": "note", "parts": ["本分部沒有自己的市場價格，無法做反向 DCF。改列兩個對照：分部價值占 TSLA 市值多少，以及用 Waymo 融資估值的趟數倍數套到 Tesla 的現況規模。兩者都只作對照，不進結論。"]},
    {"type": "table", "grid": True, "head": [f"@{VA}!A{v_sum + 1}", f"@{VA}!C{v_sum + 1}", f"@{VA}!D{v_sum + 1}", f"@{VA}!E{v_sum + 1}", f"@{VA}!F{v_sum + 1}"],
     "rows": [[f"@{VA}!A{v_ev}", f"@{VA}!C{v_ev}:F{v_ev}|n1"], [f"@{VA}!A{v_share}", f"@{VA}!C{v_share}:F{v_share}|p1"]]},
    {"type": "table", "head": ["項目", "值", "標記", "出處"],
     "rows": [[f"@{IN}!B{in_px}", f"@{IN}!E{in_px}|n2", f"@{IN}!G{in_px}|tag", f"@{IN}!I{in_px}"],
              [f"@{IN}!B{in_mcap}", f"@{IN}!E{in_mcap}|n0", f"@{IN}!G{in_mcap}|tag", f"@{IN}!I{in_mcap}"]]},
    {"type": "table", "caption": f"@{VA}!A{v_wm}", "head": ["項目", "值", "說明"],
     "rows": [[f"@{VA}!A{r}", f"@{VA}!C{r}|n0" if r != v_wm + 4 else f"@{VA}!C{r}|n1", nourl(VA, f"D{r}")] for r in range(v_wm + 1, v_wm_end)]},
    {"type": "note", "parts": [f"@{VA}!A{v_wm_end}"]},
]}

pl = {"id": "pl", "title": "每英里經濟與分部營收推導（取代標準損益表）", "blocks": [
    {"type": "theme", "title": "情境 B 每英里成本堆疊 vs 每英里價格（美元／英里）：2028–2035", "x": YRS28, "fmt": "n2", "stack": True, "pad": 1.1,
     "total": rng(PM, pb["full"], F28),
     "bars": [{"name": f"@{PM}!A{pb['opx']}", "y": rng(PM, pb["opx"], F28), "color": "#0b7a80"},
              {"name": f"@{PM}!A{pb['dep']}", "y": rng(PM, pb["dep"], F28), "color": "#6fc7cb"},
              {"name": f"@{PM}!A{pb['sga']}", "y": rng(PM, pb["sga"], F28), "color": "#c29a2b"}],
     "lines": [{"name": f"@{PM}!A{pb['price']}", "y": rng(PM, pb["price"], F28), "color": "#b4472f"}],
     "caption": ["堆疊＝每英里全成本（營運成本含能源、維修清潔、保險、遠端監看、安全員、場站；另加車輛折舊與管銷分攤），頂端數字為全成本合計；折線＝實收每英里價格。",
                 f"@{PM}!D{pb['yr']}|yr", " 年全成本 ", f"@{PM}!D{pb['full']}|n2", " 美元，超出圖上範圍未畫。模型沒有把營運成本再拆成安全員、遠端監看等細項。"]},
    {"type": "theme", "title": "稼動率：每車年付費英里（英里／年）：2027–2035，三情境", "x": YRS27, "fmt": "n0",
     "lines": [{"name": "情境 A", "y": rng(PM, pa["mpv"], FC), "color": "#0b7a80"},
               {"name": "情境 B", "y": rng(PM, pb["mpv"], FC), "color": "#c29a2b", "labels": True},
               {"name": "情境 C", "y": rng(PM, pc["mpv"], FC), "color": "#6f6bae"}],
     "caption": "每車年付費英里＝線的車隊加權每車每日付費英里 × 一年天數 × 稼動率乘數；空駛與等候不計入付費英里，稼動率越低，每付費英里分攤的折舊與固定營運成本越高。"},
    {"type": "note", "parts": ["口徑：本表不是公司揭露的 GAAP 或 Non-GAAP 損益表。Tesla 不單獨揭露 Robotaxi 分部損益，以下為分部模型推估的每英里經濟與分部 EBIT（營收減營運成本、折舊、平台研發管銷），稅後營業利益與自由現金流見估值章。"]},
    {"type": "note", "parts": [f"@{PM}!A1"]},
    {"type": "note", "parts": [f"@{PM}!A2"]},
    pm_table("B", False),
    {"type": "table", "grid": True, "caption": f"@{PM}!A{pm_obs}",
     "head": [f"@{PM}!{c}{pm_obs_h}" for c in "ABCDE"],
     "rowgen": {"sheet": PM, "rows": f"{pm_obs_h + 1}-{pm_obs_h + 5}", "cols": list("ABCDE"), "fmts": ["", "", "n2", "n2", "n2"]}},
    {"type": "kv", "items": [[f"@{PM}!A{pm_obs_full}", f"@{PM}!F{pm_obs_full}"]]},
    pm_table("A", True), pm_table("C", True),
]}

drivers = {"id": "drivers", "title": "驅動因子：滲透（城市、車隊、付費里程）與市場天花板", "blocks": [
    {"type": "theme", "title": "情境 B Tesla 全球網路付費里程（十億英里，堆疊）與美國 Robotaxi 滲透率（右軸）：2027–2035", "x": YRS27, "fmt": "n1", "fmt_r": "p0",
     "stack": True, "pad": 1.15, "total": rng(MK, mb["glob"], FC),
     "bars": [{"name": f"@{MK}!A{mb['own']}", "y": rng(MK, mb["own"], FC), "color": "#0b7a80"},
              {"name": f"@{MK}!A{mb['host']}", "y": rng(MK, mb["host"], FC), "color": "#c29a2b"}],
     "lines": [{"name": f"@{MK}!A{mb['pen']}", "y": rng(MK, mb["pen"], FC), "axis": "r", "color": "#3b3838", "labels": True}],
     "caption": "堆疊＝Tesla 自有車隊與車主車輛的付費里程（合計＝全球網路付費里程，含海外）；折線＝Robotaxi 替代的叫車趟數占比（全部營運商）。分部營收＝自有車隊里程 × 價格＋車主車輛里程 × 價格 × Tesla 抽成。"},
    {"type": "theme", "title": "分部營收推導結果：Tesla Robotaxi 營收（十億美元）：2026–2035，三情境", "x": YRS26, "fmt": "n1",
     "lines": [{"name": "情境 A", "y": rng(MK, ma["rev"], YC), "color": "#0b7a80", "labels": True},
               {"name": "情境 B", "y": rng(MK, mb["rev"], YC), "color": "#c29a2b"},
               {"name": "情境 C", "y": rng(MK, MKR["C"]["rev"], YC), "color": "#6f6bae"}],
     "caption": "營收由市場天花板頁逐年推導：美國叫車趟數 × 滲透率 × 價格誘發＋自有車轉換，乘上 Tesla 份額與每趟英里，再加海外倍數，最後乘價格（車主車輛只計抽成）。"},
    {"type": "note", "parts": [f"@{MK}!A1"]},
    {"type": "note", "parts": [f"@{MK}!A2"]},
    {"type": "h3", "text": "一、現況滲透：全國時間序列"},
    {"type": "table", "caption": f"@{PE}!A{pe_c}", "head": [f"@{PE}!{c}{pe_ch}" for c in "ABCDEFGHIK"],
     "rowgen": {"sheet": PE, "rows": f"{pe_ch + 1}-{pe_ce}", "cols": list("ABCDEFGHIK"), "fmts": ["", "n0", "n0", "", "", "n1", "n1", "n0", "n2", ""]}},
    {"type": "note", "parts": [f"@{PE}!A{pe_ce + 1}"]},
    {"type": "kv", "items": [[f"@{PE}!A{pe_wk}", f"@{PE}!H{pe_wk}|n0"]]},
    {"type": "h3", "text": "二、每城市現況與 Waymo 並排"},
    {"type": "table", "caption": f"@{PE}!A{pe_b}", "head": [f"@{PE}!{c}{pe_bh}" for c in "ABCDEFGHIJ"],
     "rowgen": {"sheet": PE, "rows": f"{pe_bh + 1}-{pe_bt}", "cols": list("ABCDEFGHIJ"), "fmts": ["", "", "", "", "", "n0", "n0", "", "n2", "tag"]}},
    {"type": "note", "parts": [f"@{PE}!A{pe_bt + 1}"]},
    {"type": "table", "caption": f"@{PE}!A{pe_d}", "head": [f"@{PE}!{c}{pe_dh}" for c in "ABCD"],
     "rowgen": {"sheet": PE, "rows": f"{pe_dh + 1}-{pe_de}", "cols": list("ABCD"), "fmts": ["", "n2", "n2", ""],
                "row_fmt": {str(pe_tx): {"B": "n0", "C": "n0"}, str(pe_cities): {"B": "n0", "C": "n0"}, str(find(PE, "Austin 服務區", pe_dh)): {"B": "n0", "C": "n0"}}}},
    {"type": "h3", "text": "三、價格：每英里實收價格與比較錨"},
    {"type": "table", "caption": f"@{PE}!A{pe_a}", "head": [f"@{PE}!{c}{pe_ah}" for c in "ABCDEFGHI"],
     "rowgen": {"sheet": PE, "rows": f"{pe_ah + 1}-{pe_el}", "cols": list("ABCDEFGHI"), "fmts": ["", "", "", "n2", "n2", "n2", "n2", "n2", "tag"],
                "row_fmt": {str(pe_vs_own - 1): {c: "p0" for c in "FGH"}}}},
    {"type": "kv", "items": [[f"@{PE}!A{pe_el}", f"@{PE}!K{pe_el}"]]},
    {"type": "h3", "text": "四、市場天花板與情境"},
    {"type": "table", "grid": True, "caption": f"@{MK}!A{mk_cmp}", "head": [f"@{MK}!{c}{mk_cmp + 1}" for c in "ABCDEF"],
     "rowgen": {"sheet": MK, "rows": f"{mk_cmp + 2}-{mk_cmp + 11}", "cols": list("ABCDEF"), "fmts": ["", "yr", "n1", "n1", "n2", "n1"],
                "row_fmt": {str(mk_cmp + 2 + i): {c: ("p1" if i in (4, 5) else "n2") for c in "CDEF"} for i in (2, 3, 4, 5)}}},
    mk_table("B", False), mk_table("A", True), mk_table("C", True),
    {"type": "table", "collapse": True, "summary": "展開：In 頁輸入（第一至第七節；藍字輸入、低／基準／高、標記）",
     "head": [f"@{IN}!{c}3" for c in "ABCDEFG"],
     "rowgen": {"sheet": IN, "rows": f"4-{in_end}", "cols": list("ABCDEFG"), "fmts": ["", "", "", "", "", "", "tag"]}},
]}

recon = {"id": "recon", "title": "對帳：線層級變數與資料來源比較", "blocks": [
    {"type": "note", "parts": ["@對帳!A1"]},
    {"type": "table", "caption": "主要對帳要點（價格、營運成本、稼動率、份額、每趟英里）", "head": [f"@對帳!{c}3" for c in "ACDGI"],
     "rows": [[f"@對帳!{c}{r}" if c != "C" or not isinstance(wb["對帳"][f"C{r}"].value, (int, float)) else f"@對帳!C{r}|n2" for c in "ACDGI"]
              for r in (find("對帳", "Tesla 車隊加權每英里價格"), find("對帳", "Tesla 車隊加權營運成本"), find("對帳", "Tesla 車隊加權每車每日付費英里"),
                        find("對帳", "Tesla 份額 2035"), find("對帳", "Tesla 每趟付費英里"))]},
    {"type": "table", "collapse": True, "summary": "展開：對帳全表（全部線層級變數、口徑、影響儲存格；出處欄含網址，請見 xlsx 對帳頁）",
     "head": [f"@對帳!{c}3" for c in "ABCDEGHI"],
     "rowgen": {"sheet": "對帳", "rows": f"4-{wb['對帳'].max_row}", "cols": list("ABCDEGHI"), "fmts": ["", "", "", "", "", "tag", "", ""]}},
    {"type": "table", "collapse": True, "summary": "展開：資料來源比較（涵蓋、頻率、可靠度、建議用途）",
     "head": [f"@資料來源比較!{c}3" for c in "ABCDEF"],
     "rowgen": {"sheet": "資料來源比較", "rows": f"4-{wb['資料來源比較'].max_row}", "cols": list("ABCDEF"), "fmts": ["", "", "", "", "", "tag"]}},
]}

status = {"id": "status", "title": "資料狀態", "blocks": [
    {"type": "kv", "items": [["版本", "$version"], ["成品日期", "$date"], ["來源 xlsx", "$xlsx"], ["雜湊指紋", "$sha"],
                             ["建置識別", "@README!A1"], ["給 Tesla 模型", f"@{T}!A1"]]},
    {"type": "h3", "text": "失效條件與範圍"},
    {"type": "list", "items": ["@README!A12", "@README!A13"]},
    {"type": "scan", "needle": "未取得", "sheets": {PE: ["A", "B"]}, "collapse": True, "summary": "展開：滲透頁標為「未取得」的資料",
     "empty": "滲透頁沒有標為未取得的資料。"},
    {"type": "table", "collapse": True, "summary": "展開：改版紀錄（前兩版之間的變更清單，工作表名見來源儲存格）",
     "head": [f"@改版紀錄_v0.2!{c}3" for c in "BCDE"],
     "rowgen": {"sheet": "改版紀錄_v0.2", "rows": f"4-{wb['改版紀錄_v0.2'].max_row}", "cols": list("BCDE"), "fmts": ["", "", "", ""]}},
]}

cfg["sections"] = [thesis, valuation, sens, revdcf, pl, drivers, recon, status]
with open(os.path.join(HERE, "html.json"), "w", encoding="utf-8", newline="\n") as fh:
    json.dump(cfg, fh, ensure_ascii=False, indent=1)
    fh.write("\n")
print("html.json 已寫出；xlsx：", os.path.basename(X))
