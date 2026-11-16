# v0.4.1（2026-11-05）：套用個股 HTML 版面規格 v1.0（docs/plan/個股HTML版面規格.md）——只改展示層。v0.5：DCF 改標準版（FCFF＋WACC），估值章 DCF 區塊改由本檔依 xlsx 列號產生（html_src.json 不動）。
# 讀 arista/html_src.json（v0.4 的版面，內容不動）→ 少量前處理 → tools/company_html/relayout.py → 寫 arista/html.json。
# 用法：python3 arista/gen_html_json.py [成品 xlsx（預設 arista/dist/ 內唯一的 xlsx）]
import copy, glob, json, os, re, sys
import openpyxl
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools", "company_html"))
from relayout import relayout  # noqa: E402
X = sys.argv[1] if len(sys.argv) > 1 else (lambda f: f[0] if len(f) == 1 else sys.exit(f"dist 應恰有一個 xlsx：{f}"))(sorted(glob.glob(os.path.join(HERE, "dist", "*.xlsx"))))
cfg = json.load(open(os.path.join(HERE, "html_src.json"), encoding="utf-8"))
wb = openpyxl.load_workbook(X, data_only=True)
IN, VA = wb["In"], wb["Val"]
def row(ws, prefix, col=1):
    return next(r for r in range(1, ws.max_row + 1) if str(ws.cell(r, col).value or "").startswith(prefix))
ink = lambda k: next(r for r in range(1, IN.max_row + 1) if IN.cell(r, 1).value == k)
secs = {s["id"]: s for s in cfg["sections"]}
DC = wb["DCF"]
r_dcf = row(VA, "DCF 目標價（情境加權）")
_rk = row(VA, "股權成本 ke"); _rr = row(VA, "永續期 RONIC"); _rp = row(VA, "明確期 FCFF 現值")
_dh = row(DC, "項目")
# 第二輪修正：DCF 彙總表格式依標籤（模組 v2 stub 期在「＋淨現金與投資」後多兩列，不寫死列偏移）
_fm = lambda r: (lambda lab: "n3" if lab.startswith("評價用稀釋股數") else "p1" if "終值現值占 EV" in lab else "x1" if "隱含終值" in lab else "n2")(str(VA.cell(r, 1).value or "").strip())
DCF_BLOCKS = [
    {"type": "note", "parts": [f"@Val!A{_rk - 2}"]},
    {"type": "table", "caption": "DCF 折現率：股權成本、稅後債務成本、WACC、永續成長率、RONIC", "head": ["項目", "數值", "說明"],
     "rows": [[f"@Val!A{r}", f"@Val!B{r}|" + ("p1" if r == _rk + 2 else "p2"), f"@Val!E{r}"] for r in range(_rk, _rr + 1)]},
    {"type": "table", "caption": "DCF 彙總：明確期現值、終值現值、企業價值、淨現金與每股價值", "head": [f"@Val!{c}{_rk - 1}" for c in "ABCD"],
     "rows": [[f"@Val!A{r}"] + [f"@Val!{c}{r}|" + _fm(r) for c in "BCD"] for r in range(_rp, r_dcf)]
             + [[f"@Val!A{r_dcf}", f"@Val!B{r_dcf}|n2", "", ""]]}]
for s_ in "ABC":
    _a = row(DC, f"情境 {s_}（"); _b = row(DC, "現值（FCFF", _a) if False else next(r for r in range(_a, DC.max_row + 1) if str(DC.cell(r, 1).value or "").startswith("現值（FCFF"))
    _cols = ["A"] + [chr(ord("C") + i) for i in range(6)]
    _pct = {str(r): {c: "p1" for c in _cols[1:]} for r in range(_a + 1, _b + 1) if str(DC.cell(r, 1).value).startswith(("稅率", "折舊攤銷 ÷", "資本支出 ÷"))}
    _pct.update({str(r): {c: "n3" for c in _cols[1:]} for r in range(_a + 1, _b + 1) if str(DC.cell(r, 1).value).startswith("折現因子")})
    DCF_BLOCKS.append({"type": "table", "caption": f"@DCF!A{_a}", "collapse": True, "summary": "展開：本情境各年無槓桿自由現金流明細",
                       "head": [f"@DCF!{c}{_dh}" for c in _cols],
                       "rowgen": {"sheet": "DCF", "rows": f"{_a + 1}-{_b}", "cols": _cols, "fmts": [""] + ["n2"] * 6, "row_fmt": _pct}})

# 前處理（只動版面，不加數字）
# 0. 第二輪修正：「情境驅動」表依 xlsx 的「情境驅動對照」節重建列號（In 頁新增 beta_du、stub 列後，html_src.json 的寫死列號會錯位）
_s0 = row(IN, "情境驅動對照") + 1
_s1 = max(r for r in range(_s0 + 1, IN.max_row + 1) if str(IN.cell(r, 1).value or "").startswith("scn_"))
_nscn = 0
for b in secs["drivers"]["blocks"]:
    if b["type"] == "table" and str(b.get("caption", "")).startswith("情境驅動"):
        b["head"] = [f"@In!{c}{_s0}" for c in "BDEFC"]
        b["rows"] = [[f"@In!B{r}"] + [f"@In!{c}{r}" for c in "DEF"] + [f"@In!C{r}"] for r in range(_s0 + 1, _s1 + 1)]
        _nscn += 1
assert _nscn == 1, _nscn
# 1. 驅動：原「單一參數」表實為輸入 → 推導量對照，交給轉換器放到附錄
for b in secs["drivers"]["blocks"]:
    if str(b.get("caption", "")).startswith("單一參數"):
        b["caption"] = "驅動 → 推導量對照"
# 2. 估值：重點觀察已移到第一章；相對倍數調整由 kv 改為表格（轉換器會把 kv 當成折現率表）；
#    DCF 折現率直接顯示、緊接在 DCF 彙總之前；無標題的 DCF、Comparable 計算表補標題以便指定主要列
va = secs["valuation"]["blocks"]
out, skip = [], False
for b in va:
    if b["type"] == "h3" and b.get("text") == "重點觀察":
        skip = True
        continue
    if skip and b["type"] == "list":
        skip = False
        continue
    if b["type"] == "kv":
        b = {"type": "table", "caption": "相對倍數調整", "head": ["項目", "數值"], "rows": [list(i) for i in b["items"]]}
    if b["type"] == "table" and b.get("caption") == "@Val!A4":      # v0.5：舊 0.1 DCF（股東現金流）兩張表 → 標準版 DCF 折現率、彙總與各情境 FCFF 明細
        out += DCF_BLOCKS
        continue
    if b["type"] == "table" and not b.get("caption") and b.get("rows") and b["rows"][0][0] == "@Val!A13":
        continue
    if b["type"] == "bars":
        b = copy.deepcopy(b)
        for it in b["items"]:
            if it["label"] == "@Val!A23":
                it["label"], it["value"] = f"@Val!A{r_dcf}", f"@Val!B{r_dcf}"
    if b["type"] == "table" and not b.get("caption") and b.get("rows") and b["rows"][0][0] == f"@Val!A{row(VA, '同業前瞻 P/E 中位數')}":
        b = dict(b, caption="Comparable 計算")
    out.append(b)
secs["valuation"]["blocks"] = out
# 3. 對帳：單一大表補標題以便指定主要列
for b in secs["recon"]["blocks"]:
    if b["type"] == "table" and not b.get("caption"):
        b["caption"] = "對帳要點"
tor_title = {b["title"]: b for b in secs["sens"]["blocks"] if b["type"] == "tornado"}
# v0.5：Sens 目標價參數表改為標準版 DCF 的十個案例（tools/fcff_dcf.py write_sens）
SE = wb["Sens"]
_st = row(SE, "目標價參數敏感度")
_se = next(r for r in range(_st + 2, SE.max_row + 1) if str(SE.cell(r, 1).value or "").startswith("檢查：閉式公式"))
for b in secs["sens"]["blocks"]:
    if b["type"] == "table" and b.get("caption") == "@Sens!A21":
        b["caption"] = f"@Sens!A{_st}"
        b["head"] = [f"@Sens!{c}{_st + 1}" for c in "ABCDEFG"]
        b["rowgen"] = dict(b["rowgen"], rows=f"{_st + 2}-{_se - 1}")

scen_tok = re.split(r"[.\s]", IN.cell(row(IN, "情境驅動對照"), 1).value)[0]
SPEC = {
 "in_skip": [scen_tok],
 "in_groups": [
  ["AI 後端乙太網路（AI 營收）", [
    ["AI 後端 scale-out 交換器市場（讀網通市場模型）", ["mkt_26", "mkt_27", "mkt_28", "mkt_29"]],
    ["AI IT 資本支出（對帳參考）", ["itc_26", "itc_27", "itc_28", "itc_29"]],
    ["情境顆數係數", ["um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C"]],
    ["乙太網路占比", ["eth_26", "eth_27", "eth_28", "eth_29"]],
    ["Arista 份額相對值", ["rel_27", "rel_28", "rel_29"]],
    ["AI 營收校準", ["ai25", "ai26"]],
    ["AI 毛利率", ["d_ai"]],
    ["市場量敏感度係數（只進敏感度）", ["mf", "wta_27", "wta_28", "wta_29"]]]],
  ["非 AI 產品與服務", [
    ["非 AI 產品營收成長", ["g_nai27", "g_nai28", "g_nai29"]],
    ["服務營收", ["svc25", "g_svc26", "g_svc27", "g_svc28", "g_svc29"]]]],
  ["全公司營運：FY2026 錨點與營業費用", [
    ["FY2026 錨點（第三季實績與第四季指引）", ["q3_rev", "q3_gp", "q3_oi", "q4_rev", "q4_gm", "q4_oim", "q_oth", "q_tx", "q_sh"]],
    ["營業費用、其他收益、稅與股數（FY2027 起）", ["phi", "opx_adj", "oth_f", "tx_f", "sh_g"]]]],
  ["全公司：股價、折現率、目標價規則與現金流", [
    ["股價與市場倍數", ["px", "pe_ttm_gaap", "pe_fwd_sa", "g2n"]],
    ["共識與先驗（只作對帳）", ["cons_rev26", "cons_eps26", "erste_eps27", "dell_top2", "idc_anet", "pe_fair_andy", "cons_tp"]],
    ["目標價規則（情境權重、DCF 權重、PL 延伸期收斂）", ["tp_wA", "tp_wB", "tp_wC", "tp_w_dcf", "tp_g", "tp_gd", "nm10_k"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "beta_u", "beta_u_med", "beta_du", "beta_de", "beta", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "bs_date", "stub_frac", "stub_div", "sh_v"]],
    ["自由現金流（股票薪酬、折舊、資本支出、營運資金）", ["fc_sbc", "fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc", "fcf_da", "fcf_capex"]]]]],
 "blocks_to": {"Arista 份額": "AI 後端乙太網路（AI 營收）", "@Units!A13": "AI 後端乙太網路（AI 營收）"},
 "main": {"DCF 彙總": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "淨現金滾到評價日", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
          "DCF 折現率": ["股權成本", "稅後債務成本", "權益占資本", "WACC", "永續成長率", "永續期 RONIC"],
          "對帳要點": ["FY2026 營收 vs 共識", "FY2026 EPS vs 共識", "FY2027 EPS vs 前瞻", "FY2026 Arista 有效份額 vs IDC", "3Q26 EPS"]},
}
# 第三輪修正 15／21：情境驅動表依 In 頁 C 欄單位格式化（% 與 pt 用百分比，避免 0.06 顯示成小數）；DCF 彙總主表直接顯示 stub 兩列
def _uf3(r_):
    u = IN.cell(r_, 3).value
    return "p1" if u in ("%", "pt") else "n3" if u == "十億股" else "n2"
for b in secs["drivers"]["blocks"]:
    if b["type"] == "table" and str(b.get("caption", "")).startswith("情境驅動"):
        b["rows"] = [[x[0]] + [(f"{c}|{_uf3(int(c.split('!')[1][1:]))}" if "|" not in c and isinstance(IN[c.split('!')[1]].value, (int, float)) else c) for c in x[1:4]] + x[4:] for x in b["rows"]]
_m3 = SPEC["main"]["DCF 彙總"]
if "＋Stub 期 FCFF" not in _m3:
    _m3.insert(_m3.index("＋淨現金與投資") + 1, "＋Stub 期 FCFF")
cfg = relayout(cfg, X, SPEC)
# 後處理（補充規則 A）：橫幅中間格＝只看 Comparable（市場主流同業法、未經相對倍數調整）；附註列進加權的調整後值
_rc, _ru, _ra = row(VA, "Comparable 目標價"), row(VA, "上檔空間：Comparable"), row(VA, "Comparable 目標價（相對倍數調整後）")
assert VA.cell(_rc, 1).value == "Comparable 目標價"
cfg["hero"]["tiles"][1] = ["只看 Comparable 目標價（市場主流同業法）", f"@Val!B{_rc}|n2", ["上檔空間 ", f"@Val!B{_ru}|sp1"]]
cfg["hero"]["note"] = list(cfg["hero"].get("note") or []) + ["。進加權的 Comparable 為相對倍數調整後 ", f"@Val!B{_ra}|n2", "（調整係數 ", f"@Val!B{_ra - 1}|n2", "）"]
# 後處理：轉換器把估值章所有長條圖移到最前面；舊法（規則值）長條圖放回舊法段落，避免與目標價並列
_vb = cfg["sections"][[s["id"] for s in cfg["sections"]].index("valuation")]["blocks"]
_old = next(b for b in _vb if b["type"] == "bars" and b.get("title") == "規則值隱含股價與現價")
_vb.remove(_old)
_h = next(i for i, b in enumerate(_vb) if b["type"] == "h3" and str(b.get("text", "")).startswith("規則值隱含股價"))
_vb.insert(_h + 2, _old)
# 後處理：單位為 pt（百分點）的參數，轉換器預設一位小數；改用百分比格式，避免 0.06 顯示成 0.1
_pt = {r for r in range(4, IN.max_row + 1) if IN.cell(r, 3).value == "pt"}
def _fix_pt(x):
    m = re.match(r"^@In!([DEF])(\d+)\|n1$", x) if isinstance(x, str) else None
    return f"@In!{m.group(1)}{m.group(2)}|p1" if m and int(m.group(2)) in _pt else x
for b in cfg["sections"][[s["id"] for s in cfg["sections"]].index("drivers")]["blocks"]:
    if b["type"] == "table" and isinstance(b.get("rows"), list):
        b["rows"] = [[_fix_pt(c) for c in r] if isinstance(r, list) else r for r in b["rows"]]
# 後處理：本模型的龍捲風圖是 FY2028 EPS（不是加權目標價），轉換器的通用標題改回原意
for b in cfg["sections"][[s["id"] for s in cfg["sections"]].index("sens")]["blocks"]:
    if b["type"] == "tornado":
        b["title"] = "FY2028 EPS 相對基準的變動（情境 A，一次動一個輸入；" + ("未達兩位數百分比）" if b.get("collapse") else "變動達兩位數百分比的案例）")
# 第三輪修正 17：橫幅金額標明幣別
_CUR3 = "美元"
def _lab3(x):
    t = str(wb[x[1:].split("!")[0]][x.split("!")[1]].value) if isinstance(x, str) and x.startswith("@") else str(x)
    return x if _CUR3 in t else ([x, "（" + _CUR3 + "）"] if isinstance(x, str) and x.startswith("@") else t + "（" + _CUR3 + "）")
for _t in cfg["hero"]["tiles"]:
    _t[0] = _lab3(_t[0])
# ================= v0.5.5（2026-11-15）：個股 HTML 版面規格 v1.3（樣板 NVIDIA v0.3.10；同 lineage 智邦 v0.4.5）＋三層協調 C0 =================
# 章節：命題 → 估值 → 敏感度 → 反向 DCF → 損益（標準損益表）→ 驅動因子（分部營收推導）→ 對帳（含 C0）→ 資料狀態；Arista 無個別模組。
# 數字一律取自 xlsx（README 頁末 v1.3 白話列、PL 頁末標準損益表、RevDCF 頁、Val 0.7、Sens 頁末、Coord 頁 C0；arista/v055_post.py 產生）。列號依標籤找。
sys.path.insert(0, os.path.join(HERE, "..", "tools", "coord"))
import coord as COORD  # noqa: E402

WBF = wb


def find(sh, prefix, start=1, col=1, exact=False):
    ws = WBF[sh]
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and ((v.strip() == prefix) if exact else v.strip().startswith(prefix)):
            return r
    raise KeyError((sh, prefix))


S = {s_["id"]: s_ for s_ in cfg["sections"]}
RD, PLW, RV = WBF["README"], WBF["PL"], WBF["RevDCF"]
r13 = find("README", "v1.3 展示層")
RM = dict(zip(["hero", "rec", "o1", "o2", "o3", "o4", "o5", "drv", "ver"], range(r13 + 1, r13 + 10)))
assert str(RD[f"A{RM['ver']}"].value).startswith("v0.5.6"), RD[f"A{RM['ver']}"].value
cfg["layout"] = "v1.3"
old_note = cfg["hero"]["note"]
cfg["hero"]["note"] = [f"@README!A{RM['hero']}"]            # 橫幅只留 ≤200 字白話；評等規則與相對倍數調整說明移到估值章

# ---- 1. 命題與核心問題 ----
rbB = find("PL", "情境 B：")
prB = lambda p, ex=True: find("PL", p, rbB, exact=ex)
hdB = rbB + 1
r_aisw, r_nap, r_epsB = prB("AI 營收"), prB("非 AI 營收"), prB("EPS（non-GAAP）")
core_chart = {"type": "theme", "title": "AI 份額下滑撐不撐得住 EPS：AI 營收、非 AI 營收（長條，十億美元）與 non-GAAP EPS（折線，美元，右軸；情境 B）",
              "x": f"@PL!C{hdB}:G{hdB}", "fmt": "n1", "fmt_r": "n2", "stack": True, "pad": 1.25,
              "bars": [{"name": ["AI 營收"], "y": f"@PL!C{r_aisw}:G{r_aisw}", "color": "#0b7a80"},
                       {"name": ["非 AI 營收"], "y": f"@PL!C{r_nap}:G{r_nap}", "color": "#6fc7cb"}],
              "lines": [{"name": ["EPS（non-GAAP，美元）"], "y": f"@PL!C{r_epsB}:G{r_epsB}", "axis": "r", "color": "#3b3838", "labels": True}],
              "caption": "回答核心問題：AI 營收＝交換器市場 × 乙太網路占比 × Arista 份額；份額相對值逐年下修，情境 B 的顆數自 FY2028 起下修（顆數係數縮減），非 AI 營收以成長率推。FY2025 為實績、FY2026 為錨點，FY2027 起為對帳中年份。"}
th = S["thesis"]["blocks"]
i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RM[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
th.insert(0, {"type": "note", "parts": [f"@README!A{RM['rec']}"]})
th.insert(0, core_chart)

# ---- 2. 估值：四個目標價與現價圖置頂（G 節位置）；原橫幅評等規則與相對倍數調整說明放在圖下 ----
va = S["valuation"]["blocks"]
i_bar = next(i for i, b in enumerate(va) if b["type"] == "bars" and b.get("title") == "目標價與現價")
g0 = find("Val", "0.7 四個目標價並列"); gr = list(range(g0 + 2, g0 + 7)); r_px = find("Val", "現價（", g0)
bars4 = {"type": "bars", "title": "四個目標價與現價（美元；豎線＝現價）", "fmt": "n0",
         "items": [{"label": lb, "value": f"@Val!B{r}", "series": sr} for r, sr, lb in zip(gr, (1, 2, 2, 0, 1),
                   ("DCF 目標價", "Comparable（調整後，進加權）", "Comparable（未調整，不進加權）", "加權目標價（橫幅主值）", "市場共識目標價"))],
         "ref": {"label": f"@Val!A{r_px}", "value": f"@Val!B{r_px}"}}
tbl4 = {"type": "table", "caption": f"@Val!A{g0}", "head": [f"@Val!{c}{g0 + 1}" for c in "ABCD"],
        "rows": [[f"@Val!A{r}", f"@Val!B{r}|n2", f"@Val!C{r}|sp1", f"@Val!D{r}"] for r in gr] + [[f"@Val!A{g0 + 7}", f"@Val!B{g0 + 7}|n2", "", f"@Val!D{g0 + 7}"]]}
note_rule = {"type": "note", "parts": list(old_note)}
conc = va[i_bar + 1]
conc = dict(conc, collapse=True, summary="展開：結論的主要項目（四個目標價之外的上檔空間、投資建議、市場共識等）")
va[i_bar + 1] = conc
del va[i_bar]
_vi = S["valuation"].pop("intro", None)
va[0:0] = [bars4, {"type": "note", "parts": [_vi]} if _vi else {"type": "note", "parts": [""]}, tbl4, note_rule]

# ---- 4. 反向 DCF ----
rv = lambda p, start=1: next(r for r in range(start, RV.max_row + 1) if str(RV.cell(r, 1).value or "").startswith(p))
ra = rv("A. DCF"); rb_ = rv("(a)"); rbb = rv("(b)"); rc = rv("(c)"); rcmp = rv("B. Comparable"); rsum = rv("C. 摘要"); yh = rv("項目（情境 B）")
m1, m0 = rv("現價隱含的營業利益率", rb_), rv("模型：基準的等效營業利益率", rb_)
s1, rv1, rv0, g1, g0_ = rv("現價隱含的營收倍數", rbb), rv("現價隱含的 FY2028 營收", rbb), rv("模型：FY2028 營收", rbb), rv("現價隱含的 FY2028 營收年增率", rbb), rv("模型：FY2028 營收年增率", rbb)
w1, w0 = rv("現價隱含的 WACC", rc), rv("模型 WACC", rc)
rev_ = {
    "id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2",
    "blocks": [
        {"type": "note", "parts": [f"@RevDCF!A{rsum + 1}"]},
        {"type": "note", "parts": [f"@RevDCF!A{rsum + 2}"]},
        {"type": "table", "caption": "現金流折現法（情境 B 路徑）：現價隱含值與模型基準", "head": ["項目", "現價隱含", "模型基準"],
         "rows": [["預測期營業利益率（FY2028–FY2032 統一一個值；模型欄為等效值，終值年權重大）", f"@RevDCF!B{m1}|p1", f"@RevDCF!B{m0}|p1"],
                  ["營收規模（FY2028 起各年同乘的倍數）", f"@RevDCF!B{s1}|n2", "—"],
                  ["FY2028 營收（十億元）", f"@RevDCF!B{rv1}|n0", f"@RevDCF!B{rv0}|n0"],
                  ["FY2028 營收年增率", f"@RevDCF!B{g1}|p1", f"@RevDCF!B{g0_}|p1"],
                  ["折現率（WACC）", f"@RevDCF!B{w1}|p2", f"@RevDCF!B{w0}|p2"]]},
        {"type": "table", "caption": f"@RevDCF!A{ra}", "head": ["項目", "值"],
         "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("n2" if r in (ra + 2, ra + 3, ra + 9) else "n3" if r == ra + 5 else "x1" if r == ra + 8 else "n1")] for r in range(ra + 2, ra + 10)]},
        {"type": "table", "caption": f"@RevDCF!A{rcmp}", "head": [f"@RevDCF!{c}{rcmp + 1}" for c in "ABCD"],
         "rows": [[f"@RevDCF!A{rcmp + 2}", "—", f"@RevDCF!C{rcmp + 2}|n2", "—"]]
                 + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("n2" if k == 1 else "n1"), f"@RevDCF!C{r}|" + ("n2" if k == 1 else "n1"), f"@RevDCF!D{r}|sp1"] for k, r in enumerate(range(rcmp + 3, rcmp + 6))]
                 + [[f"@RevDCF!A{rcmp + 6}", "—", f"@RevDCF!C{rcmp + 6}|n1", "—"]]},
        {"type": "table", "caption": f"@RevDCF!A{yh}", "head": [f"@RevDCF!{c}{yh}" for c in "ABCDEFG"], "collapse": True,
         "summary": "展開：情境 B 各年數字與反解係數（封閉解的計算過程）",
         "rows": [[f"@RevDCF!A{r}"] + [f"@RevDCF!{c}{r}|" + ("p1" if r in (yh + 3, yh + 4) else "n3" if r in (yh + 5, yh + 6, yh + 7, yh + 8) else "n1") for c in "BCDEFG"] for r in range(yh + 1, rb_ - 1)]
                 + [{"sub": f"@RevDCF!A{rb_}"}] + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("p1" if "利益率" in str(RV.cell(r, 1).value) else "n1"), "", "", "", "", ""] for r in range(rb_ + 1, rb_ + 5)]
                 + [{"sub": f"@RevDCF!A{rbb}"}] + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("p1" if "年增率" in str(RV.cell(r, 1).value) else "n2" if "倍數" in str(RV.cell(r, 1).value) else "n3" if "核對" in str(RV.cell(r, 1).value) else "n1"), "", "", "", "", ""] for r in range(rbb + 1, rc - 1)]
                 + [{"sub": f"@RevDCF!A{rc}"}] + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("p2" if str(RV.cell(r, 1).value).startswith(("現價隱含的 WACC", "模型 WACC")) else "n2"), "", "", "", "", ""] for r in range(w1, w1 + 4)]},
    ]}

# ---- 5. 損益（標準損益表） ----
p0 = find("PL", "標準損益表（v0.5.5")
ph = next(r for r in range(p0, PLW.max_row + 1) if PLW.cell(r, 1).value == "項目")
plr = lambda lab_, start=ph: next(r for r in range(start, PLW.max_row + 1) if PLW.cell(r, 1).value == lab_)
keys = ["營收", "營收年增率（YoY）", "營業成本（COGS）", "毛利", "毛利率", "營業費用（銷管＋研發）", "營業利益", "營業利益率", "其他收入（淨額：利息、匯兌、投資收益等）",
        "稅前利益", "所得稅", "少數股權損益", "稅後淨利（歸屬母公司）", "淨利率", "股數（十億股）", "EPS（non-GAAP 稀釋；美元）"]
fmt = lambda u: {"%": "p1", "US$": "n2", "十億股": "n3"}.get(u, "n2")
PC9 = "CDEFGHIJK"
rowx = lambda r: [f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmt(PLW.cell(r, 2).value)}" for c in PC9]
r_ratio = plr("比率"); r_dil = plr("GAAP 稀釋 EPS 對照（歷史；預測不估 GAAP）"); r_chk = plr("核對（應為 0）")
r_d31 = next(r for r in range(r_chk, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("差額（FY2030E、FY2031E）：毛利"))
is_tbl = {"type": "table", "caption": "標準損益表（Non-GAAP 口徑，十億美元；EPS 為美元）：FY2023A–FY2025A 實績、FY2026E–FY2031E 情境 B（基準）；標「對帳中」的年度依賴 Siliconomics 對帳中年份",
          "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in PC9],
          "rows": [rowx(plr(k)) for k in keys] + [{"sub": f"@PL!A{r_ratio}"}] + [rowx(r) for r in range(r_ratio + 1, r_dil)]
                  + [{"sub": "核對：FY2030E、FY2031E 兩種推法的營業利益差額（毛利率與營業費用率沿用前一年）"}] + [rowx(r) for r in (r_d31, r_d31 + 1)]}
dil_tbl = {"type": "table", "caption": f"@PL!A{r_dil}", "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in "CDE"], "collapse": True,
           "summary": "展開：GAAP 稀釋 EPS 對照（歷史）與核對列",
           "rows": [[f"@PL!A{r_dil + 1}"] + [f"@PL!{c}{r_dil + 1}|n2" for c in "CDE"]] + [{"sub": f"@PL!A{r_chk}"}]
                   + [[f"@PL!A{r}"] + [f"@PL!{c}{r}|n3" for c in "CDE"] for r in range(r_chk + 1, PLW.max_row + 1) if r not in (r_d31, r_d31 + 1)]}
rv_, gm_, opm_, nm_ = plr("營收"), plr("毛利率"), plr("營業利益率"), plr("淨利率")
is_chart = {"type": "theme", "title": "營收（十億美元）與毛利率、營業利益率、淨利率：FY2023A–FY2031E（情境 B）", "x": f"@PL!C{ph}:K{ph}",
            "fmt": "n1", "fmt_r": "p1", "stack": True, "pad": 1.6,
            "bars": [{"name": [f"@PL!A{rv_}"], "y": f"@PL!C{rv_}:K{rv_}", "color": "#6fc7cb"}],
            "lines": [{"name": [f"@PL!A{gm_}"], "y": f"@PL!C{gm_}:K{gm_}", "axis": "r", "color": "#0b7a80", "labels": True},
                      {"name": [f"@PL!A{opm_}"], "y": f"@PL!C{opm_}:K{opm_}", "axis": "r", "color": "#c29a2b"},
                      {"name": [f"@PL!A{nm_}"], "y": f"@PL!C{nm_}:K{nm_}", "axis": "r", "color": "#6f6bae", "dash": True}],
            "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸，non-GAAP 口徑）。FY2026E 為錨點；FY2027E 起為對帳中年份（Siliconomics 標示）；FY2030E–FY2031E 為延伸期：營收與淨利由收斂法推得，毛利率與營業費用率沿用前一年。"}
xt = find("PL", "三情境延伸到 FY2032"); xh = find("PL", "情境／延伸項目", xt)
xs = lambda s: find("PL", "情境 " + s, xh, exact=True)
xr = lambda base, lab_: next(r for r in range(base, PLW.max_row + 1) if PLW.cell(r, 1).value == lab_)
R_REV, R_EPS = "營收", "EPS（延伸表）"
xA, xB, xC = xs("A"), xs("B"), xs("C")
sc_chart = {"type": "theme", "title": "三情境並列：營收（長條，十億美元）與 non-GAAP EPS（折線，美元，右軸）：FY2025A–FY2032", "x": f"@PL!C{xh}:J{xh}",
            "fmt": "n1", "fmt_r": "n2", "stack": False, "pad": 1.1,
            "bars": [{"name": ["情境 A（樂觀）營收"], "y": f"@PL!C{xr(xA, R_REV)}:J{xr(xA, R_REV)}", "color": "#b7e3e4"},
                     {"name": ["情境 B（基準）營收"], "y": f"@PL!C{xr(xB, R_REV)}:J{xr(xB, R_REV)}", "color": "#10a0a7"},
                     {"name": ["情境 C（保守）營收"], "y": f"@PL!C{xr(xC, R_REV)}:J{xr(xC, R_REV)}", "color": "#8a8c8c"}],
            "lines": [{"name": ["情境 A EPS"], "y": f"@PL!C{xr(xA, R_EPS)}:J{xr(xA, R_EPS)}", "axis": "r", "color": "#0b7a80", "dash": True},
                      {"name": ["情境 B EPS"], "y": f"@PL!C{xr(xB, R_EPS)}:J{xr(xB, R_EPS)}", "axis": "r", "color": "#3b3838", "labels": True},
                      {"name": ["情境 C EPS"], "y": f"@PL!C{xr(xC, R_EPS)}:J{xr(xC, R_EPS)}", "axis": "r", "color": "#b4472f", "dash": True}],
            "caption": "三情境在 FY2026–FY2027 相同（錨點；顆數係數相同），差異自 FY2028 起（顆數係數不同）；FY2030 起為延伸期（收斂法）。FY2027 起皆為對帳中年份。"}
old_pl = S["pl"]["blocks"]
pls = [b for b in old_pl if b["type"] == "pl"]
notes = [p_ for b in old_pl if b["type"] == "note" for p_ in b["parts"]]
kvs = [b for b in old_pl if b["type"] == "kv"]
assert len(pls) == 2 and len(kvs) == 1, (len(pls), len(kvs))
old_blocks = [dict(pls[0], collapse=True, summary="展開：三情境摘要表與折線（FY2025–FY2029；營收、AI 營收、非 AI 營收、毛利率、EPS）"), dict(kvs[0], collapse=True, summary="展開：三情境錨點與驅動核對"),
              {"type": "note", "parts": notes, "collapse": True, "summary": "展開：延伸期（FY2030–FY2032）的收斂法說明"},
              dict(pls[1], collapse=True, summary="展開：三情境延伸到 FY2032（收斂法）")]
S["pl"].pop("intro", None)
S["pl"].update(title="損益（標準損益表）", blocks=[{"type": "note", "parts": [f"@PL!A{p0 + 1}"]}, is_chart, is_tbl, {"type": "note", "parts": [f"@PL!A{p0 + 2}"]}, sc_chart, dil_tbl] + old_blocks)

# ---- 6. 驅動因子：分部營收推導（情境 B） ----
pe = lambda p: find("PL", p, rbB)
rows_seg = [("AI 營收＝AI 後端 scale-out 交換器市場 × 乙太網路占比 × Arista 有效份額", [pe("AI 後端 scale-out 交換器市場"), prB("乙太網路占比"), prB("AI 後端乙太網路交換器市場"), pe("Arista 有效份額（FY2026 校準）"), prB("份額相對值（÷ FY2026）"), prB("Arista 有效份額"), prB("AI 營收")]),
            ("非 AI 營收＝FY2026 錨點 × 成長率（其中服務營收另以成長率推）", [prB("非 AI 營收"), prB("服務營收"), pe("產品：非 AI")]),
            ("合計：合併營收（AI＋非 AI）", [prB("營收")])]
fm_seg = lambda u: {"US$B": "n2", "%": "p1", "x": "n2"}.get(u, "n2")
seg_rows = []
for t_, rs in rows_seg:
    seg_rows.append({"sub": t_})
    seg_rows += [[f"@PL!A{r}", f"@PL!B{r}"] + [f"@PL!{c}{r}|{fm_seg(PLW.cell(r, 2).value)}" for c in "DEFG"] for r in rs]
seg_tbl = {"type": "table", "caption": "分部營收推導（情境 B；FY2026 為錨點）：驅動因子 → 各分部營收 → 合併營收", "head": [f"@PL!A{hdB}", f"@PL!B{hdB}"] + [f"@PL!{c}{hdB}" for c in "DEFG"], "rows": seg_rows}
r_tot = prB("營收")
seg_chart = {"type": "theme", "title": "分部營收（情境 B，十億美元）：AI 營收、非 AI 營收", "x": f"@PL!C{hdB}:G{hdB}", "fmt": "n1", "fmt_r": "n1", "stack": True, "pad": 1.15, "total": f"@PL!C{r_tot}:G{r_tot}",
             "bars": [{"name": ["AI 營收"], "y": f"@PL!C{r_aisw}:G{r_aisw}", "color": "#0b7a80"}, {"name": ["非 AI 營收（含服務）"], "y": f"@PL!C{r_nap}:G{r_nap}", "color": "#6fc7cb"}],
             "lines": [], "caption": "堆疊＝兩個分部營收（FY2025A 的 AI 營收為公司稱超過目標的推算值）；頂端數字＝合併營收（取 PL 頁合計列，不在產生器加總）。AI 營收由市場、乙太網路占比與份額推導，見下表。"}
S["drivers"]["blocks"] = [{"type": "note", "parts": [f"@README!A{RM['drv']}"]}, seg_chart, seg_tbl, {"type": "h3", "text": "情境驅動與單一參數"}] + S["drivers"]["blocks"]
S["drivers"]["title"] = "驅動因子：分部營收推導"

# ---- 8. 對帳：C0（實績）＋讀取層級（同源，不判定） ----
if "Coord" in WBF.sheetnames:
    CW = WBF["Coord"]
    cr = lambda p, start=1: next(r for r in range(start, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith(p))
    c0 = cr("C0 由下往上對帳"); c1 = cr("1. 讀哪一層", c0); c2 = cr("2. 讀取層級", c0); c3 = cr("3. 結論", c0); r_cc = cr("C0 結論（白話）", c3)
    c4 = cr("4. C0（實績）", c3); r_m = cr("4.1 公司實績", c4); r_p = cr("4.2 口徑換算", c4); r_cmp = cr("4.3 對帳", c4)
    r_meth = cr("方法：", c4); r_rem = cr("為什麼不判定", c4); r_end = CW.max_row
    nfmt = lambda r: next((CW[f"{c}{r}"].number_format for c in "BCDE" if isinstance(CW[f"{c}{r}"].value, (int, float)) or (isinstance(CW[f"{c}{r}"].value, str) and CW[f"{c}{r}"].value.startswith("=") and c in "BC")), "General")
    fcode = lambda a, nf: "" if ("判定：" in a or "是否都在" in a) else ("sp1" if nf.startswith("+") else "p1" if "%" in nf else "n3" if nf.startswith("0.000") else "n2" if nf in ("0.00", "#,##0.00") else "n1")
    lay_rows = [[f"@Coord!{c}{r}" for c in "ABCDE"] for r in range(c1 + 2, c2 - 1) if CW.cell(r, 1).value]
    c0_rows = []
    for r in range(c2 + 2, c3 - 1):
        a = str(CW.cell(r, 1).value or "")
        if not a:
            continue
        if not a.startswith("　"):
            c0_rows.append({"sub": f"@Coord!A{r}"})
            continue
        nf = next((CW[f"{c}{r}"].number_format for c in "BCDE" if CW[f"{c}{r}"].value is not None), "General")
        f_ = "" if "判定" in a else ("sp1" if nf.startswith("+") else "p1" if "%" in nf else "n4" if nf == "0.0000" else "n2" if nf == "0.00" else "n1")
        cx = lambda c: f"@Coord!{c}{r}" if CW[f"{c}{r}"].value else ""
        c0_rows.append([f"@Coord!A{r}"] + [f"@Coord!{c}{r}|{f_}" for c in "BCDE"] + [cx("F"), cx("G")])
    act_rows = []
    for r in range(r_cmp + 1, r_end + 1):
        a = str(CW.cell(r, 1).value or "")
        if not a:
            continue
        f_ = fcode(a, nfmt(r))
        act_rows.append([f"@Coord!A{r}", f"@Coord!B{r}|{f_}", (f"@Coord!C{r}|{f_}" if CW[f"C{r}"].value is not None and not (isinstance(CW[f"C{r}"].value, str) and not CW[f"C{r}"].value.startswith("=")) else ""),
                         (f"@Coord!D{r}" if CW[f"D{r}"].value else ""), (f"@Coord!E{r}" if CW[f"E{r}"].value else "")])
    in_rows = []
    for r in range(r_m + 1, r_cmp):
        a = str(CW.cell(r, 1).value or "")
        if not a:
            continue
        if not a.startswith("　"):
            in_rows.append({"sub": f"@Coord!A{r}"})
            continue
        f_ = fcode(a, nfmt(r))
        if "核對" in a:
            f_ = "n3"
        in_rows.append([f"@Coord!A{r}", f"@Coord!B{r}|{f_}", (f"@Coord!D{r}" if CW[f"D{r}"].value else ""), (f"@Coord!E{r}" if CW[f"E{r}"].value else "")])
    c0_blocks = [{"type": "h3", "text": "三層協調：公司實績對照（無可比實績，不判定）"},
                 {"type": "note", "parts": [f"@Coord!B{r_cc}"]},
                 {"type": "note", "parts": [f"@Coord!A{r_meth}"]},
                 {"type": "table", "caption": "公司實績對照（參考，不判定：Arista 份額由公司 AI 營收目標反推，無可比實績）：2026 年已公布季營收對層級隱含值與錨點的差（金額十億美元）",
                  "head": ["項目", "主讀數（已公布季度累計）", "第二讀數（最近一季年化）", "單位與口徑", "來源與證據等級"], "rows": act_rows},
                 {"type": "note", "parts": [f"@Coord!A{r_rem}"]},
                 {"type": "table", "caption": "實績輸入（逐季營收、核對列）與參考值", "head": ["項目", "值", "單位與口徑", "來源與證據等級"], "rows": in_rows,
                  "collapse": True, "summary": "展開：逐季營收實績、核對列與出處"},
                 {"type": "table", "caption": "讀取層級（同源，不判定）：本模型自層級讀入的值與層級值必然相同，差距恆為零，無檢驗力，不計入對帳中（情境 B 基準對 Siliconomics 基準）", "head": ["項目", "CY2026", "CY2027", "CY2028", "CY2029", "單位與口徑", "來源與證據等級"], "rows": c0_rows,
                  "collapse": True, "summary": "展開：讀取層級（同源，不判定）與層級之間的參考差距"},
                 {"type": "table", "caption": "讀哪一層、哪一版（下游只註明版本、不自行重估）", "head": ["項目", "值", "版本／檔案", "本模型用途", "備註"], "rows": lay_rows,
                  "collapse": True, "summary": "展開：本模型讀哪一層、哪一版（名稱、值、版本）"},
                 {"type": "h3", "text": "共識、研調份額與校準的對照"}]
    S["recon"]["blocks"] = c0_blocks + S["recon"]["blocks"]

# ---- 9. 資料狀態：本版要點、歷史資料來源 ----
st = S["status"]["blocks"]
r_h1 = find("Hist", "Hist — Arista損益表歷史")
st[0:0] = [{"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RM['ver']}"]},
           {"type": "h3", "text": "標準損益表歷史資料來源"}, {"type": "note", "parts": [f"@Hist!A{r_h1}"]}]

# ---- 3. 敏感度：加權目標價相對基準的低案、高案主圖置頂（H 節位置） ----
SEW = WBF["Sens"]
r_tp = find("Sens", "目標價參數敏感度")
t0 = find("Sens", "加權目標價相對基準（HTML"); r1 = t0 + 2; r2 = r1
while SEW.cell(r2 + 1, 1).value:
    r2 += 1
tors = []
for col, side in (("B", "低"), ("C", "高")):
    vals = {r: SEW[f"{col}{r}"].value for r in range(r1, r2 + 1) if isinstance(SEW[f"{col}{r}"].value, (int, float))}
    big = [r for r, v in vals.items() if abs(v) >= 0.1]
    small = [r for r, v in vals.items() if 0.0005 <= abs(v) < 0.1]
    base = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1"}
    tors.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值，變動達兩位數百分比的案例", rows=",".join(map(str, big))))
    tors.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值", rows=",".join(map(str, small)), collapse=True,
                     summary=f"展開：加權目標價相對基準：輸入取{side}值，變動未達兩位數百分比的案例（幾乎為零者見全部案例表）"))
    assert big, (side, vals)
tors.append({"type": "table", "caption": f"@Sens!A{t0}", "collapse": True, "summary": "展開：全部案例的加權目標價相對基準（低案、高案）",
             "head": [f"@Sens!{c}{t0 + 1}" for c in "ABC"], "rows": [[f"@Sens!A{r}", f"@Sens!B{r}|sp1", f"@Sens!C{r}|sp1"] for r in range(r1, r2 + 1)]})
tb0 = find("Sens", "目標價敏感度：業務主因")
n_biz = 0
while SEW.cell(tb0 + 2 + n_biz, 1).value and isinstance(SEW.cell(tb0 + 2 + n_biz, 3).value, (int, float)):
    n_biz += 1
tors.append({"type": "table", "caption": f"@Sens!A{tb0}", "collapse": True, "summary": "展開：業務主因案例的加權目標價（整本重算快照；含快照狀態）",
             "head": [f"@Sens!{c}{tb0 + 1}" for c in "ACDEFGH"],
             "rows": [[f"@Sens!A{r}", f"@Sens!C{r}|n2", f"@Sens!D{r}|sp1", f"@Sens!E{r}|n2", f"@Sens!F{r}|sp1", f"@Sens!G{r}|n2", f"@Sens!H{r}"] for r in range(tb0 + 2, tb0 + 2 + n_biz)]})
sb = S["sens"]["blocks"]
for b in sb:
    if b.get("type") == "tornado":
        b["summary"] = "展開：對照：FY2028 EPS 相對基準的變動（情境 A，一次動一個輸入）" if not b.get("collapse") else b.get("summary", "展開：FY2028 EPS 相對基準的變動（未達兩位數百分比）")
        b["collapse"] = True
    if b.get("type") == "table" and not b.get("collapse") and str(b.get("caption", "")) == f"@Sens!A{r_tp}":
        b["collapse"] = True; b["summary"] = "展開：目標價參數表（WACC、β、ERP、永續成長率等；加權目標價與相對基準）"
_it = {k: S[k].pop("intro", None) for k in ("sens",)}
tors = [tors[0], tors[2], tors[1], tors[3]] + tors[4:]
S["sens"]["blocks"] = tors[:2] + \
    [{"type": "note", "parts": ["主圖為加權目標價相對基準的變動（低案＝輸入取敏感度低值、高案＝取高值；一次動一個輸入，其他不變）：業務主因為整本重算快照，目標價參數為閉式公式；變動達兩位數百分比者直接顯示，其餘摺疊。下方摺疊為全部案例與 FY2028 EPS 對照。"]}] + tors[2:] + sb

# ---- 順序與對帳中兩讀數（tools/coord/coord.py）：橫幅一格與揭露句、估值章兩讀數表與長條圖放在四個目標價圖與表之後 ----
cfg["sections"] = [S[k] for k in ("thesis", "valuation", "sens")] + [rev_] + [S[k] for k in ("pl", "drivers", "recon", "status")]
n_before = len(S["valuation"]["blocks"])
COORD.inject_html(cfg, X)
va = S["valuation"]["blocks"]
if len(va) > n_before:                                           # inject_html 在最前面插入 4 個區塊（h3、note、table、bars）
    inj, rest = va[:4], va[4:]
    assert inj[0].get("type") == "h3" and inj[3].get("type") == "bars", inj
    S["valuation"]["blocks"] = rest[:3] + inj + rest[3:]
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")
