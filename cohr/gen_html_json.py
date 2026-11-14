# v0.3.1（2026-11-05）：套用個股 HTML 版面規格 v1.0（docs/plan/個股HTML版面規格.md）——只改展示層。v0.4：DCF 改標準版（FCFF＋WACC），估值章 DCF 區塊改由本檔依 xlsx 列號產生；html_src.json 只重對 In 列號。
# 讀 cohr/html_src.json（v0.3 的版面，內容不動）→ 少量前處理 → tools/company_html/relayout.py → 寫 cohr/html.json。
# 用法：python3 cohr/gen_html_json.py [成品 xlsx（預設 cohr/dist/ 內唯一的 xlsx）]
import glob, json, os, re, sys
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
r_dtp = row(VA, "DCF 目標價（情境加權）")
_rk = row(VA, "股權成本 ke"); _rr = row(VA, "永續期 RONIC"); _rp = row(VA, "明確期 FCFF 現值")
_dh = row(DC, "項目")
# 第二輪修正：DCF 彙總表格式依標籤（模組 v2 stub 期在「＋淨現金與投資」後多兩列，不寫死列偏移）
_fm = lambda r: (lambda lab: "n3" if lab.startswith("評價用稀釋股數") else "p1" if "終值現值占 EV" in lab else "x1" if "隱含終值" in lab else "n2")(str(VA.cell(r, 1).value or "").strip())
DCF_BLOCKS = [
    {"type": "note", "parts": [f"@Val!A{_rk - 2}"]},
    {"type": "table", "caption": "DCF 折現率：股權成本、稅後債務成本、WACC、永續成長率、RONIC", "head": ["項目", "數值", "說明"],
     "rows": [[f"@Val!A{r}", f"@Val!B{r}|" + ("p1" if r == _rk + 2 else "p2"), f"@Val!E{r}"] for r in range(_rk, _rr + 1)]},
    {"type": "table", "caption": "DCF 彙總：明確期現值、終值現值、企業價值、淨負債與每股價值", "head": [f"@Val!{c}{_rk - 1}" for c in "ABCD"],
     "rows": [[f"@Val!A{r}"] + [f"@Val!{c}{r}|" + _fm(r) for c in "BCD"] for r in range(_rp, r_dtp)]
             + [[f"@Val!A{r_dtp}", f"@Val!B{r_dtp}|n2", "", ""]]}]
for s_ in "ABC":
    _a = row(DC, f"情境 {s_}（"); _b = next(r for r in range(_a, DC.max_row + 1) if str(DC.cell(r, 1).value or "").startswith("現值（FCFF"))
    _cols = ["A"] + [chr(ord("C") + i) for i in range(6)]
    _pct = {str(r): {c: "p1" for c in _cols[1:]} for r in range(_a + 1, _b + 1) if str(DC.cell(r, 1).value).startswith(("稅率", "折舊攤銷 ÷", "資本支出 ÷"))}
    _pct.update({str(r): {c: "n3" for c in _cols[1:]} for r in range(_a + 1, _b + 1) if str(DC.cell(r, 1).value).startswith("折現因子")})
    DCF_BLOCKS.append({"type": "table", "caption": f"@DCF!A{_a}", "collapse": True, "summary": "展開：本情境各年無槓桿自由現金流明細",
                       "head": [f"@DCF!{c}{_dh}" for c in _cols],
                       "rowgen": {"sheet": "DCF", "rows": f"{_a + 1}-{_b}", "cols": _cols, "fmts": [""] + ["n2"] * 6, "row_fmt": _pct}})
# v0.4：Sens 目標價參數表改為標準版 DCF 的十個案例（tools/fcff_dcf.py write_sens；起始列不變）
SE = wb["Sens"]
_st = row(SE, "目標價參數敏感度")
_se = next(r for r in range(_st + 2, SE.max_row + 1) if str(SE.cell(r, 1).value or "").startswith("檢查：閉式公式"))
for b in secs["sens"]["blocks"]:
    if b["type"] == "table" and b.get("caption") == f"@Sens!A{_st}":
        b["head"] = [f"@Sens!{c}{_st + 1}" for c in "ABCDEFG"]
        b["rowgen"] = {"sheet": "Sens", "rows": f"{_st + 2}-{_se - 1}", "cols": list("ABCDEFG"), "fmts": ["", "", "n2", "sp1", "n2", "sp1", "n2"]}

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
# 2. 估值：重點觀察已移到第一章；兩個 kv（相對倍數調整、Comparable 計算）改為表格（轉換器會把 kv 當成折現率表）；
#    DCF 折現率直接顯示、緊接在 DCF 計算表之前
va = secs["valuation"]["blocks"]
out, skip = [], False
r_rm, r_cmp, r_dcf = row(VA, "0.5 相對倍數調整"), row(VA, "同業前瞻 P/E 中位數"), row(VA, "0.1 DCF")
for b in va:
    if b["type"] == "h3" and b.get("text") == "重點觀察":
        skip = True
        continue
    if skip and b["type"] == "list":
        skip = False
        continue
    if b["type"] == "kv":
        first = b["items"][0][0]
        cap = "Comparable 計算" if first == f"@Val!A{r_cmp}" else "相對倍數調整"
        b = {"type": "table", "caption": cap, "head": ["項目", "數值"], "rows": [list(i) for i in b["items"]]}
    if b["type"] == "table" and b.get("caption") == f"@Val!A{r_dcf}":   # v0.4：舊 0.1 DCF（股東現金流）表 → 標準版 DCF 折現率、彙總與各情境 FCFF 明細
        out += DCF_BLOCKS
        continue
    if b["type"] == "bars":
        b = json.loads(json.dumps(b))
        for it in b["items"]:
            if it["label"] == "@Val!A24":
                it["label"], it["value"] = f"@Val!A{r_dtp}", f"@Val!B{r_dtp}"
    out.append(b)
secs["valuation"]["blocks"] = out

scen_tok = re.split(r"[.\s]", IN.cell(row(IN, "情境驅動對照"), 1).value)[0]
SPEC = {
 "in_skip": ["H", scen_tok],
 "in_groups": [
  ["資料中心（AI 光模組：可插拔與 CPO）", [
    ["AI 光模組市場（讀網通市場模型；情境 A）", ["mpl_26", "mpl_27", "mpl_28", "mpl_29", "mcp_26", "mcp_27", "mcp_28", "mcp_29"]],
    ["網通 CPO 情境 B、C 的市場", ["mplB_26", "mplB_27", "mplB_28", "mplB_29", "mplC_26", "mplC_27", "mplC_28", "mplC_29",
                                  "mcpB_26", "mcpB_27", "mcpB_28", "mcpB_29", "mcpC_26", "mcpC_27", "mcpC_28", "mcpC_29"]],
    ["情境顆數係數", ["um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C"]],
    ["Coherent 份額、CPO 份額與單價", ["rel_27", "rel_28", "rel_29", "cpo_27", "cpo_28", "cpo_29", "pxr_27", "pxr_28", "pxr_29"]],
    ["會計年度換算與 FY2027 第二季", ["hw", "q2_dg"]],
    ["校準：FY2026 資料中心與通訊拆分（公司只揭露合計）", ["q3_dcc", "q4_dcc", "q4_dcc25", "q4_dg", "q4_cg", "q4_dy", "q4_cy", "h1r26", "q1_dcc"]],
    ["對帳參考（AI IT 資本支出、中際旭創份額）", ["itc_26", "itc_27", "itc_28", "itc_29", "inl_26", "inl_27", "inl_28", "inl_29"]],
    ["市場量敏感度係數（只進敏感度）", ["mf", "wta_27", "wta_28", "wta_29"]]]],
  ["通訊與工業", [
    ["通訊營收成長", ["q_cg", "g_com_27", "g_com_28", "g_com_29"]],
    ["工業營收（已排除出售事業）", ["q3_ind", "q4_ind", "q1_ind", "q_ig", "ind_run", "g_ind_27", "g_ind_28", "g_ind_29"]]]],
  ["全公司營運：實績錨點、毛利率與費用", [
    ["FY2026 第三、四季與 FY2027 第一季實績", ["q3_rev", "q4_rev", "q1_rev", "q1_gm", "q1_opx", "q1_eps", "sh_q4", "sh_q1"]],
    ["FY2025、FY2026 全年實績", ["a25_rev_in", "a26_rev_in", "a25_dcc_in", "a26_dcc_in", "a25_ind_in", "a26_ind_in", "a25_gp_in", "a26_gp_in",
                               "a25_opx_in", "a26_opx_in", "a25_int_in", "a26_int_in", "a25_tax_in", "a26_tax_in", "a25_ni_in", "a26_ni_in",
                               "a25_eps_in", "a26_eps_in", "a25_sh_in", "a26_sh_in"]],
    ["毛利率路徑（FY2027 起）", ["gm_27", "gm_28", "gm_29"]],
    ["營業費用、利息、稅、少數股權與股數", ["phi", "opx_adj", "int_f", "tx", "nci", "sh_g"]]]],
  ["全公司：股價、折現率、目標價規則與收斂法", [
    ["股價", ["px", "px_hi"]],
    ["目標價規則（情境權重、DCF 權重）", ["tp_wA", "tp_wB", "tp_wC", "tp_w_dcf"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "beta_u", "beta_u_med", "beta_du", "beta_de", "beta", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨負債與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "bs_date", "stub_frac", "stub_div", "sh_v"]],
    ["自由現金流（股票薪酬、折舊、資本支出、營運資金）", ["fc_sbc", "fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]],
    ["收斂法（FY2030 起）", ["tp_g", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]],
    ["共識目標價與前瞻 EPS（只作對照）", ["cons_tp", "cons_feps"]]]]],
 "blocks_to": {"Coherent 有效份額": "資料中心（AI 光模組：可插拔與 CPO）", "@Units!A14": "資料中心（AI 光模組：可插拔與 CPO）"},
 "main": {"DCF 彙總": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "淨現金滾到評價日", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
          "DCF 折現率": ["股權成本", "稅後債務成本", "權益占資本", "WACC", "永續成長率", "永續期 RONIC"],
          "Comparable 計算": ["同業前瞻 P/E 中位數", "12 個月後的前瞻 EPS", "Comparable 目標價"],
          "sheet:Recon": ["FY2027 EPS vs 共識", "FY2028 EPS vs 共識", "FY2027 Q1 EPS（模型，實績", "CY2028 Coherent＋中際旭創", "FY2027 資料中心 ÷ CY2027"]},
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
secs = {s["id"]: s for s in cfg["sections"]}
# 後處理：轉換器把估值章所有長條圖移到最前面；舊法（規則值）長條圖放回舊法段落，避免與目標價並列
_vb = secs["valuation"]["blocks"]
_old = next(b for b in _vb if b["type"] == "bars" and b.get("title") == "規則值隱含股價與現價")
_vb.remove(_old)
_h = next(i for i, b in enumerate(_vb) if b["type"] == "h3" and str(b.get("text", "")).startswith("規則值隱含股價"))
_vb.insert(_h + 2, _old)
# 後處理：本模型的龍捲風圖是 FY2028 EPS（不是加權目標價），轉換器的通用標題改回原意
for b in secs["sens"]["blocks"]:
    if b["type"] == "tornado":
        b["title"] = "FY2028 EPS 相對基準的變動（情境 A，一次動一個輸入；" + ("未達兩位數百分比）" if b.get("collapse") else "變動達兩位數百分比的案例）")
# 後處理（補充規則 A）：橫幅中間格＝只看 Comparable（市場主流同業法、未經相對倍數調整）；附註列進加權的調整後值
_rc, _ru, _ra = row(VA, "Comparable 目標價"), row(VA, "上檔空間：Comparable"), row(VA, "Comparable 目標價（相對倍數調整後）")
assert VA.cell(_rc, 1).value == "Comparable 目標價" and str(VA.cell(_ra - 1, 1).value).startswith("相對倍數調整係數")
cfg["hero"]["tiles"][1] = ["只看 Comparable 目標價（市場主流同業法）", f"@Val!B{_rc}|n2", ["上檔空間 ", f"@Val!B{_ru}|sp1"]]
cfg["hero"]["note"] = list(cfg["hero"].get("note") or []) + ["。進加權的 Comparable 為相對倍數調整後 ", f"@Val!B{_ra}|n2", "（調整係數 ", f"@Val!B{_ra - 1}|n2", "）"]
# 第三輪修正 17：橫幅金額標明幣別
_CUR3 = '美元'
def _lab3(x):
    t = str(wb[x[1:].split("!")[0]][x.split("!")[1]].value) if isinstance(x, str) and x.startswith("@") else str(x)
    return x if _CUR3 in t else ([x, "（" + _CUR3 + "）"] if isinstance(x, str) and x.startswith("@") else t + "（" + _CUR3 + "）")
for _t in cfg["hero"]["tiles"]:
    _t[0] = _lab3(_t[0])

# 三層協調 C4（規則 v1.4；tools/coord/coord.py）：橫幅另一側讀數一格與揭露句、估值章開頭兩讀數表與長條圖
sys.path.insert(0, os.path.join(HERE, "..", "tools", "coord"))
import coord as COORD  # noqa: E402
COORD.inject_html(cfg, X)


# ================= v0.4.5（2026-11-14）：個股 HTML 版面規格 v1.3（樣板 NVIDIA v0.3.7、Ibiden v0.5.6；docs/plan/個股HTML版面規格.md「v1.3 規格」） =================
# 章節：命題 → 估值 → 敏感度 → 反向 DCF → 損益（標準損益表）→ 驅動因子（分部營收推導）→ 對帳（含 C0 實績）→ 資料狀態；本公司無個別模組。
# 數字一律取自 xlsx（README 頁末 v1.3 白話列、PL 頁末標準損益表、RevDCF 頁、Val 0.7、Sens 頁末相對基準表、Coord 頁 F–H 節；cohr/v045_post.py 產生）。列號依標籤找。
def _v13():
    S = {x["id"]: x for x in cfg["sections"]}
    RD, PLW, RV, SW, CW = wb["README"], wb["PL"], wb["RevDCF"], wb["Sens"], wb["Coord"]
    fnd = lambda ws, p, start=1: next(r for r in range(start, ws.max_row + 1) if str(ws.cell(r, 1).value or "").strip().startswith(p))
    r13 = row(RD, "v1.3 展示層")
    RM = dict(zip(["hero", "rec", "o1", "o2", "o3", "o4", "o5", "drv", "coord", "ver"], range(r13 + 1, r13 + 11)))
    assert str(RD[f"A{RM['ver']}"].value).startswith("v0.4.5"), RD[f"A{RM['ver']}"].value
    cfg["layout"] = "v1.3"
    fmt_u = lambda u: {"%": "p1", "十億股": "n3", "x": "n2"}.get(str(u).strip(), "n2")
    # ---- 橫幅：200 字內白話＋對帳中揭露句；評等規則與歸因移到估值章 ----
    old_note = list(cfg["hero"]["note"])
    i_cd = len(old_note) - 2
    assert old_note[i_cd] == " " and str(old_note[i_cd + 1]).startswith("@Coord!B"), old_note
    cfg["hero"]["note"] = [f"@README!A{RM['hero']}"] + old_note[i_cd:]
    # ---- 1. 命題：主圖＝三情境營收與情境 B 分部（堆疊）----
    rb = row(PLW, "情境 B：")
    hdr0 = row(PLW, "情境／項目")
    sb = row(PLW, "情境 B", 1)  # 摘要區（列 17）
    assert PLW.cell(sb, 1).value == "情境 B", PLW.cell(sb, 1).value
    sa = next(r for r in range(hdr0, PLW.max_row + 1) if PLW.cell(r, 1).value == "情境 A")
    sc = next(r for r in range(hdr0, PLW.max_row + 1) if PLW.cell(r, 1).value == "情境 C")
    sfind = lambda base, lab_: next(r for r in range(base, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip() == lab_)
    r_dcB, r_cmB, r_inB, r_rvB = (sfind(sb, "資料中心"), sfind(sb, "通訊"), sfind(sb, "工業"), sfind(sb, "營收"))
    revA, revC = sfind(sa, "營收"), sfind(sc, "營收")
    th = S["thesis"]["blocks"]
    chart1 = {"type": "theme", "title": "營收分部（情境 B）與三情境營收（十億美元）", "x": f"@PL!D{hdr0}:G{hdr0}", "fmt": "n1", "stack": True, "pad": 1.35,
              "total": f"@PL!D{r_rvB}:G{r_rvB}",
              "bars": [{"name": f"@PL!A{r_dcB}", "y": f"@PL!D{r_dcB}:G{r_dcB}", "color": "#0b7a80"},
                       {"name": f"@PL!A{r_cmB}", "y": f"@PL!D{r_cmB}:G{r_cmB}", "color": "#6fc7cb"},
                       {"name": f"@PL!A{r_inB}", "y": f"@PL!D{r_inB}:G{r_inB}", "color": "#8a8c8c"}],
              "lines": [{"name": ["情境 A 營收"], "y": f"@PL!D{revA}:G{revA}", "color": "#0b7a80", "dash": True},
                        {"name": ["情境 C 營收"], "y": f"@PL!D{revC}:G{revC}", "color": "#b4472f", "dash": True}],
              "caption": "堆疊＝情境 B（基準）的三個分部，頂端數字＝情境 B 合併營收；虛線＝情境 A、情境 C 的合併營收。FY2026 為實績錨點；三情境差異來自 AI 光模組市場的顆數係數（CY2028 起）。"}
    i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
    th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RM[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
    th.insert(0, chart1)
    th.insert(1, {"type": "note", "parts": [f"@README!A{RM['rec']}"]})
    # ---- 2. 估值：第一個區塊＝四個目標價與現價 ----
    va = S["valuation"]["blocks"]
    g0 = row(VA, "0.7 估值並列"); gr = list(range(g0 + 2, g0 + 7)); r_pxg = g0 + 7
    four = {"type": "bars", "title": "四個目標價與現價（美元；豎線＝現價）", "fmt": "n2",
            "items": [{"label": lb, "value": f"@Val!B{r}", "series": sr} for r, sr, lb in zip(gr, (1, 2, 2, 0, 1),
                      ("DCF 目標價", "Comparable（調整後，進加權）", "Comparable（未調整，不進加權）", "加權目標價（橫幅主值）", "市場共識目標價"))],
            "ref": {"label": f"@Val!A{r_pxg}", "value": f"@Val!B{r_pxg}"}}
    tbl4 = {"type": "table", "caption": f"@Val!A{g0}", "head": [f"@Val!{c}{g0 + 1}" for c in "ABCD"],
            "rows": [[f"@Val!A{r}", f"@Val!B{r}|n2", f"@Val!C{r}|sp1", f"@Val!D{r}"] for r in gr] + [[f"@Val!A{r_pxg}", f"@Val!B{r_pxg}|n2", "", f"@Val!D{r_pxg}"]]}
    keep = [b for b in va if not (b.get("type") == "bars" and b.get("title") == "目標價與現價")]
    head_note = {"type": "note", "parts": old_note[:i_cd]}
    S["valuation"]["blocks"] = [four, head_note, tbl4] + keep
    intro = S["valuation"].pop("intro", None)
    if intro:
        S["valuation"]["blocks"].insert(2, {"type": "note", "parts": [intro]})
    # ---- 3. 敏感度：主圖＝加權目標價相對基準 ----
    se = S["sens"]["blocks"]
    t0 = fnd(SW, "加權目標價相對基準（HTML"); r1 = t0 + 2; r2 = r1
    while SW.cell(r2 + 1, 1).value:
        r2 += 1
    tor = []
    for col, side in (("B", "低"), ("C", "高")):
        vals = {r: SW[f"{col}{r}"].value for r in range(r1, r2 + 1) if isinstance(SW[f"{col}{r}"].value, (int, float))}
        big = [r for r, v in vals.items() if abs(v) >= 0.1]
        small = [r for r, v in vals.items() if 0.0005 <= abs(v) < 0.1]
        base = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1"}
        tor.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值，變動達兩位數百分比的案例", rows=",".join(map(str, big))))
        tor.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值", rows=",".join(map(str, small)), collapse=True,
                        summary=f"展開：加權目標價相對基準：輸入取{side}值，變動未達兩位數百分比的案例"))
    full = {"type": "table", "caption": f"@Sens!A{t0}", "collapse": True, "summary": "展開：全部案例的加權目標價相對基準（低案、高案）",
            "head": [f"@Sens!{c}{t0 + 1}" for c in "ABC"], "rows": [[f"@Sens!A{r}", f"@Sens!B{r}|sp1", f"@Sens!C{r}|sp1"] for r in range(r1, r2 + 1)]}
    snap = fnd(SW, "快照狀態")
    for b in se:
        if b.get("type") == "tornado":
            b["summary"] = ("展開：對照：FY2028 EPS 相對基準的變動未達兩位數百分比的案例（情境 A）" if b.get("collapse") else "展開：對照：FY2028 EPS 相對基準的變動達兩位數百分比的案例（情境 A，一次動一個輸入）")
            b["collapse"] = True
        if b.get("type") == "table" and not b.get("collapse"):
            b["collapse"] = True
            b["summary"] = "展開：目標價參數表（WACC、β、ERP、永續成長率等）"
    intro_s = {"type": "note", "parts": ["主圖為加權目標價相對基準的變動（低案、高案分開，一次動一個輸入；變動達兩位數百分比者直接顯示，其餘摺疊）。營運輸入案例為整本重算快照，快照狀態：",
                                          f"@Sens!B{snap}", "；估值框架與折現參數案例取自 Val 頁與 Sens 頁「目標價參數敏感度」表。下方摺疊為全部案例與 FY2028 EPS 敏感度對照。"]}
    S["sens"].pop("intro", None)
    S["sens"]["blocks"] = tor[:1] + tor[2:3] + [intro_s, tor[1], tor[3], full] + se
    # ---- 4. 反向 DCF（新）----
    rv = lambda p, start=1: next(r for r in range(start, RV.max_row + 1) if str(RV.cell(r, 1).value or "").strip().startswith(p))
    ra, rcmp, rsum, yh = rv("A. DCF"), rv("B. Comparable"), rv("C. 摘要"), rv("項目（情境 B）")
    rb_, rbb, rcc = rv("(a)"), rv("(b)"), rv("(c)")
    m1, m0 = rv("現價隱含的 EBIT 率", rb_), rv("模型：基準的等效 EBIT 率", rb_)
    s1, rv1, rv0, g1, g0b = rv("現價隱含的營收倍數", rbb), rv("現價隱含的 FY2028 營收（", rbb), rv("模型：FY2028 營收（", rbb), rv("現價隱含的 FY2028 營收年增率", rbb), rv("模型：FY2028 營收年增率", rbb)
    w1, w0 = rv("現價隱含的 WACC", rcc), rv("模型 WACC", rcc)
    a_rows = [rv(p, ra) for p in ("支持現價所需的評價日每股價值", "淨現金滾到評價日", "評價用稀釋股數", "支持現價所需的企業價值", "模型企業價值", "EV 需放大", "模型：情境 B 評價日每股價值")]
    afmt = ["n2", "n1", "n3", "n1", "n1", "n2", "n2"]
    pe1, rc1, pe2 = rv("只看 Comparable：支持現價所需的前瞻本益比", rcmp), rv("只看 Comparable：所需的相對倍數係數", rcmp), rv("加權（DCF 不變）", rcmp)
    yrs_rows = list(range(yh + 1, yh + 12))
    yfmt = ["n1", "n1", "p1", "p1", "n3", "n3", "n3", "n3", "n1", "n1", "n1"]
    rev_ = {"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2",
            "blocks": [
                {"type": "note", "parts": [f"@RevDCF!A{rsum + 1}"]},
                {"type": "note", "parts": [f"@RevDCF!A{rsum + 2}"]},
                {"type": "table", "caption": "現金流折現法（情境 B 路徑）：一次只動一個因素，反解到現價", "head": ["項目", "現價隱含", "模型基準"],
                 "rows": [["預測期營業利益率（FY2028–FY2032 統一一個值；模型欄為等效值，終值年權重大）", f"@RevDCF!B{m1}|p1", f"@RevDCF!B{m0}|p1"],
                          ["營收放大倍數（FY2028 起各年同乘；模型基準為一倍）", f"@RevDCF!B{s1}|n2", "—"],
                          ["FY2028 營收（十億美元）", f"@RevDCF!B{rv1}|n1", f"@RevDCF!B{rv0}|n1"],
                          ["FY2028 營收年增率", f"@RevDCF!B{g1}|p1", f"@RevDCF!B{g0b}|p1"],
                          ["折現率（WACC）", f"@RevDCF!B{w1}|p2", f"@RevDCF!B{w0}|p2"]]},
                {"type": "table", "caption": f"@RevDCF!A{ra}", "head": ["項目", "值"], "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|{f}"] for r, f in zip(a_rows, afmt)]},
                {"type": "table", "caption": f"@RevDCF!A{rcmp}", "head": [f"@RevDCF!{c}{rcmp + 1}" for c in "ABCD"],
                 "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|{f}", f"@RevDCF!C{r}|{f}", f"@RevDCF!D{r}|sp1"] for r, f in ((pe1, "n1"), (rc1, "n2"), (pe2, "n1"))]},
                {"type": "table", "caption": f"@RevDCF!A{yh}", "head": [f"@RevDCF!{c}{yh}" for c in "ABCDEFG"], "collapse": True,
                 "summary": "展開：情境 B 各年數字（反解的計算基礎；演算過程見 RevDCF 頁）",
                 "rows": [[f"@RevDCF!A{r}"] + [f"@RevDCF!{c}{r}|{f}" for c in "BCDEFG"] for r, f in zip(yrs_rows, yfmt)]}]}
    # ---- 5. 損益：標準損益表 ----
    p0 = fnd(PLW, "標準損益表（v0.4.5")
    ph = next(r for r in range(p0, PLW.max_row + 1) if PLW.cell(r, 1).value == "項目")
    plr = lambda lab_, start=ph: next(r for r in range(start, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip() == lab_)
    keys = ["營收", "營收年增率（YoY）", "營業成本（COGS）", "毛利", "毛利率", "營業費用（研發＋銷管）", "營業利益", "營業利益率", "其他收入（淨額）＝−利息與其他淨支出", "稅前利益", "所得稅",
            "少數股權調整（加回；歷史＝淨利 −（稅前 − 所得稅））", "稅後淨利（歸屬 Coherent）", "淨利率", "稀釋股數", "EPS（完全稀釋；non-GAAP）"]
    PC = "CDEFGHIJK"
    rowx = lambda r: [f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmt_u(PLW.cell(r, 2).value)}" for c in PC]
    r_ratio = plr("比率"); r_gaap = plr("GAAP 對照（歷史；預測不估 GAAP，故只列實績）"); r_chk = next(r for r in range(r_ratio, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("核對（應為 0"))
    r_end = next(r for r in range(r_chk, PLW.max_row + 2) if not PLW.cell(r, 1).value) - 1
    gaap_rows = [[f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmt_u(PLW.cell(r, 2).value)}" for c in "CDE"] + [""] * 6 for r in range(r_gaap + 1, r_chk)]
    is_tbl = {"type": "table", "caption": "標準損益表（十億美元；EPS 為美元）：FY2024A–FY2026A 實績、FY2027E–FY2032E 情境 B（基準；Non-GAAP）",
              "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in PC],
              "rows": [rowx(plr(k)) for k in keys] + [{"sub": f"@PL!A{r_ratio}"}] + [rowx(r) for r in range(r_ratio + 1, r_gaap)]
                      + [{"sub": f"@PL!A{r_gaap}"}] + gaap_rows}
    chk_tbl = {"type": "table", "caption": f"@PL!A{r_chk}", "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in PC], "collapse": True,
               "summary": "展開：核對列（歷史捨入差、預測期對 PL 情境 B、FY2030E–FY2032E 的推法差額）",
               "rows": [[f"@PL!A{r}"] + [f"@PL!{c}{r}|{'p1' if '占營收' in str(PLW.cell(r, 1).value) else 'n3'}" for c in PC] for r in range(r_chk + 1, r_end + 1)]}
    rv_, gm_, opm_, nm_ = plr("營收"), plr("毛利率"), plr("營業利益率"), plr("淨利率")
    is_chart = {"type": "theme", "title": "營收（十億美元）與毛利率、營業利益率、淨利率：FY2024A–FY2032E（情境 B）", "x": f"@PL!C{ph}:K{ph}",
                "fmt": "n1", "fmt_r": "p0", "stack": True, "pad": 1.6,
                "bars": [{"name": f"@PL!A{rv_}", "y": f"@PL!C{rv_}:K{rv_}", "color": "#6fc7cb"}],
                "lines": [{"name": f"@PL!A{gm_}", "y": f"@PL!C{gm_}:K{gm_}", "axis": "r", "color": "#0b7a80", "labels": True},
                          {"name": f"@PL!A{opm_}", "y": f"@PL!C{opm_}:K{opm_}", "axis": "r", "color": "#c29a2b"},
                          {"name": f"@PL!A{nm_}", "y": f"@PL!C{nm_}:K{nm_}", "axis": "r", "color": "#6f6bae", "dash": True}],
                "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸，Non-GAAP）。FY2027E 含 FY2027 第一季實績與第二季指引；FY2030E–FY2032E 為收斂法延伸期（毛利率與營業費用率沿用前一年）。"}
    old_pl = S["pl"]["blocks"]
    pls = [b for b in old_pl if b["type"] == "pl"]
    assert len(pls) == 2, [b["type"] for b in old_pl]
    ext_notes = [b for b in old_pl if b["type"] == "note"]
    kvs = [b for b in old_pl if b["type"] == "kv"]
    old_blocks = [dict(pls[0], collapse=True, summary="展開：三情境損益摘要與折線（FY2025–FY2029；營收、資料中心、毛利率、營業利益率、EPS）")]
    old_blocks += [dict(b, collapse=True, summary="展開：FY2026 模型對實績的校準核對") for b in kvs]
    old_blocks += [dict(b, collapse=True, summary="展開：延伸期（FY2030 起）的收斂法說明") for b in ext_notes]
    old_blocks += [dict(pls[1], collapse=True, summary="展開：三情境延伸到 FY2032（收斂法）")]
    S["pl"].pop("intro", None)
    S["pl"].update(title="損益（標準損益表）",
                   blocks=[{"type": "note", "parts": [f"@PL!A{p0 + 1}"]}, is_chart, is_tbl, {"type": "note", "parts": [f"@PL!A{p0 + 2}"]}, chk_tbl] + old_blocks)
    # ---- 6. 驅動因子：分部營收推導（情境 B）----
    prow = lambda lab_: next(r for r in range(rb, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip().startswith(lab_))
    hdrB = rb + 1
    rA = prow("A. 日曆年 AI 光學"); rB_ = prow("B. 會計年度營收"); rC_ = prow("C. 損益")
    r_mk = [prow(p_) for p_ in ("AI 可插拔光模組市場", "AI CPO 光學市場", "Coherent 可插拔有效份額（CY2026", "份額相對值", "可插拔有效份額（日曆年）", "單價係數", "Coherent 占 CPO", "資料中心營收（日曆年）")]
    r_dc, r_cm, r_in = prow("資料中心"), prow("通訊"), prow("工業（FY2027 起排除")
    r_rev2 = next(r for r in range(rB_, PLW.max_row + 1) if PLW.cell(r, 1).value == "營收")
    r_yoy, r_shr = r_rev2 + 1, r_rev2 + 2
    r_o = [next(r for r in range(rC_, PLW.max_row + 1) if PLW.cell(r, 1).value == k) for k in ("營業利益率", "歸屬 Coherent 淨利", "EPS")]
    seg = lambda rs: [[f"@PL!A{r}", f"@PL!B{r}"] + [f"@PL!{c}{r}|{fmt_u(PLW.cell(r, 2).value)}" for c in "DEFG"] for r in rs]
    seg_tbl = {"type": "table", "caption": "分部營收推導（情境 B；十億美元）：資料中心＋通訊＋工業＝合併營收", "head": [f"@PL!A{hdrB}", f"@PL!B{hdrB}"] + [f"@PL!{c}{hdrB}" for c in "DEFG"],
               "rows": [{"sub": f"@PL!A{rA}"}] + seg(r_mk) + [{"sub": f"@PL!A{rB_}"}] + seg([r_dc, r_cm, r_in, r_rev2, r_yoy, r_shr]) + [{"sub": f"@PL!A{rC_}"}] + seg(r_o)}
    seg_chart = {"type": "theme", "title": "分部營收（情境 B，十億美元）與資料中心占營收比重（右軸）", "x": f"@PL!D{hdrB}:G{hdrB}", "fmt": "n1", "fmt_r": "p0", "stack": True, "pad": 1.2,
                 "bars": [{"name": f"@PL!A{r_dc}", "y": f"@PL!D{r_dc}:G{r_dc}", "color": "#0b7a80"}, {"name": f"@PL!A{r_cm}", "y": f"@PL!D{r_cm}:G{r_cm}", "color": "#6fc7cb"},
                          {"name": f"@PL!A{r_in}", "y": f"@PL!D{r_in}:G{r_in}", "color": "#8a8c8c"}],
                 "lines": [{"name": f"@PL!A{r_shr}", "y": f"@PL!D{r_shr}:G{r_shr}", "axis": "r", "color": "#3b3838", "labels": True}],
                 "caption": "長條＝情境 B 的三個分部；折線＝資料中心占合併營收比重。FY2026 為實績錨點（資料中心與通訊由季度拆分推導，公司只揭露合計）。"}
    dr = S["drivers"]["blocks"]
    S["drivers"].update(title="驅動因子：分部營收推導")
    S["drivers"]["blocks"] = [{"type": "note", "parts": [f"@README!A{RM['drv']}"]}, seg_chart, seg_tbl, {"type": "h3", "text": "情境驅動與單一參數"}] + dr
    # ---- 7. 對帳：C0（實績）與讀取層級 ----
    cF, cG, cH = fnd(CW, "F. 讀取層級"), fnd(CW, "G. C0（實績）"), fnd(CW, "H. 對帳中年份")
    fh = cF + 1
    ih = next(r for r in range(cG, CW.max_row + 1) if CW.cell(r, 1).value == "輸入")
    mh = next(r for r in range(ih, CW.max_row + 1) if CW.cell(r, 1).value == "項目" and CW.cell(r, 2).value == "公司實績")
    verd = next(r for r in range(mh, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith("C0（實績）判定"))
    cmp_ = lambda ws_, r, c, f: (f"@Coord!{c}{r}|{f}" if CW[f"{c}{r}"].value is not None else "")
    ifm = lambda r: "n0" if CW.cell(r, 3).value == "月" else "p1" if CW.cell(r, 3).value == "%" else "n3"
    in_rows = [[f"@Coord!A{r}", f"@Coord!B{r}|{ifm(r)}", f"@Coord!C{r}", f"@Coord!D{r}", f"@Coord!E{r}", f"@Coord!F{r}"] for r in range(ih + 1, mh - 1) if CW.cell(r, 1).value]
    act_rows = [[f"@Coord!A{r}", cmp_(CW, r, "B", "n3"), cmp_(CW, r, "C", "n3"), cmp_(CW, r, "D", "sp1"), f"@Coord!E{r}", f"@Coord!F{r}", f"@Coord!G{r}", f"@Coord!H{r}"] if CW.cell(r, 7).value else
                [f"@Coord!A{r}", "", "", "", f"@Coord!E{r}", f"@Coord!F{r}", "", ""] for r in range(mh + 1, verd + 1)]
    lv0 = next(r for r in range(verd, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith("層級之間")); lvh = lv0 + 1
    lv_rows = [[f"@Coord!A{r}", f"@Coord!B{r}|n2", f"@Coord!C{r}|n2", f"@Coord!D{r}|sp1", f"@Coord!E{r}", f"@Coord!F{r}", f"@Coord!G{r}"] for r in range(lvh + 1, cH - 1) if CW.cell(r, 1).value]
    h_rows = [[f"@Coord!A{r}"] + [f"@Coord!{c}{r}" for c in "BCDEFG"] for r in range(cH + 2, cH + 4)]
    cnt = fnd(CW, "對帳中的 C0 項目數")
    act_blocks = [{"type": "h3", "text": "三層協調：由下往上對帳（公司實績對層級隱含值）"}, {"type": "note", "parts": [f"@Coord!A{cG + 1}"]}, {"type": "note", "parts": [f"@Coord!A{cG + 2}"]},
                  {"type": "table", "caption": "資料中心與通訊（DCC）營收：已公布季度實績（公司財報新聞稿） vs 層級隱含值（主讀數、第二讀數；層級份額由本模型校準，判定標參考）",
                   "head": [f"@Coord!{c}{mh}" for c in "ABCDEFGH"], "rows": act_rows},
                  {"type": "note", "parts": [f"@Coord!I{verd}"]},
                  {"type": "note", "parts": ["對帳中項目數：", f"@Coord!B{cnt}|n0", "。"]},
                  {"type": "table", "caption": "實績與層級隱含值的輸入、等級與出處（展開看）", "collapse": True, "summary": "展開：實績、層級值、換算參數與出處", "head": [f"@Coord!{c}{ih}" for c in "ABCDEF"], "rows": in_rows},
                  {"type": "h3", "text": "層級之間與外部估計（參考，不計入）"},
                  {"type": "table", "head": [f"@Coord!{c}{lvh}" for c in "ABCDEFG"], "rows": lv_rows},
                  {"type": "table", "caption": f"@Coord!A{cF}", "collapse": True, "summary": "展開：本模型讀哪一層、哪一版（下游只註明版本、不自行重估）", "head": [f"@Coord!{c}{fh}" for c in "ABCDE"],
                   "rows": [[f"@Coord!{c}{r}" for c in "ABCDE"] for r in range(fh + 1, cG - 2) if CW.cell(r, 1).value]},
                  {"type": "note", "parts": [f"@Coord!A{cH + 5}"]}, {"type": "note", "parts": [f"@Coord!A{cH + 6}"]} if False else {"type": "note", "parts": [f"@Coord!A{cnt + 1}"]},
                  {"type": "h3", "text": "公司目標、先驗比對與校準"}]
    S["recon"]["blocks"] = act_blocks + S["recon"]["blocks"]
    # ---- 8. 資料狀態 ----
    st = S["status"]["blocks"]
    st[0:0] = [{"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RM['ver']}"]}]
    order = ["thesis", "valuation", "sens", "revdcf", "pl", "drivers", "recon", "status"]
    S["revdcf"] = rev_
    cfg["sections"] = [S[k] for k in order]


_v13()

json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")
