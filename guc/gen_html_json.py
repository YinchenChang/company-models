# 創意：把 html_src.json（v1.0 之前的頁面設定；v0.3 起 In 列號與 Sens 目標價參數表已對應標準版 DCF）依「個股 HTML 版面規格 v1.0」（docs/plan/個股HTML版面規格.md）轉成 html.json。
# 用法：python3 guc/gen_html_json.py <成品 xlsx>（讀 guc/html_src.json，寫 guc/html.json；勿手改 html.json）；之後 python3 tools/company_html/build_company_html.py alchip
import copy, json, os, re, sys
import openpyxl
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools", "company_html"))
import relayout as RL  # noqa: E402

X = sys.argv[1]
WB = openpyxl.load_workbook(X, data_only=True)
cfg = copy.deepcopy(json.load(open(os.path.join(HERE, "html_src.json"), encoding="utf-8")))
SEC = {s["id"]: s for s in cfg["sections"]}


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def row_of(sh, prefix, need_num=False, exact=False):
    ws = WB[sh]
    for r in range(1, ws.max_row + 1):
        t = str(ws.cell(r, 1).value or "").strip()
        if (t == prefix if exact else t.startswith(prefix)) and (not need_num or num(ws.cell(r, 2).value)):
            return r
    raise KeyError((sh, prefix))


I = {str(WB["In"].cell(r, 1).value): r for r in range(4, WB["In"].max_row + 1) if WB["In"].cell(r, 1).value}


def in_tok(prefix):
    ws = WB["In"]
    for r in range(4, ws.max_row + 1):
        a = str(ws.cell(r, 1).value or "")
        if a.startswith(prefix) and ws.cell(r, 2).value in (None, ""):
            return re.split(r"[.\s]", a)[0]
    raise KeyError(prefix)


# ─── 驅動因子：舊「輸入 → 推導量」表移到附錄（轉換器依 caption「驅動 → 推導」辨認） ───
for b in SEC["drivers"]["blocks"]:
    if b["type"] == "table" and ((not b.get("caption") and b.get("head", [None])[0] == "輸入") or str(b.get("caption", "")).startswith("單一參數")):
        b["caption"] = "驅動 → 推導對照（單一參數：輸入、基準值、推導量與因果方向）"

# 第三輪修正：情境驅動表數值格式（年度不加千分位、百分比／pt 以 % 顯示、其他依大小取小數位）
def _scen_fmt(ws, r):
    u = str(ws.cell(r, 3).value or "").strip()
    vs = [ws.cell(r, c).value for c in (4, 5, 6)]
    vs = [v for v in vs if isinstance(v, (int, float)) and not isinstance(v, bool)]
    if not vs:
        return ""
    if all(float(v).is_integer() and 1990 <= v <= 2100 for v in vs) and ("年" in u or u in ("", "—")):
        return "|yr"
    if u in ("%", "pt", "百分點") or u.startswith("%"):
        return "|p2" if any(0 < abs(v) < 0.001 for v in vs) or any(abs(round(v * 1000) - v * 1000) > 1e-6 for v in vs) else "|p1"
    if all(float(v).is_integer() for v in vs):
        return "|n0"
    return "|n1" if max(abs(v) for v in vs) >= 100 else "|n2"
# ─── 情境驅動表：依 In 頁「情境驅動對照」節的實際位置重建（不寫死列號；v0.x 起 In 頁新增 M2 節，列號下移） ───
_s0 = row_of("In", "情境驅動對照"); _h0 = _s0 + 1
assert WB["In"].cell(_h0, 1).value == "鍵", WB["In"].cell(_h0, 1).value
_r1 = _h0 + 1
while WB["In"].cell(_r1, 1).value: _r1 += 1
_nsc = 0
for b in SEC["drivers"]["blocks"]:
    if b["type"] == "table" and str(b.get("caption", "")).startswith("情境驅動"):
        b["head"] = [f"@In!{c}{_h0}" for c in "BDEFC"]
        b["rows"] = [[f"@In!{c}{r}" + (_scen_fmt(WB["In"], r) if c in "DEF" else "") for c in "BDEFC"] for r in range(_h0 + 1, _r1)]
        _nsc += 1
assert _nsc == 1, _nsc

# ─── 估值：長條圖 → 結論主要項目 → DCF 折現率 → DCF 彙總 → Comparable → 其餘摺疊 ───
V = lambda k: row_of("Val", k, need_num=True)
tpw, dcft, comp, compadj, cons = V("目標價（DCF 與 Comparable 加權）"), V("DCF 目標價（情境加權）"), V("Comparable 目標價（市場主流同業法）"), V("Comparable 目標價（相對倍數調整後）"), V("市場共識目標價")
upw, upc, dcf12, px = V("上檔空間：加權目標價"), V("上檔空間：Comparable"), V("DCF 情境價值推到 12 個月後"), V("現價")
rec, grade = row_of("Val", "投資建議", exact=True), row_of("Val", "上檔空間級距（六級，供分組）", exact=True)
rk = row_of("Val", "股權成本 ke")                       # v0.3：0.1 節 ke、kd_at、w_e、WACC、g、RONIC，接著 DCF 彙總 10 列（第二輪修正：模組 v2 stub 期多兩列，依標籤定位）
assert dcft > rk + 6 and str(WB["Val"].cell(dcft, 1).value).startswith("DCF 目標價（情境加權）")
_dfm = lambda lab: ("g" if lab.startswith("評價用稀釋股數") else "p1" if "終值現值占 EV" in lab else "x1" if "隱含終值" in lab else
                    "n0" if lab.startswith(("DCF 情境", "DCF 目標價")) else "n1")
fb = next(r for r in range(row_of("DCF", "情境 B（平平）"), WB["DCF"].max_row + 1) if WB["DCF"].cell(r, 1).value == "營收")
hc = row_of("README", "展示層輔助：Comparable 目標價（相對倍數調整後）的上檔空間")
c0 = row_of("Val", "0.3 結論")
old = SEC["valuation"]["blocks"]
pick = lambda f: next(b for b in old if f(b))
bars = pick(lambda b: b["type"] == "bars")
for _it in bars["items"]:   # v0.3：0.1 節改版後 DCF 目標價列移到 dcft
    if _it["value"] == "@Val!B24": _it.update(label=f"@Val!A{dcft}", value=f"@Val!B{dcft}")
bars["items"].insert(2, {"label": f"@Val!A{compadj}", "value": f"@Val!B{compadj}", "series": 2})
peer_tbl = pick(lambda b: b.get("caption") == "@Val!A26")
comp_kv = pick(lambda b: b["type"] == "kv" and b["items"][0][0] == "@Val!A33")
px_kv = pick(lambda b: b["type"] == "kv" and b["items"][0][0] == f"@Val!A{px}")
rm_kv = pick(lambda b: b["type"] == "kv" and b["items"][0][0] == f"@Val!A{row_of('Val', '本公司 10 年中位數')}")
rule_tbls = [b for b in old if b["type"] == "table" and not b.get("caption") and not b.get("collapse")]
hist_tbl = pick(lambda b: str(b.get("summary", "")).startswith("展開：十年年底 P/E"))
notes = {b["parts"][0]: b for b in old if b["type"] == "note"}
H5 = ["項目", "數值", "", "", "說明"]
kv2tbl = lambda kv, cap, **kw: dict({"type": "table", "caption": cap, "head": ["項目", "數值"], "rows": [[k, v] for k, v in kv["items"]]}, **kw)
SEC["valuation"]["blocks"] = [
 bars,
 {"type": "table", "caption": "目標價結論（主要項目）", "head": H5, "rows": [
   [f"@Val!A{tpw}", f"@Val!B{tpw}|n0", "", "", f"@Val!E{tpw}"],
   [f"@Val!A{dcft}", f"@Val!B{dcft}|n0", "", "", f"@Val!E{dcft}"],
   [f"@Val!A{compadj}", f"@Val!B{compadj}|n0", "", "", f"@Val!E{compadj}"],
   [f"@Val!A{comp}", f"@Val!B{comp}|n0", "", "", f"@Val!E{comp}"],
   [f"@Val!A{px}", f"@Val!B{px}|n0", "", "", ""],
   [f"@Val!A{upw}", f"@Val!B{upw}|sp1", "", "", f"@Val!E{upw}"],
   [f"@Val!A{upc}", f"@Val!B{upc}|sp1", "", "", f"@Val!E{upc}"],
   ["Comparable（相對倍數調整後）的上檔空間", f"@README!B{hc}|sp1", "", "", ""],
   [f"@Val!A{rec}", f"@Val!B{rec}", "", "", f"@Val!E{rec}"],
   [f"@Val!A{dcf12}", f"@Val!B{dcf12}|n0", f"@Val!C{dcf12}|n0", f"@Val!D{dcf12}|n0", ""],
   [f"@Val!A{cons}", f"@Val!B{cons}|n0", "", "", f"@Val!E{cons}"]]},
 {"type": "table", "collapse": True, "summary": "展開：結論的其他項目（與共識的差距拆解、價格帶、十年分位、上檔空間級距）", "head": H5, "rows": [
   [f"@Val!A{r}", f"@Val!B{r}|{f}", f"@Val!C{r}|{f}" if r == c0 + 10 else "", f"@Val!D{r}|{f}" if r == c0 + 10 else "", f"@Val!E{r}"]
   for r, f in ((c0 + 7, "sp1"), (c0 + 8, "sp1"), (c0 + 9, ""), (c0 + 10, "p0"))] + [[f"@Val!A{grade}", f"@Val!B{grade}", "", "", f"@Val!E{grade}"]]},
 {"type": "table", "caption": "DCF 折現率（WACC、股權成本、稅後債務成本、RONIC、永續成長率；直接顯示）", "head": ["項目", "數值", "說明"], "rows": [
   [f"@Val!A{r}", f"@Val!B{r}|p2", f"@Val!E{r}"] for r in range(rk, rk + 6)]},
 {"type": "table", "caption": "DCF 彙總（標準版：無槓桿自由現金流 FCFF＋WACC；三情境；十億新台幣、每股新台幣）", "head": ["項目", "A（順利）", "B（平平）", "C（不如預期）", "說明"], "rows": [
   [f"@Val!A{r}", f"@Val!B{r}|{fm}", f"@Val!C{r}|{fm}" if r != dcft else "", f"@Val!D{r}|{fm}" if r != dcft else "", f"@Val!E{r}"]
   for r, fm in ((r, _dfm(str(WB["Val"].cell(r, 1).value).strip())) for r in range(rk + 6, dcft + 1))]},
 {"type": "table", "collapse": True, "summary": "展開：情境 B（平平）FCFF 逐年明細（DCF 頁；十億新台幣）", "head": ["@DCF!A5", "@DCF!C5", "@DCF!D5", "@DCF!E5", "@DCF!F5", "@DCF!G5", "@DCF!H5"],
  "rowgen": {"sheet": "DCF", "rows": f"{fb}-{fb + 10}", "cols": ["A", "C", "D", "E", "F", "G", "H"], "fmts": ["", "n1", "n1", "n1", "n1", "n1", "n1"],
             "row_fmt": {str(fb + 3): {c: "p1" for c in "CDEFGH"}, str(fb + 5): {c: "p1" for c in "CDEFGH"}, str(fb + 6): {c: "p1" for c in "CDEFGH"}}}},
 kv2tbl(comp_kv, "Comparable（同業前瞻 P/E 中位數 × 一年後的前瞻 EPS）"),
 dict(peer_tbl, collapse=True, summary="展開：同業前瞻倍數"),
 kv2tbl(rm_kv, "相對倍數調整（Comparable 目標倍數＝同業中位數 × 本公司相對同業的長期倍數比）"),
 notes["@Val!A48"],
 kv2tbl(px_kv, "價格與市場隱含倍數", collapse=True, summary="展開：現價與市場隱含倍數"),
 {"type": "h3", "text": "十年 P/E 規則值（舊法參考，不進目標價）"}] + [dict(b, collapse=True, summary=s) for b, s in zip(rule_tbls, ("展開：規則值隱含股價（FY2027 EPS × 依循環位置的倍數）", "展開：循環位置（本模型營收最高年）"))] + [
 notes["@Val!A2"], hist_tbl, notes["@Val!A82"]]

SEC["recon"]["blocks"][0]["caption"] = "本模型與公司說法、共識、券商的對照"

# 敏感度：兩張龍捲風是 FY2027／FY2028 EPS（情境 A），不是加權目標價，轉換器之後還原標題
TOR = {(b["sheet"], b["value_col"]): b["title"] for b in SEC["sens"]["blocks"] if b["type"] == "tornado"}

T1 = "Google 自研 Arm CPU（Axion）量產（TPU 顆數 × CPU 配比 × 創意份額 × 每顆營收）"
T2 = "其他 AI ASIC 量產（Maia 等：顆數 × 創意份額 × 每顆核心售價）"
SPEC = {
 "in_skip": [in_tok("G."), in_tok("情境驅動對照")],
 "in_groups": [
  [T1, [
    ["加速器顆數與情境係數（讀台積電模型＝市場層）", ["u_tpu_26", "u_tpu_27", "u_tpu_28", "g29_tpu", "um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C", "uf"]],
    ["CPU 對 TPU 配比、創意份額與每顆營收", ["att26", "att_A", "att_B", "att_C", "gsh26", "gsh28_A", "gsh28_B", "gsh28_C", "pc_cpu", "pfc_A", "pfc_B", "pfc_C"]],
    ["Wattonomics 錨定係數（只進敏感度）", ["wta_27", "wta_28", "wta_29"]]]],
  [T2, [
    ["其他 ASIC 顆數、每顆核心售價與份額（Broadcom、Marvell、聯發科份額讀上游）", ["u_oth_26", "u_oth_27", "u_oth_28", "g29_oth", "core_oth", "av_oth_26", "mv_oth_26", "mk_oth_26", "av_oth_27", "mv_oth_27", "mk_oth_27", "av_oth_28", "mv_oth_28", "mk_oth_28",
                                                    "av_tpu_26", "av_tpu_27", "av_tpu_28", "guc_oth_26", "go27_A", "go27_B", "go27_C", "go28_A", "go28_B", "go28_C"]],
    ["量產毛利率", ["m_tk", "gta_A", "gta_B", "gta_C"]]]],
  ["車用與其他營收（NRE＋其他量產）", [
    ["車用營收倍數與毛利率", ["am27_A", "am28_A", "ag29_A", "am27_B", "am28_B", "ag29_B", "am27_C", "am28_C", "ag29_C", "m_auto"]],
    ["其他營收年增與毛利率調整", ["g_o27_A", "g_oL_A", "g_o27_B", "g_oL_B", "g_o27_C", "g_oL_C", "gmadj"]]]],
  ["全公司（FY2026 錨點、費用與稅、股價、目標價規則、折現率與自由現金流）", [
    ["FY2026 錨點（上半年實績、第三季營收實績、第四季假設）", ["cpu_sh26", "asic_sh26", "auto_sh26", "cpu_sh25", "asic_sh25", "auto_sh25", "fx26", "q3_rev", "q3_gm", "q3_opx", "q4_g", "q4_gm", "q4_opx", "t_h1", "q_nonop", "q_tax", "q_sh"]],
    ["FY2027 起：營業費用、營業外、稅、股數、匯率", ["phi", "g_opx_min", "opx_adj", "nonop_f", "tax", "sh_g", "fx"]],
    ["股價與市場倍數", ["px", "pe_ttm", "fpe", "cons_feps", "cons_tp", "nim5"]],
    ["目標價規則（情境權重、DCF 權重）", ["tp_wA", "tp_wB", "tp_wC", "tp_w_dcf"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "rf_us", "beta_u", "beta_de", "beta_du", "beta", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "bs_date", "stub_frac", "stub_div", "sh_v"]],
    ["自由現金流（折舊、資本支出、營運資金）", ["fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]],
    ["收斂法（FY2030–FY2032）", ["tp_g", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]]]]],
 "blocks_to": {p_: T1 for p_ in ("1. Google", "2. 份額加總", "@Units!A27")},
 "main": {"DCF 彙總": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "＋Stub 期 FCFF", "淨現金滾到評價日", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
          "DCF 折現率": ["股權成本", "稅後債務成本", "權益占資本", "WACC", "永續成長率", "永續期 RONIC"],
          "目標價結論": ["目標價（DCF", "DCF 目標價", "Comparable 目標價（相對", "Comparable 目標價（市場", "現價", "上檔空間：加權", "上檔空間：Comparable", "Comparable（相對倍數調整後）的上檔空間", "投資建議", "DCF 情境價值推到", "市場共識"],
          "本模型與公司說法、共識、券商的對照": ["FY2026 EPS vs 共識", "CY2027 創意 CPU 顆數 vs 法人", "FY2027 EPS vs 市場共識", "FY2028 EPS vs 外資", "共識前瞻 12 月 EPS"]},
 "n_recon": 4,
}

out = RL.relayout(cfg, X, SPEC)

# 補充規則 A：橫幅中間格＝只看 Comparable（市場主流同業法，未經相對倍數調整）＋自己的上檔空間；附註寫進加權的調整後值
h = out["hero"]
h["tiles"][1] = ["只看 Comparable 目標價（市場主流同業法）", f"@Val!B{comp}|n0", ["上檔空間 ", f"@Val!B{upc}|sp1"]]
if abs(WB["Val"].cell(compadj, 2).value - WB["Val"].cell(comp, 2).value) > 1e-9:
    h["note"] = (h.get("note") or []) + ["；進加權的 Comparable 為相對倍數調整後 ", f"@Val!B{compadj}|n0", " 元。"]

for s in out["sections"]:
    for b in s["blocks"]:
        if b["type"] == "tornado":
            t0 = TOR.get((b["sheet"], b["value_col"]))
            if t0 and not t0.startswith("加權目標價"):
                b["title"] = t0 + ("（未達兩位數百分比的案例）" if b.get("collapse") else "（變動達兩位數百分比的案例）")
    if s["id"] == "status":   # 附錄「版本要點」只收版本說明列
        for b in s["blocks"]:
            if b.get("summary") == "展開：本版與前版的修改要點":
                ws = WB["README"]
                b["rows"] = [[f"@README!A{r}"] for r in range(1, ws.max_row + 1) if re.match(r"^v\d+\.\d", str(ws.cell(r, 1).value or ""))]

# 第三輪修正：橫幅金額單位（新台幣）
for _t in out["hero"]["tiles"]:
    if isinstance(_t[0], str):
        _t[0] = [_t[0], "（新台幣）"]
json.dump(out, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")
