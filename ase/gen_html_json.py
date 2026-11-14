# 日月光投控：把 html_src.json（v1.0 之前的頁面設定，內容不動）依「個股 HTML 版面規格 v1.0」轉成 html.json，再於 v0.2.6 起依版面規格 v1.3（docs/plan/個股HTML版面規格.md「v1.3 規格」）重排（本檔後半 _v13）。
# 用法：python3 ase/gen_html_json.py <成品 xlsx>（讀 ase/html_src.json，寫 ase/html.json；勿手改 html.json）
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
    if b["type"] == "table" and not b.get("caption") and b.get("head", [None])[0] == "輸入":
        b["caption"] = "驅動 → 推導對照（單一參數：輸入、基準值、推導量與因果方向）"

# ─── 估值：長條圖 → 結論主要項目 → 投資建議 → DCF 折現率 → DCF 彙總 → Comparable → 其餘摺疊 ───
V = lambda k: row_of("Val", k, need_num=True)
tpw, dcft, comp, compadj, cons = V("目標價（DCF 與 Comparable 加權）"), V("DCF 目標價（情境加權）"), V("Comparable 目標價（市場主流同業法）"), V("Comparable 目標價（相對倍數調整後）"), V("市場共識目標價")
upw, upc, dcf12, px = V("上檔空間：加權目標價"), V("上檔空間：Comparable"), V("DCF 情境價值推到 12 個月後"), V("現價")
rec, grade = row_of("Val", "投資建議", exact=True), row_of("Val", "上檔空間級距")
rk = row_of("Val", "股權成本 ke")                       # v0.2：0.1 節 ke、kd_at、w_e、WACC、g、RONIC，接著 DCF 彙總 10 列
r_pvx, r_dcf = row_of("Val", "明確期 FCFF 現值"), row_of("Val", "DCF 目標價（情境加權）")   # 第二輪修正：依標籤定位（模組 v2 stub 期多兩列）
assert r_pvx == rk + 6, (r_pvx, rk)
_fm = lambda lab: "n3" if lab.startswith("評價用") else "p1" if "終值現值占 EV" in lab else "x1" if "隱含終值" in lab else "n2" if lab.startswith(("DCF 情境", "DCF 目標價")) else "n1"
fb = next(r for r in range(row_of("DCF", "情境 B"), WB["DCF"].max_row + 1) if WB["DCF"].cell(r, 1).value == "營收")
hc = row_of("README", "展示層輔助：Comparable 目標價（相對倍數調整後）的上檔空間")
c0 = row_of("Val", "0.3 結論")
old = SEC["valuation"]["blocks"]
pick = lambda f: next(b for b in old if f(b))
bars = pick(lambda b: b["type"] == "bars")
for _it in bars["items"]:   # v0.2：0.1 節改版後 DCF 目標價列移到 dcft
    if _it["value"] == "@Val!B24": _it.update(label=f"@Val!A{dcft}", value=f"@Val!B{dcft}")
dcf_tbl = pick(lambda b: b.get("caption") == "@Val!A5")
peer_tbl = pick(lambda b: b.get("caption") == "@Val!A26")
comp_kv = pick(lambda b: b["type"] == "kv" and b["items"][0][0] == "@Val!A33")
px_kv = pick(lambda b: b["type"] == "kv" and b["items"][0][0] == f"@Val!A{px}")
rm_kv = pick(lambda b: b["type"] == "kv" and b["items"][0][0].startswith("@Val!A14"))
rule_tbls = [b for b in old if b["type"] == "table" and not b.get("caption") and not b.get("collapse")]
hist_tbl = pick(lambda b: str(b.get("summary", "")).startswith("展開：十年年底 P/E"))
notes = {b["parts"][0]: b for b in old if b["type"] == "note"}
H5 = ["項目", "數值", "", "", "說明"]
inr = lambda k, f: [f"@In!B{I[k]}", f"@In!E{I[k]}|{f}"]
kv2tbl = lambda kv, cap, **kw: dict({"type": "table", "caption": cap, "head": ["項目", "數值"], "rows": [[k, v] for k, v in kv["items"]]}, **kw)
SEC["valuation"]["blocks"] = [
 bars,
 {"type": "table", "caption": "目標價結論（主要項目）", "head": H5, "rows": [
   [f"@Val!A{tpw}", f"@Val!B{tpw}|n2", "", "", f"@Val!E{tpw}"],
   [f"@Val!A{dcft}", f"@Val!B{dcft}|n2", "", "", f"@Val!E{dcft}"],
   [f"@Val!A{compadj}", f"@Val!B{compadj}|n2", "", "", f"@Val!E{compadj}"],
   [f"@Val!A{comp}", f"@Val!B{comp}|n2", "", "", f"@Val!E{comp}"],
   [f"@Val!A{px}", f"@Val!B{px}|n2", "", "", ""],
   [f"@Val!A{upw}", f"@Val!B{upw}|sp1", "", "", f"@Val!E{upw}"],
   [f"@Val!A{upc}", f"@Val!B{upc}|sp1", "", "", f"@Val!E{upc}"],
   ["Comparable（相對倍數調整後）的上檔空間", f"@README!B{hc}|sp1", "", "", ""],
   [f"@Val!A{rec}", f"@Val!B{rec}", "", "", f"@Val!E{rec}"],
   [f"@Val!A{dcf12}", f"@Val!B{dcf12}|n1", f"@Val!C{dcf12}|n1", f"@Val!D{dcf12}|n1", ""],
   [f"@Val!A{cons}", f"@Val!B{cons}|n2", "", "", f"@Val!E{cons}"]]},
 {"type": "table", "collapse": True, "summary": "展開：結論的其他項目（與共識的差距拆解、價格帶、十年分位、上檔空間級距）", "head": H5, "rows": [
   [f"@Val!A{r}", f"@Val!B{r}|{f}", f"@Val!C{r}|{f}" if r == c0 + 10 else "", f"@Val!D{r}|{f}" if r == c0 + 10 else "", f"@Val!E{r}"]
   for r, f in ((c0 + 7, "sp1"), (c0 + 8, "sp1"), (c0 + 9, ""), (c0 + 10, "p0"))] + [[f"@Val!A{grade}", f"@Val!B{grade}", "", "", f"@Val!E{grade}"]]},
 {"type": "table", "caption": "DCF 折現率（WACC、股權成本、稅後債務成本、RONIC、永續成長率；直接顯示）", "head": ["項目", "數值", "說明"], "rows": [
   [f"@Val!A{r}", f"@Val!B{r}|p2", f"@Val!E{r}"] for r in range(rk, rk + 6)]},
 {"type": "table", "caption": "DCF 彙總（標準版：無槓桿自由現金流 FCFF＋WACC；三情境）", "head": ["項目", "A（樂觀）", "B（基準）", "C（保守）", "說明"], "rows": [
   [f"@Val!A{r}", f"@Val!B{r}|{fm}", f"@Val!C{r}|{fm}" if r != r_dcf else "", f"@Val!D{r}|{fm}" if r != r_dcf else "", f"@Val!E{r}"]
   for r, fm in ((r, _fm(str(WB["Val"].cell(r, 1).value).strip())) for r in range(r_pvx, r_dcf + 1))]},
 {"type": "table", "collapse": True, "summary": "展開：情境 B FCFF 逐年明細（DCF 頁；十億新台幣）", "head": ["@DCF!A5", "@DCF!C5", "@DCF!D5", "@DCF!E5", "@DCF!F5", "@DCF!G5", "@DCF!H5"],
  "rowgen": {"sheet": "DCF", "rows": f"{fb}-{fb + 10}", "cols": ["A", "C", "D", "E", "F", "G", "H"], "fmts": ["", "n1", "n1", "n1", "n1", "n1", "n1"],
             "row_fmt": {str(fb + 3): {c: "p1" for c in "CDEFGH"}, str(fb + 5): {c: "p1" for c in "CDEFGH"}, str(fb + 6): {c: "p1" for c in "CDEFGH"}}}},
 kv2tbl(comp_kv, "Comparable（同業前瞻 P/E 中位數 × 一年後的前瞻 EPS）"),
 dict(peer_tbl, collapse=True, summary="展開：同業前瞻倍數"),
 kv2tbl(rm_kv, "相對倍數調整（Comparable 目標倍數＝同業中位數 × 本公司相對同業的長期倍數比）"),
 notes["@Val!A48"],
 kv2tbl(px_kv, "價格與市場隱含倍數", collapse=True, summary="展開：現價與市場隱含倍數"),
 {"type": "h3", "text": "十年 P/E 規則值（舊法參考，不進目標價）"}] + [dict(b, collapse=True, summary=s) for b, s in zip(rule_tbls, ("展開：規則值隱含股價（FY2027 EPS × 依循環位置的倍數）", "展開：循環位置（本模型營收最高年）"))] + [
 notes["@Val!A2"], hist_tbl, notes["@Val!A101"]]

SEC["recon"]["blocks"][0]["caption"] = "本模型與公司指引、公司說法、券商的對照"

# 敏感度：第二張龍捲風是 FY2028 EPS（情境 A），不是加權目標價，轉換器之後還原標題
TOR = {(b["sheet"], b["value_col"]): b["title"] for b in SEC["sens"]["blocks"] if b["type"] == "tornado"}
# v0.2：舊折現率三案改為標準版 DCF 參數案例（同一張表），案例列與讀法列位置改變
_sl = next(r for r in range(5, 80) if WB["Sens"].cell(r, 1).value is None) - 1          # 最後一個案例列
_sn = row_of("Sens", "讀法：")

for b in SEC["sens"]["blocks"]:
    if b["type"] == "tornado": b["rows"] = f"5-{_sl}"
    if b["type"] == "table" and b.get("rowgen", {}).get("sheet") == "Sens": b["rowgen"]["rows"] = f"4-{_sl}"
    if b["type"] == "note" and b["parts"][0].startswith("@Sens!A") and b["parts"][0] != "@Sens!A2": b["parts"] = [f"@Sens!A{_sn}"]
pass


def scn_table_fix(sec):
    """第三輪修正第 32 點：情境驅動表依 In 頁「情境驅動對照」節標題找列範圍（不寫死列號）；百分比列以 % 顯示。"""
    ws = WB["In"]
    t = next(r for r in range(4, ws.max_row + 1) if str(ws.cell(r, 1).value or "").startswith("情境驅動對照"))
    h = t + 1; rows = []; r = h + 1
    while ws.cell(r, 1).value not in (None, ""):
        rows.append(r); r += 1
    assert ws.cell(rows[-1], 1).value == "scn_w", ws.cell(rows[-1], 1).value
    def fm(rr):
        u = str(ws.cell(rr, 3).value or "")
        vals = [ws.cell(rr, c).value for c in (4, 5, 6) if isinstance(ws.cell(rr, c).value, (int, float))]
        if u in ("%", "pt") or "%" in u: return "p1"
        return "n0" if vals and max(abs(v) for v in vals) >= 100 else "n2"
    for b in sec["blocks"]:
        if b.get("type") == "table" and str(b.get("caption", "")).startswith("情境驅動"):
            b["head"] = [f"@In!B{h}", f"@In!D{h}", f"@In!E{h}", f"@In!F{h}", f"@In!C{h}"]
            b["rows"] = [[f"@In!B{rr}", f"@In!D{rr}|{fm(rr)}", f"@In!E{rr}|{fm(rr)}", f"@In!F{rr}|{fm(rr)}", f"@In!C{rr}"] for rr in rows]
            b.pop("rowgen", None)
            return
    raise KeyError("情境驅動表")


scn_table_fix(SEC["drivers"])

SPEC = {
 "in_skip": [in_tok("G."), in_tok("情境驅動對照")],
 "in_groups": [
  ["封測（ATM）：LEAP 先進封裝（AI 加速器 × 委外份額 × 每顆內容＋全流程＋CPO）", [
    ["AI 加速器顆數與封裝面積（Siliconomics，唯讀）", ["units_26", "units_27", "units_28", "units_29", "ar_26", "ar_27", "ar_28", "ar_29"]],
    ["情境顆數係數（Siliconomics）", ["um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C"]],
    ["LEAP 錨點與每顆內容", ["leap25", "fx25", "leap26", "lsh_1h", "lsh_3q", "el"]],
    ["委外與份額相對值（三情境）", [k for k in I if re.match(r"^rel2\d_[ABC]$", k)]],
    ["全流程（CoWoS full process）", ["fp26", "fp27_A", "fp27_B", "fp27_C", "fp_g"]],
    ["CPO（共同封裝光學；Siliconomics 滲透率）", ["cpo27"] + [k for k in I if re.match(r"^cpo_[ABC]_\d\d$", k)]]]],
  ["封測（ATM）：主流封測", [
    ["主流封測成長", ["g_ms26", "g_ms27", "g_msL"]]]],
  ["EMS（環旭電子）", [
    ["EMS 成長與毛利率", ["g_ems27", "g_emsL", "ems_gm"]]]],
  ["全公司（資本支出與折舊、利潤率、FY2026 錨點、股價與目標價規則）", [
    ["匯率", ["fx"]],
    ["資本支出與折舊（核心問題之二）", ["capex25", "capex26_9m", "capex26", "capex27", "capi", "life", "q4_da"]],
    ["毛利、營業費用、營業外、稅與股數", ["ft", "phi", "nonop_f", "tax", "nci", "sh_g"]],
    ["FY2026 錨點（第四季依公司指引）", ["q4_atm", "q4_atmgm", "q4_ems", "q4_emsgm", "q4_opx", "q4_nonop", "q4_tax", "q4_nci"]],
    ["股價與市場倍數", ["px", "pe_ttm", "fpe", "cons_feps", "cons_tp"]],
    ["目標價規則（情境權重、DCF 權重）", ["tp_w_dcf", "tp_wA", "tp_wB", "tp_wC"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "rf_us", "beta_u", "beta_u_med", "beta_de", "beta_du", "beta", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "bs_date", "stub_frac", "stub_div", "sh_v"]],
    ["非控制權益（環旭電子市值、其餘帳面）", ["usi_sh", "usi_own", "usi_eq", "usi_px", "fx_cnytwd", "nci_usi_bk", "nci_usi_mv"]],
    ["自由現金流（折舊、資本支出、營運資金）", ["fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]],
    ["收斂法（FY2030–FY2032）", ["tp_g", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]],
    ["Wattonomics 錨定（只進敏感度）", ["wta_27", "wta_28", "wta_29"]]]]],
 "blocks_to": {"AI 加速器 → 日月光 LEAP": "封測（ATM）：LEAP 先進封裝（AI 加速器 × 委外份額 × 每顆內容＋全流程＋CPO）", "@Units!A16": "封測（ATM）：LEAP 先進封裝（AI 加速器 × 委外份額 × 每顆內容＋全流程＋CPO）"},
 "main": {"DCF 彙總": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "＋Stub 期 FCFF", "淨現金滾到評價日", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
          "DCF 折現率": ["股權成本", "稅後債務成本", "權益占資本", "WACC", "永續成長率", "永續期 RONIC"],
          "目標價結論": ["目標價（DCF", "DCF 目標價", "Comparable 目標價（相對", "Comparable 目標價（市場", "現價", "上檔空間：加權", "上檔空間：Comparable", "Comparable（相對倍數調整後）的上檔空間", "投資建議", "DCF 情境價值推到", "市場共識"],
          "本模型與公司指引、公司說法、券商的對照": ["4Q26 合併營收季增", "FY2027 LEAP ÷ FY2026", "FY2027 資本支出", "FY2027 EPS vs BofA", "FY2028 LEAP ÷ 封測營收"]},
 "n_recon": 4,
}

out = RL.relayout(cfg, X, SPEC)

# 補充規則 A：橫幅中間格＝只看 Comparable（市場主流同業法，未經相對倍數調整）＋自己的上檔空間；附註寫進加權的調整後值
h = out["hero"]
h["tiles"][1] = ["只看 Comparable 目標價（市場主流同業法）", f"@Val!B{comp}|n2", ["上檔空間 ", f"@Val!B{upc}|sp1"]]
if abs(WB["Val"].cell(compadj, 2).value - WB["Val"].cell(comp, 2).value) > 1e-9:
    h["note"] = (h.get("note") or []) + ["；進加權的 Comparable 為相對倍數調整後 ", f"@Val!B{compadj}|n2", " 元。"]

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

# 三層協調 C4（規則 v1.4；tools/coord/coord.py）：橫幅另一側讀數一格與揭露句、估值章開頭兩讀數表與長條圖
sys.path.insert(0, os.path.join(HERE, "..", "tools", "coord"))
import coord as COORD  # noqa: E402
COORD.inject_html(out, X)
cfg = out


# ================= v0.2.6（2026-11-14）：個股 HTML 版面規格 v1.3（Andy 2026-11-14 08:11、09:03、09:16、12:47）＋三層協調 C0 =================
# 章節：命題 → 估值 → 敏感度 → 反向 DCF → 損益（標準損益表）→ 驅動因子（分部營收推導）→ 對帳（含三層協調 C0）→ 資料狀態；日月光無個別模組。
# 數字與白話文字一律取自 xlsx（README 頁末 v1.3 白話列、PL 頁末標準損益表、RevDCF 頁、Val 0.7、Sens 頁末相對基準表、Coord 頁；ase/v026_post.py 產生）。列號依標籤找。
def _v13():
    S = {x["id"]: x for x in cfg["sections"]}
    RD, PLW, RVW, CW, SEW = WB["README"], WB["PL"], WB["RevDCF"], WB["Coord"], WB["Sens"]
    r13 = row_of("README", "v1.3 展示層")
    RM = dict(zip(["hero", "rec", "o1", "o2", "o3", "o4", "o5", "drv", "v25", "ver"], range(r13 + 1, r13 + 11)))
    assert str(RD[f"A{RM['ver']}"].value).startswith("v0.2.6"), RD[f"A{RM['ver']}"].value
    cfg["layout"] = "v1.3"
    cr = lambda p, start=1: next(r for r in range(start, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith(p))
    # ---- 橫幅：只留 ≤200 字白話與對帳中揭露；規則與歸因移到估值章 ----
    old_note = cfg["hero"]["note"]
    cfg["hero"]["note"] = [f"@README!A{RM['hero']}"] + list(old_note[-2:])          # 最後兩項＝" "、Coord 橫幅揭露
    for t in cfg["hero"]["tiles"]:
        t[0] = [t[0], "（新台幣）"] if isinstance(t[0], str) else list(t[0]) + ["（新台幣）"]
    # ---- 1. 命題與核心問題：兩張圖（LEAP 占比、資本支出）、投資建議條件式、重點觀察 ≤5 點 ----
    pa, pb, pc = (row_of("PL", f"情境 {k}", exact=True) for k in "ABC")
    bA, bB, bC = (row_of("PL", f"情境 {k}：") for k in "ABC")
    ex = lambda sh, label, start: next(r for r in range(start, WB[sh].max_row + 1) if str(WB[sh].cell(r, 1).value or "").strip() == label)
    hB = bB + 1
    leapB, mainB, emsB, revB = ex("PL", "LEAP 先進封裝（命題子列）", bB), ex("PL", "主流封測", bB), ex("PL", "EMS（環旭）", bB), ex("PL", "營收", bB)
    shr = {k: ex("PL", "LEAP ÷ 封測營收", b) for k, b in zip("ABC", (bA, bB, bC))}
    capB, depB, cpiB, opmB = ex("PL", "資本支出（新台幣）", bB), ex("PL", "折舊攤提", bB), ex("PL", "資本支出 ÷ 營收", bB), ex("PL", "營業利益率", bB)
    th = S["thesis"]["blocks"]
    c_leap = {"type": "theme", "title": "營收分部（情境 B，十億新台幣）與 LEAP 占封測營收（三情境）：FY2025A–FY2029", "x": f"@PL!C{hB}:G{hB}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.4,
              "total": f"@PL!C{revB}:G{revB}",
              "bars": [{"name": ["LEAP 先進封裝"], "y": f"@PL!C{leapB}:G{leapB}", "color": "#0b7a80"}, {"name": [f"@PL!A{mainB}"], "y": f"@PL!C{mainB}:G{mainB}", "color": "#6fc7cb"},
                       {"name": [f"@PL!A{emsB}"], "y": f"@PL!C{emsB}:G{emsB}", "color": "#c29a2b"}],
              "lines": [{"name": [f"@PL!A{pa}", " LEAP 占封測營收"], "y": f"@PL!C{shr['A']}:G{shr['A']}", "axis": "r", "color": "#0b7a80", "dash": True},
                        {"name": [f"@PL!A{pb}", " LEAP 占封測營收"], "y": f"@PL!C{shr['B']}:G{shr['B']}", "axis": "r", "color": "#3b3838", "labels": True},
                        {"name": [f"@PL!A{pc}", " LEAP 占封測營收"], "y": f"@PL!C{shr['C']}:G{shr['C']}", "axis": "r", "color": "#b4472f", "dash": True}],
              "caption": "堆疊長條＝情境 B 的 LEAP 先進封裝、主流封測、EMS（環旭）；頂端數字＝合併營收（沖銷及其他為小額負數，未畫）；三條線＝三情境的 LEAP 占封測營收（右軸）。回答核心問題一：LEAP 升到三成以上後能守住多少。"}
    c_cx = {"type": "theme", "title": "資本支出與折舊（情境 B，十億新台幣）、資本支出占營收與營業利益率：FY2025A–FY2029", "x": f"@PL!C{hB}:G{hB}", "fmt": "n0", "fmt_r": "p0", "stack": False, "pad": 1.3,
            "bars": [{"name": [f"@PL!A{capB}"], "y": f"@PL!C{capB}:G{capB}", "color": "#c29a2b"}, {"name": [f"@PL!A{depB}"], "y": f"@PL!C{depB}:G{depB}", "color": "#6fc7cb"}],
            "lines": [{"name": [f"@PL!A{cpiB}"], "y": f"@PL!C{cpiB}:G{cpiB}", "axis": "r", "color": "#b4472f", "labels": True},
                      {"name": [f"@PL!A{opmB}"], "y": f"@PL!C{opmB}:G{opmB}", "axis": "r", "color": "#0b7a80"}],
            "caption": "長條＝資本支出與折舊攤提（新台幣）；折線＝資本支出占營收與營業利益率（右軸）。回答核心問題二：資本支出約占營收三到四成時，折舊追上來、利益率下滑。"}
    i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
    th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RM[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
    S["thesis"]["blocks"] = [c_leap, c_cx, {"type": "note", "parts": [f"@README!A{RM['rec']}"]}] + th
    # ---- 2. 估值：四個目標價圖＋表置頂（G 節）；章節說明、評等規則、對帳中兩讀數區其後 ----
    va = S["valuation"]["blocks"]
    g0 = row_of("Val", "0.7 估值並列"); gr = list(range(g0 + 2, g0 + 7)); r_px = row_of("Val", "現價（2026")
    bars4 = {"type": "bars", "title": "四個目標價與現價（新台幣；豎線＝現價）", "fmt": "n2",
             "items": [{"label": lb, "value": f"@Val!B{r}", "series": sr} for r, sr, lb in zip(gr, (1, 2, 2, 0, 1),
                       ("DCF 目標價", "Comparable（調整後，進加權）", "Comparable（未調整，不進加權）", "加權目標價（橫幅主值）", "市場共識目標價（BofA 單一券商）"))],
             "ref": {"label": f"@Val!A{r_px}", "value": f"@Val!B{r_px}"}}
    tbl4 = {"type": "table", "caption": f"@Val!A{g0}", "head": [f"@Val!{c}{g0 + 1}" for c in "ABCD"],
            "rows": [[f"@Val!A{r}", f"@Val!B{r}|n2", f"@Val!C{r}|sp1", f"@Val!D{r}"] for r in gr] + [[f"@Val!A{g0 + 7}", f"@Val!B{g0 + 7}|n2", "", f"@Val!D{g0 + 7}"]]}
    rules = {"type": "note", "parts": list(old_note[:-2])}
    h_i = next(i for i, b in enumerate(va) if b.get("type") == "h3" and str(b.get("text", "")).startswith("對帳中"))
    coord_blocks = va[h_i:h_i + 4]
    assert [b["type"] for b in coord_blocks] == ["h3", "note", "table", "bars"], [b["type"] for b in coord_blocks]
    old_bars = next(b for b in va if b.get("type") == "bars" and b.get("title") == "目標價與現價")
    rest = [b for b in va if b not in coord_blocks and b is not old_bars]
    top = [bars4]
    if S["valuation"].get("intro"):
        top.append({"type": "note", "parts": [S["valuation"].pop("intro")]})
    S["valuation"]["blocks"] = top + [tbl4, rules] + coord_blocks + rest
    # ---- 3. 敏感度：加權目標價相對基準的低案、高案主圖置頂（H 節）----
    t0 = row_of("Sens", "加權目標價相對基準（HTML"); r1 = t0 + 2; r2 = r1
    while SEW.cell(r2 + 1, 1).value: r2 += 1
    tor = []
    for col, side in (("B", "低"), ("C", "高")):
        big = [r for r in range(r1, r2 + 1) if isinstance(SEW[f"{col}{r}"].value, (int, float)) and abs(SEW[f"{col}{r}"].value) >= 0.1]
        small = [r for r in range(r1, r2 + 1) if isinstance(SEW[f"{col}{r}"].value, (int, float)) and 0.0005 <= abs(SEW[f"{col}{r}"].value) < 0.1]
        base = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1"}
        tor.append((dict(base, title=f"加權目標價相對基準：輸入取{side}值，變動達兩位數百分比的案例", rows=",".join(map(str, big))),
                    dict(base, title=f"加權目標價相對基準：輸入取{side}值", rows=",".join(map(str, small)), collapse=True,
                         summary=f"展開：加權目標價相對基準：輸入取{side}值，變動未達兩位數百分比的案例（幾乎為零者見全部案例表）")))
    all_tbl = {"type": "table", "caption": f"@Sens!A{t0}", "collapse": True, "summary": "展開：全部案例的加權目標價相對基準（低案、高案）",
               "head": [f"@Sens!{c}{t0 + 1}" for c in "ABC"], "rows": [[f"@Sens!A{r}", f"@Sens!B{r}|sp1", f"@Sens!C{r}|sp1"] for r in range(r1, r2 + 1)]}
    sb = S["sens"]["blocks"]
    old_w = [b for b in sb if b.get("type") == "tornado" and str(b.get("title", "")).startswith("加權目標價相對基準的變動")]
    others = []
    for b in sb:
        if b in old_w: continue
        if b.get("type") == "tornado" and str(b.get("title", "")).startswith("FY2028 EPS") and not b.get("collapse"):
            b = dict(b, collapse=True, summary="展開：對照：FY2028 EPS（情境 B）相對基準的變動達兩位數百分比的案例")
        if b.get("type") == "tornado" and str(b.get("title", "")).startswith("FY2028 EPS") and b.get("collapse") and "對照" not in str(b.get("summary", "")):
            b = dict(b, summary="展開：對照：FY2028 EPS（情境 B）相對基準的變動未達兩位數百分比的案例")
        others.append(b)
    intro = S["sens"].pop("intro", None)
    lead = [tor[0][0], tor[1][0]]
    if intro: lead.append({"type": "note", "parts": [intro]})
    S["sens"]["blocks"] = lead + [tor[0][1], tor[1][1], all_tbl] + others
    # ---- 4. 反向 DCF（新）----
    rv = lambda p, start=1: next(r for r in range(start, RVW.max_row + 1) if str(RVW.cell(r, 1).value or "").startswith(p))
    ra = rv("A. DCF"); rb_ = rv("(a)"); rbb = rv("(b)"); rc = rv("(c)"); rcmp = rv("B. Comparable"); rsum = rv("C. 摘要"); yh = rv("項目（情境 B）")
    m1, m0 = rv("現價隱含的營業利益率", rb_), rv("模型：基準的等效營業利益率", rb_)
    s1, rv1, rv0, g1, g0_ = rv("現價隱含的營收倍數", rbb), rv("現價隱含的 FY2027 營收", rbb), rv("模型：FY2027 營收", rbb), rv("現價隱含的 FY2027 營收年增率", rbb), rv("模型：FY2027 營收年增率", rbb)
    w1, w0 = rv("現價隱含的 WACC", rc), rv("模型 WACC", rc)
    kvr = [rv(p, ra) for p in ("現價（", "支持現價所需的評價日", "淨現金滾到評價日", "評價用稀釋股數", "支持現價所需的企業價值", "模型企業價值 EV", "EV 需放大", "模型：情境 B 評價日")]
    kvf = ["n2", "n2", "n1", "n3", "n1", "n1", "x1", "n2"]
    cm = [rv(p, rcmp) for p in ("12 個月後前瞻 EPS", "只看 Comparable：支持現價所需的前瞻", "只看 Comparable：所需的相對", "加權（DCF 不變）", "同業前瞻本益比中位數")]
    ylab = ["營收（", "營業利益 EBIT", "營業利益率", "稅率", "NOPAT＋折舊", "營運資金增加", "現值（"]
    yrr = [rv(p, yh) for p in ylab]
    yfm = ["n1", "n1", "p1", "p1", "n1", "n1", "n1"]
    S["revdcf"] = {
        "id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2",
        "blocks": [
            {"type": "note", "parts": [f"@RevDCF!A{rsum + 1}"]},
            {"type": "note", "parts": [f"@RevDCF!A{rsum + 2}"]},
            {"type": "table", "caption": "現金流折現法（情境 B 路徑）：現價隱含值與模型基準", "head": ["項目", "現價隱含", "模型基準"],
             "rows": [["預測期營業利益率（FY2027–FY2032 統一一個值；模型欄為等效值，終值年權重大）", f"@RevDCF!B{m1}|p1", f"@RevDCF!B{m0}|p1"],
                      ["營收規模（FY2027 起各年同乘的倍數）", f"@RevDCF!B{s1}|n2", "—"],
                      ["FY2027 營收（十億新台幣）", f"@RevDCF!B{rv1}|n0", f"@RevDCF!B{rv0}|n0"],
                      ["FY2027 營收年增率（相對 FY2026 錨點）", f"@RevDCF!B{g1}|p0", f"@RevDCF!B{g0_}|p1"],
                      ["折現率（WACC）", f"@RevDCF!B{w1}|p2", f"@RevDCF!B{w0}|p2"]]},
            {"type": "table", "caption": f"@RevDCF!A{ra}", "head": ["項目", "值"], "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|{f}"] for r, f in zip(kvr, kvf)]},
            {"type": "table", "caption": f"@RevDCF!A{rcmp}", "head": [f"@RevDCF!{c}{rcmp + 1}" for c in "ABCD"],
             "rows": [[f"@RevDCF!A{cm[0]}", "—", f"@RevDCF!C{cm[0]}|n2", "—"]]
                     + [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|" + ("n2" if k == 1 else "n1"), f"@RevDCF!C{r}|" + ("n2" if k == 1 else "n1"), f"@RevDCF!D{r}|sp1"] for k, r in enumerate(cm[1:4])]
                     + [[f"@RevDCF!A{cm[4]}", "—", f"@RevDCF!C{cm[4]}|n1", "—"]]},
            {"type": "table", "caption": "情境 B 各年數字（反解係數的輸入；十億新台幣）", "collapse": True, "summary": "展開：情境 B 各年營收、營業利益與現值（封閉解的計算輸入）",
             "head": [f"@RevDCF!{c}{yh}" for c in "ABCDEFG"],
             "rows": [[f"@RevDCF!A{r}"] + [f"@RevDCF!{c}{r}|{f}" for c in "BCDEFG"] for r, f in zip(yrr, yfm)]},
        ]}
    # ---- 5. 損益（標準損益表）----
    p0 = row_of("PL", "標準損益表（v0.2.6")
    ph = next(r for r in range(p0, PLW.max_row + 1) if PLW.cell(r, 1).value == "項目")
    plr = lambda lab_, start=ph: next(r for r in range(start, PLW.max_row + 1) if PLW.cell(r, 1).value == lab_)
    keys = ["營收", "營收年增率（YoY）", "營業成本（COGS）", "毛利", "毛利率", "營業費用（研發＋銷管）", "營業利益", "營業利益率", "其他收入（淨額；營業外淨額）", "稅前利益", "所得稅",
            "稅後淨利（合併，含非控制權益）", "非控制權益損益", "歸屬母公司淨利（EPS 與目標價口徑）", "淨利率（歸屬母公司淨利 ÷ 營收）", "基本股數", "EPS（基本；NT$）"]
    fmt = lambda u: {"%": "p1", "NT$": "n2", "百萬股": "n0"}.get(u, "n1")
    COLS9 = "CDEFGHIJK"
    rowx = lambda r: [f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmt(PLW.cell(r, 2).value)}" for c in COLS9]
    r_ratio = plr("比率"); r_gaap = next(r for r in range(ph, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("GAAP 對照"))
    r_chk = plr("核對（應為 0）")
    r_d = next(r for r in range(r_chk, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("差額（FY2030E"))
    head = [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in COLS9]
    rp = cr("損益表年度標示")
    head[5:10] = [f"@Coord!{c}{rp}" for c in "BCDEF"]          # FY2027E–FY2031E 標「對帳中」
    is_tbl = {"type": "table", "caption": "標準損益表（十億新台幣；EPS 為新台幣）：FY2023A–FY2025A 實績、FY2026E–FY2031E 情境 B（基準）；GAAP（IFRS 合併，歸屬母公司）", "head": head,
              "rows": [rowx(plr(k)) for k in keys] + [{"sub": f"@PL!A{r_ratio}"}] + [rowx(r) for r in range(r_ratio + 1, r_gaap)]
                      + [{"sub": "核對：FY2030E、FY2031E 兩種推法的營業利益差額（毛利率與營業費用率沿用 FY2029E）"}] + [rowx(r) for r in (r_d, r_d + 1)]}
    gaap_tbl = {"type": "table", "caption": f"@PL!A{r_gaap}", "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in "CDE"], "collapse": True, "summary": "展開：稀釋 EPS（公司公告，只列實績）",
                "rows": [[f"@PL!A{r}"] + [f"@PL!{c}{r}|n2" for c in "CDE"] for r in range(r_gaap + 1, r_chk)]}
    rv_, gm_, opm_, nm_ = plr("營收"), plr("毛利率"), plr("營業利益率"), plr("淨利率")
    is_chart = {"type": "theme", "title": "營收（十億新台幣）與毛利率、營業利益率、淨利率：FY2023A–FY2031E（情境 B）", "x": f"@PL!C{ph}:K{ph}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
                "bars": [{"name": f"@PL!A{rv_}", "y": f"@PL!C{rv_}:K{rv_}", "color": "#6fc7cb"}],
                "lines": [{"name": f"@PL!A{gm_}", "y": f"@PL!C{gm_}:K{gm_}", "axis": "r", "color": "#0b7a80", "labels": True},
                          {"name": f"@PL!A{opm_}", "y": f"@PL!C{opm_}:K{opm_}", "axis": "r", "color": "#c29a2b"},
                          {"name": f"@PL!A{nm_}", "y": f"@PL!C{nm_}:K{nm_}", "axis": "r", "color": "#6f6bae", "dash": True}],
                "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸；GAAP，歸屬母公司淨利）。FY2026E 為錨點；FY2030E、FY2031E 為延伸期：營收與淨利由收斂法推得，毛利率與營業費用率沿用 FY2029E。"}
    sc_chart = {"type": "theme", "title": "三情境並列：營收（長條，十億新台幣）與 EPS（折線，新台幣，右軸）", "x": "@PL!C3:I3", "fmt": "n0", "fmt_r": "n1", "stack": False, "pad": 1.1,
                "bars": [{"name": [f"@PL!A{pa}", " 營收"], "y": f"@PL!C{pa + 1}:I{pa + 1}", "color": "#b7e3e4"},
                         {"name": [f"@PL!A{pb}", " 營收"], "y": f"@PL!C{pb + 1}:I{pb + 1}", "color": "#10a0a7"},
                         {"name": [f"@PL!A{pc}", " 營收"], "y": f"@PL!C{pc + 1}:I{pc + 1}", "color": "#8a8c8c"}],
                "lines": [{"name": [f"@PL!A{pa}", " EPS"], "y": f"@PL!C{pa + 13}:I{pa + 13}", "axis": "r", "color": "#0b7a80", "dash": True},
                          {"name": [f"@PL!A{pb}", " EPS"], "y": f"@PL!C{pb + 13}:I{pb + 13}", "axis": "r", "color": "#3b3838", "labels": True},
                          {"name": [f"@PL!A{pc}", " EPS"], "y": f"@PL!C{pc + 13}:I{pc + 13}", "axis": "r", "color": "#b4472f", "dash": True}],
                "caption": "三情境在 FY2026 相同（錨點）；差異來自 AI 加速器顆數、委外與份額相對值、全流程與 CPO；FY2030 起為延伸期（收斂法）。"}
    assert str(PLW.cell(pa + 1, 1).value) == "營收" and str(PLW.cell(pa + 13, 1).value) == "EPS", (PLW.cell(pa + 1, 1).value, PLW.cell(pa + 13, 1).value)
    old_pl = S["pl"]["blocks"]
    pls = [b for b in old_pl if b["type"] == "pl"]
    kvs = [b for b in old_pl if b["type"] == "kv"]
    assert len(pls) == 2, len(pls)
    legacy = [dict(pls[0], collapse=True, summary="展開：三情境摘要表與折線（FY2025A–FY2031；營收、淨利率、EPS）"),
              dict(pls[1], collapse=True, summary="展開：三情境 LEAP 與封測毛利率（FY2025A–FY2029）")]
    legacy += [dict(b, collapse=True, summary="展開：FY2026 錨點核對（應為零）") for b in kvs]
    S["pl"].update(title="損益（標準損益表）", blocks=[{"type": "note", "parts": [f"@PL!A{p0 + 1}"]}, is_chart, is_tbl, {"type": "note", "parts": [f"@PL!A{p0 + 2}"]}, sc_chart, gaap_tbl] + legacy)
    S["pl"].pop("intro", None)
    # ---- 6. 驅動因子：分部營收推導（情境 B）----
    lc, lf, lcp, lt, un = (ex("PL", t, bB) for t in ("LEAP 核心（2.5D oS、先進測試等）", "全流程（CoWoS full process）", "CPO（共同封裝光學）", "LEAP 合計", "AI 加速器顆數（含情境係數）"))
    rows_seg = [{"sub": "一、LEAP 先進封裝（美元）＝顆數 × 封裝面積 × 委外與份額 × 每顆核心內容，加全流程與 CPO"}] + \
               [ex("PL", t, bB) for t in ("AI 加速器顆數（含情境係數）", "封裝面積指數（2026＝1）", "委外與份額相對值（2026＝1）", "每顆加速器 LEAP 核心內容（不含全流程、CPO）", "LEAP 核心（2.5D oS、先進測試等）",
                                           "全流程（CoWoS full process）", "CPO（共同封裝光學）", "LEAP 合計", "匯率")] + \
               [{"sub": "二、營收分部（十億新台幣）：LEAP＋主流封測＝封測；封測＋EMS＋沖銷＝合併營收"}] + \
               [ex("PL", t, bB) for t in ("LEAP 先進封裝（命題子列）", "主流封測", "封測（ATM）", "EMS（環旭）", "沖銷及其他", "營收", "營收年增", "LEAP ÷ 封測營收", "封測年增")]
    seg_f = {"百萬顆": "n2", "x": "n2", "US$／顆": "n0", "US$B": "n2", "NT$/US$": "n1", "NT$B": "n1", "%": "p1"}
    seg_rows = [r if isinstance(r, dict) else [f"@PL!A{r}", f"@PL!B{r}"] + [f"@PL!{c}{r}|{seg_f.get(str(PLW.cell(r, 2).value), 'n1')}" for c in "CDEFG"] for r in rows_seg]
    seg_tbl = {"type": "table", "caption": "分部營收推導（情境 B）：驅動因子 → LEAP → 營收分部 → 合併營收", "head": [f"@PL!A{hB}", f"@PL!B{hB}"] + [f"@PL!{c}{hB}" for c in "CDEFG"], "rows": seg_rows}
    seg_chart = {"type": "theme", "title": "LEAP 先進封裝推導（情境 B，十億美元）與 AI 加速器顆數（百萬顆，右軸）：FY2026–FY2029", "x": f"@PL!D{hB}:G{hB}", "fmt": "n1", "fmt_r": "n1", "stack": True, "pad": 1.3,
                 "total": f"@PL!D{lt}:G{lt}",
                 "bars": [{"name": [f"@PL!A{lc}"], "y": f"@PL!D{lc}:G{lc}", "color": "#0b7a80"}, {"name": [f"@PL!A{lf}"], "y": f"@PL!D{lf}:G{lf}", "color": "#6fc7cb"},
                          {"name": [f"@PL!A{lcp}"], "y": f"@PL!D{lcp}:G{lcp}", "color": "#c29a2b"}],
                 "lines": [{"name": [f"@PL!A{un}"], "y": f"@PL!D{un}:G{un}", "axis": "r", "color": "#3b3838", "labels": True}],
                 "caption": "堆疊＝LEAP 核心、全流程、CPO（美元）；頂端數字＝LEAP 合計；折線＝AI 加速器顆數（含情境係數，右軸）。FY2026 為錨點年（依公司「逾四十億美元」校準），FY2027 起由顆數推導。"}
    dr = S["drivers"]["blocks"]
    S["drivers"]["blocks"] = [{"type": "note", "parts": [f"@README!A{RM['drv']}"]}, seg_chart, seg_tbl, {"type": "h3", "text": "情境驅動與單一參數"}] + dr
    S["drivers"]["title"] = "驅動因子：分部營收推導"
    S["drivers"].pop("intro", None)
    # ---- 7. 對帳：三層協調 C0（Coord 頁）＋公司指引對照 ----
    cb = cr("B. C0"); ch = cb + 1; r_f, r_l = cr("AI 加速器顆數（情境 B）CY2026"), cr("產業部署 GW CY2029")
    r_nc = cr("日月光 LEAP 營收（FY2026"); r_br0, r_br1 = cr("參考：LEAP 核心營收 ÷ 層級委外補充 CoWoS 月產能 CY2026"), cr("參考：LEAP 核心營收 ÷ 層級委外補充 CoWoS 月產能 CY2029")
    r_sum = cr("判定摘要"); r_c1, r_c2, r_c3 = cr("對帳中年份：CY"), cr("口徑註記"), cr("C0 結論")
    ca = cr("A. 讀哪一層"); cah = ca + 1; r_a1, r_a2 = cah + 1, cr("Wattonomics：本模型不直接讀")
    gfm = lambda r: "n2" if "GW" in str(CW.cell(r, 1).value) else "n3"
    c0_tbl = {"type": "table", "caption": f"@Coord!A{cb}", "head": [f"@Coord!{c}{ch}" for c in "ABCDEF"],
              "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|{gfm(r)}", f"@Coord!C{r}|{gfm(r)}", f"@Coord!D{r}|sp1", f"@Coord!E{r}", f"@Coord!F{r}" if r in (r_f, r_f + 4) else ""] for r in range(r_f, r_l + 1)]
                      + [[f"@Coord!A{r_nc}", f"@Coord!B{r_nc}", f"@Coord!C{r_nc}", f"@Coord!D{r_nc}", f"@Coord!E{r_nc}", f"@Coord!F{r_nc}"]]}
    br_tbl = {"type": "table", "collapse": True, "summary": "展開：橋接參考（日月光 LEAP 核心營收 ÷ 層級委外補充 CoWoS 月產能；不判定）", "head": [f"@Coord!{c}{ch}" for c in "ABCDE"],
              "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|n3", f"@Coord!C{r}|n1", f"@Coord!D{r}", f"@Coord!E{r}"] for r in range(r_br0, r_br1 + 1)]}
    a_tbl = {"type": "table", "collapse": True, "summary": "展開：本模型讀哪一層、哪一版（Siliconomics 版本與讀入項目）", "caption": f"@Coord!A{ca}", "head": [f"@Coord!{c}{cah}" for c in "ABCDEF"],
             "rows": [[f"@Coord!A{r}"] + [f"@Coord!{c}{r}|{'n2' if '顆數' in str(CW.cell(r, 1).value) else 'p0' if 'CPO' in str(CW.cell(r, 1).value) else 'n3'}" for c in "BCDE"] + [f"@Coord!F{r}"] for r in range(r_a1, r_a2)]}
    c0_blocks = [{"type": "h3", "text": "三層協調：由下往上對帳（本模型的產業量 vs 層級對應值）"}, {"type": "note", "parts": [f"@Coord!B{r_sum}"]}, c0_tbl,
                 {"type": "note", "parts": [f"@Coord!A{r_c1}"]}, {"type": "note", "parts": [f"@Coord!A{r_c2}"]}, {"type": "note", "parts": [f"@Coord!A{r_c3}"]}, br_tbl, a_tbl,
                 {"type": "note", "parts": [f"@Coord!A{r_a2}"]}, {"type": "h3", "text": "本模型與公司指引、公司說法、券商的對照"}]
    S["recon"]["blocks"] = c0_blocks + S["recon"]["blocks"]
    # ---- 9. 資料狀態：本版要點 ----
    S["status"]["blocks"][0:0] = [{"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RM['ver']}"]}]
    order = ["thesis", "valuation", "sens", "revdcf", "pl", "drivers", "recon", "status"]
    cfg["sections"] = [S[k] for k in order]


_v13()

json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")
