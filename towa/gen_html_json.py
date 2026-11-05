# v0.3（2026-11-05）：TOWA 首次套用「個股 HTML 版面規格 v1.0」（docs/plan/個股HTML版面規格.md），同時 DCF 改標準版（FCFF＋WACC）。寫法比照 ajinomoto/gen_html_json.py。
# 由 html_src.json（v0.2.1 的 html.json，內容不動）產生 towa/html.json；估值章 DCF 區塊、Sens 目標價參數表與情境驅動表的列號由本檔依標籤重產。
# 用法：python3 towa/gen_html_json.py [成品 xlsx]（省略時取 towa/dist/ 內唯一的 xlsx）；勿手改 html.json。
import copy, glob, json, os, re, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools", "company_html"))
from relayout import relayout  # noqa: E402

X = sys.argv[1] if len(sys.argv) > 1 else (lambda f: f[0] if len(f) == 1 else sys.exit(f"dist 內應剛好一個 xlsx：{f}"))(glob.glob(os.path.join(HERE, "dist", "*.xlsx")))
OUT = os.path.join(HERE, "html.json")
SRC = json.load(open(os.path.join(HERE, "html_src.json"), encoding="utf-8"))
cfg = copy.deepcopy(SRC)
wb = openpyxl.load_workbook(X, data_only=True)


def find(sh, prefix, start=1, col=1):
    ws = wb[sh]
    for r in range(start, ws.max_row + 1):
        if str(ws.cell(r, col).value or "").strip().startswith(prefix):
            return r
    raise KeyError((sh, prefix))


def sec(cid, c=None):
    return next(s for s in (c or cfg)["sections"] if s["id"] == cid)


def txt(ref):
    m = re.match(r"^@([^!]+)!([A-Z]+\d+)", str(ref))
    return str(wb[m.group(1)][m.group(2)].value or "") if m else str(ref or "")


V = lambda r, f="": f"@Val!B{r}" + (f"|{f}" if f else "")
VA = lambda r: f"@Val!A{r}"


# ---- 前處理零：情境驅動表的 In 列號依「情境驅動對照」一節實際位置平移（v0.3 在其上方新增 M2 節）----
def _shift_scn(blocks):
    for i_, b in enumerate(blocks):
        if b.get("type") == "table" and str(b.get("caption", "")).startswith("情境驅動"):
            old_h = int(re.search(r"In!B(\d+)", json.dumps(b["head"])).group(1))
            new_h = find("In", "情境驅動對照") + 1
            d = new_h - old_h
            if d:
                blocks[i_] = json.loads(re.sub(r"In!([A-Z]+)(\d+)", lambda m: f"In!{m.group(1)}{int(m.group(2)) + d}" if int(m.group(2)) >= old_h else m.group(0), json.dumps(b, ensure_ascii=False)))
            nb_ = len(blocks[i_]["rows"])
            assert wb["In"].cell(new_h + nb_, 1).value and not wb["In"].cell(new_h + nb_ + 1, 1).value, ("情境驅動表列數與 In 頁不符", new_h, nb_)
_shift_scn(sec("drivers")["blocks"])
for b in sec("drivers")["blocks"]:
    if str(b.get("caption", "")).startswith("單一參數"):
        b["caption"] = "驅動 → 推導對照表（In 頁基準值與 PL 頁推導量）"

# ---- 前處理一：估值章（重點觀察已在第一章；kv 改為有標題的表；DCF 前加折現率表）----
va = sec("valuation")
nb, old, i = [], va["blocks"], 0
while i < len(old):
    b = old[i]
    nxt = old[i + 1] if i + 1 < len(old) else {}
    if b["type"] == "h3" and b.get("text") == "重點觀察" and nxt.get("type") == "list":
        i += 2
        continue
    if b["type"] == "h3" and str(b.get("text", "")).startswith("相對倍數調整") and nxt.get("type") == "kv":
        nb.append({"type": "table", "caption": b["text"], "head": ["項目", "數值"], "rows": [list(x) for x in nxt["items"]]})
        i += 2
        continue
    if b["type"] == "kv":
        lab0 = txt(b["items"][0][0])
        cap = ("Comparable 目標價的計算（市場主流同業法：後段設備同業前瞻 P/E 中位數 × 前瞻 EPS）" if lab0.startswith("同業前瞻 P/E 中位數")
               else "舊法：規則值情境加權（參考，不是目標價）" if lab0.startswith("規則值") else "現價與目前倍數" if lab0.startswith("現價") else re.sub(r"[（(].*", "", lab0) + "等")
        nb.append({"type": "table", "caption": cap, "head": ["項目", "數值"], "rows": [list(x) for x in b["items"]]})
        i += 1
        continue
    if b["type"] == "bars":
        b = copy.deepcopy(b)
        for it in b.get("items", []):
            if str(it.get("label", "")).startswith("@Val!A") and not txt(it["label"]):
                r_ = find("Val", "DCF 目標價（情境加權）")
                it["label"], it["value"] = VA(r_), V(r_)
    if b["type"] == "table" and txt(b.get("caption", "")).startswith("0.1 DCF"):
        h = find("Val", "0.1 DCF") + 1
        ke, ronic, pvx, dcf = find("Val", "股權成本 ke"), find("Val", "永續期 RONIC"), find("Val", "明確期 FCFF 現值"), find("Val", "DCF 目標價（情境加權）")
        nb.append({"type": "table", "caption": "DCF 折現率（標準版 DCF：WACC、股權成本、稅後債務成本、RONIC、永續成長率）", "head": ["項目", "數值", "說明"],
                   "rows": [[VA(r), V(r, "p2"), f"@Val!E{r}"] for r in range(ke, ronic + 1)]})
        pct = {str(find("Val", "其中：終值現值占 EV")): {c: "p1" for c in "BCD"}, str(find("Val", "隱含終值 EV／EBITDA")): {c: "x1" for c in "BCD"},
               str(find("Val", "評價用稀釋股數")): {c: "n3" for c in "BCD"}}
        nb.append({"type": "table", "caption": f"@Val!A{h - 1}", "head": [f"@Val!{c}{h}" for c in "ABCDE"],
                   "rowgen": {"sheet": "Val", "rows": f"{pvx}-{dcf}", "cols": list("ABCDE"), "fmts": ["", "n0", "n0", "n0", ""], "row_fmt": pct}})
        b0 = find("DCF", "情境 B（")
        b1 = find("DCF", "檢查：終值占 EV", start=b0)
        hd = find("DCF", "項目")
        lab = lambda r: str(wb["DCF"].cell(r, 1).value or "")
        nb.append({"type": "table", "collapse": True, "summary": "展開：情境 B 的無槓桿自由現金流明細（DCF 頁；情境 A、C 同式）", "caption": f"@DCF!A{b0}",
                   "head": [f"@DCF!{c}{hd}" for c in "ACDEFGHI"],
                   "rowgen": {"sheet": "DCF", "rows": f"{b0 + 1}-{b1}", "cols": list("ACDEFGHI"), "fmts": ["", "n1", "n1", "n1", "n1", "n1", "n1", "n1"],
                              "row_fmt": {str(r): {c: ("n4" if "因子" in lab(r) else "p1") for c in "CDEFGHI"} for r in range(b0 + 1, b1 + 1) if any(k in lab(r) for k in ("率", "÷", "占", "因子"))}}})
        i += 1
        continue
    if b["type"] == "table" and txt(b.get("caption", "")).startswith("0.2 Comparable"):
        b = dict(b, collapse=True, summary="展開：後段設備同業的前瞻 P/E")
    nb.append(b)
    i += 1
va["blocks"] = nb

# ---- 前處理二：對帳表加標題；敏感度「目標價參數」表的列號改為 Sens 頁實際位置（標準版 DCF 案例）----
for b in sec("recon")["blocks"]:
    if b["type"] == "table" and not b.get("caption"):
        b["caption"] = "與公司指引、共識的對照（只作事後對照）"
for b in sec("sens")["blocks"]:
    if b["type"] == "table" and str(b.get("caption", "")).startswith("@Sens!A") and "目標價參數" in txt(b["caption"]):
        t0 = find("Sens", "目標價參數")
        t1 = find("Sens", "檢查：閉式公式的基準", start=t0) - 1
        b["caption"] = f"@Sens!A{t0}"
        b["head"] = [f"@Sens!{c}{t0 + 1}" for c in "ABCDEFG"]
        b["rowgen"] = {"sheet": "Sens", "rows": f"{t0 + 2}-{t1}", "cols": list("ABCDEFG"), "fmts": ["", "", "n0", "sp1", "n0", "sp1", "n0"]}

# ---- 轉換：版面規格 v1.0 ----
scn_tok = re.split(r"[.\s]", str(wb["In"].cell(find("In", "情境驅動對照"), 1).value))[0]
B_AI, B_NAI, B_CO = "AI 封膠：HBM 與 CoWoS 等先進封裝", "非 AI 封膠、其他半導體與其他事業", "全公司：錨點、利潤率、股價股數、折現率、目標價規則與現金流"
SPEC = {
 "in_skip": [scn_tok],
 "in_groups": [
  [B_AI, [
    ["AI 錨點（校準；公司不揭露）", ["ai_25", "w_hbm"]],
    ["HBM 出貨位元（讀 skhynix 成品；情境 A／B／C）", [f"hs{s_}_{y}" for s_ in "ABC" for y in range(25, 30)]],
    ["晶粒容量、堆疊層數與汰換", [f"den_{y}" for y in range(25, 30)] + [f"stk_{y}" for y in range(25, 30)] + ["dlt", "g_tail"]],
    ["CoWoS 片數（讀台積電模型）與情境顆數係數", [f"cw_{y}" for y in range(25, 30)] + ["um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C"]],
    ["AI 封膠市占流失（情境）", ["shl_A", "shl_B", "shl_C"]],
    ["Wattonomics 錨定係數（只用於敏感度）", ["wta_27", "wta_28", "wta_29"]]]],
  [B_NAI, [
    ["SEMI 組裝與封裝設備與非 AI 設備比（讀 markets/equipment）", [f"pk_{y}" for y in range(25, 30)] + [f"eq{t}{s_}_{y}" for s_ in "ABC" for t in ("t", "a") for y in range(25, 30)] + ["g_pk_tail"]],
    ["其他事業成長（新事業、醫療）", ["g_tss", "g_med"]],
    ["事業別營收實績（FY2025）", ["seg_semi", "seg_mold", "seg_osemi", "seg_new", "seg_med", "seg_las", "seg_adj"]]]],
  [B_CO, [
    ["FY2025 實績與訂單", ["a25_rev", "a25_gp", "a25_oi", "a25_ni", "a25_eps", "ord_25", "bl_25", "sh"]],
    ["FY2026 實績、公司指引與錨點", ["q1_rev", "q1_oi", "q1_ni", "q1_eps", "h1g_rev", "h1g_oi", "g26_rev", "g26_oi", "g26_ni", "g26_eps", "g26_semi", "g26_las", "g_anc"]],
    ["利潤率與股數", ["im", "ima_A", "ima_B", "ima_C", "nioi", "shares"]],
    ["共識與市場倍數", ["cons_rev26", "cons_oi26", "cons_ni26", "cons_tp", "cons_feps", "sa_fpe", "sa_ttm", "px"]],
    ["循環位置（設備市場見頂年）", ["pk_A", "pk_B", "pk_C"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "beta", "beta_u", "beta_u_med", "beta_de", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "sh_v"]],
    ["自由現金流（折舊、資本支出、營運資金）", ["fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]],
    ["目標價規則、情境權重與收斂期", ["tp_w_dcf", "tp_wA", "tp_wB", "tp_wC", "tp_g", "nim10"]]]]],
 "blocks_to": {},
 "main": {
  "0.1 DCF": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
  "DCF 折現率": ["股權成本", "稅後債務成本", "權益占資本", "WACC", "永續成長率", "永續期 RONIC"],
 },
}
if sec("drivers").get("intro"):
    SPEC["blocks_to"][sec("drivers")["intro"]] = B_AI
cfg = relayout(cfg, X, SPEC)

# ---- 後處理一：橫幅中間格＝只看 Comparable（市場主流同業法、未做相對倍數調整；工作單補充規則 A）；進加權的調整後值寫在附註 ----
v_cmp0, v_up0 = find("Val", "Comparable 目標價（市場主流同業法）"), find("Val", "上檔空間：Comparable 目標價")
v_adj = find("Val", "Comparable 目標價（相對倍數調整後）")
rd_upa = find("README", "上檔空間：Comparable 目標價（相對倍數調整後）", find("README", "HTML 橫幅輔助值"))
cfg["hero"]["tiles"][1] = ["只看 Comparable 目標價（市場主流同業法）", V(v_cmp0, "n0"), ["上檔空間 ", V(v_up0, "sp1")]]
cfg["hero"]["note"] = list(cfg["hero"].get("note") or []) + ["；進加權的 Comparable 為相對倍數調整後 ", V(v_adj, "n0"), "（上檔空間 ", f"@README!B{rd_upa}|sp1", "）"]

# ---- 後處理二：龍捲風圖是 EPS 的變動（情境 A），不是加權目標價 ----
_ytit = {"H": "FY2028 EPS 相對基準的變動", "E": "FY2027 EPS 相對基準的變動", "K": "FY2026 EPS 相對基準的變動"}
for b in sec("sens")["blocks"]:
    if b["type"] == "tornado":
        b["title"] = _ytit.get(b.get("value_col"), "EPS 相對基準的變動") + ("（情境 A；變動未達兩位數百分比的案例）" if b.get("collapse") else "（情境 A；變動達兩位數百分比的案例；一次動一個輸入）")
for t_ in cfg["hero"]["tiles"]:
    if isinstance(t_[1], str) and t_[1].endswith("|n2"):
        t_[1] = t_[1][:-3] + "|n0"

# ---- 後處理三：命題章補情境定義；附錄列出全部版本說明列 ----
old_kv = next((b for b in sec("thesis", SRC)["blocks"] if b["type"] == "kv"), None)
if old_kv:
    sec("thesis")["blocks"] += [{"type": "h3", "text": "情境定義"}, copy.deepcopy(old_kv)]
ws = wb["README"]
r_end = find("README", "展示層")
ver_rows = [[f"@README!A{r}"] for r in range(1, ws.max_row + 1) if re.match(r"^(v\d|\d{4}-\d{2}-\d{2}：|版本：)", str(ws.cell(r, 1).value or ""))]
st = sec("status")["blocks"]
st[:] = [b for b in st if not (b["type"] == "h3" and b.get("text") == "附錄：版本要點")]
st[:] = [b for b in st if not (b["type"] == "table" and str(b.get("summary", "")).startswith("展開：本版與前版"))]
ins = next(i for i, b in enumerate(st) if b["type"] == "h3" and "未找到" in str(b.get("text")))
st[ins:ins] = [{"type": "h3", "text": "附錄：版本要點"}, {"type": "table", "collapse": True, "summary": "展開：各版的修改要點", "head": ["要點"], "rows": ver_rows}]

json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", OUT)
