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
            # v0.3.3（第三輪修正第 15、31、39 項）：三情境欄依單位格式化（% 與 pt 列顯示百分比、年度列不加千分位）
            _UF = {"%": "p1", "pt": "p1", "x": "n2", "年": "yr", "EB": "n2", "US$B": "n1", "千片": "n0"}
            for row_ in blocks[i_]["rows"]:
                r_ = int(re.search(r"In!B(\d+)", row_[0]).group(1)); f_ = _UF.get(str(wb["In"].cell(r_, 3).value or ""))
                if f_:
                    row_[1:4] = [x.split("|")[0] + "|" + f_ for x in row_[1:4]]
_shift_scn(sec("drivers")["blocks"])
for b in sec("drivers")["blocks"]:
    if str(b.get("caption", "")).startswith("單一參數"):
        b["caption"] = "驅動 → 推導對照表（In 頁基準值與 PL 頁推導量）"

# ---- 前處理一：估值章（重點觀察已在第一章；kv 改為有標題的表；DCF 前加折現率表）----
def _dfmt(t):
    """DCF 頁明細列的顯示格式（v0.3.3）：只有稅率、折舊攤銷 ÷、資本支出 ÷、終值現值占 EV 用百分比；NOPAT、終值、每股價值不顯示成 %。"""
    if t.startswith(("稅率", "折舊攤銷 ÷", "資本支出 ÷", "終值現值占", "評價日後占")): return "p1"
    if t.startswith("折現因子"): return "n4"
    if t.startswith("隱含終值 EV／EBITDA"): return "x1"
    if t.startswith("每股價值"): return "n0"
    return "n1"


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
        for _lab in ("明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "＋Stub 期 FCFF", "淨現金滾到評價日"):   # v0.3.3：金額（¥B）列取一位小數（stub 期 FCFF 約 ¥0.6B，整數顯示為 1）
            pct[str(find("Val", _lab))] = {c: "n1" for c in "BCD"}
        nb.append({"type": "table", "caption": f"@Val!A{h - 1}", "head": [f"@Val!{c}{h}" for c in "ABCDE"],
                   "rowgen": {"sheet": "Val", "rows": f"{pvx}-{dcf}", "cols": list("ABCDE"), "fmts": ["", "n0", "n0", "n0", ""], "row_fmt": pct}})
        b0 = find("DCF", "情境 B（")
        b1 = find("DCF", "檢查：終值占 EV", start=b0)
        hd = find("DCF", "項目")
        lab = lambda r: str(wb["DCF"].cell(r, 1).value or "")
        nb.append({"type": "table", "collapse": True, "summary": "展開：情境 B 的無槓桿自由現金流明細（DCF 頁；情境 A、C 同式）", "caption": f"@DCF!A{b0}",
                   "head": [f"@DCF!{c}{hd}" for c in "ACDEFGHI"],
                   "rowgen": {"sheet": "DCF", "rows": f"{b0 + 1}-{b1}", "cols": list("ACDEFGHI"), "fmts": ["", "n1", "n1", "n1", "n1", "n1", "n1", "n1"],
                              "row_fmt": {str(r): {c: _dfmt(lab(r)) for c in "CDEFGHI"} for r in range(b0 + 1, b1 + 1) if _dfmt(lab(r)) != "n1"}}})   # v0.3.3：只有比例列用百分比（第三輪修正第 38 項）
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
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "rf_us", "beta", "beta_u", "beta_u_med", "beta_du", "beta_de", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "bs_date", "stub_frac", "stub_div", "sh_v"]],
    ["自由現金流（折舊、資本支出、營運資金）", ["fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]],
    ["目標價規則、情境權重與收斂期", ["tp_w_dcf", "tp_wA", "tp_wB", "tp_wC", "tp_g", "nim10"]]]]],
 "blocks_to": {},
 "main": {
  "0.1 DCF": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "＋Stub 期 FCFF", "淨現金滾到評價日", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
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

# ================= v0.3.5（2026-11-15）：個股 HTML 版面規格 v1.3（樣板 NVIDIA v0.3.10、泰瑞達 v0.4.5、Advantest v0.7.5；docs/plan/個股HTML版面規格.md「v1.3 規格」） =================
# 章節：命題 → 估值 → 敏感度 → 反向 DCF → 損益（標準損益表）→ 驅動因子（分部營收推導）→ 對帳（含 C0 實績）→ 資料狀態。
# 數字一律取自 xlsx（README 頁末 v1.3 白話列、PL 頁末標準損益表、RevDCF 頁、Val 0.7、Sens 頁末相對基準表、Coord 頁 C0；towa/v035_post.py 產生）。列號依標籤找；設定文字不含數字（年份與日期除外）。
def _v13(cfg):
    S = {x["id"]: x for x in cfg["sections"]}
    RD, PLW, RV, SW, CW = wb["README"], wb["PL"], wb["RevDCF"], wb["Sens"], wb["Coord"]
    fnd = lambda ws, p, start=1: next(r for r in range(start, ws.max_row + 1) if str(ws.cell(r, 1).value or "").strip().startswith(p))
    r13 = fnd(RD, "v1.3 展示層")
    RM = dict(zip(["hero", "rec", "o1", "o2", "o3", "o4", "o5", "drv", "ver"], range(r13 + 1, r13 + 10)))
    assert str(RD[f"A{RM['ver']}"].value).startswith("v0.3.5"), RD[f"A{RM['ver']}"].value
    cfg["layout"] = "v1.3"
    fmt_u = lambda u: {"%": "p1", "US$B": "n2", "¥B": "n1", "百萬股": "n1", "¥": "n1", "x": "n2"}.get(str(u).strip(), "n1")
    # ---- 橫幅：白話說明＋對帳中揭露句；評等規則與歸因移到估值章 ----
    old_note = list(cfg["hero"]["note"])
    i_cd = len(old_note) - 2   # 末兩項＝空格、Coord 橫幅揭露句
    assert old_note[i_cd] == " " and str(old_note[i_cd + 1]).startswith("@Coord!B"), old_note
    cfg["hero"]["note"] = [f"@README!A{RM['hero']}"] + old_note[i_cd:]
    # ---- 1. 命題：主圖＝情境 B 營收組成（堆疊）與三情境營收 ----
    hdr = 3   # PL 頁最上方：情境／項目 FY2025A … FY2029
    assert PLW.cell(hdr, 1).value == "情境／項目" and PLW.cell(hdr, 3).value == "FY2025A", PLW.cell(hdr, 3).value
    pa, pbb, pc = fnd(PLW, "情境 A"), fnd(PLW, "情境 B"), fnd(PLW, "情境 C")
    revA, revB, revC = pa + 1, pbb + 1, pc + 1
    assert all(PLW.cell(x, 1).value == "營收" for x in (revA, revB, revC)) and PLW.cell(revA, 2).value == "¥B"
    segs = [revB + i for i in (2, 3, 4, 5, 6, 7)]
    assert [PLW.cell(x, 1).value.strip() for x in segs] == ["AI 封膠（HBM＋2.5D）", "非 AI 封膠設備與模具", "其他半導體（切割等）", "新事業（TSS 等）", "醫療器材", "雷射"], [PLW.cell(x, 1).value for x in segs]
    rb = fnd(PLW, "情境 B（基準）：")
    prow = lambda lab_, start=rb: next(r for r in range(start, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip().startswith(lab_))
    th = S["thesis"]["blocks"]
    chart1 = {"type": "theme", "title": "營收組成（情境 B）與三情境營收（十億日圓）", "x": f"@PL!C{hdr}:G{hdr}", "fmt": "n0", "stack": True, "pad": 1.3,
              "total": f"@PL!C{revB}:G{revB}",
              "bars": [{"name": f"@PL!A{r}", "y": f"@PL!C{r}:G{r}", "color": c} for r, c in zip(segs, ("#0b7a80", "#6fc7cb", "#c29a2b", "#8a8c8c", "#b9b9b9", "#6f6bae"))],
              "lines": [{"name": [f"@PL!A{pa}", "營收"], "y": f"@PL!C{revA}:G{revA}", "color": "#0b7a80", "dash": True},
                        {"name": [f"@PL!A{pc}", "營收"], "y": f"@PL!C{revC}:G{revC}", "color": "#b4472f", "dash": True}],
              "caption": "堆疊＝情境 B（基準）的 AI 封膠、非 AI 封膠、其他半導體、新事業、醫療器材與雷射，頂端數字＝情境 B 營收；虛線＝情境 A、情境 C 的營收。FY2026 為公司全年指引錨點；三情境的差異來自 HBM 位元、CoWoS 片數（加速器顆數）、設備市場與 AI 封膠市占。"}
    i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
    th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RM[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
    th.insert(0, chart1)
    th.insert(1, {"type": "note", "parts": [f"@README!A{RM['rec']}"]})
    # ---- 2. 估值：第一個區塊＝四個目標價與現價（G 位置）----
    va = S["valuation"]["blocks"]
    g0 = fnd(wb["Val"], "0.7 估值並列"); gr = list(range(g0 + 2, g0 + 7)); r_pxg = g0 + 7
    four = {"type": "bars", "title": "四個目標價與現價（日圓；豎線＝現價）", "fmt": "n0",
            "items": [{"label": lb, "value": f"@Val!B{r}", "series": sr} for r, sr, lb in zip(gr, (1, 2, 2, 0, 1),
                      ("DCF 目標價", "Comparable（調整後，進加權）", "Comparable（未調整，不進加權）", "加權目標價（橫幅主值）", "市場共識目標價"))],
            "ref": {"label": f"@Val!A{r_pxg}", "value": f"@Val!B{r_pxg}"}}
    tbl4 = {"type": "table", "caption": f"@Val!A{g0}", "head": [f"@Val!{c}{g0 + 1}" for c in "ABCD"],
            "rows": [[f"@Val!A{r}", f"@Val!B{r}|n2", f"@Val!C{r}|sp1", f"@Val!D{r}"] for r in gr] + [[f"@Val!A{r_pxg}", f"@Val!B{r_pxg}|n2", "", f"@Val!D{r_pxg}"]]}
    keep = [b for b in va if not (b.get("type") == "bars" and b.get("title") == "目標價與現價")]
    head_note = {"type": "note", "parts": old_note[:i_cd]}
    S["valuation"]["blocks"] = [four, head_note, tbl4] + keep
    v_intro = S["valuation"].pop("intro", None)   # 章說明放在圖之後（G：章標題後第一個區塊是圖）
    if v_intro:
        S["valuation"]["blocks"].insert(2, {"type": "note", "parts": [v_intro]})
    # ---- 3. 敏感度：主圖＝加權目標價相對基準（H 位置）----
    se = S["sens"]["blocks"]
    t0 = fnd(SW, "加權目標價相對基準（HTML"); r1 = t0 + 2; r2 = r1
    while SW.cell(r2 + 1, 1).value and not str(SW.cell(r2 + 1, 1).value).startswith("快照狀態"):
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
            b["summary"] = ("展開：對照：FY2027 EPS 相對基準的變動未達兩位數百分比的案例（情境 A）" if b.get("collapse") else "展開：對照：FY2027 EPS 相對基準的變動達兩位數百分比的案例（情境 A，一次動一個輸入）")
            b["collapse"] = True
        if b.get("type") == "table" and not b.get("collapse"):
            b["collapse"] = True
            b["summary"] = "展開：目標價參數表（WACC、β、ERP、永續成長率等）"
    intro_s = {"type": "note", "parts": ["主圖為加權目標價相對基準的變動（低案、高案分開，一次動一個輸入；變動達兩位數百分比者直接顯示，其餘摺疊）。業務主因為整本重算快照，快照狀態：",
                                          f"@Sens!B{snap}", "；目標價參數取 Sens 頁「目標價參數敏感度」表。下方摺疊為全部案例與 EPS 敏感度對照。"]}
    sec_intro = S["sens"].pop("intro", None)
    S["sens"]["blocks"] = tor[:1] + tor[2:3] + [intro_s, tor[1], tor[3], full] + ([{"type": "note", "parts": [sec_intro], "collapse": True, "summary": "展開：EPS 敏感度（情境 A）的說明"}] if sec_intro else []) + se
    # ---- 4. 反向 DCF（新）----
    rv = lambda p, start=1: next(r for r in range(start, RV.max_row + 1) if str(RV.cell(r, 1).value or "").strip().startswith(p))
    ra, rcmp, rsum, yh = rv("A. DCF"), rv("B. Comparable"), rv("C. 摘要"), rv("項目（情境 B）")
    rb_, rbb, rcc = rv("(a)"), rv("(b)"), rv("(c)")
    m1, m0 = rv("現價隱含的營業利益率", rb_), rv("模型：基準的等效營業利益率", rb_)
    s1, rv1, rv0, g1, g0b = rv("現價隱含的放大倍數", rbb), rv("現價隱含的 FY2028 營收", rbb), rv("模型：FY2028 營收", rbb), rv("現價隱含的 FY2028 營收年增率", rbb), rv("模型：FY2028 營收年增率", rbb)
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
                          ["現金流放大倍數（FY2028 起各年同乘；模型基準為一倍）", f"@RevDCF!B{s1}|n2", "—"],
                          ["FY2028 營收（十億日圓）", f"@RevDCF!B{rv1}|n0", f"@RevDCF!B{rv0}|n0"],
                          ["FY2028 營收年增率", f"@RevDCF!B{g1}|p1", f"@RevDCF!B{g0b}|p1"],
                          ["折現率（WACC）", f"@RevDCF!B{w1}|p2", f"@RevDCF!B{w0}|p2"]]},
                {"type": "table", "caption": f"@RevDCF!A{ra}", "head": ["項目", "值"], "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|{f}"] for r, f in zip(a_rows, afmt)]},
                {"type": "table", "caption": f"@RevDCF!A{rcmp}", "head": [f"@RevDCF!{c}{rcmp + 1}" for c in "ABCD"],
                 "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|{f}", f"@RevDCF!C{r}|{f}", f"@RevDCF!D{r}|sp1"] for r, f in ((pe1, "n1"), (rc1, "n2"), (pe2, "n1"))]},
                {"type": "table", "caption": f"@RevDCF!A{yh}", "head": [f"@RevDCF!{c}{yh}" for c in "ABCDEFG"], "collapse": True,
                 "summary": "展開：情境 B 各年數字（反解的計算基礎；演算過程見 RevDCF 頁）",
                 "rows": [[f"@RevDCF!A{r}"] + [f"@RevDCF!{c}{r}|{f}" for c in "BCDEFG"] for r, f in zip(yrs_rows, yfmt)]}]}
    # ---- 5. 損益：標準損益表 ----
    p0 = fnd(PLW, "標準損益表（v0.3.5")
    ph = next(r for r in range(p0, PLW.max_row + 1) if PLW.cell(r, 1).value == "項目")
    plr = lambda lab_, start=ph: next(r for r in range(start, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip() == lab_)
    keys = ["營收", "營收年增率（YoY）", "營業成本（COGS）", "毛利", "毛利率", "營業費用（銷管費，含研發）", "營業利益", "營業利益率", "其他收入（淨額）", "稅前利益", "所得稅",
            "稅後淨利（歸屬母公司）", "淨利率", "基本股數（百萬股）", "EPS（J-GAAP 基本）"]
    PC = "CDEFGHIJKL"
    rowx = lambda r: [f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmt_u(PLW.cell(r, 2).value)}" for c in PC]
    r_ratio = plr("比率"); r_chk = next(r for r in range(r_ratio, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("核對（應為 0"))
    r_ycav = next(r for r in range(r_chk, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("年度與對帳中"))
    is_tbl = {"type": "table", "caption": "標準損益表（十億日圓；EPS 為日圓）：FY2023A–FY2025A 實績、FY2026E–FY2032E 情境 B（基準；GAAP 口徑＝日本會計準則，無 Non-GAAP；模型以增量營益率推導，預測期不推毛利，空格為未推或未取得）",
              "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in PC],
              "rows": [rowx(plr(k)) for k in keys] + [{"sub": f"@PL!A{r_ratio}"}] + [rowx(r) for r in range(r_ratio + 1, r_chk)]}
    chk_tbl = {"type": "table", "caption": f"@PL!A{r_chk}", "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in PC], "collapse": True,
               "summary": "展開：核對列（歷史捨入差、預測期對 PL 情境 B 與延伸表）",
               "rows": [[f"@PL!A{r}"] + [f"@PL!{c}{r}|n3" for c in PC] for r in range(r_chk + 1, r_ycav) if PLW.cell(r, 1).value]}
    rv_, gm_, opm_, nm_ = plr("營收"), plr("毛利率"), plr("營業利益率"), plr("淨利率")
    is_chart = {"type": "theme", "title": "營收（十億日圓）與毛利率、營業利益率、淨利率：FY2023A–FY2032E（情境 B）", "x": f"@PL!C{ph}:L{ph}",
                "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
                "bars": [{"name": f"@PL!A{rv_}", "y": f"@PL!C{rv_}:L{rv_}", "color": "#6fc7cb"}],
                "lines": [{"name": f"@PL!A{gm_}", "y": f"@PL!C{gm_}:L{gm_}", "axis": "r", "color": "#0b7a80", "labels": True},
                          {"name": f"@PL!A{opm_}", "y": f"@PL!C{opm_}:L{opm_}", "axis": "r", "color": "#c29a2b"},
                          {"name": f"@PL!A{nm_}", "y": f"@PL!C{nm_}:L{nm_}", "axis": "r", "color": "#6f6bae", "dash": True}],
                "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸，日本會計準則）。毛利率只有 FY2025A（預測期不推）；FY2026E 為公司指引錨點、FY2027E 起為情境推導；FY2030E–FY2032E 為收斂法延伸期（只推營收與淨利，營業利益由淨利還原）。"}
    old_pl = S["pl"]["blocks"]
    pls = [b for b in old_pl if b["type"] == "pl"]
    assert len(pls) == 2, [b["type"] for b in old_pl]
    ext_notes = [b for b in old_pl if b["type"] == "note"]
    old_blocks = [dict(pls[0], collapse=True, summary="展開：三情境損益摘要（FY2025–FY2029；營收、事業別、AI 封膠占營收、營業利益率、EPS）")]
    old_blocks += [dict(b, collapse=True, summary="展開：延伸期（FY2030 起）的收斂法說明") for b in ext_notes]
    old_blocks += [dict(pls[1], collapse=True, summary="展開：三情境延伸到 FY2032（收斂法）")]
    pl_intro = S["pl"].pop("intro", None)
    S["pl"].update(title="損益（標準損益表）",
                   blocks=[{"type": "note", "parts": [f"@PL!A{p0 + 1}"]}, is_chart, is_tbl, {"type": "note", "parts": [f"@PL!A{p0 + 2}"]}, {"type": "note", "parts": [f"@PL!A{r_ycav}"]}, chk_tbl]
                          + ([{"type": "note", "parts": [pl_intro], "collapse": True, "summary": "展開：三情境損益摘要的期間說明"}] if pl_intro else []) + old_blocks)
    # ---- 6. 驅動因子：營收推導（情境 B）----
    bh = rb + 1
    assert str(PLW.cell(bh, 3).value).startswith("CY2025") and str(PLW.cell(bh, 7).value).startswith("CY2029"), (PLW.cell(bh, 3).value, PLW.cell(bh, 7).value)
    seg = lambda rs: [[f"@PL!A{r}", f"@PL!B{r}"] + [f"@PL!{c}{r}|{fmt_u(PLW.cell(r, 2).value)}" for c in "CDEFG"] for r in rs]
    cy = [prow(x) for x in ("HBM 出貨位元", "HBM 堆疊數", "HBM 封膠新機需求", "CoWoS 片數需求", "2.5D 封膠新機需求", "非 AI 封裝設備指數", "設備總額")]
    e0 = prow("E. 損益")
    fy = [prow(x, e0) for x in ("HBM 封膠", "CoWoS 等 2.5D 封膠", "非 AI 封膠設備與模具", "其他半導體（切割等）", "新事業（TSS 等）", "醫療器材", "雷射", "AI 封膠合計", "半導體製造裝置事業營收")]
    r_rev = next(r for r in range(fy[-1], PLW.max_row + 1) if PLW.cell(r, 1).value == "營收")
    r_aish = prow("AI 封膠占營收", r_rev)
    pr = [next(r for r in range(r_rev, PLW.max_row + 1) if PLW.cell(r, 1).value == "營業利益")] + [prow(x, r_rev) for x in ("營業利益率", "淨利", "EPS")]
    seg_tbl = {"type": "table", "caption": "營收推導（情境 B）：上半段為日曆年堆疊數、片數與設備市場，下半段為會計年度營收（FY＝四分之三當年＋四分之一次年）", "head": [f"@PL!A{bh}", f"@PL!B{bh}"] + [f"@PL!{c}{bh}" for c in "CDEFG"],
               "rows": [{"sub": "一、日曆年：HBM 位元 → 堆疊數 → 封膠新機需求；CoWoS 片數 → 新機需求；非 AI 封裝設備與設備總額"}] + seg(cy)
                       + [{"sub": "二、會計年度營收：AI 封膠（校準後）＋非 AI 封膠＋其他半導體＋其他事業"}] + seg(fy + [r_rev, r_aish])
                       + [{"sub": "三、營業利益與每股盈餘"}] + seg(pr)}
    seg_chart = {"type": "theme", "title": "營收組成（十億日圓；情境 B）與 AI 封膠占營收（右軸）", "x": f"@PL!C{bh}:G{bh}", "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.25,
                 "bars": [{"name": f"@PL!A{r}", "y": f"@PL!C{r}:G{r}", "color": c} for r, c in zip(fy[:7], ("#0b7a80", "#6fc7cb", "#c29a2b", "#8a8c8c", "#b9b9b9", "#6f6bae", "#3b3838"))],
                 "lines": [{"name": f"@PL!A{r_aish}", "y": f"@PL!C{r_aish}:G{r_aish}", "axis": "r", "color": "#b4472f", "labels": True}],
                 "caption": "長條＝HBM 封膠、CoWoS 等先進封裝封膠、非 AI 封膠、其他半導體、新事業、醫療器材與雷射（會計年度）；折線＝AI 封膠占營收。FY2025 為實績校準、FY2026 為錨點、FY2027 起為情境 B 推導。"}
    dr = S["drivers"]["blocks"]
    d_intro = S["drivers"].pop("intro", None)
    S["drivers"].update(title="驅動因子：營收推導")
    S["drivers"]["blocks"] = ([{"type": "note", "parts": [f"@README!A{RM['drv']}"]}, seg_chart, seg_tbl, {"type": "h3", "text": "情境驅動與單一參數"}]
                              + ([{"type": "note", "parts": [d_intro]}] if d_intro else []) + dr)
    # ---- 7. 對帳：C0（讀取層級與實績）----
    c1 = fnd(CW, "1. 讀哪一層"); c2 = fnd(CW, "2. C0"); c3 = fnd(CW, "3. 對帳中年份")
    h1, h2 = c1 + 1, next(r for r in range(c2, CW.max_row + 1) if CW.cell(r, 1).value == "項目")
    cnt = fnd(CW, "對帳中項目數")
    cunit = lambda r: str(CW.cell(r, 2).value or "")
    cf = lambda r: "p1" if cunit(r) in ("%",) else "n3" if cunit(r) in ("x", "US$B") else "n1"
    c0_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|{cf(r)}", f"@Coord!E{r}|{cf(r)}", f"@Coord!F{r}|sp1", f"@Coord!G{r}"]
               for r in range(h2 + 1, cnt) if CW.cell(r, 4).value is not None]
    c0_tbl = {"type": "table", "caption": "讀取層級核對：個股讀入的層級值 vs 層級值（同源，不判定、不計入對帳中；有檢驗力的對帳見上表「公司實績 vs 層級隱含值」）", "collapse": True, "summary": "展開：讀取層級核對（同源，不判定）", "head": [f"@Coord!{c}{h2}" for c in "ABCDEFG"], "rows": c0_rows}
    a4 = fnd(CW, "4. C0（實績）"); ih = next(r for r in range(a4, CW.max_row + 1) if CW.cell(r, 1).value == "項目"); mh = next(r for r in range(ih, CW.max_row + 1) if CW.cell(r, 1).value == "讀數")
    a4e = next(r for r in range(mh, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith("對帳中項目數"))
    fm = lambda r: "p1" if str(CW.cell(r, 2).value) == "%" else "n3" if str(CW.cell(r, 2).value) == "US$B" else "n1"
    isprog = lambda r: str(CW.cell(r, 1).value or "").startswith("參考：已公布期間實績占")
    act_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|n1", f"@Coord!E{r}|n1", (f"@Coord!F{r}|p1" if isprog(r) else f"@Coord!F{r}|sp1"), f"@Coord!G{r}"]
                for r in range(mh + 1, a4e - 1)]
    in_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|" + ("n0" if CW.cell(r, 2).value == "月" else fm(r)), f"@Coord!E{r}", f"@Coord!H{r}"] for r in range(ih + 1, mh - 1) if CW.cell(r, 1).value]
    ot = fnd(CW, "其他公司實績"); oh = next(r for r in range(ot, CW.max_row + 1) if CW.cell(r, 1).value == "項目"); oe = next(r for r in range(oh, CW.max_row + 2) if not CW.cell(r + 1, 1).value)
    oth_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|n1", f"@Coord!E{r}|n1", f"@Coord!G{r}"] for r in range(oh + 1, oe + 1)]
    act_blocks = [{"type": "h3", "text": "公司實績對層級隱含值（實績對帳）"}, {"type": "note", "parts": [f"@Coord!A{a4 + 1}"]}, {"type": "note", "parts": [f"@Coord!A{a4 + 2}"]}, {"type": "note", "parts": [f"@Coord!A{a4 + 3}"]},
                  {"type": "table", "caption": "營收：公司決算短信實績 vs 層級隱含值（層級沒有 TOWA 份額；份額沿用 FY2025 實績的兩列為假設，只列參考、不判定；已公布期間只有第一季）", "head": [f"@Coord!{c}{mh}" for c in "ABCDEFG"], "rows": act_rows},
                  {"type": "note", "parts": [f"@Coord!A{a4e - 1}"]},
                  {"type": "note", "parts": ["對帳中項目數：", f"@Coord!G{a4e}|n0", "。"]},
                  {"type": "table", "caption": "實績與層級隱含值的輸入、等級與出處（展開看）", "collapse": True, "summary": "展開：實績、層級值、換算參數與出處", "head": [f"@Coord!{c}{ih}" for c in "ABCDE"] + [f"@Coord!H{ih}"], "rows": in_rows},
                  {"type": "table", "caption": f"@Coord!A{ot}", "collapse": True, "summary": "展開：營業利益、淨利、半導體事業營收與受注的實績（層級無可比量，列示供對照）", "head": [f"@Coord!{c}{oh}" for c in "ABCDE"] + [f"@Coord!G{oh}"], "rows": oth_rows}]
    c0_note = [{"type": "h3", "text": "讀取層級核對（同源，不判定）"}, {"type": "note", "parts": [f"@Coord!A{c2}"]}, {"type": "note", "parts": [f"@Coord!A{c2 + 1}"]}, c0_tbl, {"type": "note", "parts": [f"@Coord!A{cnt + 1}"]}]
    c0_src = {"type": "table", "caption": "層級名稱、版本、口徑與證據等級（每列的來源；展開看）", "collapse": True, "summary": "展開：每列的層級名稱、版本、口徑與證據等級",
              "head": ["項目", "年度", "層級名稱與版本／口徑／證據等級／備註"], "rows": [[f"@Coord!A{r}", f"@Coord!C{r}", f"@Coord!H{r}"] for r in range(h2 + 1, cnt) if CW.cell(r, 8).value]}

    def _rd(r):
        v = CW.cell(r, 2).value
        return [f"@Coord!A{r}", f"@Coord!B{r}" + ("|n3" if isinstance(v, (int, float)) and not isinstance(v, bool) else ""), f"@Coord!C{r}", f"@Coord!D{r}", f"@Coord!E{r}"]
    rd_tbl = {"type": "table", "caption": f"@Coord!A{c1}", "head": [f"@Coord!{c}{h1}" for c in "ABCDE"], "collapse": True, "summary": "展開：本模型讀哪一層、哪一版（下游只註明版本、不自行重估）",
              "rows": [_rd(r) for r in range(h1 + 1, c2 - 1) if CW.cell(r, 1).value]}
    rf0 = fnd(CW, "層級之間"); rfh = next(r for r in range(rf0, CW.max_row + 1) if CW.cell(r, 1).value == "項目"); rfe = next(r for r in range(rfh, CW.max_row + 1) if not CW.cell(r + 1, 1).value)
    ref_blocks = [{"type": "h3", "text": f"@Coord!A{rf0}"}, {"type": "note", "parts": [f"@Coord!A{rf0 + 1}"]},
                  {"type": "table", "head": [f"@Coord!{c}{rfh}" for c in "ABCDEFG"],
                   "rows": [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|n3", f"@Coord!E{r}|n3", f"@Coord!F{r}|sp1", f"@Coord!G{r}"] for r in range(rfh + 1, rfe + 1)]}]
    c3_notes = [{"type": "note", "parts": [f"@Coord!A{c3 + 1}"]}, {"type": "note", "parts": [f"@Coord!A{c3 + 2}"]}]
    S["recon"]["blocks"] = [{"type": "h3", "text": "三層協調：由下往上對帳（公司實績對各層）"}] + act_blocks + c0_note + ref_blocks + c3_notes + [c0_src, rd_tbl] + [{"type": "h3", "text": "公司目標、先驗比對與校準"}] + S["recon"]["blocks"]
    S["recon"]["title"] = "對帳：三層協調、公司實績與先驗比對"
    # ---- 8. 資料狀態 ----
    st = S["status"]["blocks"]
    st[0:0] = [{"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RM['ver']}"]}]
    order = ["thesis", "valuation", "sens", "revdcf", "pl", "drivers", "recon", "status"]
    S["revdcf"] = rev_
    cfg["sections"] = [S[k] for k in order]





# 三層協調 C4（規則 v1.4；tools/coord/coord.py）：橫幅另一側讀數一格與揭露句、估值章開頭兩讀數表與長條圖；之後套用版面 v1.3
sys.path.insert(0, os.path.join(HERE, "..", "tools", "coord"))
import coord as COORD  # noqa: E402
COORD.inject_html(cfg, X)
_v13(cfg)
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", OUT)
