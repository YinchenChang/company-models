# v0.3（2026-11-05）：味之素首次套用「個股 HTML 版面規格 v1.0」（docs/plan/個股HTML版面規格.md），同時 DCF 改標準版（FCFF＋WACC）。寫法比照 sndk/、kioxia/ 的 gen_html_json.py。
# 由 html_src.json（v0.2.1 的 html.json，內容不動）產生 ajinomoto/html.json；估值章 DCF 區塊、Sens 目標價參數表與情境驅動表的列號由本檔依標籤重產。
# 用法：python3 ajinomoto/gen_html_json.py [成品 xlsx]（省略時取 ajinomoto/dist/ 內唯一的 xlsx）；勿手改 html.json。
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
            # v0.3.3（第三輪修正第 15、39 項）：三情境欄依單位格式化（% 列顯示百分比）
            for row_ in blocks[i_]["rows"]:
                r_ = int(re.search(r"In!B(\d+)", row_[0]).group(1)); u_ = str(wb["In"].cell(r_, 3).value or "")
                f_ = "p1" if u_ == "%" else ("n2" if u_ == "x" else "")
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
        cap = ("Comparable 目標價的計算（市場主流同業法：食品同業前瞻 P/E 中位數 × 前瞻 EPS）" if lab0.startswith("同業前瞻 P/E 中位數")
               else "分部加總與現價隱含的 ABF 價值（參考，不進目標價）" if lab0.startswith("電子材料同業") else "現價與市場隱含倍數")
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
                              "row_fmt": {str(r): {c: _dfmt(lab(r)) for c in "CDEFGHI"} for r in range(b0 + 1, b1 + 1) if _dfmt(lab(r)) != "n1"}}})   # v0.3.3：只有比例列用百分比（第三輪修正第 38 項）
        i += 1
        continue
    if b["type"] == "table" and txt(b.get("caption", "")).startswith("0.2 Comparable"):
        b = dict(b, collapse=True, summary="展開：食品同業的前瞻 P/E")
    nb.append(b)
    i += 1
va["blocks"] = nb

# ---- 前處理二：對帳表加標題；敏感度「目標價參數」表的列號改為 Sens 頁實際位置（標準版 DCF 案例）----
for b in sec("recon")["blocks"]:
    if b["type"] == "table" and not b.get("caption"):
        b["caption"] = "與公司預測、共識與上游載板模型的對照（只作事後對照）"
for b in sec("sens")["blocks"]:
    if b["type"] == "table" and str(b.get("caption", "")).startswith("@Sens!A") and "目標價參數" in txt(b["caption"]):
        t0 = find("Sens", "目標價參數")
        t1 = find("Sens", "檢查：閉式公式的基準", start=t0) - 1
        b["caption"] = f"@Sens!A{t0}"
        b["head"] = [f"@Sens!{c}{t0 + 1}" for c in "ABCDEFG"]
        b["rowgen"] = {"sheet": "Sens", "rows": f"{t0 + 2}-{t1}", "cols": list("ABCDEFG"), "fmts": ["", "", "n0", "sp1", "n0", "sp1", "n0"]}

# ---- 轉換：版面規格 v1.0 ----
scn_tok = re.split(r"[.\s]", str(wb["In"].cell(find("In", "情境驅動對照"), 1).value))[0]
B_FM, B_CORE, B_CO = "ファンクショナルマテリアルズ（FM）：ABF 增層膜", "其他事業：調味料・食品、冷凍食品、生技醫藥等", "全公司：錨點、股價股數、折現率、目標價規則與現金流"
SPEC = {
 "in_skip": ["G", scn_tok],
 "in_groups": [
  [B_FM, [
    ["AI 加速器顆數（讀 tsmc 成品）與面積係數", ["unv_26", "unv_27", "unv_28", "unv_29", "uot_26", "uot_27", "uot_28", "uot_29", "v_ot", "g_asp"]],
    ["情境顆數係數與 ABF 價格乘數", ["um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C", "sp_A", "sp_B", "sp_C"]],
    ["ABF 需求、份額、產能、價格與成本", ["ai_sh", "g_nai", "rel_27", "rel_28", "rel_29", "cap_27", "cap_28", "cap_29", "pp_27", "pp_28", "pp_29", "c_g"]],
    ["FM 錨點與實績", ["f_fm", "f_fmbp", "q1_fm", "q1_fmbp", "h1_fm", "h1_fmbp", "a_fm", "a_fmbp"]],
    ["Wattonomics 錨定係數（只用於敏感度）", ["wta_27", "wta_28", "wta_29"]]]],
  [B_CORE, [
    ["成長率與利益率", ["g_sf", "m_sf", "g_ff", "m_ff", "g_ho", "m_ho", "g_cc"]],
    ["各事業錨點與實績", ["f_sf", "f_ff", "f_ho", "f_oth", "f_sfbp", "f_ffbp", "f_hobp", "f_othbp", "f_cc", "a_sf", "a_ff", "a_ho", "a_oth", "a_sfbp", "a_ffbp", "a_hobp", "a_othbp", "a_cc"]]]],
  [B_CO, [
    ["全公司錨點與實績", ["f_rev", "f_bp", "f_oi", "f_ni", "f_eps", "q1_rev", "q1_bp", "q1_ni", "q1_eps", "h1_rev", "h1_bp", "h1_ni", "h1_eps", "a_rev", "a_bp", "a_oi", "a_ni", "a_eps", "a_da", "a_cx"]],
    ["淨利與股數", ["ni_bp", "sh_g", "so"]],
    ["股價與市場倍數", ["px", "pe_cf", "fpe", "cons_feps", "cons_tp"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "rf_us", "beta", "beta_u", "beta_u_food", "beta_u_chem", "beta_w_fm", "beta_du", "beta_de", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "bs_date", "stub_frac", "stub_div", "sh_v"]],
    ["自由現金流（折舊、資本支出、營運資金；da_rev、cx_rev 另供 PL 頁自由現金流列）", ["fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc", "da_rev", "cx_rev"]],
    ["目標價規則與情境權重", ["tp_w_dcf", "tp_wA", "tp_wB", "tp_wC"]],
    ["收斂期（FY2030 起）", ["tp_g", "nm10_k", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]],
    ["標準損益表推算（只用於推算毛利，不進 EPS 與目標價）", ["opx_el"]]]]],
 "blocks_to": {"加速器顆數": "ファンクショナルマテリアルズ（FM）：ABF 增層膜", "@Units!A": "ファンクショナルマテリアルズ（FM）：ABF 增層膜"},
 "main": {
  "0.1 DCF": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "＋Stub 期 FCFF", "淨現金滾到評價日", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
  "DCF 折現率": ["股權成本", "稅後債務成本", "權益占資本", "WACC", "永續成長率", "永續期 RONIC"],
 },
}
if sec("drivers").get("intro"):
    SPEC["blocks_to"][sec("drivers")["intro"]] = B_FM
cfg = relayout(cfg, X, SPEC)

# ---- 後處理一：橫幅中間格＝只看 Comparable（市場主流同業法、未做相對倍數調整；工作單補充規則 A）；進加權的調整後值寫在附註 ----
v_cmp0, v_up0 = find("Val", "Comparable 目標價（市場主流同業法）"), find("Val", "上檔空間：Comparable 目標價")
v_adj = find("Val", "Comparable 目標價（相對倍數調整後）")
rd_upa = find("README", "上檔空間：Comparable 目標價（相對倍數調整後）", find("README", "HTML 橫幅輔助值"))
cfg["hero"]["tiles"][1] = ["只看 Comparable 目標價（市場主流同業法）", V(v_cmp0, "n0"), ["上檔空間 ", V(v_up0, "sp1")]]
cfg["hero"]["note"] = list(cfg["hero"].get("note") or []) + ["；進加權的 Comparable 為相對倍數調整後 ", V(v_adj, "n0"), "（上檔空間 ", f"@README!B{rd_upa}|sp1", "）"]

# ---- 後處理二：龍捲風圖是 EPS 的變動（情境 A），不是加權目標價 ----
_ytit = {"J": "FY2028 EPS 相對基準的變動", "F": "FY2027 EPS 相對基準的變動"}
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


# ================= v0.3.5（2026-11-15）：三層協調 C4 兩讀數（tools/coord）＋個股 HTML 版面規格 v1.3（樣板 NVIDIA v0.3.10；同次產業 Ibiden v0.5.6 的 _v13，同一寫法） =================
# 章節：命題 → 估值 → 敏感度 → 反向 DCF → 損益（標準損益表）→ 驅動因子（分部營收推導）→ 對帳（含 C0）→ 資料狀態；無個別模組。
# 數字一律取自 xlsx（README 頁末 v1.3 白話列、PL 頁末標準損益表、RevDCF 頁、Val 0.7、Sens 頁末相對基準表、Coord 頁；ajinomoto/v035_post.py 產生）。列號依標籤找。
sys.path.insert(0, os.path.join(HERE, "..", "tools", "coord"))
import coord as COORD  # noqa: E402
COORD.inject_html(cfg, X)


def _v13():
    S = {x["id"]: x for x in cfg["sections"]}
    RD, PLW, RV, SW, CW = wb["README"], wb["PL"], wb["RevDCF"], wb["Sens"], wb["Coord"]
    fnd = lambda ws, p, start=1: next(r for r in range(start, ws.max_row + 1) if str(ws.cell(r, 1).value or "").strip().startswith(p))
    r13 = find("README", "v1.3 展示層")
    RM = dict(zip(["hero", "rec", "o1", "o2", "o3", "o4", "o5", "drv", "ver"], range(r13 + 1, r13 + 10)))
    assert str(RD[f"A{RM['ver']}"].value).startswith("v0.3.8"), RD[f"A{RM['ver']}"].value
    cfg["layout"] = "v1.3"
    fmt_u = lambda u: {"%": "p1", "¥": "n2", "十億股": "n3", "x": "n3", "百萬顆": "n2"}.get(str(u).strip(), "n1")
    # ---- 橫幅：≤200 字白話＋對帳中揭露句；評等規則與 Comparable 說明移到估值章 ----
    old_note = list(cfg["hero"]["note"])
    i_cd = len(old_note) - 2
    assert old_note[i_cd] == " " and str(old_note[i_cd + 1]).startswith("@Coord!B"), old_note
    cfg["hero"]["note"] = [f"@README!A{RM['hero']}"] + old_note[i_cd:]
    # ---- 1. 命題：主圖＝情境 B 分部營收（堆疊）與情境 A、C 合併營收 ----
    pa, pc = find("PL", "情境 A"), find("PL", "情境 C")
    fa = lambda base, lab_: next(r for r in range(base, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip() == lab_)
    rb = find("PL", "情境 B（基準）：")
    hdrB = rb + 1
    prow = lambda lab_, ex=True: next(r for r in range(rb, rb + 60) if (str(PLW.cell(r, 1).value or "").strip() == lab_ if ex else str(PLW.cell(r, 1).value or "").strip().startswith(lab_)))
    r_fmai, r_fmn = prow("FM 營收：AI 用"), prow("FM 營收：非 AI（PC、一般伺服器等）")
    r_sf, r_ff, r_hox, r_oth2 = prow("調味料・食品營收"), prow("冷凍食品營收"), prow("ヘルスケア等（扣除 FM）營收"), prow("その他營收（含公司合計與各事業加總的四捨五入差）")
    r_rev = prow("營收")
    revA, revC = fa(pa, "營收"), fa(pc, "營收")
    th = S["thesis"]["blocks"]
    chart1 = {"type": "theme", "title": "營收分部（情境 B）與三情境營收（十億日圓）", "x": f"@PL!D{hdrB}:G{hdrB}", "fmt": "n0", "stack": True, "pad": 1.3,
              "total": f"@PL!D{r_rev}:G{r_rev}",
              "bars": [{"name": f"@PL!A{r_fmai}", "y": f"@PL!D{r_fmai}:G{r_fmai}", "color": "#0b7a80"},
                       {"name": f"@PL!A{r_fmn}", "y": f"@PL!D{r_fmn}:G{r_fmn}", "color": "#6fc7cb"},
                       {"name": f"@PL!A{r_sf}", "y": f"@PL!D{r_sf}:G{r_sf}", "color": "#c29a2b"},
                       {"name": f"@PL!A{r_ff}", "y": f"@PL!D{r_ff}:G{r_ff}", "color": "#e2c98a"},
                       {"name": f"@PL!A{r_hox}", "y": f"@PL!D{r_hox}:G{r_hox}", "color": "#6f6bae"},
                       {"name": f"@PL!A{r_oth2}", "y": f"@PL!D{r_oth2}:G{r_oth2}", "color": "#8a8c8c"}],
              "lines": [{"name": [f"@PL!A{pa}", "營收"], "y": f"@PL!D{revA}:G{revA}", "color": "#0b7a80", "dash": True},
                        {"name": [f"@PL!A{pc}", "營收"], "y": f"@PL!D{revC}:G{revC}", "color": "#b4472f", "dash": True}],
              "caption": "堆疊＝情境 B（基準）的分部營收（ファンクショナルマテリアルズ分 AI 用與非 AI），頂端數字＝情境 B 合併營收；虛線＝情境 A、情境 C 的合併營收。FY2026 為公司中間決算時修正後預測的錨點；三情境差異只來自 FM（加速器顆數與 ABF 價格乘數）。"}
    i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
    th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RM[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
    th.insert(0, chart1)
    th.insert(1, {"type": "note", "parts": [f"@README!A{RM['rec']}"]})
    # ---- 2. 估值：第一個區塊＝四個目標價與現價（G 位置）；規則說明放圖下 ----
    va = S["valuation"]["blocks"]
    g0 = find("Val", "0.7 估值並列"); gr = list(range(g0 + 2, g0 + 7)); r_pxg = g0 + 7
    four = {"type": "bars", "title": "四個目標價與現價（日圓；豎線＝現價）", "fmt": "n0",
            "items": [{"label": lb, "value": f"@Val!B{r}", "series": sr} for r, sr, lb in zip(gr, (1, 2, 2, 0, 1),
                      ("DCF 目標價", "Comparable（調整後，進加權）", "Comparable（未調整，不進加權）", "加權目標價（橫幅主值）", "市場共識目標價"))],
            "ref": {"label": f"@Val!A{r_pxg}", "value": f"@Val!B{r_pxg}"}}
    tbl4 = {"type": "table", "caption": f"@Val!A{g0}", "head": [f"@Val!{c}{g0 + 1}" for c in "ABCD"],
            "rows": [[f"@Val!A{r}", f"@Val!B{r}|n0", f"@Val!C{r}|sp1", f"@Val!D{r}"] for r in gr] + [[f"@Val!A{r_pxg}", f"@Val!B{r_pxg}|n0", "", f"@Val!D{r_pxg}"]]}
    keep = [b for b in va if not (b.get("type") == "bars" and b.get("title") == "目標價與現價")]
    S["valuation"]["blocks"] = [four, {"type": "note", "parts": old_note[:i_cd]}, tbl4] + keep
    intro = S["valuation"].pop("intro", None)
    if intro:
        S["valuation"]["blocks"].insert(2, {"type": "note", "parts": [intro]})
    # ---- 3. 敏感度：主圖＝加權目標價相對基準（H 位置）----
    se = S["sens"]["blocks"]
    t0 = find("Sens", "加權目標價相對基準（HTML"); r1 = t0 + 2; r2 = r1
    while SW.cell(r2 + 1, 1).value:
        r2 += 1
    main, rest = [], []
    for col, side in (("B", "低"), ("C", "高")):
        num = lambda r: isinstance(SW[f"{col}{r}"].value, (int, float)) and not isinstance(SW[f"{col}{r}"].value, bool)
        allr = sorted([r for r in range(r1, r2 + 1) if num(r) and abs(SW[f"{col}{r}"].value) >= 0.0005], key=lambda r: -abs(SW[f"{col}{r}"].value))
        big = [r for r in allr if abs(SW[f"{col}{r}"].value) >= 0.1]
        if len(big) >= 2:
            ttl = f"加權目標價相對基準：輸入取{side}值，變動達兩位數百分比的案例"
        else:   # 沒有兩個以上達兩位數者：主圖放變動最大的五個案例
            big = allr[:5]
            ttl = f"加權目標價相對基準：輸入取{side}值，變動最大的案例"
        small = [r for r in allr if r not in big]
        base = {"type": "tornado", "sheet": "Sens", "label_col": "A", "value_col": col, "fmt": "sp1"}
        main.append(dict(base, title=ttl, rows=",".join(map(str, sorted(big)))))
        if small:
            rest.append(dict(base, title=f"加權目標價相對基準：輸入取{side}值", rows=",".join(map(str, sorted(small))), collapse=True,
                             summary=f"展開：加權目標價相對基準：輸入取{side}值，其餘案例"))
    full = {"type": "table", "caption": f"@Sens!A{t0}", "collapse": True, "summary": "展開：全部案例的加權目標價相對基準（低案、高案）",
            "head": [f"@Sens!{c}{t0 + 1}" for c in "ABC"], "rows": [[f"@Sens!A{r}", f"@Sens!B{r}|sp1", f"@Sens!C{r}|sp1"] for r in range(r1, r2 + 1)]}
    for b in se:
        if b.get("type") == "tornado":
            b["summary"] = "展開：對照：" + str(b.get("title", "EPS 敏感度"))
            b["collapse"] = True
        if b.get("type") == "table" and not b.get("collapse"):
            b["collapse"] = True
            b["summary"] = "展開：目標價參數表（WACC、β、ERP、永續成長率、DCF 權重等）"
    intro_s = {"type": "note", "parts": ["主圖為加權目標價相對基準的變動（低案、高案分開，一次動一個輸入），取 Sens 頁「目標價參數敏感度」表、同業比較法不調整與另一側讀數；"
                                          "業務驅動（AI 占比、產能上限、ABF 價格等）對 FY2027／FY2028 EPS 的影響見下方摺疊的 EPS 敏感度對照。"]}
    S["sens"].pop("intro", None)
    S["sens"]["blocks"] = main + [intro_s] + rest + [full] + se
    # ---- 4. 反向 DCF（新）----
    rv = lambda p, start=1: next(r for r in range(start, RV.max_row + 1) if str(RV.cell(r, 1).value or "").strip().startswith(p))
    ra, rcmp, rsum, yh = rv("A. DCF"), rv("B. Comparable"), rv("C. 摘要"), rv("項目（情境 B）")
    rb_, rbb, rcc = rv("(a)"), rv("(b)"), rv("(c)")
    m1, m0 = rv("現價隱含的營業利益率", rb_), rv("模型：基準的等效營業利益率", rb_)
    s1, rv1, rv0, g1, g0b = rv("現價隱含的放大倍數", rbb), rv("現價隱含的 FY2028 營收", rbb), rv("模型：FY2028 營收", rbb), rv("現價隱含的 FY2028 營收年增率", rbb), rv("模型：FY2028 營收年增率", rbb)
    w1, w0 = rv("現價隱含的 WACC", rcc), rv("模型 WACC", rcc)
    a_rows = [rv(p, ra) for p in ("支持現價所需的評價日每股價值", "淨現金滾到評價日", "評價用稀釋股數", "支持現價所需的企業價值", "模型企業價值", "EV 需放大", "模型：情境 B 評價日每股價值")]
    afmt = ["n2", "n1", "n3", "n1", "n1", "x1", "n2"]
    pe1, rc1, pe2 = rv("只看 Comparable：支持現價所需的前瞻本益比", rcmp), rv("只看 Comparable：所需的相對倍數係數", rcmp), rv("加權（DCF 不變）", rcmp)
    yrs_rows = list(range(yh + 1, yh + 12))
    yfmt = ["n1", "n1", "p1", "p1", "n3", "n3", "n3", "n3", "n1", "n1", "n1"]
    rev_ = {"id": "revdcf", "title": "反向 DCF：現價隱含什麼", "intro": "@RevDCF!A2",
            "blocks": [
                {"type": "note", "parts": [f"@RevDCF!A{rsum + 1}"]},
                {"type": "note", "parts": [f"@RevDCF!A{rsum + 2}"]},
                {"type": "table", "caption": "現金流折現法（情境 B 路徑）：一次只動一個因素，反解到現價", "head": ["項目", "現價隱含", "模型基準"],
                 "rows": [["預測期事業利益率（FY2028–FY2032 統一一個值；模型欄為等效值，終值年權重大）", f"@RevDCF!B{m1}|p1", f"@RevDCF!B{m0}|p1"],
                          ["現金流放大倍數（FY2028 起各年同乘；模型基準為一倍）", f"@RevDCF!B{s1}|n2", "—"],
                          ["FY2028 營收（十億日圓）", f"@RevDCF!B{rv1}|n0", f"@RevDCF!B{rv0}|n0"],
                          ["FY2028 營收年增率", f"@RevDCF!B{g1}|p1", f"@RevDCF!B{g0b}|p1"],
                          ["折現率（WACC）", f"@RevDCF!B{w1}|p2", f"@RevDCF!B{w0}|p2"]]},
                {"type": "table", "caption": f"@RevDCF!A{ra}", "head": ["項目", "值"], "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|{f}"] for r, f in zip(a_rows, afmt)]},
                {"type": "table", "caption": f"@RevDCF!A{rcmp}", "head": [f"@RevDCF!{c}{rcmp + 1}" for c in "ABCD"],
                 "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|{f}", f"@RevDCF!C{r}|{f}", f"@RevDCF!D{r}|sp1"] for r, f in ((pe1, "n1"), (rc1, "n2"), (pe2, "n1"))]},
                {"type": "table", "caption": f"@RevDCF!A{yh}", "head": [f"@RevDCF!{c}{yh}" for c in "ABCDEFGH"], "collapse": True,
                 "summary": "展開：情境 B 各年數字（反解的計算基礎；演算過程見 RevDCF 頁）",
                 "rows": [[f"@RevDCF!A{r}"] + [f"@RevDCF!{c}{r}|{f}" for c in "BCDEFGH"] for r, f in zip(yrs_rows, yfmt)]}]}
    # ---- 5. 損益：標準損益表 ----
    p0 = find("PL", "標準損益表（v0.3.5")
    ph = next(r for r in range(p0, PLW.max_row + 1) if PLW.cell(r, 1).value == "項目")
    plr = lambda lab_, start=ph: next(r for r in range(start, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip() == lab_)
    keys = ["營收", "營收年增率（YoY）", "營業成本（COGS）", "毛利", "毛利率", "營業費用（販管費＋研發 − 持分法損益）", "營業利益（IFRS 事業利益）", "營業利益率（事業利益率）",
            "其他收入（淨額：其他營業收支＋金融收支）", "稅前利益", "所得稅", "稅後淨利（歸屬母公司）", "淨利率", "股數（十億股，分割後）", "EPS（基本；分割後）"]
    memo = ["（備註）非支配股東損益", "（備註）IFRS 営業利益（含其他營業收支；只列實績）"]
    PC = "CDEFGHIJK"
    rowx = lambda r: [f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmt_u(PLW.cell(r, 2).value)}" for c in PC]
    r_ratio = plr("比率"); r_chk = next(r for r in range(r_ratio, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("核對（應為 0"))
    r_end = next(r for r in range(r_chk, PLW.max_row + 1) if not PLW.cell(r, 1).value) - 1
    is_tbl = {"type": "table", "caption": "標準損益表（十億日圓；EPS 為日圓）：FY2023A–FY2025A 實績、FY2026E 錨點、FY2027E–FY2031E 情境 B（基準；GAAP＝IFRS，營業利益列為事業利益）",
              "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in PC],
              "rows": [rowx(plr(k)) for k in keys] + [[f"@PL!A{plr(k)}"] + [f"@PL!{c}{plr(k)}|n1" for c in "CDE"] + [""] * 6 for k in memo]
                      + [{"sub": f"@PL!A{r_ratio}"}] + [rowx(r) for r in range(r_ratio + 1, r_chk)]}
    chk_tbl = {"type": "table", "caption": f"@PL!A{r_chk}", "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in PC], "collapse": True,
               "summary": "展開：核對列（歷史捨入差、預測期對 PL 情境 B、FY2030E–FY2031E 的推法差額）",
               "rows": [[f"@PL!A{r}"] + [f"@PL!{c}{r}|{'p1' if '占營收' in str(PLW.cell(r, 1).value) else 'n3'}" for c in PC] for r in range(r_chk + 1, r_end + 1)]}
    rv_, gm_, opm_, nm_ = plr("營收"), plr("毛利率"), plr("營業利益率（事業利益率）"), plr("淨利率")
    is_chart = {"type": "theme", "title": "營收（十億日圓）與毛利率、事業利益率、淨利率：FY2023A–FY2031E（情境 B）", "x": f"@PL!C{ph}:K{ph}",
                "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
                "bars": [{"name": f"@PL!A{rv_}", "y": f"@PL!C{rv_}:K{rv_}", "color": "#6fc7cb"}],
                "lines": [{"name": f"@PL!A{gm_}", "y": f"@PL!C{gm_}:K{gm_}", "axis": "r", "color": "#0b7a80", "labels": True},
                          {"name": f"@PL!A{opm_}", "y": f"@PL!C{opm_}:K{opm_}", "axis": "r", "color": "#c29a2b"},
                          {"name": f"@PL!A{nm_}", "y": f"@PL!C{nm_}:K{nm_}", "axis": "r", "color": "#6f6bae", "dash": True}],
                "caption": "長條＝營收；三條線＝毛利率、事業利益率、淨利率（右軸，IFRS）。FY2024A 淨利含其他營業費用（減損等）、FY2025A 含本社大樓土地出售利益；FY2026E 為公司預測錨點；預測期毛利率為推算值；FY2030E–FY2031E 為收斂法延伸期。"}
    old_pl = S["pl"]["blocks"]
    pls = [b for b in old_pl if b["type"] == "pl"]
    assert len(pls) == 2, [b["type"] for b in old_pl]
    old_blocks = [dict(pls[0], collapse=True, summary="展開：三情境損益摘要（FY2025–FY2029；營收、FM、事業利益、EPS）")]
    old_blocks += [dict(b, collapse=True, summary="展開：三情境摘要補充") for b in old_pl if b["type"] in ("kv", "note")]
    old_blocks += [dict(pls[1], collapse=True, summary="展開：三情境延伸到 FY2032（收斂法）")]
    S["pl"].pop("intro", None)
    S["pl"].update(title="損益（標準損益表）",
                   blocks=[{"type": "note", "parts": [f"@PL!A{p0 + 1}"]}, is_chart, is_tbl, {"type": "note", "parts": [f"@PL!A{p0 + 2}"]}, {"type": "note", "parts": [f"@PL!A{r_end + 2}"]}, chk_tbl] + old_blocks)
    # ---- 6. 驅動因子：分部營收推導（情境 B）----
    r_eff = prow("有效顆數 ＝ NVIDIA ＋ 面積係數 × 非 NVIDIA")
    secA, secB, secC = prow("A. ABF 需求", False), prow("B. 其他事業", False), prow("C. 全公司", False)
    seg = lambda rs: [[f"@PL!A{r}", f"@PL!B{r}"] + [f"@PL!{c}{r}|{fmt_u(PLW.cell(r, 2).value)}" for c in "DEFG"] for r in rs
                      if PLW.cell(r, 1).value and not str(PLW.cell(r, 1).value).startswith(("檢查", "營收調整", "事業利益調整"))]
    endC = prow("AI 用 ABF 營收 ÷ 有效顆數", False)
    seg_tbl = {"type": "table", "caption": "分部營收推導（情境 B；十億日圓）：ファンクショナルマテリアルズ（ABF）＋調味料・食品＋冷凍食品＋ヘルスケア等＋その他＝合併營收",
               "head": [f"@PL!A{hdrB}", f"@PL!B{hdrB}"] + [f"@PL!{c}{hdrB}" for c in "DEFG"],
               "rows": [{"sub": "一、FM 營收＝FY2026 錨點 ×（AI 用：有效顆數成長 × 面積層數 × 份額；非 AI：成長率）× 需求滿足率 × ABF 價格指數"}] + seg(range(secA + 1, secB))
                       + [{"sub": "二、其他事業（不隨 AI 情境）"}] + seg(range(secB + 1, secC))
                       + [{"sub": "三、全公司"}] + seg(range(secC + 1, endC + 1))}
    seg_chart = {"type": "theme", "title": "FM 營收（情境 B，十億日圓）與有效加速器顆數（百萬顆，右軸）", "x": f"@PL!D{hdrB}:G{hdrB}", "fmt": "n0", "fmt_r": "n1", "stack": True, "pad": 1.25,
                 "bars": [{"name": f"@PL!A{r_fmai}", "y": f"@PL!D{r_fmai}:G{r_fmai}", "color": "#0b7a80"}, {"name": f"@PL!A{r_fmn}", "y": f"@PL!D{r_fmn}:G{r_fmn}", "color": "#6fc7cb"}],
                 "lines": [{"name": [f"@PL!A{r_eff}"], "y": f"@PL!D{r_eff}:G{r_eff}", "axis": "r", "color": "#3b3838", "labels": True}],
                 "caption": "長條＝FM 營收的 AI 用與非 AI 部分；折線＝有效加速器顆數（NVIDIA ＋ 面積係數 × 非 NVIDIA；情境 B 在 CY2028 起顆數低於原始值）。FY2027 起需求高於可出貨量上限，出貨由上限決定。FY2026 為校準年。"}
    dr = S["drivers"]["blocks"]
    S["drivers"].update(title="驅動因子：分部營收推導")
    S["drivers"].pop("intro", None)
    S["drivers"]["blocks"] = [{"type": "note", "parts": [f"@README!A{RM['drv']}"]}, seg_chart, seg_tbl, {"type": "h3", "text": "情境驅動與單一參數"}] + dr
    # ---- 7. 對帳：三層協調（C0 實績、讀取層級、層級之間）----
    c0 = fnd(CW, "三層協調 C0"); s1_ = fnd(CW, "1. 本模型讀哪一層"); s2 = fnd(CW, "2. 讀取層級"); s3 = fnd(CW, "3. 讀取層級與層級之間的結果")
    sv = fnd(CW, "單值輸入：本模型 vs 層級值", s2)
    a4 = fnd(CW, "4. C0（實績）"); t4 = fnd(CW, "C0（實績）：公司實績 vs 層級隱含值", a4); e4 = fnd(CW, "C0（實績）結果摘要", t4)
    a4a, a4b, rref = fnd(CW, "4a. 公司已公布實績", a4), fnd(CW, "4b. 層級隱含值", a4), fnd(CW, "參考：本模型 FY2026 錨點", a4)
    cnt = fnd(CW, "對帳中項目數", e4)
    isnum = lambda v: isinstance(v, (int, float)) and not isinstance(v, bool)
    cfm = lambda r: "p1" if CW.cell(r, 6).value == "%" else "n0" if CW.cell(r, 6).value == "月" else "n2" if CW.cell(r, 6).value in ("x", "US$B") else "n1"
    c0a = {"type": "table", "caption": f"@Coord!A{t4}", "head": [f"@Coord!{c}{t4 + 1}" for c in "ABC"] + [f"@Coord!F{t4 + 1}"],
           "rows": [[f"@Coord!A{r}", f"@Coord!B{r}" + ("|n1" if isnum(CW.cell(r, 2).value) else "|"), f"@Coord!C{r}" + ("|n1" if isnum(CW.cell(r, 3).value) else "|"), f"@Coord!F{r}|"]
                    for r in range(t4 + 2, rref - 1) if CW.cell(r, 1).value]}
    c0det = {"type": "table", "caption": "公司實績對照各列的口徑、來源與證據等級（展開看）", "collapse": True, "summary": "展開：各列的口徑、來源與證據等級",
             "head": [f"@Coord!{c}{t4 + 1}" for c in "AG"], "rows": [[f"@Coord!A{r}", f"@Coord!G{r}|"] for r in range(t4 + 2, rref - 1) if CW.cell(r, 7).value]}
    c0in = {"type": "table", "caption": f"@Coord!A{a4a}", "collapse": True, "summary": "展開：公司已公布實績（決算概要）的數值、來源與證據等級",
            "head": [f"@Coord!{c}{a4a + 1}" for c in "ABFG"], "rows": [[f"@Coord!A{r}", f"@Coord!B{r}" + (f"|{cfm(r)}" if isnum(CW.cell(r, 2).value) else "|"), f"@Coord!F{r}|", f"@Coord!G{r}|"]
                                                                  for r in range(a4a + 2, a4b - 1) if CW.cell(r, 1).value]}
    c0lay = {"type": "table", "caption": f"@Coord!A{a4b}", "collapse": True, "summary": "展開：層級隱含值（Siliconomics 現行版）的組成",
             "head": [f"@Coord!{c}{a4b + 1}" for c in "ABFG"], "rows": [[f"@Coord!A{r}", f"@Coord!B{r}" + (f"|{cfm(r)}" if isnum(CW.cell(r, 2).value) else "|"), f"@Coord!F{r}|", f"@Coord!G{r}|"]
                                                                   for r in range(a4b + 2, t4 - 1) if CW.cell(r, 1).value]}
    c0ref = {"type": "table", "caption": f"@Coord!A{rref}", "collapse": True, "summary": "展開：模型錨點對已公布實績與公司第二季決算時的上修（參考，不計入實績對帳）",
             "head": ["項目", "值", "單位", "口徑、來源、證據等級"], "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|{cfm(r)}", f"@Coord!F{r}|", f"@Coord!G{r}|"] for r in range(rref + 1, e4 - 1) if CW.cell(r, 1).value]}
    act = [{"type": "h3", "text": "三層協調：公司實績對層級隱含值"}, {"type": "note", "parts": [f"@Coord!A{a4 + 1}"]}]
    act += [{"type": "note", "parts": [f"@Coord!A{r}"]} for r in range(e4 + 1, cnt)]
    act += [c0a, {"type": "note", "parts": ["對帳中項目數：", f"@Coord!G{cnt}|n0", "。"]}, c0det, c0in, c0lay, c0ref, {"type": "note", "parts": [f"@Coord!A{a4 + 2}"]}]
    same = {"type": "table", "caption": "讀取層級（同源，不判定）：本模型的產業量讀自 Siliconomics（或由本模型校準後寫入），與層級值逐年對照的差距無檢驗力，不計入對帳中",
            "head": [f"@Coord!{c}{s2 + 1}" for c in "ABCDEF"],
            "rows": [[f"@Coord!A{r}"] + [f"@Coord!{c}{r}|" + ("sp1" if "差距" in str(CW.cell(r, 1).value) else "" if "判定" in str(CW.cell(r, 1).value) else ("n3" if CW.cell(r, 6).value in ("x", "百萬顆") else "n1")) for c in "BCDE"] + [f"@Coord!F{r}|"]
                     for r in range(s2 + 2, sv) if CW.cell(r, 1).value]}
    svt = {"type": "table", "caption": f"@Coord!A{sv}", "head": [f"@Coord!{c}{sv + 1}" for c in "ABCDG"],
           "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|" + ("p1" if CW.cell(r, 6).value == "%" else "n2"), f"@Coord!C{r}|" + ("p1" if CW.cell(r, 6).value == "%" else "n2"), f"@Coord!D{r}|sp1", f"@Coord!G{r}|"]
                    for r in range(sv + 2, s3 - 1) if CW.cell(r, 1).value]}
    same_d = {"type": "table", "caption": f"@Coord!A{s2}", "collapse": True, "summary": "展開：各列的單位、口徑、來源與證據等級", "head": [f"@Coord!{c}{s2 + 1}" for c in "AFG"],
              "rows": [[f"@Coord!A{r}", f"@Coord!F{r}|", f"@Coord!G{r}|"] for r in range(s2 + 2, sv) if CW.cell(r, 6).value or CW.cell(r, 7).value]}
    lay = {"type": "table", "caption": f"@Coord!A{s1_}", "collapse": True, "summary": "展開：本模型讀哪一層、哪一版、用在哪裡", "head": [f"@Coord!{c}{s1_ + 1}" for c in "ABCD"],
           "rows": [[f"@Coord!A{r}", f"@Coord!B{r}|" + ("p1" if CW.cell(r, 1).value == "每顆載板面積 × 層數年增" else "n2" if isnum(CW.cell(r, 2).value) else ""), f"@Coord!C{r}|", f"@Coord!D{r}|"]
                    for r in range(s1_ + 2, s2 - 1) if CW.cell(r, 1).value]}
    rd_ = [{"type": "h3", "text": "讀取層級（同源，不判定）與層級之間（參考）"}, {"type": "note", "parts": [f"@Coord!A{c0 + 1}"]}]
    rd_ += [{"type": "note", "parts": [f"@Coord!A{r}"]} for r in range(s3 + 1, s3 + 4) if CW.cell(r, 1).value]
    rd_ += [same, svt, same_d, lay]
    S["recon"]["blocks"] = act + rd_ + [{"type": "h3", "text": "由上而下與由下而上的對照（原有）"}] + S["recon"]["blocks"]
    S["recon"]["title"] = "對帳：三層協調、公司實績與先驗比對"
    # ---- 8. 資料狀態 ----
    st = S["status"]["blocks"]
    st[0:0] = [{"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RM['ver']}"]}]
    order = ["thesis", "valuation", "sens", "revdcf", "pl", "drivers", "recon", "status"]
    S["revdcf"] = rev_
    cfg["sections"] = [S[k] for k in order]


_v13()
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", OUT)
