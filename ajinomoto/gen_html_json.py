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
    ["FM 錨點與實績", ["f_fm", "f_fmbp", "q1_fm", "q1_fmbp", "a_fm", "a_fmbp"]],
    ["Wattonomics 錨定係數（只用於敏感度）", ["wta_27", "wta_28", "wta_29"]]]],
  [B_CORE, [
    ["成長率與利益率", ["g_sf", "m_sf", "g_ff", "m_ff", "g_ho", "m_ho", "g_cc"]],
    ["各事業錨點與實績", ["f_sf", "f_ff", "f_ho", "f_oth", "f_sfbp", "f_ffbp", "f_hobp", "f_othbp", "f_cc", "a_sf", "a_ff", "a_ho", "a_oth", "a_sfbp", "a_ffbp", "a_hobp", "a_othbp", "a_cc"]]]],
  [B_CO, [
    ["全公司錨點與實績", ["f_rev", "f_bp", "f_oi", "f_ni", "f_eps", "q1_rev", "q1_bp", "q1_ni", "q1_eps", "a_rev", "a_bp", "a_oi", "a_ni", "a_eps", "a_da", "a_cx"]],
    ["淨利與股數", ["ni_bp", "sh_g", "so"]],
    ["股價與市場倍數", ["px", "pe_cf", "fpe", "cons_feps", "cons_tp"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "rf_us", "beta", "beta_u", "beta_u_food", "beta_u_chem", "beta_w_fm", "beta_du", "beta_de", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "bs_date", "stub_frac", "stub_div", "sh_v"]],
    ["自由現金流（折舊、資本支出、營運資金；da_rev、cx_rev 另供 PL 頁自由現金流列）", ["fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc", "da_rev", "cx_rev"]],
    ["目標價規則與情境權重", ["tp_w_dcf", "tp_wA", "tp_wB", "tp_wC"]],
    ["收斂期（FY2030 起）", ["tp_g", "nm10_k", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]]]]],
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

json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", OUT)
