# v0.3.2：套用個股 HTML 版面規格 v1.0（docs/plan/個股HTML版面規格.md；只改展示層）
# 由 html_src.json（舊版面設定，內容不動）經共用轉換器 tools/company_html/relayout.py 產生 html.json（勿手改 html.json）。
# 用法：python3 amat/gen_html_json.py（讀 amat/dist 內唯一的 xlsx）
# 前處理（本檔）：把設備股模板的舊版面調成轉換器預期的形狀——單一參數表改作「驅動 → 推導量對照」（進附錄）、
#   情境驅動表與敏感度目標價參數表的列號依現況重新定位、估值章的 kv 改成具名表格、舊法規則值與同業名單摺疊、
#   補 DCF 折現率表與目標價敏感度龍捲風圖；
# 後處理：還原驅動因子章說明、EPS 龍捲風圖標題（轉換器預設寫「加權目標價」）、橫幅中間格＝只看 Comparable（市場主流同業法），調整後值進附註。
import copy, glob, json, os, re, sys
import openpyxl
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools", "company_html"))
from relayout import relayout, _unit_fmt  # noqa: E402

X = glob.glob(os.path.join(HERE, "dist", "*.xlsx"))
assert len(X) == 1, X
X = X[0]
cfg = json.load(open(os.path.join(HERE, "html_src.json"), encoding="utf-8"))
wb = openpyxl.load_workbook(X, data_only=True)


def text(spec):
    m = re.match(r"^@([^!]+)!([A-Z]+\d+)", spec) if isinstance(spec, str) else None
    return str(wb[m.group(1)][m.group(2)].value or "") if m else str(spec or "")


def row(sh, prefix, exact=False, start=1):
    ws = wb[sh]
    for r in range(start, ws.max_row + 1):
        v = str(ws.cell(r, 1).value or "").strip()
        if (v == prefix) if exact else v.startswith(prefix):
            return r
    raise KeyError((sh, prefix))


IN = {wb["In"].cell(r, 1).value: r for r in range(4, wb["In"].max_row + 1) if wb["In"].cell(r, 1).value}


def in_row(k, fmt):
    r = IN[k]
    return [f"@In!B{r}", f"@In!E{r}|{fmt}", f"@In!G{r}|tag"]


def kv_table(b, caption=None, collapse=False, summary=None):
    t = {"type": "table", "head": ["項目", "數值"], "rows": [[k, v] for k, v in b["items"]]}
    if caption:
        t["caption"] = caption
    if collapse:
        t.update(collapse=True, summary=summary)
    return t


def prep(cfg, spec):
    secs = {s["id"]: s for s in cfg["sections"]}
    # 驅動因子：單一參數表 → 驅動 → 推導量對照（轉換器移到附錄）；情境驅動表不摺疊
    dr = secs["drivers"]
    for b in dr["blocks"]:
        cap = str(b.get("caption", ""))
        if cap.startswith("單一參數"):
            b["caption"] = "驅動 → 推導量對照（單一參數的輸入、基準值、推導量與因果方向）"
        elif cap.startswith("情境驅動"):
            b.pop("collapse", None)
            b.pop("summary", None)
            # 列號依 In 頁現況重新定位（舊設定的列號在 In 頁插入 Wattonomics 錨定係數後已位移）
            ws = wb["In"]
            h = row("In", "情境驅動對照") + 1
            assert ws.cell(h, 1).value == "鍵", ws.cell(h, 1).value
            fm = lambda r: "yr" if ws.cell(r, 3).value == "年" else _unit_fmt(ws.cell(r, 3).value)
            b["head"] = [f"@In!B{h}", f"@In!D{h}", f"@In!E{h}", f"@In!F{h}", f"@In!C{h}"]
            b["rows"] = [[f"@In!B{r}", f"@In!D{r}|{fm(r)}", f"@In!E{r}|{fm(r)}", f"@In!F{r}|{fm(r)}", f"@In!C{r}"]
                         for r in range(h + 1, ws.max_row + 1) if ws.cell(r, 1).value]
    intro = dr.get("intro")
    # 估值：依規格重排
    va = secs["valuation"]
    old = va["blocks"]
    nb, i = [], 0
    caps = [text(b.get("caption", "")) for b in old]
    k_dcf = next(j for j, c in enumerate(caps) if c.startswith(spec["cap_dcf"]))
    k_comp = next(j for j, c in enumerate(caps) if c.startswith(spec["cap_comp"]))
    k_old = next(j for j, b in enumerate(old) if b["type"] == "h3" and str(b.get("text", "")).startswith(spec["h3_old"]))
    conc = old[0]
    assert text(conc.get("caption", "")).endswith("結論"), conc
    head = old[:k_dcf]
    bars = next(b for b in head if b["type"] == "bars")
    ca = row("Val", "Comparable 目標價（相對倍數調整後）")
    if not any(it["value"] == f"@Val!B{ca}" for it in bars["items"]):   # 加權目標價實際使用的 Comparable（相對倍數調整後）也畫出，避免讀成加權價高於兩法
        k = next(j for j, it in enumerate(bars["items"]) if text(it["label"]).startswith("Comparable 目標價"))
        bars["items"].insert(k + 1, {"label": f"@Val!A{ca}", "value": f"@Val!B{ca}", "series": 2})
    nb.append(bars)
    nb.append(conc)
    for j, b in enumerate(head[1:], 1):
        if b["type"] == "note":
            nb.append(b)
    rm_h3 = next(b for b in head if b["type"] == "h3" and str(b.get("text", "")).startswith("相對倍數調整"))
    rm_kv = head[head.index(rm_h3) + 1]
    assert rm_kv["type"] == "kv"
    nb.append({"type": "table", "caption": "DCF 折現率與成長率（股東權益現金流折現；直接顯示）", "head": ["項目", "數值", "標記"],
               "rows": [in_row(k, f) for k, f in spec["dcf_rates"]] + [[f"@Val!A{r}", f"@Val!B{r}|p1", ""] for r in spec.get("dcf_rate_val", [])]})
    nb.append({"type": "h3", "text": "DCF 彙總"})
    nb.append(old[k_dcf])
    nb.append({"type": "h3", "text": "Comparable（同業前瞻本益比）"})
    comp_kv = old[k_comp + 1]
    assert comp_kv["type"] == "kv"
    nb.append(kv_table(comp_kv, caption="Comparable 彙總（同業前瞻本益比中位數 × 前瞻每股盈餘）"))
    nb.append(kv_table(rm_kv, caption=rm_h3["text"]))
    peers = dict(old[k_comp], collapse=True, summary="展開：同業前瞻本益比（各家）")
    nb.append(peers)
    nb.append({"type": "h3", "text": old[k_old]["text"]})
    for b in old[k_old + 1:]:
        if b["type"] == "bars":
            continue   # 舊法隱含股價長條圖不放（轉換器會把所有長條圖移到章首，與目標價長條圖混淆）
        if b["type"] == "kv":
            nb.append(kv_table(b, collapse=True, summary="展開：舊法的價格與市場隱含倍數"))
        elif b["type"] == "table":
            if not b.get("collapse"):
                c = re.sub(r"^[\d.\s]+", "", text(b.get("caption", ""))).strip()
                b = dict(b, collapse=True, summary="展開：" + (c if c and not re.search(r"\d", re.sub(r"(FY|CY)\d{4}", "", c)) else "舊法明細表"))
            nb.append(b)
        elif b["type"] == "note":
            nb.append(dict(b, collapse=True, summary="展開：舊法說明"))
        else:
            nb.append(b)
    va["blocks"] = nb
    # 對帳：給表格標題，主要列由 SPEC main 指定
    for b in secs["recon"]["blocks"]:
        if b["type"] == "table" and not b.get("caption"):
            b["caption"] = "對帳要點"
    # 敏感度：先放目標價龍捲風圖，再放 EPS 龍捲風圖
    se = secs["sens"]
    tor = {}
    for b in se["blocks"]:
        if b["type"] == "tornado":
            tor[b["value_col"]] = b["title"].split("（")[0].replace("相對基準的變動", "").strip()
    ws = wb["Sens"]
    h0 = row("Sens", "目標價參數")
    for b in se["blocks"]:   # 目標價參數案例表：列號依 Sens 頁現況重新定位（舊設定在插入 Wattonomics 案例後可能位移）
        rg = b.get("rowgen") if b["type"] == "table" else None
        if rg and isinstance(b.get("caption"), str) and b["caption"].startswith("@Sens!A") and abs(int(b["caption"][7:]) - h0) <= 6:
            hh = h0 + 1
            assert str(ws.cell(hh, 1).value).strip() == "案例", ws.cell(hh, 1).value
            end = hh + 1
            while ws.cell(end + 1, 1).value not in (None, ""):
                end += 1
            b["caption"] = f"@Sens!A{h0}"
            b["head"] = [f"@Sens!{c}{hh}" for c in rg["cols"]]
            rg["rows"] = f"{hh + 1}-{end}"
    t0 = spec["tp_sens"]
    se["blocks"].insert(0, {"type": "tornado", "title": "加權目標價", "sheet": "Sens", "rows": t0["rows"], "label_col": "A", "value_col": t0["col"], "fmt": "sp1"})
    tor[t0["col"]] = None
    for b in se["blocks"]:
        if b["type"] == "table" and not b.get("collapse"):
            b.update(collapse=True, summary="展開：目標價參數的敏感度案例表（DCF 目標價、Comparable 目標價、加權目標價）")
    scen_kv = next((b for b in secs["thesis"]["blocks"] if b["type"] == "kv"), None)
    return intro, tor, scen_kv


def post(cfg, intro, tor, scen_kv, spec):
    secs = {s["id"]: s for s in cfg["sections"]}
    if intro:   # 轉換器把章說明當 note 放進某一塊；還原為章說明
        secs["drivers"]["intro"] = intro
        secs["drivers"]["blocks"] = [b for b in secs["drivers"]["blocks"] if not (b["type"] == "note" and b.get("parts") == [intro])]
    # 參數表格式：轉換器只認單位「%」，「%/年」等也改百分比；小於 1 的一般數值用三位小數（避免 0.0／0.1 誤讀）
    ws = wb["In"]
    for b in secs["drivers"]["blocks"]:
        for r_ in b.get("rows", []) if b["type"] == "table" else []:
            m = re.match(r"^@In!B(\d+)$", r_[0]) if isinstance(r_, list) and isinstance(r_[0], str) else None
            if not m:
                continue
            r = int(m.group(1))
            u = str(ws.cell(r, 3).value or "")
            vals = [ws.cell(r, c).value for c in (4, 5, 6)]
            for j in (1, 2, 3):
                cell, _, f = r_[j].partition("|")
                if u.startswith("%"):
                    f = "p1"
                elif f == "n1" and any(isinstance(x, (int, float)) and 0 < abs(x) < 1 for x in vals):
                    f = "n3"
                elif f == "n1" and any(isinstance(x, float) and abs(x) < 10 and x != int(x) for x in vals):
                    f = "n2"
                r_[j] = f"{cell}|{f}"
    if scen_kv:   # 舊版命題章的情境定義：放在重點觀察之後（摺疊）
        secs["thesis"]["blocks"].append(dict(scen_kv, collapse=True, summary="展開：三情境的定義"))
    for b in secs["sens"]["blocks"]:
        if b["type"] == "tornado" and tor.get(b["value_col"]):
            base = tor[b["value_col"]] + " 相對基準的變動"
            b["title"] = base + ("（未達兩位數百分比）" if b.get("collapse") else "（變動達兩位數百分比的案例；一次改一個輸入，整本重算）")
        if b["type"] == "tornado" and b.get("collapse"):
            b["summary"] = "展開：" + (tor.get(b["value_col"]) or "加權目標價") + " 變動未達兩位數百分比的案例"
    # 橫幅（補充規則 A）：中間格＝只看 Comparable 目標價（市場主流同業法，未做相對倍數調整）與其上檔空間；
    #   進加權的相對倍數調整後 Comparable 放附註（數字由 xlsx 帶出）
    v = lambda p, **kw: row("Val", p, **kw)
    tp, ca, cu, px, up, rec = v("目標價（DCF 與 Comparable 加權）", exact=True), v("Comparable 目標價（相對倍數調整後）"), v("Comparable 目標價", exact=True), v("現價（"), v("上檔空間：加權目標價"), v("投資建議", exact=True)
    rd = row("README", "HTML 橫幅附註用：Comparable 目標價（相對倍數調整後）")
    uc = v("上檔空間：Comparable")
    cfg["hero"] = {"rec": f"@Val!B{rec}",
                   "tiles": [[f"@Val!A{tp}", f"@Val!B{tp}|n2", ["上檔空間 ", f"@Val!B{up}|sp1"]],
                             ["只看 Comparable 目標價（市場主流同業法）", f"@Val!B{cu}|n2", ["上檔空間 ", f"@Val!B{uc}|sp1"]],
                             [f"@Val!A{px}", f"@Val!B{px}|n2", ""]],
                   "note": ["投資建議依加權目標價判定：", f"@Val!E{rec}", "。進加權的 Comparable 為相對倍數調整後 ", f"@Val!B{ca}|n2",
                            "（上檔空間 ", f"@README!B{rd}|sp1", "）。"]}


def in_skip():
    ws = wb["In"]
    out = []
    for r in range(4, ws.max_row + 1):
        a = str(ws.cell(r, 1).value or "")
        if a and ws.cell(r, 2).value in (None, "") and a.startswith("情境驅動對照"):
            out.append(re.split(r"[.\s]", a)[0])
    return out


def run(SPEC, PRE):
    intro, tor, scen_kv = prep(cfg, PRE)
    SPEC = dict(SPEC, in_skip=SPEC.get("in_skip", []) + in_skip())
    out = relayout(cfg, X, SPEC)
    post(out, intro, tor, scen_kv, PRE)
    json.dump(out, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("ok")



# ---------------- 應用材料專屬：參數依業務性質分塊（半導體系統／中國與出口管制／AGS 服務與其他／全公司） ----------------
PRE = {"cap_dcf": "0.1 DCF", "cap_comp": "0.2 Comparable", "h3_old": "舊法",
       "dcf_rates": [("tp_r", "p1"), ("tp_g", "p1"), ("tp_roe", "p1")],
       "dcf_rate_val": [row("Val", "再投資率（tp_g ÷ tp_roe")],
       "tp_sens": {"rows": f"{row('Sens', '折現率 +1pt')}-{row('Sens', 'DCF 權重 0.8')}", "col": "F"}}
_eq = lambda p: [f"{p}{s}_{y}" for s in "ABC" for y in range(25, 30)]
SPEC = {
 "in_groups": [
  ["半導體系統（晶圓廠設備；跟著設備市場走）", [
    ["設備市場：先進邏輯、其他代工與邏輯、其他前段（設備市場模型，三情境）", _eq("la") + _eq("lo") + _eq("of")],
    ["設備市場：DRAM（含 HBM）與 NAND（設備市場模型，三情境）", _eq("dr") + _eq("nd")],
    ["會計年度換算與見頂年度", ["fyw", "pk_A", "pk_B", "pk_C"]],
    ["終端市場權重與份額係數（FY2026 錨點拆分）", ["q4_dr", "q4_nd", "m_dr", "m_nd", "m_fl", "sf"]],
    ["第四季指引（半導體系統）", ["g_ss_q4", "q4g_ss"]],
    ["對帳用：SEMI 與公司說法（不進盈餘）", ["semi_wfe26", "semi_fl26", "semi_dr26", "semi_nd26", "ai_inc27", "co_ai27"]],
    ["Wattonomics 錨定係數（只進敏感度）", ["wta_27", "wta_28", "wta_29"]]]],
  ["中國與出口管制", [
    ["關係企業規則與中國營收", ["aff_hit", "aff_A", "aff_B", "aff_C", "cn_sh", "cn_cut"]]]],
  ["AGS 服務與其他（含顯示器）", [
    ["AGS 服務", ["g_ags_A", "g_ags_B", "g_ags_C", "g_ags_q4", "q4g_ags"]],
    ["其他（主要為顯示器）", ["g_oth_A", "g_oth_B", "g_oth_C", "g_oth_q4"]]]],
  ["全公司（指引、成本費用、股價與目標價規則）", [
    ["FY2026 第四季公司指引（錨點）", ["g_rev_q4", "g_eps_q4", "g_gm_q4", "g_opx_q4"]],
    ["營業成本與營業費用（兩點推導）", ["cogs25", "cogs26", "cv", "cF", "phic", "opex25", "opex26", "ov", "oF", "phi"]],
    ["淨利、股數與回購", ["nioi", "sh_q4", "bb"]],
    ["共識與股價", ["cons_rev26", "cons_eps26", "cons_rev27", "cons_eps27", "px", "sa_fpe"]],
    ["目標價規則與折現率", ["tp_r", "tp_g", "tp_roe", "tp_w_dcf", "tp_wA", "tp_wB", "tp_wC", "nim10", "cons_tp", "cons_feps"]]]]],
 "in_skip": [],
 "main": {"0.1 DCF": ["DCF 情境每股價值", "DCF 情境價值推到 12 個月後", "DCF 目標價"],
          "對帳要點": ["FY2026 non-GAAP EPS", "FY2027 營收", "FY2027 non-GAAP EPS", "半導體系統", "AGS FY2026 年增"]},
}
if __name__ == "__main__":
    run(SPEC, PRE)
