# 由 html_src.json（v0.2.1 以前的 HTML 設定，內容不動）套用「個股 HTML 版面規格 v1.0」（docs/plan/個股HTML版面規格.md），產生 nanya_pcb/html.json。
# 用法：python3 nanya_pcb/gen_html_json.py [成品 xlsx]（省略時取 nanya_pcb/dist/ 內唯一的 xlsx）；勿手改 html.json。
# 轉換器 tools/company_html/relayout.py 之外，本檔只做本公司專屬的前處理（估值章重排）與後處理（橫幅、龍捲風標題、附錄）。
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

# ======== 本公司設定 ========
B_AI = "ABF 載板：AI 加速器載板"
B_ABF = "ABF 載板：交換器等其他 AI、非 AI 與產能"
B_OTH = "BT 載板與一般電路板"
B_CO = "全公司：股價、折現率、目標價規則、稅與現金流"
IN_GROUPS = lambda K: [
  [B_AI, [
    ["加速器顆數與 AI IT 資本支出（讀網通市場模型）", [k for k in K if re.match(r"^(u_|g29_|itc_)", k)]],
    ["情境顆數係數（個股自訂）", ["um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C"]],
    ["每顆載板金額與 ABF 溢價（讀 Ibiden 模型）", ["v26", "g_asp", "ibpp_27", "ibpp_28", "ibpp_29", "ibrel_27", "ibrel_28", "ibrel_29", "ib_sh26", "ib_w_ot", "v_ot"]],
    ["欣興份額（讀欣興模型；份額合計檢查用）", [k for k in K if re.match(r"^un_(nv|ot)_\d\d$", k)]],
    ["南電份額", ["sh_nv_26", "sh_nv_27", "sh_nv_28", "sh_nv_29", "sh_ot_26", "sh_ot_27", "sh_ot_28", "sh_ot_29"]],
    ["Wattonomics 錨定係數（只用於敏感度）", ["wta_27", "wta_28", "wta_29"]]]],
  [B_ABF, [
    ["FY2026 起點：ABF 占營收與 AI 占比", ["mx_abf", "ai_abf"]],
    ["交換器等其他 AI 載板（交換器市場，讀網通市場模型）與非 AI ABF", ["swm_26", "swm_27", "swm_28", "swm_29", "el_o", "g_nai"]],
    ["產能、組合與漲價", ["cap_27", "cap_28", "cap_29", "g_mix", "carry"]],
    ["ABF 增量現金毛利率", ["ft_abf"]]]],
  [B_OTH, [
    ["FY2026 起點：BT 占營收", ["mx_bt"]],
    ["營收成長（BT 隨情境）", ["g_bt27_A", "g_bt28_A", "g_bt29_A", "g_bt27_B", "g_bt28_B", "g_bt29_B", "g_bt27_C", "g_bt28_C", "g_bt29_C", "g_pcb"]],
    ["增量現金毛利率", ["ft_oth"]]]],
  [B_CO, [
    ["資本支出與折舊", ["cx27", "cx28", "cx29", "capex26", "da26", "life"]],
    ["營業費用、業外、稅與股數", ["phi", "nonop27", "nonop28", "nonop29", "tax", "sh_g"]],
    ["FY2026 錨點（季報實績與下半年假設）", ["q3_rev", "q3_gm", "q3_opx", "q4_g", "q4_gm", "q4_opx", "q_nonop", "q_sh"]],
    ["共識與券商（只作對照）", ["cons_eps26", "cons_eps27", "cons_eps28", "cons_rev27", "cons_rev28", "bk_eps27", "bk_eps28", "cons_feps", "cons_tp"]],
    ["股價、市場倍數與匯率", ["px", "fpe", "pe_ttm", "fx", "nim5"]],
    ["折現率（標準版 DCF：無風險利率、β、股權風險溢酬、WACC、永續成長率、RONIC）", FCFF_KEYS["rate"]],
    ["資本結構與淨現金（WACC 權重、EV 到每股價值）", FCFF_KEYS["bs"]],
    ["自由現金流假設（DCF：折舊、資本支出、營運資金）", FCFF_KEYS["fcf"]],
    ["目標價規則與收斂法", ["tp_w_dcf", "tp_wA", "tp_wB", "tp_wC", "tp_g", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]]]]]
BLOCKS_TO = {"單一參數": B_AI, "AI 加速器載板份額": B_AI, "@Units!A": B_AI}
RECON_CAP = "本模型對照券商、共識與市場口徑"
RECON_MAIN = ["FY2027 EPS vs FactSet", "FY2028 EPS vs FactSet", "FY2028 營收 vs FactSet", "FY2028 毛利率 vs 美銀", "模型前瞻 12 月 EPS"]
SUB_OVERRIDE = {"股東現金流＝EPS ×（1 − tp_g ÷ tp_roe）": "每股股東現金流（淨利＋折舊－資本支出；本模型 DCF 實際使用）"}
# ===========================


def find(sh, prefix, start=1, col=1, must=True):
    ws = wb[sh]
    for r in range(start, ws.max_row + 1):
        if str(ws.cell(r, col).value or "").strip().startswith(prefix):
            return r
    if must:
        raise KeyError((sh, prefix))


def sec(c, cid):
    return next(s for s in c["sections"] if s["id"] == cid)


def ctext(b):
    c = b.get("caption", "")
    m = re.match(r"^@([^!]+)!([A-Z]+\d+)$", c) if isinstance(c, str) else None
    return str(wb[m.group(1)][m.group(2)].value or "") if m else str(c or "")


V = lambda r, f="": f"@Val!B{r}" + (f"|{f}" if f else "")
VA = lambda r: f"@Val!A{r}"
INK = {str(wb["In"].cell(r, 1).value): r for r in range(4, wb["In"].max_row + 1) if wb["In"].cell(r, 1).value}
# ---- v0.5：標準版 DCF（FCFF＋WACC）的估值區塊——折現率表緊接 DCF 彙總，各情境 FCFF 明細摺疊（docs/plan/個股DCF標準規格.md、個股HTML版面規格 v1.0）----
def fcff_blocks(pxf="n0", amt="n1"):
    Vw, Dw = wb["Val"], wb["DCF"]
    rk = find("Val", "股權成本 ke")
    rate = {"type": "table", "caption": "DCF 折現率：WACC、股權成本、稅後債務成本、RONIC、永續成長率（直接顯示）", "head": ["項目", "數值", "說明"],
            "rows": [[VA(r), V(r, "p1" if r == rk + 2 else "p2"), f"@Val!E{r}"] for r in range(rk, rk + 6)]}
    rows = []
    r_end = find("Val", "DCF 目標價（情境加權）", rk)   # 第二輪修正：依標籤定位（模組 v2 stub 期在「＋淨現金與投資」後多兩列）
    for r in range(rk + 6, r_end):
        lab = str(Vw.cell(r, 1).value)
        f = ("n3" if lab.startswith("評價用") else "p1" if "占 EV" in lab else "x1" if "EV／EBITDA" in lab
             else pxf if ("每股" in lab or "推到" in lab) else amt)
        rows.append([VA(r)] + [f"@Val!{c}{r}|{f}" for c in "BCD"])
    rows.append([VA(r_end), V(r_end, pxf), "", ""])
    summ = {"type": "table", "caption": "DCF 彙總：明確期現值、終值現值、企業價值、淨現金與每股價值", "head": [f"@Val!{c}{rk - 1}" for c in "ABCD"], "rows": rows}
    hd = next(r for r in range(1, Dw.max_row + 1) if Dw.cell(r, 1).value == "項目")
    ncol = next(c for c in range(3, 30) if not Dw.cell(hd, c).value) - 3
    cols = [chr(ord("C") + i) for i in range(ncol)]
    det = []
    for r in range(hd + 1, Dw.max_row + 1):
        if str(Dw.cell(r, 1).value or "").startswith("情境 ") and Dw.cell(r + 1, 1).value == "營收" and len(det) < 3:
            rf = {str(r + k): {c: f for c in cols} for k, f in ((4, "p1"), (6, "p1"), (7, "p1"), (12, "n3"))}
            det.append({"type": "table", "caption": f"@DCF!A{r}", "collapse": True, "summary": "展開：本情境各年無槓桿自由現金流明細",
                        "head": [f"@DCF!A{hd}"] + [f"@DCF!{c}{hd}" for c in cols],
                        "rowgen": {"sheet": "DCF", "rows": f"{r + 1}-{r + 13}", "cols": ["A"] + cols, "fmts": [""] + [amt] * ncol, "row_fmt": rf}})
    return [rate, summ] + det


FCFF_MAIN = {"DCF 折現率": ["股權成本", "稅後債務成本", "權益占資本", "WACC", "永續成長率", "永續期 RONIC"],
             "DCF 彙總": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "＋Stub 期 FCFF", "淨現金滾到評價日", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"]}
FCFF_KEYS = {"rate": ["rf", "rf_us", "beta_u", "beta_de", "beta_du", "beta", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"],
             "bs": ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "bs_date", "stub_frac", "stub_div", "sh_v"],
             "fcf": ["fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]}


def sens_tp_block(pxf="n0"):
    t0 = find("Sens", "目標價參數敏感度")
    return {"type": "table", "caption": f"@Sens!A{t0}", "head": [f"@Sens!{c}{t0 + 1}" for c in "ABCDEFG"],
            "rowgen": {"sheet": "Sens", "rows": f"{t0 + 2}-{find('Sens', '檢查：閉式公式', t0) - 1}", "cols": list("ABCDEFG"), "fmts": ["", "", pxf, "sp1", pxf, "sp1", pxf]}}

v_tp, v_dcf, v_cmp, v_cmp0 = find("Val", "目標價（DCF 與 Comparable 加權）"), find("Val", "DCF 目標價（情境加權）"), find("Val", "Comparable 目標價（相對倍數調整後）"), find("Val", "Comparable 目標價（市場主流同業法）")
v_px, v_up, v_up0, v_cons = find("Val", "現價（"), find("Val", "上檔空間：加權目標價"), find("Val", "上檔空間：Comparable 目標價"), find("Val", "市場共識目標價")
v_rec, v_grade, v_band, v_hist, v_meth = find("Val", "投資建議"), find("Val", "上檔空間級距"), find("Val", "價格帶"), find("Val", "歷史倍數檢查"), find("Val", "目標價方法")
v_eps_f, v_eps_m = find("Val", "與市場共識目標價的差距：EPS"), find("Val", "與市場共識目標價的差距：倍數")
rd_upc = find("README", "上檔空間：Comparable 目標價（相對倍數調整後）", find("README", "HTML 輔助數字"))

# ---- 前處理一：估值章依規格重排（長條圖 → 結論主要項目 → 折現率 → DCF → Comparable → 其他）----
va = sec(cfg, "valuation")
old = va["blocks"]
dcf_tbl = next(b for b in old if b["type"] == "table" and ctext(b).startswith("0.1 DCF"))
cmp_tbl = next(b for b in old if b["type"] == "table" and ctext(b).startswith("0.2 Comparable"))
KV_CAP = [("本公司 10 年中位數", "相對倍數調整（Comparable 目標倍數＝同業中位數 × 本公司相對同業的長期倍數比）"),
          ("同業前瞻 P/E 中位數", "Comparable：同業中位數、前瞻 EPS 與目標價"), ("現價", "現價與市場隱含倍數")]


def kv_table(b):
    lab = str(wb["Val"][b["items"][0][0][5:]].value or "") if str(b["items"][0][0]).startswith("@Val!") else str(b["items"][0][0])
    capn = next((c for p, c in KV_CAP if lab.strip().startswith(p)), "其他項目")
    return {"type": "table", "caption": capn, "head": ["項目", "數值"], "rows": [[k, v] for k, v in b["items"]]}


drop = lambda b: ((b["type"] == "table" and (ctext(b).startswith("0.3 結論") or ctext(b) == "目標價（DCF 與 Comparable 加權）"))
                  or (b["type"] == "h3" and (b.get("text") == "重點觀察" or str(b.get("text", "")).startswith("相對倍數調整")))
                  or (b["type"] == "list" and any(str(i).startswith("@README!") for i in b.get("items", [])))
                  or (b["type"] == "note" and b.get("parts", [""])[0] in (VA(v_grade), VA(v_rec), VA(v_meth)))
                  or (b["type"] == "bars" and b.get("title") == "目標價與現價"))
kvs = [kv_table(b) for b in old if b["type"] == "kv"]
rm_tbl = [t for t in kvs if t["caption"].startswith("相對倍數調整")]
cmp_kv = [t for t in kvs if t["caption"].startswith("Comparable：")]
rest = []
for b in old:
    if b is dcf_tbl or b is cmp_tbl or drop(b):
        continue
    if b["type"] == "kv":
        t = kv_table(b)
        if t["caption"].startswith(("相對倍數調整", "Comparable：")):
            continue
        rest.append(t)
    else:
        rest.append(b)
va["blocks"] = [
    {"type": "bars", "title": "目標價與現價", "fmt": "n0", "items": [
        {"label": VA(v_dcf), "value": V(v_dcf), "series": 0},
        {"label": VA(v_cmp), "value": V(v_cmp), "series": 1},
        {"label": VA(v_tp), "value": V(v_tp), "series": 2},
        {"label": VA(v_cons), "value": V(v_cons), "series": 0}],
     "ref": {"label": VA(v_px), "value": V(v_px)}},
    {"type": "table", "caption": "目標價結論", "head": ["項目", "數值"], "rows": [
        [VA(v_tp), V(v_tp, "n0")], [VA(v_dcf), V(v_dcf, "n0")], [VA(v_cmp), V(v_cmp, "n0")], [VA(v_px), V(v_px, "n0")],
        [VA(v_up), V(v_up, "sp1")], [f"@README!A{rd_upc}", f"@README!B{rd_upc}|sp1"], [VA(v_rec), V(v_rec)], [VA(v_band), V(v_band)], [VA(v_cons), V(v_cons, "n0")],
        [VA(v_cmp0), V(v_cmp0, "n0")], [VA(v_up0), V(v_up0, "sp1")], [VA(v_grade), V(v_grade)], [VA(v_eps_f), V(v_eps_f, "sp1")], [VA(v_eps_m), V(v_eps_m, "sp1")],
        [VA(v_hist), [V(v_hist, "p0"), "／", f"@Val!C{v_hist}|p0", "／", f"@Val!D{v_hist}|p0"]], [VA(v_meth), V(v_meth)]]},
    *fcff_blocks(), cmp_tbl] + cmp_kv + rm_tbl + rest

# ---- 前處理二：敏感度章「目標價參數」表的列號依標題對到 Sens 頁實際位置；記下各龍捲風圖原標題 ----
se = sec(cfg, "sens")
tor_title = {}
for b in se["blocks"]:
    if b["type"] == "tornado":
        tor_title[b["value_col"]] = b["title"]
for i_, b in enumerate(se["blocks"]):
    if b["type"] == "table" and str(b.get("caption", "")).startswith("@Sens!A") and "rowgen" in b and ctext(b).startswith("目標價參數"):
        se["blocks"][i_] = sens_tp_block()   # v0.4：目標價參數表改為標準版 DCF 案例（tools/fcff_dcf.write_sens）

# ---- 前處理三：對帳章的表加標題（供指定主要列）----
for b in sec(cfg, "recon")["blocks"]:
    if b["type"] == "table" and not b.get("caption") and RECON_CAP:
        b["caption"] = RECON_CAP

# ---- 轉換：版面規格 v1.0 ----
scn_tok = re.split(r"[.\s]", str(wb["In"].cell(find("In", "情境驅動對照"), 1).value))[0]


SPEC = {
 "in_skip": [scn_tok],
 "in_groups": IN_GROUPS(INK),
 "blocks_to": dict(BLOCKS_TO),
 "main": {
  "目標價結論": ["目標價（DCF 與 Comparable 加權）", "DCF 目標價", "Comparable 目標價（相對倍數調整後）", "現價", "上檔空間：加權目標價", "上檔空間：Comparable 目標價（相對倍數調整後）", "投資建議", "價格帶", "市場共識目標價"],
  **FCFF_MAIN,
  "0.2 Comparable": [],
  "Comparable：": ["同業前瞻 P/E 中位數", "12 個月後的前瞻 EPS", "Comparable 目標價"],
  "相對倍數調整": ["本公司 10 年中位數", "同業 10 年中位數的中位數", "相對倍數調整係數", "Comparable 目標價（相對倍數調整後）"],
  "現價與市場隱含倍數": ["現價", "現價 ÷ FY2027 EPS", "現價 ÷ FY2028 EPS", "前瞻 P/E"],
  RECON_CAP: RECON_MAIN,
 },
}
if not SPEC["main"]["0.2 Comparable"]:
    SPEC["main"].pop("0.2 Comparable")
if sec(cfg, "drivers").get("intro"):   # 驅動因子章的說明文字（轉換器轉成 note）放第一塊
    SPEC["blocks_to"][sec(cfg, "drivers")["intro"]] = IN_GROUPS(INK)[0][0]
for b in sec(cfg, "drivers")["blocks"]:
    if b["type"] == "note" and str(b.get("parts", [""])[0]).startswith("@Units!A"):
        SPEC["blocks_to"][b["parts"][0]] = BLOCKS_TO["@Units!A"]
cfg = relayout(cfg, X, SPEC)

# ---- 後處理一：橫幅中間格＝只看 Comparable（市場主流同業法、未經相對倍數調整；工作單補充規則 A），附它自己的上檔空間；進加權的調整後值寫在橫幅附註 ----
cfg["hero"]["tiles"][1] = ["只看 Comparable 目標價（市場主流同業法）", V(v_cmp0, "n2"), ["上檔空間 ", V(v_up0, "sp1")]]
cfg["hero"]["note"] = list(cfg["hero"]["note"]) + ["；進加權的 Comparable 為相對倍數調整後 ", V(v_cmp, "n0"), "（調整係數 ", f"@Val!B{find('Val', '相對倍數調整係數')}|n2", "）。"]

# ---- 後處理二：龍捲風圖量的是 EPS 變動，不是加權目標價；標題改回原口徑 ----
for b in sec(cfg, "sens")["blocks"]:
    if b["type"] == "tornado":
        t = tor_title[b["value_col"]]
        b["title"] = t + ("：變動未達兩位數百分比的案例" if b.get("collapse") else "：變動達兩位數百分比的案例")
        if b.get("collapse"):
            b["summary"] = f"展開：{t.split('相對基準')[0].strip()} 變動未達兩位數百分比的案例"

# ---- 後處理三：估值章其他長條圖放回章末（轉換器會把所有長條圖移到章首）----
vb = sec(cfg, "valuation")["blocks"]
others = [b for b in vb if b["type"] == "bars" and b.get("title") != "目標價與現價"]
for b in others:
    vb.remove(b)
vb += others

# ---- 後處理四：表格分組小標的文字修正（Val 頁舊標籤與實際算法不符者）----
for s in cfg["sections"]:
    for b in s["blocks"]:
        for row in b.get("rows", []) if isinstance(b.get("rows"), list) else []:
            if isinstance(row, dict) and "sub" in row:
                m = re.match(r"^@([^!]+)!([A-Z]+\d+)$", str(row["sub"]))
                txt = str(wb[m.group(1)][m.group(2)].value or "").strip() if m else str(row["sub"])
                if txt in SUB_OVERRIDE:
                    row["sub"] = SUB_OVERRIDE[txt]

# ---- 後處理五：命題章補情境定義（原命題章的 kv，轉換器不保留）；附錄的版本要點改列 README 各版說明列 ----
th = sec(cfg, "thesis")
old_kv = next((b for b in sec(SRC, "thesis")["blocks"] if b["type"] == "kv"), None)
if old_kv:
    th["blocks"] += [{"type": "h3", "text": "情境定義與口徑"}, copy.deepcopy(old_kv)]
ws = wb["README"]
r_end = find("README", "展示層")
ver_rows = [[f"@README!A{r}"] for r in range(2, r_end) if re.match(r"^(v\d|版本：|\d{4}-\d{2}-\d{2}：)", str(ws.cell(r, 1).value or ""))]
st = sec(cfg, "status")
for i, b in enumerate(st["blocks"]):
    if b["type"] == "h3" and b.get("text") == "附錄：版本要點":
        st["blocks"][i + 1]["rows"] = ver_rows
        break
else:
    ins = next(i for i, b in enumerate(st["blocks"]) if b["type"] == "h3" and "未找到" in str(b.get("text")))
    st["blocks"][ins:ins] = [{"type": "h3", "text": "附錄：版本要點"},
                             {"type": "table", "collapse": True, "summary": "展開：各版的修改要點", "head": ["要點"], "rows": ver_rows}]

# 第三輪修正（2026-11-06）：情境驅動表依節標題取列並自檢、情境定義 kv 自檢、DCF 明細比例列格式、橫幅幣別（r3_html.py）
import r3_html  # noqa: E402
r3_html.post(cfg, wb, '新台幣')
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", OUT)
