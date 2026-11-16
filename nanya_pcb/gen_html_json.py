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
# v0.3.5：三層協調 C4（tools/coord/coord.py）：橫幅另一側讀數一格與揭露句、估值章開頭兩讀數表與長條圖
sys.path.insert(0, os.path.join(HERE, "..", "tools", "coord"))
import coord as COORD  # noqa: E402
COORD.inject_html(cfg, X)


# ================= v0.3.5（2026-11-15）：個股 HTML 版面規格 v1.3（樣板 NVIDIA v0.3.10、Ibiden v0.5.6、欣興 v0.4.4；docs/plan/個股HTML版面規格.md「v1.3 規格」） =================
# 章節：命題 → 估值 → 敏感度 → 反向 DCF → 損益（標準損益表）→ 驅動因子（分部營收推導）→ 對帳（含 C0）→ 資料狀態；無個別模組。
# 數字一律取自 xlsx（README 頁末 v1.3 白話列、PL 頁末標準損益表、RevDCF 頁、Val 0.7、Sens 頁末相對基準表、Coord 頁；nanya_pcb/v035_post.py 產生）。列號依標籤找。
def _v13(cfg):
    S = {x["id"]: x for x in cfg["sections"]}
    RD, PLW, RV, SW, CW = wb["README"], wb["PL"], wb["RevDCF"], wb["Sens"], wb["Coord"]
    fnd = lambda ws, p, start=1: next(r for r in range(start, ws.max_row + 1) if str(ws.cell(r, 1).value or "").strip().startswith(p))
    r13 = find("README", "v1.3 展示層")
    RM = dict(zip(["hero", "rec", "o1", "o2", "o3", "o4", "o5", "drv", "ver"], range(r13 + 1, r13 + 10)))
    assert str(RD[f"A{RM['ver']}"].value).startswith("v0.3.6"), RD[f"A{RM['ver']}"].value
    cfg["layout"] = "v1.3"
    fmt_u = lambda u: {"%": "p1", "元": "n2", "十億股": "n3", "x": "n3", "百萬顆": "n2"}.get(str(u).strip(), "n1")
    # ---- 橫幅：≤200 字白話＋對帳中揭露句；評等規則與相對倍數說明移到估值章 ----
    old_note = list(cfg["hero"]["note"])
    i_cd = len(old_note) - 2
    assert old_note[i_cd] == " " and str(old_note[i_cd + 1]).startswith("@Coord!B"), old_note
    cfg["hero"]["note"] = [f"@README!A{RM['hero']}"] + old_note[i_cd:]
    # ---- 1. 命題：主圖＝情境 B 分部營收（堆疊）與三情境營收 ----
    pa, pc = find("PL", "情境 A"), find("PL", "情境 C")
    fa = lambda base, lab_: next(r for r in range(base, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip() == lab_)
    rb = find("PL", "情境 B（基準）：")
    hdrB = rb + 1
    prow = lambda lab_: next(r for r in range(rb, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip().startswith(lab_))
    r_abf, r_bt, r_pcb = prow("ABF 載板營收 ＝"), prow("BT 載板"), prow("PCB（一般電路板）")
    r_rev = next(r for r in range(prow("C. 其他產品"), PLW.max_row + 1) if PLW.cell(r, 1).value == "營收")
    revA, revC = fa(pa, "營收"), fa(pc, "營收")
    th = S["thesis"]["blocks"]
    chart1 = {"type": "theme", "title": "營收分部（情境 B）與三情境營收（十億元）", "x": f"@PL!D{hdrB}:G{hdrB}", "fmt": "n0", "stack": True, "pad": 1.35,
              "total": f"@PL!D{r_rev}:G{r_rev}",
              "bars": [{"name": "ABF 載板", "y": f"@PL!D{r_abf}:G{r_abf}", "color": "#0b7a80"},
                       {"name": "BT 載板", "y": f"@PL!D{r_bt}:G{r_bt}", "color": "#6fc7cb"},
                       {"name": f"@PL!A{r_pcb}", "y": f"@PL!D{r_pcb}:G{r_pcb}", "color": "#8a8c8c"}],
              "lines": [{"name": [f"@PL!A{pa}", "營收"], "y": f"@PL!D{revA}:G{revA}", "color": "#0b7a80", "dash": True},
                        {"name": [f"@PL!A{pc}", "營收"], "y": f"@PL!D{revC}:G{revC}", "color": "#b4472f", "dash": True}],
              "caption": "堆疊＝情境 B（基準）的三個產品分部，頂端數字＝情境 B 合併營收；虛線＝情境 A、情境 C 的合併營收。FY2026 為前三季實績＋第四季假設的錨點；三情境差異來自加速器顆數與 BT 載板成長。"}
    i_obs = next(i for i, b in enumerate(th) if b["type"] == "h3" and b.get("text") == "重點觀察")
    th[i_obs + 1] = {"type": "list", "ordered": True, "items": [f"@README!A{RM[k]}" for k in ("o1", "o2", "o3", "o4", "o5")]}
    th.insert(0, chart1)
    th.insert(1, {"type": "note", "parts": [f"@README!A{RM['rec']}"]})
    for b in th:
        if b["type"] == "h3" and b.get("text") == "情境定義與口徑":
            b["text"] = "情境定義"
    # ---- 2. 估值：第一個區塊＝四個目標價與現價（G 位置）；規則與相對倍數說明放圖下 ----
    va = S["valuation"]["blocks"]
    g0 = find("Val", "0.7 估值並列"); gr = list(range(g0 + 2, g0 + 7)); r_pxg = g0 + 7
    four = {"type": "bars", "title": "四個目標價與現價（元；豎線＝現價）", "fmt": "n2",
            "items": [{"label": lb, "value": f"@Val!B{r}", "series": sr} for r, sr, lb in zip(gr, (1, 2, 2, 0, 1),
                      ("DCF 目標價", "Comparable（調整後，進加權）", "Comparable（未調整，不進加權）", "加權目標價（橫幅主值）", "市場共識目標價"))],
            "ref": {"label": f"@Val!A{r_pxg}", "value": f"@Val!B{r_pxg}"}}
    tbl4 = {"type": "table", "caption": f"@Val!A{g0}", "head": [f"@Val!{c}{g0 + 1}" for c in "ABCD"],
            "rows": [[f"@Val!A{r}", f"@Val!B{r}|n2", f"@Val!C{r}|sp1", f"@Val!D{r}"] for r in gr] + [[f"@Val!A{r_pxg}", f"@Val!B{r_pxg}|n2", "", f"@Val!D{r_pxg}"]]}
    keep = [b for b in va if not (b.get("type") == "bars" and b.get("title") == "目標價與現價")]
    assert any(b.get("type") == "h3" and b.get("text") == "對帳中：兩讀數目標價" for b in keep)
    head_note = {"type": "note", "parts": old_note[:i_cd]}
    S["valuation"]["blocks"] = [four, head_note, tbl4] + keep
    intro = S["valuation"].pop("intro", None)
    if intro:
        S["valuation"]["blocks"].insert(2, {"type": "note", "parts": [intro]})
    # ---- 3. 敏感度：主圖＝加權目標價相對基準（H 位置）----
    se = S["sens"]["blocks"]
    t0 = find("Sens", "加權目標價相對基準（HTML"); r1 = t0 + 2; r2 = r1
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
                        summary=f"展開：加權目標價相對基準：輸入取{side}值，變動未達兩位數百分比的案例（幾乎為零者見全部案例表）"))
    full = {"type": "table", "caption": f"@Sens!A{t0}", "collapse": True, "summary": "展開：全部案例的加權目標價相對基準（低案、高案）",
            "head": [f"@Sens!{c}{t0 + 1}" for c in "ABC"], "rows": [[f"@Sens!A{r}", f"@Sens!B{r}|sp1", f"@Sens!C{r}|sp1"] for r in range(r1, r2 + 1)]}
    snap = fnd(SW, "快照狀態", t0)
    for b in se:
        if b.get("type") == "tornado":
            b["summary"] = "展開：對照：" + str(b.get("title", "EPS 相對基準的變動")).split("：")[0] + ("（未達兩位數的案例）" if b.get("collapse") else "（達兩位數的案例；情境 A，一次動一個輸入）")
            b["collapse"] = True
        if b.get("type") == "table" and not b.get("collapse"):
            b["collapse"] = True
            b["summary"] = "展開：目標價參數表（WACC、β、ERP、永續成長率等）"
    intro_s = {"type": "note", "parts": ["主圖為加權目標價相對基準的變動（低案、高案分開，一次動一個輸入；變動達兩位數百分比者直接顯示，其餘摺疊）。業務主因為整本重算快照，快照狀態：",
                                          f"@Sens!B{snap}", "；目標價參數取 Sens 頁「目標價參數敏感度」表。下方摺疊為全部案例與 EPS 敏感度對照。"]}
    S["sens"].pop("intro", None)
    S["sens"]["blocks"] = tor[:1] + tor[2:3] + [intro_s, tor[1], tor[3], full] + se
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
                 "rows": [["預測期營業利益率（FY2028–FY2032 統一一個值；模型欄為等效值，終值年權重大）", f"@RevDCF!B{m1}|p1", f"@RevDCF!B{m0}|p1"],
                          ["現金流放大倍數（FY2028 起各年同乘；模型基準為一倍）", f"@RevDCF!B{s1}|n2", "—"],
                          ["FY2028 營收（十億元）", f"@RevDCF!B{rv1}|n0", f"@RevDCF!B{rv0}|n0"],
                          ["FY2028 營收年增率", f"@RevDCF!B{g1}|p1", f"@RevDCF!B{g0b}|p1"],
                          ["折現率（WACC）", f"@RevDCF!B{w1}|p2", f"@RevDCF!B{w0}|p2"]]},
                {"type": "table", "caption": f"@RevDCF!A{ra}", "head": ["項目", "值"], "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|{f}"] for r, f in zip(a_rows, afmt)]},
                {"type": "table", "caption": f"@RevDCF!A{rcmp}", "head": [f"@RevDCF!{c}{rcmp + 1}" for c in "ABCD"],
                 "rows": [[f"@RevDCF!A{r}", f"@RevDCF!B{r}|{f}", f"@RevDCF!C{r}|{f}", f"@RevDCF!D{r}|sp1"] for r, f in ((pe1, "n1"), (rc1, "n2"), (pe2, "n1"))]},
                {"type": "table", "caption": f"@RevDCF!A{yh}", "head": [f"@RevDCF!{c}{yh}" for c in "ABCDEFG"], "collapse": True,
                 "summary": "展開：情境 B 各年數字（反解的計算基礎；演算過程見 RevDCF 頁）",
                 "rows": [[f"@RevDCF!A{r}"] + [f"@RevDCF!{c}{r}|{f}" for c in "BCDEFG"] for r, f in zip(yrs_rows, yfmt)]}]}
    # ---- 5. 損益：標準損益表 ----
    p0 = find("PL", "標準損益表（v0.3.5")
    ph = next(r for r in range(p0, PLW.max_row + 1) if PLW.cell(r, 1).value == "項目")
    plr = lambda lab_, start=ph: next(r for r in range(start, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").strip() == lab_)
    keys = ["營收", "營收年增率（YoY）", "營業成本（COGS）", "毛利", "毛利率", "營業費用（含研發）", "營業利益", "營業利益率", "其他收入（淨額；營業外收支）", "稅前利益",
            "所得稅", "稅後淨利（歸屬母公司）", "淨利率", "加權股數（十億股）", "EPS（基本）"]
    PC = "CDEFGHIJ"
    rowx = lambda r: [f"@PL!A{r}"] + [f"@PL!{c}{r}|{fmt_u(PLW.cell(r, 2).value)}" for c in PC]
    r_ratio = plr("比率"); r_chk = next(r for r in range(r_ratio, PLW.max_row + 1) if str(PLW.cell(r, 1).value or "").startswith("核對（應為 0"))
    r_end = next(r for r in range(r_chk, PLW.max_row + 1) if not PLW.cell(r, 1).value) - 1
    r_dil = plr("（備註）EPS 完全稀釋（公司財報；只列實績）"); r_nci = plr("（備註）非控制權益損益")
    is_tbl = {"type": "table", "caption": "標準損益表（十億元；EPS 為元）：FY2023A–FY2025A 實績、FY2026E–FY2030E 情境 B（基準；GAAP＝台灣 IFRS）",
              "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in PC],
              "rows": [rowx(plr(k)) for k in keys] + [[f"@PL!A{r_dil}"] + [f"@PL!{c}{r_dil}|n2" for c in "CDE"] + [""] * 5]
                      + [[f"@PL!A{r_nci}"] + [f"@PL!{c}{r_nci}|n1" for c in "CDE"] + [""] * 5]
                      + [{"sub": f"@PL!A{r_ratio}"}] + [rowx(r) for r in range(r_ratio + 1, r_chk)]}
    chk_tbl = {"type": "table", "caption": f"@PL!A{r_chk}", "head": [f"@PL!A{ph}"] + [f"@PL!{c}{ph}" for c in PC], "collapse": True,
               "summary": "展開：核對列（歷史捨入差、預測期對 PL 情境 B、FY2030E 的推法差額）",
               "rows": [[f"@PL!A{r}"] + [f"@PL!{c}{r}|{'p1' if '占營收' in str(PLW.cell(r, 1).value) else 'n3'}" for c in PC] for r in range(r_chk + 1, r_end + 1)]}
    rv_, gm_, opm_, nm_ = plr("營收"), plr("毛利率"), plr("營業利益率"), plr("淨利率")
    is_chart = {"type": "theme", "title": "營收（十億元）與毛利率、營業利益率、淨利率：FY2023A–FY2030E（情境 B）", "x": f"@PL!C{ph}:J{ph}",
                "fmt": "n0", "fmt_r": "p0", "stack": True, "pad": 1.6,
                "bars": [{"name": f"@PL!A{rv_}", "y": f"@PL!C{rv_}:J{rv_}", "color": "#6fc7cb"}],
                "lines": [{"name": f"@PL!A{gm_}", "y": f"@PL!C{gm_}:J{gm_}", "axis": "r", "color": "#0b7a80", "labels": True},
                          {"name": f"@PL!A{opm_}", "y": f"@PL!C{opm_}:J{opm_}", "axis": "r", "color": "#c29a2b"},
                          {"name": f"@PL!A{nm_}", "y": f"@PL!C{nm_}:J{nm_}", "axis": "r", "color": "#6f6bae", "dash": True}],
                "caption": "長條＝營收；三條線＝毛利率、營業利益率、淨利率（右軸，GAAP）。FY2026E 為前三季財報實績＋第四季假設；FY2030E 為收斂法延伸期（毛利率與營業費用率沿用前一年）。"}
    old_pl = S["pl"]["blocks"]
    pls = [b for b in old_pl if b["type"] == "pl"]
    assert len(pls) == 2, [b["type"] for b in old_pl]
    old_blocks = [dict(pls[0], collapse=True, summary="展開：三情境損益摘要（FY2025A–FY2029）")]
    old_blocks += [dict(b, collapse=True, summary="展開：三情境損益補充列") for b in old_pl if b["type"] not in ("pl",)]
    old_blocks += [dict(pls[1], collapse=True, summary="展開：三情境延伸到 FY2032（收斂法）")]
    S["pl"].pop("intro", None)
    S["pl"].update(title="損益（標準損益表）",
                   blocks=[{"type": "note", "parts": [f"@PL!A{p0 + 1}"]}, is_chart, is_tbl, {"type": "note", "parts": [f"@PL!A{p0 + 2}"]}, {"type": "note", "parts": [f"@PL!A{r_end + 2}"]}, chk_tbl] + old_blocks)
    # ---- 6. 驅動因子：分部營收推導（情境 B）----
    seg = lambda rs: [[f"@PL!A{r}", f"@PL!B{r}"] + [f"@PL!{c}{r}|{fmt_u(PLW.cell(r, 2).value)}" for c in "DEFG"] for r in rs]
    r_acc = prow("南電加速器載板營收（新台幣）"); r_oai = prow("其他 AI 載板需求"); r_nai = prow("非 AI ABF 需求"); r_dem = prow("ABF 需求合計"); r_cap = prow("ABF 產能上限")
    a_rows6 = [prow(p_) for p_ in ("NVIDIA 加速器顆數", "非 NVIDIA 加速器顆數", "每顆載板面積與層數指數", "ABF 價格溢價指數", "南電漲價年度化", "每顆 NVIDIA 加速器載板金額", "南電份額：NVIDIA", "南電份額：非 NVIDIA")] + [r_acc]
    b_rows6 = [r_oai, r_nai, r_dem, r_cap, r_abf, prow("需求滿足率")]
    c_rows6 = [r_bt, r_pcb, r_rev, prow("營收年增")]
    d_rows6 = [prow("毛利率"), prow("營業利益"), prow("營業利益率"), prow("歸屬母公司淨利"), prow("EPS（基本）")]
    seg_tbl = {"type": "table", "caption": "分部營收推導（情境 B；十億元）：ABF＋BT 載板＋一般電路板＝合併營收", "head": [f"@PL!A{hdrB}", "單位"] + [f"@PL!{c}{hdrB}" for c in "DEFG"],
               "rows": [{"sub": "一、加速器載板＝顆數 × 每顆金額（面積層數 × 溢價）× 南電份額 × 漲價年度化"}] + seg(a_rows6)
                       + [{"sub": "二、ABF 載板＝MIN（加速器＋交換器等其他 AI＋非 AI 需求，產能上限）"}] + seg(b_rows6)
                       + [{"sub": "三、BT 載板、一般電路板與合併"}] + seg(c_rows6)
                       + [{"sub": "四、毛利與損益（擴產折舊）"}] + seg(d_rows6)}
    seg_chart = {"type": "theme", "title": "ABF 載板需求組成與產能上限（情境 B，十億元）", "x": f"@PL!D{hdrB}:G{hdrB}", "fmt": "n0", "stack": True, "pad": 1.25,
                 "total": f"@PL!D{r_dem}:G{r_dem}",
                 "bars": [{"name": "加速器載板", "y": f"@PL!D{r_acc}:G{r_acc}", "color": "#0b7a80"}, {"name": "交換器等其他 AI 載板", "y": f"@PL!D{r_oai}:G{r_oai}", "color": "#6fc7cb"},
                          {"name": "非 AI ABF", "y": f"@PL!D{r_nai}:G{r_nai}", "color": "#8a8c8c"}],
                 "lines": [{"name": "ABF 產能上限", "y": f"@PL!D{r_cap}:G{r_cap}", "color": "#b4472f", "dash": True, "labels": True}],
                 "caption": "堆疊＝ABF 載板需求的三個組成（加速器載板、交換器等其他 AI 載板、非 AI），頂端數字＝需求合計；虛線＝產能上限（FY2026 營收 × 產能指數 × 組合 × 溢價）。ABF 營收取兩者較小值；FY2026 為校準年。"}
    dr = S["drivers"]["blocks"]
    S["drivers"].update(title="驅動因子：分部營收推導")
    S["drivers"].pop("intro", None)
    S["drivers"]["blocks"] = [{"type": "note", "parts": [f"@README!A{RM['drv']}"]}, seg_chart, seg_tbl, {"type": "h3", "text": "情境驅動與單一參數"}] + dr
    # ---- 7. 對帳：C0 ----
    c1 = fnd(CW, "1. 讀哪一層"); c2 = fnd(CW, "2. C0"); c3 = fnd(CW, "3. 對帳中年份")
    h1, h2 = c1 + 1, next(r for r in range(c2, CW.max_row + 1) if CW.cell(r, 1).value == "項目")
    cnt = fnd(CW, "讀取層級列數")
    cunit = lambda r: str(CW.cell(r, 2).value or "")
    cf = lambda r: "n3"
    c0_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|{cf(r)}", f"@Coord!E{r}|{cf(r)}", f"@Coord!F{r}|sp1", f"@Coord!G{r}"]
               for r in range(h2 + 1, cnt) if CW.cell(r, 4).value is not None]
    c0_tbl = {"type": "table", "caption": "讀取層級核對：個股讀入的層級值 vs 層級值（同源，不判定、不計入對帳中；有檢驗力的對帳見上表「公司實績 vs 層級隱含值」）", "collapse": True, "summary": "展開：讀取層級核對（同源，不判定）", "head": [f"@Coord!{c}{h2}" for c in "ABCDEFG"], "rows": c0_rows}
    a4 = fnd(CW, "4. C0（實績）"); ih = next(r for r in range(a4, CW.max_row + 1) if CW.cell(r, 1).value == "項目"); mh = next(r for r in range(ih, CW.max_row + 1) if CW.cell(r, 1).value == "讀數")
    a4e = next(r for r in range(mh, CW.max_row + 1) if str(CW.cell(r, 1).value or "").startswith("對帳中項目數"))
    fm = lambda r: "p1" if str(CW.cell(r, 2).value) == "%" else "n3" if str(CW.cell(r, 2).value) == "US$B" else "n0" if str(CW.cell(r, 2).value) == "月" else "n1"
    act_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|n1", f"@Coord!E{r}|n1", (f"@Coord!F{r}|p1" if r == a4e - 2 else f"@Coord!F{r}|sp1"), f"@Coord!G{r}"]
                for r in range(mh + 1, a4e - 1)]
    in_rows = [[f"@Coord!A{r}", f"@Coord!B{r}", f"@Coord!C{r}", f"@Coord!D{r}|" + fm(r), f"@Coord!E{r}", f"@Coord!H{r}"] for r in range(ih + 1, mh - 1) if CW.cell(r, 1).value]
    act_blocks = [{"type": "h3", "text": "公司實績對層級隱含值（實績對帳）"}, {"type": "note", "parts": [f"@Coord!A{a4 + 1}"]}, {"type": "note", "parts": [f"@Coord!A{a4 + 2}"]},
                  {"type": "table", "caption": "合併營收：月營收與季報實績 vs Siliconomics 基準隱含值（主讀數、第二讀數；層級值由本模型校準，判定標參考）", "head": [f"@Coord!{c}{mh}" for c in "ABCDEFG"], "rows": act_rows},
                  {"type": "note", "parts": [f"@Coord!A{a4e - 1}"]},
                  {"type": "note", "parts": ["對帳中項目數：", f"@Coord!G{a4e}|n0", "。"]},
                  {"type": "table", "caption": "實績與層級隱含值的輸入、等級與出處（展開看）", "collapse": True, "summary": "展開：實績、層級值、換算參數與出處", "head": [f"@Coord!{c}{ih}" for c in "ABCDE"] + [f"@Coord!H{ih}"], "rows": in_rows}]
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
    c3_notes = [{"type": "note", "parts": [f"@Coord!A{c3 + 1}"]}]
    S["recon"]["blocks"] = [{"type": "h3", "text": "三層協調：由下往上對帳（公司實績對各層）"}] + act_blocks + c0_note + ref_blocks + c3_notes + [c0_src, rd_tbl] + [{"type": "h3", "text": "券商、共識與市場口徑"}] + S["recon"]["blocks"]
    # ---- 8. 資料狀態 ----
    st = S["status"]["blocks"]
    st[0:0] = [{"type": "h3", "text": "本版要點"}, {"type": "note", "parts": [f"@README!A{RM['ver']}"]}]
    order = ["thesis", "valuation", "sens", "revdcf", "pl", "drivers", "recon", "status"]
    S["revdcf"] = rev_
    cfg["sections"] = [S[k] for k in order]


_v13(cfg)
json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", OUT)

print("ok", OUT)
