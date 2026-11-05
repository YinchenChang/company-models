# 由 html_src.json（v0.3 以前的 HTML 設定，內容不動）套用「個股 HTML 版面規格 v1.0」（docs/plan/個股HTML版面規格.md），產生 unimicron/html.json。
# 用法：python3 unimicron/gen_html_json.py [成品 xlsx]（省略時取 unimicron/dist/ 內唯一的 xlsx）；勿手改 html.json。
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
B_ABF = "ABF 載板：其他 AI、非 AI 與產能"
B_OTH = "BT 載板、HDI／PCB 與軟板"
B_CO = "全公司：股價、折現率、目標價規則、稅與現金流"
IN_GROUPS = lambda K: [
  [B_AI, [
    ["加速器顆數與 AI IT 資本支出（讀網通市場模型）", [k for k in K if re.match(r"^(u_|g29_|itc_)", k)]],
    ["情境顆數係數（個股自訂）", ["um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C"]],
    ["每顆載板金額與 ABF 溢價（讀 Ibiden 模型）", ["v26", "g_asp", "ibpp_27", "ibpp_28", "ibpp_29", "ibrel_27", "ibrel_28", "ibrel_29", "ib_sh26", "ib_w_ot", "v_ot"]],
    ["欣興份額", ["sh_nv_26", "sh_nv_27", "sh_nv_28", "sh_nv_29", "sh_ot_26", "sh_ot_27", "sh_ot_28", "sh_ot_29"]],
    ["Wattonomics 錨定係數（只用於敏感度）", ["wta_27", "wta_28", "wta_29"]]]],
  [B_ABF, [
    ["FY2026 起點：ABF 占營收與 AI 占比", ["mx_abf", "ai_abf"]],
    ["其他 AI 與非 AI ABF 需求", ["el_o", "g_nai"]],
    ["產能、組合與漲價", ["cap_27", "cap_28", "cap_29", "g_mix", "carry", "tgt_abf_gap27"]],
    ["ABF 增量現金毛利率", ["ft_abf"]]]],
  [B_OTH, [
    ["FY2026 起點：產品組合", ["mx_bt", "mx_pcb", "ai_pcb"]],
    ["營收成長", ["g_bt", "g_pai_27", "g_pai_28", "g_pai_29", "g_pcbn", "g_fpc"]],
    ["增量現金毛利率", ["ft_oth"]]]],
  [B_CO, [
    ["資本支出與折舊", ["cx27_A", "cx28_A", "cx29_A", "cx27_B", "cx28_B", "cx29_B", "cx27_C", "cx28_C", "cx29_C", "capex26", "da26", "life"]],
    ["營業費用、業外、稅與股數", ["phi", "nonop27", "nonop28", "nonop29", "tax", "sh_g"]],
    ["FY2026 錨點（季報實績與第四季假設）", ["q3_rev", "q3_gm", "q3_opx", "q3_ni", "q4_g", "q4_gm", "q4_opx", "q_nonop", "q_sh"]],
    ["共識與券商（只作對照）", ["cons_eps26", "cons_eps27", "cons_eps28", "cons_rev27", "cons_rev28", "bk_eps27", "cons_feps", "cons_tp"]],
    ["股價、市場倍數與匯率", ["px", "fpe", "pe_ttm", "fx", "nim5"]],
    ["折現率（DCF）", ["tp_r", "tp_g", "tp_roe"]],
    ["目標價規則與收斂法", ["tp_w_dcf", "tp_wA", "tp_wB", "tp_wC", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]]]]]
BLOCKS_TO = {"單一參數": B_AI, "AI 加速器載板份額": B_AI, "@Units!A": B_AI}
RECON_CAP = "本模型對照券商、共識與市場口徑"
RECON_MAIN = ["FY2027 EPS vs 2026-10 券商平均", "FY2028 EPS vs FactSet", "FY2027 ABF 需求滿足率", "FY2026 AI 占營收"]
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
v_tp, v_dcf, v_cmp, v_cmp0 = find("Val", "目標價（DCF 與 Comparable 加權）"), find("Val", "DCF 目標價（情境加權）"), find("Val", "Comparable 目標價（相對倍數調整後）"), find("Val", "Comparable 目標價（市場主流同業法）")
v_px, v_up, v_up0, v_cons = find("Val", "現價（"), find("Val", "上檔空間：加權目標價"), find("Val", "上檔空間：Comparable 目標價"), find("Val", "市場共識目標價")
v_rec, v_grade, v_band, v_hist, v_meth = find("Val", "投資建議"), find("Val", "上檔空間級距"), find("Val", "價格帶"), find("Val", "歷史倍數檢查"), find("Val", "目標價方法")
v_eps_f, v_eps_m = find("Val", "與市場共識目標價的差距：EPS"), find("Val", "與市場共識目標價的差距：倍數")
v_reinv = find("Val", "參考：再投資率", must=False) or find("Val", "再投資率", must=False)
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
disc_rows = [[f"@In!B{INK[k]}", f"@In!E{INK[k]}|p1"] for k in ("tp_r", "tp_g", "tp_roe") if k in INK] + ([[f"@Val!A{v_reinv}", V(v_reinv, "p1")]] if v_reinv else [])
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
    {"type": "table", "caption": "DCF 折現率（股東權益現金流折現；直接顯示）", "head": ["項目", "數值"], "rows": disc_rows},
    dcf_tbl, cmp_tbl] + cmp_kv + rm_tbl + rest

# ---- 前處理二：敏感度章「目標價參數」表的列號依標題對到 Sens 頁實際位置；記下各龍捲風圖原標題 ----
se = sec(cfg, "sens")
tor_title = {}
for b in se["blocks"]:
    if b["type"] == "tornado":
        tor_title[b["value_col"]] = b["title"]
    if b["type"] == "table" and str(b.get("caption", "")).startswith("@Sens!A") and "rowgen" in b:
        t0 = find("Sens", "目標價參數")
        if ctext(b).startswith("目標價參數"):
            r1 = t0 + 2
            while wb["Sens"][f"A{r1 + 1}"].value:
                r1 += 1
            b["caption"] = f"@Sens!A{t0}"
            b["head"] = [f"@Sens!{c}{t0 + 1}" for c in b["rowgen"]["cols"]]
            b["rowgen"]["rows"] = f"{t0 + 2}-{r1}"

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
  "0.1 DCF": [p for p in ("DCF 情境每股價值", "DCF 情境價值推到", "DCF 目標價（情境加權）", "參考：EPS", "再投資率")
              if any(str(wb["Val"].cell(r, 1).value or "").strip().startswith(p) for r in range(find("Val", "0.1 DCF"), find("Val", "0.2 Comparable")))][:5],
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

# ---- 後處理一：橫幅的 Comparable 目標價改用「相對倍數調整後」（加權目標價實際使用的值；轉換器預設抓到未調整的列）----
cfg["hero"]["tiles"][1] = [VA(v_cmp), V(v_cmp, "n2"), ["上檔空間 ", f"@README!B{rd_upc}|sp1"]]

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

json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", OUT)
