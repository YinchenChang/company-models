# v0.3.2（2026-11-05）：套用個股 HTML 版面規格 v1.0（docs/plan/個股HTML版面規格.md）——只改展示層。v0.4：DCF 改標準版（FCFF＋WACC），估值章 DCF 區塊改由本檔依 xlsx 列號產生；html_src.json 只重對 In 列號。
# 讀 accton/html_src.json（v0.3.1 的版面，內容不動）→ 少量前處理 → tools/company_html/relayout.py → 寫 accton/html.json。
# 用法：python3 accton/gen_html_json.py [成品 xlsx（預設 accton/dist/ 內唯一的 xlsx）]
import copy, glob, json, os, re, sys
import openpyxl
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools", "company_html"))
from relayout import relayout  # noqa: E402
X = sys.argv[1] if len(sys.argv) > 1 else (lambda f: f[0] if len(f) == 1 else sys.exit(f"dist 應恰有一個 xlsx：{f}"))(sorted(glob.glob(os.path.join(HERE, "dist", "*.xlsx"))))
cfg = json.load(open(os.path.join(HERE, "html_src.json"), encoding="utf-8"))
wb = openpyxl.load_workbook(X, data_only=True)
IN, VA = wb["In"], wb["Val"]
def row(ws, prefix, col=1):
    return next(r for r in range(1, ws.max_row + 1) if str(ws.cell(r, col).value or "").startswith(prefix))
ink = lambda k: next(r for r in range(1, IN.max_row + 1) if IN.cell(r, 1).value == k)
secs = {s["id"]: s for s in cfg["sections"]}
DC = wb["DCF"]
r_dcf = row(VA, "DCF 目標價（情境加權）")
_rk = row(VA, "股權成本 ke"); _rr = row(VA, "永續期 RONIC"); _rp = row(VA, "明確期 FCFF 現值")
_dh = row(DC, "項目")
# 第二輪修正：DCF 彙總表格式依標籤（模組 v2 stub 期在「＋淨現金與投資」後多兩列，不寫死列偏移）
_fm = lambda r: (lambda lab: "n3" if lab.startswith("評價用稀釋股數") else "p1" if "終值現值占 EV" in lab else "x1" if "隱含終值" in lab else "n2")(str(VA.cell(r, 1).value or "").strip())
DCF_BLOCKS = [
    {"type": "note", "parts": [f"@Val!A{_rk - 2}"]},
    {"type": "table", "caption": "DCF 折現率：股權成本、稅後債務成本、WACC、永續成長率、RONIC", "head": ["項目", "數值", "說明"],
     "rows": [[f"@Val!A{r}", f"@Val!B{r}|" + ("p1" if r == _rk + 2 else "p2"), f"@Val!E{r}"] for r in range(_rk, _rr + 1)]},
    {"type": "table", "caption": "DCF 彙總：明確期現值、終值現值、企業價值、淨現金與每股價值", "head": [f"@Val!{c}{_rk - 1}" for c in "ABCD"],
     "rows": [[f"@Val!A{r}"] + [f"@Val!{c}{r}|" + _fm(r) for c in "BCD"] for r in range(_rp, r_dcf)]
             + [[f"@Val!A{r_dcf}", f"@Val!B{r_dcf}|n2", "", ""]]}]
for s_ in "ABC":
    _a = row(DC, f"情境 {s_}（"); _b = next(r for r in range(_a, DC.max_row + 1) if str(DC.cell(r, 1).value or "").startswith("現值（FCFF"))
    _cols = ["A"] + [chr(ord("C") + i) for i in range(6)]
    _pct = {str(r): {c: "p1" for c in _cols[1:]} for r in range(_a + 1, _b + 1) if str(DC.cell(r, 1).value).startswith(("稅率", "折舊攤銷 ÷", "資本支出 ÷"))}
    _pct.update({str(r): {c: "n3" for c in _cols[1:]} for r in range(_a + 1, _b + 1) if str(DC.cell(r, 1).value).startswith("折現因子")})
    DCF_BLOCKS.append({"type": "table", "caption": f"@DCF!A{_a}", "collapse": True, "summary": "展開：本情境各年無槓桿自由現金流明細",
                       "head": [f"@DCF!{c}{_dh}" for c in _cols],
                       "rowgen": {"sheet": "DCF", "rows": f"{_a + 1}-{_b}", "cols": _cols, "fmts": [""] + ["n2"] * 6, "row_fmt": _pct}})

# 前處理（只動版面，不加數字）
# 0. 第二輪修正：「情境驅動」表依 xlsx 的「情境驅動對照」節重建列號（In 頁新增 beta_du、stub 列後，html_src.json 的寫死列號會錯位）
_s0 = row(IN, "情境驅動對照") + 1
_s1 = max(r for r in range(_s0 + 1, IN.max_row + 1) if str(IN.cell(r, 1).value or "").startswith("scn_"))
_nscn = 0
for b in secs["drivers"]["blocks"]:
    if b["type"] == "table" and str(b.get("caption", "")).startswith("情境驅動"):
        b["head"] = [f"@In!{c}{_s0}" for c in "BDEFC"]
        b["rows"] = [[f"@In!B{r}"] + [f"@In!{c}{r}" for c in "DEF"] + [f"@In!C{r}"] for r in range(_s0 + 1, _s1 + 1)]
        _nscn += 1
assert _nscn == 1, _nscn
# 1. 驅動：原「單一參數」表實為輸入 → 推導量對照，交給轉換器放到附錄
for b in secs["drivers"]["blocks"]:
    if str(b.get("caption", "")).startswith("單一參數"):
        b["caption"] = "驅動 → 推導量對照"
# 2. 估值：重點觀察已移到第一章；相對倍數調整由 kv 改為表格（轉換器會把 kv 當成折現率表）；
#    DCF 折現率直接顯示、緊接在 DCF 彙總之前；無標題的 DCF、Comparable 計算表補標題以便指定主要列
va = secs["valuation"]["blocks"]
out, skip = [], False
for b in va:
    if b["type"] == "h3" and b.get("text") == "重點觀察":
        skip = True
        continue
    if skip and b["type"] == "list":
        skip = False
        continue
    if b["type"] == "kv":
        b = {"type": "table", "caption": "相對倍數調整", "head": ["項目", "數值"], "rows": [list(i) for i in b["items"]]}
    if b["type"] == "table" and b.get("caption") == "@Val!A4":      # 標準版 DCF：舊 0.1 DCF（股東現金流）兩張表 → DCF 折現率、彙總與各情境 FCFF 明細
        out += DCF_BLOCKS
        continue
    if b["type"] == "table" and not b.get("caption") and b.get("rows") and b["rows"][0][0] == "@Val!A13":
        continue
    if b["type"] == "bars":
        b = copy.deepcopy(b)
        for it in b["items"]:
            if it["label"] == "@Val!A23":
                it["label"], it["value"] = f"@Val!A{r_dcf}", f"@Val!B{r_dcf}"
    if b["type"] == "table" and not b.get("caption") and b.get("rows") and b["rows"][0][0] == f"@Val!A{row(VA, '同業前瞻 P/E 中位數')}":
        b = dict(b, caption="Comparable 計算")
    out.append(b)
secs["valuation"]["blocks"] = out
# 標準版 DCF：Sens 目標價參數表改為十個案例（tools/fcff_dcf.py write_sens；起始列不變）
SE = wb["Sens"]
_st = row(SE, "目標價參數敏感度")
_se = next(r for r in range(_st + 2, SE.max_row + 1) if str(SE.cell(r, 1).value or "").startswith("檢查：閉式公式"))
for b in secs["sens"]["blocks"]:
    if b["type"] == "table" and b.get("caption") == f"@Sens!A{_st}":
        b["head"] = [f"@Sens!{c}{_st + 1}" for c in "ABCDEFG"]
        b["rowgen"] = dict(b["rowgen"], rows=f"{_st + 2}-{_se - 1}")
# 3. 對帳：單一大表補標題以便指定主要列
for b in secs["recon"]["blocks"]:
    if b["type"] == "table" and not b.get("caption"):
        b["caption"] = "對帳要點"
tor_title = {b["title"]: b for b in secs["sens"]["blocks"] if b["type"] == "tornado"}

scen_tok = re.split(r"[.\s]", IN.cell(row(IN, "情境驅動對照"), 1).value)[0]
SPEC = {
 "in_skip": [scen_tok],
 "in_groups": [
  ["白牌 AI 乙太網路交換器", [
    ["AI 交換器市場（讀網通市場模型）", ["mkt_26", "mkt_27", "mkt_28", "mkt_29", "su_26", "su_27", "su_28", "su_29"]],
    ["乙太網路占比", ["eth_26", "eth_27", "eth_28", "eth_29"]],
    ["白牌占比", ["wb_26", "wb_27", "wb_28", "wb_29"]],
    ["智邦白牌份額相對值", ["rel_27", "rel_28", "rel_29"]],
    ["市場量敏感度係數（只進敏感度）", ["mf"]]]],
  ["AI 加速卡等網路應用產品（Trainium）", [
    ["Trainium 顆數（讀網通市場模型）", ["trn_26", "trn_27", "trn_28", "trn_29"]],
    ["每顆 Trainium 的智邦營收相對值", ["nrel_27", "nrel_28", "nrel_29"]],
    ["網路應用產品毛利率", ["d_nap"]],
    ["顆數敏感度係數（只進敏感度）", ["tf"]]]],
  ["非 AI 與營收結構", [
    ["營收結構（FY2025 公司揭露、FY2026 拆分）", ["sw25s", "nap25s", "af25", "ns26"]],
    ["非 AI 營收成長", ["g_nai26", "g_nai27", "g_nai28", "g_nai29"]]]],
  ["全公司營運：FY2026 錨點、情境與費用", [
    ["FY2026 錨點（第三季營收實績與假設）", ["q3_rev", "q3_gm", "q4_g", "q4_gm", "q_opx", "q_nio", "q_sh"]],
    ["情境顆數係數（兩條 AI 線共用）", ["um28_A", "um29_A", "um28_B", "um29_B", "um28_C", "um29_C"]],
    ["Wattonomics 錨定係數（只進敏感度）", ["wta_27", "wta_28", "wta_29"]],
    ["AI IT 資本支出（對帳參考）", ["itc_26", "itc_27", "itc_28", "itc_29"]],
    ["營業費用、業外與稅、股數（FY2027 起）", ["phi", "opx_adj", "nio_f", "sh_g27", "sh_g"]],
    ["匯率", ["fx"]]]],
  ["全公司：股價、折現率與目標價規則", [
    ["股價與市場倍數", ["px", "pe_ttm"]],
    ["共識與法人資料（只作對帳）", ["cons_rev26", "cons_eps26", "cons_rev27", "cons_eps27", "cons_eps28", "br_sw26", "dell_top2", "cons_tp"]],
    ["目標價規則（情境權重、DCF 權重、PL 延伸期收斂）", ["tp_wA", "tp_wB", "tp_wC", "tp_w_dcf", "tp_g", "tp_gd", "nm10_k"]],
    ["折現率（標準版 DCF：股權成本、WACC、永續成長率、RONIC）", ["rf", "rf_us", "beta_u", "beta_u_med", "beta_du", "beta_de", "beta", "erp", "crp", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt"]],
    ["資本結構、淨現金與股數（評價用）", ["debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "bs_date", "stub_frac", "stub_div", "sh_v"]],
    ["自由現金流（折舊、資本支出、營運資金）", ["fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc"]]]]],
 "blocks_to": {"智邦白牌交換器份額": "白牌 AI 乙太網路交換器", "@Units!A15": "白牌 AI 乙太網路交換器"},
 "main": {"DCF 彙總": ["明確期 FCFF 現值", "終值現值", "企業價值 EV", "＋淨現金與投資", "淨現金滾到評價日", "DCF 情境價值推到 12 個月後", "DCF 目標價（情境加權）"],
          "DCF 折現率": ["股權成本", "稅後債務成本", "權益占資本", "WACC", "永續成長率", "永續期 RONIC"],
          "對帳要點": ["FY2026 營收 vs 共識", "FY2026 EPS vs 共識", "FY2027 EPS vs 共識", "FY2028 EPS vs 共識", "3Q26 EPS"]},
}
# 第三輪修正 15／21：情境驅動表依 In 頁 C 欄單位格式化（% 與 pt 用百分比，避免 0.06 顯示成小數）；DCF 彙總主表直接顯示 stub 兩列
def _uf3(r_):
    u = IN.cell(r_, 3).value
    return "p1" if u in ("%", "pt") else "n3" if u == "十億股" else "n2"
for b in secs["drivers"]["blocks"]:
    if b["type"] == "table" and str(b.get("caption", "")).startswith("情境驅動"):
        b["rows"] = [[x[0]] + [(f"{c}|{_uf3(int(c.split('!')[1][1:]))}" if "|" not in c and isinstance(IN[c.split('!')[1]].value, (int, float)) else c) for c in x[1:4]] + x[4:] for x in b["rows"]]
_m3 = SPEC["main"]["DCF 彙總"]
if "＋Stub 期 FCFF" not in _m3:
    _m3.insert(_m3.index("＋淨現金與投資") + 1, "＋Stub 期 FCFF")
cfg = relayout(cfg, X, SPEC)
# 後處理（補充規則 A）：橫幅中間格＝只看 Comparable（市場主流同業法、未經相對倍數調整）；附註列進加權的調整後值
_rc, _ru, _ra = row(VA, "Comparable 目標價"), row(VA, "上檔空間：Comparable"), row(VA, "Comparable 目標價（相對倍數調整後）")
assert VA.cell(_rc, 1).value == "Comparable 目標價" and str(VA.cell(_ra - 1, 1).value).startswith("相對倍數調整係數")
cfg["hero"]["tiles"][1] = ["只看 Comparable 目標價（市場主流同業法）", f"@Val!B{_rc}|n2", ["上檔空間 ", f"@Val!B{_ru}|sp1"]]
cfg["hero"]["note"] = list(cfg["hero"].get("note") or []) + ["。進加權的 Comparable 為相對倍數調整後 ", f"@Val!B{_ra}|n2", "（調整係數 ", f"@Val!B{_ra - 1}|n2", "）"]
# 後處理：轉換器把估值章所有長條圖移到最前面；舊法（規則值）長條圖放回舊法段落，避免與目標價並列
_vb = cfg["sections"][[s["id"] for s in cfg["sections"]].index("valuation")]["blocks"]
_old = next(b for b in _vb if b["type"] == "bars" and b.get("title") == "規則值隱含股價與現價")
_vb.remove(_old)
_h = next(i for i, b in enumerate(_vb) if b["type"] == "h3" and str(b.get("text", "")).startswith("規則值隱含股價"))
_vb.insert(_h + 2, _old)
# 後處理：單位為 pt（百分點）的參數，轉換器預設一位小數；改用百分比格式，避免 0.06 顯示成 0.1
_pt = {r for r in range(4, IN.max_row + 1) if IN.cell(r, 3).value == "pt"}
def _fix_pt(x):
    m = re.match(r"^@In!([DEF])(\d+)\|n1$", x) if isinstance(x, str) else None
    return f"@In!{m.group(1)}{m.group(2)}|p1" if m and int(m.group(2)) in _pt else x
for b in cfg["sections"][[s["id"] for s in cfg["sections"]].index("drivers")]["blocks"]:
    if b["type"] == "table" and isinstance(b.get("rows"), list):
        b["rows"] = [[_fix_pt(c) for c in r] if isinstance(r, list) else r for r in b["rows"]]
# 後處理：本模型的龍捲風圖是 FY2028 EPS（不是加權目標價），轉換器的通用標題改回原意
for b in cfg["sections"][[s["id"] for s in cfg["sections"]].index("sens")]["blocks"]:
    if b["type"] == "tornado":
        b["title"] = "FY2028 EPS 相對基準的變動（情境 A，一次動一個輸入；" + ("未達兩位數百分比）" if b.get("collapse") else "變動達兩位數百分比的案例）")
# 第三輪修正 17：橫幅金額標明幣別
_CUR3 = '新台幣'
def _lab3(x):
    t = str(wb[x[1:].split("!")[0]][x.split("!")[1]].value) if isinstance(x, str) and x.startswith("@") else str(x)
    return x if _CUR3 in t else ([x, "（" + _CUR3 + "）"] if isinstance(x, str) and x.startswith("@") else t + "（" + _CUR3 + "）")
for _t in cfg["hero"]["tiles"]:
    _t[0] = _lab3(_t[0])
json.dump(cfg, open(os.path.join(HERE, "html.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok")
