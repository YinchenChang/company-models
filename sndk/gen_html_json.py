# 由 html_src.json（v0.4 以前的 HTML 設定，內容不動）套用「個股 HTML 版面規格 v1.0」（docs/plan/個股HTML版面規格.md），產生 sndk/html.json。
# 用法：python3 sndk/gen_html_json.py [成品 xlsx]（省略時取 sndk/dist/ 內唯一的 xlsx）；勿手改 html.json。
# 轉換器 tools/company_html/relayout.py 之外，本檔只做 SanDisk 專屬的前處理（估值章、對帳與敏感度表）與後處理（橫幅標籤、龍捲風標題、命題章情境定義）。
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
INK = {str(wb["In"].cell(r, 1).value): r for r in range(4, wb["In"].max_row + 1) if wb["In"].cell(r, 1).value}
v_reinv = find("Val", "再投資率")
v_cmp0, v_up0 = find("Val", "Comparable 目標價（市場主流同業法）"), find("Val", "上檔空間：Comparable 目標價")

# ---- 前處理一：驅動因子章的「單一參數」表其實是驅動 → 推導對照表（輸入、推導量、因果方向），移到附錄 ----
for b in sec("drivers")["blocks"]:
    if str(b.get("caption", "")).startswith("單一參數"):
        b["caption"] = "驅動 → 推導對照表（In 頁基準值與 PL 頁推導量）"

# ---- 前處理二：估值章（重點觀察與投資建議已在第一章與橫幅；kv 改為有標題的表；DCF 前加折現率表；同業表摺疊）----
va = sec("valuation")
nb, old, i = [], va["blocks"], 0
while i < len(old):
    b = old[i]
    nxt = old[i + 1] if i + 1 < len(old) else {}
    if b["type"] == "h3" and b.get("text") == "重點觀察" and nxt.get("type") == "list":
        i += 2
        continue
    if b["type"] == "note" and b.get("parts", [""])[0] == VA(find("Val", "投資建議")):
        i += 1
        continue
    if b["type"] == "h3" and str(b.get("text", "")).startswith("相對倍數調整") and nxt.get("type") == "kv":
        nb.append({"type": "table", "caption": b["text"], "head": ["項目", "數值"], "rows": [list(x) for x in nxt["items"]]})
        i += 2
        continue
    if b["type"] == "kv":
        lab0 = txt(b["items"][0][0])
        if lab0.startswith("同業前瞻 P/B 中位數"):
            nb.append({"type": "table", "caption": "Comparable 目標價的計算（市場主流同業法：同業前瞻 P/B 中位數 × 前瞻每股淨值）", "head": ["項目", "數值"], "rows": [list(x) for x in b["items"]]})
        elif lab0.startswith("同業前瞻 P/E 中位數"):
            nb.append({"type": "table", "collapse": True, "summary": "展開：同業前瞻 P/E 法的對照值（並列參考，不進目標價）", "head": ["項目", "數值"], "rows": [list(x) for x in b["items"]]})
        else:
            raise SystemExit(f"估值章有未處理的 kv：{lab0}")
        i += 1
        continue
    if b["type"] == "table" and txt(b.get("caption", "")).startswith("0.1 DCF"):
        nb.append({"type": "table", "caption": "DCF 折現率與永續假設（直接顯示）", "head": ["項目", "數值"],
                   "rows": [[f"@In!B{INK[k]}", f"@In!E{INK[k]}|p1"] for k in ("tp_r", "tp_g", "tp_roe")] + [[VA(v_reinv), V(v_reinv, "p1")]]})
    if b["type"] == "table" and txt(b.get("caption", "")).startswith("0.2 Comparable"):
        b = dict(b, collapse=True, summary="展開：記憶體同業的前瞻 P/B")
    nb.append(b)
    i += 1
va["blocks"] = nb

# ---- 前處理三：對帳表加標題（主要列用）；敏感度「目標價參數」表的列號改為 Sens 頁實際位置（舊設定的列號已位移）----
for b in sec("recon")["blocks"]:
    if b["type"] == "table" and not b.get("caption"):
        b["caption"] = "與公司實績、指引與共識的對照（共識只作對照，未用於校準）"
for b in sec("sens")["blocks"]:
    if b["type"] == "table" and str(b.get("caption", "")).startswith("@Sens!A"):
        t0 = find("Sens", "目標價參數")
        r1 = t0 + 2
        while wb["Sens"][f"A{r1 + 1}"].value and not str(wb["Sens"][f"A{r1 + 1}"].value).startswith("核對"):
            r1 += 1
        b["caption"] = f"@Sens!A{t0}"
        b["head"] = [f"@Sens!{c}{t0 + 1}" for c in "ABCDEF"]
        b["rowgen"]["rows"] = f"{t0 + 2}-{r1}"

# ---- 轉換：版面規格 v1.0 ----
scn_tok = re.split(r"[.\s]", str(wb["In"].cell(find("In", "情境驅動對照"), 1).value))[0]
B_NAND, B_COST, B_CO = "NAND 營收：市場層份額、價格與終端市場分部", "成本與營業費用", "全公司：股價、資產負債、股東回饋、折現率與目標價規則"
SPEC = {
 "in_skip": ["D", scn_tok],
 "in_groups": [
  [B_NAND, [
    ["校準：最近一季（FY2027 第一季）實績與市場層口徑", ["g_rev", "g_gm", "g_eps", "sd_k"]],
    ["終端市場分部占營收（FY2027–FY2029，假設）", ["dcsh_27", "cssh_27", "dcsh_28", "cssh_28", "dcsh_29", "cssh_29"]],
    ["季度實績（FY2026）", [f"q{q}_{k}" for q in range(1, 5) for k in ("rev", "gm", "op", "ni", "eps")]],
    ["終端市場分部實績（FY2026）", [f"q{q}_{k}" for q in range(1, 5) for k in ("dc", "ed", "cs")] + ["a26_dc", "a26_ed", "a26_cs"]],
    ["Wattonomics 錨定係數（只用於敏感度）", ["wta_27", "wta_28", "wta_29"]]]],
  [B_COST, [
    ["每位元成本年降（情境）", ["cd_A", "cd_B", "cd_C"]],
    ["營業費用", ["g_ox", "ox_var", "ox_g"]]]],
  [B_CO, [
    ["年度實績（FY2025、FY2026）", ["a26_rev", "a26_ni", "a26_eps", "a26_op", "a26_gm", "a25_rev", "a25_ni", "a25_eps"]],
    ["股價與股數", ["px", "bb_px", "sh0", "dil", "g_sh"]],
    ["資產負債、資本支出與股東回饋", ["eq0", "nc0", "da_y", "cap_r", "payout"]],
    ["稅與淨現金收益率", ["tax", "rf"]],
    ["折現率（DCF）", ["tp_r", "tp_g", "tp_roe"]],
    ["目標價規則與情境權重", ["tp_w_dcf", "tp_wA", "tp_wB", "tp_wC", "cons_tp"]],
    ["收斂期（FY2030 起回到循環中段）", ["nim10", "cyc_mA", "cyc_mC", "fade_mA", "fade_mB", "fade_mC", "fade_gA", "fade_gB", "fade_gC"]]]]],
 "blocks_to": {},
 "main": {
  "0.1 DCF": ["再投資率", "DCF 情境每股價值", "DCF 情境價值推到", "DCF 目標價"],
  "與公司實績、指引與共識": ["FQ1 FY27 營收 vs 公司實績", "FQ1 FY27 EPS vs 實績", "FY2027 EPS vs 共識", "FY2028 EPS vs 共識", "FY2028 年底每股淨值 vs 共識"],
 },
}
if sec("drivers").get("intro"):
    SPEC["blocks_to"][sec("drivers")["intro"]] = B_NAND
cfg = relayout(cfg, X, SPEC)

# ---- 後處理一：橫幅中間格＝只看 Comparable（市場主流同業法；工作單補充規則 A）。本公司相對倍數調整係數為 1（資料不足），兩者同值，只改標籤 ----
v_rm, v_adj = find("Val", "相對倍數調整係數"), find("Val", "Comparable 目標價（相對倍數調整後）")
assert abs(wb["Val"][f"B{v_adj}"].value - wb["Val"][f"B{v_cmp0}"].value) < 1e-9, "Comparable 有相對倍數調整，橫幅附註需加調整後值"
cfg["hero"]["tiles"][1] = ["只看 Comparable 目標價（市場主流同業法）", V(v_cmp0, "n2"), ["上檔空間 ", V(v_up0, "sp1")]]

# ---- 後處理二：龍捲風圖是「FY2028 EPS（情境 A）」的變動，不是加權目標價；標題改回正確口徑 ----
for b in sec("sens")["blocks"]:
    if b["type"] == "tornado":
        b["title"] = ("FY2028 EPS 相對基準的變動（情境 A；變動未達兩位數百分比的案例）" if b.get("collapse")
                      else "FY2028 EPS 相對基準的變動（情境 A；變動達兩位數百分比的案例；一次動一個輸入）")

# ---- 後處理三：驅動因子章的因果鏈說明（README 驅動因子各列）放在情境表之後；命題章補情境定義 ----
db = sec("drivers")["blocks"]
chain = next((b for b in db if b["type"] == "list" and any(str(x).startswith("@README!") for x in b.get("items", []))), None)
if chain:
    db.remove(chain)
    db.insert(1 if db and db[0]["type"] == "table" else 0, dict(chain))
intro = sec("drivers", SRC).get("intro")
nt = next((b for b in db if b["type"] == "note" and b.get("parts", [None])[0] == intro), None)
if nt:
    db.remove(nt)
    db.insert(0, nt)
old_kv = next(b for b in sec("thesis", SRC)["blocks"] if b["type"] == "kv")
sec("thesis")["blocks"] += [{"type": "h3", "text": "情境定義"}, copy.deepcopy(old_kv)]

# ---- 後處理四：附錄的版本要點（README 各版說明列；轉換器只找得到緊接的段落，本檔改為列出全部版本說明列）----
ws = wb["README"]
r_end = find("README", "展示層")
ver_rows = [[f"@README!A{r}"] for r in range(2, r_end) if re.match(r"^(v\d|\d{4}-\d{2}-\d{2}：)", str(ws.cell(r, 1).value or ""))]
st = sec("status")["blocks"]
st[:] = [b for b in st if not (b["type"] == "h3" and b.get("text") == "附錄：版本要點")]
st[:] = [b for b in st if not (b["type"] == "table" and str(b.get("summary", "")).startswith("展開：本版與前版"))]
ins = next(i for i, b in enumerate(st) if b["type"] == "h3" and "未找到" in str(b.get("text")))
st[ins:ins] = [{"type": "h3", "text": "附錄：版本要點"},
               {"type": "table", "collapse": True, "summary": "展開：各版的修改要點", "head": ["要點"], "rows": ver_rows}]

json.dump(cfg, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok", OUT)
