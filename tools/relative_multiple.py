# 相對倍數調整（Comparable 的目標倍數＝同業前瞻倍數中位數 × 本公司 10 年相對同業的倍數比）
# 依據：Andy 2026-10-27 15:20「照建議做」（AI半導體 目標價總表 v0.2.3 評等偏向分析的建議 1）。
# 狀態（2026-10-27）：**尚未套用到任何模型**。以年底落後 P/E 試算，谷底年度雜訊與口徑不一致造成 15 檔評等翻轉、多為偽訊號
#   （見目標價總表 v0.2.4「相對倍數試算」頁）；待 Bloomberg 前瞻倍數十年月資料取得後，換 DATA 重算、確認合理再接入各建檔程式。
# 問題：Comparable 直接套同業中位數，等於假設本公司長期的折價／溢價（韓國折價、A 股地緣折價、
#       代工組裝低毛利、獨占溢價等）會在 12 個月內消失；評等因此主要反映「現價倍數高低」。
# 做法：
#   本公司 10 年中位數 ÷ 同業各自 10 年中位數的中位數 ＝ 原始比值；夾在 [RM_LO, RM_HI]；
#   資料不足（本公司有效年數 < MIN_YEARS，或有效同業 < MIN_PEERS）時係數＝1（不調整）。
#   Comparable 目標價（相對倍數調整後）＝ Comparable 目標價（市場主流同業法）× 係數；加權目標價改用調整後的值。
#   「Comparable 目標價（市場主流同業法）」與「上檔空間：Comparable 目標價」維持不調整（Andy 2026-10-25：看與市場主流的差距）。
# 資料：tools/data/rel_mult_hist.csv（年底落後 P/E 或 P/B，2016–2025；來源與網址逐列記載）。
#   這是 [Derived] 代理：理想來源是 Bloomberg 前瞻 P/E 10 年月資料（抓取單已在 Drive 選股與資料/），取得後只換本檔資料即可。
# 用法（各公司建檔程式，在 Val 頁所有內容寫完之後、存檔之前）：
#   import relative_multiple as RM
#   RM.add_block(wb, V, own="2382.TW", peers=[...], metric="PE", comp_ref=f"Val!$B${T['comptp']}", pxfmt="#,##0")
#   然後把加權目標價公式中的 Comparable 改成名稱 tp_comp_adj。
import csv
import os
import statistics
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data", "rel_mult_hist.csv")
YEARS = list(range(2016, 2026))
RM_LO, RM_HI = 0.5, 2.0          # [Assumed] 夾值上下限；Andy 可改（改這裡即全部模型一致）
MIN_YEARS, MIN_PEERS = 7, 2       # [Assumed] 資料門檻（7 年：排除只有 2021–2025 的序列，避免 AI 高峰期偏誤）
VALID = {"PE": (0.0, 200.0), "PB": (0.0, 20.0)}   # 有效值區間（不含下限）；沿用十年 P/E 規則「負值或 >200 剔除」

BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC")


def norm(t):
    t = str(t).strip()
    if "／" in t:
        t = t.split("／")[1]
    return t


def load(path=DATA):
    """回傳 {(ticker, metric): {year: (value, url)}}；value 為 float。"""
    out = {}
    with open(path, encoding="utf-8") as f:
        for r in csv.DictReader(f):
            if not r["year"] or r["value"] in ("", None):
                out.setdefault((r["ticker"], r["metric"]), {})
                continue
            out.setdefault((r["ticker"], r["metric"]), {})[int(r["year"])] = (float(r["value"]), r["source_url"])
    return out


def _valid(v, metric):
    lo, hi = VALID[metric]
    return lo < v <= hi


def summary(own, peers, metric="PE", hist=None):
    """Python 端同口徑計算（供總表、稽核與測試用；Excel 端以公式重算，兩者須一致）。"""
    hist = hist if hist is not None else load()

    def med(t):
        vals = [v for y, (v, _) in hist.get((norm(t), metric), {}).items() if y in YEARS and _valid(v, metric)]
        return (statistics.median(vals) if len(vals) >= MIN_YEARS else None), len(vals)
    om, on = med(own)
    pm = []
    for _, t in peers:
        m, _n = med(t)
        if m is not None:
            pm.append(m)
    pmed = statistics.median(pm) if pm else None
    ok = om is not None and len(pm) >= MIN_PEERS
    raw = om / pmed if ok else None
    fac = min(max(raw, RM_LO), RM_HI) if ok else 1.0
    return {"own_med": om, "own_n": on, "peer_meds": pm, "peer_med": pmed, "n_peers": len(pm), "raw": raw, "factor": fac, "ok": ok}


def add_block(wb, ws, own, peers, metric="PE", comp_ref=None, pxfmt="#,##0", own_name="本公司", r0=None, hist=None):
    """在 ws（Val 頁）最後寫入「相對倍數調整」區塊；定義名稱 rm_factor、tp_comp_adj。
    peers：[(名稱, 代號)]，與 Comparable 同業表相同的公司（順序不限）。
    comp_ref：Comparable 目標價（市場主流同業法）所在儲存格，例如 "Val!$B$130"。回傳 dict（各列列號）。"""
    hist = hist if hist is not None else load()
    if r0 is None:
        r0 = ws.max_row + 3
    lab = "P/E" if metric == "PE" else "P/B"
    ws.cell(r0, 1, f"0.5 相對倍數調整（Comparable 目標倍數＝同業中位數 × 本公司 10 年相對同業的 {lab} 比；Andy 2026-10-27 新增，放在本頁最後以免位移既有儲存格）").font = BOLD
    ws.cell(r0 + 1, 1, f"資料：2016–2025 年底落後 {lab}（tools/data/rel_mult_hist.csv，來源見 F 欄起各列網址），[Derived] 代理；"
                       f"有效值 {VALID[metric][0]:g} < 值 ≤ {VALID[metric][1]:g}，無效值以「剔除:」文字列示、不進中位數。"
                       f"本公司或同業有效年數 < {MIN_YEARS} 不納入；有效同業 < {MIN_PEERS} 家或本公司不足時係數＝1（不調整）。")
    h = r0 + 2
    heads = ["公司", "代號", "角色"] + [str(y) for y in YEARS] + ["有效年數", "10 年中位數", "納入值", "來源"]
    for j, x in enumerate(heads, 1):
        c = ws.cell(h, j, x)
        c.font = BOLD
        c.fill = HDR
    CY = {y: 4 + i for i, y in enumerate(YEARS)}          # D..M
    CN, CMED, CIN, CSRC = 14, 15, 16, 17                  # N O P Q
    from openpyxl.utils import get_column_letter as L
    rows = [(own_name, own, "本公司")] + [(n, t, "同業") for n, t in peers]
    r = h + 1
    first = r
    for nm, tk, role in rows:
        tk = norm(tk)
        ws.cell(r, 1, nm)
        ws.cell(r, 2, tk)
        ws.cell(r, 3, role)
        d = hist.get((tk, metric), {})
        urls = []
        for y in YEARS:
            if y in d:
                v, u = d[y]
                if u and u not in urls:
                    urls.append(u)
                if _valid(v, metric):
                    c = ws.cell(r, CY[y], v)
                    c.font = BLUE
                    c.number_format = "0.00"
                else:
                    ws.cell(r, CY[y], f"剔除:{v:g}")
        rng = f"{L(CY[YEARS[0]])}{r}:{L(CY[YEARS[-1]])}{r}"
        ws.cell(r, CN, f"=COUNT({rng})")
        ws.cell(r, CMED, f'=IF({L(CN)}{r}>0,MEDIAN({rng}),"")').number_format = "0.00"
        ws.cell(r, CIN, f'=IF({L(CN)}{r}>={MIN_YEARS},{L(CMED)}{r},"")').number_format = "0.00"
        ws.cell(r, CSRC, "；".join(urls) if urls else ("未找到" if not d else "未找到有效值"))
        r += 1
    last = r - 1
    pin = f"{L(CIN)}{first + 1}:{L(CIN)}{last}" if last > first else None
    r += 1
    R = {}

    def row(key, label, formula, fmt, note=""):
        nonlocal r
        ws.cell(r, 1, label)
        c = ws.cell(r, 2, formula)
        c.number_format = fmt
        if note:
            ws.cell(r, 5, note)
        R[key] = r
        r += 1
    row("own", "本公司 10 年中位數（納入值）", f"={L(CIN)}{first}", "0.00", "空白＝本公司有效年數不足")
    row("npeer", "納入的同業家數", f"=COUNT({pin})" if pin else "=0", "0")
    row("pmed", "同業 10 年中位數的中位數", f'=IF(B{r - 1}>0,MEDIAN({pin}),"")' if pin else '=""', "0.00")
    row("raw", "原始比值（本公司 ÷ 同業）", f'=IF(AND(ISNUMBER(B{R["own"]}),B{R["npeer"]}>={MIN_PEERS}),B{R["own"]}/B{R["pmed"]},"")', "0.00",
        "＜1＝本公司長期比同業便宜（折價），＞1＝長期溢價")
    row("lo", "夾值下限", RM_LO, "0.00", "[Assumed] 避免單一歷史異常主導；tools/relative_multiple.py 的 RM_LO")
    ws.cell(R["lo"], 2).font = BLUE
    row("hi", "夾值上限", RM_HI, "0.00", "[Assumed] tools/relative_multiple.py 的 RM_HI")
    ws.cell(R["hi"], 2).font = BLUE
    row("fac", "相對倍數調整係數", f'=IF(ISNUMBER(B{R["raw"]}),MIN(MAX(B{R["raw"]},B{R["lo"]}),B{R["hi"]}),1)', "0.00",
        "資料不足時＝1（不調整）")
    ws.cell(R["fac"], 1).font = BOLD
    ws.cell(R["fac"], 2).fill = KEY
    if comp_ref:
        row("adj", "Comparable 目標價（相對倍數調整後）", f"={comp_ref}*B{R['fac']}", pxfmt,
            "＝Comparable 目標價（市場主流同業法）× 調整係數；加權目標價使用本列")
        ws.cell(R["adj"], 1).font = BOLD
        ws.cell(R["adj"], 2).fill = KEY
        wb.defined_names["tp_comp_adj"] = DefinedName("tp_comp_adj", attr_text=f"{ws.title}!$B${R['adj']}")
    wb.defined_names["rm_factor"] = DefinedName("rm_factor", attr_text=f"{ws.title}!$B${R['fac']}")
    R["head"] = h
    R["first"] = first
    R["last"] = last
    R["title"] = r0
    return R
