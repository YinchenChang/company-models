# AI半導體 目標價總表產生器（v0.2.4 起由 repo 產生；之前版本由 chat 端手工彙整）
# 用法：python3 tools/target_summary/build_summary.py <輸出.xlsx> [--prev <前版總表.xlsx>]
# 讀：tools/target_summary/universe.csv（公司、代號、次產業、市場、資料夾）＋各資料夾 dist/*.xlsx 的 Val 頁（LibreOffice 重算後的值）。
# 新增（Andy 2026-10-27 15:20「照建議做」）：
#   (1) 穩健度欄：把加權目標價換成「共識前瞻 EPS」口徑後的上檔空間與投資建議（[Derived] 近似：DCF 與 Comparable 都約與 EPS 成正比，
#       故 目標價 × 共識前瞻 EPS ÷ 本模型前瞻 EPS；兩者期間可能差約一年，見各模型 Recon）；
#   (2) 「市場別分布」頁：各市場的評等分布、買進率，以及台股依次產業拆分（診斷評等是否偏向特定市場）；
#   (3) 「相對倍數調整」頁：Comparable 改用「同業中位數 × 本公司 10 年相對同業倍數比」（2026-10-29 起已進加權目標價），列調整前後與敏感度。
import csv
import glob
import math
import os
import re
import sys

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(ROOT, "tools"))
import relative_multiple as RM  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
VER, VDATE = "v0.2.6", "2026-10-29"   # 每次升版改這裡
BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC")

LAB = {  # Val 頁列標籤（與 markets/invest_map/extract_models.py 的 V2_* 相同）
    "tp": r"^目標價（DCF 與 Comparable 加權）",
    "dcf": r"^DCF 目標價（情境加權）",
    "comp": r"^Comparable 目標價（市場主流同業法）",
    "cons": r"^市場共識目標價",
    "up": r"^上檔空間：加權目標價",
    "upc": r"^上檔空間：Comparable 目標價",
    "geps": r"^與市場共識目標價的差距：EPS 因素",
    "tier": r"^上檔空間級距",
    "rec": r"^投資建議",
}
OBS = re.compile(r"^重點觀察 (\d)：")


def rec(u):
    if u is None:
        return None
    return "強力買進" if u > 0.5 else "買進" if u > 0.25 else "中立" if u >= -0.25 else "放空" if u >= -0.5 else "強力放空"


def tier(u):
    return ("大於 50%" if u > 0.5 else "25%–50%" if u > 0.25 else "0%–25%" if u >= 0 else
            "0%～−25%" if u >= -0.25 else "−25%～−50%" if u >= -0.5 else "小於 −50%")


def read_model(folder):
    fs = [f for f in glob.glob(os.path.join(ROOT, folder, "dist", "*.xlsx"))]
    assert len(fs) == 1, (folder, fs)
    f = fs[0]
    wb = openpyxl.load_workbook(f, data_only=True, read_only=True)
    out = {"file": os.path.basename(f)}
    for r in wb["Val"].iter_rows(values_only=True):
        a = str(r[0] or "")
        for k, pat in LAB.items():
            if k not in out and re.search(pat, a) and (r[1] is not None and not (isinstance(r[1], str) and k not in ("tier", "rec"))):
                out[k] = r[1]   # 取第一個有值的列（標題列、說明列的 B 欄為空）
    obs = {}
    for r in wb["README"].iter_rows(values_only=True):
        a = str(r[0] or "")
        m = OBS.match(a)
        if m:
            obs[int(m.group(1))] = a[m.end():]
    out["obs"] = obs
    return out


def fisher_one_sided(a, b, c, d):
    """2×2 表 [[a,b],[c,d]] 的 Fisher 精確檢定（雙尾，以超幾何機率 ≤ 觀察值者加總）。"""
    n = a + b + c + d
    r1, c1 = a + b, a + c
    def p(x):
        return math.comb(r1, x) * math.comb(n - r1, c1 - x) / math.comb(n, c1)
    p0 = p(a)
    lo, hi = max(0, c1 - (n - r1)), min(r1, c1)
    return sum(p(x) for x in range(lo, hi + 1) if p(x) <= p0 + 1e-12)


def main():
    out = sys.argv[1]
    prev = None
    if "--prev" in sys.argv:
        prev = sys.argv[sys.argv.index("--prev") + 1]
    U = list(csv.DictReader(open(os.path.join(HERE, "universe.csv"), encoding="utf-8")))
    P = {}
    if prev:
        ws = openpyxl.load_workbook(prev, data_only=True)["目標價總表"]
        hd = [c.value for c in ws[2]]
        need = ("上檔空間：加權", "投資建議（固定規則）", "模型檔案")
        if all(n in hd for n in need):
            iu, ir, iff = (hd.index(n) for n in need)
        else:   # v0.2.3 手工版的固定欄位；其他版面直接報錯，不靜默讀錯欄
            assert len(hd) < 30, f"--prev 表頭無法辨識：{hd}"
            iu, ir, iff = 9, 15, 22
        for r in ws.iter_rows(min_row=3, values_only=True):
            if r[1]:
                P[r[1]] = {"up": r[iu], "rec": r[ir], "file": r[iff]}
    rows = []
    for u in U:
        m = read_model(u["folder"])
        px = m["tp"] / (1 + m["up"])
        g = m.get("geps")
        tp_c = m["tp"] / (1 + g) if isinstance(g, (int, float)) else None
        up_c = tp_c / px - 1 if tp_c else None
        rows.append({**u, **m, "px": px, "tp_c": tp_c, "up_c": up_c})

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "目標價總表"
    ws["A1"] = (f"AI半導體 目標價總表 {VER}（{VDATE}；DCF 50%＋Comparable 50%；數字取自 company-models main 各模型 dist 的 Val 頁，"
                "重點觀察取自 README 頁；由 tools/target_summary/build_summary.py 產生）")
    ws["A1"].font = BOLD
    H = ["次產業", "公司", "代號", "市場", "幣別", "現價", "DCF 目標價", "Comparable 目標價", "加權目標價", "市場共識目標價",
         "上檔空間：加權", "上檔空間：Comparable", "上檔空間：市場共識", "加權 ÷ 共識 − 1", "DCF ÷ 現價",
         "上檔空間級距（六級）", "投資建議（固定規則）", "模型 Val 頁投資建議", "級距與建議一致",
         "前版上檔空間：加權", "前版投資建議", "與前版相比",
         "本模型前瞻 EPS ÷ 共識前瞻 EPS − 1", "穩健度：共識 EPS 下加權目標價", "穩健度：共識 EPS 下上檔空間", "穩健度：共識 EPS 下投資建議", "穩健度：評等是否相同",
         "重點觀察 1", "重點觀察 2", "重點觀察 3", "模型檔案"]
    for j, h in enumerate(H, 1):
        c = ws.cell(2, j, h)
        c.font = BOLD
        c.fill = HDR
        c.alignment = Alignment(wrap_text=True, vertical="top")
    rows.sort(key=lambda r: -r["up"])
    for i, r in enumerate(rows, 3):
        pv = P.get(r["公司"], {})
        rr = rec(r["up"])
        vals = [r["次產業"], r["公司"], r["代號"], r["市場"], r["幣別"], r["px"], r["dcf"], r["comp"], r["tp"], r["cons"],
                r["up"], r["upc"], r["cons"] / r["px"] - 1 if r.get("cons") else None, r["tp"] / r["cons"] - 1 if r.get("cons") else None,
                r["dcf"] / r["px"],
                tier(r["up"]), rr, r.get("rec"), "是" if r.get("rec") == rr else "否",
                pv.get("up"), pv.get("rec"),
                ("不變" if pv.get("rec") == rr else f"{pv.get('rec')} → {rr}") if pv else "前版無",
                r.get("geps"), r["tp_c"], r["up_c"], rec(r["up_c"]), ("是" if rec(r["up_c"]) == rr else "否") if r["up_c"] is not None else "無資料",
                r["obs"].get(1), r["obs"].get(2), r["obs"].get(3), r["file"]]
        for j, v in enumerate(vals, 1):
            c = ws.cell(i, j, v)
            if j in (6, 7, 8, 9, 10, 24):
                c.number_format = "#,##0.00"
            if j in (11, 12, 13, 14, 20, 23, 25):
                c.number_format = "0.0%"
            if j == 15:
                c.number_format = "0.00"
            if j in (26, 27) and v == "否":
                c.fill = KEY
    last = 2 + len(rows)
    ws.freeze_panes = "C3"
    for col, w in zip("ABCDEFGHIJKLMNOPQRSTUVWXYZ", [10, 14, 9, 7, 6] + [11] * 21):
        ws.column_dimensions[col].width = w

    # ---- 投資建議統計 ----
    st = wb.create_sheet("投資建議統計")
    st.append(["投資建議", "家數（模型）", "家數（穩健度：共識 EPS）"])
    for c in st[1]:
        c.font = BOLD
        c.fill = HDR
    for k, lab in enumerate(["強力買進", "買進", "中立", "放空", "強力放空"], 2):
        st.cell(k, 1, lab)
        st.cell(k, 2, f'=COUNTIF(目標價總表!Q3:Q{last},A{k})')
        st.cell(k, 3, f'=COUNTIF(目標價總表!Z3:Z{last},A{k})')
    st.cell(7, 1, "合計")
    st.cell(7, 2, "=SUM(B2:B6)")
    st.cell(7, 3, "=SUM(C2:C6)")
    st.cell(9, 1, "穩健度評等與模型評等不同的家數")
    st.cell(9, 2, f'=COUNTIF(目標價總表!AA3:AA{last},"否")')

    # ---- 市場別分布 ----
    md = wb.create_sheet("市場別分布")
    md["A1"] = "市場別評等分布（診斷：評等是否偏向特定市場；公式直接讀「目標價總表」頁）"
    md["A1"].font = BOLD
    md["A2"] = ("讀法：買進率＝（強力買進＋買進）÷ 家數。2026-10-27 分析結論：偏向的成因是「低倍數股」（台股伺服器硬體、A 股光模組、韓國記憶體）"
                "而非市場本身；台股半導體與載板的買進率與美股相近。樣本小（A 股 3 家），Fisher 檢定只作參考。")
    mk = ["台股", "A股", "韓股", "日股", "美股", "歐股"]
    H2 = ["市場", "家數", "強力買進", "買進", "中立", "放空", "強力放空", "買進以上", "買進率", "占全部買進以上",
          "買進以上（穩健度）", "買進率（穩健度）", "上檔空間中位數", "DCF ÷ 現價中位數", "加權 ÷ 共識 − 1 中位數"]
    for j, h in enumerate(H2, 1):
        c = md.cell(4, j, h)
        c.font = BOLD
        c.fill = HDR
    R = "目標價總表!"
    rng = lambda col: f"{R}${col}$3:${col}${last}"
    for i, m in enumerate(mk + ["合計"], 5):
        md.cell(i, 1, m)
        cond = f'{rng("D")},A{i}' if m != "合計" else None
        if m == "合計":
            md.cell(i, 2, f"=SUM(B5:B{i - 1})")
            for j in range(3, 9):
                L = "CDEFGH"[j - 3]
                md.cell(i, j, f"=SUM({L}5:{L}{i - 1})")
            md.cell(i, 11, f"=SUM(K5:K{i - 1})")
        else:
            md.cell(i, 2, f"=COUNTIF({cond})")
            for j, lab in enumerate(["強力買進", "買進", "中立", "放空", "強力放空"], 3):
                md.cell(i, j, f'=COUNTIFS({rng("D")},$A{i},{rng("Q")},"{lab}")')
            md.cell(i, 8, f"=C{i}+D{i}")
            md.cell(i, 11, f'=COUNTIFS({rng("D")},$A{i},{rng("Z")},"強力買進")+COUNTIFS({rng("D")},$A{i},{rng("Z")},"買進")')
        md.cell(i, 9, f'=IF(B{i}>0,H{i}/B{i},"")').number_format = "0%"
        md.cell(i, 10, f"=H{i}/$H${5 + len(mk)}").number_format = "0%"
        md.cell(i, 12, f'=IF(B{i}>0,K{i}/B{i},"")').number_format = "0%"
        sub = [r for r in rows if (m == "合計" or r["市場"] == m)]
        if sub:
            med = lambda xs: sorted(xs)[len(xs) // 2] if len(xs) % 2 else (sorted(xs)[len(xs) // 2 - 1] + sorted(xs)[len(xs) // 2]) / 2
            md.cell(i, 13, med([r["up"] for r in sub])).number_format = "0.0%"
            md.cell(i, 14, med([r["dcf"] / r["px"] for r in sub])).number_format = "0.00"
            md.cell(i, 15, med([r["tp"] / r["cons"] - 1 for r in sub if r.get("cons")])).number_format = "0.0%"
    r0 = 6 + len(mk) + 2
    md.cell(r0, 1, "台股依次產業拆分").font = BOLD
    for j, h in enumerate(["組別", "家數", "買進以上", "買進率", "公司"], 1):
        c = md.cell(r0 + 1, j, h)
        c.font = BOLD
        c.fill = HDR
    hw = {"伺服器組裝", "電源與散熱"}
    groups = [("台股伺服器硬體（組裝、電源與散熱）", lambda r: r["市場"] == "台股" and r["次產業"] in hw),
              ("台股其他（半導體、載板與 PCB、網通）", lambda r: r["市場"] == "台股" and r["次產業"] not in hw),
              ("台股＋A股", lambda r: r["市場"] in ("台股", "A股")),
              ("其他市場", lambda r: r["市場"] not in ("台股", "A股"))]
    buy = lambda r: rec(r["up"]) in ("強力買進", "買進")
    cnt = {}
    for k, (nm, fn) in enumerate(groups, r0 + 2):
        sub = [r for r in rows if fn(r)]
        b = [r for r in sub if buy(r)]
        cnt[nm] = (len(b), len(sub) - len(b))
        md.cell(k, 1, nm)
        md.cell(k, 2, len(sub))
        md.cell(k, 3, len(b))
        md.cell(k, 4, len(b) / len(sub) if sub else None).number_format = "0%"
        md.cell(k, 5, "、".join(r["公司"] + ("*" if buy(r) else "") for r in sub))
    k = r0 + 2 + len(groups)
    a, b_ = cnt["台股＋A股"]
    c, d = cnt["其他市場"]
    md.cell(k, 1, "Fisher 精確檢定（台股＋A股 vs 其他市場的買進率；雙尾 p 值）")
    md.cell(k, 2, fisher_one_sided(a, b_, c, d)).number_format = "0.000"
    md.cell(k + 1, 1, "＊＝買進以上。p 值 ≥ 0.05 表示以目前樣本無法排除隨機差異。")
    md.column_dimensions["A"].width = 34
    md.column_dimensions["E"].width = 60

    # ---- 相對倍數調整（2026-10-29 起已進目標價） ----
    rm = wb.create_sheet("相對倍數調整")
    rm["A1"] = ("相對倍數調整（2026-10-29 起已進加權目標價；各模型 Val 第 0.5 節）：Comparable 目標倍數＝同業中位數 × 本公司 10 年中位數 ÷ 同業 10 年中位數的中位數，"
                f"夾在 [{RM.RM_LO}, {RM.RM_HI}]；有效年數 < {RM.MIN_YEARS} 或有效同業 < {RM.MIN_PEERS} 時係數＝1。")
    rm["A1"].font = BOLD
    H3 = ["公司", "市場", "代號", "倍數", "本公司 10 年中位數", "本公司有效年數", "納入同業家數", "同業 10 年中位數的中位數", "原始比值", "係數",
          "上檔空間：加權（調整前）", "上檔空間：加權（調整後＝模型）", "投資建議（調整前）", "投資建議（調整後＝模型）", "是否翻轉", "同業（有效年數）",
          "敏感度：近 5 年窗係數", "敏感度：近 5 年窗投資建議", "兩種窗口是否都翻轉", "對照：10-27 落後代理係數"]
    for j, h in enumerate(H3, 1):
        c = rm.cell(4, j, h)
        c.font = BOLD
        c.fill = HDR
    hist = RM.load()
    TR = []
    hist_old = RM.load(os.path.join(ROOT, "tools", "data", "rel_mult_hist.csv"))
    peers_by = {}
    pf = os.path.join(HERE, "peers.csv")
    for p in csv.DictReader(open(pf, encoding="utf-8")):
        peers_by.setdefault(p["folder"], []).append((p["name"], p["ticker"]))
    for i, r in enumerate(sorted(rows, key=lambda r: (r["市場"], -r["up"])), 5):
        met = "PB" if r["folder"] in ("skhynix", "samsung", "micron", "kioxia", "sndk") else "PE"
        s = RM.summary(r["own_ticker"], peers_by.get(r["folder"], []), met, hist)
        # 2026-10-29 起已接入模型：r["up"] 已含調整；「調整前」以 Comparable（市場主流同業法）重算
        pre_up = 0.5 * (r["dcf"] / r["px"] - 1) + 0.5 * (r["comp"] / r["px"] - 1)
        new_up = 0.5 * (r["dcf"] / r["px"] - 1) + 0.5 * (r["comp"] * s["factor"] / r["px"] - 1)
        assert abs(new_up - r["up"]) < 1e-6, (r["公司"], new_up, r["up"])
        pinfo = []
        for nm, t in peers_by.get(r["folder"], []):
            n = len([1 for y, (v, _) in hist.get((RM.norm(t), met), {}).items() if y in RM.YEARS and RM._valid(v, met)])
            pinfo.append(f"{nm}（{n}）")
        y0, m0 = RM.YEARS, RM.MIN_YEARS
        RM.YEARS, RM.MIN_YEARS = y0[-5:], 4
        s5 = RM.summary(r["own_ticker"], peers_by.get(r["folder"], []), met, hist)
        RM.YEARS, RM.MIN_YEARS = y0, m0
        up5 = 0.5 * (r["dcf"] / r["px"] - 1) + 0.5 * (r["comp"] * s5["factor"] / r["px"] - 1)
        RM.YEARS = list(range(2016, 2026))
        so = RM.summary(r["own_ticker"], peers_by.get(r["folder"], []), met, hist_old)
        RM.YEARS = y0
        f10, f5 = rec(pre_up) != rec(new_up), rec(pre_up) != rec(up5)
        TR.append((rec(pre_up), rec(new_up), rec(up5)))
        vals = [r["公司"], r["市場"], r["own_ticker"], "P/B" if met == "PB" else "P/E", s["own_med"], s["own_n"], s["n_peers"], s["peer_med"],
                s["raw"], s["factor"], pre_up, new_up, rec(pre_up), rec(new_up), "是" if f10 else "否", "、".join(pinfo),
                s5["factor"], rec(up5), "是" if (f10 and f5) else ("只有一種" if (f10 or f5) else "否"), so["factor"]]
        for j, v in enumerate(vals, 1):
            c = rm.cell(i, j, v)
            if j in (5, 8, 9, 10, 17, 20):
                c.number_format = "0.00"
            if j in (11, 12):
                c.number_format = "0.0%"
            if j == 15 and v == "是":
                c.fill = KEY
    ORD = ["強力放空", "放空", "中立", "買進", "強力買進"]
    n = len(TR)
    f10 = sum(1 for a, b, c in TR if b != a)
    f5 = sum(1 for a, b, c in TR if c != a)
    both = [(a, b, c) for a, b, c in TR if b != a and c != a]
    same_dir = sum(1 for a, b, c in both if (ORD.index(b) - ORD.index(a)) * (ORD.index(c) - ORD.index(a)) > 0)
    same_rec = sum(1 for a, b, c in both if b == c)
    agree = sum(1 for a, b, c in TR if b == c)
    rm["A2"] = (f"{VDATE} 改用 Bloomberg 前瞻倍數（BEST_PE_RATIO／記憶體 BEST_PX_BPS_RATIO，BF＝未來 12 個月；2015-12 起月底資料，"
                "最後一點為抓取當日的最新值、非 10 月月底）：每年窗口（前一年 11 月～當年 10 月）取月值中位數，再取 2017–2026 十個窗口的中位數。"
                "口徑與 Comparable（現在的同業前瞻倍數）一致，10-27 落後 P/E 試算的谷底年度雜訊已排除。"
                f"結果：10 年窗翻轉 {f10} 檔、近 5 年窗（2022–2026、門檻 4 年）翻轉 {f5} 檔；兩種窗口都翻轉 {len(both)} 檔，"
                f"其中方向相同 {same_dir} 檔、翻到同一評等 {same_rec} 檔（其餘只差一級）；兩種窗口評等相同 {agree}／{n} 檔。"
                "敏感度在 Q–T 欄（S 欄＝兩種窗口是否都翻轉）。「調整前」＝DCF 與 Comparable（市場主流同業法）各半。")
    rm["A2"].alignment = Alignment(wrap_text=True)
    rm.column_dimensions["A"].width = 14
    rm.column_dimensions["P"].width = 70

    # ---- 前瞻倍數年度窗口中位數（調整的輸入，供核對） ----
    dv = wb.create_sheet("前瞻倍數資料")
    dv["A1"] = ("相對倍數調整的輸入：tools/data/fwd_mult_hist.csv。每格＝該窗口（前一年 11 月～當年 10 月）Bloomberg 月底前瞻倍數的中位數；"
                "有效月數 < 6 的窗口留空。P/E 有效區間 0 < 值 ≤ 200、P/B 0 < 值 ≤ 20（逐月套用）。")
    dv["A1"].font = BOLD
    hh = ["代號", "倍數"] + [str(y) for y in RM.YEARS] + ["有效窗口數", "10 年中位數"]
    for j, h in enumerate(hh, 1):
        c = dv.cell(3, j, h)
        c.font = BOLD
        c.fill = HDR
    for i, (k, d) in enumerate(sorted(hist.items()), 4):
        dv.cell(i, 1, k[0])
        dv.cell(i, 2, "P/B" if k[1] == "PB" else "P/E")
        vv = []
        for j, y in enumerate(RM.YEARS, 3):
            if y in d and RM._valid(d[y][0], k[1]):
                dv.cell(i, j, d[y][0]).number_format = "0.00"
                vv.append(d[y][0])
        dv.cell(i, 3 + len(RM.YEARS), len(vv))
        if vv:
            dv.cell(i, 4 + len(RM.YEARS), sorted(vv)[len(vv) // 2] if len(vv) % 2 else sum(sorted(vv)[len(vv) // 2 - 1:len(vv) // 2 + 1]) / 2).number_format = "0.00"

    ex = wb.create_sheet("說明")
    for k, t in enumerate([
        f"AI半導體 目標價總表 {VER}（{VDATE}）",
        "來源：各模型 dist 的 Val 頁（「目標價（DCF 與 Comparable 加權）」等列，標籤與 markets/invest_map/extract_models.py 相同）；重點觀察取自 README 頁。",
        "現價＝加權目標價 ÷（1＋上檔空間：加權目標價）（模型 Val 頁的現價，與模型一致）。",
        "穩健度（新增）：加權目標價 ×（共識前瞻 EPS ÷ 本模型前瞻 EPS），即把 EPS 換成共識口徑、倍數與方法不變。[Derived] 近似——DCF 與 Comparable 約與 EPS 成正比；"
        "本模型的「12 個月後前瞻 EPS」比共識前瞻 EPS 晚約一年，成長股的差距有一部分是期間差，所以本欄偏保守。用途：看評等是否主要靠 EPS 假設撐起來。",
        "市場別分布（新增）：各市場評等分布與買進率；台股另依次產業拆分。",
        "相對倍數調整（2026-10-29 起已進目標價）：見該頁第 2 列。資料檔 tools/data/fwd_mult_hist.csv（Bloomberg 前瞻倍數年度窗口中位數；10-27 落後代理 rel_mult_hist.csv 保留對照）、規則 tools/relative_multiple.py。",
        "投資建議固定規則（Andy 2026-10-25 14:17）：> +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、< −50% 強力放空。",
    ], 1):
        ex.cell(k, 1, t)
    ex.column_dimensions["A"].width = 160
    wb.save(out)


if __name__ == "__main__":
    main()
