#!/usr/bin/env python3
"""Waymo（Alphabet 分部，視同未上市）個股模型 v0.2 建置腳本（v0.2：對帳 2，市場、滲透、份額、價格、營運成本、稼動率改讀線 1a v0.2）。

Robotaxi 分部模型特別規格（專案規格 2026-11-16）：主頁＝「滲透」，最上方為「價格」區塊（每英里實收價格）；
其次「市場天花板與情境」「每英里經濟」「估值」「資料來源比較」「對帳」。頁面與口徑與 physical_ai/tesla_robotaxi/ 一致。

用法：python3 physical_ai/waymo/build_waymo.py   →  physical_ai/waymo/dist/20261117_Waymo個股模型_v0.2.xlsx
建置後自動：LibreOffice 重算 → 0 錯誤值檢查 → xlsx 與 Python 鏡像逐情境比對 → in_layout／sens_coverage／governance 核對。
"""
import datetime as dt
import hashlib
import json
import math
import os
import shutil
import subprocess
import sys
import tempfile
import glob

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as CL

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VERSION = "v0.2"
AS_OF = "2026-11-17"
OUT = os.path.join(HERE, "dist", f"20261117_Waymo個股模型_{VERSION}.xlsx")

BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
SEC = PatternFill("solid", fgColor="FFF2CC")
WRAP = Alignment(wrap_text=True, vertical="top")
YEARS = list(range(2026, 2036))
YC = {y: CL(3 + i) for i, y in enumerate(YEARS)}  # C..L
SCN = ["A", "B", "C"]
SCN_NAME = {"A": "情境 A（線高滲透：低價、低成本、Tesla 份額較高）", "B": "情境 B（線基準）", "C": "情境 C（線低滲透：高價、高成本）"}

# ---------------------------------------------------------------------------
# 〇、線 1a v0.2 讀入（線權威；對帳 2 工程結論）。情境 A＝線高滲透（配低成本、低價），B＝基準，C＝低滲透（配高成本、高價）
# ---------------------------------------------------------------------------
LINE_GLOB = os.path.join(HERE, "..", "lines", "robotaxi", "dist", "*Robotaxi單位經濟_v0.2.xlsx")
LINE_ANCH = [2026, 2028, 2030, 2032]
LINE_BLK = {"A": 85, "B": 44, "C": 3}                    # 情境頁各情境區塊首列
LINE_BLK_LAB = {"A": "高（高滲透）", "B": "基準（基準）", "C": "低（低滲透）"}
LINE_CASE = {"A": 0, "B": 1, "C": 2}                     # 成本／價格欄：低成本＋低價／基準／高成本＋高價
LINE_CASE_LAB = {"A": "低成本／低價", "B": "基準", "C": "高成本／高價"}
LINE_OFF = {"trips": (2, "美國叫車趟數"), "pen": (6, "Robotaxi 滲透率"), "wsh": (27, "Waymo 份額")}
ST_W, ST_W_PMD, CI_W_PRICE = 5, 20, 31                  # 堆疊頁 Waymo 第 1 項列、稼動率備忘列；成本輸入頁 W_PRICE 列


def _geo_path(v):
    """錨點 2026/28/30/32 → 2026–2035：奇數年取相鄰錨點幾何平均，2033 起持平於 2032（對帳 1 判斷題 2 預設）"""
    return {y: v[y] if y in v else (math.sqrt(v[y - 1] * v[y + 1]) if y < 2032 else v[2032]) for y in YEARS}


def load_line():
    path = sorted(glob.glob(LINE_GLOB))[-1]
    wb = openpyxl.load_workbook(path, data_only=True)
    sc, st, ci, inp, pr = wb["情境"], wb["每英里成本堆疊"], wb["成本輸入"], wb["輸入"], wb["每英里價格"]
    fn = os.path.basename(path)
    L = {"file": fn, "src": {}}
    for s, b in LINE_BLK.items():                         # 版面斷言：列號改動時立即失敗
        assert str(sc.cell(b, 1).value).startswith("情境："), (s, b)
        for k, (o, lab) in LINE_OFF.items():
            assert str(sc.cell(b + o, 1).value).startswith(lab), (s, k, sc.cell(b + o, 1).value)
            for i, y in enumerate(YEARS):
                L[(k, s, y)] = sc.cell(b + o, 3 + i).value
                L["src"][(k, s, y)] = f"{fn}｜情境!{CL(3 + i)}{b + o}（{LINE_BLK_LAB[s]}：{lab}）"
    assert st.cell(ST_W - 2, 1).value == "Waymo" and str(st.cell(ST_W, 1).value).startswith("1 車輛")
    assert str(st.cell(ST_W + 2, 1).value).startswith("3 能源") and str(st.cell(ST_W + 7, 1).value).startswith("8 場站")
    assert str(st.cell(ST_W_PMD, 1).value).startswith("備忘：每車每日付費英里") and ci.cell(CI_W_PRICE, 1).value == "W_PRICE"
    for s, c in LINE_CASE.items():
        ox, pm, pp = {}, {}, {}
        for i, y in enumerate(LINE_ANCH):
            sc_, ic = 3 + i * 3 + c, 5 + i * 3 + c
            ox[y] = sum(st.cell(ST_W + k, sc_).value for k in range(2, 8))
            pm[y] = st.cell(ST_W_PMD, sc_).value
            pp[y] = ci.cell(CI_W_PRICE, ic).value
            L["src"][("opx", s, y)] = f"{fn}｜每英里成本堆疊!{CL(sc_)}{ST_W + 2}:{CL(sc_)}{ST_W + 7}（Waymo 第 3–8 項，{LINE_CASE_LAB[s]}）"
            L["src"][("pmd", s, y)] = f"{fn}｜每英里成本堆疊!{CL(sc_)}{ST_W_PMD}（Waymo 每車每日付費英里，{LINE_CASE_LAB[s]}）"
            L["src"][("price", s, y)] = f"{fn}｜成本輸入!{CL(ic)}{CI_W_PRICE}（W_PRICE，{LINE_CASE_LAB[s]}）"
        for k, v in (("opx", ox), ("pmd", pm), ("price", pp)):
            for y, x in _geo_path(v).items():
                L[(k, s, y)] = x
                if y not in LINE_ANCH:
                    L["src"][(k, s, y)] = f"{fn}｜" + (f"{y - 1}／{y + 1} 錨點的幾何平均" if y < 2032 else "2032 錨點持平（線的權威只到 2032）")
    row = {inp.cell(r, 1).value: r for r in range(1, inp.max_row + 1)}
    L["trip_mkt"] = (inp.cell(row["M_TRIP_MI"], 4).value, inp.cell(row["M_TRIP_MI"], 5).value, inp.cell(row["M_TRIP_MI"], 6).value)
    L["src"]["trip_mkt"] = f"{fn}｜輸入!D{row['M_TRIP_MI']}:F{row['M_TRIP_MI']}（M_TRIP_MI，Assumed）"
    el = [inp.cell(row["P_ELAST"], c).value for c in (4, 5, 6)]  # 低／基準／高滲透
    L["eps"] = {"A": -el[2], "B": -el[1], "C": -el[0]}
    L["src"]["eps"] = f"{fn}｜輸入!D{row['P_ELAST']}:F{row['P_ELAST']}（P_ELAST，Analogy：NBER w31330）"
    rr = next(r for r in range(1, pr.max_row + 1) if str(pr.cell(r, 1).value or "").startswith("CPUC 每趟付費英里"))
    L["tl0"] = pr.cell(rr, 2).value
    L["src"]["tl0"] = f"{fn}｜每英里價格!B{rr}（CPUC Q2 2026 總英里÷趟數×付費占比）"
    return L


LINE = load_line()
LINE_KEYS = (("trips", "lt", "線 美國叫車年趟數（不含誘發）", "百萬趟"), ("pen", "lpen", "線 Robotaxi 滲透率（替代趟數÷叫車趟數）", "%"),
             ("wsh", "lws", "線 Waymo 占美國 Robotaxi 趟數份額", "%"), ("price", "lp", "線 Waymo 每付費英里價格", "US$/英里"),
             ("opx", "lx", "線 Waymo 單位營運成本第 3–8 項（不含車輛攤提、融資）", "US$/英里"), ("pmd", "lm", "線 Waymo 每車每日付費英里", "英里/日"))

# ---------------------------------------------------------------------------
# 一、輸入（In 頁）：鍵、項目、單位、低、基準、高、標記、出處、來源誘因、說明
# ---------------------------------------------------------------------------
IN = [
    ("§", "A. 市場（美國叫車趟數、滲透率、Waymo 份額皆讀線 1a v0.2，見本頁 E 節；v0.2 起本節只留轉換閘門與錨）"),
    ("ph", "人類叫車每付費英里乘客實付價格（含小費，全美；顯示錨）", "US$/英里", None, 3.7, None, "Derived",
     "Gridwise 2026 Q3：Uber 司機每英里毛收入 US$2.65、平台抽成中位數 22.44%（線 1a 框架）",
     "", "2.65 ÷（1 − 0.2244）× 1.08 ≈ 3.69。v0.2 起不進誘發公式（誘發改以 Waymo 自身價格相對 2026，對帳 2），只在滲透頁與每英里經濟頁作錨"),
    ("vmt", "美國 12 個月移動車輛行駛里程（2026-08）", "十億英里", None, 3335.695, None, "Verified", "FRED／FHWA TRAVLTRS（與 tesla_robotaxi 同值）", "", "政府統計"),
    ("ldv_sh", "VMT 中可被 Robotaxi 取代的小客車占比", "%", 0.85, 0.90, 0.92, "Assumed", "—", "", "與 tesla_robotaxi 同"),
    ("p_own", "自有車每英里全成本（AAA 2025，15,000 英里／年平均）", "US$/英里", 0.5587, 0.772, None, "Verified", "AAA Your Driving Costs 2025-09-16：年 11,577 美元；小轎車 55.87 美分（低）", "", "與 tesla_robotaxi 同"),
    ("cf", "便利溢價（Robotaxi 每英里價格可高於自有車成本的倍數，轉換閘門門檻）", "倍", None, 1.0, 1.5, "Assumed", "—", "",
     "v0.2 基準 1.5→1.0（與 tesla_robotaxi v0.2 同：對帳 1 判斷題 1 預設乙，價格 ≤ 0.77 全開、≥ 1.155 關閉）；高＝1.5（v0.1）。閘門＝MIN(1,MAX(0,(1.5×門檻−價)/(0.5×門檻)))，門檻＝p_own×cf"),
    ("§", "B. Waymo 營運參數（價格、營運成本、稼動率讀線，見 E 節；本節為 Waymo 自有參數）"),
    ("tl_end", "Waymo 每趟付費英里 2035（2026 讀線 CPUC 實測 tl0，線性收斂至此值）", "英里/趟", 3.73, 6.0, 7.1, "Assumed",
     "線 1a M_TRIP_MI：美國叫車平均每趟 6.0 英里（4.5–7.1）；CPUC：Waymo 加州 2026 Q2 每趟約 3.7 付費英里", "",
     "對帳 2 判斷題 2 預設丙：服務區擴到高速公路與機場後行程變長，2035 收斂到全美叫車平均；低＝乙（維持 2026 實測 3.73）；高＝線上緣 7.1"),
    ("util_k", "稼動率乘數（對線每車每日付費英里的調整；基準 1.0）", "倍", 0.8, 1.0, 1.2, "Assumed", "—", "", "與 tesla_robotaxi util_k 同；只在 Sens 動"),
    ("vcost26", "單車取得成本 2026（Ojai＋第六代套件）", "千美元", 100, 125, 160, "Interested-party",
     "Morgan Stanley 估計（TNW 轉述，線 1a）：Ojai 約 12.5 萬美元；舊款 I-PACE 約 20 萬", "券商估計（有覆蓋 GOOGL 的利益）", "線 W_VEH 2026 基準 14 萬；車輛攤提不讀線（與 tesla_robotaxi 同規則：只讀第 3–8 項）"),
    ("vcost_dec", "單車成本年降幅", "%/年", 0.04, 0.08, 0.12, "Assumed", "—", "", "套件成本下降、Zeekr／Hyundai 量產"),
    ("vlife", "車輛經濟壽命", "年", 5, 6, 8, "Assumed", "—", "", "生涯 30–50 萬英里（線 1a）÷ 每車年英里"),
    ("fix26", "固定營運費用 2026（研發、地圖、管銷、擴城）", "十億美元", 3.5, 5.0, 6.5, "Derived",
     "Alphabet 2026 Q3 Other Bets 營業虧損 US$20.5 億、營收 3.72 億（Fortune 2026-10-30）；年化約 82 億為 Other Bets 合計", "",
     "Waymo 未單獨揭露；扣除 Wing、Verily 等其他單位後取 3.5–6.5。不讀線第 9–10 項（線按車隊分攤，口徑不同；與 tesla_robotaxi 同規則）"),
    ("fix_g", "固定營運費用年成長", "%/年", 0.0, 0.02, 0.05, "Assumed", "—", "", "擴城與國際化 vs 研發成熟"),
    ("§", "C. 估值"),
    ("tax", "稅率（獲利時）", "%", 0.21, 0.21, 0.25, "Verified", "美國聯邦法定稅率 21%（州稅另計為敏感度高）", "", "虧損年不計稅盾（獨立估值口徑）"),
    ("nwc", "營運資金增加 ÷ 營收增量", "%", -0.05, 0.0, 0.05, "Assumed", "—", "", "預收乘車款 vs 應收（經 Uber 平台的城市）"),
    ("wacc", "折現率（WACC，未上市自駕營運商）", "%", 0.10, 0.12, 0.15, "Assumed", "—", "", "Alphabet 本身 WACC 約 9–10%；Waymo 為早期、技術與監管風險較高的分部"),
    ("g_tv", "永續成長率", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 標準規格第 3 節：預設 3%（2–4%）", "", "RONIC＝WACC（價值中性）"),
    ("netcash", "Waymo 淨現金（評價日）", "十億美元", 5, 10, 15, "Assumed",
     "2026-02 募資 160 億美元（AP、Bloomberg）；2026-10 首筆定期貸款 50 億美元（Fortune 2026-10-30）", "",
     "16 ＋ 5 − 2026 年燒現金（未取得）− 債務 5；現金餘額未揭露"),
    ("val_round", "最近一輪融資估值（投後）", "十億美元", None, 126, None, "Interested-party",
     "Waymo 2026-02-02 宣布 160 億美元融資、估值 1,260 億美元（AP、Bloomberg、Fortune 2026-02-03）", "Waymo／Alphabet 與投資人（Alphabet 為主要出資者）", "只作對照，不進模型"),
    ("§", "D. 情境專屬輸入（A 樂觀／B 基準／C 保守；全部 Assumed）"),
]
SCN_IN = {
    # 鍵前綴: (項目, 單位, A, B, C, 說明)
    "c35": ("自有車 VMT 轉換為 Robotaxi 的比例（閘門開啟時）2035", "%", 0.10, 0.03, 0.005, "與 tesla_robotaxi c35 相同；2027 起線性 × 閘門"),
    "eps": ("價格彈性（誘發需求）：誘發倍數＝(Waymo 2026 價格 ÷ 當年價格)^ε", "倍數", LINE["eps"]["A"], LINE["eps"]["B"], LINE["eps"]["C"],
            "讀線 P_ELAST（Analogy：NBER w31330；線高／基準／低滲透）；相對 2026 是因 2026 趟數已含當時價格（對帳 2 工程結論，避免重複計入）"),
    "intl35": ("國際趟數 ÷ 美國趟數 2035", "%", 0.30, 0.15, 0.0, "倫敦、東京（2026 宣布）、加拿大、德國、新加坡、西班牙（Q3 法說）；2027 起線性。線無海外，沿用本模型"),
}
SCN_W = {"A": 0.25, "B": 0.50, "C": 0.25}


def line_rows():
    """In 頁 E 節：(鍵, 標籤, 單位, 值, 出處)；值由 load_line() 從線成品讀入（不手打）"""
    rows = [("trip_mkt", "線 美國叫車平均每趟付費英里（M_TRIP_MI 基準；tl_end 的基準值同此）", "英里/趟", LINE["trip_mkt"][1], LINE["src"]["trip_mkt"]),
            ("tl0", "線 Waymo 每趟付費英里 2026（CPUC 推算）", "英里/趟", LINE["tl0"], LINE["src"]["tl0"])]
    for kind, pre, lab, unit in LINE_KEYS:
        for y in YEARS:
            for s in SCN:
                rows.append((f"{pre}_{y}_{s}", f"情境 {s}：{lab} {y}", unit, LINE[(kind, s, y)], LINE["src"][(kind, s, y)]))
    return rows


def base_params():
    p = {r[0]: r[4] for r in IN if r[0] != "§"}
    p["stub"] = 45 / 365
    for k, v in SCN_IN.items():
        for i, s in enumerate(SCN):
            p[f"{k}_{s}"] = v[2 + i]
    for k, lab, unit, v, src in line_rows():
        p[k] = v
    return p


# ---------------------------------------------------------------------------
# 二、Python 鏡像（與 xlsx 公式一一對應；Sens 與反解用）
# ---------------------------------------------------------------------------
def run(p, s):
    out = {k: {} for k in ["us", "pen", "rpm", "jev", "gate", "conv", "av", "wsh", "wus", "intl", "trips", "wk", "fleet", "pm", "L",
                           "cpm", "vc", "mpv", "dep_mi", "opc", "fix", "fix_mi", "full", "rev", "var", "dep",
                           "ebit", "nopat", "capex", "dnwc", "fcff", "pv"]}
    prev_fleet = prev_rev = None
    for y in YEARS:
        n = y - 2026
        us = p[f"lt_{y}_{s}"] / 1000
        pen = p[f"lpen_{y}_{s}"]
        rpm = p[f"lp_{y}_{s}"]
        jev = (p[f"lp_2026_{s}"] / rpm) ** p[f"eps_{s}"]
        gate = min(1, max(0, (1.5 * p["p_own"] * p["cf"] - rpm) / (0.5 * p["p_own"] * p["cf"])))
        conv = p["vmt"] * p["ldv_sh"] * p[f"c35_{s}"] * n / 9 * gate
        L = p["tl0"] + (p["tl_end"] - p["tl0"]) * n / 9
        av = us * pen * jev + conv / L
        wsh = p[f"lws_{y}_{s}"]
        wus = av * wsh * 1000
        intl = p[f"intl35_{s}"] * n / 9
        trips = wus * (1 + intl)
        pm = trips * L
        fleet = pm * 1000 / (p[f"lm_{y}_{s}"] * p["util_k"] * 365)
        cpm = p[f"lx_{y}_{s}"]
        vc = p["vcost26"] * (1 - p["vcost_dec"]) ** n
        mpv = p[f"lm_{y}_{s}"] * p["util_k"] * 365 / 1000
        dep_mi = vc / p["vlife"] / mpv
        opc = cpm + dep_mi
        fix = p["fix26"] * (1 + p["fix_g"]) ** n
        fix_mi = fix * 1000 / pm
        rev = pm * rpm / 1000
        var = pm * cpm / 1000
        dep = fleet * vc / p["vlife"] / 1000
        ebit = rev - var - fix - dep
        nopat = ebit - max(0, ebit) * p["tax"]
        pf = fleet if prev_fleet is None else prev_fleet
        capex = (max(0, fleet - pf) + pf / p["vlife"]) * vc / 1000
        dnwc = 0 if prev_rev is None else p["nwc"] * (rev - prev_rev)
        fcff = nopat + dep - capex - dnwc
        frac = p["stub"] if y == 2026 else 1
        t = p["stub"] / 2 if y == 2026 else p["stub"] + (y - 2027) + 0.5
        pv = fcff * frac / (1 + p["wacc"]) ** t
        for k, v in dict(us=us, pen=pen, rpm=rpm, jev=jev, gate=gate, conv=conv, av=av, wsh=wsh, wus=wus, intl=intl, trips=trips, wk=trips / 52,
                         fleet=fleet, pm=pm, L=L, cpm=cpm, vc=vc, mpv=mpv, dep_mi=dep_mi, opc=opc, fix=fix, fix_mi=fix_mi,
                         full=opc + fix_mi, rev=rev, var=var, dep=dep, ebit=ebit, nopat=nopat, capex=capex, dnwc=dnwc,
                         fcff=fcff, pv=pv).items():
            out[k][y] = v
        prev_fleet, prev_rev = fleet, rev
    sum_pv = sum(out["pv"].values())
    n36 = out["nopat"][2035] * (1 + p["g_tv"])
    tv = n36 * (1 - p["g_tv"] / p["wacc"]) / (p["wacc"] - p["g_tv"])
    pvtv = tv / (1 + p["wacc"]) ** (p["stub"] + 9)
    ev = sum_pv + pvtv
    out.update(sum_pv=sum_pv, tv=tv, pvtv=pvtv, ev=ev, eq=ev + p["netcash"], pen26=out["pen"][2026])
    return out


def swap(p, pre, frm, to="B"):
    """Sens 用：把情境 to 的線路徑（鍵前綴 pre）換成情境 frm 的路徑"""
    q = dict(p)
    for y in YEARS:
        q[f"{pre}_{y}_{to}"] = p[f"{pre}_{y}_{frm}"]
    return q


def weighted(p):
    return sum(SCN_W[s] * run(p, s)["eq"] for s in SCN)


def solve(p, key, target, lo, hi, s="B", wt=False):
    f = (lambda x: weighted({**p, key: x})) if wt else (lambda x: run({**p, key: x}, s)["eq"])
    flo, fhi = f(lo) - target, f(hi) - target
    if flo * fhi > 0:
        return None
    for _ in range(80):
        m = (lo + hi) / 2
        fm = f(m) - target
        if flo * fm <= 0:
            hi = m
        else:
            lo, flo = m, fm
    return (lo + hi) / 2


# ---------------------------------------------------------------------------
# 三、xlsx 建置
# ---------------------------------------------------------------------------
def put(ws, r, c, v, font=None, fmt=None, fill=None, wrap=False):
    cell = ws.cell(r, c, v)
    if font: cell.font = font
    if fmt: cell.number_format = fmt
    if fill: cell.fill = fill
    if wrap: cell.alignment = WRAP
    return cell


def name(wb, key, ref):
    wb.defined_names[key] = DefinedName(key, attr_text=ref)


def build():
    P = base_params()
    wb = openpyxl.Workbook()
    wsR = wb.active
    wsR.title = "README"
    names = ["滲透", "市場天花板與情境", "每英里經濟", "估值", "資料來源比較", "對帳", "In", "Sens"]
    W = {n: wb.create_sheet(n) for n in names}

    # ---------------- In ----------------
    ws = W["In"]
    put(ws, 1, 1, "In — 輸入（藍字；基準欄 E 進主模型；D／F＝敏感度低／高，只用於 Sens 頁（一次動一個參數）、不進目標價；情境 A／B／C 的驅動見 D 節與本頁最後一節「情境驅動對照」）。金額：十億美元；每英里：美元。", BOLD)
    put(ws, 2, 1, "標記：Verified（可公開驗證、附出處）／Interested-party（寫明是誰、誘因）／Analogy（寫明類比對象）／Assumed（一律給低／高區間）／Derived（寫明換算）。查不到寫「未取得」。")
    for c, h in enumerate(["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "出處", "來源誘因", "說明"], 1):
        put(ws, 3, c, h, BOLD, fill=HDR)
    r = 4
    inrow = {}
    for row in IN:
        if row[0] == "§":
            put(ws, r, 1, row[1], BOLD, fill=SEC); r += 1; continue
        k, lab, unit, lo, b, hi, tag, src, inc, note = row
        put(ws, r, 1, k); put(ws, r, 2, lab); put(ws, r, 3, unit)
        for c, v in ((4, lo), (5, b), (6, hi)):
            if v is not None: put(ws, r, c, v, BLUE)
        put(ws, r, 7, tag); put(ws, r, 8, src); put(ws, r, 9, inc); put(ws, r, 10, note)
        name(wb, k, f"In!$E${r}"); inrow[k] = r; r += 1
    for pre, (lab, unit, a, b, c_, note) in SCN_IN.items():
        for s, v in zip(SCN, (a, b, c_)):
            k = f"{pre}_{s}"
            put(ws, r, 1, k); put(ws, r, 2, f"情境 {s}：{lab}"); put(ws, r, 3, unit); put(ws, r, 5, v, BLUE)
            put(ws, r, 7, "Assumed"); put(ws, r, 8, "—"); put(ws, r, 10, note)
            name(wb, k, f"In!$E${r}"); inrow[k] = r; r += 1
    for s in SCN:
        k = f"scn_w{s}"
        put(ws, r, 1, k); put(ws, r, 2, f"情境權重 {s}"); put(ws, r, 3, "%"); put(ws, r, 5, SCN_W[s], BLUE)
        put(ws, r, 7, "Assumed"); put(ws, r, 8, "DCF 標準規格：預設 25／50／25"); name(wb, k, f"In!$E${r}"); inrow[k] = r; r += 1
    put(ws, r, 1, f"E. 線 1a v0.2 讀入（Derived；建置程式從 {LINE['file']} 讀值，不手打；情境 A＝線高滲透＋低成本＋低價、B＝基準、C＝低滲透＋高成本＋高價）", BOLD, fill=SEC); r += 1
    for k, lab, unit, v, src in line_rows():
        put(ws, r, 1, k); put(ws, r, 2, lab); put(ws, r, 3, unit); put(ws, r, 5, v, BLUE)
        put(ws, r, 7, "Derived"); put(ws, r, 8, src); put(ws, r, 10, "線權威（規則第 13 條）；改值請改線")
        name(wb, k, f"In!$E${r}"); inrow[k] = r; r += 1
    put(ws, r, 1, "stub"); put(ws, r, 2, "評價日（2026-11-16）至 FY2026 年底的年數"); put(ws, r, 3, "年")
    put(ws, r, 5, "=(DATE(2026,12,31)-DATE(2026,11,16))/365"); put(ws, r, 7, "Derived"); name(wb, "stub", f"In!$E${r}"); r += 2
    put(ws, r, 1, "情境驅動對照（純展示；修改請改上方藍字）", BOLD, fill=SEC); r += 1
    for c, h in ((1, "鍵"), (2, "項目"), (4, "情境 A（樂觀）"), (5, "情境 B（基準）"), (6, "情境 C（保守）"), (8, "說明")):
        put(ws, r, c, h, BOLD, fill=HDR)
    r += 1
    for pre, (lab, *_rest) in SCN_IN.items():
        put(ws, r, 1, pre); put(ws, r, 2, lab)
        for c, s in zip((4, 5, 6), SCN):
            put(ws, r, c, f"=E{inrow[f'{pre}_{s}']}")
        put(ws, r, 8, f"引用 In 頁 {pre}_A／{pre}_B／{pre}_C；全部 Assumed"); r += 1
    for kind, pre, lab, unit in LINE_KEYS:
        for y in YEARS:
            put(ws, r, 1, f"{pre}_{y}"); put(ws, r, 2, f"{lab} {y}")
            for c, s_ in zip((4, 5, 6), SCN):
                put(ws, r, c, f"=E{inrow[f'{pre}_{y}_{s_}']}")
            put(ws, r, 8, f"引用 In 頁 {pre}_{y}_A／{pre}_{y}_B／{pre}_{y}_C；讀線（E 節）"); r += 1
    put(ws, r, 1, "scn_w"); put(ws, r, 2, "情境權重")
    for c, s in zip((4, 5, 6), SCN):
        put(ws, r, c, f"=E{inrow['scn_w' + s]}")
    put(ws, r, 8, "引用 In 頁 scn_wA／scn_wB／scn_wC；合計＝1")
    for c, w in zip("ABCDEFGHIJ", (12, 46, 10, 10, 10, 10, 14, 60, 24, 50)):
        ws.column_dimensions[c].width = w

    # ---------------- 市場天花板與情境 ----------------
    ws = W["市場天花板與情境"]
    put(ws, 1, 1, "市場天花板與情境 — 美國叫車總趟數 × Robotaxi 滲透 × 誘發需求 × Waymo 份額 ×（1＋國際）→ Waymo 趟數 × 每趟英里 → 付費里程、車隊。趟數、滲透、份額、價格讀線 1a v0.2（In 頁 E 節）", BOLD)
    put(ws, 2, 1, "2026 起點滲透率（線校準值，情境 B）"); put(ws, 2, 3, "=lpen_2026_B", fmt="0.00%")
    name(wb, "pen26", "'市場天花板與情境'!$C$2")
    put(ws, 3, 1, "天花板讀法：叫車市場只是美國出行的一小部分；AAA 2025 自有車每英里 0.77 美元（小轎車 0.56），Robotaxi 價格若逼近此線，可替代的不只叫車而是自有車里程——本版兩條路徑（與 tesla_robotaxi 同式）：(1) 誘發倍數＝(Waymo 2026 價格 ÷ 當年價格)^ε（ε 讀線 P_ELAST）；(2) 價格低於 自有車成本×便利溢價 時，閘門開啟，部分小客車 VMT 轉入 Robotaxi（c35）。單位換算：轉入英里 ÷ 每趟英里＝趟數。v0.2 份額讀線：Waymo＋Tesla＋其他＝100%（對帳 2）。", wrap=True)
    VROWS = ["us", "pen", "rpm", "jev", "gate", "conv", "av", "wsh", "wus", "intl", "trips", "wk", "shus", "L", "pm", "fleet"]
    VLAB = {"us": ("美國叫車基底趟數（線）", "十億趟"), "pen": ("Robotaxi 滲透率（人類叫車替代，線）", "%"),
            "rpm": ("Waymo 每付費英里實收價格（線）", "US$/英里"), "jev": ("誘發倍數 (2026 價÷當年價)^ε", "倍"), "L": ("Waymo 每趟付費英里", "英里/趟"),
            "gate": ("自有車轉換閘門（價格 vs 自有車成本×便利溢價）", "0–1"), "conv": ("自有車里程轉入 Robotaxi", "十億英里"),
            "av": ("美國 Robotaxi 趟數", "十億趟"), "wsh": ("Waymo 占 Robotaxi 份額（線）", "%"), "wus": ("Waymo 美國趟數", "百萬趟"),
            "intl": ("國際 ÷ 美國", "%"), "trips": ("Waymo 全球付費趟數", "百萬趟"), "wk": ("每週付費趟數", "百萬趟/週"),
            "shus": ("Waymo 美國趟數 ÷ 美國叫車基底", "%"), "fleet": ("所需車隊", "千台"), "pm": ("付費（載客）英里", "百萬英里")}
    VR = {}  # (s,key)->row
    r = 5
    for s in SCN:
        put(ws, r, 1, SCN_NAME[s], BOLD, fill=SEC)
        for y in YEARS: put(ws, r, YEARS.index(y) + 3, y, BOLD, fill=SEC)
        r += 1
        for k in VROWS:
            VR[(s, k)] = r; put(ws, r, 1, VLAB[k][0]); put(ws, r, 2, VLAB[k][1]); r += 1
        r += 1
    for s in SCN:
        R = lambda k: VR[(s, k)]
        for y in YEARS:
            c = YC[y]; n = y - 2026; col = YEARS.index(y) + 3
            prev = YC.get(y - 1)
            f = {
                "us": f"=lt_{y}_{s}/1000",
                "pen": f"=lpen_{y}_{s}",
                "rpm": f"=lp_{y}_{s}",
                "jev": f"=(lp_2026_{s}/{c}{R('rpm')})^eps_{s}",
                "L": f"=tl0+(tl_end-tl0)*{n}/9",
                "gate": f"=MIN(1,MAX(0,(1.5*p_own*cf-{c}{R('rpm')})/(0.5*p_own*cf)))",
                "conv": f"=vmt*ldv_sh*c35_{s}*{n}/9*{c}{R('gate')}",
                "av": f"={c}{R('us')}*{c}{R('pen')}*{c}{R('jev')}+{c}{R('conv')}/{c}{R('L')}",
                "wsh": f"=lws_{y}_{s}",
                "wus": f"={c}{R('av')}*{c}{R('wsh')}*1000",
                "intl": f"=intl35_{s}*{n}/9",
                "trips": f"={c}{R('wus')}*(1+{c}{R('intl')})",
                "wk": f"={c}{R('trips')}/52",
                "shus": f"={c}{R('wus')}/({c}{R('us')}*1000)",
                "fleet": f"={c}{R('pm')}*1000/(lm_{y}_{s}*util_k*365)",
                "pm": f"={c}{R('trips')}*{c}{R('L')}",
            }
            for k in VROWS:
                fmt = "0.0%" if k in ("pen", "wsh", "intl", "shus") else ("0.00" if k in ("rpm", "jev", "gate", "us", "av", "wk", "L") else "#,##0.0")
                put(ws, R(k), col, f[k], fmt=fmt)
    ws.column_dimensions["A"].width = 40; ws.column_dimensions["B"].width = 12

    # ---------------- 每英里經濟 ----------------
    ws = W["每英里經濟"]
    put(ws, 1, 1, "每英里經濟 — 價格 vs 每英里成本（US$／付費英里）；收支只作估值的一環。v0.2 口徑：營運成本＝線 Waymo 第 3–8 項（每付費英里）＋車輛折舊 ÷ 每車年付費英里；全載成本＝營運成本＋固定費用 ÷ 付費英里", BOLD)
    put(ws, 3, 1, "線 Waymo 單位營運成本第 3–8 項（能源、維修清潔、保險理賠、遠端營運、安全員、場站）與稼動率（讀 In 頁 E 節）", BOLD, fill=SEC)
    for c, h in enumerate(["項目", "年", "情境 A（低成本）", "情境 B（基準）", "情境 C（高成本）"], 1): put(ws, 4, c, h, BOLD, fill=HDR)
    r = 5
    for pre, lab in (("lx", "營運成本第 3–8 項（US$/英里）"), ("lm", "每車每日付費英里")):
        for y in (2026, 2030, 2032):
            put(ws, r, 1, lab); put(ws, r, 2, y)
            for c, s_ in zip((3, 4, 5), SCN): put(ws, r, c, f"={pre}_{y}_{s_}", fmt="0.00" if pre == "lx" else "0")
            r += 1
    put(ws, 11, 1, "錨（每英里）", BOLD, fill=SEC)
    anchors = [("人類叫車乘客實付（含小費，全美）ph", "=ph", "Derived：Gridwise 2026 Q3"),
               ("AAA 2025 自有車全成本（15,000 英里／年，平均）", "=11577/15000", "Verified：AAA Your Driving Costs 2025-09-16（年 11,577 美元）"),
               ("AAA 2025 小轎車", 0.5587, "Verified：同上，55.87 美分"),
               ("Tesla 奧斯汀牌價（每英里部分，不含起跳 3.00；2026-03-12 起）", 1.40, "Verified（媒體轉述 app 價格，線 1a v0.2 更正）"),
               ("Tesla 營運成本目標", 0.30, "Interested-party：Tesla 管理層（線 1a）")]
    r = 12
    for lab, v, src in anchors:
        put(ws, r, 1, lab); put(ws, r, 3, v, BLUE if not isinstance(v, str) else None, fmt="0.00"); put(ws, r, 4, src); r += 1
    MROWS = ["rpm", "cpm", "vc", "mpv", "dep_mi", "opc", "fix", "fix_mi", "full", "gap", "flag", "vsaaa"]
    MLAB = {"rpm": ("每付費英里實收價格", "US$/英里"),
            "cpm": ("營運成本第 3–8 項（線，每付費英里）", "US$/英里"), "vc": ("單車成本", "千美元"), "mpv": ("每車年付費英里（線稼動率 × util_k）", "千英里"),
            "dep_mi": ("每付費英里折舊", "US$/英里"), "opc": ("營運成本（每付費英里，含折舊）", "US$/英里"),
            "fix": ("固定費用", "十億美元"), "fix_mi": ("固定費用 ÷ 付費英里", "US$/英里"), "full": ("全載成本（每付費英里）", "US$/英里"),
            "gap": ("價格 − 營運成本", "US$/英里"), "flag": ("補貼訊號（公式判定）", ""), "vsaaa": ("價格 ÷ AAA 自有車每英里", "倍")}
    MR = {}
    r = 19
    for s in SCN:
        put(ws, r, 1, SCN_NAME[s], BOLD, fill=SEC)
        for y in YEARS: put(ws, r, YEARS.index(y) + 3, y, BOLD, fill=SEC)
        r += 1
        for k in MROWS:
            MR[(s, k)] = r; put(ws, r, 1, MLAB[k][0]); put(ws, r, 2, MLAB[k][1]); r += 1
        r += 1
    V = "'市場天花板與情境'!"
    for s in SCN:
        R = lambda k: MR[(s, k)]
        for y in YEARS:
            c = YC[y]; n = y - 2026; col = YEARS.index(y) + 3
            f = {
                "rpm": f"={V}{c}{VR[(s, 'rpm')]}",
                "cpm": f"=lx_{y}_{s}",
                "vc": f"=vcost26*(1-vcost_dec)^{n}",
                "mpv": f"=lm_{y}_{s}*util_k*365/1000",
                "dep_mi": f"={c}{R('vc')}/vlife/{c}{R('mpv')}",
                "opc": f"={c}{R('cpm')}+{c}{R('dep_mi')}",
                "fix": f"=fix26*(1+fix_g)^{n}",
                "fix_mi": f"={c}{R('fix')}*1000/{V}{c}{VR[(s, 'pm')]}",
                "full": f"={c}{R('opc')}+{c}{R('fix_mi')}",
                "gap": f"={c}{R('rpm')}-{c}{R('opc')}",
                "flag": f'=IF({c}{R("rpm")}<{c}{R("opc")},"低於營運成本：補貼",IF({c}{R("rpm")}<{c}{R("full")},"高於營運、低於全載","覆蓋全載成本"))',
                "vsaaa": f"={c}{R('rpm')}/(11577/15000)",
            }
            for k in MROWS:
                fmt = "0.00"
                put(ws, R(k), col, f[k], fmt=fmt)
    ws.column_dimensions["A"].width = 44; ws.column_dimensions["D"].width = 12

    # ---------------- 估值 ----------------
    ws = W["估值"]
    put(ws, 1, 1, "估值 — 分部 DCF（FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加；明確期 FY2026 殘期～FY2035；永續 NOPAT(2036)×(1−g/RONIC)/(WACC−g)，RONIC＝WACC）。十億美元。評等：未上市。", BOLD)
    put(ws, 2, 1, "營收分部：Waymo 未上市、無報告分部；營收列＝「Robotaxi 乘車營收（美國＋國際）」一列（直營 app 與 Uber 平台城市的拆分未揭露）。")
    DROWS = ["rev", "var", "fix", "dep", "ebit", "tax", "nopat", "capex", "dnwc", "fcff", "frac", "t", "pv"]
    DLAB = {"rev": "Robotaxi 乘車營收（美國＋國際）", "var": "變動營運成本（付費英里 × 線第 3–8 項）", "fix": "固定費用",
            "dep": "折舊（車隊 × 單車成本 ÷ 壽命）", "ebit": "營業利益 EBIT", "tax": "稅（只在 EBIT>0）", "nopat": "NOPAT",
            "capex": "資本支出（新增車＋汰換）", "dnwc": "營運資金增加", "fcff": "FCFF", "frac": "計入比例（2026 殘期）",
            "t": "折現期（期中）", "pv": "FCFF 現值"}
    DR = {}; SUMR = {}
    r = 4
    for s in SCN:
        put(ws, r, 1, SCN_NAME[s], BOLD, fill=SEC)
        for y in YEARS: put(ws, r, YEARS.index(y) + 3, y, BOLD, fill=SEC)
        r += 1
        for k in DROWS:
            DR[(s, k)] = r; put(ws, r, 1, DLAB[k]); r += 1
        for k, lab in (("sum_pv", "明確期現值合計"), ("n36", "NOPAT 2036"), ("tv", "永續價值（2035 年底）"), ("pvtv", "永續價值現值"),
                       ("ev", "企業價值 EV"), ("tvsh", "永續價值占 EV（EV≤0 時 N/M）"), ("eq", "股權價值（EV＋淨現金）")):
            SUMR[(s, k)] = r; put(ws, r, 1, lab, BOLD if k in ("ev", "eq") else None); r += 1
        r += 1
    M = "'每英里經濟'!"
    for s in SCN:
        R = lambda k: DR[(s, k)]
        for y in YEARS:
            c = YC[y]; col = YEARS.index(y) + 3; p_ = YC.get(y - 1)
            fleet = f"{V}{c}{VR[(s, 'fleet')]}"
            pfleet = f"{V}{p_}{VR[(s, 'fleet')]}" if p_ else fleet
            f = {
                "rev": f"={V}{c}{VR[(s, 'pm')]}*{M}{c}{MR[(s, 'rpm')]}/1000",
                "var": f"={V}{c}{VR[(s, 'pm')]}*{M}{c}{MR[(s, 'cpm')]}/1000",
                "fix": f"={M}{c}{MR[(s, 'fix')]}",
                "dep": f"={fleet}*{M}{c}{MR[(s, 'vc')]}/vlife/1000",
                "ebit": f"={c}{R('rev')}-{c}{R('var')}-{c}{R('fix')}-{c}{R('dep')}",
                "tax": f"=MAX(0,{c}{R('ebit')})*tax",
                "nopat": f"={c}{R('ebit')}-{c}{R('tax')}",
                "capex": f"=(MAX(0,{fleet}-{pfleet})+{pfleet}/vlife)*{M}{c}{MR[(s, 'vc')]}/1000",
                "dnwc": (f"=nwc*({c}{R('rev')}-{p_}{R('rev')})" if p_ else "=0"),
                "fcff": f"={c}{R('nopat')}+{c}{R('dep')}-{c}{R('capex')}-{c}{R('dnwc')}",
                "frac": ("=stub" if y == 2026 else "=1"),
                "t": ("=stub/2" if y == 2026 else f"=stub+({y}-2027)+0.5"),
                "pv": f"={c}{R('fcff')}*{c}{R('frac')}/(1+wacc)^{c}{R('t')}",
            }
            for k in DROWS:
                put(ws, R(k), col, f[k], fmt="0.000" if k in ("frac", "t") else "#,##0.00")
        S_ = lambda k: SUMR[(s, k)]
        put(ws, S_("sum_pv"), 3, f"=SUM(C{R('pv')}:L{R('pv')})", fmt="#,##0.0")
        put(ws, S_("n36"), 3, f"=L{R('nopat')}*(1+g_tv)", fmt="#,##0.00")
        put(ws, S_("tv"), 3, f"=C{S_('n36')}*(1-g_tv/wacc)/(wacc-g_tv)", fmt="#,##0.0")
        put(ws, S_("pvtv"), 3, f"=C{S_('tv')}/(1+wacc)^(stub+9)", fmt="#,##0.0")
        put(ws, S_("ev"), 3, f"=C{S_('sum_pv')}+C{S_('pvtv')}", fmt="#,##0.0")
        put(ws, S_("tvsh"), 3, f'=IF(C{S_("ev")}>0,C{S_("pvtv")}/C{S_("ev")},"N/M")', fmt="0.0%")
        put(ws, S_("eq"), 3, f"=C{S_('ev')}+netcash", fmt="#,##0.0")
        name(wb, f"eq_{s}", f"'估值'!$C${S_('eq')}")
    r += 1
    put(ws, r, 1, "彙總（十億美元）", BOLD, fill=SEC); r += 1
    SUMROW = r
    for c, h in enumerate(["項目", "", "情境 A", "情境 B", "情境 C", "加權", "說明"], 1): put(ws, r, c, h, BOLD, fill=HDR)
    r += 1
    put(ws, r, 1, "股權價值（分部 DCF）"); put(ws, r, 3, "=eq_A", fmt="#,##0.0"); put(ws, r, 4, "=eq_B", fmt="#,##0.0"); put(ws, r, 5, "=eq_C", fmt="#,##0.0")
    put(ws, r, 6, "=scn_wA*eq_A+scn_wB*eq_B+scn_wC*eq_C", fmt="#,##0.0"); put(ws, r, 7, "權重 25／50／25"); name(wb, "eq_w", f"'估值'!$F${r}"); EQR = r; r += 1
    put(ws, r, 1, "估值區間下緣（全部成員含加權取 MIN）"); put(ws, r, 6, f"=MIN(C{EQR}:F{EQR})", fmt="#,##0.0"); name(wb, "eq_lo", f"'估值'!$F${r}"); r += 1
    put(ws, r, 1, "估值區間上緣（MAX）"); put(ws, r, 6, f"=MAX(C{EQR}:F{EQR})", fmt="#,##0.0"); name(wb, "eq_hi", f"'估值'!$F${r}"); r += 1
    put(ws, r, 1, "檢查：加權落在區間內"); put(ws, r, 6, '=IF(AND(eq_w>=eq_lo,eq_w<=eq_hi),"OK","FAIL")'); r += 1
    put(ws, r, 1, "最近一輪融資估值（投後，2026-02）"); put(ws, r, 6, "=val_round", fmt="#,##0.0"); put(ws, r, 7, "Interested-party；只作對照，不進模型"); r += 1
    put(ws, r, 1, "加權 DCF ÷ 融資估值"); put(ws, r, 6, "=eq_w/val_round", fmt="0.0%"); r += 1
    put(ws, r, 1, "情境 A ÷ 融資估值"); put(ws, r, 6, "=eq_A/val_round", fmt="0.0%"); r += 1
    put(ws, r, 1, "評等"); put(ws, r, 6, "未上市"); r += 1
    put(ws, r, 1, "結論（公式組字）"); put(ws, r, 3, '="在本版假設下，Waymo 分部 DCF 三情境為 "&TEXT(eq_C,"0")&"～"&TEXT(eq_A,"0")&" 十億美元（加權 "&TEXT(eq_w,"0")&"），"&IF(eq_A<val_round,"連情境 A 都低於","情境 A 高於")&" 2026-02 融資投後估值 "&TEXT(val_round,"0")&" 十億美元；融資估值隱含的滲透與價格路徑高於本版情境 A。"', wrap=True)
    name(wb, "concl", f"'估值'!$C${r}"); CONCLR = r; r += 2
    # 反解（整本重算，Python 鏡像）
    put(ws, r, 1, "反解（整本重算；Python 鏡像二分法，數值寫入；建置時核對鏡像＝xlsx）", BOLD, fill=SEC); r += 1
    REVR = r
    ws.column_dimensions["A"].width = 46; ws.column_dimensions["G"].width = 40
    for col in "CDEF": ws.column_dimensions[col].width = 12

    # ---------------- 滲透（主頁） ----------------
    build_penetration(W["滲透"], wb, P, VR)

    # ---------------- 資料來源比較、對帳 ----------------
    build_sources(W["資料來源比較"])
    build_recon(W["對帳"], P)

    # ---------------- Sens（Python 鏡像，數值） ----------------
    ws = W["Sens"]
    base_w = weighted(P)
    put(ws, 1, 1, "Sens — 一次動一個 In 頁輸入至敏感度低／高（D／F 欄），整本重算三情境後加權（Python 鏡像與 xlsx 公式同式；建置時核對基準一致）。單位：十億美元。", BOLD)
    put(ws, 2, 1, "註：tax 低＝基準（21%）只列高；cf 低＝基準（1.0）只列高；val_round 只作對照，不進模型。情境專屬輸入（*_A／_B／_C）與線讀入（In 頁 E 節）無低／高欄，其不確定性由三情境與下方「情境 B 改用 A／C 路徑」案例表達。")
    hdr = 4
    for c, h in enumerate(["驅動", "設定", "鍵", "加權目標價（股權價值）", "相對基準", "情境 B 股權價值"], 1): put(ws, hdr, c, h, BOLD, fill=HDR)
    r = hdr + 1
    put(ws, r, 1, "基準"); put(ws, r, 2, "In 頁 E 欄"); put(ws, r, 4, round(base_w, 4)); put(ws, r, 5, 0, fmt="0.0%"); put(ws, r, 6, round(run(P, "B")["eq"], 4)); r += 1
    sens_rows = []
    for row in IN:
        if row[0] == "§" or (row[3] is None and row[5] is None): continue
        k, lab, unit, lo, b, hi = row[:6]
        for side, v in (("低", lo), ("高", hi)):
            if v is None or v == b: continue
            p2 = {**P, k: v}
            w2 = weighted(p2)
            sens_rows.append((abs(w2 - base_w), lab, f"{k}＝{v}（敏感度{side}）", k, w2, run(p2, "B")["eq"]))
    # 另加：WACC ±1pt、終值替代法（FCFF×(1+g)）
    for d in (-0.01, 0.01):
        p2 = {**P, "wacc": P["wacc"] + d}
        sens_rows.append((abs(weighted(p2) - base_w), f"WACC {'+' if d > 0 else '−'}1pt", f"wacc＝{P['wacc'] + d:.2f}", "wacc", weighted(p2), run(p2, "B")["eq"]))
    # 線路徑：情境 B 改用情境 A／C 的路徑（一次只換一組；鍵為情境專屬，無低／高欄）
    for pre, lab in (("lp", "線價格路徑"), ("lx", "線營運成本路徑（第 3–8 項）"), ("lm", "線稼動率路徑"), ("lws", "線 Waymo 份額路徑"),
                     ("lpen", "線滲透率路徑"), ("lt", "線美國叫車趟數路徑")):
        for frm in ("A", "C"):
            p2 = swap(P, pre, frm)
            sens_rows.append((abs(weighted(p2) - base_w), f"{lab}：情境 B 改用情境 {frm}", f"{pre}_<年>_B ← {pre}_<年>_{frm}", f"{pre}_*_B", weighted(p2), run(p2, "B")["eq"]))
    p2 = {**P, "scn_wA": 0, "scn_wB": 1, "scn_wC": 0}
    sens_rows.append((abs(run(P, "B")["eq"] - base_w), "情境權重只用情境 B", "權重 0／100／0", "scn_w*", run(P, "B")["eq"], run(P, "B")["eq"]))
    sens_rows.sort(key=lambda x: -x[0])
    for _, lab, setting, k, w2, b2 in sens_rows:
        put(ws, r, 1, lab); put(ws, r, 2, setting); put(ws, r, 3, k); put(ws, r, 4, round(w2, 4), fmt="#,##0.0")
        put(ws, r, 5, f"=D{r}-D{hdr + 1}", fmt="#,##0.0"); put(ws, r, 6, round(b2, 4), fmt="#,##0.0"); r += 1
    r += 1
    put(ws, r, 1, "檢查：Sens 基準（Python）− 估值頁加權（xlsx）"); put(ws, r, 4, f"=D{hdr + 1}-eq_w", fmt="0.0000"); SENS_CHK = r; r += 1
    fp = hashlib.sha1(json.dumps({k: P[k] for k in sorted(P) if isinstance(P[k], (int, float))}, sort_keys=True).encode()).hexdigest()[:10]
    build_id = f"waymo-{VERSION}-{fp}"
    put(ws, r, 1, "build_id"); put(ws, r, 4, build_id); r += 1
    put(ws, r, 1, "模型狀態"); put(ws, r, 4, f'=IF(ABS(D{SENS_CHK})<0.001,"CURRENT","STALE")'); r += 1
    ws.column_dimensions["A"].width = 44; ws.column_dimensions["B"].width = 34; ws.column_dimensions["D"].width = 22

    # 反解寫回估值頁
    ws = W["估值"]
    r = REVR
    rev_items = []
    rev_items.append(("情境 B：Waymo 份額改用線低滲透情境路徑（2035 55%）時的股權價值", run(swap(P, "lws", "C"), "B")["eq"], "#,##0.0", "十億美元；份額權威在線（對帳 2 判斷題 1 的替代讀法）"))
    rev_items.append(("情境 A 股權價值＝融資估值所需的 WACC", solve(P, "wacc", P["val_round"], 0.02, 0.30, "A"), "0.0%", "g＝3%；基準 WACC 12%"))
    rev_items.append(("情境 B 股權價值＝0 所需的 2035 每趟付費英里（tl_end）", solve(P, "tl_end", 0.0, 1.0, 12.0, "B"), "0.00", "英里/趟；2026 為 CPUC 實測 3.73"))
    rev_items.append(("情境 B 股權價值＝0 的 2026 固定費用（fix26）", solve(P, "fix26", 0.0, 0.0, 30.0, "B"), "0.00", "十億美元／年，2% 成長"))
    rev_items.append(("加權股權價值＝融資估值所需的 WACC", solve(P, "wacc", P["val_round"], 0.02, 0.30, wt=True), "0.0%", "三情境 25／50／25"))
    for lab, v, fmt, note in rev_items:
        put(ws, r, 1, lab); put(ws, r, 6, v if v is not None else "區間內無解", fmt=fmt); put(ws, r, 7, note); r += 1

    # ---------------- README 頁 ----------------
    ws = wsR
    put(ws, 1, 1, f"Waymo（Alphabet 分部，視同未上市）個股模型 {VERSION}｜build_id {build_id}｜資料截止日 {AS_OF}｜成品日 {AS_OF}｜股價基準日：不適用（未上市）", BOLD)
    lines = [
        "命題：Waymo 的價值取決於 Robotaxi 在美國（及全球）叫車的滲透速度，以及每英里實收價格能降到哪裡（決定傑文斯效應能否實現）；收支只是估值的一環。",
        "v0.2（對帳 2）：美國叫車趟數、滲透率、Waymo 份額、價格、營運成本（第 3–8 項）、稼動率改讀線 1a v0.2（In 頁 E 節）；與 tesla_robotaxi 共用同一組份額（Waymo＋Tesla＋其他＝100%）；每趟英里由 CPUC 實測 3.73 收斂到全美叫車平均 6.0（2035）；誘發倍數改為相對 2026 價格、ε 讀線；便利溢價 cf 1.0。",
        "主頁＝「滲透」：最上方「價格」區塊（Waymo 每英里實收價格時間序列與人類叫車、自有車兩個錨）、全國里程碑、各城市、加州 CPUC 月度、全國份額。",
        "頁面（與 physical_ai/tesla_robotaxi/ 同一套）：滲透｜市場天花板與情境｜每英里經濟｜估值｜資料來源比較｜對帳｜In｜Sens。",
        "估值口徑：分部 DCF（FCFF，WACC 12%，g 3%，RONIC＝WACC）三情境 25／50／25（A＝線高滲透、B＝線基準、C＝線低滲透）；以 2026-02 融資投後 1,260 億美元為對照；同業倍數腿未建（理由見 README.md）。評等：未上市。",
        "不讀 Siliconomics／Tokenomics／Wattonomics 協調檔（實體 AI 線不適用）；車用晶片、LiDAR 單價屬 AI供應鏈／共用零組件層，只讀未用。",
    ]
    for i, t in enumerate(lines, 3): put(ws, i, 1, t, wrap=True)
    put(ws, 9, 1, "結論（引用估值頁）：", BOLD); put(ws, 10, 1, "=concl", wrap=True)
    put(ws, 12, 1, "加權股權價值（十億美元）"); put(ws, 12, 2, "=eq_w", fmt="#,##0.0")
    put(ws, 13, 1, "區間（MIN～MAX）"); put(ws, 13, 2, '=TEXT(eq_lo,"0.0")&"～"&TEXT(eq_hi,"0.0")')
    put(ws, 14, 1, "融資估值（對照）"); put(ws, 14, 2, "=val_round")
    put(ws, 15, 1, "評等"); put(ws, 15, 2, "未上市")
    ws.column_dimensions["A"].width = 140

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)
    return P, build_id


# ---------------------------------------------------------------------------
# 滲透頁
# ---------------------------------------------------------------------------
OBI = [  # 期間, Waymo 均價, Uber, Lyft, Tesla, 樣本, 出處, 標記
    ("2025-03-25～04-25", 20.43, 15.58, 14.44, None, "近 9 萬筆報價（舊金山）", "Obi，TechCrunch 2025-06-12", "Verified"),
    ("2025-11-27～2026-01-01", 19.69, 17.47, 15.47, None, "9.4 萬筆模擬叫車（灣區）", "Obi，TechCrunch 2026-01-27；Automotive World", "Verified"),
    ("2026-03-01～04-14", "=C9*1.18", "=D9*1.05", "=E9*1.12", None, "Obi 春季灣區研究（二手）", "RideWise 2026-05-25 轉述 Obi：Waymo +18%、Uber +5%、Lyft +12%（相對 2026-01）；原文未取得", "Derived"),
]


def build_penetration(ws, wb, P, VR):
    put(ws, 1, 1, "滲透（主頁）— Waymo｜價格 → 量 → 份額。資料截止 2026-11-16。標記見各列；查不到寫「未取得」。", BOLD)
    put(ws, 2, 1, "核心指標：每英里實收價格（專案規格 2026-11-16：決定傑文斯效應能否實現）。口徑：乘客實付（含起跳價與費用，不含小費）÷ 載客英里；人類叫車錨＝UberX／Lyft 乘客實付 ×（1＋小費率）；自有車錨＝AAA 每車英里全成本。", wrap=True)
    put(ws, 4, 1, "A1. 價格時間序列（舊金山，同路線比價；Obi）", BOLD, fill=SEC)
    put(ws, 5, 1, "小費率（人類叫車）"); put(ws, 5, 3, 0.08, BLUE, fmt="0%"); put(ws, 5, 4, "Assumed（5–15%）；與 tesla_robotaxi tip 同值（RideWise 以 15% 計，為上緣）"); name(wb, "tip", "'滲透'!$C$5")
    put(ws, 6, 1, "每趟載客英里（換算用）"); put(ws, 6, 3, 3.4, BLUE, fmt="0.0"); put(ws, 6, 4, "Derived：CPUC 加州每趟 3.4 英里（Findings Press）；Obi 同路線比價，故各家同分母"); name(wb, "mi_trip_sf", "'滲透'!$C$6")
    hdrs = ["期間", "樣本", "Waymo 每趟", "UberX 每趟", "Lyft 每趟", "UberX＋小費", "Lyft＋小費", "Waymo 每英里", "UberX＋小費 每英里", "Lyft＋小費 每英里",
            "Waymo ÷ UberX＋小費 −1", "Waymo ÷ AAA 自有車", "出處", "標記"]
    for c, h in enumerate(hdrs, 1): put(ws, 7, c, h, BOLD, fill=HDR)
    r = 8
    for per, w, u, l, t, smp, src, tag in OBI:
        put(ws, r, 1, per); put(ws, r, 2, smp)
        put(ws, r, 3, w, BLUE if not isinstance(w, str) else None, fmt="0.00"); put(ws, r, 4, u, BLUE if not isinstance(u, str) else None, fmt="0.00"); put(ws, r, 5, l, BLUE if not isinstance(l, str) else None, fmt="0.00")
        put(ws, r, 6, f"=D{r}*(1+tip)", fmt="0.00"); put(ws, r, 7, f"=E{r}*(1+tip)", fmt="0.00")
        put(ws, r, 8, f"=C{r}/mi_trip_sf", fmt="0.00"); put(ws, r, 9, f"=F{r}/mi_trip_sf", fmt="0.00"); put(ws, r, 10, f"=G{r}/mi_trip_sf", fmt="0.00")
        put(ws, r, 11, f"=C{r}/F{r}-1", fmt="0.0%"); put(ws, r, 12, f"=H{r}/(11577/15000)", fmt="0.0")
        put(ws, r, 13, src); put(ws, r, 14, tag); r += 1
    put(ws, r, 1, "讀法（公式組字）"); put(ws, r, 3, '="2026-01 舊金山 Waymo 每趟 "&TEXT(C9,"0.00")&" 美元，"&IF(C9<F9,"低於","高於")&"含小費的 UberX（"&TEXT(F9,"0.00")&"）；不含小費則高 "&TEXT(C9/D9-1,"0.0%")&"。每英里約 "&TEXT(H9,"0.0")&" 美元，是 AAA 自有車每英里的 "&TEXT(L9,"0.0")&" 倍。"', wrap=True); r += 2

    put(ws, r, 1, "A2. 依行程長度（含起跳價）", BOLD, fill=SEC); r += 1
    for c, h in enumerate(["長度區間", "期間", "Waymo 每英里", "UberX 每英里", "Lyft 每英里", "Waymo 溢價 vs UberX", "出處", "標記"], 1): put(ws, r, c, h, BOLD, fill=HDR)
    r += 1
    KM = 1.609344
    rows = [("短（< 1.4 公里，約 0.9 英里）", "2025-03～04", "未取得（Obi 報每公里約 26 美元，疑誤植）", "未取得", "未取得", 0.4148, "Obi（TechCrunch 2025-06-12）：最短行程 Waymo 比 Uber 貴 41.48%", "Verified"),
            ("中（4.3–9.3 公里，2.7–5.8 英里）", "2025-03～04", f"=3.5*{KM}", f"=2.9*{KM}", f"=2.6*{KM}", "=C{r}/D{r}-1", "Obi：每公里 Waymo 3.50、Uber 2.90、Lyft 2.60 美元", "Verified"),
            ("中（4.3–9.3 公里）", "2025-11～2026-01", "未取得（只揭露相對值）", "未取得", "未取得", 0.02, "Obi（Automotive World 2026-01-28）：Waymo 每公里約比 Uber 高 2%", "Verified"),
            ("長（> 9.3 公里）", "—", "未取得", "未取得", "未取得", "未取得", "Obi 未揭露長程區間；RideWise 費率表推算的損益兩平距離與其自身範例矛盾，不採用", "—")]
    for lab, per, w, u, l, prem, src, tag in rows:
        put(ws, r, 1, lab); put(ws, r, 2, per)
        for c, v in ((3, w), (4, u), (5, l)):
            put(ws, r, c, v, fmt="0.00")
        prem = prem.format(r=r) if isinstance(prem, str) and "{r}" in prem else prem
        put(ws, r, 6, prem, BLUE if isinstance(prem, float) else None, fmt="0.0%"); put(ws, r, 7, src); put(ws, r, 8, tag); r += 1
    put(ws, r, 1, "牌價推算（RideWise 費率表；可靠度低，只作形狀參考）：每英里＝(起跳＋每英里×d＋每分鐘×d×每英里分鐘)÷d", wrap=True); r += 1
    put(ws, r, 1, "每英里分鐘（假設平均時速 15 英里）"); put(ws, r, 3, 4.0, BLUE); put(ws, r, 4, "Assumed（3–5）"); MPM = r; r += 1
    for c, h in enumerate(["城市／服務", "起跳", "每英里", "每分鐘", "2 英里每英里", "5 英里每英里", "10 英里每英里", "出處", "標記"], 1): put(ws, r, c, h, BOLD, fill=HDR)
    r += 1
    cards = [("舊金山 Waymo", 4.95, 3.65, 0.42), ("舊金山 UberX（未含服務費）", 1.00, 1.25, None), ("鳳凰城 Waymo", 2.50, 2.10, 0.34),
             ("洛杉磯 Waymo", 3.95, 2.95, 0.38), ("奧斯汀 Waymo（Uber app）", 1.00, 0.93, 0.18)]
    for lab, b, pm, pmin in cards:
        put(ws, r, 1, lab); put(ws, r, 2, b, BLUE); put(ws, r, 3, pm, BLUE); put(ws, r, 4, pmin if pmin is not None else "未取得", BLUE if pmin else None)
        for c, d in ((5, 2), (6, 5), (7, 10)):
            per_min = f"+D{r}*{d}*$C${MPM}" if pmin is not None else ""
            put(ws, r, c, f"=(B{r}+C{r}*{d}{per_min})/{d}", fmt="0.00")
        put(ws, r, 8, "RideWise 2026-05-25（2026-09-10 更新）：Waymo app 2026-04～05 擷取；Uber 為費率卡"); put(ws, r, 9, "Interested-party"); r += 1
    put(ws, r, 1, "註：Uber 費率卡未含預約費與動態加價，與 Obi 實測均價（17.47 美元／趟）差距大；因此每英里價格以 A1 實測為準。"); r += 2

    put(ws, r, 1, "A3. 錨（每英里）", BOLD, fill=SEC); r += 1
    for c, h in enumerate(["錨", "值", "出處", "標記"], 1): put(ws, r, c, h, BOLD, fill=HDR)
    r += 1
    for lab, v, src, tag in [("人類叫車 UberX＋小費（舊金山 2026-01）", "=I9", "A1 列", "Derived"),
                              ("人類叫車（全美，乘客實付含小費）", "=ph", "In 頁 ph（Gridwise 2026 Q3 推算）", "Derived"),
                              ("自有車全成本，平均（15,000 英里／年）", "=11577/15000", "AAA Your Driving Costs，2025-09-16：年 11,577 美元", "Verified"),
                              ("自有車全成本，小轎車", 0.5587, "AAA 2025：55.87 美分／英里", "Verified"),
                              ("Waymo 全美平均 2026（讀線，情境 B）", "=lp_2026_B", "In 頁 lp_2026_B（線 W_PRICE）", "Derived")]:
        put(ws, r, 1, lab); put(ws, r, 2, v, BLUE if not isinstance(v, str) else None, fmt="0.00"); put(ws, r, 3, src); put(ws, r, 4, tag); r += 1
    r += 1
    put(ws, r, 1, "A4. 觀察到的價格彈性：Δln(付費里程) ÷ Δln(價格)（同城市）", BOLD, fill=SEC); r += 1
    for c, h in enumerate(["區間", "價格變化（舊金山 Obi）", "付費里程變化（加州 CPUC）", "粗算彈性", "判讀"], 1): put(ws, r, c, h, BOLD, fill=HDR)
    r += 1
    E_HDR = r
    r += 3  # 兩列，於 CPUC 表建好後填
    put(ws, r, 1, "結論：同城市、可識別的價格彈性＝未取得。CPUC 只有全州月資料（不分城市），且同期服務區與車隊大幅擴張（供給受限），粗算值不作依據；情境中的 ε 為 Assumed。", wrap=True); r += 2

    # B. 全國里程碑
    put(ws, r, 1, "B. 全國時間序列（里程碑）", BOLD, fill=SEC); r += 1
    for c, h in enumerate(["日期", "每週付費趟數（千）", "營運城市", "車隊（台）", "服務區（平方英里）", "累計自動駕駛英里（百萬）", "其他", "出處", "標記", "來源誘因"], 1): put(ws, r, c, h, BOLD, fill=HDR)
    r += 1
    MS = [("2025-05-05", 250, None, 1500, None, None, "", "Waymo 部落格（robotaxistats.com 彙整）", "Interested-party", "公司自揭 KPI"),
          ("2025 全年", None, None, None, None, None, "全年 1,500 萬趟；累計 2,000 萬＋", "Waymo 部落格 2026-02-02", "Interested-party", "公司自揭 KPI"),
          ("2026-02-02", 400, None, None, None, None, "融資 160 億美元，投後 1,260 億", "Waymo 部落格；AP、Bloomberg", "Interested-party", "募資宣傳"),
          ("2026-03-26", 500, 10, None, None, None, "", "Waymo（TechCrunch 2026-03-27）", "Interested-party", "公司自揭 KPI"),
          ("2026-05-13", None, 11, 3000, 1400, None, "服務區由約 1,100 擴至 1,400＋", "Electrek 2026-05-13（車隊為媒體估計）", "Derived", ""),
          ("2026-09-24", 500, 15, 4000, None, None, "德州登記 1,102 台；加州＋德州約占 80%", "TechCrunch 2026-09-24（Dealroom 轉述）；車隊資料未具名", "Derived", ""),
          ("2026-10-22", None, 14, None, None, 200, "14 城開放（robotaxistats 計，不含拉斯維加斯）", "Waymo FAQ（robotaxistats.com 彙整）", "Interested-party", "公司 FAQ"),
          ("2026-10-28", None, 15, None, None, 300, "Q3 法說：聖安東尼奧、聖地牙哥、坦帕開放", "Alphabet Q3 2026 法說（Pichai）", "Interested-party", "管理層口頭"),
          ("2026 年底（目標）", 1000, 35, None, None, None, "另 20＋城（含倫敦、東京）", "Bloomberg 2026-02-11（共同執行長 Mawakana）；Fortune 2026-10-30", "Interested-party", "管理層目標")]
    for row in MS:
        for c, v in enumerate(row, 1):
            put(ws, r, c, v, BLUE if isinstance(v, (int, float)) else None, fmt="#,##0")
        r += 1
    put(ws, r, 1, "每車每日付費趟數（2026-09）"); put(ws, r, 2, f"=B{r-4}*1000/7/D{r-4}", fmt="0.0"); put(ws, r, 3, "Derived：每週 50 萬 ÷ 7 ÷ 約 4,000 台"); r += 2

    # C. 各城市
    put(ws, r, 1, "C. 各城市（2026-10 前後）", BOLD, fill=SEC); r += 1
    for c, h in enumerate(["城市", "狀態（2026-10-22）", "叫車 app", "服務區（平方英里，2026-04 估）", "車隊（台）", "叫車市場份額（總訂單金額）", "Waymo 等候（分）", "UberX／Lyft 等候（分）", "Waymo 牌價每英里（RideWise）", "出處"], 1): put(ws, r, c, h, BOLD, fill=HDR)
    r += 1
    cities = [("舊金山灣區", "無人營運", "Waymo", 260, 1000, 0.15, 5.74, "3.15／5.14", 3.65, "服務區 Electrek；車隊 Obi（2026-01，約 1,000）；份額 YipitData 2026-06（Business Insider 轉述）；等候 Obi"),
              ("洛杉磯", "無人營運", "Waymo", 120, "未取得", 0.15, "未取得", "未取得", 2.95, "份額 YipitData 2026-06"),
              ("鳳凰城", "無人營運", "Waymo", 315, "未取得", 0.16, "未取得", "未取得", 2.10, "份額 YipitData 2026-06"),
              ("奧斯汀", "無人營運", "Uber", "90–130", "德州合計 1,102", "未取得", "未取得", "未取得", 0.93, "德州登記（TechCrunch 2026-09-24）"),
              ("亞特蘭大", "無人營運", "Uber", 65, "未取得", "未取得", "未取得", "未取得", 0.85, ""),
              ("達拉斯", "無人營運", "Waymo", 50, "德州合計", "未取得", "未取得", "未取得", 2.35, ""),
              ("休士頓", "無人營運", "Waymo", 25, "德州合計", "未取得", "未取得", "未取得", 2.30, ""),
              ("聖安東尼奧", "無人營運（2026-10-14 全面開放）", "Waymo", 60, "德州合計", "未取得", "未取得", "未取得", 2.25, ""),
              ("邁阿密", "無人營運", "Waymo", 60, "未取得", "未取得", "未取得", "未取得", 2.75, ""),
              ("奧蘭多", "無人營運", "Waymo", 60, "未取得", "未取得", "未取得", "未取得", 2.50, ""),
              ("坦帕", "無人營運（2026-10-14 全面開放）", "Waymo", "未取得", "未取得", "未取得", "未取得", "未取得", "未取得", ""),
              ("聖地牙哥", "無人營運（2026-10-14 全面開放）", "Waymo", "未取得", "未取得", "未取得", "未取得", "未取得", "未取得", ""),
              ("丹佛", "逐步開放（2026-09-01）", "Waymo", "未取得", "未取得", "未取得", "未取得", "未取得", "未取得", ""),
              ("納許維爾", "無人營運", "Waymo", "未取得", "未取得", "未取得", "未取得", "未取得", "未取得", ""),
              ("拉斯維加斯", "逐步開放（2026-09-14）", "Waymo", "未取得", "未取得", "未取得", "未取得", "未取得", "未取得", "")]
    for row in cities:
        for c, v in enumerate(row, 1):
            put(ws, r, c, v, BLUE if isinstance(v, (int, float)) else None, fmt="0%" if c == 6 and isinstance(v, float) else None)
        r += 1
    put(ws, r, 1, "城市狀態：robotaxistats.com（引 Waymo FAQ／部落格，2026-10-22～27）；服務區：Electrek 2026-05-13 估計（Derived）；牌價：RideWise（Interested-party，可靠度低）；份額：YipitData 估計（第三方卡片資料，Interested-party：Uber CFO 曾引用同系列）。", wrap=True); r += 2

    # D. CPUC
    put(ws, r, 1, "D. 加州 CPUC 月度（Waymo 無人載客服務；Platform Aeronaut 彙整 CPUC 季報；＋號月份有遮蔽，為部分加總）", BOLD, fill=SEC); r += 1
    for c, h in enumerate(["月份", "趟數", "乘客", "總行駛英里", "空駛比", "付費英里（推算）", "每趟付費英里", "每週趟數（千）", "註"], 1): put(ws, r, c, h, BOLD, fill=HDR)
    r += 1
    cpuc = [("2025-01", 550457, 783182, 3481697.93, .5305), ("2025-02", 559569, 779358, 3485921.73, .5007), ("2025-03", 708180, 1023028, 4265981.5, .4708),
            ("2025-04", 776058, 1118391, 4784485.32, .4681), ("2025-05", 875592, 1260448, 5336125.62, .4436), ("2025-06", 583262, 828029, 3567131.71, .433),
            ("2025-07", 778812, 1122471, 4875510.04, .4388), ("2025-08", 880643, 1265004, 5571862.8, .4313), ("2025-09", 1016546, 1369203, 6567950.55, .4427),
            ("2025-10", 1217872, 1600676, 7902019.29, .4389), ("2025-11", 1273473, 1650645, 8414101.98, .4415), ("2025-12", 1257187, 1601574, 8420126.85, .4343),
            ("2026-01", 1322645, 1657530, 9092739.35, .4496), ("2026-02", 1220222, 1530644, 8191719.63, .4259), ("2026-03", 1363067, 1717304, 9105108.29, .416),
            ("2026-04", 1416158, 2207679, 9414525.13, .4071), ("2026-05", 1454991, 2309049, 9704968.28, .4071), ("2026-06", 1348926, 2144678, 9023791.29, .4189)]
    days = {"01": 31, "02": 28, "03": 31, "04": 30, "05": 31, "06": 30, "07": 31, "08": 31, "09": 30, "10": 31, "11": 30, "12": 31}
    crow = {}
    for m, t, ps, mi, dh in cpuc:
        put(ws, r, 1, m); put(ws, r, 2, t, BLUE, fmt="#,##0"); put(ws, r, 3, ps, BLUE, fmt="#,##0"); put(ws, r, 4, mi, BLUE, fmt="#,##0"); put(ws, r, 5, dh, BLUE, fmt="0.0%")
        put(ws, r, 6, f"=D{r}*(1-E{r})", fmt="#,##0"); put(ws, r, 7, f"=F{r}/B{r}", fmt="0.00"); put(ws, r, 8, f"=B{r}/{days[m[-2:]]}*7/1000", fmt="#,##0")
        note = "＋（部分遮蔽）" if m >= "2026" else ("異常低（疑遮蔽），不作趨勢依據" if m == "2025-06" else "")
        put(ws, r, 9, note); crow[m] = r; r += 1
    put(ws, r, 1, "加州 2026-04～06 平均每週趟數 ÷ 全國每週 50 萬"); put(ws, r, 8, f"=AVERAGE(H{crow['2026-04']}:H{crow['2026-06']})/500", fmt="0%"); put(ws, r, 9, "Derived；全國每週 50 萬（2026-03 起）"); r += 2

    # A4 彈性列（回填）
    rr = E_HDR
    for lab, p1, p2, m1, m2 in (("2025-04 → 2026-01", "C8", "C9", "2025-04", "2026-01"), ("2026-01 → 2026-04", "C9", "C10", "2026-01", "2026-04")):
        put(ws, rr, 1, lab); put(ws, rr, 2, f"={p2}/{p1}-1", fmt="0.0%"); put(ws, rr, 3, f"=F{crow[m2]}/F{crow[m1]}-1", fmt="0.0%")
        put(ws, rr, 4, f"=LN(F{crow[m2]}/F{crow[m1]})/LN({p2}/{p1})", fmt="0.0")
        put(ws, rr, 5, "受擴張干擾，不作依據（價格降、量大增：供給驅動）" if "2025" in lab else "受擴張干擾，不作依據（價格升、量仍增：需求未飽和或供給受限）"); rr += 1

    # E. 全國份額
    put(ws, r, 1, "E. 全國份額（Derived）", BOLD, fill=SEC); r += 1
    put(ws, r, 1, "Waymo 2026 年趟數（線，Q2 年化）÷ 美國叫車年趟數"); put(ws, r, 3, "=lpen_2026_B*lws_2026_B", fmt="0.00%"); r += 1
    put(ws, r, 1, "年底目標（每週 100 萬）年化 ÷ 美國叫車年趟數（線 2027）"); put(ws, r, 3, "=1*52/lt_2027_B", fmt="0.00%"); r += 1
    put(ws, r, 1, "情境 B：2030 每週趟數（百萬）"); put(ws, r, 3, f"='市場天花板與情境'!G{VR[('B', 'wk')]}", fmt="0.0"); r += 1
    put(ws, r, 1, "情境 B：2035 每週趟數（百萬）"); put(ws, r, 3, f"='市場天花板與情境'!L{VR[('B', 'wk')]}", fmt="0.0"); r += 1
    ws.column_dimensions["A"].width = 34
    for col in "BCDEFGHIJKL": ws.column_dimensions[col].width = 14
    ws.column_dimensions["M"].width = 50


def build_sources(ws):
    put(ws, 1, 1, "資料來源比較 — 來源｜涵蓋｜頻率｜可靠度｜建議用途｜本模型使用處｜標記", BOLD)
    hd = ["來源", "涵蓋", "頻率", "可靠度", "建議用途", "本模型使用處", "標記", "網址／備註"]
    for c, h in enumerate(hd, 1): put(ws, 2, c, h, BOLD, fill=HDR)
    rows = [
        ("Waymo 部落格、FAQ、新聞稿", "全國每週付費趟數、城市開放、累計英里、融資", "不定期（里程碑）", "高（數字可信但選擇性揭露，只報好消息）", "全國量的錨；年底目標作 Interested-party", "滲透 B；線 1a 份額校準", "Interested-party", "waymo.com/blog；經 robotaxistats.com 彙整"),
        ("Alphabet 法說與 10-Q", "Other Bets 營收與營業虧損、城市數、英里", "季", "高（法定申報），但 Waymo 不單獨揭露", "燒錢上限（fix26）", "In fix26", "Verified", "Q3 2026 逐字稿 2026-10-28"),
        ("加州 CPUC AV 季報（含彙整網站）", "加州月度趟數、乘客、行駛英里、空駛比", "季（月資料）", "高（監管申報），但公開版大量遮蔽；需經彙整網站", "空駛比、每趟英里、加州量；彈性粗算", "滲透 D、線 1a（tl0、稼動率）", "Verified", "avrideshare.platformaeronaut.com；Findings Press 論文"),
        ("加州 DMV", "測試／部署許可、脫離報告", "年／即時", "高，但不含商業量", "許可追蹤", "未用", "Verified", "未取得數值"),
        ("NHTSA SGO 事故資料", "事故報告", "月", "高（強制申報），分母（英里）需另找", "停駛風險追蹤", "未用（風險登錄）", "Verified", "未取得數值"),
        ("德州 DMV 車籍（經媒體）", "德州車隊數", "不定期", "中（媒體整理，未具名資料集）", "車隊規模", "滲透 B、C", "Derived", "TechCrunch 2026-09-24"),
        ("robotaxitracker.com（專案常用）", "Tesla、Waymo、Zoox、Avride：車輛名錄、服務區地圖、乘車統計、價格、等候、政府資料（CPUC、DMV、NHTSA、奧斯汀事件、亞利桑那 MVD）", "持續（社群）", "中低：社群投稿＋等候時間遙測＋政府紀錄＋審核流程；Methodology 頁於本環境無法讀取內文（前端渲染、代理封鎖），只見來源類別；與公司無關聯", "等候與價格的高頻趨勢、車隊名錄交叉核對；不作估值錨", "未直接取數（本環境讀不到數值）", "Interested-party", "社群；robotaxitracker.com/methodology"),
        ("robotaxistats.com", "城市狀態、里程碑時間軸（引公司原文）", "週", "中：彙整公司原文，頁面日期有不一致", "城市清單與日期", "滲透 B、C", "Derived", "robotaxistats.com/services/waymo"),
        ("Obi（叫車比價）", "舊金山同路線 Waymo／Uber／Lyft／Tesla 價格與等候", "半年", "中高：大樣本（9–9.4 萬筆）但只限灣區；為比價 app 商", "價格區塊主來源；人類叫車錨", "滲透 A1、A2", "Verified", "TechCrunch 2025-06-12、2026-01-27"),
        ("YipitData（經 Business Insider）", "城市叫車訂單金額份額", "月", "中：卡片交易樣本推估；服務區擴張使口徑漂移", "城市份額", "滲透 C", "Interested-party", "Uber CFO 曾引用同系列"),
        ("Gridwise", "司機每英里收入、平台抽成", "季", "中：司機樣本", "人類叫車錨（ph）", "In ph", "Derived", "gridwise.io"),
        ("RideWise", "各城牌價、行程情境", "不定期", "低：自身範例與結論矛盾；費率卡不含服務費", "只作形狀參考", "滲透 A2、C", "Interested-party", "getridewise.com"),
        ("Uber／Lyft 財報", "全球趟數、總訂單金額", "季", "高；但美國叫車趟數未拆", "市場天花板（線 1a RT_US_TRIPS）", "線 1a", "Verified", "Lyft FY2025：945.5 百萬趟"),
        ("AAA Your Driving Costs", "自有車每英里全成本", "年", "高", "自有車錨", "滲透 A3、每英里經濟", "Verified", "2025-09-16"),
        ("券商（Morgan Stanley、MoffettNathanson）", "單車成本、車隊預測", "不定期", "中：有覆蓋利益", "單車成本", "In vcost26", "Interested-party", "經 TNW、Dealroom 轉述"),
    ]
    for i, row in enumerate(rows, 3):
        for c, v in enumerate(row, 1): put(ws, i, c, v, wrap=True)
    put(ws, len(rows) + 4, 1, "建議：量以 Waymo 揭露＋CPUC 為主，價格以 Obi 為主，份額以 YipitData 為輔；robotaxitracker 作高頻交叉核對（需能讀到數值與方法細節後再升級）。YipitData 若可直接訂閱，城市份額的可靠度可升一級。", wrap=True)
    for c, w in zip("ABCDEFGH", (26, 40, 12, 40, 30, 22, 14, 34)): ws.column_dimensions[c].width = w


def build_recon(ws, P):
    put(ws, 1, 1, f"對帳 — 本模型使用的線層級變數（線 1a Robotaxi v0.2 為權威：{LINE['file']}）。v0.2 依對帳 2（2026-11-17）改讀線；與 tesla_robotaxi v0.3 同一組市場、滲透、份額。", BOLD)
    hd = ["變數", "所屬線", "本模型使用值（或區間）", "單位", "口徑", "出處（讀線儲存格見 In 頁 E 節 H 欄）", "標記", "影響本模型哪些儲存格", "與 v0.1 的差異／對帳 2 結論"]
    for c, h in enumerate(hd, 1): put(ws, 2, c, h, BOLD, fill=HDR)
    f2 = lambda x: f"{x:.2f}"
    rows = [
        ("美國叫車年趟數", "1a", f"2026 {P['lt_2026_B']/1000:.2f}；2035 A {P['lt_2035_A']/1000:.2f}／B {P['lt_2035_B']/1000:.2f}／C {P['lt_2035_C']/1000:.2f}", "十億趟", "Uber＋Lyft，由 Lyft 趟數÷趟數份額", "線 情境頁", "Derived", "In lt_*；市場天花板 第 1 列", "v0.1 自建 3.5（2025）×1.05；tesla_robotaxi v0.2 為 3.79（每趟 7.1 英里換成里程）。三者改為同一值"),
        ("Robotaxi 滲透率（趟數）2035", "1a", f"A {P['lpen_2035_A']:.1%}／B {P['lpen_2035_B']:.1%}／C {P['lpen_2035_C']:.1%}", "%", "覆蓋率 × 都會區內份額 S 曲線；2026 校準", "線 情境頁", "Derived（輸入 Assumed）", "In lpen_*", "v0.1 自設 50／25／8%（與 tesla_robotaxi 同值、路徑不同）"),
        ("營運商份額 2035（Waymo／Tesla／其他）", "1a", f"A {P['lws_2035_A']:.0%}／50%／20%；B {P['lws_2035_B']:.0%}／35%／25%；C {P['lws_2035_C']:.0%}／20%／25%", "%", "占美國 Robotaxi 趟數；三家合計 100%", "線 情境頁", "Assumed（線）", "In lws_*", "v0.1 Waymo 55／45／25% 與 tesla_robotaxi 45／30／12% 在情境 A 合計 100%（其他 0%），C 合計 37%。改讀線後各情境合計 = 100%（對帳 2 議題 1）。情境 A 是 Tesla 勝出的高滲透世界，Waymo 份額最低"),
        ("每趟付費英里", "1a", f"2026 {P['tl0']:.2f}（CPUC）→ 2035 {P['tl_end']:.1f}（tl_end）", "英里/趟", "Waymo 自身行程；市場平均 6.0 只用於收斂終點", "線 每英里價格頁、輸入 M_TRIP_MI", "Derived／Assumed", "In tl0、tl_end；市場天花板 L 列", "v0.1 固定 4.0（3.4–5.5）；tesla_robotaxi v0.2 用人類叫車 7.1 換算市場里程。對帳 2 議題 2：預設丙（收斂），低＝乙 3.73 不變"),
        ("Waymo 每付費英里價格", "1a", f"2026 {P['lp_2026_B']:.2f}；2032 A {P['lp_2032_A']:.2f}／B {P['lp_2032_B']:.2f}／C {P['lp_2032_C']:.2f}；2033 起持平", "US$/英里", "乘客實付，不含小費", "線 成本輸入 W_PRICE", "Assumed（線）", "In lp_*", "v0.1 自設 4.5 → 2035 1.8／2.3／3.0（幾何）"),
        ("Waymo 營運成本第 3–8 項", "1a", f"2026 {P['lx_2026_B']:.2f}；2030 A {P['lx_2030_A']:.2f}／B {P['lx_2030_B']:.2f}／C {P['lx_2030_C']:.2f}", "US$/付費英里", "能源、維修清潔、保險理賠、遠端營運、安全員、場站", "線 每英里成本堆疊", "Derived（線）", "In lx_*；每英里經濟 cpm 列", "v0.1 自建每車英里 0.85、年降 12%、÷ 付費占比 59%（約 1.44／付費英里）。車輛攤提與融資不讀線（與 tesla_robotaxi 同規則）"),
        ("Waymo 稼動率（每車每日付費英里）", "1a", f"2026 {P['lm_2026_B']:.0f}；2032 A {P['lm_2032_A']:.0f}／B {P['lm_2032_B']:.0f}／C {P['lm_2032_C']:.0f}", "英里/日", "承重變數", "線 每英里成本堆疊備忘列", "Derived／Assumed（線）", "In lm_*；車隊、折舊", "v0.1 每車每日 20 趟 × 4.0 英里 ＝ 80，且不隨時間改善"),
        ("價格彈性 ε", "1a", f"A {P['eps_A']}／B {P['eps_B']}／C {P['eps_C']}", "—", "誘發倍數＝(2026 價÷當年價)^ε", "線 輸入 P_ELAST（NBER w31330）", "Analogy", "In eps_*", "v0.1 MAX(1, 人類價÷Waymo 價)^ε，ε 1.2／0.7／0.3。改為相對 2026，避免與已校準的 2026 趟數重複計入（對帳 2 工程結論）；線情境頁的市場倍數用市場價格路徑，與營運商價格不一致，不採用（線 v0.3 待修）"),
        ("自有車 VMT 轉換（閘門）", "1a／1b", "VMT 3,335.7 十億英里 × 小客車 90% × c35 × 閘門；門檻＝0.772 × cf 1.0", "十億英里", "價格≤門檻全開、≥1.5×門檻關閉", "FRED；AAA；Assumed", "Assumed", "市場天花板 gate、conv 列", "cf 1.5→1.0（對帳 1 預設乙）；Waymo 價格（2032 A 1.5）仍高於 1.155，閘門關閉，此路徑為 0"),
        ("單車取得成本、壽命", "1a", "125 千美元（年降 8%）、6 年", "千美元／年", "", "Morgan Stanley（TNW）", "Interested-party／Assumed", "In vcost26、vlife", "線 W_VEH 2026 基準 140 千美元，未讀（與 tesla_robotaxi 同：只讀第 3–8 項）"),
        ("固定營運費用", "1a", "5.0（3.5–6.5），年增 2%", "十億美元", "研發、地圖、管銷、擴城", "Alphabet Other Bets", "Derived", "In fix26", "不讀線第 9–10 項（線按車隊分攤，口徑不同；與 tesla_robotaxi 用營收 15% 不同，屬各自公司資料）"),
        ("海外", "—", "intl35 A 30%／B 15%／C 0%", "%", "國際趟數 ÷ 美國", "管理層、媒體", "Assumed", "In intl35_*", "線無海外，沿用 v0.1"),
        ("AAA 自有車每英里全成本", "1b／1a", "0.772（小轎車 0.559）", "US$/英里", "15,000 英里／年", "AAA 2025-09-16", "Verified", "閘門、錨", "不變"),
    ]
    for i, row in enumerate(rows, 3):
        for c, v in enumerate(row, 1): put(ws, i, c, v, wrap=True)
    for c, w in zip("ABCDEFGHI", (30, 8, 34, 12, 26, 30, 14, 26, 60)): ws.column_dimensions[c].width = w


# ---------------------------------------------------------------------------
def recalc(path):
    tmp = tempfile.mkdtemp()
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", tmp, path], check=True, capture_output=True, timeout=240)
    return os.path.join(tmp, os.path.basename(path))


def main():
    P, build_id = build()
    rc = recalc(OUT)
    wv = openpyxl.load_workbook(rc, data_only=True)
    errs = []
    for sh in wv.worksheets:
        for row in sh.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith(("#", "Err:")) and c.value[:4] in ("#DIV", "#VAL", "#REF", "#NAM", "#N/A", "#NUM", "#NUL", "Err:"):
                    errs.append(f"{sh.title}!{c.coordinate}={c.value}")
    print("錯誤值：", len(errs), errs[:10])
    # 鏡像比對
    ok = True
    for s in SCN:
        py = run(P, s)["eq"]
        cell = wv.defined_names[f"eq_{s}"].attr_text
        sh, ref = cell.split("!")
        xv = wv[sh.strip("'")][ref.replace("$", "")].value
        print(f"情境 {s}：xlsx {xv:.4f}  Python {py:.4f}")
        ok &= abs(xv - py) < 1e-6
    print("鏡像一致：", ok)
    # 以重算後檔案取代（保留公式與快取值）
    shutil.copy(rc, OUT)
    for cmd in (["python3", "tools/in_layout/check_in_layout.py", OUT],
                ["python3", "tools/sens_coverage/check_sens_coverage.py", OUT, "--require-keycol"],
                ["python3", "tools/model_governance/check_governance.py", OUT]):
        res = subprocess.run(cmd, cwd=ROOT, capture_output=True, text=True)
        print("$", " ".join(cmd[1:2]), "→", res.returncode); print(res.stdout[-1500:], res.stderr[-800:])
    if errs or not ok:
        sys.exit(1)
    # 摘要
    out = {s: run(P, s) for s in SCN}
    summ = {"build_id": build_id, "eq": {s: out[s]["eq"] for s in SCN}, "eq_w": weighted(P),
            "wk_B": {y: out["B"]["wk"][y] for y in (2026, 2030, 2035)}, "pen26": out["B"]["pen26"],
            "pm_2035": {s: out[s]["pm"][2035] for s in SCN}, "rev_2035": {s: out[s]["rev"][2035] for s in SCN},
            "full_2032": {s: out[s]["full"][2032] for s in SCN}, "line_file": LINE["file"]}
    json.dump(summ, open(os.path.join(HERE, "dist", "summary.json"), "w"), ensure_ascii=False, indent=1)
    print(json.dumps(summ, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
