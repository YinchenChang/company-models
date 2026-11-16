#!/usr/bin/env python3
"""Waymo（Alphabet 分部，視同未上市）個股模型 v0.1 建置腳本。

Robotaxi 分部模型特別規格（Andy 2026-11-16 22:33／22:38）：主頁＝「滲透」，最上方為「價格」區塊（每英里實收價格）；
其次「市場天花板與情境」「每英里經濟」「估值」「資料來源比較」「對帳」。頁面與口徑與 physical_ai/tesla_robotaxi/ 一致。

用法：python3 physical_ai/waymo/build_waymo.py   →  physical_ai/waymo/dist/20261116_Waymo個股模型_v0.1.xlsx
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

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as CL

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VERSION = "v0.1"
AS_OF = "2026-11-16"
OUT = os.path.join(HERE, "dist", f"20261116_Waymo個股模型_{VERSION}.xlsx")

BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
SEC = PatternFill("solid", fgColor="FFF2CC")
WRAP = Alignment(wrap_text=True, vertical="top")
YEARS = list(range(2026, 2036))
YC = {y: CL(3 + i) for i, y in enumerate(YEARS)}  # C..L
SCN = ["A", "B", "C"]
SCN_NAME = {"A": "情境 A（樂觀：價格大降＋傑文斯效應實現）", "B": "情境 B（基準）", "C": "情境 C（保守：滲透慢、價格不降）"}

# ---------------------------------------------------------------------------
# 一、輸入（In 頁）：鍵、項目、單位、低、基準、高、標記、出處、來源誘因、說明
# ---------------------------------------------------------------------------
IN = [
    ("§", "A. 市場（美國叫車，線 1a 口徑）"),
    ("us_trips25", "美國叫車（Uber＋Lyft 等）年趟數 2025", "十億趟", 3.0, 3.5, 4.2, "Derived",
     "Lyft FY2025 新聞稿（2026-02-10）：全年 945.5 百萬趟（含 2025 年中併入的 FreeNow 歐洲）；Lyft 美國份額約 24%（Second Measure，2024-03，經 Backlinko 轉述）",
     "", "≈（945.5 − FreeNow 歐洲，未取得）÷ 0.24–0.30；Uber 全球 13,567 百萬趟含外送，無法拆美國叫車，故不用"),
    ("g_mkt", "美國叫車趟數年成長（人類叫車基底）", "%/年", 0.03, 0.05, 0.07, "Assumed", "—", "",
     "Lyft 2025 趟數 +14%、Uber 全球 +20%（含外送與國際）；成熟後基底成長放緩，給區間"),
    ("w_trips26", "Waymo 2026 全年付費趟數", "百萬趟", 24, 28, 34, "Derived",
     "Waymo 部落格：2026-02-02 每週 40 萬＋；2026-03-26 每週 50 萬＋；TechCrunch 2026-09-24：平均每週約 50 萬；年底目標每週 100 萬（Bloomberg 2026-02-11，共同執行長）",
     "", "≈ 52 週 × 平均每週 0.5–0.65 百萬；Q4 爬坡未取得官方數字"),
    ("w_avshare26", "Waymo 占美國付費無人 Robotaxi 趟數份額 2026", "%", 0.85, 0.90, 0.95, "Assumed", "—", "",
     "Tesla（奧斯汀、灣區有安全員）、Zoox（拉斯維加斯 2026-07 起收費）規模遠小於 Waymo；精確份額未取得"),
    ("ph", "人類叫車每付費英里乘客實付價格（含小費，全美）", "US$/英里", 2.8, 3.7, 4.5, "Derived",
     "Gridwise 2026 Q3：Uber 司機每英里毛收入 US$2.65、平台抽成中位數 22.44%（線 1a 框架）",
     "", "2.65 ÷（1 − 0.2244）× 1.08（小費，與 tesla_robotaxi 同）≈ 3.69；舊金山 Obi 2026-01 Uber 均價 17.47 ÷ 3.4 英里 ≈ 5.1（高價城市）→ 區間取 2.8–4.5"),
    ("vmt", "美國 12 個月移動車輛行駛里程（2026-08）", "十億英里", None, 3335.695, None, "Verified", "FRED／FHWA TRAVLTRS（與 tesla_robotaxi 同值）", "", "政府統計"),
    ("ldv_sh", "VMT 中可被 Robotaxi 取代的小客車占比", "%", 0.85, 0.90, 0.92, "Assumed", "—", "", "與 tesla_robotaxi 同"),
    ("p_own", "自有車每英里全成本（AAA 2025，15,000 英里／年平均）", "US$/英里", 0.5587, 0.772, None, "Verified", "AAA Your Driving Costs 2025-09-16：年 11,577 美元；小轎車 55.87 美分（低）", "", "與 tesla_robotaxi 同"),
    ("cf", "便利溢價（Robotaxi 每英里價格可高於自有車成本的倍數，轉換閘門門檻）", "倍", 1.2, 1.5, 2.0, "Assumed", "—", "", "閘門＝MIN(1,MAX(0,(1.5×門檻−價)/(0.5×門檻)))，門檻＝p_own×cf；與 tesla_robotaxi 同式"),
    ("§", "B. Waymo 營運參數（每英里與車隊）"),
    ("pmi_trip", "每趟付費（載客）英里", "英里/趟", 3.4, 4.0, 5.5, "Analogy",
     "Findings Press（2026）以 CPUC 未遮蔽資料：加州 2023-08～2025-12 每趟載客約 3.4 英里（推算）", "",
     "類比：加州（舊金山為主、短程）＝下緣；全美含機場與高速公路行程較長；上緣 5.5 為 Assumed"),
    ("paid_sh26", "付費英里占總行駛英里 2026", "%", 0.55, 0.59, 0.62, "Derived",
     "CPUC 月報（Platform Aeronaut 彙整）：2026-01～06 空駛比 40.7%–45.0%", "", "1 − 空駛比；部分月份有遮蔽（＋號）"),
    ("paid_sh35", "付費英里占總行駛英里 2035", "%", 0.60, 0.65, 0.70, "Assumed", "—", "", "規模與調度改善；人類叫車類比 55–65%（線 1a）"),
    ("rpm26", "Waymo 每付費英里實收價格 2026（全美平均）", "US$/英里", 3.5, 4.5, 5.8, "Derived",
     "Obi（TechCrunch 2026-01-27）：舊金山 Waymo 均價 US$19.69／趟；÷ 每趟 3.4 英里 ≈ 5.79（舊金山＝上緣）", "",
     "其他城市 Waymo 牌價較低（RideWise 費率表，可靠度低）；全美平均未取得"),
    ("cpm_rao", "遠端協助每車英里 2026", "US$/英里", 0.05, 0.12, 0.20, "Assumed", "Markey 參議員報告（2026）：各家未揭露 RAO 比例；Waymo 約半數 RAO 在菲律賓", "", "RAO：車 1:10–1:40 × 時薪"),
    ("cpm_ins", "保險每車英里 2026", "US$/英里", 0.10, 0.20, 0.30, "Assumed", "未取得", "", "商用車隊保險；Waymo 未揭露"),
    ("cpm_nrg", "能源每車英里 2026", "US$/英里", 0.04, 0.06, 0.08, "Assumed", "—", "", "約 0.3 kWh/英里 × 商用電價（含充電損失）"),
    ("cpm_mnt", "維修與清潔每車英里 2026", "US$/英里", 0.15, 0.25, 0.35, "Assumed", "未取得", "", "含感測器校正、內裝清潔"),
    ("cpm_depot", "場站與車隊管理每車英里 2026", "US$/英里", 0.12, 0.22, 0.40, "Assumed", "未取得", "", "場站租金、充電樁、現場人員"),
    ("cpm_dec", "每車英里現金成本年降幅", "%/年", 0.06, 0.12, 0.15, "Assumed", "—", "", "規模經濟、場站利用率、RAO 比例改善"),
    ("vcost26", "單車取得成本 2026（Ojai＋第六代套件）", "千美元", 100, 125, 160, "Interested-party",
     "Morgan Stanley 估計（TNW 轉述，線 1a）：Ojai 約 12.5 萬美元；舊款 I-PACE 約 20 萬", "券商估計（有覆蓋 GOOGL 的利益）", ""),
    ("vcost_dec", "單車成本年降幅", "%/年", 0.04, 0.08, 0.12, "Assumed", "—", "", "套件成本下降、Zeekr／Hyundai 量產"),
    ("vlife", "車輛經濟壽命", "年", 5, 6, 8, "Assumed", "—", "", "生涯 30–50 萬英里（線 1a）÷ 每車年英里"),
    ("tpvd", "每車每日付費趟數", "趟/車/日", 15, 20, 26, "Derived",
     "每週 50 萬趟 ÷ 7 ÷ 車隊約 3,000（2026-04）≈ 24；÷ 約 4,000（2026-09）≈ 18", "", "時點不完全一致；中國同業 21–25（線 1a）"),
    ("fix26", "固定營運費用 2026（研發、地圖、管銷、擴城）", "十億美元", 3.5, 5.0, 6.5, "Derived",
     "Alphabet 2026 Q3 Other Bets 營業虧損 US$20.5 億、營收 3.72 億（Fortune 2026-10-30）；年化約 82 億為 Other Bets 合計", "",
     "Waymo 未單獨揭露；扣除 Wing、Verily 等其他單位後取 3.5–6.5"),
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
    "avpen30": ("Robotaxi 占美國叫車（人類替代部分）滲透率 2030", "%", 0.20, 0.10, 0.03, "2026 起幾何內插；2026 起點由 Waymo 趟數反推（約 0.9%）"),
    "avpen35": ("Robotaxi 占美國叫車（人類替代部分）滲透率 2035", "%", 0.50, 0.25, 0.08, "2030→2035 線性；2035 值與 tesla_robotaxi s35 相同（線 1a 共用）"),
    "c35": ("自有車 VMT 轉換為 Robotaxi 的比例（閘門開啟時）2035", "%", 0.10, 0.03, 0.005, "與 tesla_robotaxi c35 相同；2027 起線性 × 閘門"),
    "wsh30": ("Waymo 占美國 Robotaxi 趟數份額 2030", "%", 0.65, 0.55, 0.35, "Tesla、Zoox、Uber 平台二線營運商進入"),
    "wsh35": ("Waymo 占美國 Robotaxi 趟數份額 2035", "%", 0.55, 0.45, 0.25, ""),
    "rpm35": ("Waymo 每付費英里實收價格 2035", "US$/英里", 1.8, 2.3, 3.0, "2026 起幾何內插；Tesla 奧斯汀牌價 US$1/英里＋起跳 3.25 為價格戰參照"),
    "eps": ("價格彈性（傑文斯效應）：誘發需求 ＝ MAX(1, 人類價格 ÷ Robotaxi 價格)^ε", "倍數", 1.2, 0.7, 0.3, "觀察到的彈性未取得（見滲透頁 A4）"),
    "intl35": ("國際趟數 ÷ 美國趟數 2035", "%", 0.30, 0.15, 0.0, "倫敦、東京（2026 宣布）、加拿大、德國、新加坡、西班牙（Q3 法說）；2027 起線性"),
}
SCN_W = {"A": 0.25, "B": 0.50, "C": 0.25}


def base_params():
    p = {r[0]: r[4] for r in IN if r[0] != "§"}
    p["stub"] = 45 / 365
    p["cpm26"] = p["cpm_rao"] + p["cpm_ins"] + p["cpm_nrg"] + p["cpm_mnt"] + p["cpm_depot"]
    for k, v in SCN_IN.items():
        for i, s in enumerate(SCN):
            p[f"{k}_{s}"] = v[2 + i]
    return p


# ---------------------------------------------------------------------------
# 二、Python 鏡像（與 xlsx 公式一一對應；Sens 與反解用）
# ---------------------------------------------------------------------------
def run(p, s):
    p = dict(p)
    p["cpm26"] = p["cpm_rao"] + p["cpm_ins"] + p["cpm_nrg"] + p["cpm_mnt"] + p["cpm_depot"]
    us26 = p["us_trips25"] * (1 + p["g_mkt"])
    pen26 = p["w_trips26"] / 1000 / p["w_avshare26"] / us26
    out = {k: {} for k in ["us", "pen", "rpm", "jev", "gate", "conv", "av", "wsh", "wus", "intl", "trips", "wk", "fleet", "pm",
                           "psh", "vm", "cpm", "vc", "dep_mi", "opc", "fix", "fix_mi", "full", "rev", "var", "dep",
                           "ebit", "nopat", "capex", "dnwc", "fcff", "pv"]}
    prev_fleet = prev_rev = None
    for y in YEARS:
        n = y - 2026
        us = p["us_trips25"] * (1 + p["g_mkt"]) ** (y - 2025)
        pen = pen26 * (p[f"avpen30_{s}"] / pen26) ** (n / 4) if y <= 2030 else p[f"avpen30_{s}"] + (p[f"avpen35_{s}"] - p[f"avpen30_{s}"]) * (y - 2030) / 5
        rpm = p["rpm26"] * (p[f"rpm35_{s}"] / p["rpm26"]) ** (n / 9)
        jev = max(1, p["ph"] / rpm) ** p[f"eps_{s}"]
        gate = min(1, max(0, (1.5 * p["p_own"] * p["cf"] - rpm) / (0.5 * p["p_own"] * p["cf"])))
        conv = p["vmt"] * p["ldv_sh"] * p[f"c35_{s}"] * n / 9 * gate
        av = us * pen * jev + conv / p["pmi_trip"]
        wsh = p["w_avshare26"] + (p[f"wsh30_{s}"] - p["w_avshare26"]) * n / 4 if y <= 2030 else p[f"wsh30_{s}"] + (p[f"wsh35_{s}"] - p[f"wsh30_{s}"]) * (y - 2030) / 5
        wus = av * wsh * 1000
        intl = p[f"intl35_{s}"] * n / 9
        trips = wus * (1 + intl)
        fleet = trips / (p["tpvd"] * 365) * 1000
        pm = trips * p["pmi_trip"]
        psh = p["paid_sh26"] + (p["paid_sh35"] - p["paid_sh26"]) * n / 9
        vm = pm / psh
        cpm = p["cpm26"] * (1 - p["cpm_dec"]) ** n
        vc = p["vcost26"] * (1 - p["vcost_dec"]) ** n
        dep_mi = vc / p["vlife"] / (vm / fleet)
        opc = (cpm + dep_mi) / psh
        fix = p["fix26"] * (1 + p["fix_g"]) ** n
        fix_mi = fix * 1000 / pm
        rev = pm * rpm / 1000
        var = vm * cpm / 1000
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
                         fleet=fleet, pm=pm, psh=psh, vm=vm, cpm=cpm, vc=vc, dep_mi=dep_mi, opc=opc, fix=fix, fix_mi=fix_mi,
                         full=opc + fix_mi, rev=rev, var=var, dep=dep, ebit=ebit, nopat=nopat, capex=capex, dnwc=dnwc,
                         fcff=fcff, pv=pv).items():
            out[k][y] = v
        prev_fleet, prev_rev = fleet, rev
    sum_pv = sum(out["pv"].values())
    n36 = out["nopat"][2035] * (1 + p["g_tv"])
    tv = n36 * (1 - p["g_tv"] / p["wacc"]) / (p["wacc"] - p["g_tv"])
    pvtv = tv / (1 + p["wacc"]) ** (p["stub"] + 9)
    ev = sum_pv + pvtv
    out.update(sum_pv=sum_pv, tv=tv, pvtv=pvtv, ev=ev, eq=ev + p["netcash"], pen26=pen26)
    return out


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
    # 衍生：cpm26
    put(ws, r, 1, "cpm26"); put(ws, r, 2, "每車英里現金成本 2026（五項合計）"); put(ws, r, 3, "US$/英里")
    put(ws, r, 5, "=cpm_rao+cpm_ins+cpm_nrg+cpm_mnt+cpm_depot"); put(ws, r, 7, "Derived"); put(ws, r, 8, "上列五項加總")
    name(wb, "cpm26", f"In!$E${r}"); r += 1
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
    put(ws, r, 1, "scn_w"); put(ws, r, 2, "情境權重")
    for c, s in zip((4, 5, 6), SCN):
        put(ws, r, c, f"=E{inrow['scn_w' + s]}")
    put(ws, r, 8, "引用 In 頁 scn_wA／scn_wB／scn_wC；合計＝1")
    for c, w in zip("ABCDEFGHIJ", (12, 46, 10, 10, 10, 10, 14, 60, 24, 50)):
        ws.column_dimensions[c].width = w

    # ---------------- 市場天花板與情境 ----------------
    ws = W["市場天花板與情境"]
    put(ws, 1, 1, "市場天花板與情境 — 美國叫車總趟數 × Robotaxi 滲透 × 傑文斯誘發需求 × Waymo 份額 ×（1＋國際）→ Waymo 趟數、車隊、付費里程（十億／百萬趟）", BOLD)
    put(ws, 2, 1, "2026 起點滲透率（由 Waymo 2026 趟數反推）"); put(ws, 2, 3, "=w_trips26/1000/w_avshare26/(us_trips25*(1+g_mkt))", fmt="0.00%")
    name(wb, "pen26", "'市場天花板與情境'!$C$2")
    put(ws, 3, 1, "天花板讀法：叫車市場只是美國出行的一小部分；AAA 2025 自有車每英里 0.77 美元（小轎車 0.56），Robotaxi 價格若逼近此線，可替代的不只叫車而是自有車里程——本版兩條路徑（與 tesla_robotaxi 同式）：(1) 傑文斯係數 ε 放大『低於人類叫車價格』的叫車需求；(2) 價格低於 自有車成本×便利溢價 時，閘門開啟，部分小客車 VMT 轉入 Robotaxi（c35）。單位換算：轉入英里 ÷ 每趟英里＝趟數。", wrap=True)
    VROWS = ["us", "pen", "rpm", "jev", "gate", "conv", "av", "wsh", "wus", "intl", "trips", "wk", "shus", "fleet", "pm"]
    VLAB = {"us": ("美國叫車基底趟數", "十億趟"), "pen": ("Robotaxi 滲透率（人類叫車替代）", "%"),
            "rpm": ("Waymo 每付費英里實收價格", "US$/英里"), "jev": ("傑文斯誘發係數 MAX(1,人類價÷RT價)^ε", "倍"),
            "gate": ("自有車轉換閘門（價格 vs 自有車成本×便利溢價）", "0–1"), "conv": ("自有車里程轉入 Robotaxi", "十億英里"),
            "av": ("美國 Robotaxi 趟數", "十億趟"), "wsh": ("Waymo 占 Robotaxi 份額", "%"), "wus": ("Waymo 美國趟數", "百萬趟"),
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
                "us": f"=us_trips25*(1+g_mkt)^({y}-2025)",
                "pen": (f"=pen26*(avpen30_{s}/pen26)^({n}/4)" if y <= 2030 else f"=avpen30_{s}+(avpen35_{s}-avpen30_{s})*({y}-2030)/5"),
                "rpm": f"=rpm26*(rpm35_{s}/rpm26)^({n}/9)",
                "jev": f"=MAX(1,ph/{c}{R('rpm')})^eps_{s}",
                "gate": f"=MIN(1,MAX(0,(1.5*p_own*cf-{c}{R('rpm')})/(0.5*p_own*cf)))",
                "conv": f"=vmt*ldv_sh*c35_{s}*{n}/9*{c}{R('gate')}",
                "av": f"={c}{R('us')}*{c}{R('pen')}*{c}{R('jev')}+{c}{R('conv')}/pmi_trip",
                "wsh": (f"=w_avshare26+(wsh30_{s}-w_avshare26)*{n}/4" if y <= 2030 else f"=wsh30_{s}+(wsh35_{s}-wsh30_{s})*({y}-2030)/5"),
                "wus": f"={c}{R('av')}*{c}{R('wsh')}*1000",
                "intl": f"=intl35_{s}*{n}/9",
                "trips": f"={c}{R('wus')}*(1+{c}{R('intl')})",
                "wk": f"={c}{R('trips')}/52",
                "shus": f"={c}{R('wus')}/({c}{R('us')}*1000)",
                "fleet": f"={c}{R('trips')}/(tpvd*365)*1000",
                "pm": f"={c}{R('trips')}*pmi_trip",
            }
            for k in VROWS:
                fmt = "0.0%" if k in ("pen", "wsh", "intl", "shus") else ("0.00" if k in ("rpm", "jev", "gate", "us", "av", "wk") else "#,##0.0")
                put(ws, R(k), col, f[k], fmt=fmt)
    ws.column_dimensions["A"].width = 40; ws.column_dimensions["B"].width = 12

    # ---------------- 每英里經濟 ----------------
    ws = W["每英里經濟"]
    put(ws, 1, 1, "每英里經濟 — 價格 vs 每英里全成本（US$／付費英里）；收支只作估值的一環。成本口徑：營運成本＝（每車英里現金成本＋每車英里折舊）÷ 付費英里占比；全載成本＝營運成本＋固定費用 ÷ 付費英里", BOLD)
    put(ws, 3, 1, "2026 每車英里現金成本拆解（Assumed；與線 1a 同式）", BOLD, fill=SEC)
    comps = [("cpm_rao", "遠端協助"), ("cpm_ins", "保險"), ("cpm_nrg", "能源"), ("cpm_mnt", "維修清潔"), ("cpm_depot", "場站與車隊管理"), ("cpm26", "合計（每車英里）")]
    for c, h in enumerate(["項目", "鍵", "低", "基準", "高"], 1): put(ws, 4, c, h, BOLD, fill=HDR)
    r = 5
    for k, lab in comps:
        put(ws, r, 1, lab); put(ws, r, 2, k)
        if k == "cpm26":
            put(ws, r, 3, "=SUM(C5:C9)", fmt="0.00"); put(ws, r, 4, "=cpm26", fmt="0.00"); put(ws, r, 5, "=SUM(E5:E9)", fmt="0.00")
        else:
            put(ws, r, 3, f"=In!D{inrow[k]}", fmt="0.00"); put(ws, r, 4, f"={k}", fmt="0.00"); put(ws, r, 5, f"=In!F{inrow[k]}", fmt="0.00")
        r += 1
    put(ws, 11, 1, "錨（每英里）", BOLD, fill=SEC)
    anchors = [("人類叫車乘客實付（含小費，全美）ph", "=ph", "Derived：Gridwise 2026 Q3"),
               ("AAA 2025 自有車全成本（15,000 英里／年，平均）", "=11577/15000", "Verified：AAA Your Driving Costs 2025-09-16（年 11,577 美元）"),
               ("AAA 2025 小轎車", 0.5587, "Verified：同上，55.87 美分"),
               ("Tesla 奧斯汀牌價（每英里部分，不含起跳 3.25）", 1.00, "Verified（媒體轉述 app 價格，線 1a）"),
               ("Tesla 營運成本目標", 0.30, "Interested-party：Tesla 管理層（線 1a）")]
    r = 12
    for lab, v, src in anchors:
        put(ws, r, 1, lab); put(ws, r, 3, v, BLUE if not isinstance(v, str) else None, fmt="0.00"); put(ws, r, 4, src); r += 1
    MROWS = ["rpm", "psh", "vm", "cpm", "vc", "mpv", "dep_mi", "opc", "fix", "fix_mi", "full", "gap", "flag", "vsaaa"]
    MLAB = {"rpm": ("每付費英里實收價格", "US$/英里"), "psh": ("付費英里占比", "%"), "vm": ("總行駛英里", "百萬英里"),
            "cpm": ("每車英里現金成本", "US$/英里"), "vc": ("單車成本", "千美元"), "mpv": ("每車年行駛英里", "千英里"),
            "dep_mi": ("每車英里折舊", "US$/英里"), "opc": ("營運成本（每付費英里，含折舊）", "US$/英里"),
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
                "psh": f"=paid_sh26+(paid_sh35-paid_sh26)*{n}/9",
                "vm": f"={V}{c}{VR[(s, 'pm')]}/{c}{R('psh')}",
                "cpm": f"=cpm26*(1-cpm_dec)^{n}",
                "vc": f"=vcost26*(1-vcost_dec)^{n}",
                "mpv": f"={c}{R('vm')}/{V}{c}{VR[(s, 'fleet')]}",
                "dep_mi": f"={c}{R('vc')}/vlife/{c}{R('mpv')}",
                "opc": f"=({c}{R('cpm')}+{c}{R('dep_mi')})/{c}{R('psh')}",
                "fix": f"=fix26*(1+fix_g)^{n}",
                "fix_mi": f"={c}{R('fix')}*1000/{V}{c}{VR[(s, 'pm')]}",
                "full": f"={c}{R('opc')}+{c}{R('fix_mi')}",
                "gap": f"={c}{R('rpm')}-{c}{R('opc')}",
                "flag": f'=IF({c}{R("rpm")}<{c}{R("opc")},"低於營運成本：補貼",IF({c}{R("rpm")}<{c}{R("full")},"高於營運、低於全載","覆蓋全載成本"))',
                "vsaaa": f"={c}{R('rpm')}/(11577/15000)",
            }
            for k in MROWS:
                fmt = "0.0%" if k == "psh" else ("#,##0" if k in ("vm",) else "0.00")
                put(ws, R(k), col, f[k], fmt=fmt)
    ws.column_dimensions["A"].width = 44; ws.column_dimensions["D"].width = 12

    # ---------------- 估值 ----------------
    ws = W["估值"]
    put(ws, 1, 1, "估值 — 分部 DCF（FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加；明確期 FY2026 殘期～FY2035；永續 NOPAT(2036)×(1−g/RONIC)/(WACC−g)，RONIC＝WACC）。十億美元。評等：未上市。", BOLD)
    put(ws, 2, 1, "營收分部：Waymo 未上市、無報告分部；營收列＝「Robotaxi 乘車營收（美國＋國際）」一列（直營 app 與 Uber 平台城市的拆分未揭露）。")
    DROWS = ["rev", "var", "fix", "dep", "ebit", "tax", "nopat", "capex", "dnwc", "fcff", "frac", "t", "pv"]
    DLAB = {"rev": "Robotaxi 乘車營收（美國＋國際）", "var": "變動現金成本（總行駛英里 × 每車英里現金成本）", "fix": "固定費用",
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
                "var": f"={M}{c}{MR[(s, 'vm')]}*{M}{c}{MR[(s, 'cpm')]}/1000",
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
    put(ws, 2, 1, "註：tax 低＝基準（21%）只列高；val_round 只作對照，不進模型。情境專屬輸入（*_A／_B／_C）無低／高欄，其不確定性由三情境表達。")
    hdr = 4
    for c, h in enumerate(["驅動", "設定", "鍵", "加權目標價（股權價值）", "相對基準", "情境 B 股權價值"], 1): put(ws, hdr, c, h, BOLD, fill=HDR)
    r = hdr + 1
    put(ws, r, 1, "基準"); put(ws, r, 2, "In 頁 E 欄"); put(ws, r, 4, round(base_w, 4)); put(ws, r, 5, 0, fmt="0.0%"); put(ws, r, 6, round(run(P, "B")["eq"], 4)); r += 1
    sens_rows = []
    for row in IN:
        if row[0] == "§" or row[3] is None: continue
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
    rev_items.append(("情境 B：avpen35_B＝100%（其他不變）時的股權價值", run({**P, "avpen35_B": 1.0}, "B")["eq"], "#,##0.0", "十億美元；仍低於融資估值 → 單靠滲透率不可達"))
    rev_items.append(("情境 A 股權價值＝融資估值所需的 WACC", solve(P, "wacc", P["val_round"], 0.035, 0.30, "A"), "0.0%", "g＝3%；基準 WACC 12%"))
    rev_items.append(("情境 B 股權價值＝0 所需的 2035 每付費英里價格（rpm35_B）", solve(P, "rpm35_B", 0.0, 1.0, 8.0, "B"), "0.00", "US$/英里；價格提高同時壓低傑文斯誘發需求"))
    rev_items.append(("情境 B 股權價值＝0 所需的價格彈性（eps_B）", solve(P, "eps_B", 0.0, 0.0, 5.0, "B"), "0.00", "傑文斯效應需多強"))
    rev_items.append(("情境 B 股權價值＝0 的 2026 固定費用（fix26）", solve(P, "fix26", 0.0, 0.0, 15.0, "B"), "0.00", "十億美元／年，2% 成長"))
    for lab, v, fmt, note in rev_items:
        put(ws, r, 1, lab); put(ws, r, 6, v if v is not None else "區間內無解", fmt=fmt); put(ws, r, 7, note); r += 1

    # ---------------- README 頁 ----------------
    ws = wsR
    put(ws, 1, 1, f"Waymo（Alphabet 分部，視同未上市）個股模型 {VERSION}｜build_id {build_id}｜資料截止日 {AS_OF}｜成品日 {AS_OF}｜股價基準日：不適用（未上市）", BOLD)
    lines = [
        "命題：Waymo 的價值取決於 Robotaxi 在美國（及全球）叫車的滲透速度，以及每英里實收價格能降到哪裡（決定傑文斯效應能否實現）；收支只是估值的一環。",
        "主頁＝「滲透」：最上方「價格」區塊（Waymo 每英里實收價格時間序列與人類叫車、自有車兩個錨）、全國里程碑、各城市、加州 CPUC 月度、全國份額。",
        "頁面（與 physical_ai/tesla_robotaxi/ 同一套）：滲透｜市場天花板與情境｜每英里經濟｜估值｜資料來源比較｜對帳｜In｜Sens。",
        "估值口徑：分部 DCF（FCFF，WACC 12%，g 3%，RONIC＝WACC）三情境 25／50／25；以 2026-02 融資投後 1,260 億美元為對照；同業倍數腿未建（理由見 README.md）。評等：未上市。",
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
    put(ws, 2, 1, "核心指標：每英里實收價格（Andy 2026-11-16 22:38：決定傑文斯效應能否實現）。口徑：乘客實付（含起跳價與費用，不含小費）÷ 載客英里；人類叫車錨＝UberX／Lyft 乘客實付 ×（1＋小費率）；自有車錨＝AAA 每車英里全成本。", wrap=True)
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
                              ("Waymo 全美平均 2026（模型輸入）", "=rpm26", "In 頁 rpm26", "Derived")]:
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
    put(ws, r, 1, "Waymo 2026 年趟數 ÷ 美國叫車年趟數"); put(ws, r, 3, "=w_trips26/1000/(us_trips25*(1+g_mkt))", fmt="0.00%"); r += 1
    put(ws, r, 1, "年底目標（每週 100 萬）年化 ÷ 美國叫車年趟數"); put(ws, r, 3, "=1*52/1000/(us_trips25*(1+g_mkt)^2)", fmt="0.00%"); r += 1
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
        ("Waymo 部落格、FAQ、新聞稿", "全國每週付費趟數、城市開放、累計英里、融資", "不定期（里程碑）", "高（數字可信但選擇性揭露，只報好消息）", "全國量的錨；年底目標作 Interested-party", "滲透 B、In w_trips26", "Interested-party", "waymo.com/blog；經 robotaxistats.com 彙整"),
        ("Alphabet 法說與 10-Q", "Other Bets 營收與營業虧損、城市數、英里", "季", "高（法定申報），但 Waymo 不單獨揭露", "燒錢上限（fix26）", "In fix26", "Verified", "Q3 2026 逐字稿 2026-10-28"),
        ("加州 CPUC AV 季報（含彙整網站）", "加州月度趟數、乘客、行駛英里、空駛比", "季（月資料）", "高（監管申報），但公開版大量遮蔽；需經彙整網站", "空駛比、每趟英里、加州量；彈性粗算", "滲透 D、In paid_sh26、pmi_trip", "Verified", "avrideshare.platformaeronaut.com；Findings Press 論文"),
        ("加州 DMV", "測試／部署許可、脫離報告", "年／即時", "高，但不含商業量", "許可追蹤", "未用", "Verified", "未取得數值"),
        ("NHTSA SGO 事故資料", "事故報告", "月", "高（強制申報），分母（英里）需另找", "停駛風險追蹤", "未用（風險登錄）", "Verified", "未取得數值"),
        ("德州 DMV 車籍（經媒體）", "德州車隊數", "不定期", "中（媒體整理，未具名資料集）", "車隊規模", "滲透 B、C", "Derived", "TechCrunch 2026-09-24"),
        ("robotaxitracker.com（Andy 常用）", "Tesla、Waymo、Zoox、Avride：車輛名錄、服務區地圖、乘車統計、價格、等候、政府資料（CPUC、DMV、NHTSA、奧斯汀事件、亞利桑那 MVD）", "持續（社群）", "中低：社群投稿＋等候時間遙測＋政府紀錄＋審核流程；Methodology 頁於本環境無法讀取內文（前端渲染、代理封鎖），只見來源類別；與公司無關聯", "等候與價格的高頻趨勢、車隊名錄交叉核對；不作估值錨", "未直接取數（本環境讀不到數值）", "Interested-party", "社群；robotaxitracker.com/methodology"),
        ("robotaxistats.com", "城市狀態、里程碑時間軸（引公司原文）", "週", "中：彙整公司原文，頁面日期有不一致", "城市清單與日期", "滲透 B、C", "Derived", "robotaxistats.com/services/waymo"),
        ("Obi（叫車比價）", "舊金山同路線 Waymo／Uber／Lyft／Tesla 價格與等候", "半年", "中高：大樣本（9–9.4 萬筆）但只限灣區；為比價 app 商", "價格區塊主來源；人類叫車錨", "滲透 A1、A2", "Verified", "TechCrunch 2025-06-12、2026-01-27"),
        ("YipitData（經 Business Insider）", "城市叫車訂單金額份額", "月", "中：卡片交易樣本推估；服務區擴張使口徑漂移", "城市份額", "滲透 C", "Interested-party", "Uber CFO 曾引用同系列"),
        ("Gridwise", "司機每英里收入、平台抽成", "季", "中：司機樣本", "人類叫車錨（ph）", "In ph", "Derived", "gridwise.io"),
        ("RideWise", "各城牌價、行程情境", "不定期", "低：自身範例與結論矛盾；費率卡不含服務費", "只作形狀參考", "滲透 A2、C", "Interested-party", "getridewise.com"),
        ("Uber／Lyft 財報", "全球趟數、總訂單金額", "季", "高；但美國叫車趟數未拆", "市場天花板（us_trips25）", "In us_trips25", "Verified", "Lyft FY2025：945.5 百萬趟"),
        ("AAA Your Driving Costs", "自有車每英里全成本", "年", "高", "自有車錨", "滲透 A3、每英里經濟", "Verified", "2025-09-16"),
        ("券商（Morgan Stanley、MoffettNathanson）", "單車成本、車隊預測", "不定期", "中：有覆蓋利益", "單車成本", "In vcost26", "Interested-party", "經 TNW、Dealroom 轉述"),
    ]
    for i, row in enumerate(rows, 3):
        for c, v in enumerate(row, 1): put(ws, i, c, v, wrap=True)
    put(ws, len(rows) + 4, 1, "建議：量以 Waymo 揭露＋CPUC 為主，價格以 Obi 為主，份額以 YipitData 為輔；robotaxitracker 作高頻交叉核對（需能讀到數值與方法細節後再升級）。YipitData 若可直接訂閱，城市份額的可靠度可升一級。", wrap=True)
    for c, w in zip("ABCDEFGH", (26, 40, 12, 40, 30, 22, 14, 34)): ws.column_dimensions[c].width = w


def build_recon(ws, P):
    put(ws, 1, 1, "對帳 — 本模型使用的線層級變數（線 1a Robotaxi 為主）。線模型完成版本代理讀不到；取值先參考線框架 v0.1，不同處照寫、不靜默改口徑。", BOLD)
    hd = ["變數", "所屬線", "本模型使用值（或區間）", "單位", "口徑", "出處", "標記", "影響本模型哪些儲存格", "與線框架 v0.1 的差異"]
    for c, h in enumerate(hd, 1): put(ws, 2, c, h, BOLD, fill=HDR)
    rows = [
        ("美國叫車年趟數 2025", "1a", "3.5（3.0–4.2）", "十億趟", "Uber＋Lyft 等人類叫車，美國", "Lyft FY2025 ÷ 美國份額", "Derived", "In us_trips25 → 市場天花板 各情境第 1 列", "線框架未給；新增，請 1a 採同口徑"),
        ("美國叫車基底成長", "1a", "5%（3–7%）", "%/年", "趟數", "—", "Assumed", "In g_mkt", "線框架未給"),
        ("Robotaxi 滲透率（人類叫車替代）2030／2035", "1a", "A 20%／50%；B 10%／25%；C 3%／8%", "%", "占美國叫車基底（本模型以趟數、tesla_robotaxi 以里程）", "—", "Assumed", "In avpen30_*、avpen35_*", "線框架未給；2035 值與 tesla_robotaxi s35 相同；路徑不同（本模型 2026 起幾何內插＋2030 後線性；Tesla 為 S 曲線）"),
        ("自有車 VMT 轉換（閘門）", "1a／1b", "VMT 3,335.7 十億英里 × 小客車 90% × c35（A 10%／B 3%／C 0.5%）× 閘門；門檻＝0.772 × 便利溢價 1.5", "十億英里", "價格≤門檻全開、≥1.5×門檻關閉", "FRED；AAA；Assumed", "Assumed", "市場天花板 gate、conv 列", "與 tesla_robotaxi 同式同值；Waymo 各情境 2035 價格（1.8–3.0）均高於 1.5×門檻（1.73），閘門關閉 → 本模型此路徑為 0"),
        ("人類叫車小費率", "1a", "8%（5–15%）", "%", "乘客實付 ÷ 車資", "Assumed", "Assumed", "滲透 A1、In ph", "與 tesla_robotaxi 同；RideWise 用 15%"),
        ("Robotaxi 營運商份額（Waymo＋Tesla）", "1a", "Waymo 2035：A 55%／B 45%／C 25%（趟數）", "%", "Tesla 端 sig35：A 45%／B 30%／C 12%（里程）", "—", "Assumed", "In wsh35_*", "兩模型情境 A 合計 100%，不可同時成立；B 合計 75%，餘 25% 給 Zoox、Uber 平台二線營運商等"),
        ("價格彈性 ε 的結構", "1a", "Waymo：情境專屬（A 1.2／B 0.7／C 0.3）＋ MAX(1,…)（價格高於人類時不抑制）", "—", "Tesla：單一輸入 0.7（0.3–1.2）", "—", "Assumed", "In eps_*", "數值相同、結構不同，對帳時統一"),
        ("價格彈性 ε（傑文斯）", "1a", "A 1.2；B 0.7；C 0.3", "—", "誘發係數＝MAX(1, 人類價÷RT 價)^ε", "觀察值未取得", "Assumed", "In eps_*", "線框架未給；Andy 22:38 指為核心"),
        ("人類叫車乘客每英里實付（含小費）", "1a", "3.7（2.8–4.5）＝2.65 ÷（1−22.44%）× 1.08", "US$/英里", "全美；乘客實付", "Gridwise：司機 2.65 ÷（1−22.44%）", "Derived", "In ph → 傑文斯係數", "線框架給司機 2.65（司機端），本模型換算成乘客端"),
        ("Waymo 舊金山每英里票價", "1a", "5.79（＝19.69 ÷ 3.4 英里）", "US$/英里", "乘客實付，不含小費", "Obi 2026-01；CPUC 每趟 3.4 英里", "Derived", "In rpm26 上緣；滲透 A1", "線框架寫「5.72／公里（≈9.2／英里）」；本模型判斷 5.72 實為每英里：19.69 ÷ 5.72 ＝ 3.44，與 CPUC 每趟 3.4 英里吻合；若為每公里則每趟僅 2.1 英里，與 CPUC 矛盾。請 1a 更正"),
        ("每趟載客英里", "1a", "4.0（3.4–5.5）", "英里", "Waymo 全美平均（加州 3.4 為下緣）", "Findings Press（CPUC）", "Analogy", "In pmi_trip", "線框架未給；tesla_robotaxi 的人類叫車行程長度約 7.1 英里（Gridwise 推算）——Waymo 加州實測顯著較短，兩者不同口徑"),
        ("付費英里占比（1−空駛）", "1a", "2026 59%（55–62%）→ 2035 65%", "%", "總行駛英里", "CPUC 月報", "Derived", "In paid_sh26／paid_sh35", "線框架以人類叫車類比 55–65%；本模型用 CPUC 實測"),
        ("單車取得成本", "1a", "125（100–160），年降 8%", "千美元", "Ojai＋第六代套件", "Morgan Stanley（TNW）", "Interested-party", "In vcost26／vcost_dec", "同線框架"),
        ("車輛經濟壽命", "1a", "6 年（5–8）", "年", "≈ 30–50 萬英里", "—", "Assumed", "In vlife", "線框架以英里表示（30–50 萬）；本模型每車年約 5–6 萬英里，6 年≈ 33 萬英里，落在區間內"),
        ("遠端協助每車英里", "1a", "0.12（0.05–0.20）", "US$/英里", "", "Markey 報告（比例未揭露）", "Assumed", "In cpm_rao", "線框架給比例 1:10–1:40，本模型換成每英里"),
        ("保險、能源、維修清潔、場站每車英里", "1a", "0.20／0.06／0.25／0.22", "US$/英里", "", "未取得", "Assumed", "In cpm_ins／nrg／mnt／depot", "線框架列「未取得」；本模型給區間"),
        ("每車每日付費趟數", "1a", "20（15–26）", "趟", "全美平均", "Waymo 每週趟數 ÷ 車隊", "Derived", "In tpvd → 車隊", "線框架：美國未揭露；中國 21–25"),
        ("Uber 平台抽成", "1a", "22.44%", "%", "中位數", "Gridwise 2026 Q3", "Verified", "In ph 換算", "同線框架"),
        ("AAA 自有車每英里全成本", "1b／1a", "0.772（小轎車 0.559）", "US$/英里", "15,000 英里／年", "AAA 2025-09-16", "Verified", "滲透 A3、每英里經濟 錨", "線框架未列；Andy 22:38 指定"),
        ("Other Bets 營業虧損", "1a", "2026 Q3 20.5 億美元", "十億美元/季", "Alphabet 分部", "Fortune 2026-10-30", "Verified", "In fix26 依據", "同線框架"),
        ("Waymo 每週付費趟數", "1a", "50 萬＋（2026-03）；年底目標 100 萬", "趟/週", "全國", "Waymo", "Interested-party", "In w_trips26", "同線框架"),
        ("Waymo 車隊", "1a", "約 4,000（2026-09）", "台", "", "TechCrunch", "Derived", "In tpvd 依據", "同線框架"),
    ]
    for i, row in enumerate(rows, 3):
        for c, v in enumerate(row, 1): put(ws, i, c, v, wrap=True)
    for c, w in zip("ABCDEFGHI", (30, 8, 26, 12, 24, 30, 14, 30, 50)): ws.column_dimensions[c].width = w


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
            "wk_B": {y: out["B"]["wk"][y] for y in (2026, 2030, 2035)}, "pen26": out["B"]["pen26"]}
    json.dump(summ, open(os.path.join(HERE, "dist", "summary.json"), "w"), ensure_ascii=False, indent=1)
    print(json.dumps(summ, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
