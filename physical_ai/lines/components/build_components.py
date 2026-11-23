# -*- coding: utf-8 -*-
"""線 6 共用零組件單位經濟 v0.1 — build script.

核心等式：零組件需求量(c, y) = Σ_line 出貨量(line, y) × 每台用量(c, line)
          市場規模(c, y)   = Σ_line 出貨量 × 用量 × 價格乘數(c, line) × 單價(c, y)
          單價(c, y)       = 2026 單價 × (1 − 年降幅)^(y − 2026)
          稼動(c, y)       = 需求量 ÷ 產能（只對有產能資料的零組件）

輸出：dist/20261116_共用零組件單位經濟_v0.1.xlsx（全部公式，LibreOffice 重算後 0 錯誤）
用法：python3 build_components.py [--recalc]
"""
import os, sys, subprocess, shutil
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as CL

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "dist", "20261116_共用零組件單位經濟_v0.1.xlsx")

BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
YEL = PatternFill("solid", fgColor="FFF2CC")
GRY = PatternFill("solid", fgColor="EDEDED")
WRAP = Alignment(wrap_text=True, vertical="top")
YEARS = list(range(2026, 2033))
SC = ["低", "基準", "高"]
FW = "線框架"  # 出處前綴

# ---------------------------------------------------------------- lines
LINES = [  # key, 名稱, 所屬線
    ("IND", "工業機器人（傳統）", "2a"),
    ("COB", "協作機器人", "2a"),
    ("HUM", "人形機器人", "2b"),
    ("AMR", "運輸物流機器人（AMR 等）", "3"),
    ("CAR", "輕型車（全部）", "1b"),
    ("CLD", "輕型車（搭 LiDAR）", "1b"),
    ("EMB", "輕型車（搭 EMB 線控煞車）", "1b"),
    ("SBW", "輕型車（搭 SBW 線控轉向）", "1b"),
    ("RT", "Robotaxi 新增車輛", "1a"),
]
LK = [l[0] for l in LINES]

# ---------------------------------------------------------------- drivers (anchors 2026 / 2030 / 2032, each low/base/high)
# key, 名稱, 單位, kind(g=幾何插值, l=線性), [26L,26B,26H], [30L,30B,30H], [32L,32B,32H], 標記, 出處
DRV = [
    ("IND", "全球工業機器人裝機（含協作）", "萬台", "g", [62, 65.5, 68], [72, 83, 95], [75, 90, 110],
     "2026 基準 [Interested-party] IFR 預測；2030/2032 [Assumed]",
     "線2a 框架 §1：IFR 預測 2026 年 65.5 萬、2029 年 80.6 萬；2025 實績 60.3 萬 [Verified]。2030 基準 83 萬＝2029 值外推約 3%/年 [Assumed]"),
    ("COBSH", "協作占工業機器人比重", "%", "l", [0.11, 0.11, 0.12], [0.11, 0.15, 0.20], [0.12, 0.17, 0.24],
     "2026 [Verified] 2025 值延用；之後 [Assumed]", "線2a 框架 §3：IFR 2025 協作 6.6 萬台，占 11%"),
    ("HUM", "全球人形機器人出貨", "萬台", "g", [3, 5, 7], [15, 50, 120], [30, 150, 350],
     "2026 [Interested-party] TrendForce；2030/2032 [Assumed]",
     "線2b 框架 §7-9：TrendForce 2026 全球 >5 萬台；2025 IDC 約 1.8 萬台。高情境參考 Schaeffler 假設 2030 年前全球年產 ≥100 萬台 [Interested-party]、Tesla Fremont 設計產能 100 萬台/年 [Interested-party]"),
    ("AMR", "運輸物流機器人年銷量", "萬台", "g", [13, 14, 15], [20, 28, 38], [24, 38, 55],
     "2026 [Derived] 2025×(1+21%)；之後 [Assumed]",
     "線3 框架 §3：IFR 2025 運輸物流機器人 11.75 萬台，+21%（樣本 238 家，非全產業）"),
    ("CAR", "全球輕型車銷量", "萬輛", "g", [8800, 9000, 9200], [8800, 9300, 9700], [8800, 9500, 10000],
     "[Assumed]（五份框架皆未給全球值）", "未取得：線1b 框架只給中國 2026H1 861 萬輛（城市 NOA 分母）。區間為假設，待對帳"),
    ("LPEN", "車用 LiDAR 搭載率（全球）", "%", "l", [0.035, 0.045, 0.055], [0.08, 0.15, 0.25], [0.10, 0.20, 0.35],
     "2026 [Derived]；之後 [Assumed]",
     "2026：禾賽 ADAS 指引 250–300 萬顆 [Interested-party 公司指引]＋速騰 2025 ADAS 60.9 萬顆 [Verified 財報經 laserfair]＋其他 → 約 400–500 萬顆 ÷ 1.2 顆/車 ÷ 9,000 萬輛。2030：線1b 框架中國城市 NOA 2030 年 62% [Interested-party 灼識] 為上界參考"),
    ("EPEN", "EMB 搭載率（全球）", "%", "l", [0.01, 0.02, 0.03], [0.05, 0.12, 0.20], [0.08, 0.18, 0.30],
     "[Derived] 中國值 × 中國占全球約 1/3；[Assumed]",
     "線1b 框架 §3：中國 EMB 滲透率 2026 約 5%、2030 約 40% [Interested-party 產業媒體]；全球值＝中國 ×（中國占全球銷量約 1/3）＋海外少量 [Assumed]"),
    ("SPEN", "SBW 搭載率（全球）", "%", "l", [0.003, 0.005, 0.01], [0.02, 0.05, 0.10], [0.03, 0.08, 0.15],
     "[Assumed]", "線1b 框架 §4：中國 SBW 國標 2026-07 取消機械連接強制；SBW 滲透率未取得"),
    ("RT", "Robotaxi 年新增車輛（全球）", "萬台", "g", [0.6, 1.0, 1.5], [2, 8, 25], [4, 20, 60],
     "[Derived] 2026 由車隊目標推；之後 [Assumed]",
     "線1a 框架 §2：Waymo 約 4,000 台、小馬目標 2026 年底 3,000+、文遠 1,800+、Tesla 稱 300+ Cybercab；百度未取得。2026 年新增約 0.6–1.5 萬台"),
]
DK = [d[0] for d in DRV]

# ---------------------------------------------------------------- components
# key, 名稱, 數量單位, nondup(1=計入不重複合計), v0.5 次產業代碼
COMP = [
    ("HR", "諧波減速機", "顆", 1, "RED"),
    ("RV", "RV 減速機", "顆", 1, "RED"),
    ("PL", "行星減速機", "顆", 1, "RED"),
    ("PRS", "行星滾柱絲杠", "根", 1, "SCR"),
    ("BS", "滾珠絲杠", "根", 1, "SCR"),
    ("LG", "線性導軌", "組", 1, "SCR"),
    ("SM", "有框伺服馬達", "顆", 1, "MOT"),
    ("FM", "無框力矩馬達", "顆", 1, "MOT"),
    ("CM", "空心杯馬達", "顆", 1, "MOT"),
    ("SD", "伺服驅動（每軸）", "軸", 1, "MOT"),
    ("ENC", "編碼器", "顆", 1, "MOT"),
    ("F6", "六維力感測器", "顆", 1, "FRC"),
    ("TAC", "觸覺感測（指尖陣列）", "片", 1, "FRC"),
    ("PNE", "氣動元件（每台美元內容）", "美元", 1, "PNE"),
    ("FAS", "工業感測與機器視覺（每台美元內容）", "美元", 1, "FAS"),
    ("HAND", "靈巧手（模組；與空心杯、觸覺重複，不計入合計）", "隻", 0, "HAND"),
    ("LDR", "LiDAR", "顆", 1, "LDR"),
    ("CAM", "相機模組（機器人與 Robotaxi）", "顆", 1, "VIS"),
    ("CAMV", "車用相機（全部輕型車；既有需求為主，不計入合計）", "顆", 0, "VIS"),
]
CK = [c[0] for c in COMP]

# usage[c] = {line: (low, base, high)} ; missing = 0
U = {
    "HR": {"IND": (1.5, 2, 3), "COB": (6, 6, 7), "HUM": (6, 10, 16)},
    "RV": {"IND": (3, 3.5, 4)},
    "PL": {"IND": (0, 0.5, 1), "HUM": (2, 8, 14), "AMR": (2, 2, 4)},
    "PRS": {"HUM": (4, 8, 14)},
    "BS": {"IND": (0.2, 0.3, 0.5), "HUM": (0, 0, 6), "AMR": (0, 0.5, 1), "EMB": (4, 4, 4), "SBW": (1, 1, 1)},
    "LG": {"IND": (0.1, 0.3, 0.5)},
    "SM": {"IND": (6, 6, 6), "AMR": (2, 2, 4)},
    "FM": {"COB": (6, 6, 7), "HUM": (28, 28, 30)},
    "CM": {"HUM": (10, 12, 24)},
    "SD": {"IND": (6, 6, 6), "COB": (6, 6, 7), "HUM": (28, 28, 30), "AMR": (2, 2, 4)},
    "ENC": {"IND": (6, 6, 12), "COB": (12, 12, 14), "HUM": (28, 56, 60), "AMR": (2, 2, 4)},
    "F6": {"IND": (0.02, 0.05, 0.1), "COB": (0.1, 0.3, 0.5), "HUM": (2, 2, 4)},
    "TAC": {"HUM": (0, 4, 10)},
    "PNE": {"IND": (200, 500, 1500), "COB": (100, 300, 800)},
    "FAS": {"IND": (300, 800, 2000), "COB": (200, 500, 1500)},
    "HAND": {"HUM": (1.5, 2, 2)},
    "LDR": {"HUM": (0, 0.3, 1), "AMR": (0.5, 1, 2), "CLD": (1, 1.2, 1.5), "RT": (1, 4, 6)},
    "CAM": {"HUM": (2, 4, 8), "AMR": (1, 2, 4), "RT": (8, 11, 13)},
    "CAMV": {"CAR": (2, 3, 5)},
}
U_SRC = {
    "HR": ("[Analogy]／[Interested-party]", "傳統 6 軸：腕部 2–3 顆諧波（線2a §3 物理註記）；以 2024 全球工業用諧波消費 121.87 萬台 [Interested-party 中國報告網] 校準為 2 顆；協作 6–7 顆（線2a）；人形：Optimus 類全尺寸 14–16 顆（線2b §3.2）；出貨組合平均取 10（6–16），因中國中小型機多用行星減速 [Analogy，待核]"),
    "RV": ("[Analogy]", "傳統 6 軸底座與大臂用 RV，3–4 顆（線2a §3 物理註記）"),
    "PL": ("[Assumed]", "人形以行星替代部分諧波：出貨組合平均 8（2–14）（線2b §3.2 替代路線；中國中小型機為主）；AMR 驅動輪每輪 1 顆"),
    "PRS": ("[Interested-party]", "Optimus 類全尺寸 12–14 根（線2b §3.2，MS 二手轉述；36Kr 經 humanoid.guide 約 14 根）；出貨組合平均取 8（4–14）：中國中小型機多為旋轉關節、不用或少用滾柱絲杠 [Analogy，待核]；高值 14＝全部出貨為 Optimus 類"),
    "BS": ("[Analogy]", "EMB 每輪 1 根（4 根/車）、SBW 齒條致動器 1 根 [Analogy：線控致動器結構]；SCARA 滾珠花鍵；人形高值 6 根＝以滾珠絲杠替代滾柱絲杠的設計"),
    "LG": ("[Assumed]", "工業機器人第七軸／龍門應用平均；線性導軌主需求在工具機與半導體設備，不在本模型五線"),
    "SM": ("[Analogy]", "傳統 6 軸每軸 1 顆；AMR 每驅動輪 1 顆"),
    "FM": ("[Interested-party]", "協作一體化關節每關節 1 顆（線2a）；人形每關節 1 顆 28–30（線2b §3.2）"),
    "CM": ("[Interested-party]", "人形靈巧手內 10–24 顆（線2b §3.2，TechBuzz China 追蹤）"),
    "SD": ("[Analogy]", "每軸／每關節 1 組"),
    "ENC": ("[Interested-party]", "人形每關節 2 顆（馬達端＋輸出端，線2b §3.2）；協作雙編碼器；低值＝單編碼器設計"),
    "F6": ("[Interested-party]／[Assumed]", "人形 2–4 顆（線2b，X.PIN）；協作與工業為末端選配搭載率 [Assumed]"),
    "TAC": ("[Assumed]", "每台指尖觸覺陣列 0–10 片（兩手各 5 指）；目前多數機型未搭"),
    "PNE": ("[Assumed]", "每台末端夾爪與周邊氣動美元內容；氣動主需求為一般工廠自動化，非本模型五線"),
    "FAS": ("[Assumed]", "每台工作站視覺與感測美元內容；同上"),
    "HAND": ("[Interested-party]", "每台 2 隻（線2b）；低值 1.5＝部分機型用夾爪"),
    "LDR": ("[Analogy]", "Robotaxi：Waymo 第 6 代 4 顆 [Interested-party 公司]、Tesla 0 顆；乘用車搭載車平均 1–1.5 顆；AMR 1 顆；人形部分機型（如宇樹 G1）頭部 1 顆"),
    "CAM": ("[Analogy]", "Robotaxi：Waymo 第 6 代 13 顆、Tesla 約 8 顆；人形頭部 2–8 顆；AMR 1–4 顆"),
    "CAMV": ("[Assumed]", "全球輕型車平均 2–5 顆（含倒車與環景）；多為既有需求，故不計入不重複合計"),
}

# price (USD per unit, 2026) low/base/high  ; annual decline low/base/high ; tag ; source
P = {
    "HR": ((300, 350, 450), (0.03, 0.06, 0.10), "[Derived]／[Analogy]",
           "上界：哈默納科 FY03/27 營收指引 745 億日圓 ÷（2024 全球工業用諧波 121.87 萬台 × 85% 份額）≈ 7.2 萬日圓 ≈ 480 美元（營收含半導體設備與總成，故為上界）[Derived]；中國廠便宜 30–50%（線2b）"),
    "RV": ((700, 900, 1200), (0.02, 0.04, 0.06), "[Assumed]", "未取得 Nabtesco 精密減速機分部均價；假設為諧波 2–3 倍"),
    "PL": ((50, 80, 120), (0.03, 0.05, 0.08), "[Analogy]", "線2b：行星約為諧波的 1/5"),
    "PRS": ((190, 750, 2000), (0.08, 0.15, 0.20), "[Interested-party]（三口徑互相矛盾，見 PRS 口徑表）",
            "基準＝覓途諮詢占比推算；低＝中國量產價（五洲新春 2025 均價）；高＝MS 推算中值。見本頁「行星滾柱絲杠單價口徑」"),
    "BS": ((20, 30, 50), (0.02, 0.03, 0.05), "[Assumed]", "未取得；車用與小型精密滾珠絲杠"),
    "LG": ((60, 100, 150), (0.01, 0.02, 0.03), "[Assumed]", "未取得"),
    "SM": ((80, 120, 200), (0.02, 0.03, 0.05), "[Assumed]", "未取得"),
    "FM": ((100, 180, 300), (0.04, 0.08, 0.12), "[Derived]／[Assumed]", "線2b：若 Optimus BOM 5–6 萬美元、馬達占 16% → 290–340 美元/顆（上界）[Derived]；中國量產更低"),
    "CM": ((20, 50, 150), (0.04, 0.08, 0.12), "[Interested-party]", "線2b：100–4,000 人民幣（國產低階 vs Maxon/Faulhaber），TechBuzz China"),
    "SD": ((50, 90, 150), (0.03, 0.05, 0.08), "[Assumed]", "未取得"),
    "ENC": ((8, 20, 50), (0.03, 0.05, 0.08), "[Assumed]", "未取得（磁編碼器 vs 光學絕對式）"),
    "F6": ((700, 1500, 3000), (0.10, 0.18, 0.25), "[Interested-party]",
           "線2b：5,000–30,000 人民幣（X.PIN）；賽迪 2030 中國人形 120 萬顆、60 億元 → 隱含約 5,000 人民幣/顆 [Derived，經同花順 2026-08-31]"),
    "TAC": ((30, 80, 200), (0.08, 0.15, 0.20), "[Assumed]", "未取得"),
    "PNE": ((1, 1, 1), (0, 0, 0), "[Derived]", "美元內容，單價恆為 1"),
    "FAS": ((1, 1, 1), (0, 0, 0), "[Derived]", "美元內容，單價恆為 1"),
    "HAND": ((2000, 3500, 4750), (0.08, 0.15, 0.20), "[Interested-party]",
             "MS：兩隻合計約 9,500 美元 → 4,750/隻；中國量產目標單隻約 1.5 萬人民幣（覓途）（皆經線2b）"),
    "LDR": ((150, 180, 220), (0.06, 0.10, 0.15), "[Verified]／[Interested-party]",
            "禾賽 2026Q2 綜合 ASP 1,297 人民幣（年減 35%）÷ 7.1 ≈ 183 美元（長橋海豚研究 2026-08-19 整理財報）"),
    "CAM": ((15, 25, 40), (0.02, 0.04, 0.06), "[Assumed]", "未取得"),
    "CAMV": ((15, 25, 40), (0.02, 0.04, 0.06), "[Assumed]", "未取得"),
}
# price multiplier by (comp, line), default 1 — (low, base, high) for sensitivity; base used in calc
MULT = {("LDR", "RT"): (2, 3, 5), ("CAM", "RT"): (1, 1.5, 2)}
MULT_SRC = "Robotaxi 用長距與高解析感測，單價高於乘用車主流款 [Assumed]"

# supply: comp -> (cap2026 low/base/high 萬顆, growth low/base/high, tag, source)
SUP = {
    "HR": ((200, 250, 350), (0.10, 0.15, 0.30), "[Assumed]",
           "未取得哈默納科、綠的諧波產能數字。錨：2024 全球工業用諧波消費 121.87 萬台 [Interested-party 中國報告網]；產能取其約 2 倍"),
    "RV": ((200, 230, 280), (0.03, 0.05, 0.08), "[Assumed]", "未取得 Nabtesco 產能；Nabtesco 全球份額 >53% [Interested-party]"),
    "PRS": ((30, 60, 120), (0.30, 0.60, 1.00), "[Assumed]",
            "錨：五洲新春 2025 出貨約 8.7 萬根；新劍傳動據報建 100 萬根/年產線；雙林 2026-01 宣布量產（皆 [Interested-party]，經 humanoid.guide）"),
    "LDR": ((700, 900, 1200), (0.15, 0.25, 0.35), "[Interested-party]",
            "禾賽 2026 產能擴至 >400 萬顆（公司，經 laserfair）；速騰、華為、圖達通等未取得 → 全球合計 [Assumed]"),
}

# PRS price口徑 table
PRS_PTS = [
    ("MS 推算（中值）", 2000, "[Interested-party] Morgan Stanley（二手轉述，線2b §3.2）：1,350–2,700 美元/根"),
    ("覓途諮詢占比推算", None, "[Derived] 覓途：絲杠占人形 BOM 19%（線2b）× BOM 55,000 美元（MS）÷ 14 根"),
    ("中國量產價（五洲新春 2025 均價）", None, "[Interested-party] 約 1,360 人民幣，年減 39%（humanoid.guide 轉述產業網站）÷ 匯率"),
    ("中國最低報價（雙林）", None, "[Interested-party] 499 人民幣（China Daily 經 humanoid.guide，2026-01）÷ 匯率"),
]


def addr(ws, r, c, absolute=True):
    a = f"${CL(c)}${r}" if absolute else f"{CL(c)}{r}"
    return f"'{ws.title}'!{a}"


def hdr(ws, r, vals, c0=1):
    for i, v in enumerate(vals):
        x = ws.cell(r, c0 + i, v); x.font = BOLD; x.fill = HDR; x.alignment = WRAP


def inp(cell, judg=False):
    cell.font = BLUE
    if judg:
        cell.fill = YEL


def build():
    wb = openpyxl.Workbook()
    ws0 = wb.active; ws0.title = "說明"
    wi = wb.create_sheet("輸入")
    wc = wb.create_sheet("計算")
    wsn = wb.create_sheet("敏感度")
    wsc = wb.create_sheet("情境")
    wsd = wb.create_sheet("供需")
    wsh = wb.create_sheet("份額")
    wdl = wb.create_sheet("下游取數")
    wr = wb.create_sheet("對帳")
    wt = wb.create_sheet("追蹤指標")

    # ================================================================ 輸入
    R = {}  # registry of addresses
    r = 1
    wi.cell(r, 1, "輸入（藍字＝輸入；黃底＝判斷，可改；每列附標記與出處）").font = BOLD; r += 2
    wi.cell(r, 1, "A. 全域").font = BOLD; r += 1
    hdr(wi, r, ["變數", "值", "單位", "標記", "出處"]); r += 1
    for k, n, v, u, t, s in [("FXCNY", "美元兌人民幣", 7.1, "CNY/USD", "[Assumed]", "區間 6.9–7.3；僅用於換算"),
                             ("FXJPY", "美元兌日圓", 150, "JPY/USD", "[Assumed]", "區間 140–160；僅用於換算")]:
        wi.cell(r, 1, n); c = wi.cell(r, 2, v); inp(c, True); wi.cell(r, 3, u); wi.cell(r, 4, t); wi.cell(r, 5, s)
        R[k] = addr(wi, r, 2); r += 1
    r += 1
    # ---- B drivers
    wi.cell(r, 1, "B. 各線出貨量錨點（2026／2030／2032；其間插值：台數幾何、比率線性）— 對帳頁列出處").font = BOLD; r += 1
    hdr(wi, r, ["鍵", "變數", "單位", "插值", "2026 低", "2026 基準", "2026 高", "2030 低", "2030 基準", "2030 高",
                "2032 低", "2032 基準", "2032 高", "標記", "出處"]); r += 1
    R["DRV"] = {}
    for d in DRV:
        k, n, u, kind, a26, a30, a32, tag, src = d
        wi.cell(r, 1, k); wi.cell(r, 2, n); wi.cell(r, 3, u); wi.cell(r, 4, "幾何" if kind == "g" else "線性")
        vals = a26 + a30 + a32
        for j, v in enumerate(vals):
            c = wi.cell(r, 5 + j, v); inp(c, True)
            if u == "%": c.number_format = "0.0%"
        wi.cell(r, 14, tag); wi.cell(r, 15, src).alignment = WRAP
        R["DRV"][k] = {(yr, s): addr(wi, r, 5 + yi * 3 + s) for yi, yr in enumerate((2026, 2030, 2032)) for s in range(3)}
        R["DRV"][k]["kind"] = kind
        r += 1
    r += 1
    # ---- C usage blocks
    R["U"] = {}
    for s in range(3):
        wi.cell(r, 1, f"C{s+1}. 每台／每車用量矩陣 — {SC[s]}情境（數量單位見零組件；PNE、FAS 為每台美元內容）").font = BOLD; r += 1
        hdr(wi, r, ["鍵", "零組件", "單位"] + [f"{l[1]}（線{l[2]}）" for l in LINES] + ["標記", "出處"]); r += 1
        for c in COMP:
            ck = c[0]
            wi.cell(r, 1, ck); wi.cell(r, 2, c[1]); wi.cell(r, 3, c[2])
            for j, lk in enumerate(LK):
                v = U.get(ck, {}).get(lk, (0, 0, 0))[s]
                cell = wi.cell(r, 4 + j, v); inp(cell, v != 0)
            wi.cell(r, 4 + len(LK), U_SRC[ck][0]); wi.cell(r, 5 + len(LK), U_SRC[ck][1]).alignment = WRAP
            R["U"][(ck, s)] = (r, 4, 4 + len(LK) - 1)
            r += 1
        r += 1
    # ---- D PRS口徑 table (before prices so prices can link)
    wi.cell(r, 1, "D. 行星滾柱絲杠單價口徑（互相矛盾，做成敏感度；美元/根，2026）").font = BOLD; r += 1
    hdr(wi, r, ["口徑", "美元/根", "說明與標記"]); r += 1
    R["PRSPT"] = []
    # helper input cells for derived
    wi.cell(r, 5, "覓途絲杠占 BOM"); c = wi.cell(r, 6, 0.19); inp(c); R["MT_SH"] = addr(wi, r, 6)
    wi.cell(r + 1, 5, "MS 人形 BOM（美元）"); c = wi.cell(r + 1, 6, 55000); inp(c); R["MS_BOM"] = addr(wi, r + 1, 6)
    wi.cell(r + 2, 5, "Optimus 類每台絲杠根數"); c = wi.cell(r + 2, 6, 14); inp(c); R["PRS_N"] = addr(wi, r + 2, 6)
    wi.cell(r + 3, 5, "五洲新春 2025 均價（人民幣）"); c = wi.cell(r + 3, 6, 1360); inp(c); R["WZ"] = addr(wi, r + 3, 6)
    wi.cell(r + 4, 5, "雙林報價（人民幣）"); c = wi.cell(r + 4, 6, 499); inp(c); R["SL"] = addr(wi, r + 4, 6)
    for i, (n, v, s) in enumerate(PRS_PTS):
        wi.cell(r, 1, n)
        if i == 0:
            c = wi.cell(r, 2, v); inp(c)
        elif i == 1:
            wi.cell(r, 2, f"={R['MT_SH']}*{R['MS_BOM']}/{R['PRS_N']}")
        elif i == 2:
            wi.cell(r, 2, f"={R['WZ']}/{R['FXCNY']}")
        else:
            wi.cell(r, 2, f"={R['SL']}/{R['FXCNY']}")
        wi.cell(r, 2).number_format = "#,##0"
        wi.cell(r, 3, s).alignment = WRAP
        R["PRSPT"].append(addr(wi, r, 2)); r += 1
    r += 2
    # ---- E prices
    wi.cell(r, 1, "E. 2026 單價與年降幅（美元；低／高只用於敏感度，計算一律用基準）").font = BOLD; r += 1
    hdr(wi, r, ["鍵", "零組件", "單位", "2026 單價 低", "2026 單價 基準", "2026 單價 高", "年降幅 低", "年降幅 基準", "年降幅 高",
                "計入不重複合計（1/0）", "v0.5 次產業", "標記", "出處"]); r += 1
    R["P"] = {}
    for c in COMP:
        ck = c[0]; (p0, d, tag, src) = P[ck]
        wi.cell(r, 1, ck); wi.cell(r, 2, c[1]); wi.cell(r, 3, f"美元/{c[2]}")
        for j in range(3):
            if ck == "PRS":
                wi.cell(r, 4 + j, "=" + R["PRSPT"][[2, 1, 0][j]])
            else:
                cell = wi.cell(r, 4 + j, p0[j]); inp(cell, "Assumed" in tag or "Analogy" in tag)
            wi.cell(r, 4 + j).number_format = "#,##0"
            cell = wi.cell(r, 7 + j, d[j]); inp(cell, True); cell.number_format = "0%"
        cell = wi.cell(r, 10, c[3]); inp(cell)
        wi.cell(r, 11, c[4]); wi.cell(r, 12, tag); wi.cell(r, 13, src).alignment = WRAP
        R["P"][ck] = {"p0": [addr(wi, r, 4 + j) for j in range(3)], "d": [addr(wi, r, 7 + j) for j in range(3)],
                      "nd": addr(wi, r, 10), "row": r}
        r += 1
    r += 1
    # ---- F multipliers
    wi.cell(r, 1, "F. 價格乘數（零組件 × 線；預設 1；計算用基準）").font = BOLD; r += 1
    hdr(wi, r, ["鍵", "零組件"] + [l[1] for l in LINES] + ["說明"]); r += 1
    R["M"] = {}
    for c in COMP:
        ck = c[0]
        wi.cell(r, 1, ck); wi.cell(r, 2, c[1])
        for j, lk in enumerate(LK):
            v = MULT.get((ck, lk), (1, 1, 1))[1]
            cell = wi.cell(r, 3 + j, v); inp(cell, (ck, lk) in MULT)
        if any(k[0] == ck for k in MULT):
            wi.cell(r, 3 + len(LK), MULT_SRC + "；敏感度低／高：" + "、".join(f"{k[1]} {v[0]}–{v[2]}" for k, v in MULT.items() if k[0] == ck))
        R["M"][ck] = (r, 3, 3 + len(LK) - 1)
        r += 1
    r += 1
    # ---- G supply
    wi.cell(r, 1, "G. 供給（2026 有效產能，萬顆／根；年增率）— 只對有錨點的零組件").font = BOLD; r += 1
    hdr(wi, r, ["鍵", "零組件", "2026 產能 低", "2026 產能 基準", "2026 產能 高", "年增 低", "年增 基準", "年增 高", "標記", "出處"]); r += 1
    R["S"] = {}
    for ck, (cap, g, tag, src) in SUP.items():
        wi.cell(r, 1, ck); wi.cell(r, 2, dict((c[0], c[1]) for c in COMP)[ck])
        for j in range(3):
            cell = wi.cell(r, 3 + j, cap[j]); inp(cell, True)
            cell = wi.cell(r, 6 + j, g[j]); inp(cell, True); cell.number_format = "0%"
        wi.cell(r, 9, tag); wi.cell(r, 10, src).alignment = WRAP
        R["S"][ck] = {"cap": [addr(wi, r, 3 + j) for j in range(3)], "g": [addr(wi, r, 6 + j) for j in range(3)]}
        r += 1
    wi.column_dimensions["A"].width = 8; wi.column_dimensions["B"].width = 30
    for col in range(3, 14):
        wi.column_dimensions[CL(col)].width = 12
    wi.column_dimensions["O"].width = 60; wi.column_dimensions["M"].width = 40

    # ================================================================ 計算
    C = {}
    r = 1
    wc.cell(r, 1, "計算（全部公式）：需求量＝Σ 線出貨量 × 用量；市場規模＝Σ 線出貨量 × 用量 × 乘數 × 單價").font = BOLD; r += 2
    # price path
    wc.cell(r, 1, "1. 單價路徑（美元，基準）＝ 2026 單價 ×（1 − 年降幅）^(年 − 2026)").font = BOLD; r += 1
    hdr(wc, r, ["鍵", "零組件"] + YEARS); r += 1
    C["price"] = {}
    for c in COMP:
        ck = c[0]
        wc.cell(r, 1, ck); wc.cell(r, 2, c[1])
        for j, y in enumerate(YEARS):
            wc.cell(r, 3 + j, f"={R['P'][ck]['p0'][1]}*(1-{R['P'][ck]['d'][1]})^({y}-2026)").number_format = "#,##0.0"
            C["price"][(ck, y)] = addr(wc, r, 3 + j)
        r += 1
    r += 1
    C["drv"] = {}; C["vol"] = {}; C["units"] = {}; C["mkt"] = {}; C["nd"] = {}
    for s in range(3):
        # driver paths
        wc.cell(r, 1, f"2.{s+1} 驅動變數路徑 — {SC[s]}情境").font = BOLD; r += 1
        hdr(wc, r, ["鍵", "變數"] + YEARS); r += 1
        for d in DRV:
            k = d[0]; a = R["DRV"][k]
            wc.cell(r, 1, k); wc.cell(r, 2, d[1])
            for j, y in enumerate(YEARS):
                if y in (2026, 2030, 2032):
                    f = f"={a[(y, s)]}"
                elif d[3] == "g":
                    if y < 2030:
                        f = f"={a[(2026, s)]}*({a[(2030, s)]}/{a[(2026, s)]})^(({y}-2026)/4)"
                    else:
                        f = f"={a[(2030, s)]}*({a[(2032, s)]}/{a[(2030, s)]})^(({y}-2030)/2)"
                else:
                    if y < 2030:
                        f = f"={a[(2026, s)]}+({a[(2030, s)]}-{a[(2026, s)]})*({y}-2026)/4"
                    else:
                        f = f"={a[(2030, s)]}+({a[(2032, s)]}-{a[(2030, s)]})*({y}-2030)/2"
                cell = wc.cell(r, 3 + j, f)
                cell.number_format = "0.0%" if d[2] == "%" else "#,##0.0"
                C["drv"][(k, s, y)] = addr(wc, r, 3 + j, absolute=False)
            r += 1
        r += 1
        # volumes: rows=years, cols=lines
        wc.cell(r, 1, f"3.{s+1} 各線出貨量（萬台／萬輛）— {SC[s]}情境（列＝年、欄＝線）").font = BOLD; r += 1
        hdr(wc, r, ["年", ""] + [f"{l[1]}" for l in LINES]); r += 1
        for y in YEARS:
            wc.cell(r, 1, y)
            D = lambda k: C["drv"][(k, s, y)]
            forms = {
                "IND": f"={D('IND')}*(1-{D('COBSH')})", "COB": f"={D('IND')}*{D('COBSH')}", "HUM": f"={D('HUM')}",
                "AMR": f"={D('AMR')}", "CAR": f"={D('CAR')}", "CLD": f"={D('CAR')}*{D('LPEN')}",
                "EMB": f"={D('CAR')}*{D('EPEN')}", "SBW": f"={D('CAR')}*{D('SPEN')}", "RT": f"={D('RT')}"}
            for j, lk in enumerate(LK):
                wc.cell(r, 3 + j, forms[lk]).number_format = "#,##0.00"
                C["vol"][(s, y, lk)] = addr(wc, r, 3 + j)
            C["vol"][(s, y)] = f"'計算'!${CL(3)}${r}:${CL(2 + len(LK))}${r}"
            r += 1
        r += 1
        # units
        wc.cell(r, 1, f"4.{s+1} 零組件需求量（萬顆／萬根；PNE、FAS 為萬美元）— {SC[s]}情境").font = BOLD; r += 1
        hdr(wc, r, ["鍵", "零組件"] + YEARS); r += 1
        for c in COMP:
            ck = c[0]; ur = R["U"][(ck, s)]
            urng = f"'輸入'!${CL(ur[1])}${ur[0]}:${CL(ur[2])}${ur[0]}"
            wc.cell(r, 1, ck); wc.cell(r, 2, c[1])
            for j, y in enumerate(YEARS):
                wc.cell(r, 3 + j, f"=SUMPRODUCT({urng},{C['vol'][(s, y)]})").number_format = "#,##0.0"
                C["units"][(s, ck, y)] = addr(wc, r, 3 + j)
            r += 1
        r += 1
        # market
        wc.cell(r, 1, f"5.{s+1} 市場規模（百萬美元，五線歸因口徑）— {SC[s]}情境").font = BOLD; r += 1
        hdr(wc, r, ["鍵", "零組件"] + YEARS); r += 1
        first = r
        for c in COMP:
            ck = c[0]; ur = R["U"][(ck, s)]; mr = R["M"][ck]
            urng = f"'輸入'!${CL(ur[1])}${ur[0]}:${CL(ur[2])}${ur[0]}"
            mrng = f"'輸入'!${CL(mr[1])}${mr[0]}:${CL(mr[2])}${mr[0]}"
            wc.cell(r, 1, ck); wc.cell(r, 2, c[1])
            for j, y in enumerate(YEARS):
                # 萬顆 × 美元 / 100 = 百萬美元
                wc.cell(r, 3 + j, f"=SUMPRODUCT({urng},{mrng},{C['vol'][(s, y)]})*{C['price'][(ck, y)]}/100").number_format = "#,##0"
                C["mkt"][(s, ck, y)] = addr(wc, r, 3 + j)
            r += 1
        last = r - 1
        wc.cell(r, 2, "不重複合計（不含靈巧手模組）").font = BOLD
        ndr = R["P"][CK[0]]["row"]
        ndrng = f"'輸入'!$J${R['P'][CK[0]]['row']}:$J${R['P'][CK[-1]]['row']}"
        for j, y in enumerate(YEARS):
            wc.cell(r, 3 + j, f"=SUMPRODUCT({ndrng},{CL(3+j)}{first}:{CL(3+j)}{last})").number_format = "#,##0"
            C["nd"][(s, y)] = addr(wc, r, 3 + j)
        r += 2
    # 6. 2030 base by-line matrix
    Y = 2030
    wc.cell(r, 1, "6. 2030 基準：市場規模依線拆分（百萬美元）— 供敏感度與人形占比").font = BOLD; r += 1
    hdr(wc, r, ["鍵", "零組件"] + [l[1] for l in LINES] + ["合計", "人形占比"]); r += 1
    C["bl"] = {}; first = r
    for c in COMP:
        ck = c[0]; ur = R["U"][(ck, 1)]; mr = R["M"][ck]
        wc.cell(r, 1, ck); wc.cell(r, 2, c[1])
        for j, lk in enumerate(LK):
            u_ = addr(wi, ur[0], ur[1] + j); m_ = addr(wi, mr[0], mr[1] + j)
            wc.cell(r, 3 + j, f"={u_}*{m_}*{C['vol'][(1, Y, lk)]}*{C['price'][(ck, Y)]}/100").number_format = "#,##0"
            C["bl"][(ck, lk)] = addr(wc, r, 3 + j)
        tc = 3 + len(LK)
        wc.cell(r, tc, f"=SUM({CL(3)}{r}:{CL(tc-1)}{r})").number_format = "#,##0"
        C["bl"][(ck, "TOT")] = addr(wc, r, tc)
        wc.cell(r, tc + 1, f"=IF({CL(tc)}{r}=0,0,{CL(3+LK.index('HUM'))}{r}/{CL(tc)}{r})").number_format = "0%"
        C["bl"][(ck, "HUMSH")] = addr(wc, r, tc + 1)
        r += 1
    last = r - 1
    wc.cell(r, 2, "不重複合計").font = BOLD
    ndrng = f"'輸入'!$J${R['P'][CK[0]]['row']}:$J${R['P'][CK[-1]]['row']}"
    for j in range(len(LK) + 1):
        col = CL(3 + j)
        wc.cell(r, 3 + j, f"=SUMPRODUCT({ndrng},{col}{first}:{col}{last})").number_format = "#,##0"
        key = LK[j] if j < len(LK) else "TOT"
        C["bl"][("ND", key)] = addr(wc, r, 3 + j)
    tc = 3 + len(LK)
    wc.cell(r, tc + 1, f"={CL(3+LK.index('HUM'))}{r}/{CL(tc)}{r}").number_format = "0%"
    r += 1
    wc.cell(r, 2, "檢核：本表合計 − 5.2 節 2030 不重複合計（應為 0）")
    wc.cell(r, 3, f"={C['bl'][('ND','TOT')]}-{C['nd'][(1, Y)]}").number_format = "0.000"
    C["chk1"] = addr(wc, r, 3); r += 2
    # 7. humanoid content per unit
    wc.cell(r, 1, "7. 人形每台零組件內容（出貨組合平均，美元/台，基準用量 × 基準單價；不含靈巧手模組、運算、電池、結構）— 對照線2b BOM").font = BOLD; r += 1
    hdr(wc, r, ["", ""] + YEARS); r += 1
    wc.cell(r, 2, "每台零組件內容")
    hcol = LK.index("HUM")
    C["humc"] = {}
    for j, y in enumerate(YEARS):
        terms = []
        for c in COMP:
            ck = c[0]; ur = R["U"][(ck, 1)]; mr = R["M"][ck]
            terms.append(f"{R['P'][ck]['nd']}*{addr(wi, ur[0], ur[1]+hcol)}*{addr(wi, mr[0], mr[1]+hcol)}*{C['price'][(ck, y)]}")
        wc.cell(r, 3 + j, "=" + "+".join(terms)).number_format = "#,##0"
        C["humc"][y] = addr(wc, r, 3 + j)
    r += 1
    wc.cell(r, 2, "其中：行星滾柱絲杠")
    for j, y in enumerate(YEARS):
        ur = R["U"][("PRS", 1)]
        wc.cell(r, 3 + j, f"={addr(wi, ur[0], ur[1]+hcol)}*{C['price'][('PRS', y)]}").number_format = "#,##0"
        C.setdefault("humprs", {})[y] = addr(wc, r, 3 + j)
    r += 1
    wc.cell(r, 2, "絲杠占每台零組件內容")
    for j, y in enumerate(YEARS):
        wc.cell(r, 3 + j, f"={C['humprs'][y]}/{C['humc'][y]}").number_format = "0%"
    r += 1
    wc.column_dimensions["B"].width = 34
    for col in range(3, 14):
        wc.column_dimensions[CL(col)].width = 13

    # ================================================================ 敏感度
    T = C["nd"][(1, Y)]
    def uaddr(ck, lk, s):
        ur = R["U"][(ck, s)]; return addr(wi, ur[0], ur[1] + LK.index(lk))
    def drv(k, s): return R["DRV"][k][(Y, s)]
    def colsum(lks): return "(" + "+".join(C["bl"][("ND", lk)] for lk in lks) + ")"
    def ratio_pd(ck, s):
        return f"((1-{R['P'][ck]['d'][s]})/(1-{R['P'][ck]['d'][1]}))^4"
    SENS_TOTAL = [
        # name, low label expr, high label expr, affected base value, ratio_low, ratio_high, tag
        ("人形 2030 出貨量（萬台）", drv("HUM", 0), drv("HUM", 2), colsum(["HUM"]), f"{drv('HUM',0)}/{drv('HUM',1)}", f"{drv('HUM',2)}/{drv('HUM',1)}", "[Assumed]"),
        ("行星滾柱絲杠 2026 單價口徑（中國量產 ↔ MS）", R["P"]["PRS"]["p0"][0], R["P"]["PRS"]["p0"][2], C["bl"][("PRS", "TOT")],
         f"{R['P']['PRS']['p0'][0]}/{R['P']['PRS']['p0'][1]}", f"{R['P']['PRS']['p0'][2]}/{R['P']['PRS']['p0'][1]}", "[Interested-party] 三口徑矛盾"),
        ("行星滾柱絲杠年降幅（快 ↔ 慢）", R["P"]["PRS"]["d"][2], R["P"]["PRS"]["d"][0], C["bl"][("PRS", "TOT")], ratio_pd("PRS", 2), ratio_pd("PRS", 0), "[Assumed]"),
        ("人形每台滾柱絲杠根數", uaddr("PRS", "HUM", 0), uaddr("PRS", "HUM", 2), C["bl"][("PRS", "HUM")],
         f"{uaddr('PRS','HUM',0)}/{uaddr('PRS','HUM',1)}", f"{uaddr('PRS','HUM',2)}/{uaddr('PRS','HUM',1)}", "[Interested-party]"),
        ("工業機器人 2030 裝機（萬台）", drv("IND", 0), drv("IND", 2), colsum(["IND", "COB"]), f"{drv('IND',0)}/{drv('IND',1)}", f"{drv('IND',2)}/{drv('IND',1)}", "[Assumed]"),
        ("車用 LiDAR 2030 搭載率", drv("LPEN", 0), drv("LPEN", 2), colsum(["CLD"]), f"{drv('LPEN',0)}/{drv('LPEN',1)}", f"{drv('LPEN',2)}/{drv('LPEN',1)}", "[Assumed]"),
        ("LiDAR 年降幅（快 ↔ 慢）", R["P"]["LDR"]["d"][2], R["P"]["LDR"]["d"][0], C["bl"][("LDR", "TOT")], ratio_pd("LDR", 2), ratio_pd("LDR", 0), "[Assumed]"),
        ("Robotaxi 2030 新增車輛（萬台）", drv("RT", 0), drv("RT", 2), colsum(["RT"]), f"{drv('RT',0)}/{drv('RT',1)}", f"{drv('RT',2)}/{drv('RT',1)}", "[Assumed]"),
        ("全球輕型車 2030 銷量（萬輛）", drv("CAR", 0), drv("CAR", 2), colsum(["CAR", "CLD", "EMB", "SBW"]), f"{drv('CAR',0)}/{drv('CAR',1)}", f"{drv('CAR',2)}/{drv('CAR',1)}", "[Assumed]"),
        ("EMB 2030 搭載率", drv("EPEN", 0), drv("EPEN", 2), colsum(["EMB"]), f"{drv('EPEN',0)}/{drv('EPEN',1)}", f"{drv('EPEN',2)}/{drv('EPEN',1)}", "[Derived]／[Assumed]"),
        ("AMR 2030 銷量（萬台）", drv("AMR", 0), drv("AMR", 2), colsum(["AMR"]), f"{drv('AMR',0)}/{drv('AMR',1)}", f"{drv('AMR',2)}/{drv('AMR',1)}", "[Assumed]"),
        ("諧波減速機 2026 單價", R["P"]["HR"]["p0"][0], R["P"]["HR"]["p0"][2], C["bl"][("HR", "TOT")], f"{R['P']['HR']['p0'][0]}/{R['P']['HR']['p0'][1]}", f"{R['P']['HR']['p0'][2]}/{R['P']['HR']['p0'][1]}", "[Derived]／[Analogy]"),
        ("RV 減速機 2026 單價", R["P"]["RV"]["p0"][0], R["P"]["RV"]["p0"][2], C["bl"][("RV", "TOT")], f"{R['P']['RV']['p0'][0]}/{R['P']['RV']['p0'][1]}", f"{R['P']['RV']['p0'][2]}/{R['P']['RV']['p0'][1]}", "[Assumed]"),
        ("無框力矩馬達 2026 單價", R["P"]["FM"]["p0"][0], R["P"]["FM"]["p0"][2], C["bl"][("FM", "TOT")], f"{R['P']['FM']['p0'][0]}/{R['P']['FM']['p0'][1]}", f"{R['P']['FM']['p0'][2]}/{R['P']['FM']['p0'][1]}", "[Derived]／[Assumed]"),
        ("六維力 2026 單價", R["P"]["F6"]["p0"][0], R["P"]["F6"]["p0"][2], C["bl"][("F6", "TOT")], f"{R['P']['F6']['p0'][0]}/{R['P']['F6']['p0'][1]}", f"{R['P']['F6']['p0'][2]}/{R['P']['F6']['p0'][1]}", "[Interested-party]"),
        ("人形每台諧波顆數", uaddr("HR", "HUM", 0), uaddr("HR", "HUM", 2), C["bl"][("HR", "HUM")], f"{uaddr('HR','HUM',0)}/{uaddr('HR','HUM',1)}", f"{uaddr('HR','HUM',2)}/{uaddr('HR','HUM',1)}", "[Interested-party]"),
        ("Robotaxi LiDAR 價格乘數", "2", "5", C["bl"][("LDR", "RT")], f"2/{addr(wi, R['M']['LDR'][0], R['M']['LDR'][1]+LK.index('RT'))}", f"5/{addr(wi, R['M']['LDR'][0], R['M']['LDR'][1]+LK.index('RT'))}", "[Assumed]"),
    ]
    SENS_PRS = [
        ("人形 2030 出貨量（萬台）", drv("HUM", 0), drv("HUM", 2), C["bl"][("PRS", "HUM")], f"{drv('HUM',0)}/{drv('HUM',1)}", f"{drv('HUM',2)}/{drv('HUM',1)}", "[Assumed]"),
        SENS_TOTAL[1], SENS_TOTAL[2], SENS_TOTAL[3],
    ]
    HRRV = f"({C['bl'][('HR','TOT')]}+{C['bl'][('RV','TOT')]})"
    SENS_RED = [
        ("工業機器人 2030 裝機（萬台）", drv("IND", 0), drv("IND", 2), f"({C['bl'][('HR','IND')]}+{C['bl'][('HR','COB')]}+{C['bl'][('RV','IND')]})", f"{drv('IND',0)}/{drv('IND',1)}", f"{drv('IND',2)}/{drv('IND',1)}", "[Assumed]"),
        ("人形 2030 出貨量（萬台）", drv("HUM", 0), drv("HUM", 2), C["bl"][("HR", "HUM")], f"{drv('HUM',0)}/{drv('HUM',1)}", f"{drv('HUM',2)}/{drv('HUM',1)}", "[Assumed]"),
        SENS_TOTAL[11], SENS_TOTAL[12], SENS_TOTAL[15],
        ("傳統工業機器人每台諧波顆數", uaddr("HR", "IND", 0), uaddr("HR", "IND", 2), C["bl"][("HR", "IND")], f"{uaddr('HR','IND',0)}/{uaddr('HR','IND',1)}", f"{uaddr('HR','IND',2)}/{uaddr('HR','IND',1)}", "[Analogy]"),
        ("傳統工業機器人每台 RV 顆數", uaddr("RV", "IND", 0), uaddr("RV", "IND", 2), C["bl"][("RV", "IND")], f"{uaddr('RV','IND',0)}/{uaddr('RV','IND',1)}", f"{uaddr('RV','IND',2)}/{uaddr('RV','IND',1)}", "[Analogy]"),
    ]
    r = 1
    wsn.cell(r, 1, "敏感度（承重變數排序在前，基準結果在後）：一次只動一個變數，其餘維持基準；模型為乘法結構，輸出＝基準 + 受影響部分 ×（比率 − 1），為精確值").font = BOLD; r += 2
    C["rank"] = {}

    def tornado(r, title, target, rows, key):
        n = len(rows)
        wsn.cell(r, 1, title).font = BOLD; r += 1
        # ranked table first
        hdr(wsn, r, ["排名", "變數", "輸出@低", "輸出@高", "擺幅（百萬美元）", "擺幅 ÷ 基準"]); r += 1
        rk_top = r
        raw_top = r + n + 3
        for i in range(n):
            wsn.cell(r, 1, i + 1)
            m = f"MATCH({i+1},$H${raw_top}:$H${raw_top+n-1},0)"
            wsn.cell(r, 2, f"=INDEX($B${raw_top}:$B${raw_top+n-1},{m})")
            for j, colsrc in enumerate(["E", "F", "G"]):
                wsn.cell(r, 3 + j, f"=INDEX(${colsrc}${raw_top}:${colsrc}${raw_top+n-1},{m})").number_format = "#,##0"
            wsn.cell(r, 6, f"=E{r}/{target}").number_format = "0%"
            C["rank"][(key, i + 1)] = (addr(wsn, r, 2), addr(wsn, r, 3), addr(wsn, r, 4), addr(wsn, r, 5))
            r += 1
        wsn.cell(r, 2, "基準結果（百萬美元）"); wsn.cell(r, 3, f"={target}").number_format = "#,##0"
        C["rank"][(key, "base")] = addr(wsn, r, 3)
        r += 2
        hdr(wsn, r, ["序", "變數", "低值", "高值", "輸出@低", "輸出@高", "擺幅", "名次", "受影響基準值", "標記"]); r += 1
        assert r == raw_top
        for i, (nm, lo, hi, aff, rl, rh, tag) in enumerate(rows):
            wsn.cell(r, 1, i + 1); wsn.cell(r, 2, nm)
            wsn.cell(r, 3, f"={lo}"); wsn.cell(r, 4, f"={hi}")
            for cc in (3, 4):
                wsn.cell(r, cc).number_format = "#,##0.###"
            wsn.cell(r, 9, f"={aff}").number_format = "#,##0"
            wsn.cell(r, 5, f"={target}+I{r}*({rl}-1)").number_format = "#,##0"
            wsn.cell(r, 6, f"={target}+I{r}*({rh}-1)").number_format = "#,##0"
            wsn.cell(r, 7, f"=ABS(F{r}-E{r})+ROW()/1000000").number_format = "#,##0"
            wsn.cell(r, 8, f"=RANK(G{r},$G${raw_top}:$G${raw_top+n-1},0)")
            wsn.cell(r, 10, tag)
            r += 1
        return r + 1

    r = tornado(r, "S1. 目標：2030 年五線歸因零組件不重複合計（基準情境，百萬美元）", T, SENS_TOTAL, "T")
    r = tornado(r, "S2. 目標：2030 年行星滾柱絲杠市場規模（百萬美元）", C["bl"][("PRS", "TOT")], SENS_PRS, "PRS")
    r = tornado(r, "S3. 目標：2030 年諧波＋RV 減速機市場規模（百萬美元）", HRRV, SENS_RED, "RED")
    wsn.column_dimensions["B"].width = 44
    for col in "CDEFGHI":
        wsn.column_dimensions[col].width = 14

    # ================================================================ 情境
    r = 1
    wsc.cell(r, 1, "情境（2026–2032；低／基準／高＝需求面同向：出貨量與每台用量；單價一律用基準路徑，價格不確定性見敏感度與滾柱絲杠矩陣）").font = BOLD; r += 2
    C["sc"] = {}
    for blk, (title, kind) in enumerate([("A. 市場規模（百萬美元）", "mkt"), ("B. 需求量（萬顆／根；PNE、FAS 為萬美元）", "units")]):
        wsc.cell(r, 1, title).font = BOLD; r += 1
        hdr(wsc, r, ["鍵", "零組件", "情境"] + YEARS + ["CAGR 2026–2030"]); r += 1
        for c in COMP:
            ck = c[0]
            for s in range(3):
                wsc.cell(r, 1, ck); wsc.cell(r, 2, c[1]); wsc.cell(r, 3, SC[s])
                for j, y in enumerate(YEARS):
                    wsc.cell(r, 4 + j, f"={C[kind][(s, ck, y)]}").number_format = "#,##0"
                    C["sc"][(kind, s, ck, y)] = addr(wsc, r, 4 + j)
                wsc.cell(r, 4 + len(YEARS), f"=IF(D{r}<=0,0,(H{r}/D{r})^(1/4)-1)").number_format = "0%"
                if s == 1:
                    for j in range(len(YEARS) + 4):
                        wsc.cell(r, 1 + j).fill = GRY
                r += 1
        if kind == "mkt":
            for s in range(3):
                wsc.cell(r, 2, "不重複合計（不含靈巧手模組）").font = BOLD; wsc.cell(r, 3, SC[s])
                for j, y in enumerate(YEARS):
                    wsc.cell(r, 4 + j, f"={C['nd'][(s, y)]}").number_format = "#,##0"
                wsc.cell(r, 4 + len(YEARS), f"=(H{r}/D{r})^(1/4)-1").number_format = "0%"
                r += 1
        r += 1
    # PRS 3x3 matrix 2030
    wsc.cell(r, 1, "C. 行星滾柱絲杠 2030 市場規模矩陣（百萬美元）：需求情境 × 2026 單價口徑（年降幅用基準）").font = BOLD; r += 1
    hdr(wsc, r, ["", "需求情境 →"] + SC + ["2030 需求量（萬根）→"] + SC); r += 1
    C["prsmx"] = {}
    for i, lab in enumerate(["中國量產（五洲均價）", "覓途占比推算（基準）", "MS 推算中值"]):
        pidx = [2, 1, 0][i]  # PRSPT index
        wsc.cell(r, 1, lab); wsc.cell(r, 2, f"={R['PRSPT'][pidx]}").number_format = "#,##0"
        for s in range(3):
            f = f"={C['units'][(s, 'PRS', Y)]}*{R['PRSPT'][pidx]}*(1-{R['P']['PRS']['d'][1]})^4/100"
            wsc.cell(r, 3 + s, f).number_format = "#,##0"
            C["prsmx"][(i, s)] = addr(wsc, r, 3 + s)
            wsc.cell(r, 7 + s, f"={C['units'][(s, 'PRS', Y)]}").number_format = "#,##0"
        r += 1
    wsc.cell(r, 1, "註：B 欄為 2026 單價（美元/根）；2030 單價＝B ×（1 − 基準年降幅）^4")
    r += 2
    wsc.cell(r, 1, "D. 人形每台零組件內容（美元/台，基準）— 對照線2b：Optimus Gen2 類 BOM 5–6 萬美元（MS）、業界 2025 年底 3.5 萬 → 2030–35 年 1.3–1.7 萬（BofA）").font = BOLD; r += 1
    hdr(wsc, r, ["", ""] + YEARS); r += 1
    wsc.cell(r, 2, "每台零組件內容")
    for j, y in enumerate(YEARS):
        wsc.cell(r, 3 + j, f"={C['humc'][y]}").number_format = "#,##0"
    r += 1
    wsc.column_dimensions["B"].width = 36

    # ================================================================ 供需
    r = 1
    wsd.cell(r, 1, "供需：需求（三情境）對照產能（基準）；稼動＝需求 ÷ 產能。>100% 缺口、85–100% 偏緊、<60% 過剩風險。只對有產能錨點的零組件。需求為五線歸因，未含工具機、半導體設備、割草機等其他需求，實際稼動更高").font = BOLD
    wsd.cell(r, 1).alignment = WRAP; r += 2
    C["sup"] = {}
    for ck in SUP:
        wsd.cell(r, 1, f"{ck}　{dict((c[0], c[1]) for c in COMP)[ck]}").font = BOLD; r += 1
        hdr(wsd, r, ["項目"] + YEARS); r += 1
        rows = {}
        for s in range(3):
            wsd.cell(r, 1, f"需求 {SC[s]}（萬）")
            for j, y in enumerate(YEARS):
                wsd.cell(r, 2 + j, f"={C['units'][(s, ck, y)]}").number_format = "#,##0"
            rows[f"d{s}"] = r; r += 1
        wsd.cell(r, 1, "產能 基準（萬）")
        for j, y in enumerate(YEARS):
            wsd.cell(r, 2 + j, f"={R['S'][ck]['cap'][1]}*(1+{R['S'][ck]['g'][1]})^({y}-2026)").number_format = "#,##0"
        rows["cap"] = r; r += 1
        for s in range(3):
            wsd.cell(r, 1, f"稼動 {SC[s]}")
            for j, y in enumerate(YEARS):
                col = CL(2 + j)
                wsd.cell(r, 2 + j, f"={col}{rows[f'd{s}']}/{col}{rows['cap']}").number_format = "0%"
                C["sup"][(ck, s, y)] = addr(wsd, r, 2 + j)
            rows[f"u{s}"] = r; r += 1
        wsd.cell(r, 1, "價格壓力旗標（基準）")
        for j, y in enumerate(YEARS):
            col = CL(2 + j)
            wsd.cell(r, 2 + j, f'=IF({col}{rows["u1"]}>1,"缺口",IF({col}{rows["u1"]}>0.85,"偏緊",IF({col}{rows["u1"]}<0.6,"過剩風險","平衡")))')
            C["sup"][(ck, "flag", y)] = addr(wsd, r, 2 + j)
        r += 1
        wsd.cell(r, 1, "缺口（萬，基準需求 − 產能；負＝餘裕）")
        for j, y in enumerate(YEARS):
            col = CL(2 + j)
            wsd.cell(r, 2 + j, f"={col}{rows['d1']}-{col}{rows['cap']}").number_format = "#,##0"
            C["sup"][(ck, "gap", y)] = addr(wsd, r, 2 + j)
        r += 1
        wsd.cell(r, 1, "產能出處：" + SUP[ck][3]).alignment = WRAP
        r += 2
    wsd.column_dimensions["A"].width = 34

    # ================================================================ 份額
    SH = [
        ("諧波減速機", "哈默納科 6324", "RED", "日", "是", "85%", "全球工業機器人用，2023", "[Interested-party] 中國報告網（產業報告商）", "https://www.chinabaogao.com/jingzheng/202507/758546.html"),
        ("諧波減速機", "綠的諧波 688017", "RED", "陸（A 股）", "是", "8%（2023）；12%（2025）", "全球", "[Interested-party] 中國報告網；2025 值為 Zane Robotics 未註明出處", "同上；https://zanerobotics.substack.com/p/the-technical-gap-in-chinese-robot"),
        ("諧波減速機", "哈默納科／綠的諧波／來福", "RED", "", "", "40%／18%／10%；CR3 68%", "中國，2023", "[Interested-party] 中國報告網", "同上"),
        ("RV 減速機", "Nabtesco 6268", "RED", "日", "是", ">53%（全球）；60%（中國）", "2025", "[Interested-party] Zane Robotics，未註明原始出處", "https://zanerobotics.substack.com/p/the-technical-gap-in-chinese-robot"),
        ("RV 減速機", "雙環傳動 002472、中大力德 002896", "RED", "陸", "是", "未取得", "", "—", ""),
        ("行星滾柱絲杠", "五洲新春 603667", "SCR", "陸", "是", "未取得（2025 出貨約 8.7 萬根）", "", "[Interested-party] 產業網站經 humanoid.guide", "https://humanoid.guide/?p=20588"),
        ("行星滾柱絲杠", "雙林（未在 v0.5）、新劍傳動（未上市）", "SCR", "陸", "雙林：待核", "未取得", "", "[Interested-party] China Daily、36Kr 經 humanoid.guide", "https://humanoid.guide/?p=20588"),
        ("行星滾柱絲杠", "Schaeffler（含 Ewellix）、Rollvis、GSA、SKF", "SCR", "歐", "否", "未取得（歐系為現有技術領先者）", "", "[Interested-party] humanoid.guide", "https://humanoid.guide/?p=20588"),
        ("滾珠絲杠與線性導軌", "THK 6481、上銀 2049、NSK 6471", "SCR", "日／台", "是", "未取得", "", "—", ""),
        ("無框與空心杯馬達", "Maxon、Kollmorgen（RRX）、鳴志 603728", "MOT", "", "鳴志 是", "未取得", "", "—", ""),
        ("伺服驅動", "安川 6506、匯川 300124、台達電 2308", "MOT", "", "是", "未取得", "", "—", ""),
        ("六維力感測（中國人形）", "藍點觸控（未在 v0.5）", "FRC", "陸", "待核", ">80%", "中國，2025", "[Interested-party] 賽迪顧問經同花順", "https://stock.10jqka.com.cn/20260831/c679460636.shtml"),
        ("六維力感測（中國人形）", "坤維科技／宇立儀器", "FRC", "陸", "未上市", "9.2%／3.3%；CR3 92.8%", "中國，2025", "同上", "同上"),
        ("氣動元件", "SMC 6273", "PNE", "日", "是", "未取得（FY3/26 營收 8,425 億日圓，+6.4%）", "", "[Verified] 經 envalith 整理財報", "https://envalith.com/en/companies/6273/overview"),
        ("工業感測與機器視覺", "Keyence 6861", "FAS", "日", "是", "未取得", "", "—", ""),
        ("靈巧手", "因時機器人（未上市）、兆威機電 003021", "HAND", "陸", "兆威 是", "未取得（因時出貨領先，媒體）", "", "[Interested-party] 媒體", ""),
        ("LiDAR（全部用途）", "禾賽 HSAI／速騰聚創 2498", "LDR", "美港／港", "是", "兩家 2025 出貨 162.0 萬／91.2 萬顆，兩家間比 64%／36%", "2025", "[Verified] 財報經 laserfair", "https://www.laserfair.com/English/202604/07/1040.html"),
        ("相機（CIS）", "Sony 6758、豪威 603501；模組舜宇 2382.HK", "VIS", "", "是", "未取得", "", "—", ""),
    ]
    r = 1
    wsh.cell(r, 1, "主要廠份額（只列可查證或註明利害方的數字；查不到寫未取得）").font = BOLD; r += 2
    hdr(wsh, r, ["零組件", "公司（代號）", "v0.5 次產業", "市場", "亞太掛牌", "份額", "口徑／年份", "標記", "出處"]); r += 1
    C["share_rows"] = {}
    for row in SH:
        for j, v in enumerate(row):
            wsh.cell(r, 1 + j, v).alignment = WRAP
        C["share_rows"][row[1]] = r
        r += 1
    # LiDAR derived share formula
    r += 1
    wsh.cell(r, 1, "LiDAR 兩家間比（公式）"); wsh.cell(r, 2, "禾賽 2025 出貨（萬顆）"); c = wsh.cell(r, 3, 162.04); inp(c)
    wsh.cell(r, 4, "速騰 2025 出貨（萬顆）"); c = wsh.cell(r, 5, 91.2); inp(c)
    wsh.cell(r, 6, f"=C{r}/(C{r}+E{r})").number_format = "0%"; C["ldsh"] = addr(wsh, r, 6)
    for col, w in zip("ABCDEFGHI", [18, 34, 10, 10, 12, 30, 16, 34, 50]):
        wsh.column_dimensions[col].width = w

    # ================================================================ 下游取數
    r = 1
    wdl.cell(r, 1, "下游取數（給個股模型：哈默納科、Nabtesco、上銀、Fanuc、安川、Tesla 等）。口徑：「五線歸因」＝只計工業機器人含協作、人形、AMR、智慧車、Robotaxi 拉動的需求，不含工具機、半導體設備等").font = BOLD
    wdl.cell(r, 1).alignment = WRAP; r += 2
    hdr(wdl, r, ["名稱", "低", "基準", "高", "單位", "口徑", "給誰"]); r += 1
    DL = []
    def dl(name, lo, ba, hi, unit, scope, who, fmt="#,##0"):
        nonlocal r
        wdl.cell(r, 1, name)
        for j, f in enumerate((lo, ba, hi)):
            wdl.cell(r, 2 + j, f"={f}" if f else "").number_format = fmt
        wdl.cell(r, 5, unit); wdl.cell(r, 6, scope).alignment = WRAP; wdl.cell(r, 7, who)
        DL.append((name, r)); r += 1
    for y in (2026, 2030, 2032):
        dl(f"諧波減速機市場規模 {y}", C["mkt"][(0, "HR", y)], C["mkt"][(1, "HR", y)], C["mkt"][(2, "HR", y)], "百萬美元", "五線歸因；需求情境（單價基準）", "哈默納科、綠的諧波、上銀")
    for y in (2026, 2030, 2032):
        dl(f"RV 減速機市場規模 {y}", C["mkt"][(0, "RV", y)], C["mkt"][(1, "RV", y)], C["mkt"][(2, "RV", y)], "百萬美元", "五線歸因（只來自傳統工業機器人）", "Nabtesco、雙環傳動")
    dl("諧波減速機需求量 2030", C["units"][(0, "HR", Y)], C["units"][(1, "HR", Y)], C["units"][(2, "HR", Y)], "萬顆", "五線歸因", "哈默納科")
    dl("諧波需求中人形占比 2030（基準）", None, C["bl"][("HR", "HUMSH")], None, "%", "基準情境", "哈默納科、綠的諧波", "0%")
    for y in (2026, 2030, 2032):
        dl(f"行星滾柱絲杠市場規模 {y}（基準單價口徑＝覓途）", C["mkt"][(0, "PRS", y)], C["mkt"][(1, "PRS", y)], C["mkt"][(2, "PRS", y)], "百萬美元", "需求情境；單價口徑敏感度見下兩列與情境頁 C", "五洲新春、上銀、Schaeffler、Tesla")
    dl("行星滾柱絲杠 2030：基準需求 × 單價口徑（中國量產／覓途／MS）", C["prsmx"][(0, 1)], C["prsmx"][(1, 1)], C["prsmx"][(2, 1)], "百萬美元", "低＝中國量產價、基準＝覓途、高＝MS（此列三欄是價格口徑，不是需求情境）", "同上")
    dl("行星滾柱絲杠 2030 全矩陣最小／最大", C["prsmx"][(0, 0)], None, C["prsmx"][(2, 2)], "百萬美元", "低需求 × 中國價；高需求 × MS 價", "同上")
    dl("行星滾柱絲杠需求量 2030", C["units"][(0, "PRS", Y)], C["units"][(1, "PRS", Y)], C["units"][(2, "PRS", Y)], "萬根", "全來自人形", "同上")
    dl("行星滾柱絲杠 2026 稼動（需求 ÷ 基準產能）", C["sup"][("PRS", 0, 2026)], C["sup"][("PRS", 1, 2026)], C["sup"][("PRS", 2, 2026)], "%", "產能為 [Assumed]", "同上", "0%")
    dl("行星滾柱絲杠 2030 稼動", C["sup"][("PRS", 0, Y)], C["sup"][("PRS", 1, Y)], C["sup"][("PRS", 2, Y)], "%", "產能為 [Assumed]", "同上", "0%")
    dl("LiDAR 市場規模 2030", C["mkt"][(0, "LDR", Y)], C["mkt"][(1, "LDR", Y)], C["mkt"][(2, "LDR", Y)], "百萬美元", "五線歸因（不含割草機等）", "禾賽、速騰")
    dl("LiDAR 需求量 2030", C["units"][(0, "LDR", Y)], C["units"][(1, "LDR", Y)], C["units"][(2, "LDR", Y)], "萬顆", "同上", "禾賽、速騰")
    dl("六維力感測市場規模 2030", C["mkt"][(0, "F6", Y)], C["mkt"][(1, "F6", Y)], C["mkt"][(2, "F6", Y)], "百萬美元", "五線歸因", "Novanta、柯力")
    dl("滾珠絲杠需求量 2030（主要來自 EMB）", C["units"][(0, "BS", Y)], C["units"][(1, "BS", Y)], C["units"][(2, "BS", Y)], "萬根", "只計增量（EMB、SBW、機器人）；不含既有 EPS 與工具機", "上銀、THK、NSK")
    dl("零組件不重複合計 2030", C["nd"][(0, Y)], C["nd"][(1, Y)], C["nd"][(2, Y)], "百萬美元", "五線歸因，不含靈巧手模組、運算、電池", "總覽")
    dl("人形每台零組件內容 2026（基準）", None, C["humc"][2026], None, "美元/台", "不含運算、電池、結構、組裝；對照線2b BOM", "Tesla、人形整機")
    dl("人形每台零組件內容 2030（基準）", None, C["humc"][2030], None, "美元/台", "同上", "Tesla、人形整機")
    dl("人形占 2030 不重複合計比重（基準）", None, f"{C['bl'][('ND','HUM')]}/{C['bl'][('ND','TOT')]}", None, "%", "基準", "總覽", "0%")
    r += 1
    wdl.cell(r, 1, "主要廠份額（取自份額頁；未取得者不列）").font = BOLD; r += 1
    hdr(wdl, r, ["零組件", "公司", "份額", "口徑", "標記"]); r += 1
    for row in SH:
        if "未取得" in row[5] and not row[5].startswith("兩家"):
            continue
        sr = C["share_rows"][row[1]]
        wdl.cell(r, 1, f"='份額'!A{sr}"); wdl.cell(r, 2, f"='份額'!B{sr}"); wdl.cell(r, 3, f"='份額'!F{sr}")
        wdl.cell(r, 4, f"='份額'!G{sr}"); wdl.cell(r, 5, f"='份額'!H{sr}")
        r += 1
    wdl.column_dimensions["A"].width = 52; wdl.column_dimensions["F"].width = 44; wdl.column_dimensions["G"].width = 30
    for col in "BCD":
        wdl.column_dimensions[col].width = 12

    # ================================================================ 對帳
    REC = [
        ("工業機器人 2025 全球裝機", "2a", "60.3 萬台", "線2a 框架 §3（IFR [Verified]）", "校準基準"),
        ("工業機器人 2026 裝機（基準錨）", "2a", "65.5 萬台", "線2a 框架 §1（IFR 預測）", "輸入 B：IND"),
        ("工業機器人 2029 裝機", "2a", "80.6 萬台 → 本模型 2030 基準 83 萬", "線2a 框架 §1（IFR 預測）", "輸入 B：IND 2030 為外推 [Assumed]"),
        ("協作占比", "2a", "11%（2025）", "線2a 框架 §3（IFR）", "輸入 B：COBSH"),
        ("傳統 6 軸：RV 底座大臂、諧波腕部；協作全諧波", "2a", "RV 3.5、諧波 2；協作諧波 6", "線2a 框架 §3 物理註記（顆數為本模型 [Analogy]）", "輸入 C 用量"),
        ("人形 2025 出貨", "2b", "約 1.8 萬台（IDC）", "線2b 框架 Sources", "校準基準"),
        ("人形 2026 出貨（基準錨）", "2b", "5 萬台（3–7 萬）", "線2b 框架 §7-9：TrendForce >5 萬；MS 中國 2.8 萬；H1 追蹤 1.9–2.5 萬", "輸入 B：HUM；**口徑衝突待主代理對帳**"),
        ("人形 2030／2032 出貨", "2b", "50 萬（15–120 萬）／150 萬（30–350 萬）", "本模型 [Assumed]；線2b 框架未給 2030 量", "**待線2b 模型完成後替換**"),
        ("人形每台諧波", "2b", "出貨組合平均 10（6–16）；Optimus 類 14–16", "線2b 框架 §3.2（14–16 為全尺寸）", "輸入 C；**組合平均待線2b 對帳**"),
        ("人形每台滾柱絲杠", "2b", "出貨組合平均 8（4–14）；Optimus 類 12–14", "線2b 框架 §3.2（12–14 為全尺寸）", "輸入 C；**組合平均待線2b 對帳**"),
        ("人形每台無框馬達／驅動", "2b", "28（28–30）", "線2b 框架 §3.2", "輸入 C"),
        ("人形每台編碼器", "2b", "56（28–60）", "線2b 框架 §3.2（每關節 2）", "輸入 C"),
        ("人形每台六維力", "2b", "2（2–4）", "線2b 框架 §3.2", "輸入 C"),
        ("人形每台靈巧手", "2b", "2", "線2b 框架 §3.2", "輸入 C"),
        ("人形 BOM（對照）", "2b", "5–6 萬美元（MS）；3.5 萬 → 1.3–1.7 萬（BofA）", "線2b 框架 §3.2", "計算 7：每台零組件內容應低於 BOM"),
        ("滾柱絲杠單價矛盾", "2b", "MS 1,350–2,700 vs 覓途占 19%", "線2b 框架 §3.2 交叉檢核 2", "輸入 D；情境 C；敏感度 S2"),
        ("運輸物流機器人 2025", "3", "11.75 萬台，+21%", "線3 框架 §3（IFR 樣本）", "輸入 B：AMR 2026 = ×1.21"),
        ("AMR 單價", "3", "未取得", "線3 框架 §3", "本模型不需要整機單價"),
        ("中國城市 NOA 滲透率", "1b", "18.8%（2026H1）；2030 62%", "線1b 框架 §3", "LiDAR 搭載率上界參考"),
        ("中國 EMB 滲透率", "1b", "2026 約 5%、2030 約 40%", "線1b 框架 §3（產業媒體）", "輸入 B：EPEN（全球化換算）"),
        ("全球輕型車銷量", "1b", "9,000 萬輛（8,800–9,200 萬）", "**五份框架皆未給**；[Assumed]", "**待線1b 模型提供**"),
        ("SBW 滲透率", "1b", "未取得", "線1b 框架 §7", "輸入 B：SPEN [Assumed]"),
        ("Robotaxi 車隊", "1a", "Waymo 約 4,000；小馬 1,400+（目標 3,000+）；文遠 1,800+；Tesla 稱 300+", "線1a 框架 §2", "輸入 B：RT 2026 新增 1 萬（0.6–1.5 萬）"),
        ("Robotaxi 2030 新增車輛", "1a", "8 萬（2–25 萬）", "本模型 [Assumed]", "**待線1a 模型提供**"),
        ("Robotaxi 每車 LiDAR／相機", "1a", "4／11（Waymo 4／13；Tesla 0／約 8）", "線1a 框架 §7 指定 LiDAR 單價歸共用層", "輸入 C"),
    ]
    r = 1
    wr.cell(r, 1, "對帳：本模型引用的各線出貨量與每台用量（其他五線模型完成後由主代理逐列對帳；粗體＝待替換）").font = BOLD; r += 2
    hdr(wr, r, ["變數", "所屬線", "使用值", "出處", "本模型位置"]); r += 1
    for row in REC:
        for j, v in enumerate(row):
            wr.cell(r, 1 + j, v.replace("**", "")).alignment = WRAP
            if "**" in v:
                wr.cell(r, 1 + j).font = BOLD
        r += 1
    for col, w in zip("ABCDE", [34, 8, 40, 50, 36]):
        wr.column_dimensions[col].width = w

    # ================================================================ 追蹤指標
    TR = [
        ("哈默納科營收", "季", "公司決算短信", "FY03/27 Q1 營收 166.81 億日圓，+23.6%；全年指引 745 億日圓", "2026-08", "[Verified] 經 Quartr 摘要，待核原文", "諧波最上游讀數；人形訂單是否進入"),
        ("哈默納科受注（線2a 已追）", "季", "公司決算短信", "2026 年 4–6 月 241 億日圓，+55.7%", "2026-08", "[Verified]（線2a）", "訂單領先營收"),
        ("Nabtesco 集團訂單", "季", "公司決算", "2026 年 1–6 月 1,985 億日圓，+33.1%；精密減速機分部未取得", "2026-08", "[Verified]（線2a）", "RV 需求＝傳統工業機器人"),
        ("JARA 受注額", "季", "日本機器人工業會", "2026 年 7–9 月 3,254 億日圓，+46.6%", "2026-10", "[Verified]（線2a）", "工業機器人循環"),
        ("禾賽 LiDAR 出貨與 ASP", "季", "禾賽財報", "2026Q2 出貨 62.8 萬顆（ADAS 48.6 萬、機器人 14.2 萬）；ASP 1,297 人民幣，年減 35%", "2026-08", "[Verified] 經長橋海豚研究整理", "LiDAR 價格曲線與車用搭載"),
        ("禾賽 2026 指引與產能", "年", "禾賽法說", "出貨 300–350 萬顆；產能 >400 萬顆", "2026-03／08", "[Interested-party] 公司指引", "供給面"),
        ("速騰聚創出貨", "季", "公司財報", "2025 年 91.2 萬顆（ADAS 60.9 萬、機器人等 30.3 萬）", "2026-03", "[Verified] 經 laserfair", "份額"),
        ("五洲新春滾柱絲杠出貨與均價", "半年", "公司財報、法說", "2025 年約 8.7 萬根；均價約 1,360 人民幣，年減 39%", "2026", "[Interested-party] 產業網站", "解決絲杠單價矛盾"),
        ("新劍傳動、雙林滾柱絲杠產線", "事件", "公司公告、媒體", "新劍據報建 100 萬根/年；雙林 499 人民幣量產（2026-01）", "2026", "[Interested-party]", "供給擴張與價格下限"),
        ("Schaeffler 人形訂單", "半年", "Schaeffler 法說", "2030 年前人形訂單達數億歐元；目標拿下可服務 BOM（約 50%）的約 10%", "2026-05-05", "[Interested-party] 公司（Reuters 經 Yahoo）", "歐系絲杠與減速機廠進入人形"),
        ("中國人形六維力出貨與份額", "年", "賽迪顧問", "2025 藍點觸控 >80%；2030 中國 120 萬顆、60 億元", "2026-08-31", "[Interested-party] 賽迪", "感測價格曲線"),
        ("SMC 營收", "季", "公司決算", "FY3/26 營收 8,425 億日圓，+6.4%；營益率 22.6%", "2026-05", "[Verified] 經 envalith", "工廠自動化循環（氣動）"),
        ("Keyence 營收", "季", "公司決算", "2026 年 4–6 月營收 3,466 億日圓，+32.8%", "2026-07-29", "[Verified] 經 Quartr 摘要，待核原文", "工廠自動化循環（感測與視覺）"),
        ("中國工業機器人產量", "月", "國家統計局", "2026 年 9 月 96,752 台，+18.9%", "2026-10", "[Verified]（線2a）", "最高頻需求讀數"),
        ("Tesla Optimus 量產節奏", "季", "Tesla 法說", "Fremont 設計產能 100 萬台/年；達每週 2 萬台至少需一年", "2026-10", "[Interested-party]（線2b）", "人形量的最大單一變數"),
    ]
    r = 1
    wt.cell(r, 1, "追蹤指標（框架第 5 節；含最新值與日期）").font = BOLD; r += 2
    hdr(wt, r, ["指標", "頻率", "資料來源", "最新值", "日期", "標記", "為何重要"]); r += 1
    for row in TR:
        for j, v in enumerate(row):
            wt.cell(r, 1 + j, v).alignment = WRAP
        r += 1
    for col, w in zip("ABCDEFG", [28, 6, 18, 50, 12, 26, 30]):
        wt.column_dimensions[col].width = w

    # ================================================================ 說明
    txt = [
        ("線 6 共用零組件單位經濟 v0.1", True),
        ("產生：2026-11-16｜框架：線6_共用零組件_框架_v0.1.md｜公開 repo，不含個人情境", False),
        ("", False),
        ("核心等式", True),
        ("需求量(c, y) ＝ Σ_line 出貨量(line, y) × 每台用量(c, line)", False),
        ("市場規模(c, y) ＝ Σ_line 出貨量 × 用量 × 價格乘數(c, line) × 單價(c, y)；單價(c, y) ＝ 2026 單價 ×（1 − 年降幅）^(y − 2026)", False),
        ("稼動(c, y) ＝ 需求量 ÷ 產能（諧波、RV、滾柱絲杠、LiDAR）", False),
        ("線：工業機器人（傳統、協作）、人形、運輸物流機器人、輕型車（全部、搭 LiDAR、搭 EMB、搭 SBW）、Robotaxi 新增車輛", False),
        ("", False),
        ("口徑", True),
        ("1. 「五線歸因」市場：只計上列各線拉動的需求。諧波另有半導體設備需求、滾珠絲杠與線性導軌主需求在工具機、LiDAR 另有割草機等機器人需求，皆不在內 → 本模型數字低於公司營收口徑的整體市場。", False),
        ("2. 情境＝需求面同向（出貨量 × 用量），單價一律用基準路徑；單價不確定性放敏感度，滾柱絲杠另做 3×3 矩陣（情境頁 C）。", False),
        ("3. 不重複合計不含靈巧手模組（與空心杯、觸覺重複），也不含運算（AI供應鏈權威）、電池、結構件。", False),
        ("4. 氣動（SMC）與工業感測視覺（Keyence）以每台美元內容建模，數字小，代表這兩家的需求主體不在五線。", False),
        ("5. 單位：出貨量萬台；需求量萬顆；市場規模百萬美元；美元換算匯率見輸入 A。", False),
        ("", False),
        ("頁面", True),
        ("輸入｜計算｜敏感度（排序在前）｜情境｜供需｜份額｜下游取數｜對帳（引用各線出貨量，由主代理對帳）｜追蹤指標", False),
        ("", False),
        ("格式", True),
        ("藍字＝輸入；黃底＝判斷（Assumed／Analogy），可改；黑字＝公式。標記：[Verified]／[Interested-party]／[Analogy]／[Assumed]／[Derived]；Analogy、Assumed 一律給低／基準／高三值。", False),
        ("", False),
        ("已知限制", True),
        ("各線出貨量為暫用值（五線模型同時建，完成版讀不到）；人形 2030 量、全球輕型車銷量、Robotaxi 2030 新增車輛皆為本模型 [Assumed]，對帳頁粗體列待替換。", False),
        ("諧波、RV、滾柱絲杠產能為 [Assumed]（公司未揭露產能）；供需稼動只能看方向。", False),
        ("RV、滾珠絲杠、線性導軌、伺服、編碼器、觸覺、相機單價皆未取得，為 [Assumed]。", False),
    ]
    for i, (t, b) in enumerate(txt):
        c = ws0.cell(i + 1, 1, t); c.alignment = WRAP
        if b:
            c.font = BOLD
    ws0.column_dimensions["A"].width = 140
    ws0.cell(1, 3, "檢核：計算 6 依線拆分合計 − 5.2 不重複合計")
    ws0.cell(2, 3, f"={C['chk1']}")

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)
    return OUT


def recalc(path):
    tmp = "/tmp/lo_components_rc"
    shutil.rmtree(tmp, ignore_errors=True); os.makedirs(tmp)
    prof = "file:///tmp/lo_components_profile"
    subprocess.run(["soffice", f"-env:UserInstallation={prof}", "--headless", "--norestore", "--calc",
                    "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", tmp, path],
                   check=True, capture_output=True, timeout=180)
    out = os.path.join(tmp, os.path.basename(path))
    wb = openpyxl.load_workbook(out, data_only=True)
    errs = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("#") or (isinstance(c.value, str) and c.value in ("#DIV/0!", "#REF!", "#NAME?", "#VALUE!", "Err:502", "Err:504")):
                    errs.append(f"{ws.title}!{c.coordinate}={c.value}")
                if isinstance(c.value, str) and c.value.startswith("Err:"):
                    errs.append(f"{ws.title}!{c.coordinate}={c.value}")
    chk = wb["說明"]["C2"].value
    if not errs and abs(chk or 0) < 1e-6:
        shutil.copy(out, path)  # 交付含快取值的版本，供 openpyxl data_only 讀取
    print("recalc file:", out)
    print("errors:", len(errs), errs[:20])
    print("check (should be 0):", chk)
    return out, errs


if __name__ == "__main__":
    p = build()
    print("built", p)
    if "--recalc" in sys.argv:
        recalc(p)
