"""實體 AI 地圖 — 資料活頁簿（事實與判斷的權威位置）。

讀：data/*.csv（產業地位抄錄、收盤價）、physical_ai/lines/*/dist（六條線成品）、physical_ai/<公司>/dist（個股模型成品）、
    markets/invest_map/dist/data.json（跨圖個股：主圖 AI 算力地圖的模型目標價）。
寫：dist/<日期>_實體AI地圖_資料_v<版>.xlsx。黃底＝判斷輸入（評等門檻、權重）；其餘為抄錄值或公式，每格附出處（檔案｜工作表!儲存格）。
用法：python3 physical_ai/map/build_xlsx.py
"""
import csv, glob, json, os, re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VERSION = "v0.4"
DATE = "20261123"
OUT = os.path.join(HERE, "dist", f"{DATE}_實體AI地圖_資料_{VERSION}.xlsx")
YEL = PatternFill("solid", fgColor="FFF2CC")
HDR = PatternFill("solid", fgColor="DDEBF7")
BOLD = Font(bold=True)
BLUE = Font(color="0000FF")


def rcsv(name):
    with open(os.path.join(HERE, "data", name), encoding="utf-8") as f:
        return list(csv.DictReader(f))


def latest(folder, pat):
    fs = sorted(glob.glob(os.path.join(ROOT, folder, "dist", pat)))
    if not fs:
        raise SystemExit(f"找不到 {folder}/dist/{pat}")
    return fs[-1]


def rel(p):
    return os.path.relpath(p, ROOT)


# ------------------------------------------------------------------ 六條線重點數字
# (線代碼, 線名, 資料夾, 工作表, 查列方式, 鍵, 年, 欄(低,基準,高), 顯示標籤, 單位, 格式)
LINES = [
    ("1", "Robotaxi", "lines/robotaxi", "下游取數", "id", "RT_PEN_2030", None, "CDE", "美國 Robotaxi 滲透率 2030（趟數）", "%", "pct"),
    ("1", "Robotaxi", "lines/robotaxi", "下游取數", "id", "RT_PEN_2035", None, "CDE", "美國 Robotaxi 滲透率 2035（趟數）", "%", "pct"),
    ("1", "Robotaxi", "lines/robotaxi", "下游取數", "id", "RT_OPEX_WAYMO_2030", None, "CDE", "Waymo 每付費英里單位營運成本 2030", "US$/英里", "usd2"),
    ("1", "Robotaxi", "lines/robotaxi", "下游取數", "id", "RT_OPEX_TSLA_FLEET_2030", None, "CDE", "Tesla 車隊每付費英里單位營運成本 2030", "US$/英里", "usd2"),
    ("1", "Robotaxi", "lines/robotaxi", "下游取數", "id", "RT_OPPRICE_WAYMO_2030", None, "CDE", "Waymo 每付費英里價格 2030", "US$/英里", "usd2"),
    ("1", "Robotaxi", "lines/robotaxi", "下游取數", "id", "RT_PRICE_OWN_AVG", None, "CDE", "門檻：自有車每英里全成本（AAA）", "US$/英里", "usd2"),
    ("2", "智慧車", "lines/smart_car", "下游取數", "name", "FSD 活躍數（年末，含早期買斷）", 2030, "CDE", "Tesla FSD 活躍數 2030", "百萬", "f1"),
    ("2", "智慧車", "lines/smart_car", "下游取數", "name", "FSD 年末年化營收", 2030, "CDE", "FSD 年末年化營收 2030", "十億 USD", "f1"),
    ("2", "智慧車", "lines/smart_car", "下游取數", "name", "中國城市 NOA 滲透率", 2030, "CDE", "中國城市 NOA 滲透率 2030", "%", "pct"),
    ("2", "智慧車", "lines/smart_car", "下游取數", "name", "每車智能內容（中國新車平均）", 2030, "CDE", "中國新車每車智能內容 2030", "USD/車", "f0"),
    ("2", "智慧車", "lines/smart_car", "下游取數", "name", "中國智能內容市場", 2030, "CDE", "中國智能內容市場 2030", "十億 USD", "f1"),
    ("3", "工業機器人（含協作）", "lines/industrial_robot", "下游取數", "idv", "IR.inst.global.2026.base", None, "E", "全球工業機器人裝機 2026（基準）", "台", "f0"),
    ("3", "工業機器人（含協作）", "lines/industrial_robot", "下游取數", "idv", "IR.inst.global.2030.base", None, "E", "全球工業機器人裝機 2030（基準）", "台", "f0"),
    ("3", "工業機器人（含協作）", "lines/industrial_robot", "下游取數", "idv", "IR.share.cobot.2030.base", None, "E", "協作占全球裝機比 2030（基準）", "%", "pct"),
    ("3", "工業機器人（含協作）", "lines/industrial_robot", "下游取數", "idv", "IR.red.rv.units.2030.base", None, "E", "機器人 RV 減速機需求 2030（基準）", "顆", "f0"),
    ("4", "人形機器人（硬體）", "lines/humanoid", "下游取數", "id", "HUM_BOM_2026", None, "CDE", "全尺寸人形每台 BOM 2026", "USD/台", "f0"),
    ("4", "人形機器人（硬體）", "lines/humanoid", "下游取數", "id", "HUM_BOM_2030", None, "CDE", "全尺寸人形每台 BOM 2030（低／基準／高＝成本情境）", "USD/台", "f0"),
    ("4", "人形機器人（硬體）", "lines/humanoid", "下游取數", "id", "HUM_PAYBACK_2026", None, "CDE", "對美國工人回本期 2026（假設生產力已到位）", "年", "f1"),
    ("4", "人形機器人（硬體）", "lines/humanoid", "下游取數", "id", "SHIP_2030", None, "CDE", "全球人形出貨 2030", "台/年", "f0"),
    ("5", "倉儲物流自動化", "lines/warehouse", "下游取數", "id", "WH_AMR_MKT_2030", None, "CDE", "行動機器人營收 2030", "十億 USD", "f1"),
    ("5", "倉儲物流自動化", "lines/warehouse", "下游取數", "id", "WH_ASRS_MKT_2030", None, "CDE", "AS/RS 市場 2030", "十億 USD", "f1"),
    ("5", "倉儲物流自動化", "lines/warehouse", "下游取數", "id", "WH_AMR_PAYBACK", None, "CDE", "每台 AMR 回本期（穩態，悲觀／基準／樂觀）", "年", "f1"),
    ("5", "倉儲物流自動化", "lines/warehouse", "下游取數", "id", "WH_COST_MANUAL", None, "CDE", "人工揀貨每訂單行成本", "USD/行", "usd2"),
    ("5", "倉儲物流自動化", "lines/warehouse", "下游取數", "id", "WH_COST_AMR", None, "CDE", "AMR 貨到人每訂單行成本", "USD/行", "usd2"),
    ("6", "共用零組件", "lines/components", "下游取數", "name", "諧波減速機市場規模 2030", None, "BCD", "諧波減速機市場 2030", "百萬 USD", "f0"),
    ("6", "共用零組件", "lines/components", "下游取數", "name", "RV 減速機市場規模 2030", None, "BCD", "RV 減速機市場 2030", "百萬 USD", "f0"),
    ("6", "共用零組件", "lines/components", "下游取數", "name", "行星滾柱絲杠市場規模 2030（基準單價口徑＝覓途）", None, "BCD", "行星滾柱絲杠市場 2030", "百萬 USD", "f0"),
    ("6", "共用零組件", "lines/components", "下游取數", "name", "LiDAR 市場規模 2030", None, "BCD", "LiDAR 市場 2030", "百萬 USD", "f0"),
    ("6", "共用零組件", "lines/components", "下游取數", "name", "零組件不重複合計 2030", None, "BCD", "五線零組件不重複合計 2030", "百萬 USD", "f0"),
]
LINE_META = {"1": ("Robotaxi", "每英里成本何時低於價格與自有車成本？"), "2": ("智慧車（四輪機器人）", "自駕滲透率 × 每車智能內容"),
             "3": ("工業機器人（含協作）", "AI 能否把機器人從固定工位推向彈性產線？"), "4": ("人形機器人（硬體）", "每台 BOM、售價與回本期何時跨過門檻？"),
             "5": ("倉儲物流自動化（AMR、AS/RS）", "每訂單行成本與回本期"), "6": ("共用零組件", "減速機、絲杠、馬達、感測：誰供應哪幾條線？")}
# v0.2（2026-11-17）：線號改 1–6（前五為應用端、6 為共用零組件；產業地位活頁簿的「所屬線」原文 1a／1b／2a／2b／3 不改，地圖以 data/map_groups.csv 分組）


def line_values():
    out, cache = [], {}
    for code, lname, folder, sheet, how, key, yr, cols, lab, unit, fmt in LINES:
        f = latest(os.path.join("physical_ai", folder), "*.xlsx")
        wb = cache.setdefault(f, openpyxl.load_workbook(f, data_only=True))
        ws = wb[sheet]
        hit = None
        for r in range(1, ws.max_row + 1):
            a = ws.cell(r, 1).value
            if a is None:
                continue
            if how in ("id", "idv", "name") and str(a).strip() == key and (yr is None or str(ws.cell(r, 2).value) == str(yr)):
                hit = r; break
        if hit is None:
            raise SystemExit(f"{rel(f)} {sheet} 找不到 {key} {yr or ''}")
        vals = [ws[f"{c}{hit}"].value for c in cols]
        if len(vals) == 1:
            vals = [None, vals[0], None]
        if not isinstance(vals[1], (int, float)):
            raise SystemExit(f"{rel(f)} {sheet}!{cols[min(1, len(cols)-1)]}{hit} 不是數值：{vals}")
        src = f"{rel(f)}｜{sheet}!{cols[0]}{hit}" + (f":{cols[-1]}{hit}" if len(cols) > 1 else "")
        out.append(dict(code=code, lab=lab, unit=unit, fmt=fmt, lo=vals[0], base=vals[1], hi=vals[2], src=src))
    return out


# ------------------------------------------------------------------ 個股模型（主圖＝本專案）
def find_row(ws, pats):
    for r in range(1, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if a and any(re.match(p, str(a)) for p in pats):
            for c in (2, 3):
                v = ws.cell(r, c).value
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    return v, f"Val!{openpyxl.utils.get_column_letter(c)}{r}"
    return None, None


def named(wb, nm):
    sh, ref = list(wb.defined_names[nm].destinations)[0]
    return wb[sh][ref.replace("$", "")].value, f"{sh}!{ref.replace('$', '')}"


# 代號（產業地位 csv）→ (資料夾, 檔名樣式)
OWN = {"TSLA": ("tesla", "*Tesla個股模型_v*.xlsx"), "6324": ("harmonic_drive", "*.xlsx"), "6268": ("nabtesco", "*.xlsx"),
       "2049": ("hiwin", "*.xlsx"), "6954": ("fanuc", "*.xlsx"), "6506": ("yaskawa", "*.xlsx"),
       # 第二批（2026-11-23）
       "MBLY": ("mobileye", "*.xlsx"), "9660": ("horizon", "*.xlsx"), "PONY": ("pony", "*.xlsx"), "9880": ("ubtech", "*.xlsx"), "UNITREE": ("unitree", "*.xlsx"), "2590": ("geekplus", "*.xlsx"), "SYM": ("symbotic", "*.xlsx"), "1316": ("nexteer", "*.xlsx"), "6481": ("thk", "*.xlsx"), "688017": ("leaderdrive", "*.xlsx"), "6273": ("smc", "*.xlsx"), "6861": ("keyence", "*.xlsx")}
SEGM = {"WAYMO": ("waymo", "*.xlsx"), "TSLA-RT": ("tesla_robotaxi", "*.xlsx")}
# 跨圖：產業地位代號 → 算力地圖代號
CROSS = {"NVDA-C": "NVDA", "2330": "2330", "AMZN": "AMZN", "TER": "TER", "2308": "2308", "2317": "2317"}
TP_PATS = [r"^目標價（DCF 與", r"^加權目標價"]
GATE_PATS = [r"^入不敷出判定用"]
CCY = {"TSLA": "USD", "6324": "JPY", "6268": "JPY", "2049": "TWD", "6954": "JPY", "6506": "JPY",
       "MBLY": "USD", "9660": "HKD", "PONY": "USD", "9880": "HKD", "UNITREE": "CNY", "2590": "HKD", "SYM": "USD", "1316": "HKD", "6481": "JPY", "688017": "CNY", "6273": "JPY", "6861": "JPY"}


def own_models():
    out = {}
    for tk, (folder, pat) in OWN.items():
        f = latest(os.path.join("physical_ai", folder), pat)
        wb = openpyxl.load_workbook(f, data_only=True)
        V = wb["Val"]
        tp, tps = find_row(V, TP_PATS)
        gate, gs = find_row(V, GATE_PATS)
        mpx, mps = named(wb, "px")
        mpd, _ = named(wb, "px_date")
        if tp is None or gate is None:
            raise SystemExit(f"{rel(f)} 抽不到目標價或收支關卡")
        out[tk] = dict(file=rel(f), tp=tp, tp_src=f"{rel(f)}｜{tps}", gate=gate, gate_src=f"{rel(f)}｜{gs}",
                       mpx=mpx, mpx_src=f"{rel(f)}｜{mps}", mpd=str(mpd)[:10], ccy=CCY[tk])
    return out


def seg_models():
    out = {}
    f = latest("physical_ai/waymo", "*.xlsx"); wb = openpyxl.load_workbook(f, data_only=True)
    d = {}
    for k in ("eq_C", "eq_w", "eq_A"):
        v, c = named(wb, k); d[k] = (v, f"{rel(f)}｜{c}")
    out["WAYMO"] = dict(file=rel(f), lo=d["eq_C"], base=d["eq_w"], hi=d["eq_A"], note="股權價值（DCF，十億美元）；2026-02 融資投後 1,260 億美元只作對照")
    f = latest("physical_ai/tesla_robotaxi", "*.xlsx"); wb = openpyxl.load_workbook(f, data_only=True)
    d = {}
    for k in ("RT_VAL_LO", "RT_VAL_BASE", "RT_VAL_HI"):
        v, c = named(wb, k); d[k] = (v, f"{rel(f)}｜{c}")
    out["TSLA-RT"] = dict(file=rel(f), lo=d["RT_VAL_LO"], base=d["RT_VAL_BASE"], hi=d["RT_VAL_HI"], note="分部價值（DCF，十億美元）；Tesla 個股模型 SOTP 讀此值")
    return out


def cross_models():
    d = json.load(open(os.path.join(ROOT, "markets/invest_map/dist/data.json"), encoding="utf-8"))
    by = {v["tk"]: v for v in d["valuation"]}
    out = {}
    for ptk, itk in CROSS.items():
        v = by[itk]
        out[ptk] = dict(itk=itk, tp=v["center"], tp_src=f"AI 算力地圖 {d['meta']['version']} data.json → {v['c1src']}", ccy=v["ccy"],
                        gate=v.get("gate"), gsum=v.get("gsum"), glabel=v.get("glabel"), name=v["name"], mfile=v.get("mfile"))
    return out, d["meta"]["version"]


def main():
    pos = rcsv("positions.csv"); segs = rcsv("segments.csv"); watch = rcsv("watchlist.csv"); params = rcsv("params.csv")
    prices = {r["tk"]: r for r in rcsv("prices.csv")}
    L = line_values(); OM = own_models(); SM = seg_models(); CM, imv = cross_models()
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "說明"
    notes = [f"實體 AI 地圖 資料活頁簿 {VERSION}（{DATE[:4]}-{DATE[4:6]}-{DATE[6:]}）— 事實與判斷的權威位置；網站只呈現本檔重算後的值。",
             "黃底＝判斷輸入（可改）；其餘為抄錄值或公式；每列出處欄寫明檔案｜工作表!儲存格。",
             "產業地位抄錄自 AI實體應用 專案產業地位活頁簿（data/SOURCE.md）；六條線數字只讀 physical_ai/lines/*/dist 成品；個股估值只讀個股模型成品。",
             "跨圖個股：主圖＝持有個股模型的專案所屬地圖；目標價只讀主圖模型（AI 算力地圖 data.json 所引模型儲存格）。",
             "上檔空間＝最新一份模型目標價 ÷ 最新收盤 − 1；收盤一律同一日（data/prices.csv）。五級：> +50% 強力買進；+25%～+50% 買進；−25%～+25% 中立；−50%～−25% 放空；< −50% 強力放空。",
             "收支關卡：評價為買進以上且收支關卡（模型 FY2026–FY2030 自由現金流扣股利，情境 B 合計）< 0 者，最終分級改列「收支不足」；放空、強力放空、中立維持。"]
    for i, t in enumerate(notes, 1):
        ws.cell(i, 1, t)
    ws.column_dimensions["A"].width = 150

    # 參數（黃底判斷）
    P = wb.create_sheet("參數")
    for c, h in enumerate(["名稱", "值", "說明"], 1):
        P.cell(1, c, h).font = BOLD
    prm = [("t_sb", 0.50, "強力買進門檻（上檔空間 >）"), ("t_b", 0.25, "買進門檻（>）"), ("t_s", -0.25, "中立下緣（≥）"), ("t_ss", -0.50, "放空下緣（≥）")]
    for p in params:
        if not re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", p["名稱"]):
            continue
        prm.append((p["名稱"], float(p["值"]), p.get("說明（與 AI 投資地圖相同）") or p.get("說明", "")))
    for i, (k, v, n) in enumerate(prm, 2):
        P.cell(i, 1, k); c = P.cell(i, 2, v); c.fill = YEL; c.font = BLUE; P.cell(i, 3, n)
        wb.defined_names[k] = DefinedName(k, attr_text=f"參數!$B${i}")
    P.column_dimensions["A"].width = 12; P.column_dimensions["C"].width = 60

    # 線
    W = wb.create_sheet("線")
    hd = ["線", "線名", "關鍵問題", "指標", "低", "基準", "高", "單位", "格式", "出處"]
    for c, h in enumerate(hd, 1):
        W.cell(1, c, h).font = BOLD; W.cell(1, c).fill = HDR
    for i, x in enumerate(L, 2):
        nm, q = LINE_META[x["code"]]
        for c, v in enumerate([x["code"], nm, q, x["lab"], x["lo"], x["base"], x["hi"], x["unit"], x["fmt"], x["src"]], 1):
            W.cell(i, c, v)
    W.column_dimensions["D"].width = 44; W.column_dimensions["J"].width = 80

    # 地圖分組（判斷：線 6 共用零組件拆為 6A–6D；黃底）
    import csv as _csv
    GR = wb.create_sheet("地圖分組")
    grows = list(_csv.DictReader(open(os.path.join(HERE, "data", "map_groups.csv"), encoding="utf-8")))
    gh = list(grows[0].keys()) + ["標記"]
    for c, h in enumerate(gh, 1):
        GR.cell(1, c, h).font = BOLD; GR.cell(1, c).fill = HDR
    for i, r in enumerate(grows, 2):
        for c, h in enumerate(gh[:-1], 1):
            GR.cell(i, c, r[h]).fill = YEL
        GR.cell(i, len(gh), "判斷（2026-11-17；線 6 拆 6A–6D 為地圖呈現分組，不影響地位分數與估值）")
    allsub = {x for r in grows for x in r["次產業代碼"].split()}
    assert allsub == {r["代碼"] for r in segs}, ("地圖分組未涵蓋全部次產業", allsub ^ {r["代碼"] for r in segs})
    GR.column_dimensions["C"].width = 50; GR.column_dimensions["E"].width = 22

    # 次產業
    S = wb.create_sheet("次產業")
    sh = list(segs[0].keys())
    for c, h in enumerate(sh, 1):
        S.cell(1, c, h).font = BOLD; S.cell(1, c).fill = HDR
    for i, r in enumerate(segs, 2):
        for c, h in enumerate(sh, 1):
            v = r[h]
            try:
                v = float(v) if h in ("人形 BOM 占比中值", "不可替代性（1–3）", "價值鏈分量分", "覆蓋公司數") and v != "" else v
            except ValueError:
                pass
            S.cell(i, c, v)

    # 個股地位（分數以公式重算，與抄錄值核對）
    G = wb.create_sheet("個股地位")
    gh = ["代號", "公司", "市場", "亞太掛牌", "層級", "次產業代碼", "次產業", "所屬線", "子分類", "公司重要性", "市場位置", "次產業不可替代性", "價值鏈分量分",
          "不可替代性分", "市場位置分", "地位分數", "排名", "抄錄地位分數", "檢查", "地圖歸屬", "地位依據", "標記", "實體 AI 營收曝險"]
    for c, h in enumerate(gh, 1):
        G.cell(1, c, h).font = BOLD; G.cell(1, c).fill = HDR
    n = len(pos)
    sc = lambda x: f"IF({x}=1,s_irr1,IF({x}=2,s_irr2,s_irr3))"
    for i, r in enumerate(pos, 2):
        vals = [r["代號"], r["公司"], r["市場"], r["亞太掛牌"], r["層級"], r["次產業代碼"], r["次產業"], r["所屬線"], r["子分類"],
                float(r["公司重要性（1–3）"]), float(r["市場位置（1–4）"]), float(r["次產業不可替代性"]), float(r["價值鏈分量分"])]
        for c, v in enumerate(vals, 1):
            G.cell(i, c, v)
        G.cell(i, 14, f"=({sc(f'L{i}')}+{sc(f'J{i}')})/2")
        G.cell(i, 15, f"=CHOOSE(K{i},s_pos1,s_pos2,s_pos3,s_pos4)")
        G.cell(i, 16, f"=w_irr*N{i}+w_pos*O{i}+w_vc*M{i}")
        G.cell(i, 17, f"=RANK(P{i},$P$2:$P${n + 1},0)")
        G.cell(i, 18, float(r["地位分數"]))
        G.cell(i, 19, f'=IF(ABS(P{i}-R{i})<0.001,"OK","差異")')
        for c, h in zip((20, 21, 22, 23), ("地圖歸屬", "地位依據", "標記", "實體 AI 營收曝險（判斷，不進分數）")):
            G.cell(i, c, r[h])

    # 估值
    V = wb.create_sheet("估值")
    vh = ["代號", "公司", "主圖", "幣別", "目標價", "目標價出處", "收盤", "收盤日", "收盤出處", "上檔空間", "估值評等", "收支關卡（情境 B 合計）", "收支出處",
          "最終分級", "模型現價", "模型現價日", "模型檔", "台股不公開目標價"]
    for c, h in enumerate(vh, 1):
        V.cell(1, c, h).font = BOLD; V.cell(1, c).fill = HDR
    rate = lambda u: f'IF({u}>t_sb,"強力買進",IF({u}>t_b,"買進",IF({u}>=t_s,"中立",IF({u}>=t_ss,"放空","強力放空"))))'
    name_of = {r["代號"]: r["公司"] for r in pos}
    i = 2
    for tk, m in OM.items():
        p = prices[tk]
        row = [tk, name_of[tk], "實體 AI 地圖", m["ccy"], m["tp"], m["tp_src"], float(p["close"]), p["date"], p["source"]]
        for c, v in enumerate(row, 1):
            V.cell(i, c, v)
        V.cell(i, 10, f"=E{i}/G{i}-1"); V.cell(i, 11, "=" + rate(f"J{i}"))
        V.cell(i, 12, m["gate"]); V.cell(i, 13, m["gate_src"])
        V.cell(i, 14, f'=IF(AND(L{i}<0,OR(K{i}="強力買進",K{i}="買進")),"收支不足",K{i})')
        V.cell(i, 15, m["mpx"]); V.cell(i, 16, m["mpd"]); V.cell(i, 17, m["file"]); V.cell(i, 18, "是" if m["ccy"] == "TWD" else "否")
        i += 1
    for ptk, m in CM.items():
        p = prices[m["itk"]]
        row = [ptk, name_of[ptk], "AI 算力地圖（跨圖）", m["ccy"], m["tp"], m["tp_src"], float(p["close"]), p["date"], p["source"]]
        for c, v in enumerate(row, 1):
            V.cell(i, c, v)
        V.cell(i, 10, f"=E{i}/G{i}-1"); V.cell(i, 11, "=" + rate(f"J{i}"))
        V.cell(i, 12, m["gsum"]); V.cell(i, 13, f"AI 算力地圖 {imv} data.json：{m['gate']}；{m['glabel']}")
        V.cell(i, 14, f'=IF(AND(N(L{i})<0,OR(K{i}="強力買進",K{i}="買進")),"收支不足",K{i})')
        V.cell(i, 17, m["mfile"]); V.cell(i, 18, "是" if m["ccy"] == "TWD" else "否")
        i += 1
    VEND = i - 1
    V.column_dimensions["F"].width = 60; V.column_dimensions["I"].width = 50

    # 未上市分部
    U = wb.create_sheet("分部")
    for c, h in enumerate(["代號", "公司", "低（情境 C）", "基準（加權）", "高（情境 A）", "單位", "說明", "出處低", "出處基準", "出處高", "模型檔"], 1):
        U.cell(1, c, h).font = BOLD; U.cell(1, c).fill = HDR
    for j, (tk, m) in enumerate(SM.items(), 2):
        for c, v in enumerate([tk, name_of[tk], m["lo"][0], m["base"][0], m["hi"][0], "十億美元", m["note"], m["lo"][1], m["base"][1], m["hi"][1], m["file"]], 1):
            U.cell(j, c, v)

    # 觀察名單
    O = wb.create_sheet("觀察名單")
    oh = list(watch[0].keys())
    for c, h in enumerate(oh, 1):
        O.cell(1, c, h).font = BOLD
    for j, r in enumerate(watch, 2):
        for c, h in enumerate(oh, 1):
            O.cell(j, c, r[h])
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)
    # LibreOffice 重算後覆寫，讓公式格帶快取值（公式格快取值檢查）
    import shutil, subprocess, tempfile
    tmp = tempfile.mkdtemp()
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", tmp, OUT], check=True, capture_output=True)
    shutil.copy(os.path.join(tmp, os.path.basename(OUT)), OUT); shutil.rmtree(tmp, ignore_errors=True)
    print(OUT, "線指標", len(L), "個股", len(pos), "估值", VEND - 1)


if __name__ == "__main__":
    main()
