"""AI 投資地圖：從各公司模型成品（<公司>/dist/*.xlsx 的重算值）抽出估值欄位。

輸出 data/models_extract.json：每家公司一筆，每個數字附「檔案、工作表!儲存格」。
本程式只讀、不計算；換算（空間、分級、地位分數）在 AI投資地圖_資料.xlsx 以公式完成。

用法：python3 markets/invest_map/extract_models.py [--avgo <avgo dist xlsx 路徑>]
AVGO 模型已合併（PR #107，avgo/dist/）；預設讀 avgo/dist，--avgo 只在要改讀其他成品時使用。
v0.4（2026-10-25）：半導體 13 家改讀估值 v2 的統一列標籤（semis_v2）；舊的 Andy 合理 P/E、規則值專用讀法已移除
（v2 模型不再有「對照：以 Andy 合理 P/E」一節）。
"""
import argparse
import glob
import json
import os
import re

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "models_extract.json")


def dist_xlsx(folder):
    files = sorted(glob.glob(os.path.join(ROOT, folder, "dist", "*.xlsx")))
    if len(files) != 1:
        raise SystemExit(f"{folder}/dist 應恰有一個 xlsx，實際 {len(files)}")
    return files[0]


class Book:
    def __init__(self, path):
        self.path = path
        self.rel = os.path.relpath(path, ROOT)
        self.wb = openpyxl.load_workbook(path, data_only=True)

    def rows(self, sheet):
        ws = self.wb[sheet]
        for row in ws.iter_rows():
            cells = [(c.coordinate, c.value) for c in row if c.value is not None]
            if cells:
                yield cells

    def find(self, sheet, *preds):
        """回傳第一列：各條件 (欄序, 正規式) 皆成立。欄序以非空格計。"""
        for cells in self.rows(sheet):
            ok = True
            for idx, pat in preds:
                if idx >= len(cells) or not re.search(pat, str(cells[idx][1])):
                    ok = False
                    break
            if ok:
                return cells
        raise KeyError(f"{self.rel} {sheet} 找不到 {preds}")

    def num(self, sheet, preds, k):
        """該列第 k 個數值（0 起算）與其儲存格。"""
        cells = self.find(sheet, *preds)
        nums = [(c, v) for c, v in cells if isinstance(v, (int, float)) and not isinstance(v, bool)]
        c, v = nums[k]
        return {"value": v, "cell": f"{sheet}!{c}", "file": self.rel}


def lit(value, src):
    return {"value": value, "cell": None, "file": None, "source": src}


def model_px(b, sheet, label_pat):
    """模型輸入頁的股價列：回傳 {value, cell, file, date}。"""
    cells = b.find(sheet, (1, label_pat)) if sheet == "In" else b.find(sheet, (0, label_pat))
    lbl = [v for c, v in cells if isinstance(v, str) and re.search(r"\d{4}-\d{2}-\d{2}", v)]
    nums = [(c, v) for c, v in cells if isinstance(v, (int, float))]
    c, v = nums[0]
    date = re.search(r"(\d{4}-\d{2}-\d{2})", lbl[0]).group(1) if lbl else ""
    return {"value": v, "cell": f"{sheet}!{c}", "file": b.rel, "date": date}


def first_num_row(b, sheet, pat):
    """第一個「首格符合 pat 且至少有一個數值」的列：回傳 [(儲存格, 值), ...]（只含數值）。"""
    for cells in b.rows(sheet):
        if re.search(pat, str(cells[0][1])):
            nums = [(c, v) for c, v in cells[1:] if isinstance(v, (int, float)) and not isinstance(v, bool)]
            if nums:
                return nums
    raise KeyError(f"{b.rel} {sheet} 找不到含數值的列 {pat}")


def cell(b, sheet, c, v):
    return {"value": v, "cell": f"{sheet}!{c}", "file": b.rel}


# 估值 v2（AI半導體 2026-10-25，company-models #133–#140）：Val 頁「0. 目標價」區塊的統一列標籤
V2_TP = r"^目標價（DCF 與 Comparable 加權）"
V2_DCF = r"^DCF 目標價（情境加權）"
V2_COMP = r"^Comparable 目標價（市場主流同業法）"
V2_CONS = r"^市場共識目標價"
V2_SCEN12 = r"DCF 情境價值推到 12 個月後"          # A／B／C 三欄（部分樣板無此列 → 改讀價格帶）
V2_BAND = r"^價格帶（DCF 不如預期～順利）"          # 數值版：C 情境、A 情境
V2_OLD = r"^(規則值情境加權價|舊法隱含價（情境加權)"
V2_UP = r"^上檔空間：加權目標價"                    # 模型自算（÷ 模型現價 − 1）；網站直接引用
V2_UPC = r"^上檔空間：Comparable 目標價"
V2_TIER6 = r"^上檔空間級距"                          # 文字列（#164 起；尚未補的模型沒有）
V2_REC = r"^投資建議"                                # 文字列（#164 起；尚未補的模型沒有）                      # 舊法（參考，非目標價）


def text_cell(b, sheet, pat):
    """首格符合 pat 的列，取其後第一個文字值（投資建議、級距等）；沒有該列回傳 None。"""
    for cells in b.rows(sheet):
        if re.search(pat, str(cells[0][1])):
            for c, v in cells[1:]:
                if isinstance(v, str) and v.strip() and v.strip() not in ("—", "-", "US$", "US$B", "%", "x", "倍"):  # 跳過單位欄
                    return {"value": v.strip(), "cell": f"{sheet}!{c}", "file": b.rel}
            return None
    return None


def semis_px(b):
    """模型現價（上檔空間的分母）：Val 頁「現價（日期…）」優先；記憶體、台積電等放 In 頁。"""
    for sheet, pat in (("Val", r"^現價（\d{4}-\d{2}-\d{2}"), ("In", r"股價（"), ("In", r"現價（"), ("Val", r"^現價（")):
        if sheet not in b.wb.sheetnames:
            continue
        try:
            return model_px(b, sheet, pat)
        except (KeyError, IndexError):
            continue
    raise KeyError(f"{b.rel} 找不到模型現價")


def semis_v2(b):
    """估值 v2 半導體模型：加權目標價、DCF、Comparable、共識、C 情境 DCF 價值（12 個月後）、舊法參考價、現價。"""
    out = {}
    for k, pat in (("tp", V2_TP), ("tp_dcf", V2_DCF), ("tp_comp", V2_COMP), ("tp_cons", V2_CONS),
                   ("up", V2_UP), ("up_comp", V2_UPC)):
        c, v = first_num_row(b, "Val", pat)[0]
        out[k] = cell(b, "Val", c, v)
    try:
        c, v = first_num_row(b, "Val", V2_OLD)[0]
        out["old_rule"] = cell(b, "Val", c, v)
    except KeyError:
        pass                                              # 新建模型（v0.1 起即 v2）沒有舊法
    for k, pat in (("tier6", V2_TIER6), ("rec", V2_REC)):
        t = text_cell(b, "Val", pat)
        if t:
            out[k] = t
    try:
        nums = first_num_row(b, "Val", V2_SCEN12)
        out["dcf_A"], out["dcf_C"] = cell(b, "Val", *nums[0]), cell(b, "Val", *nums[2])
    except KeyError:
        nums = first_num_row(b, "Val", V2_BAND)          # B＝C 情境、C＝A 情境
        out["dcf_C"], out["dcf_A"] = cell(b, "Val", *nums[0]), cell(b, "Val", *nums[1])
    out["model_px"] = semis_px(b)
    try:
        out["pe_fy28"] = [b.num("Val", [(0, r"^現價 ÷ FY2028 EPS（模型）")], i) for i in range(3)]
    except (KeyError, IndexError):
        pass
    return out


def cash_rows(b, sheet, label_pats, hdr_row=None):
    """收支關卡：回傳 {basis, label, values:[(年,值,儲存格)], file, sheet}；依序找第一個存在的列標籤。"""
    ws = b.wb[sheet]
    for pat in label_pats:
        for row in ws.iter_rows():
            v = row[0].value
            if isinstance(v, str) and re.match(pat, v):
                r = row[0].row
                hdr = hdr_row or next(rr for rr in range(1, r) if str(ws.cell(rr, 1).value or "").startswith(("項目", "ID", "編號")))
                vals = []
                for c in row[1:]:
                    yr = ws.cell(hdr, c.column).value
                    if isinstance(c.value, (int, float)) and yr not in (None, "單位") \
                            and "2025" not in str(yr) and "H20" not in str(yr) \
                            and not re.search(r"v\d|對照", str(yr)):  # 只取模型期間（2026 起）的年度欄，不含舊版對照欄
                        vals.append((str(yr), c.value, f"{sheet}!{c.coordinate}"))
                return {"label": v, "values": vals, "file": b.rel}
    raise KeyError(f"{b.rel} {sheet} 找不到 {label_pats}")


NI_LABELS = {"淨利", "歸屬母公司淨利", "歸母淨利", "淨利（歸屬母公司）", "淨利（不含權益證券損益）", "淨利（non-GAAP）",
             "歸屬淨利", "稅後淨利", "non-GAAP 淨利"}  # v0.5：新模型的列名


def net_income_rows(b):
    """模型只有損益時：PL 頁所有「淨利」列（各情境、各年）的數值。"""
    ws = b.wb["PL"]
    hdr = {}
    for row in ws.iter_rows():                      # 表頭列：第一格為「情境／項目」或「項目」
        if str(row[0].value or "").startswith(("情境／項目", "項目")):
            hdr = {c.column: str(c.value) for c in row if c.value is not None}
            break
    vals, seen = [], set()
    for row in ws.iter_rows():
        v = row[0].value
        if isinstance(v, str) and v.strip() in NI_LABELS:
            for c in row[1:]:
                h = hdr.get(c.column, "")
                if not isinstance(c.value, (int, float)) or not h or re.search(r"\dA$|A（|實績|說明|單位", h):
                    continue                        # 排除實績年度（FY2025A 等）
                key = (h, round(c.value, 6))
                if key in seen:                     # 同一情境在摘要與推導區塊各出現一次，只計一次
                    continue
                seen.add(key)
                vals.append((h, c.value, f"PL!{c.coordinate}"))
    return {"label": "PL 頁各情境模型年度淨利（不含實績）", "values": vals, "file": b.rel}


def gap_plus_returns(b):
    """雲端模板統一口徑（2026-10-25）：營運缺口（不含融資、不含還本）＋股利＋回購。
    股東回饋可由公司自行停止，不算入不敷出；租金、利息、稅、資本支出照扣。"""
    gap = cash_rows(b, "各期收支", [r"^營運缺口（不含融資"])
    vals = [list(x) for x in gap["values"]]
    parts = ["營運缺口"]
    for pat, name in ((r"^⑧ 股利", "股利"), (r"^⑨ 回購", "回購")):
        try:
            add = cash_rows(b, "各期收支", [pat])
        except KeyError:
            continue
        amap = {y: (v, c) for y, v, c in add["values"]}
        for x in vals:
            if x[0] in amap:
                x[1] += amap[x[0]][0]
                x[2] += "＋" + amap[x[0]][1].split("!")[1]
        parts.append(name)
    label = "＋".join(parts) if len(parts) > 1 else "營運缺口（不含融資、不含還本；無股利與回購）"
    return {"basis": "現金流", "label": label,
            "values": [tuple(x) for x in vals], "file": b.rel}


def cloud(b):
    def row(lbl, k=0):
        return b.num("摘要", [(0, "^" + re.escape(lbl))], k)
    return {
        "rating": {"value": b.find("摘要", (0, r"^結論｜評等$"))[1][1], "cell": "摘要", "file": b.rel},
        "target": row("結論｜點位（加權目標價）"),
        "price": row("結論｜現價"),
        "scen_lo": row("結論｜情境區間（下緣／上緣）", 0),
        "scen_hi": row("結論｜情境區間（下緣／上緣）", 1),
        "sentence": {"value": b.find("摘要", (0, r"^結論｜結論句$"))[1][1], "cell": "摘要", "file": b.rel},
        "model_px": model_px(b, "輸入與假設", r"^現價（"),
        "cash": gap_plus_returns(b),
    }


# 半導體個股（AI半導體 模型，估值 v2）：代號 → 模型資料夾。v0.5 起 52 家（2026-10-25 Andy 18:14：個股全部接入）。
SEMIS_V2 = (("NVDA", "nvidia"), ("AVGO", "avgo"), ("AMD", "amd"), ("2454", "mediatek"), ("MRVL", "mrvl"),
            ("3661", "alchip"), ("3443", "guc"),
            ("000660", "skhynix"), ("005930", "samsung"), ("MU", "micron"), ("285A", "kioxia"), ("SNDK", "sndk"),
            ("ANET", "arista"), ("300308", "innolight"), ("COHR", "cohr"), ("LITE", "lite"), ("300502", "eoptolink"),
            ("2345", "accton"), ("FN", "fabrinet"), ("CRDO", "crdo"), ("ALAB", "alab"),
            ("2330", "tsmc"), ("2449", "kyec"), ("AMKR", "amkr"),
            ("4062", "ibiden"), ("2383", "emc"), ("3037", "unimicron"), ("8046", "nanya_pcb"), ("2368", "gce"),
            ("300476", "victory_giant"),
            ("ASML", "asml"), ("AMAT", "amat"), ("LRCX", "lrcx"), ("KLAC", "klac"), ("6857", "advantest"),
            ("6146", "disco"), ("042700", "hanmi"), ("TER", "ter"), ("BESI", "besi"), ("6315", "towa"),
            ("7729", "accretech"),
            ("2317", "foxconn"), ("2382", "quanta"), ("3231", "wistron"), ("6669", "wiwynn"), ("DELL", "dell"),
            ("SMCI", "smci"),
            ("2308", "delta"), ("3017", "avc"), ("2301", "liteon"), ("3324", "auras"), ("3653", "jentech"))
MEMORY_FCF = ("000660", "005930", "MU")   # 收支關卡用 Y 頁自由現金流（其餘先用淨利，待 AI半導體 補自由現金流列）


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--avgo", help="AVGO 成品 xlsx（預設 avgo/dist）")
    a = ap.parse_args()
    out = {}

    # --- 半導體 13 家：估值 v2 通用讀法（地圖 v0.4） ---
    for tk, folder in SEMIS_V2:
        b = Book(a.avgo if (tk == "AVGO" and a.avgo) else dist_xlsx(folder))
        out[tk] = {"model": b.rel, "v2": True, **semis_v2(b)}
    for tk in MEMORY_FCF:                              # 記憶體：收支關卡用 Y 頁自由現金流
        b = Book(dist_xlsx(dict(SEMIS_V2)[tk]))
        out[tk]["cash"] = dict(basis="現金流", **cash_rows(b, "Y", [r"^自由現金流"], hdr_row=3))

    # --- 算力平台（雲端、neocloud）：一頁摘要 ---
    for tk, folder in (("MSFT", "microsoft"), ("AMZN", "amazon"), ("GOOGL", "alphabet"),
                       ("ORCL", "oracle"), ("CRWV", "coreweave"), ("NBIS", "nebius"), ("WYFI", "whitefiber")):
        b = Book(dist_xlsx(folder))
        if "摘要" not in b.wb.sheetnames and "Val" in b.wb.sheetnames:
            # v0.6（2026-10-27）：依「個股模型原則」重建的模型（MAG v0.3、Oracle v0.4）沿用估值 v2 列標籤；
            # 收支關卡讀 CF 頁「收支關卡：自由現金流扣股利（五年合計）」的 A／B／C 三情境合計（任一情境為負即不通過，與半導體「各情境皆正」同嚴格）
            out[tk] = {"model": b.rel, "v2": True, "seg_model": True, **semis_v2(b)}
            ws = b.wb["CF"]
            for row in ws.iter_rows():
                if isinstance(row[0].value, str) and row[0].value.startswith("收支關卡：自由現金流扣股利（五年合計）"):
                    nums = [(c.coordinate, c.value) for c in row[1:] if isinstance(c.value, (int, float))][:3]
                    out[tk]["cash"] = {"basis": "現金流", "file": b.rel,
                                       "label": "自由現金流扣股利五年合計（情境 A／B／C）",
                                       "values": [(sc, v, f"CF!{c}") for sc, (c, v) in zip(("A", "B", "C"), nums)]}
                    break
            else:
                raise KeyError(f"{b.rel} CF 頁找不到收支關卡列")
        else:
            out[tk] = {"model": b.rel, **cloud(b)}

    # --- 模型商（上市） ---
    b = Book(dist_xlsx("minimax"))
    out["00100"] = {"model": b.rel,
                    "price": b.num("Valuation", [(0, "^V01$")], 0), "price_date": "2026-10-21",
                    "disc_ratio": b.num("Valuation", [(0, "^V13$")], 0)}
    b = Book(dist_xlsx("zhipu"))
    out["02513"] = {"model": b.rel,
                    "need_ratio_2028": b.num("Reverse", [(0, "^X21$")], 3),
                    "need_ratio_2028_lo": b.num("Reverse", [(0, "^X22$")], 3)}

    # --- 收支關卡：只有損益的半導體模型 ---
    for tk, folder in SEMIS_V2:
        if tk in MEMORY_FCF:
            continue
        b = Book(a.avgo if (tk == "AVGO" and a.avgo) else dist_xlsx(folder))
        out[tk]["cash"] = dict(basis="淨利", **net_income_rows(b))
    # --- 模型商 ---
    bm = Book(dist_xlsx("minimax"))
    f06 = cash_rows(bm, "Funding", [r"^F06$"])
    f07 = {y: (v, c) for y, v, c in cash_rows(bm, "Funding", [r"^F07$"])["values"]}
    out["00100"]["cash"] = {"basis": "現金流", "file": bm.rel,
                            "label": "F06 營運燒錢（經調整淨利）＋F07 算力預付與自建資本支出",
                            "values": [(y, v + f07.get(y, (0, ""))[0], c + ("＋" + f07[y][1].split("!")[1] if y in f07 else ""))
                                       for y, v, c in f06["values"]]}
    out["02513"]["cash"] = dict(basis="現金流", **cash_rows(Book(dist_xlsx("zhipu")), "Funding", [r"^F09$"]))
    out["02513"]["cash"]["label"] = "F09 自由現金流"
    # --- 未上市模型商：只取收支（無股價、無估值） ---
    for tk, folder, rowid in (("OpenAI", "openai", r"^F06$"), ("Anthropic", "anthropic", r"^F07$")):
        b = Book(dist_xlsx(folder))
        out[tk] = {"model": b.rel, "unlisted": True,
                   "cash": dict(basis="現金流", **cash_rows(b, "Funding", [rowid]))}
        out[tk]["cash"]["label"] = ("F06 " if rowid == r"^F06$" else "F07 ") + "自由現金流"
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    print(f"寫出 {OUT}：{len(out)} 家")


if __name__ == "__main__":
    main()
