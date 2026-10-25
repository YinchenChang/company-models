"""AI 投資地圖：從各公司模型成品（<公司>/dist/*.xlsx 的重算值）抽出估值欄位。

輸出 data/models_extract.json：每家公司一筆，每個數字附「檔案、工作表!儲存格」。
本程式只讀、不計算；換算（空間、分級、地位分數）在 AI投資地圖_資料.xlsx 以公式完成。

用法：python3 markets/invest_map/extract_models.py [--avgo <avgo dist xlsx 路徑>]
AVGO 模型已合併（PR #107，avgo/dist/）；預設讀 avgo/dist，--avgo 只在要改讀其他成品時使用。
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


def semis_generic(b, price_pat, rule=True, rule_scen=("A", "C")):
    """P/E 類半導體（Arista、台達電、中際旭創、廣達）：Val 頁第 1、3、5 節。"""
    out = {
        "price": b.num("Val", [(0, price_pat)], 0),
        "pe_fy28": [b.num("Val", [(0, r"^現價 ÷ FY2028 EPS")], i) for i in range(3)],
    }
    out["price_date"] = re.search(r"(\d{4}-\d{2}-\d{2})", b.find("Val", (0, price_pat))[0][1]).group(1)
    if rule:
        for s in rule_scen:
            row = b.find("Val", (0, f"^{s}$"), (2, r"^\d"))  # 第 3 節：情境列
            nums = [(c, v) for c, v in row if isinstance(v, (int, float))]
            # 第 3 節欄位：EPS…、規則倍數、隱含股價、相對現價、現價 ÷ EPS
            c, v = nums[-3]
            out[f"rule_{s}"] = {"value": v, "cell": f"Val!{c}", "file": b.rel}
            c, v = nums[-4]
            out[f"rule_mult_{s}"] = {"value": v, "cell": f"Val!{c}", "file": b.rel}
    return out


def andy_rows(b, mult_pat):
    """第 5 節（或第 4 節）「以 Andy 合理 P/E 計算的隱含價」：A 列 FY2027、C 列 FY2028。"""
    hdr_row = None
    ws = b.wb["Val"]
    for row in ws.iter_rows():
        v = row[0].value
        if v and re.search(r"對照：以 Andy 合理 P/E", str(v)):
            hdr_row = row[0].row
            mult = float(re.search(mult_pat, str(v)).group(1))
            break
    res = {"andy_pe": lit(mult, f"{b.rel} Val!A{hdr_row} 標題")}
    for r in range(hdr_row + 2, hdr_row + 5):
        s = ws.cell(r, 1).value
        vals = [(ws.cell(r, j).coordinate, ws.cell(r, j).value) for j in range(2, 9)
                if isinstance(ws.cell(r, j).value, (int, float))]
        # 欄位：FY2027 EPS、× 倍數、相對現價、FY2028 EPS、× 倍數、相對現價
        res[f"andy27_{s}"] = {"value": vals[1][1], "cell": f"Val!{vals[1][0]}", "file": b.rel}
        res[f"andy28_{s}"] = {"value": vals[4][1], "cell": f"Val!{vals[4][0]}", "file": b.rel}
    return res


def memory(b):
    out = {"pb_rule": {}, "pe_rule": {}}
    for kind, key in (("P/B", "pb_rule"), ("P/E", "pe_rule")):
        for pos in ("見頂", "中段", "谷底"):
            row = b.find("Val", (0, rf"^規則值：{pos}（{re.escape(kind)}）"))
            nums = [(c, v) for c, v in row if isinstance(v, (int, float))]
            c, v = nums[2]  # 倍數、每股基準、隱含股價
            out[key][pos] = {"value": v, "cell": f"Val!{c}", "file": b.rel}
            c2, v2 = nums[3]
            out[key][pos + "_upside"] = {"value": v2, "cell": f"Val!{c2}", "file": b.rel}
    return out


def model_px(b, sheet, label_pat):
    """模型輸入頁的股價列：回傳 {value, cell, file, date}。"""
    cells = b.find(sheet, (1, label_pat)) if sheet == "In" else b.find(sheet, (0, label_pat))
    lbl = [v for c, v in cells if isinstance(v, str) and re.search(r"\d{4}-\d{2}-\d{2}", v)]
    nums = [(c, v) for c, v in cells if isinstance(v, (int, float))]
    c, v = nums[0]
    date = re.search(r"(\d{4}-\d{2}-\d{2})", lbl[0]).group(1) if lbl else ""
    return {"value": v, "cell": f"{sheet}!{c}", "file": b.rel, "date": date}


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


NI_LABELS = {"淨利", "歸屬母公司淨利", "歸母淨利", "淨利（歸屬母公司）", "淨利（不含權益證券損益）", "淨利（non-GAAP）"}


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


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--avgo", help="AVGO 成品 xlsx（預設 avgo/dist）")
    a = ap.parse_args()
    out = {}

    # --- 運算 ---
    b = Book(dist_xlsx("nvidia"))
    d = {"price": b.num("Val", [(0, r"^現價（")], 0),
         "price_date": "2026-10-22",
         "pe_fy28": [b.num("Val", [(0, r"^現價 ÷ FY2028 EPS（模型）")], i) for i in range(3)],
         "eps_fy28_A": b.num("Val", [(0, "^A$"), (1, "FY2028"), (3, "^規則值")], 0),
         "eps_fy29_C": b.num("Val", [(0, "^C$"), (1, "FY2029"), (3, "^規則值")], 0),
         "eps_fy29_A": b.num("Val", [(0, "^A$"), (1, "FY2029"), (3, "^規則值")], 0),
         "andy_pe": lit(25, "AI半導體 先驗比對_NVDA_20261023.md N7b「合理 P/E＝現價（約 25x）」"),
         "rule_note": "規則值不適用（落後倍數含高成長期溢價；模型 Val 第 3 節）"}
    out["NVDA"] = {"model": b.rel, **d}

    if True:
        b = Book(a.avgo or dist_xlsx("avgo"))
        d = {"price": b.num("Val", [(0, r"^現價（")], 0), "price_date": "2026-10-22",
             "pe_fy28": [b.num("Val", [(0, r"^現價 ÷ FY2028 EPS（模型）")], i) for i in range(3)],
             **andy_rows(b, r"(\d+)x"),
             "rule_note": "十年倍數待前瞻 P/E 歷史（Bloomberg）；規則值未算"}
        out["AVGO"] = {"model": b.rel, **d}

    # --- 記憶體（P/B 主、P/E 對照） ---
    for tk, folder in (("000660", "skhynix"), ("005930", "samsung"), ("MU", "micron")):
        b = Book(dist_xlsx(folder))
        out[tk] = {"model": b.rel, **memory(b), "model_px": model_px(b, "In", r"股價（"),
                   "cash": dict(basis="現金流", **cash_rows(b, "Y", [r"^自由現金流"], hdr_row=3))}

    # --- 代工 ---
    b = Book(dist_xlsx("tsmc"))
    out["2330"] = {"model": b.rel, "model_px": model_px(b, "In", r"^股價（"),
                   # 列欄位：倍數、每股基準（EPS）、隱含股價、相對目前、現價 ÷ 基準
                   "rule_mid": b.num("Val", [(0, r"^規則值：中段（P/E）")], 2),
                   "rule_lo": b.num("Val", [(0, r"^規則值：低（P/E）")], 2),
                   "rule_mid_upside": b.num("Val", [(0, r"^規則值：中段（P/E）")], 3),
                   "eps27_A": b.num("Val", [(0, r"^規則值：中段（P/E）")], 1),
                   "eps28_C": b.num("Val", [(0, r"^規則值：2028 C（P/E）")], 1),
                   "andy_pe": lit(35, "AI半導體 先驗比對_台積電_20261008.md T7b「合理 P/E 35x」")}

    # --- 網通、電源、組裝 ---
    for tk, folder, mp in (("ANET", "arista", r"(\d+)x"), ("2308", "delta", r"(\d+)x"),
                           ("300308", "innolight", r"(\d+)x"), ("2382", "quanta", r"(\d+)x")):
        b = Book(dist_xlsx(folder))
        out[tk] = {"model": b.rel, **semis_generic(b, r"^現價（"), **andy_rows(b, mp)}

    # --- 載板（無規則值） ---
    b = Book(dist_xlsx("ibiden"))
    out["4062"] = {"model": b.rel, "price": b.num("Val", [(0, r"^現價（")], 0), "price_date": "2026-10-16",
                   "pe_fy28": [b.num("Val", [(0, r"^現價 ÷ FY2028 EPS（模型）")], i) for i in range(3)],
                   **andy_rows(b, r"(\d+)x"), "rule_note": "模型未算十年倍數規則值"}

    # --- 設備 ---
    b = Book(dist_xlsx("advantest"))
    out["6857"] = {"model": b.rel,
                   "price": b.num("Val", [(0, r"^現價（stockanalysis")], 0), "price_date": "2026-10-21",
                   "rule_A": b.num("Val", [(0, "^A$"), (1, "FY2027"), (3, "^規則值")], 2),
                   "rule_C": b.num("Val", [(0, "^C$"), (1, "FY2027"), (3, "^規則值")], 2)}
    b = Book(dist_xlsx("disco"))
    out["6146"] = {"model": b.rel,
                   "price": b.num("Val", [(0, r"^現價（")], 0), "price_date": "2026-10-22",
                   "rule_A": b.num("Val", [(0, "^A$"), (1, "FY2027"), (3, "^規則值")], 2),
                   "rule_C": b.num("Val", [(0, "^C$"), (1, "FY2027"), (3, "^規則值")], 2)}

    # --- 算力平台（雲端、neocloud）：一頁摘要 ---
    for tk, folder in (("MSFT", "microsoft"), ("AMZN", "amazon"), ("GOOGL", "alphabet"),
                       ("ORCL", "oracle"), ("CRWV", "coreweave"), ("NBIS", "nebius"), ("WYFI", "whitefiber")):
        b = Book(dist_xlsx(folder))
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
    for tk, folder in (("NVDA", "nvidia"), ("2330", "tsmc"), ("ANET", "arista"), ("2308", "delta"), ("300308", "innolight"),
                       ("2382", "quanta"), ("4062", "ibiden"), ("6857", "advantest"), ("6146", "disco")):
        out[tk]["cash"] = dict(basis="淨利", **net_income_rows(Book(dist_xlsx(folder))))
    out["AVGO"]["cash"] = dict(basis="淨利", **net_income_rows(Book(a.avgo or dist_xlsx("avgo"))))
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
