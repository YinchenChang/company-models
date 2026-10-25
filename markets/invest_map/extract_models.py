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
        out[tk] = {"model": b.rel, **memory(b), "model_px": model_px(b, "In", r"股價（")}

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

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    print(f"寫出 {OUT}：{len(out)} 家")


if __name__ == "__main__":
    main()
