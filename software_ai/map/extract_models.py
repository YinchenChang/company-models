"""AI 軟體地圖：從 software_ai/<公司>/dist/ 最新個股模型成品（xlsx）抽估值欄位，並讀相對倍數與單位經濟。

輸出 software_ai/map/data/models_extract.json：每家一筆，每個值附「檔案、工作表!儲存格」。
本程式只讀、不計算估值；以「列標籤」定位（不寫死儲存格位置），各公司版面不同時以多個標籤樣式依序嘗試。
缺值（例如 v0.1 成品沒有股票薪酬加回 DCF）一律寫 null，網頁顯示「—」。

最新成品＝<公司>/dist/ 下（不含 _待刪除/ 與其他子資料夾）依「日期＋版號」排序最大的 xlsx；
HTML 同法另取（可能比 xlsx 舊，記 html_matches）。

用法：python3 software_ai/map/extract_models.py
"""
import datetime
import glob
import json
import os
import re
import shutil
import subprocess
import tempfile

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))              # repo 根目錄
SW = os.path.join(ROOT, "software_ai")
OUT = os.path.join(HERE, "data", "models_extract.json")

# 資料夾 → 代號（與 AI軟體地圖_產業地位 的「代號」欄一致）
MODELS = [
    ("snowflake", "SNOW"), ("datadog", "DDOG"), ("palantir", "PLTR"),
    ("crowdstrike", "CRWD"), ("paloalto", "PANW"), ("trendmicro", "4704"),
    ("salesforce", "CRM"), ("hubspot", "HUBS"),
    ("accenture", "ACN"), ("fujitsu", "6702"), ("nri", "4307"),
    ("applovin", "APP"), ("naver", "035420"), ("lycorp", "4689"),
]

NAME_RE = re.compile(r"^(\d{8})_(.+)_v(\d+(?:\.\d+)*)\.(xlsx|html)$")


def art_key(fn):
    m = NAME_RE.match(fn)
    return (m.group(1), tuple(int(x) for x in m.group(3).split("."))) if m else None


INCLUDE_SUPERSEDED = False     # 只供開發測試（--include-superseded）：連 _待刪除/ 也找，驗證舊版面（v0.1）讀得通


def latest(folder, ext):
    """dist/ 頂層（排除 _待刪除/）最新成品；回傳 (路徑, 日期, 版號, 標題) 或 None。"""
    best = None
    pats = [os.path.join(SW, folder, "dist", "*." + ext)]
    if INCLUDE_SUPERSEDED:
        pats.append(os.path.join(SW, folder, "dist", "_待刪除", "*." + ext))
    for p in (q for pat in pats for q in glob.glob(pat)):
        fn = os.path.basename(p)
        k = art_key(fn)
        if not k or ("_待刪除" in p and not INCLUDE_SUPERSEDED):
            continue
        if best is None or k > best[0]:
            m = NAME_RE.match(fn)
            best = (k, p, m.group(1), m.group(3), m.group(2))
    return None if best is None else best[1:]


def rel(p):
    return os.path.relpath(p, ROOT)


def recalc(path, tmp):
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", tmp, path],
                   check=True, capture_output=True)
    return os.path.join(tmp, os.path.basename(path))


class Book:
    def __init__(self, path, tmp):
        self.path = path
        self.rel = rel(path)
        wb = openpyxl.load_workbook(path, data_only=True)
        self.recalculated = False
        if "Val" in wb.sheetnames and not any(isinstance(c.value, (int, float)) for r in wb["Val"].iter_rows(max_row=30) for c in r):
            wb = openpyxl.load_workbook(recalc(path, tmp), data_only=True)   # 無快取值才重算
            self.recalculated = True
        self.wb = wb

    def find(self, sheet, pat, excl=None, kind="num", last=False, single=False):
        """第一個（或最後一個）A 欄標籤符合 pat、且其後有指定型別值的列。回傳 {values:[(儲存格,值)], label, row}。"""
        if sheet not in self.wb.sheetnames:
            return None
        hit = None
        for row in self.wb[sheet].iter_rows():
            a = row[0].value
            if not isinstance(a, str) or not re.search(pat, a):
                continue
            if excl and re.search(excl, a):
                continue
            vals = []
            for c in row[1:]:
                v = c.value
                if kind == "num" and isinstance(v, (int, float)) and not isinstance(v, bool):
                    vals.append((c.coordinate, v))
                elif kind == "str" and isinstance(v, str) and v.strip() and v.strip() not in ("—", "-", "未找到", "未取得"):
                    vals.append((c.coordinate, v.strip()))
                elif kind == "any" and v is not None:
                    vals.append((c.coordinate, v))
            if single and len(vals) != 1 and not re.search(r"情境加權", a):   # 情境 A／B／C 三欄的列不是「情境加權」單值
                continue
            if vals:
                hit = {"values": vals, "label": a, "row": row[0].row}
                if not last:
                    return hit
        return hit

    def one(self, sheet, pat, excl=None, kind="num", k=0, single=False):
        h = self.find(sheet, pat, excl, kind, single=single)
        if not h or len(h["values"]) <= k:
            return None
        c, v = h["values"][k]
        return {"value": v, "cell": f"{sheet}!{c}", "file": self.rel, "label": h["label"][:60]}

    def abc(self, sheet, pat, excl=None):
        """A／B／C 三情境：標籤列後的前三個數值（版面把 A、B、C 放在連續三欄）。"""
        h = self.find(sheet, pat, excl)
        if not h or len(h["values"]) < 3:
            return None
        vals = h["values"][-3:] if len(h["values"]) > 3 else h["values"][:3]
        return {"values": [v for _, v in vals], "cells": [f"{sheet}!{c}" for c, _ in vals],
                "file": self.rel, "label": h["label"][:60]}


def iso(d):
    return d.strftime("%Y-%m-%d") if isinstance(d, (datetime.datetime, datetime.date)) else None


def price_block(b):
    h = b.find("Val", r"^現價(?!隱含)")
    if not h:
        return None
    c, v = h["values"][0]
    date = None
    row = b.wb["Val"][h["row"]]
    for cell in row:                                           # 同列的日期格
        if isinstance(cell.value, (datetime.datetime, datetime.date)):
            date = iso(cell.value)
    if not date:
        m = re.search(r"(\d{4}-\d{2}-\d{2})", h["label"])      # 標籤內的日期
        date = m.group(1) if m else None
    if not date:                                               # 說明欄寫「股價基準日 2026-11-13」
        for cell in row:
            if isinstance(cell.value, str):
                m = re.search(r"股價基準日\s*(\d{4}-\d{2}-\d{2})", cell.value)
                if m:
                    date = m.group(1)
    if not date:                                               # 獨立的「股價基準日」列
        d = b.find("Val", r"^股價基準日", kind="any")
        if d and isinstance(d["values"][0][1], (datetime.datetime, datetime.date)):
            date = iso(d["values"][0][1])
    return {"value": v, "cell": f"Val!{c}", "file": b.rel, "label": h["label"][:60], "date": date}


SBC_LABEL = r"^(對照：)?DCF.*(加回|SBC)"
SBC_EXCL = r"每股|價值|不加回|差|相對|−|若 DCF|情境 DCF|A／B／C|上檔"


def sbc_dcf(b):
    """股票薪酬加回 DCF 目標價（v0.2 起；只作對照，不進加權目標價）。
    回傳 (情境加權單值, A／B／C 三值, 說明)。先找 Val 頁，再找 Recon 頁（部分模型放對照讀數於 Recon）。
    版面：B 單值＝情境加權；B＋C/D/E 四值＝B 為情境加權；只有 A／B／C 三值＝只回三值（不自行加權）。"""
    wsum = abc = note = None
    for sheet in ("Val", "Recon"):
        if sheet not in b.wb.sheetnames:
            continue
        for row in b.wb[sheet].iter_rows():
            a = row[0].value
            if not isinstance(a, str) or not re.search(SBC_LABEL, a) or "目標價" not in a or re.search(SBC_EXCL, a):
                if isinstance(a, str) and re.search(r"^情境 DCF（股票薪酬加回）目標價", a) and abc is None:
                    nums = [(c.coordinate, c.value) for c in row[1:] if isinstance(c.value, (int, float))]
                    if len(nums) >= 3:
                        abc = {"values": [v for _, v in nums[:3]], "cells": [f"{sheet}!{c}" for c, _ in nums[:3]],
                               "file": b.rel, "label": a[:70]}
                continue
            nums = [(c.coordinate, c.value) for c in row[1:] if isinstance(c.value, (int, float)) and not isinstance(c.value, bool)]
            txt = [c.value for c in row[1:] if isinstance(c.value, str) and c.value.strip() in ("未找到", "未取得")]
            if not nums:
                if txt and note is None:
                    note = txt[0].strip()
                continue
            if len(nums) == 3 and nums[0][0].startswith("B") and "加權" not in a:
                if abc is None:
                    abc = {"values": [v for _, v in nums], "cells": [f"{sheet}!{c}" for c, _ in nums], "file": b.rel, "label": a[:70]}
                continue
            if nums[0][0].startswith("B") and wsum is None:
                c, v = nums[0]
                wsum = {"value": v, "cell": f"{sheet}!{c}", "file": b.rel, "label": a[:70]}
        if wsum:
            break
    return wsum, abc, note


def ai_net(b):
    """AI 淨效果（命題模組）：Val 頁標籤含「AI 淨效果」的第一列，A／B／C 三值；有則抽。"""
    for sheet in ("Val", "Thesis"):
        if sheet not in b.wb.sheetnames:
            continue
        for row in b.wb[sheet].iter_rows():
            a = row[0].value
            if isinstance(a, str) and re.search(r"^AI 淨效果.*A／B／C", a):
                nums = [(c.coordinate, c.value) for c in row[1:] if isinstance(c.value, (int, float))]
                if len(nums) >= 3:
                    return {"values": [v for _, v in nums[:3]], "cells": [f"{sheet}!{c}" for c, _ in nums[:3]],
                            "file": b.rel, "label": a[:70]}
    return None


def gate_block(b):
    g = {"basis": b.one("Val", r"^入不敷出判定用"), "abc": None, "verdict": None}
    h = b.abc("Val", r"^收支關卡.*A／B／C") or b.abc("Val", r"^收支關卡（FY.*合計.*情境 A／B／C")
    if h:
        g["abc"] = h
    else:                                                      # 以三列「情境 A／C 現金收支合計」補（LY）
        pass
    v = b.one("Val", r"^收支關卡判定|^入不敷出判定$", kind="str")
    if v:
        g["verdict"] = v
    return g


def scen_prices(b):
    for pat in (r"^情境加權目標價（各情境", r"^各情境加權目標價", r"^情境目標價 A／B／C", r"^情境目標價＝", r"^情境目標價（.*A／B／C"):
        h = b.abc("Val", pat)
        if h:
            return h
    # 三列：情境 A 目標價／情境 B 目標價／情境 C 目標價
    vals = []
    for s in "ABC":
        o = b.one("Val", rf"^情境 {s} 目標價")
        if not o:
            return None
        vals.append(o)
    return {"values": [o["value"] for o in vals], "cells": [o["cell"] for o in vals], "file": b.rel, "label": "情境 A／B／C 目標價"}


def read_model(folder, tk, tmp):
    x = latest(folder, "xlsx")
    rec = {"folder": folder, "ticker": tk, "status": "ok"}
    h = latest(folder, "html")
    if h:
        rec["html"] = {"file": rel(h[0]), "name": os.path.basename(h[0]), "date": h[1], "version": "v" + h[2], "title": h[3]}
    if x is None:
        rec["status"] = "dist 無現行成品（改版中；_待刪除/ 不計）"
        return rec
    path, date, ver, title = x
    rec.update({"xlsx": rel(path), "date": f"{date[:4]}-{date[4:6]}-{date[6:]}", "version": "v" + ver, "title": title})
    rec["html_matches"] = bool(h and h[2] == ver and h[1] == date)
    b = Book(path, tmp)
    rec["recalculated"] = b.recalculated
    rec["price"] = price_block(b)
    rec["dcf"] = b.one("Val", r"^DCF 目標價（(三)?情境加權", r"加回", single=True)
    rec["comp"] = (b.one("Val", r"^Comparable 目標價（相對倍數調整後", single=True)       # 有「調整後」列者以調整後為進加權的那一列
                   or b.one("Val", r"^Comparable 目標價（", r"只看|A／B／C", single=True))
    rec["tp"] = b.one("Val", r"^(加權)?目標價（DCF 與.{0,6}(Comparable|同業本益比) ?加權", single=True)
    rec["upside"] = b.one("Val", r"^(上檔空間(（相對現價）|：加權目標價)|加權目標價上檔空間)", single=True)
    rec["rating"] = b.one("Val", r"^估值評等", r"觸發|強力買進或強力放空|須診斷", kind="str")
    rec["rec"] = b.one("Val", r"^投資建議(?!為)", kind="str")
    rec["gate"] = gate_block(b)
    rec["scen"] = scen_prices(b)
    rec["sbc_dcf"], rec["sbc_abc"], rec["sbc_note"] = sbc_dcf(b)
    rec["ai_net"] = ai_net(b)
    rec["missing"] = [k for k in ("price", "dcf", "comp", "tp", "upside", "rating", "rec", "scen", "sbc_dcf", "ai_net")
                      if not rec.get(k)] + (["gate"] if not rec["gate"]["basis"] else [])
    return rec


def relmult(folder):
    p = os.path.join(SW, "data", "relmult_hist", "out", folder + ".json")
    if not os.path.exists(p):
        return None
    d = json.load(open(p, encoding="utf-8"))
    g = d["groups"]["all"]
    return {"file": rel(p), "generated": d.get("generated"), "ccy": d.get("currency"), "peers": g.get("peers"),
            "peers_missing": g.get("peers_missing"), "ps_reading": g.get("ps_reading"), "psg_reading": g.get("psg_reading"),
            "geo_base": g.get("geo_base"), "base_source": g.get("base_source"), "divergence": g.get("divergence"),
            "divergence_gt_30pct": g.get("divergence_gt_30pct"), "valid_points": g.get("valid_points"),
            "sens": g.get("sens"), "gaps": d.get("gaps"), "groups": [k for k in d["groups"] if k != "all"]}


# ---------- 單位經濟 ----------
ARCH = [("L", "按人力計價", "IT 服務、BPO"), ("S", "按席位計價", "CRM、設計與生產力、人資、工作流"),
        ("A", "按廣告計價", "數位廣告"), ("U", "按用量計價", "資料平台"), ("P", "按防護範圍計價", "資安")]
YEAR_COL = {2026: "D", 2030: "H", 2032: "J"}
SUBS5 = [("IT", "ITS", "L"), ("CRM", "CRM", "S"), ("AD", "ADS", "A"), ("DATA", "DATA", "U"), ("SEC", "SEC", "P")]


def unit_econ(tmp):
    fs = sorted(glob.glob(os.path.join(SW, "unit_economics", "dist", "*.xlsx")))
    fs = [f for f in fs if art_key(os.path.basename(f))]
    if not fs:
        return None
    path = max(fs, key=lambda f: art_key(os.path.basename(f)))
    rc = recalc(path, tmp)
    wb = openpyxl.load_workbook(rc, data_only=True)
    errs = [(ws.title, c.coordinate) for ws in wb for r in ws.iter_rows() for c in r
            if isinstance(c.value, str) and (c.value.startswith("#") or c.value.startswith("Err:"))]
    m = NAME_RE.match(os.path.basename(path))
    out = {"file": rel(path), "date": f"{m.group(1)[:4]}-{m.group(1)[4:6]}-{m.group(1)[6:]}", "version": "v" + m.group(3),
           "recalc_errors": len(errs), "archetypes": [], "subs": []}

    def rowcode(ws, code):
        for row in ws.iter_rows(min_col=1, max_col=10):
            if row[0].value == code:
                return row
        return None

    for letter, name, scope in ARCH:
        ws = wb[letter]
        title = str(ws["A1"].value or "")
        o = {"code": letter, "name": name, "scope": scope, "title": title[:60]}
        for key in ("netr", "gm", "infsh", "incrr", "erosr", "infr"):
            r = rowcode(ws, key)
            if r is None:
                continue
            o[key] = {str(y): {"value": r[openpyxl.utils.column_index_from_string(c) - 1].value,
                                "cell": f"{letter}!{c}{r[0].row}"} for y, c in YEAR_COL.items()}
            o[key + "_label"] = str(r[1].value)
        out["archetypes"].append(o)
    ws = wb["次產業"]
    for code, sub, letter in SUBS5:
        o = {"code": sub, "ue_code": code, "archetype": letter}
        head = rowcode(ws, code)
        o["label"] = str(head[1].value) if head else ""
        for key in ("base", "net", "xfer", "netsh"):
            r = rowcode(ws, f"SI_{code}_{key}")
            if r is None:
                continue
            o[key] = {str(y): {"value": r[openpyxl.utils.column_index_from_string(c) - 1].value,
                                "cell": f"次產業!{c}{r[0].row}"} for y, c in YEAR_COL.items()}
            o[key + "_unit"] = str(r[2].value)
        out["subs"].append(o)
    return out


def main():
    global INCLUDE_SUPERSEDED, OUT
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--include-superseded", action="store_true", help="開發測試：連 _待刪除/ 也找（不得用於正式建置）")
    ap.add_argument("--out", help="輸出路徑（測試用）")
    a = ap.parse_args()
    INCLUDE_SUPERSEDED = a.include_superseded
    if a.out:
        OUT = a.out
    tmp = tempfile.mkdtemp()
    try:
        models = {}
        for folder, tk in MODELS:
            rec = read_model(folder, tk, tmp)
            rec["relmult"] = relmult(folder)
            models[tk] = rec
        out = {"generated": datetime.date.today().isoformat(), "models": models, "unit_econ": unit_econ(tmp)}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    print(f"寫出 {rel(OUT)}：{len(models)} 家")
    for tk, m in models.items():
        print(f"  {tk:7s} {m['folder']:12s} {m.get('version','—'):7s} {m['status']}  缺值：{m.get('missing', '—')}")


if __name__ == "__main__":
    main()
