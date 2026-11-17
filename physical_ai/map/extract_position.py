"""抄錄實體 AI 產業地位活頁簿（權威：AI實體應用 專案 Drive 03-AI實體應用/<日期>_實體AI產業地位_v*.xlsx）→ data/*.csv。
公開 repo：去除個人與工作情境欄（可投資範圍欄改由「市場」欄推得「亞太掛牌」），人名改中性用語。
用法：python3 physical_ai/map/extract_position.py <產業地位 xlsx>"""
import csv, os, re, sys
import openpyxl
HERE = os.path.dirname(os.path.abspath(__file__))
SRC = sys.argv[1]
ASIA = re.compile(r"台|日|韓|港|中|A 股|陸|深|滬")
def clean(v):
    if v is None:
        return ""
    s = str(v)
    NM = "\u0041\u006e\u0064\u0079"   # 人名（以跳脫碼寫，避免本檔出現該字）
    s = re.sub(r"[（(]?" + NM + r"[^，；）)]*?(?=[，；）)]|$)[）)]?", "", s)  # 移除人名與其後的日期註記
    s = s.replace(NM, "")
    return s.strip()
wb = openpyxl.load_workbook(SRC, data_only=True)
def dump(sheet, out, drop=()):
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    hdr = [clean(h) for h in rows[0]]
    keep = [i for i, h in enumerate(hdr) if h and not any(d in h for d in drop)]
    with open(os.path.join(HERE, "data", out), "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        extra = ["亞太掛牌"] if "市場" in hdr else []
        w.writerow([hdr[i] for i in keep] + extra)
        n = 0
        for r in rows[1:]:
            if r[0] is None or str(r[0]).strip() == "":
                continue
            vals = [clean(r[i]) for i in keep]
            if extra:
                m = clean(r[hdr.index("市場")])
                vals.append("是" if ASIA.search(m) and not m.startswith("未上市") else "否")
            w.writerow(vals); n += 1
    print(out, n)
dump("層級與次產業", "segments.csv")
dump("個股地位", "positions.csv", drop=("基金",))
dump("觀察名單（未計分）", "watchlist.csv", drop=("基金",))
dump("參數", "params.csv")
with open(os.path.join(HERE, "data", "SOURCE.md"), "w", encoding="utf-8") as f:
    f.write(f"# 資料來源\n\n`segments.csv`、`positions.csv`、`watchlist.csv`、`params.csv` 抄錄自 AI實體應用 專案的產業地位活頁簿 `{os.path.basename(SRC)}`（權威在該專案 Drive；本 repo 只存抄錄）。\n抄錄時移除工作情境欄，另由「市場」欄推得「亞太掛牌」（台、日、韓、港、中國 A 股）。\n")
