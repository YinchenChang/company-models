# 歷史相對倍數庫（PS 與 PSG）彙總活頁簿：把 relmult_hist/out/*.json 整理成一份可讀的 xlsx。
# 輸出 software_ai/data/dist/歷史相對倍數庫_PS與PSG_<日期>.xlsx。只讀 JSON，不重算。
import glob
import json
import os
import sys

import openpyxl
from openpyxl.styles import Font

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.dirname(HERE)                      # software_ai/data/
OUT = os.path.join(DATA, "relmult_hist", "out")
DIST = os.path.join(DATA, "dist")
B = Font(bold=True)


def main():
    peer = json.load(open(os.path.join(OUT, "peer_ps_20261113.json"), encoding="utf-8"))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "相對倍數"
    hdr = ["公司", "代號", "幣別", "PS 讀數", "PSG 讀數", "geo_base（採用）", "低", "高", "PS 與 PSG 差距", "差距 >30%", "PS 有效點", "PSG 有效點", "採用依據"]
    ws.append(hdr)
    for f in sorted(glob.glob(os.path.join(OUT, "*.json"))):
        if os.path.basename(f).startswith("peer_ps_"):
            continue
        d = json.load(open(f, encoding="utf-8"))
        vp = d.get("valid_points") or {}
        ws.append([d.get("company"), d.get("ticker"), d.get("currency"), d.get("ps_reading"), d.get("psg_reading"), d.get("geo_base"),
                   (d.get("sens") or {}).get("low"), (d.get("sens") or {}).get("high"), d.get("divergence"),
                   "是" if d.get("divergence_gt_30pct") else "否", vp.get("ps"), vp.get("psg"), d.get("base_source")])
    ws2 = wb.create_sheet("同業現價P／S")
    ws2.append(["說明", "同業 P/S＝市值 ÷ 最新 TTM 營收，價格日見來源 CSV；本表直接取自 peer_ps_20261113.json（as_of " + str(peer.get("as_of")) + "）"])
    ws2.append([])
    ws2.append(["公司", "代號", "群組", "本公司 P/S", "同業中位 P/S", "同業家數", "逾期家數", "逾期明細", "資料庫最近一季", "資料庫最近一季中位", "較資料庫變動"])
    for co, v in sorted(peer["companies"].items()):
        for g, gv in v["groups"].items():
            db = gv.get("db_last_quarter") or {}
            ws2.append([co, v.get("ticker"), g, v.get("own_ps"), gv.get("peer_median_ps"), gv.get("n_peers"), gv.get("n_stale"),
                        "、".join(gv.get("stale") or []), db.get("q"), db.get("median"), gv.get("change_vs_db")])
    for w in (ws, ws2):
        for c in w[1 if w is ws else 3]:
            c.font = B
        for col in w.columns:
            w.column_dimensions[col[0].column_letter].width = 16
    os.makedirs(DIST, exist_ok=True)
    p = os.path.join(DIST, "歷史相對倍數庫_PS與PSG_20261113.xlsx")
    wb.save(p)
    print("寫出", p, "；公司", ws.max_row - 1, "家；同業列", ws2.max_row - 3)
    return 0 if ws.max_row > 1 and ws2.max_row > 3 else 1


if __name__ == "__main__":
    sys.exit(main())
