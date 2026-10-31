# Wattonomics v0.1.7 資金缺口（Gap 頁）取數：讀各個股模型的基準情境年度值，寫 builder/seed/gap_models_v0.1.7.json。
# 用法：python3 wattonomics/tools/import_gap_models.py [--out wattonomics/builder/seed/gap_models_v0.1.7.json]
#
# 原則（工作單 §2e）：
# - 模型路徑讀 markets/invest_map/data/models_extract.json 的 `model`（main 上的現行版）。
# - 每家每列的儲存格都是逐家人工開檔確認後寫死在下方 CONF（不靠關鍵字自動抓第一個命中）；
#   執行時再核對「情境 B 區塊標題、年度表頭、列標籤原文」三者，任何一項對不上即停（代碼 1），不猜。
# - 某格空白或某年模型沒有（例如 Alphabet、Meta FY2026 的 AI／非 AI 拆分只有 Q4）→ 寫 null 並標「無模型值」，不外推、不補估。
# - 多格組成的值（例如 CoreWeave 營業現金流＝營運缺口＋CapEx＋JV 出資＋overlay）任一組成格非數值 → 該年 null。
# - 值一律保留模型原單位，再依 scale 換成 US$B（WhiteFiber 為 US$m ÷ 1000）；智譜保留 RMB 億，由 Excel（In gfx_cny）換算。
import argparse, json, os, subprocess, sys
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
GL = os.path.dirname(HERE)
ROOT = os.path.dirname(GL)
EXTRACT = os.path.join(ROOT, "markets", "invest_map", "data", "models_extract.json")

# 欄位 → 財年。MAG 型：C=FY2024A、D=FY2025A、E=9M26A、F=Q4 26E、G=FY2026E、H..M=FY2027E..FY2032E（Amazon、Alphabet、Meta）
MAG7 = {"C": 2024, "D": 2025, "G": 2026, "H": 2027, "I": 2028, "J": 2029, "K": 2030, "L": 2031, "M": 2032}
MAG7_FWD = {k: v for k, v in MAG7.items() if v >= 2027}
NEO = {"C": 2026, "D": 2027, "E": 2028, "F": 2029, "G": 2030}
LAB = {"D": 2025, "E": 2026, "F": 2027, "G": 2028, "H": 2029, "I": 2030}
LAB_FWD = {k: v for k, v in LAB.items() if v >= 2026}

# 每家：key → dict(ticker, family, fye, unit, scale, tag, sheet, block=(儲存格, 須含字串), header=(列, {欄: 須等於}), rows={項目: [...]}, cash=...)
# rows 每項：dict(cells=[(列, 標籤須含)], cols={欄: 財年}, note)；cells 多於一格時為加總（全部為數值才算）
CONF = {
    "msft": dict(ticker="MSFT", family="MAG 型（CF 頁）", fye=6, unit="US$B", scale=1, sheet="CF",
                 tag="Derived（個股模型基準情境 B；FY2026A 為財報）",
                 block=("A46", "情境 B（基準）"), header=(47, {"E": "FY2026A", "F": "FY2027", "K": "FY2032"}),
                 rows=dict(
                     ocf=dict(cells=[(55, "營業現金流")], cols={"E": 2026, "F": 2027, "G": 2028, "H": 2029, "I": 2030, "J": 2031, "K": 2032},
                              note="工作單寫『MSFT 已確認 CF!A23』，但 A23 是情境 A（樂觀）區塊；基準情境 B 為第 55 列，本表採 B"),
                     capai=dict(cells=[(56, "AI 資本支出（Thesis）")], cols={"F": 2027, "G": 2028, "H": 2029, "I": 2030, "J": 2031, "K": 2032}),
                     nai=dict(cells=[(57, "非 AI 資本支出")], cols={"E": 2026, "G": 2028, "H": 2029, "I": 2030, "J": 2031, "K": 2032},
                              note="第 57 列；FY2027 另加第 58 列（補足指引下限的其他資本支出，不增 AI MW，模型以非 AI 折舊；只有 FY2027 有值）"),
                     nai27=dict(cells=[(57, "非 AI 資本支出"), (58, "補足到公司指引下限的其他資本支出")], cols={"F": 2027}, merge_into="nai"),
                     cap=dict(cells=[(59, "資本支出合計（公司口徑，含融資租賃）")], cols={"E": 2026, "F": 2027, "G": 2028, "H": 2029, "I": 2030, "J": 2031, "K": 2032}),
                     bb=dict(cells=[(64, "股利"), (67, "回購")], cols={"E": 2026, "F": 2027, "G": 2028, "H": 2029, "I": 2030, "J": 2031, "K": 2032}),
                 ),
                 cash=dict(cell="E69", label="現金與短期投資（年底）", asof="2026-06-30（FY2026A 年底；模型無 2025-12-31）")),
    "googl": dict(ticker="GOOGL", family="MAG 型（CF 頁）", fye=12, unit="US$B", scale=1, sheet="CF",
                  tag="Derived（個股模型基準情境 B；FY2024A–2025A 為財報，FY2026E 含 9M 實績）",
                  block=("A43", "情境 B"), header=(44, {"C": "FY2024A", "D": "FY2025A", "G": "FY2026E", "H": "FY2027E", "M": "FY2032E"}),
                  rows=dict(
                      ocf=dict(cells=[(50, "營運現金流")], cols=MAG7),
                      capai=dict(cells=[(51, "資本支出：AI")], cols=MAG7_FWD, note="FY2026E 欄只有 Q4（9M 未拆 AI／非 AI）→ 2026 無模型值"),
                      nai=dict(cells=[(52, "資本支出：非 AI")], cols=MAG7_FWD, note="FY2026E 欄只有 Q4 → 2026 無模型值"),
                      cap=dict(cells=[(53, "資本支出合計")], cols=MAG7),
                      bb=dict(cells=[(56, "普通股股利"), (57, "特別股股利"), (66, "回購")], cols=MAG7, note="FY2024A–2025A 回購欄空白 → 該年無模型值"),
                  ),
                  cash=dict(cell="E63", label="期初現金（現金＋債權類有價證券）", asof="2025-12-31（9M26 期初）")),
    "amzn": dict(ticker="AMZN", family="MAG 型（CF 頁）", fye=12, unit="US$B", scale=1, sheet="CF",
                 tag="Derived（個股模型基準情境 B；FY2024A–2025A 為財報，FY2026E 含 9M 實績）",
                 block=("A44", "情境 B（基準）"), header=(45, {"C": "FY2024A", "D": "FY2025A", "G": "FY2026E（9M 實績＋Q4 模型）", "H": "FY2027E", "M": "FY2032E"}),
                 rows=dict(
                     ocf=dict(cells=[(50, "營運現金流")], cols=MAG7),
                     capai=dict(cells=[(51, "資本支出：AI（Thesis）")], cols=MAG7),
                     nai=dict(cells=[(52, "資本支出：AWS 其他"), (53, "資本支出：零售")], cols=MAG7, note="AWS 其他＋零售"),
                     cap=dict(cells=[(54, "資本支出（現金，淨額）")], cols=MAG7),
                     bb=dict(cells=[(57, "股利"), (58, "回購")], cols=MAG7),
                 ),
                 cash=dict(cell="D64", label="期末現金與有價證券", asof="2025-12-31（FY2025A 期末）")),
    "meta": dict(ticker="META", family="MAG 型（CF 頁）", fye=12, unit="US$B", scale=1, sheet="CF",
                 tag="Derived（個股模型基準情境 B；FY2024A–2025A 為財報，FY2026E 含 9M 實績）",
                 block=("A43", "情境 B"), header=(44, {"C": "FY2024A", "D": "FY2025A", "G": "FY2026E", "H": "FY2027E", "M": "FY2032E"}),
                 rows=dict(
                     ocf=dict(cells=[(51, "營運現金流")], cols=MAG7),
                     capai=dict(cells=[(52, "資本支出：AI（Thesis）")], cols=MAG7_FWD, note="FY2026E 欄只有 Q4 → 2026 無模型值"),
                     nai=dict(cells=[(53, "資本支出：非 AI（Seg）")], cols=MAG7_FWD, note="FY2026E 欄只有 Q4 → 2026 無模型值"),
                     cap=dict(cells=[(54, "資本支出合計（含融資租賃本金）")], cols=MAG7),
                     bb=dict(cells=[(57, "普通股股利"), (58, "特別股股利"), (66, "回購")], cols=MAG7, note="FY2024A–2025A 回購欄空白 → 該年無模型值"),
                 ),
                 cash=dict(cell="E63", label="期初現金（現金＋有價證券）", asof="2025-12-31（9M26 期初）")),
    "orcl": dict(ticker="ORCL", family="MAG 型（CF 頁）", fye=5, unit="US$B", scale=1, sheet="CF",
                 tag="Derived（個股模型基準情境 B；FY2026A 為財報）",
                 block=("A49", "情境 B（基準"), header=(13, {"C": "FY2026A", "D": "FY2027", "I": "FY2032"}),
                 rows=dict(
                     ocf=dict(cells=[(55, "營運現金流")], cols={"C": 2026, "D": 2027, "E": 2028, "F": 2029, "G": 2030, "H": 2031, "I": 2032}),
                     capai=dict(cells=[(56, "資本支出：AI（Thesis）")], cols={"D": 2027, "E": 2028, "F": 2029, "G": 2030, "H": 2031, "I": 2032}),
                     nai=dict(cells=[(57, "資本支出：非 AI")], cols={"D": 2027, "E": 2028, "F": 2029, "G": 2030, "H": 2031, "I": 2032}, note="FY2026A 空白 → 無模型值"),
                     cap=dict(cells=[(58, "資本支出合計")], cols={"C": 2026, "D": 2027, "E": 2028, "F": 2029, "G": 2030, "H": 2031, "I": 2032}),
                     bb=dict(cells=[(60, "股利（普通股＋特別股）")], cols={"C": 2026, "D": 2027, "E": 2028, "F": 2029, "G": 2030, "H": 2031, "I": 2032}, note="只有股利；模型無回購列"),
                 ),
                 cash=dict(cell="D63", label="期初現金（FY2027＝Q1 末）", asof="2026-08-31（FY2027 Q1 末；模型無 2025-12-31）")),
    "xai": dict(ticker="SPCX", family="MAG 型（CF 頁）", fye=12, unit="US$B", scale=1, sheet="CF",
                tag="Derived（SpaceX 個股模型基準情境 B，合併 xAI；FY2026E 含 9M 實績）",
                block=("A45", "情境 B"), header=(46, {"C": "FY2025A", "G": "FY2026E", "H": "FY2027E", "M": "FY2032E"}),
                rows=dict(
                    ocf=dict(cells=[(52, "營運現金流")], cols={k: v for k, v in MAG7.items() if v >= 2026}),
                    capai=dict(cells=[(55, "資本支出：SI（現金）")], cols={k: v for k, v in MAG7.items() if v >= 2026}, note="SI（xAI 算力）＝AI"),
                    nai=dict(cells=[(53, "資本支出：Space"), (54, "資本支出：Connectivity")], cols={k: v for k, v in MAG7.items() if v >= 2026}, note="Space＋Connectivity（Starlink）"),
                    cap=dict(cells=[(56, "資本支出合計（現金）")], cols={k: v for k, v in MAG7.items() if v >= 2026}),
                ),
                cash=dict(cell="E67", label="期初現金（現金＋有價證券）", asof="2025-12-31（9M26 期初）"),
                missing=dict(bb="模型無股利、回購列")),
    "crwv": dict(ticker="CRWV", family="新雲型（各期收支頁）", fye=12, unit="US$B", scale=1, sheet="各期收支",
                 tag="Derived（個股模型；營業現金流由模型列組合）",
                 block=("A1", "支出與資金"), header=(4, {"C": "FY26", "G": "FY30"}),
                 rows=dict(
                     ocf=dict(cells=[(39, "營運缺口（不含融資、不含還本）"), (9, "① CapEx（用途用"), (18, "④ JV／策略投資出資"), (19, "電力（overlay）"), (20, "維護（overlay）")], cols=NEO,
                              note="模型無前瞻 CFO 列；營業現金流＝營運缺口＋CapEx＋JV 出資＋電力／維護 overlay（＝營運來源 − 租金 − 利息；FY26 含 1H 實際 CFO 3.663）"),
                     capai=dict(cells=[(9, "① CapEx（用途用")], cols=NEO, note="全公司為 AI 雲"),
                     cap=dict(cells=[(9, "① CapEx（用途用")], cols=NEO),
                 ),
                 cash=dict(cell="C74", label="期初現金（2025-12-31）", asof="2025-12-31"),
                 missing=dict(nai="模型未拆非 AI（全公司為 AI 雲）", bb="模型無股利、回購列（models_extract 標籤註明無股利與回購）")),
    "nbis": dict(ticker="NBIS", family="新雲型（各期收支頁）", fye=12, unit="US$B", scale=1, sheet="各期收支",
                 tag="Derived（個股模型；營業現金流由模型列組合）",
                 block=("A1", "支出與資金"), header=(4, {"C": "FY26", "G": "FY30"}),
                 rows=dict(
                     ocf=dict(cells=[(45, "營運缺口（不含融資、不含還本）"), (12, "① CapEx（用途用"), (21, "④ JV／策略投資出資"), (22, "電力（overlay）"), (23, "維護（overlay）")], cols=NEO,
                              note="模型無前瞻 CFO 列；營業現金流＝營運缺口＋CapEx＋JV 出資＋overlay（FY26 含 1H 實際 CFO 4.50）"),
                     capai=dict(cells=[(12, "① CapEx（用途用")], cols=NEO, note="全公司為 AI 雲"),
                     cap=dict(cells=[(12, "① CapEx（用途用")], cols=NEO),
                 ),
                 cash=dict(cell="C86", label="期初現金（2025-12-31）", asof="2025-12-31"),
                 missing=dict(nai="模型未拆非 AI（全公司為 AI 雲）", bb="模型無股利、回購列（第 21 列說明：NBIS 不發股息）")),
    "wyfi": dict(ticker="WYFI", family="新雲型（各期收支頁）", fye=12, unit="US$B", scale=0.001, sheet="各期收支",
                 tag="Derived（個股模型；營業現金流由模型列組合；原值 US$m ÷ 1000）",
                 block=("A1", "支出與資金"), header=(4, {"C": "FY26", "G": "FY30"}),
                 rows=dict(
                     ocf=dict(cells=[(52, "營運缺口（不含融資、不含還本）"), (13, "① CapEx（用途用"), (23, "④ JV／策略投資出資"), (24, "⑧ 股利"), (25, "電力（overlay）"), (26, "維護（overlay）")], cols=NEO,
                              note="模型無前瞻 CFO 列；營業現金流＝營運缺口＋CapEx＋JV 出資＋股利＋overlay（營運缺口已扣股利，加回；現金稅留在營業現金流內）；原值 US$m"),
                     capai=dict(cells=[(13, "① CapEx（用途用")], cols=NEO, note="雲端＋託管，全為 AI 資料中心"),
                     cap=dict(cells=[(13, "① CapEx（用途用")], cols=NEO),
                     bb=dict(cells=[(24, "⑧ 股利")], cols=NEO, note="模型股利列（0）；無回購列"),
                 ),
                 cash=dict(cell="C100", label="期初現金（2025-12-31）", asof="2025-12-31"),
                 missing=dict(nai="模型未拆非 AI")),
    "oai": dict(ticker="OpenAI", family="AI 實驗室型（Funding 頁）", fye=12, unit="US$B", scale=1, sheet="Funding",
                tag="Derived（OpenAI 收支模型；營收等輸入為 Interested-party、自有資本支出為模型假設）",
                block=("A1", "Funding"), header=(5, {"D": "2025", "E": "2026", "I": "2030"}),
                rows=dict(
                    ocf=dict(cells=[(13, "F06"), ("Cost!26", "自有資本支出（扣合作方出資）")], cols=LAB_FWD,
                             note="營業現金流＝自由現金流 F06（營收 − 算力現金成本 − 非算力成本）＋ Cost 頁自有資本支出（算力現金成本內含，屬投資）；2025 為校準年（模型註明只對照）→ 不取"),
                    capai=dict(cells=[("Cost!26", "自有資本支出（扣合作方出資）")], cols=LAB_FWD, note="模型自有資本支出（Inputs 第 180–185 列，模型標 Assumed）"),
                    cap=dict(cells=[("Cost!26", "自有資本支出（扣合作方出資）")], cols=LAB_FWD),
                ),
                cash=dict(cell="D30", label="年底現金", asof="2025-12-31"),
                missing=dict(nai="模型無非 AI 資本支出", bb="未上市；模型無股利、回購")),
    "ant": dict(ticker="Anthropic", family="AI 實驗室型（Funding 頁）", fye=12, unit="US$B", scale=1, sheet="Funding",
                tag="Derived（Anthropic 收支模型；營收等輸入為 Interested-party）",
                block=("A1", "Funding"), header=(5, {"D": "2025", "E": "2026", "I": "2030"}),
                rows=dict(
                    ocf=dict(cells=[(14, "F07"), (10, "F03")], cols=LAB_FWD,
                             note="營業現金流＝自由現金流 F07 ＋ 自建資本支出 F03（Fluidstack；算力現金成本內含，屬投資）；2025 欄模型註明只對照 → 不取"),
                    capai=dict(cells=[(10, "F03")], cols=LAB_FWD, note="自建資本支出（Fluidstack $50B 均攤 2026–2028）"),
                    cap=dict(cells=[(10, "F03")], cols=LAB_FWD),
                ),
                cash=dict(cell="D42", label="年底現金", asof="2025-12-31"),
                missing=dict(nai="模型無非 AI 資本支出", bb="未上市；模型無股利、回購")),
    "zhipu": dict(ticker="02513", family="AI 實驗室型（Funding 頁）", fye=12, unit="RMB 億", scale=1, sheet="Funding",
                  tag="Derived（智譜個股模型；2025 為財報；RMB 億）",
                  block=("A1", "Funding"), header=(5, {"D": "2025", "E": "2026", "I": "2030"}),
                  rows=dict(
                      ocf=dict(cells=[(16, "F09")], cols=LAB, note="自由現金流 F09（未扣非算力資本支出；自有算力資本支出基準 0，所以＝營業現金流）"),
                      nai=dict(cells=[(18, "F10")], cols=LAB, note="非算力資本支出"),
                      cap=dict(cells=[(18, "F10")], cols=LAB, note="自有算力資本支出基準 0（D11r），合計＝非算力資本支出"),
                  ),
                  cash=dict(cell="D39", label="期末現金", asof="2025-12-31"),
                  fx=dict(sheet="Inputs", cell="F27", lo_cell="G27", hi_cell="H27", label_cell="C27", label="匯率：USD/CNY", asof="2026-09-28"),
                  missing=dict(capai="自有算力資本支出基準 0（D11r）；智譜不是 Wattonomics 營運者（容量在中國合計）", bb="模型無股利、回購")),
}
ITEMS = ("ocf", "capai", "nai", "cap", "bb")
ITEM_NAME = {"ocf": "營業現金流", "capai": "AI 資本支出（模型）", "nai": "非 AI 資本支出", "cap": "資本支出合計（模型）", "bb": "買回＋股利"}


def commit_of(path):
    try:
        return subprocess.check_output(["git", "-C", ROOT, "log", "-1", "--format=%h", "--", path], text=True).strip()
    except Exception:
        return ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(GL, "builder", "seed", "gap_models_v0.1.7.json"))
    a = ap.parse_args()
    ex = json.load(open(EXTRACT, encoding="utf-8"))
    out = {"_doc": "Wattonomics v0.1.7 Gap 頁取數（tools/import_gap_models.py）；值為財年；null＝無模型值（不外推）。build.py 只在 In 頁 Y9 首建時讀本檔，之後 In 為權威",
           "players": {}}
    bad = []
    for p, c in CONF.items():
        rel = ex[c["ticker"]]["model"]
        path = os.path.join(ROOT, rel)
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb[c["sheet"]]
        grid = {}

        def v(sheet, ref):
            key = (sheet, ref)
            if key not in grid:
                grid[key] = wb[sheet][ref].value
            return grid[key]
        bc, btxt = c["block"]
        if btxt not in str(v(c["sheet"], bc) or ""):
            bad.append(f"{p}：區塊標題 {c['sheet']}!{bc} 不含「{btxt}」（實際：{v(c['sheet'], bc)}）")
        hr, hmap = c["header"]
        for col, want in hmap.items():
            got = v(c["sheet"], f"{col}{hr}")
            if str(got).strip() != want:
                bad.append(f"{p}：表頭 {c['sheet']}!{col}{hr} 應為「{want}」，實際「{got}」")
        rec = {"ticker": c["ticker"], "file": rel, "commit": commit_of(rel), "family": c["family"], "fye": c["fye"], "unit": c["unit"],
               "scale": c["scale"], "tag": c["tag"], "items": {}}
        merged = {}
        for item, spec in c["rows"].items():
            target = spec.get("merge_into", item)
            cells = []
            for rr, must in spec["cells"]:
                sh, row = (rr.split("!")[0], int(rr.split("!")[1])) if isinstance(rr, str) else (c["sheet"], rr)
                lab = None
                for lc in ("A", "B", "C"):
                    t = v(sh, f"{lc}{row}")
                    if isinstance(t, str) and must in t:
                        lab = t; break
                if lab is None:
                    bad.append(f"{p}.{item}：{sh} 第 {row} 列 A–C 欄找不到標籤「{must}」")
                cells.append((sh, row, lab))
            vals, refs = {}, {}
            for col, fy in spec["cols"].items():
                parts = [v(sh, f"{col}{row}") for sh, row, _ in cells]
                ref = "＋".join(f"{sh}!{col}{row}" for sh, row, _ in cells)
                if all(isinstance(x, (int, float)) for x in parts):
                    vals[str(fy)] = round(sum(parts) * c["scale"], 6)
                else:
                    vals[str(fy)] = None
                refs[str(fy)] = ref
            m = merged.setdefault(target, {"values": {}, "cells": {}, "labels": [], "notes": []})
            for fy in vals:
                if fy not in m["values"] or m["values"][fy] is None:
                    m["values"][fy] = vals[fy]; m["cells"][fy] = refs[fy]
            m["labels"] += [f"{sh}!A–C{row}「{lab}」" for sh, row, lab in cells if f"{sh}!A–C{row}「{lab}」" not in m["labels"]]
            if spec.get("note"): m["notes"].append(spec["note"])
        for item in ITEMS:
            if item in merged:
                m = merged[item]
                rec["items"][item] = {"values": dict(sorted(m["values"].items())), "cells": dict(sorted(m["cells"].items())),
                                      "labels": m["labels"], "note": "；".join(m["notes"])}
            else:
                rec["items"][item] = {"values": {}, "cells": {}, "labels": [], "note": "無模型值：" + c.get("missing", {}).get(item, "模型無此列")}
        cv = v(c["sheet"], c["cash"]["cell"])
        rec["cash"] = {"value": round(cv * c["scale"], 6) if isinstance(cv, (int, float)) else None, "cell": f"{c['sheet']}!{c['cash']['cell']}",
                       "label": c["cash"]["label"], "asof": c["cash"]["asof"]}
        # 現金標籤核對（同列 A–C 欄）
        crow = int("".join(ch for ch in c["cash"]["cell"] if ch.isdigit()))
        if not any(c["cash"]["label"] in str(v(c["sheet"], f"{lc}{crow}") or "") for lc in "ABC"):
            bad.append(f"{p}：現金列 {c['sheet']} 第 {crow} 列找不到標籤「{c['cash']['label']}」")
        if "fx" in c:
            fx = c["fx"]
            if fx["label"] not in str(v(fx["sheet"], fx["label_cell"])):
                bad.append(f"{p}：匯率標籤 {fx['sheet']}!{fx['label_cell']} 不符")
            rec["fx"] = {"value": v(fx["sheet"], fx["cell"]), "lo": v(fx["sheet"], fx["lo_cell"]), "hi": v(fx["sheet"], fx["hi_cell"]), "cell": f"{fx['sheet']}!{fx['cell']}", "label": str(v(fx["sheet"], fx["label_cell"])), "asof": fx["asof"]}
        out["players"][p] = rec
    if bad:
        print("停止（列位核對不符，不寫檔）：\n  " + "\n  ".join(bad)); sys.exit(1)
    json.dump(out, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"寫出 {a.out}")
    for p, rec in out["players"].items():
        miss = [f"{ITEM_NAME[i]}" for i in ITEMS if not any(x is not None for x in rec["items"][i]["values"].values())]
        print(f"  {p:6s} {rec['file']}（{rec['commit']}）OCF {rec['items']['ocf']['values']}  現金 {rec['cash']['value']}  整列無模型值：{miss}")


if __name__ == "__main__":
    main()
