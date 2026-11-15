# Wattonomics v0.1.7 資金缺口（Gap 頁；11 家代表玩家，不含中國，Andy 2026-10-31 12:58）取數：讀各個股模型的基準情境年度值，寫 builder/seed/gap_models_v0.1.7.json。
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
# v0.1.8：models_extract.json 尚未更新的個股（Wattonomics 不改 markets/ 檔案），在此逐家覆寫為 main 上的現行版；
# 覆寫只換檔案路徑，列位核對規則不變（工作單 20261102 E7：Anthropic v0.1.7 已合併 #289，dist 只剩 v0_1_7）
MODEL_OVERRIDE = {"Anthropic": "anthropic/dist/20261104_Anthropic收支模型_v0_1_9.xlsx"}   # v0.2.1 W4：Anthropic v0.1.8；2026-11-05 改指 v0.1.9（#327 移除 v0_1_8；F07／F03 逐位相同，種子與成品不重建）

# 欄位 → 財年。MAG 型：C=FY2024A、D=FY2025A、E=9M26A、F=Q4 26E、G=FY2026E、H..M=FY2027E..FY2032E（Amazon、Alphabet、Meta）
MAG7 = {"C": 2024, "D": 2025, "G": 2026, "H": 2027, "I": 2028, "J": 2029, "K": 2030, "L": 2031, "M": 2032}
MAG7_FWD = {k: v for k, v in MAG7.items() if v >= 2027}
NEO = {"C": 2026, "D": 2027, "E": 2028, "F": 2029, "G": 2030}
NEOTH = {"E": 2026, "F": 2027, "G": 2028, "H": 2029, "I": 2030}   # v0.3.4：新雲 Thesis 頁（C＝9M 實績、D＝Q4 模型、E＝FY2026 … I＝FY2030；FY2031–32 不取，與舊設定同範圍）
NEOCF = {"D": 2026, "E": 2027, "F": 2028, "G": 2029, "H": 2030}   # v0.3.4：新雲 CF 頁（C＝FY2025A、D＝FY2026 … H＝FY2030）
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
    # v0.3.4（2026-11-15）：CoreWeave v5.0、Nebius v0.6、WhiteFiber v0.4 起個股模型改為 Thesis／CF／Fund 頁（已無「各期收支」頁）。
    # 資本支出讀 Thesis 情境 B「資本支出合計（公司口徑；Q4 起模型）」（FY2026＝9M 實績＋Q4 模型）；營業現金流與期初現金讀 CF 情境 B。
    # 兩頁各自核對情境 B 區塊標題與年度表頭（extra）；舊設定（各期收支頁組合式營業現金流）見 git 歷史。
    "crwv": dict(ticker="CRWV", family="新雲型（Thesis／CF 頁；v5.0 起）", fye=12, unit="US$B", scale=1, sheet="CF",
                 tag="Derived（個股模型基準情境 B；FY2026＝9M 實績＋Q4 模型）",
                 block=("A40", "情境 B（基準"), header=(4, {"D": "FY2026", "H": "FY2030"}),
                 extra=[("Thesis", "A39", "情境 B（基準", "in"), ("Thesis", "E3", "FY2026", "eq"), ("Thesis", "I3", "FY2030", "eq")],
                 rows=dict(
                     ocf=dict(cells=[(47, "營運現金流")], cols=NEOCF, note="CF 情境 B 營運現金流（FY2026＝9M 實績 5.638＋Q4 模型；含客戶預付流入、扣預付認列）"),
                     capai=dict(cells=[("Thesis!66", "資本支出合計（公司口徑；Q4 起模型）")], cols=NEOTH, note="全公司為 AI 雲；公司口徑（9M＝1H 10-Q 16.139＋Q3 法說 14.511，應計；CF 頁現金資本支出 FY2026 30.64 只作對照）"),
                     cap=dict(cells=[("Thesis!66", "資本支出合計（公司口徑；Q4 起模型）")], cols=NEOTH),
                 ),
                 cash=dict(carry=True, ref_cell="D51", ref_label="期初現金（FY2026＝2026-09-30）"),
                 missing=dict(nai="模型未拆非 AI（全公司為 AI 雲）", bb="模型無股利、回購列（CF 頁說明：無股利、無回購）")),
    "nbis": dict(ticker="NBIS", family="新雲型（Thesis／CF 頁；v0.6 起）", fye=12, unit="US$B", scale=1, sheet="CF",
                 tag="Derived（個股模型基準情境 B；FY2026＝9M 實績＋Q4 模型）",
                 block=("A40", "情境 B（基準"), header=(4, {"D": "FY2026", "H": "FY2030"}),
                 extra=[("Thesis", "A39", "情境 B（基準", "in"), ("Thesis", "E3", "FY2026", "eq"), ("Thesis", "I3", "FY2030", "eq")],
                 rows=dict(
                     ocf=dict(cells=[(47, "營運現金流")], cols=NEOCF, note="CF 情境 B 營運現金流（FY2026＝9M 實績 4.231＋Q4 模型；含客戶預付流入、扣預付認列）"),
                     capai=dict(cells=[("Thesis!66", "資本支出合計（公司口徑；Q4 起模型）")], cols=NEOTH, note="全公司為 AI 雲；9M＝現金購置實績 15.431（6-K；公司未另定應計口徑）；公司 2026 指引約 27 不入（模型原則 59）"),
                     cap=dict(cells=[("Thesis!66", "資本支出合計（公司口徑；Q4 起模型）")], cols=NEOTH),
                 ),
                 cash=dict(carry=True, ref_cell="D51", ref_label="期初現金（FY2026＝2026-09-30）"),
                 missing=dict(nai="模型未拆非 AI（全公司為 AI 雲）", bb="模型無股利、回購列（CF 頁說明：無股利、無回購）")),
    "wyfi": dict(ticker="WYFI", family="新雲型（Thesis／CF 頁；v0.4 起）", fye=12, unit="US$B", scale=0.001, sheet="CF",
                 tag="Derived（個股模型基準情境 B；FY2026＝9M 實績＋Q4 模型；原值 US$m ÷ 1000）",
                 block=("A40", "情境 B（基準"), header=(4, {"D": "FY2026", "H": "FY2030"}),
                 extra=[("Thesis", "A41", "情境 B（基準", "in"), ("Thesis", "E3", "FY2026", "eq"), ("Thesis", "I3", "FY2030", "eq")],
                 rows=dict(
                     ocf=dict(cells=[(47, "營運現金流")], cols=NEOCF, note="CF 情境 B 營運現金流（FY2026＝9M 實績 211.3＋Q4 模型；含客戶預付流入、扣預付認列）；原值 US$m"),
                     capai=dict(cells=[("Thesis!70", "資本支出合計（公司口徑；Q4 起模型）")], cols=NEOTH, note="雲端＋託管，全為 AI 資料中心；9M＝現金購置與設備訂金實績 549.8（10-Q）"),
                     cap=dict(cells=[("Thesis!70", "資本支出合計（公司口徑；Q4 起模型）")], cols=NEOTH),
                 ),
                 cash=dict(carry=True, ref_cell="D51", ref_label="期初現金（FY2026＝2026-09-30）"),
                 missing=dict(nai="模型未拆非 AI", bb="模型無股利、回購列（CF 頁說明：無股利、無回購）")),
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
    ap.add_argument("--out", default=os.path.join(GL, "builder", "seed", "gap_models_v0.1.8.json"))
    ap.add_argument("--only", help="v0.2.1：只重取指定玩家（逗號分隔鍵，例如 ant），其餘沿用 --from 種子")
    ap.add_argument("--from", dest="src", default=os.path.join(GL, "builder", "seed", "gap_models_v0.1.8.json"))
    a = ap.parse_args()
    ex = json.load(open(EXTRACT, encoding="utf-8"))
    only = set(a.only.split(",")) if a.only else None
    out = {"_doc": "Wattonomics v0.1.8 Gap 頁取數（tools/import_gap_models.py）；值為財年；null＝無模型值（不外推）。build.py 只在 In 頁 Y9 首建時讀本檔，之後 In 為權威",
           "players": {}}
    bad = []
    if only:
        out = json.load(open(a.src, encoding="utf-8"))
        out["_doc"] += f"；{'、'.join(sorted(only))} 以 --only 重取（其餘沿用 {os.path.basename(a.src)}）"
    for p, c in CONF.items():
        if only and p not in only: continue
        rel = MODEL_OVERRIDE.get(c["ticker"], ex[c["ticker"]]["model"])
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
        for sh_, cell_, want_, mode_ in c.get("extra", []):   # v0.3.4：其他頁的區塊標題／表頭核對
            got_ = v(sh_, cell_)
            ok_ = (want_ in str(got_ or "")) if mode_ == "in" else (str(got_).strip() == want_)
            if not ok_:
                bad.append(f"{p}：{sh_}!{cell_} 應{'含' if mode_ == 'in' else '為'}「{want_}」，實際「{got_}」")
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
        if c["cash"].get("carry"):
            # v0.3.4：新版模型無 2025-12-31 期初現金列（CF 期初現金為 2026-09-30，與其他 8 家的 2025-12-31 口徑不同，且 Gap 頁 2026 需求為全年）
            # → 寫「模型無此列」並沿用前一種子的 2025-12-31 值（出處為舊版模型儲存格）；2026-09-30 值只寫入說明作對照，不採用
            prev = json.load(open(a.src, encoding="utf-8"))["players"][p]
            rcell, rlab = c["cash"]["ref_cell"], c["cash"]["ref_label"]
            rrow = int("".join(ch for ch in rcell if ch.isdigit()))
            if not any(rlab in str(v(c["sheet"], f"{lc}{rrow}") or "") for lc in "ABC"):
                bad.append(f"{p}：對照現金列 {c['sheet']} 第 {rrow} 列找不到標籤「{rlab}」")
            rv = v(c["sheet"], rcell)
            rv = round(rv * c["scale"], 6) if isinstance(rv, (int, float)) else None
            rec["cash"] = {"value": prev["cash"]["value"],
                           "cell": f"模型無此列（新版無 2025-12-31 期初現金）；沿用 {prev['file']} {prev['cash']['cell']}",
                           "label": prev["cash"]["label"],
                           "asof": f"{prev['cash']['asof']}；對照：新版 {c['sheet']}!{rcell}「{rlab}」{rv}（不採用：asof 與他家不同）"}
        else:
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
        if only and p not in only: continue
        miss = [f"{ITEM_NAME[i]}" for i in ITEMS if not any(x is not None for x in rec["items"][i]["values"].values())]
        print(f"  {p:6s} {rec['file']}（{rec['commit']}）OCF {rec['items']['ocf']['values']}  現金 {rec['cash']['value']}  整列無模型值：{miss}")


if __name__ == "__main__":
    main()
