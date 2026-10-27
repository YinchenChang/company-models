# Siliconomics（AI 半導體產業基準庫）建檔程式
# 用法：
#   首版（由種子建）：python3 siliconomics/builder/build.py --out <raw.xlsx>
#   之後各版（Excel 優先）：python3 siliconomics/builder/build.py --base siliconomics/model/<現行>.xlsx --out <raw.xlsx>
#     --base 時，In 頁的藍字輸入（低／基準／高、標記、來源）一律讀回現行 Excel；種子只用於 Excel 裡還沒有的鍵。
# 產出後以 LibreOffice headless 重算存檔（見 siliconomics/README.md），重算錯誤須為 0。
# 原則（比照 Tokenomics CLAUDE.md 第 1 節）：Excel 是唯一事實來源；本程式只產生公式頁與格式，不得在程式內新增或省略計算機制。
import argparse, json, os, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
SIL = os.path.dirname(HERE)
ROOT = os.path.dirname(SIL)
sys.path.insert(0, os.path.join(ROOT, "markets", "memory"))

VERSION = "v0.1"   # 由 --version 覆寫
DATE = "20261027"  # 由 --date 覆寫
TK_SNAPSHOT = os.path.join(SIL, "data", "tokenomics_snapshot_v21.7.json")
TK_OLD = os.path.join(SIL, "data", "tokenomics_snapshot_v21.7.json")   # 下游（chip_bridge）實際讀的 TK 快照；v0.4 起與本庫同一份
SEED = os.path.join(HERE, "seed", "seed_v0.1.json")
CONFLICTS = os.path.join(HERE, "seed", "conflicts_v0.1.json")
DECISIONS = os.path.join(HERE, "seed", "decisions_v0.1.json")
import glob


def _vkey(f):
    """依版本號排序 seed 檔（v0.4 在 v0.4.1 之前；字串排序會顛倒）"""
    import re as _re
    m = _re.search(r"_v([0-9.]+)\.json$", f)
    return tuple(int(x) for x in m.group(1).split(".")) if m else (0,)
REF_SKH = os.path.join(ROOT, "skhynix", "dist")

BLUE = Font(color="0000FF"); GREEN = Font(color="008000"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7"); KEY = PatternFill("solid", fgColor="FFF2CC")
PEND = PatternFill("solid", fgColor="FFFF00"); WRAP = Alignment(wrap_text=True, vertical="top")
YEARS = [2025, 2026, 2027, 2028, 2029]


def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR


def name(wb, nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


def load_base_inputs(path):
    """讀回現行 Excel 的 In 頁（Excel 優先）：{鍵: (低, 基準, 高, 標記, 來源)}"""
    if not path:
        return {}
    wb = openpyxl.load_workbook(path, data_only=False)
    ws = wb["In"]
    out = {}
    for r in range(5, ws.max_row + 1):
        k = ws.cell(r, 1).value
        if not k or ws.cell(r, 3).value is None and ws.cell(r, 5).value is None:
            continue
        out[k] = tuple(ws.cell(r, c).value for c in (4, 5, 6, 7, 8, 10))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--base")
    ap.add_argument("--out", required=True)
    ap.add_argument("--version", help="新版版本號；省略且給 --base 時沿用 base 檔名的版本（CI 重建比對用）")
    ap.add_argument("--date", help="省略且給 --base 時沿用 base 檔名的日期")
    a = ap.parse_args()
    if a.base:
        fn = os.path.basename(a.base)
        if a.version is None and "_v" in fn:
            a.version = "v" + fn.rsplit("_v", 1)[1].rsplit(".xlsx", 1)[0]
        if a.date is None and fn[:8].isdigit():
            a.date = fn[:8]
    if a.base and (a.version is None or a.date is None):
        sys.exit("--base 的檔名不是 YYYYMMDD_Siliconomics_vX.Y.xlsx 時，必須明給 --version 與 --date（避免 README／Interface 版本標示錯誤）")
    a.version = a.version or VERSION
    a.date = a.date or DATE
    globals()["VERSION"], globals()["DATE"] = a.version, a.date
    seed = json.load(open(SEED, encoding="utf-8"))
    base = load_base_inputs(a.base)
    tk = json.load(open(TK_SNAPSHOT, encoding="utf-8"))
    tk_old = json.load(open(TK_OLD, encoding="utf-8")) if os.path.exists(TK_OLD) else None
    conflicts = json.load(open(CONFLICTS, encoding="utf-8")) if os.path.exists(CONFLICTS) else []
    decisions = []
    for f in sorted(glob.glob(os.path.join(HERE, "seed", "decisions_v*.json")), key=_vkey):
        decisions += json.load(open(f, encoding="utf-8"))
    # 新增的輸入列（各版 rows_added_v*.json）：接在指定節的最後一列之後，成為新的一節；base Excel 有值時以 Excel 為準
    seed_rows = list(seed["rows"])
    for f in sorted(glob.glob(os.path.join(HERE, "seed", "rows_added_v*.json")), key=_vkey):
        blocks = json.load(open(f, encoding="utf-8"))
        for ad in (blocks if isinstance(blocks, list) else [blocks]):
            idx = [i for i, rw in enumerate(seed_rows) if "sec" in rw]
            start = next(i for i in idx if seed_rows[i]["sec"].startswith(ad["after_section_end"]))
            nxt = next((i for i in idx if i > start), len(seed_rows))
            seed_rows[nxt:nxt] = [{"sec": ad["section"]}] + ad["rows"]

    wb = openpyxl.Workbook()
    RD = wb.active; RD.title = "README"
    SC = wb.create_sheet("Scen")
    IN = wb.create_sheet("In")

    # ---------------- In：藍字輸入（Excel 擁有） ----------------
    IN["A1"] = "In — 產業輸入（藍字＝輸入，由本活頁簿擁有；要改值直接改本頁，builder 重建時讀回）。每列具名範圍＝鍵（指向 E 欄），另有 _lo／_hi（D、F 欄）"
    IN["A2"] = "標記：Verified／Interested-party（寫明是誰）／Analogy／Assumed（給低／高區間）／Derived（寫明換算）。v0.1 內容由 markets/memory/market.py 逐字搬入（數值、標記、來源文字不變）"
    hdr(IN, 4, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源／說明", "Interface 名稱", "第二來源／說明"])
    INROW = {}
    r = 5
    for row in seed_rows:
        if "sec" in row:
            IN.cell(r, 1, row["sec"]).font = BOLD; r += 1; continue
        k = row["id"]
        lo, v, hi, tag, src, src2 = base.get(k, (row["lo"], row["base"], row["hi"], row["tag"], row["src"], row.get("src2", "")))
        IN.cell(r, 1, k); IN.cell(r, 2, row["label"]); IN.cell(r, 3, row["unit"])
        for j, x in ((4, lo), (5, v), (6, hi)):
            if x is not None:
                c = IN.cell(r, j, x); c.font = BLUE
                if row["unit"] == "%": c.number_format = "0.0%"
        IN.cell(r, 7, tag); IN.cell(r, 8, src); IN.cell(r, 9, "SIF_" + k)
        if src2: IN.cell(r, 10, src2)
        name(wb, k, f"In!$E${r}")
        if lo is not None: name(wb, k + "_lo", f"In!$D${r}")
        if hi is not None: name(wb, k + "_hi", f"In!$F${r}")
        INROW[k] = (r, row["label"], row["unit"], lo is not None, hi is not None, tag)
        r += 1
    r += 1
    IN.cell(r, 1, "M2 分組：加速器平台（Units 頁各列與 market.py UNITS 的名稱、標記、來源）").font = BOLD; r += 1
    hdr(IN, r, ["分組鍵", "平台名稱", "", "", "", "", "標記", "來源／說明"]); r += 1
    GRP_ROW0 = r
    for g in seed["unit_groups"]:
        IN.cell(r, 1, "grp_" + g["key"]); IN.cell(r, 2, g["name"]); IN.cell(r, 7, g["tag"]); IN.cell(r, 8, g["src"])
        r += 1
    IN_LAST = r - 1
    for col, w in zip("ABCDEFGHIJ", (14, 52, 8, 9, 10, 9, 18, 70, 18, 50)):
        IN.column_dimensions[col].width = w

    # ---------------- Path、Mkt：記憶體價格路徑與 HBM 市場（公式與 markets/memory/market.py write() 同一份） ----------------
    import market as MKT
    ref = MKT.write(wb, hdr, Font)
    MK = ref["mk"]
    wb["Mkt"]["A2"] = "公式由 markets/memory/market.py write() 產生（與各記憶體公司模型的 Mkt 頁同一份）；輸入讀本活頁簿 In 頁"

    # ---------------- Scen：情境定義 ----------------
    SC["A1"] = "Scen — 行業情境定義（A＝順利、B＝平平、C＝不如預期）。各公司模型的 A／B／C 一律讀同字母行業情境（Andy 2026-10-25 13:13）"
    hdr(SC, 3, ["情境", "名稱", "定義（顆數與部署）", "記憶體價格路徑", "2028 顆數 ÷ 基準", "2029 顆數 ÷ 基準", "2030–2032"])
    NM = {"A": "順利", "B": "平平", "C": "不如預期"}
    DEF = {"A": "基準顆數（市場層 2026–2028 預測）；部署順利、新技術導入快",
           "B": "AI 資本支出消化：2028 顆數約為基準 0.8、2029 約 0.85",
           "C": "明顯下修：2028、2029 顆數約為基準 0.6"}
    for i, s in enumerate(seed["scen"]):
        rr = 4 + i
        SC.cell(rr, 1, s); SC.cell(rr, 2, NM[s]); SC.cell(rr, 3, DEF[s]); SC.cell(rr, 4, seed["scen_desc"][s])
        for j, y in ((5, 28), (6, 29)):
            c = SC.cell(rr, j, f"=um{y}_{s}"); c.font = GREEN; c.number_format = "0.00"
        c = SC.cell(rr, 7, "未定：待算力收支 R4 產業部署路徑（預計 10-31 後）；各公司模型目前以收斂法自推"); c.fill = PEND
    SC["A9"] = "說明"
    SC["A9"].font = BOLD
    SC["A10"] = ("情境係數只定義 2028–29（Assumed，無外部預測可校準；B、C 區間 ±0.1，見 In 頁 M2b）。"
                 "R4 上線後，2026–2030 的部署 GW 三情境改讀算力收支（權威在算力收支，第 13 條），本頁的係數改為由部署路徑推導。")
    for col, w in zip("ABCDEFG", (6, 10, 46, 60, 14, 14, 50)):
        SC.column_dimensions[col].width = w

    # ---------------- Units：平台別顆數（2026–2029；2029＝2028 ×（1＋平台別成長率）） ----------------
    UN = None; UROW = {}
    if "g29_bw" in INROW:
        UN = wb.create_sheet("Units", index=wb.sheetnames.index("Path"))
        UN["A1"] = "Units — 加速器平台別顆數（百萬顆；情境 A 基準）。2026–2028＝In 頁 M2；2029＝2028 ×（1＋In 頁 M5 平台別成長率）。情境 B、C 再乘 Scen 頁係數（2028–29）"
        hdr(UN, 3, ["平台", "鍵", "2026", "2027", "2028", "2029", "2029 成長"])
        r = 4
        for g in seed["unit_groups"]:
            k = g["key"]
            UN.cell(r, 1, g["name"]); UN.cell(r, 2, k)
            for j, y in enumerate((26, 27, 28)):
                UN.cell(r, 3 + j, f"=u_{k}_{y}").font = GREEN
            UN.cell(r, 6, f"=E{r}*(1+g29_{k})"); UN.cell(r, 7, f"=g29_{k}").number_format = "0%"
            for j in range(3, 7): UN.cell(r, j).number_format = "0.00"
            UROW[k] = r; r += 1
        r0, r1 = 4, r - 1
        UN.cell(r, 1, "合計（基準＝情境 A，未乘係數）").font = BOLD
        for j in range(3, 7):
            c = UN.cell(r, j, f"=SUM({'CDEF'[j-3]}{r0}:{'CDEF'[j-3]}{r1})"); c.number_format = "0.00"; c.fill = KEY
        UN.cell(r, 7, f"=F{r}/E{r}-1").number_format = "0.0%"; UROW["tot"] = r; r += 1
        for s_ in seed["scen"]:
            UN.cell(r, 1, f"合計（情境 {s_}）")
            UN.cell(r, 3, f"=C{UROW['tot']}"); UN.cell(r, 4, f"=D{UROW['tot']}")
            UN.cell(r, 5, f"=E{UROW['tot']}*um28_{s_}"); UN.cell(r, 6, f"=F{UROW['tot']}*um29_{s_}")
            for j in range(3, 7): UN.cell(r, j).number_format = "0.00"
            UROW["tot" + s_] = r; r += 1
        UN.column_dimensions["A"].width = 40
        for col in "BCDEFG": UN.column_dimensions[col].width = 10

    # ---------------- ShareTbl：ASIC 設計服務份額表（同一平台各設計夥伴份額加總 ≤ 100%） ----------------
    ST = None; STROW = {}
    if "sh_tpu_26" in INROW:
        ST = wb.create_sheet("ShareTbl", index=wb.sheetnames.index("Path"))
        ST["A1"] = ("ShareTbl — ASIC 設計服務顆數份額表（基準值；In 頁 M7）。同一平台各設計夥伴份額加總須 ≤ 100%，餘額＝客戶自行設計或未建模的夥伴。"
                    "Trainium 的世芯份額隨情境（公司特有：A 最有利、C 最不利），三情境各列一次")
        hdr(ST, 3, ["平台", "夥伴", "鍵（2026／2027／2028）", "2026", "2027", "2028", "說明"])
        r = 4
        TBL = [("Google TPU", [("Broadcom", "sh_tpu_{y}"), ("Marvell", "msh_tpu_{y}"), ("聯發科", "mtk_tpu_{y}")], "tpu"),
               ("其他 ASIC（MTIA、Maia、OpenAI 等）", [("Broadcom", "sh_oth_{y}"), ("Marvell", "msh_oth_{y}"), ("聯發科", "mtk_oth_{y}")], "oth")]
        for s_ in seed["scen"]:
            TBL.append((f"AWS Trainium（情境 {s_}）", [("Marvell", "msh_trn_{y}"), ("世芯", "al_{y}x" + s_)], "trn" + s_))
        def kk(pat, y, s_=None):
            if pat.startswith("al_"):
                return "al_trn_26" if y == 26 else f"al{y}_{pat[-1]}"
            return pat.format(y=y)
        for plat, parts, code in TBL:
            r_first = r
            for who, pat in parts:
                ST.cell(r, 1, plat); ST.cell(r, 2, who); ST.cell(r, 3, "／".join(kk(pat, y) for y in (26, 27, 28)))
                for j, y in enumerate((26, 27, 28)):
                    c = ST.cell(r, 4 + j, "=" + kk(pat, y)); c.font = GREEN; c.number_format = "0%"
                r += 1
            ST.cell(r, 1, plat); ST.cell(r, 2, "合計").font = BOLD
            for j in range(3):
                col = "DEF"[j]
                c = ST.cell(r, 4 + j, f"=SUM({col}{r_first}:{col}{r-1})"); c.number_format = "0%"; c.fill = KEY
            STROW["sum_" + code] = r; r += 1
            ST.cell(r, 1, plat); ST.cell(r, 2, "餘額（客戶自行設計或未建模）")
            for j in range(3):
                col = "DEF"[j]
                ST.cell(r, 4 + j, f"=1-{col}{r-1}").number_format = "0%"
            r += 2
        if "cposh_lite_27" in INROW:
            ST.cell(r, 1, "CPO 光學金額份額（欄 D–F＝2027／2028／2029；Fabrinet 為代工，與模組廠重疊，不計入合計）").font = BOLD; r += 1
            CPO_CO = [("Lumentum", "lite"), ("Coherent", "cohr"), ("中際旭創", "innolight"), ("新易盛", "eoptolink")]
            r_first = r
            for who, co in CPO_CO:
                ST.cell(r, 1, "CPO 光學"); ST.cell(r, 2, who); ST.cell(r, 3, f"cposh_{co}_27／28／29")
                for j, y in enumerate((27, 28, 29)):
                    c = ST.cell(r, 4 + j, f"=cposh_{co}_{y}"); c.font = GREEN; c.number_format = "0%"
                r += 1
            ST.cell(r, 1, "CPO 光學"); ST.cell(r, 2, "合計（模組與雷射廠）").font = BOLD
            for j in range(3):
                col = "DEF"[j]
                c = ST.cell(r, 4 + j, f"=SUM({col}{r_first}:{col}{r-1})"); c.number_format = "0%"; c.fill = KEY
            STROW["sum_cpo"] = r; r += 1
            ST.cell(r, 1, "CPO 光學"); ST.cell(r, 2, "餘額（NVIDIA、Broadcom 交換器端自製光引擎與其他，未建模）")
            for j in range(3):
                col = "DEF"[j]
                ST.cell(r, 4 + j, f"=1-{col}{r-1}").number_format = "0%"
            r += 1
            ST.cell(r, 1, "CPO 光學"); ST.cell(r, 2, "Fabrinet（代工；參考，不計入）"); ST.cell(r, 3, "cposh_fabrinet_27／28／29")
            for j, y in enumerate((27, 28, 29)):
                c = ST.cell(r, 4 + j, f"=cposh_fabrinet_{y}"); c.font = GREEN; c.number_format = "0%"
            r += 2
        if "abf_ibsh26" in INROW:
            ST.cell(r, 1, "加速器 ABF 載板份額（欄 D–G＝2026／2027／2028／2029；NVIDIA 為顆數份額，非 NVIDIA 為金額份額）").font = BOLD; r += 1
            def ibrel(y):
                return "1" if y == 26 else f"abf_ibrel_{y}"
            for side, plat, ibf in (("nv", "NVIDIA 加速器", lambda y: f"=abf_ibsh26*{ibrel(y)}"),
                                    ("ot", "非 NVIDIA 加速器", lambda y: f"=abf_wot_ib*{ibrel(y)}/abf_vot")):
                r_first = r
                ST.cell(r, 1, plat); ST.cell(r, 2, "Ibiden")
                ST.cell(r, 3, "abf_ibsh26 × abf_ibrel" if side == "nv" else "abf_wot_ib × abf_ibrel ÷ abf_vot")
                for j, y in enumerate((26, 27, 28, 29)):
                    c = ST.cell(r, 4 + j, ibf(y)); c.number_format = "0.0%"
                r += 1
                for who, co in (("欣興", "un"), ("南電", "np"), ("景碩", "ks")):
                    ST.cell(r, 1, plat); ST.cell(r, 2, who); ST.cell(r, 3, f"abf_{co}_{side}_26／27／28／29")
                    for j, y in enumerate((26, 27, 28, 29)):
                        c = ST.cell(r, 4 + j, f"=abf_{co}_{side}_{y}"); c.font = GREEN; c.number_format = "0.0%"
                    r += 1
                ST.cell(r, 1, plat); ST.cell(r, 2, "合計").font = BOLD
                for j in range(4):
                    col = "DEFG"[j]
                    c = ST.cell(r, 4 + j, f"=SUM({col}{r_first}:{col}{r-1})"); c.number_format = "0.0%"; c.fill = KEY
                STROW["sum_abf" + side] = r; r += 1
                ST.cell(r, 1, plat); ST.cell(r, 2, "餘額（三星電機、AT&S、Shinko 等，未建模）")
                for j in range(4):
                    col = "DEFG"[j]
                    ST.cell(r, 4 + j, f"=1-{col}{r-1}").number_format = "0.0%"
                r += 2
            ST.column_dimensions["G"].width = 9
        if "pwr_dsh_26" in INROW:
            ST.cell(r, 1, "AI 電源有效份額（欄 D–G＝2026／2027／2028／2029；分母＝Rack 頁 AI 電源市場；2026＝公司模型校準值，之後 × 份額相對值；光寶相對值隨情境：A 高、B 基準、C 低）").font = BOLD; r += 1
            for s_ in seed["scen"]:
                suf = {"A": "_hi", "B": "", "C": "_lo"}[s_]
                r_first = r
                ST.cell(r, 1, f"AI 電源（情境 {s_}）"); ST.cell(r, 2, "台達電"); ST.cell(r, 3, "pwr_dsh_26 × pwr_drel_27／28／29（不隨情境）")
                for j, y in enumerate((26, 27, 28, 29)):
                    c = ST.cell(r, 4 + j, "=pwr_dsh_26" + ("" if y == 26 else f"*pwr_drel_{y}")); c.number_format = "0.0%"
                r += 1
                ST.cell(r, 1, f"AI 電源（情境 {s_}）"); ST.cell(r, 2, "光寶"); ST.cell(r, 3, f"pwr_lsh_26 × pwr_lrel_27／28／29{suf}")
                for j, y in enumerate((26, 27, 28, 29)):
                    c = ST.cell(r, 4 + j, "=pwr_lsh_26" + ("" if y == 26 else f"*pwr_lrel_{y}{suf}")); c.number_format = "0.0%"
                r += 1
                ST.cell(r, 1, f"AI 電源（情境 {s_}）"); ST.cell(r, 2, "合計").font = BOLD
                for j in range(4):
                    col = "DEFG"[j]
                    c = ST.cell(r, 4 + j, f"=SUM({col}{r_first}:{col}{r-1})"); c.number_format = "0.0%"; c.fill = KEY
                STROW["sum_pwr" + s_] = r; r += 1
                ST.cell(r, 1, f"AI 電源（情境 {s_}）"); ST.cell(r, 2, "餘額（Flex、群電、康舒、Vertiv、麥格米特等，未建模）")
                for j in range(4):
                    col = "DEFG"[j]
                    ST.cell(r, 4 + j, f"=1-{col}{r-1}").number_format = "0.0%"
                r += 2
            ST.cell(r, 1, "對照：凱基台達電份額（只作對照，不入表；CF11 待 Andy）"); ST.cell(r, 2, "2025-07 Blackwell PSU／2025-09 AI 伺服器電源")
            ST.cell(r, 4, "=pwr_kgi_0701").number_format = "0.0%"; ST.cell(r, 5, "=pwr_kgi_0926").number_format = "0.0%"; r += 2
            ST.column_dimensions["G"].width = 9
        ST.column_dimensions["A"].width = 30; ST.column_dimensions["B"].width = 26; ST.column_dimensions["C"].width = 34
        for col in "DEF": ST.column_dimensions[col].width = 9

    # ---------------- TK_In：Tokenomics 鏡像（只讀） ----------------
    TKS = wb.create_sheet("TK_In")
    src = tk["source"]
    TKS["A1"] = (f"TK_In — Tokenomics 取數（只讀鏡像；權威在算力收支）：{src['file']}（{src['version']}，commit {src['commit'][:7]}，擷取 {src['extractedAt']}）。"
                 "綠字＝外部連結值，不得在本頁修改；Tokenomics 升版時以 tools/tokenomics/import_tokenomics.py 重抽快照再重建")
    TKS["A2"] = "右側 J–L 欄：company-models 下游目前讀的舊版快照（markets/chip_bridge v20.1）；M 欄＝與現行 TK 不同的格數（>0 表示下游用的值已過時）"
    hdr(TKS, 4, ["名稱（TK）", "世代／項目", "本庫名稱", "低", "基準", "高", "單位", "TK 位置", "取用者",
                 "舊快照 低", "舊快照 基準", "舊快照 高", "差異格數"])
    gens = {g["name"]: g["code"] for g in src["generations"]}
    users = {"IF_RacksPerGW": "chip_bridge、network（經橋接表）", "IF_GPUsPerGW": "chip_bridge", "IF_CapexIT": "chip_bridge、network、delta、quanta 等（部署 GW 換算）",
             "IF_CapexFacility": "chip_bridge", "IF_CapexTotal": "chip_bridge",
             "IF_RackPrice": "chip_bridge（目前抄錄值，改接排入中）", "IF_ScaleOutShare": "chip_bridge（目前抄錄值）", "IF_StoragePerGPU": "chip_bridge（目前抄錄值）"}
    r = 5; TKROWS = []
    old = tk_old["items"] if tk_old else {}
    for nm_, it in tk["items"].items():
        short = nm_[3:]
        if it.get("kind") == "gen_cost":
            for g, vv in it["values"].items():
                code = gens.get(g, g)
                sname = f"TK_{short}_{code}"
                TKS.cell(r, 1, nm_); TKS.cell(r, 2, g); TKS.cell(r, 3, sname)
                for j, c_ in ((4, "低成本"), (5, "基準"), (6, "高成本")):
                    TKS.cell(r, j, vv[c_]).font = GREEN
                TKS.cell(r, 7, it.get("unit")); TKS.cell(r, 8, it.get("ref")); TKS.cell(r, 9, users.get(nm_, "—（本庫已鏡像，待下游改接）"))
                if nm_ in old:
                    ov = old[nm_]["values"][g]
                    for j, c_ in ((10, "低成本"), (11, "基準"), (12, "高成本")):
                        TKS.cell(r, j, ov[c_])
                    TKS.cell(r, 13, f"=(ABS(D{r}-J{r})>1E-9)+(ABS(E{r}-K{r})>1E-9)+(ABS(F{r}-L{r})>1E-9)")
                name(wb, sname, f"TK_In!$E${r}")
                name(wb, sname + "_lo", f"TK_In!$D${r}"); name(wb, sname + "_hi", f"TK_In!$F${r}")
                TKROWS.append((sname, f"{it.get('label')}（{g}）", it.get("unit"), r)); r += 1
        else:
            vals = it.get("values") or it.get("value")
            sname = f"TK_{short}"
            TKS.cell(r, 1, nm_); TKS.cell(r, 2, it.get("label")); TKS.cell(r, 3, sname)
            if isinstance(vals, list):
                for j, x in zip((4, 5, 6), vals):
                    TKS.cell(r, j, x).font = GREEN
            else:
                TKS.cell(r, 5, vals).font = GREEN
            TKS.cell(r, 7, it.get("unit")); TKS.cell(r, 8, it.get("ref")); TKS.cell(r, 9, "avgo、mrvl、amd 等（目前各自假設，待改接）")
            name(wb, sname, f"TK_In!$E${r}")
            name(wb, sname + "_lo", f"TK_In!$D${r}"); name(wb, sname + "_hi", f"TK_In!$F${r}")
            TKROWS.append((sname, it.get("label"), it.get("unit"), r)); r += 1
    TK_LAST = r - 1
    r += 1
    TKS.cell(r, 1, "下游舊快照與現行 TK 不同的格數").font = BOLD
    c = TKS.cell(r, 13, f"=SUM(M5:M{TK_LAST})"); c.fill = KEY; TK_DIFF_CELL = f"TK_In!$M${r}"
    for col, w in zip("ABCDEFGHIJKLM", (22, 34, 26, 9, 9, 9, 10, 20, 34, 9, 9, 9, 8)):
        TKS.column_dimensions[col].width = w

    # ---------------- Deploy：AI IT 資本支出與部署 GW（2026–2029；情境 A／B／C） ----------------
    DP = None; DROW = {}
    if "vr26" in INROW and UN is not None and "it_GB" in INROW:
        DP = wb.create_sheet("Deploy", index=wb.sheetnames.index("Path"))
        DP["A1"] = ("Deploy — AI IT 資本支出＝Σ 平台顆數（Units 頁）× 每顆 IT capex（In 頁 M9）× 情境顆數係數；部署 GW＝AI IT 資本支出 ÷ 每 GW IT 資本支出"
                    "（Tokenomics IF_CapexIT；2026 依 vr26 混合 GB300 與 VR200，2027 起 VR200）。公式與 markets/network、各公司模型相同；R4 產業部署路徑上線後改讀算力收支")
        hdr(DP, 3, ["項目", "情境", "2026", "2027", "2028", "2029"])
        T2P = {"bw": "GB", "rb": "VR", "ru": "RU", "hp": "HP", "tpu": "TPU", "trn": "TRN", "oth": "ASIC"}
        def plat(k, y):
            return ("AMD8" if y == 2026 else "AMDR") if k == "amd" else T2P[k]
        COLS = {2026: "C", 2027: "D", 2028: "E", 2029: "F"}
        r = 4
        DP.cell(r, 1, "AI IT 資本支出（基準，未乘係數）").font = BOLD
        for y, col in COLS.items():
            c = DP.cell(r, list(COLS).index(y) + 3, "=" + "+".join(f"Units!{col}{UROW[g['key']]}*it_{plat(g['key'], y)}" for g in seed["unit_groups"]))
            c.number_format = "0.0"
        DROW["itc"] = r; r += 1
        for s_ in seed["scen"]:
            DP.cell(r, 1, "AI IT 資本支出 × 情境顆數係數"); DP.cell(r, 2, s_)
            for y, col in COLS.items():
                f = f"={col}{DROW['itc']}" + (f"*um{y % 100}_{s_}" if y >= 2028 else "")
                DP.cell(r, list(COLS).index(y) + 3, f).number_format = "0.0"
            DROW["itc" + s_] = r; r += 1
        DP.cell(r, 1, "每 GW IT 資本支出（$B／GW）")
        DP.cell(r, 3, "=TK_CapexIT_GB300*(1-vr26)+TK_CapexIT_VR200*vr26").number_format = "0.00"
        for y in (2027, 2028, 2029):
            DP.cell(r, list(COLS).index(y) + 3, "=TK_CapexIT_VR200").number_format = "0.00"
        DROW["itpg"] = r; r += 1
        for s_ in seed["scen"]:
            DP.cell(r, 1, "部署 GW（IT 電力）").font = BOLD; DP.cell(r, 2, s_)
            for y, col in COLS.items():
                c = DP.cell(r, list(COLS).index(y) + 3, f"={col}{DROW['itc' + s_]}/{col}{DROW['itpg']}"); c.number_format = "0.0"; c.fill = KEY
            DROW["gw" + s_] = r; r += 1
        DP.column_dimensions["A"].width = 36
        for col in "BCDEF": DP.column_dimensions[col].width = 10

    # ---------------- ABF：ABF 載板供需表（v0.5.1；需求與供給皆為 2026＝1 的指數，缺口以 2026 外部缺口校準） ----------------
    AB = None; ABROW = {}
    if "abfm_25" in INROW and DP is not None:
        AB = wb.create_sheet("ABF", index=wb.sheetnames.index("Path"))
        AB["A1"] = ("ABF — 載板供需表。需求（面積 × 層數口徑，2026＝1）＝加速器（顆數 ×（NVIDIA 1、非 NVIDIA v_ot）×（1＋g_asp）^t）、其他 AI（AI IT 資本支出成長 ^ 彈性）、"
                    "非 AI（年增 g_nai），以 2026 金額權重加總；供給＝各載板廠產能指數 × 2026 營收權重，再與味之素 ABF 膜可出貨量上限取小；"
                    "缺口＝（1＋2026 外部缺口）× 需求 ÷ 有效供給 − 1。價格溢價不計入（只量）")
        hdr(AB, 3, ["項目", "情境", "2026", "2027", "2028", "2029", "說明"])
        COLS = {2026: "C", 2027: "D", 2028: "E", 2029: "F"}
        NVK = ("bw", "rb", "ru", "hp"); OTK = ("amd", "tpu", "trn", "oth")
        def cell(r, y, f, fmt="0.000", fill=None):
            c = AB.cell(r, list(COLS).index(y) + 3, f); c.number_format = fmt
            if fill: c.fill = fill
            return c
        r = 4
        AB.cell(r, 1, "A. 需求").font = BOLD; r += 1
        AB.cell(r, 1, "2026 ABF 載板產值（US$B）"); cell(r, 2026, "=abfm_25*(1+abfm_g26)", "0.00"); AB.cell(r, 7, "TPCA 2025 × 2026 年增"); ABROW["mkt"] = r; r += 1
        AB.cell(r, 1, "2026 權重：加速器"); cell(r, 2026, "=abf_ai26*abf_accsh26", "0.0%"); ABROW["wacc"] = r; r += 1
        AB.cell(r, 1, "2026 權重：其他 AI（交換器、CPU、網通）"); cell(r, 2026, "=abf_ai26*(1-abf_accsh26)", "0.0%"); ABROW["woai"] = r; r += 1
        AB.cell(r, 1, "2026 權重：非 AI"); cell(r, 2026, "=1-abf_ai26", "0.0%"); ABROW["wnai"] = r; r += 1
        for s_ in seed["scen"]:
            AB.cell(r, 1, "加速器有效顆數（NVIDIA ＋ 非 NVIDIA × v_ot；百萬顆）"); AB.cell(r, 2, s_)
            for y, col in COLS.items():
                m = (f"*um{y % 100}_{s_}" if y >= 2028 else "")
                cell(r, y, "=(" + "+".join(f"Units!{col}{UROW[k]}" for k in NVK) + "+abf_vot*(" + "+".join(f"Units!{col}{UROW[k]}" for k in OTK) + f")){m}", "0.00")
            ABROW["equ" + s_] = r; r += 1
        for s_ in seed["scen"]:
            AB.cell(r, 1, "需求指數：加速器（× 面積層數）"); AB.cell(r, 2, s_)
            for y, col in COLS.items():
                cell(r, y, f"={col}{ABROW['equ' + s_]}/$C{ABROW['equ' + s_]}*(1+abf_gasp)^{y - 2026}")
            ABROW["dacc" + s_] = r; r += 1
        for s_ in seed["scen"]:
            AB.cell(r, 1, "需求指數：其他 AI（AI IT 資本支出成長 ^ 彈性）"); AB.cell(r, 2, s_)
            for y, col in COLS.items():
                cell(r, y, f"=(Deploy!{col}{DROW['itc' + s_]}/Deploy!$C{DROW['itc' + s_]})^abf_oai_el")
            ABROW["doai" + s_] = r; r += 1
        AB.cell(r, 1, "需求指數：非 AI")
        for y in COLS: cell(r, y, f"=(1+abf_gnai)^{y - 2026}")
        ABROW["dnai"] = r; r += 1
        for s_ in seed["scen"]:
            AB.cell(r, 1, "需求指數合計（2026＝1）").font = BOLD; AB.cell(r, 2, s_)
            for y, col in COLS.items():
                cell(r, y, f"=$C${ABROW['wacc']}*{col}{ABROW['dacc' + s_]}+$C${ABROW['woai']}*{col}{ABROW['doai' + s_]}+$C${ABROW['wnai']}*{col}{ABROW['dnai']}", fill=KEY)
            ABROW["dem" + s_] = r; r += 1
        r += 1
        AB.cell(r, 1, "B. 供給").font = BOLD; r += 1
        MAK = [("Ibiden", "ib", "abfrev_ib26"), ("欣興", "un", "abfrev_un26"), ("南電", "np", "abfrev_np26"), ("景碩", "ks", "abfrev_ks26")]
        wr = {}
        for who, co, rv in MAK:
            AB.cell(r, 1, f"2026 供給權重：{who}"); cell(r, 2026, f"={rv}/$C${ABROW['mkt']}", "0.0%"); wr[co] = r; r += 1
        AB.cell(r, 1, "2026 供給權重：其他（三星電機、AT&S、Shinko、臻鼎等）")
        cell(r, 2026, "=1-" + "-".join(f"C{wr[co]}" for _, co, _ in MAK), "0.0%"); wr["oth"] = r; r += 1
        ABROW["woth"] = r - 1
        for who, co, _ in MAK + [("其他", "oth", None)]:
            AB.cell(r, 1, f"產能指數：{who}")
            for y in COLS: cell(r, y, "=1" if y == 2026 else f"=abfcap_{co}_{y % 100}", "0.00")
            ABROW["cap" + co] = r; r += 1
        AB.cell(r, 1, "載板產能指數（加權）").font = BOLD
        for y, col in COLS.items():
            cell(r, y, "=" + "+".join(f"$C${wr[co]}*{col}{ABROW['cap' + co]}" for co in ("ib", "un", "np", "ks", "oth")), fill=KEY)
        ABROW["sub"] = r; r += 1
        AB.cell(r, 1, "味之素 ABF 膜可出貨量上限（2026＝1）")
        for y in COLS: cell(r, y, "=1" if y == 2026 else f"=abffilm_{y % 100}", "0.00")
        ABROW["film"] = r; r += 1
        AB.cell(r, 1, "有效供給指數（取小）").font = BOLD
        for y, col in COLS.items():
            cell(r, y, f"=MIN({col}{ABROW['sub']},{col}{ABROW['film']})", fill=KEY)
        ABROW["sup"] = r; AB.cell(r, 7, "膜上限低於載板產能時，膜為瓶頸"); r += 2
        AB.cell(r, 1, "C. 缺口（缺口 ÷ 供給；2026 以外部缺口校準）").font = BOLD; r += 1
        for s_ in seed["scen"]:
            AB.cell(r, 1, "ABF 供需缺口").font = BOLD; AB.cell(r, 2, s_)
            for y, col in COLS.items():
                cell(r, y, f"=(1+abf_gap_26)*{col}{ABROW['dem' + s_]}/{col}{ABROW['sup']}-1", "0.0%", KEY)
            ABROW["gap" + s_] = r; r += 1
        AB.cell(r, 1, "對照：外部缺口（In 頁 abf_gap_*；兩來源中點）")
        for y in (2026, 2027, 2028): cell(r, y, f"=abf_gap_{y % 100}", "0.0%")
        ABROW["ext"] = r; r += 1
        AB.cell(r, 1, "對照：外部缺口低／高")
        for y in (2026, 2027, 2028): cell(r, y, f'=TEXT(abf_gap_{y % 100}_lo,"0%")&"–"&TEXT(abf_gap_{y % 100}_hi,"0%")', "@")
        r += 1
        AB.column_dimensions["A"].width = 52; AB.column_dimensions["G"].width = 40
        for col in "BCDEF": AB.column_dimensions[col].width = 10

    # ---------------- PGW：每 GW AI 高速 PCB 與 CCL 金額（v0.5.2） ----------------
    PG = None; PGROW = {}
    if "pcb_cmp_26" in INROW and DP is not None:
        PG = wb.create_sheet("PGW", index=wb.sheetnames.index("Path"))
        PG["A1"] = ("PGW — 每 GW AI 高速 PCB 與 CCL 金額（US$M／GW，不含價格指數）＝ 2026 元件金額 × 2027–29 指數；市場＝每 GW × 部署 GW（Deploy 頁）。"
                    "PCB 口徑權威原在勝宏模型、CCL 口徑原在台光電模型；兩者口徑不同，不可相加")
        hdr(PG, 3, ["項目", "情境", "2026", "2027", "2028", "2029", "說明"])
        COLS = {2026: "C", 2027: "D", 2028: "E", 2029: "F"}
        r = 4
        for pre, ttl, comps in (("pcb", "PCB", (("cmp", "cmpx", "加速器卡／運算板"), ("sw", "swx", "交換器板"), ("mid", "midx", "中板／背板／其他"))),
                                ("ccl", "CCL", (("cmp", "cmpx", "加速器／運算板"), ("sw", "swx", "交換器板"), ("ubb", "ubbx", "UBB／中板／背板")))):
            PG.cell(r, 1, f"每 GW {ttl}").font = BOLD; r += 1
            r0 = r
            for c, cx, lab in comps:
                PG.cell(r, 1, f"每 GW {ttl}：{lab}")
                for y, col in COLS.items():
                    f = f"={pre}_{c}_26" if y == 2026 else f"={pre}_{c}_26*{pre}_{cx}_{y % 100}"
                    PG.cell(r, list(COLS).index(y) + 3, f).number_format = "#,##0"
                r += 1
            PG.cell(r, 1, f"每 GW {ttl} 合計（US$M／GW）").font = BOLD
            for y, col in COLS.items():
                c_ = PG.cell(r, list(COLS).index(y) + 3, f"=SUM({col}{r0}:{col}{r-1})"); c_.number_format = "#,##0"; c_.fill = KEY
            PGROW[pre] = r; r += 1
            for s_ in seed["scen"]:
                PG.cell(r, 1, f"AI 高速 {ttl} 市場（不含價格指數；US$B）"); PG.cell(r, 2, s_)
                for y, col in COLS.items():
                    PG.cell(r, list(COLS).index(y) + 3, f"={col}{PGROW[pre]}*Deploy!{col}{DROW['gw' + s_]}/1000").number_format = "0.0"
                PGROW[pre + "m" + s_] = r; r += 1
            r += 1
        PG.cell(r, 1, "每 GW PCB ÷ 每 GW CCL")
        for y, col in COLS.items():
            PG.cell(r, list(COLS).index(y) + 3, f"={col}{PGROW['pcb']}/{col}{PGROW['ccl']}").number_format = "0.00"
        PGROW["ratio"] = r; r += 1
        PG.column_dimensions["A"].width = 44; PG.column_dimensions["G"].width = 30
        for col in "BCDEF": PG.column_dimensions[col].width = 10

    # ---------------- Rack：AI 伺服器組裝、商用 GPU 系統、AI 電源市場（v0.6） ----------------
    RK = None; RKROW = {}
    if "pwr_pkw_26" in INROW and "asm_odm_gb" in INROW and DP is not None:
        RK = wb.create_sheet("Rack", index=wb.sheetnames.index("Path"))
        RK["A1"] = ("Rack — 機架層市場（2026–2029；情境 A／B／C）：(1) AI 伺服器組裝市場（ODM 價）＝部署 GW × 每 GW 架數（Tokenomics）× ODM 每櫃價格；"
                    "(2) 商用 GPU 系統市場（OEM／整櫃售價，含 GPU）＝商用 GPU 顆數 × 每顆系統價值（Tokenomics 機架價 ÷ 每架 GPU 數）；"
                    "(3) AI 電源市場＝部署 GW × 每 kW 電源內容。2026 部署依 vr26 混合 GB300 與 VR200，2027 起 VR200（ASIC 機架假設每 GW 價值相同）")
        hdr(RK, 3, ["項目", "情境", "2026", "2027", "2028", "2029", "說明"])
        COLS = {2026: "C", 2027: "D", 2028: "E", 2029: "F"}
        def rc(r, y, f, fmt="#,##0.00", fill=None):
            c = RK.cell(r, list(COLS).index(y) + 3, f); c.number_format = fmt
            if fill: c.fill = fill
            return c
        r = 4
        RK.cell(r, 1, "A. AI 伺服器組裝市場（ODM 價；quanta、foxconn、wistron、wiwynn 同讀）").font = BOLD; r += 1
        RK.cell(r, 1, "每 GW ODM 整櫃價值（US$B／GW）")
        rc(r, 2026, "=(TK_RacksPerGW_GB300*asm_odm_gb*(1-vr26)+TK_RacksPerGW_VR200*asm_odm_vr*vr26)/1000")
        for y in (2027, 2028, 2029): rc(r, y, "=TK_RacksPerGW_VR200*asm_odm_vr/1000")
        RK.cell(r, 7, "Tokenomics 每 GW 架數 × ODM 每櫃價格（In 頁 M18）"); RKROW["vpg"] = r; r += 1
        for s_ in seed["scen"]:
            RK.cell(r, 1, "AI 伺服器組裝市場（US$B）").font = BOLD; RK.cell(r, 2, s_)
            for y, col in COLS.items(): rc(r, y, f"=Deploy!{col}{DROW['gw' + s_]}*{col}{RKROW['vpg']}", "#,##0.0", KEY)
            RKROW["asm" + s_] = r; r += 1
        r += 1
        RK.cell(r, 1, "B. 商用 GPU 系統市場（OEM／整櫃售價；dell、smci 同讀；CF10）").font = BOLD; r += 1
        RK.cell(r, 1, "每顆系統價值：GB300 世代（US$K／顆）")
        for y in COLS: rc(r, y, "=TK_RackPrice_GB300/(TK_GPUsPerGW_GB300/TK_RacksPerGW_GB300)*1000")
        RK.cell(r, 7, "Tokenomics 機架價 ÷ 每架 GPU 數；HGX 八卡與 AMD MI355 以同值近似（Analogy）"); RKROW["vgb"] = r; r += 1
        RK.cell(r, 1, "每顆系統價值：Rubin 世代（US$K／顆）")
        for y in COLS: rc(r, y, "=TK_RackPrice_VR200/(TK_GPUsPerGW_VR200/TK_RacksPerGW_VR200)*1000")
        RK.cell(r, 7, "Rubin Ultra 與 AMD MI450 整櫃以 VR200 同值近似（Analogy）"); RKROW["vvr"] = r; r += 1
        def ucell(k, col): return f"Units!{col}{UROW[k]}"
        RK.cell(r, 1, "GB300 世代商用 GPU 顆數（Blackwell＋Hopper；2026 另含 AMD MI355；百萬顆，情境 A）")
        for y, col in COLS.items():
            rc(r, y, "=" + "+".join(ucell(k, col) for k in ("bw", "hp")) + (f"+{ucell('amd', col)}" if y == 2026 else ""), "0.000")
        RK.cell(r, 7, "世代分組依 AMD 實際產品：2026 MI355（八卡，GB300 世代）、2027 起 MI450 Helios 整櫃（Rubin 世代）"); RKROW["ugb"] = r; r += 1
        RK.cell(r, 1, "Rubin 世代商用 GPU 顆數（Rubin＋Rubin Ultra；2027 起另含 AMD MI450；百萬顆，情境 A）")
        for y, col in COLS.items():
            rc(r, y, "=" + "+".join(ucell(k, col) for k in ("rb", "ru")) + (f"+{ucell('amd', col)}" if y >= 2027 else ""), "0.000")
        RKROW["uvr"] = r; r += 1
        for s_ in seed["scen"]:
            RK.cell(r, 1, "商用 GPU 系統市場（US$B）").font = BOLD; RK.cell(r, 2, s_)
            for y, col in COLS.items():
                m = (f"*um{y % 100}_{s_}" if y >= 2028 else "")
                rc(r, y, f"=({col}{RKROW['ugb']}*{col}{RKROW['vgb']}+{col}{RKROW['uvr']}*{col}{RKROW['vvr']}){m}", "#,##0.0", KEY)
            RKROW["sys" + s_] = r; r += 1
        r += 1
        RK.cell(r, 1, "C. AI 電源市場（delta、liteon 同讀）").font = BOLD; r += 1
        RK.cell(r, 1, "AI 電源每 kW 內容（US$/kW）")
        for y in COLS: rc(r, y, f"=pwr_pkw_{y % 100}", "#,##0")
        RK.cell(r, 7, "機架電源櫃、PSU、BBU、HVDC 電源櫃（In 頁 M17）"); RKROW["pkw"] = r; r += 1
        for s_ in seed["scen"]:
            RK.cell(r, 1, "AI 電源市場（US$B）").font = BOLD; RK.cell(r, 2, s_)
            for y, col in COLS.items(): rc(r, y, f"=Deploy!{col}{DROW['gw' + s_]}*{col}{RKROW['pkw']}/1000", "#,##0.00", KEY)
            RKROW["pwr" + s_] = r; r += 1
        RK.cell(r, 1, "對照：UBS 全球 AI 伺服器電源市場（US$B；口徑未找到）")
        rc(r, 2026, "=pwr_ubs_26", "#,##0.0"); rc(r, 2027, "=pwr_ubs_27", "#,##0.0"); RKROW["ubs"] = r; r += 1
        RK.cell(r, 1, "本庫情境 A ÷ UBS")
        for y, col in ((2026, "C"), (2027, "D")): rc(r, y, f"={col}{RKROW['pwrA']}/{col}{RKROW['ubs']}", "0.00")
        RK.cell(r, 7, "UBS 高出約五到七成：口徑（是否含 BBU、HVDC、非 AI）未找到，只作對照"); r += 1
        RK.column_dimensions["A"].width = 62; RK.column_dimensions["G"].width = 60
        for col in "BCDEF": RK.column_dimensions[col].width = 10

    # ---------------- Conflicts：跨模型同一變數不同值的登錄 ----------------
    CF = wb.create_sheet("Conflicts")
    CF["A1"] = ("Conflicts — 跨模型衝突登錄：同一個行業變數在兩個以上公司模型有不同值。本庫的任務是讓每個變數只有一個值；"
                "建議值為 Claude 預設（判斷類，Andy 可改），在對應子產業併入本庫的版本生效，生效前各公司模型維持現值")
    hdr(CF, 3, ["ID", "子產業", "變數", "單位", "模型甲（位置）", "甲值", "模型乙（位置）", "乙值", "性質", "建議單一值與理由", "生效版本", "狀態", "來源"])
    r = 4
    for cf in conflicts:
        for j, k in enumerate(("id", "sub", "var", "unit", "a_loc", "a_val", "b_loc", "b_val", "kind", "proposal", "target", "status", "origin"), 1):
            c = CF.cell(r, j, cf.get(k)); c.alignment = WRAP
        if cf.get("status", "").startswith("待 Andy"):
            for j in range(1, 14): CF.cell(r, j).fill = PEND
        r += 1
    CF_LAST = max(r - 1, 4)
    for col, w in zip("ABCDEFGHIJKLM", (7, 10, 30, 9, 26, 12, 26, 12, 12, 52, 9, 14, 22)):
        CF.column_dimensions[col].width = w

    # ---------------- Interface：下游唯一連結點 ----------------
    IF = wb.create_sheet("Interface")
    IF["A1"] = ("Interface — 下游唯一連結點（各公司模型與市場層只讀本頁 SIF_ 具名範圍，經 tools/export_snapshot.py 匯成版本固定快照）。"
                "A 節＝輸入（純量，含低／高）；B 節＝推導（年度）；C 節＝Tokenomics 鏡像（SIF_TK_*，值＝TK_In）")
    IF["A2"] = f"版本：Siliconomics {VERSION}（{DATE}）；Tokenomics {src['version']}"
    hdr(IF, 4, ["名稱", "項目", "單位", "低", "基準", "高", "標記", "來源位置"])
    r = 5
    IF.cell(r, 1, "A. 輸入（純量）").font = BOLD; r += 1
    IFNAMES = []
    for k, (ir, lab, unit, has_lo, has_hi, tag) in INROW.items():
        nm_ = "SIF_" + k
        IF.cell(r, 1, nm_); IF.cell(r, 2, lab); IF.cell(r, 3, unit)
        if has_lo: IF.cell(r, 4, f"=In!D{ir}").font = GREEN
        c = IF.cell(r, 5, f"=In!E{ir}"); c.font = GREEN
        if has_hi: IF.cell(r, 6, f"=In!F{ir}").font = GREEN
        IF.cell(r, 7, f"=In!G{ir}"); IF.cell(r, 8, f"In!E{ir}")
        name(wb, nm_, f"Interface!$E${r}")
        if has_lo: name(wb, nm_ + "_lo", f"Interface!$D${r}")
        if has_hi: name(wb, nm_ + "_hi", f"Interface!$F${r}")
        IFNAMES.append(nm_); r += 1
    r += 1
    IF.cell(r, 1, "B. 推導（日曆年；欄 D–H＝2025–2029）").font = BOLD; r += 1
    hdr(IF, r, ["名稱", "項目", "單位", "2025", "2026", "2027", "2028", "2029", "來源位置"]); r += 1
    M = wb["Mkt"]
    derived = [("SIF_HBM_DemandEB", "dem", "HBM 需求（加速器裝載，情境 A 基準）", "EB"),
               ("SIF_HBM_DemandGrossEB", "demg", "HBM 需求 × 出貨係數（情境 A 基準）", "EB"),
               ("SIF_Units_Total", "units", "加速器顆數合計（情境 A 基準）", "百萬顆"),
               ("SIF_HBM_GBperUnit", "gbavg", "每顆平均 HBM", "GB"),
               ("SIF_HBM_SupplyEB", "sup", "HBM 位元供給（三情境相同）", "EB")]
    for s in seed["scen"]:
        derived += [(f"SIF_UnitsCoef_{s}", "cf" + s, f"顆數係數（{s}）", "x"),
                    (f"SIF_HBM_ShipEB_{s}", "ship" + s, f"HBM 出貨位元（{s}）", "EB"),
                    (f"SIF_HBM_GapEB_{s}", "gap" + s, f"HBM 需求 − 供給（{s}）", "EB"),
                    (f"SIF_HBM_Price_{s}", "p" + s, f"HBM 單價（{s}）", "$/GB"),
                    (f"SIF_HBM_Mkt_{s}", "m" + s, f"HBM 市場（{s}）", "$B")]
    for nm_, key, lab, unit in derived:
        mr = MK[key]
        IF.cell(r, 1, nm_); IF.cell(r, 2, lab); IF.cell(r, 3, unit)
        for j in range(5):
            c = IF.cell(r, 4 + j, f'=IF(ISBLANK(Mkt!{"CDEFG"[j]}{mr}),"",Mkt!{"CDEFG"[j]}{mr})'); c.font = GREEN; c.number_format = "0.00"
        IF.cell(r, 9, f"Mkt!C{mr}:G{mr}")
        name(wb, nm_, f"Interface!$D${r}:$H${r}")
        IFNAMES.append(nm_); r += 1
    if UN is not None:
        r += 1
        IF.cell(r, 1, "B2. 平台別顆數（Units 頁；欄 D–H＝2025–2029，2025 空白）").font = BOLD; r += 1
        items = ([(f"SIF_UnitsPlat_{g['key']}", g["key"], f"{g['name']} 顆數（情境 A）") for g in seed["unit_groups"]]
                 + [("SIF_UnitsPlat_Total", "tot", "加速器顆數合計（情境 A；含 2029）")]
                 + [(f"SIF_UnitsPlat_Total_{s_}", "tot" + s_, f"加速器顆數合計（{s_}）") for s_ in seed["scen"]])
        for nm_, key, lab in items:
            ur = UROW[key]
            IF.cell(r, 1, nm_); IF.cell(r, 2, lab); IF.cell(r, 3, "百萬顆")
            IF.cell(r, 4, "")
            for j in range(4):
                c = IF.cell(r, 5 + j, f"=Units!{'CDEF'[j]}{ur}"); c.font = GREEN; c.number_format = "0.00"
            IF.cell(r, 9, f"Units!C{ur}:F{ur}")
            name(wb, nm_, f"Interface!$D${r}:$H${r}")
            IFNAMES.append(nm_); r += 1
        IF.cell(r, 1, "SIF_UnitsGrowth29"); IF.cell(r, 2, "2029 加速器顆數年增（情境 A；平台別成長率加權）"); IF.cell(r, 3, "%")
        c = IF.cell(r, 5, f"=Units!G{UROW['tot']}"); c.font = GREEN; c.number_format = "0.0%"
        IF.cell(r, 9, f"Units!G{UROW['tot']}")
        name(wb, "SIF_UnitsGrowth29", f"Interface!$E${r}"); IFNAMES.append("SIF_UnitsGrowth29"); r += 1
    if ST is not None:
        r += 1
        IF.cell(r, 1, "B3. 份額合計（ShareTbl 頁；欄 E–G＝ASIC 2026–2028、CPO 2027–2029；ABF 載板 E–H＝2026–2029）").font = BOLD; r += 1
        for code, lab in [("tpu", "Google TPU"), ("oth", "其他 ASIC")] + [("trn" + s_, f"AWS Trainium（情境 {s_}）") for s_ in seed["scen"]] + ([("cpo", "CPO 光學（2027–2029）")] if "sum_cpo" in STROW else []) + ([("abfnv", "NVIDIA 加速器 ABF 載板（2026–2029）"), ("abfot", "非 NVIDIA 加速器 ABF 載板（2026–2029）")] if "sum_abfnv" in STROW else []):
            nm_ = {"cpo": "SIF_CPOShareSum", "abfnv": "SIF_ABFShareSum_nv", "abfot": "SIF_ABFShareSum_ot"}.get(code, f"SIF_ASICShareSum_{code}")
            sr = STROW["sum_" + code]
            n_ = 4 if code.startswith("abf") else 3
            IF.cell(r, 1, nm_); IF.cell(r, 2, f"{lab} {'光學金額份額合計' if code == 'cpo' else ('載板份額合計' if n_ == 4 else '設計夥伴份額合計')}"); IF.cell(r, 3, "%")
            for j in range(n_):
                c = IF.cell(r, 5 + j, f"=ShareTbl!{'DEFG'[j]}{sr}"); c.font = GREEN; c.number_format = "0%"
            IF.cell(r, 9, f"ShareTbl!D{sr}:{'DEFG'[n_-1]}{sr}")
            name(wb, nm_, f"Interface!$E${r}:${'EFGH'[n_-1]}${r}"); IFNAMES.append(nm_); r += 1
    if DP is not None:
        r += 1
        IF.cell(r, 1, "B4. AI IT 資本支出與部署 GW（Deploy 頁；欄 E–H＝2026–2029）").font = BOLD; r += 1
        items = ([(f"SIF_AIITCapex_{s_}", "itc" + s_, f"AI IT 資本支出（{s_}）", "$B") for s_ in seed["scen"]]
                 + [("SIF_ITCapexPerGW", "itpg", "每 GW IT 資本支出（2026 依 vr26 混合）", "$B/GW")]
                 + [(f"SIF_DeployGW_{s_}", "gw" + s_, f"部署 GW（{s_}）", "GW") for s_ in seed["scen"]])
        for nm_, key, lab, unit in items:
            dr = DROW[key]
            IF.cell(r, 1, nm_); IF.cell(r, 2, lab); IF.cell(r, 3, unit)
            for j in range(4):
                c = IF.cell(r, 5 + j, f"=Deploy!{'CDEF'[j]}{dr}"); c.font = GREEN; c.number_format = "0.00"
            IF.cell(r, 9, f"Deploy!C{dr}:F{dr}")
            name(wb, nm_, f"Interface!$E${r}:$H${r}"); IFNAMES.append(nm_); r += 1
    if AB is not None:
        r += 1
        IF.cell(r, 1, "B5. ABF 供需（ABF 頁；欄 E–H＝2026–2029）").font = BOLD; r += 1
        items = ([(f"SIF_ABFDemIdx_{s_}", "dem" + s_, f"ABF 需求指數（{s_}；2026＝1）", "x") for s_ in seed["scen"]]
                 + [("SIF_ABFSubCapIdx", "sub", "載板產能指數（加權；2026＝1）", "x"), ("SIF_ABFFilmIdx", "film", "味之素 ABF 膜上限（2026＝1）", "x"), ("SIF_ABFSupIdx", "sup", "有效供給指數", "x")]
                 + [(f"SIF_ABFGap_{s_}", "gap" + s_, f"ABF 供需缺口（{s_}）", "%") for s_ in seed["scen"]])
        for nm_, key, lab, unit in items:
            ar = ABROW[key]
            IF.cell(r, 1, nm_); IF.cell(r, 2, lab); IF.cell(r, 3, unit)
            for j in range(4):
                c = IF.cell(r, 5 + j, f"=ABF!{'CDEF'[j]}{ar}"); c.font = GREEN; c.number_format = "0.0%" if unit == "%" else "0.000"
            IF.cell(r, 9, f"ABF!C{ar}:F{ar}")
            name(wb, nm_, f"Interface!$E${r}:$H${r}"); IFNAMES.append(nm_); r += 1
    if PG is not None:
        r += 1
        IF.cell(r, 1, "B6. 每 GW PCB／CCL（PGW 頁；欄 E–H＝2026–2029）").font = BOLD; r += 1
        items = ([("SIF_PCBperGW", "pcb", "每 GW AI 高速 PCB 金額（US$M／GW）", "US$M/GW"), ("SIF_CCLperGW", "ccl", "每 GW 高速 CCL 金額（US$M／GW）", "US$M/GW")]
                 + [(f"SIF_PCBMkt_{s_}", "pcbm" + s_, f"AI 高速 PCB 市場（{s_}；不含價格指數）", "US$B") for s_ in seed["scen"]]
                 + [(f"SIF_CCLMkt_{s_}", "cclm" + s_, f"AI 高速 CCL 市場（{s_}；不含價格指數）", "US$B") for s_ in seed["scen"]])
        for nm_, key, lab, unit in items:
            pr = PGROW[key]
            IF.cell(r, 1, nm_); IF.cell(r, 2, lab); IF.cell(r, 3, unit)
            for j in range(4):
                c = IF.cell(r, 5 + j, f"=PGW!{'CDEF'[j]}{pr}"); c.font = GREEN; c.number_format = "#,##0.0"
            IF.cell(r, 9, f"PGW!C{pr}:F{pr}")
            name(wb, nm_, f"Interface!$E${r}:$H${r}"); IFNAMES.append(nm_); r += 1
    if RK is not None:
        r += 1
        IF.cell(r, 1, "B7. 機架層市場（Rack 頁；欄 E–H＝2026–2029）").font = BOLD; r += 1
        items = ([("SIF_ODMValPerGW", "vpg", "每 GW ODM 整櫃價值", "US$B/GW")]
                 + [(f"SIF_AsmMkt_{s_}", "asm" + s_, f"AI 伺服器組裝市場（ODM 價；{s_}）", "US$B") for s_ in seed["scen"]]
                 + [("SIF_SysValGB", "vgb", "每顆系統價值（GB300 世代）", "US$K"), ("SIF_SysValVR", "vvr", "每顆系統價值（Rubin 世代）", "US$K"),
                    ("SIF_SysUnitsGB", "ugb", "GB300 世代商用 GPU 顆數（情境 A）", "百萬顆"), ("SIF_SysUnitsVR", "uvr", "Rubin 世代商用 GPU 顆數（情境 A）", "百萬顆")]
                 + [(f"SIF_GPUSysMkt_{s_}", "sys" + s_, f"商用 GPU 系統市場（{s_}）", "US$B") for s_ in seed["scen"]]
                 + [(f"SIF_PwrMkt_{s_}", "pwr" + s_, f"AI 電源市場（{s_}）", "US$B") for s_ in seed["scen"]])
        for nm_, key, lab, unit in items:
            kr = RKROW[key]
            IF.cell(r, 1, nm_); IF.cell(r, 2, lab); IF.cell(r, 3, unit)
            for j in range(4):
                c = IF.cell(r, 5 + j, f"=Rack!{'CDEF'[j]}{kr}"); c.font = GREEN; c.number_format = "#,##0.000"
            IF.cell(r, 9, f"Rack!C{kr}:F{kr}")
            name(wb, nm_, f"Interface!$E${r}:$H${r}"); IFNAMES.append(nm_); r += 1
    if ST is not None and "sum_pwrA" in STROW:
        r += 1
        IF.cell(r, 1, "B8. AI 電源份額合計（ShareTbl 頁；欄 E–H＝2026–2029）").font = BOLD; r += 1
        for s_ in seed["scen"]:
            nm_ = f"SIF_PwrShareSum_{s_}"; sr = STROW["sum_pwr" + s_]
            IF.cell(r, 1, nm_); IF.cell(r, 2, f"台達電＋光寶 AI 電源有效份額合計（{s_}）"); IF.cell(r, 3, "%")
            for j in range(4):
                c = IF.cell(r, 5 + j, f"=ShareTbl!{'DEFG'[j]}{sr}"); c.font = GREEN; c.number_format = "0.0%"
            IF.cell(r, 9, f"ShareTbl!D{sr}:G{sr}")
            name(wb, nm_, f"Interface!$E${r}:$H${r}"); IFNAMES.append(nm_); r += 1
    r += 1
    IF.cell(r, 1, "C. Tokenomics 鏡像（只讀；權威在算力收支；本庫與 TK 一致由 Checks K3 檢查）").font = BOLD; r += 1
    hdr(IF, r, ["名稱", "項目", "單位", "低", "基準", "高", "", "來源位置"]); r += 1
    for sname, lab, unit, tr in TKROWS:
        nm_ = "SIF_" + sname
        IF.cell(r, 1, nm_); IF.cell(r, 2, lab); IF.cell(r, 3, unit)
        for j, col in ((4, "D"), (5, "E"), (6, "F")):
            IF.cell(r, j, f'=IF(ISBLANK(TK_In!{col}{tr}),"",TK_In!{col}{tr})').font = GREEN
        IF.cell(r, 8, f"TK_In!E{tr}")
        name(wb, nm_, f"Interface!$E${r}")
        name(wb, nm_ + "_lo", f"Interface!$D${r}"); name(wb, nm_ + "_hi", f"Interface!$F${r}")
        IFNAMES.append(nm_); r += 1
    for col, w in zip("ABCDEFGHI", (30, 50, 9, 10, 10, 10, 10, 14, 16)):
        IF.column_dimensions[col].width = w

    # ---------------- Checks ----------------
    CK = wb.create_sheet("Checks")
    CK["A1"] = "Checks — 交叉檢查與治理（ERROR 計入 SGOV_Errors，須為 0；WARN 計入 SGOV_Warn，列為待補）"
    hdr(CK, 3, ["ID", "項目", "本庫值", "參照值", "差", "判讀", "性質", "說明"])
    r = 4
    errs, warns = [], []
    # K1 搬移忠實度：HBM 市場與 SK 海力士成品 Mkt 頁逐值相同
    ref_vals = {}
    try:
        fn = [f for f in os.listdir(REF_SKH) if f.endswith(".xlsx")][0]
        rw = openpyxl.load_workbook(os.path.join(REF_SKH, fn), data_only=True)["Mkt"]
        for rr in range(1, rw.max_row + 1):
            lab = rw.cell(rr, 1).value
            if isinstance(lab, str) and lab.startswith("HBM 市場（"):
                ref_vals[lab] = [rw.cell(rr, c).value for c in range(3, 8)]
        ref_name = f"skhynix/dist/{fn}"
    except Exception as e:  # 找不到成品時只記錄
        ref_name = f"（讀取失敗：{e}）"
    for s in seed["scen"]:
        lab = f"HBM 市場（{s}）"
        for j, y in enumerate(YEARS):
            rv = ref_vals.get(lab, [None] * 5)[j]
            CK.cell(r, 1, f"K1-{s}{y}"); CK.cell(r, 2, f"{lab} {y}：本庫 vs SK 海力士模型成品")
            CK.cell(r, 3, f"=Mkt!{'CDEFG'[j]}{MK['m' + s]}")
            CK.cell(r, 4, rv)
            CK.cell(r, 5, f"=IF(ISNUMBER(D{r}),C{r}-D{r},\"\")")
            CK.cell(r, 6, f'=IF(NOT(ISNUMBER(D{r})),"WARN",IF(ABS(E{r})<1E-9,"OK","ERROR"))')
            CK.cell(r, 7, "驗證（搬移忠實度）"); CK.cell(r, 8, ref_name)
            r += 1
    # K2 情境係數
    for s in seed["scen"]:
        for y in (28, 29):
            CK.cell(r, 1, f"K2-{s}{y}"); CK.cell(r, 2, f"情境 {s} 20{y} 顆數係數：A＝1；B、C 在區間內")
            CK.cell(r, 3, f"=um{y}_{s}")
            if s == "A":
                CK.cell(r, 6, f'=IF(ABS(C{r}-1)<1E-9,"OK","ERROR")')
            else:
                CK.cell(r, 4, f'=um{y}_{s}_lo&"–"&um{y}_{s}_hi')
                CK.cell(r, 6, f'=IF(AND(C{r}>=um{y}_{s}_lo,C{r}<=um{y}_{s}_hi),"OK","ERROR")')
            CK.cell(r, 7, "驗證"); r += 1
        # 單調：A ≥ B ≥ C
    for y in (28, 29):
        CK.cell(r, 1, f"K2-ord{y}"); CK.cell(r, 2, f"20{y} 顆數係數 A ≥ B ≥ C")
        CK.cell(r, 6, f'=IF(AND(um{y}_A>=um{y}_B,um{y}_B>=um{y}_C),"OK","ERROR")'); CK.cell(r, 7, "驗證"); r += 1
    # K3 TK 一致
    CK.cell(r, 1, "K3"); CK.cell(r, 2, "下游（chip_bridge 等）讀的 Tokenomics 快照與現行 TK 不同的格數")
    CK.cell(r, 3, f"={TK_DIFF_CELL}"); CK.cell(r, 6, f'=IF(C{r}=0,"OK","WARN")'); CK.cell(r, 7, "治理（TK 一致）")
    CK.cell(r, 8, "WARN＝下游仍用舊版 TK；本庫已用現行版；差異明細見 TK_In 頁 M 欄"); r += 1
    if UN is not None:
        for j, y in enumerate((2026, 2027, 2028)):
            CK.cell(r, 1, f"K4-{y}"); CK.cell(r, 2, f"{y} 平台別顆數合計（Units）＝ HBM 市場頁顆數合計（Mkt）")
            CK.cell(r, 3, f"=Units!{'CDE'[j]}{UROW['tot']}"); CK.cell(r, 4, f"=Mkt!{'DEF'[j]}{MK['units']}")
            CK.cell(r, 5, f"=C{r}-D{r}"); CK.cell(r, 6, f'=IF(ABS(E{r})<1E-9,"OK","ERROR")'); CK.cell(r, 7, "恆等式（結構）"); r += 1
        for g in seed["unit_groups"]:
            k = g["key"]
            CK.cell(r, 1, f"K5-{k}"); CK.cell(r, 2, f"2029 成長率在區間內：{g['name']}")
            CK.cell(r, 3, f"=g29_{k}"); CK.cell(r, 4, f'=g29_{k}_lo&"–"&g29_{k}_hi')
            CK.cell(r, 6, f'=IF(AND(C{r}>=g29_{k}_lo,C{r}<=g29_{k}_hi),"OK","ERROR")'); CK.cell(r, 7, "驗證"); r += 1
    if ST is not None:
        for code, lab in [("tpu", "Google TPU"), ("oth", "其他 ASIC")] + [("trn" + s_, f"AWS Trainium（情境 {s_}）") for s_ in seed["scen"]] + ([("cpo", "CPO 光學（模組與雷射廠）")] if "sum_cpo" in STROW else []) + ([("abfnv", "NVIDIA 加速器 ABF 載板"), ("abfot", "非 NVIDIA 加速器 ABF 載板")] if "sum_abfnv" in STROW else []):
            for j, y in enumerate((2027, 2028, 2029) if code == "cpo" else ((2026, 2027, 2028, 2029) if code.startswith("abf") else (2026, 2027, 2028))):
                CK.cell(r, 1, f"K6-{code}{y}"); CK.cell(r, 2, f"{lab} {y} {'光學金額份額合計' if code == 'cpo' else ('載板份額合計' if code.startswith('abf') else '設計夥伴份額合計')} ≤ 100%")
                CK.cell(r, 3, f"=ShareTbl!{'DEFG'[j]}{STROW['sum_' + code]}"); CK.cell(r, 4, 1)
                CK.cell(r, 6, f'=IF(C{r}<=1+1E-9,"OK","ERROR")'); CK.cell(r, 7, "驗證（加總受限）"); r += 1
    if ST is not None and "sum_pwrA" in STROW:
        for s_ in seed["scen"]:
            for j, y in enumerate((2026, 2027, 2028, 2029)):
                CK.cell(r, 1, f"K6-pwr{s_}{y}"); CK.cell(r, 2, f"AI 電源 {y} 台達電＋光寶有效份額合計 ≤ 100%（情境 {s_}）")
                CK.cell(r, 3, f"=ShareTbl!{'DEFG'[j]}{STROW['sum_pwr' + s_]}"); CK.cell(r, 4, 1)
                CK.cell(r, 6, f'=IF(C{r}<=1+1E-9,"OK","ERROR")'); CK.cell(r, 7, "驗證（加總受限）"); r += 1
    if RK is not None:
        CK.cell(r, 1, "K7"); CK.cell(r, 2, "Rack 頁：商用 GPU 系統市場 2026 ≤ AI IT 資本支出 2026（系統售價不應超過全部 IT 資本支出）")
        CK.cell(r, 3, f"=Rack!C{RKROW['sysA']}"); CK.cell(r, 4, f"=Deploy!C{DROW['itcA']}")
        CK.cell(r, 6, f'=IF(C{r}<=D{r}+1E-9,"OK","ERROR")'); CK.cell(r, 7, "驗證（量級）"); r += 1
    # G 治理：輸入列標記、來源、區間
    g0 = r
    CK.cell(r, 1, "G1"); CK.cell(r, 2, "In 頁輸入列缺標記")
    CK.cell(r, 3, f'=SUMPRODUCT((In!A5:A{IN_LAST}<>"")*(In!E5:E{IN_LAST}<>"")*(In!G5:G{IN_LAST}=""))')
    CK.cell(r, 6, f'=IF(C{r}=0,"OK","ERROR")'); CK.cell(r, 7, "治理"); r += 1
    CK.cell(r, 1, "G2"); CK.cell(r, 2, "In 頁輸入列缺來源")
    CK.cell(r, 3, f'=SUMPRODUCT((In!A5:A{IN_LAST}<>"")*(In!E5:E{IN_LAST}<>"")*(In!H5:H{IN_LAST}=""))')
    CK.cell(r, 6, f'=IF(C{r}=0,"OK","ERROR")'); CK.cell(r, 7, "治理")
    CK.cell(r, 8, "v0.2 起應為 0"); r += 1
    n_noint = sum(1 for k, v in INROW.items() if ("Assumed" in str(v[5]) or "Analogy" in str(v[5])) and not (v[3] and v[4]))
    CK.cell(r, 1, "G3"); CK.cell(r, 2, "標記含 Assumed／Analogy 但未給低／高區間的輸入列（Tokenomics 規則：一律給區間）")
    CK.cell(r, 3, f'=SUMPRODUCT((In!A5:A{IN_LAST}<>"")*(In!E5:E{IN_LAST}<>"")*(ISNUMBER(SEARCH("Assumed",In!G5:G{IN_LAST}))+ISNUMBER(SEARCH("Analogy",In!G5:G{IN_LAST}))>0)*((In!D5:D{IN_LAST}="")+(In!F5:F{IN_LAST}="")>0))')
    CK.cell(r, 6, f'=IF(C{r}=0,"OK","ERROR")'); CK.cell(r, 7, "治理")
    CK.cell(r, 8, f"v0.2 起應為 0（建置時 {n_noint} 列）"); r += 1
    keys = [k for k in INROW]
    CK.cell(r, 1, "G4"); CK.cell(r, 2, "In 頁鍵重複（建置時檢查）"); CK.cell(r, 3, len(keys) - len(set(keys)))
    CK.cell(r, 6, f'=IF(C{r}=0,"OK","ERROR")'); CK.cell(r, 7, "治理"); r += 1
    CK.cell(r, 1, "G5"); CK.cell(r, 2, "衝突登錄中『待 Andy』或『待併入』的項目數")
    CK.cell(r, 3, f'=COUNTIF(Conflicts!L4:L{CF_LAST},"待*")'); CK.cell(r, 6, f'=IF(C{r}=0,"OK","WARN")'); CK.cell(r, 7, "治理")
    CK.cell(r, 8, "WARN＝仍有同一變數兩個值；各子產業併入本庫後歸零"); r += 1
    r += 1
    CK.cell(r, 1, "SGOV_Errors").font = BOLD; c = CK.cell(r, 3, f'=COUNTIF(F4:F{r-2},"ERROR")'); c.fill = KEY
    name(wb, "SGOV_Errors", f"Checks!$C${r}"); r += 1
    CK.cell(r, 1, "SGOV_Warn").font = BOLD; c = CK.cell(r, 3, f'=COUNTIF(F4:F{r-3},"WARN")'); c.fill = KEY
    name(wb, "SGOV_Warn", f"Checks!$C${r}")
    for col, w in zip("ABCDEFGH", (10, 58, 12, 12, 10, 8, 18, 60)):
        CK.column_dimensions[col].width = w

    # ---------------- Decisions、DB_Evidence ----------------
    DC = wb.create_sheet("Decisions")
    DC["A1"] = "Decisions — 決策登錄（Excel 擁有）。判斷類由 Andy 決定；Claude 依預設執行者標『預設、可撤回』"
    hdr(DC, 3, ["決策 ID", "類別", "項目", "決議", "日期", "Andy 原話／依據", "狀態"])
    r = 4
    for d in decisions:
        for j, k in enumerate(("id", "cat", "item", "decision", "date", "quote", "status"), 1):
            c = DC.cell(r, j, d.get(k)); c.alignment = WRAP
        if "預設" in str(d.get("status")):
            for j in range(1, 8): DC.cell(r, j).fill = PEND
        r += 1
    for col, w in zip("ABCDEFG", (9, 12, 24, 60, 12, 50, 16)):
        DC.column_dimensions[col].width = w
    EV = wb.create_sheet("DB_Evidence")
    EV["A1"] = "DB_Evidence — 證據登錄（所有比對過的新資訊，含不採納者）；先與本庫現值比對，再判定採納／部分採納／不採納／待查"
    hdr(EV, 3, ["ID", "日期", "主張（摘要）", "來源", "標記", "對應鍵", "本庫當時值", "新資訊值", "判定", "處理版本", "立場", "備註"])
    EV.cell(4, 1, "（v0.1 尚無；跨模型衝突先登錄於 Conflicts 頁，判定後再轉入本頁）")

    # ---------------- README ----------------
    lines = [
        (f"Siliconomics {VERSION} — AI 半導體產業基準庫（{DATE}）", True),
        ("用途：AI 半導體各子產業的『唯一事實』——加速器顆數、HBM 與記憶體價格、網通內容、CoWoS、載板、伺服器組裝、電源散熱、設備等行業變數只有一個值；各公司模型只讀本庫 Interface 的 SIF_ 名稱。", False),
        ("與 Tokenomics 的分工：每 GW 硬體量、機架價、每 GW 資本支出、NonNV 比例、產業年度部署路徑（R4）權威在 Tokenomics／算力收支，本庫只在 TK_In 頁鏡像（綠字、不得修改）；晶片產業資訊（HBM 單價與容量、每顆內容、網通內容、份額、出貨顆數、ASIC 規格、晶圓與設備）權威在本庫（《情境卡產生規則》第 13 條）。", False),
        ("v0.1 範圍：(1) Tokenomics v21.7 鏡像（22 個名稱）與下游舊快照差異；(2) 加速器顆數層、記憶體價格路徑、HBM 市場（由 markets/memory/market.py 逐字搬入，數值不變）；(3) 跨模型衝突登錄；(4) Interface、Checks、Decisions。網通、CoWoS、載板、組裝電源散熱、設備依規劃書分段併入。", False),
        ("v0.2：In 頁 Assumed 列全部補上區間、缺來源列補來源；新增 M5 節 2029 平台別顆數成長（單一來源，取代台積電、網通模型的寫死值與 NAND 模型的合計年增）與 Units 頁；Interface 新增 SIF_UnitsPlat_*、SIF_UnitsGrowth29。", False),
        ("v0.3：新增 M6 節（加速器每顆售價、ASIC 核心價、HBM 代購比例；原在 nvidia、avgo 模型）與 M7 節（ASIC 設計服務份額；原在 avgo、mrvl、mediatek、alchip 模型）、ShareTbl 頁（同一平台份額加總 ≤ 100%）；In 頁新增 J 欄第二來源。", False),
        ("v0.6：新增 M17（AI 電源每 kW 內容、台達電與光寶份額路徑、凱基與 UBS 外部錨點；原在 delta、liteon）、M18（ODM 每櫃價格、緯穎 AI 占比單一值〔CF09〕；原在 quanta、wiwynn、wistron）與 Rack 頁（AI 伺服器組裝市場、商用 GPU 系統市場〔CF10：AMD 2026 計 GB300 世代、2027 起計 Rubin 世代〕、AI 電源市場）、ShareTbl 的 AI 電源份額表、Checks K6-pwr、K7、Interface SIF_AsmMkt_*、SIF_GPUSysMkt_*、SIF_PwrMkt_*、SIF_PwrShareSum_*。", False),
        ("v0.5.2：新增 M16（每 GW AI 高速 PCB 與 CCL 金額；原在 victory_giant、emc）與 PGW 頁、Interface SIF_PCBperGW、SIF_CCLperGW、SIF_PCBMkt_*、SIF_CCLMkt_*。", False),
        ("v0.5.1：新增 M15（ABF 供需表輸入：產值、AI 占比、權重、各廠產能指數、味之素膜上限；欣興、南電、景碩產能與味之素膜上限原在各公司模型）與 ABF 頁（需求、供給、缺口三情境）、Interface SIF_ABF*。", False),
        ("v0.5：新增 M13（加速器 ABF 載板份額、每顆金額路徑、供需缺口；原在 ibiden、unimicron、nanya_pcb、kinsus）、M14（CoWoS 每片封裝數、其他用途、台積電產能、委外補充；原在 tsmc），ShareTbl 的 ABF 載板份額表、Checks K6-abf、Interface SIF_ABFShareSum_*。", False),
        ("v0.4.1：新增 M12（CPO 光學金額份額：Lumentum、Coherent、中際旭創、新易盛、Fabrinet；原在各公司模型）與 ShareTbl 的 CPO 份額表、Checks K6-cpo、Interface SIF_CPOShareSum。", False),
        ("v0.4：新增 M8–M11（每顆網通內容、每顆 IT capex〔NVIDIA 世代改由 Tokenomics 推導〕、網通技術轉換與 CPO 三情境、2026 Rubin 占比 vr26；原在 markets/network 與 11 個公司模型）與 Deploy 頁（AI IT 資本支出、部署 GW 三情境）；橋接表改讀 Tokenomics v21.7。", False),
        ("讀法：Scen（情境定義）→ In（藍字輸入）→ Path／Mkt（推導）→ Interface（下游取數）→ Checks（SGOV_Errors 須為 0）。", False),
        ("顏色：藍字＝輸入（本活頁簿擁有）；黑字＝公式；綠字＝跨頁或外部連結；淡黃底＝關鍵輸出；亮黃底＝待 Andy 決定或未定。", False),
        ("標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy 與 Assumed 一律給區間（v0.1 搬入列尚有缺區間者，見 Checks G3）。", False),
        ("Excel 優先：要改輸入，直接改 In 頁藍字；builder 重建時讀回（--base）。公式頁不要手改（重建時會被覆寫）。", False),
        ("取數規則（下游）：只取 SIF_ 具名範圍（含 SIF_TK_* 鏡像）；不得連結 In、Mkt 等內部頁；以 tools/export_snapshot.py 匯成版本固定快照，版本號寫入公司模型。", False),
        ("版本命名：YYYYMMDD_Siliconomics_vX.Y.xlsx；model/CURRENT 記錄現行檔名。只改證據或標記為修補版；輸入值或公式改變為次版；Interface 名稱刪除或定義改變為主版（v1.0 前只新增名稱以次版處理）。", False),
    ]
    for i, (t, b) in enumerate(lines, 1):
        c = RD.cell(i, 1, t); c.alignment = WRAP
        if b: c.font = Font(bold=True, size=13)
    RD.column_dimensions["A"].width = 150

    wb.save(a.out)
    print(f"寫出 {a.out}：In {len(INROW)} 鍵、Interface {len(IFNAMES)} 名稱、衝突 {len(conflicts)} 項")


if __name__ == "__main__":
    main()
