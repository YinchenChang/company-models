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
TK_SNAPSHOT = os.path.join(SIL, "data", "tokenomics_snapshot_v21.10.json")   # v0.16：Tokenomics v21.10（相對 v21.7 只有 NonNV 產出比 TPU、Trainium 改值；本庫不讀產出比）
TK_OLD = os.path.join(SIL, "data", "tokenomics_snapshot_v21.7.json")   # 下游（chip_bridge）實際讀的 TK 快照（仍 v21.7；差異 6 格皆為產出比，chip_bridge 不讀，見 Checks K3）
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


def kc(y, s_="B"):
    """基準鏈上的顆數係數乘數（v0.19-trial：有協調區塊時各年讀 crd_k_<yy>；否則沿用 2028–29 × um*_B）。"""
    if globals().get("COORD"):
        return f"*crd_k_{y % 100}"
    return f"*um{y % 100}_{s_}" if y >= 2028 else ""


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
    # v0.16（決策 S41）：本庫只給基準（＝原情境 B）；產業情境 A、C 的參考值移到市場層 markets/scen（In 頁不再有 A／C 列，各頁只算基準）
    sys.path.insert(0, os.path.join(ROOT, "markets", "scen"))
    import scen as SCN
    SCL = ["B"]
    globals()["SCL"] = SCL
    base = load_base_inputs(a.base)
    tk = json.load(open(TK_SNAPSHOT, encoding="utf-8"))
    tk_old = json.load(open(TK_OLD, encoding="utf-8")) if os.path.exists(TK_OLD) else None
    conflicts = json.load(open(CONFLICTS, encoding="utf-8")) if os.path.exists(CONFLICTS) else []
    decisions = []
    for f in sorted(glob.glob(os.path.join(HERE, "seed", "decisions_v*.json")), key=_vkey):
        decisions += json.load(open(f, encoding="utf-8"))
    # 新增的輸入列（各版 rows_added_v*.json）：接在指定節的最後一列之後，成為新的一節；base Excel 有值時以 Excel 為準
    seed_rows = [rw for rw in seed["rows"] if "id" not in rw or rw["id"] not in SCN._OWN]
    for f in sorted(glob.glob(os.path.join(HERE, "seed", "rows_added_v*.json")), key=_vkey):
        blocks = json.load(open(f, encoding="utf-8"))
        for ad in (blocks if isinstance(blocks, list) else [blocks]):
            idx = [i for i, rw in enumerate(seed_rows) if "sec" in rw]
            start = next(i for i in idx if seed_rows[i]["sec"].startswith(ad["after_section_end"]))
            nxt = next((i for i in idx if i > start), len(seed_rows))
            seed_rows[nxt:nxt] = [{"sec": ad["section"]}] + [rw for rw in ad["rows"] if rw["id"] not in SCN._OWN]

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
    # v0.19-trial（三層基準協調規則 v1.1 C3）：基準鏈上的顆數係數改讀 Coord 頁的「依規則決定的基準係數」crd_k_<yy>
    #   ＝IF(上游定基準且覆蓋已補齊, 換算值 ÷ 原始顆數讀數, 現行係數〔2026–27＝1、2028–29＝um28_B／um29_B〕)。
    #   覆蓋未補齊（crd_cov_done＝0）時各年等於現行係數，數值與 v0.18 逐格相同。
    COORD = "crd_cov_done" in INROW
    globals()["COORD"] = COORD
    for col, w in zip("ABCDEFGHIJ", (14, 52, 8, 9, 10, 9, 18, 70, 18, 50)):
        IN.column_dimensions[col].width = w

    # ---------------- Path、Mkt：記憶體價格路徑與 HBM 市場（公式與 markets/memory/market.py write() 同一份） ----------------
    import market as MKT
    ref = MKT.write(wb, hdr, Font, scen=SCL)
    MK = ref["mk"]
    if COORD:   # v0.19-trial：Mkt 頁顆數係數列（market.py 寫 1／1／um28_B／um29_B）改讀依規則決定的基準係數；覆蓋未補齊時值相同
        for j, yy in enumerate((26, 27, 28, 29)):
            c = wb["Mkt"].cell(MK["cfB"], 4 + j, f"=crd_k_{yy}"); c.number_format = "0.00"
        wb["Mkt"].cell(MK["cfB"], 1).value = str(wb["Mkt"].cell(MK["cfB"], 1).value) + "〔v0.19-trial：讀 Coord 頁 crd_k，依三層基準協調規則決定〕"
    wb["Mkt"]["A2"] = "公式由 markets/memory/market.py write() 產生（與各記憶體公司模型的 Mkt 頁同一份）；輸入讀本活頁簿 In 頁"

    # ---------------- Scen：情境定義 ----------------
    SC["A1"] = ("Scen — v0.16 起本庫只給基準（工作方式 v1.7：三層只給基準、情境在個股；Andy 2026-11-02 15:58）。基準＝原情境 B（平平；決策 S33、S35）。"
                "產業情境 A（順利）、C（不如預期）的參考值已移到市場層 markets/scen（決策 S41；自 v0.15 In 頁逐字搬出），In 頁與各推導頁不再有 A／C；"
                "個股 A／C 預設讀市場層同字母情境（Andy 2026-10-25「各股情境與產業情境一致」），覆寫時寫明相對市場層的倍數與產業 GW 相對 Wattonomics 基準 ±%")
    hdr(SC, 3, ["情境", "名稱", "定義（顆數與部署）", "記憶體價格路徑", "2028 顆數 ÷ 原始顆數", "2029 顆數 ÷ 原始顆數", "角色"])
    NM = {"B": "平平＝基準"}
    DEF = {"B": "AI 資本支出消化：2028 顆數約為原始顆數 0.8、2029 約 0.85"}
    for i, s in enumerate(SCL):
        rr = 4 + i
        SC.cell(rr, 1, s); SC.cell(rr, 2, NM[s]); SC.cell(rr, 3, DEF[s]); SC.cell(rr, 4, seed["scen_desc"][s])
        for j, y in ((5, 28), (6, 29)):
            c = SC.cell(rr, j, f"=um{y}_{s}"); c.font = GREEN; c.number_format = "0.00"
        SC.cell(rr, 7, "基準（下游讀 *_Base）")
    SC["A9"] = "說明"
    SC["A9"].font = BOLD
    SC["A10"] = ("情境係數只定義 2028–29（Assumed，無外部預測可校準；區間見 In 頁 M2b）。基準係數 2028 0.8、2029 0.85，"
                 "即本庫基準＝原始顆數（Units 頁未乘係數者）× 基準係數。部署 GW 權威在 Wattonomics（第 13 條）；兩邊差距見 Recon 頁 B 節。")
    SC["A11"] = (f"市場層情境參考值：{SCN.SRC}，{len(SCN.OWN)} 列（記憶體價格路徑 cqA／cqC、nqA／nqC，顆數係數 um28／um29 A、C，HBM 單價 hbm_p28／29 A、C，"
                 "世芯 Trainium 份額 al27／al28 A、C，CPO 滲透率 cpo A、C，設備係數 eq_ck C）。A：" + SCN.SCEN_DESC["A"] + "；C：" + SCN.SCEN_DESC["C"])
    for col, w in zip("ABCDEFG", (6, 10, 46, 60, 14, 14, 50)):
        SC.column_dimensions[col].width = w

    # ---------------- Units：平台別顆數（2026–2029；2029＝2028 ×（1＋平台別成長率）） ----------------
    UN = None; UROW = {}
    if "g29_bw" in INROW:
        UN = wb.create_sheet("Units", index=wb.sheetnames.index("Path"))
        UN["A1"] = "Units — 加速器平台別顆數（百萬顆；原始顆數＝舊情境 A，未乘係數；v0.12 起本庫基準＝合計（情境 B）列）。2026–2028＝In 頁 M2；2029＝2028 ×（1＋In 頁 M5 平台別成長率）。情境 B、C 再乘 Scen 頁係數（2028–29）"
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
        UN.cell(r, 1, "合計（原始顆數＝舊情境 A，未乘係數）").font = BOLD
        for j in range(3, 7):
            c = UN.cell(r, j, f"=SUM({'CDEF'[j-3]}{r0}:{'CDEF'[j-3]}{r1})"); c.number_format = "0.00"; c.fill = KEY
        UN.cell(r, 7, f"=F{r}/E{r}-1").number_format = "0.0%"; UROW["tot"] = r; r += 1
        for s_ in SCL:
            UN.cell(r, 1, f"合計（情境 {s_}）" + ("＝本庫基準" if s_ == "B" else "〔停用過渡〕"))
            if COORD:
                for j, yy in enumerate((26, 27, 28, 29)):
                    UN.cell(r, 3 + j, f"={'CDEF'[j]}{UROW['tot']}*crd_k_{yy}")
            else:
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
        for s_ in SCL:
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
            for s_ in SCL:
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
            if "lc_avsh_26" in INROW:
                ST.cell(r, 1, "AI 液冷有效份額（欄 D–G＝2026／2027／2028／2029；分母＝Rack 頁 AI 液冷市場〔機架內＋分歧管〕；奇鋐相對值不隨情境，雙鴻與健策 A 高、B 基準、C 低）").font = BOLD; r += 1
                for s_ in SCL:
                    suf = {"A": "_hi", "B": "", "C": "_lo"}[s_]
                    r_first = r
                    for who, f26, fy, note in (("奇鋐", "=lc_avsh_26", "*lc_avrel_{y}", "lc_avsh_26 × lc_avrel_27／28／29（不隨情境）"),
                                                ("雙鴻", "=lc_aush_26", "*lc_avrel_{y}" + suf, f"lc_aush_26 × lc_avrel_27／28／29{suf}"),
                                                ("健策（MCL／MLCP 冷板元件）", "=0", None, f"lc_jtsh_27／28／29{suf}（2026 含在其他散熱）")):
                        ST.cell(r, 1, f"AI 液冷（情境 {s_}）"); ST.cell(r, 2, who); ST.cell(r, 3, note)
                        for j, y in enumerate((26, 27, 28, 29)):
                            f = f26 if y == 26 else (f26 + fy.format(y=y) if fy else f"=lc_jtsh_{y}{suf}")
                            ST.cell(r, 4 + j, f).number_format = "0.0%"
                        r += 1
                    ST.cell(r, 1, f"AI 液冷（情境 {s_}）"); ST.cell(r, 2, "合計").font = BOLD
                    for j in range(4):
                        col = "DEFG"[j]
                        c = ST.cell(r, 4 + j, f"=SUM({col}{r_first}:{col}{r-1})"); c.number_format = "0.0%"; c.fill = KEY
                    STROW["sum_lc" + s_] = r; r += 1
                    ST.cell(r, 1, f"AI 液冷（情境 {s_}）"); ST.cell(r, 2, "餘額（酷碼、富世達、Boyd、CoolIT、Vertiv 等，未建模）")
                    for j in range(4):
                        col = "DEFG"[j]
                        ST.cell(r, 4 + j, f"=1-{col}{r-1}").number_format = "0.0%"
                    r += 2
                ST.cell(r, 1, "註：台達電液冷營收多在機架外（CDU、Liquid-to-Air），不屬本市場，不入表（CF12）；奇鋐、雙鴻分子含部分 CDU，份額可能高估"); r += 2
            ST.cell(r, 1, "對照：凱基台達電份額（只作對照，不入表；CF11 Andy 11-05 選甲）"); ST.cell(r, 2, "2025-07 Blackwell PSU／2025-09 AI 伺服器電源")
            ST.cell(r, 4, "=pwr_kgi_0701").number_format = "0.0%"; ST.cell(r, 5, "=pwr_kgi_0926").number_format = "0.0%"; r += 2
            ST.column_dimensions["G"].width = 9
        if "hbmsh_skh_25" in INROW:
            ST.cell(r, 1, "HBM 營收份額（欄 D–H＝2025／2026／2027／2028／2029；In 頁 M24；份額不隨情境；餘額＝長鑫等其他）").font = BOLD; r += 1
            r_first = r
            for who, nm_ in (("SK 海力士", "skh"), ("三星", "ss"), ("美光", "mu")):
                ST.cell(r, 1, "HBM"); ST.cell(r, 2, who); ST.cell(r, 3, f"hbmsh_{nm_}_25…29")
                for j, y in enumerate((25, 26, 27, 28, 29)):
                    ST.cell(r, 4 + j, f"=hbmsh_{nm_}_{y}").number_format = "0.0%"
                r += 1
            ST.cell(r, 1, "HBM"); ST.cell(r, 2, "合計").font = BOLD
            for j in range(5):
                col = "DEFGH"[j]
                c = ST.cell(r, 4 + j, f"=SUM({col}{r_first}:{col}{r-1})"); c.number_format = "0.0%"; c.fill = KEY
            STROW["sum_hbm"] = r; r += 1
            ST.cell(r, 1, "HBM"); ST.cell(r, 2, "餘額（長鑫等，未建模）")
            for j in range(5):
                col = "DEFGH"[j]
                ST.cell(r, 4 + j, f"=1-{col}{r-1}").number_format = "0.0%"
            r += 2
            ST.column_dimensions["H"].width = 9
        if "asm_qush_26" in INROW:
            KS = ("bw", "rb", "ru", "hp", "amd", "tpu", "trn", "oth")
            def un(k, y): return f"u_{k}_{y}" if y <= 28 else f"u_{k}_28*(1+g29_{k})"
            def usum(ks, y): return "(" + "+".join(un(k, y) for k in ks) + ")"
            ST.cell(r, 1, "AI 伺服器組裝有效份額（欄 D–G＝2026／2027／2028／2029；分母＝Rack 頁 AI 伺服器組裝市場〔ODM 價〕；2026＝公司模型校準快照，之後 × 份額相對值；不隨情境）").font = BOLD; r += 1
            ST.cell(r, 1, "AI 伺服器組裝"); ST.cell(r, 2, "緯穎 ASIC 傾斜係數（x）"); ST.cell(r, 3, "asm_wyam ×（Trainium＋其他 ASIC 顆數指數 ÷ 全部顆數指數）＋（1 − asm_wyam）")
            for j, y in enumerate((26, 27, 28, 29)):
                f = "=1" if y == 26 else f"=asm_wyam*({usum(('trn', 'oth'), y)}/{usum(('trn', 'oth'), 26)})/({usum(KS, y)}/{usum(KS, 26)})+(1-asm_wyam)"
                ST.cell(r, 4 + j, f).number_format = "0.000"
            STROW["asm_tilt"] = r; r += 1
            r_first = r
            for who, k in (("廣達", "qu"), ("鴻海", "fx"), ("緯創本體", "wi"), ("緯穎（× ASIC 傾斜）", "wy")):
                ST.cell(r, 1, "AI 伺服器組裝"); ST.cell(r, 2, who); ST.cell(r, 3, f"asm_{k}sh_26 × asm_{k}rel_27／28／29" + (" × 傾斜係數" if k == "wy" else ""))
                for j, y in enumerate((26, 27, 28, 29)):
                    f = f"=asm_{k}sh_26" + ("" if y == 26 else f"*asm_{k}rel_{y}") + (f"*{'DEFG'[j]}{STROW['asm_tilt']}" if k == "wy" and y > 26 else "")
                    ST.cell(r, 4 + j, f).number_format = "0.0%"
                r += 1
            ST.cell(r, 1, "AI 伺服器組裝"); ST.cell(r, 2, "合計").font = BOLD
            for j in range(4):
                col = "DEFG"[j]
                c = ST.cell(r, 4 + j, f"=SUM({col}{r_first}:{col}{r-1})"); c.number_format = "0.0%"; c.fill = KEY
            STROW["sum_asm"] = r; r += 1
            ST.cell(r, 1, "AI 伺服器組裝"); ST.cell(r, 2, "餘額（英業達、Celestica、Dell、美超微、HPE 等，未建模）")
            for j in range(4):
                col = "DEFG"[j]
                ST.cell(r, 4 + j, f"=1-{col}{r-1}").number_format = "0.0%"
            r += 1
            ST.cell(r, 1, "註：Dell、美超微以整櫃售價計（Rack 頁 B 節商用 GPU 系統市場），分母不同，不入本表；緯創模型自算的緯穎份額路徑與本表不同（CF23）"); r += 2
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
        for s_ in SCL:
            DP.cell(r, 1, "AI IT 資本支出 × 情境顆數係數"); DP.cell(r, 2, s_)
            for y, col in COLS.items():
                f = f"={col}{DROW['itc']}" + kc(y, s_)
                DP.cell(r, list(COLS).index(y) + 3, f).number_format = "0.0"
            DROW["itc" + s_] = r; r += 1
        DP.cell(r, 1, "每 GW IT 資本支出（$B／GW）")
        DP.cell(r, 3, "=TK_CapexIT_GB300*(1-vr26)+TK_CapexIT_VR200*vr26").number_format = "0.00"
        for y in (2027, 2028, 2029):
            DP.cell(r, list(COLS).index(y) + 3, "=TK_CapexIT_VR200").number_format = "0.00"
        DROW["itpg"] = r; r += 1
        for s_ in SCL:
            DP.cell(r, 1, "部署 GW（IT 電力）").font = BOLD; DP.cell(r, 2, s_)
            for y, col in COLS.items():
                c = DP.cell(r, list(COLS).index(y) + 3, f"={col}{DROW['itc' + s_]}/{col}{DROW['itpg']}"); c.number_format = "0.0"; c.fill = KEY
            DROW["gw" + s_] = r; r += 1
        # v0.16.3：實體部署 GW（Andy 2026-11-05「先加」；電源、散熱、組裝等按實體電力計的下游用）。
        # 成本等值 GW（上列）以 VR200 每 GW IT capex 當分母：Rubin Ultra（每 GW 成本較高）被高估、ASIC（較便宜）被低估。
        # NVIDIA＝顆數 ÷ 該世代 Tokenomics 每 GW 顆數；非 NVIDIA＝IT 資本支出 ÷（VR200 每 GW IT capex × Tokenomics 持有成本比），與 R4 橋接 ① 同法。
        NVG = {"bw": "GB300", "rb": "VR200", "ru": "RU", "hp": "H100"}
        HOLD = {"amd": "TK_NonNV_Hold_AMD", "tpu": "TK_NonNV_Hold_TPU", "trn": "TK_NonNV_Hold_Trn",
                "oth": "((TK_NonNV_Hold_Maia+TK_NonNV_Hold_Ascend+TK_NonNV_Hold_OAI)/3)"}
        def ucoef(y):
            return kc(y, "B")
        DP.cell(r, 1, "實體部署 GW：NVIDIA（Σ 顆數 ÷ 該世代每 GW 顆數）"); DP.cell(r, 2, "B")
        for y, col in COLS.items():
            f = "=(" + "+".join(f"Units!{col}{UROW[k]}/TK_GPUsPerGW_{gname}" for k, gname in NVG.items() if k in UROW) + f")*1E6{ucoef(y)}"
            DP.cell(r, list(COLS).index(y) + 3, f).number_format = "0.0"
        DROW["phnv"] = r; r += 1
        DP.cell(r, 1, "實體部署 GW：非 NVIDIA（IT 資本支出 ÷（VR200 每 GW × 持有成本比））"); DP.cell(r, 2, "B")
        for y, col in COLS.items():
            f = "=(" + "+".join(f"Units!{col}{UROW[k]}*it_{plat(k, y)}/(TK_CapexIT_VR200*{h})" for k, h in HOLD.items() if k in UROW) + f"){ucoef(y)}"
            DP.cell(r, list(COLS).index(y) + 3, f).number_format = "0.0"
        DROW["phno"] = r; r += 1
        DP.cell(r, 1, "實體部署 GW（IT 電力；電源、散熱、組裝用）").font = BOLD; DP.cell(r, 2, "B")
        for y, col in COLS.items():
            c = DP.cell(r, list(COLS).index(y) + 3, f"={col}{DROW['phnv']}+{col}{DROW['phno']}"); c.number_format = "0.0"; c.fill = KEY
        DROW["gwphB"] = r; r += 1
        DP.cell(r, 1, "實體 ÷ 成本等值")
        for y, col in COLS.items():
            DP.cell(r, list(COLS).index(y) + 3, f"={col}{DROW['gwphB']}/{col}{DROW['gwB']}").number_format = "0.000"
        DROW["phr"] = r; r += 1
        DP.column_dimensions["A"].width = 36
        for col in "BCDEF": DP.column_dimensions[col].width = 10

    # ---------------- Coord：三層基準協調（v0.19-trial；規則 docs/plan/三層基準協調規則.md v1.1 C3、C4、C7） ----------------
    CO = None; COROW = {}
    if COORD and DP is not None:
        CO = wb.create_sheet("Coord", index=wb.sheetnames.index("Path"))
        CO["A1"] = ("Coord — 三層基準協調（試作 v0.19-trial；規則 v1.1 C3）：每年依兩層證據比例判定狀態（甲：兩層都有證據；乙：一層；丙：兩層都外推；含緩衝區），"
                    "區間跨線即待定（規則 v1.3：兩層證據比例低／基準／高任一端點與基準值落在門檻兩側 → 待定，維持現行、標對帳中）。決定由誰定基準。上游定基準（丙，或 Wattonomics 有證據的乙）且 Wattonomics 覆蓋已補齊（crd_cov_done＝1）→ 基準＝Wattonomics 換算值；"
                    "覆蓋未補齊 → 維持現行（原始顆數 × S33 係數）並標「對帳中」；其餘維持現行。基準係數 crd_k＝基準 ÷ 原始顆數讀數，取代基準鏈上的 S33 係數")
        CO["A2"] = ("S33 係數（um28_B、um29_B）不再是基準鏈上的水準係數：只在『維持現行』的年份經 crd_k 帶入；改用換算值的年份不讀。"
                    "供給側讀數＝原始顆數（不乘係數）；需求側讀數＝Wattonomics 換算值（C7 兩條明列數列）；另一側讀數＝未被採為基準的那一條（C4 兩讀數目標價用）")
        hdr(CO, 3, ["列", "項目", "2026", "2027", "2028", "2029", "說明"])
        CC_ = {2026: "C", 2027: "D", 2028: "E", 2029: "F"}
        r = 4
        def co(key, lab, f, fmt="0.00", note="", fill=None, bold=False):
            nonlocal_r = COROW.setdefault("_r", 4)
            CO.cell(nonlocal_r, 1, key); CO.cell(nonlocal_r, 2, lab)
            if bold: CO.cell(nonlocal_r, 2).font = BOLD
            for y, col in CC_.items():
                c = CO.cell(nonlocal_r, list(CC_).index(y) + 3, f(y, col, y % 100)); c.number_format = fmt
                if fill: c.fill = fill
            CO.cell(nonlocal_r, 7, note)
            COROW[key] = nonlocal_r; COROW["_r"] = nonlocal_r + 1
        R_ = lambda k, col: f"{col}{COROW[k]}"
        CO.cell(COROW.setdefault("_r", 4), 1, "A. 兩個讀數").font = BOLD; COROW["_r"] += 1
        co("sup", "供給側讀數：原始顆數部署 GW（不乘係數；成本等值、出貨年）", lambda y, col, yy: f"=Deploy!{col}{DROW['itc']}/Deploy!{col}{DROW['itpg']}", bold=True,
           note="Deploy 頁 AI IT 資本支出（未乘係數）÷ 每 GW IT 資本支出；2027 已含 v0.18 HBM 截頂（u_*_27 × 0.9404）；個股情境 A（J5）")
        co("s33", "現行係數（S33：2026–27＝1，2028–29＝um28_B／um29_B）", lambda y, col, yy: "=1" if y <= 2027 else f"=um{yy}_B", note="決策 S33；只作『維持現行』年份的係數與情境參照")
        co("cur", "現行基準（原始 × 現行係數＝v0.18 取法）", lambda y, col, yy: f"={R_('sup', col)}*{R_('s33', col)}")
        co("dem", "需求側讀數：Wattonomics 換算值（平滑法 x2；暫值）", lambda y, col, yy: f"=crd_x2_{yy}", bold=True, note="In 頁 M30 crd_x2；待 Wattonomics v0.3.2 後以 baseline_bridge.py 重算")
        COROW["_r"] += 1
        CO.cell(COROW["_r"], 1, "B. 證據狀態（C3；只看狀態、不看數值高低）").font = BOLD; COROW["_r"] += 1
        co("ew", "Wattonomics 證據比例（WIF_Evidence 口徑）", lambda y, col, yy: f"=crd_ew_{yy}", "0.0%")
        co("es", "Siliconomics 證據比例", lambda y, col, yy: f"=crd_es_{yy}", "0.0%")
        # v1.3「區間跨線即待定」：兩層證據比例都用低／基準／高三值判；任一層低或高的「有證據」與基準值不同 → 待定
        co("ew_lo", "　Wattonomics 證據比例 低", lambda y, col, yy: f"=crd_ew_{yy}_lo", "0.0%", note="模型只有點值時低＝基準＝高（In 頁說明欄註明）")
        co("ew_hi", "　Wattonomics 證據比例 高", lambda y, col, yy: f"=crd_ew_{yy}_hi", "0.0%")
        co("es_lo", "　Siliconomics 證據比例 低", lambda y, col, yy: f"=crd_es_{yy}_lo", "0.0%", note="docs/reports/20261114_供給側證據比例.xlsx 區間")
        co("es_hi", "　Siliconomics 證據比例 高", lambda y, col, yy: f"=crd_es_{yy}_hi", "0.0%")
        hasev = lambda sh, pv: (lambda y, col, yy: f"=IF(ISBLANK(crd_{pv}_{yy}),IF({R_(sh, col)}>=crd_thr,1,0),IF(crd_{pv}_{yy}=1,IF({R_(sh, col)}>=crd_exit,1,0),IF({R_(sh, col)}>=crd_enter,1,0)))")
        co("hw", "Wattonomics 有證據（1／0；首次以門檻、之後套緩衝區）", hasev("ew", "prevw"), "0", note="首次判定（上次判定空白）：≥ crd_thr；原有證據：≥ crd_exit 才保留；原無證據：≥ crd_enter 才進入")
        co("hs", "Siliconomics 有證據（1／0）", hasev("es", "prevs"), "0")
        co("hw_lo", "　Wattonomics 有證據（低）", hasev("ew_lo", "prevw"), "0")
        co("hw_hi", "　Wattonomics 有證據（高）", hasev("ew_hi", "prevw"), "0")
        co("hs_lo", "　Siliconomics 有證據（低）", hasev("es_lo", "prevs"), "0")
        co("hs_hi", "　Siliconomics 有證據（高）", hasev("es_hi", "prevs"), "0")
        co("st0", "狀態（只看基準值；甲／乙／丙）", lambda y, col, yy: f'=IF(AND({R_("hw", col)}=1,{R_("hs", col)}=1),"甲",IF(OR({R_("hw", col)}=1,{R_("hs", col)}=1),"乙","丙"))', "@")
        co("pend", "區間跨線即待定（1／0；規則 v1.3）", lambda y, col, yy: (f'=IF(OR({R_("hw_lo", col)}<>{R_("hw", col)},{R_("hw_hi", col)}<>{R_("hw", col)},'
                                                                      f'{R_("hs_lo", col)}<>{R_("hs", col)},{R_("hs_hi", col)}<>{R_("hs", col)}),1,0)'), "0",
           note="任一層的區間端點與基準值落在門檻（含緩衝區）兩側 → 狀態不同 → 待定：維持現行基準、標對帳中、個股並列兩讀數，下次重判再定")
        co("st", "狀態（甲／乙／丙／待定）", lambda y, col, yy: f'=IF({R_("pend", col)}=1,"待定",{R_("st0", col)})', "@", fill=KEY, bold=True)
        co("lead", "由誰定基準", lambda y, col, yy: f'=IF({R_("st", col)}="待定","待定（維持現行；下次重判）",IF({R_("st", col)}="甲","兩層對帳（證據等級高者）",IF({R_("st", col)}="丙","Wattonomics（上游約束）",IF({R_("hw", col)}=1,"Wattonomics","Siliconomics"))))', "@")
        co("up", "上游定基準（丙，或 Wattonomics 有證據的乙；待定＝0；1／0）", lambda y, col, yy: f'=IF(OR({R_("st", col)}="丙",AND({R_("st", col)}="乙",{R_("hw", col)}=1)),1,0)', "0")
        co("res", "甲狀態殘差（現行 ÷ 換算值 − 1；容差 rc_tol）", lambda y, col, yy: f"={R_('cur', col)}/{R_('dem', col)}-1", "+0.0%;-0.0%", note="只在甲狀態用；乙、丙狀態只供參考")
        COROW["_r"] += 1
        CO.cell(COROW["_r"], 1, "C. 依規則的基準").font = BOLD; COROW["_r"] += 1
        co("use", "改用換算值（上游定基準且覆蓋已補齊；1／0）", lambda y, col, yy: f"=IF(AND({R_('up', col)}=1,crd_cov_done=1),1,0)", "0")
        co("flag", "對帳中（1／0）", lambda y, col, yy: f'=IF(OR({R_("st", col)}="待定",AND({R_("up", col)}=1,crd_cov_done<>1),AND({R_("st", col)}="甲",ABS({R_("res", col)})>rc_tol)),1,0)', "0",
           note="待定（v1.3）；上游定基準但覆蓋未補齊（C3 乙前提）；或甲狀態殘差超出容差（待診斷）")
        PV = {"D": "C", "E": "D", "F": "E"}   # 前一年欄
        co("chain", "連續法（1／0；規則 v1.4）：上游定基準且前一年為待定、甲且維持現行，或前一年本身用連續法",
           lambda y, col, yy: "=0" if col == "C" else
           (f'=IF(AND({R_("use", col)}=1,OR({R_("st", PV[col])}="待定",AND({R_("st", PV[col])}="甲",{R_("use", PV[col])}=0),{PV[col]}{COROW["_r"]}=1)),1,0)'),
           "0", note="水準等前一年的待定解決，形狀由上游決定：基準＝前一年基準 ×（該年換算值 ÷ 前一年換算值）；換算值本身列為另一側讀數。前一年也直接用換算值時不適用")
        co("base", "基準部署 GW（依規則）", lambda y, col, yy: (f"=IF({R_('use', col)}=1,{R_('dem', col)},{R_('cur', col)})" if col == "C" else
           f"=IF({R_('use', col)}=1,IF({R_('chain', col)}=1,{PV[col]}{COROW['_r']}*{R_('dem', col)}/{R_('dem', PV[col])},{R_('dem', col)}),{R_('cur', col)})"), fill=KEY, bold=True)
        co("lab", "該年標示", lambda y, col, yy: f'=IF({R_("use", col)}=1,IF({R_("chain", col)}=1,"連續法（換算值成長率）","改用換算值"),IF({R_("st", col)}="待定","待定（對帳中）",IF({R_("flag", col)}=1,"對帳中",IF({R_("st", col)}="甲","一致","維持（本層有證據）"))))', "@", fill=KEY)
        co("k", "基準係數 crd_k（＝基準 ÷ 供給側讀數；維持現行的年份＝現行係數）", lambda y, col, yy: f"=IF({R_('use', col)}=1,{R_('base', col)}/{R_('sup', col)},{R_('s33', col)})", "0.0000",
           note="基準鏈（Units、Deploy、Mkt、ABF、Rack、Equip、Cap）各年乘本列；Interface SIF_UnitsCoef_Base 隨之")
        for y, col in CC_.items():
            name(wb, f"crd_k_{y % 100}", f"Coord!${col}${COROW['k']}")
        co("chk", "檢查：Deploy 頁基準部署 GW − 本頁基準（應為 0）", lambda y, col, yy: f"=Deploy!{col}{DROW['gwB']}-{R_('base', col)}", "0.000000")
        co("alt", "另一側讀數（未被採為基準的一條；C4 兩讀數目標價用；連續法年份＝換算值本身）", lambda y, col, yy: f"=IF({R_('use', col)}=1,IF({R_('chain', col)}=1,{R_('dem', col)},{R_('sup', col)}),{R_('dem', col)})")
        co("altr", "另一側讀數 ÷ 基準", lambda y, col, yy: f"={R_('alt', col)}/{R_('base', col)}", "0.000", note="個股以基準算出的產業量 × 本比值＝另一側讀數的產業量（近似；各平台等比例）")
        co("c7", "C7 供給側 ÷ 需求側 − 1（認知落差假說；逐季記錄哪一側收斂）", lambda y, col, yy: f"={R_('sup', col)}/{R_('dem', col)}-1", "+0.0%;-0.0%")
        co("chg", "依規則基準 ÷ 現行 − 1", lambda y, col, yy: f"={R_('base', col)}/{R_('cur', col)}-1", "+0.0%;-0.0%")
        CO.column_dimensions["A"].width = 6; CO.column_dimensions["B"].width = 62; CO.column_dimensions["G"].width = 80
        for col in "CDEF": CO.column_dimensions[col].width = 12

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
        for s_ in SCL:
            AB.cell(r, 1, "加速器有效顆數（NVIDIA ＋ 非 NVIDIA × v_ot；百萬顆）"); AB.cell(r, 2, s_)
            for y, col in COLS.items():
                m = kc(y, s_)
                cell(r, y, "=(" + "+".join(f"Units!{col}{UROW[k]}" for k in NVK) + "+abf_vot*(" + "+".join(f"Units!{col}{UROW[k]}" for k in OTK) + f")){m}", "0.00")
            ABROW["equ" + s_] = r; r += 1
        for s_ in SCL:
            AB.cell(r, 1, "需求指數：加速器（× 面積層數）"); AB.cell(r, 2, s_)
            for y, col in COLS.items():
                cell(r, y, f"={col}{ABROW['equ' + s_]}/$C{ABROW['equ' + s_]}*(1+abf_gasp)^{y - 2026}")
            ABROW["dacc" + s_] = r; r += 1
        for s_ in SCL:
            AB.cell(r, 1, "需求指數：其他 AI（AI IT 資本支出成長 ^ 彈性）"); AB.cell(r, 2, s_)
            for y, col in COLS.items():
                cell(r, y, f"=(Deploy!{col}{DROW['itc' + s_]}/Deploy!$C{DROW['itc' + s_]})^abf_oai_el")
            ABROW["doai" + s_] = r; r += 1
        AB.cell(r, 1, "需求指數：非 AI")
        for y in COLS: cell(r, y, f"=(1+abf_gnai)^{y - 2026}")
        ABROW["dnai"] = r; r += 1
        for s_ in SCL:
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
        for s_ in SCL:
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
            for s_ in SCL:
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
        for s_ in SCL:
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
        for s_ in SCL:
            RK.cell(r, 1, "商用 GPU 系統市場（US$B）").font = BOLD; RK.cell(r, 2, s_)
            for y, col in COLS.items():
                m = kc(y, s_)
                rc(r, y, f"=({col}{RKROW['ugb']}*{col}{RKROW['vgb']}+{col}{RKROW['uvr']}*{col}{RKROW['vvr']}){m}", "#,##0.0", KEY)
            RKROW["sys" + s_] = r; r += 1
        r += 1
        RK.cell(r, 1, "C. AI 電源市場（delta、liteon 同讀）").font = BOLD; r += 1
        RK.cell(r, 1, "AI 電源每 kW 內容（US$/kW）")
        for y in COLS: rc(r, y, f"=pwr_pkw_{y % 100}", "#,##0")
        RK.cell(r, 7, "機架電源櫃、PSU、BBU、HVDC 電源櫃（In 頁 M17）"); RKROW["pkw"] = r; r += 1
        for s_ in SCL:
            RK.cell(r, 1, "AI 電源市場（US$B）").font = BOLD; RK.cell(r, 2, s_)
            for y, col in COLS.items(): rc(r, y, f"=Deploy!{col}{DROW['gw' + s_]}*{col}{RKROW['pkw']}/1000", "#,##0.00", KEY)
            RKROW["pwr" + s_] = r; r += 1
        RK.cell(r, 1, "對照：UBS 全球 AI 伺服器電源市場（US$B；口徑未找到）")
        rc(r, 2026, "=pwr_ubs_26", "#,##0.0"); rc(r, 2027, "=pwr_ubs_27", "#,##0.0"); RKROW["ubs"] = r; r += 1
        RK.cell(r, 1, "本庫情境 A ÷ UBS")
        for y, col in ((2026, "C"), (2027, "D")): rc(r, y, f"={col}{RKROW['pwr' + SCL[0]]}/{col}{RKROW['ubs']}", "0.00")
        RK.cell(r, 7, "UBS 高出約五到七成：口徑（是否含 BBU、HVDC、非 AI）未找到，只作對照"); r += 1
        if "lc_pen_26" in INROW:
            r += 1
            RK.cell(r, 1, "D. AI 液冷市場（機架內液冷＋分歧管；avc、auras、jentech 同讀；不含 CDU、Liquid-to-Air 等機架外冷卻）").font = BOLD; r += 1
            RK.cell(r, 1, "每 GW 液冷內容（US$M／GW）")
            rc(r, 2026, "=((1-vr26)*TK_RacksPerGW_GB300*(lc_cool_gb+lc_mani)+vr26*TK_RacksPerGW_VR200*(lc_cool_vr+lc_mani))/1000", "#,##0.0")
            for y in (2027, 2028, 2029): rc(r, y, f"=TK_RacksPerGW_VR200*(lc_cool_vr+lc_mani)*lc_adj_{y % 100}/1000", "#,##0.0")
            RK.cell(r, 7, "每架機架內液冷（Morgan Stanley BOM）＋分歧管，× 每 GW 架數；2027 起 × 內容調整（Rubin 全液冷、ASIC 液冷）"); RKROW["lcpg"] = r; r += 1
            for s_ in SCL:
                suf = {"A": "_hi", "B": "", "C": "_lo"}[s_]
                RK.cell(r, 1, "液冷滲透率（部署 GW 中液冷占比）"); RK.cell(r, 2, s_)
                for y in COLS: rc(r, y, f"=lc_pen_{y % 100}" + ("" if y == 2026 else suf), "0.0%")
                RKROW["lpen" + s_] = r; r += 1
            for s_ in SCL:
                RK.cell(r, 1, "AI 液冷市場（US$B）").font = BOLD; RK.cell(r, 2, s_)
                for y, col in COLS.items(): rc(r, y, f"=Deploy!{col}{DROW['gw' + s_]}*{col}{RKROW['lpen' + s_]}*{col}{RKROW['lcpg']}/1000", "#,##0.00", KEY)
                RKROW["lc" + s_] = r; r += 1
            RK.cell(r, 1, "對照：TrendForce AI 晶片液冷滲透率（顆數口徑）")
            rc(r, 2026, "=lc_tf_26", "0.0%"); rc(r, 2027, "=lc_tf_27", "0.0%")
            RK.cell(r, 7, "以 GW 計高於顆數計（液冷晶片功耗較高）"); r += 1
        RK.column_dimensions["A"].width = 62; RK.column_dimensions["G"].width = 60
        for col in "BCDEF": RK.column_dimensions[col].width = 10

    # ---------------- Equip：SEMI 設備路徑、SoC 測試機市場（v0.7） ----------------
    EQ = None; EQROW = {}
    if "eq_s_wfe_26" in INROW and "ts_mk_soc26" in INROW and UN is not None:
        EQ = wb.create_sheet("Equip", index=wb.sheetnames.index("Path"))
        EQ["A1"] = ("Equip — (1) SEMI 設備分項路徑（2026 新聞稿值；2025 由年增回推；2027、2028 年增連乘，測試與封裝 2027 為 2026、2028 的幾何內插；2029 × (1＋g29)）；"
                    "(2) SoC 測試機市場三情境＝AI（加速器顆數 × 每顆測試時間係數 × 每加權顆支出，以 CY2026 公司市場校準）＋非 AI；(3) 測試 ÷ WFE（CF15）。"
                    "記憶體測試機市場依 DRAM 設備支出（公司資本支出推導，markets/equipment），留在 advantest 模型")
        hdr(EQ, 3, ["項目", "情境", "2025", "2026", "2027", "2028", "2029", "說明"])
        YC = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G"}
        def ec(r, y, f, fmt="0.00", fill=None):
            c = EQ.cell(r, list(YC).index(y) + 3, f); c.number_format = fmt
            if fill: c.fill = fill
            return c
        r = 4
        EQ.cell(r, 1, "A. SEMI 設備分項（US$B；SEMI 為設備商協會，Interested-party）").font = BOLD; r += 1
        for k, nm in (("dram", "DRAM 設備"), ("nand", "NAND 設備"), ("fd", "代工與邏輯設備"), ("wfe", "晶圓前段設備（WFE）")):
            EQ.cell(r, 1, f"SEMI {nm}")
            ec(r, 2025, f"=eq_s_{k}_26/(1+eq_g_{k}_26)"); ec(r, 2026, f"=eq_s_{k}_26"); ec(r, 2027, f"=D{r}*(1+eq_g_{k}_27)")
            ec(r, 2028, f"=E{r}*(1+eq_g_{k}_28)"); ec(r, 2029, f"=F{r}*(1+eq_g29)")
            EQROW[k] = r; r += 1
        for k, nm in (("test", "測試設備"), ("pkg", "組裝與封裝設備")):
            EQ.cell(r, 1, f"SEMI {nm}")
            ec(r, 2025, f"=eq_s_{k}_26/(1+eq_g_{k}_26)"); ec(r, 2026, f"=eq_s_{k}_26"); ec(r, 2027, f"=SQRT(eq_s_{k}_26*eq_s_{k}_28)")
            ec(r, 2028, f"=eq_s_{k}_28"); ec(r, 2029, f"=F{r}*(1+eq_g29)")
            EQ.cell(r, 8, "2027 幾何內插（Derived；SEMI 年中預測未給 2027）"); EQROW[k] = r; r += 1
        EQ.cell(r, 1, "SEMI 設備合計（WFE＋測試＋封裝）").font = BOLD
        for y, col in YC.items(): ec(r, y, f"={col}{EQROW['wfe']}+{col}{EQROW['test']}+{col}{EQROW['pkg']}", fill=KEY)
        EQROW["tot"] = r; r += 2
        EQ.cell(r, 1, "B. SoC 測試機市場（日曆年，US$B）").font = BOLD; r += 1
        PL_ = ("bw", "rb", "ru", "hp", "amd", "tpu", "trn", "oth")
        for s_ in SCL:
            EQ.cell(r, 1, "測試時間加權顆數（百萬顆）"); EQ.cell(r, 2, s_)
            for y in (2026, 2027, 2028, 2029):
                col = "CDEF"[y - 2026]
                m = kc(y, s_)
                ec(r, y, "=(" + "+".join(f"Units!{col}{UROW[k]}*ts_tt_{k}" for k in PL_) + f"){m}", "0.000")
            EQROW["wu" + s_] = r; r += 1
        EQ.cell(r, 1, "每加權顆測試支出（以 CY2026 校準；US$B／百萬顆）")
        ec(r, 2026, f"=(ts_mk_soc26-ts_nonai26)/D{EQROW['wu' + SCL[0]]}", "0.0000"); EQ.cell(r, 8, "（公司 CY2026 SoC 市場 − 非 AI）÷ 加權顆數；平行測試效率提升基準為 0"); EQROW["cw"] = r; r += 1
        EQ.cell(r, 1, "非 AI SoC 測試機市場")
        for y in (2026, 2027, 2028, 2029): ec(r, y, f"=ts_nonai26*(1+ts_g_nonai)^{y - 2026}")
        EQROW["nai"] = r; r += 1
        for s_ in SCL:
            EQ.cell(r, 1, "SoC 測試機市場").font = BOLD; EQ.cell(r, 2, s_)
            ec(r, 2025, "=ts_mk_soc25")
            for y in (2026, 2027, 2028, 2029):
                col = YC[y]
                ec(r, y, f"=$D${EQROW['cw']}*{col}{EQROW['wu' + s_]}+{col}{EQROW['nai']}", fill=KEY)
            EQROW["soc" + s_] = r; r += 1
        r += 1
        EQ.cell(r, 1, "C. 測試 ÷ WFE（CF15；泰瑞達長期 7–9% 是其 ATE 口徑，SEMI 測試設備另含探針台、分選機等）").font = BOLD; r += 1
        EQ.cell(r, 1, "SEMI 測試設備 ÷ SEMI WFE")
        for y, col in YC.items(): ec(r, y, f"={col}{EQROW['test']}/{col}{EQROW['wfe']}", "0.0%")
        EQROW["r_semi"] = r; r += 1
        EQ.cell(r, 1, "SoC 測試機市場（情境 A）÷ SEMI WFE")
        for y, col in YC.items(): ec(r, y, f"={col}{EQROW['soc' + SCL[0]]}/{col}{EQROW['wfe']}", "0.0%")
        EQROW["r_soc"] = r; r += 1
        EQ.cell(r, 1, "對照：泰瑞達測試 ÷ WFE 長期區間")
        ec(r, 2026, '=TEXT(ts_tw_lo,"0%")&"–"&TEXT(ts_tw_hi,"0%")', "@"); r += 1
        if "hbm_hbp_25" in INROW:
            r += 1
            EQ.cell(r, 1, "D. 混合鍵合占 HBM 工序比例（行業技術情境：CY2028 起 A 高、B 基準、C 低；disco、hanmi、besi 同讀）").font = BOLD; r += 1
            for s_ in SCL:
                suf = {"A": "_hi", "B": "", "C": "_lo"}[s_]
                EQ.cell(r, 1, "混合鍵合占 HBM 工序比例"); EQ.cell(r, 2, s_)
                for y in YC: ec(r, y, f"=hbm_hbp_{y % 100}" + (suf if y >= 2028 else ""), "0.0%")
                EQROW["hbp" + s_] = r; r += 1
            EQ.cell(r, 1, "HBM 平均堆疊層數（基準）")
            for y in YC: ec(r, y, f"=hbm_stk_{y % 100}", "0.0")
            EQROW["stk"] = r; r += 1
        EQ.column_dimensions["A"].width = 52; EQ.column_dimensions["H"].width = 60
        for col in "BCDEFG": EQ.column_dimensions[col].width = 10

    # ---------------- Recon：由下往上對帳（v0.11；個股基準情境 → 本庫市場；R4 部署 GW → Wattonomics） ----------------
    RC = None; RCROW = {}
    rc_ptr = os.path.join(SIL, "data", "CURRENT_RECON")
    if "rc_tol" in INROW and os.path.exists(rc_ptr):
        rc_file = open(rc_ptr).read().strip()
        rsnap = json.load(open(os.path.join(SIL, "data", rc_file), encoding="utf-8"))
        evf = sorted(glob.glob(os.path.join(HERE, "seed", "recon_evidence_v*.json")), key=_vkey)
        rev = json.load(open(evf[-1], encoding="utf-8")) if evf else {"r4": {}, "soc": {}}
        RC = wb.create_sheet("Recon")
        RC["A1"] = ("Recon — 由下往上對帳（工作方式 v1.7 第 8 節；決策 S32）：覆蓋個股的基準情境（情境 " + rsnap.get("base_scen", "B") + "）加總 → 本庫市場；"
                    "本庫推得上線 GW → Wattonomics 主路徑。個股值與 R4 值為版本固定快照（tools/recon.py → data/" + rc_file + "），綠字＝外部快照；"
                    "狀態：|比值 − 1| ≤ 容差 rc_tol＝一致；超出且有已寄出的證據送件＝時差；其餘＝待診斷")
        hdr(RC, 3, ["項目", "說明", "2025", "2026", "2027", "2028", "2029", "來源／證據送件"])
        RYC = "CDEFG"
        r = 4
        soc = rsnap["markets"]["soc"]
        RC.cell(r, 1, "A. SoC 測試機（份額口徑；分母＝本庫 Equip 頁 SoC 測試機市場，兩家模型同讀）").font = BOLD; r += 1
        r_first = r
        for row in soc["rows"]:
            RC.cell(r, 1, row["who"] + " SoC 份額"); RC.cell(r, 2, row["note"])
            for j, v in enumerate(row["vals"]):
                c = RC.cell(r, 3 + j, v); c.font = GREEN; c.number_format = "0.0%"
            RC.cell(r, 8, row["ref"]); r += 1
        RC.cell(r, 1, "覆蓋個股合計").font = BOLD
        for j in range(5):
            c = RC.cell(r, 3 + j, f"=SUM({RYC[j]}{r_first}:{RYC[j]}{r-1})"); c.number_format = "0.0%"; c.fill = KEY
        RCROW["soc_sum"] = r; r += 1
        RC.cell(r, 1, "其他廠（1 − 合計）")
        for j in range(5): RC.cell(r, 3 + j, f"=1-{RYC[j]}{r-1}").number_format = "0.0%"
        r += 1
        RC.cell(r, 1, "預期合計（1 − rc_soc_oth）"); RC.cell(r, 2, "In 頁 M27")
        for j in range(5): RC.cell(r, 3 + j, "=1-rc_soc_oth").number_format = "0.0%"
        RCROW["soc_exp"] = r; r += 1
        RC.cell(r, 1, "比值（合計 ÷ 預期）")
        for j in range(5): RC.cell(r, 3 + j, f"={RYC[j]}{RCROW['soc_sum']}/{RYC[j]}{RCROW['soc_exp']}").number_format = "0.000"
        RCROW["soc_ratio"] = r; r += 1
        RC.cell(r, 1, "證據送件")
        for j, y in enumerate(soc["years"]):
            v = rev.get("soc", {}).get(str(y), "")
            if v: RC.cell(r, 3 + j, v)
        RCROW["soc_ev"] = r; r += 1
        RC.cell(r, 1, "狀態").font = BOLD
        for j in range(5):
            col = RYC[j]
            c = RC.cell(r, 3 + j, f'=IF(ABS({col}{RCROW["soc_ratio"]}-1)<=rc_tol,"一致",IF(LEN({col}{RCROW["soc_ev"]})>0,"時差","待診斷"))'); c.fill = KEY
        RCROW["soc_st"] = r; r += 2
        r4 = rsnap["markets"]["r4"]
        RC.cell(r, 1, "B. R4 部署 GW（本庫推得新增上線 IT GW vs Wattonomics 主路徑；口徑同對方主路徑〔v0.1.8 起毛額〕；本庫基準〔v0.12 起＝原情境 B〕；r3 起）").font = BOLD; r += 1
        RC.cell(r, 1, "本庫推得新增上線 IT GW"); RC.cell(r, 2, "r4_bridge b8（ASIC 實體換算、出貨年→上線年、庫存；Wattonomics 為毛額時不扣汰換）")
        for j, v in enumerate(r4["sil_online"]):
            c = RC.cell(r, 4 + j, v); c.font = GREEN; c.number_format = "0.00"
        RC.cell(r, 8, r4["ref"]); RCROW["r4_s"] = r; r += 1
        RC.cell(r, 1, "Wattonomics 年新增 IT GW（主路徑）"); RC.cell(r, 2, "r4_bridge b9")
        for j, v in enumerate(r4["wt_main"]):
            c = RC.cell(r, 4 + j, v); c.font = GREEN; c.number_format = "0.00"
        RC.cell(r, 8, r4["ref"]); RCROW["r4_w"] = r; r += 1
        RC.cell(r, 1, "比值（Wattonomics ÷ 本庫）")
        for j in range(4): RC.cell(r, 4 + j, f"={'DEFG'[j]}{RCROW['r4_w']}/{'DEFG'[j]}{RCROW['r4_s']}").number_format = "0.000"
        RCROW["r4_ratio"] = r; r += 1
        RC.cell(r, 1, "本庫基準 ÷ Wattonomics 基準 − 1"); RC.cell(r, 2, "＝1 ÷ 比值 − 1；個股自訂情境寫明產業 GW 相對 Wattonomics 基準 ±% 時的換算基礎（Interface SIF_GWvsWT_Base）")
        for j in range(4): RC.cell(r, 4 + j, f"=1/{'DEFG'[j]}{RCROW['r4_ratio']}-1").number_format = "+0.0%;-0.0%"
        RCROW["r4_vs"] = r; r += 1
        adj = rev.get("r4_wt_adj")
        if adj:
            RC.cell(r, 1, "對照：Wattonomics 側歸因調整後比值（尚未進入 Wattonomics 路徑）"); RC.cell(r, 8, adj["src"])
            for j, v in enumerate(adj["vals"]):
                c = RC.cell(r, 4 + j, v); c.font = GREEN; c.number_format = "0.000"
            r += 1
        RC.cell(r, 1, "證據送件")
        for j, y in enumerate(r4["years"]):
            v = rev.get("r4", {}).get(str(y), "")
            if v: RC.cell(r, 4 + j, v)
        RCROW["r4_ev"] = r; r += 1
        RC.cell(r, 1, "狀態").font = BOLD
        for j in range(4):
            col = "DEFG"[j]
            c = RC.cell(r, 4 + j, f'=IF(ABS({col}{RCROW["r4_ratio"]}-1)<=rc_tol,"一致",IF(LEFT({col}{RCROW["r4_ev"]},6)="兩邊都不錨定","兩邊都不錨定",IF(LEN({col}{RCROW["r4_ev"]})>0,"時差","待診斷")))'); c.fill = KEY
        RCROW["r4_st"] = r; r += 2
        RC.cell(r, 1, "C. 結構一致的市場（覆蓋個股直接讀本庫份額 × 本庫市場，比值恆等於份額合計；由 Checks K6 約束 ≤ 100%）").font = BOLD; r += 1
        for mk, k6 in (("ASIC 設計服務（TPU、Trainium、其他 ASIC）", "K6-tpu／oth／trn"), ("CPO 光學", "K6-cpo"), ("加速器 ABF 載板", "K6-abfnv／abfot"),
                       ("AI 電源", "K6-pwr"), ("AI 液冷", "K6-lc"), ("AI 伺服器組裝", "K6-asm"), ("HBM", "K6-hbm")):
            RC.cell(r, 1, mk); RC.cell(r, 2, "結構一致（份額由本庫擁有）"); RC.cell(r, 8, "Checks " + k6); r += 1
        r += 1
        RC.cell(r, 1, "D. 尚未建立對帳的市場（v0.12 起依序補：份額或營收 ÷ 市場）").font = BOLD; r += 1
        for mk in ("GPU／ASIC 晶片金額（NVIDIA、AMD、Broadcom、Marvell、聯發科、創意、世芯）", "CoWoS（台積電；Amkor、日月光委外）",
                   "晶圓前段設備（應用材料、Lam、KLA、ASML；TEL 未建模）", "AI 高速 PCB／CCL（勝宏、金像電、台光電）", "光學模組總量（中際旭創、新易盛、Lumentum、Coherent）"):
            RC.cell(r, 1, mk); RC.cell(r, 2, "未建"); r += 1
        r += 1
        nvq = rsnap["markets"].get("nv_dc")
        if nvq:
            RC.cell(r, 1, "F. 季度實績驗證（v0.13；三層校準原則：實績雙向校正）：NVIDIA 資料中心營收（出貨年 NVIDIA GW × 每 GW 內容 vs 公司實績／指引；只作診斷，不計入狀態彙總）").font = BOLD; r += 1
            for lab, key, fm in (("隱含 NVIDIA 資料中心營收（$B）", "implied", "0.0"), ("公司（2026 錨點；2027 指引推得）（$B）", "company", "0.0")):
                RC.cell(r, 1, lab)
                for j, v in enumerate(nvq[key]):
                    if v is not None:
                        c = RC.cell(r, 4 + j, v); c.font = GREEN; c.number_format = fm
                RC.cell(r, 8, nvq["ref"]); RCROW["nv_" + key] = r; r += 1
            RC.cell(r, 1, "比值（隱含 ÷ 公司）")
            for j in range(len(nvq["implied"])):
                if nvq["implied"][j] is not None:
                    RC.cell(r, 4 + j, f"={'DEFG'[j]}{RCROW['nv_implied']}/{'DEFG'[j]}{RCROW['nv_company']}").number_format = "0.000"
            RCROW["nv_ratio"] = r; r += 1
            RC.cell(r, 1, "狀態")
            for j in range(len(nvq["implied"])):
                if nvq["implied"][j] is not None:
                    col = "DEFG"[j]
                    RC.cell(r, 4 + j, f'=IF(ABS({col}{RCROW["nv_ratio"]}-1)<=rc_tol,"一致","待診斷")')
            r += 2
        RC.cell(r, 1, "E. 狀態彙總（A、B 節）").font = BOLD; r += 1
        st_cells = f"C{RCROW['soc_st']}:G{RCROW['soc_st']},D{RCROW['r4_st']}:G{RCROW['r4_st']}"
        for lab in ("一致", "時差", "兩邊都不錨定", "待診斷"):
            RC.cell(r, 1, lab)
            RC.cell(r, 3, f'=COUNTIF(C{RCROW["soc_st"]}:G{RCROW["soc_st"]},"{lab}")+COUNTIF(D{RCROW["r4_st"]}:G{RCROW["r4_st"]},"{lab}")')
            RCROW["n_" + lab] = r; r += 1
        RC.column_dimensions["A"].width = 46; RC.column_dimensions["B"].width = 44; RC.column_dimensions["H"].width = 80
        for col in "CDEFG": RC.column_dimensions[col].width = 11

    # ---------------- Cap：供給上限（v0.13；三層校準原則 Andy 2026-11-04：基準 GW＝需求側錨點，但不得超過供給產能） ----------------
    CP = None; CPROW = {}
    if UN is not None and DP is not None and "cw_capend29" in INROW:
        CP = wb.create_sheet("Cap")
        CP["A1"] = ("Cap — 供給上限（v0.13；三層校準原則〔Andy 2026-11-04〕：實績雙向校正、供給是上限、錨點可信度分時段看）。"
                    "以本庫基準顆數（原始顆數 × 基準係數）算 CoWoS 與 HBM 需求，對照產能；可支撐部署 GW＝本庫基準部署 GW × MIN(1, 產能 ÷ 需求)。"
                    "Wattonomics 錨定部署 GW（本庫口徑）＝本庫基準 × R4 比值（Recon B 節）；兩者都不得超過可支撐部署 GW（Checks K10）")
        hdr(CP, 3, ["項目", "單位", "2026", "2027", "2028", "2029", "說明"])
        CC = {2026: "C", 2027: "D", 2028: "E", 2029: "F"}
        PPK = {"bw": "GB", "rb": "VR", "ru": "RU", "hp": "HP", "trn": "TRN", "oth": "ASIC"}
        def ppk(k, y):
            if k == "amd": return "AMD8" if y == 2026 else "AMDR"
            if k == "tpu": return "TPU7" if y == 2026 else "TPU8"
            return PPK[k]
        coef = lambda y: f"crd_k_{y % 100}" if COORD else {2026: "1", 2027: "1", 2028: "um28_B", 2029: "um29_B"}[y]
        r = 4
        CP.cell(r, 1, "A. CoWoS（千片）").font = BOLD; r += 1
        r_first = r
        for g in seed["unit_groups"]:
            k = g["key"]
            CP.cell(r, 1, "  " + g["name"]); CP.cell(r, 2, "千片")
            for y, col in CC.items():
                CP.cell(r, list(CC).index(y) + 3, f"=Units!{col}{UROW[k]}*{coef(y)}/cw_pp_{ppk(k, y)}*1000").number_format = "#,##0"
            CP.cell(r, 7, "基準顆數 ÷ 每片封裝數（In M14）"); r += 1
        CP.cell(r, 1, "  其他 CoWoS 用途"); CP.cell(r, 2, "千片")
        CP.cell(r, 3, "=cw_oth26").number_format = "#,##0"
        for y in (2027, 2028, 2029):
            CP.cell(r, list(CC).index(y) + 3, f"={CC[y-1]}{r}*(1+cw_oth_g)").number_format = "#,##0"
        r += 1
        CP.cell(r, 1, "CoWoS 需求合計（基準）").font = BOLD
        for y, col in CC.items(): CP.cell(r, list(CC).index(y) + 3, f"=SUM({col}{r_first}:{col}{r-1})").number_format = "#,##0"
        CPROW["cw_dem"] = r; r += 1
        CP.cell(r, 1, "CoWoS 可用產能（台積電月產能年均 × 12 ×（1＋委外））"); CP.cell(r, 2, "千片")
        for y, col in CC.items():
            CP.cell(r, list(CC).index(y) + 3, f"=(cw_capend{y % 100 - 1}+cw_capend{y % 100})/2*12*(1+cw_osat)").number_format = "#,##0"
        CP.cell(r, 7, "產能年底值為 Interested-party（2029 Assumed）；委外 cw_osat")
        CPROW["cw_cap"] = r; r += 1
        CP.cell(r, 1, "CoWoS 需求 ÷ 產能"); CP.cell(r, 2, "%")
        for y, col in CC.items(): CP.cell(r, list(CC).index(y) + 3, f"={col}{CPROW['cw_dem']}/{col}{CPROW['cw_cap']}").number_format = "0%"
        CP.cell(r, 7, "> 100%＝CoWoS 為瓶頸"); CPROW["cw_util"] = r; r += 2
        CP.cell(r, 1, "B. HBM（EB；Mkt 頁，基準＝情境 B）").font = BOLD; r += 1
        MC = {2026: "D", 2027: "E", 2028: "F", 2029: "G"}
        CP.cell(r, 1, "HBM 需求（毛額，基準）"); CP.cell(r, 2, "EB")
        for y in CC: CP.cell(r, list(CC).index(y) + 3, f"=Mkt!{MC[y]}{MK['shipB']}+MAX(0,Mkt!{MC[y]}{MK['gapB']})").number_format = "0.00"
        CPROW["hb_dem"] = r; r += 1
        CP.cell(r, 1, "HBM 位元供給"); CP.cell(r, 2, "EB")
        for y in CC: CP.cell(r, list(CC).index(y) + 3, f"=Mkt!{MC[y]}{MK['sup']}").number_format = "0.00"
        CPROW["hb_sup"] = r; r += 1
        CP.cell(r, 1, "HBM 需求 ÷ 供給"); CP.cell(r, 2, "%")
        for y, col in CC.items(): CP.cell(r, list(CC).index(y) + 3, f"={col}{CPROW['hb_dem']}/{col}{CPROW['hb_sup']}").number_format = "0%"
        CP.cell(r, 7, "> 100%＝HBM 為瓶頸（本庫 HBM 出貨已取 MIN(需求, 供給)）"); CPROW["hb_util"] = r; r += 2
        CP.cell(r, 1, "C. 部署 GW 上限（本庫口徑：成本等值、出貨年）").font = BOLD; r += 1
        CP.cell(r, 1, "本庫基準部署 GW"); CP.cell(r, 2, "GW")
        for y, col in CC.items(): CP.cell(r, list(CC).index(y) + 3, f"=Deploy!{col}{DROW['gwB']}").number_format = "0.0"
        CPROW["gw_b"] = r; r += 1
        CP.cell(r, 1, "可支撐部署 GW（供給上限）").font = BOLD; CP.cell(r, 2, "GW")
        for y, col in CC.items():
            c = CP.cell(r, list(CC).index(y) + 3, f"={col}{CPROW['gw_b']}*MIN(1,1/{col}{CPROW['cw_util']},1/{col}{CPROW['hb_util']})"); c.number_format = "0.0"; c.fill = KEY
        CP.cell(r, 7, "＝基準 × MIN(1, 產能 ÷ 需求)；兩項皆未達瓶頸時＝基準。只算 CoWoS、HBM；電力、晶圓前段未建")
        CPROW["gw_cap"] = r; r += 1
        # v0.16：分項上限（算力收支 11-04：CoWoS、HBM 分開看；單一限制、不以基準截頂＝基準 × 產能 ÷ 需求）
        for key, lab in (("gw_cw", "  其中：只受 CoWoS 限制的可支撐部署 GW（基準 × 產能 ÷ 需求；不截頂）"), ("gw_hb", "  其中：只受 HBM 限制的可支撐部署 GW（基準 × 供給 ÷ 需求；不截頂）")):
            CP.cell(r, 1, lab); CP.cell(r, 2, "GW")
            ut = CPROW["cw_util" if key == "gw_cw" else "hb_util"]
            for y, col in CC.items():
                CP.cell(r, list(CC).index(y) + 3, f"={col}{CPROW['gw_b']}/{col}{ut}").number_format = "0.0"
            CP.cell(r, 7, "可支撐部署 GW（供給上限）＝MIN(基準, 本列, 下一列)" if key == "gw_cw" else "2026 HBM 供給為廠商出貨推估（v0.14），非獨立產能數據")
            CPROW[key] = r; r += 1
        CP.cell(r, 1, "上限留餘（產能可支撐 ÷ 基準需求 − 1）"); CP.cell(r, 2, "%")
        for y, col in CC.items():
            CP.cell(r, list(CC).index(y) + 3, f"=1/MAX({col}{CPROW['cw_util']},{col}{CPROW['hb_util']})-1").number_format = "+0%;-0%"
        CPROW["head"] = r; r += 1
        if RC is not None and "r4_ratio" in RCROW:
            CP.cell(r, 1, "Wattonomics 錨定部署 GW（本庫口徑）"); CP.cell(r, 2, "GW")
            for y, col in CC.items():
                CP.cell(r, list(CC).index(y) + 3, f"={col}{CPROW['gw_b']}*Recon!{'DEFG'[list(CC).index(y)]}{RCROW['r4_ratio']}").number_format = "0.0"
            CP.cell(r, 7, "＝本庫基準 × R4 比值（Wattonomics ÷ 本庫推得上線 GW）；近似換算，只作上限檢查")
            CPROW["gw_wt"] = r; r += 1
        CP.cell(r, 1, "狀態（本庫基準 ≤ 上限）").font = BOLD
        for y, col in CC.items():
            CP.cell(r, list(CC).index(y) + 3, f'=IF({col}{CPROW["gw_b"]}<={col}{CPROW["gw_cap"]}*1.000001,"未觸頂","超過上限")').fill = KEY
        CPROW["st"] = r; r += 1
        CP.column_dimensions["A"].width = 48; CP.column_dimensions["G"].width = 70
        for col in "BCDEF": CP.column_dimensions[col].width = 11

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
        c = IF.cell(r, 5, f'=IF(ISBLANK(In!E{ir}),"",In!E{ir})' if k.startswith("crd_prev") else f"=In!E{ir}"); c.font = GREEN   # v0.19-trial：上次判定空白＝首次判定，不得顯示為 0
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
    derived = [("SIF_HBM_DemandEB", "dem", "HBM 需求（加速器裝載，情境 A（原始顆數））", "EB"),
               ("SIF_HBM_DemandGrossEB", "demg", "HBM 需求 × 出貨係數（情境 A（原始顆數））", "EB"),
               ("SIF_Units_Total", "units", "加速器顆數合計（情境 A（原始顆數））", "百萬顆"),
               ("SIF_HBM_GBperUnit", "gbavg", "每顆平均 HBM", "GB"),
               ("SIF_HBM_SupplyEB", "sup", "HBM 位元供給（三情境相同）", "EB")]
    for s in SCL:
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
                 + [(f"SIF_UnitsPlat_Total_{s_}", "tot" + s_, f"加速器顆數合計（{s_}）") for s_ in SCL])
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
        for code, lab in [("tpu", "Google TPU"), ("oth", "其他 ASIC")] + [("trn" + s_, f"AWS Trainium（情境 {s_}）") for s_ in SCL] + ([("cpo", "CPO 光學（2027–2029）")] if "sum_cpo" in STROW else []) + ([("abfnv", "NVIDIA 加速器 ABF 載板（2026–2029）"), ("abfot", "非 NVIDIA 加速器 ABF 載板（2026–2029）")] if "sum_abfnv" in STROW else []):
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
        items = ([(f"SIF_AIITCapex_{s_}", "itc" + s_, f"AI IT 資本支出（{s_}）", "$B") for s_ in SCL]
                 + [("SIF_ITCapexPerGW", "itpg", "每 GW IT 資本支出（2026 依 vr26 混合）", "$B/GW")]
                 + [(f"SIF_DeployGW_{s_}", "gw" + s_, f"部署 GW（{s_}）", "GW") for s_ in SCL]
                 + [("SIF_DeployGWPhys_Base", "gwphB", "實體部署 GW（基準；NVIDIA 依世代每 GW 顆數、非 NVIDIA 依持有成本比；v0.16.2）", "GW")])
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
        items = ([(f"SIF_ABFDemIdx_{s_}", "dem" + s_, f"ABF 需求指數（{s_}；2026＝1）", "x") for s_ in SCL]
                 + [("SIF_ABFSubCapIdx", "sub", "載板產能指數（加權；2026＝1）", "x"), ("SIF_ABFFilmIdx", "film", "味之素 ABF 膜上限（2026＝1）", "x"), ("SIF_ABFSupIdx", "sup", "有效供給指數", "x")]
                 + [(f"SIF_ABFGap_{s_}", "gap" + s_, f"ABF 供需缺口（{s_}）", "%") for s_ in SCL])
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
                 + [(f"SIF_PCBMkt_{s_}", "pcbm" + s_, f"AI 高速 PCB 市場（{s_}；不含價格指數）", "US$B") for s_ in SCL]
                 + [(f"SIF_CCLMkt_{s_}", "cclm" + s_, f"AI 高速 CCL 市場（{s_}；不含價格指數）", "US$B") for s_ in SCL])
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
                 + [(f"SIF_AsmMkt_{s_}", "asm" + s_, f"AI 伺服器組裝市場（ODM 價；{s_}）", "US$B") for s_ in SCL]
                 + [("SIF_SysValGB", "vgb", "每顆系統價值（GB300 世代）", "US$K"), ("SIF_SysValVR", "vvr", "每顆系統價值（Rubin 世代）", "US$K"),
                    ("SIF_SysUnitsGB", "ugb", "GB300 世代商用 GPU 顆數（情境 A）", "百萬顆"), ("SIF_SysUnitsVR", "uvr", "Rubin 世代商用 GPU 顆數（情境 A）", "百萬顆")]
                 + [(f"SIF_GPUSysMkt_{s_}", "sys" + s_, f"商用 GPU 系統市場（{s_}）", "US$B") for s_ in SCL]
                 + [(f"SIF_PwrMkt_{s_}", "pwr" + s_, f"AI 電源市場（{s_}）", "US$B") for s_ in SCL]
                 + ([("SIF_LiqPerGW", "lcpg", "每 GW 液冷內容（機架內＋分歧管）", "US$M/GW")]
                    + [(f"SIF_LiqPen_{s_}", "lpen" + s_, f"液冷滲透率（{s_}）", "%") for s_ in SCL]
                    + [(f"SIF_LiqMkt_{s_}", "lc" + s_, f"AI 液冷市場（{s_}）", "US$B") for s_ in SCL] if "lcpg" in RKROW else []))
        for nm_, key, lab, unit in items:
            kr = RKROW[key]
            IF.cell(r, 1, nm_); IF.cell(r, 2, lab); IF.cell(r, 3, unit)
            for j in range(4):
                c = IF.cell(r, 5 + j, f"=Rack!{'CDEF'[j]}{kr}"); c.font = GREEN; c.number_format = "#,##0.000"
            IF.cell(r, 9, f"Rack!C{kr}:F{kr}")
            name(wb, nm_, f"Interface!$E${r}:$H${r}"); IFNAMES.append(nm_); r += 1
    if EQ is not None:
        r += 1
        IF.cell(r, 1, "B10. 設備（Equip 頁；SEMI 分項欄 D–H＝2025–2029；SoC 測試機市場 E–H＝2026–2029）").font = BOLD; r += 1
        for k, nm in (("dram", "DRAM"), ("nand", "NAND"), ("fd", "FoundryLogic"), ("wfe", "WFE"), ("test", "Test"), ("pkg", "Pkg"), ("tot", "Total")):
            nm_ = f"SIF_SEMI_{nm}"; er = EQROW[k]
            IF.cell(r, 1, nm_); IF.cell(r, 2, f"SEMI 設備：{nm}"); IF.cell(r, 3, "US$B")
            for j in range(5):
                c = IF.cell(r, 4 + j, f"=Equip!{'CDEFG'[j]}{er}"); c.font = GREEN; c.number_format = "0.00"
            IF.cell(r, 9, f"Equip!C{er}:G{er}")
            name(wb, nm_, f"Interface!$D${r}:$H${r}"); IFNAMES.append(nm_); r += 1
        for s_ in SCL:
            nm_ = f"SIF_TestSoCMkt_{s_}"; er = EQROW["soc" + s_]
            IF.cell(r, 1, nm_); IF.cell(r, 2, f"SoC 測試機市場（{s_}；2025 為公司值）"); IF.cell(r, 3, "US$B")
            for j in range(5):
                c = IF.cell(r, 4 + j, f"=Equip!{'CDEFG'[j]}{er}"); c.font = GREEN; c.number_format = "0.000"
            IF.cell(r, 9, f"Equip!C{er}:G{er}")
            name(wb, nm_, f"Interface!$D${r}:$H${r}"); IFNAMES.append(nm_); r += 1
        for key, nm_, lab in ((("r_semi", "SIF_TestWFE_SEMI", "SEMI 測試 ÷ SEMI WFE"), ("r_soc", "SIF_TestWFE_SoC", "SoC 測試機（基準）÷ SEMI WFE"))
                              + tuple((f"hbp{s_}", f"SIF_HBMHybrid_{s_}", f"混合鍵合占 HBM 工序比例（{s_}）") for s_ in SCL if f"hbp{s_}" in EQROW)):
            er = EQROW[key]
            IF.cell(r, 1, nm_); IF.cell(r, 2, lab); IF.cell(r, 3, "%")
            for j in range(5):
                c = IF.cell(r, 4 + j, f"=Equip!{'CDEFG'[j]}{er}"); c.font = GREEN; c.number_format = "0.0%"
            IF.cell(r, 9, f"Equip!C{er}:G{er}")
            name(wb, nm_, f"Interface!$D${r}:$H${r}"); IFNAMES.append(nm_); r += 1
    if ST is not None and "sum_lc" + SCL[0] in STROW:
        r += 1
        IF.cell(r, 1, "B9. AI 液冷份額合計（ShareTbl 頁；欄 E–H＝2026–2029）").font = BOLD; r += 1
        for s_ in SCL:
            nm_ = f"SIF_LiqShareSum_{s_}"; sr = STROW["sum_lc" + s_]
            IF.cell(r, 1, nm_); IF.cell(r, 2, f"奇鋐＋雙鴻＋健策液冷有效份額合計（{s_}）"); IF.cell(r, 3, "%")
            for j in range(4):
                c = IF.cell(r, 5 + j, f"=ShareTbl!{'DEFG'[j]}{sr}"); c.font = GREEN; c.number_format = "0.0%"
            IF.cell(r, 9, f"ShareTbl!D{sr}:G{sr}")
            name(wb, nm_, f"Interface!$E${r}:$H${r}"); IFNAMES.append(nm_); r += 1
    if ST is not None and "sum_asm" in STROW:
        r += 1
        IF.cell(r, 1, "B11. AI 伺服器組裝份額（ShareTbl 頁；欄 E–H＝2026–2029）").font = BOLD; r += 1
        for nm_, lab, sr, fm in (("SIF_AsmShareSum", "廣達＋鴻海＋緯創本體＋緯穎組裝有效份額合計", STROW["sum_asm"], "0.0%"),
                                 ("SIF_AsmWyTilt", "緯穎 ASIC 傾斜係數", STROW["asm_tilt"], "0.000"),
                                 ("SIF_AsmShare_wy", "緯穎組裝有效份額（× ASIC 傾斜）", STROW["sum_asm"] - 1, "0.0%")):
            IF.cell(r, 1, nm_); IF.cell(r, 2, lab); IF.cell(r, 3, "x" if "Tilt" in nm_ else "%")
            for j in range(4):
                c = IF.cell(r, 5 + j, f"=ShareTbl!{'DEFG'[j]}{sr}"); c.font = GREEN; c.number_format = fm
            IF.cell(r, 9, f"ShareTbl!D{sr}:G{sr}")
            name(wb, nm_, f"Interface!$E${r}:$H${r}"); IFNAMES.append(nm_); r += 1
    if ST is not None and "sum_hbm" in STROW:
        r += 1
        IF.cell(r, 1, "B10. HBM 營收份額合計（ShareTbl 頁；欄 D–H＝2025–2029）").font = BOLD; r += 1
        nm_ = "SIF_HBMShareSum"; sr = STROW["sum_hbm"]
        IF.cell(r, 1, nm_); IF.cell(r, 2, "SK 海力士＋三星＋美光 HBM 營收份額合計"); IF.cell(r, 3, "%")
        for j in range(5):
            c = IF.cell(r, 4 + j, f"=ShareTbl!{'DEFGH'[j]}{sr}"); c.font = GREEN; c.number_format = "0.0%"
        IF.cell(r, 9, f"ShareTbl!D{sr}:H{sr}")
        name(wb, nm_, f"Interface!$D${r}:$H${r}"); IFNAMES.append(nm_); r += 1
    if ST is not None and "sum_pwr" + SCL[0] in STROW:
        r += 1
        IF.cell(r, 1, "B8. AI 電源份額合計（ShareTbl 頁；欄 E–H＝2026–2029）").font = BOLD; r += 1
        for s_ in SCL:
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
    # D. 基準（v0.12；工作方式 v1.7「三層只給基準、情境在個股」；決策 S33、S35）：基準＝原情境 B（平平）。
    #    每個含情境字母 B 的名稱另建 *_Base 名稱（同值、指向 B 列）；A／B／C 名稱標「停用過渡」，v0.13 刪除。
    import re as _re
    _pat = _re.compile(r"^(SIF_.*?)(_B(?=_|$)|(?<=[a-z])B(?=_\d|$))(.*)$")
    def _base_name(nm_):
        m = _pat.match(nm_)
        if not m or nm_.endswith(("_lo", "_hi")):
            return None
        return m.group(1) + ("_Base" if m.group(2) == "_B" else "Base") + m.group(3)
    _bases = [(nm_, _base_name(nm_)) for nm_ in list(IFNAMES) if _base_name(nm_)]
    _scen_set = set()
    for nm_b, _ in _bases:
        m = _pat.match(nm_b)
        for x in "ABC":
            _scen_set.add(m.group(1) + m.group(2).replace("B", x) + m.group(3))
    # v0.15：推導節（B 節起）的 A／B／C 名稱已刪除（下游全部改讀 *_Base 或個股自訂；Wattonomics v0.2.2 起只讀 *_Base）；
    #        列保留在頁內作為 *_Base 的指向與參考。In 頁層的情境列（um28_A 等）名稱保留，作為個股自訂情境的參考值（v0.16 移到個股）
    _in_names = {"SIF_" + k for k in INROW}
    _to_del = []
    for rr in range(5, r):
        nm_ = IF.cell(rr, 1).value
        if nm_ in _scen_set:
            if nm_ in _in_names:
                IF.cell(rr, 2, f"{IF.cell(rr, 2).value}〔個股自訂情境參考值；基準讀 *_Base〕")
            else:
                IF.cell(rr, 2, f"{IF.cell(rr, 2).value}〔名稱已刪除（v0.15）；基準讀 *_Base〕")
                _to_del.append(nm_)
    r += 1
    IF.cell(r, 1, "D. 基準（v0.12 起；三層只給基準、情境在個股〔工作方式 v1.7〕；基準＝原情境 B〔平平〕，決策 S33、S35）。"
                  "下游基準一律讀本節 *_Base；個股的 A／C 由個股自訂，並寫明產業 GW 相對 Wattonomics 基準 ±%（見 Recon 頁 B 節、SIF_GWvsWT_Base）").font = BOLD; r += 1
    hdr(IF, r, ["名稱", "項目", "單位", "低／2025", "基準／2026", "高／2027", "2028", "2029", "等同（原情境 B 名稱）"]); r += 1
    for nm_b, nm_base in _bases:
        dst = wb.defined_names[nm_b].attr_text.split("!")[1].replace("$", "")
        src_r = int(_re.sub(r"[A-Z]", "", dst.split(":")[0]))
        IF.cell(r, 1, nm_base)
        IF.cell(r, 2, _re.sub(r"〔停用過渡.*〕$", "", str(IF.cell(src_r, 2).value)).replace("（B）", "（基準）").replace("（B；", "（基準；") + "（基準）")
        IF.cell(r, 3, IF.cell(src_r, 3).value)
        for j in range(4, 9):
            if IF.cell(src_r, j).value not in (None, ""):
                c = IF.cell(r, j, f"=IF(ISBLANK(Interface!{'ABCDEFGHI'[j-1]}{src_r}),\"\",Interface!{'ABCDEFGHI'[j-1]}{src_r})" if j > 3 else None)
                c.font = GREEN; c.number_format = IF.cell(src_r, j).number_format
        IF.cell(r, 9, nm_b)
        ref = dst.split(":")
        cols = [_re.sub(r"\d", "", x) for x in ref]
        name(wb, nm_base, f"Interface!${cols[0]}${r}" + (f":${cols[1]}${r}" if len(cols) > 1 else ""))
        if nm_b + "_lo" in wb.defined_names: name(wb, nm_base + "_lo", f"Interface!$D${r}")
        if nm_b + "_hi" in wb.defined_names: name(wb, nm_base + "_hi", f"Interface!$F${r}")
        IFNAMES.append(nm_base); r += 1
    for nm_ in _to_del:   # v0.15：*_Base 建好之後才刪（Base 列以 B 名稱的位置定址）
        for _x in (nm_, nm_ + "_lo", nm_ + "_hi"):
            if _x in wb.defined_names: del wb.defined_names[_x]
        if nm_ in IFNAMES: IFNAMES.remove(nm_)
    if RC is not None and "r4_vs" in RCROW:
        IF.cell(r, 1, "SIF_GWvsWT_Base"); IF.cell(r, 2, "本庫基準推得新增上線 IT GW 相對 Wattonomics 基準（＝本庫 ÷ Wattonomics − 1；個股情境的產業 GW 換算基礎）"); IF.cell(r, 3, "%")
        for j in range(4):
            c = IF.cell(r, 5 + j, f"=Recon!{'DEFG'[j]}{RCROW['r4_vs']}"); c.font = GREEN; c.number_format = "+0.0%;-0.0%"
        IF.cell(r, 9, f"Recon!D{RCROW['r4_vs']}:G{RCROW['r4_vs']}")
        name(wb, "SIF_GWvsWT_Base", f"Interface!$E${r}:$H${r}"); IFNAMES.append("SIF_GWvsWT_Base"); r += 1
        IF.cell(r, 1, "SIF_WTAnchor_Base"); IF.cell(r, 2, "Wattonomics 錨定係數（＝R4 比值 Wattonomics ÷ 本庫；個股敏感度「Wattonomics 錨定」案例：產業量自 CY2027 起 × 本係數）"); IF.cell(r, 3, "x")
        for j in range(4):
            c = IF.cell(r, 5 + j, f"=Recon!{'DEFG'[j]}{RCROW['r4_ratio']}"); c.font = GREEN; c.number_format = "0.000"
        IF.cell(r, 9, f"Recon!D{RCROW['r4_ratio']}:G{RCROW['r4_ratio']}")
        name(wb, "SIF_WTAnchor_Base", f"Interface!$E${r}:$H${r}"); IFNAMES.append("SIF_WTAnchor_Base"); r += 1
    if CP is not None:
        for nm_, key, lab, fm in (("SIF_CapGW_Base", "gw_cap", "可支撐部署 GW（CoWoS／HBM 供給上限；本庫口徑）", "0.0"),
                                  ("SIF_CapHead_Base", "head", "供給上限留餘（產能可支撐 ÷ 基準需求 − 1）", "+0%;-0%"),
                                  ("SIF_CoWoSUtil_Base", "cw_util", "CoWoS 需求 ÷ 產能（基準）", "0%"),
                                  ("SIF_HBMUtil_Base", "hb_util", "HBM 需求 ÷ 供給（基準）", "0%")):
            IF.cell(r, 1, nm_); IF.cell(r, 2, lab); IF.cell(r, 3, "GW" if nm_ == "SIF_CapGW_Base" else "%")
            for j in range(4):
                c = IF.cell(r, 5 + j, f"=Cap!{'CDEF'[j]}{CPROW[key]}"); c.font = GREEN; c.number_format = fm
            IF.cell(r, 9, f"Cap!C{CPROW[key]}:F{CPROW[key]}")
            name(wb, nm_, f"Interface!$E${r}:$H${r}"); IFNAMES.append(nm_); r += 1
        if RC is not None and "r4_s" in RCROW:
            IF.cell(r, 1, "SIF_CapOnlineGW_Base"); IF.cell(r, 2, "可支撐新增上線 IT GW（Wattonomics 可比口徑：本庫推得上線 GW × 可支撐 ÷ 基準部署；毛額）"); IF.cell(r, 3, "GW")
            for j in range(4):
                col = "CDEF"[j]
                c = IF.cell(r, 5 + j, f"=Recon!{'DEFG'[j]}{RCROW['r4_s']}*Cap!{col}{CPROW['gw_cap']}/Cap!{col}{CPROW['gw_b']}"); c.font = GREEN; c.number_format = "0.00"
            IF.cell(r, 9, "Recon B 節 × Cap C 節")
            name(wb, "SIF_CapOnlineGW_Base", f"Interface!$E${r}:$H${r}"); IFNAMES.append("SIF_CapOnlineGW_Base"); r += 1
        # v0.16：分項上限（算力收支 11-04 問題的後續）：部署口徑（本庫）與上線口徑（Wattonomics 可比，毛額）
        for nm_, key, lab in (("SIF_CapGW_CoWoS_Base", "gw_cw", "只受 CoWoS 限制的可支撐部署 GW（本庫口徑；不截頂）"),
                              ("SIF_CapGW_HBM_Base", "gw_hb", "只受 HBM 限制的可支撐部署 GW（本庫口徑；不截頂）")):
            IF.cell(r, 1, nm_); IF.cell(r, 2, lab); IF.cell(r, 3, "GW")
            for j in range(4):
                c = IF.cell(r, 5 + j, f"=Cap!{'CDEF'[j]}{CPROW[key]}"); c.font = GREEN; c.number_format = "0.0"
            IF.cell(r, 9, f"Cap!C{CPROW[key]}:F{CPROW[key]}")
            name(wb, nm_, f"Interface!$E${r}:$H${r}"); IFNAMES.append(nm_); r += 1
        if RC is not None and "r4_s" in RCROW:
            for nm_, key, lab in (("SIF_CapOnlineGW_CoWoS_Base", "gw_cw", "只受 CoWoS 限制的可支撐新增上線 IT GW（Wattonomics 可比口徑；毛額；不截頂）"),
                                  ("SIF_CapOnlineGW_HBM_Base", "gw_hb", "只受 HBM 限制的可支撐新增上線 IT GW（Wattonomics 可比口徑；毛額；不截頂）")):
                IF.cell(r, 1, nm_); IF.cell(r, 2, lab); IF.cell(r, 3, "GW")
                for j in range(4):
                    col = "CDEF"[j]
                    c = IF.cell(r, 5 + j, f"=Recon!{'DEFG'[j]}{RCROW['r4_s']}*Cap!{col}{CPROW[key]}/Cap!{col}{CPROW['gw_b']}"); c.font = GREEN; c.number_format = "0.00"
                IF.cell(r, 9, "Recon B 節 × Cap C 節")
                name(wb, nm_, f"Interface!$E${r}:$H${r}"); IFNAMES.append(nm_); r += 1
    if "nvr_GB300" in INROW:
        # v0.13.1：分世代每 IT GW NVIDIA 內容（Wattonomics Requests R7）＝NVIDIA 合計 ÷ IT capex（M28）× Tokenomics 每 IT GW IT capex（鏡像）
        for gen, rk in (("H100", "nvr_H100"), ("GB200", "nvr_GB200"), ("GB300", "nvr_GB300"), ("VR200_26", "nvr_VR26"), ("VR200", "nvr_VR27"), ("RU", "nvr_RU")):
            tkg = gen.split("_")[0]
            nm_ = f"SIF_NVContentGW_{gen}"
            IF.cell(r, 1, nm_); IF.cell(r, 3, "$B/GW")
            IF.cell(r, 2, f"每 IT GW NVIDIA 營收口徑內容（{gen.replace('_26', '，2026 HBM 價')}；{rk} × Tokenomics IF_CapexIT_{tkg}）")
            IF.cell(r, 4, f"={rk}_lo*TK_CapexIT_{tkg}").font = GREEN
            IF.cell(r, 5, f"={rk}*TK_CapexIT_{tkg}").font = GREEN
            IF.cell(r, 6, f"={rk}_hi*TK_CapexIT_{tkg}").font = GREEN
            for j in (4, 5, 6): IF.cell(r, j).number_format = "0.00"
            IF.cell(r, 7, f"=In!G{INROW[rk][0]}"); IF.cell(r, 9, f"In!E{INROW[rk][0]} × TK_In")
            name(wb, nm_, f"Interface!$E${r}"); name(wb, nm_ + "_lo", f"Interface!$D${r}"); name(wb, nm_ + "_hi", f"Interface!$F${r}")
            IFNAMES.append(nm_); r += 1
    if CO is not None:
        # v0.19-trial：三層基準協調（規則 v1.1 C3、C4、C7）。逐年純量名稱，值在欄 E
        r += 1
        IF.cell(r, 1, "E. 三層基準協調（v0.19-trial 試作；規則 docs/plan/三層基準協調規則.md v1.1；Coord 頁）。基準仍讀 SIF_DeployGW_Base 等 *_Base；"
                      "對帳中年份的個股另讀 SIF_DeployGW_Alt_<年>（另一側讀數）並列兩讀數目標價").font = BOLD; r += 1
        hdr(IF, r, ["名稱", "項目", "單位", "", "值", "", "", "", "來源位置"]); r += 1
        for key, stem, lab, unit, fm in (("sup", "SIF_DeployGW_Supply", "供給側讀數：原始顆數部署 GW（不乘係數；情境 A）", "GW", "0.00"),
                                         ("dem", "SIF_DeployGW_Demand", "需求側讀數：Wattonomics 換算值（平滑法）", "GW", "0.00"),
                                         ("cur", "SIF_DeployGW_Current", "現行取法（原始 × S33 係數；v0.18 基準）", "GW", "0.00"),
                                         ("alt", "SIF_DeployGW_Alt", "另一側讀數（未被採為基準的一條）", "GW", "0.00"),
                                         ("altr", "SIF_DeployGW_AltRatio", "另一側讀數 ÷ 基準", "x", "0.000"),
                                         ("st", "SIF_RecoStatus", "證據狀態（甲／乙／丙／待定；v1.3 區間跨線即待定）", "文字", "@"),
                                         ("lead", "SIF_RecoLead", "由誰定基準", "文字", "@"),
                                         ("lab", "SIF_RecoLabel", "該年標示（一致／對帳中／改用換算值／維持）", "文字", "@"),
                                         ("flag", "SIF_RecoFlag", "對帳中（1＝是）", "旗標", "0"),
                                         ("use", "SIF_RecoUseConv", "基準已改用換算值（1＝是）", "旗標", "0"),
                                         ("chain", "SIF_RecoChain", "連續法（1＝是；規則 v1.4：前一年待定或甲且維持現行，基準＝前一年基準 × 換算值成長率）", "旗標", "0")):
            for j, y in enumerate((2026, 2027, 2028, 2029)):
                nm_ = f"{stem}_{y}"; col = "CDEF"[j]
                IF.cell(r, 1, nm_); IF.cell(r, 2, f"{y} {lab}"); IF.cell(r, 3, unit)
                c = IF.cell(r, 5, f"=Coord!{col}{COROW[key]}"); c.font = GREEN; c.number_format = fm
                IF.cell(r, 9, f"Coord!{col}{COROW[key]}")
                name(wb, nm_, f"Interface!$E${r}"); IFNAMES.append(nm_); r += 1
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
    for s in SCL:
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
    # K2 基準顆數係數在區間內（v0.16：A＝1、A ≥ B ≥ C 的檢查移到市場層 markets/scen）
    for s in SCL:
        for y in (28, 29):
            CK.cell(r, 1, f"K2-{s}{y}"); CK.cell(r, 2, f"基準 20{y} 顆數係數在區間內")
            CK.cell(r, 3, f"=um{y}_{s}")
            CK.cell(r, 4, f'=um{y}_{s}_lo&"–"&um{y}_{s}_hi')
            CK.cell(r, 6, f'=IF(AND(C{r}>=um{y}_{s}_lo,C{r}<=um{y}_{s}_hi),"OK","ERROR")')
            CK.cell(r, 7, "驗證"); r += 1
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
        for code, lab in [("tpu", "Google TPU"), ("oth", "其他 ASIC")] + [("trn" + s_, f"AWS Trainium（情境 {s_}）") for s_ in SCL] + ([("cpo", "CPO 光學（模組與雷射廠）")] if "sum_cpo" in STROW else []) + ([("abfnv", "NVIDIA 加速器 ABF 載板"), ("abfot", "非 NVIDIA 加速器 ABF 載板")] if "sum_abfnv" in STROW else []):
            for j, y in enumerate((2027, 2028, 2029) if code == "cpo" else ((2026, 2027, 2028, 2029) if code.startswith("abf") else (2026, 2027, 2028))):
                CK.cell(r, 1, f"K6-{code}{y}"); CK.cell(r, 2, f"{lab} {y} {'光學金額份額合計' if code == 'cpo' else ('載板份額合計' if code.startswith('abf') else '設計夥伴份額合計')} ≤ 100%")
                CK.cell(r, 3, f"=ShareTbl!{'DEFG'[j]}{STROW['sum_' + code]}"); CK.cell(r, 4, 1)
                CK.cell(r, 6, f'=IF(C{r}<=1+1E-9,"OK","ERROR")'); CK.cell(r, 7, "驗證（加總受限）"); r += 1
    if ST is not None and "sum_pwr" + SCL[0] in STROW:
        for s_ in SCL:
            for j, y in enumerate((2026, 2027, 2028, 2029)):
                CK.cell(r, 1, f"K6-pwr{s_}{y}"); CK.cell(r, 2, f"AI 電源 {y} 台達電＋光寶有效份額合計 ≤ 100%（情境 {s_}）")
                CK.cell(r, 3, f"=ShareTbl!{'DEFG'[j]}{STROW['sum_pwr' + s_]}"); CK.cell(r, 4, 1)
                CK.cell(r, 6, f'=IF(C{r}<=1+1E-9,"OK","ERROR")'); CK.cell(r, 7, "驗證（加總受限）"); r += 1
    if ST is not None and "sum_lc" + SCL[0] in STROW:
        for s_ in SCL:
            for j, y in enumerate((2026, 2027, 2028, 2029)):
                CK.cell(r, 1, f"K6-lc{s_}{y}"); CK.cell(r, 2, f"AI 液冷 {y} 奇鋐＋雙鴻＋健策有效份額合計 ≤ 100%（情境 {s_}）")
                CK.cell(r, 3, f"=ShareTbl!{'DEFG'[j]}{STROW['sum_lc' + s_]}"); CK.cell(r, 4, 1)
                CK.cell(r, 6, f'=IF(C{r}<=1+1E-9,"OK","ERROR")'); CK.cell(r, 7, "驗證（加總受限）"); r += 1
    if ST is not None and "sum_asm" in STROW:
        for j, y in enumerate((2026, 2027, 2028, 2029)):
            CK.cell(r, 1, f"K6-asm{y}"); CK.cell(r, 2, f"AI 伺服器組裝 {y} 廣達＋鴻海＋緯創本體＋緯穎有效份額合計 ≤ 100%")
            CK.cell(r, 3, f"=ShareTbl!{'DEFG'[j]}{STROW['sum_asm']}"); CK.cell(r, 4, 1)
            CK.cell(r, 6, f'=IF(C{r}<=1+1E-9,"OK","ERROR")'); CK.cell(r, 7, "驗證（加總受限）"); r += 1
    if ST is not None and "sum_hbm" in STROW:
        for j, y in enumerate((2025, 2026, 2027, 2028, 2029)):
            CK.cell(r, 1, f"K6-hbm{y}"); CK.cell(r, 2, f"HBM {y} SK 海力士＋三星＋美光營收份額合計 ≤ 100%")
            CK.cell(r, 3, f"=ShareTbl!{'DEFGH'[j]}{STROW['sum_hbm']}"); CK.cell(r, 4, 1)
            CK.cell(r, 6, f'=IF(C{r}<=1+1E-9,"OK","ERROR")'); CK.cell(r, 7, "驗證（加總受限）"); r += 1
    if EQ is not None:
        for j, y in enumerate((2026, 2027, 2028)):
            col = "DEF"[j]
            CK.cell(r, 1, f"K8-{y}"); CK.cell(r, 2, f"{y} SoC 測試機市場（基準）÷ SEMI WFE 在泰瑞達長期區間內（7–9%；CF15）")
            CK.cell(r, 3, f"=Equip!{col}{EQROW['r_soc']}"); CK.cell(r, 4, '=TEXT(ts_tw_lo,"0%")&"–"&TEXT(ts_tw_hi,"0%")')
            CK.cell(r, 6, f'=IF(AND(C{r}>=ts_tw_lo,C{r}<=ts_tw_hi),"OK","WARN")'); CK.cell(r, 7, "驗證（量級；口徑不同，只列 WARN）"); r += 1
    if RC is not None:
        for j, y in enumerate((2025, 2026, 2027, 2028, 2029)):
            CK.cell(r, 1, f"K6-soc{y}"); CK.cell(r, 2, f"SoC 測試機 {y} Advantest＋泰瑞達份額合計 ≤ 100%（個股基準情境；Recon 頁 A 節）")
            CK.cell(r, 3, f"=Recon!{'CDEFG'[j]}{RCROW['soc_sum']}"); CK.cell(r, 4, 1)
            CK.cell(r, 6, f'=IF(C{r}<=1+1E-9,"OK","ERROR")'); CK.cell(r, 7, "驗證（加總受限）"); r += 1
        CK.cell(r, 1, "K9"); CK.cell(r, 2, "Recon 頁狀態為『待診斷』的格數（沒有證據送件的落差；規則：兩邊都更新後仍有落差須診斷錯在哪一方）")
        CK.cell(r, 3, f"=Recon!C{RCROW['n_待診斷']}+Recon!C{RCROW['n_兩邊都不錨定']}"); CK.cell(r, 6, f'=IF(C{r}=0,"OK","WARN")'); CK.cell(r, 7, "治理（對帳）")
        CK.cell(r, 8, "WARN＝仍有待診斷的落差；明細見 Recon 頁"); r += 1
    if CP is not None:
        for j, y in enumerate((2026, 2027, 2028, 2029)):
            col = "CDEF"[j]
            CK.cell(r, 1, f"K10-{y}"); CK.cell(r, 2, f"供給上限 {y}：本庫基準部署 GW ≤ CoWoS／HBM 可支撐部署 GW（Cap 頁；三層校準原則：供給是上限）")
            CK.cell(r, 3, f"=Cap!{col}{CPROW['gw_b']}"); CK.cell(r, 4, f"=Cap!{col}{CPROW['gw_cap']}"); CK.cell(r, 5, f"=C{r}-D{r}")
            CK.cell(r, 6, f'=IF(C{r}<=D{r}*1.000001,"OK","WARN")'); CK.cell(r, 7, "驗證（供給上限）")
            CK.cell(r, 8, "WARN＝本庫基準需求超過產能，基準須改以上限計（或產能證據過時）"); r += 1
        if "gw_wt" in CPROW:
            for j, y in enumerate((2026, 2027, 2028, 2029)):
                col = "CDEF"[j]
                CK.cell(r, 1, f"K10w-{y}"); CK.cell(r, 2, f"供給上限 {y}：Wattonomics 錨定部署 GW（本庫口徑）≤ 可支撐部署 GW")
                CK.cell(r, 3, f"=Cap!{col}{CPROW['gw_wt']}"); CK.cell(r, 4, f"=Cap!{col}{CPROW['gw_cap']}"); CK.cell(r, 5, f"=C{r}-D{r}")
                CK.cell(r, 6, f'=IF(C{r}<=D{r}*1.000001,"OK","WARN")'); CK.cell(r, 7, "驗證（供給上限）")
                CK.cell(r, 8, "WARN＝錨點超過供給產能，錨點須以上限計；寄交流給算力收支"); r += 1
    if RK is not None:
        CK.cell(r, 1, "K7"); CK.cell(r, 2, "Rack 頁：商用 GPU 系統市場 2026 ≤ AI IT 資本支出 2026（系統售價不應超過全部 IT 資本支出）")
        CK.cell(r, 3, f"=Rack!C{RKROW['sys' + SCL[0]]}"); CK.cell(r, 4, f"=Deploy!C{DROW['itc' + SCL[0]]}")
        CK.cell(r, 6, f'=IF(C{r}<=D{r}+1E-9,"OK","ERROR")'); CK.cell(r, 7, "驗證（量級）"); r += 1
    if CO is not None:
        for j, y in enumerate((2026, 2027, 2028, 2029)):
            col = "CDEF"[j]
            CK.cell(r, 1, f"K11-{y}"); CK.cell(r, 2, f"三層基準協調 {y}：Deploy 頁基準部署 GW＝Coord 頁依規則的基準（規則 v1.1 C3；檢查程式第 2 項的本庫側）")
            CK.cell(r, 3, f"=Deploy!{col}{DROW['gwB']}"); CK.cell(r, 4, f"=Coord!{col}{COROW['base']}"); CK.cell(r, 5, f"=C{r}-D{r}")
            CK.cell(r, 6, f'=IF(ABS(E{r})<=1E-9*ABS(D{r}),"OK","ERROR")'); CK.cell(r, 7, "驗證（基準取法）"); r += 1
        for j, y in enumerate((2026, 2027, 2028, 2029)):
            col = "CDEF"[j]
            CK.cell(r, 1, f"K12-{y}"); CK.cell(r, 2, f"三層基準協調 {y}：該年標示（對帳中須有對帳單與期限，見 tools/baseline_coordination/check.py 第 5 項）")
            CK.cell(r, 3, f"=Coord!{col}{COROW['lab']}"); CK.cell(r, 6, "INFO"); CK.cell(r, 7, "資訊（不計入 SGOV）"); r += 1
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
        ("v0.19-trial（試作，不合併）：依三層基準協調規則 v1.1 C3 決定基準：In 頁 M30（換算值、兩層證據比例、cov_done、上次判定）＋ Coord 頁（逐年狀態甲乙丙、由誰定基準、該年標示、依規則的基準與基準係數 crd_k）；基準鏈上的 S33 係數改讀 crd_k（覆蓋未補齊時＝現行係數，數值與 v0.18 相同）。Interface E 節新增供給側／需求側讀數、另一側讀數、狀態與對帳中旗標。", False),
        ("v0.18：2027 基準依 HBM 本庫自算供給截頂（三層校準原則第 2 條；Andy 11-05 同意，條件為本庫供給）：各平台 2027 顆數 × 0.9404（＝可支撐部署 GW 39.16 ÷ 原基準 41.65），不遞延到 2028；2027 部署 GW 41.65 → 39.16、HBM 需求 ÷ 供給 106% → 100%。決策 S48。", False),
        ("v0.17：HBM 供給改本庫自算（Andy 11-05「證據到哪，數字到哪」）：2026 基數＝本庫 2026 HBM 出貨 3.78 EB（三家皆稱售罄；參照 UBS 4.23）；2027 成長＝晶圓 +12% × HBM 晶圓占比 20%→30% × 每片位元 +5% ＝ +76%（0.36–1.19；參照 TrendForce +55%）。2027 供給 6.56 → 6.67 EB、需求 ÷ 供給 108% → 106%。決策 S47。", False),
        ("v0.16.3：Deploy 頁新增實體部署 GW（NVIDIA＝顆數 ÷ 該世代每 GW 顆數；非 NVIDIA＝IT 資本支出 ÷（VR200 每 GW × 持有成本比））與 Interface SIF_DeployGWPhys_Base，供電源、散熱、組裝等按實體電力計的下游讀取（Andy 11-05「先加」；讀者尚未改讀）。成本等值的 SIF_DeployGW_Base 不變。決策 S46。", False),
        ("v0.16.2（只改文字）：CF03、CF11、CF13、CF23 結案（Andy 2026-11-05 15:24；決策 S45），Checks G5 4 → 0。", False),
        ("v0.16.1：M28 改為 NVIDIA 營收口徑（算力收支 11-05 指出晶片內容橋接表第 44 列 GPU 取毛利、其餘取售價，為混合口徑；橋接表 v0.2.3 新增營收口徑列）：nvr_ 基準 GB300 0.647 → 0.772、VR200 0.624／0.581 → 0.695；SIF_NVContentGW_* 隨之改變（例 VR200 21.85 → 26.13）。Recon F 節 NVIDIA 資料中心營收比值 2026／2027 由 1.02／1.04 變為約 1.19／1.25（待診斷）。決策 S44。", False),
        ("v0.16：情境過渡完成（決策 S41）：In 頁層的產業情境 A、C 列（70 列）移到市場層 markets/scen（數值不變），本庫各頁只算基準、Scen 頁只剩基準；Cap 頁分項上限 SIF_CapGW_CoWoS_Base／SIF_CapGW_HBM_Base 與上線口徑 SIF_CapOnlineGW_CoWoS_Base／_HBM_Base（S42）；TK_In 改 Tokenomics v21.10（S43；只有本庫不讀的 NonNV 產出比改值）。唯一數值變動：SIF_TestWFE_SoC 2028–29 改讀基準（無下游讀者）。", False),
        ("v0.15：Interface 推導節的情境 A／B／C 名稱刪除（下游 9 檔改讀 *_Base 或個股自訂：tuc、arm、ase、mpi、smci、dell、jentech、ter、tsmc；Wattonomics 已於 v0.2.2 改讀）；In 頁層情境列名稱保留為個股自訂參考值。數值不變。決策 S40。", False),
        ("v0.14：HBM 供給基數改用廠商出貨推估（M3b hbm_sup26 4.23 EB；原＝本庫 2026 需求）；2027 HBM 需求 ÷ 供給 121% → 108%，可支撐部署 GW 34.4 → 38.5。Andy 2026-11-05 選甲；決策 S39。", False),
        ("v0.13.1：新增 M28（分世代 NVIDIA 合計 ÷ IT capex；GB300、VR200 取自晶片內容橋接表，Hopper、GB200、Rubin Ultra 為 Analogy）與 Interface SIF_NVContentGW_<世代>（＝比例 × Tokenomics 每 IT GW IT capex；Wattonomics Requests R7）、SIF_CapOnlineGW_Base（可支撐新增上線 GW，Wattonomics 可比口徑）。", False),
        ("v0.13：三層校準原則（Andy 2026-11-04：實績雙向校正、供給是上限、錨點可信度分時段看）落地：新增 Cap 頁（基準顆數的 CoWoS、HBM 需求 vs 產能 → 可支撐部署 GW；Checks K10、K10w）、Recon F 節季度實績驗證（NVIDIA 資料中心營收）、Interface SIF_WTAnchor_Base（個股敏感度「Wattonomics 錨定」案例用）、SIF_CapGW_Base、SIF_CapHead_Base、SIF_CoWoSUtil_Base、SIF_HBMUtil_Base。決策 S37。", False),
        ("v0.12：本庫只留基準（工作方式 v1.7；基準＝原情境 B，決策 S33、S35）：Interface 新增 D 節 *_Base 名稱（57 組，等同原 B 列）與 SIF_GWvsWT_Base（本庫基準推得上線 GW 相對 Wattonomics 基準）；A／B／C 名稱標停用過渡、v0.13 刪除；Scen、Units 頁改標；R4 對帳以基準重跑（r3，讀 Wattonomics v0.1.9），新增 NVIDIA 資料中心營收對帳點；個股情境改為個股自訂（決策 S36）。輸入值不變。", False),
        ("v0.11：新增 Recon 頁（由下往上對帳：覆蓋個股基準情境 → 本庫市場；本庫推得上線 GW → Wattonomics；狀態 一致／時差／待診斷，決策 S32、S33）與 M27 對帳參數（容差 rc_tol、SoC 其他廠份額 rc_soc_oth）、Checks K6-soc、K9；標記更正：hbm_hbp_27 高值改 0、u_rb_28 與 u_bw_28 改標 Assumed 並給區間、nd_sh_essd25 改標 Derived、混合鍵合列名改「HBM 工序比例」（決策 S34）。", False),
        ("v0.10：NAND 季價 2026Q3 改用公司實現平均售價 +25%（原合約價 +15%；Kioxia、SK 海力士實績 mid-20%），2026Q4 改為兩季累計不變的餘額 +8.1%（原 +17.5%）；2027 起價格水準不變。決策 S30。", False),
        ("v0.9：CY2026 SoC 測試機市場 ts_mk_soc26 由 11.0 改為 11.5（Advantest 2026-10-28 稱約 115 億美元；Advantest＋泰瑞達 CY2026 SoC 營收合計約 113 億美元已高於舊值）；每加權顆測試支出與 CY2027 起 SoC 市場隨之上升。決策 S29。", False),
        ("v0.8.4：來源複核後修正標記與來源文字：DRAM／NAND 情境季價路徑 2027–28 改標 Derived、2029 改標 Assumed（區間＝情境 A、B）；Advantest 測試機市場來源改為 2026-07-29 原文；輸入值不變。", False),
        ("v0.8.3：衝突登錄清理（CF02、CF05–08、CF16 結案；CF01 移到南電分部重建）；無新增輸入。", False),
        ("v0.8.2：新增 M26（共用匯率：新台幣 32.0、人民幣 6.72、韓元 1,420；台、陸、韓 21 個模型改讀；日圓為公司財測前提，不共用）。", False),
        ("v0.8.1：新增 M25（AI 伺服器組裝：廣達、鴻海、緯創本體、緯穎 FY2026 有效份額快照與 2027–29 份額相對值、緯穎 ASIC 比重；原在 quanta、foxconn、wistron、wiwynn）、ShareTbl 的組裝份額表（含緯穎 ASIC 傾斜係數）、Checks K6-asm、Interface SIF_AsmShareSum、SIF_AsmWyTilt、SIF_AsmShare_wy；CF23 登錄（緯創模型的緯穎份額路徑與緯穎模型不同）。", False),
        ("v0.8：新增 M24（HBM 營收份額：SK 海力士、三星、美光 2025–2029；原在各公司 company.py）、ShareTbl 的 HBM 份額表、Checks K6-hbm、Interface SIF_HBMShareSum；台積電的 HBM 2027 缺口改讀 SIF_HBM_GapEB_A ÷ SIF_HBM_SupplyEB（原手打 21%）。", False),
        ("v0.7.1：新增 M22（HBM 堆疊技術：核心晶粒容量、堆疊層數、混合鍵合占比 2025–2030；原在 disco，hanmi、besi 同讀）、M23（NAND 市場輸入：每顆加速器企業級 SSD 容量、2025 位元基期與應用結構、各應用與供給位元年增、各廠季營收、份額年變動；原在 markets/nand）、Equip 頁 D 節、Interface SIF_HBMHybrid_*。", False),
        ("v0.7：新增 M20（設備市場層的行業輸入：SEMI 年中預測、2029 外推、其他晶圓廠設備成長與 2025 資本支出、設備情境係數、AI 分類、券商 WFE；原在 markets/equipment）、M21（測試設備：SoC／記憶體測試機市場錨點、非 AI SoC、每顆測試時間係數、HBM 權重、泰瑞達測試 ÷ WFE 區間；原在 advantest、ter）與 Equip 頁（SEMI 分項路徑、SoC 測試機市場三情境、測試 ÷ WFE〔CF15〕）、Checks K8、Interface SIF_SEMI_*、SIF_TestSoCMkt_*、SIF_TestWFE_*。", False),
        ("v0.6.1：新增 M19（液冷：機架內液冷每架內容〔原在 markets/chip_bridge〕、分歧管、內容調整、滲透率、TrendForce 對照、奇鋐與雙鴻 FY2026 份額快照與份額相對值、健策冷板元件份額；原在 avc、auras、jentech）、Rack 頁 D 節（AI 液冷市場三情境）、ShareTbl 的 AI 液冷份額表（CF04 加入雙鴻；CF12 台達電機架外液冷不屬本市場）、Checks K6-lc、Interface SIF_LiqPerGW、SIF_LiqPen_*、SIF_LiqMkt_*、SIF_LiqShareSum_*。", False),
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
