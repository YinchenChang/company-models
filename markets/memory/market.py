# 記憶體市場層（共用）：季度價格路徑（情境 A／B／C）、加速器數量層、HBM 市場。
# 各公司模型（tools/memory_model/build.py 等 38 個建檔程式）引用本檔，確保所有公司用同一組市場假設。
# 2026-10-27 起（Siliconomics v0.1）：本檔不再持有任何數值。所有輸入（季價、顆數、每顆 HBM、情境係數、HBM 市場參數、對帳參考）
# 由 AI 半導體產業基準庫 Siliconomics 的 Excel 擁有，經 siliconomics/tools/export_snapshot.py 匯成快照；本檔只讀快照、
# 以原有的 Python 名稱（UNITS、SCEN_UNITS、CONV_* …）提供給下游，並保留 write() 的公式結構（Siliconomics 的 Path／Mkt 頁也用同一份）。
# 要改數值：改 Siliconomics Excel 的 In 頁 → 重建、重算 → 重新匯出快照 → 更新 siliconomics/data/CURRENT_SNAPSHOT。
import json, os

_SIL = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "siliconomics", "data")
SNAPSHOT_FILE = open(os.path.join(_SIL, "CURRENT_SNAPSHOT"), encoding="utf-8").read().strip()
SNAP = json.load(open(os.path.join(_SIL, SNAPSHOT_FILE), encoding="utf-8"))
SIL_VERSION = SNAP["source"]["version"]
_ROWS = SNAP["rows"]
# Excel 不分整數與浮點數；下游有把數值寫進說明文字者（例如「× 1.0／1.0」），故依單位還原原本的型別：
# 比例、倍數、百萬顆、單價一律浮點數；GB、$B、$K 為整數值時用整數（與 2026-10-27 前 market.py 的寫法相同）
_FLOAT_UNITS = {"%", "x", "百萬顆", "$/GB"}
for _r in _ROWS:
    if "id" in _r:
        for _f in ("lo", "base", "hi"):
            _x = _r[_f]
            if isinstance(_x, (int, float)) and not isinstance(_x, bool):
                _r[_f] = float(_x) if _r["unit"] in _FLOAT_UNITS else (int(_x) if float(_x).is_integer() else _x)
_V = {r["id"]: r for r in _ROWS if "id" in r}

QUARTERS = [f"{y}Q{q}" for y in (2025, 2026, 2027, 2028, 2029) for q in (1, 2, 3, 4)]  # Path!C..V（結構，非數值）
NQ = len(QUARTERS)
SCEN = tuple(SNAP["scen"])
SCEN_DESC = dict(SNAP["scen_desc"])


def _b(k):
    return _V[k]["base"]


# 加速器顆數的情境係數（2028、2029 年 ÷ 情境 A 基準）：tsmc、advantest 等同一套，由 Siliconomics 單一來源提供
SCEN_UNITS = {s: (_b(f"um28_{s}"), _b(f"um29_{s}")) for s in SCEN}
SCEN_UNITS_RNG = {s: ((_V[f"um28_{s}"]["lo"], _V[f"um28_{s}"]["hi"]), (_V[f"um29_{s}"]["lo"], _V[f"um29_{s}"]["hi"]))
                  for s in SCEN if _V[f"um28_{s}"]["lo"] is not None}


def qcol(i):  # 0..15 → Path 欄
    return chr(ord("C") + i)


CONV_COMMON = [_b(f"cq_{i}") for i in range(1, 8)]
CONV_SRC = _V["cq_1"]["src"]
CONV_S = {s: [_b(f"cq{s}_{i}") for i in range(8, 20)] for s in SCEN}
CONV_S_SRC = {s: _V[f"cq{s}_8"]["src"] for s in SCEN}
NAND_COMMON = [_b(f"nq_{i}") for i in range(1, 8)]
NAND_SRC = _V["nq_1"]["src"]
NAND_S = {s: [_b(f"nq{s}_{i}") for i in range(8, 20)] for s in SCEN}
NAND_S_SRC = {s: _V[f"nq{s}_8"]["src"] for s in SCEN}

UNITS = [(g["key"], g["name"], tuple(_b(f"u_{g['key']}_{y}") for y in (26, 27, 28)),
          tuple(_b(f"g_{g['key']}_{y}") for y in (26, 27, 28)), g["tag"], g["src"]) for g in SNAP["unit_groups"]]

# 2029 平台別顆數成長（Siliconomics v0.2 起單一來源；tsmc、network 讀這裡）與合計年增（nand 讀這裡）
G29 = {g["key"]: _V[f"g29_{g['key']}"] for g in SNAP["unit_groups"] if f"g29_{g['key']}" in _V}
G_U29 = SNAP["names"].get("SIF_UnitsGrowth29")

# 各記憶體公司模型 In 頁重播的節：inputs() 只放 M1–M3（市場層），ref_inputs() 只放 M4（對帳參考）；
# M5 起的節（例如 2029 平台別成長）不寫進記憶體公司模型，由需要的模型經上面的名稱取用
_INPUT_SECS = ("M1.", "M2.", "M2b.", "M3.")
_REF_SECS = ("M4.",)


def _replay(sec, inp, want_ref):
    on = False
    want = _REF_SECS if want_ref else _INPUT_SECS
    for r in _ROWS:
        if "sec" in r:
            on = r["sec"].startswith(want)
            if on:
                sec(r["sec"])
            continue
        if on:
            inp(r["id"], r["label"], r["unit"], r["lo"], r["base"], r["hi"], r["tag"], r["src"])


def inputs(sec, inp):
    _replay(sec, inp, False)

def write(wb, hdr, Font):
    """寫 Path、Mkt 兩頁；回傳參照表"""
    BOLD = Font(bold=True)
    P = wb.create_sheet("Path")
    P["A1"] = "Path — 記憶體合約價指數（2025Q1＝1；季增來自 In 頁 M1）。共用市場層：所有記憶體公司用同一條路徑，再依各自會計年度取平均"
    hdr(P, 3, ["路徑", "情境"] + QUARTERS)
    ref = {"conv": {}, "nand": {}}
    r = 4
    for kind, cpre, npre in (("conv", "cq", "一般 DRAM"), ("nand", "nq", "NAND")):
        for s in SCEN:
            P.cell(r, 1, npre); P.cell(r, 2, s)
            P.cell(r, 3, 1)
            for i in range(1, NQ):
                key = f"{cpre}_{i}" if i <= 7 else f"{cpre}{s}_{i}"
                P.cell(r, 3 + i, f"={qcol(i-1)}{r}*(1+{key})").number_format = "0.00"
            ref[kind][s] = r
            r += 1
        r += 1
    P.cell(r, 1, "年均指數").font = BOLD; r += 1
    hdr(P, r, ["路徑", "情境", "2025", "2026", "2027", "2028", "2029", "2026÷2025", "2027÷2026", "2028÷2027", "2029÷2028"]); r += 1
    for kind, npre in (("conv", "一般 DRAM"), ("nand", "NAND")):
        for s in SCEN:
            src = ref[kind][s]
            P.cell(r, 1, npre); P.cell(r, 2, s)
            for j in range(5):
                P.cell(r, 3 + j, f"=AVERAGE({qcol(4*j)}{src}:{qcol(4*j+3)}{src})").number_format = "0.00"
            for j, (a, b) in enumerate((("D", "C"), ("E", "D"), ("F", "E"), ("G", "F"))):
                P.cell(r, 8 + j, f"={a}{r}/{b}{r}-1").number_format = "0%"
            r += 1
    P.column_dimensions["A"].width = 14
    for i in range(NQ): P.column_dimensions[qcol(i)].width = 8

    M = wb.create_sheet("Mkt")
    M["A1"] = "Mkt — 加速器數量層（最小版）與 HBM 市場（日曆年）。需求＝顆數 × GB × 情境顆數係數（2028–29）；供給三情境相同；出貨＝MIN(需求, 供給)；市場＝出貨 × 單價（單價依情境）"
    hdr(M, 3, ["項目", "單位", "2025", "2026", "2027", "2028", "2029", "說明"])
    MK = {}; mr = [4]
    def put(label, unit, fs, key=None, fmt="0.00", note=""):
        rr = mr[0]
        M.cell(rr, 1, label); M.cell(rr, 2, unit)
        for j, f in enumerate(fs):
            if f is None or f == "": continue
            M.cell(rr, 3 + j, f).number_format = fmt
        M.cell(rr, 8, note)
        if key: MK[key] = rr
        mr[0] += 1
    YC = {25: "C", 26: "D", 27: "E", 28: "F", 29: "G"}
    M.cell(mr[0], 1, "需求：加速器裝載的 HBM（PB）").font = BOLD; mr[0] += 1
    for k, nm, *_ in UNITS:
        put(f"  {nm}", "PB", [None] + [f"=u_{k}_{y}*g_{k}_{y}" for y in (26, 27, 28)], key="pb_" + k, fmt="0")
    a, b = MK["pb_bw"], MK["pb_oth"]
    put("需求合計（裝載）", "EB", [None] + [f"=SUM({YC[y]}{a}:{YC[y]}{b})/1000" for y in (26, 27, 28)], key="dem")
    M.cell(MK["dem"], 7, f"=F{MK['dem']}*(1+hbm_dem29)").number_format = "0.00"
    put("需求 × 出貨係數（基準＝情境 A）", "EB", [None] + [f"={YC[y]}{MK['dem']}*hbm_gross" for y in (26, 27, 28, 29)], key="demg")
    put("加速器顆數合計", "百萬顆", [None] + ["=" + "+".join(f"u_{k}_{y}" for k, *_ in UNITS) for y in (26, 27, 28)], key="units", fmt="0.0", note="JPMorgan 2027：GPU 10.9M、ASIC 12.5M；情境 A 基準")
    put("每顆平均 HBM", "GB", [None] + [f"={YC[y]}{MK['dem']}*1000/{YC[y]}{MK['units']}" for y in (26, 27, 28)], key="gbavg", fmt="0")
    put("供給（三情境相同）", "EB", [None] * 5, key="sup", note="2026 供給＝需求出貨（年中已成交）；供給不隨需求情境變（產能 2027 前已訂定）")
    rr = MK["sup"]
    for j, f in enumerate(["=D{0}/(1+hbm_g26)".format(rr), "=D{0}".format(MK["demg"]), "=D{0}*(1+hbm_sup27)".format(rr), "=E{0}*(1+hbm_sup28)".format(rr), "=F{0}*(1+hbm_sup29)".format(rr)]):
        M.cell(rr, 3 + j, f).number_format = "0.00"
    # 情境層：顆數係數 → 需求 → 出貨 → 缺口（供給不變，需求隨情境）
    for s in SCEN:
        put(f"顆數係數（{s}；2028–29 ÷ 基準，與台積電、Advantest 同一套）", "x", [None, 1, 1, f"=um28_{s}", f"=um29_{s}"], key="cf" + s, fmt="0.00", note="2026–27 無情境差異")
    for s in SCEN:
        put(f"需求 × 出貨係數（{s}）", "EB", [None] + [f"={YC[y]}{MK['demg']}*{YC[y]}{MK['cf' + s]}" for y in (26, 27, 28, 29)], key="demg" + s)
    for s in SCEN:
        put(f"出貨＝MIN(需求, 供給)（{s}）", "EB", [f"=C{MK['sup']}"] + [f"=MIN({YC[y]}{MK['demg' + s]},{YC[y]}{MK['sup']})" for y in (26, 27, 28, 29)], key="ship" + s,
            note="需求低於供給時出貨＝需求（閒置產能）" if s == "A" else "")
    for s in SCEN:
        put(f"需求 − 供給（{s}；>0＝短缺）", "EB", [None] + [f"={YC[y]}{MK['demg' + s]}-{YC[y]}{MK['sup']}" for y in (26, 27, 28, 29)], key="gap" + s)
    MK["ship"], MK["gap"] = MK["shipA"], MK["gapA"]   # 舊鍵：情境 A（2026–27 三情境相同）
    for s in SCEN:
        put(f"HBM 單價（{s}）", "$/GB", [f"=hbm_mkt25/C{MK['shipA']}", "=hbm_p26", "=D{r}*(1+hbm_p27)", f"=E{{r}}*(1+hbm_p28_{s})"], key="p" + s)
        rr = MK["p" + s]
        M.cell(rr, 5, f"=D{rr}*(1+hbm_p27)"); M.cell(rr, 6, f"=E{rr}*(1+hbm_p28_{s})"); M.cell(rr, 7, f"=F{rr}*(1+hbm_p29_{s})")
    for s in SCEN:
        put(f"HBM 市場（{s}）", "$B", [f"={YC[y]}{MK['ship' + s]}*{YC[y]}{MK['p' + s]}" for y in (25, 26, 27, 28, 29)], key="m" + s, fmt="0.0")
    mr[0] += 1
    M.cell(mr[0], 1, "對帳（樣板驗收 2）").font = BOLD; mr[0] += 1
    put("由下而上：NVIDIA 裝載 HBM × 單價（2026）", "$B", [None, f"=(D{MK['pb_bw']}+D{MK['pb_rb']}+D{MK['pb_ru']}+D{MK['pb_hp']})/1000*hbm_p26"], key="bu", fmt="0.0")
    put("由上而下：NVIDIA 資料中心營收 × 每顆 HBM ÷ 每顆 NVIDIA 內容（2026）", "$B", [None, "=nv_dc26*nv_hbm/nv_content"], key="td", fmt="0.0", note="橋接表 v0.2；與顆數預測獨立")
    put("  差距", "%", [None, f"=D{MK['bu']}/D{MK['td']}-1"], fmt="0%")
    put("由營收推得 NVIDIA 顆數 vs 顆數預測", "百萬顆", [None, "=nv_dc26/nv_content", "=u_bw_26+u_rb_26+u_hp_26"], fmt="0.0", note="D＝營收推得；E＝顆數預測（2026）")
    put("已發布：SK 海力士自估 2026／BofA 2026／BofA 2027", "$B", [None, "=pub26_lo", "=pub26_hi", "=pub27"], key="pub", fmt="0.0")
    put("  本模型 2026 ÷ SK 海力士自估、2027 ÷ BofA", "%", [None, f"=D{MK['mA']}/pub26_lo-1", None, f"=E{MK['mA']}/pub27-1"], fmt="0%")
    M.column_dimensions["A"].width = 58
    for c in "CDEFG": M.column_dimensions[c].width = 10
    M.column_dimensions["H"].width = 50
    return {"path": ref, "mk": MK}


def ref_inputs(sec, inp):
    _replay(sec, inp, True)
