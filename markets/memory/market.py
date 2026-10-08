# 記憶體市場層（共用）：季度價格路徑（情境 A／B／C）、加速器數量層、HBM 市場。
# 各公司模型（tools/memory_model/build.py）引用本檔，確保所有記憶體公司用同一組市場假設。
QUARTERS = [f"{y}Q{q}" for y in (2025, 2026, 2027, 2028) for q in (1, 2, 3, 4)]  # Path!C..R
SCEN = ("A", "B", "C")
SCEN_DESC = {"A": "2027–28 高檔持平（BofA、TrendForce 型）",
             "B": "2027 Q2 見頂、2028 回落約一半（Citi、Bernstein 型）",
             "C": "2028 回到 2025 年價格水準（完整谷底；歷史循環型）"}

def qcol(i):  # 0..15 → Path 欄
    return chr(ord("C") + i)

# 價格季增（QoQ）：2025Q2–2026Q4 共用；2027Q1–2028Q4 依情境
CONV_COMMON = [-0.025, 0.125, 0.475, 0.95, 0.605, 0.155, 0.125]
CONV_SRC = "TrendForce 合約價：2Q25 0~−5%、3Q25 +10~15%、4Q25 實際 +45~50%、1Q26 實際約 +93~98%、2Q26 +58~63%、3Q26 +13~18%（MS 實際約 +15%）、4Q26 +10~15%"
CONV_S = {"A": [0.03, 0.0, -0.02, -0.03, -0.03, -0.03, -0.02, -0.02],
          "B": [0.05, 0.02, -0.08, -0.15, -0.25, -0.15, -0.10, -0.05],
          "C": [0.05, 0.02, -0.08, -0.15, -0.35, -0.30, -0.25, -0.20]}
CONV_S_SRC = {"A": "BofA：2027 持平略降、2027 均價仍高於 2026；2028 溫和修正（年均約 −10%）",
              "B": "Citi：2027 Q2 見頂；Bernstein：2028 DRAM $/GB −52.9%（本路徑 2028 年均約 −47%）",
              "C": "Assumed：2028 年底回到約 2025Q4 水準（自高點約 −79%），比照 2019、2023 年谷底"}
NAND_COMMON = [0.0, 0.05, 0.30, 0.875, 0.725, 0.15, 0.175]
NAND_SRC = "TrendForce：1Q26 +85~90%、2Q26 +70~75%、3Q26 +10~15%（MS 實際約 +20%）、4Q26 +15~20%；2025Q2–Q4 為 Assumed"
NAND_S = {"A": [0.03, 0.0, -0.05, -0.07, -0.05, -0.05, -0.05, -0.05],
          "B": [0.05, 0.02, -0.10, -0.20, -0.30, -0.25, -0.20, -0.10],
          "C": [0.05, 0.02, -0.10, -0.20, -0.35, -0.35, -0.25, -0.20]}
NAND_S_SRC = {"A": "TrendForce：2027 下半年供需平衡、價格面臨下修壓力",
              "B": "Citi：2027 Q2 見頂；Bernstein：NAND $/GB 2028 −68.8%",
              "C": "Assumed：回到約 2025 年水準"}

UNITS = [
 ("bw", "NVIDIA Blackwell（B200／B300）", (5.4, 2.4, 0.5), (260, 280, 288), "Interested-party", "顆數：Morgan Stanley 2026-07（2026 5.4M）；AC Research 2027 約 2.4M；2028 Assumed。GB：B200 186、B300 288 組合 [Derived]"),
 ("rb", "NVIDIA Rubin", (1.7, 7.0, 7.0), (288, 288, 288), "Interested-party", "2026：TrendForce 組合 Rubin 22% × 約 7.6M [Derived]；2027：Morgan Stanley 約 7.0M；2028 Assumed"),
 ("ru", "NVIDIA Rubin Ultra", (0.0, 0.3, 2.0), (1024, 1024, 1024), "Assumed", "時程 Assumed；GB：NVIDIA 1 TB／封裝"),
 ("hp", "NVIDIA Hopper（H200 等）", (0.5, 0.0, 0.0), (141, 141, 141), "Interested-party", "TrendForce 2026 組合 Hopper 7%"),
 ("amd", "AMD Instinct", (0.75, 1.5, 2.0), (288, 360, 432), "Interested-party／Assumed", "2027：Morgan Stanley MI455X 1.0M＋MI450 0.5M；其餘 Assumed"),
 ("tpu", "Google TPU", (3.85, 5.5, 7.0), (192, 250, 288), "Interested-party", "MS 3.2／5.0、GF 4.5／8.84（2026／2027）取中間偏 MS；2028 Assumed"),
 ("trn", "AWS Trainium", (2.0, 2.5, 3.0), (110, 144, 200), "Assumed", "年出貨未找到；GB：Trn2 96、Trn3 144"),
 ("oth", "其他 ASIC（MTIA、Maia、Ascend 等）", (1.0, 4.5, 5.0), (150, 200, 250), "Derived／Assumed", "2027：JPMorgan ASIC 12.5M − TPU − Trainium"),
]

def inputs(sec, inp):
    sec("M1. 記憶體價格路徑（季增，共用市場層 markets/memory）")
    for i, v in enumerate(CONV_COMMON):
        inp(f"cq_{i+1}", f"一般 DRAM 合約價季增 {QUARTERS[i+1]}", "%", None, v, None, "Interested-party", CONV_SRC if i == 0 else "同上")
    for s in SCEN:
        for i, v in enumerate(CONV_S[s]):
            inp(f"cq{s}_{i+8}", f"一般 DRAM 合約價季增 {QUARTERS[i+8]}（{s}）", "%", None, v, None, "Interested-party" if s != "C" else "Assumed", CONV_S_SRC[s] if i == 0 else "同上")
    for i, v in enumerate(NAND_COMMON):
        inp(f"nq_{i+1}", f"NAND 合約價季增 {QUARTERS[i+1]}", "%", None, v, None, "Interested-party", NAND_SRC if i == 0 else "同上")
    for s in SCEN:
        for i, v in enumerate(NAND_S[s]):
            inp(f"nq{s}_{i+8}", f"NAND 合約價季增 {QUARTERS[i+8]}（{s}）", "%", None, v, None, "Interested-party" if s != "C" else "Assumed", NAND_S_SRC[s] if i == 0 else "同上")
    sec("M2. 加速器數量層（百萬顆）與每顆 HBM（GB）")
    for k, nm, u, gb, tag, src in UNITS:
        for i, y in enumerate((26, 27, 28)):
            inp(f"u_{k}_{y}", f"{nm} 顆數 20{y}", "百萬顆", None, u[i], None, tag, src if i == 0 else "同上")
        for i, y in enumerate((26, 27, 28)):
            inp(f"g_{k}_{y}", f"{nm} 每顆 HBM 20{y}", "GB", None, gb[i], None, tag, "同上")
    sec("M3. HBM 市場")
    inp("hbm_mkt25", "HBM 市場 2025", "$B", 34, 35, 35, "Verified", "Micron TAM 約 $35B（2025-12 法說）；Yole 約 $34B")
    inp("hbm_g26", "2026 HBM 位元成長（推 2025 出貨位元）", "%", 0.6, 0.7, 0.8, "Interested-party", "TrendForce：2026 >70%")
    inp("hbm_gross", "出貨位元 ÷ 加速器裝載位元", "x", 1.0, 1.15, 1.3, "Assumed", "封裝良率損失、庫存與未列入的買方（Epoch：CoWoS-L 良率 65–95%）")
    inp("hbm_sup27", "2027 HBM 位元供給成長", "%", 0.50, 0.55, 0.63, "Interested-party", "TrendForce 2027 +50–60%；JPMorgan 2026–28 CAGR 63%")
    inp("hbm_sup28", "2028 HBM 位元供給成長", "%", 0.30, 0.45, 0.63, "Assumed", "新廠 2028 才貢獻（TrendForce）")
    inp("hbm_p26", "HBM 混合單價 2026", "$/GB", 12.5, 14.6, 16.0, "Interested-party", "BofA $14.6/GB；DigiTimes HBM3E $12–12.8、HBM4 約 $16")
    inp("hbm_p27", "HBM 單價 2027 變動（三情境相同；年約已議定）", "%", 0.38, 0.80, 1.21, "Interested-party", "BofA +38%、JPM +54%、UBS +76–79%、TrendForce +121%；Micron：CY27 大部分已簽、大幅漲價")
    inp("hbm_p28_A", "HBM 單價 2028 變動（A）", "%", 0.0, 0.10, 0.25, "Interested-party", "JPMorgan 2028 +25%（上限）")
    inp("hbm_p28_B", "HBM 單價 2028 變動（B）", "%", -0.40, -0.25, -0.10, "Assumed", "一般 DRAM 大跌時 HBM 年約重議下修")
    inp("hbm_p28_C", "HBM 單價 2028 變動（C）", "%", -0.55, -0.40, -0.25, "Assumed", "完整谷底")

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
            for i in range(1, 16):
                key = f"{cpre}_{i}" if i <= 7 else f"{cpre}{s}_{i}"
                P.cell(r, 3 + i, f"={qcol(i-1)}{r}*(1+{key})").number_format = "0.00"
            ref[kind][s] = r
            r += 1
        r += 1
    P.cell(r, 1, "年均指數").font = BOLD; r += 1
    hdr(P, r, ["路徑", "情境", "2025", "2026", "2027", "2028", "2026÷2025", "2027÷2026", "2028÷2027"]); r += 1
    for kind, npre in (("conv", "一般 DRAM"), ("nand", "NAND")):
        for s in SCEN:
            src = ref[kind][s]
            P.cell(r, 1, npre); P.cell(r, 2, s)
            for j in range(4):
                P.cell(r, 3 + j, f"=AVERAGE({qcol(4*j)}{src}:{qcol(4*j+3)}{src})").number_format = "0.00"
            for j, (a, b) in enumerate((("D", "C"), ("E", "D"), ("F", "E"))):
                P.cell(r, 7 + j, f"={a}{r}/{b}{r}-1").number_format = "0%"
            r += 1
    P.column_dimensions["A"].width = 14
    for i in range(16): P.column_dimensions[qcol(i)].width = 8

    M = wb.create_sheet("Mkt")
    M["A1"] = "Mkt — 加速器數量層（最小版）與 HBM 市場（日曆年）。需求＝顆數 × GB；出貨＝MIN(需求, 供給)；市場＝出貨 × 單價"
    hdr(M, 3, ["項目", "單位", "2025", "2026", "2027", "2028", "說明"])
    MK = {}; mr = [4]
    def put(label, unit, fs, key=None, fmt="0.00", note=""):
        rr = mr[0]
        M.cell(rr, 1, label); M.cell(rr, 2, unit)
        for j, f in enumerate(fs):
            if f is None or f == "": continue
            M.cell(rr, 3 + j, f).number_format = fmt
        M.cell(rr, 7, note)
        if key: MK[key] = rr
        mr[0] += 1
    YC = {25: "C", 26: "D", 27: "E", 28: "F"}
    M.cell(mr[0], 1, "需求：加速器裝載的 HBM（PB）").font = BOLD; mr[0] += 1
    for k, nm, *_ in UNITS:
        put(f"  {nm}", "PB", [None] + [f"=u_{k}_{y}*g_{k}_{y}" for y in (26, 27, 28)], key="pb_" + k, fmt="0")
    a, b = MK["pb_bw"], MK["pb_oth"]
    put("需求合計（裝載）", "EB", [None] + [f"=SUM({YC[y]}{a}:{YC[y]}{b})/1000" for y in (26, 27, 28)], key="dem")
    put("需求 × 出貨係數", "EB", [None] + [f"={YC[y]}{MK['dem']}*hbm_gross" for y in (26, 27, 28)], key="demg")
    put("加速器顆數合計", "百萬顆", [None] + ["=" + "+".join(f"u_{k}_{y}" for k, *_ in UNITS) for y in (26, 27, 28)], key="units", fmt="0.0", note="JPMorgan 2027：GPU 10.9M、ASIC 12.5M")
    put("每顆平均 HBM", "GB", [None] + [f"={YC[y]}{MK['dem']}*1000/{YC[y]}{MK['units']}" for y in (26, 27, 28)], key="gbavg", fmt="0")
    put("供給", "EB", ["=D{s}/(1+hbm_g26)", "=D{d}", "=D{s}*(1+hbm_sup27)", "=E{s}*(1+hbm_sup28)"], key="sup", note="2026 供給＝需求出貨（年中已成交）")
    rr = MK["sup"]
    for j, f in enumerate(["=D{0}/(1+hbm_g26)".format(rr), "=D{0}".format(MK["demg"]), "=D{0}*(1+hbm_sup27)".format(rr), "=E{0}*(1+hbm_sup28)".format(rr)]):
        M.cell(rr, 3 + j, f).number_format = "0.00"
    put("出貨＝MIN(需求, 供給)", "EB", [f"=C{MK['sup']}"] + [f"=MIN({YC[y]}{MK['demg']},{YC[y]}{MK['sup']})" for y in (26, 27, 28)], key="ship")
    put("需求 − 供給（>0＝短缺）", "EB", [None] + [f"={YC[y]}{MK['demg']}-{YC[y]}{MK['sup']}" for y in (26, 27, 28)], key="gap")
    for s in SCEN:
        put(f"HBM 單價（{s}）", "$/GB", [f"=hbm_mkt25/C{MK['ship']}", "=hbm_p26", "=D{r}*(1+hbm_p27)", f"=E{{r}}*(1+hbm_p28_{s})"], key="p" + s)
        rr = MK["p" + s]
        M.cell(rr, 5, f"=D{rr}*(1+hbm_p27)"); M.cell(rr, 6, f"=E{rr}*(1+hbm_p28_{s})")
    for s in SCEN:
        put(f"HBM 市場（{s}）", "$B", [f"={YC[y]}{MK['ship']}*{YC[y]}{MK['p' + s]}" for y in (25, 26, 27, 28)], key="m" + s, fmt="0.0")
    mr[0] += 1
    M.cell(mr[0], 1, "對帳（樣板驗收 2）").font = BOLD; mr[0] += 1
    put("由下而上：NVIDIA 裝載 HBM × 單價（2026）", "$B", [None, f"=(D{MK['pb_bw']}+D{MK['pb_rb']}+D{MK['pb_ru']}+D{MK['pb_hp']})/1000*hbm_p26"], key="bu", fmt="0.0")
    put("由上而下：NVIDIA 資料中心營收 × 每顆 HBM ÷ 每顆 NVIDIA 內容（2026）", "$B", [None, "=nv_dc26*nv_hbm/nv_content"], key="td", fmt="0.0", note="橋接表 v0.1；與顆數預測獨立")
    put("  差距", "%", [None, f"=D{MK['bu']}/D{MK['td']}-1"], fmt="0%")
    put("由營收推得 NVIDIA 顆數 vs 顆數預測", "百萬顆", [None, "=nv_dc26/nv_content", "=u_bw_26+u_rb_26+u_hp_26"], fmt="0.0", note="D＝營收推得；E＝顆數預測（2026）")
    put("已發布：SK 海力士自估 2026／BofA 2026／BofA 2027", "$B", [None, "=pub26_lo", "=pub26_hi", "=pub27"], key="pub", fmt="0.0")
    put("  本模型 2026 ÷ SK 海力士自估、2027 ÷ BofA", "%", [None, f"=D{MK['mA']}/pub26_lo-1", None, f"=E{MK['mA']}/pub27-1"], fmt="0%")
    M.column_dimensions["A"].width = 58
    for c in "CDEF": M.column_dimensions[c].width = 10
    M.column_dimensions["G"].width = 50
    return {"path": ref, "mk": MK}

def ref_inputs(sec, inp):
    sec("M4. 對帳用（橋接表 v0.1、已發布）")
    inp("nv_dc26", "NVIDIA 資料中心營收 2026（約 FY27）", "$B", 340, 374, 400, "Verified／Assumed", "FY27 Q1 75.2、Q2 89.0（8-K）；Q3 指引總營收 108（資料中心取 100）；Q4 Assumed 110")
    inp("nv_content", "每顆 GPU 的 NVIDIA 內容（GB300）", "$K", 45, 51.6, 58, "Derived", "markets/chip_bridge v0.1")
    inp("nv_hbm", "每顆 GPU HBM 金額（GB300）", "$K", 3.5, 4.46, 4.9, "Derived", "markets/chip_bridge v0.1")
    inp("pub26_lo", "已發布 2026 HBM 市場（SK 海力士自估）", "$B", None, 54.6, None, "Interested-party", "Motley Fool／글로벌이코노믹 2026-09")
    inp("pub26_hi", "已發布 2026 HBM 市場（BofA）", "$B", None, 77.4, None, "Interested-party", "BofA 2026-08-07")
    inp("pub27", "已發布 2027 HBM 市場（BofA）", "$B", None, 152.8, None, "Interested-party", "BofA；JPMorgan 2027–28 為 160–282")
