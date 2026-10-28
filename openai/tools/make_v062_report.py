#!/usr/bin/env python3
"""v0.6.2 對照 Excel：①TK 名稱前後值 ②主要輸出前後值與歸因 ③已套用的預設 ④Tokenomics 缺口與對照讀法。
輸入：tools/tk_attribution.py 的 JSON（LibreOffice 重算結果）、v0.6.1／v0.6.2 活頁簿的 TK_Link。差額、百分比與占比為本檔內的 Excel 公式。
用法：python3 tools/make_v062_report.py --attr <json> --old model/archive/<v0.6.1>.xlsx --new model/<v0.6.2>.xlsx --out docs/reports/20261022_v0.6.2_TK同步_對照.xlsx"""
import argparse
import shutil
import json
import sys
import tempfile
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "builder"))
from common import lo_recalc  # noqa: E402

YEARS = ["FY2025", "FY2026", "FY2027", "FY2028", "FY2029", "FY2030"]
GENS = ["Hopper", "GB200", "GB300", "VR200", "Rubin Ultra"]
GLAB = {"1": "①IF_TokGW_*", "2": "②IF_HoldEcon", "3": "③IF_CapexTotal", "4": "④其餘對照類"}
# 輸出區塊：(名稱, 標題, 數字格式, 是否加總歸因)
BLOCKS = [
    ("CMP_Eta", "η（有效推論 GW 校準係數；各年相同）", "0.0000"),
    ("CMP_InfGW_Eff", "有效推論 GW（未截頂）", "0.000"),
    ("CMP_InfGW", "推論 GW（截頂後）", "0.000"),
    ("CMP_CapFactor", "截頂係數（1＝未受限）", "0.000"),
    ("CMP_RDGW", "研發 GW（殘差）", "0.000"),
    ("CMP_DemandGW", "總需求 GW", "0.000"),
    ("CMP_Supply_Owned", "自建 GW", "0.000"),
    ("CMP_VReqFactor", "機隊 VR 等值係數", "0.0000"),
    ("COST_SupplierHold", "供應商持有成本 $B", "0.00"),
    ("COST_CloudGM", "雲端毛利 $B", "0.00"),
    ("COST_CloudGMPct", "雲端毛利 %", "0.0%"),
    ("COST_PropGap_VR", "命題：每 VR 等值 GW 差額（現金口徑）$B", "0.00"),
    ("COST_PropGapEcon_VR", "命題：每 VR 等值 GW 差額（經濟口徑）$B", "0.00"),
    ("COST_Coverage", "命題：營收覆蓋全成本比（現金口徑）", "0.0%"),
    ("COST_RDCompute", "研發算力成本 $B", "0.00"),
    ("COST_Q3", "Q3 研發占整體營運成本", "0.0%"),
    ("COST_Q3GapTK", "Q3 研發算力 − TK L1_Ans3 差距 $B", "0.00"),
    ("FND_ExtNeedCum", "累計外部資金需求 $B（2026–2030 欄位為 FY2026–FY2030）", "0.0"),
    ("PXQ_Ratio", "PxQ 回本比值（欄 FY2025–FY2029 位置＝2025 指數／2025 線性／2026 低／基準／高 GW；前後相同，不納入歸因）", "0.000"),
]
DEFAULTS = [
    ("Tokenomics 提交", "取工作單開工時 master HEAD `1916a06`（#144 合併）的 model/CURRENT＝v19.0；檔案 SHA-256 與 v0.6.1 的 TK_PxQ 相同（d1728705…），故與工作單寫的 `ad33e28` 同一份檔案", "ad33e28（較早提交）", "無（同一檔案、同一 SHA-256）"),
    ("L1_Ans3 是否改連 IF_", "v19.0 Interface 查無『研發用算力成本 $B/年』等義的 IF_ 名稱（IF_ProgCost_* 為每模型層級研發計畫成本 $M，非 2025 穩態年度合計；IF_AllocQ2 為占比），保留 L1_Ans3（下游契約第 15 條允許）", "改連 IF_ProgCost_*（口徑不同）", "Q3 對照倍數 10.4→6.3（2025）；列 Tokenomics 缺口"),
    ("研發占比區間名稱", "新增 RNG_IF_AllocQ1_Lo／_Hi 與 RNG_StaleFlag 到 TK_Link（工程類，只供檢查頁並列）；新增檢查 C163 驗 RNG_StaleFlag＝『一致』", "只用點值 0.184", "無（info 列）"),
    ("兩個快照合一", "TKQ_ 前綴退役；PxQ 公式改引 TK_；IF_HoldEcon、SRC_DEM_004／006–009 原在 63 名內，併入時不重複；新增 22 個取數名稱＋3 個 RNG_ 名稱（共 88 名）", "保留兩份快照", "PxQ 776 個數值格與 v0.6.1 逐格相同（容許誤差 1e-12）"),
    ("builder 參數", "移除 --tk-pxq-dir（不保留可選參數，避免兩個版本並存）", "改為可選並註明停用", "無"),
    ("Checks 編號", "標籤改變的檢查退役舊號、取新號（退役不重用）：C132–C138（TK_PxQ 專用）退役；新增 C155（營收兩來源，標籤改 TK_）、C156–C163", "沿用舊號並改標籤", "無（編號規則）"),
    ("TK_Link 讀取日", "整份重讀，讀取日 2026-10-22", "—", "無"),
    ("歸因分組", "①IF_TokGW_*（3 名）②IF_HoldEcon（1 名）③IF_CapexTotal（1 名）④其餘兩版皆有且值不同者（25 名）；每組各算『只換這組』與『只還原這組』兩次，交互項＝全換 − Σ單換", "其他分組", "—"),
    ("歸因不含 PxQ", "PxQ 在 v0.6.1 本來就用 v19.0 的 IF_HoldEcon；單一快照後『還原 IF_HoldEcon』也會動 PxQ，故 PxQ 只列前後值（不變）並不納入歸因", "—", "—"),
    ("容量截頂", "v19.0 下各年截頂係數仍為 1.000（未觸發），照模型既有機制", "—", "—"),
    ("HTML 敘述", "『2027 起研發占總需求 6–9 成』改為『5–9 成』（2030 年降至 52.7%），test_html 同步", "不改文字", "文字與 Excel 一致"),
    ("scenarios.yaml 期望值", "sheets 去 TK_PxQ；defined_names 719→709；tk_values 63→85；tk_pending 0→2；情境輸入 TKQ_→TK_（3 個）；ay 情境對應檢查 C132→C155", "—", "有意識更新"),
]


def style_header(ws, row=1):
    for c in ws[row]:
        c.font, c.fill, c.alignment = Font(bold=True), PatternFill("solid", fgColor="FFEDEDED"), Alignment(wrap_text=True, vertical="top")


def tk_vec(ws, name):
    for r in range(10, 300):
        if ws.cell(r, 1).value == name:
            return r, [ws.cell(r, c).value for c in range(10, 25)], ws.cell(r, 3).value, ws.cell(r, 4).value, ws.cell(r, 5).value
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--attr", type=Path, required=True)
    ap.add_argument("--old", type=Path, required=True)
    ap.add_argument("--new", type=Path, required=True)
    ap.add_argument("--out", type=Path, required=True)
    a = ap.parse_args()
    d = json.loads(a.attr.read_text(encoding="utf-8"))
    sc, groups = d["scen"], d["groups"]
    ow, nw = openpyxl.load_workbook(a.old, data_only=True)["TK_Link"], openpyxl.load_workbook(a.new, data_only=True)["TK_Link"]
    out = openpyxl.Workbook()
    # ① TK 名稱前後值
    w1 = out.active
    w1.title = "①TK名稱前後值"
    w1.append(["歸因組", "Tokenomics 名稱", "說明", "單位", "世代（基準欄）", "v5.26（v0.6.1）", "v19.0（v0.6.2）", "變動 %"])
    gof = {n: g for g, ns in groups.items() for n in ns}
    allnames = [nw.cell(r, 1).value for r in range(10, 300) if nw.cell(r, 1).value]
    for n in allnames:
        nv = tk_vec(nw, n)
        ov = tk_vec(ow, n)
        cols = [1, 4, 7, 10, 13] if nv[4] == 15 else list(range(nv[4]))
        changed = ov is not None and nv[1][:nv[4]] != ov[1][:ov[4]]
        if ov is not None and not changed:
            continue
        for j, k in enumerate(cols):
            gen = GENS[j] if nv[4] == 15 else "—"
            o = ov[1][k] if ov else None
            r = w1.max_row + 1
            w1.append([GLAB.get(gof.get(n), "新增（P×Q／RNG；v5.26 無此名稱）") if ov is None else GLAB[gof[n]], n, nv[2], nv[3], gen, o if o is not None else "（無）", nv[1][k],
                       f'=IF(ISNUMBER(F{r}),IF(F{r}=0,"",G{r}/F{r}-1),"")'])
    for r in range(2, w1.max_row + 1):
        w1.cell(r, 8).number_format = "0.0%"
    # ② 主要輸出
    w2 = out.create_sheet("②主要輸出與歸因")
    w2.append(["輸出", "列", *YEARS, "說明"])
    rows = {}
    for key, title, fmt in BLOCKS:
        o, nn = sc["OLD"][key], sc["NEW"][key]
        if key == "PXQ_Ratio":          # v0.6.1 的 PxQ 本來就用 v19.0；『全部還原』變體會動到 PxQ，故前後都取 v19.0 實值，歸因列為 0（不納入歸因）
            o = nn
            sc = {**sc, **{f"ONLY_{g}": {**sc[f"ONLY_{g}"], key: nn} for g in "1234"}}
        n = len(o)
        if key == "FND_ExtNeedCum":
            idx = list(range(1, 6))        # FY2026–FY2030（vector 首格為空）
        else:
            idx = list(range(n))
        def put(label, vals, f=fmt, note=""):
            r = w2.max_row + 1
            row = [title if label == "前（v0.6.1，v5.26）" else "", label] + [None] * 6
            for j, v in zip(range(n), vals):
                col = 2 + (j if key != "FND_ExtNeedCum" else j)
                row[col] = v
            w2.append(row + [note])
            for c in range(3, 9):
                w2.cell(r, c).number_format = f
            return r
        first = w2.max_row + 1
        rb = put("前（v0.6.1，v5.26）", o)
        ra = put("後（v0.6.2，v19.0）", nn)
        rd = w2.max_row + 1
        w2.append(["", "差額（後 − 前）"] + [f'=IF(AND(ISNUMBER({L(c)}{ra}),ISNUMBER({L(c)}{rb})),{L(c)}{ra}-{L(c)}{rb},"")' for c in range(3, 9)] + [""])
        for c in range(3, 9):
            w2.cell(rd, c).number_format = fmt
        gr = {}
        for g in "1234":
            lab = f"歸因 {GLAB[g]}（只換這組）"
            rr = put(lab, [x - y if isinstance(x, (int, float)) and isinstance(y, (int, float)) else None for x, y in zip(sc[f"ONLY_{g}"][key], o)])
            gr[g] = rr
        ri = w2.max_row + 1
        w2.append(["", "交互項（差額 − Σ單換）"] + [f'=IF(ISNUMBER({L(c)}{rd}),{L(c)}{rd}-SUM({L(c)}{gr["1"]}:{L(c)}{gr["4"]}),"")' for c in range(3, 9)] + ["全換結果與單換加總的差；0＝各組獨立"])
        for c in range(3, 9):
            w2.cell(ri, c).number_format = fmt
        if key in ("CMP_InfGW_Eff", "CMP_RDGW", "COST_SupplierHold", "COST_CloudGM"):
            rp = w2.max_row + 1
            w2.append(["", "差額 %（後 ÷ 前 − 1）"] + [f'=IF(AND(ISNUMBER({L(c)}{rb}),{L(c)}{rb}<>0),{L(c)}{ra}/{L(c)}{rb}-1,"")' for c in range(3, 9)] + [""])
            for c in range(3, 9):
                w2.cell(rp, c).number_format = "0.0%"
        rows[key] = (rb, ra)
        w2.append([" "])
    # 研發占總需求比（Excel 公式）
    t = w2.max_row + 1
    w2.append(["研發 GW 占總需求比", "前"] + [f"={L(c)}{rows['CMP_RDGW'][0]}/{L(c)}{rows['CMP_DemandGW'][0]}" for c in range(3, 9)] + ["研發 GW ÷ 總需求 GW"])
    w2.append(["", "後"] + [f"={L(c)}{rows['CMP_RDGW'][1]}/{L(c)}{rows['CMP_DemandGW'][1]}" for c in range(3, 9)] + [""])
    w2.append(["", "差額（百分點）"] + [f"={L(c)}{t + 1}-{L(c)}{t}" for c in range(3, 9)] + [""])
    for r in (t, t + 1, t + 2):
        for c in range(3, 9):
            w2.cell(r, c).number_format = "0.0%"
    w2.append([" "])
    w2.append(["外部資金需求峰值年", "前／後", "2029", "2029", "", "", "", "", "FND_ExtNeed 年度向量：2027／2028／2029／2030＝38.4／134.9／137.6／131.9 $B，前後相同；峰值年 2029"])
    # ③ 預設
    w3 = out.create_sheet("③已套用的預設")
    w3.append(["問題", "採用的預設", "替代選項", "對結果的影響"])
    for x in DEFAULTS:
        w3.append(list(x))
    # ④ 缺口與對照讀法
    w4 = out.create_sheet("④缺口與對照讀法")
    w4.append(["項目", "v5.26（v0.6.1）", "v19.0（v0.6.2）", "說明"])
    q1 = lambda ws, n: tk_vec(ws, n)[1][0]
    w4.append(["TK IF_AllocQ1（研發算力占比 R1）", q1(ow, "IF_AllocQ1"), q1(nw, "IF_AllocQ1"), "口徑改變（v11.0／v12.0 起：服務 GW 改用支出路線）"])
    w4.append(["RNG_IF_AllocQ1_Lo", "（無）", q1(nw, "RNG_IF_AllocQ1_Lo"), "v13.0 起視為區間；RNG_StaleFlag＝『一致』才可使用"])
    w4.append(["RNG_IF_AllocQ1_Hi", "（無）", q1(nw, "RNG_IF_AllocQ1_Hi"), ""])
    w4.append(["本模型 2025 研發 GW 占（研發＋推論）", "=1/(1+0.7)", "=1/(1+0.7)", "研發 GW 1.000、推論 GW 0.700（Checks C159；兩版相同）"])
    w4.append(["TK IF_AllocRDGW（研發 GW·年，物理下限）", q1(ow, "IF_AllocRDGW"), q1(nw, "IF_AllocRDGW"), "本模型 2025 研發 GW 1.000 為其 11 倍（v5.26）→ 4.9 倍（v19.0）"])
    w4.append(["TK L1_Ans3（研發算力成本 $B/年）", q1(ow, "L1_Ans3"), q1(nw, "L1_Ans3"), "區間 Lo／Hi 另見 TK_Link"])
    w4.append(["Q3 對照倍數 2025（研發算力 ÷ L1_Ans3）", "=" + str(sc["OLD"]["COST_RDCompute"][0]) + "/B7", "=" + str(sc["NEW"]["COST_RDCompute"][0]) + "/C7", "約 10 倍縮為約 6 倍"])
    w4.append(["Tokenomics 缺口 1", "", "", "Q3 對照需 IF_ 等義名稱（研發用算力成本 $B/年）；目前只能連 L1_Ans3"])
    w4.append(["Tokenomics 缺口 2", "", "", "IF_PxEffMix_Lo／_Hi（2026 有效牌價無區間）；check_tk_snapshot 報 MISSING 2"])
    for w in (w1, w2, w3, w4):
        style_header(w)
        w.freeze_panes = "A2"
        for i, wd in enumerate([26, 30, 40, 22, 14, 16, 16, 12, 60], start=1):
            w.column_dimensions[L(i)].width = wd
        for row in w.iter_rows(min_row=2):
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")
    w2.column_dimensions["A"].width = 34
    w2.column_dimensions["B"].width = 30
    for c in "CDEFGH":
        w2.column_dimensions[c].width = 12
    tmp = Path(tempfile.mkdtemp()) / a.out.name
    out.save(tmp)
    a.out.parent.mkdir(parents=True, exist_ok=True)
    shutil.move(str(lo_recalc(tmp, Path(tempfile.mkdtemp()))), str(a.out))  # 跨磁碟可用（/tmp 可能與工作目錄不同磁碟）
    print("saved", a.out)


if __name__ == "__main__":
    main()
