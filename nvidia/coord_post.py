# NVIDIA 三層基準協調試作（2026-11-14）：對帳中標示與兩讀數目標價（三層基準協調規則 v1.1 C4；個股 HTML 版面規格 v1.3 I 節）。
# 試作，請勿合併（步驟 5a 前置；base＝claude/nvidia-v0.3.7-layout13）。
#   write_inputs(wb, SIL)：In 頁末新增「Y. 三層協調」一節——現行基準 GW（Siliconomics SIF_DeployGW_Base）、另一側讀數 GW、對帳中旗標、
#                          兩者比值 co_r_<yy>（具名範圍）。正式建置與另一側讀數建置都呼叫（另一側讀數建置的顆數乘數讀 co_r_27..29）。
#   write_sheet(wb, here)：新增 Coord 頁——兩讀數 GW、兩讀數目標價與評等、橫幅揭露文字、損益表年度標示（HTML 用）、快照一致性核對。
#   另一側讀數的數值由 nvidia/coord_alt.py 整本重建（NV_OVERRIDE="_coord=1"：三情境 CY2027 起產業量 × co_r，其他輸入不變）後寫入 coord_alt.json。
# 集中輸入：load()。Siliconomics 快照有 SIF_DeployGW_Demand_<yyyy>／SIF_RecoFlag_<yyyy> 時改讀快照；否則讀 coord_inputs.json 的暫值（標「暫值，待 Siliconomics 試作分支」）。
import json, os
from openpyxl.styles import Font
from openpyxl.workbook.defined_name import DefinedName

YEARS = (2026, 2027, 2028, 2029)
APPLY = (2027, 2028, 2029)          # 另一側讀數套用的日曆年（FY2027≈CY2026 為錨點，沿用既有 Wattonomics 錨定案例的做法）
WTA_CO = {y: f"co_r_{y % 100}" for y in APPLY}
FY_OF = {2026: [2027], 2027: [2028], 2028: [2029], 2029: [2030, 2031]}   # 依賴各日曆年的會計年度（FY(t)≈CY(t−1)；FY2031 為 FY2030 的延伸期）
BOLD = Font(bold=True); BLUE = Font(color="0000FF")
HERE = os.path.dirname(os.path.abspath(__file__))


def load(SIL):
    """集中輸入（唯一一處）。來源三選一，依序：
       (1) sil：Siliconomics 現行快照（reader）已有 SIF_DeployGW_AltRatio_<yyyy>、SIF_RecoFlag_<yyyy>、SIF_UnitsCoef_Base（試作分支合併後自動改用）；
       (2) trial：coord_inputs.json 的 mode＝"trial"——Siliconomics 試作分支快照的摘錄（分支、SHA、檔案 sha256 記在 trial 區塊；extract() 產生）；
       (3) temp：coord_inputs.json 的暫值（派工訊息給定）。
    回傳 base（現行基準 GW）、alt（另一側讀數 GW）、ratio（另一側 ÷ 基準）、flag、ucoef（2028／2029 基準係數）、mode、src_*。"""
    N = SIL.NAMES
    J = json.load(open(os.path.join(HERE, "coord_inputs.json"), encoding="utf-8"))
    sb = dict(zip(YEARS, N["SIF_DeployGW_Base"]))
    if all(f"SIF_DeployGW_AltRatio_{y}" in N and f"SIF_RecoFlag_{y}" in N for y in YEARS) and "SIF_UnitsCoef_Base" in N:
        T = {k: v for k, v in N.items() if k.startswith(("SIF_DeployGW_", "SIF_Reco", "SIF_UnitsCoef", "SIF_crd_"))}; mode, tag = "sil", f"Siliconomics {SIL.VERSION}"
    elif J.get("mode") == "trial":
        T = J["trial"]["names"]; mode = "trial"
        tag = f"Siliconomics 試作分支 {J['trial']['branch']}@{J['trial']['sha']}（{J['trial']['version']}；摘錄於 nvidia/coord_inputs.json）"
    else:
        T = None; mode = "temp"
    if T is not None:
        uc = T["SIF_UnitsCoef_Base"]; uc = dict(zip(range(2029 - len(uc) + 1, 2030), uc))
        assert uc[2026] == 1 and uc[2027] == 1, uc   # 本模型 2026–27 顆數係數固定為 1（引擎）；不為 1 時須改引擎，交判斷
        C = dict(base={y: T[f"SIF_DeployGW_Current_{y}"] for y in YEARS}, alt={y: T[f"SIF_DeployGW_Alt_{y}"] for y in YEARS},
                 ratio={y: T[f"SIF_DeployGW_AltRatio_{y}"] for y in YEARS}, flag={y: int(T[f"SIF_RecoFlag_{y}"]) for y in YEARS},
                 rstatus={y: T.get(f"SIF_RecoStatus_{y}", "") for y in YEARS}, lead={y: T.get(f"SIF_RecoLead_{y}", "") for y in YEARS},
                 ucoef={2028: uc[2028], 2029: uc[2029]}, cov_done=T.get("SIF_crd_cov_done"), temp=False, mode=mode, status_txt=tag,
                 src_base=f"{tag} SIF_DeployGW_Current_<yyyy>", src_alt=f"{tag} SIF_DeployGW_Alt_<yyyy>", src_ratio=f"{tag} SIF_DeployGW_AltRatio_<yyyy>",
                 src_flag=f"{tag} SIF_RecoFlag_<yyyy>", src_ucoef=f"{tag} SIF_UnitsCoef_Base")
    else:
        C = dict(base=sb, alt={y: J["alt_gw"][str(y)] for y in YEARS}, ratio=None, flag={y: int(J["reco_flag"][str(y)]) for y in YEARS},
                 rstatus={y: "" for y in YEARS}, lead={y: "" for y in YEARS}, ucoef={2028: N["SIF_um28_Base"], 2029: N["SIF_um29_Base"]}, cov_done=None,
                 temp=True, mode=mode, status_txt=J["status"], src_base=f"Siliconomics {SIL.VERSION} SIF_DeployGW_Base",
                 src_alt=f"{J['status']}（nvidia/coord_inputs.json；之後改讀 SIF_DeployGW_Alt_<yyyy>）", src_ratio="本頁上兩列",
                 src_flag=f"{J['status']}（nvidia/coord_inputs.json；之後改讀 SIF_RecoFlag_<yyyy>）", src_ucoef=f"Siliconomics {SIL.VERSION} SIF_um28_Base／SIF_um29_Base")
    for y in YEARS:   # 兩讀數的「現行基準」須與本模型其餘部分讀的 Siliconomics 快照一致；不一致（例：cov_done＝1 改了基準）時停下交判斷
        assert abs(C["base"][y] - sb[y]) < 0.0005, ("現行基準 GW 與 reader 快照不一致", y, C["base"][y], sb[y])
    C["status"] = C["status_txt"]
    return C


def extract(snap, branch, sha):
    """把 Siliconomics 試作分支快照摘錄進 coord_inputs.json（mode＝trial）。"""
    d = json.load(open(snap, encoding="utf-8")); N = d["names"]
    keep = {k: v for k, v in N.items() if k.startswith(("SIF_DeployGW_Supply_", "SIF_DeployGW_Demand_", "SIF_DeployGW_Current_", "SIF_DeployGW_Alt_", "SIF_DeployGW_AltRatio_",
                                                         "SIF_RecoStatus_", "SIF_RecoLead_", "SIF_RecoLabel_", "SIF_RecoFlag_", "SIF_RecoUseConv_")) or k in ("SIF_UnitsCoef_Base", "SIF_crd_cov_done")}
    p = os.path.join(HERE, "coord_inputs.json"); J = json.load(open(p, encoding="utf-8"))
    J["mode"] = "trial"
    J["trial"] = {"branch": branch, "sha": sha, "file": os.path.basename(snap), "version": d["source"]["version"], "xlsx_sha256": d["source"]["sha256"], "names": dict(sorted(keep.items()))}
    json.dump(J, open(p, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return len(keep)


def write_inputs(wb, SIL):
    C = load(SIL); ws = wb["In"]; r = ws.max_row + 2
    ws.cell(r, 1, "Y. 三層協調：產業部署 GW 兩讀數、對帳中旗標與基準顆數係數（三層基準協調規則 v1.1 C4；試作；來源：" + C["status"] + "）").font = BOLD; r += 1
    def put(key, label, unit, val, tag, src, note="", inp=False):
        nonlocal r
        for j, v in enumerate((key, label, unit, None, val, None, tag, src, note), 1):
            if v is not None: ws.cell(r, j, v)
        if inp: ws.cell(r, 5).font = BLUE
        wb.defined_names[key] = DefinedName(key, attr_text=f"In!$E${r}"); r += 1
    for y in YEARS:
        put(f"co_gwb_{y % 100}", f"現行基準：產業部署 GW CY{y}（Siliconomics 基準）", "GW", C["base"][y], "Derived", C["src_base"],
            "本模型的三情境顆數以此為基準（情境 B＝Siliconomics 基準）；只作兩讀數比值的分母" if y == 2026 else "同上")
    for y in YEARS:
        put(f"co_gwa_{y % 100}", f"另一側讀數：產業部署 GW CY{y}（Wattonomics 換算值）", "GW", C["alt"][y], "Assumed" if C["temp"] else "Derived", C["src_alt"],
            "三層基準協調規則 v1.1 C3 平滑法換算（需求側）；只用於兩讀數目標價，不進基準" if y == 2026 else "同上", inp=C["temp"])
    for y in YEARS:
        put(f"co_flag_{y}", f"CY{y} 是否「對帳中」（1＝是）", "旗標", C["flag"][y], "Assumed" if C["temp"] else "Derived", C["src_flag"],
            "1＝兩層對帳未結案：依賴本年的數字標「對帳中」，估值章並列兩讀數目標價（規則 C4）" if y == 2026 else "同上", inp=C["temp"])
    for y in YEARS:
        put(f"co_r_{y % 100}", f"另一側 ÷ 現行基準 CY{y}", "x", C["ratio"][y] if C["ratio"] else f"=co_gwa_{y % 100}/co_gwb_{y % 100}", "Derived", C["src_ratio"],
            "另一側讀數建置：CY2027 起三情境產業量（運算顆數與網通）× 本比值；CY2026 不套用（FY2027 為錨點，見 Coord 頁說明）" if y == 2026 else
            ("另一側讀數建置的顆數乘數" if y == 2027 else "同上"))
    for y in (2028, 2029):
        put(f"co_ucoef_{y % 100}", f"基準顆數係數 CY{y}（情境 B＝Siliconomics 基準；um{y % 100}_B 讀本列）", "x", C["ucoef"][y], "Derived", C["src_ucoef"],
            "Siliconomics 協調試作起，情境 B 的顆數係數改讀 Siliconomics 基準係數（cov_done＝0 時與原 um28_B／um29_B 同值）；情境 A、C 仍為個股自訂" if y == 2028 else "同上")
    if C["ratio"]:
        put("co_chk_r", "核對：Σ|另一側 ÷ 現行 − 比值|（應為 0）", "x", "=" + "+".join(f"ABS(co_gwa_{y % 100}/co_gwb_{y % 100}-co_r_{y % 100})" for y in YEARS), "Derived", "本節", "")
    # 情境 B 的基準係數改讀上兩列（In 頁 um28_B、um29_B 由數值改為公式；數值不變）
    for row in ws.iter_rows(min_row=4):
        k = row[0].value
        if k in ("um28_B", "um29_B"):
            yy = k[2:4]
            assert abs(float(row[4].value) - C["ucoef"][2000 + int(yy)]) < 1e-12, (k, row[4].value, C["ucoef"])   # 本試作數值不變；不同時停下交判斷
            row[4].value = f"=co_ucoef_{yy}"; row[6].value = "Derived"
            row[8].value = f"Siliconomics 協調試作：改讀 {C['src_ucoef']}（In 頁 Y 節 co_ucoef_{yy}）；原值 {C['ucoef'][2000 + int(yy)]} 不變。" + str(row[8].value or "")
    return C


def locate(wb):
    """兩讀數對照用的儲存格位置（依標籤找，正式建置與另一側讀數建置結構相同）。"""
    def fr(sh, prefix, start=1):
        ws = wb[sh]
        return next(r for r in range(start, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value.strip().startswith(prefix))
    L = {"tp": ("Val", f"B{fr('Val', '目標價（DCF 與 Comparable 加權）')}"), "dcf": ("Val", f"B{fr('Val', 'DCF 目標價（情境加權）')}"),
         "comp": ("Val", f"B{fr('Val', 'Comparable 目標價')}"), "fwd": ("Val", f"B{fr('Val', '12 個月後的前瞻 EPS（情境加權）')}"),
         "up": ("Val", f"B{fr('Val', '上檔空間：加權目標價')}")}
    ra = fr("PL", "情境 A"); rb = fr("PL", "情境 B"); rc = fr("PL", "情境 C")
    for s, r0 in (("A", ra), ("B", rb), ("C", rc)):
        re_ = fr("PL", "EPS（", r0)
        for y, c in zip((2028, 2029, 2030), "EFG"): L[f"eps_{s}_{y}"] = ("PL", f"{c}{re_}")
    rv = fr("PL", "營收", rb)
    for y, c in zip((2028, 2029, 2030), "EFG"): L[f"rev_B_{y}"] = ("PL", f"{c}{rv}")
    rbb = fr("PL", "情境 B：")
    ru = fr("PL", "NVIDIA 認列顆數合計", rbb)
    for y, c in zip((2028, 2029, 2030), "EFG"): L[f"nvu_B_{y}"] = ("PL", f"{c}{ru}")
    return L


REC = lambda c: f'IF({c}>0.5,"強力買進",IF({c}>0.25,"買進",IF({c}>=-0.25,"中立",IF({c}>=-0.5,"放空","強力放空"))))'   # 同 Val 頁「投資建議」五級固定規則


def write_sheet(wb, C):
    A = json.load(open(os.path.join(HERE, "coord_alt.json"), encoding="utf-8"))
    L = locate(wb)
    ws = wb.create_sheet("Coord", index=wb.sheetnames.index("Val") + 1)
    ws.column_dimensions["A"].width = 58
    for c in "BCDEFG": ws.column_dimensions[c].width = 18
    ws["A1"] = ("Coord — 三層協調：對帳中標示與兩讀數目標價（三層基準協調規則 v1.1 C4；個股 HTML 版面規格 v1.3 I 節）。試作，請勿合併。"
                "個股不裁決：基準只讀 Siliconomics 現行基準；對帳未結案的年份並列「現行基準讀數」與「另一側讀數」兩個目標價。")
    ws["A1"].font = BOLD
    ws["A2"] = (("另一側讀數為暫值，待 Siliconomics 試作分支（SIF_DeployGW_Alt_<yyyy>、SIF_RecoFlag_<yyyy>）；集中輸入 nvidia/coord_inputs.json，In 頁 Y 節。" if C["temp"] else f"兩讀數、對帳中旗標與基準顆數係數取自 {C['status']}（集中輸入 nvidia/coord_inputs.json，In 頁 Y 節）。")
                + "另一側讀數＝只換產業部署 GW 一個輸入：三情境 CY2027 起運算顆數與網通營收 × 另一側 ÷ 現行基準（FY2027 為錨點不變；k_rec、價格、份額、折現率、情境權重與 Comparable 倍數皆不變），整本重建後讀 Val 與 PL（nvidia/coord_alt.py 快照）。")
    r = 4
    ws.cell(r, 1, "A. 產業部署 GW 兩讀數（日曆年）").font = BOLD; r += 1
    for j, h in enumerate(["年度", "現行基準 GW", "另一側讀數 GW", "另一側 ÷ 現行", "對帳中（1＝是）", "估值是否依賴（1＝是）", "依賴的會計年度", "Siliconomics 狀態／主導層"], 1): ws.cell(r, j, h).font = BOLD
    r += 1; RA = {}
    for y in YEARS:
        yy = y % 100; RA[y] = r
        ws.cell(r, 1, f"CY{y}"); ws.cell(r, 2, f"=co_gwb_{yy}").number_format = "0.000"; ws.cell(r, 3, f"=co_gwa_{yy}").number_format = "0.000"
        ws.cell(r, 4, f"=co_r_{yy}").number_format = "0.000"; ws.cell(r, 5, f"=co_flag_{y}")
        ws.cell(r, 6, 1)   # 估值依賴：DCF 用 FY2027–FY2032 全部年份、Comparable 用 FY2028–FY2029 EPS；FY2027 經 k_rec 校準依賴 CY2026
        ws.cell(r, 7, "、".join(f"FY{f}" for f in FY_OF[y])); ws.cell(r, 8, f"{C['rstatus'][y]}／{C['lead'][y]}" if C["rstatus"][y] else "—")
        wb.defined_names[f"co_dep_{y}"] = DefinedName(f"co_dep_{y}", attr_text=f"Coord!$F${r}"); r += 1
    ws.cell(r, 1, "估值依賴：DCF 用 FY2027–FY2032 全部年份，Comparable 用 FY2028–FY2029 EPS，FY2027 的認列顆數係數（k_rec）以 CY2026 顆數校準；四年都標 1。"
                  "CY2026 的另一側讀數不套用（FY2027 營收為公司實績＋指引的錨點），只標示。"); r += 2
    yrs = "MID(" + "&".join(f'IF(co_flag_{y}=1,"、{y}","")' for y in YEARS) + ",2,99)"
    anyf = "OR(" + ",".join(f"AND(co_flag_{y}=1,co_dep_{y}=1)" for y in YEARS) + ")"
    ws.cell(r, 1, "B. 兩讀數：顆數 → 營收 → EPS → 目標價（美元；營收 US$B；顆數百萬顆）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "現行基準讀數", "另一側讀數", "另一側相對現行"], 1): ws.cell(r, j, h).font = BOLD
    r += 1; RB = {}
    rows = [("nvu_B_2028", "NVIDIA 認列顆數（情境 B）FY2028", "0.00"), ("nvu_B_2029", "NVIDIA 認列顆數（情境 B）FY2029", "0.00"), ("nvu_B_2030", "NVIDIA 認列顆數（情境 B）FY2030", "0.00"),
            ("rev_B_2028", "營收（情境 B）FY2028", "#,##0.0"), ("rev_B_2029", "營收（情境 B）FY2029", "#,##0.0"), ("rev_B_2030", "營收（情境 B）FY2030", "#,##0.0")]
    for s, nm in (("A", "樂觀"), ("B", "基準"), ("C", "保守")):
        rows += [(f"eps_{s}_{y}", f"EPS 情境 {s}（{nm}）FY{y}", "0.00") for y in (2028, 2029, 2030)]
    rows += [("fwd", "12 個月後前瞻 EPS（情境加權）", "0.00"), ("dcf", "DCF 目標價（情境加權）", "#,##0.00"), ("comp", "Comparable 目標價（進加權，相對倍數調整後）", "#,##0.00"),
             ("tp", "加權目標價", "#,##0.00"), ("up", "上檔空間（加權目標價 ÷ 現價 − 1）", "0.0%")]
    for k, lab, fmt in rows:
        RB[k] = r; ws.cell(r, 1, lab)
        sh, co = L[k]
        if k == "comp":
            ws.cell(r, 2, "=tp_comp_adj")
        else:
            ws.cell(r, 2, f"={sh}!{co}")
        ws.cell(r, 3, A["alt"][k])
        ws.cell(r, 3).font = Font(color="7F7F7F")
        for c in (2, 3): ws.cell(r, c).number_format = fmt
        if k == "up":
            ws.cell(r, 4, f"=C{r}-B{r}").number_format = '+0.0%;-0.0%'
        else:
            ws.cell(r, 4, f"=C{r}/B{r}-1").number_format = '+0.0%;-0.0%'
        r += 1
    ws.cell(r, 1, "投資建議（五級固定規則，同 Val 頁）"); ws.cell(r, 2, "=" + REC(f"B{RB['up']}")); ws.cell(r, 3, "=" + REC(f"C{RB['up']}")); RB["rec"] = r; r += 1
    ws.cell(r, 1, "兩讀數評等是否相同"); ws.cell(r, 2, f'=IF(B{RB["rec"]}=C{RB["rec"]},"相同","不同")'); RB["same"] = r; r += 1
    ws.cell(r, 1, "現行基準讀數：Comparable 欄為 Val 頁 tp_comp_adj（同業倍數 × 情境加權前瞻 EPS × 相對倍數調整係數）；另一側讀數為整本重建快照（灰字）。"); r += 2
    ws.cell(r, 1, "C. 展示層文字（HTML 用；版面規格 v1.3 I 節）").font = BOLD; r += 1
    RC = {}
    RC["flag"] = r; ws.cell(r, 1, "是否有估值依賴的對帳中年份"); ws.cell(r, 2, f'=IF({anyf},"是","否")'); r += 1
    RC["years"] = r; ws.cell(r, 1, "對帳中年份（日曆年）"); ws.cell(r, 2, "=" + yrs); r += 1
    RC["hero"] = r; ws.cell(r, 1, "橫幅揭露（≤ 90 字）")
    ws.cell(r, 2, f'=IF({anyf},"對帳中（"&B{RC["years"]}&" 產業 GW）：另一側讀數目標價 "&TEXT(C{RB["tp"]},"0.00")&" 美元（"&C{RB["rec"]}&"）"&IF(B{RB["same"]}="相同","，評等與現行基準讀數相同。","，與現行基準讀數（"&B{RB["rec"]}&"）評等不同。"),"")'); r += 1
    RC["lab_cur"] = r; ws.cell(r, 1, "目標價區標籤：現行基準讀數"); ws.cell(r, 2, "現行基準讀數（Siliconomics 基準；橫幅主值）"); r += 1
    RC["lab_alt"] = r; ws.cell(r, 1, "目標價區標籤：另一側讀數"); ws.cell(r, 2, '="另一側讀數（Wattonomics 換算的產業 GW"&IF(' + ("TRUE" if C["temp"] else "FALSE") + ',"；暫值","；Siliconomics 試作")&"）"'); r += 1
    RC["note"] = r; ws.cell(r, 1, "目標價區說明")
    ws.cell(r, 2, f'="兩讀數只差產業部署 GW 一個輸入：CY2027–CY2029 另一側讀數為現行基準的 "&TEXT(co_r_27,"0.00")&"／"&TEXT(co_r_28,"0.00")&"／"&TEXT(co_r_29,"0.00")&" 倍（運算顆數與網通營收同乘），其他假設不變。個股不裁決哪一側正確，對帳在 Siliconomics 對帳單處理；結案前兩個目標價並列。"' + ('&"另一側讀數目前為暫值，待 Siliconomics 試作分支提供。"' if C["temp"] else '&"兩讀數取自 Siliconomics 協調試作分支（未合併）。"')); r += 1
    RC["pl"] = r; ws.cell(r, 1, "損益表年度標示（FY2027E–FY2031E）")
    for j, fy in enumerate(range(2027, 2032)):
        cy = next(y for y, fs in FY_OF.items() if fy in fs)
        ws.cell(r, 2 + j, f'=IF(AND(co_flag_{cy}=1,co_dep_{cy}=1),"FY{fy}E（對帳中）","FY{fy}E")')
    r += 2
    ws.cell(r, 1, "D. 快照一致性（另一側讀數為建置快照；In 頁 Y 節或模型改動而未重跑 python3 nvidia/coord_alt.py 時顯示 STALE）").font = BOLD; r += 1
    chk = []
    for y in APPLY:
        ws.cell(r, 1, f"快照所用比值 − 現行比值 CY{y}（應為 0）"); ws.cell(r, 2, f"={A['ratio'][str(y)]}-co_r_{y % 100}").number_format = "0.000000"; chk.append(f"ABS(B{r})<1E-9"); r += 1
    ws.cell(r, 1, "快照的現行基準加權目標價 − 本檔加權目標價（應為 0）"); ws.cell(r, 2, f"={A['base']['tp']}-B{RB['tp']}").number_format = "0.000000"; chk.append(f"ABS(B{r})<0.005"); r += 1
    RC["state"] = r; ws.cell(r, 1, "另一側讀數快照狀態").font = BOLD; ws.cell(r, 2, f'=IF(AND({",".join(chk)}),"CURRENT","STALE")'); r += 1
    return {"RA": RA, "RB": RB, "RC": RC}


VER = "0.3.8"


def stamp(wb):
    """版本字串：v0.3.7 → v0.3.8（試作）；build_id 改含本試作檔案的雜湊（README!A1 與 Sens 同步）。"""
    import hashlib
    import v037_post as V7
    RD, SE = wb["README"], wb["Sens"]
    old = SE.cell(V7.lab(SE, "build_id"), 4).value
    h = hashlib.sha256(old.encode())
    for fn in ("coord_post.py", "coord_inputs.json", "coord_alt.json", "gen_html_json.py"):
        h.update(fn.encode()); h.update(open(os.path.join(HERE, fn), "rb").read())
    new = f"nvidia-v{VER}-{h.hexdigest()[:10]}"
    n = 0
    for ws in (RD, SE):
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and old in c.value:
                    c.value = c.value.replace(old, new); n += 1
    a1 = RD["A1"].value
    assert "個股模型 v0.3.7" in a1, a1[:80]
    RD["A1"].value = a1.replace("個股模型 v0.3.7", f"個股模型 v{VER}（試作，請勿合併）", 1).replace(
        "標準損益表、反向 DCF 頁）", f"標準損益表、反向 DCF 頁）；v{VER} 試作：三層協調——對帳中標示與兩讀數目標價（規則 v1.1 C4；Coord 頁、In 頁 Y 節；基準數值不變）", 1)
    return f"build_id {old} → {new}（{n} 格）"


if __name__ == "__main__":   # python3 nvidia/coord_post.py extract <快照 json> <分支> <SHA>
    import sys
    if sys.argv[1] == "extract":
        print("摘錄", extract(sys.argv[2], sys.argv[3], sys.argv[4]), "個名稱")
