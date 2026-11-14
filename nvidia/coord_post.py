# NVIDIA 三層基準協調試作（2026-11-14）：對帳中標示與兩讀數目標價（三層基準協調規則 v1.1 C4；個股 HTML 版面規格 v1.3 I 節）。
# 2026-11-14 13:06 Andy 確認試作後轉正式（Siliconomics v0.19 PR #487 合併後讀正式快照）。
#   write_inputs(wb, SIL)：In 頁末新增「Y. 三層協調」一節——現行基準 GW（Siliconomics SIF_DeployGW_Base）、另一側讀數 GW、對帳中旗標、
#                          兩者比值 co_r_<yy>（具名範圍）。正式建置與另一側讀數建置都呼叫（另一側讀數建置的顆數乘數讀 co_r_27..29）。
#   write_sheet(wb, here)：新增 Coord 頁——兩讀數 GW、兩讀數目標價與評等、橫幅揭露文字、損益表年度標示（HTML 用）、快照一致性核對。
#   另一側讀數的數值由 nvidia/coord_alt.py 整本重建（NV_OVERRIDE="_coord=1"：情境 B、C 於 CY2027 起產業量 × co_r（情境 A＝供給側讀數不縮放），其他輸入不變）後寫入 coord_alt.json。
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
        bl = T.get("SIF_DeployGW_Base")   # 現行基準＝本庫基準（cov_done＝1 起由規則決定；Current_<yyyy> 為改前的現行值，只作對照）
        base = dict(zip(YEARS, bl)) if bl else {y: T[f"SIF_DeployGW_Current_{y}"] for y in YEARS}
        C = dict(base=base, prev={y: T.get(f"SIF_DeployGW_Current_{y}") for y in YEARS}, supply={y: T[f"SIF_DeployGW_Supply_{y}"] for y in YEARS}, alt={y: T[f"SIF_DeployGW_Alt_{y}"] for y in YEARS},
                 ratio={y: T[f"SIF_DeployGW_AltRatio_{y}"] for y in YEARS}, flag={y: int(T[f"SIF_RecoFlag_{y}"]) for y in YEARS},
                 rstatus={y: T.get(f"SIF_RecoStatus_{y}", "") for y in YEARS}, lead={y: T.get(f"SIF_RecoLead_{y}", "") for y in YEARS},
                 ucoef={2028: uc[2028], 2029: uc[2029]}, cov_done=T.get("SIF_crd_cov_done"), temp=False, mode=mode, status_txt=tag,
                 src_base=f"{tag} SIF_DeployGW_Base" if bl else f"{tag} SIF_DeployGW_Current_<yyyy>", src_supply=f"{tag} SIF_DeployGW_Supply_<yyyy>", src_alt=f"{tag} SIF_DeployGW_Alt_<yyyy>", src_ratio=f"{tag} SIF_DeployGW_AltRatio_<yyyy>",
                 src_flag=f"{tag} SIF_RecoFlag_<yyyy>", src_ucoef=f"{tag} SIF_UnitsCoef_Base")
    else:
        C = dict(base=sb, prev=dict(sb), supply=None, src_supply="", alt={y: J["alt_gw"][str(y)] for y in YEARS}, ratio=None, flag={y: int(J["reco_flag"][str(y)]) for y in YEARS},
                 rstatus={y: "" for y in YEARS}, lead={y: "" for y in YEARS}, ucoef={2028: N["SIF_um28_Base"], 2029: N["SIF_um29_Base"]}, cov_done=None,
                 temp=True, mode=mode, status_txt=J["status"], src_base=f"Siliconomics {SIL.VERSION} SIF_DeployGW_Base",
                 src_alt=f"{J['status']}（nvidia/coord_inputs.json；之後改讀 SIF_DeployGW_Alt_<yyyy>）", src_ratio="本頁上兩列",
                 src_flag=f"{J['status']}（nvidia/coord_inputs.json；之後改讀 SIF_RecoFlag_<yyyy>）", src_ucoef=f"Siliconomics {SIL.VERSION} SIF_um28_Base／SIF_um29_Base")
    if C["supply"]:   # 一致性：現行基準 GW＝供給側讀數（原始顆數 GW）× 基準顆數係數（2026–27 係數 1）
        for y in YEARS:
            k = C["ucoef"].get(y, 1)
            assert abs(C["base"][y] - C["supply"][y] * k) < 0.0005 * max(1, C["base"][y]), ("基準 GW ≠ 供給 × 係數", y, C["base"][y], C["supply"][y], k)
    C["status"] = C["status_txt"]
    return C


def extract(snap, branch, sha):
    """把 Siliconomics 試作分支快照摘錄進 coord_inputs.json（mode＝trial）。"""
    d = json.load(open(snap, encoding="utf-8")); N = d["names"]
    keep = {k: v for k, v in N.items() if k == "SIF_DeployGW_Base" or k.startswith(("SIF_DeployGW_Supply_", "SIF_DeployGW_Demand_", "SIF_DeployGW_Current_", "SIF_DeployGW_Alt_", "SIF_DeployGW_AltRatio_",
                                                         "SIF_RecoStatus_", "SIF_RecoLead_", "SIF_RecoLabel_", "SIF_RecoFlag_", "SIF_RecoUseConv_")) or k in ("SIF_UnitsCoef_Base", "SIF_crd_cov_done")}
    p = os.path.join(HERE, "coord_inputs.json"); J = json.load(open(p, encoding="utf-8"))
    J["mode"] = "trial"
    J["trial"] = {"branch": branch, "sha": sha, "file": os.path.basename(snap), "version": d["source"]["version"], "xlsx_sha256": d["source"]["sha256"], "names": dict(sorted(keep.items()))}
    json.dump(J, open(p, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return len(keep)


def write_inputs(wb, SIL):
    C = load(SIL); ws = wb["In"]; r = ws.max_row + 2
    ws.cell(r, 1, "Z. 三層協調：產業部署 GW 兩讀數、對帳中旗標與基準顆數係數（三層基準協調規則 v1.4 C4；來源：" + C["status"] + "）").font = BOLD; r += 1
    def put(key, label, unit, val, tag, src, note="", inp=False):
        nonlocal r
        for j, v in enumerate((key, label, unit, None, val, None, tag, src, note), 1):
            if v is not None: ws.cell(r, j, v)
        if inp: ws.cell(r, 5).font = BLUE
        wb.defined_names[key] = DefinedName(key, attr_text=f"In!$E${r}"); r += 1
    for y in YEARS:
        put(f"co_gwb_{y % 100}", f"現行基準：產業部署 GW CY{y}（Siliconomics 基準）", "GW", C["base"][y], "Derived", C["src_base"],
            "本模型的三情境顆數以此為基準（情境 B＝Siliconomics 基準）；只作兩讀數比值的分母" if y == 2026 else "同上")
    if C["supply"]:
        for y in YEARS:
            put(f"co_gws_{y % 100}", f"供給側讀數：產業部署 GW CY{y}（原始顆數換成的 GW）", "GW", C["supply"][y], "Derived", C["src_supply"],
                "情境 A（樂觀）＝供給側讀數（chat 端 2026-11-14 決定）：In 頁 um28_A、um29_A＝本列 ÷（現行基準 ÷ 基準係數）" if y == 2026 else "同上")
    for y in YEARS:
        put(f"co_gwa_{y % 100}", f"另一側讀數：產業部署 GW CY{y}（Wattonomics 換算值）", "GW", C["alt"][y], "Assumed" if C["temp"] else "Derived", C["src_alt"],
            "三層基準協調規則 v1.1 C3 平滑法換算（需求側）；只用於兩讀數目標價，不進基準" if y == 2026 else "同上", inp=C["temp"])
    for y in YEARS:
        put(f"co_flag_{y}", f"CY{y} 是否「對帳中」（1＝是）", "旗標", C["flag"][y], "Assumed" if C["temp"] else "Derived", C["src_flag"],
            "1＝兩層對帳未結案：依賴本年的數字標「對帳中」，估值章並列兩讀數目標價（規則 C4）" if y == 2026 else "同上", inp=C["temp"])
    for y in YEARS:
        put(f"co_r_{y % 100}", f"另一側 ÷ 現行基準 CY{y}", "x", C["ratio"][y] if C["ratio"] else f"=co_gwa_{y % 100}/co_gwb_{y % 100}", "Derived", C["src_ratio"],
            "另一側讀數建置：CY2027 起情境 B、C 產業量（運算顆數與網通）× 本比值（情境 A＝供給側讀數，不縮放）；CY2026 不套用（FY2027 為錨點，見 Coord 頁說明）" if y == 2026 else
            ("另一側讀數建置的顆數乘數" if y == 2027 else "同上"))
    for y in (2028, 2029):
        put(f"co_ucoef_{y % 100}", f"基準顆數係數 CY{y}（情境 B＝Siliconomics 基準；um{y % 100}_B 讀本列）", "x", C["ucoef"][y], "Derived", C["src_ucoef"],
            "Siliconomics v0.19 起，情境 B 的顆數係數改讀 Siliconomics 基準係數（cov_done＝1：2029 由規則以連續法決定）；情境 A＝供給側讀數、情境 C 仍為個股自訂" if y == 2028 else "同上")
    if C["supply"]:   # 敏感度「總量缺口全由 NVIDIA 吸收」用：NVIDIA 占原始顆數的比例（規則 v1.2 C3：總量依原始顆數等比例分配）
        NVK, OTK = ("bw", "rb", "ru", "hp"), ("amd", "tpu", "trn", "oth")
        u = lambda k, y: f"u_{k}_{y % 100}" if y <= 2028 else f"u_{k}_28*(1+g29_{k})"
        for y in APPLY:
            nv = "+".join(u(k, y) for k in NVK); al = nv + "+" + "+".join(u(k, y) for k in OTK)
            put(f"co_ns_{y % 100}", f"NVIDIA 占原始顆數比例 CY{y}（全部加速器平台）", "%", f"=({nv})/({al})", "Derived", "本頁 u_<平台>_<yy>、g29_<平台>",
                "敏感度「總量缺口全由 NVIDIA 吸收」：其他平台維持原始顆數，NVIDIA 顆數＝原始 − 總量缺口（下限 0）" if y == 2027 else "同上")
    if C["ratio"]:
        put("co_chk_r", "核對：Σ|另一側 ÷ 現行 − 比值|（應為 0）", "x", "=" + "+".join(f"ABS(co_gwa_{y % 100}/co_gwb_{y % 100}-co_r_{y % 100})" for y in YEARS), "Derived", "本節", "")
    # 情境 B 的基準係數改讀 Z 節（In 頁 um28_B、um29_B 由數值改為公式）；情境 A＝供給側讀數（um28_A、um29_A 改公式）。原值逐字保留在說明欄
    C["um_old"] = {}
    for row in ws.iter_rows(min_row=4):
        k = row[0].value
        if k in ("um28_B", "um29_B", "um28_A", "um29_A"):
            yy = k[2:4]; old = row[4].value; C["um_old"][k] = old
            if k.endswith("_B"):
                row[4].value = f"=co_ucoef_{yy}"; src = f"{C['src_ucoef']}（In 頁 Z 節 co_ucoef_{yy}）"
            else:
                if not C["supply"]: continue
                row[4].value = f"=co_gws_{yy}/(co_gwb_{yy}/co_ucoef_{yy})"; src = f"{C['src_supply']} ÷ 原始顆數 GW（In 頁 Z 節）"
            row[6].value = "Derived"
            row[8].value = f"Siliconomics 三層協調（NVIDIA v{VER}）：改讀 {src}；原值 {old}（Assumed，原說明：{row[8].value or ''}）"
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


def absorb_factor(cy, s, alt):
    """引擎乘數（nvidia/build_nvidia.py engine 的 nvs）：NVIDIA 認列顆數 ÷（原始 × 情境係數 × 另一側比值）＝ MAX(0, 1 − (1 − e) ÷ NVIDIA 比例) ÷ e，
    e＝情境係數 × 另一側比值（基準讀數時比值＝1）。其他平台維持原始顆數，總量缺口全由 NVIDIA 吸收。"""
    if cy < 2027 or s != "B":   # 只作用於情境 B（＝Siliconomics 基準）；情境 A＝供給側讀數無缺口，情境 C 為個股自訂係數（全吸收時 NVIDIA 顆數 < 0，不可行）
        return "1"
    um = "1" if cy == 2027 else f"um{cy % 100}_{s}"
    e = f"({um}*co_r_{cy % 100})" if alt else f"({um})"
    return f"(MAX(0,1-(1-{e})/co_ns_{cy % 100})/{e})"


def delete_wta(wb):
    """刪除舊「Wattonomics 錨定」敏感度案例（chat 端 2026-11-14 決定：已被 Coord 頁兩讀數取代）。
    為保持既有列號（HTML 設定與其他頁引用依列號），以清空儲存格方式刪除，首列留一行說明；In 頁 Z 節與具名範圍 wta_27..29 一併刪除；
    README 重點 3 的錨定句刪除。回傳清空格數與殘留引用（應為空）。"""
    import re
    note = "（v0.3.8 刪除：舊「Wattonomics 錨定」敏感度案例，已由 Coord 頁兩讀數目標價取代；本區保留空白以維持列號）"
    SE, DC, IN, RD = wb["Sens"], wb["DCF"], wb["In"], wb["README"]
    def rows_of(ws, pred):
        return [r for r in range(1, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, str) and pred(ws.cell(r, 1).value)]
    cleared = {}
    def clear(ws, r0, r1, msg=True):
        n = 0
        for r in range(r0, r1 + 1):
            for c in ws[r]:
                if c.value is not None:
                    c.value = None; n += 1
        if msg: ws.cell(r0, 1, note)
        cleared.setdefault(ws.title, []).append((r0, r1))
        return n
    n = 0
    s0 = rows_of(SE, lambda v: v.startswith("Wattonomics 錨定（三層校準原則）"))[0]
    s1 = rows_of(SE, lambda v: v.startswith("核對：情境 B 欄 DCF − Val 第 0.1 節情境 B 價值"))[0]
    n += clear(SE, s0, s1)
    m = rows_of(SE, lambda v: v.startswith("Wattonomics 錨定（情境 B × wta）"))
    for r in m: n += clear(SE, r, r)
    d0 = rows_of(DC, lambda v: v.startswith("Sens 案例路徑：Wattonomics 錨定"))[0]
    d1 = next(r for r in rows_of(DC, lambda v: v.startswith("檢查：終值占 EV 超過 80%")) if r > d0)   # fcff_dcf.write_detail 區塊的最後一列
    n += clear(DC, d0, d1)
    z0 = rows_of(IN, lambda v: v.startswith("Z. Wattonomics 錨定係數"))[0]
    n += clear(IN, z0, z0 + 3)
    for k in ("wta_27", "wta_28", "wta_29"):
        if k in wb.defined_names: del wb.defined_names[k]
    for row in RD.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("=") and "若改用 Wattonomics 錨定" in c.value:
                i = c.value.index("若改用 Wattonomics 錨定"); c.value = c.value[:i] + '"'
    # 殘留引用掃描：任何公式引用被清空的列或 wta_ 名稱
    pat = re.compile(r"(?:(Sens|DCF|In)!)\$?([A-Z]{1,3})\$?(\d+)")
    left = []
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not (isinstance(v, str) and v.startswith("=")): continue
                if re.search(r"\bwta_\d\d\b", v): left.append(f"{ws.title}!{c.coordinate}")
                for sh, col, rr in pat.findall(v):
                    if any(a <= int(rr) <= b for a, b in cleared.get(sh, [])): left.append(f"{ws.title}!{c.coordinate}→{sh}!{col}{rr}")
                if ws.title in cleared:   # 同頁未加頁名的引用
                    for col, rr in re.findall(r"(?<![!$A-Za-z_])\$?([A-Z]{1,3})\$?(\d+)(?![\d(])", v):
                        if any(a <= int(rr) <= b for a, b in cleared[ws.title]) and not any(a <= c.row <= b for a, b in cleared[ws.title]):
                            left.append(f"{ws.title}!{c.coordinate}→{col}{rr}")
    return n, cleared, left


def relabel(wb):
    """只改文字（數值不變）：情境 B 標籤與情境 C 說明改為讀 Siliconomics 基準係數（SIF_UnitsCoef_Base；2029 連續法）。標籤以公式帶入係數。"""
    lab = '"情境 B：{p}2028 顆數 ×"&TEXT(um28_B,"0.0#")&"、{p}2029 ×"&TEXT(um29_B,"0.000")&"（Siliconomics 基準係數；2029 連續法）"'
    T = [("Units", "A38", "情境 B：2028 顆數 ×0.8、2029 ×0.85", "=" + lab.format(p="")),
         ("PL", "A127", "情境 B：CY2028 顆數 ×0.8、CY2029 ×0.85", "=" + lab.format(p="CY")),
         ("Sens", "B20", "CY2028 ×0.8、CY2029 ×0.85；記憶體 B 路徑", '="CY2028 ×"&TEXT(um28_B,"0.0#")&"、CY2029 ×"&TEXT(um29_B,"0.000")&"（Siliconomics 基準係數）；記憶體 B 路徑"')]
    n = 0
    for sh, c, old, new in T:
        assert wb[sh][c].value == old, (sh, c, wb[sh][c].value)
        wb[sh][c].value = new; n += 1
    RD = wb["README"]; v = RD["A19"].value; o = "B＝基準（CY2028 顆數 ×0.8、CY2029 ×0.85）"
    assert isinstance(v, str) and o in v, v
    i = v.index(o)
    RD["A19"].value = '="' + v[:i].replace('"', '""') + 'B＝基準（CY2028 顆數 ×"&TEXT(um28_B,"0.0#")&"、CY2029 ×"&TEXT(um29_B,"0.000")&"，Siliconomics 基準係數；2029 連續法）' + v[i + len(o):].replace('"', '""') + '"'; n += 1
    IN = wb["In"]
    r = next(r for r in range(4, IN.max_row + 1) if IN.cell(r, 1).value == "um29_C")
    old = IN.cell(r, 9).value
    assert ("×0.706（基準係數 0.85）" in str(old)) or ("基準係數 0.71" in str(old)), old   # v0.18 口徑或 v0.19 reader 口徑
    IN.cell(r, 9).value = ('="個股自訂情境 C：＝Siliconomics 基準 ×"&TEXT(um29_C/co_ucoef_29,"0.000")&"（基準係數 "&TEXT(co_ucoef_29,"0.000")&"，SIF_UnitsCoef_Base，2029 連續法；v0.3.8）；2029 產業 GW 相對另一側讀數（Wattonomics 換算值）"'
                           '&TEXT(co_gws_29*um29_C/co_gwa_29-1,"+0%;-0%")&"（現行基準 "&TEXT(co_gwb_29/co_gwa_29-1,"+0%;-0%")&"）。原說明（Siliconomics reader 產生）：' + str(old).replace('"', '""') + '"')
    n += 1
    return n


REC = lambda c: f'IF({c}>0.5,"強力買進",IF({c}>0.25,"買進",IF({c}>=-0.25,"中立",IF({c}>=-0.5,"放空","強力放空"))))'   # 同 Val 頁「投資建議」五級固定規則


def write_sheet(wb, C):
    A = json.load(open(os.path.join(HERE, "coord_alt.json"), encoding="utf-8"))
    L = locate(wb)
    ws = wb.create_sheet("Coord", index=wb.sheetnames.index("Val") + 1)
    ws.column_dimensions["A"].width = 58
    for c in "BCDEFG": ws.column_dimensions[c].width = 18
    ws["A1"] = ("Coord — 三層協調：對帳中標示與兩讀數目標價（三層基準協調規則 v1.4 C4；個股 HTML 版面規格 v1.3 I 節）。"
                "個股不裁決：基準只讀 Siliconomics 現行基準；對帳未結案的年份並列「現行基準讀數」與「另一側讀數」兩個目標價。")
    ws["A1"].font = BOLD
    ws["A2"] = (("另一側讀數為暫值，待 Siliconomics 試作分支（SIF_DeployGW_Alt_<yyyy>、SIF_RecoFlag_<yyyy>）；集中輸入 nvidia/coord_inputs.json，In 頁 Z 節。" if C["temp"] else f"兩讀數、對帳中旗標與基準顆數係數取自 {C['status']}（集中輸入 nvidia/coord_inputs.json，In 頁 Z 節）。")
                + "另一側讀數＝只換產業部署 GW 一個輸入：情境 B、C 於 CY2027 起運算顆數與網通營收 × 另一側 ÷ 現行基準（情境 A＝供給側讀數，兩讀數相同）（FY2027 為錨點不變；k_rec、價格、份額、折現率、情境權重與 Comparable 倍數皆不變），整本重建後讀 Val 與 PL（nvidia/coord_alt.py 快照）。")
    r = 4
    ws.cell(r, 1, "A. 產業部署 GW 兩讀數（日曆年）").font = BOLD; r += 1
    for j, h in enumerate(["年度", "現行基準 GW", "另一側讀數 GW", "另一側 ÷ 現行", "對帳中（1＝是）", "估值是否依賴（1＝是）", "依賴的會計年度", "Siliconomics 狀態／主導層", "供給側讀數 GW（情境 A）", "改前現行基準 GW"], 1): ws.cell(r, j, h).font = BOLD
    r += 1; RA = {}
    for y in YEARS:
        yy = y % 100; RA[y] = r
        ws.cell(r, 1, f"CY{y}"); ws.cell(r, 2, f"=co_gwb_{yy}").number_format = "0.000"; ws.cell(r, 3, f"=co_gwa_{yy}").number_format = "0.000"
        ws.cell(r, 4, f"=co_r_{yy}").number_format = "0.000"; ws.cell(r, 5, f"=co_flag_{y}")
        ws.cell(r, 6, 1)   # 估值依賴：DCF 用 FY2027–FY2032 全部年份、Comparable 用 FY2028–FY2029 EPS；FY2027 經 k_rec 校準依賴 CY2026
        ws.cell(r, 7, "、".join(f"FY{f}" for f in FY_OF[y])); ws.cell(r, 8, f"{C['rstatus'][y]}／{C['lead'][y]}" if C["rstatus"][y] else "—")
        if C["supply"]: ws.cell(r, 9, f"=co_gws_{yy}").number_format = "0.000"
        if C.get("prev") and C["prev"][y] is not None: ws.cell(r, 10, C["prev"][y]).number_format = "0.000"
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
    RC["lab_alt"] = r; ws.cell(r, 1, "目標價區標籤：另一側讀數"); ws.cell(r, 2, '="另一側讀數（Wattonomics 換算的產業 GW"&IF(' + ("TRUE" if C["temp"] else "FALSE") + ',"；暫值","")&"）"'); r += 1
    RC["note"] = r; ws.cell(r, 1, "目標價區說明")
    ws.cell(r, 2, f'="兩讀數只差產業部署 GW 一個輸入：CY2027–CY2029 另一側讀數為現行基準的 "&TEXT(co_r_27,"0.00")&"／"&TEXT(co_r_28,"0.00")&"／"&TEXT(co_r_29,"0.00")&" 倍（情境 B、C 的運算顆數與網通營收同乘；情境 A＝供給側讀數不變），其他假設不變。個股不裁決哪一側正確，對帳在 Siliconomics 對帳單處理；結案前兩個目標價並列。"' + ('&"另一側讀數目前為暫值，待 Siliconomics 試作分支提供。"' if C["temp"] else ('&"兩讀數取自 Siliconomics 協調試作分支（未合併）。"' if C["mode"] == "trial" else '&"兩讀數取自 ' + C["status_txt"] + '（對帳單 Drive 02-AI供應鏈/對帳/，期限 2027-02-28）。"'))); r += 1
    RC["pl"] = r; ws.cell(r, 1, "損益表年度標示（FY2027E–FY2031E）")
    for j, fy in enumerate(range(2027, 2032)):
        cy = next(y for y, fs in FY_OF.items() if fy in fs)
        ws.cell(r, 2 + j, f'=IF(AND(co_flag_{cy}=1,co_dep_{cy}=1),"FY{fy}E（對帳中）","FY{fy}E")')
    r += 2
    ws.cell(r, 1, "E. 敏感度：總量缺口全由 NVIDIA 吸收（規則 v1.2 C3；其他平台維持原始顆數，NVIDIA 承擔全部總量缺口，下限 0；兩讀數各算一次；整本重建快照）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "現行基準讀數：等比例分配", "現行基準讀數：NVIDIA 全吸收", "另一側讀數：等比例分配", "另一側讀數：NVIDIA 全吸收"], 1): ws.cell(r, j, h).font = BOLD
    r += 1; RE = {}
    for k, lab, fmt in [("nvu_B_2028", "NVIDIA 認列顆數（情境 B）FY2028", "0.00"), ("nvu_B_2029", "NVIDIA 認列顆數（情境 B）FY2029", "0.00"), ("nvu_B_2030", "NVIDIA 認列顆數（情境 B）FY2030", "0.00"),
                        ("eps_B_2028", "EPS 情境 B FY2028", "0.00"), ("eps_B_2029", "EPS 情境 B FY2029", "0.00"), ("eps_B_2030", "EPS 情境 B FY2030", "0.00"),
                        ("dcf", "DCF 目標價（情境加權）", "#,##0.00"), ("comp", "Comparable 目標價（進加權）", "#,##0.00"), ("tp", "加權目標價", "#,##0.00"), ("up", "上檔空間", "0.0%")]:
        RE[k] = r; ws.cell(r, 1, lab)
        ws.cell(r, 2, f"=B{RB[k]}"); ws.cell(r, 3, A["absorb"][k]); ws.cell(r, 4, f"=C{RB[k]}")
        ws.cell(r, 5, A["alt_absorb"][k] if A.get("alt_absorb") else ("不可行" if k == "tp" else "—"))
        for c in (2, 3, 4, 5): ws.cell(r, c).number_format = fmt
        for c in (3, 5): ws.cell(r, c).font = Font(color="7F7F7F")
        r += 1
    ws.cell(r, 1, "投資建議（五級固定規則）")
    for c in ("BCDE" if A.get("alt_absorb") else "BCD"): ws[f"{c}{r}"] = "=" + REC(f"{c}{RE['up']}")
    RE["rec"] = r; r += 1
    ws.cell(r, 1, "加權目標價：全吸收相對等比例分配"); ws.cell(r, 3, f"=C{RE['tp']}/B{RE['tp']}-1").number_format = '+0.0%;-0.0%'; ws.cell(r, 5, f"=E{RE['tp']}/D{RE['tp']}-1" if A.get("alt_absorb") else "—").number_format = '+0.0%;-0.0%'; RE["d"] = r; r += 1
    for y in APPLY:   # 可行性：總量缺口 ÷ NVIDIA 原始顆數（情境 B）；> 100% 即 NVIDIA 全部顆數也不夠吸收
        um = "1" if y == 2027 else f"um{y % 100}_B"
        ws.cell(r, 1, f"總量缺口 ÷ NVIDIA 原始顆數（情境 B）CY{y}")
        ws.cell(r, 3, f"=(1-{um})/co_ns_{y % 100}").number_format = "0.0%"
        ws.cell(r, 5, f"=(1-{um}*co_r_{y % 100})/co_ns_{y % 100}").number_format = "0.0%"
        RE[f"gap_{y}"] = r; r += 1
    ws.cell(r, 1, "全吸收是否可行（各年缺口 ÷ NVIDIA 原始顆數皆 ≤ 100%）")
    for c in "CE":
        ws[f"{c}{r}"] = "=IF(MAX(" + ",".join(f"{c}{RE[f'gap_{y}']}" for y in APPLY) + ')<=1,"可行","不可行（缺口大於 NVIDIA 全部顆數）")'
    RE["feas"] = r; r += 1
    ws.cell(r, 1, "讀法：總量缺口＝原始顆數總量 −（現行基準或另一側讀數的總量）。等比例分配＝各平台按原始顆數同比例縮（Siliconomics 做法）；全吸收＝AMD、TPU、Trainium、其他 ASIC 維持原始顆數，缺口全算在 NVIDIA（NVIDIA 占原始顆數的比例見 In 頁 Z 節），NVIDIA 顆數不低於 0。極端情境，只作敏感度。"); r += 2
    ws.cell(r, 1, "D. 快照一致性（另一側讀數為建置快照；In 頁 Z 節或模型改動而未重跑 python3 nvidia/coord_alt.py 時顯示 STALE）").font = BOLD; r += 1
    chk = []
    for y in APPLY:
        ws.cell(r, 1, f"快照所用比值 − 現行比值 CY{y}（應為 0）"); ws.cell(r, 2, f"={A['ratio'][str(y)]}-co_r_{y % 100}").number_format = "0.000000"; chk.append(f"ABS(B{r})<1E-9"); r += 1
    ws.cell(r, 1, "快照的現行基準加權目標價 − 本檔加權目標價（應為 0）"); ws.cell(r, 2, f"={A['base']['tp']}-B{RB['tp']}").number_format = "0.000000"; chk.append(f"ABS(B{r})<0.005"); r += 1
    RC["state"] = r; ws.cell(r, 1, "另一側讀數快照狀態").font = BOLD; ws.cell(r, 2, f'=IF(AND({",".join(chk)}),"CURRENT","STALE")'); r += 1
    return {"RA": RA, "RB": RB, "RC": RC, "RE": RE}


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
    RD["A1"].value = a1.replace("個股模型 v0.3.7", f"個股模型 v{VER}", 1).replace(
        "標準損益表、反向 DCF 頁）", f"標準損益表、反向 DCF 頁）；v{VER}：三層協調——對帳中標示與兩讀數目標價（規則 v1.4 C4；Coord 頁、In 頁 Z 節；基準數值不變）", 1)
    return f"build_id {old} → {new}（{n} 格）"


if __name__ == "__main__":   # python3 nvidia/coord_post.py extract <快照 json> <分支> <SHA>
    import sys
    if sys.argv[1] == "extract":
        print("摘錄", extract(sys.argv[2], sys.argv[3], sys.argv[4]), "個名稱")
