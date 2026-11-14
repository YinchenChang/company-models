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
            row[8].value = f"Siliconomics 三層協調（NVIDIA v{NOTE_VER}）：改讀 {src}；原值 {old}（Assumed，原說明：{row[8].value or ''}）"
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
    r += 1
    # v0.3.9（Andy 11-14 17:47 選項 A）：損益章 FY2029 下滑說明——2028 基準係數待定維持現行、等比例套用；重判後更新
    RC["pl29"] = r; ws.cell(r, 1, "損益章說明：FY2029 營收下滑（2028 係數待定；選項 A）")
    ws.cell(r, 2, f'="FY2029E 營收較 FY2028E "&TEXT(B{RB["rev_B_2029"]}/B{RB["rev_B_2028"]}-1,"+0.0%;-0.0%")&"：2028 年 Siliconomics 基準顆數係數仍為「待定」，依規則維持現行 "&TEXT(co_ucoef_28,"0.00")&"（2027 年為 1），並等比例套用到各平台；NVIDIA 原始顆數 CY2027→CY2028 只增 "&TEXT((u_bw_28+u_rb_28+u_ru_28+u_hp_28)/(u_bw_27+u_rb_27+u_ru_27+u_hp_27)-1,"0.0%")&"（全市場原始顆數 +"&TEXT(((u_bw_28+u_rb_28+u_ru_28+u_hp_28)/co_ns_28)/((u_bw_27+u_rb_27+u_ru_27+u_hp_27)/co_ns_27)-1,"0.0%")&"，成長主要來自其他平台），乘係數後認列顆數 FY2028→FY2029 "&TEXT(B{RB["nvu_B_2029"]}/B{RB["nvu_B_2028"]}-1,"+0.0%;-0.0%")&"。這是 2028 年狀態未定造成的斷層，不是需求轉弱的判斷；待 NVIDIA FY27Q3 財報（2026-11-19）與財報季結束（2026-11-30）重判 2028 年狀態後更新，屆時一併評估 2028 年改套連續法、或總量缺口改依各平台成長分配兩案。"')
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
    r += 1
    RF = write_c0(ws, wb, r, L)
    return {"RA": RA, "RB": RB, "RC": RC, "RE": RE, "RF": RF}



def write_c0(ws, wb, r, L):
    """F 節（v0.3.9）：三層協調——讀取層級與 C0 由下往上對帳（只新增；個股不裁決、不改上游值）。
    F1 逐項列出本模型讀的 Siliconomics 名稱、值、版本與 Tokenomics 鏡像值；F2 本模型自己的量 vs Siliconomics 對應值（差距超過 ±15% 標「對帳中」）。"""
    import reader as SIL   # Siliconomics 快照（與 build_nvidia 同一份）
    N, V = SIL.NAMES, SIL.VERSION
    TKV = "Tokenomics v21.10"   # Siliconomics v0.19 鏡像的 Tokenomics 版本（siliconomics/CHANGELOG v0.19 標題）
    rp = os.path.join(os.path.dirname(os.path.abspath(SIL.__file__)), "data", f"recon_snapshot_{V}.json")
    NVD = json.load(open(rp, encoding="utf-8"))["markets"]["nv_dc"]
    NVD_I = dict(zip(NVD["years"], NVD["implied"])); NVD_C = dict(zip(NVD["years"], NVD["company"]))
    yy = lambda y: y % 100
    ix = lambda y: y - 2025   # 年度陣列：[None, 2026, 2027, 2028, 2029]
    GAP = '+0.0%;-0.0%'
    def hdr(row, cols):
        for j, h in enumerate(cols, 1): ws.cell(row, j, h).font = BOLD
    ws.cell(r, 1, f"F. 三層協調：讀取層級與 C0 由下往上對帳（三層基準協調規則 v1.4 C0、C4、C5；Siliconomics {V}、{TKV} 鏡像；只新增，不改任何上游值或既有格）").font = BOLD; r += 1
    ws.cell(r, 1, "層級：Tokenomics（權威：每 GW 顆數、機架價）→ Wattonomics（產業部署 GW）→ Siliconomics（行業變數；本模型只讀它）→ 個股。下游只註明版本、不自行重估（Instructions 第 3 條）；"
                  "差距超過門檻者標「對帳中」，個股不裁決、不改上游，證據回饋走對帳單（C5）。"); r += 2
    ws.cell(r, 1, "F1. 本模型讀取的 Siliconomics 名稱、值與版本").font = BOLD; r += 1
    hdr(r, ["名稱／項目", "年度", f"層級值（Siliconomics {V} 快照）", "本模型採用值（讀取格）", "採用 − 層級值（應為 0）", "單位", "層級與版本", "讀取方式與說明"]); r += 1
    F1a = r; R1 = {}
    def f1(name, y, val, cell, unit, layer, how, fmt="0.000"):
        nonlocal r
        R1[(name, y)] = r
        ws.cell(r, 1, name); ws.cell(r, 2, y if isinstance(y, str) else f"CY{y}")
        ws.cell(r, 3, val).number_format = fmt
        if cell is not None:
            ws.cell(r, 4, cell).number_format = fmt
            ws.cell(r, 5, f"=D{r}-C{r}").number_format = "0.000000"
        else:
            ws.cell(r, 4, "不直接讀取")
        ws.cell(r, 6, unit); ws.cell(r, 7, layer); ws.cell(r, 8, how); r += 1
    SL = f"Siliconomics {V}"
    for y in YEARS:
        f1("SIF_DeployGW_Base", y, N["SIF_DeployGW_Base"][YEARS.index(y)], f"=co_gwb_{yy(y)}", "GW", SL, "In 頁 Z 節 co_gwb_<yy>；現行基準產業部署 GW（兩讀數的分母）")
    for y in YEARS:
        f1("SIF_DeployGW_Supply_<yyyy>", y, N[f"SIF_DeployGW_Supply_{y}"], f"=co_gws_{yy(y)}", "GW", SL, "In 頁 Z 節 co_gws_<yy>；供給側讀數（情境 A）")
    for y in YEARS:
        f1("SIF_DeployGW_Alt_<yyyy>", y, N[f"SIF_DeployGW_Alt_{y}"], f"=co_gwa_{yy(y)}", "GW", SL + "（Wattonomics v0.3.2 換算值）", "In 頁 Z 節 co_gwa_<yy>；另一側讀數，只進兩讀數目標價")
    for y in YEARS:
        f1("SIF_DeployGW_AltRatio_<yyyy>", y, N[f"SIF_DeployGW_AltRatio_{y}"], f"=co_r_{yy(y)}", "x", SL, "In 頁 Z 節 co_r_<yy>；另一側 ÷ 現行基準")
    for y in YEARS:
        f1("SIF_RecoFlag_<yyyy>", y, N[f"SIF_RecoFlag_{y}"], f"=co_flag_{y}", "旗標", SL, "In 頁 Z 節 co_flag_<yyyy>；1＝對帳中（Siliconomics 對帳單期限 2027-02-28）", "0")
    for y in (2028, 2029):
        f1("SIF_UnitsCoef_Base", y, N["SIF_UnitsCoef_Base"][ix(y)], f"=co_ucoef_{yy(y)}", "x", SL, "In 頁 Z 節 co_ucoef_<yy> → um<yy>_B（情境 B 顆數係數；2026–27 係數固定為 1）")
    NV = ("bw", "rb", "ru", "hp")
    u_m = lambda y: "+".join((f"u_{k}_{yy(y)}" if y <= 2028 else f"u_{k}_28*(1+g29_{k})") for k in NV)
    for y in YEARS:
        f1("SIF_UnitsPlat_bw＋rb＋ru＋hp（NVIDIA 平台顆數，原始）", y, sum(N[f"SIF_UnitsPlat_{k}"][ix(y)] for k in NV), "=" + u_m(y), "百萬顆", SL,
           "間接讀：In 頁 u_<平台>_<yy>（markets/memory UNITS ← 台積電模型 In 頁；2029＝2028 ×（1＋g29））；與快照逐值相同", "0.0000")
    for k, nm in (("amd", "AMD"), ("tpu", "TPU"), ("trn", "Trainium"), ("oth", "其他 ASIC")):
        f1(f"SIF_pa_{k}（{nm} 每顆售價，基準）", "—", N[f"SIF_pa_{k}"], f"=pa_{k}", "US$K", SL, f"In 頁 pa_{k}（SIL.row9）；份額計算用", "0.0")
    for y in (2026, 2027):
        f1("Recon 快照 nv_dc.implied（隱含 NVIDIA 資料中心營收）", y, NVD_I[y], None, "US$B", SL + " Recon 快照", "R4 橋接 n3＝NVIDIA 成本等值 GW × 每 GW 營收內容（Siliconomics 內部對帳點；本模型只作 F2 對照，不讀入模型）", "#,##0.000")
    for y in (2026, 2027):
        f1("Recon 快照 nv_dc.company（公司資料中心營收）", y, NVD_C[y], None, "US$B", SL + " Recon 快照", "R4 橋接 n4（2026＝FY2027 錨點、2027＝FY2028 指引推得）；與本模型錨點同一來源", "#,##0.000")
    for g, rk in (("GB300", "bw"), ("VR200", "rb"), ("RU", "ru")):
        f1(f"SIF_TK_RackPrice_{g}", "—", N[f"SIF_TK_RackPrice_{g}"], None, "US$M／機架", f"{TKV}（Siliconomics 鏡像）", "本模型不直接讀；GPU 模組價 pm_" + rk + " 讀 chip_bridge v0.2.3（由 Tokenomics 機架價往下推），見 F2 對照", "0.0")
        f1(f"SIF_TK_GPUsPerGW_{g}", "—", N[f"SIF_TK_GPUsPerGW_{g}"], None, "顆／GW", f"{TKV}（Siliconomics 鏡像）", "本模型不直接讀（只用顆數，不換 GW；GW 由 Siliconomics 供給）", "#,##0")
        f1(f"SIF_TK_RacksPerGW_{g}", "—", N[f"SIF_TK_RacksPerGW_{g}"], None, "機架／GW", f"{TKV}（Siliconomics 鏡像）", "同上；機架價 ÷（每 GW 顆數 ÷ 每 GW 機架數）＝每顆機架級金額", "#,##0")
    F1b = r - 1
    ws.cell(r, 1, "F1 核對：Σ|採用 − 層級值|（應為 0）"); ws.cell(r, 5, f"=SUMPRODUCT(ABS(E{F1a}:E{F1b}))").number_format = "0.000000"; RF_chk = r; r += 2

    # ---------- F2 ----------
    ws.cell(r, 1, "F2. C0 由下往上對帳：本模型自己的量 vs Siliconomics 對應值（差距＝本模型 ÷ Siliconomics − 1；超過門檻標「對帳中」；不改上游）").font = BOLD; r += 1
    ws.cell(r, 1, "對帳門檻（±）"); ws.cell(r, 4, 0.15).number_format = "0%"; ws.cell(r, 4).font = BLUE; ws.cell(r, 8, "三層基準協調規則 C0／工作單：差距超過 ±15% 標「對帳中」；門檻為輸入（藍字）"); TOL = f"$D${r}"; r += 1
    ws.cell(r, 1, "FY 月數權重（CY(t) 占 FY(t+1) 的比例）"); ws.cell(r, 4, "=11/12").number_format = "0.0000"; ws.cell(r, 4).font = BLUE; ws.cell(r, 8, "本模型會計年度 FY(t+1)＝11/12 × CY(t)＋1/12 × CY(t+1)（FY↔CY 月數加權，v0.3.6）；層級值換成同一口徑再比"); WC = f"$D${r}"; r += 1
    hdr(r, ["項目", "年度（日曆年／會計年度）", "單位", "本模型值", "Siliconomics 值", "差距 %", "狀態", "本模型口徑與來源", "本模型證據等級", "Siliconomics 口徑與來源", "Siliconomics 證據等級"]); RF_h = r; r += 1
    ra = lambda k: int("".join(c for c in L[k][1] if c.isdigit()))
    r_nvu = ra("nvu_B_2028")
    # 情境 B 的運算與網通營收列（依標籤自情境 B 區塊起找）
    def fr(prefix, start):
        sh = wb["PL"]
        return next(x for x in range(start, sh.max_row + 1) if isinstance(sh.cell(x, 1).value, str) and sh.cell(x, 1).value.strip().startswith(prefix))
    rb0 = next(x for x in range(1, wb["PL"].max_row + 1) if isinstance(wb["PL"].cell(x, 1).value, str) and wb["PL"].cell(x, 1).value.startswith("情境 B："))
    r_comp, r_net = fr("運算營收（Σ", rb0), fr("NVIDIA 網通營收", rb0)
    SUBT = "同一缺口的分項，不另計"
    F2a = r
    def f2(item, y, unit, mv, sv, mfmt, mdesc, mtag, sdesc, stag, judge=True, sub=False, main=False):
        nonlocal r
        ws.cell(r, 1, item); ws.cell(r, 2, y); ws.cell(r, 3, unit)
        ws.cell(r, 4, mv).number_format = mfmt
        ws.cell(r, 5, sv).number_format = mfmt
        numeric = not isinstance(sv, str) or sv.startswith("=")
        if judge and numeric:
            ws.cell(r, 6, f"=D{r}/E{r}-1").number_format = GAP
            ws.cell(r, 7, f'=IF(ABS(F{r})>{TOL},"{SUBT if sub else "對帳中"}","一致")')
        else:
            ws.cell(r, 6, "—"); ws.cell(r, 7, "不判定（口徑不同）" if numeric else "無可比的層級量")
        ws.cell(r, 8, mdesc); ws.cell(r, 9, mtag); ws.cell(r, 10, sdesc); ws.cell(r, 11, stag); r += 1
    SU = f"Siliconomics {V} In 頁 M6 起的 UNITS（SIF_UnitsPlat_bw／rb／ru／hp）"
    st_u = "Interested-party（賣方與研調估計；Rubin Ultra 為 Assumed）"
    C1 = lambda name, y: f"C{R1[(name, y)]}"
    PL_ = "SIF_UnitsPlat_bw＋rb＋ru＋hp（NVIDIA 平台顆數，原始）"
    base_cy = lambda y: C1(PL_, y) if y < 2028 else f"{C1(PL_, y)}*{C1('SIF_UnitsCoef_Base', y)}"   # 基準 NVIDIA 平台顆數（日曆年；2026–27 係數 1）
    fy_w = lambda y: f"({WC}*({base_cy(y)})+(1-{WC})*({base_cy(y + 1)}))" if y < 2029 else f"({base_cy(y)})"   # FY(y+1)＝11/12 CY(y)＋1/12 CY(y+1)；FY2030 缺 CY2030，用 CY2029
    IMP, CMP = "Recon 快照 nv_dc.implied（隱含 NVIDIA 資料中心營收）", "Recon 快照 nv_dc.company（公司資料中心營收）"
    # (0) 唯一的對帳項目：認列係數校準缺口（對帳中判定只算一次）
    f2("營收校準認列係數缺口：Siliconomics 由下而上推得的 NVIDIA 顆數與營收高於公司錨點", "FY2027 錨點校準", "x（本模型 ÷ 層級）", "=k_rec", 1, "0.000",
       "認列係數 k_rec＝公司 FY2027 資料中心營收錨點 ÷ 未校準由下而上營收（Derived）；Recon 頁 B 節。差距＝k_rec − 1；層級值＝1 表示層級顆數、每顆內容照單全收", "Derived（錨點為公司實績，Verified／Assumed）",
       f"{SL}：NVIDIA 平台顆數（賣方與研調估計）與 R4 隱含營收（n5 比值：CY2026＝{NVD_I[2026] / NVD_C[2026]:.3f}、CY2027＝{NVD_I[2027] / NVD_C[2027]:.3f}，未加權即高於公司值約 {100 * (NVD_I[2026] / NVD_C[2026] - 1):.0f}–{100 * (NVD_I[2027] / NVD_C[2027] - 1):.0f}%；FY 月數加權後 FY2027 高約 {100 * ((11 / 12 * NVD_I[2026] + NVD_I[2027] / 12) / NVD_C[2026] - 1):.0f}%）", "Interested-party／Derived", main=True)
    # (1) 市場層顆數（讀同一份）
    for y in YEARS:
        f2("NVIDIA 平台顆數（市場層，認列係數前）", f"CY{y}", "百萬顆", "=" + u_m(y), "=" + C1(PL_, y), "0.000",
           "In 頁 u_bw／rb／ru／hp（markets/memory UNITS ← 台積電模型 In 頁）；2029＝2028 ×（1＋g29，Assumed）；日曆年", "Interested-party", SU + "；原始顆數（未乘基準係數）；日曆年", st_u)
    # (2) 營收校準等值顆數（情境 B）vs 基準顆數（FY 月數加權）：差距來源＝認列係數校準缺口
    for y, c in zip(YEARS, "DEFG"):
        f2("營收校準等值顆數（情境 B；非實體顆數）vs Siliconomics 基準 NVIDIA 平台顆數（FY 月數加權）", f"FY{y + 1}（CY{y}）", "百萬顆", f"=PL!{c}{r_nvu}", "=" + fy_w(y), "0.000",
           "PL 頁情境 B：k_rec × 市場層顆數 × 基準係數（市場層已依 FY 月數加權）。是為對上公司營收錨點而校準出的等值顆數，不是實體出貨顆數；與層級值的差距＝認列係數校準缺口（k_rec − 1），不是新證據", "Derived（由營收錨點反推）",
           SU + " × SIF_UnitsCoef_Base，換成同一 FY 月數加權口徑：FY(t+1)＝" + "11/12 × CY(t)＋1/12 × CY(t+1)" + ("（FY2030 缺 CY2030 層級值，只用 CY2029；本列差距仍等於認列係數 − 1，與是否加權無關）" if y == 2029 else ""), st_u, sub=True)
    # (3) 資料中心營收：同一缺口的分項
    f2("資料中心營收 FY2027 錨點 vs R4 公司錨點（同一輸入）", "FY2027（CY2026）", "US$B", "=a27_dc", "=" + C1(CMP, 2026), "#,##0.0",
       "Q 頁 FY2027 錨點＝Q1、Q2 公司實績＋Q3 指引中點＋Q4 假設（Q!K5）", "Verified／Assumed（Q4 為假設）",
       f"{SL} Recon 快照 nv_dc.company（R4 橋接 r11 n4＝nv_dc_26）", "Verified／Assumed")
    f2("資料中心營收 FY2027 錨點 vs Siliconomics 隱含（FY 月數加權；R4 n5 比值的另一面，不是個股新證據）", "FY2027（CY2026／CY2027 加權）", "US$B", "=a27_dc", f"={WC}*{C1(IMP, 2026)}+(1-{WC})*{C1(IMP, 2027)}", "#,##0.0",
       "同上（公司實績加指引的錨點）", "Verified／Assumed",
       f"{SL} Recon 快照 nv_dc.implied（R4 n3）換成 FY 月數加權：11/12 × CY2026＋1/12 × CY2027；R4 本身以 FY2027≈CY2026 比（差距 −16%，口徑差約 4 個百分點）", "Derived（顆數、每 GW 內容多屬 Interested-party）", sub=True)
    f2("資料中心營收 FY2028 公司指引推得 vs R4 公司值", "FY2028（CY2027）", "US$B", "=a27_dc*(1+g28_gr)", "=" + C1(CMP, 2027), "#,##0.0",
       "FY2027 錨點 ×（1＋公司 FY2028 營收年增指引 70%）；假設資料中心占比同 FY2027", "Interested-party（指引二手報導）",
       f"{SL} Recon 快照 nv_dc.company（nv_dc_27＝FY2028 營收約 +70% × 資料中心 92.5%）", "Derived（Interested-party 指引）")
    f2("資料中心營收 FY2028 情境 B（運算＋網通）vs Siliconomics 隱含（未加權）", "FY2028（CY2027）", "US$B", f"=PL!E{r_comp}+PL!E{r_net}", "=" + C1(IMP, 2027), "#,##0.0",
       "PL 頁情境 B：運算營收（認列顆數 × 每顆售價）＋網通營收（中國為 0）", "Derived（情境 B 為模型基準）",
       f"{SL} Recon 快照 nv_dc.implied（n1 × n2，VR200 2027 HBM 價內容）；缺 CY2028 層級值故未做 FY 月數加權，FY2028 有 1/12 屬 CY2028，口徑差預期約 1–2 個百分點", "Derived（顆數、每 GW 內容多屬 Interested-party）", sub=True)
    f2("資料中心營收 FY2029（CY2028）情境 B", "FY2029（CY2028）", "US$B", f"=PL!F{r_comp}+PL!F{r_net}", "無可比的層級量", "#,##0.0",
       "PL 頁情境 B FY2029 運算＋網通（FY2030 以後見 PL 頁）", "Derived", f"{SL} R4 橋接 nv_dc 對帳點只到 CY2027（Recon 快照 implied 2028、2029＝空）；CY2028 以後只能靠上列顆數對帳", "—")
    # (4) 模組價 vs Tokenomics 機架價
    for g, k, pm in (("GB300", "bw", "pm_bw"), ("VR200", "rb", "pm_rb")):
        per = f"={C1(f'SIF_TK_RackPrice_{g}', '—')}*1000/({C1(f'SIF_TK_GPUsPerGW_{g}', '—')}/{C1(f'SIF_TK_RacksPerGW_{g}', '—')})"
        f2(f"GPU 每顆金額：模組價 vs Tokenomics 機架價 ÷ 每機架顆數（{g}）", "2026", "US$K", f"={pm}", per, "0.0",
           "In 頁（chip_bridge v0.2.3 Rack 頁）：GPU 模組售價（NVIDIA 對買方）", "Derived", f"{TKV} 鏡像：SIF_TK_RackPrice_{g} × 1000 ÷（SIF_TK_GPUsPerGW ÷ SIF_TK_RacksPerGW）；整機架口徑（含 CPU、網通、機櫃）", "Interested-party（Tokenomics 機架價估計）", judge=False)
    # (5) 無可比
    f2("供應與產能承諾（10-Q Note 10）vs 層級量", "2026-07-26", "US$B", "=h_supply_c", "無可比的層級量", "#,##0",
       "Hist 頁 h_supply_c（NVIDIA 10-Q 供應承諾，多年期、含非 GPU 項目）", "Verified", "Siliconomics 無採購承諾變數；承諾金額為多年累計義務，與任一年的顆數或營收不同口徑", "—")
    f2("產業部署 GW 與 NVIDIA 成本等值 GW", "CY2026–2029", "GW", "不自建（只讀 Siliconomics）", "無可比的層級量", "0.0",
       "本模型以顆數建模，GW 只經 F1 的 Z 節讀入並用於兩讀數比值", "—", "R4 橋接 n1（NVIDIA 成本等值 GW）未進 Recon 快照，無法對帳", "—")
    F2b = r - 1
    ws.cell(r, 1, "對帳中項目數（差距超過門檻；同一缺口只算一次）"); ws.cell(r, 4, f'=COUNTIF(G{F2a}:G{F2b},"對帳中")'); RF_n = r; r += 1
    ws.cell(r, 1, "同一缺口的分項數（不另計）"); ws.cell(r, 4, f'=COUNTIF(G{F2a}:G{F2b},"{SUBT}")'); RF_sub = r; r += 1
    ws.cell(r, 1, "有判定的項目數（一致＋對帳中；分項不計）"); ws.cell(r, 4, f'=COUNTIF(G{F2a}:G{F2b},"一致")+COUNTIF(G{F2a}:G{F2b},"對帳中")'); RF_j = r; r += 1
    ws.cell(r, 1, "對帳摘要（HTML 用）")
    ws.cell(r, 4, f'="由下往上對帳：有判定 "&D{RF_j}&" 項、對帳中 "&D{RF_n}&" 項（差距超過 ±"&TEXT({TOL},"0%")&"）。唯一的對帳中項目是營收校準認列係數缺口（認列係數 "&TEXT(k_rec,"0.000")&"，即 Siliconomics 由下而上推得的 NVIDIA 顆數與營收高於公司錨點：顆數面認列係數比層級值低 "&TEXT(1-k_rec,"0%")&"，R4 營收面隱含值（FY 月數加權）高於公司錨點約 "&TEXT(({WC}*{C1(IMP, 2026)}+(1-{WC})*{C1(IMP, 2027)})/a27_dc-1,"0%")&"；未加權的 R4 比值為 "&TEXT({C1(IMP, 2026)}/{C1(CMP, 2026)}-1,"0%")&"／"&TEXT({C1(IMP, 2027)}/{C1(CMP, 2027)}-1,"0%")&"）；營收校準等值顆數與資料中心營收的 "&D{RF_sub}&" 列差距是同一缺口的分項，不另計。市場層顆數、公司錨點與指引推得值與 Siliconomics 一致。個股不裁決，不改上游值；證據回饋走對帳單（C5）。"'); RF_s = r; r += 1
    ws.cell(r, 1, "讀法：營收校準等值顆數＝營收校準係數 × Siliconomics 顆數，所以顆數差距必然等於係數 − 1（循環，不是獨立證據）；資料中心營收的差距是 Siliconomics 內部 R4 的 n5 比值（本模型錨點＝R4 公司值）。證據方向（公司實績 vs 賣方與研調估計）依規則 C0 第 1 點由主控依對帳單處理；層級值已換成與本模型相同的 FY 月數加權口徑（FY2028 營收與 FY2030 顆數因缺後一年層級值未加權；FY2030 顆數差距仍等於係數 − 1，FY2028 營收口徑差預期約 1–2 個百分點）。")
    return {"F1": (F1a, F1b), "F1chk": RF_chk, "F2h": RF_h, "F2": (F2a, F2b), "n": RF_n, "sum": RF_s}

VER = "0.3.9"
NOTE_VER = "0.3.8"   # In 頁既有說明欄的版本字樣維持 v0.3.8（本版只新增 Coord 頁 F 節）


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
        "標準損益表、反向 DCF 頁）", f"標準損益表、反向 DCF 頁）；v0.3.8：三層協調——對帳中標示與兩讀數目標價（規則 v1.4 C4；Coord 頁、In 頁 Z 節；基準數值不變）；v{VER}：Coord 頁新增 F 節「三層協調：讀取層級與 C0 由下往上對帳」（只新增；數值不變）", 1)
    return f"build_id {old} → {new}（{n} 格）"


if __name__ == "__main__":   # python3 nvidia/coord_post.py extract <快照 json> <分支> <SHA>
    import sys
    if sys.argv[1] == "extract":
        print("摘錄", extract(sys.argv[2], sys.argv[3], sys.argv[4]), "個名稱")
