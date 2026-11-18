# 三層基準協調狀態檔匯出（v0.19 起；規則 docs/plan/三層基準協調規則.md v1.4；欄位依 tools/baseline_coordination/README.md）
# 用法：python3 siliconomics/tools/export_baseline_status.py [<重算後 Siliconomics xlsx>] [<輸出 json>]
#   預設讀 model/CURRENT、寫 data/baseline_status.json。只讀 Interface 的 SIF_ 名稱（Coord 頁推導），本程式不做任何判定或計算。
#   對帳單與期限（recon_doc、deadline）、硬體個股清單與依賴年份讀 builder/seed/coord_meta_v*.json（判斷類，未定者留空）。
#   v0.20（規則 v1.4.3）：頂層 wattonomics_version（取自 wattonomics/model/CURRENT 檔名）；逐年 c0_fail（讀 SIF_RecoC0Fail_<年>）與
#   c0_recon_doc、c0_deadline（meta c0）；頂層 c0_actuals.<年>（協調年份以外的已公布年份；值讀 Wattonomics 快照 WIF_Recon_NV_*，對帳單與期限讀 meta）。
#   2026-11-18（三層對帳工作單 A1、A8；不改任何基準值）：逐年 timing_cum（C2 ② 累計 GW）與 timing_cum_ratio（② 累計 ÷ 當年基準）讀協調橋接檔
#   Bridge 頁 g2c、g2p 列（檔案＝docs/reports/*協調橋接_<Wattonomics 版本>*.xlsx 最新一份）；頂層新增 method、generated、source_sha256、
#   wattonomics_source_sha256、bridge（檔名、sha256、橋接當時讀的 Siliconomics 版本）。
#   v0.20 r1（2026-11-18；工作單 W2-05B 的 Siliconomics 部分）：逐年 observations＝Coord 頁 D 節觀測數列（中間估計 mid i／ii／iii／lo／hi、衰減寫法 λ＝0／0.5／1），
#   display_only＝true、rule＝"v1.5"（草案，待 Andy 核可）；不適用年份（非待定／非連續法）為 null。D 節不輸出 Interface 名稱（W2-05A），故本欄依 Coord 頁 A 欄鍵讀值；
#   只供顯示，個股不得讀（規則 v1.5 草案第 10 節第 15 項，check.py 由共用區另案實作）。legacy_s33 已可讀到新名 SIF_DeployGW_LegacyS33_<年>。
import datetime, glob, hashlib, json, os, re, sys
import openpyxl

SIL = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
YRS = [2026, 2027, 2028, 2029]


def sha256(path):
    return hashlib.sha256(open(path, "rb").read()).hexdigest()


def read_bridge(wt_ver):
    """協調橋接檔（baseline_bridge.py 的輸出）。Bridge 頁 g2c＝C2 ② 累計 GW，g2p＝② 累計 ÷ 當年基準（檢查第 7 項）。
    回傳 (路徑, {年: (g2c, g2p)}, In 頁 A1 標題)；找不到檔案回傳 (None, {}, None)，check.py 第 7 項照舊印「未量測」。"""
    cand = sorted(glob.glob(os.path.join(SIL, "docs", "reports", f"*協調橋接_{wt_ver}*.xlsx")))
    if not cand:
        return None, {}, None
    bp = cand[-1]
    bw = openpyxl.load_workbook(bp, data_only=True)
    rows = {str(r[0]): r for r in bw["Bridge"].iter_rows(values_only=True) if r and r[0]}
    # 欄：A 鍵、B 項目、C 說明、D＝2025、E..H＝2026..2029（tuple 索引 4..7）
    out = {y: (rows["g2c"][4 + i], rows["g2p"][4 + i]) for i, y in enumerate(YRS)}
    return bp, out, bw["In"]["A1"].value


def main():
    x = sys.argv[1] if len(sys.argv) > 1 else os.path.join(SIL, "model", open(os.path.join(SIL, "model/CURRENT")).read().strip())
    out = sys.argv[2] if len(sys.argv) > 2 else os.path.join(SIL, "data", "baseline_status.json")
    wb = openpyxl.load_workbook(x, data_only=True)

    def v(nm):
        sh, ref = list(wb.defined_names[nm].destinations)[0]
        c = wb[sh][ref.replace("$", "").split(":")[0]]
        return c.value

    meta = {}
    for f in sorted(glob.glob(os.path.join(SIL, "builder", "seed", "coord_meta_v*.json"))):
        meta.update(json.load(open(f, encoding="utf-8")))
    prev = lambda k: None if v(k) in (None, "") else bool(v(k))

    def three(k):
        """證據比例三值（規則 v1.3；check.py 接受 {lo, base, hi}）；沒有 _lo／_hi 名稱時以點值代入。"""
        b = v(k)
        lo = v(k + "_lo") if k + "_lo" in wb.defined_names else b
        hi = v(k + "_hi") if k + "_hi" in wb.defined_names else b
        return {"lo": b if lo is None else lo, "base": b, "hi": b if hi is None else hi}
    # Coord 頁 D 節（觀測數列；只顯示）：A 欄鍵 → 列號；欄 C–F＝2026–2029
    CO = wb["Coord"] if "Coord" in wb.sheetnames else None
    corow = {CO.cell(r, 1).value: r for r in range(1, CO.max_row + 1) if CO.cell(r, 1).value} if CO is not None else {}

    def obs(y):
        if "mid_i" not in corow:
            return None
        col = 3 + YRS.index(y)
        num = lambda k: (lambda x: x if isinstance(x, (int, float)) and not isinstance(x, bool) else None)(CO.cell(corow[k], col).value)
        return {"mid": {"i": num("mid_i"), "ii": num("mid_ii"), "iii": num("mid_iii"), "lo": num("mid_lo"), "hi": num("mid_hi")},
                "decay": {"0": num("dec0"), "0.5": num("dec05"), "1": num("dec1")},
                "display_only": True, "rule": "v1.5",
                "note": "規則 v1.5 草案 C3（待 Andy 核可）；權重未校準；只顯示，不得作為基準或個股讀數；Coord 頁 D 節（不輸出 Interface 名稱）；非適用年份為 null"}
    root = os.path.dirname(SIL)
    wt_file = open(os.path.join(root, "wattonomics", "model", "CURRENT")).read().strip()
    wt_ver = re.search(r"v\d+(?:\.\d+)+", wt_file).group(0)
    bridge_path, timing, bridge_title = read_bridge(wt_ver)
    years = {}
    coefs = []
    for y in YRS:
        yy = y % 100
        st = v(f"SIF_RecoStatus_{y}")
        lead = v(f"SIF_RecoLead_{y}")
        rec = {
            "wattonomics": {"evidence_share": three(f"SIF_crd_ew_{yy}"), "prev_has_evidence": prev(f"SIF_crd_prevw_{yy}")},
            "siliconomics": {"evidence_share": three(f"SIF_crd_es_{yy}"), "prev_has_evidence": prev(f"SIF_crd_prevs_{yy}")},
            "status": st,
            "lead": ("wattonomics" if lead == "Wattonomics" else "siliconomics") if st == "乙" else None,
            "label": v(f"SIF_RecoLabel_{y}"),
            "converted": v(f"SIF_DeployGW_Demand_{y}"),
            "supply_side": v(f"SIF_DeployGW_Supply_{y}"),
            # 舊 S33 係數取法（v0.18 基準；只供對照，不是現行基準，見 README「欄位語義」）。優先讀語義明確的新名 SIF_DeployGW_LegacyS33_<年>（下一版新增），
            # 沒有就讀舊名 SIF_DeployGW_Current_<年>；current_s33 保留一版作相容別名（無讀取者），之後刪除。
            "legacy_s33": v(f"SIF_DeployGW_LegacyS33_{y}" if f"SIF_DeployGW_LegacyS33_{y}" in wb.defined_names else f"SIF_DeployGW_Current_{y}"),
            "current_s33": v(f"SIF_DeployGW_Current_{y}"),
            "use_converted": bool(v(f"SIF_RecoUseConv_{y}")),
            "continuity": bool(v(f"SIF_RecoChain_{y}")) if f"SIF_RecoChain_{y}" in wb.defined_names else False,
            "reconciling": bool(v(f"SIF_RecoFlag_{y}")),
            "recon_doc": meta.get("recon_doc", {}).get(str(y)),
            "deadline": meta.get("deadline", {}).get(str(y)),
            "c0_fail": bool(v(f"SIF_RecoC0Fail_{y}")) if f"SIF_RecoC0Fail_{y}" in wb.defined_names else False,
            "supply_cap_check_only": "供給上限非獨立（SIF_CapOnlineGW_Base＝本庫自身推得），依規則 v1.1 只作檢查、不封頂；故不填 supply_cap",
        }
        # 基準部署 GW（Interface 向量 SIF_DeployGW_Base 欄 E–H＝2026–2029）
        sh, ref = list(wb.defined_names["SIF_DeployGW_Base"].destinations)[0]
        row = [c.value for c in wb[sh][ref.replace("$", "")][0]]
        rec["sil_base"] = row[YRS.index(y)]
        c0 = meta.get("c0", {}).get(str(y))
        if rec["c0_fail"]:   # 規則 v1.4.3 C0 第 8 點：實績不過須登錄對帳單與期限（甲年標示由 Coord 頁改「有條件一致」）
            rec["c0_recon_doc"] = (c0 or {}).get("recon_doc")
            rec["c0_deadline"] = (c0 or {}).get("deadline")
            rec["c0_point"] = (c0 or {}).get("point")
        if y in timing and timing[y][0] is not None:   # C2 ② 累計（單位 GW；check.py 第 7 項以 |timing_cum| ≤ 15% × sil_base 判定）
            rec["timing_cum"], rec["timing_cum_ratio"] = timing[y]
        o_ = obs(y)
        if o_ is not None:
            rec["observations"] = o_
        years[str(y)] = rec
        if not rec["use_converted"] and y >= 2028:
            coefs.append({"name": f"um{yy}_B（S33；經 crd_k_{yy} 帶入 {y} 基準）", "year": y, "tag": "Assumed",
                          "note": "覆蓋未補齊、該年維持現行（對帳中）；改用換算值後不再在基準鏈上"})
    wsnap = json.load(open(os.path.join(root, "wattonomics", "data", open(os.path.join(root, "wattonomics", "data", "CURRENT_SNAPSHOT")).read().strip()), encoding="utf-8"))["names"]
    c0a = {}
    for yr, m in (meta.get("c0_actuals") or {}).items():
        ratio = wsnap.get(f"WIF_Recon_NV_Ratio_{yr}")
        rec = {"layer": m.get("layer"), "point": m.get("point"),
               "implied": wsnap.get(f"WIF_Recon_NV_ImpliedRev_{yr}"), "actual": wsnap.get(f"WIF_Recon_NV_ActualRev_{yr}"), "ratio": ratio,
               "c0_fail": ratio is not None and abs(ratio - 1) > v("SIF_rc_tol")}
        if rec["c0_fail"]:
            rec["c0_recon_doc"] = m.get("recon_doc"); rec["c0_deadline"] = m.get("deadline")
        c0a[yr] = rec
    st = {
        "_note": "由 siliconomics/tools/export_baseline_status.py 自 " + os.path.basename(x) + " 匯出；換算值與 Wattonomics 證據比例（含 v0.3.6 起的低／高區間）取自 Wattonomics " + wt_ver + "（" + wt_file + "）；continuity＝規則 v1.4 連續法（基準＝前一年基準 × 換算值成長率，非換算值本身）",
        "source": os.path.basename(x),
        "source_sha256": sha256(x),
        "method": "export_baseline_status.py：只讀 Siliconomics Interface 的 SIF_ 名稱（Coord 頁推導）、Wattonomics 快照 WIF_Recon_NV_*、協調橋接檔 Bridge 頁 g2c／g2p；判斷類欄位讀 builder/seed/coord_meta_v*.json；本程式不做判定或計算（規則 v1.4.3 C3）",
        "generated": datetime.date.today().isoformat(),
        "wattonomics_version": wt_ver,
        "wattonomics_source_sha256": sha256(os.path.join(root, "wattonomics", "model", wt_file)),
        "bridge": ({"file": os.path.basename(bridge_path), "sha256": sha256(bridge_path), "in_title": bridge_title,
                    "note": "timing_cum（GW）＝Bridge g2c、timing_cum_ratio＝g2p；橋接檔於建置當時讀「前一版」Siliconomics 基準（見 in_title），其輸出再餵進本版 In 頁 crd_x2，故來源標示版本落後一版屬設計"}
                   if bridge_path else None),
        "threshold": v("SIF_crd_thr"),
        "wattonomics_coverage_complete": bool(v("SIF_crd_cov_done")),
        "years": years,
        "c0_actuals": c0a,
        "level_coefficients": coefs,
        "hardware_companies": meta.get("hardware_companies", []),
        "depends_on_years": meta.get("depends_on_years", {}),
    }
    json.dump(st, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"狀態檔：{out}（{os.path.basename(x)}）")


if __name__ == "__main__":
    main()
