# 三層基準協調狀態檔匯出（v0.19-trial；規則 docs/plan/三層基準協調規則.md v1.1；欄位依 tools/baseline_coordination/README.md）
# 用法：python3 siliconomics/tools/export_baseline_status.py [<重算後 Siliconomics xlsx>] [<輸出 json>]
#   預設讀 model/CURRENT、寫 data/baseline_status.json。只讀 Interface 的 SIF_ 名稱（Coord 頁推導），本程式不做任何判定或計算。
#   對帳單與期限（recon_doc、deadline）、硬體個股清單與依賴年份讀 builder/seed/coord_meta_v*.json（判斷類，未定者留空）。
import glob, json, os, sys
import openpyxl

SIL = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
YRS = [2026, 2027, 2028, 2029]


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
    years = {}
    coefs = []
    for y in YRS:
        yy = y % 100
        st = v(f"SIF_RecoStatus_{y}")
        lead = v(f"SIF_RecoLead_{y}")
        rec = {
            "wattonomics": {"evidence_share": v(f"SIF_crd_ew_{yy}"), "prev_has_evidence": prev(f"SIF_crd_prevw_{yy}")},
            "siliconomics": {"evidence_share": v(f"SIF_crd_es_{yy}"), "prev_has_evidence": prev(f"SIF_crd_prevs_{yy}")},
            "status": st,
            "lead": ("wattonomics" if lead == "Wattonomics" else "siliconomics") if st == "乙" else None,
            "label": v(f"SIF_RecoLabel_{y}"),
            "converted": v(f"SIF_DeployGW_Demand_{y}"),
            "supply_side": v(f"SIF_DeployGW_Supply_{y}"),
            "current_s33": v(f"SIF_DeployGW_Current_{y}"),
            "use_converted": bool(v(f"SIF_RecoUseConv_{y}")),
            "reconciling": bool(v(f"SIF_RecoFlag_{y}")),
            "recon_doc": meta.get("recon_doc", {}).get(str(y)),
            "deadline": meta.get("deadline", {}).get(str(y)),
            "supply_cap_check_only": "供給上限非獨立（SIF_CapOnlineGW_Base＝本庫自身推得），依規則 v1.1 只作檢查、不封頂；故不填 supply_cap",
        }
        # 基準部署 GW（Interface 向量 SIF_DeployGW_Base 欄 E–H＝2026–2029）
        sh, ref = list(wb.defined_names["SIF_DeployGW_Base"].destinations)[0]
        row = [c.value for c in wb[sh][ref.replace("$", "")][0]]
        rec["sil_base"] = row[YRS.index(y)]
        years[str(y)] = rec
        if not rec["use_converted"] and y >= 2028:
            coefs.append({"name": f"um{yy}_B（S33；經 crd_k_{yy} 帶入 {y} 基準）", "year": y, "tag": "Assumed",
                          "note": "覆蓋未補齊、該年維持現行（對帳中）；改用換算值後不再在基準鏈上"})
    st = {
        "_note": "v0.19-trial 試作（不合併）：由 siliconomics/tools/export_baseline_status.py 自 " + os.path.basename(x) + " 匯出；換算值與 Wattonomics 證據比例為暫值（待 Wattonomics v0.3.2）",
        "source": os.path.basename(x),
        "threshold": v("SIF_crd_thr"),
        "wattonomics_coverage_complete": bool(v("SIF_crd_cov_done")),
        "years": years,
        "level_coefficients": coefs,
        "hardware_companies": meta.get("hardware_companies", []),
        "depends_on_years": meta.get("depends_on_years", {}),
    }
    json.dump(st, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"狀態檔：{out}（{os.path.basename(x)}）")


if __name__ == "__main__":
    main()
