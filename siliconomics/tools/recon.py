# Siliconomics Recon 快照：由下往上對帳（覆蓋個股基準情境 → Siliconomics 市場；R4 部署 GW → Wattonomics；公司實績對帳點）。
# 用法：python3 siliconomics/tools/recon.py <R4 橋接重算後 xlsx> <輸出 recon_snapshot.json> [<重算後 Siliconomics xlsx，預設 model/CURRENT>]
#   讀各公司 dist 成品（data_only）、r4_bridge 的重算結果與 Siliconomics 現行 Excel（Recon D 節、Coord 狀態），寫成版本固定的快照；
#   builder 以快照建 Recon 頁 A、B、F 節（常數＋公式），所以下游個股之後改版不會讓 Siliconomics 重建結果漂移（CI 比對穩定）。
#   每次 Siliconomics 升版重跑一次（README 建置步驟）。
# 快照結構（v0.20 起；tools/baseline_coordination/check.py 第 [9] 項讀）：每個市場有
#   kind：market（由下往上市場：rows＝已建模個股＋「其他（未建模）」殘差列，market＝市場值；C0 恆等式 Σ rows＝market）／
#         structural（結構一致，份額由本庫擁有，恆等式由 Checks K6 約束）／layer（層間對帳點，R4）／actual（公司實績對帳點，NVIDIA）；
#   status：一致／時差／時差邊界／對帳中／待診斷／未建／結構一致／C3（layer 讀 Coord 頁）；status_by_year：逐年狀態。
#   市場層狀態（soc）依 Recon 頁同一判定式（|比值 − 1| ≤ rc_tol）重算；D1 與 layer、actual 的狀態讀重算後 Siliconomics 的 Recon／Coord 頁，不自判。
# 個股基準情境：工作方式 v1.7「Siliconomics 對帳只對個股的基準情境加總」；基準取情境 B（平平）為預設，待 Andy 確認（決策 S33）。
import glob, json, os, sys
import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
BASE_SCEN = "B"


def dist(co):
    fs = glob.glob(os.path.join(ROOT, co, "dist", "*.xlsx"))
    assert len(fs) == 1, (co, fs)
    return fs[0]


def scen_row(path, sheet, label, scen=BASE_SCEN, cols="CDEFG"):
    """在 sheet 找『情境 X：…』或『情境 X（…）』引擎段落內第一個 label 列，回傳 (值串, 儲存格參照)。"""
    ws = openpyxl.load_workbook(path, data_only=True)[sheet]
    start = None
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if not isinstance(v, str):
            continue
        if start is None and v.startswith(f"情境 {scen}") and len(v) > 4 and v[4] in "：（(":
            start = r
            continue
        if start is not None and v.startswith("情境 ") and len(v) > 4 and v[4] in "：（(":
            break
        if start is not None and v == label:
            vals = [ws[f"{c}{r}"].value for c in cols]
            assert all(isinstance(x, (int, float)) for x in vals), (path, label, vals)
            return vals, f"{os.path.relpath(path, ROOT)} {sheet}!{cols[0]}{r}:{cols[-1]}{r}"
    raise KeyError((path, sheet, label, scen))


def bridge_row(path, key):
    ws = openpyxl.load_workbook(path, data_only=True)["Bridge"]
    for row in ws.iter_rows(min_row=3):
        if row[0].value == key:
            return [row[j].value for j in range(4, 8)]  # 2026–2029（欄 E–H）
    raise KeyError(key)


def sil_name(wb, nm):
    sh, ref = list(wb.defined_names[nm].destinations)[0]
    rng = wb[sh][ref.replace("$", "")]
    return [c.value for row in rng for c in row] if isinstance(rng, tuple) else rng.value


def recon_rows(wb):
    """Siliconomics Recon 頁 A 欄標籤 → 列號（D1 區塊）。"""
    ws = wb["Recon"]
    return {ws.cell(r, 1).value: r for r in range(1, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, str)}


ORDER = ["待診斷", "對帳中", "時差", "時差邊界", "一致"]


def worst(sts):
    """整體狀態＝逐年最差者（待診斷 > 對帳中 > 時差 > 時差邊界 > 一致）；其他字串（層級未建、無實績）不計。"""
    xs = [x for x in sts if x in ORDER]
    return min(xs, key=ORDER.index) if xs else None


def main():
    r4x, out = sys.argv[1:3]
    silx = sys.argv[3] if len(sys.argv) > 3 else os.path.join(ROOT, "siliconomics", "model", open(os.path.join(ROOT, "siliconomics", "model", "CURRENT")).read().strip())
    S = openpyxl.load_workbook(silx, data_only=True)
    tol = sil_name(S, "rc_tol")
    oth = sil_name(S, "rc_soc_oth")
    snap = {"note": "Recon 快照（tools/recon.py 產生）。個股值取各公司 dist 成品的情境 " + BASE_SCEN + "（基準）；R4 取 r4_bridge 重算結果；D1、C3 狀態取 "
            + os.path.basename(silx) + "。結構見 tools/recon.py 檔頭與 tools/baseline_coordination/README.md。",
            "base_scen": BASE_SCEN, "siliconomics": os.path.basename(silx), "tolerance": tol, "markets": {}}
    # 1. SoC 測試機：Advantest（會計年度 4–3 月，近似日曆年）＋泰瑞達（日曆年）份額，FY2025–FY2029；市場值＝100%（份額口徑）
    adv, adv_ref = scen_row(dist("advantest"), "PL", "Advantest SoC 市占")
    ter, ter_ref = scen_row(dist("ter"), "PL", "泰瑞達 SoC 市占")
    msum = [a + b for a, b in zip(adv, ter)]
    st_y = ["一致" if abs(m_ / (1 - oth) - 1) <= tol else "待診斷" for m_ in msum]   # 證據送件（時差）目前無；Recon 頁同式
    snap["markets"]["soc"] = {
        "kind": "market", "name": "SoC 測試機（份額口徑）", "unit": "份額", "status": worst(st_y), "status_by_year": st_y,
        "years": [2025, 2026, 2027, 2028, 2029],
        "market": [1.0] * 5,
        "modeled_sum": msum,
        "expected_residual": oth,
        "rows": [
            {"who": "Advantest", "vals": adv, "ref": adv_ref,
             "note": "會計年度（4–3 月）份額，近似日曆年；FY2025＝實績隱含，FY2026 起＝CY2026 三季實績＋第四季指引推得"},
            {"who": "泰瑞達", "vals": ter, "ref": ter_ref, "note": "日曆年；FY2025、FY2026＝實績／錨點隱含，之後維持"},
            {"who": "其他（未建模）", "vals": [1 - m_ for m_ in msum], "residual": True,
             "note": f"殘差＝1 − 已建模份額合計；預期約 rc_soc_oth＝{oth}（In 頁 M27）"},
        ],
    }
    # 2. R4：Siliconomics 推得淨新增上線 GW vs Wattonomics 主路徑（r4_bridge Bridge 頁）；狀態＝C3（讀 Coord，規則 v1.4.3 唯一權威）
    snap["markets"]["r4"] = {
        "kind": "layer", "name": "R4 部署 GW（層間對帳點，非個股加總）",
        "years": [2026, 2027, 2028, 2029],
        "sil_online": bridge_row(r4x, "b8"),
        "wt_main": bridge_row(r4x, "b9"),
        "status": "C3", "status_by_year": [f"{sil_name(S, f'SIF_RecoStatus_{y}')}：{sil_name(S, f'SIF_RecoLabel_{y}')}" for y in (2026, 2027, 2028, 2029)],
        "ref": os.path.basename(r4x) + " Bridge 第 b8、b9 列",
    }
    # 3. 公司實績對帳點：NVIDIA 資料中心營收（r4_bridge n3、n4；v0.20 起主讀數＝純實績 n10–n14，規則 v1.4.3 C0）
    try:
        nv = {"kind": "actual", "name": "NVIDIA 資料中心營收（公司實績對帳點）", "years": [2026, 2027, 2028, 2029],
              "implied": bridge_row(r4x, "n3"), "company": bridge_row(r4x, "n4"),
              "ref": os.path.basename(r4x) + " Bridge 第 n3、n4 列（company＝含指引，只列參考）"}
        act = bridge_row(r4x, "n10")
        nv.update({"actual_cum": act, "actual_label": "FY27 Q1＋Q2", "months": [6, None, None, None],
                   "actual_q_annual": bridge_row(r4x, "n13"),
                   "ratio_main": bridge_row(r4x, "n12"), "ratio_second": bridge_row(r4x, "n14"),
                   "ref_actual": os.path.basename(r4x) + " Bridge 第 n10–n14 列（NVIDIA 8-K FY27 Q1 75.2、Q2 89.0）"})
        sy = []
        for m_, s_ in zip(nv["ratio_main"], nv["ratio_second"]):
            if m_ is None:
                sy.append("無實績")
            elif abs(m_ - 1) <= tol and abs(s_ - 1) <= tol:
                sy.append("一致")
            elif abs(m_ - 1) > tol and abs(s_ - 1) > tol and (m_ > 1) == (s_ > 1):
                sy.append("對帳中")
            else:
                sy.append("時差邊界")
        nv["status_by_year"] = sy; nv["status"] = worst(sy)
        snap["markets"]["nv_dc"] = nv
    except KeyError:
        pass
    # 4. Recon D 節（v0.20）：D1 GPU／ASIC 晶片金額讀 Siliconomics Recon 頁；其餘未建
    rr = recon_rows(S)
    ws = S["Recon"]
    lab = lambda k: next(x for x in rr if x.startswith(k))
    vals = lambda k: [ws.cell(rr[lab(k)], c).value for c in range(4, 8)]
    if any(x.startswith("D1. GPU") for x in rr):
        rows = []
        for k, who in (("已建模：NVIDIA", "NVIDIA"), ("已建模：AMD", "AMD"), ("已建模：Broadcom", "Broadcom（ASIC）"), ("已建模：Marvell", "Marvell（ASIC）"),
                       ("已建模：聯發科", "聯發科（ASIC）"), ("已建模：世芯", "世芯（ASIC）"), ("已建模：創意", "創意（無層級份額）")):
            rows.append({"who": who, "vals": vals(k), "ref": f"{os.path.basename(silx)} Recon!D{rr[lab(k)]}:G{rr[lab(k)]}"})
        rows.append({"who": "其他（未建模）", "vals": vals("殘差（其他"), "residual": True, "ref": f"{os.path.basename(silx)} Recon!D{rr[lab('殘差（其他')]}:G{rr[lab('殘差（其他')]}",
                     "note": "＝殘差①（非建模夥伴份額）＋殘差②（夥伴未收；可為負）"})
        sy = vals("狀態（市場層")
        snap["markets"]["gpu_asic"] = {
            "kind": "market", "name": "GPU／ASIC 晶片金額（US$B；層級值，2026 以公司實績對帳）", "unit": "US$B", "years": [2026, 2027, 2028, 2029],
            "market": vals("市場值（層級）"), "modeled_sum": vals("已建模個股加總"), "rows": rows,
            "ratio_main": vals("比值（主讀數"), "ratio_second": vals("比值（第二讀數"),
            "status_by_year": sy, "status": worst(sy) or "無實績",
            "note": "已建模個股列為層級值（份額 × 市場）；比較基準（公司實績）見 Siliconomics Recon 頁 D1 實績列與 builder/seed/recon_d_gpu_asic_v*.json"}
    else:
        snap["markets"]["gpu_asic"] = {"kind": "market", "name": "GPU／ASIC 晶片金額", "status": "未建"}
    for key, nm in (("cowos", "CoWoS（台積電；Amkor、日月光委外）"), ("wfe", "晶圓前段設備（應用材料、Lam、KLA、ASML；TEL 未建模）"),
                    ("pcb_ccl", "AI 高速 PCB／CCL（勝宏、金像電、台光電）"), ("optics", "光學模組總量（中際旭創、新易盛、Lumentum、Coherent）")):
        snap["markets"][key] = {"kind": "market", "name": nm, "status": "未建"}
    # 5. 結構一致的市場（Recon C 節；份額由本庫擁有，恆等式由 Checks K6 約束）
    for key, nm, k6 in (("asic_ds", "ASIC 設計服務（TPU、Trainium、其他 ASIC）", "K6-tpu／oth／trn"), ("cpo", "CPO 光學", "K6-cpo"), ("abf", "加速器 ABF 載板", "K6-abfnv／abfot"),
                        ("power", "AI 電源", "K6-pwr"), ("liquid", "AI 液冷", "K6-lc"), ("asm", "AI 伺服器組裝", "K6-asm"), ("hbm", "HBM", "K6-hbm")):
        snap["markets"][key] = {"kind": "structural", "name": nm, "status": "結構一致", "check": "Checks " + k6}
    json.dump(snap, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("寫出", out)
    for k, m in snap["markets"].items():
        print(k, m.get("kind"), m.get("status"), json.dumps(m.get("status_by_year"), ensure_ascii=False))


if __name__ == "__main__":
    main()
