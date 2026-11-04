# Siliconomics Recon 快照：由下往上對帳（覆蓋個股基準情境 → Siliconomics 市場；R4 部署 GW → Wattonomics）。
# 用法：python3 siliconomics/tools/recon.py <R4 橋接重算後 xlsx> <輸出 recon_snapshot.json>
#   讀各公司 dist 成品（data_only）與 r4_bridge 的重算結果，寫成版本固定的快照；builder 以快照建 Recon 頁（常數＋公式），
#   所以下游個股之後改版不會讓 Siliconomics 重建結果漂移（CI 比對穩定）。每次 Siliconomics 升版重跑一次（README 建置步驟）。
# 原則：本工具只讀取、不計算判斷；狀態（一致／時差／待診斷）由 Recon 頁公式依容差與證據送件欄判定。
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


def main():
    r4x, out = sys.argv[1:3]
    snap = {"note": "Recon 快照（tools/recon.py 產生）。個股值取各公司 dist 成品的情境 " + BASE_SCEN + "（基準）；R4 取 r4_bridge 重算結果。",
            "base_scen": BASE_SCEN, "markets": {}}
    # 1. SoC 測試機：Advantest（會計年度 4–3 月，近似日曆年）＋泰瑞達（日曆年）份額，FY2025–FY2029
    adv, adv_ref = scen_row(dist("advantest"), "PL", "Advantest SoC 市占")
    ter, ter_ref = scen_row(dist("ter"), "PL", "泰瑞達 SoC 市占")
    snap["markets"]["soc"] = {
        "years": [2025, 2026, 2027, 2028, 2029],
        "rows": [
            {"who": "Advantest", "vals": adv, "ref": adv_ref,
             "note": "會計年度（4–3 月）份額，近似日曆年；FY2025＝實績隱含，FY2026 起＝CY2026 三季實績＋第四季指引推得"},
            {"who": "泰瑞達", "vals": ter, "ref": ter_ref, "note": "日曆年；FY2025、FY2026＝實績／錨點隱含，之後維持"},
        ],
    }
    # 2. R4：Siliconomics 推得淨新增上線 GW vs Wattonomics 主路徑（r4_bridge Bridge 頁）
    snap["markets"]["r4"] = {
        "years": [2026, 2027, 2028, 2029],
        "sil_online": bridge_row(r4x, "b8"),
        "wt_main": bridge_row(r4x, "b9"),
        "ref": os.path.basename(r4x) + " Bridge 第 b8、b9 列",
    }
    # 3. 季度實績驗證（v0.13）：NVIDIA 資料中心營收對帳點（r4_bridge n3、n4）
    try:
        snap["markets"]["nv_dc"] = {"years": [2026, 2027, 2028, 2029], "implied": bridge_row(r4x, "n3"), "company": bridge_row(r4x, "n4"),
                                    "ref": os.path.basename(r4x) + " Bridge 第 n3、n4 列"}
    except KeyError:
        pass
    json.dump(snap, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("寫出", out)
    for k, m in snap["markets"].items():
        print(k, json.dumps(m, ensure_ascii=False)[:300])


if __name__ == "__main__":
    main()
