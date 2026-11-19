# Wattonomics 改讀 Siliconomics v0.21 的影響分析：把試算結果整理成活頁簿（值；不含公式，數字全部取自重算後的試算檔與快照）。
# 用法：python3 build_xlsx.py <舊 Wattonomics 快照> <試算 Wattonomics 快照> <舊 Siliconomics 快照> <試算 Siliconomics 快照> <舊橋接 xlsx> <試算橋接 xlsx> <stocks json> <輸出 xlsx>
# 試算步驟（都在暫存區，未入模型）：
#   1) python3 wattonomics/builder/build.py --base wattonomics/model/20261118_Wattonomics_v0.3.7.xlsx --version v0.3.8 --date 20261119 --out <raw>；LibreOffice 重算；export_snapshot.py
#   2) wattonomics/model/CURRENT 指向試算檔後 python3 siliconomics/tools/baseline_bridge.py <raw>；LibreOffice 重算 → 新換算值 x2
#   3) siliconomics/tools/apply_inputs.py 套 sil_in_edits_trial.json（crd_x2、crd_ew）→ builder --version v0.22-trial；重算；export_snapshot.py --allow-trial-sgov
#   4) 再跑一次 1)、2) 確認收斂（Wattonomics 0 個名稱差異、x2 不變）
#   5) bash siliconomics/tools/cascade.sh 兩輪（第二輪除第一輪建置失敗的 nvidia、jentech 外全部 SAME）
import json, sys, re, openpyxl
from openpyxl.styles import Font
wo, wn, so, sn = [json.load(open(p, encoding="utf-8"))["names"] for p in sys.argv[1:5]]
bo, bn, stocks, OUT = sys.argv[5], sys.argv[6], json.load(open(sys.argv[7], encoding="utf-8")), sys.argv[8]
B = Font(bold=True); wb = openpyxl.Workbook()
def sheet(title, head, first=False):
    ws = wb.active if first else wb.create_sheet(title); ws.title = title; ws.append(head)
    for c in ws[1]: c.font = B
    return ws
ws = sheet("Wattonomics重點", ["名稱", "年", "v0.3.7", "v0.3.8 試算", "差", "差％"], True)
for p in ["WIF_MktGW_Add", "WIF_MktGW_YE", "WIF_Evidence_Lo", "WIF_Evidence", "WIF_Evidence_Hi", "WIF_Coverage", "WIF_Recon_NV_NVGW", "WIF_Recon_NV_ContentPerGW",
          "WIF_Recon_NV_ImpliedRev", "WIF_Recon_NV_ActualRev", "WIF_Recon_NV_Ratio", "WIF_Spend_total", "WIF_Gap_Short_A_D0", "WIF_Gap_CumUnarranged_A_D0"]:
    for y in range(2024, 2033):
        k = f"{p}_{y}"
        if k in wo and isinstance(wo[k], (int, float)):
            ws.append([p, y, wo[k], wn[k], wn[k] - wo[k], (wn[k] / wo[k] - 1) if wo[k] else None])
ws = sheet("Wattonomics全部變動", ["名稱", "v0.3.7", "v0.3.8 試算", "差％"])
for k in sorted(wo):
    if k in wn and wo[k] != wn[k] and k.startswith("WIF_"):
        r = (wn[k] / wo[k] - 1) if isinstance(wo[k], (int, float)) and isinstance(wn[k], (int, float)) and wo[k] else None
        ws.append([k, wo[k], wn[k], r])
ws = sheet("橋接換算值", ["列", "項目", "版本", 2026, 2027, 2028, 2029])
for tag, p in (("現行（Siliconomics v0.21 ＋ Wattonomics v0.3.7；r2 橋接檔）", bo), ("試算（v0.21 ＋ v0.3.8）", bn)):
    for r in openpyxl.load_workbook(p, data_only=True)["Bridge"].iter_rows(values_only=True):
        if r[0] in ("w1", "s2", "x2", "s0", "g", "g2c", "g2p", "g3"): ws.append([r[0], str(r[1])[:60], tag] + list(r[4:8]))
ws.append(["crd_x2", "模型現值（Siliconomics v0.21 In 頁，仍是對 v0.20 算的）", "模型", 20.972, 24.488, 28.184, 28.482])
ws = sheet("Siliconomics全部變動", ["名稱", "v0.21", "v0.22 試算", "差％"])
flat = lambda v: v if not isinstance(v, list) else json.dumps(v, ensure_ascii=False)
for k in sorted(so):
    if k in sn and so[k] != sn[k]:
        a, b = so[k], sn[k]
        if isinstance(a, list) and isinstance(b, list) and len(a) == len(b):
            for i, (x, y) in enumerate(zip(a, b)):
                if x != y: ws.append([f"{k}[{i}]", x, y, (y / x - 1) if isinstance(x, (int, float)) and isinstance(y, (int, float)) and x else None])
        else:
            ws.append([k, flat(a), flat(b), (b / a - 1) if isinstance(a, (int, float)) and isinstance(b, (int, float)) and a else None])
ws = sheet("個股", ["資料夾", "公司", "次產業", "加權目標價 改前", "改後", "變動％", "上檔 改前", "上檔 改後", "評等 改前", "評等 改後", "評等是否改變"])
for r in sorted(stocks, key=lambda r: -(r["tp1"] / r["tp0"] - 1)):
    ws.append([r["folder"], r["name"], r["seg"], r["tp0"], r["tp1"], r["tp1"] / r["tp0"] - 1, r["up0"], r["up1"], r["r0"], r["r1"], "是" if r["r0"] != r["r1"] else ""])
for ws in wb:
    ws.column_dimensions["A"].width = 44; ws.column_dimensions["B"].width = 30
wb.save(OUT); print("wrote", OUT)
