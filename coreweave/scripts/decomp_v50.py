"""CoreWeave 目標價變動拆解 v4.13（25.51）→ v5.0（工作單第 4 節；CLAUDE.md 第 12 條）。
新格式沒有 Q2 資料版本，① 方法改為新格式與 ② Q3 實績與日曆滾動無法分開（Q3 實績是 v5.0 的起點資料）：合併為一段「①＋②」＝v5.0 在 TK v18.0、現價 43.66 下的目標價 − v4.13；
③ Tokenomics v18.0 → v22.0（現價 43.66）；④ 現價 43.66 → 60.77（TK v22.0）。四段相加＝總變動。→ data/decomp_v50.json"""
import datetime, json, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import runner_v50 as RN

V413 = 25.51
res = RN.run([("base", {}, {}), ("px", {"px": 43.66}, {}), ("tk", {"px": 43.66}, {"CRWV_TKV": "v18.0"})], 2)
a, b, c = res["base"]["tp"], res["px"]["tp"], res["tk"]["tp"]
steps = [{"id": "m", "label": "①＋② 方法改為 Oracle 收支型樣板（DCF 50%＋Comparable 50%、FCFF、新雲同業倍數、評等五級）並滾入 Q3 實績（TK v18.0、現價 43.66 下）", "delta": c - V413, "tp": c},
         {"id": "tk", "label": "③ Tokenomics v18.0 → v22.0（現價 43.66）", "delta": b - c, "tp": b},
         {"id": "px", "label": "④ 現價 43.66（2026-11-05）→ 60.77（2026-11-11）", "delta": a - b, "tp": a}]
out = {"date": datetime.date.today().isoformat(), "v413": V413, "v50": a, "total": a - V413, "steps": steps, "check": sum(s["delta"] for s in steps) - (a - V413),
       "note": "①與②無法在 v5.0 內單獨切換（工作單第 4 節的替代做法）"}
json.dump(out, open(os.path.join(RN.HERE, "data", "decomp_v50.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps(out, ensure_ascii=False, indent=1))
