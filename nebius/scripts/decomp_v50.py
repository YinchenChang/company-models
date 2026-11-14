"""Nebius 目標價變動拆解 v0.5.3（舊 JS 引擎，10.30）→ v0.6.0（CoreWeave 新雲樣板重建）。兩段相加＝總變動 → data/decomp_v50.json
  ① 方法與資料：新樣板（DCF 50%＋Comparable 50%、FCFF、新雲同業倍數、評等五級）＋Q3 實績＋Tokenomics v22.0（以舊現價 120.37 計）
  ② 現價 120.37 → 128.64（2026-11-12）
舊引擎（JS、company.json）與新樣板結構不同，無法逐項對應；① 不再細拆，README 列主要差異。"""
import datetime, json, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import runner_v50 as RN

V053, PX0 = 10.30, 120.37
res = RN.run([("s1", {"px": PX0}, {}), ("s2", {}, {})], 2)
tp1, tp2 = res["s1"]["tp"], res["s2"]["tp"]
steps = [{"id": "m", "label": "① 新樣板＋Q3 實績＋Tokenomics v22.0（以舊現價 120.37 計）", "delta": tp1 - V053, "tp": tp1},
         {"id": "px", "label": "② 現價 120.37 → 128.64（2026-11-12）", "delta": tp2 - tp1, "tp": tp2}]
out = {"date": datetime.date.today().isoformat(), "v053": V053, "v060": tp2, "total": tp2 - V053, "steps": steps,
       "check": sum(x["delta"] for x in steps) - (tp2 - V053), "full_at_old_px": {"tp": tp1, "rec": res["s1"]["rec"], "up": tp1 / PX0 - 1}}
json.dump(out, open(os.path.join(RN.HERE, "data", "decomp_v50.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps(out, ensure_ascii=False, indent=1))
