"""CoreWeave 目標價變動拆解 v4.13（25.51）→ v5.0（工作單第 4 節替代做法；CLAUDE.md 第 12 條）。五段相加＝總變動 → data/decomp_v50.json
  ① 方法：v4.13 → v5.0 新格式（以 Q2 起點、TK v18.0、現價 43.66、舊同業與 rf 計）
  ② Q3 實績：v5.0 內把 Q3 實績換回「Q2 起點模型值」（下方 Q2START，依 v5.0 同一套公式由 Q2 10-Q 推一季）再換回實績
  ③ Tokenomics v18.0 → v22.0　④ 現價 43.66 → 60.77　⑤ 同業 NTM 口徑與 rf 4.88% → 4.98%"""
import datetime, json, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import runner_v50 as RN

V413 = 25.51
# Q2 起點模型值：Q3 季末在役 1,675 MW（2026-06-30 的 1,500 到年底 1,850 線性）；每 MW 年收入 9.87、營運費用 1.912、租金 0.85（＝85% × 1.00）、管銷 5.51%（v5.0 同一套 Tokenomics 與參數）；
# 股票薪酬、利息沿用 Q2 實績；折舊＝Q2 1.393＋在建工程 11.9 ÷ 6 ÷ 4；資本支出＝Q3 新增 175 MW ×（1 − λ）× 30.6＋2027 年新增 1,250 × λ × 34.9 × 0.25＋自建機房；
# 資產負債：現金維持 Q2 5.524，缺口全由借款補（借款 Q2 帳面 35.068＋資本支出 − 營運現金流）；股數＝Q2 期末 0.5515。
mw = 1675; avg = (1500 + mw) / 2; bill = avg * 0.92
rev = bill * 9.87 * 0.25 / 1000
aeb = rev - avg * (1.912 + 0.85) * 0.25 / 1000 - rev * 0.0551
sbc, intq = 0.165, 0.640
da = 1.393 + 11.9 / 6 / 4
oi = aeb - sbc - da; ni = oi - intq; cfo = ni + da + sbc
cx = (175 * 0.65 * 30.6 + 1250 * 0.35 * 34.9 * 0.25 + 175 * 0.15 * 12.67) / 1000
Q2START = {"mw_q3": mw, "rev9m": 4.653 + rev, "oi9m": -0.193 + oi, "aeb9m": 2.667 + aeb, "aoi9m": 0.149 + aeb - da, "ni9m": -1.366 + ni, "int9m": 1.176 + intq,
           "da9m": 2.54 + da, "sbc9m": 0.318 + sbc, "cfo9m": 3.663 + cfo, "cx9m_cash": 14.117 + cx, "cx_q3_def": cx, "rev_q3": rev, "aeb_q3": aeb, "da_q3": da, "int_q3": intq,
           "cash_q3": 5.524, "rcash_q3": 1.38, "debt_q3": 35.068 + cx - cfo, "finl_q3": 0.221, "opl_q3": 16.319, "ppe_q3": 46.736 + cx - da, "sh_out": 0.5515}
OLD = {"pn_eb": 6.194, "rf": 0.0488}
P = {"px": 43.66}
runs = [("s0", {**Q2START, **P, **OLD}, {"CRWV_TKV": "v18.0"}), ("s1", {**P, **OLD}, {"CRWV_TKV": "v18.0"}), ("s2", {**P, **OLD}, {}), ("s3", dict(OLD), {}), ("s4", {}, {})]
res = RN.run(runs, 2)
tp = {k: res[k]["tp"] for k in res}
labs = [("m", "① 方法改為 Oracle 收支型樣板（DCF 50%＋Comparable 50%、FCFF、新雲同業倍數、評等五級；以 Q2 起點、TK v18.0、現價 43.66、舊同業與 rf 計）", tp["s0"] - V413, tp["s0"]),
        ("q3", "② Q3 實績取代 Q2 起點模型值（營收 3.661 對模型 " + f"{rev:.3f}" + "、在役 1,700 對 1,675 MW、借款與現金、股數）", tp["s1"] - tp["s0"], tp["s1"]),
        ("tk", "③ Tokenomics v18.0 → v22.0（本模型所用基準值不變；IF_RevGWFleet 系列 1–2% 變動只用於上限檢查）", tp["s2"] - tp["s1"], tp["s2"]),
        ("px", "④ 現價 43.66（2026-11-05）→ 60.77（2026-11-11）", tp["s3"] - tp["s2"], tp["s3"]),
        ("peer", "⑤ 同業 NTM 口徑（Nebius FY27 → NTM EBITDA）與 rf 4.88% → 4.98%", tp["s4"] - tp["s3"], tp["s4"])]
steps = [{"id": i, "label": l, "delta": d, "tp": t, "dcf": None} for i, l, d, t in labs]
# 評等變動原因：在舊價 43.66 下的 v5.0 目標價與評等
out = {"date": datetime.date.today().isoformat(), "v413": V413, "v50": tp["s4"], "total": tp["s4"] - V413, "steps": steps,
       "check": sum(s["delta"] for s in steps) - (tp["s4"] - V413), "q2start": Q2START,
       "rating_at_old_px": {"px": 43.66, "tp": tp["s3"] - 0 if False else res["s2"]["tp"], "rec": res["s2"]["rec"], "up": res["s2"]["tp"] / 43.66 - 1},
       "full_at_old_px": None}
res2 = RN.run([("oldpx", dict(P), {})], 1)
out["full_at_old_px"] = {"tp": res2["oldpx"]["tp"], "rec": res2["oldpx"]["rec"], "up": res2["oldpx"]["tp"] / 43.66 - 1}
json.dump(out, open(os.path.join(RN.HERE, "data", "decomp_v50.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps({k: v for k, v in out.items() if k != "q2start"}, ensure_ascii=False, indent=1))
