# unit-basis 交叉檢驗算式（只用 FY2027 Q2 以前資料；Q3 實績不入）
# 單位：顆數＝百萬顆；金額＝十億美元；每顆價＝千美元
S26 = dict(bw=5.4, rb=1.7, hp=0.5)                       # Siliconomics v0.20 Units 2026
S27 = dict(bw=2.256909, rb=6.582652, ru=0.282114)        # Siliconomics v0.20 Units 2027（截頂後）
S27raw = dict(bw=2.4, rb=7.0, ru=0.3)
# 每封裝 NVIDIA 運算營收（千美元）：NVIDIA 個股模型 In 頁 B 節（chip_bridge v0.2.3）
# 模組＋CPU＋LPDDR＋匣級加價×k_tray(0/0.5/1)；Blackwell 低端再放寬到 35（HGX B200 較便宜，Assumed）
P = dict(
  bw=(35.0, 40.4+2+3.069444+0.5*3.072222, 47+2+3.069444+3.072222),
  rb=(50+2.5+5.055556, 55+2.5+5.055556+0.5*19.222222, 60+2.5+5.055556+19.222222),
  ru=(60+2.5+5.055556, 75+2.5+5.055556+0.5*19.222222, 100+2.5+5.055556+19.222222),
  hp=(25.0, 30.0, 35.0))
def rev(units, half=False):
    k = 0.5 if half else 1.0
    return [sum(units[p]*k*P[p][i] for p in units) for i in range(3)]
print("每封裝價（低/基準/高）", {k:[round(x,1) for x in v] for k,v in P.items()})
# ---- 2025（FY2026）：NVIDIA 自己的 6M Blackwell 是什麼單位
dc26fy = 39.1+41.1+51.2+62.3
for ns in (0.15,0.20):
    print("FY2026 DC", round(dc26fy,1), "網通占比", ns, "→ 運算", round(dc26fy*(1-ns),1))
for k,lab in ((1,"6M 當封裝"),(0.5,"6M 當晶粒→3M 封裝")):
    print(lab, [round(6*k*x,0) for x in (35,40.4+2+3.07,55)])
print("$500B/20M＝每顆", 500/20, "千美元；×2＝每封裝", 1000/20)
# ---- 2026
h1_dc = 75.2+89.0; h1_net = 14.8+17.5; h1_cmp = h1_dc-h1_net
q2_cmp = 89.0-17.5
lo = h1_cmp + 2*q2_cmp
hi = h1_cmp + q2_cmp*1.15 + q2_cmp*1.15**2
print("H1 DC", h1_dc, "H1 運算", round(h1_cmp,1), "CY2026 運算區間", round(lo,1), round(hi,1))
print("2026 封裝口徑隱含運算營收", [round(x) for x in rev(S26)])
print("2026 全部當晶粒（÷2）", [round(x) for x in rev(S26, True)])
mix = dict(S26); mix['rb'] = S26['rb']/2
print("2026 只有 Rubin 當晶粒", [round(x) for x in rev(mix)])
# Blackwell 單獨：公司運算 − Rubin(0–60) − Hopper(12.5–17.5)，再除以 Blackwell 每封裝價
bw_need_lo = (lo-60-17.5)/P['bw'][2]; bw_need_hi = (hi-0-12.5)/P['bw'][0]
print("2026 Blackwell 需要的封裝數區間", round(bw_need_lo,2), round(bw_need_hi,2), "（Siliconomics 5.4；晶粒口徑 2.7）")
# CoWoS（2026）：36Kr 66 萬片
for u,lab in ((7.6,"封裝"),(3.8,"晶粒→封裝")):
    print("2026 每片封裝數", lab, round(u/0.66,1))
# ---- 2027
fy27_rev = 81.615+96.221+108+118.8   # NVIDIA 模型 FY2027 錨點（含 Q3 指引中點與 Q4 假設；非 Q3 實績）
dc28 = 633.7                         # Siliconomics v0.16.1：FY2028 +70% 指引推得
c28 = (dc28*0.80, dc28*0.85)
print("FY2028 DC", dc28, "運算", [round(x) for x in c28])
print("2027 封裝口徑", [round(x) for x in rev(S27)])
print("2027 晶粒口徑（÷2）", [round(x) for x in rev(S27, True)])
# 晶粒口徑下 Rubin 每封裝需要多少營收才對得上
need = (c28[0] - S27['bw']/2*P['bw'][2] - S27['ru']/2*P['ru'][2])/(S27['rb']/2)
print("晶粒口徑下 Rubin 每封裝需要營收（千美元，取最寬鬆）", round(need,1), "；VR200 機架價每封裝 低/基準/高", [round(x/72*1000,1) for x in (7.8,8.4,9.1)])
# 公司營收 → 需要的 Rubin 封裝數
for i,lab in enumerate(("低價","基準價","高價")):
    for c in c28:
        r = (c - S27['bw']*P['bw'][i] - S27['ru']*P['ru'][i])/P['rb'][i]
        print("2027 指引隱含 Rubin 封裝", lab, round(c), "→", round(r,2))
print("MS 9 萬櫃 × 72 ＝", 90*72/1000, "百萬封裝；7M 若為晶粒＝", 3.5, "百萬封裝＝", round(3.5e6/72/1000,1), "千櫃")
print("MS NVIDIA CoWoS 1.222M 片：每片封裝（封裝口徑）", round(9.7/1.222,1), "（晶粒口徑）", round(4.85/1.222,1))
print("Siliconomics 自算 2027 NVIDIA CoWoS 片數", 161.208+822.832+70.528, "千片；晶粒口徑減半", round((161.208+822.832+70.528)/2,1))
# ---- Wattonomics 438 萬
w_rb, w_ru = 14.41*291744/1e6, 0.93*194760/1e6
print("WAT 2027 Rubin", round(w_rb,3), "RU", round(w_ru,3), "合計", round(w_rb+w_ru,3), "SIL", round(sum(S27.values()),3), "比", round(sum(S27.values())/(w_rb+w_ru),2))
wr = [w_rb*P['rb'][i]+w_ru*P['ru'][i] for i in range(3)]
print("WAT 隱含 NVIDIA 運算營收", [round(x) for x in wr], "→ DC（網通 15–20%）", round(wr[0]/0.85), round(wr[2]/0.80), "基準", round(wr[1]/0.825))
print("每 Rubin 封裝 IT capex（TK）", round(37.586352e9/291744/1e3,1), "千美元")
