# CoWoS 路徑：以 NVIDIA CoWoS 用片量推算 2027 年 NVIDIA 加速器封裝數（只出分析，不改任何模型值）。
# 用法：python3 build.py <輸出 xlsx> [<結果 json>]   之後以 LibreOffice 重算（計算頁為公式；敏感度與成立條件頁為本程式算出的值）。
# 方法（相對法）：以 2025 年「營收推得的封裝數」校準片數→封裝的轉換係數 κ，再套到 2026（校驗）與 2027。
#   有效片數(y)＝( L×片數(y−1) ＋ (1−L)×片數(y) ) ×(1−非 GPU 占比(y))      L＝CoWoS 產出到 NVIDIA 認列營收的時差（年）
#   κ＝2025 年所需片數（封裝 ÷ 每片好封裝）÷ 有效片數(2025)
#   封裝(y)＝κ × 有效片數(y) ÷ Σ(世代占比 ÷ 每片好封裝)
import json, sys, math
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

OUT = sys.argv[1]; JOUT = sys.argv[2] if len(sys.argv) > 2 else None
IP, AN, AS, DE = "Interested-party", "Analogy", "Assumed", "Derived"
# 鍵: (說明, 低, 基準, 高, 標記, 來源)   低／高＝該輸入本身的低／高（不是「顆數最少／最多」）
INP = [
 ("w24", "NVIDIA CoWoS 用片量 2024（千片）", 190, 210, 230, AN, "S1 全球需求 372 千片 × NVIDIA 占比 51–62%（占比無 2024 年來源，取 2025–26 年占比區間）"),
 ("w25", "NVIDIA CoWoS 用片量 2025（千片）", 370, 410, 450, IP, "S1 全球 689 千片 × S5 約 59%＝407；S2 早期估計 450；低＝689×53.4%"),
 ("w26", "NVIDIA CoWoS 用片量 2026（千片）", 595, 780, 850, IP, "S3 780（2026-09 版）；S4 595（2025-07 版）；S6 800–850"),
 ("w27", "NVIDIA CoWoS 用片量 2027（千片）", 900, 1222, 1340, IP, "S3、S7 1,222；低＝台積電年底 17 萬片／月（S8 KGI）下的供給受限推算（年均 15 萬×12×45.6%＋委外 8 萬）；高＝S9 2,730×49%"),
 ("ng25", "NVIDIA 片數中非 GPU 用途占比 2025（Vera CPU、交換器晶片、車用）", 0.00, 0.02, 0.05, AS, "S4：委外 8 萬片含 Vera CPU 與車用；2025 年 Vera 未量產"),
 ("ng26", "同上 2026", 0.02, 0.05, 0.08, AS, "Vera CPU 2026 下半年起量"),
 ("ng27", "同上 2027", 0.03, 0.08, 0.15, AS, "每 NVL72 配 36 顆 Vera CPU：0.5 顆／GPU 封裝 ÷ 約 60 顆／片 ÷ GPU 每封裝 0.10 片 ≈ 8%（Vera 是否全用 CoWoS 未查到原文）"),
 ("lag", "CoWoS 產出到 NVIDIA 認列營收的時差（年）", 0.25, 0.375, 0.50, AS, "nv-units 報告：P1 台積電產出時點比認列早約 1–2 季"),
 ("pp_bw", "Blackwell 每片好封裝", 11.2, 13.2, 15.2, DE, "rubin-price xlsx 幾何頁（中介層面積 Analogy × 良率）"),
 ("pp_rb", "Rubin 每片好封裝", 7.2, 9.68, 12.35, DE, "同上；台積電稱 5 倍光罩已量產（S10）"),
 ("pp_ru", "Rubin Ultra 每片好封裝", 4.0, 4.4, 5.7, DE, "同上"),
 ("pp_hp", "Hopper 每片好封裝", 23.2, 29.0, 34.8, AN, "Siliconomics In cw_pp_HP"),
 ("u25_bw", "2025 Blackwell 封裝（百萬；營收推得）", 3.00, 3.42, 3.80, DE, "nv-units 報告：2025 年 B＋R 營收 1,708 億 ÷ 5.0 萬美元；低＝NVIDIA 稱 2025-10 已出 600 萬晶粒"),
 ("u25_hp", "2025 Hopper／H20 封裝（百萬）", 0.50, 0.70, 0.90, DE, "nv-units 報告 P1 錨點"),
 ("m26_bw", "2026 組合：Blackwell（百萬；只取比例）", None, 5.4, None, IP, "Siliconomics v0.21 In u_bw_26"),
 ("m26_rb", "2026 組合：Rubin", None, 1.7, None, IP, "u_rb_26"),
 ("m26_hp", "2026 組合：Hopper", None, 0.5, None, IP, "u_hp_26"),
 ("m27_bw", "2027 組合：Blackwell（只取比例）", None, 2.4, None, IP, "v0.21 決策 S55 修正後券商組合 2.4：6.7：0.3"),
 ("m27_rb", "2027 組合：Rubin", None, 6.7, None, IP, "同上"),
 ("m27_ru", "2027 組合：Rubin Ultra", None, 0.3, None, IP, "同上"),
 ("rev27", "曆年 2027 Blackwell＋Rubin 系統營收（十億美元）", 499.51, 576.003, 654.926, IP, "rubin-price xlsx A1 頁（FY2028 公司指引推得）"),
 ("p_bw", "Blackwell 每封裝 NVIDIA 營收（千美元）", 43, 50, 55, DE, "rubin-price xlsx 價格頁"),
 ("ru_x", "Rubin Ultra ÷ Rubin 每封裝營收", None, 1.996, None, DE, "Siliconomics v0.21 每 GW 內容 44.35÷33.28 × 每 GW 封裝 291,744÷194,760"),
]
SRC = [
 ("S1", "Morgan Stanley 全球 CoWoS 需求：2023 117、2024 372、2025 689、2026 1,394、2027 2,694 千片（2026-06 版）", "財聯社轉述（futunn，2026-06-25）", "https://news.futunn.com/en/post/75096132/morgan-stanley-raises-cowos-demand-forecast-to-reach-2-69", "二手"),
 ("S2", "Morgan Stanley：NVIDIA 2025 CoWoS 需求約 45 萬片", "GuruFocus 轉述（2025-03）", "https://www.gurufocus.com/news/2726257/nvidia-nvda-set-to-gain-market-share-amid-cowos-order-adjustments", "二手"),
 ("S3", "Morgan Stanley 2026-09：全球 1,394 → 2,509 千片；NVIDIA 780 → 1,222 千片", "X 貼文轉述（2026-10）", "https://x.com/SemiconductorsX/status/2106258073409221049", "二手，未見原文"),
 ("S4", "Morgan Stanley 2025-07：NVIDIA 2026 年 595 千片（占 60%），其中台積電 510、委外 80（含 Vera CPU 與車用）", "X 貼文／36Kr 轉述", "https://x.com/jukan05/status/1950102624164073553", "二手"),
 ("S5", "Morgan Stanley：NVIDIA 約占 CoWoS 59%", "htx 轉述（2026-05）", "https://www.htx.com/en-in/news/morgan-stanley-2026-semiconductor-report-buy-packaging-buy-t-lFQXxzmh/", "二手"),
 ("S6", "NVIDIA 已訂約 80–85 萬片", "INDmoney（2026-06）", "https://www.indmoney.com/blog/us-stocks/tsmc-cowos-bottleneck-ai-chip-supply-squeeze-explained", "二手"),
 ("S7", "Morgan Stanley：NVIDIA 2027 年 122.2 萬片（+57%），占比由 53.4% 降到 45.6%", "wccftech（2026-08）", "https://wccftech.com/nvidias-cowos-ai-chip-packaging-demand-to-be-foreshadowed-in-2027-by-amd-due-to-agentic-ai-cpu-ramp-says-morgan-stanley/", "二手"),
 ("S8", "台積電 CoWoS 月產能：2026 年底 11.5–14 萬片、2027 年約 17 萬片（法人估計）；另有 20 萬片說法", "TrendForce 轉述 TechNews（2026-04-16）、TrendForce（2026-10-27）", "https://www.trendforce.com/news/2026/04/16/news-tsmc-says-cowos-offers-industrys-largest-reticle-size-packaging-amid-intel-emib-rivalry-copos-advances/", "二手"),
 ("S9", "Goldman Sachs：台積電 CoWoS 2025 年 675、2026 年 1,275、2027 年 2,310–2,730 千片", "TrendForce（2026-10-27）、I/O Fund", "https://www.trendforce.com/news/2026/10/27/news-tsmc-reportedly-eyes-20-cowos-price-hike-skips-equipment-price-renegotiations-amid-ai-demand-surge/", "二手"),
 ("S10", "台積電 2026Q3 法說：5 倍光罩尺寸封裝已量產，開發 9 倍與 14 倍", "Zacks／MarketBeat 逐字稿（2026-10-15）", "https://www.marketbeat.com/earnings/reports/2026-10-15-taiwan-semiconductor-manufacturing-company-limited-stock/", "二手逐字稿"),
 ("S11", "台積電 2025Q1 法說：2025 年 CoWoS 產能倍增；2026 年「I cannot say the number」", "台積電官方逐字稿（2025-04）", "https://investor.tsmc.com/english/encrypt/files/encrypt_file/reports/2025-04/7630274eecc1197a4e3ea6a415f44a47204fe10a/TSMC%201Q25%20Transcript.pdf", "一手"),
 ("S12", "台積電 2026Q2 法說：封裝產能「so tight that now it's limiting my customers' growth」；台積電不揭露 CoWoS 產能", "Silicon Analysts 引述（2026-06／07）", "https://siliconanalysts.com/analysis/foundry-allocation-status-q1-2026", "二手"),
]
B = {r[0]: r[3] for r in INP}; LO = {r[0]: r[2] for r in INP}; HI = {r[0]: r[4] for r in INP}

def calc(p):
    eff = lambda y, a, b: (p["lag"] * p[a] + (1 - p["lag"]) * p[b]) * (1 - p[f"ng{y}"])
    e25, e26, e27 = eff(25, "w24", "w25"), eff(26, "w25", "w26"), eff(27, "w26", "w27")
    need25 = (p["u25_bw"] / p["pp_bw"] + p["u25_hp"] / p["pp_hp"]) * 1000
    k = need25 / e25
    s26 = p["m26_bw"] + p["m26_rb"] + p["m26_hp"]; s27 = p["m27_bw"] + p["m27_rb"] + p["m27_ru"]
    wpp26 = (p["m26_bw"] / p["pp_bw"] + p["m26_rb"] / p["pp_rb"] + p["m26_hp"] / p["pp_hp"]) / s26
    wpp27 = (p["m27_bw"] / p["pp_bw"] + p["m27_rb"] / p["pp_rb"] + p["m27_ru"] / p["pp_ru"]) / s27
    n26 = k * e26 / wpp26 / 1000; n27 = k * e27 / wpp27 / 1000
    avg = p["rev27"] / n27
    prb = (avg - p["m27_bw"] / s27 * p["p_bw"]) / (p["m27_rb"] / s27 + p["m27_ru"] / s27 * p["ru_x"])
    return dict(e25=e25, e26=e26, e27=e27, need25=need25, kappa=k, wpp26=wpp26, wpp27=wpp27, n26=n26, n26_exhp=n26 * (1 - p["m26_hp"] / s26), n27=n27, avg27=avg, p_rb=prb)

base = calc(B)
sens = []
for r in INP:
    k = r[0]
    if LO[k] is None: continue
    lo = calc({**B, k: LO[k]})["n27"]; hi = calc({**B, k: HI[k]})["n27"]
    if abs(lo - base["n27"]) < 1e-12 and abs(hi - base["n27"]) < 1e-12: continue   # 只影響 2026 校驗或隱含單價的輸入不列
    sens.append((k, r[1], LO[k], HI[k], lo, hi, lo / base["n27"] - 1, hi / base["n27"] - 1))
sens.sort(key=lambda s: -max(abs(s[6]), abs(s[7])))
# 合併區間：各輸入對數偏離的平方和開根號（輸入視為獨立；w25、w26、w27 來源相同，實際正相關，會互相抵銷一部分，故此區間偏寬）
dn = math.sqrt(sum(min(math.log(1 + s[6]), math.log(1 + s[7]), 0) ** 2 for s in sens))
up = math.sqrt(sum(max(math.log(1 + s[6]), math.log(1 + s[7]), 0) ** 2 for s in sens))
rng = (base["n27"] * math.exp(-dn), base["n27"] * math.exp(up))

def solve(key, target, a, b):
    f = lambda x: calc({**B, key: x})["n27"] - target
    fa, fb = f(a), f(b)
    if fa * fb > 0: return None
    for _ in range(80):
        m = (a + b) / 2; fm = f(m)
        if fa * fm <= 0: b, fb = m, fm
        else: a, fa = m, fm
    return (a + b) / 2
TARGETS = [("自建基準 613 萬", 6.1325), ("Siliconomics v0.20 券商值 912 萬", 9.12), ("需求側 438 萬", 4.38)]
must = []
for nm, t in TARGETS:
    must.append((nm, t, solve("w27", t, 100, 4000), solve("pp_rb", t, 2, 30), solve("ng27", t, -0.9, 0.9), t / base["n27"]))
paths = {"A1": (3.92, 5.844, 8.387), "D": (2.615, 4.758, 8.484), "P1": (4.004, 6.421, 11.033), "M": (2.476, 6.643, 16.457)}
med = lambda xs: (sorted(xs)[len(xs) // 2] if len(xs) % 2 else sum(sorted(xs)[len(xs) // 2 - 1:len(xs) // 2 + 1]) / 2)
med4 = med([v[1] for v in paths.values()]); med5 = med([v[1] for v in paths.values()] + [base["n27"]])
ov4 = (max(v[0] for v in paths.values()), min(v[2] for v in paths.values()))
ov5 = (max(ov4[0], rng[0]), min(ov4[1], rng[1]))
res = dict(base=base, range=rng, sens=[dict(key=s[0], lo_in=s[2], hi_in=s[3], n27_lo=s[4], n27_hi=s[5]) for s in sens],
           must=[dict(target=m[0], n=m[1], w27=m[2], pp_rb=m[3], ng27=m[4], conv_ratio=m[5]) for m in must],
           median4=med4, median5=med5, overlap4=ov4, overlap5=ov5)
if JOUT: json.dump(res, open(JOUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

# ---------- xlsx ----------
wb = openpyxl.Workbook(); BOLD = Font(bold=True); BLUE = Font(color="0000FF"); KEY = PatternFill("solid", fgColor="FFF2CC")
ws = wb.active; ws.title = "輸入"
ws.append(["CoWoS 路徑輸入（藍字＝輸入；低／高＝該輸入本身的低值／高值）。片數全部來自券商經媒體轉述，台積電與 NVIDIA 都不揭露。"]); ws["A1"].font = BOLD
ws.append(["鍵", "項目", "低", "基準", "高", "標記", "來源／說明"]); [setattr(c, "font", BOLD) for c in ws[2]]
ROW = {}
for r in INP:
    ws.append(list(r)); ROW[r[0]] = ws.max_row
    for c in "CDE": ws[f"{c}{ws.max_row}"].font = BLUE
I = lambda k: f"輸入!$D${ROW[k]}"
for col, w in zip("ABCDEFG", (9, 52, 9, 10, 9, 17, 110)): ws.column_dimensions[col].width = w
c = wb.create_sheet("計算")
c.append(["CoWoS 路徑計算（公式；基準值）。封裝單位＝百萬；片數＝千片"]); c["A1"].font = BOLD
c.append(["項目", "2025（校準）", "2026（校驗）", "2027", "說明"]); [setattr(x, "font", BOLD) for x in c[2]]
def row(label, f25, f26, f27, note="", key=False):
    c.append([label, f25, f26, f27, note])
    if key:
        for x in c[c.max_row]: x.fill = KEY
    return c.max_row
r_e = row("有效片數＝(時差×前一年＋(1−時差)×當年)×(1−非 GPU 占比)",
          f"=({I('lag')}*{I('w24')}+(1-{I('lag')})*{I('w25')})*(1-{I('ng25')})", f"=({I('lag')}*{I('w25')}+(1-{I('lag')})*{I('w26')})*(1-{I('ng26')})",
          f"=({I('lag')}*{I('w26')}+(1-{I('lag')})*{I('w27')})*(1-{I('ng27')})", "千片")
r_w = row("每封裝用片（片／封裝）＝Σ 組合占比 ÷ 每片好封裝",
          f"=({I('u25_bw')}/{I('pp_bw')}+{I('u25_hp')}/{I('pp_hp')})/({I('u25_bw')}+{I('u25_hp')})",
          f"=({I('m26_bw')}/{I('pp_bw')}+{I('m26_rb')}/{I('pp_rb')}+{I('m26_hp')}/{I('pp_hp')})/({I('m26_bw')}+{I('m26_rb')}+{I('m26_hp')})",
          f"=({I('m27_bw')}/{I('pp_bw')}+{I('m27_rb')}/{I('pp_rb')}+{I('m27_ru')}/{I('pp_ru')})/({I('m27_bw')}+{I('m27_rb')}+{I('m27_ru')})", "世代組合只取比例")
r_n = row("2025 所需片數（千片）＝營收推得封裝 × 每封裝用片", f"=({I('u25_bw')}+{I('u25_hp')})*B{r_w}*1000", None, None)
r_k = row("轉換係數 κ＝所需片數 ÷ 有效片數（2025 校準，之後各年沿用）", f"=B{r_n}/B{r_e}", f"=B{r_n}/B{r_e}", f"=B{r_n}/B{r_e}", "< 1 的部分＝良率以外的損耗、庫存增加、片數高估等，無法再拆")
r_p = row("NVIDIA 加速器封裝（百萬）＝κ × 有效片數 ÷ 每封裝用片", f"={I('u25_bw')}+{I('u25_hp')}", f"=C{r_k}*C{r_e}/C{r_w}/1000", f"=D{r_k}*D{r_e}/D{r_w}/1000", "2025＝錨點本身（循環）", key=True)
r_x = row("　其中不含 Hopper", f"={I('u25_bw')}", f"=C{r_p}*(1-{I('m26_hp')}/({I('m26_bw')}+{I('m26_rb')}+{I('m26_hp')}))", f"=D{r_p}", "2026 對照：自建推得值 620 萬（不含 Hopper）")
row("對照：自建（rubin-price 互驗頁）", None, 6.203, 6.1325, "2026＝NVIDIA「2,000 萬晶粒」推得；2027＝四路徑基準中位數")
row("對照：Siliconomics（2026＝v0.21；2027＝v0.20 券商值）", None, 7.6, 9.12, "2026 含 Hopper 0.5")
row("CoWoS 路徑 ÷ 自建 − 1", None, f"=C{r_x}/6.203-1", f"=D{r_p}/6.1325-1", "2026 為校驗", key=True)
r_a = row("2027 隱含平均每封裝營收（千美元）＝指引推得營收 ÷ 封裝", None, None, f"={I('rev27')}/D{r_p}", "對照 rubin-price 基準 98.6")
row("2027 隱含 Rubin 每封裝營收（千美元）", None, None,
    f"=(D{r_a}-{I('m27_bw')}/({I('m27_bw')}+{I('m27_rb')}+{I('m27_ru')})*{I('p_bw')})/(({I('m27_rb')}+{I('m27_ru')}*{I('ru_x')})/({I('m27_bw')}+{I('m27_rb')}+{I('m27_ru')}))",
    "對照 rubin-price 三線重疊 95.9／114.1／139.3", key=True)
for col, w in zip("ABCDE", (62, 16, 16, 16, 70)): c.column_dimensions[col].width = w
s = wb.create_sheet("敏感度")
s.append(["單因子敏感度：每次只動一個輸入到它的低值或高值，其餘留基準（值由 build.py 算出；依影響大小排序）。先看這一頁，再看計算頁的基準值。"]); s["A1"].font = BOLD
s.append(["鍵", "輸入", "輸入低", "輸入高", "2027 封裝（輸入低）", "2027 封裝（輸入高）", "低：對基準", "高：對基準"]); [setattr(x, "font", BOLD) for x in s[2]]
for x in sens: s.append([x[0], x[1], x[2], x[3], round(x[4], 3), round(x[5], 3), round(x[6], 4), round(x[7], 4)])
s.append([]); s.append(["基準", None, None, None, round(base["n27"], 3)])
s.append(["合併區間（各輸入對數偏離平方和開根號，視為獨立）", None, None, None, round(rng[0], 3), round(rng[1], 3)])
for col, w in zip("ABCDEFGH", (9, 58, 9, 9, 18, 18, 12, 12)): s.column_dimensions[col].width = w
m = wb.create_sheet("成立條件")
m.append(["要讓 CoWoS 路徑算出下列顆數，單一輸入必須是多少（其餘留基準；值由 build.py 反解）"]); m["A1"].font = BOLD
m.append(["目標", "百萬封裝", "NVIDIA 2027 用片量（千片；基準 1,222；2026 年 780）", "Rubin 每片好封裝（基準 9.68；幾何區間 7.2–12.35）", "2027 非 GPU 占比（基準 8%）", "或：2027 轉換係數 ÷ 2025 校準值"]); [setattr(x, "font", BOLD) for x in m[2]]
for x in must: m.append([x[0], x[1]] + [None if v is None else round(v, 3) for v in x[2:]])
m.append([]); m.append(["併入四路徑後", "四路徑", "加 CoWoS 路徑（五路徑）"])
m.append(["基準值中位數（百萬封裝）", round(med4, 4), round(med5, 4)])
m.append(["重疊區 下緣", round(ov4[0], 3), round(ov5[0], 3)]); m.append(["重疊區 上緣", round(ov4[1], 3), round(ov5[1], 3)])
for col, w in zip("ABCDEF", (34, 12, 44, 44, 26, 34)): m.column_dimensions[col].width = w
q = wb.create_sheet("來源")
q.append(["編號", "內容", "出處", "網址", "等級"]); [setattr(x, "font", BOLD) for x in q[1]]
for x in SRC: q.append(list(x))
for col, w in zip("ABCDE", (6, 90, 44, 90, 14)): q.column_dimensions[col].width = w
wb.save(OUT)
print(json.dumps({k: (round(v, 4) if isinstance(v, float) else v) for k, v in base.items()}, ensure_ascii=False))
print("range", [round(x, 3) for x in rng], "med4", round(med4, 4), "med5", round(med5, 4), "ov5", [round(x, 3) for x in ov5])
for x in sens: print(f"{x[0]:8s} {x[4]:.3f} {x[5]:.3f} {x[6]:+.1%} {x[7]:+.1%}")
for x in must: print(x)
