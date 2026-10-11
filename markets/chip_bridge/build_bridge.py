# 每 GW 晶片內容橋接表（步驟 0）建檔程式：openpyxl 寫公式，之後由 LibreOffice 重算
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L

BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
wb = openpyxl.Workbook()

# ---------------- In：輸入 ----------------
ws = wb.active; ws.title = "In"
ws["A1"] = "In — 輸入（藍字）。基準欄＝E（GB300）、H（VR200）。Tokenomics 列取自 v5.27（model/20261008_Tokenomics_v5.27.xlsx，SHA-256 50f678b4…7ba7）"
hdr = ["鍵", "項目", "單位", "GB300 低", "GB300 基準", "GB300 高", "VR200 低", "VR200 基準", "VR200 高", "標記", "來源／說明"]
for j, h in enumerate(hdr, 1):
    c = ws.cell(3, j, h); c.font = BOLD; c.fill = HDR
rows = [
 ("sec","A. Tokenomics v5.27（算力收支；只讀不改）"),
 ("rack","機架單價","$M/架",4.0,5.0,6.5,7.8,8.4,9.1,"TK","Spec_Rack E11:F13（非 Interface 具名範圍；已提修改申請 R1）"),
 ("racks","每 GW 機架數","架",6764,6764,6764,4052,4052,4052,"TK","IF_RacksPerGW"),
 ("gpus","每 GW GPU 數","顆",487008,487008,487008,291744,291744,291744,"TK","IF_GPUsPerGW"),
 ("so","Scale-out 網路 ÷ 機架價格","%",0.05,0.10,0.15,0.05,0.10,0.15,"TK","Inputs D13:F13（Analogy；非 Interface，R1）"),
 ("stor","儲存與管理","$/GPU",300,500,800,300,500,800,"TK","Inputs D14:F14（Assumed；非 Interface，R1）"),
 ("capit","每 GW 資本支出 — IT 設備","$B",28.5549024,37.445504,50.9505064,33.2734032,37.586352,42.6375752,"TK","IF_CapexIT"),
 ("capfac","每 GW 資本支出 — 廠房","$B",10.212,12.67,16.186,10.212,12.67,16.186,"TK","IF_CapexFacility"),
 ("captot","每 GW 資本支出 — 合計","$B",38.7669024,50.115504,67.1365064,43.4854032,50.256352,58.8235752,"TK","IF_CapexTotal"),
 ("sec","B. 機架內構件（本專案蒐集；每數字盡量兩來源）"),
 ("gpu_n","每架 GPU 封裝數","顆",72,72,72,72,72,72,"Verified","NVIDIA 規格頁"),
 ("gpu_p","GPU 模組售價（NVIDIA 對買方）","$K/顆",40,47,50,50,55,60,"Interested-party","GB300：Data Gravity 採購單 72 模組約 $3.4M（Wing VC 2026-08）；Epoch B200 $30–40K。VR200：Morgan Stanley $55K（內文；標題 $50K，Tom's Hardware 2026-05-21）"),
 ("hbm_gb","HBM 容量","GB/顆",288,288,288,288,288,288,"Verified","NVIDIA：GB300 20 TB、VR200 20.7 TB ÷ 72"),
 ("hbm_p","HBM 單價（2026 合約）","$/GB",12,15.5,17,15.6,16.7,17.2,"Derived","GB300 HBM3E：Epoch $14–17/GB；Silicon Analysts B300 約 $4,900/顆（=$17/GB）。VR200 HBM4：首爾經濟日報 12 層 36GB 約 $600（=$16.7/GB，2026-10-01）；UBS 每顆 $4,943（=$17.2/GB，2026-08-10）"),
 ("hbm_p27","HBM4 單價（2027 合約情境）","$/GB",None,None,None,30,36,37,"Interested-party","首爾經濟日報 約 $1,300/36GB（2027）；TrendForce 2027 ASP +121%（2026-09-29）；UBS +79%（低）"),
 ("tsmc","台積電：邏輯晶圓＋CoWoS","$K/顆",1.9,2.3,2.7,2.3,2.8,3.5,"Derived","B200：Silicon Analysts 邏輯 $850＋封裝 $1,100；Epoch 約 $1.9K；B300 封裝溢價 25–40%（Epoch）。Rubin：UBS『GPU 相關扣 HBM』約 $4.3K（含周邊）為上限參考；N3 晶圓較 4NP 貴"),
 ("oth","其他製造：基板、測試、輔助、封裝良率損失","$K/顆",1.0,1.5,2.0,1.0,1.5,2.0,"Derived","Epoch B200：封裝良率損失 $1,000（$430–1,700）＋輔助 $480；Rubin 為類比"),
 ("cpu_n","每架 CPU 數","顆",36,36,36,36,36,36,"Verified","NVIDIA：36 Grace／36 Vera"),
 ("cpu_p","CPU 單價","$K/顆",3,4,5,4,5,6,"Analogy","Vera $5K（Morgan Stanley）；Grace 以 Vera 類比"),
 ("lp_tb","LPDDR5X 容量（每架）","TB",17,17,36,28,28,54,"Verified／Interested-party","GB300：NVIDIA 規格頁 17 TB（參考架構頁寫每顆 Grace 1 TB＝36 TB，作高）。VR200：NVIDIA 初步規格 54 TB（高）；TrendForce 2026-06-10、GF Securities 2026-07-26：SOCAMM 減半，約 28 TB（基準）"),
 ("lp_p","LPDDR5X 單價","$/GB",8,13,21,8,13,21,"Interested-party","SemiAnalysis 2026Q1 約 $8；UBS 2026-08 隱含約 $12.9（$19,355／1.5 TB）；TrendForce 2026Q2 預估約 $19–20、長約上限約 $21（Wccftech 2026-05-04）"),
 ("nic_n","每架 SuperNIC 數","顆",72,72,72,72,144,144,"Verified／Derived","GB300：NVIDIA 參考架構 每匣 4 顆 CX-8。VR200：NVIDIA 技術部落格圖示 每匣 8 顆 CX-9（推導 144）"),
 ("nic_p","SuperNIC 單價","$K/顆",1.5,2.0,2.5,2.0,3.0,4.0,"Analogy／Assumed","CX-8 零售約 $2,089（Newegg）；大客戶價較低。CX-9 無報價"),
 ("dpu_n","每架 DPU 數","顆",18,18,18,18,18,18,"Verified","每匣 1 顆 BlueField-3／BlueField-4"),
 ("dpu_p","DPU 單價","$K/顆",1.5,2.5,3.5,3,4,5,"Assumed","無可靠報價"),
 ("nand","機架內 NAND／SSD","$M/架",0.05,0.10,0.20,0.2,1.0,1.2,"Assumed／Interested-party","GB300：每匣 4 顆 E1.S＋1 M.2（容量未揭露）。VR200：Morgan Stanley 約 $1M 以上（Tom's Hardware）；Tokenomics SRC_HW_008 引用同一篇，**不是獨立第二來源**；低值 0.2 為 Assumed（若 $1M 指的是架外 BlueField-4 STX 儲存）"),
 ("nvsw","NVLink 交換匣＋背板 ÷ 機架","%",0.04,0.05,0.06,0.04,0.05,0.06,"Derived","Wing VC：NVIDIA 內容 >92%（含 NVLink）− 運算匣約 87% ≈ 5%"),
 ("nvsw_chip","交換匣中交換晶片占比","%",0.4,0.5,0.6,0.4,0.5,0.6,"Assumed","GB300 18 顆、VR200 36 顆 NVLink 交換晶片（NVIDIA）"),
 ("nonnv","非 NVIDIA 機架內容 ÷ 機架","%",0.05,0.07,0.08,0.05,0.07,0.08,"Interested-party","Wing VC：電源架、匯流排、冷板、鋼構、線纜、整合合計 <8%"),
 ("cool","其中：機架內液冷","$K/架",50,50,50,56,56,56,"Interested-party","Morgan Stanley（Tom's Hardware）：GB300 約 $49,860、Vera Rubin 約 $56K"),
]
R = {}
r = 4
for row in rows:
    if row[0] == "sec":
        ws.cell(r, 1, row[1]).font = BOLD; r += 1; continue
    key, name, unit, *vals, tag, src = row
    ws.cell(r, 1, key); ws.cell(r, 2, name); ws.cell(r, 3, unit)
    for j, v in enumerate(vals):
        if v is not None:
            c = ws.cell(r, 4 + j, v); c.font = BLUE
    ws.cell(r, 10, tag); ws.cell(r, 11, src)
    R[key] = r; r += 1
for col, w in zip("ABCDEFGHIJK", [9, 34, 9, 9, 10, 9, 9, 10, 9, 18, 110]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "D4"

def g(key, gen):  # 基準值儲存格
    return f"In!{'E' if gen=='GB' else 'H'}{R[key]}"

# ---------------- Rack：每架拆分（$M）----------------
rk = wb.create_sheet("Rack")
rk["A1"] = "Rack — 每架拆分（$M/架；各欄皆用基準值）。D 欄：2027 HBM 價、NVIDIA 吸收（售價不變）；E 欄：2027 HBM 價、全額轉嫁（GPU 與機架售價上調）"
cols = {"B": "GB", "C": "VR", "D": "VR", "E": "VR"}
heads = ["項目", "GB300 NVL72", "VR200（2026 HBM 價）", "VR200（2027 HBM 價，NVIDIA 吸收）", "VR200（2027 HBM 價，全額轉嫁）", "說明"]
for j, h in enumerate(heads, 1):
    c = rk.cell(3, j, h); c.font = BOLD; c.fill = HDR
K = {}
def put(label, fdict, note="", key=None, fmt="0.000"):
    global rr
    rk.cell(rr, 1, label)
    for col, f in fdict.items():
        c = rk.cell(rr, " BCDE".index(col) + 1, f); c.number_format = fmt
    rk.cell(rr, 6, note)
    if key: K[key] = rr
    rr += 1
rr = 4
put("HBM 單價 $/GB", {"B": f"={g('hbm_p','GB')}", "C": f"={g('hbm_p','VR')}", "D": f"={g('hbm_p27','VR')}", "E": f"={g('hbm_p27','VR')}"}, key="hp", fmt="0.0")
put("每顆 HBM 金額 $K", {c: f"={g('hbm_gb',cols[c])}*{c}{K['hp']}/1000" for c in "BCDE"}, key="hbmk", fmt="0.00")
put("GPU 模組售價 $K", {"B": f"={g('gpu_p','GB')}", "C": f"={g('gpu_p','VR')}", "D": f"={g('gpu_p','VR')}", "E": f"={g('gpu_p','VR')}+(E{K['hbmk']}-C{K['hbmk']})"}, "E 欄：HBM 增加額全額加到售價", key="gp", fmt="0.00")
put("  台積電（邏輯＋CoWoS）$K", {c: f"={g('tsmc',cols[c])}" for c in "BCDE"}, key="tsk", fmt="0.00")
put("  其他製造 $K", {c: f"={g('oth',cols[c])}" for c in "BCDE"}, key="otk", fmt="0.00")
put("  NVIDIA 晶片毛利 $K（殘差）", {c: f"={c}{K['gp']}-{c}{K['hbmk']}-{c}{K['tsk']}-{c}{K['otk']}" for c in "BCDE"}, key="nvk", fmt="0.00")
put("  NVIDIA 晶片層級毛利率", {c: f"={c}{K['nvk']}/{c}{K['gp']}" for c in "BCDE"}, key="nvgm", fmt="0.0%")
put("  HBM ÷ GPU 售價", {c: f"={c}{K['hbmk']}/{c}{K['gp']}" for c in "BCDE"}, key="hbmsh", fmt="0.0%")
put("  台積電 ÷ GPU 售價", {c: f"={c}{K['tsk']}/{c}{K['gp']}" for c in "BCDE"}, key="tssh", fmt="0.0%")
rr += 1
put("機架售價 $M", {"B": f"={g('rack','GB')}", "C": f"={g('rack','VR')}", "D": f"={g('rack','VR')}", "E": f"={g('rack','VR')}+{g('gpu_n','VR')}*(E{K['gp']}-C{K['gp']})/1000"}, "B–D 欄＝Tokenomics 基準；E 欄＝加上轉嫁額", key="rack")
put("GPU 模組（72 顆）", {c: f"={g('gpu_n',cols[c])}*{c}{K['gp']}/1000" for c in "BCDE"}, key="mod")
put("  HBM", {c: f"={g('gpu_n',cols[c])}*{c}{K['hbmk']}/1000" for c in "BCDE"}, key="m_hbm")
put("  台積電（邏輯＋CoWoS）", {c: f"={g('gpu_n',cols[c])}*{c}{K['tsk']}/1000" for c in "BCDE"}, key="m_ts")
put("  其他製造", {c: f"={g('gpu_n',cols[c])}*{c}{K['otk']}/1000" for c in "BCDE"}, key="m_ot")
put("  NVIDIA 晶片毛利", {c: f"={g('gpu_n',cols[c])}*{c}{K['nvk']}/1000" for c in "BCDE"}, key="m_nv")
put("CPU（Grace／Vera）", {c: f"={g('cpu_n',cols[c])}*{g('cpu_p',cols[c])}/1000" for c in "BCDE"}, key="cpu")
put("非 HBM DRAM（LPDDR5X）", {c: f"={g('lp_tb',cols[c])}*{g('lp_p',cols[c])}/1000" for c in "BCDE"}, "TB × $/GB ÷ 1000 ＝ $M", key="lp")
put("NAND／SSD（機架內）", {c: f"={g('nand',cols[c])}" for c in "BCDE"}, key="nand")
put("SuperNIC", {c: f"={g('nic_n',cols[c])}*{g('nic_p',cols[c])}/1000" for c in "BCDE"}, key="nic")
put("DPU", {c: f"={g('dpu_n',cols[c])}*{g('dpu_p',cols[c])}/1000" for c in "BCDE"}, key="dpu")
put("NVLink 交換匣＋背板", {c: f"={g('nvsw',cols[c])}*{g('rack',cols[c])}" for c in "BCDE"}, "以 Tokenomics 機架價計（不隨 HBM 轉嫁放大）", key="nvsw")
put("  NVLink 交換晶片", {c: f"={c}{K['nvsw']}*{g('nvsw_chip',cols[c])}" for c in "BCDE"}, key="nvsw_c")
put("  交換匣其他（含銅纜背板）", {c: f"={c}{K['nvsw']}-{c}{K['nvsw_c']}" for c in "BCDE"}, key="nvsw_o")
put("非 NVIDIA 機架內容（電源、冷卻、機構、整合）", {c: f"={g('nonnv',cols[c])}*{g('rack',cols[c])}" for c in "BCDE"}, key="nonnv")
put("  其中：液冷", {c: f"={g('cool',cols[c])}/1000" for c in "BCDE"}, key="cool")
put("殘差：匣級電路板、供電與 NVIDIA／ODM 匣級加價", {c: f"={c}{K['rack']}-{c}{K['mod']}-{c}{K['cpu']}-{c}{K['lp']}-{c}{K['nand']}-{c}{K['nic']}-{c}{K['dpu']}-{c}{K['nvsw']}-{c}{K['nonnv']}" for c in "BCDE"}, "負值＝輸入互相矛盾（構件加總超過售價）", key="resid")
put("  殘差 ÷ 機架售價", {c: f"={c}{K['resid']}/{c}{K['rack']}" for c in "BCDE"}, key="residsh", fmt="0.0%")
put("核對：構件＋殘差－售價（應為 0）", {c: f"={c}{K['mod']}+{c}{K['cpu']}+{c}{K['lp']}+{c}{K['nand']}+{c}{K['nic']}+{c}{K['dpu']}+{c}{K['nvsw']}+{c}{K['nonnv']}+{c}{K['resid']}-{c}{K['rack']}" for c in "BCDE"}, key="chk1")
for col, w in zip("ABCDEF", [44, 14, 18, 22, 22, 50]):
    rk.column_dimensions[col].width = w

# ---------------- Bridge：每 GW 橋接表（$B/GW）----------------
br = wb.create_sheet("Bridge")
br["A1"] = "Bridge — 每 GW 晶片內容橋接表（軸 A：買方每一塊錢分到哪裡；$B/GW IT 關鍵電力）"
br["A2"] = "每架項目 × 每 GW 機架數；scale-out 與儲存依 Tokenomics 口徑；廠房＝IF_CapexFacility。E 欄的 IT 合計高於 Tokenomics（轉嫁額），見核對列"
bh = ["項目（縮排＝樹狀層級）", "GB300", "VR200 2026", "VR200 2027 吸收", "VR200 2027 轉嫁", "GB300 ÷ 總 capex", "VR200 2026 ÷ 總 capex", "VR200 2027 吸收 ÷ 總", "VR200 2027 轉嫁 ÷ 總", "下游歸屬"]
for j, h in enumerate(bh, 1):
    c = br.cell(4, j, h); c.font = BOLD; c.fill = HDR
B = {}
def racks(c): return g('racks', cols[c])
def bput(label, f, owner="", key=None, bold=False):
    global br_r
    br.cell(br_r, 1, label).font = Font(bold=bold)
    for c in "BCDE":
        x = br.cell(br_r, " BCDE".index(c) + 1, f(c)); x.number_format = "0.000"
    br.cell(br_r, 10, owner)
    if key: B[key] = br_r
    br_r += 1
br_r = 5
fr = lambda k: (lambda c: f"=Rack!{c}{K[k]}*{racks(c)}/1000")
rows_tree = [
 ("每 GW capex 合計", None, "", "tot", True),
 ("├ 廠房（機房、電力、冷卻）", lambda c: f"={g('capfac',cols[c])}", "電力基建專案", "fac", False),
 ("└ IT 設備", None, "", "it", True),
 ("   ├ 伺服器／機架", lambda c: f"=Rack!{c}{K['rack']}*{racks(c)}/1000", "", "srv", True),
 ("   │  ├ 加速器封裝（GPU 模組）", fr("mod"), "", "mod", True),
 ("   │  │  ├ GPU 晶片：NVIDIA 毛利（設計）", fr("m_nv"), "NVDA", "m_nv", False),
 ("   │  │  ├ GPU 晶片＋CoWoS：台積電", fr("m_ts"), "台積電（軸 B）", "m_ts", False),
 ("   │  │  ├ HBM", fr("m_hbm"), "SK 海力士／三星／美光", "m_hbm", False),
 ("   │  │  └ 其他製造（基板、測試、良率損失）", fr("m_ot"), "基板廠、封測", "m_ot", False),
 ("   │  ├ CPU（Grace／Vera）", fr("cpu"), "NVDA（台積電代工）", "cpu", False),
 ("   │  ├ 非 HBM DRAM（LPDDR5X）", fr("lp"), "記憶體三家", "lp", False),
 ("   │  ├ NAND／SSD（機架內）", fr("nand"), "NAND 廠", "nand", False),
 ("   │  ├ SuperNIC＋DPU（scale-out 端點，機架內）", lambda c: f"=(Rack!{c}{K['nic']}+Rack!{c}{K['dpu']})*{racks(c)}/1000", "NVDA", "nicdpu", False),
 ("   │  ├ 非 NVIDIA 機架內容（電源、冷卻、機構、整合）", fr("nonnv"), "台系供應鏈（待定）", "nonnv", False),
 ("   │  └ 殘差：匣級電路板、供電、NVIDIA／ODM 匣級加價", fr("resid"), "NVDA／ODM（未拆）", "resid", False),
 ("   ├ 網通", None, "", "net", True),
 ("   │  ├ scale-up：NVLink 交換晶片", fr("nvsw_c"), "NVDA", "nvsw_c", False),
 ("   │  ├ scale-up：交換匣其他（含銅纜背板）", fr("nvsw_o"), "NVDA／線纜", "nvsw_o", False),
 ("   │  └ scale-out：機架外（交換器、光模組、線纜；步驟 3 拆分）", lambda c: f"={g('so',cols[c])}*{g('rack',cols[c])}*{racks(c)}/1000", "AVGO、NVDA、光模組", "so", False),
 ("   └ 儲存與管理", lambda c: f"={g('stor',cols[c])}*{g('gpus',cols[c])}/1E9", "未拆", "stor", False),
]
for label, f, owner, key, bold in rows_tree:
    bput(label, f if f else (lambda c: None), owner, key, bold)
# 合計公式
for c in "BCDE":
    ci = " BCDE".index(c) + 1
    br.cell(B["net"], ci, f"={c}{B['nvsw_c']}+{c}{B['nvsw_o']}+{c}{B['so']}")
    br.cell(B["it"], ci, f"={c}{B['srv']}-{c}{B['nvsw_c']}-{c}{B['nvsw_o']}+{c}{B['net']}+{c}{B['stor']}")
    br.cell(B["tot"], ci, f"={c}{B['fac']}+{c}{B['it']}")
    for k in ("net", "it", "tot"):
        br.cell(B[k], ci).number_format = "0.000"
br.cell(B["srv"], 10, "註：NVLink 交換匣實體在機架內，此處移到『網通』列示，IT 合計已扣除避免重複")
# 占比欄
for k, rowi in B.items():
    for c, pc in zip("BCDE", "FGHI"):
        x = br[f"{pc}{rowi}"]; x.value = f"={c}{rowi}/{c}${B['tot']}"; x.number_format = "0.0%"
br_r += 1
br.cell(br_r, 1, "核對").font = BOLD; br_r += 1
B["c1"] = br_r; br.cell(br_r, 1, "IT 合計 − Tokenomics IF_CapexIT（B–D 應為 0；E＝轉嫁額）")
for c in "BCDE":
    x = br.cell(br_r, " BCDE".index(c) + 1, f"={c}{B['it']}-{g('capit',cols[c])}"); x.number_format = "0.000"
br_r += 1
B["c2"] = br_r; br.cell(br_r, 1, "合計 − Tokenomics IF_CapexTotal（B–D 應為 0）")
for c in "BCDE":
    x = br.cell(br_r, " BCDE".index(c) + 1, f"={c}{B['tot']}-{g('captot',cols[c])}"); x.number_format = "0.000"
br_r += 2
br.cell(br_r, 1, "記憶體彙總（$B/GW）").font = BOLD; br_r += 1
def sput(label, f, key, fmt="0.000"):
    global br_r
    br.cell(br_r, 1, label)
    for c in "BCDE":
        x = br.cell(br_r, " BCDE".index(c) + 1, f(c)); x.number_format = fmt
    B[key] = br_r; br_r += 1
sput("HBM", lambda c: f"={c}{B['m_hbm']}", "s_hbm")
sput("非 HBM（LPDDR5X＋機架內 NAND）", lambda c: f"={c}{B['lp']}+{c}{B['nand']}", "s_non")
sput("  其中 DRAM（LPDDR5X）", lambda c: f"={c}{B['lp']}", "s_lp")
sput("記憶體合計（不含儲存層、機架外）", lambda c: f"={c}{B['s_hbm']}+{c}{B['s_non']}", "s_mem")
sput("HBM ÷ 記憶體合計", lambda c: f"={c}{B['s_hbm']}/{c}{B['s_mem']}", "s_hsh", "0.0%")
sput("HBM ÷ IT capex", lambda c: f"={c}{B['s_hbm']}/{c}{B['it']}", "s_hit", "0.0%")
sput("DRAM（HBM＋LPDDR）÷ IT capex", lambda c: f"=({c}{B['s_hbm']}+{c}{B['s_lp']})/{c}{B['it']}", "s_dit", "0.0%")
br_r += 1
br.cell(br_r, 1, "回饋算力收支的摘要數字（每顆加速器）").font = BOLD; br_r += 1
sput("每顆加速器 HBM 金額 $K", lambda c: f"=Rack!{c}{K['hbmk']}", "f1", "0.00")
sput("HBM ÷ GPU 售價", lambda c: f"=Rack!{c}{K['hbmsh']}", "f2", "0.0%")
sput("台積電 ÷ GPU 售價", lambda c: f"=Rack!{c}{K['tssh']}", "f3", "0.0%")
sput("NVIDIA 晶片層級毛利率", lambda c: f"=Rack!{c}{K['nvgm']}", "f4", "0.0%")
sput("NVIDIA 合計 ÷ IT capex（GPU 毛利＋CPU＋NIC/DPU＋NVLink＋殘差）", lambda c: f"=({c}{B['m_nv']}+{c}{B['cpu']}+{c}{B['nicdpu']}+{c}{B['nvsw_c']}+{c}{B['nvsw_o']}+{c}{B['resid']})/{c}{B['it']}", "f5", "0.0%")
for col, w in zip("ABCDEFGHIJ", [58, 10, 12, 15, 15, 12, 14, 15, 15, 30]):
    br.column_dimensions[col].width = w

# ---------------- Sens：單變數敏感度（VR200 2026 基準）----------------
se = wb.create_sheet("Sens")
se["A1"] = "Sens — 單變數敏感度（VR200；一次只動一個藍字輸入；其餘＝基準）。先看本頁，再看 Bridge 基準"
sh = ["情境", "HBM $/GB", "LPDDR TB/架", "LPDDR $/GB", "GPU 售價 $K", "機架售價 $M", "台積電 $K/顆", "HBM $B/GW", "HBM ÷ IT capex", "LPDDR $B/GW", "NVIDIA 晶片毛利率", "機架殘差 ÷ 售價", "HBM ÷ GPU 售價"]
for j, h in enumerate(sh, 1):
    c = se.cell(3, j, h); c.font = BOLD; c.fill = HDR
V = lambda k: g(k, "VR")
base = [f"={V('hbm_p')}", f"={V('lp_tb')}", f"={V('lp_p')}", f"={V('gpu_p')}", f"={V('rack')}", f"={V('tsmc')}"]
scen = [("基準（2026 HBM 價）", {}),
        ("HBM 低（$15.6）", {0: f"=In!G{R['hbm_p']}"}),
        ("HBM 2027 低（$30，UBS +79%）", {0: f"=In!G{R['hbm_p27']}"}),
        ("HBM 2027（$36，TrendForce +121%）", {0: f"=In!H{R['hbm_p27']}"}),
        ("LPDDR 54 TB（NVIDIA 初步規格）", {1: f"=In!I{R['lp_tb']}"}),
        ("LPDDR $8／GB", {2: f"=In!G{R['lp_p']}"}),
        ("LPDDR $21／GB", {2: f"=In!I{R['lp_p']}"}),
        ("GPU 售價 $50K", {3: f"=In!G{R['gpu_p']}"}),
        ("GPU 售價 $60K", {3: f"=In!I{R['gpu_p']}"}),
        ("機架 $7.8M（MS BOM）", {4: f"=In!G{R['rack']}"}),
        ("機架 $9.1M（Bernstein）", {4: f"=In!I{R['rack']}"}),
        ("台積電 $2.3K", {5: f"=In!G{R['tsmc']}"}),
        ("台積電 $3.5K", {5: f"=In!I{R['tsmc']}"}),
        ("LPDDR 54 TB × $21／GB（記憶體最緊）", {1: f"=In!I{R['lp_tb']}", 2: f"=In!I{R['lp_p']}"})]
for i, (nm, ov) in enumerate(scen):
    rrow = 4 + i
    se.cell(rrow, 1, nm)
    for j in range(6):
        c = se.cell(rrow, 2 + j, ov.get(j, base[j])); c.font = BLUE if j in ov else Font()
    racksV, gpusV = V('racks'), V('gpus')
    hb = f"{gpusV}*{V('hbm_gb')}*B{rrow}/1E9"
    it = f"({racksV}*F{rrow}*(1+{V('so')})+{gpusV}*{V('stor')}/1E6/1000*1000)/1000"
    # IT capex $B = 機架數 × 機架價 $M × (1+so) /1000 + GPU 數 × $/GPU /1e9
    it = f"({racksV}*F{rrow}*(1+{V('so')})/1000+{gpusV}*{V('stor')}/1E9)"
    se.cell(rrow, 8, f"={hb}").number_format = "0.000"
    se.cell(rrow, 9, f"=H{rrow}/{it}").number_format = "0.0%"
    se.cell(rrow, 10, f"={racksV}*C{rrow}*D{rrow}/1E6").number_format = "0.000"
    se.cell(rrow, 11, f"=(E{rrow}-{V('hbm_gb')}*B{rrow}/1000-G{rrow}-{V('oth')})/E{rrow}").number_format = "0.0%"
    comp = (f"{V('gpu_n')}*E{rrow}/1000+{V('cpu_n')}*{V('cpu_p')}/1000+C{rrow}*D{rrow}/1000+{V('nand')}"
            f"+{V('nic_n')}*{V('nic_p')}/1000+{V('dpu_n')}*{V('dpu_p')}/1000+F{rrow}*({V('nvsw')}+{V('nonnv')})")
    se.cell(rrow, 12, f"=(F{rrow}-({comp}))/F{rrow}").number_format = "0.0%"
    se.cell(rrow, 13, f"={V('hbm_gb')}*B{rrow}/1000/E{rrow}").number_format = "0.0%"
se.cell(20, 1, "讀法：VR200 的 HBM ÷ IT capex 在 2026 價約 3.7%、2027 價（$36/GB）約 8.0%；LPDDR 54 TB × $21/GB 時非 HBM DRAM 每 GW 是 HBM 的 3 倍以上。機架殘差變負＝輸入組合不可能（售價必須上調或 NVIDIA 讓利）")
se.column_dimensions["A"].width = 36
for col in "BCDEFGHIJKLM": se.column_dimensions[col].width = 13

# ---------------- README ----------------
rd = wb.create_sheet("README", 0)
lines = [
 ("每 GW 晶片內容橋接表（AI 半導體專案 步驟 0）v0.1　2026-10-08", True),
 ("模型命題：買方每投入 1 GW 的資料中心資本支出，晶片內容（運算、HBM、非 HBM 記憶體、網通晶片、代工）各分到多少，作為後續 HBM、網通、台積電模型的共同分母。", False),
 ("", False),
 ("取數：Tokenomics v5.27（repo YinchenChang/Tokenomics，model/20261008_Tokenomics_v5.27.xlsx，SHA-256 50f678b4c78e9bfd3d5b7152c465148f3372d1ab33dfe5a18d7b089e1f207ba7）。具名範圍：IF_RacksPerGW、IF_GPUsPerGW、IF_CapexIT、IF_CapexFacility、IF_CapexTotal；機架單價、scale-out 比率、儲存 $/GPU 不在 Interface，暫讀 Spec_Rack／Inputs 原格（修改申請 R1）。", False),
 ("範圍：NVIDIA GB300 NVL72、VR200 NVL72 各一個代表性 1 GW；不含 ASIC 與世代組合（屬步驟 1 加速器數量層）。", False),
 ("", False),
 ("驅動因子 → 推導量（頁：輸入 → 推導 → 因果方向）", True),
 ("In：Tokenomics 每 GW 機架數、GPU 數、IT／廠房 capex、機架價；機架內構件單價與數量 → 供 Rack、Bridge、Sens 使用", False),
 ("Rack：GPU 售價、HBM GB × $/GB、台積電、其他製造 → NVIDIA 晶片毛利（殘差）；機架價 − 各構件 → 匣級殘差。HBM 價↑ → NVIDIA 毛利↓（吸收）或 機架價↑（轉嫁）", False),
 ("Bridge：每架金額 × 每 GW 機架數 → 每 GW 樹狀拆分；核對 IT 合計＝IF_CapexIT、合計＝IF_CapexTotal", False),
 ("Sens：一次動一個輸入 → HBM $/GW、HBM ÷ IT capex、NVIDIA 晶片毛利率、機架殘差", False),
 ("", False),
 ("標記：TK＝Tokenomics 取數；Verified 公開一手；Interested-party 有利害關係；Analogy 類比；Assumed 無實證；Derived 推導。Analogy／Assumed 一律給區間（In 頁 低／基準／高）。", False),
 ("讀法注意：(1) 殘差不是單一公司的營收，是 NVIDIA 匣級加價、電路板、供電與 ODM 毛利的合計；(2) LPDDR、NAND 以供應商價格計，NVIDIA 對它們的加價落在殘差；(3) 2027 HBM 情境 E 欄高於 Tokenomics IT capex，差額即轉嫁額。", False),
 ("SemiAnalysis 資料未使用。", False),
]
for i, (t, b) in enumerate(lines, 1):
    rd.cell(i, 1, t).font = Font(bold=b)
rd.column_dimensions["A"].width = 160

# ---------------- Requests ----------------
rq = wb.create_sheet("Requests")
rq["A1"] = "給算力收支的修改申請與回饋（由算力收支決定是否升版）"; rq["A1"].font = BOLD
reqs = [
 ("R1", "工程類（建議）", "Interface 新增 IF_RackPrice（Spec_Rack 第 11–13 列）、IF_ScaleOutShare（Inputs 第 13 列）、IF_StoragePerGPU（Inputs 第 14 列），使下游不必讀非契約儲存格。數值不變。"),
 ("R2", "判斷類（提請注意，不建議本版改）", "VR200 機架價 8.4 以 MS 2026-05 BOM 為底。之後兩個方向相反的訊息：LPDDR5X 減半（54→28 TB，約 −$0.34M/架）與 2027 HBM4 合約價約翻倍（若全額轉嫁約 +$1.4M/架，+17%）。淨效果視 NVIDIA 吸收比例而定；建議在步驟 2 HBM 市場模型完成後再提數值。"),
 ("F1", "回饋（摘要數字）", "每顆加速器 HBM 金額：GB300 約 $4.5K、VR200 2026 價約 $4.8K、2027 價約 $10.4K；每 GW HBM：$2.17B／$1.40B／$3.03B；HBM ÷ IT capex：5.8%／3.7%／8.0%（見 Bridge）。VR200 每 GW 的 HBM 位元（約 84 PB）低於 GB300（約 140 PB），因每 GW GPU 數少 40%、每顆容量相同。"),
]
for i, rrw in enumerate(reqs, 3):
    for j, v in enumerate(rrw, 1): rq.cell(i, j, v)
rq.column_dimensions["A"].width = 6; rq.column_dimensions["B"].width = 26; rq.column_dimensions["C"].width = 150

# ---------------- Sources ----------------
so = wb.create_sheet("Sources")
srcs = [
 ("Tokenomics v5.27", "https://github.com/YinchenChang/Tokenomics/blob/master/model/20261008_Tokenomics_v5.27.xlsx"),
 ("Wing VC / Data Gravity, How Much Does an NVIDIA NVL72 Cost?（2026-08）", "https://www.wing.vc/content/how-much-does-an-nvidia-nvl72-cost"),
 ("Tom's Hardware（Morgan Stanley VR200 $7.8M、Rubin $55K、Vera $5K、記憶體約 $2M，2026-05-21）", "https://www.tomshardware.com/tech-industry/artificial-intelligence/nvidias-memory-costs-soar-485-percent-latest-ai-systems-now-cost-usd7-8-million-to-build-memory-now-comprises-25-percent-of-the-total-cost-rubin-gpus-a-mere-usd50-000-apiece"),
 ("Epoch AI, B200 cost breakdown（2025-12-10）", "https://epoch.ai/data-insights/b200-cost-breakdown"),
 ("Epoch AI, AI chip component cost shares（2026-05-21）", "https://epoch.ai/data-insights/ai-chip-component-cost-shares"),
 ("Silicon Analysts, AI Accelerator Manufacturing Cost Estimates（2026-08）", "https://siliconanalysts.com/data/ai-chip-costs"),
 ("UBS Vera Rubin superchip BOM（Let's Data Science 轉述 Wccftech／ChosunBiz，2026-08-10）", "https://letsdatascience.com/news/nvidia-vera-rubin-memory-costs-dominate-superchip-bom-1c630dcd"),
 ("Goldman Sachs 記憶體占 superchip 物料 62%／53%（Crypto Briefing，2026-08-10）", "https://cryptobriefing.com/nvidia-vera-rubin-memory-costs-goldman-sachs/"),
 ("Vera Rubin LPDDR 減半（TrendForce 2026-06-10；GF Securities 2026-07-26）", "https://letsdatascience.com/news/nvidia-reportedly-reduces-vera-rubin-memory-configuration-01e36b5d"),
 ("Wccftech, LPDDR 合約價（2026-05-04）", "https://wccftech.com/mobile-dram-prices-expected-to-increase-by-100-quarter-over-quarter-as-long-term-agreements-now-getting-signed-at-prices-as-high-as-21-gb/amp/"),
 ("Silicon Analysts（TrendForce 2027 HBM ASP +121%、首爾經濟日報 HBM4 $600→$1,300，2026-10-01）", "https://siliconanalysts.com/market/trendforce-lifts-2027-hbm-asp-forecast-to-121-hbm4-stack-seen-doubling-to-1-300-2026-10-01"),
 ("NVIDIA GB300 NVL72 規格頁", "https://www.nvidia.com/en-au/data-center/gb300-nvl72"),
 ("NVIDIA Vera Rubin NVL72 規格頁", "https://nvidia.com/ko-kr/data-center/vera-rubin-nvl72/"),
 ("NVIDIA NVL72 AI Factory 參考架構（GB300 元件）", "https://docs.nvidia.com/enterprise-reference-architectures/nvl72-ai-factory/latest/components.html"),
 ("NVIDIA 技術部落格 Vera Rubin POD（2026-09）", "https://developer.nvidia.com/blog/nvidia-vera-rubin-pod-seven-chips-five-rack-scale-systems-one-ai-supercomputer/"),
 ("Newegg ConnectX-8 零售價", "https://www.newegg.com/p/N82E16833942013"),
]
so["A1"] = "來源"; so["A1"].font = BOLD
for i, (a, b) in enumerate(srcs, 2):
    so.cell(i, 1, a); so.cell(i, 2, b)
so.column_dimensions["A"].width = 90; so.column_dimensions["B"].width = 120

import sys
wb.save(sys.argv[1])
print("saved", sys.argv[1], K, B)
