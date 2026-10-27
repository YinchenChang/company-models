"""AI 投資地圖 — 產生資料活頁簿 dist/AI投資地圖_資料_v<版本>.xlsx（事實與判斷的權威位置）。

讀：
  data/tokenomics_snapshot_v*.json   Tokenomics 官方取數快照（每 GW 金額）
  data/models_extract.json           各公司模型估值欄位（extract_models.py 產生，附來源儲存格）
  data/universe_ai_semis_v*.csv      AI半導體 候選母體（60 家；權威在 AI半導體 Project）
  data/segments_ai_semis_v*.csv      AI半導體 次產業排序
  data/prices.csv                    同日收盤價（統一比較日）
寫：一本活頁簿；黃底＝判斷輸入（Andy 可直接改），其餘為抄錄值或公式。
重建時若 dist 已有同版活頁簿，黃底欄位一律讀回（Excel 優先），不以程式預設覆寫。

用法：python3 markets/invest_map/build_xlsx.py
之後：LibreOffice 重算（build_html.py 會自動做），再產生網站。
"""
import csv
import glob
import json
import os
import re

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
VERSION = "v0.6"
DATE = "20261026"
OUT = os.path.join(HERE, "dist", f"{DATE}_AI投資地圖_資料_{VERSION}.xlsx")

YELLOW = PatternFill("solid", fgColor="FFF2CC")
HDR = PatternFill("solid", fgColor="1F3A5F")
HFONT = Font(color="FFFFFF", bold=True)
BLUE = Font(color="0000CC")
WRAP = Alignment(wrap_text=True, vertical="top")


def load(name):
    with open(os.path.join(HERE, "data", name), encoding="utf-8") as f:
        return json.load(f)


def latest(pattern):
    files = sorted(glob.glob(os.path.join(HERE, "data", pattern)))
    return files[-1]


TK = load(os.path.basename(latest("tokenomics_snapshot_v*.json")))
MX = load("models_extract.json")


def tkv(name, gen=None, case="基準"):
    it = TK["items"][name]
    if gen:
        return it["values"][gen][case], it["cells"][gen][case]
    return it["values"], it["cell"]


def tkr(name, side):
    it = TK["items"][name]
    return it["range"][side], it["range"]["cells"][side]


# ---------------------------------------------------------------- 判斷輸入（預設）
# 層級與次產業。不可替代性 1 高–3 低（AI半導體 次產業排序；算力層為 Claude 判斷）
# 價值鏈分量分：依占 AI IT 資本支出比例分級（≥20%→100；10–20%→75；5–10%→50；<5%→25），例外註明。
SEGMENTS = [
    # key, 層級, 次產業, 占 IT capex 文字, 中值(比例), 不可替代性, 價值鏈分量分, 依據
    ("model", "L5 模型商", "AI 模型商", "營收端（P×Q）", None, 2, 75,
     "收入池最上游：每 GW 第一年 P×Q 營收高於持有成本，但回本比值約 0.94，未定 [Derived]；分量分為判斷"),
    ("cloud", "L4 算力平台", "AI 雲端平台", "持有算力資產", None, 2, 100,
     "持有並出租算力；年持有成本由此層支付，是上游資本支出的付款方 [判斷]"),
    ("neocloud", "L4 算力平台", "Neocloud（GPU 租賃）", "持有算力資產", None, 3, 50,
     "同為算力持有者，但客戶集中、融資依賴高、可替代 [判斷]"),
    # L3 比例（v0.5.2，2026-10-26）：改回讀 AI半導體 次產業排序 v0.2.3 的「GB300 對照 ÷ IT」欄（company-models #166；各塊互斥、合計 100%，
    # 與 Andy 10-25 11:55「L3 口徑＝GB300、2026 年價格」及每 GW 44.9 同口徑）。v0.3.2–v0.5.1 為本專案自讀晶片內容橋接表的過渡值。
    # v0.2.3 歸塊與過渡值不同：運算＝NVIDIA 晶片毛利＋CPU（不含台積電與其他製造）；代工與封測、載板（ABF）各自成塊並計入 100%；儲存與管理另列。
    # 因此「代工與封測」由 L2（軸 B）移到 L3。價值鏈分量分沿用原值（以 VR200 2027 套級距，各塊級距不變；例外見依據欄）。
    ("compute", "L3 IT 設備", "運算", "{運算}", "運算", 1, 100,
     "AI半導體 次產業排序 v0.2.3：NVIDIA GPU 晶片毛利＋CPU；HBM、台積電、網通晶片不在此塊；只含 NVIDIA 平台 [Derived]"),
    ("memory", "L3 IT 設備", "記憶體", "{記憶體}", "記憶體", 1, 100,
     "AI半導體 次產業排序 v0.2.3：HBM＋LPDDR5X＋機架內 NAND（NAND 為 Assumed）；2027 HBM 合約價約翻倍 [Derived]"),
    ("network", "L3 IT 設備", "網通", "{網通}", "網通", 2, 75,
     "AI半導體 次產業排序 v0.2.3：NVLink 交換晶片＋交換匣其他＋scale-out 機架外＋SuperNIC／DPU [Derived]"),
    ("foundry", "L3 IT 設備", "代工與封測", "{代工與封測}", "代工與封測", 1, 75,
     "AI半導體 次產業排序 v0.2.3：台積電（GPU 晶片＋CoWoS）＋其他製造（測試、良率損失），扣除 ABF；比例小但先進製程與 CoWoS 近獨占，分量分例外給 75 [Derived]／[判斷]"),
    ("substrate", "L3 IT 設備", "載板與 PCB", "{載板與 PCB}|ABF，0.2–0.4%；PCB／CCL 在伺服器組裝內", "載板與 PCB", 2, 25,
     "AI半導體 次產業排序 v0.2.3：ABF 自「其他製造」切出 [Assumed] 區間"),
    ("power", "L3 IT 設備", "電源與散熱", "{電源與散熱}|含機構、整合；電源＋散熱約 3–5%", "電源與散熱", 2, 25,
     "AI半導體 次產業排序 v0.2.3：非 NVIDIA 機架內容全額；電源＋散熱 3–5% 為判斷，級距依 3–5% 給 25 [Derived]／[Assumed]"),
    ("server", "L3 IT 設備", "伺服器組裝", "{伺服器組裝}|殘差，未拆；ODM 附加價值約 1–3%", "伺服器組裝", 3, 25,
     "AI半導體 次產業排序 v0.2.3：橋接表殘差（匣級電路板含 PCB、匣級供電、NVIDIA／ODM 匣級加價），未拆；組裝廠附加價值 1–3% 為判斷 [Derived]"),
    # v0.5.3（2026-10-27）：儲存與管理 0.8% 改為地圖上的一塊（原只寫在說明句、不顯示，地圖各塊加總 31.9 ≠ IT 32.2；Andy 10-27 發現）。
    ("storage", "L3 IT 設備", "儲存與管理", "{儲存與管理}|機架外儲存、管理網路，未拆", "儲存與管理", 3, 25,
     "AI半導體 次產業排序 v0.2.3：橋接表儲存與管理列（機架外儲存、管理網路），未拆、不建模；不可替代性 3 為判斷 [Derived]／[判斷]"),
    ("equipment", "L2 半導體製造與設備", "設備與材料", "約 7.5%（軸 B，不可與軸 A 相加）", 0.075, 1, 50, "AI半導體 次產業排序 v0.2.3（設備模型情境 A：2027 AI 相關設備 117.8B ÷ 1,565B）"),
]


def _seg_shares():
    """讀 AI半導體 次產業排序最新版的 GB300 欄與 2027（VR200）欄；回傳 {次產業: (GB300 比例, 'VR200 文字')}。"""
    out = {}
    with open(latest("segments_ai_semis_v*.csv"), encoding="utf-8") as f:
        for r in csv.DictReader(f):
            g = r.get("GB300 對照 ÷ IT", "").strip()
            if g.endswith("%"):
                out[r["次產業"]] = (float(g[:-1]) / 100, r["2027 金額 ÷ AI IT capex"].strip())
    return out


_SH = _seg_shares()
SEGMENTS = [
    (k, ly, nm, (f"{_SH[mid][0]:.1%}（VR200 2027 {_SH[mid][1]}" + ("；" + txt.split("|", 1)[1] if "|" in txt else "") + "）"
                 if isinstance(mid, str) else txt),
     (_SH[mid][0] if isinstance(mid, str) else mid), irr, vc, b)
    for (k, ly, nm, txt, mid, irr, vc, b) in SEGMENTS
]
assert abs(sum(v for (k, ly, nm, t, v, *_ ) in SEGMENTS if ly == "L3 IT 設備") - 1) < 0.002, "L3 各塊（含儲存與管理）應合計 100%"
SEG_BY_NAME = {s[2]: s for s in SEGMENTS}

# 研究覆蓋公司（有模型者）的地位判斷：代號 → (次產業, 公司重要性 1–3, 市場位置 1–4, 地位依據, 標記)
# 市場位置：1 龍頭／2 前二或寡占／3 主要參與者／4 追隨者（在所屬次產業內）
POSITION = {
    "NVDA": ("運算", 1, 1, "AI 加速器金額份額 CY2027 約 75–82%（本專案推導）", "[D]"),
    "AVGO": ("運算", 1, 2, "客製 ASIC 約 80–85%（JPMorgan）；運算整體第二", "[IP]"),
    "000660": ("記憶體", 1, 1, "2026 HBM 市占 52–55%（本專案）", "[D]"),
    "005930": ("記憶體", 1, 2, "2026 HBM 市占 28–32%（本專案）", "[D]"),
    "MU": ("記憶體", 1, 3, "2026 HBM 市占 18–20%（本專案）", "[D]"),
    "ANET": ("網通", 1, 2, "AI 後端乙太網路交換器份額約 18.7%（本專案 Arista 比對）；NVIDIA、白牌競爭", "[D]"),
    "300308": ("網通", 1, 1, "高速光模組出貨領先（媒體與券商）", "[IP]"),
    "2330": ("代工與封測", 1, 1, "先進製程晶圓與 CoWoS 接近獨占", "[判斷]"),
    "4062": ("載板與 PCB", 1, 1, "高階 AI 載板主要供應者", "[判斷]"),
    "2308": ("電源與散熱", 1, 1, "伺服器電源龍頭", "[判斷]"),
    "2382": ("伺服器組裝", 1, 2, "雲端與 NVIDIA 整櫃主要 ODM；鴻海為最大組裝廠", "[判斷]"),
    "6857": ("設備與材料", 1, 1, "SoC 測試市占 66%（公司自估 CY2025）", "[IP]"),
    "6146": ("設備與材料", 2, 1, "研削與切割市占 60–80%（兩來源）", "[IP]"),
    "AMZN": ("AI 雲端平台", 1, 1, "AWS 為最大公有雲；自研 Trainium", "[判斷]"),
    "MSFT": ("AI 雲端平台", 1, 2, "Azure 第二大公有雲；OpenAI 主要算力夥伴", "[判斷]"),
    "GOOGL": ("AI 雲端平台", 1, 2, "TPU＋Gemini 垂直整合；雲端第三", "[判斷]"),
    "ORCL": ("AI 雲端平台", 2, 3, "OCI 規模較小；Stargate 主要承建方，客戶集中", "[判斷]"),
    "CRWV": ("Neocloud（GPU 租賃）", 2, 1, "最大 neocloud；客戶集中於 Microsoft、OpenAI", "[判斷]"),
    "NBIS": ("Neocloud（GPU 租賃）", 2, 2, "歐洲為主的 neocloud；Microsoft 長約", "[判斷]"),
    "WYFI": ("Neocloud（GPU 租賃）", 3, 4, "小型 neocloud＋託管", "[判斷]"),
    "OpenAI": ("AI 模型商", 1, 1, "ChatGPT 與 API 用量最大的模型商；算力承諾最大（本專案 OpenAI 模型）", "[判斷]"),
    "Anthropic": ("AI 模型商", 1, 2, "企業與程式開發用途為主的前沿模型商（本專案 Anthropic 模型）", "[判斷]"),
    "00100": ("AI 模型商", 3, 3, "中國第二梯隊模型商；多模態、海外營收占比高", "[判斷]"),
    "02513": ("AI 模型商", 3, 3, "中國第二梯隊模型商；政企本地化部署", "[判斷]"),
}

COMPANY_META = {
    # 代號：(公司, 掛牌, 幣別, 模型資料夾)
    "NVDA": ("NVIDIA", "美", "USD", "nvidia"), "AVGO": ("Broadcom", "美", "USD", "avgo"),
    "000660": ("SK 海力士", "韓", "KRW", "skhynix"), "005930": ("三星電子", "韓", "KRW", "samsung"),
    "MU": ("美光", "美", "USD", "micron"), "2330": ("台積電", "台", "TWD", "tsmc"),
    "ANET": ("Arista", "美", "USD", "arista"), "300308": ("中際旭創", "中", "CNY", "innolight"),
    "2308": ("台達電", "台", "TWD", "delta"), "2382": ("廣達", "台", "TWD", "quanta"),
    "4062": ("Ibiden", "日", "JPY", "ibiden"), "6857": ("Advantest", "日", "JPY", "advantest"),
    "6146": ("Disco", "日", "JPY", "disco"),
    "MSFT": ("Microsoft", "美", "USD", "microsoft"), "AMZN": ("Amazon", "美", "USD", "amazon"),
    "GOOGL": ("Alphabet", "美", "USD", "alphabet"), "ORCL": ("Oracle", "美", "USD", "oracle"),
    "CRWV": ("CoreWeave", "美", "USD", "coreweave"), "NBIS": ("Nebius", "美", "USD", "nebius"),
    "WYFI": ("WhiteFiber", "美", "USD", "whitefiber"),
    "00100": ("MiniMax", "港", "HKD", "minimax"), "02513": ("智譜", "港", "HKD", "zhipu"),
    "OpenAI": ("OpenAI", "未上市", "USD", "openai"), "Anthropic": ("Anthropic", "未上市", "USD", "anthropic"),
}

# v0.5（2026-10-25 Andy 18:14）：AI半導體 其餘個股全部接入。公司名、掛牌、次產業、重要性、地位依據取自候選母體；
# 市場位置（1–4）預設由「次產業內優先序」換算（1→1、2→2、3–5→3、≥6→4），標 [判斷]，黃底可改。
CCY_BY_LISTING = {"美": "USD", "台": "TWD", "日": "JPY", "韓": "KRW", "中": "CNY", "荷": "EUR", "港": "HKD"}


def _add_semis():
    import extract_models
    with open(latest("universe_ai_semis_v*.csv"), encoding="utf-8") as f:
        uni = {r["代號"]: r for r in csv.DictReader(f)}
    for tk, folder in extract_models.SEMIS_V2:
        if tk in COMPANY_META:
            continue
        u = uni[tk]
        order = int(u["次產業內優先序"])
        posn = 1 if order == 1 else 2 if order == 2 else 3 if order <= 5 else 4
        COMPANY_META[tk] = (u["公司"], u["掛牌"], CCY_BY_LISTING[u["掛牌"]], folder)
        POSITION[tk] = (u["次產業"], int(u["供應鏈重要性（1 高–3 低）"]), posn,
                        f"{u['市場地位']}（AI半導體 候選母體 v0.2.2；市場位置由次產業內優先序 {order} 換算）", "[判斷]")


_add_semis()


# 未上市與算力層的地圖節點（不進排序）
MAP_EXTRA = []  # OpenAI、Anthropic 已納入 POSITION（2026-10-25 Andy：未上市也放，可投資性寫「未上市」）

PARAMS = [
    ("w_irr", 0.40, "地位分數權重：不可替代性"),
    ("w_pos", 0.35, "地位分數權重：市場位置"),
    ("w_vc", 0.25, "地位分數權重：價值鏈分量"),
    ("s_irr1", 100, "不可替代性／重要性 1 → 分數"), ("s_irr2", 60, "2 → 分數"), ("s_irr3", 30, "3 → 分數"),
    ("s_pos1", 100, "市場位置 1 龍頭 → 分數"), ("s_pos2", 75, "2 前二或寡占 → 分數"),
    ("s_pos3", 50, "3 主要參與者 → 分數"), ("s_pos4", 25, "4 追隨者 → 分數"),
]


# ---------------------------------------------------------------- 讀回既有黃底輸入
def restore():
    """讀回黃底判斷輸入：優先讀本版活頁簿；本版尚未產生（升版）時讀 dist/ 最新一版，避免升版時判斷被預設值覆蓋。"""
    src = OUT
    if not os.path.exists(src):
        prev = sorted(glob.glob(os.path.join(HERE, "dist", "*_AI投資地圖_資料_*.xlsx")))
        if not prev:
            return {}, {}, {}
        src = prev[-1]
        print("讀回黃底輸入：", os.path.basename(src))
    wb = openpyxl.load_workbook(src)
    p = {r[0].value: r[1].value for r in wb["參數"].iter_rows(min_row=2) if r[0].value}
    seg = {}
    for r in wb["價值鏈"].iter_rows(min_row=2):
        if r[1].value:
            seg[r[1].value] = (r[5].value, r[6].value)
    pos = {}
    for r in wb["地位"].iter_rows(min_row=2):
        if r[0].value:
            pos[str(r[0].value)] = (r[4].value, r[5].value)
    return p, seg, pos


def hdr(ws, cols, widths):
    for j, (c, w) in enumerate(zip(cols, widths), 1):
        cell = ws.cell(1, j, c)
        cell.fill, cell.font, cell.alignment = HDR, HFONT, WRAP
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "B2"


def src(d):
    if d is None:
        return ""
    if d.get("cell"):
        return f"{d['file']} {d['cell']}"
    return d.get("source", "")


def main():
    rp, rseg, rpos = restore()
    wb = openpyxl.Workbook()

    # ============ README
    ws = wb.active
    ws.title = "README"
    lines = [
        f"AI 投資地圖 資料活頁簿 {VERSION}（{DATE}）",
        "用途：AI 投資地圖網站（markets/invest_map/dist/index.html）的唯一數字來源。網站只呈現本活頁簿重算後的值。",
        "三個視角：(1) 供應鏈地圖＝「價值鏈」頁＋Tokenomics 每 GW 金額；(2) 產業地位＝「地位」頁；(3) 可投資性＝「估值」頁。",
        "黃底＝判斷輸入（Andy 可直接改；重建時讀回，不被程式覆寫）。藍字＝自各模型或 Tokenomics 抄錄的值，旁欄註明來源儲存格。其餘為公式。",
        f"Tokenomics：{TK['source']['file']}（commit {TK['source']['commit'][:7]}；擷取 {TK['source']['extractedAt']}）",
        "公司模型：company-models 各公司 dist/*.xlsx（extract_models.py 抽數）。",
        "AI半導體 候選母體 v0.2.2（2026-10-24）：次產業排序、公司重要性、市場地位文字。權威在 AI半導體 Project。",
        "收支關卡（Andy 2026-10-25 10:48）：先看收支、再看評價。模型期間（2026 起約五年）現金收支合計為負＝入不敷出、有資金缺口，"
        "評價為買進以上（或未估值）者，最終分級改為「收支不足」；評價為中立、放空、強力放空者維持原評價，入不敷出只加重放空理由（Andy 2026-10-26 修訂）。口徑：雲端／neocloud／MAG 模板一律用「營運缺口＋股東回饋（股利、回購）」；記憶體、模型商用自由現金流；"
        "只有損益的半導體模型以「各情境淨利皆為正」判定（現金流本專案未建立）。",
        "可投資性口徑（v0.5，Andy 2026-10-25 18:14：數字引用個股模型，模型沒有的才自算）：半導體 52 家的中心價＝模型「目標價（DCF 與 Comparable 加權）」、"
        "中心空間＝模型「上檔空間：加權目標價」、現價＝模型現價（上檔空間的分母）、投資建議與上檔空間級距＝模型 Val 列；模型尚未補列者依 AI半導體 固定規則"
        "（>+50% 強力買進、+25～+50% 買進、−25～+25% 中立、−50～−25% 放空、<−50% 強力放空）以模型上檔空間自算。悲觀情境空間（v0.5.1 前稱「下檔空間」）＝C 情境 DCF 價值 ÷ 模型現價 − 1（模型無此欄，自算）；不是一般所說目標價相對現價的下跌幅度。"
        "雲端與 neocloud（v0.5.1）：上檔空間＝模型加權目標價 ÷ 最近外部收盤價 − 1，依同一五級規則自算（不用模型自身評等，不再出現「賣出」）；MiniMax 用反向檢驗折現價值，依同一固定規則自算。",
        "本活頁簿是研究整理，不是投資建議；各模型的方法與假設見各自 README 與 Val／摘要頁。",
    ]
    for i, t in enumerate(lines, 1):
        ws.cell(i, 1, t).alignment = WRAP
    ws.column_dimensions["A"].width = 140

    # ============ 參數
    wp = wb.create_sheet("參數")
    hdr(wp, ["名稱", "值", "說明"], [16, 10, 50])
    for i, (k, v, d) in enumerate(PARAMS, 2):
        wp.cell(i, 1, k)
        c = wp.cell(i, 2, rp.get(k, v))
        c.fill = YELLOW
        wp.cell(i, 3, d)
        wb.defined_names[k] = openpyxl.workbook.defined_name.DefinedName(k, attr_text=f"參數!$B${i}")

    # ============ Tokenomics 每 GW
    wt = wb.create_sheet("每GW")
    hdr(wt, ["項目", "世代", "低", "基準", "高", "單位", "Tokenomics 名稱", "來源儲存格（基準）"],
        [34, 12, 10, 10, 10, 14, 26, 40])
    r = 2
    for name, label in (("IF_CapexIT", "每 GW 資本支出：IT 設備"), ("IF_CapexFacility", "每 GW 資本支出：設施（電力、機房、冷卻）"),
                        ("IF_CapexTotal", "每 GW 資本支出：合計"), ("IF_HoldEcon", "每 GW 年經濟持有成本")):
        for gen in ("GB300 NVL72", "VR200 NVL72"):
            it = TK["items"][name]
            vals = it["values"][gen]
            wt.cell(r, 1, label)
            wt.cell(r, 2, gen.split()[0])
            for j, case in enumerate(("低成本", "基準", "高成本"), 3):
                c = wt.cell(r, j, vals[case])
                c.font = BLUE
            wt.cell(r, 6, it["unit"] if name != "IF_HoldEcon" else "$B/GW/年")
            wt.cell(r, 7, name)
            wt.cell(r, 8, it["cells"][gen]["基準"])
            r += 1
    for name, label in (("L1_RevGW_PQ_VR200", "VR200 每 GW 營收 P×Q：第一年（m＝2026 Q2）"),
                        ("L1_RevGW_PQ_VR200_Life", "VR200 每 GW 營收 P×Q：壽命期平均（含 L）"),
                        ("L1_LTotalFleetVDisc", "實際 L（機隊、折現）"),
                        ("L1_LReqFleet_Full", "回本所需 L（全部算力持有成本口徑）")):
        v, cell = tkv(name)
        lo, _ = tkr(name, "低")
        hi, _ = tkr(name, "高")
        wt.cell(r, 1, label)
        wt.cell(r, 2, "VR200" if "VR200" in name else "機隊")
        for j, x in enumerate((lo, v, hi), 3):
            wt.cell(r, j, x).font = BLUE
        wt.cell(r, 6, TK["items"][name]["unit"])
        wt.cell(r, 7, name)
        wt.cell(r, 8, cell)
        r += 1
    wt.cell(r, 1, "回本比值（實際 L ÷ 所需 L）")
    wt.cell(r, 2, "機隊")
    wt.cell(r, 4, f"=D{r-2}/D{r-1}")
    wt.cell(r, 6, "倍")
    wt.cell(r, 7, "[Derived]")
    wt.cell(r, 8, "Tokenomics 交接：0.940，「未定」")
    TK_ROWS = r

    # ============ 價值鏈
    wv = wb.create_sheet("價值鏈")
    hdr(wv, ["層級", "次產業", "占 AI IT 資本支出", "比例中值", "每 GW 金額（GB300，$B）", "不可替代性（1–3）", "價值鏈分量分",
             "依據"], [14, 22, 30, 10, 14, 12, 12, 70])
    capex_it_row = 2  # 每GW!D2＝IF_CapexIT GB300 基準（D3 為 VR200）
    for i, s in enumerate(SEGMENTS, 2):
        key, layer, name, txt, mid, irr, vc, basis = s
        irr, vc = rseg.get(name, (irr, vc))
        wv.cell(i, 1, layer)
        wv.cell(i, 2, name)
        wv.cell(i, 3, txt)
        if mid is not None:
            wv.cell(i, 4, mid).font = BLUE
            if layer == "L3 IT 設備":
                wv.cell(i, 5, f"=D{i}*每GW!$D${capex_it_row}")
        wv.cell(i, 6, irr).fill = YELLOW
        wv.cell(i, 7, vc).fill = YELLOW
        wv.cell(i, 8, basis).alignment = WRAP

    # ============ 地位
    wl = wb.create_sheet("地位")
    hdr(wl, ["代號", "公司", "層級", "次產業", "公司重要性（1–3）", "市場位置（1–4）", "地位依據", "標記",
             "次產業不可替代性", "價值鏈分量分", "不可替代性分", "市場位置分", "地位分數", "排名"],
        [9, 12, 14, 20, 10, 10, 56, 8, 10, 10, 10, 10, 10, 7])
    n = len(POSITION)
    for i, (tk, (seg, imp, posn, basis, tag)) in enumerate(POSITION.items(), 2):
        imp, posn = rpos.get(tk, (imp, posn))
        segrow = [s[2] for s in SEGMENTS].index(seg) + 2
        wl.cell(i, 1, tk)
        wl.cell(i, 2, COMPANY_META[tk][0])
        wl.cell(i, 3, f"=價值鏈!A{segrow}")
        wl.cell(i, 4, seg)
        wl.cell(i, 5, imp).fill = YELLOW
        wl.cell(i, 6, posn).fill = YELLOW
        wl.cell(i, 7, basis).alignment = WRAP
        wl.cell(i, 8, tag)
        wl.cell(i, 9, f"=價值鏈!F{segrow}")
        wl.cell(i, 10, f"=價值鏈!G{segrow}")
        wl.cell(i, 11, f"=(CHOOSE(I{i},s_irr1,s_irr2,s_irr3)+CHOOSE(E{i},s_irr1,s_irr2,s_irr3))/2")
        wl.cell(i, 12, f"=CHOOSE(F{i},s_pos1,s_pos2,s_pos3,s_pos4)")
        wl.cell(i, 13, f"=w_irr*K{i}+w_pos*L{i}+w_vc*J{i}")
        wl.cell(i, 14, f"=RANK(M{i},$M$2:$M${n+1},0)")

    # ============ 價格
    wpz = wb.create_sheet("價格")
    hdr(wpz, ["代號", "幣別", "模型內現價", "模型價日期", "模型價來源", "外部收盤（非半導體模型）", "收盤來源", "採用價"],
        [9, 7, 12, 12, 52, 12, 52, 12])
    prices = {}
    pfile = os.path.join(HERE, "data", "prices.csv")
    if os.path.exists(pfile):
        with open(pfile, encoding="utf-8") as f:
            for row in csv.DictReader(f):
                prices[row["ticker"]] = row
    PRICE_ROW = {}
    for i, tk in enumerate(COMPANY_META, 2):
        m = MX[tk]
        PRICE_ROW[tk] = i
        wpz.cell(i, 1, tk)
        wpz.cell(i, 2, COMPANY_META[tk][2])
        mp, mdate, msrc = None, "", ""
        if "model_px" in m:
            mp, mdate, msrc = m["model_px"]["value"], m["model_px"]["date"], src(m["model_px"])
        elif "price" in m and isinstance(m["price"], dict):
            mp, msrc = m["price"]["value"], src(m["price"])
            mdate = m.get("price_date", "")
        if mp is not None:
            wpz.cell(i, 3, mp).font = BLUE
        wpz.cell(i, 4, mdate)
        wpz.cell(i, 5, msrc)
        if tk in prices and prices[tk]["close"]:
            wpz.cell(i, 6, float(prices[tk]["close"])).font = BLUE
            wpz.cell(i, 7, f"{prices[tk]['date']}｜{prices[tk]['source']}")
        if m.get("unlisted"):
            wpz.cell(i, 8, "未上市")
        elif m.get("v2"):   # v0.5（Andy 18:14：數字引用個股模型）：半導體用模型現價，即模型上檔空間的分母
            wpz.cell(i, 8, f"=C{i}")
        else:
            wpz.cell(i, 8, f"=IF(ISNUMBER(F{i}),F{i},C{i})")

    # ============ 估值
    we = wb.create_sheet("估值")
    cols = ["代號", "公司", "層級", "次產業", "方法", "中心候選 1", "中心候選 1 來源", "中心候選 2", "中心候選 2 來源",
            "悲觀情境候選 1", "悲觀情境候選 1 來源", "悲觀情境候選 2", "悲觀情境候選 2 來源", "採用價", "中心價", "悲觀情境價",
            "中心空間", "悲觀情境空間", "分級", "對照口徑空間", "對照口徑", "現價 ÷ 模型 FY2028 EPS（A／B／C）", "註記",
            "輔助 1", "輔助 2", "模型價日以來漲跌", "價格變動旗標",
            "收支口徑", "收支列（模型）", "收支序列", "期間合計", "最小值", "收支來源", "收支關卡", "最終分級",
            "上檔空間級距", "投資建議與上檔空間依據"]
    hdr(we, cols, [9, 12, 14, 18, 24, 12, 40, 12, 40, 12, 40, 12, 40, 11, 11, 11, 9, 9, 8, 10, 30, 22, 60, 10, 10, 10, 18, 9, 30, 46, 10, 10, 44, 14, 12, 12, 60])
    REC = 'IF(Q{i}>0.5,"強力買進",IF(Q{i}>0.25,"買進",IF(Q{i}>=-0.25,"中立",IF(Q{i}>=-0.5,"放空","強力放空"))))'
    TIER6 = ('IF(Q{i}>0.5,"大於 50%",IF(Q{i}>0.25,"25%–50%",IF(Q{i}>=0,"0%–25%",'
             'IF(Q{i}>=-0.25,"0%～−25%",IF(Q{i}>=-0.5,"−25%～−50%","小於 −50%")))))')

    def put(i, j, d, fmt_src=True):
        if d is None:
            return
        if isinstance(d, dict):
            we.cell(i, j, d["value"]).font = BLUE
            we.cell(i, j + 1, src(d))
        else:  # (公式, 說明)
            we.cell(i, j, d[0])
            we.cell(i, j + 1, d[1])

    for i, tk in enumerate(COMPANY_META, 2):
        m = MX[tk]
        seg = POSITION[tk][0]
        segrow = [s[2] for s in SEGMENTS].index(seg) + 2
        prow = PRICE_ROW[tk]
        we.cell(i, 1, tk)
        we.cell(i, 2, COMPANY_META[tk][0])
        we.cell(i, 3, f"=價值鏈!A{segrow}")
        we.cell(i, 4, seg)
        we.cell(i, 14, f"=價格!H{prow}")
        method, c1, c2, b1, b2, alt, altlbl, note = "", None, None, None, None, None, "", ""
        if m.get("v2"):
            # 估值 v2（地圖 v0.4，2026-10-25）：中心＝模型「目標價（DCF 與 Comparable 加權）」；
            # 下檔＝C 情境 DCF 價值（推到 12 個月後，即價格帶下緣）。對照欄＝模型內的舊法規則值情境加權價（參考）。
            method = "加權目標價（DCF 50%＋分部加總 50%）" if m.get("seg_model") else "加權目標價（DCF 50%＋Comparable 50%）"
            c1 = m["tp"]
            b1 = m["dcf_C"]
            if "old_rule" in m:
                alt = (f"={m['old_rule']['value']}/N{i}-1", )
                altlbl = f"舊法：規則值情境加權價（參考，{m['old_rule']['cell']}）"
            note = (f"DCF 目標價 {m['tp_dcf']['value']:,.1f}（{m['tp_dcf']['cell']}）；"
                    f"Comparable 目標價 {m['tp_comp']['value']:,.1f}（{m['tp_comp']['cell']}）；"
                    f"市場共識目標價 {m['tp_cons']['value']:,.1f}（{m['tp_cons']['cell']}，外部、僅供對照）；"
                    f"價格帶 C～A 情境 DCF {m['dcf_C']['value']:,.1f}～{m['dcf_A']['value']:,.1f}")
        elif tk in ("MSFT", "AMZN", "GOOGL", "ORCL", "CRWV", "NBIS", "WYFI"):
            method = "模型加權目標價"   # v0.5.1（Andy 2026-10-26）：不顯示模型自身評等（含「賣出」），改用本地圖五級
            c1 = m["target"]
            b1 = m["target"]
            b2 = m["scen_lo"]
            # v0.5.1：模型摘要句含模型自身評等與門檻（「賣出」等），網站不顯示；只留點位與區間，評等改依本地圖五級
            parts = m["sentence"]["value"].rstrip("。").split("；")
            head = re.sub(r"^[^：]*：點位 (\S+)，較現價 (\S+) .*$", r"模型點位 \1（模型日現價 \2）", parts[0])
            note = "；".join([head] + [x for x in parts[1:] if "門檻" not in x])
            if tk in ("MSFT", "AMZN", "GOOGL"):
                note += "｜MAG 模型終值占 EV 逾 90%，評價口徑待 MAG v0.3"
        elif tk == "00100":
            method = "反向檢驗：折現價值 ÷ 市值"
            we.cell(i, 24, m["disc_ratio"]["value"]).font = BLUE
            c1 = (f"=X{i}*價格!C{prow}", f"{src(m['disc_ratio'])} × 模型現價")
            b1 = (f"=X{i}*價格!C{prow}", "同中心（模型只有一個情境）")
        elif m.get("unlisted"):
            method = "未上市"
            note = "未上市：不評估可投資性；收支關卡結果只作參考。"
        elif tk == "02513":
            method = "未估值（反向模式）"
            note = (f"模型未出隱含價；2028 年共識營收需為正向推估的 {m['need_ratio_2028']['value']:.2f} 倍"
                    f"（最低共識 {m['need_ratio_2028_lo']['value']:.2f} 倍；{src(m['need_ratio_2028'])}）")
        we.cell(i, 5, method)
        put(i, 6, c1)
        put(i, 8, c2)
        put(i, 10, b1)
        put(i, 12, b2)
        if c1 is not None:
            we.cell(i, 15, f"=MIN(F{i},H{i})")
            we.cell(i, 16, f"=MIN(J{i},L{i})")
            we.cell(i, 18, f"=P{i}/N{i}-1")          # 悲觀情境空間（v0.5.1 前稱「下檔空間」）：模型無此欄，自算（C 情境 DCF ÷ 模型現價 − 1）
            if m.get("v2"):
                # v0.5（Andy 18:14）：模型有的數字直接引用——上檔空間、投資建議、級距；模型尚未補列者依同一固定規則自算
                we.cell(i, 17, m["up"]["value"]).font = BLUE
                basis = [f"上檔空間 {src(m['up'])}"]
                for col, key, f in ((19, "rec", REC), (36, "tier6", TIER6)):
                    if key in m:
                        we.cell(i, col, m[key]["value"]).font = BLUE
                        basis.append(f"{'投資建議' if key == 'rec' else '級距'} {src(m[key])}")
                    else:
                        we.cell(i, col, "=" + f.format(i=i))
                        basis.append(f"{'投資建議' if key == 'rec' else '級距'}：模型尚未補列，依 AI半導體 固定規則自算")
                we.cell(i, 37, "；".join(basis)).alignment = WRAP
            elif "rating" in m:                       # 雲端、neocloud、MAG
                # v0.5.1（Andy 2026-10-26）：「賣出」與放空易混，取消；模型未更新前，以模型加權目標價 ÷ 最近收盤價自算上檔空間，
                # 依本地圖五級規則評等。模型自身評等（以模型日價計）只留在依據欄供對照，網站不顯示。
                we.cell(i, 17, f"=O{i}/N{i}-1")
                we.cell(i, 19, "=" + REC.format(i=i))
                we.cell(i, 36, "=" + TIER6.format(i=i))
                we.cell(i, 37, f"上檔空間＝模型加權目標價 {src(m['target'])} ÷ 外部收盤價 − 1，依本地圖五級規則自算"
                               "（模型自身評等以模型日價計、級距不同，不採用；原值見模型摘要頁）")
            else:                                     # MiniMax：模型無評等，依同一固定規則自算
                we.cell(i, 17, f"=O{i}/N{i}-1")
                we.cell(i, 19, "=" + REC.format(i=i))
                we.cell(i, 36, "=" + TIER6.format(i=i))
                we.cell(i, 37, "模型無評等；依 AI半導體 固定規則自算")
        else:
            we.cell(i, 19, "未上市" if m.get("unlisted") else "未估值")
        if alt:
            we.cell(i, 20, alt[0])
            we.cell(i, 21, altlbl)
        pe = m.get("pe_fy28")
        if pe:
            we.cell(i, 22, " ／ ".join(f"{x['value']:.1f}" for x in pe))
        we.cell(i, 23, note).alignment = WRAP
        we.cell(i, 26, f'=IF(ISNUMBER(價格!C{prow}),價格!H{prow}/價格!C{prow}-1,"")')
        we.cell(i, 27, f'=IF(Z{i}="","",IF(ABS(Z{i})>0.2,"模型尚未反映期間變化",""))')
        # 收支關卡（Andy 2026-10-25 10:48）：入不敷出（模型期間現金收支合計為負）者不得列為可投資
        cash = m["cash"]
        vals = [v for _, v, _ in cash["values"]]
        cells = [c for _, _, c in cash["values"]]
        we.cell(i, 28, cash["basis"])
        we.cell(i, 29, cash["label"])
        if cash["basis"] == "現金流":
            we.cell(i, 30, "、".join(f"{y} {v:,.1f}" for y, v, _ in cash["values"]))
        else:
            we.cell(i, 30, f"{len(vals)} 個情境年度淨利（不含實績），全部{'為正' if min(vals) > 0 else '含負值'}")
        we.cell(i, 31, sum(vals)).font = BLUE
        we.cell(i, 32, min(vals)).font = BLUE
        we.cell(i, 33, f"[Derived] {cash['file']} {cells[0]}…{cells[-1]}")
        we.cell(i, 34, f'=IF(AB{i}="淨利",IF(AF{i}>0,"通過（淨利）","不通過"),IF(AE{i}<0,"不通過","通過（現金流）"))')
        # v0.5.1（Andy 2026-10-26）：收支不足只壓過「買進以上」（與未估值）；放空、強力放空、中立維持原評價——現金流缺口加重放空理由
        we.cell(i, 35, f'=IF(S{i}="未上市","未上市",IF(AND(AH{i}="不通過",OR(S{i}="強力買進",S{i}="買進",S{i}="未估值")),"收支不足",S{i}))')

    # ============ 地圖節點（全部候選＋算力層）
    wm = wb.create_sheet("地圖節點")
    hdr(wm, ["次產業", "次產業內優先序", "公司", "代號", "掛牌", "角色", "供應鏈重要性（1–3）", "市場地位", "標記",
             "研究狀態"], [14, 8, 16, 10, 6, 40, 10, 50, 10, 30])
    r = 2
    with open(latest("universe_ai_semis_v*.csv"), encoding="utf-8") as f:
        for row in csv.DictReader(f):
            tk = row["代號"]
            status = "有模型" if tk in COMPANY_META else row["本專案狀態"]
            vals = [row["次產業"], row["次產業內優先序"], row["公司"], tk, row["掛牌"], row["在 AI 供應鏈的角色"],
                    row["供應鏈重要性（1 高–3 低）"], row["市場地位"], row["標記"], status]
            for j, v in enumerate(vals, 1):
                wm.cell(r, j, v)
            r += 1
    for tk, (seg, imp, posn, basis, tag) in POSITION.items():
        if seg in ("AI 雲端平台", "Neocloud（GPU 租賃）", "AI 模型商"):
            vals = [seg, "", COMPANY_META[tk][0], tk, COMPANY_META[tk][1], "", imp, basis, tag, "有模型"]
            for j, v in enumerate(vals, 1):
                wm.cell(r, j, v)
            r += 1
    for name, listing, seg, status in MAP_EXTRA:
        for j, v in enumerate([seg, "", name, "", listing, "", "", "", "", status], 1):
            wm.cell(r, j, v)
        r += 1

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)
    print("寫出", OUT)


if __name__ == "__main__":
    main()
