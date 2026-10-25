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

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
VERSION = "v0.3.2"
DATE = "20261025"
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
    # L3 比例（v0.3.2，Andy 10-25 11:46 發現重疊）：改讀 晶片內容橋接表 v0.2.1 Bridge GB300 欄（與每 GW 44.9＝IT 32.2＋廠房 12.67 同口徑，六塊＋儲存合計 100%）。
    # 原 AI半導體 次產業排序 v0.2.2 的 43／23／19%：HBM 在運算與記憶體重複、HBM 價格期間混用、網通取自另一模型，合計約 91%（已發交流）。
    # 價值鏈分量分沿用原值：以 VR200 2027（AI半導體 原口徑年份）套級距，各塊級距不變（見依據欄）。
    ("compute", "L3 IT 設備", "運算", "57.3%（VR200 2027 約 37%）", 0.573, 1, 100,
     "晶片內容橋接表 v0.2.1 Bridge GB300 欄 B10＋B11＋B13＋B14（NVIDIA 晶片毛利 48.5%＋台積電 3.5%＋其他製造 2.3%＋CPU 3.0%）÷ B7；HBM 移到記憶體 [Derived]"),
    ("memory", "L3 IT 設備", "記憶體", "13.4%（VR200 2027 約 23%）", 0.134, 1, 100,
     "晶片內容橋接表 v0.2.1 Bridge GB300 欄 B12＋B15＋B16（HBM 6.7%＋LPDDR 4.6%＋機架內 SSD 2.1%）÷ B7；2027 HBM 合約價約翻倍，VR200 2027 約 23%，級距沿用 100 [Derived]"),
    ("network", "L3 IT 設備", "網通", "17.5%（VR200 約 19%）", 0.175, 2, 75,
     "晶片內容橋接表 v0.2.1 Bridge GB300 欄 B17＋B20（網卡與 DPU 4.0%＋NVLink 4.5%＋機架外交換器、光模組、線纜 9.0%）÷ B7 [Derived]"),
    ("substrate", "L3 IT 設備", "載板與 PCB", "含在運算「其他製造」2.3% 內（ABF 約 0.3%），不另加", 0.003, 2, 25,
     "AI半導體 次產業排序 v0.2.2（Ibiden 比對）；屬 晶片內容橋接表 v0.2.1 Bridge GB300 欄 B13 的一部分，金額不與運算相加 [Assumed] 區間"),
    ("power", "L3 IT 設備", "電源與散熱", "6.3%（電源、散熱、機構、整合；電源＋散熱約 3–5%）", 0.063, 2, 25,
     "晶片內容橋接表 v0.2.1 Bridge GB300 欄 B18 ÷ B7；電源＋散熱 3–5% 為 AI半導體 判斷，級距依 3–5% 給 25 [Derived]／[Assumed]"),
    ("server", "L3 IT 設備", "伺服器組裝", "4.6%（板卡、供電與組裝加價；ODM 附加價值約 1–3%）", 0.046, 3, 25,
     "晶片內容橋接表 v0.2.1 Bridge GB300 欄 B19 ÷ B7（殘差：匣級電路板、供電、NVIDIA／ODM 匣級加價）；ODM 附加價值 1–3% 為 AI半導體 判斷 [Derived]"),
    ("foundry", "L2 半導體製造與設備", "代工與封測", "約 3.5%（含在 L3 運算內，不另加；僅 NVIDIA 平台）", 0.035, 1, 75,
     "例外：晶片內容橋接表 v0.2.1 Bridge GB300 欄 B11 ÷ B7 只見 3.5%，但先進製程與 CoWoS 的價值藏在運算的晶片成本內，給 75 [判斷]"),
    ("equipment", "L2 半導體製造與設備", "設備與材料", "約 7.5%（軸 B，不可與軸 A 相加）", 0.075, 1, 50, "AI半導體 次產業排序 v0.2.2（設備模型情境 A：2027 AI 相關設備 117.8B ÷ 1,565B）"),
]
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

# 未上市與算力層的地圖節點（不進排序）
MAP_EXTRA = []  # OpenAI、Anthropic 已納入 POSITION（2026-10-25 Andy：未上市也放，可投資性寫「未上市」）

PARAMS = [
    ("tier_center", 0.20, "「機會」：中心空間至少"),
    ("tier_bear", -0.25, "「機會」：下檔空間不低於"),
    ("tier_rich", -0.10, "「偏貴」：中心空間低於"),
    ("w_irr", 0.40, "地位分數權重：不可替代性"),
    ("w_pos", 0.35, "地位分數權重：市場位置"),
    ("w_vc", 0.25, "地位分數權重：價值鏈分量"),
    ("s_irr1", 100, "不可替代性／重要性 1 → 分數"), ("s_irr2", 60, "2 → 分數"), ("s_irr3", 30, "3 → 分數"),
    ("s_pos1", 100, "市場位置 1 龍頭 → 分數"), ("s_pos2", 75, "2 前二或寡占 → 分數"),
    ("s_pos3", 50, "3 主要參與者 → 分數"), ("s_pos4", 25, "4 追隨者 → 分數"),
]


# ---------------------------------------------------------------- 讀回既有黃底輸入
def restore():
    if not os.path.exists(OUT):
        return {}, {}, {}
    wb = openpyxl.load_workbook(OUT)
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
        "不論估值空間多大，最終分級一律為「收支不足」。口徑：雲端／neocloud／MAG 模板一律用「營運缺口＋股東回饋（股利、回購）」；記憶體、模型商用自由現金流；"
        "只有損益的半導體模型以「各情境淨利皆為正」判定（現金流本專案未建立）。",
        "可投資性口徑（預設，Andy 可改）：中心價＝各模型主口徑；半導體 P/E 類取「十年倍數規則值」與「Andy 先驗合理 P/E」兩者較低者；"
        "規則值不適用者（NVDA、AVGO、Ibiden）只用 Andy 合理 P/E，標「替代口徑」；記憶體以 P/B 規則值為主、P/E 規則值為對照；"
        "雲端與 neocloud 用模型加權目標價與情境下緣；MiniMax 用反向檢驗折現價值。空間＝隱含價 ÷ 10-23 收盤價 − 1，未折現。",
        "分級：中心空間 ≥ 參數 tier_center 且下檔 ≥ tier_bear → 機會；中心 < tier_rich → 偏貴；其餘 → 觀察。",
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
    hdr(wpz, ["代號", "幣別", "模型內現價", "模型價日期", "模型價來源", "10-23 收盤", "收盤來源", "採用價"],
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
        wpz.cell(i, 8, "未上市" if m.get("unlisted") else f"=IF(ISNUMBER(F{i}),F{i},C{i})")

    # ============ 估值
    we = wb.create_sheet("估值")
    cols = ["代號", "公司", "層級", "次產業", "方法", "中心候選 1", "中心候選 1 來源", "中心候選 2", "中心候選 2 來源",
            "下檔候選 1", "下檔候選 1 來源", "下檔候選 2", "下檔候選 2 來源", "採用價", "中心價", "下檔價",
            "中心空間", "下檔空間", "分級", "對照口徑空間", "對照口徑", "現價 ÷ 模型 FY2028 EPS（A／B／C）", "註記",
            "輔助 1", "輔助 2", "模型價日以來漲跌", "價格變動旗標",
            "收支口徑", "收支列（模型）", "收支序列", "期間合計", "最小值", "收支來源", "收支關卡", "最終分級"]
    hdr(we, cols, [9, 12, 14, 18, 24, 12, 40, 12, 40, 12, 40, 12, 40, 11, 11, 11, 9, 9, 8, 10, 30, 22, 60, 10, 10, 10, 18, 9, 30, 46, 10, 10, 44, 14, 12])

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
        if tk == "NVDA":
            method = "替代口徑：Andy 合理 P/E"
            we.cell(i, 24, m["eps_fy28_A"]["value"]).font = BLUE
            we.cell(i, 25, m["eps_fy29_C"]["value"]).font = BLUE
            c1 = (f"=25*X{i}", f"Andy 先驗 25x × FY2028 EPS（A）{src(m['eps_fy28_A'])}")
            b1 = (f"=25*Y{i}", f"25x × FY2029 EPS（C）{src(m['eps_fy29_C'])}")
            note = m["rule_note"] + "；先驗 25x 為近 12 月口徑，套在前瞻 EPS 偏保守"
        elif tk in ("AVGO", "4062"):
            method = "替代口徑：Andy 合理 P/E"
            c1 = m["andy27_A"]
            b1 = m["andy28_C"]
            note = m["rule_note"]
        elif tk in ("ANET", "2308", "300308", "2382"):
            method = "P/E：規則值與 Andy 合理 P/E 取低"
            c1, c2 = m["rule_A"], m["andy27_A"]
            b1, b2 = m["rule_C"], m["andy28_C"]
            if tk == "300308":
                note = "十年 P/E 中位數 52.75x 含小型股時期，規則值偏高；取低後由 Andy 30x 決定"
        elif tk == "2330":
            method = "P/E：規則值與 Andy 合理 P/E 取低"
            we.cell(i, 24, m["eps27_A"]["value"]).font = BLUE
            we.cell(i, 25, m["eps28_C"]["value"]).font = BLUE
            c1 = m["rule_mid"]
            c2 = (f"=35*X{i}", f"Andy 先驗 35x × 2027 EPS（A）{src(m['eps27_A'])}")
            b1 = m["rule_lo"]
            b2 = (f"=35*Y{i}", f"35x × 2028 EPS（C）{src(m['eps28_C'])}")
            note = "Andy 35x 高於十年最高值（先驗比對 T7b），取低後由規則值決定"
        elif tk in ("000660", "005930", "MU"):
            method = "P/B 規則值（主）"
            c1 = m["pb_rule"]["見頂"]
            b1 = m["pb_rule"]["見頂"]
            b2 = m["pb_rule"]["谷底"]
            alt = (f"={m['pe_rule']['見頂']['value']}/N{i}-1", )
            altlbl = f"P/E 規則值：見頂（{m['pe_rule']['見頂']['cell']}）"
            note = "方法衝突：P/B 假設 ROE 回到歷史；P/E 假設高峰盈餘以歷史倍數計價（overview README v0.3）"
        elif tk in ("6857", "6146"):
            method = "P/E 規則值（設備股共用循環規則）"
            c1 = m["rule_A"]
            b1 = m["rule_C"]
        elif tk in ("MSFT", "AMZN", "GOOGL", "ORCL", "CRWV", "NBIS", "WYFI"):
            method = f"模型加權目標價（評等：{m['rating']['value']}）"
            c1 = m["target"]
            b1 = m["target"]
            b2 = m["scen_lo"]
            note = m["sentence"]["value"]
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
            we.cell(i, 17, f"=O{i}/N{i}-1")
            we.cell(i, 18, f"=P{i}/N{i}-1")
            we.cell(i, 19, f'=IF(AND(Q{i}>=tier_center,R{i}>=tier_bear),"機會",IF(Q{i}<tier_rich,"偏貴","觀察"))')
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
        we.cell(i, 35, f'=IF(S{i}="未上市","未上市",IF(AH{i}="不通過","收支不足",S{i}))')

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
