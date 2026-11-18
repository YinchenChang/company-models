import json, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L

B = "/home/claude/out/channel/"
us = json.load(open(B + "data/us_data.json"))
tw = json.load(open(B + "data/tw_quarterly.json"))
wy = json.load(open(B + "data/wiwynn.json"))
mo = json.load(open(B + "data/monthly.json"))
# 修正：Meta 2025Q3 未起租租賃取 2025-09-30 的值（避免被 2025-10-31 的其他事實覆蓋）
us["META"]["lease_nc"]["2025Q3"] = {"v": 58.14, "end": "2025-09-30", "acc": "0001628280-25-047240", "form": "10-Q", "filed": "2025-10-30", "how": "instant"}

QS = [f"{y}Q{q}" for y in (2024, 2025, 2026) for q in (1, 2, 3, 4)][:-1]
def qend(q):
    y, n = int(q[:4]), int(q[5])
    return dt.date(y, n * 3, [31, 30, 30, 31][n - 1])
def qdays(q):
    e = qend(q); n = int(q[5])
    s = dt.date(e.year, 1, 1) if n == 1 else qend(f"{e.year}Q{n-1}") + dt.timedelta(days=1)
    return (e - s).days + 1

HDR = PatternFill("solid", fgColor="1F3864"); HF = Font(bold=True, color="FFFFFF")
INP = Font(color="0000FF"); WRAP = Alignment(wrap_text=True, vertical="top")
GREY = PatternFill("solid", fgColor="F2F2F2")
def hdr(ws, row, heads, widths=None):
    for i, h in enumerate(heads, 1):
        c = ws.cell(row, i, h); c.fill = HDR; c.font = HF; c.alignment = Alignment(wrap_text=True, vertical="center")
    if widths:
        for i, w in enumerate(widths, 1): ws.column_dimensions[L(i)].width = w
    ws.row_dimensions[row].height = 45

# ---------------- 來源清單 ----------------
SRC = []  # (id, 說明, URL, 日期, 等級)
def S(i, desc, url, date, level):
    SRC.append((i, desc, url, date, level)); return i
CIKS = {"MSFT": 789019, "AMZN": 1018724, "GOOGL": 1652044, "META": 1326801, "ORCL": 1341439, "CRWV": 1769628, "NBIS": 1513845, "DELL": 1571996, "SMCI": 1375365, "HPE": 1645590}
for t, c in CIKS.items():
    S(f"X-{t}", f"{t} SEC XBRL companyfacts（公司 10-Q/10-K 的 XBRL 申報值；每格的 accession 見資料列備註/來源欄）", f"https://data.sec.gov/api/xbrl/companyfacts/CIK{c:010d}.json", "取用 2026-11-18", "[Verified]（公司申報 XBRL）")
S("WY", "緯穎 公司網站合併財報 PDF（KPMG 核閱）：Wiwynn_2025Q1/2025Q4/2026Q1/2026Q2/2026Q3_Financial.pdf、2025Q2_WYHQ_FS_EN.pdf、2025Q3_WYHQ_FS_EN.pdf（Q2 2026 核閱報告日 2026-08-07）", "https://www.wiwynn.com/hubfs/Investors/Financial_Report/", "各季財報；取用 2026-11-18", "[Verified]")
S("SA-BS", "S&P Global 資料經 stockanalysis.com 季資產負債表（廣達 2382／鴻海 2317／緯創 3231；與緯穎公司原文 11 季逐季相符，見核對結果）", "https://stockanalysis.com/quote/tpe/{2382|2317|3231}/financials/balance-sheet/?period=quarterly", "取用 2026-11-18（頁面標示更新至 2026-11-12）", "[二手]")
S("SA-IS", "S&P Global 資料經 stockanalysis.com 季損益表（營收、營業成本＝營收－毛利，由頁面推導）", "https://stockanalysis.com/quote/tpe/{2382|2317|3231}/financials/?period=quarterly", "取用 2026-11-18", "[二手]")
S("YH-M", "Yahoo 奇摩股市月營收頁（2025-01 至 2026-09，千元精確值）", "https://tw.stock.yahoo.com/quote/{2382|2317|3231|6669}.TW/revenue", "取用 2026-11-18", "[二手]（轉載公司月營收公告）")
S("HS-M", "HiStock 月營收頁（含 2026-10；數字為單精度浮點，末位進位）", "https://histock.tw/stock/financial.aspx?no={2382|2317|3231|6669}&t=2", "取用 2026-11-18", "[二手]（轉載公司月營收公告）")
S("UA-Q", "優分析 廣達 3Q26 存貨 8,561 億、季增 659 億（引用優分析資料庫，未引觀測站）", "https://uanalyze.com.tw/articles/8736859074", "2026-11-11", "[二手]")
S("SD-Q", "財報狗 廣達 Q3 營收 1.1155 兆、存貨 8,561 億、應收 9,317 億", "https://statementdog.com/analysis/2382", "取用 2026-11-18", "[二手]")
S("UA-Q2", "優分析 廣達 9 月營收／Q3 營收與新世代伺服器放量時程", "https://uanalyze.com.tw/articles/2106956697", "2026-10-07", "[二手]")
S("SD-H", "財報狗 鴻海 Q3 法說：營收 3 兆 156.34 億、AI 機櫃 Q4 高雙位數季增", "https://statementdog.com/news/18049", "2026-11-12", "[二手]")
S("UA-H1", "優分析／鉅亨 鴻海 2Q26 法說（2026-08-12）：AI 機櫃、AI 伺服器占有率與 ASIC", "https://uanalyze.com.tw/articles/3510853294", "2026-08-12", "[二手]（法說，經轉述）")
S("UA-H2", "優分析 鴻海 3Q26 法說（雲端網路 54% 營收）", "https://uanalyze.com.tw/articles/3008659196", "2026-11-12", "[二手]")
S("UA-W", "優分析／鉅亨 緯創 2025 全年法說（2026-03-12）", "https://uanalyze.com.tw/articles/5710648963", "2026-03-13", "[二手]")
S("FU-WY", "Futurum 緯穎 Q3 2026（頁面自述為 AI 生成、未經編輯審閱）", "https://futurumgroup.com/?p=97797", "2026-11-11", "[二手]，可靠度低")
S("PU-D", "Pulse 2.0 Dell Q2 FY27 業績（公司新聞稿原文導向投資人網站，未取得）", "https://pulse2.com/dell-technologies-q2-revenue-hits-record-47-billion-as-ai-server-backlog-reaches-95-billion-and-fy27-outlook-rises-to-192-billion/", "2026-09-01", "[二手]")
S("SM-R", "Supermicro FQ1 FY27 Prepared Remarks（公司簡報 PDF）", "https://s204.q4cdn.com/707617056/files/content_files/FQ127-Prepared-Remarks.pdf", "2026-11（約 11-03/04，未逐字核對日期）", "[Interested-party]（公司說法）")
S("NA-H", "Nasdaq 轉載 HPE Q3 FY26 法說重點", "https://www.nasdaq.com/articles/hewlett-packard-enterprise-q3-earnings-call-highlights", "2026-09-02", "[二手]")
# 10-Q/10-K 原文（WebFetch 讀取；SEC Archives 直連被擋）
DOC = {
 "META-3Q26": ("Meta 10-Q 2026-09-30（附註：Property and equipment；Commitments）", "https://www.sec.gov/Archives/edgar/data/1326801/000162828026066773/meta-20260930.htm", "2026-10-22"),
 "META-2Q26": ("Meta 10-Q 2026-06-30", "https://www.sec.gov/Archives/edgar/data/1326801/000162828026050705/meta-20260630.htm", "2026-07-30"),
 "META-1Q26": ("Meta 10-Q 2026-03-31", "https://www.sec.gov/Archives/edgar/data/1326801/000162828026028526/meta-20260331.htm", "2026-04-30"),
 "META-3Q25": ("Meta 10-Q 2025-09-30", "https://www.sec.gov/Archives/edgar/data/1326801/000162828025047240/meta-20250930.htm", "2025-10-30"),
 "META-2Q25": ("Meta 10-Q 2025-06-30", "https://www.sec.gov/Archives/edgar/data/1326801/000162828025036791/meta-20250630.htm", "2025-07-31"),
 "META-1Q25": ("Meta 10-Q 2025-03-31", "https://www.sec.gov/Archives/edgar/data/1326801/000132680125000054/meta-20250331.htm", "2025-05-01"),
 "MSFT-1QFY27": ("Microsoft 10-Q 2026-09-30（Note 12 租賃；Note 6 P&E）", "https://www.sec.gov/Archives/edgar/data/789019/000119312526436133/msft-20260930.htm", "2026-10-28"),
 "MSFT-FY26": ("Microsoft 10-K FY2026（2026-06-30；Note 13 租賃）", "https://www.sec.gov/Archives/edgar/data/789019/000119312526323660/msft-20260630.htm", "2026-07-29"),
 "MSFT-3QFY26": ("Microsoft 10-Q 2026-03-31", "https://www.sec.gov/Archives/edgar/data/789019/000119312526191507/msft-20260331.htm", "2026-04-29"),
 "MSFT-2QFY26": ("Microsoft 10-Q 2025-12-31", "https://www.sec.gov/Archives/edgar/data/789019/000119312526027207/msft-20251231.htm", "2026-01-28"),
 "ORCL-1QFY27": ("Oracle 10-Q 2026-08-31", "https://www.sec.gov/Archives/edgar/data/1341439/000119312526389274/orcl-20260831.htm", "2026-09-11"),
 "CRWV-3Q26": ("CoreWeave 10-Q 2026-09-30", "https://www.sec.gov/Archives/edgar/data/1769628/000176962826000465/crwv-20260930.htm", "2026-11-12"),
 "CRWV-2Q26": ("CoreWeave 10-Q 2026-06-30", "https://www.sec.gov/Archives/edgar/data/1769628/000176962826000366/crwv-20260630.htm", "2026-08-12"),
 "CRWV-1Q26": ("CoreWeave 10-Q 2026-03-31", "https://www.sec.gov/Archives/edgar/data/1769628/000176962826000222/crwv-20260331.htm", "2026-05-08"),
 "CRWV-3Q25": ("CoreWeave 10-Q 2025-09-30", "https://www.sec.gov/Archives/edgar/data/1769628/000176962825000062/crwv-20250930.htm", "2025-11-13"),
 "GOOGL-3Q26": ("Alphabet 10-Q 2026-09-30（Note 7 P&E；Note 4 租賃；Note 10 承諾）", "https://www.sec.gov/Archives/edgar/data/1652044/000165204426000084/goog-20260930.htm", "2026-10-29"),
 "GOOGL-2Q26": ("Alphabet 10-Q 2026-06-30", "https://www.sec.gov/Archives/edgar/data/1652044/000165204426000071/goog-20260630.htm", "2026-07-23"),
 "GOOGL-1Q26": ("Alphabet 10-Q 2026-03-31", "https://www.sec.gov/Archives/edgar/data/1652044/000165204426000048/goog-20260331.htm", "2026-04-30"),
 "GOOGL-3Q25": ("Alphabet 10-Q 2025-09-30", "https://www.sec.gov/Archives/edgar/data/1652044/000165204425000091/goog-20250930.htm", "2025-10-30"),
 "AMZN-3Q26": ("Amazon 10-Q 2026-09-30（僅讀前約 10 萬字元：現金流量、Note 1）", "https://www.sec.gov/Archives/edgar/data/1018724/000101872426000040/amzn-20260930.htm", "2026-10-30"),
}
for k, (d, u, dd) in DOC.items(): S("D-" + k, d, u, dd, "[Verified]（公司 SEC 申報原文）")
S("NB", "Nebius 20-F（FY2025，XBRL companyfacts：年度資本支出 4,066.0 百萬美元；FY2024 807.7 百萬美元）；季度為 6-K，未取得", "https://data.sec.gov/api/xbrl/companyfacts/CIK0001513845.json", "20-F 申報日 2026-04-30（頁面 filed 欄）", "[Verified]（年度）")
S("NAD", "Nadella 談『chips sitting in inventory that I can't plug in』（播客，非法說；媒體轉述，日期約 2025-11，未驗證）", "https://www.tomshardware.com/tech-industry/artificial-intelligence/microsoft-ceo-says-the-company-doesnt-have-enough-electricity-to-install-all-the-ai-gpus-in-its-inventory-you-may-actually-have-a-bunch-of-chips-sitting-in-inventory-that-i-cant-plug-in", "約 2025-11（未驗證）", "[二手]")

wb = Workbook()

# ---------------- ODM季資料 ----------------
ws = wb.active; ws.title = "ODM季資料"
heads = ["公司", "曆季", "期末日(曆季)", "財報實際期末日", "幣別／單位", "營收", "營業成本", "期末存貨", "合約負債（流動；Dell 為總額）", "當季天數", "存貨天數＝存貨÷營業成本×天數", "存貨金額年增（對 4 季前）", "來源ID", "等級", "備註（accession／推導方式）"]
hdr(ws, 1, heads, [14, 8, 12, 12, 14, 14, 14, 14, 16, 8, 14, 12, 12, 16, 70])
ws.freeze_panes = "C2"
COMP = [("廣達 2382", "2382", "NT$百萬"), ("鴻海 2317", "2317", "NT$百萬"), ("緯創 3231", "3231", "NT$百萬"), ("緯穎 6669", "6669", "NT$百萬"),
        ("Dell", "DELL", "US$百萬"), ("Supermicro", "SMCI", "US$百萬"), ("HPE", "HPE", "US$百萬")]
row = 2; blocks = {}
for name, key, unit in COMP:
    blocks[key] = row
    for qi, q in enumerate(QS):
        ws.cell(row, 1, name); ws.cell(row, 2, q); ws.cell(row, 3, qend(q)).number_format = "yyyy-mm-dd"; ws.cell(row, 5, unit)
        ws.cell(row, 10, qdays(q))
        rv = cs = iv = cl = None; endd = None; note = ""; sid = ""; lvl = ""
        if key in ("2382", "2317", "3231"):
            rv = tw["rev"][key].get(q); cs = tw["cost"][key].get(q); iv = tw["inv"][key].get(q); cl = tw["cl"][key].get(q)
            endd = qend(q); sid = "SA-BS; SA-IS"; lvl = "[二手]"
            note = "S&P Global 經 stockanalysis.com；營業成本＝頁面營收－毛利；合約負債＝頁面『Current Unearned Revenue』（科目名稱與公司原文是否一致僅緯穎驗證過）"
            if key == "2382" and q == "2026Q3": note += "；存貨 856,064 與優分析、財報狗『8,561 億』一致（見核對結果）；公司觀測站原文未取得"
            if key == "2317" and q == "2026Q3": note += "；營收 3,015,634 與財報狗『3 兆 156.34 億』一致"
        elif key == "6669":
            endd = qend(q); sid = "WY"; lvl = "[Verified]"
            b = wy["bs"].get(q); iss = wy["is"].get(q)
            if b: iv = b[0] / 1000; cl = b[1] / 1000
            if iss: rv = iss[0] / 1000; cs = iss[1] / 1000
            note = "公司網站 PDF（KPMG 核閱，非查核）；原單位 NT$千元，已除以 1000"
            if q in ("2024Q4", "2025Q4"): note = "營收與營業成本＝全年值－前三季（Excel 公式，全年值見表下方輔助列）；存貨取自同季財報；" + "原單位 NT$千元"
        else:
            d = us[key]
            def g(k):
                x = d[k].get(q) if d.get(k) else None
                return x
            r_, c_, i_, l_ = g("rev"), g("cost"), g("inv"), g("cl")
            if r_: rv = r_["v"]; endd = dt.date.fromisoformat(r_["end"]); acc = r_["acc"]; fl = r_["filed"]
            if c_: cs = c_["v"]
            if i_: iv = i_["v"]
            if l_: cl = l_["v"]
            sid = "X-" + key; lvl = "[Verified]（Q4＝全年值－前三季，屬 [Derived]）"
            note = f"XBRL；財報期末 {r_['end'] if r_ else (i_['end'] if i_ else '—')}；來源申報 {('https://www.sec.gov/Archives/edgar/data/%d/%s/' % (CIKS[key], r_['acc'].replace('-', ''))) if r_ else ''}（申報日 {r_['filed'] if r_ else ''}；算法：{r_['how'] if r_ else ''}）"
            if key == "HPE": note += "；營業成本：XBRL 無對應標籤，本輪未取得原文，存貨天數留空（沒找到）"
            if key == "DELL" and q == "2026Q3": note = "Dell FQ3 FY27（2026-08～10）截至 2026-11-18 尚未公布（約 11 月下旬）"
            if key == "HPE" and q == "2026Q3": note = "HPE 2026Q3 對應的財報（FQ4 FY26，截至 10-31）截至 2026-11-18 尚未公布"
            if key == "DELL" and q == "2026Q3": rv = None
        if endd: ws.cell(row, 4, endd).number_format = "yyyy-mm-dd"
        for ci, v in ((6, rv), (7, cs), (8, iv), (9, cl)):
            if v is not None:
                c = ws.cell(row, ci, v); c.font = INP; c.number_format = "#,##0.0" if key == "6669" else "#,##0"
        ws.cell(row, 11, f'=IF(AND(ISNUMBER(H{row}),ISNUMBER(G{row}),N(G{row})>0),H{row}/G{row}*J{row},"")').number_format = "0.0"
        if qi >= 4:
            ws.cell(row, 12, f'=IF(AND(ISNUMBER(H{row}),ISNUMBER(H{row-4})),H{row}/H{row-4}-1,"")').number_format = "0.0%"
        ws.cell(row, 13, sid); ws.cell(row, 14, lvl); ws.cell(row, 15, note)
        if q in ("2024Q4",) and key == "6669": pass
        row += 1
# 緯穎全年值輔助與 Q4 公式
aux = row + 2
ws.cell(aux, 1, "緯穎全年值（NT$百萬；輸入，來源 WY：Wiwynn_2025Q4_Financial.pdf，原單位千元已除以 1000）").font = Font(bold=True)
ws.cell(aux + 1, 1, "年度"); ws.cell(aux + 1, 2, "營收"); ws.cell(aux + 1, 3, "營業成本")
for i, y in enumerate(("2024", "2025")):
    ws.cell(aux + 2 + i, 1, y)
    ws.cell(aux + 2 + i, 2, wy["fy"][y][0] / 1000).font = INP
    ws.cell(aux + 2 + i, 3, wy["fy"][y][1] / 1000).font = INP
    for cc in (2, 3): ws.cell(aux + 2 + i, cc).number_format = "#,##0.0"
w0 = blocks["6669"]
for yi, y in enumerate(("2024", "2025")):
    q4 = w0 + yi * 4 + 3; q1 = w0 + yi * 4
    ws.cell(q4, 6, f"=B{aux+2+yi}-SUM(F{q1}:F{q1+2})").number_format = "#,##0.0"
    ws.cell(q4, 7, f"=C{aux+2+yi}-SUM(G{q1}:G{q1+2})").number_format = "#,##0.0"
ws.cell(aux + 5, 1, "說明：存貨天數用『當季營業成本』年化前的原始比例（存貨÷季營業成本×當季天數），不是年度成本；Dell／HPE 以 91 天近似 13 週（Dell 第四季偶有 14 週，未調整）。曆季對齊規則＝財報期間中點所在的曆季（Dell FQ1＝2～4 月→曆 1 季；HPE FQ1＝11～1 月→曆前一年第 4 季）。")
ws.cell(aux + 6, 1, "AI 與非 AI 存貨：七家公司全部『公司沒揭露』拆分；本表不自行拆。英業達 2356（可選）本輪未蒐集。")
ODM_LAST = row - 1

# ---------------- ODM月營收 ----------------
wm = wb.create_sheet("ODM月營收")
hdr(wm, 1, ["月份", "廣達 2382", "鴻海 2317", "緯創 3231", "緯穎 6669", "廣達 YoY", "鴻海 YoY", "緯創 YoY", "緯穎 YoY", "四家合計", "四家合計 MoM"], [10, 16, 16, 16, 16, 10, 10, 10, 10, 18, 12])
wm.freeze_panes = "B2"
keys = ["2382", "2317", "3231", "6669"]
for i, m in enumerate(mo["months"]):
    r = i + 2; wm.cell(r, 1, m)
    for j, k in enumerate(keys):
        c = wm.cell(r, 2 + j, mo["data"][k][i] / 1000); c.font = INP; c.number_format = "#,##0.0"
        if i >= 12: wm.cell(r, 6 + j, f"={L(2+j)}{r}/{L(2+j)}{r-12}-1").number_format = "0.0%"
    wm.cell(r, 10, f"=SUM(B{r}:E{r})").number_format = "#,##0.0"
    if i >= 1: wm.cell(r, 11, f"=J{r}/J{r-1}-1").number_format = "0.0%"
n = len(mo["months"]); last = n + 1
r0 = last + 2
wm.cell(r0, 1, "單位：NT$百萬（來源原為千元，已除以 1000）。等級 [二手]（Yahoo 奇摩股市＋HiStock 互核；轉載公司月營收公告；公司公告原文／觀測站未取得）。2026-10 僅 HiStock（廣達、鴻海）或 Yahoo＋HiStock（緯穎）；緯創 2026-10 與 technews 標題 4,026.15 億一致。").alignment = WRAP
wm.merge_cells(start_row=r0, start_column=1, end_row=r0, end_column=11); wm.row_dimensions[r0].height = 48
wm.cell(r0 + 1, 1, "來源ID：YH-M、HS-M（見來源清單）。鴻海 2026-10 為單一來源（HiStock，單精度浮點，末位進位）。")
# 季度加總 vs 季報核對
cr = r0 + 3
wm.cell(cr, 1, "月營收季加總 vs 季報營收（NT$百萬）").font = Font(bold=True)
hdr(wm, cr + 1, ["曆季", "廣達 月加總", "廣達 季報", "鴻海 月加總", "鴻海 季報", "緯創 月加總", "緯創 季報", "緯穎 月加總", "緯穎 季報", "鴻海 差異%", "緯穎 差異%"])
qq = ["2025Q1", "2025Q2", "2025Q3", "2025Q4", "2026Q1", "2026Q2", "2026Q3"]
for qi, q in enumerate(qq):
    r = cr + 2 + qi; wm.cell(r, 1, q)
    mrow = 2 + qi * 3
    for j, k in enumerate(keys):
        wm.cell(r, 2 + 2 * j, f"=SUM({L(2+j)}{mrow}:{L(2+j)}{mrow+2})").number_format = "#,##0"
        srow = blocks[k] + 4 + qi
        wm.cell(r, 3 + 2 * j, f"='ODM季資料'!F{srow}").number_format = "#,##0"
    wm.cell(r, 10, f"=D{r}/E{r}-1").number_format = "0.00%"
    wm.cell(r, 11, f"=H{r}/I{r}-1").number_format = "0.00%"
wm.cell(cr + 2 + len(qq), 1, "觀察：廣達、緯創、緯穎月加總與季報幾乎相等；鴻海月加總與季報合併營收差 −0.5%～+0.5%（2026Q3 月加總高約 0.4%；原因未驗證，可能是月營收未經核閱或口徑不同）。")

# ---------------- 買方季資料 ----------------
wbuy = wb.create_sheet("買方季資料")
bh = ["公司", "曆季", "期末日(曆季)", "財報實際期末日", "現金資本支出（US$百萬）", "融資租賃本金償還（US$百萬）", "融資租賃新增使用權資產（非現金，US$百萬）",
      "資本支出：現金＋融資租賃新增（口徑A）", "資本支出：現金＋融資租賃本金（口徑B，Meta 自稱口徑）", "在建工程／尚未投入使用資產（US$百萬）", "該欄口徑", "伺服器網路／技術設備毛額（US$百萬）", "該欄口徑",
      "已簽約未起租租賃（US$十億）", "採購／不可取消承諾（US$十億）", "承諾口徑", "CIP ÷ 年化現金資本支出（月，口徑含建物，僅供對照）", "來源ID", "等級", "備註"]
hdr(wbuy, 1, bh, [12, 8, 11, 11, 13, 12, 14, 14, 16, 14, 28, 14, 24, 12, 12, 30, 14, 20, 16, 70])
wbuy.freeze_panes = "C2"
BUY = ["META", "MSFT", "GOOGL", "AMZN", "ORCL", "CRWV", "NBIS"]
BN = {"META": "Meta", "MSFT": "Microsoft", "GOOGL": "Alphabet", "AMZN": "Amazon", "ORCL": "Oracle", "CRWV": "CoreWeave", "NBIS": "Nebius"}
# 手工（原文）值：{公司: {季: {欄: (值, 來源ID)}}}
man = {
 "META": {
  "2024Q4": {"cip": (26802, "D-META-3Q25"), "eq": (68397, "D-META-3Q25")},
  "2025Q1": {"cip": (32385, "D-META-1Q25"), "eq": (75859, "D-META-1Q25"), "cm": (30.05, "D-META-1Q25")},
  "2025Q2": {"cip": (35942, "D-META-2Q25"), "eq": (85833, "D-META-2Q25"), "cm": (27.95, "D-META-2Q25")},
  "2025Q3": {"cip": (44015, "D-META-3Q25"), "eq": (92142, "D-META-3Q25"), "cm": (81.19, "D-META-3Q25")},
  "2025Q4": {"cip": (50521, "D-META-3Q26"), "eq": (98040, "D-META-3Q26")},
  "2026Q1": {"cip": (61017, "D-META-1Q26"), "eq": (105987, "D-META-1Q26"), "cm": (237.67, "D-META-1Q26")},
  "2026Q2": {"cip": (80345, "D-META-2Q26"), "eq": (119683, "D-META-2Q26"), "cm": (349.31, "D-META-2Q26")},
  "2026Q3": {"cip": (89499, "D-META-3Q26"), "eq": (133626, "D-META-3Q26"), "cm": (454.52, "D-META-3Q26")}},
 "MSFT": {
  "2025Q2": {"eq": (132836, "D-MSFT-3QFY26")},
  "2025Q4": {"eq": (171351, "D-MSFT-2QFY26"), "nc": (155.1, "D-MSFT-2QFY26")},
  "2026Q1": {"eq": (190883, "D-MSFT-3QFY26"), "nc": (196.6, "D-MSFT-3QFY26")},
  "2026Q2": {"eq": (215874, "D-MSFT-FY26"), "nc": (329.1, "D-MSFT-FY26")},
  "2026Q3": {"eq": (244840, "D-MSFT-1QFY27"), "nc": (364.6, "D-MSFT-1QFY27")}},
 "GOOGL": {
  "2024Q4": {"cip": (50597, "X-GOOGL")}, "2025Q1": {"cip": (55641, "X-GOOGL")},
  "2025Q3": {"cip": (69856, "D-GOOGL-3Q25"), "nc": (42.6, "D-GOOGL-3Q25")},
  "2025Q4": {"cip": (78592, "D-GOOGL-3Q26")},
  "2026Q1": {"cip": (108597, "D-GOOGL-1Q26"), "nc": (75.6, "D-GOOGL-1Q26"), "cm": (332.4, "D-GOOGL-1Q26")},
  "2026Q2": {"cip": (122814, "D-GOOGL-2Q26"), "cm": (707.0, "D-GOOGL-2Q26")},
  "2026Q3": {"cip": (145270, "D-GOOGL-3Q26"), "nc": (130.5, "D-GOOGL-3Q26"), "cm": (802.1, "D-GOOGL-3Q26")}},
 "AMZN": {},
 "ORCL": {"2026Q3": {"nc": (288.0, "D-ORCL-1QFY27")}},
 "CRWV": {
  "2024Q4": {"cip": (3201, "D-CRWV-3Q25"), "eq": (9147, "D-CRWV-3Q25")},
  "2025Q3": {"cip": (6910, "D-CRWV-3Q25"), "eq": (14629, "D-CRWV-3Q25"), "nc": (39.1, "D-CRWV-3Q25")},
  "2025Q4": {"cip": (9376, "D-CRWV-3Q26"), "eq": (20903, "D-CRWV-3Q26")},
  "2026Q1": {"cip": (9581, "D-CRWV-1Q26"), "eq": (26627, "D-CRWV-1Q26"), "nc": (40.7, "D-CRWV-1Q26")},
  "2026Q2": {"cip": (11918, "D-CRWV-2Q26"), "eq": (33823, "D-CRWV-2Q26"), "nc": (35.5, "D-CRWV-2Q26")},
  "2026Q3": {"cip": (9339, "D-CRWV-3Q26"), "eq": (47023, "D-CRWV-3Q26"), "nc": (47.4, "D-CRWV-3Q26")}},
 "NBIS": {},
}
cipnote = {"META": "CIP（含建物；Meta 未拆建物與 IT）", "MSFT": "公司沒揭露 CIP（Note 6 無此欄）", "GOOGL": "assets not yet in service（含建物、IT；未拆）", "AMZN": "本輪未取得（只讀 10-Q 前 10 萬字元）", "ORCL": "CIP（含建物；未拆）", "CRWV": "CIP（含建物；未拆）", "NBIS": "未取得（季度為 6-K）"}
eqnote = {"META": "Servers and network assets（毛額）", "MSFT": "Servers, network equipment, and software（成本）", "GOOGL": "未單列（Technical infrastructure 約 60% 為伺服器與網路，未逐季取）", "AMZN": "未取得", "ORCL": "Computer, network, machinery and equipment（毛額，XBRL MachineryAndEquipmentGross）", "CRWV": "Technology equipment（毛額）", "NBIS": "未取得"}
cmnote = {"META": "Non-cancelable contractual commitments（10-Q 文字）", "GOOGL": "各季口徑不同：2026Q1＝material purchase commitments 332.4（含短期 138.0）；2026Q2/Q3＝>1 年固定或保證承諾 707.0／802.1", "ORCL": "Unrecorded unconditional purchase obligations（XBRL）", "MSFT": "未取得", "AMZN": "未取得", "CRWV": "無單一總額（只列建設與設備承諾區間，見公司說法）", "NBIS": "未取得"}
brow = {}; row = 2
for t in BUY:
    brow[t] = row
    d = us.get(t, {})
    for q in QS:
        wbuy.cell(row, 1, BN[t]); wbuy.cell(row, 2, q); wbuy.cell(row, 3, qend(q)).number_format = "yyyy-mm-dd"
        sids = []; lv = set(); notes = []
        def put(col, rec, fmt="#,##0"):
            if rec:
                c = wbuy.cell(row, col, rec); c.font = INP; c.number_format = fmt
        cx = d.get("capex", {}).get(q) if d else None
        if cx:
            put(5, cx["v"]); wbuy.cell(row, 4, dt.date.fromisoformat(cx["end"])).number_format = "yyyy-mm-dd"
            sids.append("X-" + t); lv.add("[Verified]")
            notes.append(f"現金資本支出 {cx['how']}；申報 https://www.sec.gov/Archives/edgar/data/{CIKS[t]}/{cx['acc'].replace('-', '')}/（申報日 {cx['filed']}）")
        fp = d.get("fl_princ", {}).get(q) if d else None
        if fp: put(6, fp["v"], "#,##0.0")
        fr = d.get("fl_rou", {}).get(q) if d else None
        if fr: put(7, fr["v"], "#,##0.0")
        wbuy.cell(row, 8, f'=IF(AND(ISNUMBER(E{row}),ISNUMBER(G{row})),E{row}+G{row},"")').number_format = "#,##0"
        wbuy.cell(row, 9, f'=IF(AND(ISNUMBER(E{row}),ISNUMBER(F{row})),E{row}+F{row},"")').number_format = "#,##0"
        m = man.get(t, {}).get(q, {})
        cip = m.get("cip"); eq = m.get("eq"); nc = m.get("nc"); cm = m.get("cm")
        if t == "ORCL":
            x = d["cip"].get(q);
            if x: cip = (x["v"], "X-ORCL")
            x = d["equip"].get(q)
            if x: eq = (x["v"], "X-ORCL")
            x = d["purch"].get(q)
            if x: cm = (x["v"], "X-ORCL")
            xo = d["cip"].get(q) or d["capex"].get(q)
            if xo: wbuy.cell(row, 4, dt.date.fromisoformat(xo["end"])).number_format = "yyyy-mm-dd"
        if t == "META":
            x = d["lease_nc"].get(q)
            if x: nc = (x["v"], "X-META")
        if cip: put(10, cip[0]); sids.append(cip[1]); wbuy.cell(row, 11, cipnote[t]); lv.add("[Verified]")
        elif q in ("2025Q4", "2026Q1", "2026Q2", "2026Q3") and t in ("MSFT", "AMZN"):
            wbuy.cell(row, 11, cipnote[t])
        if eq: put(12, eq[0]); sids.append(eq[1]); wbuy.cell(row, 13, eqnote[t]); lv.add("[Verified]")
        if nc: put(14, nc[0], "#,##0.0"); sids.append(nc[1]); lv.add("[Verified]")
        if cm: put(15, cm[0], "#,##0.0"); sids.append(cm[1]); wbuy.cell(row, 16, cmnote.get(t, "")); lv.add("[Verified]")
        wbuy.cell(row, 17, f'=IF(AND(ISNUMBER(J{row}),ISNUMBER(E{row}),N(E{row})>0),J{row}/(E{row}*4)*12,"")').number_format = "0.0"
        if t == "META" and nc: notes.append("未起租租賃取自 XBRL 標籤 UnrecordedUnconditionalPurchaseObligationBalanceSheetAmount；僅 2026Q3 經 10-Q 文字確認為『尚未起租的租賃 355.73B』，其餘季度依標籤、未逐季讀文字")
        if t == "GOOGL" and q in ("2024Q4", "2025Q1"): notes.append("CIP 取自 XBRL ConstructionInProgressGross（2025-04-25 的 10-Q）")
        if t == "ORCL" and q in ("2024Q2", "2025Q2", "2026Q2"): notes.append("Oracle 10-Q 只在第 1 財季列 CIP；5 月底值取自 10-K／XBRL")
        if t == "ORCL" and q == "2026Q3": notes.append("Oracle FQ1 FY27（2026-06～08）；10-Q 文字：未起租租賃 $288 billion（預計 FY27Q2～FY2029 起租）")
        if t == "CRWV" and q in ("2025Q3", "2026Q1", "2026Q2", "2026Q3"): notes.append("未起租租賃＝估計未來未折現租金；2025Q3 取自 10-Q 文字")
        if t == "CRWV" and q == "2025Q3": notes.append("融資租賃新增 XBRL 為 −120（原文如此，未驗證）")
        if t == "NBIS":
            notes.append("Nebius 只申報 20-F（年度）；FY2024 資本支出 807.7、FY2025 4,066.0（US$百萬，XBRL）；季度 6-K 未取得，故季欄留空（沒找到，非不存在）")
            if q in ("2024Q4", "2025Q4"): sids.append("NB")
        if t == "AMZN": notes.append("現金資本支出＝Purchases of property and equipment（XBRL 標籤 PaymentsToAcquireProductiveAssets）；CIP、未起租、採購承諾本輪未取得")
        if t == "MSFT": notes.append("Microsoft 現金資本支出＝Additions to property and equipment；公司對外常用的『含融資租賃』數字見公司說法，未逐季取得，故以口徑A、B自算")
        if t == "META": notes.append("Meta 公布的『含融資租賃本金』資本支出＝現金＋融資租賃本金（口徑B）")
        wbuy.cell(row, 18, "; ".join(dict.fromkeys(sids))); wbuy.cell(row, 19, " / ".join(sorted(lv)) if lv else "")
        wbuy.cell(row, 20, "；".join(notes))
        row += 1

# ---------------- 公司說法 ----------------
wst = wb.create_sheet("公司說法")
hdr(wst, 1, ["公司", "時點／場合", "類別", "內容（摘要；引文≤15字）", "來源ID", "等級", "備註"], [14, 22, 22, 90, 12, 18, 50])
ST = [
 ("鴻海 2317", "2026-11-12 3Q26 法說", "AI 機櫃出貨", "Q4 AI 機櫃出貨『高雙位數成長』（較 Q3）；2027 年預期維持高雙位數成長；雲端網路產品占營收 54%；AI 伺服器占比與金額未揭露", "SD-H; UA-H2", "[二手]（法說，經轉述）", "公司沒揭露 AI 伺服器營收金額或占比"),
 ("鴻海 2317", "2026-08-12 2Q26 法說", "AI 機櫃出貨／AI 伺服器占有率", "Q3 AI 機櫃出貨預估高雙位數季增；全年維持『倍數以上成長』目標；AI 伺服器占有率朝 50% 目標；ASIC 伺服器約占去年 AI 伺服器營收 10%，上半年 ASIC 營收已略超過去年全年；Vera Rubin 預計 Q4 放量", "UA-H1", "[二手]（法說，經轉述）", "機櫃顆數／台數未揭露"),
 ("廣達 2382", "2026-10-07 公告 9 月營收", "AI 伺服器", "新世代伺服器 Q3 底試產、本季小量出貨、明年 Q1 放量；伺服器展望維持不變；AI 伺服器占比、機櫃出貨量未揭露；併購友達龜山廠供 AI 伺服器後段組裝，預計明年 Q4 交付", "UA-Q2", "[二手]", "公司沒揭露 AI 占比（本文未見；法說逐字稿未取得）"),
 ("廣達 2382", "3Q26 財報", "毛利率", "財報狗：毛利率 4.75%、營收 1.1155 兆、存貨 8,561 億", "SD-Q", "[二手]", "與 S&P 資料一致（毛利 53,024／營收 1,115,545＝4.75%）"),
 ("緯創 3231", "2026-03-12 法說（2025 全年）", "AI 伺服器／網通", "AI 伺服器展望正向；全年營收預期高雙位數成長；網通出貨預期 10 倍成長；系統級出貨 2026Q1 開始（前一年為 Q3）；AI 占比、機櫃量未揭露", "UA-W", "[二手]", "3Q26 法說（約 2026-11-中）本輪未取得"),
 ("緯穎 6669", "2026-11 Q3 財報", "AI 伺服器占比", "『AI servers ... roughly half』of revenue（出自轉述的 CEO 說法，非公司新聞稿）；原料供應吃緊使月營收短期波動", "FU-WY", "[二手]，可靠度低", "來源頁自述 AI 生成、未編輯；公司原文未取得；不得作為數值使用"),
 ("Dell", "2026-09-01 Q2 FY27", "AI 伺服器訂單／營收／在手訂單", "AI 伺服器訂單 60.9B、AI 伺服器營收 16.4B（當季）、AI 在手訂單 95B、FY27 AI 伺服器營收展望 74B（由 60B 上修）；ISG 營收 31.8B", "PU-D", "[二手]", "Dell 新聞稿原文導向投資人首頁，未取得；『營收 vs 出貨』口徑未核"),
 ("Supermicro", "FQ1 FY27 簡報（約 2026-11-03/04）", "AI 占比、存貨、預收", "AI 解決方案約占 Q1 營收 71%（上季 59%）；存貨 15.1B（前季 12.9B）；存貨天數 74，目標 50 天（擷取工具輸出語意不清，前季值未確認）；遞延收入季增約 3.1B，大多是客戶出貨前付款，『客戶預付款支應大部分建置』；12 月前 DLC 機櫃產能目標逾 4,000 櫃；FY27 營收展望 72–82B；展望受客戶設施就緒與零組件供應限制", "SM-R", "[Interested-party]（公司說法）", "摘要為擷取工具轉述，非逐字；『存貨天數 74』與本表由 XBRL 算出的 Q3 值（15,054÷17,405×92≈79.6）口徑不同，未對齊"),
 ("HPE", "2026-09-02 Q3 FY26", "AI 系統", "AI 系統營收近 1.6B、AI 系統訂單 2.4B（較上季 +30% 以上）；存貨 11.8B，原因是商品成本上升與為支應訂單與在手訂單的目標性採購；總在手訂單創高但未揭露 AI 專屬在手訂單", "NA-H", "[二手]", ""),
 ("Alphabet", "10-Q 2026-09-30 Note 7", "已收到但尚未上線（資產）", "assets not yet in service：2025-12-31 為 78,592、2026-03-31 為 108,597、2026-06-30 為 122,814、2026-09-30 為 145,270（US$百萬）；含建物與 IT，未拆", "D-GOOGL-3Q26; D-GOOGL-2Q26; D-GOOGL-1Q26", "[Verified]", "這是三家買方中唯一以法定附註逐季揭露的『未投入使用』餘額"),
 ("Meta", "10-Q 2026-09-30", "未起租與預付", "尚未起租的租賃約 355.73B，預計 2026 年剩餘期間至 2036 年起租；受限現金 10.80B 為『多年期基礎設施採購協議』的託管，預計 2028–2030 釋放（工具轉述，未逐字核）", "D-META-3Q26", "[Verified]", "公司未使用『placed in service』字樣"),
 ("CoreWeave", "10-Q 2026-09-30", "電力尚未交付", "『453 MW』電力於某站點尚未交付（分階段 2026–2028）；另一組站點 354 MW 尚未交付；未揭露 GW 級合約電力或在役電力（在役／合約 GW 說法需讀法說，未取得）", "D-CRWV-3Q26", "[Verified]", "此為資料中心電力，不是 GPU 已收貨未上線"),
 ("Microsoft", "播客（非法說）", "已收貨未上線", "Nadella：『chips sitting in inventory that I can't plug in』（約 10 字；電力與機房殼體不足）", "NAD", "[二手]", "出處為播客訪談，非法說逐字稿；日期未驗證；10-Q 未見等價揭露"),
 ("Oracle", "—", "已收貨未上線", "本輪未讀法說逐字稿，未取得（沒找到，非不存在）", "", "—", ""),
 ("Amazon", "—", "已收貨未上線", "本輪 10-Q（Q3 2026）前 10 萬字元未見；法說逐字稿未取得", "D-AMZN-3Q26", "—", "AWS 履約義務 約 578B（2026-09-30，原文轉述）"),
 ("Nebius", "—", "全部", "只申報 20-F／6-K；季度資本支出、在建工程、承諾本輪未取得", "NB", "—", "FY2025 資本支出 4,066.0 百萬美元（XBRL）"),
 ("xAI／OpenAI", "—", "全部", "公開文件本輪未查（可選項目，未做）", "", "—", ""),
]
for i, rr in enumerate(ST, 2):
    for j, v in enumerate(rr, 1):
        c = wst.cell(i, j, v); c.alignment = WRAP

# ---------------- 核對結果 ----------------
wc = wb.create_sheet("核對結果")
hdr(wc, 1, ["#", "項目", "設計書引用值", "單位", "原文／查得值", "差異（查得−引用）", "差異%", "判定", "查得值出處（原文位置）", "來源ID", "等級", "備註"], [4, 34, 14, 10, 14, 12, 9, 22, 50, 18, 16, 70])
CK = [
 (1, "廣達 3Q26 期末存貨", 856100, "NT$百萬", 856064, "無法核對原文（二手三源一致）", "S&P Global 季資產負債表 Inventory 856,064 百萬；優分析『第 3 季底存貨 8,561 億元，比第 2 季底增加 659 億元』；財報狗『存貨 8,561 億』", "SA-BS; UA-Q; SD-Q", "[二手]", "公司財報（觀測站／廣達 IR）被代理擋住，原文未取得，故不能嚴格說『與公司財報原文相符』。內部一致性：S&P 兩季差 856,064−790,184＝65,880，與優分析『659 億』吻合。"),
 (2, "廣達 2Q26 期末存貨（設計書由季增推算）", 790200, "NT$百萬", 790184, "無法核對原文（二手一致）", "S&P Global 2026-06-30 Inventory 790,184；設計書以 8,561−659＝7,902 億推算", "SA-BS", "[二手]", ""),
 (3, "Meta 2026-09-30 在建工程（CIP）", 89500, "US$百萬", 89499, "相符", "10-Q 2026-09-30 附註 Property and equipment：Construction in progress 89,499（2025-12-31 為 50,521）", "D-META-3Q26", "[Verified]", "取得方式：WebFetch 讀 SEC 原文（第 100,000–200,000 字元區段）；工具為摘要型，數字經 6 個季度前後交叉一致（見買方季資料）。Meta 未拆建物與 IT。"),
 (4, "Meta 伺服器網路資產毛額 2025-12 → 2026-09", 98000, "US$百萬", 98040, "相符", "同上：Servers and network assets 98,040（2025-12-31）→ 133,626（2026-09-30）；設計書 98.0 → 133.6", "D-META-3Q26", "[Verified]", "設計書 133.6 對應 133,626，相符"),
 (5, "Oracle 2026-08-31 在建工程（CIP）", 48500, "US$百萬", 48546, "相符", "10-Q 2026-08-31（FQ1 FY27）Property, plant and equipment：Construction in progress 48,546（2026-05-31 為 39,973）；XBRL ConstructionInProgressGross 同值", "D-ORCL-1QFY27; X-ORCL", "[Verified]", "Oracle 未拆建物與 IT"),
 (6, "Microsoft 已簽約未起租租賃", 329100, "US$百萬", 329100, "相符（但時點已過期）", "10-K FY2026（2026-06-30）Note 13 租賃：『additional leases, primarily for datacenters, not yet commenced』$329.1 billion，預計 FY2027–FY2033 起租", "D-MSFT-FY26", "[Verified]", "設計書若標為最新值則過期：2026-09-30 10-Q 已為 364.6B（FY2027–FY2034 起租）；時序 155.1（2025-12）→196.6（2026-03）→329.1（2026-06）→364.6（2026-09）。"),
]
for i, rr in enumerate(CK, 2):
    n_, item, ref, unit, got, verdict, loc, sid, lvl, note = rr
    wc.cell(i, 1, n_); wc.cell(i, 2, item); wc.cell(i, 3, ref).font = INP; wc.cell(i, 4, unit); wc.cell(i, 5, got).font = INP
    wc.cell(i, 6, f"=E{i}-C{i}").number_format = "#,##0"; wc.cell(i, 7, f"=F{i}/C{i}").number_format = "0.00%"
    wc.cell(i, 8, verdict); wc.cell(i, 9, loc); wc.cell(i, 10, sid); wc.cell(i, 11, lvl); wc.cell(i, 12, note)
    for j in (2, 8, 9, 12): wc.cell(i, j).alignment = WRAP
    wc.cell(i, 3).number_format = "#,##0"; wc.cell(i, 5).number_format = "#,##0"
r = len(CK) + 4
wc.cell(r, 1, "二手來源可信度驗證：S&P Global（stockanalysis.com）緯穎 2024Q1–2026Q3 共 11 季的存貨、合約負債、總資產，與緯穎公司網站 KPMG 核閱 PDF 逐季相符（四捨五入至百萬）。").alignment = WRAP
wc.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
wc.cell(r + 1, 1, "未驗證：廣達、鴻海、緯創的合約負債科目名稱是否等同公司『合約負債—流動』；鴻海月營收 vs 季報差異原因。")
wc.merge_cells(start_row=r + 1, start_column=1, end_row=r + 1, end_column=12)

# ---------------- 涵蓋率（公式） ----------------
wv = wb.create_sheet("涵蓋率")
hdr(wv, 1, ["類別", "公司", "項目", "有資料季數（共 11 季）", "缺的季度／說明"], [10, 14, 34, 16, 90])
rr = 2
ODM = "'ODM季資料'"; BUYS = "'買方季資料'"
for name, key, unit in COMP:
    a = blocks[key]; b = a + 10
    for lab, col in (("營收", "F"), ("營業成本", "G"), ("期末存貨", "H"), ("合約負債", "I")):
        wv.cell(rr, 1, "ODM／OEM"); wv.cell(rr, 2, name); wv.cell(rr, 3, lab)
        wv.cell(rr, 4, f"=COUNT({ODM}!{col}{a}:{col}{b})")
        miss = [q for qi, q in enumerate(QS) if ws.cell(a + qi, ord(col) - 64).value is None]
        wv.cell(rr, 5, ("缺：" + "、".join(miss)) if miss else "")
        rr += 1
    wv.cell(rr, 1, "ODM／OEM"); wv.cell(rr, 2, name); wv.cell(rr, 3, "存貨天數（可算）")
    wv.cell(rr, 4, f"=COUNT({ODM}!K{a}:K{b})")
    miss = [q for qi, q in enumerate(QS) if ws.cell(a + qi, 7).value is None or ws.cell(a + qi, 8).value is None]
    wv.cell(rr, 5, ("缺：" + "、".join(miss)) if miss else ""); rr += 1
for t in BUY:
    a = brow[t]; b = a + 10
    for lab, col in (("現金資本支出", "E"), ("融資租賃本金", "F"), ("融資租賃新增ROU", "G"), ("在建工程／未投入使用", "J"), ("伺服器網路／技術設備毛額", "L"), ("已簽未起租租賃", "N"), ("採購／不可取消承諾", "O")):
        wv.cell(rr, 1, "買方"); wv.cell(rr, 2, BN[t]); wv.cell(rr, 3, lab)
        wv.cell(rr, 4, f"=COUNT({BUYS}!{col}{a}:{col}{b})")
        miss = [q for qi, q in enumerate(QS) if wbuy.cell(a + qi, ord(col) - 64).value is None]
        wv.cell(rr, 5, ("缺：" + "、".join(miss)) if len(miss) < 11 and miss else ("全缺（公司沒揭露或本輪沒找到，見備註）" if len(miss) == 11 else "")); rr += 1
wv.cell(2 + 0, 6, "")
wb.save(B + "20261118_通路與買方端存量_資料.xlsx")
json.dump({"blocks": blocks, "brow": brow}, open(B + "data/layout.json", "w"))
# 來源清單（最後建立，順序放在核對結果之前）
wsrc = wb.create_sheet("來源清單", 3)
hdr(wsrc, 1, ["來源ID", "說明", "URL", "日期", "等級"], [14, 90, 90, 28, 28])
for i, s in enumerate(SRC, 2):
    for j, v in enumerate(s, 1): wsrc.cell(i, j, v).alignment = WRAP
wb.move_sheet("核對結果", offset=0)
order = ["ODM季資料", "ODM月營收", "買方季資料", "公司說法", "來源清單", "核對結果", "涵蓋率"]
wb._sheets = [wb[n] for n in order]
wb.save(B + "20261118_通路與買方端存量_資料.xlsx")
print("saved", ws.max_row, wbuy.max_row)
