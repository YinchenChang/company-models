# 建立「AI 軟體產業地位 v0.1」活頁簿（方法比照 實體AI產業地位 v0.5 / AI 算力地圖 build_xlsx.py）
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

import os, sys
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dist", "AI軟體地圖_產業地位_v0.2.xlsx")
os.makedirs(os.path.dirname(OUT), exist_ok=True)
BLUE = Font(color="0000CC")
BOLD = Font(bold=True)
YEL = PatternFill("solid", fgColor="FFF2CC")
HEAD = PatternFill("solid", fgColor="D9E1F2")

wb = Workbook()

# ---------- 說明 ----------
ws = wb.active
ws.title = "說明"
notes = [
    "AI 軟體產業地位 v0.2（2026-11-16，AI軟體應用專案；供「AI 軟體地圖（暫名）」與開工排序使用）",
    "用途：比照 AI 算力地圖與實體 AI 地圖的「產業地位」頁，列出本專案要覆蓋的次產業與重要個股，給出地位分數；另加「次產業優先序」頁，建議先做哪幾個次產業。",
    "方法（與 AI 算力地圖 build_xlsx.py、實體AI產業地位 v0.5 相同）：地位分數 ＝ w_irr × 不可替代性分 ＋ w_pos × 市場位置分 ＋ w_vc × 價值鏈分量分。",
    "不可替代性分 ＝（次產業不可替代性換分 ＋ 公司重要性換分）÷ 2；1→100、2→60、3→30。市場位置（次產業內）：1 龍頭→100、2 前二或寡占→75、3 主要參與者→50、4 追隨者→25。",
    "差異 1（分量的分母）：軟體沒有 BOM，價值鏈分量改用『次產業全球營收池』分級（≥3,000 億美元→100；1,000–3,000 億→75；300–1,000 億→50；<300 億→25）。營收池為 Analogy（Gartner 等研調，Interested-party），單位經濟模型完成後改讀模型。",
    "差異 2：新增『AI 方向』欄（受益／分歧／侵蝕，判斷，不進分數）。地位分數衡量『今天的市場地位』，不衡量 AI 讓它變好或變壞；後者是本專案單位經濟模型要回答的問題。地位高＋侵蝕＝最需要先算清楚的標的。",
    "差異 3：同一公司可出現在多個次產業（例如 Microsoft、Tencent）；排名以列計。",
    "亞太掛牌：『是』＝在日、韓、台、港、澳洲或中國 A 股掛牌；『待核』＝亞太公司但只在美股掛牌；印度本版不列入亞太。",
    "地圖歸屬（沿用專案負責人 2026-11-16 22:09 規則）：主圖＝持有該公司個股模型的專案所屬地圖；另一張圖照列、計入排名，標『跨圖』，目標價與評等只讀主圖模型。Alphabet、Meta、Microsoft、Amazon、Oracle、Tencent、Alibaba、Baidu 與模型公司的整家公司模型權威在 AI供應鏈（Instructions 第 3(b) 條），本表只讀。",
    "輸入：藍字＝判斷或抄錄輸入（黃底可改）；黑字＝公式。標記：Verified／Interested-party／Analogy／Assumed／Derived；[判斷] 等同 Assumed。",
    "限制：公司名單與地位依截至 2026 年中的公開資訊與判斷；市場份額未逐家查證；上市狀態有疑者標『待核』；未做估值，故無『可投資性』頁。",
    "市場位置口徑（v0.1 核對後定）：以全球口徑給分；只在單一國家或區域稱龍頭者，市場位置最高給 2。已下市公司（EA、SCSK、NTT Data）仍列為競爭參照，亞太欄標未上市。『澳』指澳洲掛牌，視為亞太。",
    "建議批次的覆寫原則：計算排名決定『哪類問題先做』，批次則依『單位經濟原型』安排——第一批每個原型（按人力、按席位、按廣告、按用量、按防護範圍計價）各做一家樣板；同原型的其他次產業（PRD、HCM、WFL 接 CRM；BPO、AGY 接 ITS）在樣板驗收後複製，所以計算排名較高者也可能排在第二、三批。BPO 營收池 300、PRD 100 剛好落在分量門檻上，小幅調整會換檔。",
    "v0.2 變更（2026-11-16 23:30）：印度暫不考慮（專案負責人 23:17）；資安（SEC）與資料平台（DATA）的 AI 衝擊強度由 3 改 1、代表性由 3 改 1——v0.1 把『共識受益、爭議小』誤當成『衝擊小』，經專案負責人 23:18 指出後更正；兩者列入第一批。ADS 樣板定為 LY Corp。",
    "次產業優先序：準則只用『AI 衝擊強度、核心命題代表性、兩端重疊、營收池』；資料是否好取得不作選題準則（只影響怎麼做）。權重與各準則評分為判斷，可在黃底格改。",
]
for i, t in enumerate(notes, 1):
    ws.cell(i, 1, t).alignment = Alignment(wrap_text=True, vertical="top")
ws.column_dimensions["A"].width = 150

# ---------- 參數 ----------
wp = wb.create_sheet("參數")
params = [
    ("w_irr", 0.40, "地位分數權重：不可替代性"),
    ("w_pos", 0.35, "地位分數權重：市場位置"),
    ("w_vc", 0.25, "地位分數權重：價值鏈分量"),
    ("s_irr1", 100, "不可替代性／重要性 1 → 分數"),
    ("s_irr2", 60, "2 → 分數"),
    ("s_irr3", 30, "3 → 分數"),
    ("s_pos1", 100, "市場位置 1 龍頭"),
    ("s_pos2", 75, "2 前二或寡占"),
    ("s_pos3", 50, "3 主要參與者"),
    ("s_pos4", 25, "4 追隨者"),
    ("pool_t1", 300, "營收池門檻（十億美元）≥ → 分量 100"),
    ("pool_t2", 100, "≥ → 75"),
    ("pool_t3", 30, "≥ → 50；以下 25"),
    ("p_imp", 0.35, "優先序權重：AI 衝擊強度"),
    ("p_rep", 0.25, "優先序權重：核心命題代表性"),
    ("p_ovl", 0.25, "優先序權重：兩端重疊"),
    ("p_pool", 0.15, "優先序權重：營收池"),
    ("ovl_cut", 55, "兩端重疊：只計地位分數 ≥ 此值的亞太標的"),
]
wp.append(["名稱", "值", "說明"])
for r in params:
    wp.append(list(r))
wp.append(["地位權重合計檢查", "=B2+B3+B4", "應為 1"])
wp.append(["優先序權重合計檢查", "=B15+B16+B17+B18", "應為 1"])
P = {name: f"參數!$B${i+2}" for i, (name, _, _) in enumerate(params)}
for row in wp.iter_rows(min_row=2, max_row=len(params) + 1, min_col=2, max_col=2):
    for c in row:
        c.font = BLUE
        c.fill = YEL
wp.column_dimensions["A"].width = 22
wp.column_dimensions["C"].width = 44

# ---------- 層級與次產業 ----------
# 代碼, 層級, 次產業, 關鍵問題, 營收池文字, 營收池中值(十億美元), 不可替代性, AI方向, 主要定價, 依據與標記
subs = [
    ("DATA", "S1 基礎與資料層", "資料平台、資料庫與可觀測性", "AI 工作負載會不會讓資料層的用量計價營收加速？", "約 1,200–1,800 億美元", 150, 2, "受益", "用量計價",
     "DBMS＋可觀測性市場規模 [Analogy：Gartner 等，Interested-party]；資料是代理的燃料"),
    ("SEC", "S1 基礎與資料層", "資安", "代理數量暴增後，身分與端點的防護面積放大多少？", "約 1,500–2,500 億美元（含服務）", 200, 1, "受益", "訂閱＋端點／用量",
     "資安支出 [Analogy：Gartner，Interested-party]；合規剛性 → 不可替代性 1"),
    ("DEV", "S1 基礎與資料層", "開發工具與 DevOps", "AI 寫程式後，開發工具按人頭收費的基數會縮還是擴？", "約 300–800 億美元", 50, 3, "分歧", "每席位",
     "[Assumed：區間]；AI 原生寫程式工具直接競爭"),
    ("CRM", "S2 企業應用層", "前台：CRM、客服與行銷自動化", "客服與業務席位被代理取代後，按任務計價能否補回營收？", "約 1,000–1,400 億美元", 120, 2, "侵蝕", "每席位 → 每任務過渡中",
     "CRM 市場 [Analogy：Gartner／IDC]；2026 年『SaaS 末日』賣壓核心"),
    ("ERP", "S2 企業應用層", "ERP、財務與供應鏈（記錄系統）", "記錄系統是代理的資料來源：護城河更深，還是被代理繞過？", "約 1,200–1,800 億美元", 150, 1, "分歧", "每席位／模組",
     "ERP 市場 [Analogy]；轉換成本高 → 不可替代性 1"),
    ("HCM", "S2 企業應用層", "人資、薪酬與招聘平台", "企業少招人，人資席位與招聘廣告量會同步縮嗎？", "約 600–1,000 億美元", 80, 2, "侵蝕", "每員工／每職缺",
     "HCM 軟體＋線上招聘 [Analogy]；招聘量與 AI 替代率直接連動"),
    ("WFL", "S2 企業應用層", "工作流、ITSM 與流程自動化（含 RPA）", "代理平台會吃掉 RPA 與工單系統，還是工作流平台成為代理的調度層？", "約 400–800 億美元", 60, 2, "分歧", "每席位＋用量",
     "[Analogy]；ServiceNow 為定價轉型觀察指標"),
    ("PRD", "S2 企業應用層", "設計、文件與協作生產力", "生成式工具把創作門檻降到零後，專業席位還值多少？", "約 800–1,200 億美元", 100, 3, "侵蝕", "每席位",
     "[Analogy]；模型直接替代風險最高 → 不可替代性 3"),
    ("VRT", "S2 企業應用層", "垂直軟體與專業資訊（法律、醫療、金融、政府）", "專有資料是護城河，還是被模型以低價重做？", "約 1,500–2,500 億美元", 200, 2, "分歧", "訂閱（資料＋工作流）",
     "[Analogy]；2026-02 法律與資訊股因 AI 代理外掛急跌"),
    ("ITS", "S3 服務層", "IT 服務、顧問與外包", "按人天計價被按成果計價取代，AI 導入新需求能否抵銷通縮？", "約 1.2–1.6 兆美元（不含 IaaS）", 1400, 3, "侵蝕", "人天／專案 → 成果",
     "Gartner IT 服務 2026 約 1.87 兆（含 IaaS）[Interested-party]；扣 IaaS 為 [Derived]；印度廠稱每年 2–3% AI 通縮"),
    ("BPO", "S3 服務層", "業務流程外包與客服中心", "客服與後台人力被代理替代的速度？", "約 2,000–4,000 億美元", 300, 3, "侵蝕", "人力時數",
     "[Analogy]；人力替代率最直接的次產業"),
    ("ADS", "S4 網路平台層", "數位廣告：搜尋、社群與廣告科技", "AI 提升轉換率帶來的增量，大於聊天介面分流搜尋的損失嗎？", "約 7,000–9,000 億美元", 800, 2, "受益", "競價廣告",
     "全球數位廣告 [Analogy：媒體集團預估，Interested-party]"),
    ("AGY", "S4 網路平台層", "廣告代理與行銷服務", "創意與媒體採購被平台自動化後，代理商的佣金還剩多少？", "約 600–1,000 億美元", 80, 3, "侵蝕", "佣金／人力",
     "[Analogy]；平台自動投放（Advantage+、PMax）直接替代"),
    ("ECM", "S4 網路平台層", "電商、旅遊與本地服務平台", "購物代理繞過入口，抽成與廣告營收會被誰拿走？", "約 3,000–8,000 億美元（平台抽成＋廣告＋物流收入）", 500, 2, "分歧", "抽成＋廣告",
     "[Assumed：區間大]；代理式購物為新變數"),
    ("GAM", "S4 網路平台層", "遊戲", "AI 降低開發成本，是放大 IP 龍頭，還是讓內容供給氾濫？", "約 1,800–2,000 億美元", 190, 2, "受益", "買斷＋內購",
     "[Analogy：Newzoo，Interested-party]；亞太權重最高的次產業"),
    ("CNT", "S4 網路平台層", "內容、串流與 IP", "AI 生成內容讓內容成本下降，還是稀釋付費意願？", "約 2,000–3,000 億美元", 250, 2, "分歧", "訂閱＋廣告",
     "[Analogy]"),
    ("AIN", "S5 AI 原生應用", "AI 原生應用與代理", "原生應用的毛利能否撐過模型商下游化與推論成本？", "約 200–600 億美元（年化營收）", 40, 3, "受益", "訂閱＋用量／成果",
     "[Analogy：創投與研調統計，Interested-party]；成長最快、上市標的少"),
]
wsb = wb.create_sheet("層級與次產業")
hdr = ["代碼", "層級", "次產業", "關鍵問題", "營收池（文字）", "營收池中值（十億美元）", "不可替代性（1–3）", "價值鏈分量分",
       "AI 方向（判斷，不進分數）", "主要定價模式", "依據與標記", "覆蓋公司數", "亞太標的數（是）"]
wsb.append(hdr)
for i, s in enumerate(subs, 2):
    code, layer, name, q, pooltxt, pool, irr, aid, price, basis = s
    wsb.append([code, layer, name, q, pooltxt, pool, irr,
                f"=IF(F{i}>={P['pool_t1']},100,IF(F{i}>={P['pool_t2']},75,IF(F{i}>={P['pool_t3']},50,25)))",
                aid, price, basis,
                f"=COUNTIF(個股地位!$F:$F,A{i})",
                f'=COUNTIFS(個股地位!$F:$F,A{i},個股地位!$D:$D,"是*")'])
    for col in (6, 7, 9):
        wsb.cell(i, col).font = BLUE
        wsb.cell(i, col).fill = YEL
NSUB = len(subs)
last = NSUB + 1
wsb.append(["", "", "營收池合計（十億美元）", "", "", f"=SUM(F2:F{last})"])

# ---------- 個股地位 ----------
# 代號, 公司, 市場, 亞太, 層級(lookup), 次產業代碼, 重要性, 市場位置, 地位依據, 標記, AI方向(公司), 地圖歸屬
R = []
def add(t, name, mkt, apac, code, imp, pos, basis, tag="[判斷]", ai="", mapown="主圖：AI 軟體地圖"):
    R.append((t, name, mkt, apac, code, imp, pos, basis, tag, ai, mapown))

A = "是"; AS = "是（A 股）"; US = "否"; ADR = "待核（亞太公司、美股掛牌）"; IN = "否（印度，本版不列入亞太）"; PV = "未上市"
X = "跨圖（主圖：AI 算力地圖；模型權威 AI供應鏈）"
# DATA
add("DBRX", "Databricks", "未上市", PV, "DATA", 1, 1, "資料與 AI 平台營收規模已居首（年化營收超越 Snowflake，待核）", "[Interested-party：公司揭露]", "受益")
add("SNOW", "Snowflake", "美", US, "DATA", 1, 2, "雲端資料倉儲龍頭之一，用量計價", ai="受益")
add("ORCL", "Oracle（資料庫）", "美", US, "DATA", 1, 1, "企業資料庫龍頭；整家公司模型在 AI供應鏈", ai="受益", mapown=X)
add("DDOG", "Datadog", "美", US, "DATA", 2, 1, "可觀測性龍頭；AI 工作負載監控需求", ai="受益")
add("MDB", "MongoDB", "美", US, "DATA", 2, 2, "文件型資料庫龍頭", ai="受益")
add("PLTR", "Palantir", "美", US, "DATA", 1, 2, "資料整合＋AI 決策平台（AIP），政府與企業", ai="受益")
add("688031", "星環科技", "陸", AS, "DATA", 3, 4, "中國大數據平台", ai="受益")
# SEC
add("CRWD", "CrowdStrike", "美", US, "SEC", 1, 1, "端點與雲端資安平台龍頭", ai="受益")
add("PANW", "Palo Alto Networks", "美", US, "SEC", 1, 1, "網路資安平台龍頭；平台化整併", ai="受益")
add("ZS", "Zscaler", "美", US, "SEC", 2, 2, "零信任存取龍頭", ai="受益")
add("FTNT", "Fortinet", "美", US, "SEC", 2, 2, "防火牆出貨量領先", ai="受益")
add("NET", "Cloudflare", "美", US, "SEC", 2, 2, "邊緣網路與安全；代理流量管控", ai="受益")
add("OKTA", "Okta", "美", US, "SEC", 2, 2, "身分管理；代理身分為新需求", ai="受益")
add("CHKP", "Check Point", "美", US, "SEC", 3, 3, "老牌防火牆", ai="分歧")
add("4704", "Trend Micro 趨勢科技", "日", A, "SEC", 2, 3, "亞太最大上市資安軟體公司（台灣創辦、東京掛牌）", ai="分歧")
add("688561", "奇安信", "陸", AS, "SEC", 3, 3, "中國資安龍頭之一", ai="分歧")
add("300454", "深信服", "陸", AS, "SEC", 3, 3, "中國網路安全與雲", ai="分歧")
# DEV
add("MSFT-GH", "Microsoft（GitHub）", "美", US, "DEV", 1, 1, "GitHub 與 Copilot；整家公司模型在 AI供應鏈", ai="受益", mapown=X)
add("TEAM", "Atlassian", "美", US, "DEV", 1, 1, "Jira／Confluence 開發協作龍頭，每席位計價", ai="侵蝕")
add("GTLB", "GitLab", "美", US, "DEV", 2, 2, "DevOps 平台", ai="分歧")
add("FROG", "JFrog", "美", US, "DEV", 3, 3, "軟體供應鏈與套件管理（受 AI 寫碼量增加帶動）", ai="受益")
# CRM
add("CRM", "Salesforce", "美", US, "CRM", 1, 1, "CRM 全球龍頭；Agentforce 與定價轉型為產業指標", ai="侵蝕")
add("HUBS", "HubSpot", "美", US, "CRM", 2, 2, "中小企業 CRM 與行銷自動化", ai="侵蝕")
add("NICE", "NICE", "美", US, "CRM", 2, 2, "客服中心軟體龍頭；客服代理直接衝擊", ai="侵蝕")
add("ZENDESK", "Zendesk", "未上市", PV, "CRM", 3, 3, "客服工單；已改按成果計價", ai="侵蝕")
add("FRSH", "Freshworks", "美", US, "CRM", 3, 3, "中小企業 CRM（印度創辦、研發在印度；美國總部、美股掛牌，故標否）", ai="侵蝕")
add("4443", "Sansan", "日", A, "CRM", 3, 3, "日本名片與客戶資料平台", ai="分歧")
# ERP
add("SAP", "SAP", "歐（德）", US, "ERP", 1, 1, "ERP 全球龍頭；雲端轉換中", ai="分歧")
add("ORCL-ERP", "Oracle（Fusion／NetSuite）", "美", US, "ERP", 1, 2, "雲端 ERP 前二；整家公司模型在 AI供應鏈", ai="分歧", mapown=X)
add("INTU", "Intuit", "美", US, "ERP", 1, 1, "中小企業會計與報稅龍頭", ai="分歧")
add("600588", "用友網絡", "陸", AS, "ERP", 2, 2, "中國 ERP 龍頭", ai="分歧")
add("0268.HK", "金蝶國際", "港", A, "ERP", 2, 2, "中國雲端 ERP 前二", ai="分歧")
add("4684", "OBIC", "日", A, "ERP", 2, 2, "日本中堅企業 ERP，高毛利", ai="分歧")
add("XRO", "Xero", "澳", A, "ERP", 2, 2, "中小企業雲端會計（澳紐英）", ai="分歧")
add("4478", "freee", "日", A, "ERP", 3, 3, "日本中小企業雲端會計", ai="分歧")
add("3994", "Money Forward", "日", A, "ERP", 3, 3, "日本雲端會計與金融資料", ai="分歧")
add("300378", "鼎捷數智", "陸", AS, "ERP", 3, 4, "製造業 ERP（台灣鼎新體系）", ai="分歧")
# HCM
add("WDAY", "Workday", "美", US, "HCM", 1, 1, "大型企業 HCM 龍頭，每員工計價", ai="侵蝕")
add("ADP", "ADP", "美", US, "HCM", 1, 1, "薪酬外包龍頭", ai="分歧")
add("6098", "Recruit 瑞可利（Indeed）", "日", A, "HCM", 1, 1, "全球線上招聘龍頭（Indeed、Glassdoor）＋日本人資", ai="侵蝕")
add("PAYX", "Paychex", "美", US, "HCM", 2, 2, "中小企業薪酬", ai="分歧")
add("BZ", "看準科技（BOSS直聘）", "美／港（2076）", A, "HCM", 2, 2, "中國線上招聘龍頭", ai="侵蝕")
add("SEK", "SEEK", "澳", A, "HCM", 3, 3, "澳紐招聘平台", ai="侵蝕")
add("3130", "104 人力銀行", "台", A, "HCM", 3, 3, "台灣招聘平台龍頭（國內）", ai="侵蝕")
# WFL
add("NOW", "ServiceNow", "美", US, "WFL", 1, 1, "ITSM 與企業工作流龍頭；代理調度層候選", ai="分歧")
add("PATH", "UiPath", "美", US, "WFL", 2, 2, "RPA 龍頭；被代理平台替代風險", ai="侵蝕")
add("MNDY", "monday.com", "美", US, "WFL", 3, 3, "工作管理", ai="侵蝕")
add("ASAN", "Asana", "美", US, "WFL", 3, 4, "工作管理", ai="侵蝕")
add("APPN", "Appian", "美", US, "WFL", 3, 4, "低程式碼流程平台", ai="分歧")
# PRD
add("MSFT-365", "Microsoft（365 Copilot）", "美", US, "PRD", 1, 1, "辦公生產力龍頭；整家公司模型在 AI供應鏈", ai="受益", mapown=X)
add("ADBE", "Adobe", "美", US, "PRD", 1, 1, "創意軟體龍頭；生成式工具直接競爭", ai="侵蝕")
add("ADSK", "Autodesk", "美", US, "PRD", 1, 1, "設計與 CAD 龍頭（工程資料護城河較深）", ai="分歧")
add("FIG", "Figma", "美", US, "PRD", 2, 2, "介面設計協作", ai="分歧")
add("CANVA", "Canva", "未上市", PV, "PRD", 2, 2, "大眾設計平台（澳洲）", ai="受益")
add("DOCU", "DocuSign", "美", US, "PRD", 2, 2, "電子簽章龍頭", ai="侵蝕")
add("BOX", "Box", "美", US, "PRD", 3, 3, "企業檔案管理", ai="分歧")
add("DBX", "Dropbox", "美", US, "PRD", 3, 3, "檔案同步", ai="侵蝕")
add("688111", "金山辦公", "陸", AS, "PRD", 2, 2, "中國辦公軟體龍頭（WPS）", ai="分歧")
add("4776", "Cybozu", "日", A, "PRD", 3, 3, "日本群組軟體", ai="侵蝕")
add("300624", "萬興科技", "陸", AS, "PRD", 3, 3, "影音與繪圖創意軟體", ai="侵蝕")
# VRT
add("TRI", "Thomson Reuters", "美／加", US, "VRT", 1, 1, "法律與稅務資訊龍頭（Westlaw）", ai="分歧")
add("RELX", "RELX（LexisNexis）", "歐（英）", US, "VRT", 1, 1, "法律、科學與風險資訊龍頭", ai="分歧")
add("SPGI", "S&P Global", "美", US, "VRT", 1, 1, "金融資料與評等", ai="分歧")
add("WKL", "Wolters Kluwer", "歐（荷）", US, "VRT", 2, 2, "醫療、稅務、法律資訊", ai="分歧")
add("VEEV", "Veeva", "美", US, "VRT", 1, 1, "生技醫藥垂直雲龍頭", ai="分歧")
add("CSU", "Constellation Software", "加", US, "VRT", 2, 2, "垂直軟體併購平台", ai="分歧")
add("LSEG", "London Stock Exchange Group（金融資料）", "歐（英）", US, "VRT", 2, 2, "金融資料與終端（與 S&P Global 同級）", ai="分歧")
add("TYL", "Tyler Technologies", "美", US, "VRT", 2, 2, "美國地方政府軟體", ai="分歧")
add("GWRE", "Guidewire", "美", US, "VRT", 2, 2, "產險核心系統", ai="分歧")
add("600570", "恒生電子", "陸", AS, "VRT", 2, 2, "中國金融軟體龍頭", ai="分歧")
add("2413", "M3", "日", A, "VRT", 2, 2, "日本醫師平台與醫藥行銷", ai="分歧")
# ITS
add("ACN", "Accenture", "美", US, "ITS", 1, 1, "全球 IT 服務龍頭；印度廠的對照指標", ai="侵蝕")
add("TCS", "Tata Consultancy Services", "印", IN, "ITS", 1, 2, "印度 IT 服務龍頭；AI 年化營收約 26 億美元（Q1 FY27，公司揭露）", "[Interested-party：公司揭露]", "侵蝕")
add("INFY", "Infosys", "印／美 ADR", IN, "ITS", 2, 2, "印度第二大；FY27 指引下修", "[Interested-party：公司揭露]", "侵蝕")
add("HCLTECH", "HCLTech", "印", IN, "ITS", 2, 3, "工程服務與軟體產品並重；管理層明言 AI 通縮", "[Interested-party：公司揭露]", "侵蝕")
add("WIPRO", "Wipro", "印／美 ADR", IN, "ITS", 3, 3, "印度前五", ai="侵蝕")
add("TECHM", "Tech Mahindra", "印", IN, "ITS", 3, 3, "電信垂直為主", ai="侵蝕")
add("LTIM", "LTIMindtree", "印", IN, "ITS", 3, 3, "印度前五大 IT 服務", ai="侵蝕")
add("CTSH", "Cognizant", "美", US, "ITS", 2, 3, "印度交付、美股掛牌", ai="侵蝕")
add("CAP", "Capgemini", "歐（法）", US, "ITS", 2, 3, "歐洲 IT 服務龍頭", ai="侵蝕")
add("IBM", "IBM（顧問）", "美", US, "ITS", 2, 3, "顧問與基礎軟體並重", ai="分歧")
add("EPAM", "EPAM", "美", US, "ITS", 3, 4, "工程外包", ai="侵蝕")
add("6702", "Fujitsu 富士通", "日", A, "ITS", 1, 2, "日本 IT 服務龍頭之一（區域龍頭）", ai="分歧")
add("NTTDATA", "NTT Data（NTT 子公司，2025-09 下市）", "未上市（母公司 9432）", PV, "ITS", 1, 2, "日本 IT 服務龍頭，海外營收大", ai="分歧")
add("4307", "野村總研 NRI", "日", A, "ITS", 2, 2, "金融系統與顧問，高毛利", ai="分歧")
add("6701", "NEC", "日", A, "ITS", 2, 2, "政府與社會基礎 IT", ai="分歧")
add("SCSK", "SCSK（住友商事子公司，2026-03 下市）", "未上市（母公司 8053）", PV, "ITS", 3, 3, "住友系 IT 服務；已下市，列為競爭參照", ai="分歧")
add("3626", "TIS", "日", A, "ITS", 3, 3, "金融與支付系統整合", ai="分歧")
add("018260", "Samsung SDS", "韓", A, "ITS", 2, 3, "韓國 IT 服務龍頭（集團內需為主）", ai="分歧")
add("064400", "LG CNS", "韓", A, "ITS", 3, 2, "韓國第二大 IT 服務", ai="分歧")
add("301236", "軟通動力", "陸", AS, "ITS", 3, 2, "中國 IT 外包龍頭之一", ai="侵蝕")
add("0354.HK", "中軟國際", "港", A, "ITS", 3, 3, "中國 IT 外包（華為為主要客戶）", ai="侵蝕")
add("6214", "精誠資訊", "台", A, "ITS", 3, 3, "台灣 IT 服務龍頭（國內）", ai="分歧")
# BPO
add("TEP", "Teleperformance", "歐（法）", US, "BPO", 1, 1, "全球客服中心龍頭", ai="侵蝕")
add("CNXC", "Concentrix", "美", US, "BPO", 2, 2, "客服與 CX 外包", ai="侵蝕")
add("G", "Genpact", "美", US, "BPO", 2, 2, "財務與後台流程外包（印度交付）", ai="侵蝕")
add("FSL", "Firstsource", "印", IN, "BPO", 3, 3, "印度 BPO", ai="侵蝕")
add("9715", "Transcosmos", "日", A, "BPO", 3, 3, "日本客服與數位行銷外包", ai="侵蝕")
# ADS
add("GOOGL", "Alphabet（搜尋與 YouTube 廣告）", "美", US, "ADS", 1, 1, "全球數位廣告龍頭；整家公司模型在 AI供應鏈", ai="分歧", mapown=X)
add("META", "Meta", "美", US, "ADS", 1, 2, "社群廣告龍頭；AI 推薦與自動投放受益；整家公司模型在 AI供應鏈", ai="受益", mapown=X)
add("BYTEDANCE", "字節跳動（抖音／TikTok）", "未上市", PV, "ADS", 1, 2, "中國與全球短影音廣告龍頭", ai="受益")
add("0700-ADS", "騰訊（社群廣告）", "港", A, "ADS", 1, 2, "微信生態廣告；整家公司模型預設在 AI供應鏈（第 3(b) 條）", ai="受益", mapown=X)
add("APP", "AppLovin", "美", US, "ADS", 2, 2, "行動廣告科技；AI 推薦引擎", ai="受益")
add("035420", "Naver", "韓", A, "ADS", 2, 2, "韓國搜尋龍頭；自建 HyperCLOVA X", ai="分歧")
add("4689", "LY Corp（LINE Yahoo）", "日", A, "ADS", 2, 2, "日本與台灣通訊＋搜尋廣告龍頭", ai="分歧")
add("1024.HK", "快手", "港", A, "ADS", 2, 3, "中國短影音第二；可靈影片模型", ai="受益")
add("035720", "Kakao", "韓", A, "ADS", 2, 3, "韓國通訊平台廣告", ai="分歧")
add("9888", "百度（搜尋廣告）", "港／美", A, "ADS", 2, 2, "中國搜尋廣告龍頭；整家公司模型預設在 AI供應鏈（自建崑崙晶片與雲）", ai="侵蝕", mapown=X)
add("TTD", "The Trade Desk", "美", US, "ADS", 2, 3, "獨立程式化廣告平台", ai="分歧")
add("RDDT", "Reddit", "美", US, "ADS", 3, 3, "社群；資料授權給模型商", ai="受益")
add("PINS", "Pinterest", "美", US, "ADS", 3, 3, "視覺搜尋與購物廣告", ai="分歧")
add("4751", "CyberAgent", "日", A, "ADS", 2, 3, "日本網路廣告代理龍頭＋遊戲＋ABEMA", ai="分歧")
# AGY
add("PUB", "Publicis", "歐（法）", US, "AGY", 1, 1, "廣告集團中數據與科技轉型最深", ai="分歧")
add("OMC", "Omnicom（含 IPG）", "美", US, "AGY", 1, 1, "2025-11 完成併購 IPG，全球最大廣告集團", ai="侵蝕")
add("WPP", "WPP", "歐（英）", US, "AGY", 2, 2, "老牌廣告集團，客戶流失", ai="侵蝕")
add("4324", "電通 Dentsu", "日", A, "AGY", 2, 2, "日本最大廣告代理；2026-02 放棄出售海外業務、自行整頓", ai="侵蝕")
add("2433", "博報堂 DY", "日", A, "AGY", 3, 3, "日本第二大廣告代理", ai="侵蝕")
# ECM
add("AMZN-EC", "Amazon（零售與廣告）", "美", US, "ECM", 1, 1, "全球電商龍頭；整家公司模型在 AI供應鏈", ai="分歧", mapown=X)
add("BABA-EC", "阿里巴巴（淘天）", "港／美", A, "ECM", 1, 1, "中國電商龍頭；整家公司模型預設在 AI供應鏈（第 3(b) 條）", ai="分歧", mapown=X)
add("PDD", "拼多多（含 Temu）", "美", ADR, "ECM", 1, 2, "中國第二大電商＋跨境", ai="分歧")
add("3690", "美團", "港", A, "ECM", 1, 2, "中國本地生活服務龍頭", ai="分歧")
add("SHOP", "Shopify", "美／加", US, "ECM", 2, 1, "電商開店平台龍頭；代理式購物介面", ai="受益")
add("MELI", "MercadoLibre", "美", US, "ECM", 2, 2, "拉美電商龍頭", ai="分歧")
add("BKNG", "Booking Holdings", "美", US, "ECM", 2, 1, "全球線上旅遊龍頭；被代理繞過風險", ai="分歧")
add("9618", "京東", "港／美", A, "ECM", 2, 2, "中國自營電商", ai="分歧")
add("SE", "Sea（Shopee）", "美", ADR, "ECM", 2, 2, "東南亞電商龍頭（新加坡）", ai="分歧")
add("CPNG", "Coupang", "美", ADR, "ECM", 2, 2, "韓國電商龍頭", ai="分歧")
add("GRAB", "Grab", "美", ADR, "ECM", 3, 3, "東南亞外送、叫車與金融超級應用", ai="分歧")
add("9961", "攜程 Trip.com", "港／美", A, "ECM", 2, 2, "亞洲線上旅遊龍頭", ai="分歧")
add("4755", "樂天集團", "日", A, "ECM", 2, 2, "日本電商龍頭之一（與 Amazon 日本並列）", ai="分歧")
add("4385", "Mercari", "日", A, "ECM", 3, 2, "日本二手市集龍頭", ai="分歧")
add("8454", "富邦媒（momo）", "台", A, "ECM", 3, 2, "台灣電商龍頭（國內）", ai="分歧")
# GAM
add("0700-GAM", "騰訊（遊戲）", "港", A, "GAM", 1, 1, "全球遊戲營收最大；整家公司模型預設在 AI供應鏈", ai="受益", mapown=X)
add("7974", "任天堂", "日", A, "GAM", 1, 1, "主機與 IP 龍頭", ai="受益")
add("6758-GAM", "Sony（PlayStation）", "日", A, "GAM", 1, 1, "主機平台龍頭；同一公司 CIS 列於實體 AI 地圖", ai="受益", mapown="跨圖（另見實體 AI 地圖；尚無個股模型，主圖待定）")
add("MSFT-GAM", "Microsoft（Xbox／Activision）", "美", US, "GAM", 1, 2, "全球前三遊戲發行商；整家公司模型在 AI供應鏈", ai="受益", mapown=X)
add("9999", "網易", "港／美", A, "GAM", 2, 2, "中國遊戲第二", ai="受益")
add("MIHOYO", "米哈遊", "未上市", PV, "GAM", 2, 2, "全球化手遊（原神）", ai="受益")
add("RBLX", "Roblox", "美", US, "GAM", 2, 2, "使用者創作平台；AI 創作工具受益", ai="受益")
add("259960", "Krafton", "韓", A, "GAM", 2, 3, "PUBG IP", ai="受益")
add("3659", "Nexon", "日", A, "GAM", 2, 3, "線上遊戲（韓系、東京掛牌）", ai="受益")
add("9697", "Capcom", "日", A, "GAM", 2, 3, "IP 與主機遊戲", ai="受益")
add("7832", "萬代南夢宮", "日", A, "GAM", 2, 3, "IP 授權與遊戲", ai="受益")
add("TTWO", "Take-Two", "美", US, "GAM", 2, 3, "GTA IP", ai="受益")
add("CDR", "CD Projekt", "歐（波蘭）", US, "GAM", 3, 3, "巫師、電馭叛客 IP；3A 單機開發成本最高的公司之一", ai="受益")
add("EA", "Electronic Arts（2026-08 完成私有化）", "未上市", PV, "GAM", 3, 3, "運動遊戲 IP；已下市，列為競爭參照", ai="受益")
add("U", "Unity", "美", US, "GAM", 3, 3, "遊戲引擎與廣告；AI 創作工具競爭", ai="分歧")
add("9766", "Konami", "日", A, "GAM", 3, 3, "IP 與手遊", ai="受益")
add("9684", "Square Enix", "日", A, "GAM", 3, 4, "IP 與 RPG", ai="受益")
add("036570", "NCSoft", "韓", A, "GAM", 3, 4, "MMORPG", ai="分歧")
add("251270", "Netmarble", "韓", A, "GAM", 3, 4, "手遊", ai="分歧")
add("3293", "鈊象", "台", A, "GAM", 3, 4, "台灣博弈類手遊，高毛利", ai="分歧")
# CNT
add("NFLX", "Netflix", "美", US, "CNT", 1, 1, "全球串流影音龍頭", ai="分歧")
add("SPOT", "Spotify", "美", US, "CNT", 2, 1, "音樂串流龍頭；AI 音樂供給衝擊", ai="分歧")
add("1698", "騰訊音樂", "港／美", A, "CNT", 2, 2, "中國音樂串流龍頭", ai="分歧")
add("9626", "嗶哩嗶哩", "港／美", A, "CNT", 3, 3, "中國年輕族群影音社群", ai="分歧")
add("WBTN", "Webtoon Entertainment", "美", ADR, "CNT", 3, 2, "網路漫畫平台龍頭（Naver 子公司）", ai="分歧")
add("9468", "KADOKAWA", "日", A, "CNT", 3, 3, "動漫與出版 IP（Sony 入股）", ai="分歧")
add("9602", "東寶 Toho", "日", A, "CNT", 3, 3, "動畫與電影 IP（哥吉拉、動畫製作）", ai="分歧")
add("352820", "HYBE", "韓", A, "CNT", 3, 3, "K-pop IP；AI 音樂與虛擬偶像", ai="分歧")
# AIN
add("ANYSPHERE", "Anysphere（Cursor）", "未上市", PV, "AIN", 2, 1, "AI 寫程式工具年化營收領先（媒體報導）", "[Interested-party：媒體]", "受益")
add("PERPLEXITY", "Perplexity", "未上市", PV, "AIN", 3, 2, "AI 搜尋；直接分流搜尋廣告", "[Interested-party：媒體]", "受益")
add("HARVEY", "Harvey", "未上市", PV, "AIN", 3, 2, "法律 AI；VRT 的替代方", "[Interested-party：媒體]", "受益")
add("SIERRA", "Sierra", "未上市", PV, "AIN", 3, 2, "客服代理、按成果計價；CRM 的替代方", "[Interested-party：媒體]", "受益")
add("DUOL", "Duolingo", "美", US, "AIN", 3, 2, "AI 驅動內容生產的消費應用", ai="分歧")
add("1357", "美圖", "港", A, "AIN", 3, 2, "AI 修圖訂閱（亞太少數已獲利 AI 應用）", ai="受益")
add("4180", "Appier 沛星", "日", A, "AIN", 3, 3, "AI 行銷 SaaS（台灣創辦、東京掛牌）", ai="受益")
add("3993", "PKSHA Technology", "日", A, "AIN", 3, 3, "日本 AI 代理與客服 AI", ai="受益")
add("6682", "第四範式", "港", A, "AIN", 3, 3, "中國企業 AI 平台", ai="受益")
add("AI", "C3.ai", "美", US, "AIN", 3, 4, "企業 AI 應用；成長停滯", ai="分歧")

wsc = wb.create_sheet("個股地位")
hdr2 = ["代號", "公司", "市場", "亞太掛牌", "層級", "次產業代碼", "次產業", "公司重要性（1–3）", "市場位置（1–4）",
        "地位依據", "標記", "AI 方向（公司，判斷，不進分數）", "次產業不可替代性", "價值鏈分量分", "不可替代性分", "市場位置分",
        "地位分數", "全表排名", "次產業內排名", "地圖歸屬"]
wsc.append(hdr2)
sub_rng = f"層級與次產業!$A$2:$A${last}"
def lk(col, i):
    return f"INDEX(層級與次產業!${col}$2:${col}${last},MATCH($F{i},{sub_rng},0))"
N = len(R)
for k, r in enumerate(R):
    i = k + 2
    t, name, mkt, apac, code, imp, pos, basis, tag, ai, mapown = r
    irrscore = lambda ref: f"CHOOSE({ref},{P['s_irr1']},{P['s_irr2']},{P['s_irr3']})"
    wsc.append([t, name, mkt, apac,
                "=" + lk("B", i), code, "=" + lk("C", i), imp, pos, basis, tag, ai,
                "=" + lk("G", i), "=" + lk("H", i),
                f"=({irrscore(f'M{i}')}+{irrscore(f'H{i}')})/2",
                f"=CHOOSE(I{i},{P['s_pos1']},{P['s_pos2']},{P['s_pos3']},{P['s_pos4']})",
                f"=ROUND({P['w_irr']}*O{i}+{P['w_pos']}*P{i}+{P['w_vc']}*N{i},1)",
                f"=RANK(Q{i},$Q$2:$Q${N+1},0)",
                f"=COUNTIFS($F$2:$F${N+1},F{i},$Q$2:$Q${N+1},\">\"&Q{i})+1",
                mapown])
    for col in (8, 9, 12):
        wsc.cell(i, col).font = BLUE
        wsc.cell(i, col).fill = YEL

# ---------- 次產業優先序 ----------
# 代碼, AI 衝擊強度(1–3, 1=最強), 核心命題代表性(1–3), 代表的單位經濟原型, 建議樣板公司, 建議批次(判斷), 理由
prio = {
    "ITS": (1, 1, "按人力計價 → 按成果計價（人力替代率）", "Accenture（對照）＋富士通（亞太樣板）", "第一批", "營收池最大；亞太高地位標的數居前，計入印度則最多；『人天通縮 vs AI 導入新需求』可用公司揭露的 AI 營收直接驗證"),
    "CRM": (1, 1, "按席位計價 → 按任務計價", "Salesforce", "第一批", "2026 年軟體賣壓的震央；席位→任務的定價轉換在此最先發生，樣板可複製到 HCM、WFL、PRD"),
    "ADS": (1, 2, "競價廣告 × AI 轉換率（受益原型）", "LY Corp（亞太樣板）；AppLovin（美，對照）", "第一批", "最大的受益營收池；Alphabet／Meta 已在 AI供應鏈，本專案以亞太平台為樣板可補地圖空白"),
    "GAM": (2, 2, "開發成本下降 × IP 定價權", "Nexon 或 Capcom", "第二批", "亞太權重最高、兩端重疊最多；但 AI 影響目前多在成本端，命題可證偽性較弱"),
    "VRT": (1, 2, "專有資料護城河 vs 模型重做", "Thomson Reuters 或 RELX", "第二批", "2026-02 急跌的次產業；結論分歧大、可下注的落差可能最大；亞太標的少"),
    "ERP": (2, 2, "記錄系統 × 代理取用費", "SAP（全球）／用友（亞太）", "第二批", "不可替代性最高的應用層；看代理是加深還是繞過記錄系統"),
    "SEC": (1, 1, "按防護範圍計價 × 非人類身分（代理）數量", "CrowdStrike", "第一批", "代理數量暴增使身分與端點防護面積放大；按防護範圍計價是其他樣板涵蓋不到的原型（專案負責人 2026-11-16 23:18 指出 v0.1 漏列）"),
    "DATA": (1, 1, "按用量計價 × 代理工作負載", "Snowflake", "第一批", "代理要讀寫資料，用量計價最直接受益；按用量計價是其他樣板涵蓋不到的原型；也是送 AI供應鏈 的 token 需求端證據來源（專案負責人 2026-11-16 23:18 指出 v0.1 漏列）"),
    "HCM": (2, 1, "按員工計價 × 招聘量", "Recruit（亞太）", "第二批", "直接讀人力替代率；Recruit 為亞太大型股、兩端重疊"),
    "PRD": (1, 1, "按席位計價（模型直接替代）", "Adobe", "第二批", "沿用 CRM 樣板；侵蝕最直接"),
    "WFL": (2, 2, "席位＋用量混合", "ServiceNow", "第三批", "沿用 CRM 樣板"),
    "BPO": (1, 1, "按人力時數計價", "Teleperformance", "第三批", "沿用 ITS 樣板；亞太標的少"),
    "AGY": (2, 2, "佣金 × 平台自動化", "電通", "第三批", "沿用 ITS 樣板；規模小"),
    "ECM": (2, 2, "抽成＋廣告 × 代理購物", "Sea 或樂天", "第三批", "代理購物仍早期，證據不足；亞太標的多，可列觀察"),
    "CNT": (2, 3, "訂閱 × 內容成本", "Netflix", "第三批", "AI 影響方向分歧、可量化錨點少"),
    "DEV": (2, 2, "席位 vs AI 寫碼", "Atlassian", "第三批", "與 AIN 一併看"),
    "AIN": (2, 2, "AI 原生毛利 vs 推論成本", "（多未上市；以 Duolingo／美圖觀察）", "第三批", "上市標的少；推論成本讀 AI供應鏈 Tokenomics"),
}
wsp = wb.create_sheet("次產業優先序")
hdr3 = ["代碼", "次產業", "AI 衝擊強度（1 最強–3）", "核心命題代表性（1 最高–3）", "亞太高地位標的數（是）", "同上＋印度與待核",
        "衝擊分", "代表性分", "重疊分", "重疊分（含印度與待核，敏感度）", "營收池分", "優先分數", "優先分數（敏感度）", "計算排名",
        "建議批次（判斷）", "代表的單位經濟原型", "建議樣板公司", "理由"]
wsp.append(hdr3)
NN = N + 1
for k, s in enumerate(subs):
    i = k + 2
    code = s[0]
    imp, rep, arche, sample, batch, why = prio[code]
    hi = f'COUNTIFS(個股地位!$F$2:$F${NN},A{i},個股地位!$D$2:$D${NN},"是*",個股地位!$Q$2:$Q${NN},">="&{P["ovl_cut"]})'
    hi2 = (hi + f'+COUNTIFS(個股地位!$F$2:$F${NN},A{i},個股地位!$D$2:$D${NN},"待*",個股地位!$Q$2:$Q${NN},">="&{P["ovl_cut"]})')
    ovl = lambda c: f"=IF({c}>=5,100,IF({c}>=3,75,IF({c}>=1,50,25)))"
    wsp.append([code, f"=INDEX(層級與次產業!$C$2:$C${last},MATCH(A{i},{sub_rng},0))", imp, rep,
                "=" + hi, "=" + hi2,
                f"=CHOOSE(C{i},100,60,30)", f"=CHOOSE(D{i},100,60,30)",
                ovl(f"E{i}"), ovl(f"F{i}"),
                f"=INDEX(層級與次產業!$H$2:$H${last},MATCH(A{i},{sub_rng},0))",
                f"=ROUND({P['p_imp']}*G{i}+{P['p_rep']}*H{i}+{P['p_ovl']}*I{i}+{P['p_pool']}*K{i},1)",
                f"=ROUND({P['p_imp']}*G{i}+{P['p_rep']}*H{i}+{P['p_ovl']}*J{i}+{P['p_pool']}*K{i},1)",
                f"=RANK(L{i},$L$2:$L${last},0)",
                batch, arche, sample, why])
    for col in (3, 4, 15):
        wsp.cell(i, col).font = BLUE
        wsp.cell(i, col).fill = YEL

# ---------- 觀察名單 ----------
wso = wb.create_sheet("觀察名單（未計分）")
wso.append(["代號", "公司", "市場", "亞太掛牌", "說明"])
for r in [
    ("DIS", "Disney", "美", "否", "全球最大內容 IP；非網路平台，暫列觀察"),
    ("OPENAI", "OpenAI（ChatGPT 消費與企業應用）", "未上市", "未上市", "應用層最大競爭者；模型權威在 AI供應鏈，本專案只讀其應用營收作競爭者定價"),
    ("ANTHROPIC", "Anthropic（Claude Code、Cowork）", "未上市", "未上市", "同上；2026-01 Cowork 為軟體股賣壓觸發點之一"),
    ("2513.HK", "智譜", "港", "是", "模型公司，歸 AI供應鏈"),
    ("MINIMAX", "MiniMax", "港", "是", "模型公司，歸 AI供應鏈"),
    ("ZM", "Zoom", "美", "否", "通訊協作，席位侵蝕案例"),
    ("INFA", "Informatica（已被 Salesforce 收購）", "—", "—", "資料整合併購案，作為 CRM 樣板的參照"),
]:
    wso.append(list(r))

for w in (wsb, wsc, wsp, wso):
    for c in w[1]:
        c.font = BOLD
        c.fill = HEAD
        c.alignment = Alignment(wrap_text=True, vertical="top")
    w.freeze_panes = "B2"
    for col in range(1, w.max_column + 1):
        w.column_dimensions[get_column_letter(col)].width = 14
for w, cols in ((wsb, {"C": 30, "D": 44, "E": 28, "K": 50}), (wsc, {"B": 30, "D": 22, "G": 28, "J": 48, "T": 40}),
                (wsp, {"B": 30, "P": 30, "Q": 34, "R": 70}), (wso, {"B": 34, "E": 70})):
    for k, v in cols.items():
        w.column_dimensions[k].width = v

wb.save(OUT)
# W2-6（2026-11-20）：openpyxl 寫出的公式格沒有快取值；以 LibreOffice 重算後原地存回，成品開檔即有值（tools/xlsx_recalc.py）
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "tools"))
import xlsx_recalc  # noqa: E402
xlsx_recalc.recalc_inplace(OUT)
print("rows", N, "subs", NSUB, OUT)
