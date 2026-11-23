# Cerebras Systems（CBRS）收支與估值模型 v0.1（AI 半導體 Project：運算層；格式比照 nvidia/ v0.3.4）
# 用法：python3 cbrs/build_cbrs.py raw.xlsx [--html] ；再以 LibreOffice headless 重算（CI：ai-semis-build）
#   --html：另寫 cbrs/html_src.json（v0.1.2 起；再由 cbrs/gen_html_json.py <xlsx> 轉成版面 v1.3 的 html.json）
# 會計年度＝日曆年（FY2026＝CY2026）。金額：十億美元（US$B）；每 MW 年營收：百萬美元／MW／年（＝US$B／GW／年）；股數：十億股。
# 營收＝推論雲端（平均計費 MW × 每 MW 年營收）＋系統（硬體）銷售；成本、費用、股票薪酬、資本支出（新增 MW × 每 MW 資本支出）→ EBIT、FCFF。
# 估值：標準版 DCF（tools/fcff_dcf.py v2）與 Comparable（同業前瞻 P/S；上市未滿一年、獲利為負，市場主流用營收倍數）各半。
import sys, os, json
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, os.path.join(ROOT, "tools", "in_layout"))
sys.path.insert(0, os.path.join(ROOT, "tools", "segment_check"))
import fcff_dcf as FD
import apply_in_layout as INL
from addrow import add_check

VER, VDATE = "v0.1.3", "2026-11-08"   # v0.1.2（2026-11-15）：版面 v1.3＋三層協調（v012_post.py）；v0.1.3（2026-11-15）：滾入 3Q26 實績與 4Q26 指引；VDATE＝v0.1 建檔日（版本沿革用）
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7"); KEY = PatternFill("solid", fgColor="FFF2CC"); SUB = PatternFill("solid", fgColor="EDEDED")
SC = ("A", "B", "C"); SCN = {"A": "樂觀", "B": "基準", "C": "保守"}
YRS = list(range(2025, 2033)); COL = {y: get_column_letter(3 + i) for i, y in enumerate(YRS)}   # C=FY2025A … J=FY2032
FC = list(range(2027, 2033))   # 預測年度（DCF 明確期）

TK = json.load(open(os.path.join(HERE, "data", "tokenomics_snapshot_v22.0.json"), encoding="utf-8"))
TKS = TK["source"]; TKV = f"Tokenomics {TKS['version']}（{TKS['file']}，commit {TKS['commit'][:7]}；快照 cbrs/data/tokenomics_snapshot_v22.0.json）"
TKI = TK["items"]; VR = "VR200 NVL72"
tkv = lambda k: TKI[k]["values"]
tkc = lambda k, i=1: TKI[k]["cells"][i] if isinstance(TKI[k]["cells"], list) else TKI[k]["cells"][VR]["基準"]

wb = openpyxl.Workbook()


def dn(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def hdr(ws, row, labels, c0=1):
    for j, h in enumerate(labels, c0):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR


# ---------------- 來源 ----------------
S8K2 = "Cerebras 8-K 2026-08-12（2Q26 財報新聞稿）https://www.sec.gov/Archives/edgar/data/0002021728/000162828026056186/cbrsannouncesfinancialresu.htm"
S8K1 = "Cerebras 8-K（1Q26 財報新聞稿）https://www.sec.gov/Archives/edgar/data/0002021728/000162828026044941/cbrsannouncesfinancialresu.htm"
S8K3 = "Cerebras 8-K 2026-11-12（3Q26 財報新聞稿）https://www.sec.gov/Archives/edgar/data/0002021728/000162828026073366/cbrsannouncesfinancialresu.htm"
SQ3 = "Cerebras 10-Q（季度結束 2026-09-30）https://www.sec.gov/Archives/edgar/data/0002021728/000162828026073494/cbrs-20260930.htm"
SQ2 = "Cerebras 10-Q（季度結束 2026-06-30）https://www.sec.gov/Archives/edgar/data/0002021728/000162828026056357/cbrs-20260630.htm"
SQ1 = "Cerebras 10-Q（季度結束 2026-03-31）https://www.sec.gov/Archives/edgar/data/0002021728/000162828026044981/cbrs-20260331.htm"
SAPI = "sec-api.io 10-Q／424B4 摘錄（https://sec-api.io/insights/cerebras-cbrs-first-public-quarters；二手彙整）"
SFUT = "Futurum S-1 拆解（2026-04-22，https://futurumgroup.com/insights/cerebras-s-1-teardown-is-the-23b-wafer-scale-ipo-the-end-of-gpu-homogeneity/；二手）"
SPX = "equibles.com CBRS 頁：2026-11-06 收盤 89.68、當日 +7.97（https://equibles.com/stocks/cbrs/summary；2026-11-08 讀取）"
SSA = "stockanalysis.com"

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億美元（US$B）；每 MW 年營收與每 MW 資本支出：百萬美元／MW（＝十億美元／GW）；股數：十億股。每個輸入附兩個來源（找不到者寫「未找到」）；公司原文算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間。公司原文＝Cerebras 8-K 財報新聞稿與 10-Q（sec.gov）；S-1／424B4 與 OpenAI 合約條款只經二手摘要（待人工複核）"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)

sec("A. 推論雲端（OpenAI 等）：營收＝平均計費 MW × 每 MW 年營收")
inp("p_mw", "每 MW 年營收（FY2026–FY2027 合約價；＝US$B／GW／年）", "US$M/MW/年", 5.33, "=oai_val/(oai_mw/1000)/oai_yr", 8.89, "Derived",
    "OpenAI 主合作協議（MRA）：750 MW、逾 200 億美元、期間三或四年可延至五年（S-1 經 " + SFUT + "；" + SAPI + "）",
    "＝200 億 ÷ 0.75 GW ÷ 3.5 年；低＝延至 5 年（5.33）、高＝3 年（8.89）。單價與 MW 口徑（IT 或設施）合約未公開：未找到")
for s, (g, note) in zip(SC, ((0.0, "每 MW 產出隨新世代 WSE 提升，抵銷每 token 降價；OpenAI 選擇權以原價續約"), (-0.04, "合約期內固定；續約與新客戶每 MW 價格溫和下降（每 token 降價大於每 MW 產出提升）"), (-0.10, "續約時大幅降價（GPU、ASIC 競爭）"))):
    inp(f"pg{s}", f"情境 {s}：FY2028 起每 MW 年營收年變化", "%", None, g, None, "Assumed", note,
        "每 MW 營收＝每 token 價格 × 每 MW 產出；參考：Tokenomics token 牌價年變化 −9%（v21.8 Cap_In，經 AI 投資地圖轉述；非 Interface 名稱，只作參考），產出隨世代提升未另計")
MW = {"A": (330, 800, 1300, 1700, 2000, 2200), "B": (320, 680, 900, 1050, 1200, 1300), "C": (250, 420, 520, 580, 600, 600)}
MWN = {"A": "OpenAI 750 MW 於 2028 全數上線、1.25 GW 選擇權 2029 起部分執行，AWS Bedrock 等客戶擴大",
       "B": "OpenAI 750 MW 分批上線（2028 年平均約九成，接近 RPO 認列時程）、其他客戶溫和成長；未執行選擇權",
       "C": "OpenAI 上線延後、不續約擴大，其他客戶有限"}
for s in SC:
    for y, v in zip(FC, MW[s]):
        inp(f"mw{s}_{y % 100}", f"情境 {s}：FY{y} 平均計費 MW", "MW", None, v, None, "Assumed",
            MWN[s] if y == 2027 else "同上",
            "公司：700 gross MW 已簽約、2027 年底前交付（" + S8K3 + "）；OpenAI 承諾 750 MW、選擇權再 1.25 GW" if y == 2027 else "同上")
inp("mw_k", "平均計費 MW 係數（Sens 用；FY2027 起乘在各情境 MW）", "x", 0.8, 1.0, 1.2, "Assumed", "上線時程、OpenAI 實際使用量的不確定", "±20%")
for s, (v, note) in zip(SC, ((0.72, "規模化、利用率高"), (0.66, "約等於 FY2026 錨點（指引推得）：機房租金與電力約占營收三分之一"), (0.58, "價格競爭、利用率不足"))):
    inp(f"gmc{s}", f"情境 {s}：雲端現金毛利率（不含折舊；FY2029 起，FY2027–FY2028 由 FY2026 線性收斂）", "%", None, v, None, "Assumed", note,
        "公司：1Q26 核心雲端毛利率 53%、2Q26 41.8%、3Q26 46.9%（含折舊；" + S8K1 + "；" + S8K2 + "；" + S8K3 + "）。雲端毛利（含折舊）＝雲端營收 × 本值 − 雲端折舊（PL 頁），故每 MW 資本支出同時影響毛利與現金流")
inp("gm_adj", "雲端現金毛利率加減（Sens 用；基準 0）", "pt", -0.05, 0.0, 0.05, "Assumed", "Sens：±5 個百分點", "—")
inp("capex_mw", "每新增 MW 資本支出（自製 WSE 系統成本＋自建機房部分；機房多為租賃）", "US$M/MW", 7.0, 9.0,
    "=tk_capit*tk_hold_cs*(1-hw_gm)", "Derived",
    "基準：FY2026 折舊攤銷 ÷ FY2026 平均計費 MW（PL 頁）× 5 年＝約 8.3；2026-09-30 不動產廠房設備淨額 1.508B（" + S8K3 + "）；v0.1.2 以 2026-06-30 淨額 0.99B ÷ 估計已上線約 100 MW（含在建）＝約 10；兩者之間取 9（基準值不變）。低 7 為 Assumed",
    "高＝Tokenomics 隱含：VR200 每 GW IT 資本支出 × Cerebras 持有成本比 ×（1 − 硬體毛利率）（自製系統以成本入帳；" + TKV + "）")
inp("life", "雲端系統折舊年限（亦為汰換週期）", "年", None, 5, None, "Assumed", "未找到（10-Q 會計政策未取得原文）", "汰換：新增 MW 的資本支出在 5 年後再發生一次（FY2032 汰換 FY2027 新增）")
inp("oth_cx", "其他資本支出 ÷ 營收（研發、製造、IT）", "%", 0.02, 0.03, 0.05, "Assumed", "未找到", "—")
inp("da_old", "FY2025 以前資產的折舊攤銷（FY2027–FY2030 每年）", "US$B", None, 0.02, None, "Assumed", "1H25 折舊攤銷 0.009（" + S8K2 + "）年化", "—")

sec("B. 系統（硬體）銷售")
for s, (g, note) in zip(SC, ((0.25, "MBZUAI、G42 以外的主權與企業客戶增加"), (0.05, "大致持平，雲端為主"), (-0.15, "UAE 客戶（MBZUAI、G42）採購遞減"))):
    inp(f"hg{s}", f"情境 {s}：硬體核心營收年增（FY2027 起）", "%", None, g, None, "Assumed", note,
        "客戶集中：FY2025 MBZUAI 62%、G42 24%；2Q26 MBZUAI 34%、G42 9%（" + SAPI + "）；3Q26 關係人客戶 A 38%、OpenAI 約 34%（" + SQ3 + "）")
inp("hw_gm", "硬體核心毛利率", "%", 0.35, 0.40, 0.45, "Assumed", "1Q26 核心硬體毛利率 42%、2Q26 38.8%、3Q26 31.0%（" + S8K1 + "；" + S8K2 + "；" + S8K3 + "）", "—")

sec("C. FY2026 錨點（1Q26–3Q26 實績＋4Q26 指引中點；Q 頁）")
inp("q4_rev", "4Q26 核心營收指引", "US$B", 0.310, 0.320, 0.330, "Interested-party", S8K3 + "：約 3.10–3.30 億美元", "3Q26 實績 0.2309（Q 頁 F20）")
inp("q4_gm", "4Q26 核心毛利率指引", "%", 0.47, 0.48, 0.49, "Interested-party", S8K3 + "：47%–49%", "3Q26 實績 41.0%")
inp("q4_om", "4Q26 核心營業利益率指引", "%", -0.01, 0.0, 0.01, "Interested-party", S8K3 + "：(1%)–1%", "3Q26 實績 (14.3%)")
inp("fy_rev", "FY2026 核心營收指引（11 月上修；只作對帳）", "US$B", 0.940, 0.950, 0.960, "Interested-party", S8K3 + "：9.40–9.60 億美元（8 月 8.80–8.90 億）", "1–3Q26 實績＋4Q26 指引中點＝0.952；共識 FY2026 營收 8.871 億（" + SSA + "，10-05；3Q26 前）")
inp("fy_gm", "FY2026 核心毛利率指引（只作對帳）", "%", 0.42, 0.43, 0.44, "Interested-party", S8K3 + "：42%–44%（8 月 41%–43%）", "1–3Q26 實績 42.5%＋4Q26 指引中點 48% 合成約 44.4%，高於全年指引上緣；兩組指引不一致，本模型以 4Q26 指引為錨點")
inp("fy_om", "FY2026 核心營業利益率指引（只作對帳）", "%", -0.10, -0.09, -0.08, "Interested-party", S8K3 + "：(10%)–(8%)（8 月 (19%)–(17%)）", "1–3Q26 實績 (11.1%)＋4Q26 指引中點 0% 合成約 (7.4%)，優於全年指引上緣")
inp("q4_hw", "4Q26 核心硬體營收", "US$B", 0.070, 0.085, 0.100, "Assumed", "公司未給分部指引；1Q26 0.112、2Q26 0.082、3Q26 0.086（" + S8K3 + "）", "雲端＝核心營收 − 硬體")
inp("q4_da", "4Q26 折舊攤銷", "US$B", None, 0.035, None, "Assumed", "3Q26 0.030（" + S8K3 + "）、隨新上線產能增加（沿用 v0.1.2 下半年 0.07 的每季值）", "—")
inp("q4_cx", "4Q26 資本支出", "US$B", 0.1, 0.5, 1.1, "Assumed", "沿用 v0.1.2 下半年假設 0.6／1.0／1.6 扣 3Q26 實績 0.497（9M26 1.046 − 1H26 0.549；" + S8K3 + "）", "公司：700 gross MW 已簽約、2027 年底前交付；尚未起租租約未折現約 85 億（" + SQ3 + "）")
inp("q4_oth", "4Q26 其他收益淨額（利息收入扣 OpenAI 貸款利息）", "US$B", None, 0.06, None, "Assumed", "3Q26 實績 0.064（" + S8K3 + "）；沿用 v0.1.2 下半年 0.12 的每季值", "—")
inp("cl_26", "FY2026 核心雲端營收（Q 頁）", "US$B", None, "=Q!H19", None, "Derived", "Q 頁：1Q26–3Q26 實績＋4Q26＝核心營收指引中點 − 硬體假設", "—")
inp("hw_26", "FY2026 核心硬體營收（Q 頁）", "US$B", None, "=Q!H18", None, "Derived", "Q 頁", "—")
inp("gp_26", "FY2026 核心毛利（Q 頁）", "US$B", None, "=Q!H21", None, "Derived", "Q 頁：1Q26–3Q26 實績＋4Q26 指引中點", "—")
inp("gmc_26", "FY2026 雲端現金毛利率（＝(核心毛利＋折舊攤銷 − 硬體毛利率 × 硬體營收) ÷ 雲端營收）", "%", None, "=(gp_26+da_26-hw_gm*hw_26)/cl_26", None, "Derived", "上列錨點與 hw_gm；FY2026 折舊攤銷全數視為雲端成本（Assumed）", "—")
inp("opx_26", "FY2026 核心營業費用（Q 頁）", "US$B", None, "=Q!H23", None, "Derived", "Q 頁：1Q26–3Q26 實績＋4Q26 指引中點", "—")
inp("sbc_26", "FY2026 經常性股票薪酬（不含 IPO 一次性認列 0.274 與其雇主薪資稅）", "US$B", None, "=Q!H29", None, "Derived", "Q 頁：1Q26–3Q26 實績＋4Q26＝3Q26 持平（Assumed）", "—")
inp("da_26", "FY2026 折舊攤銷（Q 頁）", "US$B", None, "=Q!H30", None, "Derived", "Q 頁", "—")
inp("cx_26", "FY2026 資本支出（9M26 實績＋4Q26 假設）", "US$B", None, "=Q!F48+q4_cx", None, "Derived", "Q 頁", "—")
inp("oth_26", "FY2026 其他收益淨額", "US$B", None, "=Q!H32", None, "Derived", "Q 頁", "—")

sec("D. 費用、股票薪酬、稅、股數（FY2027 起）")
inp("phi", "核心營業費用成長對營收成長的彈性", "x", 0.4, 0.6, 0.8, "Assumed／Derived", "2Q26 核心營業費用年增 56%、核心營收年增 103%（" + S8K2 + "）比值約 0.55", "—")
inp("g_opx_min", "核心營業費用年增下限", "%", None, 0.05, None, "Assumed", "研發人力持續擴充", "—")
inp("g_opx_max", "核心營業費用年增上限", "%", 0.4, 0.6, 0.8, "Assumed", "4Q26 核心營業費用 0.1536、年化約 6.1 億（Q 頁 G23），FY2026 合計約 4.93 億（Q 頁 H23）；高成長年度費用不與營收同比例增加", "—")
inp("sbc_27", "股票薪酬 ÷ 營收：FY2027", "%", 0.10, 0.14, 0.18, "Derived",
    "2Q26 股票薪酬 3.77 億扣 IPO 一次性 2.74 億＝每季約 1.03 億，年化 4.1 億 ÷ FY2027 營收約 29 億（情境 B）", "低／高 Assumed")
inp("sbc_ss", "股票薪酬 ÷ 營收：FY2032（FY2028–FY2032 由 FY2027 線性收斂）", "%", 0.04, 0.06, 0.08, "Assumed", "大型半導體同業約 3%–8%（NVIDIA 約 4%、AMD 約 5%：本專案模型 In 頁）", "—")
inp("oth_27", "其他收益淨額 FY2027（利息收入扣利息費用）", "US$B", None, 0.15, None, "Assumed", "現金隨資本支出下降", "—")
inp("oth_28", "其他收益淨額 FY2028", "US$B", None, 0.05, None, "Assumed", "—", "—")
inp("oth_f", "其他收益淨額 FY2029 起每年", "US$B", None, 0.0, None, "Assumed", "—", "—")
inp("tx_e", "有效稅率 FY2026–FY2028（虧損與營業虧損扣抵）", "%", None, 0.0, None, "Assumed", "FY2025 前累積虧損 9.05 億（2025-12-31 累積虧損，" + S8K2 + "）", "—")
inp("tx_f", "有效稅率 FY2029 起", "%", 0.10, 0.15, 0.21, "Assumed", "美國法定 21%；扣抵與研發抵減使近期偏低", "—")
inp("sh_g", "稀釋股數年增（新授予股票報酬；FY2027 起）", "%", 0.0, 0.02, 0.04, "Assumed", "未找到", "評價用股數 sh_v 已含既有選擇權、限制股與認股權證")

sec("E. 公司揭露、合約、共識與 Tokenomics（只作 Recon 對帳，不作校準）")
inp("oai_val", "OpenAI 主合作協議金額（逾）", "US$B", None, 20, None, "Interested-party", SFUT, "letsdatascience、devtake 轉述（10B 為首次公布值、S-1 時倍增至 20B；二手）")
inp("oai_mw", "OpenAI 承諾容量", "MW", None, 750, None, "Interested-party", SAPI + "：750 MW 承諾＋1.25 GW 選擇權（合計 2 GW）", SFUT)
inp("oai_opt", "OpenAI 選擇權容量", "MW", None, 1250, None, "Interested-party", SAPI, "—")
inp("oai_yr", "OpenAI 合約期間（三或四年，可延至五年）", "年", 3, 3.5, 5, "Interested-party", SAPI, "基準取三、四年中點")
inp("rpo", "剩餘履約義務 RPO（2026-09-30）", "US$B", None, 26.1, None, "Verified", S8K3 + "：261 億美元", "2026-06-30 254 億（" + S8K2 + "）；2025-12-31 246 億、2026-03-31 250 億（" + SAPI + "）")
inp("rpo_24", "RPO 中 24 個月內認列比例（2026-09-30 起算，至 2028-09-30）", "%", None, 0.33, None, "Verified", SQ3 + "：約 33%", "2026-06-30：22%（" + SAPI + "）")
inp("rpo_48", "RPO 中第 25–48 個月認列比例", "%", None, 0.44, None, "Verified", SQ3 + "：約 44%", "其餘約 23% 在 48 個月後；2026-06-30：43%（" + SAPI + "）")
inp("mw_co27", "公司：已簽約、2027 年底前交付的資料中心容量（gross MW）", "MW", None, 700, None, "Interested-party", S8K3 + "：700 gross MW", "2Q26：逾 600 MW 上線或簽約（" + S8K2 + "）；gross 口徑可能為設施電力")
inp("cfo_x27", "公司：2027 年營收三倍（倍數）", "x", None, 3, None, "Interested-party", S8K3 + "：執行長「expect to triple revenue in 2027」", "2Q26：財務長逾三倍（" + S8K2 + "）")
inp("cons_rev26", "共識 FY2026 營收", "US$B", 0.8849, 0.8871, 0.8911, "Interested-party", SSA + " 預測頁（https://stockanalysis.com/stocks/cbrs/forecast/；分析師資料 2026-10-05 更新）", "FY2027、FY2028 共識：未找到（付費）")
inp("cons_eps26", "共識 FY2026 調整後 EPS", "US$", -0.71, -0.47, -0.16, "Interested-party", SSA + " 預測頁（同上）", "口徑為公司核心（不含股票薪酬），與本模型不同")
inp("cons_tp", "市場共識目標價（賣方平均，11 位）", "US$", 175, 282.55, 330, "Interested-party", SSA + " 預測頁（分析師資料 2026-10-05 更新；股價當時約 165–200）", "股價其後跌至 80–110，共識目標價可能尚未更新；只作對照")
inp("tk_out_cs", "Tokenomics：Cerebras WSE-3 每 GW 產出 ÷ VR200", "比", tkv("IF_NonNV_Out_CS")[0], tkv("IF_NonNV_Out_CS")[1], tkv("IF_NonNV_Out_CS")[2], "Derived",
    TKV + "：IF_NonNV_Out_CS（" + "、".join(TKI["IF_NonNV_Out_CS"]["cells"]) + "；Tokenomics 標記 Assumed）", "只作 Recon")
inp("tk_hold_cs", "Tokenomics：Cerebras WSE-3 每 GW 持有成本 ÷ VR200", "比", tkv("IF_NonNV_Hold_CS")[0], tkv("IF_NonNV_Hold_CS")[1], tkv("IF_NonNV_Hold_CS")[2], "Derived",
    TKV + "：IF_NonNV_Hold_CS（" + "、".join(TKI["IF_NonNV_Hold_CS"]["cells"]) + "；Tokenomics 標記 Assumed）", "只作 Recon 與 capex_mw 高值")
hv = tkv("IF_HoldEcon")[VR]; cv = tkv("IF_CapexIT")[VR]
inp("tk_hold_vr", "Tokenomics：VR200 每 GW 年持有成本（經濟口徑）", "US$B/GW/年", hv["低成本"], hv["基準"], hv["高成本"], "Derived",
    TKV + "：IF_HoldEcon｜VR200（" + TKI["IF_HoldEcon"]["cells"][VR]["基準"] + "）", "低／高＝低成本／高成本情境")
inp("tk_capit", "Tokenomics：VR200 每 GW IT 資本支出", "US$B/GW", cv["低成本"], cv["基準"], cv["高成本"], "Derived",
    TKV + "：IF_CapexIT｜VR200（" + TKI["IF_CapexIT"]["cells"][VR]["基準"] + "）", "—")

sec("F. 價格與估值")
inp("px", "股價（2026-11-20 收盤）", "US$", None, 82.67, None, "Verified", "工作單 px1120；Yahoo Finance 日線收盤 2026-11-20（markets/invest_map/data/closes_20261109_20261120.csv）",
    "v0.1.4 為 2026-11-13 收盤 108.72；v0.1.3 為 2026-11-05 收盤 81.71（" + SPX + "：11-05 收盤＝89.68 − 7.97＝81.71；Yahoo Finance 日線 11-05 收盤同為 81.71）。equibles 2026-11-04 收盤 91.85；stockanalysis.com 歷史股價 10-30 收盤 109.18。11-09 為 IPO 鎖定期最後解禁日、11-12 公布 3Q26 財報")
for s, w in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重 {s}", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 0.3／0.5／0.2；同 NVIDIA、AMD）", "三權重合計應為 1")
inp("tp_w_dcf", "目標價中 DCF 的權重（其餘為 Comparable）", "%", None, 0.5, None, "Assumed", "Andy 可改（Andy 2026-10-25 11:53 定 50／50）", "Sens 頁 0.3／0.7")

for _row in FD.in_rows(dict(
        date="2026-12-31", amt="US$B", shu="十億股",
        rf=(0.0500, "美國財政部每日殖利率曲線 10 Yr：2026-10-23 5.00%（同 NVIDIA、AMD v0.3 起各模型）", "FRED DGS10（同一序列）", "無風險利率（美國 10 年期公債殖利率，2026-10-23）"),
        beta_fixed=("=(beta_u-beta_du)*(1+(1-tax_w)*beta_de)", "=beta_u*(1+(1-tax_w)*beta_de)", "=(beta_u+beta_du)*(1+(1-tax_w)*beta_de)", "Derived",
                    "自算產業 β：US 半導體 同業 13 家中位數（docs/plan/自算產業β_20261105.md），以本公司市值 D/E 與稅率重新加槓桿（βL＝βU ×（1＋(1 − tax_w) × D/E））",
                    "本公司上市未滿一年，個股迴歸 β 不採用。Sens 低／高：β_u ∓ beta_du（0.3），各自同式重新加槓桿"),
        kd=(0.06, "Verified", SAPI + "：OpenAI 營運資金貸款年利率 6%（2026-01 撥款 10.05 億，最晚 2032-12-31 到期）", "循環信用額度 8.5 億未動用"),
        tax_w=("=tx_f", "Derived", "＝In 頁 tx_f（FY2029 起有效稅率）", "NOPAT 用 PL 頁各年稅率（FY2027–FY2028 為 0）"),
        debt=("=Q!F42", "Verified", S8K3 + "：OpenAI 營運資金貸款 流動 7.07 億＋非流動 1.23 億（2026-09-30；本金 8.291 億，" + SQ3 + "）", "Q 頁 F42"),
        lease=("=Q!F43", "Verified", S8K3 + "：營業租賃負債 流動 0.98 億＋非流動 5.23 億（2026-09-30）", "尚未起租租約未折現最低付款約 85 億（" + SQ3 + "）尚未入帳，以資本支出與營業成本反映"),
        cash=("=Q!F35+Q!F36", "Verified", S8K3 + "：現金及約當現金 30.75 億＋短期投資 42.86 億（2026-09-30）", "Q 頁 F35、F36"),
        inv=("=Q!F37", "Verified", S8K3 + "：受限現金 4.44 億（流動；列為非營業投資，Sens 測不計）", "受限用途：未找到原文說明"),
        mi=(0.0, "Verified", "無少數股權", "—"),
        stub=dict(bs_date="2026-09-30", val_date="2026-12-31", bs_src=S8K3 + "（最新已公布一季 3Q26）",
                  div=(0.0, "Verified", "未配息；2026-09-30～2026-12-31 無已宣告股利或買回", "—")),
        sh=("=Q!F54+Q!F55+Q!F56", "Derived", "已發行 2.4300 億股（2026-11-05，A＋B＋N 類）＋潛在稀釋 0.7233 億（2026-09-30）（" + SQ3 + "）；v0.1.2 為 2.3756＋0.7601＋0.1003 億",
            "全部視為稀釋（選擇權履約價多低於現價；客戶認股權證對應的營收已在 FCFF 內，故其攤銷不再從 EBIT 扣除）"),
        fc_da=(None, 0.10, None, "Assumed", "本模型逐年讀 PL 頁折舊攤銷列（fcff_dcf v2 da 參數）；本列只作模組預設", "—"),
        fc_cx=(None, 0.15, None, "Assumed", "本模型逐年讀 PL 頁資本支出列（cx 參數）；本列只作模組預設", "Sens：資本支出 ÷ 營收 ±1pt（fc_cx_d）"),
        fc_wc=(0.0, 0.05, 0.15, "Assumed", "客戶訂金與遞延收入（2026-09-30 合計 7.96 億）使營運資金需求低；應收 2.32 億、存貨 1.46 億（" + S8K3 + "）", "營運資金增加＝本值 × 營收增量"),
        )):
    rows.append(_row)
inp("beta_u", "產業無槓桿 β（採用）：自算 US 半導體 同業中位數 Blume 調整（0.67 × 中位數＋0.33）", "x", None, 1.884, None, "Analogy",
    "自算產業 β：US 半導體 同業 13 家無槓桿 β 中位數 2.32 → Blume 調整 1.884（docs/plan/自算產業β_20261105.md；同 NVIDIA、AMD）", "Cerebras 一半以上營收為推論雲端；新雲同業 β 未另算")
inp("beta_du", "β 敏感度幅度（低／高＝β_u ∓ 本值，各自重新加槓桿）", "x", None, 0.3, None, "Assumed", "工作單第二輪修正（Andy 2026-11-05 23:16 定案）", "—")
inp("beta_de", "市值負債權益比 D/E（重新加槓桿用）", "x", None, "=(debt_v+lease_v)/(px*sh_v)", None, "Derived", "＝(債務本金 ＋ 租賃) ÷ (股價 × 評價股數)", "—")

IKEY = {}
r = 4
for it in rows:
    if it[0] == "sec":
        IN.cell(r, 1, it[1]).font = BOLD; r += 1; continue
    k, lab, unit, lo, b, hi, tag, s1, s2 = it
    IN.cell(r, 1, k); IN.cell(r, 2, lab); IN.cell(r, 3, unit)
    for c, v in ((4, lo), (5, b), (6, hi)):
        if v is not None:
            IN.cell(r, c, v).font = BLUE
    IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2)
    dn(k, f"In!$E${r}")
    if lo is not None: dn(k + "_lo", f"In!$D${r}")
    if hi is not None: dn(k + "_hi", f"In!$F${r}")
    IKEY[k] = r; r += 1
IN.cell(IKEY["stub_frac"], 9, "本模型 stub FCFF＝4Q26 FCFF（Q 頁 H33，路徑自訂 stub_fcff），不用第一年 FCFF × 本值；同 README、DCF 頁 A2")
FD.yellow(IN)
for k in ("tp_wA", "tp_wB", "tp_wC", "tp_w_dcf"):
    for c in range(1, 10): IN.cell(IKEY[k], c).fill = KEY
for c, w in zip("ABCDEFGHI", (14, 62, 12, 10, 12, 10, 14, 70, 60)):
    IN.column_dimensions[c].width = w

# ---------------- Q（季度實績、指引、資產負債；US$B） ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 季度實績與指引（US$B；公司原文：8-K 財報新聞稿、10-Q）。核心（Core）＝公司 non-GAAP：排除客戶認股權證攤銷（沖減營收）、代墊資料中心收入與成本、股票薪酬、IPO 相關雇主薪資稅"
Q["A2"] = "1Q26–3Q26 為實績；4Q26＝4Q26 指引中點（3Q26 財報 2026-11-12；硬體、折舊、資本支出、其他收益為 Assumed）。v0.1.3 起資產負債表、現金流（9M26）、股數改到 3Q26 欄（2026-09-30）"
hdr(Q, 3, ["項目", "FY2024", "FY2025", "1Q26", "2Q26", "3Q26", "4Q26E（指引）", "FY2026E", "標記", "來源"])
def qrow(rr, label, vals, tag="Verified", src="", fmt="#,##0.000"):
    Q.cell(rr, 1, label)
    for c, v in zip("BCDEFGH", vals):
        if v is not None:
            x = Q[f"{c}{rr}"]; x.value = v; x.number_format = fmt
            if not (isinstance(v, str) and v.startswith("=")): x.font = BLUE
    Q.cell(rr, 9, tag); Q.cell(rr, 10, src)
Q["A4"] = "GAAP（公司財報）"; Q["A4"].font = BOLD
qrow(5, "硬體營收（GAAP）", [None, None, 0.110593, 0.054119, 0.068881, None, None], src=SQ1 + "；" + S8K2 + "；" + S8K3)
qrow(6, "雲端與其他服務營收（GAAP）", [None, None, 0.082813, 0.125991, 0.141928, None, None], src=SQ1 + "；" + S8K2 + "；" + S8K3)
qrow(7, "營收合計（GAAP）", [0.2903, 0.5100, 0.193406, 0.180110, 0.210809, None, None], src="FY2024、FY2025：424B4（" + SAPI + "）；季：10-Q；3Q26：" + S8K3)
qrow(8, "毛利（GAAP）", [0.1227, 0.1991, 0.086176, 0.025559, 0.042094, None, None], src="1Q26＝1H26 111.735 − 2Q26 25.559（百萬美元，" + S8K2 + "）；3Q26 42.094（" + S8K3 + "）")
qrow(9, "營業損益（GAAP）", [-0.1014, -0.1459, -0.015037, -0.477233, -0.224149, None, None], src="1Q26＝1H26 −492.270 − 2Q26 −477.233；3Q26 −224.149")
qrow(10, "股票薪酬（GAAP，含 IPO 一次性）", [0.0586, 0.0498, 0.009593, 0.377008, 0.155105, None, None], src="1Q26＝1H26 386.601 − 2Q26 377.008；3Q26 155.105（9M26 541.706 − 1H26）")
qrow(11, "其中：IPO 時一次認列的限制股費用", [None, None, None, 0.2736, None, None, None], src=SAPI + "：2026-05-13 IPO 認列 2.736 億")
qrow(12, "IPO 股票薪酬的雇主薪資稅", [None, None, None, 0.022778, 0.0, None, None], src=S8K2 + "；3Q26 無（9M26 22.778 全在 1H26，" + S8K3 + "）")
qrow(13, "客戶認股權證攤銷（沖減營收；至 2031-10）", [None, None, 0.002053, 0.044262, 0.036424, None, None], src=S8K2 + "；1Q26＝1H26 46.315 − 2Q26 44.262；3Q26 36.424")
qrow(14, "代墊資料中心收入（pass-through）", [None, None, 0.004111, 0.014503, 0.016360, None, None], src=S8K2 + "；1Q26＝1H26 18.614 − 2Q26 14.503；3Q26 16.360")
qrow(15, "淨利（GAAP）", [-0.4816, 0.2378, -0.014006, -0.450528, -0.161218, None, None], src="FY2025 含遠期合約負債消滅利益 3.633 億（一次性）；3Q26 −161.218（" + S8K3 + "）")
qrow(16, "其他收益淨額（GAAP）", [None, None, 0.002528, 0.026979, 0.063900, None, None], src="1Q26＝1H26 29.507 − 2Q26 26.979；3Q26 63.900")
Q["A17"] = "核心（non-GAAP；公司口徑）"; Q["A17"].font = BOLD
qrow(18, "核心硬體營收", [None, None, 0.1116, 0.082141, 0.085959, "=q4_hw", "=SUM(D18:G18)"], src=S8K1 + "（111.6）；" + S8K2 + "；" + S8K3 + "；4Q26＝q4_hw（Assumed）")
qrow(19, "核心雲端營收", [None, None, 0.0798, 0.127728, 0.144914, "=G20-G18", "=SUM(D19:G19)"], src="3Q26 144.914（" + S8K3 + "）；4Q26＝核心營收 − 硬體（Derived）")
qrow(20, "核心營收合計", [None, None, 0.191348, 0.209869, 0.230873, "=q4_rev", "=SUM(D20:G20)"], src="1Q26＝1H26 401.217 − 2Q26 209.869；3Q26 230.873；4Q26 指引中點。本列 FY2026E 0.95209；PL 頁 FY2026 營收＝雲端＋硬體＝0.952142，差 0.05 百萬美元來自 1Q26 核心硬體 111.6（新聞稿四捨五入）＋雲端 79.8 與核心合計 191.348 之差，不影響估值")
qrow(21, "核心毛利", [None, None, 0.089058, 0.085217, 0.094575, "=G20*q4_gm", "=SUM(D21:G21)"], src="1Q26＝1H26 174.275 − 2Q26 85.217；3Q26 94.575；4Q26＝指引中點")
qrow(22, "核心毛利率", [None, None, "=D21/D20", "=E21/E20", "=F21/F20", "=G21/G20", "=H21/H20"], tag="Derived", fmt="0.0%")
qrow(23, "核心營業費用", [None, None, 0.09257, 0.11883, 0.127683, "=G21-G24", "=SUM(D23:G23)"], src="1Q26＝1H26 211.400 − 2Q26 118.830；3Q26 127.683")
qrow(24, "核心營業損益", [None, None, -0.003512, -0.033613, -0.033108, "=G20*q4_om", "=SUM(D24:G24)"], src="1Q26＝1H26 −37.125 − 2Q26 −33.613；3Q26 −33.108；4Q26＝指引中點")
qrow(25, "核心營業利益率", [None, None, "=D24/D20", "=E24/E20", "=F24/F20", "=G24/G20", "=H24/H20"], tag="Derived", fmt="0.0%")
qrow(26, "核心硬體毛利", [None, None, "=0.42*D18", 0.031878, 0.026672, None, None], tag="Verified／Derived", src="1Q26＝核心硬體毛利率 42%（四捨五入）× 營收；2Q26 31.878；3Q26 26.672（31.0%）")
qrow(27, "核心雲端毛利", [None, None, "=D21-D26", 0.053339, 0.067903, None, None], tag="Verified／Derived", src="1Q26＝核心毛利 − 硬體；2Q26 53.339；3Q26 67.903（46.9%）")
Q["A28"] = "錨點推導（進 PL 頁 FY2026）"; Q["A28"].font = BOLD
qrow(29, "經常性股票薪酬（不含 IPO 一次性）", [None, None, "=D10", "=E10-E11", "=F10", "=F29", "=SUM(D29:G29)"], tag="Derived", src="3Q26＝GAAP 股票薪酬（無 IPO 一次性項目）；4Q26＝3Q26 持平（Assumed）")
qrow(30, "折舊攤銷", [None, None, 0.018174, 0.024377, 0.029794, "=q4_da", "=SUM(D30:G30)"], src="2Q26 24.377；1Q26＝1H26 42.551 − 24.377；3Q26 29.794（" + S8K3 + "）；4Q26 q4_da")
qrow(31, "資本支出（季別未揭露；累計見 E48、F48）", [None, None, None, None, "=F48-E48", "=q4_cx", "=F48+G31"], tag="Verified／Assumed", src="3Q26＝9M26 − 1H26（Derived）；4Q26 q4_cx")
qrow(32, "其他收益淨額", [None, None, "=D16", "=E16", "=F16", "=q4_oth", "=SUM(D32:G32)"], tag="Verified／Assumed", src="4Q26 q4_oth")
qrow(33, "4Q26 FCFF（＝stub 期 FCFF，評價日前）", [None, None, None, None, None, None,
     "=(G24-G29)*(1-tx_e)+G30-G31-fc_wc*(G20-F20)"], tag="Derived",
     src="＝(核心營業損益 − 經常性股票薪酬) ×（1 − 稅率）＋折舊 − 資本支出 − 營運資金增加（4Q26；v0.1.2 為 3Q26＋4Q26）")
Q["A34"] = "資產負債表（2026-09-30；3Q26 欄。2Q26 欄＝2026-06-30，v0.1.2 值）"; Q["A34"].font = BOLD
for rr, lab, v2, v, src in ((35, "現金及約當現金", 6.742157, 3.074752, S8K3), (36, "短期投資", 1.179390, 4.286026, S8K3), (37, "受限現金", 0.684680, 0.444042, S8K3 + "（流動）"),
                        (38, "應收帳款", 0.123361, 0.231762, S8K3), (39, "存貨", 0.032666, 0.146100, S8K3), (40, "客戶訂金", 0.242672, 0.108715, S8K3),
                        (41, "遞延收入（流動＋非流動）", 0.418264, 0.687602, S8K3 + "：234.309＋453.293"), (42, "OpenAI 營運資金貸款（流動＋非流動）", 0.918244, 0.829082, S8K3 + "：706.518＋122.564"),
                        (43, "營業租賃負債（流動＋非流動）", 0.568827, 0.621139, S8K3 + "：98.368＋522.771"), (44, "股東權益", 9.154642, 9.158913, S8K3)):
    qrow(rr, lab, [None, None, None, v2, v, None, None], src=src + "；2Q26 欄：" + S8K2)
qrow(45, "剩餘履約義務 RPO", [None, 24.6, 25.0, 25.4, 26.1, None, None], src=S8K3 + "；" + S8K2 + "；" + SAPI + "（FY2025 欄＝2025-12-31、1Q26＝2026-03-31）", fmt="#,##0.0")
Q["A46"] = "現金流量（FY；1H26 填於 2Q26 欄、9M26 填於 3Q26 欄）"; Q["A46"].font = BOLD
qrow(47, "營業現金流", [0.452, -0.0101, None, -0.047488, -0.388894, None, None], src=SAPI + "；" + S8K2 + "；9M26：" + S8K3)
qrow(48, "資本支出", [None, 0.3827, None, 0.548873, 1.046059, None, None], src=SAPI + "；" + S8K2 + "；9M26：" + S8K3)
Q["A49"] = "客戶集中（占營收）"; Q["A49"].font = BOLD
qrow(50, "MBZUAI", [None, 0.62, 0.63, 0.34, 0.38, None, None], tag="Verified／Derived", src=SAPI + "（10-Q 風險因子）；3Q26＝10-Q「客戶 A」（關係人）38%，對應前期 MBZUAI（Derived；" + SQ3 + "）", fmt="0%")
qrow(51, "G42", [0.85, 0.24, 0.11, 0.09, None, None, None], tag="Verified", src=SAPI + "；3Q26：10-Q 關係人「客戶 C」低於 10%（未揭露確數）", fmt="0%")
qrow(52, "OpenAI 營收（US$B；當季）", [None, 0, None, 0.0568, 0.0724, None, None], tag="Verified", src=SAPI + "：2Q26 第三大客戶占 32%；3Q26 72.4 百萬（扣認股權證攤銷 3.2 百萬後；9M26 146.8；" + SQ3 + "）")
Q["A53"] = "股數（十億股；v0.1.3 起讀 3Q26 欄）"; Q["A53"].font = BOLD
qrow(54, "已發行普通股（A＋B＋N 類）", [None, None, None, 0.237564041, 0.243003364, None, None], src="2Q26 欄 2026-08-05（" + SAPI + "）；3Q26 欄 2026-11-05：A 160,084,168＋B 69,203,688＋N 13,715,508（" + SQ3 + " 封面）", fmt="0.0000")
qrow(55, "潛在稀釋證券", [None, None, None, 0.076011, 0.072326, None, None], src="2Q26 欄 2026-06-30（" + SAPI + "）；3Q26 欄 2026-09-30：或有認股權證 26.108＋選擇權 19.503＋未既得 RSU 17.139＋PRSU 9.000＋提前履約 0.481＋ESPP 0.095 百萬股（" + SQ3 + "）", fmt="0.0000")
qrow(56, "N 類認股權證（2026-07 發行）", [None, None, None, 0.010033508, 0.0, None, None], src="2Q26 欄：" + SAPI + "；3Q26 欄 0：2026-09-30 已含在上兩列（已發行 N 類或或有認股權證），不重複計入", fmt="0.0000")
SEGR = add_check(Q, seg_rows=[5, 6], total_row=7, cols="DEF", note="官方分部（GAAP）：硬體、雲端與其他服務（segments.json）")
for c, w in zip("ABCDEFGHIJ", (52, 10, 10, 10, 10, 12, 12, 11, 14, 90)):
    Q.column_dimensions[c].width = w

# ---------------- 收支引擎（PL 頁三情境；Sens 頁案例共用） ----------------
ENGINE_ROWS = 31


def engine(ws, r0, s, title, over=None):
    """寫一個情境區塊（C=FY2025A … J=FY2032）。over：{參數鍵: 取代的名稱或公式}（Sens 用）。回傳 {列名: 列號}。"""
    over = over or {}
    P = lambda k: over.get(k, k)
    R = {}
    ws.cell(r0, 1, title).font = BOLD
    for cc in range(1, 11): ws.cell(r0, cc).fill = SUB
    hdr(ws, r0 + 1, ["項目", "單位"] + [f"FY{y}" + ("A" if y == 2025 else "E" if y == 2026 else "") for y in YRS])
    names = ["mw", "pmw", "cl", "hw", "rev", "yoy", "gmc", "gmh", "dac", "gp", "gmcg", "gm", "opx", "coi", "sbcr", "sbc", "ebit", "om", "oth", "pti", "txr", "tax", "ni", "sh", "eps", "nmw", "cxc", "cxo", "cx", "dao", "da"]
    for i, n in enumerate(names): R[n] = r0 + 2 + i
    lab = {"mw": ("平均計費 MW", "MW"), "pmw": ("每 MW 年營收", "US$M/MW"), "cl": ("雲端核心營收", "US$B"), "hw": ("硬體核心營收", "US$B"),
           "rev": ("核心營收合計", "US$B"), "yoy": ("營收年增", "%"), "gmc": ("雲端現金毛利率（不含折舊）", "%"), "dac": ("雲端折舊攤銷（列入雲端成本）", "US$B"), "gmcg": ("雲端毛利率（含折舊）", "%"), "gmh": ("硬體核心毛利率", "%"),
           "gp": ("核心毛利", "US$B"), "gm": ("核心毛利率", "%"), "opx": ("核心營業費用（不含股票薪酬）", "US$B"), "coi": ("核心營業損益", "US$B"),
           "sbcr": ("股票薪酬 ÷ 營收", "%"), "sbc": ("經常性股票薪酬", "US$B"),
           "ebit": ("營業利益 EBIT（含股票薪酬；不含客戶認股權證攤銷、代墊收入與 IPO 一次性費用）", "US$B"), "om": ("營業利益率", "%"),
           "oth": ("其他收益淨額", "US$B"), "pti": ("稅前淨利", "US$B"), "txr": ("有效稅率", "%"), "tax": ("所得稅", "US$B"), "ni": ("淨利", "US$B"),
           "sh": ("稀釋股數", "十億股"), "eps": ("EPS（同 EBIT 口徑）", "US$"), "nmw": ("新增平均計費 MW", "MW"),
           "cxc": ("雲端資本支出（新增 MW × 每 MW 資本支出；FY2031 起含汰換）", "US$B"), "cxo": ("其他資本支出（營收 × oth_cx）", "US$B"), "cx": ("資本支出合計", "US$B"), "dao": ("其他折舊攤銷（列入營業費用）", "US$B"), "da": ("折舊攤銷合計", "US$B")}
    pct = {"yoy", "gmc", "gmh", "gm", "gmcg", "sbcr", "om", "txr"}
    for n in names:
        ws.cell(R[n], 1, lab[n][0]); ws.cell(R[n], 2, lab[n][1])
    def put(n, y, f):
        c = ws[f"{COL[y]}{R[n]}"]; c.value = f
        c.number_format = "0.0%" if n in pct else ("#,##0" if n in ("mw", "nmw") else "#,##0.00" if n in ("pmw", "eps") else "0.0000" if n == "sh" else "#,##0.000")
    C = COL
    # FY2025A（GAAP；只列可得項目）
    put("rev", 2025, "=Q!C7"); put("gp", 2025, "=Q!C8"); put("gm", 2025, f"=C{R['gp']}/C{R['rev']}"); put("ebit", 2025, "=Q!C9"); put("om", 2025, f"=C{R['ebit']}/C{R['rev']}")
    put("sbc", 2025, "=Q!C10"); put("cx", 2025, "=Q!C48")
    # FY2026E 錨點
    put("pmw", 2026, f"={P('p_mw')}"); put("cl", 2026, "=cl_26"); put("mw", 2026, f"=D{R['cl']}*1000/D{R['pmw']}")
    put("hw", 2026, "=hw_26"); put("rev", 2026, f"=D{R['cl']}+D{R['hw']}"); put("yoy", 2026, f"=D{R['rev']}/C{R['rev']}-1")
    put("gmc", 2026, "=gmc_26"); put("gmh", 2026, f"={P('hw_gm')}")
    put("gp", 2026, "=gp_26"); put("gm", 2026, f"=D{R['gp']}/D{R['rev']}"); put("opx", 2026, "=opx_26"); put("coi", 2026, f"=D{R['gp']}-D{R['opx']}")
    put("sbc", 2026, "=sbc_26"); put("sbcr", 2026, f"=D{R['sbc']}/D{R['rev']}"); put("ebit", 2026, f"=D{R['coi']}-D{R['sbc']}"); put("om", 2026, f"=D{R['ebit']}/D{R['rev']}")
    put("oth", 2026, "=oth_26"); put("pti", 2026, f"=D{R['ebit']}+D{R['oth']}"); put("txr", 2026, f"={P('tx_e')}")
    put("tax", 2026, f"=MAX(0,D{R['pti']})*D{R['txr']}"); put("ni", 2026, f"=D{R['pti']}-D{R['tax']}"); put("sh", 2026, "=sh_v"); put("eps", 2026, f"=D{R['ni']}/D{R['sh']}")
    put("cxc", 2026, "=cx_26"); put("cxo", 2026, 0); put("cx", 2026, f"=D{R['cxc']}+D{R['cxo']}")
    put("dac", 2026, "=da_26"); put("dao", 2026, 0); put("da", 2026, f"=D{R['dac']}+D{R['dao']}"); put("gmcg", 2026, f"=(D{R['cl']}*D{R['gmc']}-D{R['dac']})/D{R['cl']}")
    for i, y in enumerate(FC):
        c, p = C[y], C[y - 1]
        put("mw", y, f"=mw{s}_{y % 100}*{P('mw_k')}")
        put("pmw", y, f"={P('p_mw')}" if y == 2027 else f"={p}{R['pmw']}*(1+{P('pg' + s)})")
        put("cl", y, f"={c}{R['mw']}*{c}{R['pmw']}/1000")
        put("hw", y, f"={p}{R['hw']}*(1+{P('hg' + s)})")
        put("rev", y, f"={c}{R['cl']}+{c}{R['hw']}"); put("yoy", y, f"={c}{R['rev']}/{p}{R['rev']}-1")
        tgt = f"({P('gmc' + s)}+{P('gm_adj')})"
        put("gmc", y, f"=D{R['gmc']}+({tgt}-D{R['gmc']})*{min(i + 1, 3)}/3" if y < 2029 else f"={tgt}")
        put("gmh", y, f"={P('hw_gm')}")
        put("gp", y, f"={c}{R['cl']}*{c}{R['gmc']}+{c}{R['hw']}*{c}{R['gmh']}-{c}{R['dac']}"); put("gm", y, f"={c}{R['gp']}/{c}{R['rev']}")
        put("gmcg", y, f"=({c}{R['cl']}*{c}{R['gmc']}-{c}{R['dac']})/{c}{R['cl']}")
        put("opx", y, f"={p}{R['opx']}*(1+MIN({P('g_opx_max')},MAX({P('g_opx_min')},{P('phi')}*{c}{R['yoy']})))")
        put("coi", y, f"={c}{R['gp']}-{c}{R['opx']}")
        put("sbcr", y, f"={P('sbc_27')}+({P('sbc_ss')}-{P('sbc_27')})*{i}/5")
        put("sbc", y, f"={c}{R['rev']}*{c}{R['sbcr']}")
        put("ebit", y, f"={c}{R['coi']}-{c}{R['sbc']}"); put("om", y, f"={c}{R['ebit']}/{c}{R['rev']}")
        put("oth", y, "=oth_27" if y == 2027 else "=oth_28" if y == 2028 else "=oth_f")
        put("pti", y, f"={c}{R['ebit']}+{c}{R['oth']}")
        put("txr", y, f"={P('tx_e')}" if y <= 2028 else f"={P('tx_f')}")
        put("tax", y, f"=MAX(0,{c}{R['pti']})*{c}{R['txr']}"); put("ni", y, f"={c}{R['pti']}-{c}{R['tax']}")
        put("sh", y, f"={p}{R['sh']}*(1+{P('sh_g')})"); put("eps", y, f"={c}{R['ni']}/{c}{R['sh']}")
        put("nmw", y, f"=MAX(0,{c}{R['mw']}-{p}{R['mw']})")
        rep = f"+{C[y - 5]}{R['nmw']}*{P('capex_mw')}/1000" if y - 5 >= 2027 else (f"+D{R['cxc']}" if y - 5 == 2026 else "")
        put("cxc", y, f"={c}{R['nmw']}*{P('capex_mw')}/1000{rep}")
        put("cxo", y, f"={c}{R['rev']}*{P('oth_cx')}")
        put("cx", y, f"={c}{R['cxc']}+{c}{R['cxo']}")
        # 折舊：直線 life 年、取得年度半年；FY2026 起的資本支出（FY2026 全數視為雲端）；舊資產 da_old 至 FY2030（列入其他）
        def dep(rowk):
            cols = [C[k] for k in range(max(2026, y - 4), y)]
            terms = [f"{cc}{R[rowk]}" for cc in cols] + [f"0.5*{c}{R[rowk]}"]
            if y - 5 >= 2026: terms.append(f"0.5*{C[y - 5]}{R[rowk]}")
            return f"({'+'.join(terms)})/life"
        put("dac", y, "=" + dep("cxc")); put("dao", y, "=" + dep("cxo") + ("+da_old" if y <= 2030 else ""))
        put("da", y, f"={c}{R['dac']}+{c}{R['dao']}")
    return R


# ---------------- PL ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — 三情境收支（US$B；FY＝日曆年）。FY2025 為 GAAP 實績（只列可得項目）；FY2026 為錨點（核心口徑，1Q26–3Q26 實績＋4Q26 指引中點）；FY2027–FY2032 為模型"
PL["A2"] = "口徑：營收、毛利、營業費用為公司核心（non-GAAP）口徑；EBIT 再扣經常性股票薪酬（標準版 DCF 視為費用），不扣客戶認股權證攤銷（認股權證已計入評價股數）、代墊資料中心收入與 IPO 一次性費用"
SDEF = {"A": "情境 A（樂觀）：OpenAI 750 MW 2028 全數上線、選擇權部分執行（FY2032 平均計費 2,200 MW）；每 MW 年營收持平；雲端現金毛利率 72%；硬體年增 25%",
        "B": "情境 B（基準）：OpenAI 750 MW 分批上線（接近 RPO 認列時程）、其他客戶溫和成長（FY2032 平均計費 1,300 MW）；每 MW 年營收年降 4%；雲端現金毛利率 66%；硬體年增 5%",
        "C": "情境 C（保守）：OpenAI 上線延後、不續約擴大（FY2032 平均計費 600 MW）；每 MW 年營收年降 10%；雲端現金毛利率 58%；硬體年減 15%"}
PR = {}
rr = 4
for s in SC:
    PR[s] = engine(PL, rr, s, SDEF[s]); rr += ENGINE_ROWS + 3
PL.column_dimensions["A"].width = 70; PL.column_dimensions["B"].width = 11
for c in "CDEFGHIJ": PL.column_dimensions[c].width = 11

# ---------------- DCF（tools/fcff_dcf.py v2） ----------------
DCFW = wb.create_sheet("DCF")
DCFW["A1"] = "DCF — 標準版：無槓桿自由現金流 FCFF＝NOPAT＋折舊攤銷 − 資本支出 − 營運資金增加（股票薪酬視為費用），以 WACC 折現到評價日 2026-12-31（期中慣例），終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g）"
DCFW["A2"] = "資本支出、折舊攤銷逐年讀 PL 頁（雲端擴建：新增 MW × 每 MW 資本支出；FY2031 起含汰換）；稅率讀 PL 頁各年（FY2027–FY2028 為 0）。Stub 期（2026-09-30 → 12-31）FCFF 取 Q 頁 H33（4Q26 FCFF），已從淨現金扣除"


def paths_for(ws_name, R_, s, label, name=None):
    c = COL
    return dict(name=name or s, label=label,
                rev={y: f"{ws_name}!{c[y]}{R_['rev']}" for y in FC}, rev0=f"{ws_name}!D{R_['rev']}",
                ebit={y: f"{ws_name}!{c[y]}{R_['ebit']}" for y in FC}, tax={y: f"{ws_name}!{c[y]}{R_['txr']}" for y in FC},
                da={y: f"{ws_name}!{c[y]}{R_['da']}" for y in FC}, cx={y: f"{ws_name}!{c[y]}{R_['cx']}" for y in FC},
                stub_fcff="Q!$H$33")


RD_ = FD.write_detail(DCFW, 4, [paths_for("PL", PR[s], s, f"情境 {s}（{SCN[s]}）") for s in SC], FC, 1.0, stub=True)
DCF_END = DCFW.max_row
DCFW.column_dimensions["A"].width = 46; DCFW.column_dimensions["B"].width = 40

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — 目標價（DCF 與 Comparable 加權；美元）。Comparable 用同業前瞻 P/S（本公司 FY2026–FY2027 虧損、上市未滿一年，市場主流以營收倍數評價）；十年歷史倍數不適用"
V["A2"] = "目標時點：2027-12-31（評價日 2026-12-31 後 12 個月）；Comparable 的 12 個月後前瞻營收＝FY2028（日曆年 2028）營收"
V["A3"] = "0. 目標價（DCF 與 Comparable 加權）"; V["A3"].font = BOLD
SV = FD.write_summary(wb, V, 4, RD_, ["A", "B", "C"], ["tp_wA", "tp_wB", "tp_wC"], beta_note="自算產業 β 重新加槓桿")
r = 4 + FD.summary_rows(True) + 2
V.cell(r, 1, "0.2 Comparable（市場主流同業前瞻 P/S）").font = BOLD; R_C0 = r; r += 1
hdr(V, r, ["同業與前瞻 P/S", "代號", "前瞻 P/S（11-20）", "日期", "來源", "標記"]); R_PH = r; r += 1
PEERS = [("NVIDIA", "NVDA", 8.37, 186.64, "11-06", 125.53), ("AMD", "AMD", 8.76, 449.04, "11-06", 424.39), ("Broadcom", "AVGO", 8.91, 293.69, "11-06", 304.41),
         ("Marvell", "MRVL", 7.90, 141.20, "11-06", 122.28)]   # px1120：最後一欄由 11-13 收盤（162.85、534.06、347.01、156.76）改 11-20 收盤；px1113：最後一欄由 11-05 收盤（181.60、404.40、291.53、133.55）改 11-13 收盤（共識 P/S 與頁面價格不變）
R_P1 = r
for nm, tk_, fps, ppx, pdt, px5 in PEERS:
    V.cell(r, 1, nm); V.cell(r, 2, tk_)
    c = V.cell(r, 3, f"={fps}*{px5}/{ppx}"); c.number_format = "0.00"
    V.cell(r, 4, "2026-11-20")
    src5 = {"NVDA": "本專案 NVIDIA 模型 In 頁 px", "AMD": "本專案 AMD 模型 In 頁 px", "AVGO": "本專案 Broadcom 模型 In 頁 px", "MRVL": "本專案 Marvell 模型 In 頁 px",
            "CRWV": "本專案 CoreWeave README（v4.13 現價）；equibles 11-06 收盤 46.40 − 2.74", "NBIS": "equibles 11-06 收盤 129.49 − 9.12"}[tk_]
    V.cell(r, 5, f"stockanalysis.com 統計頁 Forward PS {fps}（2026-{pdt} 頁面，價格 {ppx}）× 11-20 收盤 {px5}（{src5}；工作單 px1120，closes_20261109_20261120.csv）÷ 頁面價格；v0.1.4 為 11-13 收盤、v0.1.3 為 11-05 收盤換算")
    V.cell(r, 6, "Interested-party（資料商彙整賣方營收預估）"); r += 1
R_P2 = r - 1
V.cell(r, 1, "  同業選擇"); V.cell(r, 5, "只用 AI 運算晶片同業（NVIDIA、AMD、Broadcom、Marvell；Andy 2026-11-08 21:21 定：Comparable 只用晶片同業）。推論雲端同業（CoreWeave 前瞻 P/S 1.2、Nebius 4.6）不納入；Cerebras 2Q26 營收七成為雲端服務，此選擇使倍數偏高。前瞻 P/S＝股價 ÷ 資料商次一會計年度營收預估"); R_PNOTE = r; r += 1
V.cell(r, 1, "同業前瞻 P/S 中位數"); V.cell(r, 2, f"=MEDIAN(C{R_P1}:C{R_P2})").number_format = "0.00"; R_MED = r; r += 1
V.cell(r, 1, "  各情境 12 個月後的前瞻營收（FY2028）")
for j, s in enumerate(SC):
    V.cell(r, 2 + j, f"=PL!F{PR[s]['rev']}").number_format = "#,##0.000"
V.cell(r, 5, "目標時點 2027-12-31 的未來 12 個月＝日曆年 2028（US$B）"); R_FR = r; r += 1
V.cell(r, 1, "12 個月後的前瞻營收（情境加權）"); V.cell(r, 2, f"=tp_wA*B{R_FR}+tp_wB*C{R_FR}+tp_wC*D{R_FR}").number_format = "#,##0.000"; R_FRW = r; r += 1
V.cell(r, 1, "評價用稀釋股數"); V.cell(r, 2, "=sh_v").number_format = "0.0000"; R_SH = r; r += 1
V.cell(r, 1, "Comparable 目標價"); V.cell(r, 2, f"=B{R_MED}*B{R_FRW}/B{R_SH}").number_format = "#,##0.00"
V.cell(r, 5, "同業前瞻 P/S 中位數 × 12 個月後的前瞻營收（情境加權）÷ 稀釋股數"); R_COMP = r; r += 2
V.cell(r, 1, "0.3 結論").font = BOLD; R_CON = r; r += 1
hdr(V, r, ["項目", "數值", "對照", "對照", "說明"]); R_CH = r; r += 1
T = {}
def vrow(key, lab, b, c=None, d=None, e=None, fmt="#,##0.00"):
    global r
    V.cell(r, 1, lab)
    for col, v in (("B", b), ("C", c), ("D", d)):
        if v is not None:
            V[f"{col}{r}"] = v; V[f"{col}{r}"].number_format = fmt if col == "B" or not isinstance(fmt, tuple) else fmt
    if e is not None: V.cell(r, 5, e)
    T[key] = r; r += 1
vrow("tp", "目標價（DCF 與 Comparable 加權）", f"=tp_w_dcf*B{SV['dcf']}+(1-tp_w_dcf)*tp_comp_adj", e='="DCF 權重 "&TEXT(tp_w_dcf,"0%")&"、Comparable 權重 "&TEXT(1-tp_w_dcf,"0%")&"（In 頁 tp_w_dcf，Andy 可改）"')
vrow("comp", "Comparable 目標價（市場主流同業法）", f"=B{R_COMP}")
vrow("cons", "市場共識目標價", "=cons_tp", e="stockanalysis.com：11 位分析師平均（分析師資料 2026-10-05 更新，股價當時約 165–200）；Interested-party")
vrow("up", "上檔空間：加權目標價", f"=B{T['tp']}/px-1", fmt="0.0%")
vrow("upc", "上檔空間：Comparable 目標價", f"=B{T['comp']}/px-1", fmt="0.0%")
vrow("geps", "與市場共識目標價的差距：EPS 因素", "不適用", e="本模型 Comparable 用營收倍數；共識 FY2027–FY2028 營收與 EPS 未取得（付費），無法拆分")
vrow("gmul", "與市場共識目標價的差距：倍數因素", f"=C{r}/D{r}-1", f"=cons_tp*sh_v/B{R_FRW}", f"=B{R_MED}", e="C＝市場共識目標價 × 股數 ÷ 本模型前瞻營收（情境加權）隱含的 P/S；D＝同業中位數；B＝C ÷ D − 1", fmt="0.00")
vrow("band", "價格帶（DCF 保守～樂觀）", f"=D{SV['v12']}", f"=B{SV['v12']}", e="B＝C 情境 DCF 價值、C＝A 情境 DCF 價值（皆已推到 12 個月後）")
vrow("ps", "  檢查用：現價隱含 P/S（市值 ÷ 各情境 FY2028 營收）", f"=px*sh_v/PL!F{PR['A']['rev']}", f"=px*sh_v/PL!F{PR['B']['rev']}", f"=px*sh_v/PL!F{PR['C']['rev']}", fmt="0.00")
vrow("meth", "目標價方法", "DCF（標準版）：無槓桿自由現金流 FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加（股票薪酬視為費用），以 WACC 折現到 2026 年底，終值 RONIC＝WACC、g＝MIN(3%, rf)；"
     "Comparable：同業前瞻 P/S 中位數 × FY2028 營收（情境加權）÷ 稀釋股數；兩者各半。十年 P/E 規則值不適用（上市未滿一年、FY2026–FY2027 虧損）")
vrow("wchk", "  檢查：情境權重合計（應為 1）", "=ROUND(tp_wA+tp_wB+tp_wC,6)", fmt="0.000")
vrow("tier", "上檔空間級距（六級，供分組）", f'=IF(B{T["up"]}>0.5,"大於 50%",IF(B{T["up"]}>0.25,"25%–50%",IF(B{T["up"]}>=0,"0%–25%",IF(B{T["up"]}>=-0.25,"0%～−25%",IF(B{T["up"]}>=-0.5,"−25%～−50%","小於 −50%")))))',
     e="以「上檔空間：加權目標價」分六級（Andy 2026-10-25 14:00）；投資建議另依五級固定規則（下一列）")
vrow("rec", "投資建議", f'=IF(B{T["up"]}>0.5,"強力買進",IF(B{T["up"]}>0.25,"買進",IF(B{T["up"]}>=-0.25,"中立",IF(B{T["up"]}>=-0.5,"放空","強力放空"))))',
     e="依上檔空間固定規則（Andy 2026-10-25 14:17）：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空")
r += 1
V.cell(r, 1, "1. 價格與市場隱含倍數").font = BOLD; r += 1
hdr(V, r, ["項目", "A", "B", "C", "說明"]); R_H1 = r; r += 1
vrow("px", "現價（2026-11-20 收盤）", "=px")
vrow("mcap", "市值（稀釋股數；US$B）", "=px*sh_v", fmt="#,##0.00")
vrow("ev", "企業價值（市值 − 淨現金與投資；US$B）", "=px*sh_v-dcf_nc", fmt="#,##0.00", e="淨現金與投資為 2026-09-30（未滾到評價日）")
vrow("evs27", "現價隱含 EV ÷ FY2027 營收", f"=B{T['ev']}/PL!E{PR['A']['rev']}", f"=B{T['ev']}/PL!E{PR['B']['rev']}", f"=B{T['ev']}/PL!E{PR['C']['rev']}", fmt="0.00")
vrow("evs28", "現價隱含 EV ÷ FY2028 營收", f"=B{T['ev']}/PL!F{PR['A']['rev']}", f"=B{T['ev']}/PL!F{PR['B']['rev']}", f"=B{T['ev']}/PL!F{PR['C']['rev']}", fmt="0.00")
r += 1
V.cell(r, 1, "0.5 相對倍數調整（Comparable 目標倍數＝同業中位數 × 本公司 10 年相對同業的倍數比）").font = BOLD; r += 1
V.cell(r, 1, "本公司 2026-05-14 上市，無十年前瞻倍數歷史（tools/data/fwd_mult_hist.csv 無 CBRS）；依 tools/relative_multiple.py 規則「本公司有效年數不足 → 係數＝1（不調整）」"); r += 1
vrow("rmk", "相對倍數調整係數", 1, e="資料不足時＝1（不調整）", fmt="0.00")
vrow("compadj", "Comparable 目標價（相對倍數調整後）", f"=Val!$B${R_COMP}*B{T['rmk']}", e="＝Comparable 目標價（市場主流同業法）× 調整係數；加權目標價使用本列")
dn("tp_comp_adj", f"Val!$B${T['compadj']}")
# 外部稽核優化 W3-1（2026-11-23）：0.3 結論的未調整 Comparable 列加註「未調整，不進加權」並並列調整後的值（共用 tools/relative_multiple.py；不改數值格）
import relative_multiple as _RM_W3  # noqa: E402
_RM_W3.mark_unadjusted(V, "#,##0.00", upto=T["rmk"])
r += 1
r = FD.write_checks(V, r, nc_bs="Q!F35+Q!F36+inv_k*Q!F37-Q!F42-Q!F43", nc_note="現金＋短期投資＋受限現金 − OpenAI 貸款 − 租賃（Q 頁 2026-09-30）")
for c, w in zip("ABCDEF", (60, 14, 14, 14, 90, 30)): V.column_dimensions[c].width = w

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 對帳：本模型與公司指引、揭露、共識、Tokenomics（只對帳，不校準）"
hdr(RC, 3, ["項目", "本模型（情境 A）", "本模型（情境 B）", "本模型（情境 C）", "對照值", "B 相對對照", "對照來源", "說明"])
RCR = {}
rr = 4
def rec(key, lab, a, b, c, ref, src, note="", fmt="#,##0.000", rel=True):
    global rr
    RC.cell(rr, 1, lab)
    for col, v in (("B", a), ("C", b), ("D", c), ("E", ref)):
        if v is not None:
            RC[f"{col}{rr}"] = v; RC[f"{col}{rr}"].number_format = fmt
    if rel and b is not None and ref is not None:
        RC[f"F{rr}"] = f"=C{rr}/E{rr}-1"; RC[f"F{rr}"].number_format = "0.0%"
    RC.cell(rr, 7, src); RC.cell(rr, 8, note)
    RCR[key] = rr; rr += 1
pa, pb, pc = PR["A"], PR["B"], PR["C"]
RC.cell(rr, 1, "A. FY2026 錨點與公司指引").font = BOLD; rr += 1
rec("rev26", "FY2026 核心營收 vs 共識", "=PL!D%d" % pa["rev"], "=PL!D%d" % pb["rev"], "=PL!D%d" % pc["rev"], "=cons_rev26", "In cons_rev26（stockanalysis）", "錨點＝公司指引中點，三情境相同")
rec("gm26", "FY2026 核心毛利率 vs 指引中點", "=PL!D%d" % pa["gm"], "=PL!D%d" % pb["gm"], "=PL!D%d" % pc["gm"], "=fy_gm", "In fy_gm", "", fmt="0.0%")
rec("om26", "FY2026 核心營業利益率 vs 指引中點", "=Q!H25", "=Q!H25", "=Q!H25", "=fy_om", "In fy_om", "Q 頁核心口徑（未扣經常性股票薪酬）", fmt="0.0%", rel=False)
RC.cell(rr, 1, "B. 2027 年營收、RPO 與容量").font = BOLD; rr += 1
rec("x27", "FY2027 營收 ÷ FY2026（公司：三倍）", "=PL!E%d/PL!D%d" % (pa["rev"], pa["rev"]), "=PL!E%d/PL!D%d" % (pb["rev"], pb["rev"]), "=PL!E%d/PL!D%d" % (pc["rev"], pc["rev"]), "=cfo_x27", "In cfo_x27", "公司說法無區間；情境 C 低於三倍", fmt="0.00")
rec("rpo24", "雲端營收 4Q26–3Q28（≈ 4Q26＋FY2027＋FY2028 × 0.75）vs RPO 24 個月內認列", f"=Q!G19+PL!E{pa['cl']}+0.75*PL!F{pa['cl']}", f"=Q!G19+PL!E{pb['cl']}+0.75*PL!F{pb['cl']}", f"=Q!G19+PL!E{pc['cl']}+0.75*PL!F{pc['cl']}", "=rpo*rpo_24", "In rpo、rpo_24（10-Q 附註；2026-09-30 起算）",
    "RPO 含硬體與支援服務；模型雲端營收亦含 RPO 以外的新合約，故高於 RPO 認列額屬合理，低於則表示上線時程比合約慢")
rec("rpo48", "雲端營收 4Q28–3Q30（≈ FY2028 × 0.25＋FY2029＋FY2030 × 0.75）vs RPO 第 25–48 個月", f"=0.25*PL!F{pa['cl']}+PL!G{pa['cl']}+0.75*PL!H{pa['cl']}", f"=0.25*PL!F{pb['cl']}+PL!G{pb['cl']}+0.75*PL!H{pb['cl']}", f"=0.25*PL!F{pc['cl']}+PL!G{pc['cl']}+0.75*PL!H{pc['cl']}", "=rpo*rpo_48", "In rpo、rpo_48", "同上")
rec("mw28", "FY2028 平均計費 MW vs 公司：700 gross MW 已簽約、2027 年底前交付", f"=PL!F{pa['mw']}", f"=PL!F{pb['mw']}", f"=PL!F{pc['mw']}", "=mw_co27", "In mw_co27", "公司口徑為容量（含尚未計費），可能為設施電力，與計費 MW 不同", fmt="#,##0")
rec("oai", "OpenAI 承諾 750 MW 是否在 FY2028 前上線（FY2028 平均計費 MW ÷ 750）", f"=PL!F{pa['mw']}/oai_mw", f"=PL!F{pb['mw']}/oai_mw", f"=PL!F{pc['mw']}/oai_mw", None, "In oai_mw", "含 OpenAI 以外客戶，比值為上限", fmt="0.00", rel=False)
RC.cell(rr, 1, "C. 與 Tokenomics 對照（跨層一致性；只對帳）").font = BOLD; rr += 1
rec("tkhold", "每 MW 年營收（FY2027）vs Tokenomics 隱含 Cerebras 每 GW 年持有成本（VR200 × 持有成本比）", f"=PL!E{pa['pmw']}", f"=PL!E{pb['pmw']}", f"=PL!E{pc['pmw']}", "=tk_hold_vr*tk_hold_cs", TKV,
    "合約隱含單價低於 Tokenomics 的持有成本：若 Tokenomics 比值正確，Cerebras 以低於買方自建成本的價格出租（毛利 40%–50% 與此矛盾）；較可能是 Tokenomics 持有成本比 1.0（Assumed）偏高，或 750 MW 為設施電力口徑。待送證據給算力收支", fmt="#,##0.00")
rec("tkcx", "每 MW 資本支出（基準）vs Tokenomics 隱含（VR200 IT 資本支出 × 持有成本比 ×（1 − 硬體毛利率））", "=capex_mw", "=capex_mw", "=capex_mw", "=capex_mw_hi", TKV,
    "基準取公司不動產廠房設備推估；Tokenomics 隱含值作為 Sens 高值", fmt="#,##0.00")
RC.cell(rr, 1, "D. 本模型 FY2027–FY2028 關鍵數字").font = BOLD; rr += 1
rec("rev27", "FY2027 核心營收", f"=PL!E{pa['rev']}", f"=PL!E{pb['rev']}", f"=PL!E{pc['rev']}", "=PL!D%d*cfo_x27" % pb["rev"], "FY2026 錨點 × 3（公司說法）", "", rel=True)
rec("ebit28", "FY2028 EBIT", f"=PL!F{pa['ebit']}", f"=PL!F{pb['ebit']}", f"=PL!F{pc['ebit']}", None, "—", "", rel=False)
rec("cx27", "FY2027–FY2028 資本支出合計 vs 2026-09-30 現金與投資（含受限現金）", f"=PL!E{pa['cx']}+PL!F{pa['cx']}", f"=PL!E{pb['cx']}+PL!F{pb['cx']}", f"=PL!E{pc['cx']}+PL!F{pc['cx']}", "=Q!F35+Q!F36+Q!F37", "Q 頁",
    "加上 4Q26 資本支出與營業虧損後，情境 A 可能需外部融資（循環額度 8.5 億、OpenAI 貸款）", rel=True)
for c, w in zip("ABCDEFGH", (70, 14, 14, 14, 14, 12, 50, 90)): RC.column_dimensions[c].width = w

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 敏感度（一次動一個輸入；以情境 B 為底）：FY2028 營收、FY2028 EPS、情境 B DCF 價值與情境 B 加權價值（DCF 與 Comparable 各半；Comparable 同業倍數不變）"
hdr(SE, 3, ["案例", "設定", "FY2028 營收", "變動", "FY2028 EPS", "變動", "情境 B DCF 價值", "變動", "情境 B 加權價值", "變動"])
CASES = [("mw_k", "平均計費 MW 係數 ×0.8／×1.2", "mw_k_lo", "mw_k_hi"),
         ("p_mw", "每 MW 年營收（合約期 5 年／3 年）", "p_mw_lo", "p_mw_hi"),
         ("gm_adj", "雲端現金毛利率 −5pt／＋5pt", "gm_adj_lo", "gm_adj_hi"),
         ("capex_mw", "每 MW 資本支出 7／Tokenomics 隱含", "capex_mw_lo", "capex_mw_hi"),
         ("hw_gm", "硬體毛利率 35%／45%", "hw_gm_lo", "hw_gm_hi"),
         ("phi", "營業費用彈性 0.4／0.8", "phi_lo", "phi_hi"),
         ("g_opx_max", "營業費用年增上限 40%／80%", "g_opx_max_lo", "g_opx_max_hi"),
         ("sbc_ss", "長期股票薪酬率 4%／8%", "sbc_ss_lo", "sbc_ss_hi"),
         ("sbc_27", "FY2027 股票薪酬率 10%／18%", "sbc_27_lo", "sbc_27_hi"),
         ("tx_f", "FY2029 起稅率 10%／21%", "tx_f_lo", "tx_f_hi"),
         ("oth_cx", "其他資本支出 2%／5%", "oth_cx_lo", "oth_cx_hi")]
# 基準列
SR = {}
SR["base"] = 4
SE.cell(4, 1, "基準（情境 B）"); SE.cell(4, 2, "—")
SE["C4"] = f"=PL!F{pb['rev']}"; SE["E4"] = f"=PL!F{pb['eps']}"; SE["G4"] = f"=Val!C{SV['v12']}"
SE["I4"] = f"=tp_w_dcf*G4+(1-tp_w_dcf)*Val!$B${R_MED}*C4/sh_v*Val!$B${T['rmk']}"
rr = 5
case_rows = []
for k, lab, lo, hi in CASES:
    for side, nm in (("低", lo), ("高", hi)):
        SE.cell(rr, 1, f"{lab}：{side}"); SE.cell(rr, 2, f"{k}＝{nm}")
        case_rows.append((rr, k, nm, side, lab)); rr += 1
CASE_END = rr - 1
for c, w in zip("ABCDEFGHIJ", (44, 26, 12, 9, 12, 9, 14, 9, 14, 9)): SE.column_dimensions[c].width = w
rr += 1
c1, c2, rr = FD.write_sens(SE, rr, RD_, ["A", "B", "C"], ["tp_wA", "tp_wB", "tp_wC"], comp_ref="tp_comp_adj", base_ref=f"Val!$B${T['tp']}")
STD_SENS = (c1, c2)
rr += 2
SE.cell(rr, 1, "案例明細（每案例一個收支區塊；以情境 B 為底，只改一個輸入）").font = BOLD; rr += 1
dcf_r = DCF_END + 3
for (row_, k, nm, side, lab) in case_rows:
    Rk = engine(SE, rr, "B", f"案例：{lab}（{side}；{k}＝{nm}）", over={k: nm})
    rr += ENGINE_ROWS + 3
    Rd = FD.write_detail(DCFW, dcf_r, [paths_for("Sens", Rk, "B", f"Sens 案例：{lab}（{side}）", name=f"S{row_}")], FC, 1.0, stub=True,
                         title=f"FCFF 明細（Sens 案例：{lab}，{side}）")
    dcf_r = DCFW.max_row + 3
    vexpr = FD.value_expr(Rd, f"S{row_}")
    SE[f"C{row_}"] = f"=F{Rk['rev']}"; SE[f"E{row_}"] = f"=F{Rk['eps']}"; SE[f"G{row_}"] = f"={vexpr}"
    SE[f"I{row_}"] = f"=tp_w_dcf*G{row_}+(1-tp_w_dcf)*Val!$B${R_MED}*C{row_}/sh_v*Val!$B${T['rmk']}"
for row_ in [4] + [x[0] for x in case_rows]:
    for c, f in (("C", "#,##0.000"), ("E", "#,##0.00"), ("G", "#,##0.00"), ("I", "#,##0.00")):
        SE[f"{c}{row_}"].number_format = f
    if row_ != 4:
        for c, b in (("D", "C"), ("F", "E"), ("H", "G"), ("J", "I")):
            SE[f"{c}{row_}"] = f"=IF(ABS({b}$4)<1E-9,\"\",{b}{row_}/{b}$4-1)"
            SE[f"{c}{row_}"].number_format = "0.0%"
# 基準核對：Sens 引擎（無覆寫）＝PL 情境 B
SE.cell(CASE_END + 1, 1, "  檢查：Sens 基準列取自 PL 情境 B 與 Val（無覆寫）")

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources")
SO["A1"] = "來源（公司原文＝Cerebras 8-K 財報新聞稿與 10-Q，sec.gov；S-1／424B4 與合約條款只經二手摘要）"; SO["A1"].font = BOLD
hdr(SO, 2, ["來源", "用途"])
for i, (s_, u_) in enumerate((
        (S8K3, "3Q26 GAAP 與核心損益、分部調節表、2026-09-30 資產負債、9M26 現金流、4Q26 與 FY2026 指引、RPO 261 億、700 gross MW、2027 年營收三倍說法（v0.1.3）"),
        (SQ3, "股數（2026-11-05 封面）、潛在稀釋證券、RPO 認列時點、客戶集中、OpenAI 營收、尚未起租租約、OpenAI 貸款（v0.1.3）"),
        (S8K2, "2Q26 GAAP 與核心損益、調節表、2026-06-30 資產負債、1H26 現金流、3Q26 與 FY2026 指引（8 月）、600 MW、AWS、OpenAI 更新"),
        (S8K1, "1Q26 GAAP 與核心損益、分部毛利率"),
        (SQ2, "2Q26 營收分部說明、客戶認股權證攤銷、OpenAI 營收"),
        (SQ1, "1Q26 營收分部"),
        (SAPI, "FY2024–FY2025（424B4）、客戶集中、OpenAI 貸款與循環額度、股數、RPO 認列時點、租約、潛在稀釋"),
        (SFUT, "S-1：OpenAI 主合作協議 750 MW、逾 200 億美元、AWS 條款書、2025 客戶集中"),
        ("Yahoo Finance 日線收盤 2026-11-20（工作單 px1120；closes_20261109_20261120.csv）；v0.1.4 為 11-13 收盤 108.72；v0.1.3 的 2026-11-05 收盤 81.71 見 " + SPX, "現價（2026-11-20 收盤）"),
        ("stockanalysis.com 預測頁 https://stockanalysis.com/stocks/cbrs/forecast/ ；統計頁 https://stockanalysis.com/stocks/cbrs/statistics/", "共識營收、EPS、目標價；股數與現金（對照）"),
        ("stockanalysis.com 各同業統計頁（NVDA、AMD、AVGO、MRVL；2026-11-08 讀取）", "同業前瞻 P/S"),
        (TKV, "Cerebras 產出比、持有成本比、VR200 持有成本與 IT 資本支出（Recon 與 capex_mw 高值）"),
        ("docs/plan/自算產業β_20261105.md", "自算產業 β"),
        ("docs/plan/個股DCF標準規格.md；tools/fcff_dcf.py v2", "標準版 DCF"),), 3):
    SO.cell(i, 1, s_); SO.cell(i, 2, u_)
SO.column_dimensions["A"].width = 120; SO.column_dimensions["B"].width = 80

# ---------------- README ----------------
RM_ = wb.create_sheet("README", 0)
pB = PR["B"]
RL = [
    f"Cerebras Systems（CBRS）收支與估值模型 {VER}（AI 半導體 Project：運算層；格式比照 NVIDIA v0.3.4）　{VDATE}",
    "建置：python3 cbrs/build_cbrs.py raw.xlsx，再 LibreOffice headless 重算；HTML：python3 cbrs/build_cbrs.py raw.xlsx --html 寫 html.json → tools/company_html/build_company_html.py cbrs",
    "模型命題：Cerebras 的營收＝推論雲端（平均計費 MW × 每 MW 年營收）＋系統銷售。OpenAI 750 MW 合約使 2027 年營收可達三倍；股價取決於 OpenAI 以外能否擴大、每 MW 價格與毛利能否撐住，以及擴建資本支出能否由 IPO 現金支應。",
    "核心問題：OpenAI 合約之後，Cerebras 能否把推論雲端擴大到 OpenAI 以外，並以自有現金完成建置。",
    "",
    "會計年度＝日曆年（FY2026＝CY2026）。金額：十億美元；每 MW 年營收與每 MW 資本支出：百萬美元／MW（＝十億美元／GW）",
    "",
    "驅動因子 → 推導量",
    "平均計費 MW（各情境 FY2027–FY2032，In 頁 A 節）× 每 MW 年營收（OpenAI 合約：200 億 ÷ 0.75 GW ÷ 3.5 年；FY2028 起逐年調整）→ 雲端營收；硬體＝FY2026 × 情境年增",
    "毛利率：雲端現金毛利率（不含折舊）由 FY2026 錨點（指引推得）線性收斂到情境長期值（FY2029），再扣雲端折舊得含折舊毛利；硬體 40%。營業費用依營收成長彈性 0.6（年增下限 5%、上限 60%）",
    "EBIT＝核心營業損益 − 經常性股票薪酬（FY2027 占營收 14%，收斂到 FY2032 的 6%）；不扣客戶認股權證攤銷（認股權證已計入評價股數 3.15 億股）",
    "資本支出＝新增平均計費 MW × 每 MW 資本支出 9（百萬美元，以 FY2026 折舊與不動產廠房設備推估；Tokenomics 隱含約 22.6 為高值）＋營收 × 3%；FY2031 起含 5 年汰換；折舊直線 5 年，雲端折舊列入雲端成本",
    "估值：DCF（標準版 FCFF、WACC）與 Comparable（同業前瞻 P/S 中位數 × FY2028 營收）各半；情境權重 0.3／0.5／0.2",
    "",
    "標記：Verified、Interested-party、Analogy、Assumed、Derived。S-1 與 OpenAI 合約條款只經二手摘要，待人工複核",
    "已知限制：(1) OpenAI 合約單價與 MW 口徑未公開（Interested-party 推得）；(2) 每 MW 資本支出無公司揭露；(3) FY2026 錨點含 4Q26 指引（未實現）；(4) 共識 FY2027–FY2028 未取得；(5) 上市未滿一年，β 用產業值、無十年倍數",
    "",
    "",
    "",
    "",
    "",
]
for i, t in enumerate(RL, 1):
    RM_.cell(i, 1, t)
RM_["A1"].font = BOLD
# 投資建議與重點觀察（公式帶數字）
up, tpv, dcf_, cmp_ = f"Val!B{T['up']}", f"Val!B{T['tp']}", f"Val!B{SV['dcf']}", f"tp_comp_adj"
RM_["A21"] = f'="投資建議："&Val!B{T["rec"]}&"（依上檔空間固定規則）"'
RM_["A22"] = (f'="重點觀察 1：目標價 US$"&TEXT({tpv},"0.00")&"（DCF "&TEXT({dcf_},"0.00")&"、Comparable "&TEXT({cmp_},"0.00")&"；DCF 權重 "&TEXT(tp_w_dcf,"0%")&"），比現價 "&TEXT(px,"0.00")&" "&IF({up}>=0,"高 ","低 ")&TEXT(ABS({up}),"0.0%")&"；Comparable 單獨為 "&TEXT(Val!B{T["upc"]},"+0.0%;-0.0%")&"。'
               f'"&IF(ABS({up})>0.25,"上檔空間大的主因："&IF({up}>0,"同業前瞻 P/S 中位數 "&TEXT(Val!B{R_MED},"0.0")&" 倍套在 FY2028 營收（情境加權 "&TEXT(Val!B{R_FRW},"0.00")&" 十億美元），而現價只隱含 "&TEXT(Val!C{T["ps"]},"0.0")&" 倍（情境 B）","DCF 偏低：擴建資本支出使 FY2027–FY2029 自由現金流為負"),"兩種方法分歧：DCF 只有 "&TEXT({dcf_},"0.00")&"（擴建期自由現金流為負），Comparable 為 "&TEXT({cmp_},"0.00")&"。")')
RM_["A23"] = (f'="重點觀察 2：FY2027 營收（情境 B）"&TEXT(PL!E{pB["rev"]},"0.00")&" 十億美元，為 FY2026 的 "&TEXT(PL!E{pB["rev"]}/PL!D{pB["rev"]},"0.0")&" 倍（公司：三倍）；'
               f'FY2028 營收 A／B／C 為 "&TEXT(PL!F{PR["A"]["rev"]},"0.00")&"／"&TEXT(PL!F{pB["rev"]},"0.00")&"／"&TEXT(PL!F{PR["C"]["rev"]},"0.00")&"，差異主要來自平均計費 MW（"&TEXT(PL!F{PR["A"]["mw"]},"#,##0")&"／"&TEXT(PL!F{pB["mw"]},"#,##0")&"／"&TEXT(PL!F{PR["C"]["mw"]},"#,##0")&"）。"')
RM_["A24"] = (f'="重點觀察 3：最敏感的是每 MW 年營收與平均計費 MW：合約期以 5 年計（每 MW 年營收 "&TEXT(p_mw_lo,"0.0")&"）時情境 B 加權價值 "&TEXT(Sens!J{[x[0] for x in case_rows if x[1]=="p_mw" and x[3]=="低"][0]},"+0.0%;-0.0%")&"，MW ×0.8 時 "&TEXT(Sens!J{[x[0] for x in case_rows if x[1]=="mw_k" and x[3]=="低"][0]},"+0.0%;-0.0%")&"，每 MW 資本支出改用 Tokenomics 隱含值時 "&TEXT(Sens!J{[x[0] for x in case_rows if x[1]=="capex_mw" and x[3]=="高"][0]},"+0.0%;-0.0%")&"。'
               f'下一個檢驗點：4Q26 財報的核心營收對照指引 3.10–3.30 億美元、毛利率能否到 47%–49%，以及 2027 年三倍成長的具體容量上線時程。"')
# 展示層
RM_["A31"] = "展示層：命題摘要、命題分項與白話重點（HTML 用）"; RM_["A31"].font = BOLD
RM_["A32"] = "命題：Cerebras 的營收＝推論雲端（平均計費 MW × 每 MW 年營收）＋系統銷售；OpenAI 750 MW 合約撐起 2027 年的高成長，但價值取決於 OpenAI 以外的擴張、單價與毛利，以及擴建所需資本支出。"
RM_["A33"] = "命題 1：2027 年營收能否達到公司說的「三倍」（FY2026 錨點約 9.52 億美元的三倍約 28.6 億）。"
RM_["A34"] = "命題 2：OpenAI 合約隱含的每 MW 年營收，是否足以覆蓋自建成本——毛利率能否回到 45% 以上，以及與 Tokenomics 持有成本的差距。"
RM_["A35"] = "命題 3：擴建數百 MW 的資本支出，2026-09-30 約 78 億美元的現金與投資（含受限現金）是否足夠，不需再融資或稀釋。"
RM_["A36"] = "命題 4：現價（11-20 收盤 82.67 美元）相對同業前瞻 P/S 是低估，還是反映客戶集中（UAE、OpenAI）與解禁賣壓的合理折價。"
RM_["A37"] = (f'="重點 1：加權目標價約 "&TEXT({tpv},"0.00")&" 美元，比現價 "&TEXT(px,"0.00")&" "&IF({up}>=0,"高 ","低 ")&TEXT(ABS({up}),"0.0%")&"，投資建議「"&Val!B{T["rec"]}&"」。'
               f'兩種方法差距大：Comparable（同業 P/S）"&TEXT({cmp_},"0.00")&"，DCF 只有 "&TEXT({dcf_},"0.00")&"，因為擴建期自由現金流為負、長期營業利益率有限。"')
RM_["A38"] = (f'="重點 2：情境 B 的 FY2027 營收 "&TEXT(PL!E{pB["rev"]},"0.00")&" 十億美元（FY2026 的 "&TEXT(PL!E{pB["rev"]}/PL!D{pB["rev"]},"0.0")&" 倍）、FY2028 "&TEXT(PL!F{pB["rev"]},"0.00")&"；'
               f'營業利益率 FY2028 "&TEXT(PL!F{pB["om"]},"0.0%")&"、FY2032 "&TEXT(PL!J{pB["om"]},"0.0%")&"（含股票薪酬）。"')
RM_["A39"] = (f'="重點 3：資本支出是 DCF 的關鍵。情境 B FY2027–FY2028 資本支出合計 "&TEXT(PL!E{pB["cx"]}+PL!F{pB["cx"]},"0.0")&" 十億美元，對照 2026-09-30 現金與投資 "&TEXT(Q!F35+Q!F36+Q!F37,"0.0")&"；'
               f'每 MW 資本支出若接近 Tokenomics 隱含值，情境 B 加權價值 "&TEXT(Sens!J{[x[0] for x in case_rows if x[1]=="capex_mw" and x[3]=="高"][0]},"+0.0%;-0.0%")&"。"')
RM_["A40"] = (f'="重點 4：OpenAI 合約隱含的每 MW 年營收約 "&TEXT(p_mw,"0.0")&" 百萬美元，低於 Tokenomics 隱含的 Cerebras 每 GW 年持有成本 "&TEXT(tk_hold_vr*tk_hold_cs,"0.0")&"（十億美元／GW）。'
               f'若 Tokenomics 比值正確，合約毛利不足；本模型以公司毛利率實績為準，差距列在 Recon 頁並待送證據給算力收支。"')
RM_["A42"] = "命題回覆（HTML 用；依本模型計算回答開頭四個命題）"; RM_["A42"].font = BOLD
RM_["A43"] = (f'="2027 年營收能否三倍？答：情境 A／B／C 的 FY2027 營收為 FY2026 的 "&TEXT(PL!E{PR["A"]["rev"]}/PL!D{PR["A"]["rev"]},"0.0")&"／"&TEXT(PL!E{pB["rev"]}/PL!D{pB["rev"]},"0.0")&"／"&TEXT(PL!E{PR["C"]["rev"]}/PL!D{PR["C"]["rev"]},"0.0")&" 倍；'
               f'情境 B 要 2027 年平均計費約 "&TEXT(PL!E{pB["mw"]},"#,##0")&" MW，對照公司「700 gross MW 已簽約、2027 年底前交付」。"')
RM_["A44"] = (f'="單價夠不夠覆蓋成本？答：以公司毛利率實績推算可以——雲端現金毛利率 FY2026 約 "&TEXT(gmc_26,"0.0%")&"、情境 B 長期 "&TEXT(gmcB,"0%")&"，扣折舊後 FY2030 雲端毛利率 "&TEXT(PL!H{pB["gmcg"]},"0.0%")&"；但與 Tokenomics 對照時，合約單價 "&TEXT(p_mw,"0.0")&" 只有 Tokenomics 隱含持有成本 "&TEXT(tk_hold_vr*tk_hold_cs,"0.0")&" 的 "&TEXT(p_mw/(tk_hold_vr*tk_hold_cs),"0%")&"，兩者至少一方需修正。"')
RM_["A45"] = (f'="現金夠不夠？答：情境 B 下 FY2027–FY2029 自由現金流合計 "&TEXT(DCF!C{RD_["paths"]["B"]["fcff"]}+DCF!D{RD_["paths"]["B"]["fcff"]}+DCF!E{RD_["paths"]["B"]["fcff"]},"0.0")&" 十億美元，對照評價日淨現金 "&TEXT(Val!C{SV["ncr"]},"0.0")&"；'
               f'情境 A 擴建更快，"&IF(Val!B{SV["ncr"]}+DCF!C{RD_["paths"]["A"]["fcff"]}+DCF!D{RD_["paths"]["A"]["fcff"]}+DCF!E{RD_["paths"]["A"]["fcff"]}<0,"需要外部融資","仍可自行支應")&"。"')
RM_["A46"] = (f'="現價是低估還是合理折價？答：現價隱含情境 B FY2028 P/S "&TEXT(Val!C{T["ps"]},"0.0")&" 倍，低於同業中位數 "&TEXT(Val!B{R_MED},"0.0")&" 倍，Comparable 認為低估；'
               f'但標準版 DCF 只有 "&TEXT({dcf_},"0.00")&"，"&IF({dcf_}<px,"低於現價","高於現價")&"。加權結果為「"&Val!B{T["rec"]}&"」。"')
RM_["A48"] = "本版要點與版本沿革（HTML 附錄用）"; RM_["A48"].font = BOLD
RM_["A49"] = "v0.1.1（2026-11-08）：Comparable 同業改為只用晶片同業（NVIDIA、AMD、Broadcom、Marvell；移除 CoreWeave、Nebius；Andy 2026-11-08 21:21）。其他判斷項照預設。｜" + f"v0.1（{VDATE}）：新建。格式比照 NVIDIA v0.3.4（README、In、Q、PL、Val、DCF、Recon、Sens、Sources；HTML 版面規格 v1.0）。Tokenomics {TKS['version']} 快照（Cerebras 產出比、持有成本比、VR200 持有成本與 IT 資本支出，只作 Recon 與 Sens 高值）。現價為 2026-11-05 收盤（與 AI半導體 其他個股同日）。"
RM_.column_dimensions["A"].width = 160

# ---------------- In 版面、儲存 ----------------
INL.apply(wb, "v0.1.1", VDATE)   # v0.1.2 起固定：In 頁版面說明與 README 版本沿革句屬 v0.1.1 的變更（不隨版號改寫）
for ws in (PL, V, DCFW, SE):
    ws.freeze_panes = "B3" if ws is not V else None
# v0.1.2（2026-11-15）：個股 HTML 版面規格 v1.3＋三層協調（本資料夾 v012_post.py；只新增頁與列，既有數值格不變）
sys.path.insert(0, HERE)
import v012_post  # noqa: E402
print(v012_post.apply(wb, PR, TKV))
import px1113_post  # noqa: E402   v0.1.4：現價與同業報價改 2026-11-13 收盤（工作單 px1113）
print(px1113_post.apply(wb))
import px1120_post  # noqa: E402   v0.1.5：現價與同業報價改 2026-11-20 收盤（工作單 px1120）
print(px1120_post.apply(wb))
out = sys.argv[1]
wb.save(out)

# ---------------- html.json ----------------
if "--html" in sys.argv:
    def inrow(k):
        return IKEY[k]
    def irow(k, f="n2"):
        n = inrow(k)
        return [f"@In!B{n}", f"@In!D{n}|{f}", f"@In!E{n}|{f}", f"@In!F{n}|{f}", f"@In!C{n}", f"@In!G{n}|tag"]
    IH = ["@In!B3", "@In!D3", "@In!E3", "@In!F3", "@In!C3", "@In!G3"]
    # 情境驅動對照節（in_layout 產生）
    scn0 = None
    for x in range(1, IN.max_row + 1):
        v = IN.cell(x, 1).value
        if isinstance(v, str) and "情境驅動對照" in v:
            scn0 = x
    scn_rows = [x for x in range(scn0 + 2, IN.max_row + 1) if IN.cell(x, 2).value]
    def sfmt(x):
        k = IN.cell(x, 1).value or ""
        lab = IN.cell(x, 2).value or ""
        return "n0" if "MW" in lab and "每 MW" not in lab else "p0" if ("%" in (IN.cell(x, 3).value or "") or k == "scn_w") else "n2"
    def blk(keys, fmts):
        sens = [k for k in keys if IN.cell(inrow(k), 4).value is not None or IN.cell(inrow(k), 6).value is not None]
        fixed = [k for k in keys if k not in sens]
        return sens, fixed
    def tables(title, keys, fmts):
        sens, fixed = blk(keys, fmts)
        out = [{"type": "h3", "text": title}]
        if sens:
            out.append({"type": "table", "caption": f"{title}：有敏感度區間的參數（基準進目標價；低／高只用於敏感度頁）", "head": IH, "rows": [irow(k, fmts.get(k, "n2")) for k in sens]})
        if fixed:
            out.append({"type": "table", "collapse": True, "summary": f"展開：{title}的其他固定輸入（實績、公司揭露、換算值）", "head": IH, "rows": [irow(k, fmts.get(k, "n2")) for k in fixed]})
        return out
    PCT = lambda ks: {k: "p1" for k in ks}
    cloud_keys = ["p_mw", "mw_k", "gm_adj", "capex_mw", "life", "oth_cx", "da_old", "cl_26", "gmc_26"]
    hw_keys = ["hw_gm", "hw_26", "q4_hw"]
    co_keys = ["q4_rev", "q4_gm", "q4_om", "fy_rev", "fy_gm", "fy_om", "q4_da", "q4_cx", "q4_oth", "gp_26", "opx_26", "sbc_26", "da_26", "cx_26", "oth_26",
               "phi", "g_opx_min", "sbc_27", "sbc_ss", "oth_27", "oth_28", "oth_f", "tx_e", "tx_f", "sh_g", "px", "tp_w_dcf"]
    ref_keys = ["oai_val", "oai_mw", "oai_opt", "oai_yr", "rpo", "rpo_24", "rpo_48", "mw_co27", "cfo_x27", "cons_rev26", "cons_eps26", "cons_tp", "tk_out_cs", "tk_hold_cs", "tk_hold_vr", "tk_capit"]
    fm = PCT(["oth_cx", "gm_adj", "gmc_26", "hw_gm", "q4_gm", "q4_om", "fy_gm", "fy_om", "g_opx_min", "sbc_27", "sbc_ss", "tx_e", "tx_f", "sh_g", "tp_w_dcf", "rpo_24", "rpo_48"])
    fm.update({"oai_mw": "n0", "oai_opt": "n0", "mw_co27": "n0", "q4_rev": "n3", "fy_rev": "n3", "q4_hw": "n3", "q4_da": "n3", "q4_cx": "n3", "q4_oth": "n3", "hw_26": "n3", "cl_26": "n3", "cons_rev26": "n3", "da_old": "n3",
               "gp_26": "n3", "opx_26": "n3", "sbc_26": "n3", "da_26": "n3", "oth_26": "n3", "oth_27": "n3", "oth_28": "n3", "oth_f": "n3", "life": "n0", "oai_val": "n0", "rpo": "n1", "cons_tp": "n2", "px": "n2"})
    drivers = [{"type": "table", "caption": "情境驅動（三情境並排，進加權目標價；取自 In 頁最後一節「情境驅動對照」）",
                "head": [f"@In!B{scn0 + 1}", f"@In!D{scn0 + 1}", f"@In!E{scn0 + 1}", f"@In!F{scn0 + 1}", f"@In!C{scn0 + 1}"],
                "rows": [[f"@In!B{x}", f"@In!D{x}|{sfmt(x)}", f"@In!E{x}|{sfmt(x)}", f"@In!F{x}|{sfmt(x)}", f"@In!C{x}"] for x in scn_rows]}]
    drivers += tables("一、推論雲端（平均計費 MW × 每 MW 年營收；擴建資本支出）", cloud_keys, fm)
    drivers += tables("二、系統（硬體）銷售", hw_keys, fm)
    drivers += tables("三、全公司：錨點、費用、股票薪酬、稅與股數、股價與目標價規則", co_keys, fm)
    drivers += tables("四、合約、揭露、共識與 Tokenomics（只作對帳）", ref_keys, fm)
    def plm(name_cell, rowkey, fmt, chart=False, ymin=None):
        m = {"name": name_cell, "fmt": fmt, "rows": {s: f"@PL!D{PR[s][rowkey]}:J{PR[s][rowkey]}" for s in SC}}
        if chart: m["chart"] = True
        if ymin: m["ymin"] = ymin
        return m
    pl_sec = {"id": "pl", "title": "三情境損益", "intro": "@PL!A1", "blocks": [
        {"type": "note", "parts": ["@PL!A2"]},
        {"type": "pl", "years": f"@PL!D{PR['A']['mw'] - 1}:J{PR['A']['mw'] - 1}", "metrics": [
            plm(["營收（", f"@PL!B{PR['A']['rev']}", "）"], "rev", "n2", True),
            plm(["平均計費 MW（", f"@PL!B{PR['A']['mw']}", "）"], "mw", "n0", True),
            plm("核心毛利率（含折舊）", "gm", "p1", True, "auto"),
            plm("雲端毛利率（含折舊）", "gmcg", "p1", False, "auto"),
            plm(["EBIT（", f"@PL!B{PR['A']['ebit']}", "）"], "ebit", "n2"),
            plm("營業利益率（含股票薪酬）", "om", "p1", True, "auto"),
            plm(["EPS（", f"@PL!B{PR['A']['eps']}", "）"], "eps", "n2"),
            plm(["資本支出（", f"@PL!B{PR['A']['cx']}", "）"], "cx", "n2", True)]},
        {"type": "table", "collapse": True, "summary": "展開：情境 B 完整收支區塊", "caption": f"@PL!A{PR['B']['mw'] - 2}",
         "head": [f"@PL!A{PR['B']['mw'] - 1}", f"@PL!B{PR['B']['mw'] - 1}"] + [f"@PL!{COL[y]}{PR['B']['mw'] - 1}" for y in YRS],
         "rowgen": {"sheet": "PL", "rows": f"{PR['B']['mw']}-{PR['B']['da']}", "cols": ["A", "B"] + [COL[y] for y in YRS], "fmts": ["", ""] + ["n3"] * 8,
                    "row_fmt": {str(PR["B"][k]): {COL[y]: f for y in YRS} for k, f in (("mw", "n0"), ("nmw", "n0"), ("pmw", "n2"), ("yoy", "p1"), ("gmc", "p1"), ("gmh", "p1"), ("gm", "p1"), ("gmcg", "p1"), ("sbcr", "p1"), ("om", "p1"), ("txr", "p1"), ("eps", "n2"), ("sh", "n3"))}}}]}
    # 估值
    SVr = SV
    dcf_rate_rows = [k for k in ("ke", "kd_at", "w_e", "wacc", "g", "ronic") if k in SVr]
    val_blocks = [
        {"type": "bars", "title": "目標價與現價", "items": [{"label": f"@Val!A{T['tp']}", "value": f"@Val!B{T['tp']}", "series": 0},
                                                       {"label": f"@Val!A{SV['dcf']}", "value": f"@Val!B{SV['dcf']}", "series": 1},
                                                       {"label": f"@Val!A{T['comp']}", "value": f"@Val!B{T['comp']}", "series": 2},
                                                       {"label": f"@Val!A{T['cons']}", "value": f"@Val!B{T['cons']}", "series": 1}],
         "ref": {"label": "現價", "value": f"@Val!B{T['px']}"}, "fmt": "n2"},
        {"type": "table", "caption": f"@Val!A{R_CON}", "head": [f"@Val!A{R_CH}", f"@Val!B{R_CH}", f"@Val!C{R_CH}", f"@Val!D{R_CH}"],
         "rows": [[f"@Val!A{T['tp']}", f"@Val!B{T['tp']}|n2", f"@Val!E{T['tp']}", ""], [f"@Val!A{T['comp']}", f"@Val!B{T['comp']}|n2", "", ""],
                  [f"@Val!A{T['cons']}", f"@Val!B{T['cons']}|n2", f"@Val!E{T['cons']}", ""], [f"@Val!A{T['up']}", f"@Val!B{T['up']}|sp1", "", ""],
                  [f"@Val!A{T['upc']}", f"@Val!B{T['upc']}|sp1", "", ""]]},
        {"type": "table", "collapse": True, "summary": "展開：結論的其他項目", "head": [f"@Val!A{R_CH}", f"@Val!B{R_CH}", f"@Val!C{R_CH}", f"@Val!D{R_CH}"],
         "rows": [[f"@Val!A{T['geps']}", f"@Val!B{T['geps']}", "", ""], [f"@Val!A{T['gmul']}", f"@Val!B{T['gmul']}|sp1", f"@Val!C{T['gmul']}|x1", f"@Val!D{T['gmul']}|x1"],
                  [f"@Val!A{T['band']}", f"@Val!B{T['band']}|n2", f"@Val!C{T['band']}|n2", ""], [f"@Val!A{T['ps']}", f"@Val!B{T['ps']}|x1", f"@Val!C{T['ps']}|x1", f"@Val!D{T['ps']}|x1"],
                  [f"@Val!A{T['meth']}", f"@Val!B{T['meth']}", "", ""]]},
        {"type": "table", "caption": "相對倍數調整（Comparable 目標倍數＝同業中位數 × 本公司相對同業的長期倍數比）", "head": ["項目", "數值"],
         "rows": [[f"@Val!A{T['rmk']}", f"@Val!B{T['rmk']}|n2"], [f"@Val!A{T['compadj']}", f"@Val!B{T['compadj']}|n2"]]},
        {"type": "note", "parts": [f"@Val!A{T['rmk'] - 1}"]},
        {"type": "note", "parts": [f"@Val!A{T['tier']}", "：", f"@Val!B{T['tier']}"]},
        {"type": "note", "parts": [f"@Val!A{T['rec']}", "：", f"@Val!B{T['rec']}", "（依上檔空間固定規則）"]},
        {"type": "note", "parts": ["@Val!A4"]},
    ]
    # DCF 折現率與彙總：依標籤定位
    def vfind(lbl, start=4, end=6 + FD.summary_rows(True)):
        for x in range(start, end):
            v = V.cell(x, 1).value
            if isinstance(v, str) and v.startswith(lbl): return x
        raise KeyError(lbl)
    rate = [(vfind("股權成本"), "p2"), (vfind("稅後債務成本"), "p2"), (vfind("權益占資本"), "p1"), (vfind("WACC"), "p2"), (vfind("永續成長率"), "p2"), (vfind("永續期 RONIC"), "p2")]
    val_blocks.append({"type": "table", "caption": "DCF 折現率：WACC、股權成本、稅後債務成本、RONIC、永續成長率", "head": ["項目", "數值", "說明"],
                       "rows": [[f"@Val!A{x}", f"@Val!B{x}|{f}", f"@Val!E{x}"] for x, f in rate]})
    main_ = [vfind("明確期 FCFF 現值"), vfind("終值現值"), vfind("企業價值 EV"), vfind("＋淨現金與投資"), vfind("＋Stub 期 FCFF"), vfind("淨現金滾到評價日"), vfind("DCF 情境價值推到 12 個月後"), vfind("DCF 目標價")]
    oth_ = [vfind("評價用稀釋股數"), vfind("DCF 情境每股價值"), vfind("  其中：終值現值占 EV"), vfind("  隱含終值")]
    VH = ["@Val!A5", "@Val!B5", "@Val!C5", "@Val!D5"]
    val_blocks.append({"type": "table", "caption": "DCF 彙總：明確期現值、終值現值、企業價值、淨現金與每股價值", "head": VH,
                       "rows": [[f"@Val!A{x}"] + ([f"@Val!B{x}|n2", "", ""] if x == main_[-1] else [f"@Val!{c}{x}|{'n2' if x == main_[-2] else 'n1'}" for c in "BCD"]) for x in main_]})
    ofm = {oth_[0]: "n3", oth_[1]: "n2", oth_[2]: "p1", oth_[3]: "x1"}
    val_blocks.append({"type": "table", "collapse": True, "summary": "展開：DCF 彙總的其他項目", "head": VH, "rows": [[f"@Val!A{x}"] + [f"@Val!{c}{x}|{ofm[x]}" for c in "BCD"] for x in oth_]})
    dcols = ["A"] + [get_column_letter(3 + i) for i in range(len(FC))]
    for s in SC:
        Pp = RD_["paths"][s]
        r_lab = min(v for v in Pp.values() if isinstance(v, int)) - 1
        r_last = max(v for v in Pp.values() if isinstance(v, int))
        val_blocks.append({"type": "table", "caption": f"@DCF!A{r_lab}", "collapse": True, "summary": "展開：本情境各年無槓桿自由現金流明細",
                           "head": [f"@DCF!A{RD_['hdr']}"] + [f"@DCF!{c}{RD_['hdr']}" for c in dcols[1:]],
                           "rowgen": {"sheet": "DCF", "rows": f"{r_lab + 1}-{r_last}", "cols": dcols, "fmts": [""] + ["n2"] * len(FC)}})
    val_blocks += [
        {"type": "table", "caption": f"@Val!A{R_C0}", "head": [f"@Val!A{R_PH}", f"@Val!B{R_PH}", f"@Val!C{R_PH}", f"@Val!D{R_PH}", f"@Val!E{R_PH}", f"@Val!F{R_PH}"],
         "rowgen": {"sheet": "Val", "rows": f"{R_P1}-{R_P2}", "cols": ["A", "B", "C", "D", "E", "F"], "fmts": ["", "", "x1", "", "", "tag"]}, "collapse": True, "summary": "展開：同業前瞻 P/S"},
        {"type": "note", "parts": [f"@Val!E{R_PNOTE}"]},
        {"type": "table", "head": VH, "caption": "Comparable 彙總：同業倍數 × 前瞻營收",
         "rows": [[f"@Val!A{R_MED}", f"@Val!B{R_MED}|x1", "", ""], [f"@Val!A{R_FR}", f"@Val!B{R_FR}|n2", f"@Val!C{R_FR}|n2", f"@Val!D{R_FR}|n2"],
                  [f"@Val!A{R_FRW}", f"@Val!B{R_FRW}|n2", "", ""], [f"@Val!A{R_SH}", f"@Val!B{R_SH}|n3", "", ""], [f"@Val!A{R_COMP}", f"@Val!B{R_COMP}|n2", "", ""]]},
        {"type": "table", "caption": "現價與市場隱含倍數", "head": [f"@Val!A{R_H1}", f"@Val!B{R_H1}", f"@Val!C{R_H1}", f"@Val!D{R_H1}"], "collapse": True, "summary": "展開：現價與市場隱含倍數",
         "rows": [[f"@Val!A{T['px']}", f"@Val!B{T['px']}|n2", "", ""], [f"@Val!A{T['mcap']}", f"@Val!B{T['mcap']}|n2", "", ""], [f"@Val!A{T['ev']}", f"@Val!B{T['ev']}|n2", "", ""],
                  [f"@Val!A{T['evs27']}", f"@Val!B{T['evs27']}|x1", f"@Val!C{T['evs27']}|x1", f"@Val!D{T['evs27']}|x1"],
                  [f"@Val!A{T['evs28']}", f"@Val!B{T['evs28']}|x1", f"@Val!C{T['evs28']}|x1", f"@Val!D{T['evs28']}|x1"]]},
    ]
    # 對帳
    RH = [f"@Recon!{c}3" for c in "ABCDEFGH"]
    def rrow(k, f):
        n = RCR[k]; rel = RC[f"F{n}"].value is not None
        return [f"@Recon!A{n}", f"@Recon!B{n}|{f}", f"@Recon!C{n}|{f}", f"@Recon!D{n}|{f}", f"@Recon!E{n}|{f}" if RC[f"E{n}"].value is not None else "",
                f"@Recon!F{n}|sp1" if rel else "", f"@Recon!G{n}", f"@Recon!H{n}"]
    RF = {"rev26": "n3", "gm26": "p1", "om26": "p1", "x27": "x1", "rpo24": "n2", "rpo48": "n2", "mw28": "n0", "oai": "n2", "tkhold": "n2", "tkcx": "n2", "rev27": "n2", "ebit28": "n2", "cx27": "n2"}
    main_r = ["x27", "rpo24", "rpo48", "tkhold", "cx27"]
    recon = [{"type": "table", "head": RH, "caption": "對帳要點：公司指引、RPO、容量、Tokenomics 與現金", "rows": [rrow(k, RF[k]) for k in main_r]},
             {"type": "table", "collapse": True, "summary": "展開：其他對帳項目", "head": RH, "rows": [rrow(k, RF[k]) for k in RF if k not in main_r]}]
    # 敏感度
    cr = [x[0] for x in case_rows]
    big = lambda: None
    sens = [{"type": "tornado", "title": "情境 B 加權價值相對基準的變動（一次動一個輸入）", "sheet": "Sens", "rows": ",".join(str(x) for x in cr), "label_col": "A", "value_col": "J", "fmt": "sp1"},
            {"type": "table", "collapse": True, "summary": "展開：全部案例的 FY2028 營收、EPS、DCF 與加權價值",
             "head": [f"@Sens!{c}3" for c in "ABCDEFGHIJ"],
             "rowgen": {"sheet": "Sens", "rows": f"4-{CASE_END}", "cols": list("ABCDEFGHIJ"), "fmts": ["", "", "n2", "sp1", "n2", "sp1", "n2", "sp1", "n2", "sp1"]}},
            {"type": "table", "caption": f"@Sens!A{STD_SENS[0] - 2}", "head": [f"@Sens!{c}{STD_SENS[0] - 1}" for c in "ABCDEFG"],
             "rowgen": {"sheet": "Sens", "rows": f"{STD_SENS[0]}-{STD_SENS[1]}", "cols": list("ABCDEFG"), "fmts": ["", "", "n2", "sp1", "n2", "sp1", "n2"]}}]
    H = {"company": "Cerebras Systems", "ticker": "CBRS", "thesis": "@README!A32", "question": "OpenAI 合約之後，Cerebras 能否把推論雲端擴大到 OpenAI 以外，並以自有現金完成建置",
         "scenarios": [{"key": s, "name": f"情境 {s}"} for s in SC],
         "sections": [
             {"id": "thesis", "title": "命題與核心問題", "blocks": [
                 {"type": "h3", "text": "命題"}, {"type": "list", "ordered": True, "items": ["@README!A33", "@README!A34", "@README!A35", "@README!A36"]},
                 {"type": "list", "ordered": True, "collapse": True, "summary": "展開：命題回覆（依本模型計算的答案；想先看結論可展開，或讀完全文再看）", "items": ["@README!A43", "@README!A44", "@README!A45", "@README!A46"]},
                 {"type": "h3", "text": "重點觀察"}, {"type": "list", "ordered": True, "items": ["@README!A37", "@README!A38", "@README!A39", "@README!A40"]},
                 {"type": "h3", "text": "情境定義與會計年度"},
                 {"type": "kv", "items": [["情境 A", f"@PL!A{PR['A']['mw'] - 2}"], ["情境 B", f"@PL!A{PR['B']['mw'] - 2}"], ["情境 C", f"@PL!A{PR['C']['mw'] - 2}"], ["會計年度", "@README!A6"]]}]},
             {"id": "drivers", "title": "驅動因子 → 推導量", "blocks": drivers},
             pl_sec,
             {"id": "valuation", "title": "估值", "intro": "@Val!A1", "blocks": val_blocks},
             {"id": "recon", "title": "對帳", "intro": "@Recon!A1", "blocks": recon},
             {"id": "sens", "title": "敏感度", "intro": "@Sens!A1", "blocks": sens},
             {"id": "status", "title": "資料狀態", "blocks": [
                 {"type": "kv", "items": [["模型版本", ["$version", "（", "$date", "）"]], ["來源 xlsx", "$xlsx"], ["來源 xlsx 的雜湊指紋（取代建置 commit）", "$sha"],
                                          ["Tokenomics 快照版本", f"@In!H{inrow('tk_hold_cs')}"]]},
                 {"type": "h3", "text": "上游成品"}, {"type": "upstream", "sheets": ["In"]},
                 {"type": "h3", "text": "待複核項目"},
                 {"type": "list", "items": ["公司原文為財報新聞稿與季報（sec.gov）；招股說明書、OpenAI 合約條款與客戶集中只經二手摘要，原文待人工複核。", "@Q!A2",
                                            "標為「未找到」的項目（見下方清單）待找到來源後補上。"]},
                 {"type": "h3", "text": "已知限制"}, {"type": "note", "parts": ["@README!A16"]},
                 {"type": "h3", "text": "附錄：版本要點"},
                 {"type": "table", "collapse": True, "summary": "展開：本版要點", "head": ["要點"], "rows": [["@README!A49"]]},
                 {"type": "h3", "text": "「未找到」清單"},
                 {"type": "scan", "needle": "未找到", "collapse": True, "summary": "展開：標為「未找到」的輸入與來源", "sheets": {"In": ["B"], "Q": ["A"], "Val": ["A"], "Sources": ["B"]}, "empty": "沒有標為未找到的項目。"}]}],
         "wide_text": True,
         "hero": {"rec": f"@Val!B{T['rec']}",
                  "tiles": [[[f"@Val!A{T['tp']}", "（美元）"], f"@Val!B{T['tp']}|n2", ["上檔空間 ", f"@Val!B{T['up']}|sp1"]],
                            [["只看 Comparable 目標價（市場主流同業法）", "（美元）"], f"@Val!B{T['comp']}|n2", ["上檔空間 ", f"@Val!B{T['upc']}|sp1"]],
                            [[f"@Val!A{T['px']}", "（美元）"], f"@Val!B{T['px']}|n2", ""]],
                  "note": ["投資建議依加權目標價判定：", f"@Val!E{T['rec']}"]}}
    json.dump(H, open(os.path.join(HERE, "html_src.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)   # v0.1.2：html.json 由 gen_html_json.py <xlsx> 依版面 v1.3 轉出
