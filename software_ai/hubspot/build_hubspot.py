#!/usr/bin/env python3
"""HubSpot（HUBS）個股模型 v0.1 建置腳本（AI 軟體應用 Project；原型 S：按席位計價，AI 代理按任務計價）。

用法（由 make_v01.sh 呼叫；敏感度與反解由 敏感度程式 呼叫本檔）：
  python3 build_hubspot.py <輸出 raw.xlsx> [--sens data/sens_v01.json] [--kreq data/kreq_v01.json] [--set 鍵=值 ...]
公開 repo：只含公司與市場事實，不含任何個人情境。本版未做先驗。
"""
import datetime as dt
import hashlib
import json
import os
import sys

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L_
from openpyxl.workbook.defined_name import DefinedName

import hs_inputs as HI
from hs_inputs import D, HCOLS, HIST, HLAB, HSUF, INP, SCN_ROWS, SEC

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
VERSION = "v0.1"
BUILD_DATE = "2026-11-16"
FILE_DATE = "20261116"
YEARS = list(range(2023, 2033))
YC = {y: L_(3 + i) for i, y in enumerate(YEARS)}          # C..L；FY2026E＝F
FC = list(range(2027, 2033))
SC = ["A", "B", "C"]
SCN_NAME = {"A": "情境 A（席位取代慢、agent 放量）", "B": "情境 B（基準）", "C": "情境 C（席位取代快、agent 慢）"}
BLUE, BOLD = Font(color="0000FF"), Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")

# ───────── 列號對照（模組層級；gen_html_json.py 匯入）─────────
P_OFF = {"title": 0, "yr": 1, "sub": 2, "ps": 3, "rev": 4, "g": 5, "chk": 6, "cust": 7, "asrpc": 8, "chk2": 9, "cogs": 10, "gp": 11, "gm": 12, "opx": 13,
         "ebit": 14, "ebitm": 15, "oth": 16, "pt": 17, "tax": 18, "ni": 19, "nim": 20, "eps": 21, "nig": 22, "sh": 23, "she": 24, "ngcogs": 25, "ngopx": 26,
         "ngop": 27, "ngopm": 28, "sbc": 29, "adj": 30, "restr": 31, "ngeps": 32, "inf": 33}
C_OFF = {"title": 0, "yr": 1, "rev": 2, "gr": 3, "ebit": 4, "tax": 5, "nopat": 6, "da": 7, "cx": 8, "wc": 9, "sbcab": 10, "fcff": 11, "int": 12, "fcf": 13, "bb": 14,
         "cash": 15, "t": 16, "cf": 17, "df": 18, "pv": 19, "spv": 20, "n33": 21, "tv": 22, "tvdf": 23, "tvpv": 24, "ev": 25, "nc": 26, "eq": 27, "shv": 28, "ps0": 29,
         "ps1": 30, "tvsh": 31, "gate": 32, "mincash": 33, "sbcpv": 34}
T_OFF = {"title": 0, "yr": 1, "n": 2, "na": 3, "cust": 4, "avg": 5, "gup": 6, "e": 7, "arpc0": 8, "arpc": 9, "core": 10, "adopt": 11, "tasks": 12, "price": 13,
         "aiRev": 14, "sub": 15, "ps": 16, "rev": 17, "gr": 18, "loss": 19, "ttot": 20, "cpt": 21, "pcost": 22, "emb": 23, "inc": 24, "net": 25, "npc": 26,
         "brk": 27, "share": 28, "chk": 29}
PB = {s: {k: 5 + i * 37 + v for k, v in P_OFF.items()} for i, s in enumerate(SC)}
CB = {s: {k: 5 + i * 37 + v for k, v in C_OFF.items()} for i, s in enumerate(SC)}
TB = {s: {k: 14 + i * 33 + v for k, v in T_OFF.items()} for i, s in enumerate(SC)}
# 標準損益表（情境 B；HTML 用）放在 PL 頁最後
SIS0 = 5 + 3 * 37 + 1
SIS = {k: SIS0 + i for i, k in enumerate(["title", "yr", "rev", "g", "cogs", "gp", "gm", "opx", "ebit", "ebitm", "oth", "pt", "tax", "ni", "nim", "eps", "nig", "ngeps", "ngop", "ngopm"])}

# Q4E 頁列號
Q4 = {k: 5 + i for i, k in enumerate(["cust", "avg", "asrpc", "sub", "ps", "rev", "ngcogs", "ngopx", "ngop", "sbc", "adj", "restr", "cogs", "ebit", "oth", "pt", "tax", "ni", "da", "cx", "nwc",
                                      "nopat", "fcff", "fcf", "bb", "cash", "ngeps", "eps"])}


def f_(ws, cell, v, fmt=None, bold=False, blue=False, fill=None, wrap=False):
    c = ws[cell]
    c.value = v
    if fmt:
        c.number_format = fmt
    if bold:
        c.font = BOLD
    if blue:
        c.font = BLUE
    if fill:
        c.fill = fill
    if wrap:
        c.alignment = Alignment(wrap_text=True, vertical="top")
    return c


def addname(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def tf(*parts):
    out = []
    for p in parts:
        if isinstance(p, str):
            for i in range(0, len(p), 200):
                out.append('"' + p[i:i + 200].replace('"', '""') + '"')
        else:
            out.append(f'TEXT({p[0]},"{p[1]}")')
    return "=" + "&".join(out)


def tier_f(cell):
    return f'IF({cell}>0.5,"強力買進",IF({cell}>0.25,"買進",IF({cell}>=-0.25,"中立",IF({cell}>=-0.5,"放空","強力放空"))))'


def parse_ov(argv):
    ov = {}
    for i, a in enumerate(argv):
        if a == "--set":
            k, v = argv[i + 1].split("=", 1)
            ov[k] = float(v)
    return ov


EXTRA_CASES = [   # (標籤, 設定文字, 方向, [(鍵, 值)...], 鍵欄文字)
    ("token 價格與任務售價聯動（價格 2 倍）", "tk_pscale＝2；link_pass＝1（成本與售價同步）", "高", [("tk_pscale", 2.0), ("link_pass", 1.0)], "tk_pscale+link_pass"),
    ("token 價格與任務售價聯動（價格 0.5 倍）", "tk_pscale＝0.5；link_pass＝1（成本與售價同步）", "低", [("tk_pscale", 0.5), ("link_pass", 1.0)], "tk_pscale+link_pass"),
]
SIDE_NOTE = {"tk_pscale": "（上界：只動成本一邊，售價不連動；聯動案另列）", "route": "（路由到便宜模型：全部走 Luna）", "link_pass": "（價格傳導到售價）"}


def case_jobs():
    """敏感度案例清單：[(key, side, label, setting, [(k, v)])]，供 敏感度程式 使用"""
    out = []
    for it in INP:
        if it[0] == SEC:
            continue
        key, lab, unit, lo, base, hi = it[:6]
        for side, val in (("低", lo), ("高", hi)):
            if val is None:
                continue
            if not isinstance(base, (str, D)) and abs(val - base) < 1e-12:
                continue
            if isinstance(base, D):
                continue
            fm = (lambda x: f"{x:.1%}" if unit == "%" else f"{x:g}")
            bs = "公式" if isinstance(base, str) else fm(base)
            out.append((key, side, f"{lab.split('（')[0]}：{side}{SIDE_NOTE.get(key, '') if side == '高' or key == 'route' else ''}",
                        f"低／基準／高＝{fm(lo) if lo is not None else '—'}／{bs}／{fm(hi) if hi is not None else '—'}", [(key, val)]))
    for lab, st, side, sets, kc in EXTRA_CASES:
        out.append((kc, side, lab, st, sets))
    return out


def build(out, ov=None, sens=None, kreq=None):
    ov = ov or {}
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    names = ["README", "In", "Hist", "Q4E", "Thesis", "PL", "CF", "Val", "RevDCF", "Recon", "Sens", "Sources"]
    W = {n: wb.create_sheet(n) for n in names}
    keyrow = {}
    # ───────── In ─────────
    In = W["In"]
    f_(In, "A1", "In — 輸入（藍字；基準欄 E 進主模型；D／F＝敏感度低／高，只用於 Sens 頁、不進目標價；情境 A／B／C 的驅動見本頁最後一節「情境驅動對照」）。金額 US$M；每股 US$；股數百萬股。標記：Verified／Interested-party／Analogy／Assumed／Derived；Interested-party 列須有來源誘因。")
    f_(In, "A2", "事實總帳：data/facts_hubspot.csv（逐列出處）與 data/research_hubspot.md；Tokenomics v22.0 快照 data/tokenomics_snapshot_v22.0.json。本版未做先驗。")
    for j, h in enumerate(["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源", "說明", "來源誘因（Interested-party 列）"], 1):
        f_(In, f"{L_(j)}3", h, bold=True, fill=HDR)
    r = 4
    for it in INP:
        if it[0] == SEC:
            f_(In, f"A{r}", it[1], bold=True)
            r += 1
            continue
        key, lab, unit, lo, base, hi, tag, src, note, inc = it
        if key in ov:
            base = ov[key]
        keyrow[key] = r
        f_(In, f"A{r}", key)
        f_(In, f"B{r}", lab)
        f_(In, f"C{r}", unit)
        if lo is not None:
            f_(In, f"D{r}", lo, blue=True)
        f_(In, f"E{r}", base, blue=not (isinstance(base, str) and base.startswith("=")))
        if hi is not None:
            f_(In, f"F{r}", hi, blue=True)
        f_(In, f"G{r}", tag)
        f_(In, f"H{r}", src)
        if note:
            f_(In, f"I{r}", note)
        if inc:
            f_(In, f"J{r}", inc)
        if isinstance(base, D):
            In[f"E{r}"].number_format = "yyyy-mm-dd"
        elif unit == "%":
            for c in "DEF":
                In[f"{c}{r}"].number_format = "0.0%"
        addname(wb, key, f"In!$E${r}")
        r += 1
    # 派生比率鍵（不進 In 區間表，放在 Thesis 頂端）：cogs_adj_sh 在 Hist 後定義
    r += 1
    f_(In, f"A{r}", "J. 情境驅動對照（三情境並排，進加權目標價；本節只引用上方各列，修改請改上方藍字）", bold=True)
    r += 1
    for j, h in enumerate(["鍵", "項目", "單位", "情境 A（取代慢）", "情境 B（基準）", "情境 C（取代快）", "標記", "說明"], 1):
        f_(In, f"{L_(j)}{r}", h, bold=True, fill=HDR)
    SCN_HDR = r
    r += 1
    SCNROW = {}
    for k, lab, unit, ks in SCN_ROWS:
        f_(In, f"A{r}", k)
        f_(In, f"B{r}", lab)
        f_(In, f"C{r}", unit)
        for c, kk in zip("DEF", ks):
            f_(In, f"{c}{r}", f"=E{keyrow[kk]}", "0.0%" if unit == "%" else None)
        f_(In, f"G{r}", "Derived")
        f_(In, f"H{r}", "引用 In 頁 " + "／".join(ks) + "；說明見上方各列")
        SCNROW[k] = r
        r += 1
    In.column_dimensions["B"].width = 62
    In.column_dimensions["H"].width = 60
    In.column_dimensions["I"].width = 50
    # ───────── Hist ─────────
    H = W["Hist"]
    f_(H, "A1", "Hist — 歷史財務（US$M；股數百萬股；FY2023 為 FY2025 10-K 重編後數；來源：10-K FY2025、10-K FY2023、各季新聞稿與 10-Q，逐列出處見 data/facts_hubspot.csv）。公式只引用本頁。", bold=True)
    for j, h in enumerate(["項目", "鍵"] + HLAB + ["標記", "來源", "說明／來源誘因"], 1):
        f_(H, f"{L_(j)}2", h, bold=True, fill=HDR)
    hr = {k: 3 + i for i, (k, *_r) in enumerate(HIST)}
    for key, lab, vals, tag in HIST:
        rr = hr[key]
        f_(H, f"A{rr}", lab)
        f_(H, f"B{rr}", key)
        for c, sf, v in zip(HCOLS, HSUF, vals):
            if v == "sum":
                f_(H, f"{c}{rr}", f"=H{rr}+G{rr}", "#,##0.000")
            elif isinstance(v, str):
                f_(H, f"{c}{rr}", v.format(r=rr, **hr), "#,##0.000")
            elif v is None:
                f_(H, f"{c}{rr}", "未找到")
            else:
                f_(H, f"{c}{rr}", v, "#,##0.000" if abs(v) < 1e5 else "#,##0", blue=True)
            if v is not None or c == "I":
                addname(wb, f"h_{key}_{sf}", f"Hist!${c}${rr}")
        f_(H, f"M{rr}", tag)
        f_(H, f"N{rr}", "10-K FY2025／10-K FY2023／各季新聞稿與 10-Q（facts_hubspot.csv）")
        if tag == "Interested-party":
            f_(H, f"O{rr}", "來源誘因：公司自訂 non-GAAP 或 KPI 口徑，傾向呈現成長與獲利改善；定義可變更（例：Customers 自 2026Q3 起納入 YouSpot Pro 訂閱者）")
        elif key == "cogs":
            f_(H, f"O{rr}", "FY2023 取 FY2025 10-K 重編後數；FY2023 10-K 原數為營業損失 −208.056、訂閱成本 290.802，重編原因未找到")
        elif key == "rev_ltm1":
            f_(H, f"O{rr}", "＝9M25 營收＋Q4'24 營收 703.172（Q4 2025 新聞稿比較欄）")
    # 欄 L 在 Hist 為 Q3'26，標記欄實際從 M 起：修正表頭
    for j, h in enumerate(["項目", "鍵"] + HLAB + ["標記", "來源", "說明／來源誘因"], 1):
        f_(H, f"{L_(j)}2", h, bold=True, fill=HDR)
    rr = 3 + len(HIST) + 1
    f_(H, f"A{rr}", "核對：訂閱＋專業服務 − 營收合計（應為 0）", bold=True)
    for c in HCOLS:
        if c in ("F", "G", "H", "I", "J", "K", "L", "C", "D", "E"):
            f_(H, f"{c}{rr}", f"=IFERROR({c}{hr['sub']}+{c}{hr['ps']}-{c}{hr['rev']},\"n/a\")", "0.000")
    addname(wb, "cogs_adj_sh", f"Hist!$C${rr + 2}")
    f_(H, f"A{rr + 2}", "GAAP 營業成本與非 GAAP 營業成本之差 ÷ 非 GAAP 調節合計（近四季；成本內含的 SBC 與攤銷占調節比）")
    f_(H, f"C{rr + 2}", "=(h_cogs_ltm-h_ngcogs_ltm)/(h_ngoi_ltm-h_oi_ltm)", "0.0%")
    H.column_dimensions["A"].width = 52
    # ───────── Q4E ─────────
    Q = W["Q4E"]
    f_(Q, "A1", "Q4E — Q4 2026 預測（US$M；單季；三情境相同）。Q4 為尚未公布的一季：客戶淨增、ASRPC 季增為 Assumed；成本率持平於近四季（工作單 G2.3）；重組費用為公司預估。公司指引只在 Recon 對照。", bold=True)
    f_(Q, "A2", "FY2026E＝9M26 實績（Hist）＋本頁；Q4 自由現金流為 DCF 的第一期（折現期中點＝距評價日 0.5 個單季）。")
    for j, h in enumerate(["項目", "Q4 2026E", "說明"], 1):
        f_(Q, f"{L_(j)}4", h, bold=True, fill=HDR)
    q = Q4

    def qrow(k, lab, fm, nf="#,##0.0", note=""):
        f_(Q, f"A{q[k]}", lab)
        f_(Q, f"B{q[k]}", fm, nf)
        if note:
            f_(Q, f"C{q[k]}", note)
    qrow("cust", "期末客戶數（戶）", "=h_cust_q326+na_q4", "#,##0", "Q3 期末 311,704＋淨增")
    qrow("avg", "平均客戶數（戶）", f"=AVERAGE(h_cust_q326,B{q['cust']})", "#,##0")
    qrow("asrpc", "ASRPC（US$，年化）", "=h_asrpc_q326*(1+q4_asrpc)", "#,##0", "ASRPC＝訂閱營收×4÷平均客戶（Q3 實績驗算 12,038）")
    qrow("sub", "訂閱營收", f"=B{q['avg']}*B{q['asrpc']}/4/1000000")
    qrow("ps", "專業服務與其他營收", f"=ps_pct*B{q['sub']}")
    qrow("rev", "營收合計", f"=B{q['sub']}+B{q['ps']}")
    qrow("ngcogs", "非 GAAP 營業成本", f"=ng_cogs*B{q['rev']}")
    qrow("ngopx", "非 GAAP 營業費用", f"=ng_opex*B{q['rev']}")
    qrow("ngop", "非 GAAP 營業利益", f"=B{q['rev']}-B{q['ngcogs']}-B{q['ngopx']}")
    qrow("sbc", "股票薪酬", f"=sbc_pct*B{q['rev']}")
    qrow("adj", "非 GAAP 調節其他項", f"=adj_pct*B{q['rev']}")
    qrow("restr", "重組費用（一次性）", "=restr_q4")
    qrow("cogs", "營業成本（GAAP）", f"=B{q['ngcogs']}+cogs_adj_sh*(B{q['sbc']}+B{q['adj']}+B{q['restr']})")
    qrow("ebit", "營業利益（GAAP）", f"=B{q['ngop']}-B{q['sbc']}-B{q['adj']}-B{q['restr']}")
    qrow("oth", "其他收入（現金與投資收益）", "=int_y*cash_0/4")
    qrow("pt", "稅前利益", f"=B{q['ebit']}+B{q['oth']}")
    qrow("tax", "所得稅", f"=MAX(0,B{q['pt']})*tax")
    qrow("ni", "稅後淨利（GAAP）", f"=B{q['pt']}-B{q['tax']}")
    qrow("da", "折舊與攤銷", f"=da_pct*B{q['rev']}")
    qrow("cx", "資本支出（含資本化軟體）", f"=cx_pct*B{q['rev']}")
    qrow("nwc", "營運資金流入", f"=nwc_k*(B{q['rev']}-h_rev_q425)")
    qrow("nopat", "NOPAT", f"=IF(B{q['ebit']}>0,B{q['ebit']}*(1-tax),B{q['ebit']})")
    qrow("fcff", "FCFF（SBC 費用化；sbc_ab＝1 時加回）", f"=B{q['nopat']}+B{q['da']}-B{q['cx']}+B{q['nwc']}+sbc_ab*B{q['sbc']}", note="第一期 DCF 現金流（折現期中點）")
    qrow("fcf", "自由現金流（現金口徑近似）", f"=B{q['fcff']}+(1-sbc_ab)*B{q['sbc']}+B{q['oth']}*(1-IF(B{q['pt']}>0,tax,0))")
    qrow("bb", "回購", "=bb_q4")
    qrow("cash", "期末現金與投資", f"=cash_0+B{q['fcf']}-B{q['bb']}")
    qrow("ngeps", "非 GAAP EPS（US$）", f"=(B{q['ngop']}+B{q['oth']})*(1-tax)/sh_eps0")
    qrow("eps", "GAAP EPS（US$）", f"=B{q['ni']}/sh_eps0")
    Q.column_dimensions["A"].width = 46
    Q.column_dimensions["C"].width = 60
    # ───────── Thesis ─────────
    T = W["Thesis"]
    f_(T, "A1", "Thesis — 命題模組（原型 S：按席位計價）：每客戶訂閱營收＝上年×(1＋年增−席位被 AI 取代侵蝕)；AI 淨效果＝按任務計價營收（增量）− 席位取代損失（侵蝕）− 增量推論成本。US$M；單位為戶。", bold=True)
    f_(T, "A2", "推論成本：每任務 token 取 Tokenomics 工作負載（公司未揭露 token，標 Analogy）× 市場價（組合價）÷ 成功率；內嵌（未計價）AI 成本另計，兩者只扣增量（基期已含在近四季成本率）。價格年降標「待 AI供應鏈 登錄」。")
    f_(T, "A4", "有效 token 價格與每任務成本（US$；所有情境共用）", bold=True)
    pr = {"pf": 5, "pc": 6, "po": 7, "cost0": 8}
    for k, lab, fm, nf in [("pf", "有效新鮮輸入價（US$／百萬 token；組合價，route＝1 時全 Luna）", "=IF(route=1,tkp_fL,mix_L*tkp_fL+mix_T*tkp_fT+mix_S*tkp_fS)", "0.0000"),
                           ("pc", "有效快取輸入價", "=IF(route=1,tkp_cL,mix_L*tkp_cL+mix_T*tkp_cT+mix_S*tkp_cS)", "0.0000"),
                           ("po", "有效輸出價", "=IF(route=1,tkp_oL,mix_L*tkp_oL+mix_T*tkp_oT+mix_S*tkp_oS)", "0.0000"),
                           ("cost0", "每解決任務推論成本（FY2026 價格；US$）＝倍數×(新鮮×價＋快取×價＋輸出×價)÷10^6 × 價格倍數 ÷ 成功率",
                            "=ctx_m*(tok_fresh*C5+tok_cache*C6+tok_dec*C7)/1000000*tk_pscale/succ", "0.00000")]:
        f_(T, f"A{pr[k]}", lab)
        f_(T, f"C{pr[k]}", fm, nf)
    addname(wb, "tk_cost0", "Thesis!$C$8")
    f_(T, "A9", "對照：每次解決對話售價（US$）／推論成本占售價比")
    f_(T, "C9", "=price_0", "0.00")
    f_(T, "D9", "=C8/price_0", "0.0%")
    f_(T, "A10", "推論成本是否吃掉單價：FY2026 價格下每任務推論成本占售價比（右）；損益兩平單價見各情境區塊「損益兩平每任務單價」")
    f_(T, "A11", "內嵌（未計價）AI 成本：FY2026 占營收 ＝ In 頁 emb_pct；使用量年增 emb_g；token 價格年變動 pd_tok（待 AI供應鏈 登錄）")
    for s in SC:
        R = TB[s]
        f_(T, f"A{R['title']}", f"{SCN_NAME[s]}：命題模組", bold=True)
        labels = {"yr": "年度", "n": "年數（自 FY2026 起）", "na": "客戶淨增率（期末年增）", "cust": "期末客戶數（戶）", "avg": "平均客戶數（戶）",
                  "gup": "每客戶訂閱營收年增（不含席位取代）", "e": "席位被 AI 取代的年侵蝕＝取代率 × 席位占比", "arpc0": "每客戶訂閱營收（無侵蝕路徑，US$）",
                  "arpc": "每客戶核心訂閱營收（含侵蝕，US$）", "core": "核心訂閱營收（US$M）＝平均客戶 × 每客戶", "adopt": "付費採用 agent 的客戶占比",
                  "tasks": "每採用客戶每年任務數（次）", "price": "每任務單價（US$）", "aiRev": "AI 增量：按任務計價營收（US$M）", "sub": "訂閱營收（US$M）＝核心＋按任務計價",
                  "ps": "專業服務與其他營收（US$M）", "rev": "營收合計（US$M）", "gr": "營收年增率", "loss": "AI 侵蝕：席位被取代的營收損失（US$M，累積）",
                  "ttot": "按任務計價的總任務數（百萬次）", "cpt": "每任務推論成本（US$，含價格年變動）", "pcost": "按任務計價的推論成本（US$M）",
                  "emb": "內嵌（未計價）AI 推論成本（US$M）", "inc": "增量推論成本（US$M）＝(按任務成本 − FY2026) ＋ (內嵌 − 基期成本率已含部分)",
                  "net": "AI 淨效果（US$M）＝按任務營收 − 席位取代損失 − 增量推論成本", "npc": "AI 淨效果每客戶每年（US$）", "brk": "損益兩平每任務單價（US$；使 AI 淨效果＝0）",
                  "share": "推論成本（按任務＋內嵌）占非 GAAP 營業成本比", "chk": "檢查：營收 − (核心＋按任務計價＋專業服務)"}
        for k, lab in labels.items():
            f_(T, f"A{R[k]}", lab)
        for y in range(2025, 2033):
            c, pc = YC[y], YC[y - 1]
            n = f"{c}{R['n']}"
            if y == 2025:
                f_(T, f"{c}{R['yr']}", "FY2025A", bold=True)
                f_(T, f"{c}{R['cust']}", "=h_cust_25", "#,##0")
                continue
            f_(T, f"{c}{R['yr']}", f"FY{y}{'E' if y >= 2026 else 'A'}", bold=True)
            f_(T, f"{c}{R['n']}", y - 2026)
            if y == 2026:
                f_(T, f"{c}{R['cust']}", f"=Q4E!B{Q4['cust']}", "#,##0")
                f_(T, f"{c}{R['avg']}", f"=AVERAGE({pc}{R['cust']},{c}{R['cust']})", "#,##0")
                f_(T, f"{c}{R['arpc0']}", f"=(h_sub_9m26+Q4E!B{Q4['sub']}-ai_rev_26)*1000000/{c}{R['avg']}", "#,##0.0")
                f_(T, f"{c}{R['arpc']}", f"={c}{R['arpc0']}", "#,##0.0")
                f_(T, f"{c}{R['adopt']}", f"=ai_rev_26*1000000/({c}{R['avg']}*tasks_0*price_0)", "0.00%")
                f_(T, f"{c}{R['tasks']}", "=tasks_0", "#,##0")
                f_(T, f"{c}{R['price']}", "=price_0", "0.000")
                f_(T, f"{c}{R['sub']}", f"=h_sub_9m26+Q4E!B{Q4['sub']}", "#,##0.0")
                f_(T, f"{c}{R['ps']}", f"=h_ps_9m26+Q4E!B{Q4['ps']}", "#,##0.0")
                f_(T, f"{c}{R['loss']}", 0, "#,##0.0")
                f_(T, f"{c}{R['gr']}", f"={c}{R['rev']}/h_rev_25-1", "0.0%")
            else:
                f_(T, f"{c}{R['na']}", f"=na_27+(na_32-na_27)*({n}-1)/5", "0.0%")
                f_(T, f"{c}{R['cust']}", f"={pc}{R['cust']}*(1+{c}{R['na']})", "#,##0")
                f_(T, f"{c}{R['avg']}", f"=AVERAGE({pc}{R['cust']},{c}{R['cust']})", "#,##0")
                f_(T, f"{c}{R['gup']}", "=g_up", "0.0%")
                f_(T, f"{c}{R['e']}", f"=repl_{s}*seat_sh", "0.00%")
                f_(T, f"{c}{R['arpc0']}", f"={pc}{R['arpc0']}*(1+g_up)", "#,##0.0")
                f_(T, f"{c}{R['arpc']}", f"={pc}{R['arpc']}*(1+g_up-{c}{R['e']})", "#,##0.0")
                f_(T, f"{c}{R['adopt']}", f"=$F{R['adopt']}+(adopt_{s}-$F{R['adopt']})*{n}/6", "0.00%")
                f_(T, f"{c}{R['tasks']}", f"={pc}{R['tasks']}*(1+g_tasks)", "#,##0")
                f_(T, f"{c}{R['price']}", f"=price_0*(1+dp_{s})^{n}*tk_pscale^link_pass", "0.000")
                f_(T, f"{c}{R['sub']}", f"={c}{R['core']}+{c}{R['aiRev']}", "#,##0.0")
                f_(T, f"{c}{R['ps']}", f"=ps_pct*{c}{R['sub']}", "#,##0.0")
                f_(T, f"{c}{R['loss']}", f"={c}{R['avg']}*({c}{R['arpc0']}-{c}{R['arpc']})/1000000", "#,##0.0")
                f_(T, f"{c}{R['gr']}", f"={c}{R['rev']}/{pc}{R['rev']}-1", "0.0%")
            f_(T, f"{c}{R['core']}", f"={c}{R['avg']}*{c}{R['arpc']}/1000000", "#,##0.0")
            f_(T, f"{c}{R['aiRev']}", f"={c}{R['avg']}*{c}{R['adopt']}*{c}{R['tasks']}*{c}{R['price']}/1000000", "#,##0.0")
            f_(T, f"{c}{R['rev']}", f"={c}{R['sub']}+{c}{R['ps']}", "#,##0.0")
            f_(T, f"{c}{R['ttot']}", f"={c}{R['avg']}*{c}{R['adopt']}*{c}{R['tasks']}/1000000", "#,##0.00")
            f_(T, f"{c}{R['cpt']}", f"=tk_cost0*(1+pd_tok)^{n}", "0.00000")
            f_(T, f"{c}{R['pcost']}", f"={c}{R['ttot']}*{c}{R['cpt']}", "#,##0.00")
            f_(T, f"{c}{R['emb']}", f"=emb_pct*$F{R['rev']}*((1+emb_g)*(1+pd_tok))^{n}*IF({n}>0,tk_pscale,1)", "#,##0.0")
            f_(T, f"{c}{R['inc']}", f"=({c}{R['pcost']}-$F{R['pcost']})+({c}{R['emb']}-emb_pct*{c}{R['rev']})", "#,##0.0")
            f_(T, f"{c}{R['net']}", f"={c}{R['aiRev']}-{c}{R['loss']}-{c}{R['inc']}", "#,##0.0")
            f_(T, f"{c}{R['npc']}", f"={c}{R['net']}/{c}{R['avg']}*1000000", "#,##0.0")
            f_(T, f"{c}{R['brk']}", f"=IFERROR(({c}{R['loss']}+{c}{R['inc']})/{c}{R['ttot']},\"n/m\")", "0.00")
            f_(T, f"{c}{R['chk']}", f"={c}{R['rev']}-({c}{R['core']}+{c}{R['aiRev']}+{c}{R['ps']})", "0.000")
            f_(T, f"{c}{R['share']}", f"=({c}{R['pcost']}+{c}{R['emb']})/PL!{c}{PB[s]['ngcogs']}", "0.0%")
    T.column_dimensions["A"].width = 70
    # ───────── PL ─────────
    P = W["PL"]
    f_(P, "A1", "PL — 損益（GAAP 為主；Non-GAAP 為公司自訂口徑，另列；US$M；EPS US$）。官方報告分部：單一應報導分部；營收線＝訂閱、專業服務與其他（segments.json）；地理（Americas／Europe／APAC）為揭露備註，不建模。", bold=True)
    f_(P, "A2", "FY2023A–FY2025A 為實績（Hist 頁）；FY2026E＝9M 實績＋Q4E；FY2027E 起為模型（三情境 A／B／C 並列；成本率持平於近四季）。會計年度＝曆年。最下方為標準損益表（情境 B）。")
    cf = W["CF"]
    f_(cf, "A1", "CF — 現金流與 DCF（FCFF＝NOPAT＋折舊攤銷 − 資本支出＋營運資金流入；股票薪酬不加回，加回只作敏感度）。US$M；DCF 折現到評價日 2026-09-30（期中慣例），FY2026 欄為 Q4 2026 單季（第一期）。", bold=True)
    f_(cf, "A2", "FY2026 欄的自由現金流為全年（9M 實績＋Q4E）；DCF 現金流只用 Q4E。收支關卡＝FY2026–FY2030 自由現金流減股利（HubSpot 無股利）。")
    f_(cf, "A3", "收支關卡：FY2026–FY2030 自由現金流減股利合計（US$M；A／B／C 三情境；正值＝通過）")
    for s, c in zip(SC, "BCD"):
        f_(cf, f"{c}3", f"=C{CB[s]['gate']}", "#,##0")
    f_(cf, "A4", "預測期最低年末現金與投資（US$M；A／B／C）")
    for s, c in zip(SC, "BCD"):
        f_(cf, f"{c}4", f"=C{CB[s]['mincash']}", "#,##0")
    for s in SC:
        R, B_, C_ = PB[s], TB[s], CB[s]
        f_(P, f"A{R['title']}", f"{SCN_NAME[s]}：損益表（GAAP；US$M；EPS 為 US$）", bold=True)
        pl_labels = {"yr": "年度", "sub": "訂閱", "ps": "專業服務與其他", "rev": "營收", "g": "營收年增率", "chk": "檢查：訂閱＋專業服務 − 營收合計",
                     "cust": "期末客戶數（命題子列，戶）", "asrpc": "每客戶訂閱營收（命題子列，US$；歷史為公司 ASRPC）", "chk2": "檢查：命題子列（平均客戶×每客戶）− 訂閱營收",
                     "cogs": "營業成本（GAAP）", "gp": "毛利（GAAP）", "gm": "毛利率（GAAP）", "opx": "營業費用（GAAP，含重組）", "ebit": "營業利益（GAAP）",
                     "ebitm": "營業利益率（GAAP）", "oth": "其他收入（淨利息與其他）", "pt": "稅前利益", "tax": "所得稅", "ni": "稅後淨利（GAAP）", "nim": "淨利率（GAAP）",
                     "eps": "EPS（GAAP 稀釋，US$）", "nig": "淨利年增率", "sh": "稀釋股數（百萬股；年平均）", "she": "期末稀釋股數（百萬股；EPS 口徑）",
                     "ngcogs": "非 GAAP 營業成本（含增量 AI 推論）", "ngopx": "非 GAAP 營業費用", "ngop": "非 GAAP 營業利益（公司口徑）", "ngopm": "非 GAAP 營業利益率",
                     "sbc": "股票薪酬", "adj": "非 GAAP 調節其他項（收購攤銷、收購費用、SBC 雇主稅等）", "restr": "重組費用", "ngeps": "非 GAAP 稀釋 EPS（US$）",
                     "inf": "增量 AI 推論成本（已含於非 GAAP 營業成本）"}
        for k, lab in pl_labels.items():
            f_(P, f"A{R[k]}", lab)
        hist_col = {2023: "C", 2024: "D", 2025: "E"}
        for y in YEARS:
            c = YC[y]
            pc = YC[y - 1] if y > 2023 else None
            f_(P, f"{c}{R['yr']}", f"FY{y}{'A' if y <= 2025 else 'E'}", bold=True)
            hs = {2023: "23", 2024: "24", 2025: "25"}.get(y)
            if y <= 2025:
                H_ = lambda k: f"h_{k}_{hs}"
                f_(P, f"{c}{R['sub']}", f"={H_('sub')}", "#,##0.0")
                f_(P, f"{c}{R['ps']}", f"={H_('ps')}", "#,##0.0")
                f_(P, f"{c}{R['cogs']}", f"={H_('cogs')}", "#,##0.0")
                f_(P, f"{c}{R['ebit']}", f"={H_('oi')}", "#,##0.0")
                f_(P, f"{c}{R['opx']}", f"={c}{R['gp']}-{c}{R['ebit']}", "#,##0.0")
                f_(P, f"{c}{R['pt']}", f"={H_('pt')}", "#,##0.0")
                f_(P, f"{c}{R['oth']}", f"={c}{R['pt']}-{c}{R['ebit']}", "#,##0.0")
                f_(P, f"{c}{R['tax']}", f"={H_('tx')}", "#,##0.0")
                f_(P, f"{c}{R['ni']}", f"={H_('ni')}", "#,##0.0")
                f_(P, f"{c}{R['sh']}", f"={H_('dsh')}", "#,##0.0")
                f_(P, f"{c}{R['eps']}", f"={H_('eps')}", "0.00")
                f_(P, f"{c}{R['sbc']}", f"={H_('sbc')}", "#,##0.0")
                f_(P, f"{c}{R['ngop']}", f"={H_('ngoi')}", "#,##0.0")
                f_(P, f"{c}{R['restr']}", f"={H_('restr')}", "#,##0.0")
                f_(P, f"{c}{R['cust']}", f"={H_('cust')}", "#,##0")
                f_(P, f"{c}{R['asrpc']}", f"={H_('asrpc')}", "#,##0")
                f_(P, f"{c}{R['chk2']}", "n/a")
                for k in ("ngcogs", "ngopx", "adj", "inf"):
                    f_(P, f"{c}{R[k]}", "未找到")
                if y == 2025:
                    f_(P, f"{c}{R['ngeps']}", f"={H_('ngeps')}", "0.00")
                else:
                    f_(P, f"{c}{R['ngeps']}", "未找到")
                f_(P, f"{c}{R['she']}", "未找到")
                f_(P, f"{c}{R['gp']}", f"={c}{R['rev']}-{c}{R['cogs']}", "#,##0.0")
            elif y == 2026:
                f_(P, f"{c}{R['sub']}", f"=Thesis!{c}{B_['sub']}", "#,##0.0")
                f_(P, f"{c}{R['ps']}", f"=Thesis!{c}{B_['ps']}", "#,##0.0")
                f_(P, f"{c}{R['cust']}", f"=Thesis!{c}{B_['cust']}", "#,##0")
                f_(P, f"{c}{R['asrpc']}", f"=Thesis!{c}{B_['arpc']}", "#,##0")
                f_(P, f"{c}{R['chk2']}", f"=Thesis!{c}{B_['avg']}*Thesis!{c}{B_['arpc']}/1000000+Thesis!{c}{B_['aiRev']}-{c}{R['sub']}", "0.000")
                f_(P, f"{c}{R['cogs']}", f"=h_cogs_9m26+Q4E!B{Q4['cogs']}", "#,##0.0")
                f_(P, f"{c}{R['ebit']}", f"=h_oi_9m26+Q4E!B{Q4['ebit']}", "#,##0.0")
                f_(P, f"{c}{R['opx']}", f"={c}{R['gp']}-{c}{R['ebit']}", "#,##0.0")
                f_(P, f"{c}{R['gp']}", f"={c}{R['rev']}-{c}{R['cogs']}", "#,##0.0")
                f_(P, f"{c}{R['oth']}", f"=(h_pt_9m26-h_oi_9m26)+Q4E!B{Q4['oth']}", "#,##0.0")
                f_(P, f"{c}{R['pt']}", f"={c}{R['ebit']}+{c}{R['oth']}", "#,##0.0")
                f_(P, f"{c}{R['tax']}", f"=h_tx_9m26+Q4E!B{Q4['tax']}", "#,##0.0")
                f_(P, f"{c}{R['ni']}", f"={c}{R['pt']}-{c}{R['tax']}", "#,##0.0")
                f_(P, f"{c}{R['she']}", "=sh_eps0", "#,##0.0")
                f_(P, f"{c}{R['sh']}", "=AVERAGE(h_dsh_q126,h_dsh_q226,h_dsh_q326,sh_eps0)", "#,##0.0")
                f_(P, f"{c}{R['eps']}", f"=h_eps_9m26+Q4E!B{Q4['eps']}", "0.00")
                f_(P, f"{c}{R['ngop']}", f"=h_ngoi_9m26+Q4E!B{Q4['ngop']}", "#,##0.0")
                f_(P, f"{c}{R['ngcogs']}", f"=h_ngcogs_9m26+Q4E!B{Q4['ngcogs']}", "#,##0.0")
                f_(P, f"{c}{R['ngopx']}", f"=h_ngopx_9m26+Q4E!B{Q4['ngopx']}", "#,##0.0")
                f_(P, f"{c}{R['sbc']}", f"=h_sbc_9m26+Q4E!B{Q4['sbc']}", "#,##0.0")
                f_(P, f"{c}{R['restr']}", f"=Q4E!B{Q4['restr']}", "#,##0.0")
                f_(P, f"{c}{R['adj']}", f"={c}{R['ngop']}-{c}{R['ebit']}-{c}{R['sbc']}-{c}{R['restr']}", "#,##0.0")
                f_(P, f"{c}{R['ngeps']}", f"=h_ngeps_9m26+Q4E!B{Q4['ngeps']}", "0.00")
                f_(P, f"{c}{R['inf']}", 0, "#,##0.0")
            else:
                n_ = f"Thesis!{c}{B_['n']}"
                f_(P, f"{c}{R['sub']}", f"=Thesis!{c}{B_['sub']}", "#,##0.0")
                f_(P, f"{c}{R['ps']}", f"=Thesis!{c}{B_['ps']}", "#,##0.0")
                f_(P, f"{c}{R['cust']}", f"=Thesis!{c}{B_['cust']}", "#,##0")
                f_(P, f"{c}{R['asrpc']}", f"=Thesis!{c}{B_['arpc']}", "#,##0")
                f_(P, f"{c}{R['chk2']}", f"=Thesis!{c}{B_['avg']}*Thesis!{c}{B_['arpc']}/1000000+Thesis!{c}{B_['aiRev']}-{c}{R['sub']}", "0.000")
                f_(P, f"{c}{R['inf']}", f"=Thesis!{c}{B_['inc']}", "#,##0.0")
                f_(P, f"{c}{R['ngcogs']}", f"=ng_cogs*{c}{R['rev']}+{c}{R['inf']}", "#,##0.0")
                f_(P, f"{c}{R['ngop']}", f"=h_ngoi_ltm+inc_m*({c}{R['rev']}-h_rev_ltm)-{c}{R['inf']}", "#,##0.0")
                f_(P, f"{c}{R['ngopx']}", f"={c}{R['rev']}-{c}{R['ngcogs']}-{c}{R['ngop']}", "#,##0.0")
                f_(P, f"{c}{R['sbc']}", f"=h_sbc_ltm+sbc_inc*({c}{R['rev']}-h_rev_ltm)", "#,##0.0")
                f_(P, f"{c}{R['adj']}", f"=adj_pct*{c}{R['rev']}", "#,##0.0")
                f_(P, f"{c}{R['restr']}", 0, "#,##0.0")
                f_(P, f"{c}{R['ebit']}", f"={c}{R['ngop']}-{c}{R['sbc']}-{c}{R['adj']}", "#,##0.0")
                f_(P, f"{c}{R['cogs']}", f"={c}{R['ngcogs']}+cogs_adj_sh*({c}{R['sbc']}+{c}{R['adj']})", "#,##0.0")
                f_(P, f"{c}{R['gp']}", f"={c}{R['rev']}-{c}{R['cogs']}", "#,##0.0")
                f_(P, f"{c}{R['opx']}", f"={c}{R['gp']}-{c}{R['ebit']}", "#,##0.0")
                f_(P, f"{c}{R['oth']}", f"=int_y*CF!{pc}{C_['cash']}", "#,##0.0")
                f_(P, f"{c}{R['pt']}", f"={c}{R['ebit']}+{c}{R['oth']}", "#,##0.0")
                f_(P, f"{c}{R['tax']}", f"=MAX(0,{c}{R['pt']})*tax", "#,##0.0")
                f_(P, f"{c}{R['ni']}", f"={c}{R['pt']}-{c}{R['tax']}", "#,##0.0")
                f_(P, f"{c}{R['she']}", f"={pc}{R['she']}+({c}{R['sbc']}-CF!{c}{C_['bb']})/(px*(1+ke)^({n_}-0.5))", "#,##0.00")
                f_(P, f"{c}{R['sh']}", f"=AVERAGE({pc}{R['she']},{c}{R['she']})", "#,##0.00")
                f_(P, f"{c}{R['eps']}", f"={c}{R['ni']}/{c}{R['sh']}", "0.00")
                f_(P, f"{c}{R['ngeps']}", f"=({c}{R['ngop']}+{c}{R['oth']})*(1-tax)/{c}{R['sh']}", "0.00")
            f_(P, f"{c}{R['rev']}", f"={c}{R['sub']}+{c}{R['ps']}", "#,##0.0")
            f_(P, f"{c}{R['chk']}", f"={c}{R['sub']}+{c}{R['ps']}-" + (f"h_rev_{hs}" if hs else f"{c}{R['rev']}"), "0.000")
            if y > 2023:
                f_(P, f"{c}{R['g']}", f"={c}{R['rev']}/{pc}{R['rev']}-1", "0.0%")
                f_(P, f"{c}{R['nig']}", f'=IF({pc}{R["ni"]}<=0,"n/m",{c}{R["ni"]}/{pc}{R["ni"]}-1)', "0.0%")
            else:
                f_(P, f"{c}{R['g']}", "n/m")
                f_(P, f"{c}{R['nig']}", "n/m")
            f_(P, f"{c}{R['gm']}", f"={c}{R['gp']}/{c}{R['rev']}", "0.0%")
            f_(P, f"{c}{R['ebitm']}", f"={c}{R['ebit']}/{c}{R['rev']}", "0.0%")
            f_(P, f"{c}{R['nim']}", f"={c}{R['ni']}/{c}{R['rev']}", "0.0%")
            f_(P, f"{c}{R['ngopm']}", f"={c}{R['ngop']}/{c}{R['rev']}", "0.0%")
        # 修正：FY2023 的 g（無 FY2022 數）已為 n/m
        # ── CF 區塊 ──
        f_(cf, f"A{C_['title']}", f"{SCN_NAME[s]}：現金流與 DCF（US$M）", bold=True)
        cl = {"yr": "年度", "rev": "營收", "gr": "營收年增率", "ebit": "營業利益（GAAP）", "tax": "稅率", "nopat": "NOPAT＝EBIT ×（1 − 稅率）（EBIT 為負時不課稅）", "da": "＋折舊與攤銷",
              "cx": "− 資本支出（含資本化軟體）", "wc": "＋營運資金流入（遞延營收等）", "sbcab": "＋股票薪酬加回（預設 0）", "fcff": "＝無槓桿自由現金流（FCFF；SBC 費用化）",
              "int": "稅後其他收入（現金收益）", "fcf": "自由現金流（現金口徑近似）＝FCFF＋非現金 SBC＋稅後其他收入", "bb": "回購", "cash": "年末現金與投資",
              "t": "折現時點（自評價日，年；期中）", "cf": "DCF 現金流（FY2026 欄＝Q4 2026 單季）", "df": "折現因子", "pv": "現值",
              "spv": "明確期現值合計", "n33": "FY2032 NOPAT × (1＋g)（即 FY2033 NOPAT）", "tv": "終值", "tvdf": "終值折現因子", "tvpv": "終值現值", "ev": "企業價值 EV",
              "nc": "＋淨現金（現金與投資 − 債務）", "eq": "股權價值", "shv": "評價股數（百萬股；含 RSU 扣預期沒收與選擇權）", "ps0": "每股價值（評價日，US$）",
              "ps1": "每股價值（目標時點，US$）", "tvsh": "終值占 EV", "gate": "收支關卡：FY2026–FY2030 自由現金流減股利合計（US$M）", "mincash": "預測期最低年末現金與投資（US$M）",
              "sbcpv": "對照：SBC 稅後現值（含終值；每股，US$，目標時點）"}
        for k, lab in cl.items():
            f_(cf, f"A{C_[k]}", lab)
        for y in range(2026, 2033):
            c = YC[y]
            pc = YC[y - 1] if y > 2026 else None
            f_(cf, f"{c}{C_['yr']}", f"FY{y}{'（Q4 單季）' if y == 2026 else 'E'}", bold=True)
            f_(cf, f"{c}{C_['rev']}", f"=PL!{c}{R['rev']}", "#,##0.0")
            f_(cf, f"{c}{C_['gr']}", f"=PL!{c}{R['g']}", "0.0%")
            f_(cf, f"{c}{C_['tax']}", "=tax", "0.0%")
            if y == 2026:
                for k, qk in (("ebit", "ebit"), ("nopat", "nopat"), ("da", "da"), ("cx", "cx"), ("wc", "nwc"), ("fcff", "fcff")):
                    f_(cf, f"{c}{C_[k]}", f"=Q4E!B{Q4[qk]}", "#,##0.0")
                f_(cf, f"{c}{C_['sbcab']}", f"=sbc_ab*Q4E!B{Q4['sbc']}", "#,##0.0")
                f_(cf, f"{c}{C_['int']}", f"=Q4E!B{Q4['oth']}*(1-IF(Q4E!B{Q4['pt']}>0,tax,0))", "#,##0.0")
                f_(cf, f"{c}{C_['fcf']}", f"=(h_cfo_9m26-h_ppe_9m26-h_capsw_9m26)+Q4E!B{Q4['fcf']}", "#,##0.0")
                f_(cf, f"{c}{C_['bb']}", "=bb_q4", "#,##0.0")
                f_(cf, f"{c}{C_['cash']}", f"=Q4E!B{Q4['cash']}", "#,##0.0")
                f_(cf, f"{c}{C_['t']}", "=((DATE(2026,12,31)-val_date)/365)/2", "0.000")
                f_(cf, f"{c}{C_['cf']}", f"={c}{C_['fcff']}", "#,##0.0")
            else:
                f_(cf, f"{c}{C_['ebit']}", f"=PL!{c}{R['ebit']}", "#,##0.0")
                f_(cf, f"{c}{C_['nopat']}", f"=IF({c}{C_['ebit']}>0,{c}{C_['ebit']}*(1-{c}{C_['tax']}),{c}{C_['ebit']})", "#,##0.0")
                f_(cf, f"{c}{C_['da']}", f"=da_pct*{c}{C_['rev']}", "#,##0.0")
                f_(cf, f"{c}{C_['cx']}", f"=cx_pct*{c}{C_['rev']}", "#,##0.0")
                f_(cf, f"{c}{C_['wc']}", f"=nwc_k*({c}{C_['rev']}-{pc}{C_['rev']})", "#,##0.0")
                f_(cf, f"{c}{C_['sbcab']}", f"=sbc_ab*PL!{c}{R['sbc']}", "#,##0.0")
                f_(cf, f"{c}{C_['fcff']}", f"={c}{C_['nopat']}+{c}{C_['da']}-{c}{C_['cx']}+{c}{C_['wc']}+{c}{C_['sbcab']}", "#,##0.0")
                f_(cf, f"{c}{C_['int']}", f"=PL!{c}{R['oth']}*(1-IF(PL!{c}{R['pt']}>0,tax,0))", "#,##0.0")
                f_(cf, f"{c}{C_['fcf']}", f"={c}{C_['fcff']}+(1-sbc_ab)*PL!{c}{R['sbc']}+{c}{C_['int']}", "#,##0.0")
                f_(cf, f"{c}{C_['bb']}", f"=bb_pct*{c}{C_['fcf']}", "#,##0.0")
                f_(cf, f"{c}{C_['cash']}", f"={pc}{C_['cash']}+{c}{C_['fcf']}-{c}{C_['bb']}", "#,##0.0")
                f_(cf, f"{c}{C_['t']}", f"=(DATE(2026,12,31)-val_date)/365+mid+{y - 2027}", "0.000")
                f_(cf, f"{c}{C_['cf']}", f"={c}{C_['fcff']}", "#,##0.0")
            f_(cf, f"{c}{C_['df']}", f"=1/(1+wacc)^{c}{C_['t']}", "0.0000")
            f_(cf, f"{c}{C_['pv']}", f"={c}{C_['cf']}*{c}{C_['df']}", "#,##0.0")
        X = lambda k: f"C{C_[k]}"
        f_(cf, X("spv"), f"=SUM(F{C_['pv']}:L{C_['pv']})", "#,##0.0")
        f_(cf, X("n33"), f"=L{C_['nopat']}*(1+g_tv)", "#,##0.0")
        f_(cf, X("tv"), f"=IF(tv_alt=1,L{C_['fcff']}*(1+g_tv)/(wacc-g_tv),{X('n33')}*(1-g_tv/ronic)/(wacc-g_tv))", "#,##0.0")
        f_(cf, X("tvdf"), "=1/(1+wacc)^((DATE(2026,12,31)-val_date)/365+6)", "0.0000")
        f_(cf, X("tvpv"), f"={X('tv')}*{X('tvdf')}", "#,##0.0")
        f_(cf, X("ev"), f"={X('spv')}+{X('tvpv')}", "#,##0.0")
        f_(cf, X("nc"), "=cash_0-debt_0", "#,##0.0")
        f_(cf, X("eq"), f"={X('ev')}+{X('nc')}", "#,##0.0")
        f_(cf, X("shv"), "=sh_val", "0.000")
        f_(cf, X("ps0"), f"={X('eq')}/{X('shv')}", "0.00")
        f_(cf, X("ps1"), f"={X('ps0')}*(1+ke)^t_tgt", "0.00")
        f_(cf, X("tvsh"), f"={X('tvpv')}/{X('ev')}", "0.0%")
        f_(cf, X("gate"), f"=SUM(F{C_['fcf']}:J{C_['fcf']})", "#,##0.0")
        f_(cf, X("mincash"), f"=MIN(F{C_['cash']}:L{C_['cash']})", "#,##0.0")
        f_(cf, X("sbcpv"), f"=((SUMPRODUCT(PL!G{R['sbc']}:L{R['sbc']},G{C_['df']}:L{C_['df']})+Q4E!B{Q4['sbc']}*F{C_['df']})*(1-tax)"
                           f"+PL!L{R['sbc']}*(1+g_tv)*(1-tax)*(1-g_tv/ronic)/(wacc-g_tv)*{X('tvdf')})/{X('shv')}*(1+ke)^t_tgt", "0.00")
    cf.column_dimensions["A"].width = 62
    P.column_dimensions["A"].width = 58
    # 標準損益表（情境 B；HTML 用）
    S0 = SIS
    f_(P, f"A{S0['title']}", "標準損益表（情境 B 基準；GAAP；US$M；EPS US$；FY2023A–FY2025A 實績、FY2026E 起預測）。GAAP 為主，Non-GAAP 為公司自訂口徑（排除股票薪酬、收購攤銷、重組等），列於最後兩列。", bold=True)
    rb = PB["B"]
    sis_map = [("rev", "營收", "rev", "#,##0.0"), ("g", "營收年增率", "g", "0.0%"), ("cogs", "營業成本", "cogs", "#,##0.0"), ("gp", "毛利", "gp", "#,##0.0"),
               ("gm", "毛利率", "gm", "0.0%"), ("opx", "營業費用", "opx", "#,##0.0"), ("ebit", "營業利益", "ebit", "#,##0.0"), ("ebitm", "營業利益率", "ebitm", "0.0%"),
               ("oth", "其他收入", "oth", "#,##0.0"), ("pt", "稅前利益", "pt", "#,##0.0"), ("tax", "所得稅", "tax", "#,##0.0"), ("ni", "淨利", "ni", "#,##0.0"),
               ("nim", "淨利率", "nim", "0.0%"), ("eps", "EPS（GAAP 稀釋，US$）", "eps", "0.00"), ("nig", "淨利年增率", "nig", "0.0%"),
               ("ngeps", "非 GAAP EPS（公司口徑，US$）", "ngeps", "0.00"), ("ngop", "非 GAAP 營業利益（公司口徑）", "ngop", "#,##0.0"), ("ngopm", "非 GAAP 營業利益率", "ngopm", "0.0%")]
    f_(P, f"A{S0['yr']}", "年度")
    for y in YEARS:
        c = YC[y]
        f_(P, f"{c}{S0['yr']}", f"=PL!{c}{rb['yr']}", bold=True)
    for k, lab, src, nf in sis_map:
        f_(P, f"A{S0[k]}", lab)
        for y in YEARS:
            c = YC[y]
            f_(P, f"{c}{S0[k]}", f"={c}{rb[src]}", nf)
    # ───────── Val ─────────
    V = W["Val"]
    VKEYS = ["tpdcf", "tpcomp", "tp", "px", "cons", "up", "upc", "updcf", "tier", "rec_val", "gatev", "rec", "per", "gate_all", "tvsh2", "bdist", "uattr", "dep", "perel",
             "gap", "gap_sbc", "gap_time", "gap_lev", "gap_mult", "pe_hubs_now", "san_dcf", "san_ev", "san_rec", "san_g", "san_m27", "san_diag"]
    VR = {k: 6 + i for i, k in enumerate(VKEYS)}
    vw0 = VR["san_diag"] + 3
    vw = {k: vw0 + 1 + i for i, k in enumerate(["rf", "beta", "erp", "ke", "kd", "tax", "mc", "wd", "wacc", "g", "ronic", "sh_basic", "rsu", "opt", "sh_val"])}
    VR["wacc_hdr"] = vw0
    d0 = vw["sh_val"] + 3
    VR["dcf_hdr"] = d0
    dsum_keys = ["spv", "tvpv", "ev", "nc", "eq", "shv", "ps0", "ps1", "tvsh"]
    VR["dsum0"] = d0 + 2
    c0 = VR["dsum0"] + len(dsum_keys) + 2
    VR["cmp_hdr"] = c0
    cmp_keys = ["eps27", "eps28", "w27", "ntm", "pe", "cmpps"]
    VR["cmp0"] = c0 + 2
    p0 = VR["cmp0"] + len(cmp_keys) + 2
    VR["peer_hdr"] = p0
    VR["peer0"] = p0 + 2
    f_(V, "A1", tf("Val — 目標價（股價基準日＋12 個月＝", ("tp_date", "yyyy-mm-dd"), "）＝DCF ", ("w_dcf", "0%"), "＋Comparable ", ("1-w_dcf", "0%"), "（情境 A／B／C 依權重加權）"))
    f_(V, "A2", "DCF：FCFF（NOPAT＋折舊攤銷 − 資本支出＋營運資金流入；股票薪酬不加回）以 WACC 折現到評價日（期中慣例），終值＝NOPAT(FY2033) ×(1 − g ÷ RONIC)÷(WACC − g)，再以 ke 推到目標時點。")
    f_(V, "A3", "Comparable：同業前瞻 P/E 中位數（股權倍數、有槓桿）× 相對倍數係數 × HUBS 目標時點 NTM 非 GAAP 每股盈餘（FY2027 與 FY2028 依日期加權）；單一分部，不做分部加總；同業 11 家，分母口徑未核，不做迴歸。")
    f_(V, "A4", "結論", bold=True)
    for j, h in enumerate(["項目", "數值", "對照", "對照", "說明"], 1):
        f_(V, f"{L_(j)}5", h, bold=True, fill=HDR)

    def vr(k, label, fb=None, nf=None, fc=None, fd=None, fe=None, nfc=None, nfd=None):
        r_ = VR[k]
        f_(V, f"A{r_}", label)
        if fb is not None:
            f_(V, f"B{r_}", fb, nf)
        if fc is not None:
            f_(V, f"C{r_}", fc, nfc or nf)
        if fd is not None:
            f_(V, f"D{r_}", fd, nfd or nf)
        if fe is not None:
            f_(V, f"E{r_}", fe)
    cmpr = {k: VR["cmp0"] + i for i, k in enumerate(cmp_keys)}
    vr("tpdcf", "DCF 目標價（情境加權）", "=" + "+".join(f"tp_w{s}*CF!C{CB[s]['ps1']}" for s in SC), "0.00", fe="情境權重 25／50／25")
    vr("tpcomp", "Comparable 目標價（同業前瞻 P/E 中位數法；情境加權）", "=" + "+".join(f"tp_w{s}*{c}{cmpr['cmpps']}" for s, c in zip(SC, "BCD")), "0.00")
    vr("tp", "目標價（DCF 與 Comparable 加權）", f"=w_dcf*B{VR['tpdcf']}+(1-w_dcf)*B{VR['tpcomp']}", "0.00", fe=tf("DCF 權重 ", ("w_dcf", "0%"), "、Comparable ", ("1-w_dcf", "0%")))
    vr("px", "現價", "=px", "0.00", fe=tf("股價基準日 ", ("px_date", "yyyy-mm-dd"), "（QuiverQuant 與 ChartRow 兩處同值）"))
    vr("cons", "賣方平均目標價（只作對照）", "=c_tp", "0.00", fe="TickFlow 2026-11-16，49 位；只作對照")
    vr("up", "上檔空間：加權目標價", f"=B{VR['tp']}/px-1", "+0.0%;-0.0%")
    vr("upc", "上檔空間：Comparable 目標價", f"=B{VR['tpcomp']}/px-1", "+0.0%;-0.0%")
    vr("updcf", "上檔空間：DCF 目標價", f"=B{VR['tpdcf']}/px-1", "+0.0%;-0.0%")
    vr("tier", "上檔空間級距", "=" + tier_f(f"B{VR['up']}"))
    vr("rec_val", "估值評等（依加權目標價上檔空間；五級）", "=" + tier_f(f"B{VR['up']}"), fe="五級固定規則（原則第 24 條）：> +25% 買進、> +50% 強力買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空")
    vr("gatev", "入不敷出判定用：收支關卡（情境 B；FY2026–FY2030 自由現金流減股利合計，US$M）", f"=CF!C{CB['B']['gate']}", "#,##0", fe="＝CF 頁收支關卡（情境 B）；負值＝入不敷出")
    vr("rec", "投資建議", f'=IF(AND(B{VR["gatev"]}<0,OR(B{VR["rec_val"]}="買進",B{VR["rec_val"]}="強力買進")),"收支不足",B{VR["rec_val"]})', fe="收支關卡未過時不得為買進（原則第 63 條）")
    vr("per", "期間：折現期間（評價日起）／持有期間（股價基準日起）（年）", "=t_tgt", "0.00", "=t_hold", fe=tf("評價日 ", ("val_date", "yyyy-mm-dd"), " → 目標時點 ", ("tp_date", "yyyy-mm-dd"), "；上檔空間不是年化報酬"))
    vr("gate_all", "收支關卡：FY2026–FY2030 自由現金流減股利合計（US$M；情境 A／B／C）", f"=CF!C{CB['A']['gate']}", "#,##0", f"=CF!C{CB['B']['gate']}", f"=CF!C{CB['C']['gate']}")
    vr("tvsh2", "DCF 終值占 EV（情境 A／B／C；超過約 80% 標示）", f"=CF!C{CB['A']['tvsh']}", "0.0%", f"=CF!C{CB['B']['tvsh']}", f"=CF!C{CB['C']['tvsh']}",
       fe=f'=IF(MAX(B{VR["tvsh2"]}:D{VR["tvsh2"]})>0.8,"超過 80%：估值對永續成長與 WACC 高度敏感","未超過 80%")')
    vr("bdist", "距最近評等界線（百分點；上檔空間 − 界線）／最近界線", f"=MIN(ABS(B{VR['up']}-0.5),ABS(B{VR['up']}-0.25),ABS(B{VR['up']}+0.25),ABS(B{VR['up']}+0.5))*100", "0.0",
       f'=IF(ABS(B{VR["up"]}-0.5)=B{VR["bdist"]}/100,"+50%",IF(ABS(B{VR["up"]}-0.25)=B{VR["bdist"]}/100,"+25%",IF(ABS(B{VR["up"]}+0.25)=B{VR["bdist"]}/100,"−25%","−50%")))',
       fe=f'=IF(B{VR["bdist"]}<5,"距評等界線不足 5 個百分點：評等對輸入微小變動敏感，翻轉條件見 Recon 與 Sens","距評等界線 5 個百分點以上")')
    vr("uattr", "上檔空間三段歸因：DCF 貢獻／未調整 Comparable 貢獻（係數＝1）／相對倍數調整", f"=w_dcf*B{VR['updcf']}", "+0.0%;-0.0%", f"=(1-w_dcf)*((B{VR['tpcomp']}/rel_mult)/px-1)",
       f"=B{VR['up']}-B{VR['uattr']}-C{VR['uattr']}", fe="相對倍數係數基準＝1.00（未調整：缺 10 年前瞻倍數），故第三項為 0；Sens 有係數低／高案")
    vr("dep", "加權目標價中 DCF／Comparable 各占的美元", f"=w_dcf*B{VR['tpdcf']}", "0.00", f"=(1-w_dcf)*B{VR['tpcomp']}")
    vr("perel", "現行相對倍數：HUBS 前瞻 P/E ÷ 同業中位數／隱含重估（係數 ÷ 現行 − 1）", "=c_pe_hubs/pe_med", "0.00", "=rel_mult/B" + str(VR["perel"]) + "-1", nfc="+0.0%;-0.0%",
       fe=tf("HUBS 現行前瞻 P/E ", ("c_pe_hubs", "0.0"), " 倍對同業中位數 ", ("pe_med", "0.0"), " 倍；相對倍數係數基準 ", ("rel_mult", "0.00"), "（未調整：缺 10 年前瞻倍數）"))
    vr("gap", "DCF 與 Comparable 落差（每股，US$；Comparable − DCF，情境加權）", f"=B{VR['tpcomp']}-B{VR['tpdcf']}", "+0.00;-0.00", fe="以下四列公式分解（工作單 G2.8）")
    sbcw = "(" + "+".join(f"tp_w{s}*CF!C{CB[s]['sbcpv']}" for s in SC) + ")"
    vr("gap_sbc", "落差分解 1：股票薪酬（DCF 費用化、Comparable 的非 GAAP EPS 不扣；SBC 稅後現值每股）", f"={sbcw}", "+0.00;-0.00")
    vr("gap_time", "落差分解 2：時點（DCF 以 ke 推到目標時點；Comparable 直接用目標時點 NTM EPS）", "=0", "+0.00;-0.00", fe="兩法同為目標時點價，無時點落差；DCF 評價日價格另列於彙總表")
    vr("gap_lev", "落差分解 3：槓桿（淨現金在 DCF 另加；在股權倍數內含）", "=0", "+0.00;-0.00", fe="HUBS 無債務；淨現金每股（目標時點）見彙總表；P/E 為股權倍數，已含淨現金，無槓桿落差")
    vr("gap_mult", "落差分解 4：倍數與成長（剩餘）＝落差 − 股票薪酬 − 時點 − 槓桿", f"=B{VR['gap']}-B{VR['gap_sbc']}-B{VR['gap_time']}-B{VR['gap_lev']}", "+0.00;-0.00",
       fe="剩餘項：同業 P/E 中位數對應的盈餘成長與風險，高於本模型以 WACC 折現的 FCFF 所隱含者")
    vr("pe_hubs_now", "現價隱含 P/E：現價 ÷ 本模型 NTM 非 GAAP EPS（目標時點，情境 B）", f"=px/C{cmpr['ntm']}", "0.0", fe="對照同業中位數；本模型 EPS 以 48.9M 庫藏股法股數，DCF 以含全數 RSU 的評價股數")
    vr("san_dcf", "合理性關卡 1（G3.3）：DCF 目標價 ÷ 現價（<25% 或 >4 倍須診斷）", f"=B{VR['tpdcf']}/px", "0.00", f'=IF(OR(B{VR["san_dcf"]}<0.25,B{VR["san_dcf"]}>4),"觸發","通過")')
    vr("san_ev", "合理性關卡 2：企業價值（情境 B，US$M；為負須診斷）", f"=CF!C{CB['B']['ev']}", "#,##0", f'=IF(B{VR["san_ev"]}<=0,"觸發","通過")')
    vr("san_rec", "合理性關卡 3：估值評等（強力買進或強力放空須診斷）", f"=B{VR['rec_val']}", "@", f'=IF(OR(B{VR["san_rec"]}="強力放空",B{VR["san_rec"]}="強力買進"),"觸發","通過")')
    vr("san_g", "合理性關卡 4：FY2026 營收與非 GAAP 營業利益率對公司指引的最大差距（>10% 須診斷）", f"=MAX(ABS(PL!F{PB['B']['rev']}/g_rev26-1),ABS(PL!F{PB['B']['ngopm']}/(g_ngoi26/g_rev26)-1))", "0.0%", f'=IF(B{VR["san_g"]}>0.1,"觸發","通過")')
    vr("san_m27", "對照：FY2027 非 GAAP 營業利益率（模型）對管理層目標中點的差距（目標非指引，不入關卡）", f"=PL!G{PB['B']['ngopm']}/g_m27-1", "+0.0%;-0.0%")
    vr("san_diag", "診斷結論", f'=IF(COUNTIF(C{VR["san_dcf"]}:C{VR["san_rec"]},"觸發")+IF(C{VR["san_g"]}="觸發",1,0)=0,"四項關卡皆通過：DCF 目標價為現價的 "&TEXT(B{VR["san_dcf"]},"0%")&"，EV 為正，評等 "&B{VR["rec_val"]}&"，FY2026 對指引最大差距 "&TEXT(B{VR["san_g"]},"0.0%")&"；DCF 低於現價主因為股票薪酬費用化（SBC 稅後現值每股 US$"&TEXT(B{VR["gap_sbc"]},"0")&"）與 WACC 對高成長終值的折現，屬經濟判斷而非建模錯誤","有關卡觸發：須先診斷輸入或口徑再發布（見 README）")')
    # WACC 區
    f_(V, f"A{vw0}", "折現率與資本結構（個股 DCF 標準規格第 2 節）", bold=True)
    rows_w = [("rf", "無風險利率 rf", "=rf", "0.00%"), ("beta", "β（Blume 調整後兩來源平均）", "=beta", "0.000"), ("erp", "股權風險溢酬 ERP", "=erp", "0.00%"),
              ("ke", "股權成本 ke＝rf＋β × ERP", "=rf+beta*erp", "0.00%"), ("kd", "稅後債務成本（無債務，不計）", "=0", "0.00%"), ("tax", "稅率", "=tax", "0.0%"),
              ("mc", "股權市值（現價 × 評價股數，US$M）", f"=px*B{vw['sh_val']}", "#,##0"), ("wd", "債務權重 D÷(D＋E)", f"=debt_0/(debt_0+B{vw['mc']})", "0.00%"),
              ("wacc", "WACC（含敏感度調整）", f"=(1-B{vw['wd']})*B{vw['ke']}+B{vw['wd']}*B{vw['kd']}+wacc_adj", "0.00%"), ("g", "永續成長率 g", "=g_tv", "0.0%"),
              ("ronic", "永續期 RONIC＝WACC＋差", f"=B{vw['wacc']}+ronic_sp", "0.00%"),
              ("sh_basic", "評價股數：流通股（10-Q 封面，百萬股）", "=sh_basic", "0.000"),
              ("rsu", "評價股數：RSU 與 PSU（扣預期沒收）", "=rsu_dil*(1-forfeit)", "0.000"),
              ("opt", "評價股數：選擇權庫藏股法淨增股數", "=opt_dil", "0.000"),
              ("sh_val", "評價股數合計（百萬股）", f"=B{vw['sh_basic']}+B{vw['rsu']}+B{vw['opt']}", "0.000")]
    for k, lab, fm, nf in rows_w:
        f_(V, f"A{vw[k]}", lab)
        f_(V, f"B{vw[k]}", fm, nf)
    addname(wb, "ke", f"Val!$B${vw['ke']}")
    addname(wb, "wacc", f"Val!$B${vw['wacc']}")
    addname(wb, "ronic", f"Val!$B${vw['ronic']}")
    addname(wb, "sh_val", f"Val!$B${vw['sh_val']}")
    # DCF 彙總
    f_(V, f"A{d0}", "DCF 彙總（三情境）", bold=True)
    for j, h in enumerate(["項目", "情境 A", "情境 B", "情境 C"], 1):
        f_(V, f"{L_(j)}{d0 + 1}", h, bold=True, fill=HDR)
    dsum = {"spv": ("明確期現值合計（US$M）", "#,##0"), "tvpv": ("終值現值（US$M）", "#,##0"), "ev": ("企業價值 EV（US$M）", "#,##0"), "nc": ("淨現金（US$M）", "#,##0"),
            "eq": ("股權價值（US$M）", "#,##0"), "shv": ("評價股數（百萬股）", "0.000"), "ps0": ("每股價值：評價日（US$）", "0.00"), "ps1": ("每股價值：目標時點（US$）", "0.00"),
            "tvsh": ("終值占 EV", "0.0%")}
    for i, k in enumerate(dsum_keys):
        rr = VR["dsum0"] + i
        VR["dsum_" + k] = rr
        f_(V, f"A{rr}", dsum[k][0])
        for c, s in zip("BCD", SC):
            f_(V, f"{c}{rr}", f"=CF!C{CB[s][k]}", dsum[k][1])
    # Comparable 表
    f_(V, f"A{c0}", "Comparable（情境 A／B／C；同業前瞻 P/E 中位數 × 相對倍數係數 × 目標時點 NTM 非 GAAP EPS）", bold=True)
    for j, h in enumerate(["項目", "情境 A", "情境 B", "情境 C"], 1):
        f_(V, f"{L_(j)}{c0 + 1}", h, bold=True, fill=HDR)
    clab = {"eps27": ("FY2027E 非 GAAP EPS（US$）", "0.00"), "eps28": ("FY2028E 非 GAAP EPS（US$）", "0.00"),
            "w27": ("FY2027 權重＝(2027-12-31 − 目標時點) ÷ 365", "0.000"), "ntm": ("目標時點 NTM 非 GAAP EPS（US$）", "0.00"),
            "pe": ("採用 P/E（同業中位數 × 相對倍數係數；倍數分母＝前瞻一年每股盈餘）", "0.0"), "cmpps": ("每股價值（目標時點，US$）", "0.00")}
    for i, k in enumerate(cmp_keys):
        f_(V, f"A{cmpr[k]}", clab[k][0])
    for c, s in zip("BCD", SC):
        f_(V, f"{c}{cmpr['eps27']}", f"=PL!H{PB[s]['ngeps']}", "0.00")
        f_(V, f"{c}{cmpr['eps28']}", f"=PL!I{PB[s]['ngeps']}", "0.00")
        f_(V, f"{c}{cmpr['w27']}", "=(DATE(2027,12,31)-tp_date)/365", "0.000")
        f_(V, f"{c}{cmpr['ntm']}", f"={c}{cmpr['w27']}*{c}{cmpr['eps27']}+(1-{c}{cmpr['w27']})*{c}{cmpr['eps28']}", "0.00")
        f_(V, f"{c}{cmpr['pe']}", "=pe_med*rel_mult", "0.0")
        f_(V, f"{c}{cmpr['cmpps']}", f"={c}{cmpr['pe']}*{c}{cmpr['ntm']}", "0.00")
    # 同業表
    f_(V, f"A{p0}", "同業前瞻 P/E（ChartRow，2026-11-13；分母口徑與租賃處理未核；皆為美國 GAAP 申報人）", bold=True)
    for j, h in enumerate(["同業", "前瞻 P/E（倍）"], 1):
        f_(V, f"{L_(j)}{p0 + 1}", h, bold=True, fill=HDR)
    peers = [it for it in INP if it[0].startswith("pe_") and it[0] not in ("pe_med", "pe_date")]
    for i, it in enumerate(peers):
        f_(V, f"A{VR['peer0'] + i}", it[1].split("：")[1].split("（")[0])
        f_(V, f"B{VR['peer0'] + i}", f"={it[0]}", "0.0")
    VR["peer_med"] = VR["peer0"] + len(peers)
    VR["peer_n"] = len(peers)
    f_(V, f"A{VR['peer_med']}", "中位數")
    f_(V, f"B{VR['peer_med']}", "=pe_med", "0.0")
    f_(V, f"A{VR['peer_med'] + 1}", "HUBS 現行前瞻 P/E（ChartRow）")
    f_(V, f"B{VR['peer_med'] + 1}", "=c_pe_hubs", "0.0")
    VR["peer_hubs"] = VR["peer_med"] + 1
    V.column_dimensions["A"].width = 76
    V.column_dimensions["E"].width = 90
    for c in "BCD":
        V.column_dimensions[c].width = 14
    # ───────── RevDCF ─────────
    RV = W["RevDCF"]
    RVA = revdcf(wb, RV, VR, vw)
    # ───────── Recon ─────────
    RCA = recon(wb, W["Recon"], VR, kreq, keyrow)
    # ───────── README ─────────
    inhash = hashlib.sha256(json.dumps([[str(x) for x in it] for it in INP] + [sorted((k, str(v)) for k, v in ov.items())] + [VERSION, hashlib.sha256(open(os.path.abspath(__file__), "rb").read() + open(os.path.join(HERE, "hs_inputs.py"), "rb").read()).hexdigest()], ensure_ascii=False).encode()).hexdigest()[:10]
    build_id = f"hubspot-{VERSION}-{inhash}"
    SENS = sens_sheet(wb, W["Sens"], sens, build_id, inhash)
    readme(wb, W["README"], build_id, VR, RVA, RCA, vw, SENS)
    # ───────── Sources ─────────
    sources(W["Sources"])
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    wb.save(out)
    json.dump({"VR": VR, "vw": vw, "RVA": RVA, "RCA": RCA, "SENS": SENS, "PB": PB, "CB": CB, "TB": TB, "SIS": SIS, "Q4": Q4, "SCN": SCNROW, "SCN_HDR": SCN_HDR, "keyrow": keyrow,
               "build_id": build_id}, open(os.environ.get("ROWMAP", out + ".rowmap.json"), "w"), ensure_ascii=False)
    return build_id


# ═════════════════════════════════════════════════════════════════════════════
def revdcf(wb, RV, VR, vw):
    CBB = CB["B"]
    f_(RV, "A1", "RevDCF — 反向 DCF：現價隱含什麼（一次只動一個驅動、其餘維持基準，反解到「現價」；全部為 Excel 公式，封閉解＋二分法）。本頁只讀 In、CF、Val 頁，無任何頁引用本頁，不影響目標價", bold=True)
    f_(RV, "A2", "讀法：以情境 B 路徑反解「情境 B 的 DCF 每股價值推到目標時點＝現價」（與目標價同一時點口徑：評價日每股價值 ×（1＋股權成本）^年數）。營業利益率與營收規模兩項只動 FY2028 起的營業利益（FY2026 Q4 與 FY2027 為錨點；折舊、資本支出、營運資金不變）；現價低於模型值時，「現價隱含」是把模型往下調到現價所需的幅度。")
    A = {}
    r = 4
    f_(RV, f"A{r}", "A. DCF（情境 B 路徑）：支持現價所需的企業價值", bold=True)
    r += 1
    for j, h in enumerate(["項目", "值", "說明"], 1):
        f_(RV, f"{L_(j)}{r}", h, bold=True, fill=HDR)
    r += 1

    def kv(key, label, f, fmt, note=""):
        nonlocal r
        f_(RV, f"A{r}", label)
        f_(RV, f"B{r}", f, fmt)
        if note:
            f_(RV, f"C{r}", note)
        A[key] = r
        r += 1
    kv("px", "現價（美元）", "=px", "0.00")
    kv("v0", "支持現價所需的評價日每股價值（美元）", "=px/(1+ke)^t_tgt", "0.00", "＝現價 ÷（1＋股權成本）^目標時點年數")
    kv("sh", "評價股數（百萬股）", "=sh_val", "0.000")
    kv("oth", "淨現金（US$M）", "=cash_0-debt_0", "#,##0.0")
    kv("evr", "支持現價所需的企業價值 EV（US$M）", f"=B{A['v0']}*B{A['sh']}-B{A['oth']}", "#,##0.0")
    kv("ev", "模型企業價值 EV（情境 B）", f"=CF!C{CBB['ev']}", "#,##0.0")
    kv("evx", "現價所需 EV 為模型的倍數", f"=B{A['evr']}/B{A['ev']}", "0.00")
    kv("v_b", "模型：情境 B 評價日每股價值（美元）", f"=(B{A['ev']}+B{A['oth']})/B{A['sh']}", "0.00")
    r += 1
    yh = r
    for j, h in enumerate(["項目（情境 B）", "Q4 2026", "FY2027", "FY2028", "FY2029", "FY2030", "FY2031", "FY2032"], 1):
        f_(RV, f"{L_(j)}{r}", h, bold=True, fill=HDR)
    r += 1
    Y = {}
    cols = ["F", "G", "H", "I", "J", "K", "L"]

    def yr(key, label, fn, fmt):
        nonlocal r
        f_(RV, f"A{r}", label)
        for i, vc in enumerate(cols):
            f_(RV, f"{L_(2 + i)}{r}", fn(vc, L_(2 + i)), fmt)
        Y[key] = r
        r += 1
    yr("rev", "營收（US$M；Q4 為單季）", lambda vc, c: f"=CF!{vc}{CBB['rev']}" if vc != "F" else f"=Q4E!B{Q4['rev']}", "#,##0.0")
    yr("ebit", "營業利益 EBIT（US$M；財報口徑）", lambda vc, c: f"=CF!{vc}{CBB['ebit']}", "#,##0.0")
    yr("m", "營業利益率", lambda vc, c: f"={c}{Y['ebit']}/{c}{Y['rev']}", "0.0%")
    yr("df", "折現因子", lambda vc, c: f"=CF!{vc}{CBB['df']}", "0.0000")
    yr("tvf", "終值係數（只在 FY2032：(1＋g)(1 − g ÷ RONIC) ÷ (WACC − g) ×終值折現因子；tv_alt＝1 時 (1＋g) ÷ (WACC − g)）",
       lambda vc, c: ("=IF(tv_alt=1,(1+g_tv)/(wacc-g_tv),(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv))*CF!C" + str(CBB["tvdf"])) if vc == "L" else "=0", "0.0000")
    yr("k", "每 1 美元營業利益對 EV 的貢獻（Q4 2026 與 FY2027 為錨點不動，記 0）", lambda vc, c: "=0" if vc in ("F", "G") else f"=(1-tax)*({c}{Y['df']}+{c}{Y['tvf']})", "0.0000")
    r += 1
    f_(RV, f"A{r}", "(a) 現價隱含的預測期營業利益率（FY2028–FY2032 統一一個營業利益率；營收、折舊、資本支出、營運資金不變）", bold=True)
    r += 1
    kv("S", "營收 × 貢獻係數合計（US$M）", f"=SUMPRODUCT(B{Y['rev']}:H{Y['rev']},B{Y['k']}:H{Y['k']})", "#,##0.0")
    kv("E0", "與 FY2028 起營業利益無關的 EV（US$M）", f"=B{A['ev']}-SUMPRODUCT(B{Y['ebit']}:H{Y['ebit']},B{Y['k']}:H{Y['k']})", "#,##0.0")
    kv("m0", "模型：基準的等效營業利益率（FY2028–FY2032，以同一係數加權）", f"=(B{A['ev']}-B{A['E0']})/B{A['S']}", "0.0%", "終值年（FY2032）權重大")
    kv("m1", "現價隱含的營業利益率（GAAP，SBC 費用化）", f"=(B{A['evr']}-B{A['E0']})/B{A['S']}", "0.0%")
    kv("m_chk", "判讀", f'=IF(B{A["m1"]}>1,"超過 100%：單靠營業利益率無法支持現價",IF(B{A["m1"]}<0,"低於 0：現價低於任何正營業利益率的價值",IF(B{A["m1"]}>MAX(D{Y["m"]}:H{Y["m"]}),"高於模型預測期任一年","在模型預測期範圍內")))', "@")
    r += 1
    f_(RV, f"A{r}", "(b) 現價隱含的營收規模（FY2028–FY2032 營收與營業利益等比例調整；營業利益率不變，折舊、資本支出、營運資金不變）", bold=True)
    r += 1
    kv("EB", "FY2028–FY2032 營業利益 × 貢獻係數合計（US$M；調整一倍所改變的 EV）", f"=SUMPRODUCT(B{Y['ebit']}:H{Y['ebit']},B{Y['k']}:H{Y['k']})", "#,##0.0")
    kv("s1", "現價隱含的營收倍數（模型＝1）", f"=1+(B{A['evr']}-B{A['ev']})/B{A['EB']}", "0.00")
    kv("rv1", "現價隱含的 FY2028 營收（US$M）", f"=B{A['s1']}*D{Y['rev']}", "#,##0.0")
    kv("rv0", "模型：FY2028 營收（情境 B；US$M）", f"=D{Y['rev']}", "#,##0.0")
    kv("g1", "現價隱含的 FY2028 營收年增率", f"=B{A['rv1']}/C{Y['rev']}-1", "0.0%")
    kv("g0", "模型：FY2028 營收年增率（情境 B）", f"=D{Y['rev']}/C{Y['rev']}-1", "0.0%")
    r += 1
    f_(RV, f"A{r}", "(c) 現價隱含的 WACC（各年現金流不變；永續期 RONIC＝WACC＋超額報酬差；二分法 40 次）", bold=True)
    r += 1
    for j, h in enumerate(["次", "下界", "上界", "中點", "中點的評價日每股價值（美元）"], 1):
        f_(RV, f"{L_(j)}{r}", h, bold=True, fill=HDR)
    r += 1
    cfr, tr, nor, ffr = CBB["cf"], CBB["t"], CBB["nopat"], CBB["fcff"]
    tvdf_t = f"(DATE(2026,12,31)-val_date)/365+6"

    def ev_at(w):
        cf = f"CF!$F${cfr}:$L${cfr}"
        tvc = (f"IF(tv_alt=1,CF!$L${ffr}*(1+g_tv)/({w}-g_tv),CF!$L${nor}*(1+g_tv)*(1-g_tv/({w}+ronic_sp))/({w}-g_tv))")
        return f"SUMPRODUCT({cf},1/(1+{w})^CF!$F${tr}:$L${tr})+{tvc}/(1+{w})^({tvdf_t})"
    b0 = r
    for i in range(40):
        f_(RV, f"A{r}", i + 1)
        if i == 0:
            f_(RV, f"B{r}", "=g_tv+0.005")
            f_(RV, f"C{r}", "=0.4")
        else:
            f_(RV, f"B{r}", f"=IF(E{r - 1}>$B${A['v0']},D{r - 1},B{r - 1})")
            f_(RV, f"C{r}", f"=IF(E{r - 1}>$B${A['v0']},C{r - 1},D{r - 1})")
        f_(RV, f"D{r}", f"=(B{r}+C{r})/2")
        f_(RV, f"E{r}", f"=({ev_at(f'D{r}')}+$B${A['oth']})/$B${A['sh']}", "0.00")
        for c in "BCD":
            RV[f"{c}{r}"].number_format = "0.0000%"
        r += 1
    kv("w1", "現價隱含的 WACC", f"=D{r - 1}", "0.00%", "搜尋區間：永續成長率＋0.5 個百分點～40%")
    kv("w0", "模型 WACC", "=wacc", "0.00%")
    kv("w_chk", "核對：以模型 WACC 代入同一公式的 EV − 情境 B EV（應為 0）", f"={ev_at('wacc')}-B{A['ev']}", "0.000000")
    kv("w_res", "核對：隱含 WACC 下的評價日每股價值 − 支持現價所需值（美元；在區間內有解時應接近 0）", f"=E{b0 + 39}-B{A['v0']}", "0.000000")
    kv("w_flag", "判讀", f'=IF(ABS(B{r - 1})<0.5,"區間內有解","無解：即使折現率降到搜尋下界，每股價值仍低於支持現價所需值")', "@")
    r += 1
    f_(RV, f"A{r}", "B. Comparable（同業前瞻 P/E）：支持現價所需的倍數", bold=True)
    r += 1
    for j, h in enumerate(["項目", "現價隱含", "模型", "差距（現價隱含 ÷ 模型 − 1）"], 1):
        f_(RV, f"{L_(j)}{r}", h, bold=True, fill=HDR)
    r += 1
    cmpb = VR["cmp0"]
    ntm = f"Val!$C${cmpb + 3}"

    def cmp_(key, label, fi, fm, fmt):
        nonlocal r
        f_(RV, f"A{r}", label)
        f_(RV, f"B{r}", fi, fmt)
        f_(RV, f"C{r}", fm, fmt)
        f_(RV, f"D{r}", f"=B{r}/C{r}-1", "+0.0%;-0.0%")
        A[key] = r
        r += 1
    cmp_("sps", "情境 B 每股價值（目標時點；Comparable，美元）", "=px", f"=Val!C{cmpb + 5}", "0.00")
    cmp_("spe", "隱含 P/E（現價 ÷ 目標時點 NTM 非 GAAP EPS；倍）", f"=px/{ntm}", "=pe_med*rel_mult", "0.0")
    cmp_("scmp", "加權（DCF 不變）：支持現價所需的 Comparable 每股價值（情境加權；美元）", f"=(px-w_dcf*Val!$B${VR['tpdcf']})/(1-w_dcf)", f"=Val!$B${VR['tpcomp']}", "0.00")
    r += 1
    f_(RV, f"A{r}", "C. 摘要（HTML 用；公式帶數字）", bold=True)
    r += 1
    f_(RV, f"A{r}", tf("現金流折現法：情境 B 的企業價值 ", (f"B{A['ev']}", "#,##0"), " 百萬美元，每股價值 ", (f"B{A['v_b']}", "0.00"), " 美元（評價日）；現價 ", ("px", "0.00"),
                      " 美元相當於企業價值為模型的 ", (f"B{A['evx']}", "0.00"), " 倍。單靠一個因素達成時，需要：折現率 ", (f"IF(ABS(B{A['w_res']})<0.5,TEXT(B{A['w1']},\"0.0%\"),\"無解（下界 \"&TEXT(B{A['w1']},\"0.0%\")&\" 仍不足）\")", "@"), "（模型 ", (f"B{A['w0']}", "0.0%"),
                      "）；或 FY2028–FY2032 營收與營業利益為模型的 ", (f"B{A['s1']}", "0.00"), " 倍（FY2028 營收 ", (f"B{A['rv1']}", "#,##0"), " 百萬美元，模型 ", (f"B{A['rv0']}", "#,##0"),
                      "）；或預測期營業利益率 ", (f"B{A['m1']}", "0%"), "（模型等效約 ", (f"B{A['m0']}", "0%"), "，終值年權重大）。"))
    A["sum_dcf"] = r
    r += 1
    f_(RV, f"A{r}", tf("同業倍數法：現價 ", ("px", "0.00"), " 美元相當於目標時點 NTM 非 GAAP EPS 的 ", (f"B{A['spe']}", "0.0"), " 倍 P/E，模型採用同業中位數 ", (f"C{A['spe']}", "0.0"),
                      " 倍；現價對模型 EPS 的倍數比同業中位數", ("IF(B" + str(A['spe']) + "<C" + str(A['spe']) + ",\"低\",\"高\")", "@"), (f"ABS(D{A['spe']})", "0%"), "。"))
    A["sum_cmp"] = r
    RV.column_dimensions["A"].width = 76
    for c in "BCDEFGH":
        RV.column_dimensions[c].width = 14
    A["yh"] = yh
    A["b0"] = b0
    A.update({f"y_{k}": v for k, v in Y.items()})
    return A


def recon(wb, RC, VR, kreq, keyrow):
    PBB, TBB = PB["B"], TB["B"]
    A = {}
    f_(RC, "A1", "Recon — 對帳：公司指引、賣方共識與模型（情境 B）。指引與共識為 Interested-party，只在本頁對照，不進基準（原則第 10 節第 59 條）。口徑：營收＝GAAP 合併營收；非 GAAP 為公司自訂口徑。", bold=True)
    r = 3

    def sec(title, heads):
        nonlocal r
        f_(RC, f"A{r}", title, bold=True)
        r += 1
        for j, h in enumerate(heads, 1):
            f_(RC, f"{L_(j)}{r}", h, bold=True, fill=HDR)
        r += 1

    def row(key, label, vals, fmts, note=""):
        nonlocal r
        f_(RC, f"A{r}", label)
        for j, (v, nf) in enumerate(zip(vals, fmts)):
            if v is not None:
                f_(RC, f"{L_(2 + j)}{r}", v, nf)
        if note:
            f_(RC, f"{L_(2 + len(vals))}{r}", note)
        A[key] = r
        r += 1
    q = Q4
    sec("1. 公司指引對帳（2026-11-04 指引；只作對照）", ["項目", "模型（情境 B）", "指引或共識", "差距（模型相對對照）", "說明"])
    for key, lab, fm, ref, nf, note in [
        ("q4rev", "Q4 2026 營收（US$M）：模型 vs 指引中點", f"=Q4E!B{q['rev']}", "=g_rev_q4", "#,##0.0", "指引 945–947；模型為 Q4E（ASRPC 季增 Assumed）"),
        ("q4ngoi", "Q4 2026 非 GAAP 營業利益（US$M）", f"=Q4E!B{q['ngop']}", "=g_ngoi_q4", "#,##0.0", "指引 216–217；Q4 單季成本率持平於近四季，不含重組節省"),
        ("fy26rev", "FY2026 營收（US$M）", f"=PL!F{PBB['rev']}", "=g_rev26", "#,##0.0", "指引 3,685–3,687"),
        ("fy26ngoi", "FY2026 非 GAAP 營業利益（US$M）", f"=PL!F{PBB['ngop']}", "=g_ngoi26", "#,##0.0", "指引 772–773"),
        ("fy26eps", "FY2026 非 GAAP EPS（US$）", f"=PL!F{PBB['ngeps']}", "=g_eps26", "0.00", "指引 13.59–13.61"),
        ("fy26c", "FY2026 營收（US$M）：模型 vs 賣方共識", f"=PL!F{PBB['rev']}", "=c_rev26", "#,##0.0", "TwelveData 32 位，Q3 後"),
        ("fy27c", "FY2027 營收（US$M）：模型 vs 賣方共識", f"=PL!G{PBB['rev']}", "=c_rev27", "#,##0.0", "TwelveData 32 位（區間 3,980–4,220）；FY2027 EPS 共識（Q3 後）未找到"),
        ("tpc", "目標價（US$）：模型加權 vs 賣方平均", f"=Val!B{VR['tp']}", "=c_tp", "0.00", "TickFlow 49 位；只作對照")]:
        f_(RC, f"A{r}", lab)
        f_(RC, f"B{r}", fm, nf)
        f_(RC, f"C{r}", ref, nf)
        f_(RC, f"D{r}", f"=B{r}/C{r}-1", "+0.0%;-0.0%")
        f_(RC, f"E{r}", tf(("IF(D" + str(r) + ">0,\"高於\",\"低於\")", "@"), "對照 ", (f"ABS(D{r})*100", "0.0"), " 個百分點；", note))
        A[key] = r
        r += 1
    r += 1
    sec("2. 指引差距歸因（Q4 2026 營收）", ["項目", "數值", "說明"])
    row("q4gap", "Q4 營收：模型 − 指引中點（US$M）", [f"=Q4E!B{q['rev']}-g_rev_q4"], ["+#,##0.0;-#,##0.0"], "落差全數來自 Q4 ASRPC 季增假設")
    row("q4_implied", "指引中點隱含的 Q4 ASRPC 季增（假設專業服務與淨增同基準）", [f"=((g_rev_q4/(1+ps_pct))*4/Q4E!B{q['avg']}*1000000)/h_asrpc_q326-1"], ["0.0%"], "模型基準取持平（0%）與此值之間（In 頁第四季每客戶營收季增）")
    row("q4_pct", "落差占 Q4 營收", [f"=B{A['q4gap']}/g_rev_q4"], ["+0.0%;-0.0%"], "")
    row("m27", "管理層 2027 非 GAAP 營業利益率目標（中點）vs 模型 FY2027 非 GAAP 營業利益率", [f"=PL!G{PBB['ngopm']}", "=g_m27"], ["0.0%", "0.0%"], "模型以歷史增量利益率（增量利益率，工作單 G3.1）延伸，不採用管理層目標；管理層 23–24%（2027）、30%（2030）為 Interested-party，只在此對照")
    r += 1
    sec("3. 反解：支持現價所需的主導變數（整本重算；快照）", ["項目", "數值", "說明"])
    kr = kreq or {}
    row("k_key", "主導變數：增量非 GAAP 營業利益率（增量利益率；基準＝FY2023–FY2025 斜率）", [kr.get("base", None)], ["0.0%"], "營收每增 1 美元的營業利益增量；建置時整本重算快照")
    row("k_req", "反解值：使加權目標價＝現價所需的增量利益率", [kr.get("required", None)], ["0.0%"], "割線法，LibreOffice 整本重算；超過 100% 表示單靠此變數無解")
    row("k_m", "反解值對應的 FY2027 非 GAAP 營業利益率", [kr.get("margin27", None)], ["0.0%"], "對照：管理層 2027 目標 23–24%")
    row("k_tp", "核對：反解值代入後的加權目標價（US$）", [kr.get("tp", None)], ["0.00"], "應≈現價")
    row("k_dn", "評等翻轉（中立→放空）：加權目標價上檔空間觸及 −25% 所需的增量利益率", [kr.get("flip_m25", None)], ["0.0%"], "整本重算快照")
    row("k_up", "評等翻轉（中立→買進）：上檔空間觸及 +25% 所需的增量利益率", [kr.get("flip_p25", None)], ["0.0%"], "整本重算快照")
    r += 1
    sec("4. 資料日期核對（相對股價基準日；超過 2 個交易日標逾期；工作單 G2.10）", ["項目", "資料日", "與股價基準日相差（交易日）", "狀態"])
    for key, lab, dn in [("d_rf", "無風險利率", "rf_date"), ("d_pe", "同業前瞻 P/E", "pe_date"), ("d_cons", "賣方共識", "c_date"), ("d_px", "現價", "px_date")]:
        f_(RC, f"A{r}", lab)
        f_(RC, f"B{r}", f"={dn}", "yyyy-mm-dd")
        f_(RC, f"C{r}", f"=NETWORKDAYS(MIN(B{r},px_date),MAX(B{r},px_date))-1", "0")
        f_(RC, f"D{r}", f'=IF(C{r}>2,"逾期","未逾期")')
        A[key] = r
        r += 1
    f_(RC, f"A{r}", "β（StockAnalysis 2026-10-26；MarketBeat 無日期）與同業 EV／Sales 為參考欄，不進模型；β 兩來源日期未對齊，標未核。")
    r += 2
    sec("5. 同業可比性（工作單 G2.6；倍數分母期間、會計準則、租賃處理）", ["同業", "前瞻 P/E（倍）", "會計準則／租賃處理", "備註"])
    pe_rows = [("Salesforce", "pe_crm", "美國 GAAP；營業租賃入資產負債表，不影響 P/E"), ("ServiceNow", "pe_now", "美國 GAAP"), ("Workday", "pe_wday", "美國 GAAP"), ("Adobe", "pe_adbe", "美國 GAAP"),
               ("Intuit", "pe_intu", "美國 GAAP"), ("Atlassian", "pe_team", "美國 GAAP；2022 年起美國申報人"), ("monday.com", "pe_mndy", "美國 GAAP（以色列公司，向 SEC 申報 10-K／20-F 以頁面為準，未核）"),
               ("DocuSign", "pe_docu", "美國 GAAP"), ("Twilio", "pe_twlo", "美國 GAAP"), ("Zoom", "pe_zm", "美國 GAAP"), ("Datadog", "pe_ddog", "美國 GAAP；成長遠高於 HUBS")]
    A["peer0"] = r
    for n, k, note in pe_rows:
        f_(RC, f"A{r}", n)
        f_(RC, f"B{r}", f"={k}", "0.0")
        f_(RC, f"C{r}", note)
        f_(RC, f"D{r}", "分母為前瞻一年每股盈餘，GAAP 或非 GAAP 頁面未載（未核）")
        r += 1
    A["peer_end"] = r - 1
    row("pe_med", "同業中位數（11 家）", ["=pe_med"], ["0.0"], "股權倍數（有槓桿）配本公司非 GAAP EPS（有槓桿）；乘目標時點 NTM EPS（FY2027、FY2028 依日期加權）")
    row("pe_hubs", "HUBS 現行前瞻 P/E（ChartRow）", ["=c_pe_hubs"], ["0.0"], "")
    row("pe_ratio", "現行倍數 ÷ 同業中位數", [f"=B{A['pe_hubs']}/B{A['pe_med']}"], ["0.00"], "")
    row("pe_rr", "隱含重估：若 HUBS 回到同業中位數的倍數，相對現行倍數的變動", [f"=1/B{A['pe_ratio']}-1"], ["+0.0%;-0.0%"], "模型相對倍數係數基準 1.00（未調整）等於假設 HUBS 以同業中位數倍數估值")
    r += 1
    sec("6. AI 淨效果（情境 A／B／C，FY2032E，US$M）", ["項目", "情境 A", "情境 B", "情境 C"])
    for k, lab, key, nf in [("ai_rev", "AI 增量：按任務計價營收", "aiRev", "#,##0.0"), ("ai_loss", "AI 侵蝕：席位被取代的營收損失", "loss", "#,##0.0"), ("ai_inc", "增量推論成本", "inc", "#,##0.0"),
                            ("ai_net", "AI 淨效果", "net", "#,##0.0"), ("ai_brk", "損益兩平每任務單價（US$）", "brk", "0.00"), ("ai_price", "模型每任務單價（US$）", "price", "0.00"),
                            ("ai_npc", "AI 淨效果每客戶每年（US$）", "npc", "#,##0.0"), ("ai_share", "推論成本占非 GAAP 營業成本", "share", "0.0%")]:
        f_(RC, f"A{r}", lab)
        for c, s in zip("BCD", SC):
            f_(RC, f"{c}{r}", f"=Thesis!L{TB[s][key]}", nf)
        A[k] = r
        r += 1
    r += 1
    sec("7. 未找到清單", ["項目", "狀態", "出處"])
    import facts_hubspot as FH
    A["nf0"] = r
    for t in FH.NOT_FOUND:
        f_(RC, f"A{r}", t)
        f_(RC, f"B{r}", "未找到")
        r += 1
    A["nf1"] = r - 1
    RC.column_dimensions["A"].width = 80
    RC.column_dimensions["E"].width = 60
    return A


def readme(wb, RD, build_id, VR, RVA, RCA, vw, SENS):
    TBB, PBB = TB["B"], PB["B"]
    f_(RD, "A1", tf(f"HubSpot（HUBS）個股模型 {VERSION}｜build_id {build_id}｜資料截止日 ", ("data_date", "yyyy-mm-dd"), f"｜成品日 {BUILD_DATE}｜股價基準日 ", ("px_date", "yyyy-mm-dd"), "｜評價日 ", ("val_date", "yyyy-mm-dd"), "｜本版未做先驗"))
    f_(RD, "A2", f"{VERSION} 要點：首版。依 工作單 G 與 G2 建立：命題模組＝原型 S（按席位計價）單位經濟：每客戶訂閱營收受席位被 AI 取代侵蝕，AI 代理按任務計價為增量，三情境；官方分部為單一分部，營收線＝訂閱、專業服務與其他；FCFF DCF 到 FY2032＋同業前瞻 P/E 法各半；推論成本讀 Tokenomics v22.0 快照。未做先驗。")
    f_(RD, "A3", tf("命題：AI 進入 CRM 後，HubSpot 的價值被重新分配：按任務計價的 AI 代理（Customer Agent 每次解決 0.50 美元）是增量，但席位被取代是侵蝕。情境 B 下 FY2032 AI 淨效果 US$M ", (f"Thesis!L{TB['B']['net']}", "#,##0"),
                    "（按任務營收 ", (f"Thesis!L{TB['B']['aiRev']}", "#,##0"), " − 席位取代損失 ", (f"Thesis!L{TB['B']['loss']}", "#,##0"), " − 增量推論成本 ", (f"Thesis!L{TB['B']['inc']}", "#,##0"), "）。"))
    f_(RD, "A4", tf("核心問題：FY2032 的 AI 淨效果在情境 A／B／C 各為 US$M ", (f"Thesis!L{TB['A']['net']}", "#,##0"), "／", (f"Thesis!L{TB['B']['net']}", "#,##0"), "／", (f"Thesis!L{TB['C']['net']}", "#,##0"),
                    "；FY2032 營收情境 B 為 US$M ", (f"PL!L{PBB['rev']}", "#,##0"), "。每任務推論成本為單價的 ", ("Thesis!D9", "0.0%"), "，不是主要風險；席位取代與費用率才是。"))
    f_(RD, "A5", tf("投資建議：", (f"Val!B{VR['rec']}", "@"), "——在『客戶數與每客戶營收依情境路徑、AI 取代席位的侵蝕依情境、非 GAAP 成本率持平於近四季、股票薪酬依規格費用化（不加回）』的基準假設下；加權目標價 US$",
                    (f"Val!B{VR['tp']}", "0.00"), "（DCF ", (f"Val!B{VR['tpdcf']}", "0.00"), "、Comparable ", (f"Val!B{VR['tpcomp']}", "0.00"), "），上檔空間 ", (f"Val!B{VR['up']}", "+0.0%;-0.0%"),
                    "（現價 US$", ("px", "0.00"), "，", ("px_date", "yyyy-mm-dd"), "）。"))
    f_(RD, "A6", tf("重點觀察 1：DCF 與 Comparable 差 US$", (f"ABS(Val!B{VR['gap']})", "0.00"), "；股票薪酬費用化占 US$", (f"Val!B{VR['gap_sbc']}", "0.00"), "，其餘為同業倍數與成長差。現行 HUBS 前瞻 P/E 為同業中位數的 ",
                    (f"Val!B{VR['perel']}", "0.00"), " 倍（相對倍數係數基準 ", ("rel_mult", "0.00"), "，未調整：缺 10 年前瞻倍數）。"))
    f_(RD, "A7", tf("重點觀察 2：距評等界線 ", (f"Val!B{VR['bdist']}", "0.0"), " 個百分點（最近界線 ", (f"Val!C{VR['bdist']}", "@"), "）；", (f"Val!E{VR['bdist']}", "@"), "。"))
    f_(RD, "A8", tf("重點觀察 3：收支關卡（FY2026–FY2030 自由現金流減股利，情境 B）US$M ", (f"Val!B{VR['gatev']}", "#,##0"), "，", (f"IF(Val!B{VR['gatev']}<0,\"未通過\",\"通過\")", "@"),
                    "；DCF 終值占 EV（情境 B）", (f"Val!C{VR['tvsh2']}", "0%"), "；現價要成立，FY2028 起現金流須為模型基準的 ", (f"RevDCF!B{RVA['evx']}", "0.0"), " 倍。"))
    f_(RD, "A9", tf("指引對照：Q4 2026 營收模型 US$M ", (f"Recon!B{RCA['q4rev']}", "#,##0"), " 對指引中點 ", (f"Recon!C{RCA['q4rev']}", "#,##0"), "（", (f"Recon!D{RCA['q4rev']}", "+0.0%;-0.0%"), "）；"
                    "FY2027 非 GAAP 營業利益率模型 ", (f"Recon!B{RCA['m27']}", "0.0%"), " 對管理層目標中點 ", (f"Recon!C{RCA['m27']}", "0.0%"), "（模型不預設費用率改善；指引與目標只在 Recon 對照）。"))
    f_(RD, "A10", "驅動 → 推導對照表（每頁輸入、推導量、因果方向）", bold=True)
    for j, h in enumerate(["頁", "輸入（In 頁鍵）", "推導量", "因果方向"], 1):
        f_(RD, f"{L_(j)}11", h, bold=True)
    mp = [("Thesis", "na_q4、q4_asrpc、na_27／na_32、g_up", "客戶數、每客戶訂閱營收", "客戶與每客戶營收年增上升 → 營收上升"),
          ("Thesis", "seat_sh、repl_A／B／C", "席位被取代的侵蝕與損失", "取代率與席位占比上升 → 每客戶營收下降"),
          ("Thesis", "ai_rev_26、tasks_0、g_tasks、price_0、adopt_A／B／C、dp_A／B／C", "按任務計價營收（AI 增量）", "採用率、任務數、單價上升 → 增量營收上升"),
          ("Thesis", "tok_*、ctx_m、succ、mix_*、tkp_*、route、tk_pscale、link_pass、pd_tok、emb_pct、emb_g（Tokenomics 快照）", "增量推論成本與損益兩平單價", "token 價格與用量上升 → 推論成本上升 → AI 淨效果下降"),
          ("PL", "ng_cogs、ng_opex、sbc_pct、adj_pct、ps_pct、cogs_adj_sh", "毛利、營業利益、EPS（GAAP 與非 GAAP）", "費用率下降 → 營業利益率上升"),
          ("CF", "da_pct、cx_pct、nwc_k、tax、int_y、bb_pct", "FCFF、收支關卡、現金與股數", "資本支出上升 → FCFF 下降；營運資金流入上升 → FCFF 上升"),
          ("Val", "rf、beta、erp、g_tv、ronic_sp、pe_*、rel_mult、w_dcf、股數（sh_basic、rsu_dil、forfeit、opt_dil）", "DCF、Comparable、加權目標價、投資建議", "WACC 上升 → DCF 下降；同業倍數與係數上升 → Comparable 上升")]
    for i, rw in enumerate(mp):
        for j, v in enumerate(rw):
            f_(RD, f"{L_(j + 1)}{12 + i}", v)
    f_(RD, "A20", "版本要點", bold=True)
    f_(RD, "A21", "首版，無前版可比；數值皆為本版輸入的函數。")
    f_(RD, "A22", "官方分部：單一應報導分部；營收線訂閱、專業服務與其他；對帳核對列為 0（PL 頁各情境區塊、Hist 頁）。")
    f_(RD, "A23", "推論成本讀 Tokenomics v22.0 快照（data/tokenomics_snapshot_v22.0.json）；不自行估 token 價格；價格年降標「待 AI供應鏈 登錄」。")
    f_(RD, "A24", "研究包：data/research_hubspot.md、facts_hubspot.csv、segments.json。")
    f_(RD, "A25", "模型狀態（README、Sens、HTML 三處 build_id 相同；Sens 頁比對輸入指紋）", bold=True)
    f_(RD, "A26", f'="build_id "&Sens!D{SENS["bid"]}&"；模型狀態 "&Sens!D{SENS["status"]}&"（Sens 與 Recon 反解為建置快照：只在 Excel 改 In 頁而未重跑 make_v01.sh 時顯示 STALE）"')
    f_(RD, "A28", "市場資料日期、失效條件與可證偽指標", bold=True)
    f_(RD, "A29", tf("市場資料日期：現價 ", ("px_date", "yyyy-mm-dd"), "（兩處同值）；rf ", ("rf_date", "yyyy-mm-dd"), "（單一來源）；同業倍數 ", ("pe_date", "yyyy-mm-dd"), "；賣方共識 ", ("c_date", "yyyy-mm-dd"), "；公司指引 2026-11-04（Q3 2026）；逾期核對見 Recon 第 4 節。"))
    f_(RD, "A30", "下一個事件：Q4 2026 財報預計 2027-02 公布（含 FY2027 指引與 seats-and-credits EMEA 試點結果）；共識與指引屆時更新。")
    f_(RD, "A31", tf("承重 Assumed 1（增量非 GAAP 營業利益率 inc_m，歷史斜率延伸；全期持平為 Sens 下緣，反解值見 Recon 第 3 節）。可證偽指標：FY2027 各季非 GAAP 營業利益率對管理層 2027 目標 23–24%；模型 FY2027 ", (f"PL!G{PBB['ngopm']}", "0.0%"), "，FY2032 ", (f"PL!L{PBB['ngopm']}", "0.0%"), "。"))
    f_(RD, "A32", tf("承重 Assumed 2（席位取代：repl_*、seat_sh）。可證偽指標：淨營收留存率（FY2026 預期較 FY2025 降約 2 個百分點；Q3 為 101%）、ASRPC 軌跡（Q3 12,038）；模型 FY2027 每客戶訂閱營收年增 ", (f"Thesis!G{TBB['arpc']}/Thesis!F{TBB['arpc']}-1", "0.0%"), "。"))
    f_(RD, "A33", tf("承重 Assumed 3（客戶淨增率與每客戶年增：na_27／na_32、g_up）。可證偽指標：Q4 2026 淨增（管理層稱低於 5,000）與 FY2027 各季客戶淨增；模型 FY2027 客戶年增 ", (f"Thesis!G{TBB['na']}", "0.0%"), "。"))
    f_(RD, "A34", "失效條件（可觀察）：seats-and-credits 方案在 EMEA 試點後擴及現有客戶而 ASP 明顯下降；NRR 低於 100%；agent 付費任務數公司揭露且高於情境 A；非 GAAP 營業利益率超過 23%（費用率證據出現）。")
    f_(RD, "A36", "風險登錄", bold=True)
    for j, h in enumerate(["風險", "是否建模", "理由", "觀察指標"], 1):
        f_(RD, f"{L_(j)}37", h, bold=True)
    risks = [("AI 改變客戶發現、評估與購買軟體的方式（管理層：Q2 逆風在 Q3 更明顯，預期延續到 2027）", "是（客戶淨增率與 g_up 低案；情境 C）", "以客戶增速與每客戶營收路徑呈現，不單獨建事件", "客戶淨增、Net New ARR、新客 ASP"),
             ("2026 Restructuring Plan（約 7% 員工）執行風險與節省", "部分（Q4 費用建模；節省不入基準）", "節省無證據，只放 Sens（ng_opex 低案）", "2027 各季非 GAAP 營業費用率"),
             ("seats-and-credits 新方案降低初始 ASP", "是（g_up 低案、seat_sh）", "試點 ASP 較低、credits 用量 2–3 倍，數字未揭露", "2027-02 EMEA 試點結果"),
             ("股票薪酬占營收 14%，費用化後 GAAP 利益率低", "是（DCF 不加回；sbc_ab 敏感度）", "SBC 比率持平於近四季", "SBC 占營收、稀釋股數"),
             ("回購對每股盈餘的支撐（9M26 回購 973.7 百萬美元）", "是（bb_pct，只影響 EPS 與現金）", "不影響 FCFF DCF；影響 Comparable 的 EPS", "季回購金額、授權餘額 964.1 百萬美元")]
    for i, rw in enumerate(risks):
        for j, v in enumerate(rw):
            f_(RD, f"{L_(j + 1)}{38 + i}", v)
    f_(RD, "A44", "命題回覆", bold=True)
    f_(RD, "A45", tf("1. AI 為 HubSpot 帶來增量還是侵蝕：FY2032 情境 A／B／C 的 AI 淨效果 US$M ", (f"Thesis!L{TB['A']['net']}", "#,##0"), "／", (f"Thesis!L{TB['B']['net']}", "#,##0"), "／", (f"Thesis!L{TB['C']['net']}", "#,##0"),
                    "（按任務營收減席位取代損失與增量推論成本）；按任務計價營收在情境 A 亦僅占 FY2032 營收 ", (f"Thesis!L{TB['A']['aiRev']}/Thesis!L{TB['A']['rev']}", "0.0%"), "。"))
    f_(RD, "A46", tf("2. 推論成本是否吃掉毛利：情境 B 的 FY2032 推論成本占非 GAAP 營業成本 ", (f"Thesis!L{TB['B']['share']}", "0.0%"), "；每解決任務推論成本占單價 ", ("Thesis!D9", "0.0%"), "；損益兩平每任務單價 US$", (f"Thesis!L{TB['B']['brk']}", "0.00"), "。"))
    f_(RD, "A47", tf("3. 現價隱含什麼：FY2028 起現金流須為情境 B 基準的 ", (f"RevDCF!B{RVA['evx']}", "0.0"), " 倍；或目標時點 NTM 非 GAAP EPS 的 P/E 為 ", (f"RevDCF!B{RVA['spe']}", "0.0"), " 倍（同業中位數 ", ("pe_med", "0.0"), "）。"))
    f_(RD, "A49", "未做先驗：本版未套用任何先驗（依工作單先跳過）；公開 repo，內容只含公司與市場事實。")
    f_(RD, "A50", "不適用：算力 MW、Tokenomics 自用 AI 讀法 2、Wattonomics（HubSpot 不自建算力，推論成本向模型商購買）。")
    f_(RD, "A51", "不適用：三層協調（Coord 頁、C0、兩讀數）：非 Siliconomics 覆蓋。")
    f_(RD, "A52", "未做：現金稅率橋（原則第 10 節第 40 條）；事件型尾端 Sens；選擇權庫藏股法（履約價未找到）與期末未歸屬 RSU（用 Q3 反稀釋加權代理）；每任務 token 公司揭露（未找到，以 Tokenomics 類比）。")
    f_(RD, "A53", "v0.2 待辦：改讀 software_ai/unit_economics 的 S 原型；補 seats-and-credits 揭露；以 10-K 核對 FY2021–FY2022 欄位。")
    f_(RD, "A54", "工作單 G2 十條落實：見 data/research_hubspot.md 末節與 README.md。")
    f_(RD, "A56", tf("席位被 AI 取代的侵蝕大於按任務計價的增量；加權目標價 US$", (f"Val!B{VR['tp']}", "0.00"), "，現價 US$", ("px", "0.00"), "，評等", (f"Val!B{VR['rec']}", "@"), "；DCF 受股票薪酬費用化拖累。"))
    RD.column_dimensions["A"].width = 120


def sens_sheet(wb, S, sens, build_id, inhash):
    f_(S, "A1", "Sens — 一次改一個輸入（In 頁敏感度低／高值），整本活頁簿以 LibreOffice 重算後的加權目標價（Val 頁）與 DCF、Comparable 單法目標價；由 敏感度程式 產生。")
    f_(S, "A2", "相對基準＝案例目標價 ÷ 基準加權目標價 − 1。token 價格高案只動成本一邊（上界）；另列價格與成本聯動案與路由案（工作單 G2.2）。情境專屬列（席位取代率、採用率、單價變動）以 A／B／C 三情境並列為其區間，不另列案。")
    for j, h in enumerate(["驅動", "設定", "方向", "加權目標價", "相對基準", "鍵", "說明", "DCF 單法目標價", "Comparable 單法目標價"], 1):
        f_(S, f"{L_(j)}3", h, bold=True, fill=HDR)
    cases = (sens or {}).get("cases", [])
    base_tp = (sens or {}).get("base", {})
    rr = 4
    case_row = {}
    for cs in cases:
        f_(S, f"A{rr}", cs["label"])
        f_(S, f"B{rr}", cs["setting"])
        f_(S, f"C{rr}", cs["side"])
        f_(S, f"D{rr}", cs["tp"], "0.00")
        f_(S, f"E{rr}", f"=D{rr}/Val!$B$8-1", "+0.0%;-0.0%")
        f_(S, f"F{rr}", cs["key"])
        f_(S, f"G{rr}", "重算值（整本重算）")
        f_(S, f"H{rr}", cs["dcf"], "0.00")
        f_(S, f"I{rr}", cs["cmp"], "0.00")
        case_row[(cs["key"], cs["side"])] = rr
        rr += 1
    rr += 1
    f_(S, f"A{rr}", "build_id")
    f_(S, f"D{rr}", build_id)
    bid = rr
    rr += 1
    f_(S, f"A{rr}", "模型狀態")
    f_(S, f"D{rr}", "CURRENT" if sens else "未跑敏感度")
    st = rr
    rr += 1
    f_(S, f"A{rr}", "註：In 頁輸入指紋（敏感度重算時的基準輸入雜湊）")
    f_(S, f"D{rr}", inhash)
    rr += 1
    f_(S, f"A{rr}", "註：基準加權目標價（敏感度重算時）／本檔 Val 頁")
    f_(S, f"D{rr}", base_tp.get("tp"), "0.00")
    f_(S, f"E{rr}", "=Val!B8", "0.00")
    f_(S, f"F{rr}", f'=IF(ABS(D{rr}-E{rr})<0.005,"一致","STALE")')
    rr += 2
    f_(S, f"A{rr}", "加權目標價相對基準（HTML 龍捲風圖用；白話標籤）", bold=True)
    t0 = rr
    rr += 1
    for j, h in enumerate(["輸入（白話）", "輸入取低值", "輸入取高值"], 1):
        f_(S, f"{L_(j)}{rr}", h, bold=True, fill=HDR)
    rr += 1
    keys_seen = []
    for cs in cases:
        if cs["key"] not in keys_seen:
            keys_seen.append(cs["key"])
    for k in keys_seen:
        lab = next(c["label"] for c in cases if c["key"] == k).split("：")[0]
        if "+" in k:
            lab = "token 價格與任務售價聯動（0.5／2 倍）"
        f_(S, f"A{rr}", lab)
        if (k, "低") in case_row:
            f_(S, f"B{rr}", f"=E{case_row[(k, '低')]}", "+0.0%;-0.0%")
        if (k, "高") in case_row:
            f_(S, f"C{rr}", f"=E{case_row[(k, '高')]}", "+0.0%;-0.0%")
        rr += 1
    S.column_dimensions["A"].width = 60
    return {"bid": bid, "status": st, "t0": t0, "first": 4, "last": 3 + len(cases)}


def sources(SR):
    f_(SR, "A1", "Sources — 來源（編號、文件、日期、網址、標記、用途）；逐列出處見 data/facts_hubspot.csv", bold=True)
    for j, h in enumerate(["編號", "文件", "日期", "網址", "標記", "用途"], 1):
        f_(SR, f"{L_(j)}3", h, bold=True, fill=HDR)
    import facts_hubspot as FH
    use = {"K25": "FY2021–FY2025 財報、KPI、員工", "K23": "FY2021–FY2022 財報與 KPI", "PR23": "FY2022–FY2023 非 GAAP", "PR_Q4": "Q4 2025 與 FY2025", "PR_Q1": "Q1 2026", "PR_Q2": "Q2 2026",
           "PR_Q3": "Q3 2026、指引、資產負債表", "Q3Q": "10-Q：股數、債務、重組、AI 成本科目", "PPC": "法說摘要（第二來源）", "WB": "法說摘要（第二來源）", "HSP": "Customer Agent／Prospecting Agent 價格",
           "SAAS": "價格轉換（第二來源）", "CR": "同業倍數、現價（第二來源）", "QQ": "現價", "TE": "rf", "SA": "β", "MB": "β", "TF": "賣方目標價", "TG": "賣方目標價（對照）", "TD": "賣方營收共識",
           "SAF": "財報前共識（對照，不採）", "MS": "財報前共識（對照，不採）"}
    for i, (k, (url, doc)) in enumerate(FH.SRC.items()):
        tag = "Verified" if k in ("K25", "K23", "PR23", "PR_Q4", "PR_Q1", "PR_Q2", "PR_Q3", "Q3Q", "QQ", "TE") else "Interested-party" if k in ("PPC", "WB", "HSP", "SAAS", "TF", "TG", "TD", "SAF", "MS") else "Analogy"
        for j, v in enumerate([f"S{i + 1:02d}", doc, "", url, tag, use.get(k, "")]):
            f_(SR, f"{L_(j + 1)}{4 + i}", v)
    n = 4 + len(FH.SRC)
    f_(SR, f"A{n}", f"S{len(FH.SRC) + 1:02d}")
    f_(SR, f"B{n}", "Tokenomics v22.0 快照（data/tokenomics_snapshot_v22.0.json）")
    f_(SR, f"C{n}", "2026-11-16")
    f_(SR, f"D{n}", "tools/tokenomics/import_tokenomics.py 匯出；--check 通過")
    f_(SR, f"E{n}", "Derived")
    f_(SR, f"F{n}", "每任務 token、token 市場價、成功率（Analogy）")
    SR.column_dimensions["B"].width = 60
    SR.column_dimensions["D"].width = 90


def main():
    argv = sys.argv[1:]
    out = argv[0]
    sens = json.load(open(argv[argv.index("--sens") + 1], encoding="utf-8")) if "--sens" in argv and os.path.exists(argv[argv.index("--sens") + 1]) else None
    kreq = json.load(open(argv[argv.index("--kreq") + 1], encoding="utf-8")) if "--kreq" in argv and os.path.exists(argv[argv.index("--kreq") + 1]) else None
    bid = build(out, ov=parse_ov(argv), sens=sens, kreq=kreq)
    print(bid)


if __name__ == "__main__":
    main()
