"""智譜 v0.1-Z2：需求頁 Demand 與營收頁 Revenue，以及 Z2 的 Checks 列。

版面：A 編號、B 項目、C 單位、D–I 2025–2030、J 空白、K 1H2026、L 2H2026、M 說明。
2026（E 欄）＝1H26 實際＋2H26 驅動（規格 D14）：流量列 E＝K＋L；存量列（訂閱者）E＝(K＋L)÷半年數；牌價列 E 只作顯示（半年平均）。
所有數值引用 SRC_ZP／Inputs／TK_Link（E6：公式不內含常數；恆等式 1−x、1＋成長率除外）。Python 只組公式文字。
跨頁列號以「«D.鍵»」「«V.鍵»」佔位，全部列建好後由 fill() 代換。
"""
from __future__ import annotations

import re

from openpyxl.workbook.defined_name import DefinedName

from common import F_BOLD, F_CALC, F_NOTE, FILL_SEC, header, put, title

YC = "DEFGHI"                   # 2025–2030
HC = "KL"                       # 1H2026、2H2026
ALL = ("D", "K", "L", "E", "F", "G", "H", "I")
LATE = "FGHI"                   # 2027–2030
YEARS = (2025, 2026, 2027, 2028, 2029, 2030)
PREV = {"F": "E", "G": "F", "H": "G", "I": "H"}          # 年度鏈（2027 起以 2026 全年為基期）
PLANS = ("lite", "pro", "max")
PLAN_ZH = {"lite": "Lite", "pro": "Pro", "max": "Max"}
TIERS = ("sol", "luna")
TIER_ZH = {"sol": "Sol（GLM 旗艦）", "luna": "Luna（Air／Flash）"}
IO = ("in", "cache", "out")
IO_ZH = {"in": "輸入", "cache": "快取命中輸入", "out": "輸出"}
SEGS = ("agent", "gpllm", "techsvc")
SEG_ZH = {"agent": "②企業級智能體", "gpllm": "③企業級通用大模型", "techsvc": "④技術服務及其他（含 C 端）"}
PH = re.compile(r"«([DVCKFX])\.([A-Za-z0-9_]+)»")


def nm(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


class Sheet:
    def __init__(self, wb, name, prefix, t1, t2, t3=None):
        self.wb, self.name, self.prefix = wb, name, prefix
        self.ws = wb.create_sheet(name)
        title(self.ws, t1, t2, t3)
        header(self.ws, 5, ["編號", "項目", "單位"] + [str(y) for y in YEARS] + ["", "1H2026", "2H2026", "說明"])
        put(self.ws, "B6", "年度", F_BOLD)
        self.r, self.n, self.rows, self.meta = 7, 0, {}, []

    def years(self, D):
        for i, y in enumerate(YEARS):
            put(self.ws, f"{YC[i]}6", f"={D.I(f'year_{y}')}")
        put(self.ws, "K6", "1H2026（實際）", F_BOLD)
        put(self.ws, "L6", "2H2026（驅動）", F_BOLD)

    def section(self, label):
        put(self.ws, f"A{self.r}", label, F_BOLD, fill=FILL_SEC)
        self.r += 1

    def add(self, label, unit, fs, note="", key=None, name=None, name_h=None, name_col=None):
        """fs：dict 欄→公式，或 callable(欄)→公式／None。name：D:I（或 name_col 指定單格）；name_h：K:L。"""
        self.n += 1
        code = f"{self.prefix}{self.n:02d}"
        r = self.r
        self.rows[key or code] = r
        put(self.ws, f"A{r}", code, F_CALC)
        put(self.ws, f"B{r}", label, F_CALC)
        put(self.ws, f"C{r}", unit, F_CALC)
        for c in ALL:
            f = fs(c) if callable(fs) else fs.get(c)
            if f is not None:
                put(self.ws, f"{c}{r}", f)
        put(self.ws, f"M{r}", note, F_NOTE)
        if name:
            nm(self.wb, name, f"{self.name}!${name_col}${r}" if name_col else f"{self.name}!$D${r}:$I${r}")
        if name_h:
            nm(self.wb, name_h, f"{self.name}!$K${r}:$L${r}")
        self.meta.append((code, label, unit, note, name, key))
        self.r += 1
        return r


def fill(wb, Z):
    rows = {k: Z[k].rows for k in ("D", "V", "C", "K", "F", "X") if k in Z}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and "«" in c.value:
                    c.value = PH.sub(lambda m: str(rows[m.group(1)][m.group(2)]), c.value)


def build(wb, D, tk_cells):
    S, I = D.S, D.I
    halves, pct, divrev = I("halves_year"), I("pct"), I("div_rev")
    days = {"D": I("days_year"), "K": I("days_1h26"), "L": I("days_2h26"), **{c: I("days_year") for c in LATE}}
    months = {"D": I("months_year"), "K": I("months_half"), "L": I("months_half"), **{c: I("months_year") for c in LATE}}
    fx = I("fx_usdcny")

    # ═════════════════════ Demand（先建；價格列以 «V.» 佔位） ═════════════════════
    Dm = Sheet(wb, "Demand", "D",
               "Demand — 需求：API 按量計費 token、GLM Coding Plan 訂閱者與 token、智譜清言（免費）token；輸出 DEM_Tok_*（層級 Sol／Luna × 付費／免費）",
               "2025 與 1H26 為校準期：API 按量計費 token＝（開放平台及 API 實際營收 − Coding Plan 估計營收）÷ 組合有效單價（規格 D14；單價為驅動、token 為倒推）。"
               "2H26 起：任務數 ×（1＋成長）× 每任務 token ×（1＋成長）× 價格反應 (p/p前)^ε（D15）。本地化部署不耗智譜算力，不進本頁（D4）。",
               "token 單位 T＝10^12；Coding Plan 訂閱者單位萬人；清言用戶百萬人；任務數十億。容量上限於 Z3 接 Compute；本頁不截頂。")
    Dm.years(D)
    Dm.section("一、API 按量計費（開放平台及 API 扣除 Coding Plan）")
    Dm.add("價格反應底數 p/p前（組合有效單價）", "倍",
           {"L": "=Revenue!L«V.p»/Revenue!K«V.p»", **{c: f"=Revenue!{c}«V.p»/Revenue!{'L' if c == 'F' else PREV[c]}«V.p»" for c in LATE}},
           "2H26＝2H26 ÷ 1H26；2027＝2027 ÷ 2H26（最新價格水準）；之後逐年", key="ratio")
    g_tpt = {"L": I("g_tpt_2h26"), **{c: I("g_tpt") for c in LATE}}
    g_task = {"L": I("g_task_2h26"), "F": I("g_task_2027"), "G": I("g_task_2028"), "H": I("g_task_2029"), "I": I("g_task_2030")}
    elK, elN = I("el_tpt"), I("el_task")

    def tpt(c):
        if c in ("D", "K"):
            return f"={I('tpt_base')}"
        if c == "L":
            return f"=K«D.tpt»*(1+{g_tpt['L']})*L«D.ratio»^{elK}"
        if c == "E":
            return "=E«D.api_tok»/E«D.tasks»"
        return f"={PREV[c]}«D.tpt»*(1+{g_tpt[c]})*{c}«D.ratio»^{elK}"
    Dm.add("API 每任務 token", "K tok／任務", tpt,
           "校準期＝Inputs 基準（只影響任務數與每任務 token 的水準拆分，不影響營收）；2H26 起 ×（1＋成長）× 價格反應^ε（每任務 token）；2026 欄＝全年 token ÷ 全年任務", key="tpt")

    def tasks(c):
        if c in ("D", "K"):
            return f"={c}«D.api_tok»/{c}«D.tpt»"
        if c == "L":
            return f"=K«D.tasks»*(1+{g_task['L']})*L«D.ratio»^{elN}"
        if c == "E":
            return "=K«D.tasks»+L«D.tasks»"
        return f"={PREV[c]}«D.tasks»*(1+{g_task[c]})*{c}«D.ratio»^{elN}"
    Dm.add("API 任務數", "十億", tasks, "校準期＝token ÷ 每任務 token；2H26＝1H26 ×（1＋半年成長）× 價格反應^ε；2027 起以 2026 全年為基期", key="tasks")

    def api_tok(c):
        if c in ("D", "K"):
            return f"=Revenue!{c}«V.calib»*{divrev}/Revenue!{c}«V.p»"
        if c == "E":
            return "=K«D.api_tok»+L«D.api_tok»"
        return f"={c}«D.tasks»*{c}«D.tpt»"
    Dm.add("API 按量計費 token", "T", api_tok,
           "2025、1H26：按量計費營收（API 實際 − Coding Plan 估計）× 100 ÷ 組合有效單價（元／百萬 token；×100 為 Inputs 定義常數）；之後＝任務數 × 每任務 token",
           key="api_tok", name="DEM_Tok_API", name_h="DEM_Tok_API_H")
    for t in TIERS:
        Dm.add(f"API 按量計費 token：{TIER_ZH[t]}", "T",
               lambda c, t=t: "=K«D.api_tok_%s»+L«D.api_tok_%s»" % (t, t) if c == "E" else f"={c}«D.api_tok»*Revenue!{c}«V.mix_{t}»",
               "token × 層級組合（Revenue）", key=f"api_tok_{t}", name=f"DEM_Tok_API_{t.capitalize()}")
    Dm.add("免費 API 型號 token（Luna；耗算力無營收）", "T",
           lambda c: "=K«D.api_free»+L«D.api_free»" if c == "E" else f"={c}«D.api_tok»*{I('free_api_ratio')}",
           "GLM-4.7-Flash 等免費型號（SRC_ZP_333）＝付費 API token × Inputs 比例（D24r：免費型號單列）", key="api_free", name="DEM_Tok_APIFree")

    Dm.section("二、GLM Coding Plan（訂閱者 × 每訂閱者 token；收入在開放平台及 API 內，D2r）")
    cpg = {"F": I("cp_g_2027"), "G": I("cp_g_2028"), "H": I("cp_g_2029"), "I": I("cp_g_2030")}

    def subs(c):
        if c == "D":
            return f"={S('usr_cp_paying_dev')}*{I('cp_subs_r2025')}"
        if c == "K":
            return f"={S('usr_cp_paying_dev')}*{I('cp_subs_r1h26')}"
        if c == "L":
            return f"=K«D.subs»*(1+{I('cp_g_2h26')})"
        if c == "E":
            return f"=(K«D.subs»+L«D.subs»)/{halves}"
        return f"={PREV[c]}«D.subs»*(1+{cpg[c]})"
    Dm.add("Coding Plan 付費訂閱者（期間平均）", "萬人", subs,
           "2025、1H26＝2026-03 付費開發者 24.2 萬（SRC_ZP_402）× Inputs 倍數（營收未單列，無法倒推；工作單：給區間）；2H26＝1H26 ×（1＋半年成長）；2026 欄＝兩個半年平均；2027 起 ×（1＋年增率）",
           key="subs", name="DEM_Subs_CP", name_h="DEM_Subs_CP_H")
    for p in PLANS:
        Dm.add(f"Coding Plan 訂閱者：{PLAN_ZH[p]}", "萬人", lambda c, p=p: f"={c}«D.subs»*Revenue!{c}«V.mix_{p}»", "訂閱者 × 方案組合（Revenue）",
               key=f"subs_{p}", name=f"DEM_Subs_{PLAN_ZH[p]}")
    for p in PLANS:
        f = f"={S(f'cp_quota_{p}_5h')}*{I('windows_day')}*{I('cp_util')}*{I('cp_calls_prompt')}*{I('cp_ktok_call')}/{I('div_k_m')}"
        Dm.add(f"每訂閱者每日 token：{PLAN_ZH[p]}", "M tok／日", {c: f for c in ALL},
               "每 5 小時額度（SRC_ZP_376–378）× 每日窗口數 × 額度使用率 × 每 prompt 呼叫次數 × 每次呼叫 token（K）÷ 1000", key=f"cptpd_{p}")
    for p in PLANS:
        Dm.add(f"Coding Plan token：{PLAN_ZH[p]}", "T",
               lambda c, p=p: "=K«D.cptok_%s»+L«D.cptok_%s»" % (p, p) if c == "E" else
               f"={c}«D.subs_{p}»*{c}«D.cptpd_{p}»*{days[c]}/{I('div_wan_m_t')}",
               "訂閱者（萬）× 每日 token（M）× 天數 ÷ 100（萬 × M → T）", key=f"cptok_{p}")
    Dm.add("Coding Plan token 合計", "T", lambda c: "=" + "+".join(f"{c}«D.cptok_{p}»" for p in PLANS), "", key="cptok", name="DEM_Tok_CP", name_h="DEM_Tok_CP_H")
    Dm.add("Coding Plan token：Sol", "T", lambda c: f"={c}«D.cptok»*{I('cp_sol_share')}", "Coding Plan 以 GLM-5.x 旗艦為主（D24r）", key="cptok_sol")
    Dm.add("Coding Plan token：Luna", "T", lambda c: f"={c}«D.cptok»*(1-{I('cp_sol_share')})", "", key="cptok_luna")

    Dm.section("三、智譜清言（C 端；無單列收入，免費 token 耗算力；年度）")
    Dm.add("智譜清言月活用戶", "百萬人",
           lambda c: f"={I('qy_mau_2025')}" if c == "D" else (f"={YC[YC.index(c) - 1]}«D.qy_users»*(1+{I('qy_g')})" if c in YC else None),
           "2025 Inputs（QuestMobile 未進前 20，SRC_ZP_434；Assumed 區間）；之後 ×（1＋年增率）", key="qy_users", name="DEM_Users_Consumer")
    Dm.add("智譜清言 token", "T",
           lambda c: f"={c}«D.qy_users»*{I('qy_tasks')}*{I('qy_ktok_task')}*{I('days_year')}/{I('div_m_k_t')}" if c in YC else None,
           "用戶（百萬）× 每日任務 × 每任務 token（K）× 天數 ÷ 1000（百萬 × K → T）", key="qy_tok", name="DEM_Tok_Consumer")
    Dm.add("智譜清言 token：Sol", "T", lambda c: f"={c}«D.qy_tok»*{I('qy_sol_share')}" if c in YC else None, "", key="qy_sol")
    Dm.add("智譜清言 token：Luna", "T", lambda c: f"={c}«D.qy_tok»*(1-{I('qy_sol_share')})" if c in YC else None, "", key="qy_luna")

    Dm.section("四、輸出：token 量（層級 × 付費／免費；供 Z3 推論 GW）")
    yo = lambda f: (lambda c: f(c) if c in YC else None)  # noqa: E731
    Dm.add("付費 token：Sol", "T", yo(lambda c: f"={c}«D.api_tok_sol»+{c}«D.cptok_sol»"), "API 按量計費 Sol ＋ Coding Plan Sol", key="paid_sol", name="DEM_Tok_Paid_Sol")
    Dm.add("付費 token：Luna", "T", yo(lambda c: f"={c}«D.api_tok_luna»+{c}«D.cptok_luna»"), "API 按量計費 Luna ＋ Coding Plan Luna", key="paid_luna", name="DEM_Tok_Paid_Luna")
    Dm.add("免費 token：Sol", "T", yo(lambda c: f"={c}«D.qy_sol»"), "智譜清言 Sol", key="free_sol", name="DEM_Tok_Free_Sol")
    Dm.add("免費 token：Luna", "T", yo(lambda c: f"={c}«D.qy_luna»+{c}«D.api_free»"), "智譜清言 Luna ＋ 免費 API 型號", key="free_luna", name="DEM_Tok_Free_Luna")
    Dm.add("付費 token 合計", "T", yo(lambda c: f"={c}«D.paid_sol»+{c}«D.paid_luna»"), "", key="paid", name="DEM_Tok_Paid")
    Dm.add("免費 token 合計", "T", yo(lambda c: f"={c}«D.free_sol»+{c}«D.free_luna»"), "", key="free", name="DEM_Tok_Free")
    Dm.add("token 合計（付費＋免費）", "T", yo(lambda c: f"={c}«D.paid»+{c}«D.free»"), "", key="tok_all", name="DEM_Tok_Total")

    Dm.section("五、對照（只列差距，不校準）")
    Dm.add("2H25→1H26 實際隱含量成長（營收倍數 ÷ 模型單價倍數）", "倍",
           {"K": f"=({S('rev_api_1h26')}/({S('rev_api_fy25')}-{S('rev_api_1h25')}))/(Revenue!K«V.p»/Revenue!D«V.p»)"},
           "開放平台及 API：1H26 ÷ 2H25（＝FY25 − 1H25）實際營收，除以 1H26 ÷ 2025 組合有效單價；Inputs 2H26 任務數成長的參考（實際值，非公司說法）",
           key="impl_vol", name="DEM_ImpliedVolGrowth", name_col="K")
    Dm.add("模型 2H26 ÷ 1H26 API 按量計費 token", "倍", {"L": "=L«D.api_tok»/K«D.api_tok»"}, "對照：公司稱 MaaS token 呼叫量『較年初 40 倍以上』（SRC_ZP_410；基期為年初時點，口徑不同）", key="vol_ratio")
    Dm.add("公司說法：MaaS token 呼叫量較年初倍數", "倍", {"K": f"={S('tok_maas_growth_ytd')}"}, "Interested-party；只對照，不反推（D14r）", key="vol_claim")
    Dm.add("模型 2025 token 合計 ÷ 天數", "T／日", {"D": f"=D«D.tok_all»/{I('days_year')}"}, "對照：招股章程 2025-11 日均 4.2T（SRC_ZP_991，招股章程第 32 頁；r2 P2 改引用一手列；口徑可能含全部雲端與免費）", key="daily25")
    Dm.add("公司揭露：2025-11 日均 token", "T／日", {"D": f"={S('tok_daily_nov25_p')}"}, "Interested-party（SRC_ZP_991，招股章程月度點）", key="daily_src")
    Dm.add("2025 計費比例＝API 按量計費 token ÷（日均 4.2T × 365）", "比例", {"D": f"=D«D.api_tok»/({S('tok_daily_nov25_p')}*{I('days_year')})"},
           "類比 OpenAI v0.6 DEM_ApiBilledRatio2025：揭露量多為免費、Coding Plan 與折扣流量", key="billed", name="DEM_ApiBilledRatio2025", name_col="D")

    # ═════════════════════ Revenue ═════════════════════
    V = Sheet(wb, "Revenue", "R",
              "Revenue — 營收：①開放平台及 API（按量計費＋Coding Plan 子列）②企業級智能體 ③企業級通用大模型 ④技術服務及其他；總額＝淨額；雲端 vs 本地化、企業 vs 個人",
              "API 牌價：2025＝GLM-4.7／GLM-4.5-Air 牌價；1H26＝GLM-4.7→GLM-5→GLM-5.1 依發布日天數加權；2H26＝GLM-5.2／5.3（Sol）、GLM-4.5-Air→GLM-5.3-Flash（Luna）；"
              "2027 起 ×（1＋年變動率）。輸入價依長上下文占比混合兩檔。有效單價＝［(1−輸出占比)×(快取命中×快取價＋(1−快取命中)×輸入價)＋輸出占比×輸出價］×(1−折扣與免費額度)。",
              "金額 RMB 億；牌價 元／百萬 token；Coding Plan 價格 元／月。②③④＝本地化部署分部（不耗智譜算力，D4、D2r）。容量上限係數取自 Compute（Z3），只回乘雲端營收。")
    V.years(D)
    s_long = I("long_ctx_share")
    V.section("一、價格期間權重（由發布日推得的天數；Inputs Derived）")
    V.add("1H26 GLM-4.7 期間占比", "比例", {"K": f"={I('days_era47_1h26')}/{I('days_1h26')}"}, "2026-01-01～02-11 ÷ 181 天", key="w47")
    V.add("1H26 GLM-5 期間占比", "比例", {"K": f"={I('days_era5_1h26')}/{I('days_1h26')}"}, "2026-02-12～04-06 ÷ 181 天", key="w5")
    V.add("1H26 GLM-5.1 期間占比（＝1 − 前兩者）", "比例", {"K": "=(1-K«V.w47»-K«V.w5»)"}, "2026-04-07～06-30（GLM-5.2 於 06-16 發布，牌價同 GLM-5.1 ≥32K 檔）", key="w51")
    V.add("2H26 GLM-5.3-Flash 期間占比（Luna）", "比例", {"L": f"={I('days_flash_2h26')}/{I('days_2h26')}"}, "2026-08-26～12-31 ÷ 184 天；之前 Luna 以 GLM-4.5-Air 計", key="wfl")

    V.section("二、API 牌價（元／百萬 token；層級 × 輸入／快取／輸出）")
    chg = {"sol": I("px_chg_sol"), "luna": I("px_chg_luna")}

    def blend(a, b, io):
        return f"(1-{s_long})*{S(f'{a}_{io}')}+{s_long}*{S(f'{b}_{io}')}"
    for t in TIERS:
        for io in IO:
            key = f"lp_{t}_{io}"

            def lp(c, t=t, io=io, key=key):
                if t == "sol":
                    if c == "D":
                        return "=" + blend("px_rmb_glm47_long_out", "px_rmb_glm47_32k200k", io)
                    if c == "K":
                        return (f"=K«V.w47»*D«V.{key}»+K«V.w5»*({blend('px_rmb_glm5_lt32k', 'px_rmb_glm5_ge32k', io)})"
                                f"+K«V.w51»*({blend('px_rmb_glm51_lt32k', 'px_rmb_glm51_ge32k', io)})")
                    if c == "L":
                        return f"={S(f'px_rmb_glm53_{io}')}"
                else:
                    if c == "D":
                        return "=" + blend("px_rmb_glm45air_long_out", "px_rmb_glm45air_32k128k", io)
                    if c == "K":
                        return f"=D«V.{key}»"
                    if c == "L":
                        return f"=(1-L«V.wfl»)*K«V.{key}»+L«V.wfl»*{S(f'px_rmb_glm53flash_{io}')}"
                if c == "E":
                    return f"=(K«V.{key}»+L«V.{key}»)/{halves}"
                return f"={'L' if c == 'F' else PREV[c]}«V.{key}»*(1+{chg[t]})"
            note = ("2025：GLM-4.7（<32K 且輸出 ≥0.2K 檔 與 32K–200K 檔，依長上下文占比混合）；1H26：期間加權 GLM-4.7／GLM-5／GLM-5.1；2H26：GLM-5.2／5.3 1M 全量；2027 起 ×（1＋年變動率）；2026 欄為兩個半年的平均（顯示用）"
                    if t == "sol" else
                    "2025、1H26：GLM-4.5-Air（<32K 且輸出 ≥0.2K 檔 與 32K–128K 檔混合）；2H26：GLM-4.5-Air 與 GLM-5.3-Flash（08-26 起）期間加權；2027 起 ×（1＋年變動率）")
            V.add(f"{TIER_ZH[t]} {IO_ZH[io]}牌價", "元／M", lp, note, key=key, name=f"REV_ListPrice_{t.capitalize()}_{io.capitalize()}")

    V.section("三、有效單價（元／百萬 token；含快取、輸出占比、折扣與免費額度）")
    phi, h, dsc = I("out_share"), I("cache_hit"), I("discount_free")
    for t in TIERS:
        V.add(f"{TIER_ZH[t]} 有效單價", "元／M",
              lambda c, t=t: f"=(K«V.ep_{t}»+L«V.ep_{t}»)/{halves}" if c == "E" else
              f"=((1-{phi})*({h}*{c}«V.lp_{t}_cache»+(1-{h})*{c}«V.lp_{t}_in»)+{phi}*{c}«V.lp_{t}_out»)*(1-{dsc})",
              "［(1−輸出占比)×(快取命中×快取價＋(1−快取命中)×輸入價)＋輸出占比×輸出價］×(1−折扣與免費額度)", key=f"ep_{t}", name=f"REV_EffPrice_{t.capitalize()}")
    solmix = {"D": I("sol_share_2025"), "K": I("sol_share_2026"), "L": I("sol_share_2026"), "E": I("sol_share_2026"), **{c: I("sol_share_2027") for c in LATE}}
    V.add("付費 API token 組合：Sol", "比例", lambda c: f"={solmix[c]}", "Inputs（2025、2026、2027 起）", key="mix_sol")
    V.add("付費 API token 組合：Luna", "比例", lambda c: f"=(1-{c}«V.mix_sol»)", "＝1 − Sol（D24：無 Astra 級）", key="mix_luna")
    V.add("組合有效單價", "元／M",
          lambda c: f"=E«V.paygo»*{divrev}/Demand!E«D.api_tok»" if c == "E" else f"={c}«V.mix_sol»*{c}«V.ep_sol»+{c}«V.mix_luna»*{c}«V.ep_luna»",
          "Σ 層級組合 × 有效單價；2026 欄＝全年按量計費營收 ÷ 全年 token（token 加權）", key="p", name="REV_ApiPrice", name_h="REV_ApiPrice_H")

    V.section("四、GLM Coding Plan 方案價格（元／月／訂閱者）")
    bill, intl = I("cp_bill_disc"), I("cp_intl_share")
    for p in PLANS:
        V.add(f"{PLAN_ZH[p]} 國內月付牌價", "元／月", {"K": f"={S(f'cp_rmb_{p}_month')}", "L": f"={S(f'cp_rmb_{p}_month')}"},
              "2026-02-12 漲價後（SRC_ZP_356／359／362）", key=f"cpdom_{p}")
        V.add(f"{PLAN_ZH[p]} 國際站月價（換人民幣）", "元／月",
              {"K": f"={S(f'cp_usd_{p}_qtr_feb2026')}/{I('months_qtr')}*{fx}", "L": f"={S(f'cp_usd_{p}_list')}*{fx}"},
              "1H26：2026-02～07 季付價 ÷ 3（SRC_ZP_373–375）；2H26：2026-07 起牌價（SRC_ZP_367／369／371）；× USD/CNY", key=f"cpint_{p}")

        def cpe(c, p=p):
            if c in ("K", "L"):
                return f"=((1-{intl})*{c}«V.cpdom_{p}»+{intl}*{c}«V.cpint_{p}»)*(1-{bill})"
            if c == "D":
                return f"=K«V.cpe_{p}»/(1+{S('pxev_cp_ar_feb2026')}/{pct})"
            if c == "E":
                return f"=(K«V.cpe_{p}»+L«V.cpe_{p}»)/{halves}"
            return f"={'L' if c == 'F' else PREV[c]}«V.cpe_{p}»*(1+{I('cp_px_chg')})"
        V.add(f"{PLAN_ZH[p]} 有效月價（實收）", "元／月", cpe,
              "［(1−國際占比)×國內＋國際占比×國際］×(1−季／年付折扣)；2025＝1H26 ÷（1＋2026-02 漲價 30%，SRC_ZP_350）；2027 起 ×（1＋年變動率）", key=f"cpe_{p}")
    V.add("方案組合：Lite", "比例", {c: f"={I('cp_mix_lite')}" for c in ALL}, "Inputs", key="mix_lite")
    V.add("方案組合：Pro", "比例", {c: f"={I('cp_mix_pro')}" for c in ALL}, "Inputs", key="mix_pro")
    V.add("方案組合：Max（＝1 − Lite − Pro）", "比例", {c: f"=(1-{c}«V.mix_lite»-{c}«V.mix_pro»)" for c in ALL}, "不另輸入", key="mix_max")
    V.add("每訂閱者加權有效月價", "元／月", lambda c: "=" + "+".join(f"{c}«V.mix_{p}»*{c}«V.cpe_{p}»" for p in PLANS), "", key="cp_arpu", name="REV_CP_ARPU")

    V.section("五、營收（RMB 億）")
    for p in PLANS:
        V.add(f"①-Coding Plan 子列：{PLAN_ZH[p]}", "RMB 億",
              lambda c, p=p: "=K«V.cp_%s»+L«V.cp_%s»" % (p, p) if c == "E" else
              f"=Demand!{c}«D.subs_{p}»*{c}«V.cpe_{p}»*{months[c]}/{I('div_wan_yuan_yi')}",
              "訂閱者（萬）× 有效月價 × 月數 ÷ 10,000（萬 × 元 → 億）", key=f"cp_{p}", name=f"REV_Sub_{PLAN_ZH[p]}")
    V.add("①-Coding Plan 子列合計（訂閱）", "RMB 億", lambda c: "=" + "+".join(f"{c}«V.cp_{p}»" for p in PLANS), "財報歸入開放平台及 API（D2r）；本模型以訂閱者驅動", key="cp", name="REV_Sub", name_h="REV_Sub_H")
    V.add("校準：按量計費營收＝開放平台及 API 實際 − Coding Plan 估計", "RMB 億",
          {"D": f"={S('rev_api_fy25')}-D«V.cp»", "K": f"={S('rev_api_1h26')}-K«V.cp»"},
          "2025（年報）、1H26（中期公告）；按量計費 token 由此倒推（D14）", key="calib")
    for t in TIERS:
        V.add(f"①-按量計費：{TIER_ZH[t]}", "RMB 億",
              lambda c, t=t: "=K«V.pg_%s»+L«V.pg_%s»" % (t, t) if c == "E" else f"=Demand!{c}«D.api_tok»*{c}«V.mix_{t}»*{c}«V.ep_{t}»/{divrev}",
              "token × 組合 × 有效單價 ÷ 100（T × 元/M → 億）", key=f"pg_{t}", name=f"REV_API_{t.capitalize()}")
    V.add("①-按量計費合計", "RMB 億", lambda c: f"={c}«V.pg_sol»+{c}«V.pg_luna»", "校準期恆等於上方校準列", key="paygo", name="REV_API_PayGo")
    V.add("①開放平台及 API（按量計費＋Coding Plan）", "RMB 億", lambda c: f"={c}«V.paygo»+{c}«V.cp»", "＝雲端部署分部（D2r）",
          key="api", name="REV_API", name_h="REV_API_H")
    g_on = I("onprem_g")
    k_on = I("onprem_k_2h26")
    for sg in SEGS:
        def seg(c, sg=sg):
            if c == "D":
                return f"={S(f'rev_{sg}_fy25')}"
            if c == "K":
                return f"={S(f'rev_{sg}_1h26')}"
            if c == "L":
                return (f"=({S(f'rev_{sg}_fy25')}-{S(f'rev_{sg}_1h25')})*((1-{k_on})+{k_on}*{S(f'rev_{sg}_1h26')}/{S(f'rev_{sg}_1h25')})")
            if c == "E":
                return f"=K«V.{sg}»+L«V.{sg}»"
            return f"={PREV[c]}«V.{sg}»*(1+{g_on})"
        V.add(SEG_ZH[sg], "RMB 億", seg,
              "2025、1H26＝財報；2H26＝2H25（FY25 − 1H25）×［(1−延續係數)＋延續係數 × 1H26 ÷ 1H25］（工作單：依 1H26 實際、各線分開）；2027 起 ×（1＋年增率）",
              key=sg, name={"agent": "REV_OnPrem_Agent", "gpllm": "REV_OnPrem_GPLLM", "techsvc": "REV_Other"}[sg])
    V.add("本地化部署分部（②＋③＋④；不耗智譜算力）", "RMB 億", lambda c: "=" + "+".join(f"{c}«V.{sg}»" for sg in SEGS), "D4：不進 Demand、不受容量上限截頂",
          key="onprem", name="REV_OnPrem", name_h="REV_OnPrem_H")
    V.add("廣告", "RMB 億", {c: f"={I('ads_rev')}" for c in ALL}, "規格 D3：無廣告業務（Decision＝0）", key="ads", name="REV_Ads")
    V.add("營收總額（①＋本地化＋廣告）", "RMB 億", lambda c: f"={c}«V.api»+{c}«V.onprem»+{c}«V.ads»", "未截頂", key="gross", name="REV_Gross", name_h="REV_Gross_H")
    V.add("營收淨額（＝總額；無雲端夥伴分成，D5）", "RMB 億", lambda c: f"={c}«V.gross»", "REV_PartnerShare 不建", key="net", name="REV_Net")
    V.add("雲端營收（①；跑在智譜算力上）", "RMB 億", lambda c: f"={c}«V.api»", "規格 r1：②③④屬本地化部署分部；C 端收入在④內", key="cloud", name="REV_Cloud")
    V.add("容量上限係數（Compute）", "比例",
          {**{c: f"=Compute!{c}«C.cap»" for c in YC if c != "E"}, "E": "=(K«V.cloud»+L«V.cloud»*Compute!L«C.cap»)/E«V.cloud»"},
          "Compute CMP_CapFactor（1＝未受限；Z3 取代 Z2 的佔位 1）；2026＝1H26 實際不截頂、2H26 依 Compute 係數，以雲端營收加權", key="cap", name="REV_CapFactor")
    V.add("營收總額（截頂後）＝雲端 × 係數＋本地化＋廣告", "RMB 億", {c: f"={c}«V.cloud»*{c}«V.cap»+{c}«V.onprem»+{c}«V.ads»" for c in YC},
          "容量上限只回乘雲端營收（工作單 Z2 第 5 步）", key="gross_c", name="REV_GrossCapped")
    V.add("營收淨額（截頂後）", "RMB 億", {c: f"={c}«V.gross_c»" for c in YC}, "", key="net_c", name="REV_NetCapped")
    V.add("企業（按量計費＋本地化部署）", "RMB 億", lambda c: f"={c}«V.paygo»+{c}«V.onprem»", "D2：企業 vs 個人；按量計費以企業與開發者為主，歸企業", key="ent", name="REV_Enterprise")
    V.add("個人（Coding Plan 訂閱）", "RMB 億", lambda c: f"={c}«V.cp»", "Coding Plan 以個人開發者為主；C 端清言收入在④內未拆，歸企業列（資料缺口）", key="ind", name="REV_Individual")
    V.add("檢查：企業＋個人＋廣告 − 總額（應為 0）", "RMB 億", lambda c: f"={c}«V.ent»+{c}«V.ind»+{c}«V.ads»-{c}«V.gross»", "", key="sumchk")
    V.add("雲端占總額", "比例", lambda c: f"={c}«V.cloud»/{c}«V.gross»", "2025 約 26%、1H26 約 87%（財報）", key="cloudshare")
    V.add("總額年增率", "比例", {c: f"=({c}«V.gross»-{YC[YC.index(c) - 1]}«V.gross»)/{YC[YC.index(c) - 1]}«V.gross»" for c in YC[1:]}, "", key="growth")

    V.section("六、校準差距（應為 0）與對照（公司說法、TK、未校準值；只列差距，不反推）")
    segkey = {"api": "rev_api", "agent": "rev_agent", "gpllm": "rev_gpllm", "techsvc": "rev_techsvc"}
    for k, sk in segkey.items():
        V.add(f"2025 校準差距：{'①開放平台及 API' if k == 'api' else SEG_ZH[k]}（模型 − 年報）", "RMB 億", {"D": f"=D«V.{k}»-{S(sk + '_fy25')}"}, "", key=f"g25_{k}")
        V.add(f"1H26 校準差距：{'①開放平台及 API' if k == 'api' else SEG_ZH[k]}（模型 − 中期公告）", "RMB 億", {"K": f"=K«V.{k}»-{S(sk + '_1h26')}"}, "", key=f"g1h_{k}")
    V.add("2025 校準差距：營收總額（模型 − 年報）", "RMB 億", {"D": f"=D«V.gross»-{S('rev_total_fy25')}"}, "", key="g25", name="REV_Gap2025", name_col="D")
    V.add("1H26 校準差距：營收總額（模型 − 中期公告）", "RMB 億", {"K": f"=K«V.gross»-{S('rev_total_1h26')}"},
          "公告以千元四捨五入至億元四位小數：四線合計 9.53892 vs 總額 9.5389", key="g1h", name="REV_Gap1H26", name_col="K")
    V.add("2025 未校準 ①：日均 4.2T × 365 × 組合有效單價 ÷ 100 ＋ Coding Plan", "RMB 億",
          {"D": f"={S('tok_daily_nov25_p')}*{I('days_year')}*D«V.p»/{divrev}+D«V.cp»"}, "自下而上口徑（假設揭露 token 全數按有效單價計費）", key="unc25")
    V.add("2025 未校準差距：未校準 − 年報 ①", "RMB 億", {"D": f"=D«V.unc25»-{S('rev_api_fy25')}"}, "差距大＝揭露量多為免費、Coding Plan 或年末時點", key="uncgap25", name="REV_GapUncal2025", name_col="D")
    V.add("模型組合有效單價變動：1H26 對 2025", "比例", {"K": "=(K«V.p»-D«V.p»)/D«V.p»"}, "對照：公司稱 API 平均售價較年初 +101%（SRC_ZP_347）", key="asp_1h")
    V.add("模型組合有效單價變動：2H26 對 2025", "比例", {"L": "=(L«V.p»-D«V.p»)/D«V.p»"}, "對照：較 2025 年底 +83%（SRC_ZP_346，截至 2026-03）", key="asp_2h")
    V.add("公司說法：API 平均售價較年初上升（中期）", "比例", {"K": f"={S('pxev_api_asp_1h26')}/{pct}"}, "Interested-party；實現平均售價口徑（收入 ÷ 計費 token），只對照", key="asp_claim")
    V.add("公司說法：API 呼叫定價較 2025 年底上升", "比例", {"L": f"={S('pxev_api_price_vs_ye2025')}/{pct}"}, "Interested-party；只對照", key="asp_claim83")
    V.add("模型 2H26 雲端營收年化（2H26 × 2）", "RMB 億", {"L": f"=L«V.cloud»*{halves}"}, "對照 MaaS ARR（D14r）", key="arr_model")
    V.add("公司說法：MaaS ARR（2026-08 月度年化）× 匯率", "RMB 億", {"L": f"={S('arr_maas_aug26_monthly')}*{fx}"}, "SRC_ZP_417（USD 億）；Interested-party", key="arr_src")
    V.add("差距：模型 ÷ ARR − 1", "比例", {"L": "=(L«V.arr_model»-L«V.arr_src»)/L«V.arr_src»"}, "Checks WARN 門檻 25%（D14r）", key="arr_gap", name="REV_ARRGapMaaS", name_col="L")
    V.add("模型 2H26 營收總額年化（2H26 × 2）", "RMB 億", {"L": f"=L«V.gross»*{halves}"}, "對照全業務 ARR", key="arrt_model")
    V.add("公司說法：全業務 ARR（2026-09）× 匯率", "RMB 億", {"L": f"={S('arr_total_sep26')}*{fx}"}, "SRC_ZP_419（USD 億）；Interested-party", key="arrt_src")
    V.add("差距：模型 ÷ 全業務 ARR − 1", "比例", {"L": "=(L«V.arrt_model»-L«V.arrt_src»)/L«V.arrt_src»"}, "Checks WARN 門檻 25%", key="arrt_gap", name="REV_ARRGapTotal", name_col="L")
    V.add("分析師共識營收 FY2026（Simply Wall St）", "RMB 億", {"E": f"={S('cons_rev_fy26')}/{I('div_mn_yi')}"}, "SRC_ZP_582；Interested-party（券商交易誘因）；只對照", key="cons26")
    V.add("差距：模型 2026 總額 − 共識", "RMB 億", {"E": "=E«V.gross»-E«V.cons26»"}, "", key="cons_gap")
    tk = tk_cells
    V.add("TK：GLM-5.3 國際站輸入價 × 匯率 ÷ 本模型 2H26 Sol 人民幣輸入牌價", "倍", {"L": f"={tk['RD_CapIn_GLM53'][0]}*{fx}/L«V.lp_sol_in»"},
          "Tokenomics Cap_In F 表（讀表）；國際站美元價換人民幣後對國內牌價", key="tk_in")
    V.add("TK：GLM-5.3 國際站輸出價 × 匯率 ÷ 本模型 2H26 Sol 人民幣輸出牌價", "倍", {"L": f"={tk['RD_CapIn_GLM53'][2]}*{fx}/L«V.lp_sol_out»"}, "同上", key="tk_out")
    V.add("TK：中國合格前緣（Luna）混合單價 × 匯率", "元／M", {"L": f"={tk['RD_PF_ChinaFront'][0]}*{fx}"},
          "Tokenomics Price_Frontier B 表（讀表；前緣模型 GLM-5.3-Flash）；Tokenomics 參考請求混合口徑，未扣折扣", key="tk_front")
    V.add("TK：GLM-5.3 參考請求混合單價 × 匯率", "元／M", {"L": f"={tk['RD_PF_GLM53_Mix'][1]}*{fx}"}, "Tokenomics Price_Frontier A 表（讀表，Sol 參考請求）", key="tk_mix")
    V.add("本模型 2H26 有效單價 ÷ TK 混合單價（Sol）", "倍", {"L": "=L«V.ep_sol»/L«V.tk_mix»"}, "口徑：本模型含快取、輸出占比與折扣；只列差距", key="tk_ratio_sol")
    V.add("本模型 2H26 有效單價 ÷ TK 中國合格前緣（Luna）", "倍", {"L": "=L«V.ep_luna»/L«V.tk_front»"}, "同上", key="tk_ratio_luna")

    for ws, w in ((Dm.ws, 54), (V.ws, 60)):
        ws.column_dimensions["A"].width = 7
        ws.column_dimensions["B"].width = w
        ws.column_dimensions["C"].width = 12
        for c in YC + HC:
            ws.column_dimensions[c].width = 11
        ws.column_dimensions["J"].width = 2
        ws.column_dimensions["M"].width = 80
        ws.freeze_panes = "D7"
    return dict(D=Dm, V=V)


def checks(D, Z, summary, snap, e6, oref):
    """Z2 的 Checks 列：(label, 公式, 期望, 種類, 說明)。«D.»／«V.» 佔位由 fill() 代換。"""
    S, I = D.S, D.I
    n_src, n_inp = summary["src"], summary["inp"]
    n_empty = sum(1 for r in D.src if all(r[k] is None for k in ("value", "lo", "hi")))
    rngD = lambda key: f"Demand!$D${{r}}:$I${{r}}".replace("{r}", f"«D.{key}»")  # noqa: E731
    seg25 = "+".join(f"ABS(Revenue!$D$«V.g25_{k}»)" for k in ("api", "agent", "gpllm", "techsvc"))
    seg1h = "+".join(f"ABS(Revenue!$K$«V.g1h_{k}»)" for k in ("api", "agent", "gpllm", "techsvc"))
    cols8 = ("D", "E", "F", "G", "H", "I", "K", "L")
    mixv = "+".join(f"(Revenue!{c}«V.mix_max»<0)+(Revenue!{c}«V.mix_max»>1)" for c in cols8)
    solv = "+".join(f"(Revenue!{c}«V.mix_sol»<0)+(Revenue!{c}«V.mix_sol»>1)" for c in cols8)
    segsum = "+".join(f"ABS(Revenue!{c}«V.api»+Revenue!{c}«V.onprem»+Revenue!{c}«V.ads»-Revenue!{c}«V.gross»)" for c in cols8)
    entsum = "+".join(f"ABS(Revenue!{c}«V.sumchk»)" for c in cols8)
    rows = [
        # ── 資料表
        ("SRC_ZP 列數", f"=SUMPRODUCT(--(LEN({preserve_src()}!$A$5:$A$2000)>0))", n_src, "eq", "Z1 604 筆（SRC_ZP_001–604）＋ Z4 補 6 筆＋招股章程 Z1-P 444 筆（SRC_ZP_611–1054）"),
        ("SRC_ZP 無數值列數（日期、事件、找不到數值的說明列）", f"=SUM({preserve_src()}!$U$5:$U$2000)", n_empty, "eq", "值、低、高皆空的列數；與 yaml 一致（列被誤刪或誤填時轉 ERR）"),
        ("SRC_ZP 缺出處列數", f"=SUM({preserve_src()}!$V$5:$V$2000)", 0, "eq", ""),
        ("SRC_ZP 區間順序異常列數（低 ≤ 值 ≤ 高）", f"=SUM({preserve_src()}!$W$5:$W$2000)", 0, "eq", ""),
        ("Inputs 列數", "=SUMPRODUCT(--(LEN(Inputs!$A$5:$A$600)>0))", n_inp, "eq", "data/zhipu_inputs.yaml"),
        ("Inputs 區間順序異常列數", "=SUM(Inputs!$L$5:$L$600)", 0, "eq", "低 ≤ 值 ≤ 高"),
        ("Inputs Analogy／Assumed 缺區間列數", "=SUM(Inputs!$M$5:$M$600)", 0, "eq", "共同規則第 4 節：Analogy／Assumed 一律給區間"),
        ("TK_Link 具名名稱數（同 OpenAI v0.6 的 63 名）", '=COUNTIF(TK_Link!$F$10:$F$200,"OK")', summary["tk_ok"], "eq", "取自 Tokenomics master CURRENT"),
        ("TK_Link 讀表列數（Cap_In／Price_Frontier 的 GLM 列與中國合格前緣）", f'=COUNTIF(TK_Link!$F$10:$F$200,"{_table_note()}")', summary["tk_table"], "eq",
         "以列標籤讀表；Tokenomics 改版找不到標籤時狀態轉 MISSING"),
        ("TK_Link 錯誤值格數", "=SUMPRODUCT(--ISERROR(TK_Link!$J$10:$X$200))", 0, "eq", ""),
        ("TK 快照版本", "=TK_Version", snap["version"], "eq", f"共同規則第 8 節：master {snap['sha'][:7]}；tools/check_tk_snapshot.py 結果見報告"),
        ("OAI_Link 已取值列數", '=COUNTIF(OAI_Link!$E$10:$E$60,"OK")', summary["oai"], "eq", "OpenAI v0.6 命題輸出（美元）"),
        ("OAI_Link 被計算頁引用的公式格數（builder 掃描）", oref, 0, "eq",
         "規格 D26：OAI_Link 只被 Checks 與 HTML 引用；值為 builder 建置時掃描結果，tests/parity/test_builder.py 每次逐格重驗"),
        ("E6：計算頁公式含常數的格數（builder 掃描）", e6, 0, "eq", "Demand、Revenue（Z3 起含 Compute、Cost）；tests/parity/test_builder.py 每次逐格重驗"),
        # ── 校準與恆等式
        ("2025 各分部校準差距（四線絕對值合計，RMB 億）", "=" + seg25, 0, "tol", "①API（含 Coding Plan）②③④ 模型 − 年報；按量計費為倒推項，恆為 0"),
        ("1H26 各分部校準差距（四線絕對值合計，RMB 億）", "=" + seg1h, 0, "tol", "模型 − 中期公告"),
        ("2025 營收總額校準差距（RMB 億）", "=REV_Gap2025", 0, "tol", "四線合計＝年報總額 7.2433"),
        ("1H26 營收總額校準差距（RMB 億；公告四捨五入容差 0.0005）", "=REV_Gap1H26", 0, "tolr", "四線合計 9.53892（千元加總）vs 總額 9.5389"),
        ("分部合計＝總營收（API＋本地化＋廣告 − 總額；各期絕對值合計）", "=" + segsum, 0, "tol", "2025–2030 與 1H26、2H26"),
        ("企業＋個人＋廣告 − 總額（各期絕對值合計）", "=" + entsum, 0, "tol", "D2 彙總列"),
        ("校準期按量計費營收 ≤0 的期數（Coding Plan 估計超過 API 實際）", "=(Revenue!D«V.calib»<=0)+(Revenue!K«V.calib»<=0)", 0, "eq",
         "Coding Plan 訂閱者 × 價格的 Inputs 使估計值超過財報 API 營收時轉 ERR"),
        ("token 層級拆分前後總量不一致的年數（相對差 >1e-9）",
         "=SUMPRODUCT(--(ABS(" + "+".join(rngD(k) for k in ("api_tok", "cptok", "qy_tok", "api_free")) + "-" + rngD("tok_all") + ")>0.000000001*" + rngD("tok_all") + "))",
         0, "eq", "API＋Coding Plan＋清言＋免費 API ＝ 付費＋免費（Sol／Luna 拆分前後）"),
        ("Coding Plan 方案組合 Max（＝1 − Lite − Pro）∉[0,1] 的期數", "=" + mixv, 0, "eq", "Inputs Lite＋Pro ≤ 1"),
        ("API 層級組合 Sol ∉[0,1] 的期數", "=" + solv, 0, "eq", ""),
        ("1H26 價格期間權重 GLM-5.1 ＜0（天數超過半年）", "=IF(Revenue!K«V.w51»<0,1,0)", 0, "eq", "Inputs 期間天數"),
        ("本地化部署延續係數 ∉[0,1]", f"=IF(OR({I('onprem_k_2h26')}<0,{I('onprem_k_2h26')}>1),1,0)", 0, "eq", "工作單：0＝下半年持平、1＝延續上半年同比"),
        ("API 任務數或每任務 token ≤0 的期數", "=SUMPRODUCT(--(" + rngD("tasks") + "<=0))+SUMPRODUCT(--(" + rngD("tpt") + "<=0))", 0, "eq", ""),
        # ── 公司說法對照（WARN／INFO）
        ("D14r：模型 2H26 年化雲端營收 ÷ MaaS ARR（2026-08，US$16 億）− 1", "=REV_ARRGapMaaS", None, "warn", "絕對值 >25% 為 WARN；ARR 只對照、不反推參數；原因見報告"),
        ("D14r：模型 2H26 年化總營收 ÷ 全業務 ARR（2026-09，US$18 億）− 1", "=REV_ARRGapTotal", None, "warn", "同上"),
        ("對照：模型組合有效單價 1H26 對 2025 變動", "=Revenue!K«V.asp_1h»", None, "info", "公司稱平均售價較年初 +101%（實現售價口徑）；報告討論"),
        ("對照：模型組合有效單價 2H26 對 2025 變動", "=Revenue!L«V.asp_2h»", None, "info", "公司稱較 2025 年底 +83%（截至 2026-03）"),
        ("對照：模型 2H26 ÷ 1H26 API token", "=Demand!L«D.vol_ratio»", None, "info", "公司稱 MaaS 呼叫量較年初 40 倍以上（時點對時點，口徑不同）"),
        ("對照：2H25→1H26 實際隱含量成長（倍）", "=DEM_ImpliedVolGrowth", None, "info", "Inputs 2H26 任務數成長（R1：延續此實際動能 2.24）的依據"),
        ("對照：2025 未校準差距（RMB 億）", "=REV_GapUncal2025", None, "info", "日均 4.2T 全數按有效單價計的假想營收 − 年報"),
        ("對照：2025 計費比例（按量計費 token ÷ 揭露日均 × 365）", "=DEM_ApiBilledRatio2025", None, "info", ""),
        ("對照：模型 2026 總額 − 分析師共識（RMB 億）", "=Revenue!E«V.cons_gap»", None, "info", "共識 60.48 億（Interested-party）"),
        ("對照：本模型 2H26 Sol 有效單價 ÷ TK GLM-5.3 混合單價 × 匯率", "=Revenue!L«V.tk_ratio_sol»", None, "info", "TK_Link 讀表列"),
        ("對照：本模型 2H26 Luna 有效單價 ÷ TK 中國合格前緣 × 匯率", "=Revenue!L«V.tk_ratio_luna»", None, "info", "TK_Link 讀表列"),
        ("對照：OpenAI v0.6 2030 每 VR 等值 GW 差額（$B/GW）", "=INDEX(OAI_COST_PropGap_VR,1,6)", None, "info", "OAI_Link；只並排（Z5 HTML）"),
        ("預覽：2030 營收總額（截頂後，RMB 億）", "=INDEX(REV_GrossCapped,1,6)", None, "info", ""),
        ("預覽：2030 token 合計（T）", "=INDEX(DEM_Tok_Total,1,6)", None, "info", "供 Z3"),
    ]
    return rows


def preserve_src():
    import preserve
    return preserve.SRC_SHEET


def _table_note():
    from tk_link import TABLE_NOTE
    return TABLE_NOTE
