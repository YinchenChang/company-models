"""v0.6-P2（工作單 S2）：需求頁 Demand 與營收頁 Revenue。

機制沿用 v0.5（reference/v0.5 的 JSON 決定 D、N、S 與 Excel「訂閱與廣告」「API」「營收彙總」頁的公式結構）；
數值一律引用 SRC_OAI／Inputs／Derived_V9／TK_ 具名範圍（E6：公式不內含常數；恆等式 1−x、1＋成長率 除外）。
Python 只產生結構：每列的公式文字由此組裝，計算由 Excel（LibreOffice）與 engine 執行。
「v0.5 原值」欄為 builder 由 v0.5 Excel 快取值讀出的對照常數，不參與計算。
"""
from __future__ import annotations

from pathlib import Path

import openpyxl
from openpyxl.workbook.defined_name import DefinedName

from common import F_BOLD, F_CALC, F_NOTE, FILL_SEC, header, put, title

YEARS = (2025, 2026, 2027, 2028, 2029, 2030)
YC = "DEFGHI"                      # 年度欄 D–I
V05C = "KLMNOP"                    # v0.5 原值欄 K–P
V05_XLSX = Path(__file__).resolve().parents[1] / "reference" / "v0.5" / "20260929_OpenAI收支模型_討論版_v0_5.xlsx"
PLANS = ("free", "go", "plus", "pro", "seats")
PLAN_ZH = {"free": "Free", "go": "Go", "plus": "Plus", "pro": "Pro", "seats": "企業席位"}
TIERS = ("top", "mid", "low")
TIER_ZH = {"top": "頂層（Astra）", "mid": "中層（Sol）", "low": "低層（Luna）"}
USAGE = {"top": "astra 占比", "mid": "sol 占比", "low": "luna 占比"}   # Inputs「方案模型組合」的索引


def nm(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


class V05:
    """v0.5 Excel 的快取值（只供對照欄）。"""

    def __init__(self):
        self.wb = openpyxl.load_workbook(V05_XLSX, data_only=True)

    def row(self, sheet, r):
        return [self.wb[sheet].cell(r, c).value for c in range(3, 9)]


class Sheet:
    """年度表：A 編號、B 項目、C 單位、D–I 2025–2030、K–P v0.5 原值、Q 說明。"""

    def __init__(self, wb, name, prefix, t1, t2, t3=None):
        self.wb, self.name, self.prefix = wb, name, prefix
        self.ws = wb.create_sheet(name)
        title(self.ws, t1, t2, t3)
        header(self.ws, 5, ["編號", "項目", "單位"] + [str(y) for y in YEARS] + [""] + [f"v0.5 {y}" for y in YEARS] + ["說明"])
        put(self.ws, "B6", "年度", F_BOLD)
        for i in range(6):
            put(self.ws, f"{YC[i]}6", f"=INP_{5 + i:03d}")
        self.r = 7
        self.n = 0
        self.rows = {}           # code → 列號
        self.meta = []           # (code, 項目, 單位, 說明, 名稱)

    def section(self, label):
        put(self.ws, f"A{self.r}", label, F_BOLD, fill=FILL_SEC)
        self.r += 1

    def add(self, label, unit, fs, v05=None, note="", name=None, key=None, name_col=None):
        """fs：長度 6 的公式清單（None＝該年空白），或 callable(i, col) → 公式。"""
        self.n += 1
        code = f"{self.prefix}{self.n:02d}"
        r = self.r
        self.rows[key or code] = r
        put(self.ws, f"A{r}", code, F_CALC)
        put(self.ws, f"B{r}", label, F_CALC)
        put(self.ws, f"C{r}", unit, F_CALC)
        for i in range(6):
            f = fs(i, YC[i]) if callable(fs) else fs[i]
            if f is not None:
                put(self.ws, f"{YC[i]}{r}", f)
        if v05 is not None:
            for i, v in enumerate(v05):
                if v is not None:
                    put(self.ws, f"{V05C[i]}{r}", v, F_NOTE)
        put(self.ws, f"Q{r}", note, F_NOTE)
        if name:
            nm(self.wb, name, f"{self.name}!${name_col}${r}" if name_col else f"{self.name}!$D${r}:$I${r}")
        self.rows[key or code] = r
        self.meta.append((code, label, unit, note, name))
        self.r += 1
        return r

    def ref(self, key, i, absolute=True):
        r = self.rows[key]
        return f"{self.name}!${YC[i]}${r}" if absolute else f"{YC[i]}{r}"


def build_p2(wb, R, sid, inp_id, inp_row):
    """建立 Demand、Revenue 兩頁；回傳供 Checks 使用的位置字典。"""
    v05 = V05()
    S = sid
    I = inp_id
    yr = [f"Demand!${YC[i]}$6" for i in range(6)]

    def yidx(name, y):
        return I(name, str(y))

    # ── 共用 Inputs ───────────────────────────────────────────
    days, months, mb, bt, tm = (I("每年天數"), I("每年月數"), I("單位換算：$M → $B"), I("單位換算：B token → T token"),
                                I("單位換算：T token → M token"))
    el_from = I("彈性起算年")

    # ═════════════════════ Demand ═════════════════════
    D = Sheet(wb, "Demand", "D", "Demand — 需求：訂閱各方案人數、每日任務數 × 每任務 token、API 任務數與計費 token；token 量（層級 × 付費／免費）",
              "機制沿用 v0.5「訂閱與廣告」「API」頁（D、N、S 決定）。2025 為校準年：API 計費 token 由 FY2025 實際營收倒推（v0.5 唯一倒推）。"
              "2027 起需求成長（不含價格）× 價格反應 (p/p前)^ε。免費用戶只產生 token 需求，不產生訂閱營收。",
              "token 單位 T＝10^12；人數 M＝百萬；任務數（API）十億。v0.5 原值欄（K–P）為 v0.5 Excel 快取值，只作對照。容量上限（V1a）於 P3 加入，本頁不截頂。")
    yr = [f"Demand!${YC[i]}$6" for i in range(6)]
    plan = {"go": [yidx("FY 平均用戶數：go", y) for y in YEARS],
            "plus": [None] + [yidx("FY 平均用戶數：plus", y) for y in YEARS[1:]],
            "seats": [yidx("FY 平均用戶數：seats", y) for y in YEARS]}
    frac = I("T4 遷移幅度（占公司計畫）— 基準")
    pg, pp = S("公司 2026 計畫：Go 用戶數"), S("公司 2026 計畫：Plus 用戶數")
    go26, pl26 = plan["go"][1], plan["plus"][1]

    D.section("一、訂閱人數（M，FY 平均；T4：2026 起 Plus→Go 遷移與 Go 額外新增）")
    D.add("遷移率 μ（Plus→Go）", "比例", [None] + [f"={frac}*(1-{pp}/{pl26})"] * 5, v05.row("訂閱與廣告", 5),
          "計畫幅度（Inputs 基準 0.3）×（1 − 公司計畫 Plus ÷ 無遷移 Plus 2026）；2026 起各年同值（v0.5）", key="mu")
    D.add("Go 額外新增倍數 λ", "倍", [None] + [f"={frac}*({pg}-{pl26}*(1-{pp}/{pl26})-{go26})/{go26}"] * 5,
          v05.row("訂閱與廣告", 6), "＝計畫幅度 ×（公司計畫 Go −（無遷移 Plus 2026 − 公司計畫 Plus）− 無遷移 Go 2026）÷ 無遷移 Go 2026；"
          "計畫幅度＝1 時，2026 Go 恰等於公司計畫 112M（v0.5；與 v0.5 寫法代數相同）", key="lam")
    free = [yidx("FY 平均用戶數：free", 2025), "V9_Free2026"] + [yidx("FY 平均用戶數：free", y) for y in YEARS[2:]]
    D.add("Free 人數", "M", lambda i, c: f"={free[i]}", v05.row("訂閱與廣告", 7),
          "2025、2027+ Inputs；2026＝SRC WAU × 換算係數（Derived_V9 V01）", name="DEM_Users_Free", key="u_free")
    D.add("Go 人數", "M", lambda i, c: f"={plan['go'][0]}" if i == 0 else
          f"={plan['go'][i]}*(1+{D.ref('lam', i, False)})+{plan['plus'][i]}*{D.ref('mu', i, False)}",
          v05.row("訂閱與廣告", 8), "2026 起＝無遷移 Go ×（1＋λ）＋無遷移 Plus × μ", name="DEM_Users_Go", key="u_go")
    D.add("Plus 人數", "M", lambda i, c: "=V9_Plus2025" if i == 0 else f"={plan['plus'][i]}*(1-{D.ref('mu', i, False)})",
          v05.row("訂閱與廣告", 9), "2025＝Derived_V9 V02（SRC 35M × 換算係數）；2026 起＝無遷移 Plus ×（1 − μ）", name="DEM_Users_Plus", key="u_plus")
    D.add("Pro 人數", "M", lambda i, c: f"=V9_Pro{YEARS[i]}", v05.row("訂閱與廣告", 10),
          "V12／V15：占比 × 付費總數（Derived_V9 V03、V04、V10–V13；付費總數口徑＝無遷移基準，沿用 P1 已審公式）", name="DEM_Users_Pro", key="u_pro")
    D.add("企業席位人數", "M", lambda i, c: f"={plan['seats'][i]}", v05.row("訂閱與廣告", 11), "Inputs", name="DEM_Users_Seats", key="u_seats")
    D.add("付費人數合計（Go＋Plus＋Pro＋席位）", "M",
          lambda i, c: "=" + "+".join(D.ref(f"u_{p}", i, False) for p in PLANS[1:]), None, "", name="DEM_Users_Paid", key="u_paid")
    D.add("消費端付費人數（Go＋Plus＋Pro）", "M", lambda i, c: "=" + "+".join(D.ref(f"u_{p}", i, False) for p in ("go", "plus", "pro")),
          v05.row("訂閱與廣告", 12), "", key="u_cons")

    D.section("二、訂閱使用量：每日任務數 × 每任務 token（2025、2026 為 Inputs；2027 起年增）")
    g_task, g_tok = I("每日任務數年增率"), I("每任務 token 年增率")
    for p, vr in zip(PLANS, range(21, 26)):
        a, b = yidx(f"每日任務數：{p}", 2025), yidx(f"每日任務數：{p}", 2026)
        D.add(f"{PLAN_ZH[p]} 每日任務數", "任務/日",
              lambda i, c, a=a, b=b, p=p: f"={a}" if i == 0 else f"={b}" if i == 1 else f"={YC[i - 1]}{D.rows[f'task_{p}']}*(1+{g_task})",
              v05.row("訂閱與廣告", vr), "2027 起 ×（1＋每日任務數年增率）", key=f"task_{p}")
    for p, vr in zip(PLANS, range(26, 31)):
        a, b = yidx(f"每任務 token：{p}", 2025), yidx(f"每任務 token：{p}", 2026)
        D.add(f"{PLAN_ZH[p]} 每任務 token", "K tok/任務",
              lambda i, c, a=a, b=b, p=p: f"={a}" if i == 0 else f"={b}" if i == 1 else f"={YC[i - 1]}{D.rows[f'tpt_{p}']}*(1+{g_tok})",
              v05.row("訂閱與廣告", vr), "2027 起 ×（1＋每任務 token 年增率；D11 情境軸）", key=f"tpt_{p}")
    for p, vr in zip(PLANS, range(31, 36)):
        D.add(f"{PLAN_ZH[p]} token", "T",
              lambda i, c, p=p: f"={D.ref(f'u_{p}', i, False)}*{D.ref(f'task_{p}', i, False)}*{D.ref(f'tpt_{p}', i, False)}*{days}/{bt}",
              v05.row("訂閱與廣告", vr), "人數（M）× 每日任務 × 每任務 token（K）× 每年天數 ÷ 1000（B→T）", name=f"DEM_Tok_Plan_{p.capitalize()}", key=f"tok_{p}")
    D.add("訂閱＋Free token 合計", "T", lambda i, c: "=" + "+".join(D.ref(f"tok_{p}", i, False) for p in PLANS),
          v05.row("訂閱與廣告", 36), "workbook 口徑（v0.5 未校正）", key="tok_sub")

    # API 需求（價格取自 Revenue；Revenue 的列號先預留，於 Revenue 建好後回填）
    D.section("三、API 需求：任務數 × 每任務 token（2025 由營收倒推；2027 起價格反應）")
    pr = {}
    D.add("價格反應底數 p/p前（組合有效單價）", "倍", [None] + ["{P_RATIO}"] * 5, v05.row("API", 21),
          "＝Revenue 組合有效單價 ÷ 前一年；價格反應只在年度 ≥ 彈性起算年（Inputs 2027）生效：底數^(ε×(年度≥起算年))", key="api_b")
    eK, eN = I("價格彈性：每任務 token"), I("價格彈性：任務數")
    D.add("API 每任務 token", "K tok/任務",
          lambda i, c: f"={I('基準每任務 token')}" if i == 0 else
          f"={YC[i - 1]}{{R}}*(1+{yidx('每任務 token 成長率（不含價格）', YEARS[i])})*{c}{{B}}^({eK}*({c}$6>={el_from}))",
          v05.row("API", 22), "層級正規化（只影響任務數與每任務 token 的水準拆分，不影響營收）", key="api_k")
    D.add("API 任務數", "十億",
          lambda i, c: "={C}{TOK}/{C}{K}" if i == 0 else
          f"={YC[i - 1]}{{R}}*(1+{yidx('任務數成長率（不含價格）', YEARS[i])})*{c}{{B}}^({eN}*({c}$6>={el_from}))",
          v05.row("API", 23), "2025＝計費 token ÷ 每任務 token；之後 ×（1＋成長）× 價格反應", key="api_n")
    D.add("API 計費 token", "T", lambda i, c: "{CALIB}" if i == 0 else f"={c}{{N}}*{c}{{K}}", v05.row("API", 24),
          "2025：（FY2025 實際營收 − 訂閱 − 廣告 − 其他）÷ 組合有效單價（v0.5 校準，唯一倒推）；之後＝任務數 × 每任務 token",
          name="DEM_Tok_API", key="api_tok")
    for t, vr in zip(TIERS, (28, 30, 32)):
        D.add(f"API token：{TIER_ZH[t]}", "T", [f"={{TOK}}*{{MIX_{t}}}"] * 6, v05.row("API", vr), "計費 token × 層級占比（Revenue）",
              name=f"DEM_Tok_API_{t.capitalize()}", key=f"api_tok_{t}")
    # 回填自身列號
    for key in ("api_k", "api_n", "api_tok") + tuple(f"api_tok_{t}" for t in TIERS):
        r = D.rows[key]
        for i in range(6):
            c = D.ws[f"{YC[i]}{r}"]
            if isinstance(c.value, str):
                c.value = (c.value.replace("{R}", str(r)).replace("{B}", str(D.rows["api_b"])).replace("{N}", str(D.rows["api_n"]))
                           .replace("{K}", str(D.rows["api_k"])).replace("{TOK}", str(D.rows["api_tok"])).replace("{C}", YC[i]))

    D.section("四、輸出：token 量（層級 × 付費／免費；供 P3 推論 GW）")
    paid_plans = PLANS[1:]
    for t in TIERS:
        D.add(f"免費 token：{TIER_ZH[t]}", "T",
              lambda i, c, t=t: f"={D.ref('tok_free', i, False)}*{I('方案模型組合：free', USAGE[t])}",
              None, f"Free token × 方案模型組合（Inputs {USAGE[t]}；v0.5 inferenceCost.usage）", name=f"DEM_Tok_Free_{t.capitalize()}", key=f"free_{t}")
    for t in TIERS:
        D.add(f"付費 token：{TIER_ZH[t]}", "T",
              lambda i, c, t=t: "=" + "+".join(f"{D.ref(f'tok_{p}', i, False)}*{I(f'方案模型組合：{p}', USAGE[t])}" for p in paid_plans)
              + f"+{D.ref(f'api_tok_{t}', i, False)}",
              None, "Σ 付費方案 token × 方案模型組合 ＋ API 該層級 token", name=f"DEM_Tok_Paid_{t.capitalize()}", key=f"paid_{t}")
    D.add("免費 token 合計", "T", lambda i, c: "=" + "+".join(D.ref(f"free_{t}", i, False) for t in TIERS), None, "", name="DEM_Tok_Free", key="free_all")
    D.add("付費 token 合計", "T", lambda i, c: "=" + "+".join(D.ref(f"paid_{t}", i, False) for t in TIERS), None, "", name="DEM_Tok_Paid", key="paid_all")
    D.add("token 合計（付費＋免費）", "T", lambda i, c: f"={D.ref('free_all', i, False)}+{D.ref('paid_all', i, False)}", None, "",
          name="DEM_Tok_Total", key="tok_all")
    D.add("對照：訂閱＋免費＋API 計費 token（應＝合計）", "T", lambda i, c: f"={D.ref('tok_sub', i, False)}+{D.ref('api_tok', i, False)}", None,
          "層級拆分前的總量；與上一列之差應為 0（Checks）", key="tok_chk")

    D.section("五、2025 對照（只列差距，不校準）")
    D.add("Tokenomics IF_AllocDemand（需求 D 合計，2025）", "M tok/年", ["=TK_IF_AllocDemand"] + [None] * 5, None,
          "TK 快照；口徑：API＋ChatGPT；付費＋免費 token", key="tk_dem")
    D.add("本模型 2025 token 合計", "M tok/年", [f"={D.ref('tok_all', 0, False)}*{tm}"] + [None] * 5, None, "T × 10^6（Inputs 定義常數）", key="our_dem")
    D.add("差距：本模型 − Tokenomics", "M tok/年", [f"=D{{OUR}}-D{{TK}}"] + [None] * 5, None, "", name="DEM_TkGap2025", name_col="D", key="dem_gap")
    D.add("比值：本模型 ÷ Tokenomics", "倍", [f"=D{{OUR}}/D{{TK}}"] + [None] * 5, None, "", name="DEM_TkRatio2025", name_col="D", key="dem_ratio")
    D.add("API 處理 token（揭露推估，FY2025）", "T", ["=V9_ApiTok"] + [None] * 5, [2628] + [None] * 5,
          "Derived_V9 V22：Inputs 每分鐘 5B × 525,600 ÷ 1000", key="api_proc")
    D.add("API 計費比例＝計費 token ÷ 處理 token（2025）", "比例", [f"=D{D.rows['api_tok']}/D{{PROC}}"] + [None] * 5,
          v05.row("API", 41), "v0.5 檢查項；計費 token 為營收倒推值", name="DEM_ApiBilledRatio2025", name_col="D", key="billed")
    for key in ("dem_gap", "dem_ratio", "billed"):
        r = D.rows[key]
        D.ws[f"D{r}"].value = (D.ws[f"D{r}"].value.replace("{OUR}", str(D.rows["our_dem"])).replace("{TK}", str(D.rows["tk_dem"]))
                               .replace("{PROC}", str(D.rows["api_proc"])))

    # ═════════════════════ Revenue ═════════════════════
    V = Sheet(wb, "Revenue", "R", "Revenue — 營收：API 單價路徑、訂閱、API、廣告、其他、總額、Microsoft 分成、淨額（估值與命題用淨額）",
              "機制沿用 v0.5「API」「訂閱與廣告」「營收彙總」頁。API 單價：2025 Inputs（Analogy）；2026 由價格事件按天數加權（E8c，事件時點見本頁下方）；"
              "2027 起 ×（1＋牌價年變動率）。訂閱營收＝人數 × 實收 ARPU × 12。",
              "營收單位 $B；單價 $/M tokens。容量上限（V1a）於 P3 加入，本頁營收未截頂。v0.5 原值欄（K–P）只作對照。")
    phi, h, rr, disc = I("輸出 token 占比（API）"), I("快取命中率（API）"), S("快取輸入價占輸入價比"), I("加權折扣率")
    chg = {t: I(f"牌價年變動率：{t} 層") for t in TIERS}

    V.section("一、API 牌價（FY 平均，$/M tokens；N3a 層級角色口徑）")
    lp = {}
    for t, io, vr in (("top", "in", 5), ("top", "out", 6), ("mid", "in", 7), ("mid", "out", 8), ("low", "in", 9), ("low", "out", 10)):
        base25 = I(f"FY 平均牌價（{'輸入' if io == 'in' else '輸出'}）：{t} 層")
        key = f"lp_{t}_{io}"
        V.add(f"{TIER_ZH[t]} {'輸入' if io == 'in' else '輸出'}牌價", "$/M",
              lambda i, c, base25=base25, key=key, t=t: f"={base25}" if i == 0 else "{EV_" + key + "}" if i == 1 else
              f"={YC[i - 1]}{V.rows[key]}*(1+{chg[t]})",
              v05.row("API", vr), "2025 Inputs（Analogy）；2026＝事件時點天數加權（下方第六節）；2027 起 ×（1＋年變動率）",
              name=f"REV_ListPrice_{t.capitalize()}_{io.capitalize()}", key=key)
    V.section("二、有效單價（$/M；含快取、輸出占比；折扣於組合列扣除）")
    for t, vr in zip(TIERS, (12, 13, 14)):
        V.add(f"{TIER_ZH[t]}每 token 單價", "$/M",
              lambda i, c, t=t: f"=(1-{phi})*({h}*{rr}+(1-{h}))*{c}{V.rows[f'lp_{t}_in']}+{phi}*{c}{V.rows[f'lp_{t}_out']}",
              v05.row("API", vr), "（1−輸出占比）×（快取命中 × 快取價比 ＋（1−快取命中））× 輸入價 ＋ 輸出占比 × 輸出價", key=f"ep_{t}")
    mixin = {"top": [I("API token 層級占比：top 層", "陣列[0]")] + [I("API token 層級占比：top 層", "陣列[1]")] * 5,
             "mid": [I("API token 層級占比：mid 層", "陣列[0]")] + [I("API token 層級占比：mid 層", "陣列[1]")] * 5}
    V.add("API token 組合：頂層", "比例", lambda i, c: f"={mixin['top'][i]}", v05.row("API", 15), "Inputs（2025、2026；2027 起沿用 2026，v0.5 以最後一值延伸）", key="mix_top")
    V.add("API token 組合：中層", "比例", lambda i, c: f"={mixin['mid'][i]}", v05.row("API", 16), "同上", key="mix_mid")
    V.add("API token 組合：低層", "比例", lambda i, c: f"=(1-{c}{V.rows['mix_top']}-{c}{V.rows['mix_mid']})", v05.row("API", 17),
          "＝1 − 頂層 − 中層（v0.5：不另輸入）", key="mix_low")
    V.add("組合有效單價（扣折扣）", "$/M",
          lambda i, c: "=(" + "+".join(f"{c}{V.rows[f'mix_{t}']}*{c}{V.rows[f'ep_{t}']}" for t in TIERS) + f")*(1-{disc})",
          v05.row("API", 18), "Σ 層級占比 × 每 token 單價 ×（1 − 加權折扣）", name="REV_ApiPrice", key="mixprice")

    V.section("三、訂閱營收（$B＝人數 × 實收 ARPU × 12 ÷ 1000）")
    arpu = {"go": [I("ARPU（實收）：go")] * 6, "plus": [I("ARPU（實收）：plus", "2025")] + [I("ARPU（實收）：plus", "2026")] * 5,
            "pro": [I("ARPU（實收）：pro", "2025")] + [I("ARPU（實收）：pro", "2026")] * 5, "seats": [I("ARPU（實收）：seats")] * 6}
    for p, vr in zip(PLANS[1:], (14, 15, 16, 17)):
        V.add(f"{PLAN_ZH[p]} 訂閱營收", "$B", lambda i, c, p=p: f"=Demand!{c}{D.rows[f'u_{p}']}*{arpu[p][i]}*{months}/{mb}",
              v05.row("訂閱與廣告", vr), "ARPU：plus、pro 2025 一欄、2026 起一欄（v0.5 以最後一值延伸）；go、seats 全期一欄",
              name=f"REV_Sub_{p.capitalize()}", key=f"sub_{p}")
    V.add("訂閱合計", "$B", lambda i, c: "=" + "+".join(f"{c}{V.rows[f'sub_{p}']}" for p in PLANS[1:]), v05.row("訂閱與廣告", 19), "",
          name="REV_Sub", key="sub")

    V.section("四、廣告與其他（廣告為獨立、慢成長的一列，不進核心分析，N1b）")
    V.add("廣告曝光用戶（Free＋Go）", "M",
          lambda i, c: f"=(Demand!{c}{D.rows['u_free']}+Demand!{c}{D.rows['u_go']})*{yidx('廣告曝光占比', YEARS[i])}",
          v05.row("訂閱與廣告", 47), "曝光占比 Inputs（Andy 2026-09-29：2027→2030 不宜 8 倍）", key="ad_users")
    V.add("廣告營收", "$B", lambda i, c: f"={c}{V.rows['ad_users']}*{yidx('每曝光用戶年廣告收入', YEARS[i])}/{mb}",
          v05.row("訂閱與廣告", 48), "每曝光用戶年收入 Inputs（Analogy；上限參照 Meta ARPP）", name="REV_Ads", key="ads")
    V.add("其他營收（授權、一次性交易）", "$B", [f"={I('其他營收')}"] * 6, [0] * 6, "Inputs（v0.5 residual，維持 0）", name="REV_Other", key="other")

    V.section("五、API 營收、總額、Microsoft 分成與淨額")
    V.add("2025 校準：API 營收＝實際營收 − 訂閱 − 廣告 − 其他", "$B",
          [f"={S('營收：FY2025')}-D{V.rows['sub']}-D{V.rows['ads']}-D{V.rows['other']}"] + [None] * 5, None,
          "v0.5 calibration：2025 API 為唯一倒推項", key="calib")
    V.add("API 營收", "$B", lambda i, c: f"=Demand!{c}{D.rows['api_tok']}*{c}{V.rows['mixprice']}/{mb}", v05.row("API", 25),
          "計費 token × 組合有效單價 ÷ 1000", name="REV_API", key="api")
    for t, vr in zip(TIERS, (27, 29, 31)):
        V.add(f"API 營收：{TIER_ZH[t]}", "$B",
              lambda i, c, t=t: f"=Demand!{c}{D.rows['api_tok']}*{c}{V.rows[f'mix_{t}']}*{c}{V.rows[f'ep_{t}']}*(1-{disc})/{mb}",
              v05.row("API", vr), "", name=f"REV_API_{t.capitalize()}", key=f"api_{t}")
    V.add("總額（訂閱＋API＋廣告＋其他）", "$B",
          lambda i, c: f"={c}{V.rows['sub']}+{c}{V.rows['api']}+{c}{V.rows['ads']}+{c}{V.rows['other']}", v05.row("營收彙總", 22),
          "容量上限（V1a）於 P3 加入", name="REV_Gross", key="gross")
    rate, cap, thru, cfrom = S("Microsoft 營收分成率"), S("Microsoft 分成總額上限"), S("Microsoft 分成付款持續至"), I("分成上限起算年")
    gr, msr, cumr = V.rows["gross"], V.r, V.r + 1

    def ms(i, c):
        pre = f"{rate}*{c}{gr}*({c}$6<{cfrom})"
        if i == 0:
            return f"={pre}+({c}$6>={cfrom})*({c}$6<={thru})*MIN({rate}*{c}{gr},{cap})"
        return f"={pre}+({c}$6>={cfrom})*({c}$6<={thru})*MIN({rate}*{c}{gr},{cap}-{YC[i - 1]}{cumr})"

    V.add("Microsoft 分成", "$B", ms, v05.row("營收彙總", 23),
          "起算年前：比率 × 總額；起算年至付款終年：MIN（比率 × 總額，上限 − 前一年累計）；之後 0（v0.5；以邏輯值相乘表示各期間）", name="REV_MSShare", key="ms")
    V.add("Microsoft 分成累計（自起算年）", "$B",
          lambda i, c: f"={c}{msr}*({c}$6>={cfrom})" if i == 0 else f"={YC[i - 1]}{cumr}+{c}{msr}*({c}$6>={cfrom})",
          v05.row("營收彙總", 24), "上限累計自 Inputs 起算年（N2a）", name="REV_MSCum", key="mscum")
    V.add("淨額（總額 − Microsoft 分成）", "$B", lambda i, c: f"={c}{gr}-{c}{msr}", v05.row("營收彙總", 25), "估值與命題用淨額", name="REV_Net", key="net")
    V.add("消費端（Go＋Plus＋Pro＋廣告）", "$B",
          lambda i, c: f"={c}{V.rows['sub_go']}+{c}{V.rows['sub_plus']}+{c}{V.rows['sub_pro']}+{c}{V.rows['ads']}", v05.row("營收彙總", 10), "", key="cons")
    V.add("企業端（席位＋API）", "$B", lambda i, c: f"={c}{V.rows['sub_seats']}+{c}{V.rows['api']}", v05.row("營收彙總", 15), "", key="ent")
    V.add("企業端占（消費端＋企業端）", "比例", lambda i, c: f"={c}{V.rows['ent']}/({c}{V.rows['cons']}+{c}{V.rows['ent']})", v05.row("營收彙總", 16),
          "對照：CFO 2026 年初消費／企業 60／40（SRC_OAI_007）", key="entshare")
    V.add("檢查：消費端＋企業端＋其他 − 總額（應為 0）", "$B",
          lambda i, c: f"={c}{V.rows['cons']}+{c}{V.rows['ent']}+{c}{V.rows['other']}-{c}{gr}", v05.row("營收彙總", 28), "", key="sumchk")
    V.add("總額年增率", "比例", lambda i, c: None if i == 0 else f"=({c}{gr}-{YC[i - 1]}{gr})/{YC[i - 1]}{gr}",
          v05.row("營收彙總", 27), "", key="growth")
    V.add("廣告占總額", "比例", lambda i, c: f"={c}{V.rows['ads']}/{c}{gr}", v05.row("營收彙總", 29), "", key="adshare")

    V.section("六、對照與差距（只列差距，不校準）")
    V.add("2025：模型總額 − SRC 實際營收（v0.5 校準方式，應為 0）", "$B", [f"=D{gr}-{S('營收：FY2025')}"] + [None] * 5, None,
          "2025 API 為倒推項，總額恆等於 SRC 13.07", name="REV_Gap2025", name_col="D", key="gap25")
    V.add("2025 未校準總額：API 改以處理 token × 組合有效單價", "$B",
          [f"=D{V.rows['sub']}+D{V.rows['ads']}+D{V.rows['other']}+V9_ApiTok*D{V.rows['mixprice']}/{mb}"] + [None] * 5, None,
          "處理 token＝Derived_V9 V22（2,628T）；顯示自下而上口徑與實際的差距", key="unc25")
    V.add("2025 未校準差距：未校準總額 − SRC 實際營收", "$B", [f"=D{V.rows['unc25']}-{S('營收：FY2025')}"] + [None] * 5, None, "",
          name="REV_GapUncal2025", name_col="D", key="uncgap25")
    V.add("2025 未校準差距（占實際）", "比例", [f"=D{V.rows['uncgap25']}/{S('營收：FY2025')}"] + [None] * 5, None, "", key="uncgap25p")
    V.add("FY2026 獨立推估＝Q1＋Q2＋2026-08 ARR × 下半年月數 ÷ 12", "$B",
          [None, f"={S('營收：2026 Q1')}+{S('營收：2026 Q2')}+{S('年化營收（ARR）：2026-08')}*{I('FY2026 獨立推估：下半年月數')}/{months}"] + [None] * 4,
          [None, 32.9] + [None] * 4, "v0.5 fy2026Independent（Derived）；v0.5 值 32.9（31–36）的下半年外推細節未記錄，本列為最簡口徑", name="REV_Indep2026", name_col="E", key="ind26")
    V.add("2026：模型總額 − 獨立推估", "$B", [None, f"=E{gr}-E{V.rows['ind26']}"] + [None] * 4, None, "", key="gap26")
    V.add("2026 下半年隱含（模型全年 − 上半年實際）", "$B", [None, f"=E{gr}-{S('營收：2026 Q1')}-{S('營收：2026 Q2')}"] + [None] * 4,
          v05.row("營收彙總", 30), "", key="h2")
    V.add("對照：管理層營收目標（只對照，D4c）", "$B",
          [None, f"={S('管理層營收目標：2026')}", None, None, None, f"={S('管理層營收目標：2030')}"], None,
          "7 月投資人簡報（SRC_OAI_043、044）；不驅動任何計算", key="mgmt")
    V.add("模型總額 − 管理層目標", "$B", [None, f"=E{gr}-E{{M}}", None, None, None, f"=I{gr}-I{{M}}"], None, "", key="mgmtgap")
    r = V.rows["mgmtgap"]
    for c in "EI":
        V.ws[f"{c}{r}"].value = V.ws[f"{c}{r}"].value.replace("{M}", str(V.rows["mgmt"]))

    # ── 第六節：價格事件時點（E8c）：基準／全部提前／全部延後 三欄 ──
    ev_row0 = V.r + 1
    put(V.ws, f"A{ev_row0}", "七、2026 API 牌價：價格事件時點天數加權（E8c；事件時點基準＝SRC_OAI 公告日，區間＝Inputs 偏移 ±1 季）", F_BOLD, fill=FILL_SEC)
    header(V.ws, ev_row0 + 1, ["編號", "項目", "單位", "基準", "全部提前（偏移取低）", "全部延後（偏移取高）", "2026 低", "2026 高", "", "v0.5 2026（對照）", "v0.5 低", "v0.5 高"])
    rr0 = ev_row0 + 2
    fy0 = I("FY2026 起日")
    evs = {}
    for k in (2, 3, 4, 5):
        d = S(f"API 價格事件 {k} 公告日（Excel 日期序列值）")
        off = I(f"價格事件 {k}（")
        orow = inp_row(off)
        evs[k] = rr0
        put(V.ws, f"A{rr0}", f"E{k}", F_CALC)
        put(V.ws, f"B{rr0}", f"價格事件 {k} 之後占 FY2026 的比例（天數加權）", F_CALC)
        put(V.ws, f"C{rr0}", "比例", F_CALC)
        for col, oc in (("D", "E"), ("E", "F"), ("F", "G")):
            o = off if oc == "E" else f"Inputs!${oc}${orow}"
            dd = f"MAX({fy0},MIN({fy0}+{days},{d}+{o}))"
            put(V.ws, f"{col}{rr0}", f"=({fy0}+{days}-{dd})/{days}")
        put(V.ws, f"Q{rr0}", f"＝（年底次日 − 生效日）÷ 365；生效日＝SRC 公告日＋Inputs 偏移，限於 FY2026 內", F_NOTE)
        rr0 += 1
    # 各層級價格序列：(期初＝事件 1 價回填, [(事件, 事件後價)])
    snap = {("top", "in"): S("API 牌價：top 層 輸入"), ("top", "out"): S("API 牌價：top 層 輸出"), ("mid", "in"): S("API 牌價：mid 層 輸入"),
            ("mid", "out"): S("API 牌價：mid 層 輸出"), ("low", "in"): S("API 牌價：low 層 輸入"), ("low", "out"): S("API 牌價：low 層 輸出")}
    model_ = {"top": "GPT-5.6 Sol", "mid": "GPT-5.6 Terra", "low": "GPT-5.6 Luna"}
    v05_26 = {("top", "in"): (6.6, 5.0, 7.5), ("top", "out"): (35.8, 28.0, 40.0), ("mid", "in"): (2.29, 2.0, 2.5),
              ("mid", "out"): (13.25, 11.0, 15.0), ("low", "in"): (0.64, 0.4, 0.8), ("low", "out"): (3.83, 2.5, 5.0)}
    v05_26 = {k: tuple(R.get(f"tokenRevenue/api/listPriceFYAvg/{k[1]}/{k[0]}/{x}/1") for x in ("v", "lo", "hi")) for k in v05_26}
    for t in TIERS:
        for io in ("in", "out"):
            zh = "輸入" if io == "in" else "輸出"
            p0 = S(f"API 價格事件 1 牌價：{t} 層 {zh}（{model_[t]}）")
            if t == "top":
                steps = [(3, S(f"API 價格事件 3 牌價：top 層 {zh}（Sol 促銷）")), (4, snap[(t, io)])]
            else:
                nm2 = "Terra 降價" if t == "mid" else "Luna 降價"
                steps = [(2, S(f"API 價格事件 2 牌價：{t} 層 {zh}（{nm2}）")), (5, snap[(t, io)])]
            put(V.ws, f"A{rr0}", f"P{t[0].upper()}{io[0].upper()}", F_CALC)
            put(V.ws, f"B{rr0}", f"FY2026 平均牌價：{TIER_ZH[t]} {zh}", F_CALC)
            put(V.ws, f"C{rr0}", "$/M", F_CALC)
            for col in "DEF":
                prev, f = p0, f"={p0}"
                for k, p in steps:
                    f += f"+({p}-{prev})*{col}{evs[k]}"
                    prev = p
                put(V.ws, f"{col}{rr0}", f)
            put(V.ws, f"G{rr0}", f"=MIN(D{rr0}:F{rr0})")
            put(V.ws, f"H{rr0}", f"=MAX(D{rr0}:F{rr0})")
            for col, v in zip("JKL", v05_26[(t, io)]):
                put(V.ws, f"{col}{rr0}", v, F_NOTE)
            put(V.ws, f"Q{rr0}", f"期初（1/1 至首次變價）回填事件 1 價（{model_[t]}）；事件 {steps[0][0]}、{steps[1][0]} 依序變價；事件 4、5 後＝2026-09-25 牌價快照", F_NOTE)
            nm(wb, f"REV_LP2026_{t.capitalize()}_{io.capitalize()}_Lo", f"Revenue!$G${rr0}")
            nm(wb, f"REV_LP2026_{t.capitalize()}_{io.capitalize()}_Hi", f"Revenue!$H${rr0}")
            r = V.rows[f"lp_{t}_{io}"]
            V.ws[f"E{r}"].value = f"=D{rr0}"
            V.rows[f"ev_{t}_{io}"] = rr0
            rr0 += 1
    V.ev_first, V.ev_last = ev_row0 + 2, rr0 - 1

    # ── 回填 Demand 對 Revenue 的引用 ──
    rb = D.rows["api_b"]
    for i in range(1, 6):
        D.ws[f"{YC[i]}{rb}"].value = f"=Revenue!{YC[i]}{V.rows['mixprice']}/Revenue!{YC[i - 1]}{V.rows['mixprice']}"
    ra = D.rows["api_tok"]
    D.ws[f"D{ra}"].value = f"=Revenue!D{V.rows['calib']}/Revenue!D{V.rows['mixprice']}*{mb}"
    for t in TIERS:
        r = D.rows[f"api_tok_{t}"]
        for i in range(6):
            D.ws[f"{YC[i]}{r}"].value = f"={YC[i]}{D.rows['api_tok']}*Revenue!{YC[i]}{V.rows[f'mix_{t}']}"

    for ws, w in ((D.ws, 46), (V.ws, 52)):
        ws.column_dimensions["A"].width = 7
        ws.column_dimensions["B"].width = w
        ws.column_dimensions["C"].width = 10
        for c in YC + V05C:
            ws.column_dimensions[c].width = 11
        ws.column_dimensions["Q"].width = 70
        ws.freeze_panes = "D7"
    return dict(D=D, V=V)
