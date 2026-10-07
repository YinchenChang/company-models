"""v0.5 JSON → v0.6 遷移對照（工作單 P1-2、P1-4、P1-5）。

這支程式只做「搬運與分類」，不產生任何新數值：
- 每個 JSON 葉節點（654 個）都必須有去處：SRC_OAI、Inputs、公式（後續工作包）、重複併入、或「不遷入＋理由」。
- 分類規則（工作單 P1-2）：Assumed／Analogy → Inputs；Derived → 公式；Interested-party／Verified → SRC_OAI；
  分類不明（混合標記、標記與註記互相矛盾）→ 依下列「保留原狀」原則處理並列入「待 Project 判斷」。
- 值一律從 reference/v0.5/openai_token_revenue.json 讀出，不在此處手寫數值。
"""
from __future__ import annotations

import json
from pathlib import Path

import source_rules

REPO = Path(__file__).resolve().parent.parent
JSON_PATH = REPO / "reference" / "v0.5" / "openai_token_revenue.json"

YEARS = [2025, 2026, 2027, 2028, 2029, 2030]
META_KEYS = {"tag", "src", "chk", "_note", "unit", "gwNote", "cellRefs", "note"}
SEP = "/"


def load():
    return json.loads(JSON_PATH.read_text(encoding="utf-8"))


def walk(o, path=()):
    """展開為葉節點：(path 字串, 值)。陣列元素以索引為路徑段。"""
    if isinstance(o, dict):
        for k, v in o.items():
            yield from walk(v, path + (k,))
    elif isinstance(o, list):
        for i, v in enumerate(o):
            yield from walk(v, path + (str(i),))
    else:
        yield SEP.join(path), o


class Recorder:
    """記錄每個葉節點的去處；同一葉節點只能有一個去處。"""

    def __init__(self, data):
        self.data = data
        self.leaves = dict(walk(data))
        self.dest: dict[str, tuple[str, str, str]] = {}      # path → (類別, 去處, 理由)
        self.src_rows: list[dict] = []
        self.inp_rows: list[dict] = []
        self.pending: list[dict] = []                         # 待 Project 判斷（P1）
        self.v9: list[dict] = []                              # V9 跨年拆開清單（含掃描後未拆者）

    # ── 取值 ────────────────────────────────────────────────
    def get(self, path):
        o = self.data
        for seg in path.split(SEP):
            o = o[int(seg)] if isinstance(o, list) else o[seg]
        return o

    def has(self, path):
        try:
            self.get(path)
            return True
        except (KeyError, IndexError, TypeError):
            return False

    def under(self, prefix):
        return [p for p in self.leaves if p == prefix or p.startswith(prefix + SEP)]

    # ── 去處登記 ────────────────────────────────────────────
    def assign(self, path, kind, ref, why=""):
        if path not in self.leaves:
            raise KeyError(f"不存在的葉節點：{path}")
        if path in self.dest:
            raise ValueError(f"葉節點重複登記：{path}：{self.dest[path]} 與 {(kind, ref)}")
        self.dest[path] = (kind, ref, why)

    def take(self, prefix, kind, ref, why="", only_unassigned=True):
        """prefix 之下尚未登記者一律登記。"""
        n = 0
        for p in self.under(prefix):
            if p in self.dest:
                if only_unassigned:
                    continue
            self.assign(p, kind, ref, why)
            n += 1
        return n

    def tag_of(self, node, default=None):
        """節點本身或最近的祖先標記。"""
        parts = node.split(SEP)
        while parts:
            p = SEP.join(parts + ["tag"])
            if p in self.leaves:
                return self.leaves[p]
            parts.pop()
        return default

    def meta_of(self, node, key, default=""):
        parts = node.split(SEP)
        while parts:
            p = SEP.join(parts + [key])
            if p in self.leaves:
                return self.leaves[p]
            parts.pop()
        return default

    # ── SRC_OAI 列 ─────────────────────────────────────────
    def src(self, node, metric, *, v=None, lo=None, hi=None, vlit=None, unit="", scope="", date="—", subject="OpenAI",
            source=None, status="Active", alt_of="—", use="", note="", tag=None, extra_leaves=(), why_extra=""):
        """登記一列 SRC_OAI。v／lo／hi 為葉節點路徑（相對 JSON 根）；vlit 為文字值（葉節點本身為字串者）。
        node 之下尚未登記的中介資料（tag、src、chk、_note…）一併併入此列。"""
        rid = f"SRC_OAI_{len(self.src_rows) + 1:03d}"
        row = dict(id=rid, metric=metric, value="", lo="", hi="", unit=unit, scope=scope,
                   subject=subject, date=date, status=status, alt_of=alt_of, use=use, note=note, v05=node)
        val = self.get(v) if v else vlit
        row["value"] = "" if val is None else val
        if lo:
            row["lo"] = self.get(lo)
        if hi:
            row["hi"] = self.get(hi)
        row["tag"] = tag or self.tag_of(node, "未標")
        row["source"] = source if source is not None else (self.meta_of(node, "src") or "—")
        row["chk"] = self.meta_of(node, "chk") or "—"
        for p in (v, lo, hi):
            if p:
                self.assign(p, "SRC", rid)
        for p in extra_leaves:
            self.assign(p, "SRC", rid, why_extra)
        # 同節點的中介資料併入本列（只併入尚未登記者）
        for p in self.under(node):
            leaf = p.split(SEP)[-1]
            if p not in self.dest and leaf in META_KEYS and p.rsplit(SEP, 1)[0] == node:
                self.assign(p, "SRC", rid, "隨附資訊")
        self.src_rows.append(row)
        return rid

    # ── Inputs 列 ──────────────────────────────────────────
    def inp(self, node, name, *, unit="", index="—", v=None, lo=None, hi=None, vlit=None, decision="—", note="", tag=None,
            status="v0.6 沿用", meta=True):
        rid = f"INP_{len(self.inp_rows) + 1:03d}"
        val = self.get(v) if v else vlit
        row = dict(id=rid, name=name, index=index, unit=unit, value="" if val is None else val,
                   lo="" if not lo else self.get(lo), hi="" if not hi else self.get(hi),
                   tag=tag or self.tag_of(node, "Decision"), decision=decision, v05=node, note=note, status=status)
        for p in (v, lo, hi):
            if p:
                self.assign(p, "INP", rid)
        if meta:
            for p in self.under(node):
                leaf = p.split(SEP)[-1]
                if p not in self.dest and leaf in META_KEYS and p.rsplit(SEP, 1)[0] == node:
                    self.assign(p, "INP", rid, "隨附資訊")
        self.inp_rows.append(row)
        return rid

    def split(self, item, detail, dest, scanned_only=False):
        self.v9.append(dict(item=item, detail=detail, dest=dest, scanned_only=scanned_only))

    def todo(self, topic, detail, where=""):
        self.pending.append(dict(topic=topic, detail=detail, where=where))


# ─────────────────────────────────────────────────────────────
#  建立對照
# ─────────────────────────────────────────────────────────────
def build_map(data=None, write_registry=False) -> Recorder:
    R = Recorder(data or load())
    S, I = R.src, R.inp

    # 一、模型設定與文件（meta、_readme）
    R.assign("_readme", "SKIP", "—", "版本說明文字；v0.6 的版本說明在 README／CHANGELOG")
    R.assign("tokenRevenue/meta/company", "SKIP", "—", "公司名稱：表頭文字，不是參數")
    R.assign("tokenRevenue/meta/fiscalYearEndMonth", "SKIP", "—", "工作單第 0 節已定：曆年制（12 月）；v0.6 不設此參數")
    R.assign("tokenRevenue/meta/version", "SKIP", "—", "v0.4 的版本字串，已過時")
    R.assign("tokenRevenue/meta/decisions", "SKIP", "—", "決策 ID 清單：v0.6 逐項在 Inputs「v0.5 原決策」欄保留；完整文字見 reference/v0.5")
    for k in ("revenue", "tokens", "price", "subs", "arpu"):
        R.assign(f"tokenRevenue/meta/units/{k}", "SKIP", "—", "單位字串：v0.6 每列自帶單位欄")
    I("tokenRevenue/meta/baseYear", "基準年", unit="年", v="tokenRevenue/meta/baseYear", tag="Decision", decision="—", meta=False)
    I("tokenRevenue/meta/calibrationYears/0", "校準年（第 1 個）", unit="年", v="tokenRevenue/meta/calibrationYears/0",
            tag="Decision", note="v0.5：2025、2026 為校準年；2026 為預測校準（見 demandGrowthExPrice 註記）", meta=False)
    I("tokenRevenue/meta/calibrationYears/1", "校準年（第 2 個）", unit="年", v="tokenRevenue/meta/calibrationYears/1", tag="Decision", meta=False)
    I("tokenRevenue/meta/elasticityFromYear", "彈性起算年", unit="年", v="tokenRevenue/meta/elasticityFromYear", tag="Decision", meta=False)
    for i, y in enumerate(YEARS):
        I(f"tokenRevenue/meta/years/{i}", "模型年度", index=f"第 {i + 1} 年", unit="年", v=f"tokenRevenue/meta/years/{i}", tag="Decision",
          decision="D10→FY2030", note="曆年制；2025 為實際校準年（工作單第 0 節）", meta=False)

    # 二、營收與營運實績（SRC_OAI）
    A = "tokenRevenue/actuals"
    S(f"{A}/revenue2024", "營收：FY2024", v=f"{A}/revenue2024/v", unit="$B", scope="FY2024 認列營收", use="P2 校準；P5 對帳")
    S(f"{A}/revenue2025", "營收：FY2025", v=f"{A}/revenue2025/v", unit="$B", scope="FY2025 認列營收（外流經查核財報）", date="2026-06",
      use="P2 校準；P5 對帳")
    S(f"{A}/cogs2025", "營業成本（cost of revenue）：FY2025", v=f"{A}/cogs2025/v", unit="$B", scope="含推論；FY2025", date="2026-06",
      use="P4 對帳")
    S(f"{A}/revenueQ1_2026", "營收：2026 Q1", v=f"{A}/revenueQ1_2026/v", unit="$B", scope="2026 Q1（股東文件）", date="2026-06-16",
      use="P2 對照（fy2026Independent 的組成）")
    S(f"{A}/revenueQ2_2026", "營收：2026 Q2", v=f"{A}/revenueQ2_2026/v", unit="$B", scope="2026 Q2（告知投資人）", date="2026-08-18",
      use="P2 對照")
    S(f"{A}/arrAug2026", "年化營收（ARR）：2026-08", v=f"{A}/arrAug2026/v", unit="$B", scope="ARR；2026 年 8 月", date="2026-08-13",
      use="P2 對照")
    S(f"{A}/consumerShareStart2026", "消費端營收占比：2026 年初", v=f"{A}/consumerShareStart2026/v", unit="比例",
      scope="年初消費／企業 60／40；8 月企業已超過消費端", date="2026-08-14", use="P2 對照")
    S(f"{A}/apiTokensPerMinOct2025", "API 每分鐘處理 token：2025-10", v=f"{A}/apiTokensPerMinOct2025/v", unit="B tok/min",
      scope="API 每分鐘 token；DevDay 揭露", date="2025-10", use="P2 校準檢查", note="2026 年更新揭露檢索未得（v0.5 註記）")
    # apiTokensFY2025：Derived → 公式（E8a：驅動值進 Inputs，2628＝每分鐘 token × 每年分鐘數 ÷ 單位換算）
    R.take(f"{A}/apiTokensFY2025", "FORMULA", "Derived_V9（V10–V12：API FY2025 token＝每分鐘 token×每年分鐘數÷單位換算）",
           "E8a：Derived→公式；驅動值『FY 平均每分鐘 5B token』進 Inputs（Assumed，3.8–6.1）；v0.5 值 2,628（2,000–3,200）並列於 Derived_V9")
    I("r3/E8a/每分鐘 token", "API FY2025 平均每分鐘 token（驅動值）", unit="B tok/min", vlit=5, tag="Assumed", decision="E8a",
      note="E8a：FY 平均每分鐘 5B token，區間 3.8–6.1（v0.5 出處：workbook 列 242 估計；DevDay 2025-10 揭露 6.0 見 SRC_OAI）", meta=False)
    R.inp_rows[-1]["lo"], R.inp_rows[-1]["hi"] = 3.8, 6.1
    I("r3/E8a/每年分鐘數", "每年分鐘數（定義常數）", unit="分", vlit=525600, tag="Decision", decision="E6",
      note="來源：定義常數（365×24×60）；區間不適用（低＝高）；E6：公式內不得含常數，故進 Inputs", meta=False)
    R.inp_rows[-1]["lo"], R.inp_rows[-1]["hi"] = 525600, 525600
    I("r3/E8a/單位換算", "單位換算：B token → T token 的除數（定義常數）", unit="B／T", vlit=1000, tag="Decision", decision="E6",
      note="來源：定義常數；區間不適用（低＝高）；E6：公式內不得含常數，故進 Inputs", meta=False)
    R.inp_rows[-1]["lo"], R.inp_rows[-1]["hi"] = 1000, 1000

    M = "tokenRevenue/msShare"
    S(f"{M}/rate", "Microsoft 營收分成率", v=f"{M}/rate/v", unit="比例", scope="OpenAI 付 Microsoft 之營收分成", date="2026-04-27",
      subject="Microsoft", use="P2 淨額")
    S(f"{M}/capTotal", "Microsoft 分成總額上限", v=f"{M}/capTotal/v", unit="$B", scope="分成累計上限", date="2026-05", subject="Microsoft",
      use="P2 淨額")
    S(f"{M}/capThroughYear", "Microsoft 分成付款持續至", v=f"{M}/capThroughYear/v", unit="年", scope="付款持續至該年；受總額上限",
      date="2026-04-27", subject="Microsoft", use="P2 淨額")
    S(f"{M}/azureInflowAfterApr2026", "Microsoft 對 OpenAI 的分潤（2026-04 後）", v=f"{M}/azureInflowAfterApr2026/v", unit="$B",
      scope="2026-04-27 起 Microsoft 不再支付 OpenAI 分潤", date="2026-04-27", subject="Microsoft", use="P2 淨額")
    I(f"{M}/capCountFromYear", "分成上限起算年", unit="年", v=f"{M}/capCountFromYear/v", decision="N2a",
      note="起算點未揭露；Andy 2026-09-29 決定 N2a（v0.5 無區間，待 Project 補）")

    # 三、API 單價快照、訂閱單價快照（SRC_OAI）
    API = "tokenRevenue/api"
    for tier in ("top", "mid", "low"):
        for f, lab in (("in", "輸入"), ("cached", "快取輸入"), ("out", "輸出")):
            S(f"{API}/listPriceSnapshot", f"API 牌價：{tier} 層 {lab}", v=f"{API}/listPriceSnapshot/{tier}/{f}", unit="$/M tokens",
              scope=f"短上下文 Standard；層級＝{R.get(f'{API}/tierMapSep2026/{tier}')}", date="2026-09-25", use="P2 單價路徑起點")
    for tier in ("top", "mid", "low"):
        pth = f"{API}/tierMapSep2026/{tier}"
        S(pth, f"API 層級對應：{tier}", vlit=R.get(pth), unit="文字", scope="2026-09 各層級的旗艦型號", date="2026-09-25",
          tag="Verified", source="OpenAI 價格頁 developers.openai.com/api/docs/pricing（2026-09-25）", use="P2 層級對照",
          note="E8b：文字列，只作標籤（Verified）", extra_leaves=(pth,))
    for i, ev in enumerate(R.get(f"{API}/priceEvents2026/list")):
        S(f"{API}/priceEvents2026", f"API 價格事件 {i + 1}", vlit=ev, unit="文字", scope="價格事件（輸入／輸出 $/M）",
          date="2026-" + ev[:5] if ev[:2].isdigit() else "—", use="P2 FY 平均牌價（2026）的月加權依據", extra_leaves=(f"{API}/priceEvents2026/list/{i}",))
    S(f"{API}/cachedInputRatio", "快取輸入價占輸入價比", v=f"{API}/cachedInputRatio/v", unit="比例", scope="定價頁：cached input＝輸入價 10%",
      date="2026-09", use="P2")

    SUB = "tokenRevenue/subscription"
    S(f"{SUB}/listPriceSep2026", "訂閱牌價：Go", v=f"{SUB}/listPriceSep2026/go", unit="$/月", scope="ChatGPT 定價頁；美國牌價", date="2026-09",
      use="P2 ARPU 上限參照")
    S(f"{SUB}/listPriceSep2026", "訂閱牌價：Plus", v=f"{SUB}/listPriceSep2026/plus", unit="$/月", scope="ChatGPT 定價頁；美國牌價", date="2026-09",
      use="P2")
    S(f"{SUB}/listPriceSep2026", "訂閱牌價：Pro（低檔）", v=f"{SUB}/listPriceSep2026/pro/0", unit="$/月", scope="2026 起 Pro 兩檔之一",
      date="2026-09", use="P2")
    S(f"{SUB}/listPriceSep2026", "訂閱牌價：Pro（高檔）", v=f"{SUB}/listPriceSep2026/pro/1", unit="$/月", scope="2026 起 Pro 兩檔之一",
      date="2026-09", use="P2")
    S(f"{SUB}/listPriceSep2026", "訂閱牌價：Business 每席", lo=f"{SUB}/listPriceSep2026/businessSeat/0", hi=f"{SUB}/listPriceSep2026/businessSeat/1",
      unit="$/席/月", scope="Business 席位區間；Enterprise 為議價", date="2026-09", use="P2")
    # 無區間的數值列：若 value 為空而 lo/hi 有值，Excel 側以低／高呈現（見 builder 檢查：不視為缺值）

    MG = f"{SUB}/migration"
    S(f"{MG}/planGo2026", "公司 2026 計畫：Go 用戶數", v=f"{MG}/planGo2026/v", unit="M", scope="年初內部預測", date="2026-04-28", use="P2 T4 遷移幅度")
    S(f"{MG}/planPlus2026", "公司 2026 計畫：Plus 用戶數", v=f"{MG}/planPlus2026/v", unit="M", scope="年初內部預測", date="2026-04-28", use="P2 T4 遷移幅度")
    I(f"{MG}/planFraction", "T4 遷移幅度（占公司計畫）— 低", unit="比例", v=f"{MG}/planFraction/low", decision="T4", tag="Decision",
      note="情境值：Plus→Go 遷移與 Go 額外新增的幅度比例", meta=False)
    I(f"{MG}/planFraction", "T4 遷移幅度（占公司計畫）— 基準", unit="比例", v=f"{MG}/planFraction/base", decision="T4", tag="Decision",
      note="Andy 2026-09-29 採預設", meta=False)
    I(f"{MG}/planFraction", "T4 遷移幅度（占公司計畫）— 高", unit="比例", v=f"{MG}/planFraction/high", decision="T4", tag="Decision", meta=True)

    # 四、需求與價格假設（Inputs）
    FY = f"{API}/listPriceFYAvg"
    R.assign(f"{FY}/_note", "INP", "INP（listPriceFYAvg）", "說明：2025 為 Analogy、2026 由事件月加權 [Derived]")
    for io in ("in", "out"):
        for tier in ("top", "mid", "low"):
            n = f"{FY}/{io}/{tier}"
            I(n, f"FY 平均牌價（{'輸入' if io == 'in' else '輸出'}）：{tier} 層", index="2025", unit="$/M tokens",
              v=f"{n}/v/0", lo=f"{n}/lo/0", hi=f"{n}/hi/0", tag="Analogy", decision="—",
              note="2025 為 Analogy（v0.5 標 Analogy/Derived 混合；2025 屬 Analogy）", meta=False)
            for p in (f"{n}/v/1", f"{n}/lo/1", f"{n}/hi/1"):
                R.assign(p, "FORMULA", "Revenue R01–R06 的 2026 欄（P2 實作：價格事件時點天數加權，Revenue 第七節；E8c）",
                         "Derived：2026 由價格事件按月加權；v0.5 標 Analogy/Derived 混合，依元素分開（2025→Inputs、2026→公式）")
            R.assign(f"{n}/tag", "INP", "listPriceFYAvg 各列", "隨附資訊")
    ch = f"{API}/listPriceAnnualChange"
    R.assign(f"{ch}/_note", "INP", "listPriceAnnualChange", "說明文字")
    for tier in ("top", "mid", "low"):
        I(f"{ch}/{tier}", f"牌價年變動率：{tier} 層", unit="比例/年", v=f"{ch}/{tier}/v", lo=f"{ch}/{tier}/lo", hi=f"{ch}/{tier}/hi",
          decision="D11", note="2027 起每年；情境軸之一")

    # tokenMix：low 為 1−top−mid（v0.5 src 註明由 Excel 計算）→ 公式
    tm = f"{API}/tokenMix"
    for tier in ("top", "mid"):
        for i, lab in enumerate(("陣列[0]", "陣列[1]")):
            I(f"{tm}/{tier}", f"API token 層級占比：{tier} 層", index=lab, unit="比例", v=f"{tm}/{tier}/v/{i}", lo=f"{tm}/{tier}/lo/{i}",
              hi=f"{tm}/{tier}/hi/{i}", note="v0.5 陣列兩位（年份對應待 P2 確認：占比 top 0.40→0.15 判斷為 2025→2026）",
              meta=(i == 1))
    R.take(f"{tm}/low", "FORMULA", "Revenue R12（P2 實作：low＝1−top−mid）", "v0.5 出處註明『low＝1−top−mid，由 Excel 計算，不另輸入』；與 Assumed 標記矛盾，依註明改為公式")
    I(f"{API}/outputShare", "輸出 token 占比（API）", unit="比例", v=f"{API}/outputShare/v", lo=f"{API}/outputShare/lo", hi=f"{API}/outputShare/hi",
      tag="Analogy")
    I(f"{API}/cacheHitRate", "快取命中率（API）", unit="比例", v=f"{API}/cacheHitRate/v", lo=f"{API}/cacheHitRate/lo", hi=f"{API}/cacheHitRate/hi")
    I(f"{API}/discountRate", "加權折扣率（Batch/Flex／企業）", unit="比例", v=f"{API}/discountRate/v", lo=f"{API}/discountRate/lo",
      hi=f"{API}/discountRate/hi")
    I(f"{API}/tokensPerTaskBase", "基準每任務 token（層級正規化）", unit="K tok/task", v=f"{API}/tokensPerTaskBase/v",
      lo=f"{API}/tokensPerTaskBase/lo", hi=f"{API}/tokensPerTaskBase/hi", note="只影響 N 與 k 的水準拆分，不影響營收")
    dg = f"{API}/demandGrowthExPrice"
    R.assign(f"{dg}/_note", "INP", "demandGrowthExPrice", "說明：2026 為校準年（v0.4 調整）")
    for k, lab in (("tasks", "任務數成長率（不含價格）"), ("tokensPerTask", "每任務 token 成長率（不含價格）")):
        for i, y in enumerate(YEARS):
            I(f"{dg}/{k}", lab, index=str(y), unit="比例/年", v=f"{dg}/{k}/v/{i}", note="2025 欄為 0（基準年）" if i == 0 else "",
              meta=False)
        I(f"{dg}/{k}", lab + "：低情境倍數", unit="倍", v=f"{dg}/{k}/lo_mult", note="區間＝基準 × 倍數", meta=False)
        I(f"{dg}/{k}", lab + "：高情境倍數", unit="倍", v=f"{dg}/{k}/hi_mult", note="區間＝基準 × 倍數", meta=True)
    el = f"{API}/elasticity"
    for k, lab in (("tasks", "價格彈性：任務數"), ("tokensPerTask", "價格彈性：每任務 token")):
        I(f"{el}/{k}", lab, unit="彈性", v=f"{el}/{k}/v", lo=f"{el}/{k}/lo", hi=f"{el}/{k}/hi", decision="D11",
          note="無法由公開資料分離（價格與能力同時變動）" if k == "tasks" else "")
    R.assign(f"{API}/calibration", "SKIP", "—", "校準方式的文字說明；P2 以公式實作，說明保留於 CLAUDE.md／報告")
    I(f"{API}/priceBasis", "價格基準（開關）", unit="文字（tierRole）", vlit=R.get(f"{API}/priceBasis/v"), tag="Decision", decision="N3a",
      note="新旗艦上市視為該層價格上升，不另立層級（Decision 標記 → Inputs 開關；補充 1 第 3 點）", meta=False)
    R.assign(f"{API}/priceBasis/v", "INP", f"INP_{len(R.inp_rows):03d}")
    R.assign(f"{API}/priceBasis/tag", "INP", f"INP_{len(R.inp_rows):03d}", "隨附資訊")
    R.assign(f"{API}/priceBasis/src", "INP", f"INP_{len(R.inp_rows):03d}", "隨附資訊")
    R.take(f"{API}/tiers", "SKIP", "—", "層級清單（top／mid／low）：結構，由公式頁的列結構體現")
    R.assign(f"{API}/listPriceSnapshot/date", "SRC", "API 牌價快照列", "快照日期：寫入各列日期欄")

    # 訂閱側假設
    R.take(f"{SUB}/tiers", "SKIP", "—", "方案清單：結構")
    R.take(f"{SUB}/line", "SKIP", "—", "方案所屬營收線（consumer／enterprise）：結構，P2 以列結構體現")
    # apiEquivalentTier：字串值 → 每方案一列
    ae = f"{SUB}/apiEquivalentTier"
    for plan in ("free", "go", "plus", "pro", "seats"):
        I(ae, f"方案對應 API 層級：{plan}", index=plan, unit="文字", vlit=R.get(f"{ae}/{plan}"), tag=R.get(f"{ae}/tag"),
          note=R.get(f"{ae}/src"), meta=False)
        R.assign(f"{ae}/{plan}", "INP", f"INP_{len(R.inp_rows):03d}")
    R.assign(f"{ae}/tag", "INP", "apiEquivalentTier 各列", "隨附資訊")
    R.assign(f"{ae}/src", "INP", "apiEquivalentTier 各列", "隨附資訊")

    sf = f"{SUB}/subsFYAvg"
    R.assign(f"{sf}/_note", "INP", "subsFYAvg", "說明：FY 平均人數；free 僅用於成本與 token")
    MSRC = "r2/V9 出處文字"       # 出處文字內的觀測值（v0.5 無獨立葉節點）
    # V9：跨年混合依年度拆開。formula＝改為公式的元素索引；src_rows＝由出處文字新增的 SRC_OAI 觀測列；src_elem＝元素本身即觀測值
    plan_cfg = {
        "free": dict(formula=[1], tag_in="Analogy", src_rows=[
            dict(metric="ChatGPT 週活躍用戶（WAU）：2026-02", vlit=900, unit="M", scope="OpenAI 公告；週活躍用戶", date="2026-02",
                 source="OpenAI 公告（v0.5 free 出處：WAU 900M，2026-02）", tag="Interested-party", use="P2 免費用戶人數（2026 公式的依據）")]),
        "go": dict(formula=[], tag_in="Assumed", src_rows=[
            dict(metric="Go 付費用戶：2025 年底", vlit=3, unit="M", scope="年底人數（非 FY 平均）", date="2025-12",
                 source="The Information（v0.5 go 出處：2025 年 3M，年底）", tag="Interested-party", use="P2 Go 2025 起點（錨點；FY 平均仍為 Assumed）")]),
        "plus": dict(formula=[0], tag_in="Assumed", src_rows=[
            dict(metric="ChatGPT Plus＋Pro 付費訂閱：2025-07", vlit=35, unit="M", scope="Plus＋Pro 合計約 35M", date="2025-07",
                 source="Reuters／The Information 2025-11（v0.5 plus 出處）", tag="Interested-party", use="P2 Plus 2025 FY 平均（公式依據）"),
            dict(metric="ChatGPT 付費訂閱：2025 年（內部文件）", vlit=44, unit="M", scope="OpenAI 內部文件 2025 年 44M", date="2026-04-28",
                 source="The Information 2026-04-28（v0.5 plus 出處）", tag="Interested-party", use="P2 Plus 2025 FY 平均（公式依據）")]),
        "pro": dict(formula=[0, 1, 2, 3, 4, 5], tag_in="Assumed", src_rows=[
            dict(metric="Pro 用戶倍數：2026 對 2025（內部預測：倍增）", vlit=2, unit="倍", scope="內部預測：Pro 用戶 2026 為 2025 的 2 倍", date="2026-04-28",
                 source="The Information 2026-04-28（v0.5 pro 出處：內部預測）", tag="Interested-party", use="V12 約束一：2025＝2026÷此倍數"),
            dict(metric="Pro 占付費訂閱總數：2026 上限（內部預測：<1%）", hi=0.01, unit="比例", scope="內部預測：Pro 占付費總數低於 1%（上限，非點值）", date="2026-04-28",
                 source="The Information 2026-04-28（v0.5 pro 出處：內部預測）", tag="Interested-party", use="V12 約束二：2026 占比不得超過此上限")]),
        "seats": dict(formula=[], tag_in="Assumed", src_rows=[
            dict(metric="付費企業用戶：2026-02", vlit=9, lo=9, unit="M", scope="逾 9M 付費企業用戶（『逾』＝下限）", date="2026-02",
                 source="OpenAI 2026-02（v0.5 seats 出處）", tag="Interested-party", use="P2 企業席位錨點（FY 平均仍為 Assumed）")]),
    }
    for plan in ("free", "go", "plus", "pro", "seats"):
        n = f"{sf}/{plan}"
        cfg = plan_cfg[plan]
        for sr in cfg.get("src_rows", []):
            S(f"{n}/{MSRC}", sr["metric"], vlit=sr.get("vlit"), unit=sr["unit"], scope=sr["scope"], date=sr["date"], source=sr["source"],
              tag=sr["tag"], use=sr["use"], note="v0.5 出處文字內的觀測值（V9：拆出為 SRC_OAI 列）")
            if "lo" in sr:
                R.src_rows[-1]["lo"] = sr["lo"]
            if "hi" in sr:
                R.src_rows[-1]["hi"] = sr["hi"]
        for i, y in enumerate(YEARS):
            vp = f"{n}/v/{i}"
            if i in cfg.get("src_elem", {}):
                e = cfg["src_elem"][i]
                S(n, e["metric"], v=vp, unit=e["unit"], scope=e["scope"], date=e["date"], source=e["source"], tag=e["tag"], use=e["use"],
                  note="V9：2026 元素為內部預測（觀測），進 SRC_OAI")
            elif plan == "pro" and i in cfg["formula"]:
                R.assign(vp, "FORMULA", "Derived_V9（%s：pro %s＝%s）" % ("V15" if i >= 2 else "V12", y, "2026÷SRC 倍數" if i == 0 else ("比例×付費總數（含席位）" if i == 1 else "比例_t/(1−比例_t)×(go＋plus＋seats)")),
                         "%s：v0.5 值 %s 改為公式；v0.5 原值並列於 Derived_V9" % ("V15" if i >= 2 else "V12", R.get(vp)))
            elif i in cfg["formula"]:
                R.assign(vp, "FORMULA", "Derived_V9（FY 平均人數 %s：觀測值×換算係數）" % y,
                         "V9：該年有觀測值，v0.5 值 %s 改為由 SRC 推得的公式（P2 定義換算；v0.5 值供對照）" % R.get(vp))
            else:
                tg = cfg["tag_in"] if i == 0 else "Assumed"
                I(n, f"FY 平均用戶數：{plan}", index=str(y), unit="M", v=vp, tag=tg,
                  note=("V9：2026+ 為 Assumed（無遷移基準）" if i > 0 else "") + ("；2025 元素 Analogy（v0.5 標記）" if tg == "Analogy" else ""), meta=False)
        if plan == "pro":      # V15：pro 全期改比例公式，區間由比例的區間傳遞，不再有絕對值倍數
            R.take(f"{n}/lo_mult", "SKIP", "—", "V15：pro 全期改為比例公式，區間由比例 0.5–2.0% 傳遞；倍數不再使用（v0.5 原值 0.6／1.2 見對照表）")
            R.take(f"{n}/hi_mult", "SKIP", "—", "V15：同上")
            continue
        I(n, f"FY 平均用戶數：{plan}：低情境倍數", unit="倍", v=f"{n}/lo_mult", tag="Assumed", note="區間＝基準 × 倍數（適用 Inputs 內的年份）", meta=False)
        I(n, f"FY 平均用戶數：{plan}：高情境倍數", unit="倍", v=f"{n}/hi_mult", tag="Assumed", note="區間＝基準 × 倍數（適用 Inputs 內的年份）", meta=True)
    # E6：V9 公式內不得有常數；換算係數全進 Inputs（標記＋區間）
    wau = plan_cfg["free"]["src_rows"][0]["vlit"]
    p35 = plan_cfg["plus"]["src_rows"][0]["vlit"]
    I(f"{sf}/free/r3換算係數", "換算係數：free FY2026 平均 ÷ 2026-02 WAU", unit="倍", vlit=R.get(f"{sf}/free/v/1") / wau, tag="Assumed", decision="E6",
      note="由 v0.5 值 950 ÷ WAU 900 還原；區間取 v0.5 free 的低／高倍數 0.8／1.2（出處無其他區間）。P2 可改為 WAU 成長路徑", meta=False)
    R.inp_rows[-1]["lo"] = R.get(f"{sf}/free/v/1") * R.get(f"{sf}/free/lo_mult") / wau
    R.inp_rows[-1]["hi"] = R.get(f"{sf}/free/v/1") * R.get(f"{sf}/free/hi_mult") / wau
    I(f"{sf}/plus/r3換算係數", "換算係數：Plus FY2025 平均 ÷ 2025-07 Plus＋Pro 付費訂閱", unit="倍", vlit=R.get(f"{sf}/plus/v/0") / p35, tag="Assumed",
      decision="E6", note="由 v0.5 值 30 ÷ 35 還原；區間取 v0.5 出處文字『FY 平均 28–33M』÷ 35", meta=False)
    R.inp_rows[-1]["lo"], R.inp_rows[-1]["hi"] = 28 / p35, 33 / p35
    I(f"{sf}/pro/r3比例", "Pro 占付費訂閱總數比例（2026）", unit="比例", vlit=0.008, tag="Assumed", decision="V12",
      note="V12：0.8%（0.5–1.0%）；2026 pro＝比例×付費總數（含席位）；須低於 SRC_OAI 上限 1%", meta=False)
    R.inp_rows[-1]["lo"], R.inp_rows[-1]["hi"] = 0.005, 0.01
    for y in (2027, 2028, 2029, 2030):
        I(f"{sf}/pro/r3比例", "Pro 占付費訂閱總數比例（2027+ 逐年）", index=str(y), unit="比例", vlit=0.008, tag="Assumed", decision="V15",
          note="<1% 約束只是 2026 內部預測，2027+ 無來源，上限放寬至 2.0% 以保留 $100／$200 兩檔擴大客群的可能", meta=False)
        R.inp_rows[-1]["lo"], R.inp_rows[-1]["hi"] = 0.005, 0.02
    R.split("subscription.subsFYAvg.free", "2026-02 WAU 900M（OpenAI 公告）為觀測；2026 元素 950→公式（WAU×Inputs 換算係數）；2025（Analogy）與 2027+（Assumed）→Inputs", "SRC_OAI 1 列＋公式 1 格＋Inputs 4 格")
    R.split("subscription.subsFYAvg.plus", "2025 報導值（Plus＋Pro 約 35M、內部文件 44M）為觀測；2025 元素 30（28–33）→公式（35M×Inputs 換算係數）；2026+（Assumed）→Inputs", "SRC_OAI 2 列＋公式 1 格＋Inputs 5 格")
    R.split("subscription.subsFYAvg.pro", "V12／V15：SRC 只登兩個約束（2026 為 2025 的 2 倍；2026 占付費總數 <1%）；Inputs：2026 占比 0.8%（0.5–1.0%）、2027–2030 逐年占比 0.8%（0.5–2.0%）；2025–2030 全為公式（2025＝2026÷倍數；其餘＝比例/(1−比例)×(go＋plus＋seats)）", "SRC_OAI 2 列＋公式 6 格＋占比 Inputs 5 列")
    R.split("subscription.subsFYAvg.go", "2025 年底 3M（The Information）為年底人數，與 FY 平均口徑不同；v0.5 標 Assumed、換算規則未定：只新增 SRC_OAI 錨點列，6 個元素維持 Inputs", "SRC_OAI 1 列（錨點）＋Inputs 6 格", scanned_only=True)
    R.split("subscription.subsFYAvg.seats", "逾 9M 付費企業用戶（2026-02）為觀測，與 FY 平均口徑不同；v0.5 標 Assumed、換算規則未定：只新增 SRC_OAI 錨點列，6 個元素維持 Inputs", "SRC_OAI 1 列（錨點）＋Inputs 6 格", scanned_only=True)
    ar = f"{SUB}/arpu"
    R.assign(f"{ar}/_note", "INP", "arpu", "說明：實收 $/月")
    for plan in ("go", "plus", "pro", "seats"):
        n = f"{ar}/{plan}"
        nv = len(R.get(f"{n}/v"))
        for i in range(nv):
            lab = "全期" if nv == 1 else ("2025" if i == 0 else "2026 起（v0.5 未標年份；待 P2 確認）")
            I(n, f"ARPU（實收）：{plan}", index=lab, unit="$/月", v=f"{n}/v/{i}", lo=f"{n}/lo/{i}", hi=f"{n}/hi/{i}", meta=(i == nv - 1))
    for sec, lab, unit in (("tasksPerDay", "每日任務數", "任務/日"), ("tokensPerTask", "每任務 token", "K tok/task")):
        n0 = f"{SUB}/{sec}"
        R.assign(f"{n0}/_note", "INP", sec, "說明文字")
        for plan in ("free", "go", "plus", "pro", "seats"):
            for i, y in enumerate((2025, 2026)):
                I(f"{n0}/{plan}", f"{lab}：{plan}", index=str(y), unit=unit, v=f"{n0}/{plan}/v/{i}", tag=R.get(f"{n0}/tag"),
                  note="2027 起年增（見成長率）", meta=False)
        for k in ("tag", "src", "chk"):
            R.assign(f"{n0}/{k}", "INP", f"{sec} 各列", "隨附資訊")
    for k, lab in (("tasksPerDayGrowth", "每日任務數年增率（2027 起）"), ("tokensPerTaskGrowth", "每任務 token 年增率（2027 起）")):
        I(f"{SUB}/{k}", lab, unit="比例/年", v=f"{SUB}/{k}/v", lo=f"{SUB}/{k}/lo", hi=f"{SUB}/{k}/hi",
          decision="D11" if k == "tokensPerTaskGrowth" else "—", note="情境軸" if k == "tokensPerTaskGrowth" else "")

    # 廣告與其他
    ads = "tokenRevenue/other/ads"
    R.assign(f"{ads}/_note", "INP", "ads", "說明：N1b 廣告為獨立列")
    R.take(f"{ads}/adTiers", "SKIP", "—", "廣告曝光方案清單（free、go）：結構")
    for k, lab, unit in (("exposureShare", "廣告曝光占比", "比例"), ("arpuPerExposedUserYear", "每曝光用戶年廣告收入", "$/用戶/年")):
        for i, y in enumerate(YEARS):
            I(f"{ads}/{k}", lab, index=str(y), unit=unit, v=f"{ads}/{k}/v/{i}", lo=f"{ads}/{k}/lo/{i}", hi=f"{ads}/{k}/hi/{i}",
              decision="N1b", note="Andy 2026-09-29：商業模式未成熟，2027→2030 不宜 8 倍；核心分析排除廣告" if (k == "exposureShare" and i == 0)
              else ("上限參照 Meta FY2025 全球 ARPP $57.03（10-K）" if (k != "exposureShare" and i == 0) else ""), meta=(i == 5))
    ci = f"{ads}/comparisonInternalTarget"
    for y in (2026, 2027, 2028, 2029, 2030):
        S(ci, f"內部廣告營收目標：{y}", v=f"{ci}/{y}", unit="$B", scope="投資人簡報；假設 2030 年 2.75B 週用戶；僅對照", date="2026-04-09",
          use="P2 對照（廣告）")
    for i, y in enumerate(YEARS):                    # reverse.comparisonSetAdsPlan 與上列同源（Axios 2026-04-09）
        tgt = f"reverse/comparisonSetAdsPlan/v/{i}"
        if y == 2025:
            R.assign(tgt, "DUP", "—", "2025 為基準年（值 0）：Axios 簡報未揭露 2025；不另列 SRC_OAI")
        else:
            R.assign(tgt, "DUP", f"SRC_OAI（內部廣告營收目標：{y}）", "與 other.ads.comparisonInternalTarget 同值同源")
    R.assign("reverse/comparisonSetAdsPlan/tag", "DUP", "SRC_OAI（內部廣告營收目標）", "同源")
    R.assign("reverse/comparisonSetAdsPlan/src", "DUP", "SRC_OAI（內部廣告營收目標）", "同源")
    I("tokenRevenue/other/residual", "其他營收（授權、一次性交易）", unit="$B", v="tokenRevenue/other/residual/v", note="維持 0")

    # 成本介面（Tokenomics v4 快照；v0.6 由 TK_Link 取代）
    R.take("tokenRevenue/costInterface/hwCostPerM", "TK", "TK_Link IF_HoldEcon、IF_TokGW_*",
           "不遷入；改取 TK_Link（分層優先於標記；補充 1 第 2 點）。v0.5 原值見對照表")
    R.assign("tokenRevenue/costInterface/overheadMultiplier", "SKIP", "—", "v0.4 已廢止（文字說明）")
    R.take("tokenRevenue/costInterface/capacityTokensT", "SKIP", "—", "N/A：D8a 容量模組建成前不適用；v0.6 P3 以 TK_Link 產能實作")

    # 對照欄
    cmp_ = "tokenRevenue/comparison"
    R.assign(f"{cmp_}/_note", "SKIP", "—", "說明文字（D4c：僅作對照欄）")
    S(f"{cmp_}/deckJul2026", "管理層營收目標：2026", v=f"{cmp_}/deckJul2026/2026", unit="$B", scope="7 月投資人簡報；2027–29 未揭露", date="2026-09-18",
      use="P5 反向模式")
    S(f"{cmp_}/deckJul2026", "管理層營收目標：2030", v=f"{cmp_}/deckJul2026/2030", unit="$B", scope="7 月投資人簡報", date="2026-09-18", use="P5 反向模式")
    S(f"{cmp_}/deckJul2026", "管理層營收目標：2026–2030 累計", v=f"{cmp_}/deckJul2026/sum2026_2030", unit="$B", scope="7 月投資人簡報", date="2026-09-18",
      use="P5 反向模式")
    R.take(f"{cmp_}/fy2026Independent", "FORMULA", "Revenue R41（P2 實作：FY2026 獨立推估＝Q1＋Q2＋ARR×下半年月數÷12）",
           "Derived（v0.5 標 Derived；組成值 Q1、Q2、ARR 已在 SRC_OAI）；區間 31–36 隨公式處理")
    S(f"{cmp_}/weeklyUsers2030Plan", "2030 週用戶數計畫", v=f"{cmp_}/weeklyUsers2030Plan/v", unit="M", scope="廣告預測之用戶假設", date="2026-04-09",
      use="P2 對照（廣告）")

    # 情境
    sc = "tokenRevenue/scenarios"
    R.assign(f"{sc}/_note", "INP", "scenarios", "說明文字（D11）")
    for s_ in ("low", "base", "high"):
        for key, lab, unit in (("elasticity.tasks", "價格彈性：任務數", "彈性"), ("elasticity.tokensPerTask", "價格彈性：每任務 token", "彈性"),
                               ("tokensPerTaskGrowth", "每任務 token 年增率", "比例/年")):
            I(f"{sc}/{s_}/{key}", f"情境 {s_}：{lab}", unit=unit, v=f"{sc}/{s_}/{key}", tag="Decision", decision="D11",
              note="情境軸（D11）", meta=False)
        for tier in ("top", "mid", "low"):
            I(f"{sc}/{s_}/listPriceAnnualChange/{tier}", f"情境 {s_}：牌價年變動率 {tier} 層", unit="比例/年",
              v=f"{sc}/{s_}/listPriceAnnualChange/{tier}", tag="Decision", decision="D11", note="情境軸（D11）", meta=False)

    # 五、推論成本（v0.5 roofline）：Derived 的 Tokenomics 快照不遷入；Assumed／Analogy 依規則遷入 Inputs
    ic = "tokenRevenue/inferenceCost"
    R.assign(f"{ic}/_note", "SKIP", "—", "說明文字")
    TKW = "不遷入；改取 TK_Link（分層優先於標記：屬 Tokenomics 領域；補充 1 第 2 點）。v0.5 原值見對照表"
    R.take(f"{ic}/rack", "TK", "TK_Link IF_Util（利用率）、IF_TokGW_*（產能）", TKW)
    R.take(f"{ic}/classes", "TK", "TK_Link IF_TokGW_Luna／Sol／Astra", TKW)
    R.take(f"{ic}/serving", "TK", "TK_Link IF_Util、IF_TokGW_*", TKW)
    us = f"{ic}/usage"
    R.assign(f"{us}/_note", "INP", "usage", "說明文字")
    for plan in ("free", "go", "plus", "pro", "seats"):
        for i, lab in enumerate(("luna 占比", "sol 占比", "astra 占比", "輸出占比（含思考）", "快取命中")):
            I(f"{us}/{plan}", f"方案模型組合：{plan}", index=lab, unit="比例", v=f"{us}/{plan}/{i}", tag=R.get(f"{us}/tag"), meta=False,
              status="待 Project 判斷（v0.5 分層成本假設；P3 起產能取自 TK_Link）")
    R.assign(f"{us}/tag", "INP", "usage 各列", "隨附資訊")
    R.assign(f"{us}/chk", "INP", "usage 各列", "隨附資訊")
    cal = f"{ic}/calibration2025"
    S(f"{cal}/inferenceTotal", "推論成本合計：2025", v=f"{cal}/inferenceTotal/v", unit="$B", scope="2025 推論支出（含付費與非付費用戶）", date="2026-02",
      use="P3 V2 η 的支出換算分子", note="T1(a)")
    S(f"{cal}/freeUsers", "推論成本：2025 非付費用戶", v=f"{cal}/freeUsers/v", unit="$B", scope="非付費用戶推論支出", date="2026-02", use="P4 對照")
    S(f"{cal}/paidUsers", "推論成本：2025 付費用戶", v=f"{cal}/paidUsers/v", unit="$B", scope="付費用戶推論支出", date="2026-02", use="P4 對照")
    I(f"{cal}/apiMargin", "API 毛利率（2025 校準）", unit="比例", v=f"{cal}/apiMargin/v", lo=f"{cal}/apiMargin/lo", hi=f"{cal}/apiMargin/hi", decision="T5")
    R.assign(f"{cal}/cogsReported/v", "DUP", "SRC_OAI（營業成本：FY2025）", "與 actuals.cogs2025 同值（7.5）同源（外流財報；FT/Fortune 2026-06）")
    for k in ("tag", "src", "chk"):
        R.assign(f"{cal}/cogsReported/{k}", "DUP", "SRC_OAI（營業成本：FY2025）", "同源")
    ck = f"{ic}/checks2026"
    S(f"{ck}/inference", "計畫推論成本：2026", v=f"{ck}/inference/v", unit="$B", scope="OpenAI 2026 預測；付費占比見下列",
      date="2026-02", use="P3 檢查", extra_leaves=(f"{ck}/inference/paidShare",), note="含付費占比 0.66（v0.5 同節點）")
    S(f"{ck}/inference2030", "計畫推論成本：2030", v=f"{ck}/inference2030/v", unit="$B", scope="OpenAI 2030 預測", date="2026-02",
      use="P3 檢查", extra_leaves=(f"{ck}/inference2030/paidShare", f"{ck}/inference2030/gmTarget"),
      note="含付費占比 0.94、毛利率目標 0.67（v0.5 同節點）")
    R.take(f"{ic}/nonText", "FORMULA", "P4（免費層非文字推論殘差）", "Derived 說明；P4 以公式實作")
    R.assign(f"{ic}/hwCostIndexSrc", "SKIP", "—", "文字說明：沿用 Tokenomics v4 艦隊硬體成本指數；v0.6 由 TK_Link 取代")

    # 六、反向模式、支出、融資、訓練、算力
    rv = "reverse"
    R.assign(f"{rv}/_note", "SKIP", "—", "說明文字（R1–R3）")
    R.assign(f"{rv}/targetVintage", "SRC", "反向目標各列（日期欄）", "簡報版本 Jul-2026：寫入各列口徑")
    tr_ = f"{rv}/targetRevenue"
    # targetRevenue 的結構：{"v":[…6…],"tag":…,"src":…}
    rid26 = next(r["id"] for r in R.src_rows if r["metric"] == "管理層營收目標：2026")
    rid30 = next(r["id"] for r in R.src_rows if r["metric"] == "管理層營收目標：2030")
    R.assign(f"{tr_}/v/1", "DUP", rid26, "E8e：與 comparison.deckJul2026.2026 同一來源（FT 2026-09-18），合併為一組 SRC 列；reverse 引用該列")
    R.assign(f"{tr_}/v/5", "DUP", rid30, "E8e：與 comparison.deckJul2026.2030 同一來源，合併；reverse 引用該列")
    R.assign(f"{tr_}/tag", "DUP", rid26, "同源")
    R.assign(f"{tr_}/src", "DUP", rid26, "同源")
    R.assign(f"{tr_}/v/0", "DUP", "SRC_OAI（營收：FY2025）", "2025 為實際值（同 actuals.revenue2025）")
    # P5（S5）：沿用 v0.5／Tokenomics v4 FT_Sep2026 列 24 的結構——2027、2028 為擬合值（線性遞減成長，閉合 36→350 與 Σ840），2029＝Σ840 殘差（公式）
    for i, y in ((2, 2027), (3, 2028)):
        I(tr_, "管理層營收目標擬合值（反向模式）", index=str(y), unit="$B", v=f"{tr_}/v/{i}", tag="Decision", decision="R1",
          status="P5 新增（反向模式；S5）", meta=False,
          note="Derived 擬合值（Tokenomics v4 FT_Sep2026 列 24：成長率線性遞減，閉合 2026 $36B→2030 $350B 且 2026–30 合計 $840B；v0.5 沿用）。"
               "只驅動 Reverse 頁（R1），不回饋基準。替代擬合：二次式 77.0／143.0（2029 殘差 234.0）")
        R.inp_rows[-1]["lo"] = R.inp_rows[-1]["hi"] = R.get(f"{tr_}/v/{i}")
    R.assign(f"{tr_}/v/4", "FORMULA", "Reverse X01（2029＝Σ840 − 2026 − 2027 − 2028 − 2030；P5）",
             "Derived：v0.5／Tokenomics v4 以 2029 為殘差使 2026–30 合計＝$840B（v0.5 標 Interested-party/Derived 混合；依元素拆分）")
    S(f"{rv}/targetFCFcum2026_2030", "管理層目標：2026–2030 累計自由現金流", v=f"{rv}/targetFCFcum2026_2030/v", unit="$B",
      scope="Jul-2026 簡報", date="2026-09-18", use="P5 反向模式")
    # 計畫算力：reverse.targetComputeCum、anchors.computeCum_FT、training.planCompute 同源（FT 2026-09-18）
    pc = "training/planCompute2026_2030"
    rid_pc = S(pc, "計畫算力支出：2026–2030 累計", v=f"{pc}/v", lo=f"{pc}/lo", hi=f"{pc}/hi", unit="$B", scope="FT 2026-09-18（7 月簡報）$856B；Reuters 2026-02 約 $600B 為下限",
               date="2026-09-18", use="P3 檢查（計畫算力）；P5 反向目標",
               note="併入 reverse.targetComputeCum2026_2030 與 spending.anchors.computeCum2026_2030_FT（同值同源）")
    R.assign(f"{rv}/targetComputeCum2026_2030/v", "DUP", rid_pc, "與 training.planCompute2026_2030 同值同源（FT 2026-09-18）")
    R.assign(f"{rv}/targetComputeCum2026_2030/tag", "DUP", rid_pc, "同源")
    R.assign(f"{rv}/targetComputeCum2026_2030/src", "DUP", rid_pc, "同源")
    R.assign("spending/anchors/computeCum2026_2030_FT/v", "DUP", rid_pc, "與 training.planCompute2026_2030 同值同源（FT 2026-09-18）")
    R.assign("spending/anchors/computeCum2026_2030_FT/tag", "DUP", rid_pc, "同源")
    R.assign("spending/anchors/computeCum2026_2030_FT/src", "DUP", rid_pc, "同源")
    pv = f"{rv}/priorVintages"
    S(pv, "管理層營收目標（舊版）：2030（2026-02 版）", v=f"{pv}/Feb-2026/2030", unit="$B", scope="2026-02 版本的 2030 營收目標；只對照", date="2026-02",
      use="P5 對照")
    R.take(f"{rv}/sufficientSets", "SKIP", "—", "反向模式充分條件組合（apiOnly／subOnly／proportional）：P5 以公式結構實作（R1–R3）")
    R.assign(f"{rv}/adsInReverse", "SKIP", "—", "決策記錄：反向模式廣告列保持正向（N1b）；P5 以公式結構體現")

    sp = "spending"
    R.assign(f"{sp}/_note", "SKIP", "—", "說明文字（S1c、S2、N4(b)）")
    I(f"{sp}/unscheduledProfile", "未揭露時程合約的攤提方式", unit="文字（even／ramp）", vlit=R.get(f"{sp}/unscheduledProfile/v"),
      decision="N4(b)", tag="Decision", note="期間內線性爬升（Andy 2026-09-29）；2026 算力現金 $53B 符合證詞 $50B", meta=False)
    R.assign(f"{sp}/unscheduledProfile/v", "INP", f"INP_{len(R.inp_rows):03d}")
    for k in ("tag", "src"):
        R.assign(f"{sp}/unscheduledProfile/{k}", "INP", f"INP_{len(R.inp_rows):03d}", "隨附資訊")
    R.take(f"{sp}/unscheduledProfile/options", "SKIP", "—", "選項清單（even／ramp）：結構")

    ct = f"{sp}/contracts"
    # Microsoft Azure
    az = f"{ct}/microsoftAzure"
    S(az, "合約總額：Microsoft Azure（OpenAI 增購）", v=f"{az}/total", unit="$B", scope="2025-10 公告增購 Azure 承諾", date="2025-10", subject="Microsoft",
      tag="Verified", source="OpenAI／Microsoft 2025-10 公告增購 $250B Azure（v0.5：總額 [Verified]）", use="P3 供給；P4 成本")
    I(az, "合約期間起：Microsoft Azure", unit="年", v=f"{az}/start", tag="Analogy", decision="V16", status="v0.6 沿用",
      note="2025-10 公告；實際付款可能延至 2026（V16）", meta=False)
    R.inp_rows[-1]["lo"], R.inp_rows[-1]["hi"] = 2025, 2026
    I(az, "合約期間迄：Microsoft Azure", unit="年", v=f"{az}/end", tag="Assumed", decision="V13",
      note="V13：終點 2030、區間 2029–2032、Assumed", meta=False)
    R.inp_rows[-1]["lo"], R.inp_rows[-1]["hi"] = 2029, 2032
    R.assign(f"{az}/annual", "SKIP", "—", "null：未揭露年度金額（P3 由 N4(b) 爬升決定）")
    # Oracle
    oc = f"{ct}/oracle"
    S(oc, "合約總額：Oracle", v=f"{oc}/total", unit="$B", scope="5 年、2027 起、約 4.5GW", date="2025-09", subject="Oracle", use="P3 供給；P3 V2 合約價")
    S(oc, "合約期間起：Oracle", v=f"{oc}/start", unit="年", scope="Oracle 公告", date="2025-09", subject="Oracle", use="P3")
    S(oc, "合約期間迄：Oracle", v=f"{oc}/end", unit="年", scope="5 年（2027–2031）", date="2025-09", subject="Oracle", use="P3")
    S(oc, "合約容量：Oracle", v=f"{oc}/gw", unit="GW", scope="約 4.5GW；電力口徑未明", date="2025-09", subject="Oracle", use="P3 V2 合約價")
    R.assign(f"{oc}/annual", "FORMULA", "Checks C14（E8f：Oracle annual＝total÷年數）", "60＝300÷5，算術推得（v0.5 以數值寫入）")
    # AWS Nvidia
    an = f"{ct}/awsNvidia"
    S(an, "合約總額：AWS（Nvidia 晶片）", v=f"{an}/total", unit="$B", scope="AWS／OpenAI 公告；7 年", date="2025-11-03", subject="AWS", use="P3；P4")
    S(an, "合約期間起：AWS（Nvidia 晶片）", v=f"{an}/start", unit="年", scope="同上", date="2025-11-03", subject="AWS", use="P3")
    S(an, "合約期間迄：AWS（Nvidia 晶片）", v=f"{an}/end", unit="年", scope="7 年", date="2025-11-03", subject="AWS", use="P3")
    R.assign(f"{an}/annual", "SKIP", "—", "null：未揭露年度金額")
    # AWS Trainium
    at = f"{ct}/awsTrainium"
    S(at, "合約總額：AWS（Trainium）", v=f"{at}/total", unit="$B", scope="AWS 追加；8 年", date="2026-02", subject="AWS", use="P3；P4")
    S(at, "合約期間起：AWS（Trainium）", v=f"{at}/start", unit="年", scope="同上", date="2026-02", subject="AWS", use="P3")
    S(at, "合約期間迄：AWS（Trainium）", v=f"{at}/end", unit="年", scope="8 年", date="2026-02", subject="AWS", use="P3")
    S(at, "合約容量：AWS（Trainium）", v=f"{at}/gw", hi=None, unit="GW", scope="2026-02 公告 2GW Trainium", date="2026-02", subject="AWS",
      use="P3 供給", note=R.get(f"{at}/gwNote"))
    # E8g：WSJ 2026-07 揭露登兩筆（推論 3GW、訓練 2GW），「5」改公式；Inputs 採用值維持 2（2–5）
    S(f"{at}/src文字（E8g）", "合約容量（WSJ 口徑）：AWS Trainium 推論", vlit=3, unit="GW", scope="WSJ 2026-07：3GW 推論", date="2026-07", subject="AWS",
      status="Alt", alt_of="見『合約容量：AWS（Trainium）』（2026-02 公告 2GW）", tag="Interested-party", source="WSJ 2026-07（v0.5 gwNote 轉述）",
      use="P3 區間上緣（與下列訓練 2GW 相加＝5，Checks 公式）", note="揭露口徑衝突：公告 2GW 對 WSJ 3＋2")
    S(f"{at}/src文字（E8g）", "合約容量（WSJ 口徑）：AWS Trainium 訓練（VR）", vlit=2, unit="GW", scope="WSJ 2026-07：2GW 訓練（VR）", date="2026-07", subject="AWS",
      status="Alt", alt_of="見『合約容量：AWS（Trainium）』（2026-02 公告 2GW）", tag="Interested-party", source="WSJ 2026-07（v0.5 gwNote 轉述）",
      use="P3 區間上緣", note="揭露口徑衝突：公告 2GW 對 WSJ 3＋2")
    I(f"{at}/r3採用值", "合約容量（採用值）：AWS（Trainium）", unit="GW", vlit=2, tag="Assumed", decision="E8g",
      note="E8g：維持 2，區間 2–5（上緣 5＝WSJ 推論 3＋訓練 2，Checks 公式驗證）", meta=False)
    R.inp_rows[-1]["lo"], R.inp_rows[-1]["hi"] = 2, 5
    R.assign(f"{at}/annual", "SKIP", "—", "null：未揭露年度金額")
    # CoreWeave
    cw = f"{ct}/coreweave"
    S(cw, "合約總額：CoreWeave", v=f"{cw}/total", unit="$B", scope="$11.9B＋$4B＋$6.5B；日後可改接 CRWV 模型之 OpenAI 時程", date="2025", subject="CoreWeave",
      use="P3；P4")
    S(cw, "合約期間起：CoreWeave", v=f"{cw}/start", unit="年", scope="CoreWeave 揭露", date="2025", subject="CoreWeave", use="P3")
    S(cw, "合約期間迄：CoreWeave", v=f"{cw}/end", unit="年", scope="CoreWeave 揭露", date="2025", subject="CoreWeave", use="P3")
    R.assign(f"{cw}/annual", "SKIP", "—", "null：未揭露年度金額")
    # Cerebras
    ce = f"{ct}/cerebras"
    S(ce, "合約總額：Cerebras", v=f"{ce}/total", lo=f"{ce}/lo", unit="$B", tag="Verified", scope="總值逾 $20B（下限）；10-Q／8-K", date="2026", subject="Cerebras", use="P3；P4",
      note="v0.5 total＝lo＝20；上緣 25 見 Inputs（Analogy）")
    I(ce, "合約總額（估計）：Cerebras", unit="$B", vlit=20, hi=f"{ce}/hi", tag="Assumed", decision="V13",
      note="V13：總額 20、區間 20–25、Assumed（SRC_OAI 另有觀測『逾 $20B』）；上限 25 為 v0.5 hi 葉節點", meta=False)
    R.inp_rows[-1]["lo"] = 20
    S(ce, "合約期間起：Cerebras", v=f"{ce}/start", unit="年", tag="Verified", scope="2026–28 分批部署", date="2026", subject="Cerebras", use="P3")
    I(ce, "合約期間迄：Cerebras", unit="年", v=f"{ce}/end", tag="Assumed", decision="V13",
      note="V13：終點 2031、區間 2029–2032、Assumed（v0.5：期間終點為推估，每批 3–4 年）", meta=False)
    R.inp_rows[-1]["lo"], R.inp_rows[-1]["hi"] = 2029, 2032
    S(ce, "合約容量：Cerebras", v=f"{ce}/gw", unit="GW", tag="Verified", scope="750MW", date="2026", subject="Cerebras", use="P3")
    R.assign(f"{ce}/annual", "SKIP", "—", "null：未揭露年度金額")
    for key in ("microsoftAzure", "oracle", "awsNvidia", "awsTrainium", "coreweave", "cerebras"):
        for k in ("tag", "src"):
            p = f"{ct}/{key}/{k}"
            if p not in R.dest:
                R.assign(p, "SRC", "合約各列", "隨附資訊")
    exc = f"{sp}/excludedOrComparison"
    cmp_note = "比較用，不計入現金流（S4b：避免與股權重複）"
    h1 = [
        ("Nvidia 意向：10GW", 10, "GW", "Nvidia", "2025-09-22",
         "OpenAI 與 NVIDIA 2025-09-22 聯合公告（意向書，非約束性）",
         "意向書，非約束性；" + cmp_note, "協議性質：非約束性意向書", (f"{exc}/nvidia10GW",)),
        ("Nvidia 意向：$100B", 100, "$B", "Nvidia", "2025-09-22",
         "OpenAI 與 NVIDIA 2025-09-22 聯合公告（意向書，非約束性）",
         "意向書，非約束性；$100B 為 NVIDIA 對 OpenAI 的投資上限，隨每 GW 部署逐步投入，不是 OpenAI 的算力支出；" + cmp_note, "協議性質：非約束性意向書", ()),
        ("Broadcom 晶片：10GW", 10, "GW", "Broadcom", "2025-10-13",
         "OpenAI 與 Broadcom 2025-10-13 聯合公告（自研加速器部署，條款書）",
         "自研加速器部署；條款書；" + cmp_note, "協議性質：條款書", (f"{exc}/broadcomAmdChips",)),
        ("AMD 晶片：6GW", 6, "GW", "AMD", "2025-10-06",
         "OpenAI 與 AMD 2025-10-06 聯合公告（最終協議，附 AMD 認股權證）",
         "最終協議，附 AMD 認股權證；" + cmp_note, "協議性質：最終協議，附 AMD 認股權證", ()),
    ]
    for metric, val, unit, subj, dt, src_, scope_, nature, extra in h1:
        S(f"{exc}/src文字（E8h）", metric, vlit=val, unit=unit, scope=scope_, date=dt, subject=subj, tag="未標（v0.5 決策文字）",
          source=src_, use="比較用", note=cmp_note + "；v0.5 excludedOrComparison 決策文字", extra_leaves=extra)
        # H1：一手、利害關係方；立場說明沿用 G1 文字並加註協議性質；等級比照 SRC_OAI_062（Verified 合約列＝2）
        R.src_rows[-1]["cl"] = dict(grade="2", stance="利害關係方", why=source_rules.CONTRACT_WHY + "；" + nature, hand="一手")
    R.assign(f"{exc}/stargateSoftBank", "SKIP", "—", "決策記錄（S3）：合資建設由合作方出資部分不進 OpenAI 現金流；無揭露數字")
    oc2 = f"{sp}/ownedCapex"
    for i, y in enumerate(YEARS):
        I(oc2, "自有資本支出", index=str(y), unit="$B", v=f"{oc2}/v/{i}", note="例：Project Camellia $20B（TechCrunch 2026-07-22）" if i == 1 else "", meta=False)
    I(oc2, "自有資本支出：低情境倍數", unit="倍", v=f"{oc2}/lo_mult", meta=False)
    I(oc2, "自有資本支出：高情境倍數", unit="倍", v=f"{oc2}/hi_mult", meta=True)
    I(f"{sp}/partnerFinancedShare", "合作方出資占比（S3 槓桿）", unit="比例", v=f"{sp}/partnerFinancedShare/v", lo=f"{sp}/partnerFinancedShare/lo",
      hi=f"{sp}/partnerFinancedShare/hi", decision="S3", note="沿用 v4 槓桿；情境軸")
    R.take(f"{sp}/leases", "SKIP", "—", "N/A：OpenAI 未揭露承租合約；雲端合約內含之租賃不另列")
    ox = f"{sp}/opexExComputePctRevenue"
    R.assign(f"{ox}/v/0", "FORMULA", "P4（2025 非算力營運費用占營收比：由財報數據推得）",
             "V9：2025 由財報數據推得（研發不含付 Microsoft、銷售、管理 ÷ 營收）→公式；v0.5 值 1.216 供對照（原 Interested-party／Derived 混合）")
    for i, y in enumerate(YEARS):
        if i == 0:
            continue
        I(ox, "非算力營運費用占營收比（v0.5）", index=str(y), unit="比例", v=f"{ox}/v/{i}", decision="S5a",
          status="P4 將取代（V4：人數×每人成本）；v0.5 值暫留作暫代", meta=False)
    I(ox, "非算力營運費用占營收比：低情境倍數", unit="倍", v=f"{ox}/lo_mult", decision="S5a", status="P4 將取代", meta=False)
    I(ox, "非算力營運費用占營收比：高情境倍數", unit="倍", v=f"{ox}/hi_mult", decision="S5a", status="P4 將取代", meta=True)
    R.split("spending.opexExComputePctRevenue", "2025 元素 1.216 由財報數據推得→公式；2026–2030（Assumed）→Inputs 並標『P4 將取代』", "公式 1 格＋Inputs 5 格＋倍數 2 格")
    an2 = f"{sp}/anchors"
    S(f"{an2}/computeCum2026_2030_WSJ", "計畫算力支出：2026–2030 累計（WSJ 口徑）", v=f"{an2}/computeCum2026_2030_WSJ/v", unit="$B", scope="WSJ 2026-07-22",
      date="2026-07-22", use="P3 檢查；P5")
    S(f"{an2}/offBalanceCommitmentsStock", "表外承諾存量", v=f"{an2}/offBalanceCommitmentsStock/v", unit="$B", scope="S-1 草案轉述（Tokenomics v1_3 Sources #19）",
      use="P5 對照")
    S(f"{an2}/compute2026Testimony", "2026 算力支出（證詞）", v=f"{an2}/compute2026Testimony/v", unit="$B", scope="Brockman 證詞（commitments tracker 2026-05 轉述）",
      date="2026-05", use="P4 對照")

    fd = "funding"
    R.assign(f"{fd}/_note", "SKIP", "—", "說明文字（S7、S4b）")
    S(f"{fd}/cashEnd2025", "現金：2025 年底", v=f"{fd}/cashEnd2025/v", lo=f"{fd}/cashEnd2025/lo", hi=f"{fd}/cashEnd2025/hi", unit="$B",
      scope="2025-12 約 $40B；2026 Q1 末 $73B（The Information 經 TNW 轉述）", use="P5 融資起點")
    er = f"{fd}/equityRound2026Mar"
    S(er, "股權融資：2026-03 輪總額", v=f"{er}/total", unit="$B", scope="2026-03 輪（headline）", date="2026-03", tag="Interested-party",
      source="OpenAI 公告（2026-03 融資輪；v0.5 該葉節點無出處，依 E8i 補記）", use="P5")
    ur = f"{er}/unconditional2026"
    usrc = R.get(f"{ur}/src")
    for lab, val, extra in (("Amazon 首筆", 15, ""), ("SoftBank（三期）", 30, ""), ("Nvidia", 30, "；Nvidia 以現金計為 v0.5 的處理（Inputs『Nvidia $30B 現金比例』1），另有報導稱多為算力"), ("其他", 12, "（約）")):
        S(f"{ur}/src文字（E8j）", f"股權融資 2026-03 輪：{lab}", vlit=val, unit="$B", scope="2026 無條件部分的組成" + extra, date="2026-03", tag="Interested-party",
          source=usrc, use="P5（組成加總＝v0.5 無條件部分 87）")
    R.take(ur, "FORMULA", "Derived_V9／Checks（SRC 四個組成加總）", "E8j：87 拆為 SRC 組成（Amazon 15、SoftBank 30、Nvidia 30、其他 12）；加總驗證＝v0.5 值 87")
    I(f"{fd}/r3Nvidia現金", "Nvidia $30B 現金比例", unit="比例", vlit=1, tag="Assumed", decision="S4b",
      note="E8j：Nvidia $30B 以現金計的比例 1（0–1）；沿用 v0.5 S4b（另有報導稱多為算力）", meta=False)
    R.inp_rows[-1]["lo"], R.inp_rows[-1]["hi"] = 0, 1
    S(f"{er}/conditional", "股權融資：條件式部分（Amazon）", v=f"{er}/conditional/v", unit="$B", scope="Amazon 其餘以 IPO 或 AGI 里程碑為條件；base 不計入", use="P5")
    R.src_rows[-1]["cl"] = dict(grade="2", stance="未評", why="原始發布者不明（多家報導）；base 不計入，P5 情境使用時再評", hand="二手")   # H4
    I(f"{fd}/minCash", "最低現金", unit="$B", v=f"{fd}/minCash/v", lo=f"{fd}/minCash/lo", hi=f"{fd}/minCash/hi", decision="S7")
    cl = f"{fd}/contingentLiabilities"
    S(f"{cl}/nvidiaGuarantees", "或有負債：Nvidia 擔保", v=f"{cl}/nvidiaGuarantees/v", unit="$B", scope="WSJ（Tokenomics v1_3 Sources #21）", use="P5（S4b 只列示）")
    S(f"{cl}/nvidiaChipFinancingTalks", "或有負債：Nvidia 晶片融資洽談", v=f"{cl}/nvidiaChipFinancingTalks/v", unit="$B", scope="WSJ；洽談中", use="P5（S4b 只列示）")
    R.assign(f"{cl}/treatment", "SKIP", "—", "決策記錄（S4b）：只列示，不作資金來源；P5 以公式結構體現")
    R.assign(f"{fd}/externalNeed", "SKIP", "—", "外部資金需求的定義（年底現金低於最低現金之差額累計）：P5 以公式實作")

    tn = "training"
    R.assign(f"{tn}/_note", "SKIP", "—", "說明文字")
    S(f"{tn}/actual2024", "訓練算力支出：2024", v=f"{tn}/actual2024/v", unit="$B", scope="Epoch（引 The Information/NYT）：2024 訓練算力 $3B、推論 $1.8B", use="P3 對照")
    S(f"{tn}/actual2025/src文字（r1 V6）", "研發費用中付 Microsoft 部分：2025", vlit=10.59, unit="$B", scope="2025 R&D 付予 Microsoft（外流財報）",
      source="外流財報（v0.5 training.actual2025 出處：2025 R&D 付予 Microsoft $10.59B）", tag="Interested-party", use="P3 研發 GW；P4 V4（訓練支出＝10.59＋其他雲端）",
      note="V6：訓練支出 12.0 拆開，付 Microsoft 的 10.59 進 SRC_OAI；其他雲端 1.41（0–2.91）見 Inputs；12.0（10.59–13.50）＝兩者相加的公式（Checks 衍生值列）")
    R.take(f"{tn}/actual2025", "FORMULA", "訓練支出 2025＝SRC_OAI 10.59＋Inputs 其他雲端（Checks 衍生值列）",
           "V6：Derived→公式；v0.5 值 12.0（10.6–13.5）供對照")
    S(f"{tn}/rd2025", "研發費用總額：2025", v=f"{tn}/rd2025/v", unit="$B", scope="外流財報 R&D 總額（含人事）", use="P4 V4")
    # planCompute：已在上方登記；planInference 併入 checks2026
    R.assign(f"{tn}/planInference/2026", "DUP", "SRC_OAI（計畫推論成本：2026）", "與 checks2026.inference.v（14.1）同值同源")
    R.assign(f"{tn}/planInference/2030", "DUP", "SRC_OAI（計畫推論成本：2030）", "與 checks2026.inference2030.v（85）同值同源")
    R.assign(f"{tn}/planInference/tag", "DUP", "SRC_OAI（計畫推論成本）", "同源")
    R.assign(f"{tn}/planInference/src", "DUP", "SRC_OAI（計畫推論成本）", "同源；2027–29 幾何內插為 Derived（P3 公式）")
    for k in ("tag", "src", "chk"):
        p = f"{tn}/planCompute2026_2030/{k}"
        if p not in R.dest:
            R.assign(p, "SRC", rid_pc, "隨附資訊")
    R.assign(f"{tn}/allocation", "SKIP", "—", "方法說明（計畫訓練總額依各年算力現金−推論成本占比分配）：P3／P4 以公式實作")
    R.assign(f"{tn}/tokenAllocation", "SKIP", "—", "方法說明（每層級訓練分攤）：P4 以公式實作")

    cm = "computeMW"
    R.assign(f"{cm}/_note", "SKIP", "—", "說明文字：單位 GW（IT 電力）；VR-eq 概念由 TK_Link 世代產能取代")
    R.take(f"{cm}/tokenomicsRef", "TK", "TK_Link IF_TokGW_*、IF_HoldEcon", "不遷入；改取 TK_Link（補充 1 第 2 點）。v0.5 原值見對照表")
    R.take(f"{cm}/genFactor/hopper", "TK", "TK_Link IF_TokGW_*（Hopper 世代欄）",
           "不遷入；改取 TK_Link（分層優先於標記；補充 1 第 2 點）。v0.5 原值見對照表")
    R.take(f"{cm}/genFactor/nextGen", "TK", "TK_Link IF_TokGW_*（Rubin Ultra 世代欄）",
           "不遷入；改取 TK_Link（分層優先於標記；補充 1 第 2 點）。v0.5 原值見對照表")
    I(f"{cm}/genFactor/custom", "世代產能係數：自研／其他（相對 VR200）", unit="倍", v=f"{cm}/genFactor/custom/v", lo=f"{cm}/genFactor/custom/lo",
      hi=f"{cm}/genFactor/custom/hi", status="P3 沿用（自研／其他產能＝VR200 × 此係數）", note="Cerebras、Broadcom、AMD、Trainium")
    fm = f"{cm}/fleetMix"
    R.assign(f"{fm}/_note", "INP", "fleetMix", "說明文字：實體 GW 世代占比；自研／其他＝1−其餘")
    for g_, label in (("hopper", "Hopper"), ("blackwell", "Blackwell"), ("veraRubin", "Vera Rubin"), ("nextGen", "下一代")):
        for i, y in enumerate(YEARS):
            I(f"{fm}/{g_}", f"實體 GW 世代占比：{label}", index=str(y), unit="比例", v=f"{fm}/{g_}/{i}", decision="V1",
              note="P3 將 Blackwell 拆為 GB200／GB300（2025：GB200 100%；2026 起 50／50；工作單 P3-2）" if (g_ == "blackwell" and i == 0) else "",
              status="P1 照 v0.5 遷入；P3 依 V1 改世代拆分", meta=False)
    for k in ("tag", "chk"):
        R.assign(f"{fm}/{k}", "INP", "fleetMix 各列", "隨附資訊")
    ac = f"{cm}/actuals"
    S(f"{ac}/gwEnd2023", "算力規模：2023 年底", v=f"{ac}/gwEnd2023", unit="GW", scope="CFO 部落格；口徑未明（IT／設施未標）", date="2026-01-18", use="P3 總需求 GW 對照")
    S(f"{ac}/gwEnd2024", "算力規模：2024 年底", v=f"{ac}/gwEnd2024", unit="GW", scope="同上", date="2026-01-18", use="P3 對照（工作單 P3-5：0.6）")
    S(f"{ac}/gwEnd2025", "算力規模：2025 年底", v=f"{ac}/gwEnd2025", unit="GW", scope="同上（約 1.9）", date="2026-01-18", use="P3 對照（工作單 P3-5：1.9）")
    S(f"{ac}/arrEnd2025", "年化營收（ARR）：2025 年底", v=f"{ac}/arrEnd2025", unit="$B", scope="同上（$20B+）", date="2026-01-18", use="P2 對照")
    for k in ("tag", "src", "chk"):
        R.assign(f"{ac}/{k}", "SRC", "算力規模各列", "隨附資訊")
    lp = f"{cm}/leasePricePerGWyr"
    I(lp, "每 GW 年合約價（leasePricePerGWyr）", unit="$B/GW/年", v=f"{lp}/v", lo=f"{lp}/lo", hi=f"{lp}/hi", tag="Analogy", decision="V8",
      status="V8：本模型唯一的合約價參數（P3 V2 使用）",
      note="基準 12、區間 8–20（v0.5 上限 16 依 V8 放寬為 20）。可比：Oracle 遠期合約隱含 13.3；差異：交付年份與世代；另參 2025 合約現金÷平均 GW≈11.8")
    R.inp_rows[-1]["hi"] = 20
    I("r1/新增/其他雲端", "2025 其他雲端研發算力支出（訓練支出的非 Microsoft 部分）", unit="$B", vlit=1.41, tag="Assumed", decision="V6",
      status="V6 新增", note="訓練支出 12.0＝SRC_OAI 10.59＋本列 1.41；區間 0–2.91（＝13.50−10.59）", meta=False)
    R.inp_rows[-1]["lo"], R.inp_rows[-1]["hi"] = 0, 2.91
    I(f"{cm}/idleReserve", "閒置保留比例", unit="比例", v=f"{cm}/idleReserve/v")
    I(f"{cm}/minTrainingShare", "訓練最低占比", unit="比例", v=f"{cm}/minTrainingShare/v", decision="V1a",
      note="V1a 容量上限：推論可用＝VR-eq 供給×(1−閒置)×(1−訓練最低占比)")
    R.assign(f"{cm}/decisions", "SKIP", "—", "決策記錄（V1a、V2a）：P3、P5 以公式結構體現；v0.5 決策 ID 保留於 Inputs 欄")

    # V9 掃描：其餘多年期項目（出處文字內的觀測錨點新增為 SRC_OAI 列；元素維持原分類）
    S("spending/ownedCapex/src文字（V9 掃描）", "自有園區資本支出例：Project Camellia", vlit=20, unit="$B", scope="自有園區與晶片（例）",
      date="2026-07-22", source="TechCrunch 2026-07-22（v0.5 ownedCapex 出處）", tag="Interested-party", use="P4 自有資本支出 2026 的參照（錨點）",
      note="v0.5 出處為『例』；2026 元素 20 為 Assumed 整體額，不等於此單一專案")
    S("tokenRevenue/subscription/tasksPerDay/src文字（V9 掃描）", "ChatGPT 每日訊息數：2025-07", vlit=2.5, unit="B 則/日", scope="每日 2.5B 則訊息（總量對照）",
      date="2025-07", source="OpenAI（v0.5 tasksPerDay 出處：總量對照）", tag="Interested-party", use="P2 每日任務數的總量對照")
    R.split("spending.ownedCapex", "出處『Project Camellia $20B（TechCrunch 2026-07-22）』為單一專案例，非整年觀測；新增 SRC_OAI 錨點列，6 個元素維持 Inputs（Assumed）", "SRC_OAI 1 列（錨點）＋Inputs 6 格", scanned_only=True)
    R.split("subscription.tasksPerDay／tokensPerTask", "出處（Tokenomics workbook，未經查核）；僅『2025-07 每日 2.5B 則訊息』為可引用觀測，與任務口徑不同；新增 SRC_OAI 錨點列，元素維持 Inputs", "SRC_OAI 1 列（錨點）＋Inputs 20 格", scanned_only=True)
    R.split("api.listPriceFYAvg", "V7 已依元素拆分：2025（Analogy）→Inputs；2026（由事件價推得）→公式", "Inputs 與公式（見分類規則 2）")
    R.split("reverse.targetRevenue", "V7 已依元素拆分：2026、2030→SRC_OAI；2025 併入實際營收；2027–28→Inputs（擬合值，P5）；2029→公式（Σ840 殘差，P5）", "SRC_OAI 2 列＋重複併入 1＋Inputs 2＋公式 1")
    R.split("other.ads.arpuPerExposedUserYear", "出處『Meta FY2025 全球 ARPP $57.03（10-K）』是第三方上限參照，非 OpenAI 觀測；元素全屬 Analogy→Inputs，不拆", "Inputs 6 格", scanned_only=True)
    R.split("other.ads.exposureShare、demandGrowthExPrice、fleetMix、arpu、tokenMix、listPriceAnnualChange", "出處皆無年度觀測值（Assumed／Analogy 全期）→不拆", "Inputs", scanned_only=True)
    R.split("api.demandGrowthExPrice.tasks（2026）", "2026 為校準年（v0.4 把 53%→134% 使 FY2026 對上獨立推估 32.9）：值由校準得出，屬 Assumed、無觀測元素→不拆；校準目標 32.9 為 Derived（公式）", "Inputs", scanned_only=True)
    p2_rows(R)
    p3_rows(R)
    p4_rows(R)
    p5_rows(R)
    # 殘餘：補上未登記的中介資料葉節點（tag/src/chk/_note…）隨最近的已登記兄弟
    for p in list(R.leaves):
        if p in R.dest:
            continue
        leaf = p.split(SEP)[-1]
        if leaf in META_KEYS:
            parent = SEP.join(p.split(SEP)[:-1])
            sib = next((R.dest[q] for q in R.leaves if q in R.dest and q.startswith(parent + SEP)), None)
            if sib:
                R.assign(p, sib[0], sib[1], "隨附資訊")
    R.traced = source_rules.trace_sources(R.src_rows)      # H2：「同上」追溯為實際出處
    finalize_ids(R, write_registry)
    return R


def p2_rows(R: Recorder):
    """v0.6-P2（S2）新增列：價格事件的數值（SRC_OAI，由 SRC_OAI_025–029 文字列拆出）、事件時點 Inputs（E8c）、
    定義常數（E6：公式不得內含常數）。不改 v0.5 葉節點的去處。"""
    S, I = R.src, R.inp
    ev = "tokenRevenue/api/priceEvents2026/數值拆分（P2 E8c）"
    src_txt = R.get("tokenRevenue/api/priceEvents2026/src")
    base = dict(unit="$/M tokens", tag="Verified", source=src_txt, use="P2 FY2026 平均牌價（事件時點天數加權）")
    for metric, val, scope, dt in (
        ("API 價格事件 1 牌價：top 層 輸入（GPT-5.6 Sol）", 5, "07-09 GPT-5.6 Sol 輸入；N3a 層級角色＝top（GPT-6 Astra 上市前）", "2026-07-09"),
        ("API 價格事件 1 牌價：top 層 輸出（GPT-5.6 Sol）", 30, "07-09 GPT-5.6 Sol 輸出；層級角色＝top", "2026-07-09"),
        ("API 價格事件 1 牌價：mid 層 輸入（GPT-5.6 Terra）", 2.5, "07-09 GPT-5.6 Terra 輸入；層級角色＝mid", "2026-07-09"),
        ("API 價格事件 1 牌價：mid 層 輸出（GPT-5.6 Terra）", 15, "07-09 GPT-5.6 Terra 輸出；層級角色＝mid", "2026-07-09"),
        ("API 價格事件 1 牌價：low 層 輸入（GPT-5.6 Luna）", 1, "07-09 GPT-5.6 Luna 輸入；層級角色＝low", "2026-07-09"),
        ("API 價格事件 1 牌價：low 層 輸出（GPT-5.6 Luna）", 6, "07-09 GPT-5.6 Luna 輸出；層級角色＝low", "2026-07-09"),
        ("API 價格事件 2 牌價：mid 層 輸入（Terra 降價）", 2, "07-30 Terra 降價後輸入（−20%）", "2026-07-30"),
        ("API 價格事件 2 牌價：mid 層 輸出（Terra 降價）", 12, "07-30 Terra 降價後輸出（−20%）", "2026-07-30"),
        ("API 價格事件 2 牌價：low 層 輸入（Luna 降價）", 0.2, "07-30 Luna 降價後輸入（−80%）", "2026-07-30"),
        ("API 價格事件 2 牌價：low 層 輸出（Luna 降價）", 1.2, "07-30 Luna 降價後輸出（−80%）", "2026-07-30"),
        ("API 價格事件 3 牌價：top 層 輸入（Sol 促銷）", 4, "08-22 Sol 促銷輸入", "2026-08-22"),
        ("API 價格事件 3 牌價：top 層 輸出（Sol 促銷）", 20, "08-22 Sol 促銷輸出", "2026-08-22"),
    ):
        S(ev, metric, vlit=val, scope=scope, date=dt, note="由 SRC_OAI 價格事件文字列拆出的數值（同一出處）；事件 4、5 之後的價格＝2026-09-25 牌價快照列", **base)
    for k, (dt, serial, what) in enumerate((("2026-07-30", 46233, "Terra、Luna 降價"), ("2026-08-22", 46256, "Sol 促銷"),
                                              ("2026-09-03", 46268, "GPT-6 Astra 上市（top 層價格上升）"),
                                              ("2026-09-22", 46287, "GPT-6 Sol／Luna 上市（mid、low 層降價）")), start=2):
        S(ev, f"API 價格事件 {k} 公告日（Excel 日期序列值）", vlit=serial, unit="日期序列值", tag="Verified", source=src_txt, date=dt,
          scope=f"{dt}（{what}）；公告日＝生效日（v0.5 事件清單未區分）；序列值 {serial}＝{dt}",
          use="P2 價格事件時點基準（E8c）", note="E8c：事件時點的基準值（Verified）；區間由 Inputs 偏移天數（±1 季，Assumed）給出")
    for k, dt in ((2, "07-30"), (3, "08-22"), (4, "09-03"), (5, "09-22")):
        I("r5/E8c/價格事件時點", f"價格事件 {k}（{dt}）生效日偏移", unit="天", vlit=0, tag="Assumed", decision="E8c",
          status="P2 新增（E8c）", note="基準 0＝SRC_OAI 公告日（Verified）；區間 ±91 天（±1 季，Assumed；S2 工作單預設）。偏移後日期限於 FY2026 內", meta=False)
        R.inp_rows[-1]["lo"], R.inp_rows[-1]["hi"] = -91, 91
    for name, unit, val, why in (
        ("FY2026 起日（定義常數；Excel 日期序列值）", "日期序列值", 46023, "46023＝2026-01-01；價格事件天數加權的年度起點"),
        ("每年天數（定義常數）", "天", 365, "每日→年換算與 FY2026 天數加權的分母（2026 為 365 天）"),
        ("每年月數（定義常數）", "月", 12, "月 ARPU → 年營收"),
        ("單位換算：$M → $B 的除數（定義常數）", "$M／$B", 1000, "人數（M）× $ → $M；T token × $/M token → $M"),
        ("單位換算：T token → M token 的乘數（定義常數）", "M／T", 1000000, "與 Tokenomics IF_AllocDemand（M tok/年）並列"),
        ("FY2026 獨立推估：下半年月數（定義常數）", "月", 6, "v0.5 fy2026Independent：Q1＋Q2＋下半年以 2026-08 ARR 外推"),
    ):
        I("r5/E6/P2 定義常數", name, unit=unit, vlit=val, tag="Decision", decision="E6", status="P2 新增（E6）",
          note=f"來源：定義常數；區間不適用（低＝高）；E6：公式內不得含常數，故進 Inputs。用途：{why}", meta=False)
        R.inp_rows[-1]["lo"], R.inp_rows[-1]["hi"] = val, val


def p3_rows(R: Recorder):
    """v0.6-P3（S3）新增 Inputs：V1 Blackwell 拆分（r6 P3-2）、V2 η 路徑（S3 預設）、E6 定義常數。不改 v0.5 葉節點的去處。"""
    I = R.inp
    for idx, val, why in (("2025", 0, "r6 P3-2：2025 GB200 100%"), ("2026 起", 0.5, "r6 P3-2：2026 起 GB200 50%／GB300 50%")):
        I("r6/P3-2/Blackwell 拆分", "GB300 占 Blackwell 比例", index=idx, unit="比例", vlit=val, tag="Assumed", decision="V1",
          status="P3 新增（V1、r6 P3-2）", note=f"{why}；區間 GB300 占 Blackwell 0–80%（Assumed，r6 P3-2）。GB200＝Blackwell×(1−本列)", meta=False)
        R.inp_rows[-1]["lo"], R.inp_rows[-1]["hi"] = 0, 0.8
    for name, unit, val, lo, hi, why in (
        ("η 逐年路徑開關（0＝沿用 2025 η；1＝線性回升至目標）", "開關", 0, 0, 1,
         "S3 工作單預設：基準沿用 2025 的 η（0）；情境 1＝自 2025 線性回升至目標年的目標值（r6 P3-3 待判斷事項，Andy 2026-10-07 授權依預設）"),
        ("η 回升情境：目標值", "倍", 1, 1, 1, "S3 工作單預設：回升到 1（模型產能完全兌現）"),
        ("η 回升情境：目標年", "年", 2030, 2030, 2030, "S3 工作單預設：2030 年"),
    ):
        I("r6/P3-3/η 路徑", name, unit=unit, vlit=val, tag="Decision", decision="V2", status="P3 新增（V2）", note=why, meta=False)
        R.inp_rows[-1]["lo"], R.inp_rows[-1]["hi"] = lo, hi
    I("r5/E6/P3 定義常數", "線性爬升付款：等差級數和 n(n+1)/2 的除數（定義常數）", unit="—", vlit=2, tag="Decision", decision="E6",
      status="P3 新增（E6）", note="來源：定義常數；區間不適用（低＝高）；E6：公式內不得含常數。用途：N4(b) 期間內線性爬升的付款權重（年序 ÷ Σ年序）", meta=False)
    R.inp_rows[-1]["lo"], R.inp_rows[-1]["hi"] = 2, 2


LEAK_SRC = ("Where's Your Ed At 2026-06-15『Exclusive: OpenAI Losses Increased Nearly 8X in 2025』（外流 2025 經查核財報；FT 獨立核實）"
            "https://www.wheresyoured.at/exclusive-openai-financials/；Fortune 2026-06-16 https://fortune.com/2026/06/16/openai-financials-leaked-losses-revenue-profit/")
WSJ_SBC_SRC = ("WSJ 2025-12-30『OpenAI Is Paying Employees More Than Any Major Tech Startup in History』（OpenAI 向投資人提供的財務預測；Equilar 分析），"
               "經 The Decoder https://the-decoder.com/openais-stock-compensation-averages-1-5-million-per-employee-dwarfing-every-tech-startup-in-history/ 、"
               "Fortune 2026-02-18 https://fortune.com/2026/02/18/openai-chatgpt-creator-record-million-dollar-equity-compensation-ai-tech-talent-war-career-retention-sam-altman-millionaire-staff/ 轉述")
FT_HC_SRC = ("FT 2026-03-21（兩位知情人士；招募計畫），經 Engadget https://www.engadget.com/ai/openai-reportedly-plans-to-double-its-workforce-to-8000-employees-161028377.html 轉述")


def p4_rows(R: Recorder):
    """v0.6-P4（S4）新增列：r6 P4-2 人數、每人成本、股權報酬的公開來源（SRC_OAI；擷取日 2026-10-07）、
    2025 銷售／管理費用（v0.5 opexExComputePctRevenue 2025 的拆分來源）、P4 假設（Inputs）與 E6 定義常數。不改 v0.5 葉節點的去處。"""
    S, I = R.src, R.inp
    ex = "（擷取 2026-10-07）"
    for metric, val, scope, use in (
        ("銷售費用（sales and marketing）：FY2025", 5.73, "FY2025 經查核財報（GAAP 費用列；是否含股權報酬原文未交代，本模型假設含）", "P4 V4：2025 銷售費用（v0.5 opexExComputePctRevenue 2025 的拆分來源）"),
        ("管理費用（general and administrative）：FY2025", 1.57, "FY2025 經查核財報（GAAP 費用列；是否含股權報酬原文未交代，本模型假設含）", "P4 V4：2025 管理費用（同上）"),
        ("營業損失（loss from operations）：FY2025", 20.92, "FY2025 經查核財報；＝營收 −（營業成本＋研發＋銷售＋管理）", "P4 Checks：費用各列與營業損失對帳"),
    ):
        S("r6/P4/外流 2025 財報（S4 搜尋）", metric, vlit=val, unit="$B", scope=scope, date="2026-06-15", tag="Interested-party",
          source=LEAK_SRC, use=use, note="S4 公開來源搜尋" + ex + "；與 SRC_OAI_002、003、087、098 同一份外流財報")
    for metric, val, unit, scope, use in (
        ("員工人數：2025（約）", 4000, "人", "2025 年員工約 4,000 人（每人股權報酬的分母；視為年均）", "P4 V4：人數 2025 值"),
        ("平均每人股權報酬：2025", 1.5, "$M/人/年", "2025 年平均每人股權報酬約 $1.5M（投資人財務預測）", "P4 V4：每人股權報酬 2025 值"),
        ("股權報酬占營收比：2025（投資人預測）", 0.462, "比例", "2025 年股權報酬約為營收的 46.2%（投資人財務預測）", "P4 Checks 對照"),
        ("股權報酬年增額：至 2030（投資人預測）", 3, "$B/年", "股權報酬預期至 2030 年每年增加約 $3B（投資人財務預測）", "P4 對照列（不回饋基準）"),
    ):
        S("r6/P4/WSJ 股權報酬（S4 搜尋）", metric, vlit=val, unit=unit, scope=scope, date="2025-12-30", tag="Interested-party",
          source=WSJ_SBC_SRC, use=use, note="S4 公開來源搜尋" + ex + "；原文 WSJ 付費牆，數值經兩家轉述一致")
    for metric, val, scope, use in (
        ("員工人數：2026-03（約）", 4500, "2026-03 員工約 4,500 人", "P4 Checks 對照（人數成長路徑）"),
        ("員工人數目標：2026 年底（約）", 8000, "2026 年底目標約 8,000 人（公司招募計畫；只作對照，不回饋基準）", "P4 Checks 對照（人數成長路徑）"),
    ):
        S("r6/P4/FT 人數（S4 搜尋）", metric, vlit=val, unit="人", scope=scope, date="2026-03-21", tag="Interested-party",
          source=FT_HC_SRC, use=use, note="S4 公開來源搜尋" + ex + "；原文 FT 付費牆")
    lag = I("r6/P4-1/自建 GW", "自建 GW 投產落後年數", unit="年", vlit=1, tag="Assumed", decision="v0.5 算力MW 第 11 列",
            status="P4 新增（S4 預設：自建 GW 計入供給）", note="v0.5：自建 GW＝累計自有資本支出（至前一年）÷ 每 GW 資本支出；區間 1–2 年（Assumed）", meta=False)
    R.inp_rows[-1]["lo"], R.inp_rows[-1]["hi"] = 1, 2
    for y, v, lo, hi in ((2026, 0.55, 0.3, 0.8), (2027, 0.30, 0.1, 0.5), (2028, 0.20, 0.05, 0.35), (2029, 0.15, 0, 0.3), (2030, 0.10, 0, 0.25)):
        I("r6/P4-2/人數成長", "員工人數年增率（年均人數）", index=str(y), unit="比例", vlit=v, tag="Assumed", decision="V4",
          status="P4 新增（V4；S4 預設：成長路徑用 Inputs）",
          note=("2026：FT 2026-03 約 4,500、年底目標約 8,000（SRC_OAI）隱含年均約 +56%，取 0.55；" if y == 2026 else "")
               + "Assumed；不以公司目標反推（共同規則第 4 節）", meta=False)
        R.inp_rows[-1]["lo"], R.inp_rows[-1]["hi"] = lo, hi
    I("r6/P4-2/每人成本", "每人年成本（不含股權報酬）年變動率", index="2026 起", unit="比例", vlit=0.03, tag="Assumed", decision="V4",
      status="P4 新增（V4）", note="2025 值由財報推得（Derived：非算力營運費用不含股權報酬 ÷ 人數）；之後每年變動率 0.03（0–0.08；薪資與非人事費用成長，Assumed）", meta=False)
    R.inp_rows[-1]["lo"], R.inp_rows[-1]["hi"] = 0, 0.08
    I("r6/P4-2/股權報酬", "每人股權報酬年變動率", index="2026 起", unit="比例", vlit=0, tag="Assumed", decision="V4",
      status="P4 新增（V4）", note="2025 值 $1.5M（SRC_OAI，WSJ）；之後每人不變（0；−0.2–0.2，Assumed）；WSJ『每年約 +$3B』只作對照", meta=False)
    R.inp_rows[-1]["lo"], R.inp_rows[-1]["hi"] = -0.2, 0.2
    I("r5/E6/P4 定義常數", "單位換算：$ → $B 的除數（定義常數）", unit="$／$B", vlit=1000000000, tag="Decision", decision="E6",
      status="P4 新增（E6）", note="來源：定義常數；區間不適用（低＝高）；E6：公式內不得含常數。用途：TK $/M token × M tok/GW/年 → $B/GW/年", meta=False)
    R.inp_rows[-1]["lo"], R.inp_rows[-1]["hi"] = 1000000000, 1000000000
    return lag


def p5_rows(R: Recorder):
    """v0.6-P5（S5）新增 Inputs：融資輪到位年、Amazon 條件式開關與到位年（S5 預設：不計入基準，列情境）。不改 v0.5 葉節點的去處。"""
    I = R.inp
    I("r6/P5-1/融資", "2026-03 融資輪無條件部分的現金到位年", unit="年", vlit=2026, tag="Decision", decision="S7",
      status="P5 新增（S7）", note="v0.5 資金第 10 列：無條件部分全額計入 2026 股權流入；SoftBank 三期的實際分期未揭露（區間上限 2027）", meta=False)
    R.inp_rows[-1]["lo"], R.inp_rows[-1]["hi"] = 2026, 2027
    I("r6/P5-1/融資", "計入 Amazon 條件式 $35B（1＝是；0＝否）", unit="開關", vlit=0, tag="Decision", decision="S7",
      status="P5 新增（S5 預設：不計入基準，列情境）", note="v0.5 輸入第 7 列（base 不計入）；Funding 第四節另列計入情境，不受本開關影響", meta=False)
    R.inp_rows[-1]["lo"], R.inp_rows[-1]["hi"] = 0, 1
    I("r6/P5-1/融資", "Amazon 條件式 $35B 到位年", unit="年", vlit=2026, tag="Assumed", decision="S7",
      status="P5 新增（S5）", note="條件為 IPO 或 AGI 里程碑（SRC_OAI_082）；v0.5 計入時與無條件部分同在 2026；區間 2026–2027（OpenAI 2026-06 已機密遞交 S-1 草稿，上市時點未定）", meta=False)
    R.inp_rows[-1]["lo"], R.inp_rows[-1]["hi"] = 2026, 2027


def coverage_report(R: Recorder):
    missing = [p for p in R.leaves if p not in R.dest]
    return missing


# ─────────────────────────────────────────────────────────────
#  穩定 ID：SRC_OAI_nnn、INP_nnn 一經發出即固定（registry）；新列取下一個號碼；被移除的號碼退役、不重用。
#  理由：Excel 優先的保留機制與 chat 端審查都以 ID 為鍵，插入新列不得讓既有 ID 位移。
# ─────────────────────────────────────────────────────────────
import re as _re

REGISTRY_PATH = Path(__file__).resolve().parent / "id_registry.json"


def _src_key(r):
    return r["metric"]


def _inp_key(r):
    return "|".join([str(r["name"]), str(r["index"]), r["v05"].replace(SEP, ".")])


def finalize_ids(R: Recorder, write_registry=True):
    reg = json.loads(REGISTRY_PATH.read_text(encoding="utf-8")) if REGISTRY_PATH.exists() else {"SRC": {}, "INP": {}}
    reg.setdefault("retired", {"SRC": [], "INP": []})
    out = {}
    for kind, rows, keyf, prefix in (("SRC", R.src_rows, _src_key, "SRC_OAI_"), ("INP", R.inp_rows, _inp_key, "INP_")):
        known = reg[kind]
        used = [int(v.split("_")[-1]) for v in known.values()] + [int(v.split("_")[-1]) for v in reg["retired"][kind]]
        nxt = max(used + [0]) + 1
        mapping = {}
        keys = [keyf(r) for r in rows]
        assert len(set(keys)) == len(keys), f"{kind} 鍵重複：{[k for k in keys if keys.count(k) > 1][:3]}"
        for r, k in zip(rows, keys):
            if k in known:
                new = known[k]
            else:
                new = f"{prefix}{nxt:03d}"
                nxt += 1
                known[k] = new
            mapping[r["id"]] = new
        # 退役：registry 內有、但本次沒有的鍵
        for k in [k for k in known if k not in set(keys)]:
            reg["retired"][kind].append(known.pop(k))
        out[kind] = mapping
        for r in rows:
            r["id"] = mapping[r["id"]]
        rows.sort(key=lambda r: r["id"])
    allmap = {**out["SRC"], **out["INP"]}
    pat = _re.compile(r"(SRC_OAI_\d{3}|INP_\d{3})")
    R.dest = {p: (k, pat.sub(lambda m: allmap.get(m.group(1), m.group(1)), ref), why) for p, (k, ref, why) in R.dest.items()}
    if write_registry:
        REGISTRY_PATH.write_text(json.dumps(reg, ensure_ascii=False, indent=0), encoding="utf-8")
    return reg


if __name__ == "__main__":
    R = build_map()
    miss = coverage_report(R)
    from collections import Counter
    print("leaves", len(R.leaves), "assigned", len(R.dest), "missing", len(miss))
    print(Counter(k for k, _, _ in R.dest.values()))
    print("SRC rows", len(R.src_rows), "INP rows", len(R.inp_rows), "pending", len(R.pending))
    for p in miss[:60]:
        print("MISSING", p, repr(R.leaves[p])[:60])
