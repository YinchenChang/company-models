# 把 wattonomics/evidence_raw/*.csv 整理成 builder/seed/evidence_v0.1.json（DB_Evidence 頁的來源），附 C 段獨立核對判定。
# 用法：python3 wattonomics/tools/import_evidence.py
# 列識別：A#n＝A_us_cloud.csv 第 n 行；B#n＝B_labs_china_totals.csv 第 n 行；D#n＝B_delay.csv 第 n 行（第 1 行為表頭）。
# 判定依 evidence_raw/C_review.md（以它為準）；C 段未列者為「未抽查」。
import csv, json, os

GL = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW = os.path.join(GL, "evidence_raw")
OUT = os.path.join(GL, "builder", "seed", "evidence_v0.1.json")

ERR = {"A33"}
MINOR = {"A19", "A45", "A21", "B3", "B47", "A116", "A16", "A58", "B16", "B36", "D12", "D7"}
UNVER = {"A2", "A27", "A30", "B30", "A36", "A83", "B34", "D18"}
TRUE = set("""A10 A11 A12 A23 A24 A38 A39 A51 A52 A53 A67 A68 A77 A78 A79 A25 A3 A4 A5 A6 A7 A8 A15 A17 A18 A29 B29 A35 A37 A41 A42
A46 B22 A50 A54 A55 A57 A59 A63 A64 A65 A69 B7 A72 A73 A74 A75 A76 A81 A82 A84 A85 A94 A87 A88 A95 A104 A105 A114 A117
B2 B4 B6 B23 B24 B33 B44 B45 B46 B48 D25 B49 B51 B58 B63 B64 B65 B67 B70 B72 D15 D2 D3 D5 D6 D8 D11 D13 D14 D16 D17 D19 D20 D21 D22 D23 D24""".split())
NOTE = {
    "A33": "方向寫反：營運者＝SpaceX／xAI、使用者＝Google（B#49）；不計入 Google 營運者 GW",
    "A19": "H 段更正（取代 C 段『全案上限 1 GW』）：$30B Azure 確定採購（GW 未揭露）＋額外至多 1 GW（上限）", "A45": "標記改 Interested-party（GW 數為 Anthropic 自述／The Information）",
    "A21": "差額原因另含耐用年限 15→25 年", "B3": "近 90 天簽 3 GW（含 AWS 2 GW），非 >3 GW",
    "B47": "946 MW-IT 屬實；1,531 MW（2027）無出處，刪除", "A116": "數值屬實；『不含中國』無出處，中國買家晶片可能在 Other 內",
    "A16": "as_of 2026-05-15", "A58": "Entergy 7 座燃氣電廠；租金起付日不在原文；2030 前 2 GW 屬實",
    "B16": "全園 8 IT-GW；NVIDIA 首期 4.25 IT-GW、3.75 為選擇權；2028 起分期", "B36": "首期 2027H2、2028 初全部交付",
    "D12": "低值依據不成立；中值 30 月屬實", "D7": "措辭：『許多案子』渦輪不是卡關項",
    "A30": "原頁被擋；標記統一為 Interested-party（A／B 兩段原不一致）", "B30": "同 A#30（標記已統一）",
    "A79": "Epoch 交叉屬實；10-K 原頁被擋", "A83": "10-K 被擋", "A72": "屬實（設施）；首棟已延到 2027-08（B#10）",
}


PRE = {"A_us_cloud.csv": "A", "B_labs_china_totals.csv": "B", "B_delay.csv": "D"}


def h_notes():
    """H 段（10-28 API 旗標人工複核）更正理由：{A18: [...]}；併入 cnote"""
    out = {}
    p = os.path.join(RAW, "H_corrections.csv")
    if os.path.exists(p):
        for c in csv.DictReader(open(p, encoding="utf-8-sig")):
            pre = PRE.get(c["file"])
            if pre:
                out.setdefault(f"{pre}{c['row']}", []).append(c["reason"] if c["field"] in ("(none)", "(note)") else f"{c['field']}：{c['reason']}")
    return out


def verdict(i):
    if i in ERR: return "錯誤"
    if i in UNVER: return "無法核對"
    if i in MINOR: return "小誤（已更正）"
    if i in TRUE: return "屬實"
    return "未抽查"


def main():
    rows = []
    HN = h_notes()
    for pre, fn in (("A", "A_us_cloud.csv"), ("B", "B_labs_china_totals.csv"), ("D", "B_delay.csv")):
        with open(os.path.join(RAW, fn), encoding="utf-8-sig") as f:
            rd = list(csv.reader(f))
        h = rd[0]
        for n, row in enumerate(rd[1:], 2):
            d = dict(zip(h, row)); i = f"{pre}{n}"
            if pre == "D":
                claim = f"{d.get('bottleneck')}｜{d.get('region')}｜{d.get('metric')}"
                val = "／".join(x for x in (d.get("low_months"), d.get("mid_months"), d.get("high_months")) if x) + (" 月" if d.get("mid_months") else "")
                player = d.get("region", ""); basis = ""
            else:
                claim = f"{d.get('player')}｜{d.get('site_or_scope') or d.get('site_or_contract')}｜{d.get('metric')}"
                val = f"{d.get('value')} {d.get('unit')}".strip(); player = d.get("player", ""); basis = d.get("basis", "")
            tag = d.get("tag", "")
            if i in ("A30", "B30"):   # 同一份 Broadcom 8-K：統一為 Interested-party（原頁被擋未核；GW 數為 Broadcom／Anthropic 揭露的合約上限）
                tag = "[Interested-party]（統一；原 A 段 Interested-party、B 段 Verified）"
            rows.append(dict(id=f"{pre}#{n}", key=i, player=player, claim=claim, value=val, basis=basis, as_of=d.get("as_of", ""),
                             tag=tag, url=d.get("url", ""), verdict=verdict(i),
                             cnote="；".join(x for x in (NOTE.get(i, ""), ("H 段（10-28）：" + "；".join(HN[i])) if i in HN else "") if x)))
    # S 段（園區候選，v0.1.2）：判定＝API 回查（Sonnet）；合併入 T2 的列另經 I 段人工開原文核對（I_review_sites.md）
    sp = os.path.join(RAW, "S_sites.csv")
    if os.path.exists(sp):
        api = {}
        with open(os.path.join(RAW, "S_sites_audit.csv"), encoding="utf-8-sig") as f:
            for a in csv.DictReader(f):
                api[int(a["line"])] = a
        ir = open(os.path.join(RAW, "I_review_sites.md"), encoding="utf-8").read() if os.path.exists(os.path.join(RAW, "I_review_sites.md")) else ""
        import re
        icited = {int(x) for x in re.findall(r"S#(\d+)", ir)}
        with open(sp, encoding="utf-8-sig") as f:
            rd = list(csv.reader(f))
        h = rd[0]
        for n, row in enumerate(rd[1:], 2):
            d = dict(zip(h, row)); a = api.get(n, {})
            rows.append(dict(id=f"S#{n}", key=f"S{n}", player=d.get("player", ""), claim=f"{d.get('player')}｜{d.get('site_or_scope')}｜{d.get('metric')}",
                             value=f"{d.get('value')} {d.get('unit')}".strip(), basis=d.get("basis", ""), as_of=d.get("as_of", ""), tag=d.get("tag", ""),
                             url=d.get("url", ""), verdict="API：" + (a.get("verdict") or "未判定"),
                             cnote="；".join(x for x in ((a.get("note") or "")[:200], "I 段已開原文核對（I_review_sites.md）" if n in icited else "") if x)))
    json.dump(rows, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    from collections import Counter
    print(f"{OUT}：{len(rows)} 列；", dict(Counter(r["verdict"] for r in rows)))


if __name__ == "__main__":
    main()
