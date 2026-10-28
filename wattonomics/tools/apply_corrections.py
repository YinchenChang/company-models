# 把複核段的欄位更正表（例：evidence_raw/H_corrections.csv）套用到 evidence_raw/*.csv，並寫出套用紀錄。
# 用法：python3 wattonomics/tools/apply_corrections.py wattonomics/evidence_raw/H_corrections.csv
# 規則：
#   - old 與現值完全相同 → 整格換成 new。
#   - old 含「…」（節錄）→ 依序找出各段，把第一段起點到最後一段終點換成 new（去掉頭尾「…」）。
#   - field 為 (none)／(note) → 只記「維持」，不改。
#   - 複合欄位（low/high、high/tag）→ 依下方 COMPOSITE 逐筆處理（寫明怎麼拆）。
#   - 找不到 old → 不改、記「未套用」，程式以非 0 結束，需人工處理。
# 重跑安全：現值已等於 new（或已含 new）者記「已套用」。
import csv, os, sys

WT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW = os.path.join(WT, "evidence_raw")

# 複合欄位：(檔案, 列) → {欄位: 新值}；寫入前核對舊值
COMPOSITE = {
    ("P_facility_power.csv", 5): dict(check={"low": "13", "high": "15"},
                                      set={"tag": "[Analogy] Epoch AI（經 CorAdvisors 轉述，Epoch 原頁只見總額 38B）；low 13／high 15 [Assumed]（原文無）"}),
    ("P_facility_power.csv", 27): dict(check={"low": "0.8", "high": "1.2"},
                                       set={"low": "0.5",
                                            "tag": "[Analogy] NREL（基準 1.0）；區間 0.5／1.2 [Assumed]（low 參考 Groww 單機 2–3MW $0.5–2M＝0.17–1.0 $M/MW 設備價＋安裝）"}),
    ("P_facility_power.csv", 29): dict(check={"high": "0.5"},
                                       set={"tag": "記者轉述 Microsoft（Data Centre Magazine）；high 0.5 [Assumed]（非 GPU 負載可能仍設發電機，原文未說）"}),
}


def load(fn):
    with open(os.path.join(RAW, fn), encoding="utf-8-sig", newline="") as f:
        return list(csv.reader(f))


def save(fn, rows):
    bom = open(os.path.join(RAW, fn), "rb").read(3) == b"\xef\xbb\xbf"   # 保留原檔 BOM
    with open(os.path.join(RAW, fn), "w", encoding="utf-8-sig" if bom else "utf-8", newline="") as f:
        csv.writer(f).writerows(rows)


def frag_replace(cur, old, new):
    segs = [s for s in old.split("…") if s.strip()]
    pos, start, end = 0, None, None
    for s in segs:
        i = cur.find(s, pos)
        if i < 0:
            return None
        if start is None:
            start = i
        end = i + len(s); pos = end
    return cur[:start] + new.strip("…") + cur[end:]


def main(corr_path):
    corr = list(csv.DictReader(open(corr_path, encoding="utf-8-sig")))
    files = {}
    log = []
    done_comp = set()
    for c in corr:
        fn, n, fld = c["file"], int(c["row"]), c["field"]
        rows = files.setdefault(fn, load(fn)); h = rows[0]
        if fld in ("(none)", "(note)"):
            log.append((fn, n, fld, "維持", c["reason"])); continue
        if (fn, n) in COMPOSITE:
            if (fn, n) in done_comp:
                log.append((fn, n, fld, "併入同列複合更正", c["reason"])); continue
            spec = COMPOSITE[(fn, n)]; row = rows[n - 1]
            st = "套用"
            for k, v in spec["check"].items():
                if row[h.index(k)] != v and not (k in spec["set"] and row[h.index(k)] == spec["set"][k]):
                    st = f"未套用（{k} 現值 {row[h.index(k)]}）"
            if st == "套用":
                for k, v in spec["set"].items():
                    row[h.index(k)] = v
            done_comp.add((fn, n)); log.append((fn, n, fld, st, c["reason"])); continue
        j = h.index(fld); cur = rows[n - 1][j]; old, new = c["old"], c["new"]
        if cur.strip() == new.strip() or (("…" in new or "…" in old) and new.strip("…") in cur):
            log.append((fn, n, fld, "已套用", c["reason"])); continue
        if cur.strip() == old.strip():
            rows[n - 1][j] = new; log.append((fn, n, fld, "套用", c["reason"])); continue
        if "…" in old:
            r = frag_replace(cur, old, new)
            if r is not None:
                rows[n - 1][j] = r; log.append((fn, n, fld, "套用（節錄比對）", c["reason"])); continue
        log.append((fn, n, fld, "未套用（舊值不符）", c["reason"]))
    for fn, rows in files.items():
        save(fn, rows)
    out = os.path.splitext(corr_path)[0] + "_applied.csv"
    with open(out, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f); w.writerow(["file", "row", "field", "status", "reason"]); w.writerows(log)
    from collections import Counter
    cnt = Counter(x[3] for x in log)
    print(out, dict(cnt))
    bad = [x for x in log if x[3].startswith("未套用")]
    for x in bad:
        print("未套用：", x)
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main(sys.argv[1])
