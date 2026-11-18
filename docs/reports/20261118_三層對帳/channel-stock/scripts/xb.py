import json,datetime as dt
def D(s): return dt.date.fromisoformat(s)
def load(t): return json.load(open(f"/home/claude/out/channel/raw/{t}.json"))
def series(d,concept,unit="USD",tax="us-gaap"):
    f=d["facts"][tax].get(concept)
    if not f: return None
    return f["units"].get(unit,[])
def latest_by(entries,key):
    out={}
    for e in entries:
        k=key(e)
        if k not in out or e["filed"]>out[k]["filed"]: out[k]=e
    return out
def instant(d,concept,**kw):
    en=series(d,concept,**kw)
    if not en: return {}
    en=[e for e in en if "start" not in e]
    m=latest_by(en,lambda e:e["end"])
    return {k:(v["val"],v["accn"],v["form"],v["filed"]) for k,v in m.items()}
def discrete(d,concept,**kw):
    en=series(d,concept,**kw)
    if not en: return {}
    m=latest_by([e for e in en if "start" in e],lambda e:(e["start"],e["end"]))
    out={}
    bystart={}
    for (s,e),v in m.items():
        bystart.setdefault(s,[]).append((e,v))
    for (s,e),v in m.items():
        days=(D(e)-D(s)).days
        if 80<=days<=100:
            out[e]=(v["val"],v["accn"],v["form"],v["filed"],"Q")
        elif days>100:
            # find shorter with same start ending ~3 months before
            prevs=[(pe,pv) for pe,pv in bystart[s] if 80<(D(e)-D(pe)).days<100]
            if prevs and e not in out:
                pe,pv=prevs[0]
                out[e]=(v["val"]-pv["val"],v["accn"],v["form"],v["filed"],"YTDdiff")
    return out
