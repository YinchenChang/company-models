import json,sys,subprocess
CIK={"MSFT":789019,"AMZN":1018724,"GOOGL":1652044,"META":1326801,"ORCL":1341439,"CRWV":1769628,"NBIS":1513845,"DELL":1571996,"SMCI":1375365,"HPE":1645590}
def lst(t,forms=("10-Q","10-K"),since="2024-01-01"):
    d=json.load(open(f"/home/claude/out/channel/raw/sub_{t}.json"))
    r=d["filings"]["recent"]
    out=[]
    for i,f in enumerate(r["form"]):
        if f in forms and r["filingDate"][i]>=since:
            out.append((r["filingDate"][i],f,r["reportDate"][i],r["accessionNumber"][i],r["primaryDocument"][i]))
    return sorted(out)
if __name__=="__main__":
    for t in sys.argv[1:]:
        for x in lst(t): print(t,*x)
