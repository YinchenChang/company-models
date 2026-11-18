import sys,subprocess,os,re
from filings import CIK,lst
from bs4 import BeautifulSoup
UA="Mozilla/5.0 research"
def get(t,acc,doc):
    d=f"/home/claude/out/channel/raw/docs/{t}"
    os.makedirs(d,exist_ok=True)
    p=f"{d}/{acc}.txt"
    if os.path.exists(p) and os.path.getsize(p)>1000: return p
    url=f"https://www.sec.gov/Archives/edgar/data/{CIK[t]}/{acc.replace('-','')}/{doc}"
    subprocess.run(["curl","-sS","-m","120","-A",UA,"-o",p+".html",url],check=True)
    soup=BeautifulSoup(open(p+".html",errors="ignore").read(),"lxml")
    for tr in soup.find_all("tr"):
        tr.append("\n")
    txt=soup.get_text(" ")
    txt=re.sub(r"[ \t\xa0]+"," ",txt)
    txt=re.sub(r"\n\s*\n+","\n",txt)
    open(p,"w").write(txt)
    return p
if __name__=="__main__":
    t=sys.argv[1]; since=sys.argv[2] if len(sys.argv)>2 else "2024-01-01"
    for f in lst(t,since=since):
        p=get(t,f[3],f[4]); print(t,f[0],f[1],f[2],p,os.path.getsize(p))
