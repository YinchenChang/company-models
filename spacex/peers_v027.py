# SpaceX v0.2.7 同業資料（驗收 B3：統一估值日與 NTM 口徑）：MarketScreener 估值頁、財務頁原值（2026-11-06 收盤，2026-11-09 以內建瀏覽器讀取）寫在本檔，產生 data/peers_connectivity_20261106.json、data/peers_si_20261106.json。
# 用法：cd spacex && python3 peers_v027.py
import json, statistics
def num(s): return float(s.replace(',','').replace('x','').replace('%',''))
RAW={
"CMCSA":dict(name="Comcast",id="COMCAST-CORPORATION-4864",fy=list(range(2021,2029)),ev="321,263 241,252 267,375 235,381 198,372 167,382 163,400 162,339",ebitda="34,708 36,459 37,633 38,069 37,384 33,122 32,776 33,101",sales="116,385 121,427 121,572 123,731 123,707 119,783 116,430 119,247",nd="91,309 90,062 90,875 91,771 89,456 80,773 76,791 75,730",fcf="17,089 12,646 12,962 12,543 19,235 13,595 13,039 13,362",px=24.39),
"TMUS":dict(name="T-Mobile US",id="T-MOBILE-US-INC-24717887",fy=list(range(2021,2029)),ev="215,002 244,166 257,797 331,336 310,056 252,492 256,064 262,278",ebitda="26,924 27,821 29,428 31,864 33,937 37,418 40,381 42,954",sales="80,118 79,571 78,558 81,400 88,309 94,805 98,946 102,905",nd="70,137 69,984 72,379 75,182 82,954 80,880 84,451 90,665",fcf="5,646 7,656 13,586 17,032 17,995 18,615 20,281 21,766",px=162.01),
"T":dict(name="AT&T",id="AT-T-INC-14324",fy=list(range(2021,2029)),ev="331,854 263,382 250,586 283,465 293,468 293,112 293,837 293,958",ebitda="51,469 41,465 43,400 44,760 46,361 48,245 49,672 51,933",sales="168,864 120,741 122,428 122,336 125,648 129,946 129,955 132,409",nd="156,185 132,189 130,609 120,084 117,366 147,402 148,127 148,248",fcf="25,430 16,186 20,461 17,644 16,586 17,945 19,216 20,906",px=21.63),
"IRDM":dict(name="Iridium Communications",id="IRIDIUM-COMMUNICATIONS-IN-5602836",fy=list(range(2021,2029)),ev="6,736 7,794 6,526 5,018 3,499 6,370 6,218 6,161",ebitda="378.2 424 463.1 470.6 495.3 508.7 532 524.3",sales="614.5 721 790.7 830.7 871.7 928.5 1,002 1,034",nd="1,277 1,336 1,428 1,714 1,678 1,695 1,543 1,486",fcf="264.4 - 241.4 304 307.5 295.8 378.8 333.5",px=44.04),
"SATS":dict(name="EchoStar",id="ECHOSTAR-CORPORATION-782521",fy=list(range(2021,2029)),ev="2,227 1,205 24,817 27,608 54,288 23,036 26,477 27,674",ebitda="718.9 656 1,320 1,626 1,494 2,311 2,090 1,859",sales="1,986 1,998 17,016 15,826 15,005 12,723 13,203 12,527",nd="-50.4 -181.7 20,320 21,056 22,996 -4,497 -1,055 141.8",fcf="193.8 203.7 -668.3 -292.2 -1,065 2,161 2,562 1,937",px=94.62,note="EchoStar 處分頻譜（含售予 SpaceX）後轉為淨現金，倍數受交易影響"),
}
W26=55/365
def rec(k,v,cur,nxt,w_cur,ntm_note):
    f=lambda s:[None if x=='-' else num(x) for x in v[s].split()]
    ev,eb,sa,nd,fc=f('ev'),f('ebitda'),f('sales'),f('nd'),f('fcf')
    i=v['fy'].index(cur); j=v['fy'].index(nxt)
    ntm=w_cur*eb[i]+(1-w_cur)*eb[j]
    rd=lambda a,b,n: round(a/b,n) if (a is not None and b) else None
    return {"ticker":k,"name":v['name'],"url":f"https://www.marketscreener.com/quote/stock/{v['id']}/valuation/",
      "refPrice":v['px'],"evCurrent":ev[i],"ebitda":{str(y):eb[n] for n,y in enumerate(v['fy'])},"sales":{str(y):sa[n] for n,y in enumerate(v['fy'])},
      "netDebt":{str(y):nd[n] for n,y in enumerate(v['fy'])},"fcf":{str(y):fc[n] for n,y in enumerate(v['fy'])},
      "ntmEbitda":round(ntm,1),"ntmEvEbitda":round(ev[i]/ntm,3),
      "salesGrowthNext":round(sa[j]/sa[i]-1,4),"ebitdaMarginNext":round(eb[j]/sa[j],4),"netDebtToEbitdaNext":rd(nd[j],eb[j],3),
      "fcfToEbitdaNext":rd(fc[j],eb[j],3),"ntmNote":ntm_note,"tag":"[Derived]","note":v.get('note','')}
common={"source":"MarketScreener 估值頁（valuation）與財務頁（finances），以 Andy 電腦上的內建瀏覽器讀取","retrieved":"2026-11-09","asOf":"2026-11-06（收盤；頁面 Reference price）",
 "method":"NTM EV/EBITDA＝當期 EV（頁面當年度欄：現價市值＋該年預估淨負債）÷ NTM EBITDA；NTM＝評估日 2026-11-06 起 12 個月＝當年度剩餘 55/365 × 當年度＋其餘 × 下一年度（12 月年制）；Oracle 財年 5 月：FY2027 剩 205/365",
 "dateGap":"模型現價日 2026-11-05，同業取 2026-11-06 收盤（差一個交易日；同一來源全組 11-05 資料未取得）","capex":"資本支出率：MarketScreener 財務頁未列 Capex，未取得；以 FCF ÷ EBITDA 作資本強度代理（Derived）"}
pc={"_readme":"SpaceX Connectivity 分部同業（v0.2.7；取代 peers_connectivity_20261030.json 的 FY2026 口徑，統一估值日與 NTM）",**common,"group":"connectivity","label":"衛星與電信",
    "peers":[rec(k,RAW[k],2026,2027,W26,"55/365 × FY2026＋310/365 × FY2027") for k in ("IRDM","TMUS","T","CMCSA","SATS")]}
pc["median"]=round(statistics.median(p["ntmEvEbitda"] for p in pc["peers"]),3)
json.dump(pc,open('data/peers_connectivity_20261106.json','w',encoding='utf-8'),ensure_ascii=False,indent=1)
SI={"CRWV":dict(name="CoreWeave",id="COREWEAVE-INC-185380761",fy=[2025,2026,2027,2028],ev="53,897 70,817 99,918 120,392",ebitda="3,093 7,554 16,711 29,237",sales="5,131 12,873 26,445 43,662",nd="- - - -",fcf="- - - -",px=46.40),
    "NBIS":dict(name="Nebius Group",id="NEBIUS-GROUP-8037501",fy=list(range(2021,2029)),ev="20,277 6,846 6,846 6,522 21,527 49,152 69,809 86,221",ebitda="424.5 859.6 - -266.4 -64.9 1,380 6,485 13,418",sales="4,704 6,992 - 117.5 529.8 3,346 12,281 23,137",nd="- - - - - - - -",fcf="- - - - - - - -",px=129.49),
    "ORCL":dict(name="Oracle",id="ORACLE-CORPORATION-13620698",fy=list(range(2022,2030)),ev="245,851 366,311 398,306 545,550 747,001 477,563 510,525 517,922",ebitda="21,565 23,428 26,184 28,900 36,549 52,090 74,148 105,490",sales="42,440 49,954 52,961 57,399 67,357 90,447 131,819 184,052",nd="- - - - - - - -",fcf="- - - - - - - -",px=116.10)}
ps=[rec(k,SI[k],2026,2027,W26,"55/365 × FY2026＋310/365 × FY2027") for k in ("CRWV","NBIS")]
r=rec("ORCL",SI["ORCL"],2027,2028,205/365,"205/365 × FY2027（至 2027-05）＋160/365 × FY2028"); r["note"]="財年 5 月；FY2027＝2026-06～2027-05"; ps.append(r)
psi={"_readme":"SpaceX SI 分部同業（v0.2.7；只用於讀法 1 自用 AI 營運成本扣除倍數；統一估值日與 NTM，取代 peers_si_20261030.json 混用 FY2027 與 NTM）",**common,"group":"si_ai_infra","label":"AI 基礎設施／雲","peers":ps}
psi["median"]=round(statistics.median(p["ntmEvEbitda"] for p in ps),3)
json.dump(psi,open('data/peers_si_20261106.json','w',encoding='utf-8'),ensure_ascii=False,indent=1)
for p in pc["peers"]+ps: print(p["ticker"],p["ntmEvEbitda"],p["salesGrowthNext"],p["ebitdaMarginNext"],p["netDebtToEbitdaNext"],p["fcfToEbitdaNext"])
print(pc["median"],psi["median"])
