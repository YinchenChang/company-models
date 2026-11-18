# 由 XBRL companyfacts 整理美股季資料，輸出 data/us_data.json
import json,datetime as dt
from xb import *
CIK={"MSFT":789019,"AMZN":1018724,"GOOGL":1652044,"META":1326801,"ORCL":1341439,"CRWV":1769628,"NBIS":1513845,"DELL":1571996,"SMCI":1375365,"HPE":1645590}
def cq(end):
    e=D(end)-dt.timedelta(days=45)
    return f"{e.year}Q{(e.month-1)//3+1}"
QS=[f"{y}Q{q}" for y in (2024,2025,2026) for q in (1,2,3,4)][:-1]
def src(t,acc): return f"https://www.sec.gov/Archives/edgar/data/{CIK[t]}/{acc.replace('-','')}/"
def pick(d,concept,kind,scale=1e6):
    r=discrete(d,concept) if kind=="d" else instant(d,concept)
    out={}
    for end,v in r.items():
        q=cq(end)
        if q in QS: out[q]={"v":round(v[0]/scale,3),"end":end,"acc":v[1],"form":v[2],"filed":v[3],"how":(v[4] if kind=="d" else "instant")}
    return out
res={}
# ODM-US
for t,rev,cost,cl in [("DELL","Revenues","CostOfRevenue","ContractWithCustomerLiability"),
    ("SMCI","RevenueFromContractWithCustomerExcludingAssessedTax","CostOfRevenue","ContractWithCustomerLiabilityCurrent"),
    ("HPE","Revenues",None,"ContractWithCustomerLiabilityCurrent")]:
    d=load(t)
    res[t]={"rev":pick(d,rev,"d"),"cost":pick(d,cost,"d") if cost else {},"inv":pick(d,"InventoryNet","i"),"cl":pick(d,cl,"i")}
# buyers
for t,capex in [("META","PaymentsToAcquirePropertyPlantAndEquipment"),("MSFT","PaymentsToAcquirePropertyPlantAndEquipment"),
   ("AMZN","PaymentsToAcquireProductiveAssets"),("GOOGL","PaymentsToAcquirePropertyPlantAndEquipment"),
   ("ORCL","PaymentsToAcquirePropertyPlantAndEquipment"),("CRWV","PaymentsToAcquirePropertyPlantAndEquipment")]:
    d=load(t)
    res[t]={"capex":pick(d,capex,"d"),"fl_princ":pick(d,"FinanceLeasePrincipalPayments","d"),
            "fl_rou":pick(d,"RightOfUseAssetObtainedInExchangeForFinanceLeaseLiability","d"),
            "unpaid":pick(d,"CapitalExpendituresIncurredButNotYetPaid","i")}
d=load("ORCL")
res["ORCL"]["cip"]=pick(d,"ConstructionInProgressGross","i")
res["ORCL"]["equip"]=pick(d,"MachineryAndEquipmentGross","i")
res["ORCL"]["purch"]=pick(d,"UnrecordedUnconditionalPurchaseObligationBalanceSheetAmount","i",1e9)
d=load("META")
res["META"]["lease_nc"]=pick(d,"UnrecordedUnconditionalPurchaseObligationBalanceSheetAmount","i",1e9)
d=load("GOOGL")
res["GOOGL"]["cip_xbrl"]=pick(d,"ConstructionInProgressGross","i")
json.dump(res,open("/home/claude/out/channel/data/us_data.json","w"),indent=1)
for t in res:
    for k,v in res[t].items(): print(t,k,sorted(v.keys()))
