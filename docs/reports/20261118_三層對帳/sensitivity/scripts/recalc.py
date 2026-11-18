# LibreOffice headless 重算並存回（xlsx），再數錯誤格。用法：python3 recalc.py <xlsx>
import uno, subprocess, time, sys, os
from com.sun.star.beans import PropertyValue
src = os.path.abspath(sys.argv[1]); tmp = "/tmp/sens/recalc_out.xlsx"
port = 2700 + os.getpid() % 200
proc = subprocess.Popen(['soffice', '-env:UserInstallation=file:///tmp/sens/lo_prof2', '--headless', '--invisible', '--norestore',
                         f'--accept=socket,host=localhost,port={port};urp;'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
ctx = None
for i in range(90):
    try:
        local = uno.getComponentContext()
        res = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
        ctx = res.resolve(f'uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext'); break
    except Exception: time.sleep(1)
desk = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
def pv(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p
doc = desk.loadComponentFromURL('file://' + src, '_blank', 0, (pv('Hidden', True),))
doc.calculateAll()
doc.storeToURL('file://' + tmp, (pv('FilterName', 'Calc MS Excel 2007 XML'),))
doc.close(True)
try: desk.terminate()
except Exception: pass
proc.wait(timeout=60)
import openpyxl
wb = openpyxl.load_workbook(tmp, data_only=True)
ERR = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!", "Err:")
bad = [(ws.title, c.coordinate, c.value) for ws in wb for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith(ERR)]
print("錯誤格", len(bad), bad[:10])
os.replace(tmp, src)
