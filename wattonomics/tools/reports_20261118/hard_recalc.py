import uno, subprocess, time, sys, os
from com.sun.star.beans import PropertyValue
src=os.path.abspath(sys.argv[1]); dst=os.path.abspath(sys.argv[2])
port=2002+os.getpid()%500
proc=subprocess.Popen(['soffice','--headless','--invisible','--norestore',f'--accept=socket,host=localhost,port={port};urp;'],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
ctx=None
for i in range(60):
    try:
        local=uno.getComponentContext()
        resolver=local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',local)
        ctx=resolver.resolve(f'uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext'); break
    except Exception: time.sleep(1)
smgr=ctx.ServiceManager
desktop=smgr.createInstanceWithContext('com.sun.star.frame.Desktop',ctx)
def pv(n,v):
    p=PropertyValue(); p.Name=n; p.Value=v; return p
doc=desktop.loadComponentFromURL('file://'+src,'_blank',0,(pv('Hidden',True),))
doc.calculateAll()
doc.storeToURL('file://'+dst,(pv('FilterName','Calc MS Excel 2007 XML'),))
doc.close(True)
try: desktop.terminate()
except Exception: pass
proc.wait(timeout=30)
