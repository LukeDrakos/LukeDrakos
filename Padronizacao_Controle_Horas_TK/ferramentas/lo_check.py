import uno, sys, time, subprocess, os
from com.sun.star.beans import PropertyValue
def pv(n,v): p=PropertyValue(); p.Name=n; p.Value=v; return p
proc=subprocess.Popen(['soffice','--headless','--invisible','--norestore','--accept=socket,host=localhost,port=2003;urp;'],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
for i in range(60):
    try:
        l=uno.getComponentContext(); r=l.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',l)
        ctx=r.resolve('uno:socket,host=localhost,port=2003;urp;StarOffice.ComponentContext'); break
    except Exception: time.sleep(1)
desk=ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop',ctx)
doc=desk.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(sys.argv[1])),'_blank',0,(pv('Hidden',True),))
libs=doc.BasicLibraries
for ln in libs.getElementNames():
    lib=libs.getByName(ln)
    for m in lib.getElementNames():
        code=lib.getByName(m); print(ln,m,len(code),'NomeArquivoRelatorio' in code, 'Auto_Open' in code)
doc.close(True); proc.terminate()
