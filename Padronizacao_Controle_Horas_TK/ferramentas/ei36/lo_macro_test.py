"""Roda uma macro do .bas dentro do LibreOffice (modo VBA) numa cópia da planilha e salva o resultado.
uso: python3 lo_macro_test.py modulo.bas planilha.xlsx Macro saida.xlsx
MsgBox é trocado por um log (headless não pode abrir diálogo)."""
import uno, sys, time, subprocess, os, re, shutil
from com.sun.star.beans import PropertyValue

bas, xlsx, macro, out = sys.argv[1:5]
def pv(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p

code = open(bas, encoding='ascii').read().replace('\r\n', '\n')
code = re.sub(r'^Attribute .*\n', '', code, flags=re.M)
code = re.sub(r'\bMsgBox\b', 'LogMsg', code)
code = 'Option VBASupport 1\n' + code + '''
Public gLog As String
Public Function LogMsg(Optional ByVal a As Variant, Optional ByVal b As Variant, Optional ByVal c As Variant) As Long
    gLog = gLog & CStr(a) & " || "
    LogMsg = 1
End Function
Public Function PegaLog() As String
    PegaLog = gLog
End Function
'''
port = 2011
proc = subprocess.Popen(['soffice', '--headless', '--invisible', '--norestore', '--nologo',
                         f'--accept=socket,host=localhost,port={port};urp;'],
                        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
ctx = None
for i in range(90):
    try:
        l = uno.getComponentContext()
        r = l.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', l)
        ctx = r.resolve(f'uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext'); break
    except Exception:
        time.sleep(1)
desk = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
doc = desk.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(xlsx)), '_blank', 0,
                                (pv('Hidden', False), pv('MacroExecutionMode', 4)))
libs = doc.BasicLibraries
if not libs.hasByName('Standard'):
    libs.createLibrary('Standard')
lib = libs.getByName('Standard')
if lib.hasByName('modEI'):
    lib.removeByName('modEI')
lib.insertByName('modEI', code)
sp = doc.getScriptProvider()
t0 = time.time()
err = None
try:
    sp.getScript(f'vnd.sun.star.script:Standard.modEI.{macro}?language=Basic&location=document').invoke((), (), ())
except Exception as e:
    err = repr(e)
log = sp.getScript('vnd.sun.star.script:Standard.modEI.PegaLog?language=Basic&location=document').invoke((), (), ())[0]
print('tempo', round(time.time() - t0, 1), 's')
print('ERRO:', err)
print('LOG:', log)
lib.removeByName('modEI')
doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(out)), (pv('FilterName', 'Calc MS Excel 2007 XML'),))
doc.close(True)
proc.terminate()
