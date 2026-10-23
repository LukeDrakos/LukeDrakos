# uso: render.py arquivo.xlsx saida_dir Aba1,Aba2  -> PDF/PNG só das abas pedidas (mantém valores calculados)
import sys, zipfile, re, os, subprocess, shutil
src, outdir, abas = sys.argv[1], sys.argv[2], sys.argv[3].split(',')
os.makedirs(outdir, exist_ok=True)
tmp = os.path.join(outdir, 'r.xlsx')
ocultar = set()
zin = zipfile.ZipFile(src); zout = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
for it in zin.infolist():
    data = zin.read(it.filename)
    if it.filename == 'xl/workbook.xml':
        s = data.decode('utf8')
        def f(m):
            tag = m.group(0); nm = re.search(r'name="([^"]+)"', tag).group(1)
            tag = re.sub(r'\sstate="[^"]+"', '', tag)
            return tag if nm in abas else ''
        s = re.sub(r'<sheet [^>]+/>', f, s)
        s = re.sub(r'activeTab="\d+"', '', s)
        s = re.sub(r'<definedName [^>]*localSheetId[^>]*>.*?</definedName>', '', s)
        data = s.encode('utf8')
    if ocultar and it.filename == alvo:
        x = data.decode('utf8')
        x = re.sub(r'<row r="(\d+)"', lambda m: m.group(0) + (' hidden="1"' if int(m.group(1)) in ocultar else ''), x)
        data = x.encode('utf8')
    zout.writestr(it, data)
zout.close()
subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', outdir, tmp], check=True, capture_output=True)
pdf = tmp[:-5] + '.pdf'
subprocess.run(['pdftoppm', '-png', '-r', '60', pdf, os.path.join(outdir, 'p')], check=True)
print(sorted(os.listdir(outdir)))
