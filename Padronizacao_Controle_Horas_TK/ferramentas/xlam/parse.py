import olefile, struct, sys
from oletools.olevba import decompress_stream
o=olefile.OleFileIO('xl/vbaProject.bin')
d=decompress_stream(bytearray(o.openstream('VBA/dir').read()))
def recs(d):
    i=0; out=[]
    while i<len(d):
        rid,sz=struct.unpack_from('<HI',d,i)
        if rid==0x0009: sz=6
        out.append((rid,i,sz,bytes(d[i+6:i+6+sz]))); i+=6+sz
    return out
R=recs(d)
for rid,i,sz,data in R:
    if rid in (0x19,0x1A,0x31,0x21,0x22,0x2B,0x0F,0x03,0x10): print(hex(rid),i,sz,data[:40])
m=o.openstream('VBA/modEI32').read()
off=[struct.unpack('<I',data)[0] for rid,i,sz,data in R if rid==0x31]
print('offsets',off)
