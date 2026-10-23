import sys, openpyxl, datetime
def h(v):
    if isinstance(v, datetime.timedelta): return v.total_seconds()/3600
    if isinstance(v,(int,float)): return v*24
    return 0
def get(f):
    wb=openpyxl.load_workbook(f,data_only=True); ws=wb["Resumo_Mensal"]
    out={}
    for r in range(5,17):
        m=ws.cell(r,1).value
        out[m.month]=dict(bruto=h(ws.cell(r,3).value), fer=h(ws.cell(r,4).value), comp=h(ws.cell(r,5).value), desc=h(ws.cell(r,7).value), abat=ws.cell(r,8).value or 0, sla=ws.cell(r,9).value or 0, final=ws.cell(r,10).value or 0)
    return out
if __name__=="__main__":
    a=get(sys.argv[1]); b=get(sys.argv[2]) if len(sys.argv)>2 else None
    for m in range(1,13):
        x=a[m]
        line=f"{m:2d} bruto {x['bruto']:8.2f} fer {x['fer']:7.2f} comp {x['comp']:7.2f} desc {x['desc']:8.2f} abat {x['abat']:11.2f}"
        if b:
            y=b[m]; line+=f" | bruto {y['bruto']:8.2f} fer {y['fer']:7.2f} comp {y['comp']:7.2f} desc {y['desc']:8.2f} abat {y['abat']:11.2f} | dif {y['abat']-x['abat']:10.2f}"
        print(line)
