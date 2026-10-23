import sys, openpyxl, datetime
def h(v):
    if isinstance(v, datetime.timedelta): return v.total_seconds()/3600
    if isinstance(v,(int,float)): return v*24
    return 0
a=openpyxl.load_workbook(sys.argv[1],data_only=True)["Lancamentos"]
b=openpyxl.load_workbook(sys.argv[2],data_only=True)["Lancamentos"]
for r in range(4,3490):
    if a[f"A{r}"].value is None: continue
    va=a[f"AD{r}"].value or 0; vb=(b[f"AD{r}"].value or 0)+(b[f"BA{r}"].value or 0)
    ua=h(a[f"U{r}"].value); ub=h(b[f"U{r}"].value)
    if abs(va-vb)>0.005 or abs(ua-ub)>0.001:
        if sys.argv[3:] and b[f"A{r}"].value.month not in map(int,sys.argv[3:]): continue
        print(r, b[f"A{r}"].value.date(), b[f"B{r}"].value, "|", b[f"C{r}"].value, "| jorn", round(ua,2), "->", round(ub,2), "| desc", round(h(a[f'AC{r}'].value),2), "->", round(h(b[f'AC{r}'].value),2), "| fer", round(h(b[f'Z{r}'].value),2), "comp", round(h(b[f'AA{r}'].value),2), "| R$", round(va,2), "->", round(vb,2), "VF", b[f"BA{r}"].value)
