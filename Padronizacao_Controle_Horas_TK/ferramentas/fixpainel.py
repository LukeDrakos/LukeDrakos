import openpyxl,sys
wb=openpyxl.load_workbook(sys.argv[1]); P=wb['Painel_De_Controle']
n=wb['Lanc_HORAS'].max_row
R=lambda c:f'Lanc_HORAS!${c}$2:${c}${n}'
for col,team in (('C','Elétrica'),('F','Climatização')):
    base=f'({R("B")}=$C$2)*({R("D")}="{team}")*({R("AC")}>0)'
    vv=f'IFERROR({R("V")}*1,0)'
    P[f'{col}33']=f'=SUMPRODUCT({base}*({R("AC")}>={vv})*{R("AC")})'
    P[f'{col}34']=f'=SUMPRODUCT({base}*({R("AC")}<{vv})*{R("AC")})'
    P[f'{col}35']=f'=SUMPRODUCT({base}*({R("AC")}<{vv}))'
    P[f'{col}36']=f'=SUMIFS({R("AM")},{R("B")},$C$2,{R("D")},"{team}")'
wb.save(sys.argv[2]); print('ok')
