"""v52 — novos status (EI36), adicional não solicitado cobre faltas, aba SLA (manual, igual EI36)."""
import openpyxl, copy, sys, re
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as L
from openpyxl.comments import Comment

SRC, OUT = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(SRC)
TK, NAVY, TXT, CINZA = 'FF0095D8', 'FF1F3864', 'FF4D4D4D', 'FF7F7F7F'
VERDE, VERM, INPUT = 'FFD5F5E3', 'FFFADBD8', 'FFF2F2F2'
fill = lambda c: PatternFill('solid', fgColor=c)
ln = Side(style='thin', color='FFD9D9D9'); B = Border(left=ln, right=ln, top=ln, bottom=ln)
NL = 3884
ADIC = 'Adicional Não Solicitado'
NOVOS = [('Recusa de Atividade S/ Motivo Legítimo', 'Sim'),
         ('Recusa de Atividade C/ Motivo Legítimo', 'Não'),
         ('Ausências Devido a Integração / Início de Atividades', 'Não'),
         ('Ausências Devido a Afastamento por Acidente', 'Não'),
         ('Ausências Devido a Desligamento / Pedido de Demissão', 'Sim')]
NOVOS_AUS = [n for n, _ in NOVOS]
JUST_NOVOS = [n for n, c in NOVOS if c == 'Não']
ENT_CJ = 'Entrada Atrasada ou Saída Antecipada C/ Justificativa'

def soma_justificadas(f):
    """duplica o termo SUMIFS do 'Entrada ... C/ Justificativa' para cada status justificado novo"""
    i = f.find('"' + ENT_CJ + '")')
    if i < 0: return f
    ini = f.rfind('SUMIFS(', 0, i)
    fim = i + len(ENT_CJ) + 3
    termo = f[ini:fim]
    extra = ''.join('+' + termo.replace(ENT_CJ, n) for n in JUST_NOVOS)
    return f[:fim] + extra + f[fim:]

# ------------------------------------------------------------------ 1. Tabelas: novos status
T = wb['Tabelas']
ult = max(r for r in range(3, 150) if T.cell(r, 1).value)
existentes = {T.cell(r, 1).value for r in range(3, ult + 1)}
novos = [(n, 'Absenteísmo', c) for n, c in NOVOS] + [(ADIC, 'Controle', 'Não')]
r = ult + 1
for st, gr, ct in novos:
    if st in existentes: continue
    for c, v in ((1, st), (2, gr), (3, ct)):
        x = T.cell(r, c); x.value = v; x._style = copy.copy(T.cell(ult, c)._style)
    r += 1

# ------------------------------------------------------------------ 2. Lanc_HORAS
lh = wb['Lanc_HORAS']
lista_aus = '$J{r}="Remanejamento"'
for k in range(2, NL + 1):
    v = lh[f'V{k}'].value
    alvo = f'$J{k}="Remanejamento"),$AD{k}'
    assert alvo in v, k
    extra = ''.join(f',$J{k}="{s}"' for s in NOVOS_AUS)
    lh[f'V{k}'] = v.replace(alvo, f'$J{k}="Remanejamento"{extra}),$AD{k}', 1)
    u = lh[f'U{k}'].value
    a2 = f'$J{k}="Horas Recebidas"),'
    assert a2 in u, k
    lh[f'U{k}'] = u.replace(a2, f'$J{k}="Horas Recebidas",$J{k}="{ADIC}"),', 1)
    ab = lh[f'AB{k}'].value
    alvo_ab = f'$J{k}="{ENT_CJ}"),'
    assert alvo_ab in ab, k
    lh[f'AB{k}'] = ab.replace(alvo_ab, f'$J{k}="{ENT_CJ}"' + ''.join(f',$J{k}="{n}"' for n in JUST_NOVOS) + '),', 1)
    aa = lh[f'AA{k}'].value
    alvo_aa = f'$J{k}="{ENT_CJ}"),"Justificada: '
    assert alvo_aa in aa, k
    aa = aa.replace(alvo_aa, f'$J{k}="{ENT_CJ}"' + ''.join(f',$J{k}="{n}"' for n in JUST_NOVOS) + '),"Justificada: ', 1)
    ini = '=IF($A{0}="","",TRIM('.format(k)
    assert aa.startswith(ini), k
    lh[f'AA{k}'] = aa.replace(ini, ini + 'IF($J{0}="{1}","Adicional não solicitado: trabalhou "&TEXT(N($U{0}),"[h]:mm")&"h; essas horas abatem o absenteísmo do mês da equipe (1:1, como compensação). ","")&'.format(k, ADIC), 1)

# ------------------------------------------------------------------ 3. Analises: adicional abate o absenteísmo do mês direto
A = wb['Analises']
for col, tit in (('BF', 'Horas Adicional não\nsolicitado (h)'), ('BG', 'Horas Adicional\nUTILIZADAS (h)')):
    A[f'{col}2'] = tit; A[f'{col}2']._style = copy.copy(A['BE2']._style)
    A.column_dimensions[col].width = 16
for base, team in ((3, 'Elétrica'), (19, 'Climatização')):
    if base == 19:
        for col in ('BF', 'BG'):
            A[f'{col}18'] = A[f'{col}2'].value; A[f'{col}18']._style = copy.copy(A['BE18']._style)
    for rr in range(base, base + 12):
        A[f'BF{rr}'] = (f'=SUMIFS(Lanc_HORAS!$U$2:$U${NL},Lanc_HORAS!$B$2:$B${NL},$A{rr},Lanc_HORAS!$D$2:$D${NL},"{team}",'
                        f'Lanc_HORAS!$J$2:$J${NL},"{ADIC}")')
        A[f'BG{rr}'] = f'=MIN(N($BF{rr}),MAX(N($AU{rr})-N($AW{rr}),0))'
        for col in ('BF', 'BG'): A[f'{col}{rr}']._style = copy.copy(A[f'BE{rr}']._style)
        assert A[f'B{rr}'].value == f'=MAX(N($AU{rr})-N($AW{rr}),0)', A[f'B{rr}'].value
        A[f'B{rr}'] = f'=MAX(N($AU{rr})-N($AW{rr})-N($BG{rr}),0)'
        assert A[f'AT{rr}'].value == f'=IF(N($AU{rr})<=0,0,N($AV{rr})*N($AW{rr})/N($AU{rr}))', A[f'AT{rr}'].value
        A[f'AT{rr}'] = f'=IF(N($AU{rr})<=0,0,N($AV{rr})*(N($AW{rr})+N($BG{rr}))/N($AU{rr}))'

for rr in list(range(3, 15)) + list(range(19, 31)):
    A[f'V{rr}'] = soma_justificadas(A[f'V{rr}'].value)
for rr in range(38, 100):
    v = A[f'P{rr}'].value
    if isinstance(v, str) and v.startswith('=') and ENT_CJ in v:
        A[f'P{rr}'] = soma_justificadas(v)

def soma_novos(col, crit_mes, team):
    return '+'.join(f'SUMIFS(Lanc_HORAS!${col}$2:${col}${NL},Lanc_HORAS!$B$2:$B${NL},{crit_mes},Lanc_HORAS!$D$2:$D${NL},"{team}",Lanc_HORAS!$J$2:$J${NL},"{s}")'
                    for s in NOVOS_AUS)
LBL = 'Outras (recusa, integração, afastamento, desligamento)'

# ------------------------------------------------------------------ 4. Gestão: linha "Outras" + SLA
G = wb['Gestao_Contrato']
for cols, team in (('BCD', 'Elétrica'), ('FGH', 'Climatização')):
    lab, h, val = cols
    # desce compensadas (19) e total (20) para 20 e 21
    for src, dst in ((20, 21), (19, 20)):
        for c in cols:
            G[f'{c}{dst}'].value = G[f'{c}{src}'].value
            G[f'{c}{dst}']._style = copy.copy(G[f'{c}{src}']._style)
    for c in cols: G[f'{c}19']._style = copy.copy(G[f'{c}18']._style)
    G[f'{lab}19'] = LBL
    G[f'{h}19'] = '=' + soma_novos('AM', '$B$3', team)
    G[f'{val}19'] = '=' + soma_novos('W', '$B$3', team)
    G[f'{h}21'] = f'=MAX(SUM({h}11:{h}19)-{h}20,0)'
    G[f'{val}21'] = f'=MAX(SUM({val}11:{val}19)-{val}20,0)'
for lab, h, base in (('B', 'C', 3), ('F', 'G', 19)):
    G[f'{lab}20'] = '( − ) Compensadas + adicional não solicitado'
    G[f'{h}20'] = (f'=IFERROR(INDEX(Analises!$AW${base}:$AW${base+11},MATCH($B$3,Analises!$A${base}:$A${base+11},0))'
                   f'+INDEX(Analises!$BG${base}:$BG${base+11},MATCH($B$3,Analises!$A${base}:$A${base+11},0)),0)')
G.row_dimensions[21].height = G.row_dimensions[20].height
G.row_dimensions[20].height = 15
# SLA na linha 8
for lab, v, base in (('B', 'C', 3), ('F', 'G', 19)):
    G[f'{lab}8'] = 'Desconto SLA (já descontado do valor a pagar)'
    G[f'{lab}8'].font = Font(name='Arial', size=8, color=CINZA, italic=True)
    G[f'{lab}8'].alignment = Alignment(horizontal='right', vertical='center')
    G[f'{v}8'] = f'=IFERROR(INDEX(Analises!$AG${base}:$AG${base+11},MATCH($B$3,Analises!$A${base}:$A${base+11},0)),0)'
    G[f'{v}8'].font = Font(name='Arial', size=9, bold=True, color=TXT)
    G[f'{v}8'].number_format = '"R$ "#,##0.00'
    G[f'{v}8'].alignment = Alignment(horizontal='center', vertical='center')
# atalho SLA no menu
G['I2'] = 'SLA'
G['I2']._style = copy.copy(G['H2']._style)
G['I2'].hyperlink = openpyxl.worksheet.hyperlink.Hyperlink(ref='I2', location="'SLA'!A1", display='SLA')

# ------------------------------------------------------------------ 5. Painel: linha "Outras" antes do TOTAL
P = wb['Painel_De_Controle']
for tot, team in ((18, 'Elétrica'), (35, 'Climatização')):
    first = 6 if team == 'Elétrica' else 23
    for c in 'HIJK':
        P[f'{c}{tot+1}'].value = P[f'{c}{tot}'].value
        P[f'{c}{tot+1}']._style = copy.copy(P[f'{c}{tot}']._style)
        P[f'{c}{tot}']._style = copy.copy(P[f'{c}{tot-1}']._style)
    P[f'H{tot}'] = LBL
    P[f'I{tot}'] = '=' + soma_novos('V', '$C$2', team) + '+' + soma_novos('M', '$C$2', team)
    P[f'J{tot}'] = '=' + soma_novos('AC', '$C$2', team)
    P[f'K{tot}'] = '=' + soma_novos('W', '$C$2', team)
    for c in 'IJK':
        P[f'{c}{tot+1}'] = f'=SUM({c}{first}:{c}{tot})'
    P.row_dimensions[tot + 1].height = P.row_dimensions[tot].height

# ------------------------------------------------------------------ 6. Aba SLA
S = wb.create_sheet('SLA', wb.sheetnames.index('Analises') + 1)
S.sheet_properties.tabColor = TK
S.sheet_view.showGridLines = False
W = {'A': 1.5, 'B': 40, 'C': 11, 'D': 12}
for k, w in W.items(): S.column_dimensions[k].width = w
for i in range(12): S.column_dimensions[L(5 + i)].width = 11
S.column_dimensions['Q'].width = 44
def cel(a, v=None, size=9, b=False, color=TXT, fl=None, h='left', wrap=False, fmt=None, brd=True):
    x = S[a]
    if v is not None: x.value = v
    x.font = Font(name='Arial', size=size, bold=b, color=color)
    x.alignment = Alignment(horizontal=h, vertical='center', wrap_text=wrap)
    if fl: x.fill = fill(fl)
    if brd: x.border = B
    if fmt: x.number_format = fmt
    return x
def barra(r, txt, c1='B', c2='Q', color=TK):
    S.merge_cells(f'{c1}{r}:{c2}{r}'); cel(f'{c1}{r}', txt, size=10, b=True, color='FFFFFFFF', fl=color, brd=False)
    for c in range(openpyxl.utils.column_index_from_string(c1), openpyxl.utils.column_index_from_string(c2) + 1):
        S.cell(r, c).fill = fill(color)
    S.row_dimensions[r].height = 20

S.merge_cells('B1:Q1')
cel('B1', '=HYPERLINK("#Gestao_Contrato!B3","EI32 — SLA: DESCONTO E INDICADORES   ◄ Gestão")', size=16, b=True, color=TK, brd=False)
S.row_dimensions[1].height = 30
S.merge_cells('B2:Q2')
cel('B2', 'Igual à EI36: o desconto de SLA é informado à mão, por mês e equipe (células cinza), e já sai do valor final a pagar '
          '(Gestão, Análises e Relatório). Os indicadores servem para registrar a meta e o realizado de cada mês; '
          'meta em branco = sem cor. Verde = dentro da meta; vermelho = fora.', size=8, color=CINZA, wrap=True, brd=False)
S.row_dimensions[2].height = 30

r = 4
barra(r, 'DESCONTO SLA (R$) — informar por mês', color=NAVY); r += 1
for c, t in (('B', 'Mês'), ('C', 'Elétrica (R$)'), ('D', 'Climatização (R$)')):
    cel(f'{c}{r}', t, b=True, fl='FFF2F2F2', h='center', wrap=True)
S.merge_cells(f'E{r}:Q{r}'); cel(f'E{r}', 'Referência / motivo (indicador, período, aprovação)', b=True, fl='FFF2F2F2', h='center')
S.row_dimensions[r].height = 26; r += 1
SLA0 = r
for i in range(12):
    cel(f'B{r}', f'=Analises!$A${3+i}', h='left', fmt='mmmm/yyyy')
    cel(f'C{r}', None, fl=INPUT, h='right', fmt='#,##0.00')
    cel(f'D{r}', None, fl=INPUT, h='right', fmt='#,##0.00')
    S.merge_cells(f'E{r}:Q{r}'); cel(f'E{r}', None, fl=INPUT, size=8, wrap=True)
    for c in range(6, 18): S.cell(r, c).fill = fill(INPUT); S.cell(r, c).border = B
    S.row_dimensions[r].height = 16; r += 1
cel(f'B{r}', 'TOTAL NO ANO', b=True, fl='FFD9D9D9')
cel(f'C{r}', f'=SUM(C{SLA0}:C{r-1})', b=True, fl='FFD9D9D9', h='right', fmt='#,##0.00')
cel(f'D{r}', f'=SUM(D{SLA0}:D{r-1})', b=True, fl='FFD9D9D9', h='right', fmt='#,##0.00')
r += 2
dv_sent = DataValidation(type='list', formula1='"≥ meta,≤ meta"', allow_blank=True)
S.add_data_validation(dv_sent)
funcs = {
    'Elétrica': ['TÉC. ELETRICISTA I', 'TÉC. ELETRICISTA II', 'TÉC. TELECOM', 'ENCARREGADO DA ELÉTRICA'],
    'Climatização': ['TÉC. CLIMATIZAÇÃO I', 'TÉC. CLIMATIZAÇÃO II', 'ENCARREGADO DA CLIMATIZAÇÃO'],
}
# funções ativas reais do cadastro, se existirem
C = wb['Cadastro']
for team in funcs:
    reais = sorted({C.cell(k, 5).value for k in range(3, 151)
                    if C.cell(k, 4).value == team and C.cell(k, 5).value and str(C.cell(k, 10).value).strip().lower() == 'ativo'})
    if reais: funcs[team] = reais
for team in ('Elétrica', 'Climatização'):
    barra(r, f'INDICADORES DE SLA — {team.upper()}'); r += 1
    hdr = [('B', 'Indicador'), ('C', 'Sentido'), ('D', 'Meta')] + [(L(5 + i), None) for i in range(12)]
    for c, t in hdr: cel(f'{c}{r}', t, b=True, fl='FFF2F2F2', h='center', wrap=True)
    for i in range(12):
        S[f'{L(5+i)}{r}'] = f'=Analises!$A${3+i}'; S[f'{L(5+i)}{r}'].number_format = 'mmm/yy'
    cel(f'Q{r}', 'Observação', b=True, fl='FFF2F2F2', h='center')
    S.row_dimensions[r].height = 22; r += 1
    i0 = r
    if team == 'Elétrica':
        base_ind = [('Nº de atendimentos — Elétrica predial', '≥ meta', 169), ('Backlog — Elétrica predial', '≤ meta', 178)]
    else:
        base_ind = [('Nº de atendimentos — Climatização', '≥ meta', 119), ('Nº de atendimentos — Refrigeração', '≥ meta', 119),
                    ('Backlog — Climatização', '≤ meta', 10.5), ('Backlog — Refrigeração', '≤ meta', 10.5)]
    linhas = base_ind + [(f'Quadro — {f}', '≥ meta', None) for f in funcs[team]]
    for nome, sent, meta in linhas:
        cel(f'B{r}', nome, fl=INPUT)
        cel(f'C{r}', sent, fl=INPUT, h='center'); dv_sent.add(f'C{r}')
        cel(f'D{r}', meta, fl=INPUT, h='center', fmt='#,##0.0;;' if isinstance(meta, float) else '#,##0')
        for i in range(12): cel(f'{L(5+i)}{r}', None, fl=INPUT, h='center')
        cel(f'Q{r}', None, fl=INPUT, size=8, wrap=True)
        S.row_dimensions[r].height = 16; r += 1
    for extra in range(3):   # linhas livres para outros indicadores
        cel(f'B{r}', None, fl=INPUT); cel(f'C{r}', None, fl=INPUT, h='center'); dv_sent.add(f'C{r}')
        cel(f'D{r}', None, fl=INPUT, h='center')
        for i in range(12): cel(f'{L(5+i)}{r}', None, fl=INPUT, h='center')
        cel(f'Q{r}', None, fl=INPUT, size=8, wrap=True)
        S.row_dimensions[r].height = 16; r += 1
    rng = f'E{i0}:P{r-1}'
    ok = f'AND($D{i0}<>"",E{i0}<>"",OR(AND(LEFT($C{i0},1)="≥",E{i0}>=$D{i0}),AND(LEFT($C{i0},1)="≤",E{i0}<=$D{i0})))'
    nok = f'AND($D{i0}<>"",E{i0}<>"",OR(AND(LEFT($C{i0},1)="≥",E{i0}<$D{i0}),AND(LEFT($C{i0},1)="≤",E{i0}>$D{i0})))'
    S.conditional_formatting.add(rng, FormulaRule(formula=[ok], fill=fill(VERDE)))
    S.conditional_formatting.add(rng, FormulaRule(formula=[nok], fill=fill(VERM)))
    r += 1
barra(r, 'METAS DE SLA DE REFERÊNCIA — todas as divisões (fonte: tabela de metas do contrato)', color=NAVY); r += 1
for c, t in (('B', 'Divisão'), ('C', 'Equipe'), ('D', 'Meta atend.'), ('E', 'Meta backlog')):
    cel(f'{c}{r}', t, b=True, fl='FFF2F2F2', h='center', wrap=True)
S.row_dimensions[r].height = 22; r += 1
for dv_, eq_, m1, m2 in [('EI36 - Serralheria', 'Serralheria', 276, 321),
                         ('EI32 - Elétrica Predial e Climatização', 'Elétrica predial', 169, 178),
                         ('EI32 - Elétrica Predial e Climatização', 'Refrigeração', 119, 10.5),
                         ('EI32 - Elétrica Predial e Climatização', 'Climatização', 119, 10.5),
                         ('EI31 - Manutenção Civil e Jardinagem', 'Civil', 250, 240)]:
    cel(f'B{r}', dv_); cel(f'C{r}', eq_); cel(f'D{r}', m1, h='center', fmt='#,##0'); cel(f'E{r}', m2, h='center', fmt='#,##0.0;;' if isinstance(m2, float) else '#,##0')
    S.row_dimensions[r].height = 16; r += 1
S.column_dimensions['C'].width = 16
S.freeze_panes = 'C4'
S.page_setup.orientation = 'landscape'; S.page_setup.fitToWidth = 1; S.page_setup.fitToHeight = 0
S.sheet_properties.pageSetUpPr.fitToPage = True

# Analises AG = desconto SLA informado
for base, col in ((3, 'C'), (19, 'D')):
    for i in range(12):
        A[f'AG{base+i}'] = f'=N(SLA!${col}${SLA0+i})'

# ------------------------------------------------------------------ 7. Relatório: linha do SLA abaixo do resumo
R = wb['Relatorio']
R.merge_cells('B9:J9')
R['B9'] = ('="Desconto SLA do mês (já descontado do valor final):  Elétrica R$ "&FIXED(IFERROR(INDEX(Analises!$AG$3:$AG$14,MATCH(Gestao_Contrato!$B$3,Analises!$A$3:$A$14,0)),0),2)'
           '&"   ·   Climatização R$ "&FIXED(IFERROR(INDEX(Analises!$AG$19:$AG$30,MATCH(Gestao_Contrato!$B$3,Analises!$A$19:$A$30,0)),0),2)')
R['B9'].font = Font(name='Arial', size=9, bold=True, color=NAVY)
R['B9'].alignment = Alignment(horizontal='left', vertical='center', indent=1)
R['N9'] = ('=IF(IFERROR(INDEX(Analises!$AG$3:$AG$14,MATCH(Gestao_Contrato!$B$3,Analises!$A$3:$A$14,0)),0)'
           '+IFERROR(INDEX(Analises!$AG$19:$AG$30,MATCH(Gestao_Contrato!$B$3,Analises!$A$19:$A$30,0)),0)=0,1,"")')
R.row_dimensions[9].height = 18

# ------------------------------------------------------------------ 8. Como usar
CU = wb['Como_Usar']
CU['C11'].value = (CU['C11'].value or '') + ' Inclui os status novos (recusa, integração, afastamento, desligamento, adicional não solicitado).'
CU['C24'].value = 'Gestao_Contrato no mês certo (B3), conferida com a aba Analises, e desconto SLA informado na aba SLA (se houver).'

wb.save(OUT)
print('ok SLA0', SLA0, 'novos status ate linha', r)
