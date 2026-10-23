"""v53 — padrão visual único (não altera valores nem fórmulas)."""
import openpyxl, sys, copy
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI
from openpyxl.worksheet.dimensions import ColumnDimension

SRC, OUT = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(SRC)
TK, NAVY, TXT, CINZA, BORDA = 'FF0095D8', 'FF1F3864', 'FF4D4D4D', 'FF7F7F7F', 'FFD9D9D9'
VERDE, VERM, ZEBRA = 'FFD5F5E3', 'FFFADBD8', 'FFF5F8FB'
fill = lambda c: PatternFill('solid', fgColor=c)
fino = Side(style='thin', color=BORDA)
BOX = Border(left=fino, right=fino, top=fino, bottom=fino)
branco = Side(style='thin', color='FFFFFFFF')
BOXH = Border(left=branco, right=branco, top=branco, bottom=branco)
CINZA_ENTRADA = PatternFill('solid', fgColor='FFF2F2F2')

# ------------------------------------------------------------ cores das guias por grupo
grupos = {
    'Como_Usar': 'FF7FC4E8',
    'Gestao_Contrato': TK, 'Painel_De_Controle': TK, 'Relatorio': TK, 'Analises': TK, 'SLA': TK,
    'Lanc_HORAS': NAVY, 'Lanc_HE': NAVY, 'Compensacoes': NAVY, 'Conferencia_Ponto': NAVY,
    'Cadastro': 'FF8C8C8C', 'Vigencias': 'FF8C8C8C', 'Tabelas': 'FF8C8C8C',
}
for n, c in grupos.items():
    ws = wb[n]
    ws.sheet_properties.tabColor = c
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85 if c in (NAVY, 'FF8C8C8C') else 90

# ------------------------------------------------------------ abas de lançamento / cadastro
# (aba, linha do cabeçalho, última coluna, congelar, largura máxima por coluna)
dados = [
    ('Lanc_HORAS', 1, 'AX', 'F2', {'X': 60, 'AA': 90, 'AV': 70, 'Z': 55, 'J': 46}),
    ('Lanc_HE', 1, 'AI', 'F2', {'AA': 70, 'AH': 70}),
    ('Compensacoes', 3, 'O', 'C4', {'G': 50}),
    ('Cadastro', 2, 'K', 'C3', {}),
    ('Vigencias', 2, 'H', 'C3', {}),
]
def eh_entrada(ws, hr, c):
    vals = [ws.cell(r, c).value for r in range(hr + 1, min(ws.max_row, hr + 400) + 1)]
    return not any(isinstance(v, str) and v.startswith('=') for v in vals)

for nome, hr, ultc, frz, caps in dados:
    ws = wb[nome]
    last = CI(ultc)
    nrow = ws.max_row
    entrada = {c: eh_entrada(ws, hr, c) for c in range(1, last + 1)}
    # cabeçalho: azul TK = você preenche; marinho = automático
    for c in range(1, last + 1):
        h = ws.cell(hr, c)
        if h.value is None: continue
        h.fill = fill(TK if entrada[c] else NAVY)
        h.font = Font(name='Arial', size=9, bold=True, color='FFFFFFFF')
        h.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        h.border = BOXH
    ws.row_dimensions[hr].height = max(ws.row_dimensions[hr].height or 0, 38)
    # corpo: fonte, alinhamento vertical e bordas finas (preserva preenchimento e formato numérico)
    f9 = Font(name='Arial', size=9, color=TXT)
    for row in ws.iter_rows(min_row=hr + 1, max_row=nrow, min_col=1, max_col=last):
        for x in row:
            fo = x.font
            if not (fo.b or (fo.color is not None and fo.color.rgb not in (None, TXT, 'FF000000'))):
                x.font = f9
            al = x.alignment
            x.alignment = Alignment(horizontal=al.horizontal, vertical='center', wrap_text=al.wrap_text, indent=al.indent)
            x.border = BOX
            if x.fill is not None and x.fill.fill_type == 'solid' and x.fill.fgColor.rgb == 'FFD9D9D9':
                x.fill = CINZA_ENTRADA
    if nome in ('Lanc_HORAS', 'Lanc_HE'):
        for r in range(hr + 1, nrow + 1):
            ws.row_dimensions[r].height = 15
    # larguras
    for col, w in caps.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = frz

# Lanc_HORAS: colunas depois de AX com largura 255 -> padrão
lh = wb['Lanc_HORAS']
for k in list(lh.column_dimensions.keys()):
    d = lh.column_dimensions[k]
    if (d.min or CI(k)) > CI('AX'):
        del lh.column_dimensions[k]
lh.column_dimensions['AY'] = ColumnDimension(lh, index='AY', width=3)

# ------------------------------------------------------------ títulos das abas (padrão único)
titulos = {'Compensacoes': 'A1', 'Cadastro': 'A1', 'Vigencias': 'A1', 'Conferencia_Ponto': 'B1',
           'Como_Usar': 'B1', 'Relatorio': 'B1', 'SLA': 'B1'}
for n, a in titulos.items():
    c = wb[n][a]
    c.font = Font(name='Arial', size=16, bold=True, color=TK, underline=c.font.u)
    wb[n].row_dimensions[c.row].height = max(wb[n].row_dimensions[c.row].height or 0, 30)
for n in ('Compensacoes', 'Cadastro', 'Vigencias'):
    ws = wb[n]
    a2 = ws['A2']
    if a2.value:
        a2.font = Font(name='Arial', size=9, italic=True, color=CINZA)
        a2.alignment = Alignment(vertical='center', wrap_text=False)

# ------------------------------------------------------------ Tabelas: cabeçalhos e larguras
T = wb['Tabelas']
for c in range(1, CI('AG') + 1):
    h = T.cell(2, c)
    if h.value is None: continue
    h.fill = fill(TK); h.font = Font(name='Arial', size=9, bold=True, color='FFFFFFFF')
    h.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True); h.border = BOXH
    t = T.cell(1, c)
    if t.value:
        t.fill = fill(NAVY); t.font = Font(name='Arial', size=10, bold=True, color='FFFFFFFF')
        t.alignment = Alignment(horizontal='left', vertical='center', indent=1)
T.row_dimensions[1].height = 22; T.row_dimensions[2].height = 32
for col, w in {'K': 45, 'L': 60}.items():
    T.column_dimensions[col].width = w
for row in T.iter_rows(min_row=3, max_row=150, min_col=1, max_col=CI('AG')):
    for x in row:
        if x.value is not None:
            if not x.font.b: x.font = Font(name='Arial', size=9, color=TXT)
            x.border = BOX
            x.alignment = Alignment(horizontal=x.alignment.horizontal, vertical='center', wrap_text=x.alignment.wrap_text)

# ------------------------------------------------------------ Gestão: menu legível + regra do contador (H4)
G = wb['Gestao_Contrato']
for r, cor in ((2, TK), (4, NAVY)):
    for c in range(2, 9):
        x = G.cell(r, c)
        if x.value is None or (r == 4 and c == 8): continue
        x.fill = fill(cor)
        x.font = Font(name='Arial', size=9, bold=True, color='FFFFFFFF', underline=None)
        x.alignment = Alignment(horizontal='center', vertical='center')
        x.border = BOXH
d3 = G['D3']
d3.fill = fill(TK); d3.font = Font(name='Arial', size=9, bold=True, color='FFFFFFFF'); d3.alignment = Alignment(horizontal='center', vertical='center')
G['H4'].font = Font(name='Arial', size=9, bold=True, color=TXT)
G['H4'].alignment = Alignment(horizontal='center', vertical='center')
G.conditional_formatting.add('H4', FormulaRule(formula=['SUM(Painel_De_Controle!$N$17:$Y$17)>0'], fill=fill(VERM)))
G.conditional_formatting.add('H4', FormulaRule(formula=['SUM(Painel_De_Controle!$N$17:$Y$17)=0'], fill=fill(VERDE)))

# ------------------------------------------------------------ Como usar: legenda de cores
CU = wb['Como_Usar']
for row in CU.iter_rows(min_row=1, max_row=CU.max_row, min_col=2, max_col=2):
    x = row[0]
    if x.value == 'Cores':
        CU.cell(x.row, 3).value = ('Guias: azul = painéis e relatórios; marinho = lançamentos; cinza = cadastros e tabelas. '
                                   'Cabeçalho azul = você preenche; cabeçalho marinho = automático (não digite por cima). '
                                   'Verde/vermelho só no fundo das células de situação.')

wb.save(OUT)
print('ok')
