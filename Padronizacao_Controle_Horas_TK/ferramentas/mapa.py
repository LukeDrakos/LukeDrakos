import openpyxl, sys
from openpyxl.utils import get_column_letter as L
wb = openpyxl.load_workbook(sys.argv[1])
out = ['# Mapa de colunas e linhas — EI32 v55', '', 'Gerado automaticamente a partir de `planilhas/Lancamento horas - EI32.xlsx` (v55).',
       '"Tipo": **entrada** = preenchida à mão; **fórmula** = automática (não digitar). A fórmula mostrada é a da linha 2/3 (exemplo), truncada.', '']
def short(v, n=170):
    s = str(v).replace('\n', ' ').replace('|', '¦')
    return s if len(s) <= n else s[:n] + ' …'
def tab(ws, hr, r_ex, last=None):
    out.append(f'## {ws.title} (cabeçalho na linha {hr}, {ws.max_row} linhas, estado: {ws.sheet_state})'); out.append('')
    out.append('| Col | Cabeçalho | Tipo | Oculta | Exemplo (linha %d) |' % r_ex); out.append('|---|---|---|---|---|')
    mc = last or ws.max_column
    for c in range(1, mc + 1):
        h = ws.cell(hr, c).value
        if h is None: continue
        v = ws.cell(r_ex, c).value
        tipo = 'fórmula' if isinstance(v, str) and v.startswith('=') else 'entrada'
        hid = 'sim' if ws.column_dimensions[L(c)].hidden else ''
        out.append(f'| {L(c)} | {short(h,60)} | {tipo} | {hid} | `{short(v)}` |' if v is not None else f'| {L(c)} | {short(h,60)} | {tipo} | {hid} | |')
    out.append('')
tab(wb['Lanc_HORAS'], 1, 2)
tab(wb['Lanc_HE'], 1, 5, 35)
tab(wb['Compensacoes'], 3, 4)
tab(wb['Cadastro'], 2, 3)
tab(wb['Vigencias'], 2, 3)
tab(wb['Analises'], 2, 3)
def grid(ws, maxr, cols):
    out.append(f'## {ws.title} (célula a célula, linhas 1–{maxr})'); out.append('')
    out.append('| Célula | Conteúdo |'); out.append('|---|---|')
    for r in range(1, maxr + 1):
        for c in cols:
            v = ws[f'{c}{r}'].value
            if v is not None: out.append(f'| {c}{r} | `{short(v,200)}` |')
    out.append('')
grid(wb['Gestao_Contrato'], wb['Gestao_Contrato'].max_row, 'BCDFGH')
grid(wb['Painel_De_Controle'], 48, 'BCEFHIJK')
t = wb['Tabelas']
out.append('## Tabelas (blocos)'); out.append('')
for c in range(1, 40):
    if t.cell(1, c).value: out.append(f'- **{L(c)}**: {t.cell(1,c).value} — colunas: ' + ', '.join(str(t.cell(2, k).value) for k in range(c, c + 4) if t.cell(2, k).value))
out.append(''); out.append('### Parâmetros (J:L)'); out.append(''); out.append('| Linha | Parâmetro | Valor | Observação |'); out.append('|---|---|---|---|')
for r in range(3, 30):
    if t.cell(r, 10).value: out.append(f'| {r} | {short(t.cell(r,10).value,60)} | {short(t.cell(r,11).value,60)} | {short(t.cell(r,12).value,120)} |')
out.append(''); out.append('### Regras dos status (A:C) e rótulos do relatório (AB:AC)'); out.append(''); out.append('| Status | Grupo | Conta no absenteísmo? |'); out.append('|---|---|---|')
for r in range(3, 60):
    if t.cell(r, 1).value: out.append(f'| {t.cell(r,1).value} | {t.cell(r,2).value} | {t.cell(r,3).value} |')
out.append(''); out.append('| Status | Como aparece no relatório |'); out.append('|---|---|')
for r in range(3, 60):
    if t.cell(r, 28).value: out.append(f'| {t.cell(r,28).value} | {t.cell(r,29).value} |')
out.append(''); out.append('## Nomes definidos'); out.append('')
for n, d in wb.defined_names.items(): out.append(f'- `{n}` → `{d.attr_text}`')
rel = wb['Relatorio']
out.append(''); out.append('## Relatorio — âncoras'); out.append('')
for c in rel['B'] + rel['C']:
    if isinstance(c.value, str) and not c.value.startswith('=') and (c.value.isupper() or c.value.startswith('(') or 'TOTAL' in c.value):
        out.append(f'- {c.coordinate}: {short(c.value,90)}')
out.append(f'- Área de impressão: {rel.print_area}; quebras manuais após as linhas: {[b.id for b in rel.row_breaks.brk]}')
out.append('- Coluna N = 1 → a macro CriarRelatorio oculta a linha antes de gerar o PDF. Colunas L/M = índice da ocorrência no detalhe.')
open(sys.argv[2], 'w').write('\n'.join(out))
print(len(out))
