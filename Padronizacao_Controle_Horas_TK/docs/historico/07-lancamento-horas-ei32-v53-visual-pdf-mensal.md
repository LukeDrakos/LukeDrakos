# EI32 v53 — padrão visual, um PDF por mês e correções (22/10/2026)

**Entrega final:** `Controle_de_Horas_TK_v53.zip`, uma pasta única chamada "Controle de Horas TK - pacote v53". Ela substitui o `EI32_v53.zip` enviado antes.

O pacote foi feito sobre o arquivo do Lucas de 22/10 às 08:37, que já tinha lançamentos de outubro e o suplemento instalado e funcionando.

## Pacote de pasta única (pedido do Lucas: "baixar e executar")

| Arquivo | O que é |
|---|---|
| `Controle de Horas TK.xlam` | O suplemento do Lucas, já com o módulo v53 dentro. Não precisa importar o .bas. |
| `Lancamento horas - EI32.xlsx` | Planilha v53 |
| `Pontos 2026.xlsx` | Base de pontos |
| `INSTALAR.vbs` | Instalador de dois cliques |
| `Modulo_Controle_Horas.bas` | Cópia de segurança do código |
| `LEIA_PRIMEIRO.txt` | Instruções |
| `Relatorios\01 - EI31`, `02 - EI32`, `03 - EI36` | Pastas dos PDFs |

**O que o INSTALAR.vbs faz:**

1. Pede para fechar o Excel, se estiver aberto.
2. Desbloqueia os arquivos (Unblock-File).
3. Grava a pasta como Local Confiável do Excel (HKCU\…\Office\16.0 e 15.0\Excel\Security\Trusted Locations\ControleHorasTK, incluindo subpastas).
4. Cria as pastas dos relatórios.
5. Ativa o suplemento (`AddIns.Add` + `Installed = True`).
6. Abre cada "Lancamento horas*.xlsx" da pasta, roda `InstalarBotoes` e salva.

**Como o .xlam foi atualizado fora do Excel ("VBA purging"):**

- O código-fonte do modEI32 foi substituído e comprimido no formato MS-OVBA. O compressor foi escrito em Python e validado com o descompressor do oletools.
- No registro do módulo, MODULEOFFSET passou a 0.
- `_VBA_PROJECT` virou só o cabeçalho com a versão 0xFFFF, o que força o Excel a recompilar a partir do código-fonte.
- Os streams `__SRP_*` foram removidos.
- O arquivo foi regravado com a biblioteca `cfb` do SheetJS, que acrescenta um stream "Sh33tJ5" inofensivo.
- **Validação:** o olevba extrai um código idêntico ao `m53.bas`, e o LibreOffice importa o módulo com `NomeArquivoRelatorio` e `Auto_Open`. A recompilação dentro do Excel só pode ser confirmada no computador do Lucas. Se algo falhar, o .bas reserva fica no pacote.

**Uso:** feche o Excel, copie tudo para a pasta oficial "01 - Padronização Definitiva" (substituindo os arquivos) e dê dois cliques no INSTALAR.vbs. Nos outros computadores, basta rodar o INSTALAR.vbs.

## Ajustes do Lucas preservados no .bas

Na Gestão, D3 tem o link HYPERLINK para a aba SLA, e os botões ficam em G3 (Atualizar Pontos) e H3 (Criar Relatório).

O .xlam que o Lucas enviou ainda tinha o código da v52 com os botões em D3:E3 e G3:H3. Ou seja, ele editou o .bas, mas não reimportou no suplemento.

## PDF: um arquivo por mês (decisão: pasta por ano)

- **Caminho:** `Relatorios\<nn - EIxx>\<ano>\<mm - Mês> - EIxx.pdf`. Exemplo: `Relatorios\02 - EI32\2026\09 - Setembro - EI32.pdf`.
- **Funções:** `NomeArquivoRelatorio` e `NomeMes`.
  - Criam a pasta do ano.
  - Apagam o PDF do mesmo mês para substituí-lo.
  - Se o arquivo estiver aberto, avisam e não geram.
  - Apagam o formato antigo `EIxx_Relatorio_aaaa-mm.pdf` do mesmo mês.
- Testado no LibreOffice, inclusive o nome "03 - Março".

## Padrão visual (sem mudar valores nem fórmulas)

- **Cores das guias:**
  - azul TK: Gestão, Painel, Relatório, Análises e SLA;
  - marinho: Lanc_HORAS, Lanc_HE, Compensações e Conferência;
  - cinza: Cadastro, Vigências e Tabelas;
  - azul claro: Como_Usar.
- **Cabeçalho das abas de dados:** azul TK quando a coluna é preenchida à mão, marinho quando é automática (detectado pela presença de fórmula).
- **Corpo:** Arial 9, bordas finas, linhas de 15 pt na Lanc_HORAS e na Lanc_HE. O cinza das células de preencher passou de D9D9D9 para F2F2F2.
- **Larguras:**
  - Lanc_HORAS: X 135 → 60, AA 202 → 90, AV 182 → 70; da coluna AY em diante, a largura 255 voltou ao padrão;
  - Lanc_HE: AA e AH → 70;
  - Tabelas: K → 45, L → 60.
- **Congelamento:** F2 na Lanc_HORAS e na Lanc_HE (antes estava em F414 e F18).
- **Gestão:** menu legível e regra do contador H4 recriada. A versão anterior estava em extensão x14, que o openpyxl descarta.
- As listras alternadas foram testadas e removidas, porque formavam xadrez com o cinza das células de preencher.

## Correções encontradas na auditoria

1. **Linha apagada na Lanc_HORAS, perto da 1045:** o contador AS/AU (`=$AS{r-1}+$AR{r}`) virou #REF! em 5.656 células. Por isso o "detalhe dia a dia" do relatório saía vazio; em setembro só apareciam as 3 HE.
   - Novo contador, à prova de exclusão de linha: `=IF($AR{r}=1,COUNTIF($AR$2:$AR{r},1),0)`.
   - O mesmo contador foi aplicado em Lanc_HE AG e Compensações O.
   - Setembro voltou a ter 146 linhas no detalhe.
2. **Painel, linhas 33–36:** o SUMPRODUCT com N(intervalo) dava sempre 0 no Excel, porque o N() lê só o primeiro elemento. As fórmulas foram reescritas com IFERROR(V*1,0) e SUMIFS.
   - Setembro, Elétrica: 155:21 integral + 1:00 de doação = 156:21, o que bate com a Gestão; 12:54 descontadas.

## Validação

- Recálculo sem erros.
- Os valores de Gestão, Análises, SLA e Compensações são idênticos aos calculados pelo Excel do Lucas.
- O Painel difere só nas linhas corrigidas.

**Regra:** nenhum lançamento pode ser feito na planilha antiga entre o envio do arquivo e a volta.
