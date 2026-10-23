# EI36 v3: lógica da EI32 no layout do Felipe e macro única (23/10/2026)

**Base:** "Lancamento horas - EI36 v2" do Felipe (dados até 07/10/2026). Uma cópia dela está em `planilhas/referencia/`.

**Decisões do Lucas (23/10):**

- a compensação passa a ser **em bloco**, igual à EI32;
- a lógica da EI32 vale para **o ano todo**, desde 01/01/2026;
- o layout do Felipe fica: mesmas abas, colunas e posições, só com as cores da EI32.

## O que mudou na planilha

### Lancamentos

As colunas A–AU são as mesmas. Algumas fórmulas mudaram e entraram colunas novas no fim.

| Coluna | Mudança |
|---|---|
| U | Jornada com **Turno 2 automático**: ADM que entra antes das 07:00 e sai até 14:00 faz turno 2. O auxiliar fica em BH |
| AR / AV | Ordem de cobertura do ferista: **férias primeiro** (+1000), depois a maior falta |
| AT | **Cascata da EI32**: a sobra só é doada se for de pelo menos 1h (Tabelas B36). Se a capacidade for ≥ falta − 1h, cobre a falta inteira, como na EI32 |
| AB / AA / AC | Acordo primeiro. Depois a **compensação em bloco**: AA = carga × fator do mês (Resumo_Mensal AN), igual a `AT = AV × (AW + BG) / AU` da EI32 |
| AW | Carga antes da compensação (auxiliar) |
| AX:AZ, BA, BB | **Abatimento de Valor/Função**: ferista de cargo menor que o coberto. Início em Tabelas B37 (01/01/2026) |
| BC / BD | Divergência do ponto e "Conferido?". **A macro escreve aqui** |
| BE / BF | Linha do relatório do mês |
| BG | Valor compensado (auxiliar) |

- **AG (de onde vem):** passa a explicar a compensação em bloco, a doação do ferista, o Turno 2 automático e o Valor/Função.
- **AH (alerta):** ganhou "PONTO DIVERGENTE" e "TURNO". O alerta "VINCULAR" agora pede só o registro na aba Compensações.

### Demais abas

- **Compensacoes:**
  - "Compensa a falta de" e "Data da falta" ficaram opcionais.
  - Colunas novas: J = horas que contam (só com Sim), K = mês, O = equipe.
  - A Situação segue o padrão da EI32.
- **Lanc_HE:**
  - AE / AF: divergência do ponto e "Conferido?";
  - tratamentos novos: HORAS COMPENSADAS e REMANEJADO P/ OUTRA ÁREA;
  - alerta "REGISTRAR" quando a HE é compensada mas não está na aba Compensações.
- **Resumo_Mensal:**
  - colunas novas AI:AQ: absenteísmo antes da compensação, compensação aprovada, adicional, usados, fator, sobra (horas positivas), valor compensado e Valor/Função;
  - **H (abatimento) = Σ valor descontado + Valor/Função.**
- **Painel:**
  - linha 20 = (+) Valor/Função; linha 21 = TOTAL; linha 22 = "Confere";
  - menu de atalhos na linha 2;
  - botões da macro em F3:I3;
  - L19 = ponto divergente não conferido.
- **Tabelas:**
  - 6 status novos, iguais à EI32, com o rótulo do relatório na coluna U;
  - parâmetros novos:
    - B18 = arquivo de pontos;
    - B19 = última atualização;
    - B36 = doação mínima;
    - B37 = início do Valor/Função;
    - B38 = "Atualizar Pontos a partir de";
  - B7 (início das regras) passou para 01/01/2026;
  - calendário PCAM em AF:AH.
- **Abas novas:**
  - **Relatorio**: fonte do PDF. Coluna N = 1 oculta a linha, L1 = última linha, L2 = mês.
  - **Conferencia_Ponto**: preenchida pela macro.
- **Cadastro:** 10 RGs provisórios (9000xx) foram trocados pela matrícula do relógio, por nome igual. A observação (coluna D) registra a troca.
  - Ficaram com RG provisório:
    - Ricardo Vicentin Gramacho, desligado e fora da base;
    - Alex Batista dos Santos, desligado; na base aparece "ALEX BASTIDA DOS SANTOS";
    - "Murilo de Souza Modesto- Spot".
- **Visual:** paleta da EI32.
  - Guias: azul = painéis, marinho = lançamentos, cinza = cadastros.
  - Cabeçalho azul TK = você preenche; marinho = automático.
  - Textos em Arial cinza 4D4D4D; verde e vermelho-claro nas regras de cor.

## Efeito financeiro (abatimento R$, Serralheria)

| Mês | v2 (Felipe) | v3 | Diferença | Por quê |
|---|---:|---:|---:|---|
| jan–abr | — | — | 0,00 | sem efeito |
| mai | 4.860,40 | 4.146,89 | **−713,51** | ver abaixo |
| jun–ago | — | — | 0,00 | sem efeito |
| set | 26.441,78 | 22.100,84 | **−4.340,94** | ver abaixo |
| out (parcial) | 11.653,28 | 10.321,47 | **−1.331,80** | 24h aprovadas em 03/10 (3 × 8h) |

**Maio:**

- O ferista Roniel (SERRALHEIRO I) passou a cobrir:
  - Robson, em 11/05: 9h;
  - Ivanildo, em 28/05: 6:17h. O ferista trabalhou 6h; pela regra da EI32, se a sobra é ≥ falta − 1h, ele cobre a falta inteira.
- Isso gerou R$ 60,83 de Valor/Função, porque cobriu SERRALHEIRO II.
- Leve atraso desde 01/01 entrou em 11/05 (Ivanildo, 13 min) e 25/05 (Ronaldo Rosendo, 25 min).

**Setembro:**

- Compensação em bloco: 133h aprovadas no mês contra 54h na v2.
- Na v2, 12 registros aprovados não tinham "de quem é a falta" e por isso não abatiam.

**Detalhes técnicos:**

- O SLA não mudou.
- Turno 2 automático não teve efeito nos dados atuais.
- Os "Adicional não solicitado" do Diego (19 linhas em jun) continuam lançados como "Horas Trabalhadas". Se forem reclassificados para o status novo, abatem junho.

## Macro (`Modulo_Controle_Horas.bas` / `Controle de Horas TK.xlam`)

- O mesmo módulo reconhece os dois layouts sozinho, por `DefinirLayout`:
  - EI32: Gestao_Contrato / Lanc_HORAS;
  - EI36: Painel / Lancamentos.
- Colunas da EI36 usadas pela macro:
  - Lancamentos B, C, D, BC, BD;
  - Lanc_HE B, E, F, K, L, AE, AF;
  - Cadastro com dados a partir da linha 4 e equipe na coluna E.
- Novo nome opcional `cfgPontosDesde`: datas anteriores não são conferidas nem criadas.
  - Na EI36 = 01/10/2026 (Tabelas B38). Na EI32 não existe, então não há limite.
  - **Por que existe:** a EI36 tem ponto **bruto**, sem os 30 min de almoço, em mai–jun, e só exceções lançadas em jan–abr.
  - **Atualizar desde 01/01 mudaria os meses:**
    - jan +84,93
    - fev +70,77
    - mar +102,12
    - abr +78,74
    - mai +470,00
    - jun +692,60
    - jul −22,53
    - ago +116,66
    - set +182,25
  - Também seriam criadas cerca de 870 linhas de "Horas Trabalhadas" em jan–abr.
  - Baixar a data é decisão do Lucas.
- **Teste no LibreOffice (modo VBA):**
  - AtualizarPontos rodou na EI36 com o limite em 01/09: 232 históricos, 22 divergências, 1 linha criada, Conferência montada.
  - Na EI32, a macro nova deu **resultado idêntico** ao da macro antiga: 0 diferenças em Lanc_HORAS, Lanc_HE e Conferencia_Ponto.
  - Botões e PDF só podem ser testados no Excel (CalculateFull e Shapes não existem no LibreOffice).
- O `.xlam` foi gerado sem Excel, pelo pipeline `ferramentas/xlam`: stream `modEI32` novo, mesmo `dir` e `_VBA_PROJECT`. Validado com olevba.

## Validação

- Recálculo no LibreOffice: **0 erros em 189.683 fórmulas**.
- Identidades conferidas em todos os meses:
  - Resumo H = Σ AD + Σ BA das linhas;
  - Painel H21 = Resumo H = Relatório J33 ("✔ confere").
- Valor compensado de setembro: 7.385,46 = 29.486,30 × 133 / 531, a mesma fórmula da EI32.

## Pendências para o Lucas e o Felipe

1. Reinstalar os botões: Suplementos › Controle de Horas TK › Instalar botões nesta planilha › Ctrl + S.
2. Mai, set e out mudaram de valor (tabela acima). Alinhar com o cliente se setembro já foi apresentado.
3. Decidir se baixa a data de `cfgPontosDesde` (Tabelas B38).
4. Diego (jun): reclassificar como "Adicional Não Solicitado"?
5. SPOT: o Murilo tem dois cadastros (normal e "- Spot"). O ponto do relógio vai para o cadastro normal (518607).
6. Observações de atestado com motivo médico ("Atestado médico", "Declaração médica"). A regra da EI32 é só "Dia Atestado".

## Padronização visual e orientações (EI32 e EI36, 23/10)

Script: `ferramentas/ei36/padronizar.py EI32|EI36 entrada saida [logo]`. Não mexe em fórmula de cálculo: Analises I (EI32) e Resumo_Mensal H (EI36) ficaram idênticos.

- **Aba `Como_Usar` igual nas duas** (na EI36 substitui a LEIA-ME). Seções:
  - onde cada coisa fica;
  - rotina do dia;
  - **COMO LANÇAR CADA SITUAÇÃO** (situação → status → o que preencher → desconta?);
  - o que a planilha faz sozinha;
  - botões;
  - fechamento do mês;
  - mudanças de cadastro;
  - cores.
- **Dicas em todas as colunas de entrada:**
  - mensagem ao clicar na célula (validação com mensagem);
  - balão no cabeçalho (comentário).
  - Abas: Lanc_HORAS/Lancamentos, Lanc_HE, Compensacoes, Cadastro e Vigencias.
- **Painel da EI36 com o topo da Gestão da EI32:**
  - título centralizado "EI36 — GESTÃO DE CONTRATO";
  - faixa da equipe;
  - logo thyssenkrupp;
  - atalho "Como usar".
- **Nas duas:**
  - cabeçalhos auxiliares em cinza;
  - zoom 90 nos painéis e 85 nos lançamentos;
  - cabeçalho/rodapé de impressão "EIxx | thyssenkrupp | aba" e "Página x de y".
- Recálculo: 0 erros (EI32 174.851 fórmulas; EI36 189.696).
