# EI32 — pontos automáticos, relatório e menu (v48 → v50, 15/10/2026)

**Versão atual: v50.** Pacote `EI32_v50_final.zip`, com quatro arquivos:

- `Lancamento horas - EI32.xlsx`
- `Pontos 2026.xlsx`
- `Modulo_Controle_Horas.bas` (módulo único `modEI32`; ASCII com CRLF)
- `LEIA_PRIMEIRO.txt`

## v50: aba Análises e botões (15/10)

**Abas visíveis (12):** Como_Usar, Gestao_Contrato, Painel_De_Controle, Relatorio, Analises, Lanc_HORAS, Lanc_HE, Compensacoes, Conferencia_Ponto, Cadastro, Vigencias, Tabelas. As abas Painel_Aux (veryHidden), IT e Legenda ficam ocultas.

**Aba `Analises`:** é a antiga Resumo_Mensal renomeada, e agora inclui os blocos que antes ficavam em abas separadas.

- **Linhas 1–31:** o motor de cálculo, oculto como antes. As colunas continuam as mesmas, então Gestão, Painel, Relatório e `lstMeses` só tiveram o nome da aba trocado.
- **Colunas A–J, a partir da linha 33:** os cartões do mês.
- **Coluna L em diante:**
  - Horas por Colaborador a partir de L37;
  - Horas por Função a partir de L72;
  - Cobertura do Ferista a partir de L92.
- **Linha 34:** atalhos internos, com HYPERLINK. O painel fica congelado em A36.
- **Menu da Gestão (linha 2):** Análises, Colaboradores, Funções e Ferista apontam para essas âncoras.
- As abas Horas_por_Colaborador, Horas_por_Funcao e Ferista foram removidas.

**Botões sem cobrir texto:**

- **Relatorio:** o título passou a ocupar só B1:H1; o botão fica em I1:J1.
- **Conferencia_Ponto:** o título ocupa B1:F1; os botões ficam em G1:H1 e I1:J1.
- **Gestão:** os botões continuam em D3:E3 e G3:H3.
- **InstalarBotoes:** agora chama `LimparBotoesAntigos`, que apaga as formas `btn*` e as formas ligadas às macros antes de recriar os botões.

**Auditoria da v50:**

- Recálculo sem erros (154.751 fórmulas).
- Os valores de todas as abas e dos blocos movidos são idênticos aos da v49.
- As identidades do Resumo conferem (B = AU − AW; AH = Z − I − AG).
- 128 links internos verificados, nenhum quebrado e nenhum alvo externo.
- A macro AtualizarPontos rodou na emulação: 0 novos e 12 INFO.

## Macros (módulo único)

| Macro | O que faz |
|---|---|
| AtualizarPontos | Lê a aba "Base unificada" da planilha de Pontos (arquivo do SharePoint sincronizado; o caminho fica em Tabelas › `cfgArquivoPontos`). Preenche a coluna Y e marca as divergências na coluna AV (Lanc_HORAS) ou AH (Lanc_HE). A linha fica vermelha por formatação condicional enquanto "Conferido?" (AW/AI) ≠ Sim. Recria a aba Conferencia_Ponto. |
| EscolherArquivoPontos | Troca o arquivo de Pontos. |
| CriarRelatorio | Gera o PDF do mês filtrado em Gestão B3 (aba Relatorio). |
| InstalarBotoes | Apaga os botões antigos e cria os novos na Gestão, no Relatorio e na Conferencia_Ponto. Também conserta os links internos. |

## Regras de divergência (decididas pelo Lucas)

- **Y diferente do relógio:** substitui pelo relógio e marca a linha em vermelho.
- **Horas Trabalhadas sem marcação no relógio:** Y = "SEM REGISTRO DE PONTO = 00:00" (zera as horas). A linha não é apagada e fica vermelha.
- **Ausência lançada, mas o relógio mostra presença:** linha vermelha; o status não muda.
- **Dia útil com ponto e sem lançamento:** a macro cria a linha (Horas Trabalhadas) e marca em vermelho.
- **Fim de semana ou folga com ponto e sem HE:** a macro cria a linha na Lanc_HE e marca em vermelho.
- **Calendário PCAM 2026:** fica em Tabelas AE:AG. A terceira não segue o calendário à risca, então qualquer diferença vira linha vermelha. A jornada não muda.
- **Base de pontos com falha** (dia sem nenhum registro, marcação incompleta): não altera nada; aparece só como INFO.

## Carga inicial já aplicada (base de 02/01 a 30/09)

- 86 avisos vermelhos (79 linhas) e 12 INFO.
- 3 HE de setembro criadas: Eder 05/09, Delmiro 27/09 e Eder 27/09, sem tratamento.

**Base de Pontos compartilhada:**

- Não tem 30/07 nem 31/07.
- Em 29/07 só há uma marcação por pessoa.
- Faltam alguns dias de julho (Eder 10/07, Washigton 20/07).

**Impacto em valor:** só Climatização em setembro. O Alex em 15/09 ficou sem ponto e foi zerado; o ferista deixou de cobrir 0:37 e o abatimento foi de R$ 375,17 para R$ 414,05.

## Menu

- **Gestão linha 2 (azul TK):** Relatório, Análises, Colaboradores, Funções, Painel, Ferista, Como usar.
- **Gestão linha 4 (marinho):** Lançamentos, HE, Compensações, Conferência de ponto, Cadastro, Tabelas.
- **Gestão H4:** contador de divergências.
- **Painel linha 17:** divergências de ponto não conferidas, por mês.
- **Volta à Gestão:** o título de cada aba tem o link "◄ Gestão".

## Próximo passo (depois da reunião)

Padronizar EI31 e EI36 na estrutura da EI32 para que o mesmo .bas controle as três planilhas.
