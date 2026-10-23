# EI32 v44 — estrutura e usabilidade (14/10/2026)

Arquivo: `EI32_44_estrutura_e_usabilidade.xlsx`, feito sobre a versão do Lucas de 14/10 (com outubro lançado).

## Base

A versão de 14/10 tinha sido feita a partir da de 07/10, sem as correções da v43. Todas foram reaplicadas:

- regra de horas × valor;
- os 44 atrasos de setembro;
- Bruno 28/09;
- Rodrigo 28/09;
- Gabriel 18/09;
- coluna O do Guilherme;
- HE: 4 linhas acertadas pelo ponto, duplicata do Duilio removida e "REMANEJADO" corrigido.

Além disso, a requisição do José em 29/08 passou para TÉC. TELECOM. Com isso, a HE de agosto paga por EI caiu mais R$ 7,34.

## Nova estrutura

| Aba | Conteúdo |
|---|---|
| Como_Usar | Guia de uma página: onde fica cada coisa, rotina do dia, fechamento e regras que geram dúvida |
| Cadastro | Antigo Apoio A:K (uma linha por pessoa) |
| Vigencias | Antigo Apoio O:V (histórico por data; adicionar linha, não editar) |
| Tabelas | Regras dos status, funções e Valor HH, parâmetros, turnos, tratamentos de HE, aprovação e contrato por data |
| Compensacoes | Quem trabalhou para compensar, horas, falta (opcional) e "Aprovado?". Só conta com "Sim" |

A aba Apoio deixou de existir. As fórmulas foram reescritas para as abas novas, e os valores ficaram idênticos em todas as abas.

## Compensação

O Resumo_Mensal Q passou a somar só as horas aprovadas na aba Compensações. A regra da equipe continua a mesma: pendências primeiro, a sobra da Elétrica cobre a Climatização e o resto abate o absenteísmo.

- As 4 compensações de agosto (Copa) foram migradas como aprovadas, e os valores não mudaram.
- Lanc_HE AE avisa "REGISTRAR na aba Compensações" quando uma HE com tratamento HORAS COMPENSADAS não tem registro na aba.
- No Painel, a linha 40 passou a ser "Compensações aguardando aprovação".

## Usabilidade

- **Lanc_HORAS, coluna AA "De onde vem":** explica a linha em texto. Mostra ferista com nome, doação de sobra (quem doou, quanto, de quanto), acordo, pendente, justificada, leve atraso, valor descontado e compensação.
- **Colunas auxiliares:** AB:AP ficam agrupadas e recolhidas.
- **Notas de cabeçalho:** colunas J, O, P, Q, V, W, Y, Z e AA, e Resumo Q.
- **Mensagens de entrada:** em Cadastro, Vigencias, Compensacoes e Lanc_HE (AB, AC, AD).

## Auditoria da v44 → v45 (15/10/2026)

Arquivo: `EI32_45_auditada.xlsx`.

**Lógica das fórmulas: OK.**

- Cada coluna tem a mesma fórmula em todas as linhas.
- O Resumo_Mensal bate com a soma das linhas em todos os meses e equipes.
- A planilha recalcula sem nenhum erro de fórmula.

**Correções aplicadas:**

| Item | O que mudou | Efeito em valor |
|---|---|---|
| Vigências duplicadas de 01/08 | Removidas as linhas antigas: Eder (TÉC. ELETRICISTA I) e José (TÉC. ELETRICISTA II) | Nenhum |
| Cadastro | Eder passou para TÉC. ELETRICISTA II e José para TÉC. TELECOM | Nenhum |
| Vigência do Alex | Criada a partir de 10/09/2026 | Nenhum |
| Bruno Domingues, 02/10 | Ausência C/ Justificativa, observação "Dia Atestado". Decisão do Lucas: o 2º atestado cobre 01 e 02/10 | Outubro Elétrica: abatimento R$ 954,96 → R$ 520,62 |

**Decisão registrada:** 09/10 foi feriado regional, transferido de 07/10 para formar 4 dias de folga seguidos. Não há lançamento nesse dia.

## Pendente

- **13/10:** Bruno Salvador e Sergio estão justificados sem observação. O Lucas vai conferir os documentos.
- **Ponto de outubro:** ainda não foi carregado; 109 linhas aparecem como "sem histórico".
- **HE de setembro sem lançamento:** Eder 05/09, Delmiro 27/09 e Eder 27/09.
- **Casos antigos em aberto:**
  - Eder, 21/09;
  - Alex, 15/09;
  - Washigton, 20/07;
  - Bruno Salvador, 10/07;
  - Delmiro e Gabriel, 17/02;
  - falta de contratação em 21/04.
- **Travas:** calendário de feriados, alerta de HE acima do ponto, "mês fechado até" e proteção das abas.
