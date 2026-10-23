# EI36 → EI32: situações e recursos que podem ser aproveitados (15/10/2026)

Arquivos analisados:

- **"Cópia de Lancamento horas - EI36 06.10":** estrutura antiga, igual à EI32 antes da v44.
- **"Lancamento horas - EI36 v2":** a versão nova feita pelo gestor da EI36.

## O que é igual

- **Status:** a lista dos 19 é a mesma nas duas planilhas, e a regra "conta no absenteísmo?" também.
- **Equivalentes que a EI32 já tem:**

| Recurso | Na EI36 v2 | Na EI32 |
|---|---|---|
| Posto vago | função na coluna H | coluna R |
| Turno diferente no dia | coluna I | coluna S |
| Situação atípica | J/K/L (Presente/Saída) | coluna L + acordo |
| Perdoado por acordo | coluna M | coluna P |
| Ferista automático | sim, maior falta primeiro | sim, mais a doação de sobra |
| Leve atraso | 10–30 min | 10–30 min |
| Regras de HE | sábado 4h a 100%, almoço de 1h acima de 6h | mesmas regras |

- **Alertas que as duas já têm:** duplicado, fora do cadastro, atraso acima de 30 min a classificar, leve atraso lançado como Horas Trabalhadas, sem ponto, ferista insuficiente, horas manuais maiores que a jornada e desconto sem observação.

## O que a EI36 v2 tem e a EI32 não

1. **Compensação vinculada à falta.** Na aba Compensações, cada registro diz de quem é a falta e a data. O abatimento cai na linha da falta: horas brutas − ferista − compensado − perdoado = descontado.
   - O tipo é automático: Autocompensação, Colega cobriu colega ou Ferista cobriu colega.
   - A coluna Situação mostra um destes: PREENCHER, CONFERIR (não há lançamento ou não havia desconto), PENDENTE, ATENÇÃO (aprovado acima da falta) ou OK.
   - A EI32 abate em bloco: primeiro as pendências, depois 1:1 sobre o absenteísmo da equipe.
2. **Alerta VINCULAR:** "Horas Compensadas" sem registro na aba Compensações. Alerta APROVAR: compensação pendente aparecendo na linha da falta.
3. **Alerta CADASTRO:** pessoa sem vigência na data do lançamento.
4. **Cadastro:** colunas "Último lançamento" e "Dias sem lançar". No Painel: "Ativo sem lançamento há mais de 5 dias".
5. **"Mês fechado até":** as linhas fechadas ficam cinza para evitar edição.
6. **Tabela de status completa:**
   - modo das horas de ausência: Dia inteiro / Pelo ponto / Leve atraso / Não;
   - se conta como trabalhado;
   - categoria;
   - texto "o que significa".
7. **Painel:** coluna "% do total" do abatimento e a checagem "Confere com o Resumo? ✔/✘".
8. **HE:** coluna "Mesma função?" (função solicitada × função do colaborador) e alerta "PONTO: aguardando histórico".
9. **Por colaborador:** dias lançados, leves atrasos descontados, horas compensadas para outros e HE em horas e em R$.
10. **SLA:** desconto manual em R$ no Resumo, mais uma tabela de indicadores (atendimentos, backlog, quadro por função × meta).
    - **Decisão do Lucas:** a EI32 não precisa, porque a demanda é sempre atendida. Manter o campo (Analises AG = 0) só para padronizar com o .bas comum.

**Checagens da EI36 antiga que a EI32 não tem:**

- HE acima de 40h no mês por pessoa;
- abatimento fora da curva (mais de 1,5 vez a média);
- posto vago nos últimos 7 dias;
- ativo sem vaga no relatório.

## Situações reais que aparecem nas observações da EI36 e não têm categoria própria

| Situação | Ocorrências | Como entra hoje |
|---|---|---|
| Jogo da seleção | ~25 | Saída antecipada ou atípica |
| "Adicional não solicitado" | ~18 | Pessoa a mais, não pedida |
| Integração TK / início de atividades | ~10 | — |
| Atestado médico, dentista ou declaração | ~60 | Justificada |
| Acidente ("torceu o pé – 7 dias") | 5 | Justificada |
| Falta de uniforme | 4 | Falta de EPI |
| Recusa a fazer a atividade | 2 | — |
| Pedido de demissão / efetivação | 2 | — |

## Uso real na EI36 v2

- **Lançamentos:** 1.707 linhas.
- **Status mais usados:** Horas Trabalhadas 1.131, falta S/J 157, férias 125, atraso S/J 114, falta C/J 108, Horas Compensadas 26.
- **Compensações:**
  - 12 de 19 estão em "PREENCHER: de quem é a falta";
  - 2 foram aprovadas acima da falta.
- **HE:** 9 de 13 estão sem tratamento.
- **SLA descontado:**
  - jul: R$ 4.466,14
  - ago: R$ 10.412,60
  - set: R$ 11.860,96
