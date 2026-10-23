# EI32 v51 — relatório completo, situação por situação (15/10/2026)

O pacote é `EI32_v51_final.zip`. O módulo .bas não mudou desde a v50. A aba Relatorio foi refeita a partir da linha 9; as linhas 1 a 8 (título e Resumo do mês) continuam como estavam.

## O que o PDF mostra, na ordem

1. **Resumo do mês.** Contrato, abatimento, ocorrências, horas descontadas, ferista, justificadas, valor final e HE paga por EI.
2. **Pendências para fechar o mês.** Cada item tem a quantidade, em verde quando é zero e em vermelho quando não é:
   - ponto divergente não conferido na Lanc_HORAS (AV preenchida e AW ≠ Sim);
   - ponto divergente não conferido na Lanc_HE (AH preenchida e AI ≠ Sim);
   - HE sem tratamento;
   - compensação com "Aprovado?" diferente de Sim;
   - ocorrência sem motivo na Observação (X).
3. **Ferista do mês.** Horas trabalhadas, número de coberturas, horas cobertas, horas sem cobertura e faltas injustificadas do próprio ferista. Os dados vêm de Analises M95:S96.
4. **Ocorrências por tipo, por equipe.** São quatro grupos:
   - **Descontam:** falta S/J, atraso >30, leve atraso, posto vago, remanejamento, férias, recesso, falta de EPI, situação atípica.
   - **Não descontam:** falta justificada, atraso justificado, treinamento.
   - **Ajustes e controles:** troca de turno, acordo, pendência, ajuste manual, compensação, reposição de férias, banco de horas, horas transferidas, horas recebidas.
   - **Para conferir:** ponto divergente, e HE lançada na Lanc_HORAS.

   A linha "Outros (status sem rótulo)" garante que o TOTAL bate com o número de ocorrências.
5. **Detalhe dia a dia.** Mostra:
   - **Situação:** inclui "· turno X (cadastro Y)" e "· ponto a conferir".
   - **Ferista:** "Cobriu tudo", "Parcial: x de y (+ acordo)", "Não cobriu" (vermelho) ou "Não desconta / Não precisou (acordo)".
   - Horas descontadas, valor e explicação (coluna AA).
   - **Motivo:** vem da coluna X; quando o motivo é obrigatório e está vazio, aparece "⚠ SEM MOTIVO" em vermelho.
6. **HE por tratamento.** Para cada tratamento: horas por equipe, HE 100%, valor e o efeito no contrato.
   - **Pagas por EI:** valor a mais, fora do contrato; entra no saldo (abatimento − HE paga por EI).
   - **Pagas pela área:** não entra no valor da EI.
   - **Compensadas:** não pagas; abatem faltas ou pendências quando aprovadas.
   - **Remanejadas:** não pagas por este contrato.
   - **Sem tratamento:** pendente.
7. **HE caso a caso.** Data, funcionário, tratamento, tipo do dia, HE 50% e 100%, valor, descrição, solicitante, área, requisição e função solicitada. Quando houver "⚠ SEM TRATAMENTO" ou "⚠ PONTO A CONFERIR", a linha fica vermelha.
8. **Compensações.** "Aprovado?" fica verde quando é Sim e vermelho nos outros casos.

## Mudanças nas fórmulas

- **Lanc_HORAS AQ:** classifica também "Troca de turno" (S preenchida e diferente de H) e "Ponto divergente (a conferir)".
- **Lanc_HORAS AX (oculta):** marca 1 quando falta o motivo. O motivo é obrigatório para justificadas, remanejamento, situação atípica, acordo, ajuste manual e treinamento.
- **Tabelas AB:AC:** ganharam os rótulos "Horas Extras" e "Situação Atípica".
- **Capacidade do relatório:** 150 ocorrências por equipe, 60 HE e 30 compensações.

## Validação

- Recálculo sem erros.
- Os valores de Gestão, Analises, Painel, Lanc_HE e Compensações são idênticos aos da v50.
- **Setembro:** o TOTAL por tipo bate com o número de ocorrências (79 e 66), com as horas descontadas e com o abatimento (R$ 772,32 e R$ 414,05).
- **Maio:**
  - HE paga por EI = R$ 12.338,12, igual ao Resumo.
  - HE paga pela área = R$ 14.267,92, igual a Analises M.
- A macro AtualizarPontos continua funcionando.

## Pendente de decisão do Lucas

- Troca de turno: exigir motivo na Observação? Hoje 331 das 333 trocas no ano estão sem motivo, então ficou fora da regra de motivo obrigatório.
- HE paga por EI: exigir número de requisição (Lanc_HE AD)? Hoje o número só aparece no relatório quando está preenchido.
