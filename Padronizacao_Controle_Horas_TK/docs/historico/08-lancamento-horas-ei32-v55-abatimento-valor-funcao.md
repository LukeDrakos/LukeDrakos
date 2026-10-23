# EI32 v55 — Abatimento de Valor/Função (22/10/2026)

## Regra (definida pelo Lucas após reunião de compliance)
- O ferista abate as horas de absenteísmo dos TÉCNICOS (nunca do encarregado — inclusive na cobertura manual, coluna O).
- Horas cobertas não geram absenteísmo.
- Se o técnico coberto tem valor-hora MAIOR que o do ferista: débito "Abatimento de Valor/Função" = (HH coberto − HH ferista) × horas cobertas. Proporcional também na doação parcial.
- Ferista de cargo igual ou maior: zero (nunca gera crédito).
- Vale desde 01/01/2026 (Tabelas!K18 "Início do Abatimento de Valor/Função").
- Cadastro corrigido: Rodrigo Chanchencow Soares (544631) = TÉC. ELETRICISTA I (Cadastro + Vigências 27/07 e 01/08). Ferista da Climatização (Alex) é Téc. Climatização I; não existe Téc. II na Climatização hoje.

## Implementação
- Lanc_HORAS: AY RG do ferista (aux, oculta), AZ HH do ferista por vigência (aux, oculta), BA Abatimento Valor/Função (R$), BB texto do cálculo. AC ganhou trava de encarregado.
- Analises: BH valor, BI horas cobertas com diferença; I (abatimento) = MAX(AV−AT,0)+BH → entra no Valor final a pagar.
- Gestão: linha 21 "( + ) Abatimento de Valor/Função (ver quadro abaixo)" SÓ COM VALOR (sem horas — Lucas: horas ali confundem com compensação); linha 20 = valor compensado (AT); TOTAL ABATIDO = soma − compensadas + valor/função (linhas 21–37 desceram 1).
- Gestão, quadro novo (linhas 40–48, uma coluna por equipe): "FERISTA × CARGO COBERTO — RELAÇÃO DO VALOR-HORA" — ferista do mês, função e valor-hora do ferista (vigência do mês), cada cargo técnico da equipe com valor-hora e diferença para o ferista em R$/h ("sem desconto" quando zero), e ABATIMENTO DE VALOR/FUNÇÃO DO MÊS. Sem relação de horas, por pedido do Lucas.
- Painel de Controle: bloco "ABATIMENTO DE VALOR/FUNÇÃO" (linhas 46–48: valor-hora do ferista no mês + valor abatido). Corrigido o rótulo E40 (estava "Valor R$", é "HORAS PENDENTES").
- Relatório: linha "( + ) Abatimento de Valor/Função" antes do ABATIMENTO DO MÊS (só valor, sem horas); FERISTA DO MÊS mostra o valor; detalhe dia a dia mostra "Acobertado pelo ferista / Abatimento de Valor/Função", o valor no Ferista/Valor e o cálculo na Explicação.
- Legenda: seção nova explicando a regra.

## Efeito (só Elétrica; Climatização = 0)
| Mês | Abatimento antes | Valor/Função | Abatimento depois |
|---|---|---|---|
| ago/2026 | 2.821,04 | 72,71 | 2.893,75 |
| set/2026 | 772,32 | 549,05 | 1.321,37 |
| out/2026 (parcial) | 2.545,27 | 213,60 | 2.758,87 |
Jan–jul: zero. 50 linhas afetadas. Diferenças (ago+): Téc. II R$ 3,56/h; Téc. Telecom R$ 2,96/h; Téc. I sem desconto. HE do Rodrigo: nenhum lançamento, sem efeito.

## Validação
Recalc 0 erros em 174.838 fórmulas; AC/W/AM idênticos à v53 (só ruído de 1e-12); Relatório e Gestão: 772,32 − 0 + 549,05 = 1.321,37. Painel_Aux veryHidden. Botões precisam ser reinstalados (Suplementos › Controle de Horas TK › Instalar botões nesta planilha).
Scripts: ferramentas/fer55.py + fer55b.py (aplicáveis sobre o arquivo atual do Lucas).
