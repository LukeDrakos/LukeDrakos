# Controle de Horas TK: padronização EI31 / EI32 / EI36

Documento de passagem para o Claude Code. Estado em 23/10/2026: **EI32 v55** + **EI36 v3** (layout do Felipe com a lógica da EI32) + **macro única** para as duas.
Autor do pedido: Lucas (estagiário de Engenharia Elétrica na thyssenkrupp, área de Facilities, contratos EI3x).
Leia este arquivo inteiro antes de mexer em qualquer planilha. O detalhe de cada coluna está em `docs/MAPA_COLUNAS.md`.

---

## 0. Regras de trabalho (obrigatórias)

1. **Rigor profissional e confirmação com o Lucas.** Pergunte antes de decidir qualquer regra de negócio que mexa em dinheiro. Não adivinhe.
2. **Meses fechados e já apresentados ao cliente não mudam sem autorização explícita.** Toda mudança precisa vir com o **efeito financeiro mês a mês**: abatimento antes, abatimento depois e diferença, por equipe.
3. **Nada é entregue sem recálculo com 0 erros.** Confira as identidades da seção 5.4 e se Relatório, Gestão e Análises batem entre si.
4. **Correções valem para todas as linhas e todas as abas afetadas**, nunca em uma linha só.
5. **Não altere configurações de segurança do Windows ou do Excel** (Central de Confiabilidade, desbloqueio de arquivo, "Confiar no acesso ao modelo de objeto do VBA"). Quando for preciso, peça ao Lucas, passo a passo.
6. **Nunca altere o histórico de pontos (coluna Y) para refletir uma suposição.** Isso seria adulteração. Ajuste vai em coluna manual (M, N, P…) com motivo na Observação (X).
7. **Trabalhe sempre numa cópia.** Ninguém pode lançar na planilha oficial enquanto você edita a cópia, senão os lançamentos se perdem. Antes de começar, confirme com o Lucas qual é o arquivo mais recente.
8. **Prontidão para reunião é requisito crítico.** Valor errado exibido diante do cliente é bug grave. Trava de mês, por exemplo, já foi abandonada porque zerou valores numa reunião.
9. **Comunicação.** Responda em português, sem jargão de programação desnecessário. O Lucas usa muito voz-para-texto, então leia o pedido com cuidado. Ele audita os arquivos sozinho e manda prints; formato de data, largura de coluna e mesclagem importam.
10. **Arquivos de entrega:** a planilha fica `.xlsx` (sem macro dentro). A macro fica só no suplemento `.xlam`.

---

## 1. Contexto do negócio

- A thyssenkrupp contrata uma empresa terceira por **posto fixo**. Quando falta alguém (falta, atraso, férias não coberta, posto vago…), a TK **abate** do valor mensal do contrato as horas não entregues × valor-hora da função (**absenteísmo**).
- **Contratos e equipes:**

| Planilha | Equipe(s) | Situação |
|---|---|---|
| **EI32** | Elétrica (Elétrica predial) e Climatização (inclui Refrigeração) | **Molde**: estrutura v55 definitiva |
| **EI36** | Serralheria (1 equipe; ambiente manual, sem ponto eletrônico confiável; tem SPOT) | Gestor: **Felipe**. **v3 (23/10)**: layout dele mantido, lógica + cores da EI32, macro funcionando (ver seção 11) |
| **EI31** | Civil / Manutenção predial e jardinagem | Despriorizada até agora; também vai para o padrão |

- **Objetivo da padronização:** as 3 planilhas com **a mesma estrutura de abas e colunas**, controladas por **um único suplemento** (`Controle de Horas TK.xlam`), com o mesmo visual (paleta thyssenkrupp) e o mesmo relatório PDF. O Lucas quer entregar à EI36 um modelo que o gestor consiga tocar sozinho.
- **Metas SLA** (aba SLA), em atendimentos / backlog:
  - EI36 Serralheria: 276 / 321
  - EI32 Elétrica predial: 169 / 178
  - EI32 Refrigeração: 119 / 10,5
  - EI32 Climatização: 119 / 10,5
  - EI31 Civil: 250 / 240

---

## 2. Onde as coisas estão

### Pasta oficial (computador do Lucas, drive G: sincronizado do SharePoint)

`G:\…\01 - Padronização Definitiva\Controle de Horas TK - pacote v53\`

- `Controle de Horas TK.xlam`: o suplemento já instalado e funcionando, com o código v53, que é o mesmo do `suplemento/Modulo_Controle_Horas.bas` deste pacote.
- `Lancamento horas - EI32.xlsx`: a planilha. Pode estar na v53 com botões ou já na v55; confirme com o Lucas.
- `Pontos 2026.xlsx`: base do relógio de ponto, aba **"Base unificada"**.
- `Relatorios\01 - EI31`, `Relatorios\02 - EI32` e `Relatorios\03 - EI36`: um PDF por mês, em `Relatorios\02 - EI32\2026\09 - Setembro - EI32.pdf`.
- **Segurança:** a pasta é Local Confiável com "Permitir Locais Confiáveis na minha rede". Foi o Lucas quem configurou.

### Este pacote

| Pasta | Conteúdo |
|---|---|
| `CLAUDE.md` | este documento |
| `docs/MAPA_COLUNAS.md` | todas as colunas de Lanc_HORAS, Lanc_HE, Compensacoes, Cadastro, Vigencias e Analises (entrada × fórmula, com exemplo), Gestão e Painel célula a célula, parâmetros, status, rótulos e nomes definidos |
| `docs/historico/` | documentos de cada versão (revisão 08/10, v44 → v55, análise EI36) com decisões e efeitos financeiros |
| `planilhas/Lancamento horas - EI32.xlsx` | **v55**, versão de referência atual (sem botões: rodar "Instalar botões") |
| `planilhas/Pontos 2026.xlsx` | base de pontos usada até agora |
| `planilhas/referencia/Lancamento horas - EI36 v2 (Felipe, 15-10).xlsx` | EI36 v2 feita pelo Felipe (cópia recebida em 15/10; peça a atual) |
| `planilhas/referencia/Lancamento horas - EI36 antiga (06-10).xlsx` | EI36 na estrutura antiga (Apoio), com os dados até 06/10 |
| `suplemento/` | `.xlam` v53, `Modulo_Controle_Horas.bas`, `INSTALAR.vbs`, `LEIA_PRIMEIRO.txt` |
| `ferramentas/` | scripts Python e Node usados para gerar as versões (seção 8) |

---

## 3. Estrutura das abas (padrão EI32 v55)

Cor das guias: azul TK = painéis; marinho = lançamentos; cinza = cadastros; azul claro = guia.

| Aba | Papel | Quem preenche |
|---|---|---|
| Como_Usar | Guia de uma página (onde fica cada coisa, rotina, fechamento) | — |
| IT_Instrucao_Trabalho, Legenda | Ocultas. Instrução de trabalho e legenda das regras (inclui a seção "Abatimento de Valor/Função") | — |
| **Gestao_Contrato** | Painel principal. **B3 = mês filtrado** (lista `lstMeses`): tudo segue essa célula. Menus nas linhas 2 e 4, D3 = link SLA, botões em G3 e H3, contador de divergências em H4 | só B3 |
| Painel_De_Controle | Detalhe do mês por equipe (B:C Elétrica, E:F Climatização), detalhe por situação (H:K) e painel anual de verificação (M:AA) | — |
| Relatorio | Fonte do PDF. A coluna N = 1 oculta a linha no PDF. Detalhe dia a dia: 150 ocorrências por equipe, 60 HE, 30 compensações | — |
| **Analises** | **Motor mensal** (linhas 3–14 Elétrica, 19–30 Climatização, 15/31 totais) + painéis (linha 33 em diante: colaboradores L37, funções L72, ferista L92/M94:S96) | — |
| SLA | Desconto SLA manual em R$ por mês e equipe, indicadores e metas | gestor |
| **Lanc_HORAS** | Um lançamento por pessoa por dia (até a linha 3873) | diário |
| **Lanc_HE** | Horas extras (até a linha 693) | diário |
| Compensacoes | Quem trabalhou para compensar, horas e "Aprovado?" (só conta com Sim), linhas 4–503 | gestor |
| Conferencia_Ponto | Recriada pela macro AtualizarPontos: lista as divergências | macro |
| Cadastro | Uma linha por pessoa (RG, nome, estilo do contrato, equipe, função, turno, valores, status) | gestor |
| Vigencias | Histórico por data (função, valor HH, turno e estilo a partir de uma data). **Mudança = linha nova, nunca editar a antiga** | gestor |
| Tabelas | Regras dos status (A:C), funções e valor HH (E:H), parâmetros (J:L), turnos (N:P), listas (R:S), contrato por data (U:X), aprovação (Z), rótulos do relatório (AB:AC), calendário PCAM 2026 (AE:AG) | gestor |
| Painel_Aux | **veryHidden**. Auxiliares | — |

- **Estilo do Contrato** (Cadastro C / Vigências H): `Normal / Fixo`, `Ferista`, `SPOT`, `Compensação` (pessoas-fantasma "Colaborador … Compensação", usadas só para horas compensadas) e `Desligado`.
  - O estilo é independente do status de emprego (Cadastro J: Ativo / Desligado).
  - Um ferista desligado continua `Ferista` em C.
- **Posto vago:** Funcionário vazio ou "Falta de contratação", com a função em R. Gera desconto da jornada inteira.

---

## 4. Regras de negócio vigentes (todas decididas pelo Lucas)

### 4.1 Jornada (Lanc_HORAS AD)

- **ADM:** 9h de segunda a quinta e 8h na sexta (Tabelas K11); sábado e domingo 0.
- **TURNO 1:** dias úteis. **TURNO 2 e TURNO 3:** segunda a sábado. Horas em Tabelas P3:P6.
- **Coluna S (Alteração de Turno):** se preenchida, prevalece.
- **Turno 2 automático:** quem é ADM, entrou antes das 07:00 e saiu até as 14:00 faz turno 2 naquele dia. Se entrou cedo e saiu depois das 14:00, continua ADM e gera alerta, porque pode ser compensação.
- O ponto (Y) e a jornada já vêm líquidos do almoço de 30 min nos dias úteis.

### 4.2 Status (Tabelas A:C) e efeito

- **Descontam (jornada cheia):**
  - Ausências S/ Justificativa
  - Férias (se não coberta)
  - Falta de Contratação
  - Falta de EPI
  - Recesso
  - Remanejamento (se o ferista não cobrir)
  - Recusa de Atividade S/ Motivo Legítimo
  - Desligamento / Pedido de Demissão
- **Não descontam (justificadas):**
  - Ausências C/ Justificativa (atestado: Observação só "Dia Atestado", sem o motivo médico)
  - Entrada Atrasada ou Saída Antecipada C/ Justificativa
  - Treinamento
  - Recusa C/ Motivo Legítimo
  - Integração / Início de Atividades
  - Afastamento por Acidente
- **Atraso ou saída:**
  - "Entrada Atrasada ou Saída Antecipada S/ Justificativa" desconta pelo ponto (jornada − horas do ponto).
  - Sem ponto válido, desconta a jornada cheia.
- **Leve atraso (10 a 30 min):** desconta. O status correto é **"Entrada Atrasada… S/ Justificativa"**, não "Horas Trabalhadas"; a planilha avisa ("AJUSTAR STATUS").
- **Mais de 30 min com "Horas Trabalhadas":** alerta "CLASSIFICAR".
- **Adicional Não Solicitado:** pessoa a mais, não pedida. As horas trabalhadas abatem o absenteísmo do mês da equipe 1:1, como compensação. **Não é ferista**; não cobre falta do dia.
- **Motivo obrigatório na Observação (X):**
  - justificadas
  - remanejamento
  - situação atípica
  - acordo
  - ajuste manual
  - treinamento
  - recusas
  - afastamento
  - desligamento

  O relatório mostra "⚠ SEM MOTIVO".

### 4.3 Ferista (motor das colunas AC, AF, AG, AH, AP, AO)

- O ferista **nunca gera absenteísmo próprio** nas métricas de abatimento.
- **Quem ele cobre:** só **técnicos** da mesma equipe com estilo `Normal / Fixo`.
  - **Nunca o encarregado**, nem na cobertura manual (coluna O). Há uma trava em AC.
- **Capacidade do dia** = horas que o ferista trabalhou no dia (AH), exceto quando ele próprio está em Remanejamento.
- **Ordem de cobertura (AP/AG):** férias primeiro (+1000), depois a maior falta.
- **Cascata (regra definitiva, confirmada pelo Lucas após mudar de ideia 2 vezes):** cobre integralmente na ordem enquanto houver capacidade. A sobra é **doada** à próxima falta, como cobertura parcial.
  - Sobra abaixo da **doação mínima de 1h** (Tabelas K15) não é doada.
  - A doação sempre aparece explicada (AA e AI).
- **Cobertura manual (O)** prevalece sobre a automática.
- **Abatimento de Valor/Função (v55, compliance, auditoria a caminho):**
  - Horas cobertas **não geram absenteísmo**.
  - Se o valor-hora do técnico coberto for **maior** que o do ferista, desconta `(HH coberto − HH ferista) × horas cobertas` (proporcional na doação parcial).
  - Ferista de cargo igual ou maior: **zero, nunca crédito**.
  - Vale desde **01/01/2026** (Tabelas K18).
  - Colunas Lanc_HORAS AY:BB; Analises BH (R$) e BI (horas).
  - Na Gestão, linha 21 **só com o valor**, mais o quadro "FERISTA × CARGO COBERTO — RELAÇÃO DO VALOR-HORA" (linhas 40–48).
  - O Lucas pediu explicitamente para **não mostrar relação de horas** ali, porque confunde com compensação.
- **Rodrigo Chanchencow Soares** (RG 544631), ferista da Elétrica desde 27/07/2026, é **TÉC. ELETRICISTA I** (cadastro corrigido na v55).
- Alex Rodrigues Santos Souza (445979) é o ferista da Climatização desde 10/09/2026, TÉC. CLIMATIZAÇÃO I. Guilherme (457855) é ferista desligado.

### 4.4 Acordo, pendência, compensação e SLA

- **P (Horas Acordadas):** perdoadas por acordo comercial; não descontam.
- **Q (Horas Pendentes):** dívida, quitada depois por compensação aprovada.
- **L/M (Situação Atípica):** horas não acordadas descontam (AL = pendente). Exemplos: jogos da seleção em 24/06 e 29/06, que sempre abatem.
- **Ordem da compensação no mês (Analises AW/BD):**
  1. quita as pendências;
  2. a sobra da Elétrica cobre a Climatização;
  3. o resto abate o absenteísmo 1:1.
  - Só conta compensação com "Aprovado? = Sim".
  - A EI36 v2 vincula a compensação à falta específica; isso ainda não está na EI32 (ver seção 9).
- **SLA:** desconto manual em R$ (Analises AG = SLA!C6) que sai do valor final. Na EI32 é normalmente 0, mas a estrutura existe para padronizar com a EI36.

### 4.5 Hora extra (Lanc_HE)

- Só se trata HE na Lanc_HE. O que estiver na Lanc_HORAS em dia não útil é troca de dia combinada.
- **Sábado:** as primeiras 4h a 100% (K5), o resto a 50%. Domingo e feriado: 100%. Dia útil: 50%.
- **Desconto de almoço de 1h em sábado, domingo e feriado** acima do limite em Tabelas K12.
  - ⚠ **Confirmar com o Lucas:** hoje K12 = 6h, mas em 30/07 ele tinha definido 4h.
- **Tratamento (V):**
  - PAGAS POR EI: valor a mais, fora do contrato; entra no saldo.
  - PAGAS PELA ÁREA
  - HORAS COMPENSADAS: registrar na aba Compensações.
  - REMANEJADO P/ OUTRA ÁREA
  - Sem tratamento = pendência.
- **Valor:** pela função solicitada (C) quando houver; senão pela vigência da pessoa na data.
- **Serviço não solicitado** não é pago; entra só como histórico.

### 4.6 Pontos (macro AtualizarPontos)

- **Y diferente do relógio:** substitui pelo relógio e marca a divergência em AV (vermelho até "Conferido? = Sim" em AW). Na Lanc_HE, AH/AI.
- **Horas Trabalhadas sem marcação:** Y = "SEM REGISTRO DE PONTO = 00:00".
- **Ausência com presença no relógio:** só marca, não muda o status.
- **Dia útil com ponto e sem lançamento:** cria a linha. **Fim de semana ou folga com ponto:** cria HE.
- **Ponto incompleto** (só entrada ou só saída): é erro do relógio. Vira só INFO e **nunca** gera desconto automático.

---

## 5. Motor de cálculo

### 5.1 Uma linha da Lanc_HORAS, na ordem

1. **Identificação:**
   - F (RG), C (estilo), D (equipe por palavra-chave), G (função), H (turno), I (valor HH) vêm do Cadastro.
   - Se houver vigência na data (AN), os valores vêm de Vigências.
2. **AD:** jornada do dia.
3. **V:** horas a subtrair.
4. **AE / AK:** conta no absenteísmo? (Tabelas C, com exceções).
5. **Ferista:**
   - AF: chave dia|equipe, só técnico Normal/Fixo que conta.
   - AP: prioridade.
   - AG: ordem.
   - AH: capacidade.
   - **AC: horas cobertas.**
   - AO: tipo de cobertura.
6. **AJ:** horas perdoadas por acordo.
7. **AM:** carga líquida = MAX(V − AC (se conta) + M − AJ − Q, 0); ferista = 0.
8. **W:** valor do absenteísmo = carga × 24 × I.
9. **BA:** abatimento de valor/função. **BB:** texto do cálculo.
10. **Textos:**
    - Z: alerta.
    - AA: "de onde vem".
    - AI: qualidade de dado.
    - AQ: situação no relatório (rótulos em Tabelas AB:AC).
    - AX: sem motivo.
11. **Relatório:** AR/AT marcam a linha do mês por equipe; AS/AU são o contador. O contador usa `COUNTIF($AR$2:$ARr,1)`, à prova de exclusão de linha.

### 5.2 Analises (por mês × equipe)

| Coluna | O que é |
|---|---|
| AU | horas brutas = Σ AM |
| AV | valor bruto = Σ W |
| AW | compensação usada |
| BG | adicional usado |
| B | horas de absenteísmo = AU − AW − BG |
| AT | valor compensado = AV × (AW + BG) / AU |
| BH | abatimento de valor/função |
| **I** | **abatimento = MAX(AV − AT, 0) + BH** |
| Z | contrato do mês (Tabelas U:W por data) |
| AG | SLA |
| **AH** | **valor final = Z − I − AG** |
| D | sem justificativa (residual: B − AS − AM, ou seja, horas − férias não cobertas − leves atrasos) |

### 5.3 Gestão

Usa INDEX/MATCH nas Analises pelo mês em B3.

| Linhas | Conteúdo |
|---|---|
| 11–19 | detalhe por status (Σ W) |
| 20 | (−) compensadas = AT |
| 21 | (+) valor/função = BH |
| 22 | TOTAL = Σ11:19 − 20 + 21 = I |

### 5.4 Identidades que precisam bater em toda entrega

- Gestão D22 = Analises I = Relatório "ABATIMENTO DO MÊS".
- Relatório: TOTAL por tipo − compensações + valor/função = abatimento.
- Painel "Horas subtraídas" = Σ AM.
- Soma das linhas da Lanc_HORAS por mês e equipe = Analises.
- O TOTAL de ocorrências do relatório = Σ AR (a linha "Outros (status sem rótulo)" deve dar 0).

**Valores de referência v55** (abatimento por mês, Elétrica / Climatização):

| Mês | Elétrica | Climatização |
|---|---|---|
| jan | 18.529,80 | 10.130,72 |
| fev | 13.445,36 | 2.895,03 |
| mar | 10.949,05 | 728,47 |
| abr | 10.643,83 | 221,40 |
| mai | 10.411,88 | 272,12 |
| jun | 1.308,57 | 1.348,12 |
| jul | 3.767,84 | 3.344,16 |
| ago | 2.893,75 | 2.028,05 |
| set | 1.321,37 | 414,05 |
| out (parcial) | 2.758,87 | 0 |

Contrato mensal: Elétrica R$ 86.991,64 e Climatização R$ 60.618,57 (a partir de ago). Tabelas K7/K8 e U:W têm o histórico.

---

## 6. Suplemento / macro (`suplemento/Modulo_Controle_Horas.bas`, módulo `modEI32`)

- **Funcionamento geral:**
  - Age sempre sobre a **planilha ativa** (`wbAlvo` = ActiveWorkbook).
  - `Alvo()` exige as abas Gestao_Contrato, Lanc_HORAS, Lanc_HE, Tabelas e Cadastro.
  - `CodigoEquipe()` tira "EIxx" do nome do arquivo.
  - `Auto_Open` → `CriarMenu`: menu "Controle de Horas TK" na guia Suplementos.
  - Os botões chamam `'Controle de Horas TK.xlam'!Macro` (`NomeMacro`).
- **Procedimentos públicos:**

| Procedimento | O que faz |
|---|---|
| AtualizarPontos | Lê "Base unificada" do arquivo em `cfgArquivoPontos` (Tabelas K16); grava `cfgUltimaAtualizacao` (K17) |
| EscolherArquivoPontos | Troca o arquivo de pontos |
| CriarRelatorio | Oculta as linhas com N = 1 e gera o PDF |
| InstalarBotoes | Gestão G3 e H3, Relatorio I1:J1, Conferencia G1:H1 e I1:J1; chama ConsertarLinks e LimparBotoesAntigos |
| ConsertarLinks | Corrige os links internos |
| PastaLocal, PastaRelatorios | Convertem URL do OneDrive/SharePoint em pasta local (registro ou ambiente) |

- **CriarRelatorio:** `NomeArquivoRelatorio` grava `Relatorios\<nn - EIxx>\<ano>\<mm - Mês> - EIxx.pdf`. O mesmo mês sobrescreve; se o PDF estiver aberto, avisa e não gera.
- **Textos com acento no VBA** usam `Tx("N{a~}o")`, porque o .bas é ASCII com CRLF.
- **Atualizar o .xlam:**
  - **Preferencial, no Windows:** Excel aberto, VBE, remover modEI32, importar o .bas e salvar o .xlam. Pode ser automatizado por COM (pywin32), mas exige "Confiar no acesso ao modelo de objeto do projeto VBA". **Quem liga isso é o Lucas.**
  - **Alternativa sem Excel** (usada na v53): `ferramentas/xlam/`.
    - `ovba.py` comprime no formato MS-OVBA.
    - Os streams dir, modEI32 e _VBA_PROJECT são reescritos: MODULEOFFSET = 0, _VBA_PROJECT = `CC 61 FF FF 00 00 00`, `__SRP_*` removidos.
    - `node mk.js` regrava o CFB.
    - Validar com `olevba`.
- **Não usar `base`/`BASE` como identificador**: é palavra reservada no LibreOffice Basic.

---

## 7. Armadilhas técnicas já vividas (não repita)

### openpyxl

- **Apaga formas e botões** e a formatação condicional em extensão x14.
  - Depois de salvar por openpyxl, é preciso rodar "Instalar botões nesta planilha".
  - Formatação condicional importante (Gestão H4) tem que ser recriada.
- **Não insere linha ajustando fórmulas:** desloque manualmente com regex (`ferramentas/fer55.py`, `shift_local`).
- **Fórmulas novas precisam de `_xlfn.`** (MAXIFS, IFS…), senão dá #NAME? no Excel. Prefira SUMIFS, COUNTIFS e INDEX/MATCH.

### Fórmulas

- **Nunca** coloque `N(intervalo)`, `IF` ou `IFERROR` dentro de SUMPRODUCT gravado por biblioteca externa. `N()` lê só o 1º elemento, e IF vira #VALUE!. Use SUMIFS ou `IFERROR(V*1,0)` testado no Excel.
- `LOOKUP(2,1/((cond1)*(cond2)),resultado)` funciona no Excel sem CSE e é o padrão para "último registro de vigência até a data".
- **Faixas fixas e capacidade:**

| Aba | Faixa |
|---|---|
| Lanc_HORAS | `$2:$3873` |
| Lanc_HE | `$2:$693` |
| Compensacoes | `$4:$503` |
| Cadastro | `$3:$150` |
| Vigencias | `$3:$149` |
| Relatório | 150 ocorrências por equipe |

  - Em SUMIFS, **todas as faixas do mesmo tamanho**; faixa diferente vira #VALUE!.
  - Ao crescer, atualize todas as faixas.
- **Contador acumulado `=$AS{r-1}+…` quebra com exclusão de linha** (já gerou #REF! em 5.656 células e esvaziou o relatório). Use `COUNTIF($X$2:$Xr,1)`.
- **Texto começando com "="** vira fórmula. Exemplo: "= ABATIMENTO" deu #VALUE!.
- **Locale:**
  - O Excel do Lucas é pt-BR (separador `;`, vírgula decimal). Fórmulas gravadas via XML são em inglês, com `,`.
  - Para número em texto, use `FIXED(x,2)`, que respeita o locale. **Não use** `TEXT(x,"0,00")`.
  - `TEXT(h,"[h]:mm")` é seguro.
- **Escala de horas:** valores de hora no Excel são fração de dia. Não divida por 24 duas vezes (já houve um bug assim em HE). Colunas em "escala de horas" (×24) usam o formato `0.00`, não `[h]:mm`.
- **Equipe por palavra-chave** (D): faltar palavra (SERRALH, ENCARREGADO, REFRIG) já escondeu muitas linhas dos relatórios. Ao incluir equipe ou função nova, audite D.
- **Nomes de equipe estão fixos** ("Elétrica", "Climatização") em centenas de SUMIFS de Analises, Gestão, Painel e Relatório. Os blocos são copiados por equipe: Analises 3–14 / 19–30, Gestão B:D / F:H, Painel B:C / E:F, Relatório em dois blocos.
- **Parâmetro de data** que filtra histórico (ex.: "início da cobertura") já zerou meses inteiros. Hoje os antigos estão desativados (K13/K14). O novo K18 vale desde 01/01/2026.
- **Painel_Aux tem que ficar veryHidden.** O LibreOffice regrava como hidden; corrija em `xl/workbook.xml`.
- **Colunas auxiliares:** a regra antiga do Lucas era "auxiliares só no Painel_Aux". Na prática o motor tem auxiliares na Lanc_HORAS (AB:AP, AR:AU, AX, AY:AZ, agrupadas e ocultas). Mantenha a organização e pergunte antes de espalhar mais.

### Windows e Excel

- Arquivo extraído de zip ganha "Mark of the Web" e dá Modo Protegido ou erro.
- Pasta de rede exige Local Confiável com "Permitir… na minha rede".
- O Lucas desbloqueia e configura isso.

### Outros

- **Vigências têm reajuste em 01/08/2026.** Exemplos: Téc. I 49,64 → 54,29; Téc. II 56,08 → 57,85. Compare valores sempre **arredondados a 2 casas** (Cadastro arredonda, Vigências não).
- **Desfazer mesclagem** antes de mover linhas e recriar depois. A Gestão e o Relatório têm muitas mesclagens.

---

## 8. Fluxo de trabalho recomendado

1. Peça ao Lucas o arquivo atual (EI32, EI36 ou EI31) e confirme que ninguém vai lançar nele até a volta.
2. Copie para uma pasta de trabalho. Guarde o original intacto para comparar.
3. Altere por **script versionado**: `ferramentas/fer55.py` e `fer55b.py` servem de modelo, com `assert` conferindo a fórmula antiga antes de trocar.
4. **Recalcule:**
   - No Windows, de preferência pelo próprio Excel via COM (`Application.CalculateFull`; também preserva botões quando a edição é feita pelo Excel).
   - Alternativa: LibreOffice headless (o script `recalc.py` da skill xlsx, se existir).
   - Contar erros (#REF!, #VALUE!, #NAME?, #DIV/0!, #N/A) = 0.
5. **Compare antes × depois** com `ferramentas/cmp.py` como base:
   - Analises I, AV, AT, AH, B por mês e equipe;
   - colunas-chave da Lanc_HORAS (AC, AM, W).
   - Tudo que mudou precisa ser explicado.
6. Ajuste veryHidden do Painel_Aux e reinstale os botões (Suplementos › Controle de Horas TK › Instalar botões nesta planilha), depois Ctrl+S.
7. Gere o PDF de um mês fechado e confira se bate com a Gestão.
8. Relate ao Lucas o que mudou, o efeito em R$ por mês e o que precisa da decisão dele.

### Ferramentas incluídas (`ferramentas/`)

Os caminhos dentro dos scripts (`/home/claude/...`, `/mnt/...`) são do ambiente antigo: troque pelos caminhos locais.
Dependências: `pip install openpyxl oletools olefile` e, para `xlam/mk.js`, `npm install cfb`.

| Arquivo | O que é |
|---|---|
| `build50.py` | v50: Analises consolidada |
| `build51.py` | v51: Relatório completo |
| `build52.py` | v52: SLA, status novos, adicional |
| `style53.py` | padrão visual |
| `fixcount.py` | contadores COUNTIF |
| `fixpainel.py` | Painel 33–36 |
| `fer55.py` / `fer55b.py` | v55: valor/função e quadro da Gestão |
| `mapa.py` | gera o MAPA_COLUNAS |
| `cmp.py` | comparação |
| `lo_check.py` | checa o módulo do .xlam no LibreOffice |
| `xlam/` | pipeline do .xlam sem Excel |

Esses scripts foram escritos para o layout da época; leia antes de reaproveitar.

---

## 9. Próximo passo: padronizar EI36 e EI31 (e alterações maiores)

### Ideias e pontos de atenção, para **discutir com o Lucas antes**

1. **Equipe parametrizada.** Hoje as equipes estão escritas dentro das fórmulas. Para 3 planilhas com 1, 2 ou mais equipes, considere uma lista de equipes em Tabelas e blocos gerados por script, ou fórmulas que leiam o nome da equipe de uma célula de cabeçalho do bloco.
   - A EI36 tem 1 equipe (Serralheria) + **SPOT** (equipe própria, contrato à parte; desconta independente de justificativa, jornada ADM 08:00–17:30).
   - A EI31 tem Civil e Jardinagem.
2. **EI36 v2 (do Felipe) tem recursos que a EI32 não tem** (detalhes em `docs/historico/05-ei36-situacoes-para-ei32.md`):
   - compensação vinculada à falta (de quem é a falta e a data), com tipo automático e situação PREENCHER / CONFERIR / PENDENTE / ATENÇÃO / OK;
   - alertas VINCULAR / APROVAR / CADASTRO (sem vigência);
   - "Último lançamento" e "Dias sem lançar" no Cadastro;
   - "Mês fechado até" (cinza);
   - tabela de status completa (modo das horas, categoria, significado);
   - "% do total" e checagem "Confere com o Resumo ✔/✘";
   - HE "Mesma função?".
3. **Regras específicas da EI36 já definidas** (jul/2026, confirmar se continuam valendo):
   - ferista cobre SERRALHEIRO I e II (não o encarregado);
   - "Horas Positivas Ferista" em vez de banco de horas;
   - saldo do ferista reseta todo mês e só cobre com saldo suficiente;
   - férias não compete pela vaga de cobertura do dia;
   - jornada própria: seg–qui 9h, sex–sáb 8h.

   Isso **difere** da cascata da EI32: decidir com o Lucas se unifica ou se vira parâmetro por planilha.
4. **Funções e valores HH** já estão em Tabelas E:H para Civil, Serralheria e outras.
5. **Faltam as metas SLA por função** (quadro).
6. **Um único .bas para as 3:** qualquer diferença de estrutura entre as planilhas quebra a macro (colunas fixas: Lanc_HORAS Y/AV/AW, Lanc_HE AB/AH/AI/P/N; abas ABA_REL e ABA_CONF; Gestão B3, G3, H3). Mantenha os mesmos nomes de aba e as mesmas letras de coluna nas 3.

### Pendências abertas (EI32)

- **Setembro/2026** subiu R$ 549,05 com o valor/função (já apresentado). O Lucas precisa alinhar o reenvio do PDF.
- **Confirmar o limite do almoço em HE de fim de semana:** 4h ou 6h (Tabelas K12 = 6h).
- **Troca de turno:** exigir motivo? 331 de 333 estão sem motivo.
- **HE paga por EI:** exigir número de requisição (Lanc_HE AD)?
- **Casos antigos a conferir:**
  - Eder 21/09
  - Alex 15/09
  - Washigton 20/07
  - Bruno Salvador 10/07
  - Delmiro e Gabriel 17/02
  - falta de contratação em 21/04
  - 13/10: Bruno Salvador e Sergio sem observação
- **Base de Pontos:**
  - sem 30–31/07;
  - em 29/07 só uma marcação;
  - faltam Eder 10/07 e Washigton 20/07.
- **Travas ainda não feitas:**
  - alerta de HE acima do ponto;
  - "mês fechado até" com proteção (proteção real foi abandonada uma vez por risco em reunião; discutir antes).

---

## 10. Histórico resumido de versões (EI32)

| Versão | Data | O que mudou |
|---|---|---|
| v44/v45 | 14–15/10 | Apoio dividido em Cadastro / Vigencias / Tabelas; aba Compensações com aprovação; coluna AA "De onde vem"; auditoria |
| v48–v50 | 15/10 | Macro de pontos, Conferencia_Ponto, menu; Analises consolidada (painéis de colaborador, função e ferista); botões sem cobrir texto |
| v51 | 15/10 | Relatório completo situação por situação, detalhe dia a dia e HE caso a caso |
| v52 | 15/10 | SLA (igual à EI36); 5 status novos + Adicional Não Solicitado; PDF por pasta de equipe; macro virou suplemento .xlam |
| v53 | 22/10 | Padrão visual único; PDF um por mês (pasta por ano); contadores COUNTIF (#REF! corrigido); Painel 33–36 corrigido; pacote com INSTALAR.vbs; .xlam gerado fora do Excel |
| v55 | 22/10 | Abatimento de Valor/Função (compliance); Rodrigo = Téc. I; trava "ferista nunca cobre encarregado"; quadro ferista × cargo na Gestão |

Detalhes, efeitos financeiros e validações de cada versão estão em `docs/historico/`.

---

## 11. EI36 v3 (23/10/2026): o que já está feito

Detalhes, efeito em R$ e pendências estão em `docs/historico/09-ei36-v3-padrao-ei32-e-macro-unica.md`.

**Decisões do Lucas:**

- o layout do Felipe fica (abas Painel, Lancamentos, Compensacoes, Lanc_HE, Resumo_Mensal, Por_Colaborador, Cadastro, Vigencias, Tabelas);
- a lógica da EI32 vale para **o ano todo**;
- a compensação é **em bloco**;
- o visual segue a paleta da EI32.

**Macro:** um só `.bas`/`.xlam` para as duas. `DefinirLayout` detecta o layout pelo nome das abas.

- Na EI36, a macro usa:
  - Lancamentos B, C, D e as colunas BC/BD (divergência / conferido);
  - Lanc_HE B, E, F, K, L e as colunas AE/AF;
  - Cadastro a partir da linha 4.
- Abas novas na EI36: Relatorio e Conferencia_Ponto.
- `cfgPontosDesde` (Tabelas B38 na EI36) limita o Atualizar Pontos e protege meses fechados.

**Script gerador:** `ferramentas/ei36/build_ei36.py`. Os `assert` conferem as fórmulas antigas antes de trocar.

- Entrada: a EI36 v2 em `planilhas/referencia/`.
- `build_bas.py` gera o `.bas`.
- `lo_macro_test.py` roda a macro no LibreOffice.

**Mapa das colunas EI36 ↔ EI32:**

| EI36 | EI32 |
|---|---|
| Y (brutas) | V |
| Z (ferista) | AC |
| AA (compensado) | parte de AT |
| AB (acordo) | AJ |
| AC / AD (descontado) | AM / W |
| BA (Valor/Função) | BA |
| Resumo_Mensal | Analises |
| Painel C3 (mês) | Gestão B3 |
