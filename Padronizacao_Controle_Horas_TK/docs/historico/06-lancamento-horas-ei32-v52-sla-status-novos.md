# EI32 v52 — SLA, status novos, pasta de relatórios e suplemento (15/10/2026)

O pacote é `EI32_v52_final.zip`. Ele traz a planilha, o módulo .bas, o Pontos 2026, o LEIA_PRIMEIRO e o INSTALAR_SUPLEMENTO. A v52 parte da v51 e é o padrão que a EI36 e a EI31 vão seguir.

## Macro como suplemento (.xlam): instalação definitiva

**Problema relatado pelo Lucas:** depois de salvar, os botões paravam de funcionar ao reabrir a planilha ou abri-la em outro computador. A causa é que a macro ficava dentro da planilha:

- ela some quando o arquivo é salvo como .xlsx;
- o Excel bloqueia macro de arquivo aberto em outro computador ou vindo do SharePoint;
- os botões ficavam apontando para uma macro que não existia mais.

**Solução:** o módulo vira o suplemento `Controle de Horas TK.xlam`, salvo na pasta "01 - Padronização Definitiva".

- Cada computador instala o suplemento uma vez, pelo caminho Opções > Suplementos > Procurar. Quando o Excel perguntar se quer copiar para a pasta de suplementos, a resposta é **Não**, para todos usarem o mesmo arquivo. Uma atualização da macro passa a ser só a troca do .xlam.
- As planilhas EI31, EI32 e EI36 ficam .xlsx, sem macro dentro.
- O código age sempre sobre a planilha ativa (`wbAlvo` = ActiveWorkbook). A função `Alvo()` confere se ela tem as abas Gestao_Contrato, Lanc_HORAS, Lanc_HE, Tabelas e Cadastro.
- Os botões usam o nome qualificado da macro (`'Controle de Horas TK.xlam'!Macro`, pela função `NomeMacro`).
- `Auto_Open` cria o menu "Controle de Horas TK" na guia Suplementos, com Atualizar Pontos, Criar Relatório, Trocar arquivo de Pontos e Instalar botões nesta planilha.
- O título das mensagens é o código da equipe (`CodigoEquipe()`).
- O .xlam não pode ser gerado fora do Excel; o Lucas cria uma vez (passo a passo no INSTALAR_SUPLEMENTO.txt).

## Pasta dos relatórios

O PDF gerado vai para `Relatorios\<pasta com EIxx no nome>`, ao lado da planilha. Para a EI32, isso é `Relatorios\02 - EI32`.

- O código da equipe vem do nome do arquivo.
- Se a pasta não existir, a macro cria `Relatorios\EIxx`.
- **OneDrive / SharePoint:** `PastaLocal()` converte o endereço https para a pasta do computador, primeiro pelo registro do OneDrive (UrlNamespace → MountPoint) e, se não achar, pelas variáveis OneDriveCommercial e OneDrive.
- A conversão de https só pode ser testada no Excel do Lucas.

## SLA (igual EI36, manual)

- **Aba SLA:**
  - desconto em R$ por mês, separado para Elétrica e Climatização, com campo de referência;
  - indicadores com meta e realizado de janeiro a dezembro: verde quando está dentro da meta, vermelho quando está fora;
  - mais 3 linhas livres para outros indicadores.
- **Metas preenchidas** (tabela de metas do contrato):

| Equipe | Meta de atendimentos | Meta de backlog |
|---|---|---|
| Elétrica predial | 169 | 178 |
| Climatização | 119 | 10,5 |
| Refrigeração | 119 | 10,5 |

- Climatização e Refrigeração aparecem como linhas separadas dentro do bloco da Climatização.
- As metas de quadro por função continuam em branco.
- A aba guarda também a tabela de referência de todas as divisões:
  - EI36 Serralheria: 276 / 321;
  - EI31 Civil: 250 / 240.
- **Analises AG** = desconto informado. AH = Z − I − AG, ou seja, o desconto sai do valor final.
- **Gestão:** a linha 8 mostra o desconto SLA, e o menu I2 tem o atalho "SLA".
- **Relatório:** a linha 9 mostra o desconto e só aparece quando é maior que zero.
- A EI32 normalmente não desconta SLA, porque a demanda é atendida. A estrutura existe para padronizar com a EI36.

## Status novos e regras (decididos pelo Lucas)

| Status | Absenteísmo | Observação |
|---|---|---|
| Recusa de Atividade S/ Motivo Legítimo | Desconta | Motivo obrigatório |
| Recusa de Atividade C/ Motivo Legítimo | Não desconta (justificada) | Motivo obrigatório |
| Ausências Devido a Integração / Início de Atividades | Não desconta (justificada) | |
| Ausências Devido a Afastamento por Acidente | Não desconta (justificada) | Motivo obrigatório |
| Ausências Devido a Desligamento / Pedido de Demissão | Desconta | Motivo obrigatório |
| Adicional Não Solicitado | — | As horas trabalhadas abatem o absenteísmo do mês da equipe 1:1, como compensação. **Não** é ferista e não cobre falta do dia. |

**Onde as regras foram aplicadas:**

- **Lanc_HORAS:**
  - V conta o dia inteiro para os 5 status de ausência;
  - U inclui o adicional;
  - AB e AA tratam as 3 justificadas novas.
- **Analises:**
  - V e a coluna P incluem as justificadas novas;
  - BF = horas de adicional do mês;
  - BG = MIN(BF, AU − AW), as horas usadas;
  - B = AU − AW − BG;
  - AT = AV × (AW + BG) / AU.
- **Gestão:** a linha 19 traz "Outras", a linha 20 traz "(−) Compensadas + adicional" e a linha 21 traz o TOTAL.
- **Painel:** "Outras" fica nas linhas 18 e 35, e o TOTAL fica em 19 e 36.
- **Relatório:** grupos atualizados, mais as linhas "(−) Compensações + adicional" e "ABATIMENTO DO MÊS (resultado)", que batem com o Resumo.
- **AtualizarPontos:** o adicional é tratado como dia trabalhado.

## Validação

- Recálculo sem erros.
- Sem dados novos, os valores são idênticos aos da v51.
- **Teste em setembro, Elétrica:**
  - afastamento: não desconta;
  - recusa S/ motivo: desconta R$ 459,88;
  - adicional de 9h: abate R$ 531,71;
  - SLA de R$ 1.000: sai do valor final.
- O módulo refeito para suplemento rodou no LibreOffice: Alvo, PastaRelatorios, CodigoEquipe e AtualizarPontos funcionaram.
- O menu e os botões só podem ser testados no Excel.
