# Mapa de colunas e linhas — EI32 v55

Gerado automaticamente a partir de `planilhas/Lancamento horas - EI32.xlsx` (v55).
"Tipo": **entrada** = preenchida à mão; **fórmula** = automática (não digitar). A fórmula mostrada é a da linha 2/3 (exemplo), truncada.

## Lanc_HORAS (cabeçalho na linha 1, 3873 linhas, estado: visible)

| Col | Cabeçalho | Tipo | Oculta | Exemplo (linha 2) |
|---|---|---|---|---|
| A | Data | entrada |  | `2026-01-05 00:00:00` |
| B | Mês | fórmula |  | `=IF($A2="","",DATE(YEAR($A2),MONTH($A2),1))` |
| C | Contrato | fórmula |  | `=IF($A2="","",IF(OR(TRIM($E2)="",TRIM($E2)="Falta de contratação"),"Posto Vago",IF($AN2>0,INDEX(Vigencias!$H$1:$H$149,$AN2),IFERROR(INDEX(Cadastro!$C$3:$C$150,MATCH(TRIM( …` |
| D | Equipe | fórmula |  | `=IF($A2="","",IF(TRIM($E2)="",IF(TRIM($R2)="","",IF(OR(ISNUMBER(SEARCH("CLIMAT",UPPER(TRIM($E2&" "&$R2)))),ISNUMBER(SEARCH("REFRIG",UPPER(TRIM($E2&" "&$R2)))),ISNUMBER(SE …` |
| E | Funcionário | entrada |  | `Bruno Domingues Pereira Barbosa` |
| F | RG | fórmula |  | `=IF(TRIM($E2)="","",IFERROR(INDEX(Cadastro!$A$3:$A$150,MATCH(TRIM($E2),Cadastro!$B$3:$B$150,0)),""))` |
| G | Função / Custo | fórmula |  | `=IF($A2="","",IF(TRIM($E2)<>"",IF(OR($C2="Ferista",NOT(ISNUMBER($F2))),IFERROR(INDEX(Cadastro!$E$3:$E$150,MATCH(TRIM($E2),Cadastro!$B$3:$B$150,0)),IF($R2<>"",$R2,"")),IF( …` |
| H | Turno | fórmula |  | `=IF($A2="","",IF($G2="","",IF(OR(TRIM($E2)="Falta de contratação",NOT(ISNUMBER($F2))),"ADM",IF($C2="Ferista",IFERROR(INDEX(Cadastro!$F$3:$F$150,MATCH(TRIM($E2),Cadastro!$ …` |
| I | Valor  HH (R$) | fórmula |  | `=IF($G2="","",IF(TRIM($E2)<>"",IF(OR($C2="Ferista",NOT(ISNUMBER($F2))),IFERROR(INDEX(Cadastro!$G$3:$G$150,MATCH(TRIM($E2),Cadastro!$B$3:$B$150,0)),IFERROR(INDEX(Tabelas!$ …` |
| J | Status | entrada |  | `Horas Trabalhadas` |
| K | Horas  Tranferidas | entrada |  | |
| L | Situção  Atípica | entrada |  | |
| M | Horas  Absenteismo  (Manual) | entrada |  | |
| N | Horas  Justificadas (Manual) | entrada |  | |
| O | Horas  Ferista (Cobertura MANUAL) | entrada |  | |
| P | Horas  Acordadas | entrada |  | |
| Q | Horas  Pendentes | entrada |  | |
| R | Falta de Contratação | entrada |  | |
| S | Alteração de Turno | entrada |  | `ADM` |
| T | Conta no  absenteísmo? | fórmula |  | `=IF($A2="","",IF(AND((IF(AND($AK2="Sim",$V2<>""),MAX(N($V2)-N($AC2),0),0)+N($M2))>0,(IF(AND($AK2="Sim",$V2<>""),MAX(N($V2)-N($AC2),0),0)+N($M2))<=N($P2)),"Não",IF(N($M2)> …` |
| U | Horas  trabalhadas | fórmula |  | `=IF($A2="","",IF(OR($J2="Horas Trabalhadas",$J2="Horas Compensadas",$J2="Banco de Horas",$J2="Remanejamento",$J2="Reposição de Férias",$J2="Reposição de Férias - HT",$J2= …` |
| V | Horas a  subtrair | fórmula |  | `=IF($A2="","",IF(AND($J2="Horas Compensadas",$K2="Horas Transferidas"),SUMIFS($O$2:$O$3873,$A$2:$A$3873,$A2,$D$2:$D$3873,$D2,$K$2:$K$3873,"Horas Recebidas"),IF(OR($J2="Au …` |
| W | Valor  absenteísmo (R$) | fórmula |  | `=IF($I2="",0,MAX((IF(AND($AK2="Sim",$V2<>""),MAX(N($V2)-N($AC2),0),0)+N($M2))-N($AJ2)-N($Q2),0)*24*$I2)` |
| X | Observação | entrada |  | |
| Y | Histórico dos pontos | entrada |  | `07:45#17:21 = 09:06` |
| Z | Alerta / Conferência | fórmula |  | `=IF($A2="","",IF(AND($J2="Horas Compensadas",$K2="Horas Transferidas"),"Horas Compensadas: COBERTURA via Remanejamento/Ferista (nao e autocompensacao)",IF(AND($J2="Horas  …` |
| AA | De onde vem (explicação da linha) | fórmula |  | `=IF($A2="","",TRIM(IF($J2="Adicional Não Solicitado","Adicional não solicitado: trabalhou "&TEXT(N($U2),"[h]:mm")&"h; essas horas abatem o absenteísmo do mês da equipe (1 …` |
| AB | Valor Justificadas (aux) | fórmula |  | `=IF($A2="","",IF($I2="",0,(IF(AND(OR($J2="Ausências C/ Justificativa",$J2="Entrada Atrasada ou Saída Antecipada C/ Justificativa",$J2="Recusa de Atividade C/ Motivo Legít …` |
| AC | Cobertura do Ferista (h) | fórmula |  | `=IF(OR($A2="",ISNUMBER(SEARCH("ENCARREGADO",$G2))),0,IF(N($O2)>0,MIN(N($O2),N($V2)),IF($AF2="",0,IF(($AH2-SUMIFS($V$2:$V$3873,$AF$2:$AF$3873,$AF2,$AG$2:$AG$3873,"<"&$AG2) …` |
| AD | Jornada do dia (aux) | fórmula |  | `=IF($A2="",0,IF(AND($S2="",UPPER(TRIM($H2))="ADM",WEEKDAY($A2,2)<=5,ISNUMBER(SEARCH(" = ",$Y2)),IFERROR(TIMEVALUE(LEFT($Y2,5)),1)<TIME(7,0,0),IFERROR(TIMEVALUE(RIGHT(LEFT …` |
| AE | Conta abs. base (aux) | fórmula |  | `=IF($A2="","",IFERROR(VLOOKUP($J2,Tabelas!$A$3:$C$150,3,FALSE()),""))` |
| AF | Chave cobertura (aux) | fórmula |  | `=IF($A2="","",IF(AND(TRUE(),$C2="Normal / Fixo",$G2<>"",ISERROR(SEARCH("ENCARREGADO",$G2)),$AE2="Sim",$K2<>"Horas Recebidas",N($V2)>0),TEXT($A2,"0")&"¦"&$D2,""))` |
| AG | Ordem cobertura (aux) | fórmula |  | `=IF($AF2="",0,COUNTIFS($AF$2:$AF$3873,$AF2,$AP$2:$AP$3873,">"&$AP2)+COUNTIFS($AF$2:$AF2,$AF2,$AP$2:$AP2,$AP2))` |
| AH | Horas Ferista no dia (aux) | fórmula |  | `=IF($A2="",0,SUMIFS($U$2:$U$3873,$A$2:$A$3873,$A2,$D$2:$D$3873,$D2,$C$2:$C$3873,"Ferista",$J$2:$J$3873,"<>Remanejamento"))` |
| AI | Alerta Qualidade de Dado | fórmula |  | `=IF(AND(N($AC2)>0,N($AC2)<N($V2)),"DOAÇÃO DE HORAS RESIDUAIS DO FERISTA sobre falta já acobertada ("&TEXT($AC2,"[h]:mm")&" de "&TEXT($V2,"[h]:mm")&")",IF($A2="","",IF(AND …` |
| AJ | Horas Perdoadas (aux) | fórmula |  | `=IF($A2="","",MIN((IF(AND($AK2="Sim",$V2<>""),MAX(N($V2)-N($AC2),0),0)+N($M2)),N($P2)))` |
| AK | Conta no absenteísmo? (aux) | fórmula |  | `=IF($A2="","",IF($C2="Ferista","Não",IF(N($AC2)>=N($V2),"Não",IF($K2="Horas Recebidas","Não",IF(AND($J2="Horas Trabalhadas",$V2<>""),"Sim",IF(OR($L2<>"",$M2<>"",$N2<>""), …` |
| AL | Situação Atípica Pendente (aux) | fórmula |  | `=IF(OR($A2="",$L2=""),"",MAX(N($M2)-N($AJ2),0))` |
| AM | Carga Líquida de Absenteísmo (aux) | fórmula |  | `=IF($A2="","",IF($C2="Ferista",0,MAX((IF(AND($AK2="Sim",$V2<>""),MAX(N($V2)-N($AC2),0),0)+N($M2))-N($AJ2)-N($Q2),0)))` |
| AN | Linha vigência (aux) | fórmula |  | `=IF($A2="","",IFERROR(LOOKUP(2,1/((Vigencias!$A$3:$A$149=$F2)*(Vigencias!$D$3:$D$149<=$A2)),ROW(Vigencias!$A$3:$A$149)),0))` |
| AO | Tipo de Cobertura do Ferista | fórmula |  | `=IF($A2="","",IF(N($AC2)=0,"",IF(N($O2)>0,"COBERTURA MANUAL do ferista",IF(N($AC2)>=N($V2),"Cobertura integral","DOAÇÃO PARCIAL — residual do ferista sobre falta já acobe …` |
| AP | Prioridade cobertura (aux) | fórmula |  | `=IF($AF2="","",IF($J2="Ausências Devido a Férias",1000,0)+MAX(N($V2)-N($P2)-N($Q2),0)*24)` |
| AQ | Situação no relatório (aux) | fórmula |  | `=IF($A2="","",IF(AND($C2="Ferista",$J2="Horas Trabalhadas"),IF(AND($AV2<>"",$AW2<>"Sim"),"Ponto divergente (a conferir)",""),IF($L2<>"","Situação atípica",IF($J2="Horas T …` |
| AR | Rel. Elétrica (aux) | fórmula |  | `=IF(AND($B2=Gestao_Contrato!$B$3,$D2="Elétrica",$AQ2<>""),1,0)` |
| AS | Rel. Elétrica nº (aux) | fórmula |  | `=IF($AR2=1,COUNTIF($AR$2:$AR2,1),0)` |
| AT | Rel. Climatização (aux) | fórmula |  | `=IF(AND($B2=Gestao_Contrato!$B$3,$D2="Climatização",$AQ2<>""),1,0)` |
| AU | Rel. Climatização nº (aux) | fórmula |  | `=IF($AT2=1,COUNTIF($AT$2:$AT2,1),0)` |
| AV | Divergência do ponto (Atualizar Pontos) | entrada |  | |
| AW | Conferido? (Sim) | entrada |  | |
| AX | Sem motivo (aux) | fórmula |  | `=IF(AND($AQ2<>"",TRIM($X2)="",OR($AQ2="Falta justificada (atestado/comprovante)",$AQ2="Atraso/saída justificado (comprovante)",$AQ2="Remanejamento",$AQ2="Situação atípica …` |
| AY | RG do ferista que cobriu (aux) | fórmula | sim | `=IF(N($AC2)<=0,"",IF(COUNTIFS($A$2:$A$3873,$A2,$D$2:$D$3873,$D2,$C$2:$C$3873,"Ferista")=1,SUMIFS($F$2:$F$3873,$A$2:$A$3873,$A2,$D$2:$D$3873,$D2,$C$2:$C$3873,"Ferista"),IF …` |
| AZ | Valor HH do ferista (aux) | fórmula | sim | `=IF($AY2="","",IFERROR(LOOKUP(2,1/((Vigencias!$A$3:$A$149=$AY2)*(Vigencias!$D$3:$D$149<=$A2)),Vigencias!$F$3:$F$149),IFERROR(INDEX(Cadastro!$G$3:$G$150,MATCH($AY2,Cadastr …` |
| BA | Abatimento Valor/Função (R$) | fórmula |  | `=IF(OR($A2="",$AZ2="",N($AC2)<=0,NOT(ISNUMBER($I2)),$C2="Ferista",ISNUMBER(SEARCH("ENCARREGADO",$G2))),0,IF($A2<N(Tabelas!$K$18),0,ROUND(MAX(ROUND($I2,2)-ROUND($AZ2,2),0) …` |
| BB | Abatimento Valor/Função — como foi calculado | fórmula |  | `=IF(N($BA2)<=0,"","Acobertado pelo ferista / Abatimento de Valor/Função: ferista "&IFERROR(LOOKUP(2,1/((Vigencias!$A$3:$A$149=$AY2)*(Vigencias!$D$3:$D$149<=$A2)),Vigencia …` |

## Lanc_HE (cabeçalho na linha 1, 693 linhas, estado: visible)

| Col | Cabeçalho | Tipo | Oculta | Exemplo (linha 5) |
|---|---|---|---|---|
| A | Data | entrada |  | `2026-01-03 00:00:00` |
| B | Mês | fórmula |  | `=IF($A5="","",DATE(YEAR($A5),MONTH($A5),1))` |
| C | Função Solicitada | entrada |  | `-` |
| D | Equipe | fórmula |  | `=IF(TRIM($E5)="","",IFERROR(INDEX(Cadastro!$D$3:$D$150,MATCH(TRIM($E5),Cadastro!$B$3:$B$150,0)),""))` |
| E | Funcionário | entrada |  | `Delmiro Gomes dos Santos` |
| F | RG | fórmula |  | `=IF(TRIM($E5)="","",IFERROR(INDEX(Cadastro!$A$3:$A$150,MATCH(TRIM($E5),Cadastro!$B$3:$B$150,0)),""))` |
| G | Função / Custo | fórmula |  | `=IF(TRIM($E5)="","",IFERROR(IF(IFERROR(LOOKUP(2,1/((Vigencias!$A$3:$A$149=$F5)*(Vigencias!$D$3:$D$149<=$A5)),Vigencias!$E$3:$E$149),"")="",INDEX(Cadastro!$E$3:$E$149,MATC …` |
| H | Funcionario conforme Solicitação da Função | fórmula |  | `=IF(OR(TRIM($C5)="",TRIM($C5)="-"),"-",IF(TRIM($G5)=TRIM($C5),"Sim","Não"))` |
| I | Turno | fórmula |  | `=IF($G5="","",IFERROR(INDEX(Cadastro!$F$3:$F$150,MATCH(TRIM($E5),Cadastro!$B$3:$B$150,0)),""))` |
| J | Valor  HH (R$) | fórmula |  | `=IF($G5="","",IFERROR(INDEX(Tabelas!$F$3:$F$149,MATCH(TRIM($C5),Tabelas!$E$3:$E$149,0)),IFERROR(IF(IFERROR(LOOKUP(2,1/((Vigencias!$A$3:$A$149=$F5)*(Vigencias!$D$3:$D$149< …` |
| K | Valor HE  50% (R$) | fórmula |  | `=IF($J5="","",$J5*1.5)` |
| L | Valor HE  100% (R$) | fórmula |  | `=IF($J5="","",$J5*2)` |
| M | Horario solicitado | entrada |  | `-` |
| N | É  feriado? | entrada |  | `Não` |
| O | Tipo  do dia | fórmula |  | `=IF($A5="","",IF($N5="Sim","Feriado",IF(WEEKDAY($A5,2)=6,"Sabado",IF(WEEKDAY($A5,2)=7,"Domingo","Dia util"))))` |
| P | Horas Totais Trabalhadas | entrada |  | `12:06:00` |
| Q | Desconto  almoço (auto) | fórmula |  | `=IF($P5="","",IF($A5="",0,IF(AND(OR($N5="Sim",WEEKDAY($A5,2)=6,WEEKDAY($A5,2)=7),$P5>Tabelas!$K$12),Tabelas!$K$6,0)))` |
| R | Horas líquidas  HE (auto) | fórmula |  | `=IF($P5="","",MAX($P5-$Q5,0))` |
| S | HE  50% (h) | fórmula |  | `=IF($R5="","",IF($O5="Dia util",$R5,IF($O5="Sabado",MAX($R5-Tabelas!$K$5,0),0)))` |
| T | HE  100% (h) | fórmula |  | `=IF($R5="","",IF(OR($O5="Domingo",$O5="Feriado"),$R5,IF($O5="Sabado",MIN($R5,Tabelas!$K$5),0)))` |
| U | Valor  HE (R$) | fórmula |  | `=IF(OR($R5="",$K5="",$L5=""),"",$S5*24*$K5+$T5*24*$L5)` |
| V | Tratamento  HE | entrada |  | `PAGAS POR EI` |
| W | HE pagas  por EI (R$) | fórmula |  | `=IF($V5="PAGAS POR EI",IFERROR($U5,0),0)` |
| X | HE pagas  pela área (R$) | fórmula |  | `=IF($V5="PAGAS PELA ÁREA",IFERROR($U5,0),0)` |
| Y | Solicitante | entrada |  | |
| Z | Área  solicitação | entrada |  | |
| AA | Descrição | entrada |  | |
| AB | Histórico dos pontos | entrada |  | `05:10#17:46 = 12:06` |
| AC | Recebimento do Historico dos Pontos | entrada |  | `Sim` |
| AD | Requisição de compra e pedido | entrada |  | |
| AE | Conferência | fórmula |  | `=IF($A5="","",IF(AND($V5="HORAS COMPENSADAS",COUNTIFS(Compensacoes!$A$4:$A$503,$A5,Compensacoes!$B$4:$B$503,$E5)=0),"REGISTRAR na aba Compensações (quem compensou e se fo …` |
| AF | Rel. HE (aux) | fórmula |  | `=IF(AND($A5<>"",$B5=Gestao_Contrato!$B$3),1,0)` |
| AG | Rel. HE nº (aux) | fórmula |  | `=IF($AF5=1,COUNTIF($AF$2:$AF5,1),0)` |
| AH | Divergência do ponto (Atualizar Pontos) | entrada |  | |
| AI | Conferido? (Sim) | entrada |  | |

## Compensacoes (cabeçalho na linha 3, 503 linhas, estado: visible)

| Col | Cabeçalho | Tipo | Oculta | Exemplo (linha 4) |
|---|---|---|---|---|
| A | Data em que trabalhou | entrada |  | `2026-08-08 00:00:00` |
| B | Quem trabalhou | entrada |  | `Delmiro Gomes dos Santos` |
| C | Horas compensadas | entrada |  | `12:28:00` |
| D | Compensa a falta de (opcional) | entrada |  | |
| E | Data da falta (opcional) | entrada |  | |
| F | Aprovado? | entrada |  | `Sim` |
| G | Aprovado por / observação | entrada |  | `Copa (ago/26): quita as pendências das duas equipes; a sobra abateu o absenteísmo da Elétrica` |
| H | Mês | fórmula |  | `=IF($A4="","",DATE(YEAR($A4),MONTH($A4),1))` |
| I | Equipe | fórmula |  | `=IF(TRIM($B4)="","",IFERROR(INDEX(Cadastro!$D$3:$D$150,MATCH(TRIM($B4),Cadastro!$B$3:$B$150,0)),"?"))` |
| J | Horas que contam | fórmula |  | `=IF(AND($F4="Sim",$A4<>""),N($C4),"")` |
| K | Tem HE "HORAS COMPENSADAS"? | fórmula |  | `=IF($A4="","",IF(COUNTIFS(Lanc_HE!$A$2:$A$693,$A4,Lanc_HE!$E$2:$E$693,$B4,Lanc_HE!$V$2:$V$693,"HORAS COMPENSADAS")+COUNTIFS(Lanc_HORAS!$A$2:$A$3873,$A4,Lanc_HORAS!$E$2:$E …` |
| L | Situação | fórmula |  | `=IF($A4="","",IF(OR(TRIM($B4)="",N($C4)<=0),"PREENCHER: quem trabalhou e as horas",IF($I4="?","CONFERIR: funcionário fora do Cadastro",IF($F4="","PREENCHER: aprovado? (Si …` |
| M | Falta encontrada? | fórmula |  | `=IF(OR($A4="",TRIM($D4)=""),"",IF(COUNTIFS(Lanc_HORAS!$A$2:$A$3873,$E4,Lanc_HORAS!$E$2:$E$3873,$D4)>0,"Sim","Não"))` |
| N | Rel. (aux) | fórmula |  | `=IF(AND($A4<>"",$H4=Gestao_Contrato!$B$3),1,0)` |
| O | Rel. nº (aux) | fórmula |  | `=IF($N4=1,COUNTIF($N$4:$N4,1),0)` |

## Cadastro (cabeçalho na linha 2, 978 linhas, estado: visible)

| Col | Cabeçalho | Tipo | Oculta | Exemplo (linha 3) |
|---|---|---|---|---|
| A | Registro | entrada |  | `520805` |
| B | Funcionário | entrada |  | `Bruno Domingues Pereira Barbosa` |
| C | Estilo do Contrato | entrada |  | `Normal / Fixo` |
| D | Equipe | entrada |  | `Elétrica` |
| E | Função / Custo | entrada |  | `TÉC. ELETRICISTA I` |
| F | Turno | entrada |  | `TURNO 2` |
| G | Valor HH (R$) | fórmula |  | `=IF($E3="","",IFERROR(INDEX(Tabelas!$F$3:$F$150,MATCH(TRIM($E3),Tabelas!$E$3:$E$150,0)),""))` |
| H | Valor HE 50% (R$) | fórmula |  | `=IF($E3="","",IFERROR(INDEX(Tabelas!$G$3:$G$150,MATCH(TRIM($E3),Tabelas!$E$3:$E$150,0)),""))` |
| I | Valor HE 100% (R$) | fórmula |  | `=IF($E3="","",IFERROR(INDEX(Tabelas!$H$3:$H$150,MATCH(TRIM($E3),Tabelas!$E$3:$E$150,0)),""))` |
| J | Status | entrada |  | `Ativo` |
| K | Observação | entrada |  | |

## Vigencias (cabeçalho na linha 2, 977 linhas, estado: visible)

| Col | Cabeçalho | Tipo | Oculta | Exemplo (linha 3) |
|---|---|---|---|---|
| A | Registro | entrada |  | `520805` |
| B | Funcionário | entrada |  | `Bruno Domingues Pereira Barbosa` |
| C | Equipe | entrada |  | `Elétrica` |
| D | Início vigência | entrada |  | `2026-01-01 00:00:00` |
| E | Função custo | entrada |  | `TÉC. ELETRICISTA I` |
| F | Valor HH (R$) | entrada |  | `49.6356094764629` |
| G | Turno | entrada |  | `TURNO 2` |
| H | Estilo do Contrato | entrada |  | `Normal / Fixo` |

## Analises (cabeçalho na linha 2, 103 linhas, estado: visible)

| Col | Cabeçalho | Tipo | Oculta | Exemplo (linha 3) |
|---|---|---|---|---|
| A | Mês | fórmula |  | `=DATE(Tabelas!$K$3,1,1)` |
| B | Horas  absenteismo | fórmula |  | `=MAX(N($AU3)-N($AW3)-N($BG3),0)` |
| C | Falta de Contratação | fórmula |  | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$A3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Falta de Contratação")` |
| D | Sem Justificativa | fórmula |  | `=$B3-$AS3-$AM3` |
| E | Falta de EPI's | fórmula |  | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$A3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Falta de EPI")` |
| F | Treinamento | fórmula |  | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$A3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Treinamento")` |
| G | Recesso | fórmula |  | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$A3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Recesso")` |
| H | Férias | fórmula |  | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$A3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Férias")` |
| I | Abatimento do  contrato (R$) | fórmula |  | `=MAX(N($AV3)-N($AT3),0)+N($BH3)` |
| J | Horas HE 50% | fórmula |  | `=SUMIFS(Lanc_HE!$S$2:$S$693,Lanc_HE!$B$2:$B$693,$A3,Lanc_HE!$D$2:$D$693,"Elétrica",Lanc_HE!$V$2:$V$693,"PAGAS POR EI")` |
| K | Horas HE 100% | fórmula |  | `=SUMIFS(Lanc_HE!$T$2:$T$693,Lanc_HE!$B$2:$B$693,$A3,Lanc_HE!$D$2:$D$693,"Elétrica",Lanc_HE!$V$2:$V$693,"PAGAS POR EI")` |
| L | HE pagas por EI  lançada (R$) | fórmula |  | `=SUMIFS(Lanc_HE!$W$2:$W$693,Lanc_HE!$B$2:$B$693,$A3,Lanc_HE!$D$2:$D$693,"Elétrica")` |
| M | HE pagas  pela área (R$) | fórmula |  | `=SUMIFS(Lanc_HE!$X$2:$X$693,Lanc_HE!$B$2:$B$693,$A3,Lanc_HE!$D$2:$D$693,"Elétrica")` |
| N | HE pagas por  EI total (R$) | fórmula |  | `=L3` |
| O | Base do  abatimento (R$) | fórmula |  | `=I3` |
| P | HE pagas por  EI usada (R$) | fórmula |  | `=N3` |
| Q | Horas Compensadas (aprovadas) | fórmula |  | `=SUMIFS(Compensacoes!$J$4:$J$503,Compensacoes!$H$4:$H$503,$A3,Compensacoes!$I$4:$I$503,"Elétrica")` |
| R | HE a pagar  pela área (R$) | fórmula |  | `=M3` |
| S | Saldo do  mês (R$) | fórmula |  | `=IFERROR(I3+AG3-P3,0)` |
| T | Qtd.  eventos abs. | fórmula |  | `=COUNTIFS(Lanc_HORAS!$B$2:$B$3873,$A3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$T$2:$T$3873,"Sim",Lanc_HORAS!$V$2:$V$3873,">0")+COUNTIFS(Lanc_HORAS!$B$2:$B$3873,$A3, …` |
| U | Status  do saldo | fórmula |  | `=IF(S3<0,"NEGATIVO",IF(S3>0,"POSITIVO","ZERADO"))` |
| V | Horas  justificadas | fórmula |  | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$A3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências C/ Justificativa")+SUMIFS(Lanc_HORAS!$V$ …` |
| W | Valor  justificadas (R$) | fórmula |  | `=SUMIFS(Lanc_HORAS!$AB$2:$AB$3873,Lanc_HORAS!$B$2:$B$3873,$A3,Lanc_HORAS!$D$2:$D$3873,"Elétrica")` |
| X | Horas SPOT | fórmula |  | `=SUMIFS(Lanc_HORAS!$U$2:$U$3873,Lanc_HORAS!$B$2:$B$3873,$A3,Lanc_HORAS!$D$2:$D$3873,"SPOT")` |
| Y | Horas Ferista sem Cobertura (h) | fórmula |  | `=MAX(SUMIFS(Lanc_HORAS!$U$2:$U$3873,Lanc_HORAS!$B$2:$B$3873,$A3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$C$2:$C$3873,"Ferista")-SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lan …` |
| Z | Contrato  base (R$) | fórmula |  | `=IFERROR(LOOKUP(2,1/((Tabelas!$U$3:$U$30="Elétrica")*(Tabelas!$V$3:$V$30<=$A3)*(Tabelas!$W$3:$W$30<>"")),Tabelas!$W$3:$W$30),Tabelas!$K$7)` |
| AA | Saldo do  mês (R$) | fórmula |  | `=S3` |
| AB | Saldo acumulado  ano fiscal (R$) | fórmula |  | `=S3` |
| AC | Horas equivalentes  do abatimento | fórmula |  | `=B3` |
| AD | Horas HE  pagas por EI | fórmula |  | `=J3+K3` |
| AE | Referência da  HE paga por EI | fórmula |  | `=IF(P3=0,"Sem uso do banco","HE paga por EI")` |
| AF | Valor a  pagar no mês (R$) | fórmula |  | `=Z3-I3` |
| AG | Abatimento  SLA (R$) | fórmula |  | `=N(SLA!$C$6)` |
| AH | Valor final a  pagar no mês (R$) | fórmula |  | `=IFERROR(Z3-I3-AG3,0)` |
| AI | Saldo virtual  disponível acumulado (R$) | fórmula |  | `=AB3` |
| AJ | Horas  trabalhadas | fórmula |  | `=SUMIFS(Lanc_HORAS!$U$2:$U$3873,Lanc_HORAS!$B$2:$B$3873,$A3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$C$2:$C$3873,"<>Ferista")+SUMIFS(Lanc_HE!$R$2:$R$693,Lanc_HE!$B$ …` |
| AK | Cobertura do Ferista (h) | fórmula |  | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$A3,Lanc_HORAS!$D$2:$D$3873,"Elétrica")` |
| AL | Horas de Remanejamento (h) | fórmula |  | `=SUMIFS(Lanc_HORAS!$U$2:$U$3873,Lanc_HORAS!$B$2:$B$3873,$A3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Remanejamento")` |
| AM | Leves Atrasos (± 10 mn dia) | fórmula |  | `=SUMIFS(Lanc_HORAS!$AM$2:$AM$3873,Lanc_HORAS!$B$2:$B$3873,$A3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Horas Trabalhadas",Lanc_HORAS!$V$2:$V$3873,">"&1 …` |
| AN | Situação Não Acordada (h) | fórmula |  | `=SUMIFS(Lanc_HORAS!$AL$2:$AL$3873,Lanc_HORAS!$B$2:$B$3873,$A3,Lanc_HORAS!$D$2:$D$3873,"Elétrica")` |
| AO | HE Total Acordado (R$) | fórmula |  | `=SUMIFS(Lanc_HE!$W$2:$W$693,Lanc_HE!$B$2:$B$693,$A3,Lanc_HE!$D$2:$D$693,"Elétrica")+SUMIFS(Lanc_HE!$X$2:$X$693,Lanc_HE!$B$2:$B$693,$A3,Lanc_HE!$D$2:$D$693,"Elétrica")` |
| AP | HE Total Acordado (h) | fórmula |  | `=SUMIFS(Lanc_HE!$R$2:$R$693,Lanc_HE!$B$2:$B$693,$A3,Lanc_HE!$D$2:$D$693,"Elétrica",Lanc_HE!$V$2:$V$693,"PAGAS POR EI")+SUMIFS(Lanc_HE!$R$2:$R$693,Lanc_HE!$B$2:$B$693,$A3, …` |
| AQ | Horas Trabalhadas Ferista (h) | fórmula |  | `=SUMIFS(Lanc_HORAS!$U$2:$U$3873,Lanc_HORAS!$B$2:$B$3873,$A3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$C$2:$C$3873,"Ferista")` |
| AR | Situação Resolvida (Acordo) (h) | fórmula |  | `=SUMIFS(Lanc_HORAS!$AJ$2:$AJ$3873,Lanc_HORAS!$B$2:$B$3873,$A3,Lanc_HORAS!$D$2:$D$3873,"Elétrica")` |
| AS | Horas das Férias Não Cobertas (h) | fórmula |  | `=SUMIFS(Lanc_HORAS!$AM$2:$AM$3873,Lanc_HORAS!$B$2:$B$3873,$A3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Férias")` |
| AT | Valor Compensado (R$) (proporcional ao absenteísmo) | fórmula |  | `=IF(N($AU3)<=0,0,N($AV3)*(N($AW3)+N($BG3))/N($AU3))` |
| AU | Horas Absenteísmo BRUTO (aux) | fórmula |  | `=SUMIFS(Lanc_HORAS!$AM$2:$AM$3873,Lanc_HORAS!$B$2:$B$3873,$A3,Lanc_HORAS!$D$2:$D$3873,"Elétrica")` |
| AV | Valor Absenteísmo BRUTO (aux) | fórmula |  | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$A3,Lanc_HORAS!$D$2:$D$3873,"Elétrica")` |
| AW | Horas Compensação UTILIZADAS | fórmula |  | `=MIN(MAX(N($Q3)-N($BD3)-MIN(MAX((0+N($BC19))-MIN((0+N($BC19)),N($Q19)),0),MAX(N($Q3)-N($BD3),0)),0),N($AU3))` |
| AX | Horas Positivas (excedente p/ acerto comercial) | fórmula |  | `=MAX(MAX(N($Q3)-N($BD3)-MIN(MAX((0+N($BC19))-MIN((0+N($BC19)),N($Q19)),0),MAX(N($Q3)-N($BD3),0)),0)-N($AW3),0)` |
| AY | HE NÃO paga (h) | fórmula |  | `=MAX(SUMIFS(Lanc_HE!$R$2:$R$693,Lanc_HE!$B$2:$B$693,$A3,Lanc_HE!$D$2:$D$693,"Elétrica")-N($AP3)-N($BA3),0)` |
| AZ | HE NÃO paga (R$) | fórmula |  | `=MAX(SUMIFS(Lanc_HE!$U$2:$U$693,Lanc_HE!$B$2:$B$693,$A3,Lanc_HE!$D$2:$D$693,"Elétrica")-N($AO3)-N($BB3),0)` |
| BA | HE Remanejada p/ outra área (h) (só registro) | fórmula |  | `=SUMIFS(Lanc_HE!$R$2:$R$693,Lanc_HE!$B$2:$B$693,$A3,Lanc_HE!$D$2:$D$693,"Elétrica",Lanc_HE!$V$2:$V$693,"REMANEJADO")+SUMIFS(Lanc_HE!$R$2:$R$693,Lanc_HE!$B$2:$B$693,$A3,La …` |
| BB | HE Remanejada p/ outra área (R$) (só registro) | fórmula |  | `=SUMIFS(Lanc_HE!$U$2:$U$693,Lanc_HE!$B$2:$B$693,$A3,Lanc_HE!$D$2:$D$693,"Elétrica",Lanc_HE!$V$2:$V$693,"REMANEJADO")+SUMIFS(Lanc_HE!$U$2:$U$693,Lanc_HE!$B$2:$B$693,$A3,La …` |
| BC | Pendentes geradas no mês (h) | fórmula |  | `=SUMIFS(Lanc_HORAS!$Q$2:$Q$3873,Lanc_HORAS!$B$2:$B$3873,$A3,Lanc_HORAS!$D$2:$D$3873,"Elétrica")` |
| BD | Pendentes QUITADAS no mês (h) | fórmula |  | `=MIN((0+N($BC3)),N($Q3))` |
| BE | Saldo PENDENTE acumulado (h) | fórmula |  | `=MAX((0+N($BC3))-N($BD3),0)` |
| BF | Horas Adicional não solicitado (h) | fórmula |  | `=SUMIFS(Lanc_HORAS!$U$2:$U$3873,Lanc_HORAS!$B$2:$B$3873,$A3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Adicional Não Solicitado")` |
| BG | Horas Adicional UTILIZADAS (h) | fórmula |  | `=MIN(N($BF3),MAX(N($AU3)-N($AW3),0))` |
| BH | Abatimento Valor/Função (R$) | fórmula |  | `=SUMIFS(Lanc_HORAS!$BA$2:$BA$3873,Lanc_HORAS!$B$2:$B$3873,$A3,Lanc_HORAS!$D$2:$D$3873,"Elétrica")` |
| BI | Horas cobertas com diferença de função (h) | fórmula |  | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$A3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$BA$2:$BA$3873,">0")` |

## Gestao_Contrato (célula a célula, linhas 1–48)

| Célula | Conteúdo |
|---|---|
| B2 | `Relatório do mês` |
| C2 | `Análises (resumo do mês)` |
| D2 | `Colaboradores` |
| F2 | `Painel de Controle` |
| G2 | `Ferista` |
| H2 | `Como usar` |
| B3 | `2026-09-01 00:00:00` |
| C3 | `◄  FILTRO DE ANÁLISE DO MÊS` |
| D3 | `=HYPERLINK("#'SLA'!A1","SLA")` |
| B4 | `Lançamentos (horas)` |
| C4 | `Horas extras` |
| D4 | `Compensações` |
| F4 | `Cadastro / Vigências` |
| G4 | `Tabelas e calendário` |
| H4 | `=IF(SUM(Painel_De_Controle!N17:Y17)>0,SUM(Painel_De_Controle!N17:Y17)&" divergência(s) de ponto","Ponto conferido")` |
| B5 | `ELÉTRICA` |
| F5 | `CLIMATIZAÇÃO` |
| B6 | `CONTRATO` |
| C6 | `ABATIMENTO` |
| D6 | `VALOR A PAGAR` |
| F6 | `CONTRATO` |
| G6 | `ABATIMENTO` |
| H6 | `VALOR A PAGAR` |
| B7 | `=IFERROR(INDEX(Analises!$Z$3:$Z$14,MATCH($B$3,Analises!$A$3:$A$14,0)),0)` |
| C7 | `=IFERROR(INDEX(Analises!$I$3:$I$14,MATCH($B$3,Analises!$A$3:$A$14,0)),0)` |
| D7 | `=IFERROR(INDEX(Analises!$AH$3:$AH$14,MATCH($B$3,Analises!$A$3:$A$14,0)),0)` |
| F7 | `=IFERROR(INDEX(Analises!$Z$19:$Z$30,MATCH($B$3,Analises!$A$19:$A$30,0)),0)` |
| G7 | `=IFERROR(INDEX(Analises!$I$19:$I$30,MATCH($B$3,Analises!$A$19:$A$30,0)),0)` |
| H7 | `=IFERROR(INDEX(Analises!$AH$19:$AH$30,MATCH($B$3,Analises!$A$19:$A$30,0)),0)` |
| B8 | `Desconto SLA (já descontado do valor a pagar)` |
| C8 | `=IFERROR(INDEX(Analises!$AG$3:$AG$14,MATCH($B$3,Analises!$A$3:$A$14,0)),0)` |
| F8 | `Desconto SLA (já descontado do valor a pagar)` |
| G8 | `=IFERROR(INDEX(Analises!$AG$19:$AG$30,MATCH($B$3,Analises!$A$19:$A$30,0)),0)` |
| B9 | `DE ONDE VEM O ABATIMENTO` |
| F9 | `DE ONDE VEM O ABATIMENTO` |
| B10 | `DETALHE POR SITUAÇÃO` |
| C10 | `HORAS` |
| D10 | `VALOR (R$)` |
| F10 | `DETALHE POR SITUAÇÃO` |
| G10 | `HORAS` |
| H10 | `VALOR (R$)` |
| B11 | `Férias Não Cobertas` |
| C11 | `=SUMIFS(Lanc_HORAS!$AM$2:$AM$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Férias")` |
| D11 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Férias")` |
| F11 | `Férias Não Cobertas` |
| G11 | `=SUMIFS(Lanc_HORAS!$AM$2:$AM$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Férias")` |
| H11 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Férias")` |
| B12 | `FALTAS Injustificadas` |
| C12 | `=SUMIFS(Lanc_HORAS!$AM$2:$AM$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências S/ Justificativa")` |
| D12 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências S/ Justificativa")` |
| F12 | `FALTAS Injustificadas` |
| G12 | `=SUMIFS(Lanc_HORAS!$AM$2:$AM$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências S/ Justificativa")` |
| H12 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências S/ Justificativa")` |
| B13 | `ATRASOS / Saídas Antecipadas` |
| C13 | `=SUMIFS(Lanc_HORAS!$AM$2:$AM$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Entrada Atrasada ou Saída Antecipada S/ Justificativa")+SUMIFS(Lanc_HORAS!$AM …` |
| D13 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Entrada Atrasada ou Saída Antecipada S/ Justificativa")+SUMIFS(Lanc_HORAS!$W$2: …` |
| F13 | `ATRASOS / Saídas Antecipadas` |
| G13 | `=SUMIFS(Lanc_HORAS!$AM$2:$AM$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Entrada Atrasada ou Saída Antecipada S/ Justificativa")+SUMIFS(Lanc_HORAS …` |
| H13 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Entrada Atrasada ou Saída Antecipada S/ Justificativa")+SUMIFS(Lanc_HORAS!$ …` |
| B14 | `Falta de Contratação` |
| C14 | `=SUMIFS(Lanc_HORAS!$AM$2:$AM$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Falta de Contratação")` |
| D14 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Falta de Contratação")` |
| F14 | `Falta de Contratação` |
| G14 | `=SUMIFS(Lanc_HORAS!$AM$2:$AM$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Falta de Contratação")` |
| H14 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Falta de Contratação")` |
| B15 | `Falta de EPI` |
| C15 | `=SUMIFS(Lanc_HORAS!$AM$2:$AM$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Falta de EPI")` |
| D15 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Falta de EPI")` |
| F15 | `Falta de EPI` |
| G15 | `=SUMIFS(Lanc_HORAS!$AM$2:$AM$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Falta de EPI")` |
| H15 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Falta de EPI")` |
| B16 | `Ausências Devido a Recesso` |
| C16 | `=SUMIFS(Lanc_HORAS!$AM$2:$AM$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Recesso")` |
| D16 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Recesso")` |
| F16 | `Ausências Devido a Recesso` |
| G16 | `=SUMIFS(Lanc_HORAS!$AM$2:$AM$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Recesso")` |
| H16 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Recesso")` |
| B17 | `Remanejamento` |
| C17 | `=SUMIFS(Lanc_HORAS!$AM$2:$AM$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Remanejamento")` |
| D17 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Remanejamento")` |
| F17 | `Remanejamento` |
| G17 | `=SUMIFS(Lanc_HORAS!$AM$2:$AM$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Remanejamento")` |
| H17 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Remanejamento")` |
| B18 | `Situação Não Acordada` |
| C18 | `=SUMIFS(Lanc_HORAS!$AM$2:$AM$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Situação Atipica")` |
| D18 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Situação Atipica")` |
| F18 | `Situação Não Acordada` |
| G18 | `=SUMIFS(Lanc_HORAS!$AM$2:$AM$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Situação Atipica")` |
| H18 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Situação Atipica")` |
| B19 | `Outras (recusa, integração, afastamento, desligamento)` |
| C19 | `=SUMIFS(Lanc_HORAS!$AM$2:$AM$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Recusa de Atividade S/ Motivo Legítimo")+SUMIFS(Lanc_HORAS!$AM$2:$AM$3873,Lan …` |
| D19 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Recusa de Atividade S/ Motivo Legítimo")+SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HO …` |
| F19 | `Outras (recusa, integração, afastamento, desligamento)` |
| G19 | `=SUMIFS(Lanc_HORAS!$AM$2:$AM$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Recusa de Atividade S/ Motivo Legítimo")+SUMIFS(Lanc_HORAS!$AM$2:$AM$3873 …` |
| H19 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Recusa de Atividade S/ Motivo Legítimo")+SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lan …` |
| B20 | `( − ) Compensadas + adicional não solicitado` |
| C20 | `=IFERROR(INDEX(Analises!$AW$3:$AW$14,MATCH($B$3,Analises!$A$3:$A$14,0))+INDEX(Analises!$BG$3:$BG$14,MATCH($B$3,Analises!$A$3:$A$14,0)),0)` |
| D20 | `=IFERROR(INDEX(Analises!$AT$3:$AT$14,MATCH($B$3,Analises!$A$3:$A$14,0)),0)` |
| F20 | `( − ) Compensadas + adicional não solicitado` |
| G20 | `=IFERROR(INDEX(Analises!$AW$19:$AW$30,MATCH($B$3,Analises!$A$19:$A$30,0))+INDEX(Analises!$BG$19:$BG$30,MATCH($B$3,Analises!$A$19:$A$30,0)),0)` |
| H20 | `=IFERROR(INDEX(Analises!$AT$19:$AT$30,MATCH($B$3,Analises!$A$19:$A$30,0)),0)` |
| B21 | `( + ) Abatimento de Valor/Função (ver quadro abaixo)` |
| D21 | `=IFERROR(INDEX(Analises!$BH$3:$BH$14,MATCH($B$3,Analises!$A$3:$A$14,0)),0)` |
| F21 | `( + ) Abatimento de Valor/Função (ver quadro abaixo)` |
| H21 | `=IFERROR(INDEX(Analises!$BH$19:$BH$30,MATCH($B$3,Analises!$A$19:$A$30,0)),0)` |
| B22 | `TOTAL ABATIDO` |
| C22 | `=MAX(SUM(C11:C19)-C20,0)` |
| D22 | `=MAX(SUM(D11:D19)-D20,0)+D21` |
| F22 | `TOTAL ABATIDO` |
| G22 | `=MAX(SUM(G11:G19)-G20,0)` |
| H22 | `=MAX(SUM(H11:H19)-H20,0)+H21` |
| B23 | `COBERTURA DO FERISTA ELETRICA E COMPENSAÇÃO` |
| F23 | `COBERTURA DO FERISTA CLIMATIZAÇÃO E COMPENSAÇÃO` |
| B24 | `Situação coberta` |
| C24 | `HORAS` |
| F24 | `Situação coberta` |
| G24 | `HORAS` |
| B25 | `Total Compensação` |
| C25 | `=IFERROR(INDEX(Analises!$Q$3:$Q$14,MATCH($B$3,Analises!$A$3:$A$14,0)),0)` |
| F25 | `Total Compensação` |
| G25 | `=IFERROR(INDEX(Analises!$Q$19:$Q$30,MATCH($B$3,Analises!$A$19:$A$30,0)),0)` |
| B26 | `Total Pendencias (ambas as equipes)` |
| C26 | `=IFERROR(INDEX(Analises!$BE$3:$BE$14,MATCH($B$3,Analises!$A$3:$A$14,0)),0)+IFERROR(INDEX(Analises!$BE$19:$BE$30,MATCH($B$3,Analises!$A$19:$A$30,0)),0)` |
| F26 | `Total Pendencias (ambas as equipes)` |
| G26 | `=IFERROR(INDEX(Analises!$BE$3:$BE$14,MATCH($B$3,Analises!$A$3:$A$14,0)),0)+IFERROR(INDEX(Analises!$BE$19:$BE$30,MATCH($B$3,Analises!$A$19:$A$30,0)),0)` |
| B27 | `Sobra de Horas de Compensação Apenas da Equipe Eletrica` |
| C27 | `=IFERROR(INDEX(Analises!$AW$3:$AW$14,MATCH($B$3,Analises!$A$3:$A$14,0)),0)` |
| F27 | `Sobra de Horas de Compensação Apenas da Equipe Climatização` |
| G27 | `=IFERROR(INDEX(Analises!$AW$19:$AW$30,MATCH($B$3,Analises!$A$19:$A$30,0)),0)` |
| B28 | `Faltas cobertas` |
| C28 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências S/ Justificativa")` |
| F28 | `Faltas cobertas` |
| G28 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências S/ Justificativa")` |
| B29 | `Atrasos / Saídas cobertas` |
| C29 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Entrada Atrasada ou Saída Antecipada S/ Justificativa")+SUMIFS(Lanc_HORAS!$AC …` |
| F29 | `Atrasos / Saídas cobertas` |
| G29 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Entrada Atrasada ou Saída Antecipada S/ Justificativa")+SUMIFS(Lanc_HORAS …` |
| B30 | `Férias cobertas` |
| C30 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Férias")` |
| F30 | `Férias cobertas` |
| G30 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Férias")` |
| B31 | `Remanejamento coberto` |
| C31 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Remanejamento")` |
| F31 | `Remanejamento coberto` |
| G31 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Remanejamento")` |
| B32 | `Total coberto pelo ferista` |
| C32 | `=SUM(C28:C31)` |
| F32 | `Total coberto pelo ferista` |
| G32 | `=SUM(G28:G31)` |
| B33 | `Horas Pendentes Quitadas Neste Mês` |
| C33 | `=IFERROR(INDEX(Analises!$BD$3:$BD$14,MATCH($B$3,Analises!$A$3:$A$14,0)),0)` |
| F33 | `Horas Pendentes Quitadas Neste Mês` |
| G33 | `=IFERROR(INDEX(Analises!$BD$19:$BD$30,MATCH($B$3,Analises!$A$19:$A$30,0)),0)` |
| B34 | `Horas Pendentes` |
| C34 | `=IFERROR(INDEX(Analises!$BE$3:$BE$14,MATCH($B$3,Analises!$A$3:$A$14,0)),0)` |
| F34 | `Horas Pendentes` |
| G34 | `=IFERROR(INDEX(Analises!$BE$19:$BE$30,MATCH($B$3,Analises!$A$19:$A$30,0)),0)` |
| B36 | `NÃO ABATEM VALOR DO CONTRATO` |
| F36 | `NÃO ABATEM VALOR DO CONTRATO` |
| B37 | `Horas Justificadas` |
| C37 | `=IFERROR(INDEX(Analises!$V$3:$V$14,MATCH($B$3,Analises!$A$3:$A$14,0)),0)` |
| F37 | `Horas Justificadas` |
| G37 | `=IFERROR(INDEX(Analises!$V$19:$V$30,MATCH($B$3,Analises!$A$19:$A$30,0)),0)` |
| B38 | `Horas Acordadas` |
| C38 | `=SUMIFS(Lanc_HORAS!$P$2:$P$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Elétrica")` |
| F38 | `Horas Acordadas` |
| G38 | `=SUMIFS(Lanc_HORAS!$P$2:$P$3873,Lanc_HORAS!$B$2:$B$3873,$B$3,Lanc_HORAS!$D$2:$D$3873,"Climatização")` |
| B40 | `FERISTA × CARGO COBERTO — RELAÇÃO DO VALOR-HORA (ELÉTRICA)` |
| F40 | `FERISTA × CARGO COBERTO — RELAÇÃO DO VALOR-HORA (CLIMATIZAÇÃO)` |
| B41 | `Ferista do mês` |
| C41 | `=IF(OR(Analises!$N$95="",Analises!$N$95="—"),"—",Analises!$N$95)` |
| F41 | `Ferista do mês` |
| G41 | `=IF(OR(Analises!$N$96="",Analises!$N$96="—"),"—",Analises!$N$96)` |
| B42 | `Função e valor-hora do ferista` |
| C42 | `=IFERROR(LOOKUP(2,1/((Vigencias!$A$3:$A$149=IFERROR(INDEX(Cadastro!$A$3:$A$150,MATCH(Analises!$N$95,Cadastro!$B$3:$B$150,0)),""))*(Vigencias!$D$3:$D$149<=EOMONTH($B$3,0))),Vigencias!$E$3:$E$149),"")` |
| D42 | `=IFERROR(LOOKUP(2,1/((Vigencias!$A$3:$A$149=IFERROR(INDEX(Cadastro!$A$3:$A$150,MATCH(Analises!$N$95,Cadastro!$B$3:$B$150,0)),""))*(Vigencias!$D$3:$D$149<=EOMONTH($B$3,0))),Vigencias!$F$3:$F$149),"")` |
| F42 | `Função e valor-hora do ferista` |
| G42 | `=IFERROR(LOOKUP(2,1/((Vigencias!$A$3:$A$149=IFERROR(INDEX(Cadastro!$A$3:$A$150,MATCH(Analises!$N$96,Cadastro!$B$3:$B$150,0)),""))*(Vigencias!$D$3:$D$149<=EOMONTH($B$3,0))),Vigencias!$E$3:$E$149),"")` |
| H42 | `=IFERROR(LOOKUP(2,1/((Vigencias!$A$3:$A$149=IFERROR(INDEX(Cadastro!$A$3:$A$150,MATCH(Analises!$N$96,Cadastro!$B$3:$B$150,0)),""))*(Vigencias!$D$3:$D$149<=EOMONTH($B$3,0))),Vigencias!$F$3:$F$149),"")` |
| B43 | `Cargo coberto (técnico)` |
| C43 | `Valor-hora do cargo` |
| D43 | `Diferença para o ferista` |
| F43 | `Cargo coberto (técnico)` |
| G43 | `Valor-hora do cargo` |
| H43 | `Diferença para o ferista` |
| B44 | `TÉC. ELETRICISTA I` |
| C44 | `=IFERROR(LOOKUP(2,1/((Vigencias!$E$3:$E$149=B44)*(Vigencias!$D$3:$D$149<=EOMONTH($B$3,0))),Vigencias!$F$3:$F$149),IFERROR(INDEX(Tabelas!$F$3:$F$150,MATCH(B44,Tabelas!$E$3:$E$150,0)),""))` |
| D44 | `=IF(OR($D$42="",C44=""),"",MAX(ROUND(C44,2)-ROUND($D$42,2),0))` |
| F44 | `TÉC. CLIMATIZAÇÃO I` |
| G44 | `=IFERROR(LOOKUP(2,1/((Vigencias!$E$3:$E$149=F44)*(Vigencias!$D$3:$D$149<=EOMONTH($B$3,0))),Vigencias!$F$3:$F$149),IFERROR(INDEX(Tabelas!$F$3:$F$150,MATCH(F44,Tabelas!$E$3:$E$150,0)),""))` |
| H44 | `=IF(OR($H$42="",G44=""),"",MAX(ROUND(G44,2)-ROUND($H$42,2),0))` |
| B45 | `TÉC. ELETRICISTA II` |
| C45 | `=IFERROR(LOOKUP(2,1/((Vigencias!$E$3:$E$149=B45)*(Vigencias!$D$3:$D$149<=EOMONTH($B$3,0))),Vigencias!$F$3:$F$149),IFERROR(INDEX(Tabelas!$F$3:$F$150,MATCH(B45,Tabelas!$E$3:$E$150,0)),""))` |
| D45 | `=IF(OR($D$42="",C45=""),"",MAX(ROUND(C45,2)-ROUND($D$42,2),0))` |
| B46 | `TÉC. TELECOM` |
| C46 | `=IFERROR(LOOKUP(2,1/((Vigencias!$E$3:$E$149=B46)*(Vigencias!$D$3:$D$149<=EOMONTH($B$3,0))),Vigencias!$F$3:$F$149),IFERROR(INDEX(Tabelas!$F$3:$F$150,MATCH(B46,Tabelas!$E$3:$E$150,0)),""))` |
| D46 | `=IF(OR($D$42="",C46=""),"",MAX(ROUND(C46,2)-ROUND($D$42,2),0))` |
| B47 | `ABATIMENTO DE VALOR/FUNÇÃO DO MÊS` |
| D47 | `=IFERROR(INDEX(Analises!$BH$3:$BH$14,MATCH($B$3,Analises!$A$3:$A$14,0)),0)` |
| F47 | `ABATIMENTO DE VALOR/FUNÇÃO DO MÊS` |
| H47 | `=IFERROR(INDEX(Analises!$BH$19:$BH$30,MATCH($B$3,Analises!$A$19:$A$30,0)),0)` |
| B48 | `Só desconta quando o cargo coberto vale mais que o ferista; ferista igual ou maior = sem desconto. Encarregado nunca é coberto. O cálculo de cada dia está na Lanc_HORAS (col. BB) e no Relatório.` |
| F48 | `Só desconta quando o cargo coberto vale mais que o ferista; ferista igual ou maior = sem desconto. Encarregado nunca é coberto. O cálculo de cada dia está na Lanc_HORAS (col. BB) e no Relatório.` |

## Painel_De_Controle (célula a célula, linhas 1–48)

| Célula | Conteúdo |
|---|---|
| B2 | `Mês analisado` |
| C2 | `=Gestao_Contrato!$B$3` |
| B3 | `ELÉTRICA` |
| E3 | `CLIMATIZAÇÃO` |
| H3 | `DETALHE POR SITUAÇÃO — ELÉTRICA` |
| B5 | `OPERAÇÃO TOTAL` |
| E5 | `OPERAÇÃO TOTAL` |
| H5 | `Situação` |
| I5 | `Horas lançadas` |
| J5 | `Coberto p/ Ferista` |
| K5 | `Abateu (R$)` |
| B6 | `Horas Total Trabalhadas` |
| C6 | `=IFERROR(INDEX(Analises!$AJ$3:$AJ$14,MATCH($C$2,Analises!$A$3:$A$14,0)),0)` |
| E6 | `Horas Total Trabalhadas` |
| F6 | `=IFERROR(INDEX(Analises!$AJ$19:$AJ$30,MATCH($C$2,Analises!$A$19:$A$30,0)),0)` |
| H6 | `Ausências C/ Justificativa` |
| I6 | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências C/ Justificativa")+SUMIFS(Lanc_HORAS!$M$2:$M$3873,Lanc_HORAS!$B$2:$B$ …` |
| J6 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências C/ Justificativa")` |
| K6 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências C/ Justificativa")` |
| B7 | `FÉRIAS` |
| E7 | `FÉRIAS` |
| H7 | `Ausências S/ Justificativa` |
| I7 | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências S/ Justificativa")+SUMIFS(Lanc_HORAS!$M$2:$M$3873,Lanc_HORAS!$B$2:$B$ …` |
| J7 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências S/ Justificativa")` |
| K7 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências S/ Justificativa")` |
| B8 | `Férias (Total)` |
| C8 | `=IFERROR(INDEX(Analises!$H$3:$H$14,MATCH($C$2,Analises!$A$3:$A$14,0)),0)` |
| E8 | `Férias (total)` |
| F8 | `=IFERROR(INDEX(Analises!$H$19:$H$30,MATCH($C$2,Analises!$A$19:$A$30,0)),0)` |
| H8 | `Atraso/Saída C/ Justificativa` |
| I8 | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Entrada Atrasada ou Saída Antecipada C/ Justificativa")+SUMIFS(Lanc_HORAS!$M$2: …` |
| J8 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Entrada Atrasada ou Saída Antecipada C/ Justificativa")` |
| K8 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Entrada Atrasada ou Saída Antecipada C/ Justificativa")` |
| B9 | `Férias Cobertas` |
| C9 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Férias")` |
| E9 | `Férias Cobertas` |
| F9 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Férias")` |
| H9 | `Atraso/Saída S/ Justificativa` |
| I9 | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Entrada Atrasada ou Saída Antecipada S/ Justificativa")+SUMIFS(Lanc_HORAS!$M$2: …` |
| J9 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Entrada Atrasada ou Saída Antecipada S/ Justificativa")` |
| K9 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Entrada Atrasada ou Saída Antecipada S/ Justificativa")` |
| B10 | `Férias Não Cobertas` |
| C10 | `=IFERROR(INDEX(Analises!$AS$3:$AS$14,MATCH($C$2,Analises!$A$3:$A$14,0)),0)` |
| E10 | `Férias Não Cobertas` |
| F10 | `=IFERROR(INDEX(Analises!$AS$19:$AS$30,MATCH($C$2,Analises!$A$19:$A$30,0)),0)` |
| H10 | `Ausência p/ Férias` |
| I10 | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Férias")+SUMIFS(Lanc_HORAS!$M$2:$M$3873,Lanc_HORAS!$B$2:$B$3 …` |
| J10 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Férias")` |
| K10 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Férias")` |
| B11 | `HORA EXTRA` |
| E11 | `HORA EXTRA` |
| H11 | `Falta de Contratação` |
| I11 | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Falta de Contratação")+SUMIFS(Lanc_HORAS!$M$2:$M$3873,Lanc_HOR …` |
| J11 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Falta de Contratação")` |
| K11 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Falta de Contratação")` |
| B12 | `HE PAGA — total (R$)` |
| C12 | `=IFERROR(INDEX(Analises!$AO$3:$AO$14,MATCH($C$2,Analises!$A$3:$A$14,0)),0)` |
| E12 | `HE total Acordado (R$)` |
| F12 | `=IFERROR(INDEX(Analises!$AO$19:$AO$30,MATCH($C$2,Analises!$A$19:$A$30,0)),0)` |
| H12 | `Falta de EPI` |
| I12 | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Falta de EPI")+SUMIFS(Lanc_HORAS!$M$2:$M$3873,Lanc_HORAS!$B$2: …` |
| J12 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Falta de EPI")` |
| K12 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Falta de EPI")` |
| B13 | `HE PAGA — total (h)` |
| C13 | `=IFERROR(INDEX(Analises!$AP$3:$AP$14,MATCH($C$2,Analises!$A$3:$A$14,0)),0)` |
| E13 | `HE total Acordado (H)` |
| F13 | `=IFERROR(INDEX(Analises!$AP$19:$AP$30,MATCH($C$2,Analises!$A$19:$A$30,0)),0)` |
| H13 | `Treinamento` |
| I13 | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Treinamento")+SUMIFS(Lanc_HORAS!$M$2:$M$3873,Lanc_HORAS!$B$2:$ …` |
| J13 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Treinamento")` |
| K13 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Treinamento")` |
| B14 | `HE pagas por EI (R$)` |
| C14 | `=IFERROR(INDEX(Analises!$L$3:$L$14,MATCH($C$2,Analises!$A$3:$A$14,0)),0)` |
| E14 | `HE pagas por EI (R$)` |
| F14 | `=IFERROR(INDEX(Analises!$L$19:$L$30,MATCH($C$2,Analises!$A$19:$A$30,0)),0)` |
| H14 | `Recesso` |
| I14 | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Recesso")+SUMIFS(Lanc_HORAS!$M$2:$M$3873,Lanc_HORAS!$B$2:$B$ …` |
| J14 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Recesso")` |
| K14 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Recesso")` |
| B15 | `HE pagas pela área (R$)` |
| C15 | `=IFERROR(INDEX(Analises!$M$3:$M$14,MATCH($C$2,Analises!$A$3:$A$14,0)),0)` |
| E15 | `HE pagas pela área (R$)` |
| F15 | `=IFERROR(INDEX(Analises!$M$19:$M$30,MATCH($C$2,Analises!$A$19:$A$30,0)),0)` |
| H15 | `Situação Atípica` |
| I15 | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Situação Atipica")+SUMIFS(Lanc_HORAS!$M$2:$M$3873,Lanc_HORAS!$B$2:$B$3873,$C$2, …` |
| J15 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Situação Atipica")` |
| K15 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Situação Atipica")` |
| B16 | `SITUAÇÕES` |
| E16 | `SITUAÇÕES` |
| H16 | `Remanejamento` |
| I16 | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Remanejamento")+SUMIFS(Lanc_HORAS!$M$2:$M$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lan …` |
| J16 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Remanejamento")` |
| K16 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Remanejamento")` |
| B17 | `Situação não acordada (pendente)` |
| C17 | `=IFERROR(INDEX(Analises!$AN$3:$AN$14,MATCH($C$2,Analises!$A$3:$A$14,0)),0)` |
| E17 | `Situação não acordada (pendente)` |
| F17 | `=IFERROR(INDEX(Analises!$AN$19:$AN$30,MATCH($C$2,Analises!$A$19:$A$30,0)),0)` |
| H17 | `Ausência parcial em dia trabalhado` |
| I17 | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Horas Trabalhadas")+SUMIFS(Lanc_HORAS!$M$2:$M$3873,Lanc_HORAS!$B$2:$B$3873,$C$2 …` |
| J17 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Horas Trabalhadas")` |
| K17 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Horas Trabalhadas")` |
| B18 | `Situação resolvida (acordo)` |
| C18 | `=IFERROR(INDEX(Analises!$AR$3:$AR$14,MATCH($C$2,Analises!$A$3:$A$14,0)),0)` |
| E18 | `Situação resolvida (acordo)` |
| F18 | `=IFERROR(INDEX(Analises!$AR$19:$AR$30,MATCH($C$2,Analises!$A$19:$A$30,0)),0)` |
| H18 | `Outras (recusa, integração, afastamento, desligamento)` |
| I18 | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Recusa de Atividade S/ Motivo Legítimo")+SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HO …` |
| J18 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Recusa de Atividade S/ Motivo Legítimo")+SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lan …` |
| K18 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Recusa de Atividade S/ Motivo Legítimo")+SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HO …` |
| B19 | `Horas Compensadas` |
| C19 | `=IFERROR(INDEX(Analises!$Q$3:$Q$14,MATCH($C$2,Analises!$A$3:$A$14,0)),0)` |
| E19 | `Horas Compensadas` |
| F19 | `=IFERROR(INDEX(Analises!$Q$19:$Q$30,MATCH($C$2,Analises!$A$19:$A$30,0)),0)` |
| H19 | `TOTAL` |
| I19 | `=SUM(I6:I18)` |
| J19 | `=SUM(J6:J18)` |
| K19 | `=SUM(K6:K18)` |
| B20 | `SALDO` |
| E20 | `SALDO` |
| H20 | `DETALHE POR SITUAÇÃO — CLIMATIZAÇÃO` |
| B21 | `Saldo do mês (R$)` |
| C21 | `=IFERROR(INDEX(Analises!$S$3:$S$14,MATCH($C$2,Analises!$A$3:$A$14,0)),0)` |
| E21 | `Saldo do mês (R$)` |
| F21 | `=IFERROR(INDEX(Analises!$S$19:$S$30,MATCH($C$2,Analises!$A$19:$A$30,0)),0)` |
| B22 | `Saldo acumulado do ano (R$)` |
| C22 | `=IFERROR(INDEX(Analises!$AB$3:$AB$14,MATCH($C$2,Analises!$A$3:$A$14,0)),0)` |
| E22 | `Saldo acumulado do ano (R$)` |
| F22 | `=IFERROR(INDEX(Analises!$AB$19:$AB$30,MATCH($C$2,Analises!$A$19:$A$30,0)),0)` |
| H22 | `Situação` |
| I22 | `Horas lançadas` |
| J22 | `Coberto p/ Ferista` |
| K22 | `Abateu (R$)` |
| B23 | `Qtd. de eventos de absenteísmo` |
| C23 | `=IFERROR(INDEX(Analises!$T$3:$T$14,MATCH($C$2,Analises!$A$3:$A$14,0)),0)` |
| E23 | `Qtd. de eventos de absenteísmo` |
| F23 | `=IFERROR(INDEX(Analises!$T$19:$T$30,MATCH($C$2,Analises!$A$19:$A$30,0)),0)` |
| H23 | `Ausências C/ Justificativa` |
| I23 | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências C/ Justificativa")+SUMIFS(Lanc_HORAS!$M$2:$M$3873,Lanc_HORAS!$B$2 …` |
| J23 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências C/ Justificativa")` |
| K23 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências C/ Justificativa")` |
| B24 | `FERISTA` |
| E24 | `FERISTA` |
| H24 | `Ausências S/ Justificativa` |
| I24 | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências S/ Justificativa")+SUMIFS(Lanc_HORAS!$M$2:$M$3873,Lanc_HORAS!$B$2 …` |
| J24 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências S/ Justificativa")` |
| K24 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências S/ Justificativa")` |
| B25 | `Horas Trabalhadas do Ferista` |
| C25 | `=IFERROR(INDEX(Analises!$AQ$3:$AQ$14,MATCH($C$2,Analises!$A$3:$A$14,0)),0)` |
| E25 | `Horas Trabalhadas do Ferista` |
| F25 | `=IFERROR(INDEX(Analises!$AQ$19:$AQ$30,MATCH($C$2,Analises!$A$19:$A$30,0)),0)` |
| H25 | `Atraso/Saída C/ Justificativa` |
| I25 | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Entrada Atrasada ou Saída Antecipada C/ Justificativa")+SUMIFS(Lanc_HORAS!$ …` |
| J25 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Entrada Atrasada ou Saída Antecipada C/ Justificativa")` |
| K25 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Entrada Atrasada ou Saída Antecipada C/ Justificativa")` |
| B26 | `Cobertura do Ferista em Ausências ` |
| C26 | `=IFERROR(INDEX(Analises!$AK$3:$AK$14,MATCH($C$2,Analises!$A$3:$A$14,0)),0)` |
| E26 | `Cobertura do Ferista em Ausências ` |
| F26 | `=IFERROR(INDEX(Analises!$AK$19:$AK$30,MATCH($C$2,Analises!$A$19:$A$30,0)),0)` |
| H26 | `Atraso/Saída S/ Justificativa` |
| I26 | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Entrada Atrasada ou Saída Antecipada S/ Justificativa")+SUMIFS(Lanc_HORAS!$ …` |
| J26 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Entrada Atrasada ou Saída Antecipada S/ Justificativa")` |
| K26 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Entrada Atrasada ou Saída Antecipada S/ Justificativa")` |
| B27 | `Horas Ferista Sem a Cobertura` |
| C27 | `=IFERROR(INDEX(Analises!$Y$3:$Y$14,MATCH($C$2,Analises!$A$3:$A$14,0)),0)` |
| E27 | `Horas Ferista Sem Cobertura` |
| F27 | `=IFERROR(INDEX(Analises!$Y$19:$Y$30,MATCH($C$2,Analises!$A$19:$A$30,0)),0)` |
| H27 | `Ausência p/ Férias` |
| I27 | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Férias")+SUMIFS(Lanc_HORAS!$M$2:$M$3873,Lanc_HORAS!$B$2: …` |
| J27 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Férias")` |
| K27 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Férias")` |
| B28 | `REMANEJAMENTO` |
| E28 | `REMANEJAMENTO` |
| H28 | `Falta de Contratação` |
| I28 | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Falta de Contratação")+SUMIFS(Lanc_HORAS!$M$2:$M$3873,Lanc …` |
| J28 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Falta de Contratação")` |
| K28 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Falta de Contratação")` |
| B29 | `Horas de Remanejamento` |
| C29 | `=IFERROR(INDEX(Analises!$AL$3:$AL$14,MATCH($C$2,Analises!$A$3:$A$14,0)),0)` |
| E29 | `Horas de Remanejamento` |
| F29 | `=IFERROR(INDEX(Analises!$AL$19:$AL$30,MATCH($C$2,Analises!$A$19:$A$30,0)),0)` |
| H29 | `Falta de EPI` |
| I29 | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Falta de EPI")+SUMIFS(Lanc_HORAS!$M$2:$M$3873,Lanc_HORAS!$ …` |
| J29 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Falta de EPI")` |
| K29 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Falta de EPI")` |
| B30 | `Horas de Remanejamento Coberto` |
| C30 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica",Lanc_HORAS!$J$2:$J$3873,"Remanejamento")` |
| E30 | `Horas de Remanejamento Coberto` |
| F30 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Remanejamento")` |
| H30 | `Treinamento` |
| I30 | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Treinamento")+SUMIFS(Lanc_HORAS!$M$2:$M$3873,Lanc_HORAS!$B …` |
| J30 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Treinamento")` |
| K30 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido Treinamento")` |
| B31 | `Horas Acordadas (acordo comercial)` |
| C31 | `=SUMIFS(Lanc_HORAS!$P$2:$P$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica")` |
| E31 | `Horas Acordadas (acordo comercial)` |
| F31 | `=SUMIFS(Lanc_HORAS!$P$2:$P$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização")` |
| H31 | `Recesso` |
| I31 | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Recesso")+SUMIFS(Lanc_HORAS!$M$2:$M$3873,Lanc_HORAS!$B$2 …` |
| J31 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Recesso")` |
| K31 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Ausências Devido a Recesso")` |
| B32 | `COBERTURA DO FERISTA` |
| E32 | `COBERTURA DO FERISTA` |
| H32 | `Situação Atípica` |
| I32 | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Situação Atipica")+SUMIFS(Lanc_HORAS!$M$2:$M$3873,Lanc_HORAS!$B$2:$B$3873,$ …` |
| J32 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Situação Atipica")` |
| K32 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Situação Atipica")` |
| B33 | `Horas cobertas INTEGRALMENTE` |
| C33 | `=SUMPRODUCT((Lanc_HORAS!$B$2:$B$3873=$C$2)*(Lanc_HORAS!$D$2:$D$3873="Elétrica")*(Lanc_HORAS!$AC$2:$AC$3873>0)*(Lanc_HORAS!$AC$2:$AC$3873>=IFERROR(Lanc_HORAS!$V$2:$V$3873*1,0))*Lanc_HORAS!$AC$2:$AC$387 …` |
| E33 | `Horas cobertas INTEGRALMENTE` |
| F33 | `=SUMPRODUCT((Lanc_HORAS!$B$2:$B$3873=$C$2)*(Lanc_HORAS!$D$2:$D$3873="Climatização")*(Lanc_HORAS!$AC$2:$AC$3873>0)*(Lanc_HORAS!$AC$2:$AC$3873>=IFERROR(Lanc_HORAS!$V$2:$V$3873*1,0))*Lanc_HORAS!$AC$2:$AC …` |
| H33 | `Remanejamento` |
| I33 | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Remanejamento")+SUMIFS(Lanc_HORAS!$M$2:$M$3873,Lanc_HORAS!$B$2:$B$3873,$C$2 …` |
| J33 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Remanejamento")` |
| K33 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Remanejamento")` |
| B34 | `Horas cobertas por DOAÇÃO residual` |
| C34 | `=SUMPRODUCT((Lanc_HORAS!$B$2:$B$3873=$C$2)*(Lanc_HORAS!$D$2:$D$3873="Elétrica")*(Lanc_HORAS!$AC$2:$AC$3873>0)*(Lanc_HORAS!$AC$2:$AC$3873<IFERROR(Lanc_HORAS!$V$2:$V$3873*1,0))*Lanc_HORAS!$AC$2:$AC$3873 …` |
| E34 | `Horas cobertas por DOAÇÃO residual` |
| F34 | `=SUMPRODUCT((Lanc_HORAS!$B$2:$B$3873=$C$2)*(Lanc_HORAS!$D$2:$D$3873="Climatização")*(Lanc_HORAS!$AC$2:$AC$3873>0)*(Lanc_HORAS!$AC$2:$AC$3873<IFERROR(Lanc_HORAS!$V$2:$V$3873*1,0))*Lanc_HORAS!$AC$2:$AC$ …` |
| H34 | `Ausência parcial em dia trabalhado` |
| I34 | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Horas Trabalhadas")+SUMIFS(Lanc_HORAS!$M$2:$M$3873,Lanc_HORAS!$B$2:$B$3873, …` |
| J34 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Horas Trabalhadas")` |
| K34 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Horas Trabalhadas")` |
| B35 | `Qtd. de casos de doação parcial` |
| C35 | `=SUMPRODUCT((Lanc_HORAS!$B$2:$B$3873=$C$2)*(Lanc_HORAS!$D$2:$D$3873="Elétrica")*(Lanc_HORAS!$AC$2:$AC$3873>0)*(Lanc_HORAS!$AC$2:$AC$3873<IFERROR(Lanc_HORAS!$V$2:$V$3873*1,0)))` |
| E35 | `Qtd. de casos de doação parcial` |
| F35 | `=SUMPRODUCT((Lanc_HORAS!$B$2:$B$3873=$C$2)*(Lanc_HORAS!$D$2:$D$3873="Climatização")*(Lanc_HORAS!$AC$2:$AC$3873>0)*(Lanc_HORAS!$AC$2:$AC$3873<IFERROR(Lanc_HORAS!$V$2:$V$3873*1,0)))` |
| H35 | `Outras (recusa, integração, afastamento, desligamento)` |
| I35 | `=SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Recusa de Atividade S/ Motivo Legítimo")+SUMIFS(Lanc_HORAS!$V$2:$V$3873,Lan …` |
| J35 | `=SUMIFS(Lanc_HORAS!$AC$2:$AC$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Recusa de Atividade S/ Motivo Legítimo")+SUMIFS(Lanc_HORAS!$AC$2:$AC$3873 …` |
| K35 | `=SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização",Lanc_HORAS!$J$2:$J$3873,"Recusa de Atividade S/ Motivo Legítimo")+SUMIFS(Lanc_HORAS!$W$2:$W$3873,Lan …` |
| B36 | `Horas subtraidas do absenteísmo` |
| C36 | `=SUMIFS(Lanc_HORAS!$AM$2:$AM$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Elétrica")` |
| E36 | `Horas subtraidas do absenteísmo` |
| F36 | `=SUMIFS(Lanc_HORAS!$AM$2:$AM$3873,Lanc_HORAS!$B$2:$B$3873,$C$2,Lanc_HORAS!$D$2:$D$3873,"Climatização")` |
| H36 | `TOTAL` |
| I36 | `=SUM(I23:I35)` |
| J36 | `=SUM(J23:J35)` |
| K36 | `=SUM(K23:K35)` |
| B37 | `COMPENSAÇÃO DE HORAS` |
| E37 | `COMPENSAÇÃO DE HORAS` |
| B38 | `Total Horas Compensadas` |
| C38 | `=IFERROR(INDEX(Analises!$AY$3:$AY$14,MATCH($C$2,Analises!$A$3:$A$14,0)),0)` |
| E38 | `Total Horas Compensadas` |
| F38 | `=IFERROR(INDEX(Analises!$AY$19:$AY$30,MATCH($C$2,Analises!$A$19:$A$30,0)),0)` |
| B39 | `Valor R$` |
| C39 | `=IFERROR(INDEX(Analises!$AZ$3:$AZ$14,MATCH($C$2,Analises!$A$3:$A$14,0)),0)` |
| E39 | `Valor R$` |
| F39 | `=IFERROR(INDEX(Analises!$AZ$19:$AZ$30,MATCH($C$2,Analises!$A$19:$A$30,0)),0)` |
| B40 | `HORAS PENDENTES` |
| E40 | `HORAS PENDENTES` |
| B41 | `Saldo Gerado no Mes Pendente` |
| C41 | `=IFERROR(INDEX(Analises!$BC$3:$BC$14,MATCH($C$2,Analises!$A$3:$A$14,0)),0)` |
| E41 | `Saldo Gerado no Mes Pendente` |
| F41 | `=IFERROR(INDEX(Analises!$BC$19:$BC$30,MATCH($C$2,Analises!$A$19:$A$30,0)),0)` |
| B42 | `SALDO PENDENTE acumulado` |
| C42 | `=IFERROR(INDEX(Analises!$BE$3:$BE$14,MATCH($C$2,Analises!$A$3:$A$14,0)),0)` |
| E42 | `SALDO PENDENTE acumulado` |
| F42 | `=IFERROR(INDEX(Analises!$BE$19:$BE$30,MATCH($C$2,Analises!$A$19:$A$30,0)),0)` |
| B43 | `Pendentes QUITADAS no mês` |
| C43 | `=IFERROR(INDEX(Analises!$BD$3:$BD$14,MATCH($C$2,Analises!$A$3:$A$14,0)),0)` |
| E43 | `Pendentes QUITADAS no Mês` |
| F43 | `=IFERROR(INDEX(Analises!$BD$19:$BD$30,MATCH($C$2,Analises!$A$19:$A$30,0)),0)` |
| B44 | `Compensação que sobrou p/ abater absenteísmo` |
| C44 | `=IFERROR(INDEX(Analises!$AW$3:$AW$14,MATCH($C$2,Analises!$A$3:$A$14,0)),0)` |
| E44 | `Compensação que sobrou p/ abater absenteísmo` |
| F44 | `=IFERROR(INDEX(Analises!$AW$19:$AW$30,MATCH($C$2,Analises!$A$19:$A$30,0)),0)` |
| B46 | `ABATIMENTO DE VALOR/FUNÇÃO` |
| E46 | `ABATIMENTO DE VALOR/FUNÇÃO` |
| B47 | `Valor-hora do ferista no mês (R$/h)` |
| C47 | `=IFERROR(LOOKUP(2,1/((Vigencias!$A$3:$A$149=IFERROR(INDEX(Cadastro!$A$3:$A$150,MATCH(Analises!$N$95,Cadastro!$B$3:$B$150,0)),""))*(Vigencias!$D$3:$D$149<=EOMONTH($C$2,0))),Vigencias!$F$3:$F$149),"—")` |
| E47 | `Valor-hora do ferista no mês (R$/h)` |
| F47 | `=IFERROR(LOOKUP(2,1/((Vigencias!$A$3:$A$149=IFERROR(INDEX(Cadastro!$A$3:$A$150,MATCH(Analises!$N$96,Cadastro!$B$3:$B$150,0)),""))*(Vigencias!$D$3:$D$149<=EOMONTH($C$2,0))),Vigencias!$F$3:$F$149),"—")` |
| B48 | `Valor abatido — diferença de função (R$)` |
| C48 | `=IFERROR(INDEX(Analises!$BH$3:$BH$14,MATCH($C$2,Analises!$A$3:$A$14,0)),0)` |
| E48 | `Valor abatido — diferença de função (R$)` |
| F48 | `=IFERROR(INDEX(Analises!$BH$19:$BH$30,MATCH($C$2,Analises!$A$19:$A$30,0)),0)` |

## Tabelas (blocos)

- **A**: REGRAS DOS STATUS — colunas: Tipo, Grupo, Conta no absenteísmo?
- **E**: FUNÇÕES E VALOR HH — colunas: Função / Custo, Valor HH (R$), Valor HE 50% (R$), Valor HE 100% (R$)
- **J**: PARÂMETROS — colunas: Parâmetro, Valor, Observação
- **N**: TURNOS — colunas: Estilo, Horario, HORAS TRABALHADAS
- **R**: LISTAS (HE / SIM-NÃO) — colunas: Tratamento HE, Sim/Não, Equipe
- **U**: VALOR DO CONTRATO POR DATA — colunas: Equipe, Início vigência, Valor mensal (R$), Observação
- **Z**: APROVAÇÃO — colunas: Aprovação, Status, Como aparece no relatório
- **AB**: RÓTULOS DO RELATÓRIO — colunas: Status, Como aparece no relatório, Data
- **AE**: CALENDÁRIO PCAM 2026 (Campo Limpo) — colunas: Data, Descrição, Tipo

### Parâmetros (J:L)

| Linha | Parâmetro | Valor | Observação |
|---|---|---|---|
| 3 | Ano de controle | 2026 | Usado nos resumos mensais |
| 4 | Saldo inicial do banco (R$) | 0 | Saldo inicial comum |
| 5 | Horas sábado a 100% | 04:00:00 | Regra usada na HE de sábado |
| 6 | Desconto almoço fim de semana (h) | 01:00:00 | Desconto automático quando sábado/domingo |
| 7 | Contrato mensal - Elétrica (R$) | 79852.8 | Equipe Elétrica Predial |
| 8 | Contrato mensal - Climatização (R$) | 55644 | Equipe Climatização |
| 9 | Contrato mensal - SPOT (R$) | - | Contrato à parte - Preencher quando houver SPOT |
| 10 | Desconto Horas de almoço Semanal | 00:30:00 | Ja embutido: o historico de ponto (coluna W) e a jornada do cadastro ja vem liquidos do almoco. Nao descontar de novo du … |
| 11 | Jornada ADM sexta-feira (h) | 8:00:00 | ADM = 9h seg-qui e 8h na sexta (sab/dom nao ha jornada) |
| 12 | Limite p/ desconto almoco fim de semana (h) | 6:00:00 | Na HE de sab/dom/feriado so desconta 1h de almoco acima deste total |
| 13 | (desativado) Inicio dos ALERTAS | None | Alertas agora valem para TODAS as datas. |
| 14 | (desativado) Inicio da COBERTURA DO FERISTA | None | Cobertura do ferista agora vale para TODAS as datas. |
| 15 | Doação mínima do residual do ferista (h) | 1:00:00 | Residual abaixo disso NÃO é doado nem cobre nada. |
| 16 | Arquivo da planilha de Pontos | G:\EI\EI - 3\EI - 3 - Facilities\17 - Fechamento dos Indicad … | Caminho completo do arquivo (OneDrive/SharePoint sincronizado). O botão "Atualizar Pontos" pede o arquivo na primeira ve … |
| 17 | Última atualização dos pontos | 19/10/2026 09:54 — base de pontos de 02/01/2026 a 30/09/2026 | Preenchido pela macro: data/hora e período da base de pontos lida. |
| 18 | Início do Abatimento de Valor/Função | 2026-01-01 00:00:00 | Ferista de cargo menor cobrindo técnico de cargo maior: desconta a diferença de valor-hora × horas cobertas a partir des … |

### Regras dos status (A:C) e rótulos do relatório (AB:AC)

| Status | Grupo | Conta no absenteísmo? |
|---|---|---|
| Ausências C/ Justificativa | Absenteísmo | Não |
| Ausências S/ Justificativa | Absenteísmo | Sim |
| Ausências Devido Falta de Contratação | Absenteísmo | Sim |
| Ausências Devido Treinamento | Absenteísmo | Não |
| Ausências Devido Falta de EPI | Absenteísmo | Sim |
| Ausências Devido a Férias | Absenteísmo | Sim |
| Entrada Atrasada ou Saída Antecipada C/ Justificativa | Absenteísmo | Não |
| Entrada Atrasada ou Saída Antecipada S/ Justificativa | Absenteísmo | Sim |
| Ausências Devido a Recesso | Absenteísmo | Sim |
| Reposição de Férias | Controle | Não |
| Reposição de Férias - HT | Controle | Não |
| Horas Compensadas | Controle | Não |
| Horas Trabalhadas | Controle | Não |
| Horas Transferidas | Controle | Não |
| Horas Recebidas | Controle | Não |
| Remanejamento | Controle | Sim |
| Banco de Horas | Controle | Não |
| Horas Extras | HE | Não |
| Situação Atípica | Absenteísmo | Sim |
| Recusa de Atividade S/ Motivo Legítimo | Absenteísmo | Sim |
| Recusa de Atividade C/ Motivo Legítimo | Absenteísmo | Não |
| Ausências Devido a Integração / Início de Atividades | Absenteísmo | Não |
| Ausências Devido a Afastamento por Acidente | Absenteísmo | Não |
| Ausências Devido a Desligamento / Pedido de Demissão | Absenteísmo | Sim |
| Adicional Não Solicitado | Controle | Não |

| Status | Como aparece no relatório |
|---|---|
| Ausências S/ Justificativa | Falta sem justificativa |
| Entrada Atrasada ou Saída Antecipada S/ Justificativa | Atraso/saída sem justificativa |
| Ausências C/ Justificativa | Falta justificada (atestado/comprovante) |
| Entrada Atrasada ou Saída Antecipada C/ Justificativa | Atraso/saída justificado (comprovante) |
| Ausências Devido a Férias | Férias |
| Ausências Devido Falta de Contratação | Posto vago (falta de contratação) |
| Remanejamento | Remanejamento |
| Ausências Devido Treinamento | Treinamento |
| Ausências Devido a Recesso | Recesso |
| Ausências Devido Falta de EPI | Falta de EPI |
| Horas Compensadas | Compensação (dia trabalhado) |
| Reposição de Férias | Reposição de férias |
| Reposição de Férias - HT | Reposição de férias |
| Banco de Horas | Banco de horas |
| Horas Transferidas | Horas transferidas |
| Horas Recebidas | Horas recebidas |
| Horas Extras | HE lançada na Lanc_HORAS |
| Situação Atípica | Situação atípica |
| Recusa de Atividade S/ Motivo Legítimo | Recusa de atividade (sem motivo legítimo) |
| Recusa de Atividade C/ Motivo Legítimo | Recusa de atividade (motivo legítimo) |
| Ausências Devido a Integração / Início de Atividades | Integração / início de atividades |
| Ausências Devido a Afastamento por Acidente | Afastamento por acidente |
| Ausências Devido a Desligamento / Pedido de Demissão | Desligamento / pedido de demissão |
| Adicional Não Solicitado | Adicional não solicitado |

## Nomes definidos

- `cfgArquivoPontos` → `Tabelas!$K$16`
- `cfgUltimaAtualizacao` → `Tabelas!$K$17`
- `lstAprovacao` → `Tabelas!$Z$3:$Z$5`
- `lstEquipes` → `Cadastro!$D$3:INDEX(Cadastro!$D$3:$D$150,MAX(1,COUNTA(Cadastro!$D$3:$D$150)))`
- `lstFuncionarios` → `Cadastro!$B$3:INDEX(Cadastro!$B$3:$B$150,MAX(1,COUNTA(Cadastro!$B$3:$B$150)))`
- `lstFuncoes` → `Tabelas!$E$3:INDEX(Tabelas!$E$3:$E$150,MAX(1,COUNTA(Tabelas!$E$3:$E$150)))`
- `lstMeses` → `Analises!$A$3:$A$14`
- `lstSimNao` → `Tabelas!$S$3:INDEX(Tabelas!$S$3:$S$10,MAX(1,COUNTA(Tabelas!$S$3:$S$10)))`
- `lstStatusOcorrencias` → `Tabelas!$A$3:INDEX(Tabelas!$A$3:$A$150,MAX(1,COUNTA(Tabelas!$A$3:$A$150)))`
- `lstTratamentoHE` → `Tabelas!$R$3:INDEX(Tabelas!$R$3:$R$10,MAX(1,COUNTA(Tabelas!$R$3:$R$10)))`
- `lstTurnos` → `Tabelas!$N$3:INDEX(Tabelas!$N$3:$N$20,MAX(1,COUNTA(Tabelas!$N$3:$N$20)))`
- `tabCalendario` → `Tabelas!$AE$3:$AG$62`

## Relatorio — âncoras

- B5: RESUMO DO MÊS
- B18: FERISTA DO MÊS
- B25: ELÉTRICA — OCORRÊNCIAS POR TIPO
- B46: AJUSTES, ACORDOS E CONTROLES
- B57: PARA CONFERIR
- B65: ELÉTRICA — DETALHE DIA A DIA
- B219: CLIMATIZAÇÃO — OCORRÊNCIAS POR TIPO
- B240: AJUSTES, ACORDOS E CONTROLES
- B251: PARA CONFERIR
- B259: CLIMATIZAÇÃO — DETALHE DIA A DIA
- B423: HORAS EXTRAS DO MÊS — CASO A CASO
- B487: COMPENSAÇÕES DO MÊS
- C61: TOTAL
- C62: ( − ) Compensações aprovadas + adicional não solicitado
- C63: ( + ) Abatimento de Valor/Função (ferista de cargo menor que o coberto)
- C255: TOTAL
- C256: ( − ) Compensações aprovadas + adicional não solicitado
- C257: ( + ) Abatimento de Valor/Função (ferista de cargo menor que o coberto)
- C415: PAGAS POR EI
- C416: PAGAS PELA ÁREA
- C417: HORAS COMPENSADAS
- C418: REMANEJADO P/ OUTRA ÁREA
- C419: SEM TRATAMENTO
- C420: TOTAL
- Área de impressão: 'Relatorio'!$B$1:$J$520; quebras manuais após as linhas: [64, 218, 412]
- Coluna N = 1 → a macro CriarRelatorio oculta a linha antes de gerar o PDF. Colunas L/M = índice da ocorrência no detalhe.