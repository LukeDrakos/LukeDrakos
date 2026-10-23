import sys
s = open(sys.argv[1], encoding='ascii').read()
def R(a, b, n=1):
    global s
    c = s.count(a)
    assert c == n, (c, a[:80])
    s = s.replace(a, b)

R("""'  Tudo age sobre a planilha ATIVA (ActiveWorkbook).
' ==========================================================================
""", """'  Tudo age sobre a planilha ATIVA (ActiveWorkbook).
'
'  DOIS LAYOUTS (detectados sozinhos em Alvo/DefinirLayout):
'   - padrao EI32 : Gestao_Contrato, Lanc_HORAS (cab. linha 1), Lanc_HE (cab. linha 1), Cadastro (cab. linha 2)
'   - padrao EI36 : Painel, Lancamentos (cab. linha 3), Lanc_HE (cab. linha 3), Cadastro (cab. linha 3)
'  As abas Relatorio, Conferencia_Ponto e Tabelas e os nomes cfgArquivoPontos, cfgUltimaAtualizacao e
'  tabCalendario existem nas duas.
' ==========================================================================
""")
R("""Private wbAlvo As Workbook   ' planilha de lancamento ativa (EI31, EI32 ou EI36)
""", """Private wbAlvo As Workbook   ' planilha de lancamento ativa (EI31, EI32 ou EI36)

' ---------- layout da planilha ativa (preenchido por DefinirLayout) ----------
Private gLay As String                      ' "EI32" ou "EI36"
Private ABA_GEST As String, ABA_LANC As String
Private L_HDR As Long, L_INI As Long        ' Lancamentos: linha do cabecalho / 1a linha de dados
Private L_NOME As String, L_ST As String, L_PONTO As String, L_DIV As String, L_OK As String, L_FIM As String
Private H_HDR As Long, H_INI As Long        ' Lanc_HE
Private H_NOME As String, H_PONTO As String, H_HORAS As String, H_FER As String, H_RECEB As String
Private H_DIV As String, H_OK As String, H_FIM As String
Private C_INI As Long, C_FIM As Long, C_EQ As Long   ' Cadastro: linhas e coluna da equipe

Private Sub DefinirLayout()
    If TemAba(wbAlvo, "Lancamentos") And TemAba(wbAlvo, "Painel") Then
        gLay = "EI36"
        ABA_GEST = "Painel": ABA_LANC = "Lancamentos"
        L_HDR = 3: L_INI = 4
        L_NOME = "B": L_ST = "C": L_PONTO = "D": L_DIV = "BC": L_OK = "BD": L_FIM = "BH"
        H_HDR = 3: H_INI = 4
        H_NOME = "B": H_PONTO = "K": H_HORAS = "F": H_FER = "E": H_RECEB = "L"
        H_DIV = "AE": H_OK = "AF": H_FIM = "AH"
        C_INI = 4: C_FIM = 200: C_EQ = 5
    Else
        gLay = "EI32"
        ABA_GEST = "Gestao_Contrato": ABA_LANC = "Lanc_HORAS"
        L_HDR = 1: L_INI = 2
        L_NOME = "E": L_ST = "J": L_PONTO = "Y": L_DIV = "AV": L_OK = "AW": L_FIM = "AW"
        H_HDR = 1: H_INI = 2
        H_NOME = "E": H_PONTO = "AB": H_HORAS = "P": H_FER = "N": H_RECEB = "AC"
        H_DIV = "AH": H_OK = "AI": H_FIM = "AI"
        C_INI = 3: C_FIM = 150: C_EQ = 4
    End If
End Sub
""")
R("""    ok = True
    For Each nm In Array("Gestao_Contrato", "Lanc_HORAS", "Lanc_HE", "Tabelas", "Cadastro")
        If Not TemAba(wbAlvo, CStr(nm)) Then ok = False
    Next nm
    If ok Then
        Alvo = True
        Exit Function
    End If""", """    ok = True
    For Each nm In Array("Gestao_Contrato", "Lanc_HORAS", "Lanc_HE", "Tabelas", "Cadastro")
        If Not TemAba(wbAlvo, CStr(nm)) Then ok = False
    Next nm
    If Not ok Then
        ok = True
        For Each nm In Array("Painel", "Lancamentos", "Lanc_HE", "Tabelas", "Cadastro")
            If Not TemAba(wbAlvo, CStr(nm)) Then ok = False
        Next nm
    End If
    If ok Then
        DefinirLayout
        Alvo = True
        Exit Function
    End If""")
R("""Private Sub GarantirAlvo()
    If wbAlvo Is Nothing Then Set wbAlvo = ActiveWorkbook
End Sub""", """Private Sub GarantirAlvo()
    If wbAlvo Is Nothing Then Set wbAlvo = ActiveWorkbook
    If gLay = "" Then DefinirLayout
End Sub""")
# ---- AtualizarPontos
R("""    Set wsL = wbAlvo.Worksheets("Lanc_HORAS")""", """    Set wsL = wbAlvo.Worksheets(ABA_LANC)""")
R("""    For lin = 3 To 150
        If CStr(wsCad.Cells(lin, 1).Value) <> "" Then""", """    For lin = C_INI To C_FIM
        If CStr(wsCad.Cells(lin, 1).Value) <> "" And IsNumeric(wsCad.Cells(lin, 1).Value) Then""")
R("""    For lin = 2 To ultL
        nm = CStr(wsL.Cells(lin, "E").Value)
        If CTem(dNomeRG, nm) And IsDate(wsL.Cells(lin, "A").Value) Then""", """    For lin = L_INI To ultL
        nm = CStr(wsL.Cells(lin, L_NOME).Value)
        If CTem(dNomeRG, nm) And IsDate(wsL.Cells(lin, "A").Value) Then""")
R("""    For lin = 2 To ultH
        nm = CStr(wsH.Cells(lin, "E").Value)
        If CTem(dNomeRG, nm) And IsDate(wsH.Cells(lin, "A").Value) Then
            k = Chave(dNomeRG(nm), CDbl(wsH.Cells(lin, "A").Value))""", """    For lin = H_INI To ultH
        nm = CStr(wsH.Cells(lin, H_NOME).Value)
        If CTem(dNomeRG, nm) And IsDate(wsH.Cells(lin, "A").Value) Then
            k = Chave(dNomeRG(nm), CDbl(wsH.Cells(lin, "A").Value))""")
R("""    Application.StatusBar = "Conferindo Lanc_HORAS..."
    ' ---------------- Lanc_HORAS ----------------
    For lin = 2 To ultL
        nm = CStr(wsL.Cells(lin, "E").Value)""", """    Application.StatusBar = "Conferindo " & ABA_LANC & "..."
    ' ---------------- Lanc_HORAS (EI32) / Lancamentos (EI36) ----------------
    For lin = L_INI To ultL
        nm = CStr(wsL.Cells(lin, L_NOME).Value)""")
R("""                st = CStr(wsL.Cells(lin, "J").Value)
                y = Trim$(CStr(wsL.Cells(lin, "Y").Value))""", """                st = CStr(wsL.Cells(lin, L_ST).Value)
                y = Trim$(CStr(wsL.Cells(lin, L_PONTO).Value))""")
R('''info.Add Array(d, "Lanc_HORAS", nm,''', '''info.Add Array(d, ABA_LANC, nm,''', 2)
R('''wsL.Cells(lin, "Y").Value''', '''wsL.Cells(lin, L_PONTO).Value''', 3)
R('''Marcar(wsL, "AV", lin,''', '''Marcar(wsL, L_DIV, lin,''', 7)
R("""    For lin = 2 To ultH
        nm = CStr(wsH.Cells(lin, "E").Value)
        If CTem(dNomeRG, nm) And IsDate(wsH.Cells(lin, "A").Value) Then
            d = CDbl(wsH.Cells(lin, "A").Value)""", """    For lin = H_INI To ultH
        nm = CStr(wsH.Cells(lin, H_NOME).Value)
        If CTem(dNomeRG, nm) And IsDate(wsH.Cells(lin, "A").Value) Then
            d = CDbl(wsH.Cells(lin, "A").Value)""")
R('''                y = Trim$(CStr(wsH.Cells(lin, "AB").Value))''', '''                y = Trim$(CStr(wsH.Cells(lin, H_PONTO).Value))''')
R('''                            wsH.Cells(lin, "AB").Value = "'" & p
                            wsH.Cells(lin, "P").Value = TotalDoPonto(p)
                            wsH.Cells(lin, "P").NumberFormat = "[h]:mm"''', '''                            wsH.Cells(lin, H_PONTO).Value = "'" & p
                            wsH.Cells(lin, H_HORAS).Value = TotalDoPonto(p)
                            wsH.Cells(lin, H_HORAS).NumberFormat = "[h]:mm"''')
R('''Marcar(wsH, "AH", lin,''', '''Marcar(wsH, H_DIV, lin,''', 4)
R('''wsH.Cells(lin, "N").Value''', '''wsH.Cells(lin, H_FER).Value''', 2)
R("""                    wsH.Cells(ultH, "A").Value = CDate(d)
                    wsH.Cells(ultH, "C").Value = "-"
                    wsH.Cells(ultH, "E").Value = nm
                    wsH.Cells(ultH, "M").Value = "-"
                    If CTem(dFolga, CStr(CLng(d))) Then
                        wsH.Cells(ultH, "N").Value = "Sim"
                    Else
                        wsH.Cells(ultH, "N").Value = Tx("N{a~}o")
                    End If
                    wsH.Cells(ultH, "P").Value = TotalDoPonto(p)
                    wsH.Cells(ultH, "P").NumberFormat = "[h]:mm"
                    wsH.Cells(ultH, "AB").Value = "'" & p
                    wsH.Cells(ultH, "AC").Value = "Sim"
                    wsH.Cells(ultH, "AH").Value =""", """                    wsH.Cells(ultH, "A").Value = CDate(d)
                    wsH.Cells(ultH, H_NOME).Value = nm
                    If gLay = "EI32" Then
                        wsH.Cells(ultH, "C").Value = "-"
                        wsH.Cells(ultH, "M").Value = "-"
                    End If
                    If CTem(dFolga, CStr(CLng(d))) Then
                        wsH.Cells(ultH, H_FER).Value = "Sim"
                    Else
                        wsH.Cells(ultH, H_FER).Value = Tx("N{a~}o")
                    End If
                    wsH.Cells(ultH, H_HORAS).Value = TotalDoPonto(p)
                    wsH.Cells(ultH, H_HORAS).NumberFormat = "[h]:mm"
                    wsH.Cells(ultH, H_PONTO).Value = "'" & p
                    wsH.Cells(ultH, H_RECEB).Value = "Sim"
                    wsH.Cells(ultH, H_DIV).Value =""")
R("""                    wsL.Cells(ultL, "A").Value = CDate(d)
                    wsL.Cells(ultL, "E").Value = nm
                    wsL.Cells(ultL, "J").Value = "Horas Trabalhadas"
                    wsL.Cells(ultL, "Y").Value = "'" & p
                    wsL.Cells(ultL, "AV").Value =""", """                    wsL.Cells(ultL, "A").Value = CDate(d)
                    wsL.Cells(ultL, L_NOME).Value = nm
                    wsL.Cells(ultL, L_ST).Value = "Horas Trabalhadas"
                    wsL.Cells(ultL, L_PONTO).Value = "'" & p
                    wsL.Cells(ultL, L_DIV).Value =""")
R("""        wsL.Range("A1:AW" & ultL).Sort Key1:=wsL.Range("A2"), Order1:=xlAscending, Key2:=wsL.Range("E2"), Order2:=xlAscending, Header:=xlYes
        wsH.Range("A1:AI" & ultH).Sort Key1:=wsH.Range("A2"), Order1:=xlAscending, Key2:=wsH.Range("E2"), Order2:=xlAscending, Header:=xlYes""",
"""        wsL.Range("A" & L_HDR & ":" & L_FIM & ultL).Sort Key1:=wsL.Range("A" & L_INI), Order1:=xlAscending, _
            Key2:=wsL.Range(L_NOME & L_INI), Order2:=xlAscending, Header:=xlYes
        wsH.Range("A" & H_HDR & ":" & H_FIM & ultH).Sort Key1:=wsH.Range("A" & H_INI), Order1:=xlAscending, _
            Key2:=wsH.Range(H_NOME & H_INI), Order2:=xlAscending, Header:=xlYes""")
# ---- MontarConferencia
R("""    For lin = 3 To 150
        If CStr(wsCad.Cells(lin, 2).Value) <> "" And Not CTem(eq, CStr(wsCad.Cells(lin, 2).Value)) Then CPor eq, CStr(wsCad.Cells(lin, 2).Value), CStr(wsCad.Cells(lin, 4).Value)
    Next lin""", """    For lin = C_INI To C_FIM
        If CStr(wsCad.Cells(lin, 2).Value) <> "" And Not CTem(eq, CStr(wsCad.Cells(lin, 2).Value)) Then CPor eq, CStr(wsCad.Cells(lin, 2).Value), CStr(wsCad.Cells(lin, C_EQ).Value)
    Next lin""")
R("""    abas = Array("Lanc_HORAS", "Lanc_HE"): cDiv = Array("AV", "AH"): cOk = Array("AW", "AI")
    For a = 0 To 1
        Set ws = wbAlvo.Worksheets(abas(a))
        ult = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        For lin = 2 To ult
            If CStr(ws.Range(cDiv(a) & lin).Value) <> "" And r <= 2000 Then
                EscreverConf wsC, r, ws.Cells(lin, "A").Value, CStr(abas(a)), CStr(ws.Cells(lin, "E").Value), eq, CStr(ws.Range(cDiv(a) & lin).Value), lin, CStr(cOk(a))""",
"""    abas = Array(ABA_LANC, "Lanc_HE"): cDiv = Array(L_DIV, H_DIV): cOk = Array(L_OK, H_OK)
    Dim ini As Variant, cNome As Variant
    ini = Array(L_INI, H_INI): cNome = Array(L_NOME, H_NOME)
    For a = 0 To 1
        Set ws = wbAlvo.Worksheets(abas(a))
        ult = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        For lin = CLng(ini(a)) To ult
            If CStr(ws.Range(cDiv(a) & lin).Value) <> "" And IsDate(ws.Cells(lin, "A").Value) And r <= 2000 Then
                EscreverConf wsC, r, ws.Cells(lin, "A").Value, CStr(abas(a)), CStr(ws.Cells(lin, CStr(cNome(a))).Value), eq, CStr(ws.Range(cDiv(a) & lin).Value), lin, CStr(cOk(a))""")
# ---- botoes
R("""    Set g = wbAlvo.Worksheets("Gestao_Contrato")
    g.Rows(3).RowHeight = 26
    g.Range("D3").Formula = "=HYPERLINK(""#'SLA'!A1"",""SLA"")"
    NovoBotao g, "btnAtualizarPontos", "G3", Tx("Atualizar Pontos"), "AtualizarPontos"
    NovoBotao g, "btnCriarRelatorio", "H3", Tx("Criar Relat{o'}rio (PDF)"), "CriarRelatorio\"""".replace('\\"', ''),
"""    Set g = wbAlvo.Worksheets(ABA_GEST)
    g.Rows(3).RowHeight = 26
    If gLay = "EI32" Then
        g.Range("D3").Formula = "=HYPERLINK(""#'SLA'!A1"",""SLA"")"
        NovoBotao g, "btnAtualizarPontos", "G3", Tx("Atualizar Pontos"), "AtualizarPontos"
        NovoBotao g, "btnCriarRelatorio", "H3", Tx("Criar Relat{o'}rio (PDF)"), "CriarRelatorio"
    Else
        NovoBotao g, "btnAtualizarPontos", "F3:G3", Tx("Atualizar Pontos"), "AtualizarPontos"
        NovoBotao g, "btnCriarRelatorio", "H3:I3", Tx("Criar Relat{o'}rio (PDF)"), "CriarRelatorio"
    End If""")
R("""Public Sub ConsertarLinks()
    Dim ws As Worksheet, hl As Hyperlink
    If Not Alvo() Then Exit Sub""", """Public Sub ConsertarLinks()
    Dim ws As Worksheet, hl As Hyperlink
    If wbAlvo Is Nothing Then
        If Not Alvo() Then Exit Sub
    End If""")
R("""    If dMax = 0 Then Err.Raise vbObjectError + 1, , Tx("A aba 'Base unificada' est{a'} vazia.")
""", """    If dMax = 0 Then Err.Raise vbObjectError + 1, , Tx("A aba 'Base unificada' est{a'} vazia.")
    ' cfgPontosDesde (opcional): datas anteriores nao sao conferidas nem criadas (protege meses fechados)
    If PontosDesde() > dMin Then dMin = PontosDesde()
""")
R("""Public Sub EscolherArquivoPontos()""", """' data a partir da qual o Atualizar Pontos age (nome cfgPontosDesde; 0 = sem limite)
Private Function PontosDesde() As Double
    Dim v As Variant
    PontosDesde = 0
    On Error Resume Next
    v = wbAlvo.Names("cfgPontosDesde").RefersToRange.Value
    If Err.Number <> 0 Then Err.Clear: Exit Function
    On Error GoTo 0
    If IsDate(v) Then PontosDesde = CDbl(CDate(v))
End Function

Public Sub EscolherArquivoPontos()""")
R("""Attribute VB_Name = "modEI32\"""".replace('\\"',''), """Attribute VB_Name = "modEI32\"""".replace('\\"',''))
s.encode('ascii')
open(sys.argv[2], 'w', encoding='ascii', newline='\r\n').write(s)
print('ok', len(s.splitlines()), 'linhas')
