Attribute VB_Name = "modEI32"
Option Explicit

' ==========================================================================
'  Controle de Lancamento de Horas - macros (mesmo modulo para EI31, EI32 e EI36)
'  AtualizarPontos      : le a planilha de Pontos, preenche o historico de ponto
'                         e marca em vermelho tudo que nao bate com o relogio
'  EscolherArquivoPontos: troca o arquivo de Pontos usado pela atualizacao
'  CriarRelatorio       : gera o PDF do mes (filtro B3 da Gestao_Contrato) e salva em
'                         Relatorios\<nn - EIxx>\<ano>\<mm - Mes> - EIxx.pdf (1 arquivo por mes)
'  InstalarBotoes       : cria os botoes na planilha ativa (rodar uma vez por planilha)
'
'  INSTALACAO DEFINITIVA: este modulo fica num SUPLEMENTO (Controle de Horas TK.xlam) salvo na
'  pasta compartilhada. Cada computador instala o suplemento uma unica vez. As planilhas EI31,
'  EI32 e EI36 continuam .xlsx (sem macro dentro); os botoes delas chamam o suplemento.
'  Tudo age sobre a planilha ATIVA (ActiveWorkbook).
'
'  DOIS LAYOUTS (detectados sozinhos em Alvo/DefinirLayout):
'   - padrao EI32 : Gestao_Contrato, Lanc_HORAS (cab. linha 1), Lanc_HE (cab. linha 1), Cadastro (cab. linha 2)
'   - padrao EI36 : Painel, Lancamentos (cab. linha 3), Lanc_HE (cab. linha 3), Cadastro (cab. linha 3)
'  As abas Relatorio, Conferencia_Ponto e Tabelas e os nomes cfgArquivoPontos, cfgUltimaAtualizacao e
'  tabCalendario existem nas duas.
' ==========================================================================

Private Const ABA_REL As String = "Relatorio"
Private Const ABA_CONF As String = "Conferencia_Ponto"
Private Const BASE_PONTOS As String = "Base unificada"
Private wbAlvo As Workbook   ' planilha de lancamento ativa (EI31, EI32 ou EI36)

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

' ---------- texto com acento (codigo ASCII) ----------
Private Function Tx(ByVal s As String) As String
    s = Replace(s, "{a~}", ChrW(227)): s = Replace(s, "{a'}", ChrW(225))
    s = Replace(s, "{e'}", ChrW(233)): s = Replace(s, "{e^}", ChrW(234))
    s = Replace(s, "{i'}", ChrW(237)): s = Replace(s, "{o'}", ChrW(243))
    s = Replace(s, "{o~}", ChrW(245)): s = Replace(s, "{o^}", ChrW(244))
    s = Replace(s, "{u'}", ChrW(250)): s = Replace(s, "{c,}", ChrW(231))
    s = Replace(s, "{A~}", ChrW(195)): s = Replace(s, "{E^}", ChrW(202))
    s = Replace(s, "{C,}", ChrW(199)): s = Replace(s, "{A'}", ChrW(193))
    s = Replace(s, "{O~}", ChrW(213)): s = Replace(s, "{E'}", ChrW(201))
    s = Replace(s, "{O'}", ChrW(211)): s = Replace(s, "{-}", ChrW(8212))
    Tx = s
End Function


' ---------- planilha alvo: a planilha de lancamento que esta ativa ----------
Private Function Alvo() As Boolean
    Dim ok As Boolean, nm As Variant
    Alvo = False
    Set wbAlvo = Nothing
    On Error Resume Next
    Set wbAlvo = ActiveWorkbook
    On Error GoTo 0
    If wbAlvo Is Nothing Then GoTo naoAchou
    ok = True
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
    End If
naoAchou:
    MsgBox Tx("Abra a planilha ""Lancamento horas - EIxx"" e deixe ela ativa antes de usar este bot{a~}o."), vbExclamation, "Controle de Horas"
End Function

Private Function TemAba(ByVal wb As Workbook, ByVal nome As String) As Boolean
    Dim ws As Object
    On Error Resume Next
    Set ws = wb.Worksheets(nome)
    TemAba = Not ws Is Nothing
    Err.Clear
End Function

Private Sub GarantirAlvo()
    If wbAlvo Is Nothing Then Set wbAlvo = ActiveWorkbook
    If gLay = "" Then DefinirLayout
End Sub

' ---------- menu na guia Suplementos (criado quando o suplemento abre) ----------
Public Sub Auto_Open()
    CriarMenu
End Sub

Public Sub Auto_Close()
    On Error Resume Next
    Application.CommandBars("Worksheet Menu Bar").Controls("Controle de Horas TK").Delete
End Sub

Public Sub CriarMenu()
    Dim barra As Object, pop As Object
    On Error Resume Next
    Application.CommandBars("Worksheet Menu Bar").Controls("Controle de Horas TK").Delete
    Set barra = Application.CommandBars("Worksheet Menu Bar")
    If barra Is Nothing Then Exit Sub
    Set pop = barra.Controls.Add(Type:=10, Temporary:=True)
    pop.Caption = "Controle de Horas TK"
    ItemMenu pop, Tx("Atualizar Pontos"), "AtualizarPontos", 1977
    ItemMenu pop, Tx("Criar Relat{o'}rio (PDF)"), "CriarRelatorio", 3355
    ItemMenu pop, Tx("Trocar arquivo de Pontos"), "EscolherArquivoPontos", 23
    ItemMenu pop, Tx("Instalar bot{o~}es nesta planilha"), "InstalarBotoes", 2950
End Sub

Private Sub ItemMenu(ByVal pop As Object, ByVal texto As String, ByVal macro As String, ByVal icone As Long)
    Dim b As Object
    Set b = pop.Controls.Add(Type:=1, Temporary:=True)
    b.Caption = texto
    b.OnAction = NomeMacro(macro)
    b.FaceId = icone
    b.Style = 3
End Sub

' nome qualificado da macro (funciona com o modulo no suplemento ou dentro da planilha)
Private Function NomeMacro(ByVal macro As String) As String
    NomeMacro = "'" & ThisWorkbook.Name & "'!" & macro
End Function

' ---------- dicionario simples (Collection) - funciona em qualquer Excel, sem bibliotecas ----------
Private Function CTem(ByVal c As Collection, ByVal k As String) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = c(k)
    CTem = (Err.Number = 0)
    Err.Clear
End Function

Private Sub CPor(ByVal c As Collection, ByVal k As String, ByVal v As Variant)
    If CTem(c, k) Then c.Remove k
    c.Add v, k
End Sub

Private Function NormRG(ByVal v As Variant) As String
    Dim s As String, i As Long, ch As String, outS As String
    s = CStr(v)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then outS = outS & ch
    Next i
    Do While Len(outS) > 1 And Left$(outS, 1) = "0"
        outS = Mid$(outS, 2)
    Loop
    NormRG = outS
End Function

Private Function MinutosDe(ByVal hhmm As String) As Long
    MinutosDe = CLng(Val(Left$(hhmm, 2))) * 60 + CLng(Val(Mid$(hhmm, 4, 2)))
End Function

' "07:46#17:20" -> "07:46#17:20 = 09:04"  (ultima - primeira - 30 min)
Private Function FmtPonto(ByVal marcas As String) As String
    Dim partes() As String, lista As String, n As Long, i As Long, dif As Long, p As String
    partes = Split(Trim$(marcas), "#")
    For i = LBound(partes) To UBound(partes)
        p = Trim$(partes(i))
        If Len(p) = 5 And Mid$(p, 3, 1) = ":" Then
            If lista <> "" Then lista = lista & "#"
            lista = lista & p
            n = n + 1
        End If
    Next i
    If n = 0 Then FmtPonto = "": Exit Function
    If n = 1 Then FmtPonto = lista: Exit Function
    partes = Split(lista, "#")
    dif = MinutosDe(partes(UBound(partes))) - MinutosDe(partes(0))
    If dif < 0 Then dif = dif + 1440
    dif = dif - 30
    If dif <= 0 Then
        FmtPonto = lista
    Else
        FmtPonto = lista & " = " & Format$(dif \ 60, "00") & ":" & Format$(dif Mod 60, "00")
    End If
End Function

Private Function TotalDoPonto(ByVal y As String) As Double
    Dim k As Long
    k = InStr(y, " = ")
    If k = 0 Then TotalDoPonto = -1: Exit Function
    TotalDoPonto = MinutosDe(Mid$(y, k + 3, 5)) / 1440#
End Function

Private Function SemEspaco(ByVal s As String) As String
    SemEspaco = Replace(Trim$(s), " ", "")
End Function

Private Function Chave(ByVal rg As String, ByVal d As Double) As String
    Chave = rg & "|" & CStr(CLng(Int(d)))
End Function

' acrescenta o aviso na coluna de divergencia (sem repetir o mesmo aviso)
Private Function Marcar(ByVal ws As Worksheet, ByVal colDiv As String, ByVal lin As Long, ByVal aviso As String) As Boolean
    Dim atual As String
    atual = CStr(ws.Range(colDiv & lin).Value)
    If InStr(1, atual, aviso, vbBinaryCompare) > 0 Then Marcar = False: Exit Function
    If atual = "" Then
        ws.Range(colDiv & lin).Value = aviso
    Else
        ws.Range(colDiv & lin).Value = atual & " | " & aviso
    End If
    Marcar = True
End Function

' ---------- pasta local da planilha (resolve OneDrive / SharePoint sincronizado) ----------
' Quando o arquivo esta numa pasta sincronizada, ThisWorkbook.Path devolve um endereco https.
' Esta funcao converte para o caminho do disco (C:\Users\...), usando o registro do OneDrive.
Public Function PastaLocal() As String
    Dim p As String, r As String
    GarantirAlvo
    p = wbAlvo.Path
    If p = "" Then PastaLocal = "": Exit Function
    If LCase$(Left$(p, 4)) <> "http" Then PastaLocal = p: Exit Function
    r = PastaPeloRegistro(p)
    If r = "" Then r = PastaPeloAmbiente(p)
    PastaLocal = r
End Function

Private Function DecodUrl(ByVal s As String) As String
    Dim i As Long, out As String, h As String
    i = 1
    Do While i <= Len(s)
        If Mid$(s, i, 1) = "%" And i + 2 <= Len(s) Then
            h = Mid$(s, i + 1, 2)
            out = out & Chr$(CLng("&H" & h))
            i = i + 3
        Else
            out = out & Mid$(s, i, 1)
            i = i + 1
        End If
    Loop
    DecodUrl = Replace(out, "/", Application.PathSeparator)
End Function

Private Function PastaExiste(ByVal p As String) As Boolean
    On Error Resume Next
    PastaExiste = False
    If p <> "" Then PastaExiste = ((GetAttr(p) And vbDirectory) = vbDirectory)
    If Err.Number <> 0 Then PastaExiste = False
    Err.Clear
End Function

' HKCU\Software\SyncEngines\Providers\OneDrive\<id> : UrlNamespace -> MountPoint
Private Function PastaPeloRegistro(ByVal url As String) As String
    Dim reg As Object, chaves As Variant, k As Variant, ns As String, mp As String, cand As String
    Const HKCU As Long = &H80000001
    Const CHAVE_OD As String = "Software\SyncEngines\Providers\OneDrive"
    PastaPeloRegistro = ""
    On Error Resume Next
    Set reg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    If reg Is Nothing Then Exit Function
    reg.EnumKey HKCU, CHAVE_OD, chaves
    If Not IsArray(chaves) Then Exit Function
    For Each k In chaves
        ns = "": mp = ""
        reg.GetStringValue HKCU, CHAVE_OD & "\" & k, "UrlNamespace", ns
        reg.GetStringValue HKCU, CHAVE_OD & "\" & k, "MountPoint", mp
        If ns <> "" And mp <> "" Then
            If Right$(ns, 1) <> "/" Then ns = ns & "/"
            If LCase$(Left$(url & "/", Len(ns))) = LCase$(ns) Then
                cand = mp & Application.PathSeparator & DecodUrl(Mid$(url, Len(ns) + 1))
                If Right$(cand, 1) = Application.PathSeparator Then cand = Left$(cand, Len(cand) - 1)
                If PastaExiste(cand) Then PastaPeloRegistro = cand: Exit Function
            End If
        End If
    Next k
End Function

' plano B: tenta encaixar o final do endereco dentro das pastas do OneDrive do usuario
Private Function PastaPeloAmbiente(ByVal url As String) As String
    Dim raizes(1 To 3) As String, i As Long, partes() As String, j As Long, resto As String, n As Long, cand As String
    raizes(1) = Environ$("OneDriveCommercial"): raizes(2) = Environ$("OneDrive"): raizes(3) = Environ$("USERPROFILE")
    partes = Split(Mid$(url, InStr(9, url, "/") + 1), "/")
    PastaPeloAmbiente = ""
    For i = 1 To 3
        If raizes(i) <> "" Then
            For j = 0 To UBound(partes)
                resto = ""
                For n = j To UBound(partes)
                    resto = resto & Application.PathSeparator & partes(n)
                Next n
                cand = raizes(i) & DecodUrl(resto)
                If PastaExiste(cand) Then PastaPeloAmbiente = cand: Exit Function
            Next j
        End If
    Next i
End Function

' codigo da equipe pelo nome do arquivo ("Lancamento horas - EI32.xlsm" -> "EI32")
Public Function CodigoEquipe() As String
    Dim nm As String, i As Long, d As String
    GarantirAlvo
    nm = UCase$(wbAlvo.Name)
    CodigoEquipe = "EI"
    For i = 1 To Len(nm) - 2
        If Mid$(nm, i, 2) = "EI" And IsNumeric(Mid$(nm, i + 2, 1)) Then
            d = ""
            i = i + 2
            Do While i <= Len(nm)
                If Not IsNumeric(Mid$(nm, i, 1)) Then Exit Do
                d = d & Mid$(nm, i, 1)
                i = i + 1
            Loop
            CodigoEquipe = "EI" & d
            Exit Function
        End If
    Next i
End Function

' pasta Relatorios\<nn - EIxx> ao lado da planilha; cria se nao existir
Public Function PastaRelatorios() As String
    Dim raiz As String, rel As String, f As String, cod As String, achou As String, sep As String
    GarantirAlvo
    sep = Application.PathSeparator
    raiz = PastaLocal()
    PastaRelatorios = ""
    If raiz = "" Then Exit Function
    cod = CodigoEquipe()
    rel = raiz & sep & "Relatorios"
    If Not PastaExiste(rel) Then MkDir rel
    achou = ""
    f = Dir(rel & sep & "*", vbDirectory)
    Do While f <> ""
        If f <> "." And f <> ".." Then
            If InStr(1, UCase$(f), cod, vbBinaryCompare) > 0 Then
                If Not IsNumeric(Mid$(UCase$(f), InStr(1, UCase$(f), cod) + Len(cod), 1)) Then
                    If PastaExiste(rel & sep & f) Then achou = f: Exit Do
                End If
            End If
        End If
        f = Dir()
    Loop
    If achou = "" Then
        achou = cod
        MkDir rel & sep & achou
    End If
    PastaRelatorios = rel & sep & achou
End Function

' procura "Pontos*.xls*" na mesma pasta da planilha (EI31, EI32 e EI36 ficam juntas com o Pontos 2026)
Private Function PontosNaPasta() As String
    Dim pasta As String, f As String
    GarantirAlvo
    pasta = PastaLocal()
    PontosNaPasta = ""
    If pasta = "" Then Exit Function
    f = Dir(pasta & Application.PathSeparator & "Pontos*.xls*")
    If f <> "" Then PontosNaPasta = pasta & Application.PathSeparator & f
End Function

' data a partir da qual o Atualizar Pontos age (nome cfgPontosDesde; 0 = sem limite)
Private Function PontosDesde() As Double
    Dim v As Variant
    PontosDesde = 0
    On Error Resume Next
    v = wbAlvo.Names("cfgPontosDesde").RefersToRange.Value
    If Err.Number <> 0 Then Err.Clear: Exit Function
    On Error GoTo 0
    If IsDate(v) Then PontosDesde = CDbl(CDate(v))
End Function

Public Sub EscolherArquivoPontos()
    Dim fd As Object
    If Not Alvo() Then Exit Sub
    Set fd = Application.FileDialog(3)   ' msoFileDialogFilePicker
    fd.Title = "Escolha a planilha de Pontos (Base unificada)"
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "Excel", "*.xlsx;*.xlsm;*.xls"
    If fd.Show = -1 Then
        wbAlvo.Names("cfgArquivoPontos").RefersToRange.Value = fd.SelectedItems(1)
        MsgBox Tx("Arquivo de Pontos gravado:") & vbCrLf & fd.SelectedItems(1), vbInformation, CodigoEquipe()
    End If
End Sub

Public Sub AtualizarPontos()
    Dim caminho As String, wbP As Workbook, wsB As Worksheet, dados As Variant
    Dim dP As Collection, dDia As Collection, dNomeRG As Collection, dRGNome As Collection, dFolga As Collection, kP As Collection
    Dim dLanc As Collection, dHE As Collection
    Dim wsL As Worksheet, wsH As Worksheet, wsC As Worksheet, wsCad As Worksheet, wsT As Worksheet
    Dim i As Long, lin As Long, ultL As Long, ultH As Long, cal As Range
    Dim rg As String, nm As String, st As String, y As String, p As String, k As String
    Dim d As Double, dMin As Double, dMax As Double, ausente As Boolean, trab As Boolean, completo As Boolean
    Dim nNovas As Long, nPreench As Long, nInfo As Long, nCriadas As Long
    Dim info As Collection, it As Variant, chv As Variant, partes() As String
    If Not Alvo() Then Exit Sub

    On Error GoTo Falhou
    caminho = CStr(wbAlvo.Names("cfgArquivoPontos").RefersToRange.Value)
    If caminho = "" Or Dir(caminho) = "" Then caminho = PontosNaPasta()
    If caminho <> "" Then wbAlvo.Names("cfgArquivoPontos").RefersToRange.Value = caminho
    If caminho = "" Or Dir(caminho) = "" Then
        EscolherArquivoPontos
        caminho = CStr(wbAlvo.Names("cfgArquivoPontos").RefersToRange.Value)
        If caminho = "" Or Dir(caminho) = "" Then Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Lendo a planilha de Pontos..."

    Set wbP = Workbooks.Open(Filename:=caminho, ReadOnly:=True, UpdateLinks:=0)
    Set wsB = wbP.Worksheets(BASE_PONTOS)
    dados = wsB.Range("A5:H" & wsB.Cells(wsB.Rows.Count, "A").End(xlUp).Row).Value
    wbP.Close SaveChanges:=False
    Set wbP = Nothing

    Set wsL = wbAlvo.Worksheets(ABA_LANC)
    Set wsH = wbAlvo.Worksheets("Lanc_HE")
    Set wsC = wbAlvo.Worksheets(ABA_CONF)
    Set wsCad = wbAlvo.Worksheets("Cadastro")
    Set wsT = wbAlvo.Worksheets("Tabelas")

    Set dNomeRG = New Collection
    Set dRGNome = New Collection
    For lin = C_INI To C_FIM
        If CStr(wsCad.Cells(lin, 1).Value) <> "" And IsNumeric(wsCad.Cells(lin, 1).Value) Then
            rg = NormRG(wsCad.Cells(lin, 1).Value)
            nm = CStr(wsCad.Cells(lin, 2).Value)
            If Not CTem(dNomeRG, nm) Then CPor dNomeRG, nm, rg
            If Not CTem(dRGNome, rg) Then CPor dRGNome, rg, nm
        End If
    Next lin

    Set dFolga = New Collection
    Set cal = wbAlvo.Names("tabCalendario").RefersToRange
    For lin = 1 To cal.Rows.Count
        If IsDate(cal.Cells(lin, 1).Value) And CStr(cal.Cells(lin, 3).Value) = "Folga" Then
            k = CStr(CLng(CDate(cal.Cells(lin, 1).Value)))
            If Not CTem(dFolga, k) Then CPor dFolga, k, True
        End If
    Next lin

    Set dP = New Collection
    Set kP = New Collection
    Set dDia = New Collection
    dMin = 0: dMax = 0
    For i = 1 To UBound(dados, 1)
        If CStr(dados(i, 1)) <> "" And IsDate(dados(i, 3)) Then
            d = CDbl(CDate(dados(i, 3)))
            rg = NormRG(dados(i, 2))
            If dMin = 0 Or d < dMin Then dMin = d
            If d > dMax Then dMax = d
            p = FmtPonto(CStr(dados(i, 6)))
            If p <> "" Then
                k = Chave(rg, d)
                If Not CTem(dP, k) Then kP.Add k
                CPor dP, k, p
                If CTem(dRGNome, rg) Then
                    k = CStr(CLng(d))
                    If CTem(dDia, k) Then
                        CPor dDia, k, dDia(k) + 1
                    Else
                        CPor dDia, k, 1
                    End If
                End If
            End If
        End If
    Next i
    If dMax = 0 Then Err.Raise vbObjectError + 1, , Tx("A aba 'Base unificada' est{a'} vazia.")
    ' cfgPontosDesde (opcional): datas anteriores nao sao conferidas nem criadas (protege meses fechados)
    If PontosDesde() > dMin Then dMin = PontosDesde()

    Set info = New Collection
    Set dLanc = New Collection
    Set dHE = New Collection
    ultL = wsL.Cells(wsL.Rows.Count, "A").End(xlUp).Row
    ultH = wsH.Cells(wsH.Rows.Count, "A").End(xlUp).Row
    For lin = L_INI To ultL
        nm = CStr(wsL.Cells(lin, L_NOME).Value)
        If CTem(dNomeRG, nm) And IsDate(wsL.Cells(lin, "A").Value) Then
            k = Chave(dNomeRG(nm), CDbl(wsL.Cells(lin, "A").Value))
            If Not CTem(dLanc, k) Then CPor dLanc, k, lin
        End If
    Next lin
    For lin = H_INI To ultH
        nm = CStr(wsH.Cells(lin, H_NOME).Value)
        If CTem(dNomeRG, nm) And IsDate(wsH.Cells(lin, "A").Value) Then
            k = Chave(dNomeRG(nm), CDbl(wsH.Cells(lin, "A").Value))
            If Not CTem(dHE, k) Then CPor dHE, k, lin
        End If
    Next lin

    Application.StatusBar = "Conferindo " & ABA_LANC & "..."
    ' ---------------- Lanc_HORAS (EI32) / Lancamentos (EI36) ----------------
    For lin = L_INI To ultL
        nm = CStr(wsL.Cells(lin, L_NOME).Value)
        If CTem(dNomeRG, nm) And IsDate(wsL.Cells(lin, "A").Value) And Left$(nm, 11) <> "Colaborador" Then
            d = CDbl(wsL.Cells(lin, "A").Value)
            If d >= dMin And d <= dMax Then
                rg = dNomeRG(nm)
                st = CStr(wsL.Cells(lin, L_ST).Value)
                y = Trim$(CStr(wsL.Cells(lin, L_PONTO).Value))
                If y = "-" Then y = ""
                ausente = (Left$(st, 8) = Tx("Aus{e^}ncia"))
                trab = (st = "Horas Trabalhadas" Or st = Tx("Adicional N{a~}o Solicitado"))
                k = Chave(rg, d)
                If Not CTem(dDia, CStr(CLng(d))) Then
                    ' a base de pontos nao tem nenhum registro deste dia
                    If trab And y = "" Then info.Add Array(d, ABA_LANC, nm, Tx("INFO {-} Base de pontos sem registros deste dia"), lin)
                ElseIf CTem(dP, k) Then
                    p = dP(k): completo = (InStr(p, " = ") > 0)
                    If y = "" Then
                        wsL.Cells(lin, L_PONTO).Value = "'" & p: nPreench = nPreench + 1
                        If ausente And completo Then
                            If Marcar(wsL, L_DIV, lin, Tx("AUS{E^}NCIA COM PONTO: lan{c,}ado como """ & st & """, mas o rel{o'}gio mostra presen{c,}a (" & p & ").")) Then nNovas = nNovas + 1
                        End If
                    ElseIf SemEspaco(y) <> SemEspaco(p) Then
                        If completo Then
                            wsL.Cells(lin, L_PONTO).Value = "'" & p
                            If Marcar(wsL, L_DIV, lin, Tx("PONTO CORRIGIDO: estava """ & y & """, rel{o'}gio mostra """ & p & """.")) Then nNovas = nNovas + 1
                        Else
                            info.Add Array(d, ABA_LANC, nm, Tx("INFO {-} Base de pontos incompleta neste dia (planilha: " & y & " {-} base: " & p & "). Mantido."), lin)
                        End If
                    ElseIf ausente And completo Then
                        If Marcar(wsL, L_DIV, lin, Tx("AUS{E^}NCIA COM PONTO: lan{c,}ado como """ & st & """, mas o rel{o'}gio mostra presen{c,}a (" & p & ").")) Then nNovas = nNovas + 1
                    End If
                Else
                    If trab Then
                        If y = "" Then
                            wsL.Cells(lin, L_PONTO).Value = "SEM REGISTRO DE PONTO = 00:00"
                            If Marcar(wsL, L_DIV, lin, Tx("SEM PONTO: lan{c,}ado como trabalhado e n{a~}o h{a'} marca{c,}{a~}o no rel{o'}gio. Horas zeradas {-} analisar.")) Then nNovas = nNovas + 1
                        ElseIf Left$(y, 12) <> "SEM REGISTRO" Then
                            If Marcar(wsL, L_DIV, lin, Tx("SEM PONTO NO REL{O'}GIO: o hist{o'}rico """ & y & """ n{a~}o existe na base de pontos {-} conferir.")) Then nNovas = nNovas + 1
                        End If
                    ElseIf Left$(st, 7) = "Entrada" Or st = "Horas Compensadas" Or st = "Remanejamento" Then
                        If Marcar(wsL, L_DIV, lin, Tx("SEM PONTO NO REL{O'}GIO: """ & st & """ lan{c,}ado sem marca{c,}{a~}o na base de pontos {-} conferir.")) Then nNovas = nNovas + 1
                    End If
                End If
                If CTem(dFolga, CStr(CLng(d))) Then
                    If (trab And Not CTem(dHE, k)) Or (ausente And st <> Tx("Aus{e^}ncias C/ Justificativa")) Then
                        If Marcar(wsL, L_DIV, lin, Tx("DIA DE FOLGA NO CALEND{A'}RIO PCAM: lan{c,}ado como """ & st & """ {-} conferir.")) Then nNovas = nNovas + 1
                    End If
                End If
            End If
        End If
    Next lin

    Application.StatusBar = "Conferindo Lanc_HE..."
    ' ---------------- Lanc_HE ----------------
    For lin = H_INI To ultH
        nm = CStr(wsH.Cells(lin, H_NOME).Value)
        If CTem(dNomeRG, nm) And IsDate(wsH.Cells(lin, "A").Value) Then
            d = CDbl(wsH.Cells(lin, "A").Value)
            If d >= dMin And d <= dMax Then
                k = Chave(dNomeRG(nm), d)
                y = Trim$(CStr(wsH.Cells(lin, H_PONTO).Value))
                If CTem(dP, k) Then
                    p = dP(k)
                    If InStr(p, " = ") > 0 Then
                        If SemEspaco(y) <> SemEspaco(p) Then
                            wsH.Cells(lin, H_PONTO).Value = "'" & p
                            wsH.Cells(lin, H_HORAS).Value = TotalDoPonto(p)
                            wsH.Cells(lin, H_HORAS).NumberFormat = "[h]:mm"
                            If y = "" Then y = "vazio"
                            If Marcar(wsH, H_DIV, lin, Tx("PONTO CORRIGIDO: estava """ & y & """, rel{o'}gio mostra """ & p & """. Horas da HE ajustadas.")) Then nNovas = nNovas + 1
                        End If
                    Else
                        info.Add Array(d, "Lanc_HE", nm, Tx("INFO {-} Base de pontos incompleta neste dia (base: " & p & "). Mantido."), lin)
                    End If
                ElseIf CTem(dDia, CStr(CLng(d))) Then
                    If Marcar(wsH, H_DIV, lin, Tx("SEM PONTO NO REL{O'}GIO: hora extra sem marca{c,}{a~}o na base de pontos {-} conferir.")) Then nNovas = nNovas + 1
                End If
                If CTem(dFolga, CStr(CLng(d))) And CStr(wsH.Cells(lin, H_FER).Value) <> "Sim" Then
                    If Marcar(wsH, H_DIV, lin, Tx("DIA DE FOLGA NO CALEND{A'}RIO PCAM: marcar ""{E'} feriado?"" = Sim (HE 100%) {-} conferir.")) Then nNovas = nNovas + 1
                End If
                If CStr(wsH.Cells(lin, H_FER).Value) = "Sim" And Not CTem(dFolga, CStr(CLng(d))) And Weekday(d, vbMonday) <= 5 Then
                    If Marcar(wsH, H_DIV, lin, Tx("MARCADO COMO FERIADO, mas o dia n{a~}o {e'} folga no calend{a'}rio PCAM {-} conferir.")) Then nNovas = nNovas + 1
                End If
            End If
        End If
    Next lin

    Application.StatusBar = "Procurando dias sem lancamento..."
    ' ---------------- dias com ponto e sem lancamento ----------------
    For Each chv In kP
        p = dP(chv)
        partes = Split(CStr(chv), "|")
        rg = partes(0): d = CDbl(partes(1))
        If CTem(dRGNome, rg) And InStr(p, " = ") > 0 And d >= dMin And d <= dMax Then
            If Not CTem(dLanc, chv) And Not CTem(dHE, chv) Then
                nm = dRGNome(rg)
                If Weekday(d, vbMonday) >= 6 Or CTem(dFolga, CStr(CLng(d))) Then
                    ultH = ultH + 1
                    wsH.Cells(ultH, "A").Value = CDate(d)
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
                    wsH.Cells(ultH, H_DIV).Value = Tx("HE CRIADA PELA ATUALIZA{C,}{A~}O DE PONTOS: havia ponto em fim de semana/folga e nenhum lan{c,}amento. Definir o tratamento da HE.")
                    CPor dHE, chv, ultH
                Else
                    ultL = ultL + 1
                    wsL.Cells(ultL, "A").Value = CDate(d)
                    wsL.Cells(ultL, L_NOME).Value = nm
                    wsL.Cells(ultL, L_ST).Value = "Horas Trabalhadas"
                    wsL.Cells(ultL, L_PONTO).Value = "'" & p
                    wsL.Cells(ultL, L_DIV).Value = Tx("LINHA CRIADA PELA ATUALIZA{C,}{A~}O DE PONTOS: havia ponto e nenhum lan{c,}amento neste dia. Conferir o status.")
                    CPor dLanc, chv, ultL
                End If
                nCriadas = nCriadas + 1: nNovas = nNovas + 1
            End If
        End If
    Next chv
    If nCriadas > 0 Then
        wsL.Range("A" & L_HDR & ":" & L_FIM & ultL).Sort Key1:=wsL.Range("A" & L_INI), Order1:=xlAscending, _
            Key2:=wsL.Range(L_NOME & L_INI), Order2:=xlAscending, Header:=xlYes
        wsH.Range("A" & H_HDR & ":" & H_FIM & ultH).Sort Key1:=wsH.Range("A" & H_INI), Order1:=xlAscending, _
            Key2:=wsH.Range(H_NOME & H_INI), Order2:=xlAscending, Header:=xlYes
    End If

    wbAlvo.Names("cfgUltimaAtualizacao").RefersToRange.Value = Format$(Now, "dd/mm/yyyy hh:nn") & Tx(" {-} base de pontos de ") & Format$(CDate(dMin), "dd/mm/yyyy") & " a " & Format$(CDate(dMax), "dd/mm/yyyy")

    Application.StatusBar = "Montando a Conferencia de Ponto..."
    MontarConferencia info
    nInfo = info.Count

    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    Application.ScreenUpdating = True
    wsC.Activate
    wsC.Range("B6").Select
    MsgBox Tx("Pontos atualizados.") & vbCrLf & vbCrLf & _
        Tx("Hist{o'}ricos preenchidos: ") & nPreench & vbCrLf & _
        Tx("Novas diverg{e^}ncias (linhas vermelhas): ") & nNovas & vbCrLf & _
        Tx("Linhas criadas (dia sem lan{c,}amento): ") & nCriadas & vbCrLf & _
        Tx("Avisos sobre a base de pontos (INFO): ") & nInfo & vbCrLf & vbCrLf & _
        Tx("Confira tudo na aba Confer{e^}ncia de Ponto e marque ""Conferido? = Sim"" em cada linha analisada."), vbInformation, CodigoEquipe()
    Exit Sub

Falhou:
    Dim msgErro As String
    msgErro = Err.Number & " - " & Err.Description
    On Error Resume Next
    If Not wbP Is Nothing Then wbP.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox Tx("N{a~}o foi poss{i'}vel atualizar os pontos.") & vbCrLf & "Detalhe: " & msgErro, vbExclamation, CodigoEquipe()
End Sub

' lista todas as linhas com divergencia + avisos INFO desta atualizacao
Private Sub MontarConferencia(ByVal info As Collection)
    Dim wsC As Worksheet, wsL As Worksheet, wsH As Worksheet, wsCad As Worksheet
    Dim lin As Long, r As Long, abas As Variant, cDiv As Variant, cOk As Variant, a As Long
    Dim ws As Worksheet, ult As Long, eq As Collection, it As Variant
    Set wsC = wbAlvo.Worksheets(ABA_CONF)
    Set wsCad = wbAlvo.Worksheets("Cadastro")
    Set eq = New Collection
    For lin = C_INI To C_FIM
        If CStr(wsCad.Cells(lin, 2).Value) <> "" And Not CTem(eq, CStr(wsCad.Cells(lin, 2).Value)) Then CPor eq, CStr(wsCad.Cells(lin, 2).Value), CStr(wsCad.Cells(lin, C_EQ).Value)
    Next lin
    wsC.Range("B6:J2000").UnMerge
    wsC.Range("B6:J2000").ClearContents
    On Error Resume Next
    wsC.Range("B6:J2000").Hyperlinks.Delete
    On Error GoTo 0
    r = 6
    abas = Array(ABA_LANC, "Lanc_HE"): cDiv = Array(L_DIV, H_DIV): cOk = Array(L_OK, H_OK)
    Dim ini As Variant, cNome As Variant
    ini = Array(L_INI, H_INI): cNome = Array(L_NOME, H_NOME)
    For a = 0 To 1
        Set ws = wbAlvo.Worksheets(abas(a))
        ult = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        For lin = CLng(ini(a)) To ult
            If CStr(ws.Range(cDiv(a) & lin).Value) <> "" And IsDate(ws.Cells(lin, "A").Value) And r <= 2000 Then
                EscreverConf wsC, r, ws.Cells(lin, "A").Value, CStr(abas(a)), CStr(ws.Cells(lin, CStr(cNome(a))).Value), eq, CStr(ws.Range(cDiv(a) & lin).Value), lin, CStr(cOk(a))
                r = r + 1
            End If
        Next lin
    Next a
    For Each it In info
        If r <= 2000 Then
            EscreverConf wsC, r, it(0), CStr(it(1)), CStr(it(2)), eq, CStr(it(3)), CLng(it(4)), ""
            r = r + 1
        End If
    Next it
End Sub

Private Sub EscreverConf(ByVal wsC As Worksheet, ByVal r As Long, ByVal d As Variant, ByVal aba As String, ByVal nm As String, ByVal eq As Collection, ByVal txt As String, ByVal lin As Long, ByVal colOk As String)
    wsC.Cells(r, "B").Formula = "=HYPERLINK(""#'" & aba & "'!A" & lin & """,""" & Format$(CDate(d), "dd/mm/yyyy") & """)"
    wsC.Cells(r, "B").Font.Underline = True
    wsC.Cells(r, "C").Value = aba
    wsC.Cells(r, "D").Value = nm
    If CTem(eq, nm) Then wsC.Cells(r, "E").Value = eq(nm)
    wsC.Range("F" & r & ":H" & r).Merge
    wsC.Cells(r, "F").Value = txt
    wsC.Cells(r, "F").WrapText = True
    If colOk = "" Then
        wsC.Cells(r, "I").Value = "Nada alterado"
        wsC.Cells(r, "J").Value = ChrW(8212)
    Else
        wsC.Cells(r, "I").Value = "Linha marcada"
        wsC.Cells(r, "J").Formula = "=IF('" & aba & "'!" & colOk & lin & "=""Sim"",""Sim"","""")"
    End If
    With wsC.Range("B" & r & ":J" & r)
        .Font.Name = "Arial": .Font.Size = 9: .Font.Color = RGB(77, 77, 77)
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(217, 217, 217)
    End With
    wsC.Cells(r, "F").Font.Size = 8: wsC.Cells(r, "F").Font.Bold = True
    wsC.Cells(r, "J").HorizontalAlignment = xlCenter: wsC.Cells(r, "J").Font.Bold = True
    wsC.Rows(r).RowHeight = 30
End Sub

' ---------------- Relatorio em PDF ----------------
' ---------- nome do PDF: Relatorios\<equipe>\<ano>\<mm - Mes> - EIxx.pdf (um arquivo por mes) ----------
Private Function NomeMes(ByVal m As Integer) As String
    NomeMes = Choose(m, "Janeiro", "Fevereiro", Tx("Mar{c,}o"), "Abril", "Maio", "Junho", "Julho", "Agosto", _
                     "Setembro", "Outubro", "Novembro", "Dezembro")
End Function

Private Function NomeArquivoRelatorio(ByVal pasta As String, ByVal mesTxt As String) As String
    Dim ano As String, mm As String, sep As String, pastaAno As String, arq As String, antigo As String
    NomeArquivoRelatorio = ""
    sep = Application.PathSeparator
    ano = Left$(mesTxt, 4)
    mm = Right$(mesTxt, 2)
    pastaAno = pasta & sep & ano
    If Not PastaExiste(pastaAno) Then MkDir pastaAno
    arq = pastaAno & sep & mm & " - " & NomeMes(CInt(mm)) & " - " & CodigoEquipe() & ".pdf"
    ' mesmo mes ja emitido: apaga para substituir (se estiver aberto, avisa)
    If Dir(arq) <> "" Then
        On Error Resume Next
        Kill arq
        If Err.Number <> 0 Or Dir(arq) <> "" Then
            Err.Clear
            On Error GoTo 0
            MsgBox Tx("O relat{o'}rio de ") & LCase$(NomeMes(CInt(mm))) & "/" & ano & Tx(" est{a'} aberto." & vbCrLf & _
                   "Feche o PDF e clique em Criar Relat{o'}rio de novo."), vbExclamation, CodigoEquipe()
            Exit Function
        End If
        On Error GoTo 0
    End If
    ' remove o arquivo no formato antigo (EIxx_Relatorio_aaaa-mm.pdf) para nao duplicar o mes
    antigo = pasta & sep & CodigoEquipe() & "_Relatorio_" & mesTxt & ".pdf"
    On Error Resume Next
    If Dir(antigo) <> "" Then Kill antigo
    Err.Clear
    On Error GoTo 0
    NomeArquivoRelatorio = arq
End Function

Public Sub CriarRelatorio()
    Dim wsRel As Worksheet, ultima As Long, lin As Long
    Dim pasta As String, arquivo As String, mesTxt As String
    If Not Alvo() Then Exit Sub
    On Error GoTo Falhou
    Set wsRel = wbAlvo.Worksheets(ABA_REL)
    Application.ScreenUpdating = False
    Application.CalculateFull
    ultima = CLng(wsRel.Range("L1").Value)
    mesTxt = CStr(wsRel.Range("L2").Value)
    wsRel.Rows("1:" & ultima).Hidden = False
    For lin = 1 To ultima
        If CStr(wsRel.Cells(lin, "N").Value) = "1" Then wsRel.Rows(lin).Hidden = True
    Next lin
    pasta = ""
    On Error Resume Next
    pasta = PastaRelatorios()
    On Error GoTo Falhou
    If pasta = "" Then
        pasta = Environ$("USERPROFILE") & "\Downloads"
        MsgBox Tx("N{a~}o achei a pasta Relatorios ao lado da planilha. O PDF vai para Downloads."), vbExclamation, CodigoEquipe()
    End If
    arquivo = NomeArquivoRelatorio(pasta, mesTxt)
    If arquivo = "" Then GoTo Restaurar
    wsRel.ExportAsFixedFormat Type:=xlTypePDF, Filename:=arquivo, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    wsRel.Rows("1:" & ultima).Hidden = False
    Application.ScreenUpdating = True
    MsgBox Tx("Relat{o'}rio salvo (substitui o anterior do mesmo m{e^}s):") & vbCrLf & arquivo, vbInformation, CodigoEquipe()
    Exit Sub
Restaurar:
    wsRel.Rows("1:" & ultima).Hidden = False
    Application.ScreenUpdating = True
    Exit Sub
Falhou:
    Dim msgErroRel As String
    msgErroRel = Err.Number & " - " & Err.Description
    On Error Resume Next
    If Not wsRel Is Nothing Then wsRel.Rows("1:" & ultima).Hidden = False
    Application.ScreenUpdating = True
    MsgBox Tx("N{a~}o foi poss{i'}vel criar o relat{o'}rio.") & vbCrLf & "Detalhe: " & msgErroRel, vbExclamation, CodigoEquipe()
End Sub

' ---------------- botoes ----------------
Public Sub InstalarBotoes()
    Dim g As Worksheet
    If Not Alvo() Then Exit Sub
    LimparBotoesAntigos
    Set g = wbAlvo.Worksheets(ABA_GEST)
    g.Rows(3).RowHeight = 26
    If gLay = "EI32" Then
        g.Range("D3").Formula = "=HYPERLINK(""#'SLA'!A1"",""SLA"")"
        NovoBotao g, "btnAtualizarPontos", "G3", Tx("Atualizar Pontos"), "AtualizarPontos"
        NovoBotao g, "btnCriarRelatorio", "H3", Tx("Criar Relat{o'}rio (PDF)"), "CriarRelatorio"
    Else
        NovoBotao g, "btnAtualizarPontos", "F3:G3", Tx("Atualizar Pontos"), "AtualizarPontos"
        NovoBotao g, "btnCriarRelatorio", "H3:I3", Tx("Criar Relat{o'}rio (PDF)"), "CriarRelatorio"
    End If
    NovoBotao wbAlvo.Worksheets(ABA_REL), "btnCriarRelatorio", "I1:J1", Tx("Criar Relat{o'}rio (PDF)"), "CriarRelatorio"
    NovoBotao wbAlvo.Worksheets(ABA_CONF), "btnAtualizarPontos", "G1:H1", Tx("Atualizar Pontos"), "AtualizarPontos"
    NovoBotao wbAlvo.Worksheets(ABA_CONF), "btnTrocarArquivo", "I1:J1", Tx("Trocar arquivo de Pontos"), "EscolherArquivoPontos"
    ConsertarLinks
    MsgBox Tx("Bot{o~}es criados e links corrigidos." & vbCrLf & "Salve a planilha normalmente (Ctrl + S). Ela pode ficar .xlsx: as macros ficam no suplemento Controle de Horas TK."), vbInformation, CodigoEquipe()
End Sub

' remove botoes de versoes anteriores (qualquer forma ligada as macros deste modulo)
Private Sub LimparBotoesAntigos()
    Dim ws As Worksheet, i As Long, acao As String
    For Each ws In wbAlvo.Worksheets
        For i = ws.Shapes.Count To 1 Step -1
            acao = ""
            On Error Resume Next
            acao = ws.Shapes(i).OnAction
            On Error GoTo 0
            If InStr(1, acao, "AtualizarPontos", vbTextCompare) > 0 Or InStr(1, acao, "CriarRelatorio", vbTextCompare) > 0 _
               Or InStr(1, acao, "EscolherArquivoPontos", vbTextCompare) > 0 Or Left$(ws.Shapes(i).Name, 3) = "btn" Then
                ws.Shapes(i).Delete
            End If
        Next i
    Next ws
End Sub

' links internos (menu da Gestao, titulos "voltar a Gestao") apontando para a propria planilha
Public Sub ConsertarLinks()
    Dim ws As Worksheet, hl As Hyperlink
    If wbAlvo Is Nothing Then
        If Not Alvo() Then Exit Sub
    End If
    For Each ws In wbAlvo.Worksheets
        For Each hl In ws.Hyperlinks
            If hl.SubAddress <> "" And hl.Address <> "" Then hl.Address = ""
        Next hl
    Next ws
End Sub

Private Sub NovoBotao(ByVal ws As Worksheet, ByVal nome As String, ByVal faixa As String, ByVal texto As String, ByVal acao As String)
    Dim shp As Shape, rg As Range
    On Error Resume Next
    ws.Shapes(nome).Delete
    On Error GoTo 0
    Dim larg As Single
    Set rg = ws.Range(faixa)
    larg = rg.Width - 6
    If larg > 180 Then larg = 180
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, rg.Left + rg.Width - larg - 3, rg.Top + 2, larg, rg.Height - 4)
    shp.Name = nome
    shp.Fill.ForeColor.RGB = RGB(31, 56, 100)
    shp.Line.Visible = msoFalse
    With shp.TextFrame2
        .TextRange.Text = texto
        .TextRange.Font.Size = 10
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Name = "Arial"
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
    End With
    shp.Placement = xlMove
    On Error Resume Next
    shp.DrawingObject.PrintObject = False
    On Error GoTo 0
    shp.OnAction = NomeMacro(acao)
End Sub
