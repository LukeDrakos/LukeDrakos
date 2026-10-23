' =====================================================================
'  INSTALAR - Controle de Horas TK (EI31 / EI32 / EI36)
'  De dois cliques neste arquivo. Ele faz tudo sozinho:
'   1. desbloqueia os arquivos desta pasta (Windows marca arquivos baixados)
'   2. marca esta pasta como Local Confiavel do Excel (macros liberadas)
'   3. cria Relatorios\01 - EI31, 02 - EI32, 03 - EI36 (se nao existirem)
'   4. ativa o suplemento "Controle de Horas TK.xlam" no Excel
'   5. abre cada "Lancamento horas - EIxx.xlsx", cria os botoes e salva
' =====================================================================
Option Explicit
Dim fso, sh, pasta, xlam, xl, wbTmp, ad, f, wb, n, ver, chave, existe, msg

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
pasta = fso.GetParentFolderName(WScript.ScriptFullName)
xlam = pasta & "\Controle de Horas TK.xlam"

If Not fso.FileExists(xlam) Then
    MsgBox "Nao achei o arquivo 'Controle de Horas TK.xlam' nesta pasta:" & vbCrLf & pasta, vbCritical, "Instalar Controle de Horas TK"
    WScript.Quit
End If

' Excel aberto atrapalha (arquivos ficam travados)
existe = False
On Error Resume Next
Set xl = GetObject(, "Excel.Application")
If Err.Number = 0 Then existe = True
Err.Clear
On Error GoTo 0
If existe Then
    If MsgBox("O Excel esta aberto. Salve o que estiver fazendo e FECHE o Excel." & vbCrLf & vbCrLf & _
              "Depois clique em OK para continuar (ou Cancelar para sair).", vbOKCancel + vbExclamation, "Instalar Controle de Horas TK") = vbCancel Then WScript.Quit
    On Error Resume Next
    Set xl = Nothing
    Set xl = GetObject(, "Excel.Application")
    If Err.Number = 0 Then
        MsgBox "O Excel ainda esta aberto. Feche e rode o instalador de novo.", vbCritical, "Instalar Controle de Horas TK"
        WScript.Quit
    End If
    Err.Clear
    On Error GoTo 0
End If

' 1. desbloquear arquivos baixados
sh.Run "powershell -NoProfile -ExecutionPolicy Bypass -Command ""Get-ChildItem -LiteralPath '" & Replace(pasta, "'", "''") & "' -Recurse | Unblock-File""", 0, True

' 2. local confiavel (Office 2016/2019/365 = 16.0; 2013 = 15.0)
For Each ver In Array("16.0", "15.0")
    chave = "HKCU\Software\Microsoft\Office\" & ver & "\Excel\Security\Trusted Locations\ControleHorasTK\"
    On Error Resume Next
    sh.RegWrite chave & "Path", pasta & "\", "REG_SZ"
    sh.RegWrite chave & "AllowSubfolders", 1, "REG_DWORD"
    sh.RegWrite chave & "Description", "Controle de Horas TK", "REG_SZ"
    Err.Clear
    On Error GoTo 0
Next

' 3. pastas dos relatorios
If Not fso.FolderExists(pasta & "\Relatorios") Then fso.CreateFolder pasta & "\Relatorios"
For Each f In Array("01 - EI31", "02 - EI32", "03 - EI36")
    If Not fso.FolderExists(pasta & "\Relatorios\" & f) Then fso.CreateFolder pasta & "\Relatorios\" & f
Next

' 4. ativar o suplemento
Set xl = CreateObject("Excel.Application")
xl.Visible = True
Set wbTmp = xl.Workbooks.Add
On Error Resume Next
Set ad = xl.AddIns.Add(xlam, False)
If Err.Number <> 0 Then
    MsgBox "Nao consegui registrar o suplemento: " & Err.Description, vbCritical, "Instalar Controle de Horas TK"
    WScript.Quit
End If
ad.Installed = False
ad.Installed = True
Err.Clear
On Error GoTo 0

' 5. botoes em cada planilha de lancamento desta pasta
n = 0
For Each f In fso.GetFolder(pasta).Files
    If LCase(fso.GetExtensionName(f.Name)) = "xlsx" And InStr(1, f.Name, "Lancamento horas", vbTextCompare) = 1 Then
        Set wb = xl.Workbooks.Open(f.Path)
        wb.Activate
        On Error Resume Next
        xl.Run "'Controle de Horas TK.xlam'!InstalarBotoes"
        If Err.Number = 0 Then
            wb.Save
            n = n + 1
        Else
            MsgBox "Nao consegui criar os botoes em " & f.Name & ": " & Err.Description, vbExclamation, "Instalar Controle de Horas TK"
        End If
        Err.Clear
        On Error GoTo 0
    End If
Next
wbTmp.Close False

msg = "Pronto!" & vbCrLf & vbCrLf & _
      "- Suplemento ativo (guia Suplementos > Controle de Horas TK)" & vbCrLf & _
      "- Botoes criados em " & n & " planilha(s)" & vbCrLf & _
      "- Relatorios vao para Relatorios\<equipe>\<ano>\" & vbCrLf & vbCrLf & _
      "Nos outros computadores, basta dar dois cliques neste INSTALAR de novo."
MsgBox msg, vbInformation, "Instalar Controle de Horas TK"
