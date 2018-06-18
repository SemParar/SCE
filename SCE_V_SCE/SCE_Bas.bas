Attribute VB_Name = "SCE_BAS"


Option Explicit

Public dbApp As ADODB.Connection
Public rsGeral As New Recordset
Public rsctl As New Recordset
Public rs As New Recordset
Public frase As String
Public gsPathIniFile As String
Public gsPathLogListas As String
Public gsPath As String
Public gsPathZ As String
Public gsPathZL As String
Public gsTrnSomenteStatCatPositivo As String
Public gsSepararManuaisSemParar As String
Public bEnvioTravado As Boolean
Public bGeracaoTravada As Boolean

'VARIAVEIS AUTOEXPRESSO
Public gsGerarTrnAE  As String
Public gsSepararManuaisAE As String
Public gsAcessoPadraoAE As String
Public gsTagPadraoAE As String

Public gsPath_REL As String
Public gsPath_SCEExtra As String
Public gsPath_CD As String
Public gsPath_LN As String
Public gsPath_TRT As String
Public gsPath_OUTROS As String
Public gsPath_DB As String
Public gsPath_DBCAD As String
Public gsPath_DS As String
Public gsPWD As String
Public gsuid As String
Public gssceln As String
Public gsscetg As String
Public gsscetr As String
Public gsPath_msgbox As String
Public gsPermiteManuais As String

Public gsPath_CGMP_Files As String
Public gsPath_CGMP_Files_TRN As String
Public gsPath_CGMP_Files_NEL As String
Public gsPath_CGMP_Files_LNT As String
Public gsPath_CGMP_Files_TAG As String
Public gsPath_CGMP_Files_TGT As String
Public gsPath_CGMP_Files_TRF As String
Public gsPath_CGMP_Files_TRT As String
Public gsPath_CGMP_Files_MSG As String


Public gsPath_CGMPEnvia As String
Public gsPath_CGMPRecebe As String
Public gsSourceFile As String
Public gsDestFile As String

Public gsPath_SCE_Files As String
Public gbTransferStatus As Boolean
Public giTransfere As Long
Public gsNextTrn As Variant
Public gsNextTrnDelta As String
Public gsNextClear As Variant
Public gsNextTrnNr As String
Public gbTRNOK As Boolean
Public gbTicket As String
Public gbZeradas As Variant
Public giDebug As Variant
Public gsTempoBusca As Integer

Public gsUser As String
Public gsPassword As String
Public gsUserGroup As String
Public gsstrErro_Posicao As String

Public gsEst_Nome As String
Public gsListasFaltantes As String
Public gsEst_Codigo As String
Public gsEst_EntradaM As String
Public gsEst_SaidaM As String
Public gsEst_Horario As String
Public gsEst_Logo As String
Public giRespMsg As Integer

Public gsdata As String
Public gstag As String
Public gsacao As String
Public gsDelimitador As String
Public gsCodNelaLivre As String
Public gsMinutosTicketCancelados As String
Public GSSequencia As Boolean
Public gsListas As String
Public gsfile As String
Public gsseqfile As String

Public gbytNivel_Acesso_Usuario As Byte
Public Const gintNIVEL_ADMINISTRADOR = 1
Public Const gintNIVEL_USUARIO = 2

               
Public Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type

Public Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Long
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type

Public Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOrfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        dwReserved As Long
End Type

Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
     
Public Type SHELLEXECUTEINFO
        cbSize As Long
        fMask As Long
        hwnd As Long
        lpVerb As String
        lpFile As String
        lpParameters As String
        lpDirectory As String
        nShow As Long
        hInstApp As Long
        '  Optional fields
        lpIDList As Long
        lpClass As String
        hkeyClass As Long
        dwHotKey As Long
        hIcon As Long
        hProcess As Long
End Type

Type LCHeader
    regtipo As String * 2
    seq     As String * 6
    dtger   As String * 8
    hrger   As String * 6
    tag     As String * 12
    f1      As String * 3
    fcrlf   As String * 2
End Type


'---- tipos da aplicação TRATATG
Type CDHeader
    regtipo As String * 2
    seq     As String * 5
    dtger   As String * 8
    hrger   As String * 6
    tag     As String * 13
    f1      As String * 3
    fcrlf   As String * 2
End Type

Type CDdet
    regtipo As String * 1
    pais    As String * 4
    tag     As String * 15
    placa   As String * 7
    f1      As String * 2
    oper    As String * 1
    f2      As String * 2
    f3      As String * 2
    f4      As String * 3
    fcrlf   As String * 2
End Type

' Tipo de Estrutura de LN
'LN02352200404140500040005500034422
'LT0387020060529051453091996

Type LNHeader
    regtipo As String * 2
    seq     As String * 5
    dtger   As String * 8
    hrger   As String * 6
    reg     As String * 13
    tag     As String * 7
    filler  As String * 20
    fcrlf   As String * 2
End Type


'D0618002900000037066KADETT              BIA3132  R
Type LNdet
    regtipo As String * 1
    pais    As String * 4
    tag     As String * 15
    modelo  As String * 20
    placa   As String * 7
    f1      As String * 2
    oper    As String * 1
    st      As String * 2
    f2      As String * 2
    fcrlf   As String * 2
End Type

Type TRHeader
    tipo As String * 2
    pais    As String * 4
    id      As String * 5
    seqfile As String * 5
    dtger   As String * 8
    hrger   As String * 6
    reg     As String * 6
    valtot  As String * 12
    crlf    As String * 2
End Type

'D0618002900000037066KADETT              BIA3132  R
Type TRdet
    regtipo As String * 1
    regseq  As String * 6
    pais As String * 4
    tag As String * 15
    acesso As String * 4
    dtent As String * 8
    hrent As String * 6
    stent As String * 1
    dtsai As String * 8
    hrsai As String * 6
    valor As String * 8
    stcobranca As String * 1
    stsaida As String * 1
    fbateria As String * 1
    fviolacao As String * 1
    contador  As String * 8
    placa As String * 7
    antpais As String * 4
    antconc As String * 5
    antpraca As String * 4
    antpista As String * 3
    antdt As String * 8
    anthr As String * 6
    motimagem As String * 2
    sttransacao As String * 1
    stmac As String * 8
    filler As String * 30
    crlf As String * 2
End Type

Type RTHeader
    tipo            As String * 2
    pais            As String * 4
    id              As String * 5
    seqfile         As String * 5
    dtger           As String * 8
    hrger           As String * 6
    rejeicao        As String * 2
    reginformado    As String * 6
    regencontrado   As String * 6
    totinformado    As String * 12
    totencontrado   As String * 12
    crlf As String * 2
End Type


Type RFTipo
    tipo            As String * 2
End Type

Type RFHeader
    pais            As String * 4
    id              As String * 5
    seqfile         As String * 5
    dtger           As String * 8
    hrger           As String * 6
    regrejeitado    As String * 6
    totAceito       As String * 12
    totNaoAceito    As String * 12
    crlf            As String * 2
End Type

Type RFDetalhe1
    seqreg          As String * 6
    pais            As String * 4
    tag             As String * 15
    acesso          As String * 4
    entradadia      As String * 8
    entradahora     As String * 6
    saidadia        As String * 8
    saidahora       As String * 6
    valor           As String * 8
    codigo          As String * 2
    crlf            As String * 2
End Type

Type RFDetalhe2
    datapagamento    As String * 8
    valorCGMP       As String * 12
    valorSGMP       As String * 12
    crlf            As String * 2
End Type


Declare Function apiCopyFile Lib "kernel32" Alias "CopyFileA" _
 (ByVal lpExistingFileName As String, _
  ByVal lpNewFileName As String, _
  ByVal bFailIfExists As Long) As Long
  
  
' Enter each of the following Declare statements on one, single line:
Declare Function apiGetPrivateProfileString Lib "kernel32" _
         Alias "GetPrivateProfileStringA" (ByVal lpApplicationName _
         As String, ByVal lpKeyName As Any, ByVal lpDefault As _
         String, ByVal lpReturnedString As String, ByVal nSize As _
         Long, ByVal lpFileName As String) As Long
'
Declare Function apiWritePrivateProfileString Lib _
         "kernel32" Alias "WritePrivateProfileStringA" _
         (ByVal lpApplicationName As String, ByVal lpKeyName _
         As Any, ByVal lpString As Any, ByVal lpFileName As _
         String) As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExA" (ByRef lpExecInfo As SHELLEXECUTEINFO) As Long


Public Sub Shellexe(frase As String, PathWork As String, Tempo As Long, WindowsMode As Integer)

Dim Sei As SHELLEXECUTEINFO
Dim ret As Long
Dim cont As Long
Const passo = 10
'Dim sApp As String
'Dim sParams As String
   
'  sApp = "Bcp.exe"
'  sParams = " GTMOTE_TEST.DBO.EMPL FORMAT nul -T -n -f C:\Data_Empl.fmt -U ## -P ###### -S BRORDSQL3"
    
    Sei.cbSize = Len(Sei)
    Sei.fMask = &H40
    'Sei.hwnd = Me.hwnd
    Sei.lpVerb = "Open"
    Sei.lpFile = "cmd.exe"
    Sei.lpParameters = "/c " & frase
    Sei.lpDirectory = PathWork
    Sei.nShow = WindowsMode
            
    
    Call ShellExecuteEx(Sei)
    
    cont = 0
    ret = WaitForSingleObject(Sei.hProcess, passo)
    While Not ret = 0 And cont * passo < Tempo
        ret = WaitForSingleObject(Sei.hProcess, passo)
        MDI.sta_Barra_MDI.Panels(2).Text = Format(cont * passo)
        cont = cont + 1
        DoEvents
    Wend
        
    Call CloseHandle(Sei.hProcess)

End Sub

Public Sub SuperShell(frase As String, PathWork As String, Tempo As Long, WindowsMode As Integer)
' WindowsMode 0 = oculto, 6 minimizado
Dim pi7z As PROCESS_INFORMATION
Dim sec1 As SECURITY_ATTRIBUTES
Dim sec2 As SECURITY_ATTRIBUTES
Dim sti As STARTUPINFO
Dim ret As Long

sec1.nLength = Len(sec1)
sec2.nLength = Len(sec2)
sti.cb = Len(sti)
sti.dwFlags = 1
sti.wShowWindow = vbMaximizedFocus
'ret = Shell(frase, vbMaximizedFocus) "c:\sce\extra\7zbat.bat"

ret = CreateProcess(vbNullString, frase, sec1, sec2, False, &H28, 0&, PathWork, sti, pi7z)
ret = WaitForSingleObject(pi7z.hProcess, Tempo)

Call CloseHandle(pi7z.hProcess)


End Sub
         

Function Data_Valida(frmForm As Form)
    Dim dtmData
    
    Data_Valida = False

    ' verifica o dia
    If Val(Left(frmForm.mstrData, 2)) > 31 Or Val(Left(frmForm.mstrData, 2)) < 1 Then
        Exit Function
    End If
    
    ' verifica o mes
    If Val(Mid(frmForm.mstrData, 4, 2)) > 12 Or Val(Mid(frmForm.mstrData, 4, 2)) < 1 Then
        Exit Function
    End If
    
    'formata a variavel para data
    dtmData = Format(Date, "Long Date")
    
    'caso "frmForm.mstrData_Hora" seja uma data valida, dtmData devera ser o dia da semana
    dtmData = Format(frmForm.mstrData, "dddd")
    'faco o teste para ver se a data e valida
    If dtmData <> frmForm.mstrData Then
        dtmData = Format(frmForm.mstrData, "dd/mm/yyyy")
        Data_Valida = True
    End If
    If VarType(dtmData) <> vbError Then
        frmForm.mstrData = dtmData
    End If
End Function

Function So_Numeros(Tecla As Integer)
    Select Case Tecla
        Case 8, 48 To 57, 127 'O Primeiro Case para checar os números de 0 a 9, o ponto e a tecla ESC
            So_Numeros = Tecla
        Case 13 '--- Testa se teclou ENTER
        Case Else
            So_Numeros = 0 '-- Se for qualquer tecla diferente de Números, ESC ou ENTER ele não processa
    End Select
End Function

Function Formata_Ano_Mes_Dia(strData As String)
    Dim strAno As String
    Dim strMes As String
    Dim strDia As String
        strAno = Right(strData, 4)
        strMes = Mid(strData, 4, 2)
        strDia = Left(strData, 2)
        Formata_Ano_Mes_Dia = strAno & "," & strMes & "," & strDia
End Function

Sub CopyFile(Sourcefile As String, Destfile As String)
'---------------------------------------------------------------
' PURPOSE: Copy a file on disk from one location to another.
' ACCEPTS: The name of the source file and destination file.
' RETURNS: Nothing
'---------------------------------------------------------------
 Dim result As Long
 If Dir(Sourcefile) = "" Then
    MsgBoxService Chr(34) & Sourcefile & Chr(34) & " is not valid file name."
 Else
    result = apiCopyFile(Sourcefile, Destfile, False)
 End If
 
 End Sub

Public Function CriaDir(nivel1 As String, nivel2 As String, nivel3 As String, nivel4 As String) As String
Dim result As String

result = ""
If nivel1 <> "" Then
    result = result + nivel1
    If Dir(result, vbDirectory) <> "." Then
       MkDir (result)
    End If
    If nivel2 <> "" Then
        result = result + nivel2
        If Dir(result, vbDirectory) <> "." Then
           MkDir (result)
        End If
        If nivel3 <> "" Then
            result = result + nivel3
            If Dir(result, vbDirectory) <> "." Then
               MkDir (result)
            End If
            If nivel4 <> "" Then
               result = result + nivel4
               If Dir(result, vbDirectory) <> "." Then
                  MkDir (result)
               End If
            End If
        End If
    End If
End If
CriaDir = result

End Function

Function Stuff(ByVal Dado As String, Atual As String, Novo As String)
    Dim aux As String
    Dim aux2 As String
    aux2 = Dado
    aux = ""
    Do While InStr(Dado, Atual) <> 0
        aux = aux & Left(Dado, InStr(Dado, Atual) - 1) & Novo
        Dado = Mid(Dado, InStr(Dado, Atual) + Len(Atual))
    Loop
    aux = aux & Dado
    Stuff = aux
End Function


Function Poe_Zero_Esquerda(strString As Variant, bytTamanho As Byte)
On Error GoTo tratarerr
Dim aux As String, aux1 As String

aux = String(bytTamanho, "0")
If Not IsNull(strString) Then
    aux1 = IIf(Val(Trim(strString)) < 0, 0, Val(Trim(strString)))
    aux = Right(aux & Format(aux1), bytTamanho)
    Poe_Zero_Esquerda = aux
Else
    Poe_Zero_Esquerda = aux
End If
Exit Function

tratarerr:
    aux = aux


End Function


Function poevirgula(ivalor As Long) As String

    If ivalor = 0 Then
        poevirgula = "0,00"
    Else
        poevirgula = Left(Format(ivalor), Len(Format(ivalor)) - 2) + "," + Right(Format(ivalor), 2)
    End If
    
End Function

Function Centraliza_Form_Top(frmForm As Form)
    Centraliza_Form_Top = (Screen.Height / 2) - (frmForm.Height / 2)
End Function

Function Centraliza_Form_Left(frmForm As Form)
    Centraliza_Form_Left = (Screen.Width / 2) - (frmForm.Width / 2)
End Function

Function Criptografa(st1 As String) As String
 Dim st2 As String
 Dim ind1 As Integer, ind2 As Integer
 
    st2 = ""
    For ind1 = 1 To Len(st1)
        ind2 = Asc(Mid(st1, ind1, 1)) Xor (255 - ind1)
        st2 = st2 + Chr(ind2)
    Next ind1
    Criptografa = st2
End Function

Public Sub TrataErro(Optional Titulo As String, Optional strArquivo As Variant, Optional strRotina As String)

    Dim i As Integer
    Dim msg As String
    Dim frase As String
    
    On Error GoTo trataerr
    
    msg = Err.Description & vbLf
    If Errors.Count > 0 Then
        If Err.Number = Errors(Errors.Count - 1).Number Then
            For i = 0 To Errors.Count - 2
                msg = msg & Errors(i).Description & vbLf
            Next i
        End If
    End If
    Screen.MousePointer = vbDefault
    frase = Format(msg) & "  Arquivo : " & strArquivo & "  Rotina   : " & strRotina
    Call LogErro(strRotina, frase)
    
    End
    
trataerr:
    End

End Sub


Function DbOpen(ds As String, db As String) As Integer
On Error GoTo DbOpenErr

    DbOpen = False
    Screen.MousePointer = 11
    
    ' Cria connection com o DB
    Set dbApp = New Connection
    dbApp.Provider = "SQLOLEDB"
    dbApp.ConnectionString = "Data Source = " & ds & "; Initial Catalog = " & db & ";" & _
                                  "User Id = " & gsuid & "; Password = " & gsPWD & ";"
    dbApp.ConnectionTimeout = 15
    dbApp.CommandTimeout = 600
    dbApp.Open
    Screen.MousePointer = 0
    DbOpen = True

Exit Function
'
DbOpenErr:
    If Not dbApp Is Nothing Then
       If dbApp.State = adStateOpen Then dbApp.Close
    End If
    Set dbApp = Nothing
    If Err <> 0 Then
        MsgBoxService Err.Source & "-->" & Err.Description, , "Error"
    End If
    End
End Function

Sub FormataGridx(mygrid As Control, myrs As Recordset)

Dim i As Integer
Dim g As MSHFlexGrid
Set g = mygrid

g.Cols = myrs.Fields.Count

'If g.Rows < 18 Then
'   g.Rows = 18
'End If

i = 0
For i = 0 To myrs.Fields.Count - 1
    g.ColWidth(i) = Len(mygrid.TextMatrix(0, i)) * 100 + 100
Next

End Sub

Public Sub ImprimeHeader(Filename As String, Header As String)
Dim arch As Integer
arch = FreeFile
Open gsPath_REL & Filename For Output As arch
 
Print #arch, "<HTML> <HEAD> <TITLE>" & Filename & "</TITLE> </HEAD><BODY>"
Print #arch, "<DIV align=center>"
Print #arch, "<TABLE cellSpacing=0 cellPadding=0 width=""100%"" border=0>"
Print #arch, "<TBODY>"
Print #arch, "  <TR>"
Print #arch, "    <TD align=middle width=200><img src = """ & gsPath_SCEExtra & "logospvf.png"" width=103></td>"
Print #arch, "    <TD align=middle width=""100%"">"
Print #arch, "      <TABLE cellSpacing=0 cellPadding=0 width=""100%"" border=0>"
Print #arch, "        <TBODY>"
Print #arch, "        <TR>"
 Print #arch, "         <TD align=middle width=""100%"">"
Print #arch, "            <P><SPAN style='FONT-SIZE: 16pt'>Controle Financeiro</SPAN></P>"
Print #arch, "          </TD>"
Print #arch, "        </TR>"
Print #arch, "        <TR>"
Print #arch, "          <TD align=middle width=""100%"">"
Print #arch, "            <P><SPAN style=""FONT-SIZE: 16pt""><FONT SIZE=4><B>" & gsEst_Codigo & " - " & gsEst_Nome & " - " & Header & "</B></FONT></SPAN></P>"
Print #arch, "          </TD>"
Print #arch, "        </TR>"
Print #arch, "        </TBODY>"
Print #arch, "      </TABLE>"
Print #arch, "    </TD>"
'
Print #arch, "    <TD align=middle width=200><img src = """ & gsPath_SCEExtra & "logospvf.png"" width=103></TD>"
Print #arch, "  </TR>"
Print #arch, " </TBODY>"
Print #arch, " </TABLE>"
Print #arch, "</DIV>"



'Print #arch, "<HTML> <HEAD> <TITLE>" & Filename & "</TITLE> </HEAD><BODY>"
'Print #arch, "<P ALIGN=CENTER><FONT SIZE=4><B>" & gsEst_Codigo & " - " & gsEst_Nome & " - " & Header & "</B></FONT></P>"
'Print #arch, "<P> <BR> </P>"



Close arch

End Sub

Public Sub ImprimeFooter(Filename As String, Footer As String)
Dim arch As Integer

arch = FreeFile
Open gsPath_REL & Filename For Append As arch
Print #arch, "<P><BR></P>"
Print #arch, "<P ALIGN=LEFT><FONT SIZE=2><B>" & Footer & "</B> </FONT> </P>"
Print #arch, "</BODY></HTML>"
 
Close arch
 
End Sub

Public Sub GravaEventos(codigo As Integer, placa As String, usuario As String, turno As Integer, texto As String, aparam As Long, bparam As Long)
On Error GoTo trataerr
'Id autoid
'tsDataHora
'ICodigo
'CPlaca
'CUsuario
'LTurno
'SzTexto
'LParam
'IParam
'Dim rs As New Recordset
Dim frase As String

frase = ""
frase = frase & "INSERT INTO " + gsPath_DB + ".dbo.tb_eventos (tsdatahora,icodigo,cplaca,cusuario,lturno,sztexto,iparam,lparam)"
frase = frase & " values ("
frase = frase & "'" & Format(Now, "yyyymmdd hh:mm:ss") & "',"
frase = frase & "'" & codigo & "',"
frase = frase & "'" & placa & "',"
frase = frase & "'" & usuario & "',"
frase = frase & "'" & turno & "',"
frase = frase & "'" & Mid(texto, 1, 50) & "',"
'frase = frase & "'" & aparam & "',"
frase = frase & "'0',"
frase = frase & "'" & bparam & "')"
Set rsGeral = dbApp.Execute(frase)
Set rsGeral = Nothing

Exit Sub

trataerr:
Call TrataErro(app.title, Error, "GravaEventos")

End Sub

Public Sub ImprimeExtra(Filename As String, extra() As String)
On Error Resume Next

Dim arch As Integer
Dim x As Integer
arch = FreeFile
Open gsPath_REL & Filename For Append As arch
For x = 0 To UBound(extra) - 1
    Print #arch, "<P ALIGN=LEFT><FONT SIZE=2><B>" & extra(x) & "</B> </FONT> </P>"
Next
'Print #arch, "<P><BR></P>"
Close arch

End Sub

Public Sub Imprimegrid(Filename As String, Grid1 As MSHFlexGrid)
On Error Resume Next

Dim arch As Integer
Dim cont As Integer
Dim Y, x As Integer

arch = FreeFile
Open gsPath_REL & Filename For Append As arch

Print #arch, "<TABLE WIDTH=100% BORDER=1 CELLPADDING=1 CELLSPACING=1>"
Print #arch, "<THEAD>"
Print #arch, "<TR ALIGN=RIGHT>"
For cont = 0 To Grid1.Cols - 1
    Print #arch, "<TH><P><FONT SIZE=2>" & Grid1.TextMatrix(0, cont) & "</FONT> </P> </TH>"
Next
Print #arch, "</TR></THEAD><TBODY>"


For Y = 1 To Grid1.Rows - 1
    Print #arch, "<TR ALIGN=RIGHT>"
    For x = 0 To Grid1.Cols - 1
        Print #arch, "<TD ><P><FONT SIZE=2>" & IIf(Grid1.TextMatrix(Y, x) <> "", Grid1.TextMatrix(Y, x), "-") & "</FONT></P></TD>"
    Next
    Print #arch, "</TR>"
Next
Print #arch, "</TBODY></TABLE>"
'Print #arch, "<P><BR></P>"

Close arch

End Sub

Public Sub ImprimegridFnac(Filename As String, Grid1 As MSHFlexGrid)
On Error Resume Next
Dim aux1, aux2, aux3 As String
Dim arch As Integer
Dim cont As Integer
Dim Y, x As Integer
Dim folha As Integer

folha = 1
arch = FreeFile
'kill gsPath_REL & Filename & Format(Poe_Zero_Esquerda(folha, 5))
Open gsPath_REL & Filename For Output As arch
 
'Print #arch, "<TR ALIGN=RIGHT>"
'For cont = 0 To Grid1.Cols - 1
'    Print #arch, "<TH><P><FONT SIZE=2>" & Grid1.TextMatrix(0, cont) & "</FONT> </P> </TH>"
'Next
'Print #arch, "</TR></THEAD><TBODY>"
 
Print #arch, "<TABLE WIDTH=100% BORDER=1 CELLPADDING=1 CELLSPACING=1>"
Print #arch, "<THEAD>"
 
For Y = 1 To Grid1.Rows - 1 Step 3
    Print #arch, "<TR ALIGN=RIGHT>"
        aux1 = ""
        Print #arch, "<TD ALIGN=CENTER><P><FoNT SIZE=3><br><br><br><br><br><br>" & aux1 & "<br></FONT></P></TD>"
        aux2 = ""
        Print #arch, "<TD ALIGN=CENTER><P><FONT SIZE=4><br>" & aux2 & "<br><br></FONT></P></TD>"
        aux3 = ""
        Print #arch, "<TD ALIGN=CENTER><P><FONT SIZE=4><br>" & aux3 & "<br><br></FONT></P></TD>"
    
        aux1 = Poe_Zero_Esquerda(IIf(Grid1.TextMatrix(Y + 0, 0) <> "", Grid1.TextMatrix(Y + 0, 0), 0), 14)
        aux2 = Poe_Zero_Esquerda(IIf(Grid1.TextMatrix(Y + 1, 0) <> "", Grid1.TextMatrix(Y + 1, 0), 0), 14)
        aux3 = Poe_Zero_Esquerda(IIf(Grid1.TextMatrix(Y + 2, 0) <> "", Grid1.TextMatrix(Y + 2, 0), 0), 14)
    
    Print #arch, "<TR ALIGN=RIGHT>"
        Print #arch, "<TD ALIGN=CENTER><P><FoNT SIZE=2>" & aux1 & "</FONT></P></TD>"
        Print #arch, "<TD ALIGN=CENTER><P><FONT SIZE=2>" & aux2 & "</FONT></P></TD>"
        Print #arch, "<TD ALIGN=CENTER><P><FONT SIZE=2>" & aux3 & "</FONT></P></TD>"
'font-family: C39P24DhTt
    Print #arch, "</TR>"
    Print #arch, "<TR ALIGN=RIGHT border=0>"
        Print #arch, "<TD ALIGN=CENTER><P><font style='font-size: 42px;font-family: C39P24DhTt;'>" & "*" & aux1 & "*" & "</FONT></P></TD>"
        Print #arch, "<TD ALIGN=CENTER><P><font style='font-size: 42px;font-family: C39P24DhTt;'>" & "*" & aux2 & "*" & "</FONT></P></TD>"
        Print #arch, "<TD ALIGN=CENTER><P><font style='font-size: 42px;font-family: C39P24DhTt;'>" & "*" & aux3 & "*" & "</FONT></P></TD>"
    Print #arch, "</TR>"
    Print #arch, "<TR ALIGN=RIGHT>"
        Print #arch, "<TD ALIGN=CENTER><P><FoNT style='font-size: 3px;font-family: arial;'>**********************************</FONT></P></TD>"
        Print #arch, "<TD ALIGN=CENTER><P><FONT style='font-size: 3px;font-family: arial;'>***********************************</FONT></P></TD>"
        Print #arch, "<TD ALIGN=CENTER><P><FoNT style='font-size: 3px;font-family: arial;'>**********************************</FONT></P></TD>"
    Print #arch, "</TR>"
Next
Print #arch, "</TBODY></TABLE>"
aux1 = "----------------------------------------------------------------------------------------------------------"
'Print #arch, "<font style='font-size: 2px;font-face: arial;'>" & aux1 & "</FONT></P>"
 
Close arch

End Sub

Public Sub ImprimeRel(Filename As String)
On Error Resume Next
 
Dim arch As Integer

arch = FreeFile
Open gsPath_REL & Mid(Filename, 1, Len(Filename) - 5) & ".bat" For Output As arch
Print #arch, "start " & gsPath_REL & Filename
Close arch

    Shell (gsPath_REL & Mid(Filename, 1, Len(Filename) - 5) & ".bat")


End Sub

Public Sub ImprimeRelDel(Filename As String)
On Error Resume Next

Kill (gsPath_REL & Mid(Filename, 1, Len(Filename) - 5) & ".bat")


End Sub

Function FgetParam(fname As String, Sessao As String, app As String, param As String, dt As Date, valordefault As String) As String
    Dim frase As String
    Dim rs As New Recordset
    Dim retorno As String
    Dim ret As String
    
    Dim dtstr As String
    dtstr = Format(dt, "yyyy-MM-dd")
    
    'USE DO BANCO CORRETO
    frase = ""
    frase = frase & " use " + gsPath_DB & vbCrLf
    dbApp.Execute (frase)
    
    frase = ""
    frase = frase & "select " + gsPath_DB + ".dbo.fGetIni('VMSCE','SCE','SCE1','" & param & "','')"
    Set rs = dbApp.Execute(frase)
    retorno = rs(0)
         
    'CASO O PARAMETRO NAO EXISTA, LER DO INI E GRAVAR NO BANCO DE DADOS
    If retorno = "" Then
     
        retorno = MyGetPar(fname, Sessao, param, "")
        If retorno = "" Then retorno = valordefault
        ret = FsetParam(app, param, DateTime.Now, retorno)

        frase = ""
        frase = frase & "select dbo.fGetIni('VMSCE','SCE','SCE1','" & param & "','')"
        Set rs = dbApp.Execute(frase)
        retorno = rs(0)
    
    End If

    FgetParam = retorno
 
End Function

Function FsetParam(app As String, param As String, dt As Date, valor As String) As String
    Dim frase As String
    Dim rs As New Recordset
    
    Dim retorno As String
    
    rs.CursorType = adOpenStatic
    
    Dim dtstr As String
    dtstr = Format(dt, "yyyy-MM-dd HH:mm:ss")
    
    'USE DO BANCO CORRETO
    frase = ""
    frase = frase & " use " + gsPath_DB & vbCrLf
    dbApp.Execute (frase)
    
    'GRAVA PARAMETRO NO BANCO DE DADOS
    frase = ""
    frase = frase & "EXEC PR_SETINI '" & app & ".SCE1." & param & "','" & valor & "','VMSCE'"
    Set rs = dbApp.Execute(frase)

    
    frase = ""
    frase = frase & "select dbo.fGetIni('VMSCE','SCE','SCE1','" & param & "','-1')"
    Set rs = dbApp.Execute(frase)
    
    If (rs(0) <> -1) Then
        retorno = "1"
    Else
        retorno = "-1"
    End If
        
    FsetParam = retorno
 
End Function



Function MyGetPar(fname As String, Sessao As String, Chave As String, inicial As String) As String
' "APP", "Sessao", "Key", "default"
   ' Find the parent app for a file with the given extension
   
   Dim ret As String
   Dim nSize As Integer
   ' Dim fname As String
   ' fname = App.Path & "\" & Apl & ".ini"
   ret = String$(255, " ")
   nSize = apiGetPrivateProfileString(Sessao, Chave, inicial, ret, Len(ret), fname)
   MyGetPar = Mid(ret, 1, nSize)
   
End Function

Public Function MyMsgBox(frase, tipo, Titulo, alerta) As Integer
On Error GoTo trataerr

Dim resp As Integer

giRespMsg = 99

frmMensagem.txtMSG = frase
frmMensagem.lblTitulo = Titulo
frmMensagem.lblAlerta = alerta
frmMensagem.Show 1

MyMsgBox = giRespMsg

Exit Function

trataerr:
    Call TrataErro(app.title, Error, "MyMsgBox")

End Function


Public Sub MySetPar(fname As String, Sessao As String, Chave As String, valor As String)
' SaveSettingx App.Title, "SISTEMA", "NextTRN", CVDate(gsNextTrn)
   
   Dim nSize As Integer
'   Dim fname As String
'   fname = App.Path & "\" & Apl & ".ini"
   nSize = apiWritePrivateProfileString(Sessao, Chave, valor, fname)
   
End Sub


Public Sub ArqGrid(Filename As String, Grid1 As MSHFlexGrid)
On Error Resume Next
 
Dim arch As Integer
Dim cont As Integer
Dim Y, x As Integer
Dim frase As String

arch = FreeFile
Open gsPath_REL & Filename For Output As arch
 
frase = ""
For cont = 0 To Grid1.Cols - 1
    frase = frase & LTrim(RTrim(Grid1.TextMatrix(0, cont))) & gsDelimitador
Next
frase = Left(frase, Len(frase) - 1)
Print #arch, frase
 
For Y = 1 To Grid1.Rows - 1
    frase = ""
    For x = 0 To Grid1.Cols - 1
        frase = frase + IIf(Grid1.TextMatrix(Y, x) <> "", Grid1.TextMatrix(Y, x), "0") & gsDelimitador
    Next
    frase = Left(frase, Len(frase) - 1)
    Print #arch, frase
Next
Close arch

End Sub

Public Sub Espera(Optional Tempo As Integer)
Dim bytTempo_Espera As Byte
Dim dblInicio_Tempo As Double
    
bytTempo_Espera = IIf(Tempo = 0, 2, Tempo)
dblInicio_Tempo = Timer
Do While Timer < dblInicio_Tempo + bytTempo_Espera
    DoEvents
Loop

End Sub

Function isFunctionExists(fname As String) As Boolean
    Dim retorno As Boolean
    
    frase = ""
    frase = frase & " use " + gsPath_DB & vbCrLf
    Set rs = dbApp.Execute(frase)
    
    frase = ""
    frase = frase & "select top 1 szValor from tb_AppParam where szApp='VER' and szParam LIKE '%" + fname + "_%'  order by tsvalidade desc"
    Set rs = dbApp.Execute(frase)
    
    If rs.BOF And rs.EOF Then
        retorno = False
    Else: retorno = True
    End If
    
    isFunctionExists = retorno
    
End Function


Public Sub checkFunctions()
Dim rsFunction As Recordset
Dim functionExists As Boolean

If Not DbOpen(gsPath_DS, gsPath_DB) Then
    MsgBoxService "O banco de dados não pode ser aberto! Aplicação será finalizada!", vbCritical, "Erro"
    End
End If

'***********************************************************************'
'**************************** FGETINI **********************************'
'***********************************************************************'
functionExists = isFunctionExists("fgetini")

If functionExists = False Then

frase = ""
frase = frase & " use " + gsPath_DB & vbCrLf
frase = frase + " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[FGETINI]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))" & vbCrLf
frase = frase + " DROP FUNCTION [dbo].[fGetIni]" & vbCrLf
Set rsFunction = dbApp.Execute(frase)

'ATUALIZA FGETINI
frase = ""
frase = frase & " use " + gsPath_DB & vbCrLf
Set rsFunction = dbApp.Execute(frase)

'CASO NAO TENHA VERSAO GRAVADA OU VERSAO MAIS ANTIGA, ENTAO ATUALIZAR
frase = ""
frase = frase + "IF  NOT EXISTS (SELECT * FROM SYS.OBJECTS WHERE OBJECT_ID = OBJECT_ID(N'" + gsPath_DB + ".[DBO].[FGETINI]') AND TYPE IN (N'FN', N'IF', N'TF', N'FS', N'FT'))" & vbCrLf
frase = frase + " BEGIN " & vbCrLf
frase = frase + "  DECLARE @frase AS NVARCHAR(MAX)   " & vbCrLf
frase = frase + "  SET @frase = N'CREATE FUNCTION [DBO].[FGETINI] (@CHOST NVARCHAR(50),@CFILE NVARCHAR(50),@CSECTION NVARCHAR(50),@CPARAM NVARCHAR(50),   " & vbCrLf
frase = frase + "  @CDEFAULT NVARCHAR(100)) RETURNS NVARCHAR(100)   " & vbCrLf
frase = frase + "  AS   " & vbCrLf
frase = frase + "  BEGIN   " & vbCrLf
frase = frase + "  DECLARE @CRET NVARCHAR(100)   " & vbCrLf
frase = frase + "  IF @CHOST IS NULL   " & vbCrLf
frase = frase + "  SET @CHOST=HOST_NAME()   " & vbCrLf
frase = frase + "  SET @CRET=DBO.FGETPARAM(''INI'',@CHOST+''_''+@CFILE+''.''+@CSECTION+''.''+@CPARAM,GETDATE())   " & vbCrLf
frase = frase + "  IF @CRET IS NULL OR @CRET=''''   " & vbCrLf
frase = frase + "  SET @CRET=@CDEFAULT   " & vbCrLf
frase = frase + "  RETURN @CRET   " & vbCrLf
frase = frase + "  END   " & vbCrLf
frase = frase + "  '   " & vbCrLf
frase = frase + "  EXECUTE SP_EXECUTESQL @frase     " & vbCrLf
frase = frase + " END"
frase = frase + " EXEC pr_setver 'fgetini', '0.0.1'"
Set rsFunction = dbApp.Execute(frase)

End If




'***********************************************************************'
'**************************** SETINI ***********************************'
'***********************************************************************'

functionExists = isFunctionExists("PR_SETINI")

If functionExists = False Then

'DROPA PR_SETINI
frase = ""
frase = frase & " use " + gsPath_DB & vbCrLf
Set rsFunction = dbApp.Execute(frase)
frase = ""

frase = frase + " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[pr_setini]') AND type in (N'P', N'PC'))" & vbCrLf
frase = frase + " DROP PROCEDURE [dbo].[pr_setini]" & vbCrLf
Set rsFunction = dbApp.Execute(frase)

'RECRIA PR_SETINI
frase = ""
frase = frase + " CREATE PROCEDURE [DBO].[pr_setini] @APLIC NVARCHAR(60),@VER NVARCHAR(120), @HOSTNAME NVARCHAR(255)=NULL" & vbCrLf
frase = frase + " AS" & vbCrLf
frase = frase + " BEGIN" & vbCrLf
frase = frase + "     IF @HOSTNAME IS NULL SET @HOSTNAME = HOST_NAME()" & vbCrLf
frase = frase + "     DECLARE @RET INT" & vbCrLf
frase = frase + "     SET @RET=0" & vbCrLf
frase = frase + "     SET @APLIC=@HOSTNAME+'_'+@APLIC" & vbCrLf
frase = frase + "     DECLARE @PREV NVARCHAR(120)" & vbCrLf
frase = frase + "     SET @PREV=DBO.FGETPARAM('INI',@APLIC,NULL)" & vbCrLf
frase = frase + "     IF @PREV<>@VER" & vbCrLf
frase = frase + "     BEGIN" & vbCrLf
frase = frase + "         INSERT INTO TB_APPPARAM (SZAPP,SZPARAM,SZVALOR,TSVALIDADE) VALUES ('INI',@APLIC,@VER,GETDATE())" & vbCrLf
frase = frase + "         SET @RET=1" & vbCrLf
frase = frase + "     END" & vbCrLf
frase = frase + "     ELSE" & vbCrLf
frase = frase + "     BEGIN" & vbCrLf
frase = frase + "         IF @PREV<>@VER" & vbCrLf
frase = frase + "         BEGIN" & vbCrLf
frase = frase + "             INSERT INTO TB_APPPARAM (SZAPP,SZPARAM,SZVALOR,TSVALIDADE) VALUES ('INI',@APLIC,@VER,GETDATE())" & vbCrLf
frase = frase + "             SET @RET=2" & vbCrLf
frase = frase + "         END" & vbCrLf
frase = frase + "     END" & vbCrLf
frase = frase + "     SELECT @RET RET" & vbCrLf
frase = frase + " END"
Set rsFunction = dbApp.Execute(frase)
frase = ""

frase = ""
frase = frase & " use " + gsPath_DB & vbCrLf
Set rsFunction = dbApp.Execute(frase)

frase = frase + " EXEC pr_setver 'pr_setini', '0.0.1'"
Set rsFunction = dbApp.Execute(frase)
frase = ""

End If




'***********************************************************************'
'*************************** TB_CADEST *********************************'
'***********************************************************************'
frase = ""
frase = frase & " use " + gsPath_DBCAD & vbCrLf
Set rsFunction = dbApp.Execute(frase)
frase = ""

frase = frase + " IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tb_cadest]') AND type in (N'U'))" & vbCrLf
frase = frase + " BEGIN" & vbCrLf
frase = frase + " CREATE TABLE [dbo].[tb_cadest](" & vbCrLf
frase = frase + "     [IIssuer] [smallint] NOT NULL," & vbCrLf
frase = frase + "     [LTag] [int] NOT NULL," & vbCrLf
frase = frase + "     [SzModelo] [nvarchar](20) NULL," & vbCrLf
frase = frase + "     [CPLACA] [nchar](7) NULL," & vbCrLf
frase = frase + "     [TID] [nchar](20) NULL," & vbCrLf
frase = frase + "     [COper] [nchar](1) NULL," & vbCrLf
frase = frase + "     [Cst] [nchar](2) NULL," & vbCrLf
frase = frase + "     [LSEQFILE] [int] NULL" & vbCrLf
frase = frase + " ) ON [PRIMARY]" & vbCrLf
frase = frase + " END"
Set rsFunction = dbApp.Execute(frase)
frase = ""

'***********************************************************************'
'************************** TB_CADEST_INC ******************************'
'***********************************************************************'
frase = frase & " use " + gsPath_DBCAD & vbCrLf
Set rsFunction = dbApp.Execute(frase)
frase = ""

frase = frase + " IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tb_cadest_inc]') AND type in (N'U'))" & vbCrLf
frase = frase + " BEGIN" & vbCrLf
frase = frase + " CREATE TABLE [dbo].[tb_cadest_inc](" & vbCrLf
frase = frase + "     [IIssuer] [smallint] NOT NULL," & vbCrLf
frase = frase + "     [LTag] [int] NOT NULL," & vbCrLf
frase = frase + "     [SzModelo] [nvarchar](20) NULL," & vbCrLf
frase = frase + "     [CPLACA] [nchar](7) NULL," & vbCrLf
frase = frase + "     [TID] [nchar](20) NULL," & vbCrLf
frase = frase + "     [COper] [nchar](1) NULL," & vbCrLf
frase = frase + "     [CSt] [nchar](2) NULL," & vbCrLf
frase = frase + "     [LSeqfile] [int] NULL," & vbCrLf
frase = frase + "     [TbId] [int] IDENTITY(1,1) NOT NULL," & vbCrLf
frase = frase + " PRIMARY KEY CLUSTERED" & vbCrLf
frase = frase + " (" & vbCrLf
frase = frase + "     [TbId] Asc" & vbCrLf
frase = frase + " )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]" & vbCrLf
frase = frase + " ) ON [PRIMARY]" & vbCrLf
frase = frase + " END"
Set rsFunction = dbApp.Execute(frase)
frase = ""


'***********************************************************************'
'************************** TB_CADEST_TMP ******************************'
'***********************************************************************'
frase = frase & " use " + gsPath_DBCAD & vbCrLf
Set rsFunction = dbApp.Execute(frase)
frase = ""

frase = frase + " IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tb_cadest_tmp]') AND type in (N'U'))" & vbCrLf
frase = frase + " BEGIN" & vbCrLf
frase = frase + " CREATE TABLE [dbo].[tb_cadest_tmp](" & vbCrLf
frase = frase + "     [IIssuer] [smallint] NOT NULL," & vbCrLf
frase = frase + "     [LTag] [int] NOT NULL," & vbCrLf
frase = frase + "     [SzModelo] [nvarchar](20) NULL," & vbCrLf
frase = frase + "     [CPLACA] [nchar](7) NULL," & vbCrLf
frase = frase + "     [TID] [nchar](20) NULL," & vbCrLf
frase = frase + "     [COper] [nchar](1) NULL," & vbCrLf
frase = frase + "     [CSt] [nchar](2) NULL," & vbCrLf
frase = frase + "     [LSeqfile] [int] NULL" & vbCrLf
frase = frase + " ) ON [PRIMARY]" & vbCrLf
frase = frase + " END"
Set rsFunction = dbApp.Execute(frase)
frase = ""

'***********************************************************************'
'************************** TB_CADESTCTL *******************************'
'***********************************************************************'
frase = frase & " use " + gsPath_DBCAD & vbCrLf
Set rsFunction = dbApp.Execute(frase)
frase = ""

frase = frase + " IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tb_cadestCTL]') AND type in (N'U'))" & vbCrLf
frase = frase + " BEGIN" & vbCrLf
frase = frase + " CREATE TABLE [dbo].[tb_cadestCTL](" & vbCrLf
frase = frase + "     [LSeqfile] [int] NULL," & vbCrLf
frase = frase + "     [SzArquivo] [nvarchar](25) NULL," & vbCrLf
frase = frase + "     [TsAtualizacao] [datetime] NULL," & vbCrLf
frase = frase + "     [CTipo] [nchar](2) NULL," & vbCrLf
frase = frase + "     [LTotal] [int] NULL," & vbCrLf
frase = frase + "     [LRegistros] [int] NULL," & vbCrLf
frase = frase + "     [LIncl] [int] NULL," & vbCrLf
frase = frase + "     [LRemo] [int] NULL," & vbCrLf
frase = frase + "     [LAlte] [int] NULL" & vbCrLf
frase = frase + " ) ON [PRIMARY]" & vbCrLf
frase = frase + " END"
Set rsFunction = dbApp.Execute(frase)
frase = ""

'***********************************************************************'
'************************** VIEW_CADTAG ********************************'
'***********************************************************************'
functionExists = isFunctionExists("VIEW_CADTAG")

If functionExists = False Then

frase = frase & " use " + gsPath_DB & vbCrLf
Set rsFunction = dbApp.Execute(frase)
frase = ""

frase = frase + " IF  EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[tb_cadtag]'))" & vbCrLf
frase = frase + " DROP VIEW [dbo].[tb_cadtag]" & vbCrLf
Set rsFunction = dbApp.Execute(frase)
frase = ""

'RECRIA VIEW CADTAG
frase = frase & " use " + gsPath_DB & vbCrLf
Set rsFunction = dbApp.Execute(frase)
frase = ""

frase = frase + " CREATE VIEW [dbo].[tb_cadtag]" & vbCrLf
frase = frase + " AS" & vbCrLf
frase = frase + " SELECT IIssuer,LTag,CPLACA,NULL CF1,COper,NULL CF2,NULL CF4,LSEQFILE" & vbCrLf
frase = frase + " From SCECAD.dbo.TB_CADEST" & vbCrLf
frase = frase + " Where 1 = 1" & vbCrLf
frase = frase + " AND 'EST'= dbo.fgetparam('INI','VMSCE_SCE.SCE1.tipoLista',GETDATE())" & vbCrLf
frase = frase + " AND Cst in ('00','06','99')" & vbCrLf
frase = frase + " Union" & vbCrLf
frase = frase + " SELECT  IIssuer,LTag,CPLACA,CF1,COper,CF2,CF4,LSEQFILE" & vbCrLf
frase = frase + " From SCECAD.dbo.tb_cadtag" & vbCrLf
frase = frase + " Where 'TIV'=dbo.fgetparam('INI','VMSCE_SCE.SCE1.tipoLista',GETDATE())"
Set rsFunction = dbApp.Execute(frase)
frase = ""
frase = frase + " EXEC pr_setver 'VIEW_CADTAG', '0.0.1'"
Set rsFunction = dbApp.Execute(frase)
frase = ""

End If

'***********************************************************************'
'************************** VIEW_CADNELA *******************************'
'***********************************************************************'
functionExists = isFunctionExists("VIEW_CADNELA")

If functionExists = False Then

frase = frase & " use " + gsPath_DB & vbCrLf
Set rsFunction = dbApp.Execute(frase)
frase = ""

frase = frase + " IF  EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[tb_cadnela]'))" & vbCrLf
frase = frase + " DROP VIEW [dbo].[tb_cadnela]" & vbCrLf
Set rsFunction = dbApp.Execute(frase)
frase = ""

'RECRIA VIEW CADNELA
frase = frase & " use " + gsPath_DB & vbCrLf
Set rsFunction = dbApp.Execute(frase)
frase = ""

frase = frase + " CREATE VIEW [dbo].[tb_cadnela]" & vbCrLf
frase = frase + " AS" & vbCrLf
frase = frase + " SELECT  IIssuer,LTag,SzModelo,CPLACA,COper,Cst,LSEQFILE" & vbCrLf
frase = frase + " From SCECAD.dbo.tb_cadest" & vbCrLf
frase = frase + " Where 1 = 1" & vbCrLf
frase = frase + " AND 'EST'= dbo.fgetparam('INI','VMSCE_SCE.SCE1.tipoLista',GETDATE())" & vbCrLf
frase = frase + " AND cst not in ('00','06','99')" & vbCrLf
frase = frase + " Union" & vbCrLf
frase = frase + " SELECT  IIssuer,LTag,SzModelo,CPLACA,COper,CSt,LSEQFILE" & vbCrLf
frase = frase + " From SCECAD.dbo.tb_cadnela" & vbCrLf
frase = frase + " Where 'TIV'= dbo.fgetparam('INI','VMSCE_SCE.SCE1.tipoLista',GETDATE())" & vbCrLf
Set rsFunction = dbApp.Execute(frase)
frase = ""

frase = ""
frase = frase & " use " + gsPath_DB & vbCrLf
Set rsFunction = dbApp.Execute(frase)

frase = frase + " EXEC pr_setver 'VIEW_CADNELA', '0.0.1'"
Set rsFunction = dbApp.Execute(frase)
frase = ""

End If

'***********************************************************************'
'************************** VIEW_CADESTCTL *****************************'
'***********************************************************************'
functionExists = isFunctionExists("VIEW_CADEST")

If functionExists = False Then

'DROPA VIEW CADESTCTL
frase = frase & " use " + gsPath_DB & vbCrLf
Set rsFunction = dbApp.Execute(frase)
frase = ""

frase = frase + " IF  EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[tb_cadestctl]'))" & vbCrLf
frase = frase + " DROP VIEW [dbo].[tb_cadestctl]" & vbCrLf
Set rsFunction = dbApp.Execute(frase)
frase = ""

'RECRIA VIEW CADESTCTL
frase = frase & " use " + gsPath_DB & vbCrLf
Set rsFunction = dbApp.Execute(frase)
frase = ""

frase = frase + " CREATE VIEW [dbo].[tb_cadestctl]" & vbCrLf
frase = frase + " AS" & vbCrLf
frase = frase + " SELECT * FROM scecad.dbo.tb_cadestctl" & vbCrLf
Set rsFunction = dbApp.Execute(frase)
frase = ""

frase = ""
frase = frase & " use " + gsPath_DB & vbCrLf
Set rsFunction = dbApp.Execute(frase)

frase = frase + " EXEC pr_setver 'VIEW_CADESTCTL', '0.0.1'"
Set rsFunction = dbApp.Execute(frase)
frase = ""

End If



'***********************************************************************'
'************************** PR_ATUETT **********************************'
'***********************************************************************'
functionExists = isFunctionExists("pr_AtuETT")

If functionExists = False Then


'USE SCECAD
frase = frase & " use " + gsPath_DBCAD & vbCrLf
Set rsFunction = dbApp.Execute(frase)
frase = ""

frase = frase + " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[pr_AtuETT]') AND type in (N'P', N'PC'))" & vbCrLf
frase = frase + " DROP PROCEDURE [dbo].[pr_AtuETT]" & vbCrLf
Set rsFunction = dbApp.Execute(frase)
frase = ""

frase = frase + " CREATE  Procedure [dbo].[pr_AtuETT]" & vbCrLf
frase = frase + " @lseqfile as integer = 1" & vbCrLf
frase = frase + " as" & vbCrLf
frase = frase + " begin" & vbCrLf
frase = frase + " SET NOCOUNT ON" & vbCrLf
frase = frase + " if (select COUNT(*) from tb_cadest_tmp) <> 0" & vbCrLf
frase = frase + " begin" & vbCrLf
frase = frase + "     drop table tb_cadest" & vbCrLf
frase = frase + "     SELECT [IIssuer],[LTag],[SzModelo],[CPLACA],[TID],[COper],[Cst],@lseqfile LSEQFILE" & vbCrLf
frase = frase + "     into tb_cadest from tb_cadest_tmp" & vbCrLf
frase = frase + "     if exists (select * from dbo.sysindexes where name = N'ix_tag' and id = object_id(N'[dbo].[tb_cadest]'))" & vbCrLf
frase = frase + "     drop index [dbo].[tb_cadest].[ix_tag]" & vbCrLf
frase = frase + "    CREATE  INDEX [ix_tag] ON [dbo].[tb_cadest]([IIssuer], [LTag]) ON [PRIMARY]" & vbCrLf
frase = frase + "     if exists (select * from dbo.sysindexes where name = N'ix_placa' and id = object_id(N'[dbo].[[tb_cadest]]'))" & vbCrLf
frase = frase + "     drop index [dbo].[tb_cadest].[ix_placa]" & vbCrLf
frase = frase + "     CREATE  INDEX [ix_placa] ON [dbo].[tb_cadest]([CPLACA]) ON [PRIMARY]" & vbCrLf
frase = frase + "     truncate table tb_cadest_tmp" & vbCrLf
frase = frase + " End" & vbCrLf
frase = frase + " End" & vbCrLf
Set rsFunction = dbApp.Execute(frase)

frase = ""
frase = frase & " use " + gsPath_DB & vbCrLf
Set rsFunction = dbApp.Execute(frase)

frase = ""
frase = frase + " EXEC pr_setver 'PR_ATUETT', '0.0.1'"
Set rsFunction = dbApp.Execute(frase)
frase = ""

End If



'***********************************************************************'
'************************** pr_AtuEST_INC ********************************'
'***********************************************************************'
functionExists = isFunctionExists("pr_AtuEST_INC")

If functionExists = False Then

'CRIACAO DA ATUEST_INC
frase = ""
frase = frase & " use " + gsPath_DBCAD & vbCrLf
Set rsFunction = dbApp.Execute(frase)

frase = ""
frase = frase + " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[pr_AtuEST_INC]') AND type in (N'P', N'PC'))" & vbCrLf
frase = frase + " DROP PROCEDURE [dbo].[pr_AtuEST_INC]" & vbCrLf
Set rsFunction = dbApp.Execute(frase)

frase = ""
frase = frase + " CREATE PROC [dbo].[pr_AtuEST_INC] @lseqfile AS integer = 1" & vbCrLf
frase = frase + " AS" & vbCrLf
frase = frase + " BEGIN" & vbCrLf
frase = frase + "     SET NOCOUNT ON" & vbCrLf
frase = frase + "     DECLARE @inc AS int" & vbCrLf
frase = frase + "     DECLARE @rem AS int" & vbCrLf
frase = frase + "     DECLARE @alt AS int" & vbCrLf
frase = frase + "     DECLARE @inicial AS int" & vbCrLf
frase = frase + "     DECLARE @final AS int" & vbCrLf
frase = frase + "     DECLARE @reg AS int" & vbCrLf
frase = frase + "     DECLARE @tblOutput AS TABLE (" & vbCrLf
frase = frase + "         totRem bigint" & vbCrLf
frase = frase + "         , totIns bigint" & vbCrLf
frase = frase + "     )" & vbCrLf
frase = frase + "     SET @INICIAL = (SELECT CONVERT(bigint, rows)" & vbCrLf
frase = frase + "     From sysindexes" & vbCrLf
frase = frase + "     WHERE id = OBJECT_ID('tb_cadest')" & vbCrLf
frase = frase + "     AND indid < 2)" & vbCrLf
frase = frase + "     MERGE tb_cadest AS Destino" & vbCrLf
frase = frase + "     USING (SELECT" & vbCrLf
frase = frase + "         iissuer," & vbCrLf
frase = frase + "         ltag," & vbCrLf
frase = frase + "         cplaca," & vbCrLf
frase = frase + "         coper" & vbCrLf
frase = frase + "     FROM tb_cadest_inc) AS Origem" & vbCrLf
frase = frase + "         ON Destino.iissuer = Origem.iissuer" & vbCrLf
frase = frase + "         AND Destino.ltag = Origem.ltag" & vbCrLf
frase = frase + "     WHEN MATCHED AND UPPER(Origem.coper) = 'R'" & vbCrLf
frase = frase + "         THEN DELETE" & vbCrLf
frase = frase + "     WHEN MATCHED AND (UPPER(Origem.coper) = 'I' OR UPPER(Origem.coper) = 'S') THEN" & vbCrLf
frase = frase + "         UPDATE SET Destino.CPlaca = Origem.CPlaca" & vbCrLf
frase = frase + "     WHEN NOT MATCHED AND UPPER(Origem.coper) = 'I' THEN" & vbCrLf
frase = frase + "         INSERT (" & vbCrLf
frase = frase + "            iissuer , ltag, cplaca, coper" & vbCrLf
frase = frase + "         ) VALUES (" & vbCrLf
frase = frase + "             Origem.iissuer , Origem.ltag, Origem.cplaca, Origem.coper" & vbCrLf
frase = frase + "         )" & vbCrLf
frase = frase + "     OUTPUT" & vbCrLf
frase = frase + "         deleted.ltag,inserted.ltag into @tblOutput;" & vbCrLf
frase = frase + "     SELECT @reg = COUNT(totRem) + COUNT(totIns)" & vbCrLf
frase = frase + "         , @rem = COUNT(totRem)" & vbCrLf
frase = frase + "         , @inc = COUNT(totIns)" & vbCrLf
frase = frase + "     FROM @tblOutput" & vbCrLf
frase = frase + "     SET @FINAL = (SELECT CONVERT(bigint, rows)" & vbCrLf
frase = frase + "     From sysindexes" & vbCrLf
frase = frase + "     WHERE id = OBJECT_ID('tb_cadest')" & vbCrLf
frase = frase + "     AND indid < 2)" & vbCrLf
frase = frase + "     SELECT" & vbCrLf
frase = frase + "         @INICIAL 'INICIAL'," & vbCrLf
frase = frase + "         @REG 'LRegistros'," & vbCrLf
frase = frase + "         @INC 'LIncl'," & vbCrLf
frase = frase + "         @REM 'LRemo'," & vbCrLf
frase = frase + "         @ALT 'LAlte'," & vbCrLf
frase = frase + "         @FINAL 'LTotal'" & vbCrLf
frase = frase + " End" & vbCrLf
Set rsFunction = dbApp.Execute(frase)

frase = ""
frase = frase & " use " + gsPath_DB & vbCrLf
Set rsFunction = dbApp.Execute(frase)

frase = ""
frase = frase + " EXEC pr_setver 'pr_AtuEST_INC', '0.0.1'"
Set rsFunction = dbApp.Execute(frase)
frase = ""

End If


End Sub

Public Sub LerGlobais()

'On Error GoTo trataerr

gsstrErro_Posicao = "MyGetPar"

'OBTEM PARAMETROS DE CONEXAO COM BANCO DE DADOS DO SCEINI
gsPath_DB = MyGetPar(gsPathIniFile, "SISTEMA", "DB", "")
MySetPar gsPathIniFile, "SISTEMA", "DB", gsPath_DB

gsPath_DBCAD = MyGetPar(gsPathIniFile, "SISTEMA", "DBCAD", gsPath_DB)
MySetPar gsPathIniFile, "SISTEMA", "DBCAD", gsPath_DBCAD

gsPWD = MyGetPar(gsPathIniFile, "SISTEMA", "PWD", "parkavi")
MySetPar gsPathIniFile, "SISTEMA", "PWD", gsPWD

gsuid = MyGetPar(gsPathIniFile, "SISTEMA", "UID", "parkavi")
MySetPar gsPathIniFile, "SISTEMA", "UID", gsuid

gsPath_DS = MyGetPar(gsPathIniFile, "SISTEMA", "DS", "")
MySetPar gsPathIniFile, "SISTEMA", "DS", gsPath_DS

'VERIFICA SE O SCE.INI EH ANTIGO

'ABRE CONEXAO COM O BANCO DE DADOS
    If Not DbOpen(gsPath_DS, gsPath_DB) Then
        MsgBoxService "O banco de dados não pode ser aberto! Aplicação será finalizada!", vbCritical, "Erro"
        End
    End If
    


    
    'OBTEM PARAMETROS DO BANCO DE DADOS
    'CASO O PARAMETRO NAO EXISTA, LER DO INI E GRAVAR NO BANCO
    
    giDebug = FgetParam(gsPathIniFile, "Sistema", "SCE", "Debug", DateTime.Now, "0")
    
    gsEst_Nome = FgetParam(gsPathIniFile, "Estacionamento", "SCE", "Nome", DateTime.Now, "Nome Shopping")
    gsEst_Logo = FgetParam(gsPathIniFile, "Estacionamento", "SCE", "Logo", DateTime.Now, "Completar.gif")
    gsDelimitador = FgetParam(gsPathIniFile, "Estacionamento", "SCE", "Delimitador", DateTime.Now, ";")
    gsEst_Codigo = FgetParam(gsPathIniFile, "Estacionamento", "SCE", "Codigo", DateTime.Now, "999")
    gsEst_EntradaM = FgetParam(gsPathIniFile, "Estacionamento", "SCE", "EntradaM", DateTime.Now, "1")
    gsEst_SaidaM = FgetParam(gsPathIniFile, "Estacionamento", "SCE", "SaidaM", DateTime.Now, "51")
    gsEst_Horario = FgetParam(gsPathIniFile, "Estacionamento", "SCE", "Horario", DateTime.Now, "00:00:00")
    gsListas = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "tipoLista", DateTime.Now, "TIV")
    gbTicket = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "QTD_TICKET", DateTime.Now, "1")
    gsNextClear = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "NextClear", DateTime.Now, Format(Date + 1, "DD/MM/YYYY") + " " + "06:00:00")
    gsNextTrn = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "NextTRN", DateTime.Now, Format(Date + 1, "DD/MM/YYYY") + " " + "04:00:00")
    gsNextTrnDelta = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "NextTRNDelta", DateTime.Now, "60")
    gsNextTrnNr = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "NextTRNnr", DateTime.Now, "0")
    gsCodNelaLivre = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "CodNelaLivre", DateTime.Now, "6,99")
    gsMinutosTicketCancelados = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "MinutosTicketCancelados", DateTime.Now, "120")
    gbZeradas = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "Zeradas", DateTime.Now, "0")
    gsPermiteManuais = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "PermiteManuais", DateTime.Now, "1")
    gssceln = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "SCELN", DateTime.Now, "2")
    gsTempoBusca = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "TempoBusca", DateTime.Now, 60)
    gsscetg = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "SCETG", DateTime.Now, "2")
    gsscetr = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "SCETR", DateTime.Now, "2")
    gsPath = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "Path", DateTime.Now, "C:\SCE\")
    gsPathZ = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "PathZ", DateTime.Now, "C:\SCE\")
    gsPathZL = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "PathZL", DateTime.Now, "C:\SCE\")
    gsTrnSomenteStatCatPositivo = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "TRNSomenteStatCatPositivo", DateTime.Now, "0")
    gsSepararManuaisSemParar = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "SepararManuaisSemParar", DateTime.Now, "0")
    gsGerarTrnAE = FgetParam(gsPathIniFile, "AUTOEXPRESSO", "SCE", "GerarTrnAE", DateTime.Now, "0")
    gsSepararManuaisAE = FgetParam(gsPathIniFile, "AUTOEXPRESSO", "SCE", "SepararManuaisAE", DateTime.Now, "0")
    gsAcessoPadraoAE = FgetParam(gsPathIniFile, "AUTOEXPRESSO", "SCE", "AcessoPadraoAE", DateTime.Now, "9999")
    gsTagPadraoAE = FgetParam(gsPathIniFile, "AUTOEXPRESSO", "SCE", "TagPadraoAE", DateTime.Now, "700000000")
    gsPath_CGMP_Files_TRN = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "CGMPTRN", DateTime.Now, "c:\cgmp-edi\1.envio\trn\")
    gsPath_CGMP_Files_NEL = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "CGMPNEL", DateTime.Now, "c:\cgmp-edi\2.recebimento\nel\")
    gsPath_CGMP_Files_LNT = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "CGMPLNT", DateTime.Now, "c:\cgmp-edi\2.recebimento\lnt\")
    gsPath_CGMP_Files_TAG = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "CGMPTAG", DateTime.Now, "c:\cgmp-edi\2.recebimento\tag\")
    gsPath_CGMP_Files_TGT = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "CGMPTGT", DateTime.Now, "c:\cgmp-edi\2.recebimento\tgt\")
    gsPath_CGMP_Files_TRF = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "CGMPTRF", DateTime.Now, "c:\cgmp-edi\2.recebimento\trf\")
    gsPath_CGMP_Files_TRT = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "CGMPTRT", DateTime.Now, "c:\cgmp-edi\2.recebimento\trt\")
    gsPath_CGMP_Files_MSG = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "CGMPMSG", DateTime.Now, "c:\cgmp-edi\4.mensagens\")
    gsPath_msgbox = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "MSGBOX", DateTime.Now, "0")

    If gsPath_CGMP_Files_TRN = "c:\cgmp-edi\1.envio\trn\" Then
    bEnvioTravado = False
    Else: bEnvioTravado = True
    End If
    
    'strTest = Mid$("Visual Basic", 3, 3) 'strTest would be "ual"
    'DD/MM/YYYY"
    If Mid$(gsNextTrn, 7, 4) = Year(DateTime.Now) Then
    bGeracaoTravada = False
    Else: bGeracaoTravada = True
    End If


gsPathZ = MyGetPar(gsPathIniFile, "SISTEMA", "PathZ", "")

If gsPathZ <> "" Then
    'INI ANTIGO, RENOMEAR PARA SCE_OLDBKP.INI
    FileCopy gsPath + "\SCE.ini", gsPath + "\SCE_OLDBKP_" + Format(Now, "yyyyMMddHHmmss") + ".ini"
    
    'DELETA O INI ANTIGO
    Kill gsPath + "\SCE.ini"
    
    'CRIAR NOVO INI
    If Dir(gsPath + "\SCE.ini", vbArchive) = "" Then
        Dim i As Integer
        i = FreeFile
        Open gsPath + "\SCE.ini" For Output As i
        Print #i, "***** Arquivo INI criado em " + Format(Now, "DD/MM/YY hh:mm:ss")
        Close i
        
    
        MySetPar gsPathIniFile, "SISTEMA", "PWD", gsPWD
        MySetPar gsPathIniFile, "SISTEMA", "UID", gsuid
        MySetPar gsPathIniFile, "SISTEMA", "DS", gsPath_DS
        MySetPar gsPathIniFile, "SISTEMA", "DBCAD", gsPath_DBCAD
        MySetPar gsPathIniFile, "SISTEMA", "DB", gsPath_DB
        
    End If
End If


'CRIA OS DIRETORIOS NECESSARIOS PARA GERENCIAMENTO DE ARQUIVOS
gsPathZ = FgetParam(gsPathIniFile, "SISTEMA", "SCE", "PathZ", DateTime.Now, "C:\SCE\")
gsPath = CriaDir(gsPath, "", "", "")
gsPath_REL = CriaDir(gsPath, "Relatorios\", "", "")
gsPath_SCEExtra = CriaDir(gsPath, "SCEExtra\", "", "")
gsPath_CGMPEnvia = CriaDir(gsPath, "CGMPEnvia\", "", "")
gsPath_CGMPRecebe = CriaDir(gsPath, "CGMPRecebe\", "", "")
gsPath_SCE_Files = CriaDir(gsPath, "Arquivos\", "", "")
gsPath_LN = CriaDir(gsPath, "Arquivos\", "LN\", "")
gsPath_CD = CriaDir(gsPath, "Arquivos\", "CD\", "")
gsPath_TRT = CriaDir(gsPath, "Arquivos\", "TRT\", "")
gsPath_OUTROS = CriaDir(gsPath, "Arquivos\", "OUTROS\", "")
    
Exit Sub

trataerr:
    Call TrataErro(app.title, Error, "LerGlobais")


End Sub


Function Busca_Indice_Combo(cmbCombo As ComboBox, strItem As String)
'** Verifica se item digitado esta no combo. Caso positivo retorna o ListIndex
    Dim i As Byte
    Dim bytPosicao As Byte
    Dim blnAchou As Boolean
    
    Busca_Indice_Combo = -1
    
    If cmbCombo.ListCount = 0 Then
        Exit Function
    End If
    
    'garanto que vira' uma string para a comparacao
    strItem = Trim(strItem & "")
    
    For i = 0 To cmbCombo.ListCount - 1
        bytPosicao = InStr(1, cmbCombo.List(i), strItem)
        If bytPosicao <> 0 Then
            blnAchou = Left(cmbCombo.List(i), bytPosicao - 1) = Left(strItem, bytPosicao)
        Else
            blnAchou = cmbCombo.List(i) = strItem
        End If
        
        If blnAchou Then
            Busca_Indice_Combo = i
            Exit For
        End If
    Next i
End Function

Public Sub LimparArquivos(Optional opc As Integer)
On Error GoTo trataerr
Dim fraseaux As String
Dim frase As String
Dim rs As New Recordset
Dim rsaux As New Recordset
Dim aux As String



If Now > CVDate(gsNextClear) And gsNextClear <> "01/01/1980 00:00:00" Then

    gsNextClear = CVDate(Format(Now, "dd/mm/yyyy")) + 1
    'MySetPar gsPathIniFile, "SISTEMA", "NextClear", Format(gsNextClear, "DD/MM/YYYY hh:mm:ss")
    
    frase = ""
    
    frase = frase & "UPDATE " + gsPath_DB + ".dbo.tb_appparam SET szvalor ='" + Format(gsNextClear, "DD/MM/YYYY hh:mm:ss") + "'WHERE szparam LIKE '%NextClear'"
    
    'frase = frase & " UPDATE " + gsPath_DB + ".dbo.tb_appparam SET szvalor = '" + gsNextClear + "' WHERE szparam LIKE '%NextClear'"
    Set rs = dbApp.Execute(frase)

'    frase = ""
'    frase = frase + " dbcc dbreindex (tb_transacao) "
'    Set rsGeral = dbApp.Execute(frase)

End If


gbTRNOK = True

'End If

Exit Sub

'On Error GoTo trataerr
trataerr:
Call TrataErro(app.title, Error, "LimpaArquivos")

gbTRNOK = False

  
End Sub




Public Sub TrataETT()
Dim Datafile As String
Dim Datafilenum As Long
Dim rs As New Recordset
Dim frase As String
Dim i As Long
Dim Header As CDHeader
Dim file As String

On Error GoTo trataerr
rs.CursorType = adOpenStatic

'VERIFICA SE A LINHA REFERENTE AO ARQUIVO JA FOI INSERIDA NA TB_CADESTCTL
frase = ""
frase = frase & "SELECT top 1 szArquivo, ISNULL(lseqfile,0) , ctipo FROM " + gsPath_DBCAD + ".dbo.TB_CADESTCTL "
frase = frase & "WHERE tsatualizacao = '19800101' AND ctipo = 'TT' "
frase = frase & "ORDER BY szarquivo DESC, lseqfile DESC, ctipo DESC"
Set rs = dbApp.Execute(frase)
If rs.BOF And rs.EOF Then
    Exit Sub
End If
file = rs(0)

Datafile = gsPath_CGMPRecebe & file
Datafilenum = FreeFile
Open Datafile For Binary As Datafilenum
'LE header
Get #Datafilenum, 1, Header

'VERIFICA O ULTIMO ARQUIVO PROCESSADO NA CADESTCTL
frase = ""
frase = frase & "SELECT TOP 1 szArquivo, ISNULL(lseqfile,0) FROM " + gsPath_DBCAD + ".dbo.TB_CADESTCTL "
frase = frase & "WHERE tsatualizacao <> '19800101' AND (ctipo = 'TT' OR ctipo = 'TG') "
frase = frase & "ORDER BY lseqfile DESC, ctipo DESC"
Set rs = dbApp.Execute(frase)


MDI.sta_Barra_MDI.Panels(2).Text = "Tratando Arquivo : " + Datafile


If Not rs.BOF And Not rs.EOF Then

    'VERIFICA SE A SEQUENCIA ATUAL EH MENOR QUE A ULTIMA PROCESSADA NO BANCO DE DADOS
    If Val(Header.seq) < Val(rs(1)) Then
        Close Datafilenum
        
       'MOVE ARQUIVO PARA A PASTA OUTROS
       gsDestFile = gsPath_OUTROS & file
       Call CopyFile(Datafile, gsDestFile)
       Kill Datafile

       'GRAVA EVENTO
       Call GravaEventos(403, " ", gsUser, 0, "Trata ETT - " + file + " Abaixo da Sequencia ", Format(Header.seq), Right(Format(Header.tag), 9))
       
       'DELETA O REGISTRO DA TB_CADESTCTL
       frase = ""
       frase = frase & "DELETE FROM " + gsPath_DBCAD + ".dbo.TB_CADESTCTL "
       frase = frase & "WHERE tsatualizacao = '19800101' and szarquivo = '" & file & "'"
       Set rs = dbApp.Execute(frase)
       Exit Sub
    End If
End If

'SE NAO EXISTIR A TABELA TB_CADEST_TMP, CRIA
frase = ""
frase = " if NOT exists (select * from " + gsPath_DBCAD + "..sysobjects where id = object_id(N'" + gsPath_DBCAD + ".DBO.tb_cadest_tmp')) "
frase = frase & "SELECT * INTO " + gsPath_DBCAD + ".dbo.TB_CADEST_TMP from " + gsPath_DBCAD + ".dbo.TB_CADEST where 0=1"
Set rs = dbApp.Execute(frase)

'TRUNCA A TABELA TB_CADEST_TMP
frase = ""
frase = frase & " truncate table " + gsPath_DBCAD + ".dbo.TB_CADEST_TMP "
Set rs = dbApp.Execute(frase)

Close Datafilenum
Datafilenum = FreeFile

'COPIA O ARQUIVO PARA O SERVIDOR DO BANCO DE DADOS PATHZ
MDI.sta_Barra_MDI.Panels(2).Text = "Copiando Arquivo : " + file
gsSourceFile = gsPath_CGMPRecebe & file
gsDestFile = gsPathZ & file
Call CopyFile(gsSourceFile, gsDestFile)


'EXECUTA O BULK COPY
frase = ""
frase = frase & " BULK INSERT " + gsPath_DBCAD + ".dbo.TB_CADest_tmp"
frase = frase & " From '" & gsPathZL & file & " '"
frase = frase & " WITH (FORMATFILE = '" & gsPathZL & "SCEExtra\cadest_tmp_bcp.fmt',FIRSTROW = 2);"
Set rs = dbApp.Execute(frase)

MDI.sta_Barra_MDI.Panels(2).Text = "Apagando Arquivo : " + file
Kill gsDestFile

'VERIFICA SE O BULK COPY EFETUOU O INSERT DOS REGISTROS NA TABELA TEMPORARIA
frase = ""
frase = frase & "SELECT count(*) from " + gsPath_DBCAD + ".dbo.TB_CADEST_TMP "
Set rs = dbApp.Execute(frase)

'SE A QUANTIDADE DE REGISTROS FOR MAIOR QUE ZERO E IGUAL AO HEADER DO ARQUIVO
If rs(0) > 0 And rs(0) = Header.tag + 0 Then
    MDI.sta_Barra_MDI.Panels(2).Text = "Atualizando Tabelas"
    
    'CHAMA A PROC ATUETT PARA PROCESSAMENTO DOS ARQUIVOS
    frase = gsPath_DBCAD + ".dbo.pr_atuett " + Header.seq
    Set rs = dbApp.Execute(frase)
    
    'ATUALIZA A TABELA TB_CADESTCTL
    frase = ""
    frase = frase & "UPDATE " + gsPath_DBCAD + ".dbo.TB_CADESTCTL SET "
    frase = frase & "lseqfile = '" & Header.seq & "',"
    frase = frase & "lregistros = 0,"
    frase = frase & "ltotal = '" & Val(Header.tag) & "',"
    frase = frase & "tsatualizacao = '" & Format(Now, "YYYYMMDD HH:MM:SS") & "',"
    frase = frase & "lincl = 0,"
    frase = frase & "lremo = 0,"
    frase = frase & "lalte = 0 "
    frase = frase & "WHERE szArquivo = '" & file & "' "
    Set rs = dbApp.Execute(frase)
    MDI.sta_Barra_MDI.Panels(2).Text = "Transferindo Arquivo : " + Datafile
    
    'GRAVA EVENTO
    Call GravaEventos(403, " ", gsUser, 0, "Trata ETT - " + file, 0, 0)
    
    'MOVE ARQUIVO PARA A PASTA ATAUALIZADOS
    gsDestFile = CriaDir(gsPath_CD, "Atualizados\", Mid(file, 1, 6) + "\", "") & file
    Close Datafilenum
    Call CopyFile(Datafile, gsDestFile)
    Kill Datafile
Else
    MsgBoxService "Trata Cadastro Total > Quantidade de Registro Errados"
End If
Exit Sub

trataerr:
Call TrataErro("Cadastro de TAG ", Error, " TrataTGT ")
gbTransferStatus = False
End Sub




Public Sub TrataTGT()
Dim Datafile As String
Dim Datafilenum As Long
Dim rs As New Recordset
Dim frase As String
Dim i As Long
Dim Header As CDHeader
Dim file As String

On Error GoTo trataerr
rs.CursorType = adOpenStatic

frase = ""
frase = frase & "SELECT top 1 szArquivo, ISNULL(lseqfile,0) , ctipo FROM " + gsPath_DBCAD + ".dbo.TB_CADtagCtl "
frase = frase & "WHERE tsatualizacao = '19800101' AND ctipo = 'TT' "
frase = frase & "ORDER BY szarquivo DESC, lseqfile DESC, ctipo DESC"
Set rs = dbApp.Execute(frase)
If rs.BOF And rs.EOF Then
    Exit Sub
End If
file = rs(0)

Datafile = gsPath_CGMPRecebe & file
Datafilenum = FreeFile
Open Datafile For Binary As Datafilenum
'LE header
Get #Datafilenum, 1, Header

frase = ""
frase = frase & "SELECT TOP 1 szArquivo, ISNULL(lseqfile,0) FROM " + gsPath_DBCAD + ".dbo.TB_CADtagCtl "
frase = frase & "WHERE tsatualizacao <> '19800101' AND (ctipo = 'TT' OR ctipo = 'TG') "
frase = frase & "ORDER BY lseqfile DESC, ctipo DESC"
Set rs = dbApp.Execute(frase)
MDI.sta_Barra_MDI.Panels(2).Text = "Tratando Arquivo : " + Datafile
If Not rs.BOF And Not rs.EOF Then
    If Val(Header.seq) < Val(rs(1)) Then
        Close Datafilenum
        
       gsDestFile = gsPath_OUTROS & file
       Call CopyFile(Datafile, gsDestFile)
       Kill Datafile

       Call GravaEventos(403, " ", gsUser, 0, "Trata TGT - " + file + " Abaixo da Sequencia ", Format(Header.seq), Right(Format(Header.tag), 9))
       frase = ""
       frase = frase & "DELETE FROM " + gsPath_DBCAD + ".dbo.TB_CADtagCtl "
       frase = frase & "WHERE tsatualizacao = '19800101' and szarquivo = '" & file & "'"
       Set rs = dbApp.Execute(frase)
       Exit Sub
    End If
End If

frase = ""
frase = " if NOT exists (select * from " + gsPath_DBCAD + "..sysobjects where id = object_id(N'" + gsPath_DBCAD + ".DBO.tb_cadtag_tmp')) "
frase = frase & "SELECT * INTO " + gsPath_DBCAD + ".dbo.TB_CADtag_tmp from " + gsPath_DBCAD + ".dbo.TB_CADtag where 0=1"
Set rs = dbApp.Execute(frase)

frase = ""
frase = frase & " truncate table " + gsPath_DBCAD + ".dbo.TB_CADtag_tmp "
Set rs = dbApp.Execute(frase)

Close Datafilenum
Datafilenum = FreeFile

MDI.sta_Barra_MDI.Panels(2).Text = "Copiando Arquivo : " + file
gsSourceFile = gsPath_CGMPRecebe & file
gsDestFile = gsPathZ & file
Call CopyFile(gsSourceFile, gsDestFile)



frase = ""
frase = frase & " BULK INSERT " + gsPath_DBCAD + ".dbo.TB_CADtag_tmp"
frase = frase & " From '" & gsPathZL & file & " '"
frase = frase & " WITH (FORMATFILE = '" & gsPathZL & "SCEExtra\cadtag_tmp_bcp.fmt',FIRSTROW = 2);"
Set rs = dbApp.Execute(frase)

MDI.sta_Barra_MDI.Panels(2).Text = "Apagando Arquivo : " + file
'gsdestfile = gsPathZ & mynames(i)
Kill gsDestFile

frase = ""
frase = frase & "SELECT count(*) from " + gsPath_DBCAD + ".dbo.TB_CADtag_tmp "
Set rs = dbApp.Execute(frase)

If rs(0) > 0 And rs(0) = Header.tag + 0 Then
    MDI.sta_Barra_MDI.Panels(2).Text = "Atualizando Tabelas"
    frase = gsPath_DBCAD + ".dbo.pr_atutgt " + Header.seq
    Set rs = dbApp.Execute(frase)
    'Atualiza cadctl
    frase = ""
    frase = frase & "UPDATE " + gsPath_DBCAD + ".dbo.TB_CADtagCtl SET "
    frase = frase & "lseqfile = '" & Header.seq & "',"
    frase = frase & "lregistros = 0,"
    frase = frase & "ltotal = '" & Val(Header.tag) & "',"
    frase = frase & "tsatualizacao = '" & Format(Now, "YYYYMMDD HH:MM:SS") & "',"
    frase = frase & "lincl = 0,"
    frase = frase & "lremo = 0,"
    frase = frase & "lalte = 0 "
    frase = frase & "WHERE szArquivo = '" & file & "' "
    Set rs = dbApp.Execute(frase)
    MDI.sta_Barra_MDI.Panels(2).Text = "Transferindo Arquivo : " + Datafile
    Call GravaEventos(403, " ", gsUser, 0, "Trata TG - " + file, 0, 0)
    gsDestFile = CriaDir(gsPath_CD, "Atualizados\", Mid(file, 1, 6) + "\", "") & file
    Close Datafilenum
    Call CopyFile(Datafile, gsDestFile)
    Kill Datafile
Else
    MsgBoxService "Trata Cadastro Total > Quantidade de Registro Errados"
End If
Exit Sub

trataerr:
Call TrataErro("Cadastro de TAG ", Error, " TrataTGT ")
gbTransferStatus = False
End Sub

Public Sub TrataLCT()
Dim Datafile As String
Dim Datafilenum As Long
Dim rs As New Recordset
Dim frase As String
Dim i As Long
Dim Header As LCHeader
Dim file As String

On Error GoTo trataerr
rs.CursorType = adOpenStatic

frase = ""
frase = frase & "SELECT top 1 szArquivo, ISNULL(lseqfile,0) , ctipo FROM " + gsPath_DBCAD + ".dbo.TB_ComboCtl "
frase = frase & "WHERE tsatualizacao = '19800101' AND ctipo = 'TT' "
frase = frase & "ORDER BY szarquivo DESC, lseqfile DESC, ctipo DESC"
Set rs = dbApp.Execute(frase)
If rs.BOF And rs.EOF Then
    Exit Sub
End If
file = rs(0)

Datafile = gsPath_CGMPRecebe & file
Datafilenum = FreeFile
Open Datafile For Binary As Datafilenum
'LE header
Get #Datafilenum, 1, Header

frase = ""
frase = frase & "SELECT TOP 1 szArquivo, ISNULL(lseqfile,0) FROM " + gsPath_DBCAD + ".dbo.TB_ComboCtl "
frase = frase & "WHERE tsatualizacao <> '19800101' AND (ctipo = 'TT' OR ctipo = 'TG') "
frase = frase & "ORDER BY lseqfile DESC, ctipo DESC"
Set rs = dbApp.Execute(frase)
MDI.sta_Barra_MDI.Panels(2).Text = "Tratando Arquivo : " + Datafile
If Not rs.BOF And Not rs.EOF Then
    If Val(Header.seq) < Val(rs(1)) Then
       Close Datafilenum
       gsDestFile = gsPath_OUTROS & file
       Call CopyFile(Datafile, gsDestFile)
       Kill Datafile
       Call GravaEventos(403, " ", gsUser, 0, "Trata LCT - " + file + " Abaixo da Sequencia ", Format(Header.seq), Right(Format(Header.tag), 9))
       frase = ""
       frase = frase & "DELETE FROM " + gsPath_DBCAD + ".dbo.TB_comboCtl "
       frase = frase & "WHERE tsatualizacao = '19800101' and szarquivo = '" & file & "'"
       Set rs = dbApp.Execute(frase)
       Exit Sub
    End If
End If

frase = ""
frase = " if NOT exists (select * from " + gsPath_DBCAD + "..sysobjects where id = object_id(N'" + gsPath_DBCAD + ".DBO.tb_combo_tmp')) "
frase = frase & "SELECT * INTO " + gsPath_DBCAD + ".dbo.TB_combo_tmp from " + gsPath_DBCAD + ".dbo.TB_combo where 0=1"
Set rs = dbApp.Execute(frase)

frase = ""
frase = frase & " truncate table " + gsPath_DBCAD + ".dbo.TB_combo_tmp "
Set rs = dbApp.Execute(frase)

Close Datafilenum
Datafilenum = FreeFile

MDI.sta_Barra_MDI.Panels(2).Text = "Copiando Arquivo : " + file
gsSourceFile = gsPath_CGMPRecebe & file
gsDestFile = gsPathZ & file
Call CopyFile(gsSourceFile, gsDestFile)



frase = ""
frase = frase & " BULK INSERT " + gsPath_DBCAD + ".dbo.TB_combo_tmp"
frase = frase & " From '" & gsPathZL & file & " '"
frase = frase & " WITH (FORMATFILE = '" & gsPathZL & "SCEExtra\LCT.fmt',FIRSTROW = 2);"
Set rs = dbApp.Execute(frase)

MDI.sta_Barra_MDI.Panels(2).Text = "Apagando Arquivo : " + file
'gsdestfile = gsPathZ & mynames(i)
Kill gsDestFile

frase = ""
frase = frase & "SELECT count(*) from " + gsPath_DBCAD + ".dbo.TB_Combo_tmp "
Set rs = dbApp.Execute(frase)

If rs(0) > 0 Then
    MDI.sta_Barra_MDI.Panels(2).Text = "Atualizando Tabelas"
    frase = gsPath_DBCAD + ".dbo.pr_atuLCT " + Header.seq
    Set rs = dbApp.Execute(frase)
    'Atualiza cadctl
    frase = ""
    frase = frase & "UPDATE " + gsPath_DBCAD + ".dbo.TB_ComboCtl SET "
    frase = frase & "lseqfile = '" & Header.seq & "',"
    frase = frase & "lregistros = 0,"
    frase = frase & "ltotal = '" & Val(Header.tag) & "',"
    frase = frase & "tsatualizacao = '" & Format(Now, "YYYYMMDD HH:MM:SS") & "',"
    frase = frase & "lincl = 0,"
    frase = frase & "lremo = 0,"
    frase = frase & "lalte = 0 "
    frase = frase & "WHERE szArquivo = '" & file & "' "
    Set rs = dbApp.Execute(frase)
    MDI.sta_Barra_MDI.Panels(2).Text = "Transferindo Arquivo : " + Datafile
    Call GravaEventos(403, " ", gsUser, 0, "Trata TG - " + file, 0, Right(Format(Header.tag), 9))
    gsDestFile = CriaDir(gsPath_CD, "Atualizados\", Mid(file, 1, 6) + "\", "") & file
    Close Datafilenum
    Call CopyFile(Datafile, gsDestFile)
    Kill Datafile
Else
    MsgBoxService "Trata Cadastro Total > Quantidade de Registro Errados"
End If
Exit Sub

trataerr:
Call TrataErro("Cadastro de TAG ", Error, " TrataTGT ")
gbTransferStatus = False
End Sub

Public Sub TrataLC(file As String)

Dim Datafile As String
Dim Datafilenum As Long
Dim rs As New Recordset
Dim frase As String
Dim oldlista As Long
Dim i As Long
Dim Header As LCHeader

On Error GoTo trataerr
rs.CursorType = adOpenStatic

Datafile = gsPath_CGMPRecebe & file
Datafilenum = FreeFile
Open Datafile For Binary As Datafilenum

'LE header
Get #Datafilenum, 1, Header

frase = ""
frase = frase & "SELECT isnull(max(lseqfile),0) FROM " + gsPath_DBCAD + ".dbo.TB_comboCtl "
frase = frase & "WHERE tsatualizacao <> '19800101' "
Set rs = dbApp.Execute(frase)
If Not rs.BOF And Not rs.EOF Then
    oldlista = Val(rs(0))
    If Val(Header.seq) <= oldlista Then
       Close Datafilenum
       gsDestFile = gsPath_OUTROS & file
       Call CopyFile(Datafile, gsDestFile)
       Kill Datafile
       Call GravaEventos(403, " ", gsUser, 0, "Trata LC - " + file + " Abaixo da Sequencia ", Format(Header.seq), Right(Format(Header.tag), 9))
       frase = ""
       frase = frase & "DELETE FROM " + gsPath_DBCAD + ".dbo.TB_ComboCtl "
       frase = frase & "WHERE tsatualizacao = '19800101' and szarquivo = '" & file & "'"
       Set rs = dbApp.Execute(frase)
       
       Exit Sub
    Else
        If Val(Header.seq) > oldlista + 1 Then
            frase = ""
            frase = frase & "DELETE FROM " + gsPath_DBCAD + ".dbo.TB_ComboCtl "
            frase = frase & "WHERE tsatualizacao = '19800101' and szarquivo = '" & file & "'"
            Set rs = dbApp.Execute(frase)
            Close Datafilenum
            frase = ""
            frase = " FALTA SEQUENCIAL DE CADASTRO :: ABRIR FALHA NO HELP DESK - "
            frase = frase + "Sequencia LCI : " + Format(oldlista + 1) + "  ATÉ " + Format(Val(Header.seq) - 1)
            
            Call LogErro("TrataLC - LISTAS FALTANTES", frase)
            
            gsListasFaltantes = Format(oldlista + 1) + "-" + Format(Val(Header.seq) - 1)
            
            MySetPar gsPathLogListas, "LISTAS", "LCI", gsListasFaltantes
            
            GSSequencia = False
            Exit Sub
        End If
    End If
End If
i = 0
MDI!sta_Barra_MDI.Panels(2).Text = "Atualizando TG : " + Header.tag
frase = ""
frase = frase & " use " + gsPath_DBCAD
frase = frase & " if not exists (select * from " + gsPath_DBCAD + "..sysobjects where id = object_id(N'[tb_combo_inc]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) "
frase = frase & "SELECT * INTO " + gsPath_DBCAD + ".dbo.TB_combo_inc from " + gsPath_DBCAD + ".dbo.TB_combo where 0=1"
Set rs = dbApp.Execute(frase)

Close Datafilenum
Datafilenum = FreeFile

frase = ""
frase = frase & " truncate table " + gsPath_DBCAD + ".dbo.TB_Combo_inc "
Set rs = dbApp.Execute(frase)

MDI.sta_Barra_MDI.Panels(2).Text = "Copiando Arquivo : " + file
gsSourceFile = gsPath_CGMPRecebe & file
gsDestFile = gsPathZ & file
Call CopyFile(gsSourceFile, gsDestFile)

Espera (10)

frase = ""
frase = frase & " BULK INSERT " + gsPath_DBCAD + ".dbo.TB_Combo_inc"
frase = frase & " From '" & gsPathZL & file & " '"
frase = frase & " WITH (FORMATFILE = '" & gsPathZL & "SCEExtra\lct.fmt',FIRSTROW = 2);"
Set rs = dbApp.Execute(frase)

Espera (10)

MDI.sta_Barra_MDI.Panels(2).Text = "Apagando Arquivo : " + file
'gsdestfile = gsPathZ & mynames(i)
Kill gsDestFile

frase = ""
frase = frase & "SELECT count(*) from " + gsPath_DBCAD + ".dbo.TB_Combo_inc "
Set rs = dbApp.Execute(frase)

If rs(0) > 0 Then
    MDI.sta_Barra_MDI.Panels(2).Text = "Atualizando Tabelas"
    frase = ""
    frase = frase & "pr_atulci_INC " + Header.seq
    Set rs = dbApp.Execute(frase)
    'Atualiza cadctl
    frase = ""
    frase = frase & "UPDATE " + gsPath_DBCAD + ".dbo.TB_comboCtl SET "
    frase = frase & "lseqfile = '" & Val(Header.seq) & "',"
    frase = frase & "lregistros = '" + Format(rs(1)) + "' ,"
    frase = frase & "ltotal = '" + Format(rs(5)) + "' ,"
    frase = frase & "tsatualizacao = '" & Format(Now, "YYYYMMDD HH:MM:SS") & "',"
    frase = frase & "lincl = '" + Format(rs(2)) + "' ,"
    frase = frase & "lremo = '" + Format(rs(3)) + "' ,"
    frase = frase & "lalte = '" + Format(rs(4)) + "'"
    frase = frase & "WHERE szArquivo = '" & file & "' "
    Set rs = dbApp.Execute(frase)
    MDI.sta_Barra_MDI.Panels(2).Text = "Transferindo Arquivo : " + Datafile
    Call GravaEventos(403, " ", gsUser, 0, "Trata LC - " + file, Format(Header.seq), Right(Format(Header.tag), 9))
    gsDestFile = CriaDir(gsPath_CD, "Atualizados\", Mid(file, 1, 6) + "\", "") & file
    Close Datafilenum
    Call CopyFile(Datafile, gsDestFile)
    Kill Datafile
    
    gsListasFaltantes = "0-0"
    MySetPar gsPathLogListas, "LISTAS", "LCI", gsListasFaltantes
    
'    'Atualiza usertag preenchendo a placa quando estiver vazia
'    frase = ""
'    frase = frase & " Update  " + gsPath_DB + ".dbo.tb_usertag"
'    frase = frase & " Set Cplaca = c.Cplaca"
'    frase = frase & " from " + gsPath_DB + ".dbo.tb_UserTag u left join " + gsPath_DB + ".dbo.tb_cadtag c on (c.LTag = u.Ltag)"
'    frase = frase & " Where u.Cplaca Is Null"
'    Set rs = dbApp.Execute(frase)
'
'    'Atualiza usertag com o TIV quando houver o TAG
'    frase = ""
'    frase = frase & " Update " + gsPath_DB + ".dbo.tb_usertag"
'    frase = frase & " set iIssuer=c.IIssuer,lTag=c.LTag"
'    frase = frase & " from " + gsPath_DB + ".dbo.tb_usertag u left join " + gsPath_DB + ".dbo.tb_cadtag c on (c.CPLACA=u.cPlaca)"
'    frase = frase & " Where u.LTag < c.LTag"
'    Set rs = dbApp.Execute(frase)
    
Else
    MsgBoxService "Trata Cadastro LC > Quantidade de Registro Errados"
End If

Exit Sub

trataerr:
Call TrataErro("Cadastro de LC ", Error, " TrataLC")
gbTransferStatus = False

End Sub
Public Sub TrataTG(file As String)

Dim Datafile As String
Dim Datafilenum As Long
Dim rs As New Recordset
Dim frase As String
Dim oldlista As Long
Dim i As Long
Dim Header As CDHeader

On Error GoTo trataerr
rs.CursorType = adOpenStatic

Datafile = gsPath_CGMPRecebe & file
Datafilenum = FreeFile
Open Datafile For Binary As Datafilenum

'LE header
Get #Datafilenum, 1, Header

frase = ""
frase = frase & "SELECT isnull(max(lseqfile),0) FROM " + gsPath_DBCAD + ".dbo.TB_CADtagCtl "
frase = frase & "WHERE tsatualizacao <> '19800101' "
Set rs = dbApp.Execute(frase)
If Not rs.BOF And Not rs.EOF Then
    oldlista = Val(rs(0))
    If Val(Header.seq) <= oldlista Then
       Close Datafilenum
       gsDestFile = gsPath_OUTROS & file
       Call CopyFile(Datafile, gsDestFile)
       Kill Datafile
       Call GravaEventos(403, " ", gsUser, 0, "Trata TG - " + file + " Abaixo da Sequencia ", Format(Header.seq), Right(Format(Header.tag), 9))
       frase = ""
       frase = frase & "DELETE FROM " + gsPath_DBCAD + ".dbo.TB_CADtagctl "
       frase = frase & "WHERE tsatualizacao = '19800101' and szarquivo = '" & file & "'"
       Set rs = dbApp.Execute(frase)
       Exit Sub
    Else
        If Val(Header.seq) > oldlista + 1 Then
            frase = ""
            frase = frase & "DELETE FROM " + gsPath_DBCAD + ".dbo.TB_CADtagctl "
            frase = frase & "WHERE tsatualizacao = '19800101' and szarquivo = '" & file & "'"
            Set rs = dbApp.Execute(frase)
            Close Datafilenum
            frase = ""
            frase = " FALTA SEQUENCIAL DE CADASTRO :: ABRIR FALHA NO HELP DESK "
            frase = frase + "Sequencia TAG : " + Format(oldlista + 1) + "  ATÉ " + Format(Val(Header.seq) - 1)
            'i = MyMsgBox(frase, 0, "LISTAS FALTANTES", "")
            Call LogErro("TrataTG - LISTAS FALTANTES", frase)
            
            gsListasFaltantes = Format(oldlista + 1) + "-" + Format(Val(Header.seq) - 1)
            MySetPar gsPathLogListas, "LISTAS", "TIV", gsListasFaltantes
            
            GSSequencia = False
            Exit Sub
        End If
    End If
End If

i = 0
MDI!sta_Barra_MDI.Panels(2).Text = "Atualizando TG : " + Header.tag
frase = ""
frase = frase & " use " + gsPath_DBCAD
frase = frase & " if not exists (select * from " + gsPath_DBCAD + "..sysobjects where id = object_id(N'[tb_cadtag_inc]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) "
frase = frase & "SELECT * INTO " + gsPath_DBCAD + ".dbo.TB_CADtag_inc from " + gsPath_DBCAD + ".dbo.TB_CADtag where 0=1"
Set rs = dbApp.Execute(frase)

Close Datafilenum
Datafilenum = FreeFile

frase = ""
frase = frase & " truncate table " + gsPath_DBCAD + ".dbo.TB_CADtag_inc "
Set rs = dbApp.Execute(frase)

MDI.sta_Barra_MDI.Panels(2).Text = "Copiando Arquivo : " + file
gsSourceFile = gsPath_CGMPRecebe & file
gsDestFile = gsPathZ & file
Call CopyFile(gsSourceFile, gsDestFile)

Espera (5)

frase = ""
frase = frase & " BULK INSERT " + gsPath_DBCAD + ".dbo.TB_CADtag_inc"
frase = frase & " From '" & gsPathZL & file & " '"
frase = frase & " WITH (FORMATFILE = '" & gsPathZL & "SCEExtra\cadtag_tmp_bcp.fmt',FIRSTROW = 2);"
Set rs = dbApp.Execute(frase)

MDI.sta_Barra_MDI.Panels(2).Text = "Apagando Arquivo : " + file
'gsdestfile = gsPathZ & mynames(i)
Kill gsDestFile


frase = ""
frase = frase & "SELECT count(*) from " + gsPath_DBCAD + ".dbo.TB_CADtag_inc "
Set rs = dbApp.Execute(frase)

If rs(0) > 0 Then
    MDI.sta_Barra_MDI.Panels(2).Text = "Atualizando Tabelas"
    frase = ""
    frase = frase & "pr_atuTAG_INC " + Header.seq
    Set rs = dbApp.Execute(frase)
    'Atualiza cadctl
    frase = ""
    frase = frase & "UPDATE " + gsPath_DBCAD + ".dbo.TB_CADtagCtl SET "
    frase = frase & "lseqfile = '" & Val(Header.seq) & "',"
    frase = frase & "lregistros = '" + Format(rs(1)) + "' ,"
    frase = frase & "ltotal = '" + Format(rs(5)) + "' ,"
    frase = frase & "tsatualizacao = '" & Format(Now, "YYYYMMDD HH:MM:SS") & "',"
    frase = frase & "lincl = '" + Format(rs(2)) + "' ,"
    frase = frase & "lremo = '" + Format(rs(3)) + "' ,"
    frase = frase & "lalte = '" + Format(rs(4)) + "'"
    frase = frase & "WHERE szArquivo = '" & file & "' "
    Set rs = dbApp.Execute(frase)
    MDI.sta_Barra_MDI.Panels(2).Text = "Transferindo Arquivo : " + Datafile
    Call GravaEventos(403, " ", gsUser, 0, "Trata TG - " + file, Format(Header.seq), Right(Format(Header.tag), 9))
    gsDestFile = CriaDir(gsPath_CD, "Atualizados\", Mid(file, 1, 6) + "\", "") & file
    Close Datafilenum
    Call CopyFile(Datafile, gsDestFile)
    Kill Datafile
    
    gsListasFaltantes = "0-0"
    MySetPar gsPathLogListas, "LISTAS", "TIV", gsListasFaltantes
    
Else
    MsgBoxService "Trata Cadastro TAG > Quantidade de Registro Errados"
End If

Exit Sub

trataerr:
Call TrataErro("Cadastro de TAG ", Error, " TrataTG")
gbTransferStatus = False

End Sub


Sub TrataFileLC()
Dim i As Integer
Dim myname As String
Dim TrataOK As Integer

On Error GoTo trataerr


If TrataOK = 0 Then
    MDI.sta_Barra_MDI.Panels(2).Text = "Procurando Arquivo LCT"
    i = 0
    ReDim mynames(i)
    myname = UCase(Dir(gsPath_CGMPRecebe & "*.LCT"))   ' Retrieve the first entry.
    Do While myname <> ""   ' Start the loop.
        If Right(myname, 4) = ".LCT" Then
            mynames(i) = UCase(myname)
            i = i + 1
            ReDim Preserve mynames(i)
        End If
        myname = UCase(Dir)
    Loop
    For i = UBound(mynames) - 1 To 0 Step -1
        If Right(mynames(i), 4) = ".LCT" Then
            DoEvents
            MDI.sta_Barra_MDI.Panels(2).Text = "Abrindo Arquivo : " + mynames(i)
            frase = ""
            frase = frase & "SELECT COUNT(*) FROM " + gsPath_DBCAD + ".dbo.TB_ComboCtl "
            frase = frase & "WHERE szArquivo = '" + mynames(i) + "' "
            Set rs = dbApp.Execute(frase)
            If rs(0) = 0 Then
                frase = ""
                frase = frase & "INSERT INTO " + gsPath_DBCAD + ".dbo.TB_ComboCtl (szarquivo,ctipo,tsatualizacao) values "
                frase = frase & "('" & mynames(i) & "','TT','19800101')"
                Set rs = dbApp.Execute(frase)
                Call TrataLCT
                TrataOK = 1
             Else
                MDI.sta_Barra_MDI.Panels(2).Text = "Transferindo Arquivo : " + mynames(i)
                gsSourceFile = gsPath_CGMPRecebe & mynames(i)
                gsDestFile = gsPath_OUTROS & mynames(i)
                Call CopyFile(gsSourceFile, gsDestFile)
                Kill gsSourceFile
            End If

        End If
    Next
End If
DoEvents
If TrataOK = 0 Then
    i = 0
    ReDim mynames(i)
    myname = UCase(Dir(gsPath_CGMPRecebe & "*.LCI"))   ' Retrieve the first entry.
    Do While myname <> ""   ' Start the loop.
        If UCase(Right(myname, 4)) = ".LCI" Then
            mynames(i) = UCase(myname)
            i = i + 1
            ReDim Preserve mynames(i)
        End If
        myname = UCase(Dir)
    Loop

    For i = 0 To UBound(mynames) - 1
        If UCase(Right(mynames(i), 4)) = ".LCI" And TrataOK = 0 Then
           DoEvents
           MDI.sta_Barra_MDI.Panels(2).Text = "Abrindo Arquivo : " + mynames(i)
           frase = ""
           frase = frase & "SELECT COUNT(*) FROM " + gsPath_DBCAD + ".dbo.TB_ComboCtl "
           frase = frase & "WHERE szArquivo = '" + mynames(i) + "' "
           Set rs = dbApp.Execute(frase)
           If rs(0) = 0 Then
              frase = ""
              frase = frase & "INSERT INTO " + gsPath_DBCAD + ".dbo.TB_ComboCtl (szarquivo,ctipo,tsatualizacao) values "
              frase = frase & "('" & mynames(i) & "','TG','19800101')"
              Set rs = dbApp.Execute(frase)
              GSSequencia = True
              TrataLC (mynames(i))
              TrataOK = 1
              If Not GSSequencia Then Exit For
              Exit For
           Else
              MDI.sta_Barra_MDI.Panels(2).Text = "Transferindo Arquivo : " + mynames(i)
              gsSourceFile = gsPath_CGMPRecebe & mynames(i)
              gsDestFile = gsPath_OUTROS & mynames(i)
              Call CopyFile(gsSourceFile, gsDestFile)
              Kill gsSourceFile
           End If
        End If
    Next
End If
DoEvents

Exit Sub

trataerr:
Call TrataErro("Cadastro de Tag ", Error, " TrataFileLCT")
gbTransferStatus = False

End Sub


Sub TrataFileEST()
Dim i As Integer
Dim ext_1 As String
Dim ext_2 As String
Dim myname As String
Dim TrataOK As Integer

On Error GoTo trataerr

'CAPTURA QUAL A LISTA QUE ESTA CONFIGURADA PARA CONSUMO
    If UCase(gsListas) = "EST" Then
        ext_1 = "*.ETT"
        ext_2 = "*.EST"
    End If
    
    'VERIFICA SE EXISTE TOTAL DE OUTRO TIPO NA PASTA
    Dim extVerifica As String
    If ext_1 = "*.ETT" Then extVerifica = "*.TTV"
    
    ReDim mynames(0)
    myname = UCase(Dir(gsPath_CGMPRecebe & extVerifica))
    
    'SE EXISTIR TOTAL DE OUTRO TIPO, ALTERAR CONFIGURACAO DO SCE
    If Right(myname, 4) = Right(extVerifica, 4) Then
    
        Dim DelArray(2) As String
        Dim ret As String
        Dim d As Integer
        DelArray(1) = "*.ETT"
        DelArray(2) = "*.EST"
        
        For d = 1 To 2
        ret = MoveRecursivo(gsPath_CGMPRecebe, gsPath_OUTROS, DelArray(d))
        DoEvents
        Next d
        
        'RECONFIGURA O SCE PARA TRATAR AS NOVAS LISTAS
        ext_1 = FsetParam("SCE", "tipoLista", DateTime.Now, "TIV")
        DoEvents
    End If
    
If TrataOK = 0 Then

    'PROCURA ARQUIVO DE LISTA TOTAL PARA CONSUMO
    MDI.sta_Barra_MDI.Panels(2).Text = "Procurando Arquivo " & ext_1
    i = 0
    ReDim mynames(i)
    myname = UCase(Dir(gsPath_CGMPRecebe & ext_1))
    
    'LOOP PARA CAPTURA DOS NOMES DOS ARQUIVOS
    Do While myname <> ""
        If Right(myname, 4) = Right(ext_1, 4) Then
            mynames(i) = UCase(myname)
            i = i + 1
            ReDim Preserve mynames(i)
            
        End If
        myname = UCase(Dir)
        DoEvents
    Loop
    
    'LOOP PARA TRATAMENTO INDIVIDUAL DOS ARQUIVOS
    For i = UBound(mynames) - 1 To 0 Step -1
    
    'VERIFICA SE A EXTENSAO LIDA EH DO TIPO CORRETO
        If Right(mynames(i), 4) = Right(ext_1, 4) Then
            DoEvents
            MDI.sta_Barra_MDI.Panels(2).Text = "Abrindo Arquivo : " + mynames(i)
            
            'VERIFICA SE O ARQUIVO JA FOI PROCESSADO
            frase = ""
            frase = frase & "SELECT COUNT(*) FROM " + gsPath_DBCAD + ".dbo.TB_CADESTCTL "
            frase = frase & "WHERE szArquivo = '" + mynames(i) + "' "
            Set rs = dbApp.Execute(frase)
            
            'CASO AINDA NAO TENHA SIDO PROCESSADO, CRIA REGISTRO NA TB_CADEST_CTL
            If rs(0) = 0 Then
                frase = ""
                frase = frase & "INSERT INTO " + gsPath_DBCAD + ".dbo.TB_CADESTCTL (szarquivo,ctipo,tsatualizacao) values "
                frase = frase & "('" & mynames(i) & "','TT','19800101')"
                Set rs = dbApp.Execute(frase)
                Call TrataETT
                TrataOK = 1
                
            'CASO JA TENHA SIDO PROCESSADO, MOVE PARA A PASTA OUTROS
             Else
                MDI.sta_Barra_MDI.Panels(2).Text = "Transferindo Arquivo : " + mynames(i)
                gsSourceFile = gsPath_CGMPRecebe & mynames(i)
                gsDestFile = gsPath_OUTROS & mynames(i)
                Call CopyFile(gsSourceFile, gsDestFile)
                Kill gsSourceFile
            End If

        End If
        DoEvents
    Next
End If

DoEvents
If TrataOK = 0 Then

    'LIMPA ARQUIVOS INCREMENTAIS DE OUTROS FORMATOS
    i = 0
    ReDim mynames(i)
    myname = UCase(Dir(gsPath_CGMPRecebe & "*.TIV"))

    'LOOP PARA CAPTURA DOS NOMES DOS ARQUIVOS
    Do While myname <> ""
        If UCase(Right(myname, 4)) = Right("*.TIV", 4) Then
            mynames(i) = UCase(myname)
            i = i + 1
            ReDim Preserve mynames(i)
        End If
        myname = UCase(Dir)
        DoEvents
    Loop
    
    'LOOP PARA TRATAMENTO INDIVIDUAL DOS ARQUIVOS
    For i = 0 To UBound(mynames) - 1
        MDI.sta_Barra_MDI.Panels(2).Text = "Transferindo Arquivo : " + mynames(i)
              gsSourceFile = gsPath_CGMPRecebe & mynames(i)
              gsDestFile = gsPath_OUTROS & mynames(i)
              Call CopyFile(gsSourceFile, gsDestFile)
              Kill gsSourceFile
              DoEvents
    Next
    
    

    'PROCURA ARQUIVO DE LISTA INCREMENTAL PARA CONSUMO
    i = 0
    ReDim mynames(i)
    myname = UCase(Dir(gsPath_CGMPRecebe & ext_2))
    
    'LOOP PARA CAPTURA DOS NOMES DOS ARQUIVOS
    Do While myname <> ""
        If UCase(Right(myname, 4)) = Right(ext_2, 4) Then
            mynames(i) = UCase(myname)
            i = i + 1
            ReDim Preserve mynames(i)
        End If
        myname = UCase(Dir)
        DoEvents
    Loop

    'LOOP PARA TRATAMENTO INDIVIDUAL DOS ARQUIVOS
    For i = 0 To UBound(mynames) - 1
    
    'VERIFICA SE A EXTENSAO LIDA EH DO TIPO CORRETO
        If UCase(Right(mynames(i), 4)) = Right(ext_2, 4) And TrataOK = 0 Then
           DoEvents
           MDI.sta_Barra_MDI.Panels(2).Text = "Abrindo Arquivo : " + mynames(i)
           
           'VERIFICA SE O ARQUIVO JA FOI PROCESSADO
           frase = ""
           frase = frase & "SELECT COUNT(*) FROM " + gsPath_DBCAD + ".dbo.TB_CADESTCTL "
           frase = frase & "WHERE szArquivo = '" + mynames(i) + "' "
           Set rs = dbApp.Execute(frase)
           
           'CASO AINDA NAO TENHA SIDO PROCESSADO, CRIA REGISTRO NA TB_CADEST_CTL
           If rs(0) = 0 Then
              frase = ""
              frase = frase & "INSERT INTO " + gsPath_DBCAD + ".dbo.TB_CADESTCTL (szarquivo,ctipo,tsatualizacao) values "
              frase = frase & "('" & mynames(i) & "','TG','19800101')"
              Set rs = dbApp.Execute(frase)
              GSSequencia = True
              TrataEST (mynames(i))
              TrataOK = 1
              If Not GSSequencia Then Exit For
              Exit For
              
           'CASO JA TENHA SIDO PROCESSADO, MOVE PARA A PASTA OUTROS
           Else
              MDI.sta_Barra_MDI.Panels(2).Text = "Transferindo Arquivo : " + mynames(i)
              gsSourceFile = gsPath_CGMPRecebe & mynames(i)
              gsDestFile = gsPath_OUTROS & mynames(i)
              Call CopyFile(gsSourceFile, gsDestFile)
              Kill gsSourceFile
           End If
        End If
        DoEvents
    Next
End If
DoEvents

Exit Sub

trataerr:
Call TrataErro("Cadastro de Tag ", Error, " TrataFileTAG")
gbTransferStatus = False

End Sub

Function MoveRecursivo(ByVal SourcePath As String, ByVal DestinyPath As String, ByVal Filter As String)
Dim i As Integer
Dim myname As String
i = 0
ReDim mynames(i)
    myname = UCase(Dir(SourcePath & Filter))
    Do While myname <> ""
        If Right(myname, 4) = Right(Filter, 4) Then
            mynames(i) = UCase(myname)
            i = i + 1
            ReDim Preserve mynames(i)
        End If
        myname = UCase(Dir)
    Loop
    For i = UBound(mynames) - 1 To 0 Step -1
        gsSourceFile = SourcePath & mynames(i)
        gsDestFile = DestinyPath & mynames(i)
        Call CopyFile(gsSourceFile, gsDestFile)
        Kill gsSourceFile
    Next
End Function

Sub TrataFileTG()
Dim i As Integer
Dim ext_1 As String
Dim ext_2 As String
Dim myname As String
Dim TrataOK As Integer

On Error GoTo trataerr

    'CAPTURA QUAL A LISTA QUE ESTA CONFIGURADA PARA CONSUMO
    If UCase(gsListas) = "TAG" Then
        ext_1 = "*.TGT"
        ext_2 = "*.TAG"
    Else
        ext_1 = "*.TTV"
        ext_2 = "*.TIV"
    End If
    
    'VERIFICA SE EXISTE TOTAL DE OUTRO TIPO NA PASTA
    Dim extVerifica As String
    If ext_1 = "*.TTV" Then extVerifica = "*.ETT"
    If ext_1 = "*.ETT" Then extVerifica = "*.TTV"
    
    ReDim mynames(0)
    myname = UCase(Dir(gsPath_CGMPRecebe & extVerifica))
    
    'SE EXISTIR TOTAL DE OUTRO TIPO, ALTERAR CONFIGURACAO DO SCE
    If Right(myname, 4) = Right(extVerifica, 4) Then
    
        Dim DelArray(4) As String
        Dim ret As String
        Dim d As Integer
        DelArray(1) = "*.TTV"
        DelArray(2) = "*.TIV"
        DelArray(3) = "*.NTT"
        DelArray(4) = "*.NTV"
        
        For d = 1 To 4
        ret = MoveRecursivo(gsPath_CGMPRecebe, gsPath_OUTROS, DelArray(d))
        Next d
        
        'RECONFIGURA O SCE PARA TRATAR AS NOVAS LISTAS
        ext_1 = FsetParam("SCE", "tipoLista", DateTime.Now, "EST")
    End If
   
If TrataOK = 0 Then

    'PROCURA ARQUIVO DE LISTA PARA CONSUMO, CONFORME CONFIGURACAO DO SCE
    MDI.sta_Barra_MDI.Panels(2).Text = "Procurando Arquivo " & ext_1
    i = 0
    ReDim mynames(i)
    myname = UCase(Dir(gsPath_CGMPRecebe & ext_1))
    
    Do While myname <> ""
        If Right(myname, 4) = Right(ext_1, 4) Then
            mynames(i) = UCase(myname)
            i = i + 1
            ReDim Preserve mynames(i)
        End If
        myname = UCase(Dir)
    Loop
    For i = UBound(mynames) - 1 To 0 Step -1
        If Right(mynames(i), 4) = Right(ext_1, 4) Then
            DoEvents
            MDI.sta_Barra_MDI.Panels(2).Text = "Abrindo Arquivo : " + mynames(i)
            frase = ""
            frase = frase & "SELECT COUNT(*) FROM " + gsPath_DBCAD + ".dbo.TB_CADtagCtl "
            frase = frase & "WHERE szArquivo = '" + mynames(i) + "' "
            Set rs = dbApp.Execute(frase)
            If rs(0) = 0 Then
                frase = ""
                frase = frase & "INSERT INTO " + gsPath_DBCAD + ".dbo.TB_CADtagCtl (szarquivo,ctipo,tsatualizacao) values "
                frase = frase & "('" & mynames(i) & "','TT','19800101')"
                Set rs = dbApp.Execute(frase)
                Call TrataTGT
                TrataOK = 1
             Else
                MDI.sta_Barra_MDI.Panels(2).Text = "Transferindo Arquivo : " + mynames(i)
                gsSourceFile = gsPath_CGMPRecebe & mynames(i)
                gsDestFile = gsPath_OUTROS & mynames(i)
                Call CopyFile(gsSourceFile, gsDestFile)
                Kill gsSourceFile
            End If

        End If
    Next
End If
DoEvents
If TrataOK = 0 Then


    'LIMPA ARQUIVOS INCREMENTAIS DE OUTROS FORMATOS
    i = 0
    ReDim mynames(i)
    myname = UCase(Dir(gsPath_CGMPRecebe & "*.EST"))

    'LOOP PARA CAPTURA DOS NOMES DOS ARQUIVOS
    Do While myname <> ""
        If UCase(Right(myname, 4)) = Right("*.EST", 4) Then
            mynames(i) = UCase(myname)
            i = i + 1
            ReDim Preserve mynames(i)
        End If
        myname = UCase(Dir)
    Loop
    
    'LOOP PARA TRATAMENTO INDIVIDUAL DOS ARQUIVOS
    For i = 0 To UBound(mynames) - 1
        MDI.sta_Barra_MDI.Panels(2).Text = "Transferindo Arquivo : " + mynames(i)
              gsSourceFile = gsPath_CGMPRecebe & mynames(i)
              gsDestFile = gsPath_OUTROS & mynames(i)
              Call CopyFile(gsSourceFile, gsDestFile)
              Kill gsSourceFile
    Next
    
    

    'TRATA CONSUMO DE LISTAS INCREMENTAIS TIV
    i = 0
    ReDim mynames(i)
    myname = UCase(Dir(gsPath_CGMPRecebe & ext_2))
    Do While myname <> ""   ' Start the loop.
        If UCase(Right(myname, 4)) = Right(ext_2, 4) Then
            mynames(i) = UCase(myname)
            i = i + 1
            ReDim Preserve mynames(i)
        End If
        myname = UCase(Dir)
    Loop

    For i = 0 To UBound(mynames) - 1
        If UCase(Right(mynames(i), 4)) = Right(ext_2, 4) And TrataOK = 0 Then
           DoEvents
           MDI.sta_Barra_MDI.Panels(2).Text = "Abrindo Arquivo : " + mynames(i)
           frase = ""
           frase = frase & "SELECT COUNT(*) FROM " + gsPath_DBCAD + ".dbo.TB_CADtagCtl "
           frase = frase & "WHERE szArquivo = '" + mynames(i) + "' "
           Set rs = dbApp.Execute(frase)
           If rs(0) = 0 Then
              frase = ""
              frase = frase & "INSERT INTO " + gsPath_DBCAD + ".dbo.TB_CADtagCtl (szarquivo,ctipo,tsatualizacao) values "
              frase = frase & "('" & mynames(i) & "','TG','19800101')"
              Set rs = dbApp.Execute(frase)
              GSSequencia = True
              TrataTG (mynames(i))
              TrataOK = 1
              If Not GSSequencia Then Exit For
              Exit For
           Else
              MDI.sta_Barra_MDI.Panels(2).Text = "Transferindo Arquivo : " + mynames(i)
              gsSourceFile = gsPath_CGMPRecebe & mynames(i)
              gsDestFile = gsPath_OUTROS & mynames(i)
              Call CopyFile(gsSourceFile, gsDestFile)
              Kill gsSourceFile
           End If
        End If
    Next
End If
DoEvents

Exit Sub

trataerr:
Call TrataErro("Cadastro de Tag ", Error, " TrataFileTAG")
gbTransferStatus = False

End Sub

Sub TrataFileGZ()
Dim i As Integer
Dim myname As String
Dim TrataOK As Integer
On Error GoTo trataerr

If TrataOK = 0 Then
    MDI.sta_Barra_MDI.Panels(2).Text = "Procurando Arquivo Compactado"
    i = 0
    ReDim mynames(i)
    myname = UCase(Dir(gsPath_CGMPRecebe & "*.*"))   ' Retrieve the first entry.
    Do While myname <> ""   ' Start the loop.
        If UCase(Right(myname, 3)) = ".GZ" Or UCase(Right(myname, 4)) = ".RAR" Or UCase(Right(myname, 4)) = ".ZIP" Then
            mynames(i) = UCase(myname)
            i = i + 1
            ReDim Preserve mynames(i)
        End If
        myname = UCase(Dir)
    Loop
    For i = UBound(mynames) - 1 To 0 Step -1
        frase = gsPath_SCEExtra & "7z.exe -y -o" & gsPath_CGMPRecebe & " e " & gsPath_CGMPRecebe
        frase = frase + mynames(i)
        
        'Shell ((frase))
        ' C:\sce\extra\7z.exe -y -oC:\sce\CGMPRecebe\ e C:\sce\CGMPRecebe\2010110804463317928.GZ
        MDI.sta_Barra_MDI.Panels(2).Text = "Abrindo Arquivo : " + mynames(i)
        TrataOK = 1
        Call SuperShell(frase, gsPath_CGMPRecebe, 120000, 0)
        MDI.sta_Barra_MDI.Panels(2).Text = "Descompactado Arquivo : " + mynames(i)
        gsSourceFile = gsPath_CGMPRecebe & mynames(i)
        gsDestFile = gsPath_OUTROS & mynames(i)
        Call CopyFile(gsSourceFile, gsDestFile)
        Kill gsSourceFile
        DoEvents
    Next
End If


Exit Sub

trataerr:
Call TrataErro("Cadastro de Tag ", Error, " TrataFileGZ")
gbTransferStatus = False

End Sub


Sub TrataFileLN()
Dim i As Integer
Dim TrataOK As Integer
Dim myname As String
Dim ext_1 As String
Dim ext_2 As String



On Error GoTo trataerr

If UCase(gsListas) = "TAG" Then
    ext_1 = "*.LNT"
    ext_2 = "*.NEL"
Else
    ext_1 = "*.NTT"
    ext_2 = "*.NTV"
End If


If TrataOK = 0 Then
    MDI.sta_Barra_MDI.Panels(2).Text = "Procurando Arquivo " + ext_1
    i = 0
    ReDim mynames(i)
    myname = UCase(Dir(gsPath_CGMPRecebe & ext_1))   ' Retrieve the first entry.
    Do While myname <> ""   ' Start the loop.
        If Right(myname, 4) = Right(ext_1, 4) Then
            mynames(i) = UCase(myname)
            i = i + 1
            ReDim Preserve mynames(i)
        End If
        myname = UCase(Dir)
    Loop
    For i = UBound(mynames) - 1 To 0 Step -1
            If Right(mynames(i), 4) = Right(ext_1, 4) And TrataOK = 0 Then
            DoEvents
            MDI.sta_Barra_MDI.Panels(2).Text = "Abrindo Arquivo : " + mynames(i)
            frase = ""
            frase = frase & "SELECT count(*) FROM " + gsPath_DBCAD + ".dbo.TB_CADnelactl "
            frase = frase & "WHERE szArquivo = '" + mynames(i) + "' "
            Set rs = dbApp.Execute(frase)
            If rs(0) = 0 Then
               frase = ""
               frase = frase & "INSERT INTO " + gsPath_DBCAD + ".dbo.TB_CADnelactl (szarquivo,ctipo,tsAtualizacao) values "
               frase = frase & "('" & mynames(i) & "','LT','19800101')"
               Set rs = dbApp.Execute(frase)
               Call TrataLNT
               TrataOK = 1
            Else
               MDI.sta_Barra_MDI.Panels(2).Text = "Transferindo Arquivo : " + mynames(i)
               gsSourceFile = gsPath_CGMPRecebe & mynames(i)
               gsDestFile = gsPath_OUTROS & mynames(i)
               Call CopyFile(gsSourceFile, gsDestFile)
               Kill gsSourceFile
            End If
        End If
    Next
End If
DoEvents
If TrataOK = 0 Then
    i = 0
    ReDim mynames(i)
    myname = UCase(Dir(gsPath_CGMPRecebe & ext_2))   ' Retrieve the first entry.
    Do While myname <> ""   ' Start the loop.
        If UCase(Right(myname, 4)) = Right(ext_2, 4) Then
            mynames(i) = UCase(myname)
            i = i + 1
            ReDim Preserve mynames(i)
        End If
        myname = UCase(Dir)
    Loop
    For i = 0 To UBound(mynames) - 1
        If Right(mynames(i), 4) = Right(ext_2, 4) Then
            MDI.sta_Barra_MDI.Panels(2).Text = "Abrindo Arquivo : " + mynames(i)
            frase = ""
            frase = frase & "SELECT count(*) FROM " + gsPath_DBCAD + ".dbo.TB_CADnelaCtl "
            frase = frase & "WHERE szArquivo = '" + mynames(i) + "' "
            Set rs = dbApp.Execute(frase)
            If rs(0) = 0 And TrataOK = 0 Then
               frase = ""
               frase = frase & "INSERT INTO " + gsPath_DBCAD + ".dbo.TB_CADnelaCtl (szarquivo,Ctipo,tsatualizacao) values "
               frase = frase & "('" & mynames(i) & "','LN','19800101')"
               Set rs = dbApp.Execute(frase)
               GSSequencia = True
               TrataLN (mynames(i))
'               TrataOK = 1
'               If Not GSSequencia Then Exit For
'               If Not TrataOK = 0 Then Exit For
                Exit For
            Else
               MDI.sta_Barra_MDI.Panels(2).Text = "Transferindo Arquivo : " + mynames(i)
               gsSourceFile = gsPath_CGMPRecebe & mynames(i)
               gsDestFile = gsPath_OUTROS & mynames(i)
               Call CopyFile(gsSourceFile, gsDestFile)
               Kill gsSourceFile
            End If
        End If
    Next
End If
DoEvents

Exit Sub

trataerr:
TrataErro "Cadastro de Nela", Error, "tratafileln"
gbTransferStatus = False


End Sub


Public Sub TrataLNT()

Dim Datafile As String
Dim Datafilenum As Long
Dim rs As New Recordset
Dim frase As String
Dim i As Long
Dim Header As LNHeader
Dim file As String

On Error GoTo trataerr
Set rs = Nothing
rs.CursorType = adOpenStatic

frase = ""
frase = frase & "SELECT top 1 szArquivo, lseqfile , ctipo FROM " + gsPath_DBCAD + ".dbo.TB_CadnelaCtl "
frase = frase & "WHERE tsatualizacao = '19800101' AND ctipo = 'LT' "
frase = frase & "ORDER BY szarquivo DESC, lseqfile DESC, ctipo DESC"
Set rs = dbApp.Execute(frase)

If rs.BOF And rs.EOF Then
    Exit Sub
End If
DoEvents
file = rs(0)
Datafile = gsPath_CGMPRecebe & file
Datafilenum = FreeFile
Open Datafile For Binary As Datafilenum
'LE header
Get #Datafilenum, 1, Header

frase = ""
frase = frase & "SELECT top 1 szArquivo, lseqfile , ctipo FROM " + gsPath_DBCAD + ".dbo.TB_CadnelaCtl "
frase = frase & "WHERE tsatualizacao <> '19800101' AND (ctipo = 'LN' OR ctipo = 'LT') "
frase = frase & "ORDER BY lseqfile DESC, ctipo DESC"
Set rs = dbApp.Execute(frase)
MDI.sta_Barra_MDI.Panels(2).Text = "Tratando Arquivo : " + Datafile
If Not rs.BOF And Not rs.EOF Then
    'perguntar sobre a LNT ????
    
    If Val(Header.seq) < Val(rs(1)) Then
       Close Datafilenum

       gsDestFile = gsPath_OUTROS & file
       Call CopyFile(Datafile, gsDestFile)
       Kill Datafile
       
       Call GravaEventos(402, " ", gsUser, 0, "Trata LNT - " + file + " Abaixo da Sequencia ", Format(Header.seq), Right(Format(Header.tag), 9))
       frase = ""
       frase = frase & "DELETE FROM " + gsPath_DBCAD + ".dbo.TB_CADnelaCtl "
       frase = frase & "WHERE tsatualizacao = '19800101' and szArquivo = '" & file & "'"
       Set rs = dbApp.Execute(frase)
       Exit Sub
    End If
End If

'CRIA A ESTRUTURA DA TABELA TB_CADNELA_TMP SE NAO EXISTIR
frase = ""
frase = " if NOT exists (select * from " + gsPath_DBCAD + "..sysobjects where id = object_id(N'" + gsPath_DBCAD + ".DBO.tb_cadnela_tmp'))"
frase = frase & "SELECT * INTO " + gsPath_DBCAD + ".dbo.TB_Cadnela_tmp from " + gsPath_DBCAD + ".dbo.TB_Cadnela where 0=1"
Set rs = dbApp.Execute(frase)

'TRUNCA A TABELA TB_CADNELA_TMP
frase = ""
frase = frase & " truncate table " + gsPath_DBCAD + ".dbo.TB_Cadnela_tmp "
Set rs = dbApp.Execute(frase)

Close Datafilenum
Datafilenum = FreeFile

'COPIA O ARQUIVO PARA O PATHZ
MDI.sta_Barra_MDI.Panels(2).Text = "Copiando Arquivo : " + file
gsSourceFile = gsPath_CGMPRecebe & file
gsDestFile = gsPathZ & file
Call CopyFile(gsSourceFile, gsDestFile)

'FAZ O BULK INSERT DO ARQUIVO NA TABELA TB_CADNELA_TMP
frase = " BULK INSERT " + gsPath_DBCAD + ".dbo.TB_Cadnela_tmp"
frase = frase & " From '" & gsPathZL & file & " '"
frase = frase & " WITH (FORMATFILE = '" & gsPathZL & "SCEExtra\cadnela_tmp_bcp.fmt',FIRSTROW = 2);"
Set rs = dbApp.Execute(frase)

'DELETA ARQUIVO DO PATHZ
MDI.sta_Barra_MDI.Panels(2).Text = "Apagando Arquivo : " + file
Kill gsDestFile

'VERIFICA A QUANTIDADE DE REGISTROS DA TB_CADNELA_TMP
frase = "SELECT count(*) from " + gsPath_DBCAD + ".dbo.TB_CADnela_tmp "
Set rs = dbApp.Execute(frase)

Dim Ireg
Ireg = Val(rs(0))

'SE A QUANTIDADE DE REGISTROS FOR MAIOR QUE ZERO ATUALIZA TABELAS
If Ireg > 0 And Ireg = Header.reg + 0 Then

    'SE A QUANTIDADE DE REGISTROS FOR IGUAL A DO CABECALHO DO ARQUIVO
    If Val(Header.reg) = Ireg Then
        MDI.sta_Barra_MDI.Panels(2).Text = "Atualizando Tabelas"
        frase = gsPath_DBCAD + ".dbo.pr_atuLNT " + Header.seq
        Set rs = dbApp.Execute(frase)
        
        'ATUALIZA A TB_CADNELACTL
        frase = ""
        frase = frase & "UPDATE " + gsPath_DBCAD + ".dbo.TB_CADnelaCtl SET "
        frase = frase & "lseqfile = '" & Header.seq & "',"
        frase = frase & "lregistros = '" & Ireg & "',"
        frase = frase & "ltotal = '" & Ireg & "',"
        frase = frase & "tsatualizacao = '" & Format(Now, "YYYYMMDD HH:MM:SS") & "',"
        frase = frase & "lincl = 0,"
        frase = frase & "lremo = 0,"
        frase = frase & "lalte = 0 "
        frase = frase & "WHERE szArquivo = '" & file & "' "
        Set rs = dbApp.Execute(frase)
        MDI.sta_Barra_MDI.Panels(2).Text = "Transferindo Arquivo Atualizado: " + Datafile
        Call GravaEventos(402, " ", gsUser, 0, "Trata LNT - " + file, 0, Right(Format(Header.reg), 9))
        
        'LIMPA O LOG DE LISTA FALTANTE
        gsListasFaltantes = "0-0"
        MySetPar gsPathLogListas, "LISTAS", "NTV", gsListasFaltantes
        
        'MOVE ARQUIVO PARA A PASTA ATUALIZADOS
        gsDestFile = CriaDir(gsPath_LN, "Atualizados\", Mid(file, 1, 6) + "\", "") & file
        Call CopyFile(Datafile, gsDestFile)
        Close Datafilenum
        
        'APAGA O ARQUIVO DA PASTA CGMPRECEBE
        MDI.sta_Barra_MDI.Panels(2).Text = "Apagando Arquivo : " + file
        Kill Datafile
    Else
        MsgBoxService "Trata Cadastro NELA TOTAL > Quantidade de Registro Errados"
    End If '
Else
    MsgBoxService "Trata Cadastro NELA TOTAL > Quantidade de Registro Zerada"

End If
Exit Sub

trataerr:
Call TrataErro("Cadastro de NELA ", Error, " trataLNT")
gbTransferStatus = False

End Sub

Public Sub TrataEST(file As String)

Dim Datafile As String
Dim Datafilenum As Long
Dim i As Long
Dim rs As New Recordset
Dim frase As String
Dim oldlista As Long
Dim Header As LNHeader

On Error GoTo trataerr
rs.CursorType = adOpenStatic
i = 0

Datafile = gsPath_CGMPRecebe & file
Datafilenum = FreeFile
Open Datafile For Binary As Datafilenum

'LE header
Get #Datafilenum, 1, Header



'Verifica se o arquivo ja consta nos registros da cadnelactl
frase = ""
frase = frase & "SELECT isnull(max(lseqfile),0) FROM " + gsPath_DBCAD + ".dbo.TB_CADESTCTL "
frase = frase & "WHERE tsatualizacao <> '19800101' AND (ctipo = 'TG' OR ctipo = 'TT') "
Set rs = dbApp.Execute(frase)

If Not rs.BOF And Not rs.EOF Then
    oldlista = Val(rs(0))
    
    'ARQUIVO JA FOI PROCESSADO, MOVER PARA A PASTA OUTROS
    If Val(Header.seq) <= oldlista Then
    
       Close Datafilenum
       
       'MOVE O ARQUIVO PARA A PASTA OUTROS
       gsDestFile = gsPath_OUTROS & file
       Call CopyFile(Datafile, gsDestFile)
       Kill Datafile
       
       'DELETA A LINHA DA TB_CADESTCTL
       frase = ""
       frase = frase & "DELETE FROM " + gsPath_DBCAD + ".dbo.TB_CADESTCTL "
       frase = frase & "WHERE tsatualizacao = '19800101' and szarquivo = '" & file & "'"
       Set rs = dbApp.Execute(frase)

       Exit Sub
    Else
    
    'VERIFICA SE EXISTE PULO DE SEQUENCIAL
       If Val(Header.seq) > oldlista + 1 Then
       
            'DELETA A LINHA DA TB_TB_CADESTCTL, POIS HA UM PULO DE SEQUENCIAL
            frase = ""
            frase = frase & "DELETE FROM " + gsPath_DBCAD + ".dbo.TB_CADESTCTL "
            frase = frase & "WHERE tsatualizacao = '19800101' and szarquivo = '" & file & "'"
            Set rs = dbApp.Execute(frase)
            
            Close Datafilenum
            
            'GRAVA LOG DE ERRO
            frase = " FALTA SEQUENCIAL DE NELA :: ABRIR FALHA NO HELP DESK "
            frase = frase + "Sequencia NEL : " + Format(oldlista + 1) + "  ATÉ " + Format(Val(Header.seq) - 1)
            Call LogErro("TrataLN - LISTAS FALTANTES", frase)
            
            'ATUALIZA INFORMACAO DE LISTA FALTANTE
            gsListasFaltantes = Format(oldlista + 1) + "-" + Format(Val(Header.seq) - 1)
            MySetPar gsPathLogListas, "LISTAS", "EST", gsListasFaltantes
            
            GSSequencia = False
            Exit Sub
       End If
    End If
End If

i = 0
MDI!sta_Barra_MDI.Panels(2).Text = "Atualizando EST : " + Header.seq

'SE NAO EXISTIR A TABELA TB_CADEST_INC, CRIA
frase = ""
frase = " if not exists (select * from " + gsPath_DBCAD + "..sysobjects where name = 'tb_cadest_inc') "
frase = frase & "begin SELECT * INTO " + gsPath_DBCAD + ".dbo.tb_cadest_inc from " + gsPath_DBCAD + ".dbo.tb_cadest where 0=1 end"
Set rs = dbApp.Execute(frase)

'TRUNCA TABELA TB_CADEST_INC
frase = ""
frase = frase & " truncate table " + gsPath_DBCAD + ".dbo.tb_cadest_inc "
Set rs = dbApp.Execute(frase)
Close Datafilenum
Datafilenum = FreeFile

'COPIA ARQUIVO PARA A MAQUINA DO BANCO (PATHZ)
MDI.sta_Barra_MDI.Panels(2).Text = "Copiando Arquivo : " + file
gsSourceFile = gsPath_CGMPRecebe & file
gsDestFile = gsPathZ & file
Call CopyFile(gsSourceFile, gsDestFile)

'FAZ O BULK INSERT NA TB_CADEST_INC
frase = ""
frase = frase & " BULK INSERT " + gsPath_DBCAD + ".dbo.tb_cadest_inc"
frase = frase & " From '" & gsPathZL & file & " '"
frase = frase & " WITH (FORMATFILE = '" & gsPathZL & "SCEExtra\cadest_tmp_bcp.fmt',FIRSTROW = 2);"
Set rs = dbApp.Execute(frase)

'DELETE ARQUIVO DA MAQUINA DO BANCO APOS BULK INSERT
MDI.sta_Barra_MDI.Panels(2).Text = "Apagando Arquivo : " + file
Kill gsDestFile

'VERIFICA SE FORAM CARREGADOS OS REGISTROS NA TB_CADEST_INC
frase = ""
frase = frase & "SELECT count(*)from " + gsPath_DBCAD + ".dbo.tb_cadest_inc "
Set rs = dbApp.Execute(frase)

'SE EXISTEM REGISTROS NA TABELA
If rs(0) > 0 Then
    MDI.sta_Barra_MDI.Panels(2).Text = "Atualizando Tabelas"
    
    'EXECUTE A PROC ATUEST_INC
    frase = ""
    frase = frase + gsPath_DBCAD + ".dbo.pr_atuEST_INC " + Header.seq
    Set rs = dbApp.Execute(frase)
    
    'ATUALIZA A TABELA TB_CADESTCTL
    frase = ""
    frase = frase & "UPDATE " + gsPath_DBCAD + ".dbo.TB_CADESTCTL SET "
    frase = frase & "lseqfile = '" & Header.seq & "',"
    frase = frase & "lregistros = '" + Format(rs(1)) + "' ,"
    frase = frase & "ltotal = '" & Format(rs(5)) & "',"
    frase = frase & "tsatualizacao = '" & Format(Now, "YYYYMMDD HH:MM:SS") & "',"
    frase = frase & "lincl = '" + Format(rs(2)) + "' ,"
    frase = frase & "lremo = '" + Format(rs(3)) + "' ,"
    frase = frase & "lalte = '" + Format(rs(4)) + "' "
    frase = frase & "WHERE szArquivo = '" & file & "' "
    Set rs = dbApp.Execute(frase)
    
    'MOVE ARQUIVO PARA A PASTA ATUALIZADOS
    MDI.sta_Barra_MDI.Panels(2).Text = "Transferindo Arquivo : " + Datafile
    gsDestFile = CriaDir(gsPath_LN, "Atualizados\", Mid(file, 1, 6) + "\", "") & file
    
    'GRAVA EVENTO
    Call GravaEventos(402, " ", gsUser, 0, "Trata EST - " + file, Format(Header.seq), Right(Format(Header.reg), 9))
    Close Datafilenum
    Call CopyFile(Datafile, gsDestFile)
    Kill Datafile
    
    'ATUALIZA ARQUIVO DE LISTA FALTANTE
    gsListasFaltantes = "0-0"
    MySetPar gsPathLogListas, "LISTAS", "EST", gsListasFaltantes
    
Else
    MsgBoxService "Trata Cadastro EST > Quantidade de Registro Errados"
End If
Exit Sub

trataerr:
TrataErro "Cadastro de EST", Error, "trataln"
gbTransferStatus = False
End Sub


Public Sub TrataLN(file As String)
Dim Datafile As String
Dim Datafilenum As Long
Dim i As Long
Dim rs As New Recordset
Dim frase As String
Dim oldlista As Long
Dim Header As LNHeader

On Error GoTo trataerr
rs.CursorType = adOpenStatic
i = 0

Datafile = gsPath_CGMPRecebe & file
Datafilenum = FreeFile
Open Datafile For Binary As Datafilenum

'LE header
Get #Datafilenum, 1, Header

'Type LNdet
frase = ""
frase = frase & "SELECT isnull(max(lseqfile),0) FROM " + gsPath_DBCAD + ".dbo.tb_cadnelactl "
frase = frase & "WHERE tsatualizacao <> '19800101' AND (ctipo = 'LN' OR ctipo = 'LT') "
Set rs = dbApp.Execute(frase)
If Not rs.BOF And Not rs.EOF Then
    oldlista = Val(rs(0))
    If Val(Header.seq) <= oldlista Then
       Close Datafilenum
       gsDestFile = gsPath_OUTROS & file
       Call CopyFile(Datafile, gsDestFile)
       Kill Datafile
       frase = ""
       frase = frase & "DELETE FROM " + gsPath_DBCAD + ".dbo.tb_cadnelactl "
       frase = frase & "WHERE tsatualizacao = '19800101' and szarquivo = '" & file & "'"
       Set rs = dbApp.Execute(frase)
       Exit Sub
    Else
       If Val(Header.seq) > oldlista + 1 Then
            frase = ""
            frase = frase & "DELETE FROM " + gsPath_DBCAD + ".dbo.tb_cadnelactl "
            frase = frase & "WHERE tsatualizacao = '19800101' and szarquivo = '" & file & "'"
            Set rs = dbApp.Execute(frase)
            Close Datafilenum
            frase = " FALTA SEQUENCIAL DE NELA :: ABRIR FALHA NO HELP DESK "
            frase = frase + "Sequencia NEL : " + Format(oldlista + 1) + "  ATÉ " + Format(Val(Header.seq) - 1)

            Call LogErro("TrataLN - LISTAS FALTANTES", frase)
            
            gsListasFaltantes = Format(oldlista + 1) + "-" + Format(Val(Header.seq) - 1)
            MySetPar gsPathLogListas, "LISTAS", "NTV", gsListasFaltantes
            
            GSSequencia = False
            Exit Sub
       End If
    End If
End If

i = 0
MDI!sta_Barra_MDI.Panels(2).Text = "Atualizando LN : " + Header.seq

frase = ""
frase = " if not exists (select * from " + gsPath_DBCAD + "..sysobjects where name = 'tb_cadnela_inc') "
frase = frase & "begin SELECT * INTO " + gsPath_DBCAD + ".dbo.tb_cadnela_inc from " + gsPath_DBCAD + ".dbo.tb_cadnela where 0=1 end"
Set rs = dbApp.Execute(frase)

frase = ""
frase = frase & " truncate table " + gsPath_DBCAD + ".dbo.tb_cadnela_inc "
Set rs = dbApp.Execute(frase)

Close Datafilenum
Datafilenum = FreeFile

MDI.sta_Barra_MDI.Panels(2).Text = "Copiando Arquivo : " + file
gsSourceFile = gsPath_CGMPRecebe & file
gsDestFile = gsPathZ & file
Call CopyFile(gsSourceFile, gsDestFile)

frase = ""
frase = frase & " BULK INSERT " + gsPath_DBCAD + ".dbo.tb_cadnela_inc"
frase = frase & " From '" & gsPathZL & file & " '"
frase = frase & " WITH (FORMATFILE = '" & gsPathZL & "SCEExtra\cadnela_tmp_bcp.fmt',FIRSTROW = 2);"
Set rs = dbApp.Execute(frase)

MDI.sta_Barra_MDI.Panels(2).Text = "Apagando Arquivo : " + file
Kill gsDestFile

frase = ""
frase = frase & "SELECT count(*)from " + gsPath_DBCAD + ".dbo.tb_cadnela_inc "
Set rs = dbApp.Execute(frase)

If rs(0) > 0 Then
    MDI.sta_Barra_MDI.Panels(2).Text = "Atualizando Tabelas"
    frase = " "
    frase = frase & gsPath_DBCAD + ".dbo.pr_atuNEL_INC " + Header.seq
    Set rs = dbApp.Execute(frase)
    
    'Atualiza cadctl
    frase = ""
    frase = frase & "UPDATE " + gsPath_DBCAD + ".dbo.tb_cadnelactl SET "
    frase = frase & "lseqfile = '" & Header.seq & "',"
    frase = frase & "lregistros = '" + Format(rs(1)) + "' ,"
    frase = frase & "ltotal = '" & Format(rs(5)) & "',"
    frase = frase & "tsatualizacao = '" & Format(Now, "YYYYMMDD HH:MM:SS") & "',"
    frase = frase & "lincl = '" + Format(rs(2)) + "' ,"
    frase = frase & "lremo = '" + Format(rs(3)) + "' ,"
    frase = frase & "lalte = '" + Format(rs(4)) + "' "
    frase = frase & "WHERE szArquivo = '" & file & "' "
    Set rs = dbApp.Execute(frase)
    MDI.sta_Barra_MDI.Panels(2).Text = "Transferindo Arquivo : " + Datafile
    gsDestFile = CriaDir(gsPath_LN, "Atualizados\", Mid(file, 1, 6) + "\", "") & file
    
    Call GravaEventos(402, " ", gsUser, 0, "Trata LN - " + file, Format(Header.seq), Right(Format(Header.reg), 9))
    Close Datafilenum
    Call CopyFile(Datafile, gsDestFile)
    Kill Datafile
    
    gsListasFaltantes = "0-0"
    MySetPar gsPathLogListas, "LISTAS", "NTV", gsListasFaltantes
    
Else
    MsgBoxService "Trata Cadastro NELA > Quantidade de Registro Errados"
End If
Exit Sub

trataerr:
TrataErro "Cadastro de Nela", Error, "trataln"
gbTransferStatus = False

End Sub

Public Sub TrataTRNAutoExpresso()
On Error GoTo trataerr
Dim rs As New Recordset
Dim frase As String
Dim DataMov As String
Dim seqfile As Integer

rs.CursorType = adOpenStatic

'[x] correção da versão 9.0.0
DataMov = Format(CVDate(gsNextTrn) - 1, "yyyymmdd")

'Elimina Transacao Zeradas
If gbZeradas = 0 Then
    frase = ""
    frase = frase & " UPDATE " + gsPath_DB + ".dbo.tb_transacao "
    frase = frase & " SET "
    frase = frase & " lseqreg = 0,"
    frase = frase & " lseqfile = 0,"
    frase = frase & " tsdatamovimento = '" + DataMov + "', "
    frase = frase & " ccontador = abs(CAST(Ccontador as int))%1000000 "
    frase = frase & " WHERE "
    frase = frase & " tsdataoperacao <= '" & DataMov & "' and "
    frase = frase & " Lseqfile is null and "
    
    'INCLUIDO NA VERSÃO 8.7.6
    frase = frase & " (ivalor = 0) "
    
    'ALTERADO NA VERSAO 9.2.8
    'TRATAMENTO PARA MENSALISTAS
    'NAO CONSIDERAR TRANSACOES COM ISTATCAT<0 NOS TRNS
    'POIS EXISTE UM TRATAMENTO DESSAS TRANSACOES NO RELATORIO DE NAO-CONFORMIDADE
    If gsTrnSomenteStatCatPositivo = 1 Then
        frase = frase & " and istatcat >= 0"
    End If
    
    Set rs = dbApp.Execute(frase)
End If

'CAPTURA AS TRANSACOES COM LSEQFILE NULL, VALOR>0 E IISSUERS AUTOEXPRESSO
frase = ""
frase = frase & "SELECT count(*) from " + gsPath_DB + ".dbo.tb_transacao "
frase = frase & "WHERE "
frase = frase & " tsdataoperacao <= '" & DataMov & "' and "
frase = frase & " Lseqfile is null and ivalor is not null and"
frase = frase & " (iissuer = 100 or iissuer=101)"
    
Set rs = dbApp.Execute(frase)

If rs(0) = 0 Then
'(tsdatahora,icodigo,cplaca,cusuario,lturno,sztexto,iparam,lparam)
    Call GravaEventos(401, " ", gsUser, 0, "Grava TRN AE - Nada a GRAVAR", 0, 0)
Else
    ' criar TRN
    frase = ""
    frase = frase & "SELECT MAX(Lseqfile) from " + gsPath_DB + ".dbo.tb_transacao "
    frase = frase & "WHERE Lseqfile is not null"
    
    Set rs = dbApp.Execute(frase)
    
    If IsNull(rs(0)) Then
        seqfile = IIf(Val(gsNextTrnNr) > 0, Val(gsNextTrnNr), 1)
    Else
        seqfile = Val(rs(0)) + 1
    End If
    MDI!sta_Barra_MDI.Panels(2).Text = "Criando TRN : " + Format(seqfile)
    
    If gsGerarTrnAE = 1 Then
        If gsSepararManuaisAE = 1 Then
            'Gera TRN de transacoes com saidas automaticas
            AtualizaTRN DataMov, DataMov, seqfile, 1, "AE"
            
            'Gera TRN de transacoes com saidas manuais, incrementando o seqfile
            AtualizaTRN DataMov, DataMov, seqfile + 1, 0, "AE"
        Else
            'gera TRN do dia com transacoes automaticas e manuais
             AtualizaTRN DataMov, DataMov, seqfile, 2, "AE"
        End If
    End If
    
End If

'[+] Corrigindo versão 9.0.0
If CVDate(gsNextTrn) < Now Then
    gsNextTrn = CVDate(gsNextTrn) + 1
End If

'[-] Removido da versão 9.0.0
'gsNextTrn = DateAdd("n", gsNextTrnDelta, gsNextTrn)




'MySetPar gsPathIniFile, "SISTEMA", "NextTRN", Format(gsNextTrn, "DD/MM/YYYY hh:mm:ss")
frase = ""
frase = frase & " UPDATE " + gsPath_DB + ".dbo.tb_appparam SET szvalor = " + Format(gsNextTrn, "DD/MM/YYYY hh:mm:ss") + " WHERE szparam = 'NextTRN'"
Set rs = dbApp.Execute(frase)

gsNextTrnNr = seqfile

'MySetPar gsPathIniFile, "SISTEMA", "NextTRNnr", gsNextTrnNr
frase = ""
frase = frase & " UPDATE " + gsPath_DB + ".dbo.tb_appparam SET szvalor = " + gsNextTrnNr + " WHERE szparam = 'NextTRNnr'"
Set rs = dbApp.Execute(frase)

Exit Sub
'On Error GoTo trataerr
trataerr:
Call TrataErro(app.title, Error, "trataTRN")

End Sub



Public Sub TrataTRN()
On Error GoTo trataerr
Dim rs As New Recordset
Dim frase As String
Dim DataMov As String
Dim seqfile As Integer

rs.CursorType = adOpenStatic

'[x] correção da versão 9.0.0
DataMov = Format(CVDate(gsNextTrn) - 1, "yyyymmdd")

'Elimina Transacao Zeradas
If gbZeradas = 0 Then
    frase = ""
    frase = frase & " UPDATE " + gsPath_DB + ".dbo.tb_transacao "
    frase = frase & " SET "
    frase = frase & " lseqreg = 0,"
    frase = frase & " lseqfile = 0,"
    frase = frase & " tsdatamovimento = '" + DataMov + "', "
    frase = frase & " ccontador = abs(CAST(Ccontador as int))%1000000 "
    frase = frase & " WHERE "
    frase = frase & " tsdataoperacao <= '" & DataMov & "' and "
    frase = frase & " Lseqfile is null and "
    
    
    'INCLUIDO NA VERSÃO 8.7.6
    frase = frase & " (ivalor = 0) "
    
    'ALTERADO NA VERSAO 9.2.8
    'TRATAMENTO PARA MENSALISTAS
    'NAO CONSIDERAR TRANSACOES COM ISTATCAT<0 NOS TRNS
    'POIS EXISTE UM TRATAMENTO DESSAS TRANSACOES NO RELATORIO DE NAO-CONFORMIDADE
    If gsTrnSomenteStatCatPositivo = 1 Then
        frase = frase & " and istatcat >= 0"
    End If
    
    Set rs = dbApp.Execute(frase)
End If

frase = ""
frase = frase & "SELECT count(*) from " + gsPath_DB + ".dbo.tb_transacao "
frase = frase & "WHERE "
frase = frase & " tsdataoperacao <= '" & DataMov & "' and "
frase = frase & " Lseqfile is null and ivalor >0 "
'frase = frase & " and Iissuer=290"

If gsTrnSomenteStatCatPositivo = 1 Then
        frase = frase & " and istatcat >= 0"
    End If
    
Set rs = dbApp.Execute(frase)

If rs(0) = 0 Then
'(tsdatahora,icodigo,cplaca,cusuario,lturno,sztexto,iparam,lparam)
    Call GravaEventos(401, " ", gsUser, 0, "Grava TRN - Nada a GRAVAR", 0, 0)
Else
    ' criar TRN
    frase = ""
    frase = frase & "SELECT MAX(Lseqfile) from " + gsPath_DB + ".dbo.tb_transacao "
    frase = frase & "WHERE Lseqfile is not null"
    
    Set rs = dbApp.Execute(frase)
    
    If IsNull(rs(0)) Or rs(0) = 0 Then
        seqfile = IIf(Val(gsNextTrnNr) > 0, Val(gsNextTrnNr), 1)
    Else
        seqfile = Val(rs(0)) + 1
    End If
    MDI!sta_Barra_MDI.Panels(2).Text = "Criando TRN : " + Format(seqfile)
    
    'Se o parametro GerarTRNAE estiver desabilitado, gerar o TRN contendo todas as transacoes
    'Independente se eh Sem Parar ou Autoexpresso
    If gsGerarTrnAE = 0 Then
        If gsSepararManuaisSemParar = 1 Then
            'Gera TRN de transacoes com saidas automaticas
            AtualizaTRN DataMov, DataMov, seqfile, 1, "TODAS"
            
            'Gera TRN de transacoes com saidas manuais, incrementando o seqfile
            AtualizaTRN DataMov, DataMov, seqfile + 1, 0, "TODAS"
        
        Else
            'gera TRN do dia com transacoes automaticas e manuais
             AtualizaTRN DataMov, DataMov, seqfile, 2, "TODAS"
        End If
    'Se o parametro GerarTRNAE estiver HABILITADO, gerar o TRN somente de TransacoesSemParar nesta rotina
    Else
        If gsSepararManuaisSemParar = 1 Then
            'Gera TRN de transacoes com saidas automaticas
            AtualizaTRN DataMov, DataMov, seqfile, 1, "SP"
            
            'Gera TRN de transacoes com saidas manuais, incrementando o seqfile
            AtualizaTRN DataMov, DataMov, seqfile + 1, 0, "SP"
        
        Else
            'gera TRN do dia com transacoes automaticas e manuais
             AtualizaTRN DataMov, DataMov, seqfile, 2, "SP"
        End If
    
    End If
End If

'[+] Corrigindo versão 9.0.0
If CVDate(gsNextTrn) < Now Then
    gsNextTrn = CVDate(gsNextTrn) + 1
End If

'[-] Removido da versão 9.0.0
'gsNextTrn = DateAdd("n", gsNextTrnDelta, gsNextTrn)




'MySetPar gsPathIniFile, "SISTEMA", "NextTRN", Format(gsNextTrn, "DD/MM/YYYY hh:mm:ss")
frase = ""
frase = frase & " UPDATE " + gsPath_DB + ".dbo.tb_appparam SET szvalor = '" + Format(gsNextTrn, "DD/MM/YYYY hh:mm:ss") + "' WHERE szparam like '%NextTRN'"
Set rs = dbApp.Execute(frase)

gsNextTrnNr = seqfile

'MySetPar gsPathIniFile, "SISTEMA", "NextTRNnr", gsNextTrnNr
frase = ""
frase = frase & " UPDATE " + gsPath_DB + ".dbo.tb_appparam SET szvalor = " + gsNextTrnNr + " WHERE szparam like 'NextTRNnr'"
Set rs = dbApp.Execute(frase)

Exit Sub
'On Error GoTo trataerr
trataerr:
Call TrataErro(app.title, Error, "trataTRN")

End Sub

Public Sub AtualizaTRNAutoExpresso(Pdatamovimento As String, Pdataoperacao As String, PSeqfile As Integer, PSoAutomaticas As Integer)
On Error GoTo trataerr
Dim rs As New Recordset
Dim rsupdate As New Recordset
Dim frase As String
Dim i As Long

'Rotina sem begin tran rodado pelo VB
rs.CursorType = adOpenStatic
rsupdate.CursorType = adOpenStatic

'TRN DE SAIDAS AUTOMATICAS
frase = ""
frase = frase & " Declare @@contador as int"
frase = frase & " Set @@contador = 0"
frase = frase & " UPDATE " + gsPath_DB + ".dbo.tb_transacao "
frase = frase & " SET "
frase = frase & " @@contador = @@contador +1,"
frase = frase & " Lseqreg = @@contador,"
frase = frase & " Lseqfile = '" + Format(PSeqfile) + "',"
frase = frase & " tsdatamovimento = '" & Pdatamovimento & "', "
frase = frase & " ccontador = abs(CAST(Ccontador as int))%1000000 "
frase = frase & " WHERE "
frase = frase & " tsdataoperacao <= '" & Pdatamovimento & "' and "
frase = frase & " lseqfile is null and "
frase = frase & " cplaca is not null and "
frase = frase & " ltag is not null and "
frase = frase & " ivalor is not null "
frase = frase & " iissuer = 100 or iissuer=101"


'SEPARAR TRN DE SAIDAS AUTOMATICAS / MANUAIS
'SE O PARAMETRO ESTIVER SETADO PARA 1, GERAR TRN DE SAIDAS AUTOMATICAS
If PSoAutomaticas = 1 Then
    frase = frase & "and istsaida=1 "
'SE O PARAMETRO ESTIVER SETADO PARA 0, GERAR TRN DE SAIDAS MANUAIS
ElseIf PSoAutomaticas = 0 Then
    frase = frase & "and istsaida=0 "
'DIFERENTE DE 1 E 0, GERAR TRN DE TODAS AS SAIDAS MANUAIS E AUTOMATICAS
Else
    frase = frase
End If


Set rsupdate = dbApp.Execute(frase)

DoEvents
'Call CriaFileTRNAutoExpresso(Pdatamovimento, PSeqfile)
Call CriaFileTRN(Pdatamovimento, PSeqfile)

' Atualiza tabela de reltec
Set rs = Nothing
Set rsupdate = Nothing

Exit Sub
'On Error GoTo trataerr
trataerr:
Call TrataErro(app.title, Error, "AtualizaTrn")
If gbTransferStatus Then
    MsgBoxService "Erro na geração do Arquivo" + Error
'    Set rsupdate = dbApp.Execute("rollback tran")
End If


End Sub


Public Sub AtualizaTRN(Pdatamovimento As String, Pdataoperacao As String, PSeqfile As Integer, PSoAutomaticas As Integer, PMeioPagamento As String)
On Error GoTo trataerr
Dim rs As New Recordset
Dim rsupdate As New Recordset
Dim frase As String
Dim i As Long

'Rotina sem begin tran rodado pelo VB
rs.CursorType = adOpenStatic
rsupdate.CursorType = adOpenStatic

frase = ""
frase = frase & " Declare @@contador as int"
frase = frase & " Set @@contador = 0"
frase = frase & " UPDATE " + gsPath_DB + ".dbo.tb_transacao "
frase = frase & " SET "
frase = frase & " @@contador = @@contador +1,"
frase = frase & " Lseqreg = @@contador,"
frase = frase & " Lseqfile = '" + Format(PSeqfile) + "',"
frase = frase & " tsdatamovimento = '" & Pdatamovimento & "', "
frase = frase & " ccontador = abs(CAST(Ccontador as int))%1000000 "
frase = frase & " WHERE "
frase = frase & " tsdataoperacao <= '" & Pdatamovimento & "' and "
frase = frase & " lseqfile is null and "
frase = frase & " cplaca is not null and "
frase = frase & " ltag is not null and "
frase = frase & " ivalor > 0 "

'ALTERADO NA VERSAO 9.2.8
'TRATAMENTO PARA MENSALISTAS
'NAO CONSIDERAR TRANSACOES COM ISTATCAT<0 NOS TRNS
'POIS EXISTE UM TRATAMENTO DESSAS TRANSACOES NO RELATORIO DE NAO-CONFORMIDADE
If gsTrnSomenteStatCatPositivo = 1 Then
    frase = frase & " and istatcat >= 0"
End If
    

'VERSAO 9.2.8
'SEPARAR TRN DE SAIDAS AUTOMATICAS / MANUAIS
'SE O PARAMETRO ESTIVER SETADO PARA 1, GERAR TRN DE SAIDAS AUTOMATICAS
If PSoAutomaticas = 1 Then
    frase = frase & "and istsaida=1 "
    
'SE O PARAMETRO ESTIVER SETADO PARA 0, GERAR TRN DE SAIDAS MANUAIS
ElseIf PSoAutomaticas = 0 Then
    frase = frase & "and istsaida=0 "
    
'DIFERENTE DE 1 E 0, GERAR TRN DE TODAS AS SAIDAS MANUAIS E AUTOMATICAS
Else
    frase = frase
End If

'VERIFICA TIPO DO MEIO DO PAGAMENTO PARA GERAR O TRN
'SOMENTE TRANSACOES SEM PARAR
If PMeioPagamento = "SP" Then
    frase = frase & " and iissuer=290 "
    
'SOMENTE TRANSACOES AUTOEXPRESSO
ElseIf PMeioPagamento = "AE" Then
    frase = frase & " and (iissuer=100 or iissuer=101) "
    
'TODAS AS TRANSACOES EM UM UNICO ARQUIVO
Else
    frase = frase
End If


Set rsupdate = dbApp.Execute(frase)

DoEvents

Call CriaFileTRN(Pdatamovimento, PSeqfile)

' Atualiza tabela de reltec
Set rs = Nothing
Set rsupdate = Nothing

Exit Sub
'On Error GoTo trataerr
trataerr:
Call TrataErro(app.title, Error, "AtualizaTrn")
If gbTransferStatus Then
    MsgBoxService "Erro na geração do Arquivo" + Error
'    Set rsupdate = dbApp.Execute("rollback tran")
End If


End Sub

Sub TrataFileRT()
On Error GoTo trataerr
Dim i As Integer
Dim myname As String
Dim auxb  As Boolean
ReDim mynames(i)

myname = Dir(gsPath_CGMPRecebe & "*.*")   ' Retrieve the first entry.
Do While myname <> ""   ' Start the loop.
    If Right(myname, 4) = ".TRT" Or Right(myname, 4) = ".TRF" Or Right(myname, 4) = ".TRN" Then
        mynames(i) = myname
        i = i + 1
        ReDim Preserve mynames(i)
    End If
    myname = Dir
Loop

auxb = False
DoEvents

For i = 0 To UBound(mynames) - 1
    If Right(mynames(i), 4) = ".TRN" Then
        TrataTR (mynames(i))
    End If
    If Right(mynames(i), 4) = ".TRF" Then
        TrataRF (mynames(i))
    End If
    If Right(mynames(i), 4) = ".TRT" Then
        TrataRT (mynames(i))
    End If
Next

Exit Sub

Exit Sub
'On Error GoTo trataerr
trataerr:
Call TrataErro(app.title, Error, "TrataFileRT")


End Sub
Public Sub TrataTR(file As String)
On Error GoTo trataerr
Dim Datafile As String
Dim Datafilenum As Long
Dim myname As String
Dim rs As New Recordset
Dim frase As String
Dim tipo As String
Dim aux As String

Dim Header As TRHeader
rs.CursorType = adOpenStatic

Datafile = gsPath_CGMPRecebe & file
Datafilenum = FreeFile
Open Datafile For Binary As Datafilenum

'LE header
Get #Datafilenum, 1, Header

frase = ""
frase = frase & "SELECT * FROM " + gsPath_DB + ".dbo.tb_trnctl "
frase = frase & "WHERE tipo = 'TR' and arquivo = '" & file & "' "
Set rs = dbApp.Execute(frase)

If Not rs.BOF And Not rs.EOF Then
    Close Datafilenum
    gsDestFile = gsPath_OUTROS & file
    Call CopyFile(Datafile, gsDestFile)
    Kill Datafile
'    frmRT!msglbl = "Arquivo : " & file & "removido para OUTROS "
Else

'atualiza tb_trnctl
    frase = ""
    frase = frase & " delete " + gsPath_DB + ".dbo.tb_trnctl "
    frase = frase & " where seqfile = " & Header.seqfile
    Set rs = dbApp.Execute(frase)
    
    frase = ""
    frase = frase & "INSERT INTO " + gsPath_DB + ".dbo.tb_trnctl ("
    frase = frase & "tipo,pais,id,seqfile,arquivo,dtger,hrger,"
    frase = frase & "reginformado,totinformado)"
    frase = frase & " values ("
    frase = frase & "'" & Header.tipo & "',"
    frase = frase & "'" & Header.pais & "',"
    frase = frase & "'" & Header.id & "',"
    frase = frase & "'" & Header.seqfile & "',"
    frase = frase & "'" & file & "',"
    frase = frase & "'" & Header.dtger & "',"
    frase = frase & "'" & Header.hrger & "',"
    frase = frase & "'" & Header.reg & "',"
    frase = frase & "'" & Header.valtot & "')"
    Set rs = dbApp.Execute(frase)
    
    'atualiza datamovimento em tb_trnctl
    frase = ""
    frase = frase & " select top 1 tsdatamovimento "
    frase = frase & " from " + gsPath_DB + ".dbo.tb_transacao "
    frase = frase & " where lseqfile = " & Header.seqfile
    Set rs = dbApp.Execute(frase)
    
    If Not rs.BOF And Not rs.EOF Then
        aux = "'" + rs(0) + "'"
    Else
        aux = "'" + Format(DateAdd("d", -1, CVDate(Mid(Header.dtger, 7, 2) + "/" + Mid(Header.dtger, 5, 2) + "/" + Mid(Header.dtger, 1, 4))), "YYYYMMDD") + "'"
    End If
    frase = ""
    frase = frase & " update " + gsPath_DB + ".dbo.tb_trnctl set datamovimento = " & aux
    frase = frase & " where datamovimento is null and seqfile = " & Header.seqfile
    Set rs = dbApp.Execute(frase)
    Call GravaEventos(401, " ", gsUser, 0, "Trata TR - " + file, Format(Header.seqfile), Right(Format(Header.reg), 9))
    gsDestFile = CriaDir(gsPath_TRT, "Atualizados\", Mid(file, 1, 6) + "\", "") & file
   
    Call CopyFile(Datafile, gsDestFile)
    Close Datafilenum
    Kill Datafile
End If

Exit Sub
'On Error GoTo trataerr
trataerr:
Call TrataErro(app.title, Error, "TrataTR")


End Sub
Public Sub TrataRF(file As String)
On Error GoTo trataerr
Dim Datafile As String
Dim Datafilenum As Long
Dim myname As String
Dim rs As New Recordset
Dim frase As String
Dim aux As String
Dim tipo As RFTipo
Dim Header As RFHeader
Dim Detalhe1  As RFDetalhe1
Dim Detalhe2  As RFDetalhe2

rs.CursorType = adOpenStatic

Datafile = gsPath_CGMPRecebe & file
Datafilenum = FreeFile
Open Datafile For Binary As Datafilenum

frase = ""
frase = frase & "SELECT * FROM " + gsPath_DB + ".dbo.tb_trfctl "
frase = frase & "WHERE tipo = 'RF' and arquivo = '" & file & "' "
Set rs = dbApp.Execute(frase)

If Not rs.BOF And Not rs.EOF Then
    Close Datafilenum
    gsDestFile = gsPath_OUTROS & file
    Call CopyFile(Datafile, gsDestFile)
    Kill Datafile
'    frmRT!msglbl = "Arquivo : " & file & "removido para OUTROS "
Else
    'LE header
    Get #Datafilenum, 1, tipo
    Do While Not EOF(Datafilenum)
    Select Case tipo.tipo
    Case "RF"
        Get #Datafilenum, , Header
        
        ' limpa tb_trfctl
        frase = ""
        frase = frase & " delete " + gsPath_DB + ".dbo.tb_trfctl "
        frase = frase & " where seqfile = " & Header.seqfile
        Set rs = dbApp.Execute(frase)
        ' limpa tb_relfin1
        frase = ""
        frase = frase & " delete " + gsPath_DB + ".dbo.tb_relfin1 "
        frase = frase & " where seqfile = " & Header.seqfile
        Set rs = dbApp.Execute(frase)
        ' limpa tb_relfin2
        frase = ""
        frase = frase & " delete " + gsPath_DB + ".dbo.tb_relfin2 "
        frase = frase & " where seqfile = " & Header.seqfile
        Set rs = dbApp.Execute(frase)
        
        
        frase = ""
        frase = frase & "INSERT INTO " + gsPath_DB + ".dbo.tb_trfctl ("
        frase = frase & "tipo,arquivo,seqfile,dtger,hrger,"
        frase = frase & "regencontrado,totinformado,totencontrado)"
        frase = frase & " values ("
        frase = frase & "'" & tipo.tipo & "',"
        frase = frase & "'" & file & "',"
        frase = frase & "'" & Header.seqfile & "',"
        frase = frase & "'" & Header.dtger & "',"
        frase = frase & "'" & Header.hrger & "',"
        frase = frase & "'" & Header.regrejeitado & "',"
        frase = frase & "'" & Header.totAceito & "',"
        frase = frase & "'" & Header.totNaoAceito & "')"
        Set rs = dbApp.Execute(frase)
    Case "DR"
        Get #Datafilenum, , Detalhe1
        frase = ""
        frase = frase & "INSERT INTO " + gsPath_DB + ".dbo.tb_relFin1 ("
        frase = frase & "pais,tag,acesso,entradadia,entradahora,"
        frase = frase & "saidadia,saidahora,valor,codigo,"
        frase = frase & "seqfile,seqreg,arquivo)"
        frase = frase & " values ("
        frase = frase & "'" & Detalhe1.pais & "',"
        frase = frase & "'" & Detalhe1.tag & "',"
        frase = frase & "'" & Detalhe1.acesso & "',"
        frase = frase & "'" & Detalhe1.entradadia & "',"
        frase = frase & "'" & Detalhe1.entradahora & "',"
        frase = frase & "'" & Detalhe1.saidadia & "',"
        frase = frase & "'" & Detalhe1.saidahora & "',"
        frase = frase & "'" & Detalhe1.valor & "',"
        frase = frase & "'" & Detalhe1.codigo & "',"
        frase = frase & "'" & Header.seqfile & "',"
        frase = frase & "'" & Detalhe1.seqreg & "',"
        frase = frase & "'" & file & "')"
        Set rs = dbApp.Execute(frase)
    Case "DF"
        Get #Datafilenum, , Detalhe2
        frase = ""
        frase = frase & "INSERT INTO " + gsPath_DB + ".dbo.tb_relfin2 ("
        frase = frase & "Datapagamento,ValorCGMP,ValorSGMP,"
        frase = frase & "Seqfile,Arquivo)"
        frase = frase & " values ("
        frase = frase & "'" & Detalhe2.datapagamento & "',"
        frase = frase & "'" & Detalhe2.valorCGMP & "',"
        frase = frase & "'" & Detalhe2.valorSGMP & "',"
        frase = frase & "'" & Header.seqfile & "',"
        frase = frase & "'" & file & "')"
        Set rs = dbApp.Execute(frase)
    End Select
    Get #Datafilenum, , tipo
Loop
    
    'atualiza datamovimento em tb_trfctl
    frase = ""
    frase = frase & " select top 1 datamovimento "
    frase = frase & " from " + gsPath_DB + ".dbo.tb_trnctl "
    frase = frase & " where seqfile = " & Header.seqfile
    Set rs = dbApp.Execute(frase)
    
    If Not rs.BOF And Not rs.EOF Then
        aux = rs(0)
        frase = ""
        frase = frase & " update " + gsPath_DB + ".dbo.tb_trfctl set datamovimento = " & aux
        frase = frase & " where datamovimento is null and seqfile = " & Header.seqfile
        Set rs = dbApp.Execute(frase)
    End If
    
    Call GravaEventos(401, " ", gsUser, 0, "Trata RF - " + file, Format(Header.seqfile), Right(Format(Header.regrejeitado), 9))
    gsDestFile = CriaDir(gsPath_TRT, "Atualizados\", Mid(file, 1, 6) + "\", "") & file
    
    Call CopyFile(Datafile, gsDestFile)
    Close Datafilenum
    Kill Datafile
End If

Exit Sub
'On Error GoTo trataerr
trataerr:
Call TrataErro(app.title, Error, "TrataRF")


End Sub

Public Sub CriaFileTRN(Pdatamovimento As String, PSeqfile As Integer)
On Error GoTo trataerr
Dim Datafile As String
Dim gsDestFile As String
Dim Datafilenum As Long
Dim rs As New Recordset
Dim rsupdate As New Recordset
Dim frase As String
Dim tipo As String
Dim i As Long
Dim incl As Long
Dim valor As Long
Dim Header As TRHeader
Dim Detalhe As TRdet

Dim file As String
Dim DataMov As Variant
Dim seqfile As String
Dim Qtde    As Long
Dim DataOper As String

rs.CursorType = adOpenStatic

seqfile = Format(PSeqfile, "00000")
file = Format(Now, "YYYYMMDDHHMMSS") + seqfile + ".TRN"

frase = ""
frase = frase & " SELECT sum(cast(ivalor as int)) as Valor, count(*) as qtde "
frase = frase & " from " + gsPath_DB + ".dbo.tb_transacao "
frase = frase & " WHERE "
frase = frase & " Lseqfile = '" & PSeqfile & "' "
frase = frase & " and Lseqreg > 0 "
Set rs = dbApp.Execute(frase)

If (rs.BOF And rs.EOF) Or IsNull(rs(0)) Then
    Exit Sub
Else
    Qtde = rs("qtde")
    valor = rs("valor")
End If


rs.MoveFirst
Header.tipo = "TR"
Header.pais = "0618"
Header.id = Format(Val(gsEst_Codigo), "00000")
Header.seqfile = Mid(file, 15, 5)
Header.dtger = Mid(file, 1, 8)
Header.hrger = Mid(file, 9, 6)
Header.reg = Format(Qtde, "000000")
Header.valtot = Format(valor, "000000000000")
Header.crlf = Chr(13) & Chr(10)

frase = ""
frase = frase & " SELECT * from " + gsPath_DB + ".dbo.tb_transacao "
frase = frase & " WHERE "
frase = frase & " lseqfile = '" & PSeqfile & "' "
frase = frase & " and lseqreg > 0 "
frase = frase & " order by lseqfile,lseqreg "
Set rs = dbApp.Execute(frase)

' chico frmRT!msglbl = "Criando Arquivo " & file
Datafile = gsPath_TRT & file
Datafilenum = FreeFile
Open Datafile For Binary As Datafilenum

'grava header
Put #Datafilenum, 1, Header

'cria detalhe
rs.MoveFirst
i = 0
Do While Not rs.EOF
    i = i + 1
    Detalhe.regtipo = "D"
    Detalhe.regseq = Poe_Zero_Esquerda(rs("Lseqreg"), 6)
    Detalhe.pais = "0618"
    Detalhe.tag = Poe_Zero_Esquerda(rs("Iissuer"), 5) + Poe_Zero_Esquerda(rs("Ltag"), 10)
    Detalhe.acesso = Poe_Zero_Esquerda(rs("Iacesso"), 4)
    Detalhe.dtent = Format(rs("tsentrada"), "yyyymmdd")
    Detalhe.hrent = Format(rs("tsentrada"), "hhmmss")
    Detalhe.stent = Poe_Zero_Esquerda(1 - rs("Istentrada"), 1)
    Detalhe.dtsai = Format(rs("tssaida"), "yyyymmdd")
    Detalhe.hrsai = Format(rs("tssaida"), "hhmmss")
    Detalhe.valor = Poe_Zero_Esquerda(rs("ivalor"), 8)
    Detalhe.stcobranca = Poe_Zero_Esquerda(rs("istcobranca"), 1)
    Detalhe.stsaida = Poe_Zero_Esquerda(1 - rs("istsaida"), 1)
    Detalhe.fbateria = Poe_Zero_Esquerda(rs("cfbateria"), 1)
    Detalhe.fviolacao = Poe_Zero_Esquerda(rs("cfviolacao"), 1)
    Detalhe.contador = Poe_Zero_Esquerda(rs("ccontador"), 8)
    Detalhe.placa = Left(rs("cplaca") + "0000", 7)
    Detalhe.antpais = Poe_Zero_Esquerda(rs("cconcpais"), 4)
    Detalhe.antconc = Poe_Zero_Esquerda(rs("cconc"), 5)
    Detalhe.antpraca = Poe_Zero_Esquerda(rs("cconcpraca"), 4)
    Detalhe.antpista = Poe_Zero_Esquerda(rs("cconcpista"), 3)
    Detalhe.antdt = "00000000"
'    Detalhe.antdt = Poe_Zero_Esquerda(rs("cantenadia"), 8)
    Detalhe.anthr = "000000"
'    Detalhe.anthr = Poe_Zero_Esquerda(rs("cantenahora"), 6)
    Detalhe.motimagem = Poe_Zero_Esquerda(rs("cmotimagem"), 2)
    Detalhe.sttransacao = Poe_Zero_Esquerda(rs("cstatustransacao"), 1)
    Detalhe.stmac = Poe_Zero_Esquerda(rs("cstmac"), 8)
    Detalhe.filler = Left("0000000000000000000000000000000", 30)
    Detalhe.crlf = Chr(13) & Chr(10)

    
    'SE FOR EMISSOR AUTOEXPRESSO FAZ AS SUBSTIUICOES DOS CAMPOS
    If Val(rs("Iissuer")) = 100 Or Val(rs("Iissuer")) = 101 Then
        Detalhe.tag = Poe_Zero_Esquerda("290", 5) + Poe_Zero_Esquerda(gsTagPadraoAE, 10)
        Detalhe.acesso = Poe_Zero_Esquerda(gsAcessoPadraoAE, 4)
    End If
    
    
    Put #Datafilenum, , Detalhe
    DoEvents
    rs.MoveNext
    If i Mod 50 = 0 Then MDI!sta_Barra_MDI.Panels(2).Text = "Criando TRN : " + Format(seqfile) + ":" + Format(i)
Loop


' Atualiza tabela de tb_TRNCTL

frase = ""
frase = frase & "SELECT * FROM " + gsPath_DB + ".dbo.tb_TRNCTL "
frase = frase & "WHERE tipo = 'TR' and seqfile = '" & Header.seqfile & "' "
Set rs = dbApp.Execute(frase)

If rs.BOF And rs.EOF Then
    frase = ""
    frase = frase & "INSERT INTO " + gsPath_DB + ".dbo.tb_TRNCTL ("
    frase = frase & "tipo,pais,id,seqfile,arquivo,dtger,hrger,"
    frase = frase & "reginformado,totinformado,datamovimento)"
    frase = frase & " values ("
    frase = frase & "'" & Header.tipo & "',"
    frase = frase & "'" & Header.pais & "',"
    frase = frase & "'" & Header.id & "',"
    frase = frase & "'" & Header.seqfile & "',"
    frase = frase & "'" & file & "',"
    frase = frase & "'" & Header.dtger & "',"
    frase = frase & "'" & Header.hrger & "',"
    frase = frase & "'" & Header.reg & "',"
    frase = frase & "'" & Header.valtot & "',"
    frase = frase & "'" & Pdatamovimento & "')"
Else
    frase = ""
    frase = frase & " update " + gsPath_DB + ".dbo.tb_TRNCTL "
    frase = frase & " set "
    frase = frase & " arquivo = '" & file & "',"
    frase = frase & " dtger = '" & Header.dtger & "',"
    frase = frase & " hrger = '" & Header.hrger & "',"
    frase = frase & " reginformado = '" & Header.reg & "',"
    frase = frase & " totinformado = '" & Header.valtot & "',"
    frase = frase & " datamovimento = '" & Pdatamovimento & "'"
    frase = frase & " WHERE tipo = 'TR' and seqfile = '" & Header.seqfile & "' "
End If
Set rs = dbApp.Execute(frase)

Call GravaEventos(401, " ", gsUser, 0, "Grava TRN - " + file, Format(Header.seqfile), Right(Format(Header.reg), 9))

Close Datafilenum
Call CopyFile(Datafile, gsPath_CGMPEnvia & file)

gsDestFile = CriaDir(gsPath_TRT, "Atualizados\", Mid(file, 1, 6) + "\", "") & file

Call CopyFile(Datafile, gsDestFile)

Kill Datafile

Set rs = Nothing

Exit Sub
'On Error GoTo trataerr
trataerr:
If gbTransferStatus Then
    MsgBoxService "Erro na geração do Arquivo" + Error
    Set rsupdate = dbApp.Execute("rollback tran")
End If
Call TrataErro(app.title, Error, "CriaFileTRN")


End Sub



Public Sub CriaFileTRNAutoExpresso(Pdatamovimento As String, PSeqfile As Integer)
On Error GoTo trataerr
Dim Datafile As String
Dim gsDestFile As String
Dim Datafilenum As Long
Dim rs As New Recordset
Dim rsupdate As New Recordset
Dim frase As String
Dim tipo As String
Dim i As Long
Dim incl As Long
Dim valor As Long
Dim Header As TRHeader
Dim Detalhe As TRdet

Dim file As String
Dim DataMov As Variant
Dim seqfile As String
Dim Qtde    As Long
Dim DataOper As String

rs.CursorType = adOpenStatic

seqfile = Format(PSeqfile, "00000")
file = Format(Now, "YYYYMMDDHHMMSS") + seqfile + ".TRN"

frase = ""
frase = frase & " SELECT sum(cast(ivalor as int)) as Valor, count(*) as qtde "
frase = frase & " from " + gsPath_DB + ".dbo.tb_transacao "
frase = frase & " WHERE "
frase = frase & " Lseqfile = '" & PSeqfile & "' "
frase = frase & " and Lseqreg > 0 "
Set rs = dbApp.Execute(frase)

If (rs.BOF And rs.EOF) Or IsNull(rs(0)) Then
    Exit Sub
Else
    Qtde = rs("qtde")
    valor = rs("valor")
End If


rs.MoveFirst
Header.tipo = "TR"
Header.pais = "0618"
Header.id = Format(Val(gsEst_Codigo), "00000")
Header.seqfile = Mid(file, 15, 5)
Header.dtger = Mid(file, 1, 8)
Header.hrger = Mid(file, 9, 6)
Header.reg = Format(Qtde, "000000")
Header.valtot = Format(valor, "000000000000")
Header.crlf = Chr(13) & Chr(10)

frase = ""
frase = frase & " SELECT * from " + gsPath_DB + ".dbo.tb_transacao "
frase = frase & " WHERE "
frase = frase & " lseqfile = '" & PSeqfile & "' "
frase = frase & " and lseqreg > 0 "
frase = frase & " order by lseqfile,lseqreg "
Set rs = dbApp.Execute(frase)

' chico frmRT!msglbl = "Criando Arquivo " & file
Datafile = gsPath_TRT & file
Datafilenum = FreeFile
Open Datafile For Binary As Datafilenum

'grava header
Put #Datafilenum, 1, Header

'cria detalhe
rs.MoveFirst
i = 0
Do While Not rs.EOF
    i = i + 1
    Detalhe.regtipo = "D"
    Detalhe.regseq = Poe_Zero_Esquerda(rs("Lseqreg"), 6)
    Detalhe.pais = "0618"
    Detalhe.tag = Poe_Zero_Esquerda("290", 5) + Poe_Zero_Esquerda(gsTagPadraoAE, 10)
    Detalhe.acesso = Poe_Zero_Esquerda(gsAcessoPadraoAE, 4)
    Detalhe.dtent = Format(rs("tsentrada"), "yyyymmdd")
    Detalhe.hrent = Format(rs("tsentrada"), "hhmmss")
    Detalhe.stent = Poe_Zero_Esquerda(1 - rs("Istentrada"), 1)
    Detalhe.dtsai = Format(rs("tssaida"), "yyyymmdd")
    Detalhe.hrsai = Format(rs("tssaida"), "hhmmss")
    Detalhe.valor = Poe_Zero_Esquerda(rs("ivalor"), 8)
    Detalhe.stcobranca = Poe_Zero_Esquerda(rs("istcobranca"), 1)
    Detalhe.stsaida = Poe_Zero_Esquerda(1 - rs("istsaida"), 1)
    Detalhe.fbateria = Poe_Zero_Esquerda(rs("cfbateria"), 1)
    Detalhe.fviolacao = Poe_Zero_Esquerda(rs("cfviolacao"), 1)
    Detalhe.contador = Poe_Zero_Esquerda(rs("ccontador"), 8)
    Detalhe.placa = Left(rs("cplaca") + "0000", 7)
    Detalhe.antpais = Poe_Zero_Esquerda(rs("cconcpais"), 4)
    Detalhe.antconc = Poe_Zero_Esquerda(rs("cconc"), 5)
    Detalhe.antpraca = Poe_Zero_Esquerda(rs("cconcpraca"), 4)
    Detalhe.antpista = Poe_Zero_Esquerda(rs("cconcpista"), 3)
    Detalhe.antdt = "00000000"
'    Detalhe.antdt = Poe_Zero_Esquerda(rs("cantenadia"), 8)
    Detalhe.anthr = "000000"
'    Detalhe.anthr = Poe_Zero_Esquerda(rs("cantenahora"), 6)
    Detalhe.motimagem = Poe_Zero_Esquerda(rs("cmotimagem"), 2)
    Detalhe.sttransacao = Poe_Zero_Esquerda(rs("cstatustransacao"), 1)
    Detalhe.stmac = Poe_Zero_Esquerda(rs("cstmac"), 8)
    Detalhe.filler = Left("0000000000000000000000000000000", 30)
    Detalhe.crlf = Chr(13) & Chr(10)
    Put #Datafilenum, , Detalhe
    DoEvents
    rs.MoveNext
    If i Mod 50 = 0 Then MDI!sta_Barra_MDI.Panels(2).Text = "Criando TRN AE : " + Format(seqfile) + ":" + Format(i)
Loop


' Atualiza tabela de tb_TRNCTL

frase = ""
frase = frase & "SELECT * FROM " + gsPath_DB + ".dbo.tb_TRNCTL "
frase = frase & "WHERE tipo = 'TR' and seqfile = '" & Header.seqfile & "' "
Set rs = dbApp.Execute(frase)

If rs.BOF And rs.EOF Then
    frase = ""
    frase = frase & "INSERT INTO " + gsPath_DB + ".dbo.tb_TRNCTL ("
    frase = frase & "tipo,pais,id,seqfile,arquivo,dtger,hrger,"
    frase = frase & "reginformado,totinformado,datamovimento)"
    frase = frase & " values ("
    frase = frase & "'" & Header.tipo & "',"
    frase = frase & "'" & Header.pais & "',"
    frase = frase & "'" & Header.id & "',"
    frase = frase & "'" & Header.seqfile & "',"
    frase = frase & "'" & file & "',"
    frase = frase & "'" & Header.dtger & "',"
    frase = frase & "'" & Header.hrger & "',"
    frase = frase & "'" & Header.reg & "',"
    frase = frase & "'" & Header.valtot & "',"
    frase = frase & "'" & Pdatamovimento & "')"
Else
    frase = ""
    frase = frase & " update " + gsPath_DB + ".dbo.tb_TRNCTL "
    frase = frase & " set "
    frase = frase & " arquivo = '" & file & "',"
    frase = frase & " dtger = '" & Header.dtger & "',"
    frase = frase & " hrger = '" & Header.hrger & "',"
    frase = frase & " reginformado = '" & Header.reg & "',"
    frase = frase & " totinformado = '" & Header.valtot & "',"
    frase = frase & " datamovimento = '" & Pdatamovimento & "'"
    frase = frase & " WHERE tipo = 'TR' and seqfile = '" & Header.seqfile & "' "
End If
Set rs = dbApp.Execute(frase)

Call GravaEventos(401, " ", gsUser, 0, "Grava TRN AE - " + file, Format(Header.seqfile), Right(Format(Header.reg), 9))

Close Datafilenum
Call CopyFile(Datafile, gsPath_CGMPEnvia & file)

gsDestFile = CriaDir(gsPath_TRT, "Atualizados\", Mid(file, 1, 6) + "\", "") & file

Call CopyFile(Datafile, gsDestFile)

Kill Datafile

Set rs = Nothing

Exit Sub
'On Error GoTo trataerr
trataerr:
If gbTransferStatus Then
    MsgBoxService "Erro na geração do Arquivo" + Error
    Set rsupdate = dbApp.Execute("rollback tran")
End If
Call TrataErro(app.title, Error, "CriaFileTRN")


End Sub

Public Sub Atualiza_NextTRN()
On Error GoTo trataerr
Dim rs As New Recordset
Dim frase As String
Dim NextData As Variant
rs.CursorType = adOpenStatic
    
    
Exit Sub
'On Error GoTo trataerr
trataerr:
Call TrataErro(app.title, Error, "Atualiza_NextTRN")


End Sub

Sub CGMPEnvia()
On Error GoTo trataerr
Dim i As Integer
Dim myname As String
    
If Dir(gsPath_CGMPEnvia, vbDirectory) <> "." Then
   MkDir (gsPath & "CGMPEnvia\")
   gsPath_CGMPEnvia = gsPath & "CGMPEnvia\"
End If

i = 0
ReDim mynames(i)
myname = Dir(gsPath_CGMPEnvia & "*.*")
Do While myname <> ""   ' Start the loop.
   mynames(i) = myname
   i = i + 1
   ReDim Preserve mynames(i)
   myname = Dir
Loop
  
For i = 0 To UBound(mynames) - 1
   gsSourceFile = gsPath_CGMPEnvia & mynames(i)
   gsDestFile = gsPath_CGMP_Files_TRN & mynames(i)
   If gsDestFile <> "" Then
      Call CopyFile(gsSourceFile, gsDestFile)
      myname = Dir(gsDestFile)
      If myname <> "" Then ' Start the loop.
         Kill gsSourceFile
      End If
   End If
DoEvents
Next

Exit Sub
'On Error GoTo trataerr
trataerr:
Call TrataErro(app.title, Error, "CGMPEnvia")


End Sub
Sub CGMPEnviaInd(Filename As String, ref As String)
On Error GoTo trataerr
Dim i As Integer
Dim myname As String
Dim gsPath_CGMPEnviaInd As String
Dim retval As Variant

'retval = Shell("NET USE z: /DELETE")
'retval = Shell("NET USE z: \\172.16.242.186\ind_fcmn")
'gsPath_CGMPEnviaInd = "z:\ind_fcmn"

If Dir("z:\" + ref, vbDirectory) <> ref Then
   MkDir ("z:\" + ref)
End If

gsSourceFile = gsPath_REL & Filename
gsDestFile = "z:\" + ref + "\" & Filename
If gsDestFile <> "" Then
   Call CopyFile(gsSourceFile, gsDestFile)
End If
'retval = Shell("NET USE z: /DELETE")

Exit Sub
'On Error GoTo trataerr
trataerr:
Call TrataErro(app.title, Error, "CGMPEnviaInd")
    
End Sub

Sub CGMPRecebeTR()
On Error GoTo trataerr
Dim i As Integer
Dim myname As String
 
If Dir(gsPath_CGMPRecebe, vbDirectory) <> "." Then
   MkDir (gsPath & "CGMPRecebe\")
   gsPath_CGMPRecebe = gsPath & "CGMPRecebe\"
End If


'RECEBE TRF
i = 0
ReDim mynames(i)
myname = Dir(gsPath_CGMP_Files_TRF & "*.*")
Do While myname <> ""   ' Start the loop.
   mynames(i) = myname
   i = i + 1
   ReDim Preserve mynames(i)
   myname = Dir
Loop
For i = 0 To UBound(mynames) - 1
   gsSourceFile = gsPath_CGMP_Files_TRF & mynames(i)
   gsDestFile = gsPath_CGMPRecebe & mynames(i)
   If mynames(i) <> "" Then
      Call CopyFile(gsSourceFile, gsDestFile)
      myname = Dir(gsDestFile)
      If myname <> "" Then ' Start the loop.
         Kill gsSourceFile
      End If
   End If
Next

DoEvents

'RECEBE TRT
i = 0
ReDim mynames(i)
myname = Dir(gsPath_CGMP_Files_TRT & "*.*")
Do While myname <> ""   ' Start the loop.
   mynames(i) = myname
   i = i + 1
   ReDim Preserve mynames(i)
   myname = Dir
Loop
For i = 0 To UBound(mynames) - 1
   gsSourceFile = gsPath_CGMP_Files_TRT & mynames(i)
   gsDestFile = gsPath_CGMPRecebe & mynames(i)
   If mynames(i) <> "" Then
      Call CopyFile(gsSourceFile, gsDestFile)
      myname = Dir(gsDestFile)
      If myname <> "" Then ' Start the loop.
         Kill gsSourceFile
      End If
   End If
Next

DoEvents

'RECEBE MSG
i = 0
ReDim mynames(i)
myname = Dir(gsPath_CGMP_Files_MSG & "*.*")
Do While myname <> ""   ' Start the loop.
   mynames(i) = myname
   i = i + 1
   ReDim Preserve mynames(i)
   myname = Dir
Loop
For i = 0 To UBound(mynames) - 1
   gsSourceFile = gsPath_CGMP_Files_MSG & mynames(i)
   gsDestFile = gsPath_CGMPRecebe & mynames(i)
   If mynames(i) <> "" Then
      Call CopyFile(gsSourceFile, gsDestFile)
      myname = Dir(gsDestFile)
      If myname <> "" Then ' Start the loop.
         Kill gsSourceFile
      End If
   End If
Next

Exit Sub
'On Error GoTo trataerr
trataerr:
Call TrataErro(app.title, Error, "CGMPRecebeTR")

End Sub

Sub CGMPRecebeTG()
On Error GoTo trataerr
Dim i As Integer
Dim myname As String
 
If Dir(gsPath_CGMPRecebe, vbDirectory) <> "." Then
   MkDir (gsPath & "CGMPRecebe\")
   gsPath_CGMPRecebe = gsPath & "CGMPRecebe\"
End If

'RECEBE TAG
i = 0
ReDim mynames(i)
myname = Dir(gsPath_CGMP_Files_TAG & "*.*")
Do While myname <> ""   ' Start the loop.
   mynames(i) = myname
   i = i + 1
   ReDim Preserve mynames(i)
   myname = Dir
Loop
DoEvents

For i = 0 To UBound(mynames) - 1
   gsSourceFile = gsPath_CGMP_Files_TAG & mynames(i)
   gsDestFile = gsPath_CGMPRecebe & mynames(i)
   If mynames(i) <> "" Then
      Call CopyFile(gsSourceFile, gsDestFile)
      myname = Dir(gsDestFile)
      If myname <> "" Then ' Start the loop.
         Kill gsSourceFile
      End If
   End If
Next

DoEvents

'RECEBE TGT
i = 0
ReDim mynames(i)
myname = Dir(gsPath_CGMP_Files_TGT & "*.*")
Do While myname <> ""   ' Start the loop.
   mynames(i) = myname
   i = i + 1
   ReDim Preserve mynames(i)
   myname = Dir
Loop

For i = 0 To UBound(mynames) - 1
   gsSourceFile = gsPath_CGMP_Files_TGT & mynames(i)
   gsDestFile = gsPath_CGMPRecebe & mynames(i) & "_" & Format(Now, "yyyymmddhhmmss")
   If mynames(i) <> "" Then
      Call CopyFile(gsSourceFile, gsDestFile)
      myname = Dir(gsDestFile)
      If myname <> "" Then ' Start the loop.
         Kill gsSourceFile
      End If
   End If
Next


Exit Sub
'On Error GoTo trataerr
trataerr:
Call TrataErro(app.title, Error, "CGMPRecebeTG")

End Sub

Sub CGMPRecebeLCT()
On Error GoTo trataerr
Dim i As Integer
Dim myname As String
 
If Dir(gsPath_CGMPRecebe, vbDirectory) <> "." Then
   MkDir (gsPath & "CGMPRecebe\")
   gsPath_CGMPRecebe = gsPath & "CGMPRecebe\"
End If

'RECEBE TAG
i = 0
ReDim mynames(i)
myname = Dir(gsPath_CGMP_Files_TAG & "*.*")
Do While myname <> ""   ' Start the loop.
   mynames(i) = myname
   i = i + 1
   ReDim Preserve mynames(i)
   myname = Dir
Loop
DoEvents

For i = 0 To UBound(mynames) - 1
   gsSourceFile = gsPath_CGMP_Files_TAG & mynames(i)
   gsDestFile = gsPath_CGMPRecebe & mynames(i)
   If mynames(i) <> "" Then
      Call CopyFile(gsSourceFile, gsDestFile)
      myname = Dir(gsDestFile)
      If myname <> "" Then ' Start the loop.
         Kill gsSourceFile
      End If
   End If
Next

DoEvents

'RECEBE TGT
i = 0
ReDim mynames(i)
myname = Dir(gsPath_CGMP_Files_TGT & "*.*")
Do While myname <> ""   ' Start the loop.
   mynames(i) = myname
   i = i + 1
   ReDim Preserve mynames(i)
   myname = Dir
Loop

For i = 0 To UBound(mynames) - 1
   gsSourceFile = gsPath_CGMP_Files_TGT & mynames(i)
   gsDestFile = gsPath_CGMPRecebe & mynames(i) & "_" & Format(Now, "yyyymmddhhmmss")
   If mynames(i) <> "" Then
      Call CopyFile(gsSourceFile, gsDestFile)
      myname = Dir(gsDestFile)
      If myname <> "" Then ' Start the loop.
         Kill gsSourceFile
      End If
   End If
Next


Exit Sub
'On Error GoTo trataerr
trataerr:
Call TrataErro(app.title, Error, "CGMPRecebeLCT")

End Sub

Sub CGMPRecebeLN()
On Error GoTo trataerr
Dim i As Integer
Dim myname As String
 
If Dir(gsPath_CGMPRecebe, vbDirectory) <> "." Then
   MkDir (gsPath & "CGMPRecebe\")
   gsPath_CGMPRecebe = gsPath & "CGMPRecebe\"
End If

'RECEBE NEL
i = 0
ReDim mynames(i)
myname = Dir(gsPath_CGMP_Files_NEL & "*.*")
Do While myname <> ""   ' Start the loop.
   mynames(i) = myname
   i = i + 1
   ReDim Preserve mynames(i)
   myname = Dir
Loop

For i = 0 To UBound(mynames) - 1
   gsSourceFile = gsPath_CGMP_Files_NEL & mynames(i)
   gsDestFile = gsPath_CGMPRecebe & mynames(i)
   If mynames(i) <> "" Then
      Call CopyFile(gsSourceFile, gsDestFile)
      myname = Dir(gsDestFile)
      If myname <> "" Then ' Start the loop.
         Kill gsSourceFile
      End If
   End If
Next

'RECEBE LNT
i = 0
ReDim mynames(i)
myname = Dir(gsPath_CGMP_Files_LNT & "*.*")
Do While myname <> ""   ' Start the loop.
   mynames(i) = myname
   i = i + 1
   ReDim Preserve mynames(i)
   myname = Dir
Loop
For i = 0 To UBound(mynames) - 1
   gsSourceFile = gsPath_CGMP_Files_LNT & mynames(i)
   gsDestFile = gsPath_CGMPRecebe & mynames(i) & "_" & Format(Now, "yyyymmddhhmmss")
   If mynames(i) <> "" Then
      Call CopyFile(gsSourceFile, gsDestFile)
      myname = Dir(gsDestFile)
      If myname <> "" Then ' Start the loop.
         Kill gsSourceFile
      End If
   End If
Next


Exit Sub
'On Error GoTo trataerr
trataerr:
Call TrataErro(app.title, Error, "CGMPRecebeLN")

End Sub


Public Sub TrataRT(file As String)
On Error GoTo trataerr
Dim Datafile As String
Dim Datafilenum As Long
Dim myname As String
Dim rs As New Recordset
Dim frase As String
Dim aux As String
Dim tipo As String
Dim Header As RTHeader

rs.CursorType = adOpenStatic

Datafile = gsPath_CGMPRecebe & file
Datafilenum = FreeFile
Open Datafile For Binary As Datafilenum

'LE header
Get #Datafilenum, 1, Header

frase = ""
frase = frase & "SELECT * FROM " + gsPath_DB + ".dbo.tb_TRTCTL "
frase = frase & "WHERE tipo = 'RT' and arquivo = '" & file & "' "
Set rs = dbApp.Execute(frase)

If Not rs.BOF And Not rs.EOF Then
    Close Datafilenum
    gsDestFile = gsPath_OUTROS & file
    Call CopyFile(Datafile, gsDestFile)
    Kill Datafile
'    frmRT!msglbl = "Arquivo : " & file & "removido para OUTROS "
Else
    frase = ""
    frase = frase & "SELECT * FROM " + gsPath_DB + ".dbo.tb_trtctl "
    frase = frase & "WHERE tipo = 'RT' and seqfile = '" & Mid(file, 15, 5) & "' "
    Set rs = dbApp.Execute(frase)
    If Not rs.BOF And Not rs.EOF Then
        frase = ""
        frase = frase & "delete " + gsPath_DB + ".dbo.tb_trtctl "
        frase = frase & "WHERE tipo = 'RT' and seqfile = '" & Mid(file, 15, 5) & "' "
        Set rs = dbApp.Execute(frase)
    End If
    frase = ""
    frase = frase & "INSERT INTO " + gsPath_DB + ".dbo.tb_trtctl ("
    frase = frase & "tipo,arquivo,seqfile,dtger,hrger,rejeicao,"
    frase = frase & "reginformado,regencontrado,totinformado,totencontrado)"
    frase = frase & " values ("
    frase = frase & "'" & Header.tipo & "',"
    frase = frase & "'" & file & "',"
    frase = frase & "'" & Header.seqfile & "',"
    frase = frase & "'" & Header.dtger & "',"
    frase = frase & "'" & Header.hrger & "',"
    frase = frase & "'" & Header.rejeicao & "',"
    frase = frase & "'" & Header.reginformado & "',"
    frase = frase & "'" & Header.regencontrado & "',"
    frase = frase & "'" & Header.totinformado & "',"
    frase = frase & "'" & Header.totencontrado & "')"
    Set rs = dbApp.Execute(frase)
    Call GravaEventos(401, " ", gsUser, 0, "Trata RT - " + file, Format(Header.seqfile), Right(Format(Header.reginformado), 9))
    gsDestFile = CriaDir(gsPath_TRT, "Atualizados\", Mid(file, 1, 6) + "\", "") & file
    
    Call CopyFile(Datafile, gsDestFile)
    Close Datafilenum
    Kill Datafile

    'atualiza datamovimento em tb_trtctl
    frase = ""
    frase = frase & " select top 1 datamovimento "
    frase = frase & " from " + gsPath_DB + ".dbo.tb_trnctl "
    frase = frase & " where seqfile = " & Header.seqfile
    Set rs = dbApp.Execute(frase)
    
    If Not rs.BOF And Not rs.EOF Then
        aux = rs(0)
        frase = ""
        frase = frase & " update " + gsPath_DB + ".dbo.tb_trtctl set datamovimento = " & aux
        frase = frase & " where datamovimento is null and seqfile = " & Header.seqfile
        Set rs = dbApp.Execute(frase)
    End If
End If

Exit Sub
'On Error GoTo trataerr
trataerr:
Call TrataErro(app.title, Error, "TrataRT")


End Sub

Public Sub LogErro(modulo As String, texto As String)

On Error GoTo trataerr

Dim arch As Integer

arch = FreeFile
If gsPath = "" Then
    Open app.Path & "\logErro.err" For Append As arch
Else
    Open gsPath & "logErro.err" For Append As arch
End If
 
Print #arch, "Data: " + Format(Now, "dd/MM/yyyy hh:mm:ss") + " - " + modulo + " : " + texto

Close arch

Exit Sub

trataerr:
    Call TrataErro(app.title, Error, "LogErro")

End Sub

Public Sub MsgBoxService(ByVal prompt As String, Optional ByVal buttons As Integer = 0, Optional ByVal title As String = "title", Optional ByVal helpfile As String = "helpfile")
    If gsPath_msgbox = "0" Then
        Call LogErro(title + " - " + helpfile, prompt + " cod: " + CStr(buttons))
    Else
        MsgBox prompt, vbOKOnly, "Alerta"
    End If
End Sub





