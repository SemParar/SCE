Attribute VB_Name = "SCE_SCE"
Option Explicit

Sub Main()

'On Error GoTo trataerr

Dim horario As String

' Set Title
app.title = "SCE"
If app.PrevInstance Then
    MsgBoxService "Aplicação SCE já esta sendo Executada"
    End
End If

If Dir(app.Path + "\" + app.title + ".ini", vbArchive) = "" Then
'criar o primeiro .ini
    Dim i As Integer
    i = FreeFile
    Open app.Path + "\" + app.title + ".ini" For Output As i
    Print #i, "***** Arquivo INI criado em " + Format(Now, "DD/MM/YY hh:mm:ss")
    Close i
    
    frmConfigBanco.Show

    Do
        If frmConfigBanco.Visible = False Then
            Exit Do
        End If
        DoEvents
    Loop
        
        gsPathIniFile = app.Path + "\" + app.title + ".ini"
     
        MySetPar gsPathIniFile, "SISTEMA", "PWD", gsPWD
        MySetPar gsPathIniFile, "SISTEMA", "UID", gsuid
        MySetPar gsPathIniFile, "SISTEMA", "DS", gsPath_DS
        MySetPar gsPathIniFile, "SISTEMA", "DBCAD", gsPath_DBCAD
        MySetPar gsPathIniFile, "SISTEMA", "DB", gsPath_DB

End If
    
' Parametros de Operacao
' le ou cria todas as variaveis de parametros

gsPathIniFile = app.Path + "\" + app.title + ".ini"
gsPathLogListas = app.Path + "\listas.txt"

gsPath = MyGetPar(gsPathIniFile, "SISTEMA", "Path", "")
If gsPath = "" Then gsPath = app.Path + "\"
MySetPar gsPathIniFile, "SISTEMA", "Path", gsPath
gsPathIniFile = gsPath + app.title + ".ini"


Call LerGlobais

gsstrErro_Posicao = "DbOpen"
If Not DbOpen(gsPath_DS, gsPath_DB) Then
   MsgBoxService "DataBase SCE não pode ser aberto! Aplicação será finalizada!", vbCritical, "Erro"
   End
End If

Dim ver
ver = Format(app.Major) + "." + Format(app.Minor) + "." + Format(app.Revision)
frase = " if exists(select name from " + gsPath_DB + "..sysobjects where name = 'pr_setver') exec "
frase = frase + gsPath_DB + "..pr_setver '" + app.title + "','" + ver + "'"
Set rs = dbApp.Execute(frase)



'giTransfere = 10 'forca a primera busca 1 minutos
'gbTransferStatus = False

frmTela_Apresentacao.Show vbModal

Call LimparArquivos



Exit Sub

trataerr:
    Call TrataErro(app.title, "SCE Main", "SCE Main")

End Sub

