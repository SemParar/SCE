Attribute VB_Name = "SCE_TG"

Sub Main()

Dim horario As String
Dim aux As String

app.title = "SCETG"
If app.PrevInstance Then
    End
End If

If Dir(app.Path + "\SCE.ini", vbArchive) = "" Then
'criar o primeiro .ini
    Dim i As Integer
    i = FreeFile
    Open app.Path + "\SCE.ini" For Output As i
    Print #i, "***** Arquivo INI criado em " + Format(Now, "DD/MM/YY hh:mm:ss")
    Close i
    MsgBoxService "Arquivo de Config Criado - SCE.INI"
End If


If Dir(app.Path + "\listas.txt", vbArchive) = "" Then
'criar o primeiro listas.txt
    Dim j As Integer
    j = FreeFile
    Open app.Path + "\listas.txt" For Output As j
    Close j
End If
    
gsPathIniFile = app.Path + "\SCE.ini"
gsPath = MyGetPar(gsPathIniFile, "SISTEMA", "Path", "c:\SCE\")
MySetPar gsPathIniFile, "SISTEMA", "Path", gsPath
gsPathIniFile = gsPath + "SCE.ini"

gsPathLogListas = app.Path + "\listas.txt"

gsListasFaltantes = "0-0"
MySetPar gsPathLogListas, "LISTAS", "TIV", gsListasFaltantes

gsListasFaltantes = "0-0"
MySetPar gsPathLogListas, "LISTAS", "LCI", gsListasFaltantes
' Parametros de Operacao
' le ou cria todas as variaveis de parametros

Call LerGlobais
    
gsstrErro_Posicao = "DbOpen - SCECAD"
If Not DbOpen(gsPath_DS, gsPath_DBCAD) Then
   MsgBoxService "DataBase SCECAD não pode ser aberto! Aplicação será finalizada!", vbCritical, "ERRO"
   End
End If

frase = ""
frase = frase + " use " + gsPath_DBCAD
Set rsGeral = dbApp.Execute(frase)

Dim ver
ver = Format(app.Major) + "." + Format(app.Minor) + "." + Format(app.Revision)
frase = " if exists(select name from " + gsPath_DB + "..sysobjects where name = 'pr_setver') exec "
frase = frase + gsPath_DB + "..pr_setver '" + app.title + "','" + ver + "'"
Set rs = dbApp.Execute(frase)
frase = ""


Transfere = 10
gbTransferStatus = False

Dim SCECALL As Integer
SCECALL = gsscetg

gsUser = "SCECALL"
gbytNivel_Acesso_Usuario = gsscetg
Load MDI
MDI.sta_Barra_MDI.Panels(1).Text = "MSG " + Format(SCECALL)
MDI.Show
MDI.WindowState = 1

End Sub


