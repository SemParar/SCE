Attribute VB_Name = "SCE_TR"

Sub Main()

Dim horario As String
Dim aux As String

' Set Title
App.title = "SCETR"
If App.PrevInstance Then
    End
End If

If Dir(App.Path + "\SCE.ini", vbArchive) = "" Then
'criar o primeiro .ini
    Dim i As Integer
    i = FreeFile
    Open App.Path + "\SCE.ini" For Output As i
    Print #i, "***** Arquivo INI criado em " + Format(Now, "DD/MM/YY hh:mm:ss")
    Close i
    MsgBoxService "Arquivo de Config Criado - SCE.INI"
End If
    
gsPathIniFile = App.Path + "\SCE.ini"
gsPath = MyGetPar(gsPathIniFile, "SISTEMA", "Path", "c:\SCE\")
MySetPar gsPathIniFile, "SISTEMA", "Path", gsPath
gsPathIniFile = gsPath + "SCE.ini"

' Parametros de Operacao
' le ou cria todas as variaveis de parametros

Call LerGlobais

gsstrErro_Posicao = "DbOpen"
If Not DbOpen(gsPath_DS, gsPath_DB) Then
   MsgBoxService "DataBase SCE não pode ser aberto! Aplicação será finalizada!", vbCritical, "Erro"
   End
End If

frase = ""
frase = frase + " use " + gsPath_DB
Set rsGeral = dbApp.Execute(frase)

Dim ver
ver = Format(App.Major) + "." + Format(App.Minor) + "." + Format(App.Revision)
frase = " if exists(select name from " + gsPath_DB + "..sysobjects where name = 'pr_setver') exec "
frase = frase + gsPath_DB + "..pr_setver '" + App.title + "','" + ver + "'"
Set rs = dbApp.Execute(frase)


giTransfere = 10
gbTransferStatus = False

Dim SCECALL As Integer
SCECALL = MyGetPar(gsPathIniFile, "Sistema", "SCETR", 0)
'[x] Elimina LOGIN da versão 9.0.0 ( 1 = 1 )
gsUser = "SCECALL " + Format(SCECALL)
gbytNivel_Acesso_Usuario = gsscetr
Load MDI
MDI.sta_Barra_MDI.Panels(1).Text = gsUser
MDI.Show
MDI.WindowState = 1

Call LimparArquivos

End Sub

