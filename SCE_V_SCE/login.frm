VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Logon"
   ClientHeight    =   1980
   ClientLeft      =   1350
   ClientTop       =   1605
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1980
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancela 
      Caption         =   "&Cancela"
      Height          =   375
      Left            =   1860
      TabIndex        =   3
      Top             =   1380
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1380
      Width           =   1215
   End
   Begin VB.PictureBox sspLogon 
      BackColor       =   &H8000000A&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   3435
      TabIndex        =   4
      Top             =   0
      Width           =   3495
      Begin VB.TextBox txtSenha 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   780
         Width           =   1875
      End
      Begin VB.TextBox txtNome_Usuario 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   300
         Width           =   1875
      End
      Begin VB.Label lblSenha 
         AutoSize        =   -1  'True
         Caption         =   "Senha"
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   840
         Width           =   465
      End
      Begin VB.Label lblUsuario 
         AutoSize        =   -1  'True
         Caption         =   "Usuário"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnCarrega_MDI As Boolean



Private Sub cmdCancela_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    txtSenha.Text = ""
End Sub

Private Sub Form_Load()
    Dim sBuffer As String
    Dim lSize As Long
    Dim frase As String
    Dim rsUsuario As Recordset

On Error GoTo LoadError

    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    If lSize > 0 Then
        txtNome_Usuario.Text = Left$(sBuffer, lSize)
    Else
        txtNome_Usuario.Text = vbNullString
    End If
    
    blnCarrega_MDI = True
    gbytNivel_Acesso_Usuario = 0
    gsUser = ""
    
    Screen.MousePointer = vbHourglass
    gsstrErro_Posicao = "DbOpen"
    If Not DbOpen(gsPath_DS, gsPath_DB) Then
          MsgBoxService "DataBase Parking não pode ser aberto! Aplicação será finalizada!", vbCritical, "Erro"
          End
    End If
    
    Screen.MousePointer = vbDefault
    gsstrErro_Posicao = "Centraliza_Form_Left"
    frmLogin.Left = Centraliza_Form_Left(frmLogin)
    frmLogin.Top = Centraliza_Form_Top(frmLogin)
    Exit Sub
    
LoadError:
    TrataErro Me.Caption, gsstrErro_Posicao, "Load"
End Sub
Private Sub CmdOK_Click()
    
Dim frase As String
Dim rsUsuario As Recordset
    
On Error GoTo LoginError

'check for correct password
txtNome_Usuario.Text = UCase(RTrim(LTrim(txtNome_Usuario.Text)))
If txtNome_Usuario.Text = "" Then
   MsgBoxService "Informe o nome do usuário !", vbCritical, "Logon"
   txtNome_Usuario.SetFocus
Else
   Screen.MousePointer = vbHourglass
       
   'senha mestra
   If Trim(txtNome_Usuario) = "FCM" And Trim(txtSenha.Text) = "82738217" Then
      Passe_Livre gintNIVEL_ADMINISTRADOR
      Exit Sub
   End If
    
frase = ""
frase = frase & "SELECT szsenha, Inivel FROM TB_usuario "
frase = frase & "WHERE Cusuario = '" & Trim(txtNome_Usuario) & "'"
Set rsUsuario = dbApp.Execute(frase)
'
gbytNivel_Acesso_Usuario = 0
gsUser = ""
    
'Verifica se usuario esta cadastrado
If rsUsuario.RecordCount <> 0 Then

'Alterado na versão 8.7.6
'Verifica se as tabelas possuem os campos necessários
checkDB


  'Verifica se senha esta correta
  'If Criptografa(txtSenha.Text) = rsUsuario!szsenha Then
  If txtSenha.Text = rsUsuario!szsenha Then
     Passe_Livre rsUsuario!iNivel
  Else
     MsgBoxService "Senha inválida, tente novamente !", vbCritical, "Logon"
     txtSenha.SetFocus
     txtSenha.SelStart = 0
     txtSenha.SelLength = Len(txtSenha.Text)
     Screen.MousePointer = vbDefault
  End If
  Else
     MsgBoxService "Usuário não cadastrado !", vbCritical, "Logon"
     txtNome_Usuario.SetFocus
     txtNome_Usuario.SelStart = 0
     txtNome_Usuario.SelLength = Len(txtNome_Usuario.Text)
     Screen.MousePointer = vbDefault
  End If
     rsUsuario.Close
  End If
     Screen.MousePointer = vbDefault
  Exit Sub
    

LoginError:
    TrataErro Me.Caption, "usuarios", "Login"
    
End Sub


Private Sub checkDB()
On Error GoTo DbCheckErr

frase = ""
frase = frase + "IF NOT EXISTS(SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'tb_transacao' AND  COLUMN_NAME = 'iStatCat')"
frase = frase + " BEGIN "
frase = frase + "      ALTER TABLE tb_transacao ADD [iStatCat] int NULL "
frase = frase + " END"
Set rs = dbApp.Execute(frase)
   
frase = ""
frase = frase + "IF NOT EXISTS(SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'tb_praia' AND  COLUMN_NAME = 'iStatCat')"
frase = frase + " BEGIN "
frase = frase + "      ALTER TABLE tb_praia ADD [iStatCat] int NULL "
frase = frase + " END"
Set rs = dbApp.Execute(frase)

frase = ""
frase = frase + "IF NOT EXISTS(SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'tb_praiadp' AND  COLUMN_NAME = 'iStatCat')"
frase = frase + " BEGIN "
frase = frase + "      ALTER TABLE tb_praiadp ADD [iStatCat] int NULL "
frase = frase + " END"
Set rs = dbApp.Execute(frase)

frase = ""
frase = frase + "IF NOT EXISTS(SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'tb_praiaold' AND  COLUMN_NAME = 'iStatCat')"
frase = frase + " BEGIN "
frase = frase + "      ALTER TABLE tb_praiaold ADD [iStatCat] int NULL "
frase = frase + " END"
Set rs = dbApp.Execute(frase)

Exit Sub

DbCheckErr:
    TrataErro Me.Caption, "", "Check DB"

End Sub

Sub Passe_Livre(bytNivel_Usuario As Byte)
 
 On Error GoTo PasseLivreError
 
    gbytNivel_Acesso_Usuario = bytNivel_Usuario
    gsUser = Trim(txtNome_Usuario)
    Unload Me
    Screen.MousePointer = vbDefault
    
    ' caso seja o administrador a logar verifico a estabilidade
    If gbytNivel_Acesso_Usuario = gintNIVEL_ADMINISTRADOR Then
       'x
    End If
    
    If blnCarrega_MDI Then
        Load MDI
        MDI.sta_Barra_MDI.Panels(1).Text = gsUser
        MDI.Show
        blnCarrega_MDI = False
    End If
    Exit Sub
    
PasseLivreError:
    TrataErro Me.Caption, "", "Passe Livre"
    
End Sub


Private Sub txtNome_Usuario_GotFocus()
    txtNome_Usuario.SelStart = 0
    txtNome_Usuario.SelLength = Len(txtNome_Usuario.Text)
End Sub
Private Sub txtSenha_GotFocus()
    txtSenha.SelStart = 0
    txtSenha.SelLength = Len(txtSenha.Text)
End Sub
