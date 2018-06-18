VERSION 5.00
Begin VB.Form frmConfirma_Senha 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirme Senha"
   ClientHeight    =   1752
   ClientLeft      =   1620
   ClientTop       =   5448
   ClientWidth     =   3516
   Icon            =   "CONF_SEN.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1752
   ScaleWidth      =   3516
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
      ScaleHeight     =   1284
      ScaleWidth      =   3444
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
         Enabled         =   0   'False
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
Attribute VB_Name = "frmConfirma_Senha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancela_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    txtSenha.Text = ""
End Sub
Private Sub Form_Load()
   
    txtNome_Usuario.Text = frmOperadores_Sistema.cmbUserName.Text
End Sub
Private Sub CmdOK_Click()
    frmOperadores_Sistema.Password = txtSenha.Text
    Unload Me
End Sub

Private Sub txtSenha_GotFocus()
    txtSenha.SelStart = 0
    txtSenha.SelLength = Len(txtSenha.Text)
End Sub
