VERSION 5.00
Begin VB.Form frmSalvarConfiguracao 
   Caption         =   "Configuracao"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSalvar 
      Caption         =   "Salvar"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox txtNewValue 
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   2280
      Width           =   4095
   End
   Begin VB.TextBox txtValue 
      Enabled         =   0   'False
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   1440
      Width           =   4095
   End
   Begin VB.TextBox txtParam 
      Enabled         =   0   'False
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label lblNewValue 
      Caption         =   "Novo Valor"
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblValue 
      Caption         =   "Valor"
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblParam 
      Caption         =   "Parametro"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmSalvarConfiguracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSalvar_Click()
Dim retorno As String

gsstrErro_Posicao = "DbOpen"
    If Not DbOpen(gsPath_DS, gsPath_DB) Then
          MsgBoxService "DataBase SCE não pode ser aberto! Aplicação será finalizada!", vbCritical, "Erro"
          End
    End If
    
If Me.txtValue = Me.txtNewValue Then
MsgBox ("Novo valor e valor antigo sao iguais")
Else

retorno = FsetParam("SCE", Me.txtParam, DateTime.Now, Me.txtNewValue)
frmConfig.dg.Refresh
frmConfig.adogrid.Refresh

Unload Me
End If


End Sub
