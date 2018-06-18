VERSION 5.00
Begin VB.Form frmConfigBanco 
   Caption         =   "Configuracao"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDBCAD 
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Text            =   "SCECAD"
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton btLimpar 
      Caption         =   "Limpar"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton btSalvar 
      Caption         =   "Salvar"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtPWD 
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Text            =   "parkavi"
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox txtUID 
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Text            =   "parkavi"
      Top             =   1680
      Width           =   2655
   End
   Begin VB.TextBox txtDB 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Text            =   "SCEWXXXX"
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox txtDs 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Text            =   "172.188.0.2"
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Database SCECAD"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Senha"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Usuario"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Datasource"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Database SCE"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmConfigBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btLimpar_Click()
txtDs = ""
txtDB = ""
txtDBCAD = ""
txtUID = ""
txtPWD = ""
End Sub

Private Sub btSalvar_Click()
gsPWD = txtPWD
gsuid = txtUID
gsPath_DS = txtDs
gsPath_DBCAD = txtDBCAD
gsPath_DB = txtDB
Unload Me
End Sub
