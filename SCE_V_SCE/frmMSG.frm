VERSION 5.00
Begin VB.Form frmMensagem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mensagem do Sistema"
   ClientHeight    =   5730
   ClientLeft      =   2760
   ClientTop       =   2340
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   7200
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1800
      Top             =   5280
   End
   Begin VB.Frame Frame1 
      Height          =   4692
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7212
      Begin VB.TextBox txtMSG 
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3252
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   720
         Width           =   6972
      End
      Begin VB.Label lblAlerta 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   5
         Top             =   4080
         Width           =   6972
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6972
      End
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   852
      Left            =   5760
      Picture         =   "frmMSG.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Width           =   1092
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Height          =   852
      Left            =   240
      Picture         =   "frmMSG.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4800
      Width           =   1092
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2640
      TabIndex        =   6
      Top             =   5040
      Width           =   1932
   End
End
Attribute VB_Name = "frmMensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lblAlertaAux As String
Dim TempoFixo As Integer

Option Explicit


Private Sub CmdCancel_Click()

giRespMsg = 0
Unload Me

End Sub

Private Sub CmdOK_Click()

giRespMsg = 1
Unload Me

End Sub

Private Sub Form_Activate()
lblAlertaAux = lblAlerta
End Sub

Private Sub Form_Load()

lblAlertaAux = lblAlerta
TempoFixo = 60
Label1 = Time

End Sub

Private Sub Timer1_Timer()

Label1 = Time
TempoFixo = TempoFixo - 1
If TempoFixo <= 0 Then Unload Me

If lblAlertaAux = lblAlerta Then
    lblAlerta = ""
Else
    lblAlerta = lblAlertaAux
End If
End Sub

