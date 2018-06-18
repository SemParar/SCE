VERSION 5.00
Begin VB.Form frmTela_Apresentacao 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5580
   ClientLeft      =   1335
   ClientTop       =   2355
   ClientWidth     =   6735
   FillColor       =   &H00E0E0E0&
   ForeColor       =   &H8000000C&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5580
   ScaleWidth      =   6735
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3645
      Left            =   1200
      Picture         =   "apresent.frx":0000
      ScaleHeight     =   3645
      ScaleWidth      =   4500
      TabIndex        =   2
      Top             =   240
      Width           =   4500
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
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
      Height          =   375
      Left            =   420
      TabIndex        =   3
      Top             =   4680
      Width           =   5895
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Estacionamento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   420
      TabIndex        =   1
      Top             =   4080
      Width           =   5895
   End
   Begin VB.Label lblVersao 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   0
      Top             =   5160
      Width           =   1935
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmTela_Apresentacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnUsuario_Clicou As Boolean

Private Sub Form_Load()
    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)
    lblVersao.Caption = "versão: " & App.Major & "." & App.Minor & "." & App.Revision
    Label2 = gsEst_Nome
End Sub


Private Sub Form_Activate()
    Dim bytTempo_Espera As Byte
    Dim dblInicio_Tempo As Double
    
    'mostro a abertura por 2 segundos
    bytTempo_Espera = 4
    dblInicio_Tempo = Timer
    blnUsuario_Clicou = False
    
    Do While Timer < dblInicio_Tempo + bytTempo_Espera
        DoEvents
        If blnUsuario_Clicou Then Exit Do
    Loop
    If Not blnUsuario_Clicou Then
        fraTela_Apresentacao_Click
    End If
End Sub

Private Sub fraTela_Apresentacao_Click()
    Unload Me
    frmLogin.Show vbModal
    blnUsuario_Clicou = True
End Sub

