VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmManLer 
   Caption         =   "Ler Tabelas"
   ClientHeight    =   5208
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   10908
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   10.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5208
   ScaleWidth      =   10908
   Begin VB.CommandButton cmdAbreTelas 
      Caption         =   "Abre Telas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   9000
      TabIndex        =   7
      Top             =   4680
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Width           =   2772
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1572
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   480
      Width           =   8412
   End
   Begin VB.CommandButton saircmd 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   9120
      TabIndex        =   3
      Top             =   1440
      Width           =   1572
   End
   Begin VB.CommandButton imprimecmd 
      Caption         =   "Imprime"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   9120
      TabIndex        =   2
      Top             =   960
      Width           =   1572
   End
   Begin VB.CommandButton Lercmd 
      Caption         =   "Ler Dados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   9120
      TabIndex        =   0
      Top             =   480
      Width           =   1572
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   2412
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   10452
      _ExtentX        =   18436
      _ExtentY        =   4255
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   9
   End
   Begin VB.Label Label2 
      Caption         =   "Frase :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1332
   End
End
Attribute VB_Name = "frmManLer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Filename As String

Private Sub Form_Load()

Me.Top = 10
Me.Left = 10

Call Lercmd_Click
 
End Sub

Private Sub Form_Unload(Cancel As Integer)

Call ImprimeRelDel(Filename)

End Sub

Private Sub imprimecmd_Click()
Dim extra() As String

If Text2 <> "" Then
    ReDim extra(1)
    extra(0) = "Filtro : " + Text2
    Filename = "select" & Format(Date, "YYYYMMDDhhmmss") & ".html"
    Call ImprimeHeader(Filename, "select")
    Call ImprimeExtra(Filename, extra)
    Call Imprimegrid(Filename, Grid1)
    Call ImprimeFooter(Filename, "Impresso em : " + Format(Now, "DD/MM/YY HH:MM:SS"))
    Call ImprimeRel(Filename)
End If

End Sub

Private Sub Lercmd_Click()

Dim frase As String
Dim rs As New Recordset
Dim aux As String
Dim fraseaux As String
Dim extra As String
On Error Resume Next
rs.CursorType = adOpenStatic

If Text2 <> "" Then
    frase = Text2
    Set rs = dbApp.Execute(frase)
    Grid1.Clear
    Set Grid1.DataSource = rs
    Call FormataGridx(Grid1, rs)
End If

End Sub

Private Sub saircmd_Click()

Unload Me

End Sub

Private Sub Text1_LostFocus()

    If RTrim(LTrim(UCase(Text1))) = "SCE" Then
        Text2.Enabled = True
    Else
        Text2.Enabled = False
    End If

End Sub
