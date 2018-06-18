VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRelPerm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumo Permanencia"
   ClientHeight    =   6240
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   10305
   Begin VB.CommandButton cmdgeraarq 
      Caption         =   "Trans. Arquivo"
      Height          =   732
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4680
      Width           =   1212
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4680
      Width           =   1212
   End
   Begin VB.TextBox Text1 
      Height          =   408
      Left            =   120
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   5640
      Width           =   10095
   End
   Begin VB.CommandButton imprimecmd 
      Caption         =   "Imprime"
      Height          =   732
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   1212
   End
   Begin VB.CommandButton Lercmd 
      Caption         =   "Ler Dados"
      Height          =   732
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   1212
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   4335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   7646
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   9
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleção"
      Height          =   732
      Left            =   120
      TabIndex        =   3
      Top             =   4680
      Width           =   3975
      Begin VB.TextBox TxtMes 
         Height          =   288
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   372
      End
      Begin VB.TextBox TxtAno 
         Height          =   288
         Left            =   2760
         TabIndex        =   4
         Text            =   "2004"
         Top             =   240
         Width           =   492
      End
      Begin VB.Label Label1 
         Caption         =   "MES : "
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "ANO : "
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmRelPerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Filename As String


Private Sub cmdsair_Click()

Unload Me

End Sub

Private Sub Form_Load()
Dim frase As String
Dim rs As New Recordset
rs.CursorType = adOpenStatic

Me.Top = 10
Me.Left = 10

TxtDia = Format(DateAdd("m", -1, Date), "DD")
TxtMes = Format(DateAdd("m", -1, Date), "MM")
TxtAno = Format(DateAdd("m", -1, Date), "YYYY")
Text1 = "Digite a DATA e mande LER DADOS"
'Call Lercmd_Click
 
End Sub

Private Sub Form_Unload(Cancel As Integer)

Call ImprimeRelDel(Filename)

End Sub

Private Sub imprimecmd_Click()
Dim extra() As String

ReDim extra(1)
extra(0) = Text1

Filename = "PER_" & gsEst_Codigo & "_" & TxtAno + TxtMes + TxtDia & "_" + Format(Now, "YYYYMMDDHHMMSS") & ".html"
Call ImprimeHeader(Filename, "Estacionamento: " + gsEst_Codigo + "  ( Indice de Permanencia : " & TxtMes & "/" & TxtAno + " )")
Call Imprimegrid(Filename, Grid1)
Call ImprimeExtra(Filename, extra)
Call ImprimeFooter(Filename, "Impresso em : " + Format(Now, "DD/MM/YY HH:MM:SS"))
Call ImprimeRel(Filename)

Filename = "PER_" & gsEst_Codigo & "_" & TxtAno + TxtMes + TxtDia & "_" + Format(Now, "YYYYMMDDHHMMSS") & ".CSV"
Call ArqGrid(Filename, Grid1)

End Sub


Private Sub Lercmd_Click()
'permanencias
Dim frase As String
Dim rs As New Recordset
Dim extra As String
Dim dataini As Date
Dim datafim As Date
On Error Resume Next


rs.CursorType = adOpenStatic
Lercmd.Enabled = False

If Val(TxtMes) = 0 Then
dataini = CVDate("01/01/" + TxtAno)
datafim = DateAdd("y", 1, dataini)
frase = ""
frase = frase + " SELECT"
frase = frase + " convert(char(8),tsentrada,112) as dia,"
frase = frase + " datepart(dw,tsentrada) as diasemana,"
frase = frase + " DATEDIFF(n,tsentrada,tssaida)/60+1 as permanencia,"
frase = frase + " count(*) as qtde"
frase = frase + " From tb_transacao"
frase = frase + " Where"
frase = frase + " tsentrada >= '" + Format(dataini, "YYYYMMDD") + "' and tsentrada < '" + Format(datafim, "YYYYMMDD") + "'"
frase = frase + " Group By"
frase = frase + " convert(char(8),tsentrada,112),"
frase = frase + " datepart(dw,tsentrada),"
frase = frase + " DATEDIFF(n,tsentrada,tssaida)/60+1"
frase = frase + " Order By"
frase = frase + " convert(char(8),tsentrada,112),"
frase = frase + " datepart(dw,tsentrada),"
frase = frase + " DATEDIFF(n,tsentrada,tssaida)/60+1"
Else
dataini = CVDate("01/" + TxtMes + "/" + TxtAno)
datafim = DateAdd("m", 1, dataini)  ' DateAdd("d", -1, DateAdd("m", 1, dataini))
frase = ""
frase = frase + " SELECT"
frase = frase + " convert(char(8),tsentrada,112) as dia,"
frase = frase + " datepart(dw,tsentrada) as diasemana,"
frase = frase + " DATEDIFF(n,tsentrada,tssaida)/60+1 as permanencia,"
frase = frase + " count(*) as qtde"
frase = frase + " From tb_transacao"
frase = frase + " Where"
frase = frase + " tsentrada >= '" + Format(dataini, "YYYYMMDD") + "' and tsentrada < '" + Format(datafim, "YYYYMMDD") + "'"
frase = frase + " Group By"
frase = frase + " convert(char(8),tsentrada,112),"
frase = frase + " datepart(dw,tsentrada),"
frase = frase + " DATEDIFF(n,tsentrada,tssaida)/60+1"
frase = frase + " Order By"
frase = frase + " convert(char(8),tsentrada,112),"
frase = frase + " datepart(dw,tsentrada),"
frase = frase + " DATEDIFF(n,tsentrada,tssaida)/60+1"
End If

Set rs = dbApp.Execute(frase)

Grid1.Clear
Set Grid1.DataSource = rs
Grid1.TextMatrix(0, 0) = "DIA               "
Grid1.TextMatrix(0, 1) = "DIA SEMANA               "
Grid1.TextMatrix(0, 2) = "PERMANENCIA              "
Grid1.TextMatrix(0, 3) = "QUANTIDADE               "
Call FormataGridx(Grid1, rs)
Grid1.Refresh
rs.MoveFirst
If rs.BOF And rs.EOF Then
    extra = extra + "   ====>  Nenhum Movimento para esta Data "
Else
    extra = extra + "   "
End If
Text1 = extra


Lercmd.Enabled = True
End Sub

