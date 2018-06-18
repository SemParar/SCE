VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRelInd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumo Diario"
   ClientHeight    =   6240
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   11550
   Begin VB.CommandButton cmdgeraarq 
      Caption         =   "Trans. Arquivo"
      Height          =   732
      Left            =   7560
      Picture         =   "frmRelInd.frx":0000
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
      Left            =   10200
      Picture         =   "frmRelInd.frx":0442
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
      Width           =   9732
   End
   Begin VB.CommandButton imprimecmd 
      Caption         =   "Imprime"
      Height          =   732
      Left            =   8880
      Picture         =   "frmRelInd.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   1212
   End
   Begin VB.CommandButton Lercmd 
      Caption         =   "Ler Dados"
      Height          =   732
      Left            =   5280
      Picture         =   "frmRelInd.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   1212
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   4332
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   11292
      _ExtentX        =   19923
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
      Width           =   4932
      Begin VB.TextBox TxtMes 
         Height          =   288
         Left            =   2280
         TabIndex        =   5
         Top             =   240
         Width           =   372
      End
      Begin VB.TextBox TxtAno 
         Height          =   288
         Left            =   4080
         TabIndex        =   4
         Text            =   "2004"
         Top             =   240
         Width           =   492
      End
      Begin VB.Label Label1 
         Caption         =   "MES : "
         Height          =   252
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   612
      End
      Begin VB.Label Label2 
         Caption         =   "ANO : "
         Height          =   252
         Left            =   3240
         TabIndex        =   6
         Top             =   240
         Width           =   612
      End
   End
End
Attribute VB_Name = "frmRelInd"
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

frase = ""
frase = frase & "Select max(tsdatamovimento) from tb_transacao "
Set rs = dbApp.Execute(frase)
If Not rs.BOF And Not rs.EOF And Not IsNull(rs(0)) Then
   TxtDia = Format(rs(0), "DD")
   TxtMes = Format(rs(0), "MM")
   TxtAno = Format(rs(0), "YYYY")
Else
   TxtDia = Format(Date, "DD")
   TxtMes = Format(Date, "MM")
   TxtAno = Format(Date, "YYYY")
End If

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

Filename = "IND_" & gsEst_Codigo & "_" & TxtAno + TxtMes + TxtDia & "_" + Format(Now, "YYYYMMDDHHMMSS") & ".HTML"

'Filename = "IND_" & gsEst_Codigo & "_" & TxtAno + TxtMes & ".html"
Call ImprimeHeader(Filename, "Indice de Passagens : " & TxtMes & "/" & TxtAno)
Call Imprimegrid(Filename, Grid1)
Call ImprimeExtra(Filename, extra)
Call ImprimeFooter(Filename, "Impresso em : " + Format(Now, "DD/MM/YY HH:MM:SS"))
Call ImprimeRel(Filename)

Filename = "IND_" & gsEst_Codigo & "_" & TxtAno + TxtMes + TxtDia & "_" + Format(Now, "YYYYMMDDHHMMSS") & ".CSV"

'Filename = "IND_" & gsEst_Codigo & "_" & TxtAno + TxtMes & ".csv"
Call ArqGrid(Filename, Grid1)



End Sub




Private Sub Lercmd_Click()

Dim frase As String
Dim rs As New Recordset
Dim extra As String
Dim dataini As Date
Dim datafim As Date
On Error Resume Next

rs.CursorType = adOpenStatic

dataini = CVDate("01/" + TxtMes + "/" + TxtAno)
datafim = dataini
frase = ""
frase = frase & " " & "pr_relind '" + Format(dataini, "YYYYMMDD") + "','" + Format(datafim, "YYYYMMDD") + "'"
Set rs = dbApp.Execute(frase)

dataini = CVDate("01/" + TxtMes + "/" + TxtAno)
datafim = DateAdd("d", -1, DateAdd("m", 1, dataini))
frase = ""
frase = frase & " " & "pr_relind '" + Format(dataini, "YYYYMMDD") + "','" + Format(datafim, "YYYYMMDD") + "'"
Set rs = dbApp.Execute(frase)

frase = ""
frase = frase + " select"
frase = frase + " distinct(convert(varchar(8),data,112)) as DATA,"
frase = frase + " (select isnull(sum(nro),0) from  tb_indpassagem     T where tb_indpassagem.data = t.data) as TOTAL,"
frase = frase + " (select isnull(sum(nro),0) from  tb_indpassagem     T where tb_indpassagem.data = t.data and sai = 1 and ent = 1) as AA,"
frase = frase + " (select isnull(sum(nro),0) from  tb_indpassagem     T where tb_indpassagem.data = t.data and sai = 0 and ent = 1) as AM,"
frase = frase + " (select isnull(sum(nro),0) from  tb_indpassagem     T where tb_indpassagem.data = t.data and sai = 1 and ent = 0) as MA,"
frase = frase + " (select isnull(sum(nro),0) from  tb_indpassagem     T where tb_indpassagem.data = t.data and sai = 0 and ent = 0) as MM,"
frase = frase + " (select isnull(sum(nro),0) from  tb_indln               T where tb_indpassagem.data = t.data) as LN,"
frase = frase + " (select isnull(sum(nro),0) from  tb_indsc               T where tb_indpassagem.data = t.data) as SC,"
frase = frase + " (select isnull(sum(nro),0) from  tb_indcan              T where tb_indpassagem.data = t.data) as CAN,"
frase = frase + " (select isnull(sum(nro),0) from  tb_indvalorzero    T where tb_indpassagem.data = t.data) as Zero,"
frase = frase + " (select isnull(sum(nro),0) from  tb_indvalor            T where tb_indpassagem.data = t.data) as Valor"
frase = frase + " From"
frase = frase + " tb_indpassagem"
Set rs = dbApp.Execute(frase)

Grid1.Clear
Set Grid1.DataSource = rs


Grid1.TextMatrix(0, 0) = "DATA      "
Grid1.TextMatrix(0, 1) = "TOTAL    "
Grid1.TextMatrix(0, 2) = "AA       "
Grid1.TextMatrix(0, 3) = "AM       "
Grid1.TextMatrix(0, 4) = "MA       "
Grid1.TextMatrix(0, 5) = "MM       "
Grid1.TextMatrix(0, 6) = "LN       "
Grid1.TextMatrix(0, 7) = "SC       "
Grid1.TextMatrix(0, 8) = "CAN      "
Grid1.TextMatrix(0, 9) = "SemV     "
Grid1.TextMatrix(0, 10) = "ComV    "
Call FormataGridx(Grid1, rs)
Grid1.Refresh

rs.MoveFirst
If rs.BOF And rs.EOF Then
    extra = extra + "   ====>  Nenhum Movimento para esta Data "
End If
Text1 = extra

End Sub

