VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRelDia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumo Diario"
   ClientHeight    =   6240
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   11115
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
      Left            =   9720
      Picture         =   "frmRelDia.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4680
      Width           =   1212
   End
   Begin VB.TextBox Text1 
      Height          =   408
      Left            =   120
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   5640
      Width           =   9732
   End
   Begin VB.CommandButton imprimecmd 
      Caption         =   "Imprime"
      Height          =   732
      Left            =   8400
      Picture         =   "frmRelDia.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   1212
   End
   Begin VB.CommandButton Lercmd 
      Caption         =   "Ler Dados"
      Height          =   732
      Left            =   7080
      Picture         =   "frmRelDia.frx":0614
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
      Width           =   10812
      _ExtentX        =   19076
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
      Begin VB.TextBox TxtDia 
         Height          =   288
         Left            =   852
         TabIndex        =   9
         Top             =   240
         Width           =   360
      End
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
      Begin VB.Label Label3 
         Caption         =   "DIA : "
         Height          =   252
         Left            =   252
         TabIndex        =   8
         Top             =   240
         Width           =   1092
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
Attribute VB_Name = "frmRelDia"
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

Filename = "CPE_" & gsEst_Codigo & "_" & TxtAno + TxtMes + TxtDia & "_" + Format(Now, "YYYYMMDDHHMMSS") & ".HTML"

'Filename = "Periodo" & TxtAno + TxtMes + TxtDia & ".html"
Call ImprimeHeader(Filename, "Controle de Transações por Periodo : " & TxtDia & "/" & TxtMes & "/" & TxtAno)
Call Imprimegrid(Filename, Grid1)
Call ImprimeExtra(Filename, extra)
Call ImprimeFooter(Filename, "Impresso em : " + Format(Now, "DD/MM/YY HH:MM:SS"))
Call ImprimeRel(Filename)

End Sub

Private Sub Lercmd_Click()
'resumo de ocupação por periodo dos dias
Dim frase As String
Dim rs As New Recordset
Dim extra As String
rs.CursorType = adOpenStatic

frase = ""
frase = frase & " " & "pr_relperiodo '" + TxtAno + TxtMes + TxtDia + " " + gsEst_Horario + "','60'"
Set rs = dbApp.Execute(frase)

frase = ""
frase = frase & " " & "select"
frase = frase & " " & "convert(varchar(10),cast(data as datetime),103) as Data,"
frase = frase & " " & "convert(varchar(9),horaini,108) as HoraIni,"
frase = frase & " " & "convert(varchar(9),horafim,108) as HoraFim,"
frase = frase & " " & "per as GRP,"
frase = frase & " " & "tent As tent, tsai As tsai, cvalor As ComValor, tolerancia As tolerancia, Promocao As Promocao"
frase = frase & " " & "from tb_periodo"
frase = frase & " " & "Union"
frase = frase & " " & "(select null,null,null,99,sum(tent),sum(tsai),sum(cvalor),sum(tolerancia),sum(promocao) from tb_periodo)"
frase = frase & " " & "order by data,per"

Set rs = dbApp.Execute(frase)

Grid1.Clear
Set Grid1.DataSource = rs
Grid1.TextMatrix(0, 0) = "DIA          "
Grid1.TextMatrix(0, 1) = "HR Inicio     "
Grid1.TextMatrix(0, 2) = "HR Fim        "
Grid1.TextMatrix(0, 3) = "PER     "
Grid1.TextMatrix(0, 4) = "ENT     "
Grid1.TextMatrix(0, 5) = "SAI     "
Grid1.TextMatrix(0, 6) = "Pagas     "
Grid1.TextMatrix(0, 7) = "Toler     "
Grid1.TextMatrix(0, 8) = "Outras    "
Call FormataGridx(Grid1, rs)
Grid1.Refresh

rs.MoveFirst
If rs.BOF And rs.EOF Then
    extra = extra + "   ====>  Nenhum Movimento para esta Data "
End If
Text1 = extra

End Sub

