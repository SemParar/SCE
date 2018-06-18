VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRelOperacional 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumo Diario"
   ClientHeight    =   6015
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   14460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   14460
   Begin VB.CommandButton cmdgeraarq 
      Caption         =   "Trans. Arquivo"
      Height          =   732
      Left            =   7560
      Picture         =   "frmRelOperacional.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   360
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
      Picture         =   "frmRelOperacional.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   360
      Width           =   1212
   End
   Begin VB.TextBox Text1 
      Height          =   408
      Left            =   120
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   7080
      Width           =   11295
   End
   Begin VB.CommandButton imprimecmd 
      Caption         =   "Imprime"
      Height          =   732
      Left            =   8880
      Picture         =   "frmRelOperacional.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1212
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   6588
      _Version        =   393216
      Cols            =   10
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
      _Band(0).Cols   =   10
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleção DATA SAÍDA"
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   5175
      Begin VB.TextBox txtDiaFim 
         Height          =   288
         Left            =   2160
         TabIndex        =   17
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtAnoFIm 
         Height          =   288
         Left            =   3840
         TabIndex        =   16
         Text            =   "2004"
         Top             =   960
         Width           =   492
      End
      Begin VB.TextBox txtMesFim 
         Height          =   288
         Left            =   3000
         TabIndex        =   15
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtDiaIni 
         Height          =   288
         Left            =   2160
         TabIndex        =   11
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox TxtMesIni 
         Height          =   288
         Left            =   3000
         TabIndex        =   5
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox TxtAnoIni 
         Height          =   288
         Left            =   3840
         TabIndex        =   4
         Text            =   "2004"
         Top             =   480
         Width           =   492
      End
      Begin VB.Label Label5 
         Caption         =   "Data Final :"
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Data Inicial :"
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Dia"
         Height          =   255
         Left            =   2280
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Mês "
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Ano"
         Height          =   255
         Left            =   3960
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton Lercmd 
      Caption         =   "Ler Dados"
      Height          =   732
      Left            =   5520
      Picture         =   "frmRelOperacional.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1212
   End
End
Attribute VB_Name = "frmRelOperacional"
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

'' pega a data do server
frase = ""
frase = frase & "Select getdate() "
Set rs = dbApp.Execute(frase)
Dim dataHoje As Date
If Not rs.BOF And Not rs.EOF And Not IsNull(rs(0)) Then
   ' -- data de hoje server
   dataHoje = rs(0)
Else
   ' -- data de hoje local
   dataHoje = Now
End If
   
Dim mesAtual As Integer
Dim mesTratado As Integer
Dim anoTratado As Integer
mesAtual = Month(dataHoje)

If Month(dataHoje) > 1 Then
    mesTratado = Month(dataHoje) - 1
    anoTratado = Year(dataHoje)
Else
    mesTratado = 12
    anoTratado = Year(dataHoje) - 1
End If

   
' -- preenche a data final com 30 dias antes
Dim dataini As Date
Dim dataininext As Date
Dim datafim As Date
Dim ultimodiamesanterior As Date

dataini = CVDate("01/" + Format(mesTratado, "00") + "/" + Format(anoTratado, "0000"))

daininext = CVDate("01/" + Format(mesAtual, "00") + "/" + Format(anoTratado, "0000"))
datafim = DateAdd("d", -1, daininext)

txtDiaIni = Format(Day(dataini), "00")
TxtMesIni = Format(Month(dataini), "00")
TxtAnoIni = Year(dataini)

txtDiaFim = Format(Day(datafim), "00")
txtMesFim = Format(Month(datafim), "00")
txtAnoFIm = Year(datafim)


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

Filename = "OPE_" & gsEst_Codigo & "_" & TxtAnoIni + TxtMesIni + txtDiaIni & "_" + Format(Now, "YYYYMMDDHHMMSS") & ".HTML"

'Filename = "IND_" & gsEst_Codigo & "_" & TxtAnoIni + TxtMesIni& ".html"
Call ImprimeHeader(Filename, "Controle de Operações : " & TxtMesIni & "/" & TxtAnoIni)
Call Imprimegrid(Filename, Grid1)
Call ImprimeExtra(Filename, extra)
Call ImprimeFooter(Filename, "Impresso em : " + Format(Now, "DD/MM/YY HH:MM:SS"))
Call ImprimeRel(Filename)

Filename = "OPE_" & gsEst_Codigo & "_" & TxtAnoIni + TxtMesIni + txtDiaIni & "_" + Format(Now, "YYYYMMDDHHMMSS") & ".CSV"

'Filename = "IND_" & gsEst_Codigo & "_" & TxtAnoIni + TxtMesIni& ".csv"
Call ArqGrid(Filename, Grid1)



End Sub




Private Sub Lercmd_Click()

Dim frase As String
Dim rs As New Recordset
Dim extra As String
Dim dataini As Date
Dim datafim As Date
On Error Resume Next

'-- MsgBox (mvDataInicio.Value)

rs.CursorType = adOpenStatic

dataini = CVDate(txtDiaIni + "/" + TxtMesIni + "/" + TxtAnoIni)
datafim = CVDate(txtDiaFim + "/" + txtMesFim + "/" + txtAnoFIm)

frase = ""
frase = frase & " " & "exec dbo.pr_relOperacao '" + Format(dataini, "YYYYMMDD 00:00") + "','" + Format(datafim, "YYYYMMDD 23:59") + "'"
Set rs = dbApp.Execute(frase)

Grid1.Clear

Set Grid1.DataSource = rs

Grid1.ColWidth(0) = 1000
Grid1.ColAlignment(0) = 3
Grid1.ColAlignment(1) = 6
Grid1.ColAlignment(2) = 6
Grid1.ColAlignment(3) = 6
Grid1.ColAlignment(4) = 6
Grid1.ColAlignment(5) = 6
Grid1.ColAlignment(6) = 6

'9.2.3
Grid1.ColAlignment(7) = 6
Grid1.ColAlignment(8) = 6

Grid1.TextMatrix(0, 0) = "Data        "
Grid1.TextMatrix(0, 1) = "Qtde Total   "
Grid1.TextMatrix(0, 2) = "Qtde Enviada   "
Grid1.TextMatrix(0, 3) = "Valor Enviado  "
Grid1.TextMatrix(0, 4) = "Qtde Zerada     "
Grid1.TextMatrix(0, 5) = "Qtde  Rejeitada "
Grid1.TextMatrix(0, 6) = "Valor Rejeitado "
Grid1.TextMatrix(0, 7) = "Qtde Aceita "
Grid1.TextMatrix(0, 8) = "Valor Aceito "


Call FormataGridx(Grid1, rs)
Grid1.Refresh

rs.MoveFirst
If rs.BOF And rs.EOF Then
    extra = extra + "   ====>  Nenhum Movimento para esta Data "
End If
Text1 = extra

End Sub

