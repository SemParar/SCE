VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRelRej 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumo Rejeição"
   ClientHeight    =   6000
   ClientLeft      =   90
   ClientTop       =   480
   ClientWidth     =   15480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   15480
   Begin VB.Frame Frame2 
      Caption         =   "Reenvio"
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   6000
      Visible         =   0   'False
      Width           =   10695
      Begin VB.CommandButton cmdReenvio 
         Caption         =   "Reenvio"
         Height          =   375
         Left            =   8040
         TabIndex        =   21
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtdata 
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Text            =   "data"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtReg 
         Height          =   375
         Left            =   6480
         TabIndex        =   18
         Text            =   "Reg"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtTrn 
         Height          =   375
         Left            =   5160
         TabIndex        =   17
         Text            =   "TRN"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TxtTag 
         Height          =   375
         Left            =   1920
         TabIndex        =   16
         Text            =   "Tag"
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.CommandButton Saircmd 
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
      Height          =   855
      Left            =   11880
      Picture         =   "frmRelRejeitados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   120
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   5520
      Width           =   14775
   End
   Begin VB.CheckBox ChkZerados 
      Caption         =   "Excluir Zerados"
      Height          =   252
      Left            =   6480
      TabIndex        =   12
      Top             =   480
      Width           =   1812
   End
   Begin VB.OptionButton Opt1 
      Caption         =   "Mês"
      Height          =   312
      Index           =   1
      Left            =   2160
      TabIndex        =   4
      Top             =   480
      Width           =   732
   End
   Begin VB.OptionButton Opt1 
      Caption         =   "Total"
      Height          =   312
      Index           =   2
      Left            =   3480
      TabIndex        =   3
      Top             =   480
      Width           =   732
   End
   Begin VB.OptionButton Opt1 
      Caption         =   "Dia"
      Height          =   312
      Index           =   0
      Left            =   960
      TabIndex        =   2
      Top             =   480
      Width           =   732
   End
   Begin VB.CommandButton imprimecmd 
      Caption         =   "Imprime"
      Height          =   855
      Left            =   10200
      Picture         =   "frmRelRejeitados.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Lercmd 
      Caption         =   "Ler Dados"
      Height          =   855
      Left            =   8640
      Picture         =   "frmRelRejeitados.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleção"
      Height          =   1212
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   8052
      Begin VB.TextBox TxtDia 
         Height          =   288
         Left            =   960
         TabIndex        =   11
         Top             =   720
         Width           =   1092
      End
      Begin VB.TextBox TxtMes 
         Height          =   288
         Left            =   3840
         TabIndex        =   7
         Top             =   720
         Width           =   1092
      End
      Begin VB.TextBox TxtAno 
         Height          =   288
         Left            =   6480
         TabIndex        =   6
         Text            =   "2004"
         Top             =   720
         Width           =   1092
      End
      Begin VB.Label Label3 
         Caption         =   "DIA : "
         Height          =   252
         Left            =   360
         TabIndex        =   10
         Top             =   720
         Width           =   1092
      End
      Begin VB.Label Label1 
         Caption         =   "MES : "
         Height          =   252
         Left            =   3120
         TabIndex        =   9
         Top             =   720
         Width           =   1092
      End
      Begin VB.Label Label2 
         Caption         =   "ANO : "
         Height          =   252
         Left            =   5880
         TabIndex        =   8
         Top             =   720
         Width           =   1092
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   3855
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      AllowBigSelection=   0   'False
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
         Name            =   "Arial"
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
End
Attribute VB_Name = "frmRelRej"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Filename As String
Dim rs As New Recordset


Private Sub Command1_Click()

End Sub

Private Sub cmdReenvio_Click()

If RTrim(LTrim(txtTrn)) <> "" Then
    frmTagReenvio.Show
    frmTagReenvio.cmdLimpa_Click
    frmTagReenvio.txtTrn = txtTrn
    frmTagReenvio.TxtReg = TxtReg
    frmTagReenvio.TxtTag = Right(TxtTag, 10)
    frmTagReenvio.TxtEmissor = Left(TxtTag, 3)
    frmTagReenvio.cmdPes_Click
End If

End Sub

Private Sub Form_Load()
Dim frase As String

Me.Top = 10
Me.Left = 10

frase = ""
frase = frase & "Select max(dtger) from tb_trfctl"
Set rs = dbApp.Execute(frase)
If Not rs.BOF And Not rs.EOF And Not IsNull(rs(0)) Then
   TxtDia = Mid(rs(0), 7, 2)
   TxtMes = Mid(rs(0), 5, 2)
   TxtAno = Mid(rs(0), 1, 4)
Else
   TxtDia = Format(Date, "DD")
   TxtMes = Format(Date, "MM")
   TxtAno = Format(Date, "YYYY")
End If

Opt1(1).Value = True
ChkZerados.Value = 0
Grid1.Rows = 15

 
Call Lercmd_Click
 
End Sub

Private Sub Form_Unload(Cancel As Integer)

Call ImprimeRelDel(Filename)

End Sub

Private Sub Grid1_Click()

Grid1.ColSel = Grid1.Col
Grid1.RowSel = Grid1.Row
'Grid1.BackColorSel = vbBlue


TxtTag = UCase(LTrim(RTrim(Grid1.TextMatrix(Grid1.Row, 3))))
txtTrn = UCase(LTrim(RTrim(Grid1.TextMatrix(Grid1.Row, 1))))
TxtReg = UCase(LTrim(RTrim(Grid1.TextMatrix(Grid1.Row, 2))))
txtdata = UCase(LTrim(RTrim(Grid1.TextMatrix(Grid1.Row, 0))))
'txtTrn = Grid1.Row
'TxtReg = UCase(LTrim(RTrim(Grid1.TextMatrix(Grid1.RowSel, 2))))



End Sub

Private Sub imprimecmd_Click()
Dim extra() As String

ReDim extra(1)
extra(0) = Text1

Filename = "REJ_" & gsEst_Codigo & "_" & TxtAno + TxtMes + TxtDia & "_" + Format(Now, "YYYYMMDDHHMMSS") & ".HTML"

'Filename = "Rejeitadas" & Format(Date, "YYYYMMDD") & ".html"
Call ImprimeHeader(Filename, "Controle de Transações Rejeitadas")
Call Imprimegrid(Filename, Grid1)
Call ImprimeExtra(Filename, extra)
Call ImprimeFooter(Filename, "Impresso em : " + Format(Now, "DD/MM/YY HH:MM:SS"))
Call ImprimeRel(Filename)
'Call ImprimeRelDel(Filename)


End Sub

Private Sub Lercmd_Click()

Dim frase As String
Dim aux As String
Dim fraseaux As String
Dim extra As String

If Opt1(0) Then
    aux = TxtAno + TxtMes + TxtDia
    extra = "Filtrado pelo Dia " + TxtDia + "/" + TxtMes + "/" + TxtAno
ElseIf Opt1(1) Then
    aux = TxtAno + TxtMes
    extra = "Filtrado pelo Mes " + TxtMes + "/" + TxtAno
Else
    aux = TxtAno
    extra = "Filtrado pelo Ano " + TxtAno
End If

If ChkZerados Then
    fraseaux = "  cast(ta.valor AS INT) > 0 AND"
Else
    fraseaux = ""
End If


frase = ""
frase = frase & " Select "
frase = frase & " isnull((select top 1 convert(varchar,ty.tsdatamovimento,112) from tb_transacao ty where ty.lseqfile = ta.seqfile),'-') as movimento ,"
frase = frase & " cast(ta.seqfile as integer) as SeqFile,"
frase = frase & " cast(ta.seqreg as integer) as SeqReg,"
frase = frase & " substring(str(ta.tag,15),1,5) + ' - ' + substring(str(ta.tag,15),6,10) as TAG,"
frase = frase & " ta.entradadia + ' ' + ta.entradahora as Entrada,"
frase = frase & " isnull((select istentrada from tb_transacao ty where ty.lseqfile = ta.seqfile and ty.lseqreg = ta.seqreg),'-') as E_AUT,"
frase = frase & " ta.saidadia + ' ' + ta.saidahora as Saida,"
frase = frase & " isnull((select istsaida from tb_transacao ty where ty.lseqfile = ta.seqfile and ty.lseqreg = ta.seqreg),'-') as S_AUT,"
frase = frase & " cast(acesso as integer) as Acesso,"
'frase = frase & " str(cast(cast(ta.valor as dec(8,2))/100 as dec(8,2)),10,2) as Valor,"
frase = frase & " dbo.fpoev2(ta.valor) as Valor,"
'frase = frase & " cast(ta.valor as integer) as Valor,"
frase = frase & " cast(ta.Codigo as integer) as Cod,"
frase = frase & " (select top 1 cDescricao from tb_codrej where iCod = ta.codigo) as Descricao"
frase = frase & " From"
frase = frase & " tb_relfin1 ta"
frase = frase & " Where"
frase = frase & fraseaux
frase = frase & " ta.saidadia like '" & aux & "%' /*@@psaidadia*/"
Set rs = dbApp.Execute(frase)
Grid1.Clear
Set Grid1.DataSource = rs

Grid1.TextMatrix(0, 0) = "Data       "
Grid1.TextMatrix(0, 1) = "Seq   "
Grid1.TextMatrix(0, 2) = "Reg   "
Grid1.TextMatrix(0, 3) = "TAG                 "
Grid1.TextMatrix(0, 4) = "ENTRADA            "
Grid1.TextMatrix(0, 5) = "EA "
Grid1.TextMatrix(0, 6) = "SAIDA              "
Grid1.TextMatrix(0, 7) = "SA "
Grid1.TextMatrix(0, 8) = "AC  "
Grid1.TextMatrix(0, 9) = "VALOR        "
Grid1.TextMatrix(0, 10) = "COD   "
Grid1.TextMatrix(0, 11) = "DESCRIÇÃO                       "
Grid1.AllowBigSelection = False
Call FormataGridx(Grid1, rs)
Grid1.Refresh

If Grid1.Rows > 1 Then
    Grid1.Row = 1
    Grid1.Col = 0
End If
TxtTag = "" ' UCase(LTrim(RTrim(Grid1.TextMatrix(Grid1.Row, 3))))
txtTrn = "" ' UCase(LTrim(RTrim(Grid1.TextMatrix(Grid1.Row, 1))))
TxtReg = "" 'UCase(LTrim(RTrim(Grid1.TextMatrix(Grid1.Row, 2))))
txtdata = "" 'UCase(LTrim(RTrim(Grid1.TextMatrix(Grid1.Row, 0))))


extra = extra + "   ====>  " + Format(Grid1.Rows - 1, "0000") + " -- Rejeições"
Text1 = extra

End Sub

Private Sub saircmd_Click()

Unload Me

End Sub

Private Sub Text2_Change()

End Sub
