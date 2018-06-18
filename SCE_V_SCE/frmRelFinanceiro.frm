VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRelFin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumo Financeiro"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   16485
   Begin VB.CommandButton saircmd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11880
      Picture         =   "frmRelFinanceiro.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   360
      Width           =   1332
   End
   Begin VB.OptionButton Opt1 
      Caption         =   "Mês de Movimento"
      Height          =   312
      Index           =   0
      Left            =   480
      TabIndex        =   10
      Top             =   360
      Width           =   1812
   End
   Begin VB.OptionButton Opt1 
      Caption         =   "Total"
      Height          =   312
      Index           =   2
      Left            =   6000
      TabIndex        =   5
      Top             =   360
      Width           =   2052
   End
   Begin VB.OptionButton Opt1 
      Caption         =   "Mês de Pagamento"
      Height          =   312
      Index           =   1
      Left            =   3240
      TabIndex        =   4
      Top             =   360
      Width           =   1812
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleção"
      Height          =   1212
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8052
      Begin VB.CheckBox ChkZerados 
         Caption         =   "Excluir Zerados"
         Height          =   252
         Left            =   5880
         TabIndex        =   11
         Top             =   720
         Value           =   1  'Checked
         Width           =   1812
      End
      Begin VB.TextBox Txtano 
         Height          =   288
         Left            =   3720
         TabIndex        =   9
         Text            =   "2004"
         Top             =   720
         Width           =   1092
      End
      Begin VB.TextBox TxtMes 
         Height          =   288
         Left            =   1080
         TabIndex        =   7
         Top             =   720
         Width           =   1092
      End
      Begin VB.Label Label2 
         Caption         =   "ANO : "
         Height          =   252
         Left            =   3120
         TabIndex        =   8
         Top             =   720
         Width           =   1092
      End
      Begin VB.Label Label1 
         Caption         =   "MES : "
         Height          =   252
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   1092
      End
   End
   Begin VB.CommandButton imprimecmd 
      Caption         =   "Imprime"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10200
      Picture         =   "frmRelFinanceiro.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1332
   End
   Begin VB.CommandButton Lercmd 
      Caption         =   "Ler Dados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8640
      Picture         =   "frmRelFinanceiro.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1332
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   3836
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   2990
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
End
Attribute VB_Name = "frmRelFin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Filename As String
Dim rs As New Recordset


Private Sub Form_Load()

On Error GoTo trataerr

ChkZerados.Value = 0
Opt1(0).Value = True
TxtMes = Format(Date, "MM")
Txtano = Format(Date, "YYYY")
rs.CursorType = adOpenStatic

Call Lercmd_Click


Exit Sub

'On Error GoTo trataerr
trataerr:
Call TrataErro(App.title, Error, "Form_Load")

End Sub

Private Sub Form_Unload(Cancel As Integer)

Call ImprimeRelDel(Filename)

End Sub


Private Sub imprimecmd_Click()
Dim fraseaux() As String

Filename = "FIN_" & gsEst_Codigo & "_" & Txtano + TxtMes + TxtDia & "_" + Format(Now, "YYYYMMDDHHMMSS") & ".HTML"
'Filename = "Financeiro" & Format(Date, "YYYYMMDD") & ".html"

ReDim fraseaux(1)

If Opt1(0) Then
    fraseaux(0) = " Dados para os Movimentos de : " & TxtMes & "/" & Txtano
ElseIf Opt1(1) Then
    fraseaux(0) = " Dados para os Pagamentos de : " & TxtMes & "/" & Txtano
Else
    fraseaux(0) = " Dados Totais Arquivados "
End If

Call Lercmd_Click
Call ImprimeHeader(Filename, "Controle Financeiro")
Call ImprimeExtra(Filename, fraseaux)
Call Imprimegrid(Filename, Grid1)
fraseaux(0) = "Dados Totalizados por data de PAGAMENTO"
Call ImprimeExtra(Filename, fraseaux)
Call Imprimegrid(Filename, Grid2)
Call ImprimeFooter(Filename, "Impresso em : " + Format(Now, "DD/MM/YY HH:MM:SS"))
Call ImprimeRel(Filename)


End Sub

Private Sub Lercmd_Click()

Dim frase As String
Dim fraseaux As String


'QUERY ANTIGA
'fraseaux = ""
'If Opt1(0) Then
'    fraseaux = "      ra.datamovimento Like '" & TxtAno & TxtMes & "%' and"
'    fraseaux = fraseaux & "         ra.seqfile *= rc.seqfile and"
'ElseIf Opt1(1) Then
'    fraseaux = "      rc.datapagamento Like '" & TxtAno & TxtMes & "%' and"
'    fraseaux = fraseaux & "         ra.seqfile *= rc.seqfile and"
'Else
'    fraseaux = fraseaux & "         ra.seqfile *= rc.seqfile and"
'End If
'
'Set rs = Nothing
'frase = ""

'frase = frase & " select"
'frase = frase & "         isnull(convert(char(8),cast(ra.datamovimento as smalldatetime),3),'-') as DtMovimento,"
'frase = frase & "         isnull(ra.seqfile,'-') as SeqArquivo,"
'frase = frase & "         isnull(str(cast(ra.reginformado as int),7),'-') as QtdeReg ,"
'frase = frase & "         isnull(str(cast(ra.totinformado as int)/100.00,10,2),'-') as ValorEnviado ,"
'frase = frase & "         isnull(str((select count(*) from tb_relfin1 as rd where " & IIf(ChkZerados, "cast(lvalor as int) > 0  and ", "") & " ra.seqfile = rd.seqfile),8),'-') as QtdeRejeitada ,"
'frase = frase & "         isnull(str(cast(rb.totencontrado as int)/100.00,10,2),'-') as ValorRejeitado,"
'frase = frase & "         cast(ra.reginformado as int) - (select count(*) from tb_relfin1 as rd where  ra.seqfile = rd.seqfile) as QtdeAceito ,"
'frase = frase & "         isnull(str(cast(rb.totinformado as int)/100.00,10,2),'Faltante') as ValorAceito ,"
'frase = frase & "         isnull(convert(char(8),cast(rc.datapagamento as smalldatetime),3),'-') as DtPGTO,"
'frase = frase & "         isnull(str(cast(rc.valorcgmp as int)/100.00,10,2),'Faltante') as ValorCGMP ,"
'frase = frase & "         isnull(str(cast(rc.valorsgmp as int)/100.00,10,2),'Faltante') as ValorSGMP"
'frase = frase & " From"
'frase = frase & "         tb_trnctl as ra ,"
'frase = frase & "         tb_trfctl as rb ,"
'frase = frase & "         tb_relfin2 As rc"
'frase = frase & " Where"
'frase = frase & fraseaux
'frase = frase & "          ra.seqfile *= rb.seqfile"
'frase = frase & " Order By"
'frase = frase & "         ra.seqfile"

'QUERY CORRIGIDA, ALTERANDO *= POR LEFT JOIN
fraseaux = ""
If Opt1(0) Then
    fraseaux = " and ra.datamovimento Like '" & Txtano & TxtMes & "%' "

ElseIf Opt1(1) Then
    fraseaux = " and rc.datapagamento Like '" & Txtano & TxtMes & "%' "

Else
    fraseaux = fraseaux
End If

Set rs = Nothing
frase = ""
frase = frase & " select"
frase = frase & "         isnull(convert(char(8),cast(ra.datamovimento as smalldatetime),3),'-') as DtMovimento,"
frase = frase & "         isnull(ra.seqfile,'-') as SeqArquivo,"
frase = frase & "         isnull(str(cast(ra.reginformado as int),7),'-') as QtdeReg ,"
frase = frase & "         isnull(str(cast(ra.totinformado as int)/100.00,10,2),'-') as ValorEnviado ,"
frase = frase & "         isnull(str((select count(*) from tb_relfin1 as rd where " & IIf(ChkZerados, "cast(lvalor as int) > 0  and ", "") & " ra.seqfile = rd.seqfile),8),'-') as QtdeRejeitada ,"
frase = frase & "         isnull(str(cast(rb.totencontrado as int)/100.00,10,2),'-') as ValorRejeitado,"
frase = frase & "         cast(ra.reginformado as int) - (select count(*) from tb_relfin1 as rd where  ra.seqfile = rd.seqfile) as QtdeAceito ,"
frase = frase & "         isnull(str(cast(rb.totinformado as int)/100.00,10,2),'Faltante') as ValorAceito, "
frase = frase & "         isnull(convert(char(8),cast(rc.datapagamento as smalldatetime),3),'-') as DtPGTO,"
frase = frase & "         isnull(str(cast(rc.valorcgmp as int)/100.00,10,2),'Faltante') as ValorCGMP ,"
frase = frase & "         isnull(str(cast(rc.valorsgmp as int)/100.00,10,2),'Faltante') as ValorSGMP"
frase = frase & " From"
frase = frase & "         tb_trnctl as ra "
frase = frase & "         left join  tb_trfctl as rb on (ra.seqfile = rb.seqfile) "
frase = frase & "         left join tb_relfin2 As rc  on (ra.seqfile = rc.seqfile)"
frase = frase & " Where 1=1 "
frase = frase & fraseaux
frase = frase & " Order By"
frase = frase & "         ra.seqfile"

Set rs = dbApp.Execute(frase)
Set Grid1.DataSource = rs
Grid1.TextMatrix(0, 0) = "Data           "
Grid1.TextMatrix(0, 1) = "Sequencial     "
Grid1.TextMatrix(0, 2) = "Env. (Qt)   "
Grid1.TextMatrix(0, 3) = "Env. (R$)      "
Grid1.TextMatrix(0, 4) = "Rej. (Qt)   "
Grid1.TextMatrix(0, 5) = "Rej. (R$)      "
Grid1.TextMatrix(0, 6) = "Aceito (Qt)   "
Grid1.TextMatrix(0, 7) = "Aceito (R$)      "
Grid1.TextMatrix(0, 8) = "Data Pgto      "
Grid1.TextMatrix(0, 9) = "CGMP (R$)        "
Grid1.TextMatrix(0, 10) = "SGMP (R$)        "
Call FormataGridx(Grid1, rs)

Grid1.Refresh

'totalizadores

'QUERY ANTIGA
'frase = ""
'frase = frase & " select"
'frase = frase & "         str(sum(cast(ra.reginformado as int)),8) as TrEnv ,"
'frase = frase & "         str(sum(cast(ra.totinformado as int))/100.00,12,2) as VlEnv ,"
'frase = frase & "         str(sum(cast(rb.regencontrado as int)-1),8) as TrRej ,"
'frase = frase & "         str(sum(cast(rb.totencontrado as int))/100.00,10,2) as VlRej,"
'frase = frase & "         str(sum(cast(rb.totinformado as int))/100.00,12,2) as VlAceito ,"
'frase = frase & "         convert(char(8),cast(rc.datapagamento as smalldatetime),3) as DtPGTO,"
'frase = frase & "         str(sum(cast(rc.valorcgmp as int))/100.00,12,2) as ValorCGMP ,"
'frase = frase & "         str(sum(cast(rc.valorsgmp as int))/100.00,12,2) as ValorSGMP"
'frase = frase & " From"
'frase = frase & "         tb_trnctl as ra ,"
'frase = frase & "         tb_trfctl as rb ,"
'frase = frase & "         tb_relfin2 As rc"
'frase = frase & " Where"
'frase = frase & fraseaux
'frase = frase & "         ra.tipo = 'TR' and"
'frase = frase & "         rb.tipo = 'RF' and"
'frase = frase & "         rb.seqfile =* ra.seqfile"
'frase = frase & " Group By rc.datapagamento "

'QUERY CORRIGIDA , ALTERANDO *= POR LEFT JOIN
frase = ""
frase = frase & " select"
frase = frase & "         str(sum(cast(ra.reginformado as int)),8) as TrEnv ,"
frase = frase & "         str(sum(cast(ra.totinformado as int))/100.00,12,2) as VlEnv ,"
frase = frase & "         str(sum(cast(rb.regencontrado as int)-1),8) as TrRej ,"
frase = frase & "         str(sum(cast(rb.totencontrado as int))/100.00,10,2) as VlRej,"
frase = frase & "         str(sum(cast(rb.totinformado as int))/100.00,12,2) as VlAceito ,"
frase = frase & "         convert(char(8),cast(rc.datapagamento as smalldatetime),3) as DtPGTO,"
frase = frase & "         str(sum(cast(rc.valorcgmp as int))/100.00,12,2) as ValorCGMP ,"
frase = frase & "         str(sum(cast(rc.valorsgmp as int))/100.00,12,2) as ValorSGMP"
frase = frase & " From"
frase = frase & "         tb_trnctl as ra "
frase = frase & "         left join tb_relfin2 As rc on (ra.seqfile = rc.seqfile)"
frase = frase & "         left join tb_trfctl as rb on (ra.seqfile=rb.seqfile)"
frase = frase & " Where 1=1 "
frase = frase & fraseaux
frase = frase & "          and ra.tipo = 'TR' and"
frase = frase & "         (rb.tipo = 'RF' or rb.tipo is null )"
frase = frase & " Group By rc.datapagamento "

Set rs = dbApp.Execute(frase)


Grid2.Clear
Set Grid2.DataSource = rs
Grid2.TextMatrix(0, 0) = "Env. (Qt)          "
Grid2.TextMatrix(0, 1) = "Env. (R$)              "
Grid2.TextMatrix(0, 2) = "Rej. (Qt)         "
Grid2.TextMatrix(0, 3) = "Rej. (R$)             "
Grid2.TextMatrix(0, 4) = "Aceito (R$)           "
Grid2.TextMatrix(0, 5) = "Data Pgto            "
Grid2.TextMatrix(0, 6) = "CGMP (R$)            "
Grid2.TextMatrix(0, 7) = "SGMP (R$)            "
Call FormataGridx(Grid2, rs)
Grid2.Refresh


Set rs = Nothing

End Sub

Private Sub saircmd_Click()

Unload Me

End Sub

