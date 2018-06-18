VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRelTra 
   Caption         =   "Transações"
   ClientHeight    =   7404
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7404
   ScaleWidth      =   10500
   Begin VB.CommandButton cmdGerar 
      Caption         =   "Gerar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9000
      Picture         =   "frmRelTransacao.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton CmdRegerar 
      Caption         =   "Regerar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9000
      Picture         =   "frmRelTransacao.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   " DIGITE A DATA DO MOVIMENTO   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   5400
      TabIndex        =   17
      Top             =   4920
      Width           =   4812
      Begin VB.TextBox TxtDia 
         Alignment       =   1  'Right Justify
         Height          =   288
         Left            =   840
         TabIndex        =   21
         Top             =   360
         Width           =   528
      End
      Begin VB.TextBox TxtMes 
         Alignment       =   1  'Right Justify
         Height          =   288
         Left            =   2520
         TabIndex        =   20
         Top             =   360
         Width           =   528
      End
      Begin VB.TextBox TxtAno 
         Alignment       =   1  'Right Justify
         Height          =   288
         Left            =   3840
         TabIndex        =   19
         Text            =   "2004"
         Top             =   360
         Width           =   612
      End
      Begin VB.Label Label3 
         Caption         =   "DIA : "
         Height          =   252
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   492
      End
      Begin VB.Label Label1 
         Caption         =   "MES : "
         Height          =   252
         Left            =   1800
         TabIndex        =   23
         Top             =   360
         Width           =   492
      End
      Begin VB.Label Label2 
         Caption         =   "ANO : "
         Height          =   252
         Left            =   3240
         TabIndex        =   22
         Top             =   360
         Width           =   372
      End
   End
   Begin VB.CommandButton Lercmd 
      Caption         =   "Ler Dados"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5400
      Picture         =   "frmRelTransacao.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Imprimecmd 
      Caption         =   "Imprime"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7440
      Picture         =   "frmRelTransacao.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Saircmd 
      Caption         =   "&Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9120
      Picture         =   "frmRelTransacao.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resumo do Arquivo "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2052
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   4692
      Begin VB.TextBox TxtSeqFile 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2880
         TabIndex        =   12
         Top             =   600
         Width           =   1692
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1320
         TabIndex        =   11
         Top             =   1080
         Width           =   1452
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1320
         TabIndex        =   10
         Top             =   600
         Width           =   1452
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2880
         TabIndex        =   8
         Top             =   1560
         Width           =   1692
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1320
         TabIndex        =   7
         Top             =   1560
         Width           =   1452
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Sequencial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   2880
         TabIndex        =   13
         Top             =   360
         Width           =   1692
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   2880
         TabIndex        =   9
         Top             =   1200
         Width           =   1692
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   1092
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Com Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1092
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Sem Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1092
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Quantidade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   1452
      End
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   6960
      Width           =   9852
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
      Height          =   1932
      Left            =   4200
      TabIndex        =   1
      Top             =   2760
      Width           =   4092
      _ExtentX        =   7218
      _ExtentY        =   3408
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   2532
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   10092
      _ExtentX        =   17801
      _ExtentY        =   4466
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid3 
      Height          =   1932
      Left            =   120
      TabIndex        =   25
      Top             =   2760
      Width           =   3972
      _ExtentX        =   7006
      _ExtentY        =   3408
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
End
Attribute VB_Name = "frmRelTra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Filename As String

Private Sub cmdGerar_Click()
Call trataTRN
End Sub

Private Sub CmdRegerar_Click()

If Val(TxtSeqFile) > 0 Then
    Call CriaFileTRN(TxtSeqFile)
End If

End Sub

Private Sub Form_Load()
Dim frase As String
Dim rs As New Recordset

frase = "select * from tb_sceform where szform = '" + Me.Name + "'"
Set rsGeral = dbApp.Execute(frase)
If Not rsGeral.EOF And Not rsGeral.BOF Then
    Me.Top = rsGeral(1)
    Me.Left = rsGeral(2)
    Me.Width = rsGeral(3)
    Me.Height = rsGeral(4)
End If

frase = ""
frase = frase & "Select isnull(max(datamovimento),convert(char(8),getdate(),112)) from tb_reltec where tipo = 'TR'"
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

 
Call Lercmd_Click
 
End Sub

Private Sub Form_Unload(Cancel As Integer)

Call ImprimeRelDel(Filename)

frase = "delete tb_sceform where szform = '" + Me.Name + "'"
Set rsGeral = dbApp.Execute(frase)
frase = "insert tb_sceform values('" + Format(Me.Name) + "'," + Format(Me.Top) + "," + Format(Me.Left) + "," + Format(Me.Width) + "," + Format(Me.Height) + ")"
Set rsGeral = dbApp.Execute(frase)


End Sub






Private Sub Grid1_DblClick()
Grid1.Sort = 7
End Sub

Private Sub imprimecmd_Click()
Dim extra() As String

ReDim extra(3)
extra(0) = Text1
extra(1) = "Quantidade  : " + Text2
extra(2) = "Valor Total : " + Text3

Filename = "Transacao" & Format(Date, "YYYYMMDD") & ".html"
Call ImprimeHeader(Filename, "Controle de Transações")
Call ImprimeExtra(Filename, extra)
Call Imprimegrid(Filename, Grid1)
ReDim extra(1)
extra(0) = "Quantidade por Tarifas"
Call ImprimeExtra(Filename, extra)
Call Imprimegrid(Filename, Grid2)
Call ImprimeFooter(Filename, "Impresso em : " + Format(Now, "DD/MM/YY HH:MM:SS"))
Call ImprimeRel(Filename)


End Sub

Private Sub Lercmd_Click()

Dim frase As String
Dim rs As New Recordset
Dim fraseaux As String

Text1 = "Filtrado pelo Dia " + TxtDia + "/" + TxtMes + "/" + TxtAno

Lercmd.Enabled = False

If TxtDia = "00" Then
    fraseaux = " tsdatamovimento is null and "
Else
    fraseaux = " tsdatamovimento = '" & TxtAno + TxtMes + TxtDia & "' and "
End If


frase = ""
frase = frase & " select"
frase = frase & " str(count(*),10) as qtde,"
frase = frase & " str(sum(ta.lvalor)/100.00,18,2) as Valor"
frase = frase & " from tb_transacao ta"
frase = frase & " where "
frase = frase & fraseaux
frase = frase & " 1 = 1 "
Set rs = dbApp.Execute(frase)

Text2 = rs(0)
Text3 = IIf(IsNull(rs(1)), "0.00", rs(1))

frase = ""
frase = frase & " select"
frase = frase & " str(count(*),10) as qtde"
frase = frase & " from tb_transacao ta"
frase = frase & " where "
frase = frase & fraseaux
frase = frase & " 1 = 1 and lvalor = 0"
Set rs = dbApp.Execute(frase)
Text4 = rs(0)
Text5 = Text2 - Text4

frase = ""
frase = frase & " select"
frase = frase & " max(lseqfile) as Seqfile"
frase = frase & " from tb_transacao ta"
frase = frase & " where "
frase = frase & fraseaux
frase = frase & " 1 = 1"
Set rs = dbApp.Execute(frase)
TxtSeqFile = IIf(IsNull(rs(0)), "ND", rs(0))

frase = ""
frase = frase & " select"
frase = frase & " str(lvalor/100.00,10,2),"
frase = frase & " str(count(*),4),"
frase = frase & " str(sum(lvalor)/100.00,10,2)"
frase = frase & " From tb_transacao"
frase = frase & " where "
frase = frase & fraseaux
frase = frase & " 1 = 1"
frase = frase & " group by str(lvalor/100.00,10,2) with rollup"


Set rs = dbApp.Execute(frase)

Set Grid2.DataSource = rs
Call FormataGridx(Grid2, rs)
Grid2.TextMatrix(0, 0) = "Tarifa "
Grid2.TextMatrix(0, 1) = "Qtde   "
Grid2.TextMatrix(0, 2) = "Valor(R$)"


' Mostra todas as transaçoes separadas por data de saida
frase = ""
frase = frase & " select"
frase = frase & " convert(varchar(10),tssaida,103) as Saidadia,"
frase = frase & " str(count(*),4),"
frase = frase & " str(sum(ta.lvalor)/100.00,10,2)"
frase = frase & " From tb_transacao ta"
frase = frase & " where "
frase = frase & fraseaux
frase = frase & " 1 = 1"
frase = frase & " group by convert(varchar(10),tssaida,103) with rollup"


Set rs = dbApp.Execute(frase)

Set Grid3.DataSource = rs
Call FormataGridx(Grid3, rs)
Grid3.TextMatrix(0, 0) = "Data de Saida "
Grid3.TextMatrix(0, 1) = "Qtde   "
Grid3.TextMatrix(0, 2) = "Valor(R$)"

'base de dados
frase = ""
frase = frase & " select"
frase = frase & " cPlaca as Placa,"
frase = frase & " right('00000' + cast(iissuer as varchar(5)),5) + '-' + right('0000000000' + cast(ltag as varchar(10)),10) as Tag,"
frase = frase & " convert(char(16),convert(nvarchar(8),tsentrada,3)+ ' ' + convert(nvarchar(5),tsentrada,8) + '-' + "
frase = frase & " replace(replace(istentrada,'0','A'),'1','M')) as Entrada,"
frase = frase & " convert(char(16),convert(nvarchar(8),tssaida,3)+ ' ' + convert(nvarchar(5),tssaida,8) + '-' + "
frase = frase & " replace(replace(istsaida,'0','A'),'1','M')) as Saida,"
frase = frase & " str(cast(lvalor/100 as dec(10,2)),12,2) as Valor,"
frase = frase & " str(lseqfile,10) as Seq,"
frase = frase & " str(lseqreg,10) as Reg"
frase = frase & " from tb_transacao ta"
frase = frase & " where "
frase = frase & fraseaux
frase = frase & " 1 = 1 "
frase = frase & " order by lseqreg"
Set rs = dbApp.Execute(frase)

Grid1.Clear

Set Grid1.DataSource = rs
Call FormataGridx(Grid1, rs)
Grid1.TextMatrix(0, 0) = "Placa   "
Grid1.TextMatrix(0, 1) = "Tag     "
Grid1.TextMatrix(0, 2) = "Entrada "
Grid1.TextMatrix(0, 3) = "Saida   "
Grid1.TextMatrix(0, 4) = "Valor   "
Grid1.TextMatrix(0, 5) = "Seq     "
Grid1.TextMatrix(0, 6) = "Reg     "
Grid1.ColAlignment = flexAlignLeftCenter

Grid1.Refresh

'Lercmd.Enabled = True

End Sub

Private Sub saircmd_Click()

Unload Me

End Sub

Private Sub TxtAno_Change()
Lercmd.Enabled = True
End Sub

Private Sub TxtAno_LostFocus()

If Val(TxtAno) > 2010 Or Val(TxtAno) < 2000 Then TxtAno = "2006"

TxtAno = Format(TxtAno, "0000")

End Sub

Private Sub TxtDia_Change()

Lercmd.Enabled = True

End Sub

Private Sub TxtDia_LostFocus()

If Val(TxtDia) > 31 Or Val(TxtDia) < 0 Then TxtDia = "00"

TxtDia = Format(TxtDia, "00")


End Sub

Private Sub TxtMes_Change()
Lercmd.Enabled = True
End Sub

Private Sub TxtMes_LostFocus()

If Val(TxtMes) > 12 Or Val(TxtMes) < 1 Then TxtMes = "1"

TxtMes = Format(TxtMes, "00")

End Sub
