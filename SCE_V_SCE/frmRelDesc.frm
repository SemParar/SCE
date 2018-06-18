VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRelDesc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório do Movimento"
   ClientHeight    =   6015
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   13065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   13065
   Begin VB.Frame Frame5 
      Caption         =   "Ticket Nao Usados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3852
      Left            =   120
      TabIndex        =   23
      Top             =   2040
      Width           =   6252
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
         Height          =   3252
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   5772
         _ExtentX        =   10186
         _ExtentY        =   5741
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
   Begin VB.Frame Frame3 
      Caption         =   "Ticket Usados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3852
      Left            =   6480
      TabIndex        =   21
      Top             =   2040
      Width           =   6492
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdFiles 
         Height          =   3252
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   6012
         _ExtentX        =   10610
         _ExtentY        =   5741
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
   Begin VB.CommandButton CmdRegerar 
      Caption         =   "TICKET"
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
      Left            =   9480
      Picture         =   "frmRelDesc.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   " Data da Pesquisa   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   6252
      Begin VB.TextBox TxtDia 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   408
         Left            =   1080
         TabIndex        =   16
         Top             =   360
         Width           =   768
      End
      Begin VB.TextBox TxtMes 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   408
         Left            =   3120
         TabIndex        =   15
         Top             =   360
         Width           =   768
      End
      Begin VB.TextBox TxtAno 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   408
         Left            =   4800
         TabIndex        =   14
         Text            =   "2004"
         Top             =   360
         Width           =   1332
      End
      Begin VB.Label Label3 
         Caption         =   "DIA : "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   480
         TabIndex        =   19
         Top             =   360
         Width           =   492
      End
      Begin VB.Label Label1 
         Caption         =   "MES : "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   2400
         TabIndex        =   18
         Top             =   360
         Width           =   492
      End
      Begin VB.Label Label2 
         Caption         =   "ANO : "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   4200
         TabIndex        =   17
         Top             =   360
         Width           =   492
      End
   End
   Begin VB.CommandButton Lercmd 
      Caption         =   "Ler Dados"
      Enabled         =   0   'False
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
      Left            =   6960
      Picture         =   "frmRelDesc.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Imprimecmd 
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
      Height          =   855
      Left            =   8280
      Picture         =   "frmRelDesc.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   1095
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
      Picture         =   "frmRelDesc.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resumo do Ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   12852
      Begin VB.TextBox txtVal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3240
         TabIndex        =   9
         Top             =   360
         Width           =   1452
      End
      Begin VB.TextBox TxtZero 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7080
         TabIndex        =   8
         Top             =   360
         Width           =   1452
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1320
         TabIndex        =   6
         Top             =   2520
         Visible         =   0   'False
         Width           =   1452
      End
      Begin VB.TextBox txtTot 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   10920
         TabIndex        =   5
         Top             =   360
         Width           =   1452
      End
      Begin VB.Line Line1 
         Visible         =   0   'False
         X1              =   240
         X2              =   2760
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   360
         TabIndex        =   7
         Top             =   2640
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Emitidos: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   9720
         TabIndex        =   4
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Nao Usados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   1920
         TabIndex        =   3
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Usados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   6000
         TabIndex        =   2
         Top             =   480
         Width           =   732
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Quantidade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1092
      End
   End
End
Attribute VB_Name = "frmRelDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Filename As String

Private Sub cmdGerar_Click()
Dim aux As Integer

aux = MyMsgBox("Confirma a Geração de 60 Tickets", vbOKCancel, "GERAR TICKET DESCONTO", "Tem Certeza da Operação")

If aux = 1 Then
    frase = ""
    frase = frase & "select max(lticket) "
    frase = frase & "       from tb_ticket"
    Set rs = dbApp.Execute(frase)
    aux = rs(0)
        
    For i = aux + 1 To aux + 61
        frase = ""
        frase = frase & "insert tb_ticket  "
        frase = frase & " (dvalor) values ('700') "
        Set rs = dbApp.Execute(frase)
    Next
End If

End Sub

Private Sub CmdRegerar_Click()

Dim aux As Integer

frmIMPDesc.Show

End Sub



Private Sub Form_Load()
Dim frase As String
Dim rs As New Recordset
rs.CursorType = adOpenStatic

If gbytNivel_Acesso_Usuario <> gintNIVEL_ADMINISTRADOR Then
    MsgBoxService "Somente Administrator Tem Acesso ao Recurso"
End If

Me.Top = 10
Me.Left = 10

TxtDia = Format(Date, "dd")
TxtMes = Format(Date, "MM")
TxtAno = Format(Date, "YYYY")

Call Lercmd_Click
 
End Sub

Private Sub Form_Unload(Cancel As Integer)

Call ImprimeRelDel(Filename)

End Sub

Private Sub GrdFiles_Click()

GrdFiles.ColSel = GrdFiles.Col
GrdFiles.RowSel = GrdFiles.Row
GrdFiles.BackColorSel = 1
txtseqfile = UCase(LTrim(RTrim(GrdFiles.TextMatrix(GrdFiles.Row, 0))))

End Sub

Private Sub Grid1_DblClick()
Grid1.Sort = 7
End Sub

Private Sub imprimecmd_Click()
Dim extra() As String

ReDim extra(3)
'extra(0) = Text1
'extra(1) = "Quantidade  : " + txtTot + "   -   ( Zeradas: " + TxtZero + " - c/ Valor: " + txtVal + " )"
'extra(2) = "Valor Total : " + txtValor

Filename = "FENAC_" & gsEst_Codigo & "_" & Format(Now, "YYYYMMDDHHMMSS") & ".HTML"

'Filename = "TicketFenac" & Format(Date, "YYYYMMDD") & ".html"
Call ImprimeHeader(Filename, "Controle de Tickets")
Call ImprimeExtra(Filename, extra)
ReDim extra(1)
'extra(0) = "Quantidade por Tarifas"
'Call ImprimeExtra(Filename, extra)
'Call Imprimegrid(Filename, Grid2)
extra(0) = "Ticket Usadados"
Call ImprimeExtra(Filename, extra)
Call Imprimegrid(Filename, GrdFiles)
extra(0) = "Ticket Nao Usados"
Call ImprimeExtra(Filename, extra)
Call Imprimegrid(Filename, Grid2)



Call ImprimeFooter(Filename, "Impresso em : " + Format(Now, "DD/MM/YY HH:MM:SS"))
Call ImprimeRel(Filename)


End Sub

Private Sub Lercmd_Click()

Dim frase As String
Dim rs As New Recordset
Dim fraseaux As String
Dim aux As Boolean
rs.CursorType = adOpenStatic

Text1 = "Filtrado pelo Dia " + TxtDia + "/" + TxtMes + "/" + TxtAno

Lercmd.Enabled = False

fraseaux = ""

frase = ""
frase = frase & "select count(*) "
frase = frase & "       from tb_ticket"
frase = frase & "       where  "
frase = frase & fraseaux
frase = frase & "       cplaca is null and "
frase = frase & "       1 = 1 "
Set rs = dbApp.Execute(frase)
txtVal = rs(0)
'If rs(0) <> 0 Then
'    txtValor = FormatCurrency(rs(1))
'Else
'    txtValor = 0
'End If

frase = ""
frase = frase & "select count(*) "
frase = frase & "       from tb_ticket"
frase = frase & "       where  "
frase = frase & fraseaux
frase = frase & "       cplaca is not null and "
frase = frase & "       1 = 1 "
Set rs = dbApp.Execute(frase)
TxtZero = rs(0)
txtTot = Val(TxtZero) + Val(txtVal)

'grid USADO

frase = ""
frase = frase & " select "
frase = frase & " convert(varchar(15),lticket) as Ticket,"
frase = frase & " convert(varchar(10),Cplaca) as Placa,"
frase = frase & " convert(varchar(10),dvalor) as Valor,"
frase = frase & " convert(varchar(20),tsdatauso,103)+ ' ' + convert(varchar(30),tsdatauso,108) as Usado,"
frase = frase & " convert(varchar(20),tsdatabaixado,103)+ ' ' + convert(varchar(30),tsdatabaixado,108) as Baixado"
frase = frase & " from tb_Ticket"
frase = frase & " where"
frase = frase & " cplaca is not null "
frase = frase & " order by lticket"
Set rs = dbApp.Execute(frase)
Set GrdFiles.DataSource = rs
GrdFiles.TextMatrix(0, 0) = "Ticket "
GrdFiles.TextMatrix(0, 1) = "Placa    "
GrdFiles.TextMatrix(0, 2) = "Valor   "
GrdFiles.TextMatrix(0, 3) = "Validado          "
GrdFiles.TextMatrix(0, 4) = "Baixado          "
GrdFiles.ColAlignment = flexAlignLeftCenter
Call FormataGridx(GrdFiles, rs)


frase = ""
frase = frase & " select "
frase = frase & " convert(varchar(15),lticket) as Ticket,"
frase = frase & " convert(varchar(10),Cplaca) as Placa,"
frase = frase & " convert(varchar(10),dvalor) as Valor,"
frase = frase & " convert(varchar(20),tsdatauso,103)+ ' ' + convert(varchar(30),tsdatauso,108) as Usado,"
frase = frase & " convert(varchar(20),tsdatabaixado,103)+ ' ' + convert(varchar(30),tsdatabaixado,108) as Baixado"
frase = frase & " from tb_Ticket"
frase = frase & " where"
frase = frase & " cplaca is null "
frase = frase & " order by lticket"
Set rs = dbApp.Execute(frase)
Set Grid2.DataSource = rs
Grid2.TextMatrix(0, 0) = "Ticket "
Grid2.TextMatrix(0, 1) = "Placa    "
Grid2.TextMatrix(0, 2) = "Valor   "
Grid2.TextMatrix(0, 3) = "Validado          "
Grid2.TextMatrix(0, 4) = "Baixado           "
Grid2.ColAlignment = flexAlignLeftCenter
Call FormataGridx(Grid2, rs)

Lercmd.Enabled = True

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
