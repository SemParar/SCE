VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRelPat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contrôle do Pátio"
   ClientHeight    =   6630
   ClientLeft      =   90
   ClientTop       =   480
   ClientWidth     =   12075
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRelPateo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   12075
   Begin VB.Frame Frame7 
      Height          =   732
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   11772
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Posição do Dia:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   3000
         TabIndex        =   25
         Top             =   240
         Width           =   2520
      End
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
      Height          =   855
      Left            =   10680
      Picture         =   "frmRelPateo.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   " Pátio    "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5652
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   7455
      Begin VB.Frame Frame4 
         Caption         =   "Pesquisar Tag"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   972
         Left            =   120
         TabIndex        =   13
         Top             =   4560
         Width           =   7212
         Begin VB.TextBox Text9 
            Alignment       =   1  'Right Justify
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
            Left            =   1320
            TabIndex        =   27
            Text            =   "Text9"
            Top             =   360
            Width           =   1692
         End
         Begin VB.TextBox Text8 
            Alignment       =   1  'Right Justify
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
            Left            =   4080
            TabIndex        =   26
            Text            =   "Text8"
            Top             =   360
            Width           =   1572
         End
         Begin VB.CommandButton cmdpesqTag 
            Caption         =   "Pesquisar"
            Height          =   396
            Left            =   5880
            TabIndex        =   22
            Top             =   360
            Width           =   1212
         End
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
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
            Left            =   3240
            TabIndex        =   14
            Text            =   "Text7"
            Top             =   360
            Width           =   732
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Tag:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   396
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   612
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
         Height          =   3975
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   7011
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
            Size            =   12
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
   Begin VB.Frame Frame2 
      Caption         =   " Transações    "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4332
      Left            =   7680
      TabIndex        =   2
      Top             =   840
      Width           =   4215
      Begin VB.Frame Frame6 
         Caption         =   "Saldo no Patio"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   972
         Left            =   120
         TabIndex        =   19
         Top             =   3120
         Width           =   3975
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   444
            Left            =   1920
            TabIndex        =   20
            Text            =   "Text6"
            Top             =   360
            Width           =   1692
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Total: "
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   600
            TabIndex        =   21
            Top             =   360
            Width           =   1332
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Quantidades Entradas"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   852
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   3975
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1920
            TabIndex        =   17
            Text            =   "Text2"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   396
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   1692
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Quantidades Saidas"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1692
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   3975
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   396
            Left            =   1920
            TabIndex        =   8
            Text            =   "Text4"
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1920
            TabIndex        =   7
            Text            =   "Text5"
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   396
            Left            =   1920
            TabIndex        =   6
            Text            =   "Text3"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   396
            Left            =   120
            TabIndex        =   11
            Top             =   1200
            Width           =   1692
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Com Valor:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   396
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   1692
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Sem Valor:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   396
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   1692
         End
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
      Height          =   855
      Left            =   9360
      Picture         =   "frmRelPateo.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   1095
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
      Height          =   855
      Left            =   7920
      Picture         =   "frmRelPateo.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5520
      Width           =   1095
   End
End
Attribute VB_Name = "frmRelPat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Filename As String



Private Sub cmdpesqTag_Click()

frmTagPes.Show
frmTagPes.cmdLimpa_Click
frmTagPes.TxtEmissor = Text7
frmTagPes.TxtTag = Text8
frmTagPes.TxtPlaca = ""
frmTagPes.cmdPes_Click

'Call Lercmd_Click

End Sub

Private Sub cmdsair_Click()
Unload Me
End Sub


Private Sub Form_Load()

Me.Top = 10
Me.Left = 10

Call Lercmd_Click
 
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Call ImprimeRelDel(Filename)

End Sub

Private Sub Grid1_DblClick()

Grid1.Sort = 7

End Sub

Private Sub Grid1_Click()

Grid1.ColSel = Grid1.Col
Grid1.RowSel = Grid1.Row
Grid1.BackColorSel = vbBlue
Text7 = UCase(LTrim(RTrim(Grid1.TextMatrix(Grid1.Row, 1))))
Text8 = UCase(LTrim(RTrim(Grid1.TextMatrix(Grid1.Row, 2))))
Text9 = UCase(LTrim(RTrim(Grid1.TextMatrix(Grid1.Row, 0))))

End Sub



Private Sub imprimecmd_Click()
Dim extra() As String

ReDim extra(4)
extra(0) = "DATA : " + Format(Now, "DD/MM/YY hh:mm")
extra(1) = "Quantidade : " + Text6

Grid1.Col = 0
Grid1.Sort = 7

Filename = "PAT_" & gsEst_Codigo & "_" & Format(Now, "YYYYMMDDHHMMSS") & ".HTML"
'Filename = "Veiculos" & Format(Date, "YYYYMMDD") & ".html"
Call ImprimeHeader(Filename, "Veículos no Páteo")
Call ImprimeExtra(Filename, extra)
Call Imprimegrid(Filename, Grid1)
Call ImprimeFooter(Filename, "Impresso em : " + Format(Now, "DD/MM/YY HH:MM:SS"))
Call ImprimeRel(Filename)

End Sub

Private Sub Lercmd_Click()
' relatorio do patio.
Dim frase As String
Dim rs As New Recordset
Dim aux As String
Dim fraseaux As String
Dim extra As String
rs.CursorType = adOpenStatic

Text1 = Format(CVDate(Now), "DD/MM/YYYY")

frase = ""
frase = frase + " select"
frase = frase + " count(*) as MovDia,"
frase = frase + " count(case when Ivalor = 0 then 1 end) as MovDiaZerado,"
frase = frase + " count(case when Ivalor > 0 then 1 end) as MovDiaValor"
frase = frase + " From"
frase = frase + " tb_transacao"
frase = frase + " Where"
frase = frase + " lseqfile is null"
Set rs = dbApp.Execute(frase)


Text2 = IIf(IsNull(rs("MovDia")), 0, rs("MovDia"))
Text3 = IIf(IsNull(rs("MovDiaZerado")), 0, rs("MovDiaZerado"))
Text4 = IIf(IsNull(rs("MovDiaValor")), 0, rs("MovDiaValor"))
Text5 = Val(Text3) + Val(Text4)

frase = ""
frase = frase & "select count(*) "
frase = frase & " from tb_praia "
frase = frase & " where tsegresso <= '20040101'"
Set rs = dbApp.Execute(frase)
Text6 = rs(0)
Text2 = Val(Text2) + Val(Text6)

Text7 = ""
Text8 = ""
Text9 = ""

frase = ""
frase = frase & "select ' ' + Cplaca as PLACA, ' ' + iissuer as EMISSOR, ' ' + ltag as TAG,"
frase = frase & " ' ' + (convert(char(11),tsingresso,103)+ convert(char(5),tsingresso,108))as ""DATA ENTRADA"", "
'frase = frase & " (select SUBSTRING(cdescricao,1,8) from tb_pista where ipista = ientpista)as Ent"
frase = frase & " SUBSTRING(p.cdescricao,1,8)"
'frase = frase & " from tb_praia t , tb_pista P"
frase = frase & " from tb_praia t left join tb_pista P"
frase = frase & " on p.ipista = t.ientpista"
frase = frase & " where tsegresso <= '20040101'"
'frase = frase & " and p.ipista =* t.ientpista"
frase = frase & " order by cplaca"
Set rs = dbApp.Execute(frase)


Grid1.Clear
Set Grid1.DataSource = rs
Call FormataGridx(Grid1, rs)

Grid1.TextMatrix(0, 0) = "Placa"
Grid1.TextMatrix(0, 1) = "Cod"
Grid1.TextMatrix(0, 2) = "Tag"
Grid1.TextMatrix(0, 3) = "Data de Entrada"
Grid1.TextMatrix(0, 4) = "Entrada"
Grid1.ColAlignment = flexAlignLeftCenter
Grid1.ColWidth(0) = 1400
Grid1.ColWidth(1) = 800
Grid1.ColWidth(2) = 1200
Grid1.ColWidth(3) = 2300
Grid1.ColWidth(4) = 1200
Grid1.AllowBigSelection = False
Grid1.Refresh

End Sub

Private Sub saircmd_Click()

Unload Me

End Sub

