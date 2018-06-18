VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCadLN 
   Caption         =   "Processa Cadastro LN"
   ClientHeight    =   4884
   ClientLeft      =   120
   ClientTop       =   432
   ClientWidth     =   12324
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4972.341
   ScaleMode       =   0  'User
   ScaleWidth      =   12468.56
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   11280
      Top             =   2280
   End
   Begin VB.TextBox txtseq 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11160
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   360
      Width           =   972
   End
   Begin VB.CommandButton cmdLimpa 
      Caption         =   "Limpar"
      Height          =   612
      Left            =   11280
      Picture         =   "frmCadLN.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   840
      Width           =   732
   End
   Begin VB.CommandButton cmdReindex 
      Caption         =   "Reindex"
      Height          =   612
      Left            =   11280
      Picture         =   "frmCadLN.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1560
      Width           =   732
   End
   Begin VB.CommandButton cmdArquivo 
      Caption         =   "ATUALIZAR"
      Height          =   732
      Left            =   8640
      Picture         =   "frmCadLN.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2640
      Width           =   972
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   732
      Left            =   10080
      Picture         =   "frmCadLN.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2640
      Width           =   972
   End
   Begin VB.Frame Frame1 
      Caption         =   "ULTIMO PROCESSADO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   10932
      Begin VB.TextBox txtULTFILE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   7560
         TabIndex        =   9
         Top             =   240
         Width           =   3252
      End
      Begin VB.TextBox txtULTSEQ 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   600
         TabIndex        =   8
         Top             =   240
         Width           =   972
      End
      Begin VB.TextBox txtULTDT 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   3480
         TabIndex        =   7
         Top             =   240
         Width           =   2652
      End
      Begin VB.Label Label2 
         Caption         =   "Arq:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   6960
         TabIndex        =   18
         Top             =   240
         Width           =   372
      End
      Begin VB.Label Label5 
         Caption         =   "Data:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2640
         TabIndex        =   11
         Top             =   240
         Width           =   612
      End
      Begin VB.Label Label6 
         Caption         =   "Seq:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   492
      End
   End
   Begin VB.TextBox txttotcad 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6480
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3000
      Width           =   1452
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1980
      Left            =   180
      TabIndex        =   0
      Top             =   300
      Width           =   10812
      _ExtentX        =   19071
      _ExtentY        =   3493
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   3
      _Band(0).TextStyleHeader=   4
   End
   Begin VB.FileListBox File2 
      Height          =   1032
      Left            =   119
      TabIndex        =   1
      Top             =   2593
      Width           =   5883
   End
   Begin VB.Label Label8 
      Caption         =   "Sequencial"
      Height          =   252
      Left            =   11160
      TabIndex        =   17
      Top             =   120
      Width           =   972
   End
   Begin VB.Label Label4 
      Caption         =   "Total de Reg no Cadastro"
      Height          =   252
      Left            =   6480
      TabIndex        =   4
      Top             =   2760
      Width           =   1812
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Falta Atualizar"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   156
      Left            =   1680
      TabIndex        =   3
      Top             =   2400
      Width           =   888
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "HISTÓRICO"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   165
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   765
   End
End
Attribute VB_Name = "frmCadLN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsctl As New Recordset

Private Sub cmdArquivo_Click()

If cmdArquivo.Caption = "ATUALIZAR" Then
    msglbl = ""
    Call AtualizaGrid1
End If

End Sub

Private Sub cmdsair_Click()

Unload Me

End Sub

Private Sub Form_Load()
On Error Resume Next

If gbytNivel_Acesso_Usuario = 1 Then
    cmdLimpa.Visible = True
    cmdReindex.Visible = True
    txtseq.Visible = True
    Label8.Visible = True
Else
    cmdLimpa.Visible = False
    cmdReindex.Visible = False
    txtseq.Visible = False
    Label8.Visible = False
    Me.BorderStyle = 1
End If

Me.Top = 50
Me.Left = 50
MDI.Width = Me.Width + 500
MDI.Height = Me.Height + 1300

Set rsGeral = Nothing

File2.Path = gsPath_CGMPRecebe
File2.Pattern = "*.NEL;*.LNT"
File2.Refresh
cmdArquivo.Caption = "ATUALIZAR"
cmdArquivo.Enabled = True

Call AtualizaGrid1

End Sub

Private Sub AtualizaGrid1()
On Error GoTo trataerr

cmdArquivo.Caption = "ESPERE"
cmdArquivo.Enabled = False
MSHFlexGrid1.Clear
Set MSHFlexGrid1.DataSource = Nothing
Set rsctl = Nothing

File2.Refresh
frase = ""
frase = frase & "SELECT top 100 "
frase = frase & " ctipo as TIPO,"
frase = frase & " lseqfile as SEQ,"
frase = frase & " tsatualizacao as ""DATA ATU"","
frase = frase & " szarquivo as ARQUIVO,"
frase = frase & " ltotal as ""REG FINAL "","
frase = frase & " lregistros as REG,"
frase = frase & " lremo as REMO,"
frase = frase & " lincl as INCL,"
frase = frase & " lalte as ALTE"
frase = frase & " FROM tb_cadnelaCtl "
frase = frase & " WHERE cTipo = 'LN' OR cTipo = 'LT' "
frase = frase & " ORDER BY lseqfile DESC"
Set rsctl = dbApp.Execute(frase)
Set MSHFlexGrid1.DataSource = rsctl
MSHFlexGrid1.TextMatrix(0, 0) = "TIPO   "
MSHFlexGrid1.TextMatrix(0, 1) = "SEQ    "
MSHFlexGrid1.TextMatrix(0, 2) = "ATUALIZACAO           "
MSHFlexGrid1.TextMatrix(0, 3) = "ARQUIVO               "
MSHFlexGrid1.TextMatrix(0, 4) = "FINAL       "
MSHFlexGrid1.TextMatrix(0, 5) = "REG     "
MSHFlexGrid1.TextMatrix(0, 6) = "REMO    "
MSHFlexGrid1.TextMatrix(0, 7) = "INCL    "
MSHFlexGrid1.TextMatrix(0, 8) = "ALTE    "
MSHFlexGrid1.ColAlignment = flexRightLeftBottom
Call FormataGridx(MSHFlexGrid1, rsctl)
MSHFlexGrid1.Refresh

frase = ""
frase = frase & "SELECT top 1 "
frase = frase & " min(lseqfile) "
frase = frase & " FROM tb_CadnelaCtl "
Set rsGeral = dbApp.Execute(frase)
txtseq = IIf(IsNull(rsGeral(0)), 0, rsGeral(0))

frase = ""
frase = frase & "SELECT top 1 "
frase = frase & " lseqfile, tsatualizacao, szarquivo "
frase = frase & " FROM tb_CadnelaCtl order by lseqfile desc "
Set rsGeral = dbApp.Execute(frase)
If Not (rsGeral.EOF And rsGeral.BOF) Then
    txtULTSEQ = IIf(IsNull(rsGeral(0)), 0, rsGeral(0))
    txtULTDT = IIf(IsNull(rsGeral(1)), 0, rsGeral(1))
    txtULTFILE = IIf(IsNull(rsGeral(2)), 0, rsGeral(2))
Else
    txtULTSEQ = 0
    txtULTDT = 0
    txtULTFILE = 0
End If

frase = ""
frase = frase & "SELECT count(*) FROM tb_Cadnela "
Set rsGeral = dbApp.Execute(frase)
txttotcad = IIf(IsNull(rsGeral(0)), 0, rsGeral(0))
Set rsGeral = Nothing

cmdArquivo.Caption = "ATUALIZAR"
cmdArquivo.Enabled = True
Me.Refresh

Exit Sub
'On Error GoTo trataerr
trataerr:
Call TrataErro(App.Title, Me.Name, "Atualiza GRID")
cmdArquivo.Caption = "ATUALIZAR"
cmdArquivo.Enabled = True


End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo trataerr

If gbytNivel_Acesso_Usuario = 1 Then
    frase = "delete " + gsPath_DB + ".dbo.tb_sceform where szform = '" + Me.Name + "'"
    Set rsGeral = dbApp.Execute(frase)
    frase = "insert " + gsPath_DB + ".dbo.tb_sceform values('" + Format(Me.Name) + "'," + Format(Me.Top) + "," + Format(Me.Left) + "," + Format(Me.Width) + "," + Format(Me.Height) + ")"
    Set rsGeral = dbApp.Execute(frase)
    Set rsGeral = Nothing
End If

Exit Sub
'On Error GoTo trataerr
trataerr:
Call TrataErro(App.Title, Me.Name, "Unload")

End Sub

Private Sub Timer1_Timer()

Timer1.Interval = 60000
Call cmdArquivo_Click

End Sub

Private Sub cmdreindex_Click()

frase = "pr_Create_IND_NELA"
Set rsGeral = dbApp.Execute(frase)

End Sub

Private Sub cmdLimpa_Click()
Dim frase

If Val(txtseq) > 0 Then
    frase = ""
    frase = frase & " delete tb_CadnelaCtl "
    frase = frase & " WHERE lseqfile <= " + Format(Val(txtseq))
    Set rsGeral = dbApp.Execute(frase)
End If
Set rsGeral = Nothing
Call AtualizaGrid1

End Sub

