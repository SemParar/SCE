VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCadEst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Processa Cadastro EST"
   ClientHeight    =   7515
   ClientLeft      =   4455
   ClientTop       =   435
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleMode       =   0  'User
   ScaleWidth      =   36972.55
   Begin VB.FileListBox File2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   14
      Top             =   4800
      Width           =   6795
   End
   Begin VB.TextBox txttotcad 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7320
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "ULTIMO PROCESSADO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   6480
      Width           =   9375
      Begin VB.TextBox txtULTDT 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtULTSEQ 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   240
         Width           =   972
      End
      Begin VB.TextBox txtULTFILE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   6
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label6 
         Caption         =   "Seq:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Data:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Arq:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   12120
      Picture         =   "frmCadEst.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6480
      Width           =   852
   End
   Begin VB.CommandButton cmdArquivo 
      Caption         =   "&Atualizar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   9840
      Picture         =   "frmCadEst.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   2175
   End
   Begin VB.CommandButton cmdReindex 
      Caption         =   "Reindex"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12000
      Picture         =   "frmCadEst.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdLimpa 
      Caption         =   "Limpar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10920
      Picture         =   "frmCadEst.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox txtseq 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   5880
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   7320
      Top             =   5160
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3780
      Left            =   240
      TabIndex        =   13
      Top             =   360
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   6668
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   3
      _Band(0).TextStyleHeader=   4
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "HISTÓRICO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   120
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "FALTA ATUALIZAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   4560
      Width           =   1770
   End
   Begin VB.Label Label4 
      Caption         =   "TOTAL DE REGISTROS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7320
      TabIndex        =   16
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "SEQ"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   9840
      TabIndex        =   15
      Top             =   5640
      Width           =   375
   End
End
Attribute VB_Name = "frmCadEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsctl As New Recordset

Private Sub cmdArquivo_Click()
On Error GoTo trataerr

If cmdArquivo.Caption = "Atualizar" Then
    cmdArquivo.Caption = "ESPERA"
    msglbl = ""
    Call AtualizaGrid1
    cmdArquivo.Caption = "Atualizar"
End If

Exit Sub

trataerr:
    Call TrataErro(app.title, Error, "cmdArquivo_Click")

End Sub

Private Sub cmdLimpa_Click()
Dim frase

If Val(txtseq) > 0 Then
    frase = ""
    frase = frase & " delete tb_cadestctl "
    frase = frase & " WHERE lseqfile <= " + Format(Val(txtseq))
    Set rsGeral = dbApp.Execute(frase)
End If
Set rsGeral = Nothing
Call AtualizaGrid1
End Sub

Private Sub cmdReindex_Click()
cmdReindex.Enabled = False
frase = gsPath_DBCAD & ".dbo.pr_Create_IND_EST"
Set rsGeral = dbApp.Execute(frase)
cmdReindex.Enabled = True
End Sub

Private Sub cmdSair_Click()
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
End If

Me.Top = 50
Me.Left = 50
MDI.Height = Me.Height + 1200
MDI.Width = Me.Width + 450
Set rsGeral = Nothing

File2.Path = gsPath_CGMPRecebe
Dim ext_1 As String
Dim ext_2 As String

File2.Pattern = "*.ETT;*.EST"
File2.Refresh
cmdArquivo.Caption = "Atualizar"
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

frase = ""
frase = frase & "SELECT top 100"
frase = frase & " ctipo as TIPO,"
frase = frase & " lseqfile as SEQ,"
frase = frase & " tsatualizacao as ""DATA ATU"","
frase = frase & " szarquivo as ARQUIVO,"
frase = frase & " ltotal as ""REG FINAL "","
frase = frase & " lregistros as REG,"
frase = frase & " lremo as REMO,"
frase = frase & " lincl as INCL,"
frase = frase & " lalte as ALTE"
frase = frase & " FROM TB_CADESTCTL "
frase = frase & " WHERE cTipo='TG' OR cTipo='TT'"
frase = frase & " ORDER BY lSeqfile DESC"
Set rsctl = dbApp.Execute(frase)
Set MSHFlexGrid1.DataSource = rsctl
MSHFlexGrid1.TextMatrix(0, 0) = "TIPO    "
MSHFlexGrid1.TextMatrix(0, 1) = "SEQ      "
MSHFlexGrid1.TextMatrix(0, 2) = "ATUALIZACAO            "
MSHFlexGrid1.TextMatrix(0, 3) = "ARQUIVO                     "
MSHFlexGrid1.TextMatrix(0, 4) = "FINAL         "
MSHFlexGrid1.TextMatrix(0, 5) = "REG      "
MSHFlexGrid1.TextMatrix(0, 6) = "REMO    "
MSHFlexGrid1.TextMatrix(0, 7) = "INCL    "
MSHFlexGrid1.TextMatrix(0, 8) = "ALTE    "
MSHFlexGrid1.ColAlignment = flexRightLeftBottom
Call FormataGridx(MSHFlexGrid1, rsctl)
MSHFlexGrid1.Refresh

frase = ""
frase = frase & "SELECT top 1 "
frase = frase & " min(lseqfile) "
frase = frase & " FROM TB_CADESTCTL "
Set rsGeral = dbApp.Execute(frase)
txtseq = IIf(IsNull(rsGeral(0)), 0, rsGeral(0))

frase = ""
frase = frase & "SELECT top 1 "
frase = frase & " lseqfile, tsatualizacao, szarquivo "
frase = frase & " FROM TB_CADESTCTL order by lseqfile desc "
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
frase = frase & "SELECT count(*) FROM " & gsPath_DB & ".dbo.tb_cadtag"
Set rsGeral = dbApp.Execute(frase)
txttotcad = IIf(IsNull(rsGeral(0)), 0, rsGeral(0))
Set rsGeral = Nothing

cmdArquivo.Caption = "Atualizar"
cmdArquivo.Enabled = True

File2.Refresh
Me.Refresh

Exit Sub
'On Error GoTo trataerr
trataerr:
Call TrataErro(app.title, Error, "Atualiza GRID")
cmdArquivo.Caption = "Atualizar"
cmdArquivo.Enabled = True


End Sub

Private Sub Timer1_Timer()

Timer1.Interval = 60000

Call cmdArquivo_Click

End Sub

