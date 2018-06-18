VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCadLN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Processa Lista Nela"
   ClientHeight    =   4125
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   13080
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   9840
      Top             =   3480
   End
   Begin VB.CommandButton cmdsair 
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
      Left            =   12000
      Picture         =   "frmCadLN.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3240
      Width           =   852
   End
   Begin VB.TextBox txtseq 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   9720
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdLimpa 
      Caption         =   "Limpar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   12240
      Picture         =   "frmCadLN.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2400
      Width           =   732
   End
   Begin VB.CommandButton cmdReindex 
      Caption         =   "Reindex"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   11400
      Picture         =   "frmCadLN.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2400
      Width           =   732
   End
   Begin VB.CommandButton cmdArquivo 
      Caption         =   "Atualizar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   10320
      Picture         =   "frmCadLN.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3240
      Width           =   852
   End
   Begin VB.Frame Frame1 
      Caption         =   "ULTIMO PROCESSADO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   9375
      Begin VB.TextBox txtULTFILE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5520
         TabIndex        =   9
         Top             =   240
         Width           =   3735
      End
      Begin VB.TextBox txtULTSEQ 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   600
         TabIndex        =   8
         Top             =   240
         Width           =   972
      End
      Begin VB.TextBox txtULTDT 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   2652
      End
      Begin VB.Label Label2 
         Caption         =   "Arq:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   17
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Data:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Seq:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   492
      End
   End
   Begin VB.TextBox txttotcad 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7560
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2640
      Width           =   2055
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1980
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   3493
      _Version        =   393216
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
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   3
      _Band(0).TextStyleHeader=   4
   End
   Begin VB.FileListBox File2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   5952
   End
   Begin VB.Label Label8 
      Caption         =   "Seq"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   9960
      TabIndex        =   16
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Total de Reg"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7920
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Falta Atualizar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   1680
      TabIndex        =   3
      Top             =   2280
      Width           =   984
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "HISTÓRICO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   300
      TabIndex        =   2
      Top             =   0
      Width           =   852
   End
End
Attribute VB_Name = "frmCadLN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsctl As New Recordset

Private Sub cmdArquivo_Click()
On Error GoTo trataerr
If cmdArquivo.Caption = "Atualizar" Then
    cmdArquivo.Caption = "Esperar"
    msglbl = ""
    Call AtualizaGrid1
    cmdArquivo.Caption = "Atualizar"
End If
Exit Sub

trataerr:
    Call TrataErro(App.title, Error, "cmdArquivo_Click")


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

Dim Ext_1 As String
Dim Ext_2 As String

If gsListas = "TAG" Then
    Ext_1 = "*.LNT"
    Ext_2 = "*.NEL"
Else
    Ext_1 = "*.NTT"
    Ext_2 = "*.NTV"
End If



File2.Path = gsPath_CGMPRecebe
File2.Pattern = Ext_1 & ";" & Ext_2
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
frase = frase & " FROM " + gsPath_DBCAD + ".dbo.TB_CADnelaCtl "
frase = frase & " WHERE cTipo = 'LN' OR cTipo = 'LT' "
frase = frase & " ORDER BY lseqfile DESC"
Set rsctl = dbApp.Execute(frase)
Set MSHFlexGrid1.DataSource = rsctl
MSHFlexGrid1.TextMatrix(0, 0) = "TIPO      "
MSHFlexGrid1.TextMatrix(0, 1) = "SEQ       "
MSHFlexGrid1.TextMatrix(0, 2) = "ATUALIZACAO           "
MSHFlexGrid1.TextMatrix(0, 3) = "ARQUIVO                    "
MSHFlexGrid1.TextMatrix(0, 4) = "FINAL         "
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
frase = frase & " FROM " + gsPath_DBCAD + ".dbo.TB_CADnelaCtl "
Set rsGeral = dbApp.Execute(frase)
txtseq = IIf(IsNull(rsGeral(0)), 0, rsGeral(0))

frase = ""
frase = frase & "SELECT top 1 "
frase = frase & " lseqfile, tsatualizacao, szarquivo "
frase = frase & " FROM " + gsPath_DBCAD + ".dbo.TB_CADnelaCtl order by lseqfile desc "
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
frase = frase & "SELECT count(*) FROM " + gsPath_DBCAD + ".dbo.TB_CADnela "
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
Call TrataErro(App.title, Me.Name, "Atualiza GRID")
cmdArquivo.Caption = "Atualizar"
cmdArquivo.Enabled = True


End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo trataerr

Exit Sub
'On Error GoTo trataerr
trataerr:
Call TrataErro(App.title, Me.Name, "Unload")

End Sub

Private Sub Timer1_Timer()
' somente para dar refhesh na tela
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
    frase = frase & " delete " + gsPath_DBCAD + ".dbo.TB_CADnelaCtl "
    frase = frase & " WHERE lseqfile <= " + Format(Val(txtseq))
    Set rsGeral = dbApp.Execute(frase)
End If
Set rsGeral = Nothing
Call AtualizaGrid1

End Sub


