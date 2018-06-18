VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCadRT 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Processa Relatório Técnico"
   ClientHeight    =   5100
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   8685
   Icon            =   "frmCadRT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   8685
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1140
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   8412
      _ExtentX        =   14843
      _ExtentY        =   2011
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
      Height          =   675
      Left            =   120
      TabIndex        =   2
      Top             =   4140
      Width           =   5832
   End
   Begin VB.CommandButton cmdArquivo 
      Caption         =   "Ler Dados"
      Height          =   855
      Left            =   6120
      Picture         =   "frmCadRT.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   7440
      Picture         =   "frmCadRT.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   1140
      Left            =   120
      TabIndex        =   6
      Top             =   1500
      Width           =   8412
      _ExtentX        =   14843
      _ExtentY        =   2011
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   3
      _Band(0).TextStyleHeader=   4
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
      Height          =   1140
      Left            =   120
      TabIndex        =   8
      Top             =   2820
      Width           =   8412
      _ExtentX        =   14843
      _ExtentY        =   2011
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   3
      _Band(0).TextStyleHeader=   4
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "FINANCEIRO"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   156
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "TÉCNICO"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   156
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   612
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "TRANSFERIDOS"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   168
      Left            =   2400
      TabIndex        =   5
      Top             =   3960
      Width           =   1116
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TRANSACAO"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   156
      Left            =   180
      TabIndex        =   4
      Top             =   0
      Width           =   876
   End
End
Attribute VB_Name = "frmCadRT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstrnctl As New Recordset
Dim rstrtctl As New Recordset
Dim rstrfctl As New Recordset

Dim rscount As New Recordset


Private Sub cmdArquivo_Click()

On Error GoTo trataerr

msglbl = ""
Call AtualizaGrid1

Exit Sub

trataerr:
    Call TrataErro(App.title, Error, "cmdArquivo_Click")
End Sub



Private Sub cmdsair_Click()

Set rsctl = Nothing
Set rscount = Nothing

Unload Me

End Sub



Private Sub Form_Load()
On Error Resume Next

Me.Top = 50
Me.Left = 50
MDI.Width = Me.Width + 500
MDI.Height = Me.Height + 1300

Set rsGeral = Nothing
rstrfctl.CursorType = adOpenStatic
File2.Path = gsPath_CGMPRecebe
File2.Pattern = "*.TRT;*.TRF;*.TRN"
File2.Refresh

Call AtualizaGrid1

End Sub

Private Sub AtualizaGrid1()

File2.Refresh
frase = ""
frase = frase & "SELECT top 100 "
frase = frase & " cast(tipo as char(2)),"
frase = frase & " cast(seqfile as char(6)),"
frase = frase & " substring(dtger,7,2)+ '/' + Substring(dtger,5,2)+ '/' + Substring(dtger,1,4)+ ' ' + Substring(hrger,1,2) + ':' + Substring(hrger,3,2)+ ':' +  Substring(hrger,5,2),"
frase = frase & " arquivo , ' ' "
frase = frase & " FROM tb_trnctl "
frase = frase & " ORDER BY Seqfile DESC, tipo "
Set rstrnctl = dbApp.Execute(frase)
Set MSHFlexGrid1.DataSource = rstrnctl
MSHFlexGrid1.TextMatrix(0, 0) = "TP          "
MSHFlexGrid1.TextMatrix(0, 1) = "SEQ           "
MSHFlexGrid1.TextMatrix(0, 2) = "DATA                "
MSHFlexGrid1.TextMatrix(0, 3) = "ARQUIVO                   "
MSHFlexGrid1.TextMatrix(0, 4) = "    "
Call FormataGridx(MSHFlexGrid1, rstrnctl)
MSHFlexGrid1.Refresh

frase = ""
frase = frase & "SELECT top 100 "
frase = frase & " cast(tipo as char(2)),"
frase = frase & " cast(seqfile as char(6)),"
frase = frase & " substring(dtger,7,2)+ '/' + Substring(dtger,5,2)+ '/' + Substring(dtger,1,4)+ ' ' + Substring(hrger,1,2) + ':' + Substring(hrger,3,2)+ ':' +  Substring(hrger,5,2),"
frase = frase & " arquivo, rejeicao"
frase = frase & " FROM tb_TRTCTL ra"
frase = frase & " ORDER BY Seqfile DESC, tipo "
Set rstrtctl = dbApp.Execute(frase)
Set MSHFlexGrid2.DataSource = rstrtctl
MSHFlexGrid2.TextMatrix(0, 0) = "TP          "
MSHFlexGrid2.TextMatrix(0, 1) = "SEQ           "
MSHFlexGrid2.TextMatrix(0, 2) = "DATA                "
MSHFlexGrid2.TextMatrix(0, 3) = "ARQUIVO                   "
MSHFlexGrid2.TextMatrix(0, 4) = "REJ "
Call FormataGridx(MSHFlexGrid2, rstrtctl)
MSHFlexGrid2.Refresh

frase = ""
frase = frase & "SELECT top 100 "
frase = frase & " cast(tipo as char(2)),"
frase = frase & " cast(seqfile as char(6)),"
frase = frase & " substring(dtger,7,2)+ '/' + Substring(dtger,5,2)+ '/' + Substring(dtger,1,4)+ ' ' + Substring(hrger,1,2) + ':' + Substring(hrger,3,2)+ ':' +  Substring(hrger,5,2),"
frase = frase & " arquivo, ' ' "
frase = frase & " FROM tb_TRFCTL "
frase = frase & " ORDER BY Seqfile DESC, tipo "
Set rstrfctl = dbApp.Execute(frase)
Set MSHFlexGrid3.DataSource = rstrfctl
MSHFlexGrid3.TextMatrix(0, 0) = "TP          "
MSHFlexGrid3.TextMatrix(0, 1) = "SEQ           "
MSHFlexGrid3.TextMatrix(0, 2) = "DATA                "
MSHFlexGrid3.TextMatrix(0, 3) = "ARQUIVO                   "
MSHFlexGrid3.TextMatrix(0, 4) = "    "
Call FormataGridx(MSHFlexGrid3, rstrfctl)
MSHFlexGrid3.Refresh

Me.Refresh

End Sub

Sub desabilitabotoes()

cmdSair.Enabled = False
cmdTecnico.Enabled = False
cmdTransacao.Enabled = False
cmdTransfer.Enabled = False

End Sub

Sub habilitabotoes()

cmdSair.Enabled = True
cmdTecnico.Enabled = True
cmdTransacao.Enabled = True
cmdTransfer.Enabled = True

End Sub

