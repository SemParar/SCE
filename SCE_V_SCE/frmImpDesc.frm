VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmIMPDesc 
   Caption         =   "Imprime Ticket"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   8550
   Begin VB.TextBox txtfim 
      Height          =   372
      Left            =   4440
      TabIndex        =   17
      Text            =   "TXTfim"
      Top             =   5400
      Width           =   1212
   End
   Begin VB.TextBox txtGerPag 
      Height          =   372
      Left            =   1560
      TabIndex        =   16
      Top             =   5400
      Width           =   1212
   End
   Begin VB.TextBox TXTINI 
      Height          =   372
      Left            =   3120
      TabIndex        =   13
      Text            =   "TXTINI"
      Top             =   5400
      Width           =   1212
   End
   Begin VB.CommandButton cmdGerar 
      Caption         =   "Gerar"
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
      Left            =   240
      Picture         =   "frmImpDesc.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Frame Frame5 
      Caption         =   "Ticket Nao Impressos"
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
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   6972
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
         Height          =   3252
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   6612
         _ExtentX        =   11668
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
   Begin VB.Frame Frame2 
      Caption         =   " Data da Impressao"
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
      TabIndex        =   3
      Top             =   120
      Width           =   6972
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
         Left            =   1200
         TabIndex        =   6
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
         Left            =   3240
         TabIndex        =   5
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
         Left            =   4920
         TabIndex        =   4
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
         Left            =   600
         TabIndex        =   9
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
         Left            =   2520
         TabIndex        =   8
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
         Left            =   4320
         TabIndex        =   7
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
      Left            =   7320
      Picture         =   "frmImpDesc.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
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
      Left            =   6120
      Picture         =   "frmImpDesc.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
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
      Left            =   7320
      Picture         =   "frmImpDesc.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Paginas"
      Height          =   252
      Left            =   1680
      TabIndex        =   15
      Top             =   5160
      Width           =   972
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "INICIO"
      Height          =   252
      Left            =   3240
      TabIndex        =   14
      Top             =   5160
      Width           =   972
   End
End
Attribute VB_Name = "frmIMPDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGerar_Click()
Dim aux, aux1 As Integer
Dim vdesc As String

aux = MyMsgBox("Confirma a Geração de Tickets " + txtGerPag + " Pagina(s)", vbOKCancel, "GERAR TICKET DESCONTO", "Tem Certeza da Operação")


If aux = 1 Then
    frase = ""
    frase = frase & "select max(lticket) "
    frase = frase & "       from tb_ticket "
    Set rs = dbApp.Execute(frase)
    aux = IIf(IsNull(rs(0)), 0, rs(0))
        
    frase = ""
    frase = frase & " Select cvalor from tb_parametros "
    frase = frase & " where cdescr = 'DESCFENAC' "
    Set rs = dbApp.Execute(frase)
    If rs.EOF And rs.BOF Then
        MsgBoxService "Valor de Desconto não Cadastrado"
        Exit Sub
    Else
        aux1 = MyMsgBox("Confirma a Geração de " + Format(Val(txtGerPag)) + " Paginas Tickets", vbOKCancel, "GERAR TICKET DESCONTO", "Tem Certeza da Operação")
        If aux1 <> 1 Then
            Exit Sub
        Else
            vdesc = rs(0)
        End If
    End If
    
    For i = 1 To Val(txtfim) - Val(TXTINI) + 1
        frase = ""
        frase = frase & "insert tb_ticket  "
        frase = frase & " (dvalor) values ('" + Format(Val(vdesc)) + "')"
        Set rs = dbApp.Execute(frase)
    Next
End If

Call Lercmd_Click

End Sub

Private Sub Form_Load()
Dim frase As String
Dim rs As New Recordset

rs.CursorType = adOpenStatic


Me.Top = 10
Me.Left = 10


TxtDia = Format(Date, "dd")
TxtMes = Format(Date, "MM")
TxtAno = Format(Date, "YYYY")
TXTINI.Enabled = False
txtfim.Enabled = False

Call Lercmd_Click
 
End Sub

Private Sub Grid1_DblClick()
Grid1.Sort = 7
End Sub

Private Sub imprimecmd_Click()
'http://192.168.100.103/sceweb/aspx/ticket/default.aspx?vINI=xxx&vFIM=YYY
Dim aux As String

aux = "http://localhost/sceweb/aspx/ticket/default.aspx?vINI=" + Format(Val(TXTINI)) + "&vFIM=" + Format(Val(txtfim))
Dim ie As Object
Set ie = CreateObject("InternetExplorer.Application")
ie.Visible = True
ie.Navigate aux
Call Lercmd_Click

End Sub

Private Sub Lercmd_Click()

Dim frase As String
Dim rs As New Recordset
Dim fraseaux As String
Dim aux As Boolean
rs.CursorType = adOpenStatic

Lercmd.Enabled = False

frase = ""
frase = frase & " select "
frase = frase & " convert(varchar(15),lticket) as Ticket,"
frase = frase & " convert(varchar(10),Cplaca) as Placa,"
frase = frase & " convert(varchar(10),dvalor) as Valor,"
frase = frase & " convert(varchar(20),tsdatauso,103)+ ' ' + convert(varchar(30),tsdatauso,108) as Usado,"
frase = frase & " convert(varchar(20),tsdatabaixado,103)+ ' ' + convert(varchar(30),tsdatabaixado,108) as Baixado"
frase = frase & " from tb_Ticket"
frase = frase & " where"
frase = frase & " dtimp is null "
frase = frase & " order by lticket"
Set rs = dbApp.Execute(frase)

Set Grid2.DataSource = rs
Grid2.TextMatrix(0, 0) = "Ticket "
Grid2.TextMatrix(0, 1) = "Placa    "
Grid2.TextMatrix(0, 2) = "Valor   "
Grid2.TextMatrix(0, 3) = "Validado          "
Grid2.TextMatrix(0, 4) = "Saida             "
Grid2.ColAlignment = flexAlignLeftCenter
Call FormataGridx(Grid2, rs)

frase = ""
frase = frase & "Select max(lticket)+ 1  from tb_ticket where dtimp is not null "
Set rs = dbApp.Execute(frase)
TXTINI = IIf(IsNull(rs(0)), "", rs(0))

frase = ""
frase = frase & "Select max(lticket)  from tb_ticket where dtimp is null "
Set rs = dbApp.Execute(frase)
txtfim = IIf(IsNull(rs(0)), TXTINI, rs(0))

If TXTINI = txtfim Then
    txtGerPag = 0
    Imprimecmd.Enabled = False
Else
    txtGerPag = (Val(txtfim) - Val(TXTINI) + 1) / 12
    Imprimecmd.Enabled = True
End If

Me.Refresh
Lercmd.Enabled = True

End Sub

Private Sub saircmd_Click()

Unload Me

End Sub

Private Sub txtGerPag_Change()

txtfim = Val(TXTINI) + Val(txtGerPag) * 12 - 1

End Sub
