VERSION 5.00
Begin VB.Form frmCADArq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastros"
   ClientHeight    =   4050
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   13395
   Icon            =   "frmCadArq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   13395
   Begin VB.Frame Frame2 
      Caption         =   "Cadastro de Tags  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1452
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   12972
      Begin VB.TextBox TxtCD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox TxtCD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   960
         TabIndex        =   16
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TxtCD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   7440
         TabIndex        =   15
         Top             =   600
         Width           =   5292
      End
      Begin VB.TextBox TxtCD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   5040
         TabIndex        =   14
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox TxtCD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   2400
         TabIndex        =   13
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label lblCD 
         Caption         =   "Label1"
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
         Index           =   5
         Left            =   5040
         TabIndex        =   23
         Top             =   1080
         Width           =   7692
      End
      Begin VB.Label lblCD 
         Caption         =   "Label1"
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
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   612
      End
      Begin VB.Label lblCD 
         Caption         =   "Label1"
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
         Left            =   960
         TabIndex        =   21
         Top             =   360
         Width           =   1332
      End
      Begin VB.Label lblCD 
         Caption         =   "Label1"
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
         Index           =   2
         Left            =   2400
         TabIndex        =   20
         Top             =   360
         Width           =   2532
      End
      Begin VB.Label lblCD 
         Caption         =   "Label1"
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
         Index           =   3
         Left            =   5040
         TabIndex        =   19
         Top             =   360
         Width           =   2292
      End
      Begin VB.Label lblCD 
         Caption         =   "Label1"
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
         Index           =   4
         Left            =   7440
         TabIndex        =   18
         Top             =   360
         Width           =   3132
      End
   End
   Begin VB.CommandButton cmdArquivo 
      Caption         =   "&Arquivos"
      Height          =   855
      Left            =   10440
      Picture         =   "frmCadArq.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   11880
      Picture         =   "frmCadArq.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Cadastro de Lista Nela  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1452
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   12972
      Begin VB.TextBox TxtLN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   2400
         TabIndex        =   11
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox TxtLN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   5040
         TabIndex        =   10
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox TxtLN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   7440
         TabIndex        =   9
         Top             =   600
         Width           =   5292
      End
      Begin VB.TextBox TxtLN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   960
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TxtLN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblLN 
         Caption         =   "Label1"
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
         Index           =   5
         Left            =   5040
         TabIndex        =   25
         Top             =   1080
         Width           =   7692
      End
      Begin VB.Label lblLN 
         Caption         =   "Label1"
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
         Index           =   3
         Left            =   5040
         TabIndex        =   24
         Top             =   360
         Width           =   2292
      End
      Begin VB.Label lblLN 
         Caption         =   "Label1"
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
         Index           =   4
         Left            =   7440
         TabIndex        =   8
         Top             =   360
         Width           =   3132
      End
      Begin VB.Label lblLN 
         Caption         =   "Label1"
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
         Index           =   2
         Left            =   2400
         TabIndex        =   7
         Top             =   360
         Width           =   2532
      End
      Begin VB.Label lblLN 
         Caption         =   "Label1"
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
         Left            =   960
         TabIndex        =   6
         Top             =   360
         Width           =   1332
      End
      Begin VB.Label lblLN 
         Caption         =   "Label1"
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
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   612
      End
   End
End
Attribute VB_Name = "frmCADArq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsctl As New Recordset

Private Sub cmdArquivo_Click()

'LN
frase = ""
frase = frase & "SELECT top 1 ctipo as TIPO, lseqfile as SEQ, ltotal as REGFINAL,"
frase = frase & " tsatualizacao as DTATU, szarquivo as ARQ"
frase = frase & " FROM tb_CadnelaCtl "
frase = frase & " ORDER BY lseqfile DESC, szArquivo DESC"
Set rsctl = dbApp.Execute(frase)
lblLN(0) = "Tipo"
lblLN(1) = "Sequencial"
lblLN(2) = "Registros"
lblLN(3) = "Data de Atualização"
lblLN(4) = "Ultimo Arquivo Recebido"
lblLN(5) = ""

If rsctl.EOF Or rsctl.BOF Then
    TxtLN(4) = "Nenhum Arquivo Registrado"
Else
    TxtLN(0) = rsctl("Tipo")
    TxtLN(1) = rsctl("SEQ")
    TxtLN(2) = rsctl("REGFINAL")
    TxtLN(3) = rsctl("DTATU")       'Mid(rsctl("DTATU"), 7, 2) + "/" + Mid(rsctl("DTATU"), 5, 2) + "/" + Mid(rsctl("DTATU"), 3, 2) + " " + Mid(rsctl("DTATU"), 9, 2) + ":" + Mid(rsctl("DTATU"), 11, 2) + ":" + Mid(rsctl("DTATU"), 13, 2)
    TxtLN(4) = rsctl("ARQ")
    frase = Mid(rsctl("ARQ"), 7, 2) + "/" + Mid(rsctl("ARQ"), 5, 2) + "/" + Mid(rsctl("ARQ"), 1, 4) + " " + Mid(rsctl("ARQ"), 9, 2) + ":" + Mid(rsctl("ARQ"), 11, 2) + ":" + Mid(rsctl("ARQ"), 13, 2)
    If DateDiff("h", frase, Now()) < 24 Then
        Frame1.BackColor = &H8000000F
    Else
        Frame1.BackColor = &H80FFFF
        lblLN(5) = "Verificar a Recepção de LN"
    End If
End If

'CD
frase = ""
frase = frase & "SELECT top 1 ctipo as TIPO, lseqfile as SEQ, ltotal as REGFINAL,"
frase = frase & " tsatualizacao as DTATU, szarquivo as ARQ"
frase = frase & " FROM tb_CadtagCtl "
frase = frase & " ORDER BY szArquivo DESC ,lSeqfile DESC "
Set rsctl = dbApp.Execute(frase)
lblCD(0) = "Tipo"
lblCD(1) = "Sequencial"
lblCD(2) = "Registros"
lblCD(3) = "Data de Atualização"
lblCD(4) = "Ultimo Arquivo Recebido"
lblCD(5) = ""
If rsctl.EOF Or rsctl.BOF Then
    TxtCD(4) = "Nenhum Arquivo Registrado"
Else
    TxtCD(0) = rsctl("Tipo")
    TxtCD(1) = rsctl("SEQ")
    TxtCD(2) = rsctl("REGFINAL")
    TxtCD(3) = rsctl("DTATU") 'Mid(rsctl("DTATU"), 7, 2) + "/" + Mid(rsctl("DTATU"), 5, 2) + "/" + Mid(rsctl("DTATU"), 3, 2) + " " + Mid(rsctl("DTATU"), 9, 2) + ":" + Mid(rsctl("DTATU"), 11, 2) + ":" + Mid(rsctl("DTATU"), 13, 2)
    TxtCD(4) = rsctl("ARQ")
    frase = Mid(rsctl("ARQ"), 7, 2) + "/" + Mid(rsctl("ARQ"), 5, 2) + "/" + Mid(rsctl("ARQ"), 1, 4)
    If DateDiff("h", frase, Now()) < 24 Then
        Frame2.BackColor = &H8000000F
    Else
        Frame2.BackColor = &H80FFFF
        lblCD(5) = "Verificar a Recepção de Cadastro de TAGS"
    End If
End If

Me.Refresh

End Sub



Private Sub cmdsair_Click()

Unload Me

End Sub


Private Sub Form_Load()

Me.Top = 10
Me.Left = 10

Call cmdArquivo_Click

End Sub

Private Sub Form_Unload(Cancel As Integer)

Set rsctl = Nothing

End Sub
