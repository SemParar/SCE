VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTagTicket 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissao de Tickets"
   ClientHeight    =   9225
   ClientLeft      =   90
   ClientTop       =   480
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   13245
   Begin VB.Frame Frame4 
      Caption         =   "Movimentos"
      Height          =   2535
      Left            =   120
      TabIndex        =   40
      Top             =   0
      Width           =   12852
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridMovimentos 
         Height          =   2175
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   3836
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
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
   Begin VB.Frame Frame1 
      Caption         =   "Ultimos  Movimentos"
      Height          =   1335
      Left            =   10560
      TabIndex        =   38
      Top             =   6480
      Width           =   2415
      Begin VB.TextBox txtid 
         Height          =   375
         Left            =   1800
         TabIndex        =   44
         Text            =   "txtid"
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdPesMov 
         Caption         =   "Pesquisar"
         Default         =   -1  'True
         Height          =   732
         Left            =   600
         Picture         =   "frmTagTicket.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   360
         Width           =   972
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ultima Ocorrencia"
      Height          =   852
      Left            =   120
      TabIndex        =   28
      Top             =   5640
      Width           =   12852
      Begin VB.TextBox txtLast 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   29
         Text            =   "txtLast"
         Top             =   360
         Width           =   12492
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pesquisa de STATUS"
      Height          =   735
      Left            =   120
      TabIndex        =   27
      Top             =   4920
      Width           =   12852
      Begin VB.Label LblTRN 
         Alignment       =   1  'Right Justify
         Caption         =   "lblTRN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11640
         TabIndex        =   37
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Registro :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Index           =   3
         Left            =   10320
         TabIndex        =   36
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblPatio 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPatio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   35
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblNela 
         Alignment       =   1  'Right Justify
         Caption         =   "lblNela"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         TabIndex        =   34
         Top             =   240
         Width           =   855
      End
      Begin VB.Label LBLCAD 
         Alignment       =   1  'Right Justify
         Caption         =   "lblCAd"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   33
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "P�tio :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Cad Nela :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Index           =   1
         Left            =   6840
         TabIndex        =   31
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Cad Tag :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Index           =   0
         Left            =   3480
         TabIndex        =   30
         Top             =   240
         Width           =   1215
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2535
      Left            =   120
      TabIndex        =   13
      Top             =   6600
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4471
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   420
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Dar Entrada"
      TabPicture(0)   =   "frmTagTicket.frx":0102
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdDarEntrada"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FraDarEntrada"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmbEntrada"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtReferencia"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Dar Saida"
      TabPicture(1)   =   "frmTagTicket.frx":011E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmbSaida"
      Tab(1).Control(1)=   "FraDarSaida"
      Tab(1).Control(2)=   "cmdDarSaida"
      Tab(1).Control(3)=   "cmdDarBaixa"
      Tab(1).Control(4)=   "Label2"
      Tab(1).ControlCount=   5
      Begin VB.TextBox txtReferencia 
         Height          =   375
         Left            =   1560
         TabIndex        =   43
         Text            =   "txtreferencia"
         Top             =   2040
         Width           =   4455
      End
      Begin VB.ComboBox cmbSaida 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         ItemData        =   "frmTagTicket.frx":013A
         Left            =   -74040
         List            =   "frmTagTicket.frx":013C
         TabIndex        =   24
         Text            =   "cmbsaida"
         Top             =   1560
         Width           =   5532
      End
      Begin VB.ComboBox cmbEntrada 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         ItemData        =   "frmTagTicket.frx":013E
         Left            =   1560
         List            =   "frmTagTicket.frx":0140
         TabIndex        =   23
         Text            =   "cmbEntrada"
         Top             =   1440
         Width           =   4455
      End
      Begin VB.Frame FraDarSaida 
         Caption         =   "Ultima Saida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74880
         TabIndex        =   20
         Top             =   360
         Width           =   3492
         Begin MSComCtl2.DTPicker DTPSaiDia 
            Height          =   492
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   1812
            _ExtentX        =   3201
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   16580611
            CurrentDate     =   37987
            MinDate         =   37987
         End
         Begin MSComCtl2.DTPicker DTPSaiHora 
            Height          =   492
            Left            =   1920
            TabIndex        =   22
            Top             =   360
            Width           =   1452
            _ExtentX        =   2566
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "hh:mm"
            Format          =   16580610
            UpDown          =   -1  'True
            CurrentDate     =   37987
            MinDate         =   37987
         End
      End
      Begin VB.CommandButton cmdDarSaida 
         Caption         =   "Sai"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Left            =   -70560
         Picture         =   "frmTagTicket.frx":0142
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   600
         Width           =   972
      End
      Begin VB.CommandButton cmdDarBaixa 
         Caption         =   "M�nima  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Left            =   -69480
         TabIndex        =   18
         Top             =   600
         Width           =   972
      End
      Begin VB.Frame FraDarEntrada 
         Caption         =   "Ultima Entrada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   3492
         Begin MSComCtl2.DTPicker DTPEntHora 
            Height          =   492
            Left            =   1920
            TabIndex        =   16
            Top             =   360
            Width           =   1452
            _ExtentX        =   2566
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "hh:mm"
            Format          =   16580610
            UpDown          =   -1  'True
            CurrentDate     =   37987
            MinDate         =   37987
         End
         Begin MSComCtl2.DTPicker DTPEntDia 
            Height          =   492
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   1812
            _ExtentX        =   3201
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   16580611
            CurrentDate     =   37987
            MinDate         =   37987
         End
      End
      Begin VB.CommandButton cmdDarEntrada 
         Caption         =   "Entra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Left            =   4920
         Picture         =   "frmTagTicket.frx":0844
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   600
         Width           =   972
      End
      Begin VB.Label Label4 
         Caption         =   "Refer�ncia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   42
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "PISTA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -74880
         TabIndex        =   26
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Acesso:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   25
         Top             =   1560
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdimprime 
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
      Height          =   732
      Left            =   10680
      Picture         =   "frmTagTicket.frx":0F46
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8160
      Width           =   972
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
      Left            =   11880
      Picture         =   "frmTagTicket.frx":1250
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8160
      Width           =   972
   End
   Begin VB.Frame FraGrid 
      Caption         =   "Pesquisa do Tag"
      Height          =   2175
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   12852
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
         Height          =   1815
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   3201
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
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
   Begin VB.Frame FraTag 
      Caption         =   "Tag"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   6840
      TabIndex        =   0
      Top             =   6480
      Width           =   3615
      Begin VB.CommandButton cmdPes 
         Caption         =   "Pesquisar"
         Height          =   732
         Left            =   600
         Picture         =   "frmTagTicket.frx":155A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1680
         Width           =   972
      End
      Begin VB.CommandButton cmdLimpa 
         Caption         =   "Limpa"
         Height          =   732
         Left            =   2040
         Picture         =   "frmTagTicket.frx":165C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1680
         Width           =   972
      End
      Begin VB.TextBox TxtEmissor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   1320
         TabIndex        =   4
         Text            =   "Emi"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox TxtTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   1320
         TabIndex        =   3
         Text            =   "tag"
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox TxtPlaca 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1320
         TabIndex        =   2
         Text            =   "Placa"
         Top             =   1200
         Width           =   1812
      End
      Begin VB.Label lblemissor 
         Alignment       =   1  'Right Justify
         Caption         =   "Emissor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblPlaca 
         Alignment       =   1  'Right Justify
         Caption         =   "Placa:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Lbltag 
         Alignment       =   1  'Right Justify
         Caption         =   "Tag:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmTagTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AuxEntData As String
Dim AuxSaiData As String
Dim Auxtag As String
Dim Auxemissor As String
Dim Auxplaca As String
Dim frase As String
Dim rs As New Recordset
Dim rshist As New Recordset
Dim tadentro As Boolean
Dim tacadastro As Boolean
Dim tanela As Boolean

Private Sub cmdDarBaixa_Click()

DTPSaiDia.MinDate = DTPEntDia.Value
DTPSaiHora.MinDate = DTPSaiDia.MinDate
DTPSaiDia.Value = DTPEntDia.Value
DTPSaiHora.Value = DTPEntDia.Value
DTPSaiHora.Minute = (Val(DTPSaiHora.Minute)) Mod 60
AuxSaiData = DTPSaiHora
cmbSaida.ListIndex = cmbSaida.ListCount - 1
cmbSaida.Refresh
Call cmdDarSaida_Click

End Sub

Private Sub cmdDarEntrada_Click()
Dim aux As String
If cmbEntrada = "" Then
    MsgBoxService "ESCOLHA UMA PISTA", vbOKOnly, "MENSAGEM"
    cmbEntrada.SetFocus
    Exit Sub
End If

'Removida obrigatoriedade do campo referencia
'2015-05-01 - Rodrigo

'If txtReferencia = "" Then
'   MsgBoxService "Digite Marca Modelo", vbOKOnly, "MENSAGEM"
'    txtReferencia.SetFocus
'    Exit Sub
'End If


AuxEntData = Format(DTPEntHora, "DD/MM/YYYY HH:MM")
AuxSaiData = "19800101 00:00"
Auxtag = Format(TxtEmissor, "00000") + Format(TxtTag, "0000000000")

frase = ""
frase = frase + "Placa:"
frase = frase + String(30, " ") + TxtPlaca + vbCrLf
frase = frase + "Tag:"
frase = frase + String(30, " ") + TxtEmissor + "-" + TxtTag + vbCrLf
frase = frase + "Marca Modelo:"
frase = frase + String(30, " ") + txtReferencia + vbCrLf
frase = frase + "ENTRAR:" + vbCrLf
frase = frase + String(30, " ") + AuxEntData + vbCrLf
frase = frase + vbCrLf
frase = frase + "Confirma ?  "
If CVDate(AuxEntData) < CVDate(Format(Now(), "DD/MM/YYYY 00:00")) Then
    fraseaux = "ENTRADA DE DIAS ANTERIORES"
Else
    fraseaux = ""
End If

aux = MyMsgBox(frase, vbOKCancel, "ENTRADA TCK MANUAL", fraseaux)

If aux = 1 Then
    frase = ""
    frase = frase & "INSERT INTO tb_praia ("
    frase = frase & "tsingresso,tsegresso,cplaca,lticket,sztexto,"
    frase = frase & "lparam,Ientpista,Ientauto,szmodelo,iissuer,ltag,lvalor)"
    frase = frase & " values ("
    frase = frase & "'" & Format(AuxEntData, "yyyymmdd hh:mm:00") & "',"
    frase = frase & "'19800101',"
    frase = frase & "'" & TxtPlaca & "',"
    frase = frase & "0,"
    frase = frase & "'Ent TCK',"
    frase = frase & Format(cmbEntrada.ItemData(cmbEntrada.ListIndex), "00") & ","
    frase = frase & Format(cmbEntrada.ItemData(cmbEntrada.ListIndex), "00") & ","
    frase = frase & "0,"
    frase = frase & "Null,"
    frase = frase & "'" & TxtEmissor & "',"
    frase = frase & "'" & TxtTag & "',"
    frase = frase & "0)"
    Set rs = dbApp.Execute(frase)
    
    If txtid <> 0 Then
        frase = ""
        frase = frase & "UPDATE TB_EVENTOS "
        frase = frase & "SET ICODIGO = 140 "
        frase = frase & "where id = " + Format(txtid)
        Set rs = dbApp.Execute(frase)
    End If
    
    Call GravaEventos(12, TxtPlaca, gsUser, Format(cmbEntrada.ItemData(cmbEntrada.ListIndex), "00"), "Ent TCK " + Auxtag + " = " + Format(AuxEntData, "yyyymmdd hh:mm:00"), Val(TxtEmissor), Val(TxtTag))
  
    For i = 1 To Val(gbTicket)
        imprime_ticket (AuxEntData)
    Next i
        
    Call cmdPesMov_Click

End If

Call cmdLimpa_Click
TxtTag = Mid(Auxtag, 6, 10)
TxtEmissor = Mid(Auxtag, 1, 5)
TxtPlaca = Auxplaca
Call cmdPes_Click

End Sub
Private Sub imprime_ticket(ByVal p_data As String)

'Define a quantidade de c�pias a serem impressas

' CABECALHO
Dim strBuff As String
Open gsPath_REL + "ticket_cabec.txt" For Input As #1
While Not EOF(1)
    Line Input #1, strBuff
    Printer.Print strBuff
Wend
Close #1

Printer.FontSize = 10
Print "Ticket : " & TxtEmissor.Text & "-" & TxtTag.Text
Printer.Print "Data : " & p_data
Printer.Print "Placa do Carro: " & TxtPlaca.Text
Printer.Print "Marca/Modelo: " & txtReferencia.Text
Printer.Print "Acesso:  " & cmbEntrada.Text
Printer.FontSize = 8

'RODAPE
Open gsPath_REL + "ticket_rodape.txt" For Input As #1
While Not EOF(1)
    Line Input #1, strBuff
    Printer.Print strBuff
Wend
Close #1
Printer.FontSize = 13
Printer.Print "       SEM PARAR - VIA F�CIL"
Printer.FontSize = 8
Printer.Print "."
Printer.FontSize = 10
'Printer.Print "   SCE - Controle de Estacionamentos "



Printer.EndDoc



End Sub


Private Sub cmdDarSaida_Click()
Dim aux As String
Dim Auxvalor  As String
Dim rsaux As Recordset

If cmbSaida = "" Then
    MsgBoxService "ESCOLHA UMA PISTA", vbOKOnly, "MENSAGEM"
    cmbSaida.SetFocus
    Exit Sub
End If


AuxSaiData = Format(DTPSaiHora, "DD/MM/YYYY HH:MM")
Auxtag = Format(TxtEmissor, "00000") + Format(TxtTag, "0000000000")

If CVDate(AuxEntData) > CVDate(AuxSaiData) Then AuxSaiData = AuxEntData
frase = ""
frase = frase + "Placa:" + vbCrLf
frase = frase + String(30, " ") + TxtPlaca + vbCrLf
frase = frase + "Tag:" + vbCrLf
frase = frase + String(30, " ") + TxtEmissor + "-" + TxtTag + vbCrLf
frase = frase + "Entrou:" + vbCrLf
frase = frase + String(30, " ") + AuxEntData + vbCrLf
frase = frase + "Sair: " + vbCrLf
frase = frase + String(30, " ") + AuxSaiData + vbCrLf
frase = frase + "Confirma   ? "

If DateDiff("h", AuxEntData, AuxSaiData) > 18 Then
    fraseaux = "Perman�ncia Acima de 18 horas"
Else
    fraseaux = ""
End If

aux = MyMsgBox(frase, vbOKCancel, "SAIDA MANUAL", fraseaux)

If aux = 1 Then
    frase = ""
    frase = frase & "select * from tb_praia "
    frase = frase & "where iissuer = '" + TxtEmissor + "' and ltag = '" & TxtTag & "' and tsegresso < '20040101'"
    Set rsaux = dbApp.Execute(frase)
    If rsaux.EOF = True Or rsaux.BOF = True Then
        MsgBoxService "Tag nao esta no estacionamento "
    Else
        frase = ""
        frase = frase & "update tb_praia set "
        frase = frase & "tsegresso = '" & Format(AuxSaiData, "yyyymmdd hh:mm:ss") & "',"
        frase = frase & "sztexto = '" & rsaux("sztexto") & " Sai Man',"
        frase = frase & "lparam = " & Val(IIf(IsNull(rsaux("lparam")), " ", rsaux("lparam"))) + cmbSaida.ItemData(cmbSaida.ListIndex) * 100 & ","
        frase = frase & "ISaiPista = " & cmbSaida.ItemData(cmbSaida.ListIndex) & ","
        frase = frase & "ISaiAuto = 0 "
        frase = frase & "where iissuer = '" + TxtEmissor + "' and ltag = '" & TxtTag & "' and tsegresso < '20050101'"
        Set rs = dbApp.Execute(frase)
        Set rs = Nothing
        Call GravaEventos(13, TxtPlaca, gsUser, Format(cmbSaida.ItemData(cmbSaida.ListIndex), "00"), "Sai Man " + Auxtag + " = " + Format(AuxSaiData, "yyyymmdd hh:mm:ss"), TxtEmissor, TxtTag)

    End If
End If


Call cmdLimpa_Click
TxtTag = Mid(Auxtag, 6, 10)
TxtEmissor = Auxemissor
TxtPlaca = Auxplaca
Call cmdPes_Click

End Sub

Private Sub cmdimprime_Click()
Dim extra() As String
Dim Filename As String


ReDim extra(1)
extra(0) = txtLast

Filename = "TAG_" & gsEst_Codigo & "_" & TxtTag & "_" & TxtAno + TxtMes + TxtDia & "_" + Format(Now, "YYYYMMDDHHMMSS") & ".HTML"
'Filename = "" & TxtTag & "_" & Format(Date, "YYYYMMDD") & ".html"

    Call ImprimeHeader(Filename, "Historico de Tag ")
    Call Imprimegrid(Filename, Grid1)
    Call ImprimeExtra(Filename, extra)
    Call ImprimeFooter(Filename, "Impresso em : " + Format(Now, "DD/MM/YY HH:MM:SS"))
    Call ImprimeRel(Filename)
    Call ImprimeRelDel(Filename)



End Sub

Public Sub cmdLimpa_Click()

cmbSaida.ListIndex = 0
cmbEntrada.ListIndex = 0

TxtPlaca.Enabled = True
TxtTag.Enabled = True
TxtEmissor.Enabled = True

TxtTag = ""
TxtEmissor = ""
TxtPlaca = ""
txtLast = ""
txtReferencia = ""
txtid = 0
LBLCAD.Caption = ""
lblPatio.Caption = ""
lblNela.Caption = ""
LblTRN.Caption = ""

FraDarEntrada.Enabled = False
FraDarSaida.Enabled = False

DTPEntDia = Format(Now(), "DD/MM/YYYY HH:MM")
DTPEntDia.MinDate = "01/01/1980"
DTPEntDia.MaxDate = "01/01/2999"

DTPEntHora = DTPEntDia
DTPEntHora.MinDate = "01/01/1980"
DTPEntHora.MaxDate = "01/01/2999"
AuxEntData = DTPEntHora

DTPSaiDia = Format(Now(), "DD/MM/YYYY HH:MM")
DTPSaiDia.MinDate = "01/01/1980"
DTPSaiDia.MaxDate = "01/01/2999"
DTPSaiHora = DTPSaiDia
DTPSaiHora.MinDate = "01/01/1980"
DTPSaiHora.MaxDate = "01/01/2999"
AuxSaiData = DTPSaiHora

SSTab1.Enabled = False
SSTab1.Tab = 0
cmdimprime.Enabled = False
cmdDarEntrada.Enabled = False
cmdDarBaixa.Enabled = False
cmdDarSaida.Enabled = False
cmdLimpa.Enabled = False
cmdPes.Enabled = True
Grid1.Clear
Grid1.Refresh


End Sub

Public Sub cmdPes_Click()
Dim aux As Integer
Dim PesqPlaca As Boolean
Dim NaoAchei As Boolean

cmdPes.Enabled = False
cmdLimpa.Enabled = True
cmbSaida.ListIndex = 0
cmbEntrada.ListIndex = 1

NaoAchei = True

Auxtag = ""
TxtTag = UCase(LTrim(RTrim(TxtTag)))
TxtPlaca = UCase(LTrim(RTrim(TxtPlaca)))
TxtEmissor = UCase(LTrim(RTrim(TxtEmissor)))
If TxtEmissor = "" Then TxtEmissor = "290"

If TxtTag <> "" Then
    If Val(TxtTag) > 2000000000 Then
        MsgBoxService "Numero de TAG Invalido - Acima do Limite"
        Call cmdLimpa_Click
        Exit Sub
    End If
    TxtPlaca = ""
    PesqPlaca = False
Else
    'elimina qualquer pesquisa com placa XXX / YYY
    If TxtPlaca = "XXX9999" Or TxtPlaca = "YYY9999" Then
        MsgBoxService "Placa : " + TxtPlaca + " - N�o pode ser pesquisada", vbOKOnly, "Alerta"
        Call cmdLimpa_Click
        Exit Sub
    End If
    If TxtPlaca = "" Then
        MsgBoxService "DIGITE = Numero de TAG ou Placa", vbOKOnly, "Alerta"
        Call cmdLimpa_Click
        Exit Sub
    End If
    TxtEmissor = ""
    TxtTag = ""
    PesqPlaca = True
End If

'procura na praia o TAG ou Placa -- OK
lblPatio.Caption = ""

If Not PesqPlaca Then
    frase = ""
    frase = frase + " select top 1 iissuer,ltag,cplaca from tb_praia "
    frase = frase + " where iissuer = '" + TxtEmissor + "' and ltag = '" + TxtTag + "'"
Else
    frase = ""
    frase = frase + " select top 1 iissuer,ltag,cplaca from tb_praia "
    frase = frase + " where  cplaca = '" + TxtPlaca + "'"
End If

Set rs = dbApp.Execute(frase)
If Not rs.EOF And Not rs.BOF Then
   If NaoAchei Then Call Achei(rs(0), rs(1), rs(2), NaoAchei)
   lblPatio.Caption = "SIM"
Else
   lblPatio.Caption = "NAO"
End If

'procura no Cadastro
'LBLCAD.Caption = ""
'If Not PesqPlaca Then
'    frase = ""
'    frase = frase + " select top 1 iissuer,ltag,cplaca from tb_cadtag "
'    frase = frase + " where iissuer = '" + TxtEmissor + "' and ltag = '" + TxtTag + "'"
'    Set rs = dbApp.Execute(frase)
'    If Not rs.EOF Then
'        If NaoAchei Then Call Achei(rs(0), rs(1), rs(2), NaoAchei)
'        LBLCAD.Caption = "SIM"
'    Else
'        LBLCAD.Caption = "NAO"
'    End If
'Else
'    frase = " select count(*) from tb_cadtag  where cplaca = '" + TxtPlaca + "'"
''    frase = frase + " and not exists (select ltag from tb_cadnela N where n.ltag = c.ltag )"
'    Set rs = dbApp.Execute(frase)
'    If rs(0) = 0 Then
'        LBLCAD.Caption = "NAO"
'    ElseIf rs(0) = 1 Then
'        frase = " select top 1 iissuer,ltag,cplaca from tb_cadtag where cplaca = '" + TxtPlaca + "'"
'        Set rs = dbApp.Execute(frase)
'        If Not rs.EOF Then
'            If NaoAchei Then Call Achei(rs(0), rs(1), rs(2), NaoAchei)
'            LBLCAD.Caption = "SIM"
'        End If
'    ElseIf rs(0) > 1 Then
'        frase = ""
'        frase = frase & " select t.iissuer, t.ltag,"
'        frase = frase & " (select 'Lista Nela nr.' + cast(lseqfile as varchar(10)) from tb_cadnela l where l.ltag = t.ltag and l.iissuer = t.iissuer)"
'        frase = frase & " from tb_cadtag t"
'        frase = frase & " where cplaca = '" + TxtPlaca + "'"
'        Set rs = dbApp.Execute(frase)
'        rs.MoveFirst
'        Do While Not rs.EOF
'            msg = msg & rs(0) & " - " & rs(1) & " : " & rs(2) & vbCr
'            rs.MoveNext
'        Loop
'        msg = "Placa : " + TxtPlaca + " - Tem mais que um tag no cadastro. Digite um Tag" & vbCr & vbCr & msg
'        MsgBoxService msg, vbOKOnly, "Alerta"
'        LBLCAD.Caption = "NAO"
'    End If
' End If

'procura no Cadastro
LBLCAD.Caption = ""
If Not PesqPlaca Then
    frase = ""
    frase = frase + " select top 1 iissuer,ltag,cplaca from tb_cadtag "
    frase = frase + " where iissuer = '" + TxtEmissor + "' and ltag = '" + TxtTag + "'"
    Set rs = dbApp.Execute(frase)
    If Not rs.EOF Then
        If NaoAchei Then Call Achei(rs(0), rs(1), rs(2), NaoAchei)
        PesqPlaca = False
        LBLCAD.Caption = "SIM"
    Else
        LBLCAD.Caption = "NAO"
    End If
Else
    frase = " select count(*) from tb_cadtag where cplaca = '" + TxtPlaca + "'"
    Set rs = dbApp.Execute(frase)
    If rs(0) = 0 Then
        LBLCAD.Caption = "NAO"
    End If
'    If rs(0) > 1 Then
'        frase = ""
'        frase = frase & " select t.iissuer, t.ltag,"
'        frase = frase & " (select 'Lista Nela nr.' + cast(lseqfile as varchar(10)) from tb_cadnela l where l.ltag = t.ltag and l.iissuer = t.iissuer)"
'        frase = frase & " from tb_cadtag t"
'        frase = frase & " where cplaca = '" + TxtPlaca + "'"
'        frase = frase & " order by t.iissuer,t.ltag "
'        Set rs = dbApp.Execute(frase)
'        rs.MoveFirst
'        Do While Not rs.EOF
'            msg = msg & rs(0) & " - " & rs(1) & " : " & rs(2) & vbCr
'            rs.MoveNext
'        Loop
'        msg = "Placa : " + TxtPlaca + " - Tem mais que um tag no cadastro. Digite um Tag" & vbCr & vbCr & msg
'        MsgBoxService msg, vbOKOnly, "Alerta"
'        LBLCAD.Caption = "NAO"
'   End If
   If rs(0) >= 1 Then
        frase = " select top 1 iissuer,ltag,cplaca from tb_cadtag where cplaca = '" + TxtPlaca + "'"
        Set rs = dbApp.Execute(frase)
        If Not rs.EOF Then
            If NaoAchei Then Call Achei(rs(0), rs(1), rs(2), NaoAchei)
            PesqPlaca = False
            LBLCAD.Caption = "SIM"
        End If
   End If
End If

'procura na Lista Nela
lblNela.Caption = ""
If Not PesqPlaca Then
    frase = ""
    frase = frase + " select top 1 iissuer,ltag,cplaca from tb_cadNela "
    frase = frase + " where iissuer = '" + TxtEmissor + "' and ltag = '" + TxtTag + "'"
    frase = frase + " and cst not in (" + gsCodNelaLivre + ")"
Else
    frase = ""
    frase = frase + " select top 1 iissuer,ltag,cplaca from tb_cadNela N"
    frase = frase + " where  cplaca = '" + TxtPlaca + "'"
    frase = frase + " and cst not in (" + gsCodNelaLivre + ")"
    frase = frase + " and (select count(*) from tb_cadtag where ltag = n.ltag) >1"
    
End If
Set rs = dbApp.Execute(frase)
If Not rs.EOF And Not rs.BOF Then
    If NaoAchei Then Call Achei(rs(0), rs(1), rs(2), NaoAchei)
    lblNela.Caption = "SIM"
Else
    lblNela.Caption = "NAO"
End If

'procura na Lista TRANSACAO
LblTRN.Caption = ""
If Not PesqPlaca Then
    frase = ""
    frase = frase + " select top 1 iissuer,ltag,cplaca from tb_transacao "
    frase = frase + " where iissuer = '" + TxtEmissor + "' and ltag = '" + TxtTag + "'"
Else
    frase = ""
    frase = frase + " select top 1 iissuer,ltag,cplaca from tb_transacao "
    frase = frase + " where  cplaca = '" + TxtPlaca + "'"
End If
Set rs = dbApp.Execute(frase)
If Not rs.EOF And Not rs.BOF Then
    If NaoAchei Then Call Achei(rs(0), rs(1), rs(2), NaoAchei)
    LblTRN.Caption = "SIM"
Else
    LblTRN.Caption = "NAO"
End If

If NaoAchei Then
    MsgBoxService "Digite uma Placa ou Tag", vbOKOnly, "Alerta"
    Call cmdLimpa_Click
    Exit Sub
End If

frase = "exec pr_PesqTAGPLACA " + TxtEmissor + "," + TxtTag + "," + TxtPlaca


'atualizar grid de historico
Grid1.Clear
Set rshist = Nothing
Set rshist = dbApp.Execute(frase)
Set Grid1.DataSource = rshist
Grid1.TextMatrix(0, 0) = "Placa   "
Grid1.TextMatrix(0, 1) = "Tag                "
Grid1.TextMatrix(0, 2) = "Entrada        "
Grid1.TextMatrix(0, 3) = "PE     "
Grid1.TextMatrix(0, 4) = "Saida          "
Grid1.TextMatrix(0, 5) = "PS     "
Grid1.TextMatrix(0, 6) = "Valor      "
Grid1.TextMatrix(0, 7) = "Obs                              "
Grid1.TextMatrix(0, 8) = "Tabela                           "
Grid1.ColAlignment = 1
Call FormataGridx(Grid1, rshist)
Grid1.Refresh

'atualizar txtln
'Set rs = Nothing
'txtLN = ""
'frase = ""
'frase = frase & " select top 1"
'frase = frase & "  lseqfile as seq,"
'frase = frase & "  szmodelo as modelo,"
'frase = frase & "  cplaca as placa,"
'frase = frase & "  'LN'"
'frase = frase & "  From tb_cadnela a"
'frase = frase & "  where iissuer = " + Format(Val(TxtEmissor)) + " and ltag = '" + Format(Val(TxtTag)) + "'"
'Set rs = dbApp.Execute(frase)
'If Not rs.EOF And Not rs.BOF Then
'    txtLN = "LN (" + Format(rs(0)) + ", " + UCase(rs(2)) + ", " + RTrim(UCase(rs(1))) + " )"
'    MsgBoxService "TAG NA LISTA NELA"
'End If

SSTab1.Enabled = False
FraDarSaida.Enabled = False
FraDarSaida.Enabled = False
cmdDarSaida.Enabled = False
cmdDarBaixa.Enabled = False
cmdDarEntrada.Enabled = False
   
AuxEntData = Format(Now(), "DD/MM/YYYY HH:MM:SS")
AuxSaiData = Format(Now(), "DD/MM/YYYY HH:MM:SS")
DTPEntDia.Value = Format(Now(), "DD/MM/YYYY HH:MM:SS")
DTPEntHora.Value = Format(Now(), "DD/MM/YYYY HH:MM:SS")
DTPSaiDia.Value = Format(Now(), "DD/MM/YYYY HH:MM:SS")
DTPSaiHora.Value = Format(Now(), "DD/MM/YYYY HH:MM:SS")

aux = 0
rshist.Requery

If rshist.BOF And rshist.EOF Then
   txtLast = "Nenhum registro deste Tag"
   FraDarEntrada.Enabled = True
   If Not tanela And tacadastro And Not tadentro Then
        SSTab1.Enabled = True
        SSTab1.Tab = 0
        cmdDarEntrada.Enabled = True
        FraDarSaida.Enabled = False
   End If
   aux = 0
Else
   rshist.MoveFirst
   AuxEntData = RTrim(Format(rshist("Entrada"), "DD/MM/YYYY HH:MM:SS"))
   AuxSaiData = RTrim(Format(rshist("Saida"), "DD/MM/YYYY HH:MM:SS"))
   If AuxSaiData <> "" Then
        If CVDate(AuxSaiData) >= CVDate("01/01/2004") Then
           'nao pode entrar antes da ultima saida
            DTPEntDia.MinDate = DateAdd("n", 1, CVDate(AuxSaiData))
            DTPEntHora.MinDate = DateAdd("n", 1, CVDate(AuxSaiData))
            If AuxEntData = "" Then
               If CVDate(AuxSaiData) < CVDate(Format(Now(), "DD/MM/YYYY 00:00")) Then
                    txtLast = AuxSaiData + " ====> " + RTrim(rshist("Obs"))
               Else
                    txtLast = AuxSaiData + " ====> " + RTrim(rshist("Obs"))
                    DTPEntDia = DateAdd("n", 1, CVDate(AuxSaiData))
                    DTPEntHora = DateAdd("n", 1, CVDate(AuxSaiData))
               End If
            Else
                txtLast = AuxSaiData + " ====> " + RTrim(rshist("Obs"))
            End If
            FraDarEntrada.Enabled = True
            If Not tanela And tacadastro And Not tadentro Then
               'carro nao esta dentro e nao esta na LN
               'Liberar Entrada
               SSTab1.Enabled = True
               SSTab1.Tab = 0
               cmdDarEntrada.Enabled = True
            End If
        End If
   Else
        txtLast = AuxEntData + " ====> " + RTrim(rshist("Obs")) + " N�o Saiu "
        FraDarSaida.Enabled = True
        'atualiza campo de ultima entrada
        DTPEntDia = AuxEntData
        DTPEntHora = DTPEntDia
        'nao pode entrar ou sair antes da ultima entrada
        DTPSaiDia.MinDate = DateAdd("n", 1, CVDate(AuxEntData))
        DTPSaiHora.MinDate = DTPSaiDia.MinDate
        DTPSaiDia = Format(Now(), "DD/MM/YYYY HH:MM:SS")
        DTPSaiHora = DTPSaiDia
        'carro entrou e nao saiu
        'liberar saida ou baixa
        SSTab1.Enabled = True
        SSTab1.Tab = 1
        cmdDarSaida.Enabled = True
        cmdDarBaixa.Enabled = True
   End If
End If

Auxplaca = TxtPlaca
Auxtag = TxtTag
Auxemissor = TxtEmissor

TxtPlaca.Enabled = False
TxtTag.Enabled = False
TxtEmissor.Enabled = False

cmdimprime.Enabled = True
cmdLimpa.SetFocus

End Sub


Private Sub cmdPesMov_Click()

frase = "exec pr_PesqTagTicket '" + Format(DateAdd("n", -gsMinutosTicketCancelados, Now()), "yyyymmdd hh:mm:ss") + "'"

'atualizar grid de historico
gridMovimentos.Clear
Set rshist = Nothing
Set rshist = dbApp.Execute(frase)
Set gridMovimentos.DataSource = rshist
gridMovimentos.TextMatrix(0, 0) = "Placa   "
gridMovimentos.TextMatrix(0, 1) = "Tag                "
gridMovimentos.TextMatrix(0, 2) = "Entrada        "
gridMovimentos.TextMatrix(0, 3) = "PE     "
gridMovimentos.TextMatrix(0, 4) = "Saida          "
gridMovimentos.TextMatrix(0, 5) = "PS     "
gridMovimentos.TextMatrix(0, 6) = "Valor      "
gridMovimentos.TextMatrix(0, 7) = "Obs                              "
gridMovimentos.TextMatrix(0, 8) = "Tabela                           "
gridMovimentos.TextMatrix(0, 9) = "ID    "

gridMovimentos.ColAlignment = 1
Call FormataGridx(gridMovimentos, rshist)
gridMovimentos.Refresh

SSTab1.Enabled = False
FraDarSaida.Enabled = False
FraDarSaida.Enabled = False
cmdDarSaida.Enabled = False
cmdDarBaixa.Enabled = False
cmdDarEntrada.Enabled = False
   
AuxEntData = Format(Now(), "DD/MM/YYYY HH:MM:SS")
AuxSaiData = Format(Now(), "DD/MM/YYYY HH:MM:SS")
DTPEntDia.Value = Format(Now(), "DD/MM/YYYY HH:MM:SS")
DTPEntHora.Value = Format(Now(), "DD/MM/YYYY HH:MM:SS")
DTPSaiDia.Value = Format(Now(), "DD/MM/YYYY HH:MM:SS")
DTPSaiHora.Value = Format(Now(), "DD/MM/YYYY HH:MM:SS")

aux = 0
rshist.Requery

If rshist.BOF And rshist.EOF Then
   txtLast = "Nenhum registro deste Tag"
   FraDarEntrada.Enabled = True
   If Not tanela And tacadastro And Not tadentro Then
        SSTab1.Enabled = True
        SSTab1.Tab = 0
        cmdDarEntrada.Enabled = True
        FraDarSaida.Enabled = False
   End If
   aux = 0
Else
   rshist.MoveFirst
   AuxEntData = RTrim(Format(rshist("Entrada"), "DD/MM/YYYY HH:MM:SS"))
   AuxSaiData = RTrim(Format(rshist("Saida"), "DD/MM/YYYY HH:MM:SS"))
   If AuxSaiData <> "" Then
        If CVDate(AuxSaiData) >= CVDate("01/01/2004") Then
           'nao pode entrar antes da ultima saida
            DTPEntDia.MinDate = DateAdd("n", 1, CVDate(AuxSaiData))
            DTPEntHora.MinDate = DateAdd("n", 1, CVDate(AuxSaiData))
            If AuxEntData = "" Then
               If CVDate(AuxSaiData) < CVDate(Format(Now(), "DD/MM/YYYY 00:00")) Then
                    txtLast = AuxSaiData + " ====> " + RTrim(rshist("Obs"))
               Else
                    txtLast = AuxSaiData + " ====> " + RTrim(rshist("Obs"))
                    DTPEntDia = DateAdd("n", 1, CVDate(AuxSaiData))
                    DTPEntHora = DateAdd("n", 1, CVDate(AuxSaiData))
               End If
            Else
                txtLast = AuxSaiData + " ====> " + RTrim(rshist("Obs"))
            End If
            FraDarEntrada.Enabled = True
            If Not tanela And tacadastro And Not tadentro Then
               'carro nao esta dentro e nao esta na LN
               'Liberar Entrada
               SSTab1.Enabled = True
               SSTab1.Tab = 0
               cmdDarEntrada.Enabled = True
            End If
        End If
   Else
        txtLast = AuxEntData + " ====> " + RTrim(rshist("Obs")) + " N�o Saiu "
        FraDarSaida.Enabled = True
        'atualiza campo de ultima entrada
        DTPEntDia = AuxEntData
        DTPEntHora = DTPEntDia
        'nao pode entrar ou sair antes da ultima entrada
        DTPSaiDia.MinDate = DateAdd("n", 1, CVDate(AuxEntData))
        DTPSaiHora.MinDate = DTPSaiDia.MinDate
        DTPSaiDia = Format(Now(), "DD/MM/YYYY HH:MM:SS")
        DTPSaiHora = DTPSaiDia
        'carro entrou e nao saiu
        'liberar saida ou baixa
        SSTab1.Enabled = True
        SSTab1.Tab = 1
        cmdDarSaida.Enabled = True
        cmdDarBaixa.Enabled = True
   End If
End If

'Auxplaca = TxtPlaca
'Auxtag = TxtTag
'Auxemissor = TxtEmissor
'
'TxtPlaca.Enabled = False
'TxtTag.Enabled = False
'TxtEmissor.Enabled = False
'
'cmdimprime.Enabled = True
'cmdLimpa.SetFocus

End Sub

Private Sub DTPEntDia_change()
DTPEntDia.Value = Format(DTPEntDia.Value, "DD/MM/YYYY HH:MM:SS")
DTPEntDia.MaxDate = Now()
DTPEntHora.MaxDate = DTPEntDia.MaxDate
DTPEntHora.Value = DTPEntDia.Value
AuxEntData = DTPEntHora
End Sub
Private Sub DTPSaiDia_change()
DTPSaiDia.Value = Format(DTPSaiDia.Value, "DD/MM/YYYY HH:MM:SS")
DTPSaiDia.MaxDate = Now()
DTPSaiHora.MaxDate = DTPSaiDia.MaxDate
DTPSaiHora.Value = DTPSaiDia.Value
AuxSaiData = DTPSaiHora
End Sub
Private Sub DTPEnthora_change()
DTPEntHora.Value = Format(DTPEntHora.Value, "DD/MM/YYYY HH:MM:SS")
DTPEntDia.MaxDate = Now()
DTPEntHora.MaxDate = DTPEntDia.MaxDate
DTPEntDia.Value = DTPEntHora.Value
AuxEntData = DTPEntHora
End Sub
Private Sub DTPsaihora_change()
DTPSaiHora.Value = Format(DTPSaiHora.Value, "DD/MM/YYYY HH:MM:SS")
DTPSaiDia.MaxDate = Now()
DTPSaiHora.MaxDate = DTPSaiDia.MaxDate
DTPSaiDia.Value = DTPSaiHora.Value
AuxSaiData = DTPSaiHora
End Sub


Private Sub Form_Load()

On Error GoTo trataerr

'carrega combo de entradas
frase = "select * from tb_pista where ctipo = 'E' order by cdescricao"
Set rsGeral = dbApp.Execute(frase)
cmbEntrada.Clear
cmbEntrada.AddItem ""
ind = 0
cmbEntrada.ItemData(ind) = 0
If Not rsGeral.EOF And Not rsGeral.BOF Then
   rsGeral.MoveFirst
   Do While Not rsGeral.EOF
        cmbEntrada.AddItem Format(rsGeral("ipista"), "00") + " - " + LTrim(RTrim(rsGeral("cdescricao")))
        ind = ind + 1
        cmbEntrada.ItemData(ind) = rsGeral("ipista")
        rsGeral.MoveNext
   Loop
   cmbEntrada.ListIndex = 1
End If
'carrega combo de saidas
frase = "select * from tb_pista where ctipo = 'S' order by cdescricao"
Set rsGeral = dbApp.Execute(frase)
cmbSaida.Clear
cmbSaida.AddItem ""
ind = 0
cmbSaida.ItemData(ind) = 0
If Not rsGeral.EOF And Not rsGeral.BOF Then
   rsGeral.MoveFirst
   Do While Not rsGeral.EOF
        cmbSaida.AddItem Format(rsGeral("ipista"), "00") + " - " + LTrim(RTrim(rsGeral("cdescricao")))
        ind = ind + 1
        cmbSaida.ItemData(ind) = rsGeral("ipista")
        rsGeral.MoveNext
   Loop
   cmbSaida.ListIndex = 1
End If
Set rsGeral = Nothing


Call cmdLimpa_Click

Exit Sub

trataerr:
Call TrataErro(App.title, Me.Name, "form_load")
End Sub


Private Sub cmdsair_Click()

Unload Me

End Sub



Private Sub gridMovimentos_Click()

Call cmdLimpa_Click

gridMovimentos.ColSel = gridMovimentos.Col
gridMovimentos.RowSel = gridMovimentos.Row
gridMovimentos.BackColorSel = vbBlue
Text7 = UCase(LTrim(RTrim(gridMovimentos.TextMatrix(gridMovimentos.Row, 1))))
TxtEmissor.Text = Mid(Text7, 1, 5)
TxtTag.Text = Mid(Text7, 6, 10)
If Val(gridMovimentos.TextMatrix(gridMovimentos.Row, 9)) > 0 Then
    txtid = Val(gridMovimentos.TextMatrix(gridMovimentos.Row, 9))
Else
    txtid = 0
End If

Call cmdPes_Click

End Sub

Private Sub LBLCAD_Change()

If LBLCAD.Caption = "SIM" Then
    tacadastro = True
Else
    tacadastro = False
End If

End Sub
Private Sub LBLPatio_Change()

If lblPatio.Caption = "SIM" Then
    tadentro = True
Else
    tadentro = False
End If

End Sub
Private Sub LBLNela_Change()

If lblNela.Caption = "SIM" Then
    tanela = True
Else
    tanela = False
End If

End Sub

Private Sub Achei(Emi, tag, placa, NAchei As Boolean)

TxtEmissor = Emi
TxtTag = tag
TxtPlaca = placa
Auxtag = Format(TxtEmissor, "00000") + Format(TxtTag, "0000000000")
Auxplaca = TxtPlaca
NAchei = False

End Sub

Private Sub LblTRN_Click()

LblTRN.Caption = "SIM"

End Sub

