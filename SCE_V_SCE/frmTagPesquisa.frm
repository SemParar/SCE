VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTagPes 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisa de Tags"
   ClientHeight    =   8370
   ClientLeft      =   90
   ClientTop       =   480
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   12030
   Begin VB.Frame Frame3 
      Caption         =   "Ultima Ocorrencia"
      Height          =   732
      Left            =   120
      TabIndex        =   28
      Top             =   3000
      Width           =   11652
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
         Height          =   408
         Left            =   120
         TabIndex        =   29
         Text            =   "txtLast"
         Top             =   240
         Width           =   11412
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pesquisa de STATUS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   27
      Top             =   3840
      Width           =   11652
      Begin VB.Label lblCadLAtivo 
         BackStyle       =   0  'Transparent
         Caption         =   "lblCadLAtivo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10080
         TabIndex        =   46
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label label3 
         Alignment       =   2  'Center
         Caption         =   "Patio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   45
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label label3 
         Alignment       =   2  'Center
         Caption         =   "Cadastro Local"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   6
         Left            =   8520
         TabIndex        =   44
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label label3 
         Alignment       =   2  'Center
         Caption         =   "Cadastro Central"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   5
         Left            =   4320
         TabIndex        =   43
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Ativo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   8040
         TabIndex        =   42
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Cad Local:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   8040
         TabIndex        =   41
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblCadL 
         BackStyle       =   0  'Transparent
         Caption         =   "lblCadL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10080
         TabIndex        =   40
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   8040
         TabIndex        =   39
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblCadLTipo 
         BackStyle       =   0  'Transparent
         Caption         =   "lblCadLTipo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10080
         TabIndex        =   38
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label LblTRN 
         Caption         =   "lblTRN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1710
         TabIndex        =   37
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label label3 
         Caption         =   "Registro:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   36
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblPatio 
         Caption         =   "lblPatio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1710
         TabIndex        =   35
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblNela 
         Caption         =   "lblNela"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   34
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label LBLCAD 
         BackStyle       =   0  'Transparent
         Caption         =   "lblCAd"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   33
         Top             =   720
         Width           =   975
      End
      Begin VB.Label label3 
         Caption         =   "Pátio :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   32
         Top             =   720
         Width           =   975
      End
      Begin VB.Label label3 
         Caption         =   "Cad Nela:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3840
         TabIndex        =   31
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label label3 
         Caption         =   "Cad Tag:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3840
         TabIndex        =   30
         Top             =   720
         Width           =   1335
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1815
      Left            =   120
      TabIndex        =   13
      Top             =   6240
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3201
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
      TabPicture(0)   =   "frmTagPesquisa.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdDarEntrada"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FraDarEntrada"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmbEntrada"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Dar Saida"
      TabPicture(1)   =   "frmTagPesquisa.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmbSaida"
      Tab(1).Control(1)=   "FraDarSaida"
      Tab(1).Control(2)=   "cmdDarSaida"
      Tab(1).Control(3)=   "cmdDarBaixa"
      Tab(1).Control(4)=   "Label2"
      Tab(1).ControlCount=   5
      Begin VB.ComboBox cmbSaida 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmTagPesquisa.frx":0038
         Left            =   -74160
         List            =   "frmTagPesquisa.frx":003A
         TabIndex        =   24
         Text            =   "cmbsaida"
         Top             =   1320
         Width           =   5415
      End
      Begin VB.ComboBox cmbEntrada 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmTagPesquisa.frx":003C
         Left            =   840
         List            =   "frmTagPesquisa.frx":003E
         TabIndex        =   23
         Text            =   "cmbEntrada"
         Top             =   1320
         Width           =   5295
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
         Top             =   240
         Width           =   4095
         Begin MSComCtl2.DTPicker DTPSaiDia 
            Height          =   495
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   1800
            _ExtentX        =   3175
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
            Format          =   57081859
            CurrentDate     =   37987
            MinDate         =   37987
         End
         Begin MSComCtl2.DTPicker DTPSaiHora 
            Height          =   495
            Left            =   2160
            TabIndex        =   22
            Top             =   360
            Width           =   1800
            _ExtentX        =   3175
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
            Format          =   57081858
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
         Height          =   855
         Left            =   -70680
         Picture         =   "frmTagPesquisa.frx":0040
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdDarBaixa 
         Caption         =   "Mínima  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -69600
         TabIndex        =   18
         Top             =   360
         Width           =   855
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
         Top             =   240
         Width           =   4095
         Begin MSComCtl2.DTPicker DTPEntHora 
            Height          =   495
            Left            =   2160
            TabIndex        =   16
            Top             =   360
            Width           =   1800
            _ExtentX        =   3175
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
            Format          =   57081858
            UpDown          =   -1  'True
            CurrentDate     =   37987
            MinDate         =   37987
         End
         Begin MSComCtl2.DTPicker DTPEntDia 
            Height          =   495
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   1800
            _ExtentX        =   3175
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
            Format          =   57081859
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
         Height          =   855
         Left            =   4320
         Picture         =   "frmTagPesquisa.frx":0742
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "PISTA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   216
         Left            =   -74760
         TabIndex        =   26
         Top             =   1320
         Width           =   732
      End
      Begin VB.Label Label1 
         Caption         =   "PISTA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   168
         Left            =   240
         TabIndex        =   25
         Top             =   1320
         Width           =   612
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
      Height          =   735
      Left            =   10920
      Picture         =   "frmTagPesquisa.frx":0E44
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   852
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
      Height          =   735
      Left            =   10920
      Picture         =   "frmTagPesquisa.frx":114E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7200
      Width           =   852
   End
   Begin VB.Frame FraGrid 
      Caption         =   "Ultimas Ocorrêcias"
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   11652
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
         Height          =   2055
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   3625
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
      Height          =   1932
      Left            =   6600
      TabIndex        =   0
      Top             =   6120
      Width           =   4215
      Begin VB.CommandButton cmdPes 
         Caption         =   "Pesquisar"
         Default         =   -1  'True
         Height          =   612
         Left            =   3240
         Picture         =   "frmTagPesquisa.frx":1458
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   852
      End
      Begin VB.CommandButton cmdLimpa 
         Caption         =   "Limpa"
         Height          =   612
         Left            =   3240
         Picture         =   "frmTagPesquisa.frx":155A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1080
         Width           =   852
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
         Top             =   360
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
         Top             =   840
         Width           =   1815
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
         Top             =   1320
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
         Top             =   360
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
         Top             =   1440
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
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Central de Serviços Sem Parar: (11) 3004-9599"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   47
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmTagPes"
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

On Error GoTo trataerr

DTPSaiDia.MinDate = DTPEntDia.value
DTPSaiHora.MinDate = DTPSaiDia.MinDate
DTPSaiDia.value = DTPEntDia.value
DTPSaiHora.value = DTPEntDia.value
DTPSaiHora.Minute = (Val(DTPSaiHora.Minute)) Mod 60
AuxSaiData = DTPSaiHora
cmbSaida.ListIndex = cmbSaida.ListCount - 1
cmbSaida.Refresh
Call cmdDarSaida_Click

Exit Sub

trataerr:
    Call TrataErro(app.title, Error, "cmdDarBaixa_Click")

End Sub

Private Sub cmdDarEntrada_Click()

On Error GoTo trataerr

Dim aux As String

If cmbEntrada = "" Then
    'MsgBox "ESCOLHA UMA PISTA", vbOKOnly, "MENSAGEM"
    MsgBoxService "ESCOLHA UMA PISTA", vbOKOnly, "MENSAGEM"
    cmbEntrada.SetFocus
    Exit Sub
End If

AuxEntData = Format(DTPEntHora, "DD/MM/YYYY HH:MM")
AuxSaiData = "19800101 00:00"
Auxtag = Format(TxtEmissor, "00000") + Format(TxtTag, "0000000000")

frase = ""
frase = frase + "Placa:" + vbCrLf
frase = frase + String(30, " ") + TxtPlaca + vbCrLf
frase = frase + "Tag:" + vbCrLf
frase = frase + String(30, " ") + TxtEmissor + "-" + TxtTag + vbCrLf
frase = frase + "ENTRAR:" + vbCrLf
frase = frase + String(30, " ") + AuxEntData + vbCrLf
frase = frase + vbCrLf
frase = frase + "Confirma ?  "
If CVDate(AuxEntData) < CVDate(Format(Now(), "DD/MM/YYYY 00:00")) Then
    fraseaux = "ENTRADA DE DIAS ANTERIORES"
Else
    fraseaux = ""
End If

aux = MyMsgBox(frase, vbOKCancel, "ENTRADA MANUAL", fraseaux)

If aux = 1 Then
    frase = ""
    frase = frase & "SELECT COUNT(*) FROM TB_PRAIA "
    frase = frase & "WHERE IISSUER = '" & TxtEmissor & "' AND LTAG = '" & TxtTag & "'"
    Set rs = dbApp.Execute(frase)
    If rs(0) > 0 Then
        Call GravaEventos(12, TxtPlaca, gsUser, 0, "Tentativa de Entrada Já na Praia - Ent Man", Val(TxtEmissor), Val(TxtTag))
    Else
        frase = ""
        frase = frase & "INSERT INTO tb_praia ("
        frase = frase & "tsingresso,tsegresso,cplaca,lticket,sztexto,"
        frase = frase & "lparam,Ientpista,Ientauto,szmodelo,iissuer,ltag,lvalor)"
        frase = frase & " values ("
        frase = frase & "'" & Format(AuxEntData, "yyyymmdd hh:mm:00") & "',"
        frase = frase & "'19800101',"
        frase = frase & "'" & TxtPlaca & "',"
        frase = frase & "0,"
        frase = frase & "'Ent Man',"
        frase = frase & Format(cmbEntrada.ItemData(cmbEntrada.ListIndex), "00") & ","
        frase = frase & Format(cmbEntrada.ItemData(cmbEntrada.ListIndex), "00") & ","
        frase = frase & "0,"
        frase = frase & "Null,"
        frase = frase & "'" & TxtEmissor & "',"
        frase = frase & "'" & TxtTag & "',"
        frase = frase & "0)"
        Set rs = dbApp.Execute(frase)
        Call GravaEventos(12, TxtPlaca, gsUser, Format(cmbEntrada.ItemData(cmbEntrada.ListIndex), "00"), "Ent Man " + Auxtag + " = " + Format(AuxEntData, "yyyymmdd hh:mm:00"), Val(TxtEmissor), Val(TxtTag))
        End If
End If

Call cmdLimpa_Click
TxtTag = Mid(Auxtag, 6, 10)
TxtEmissor = Mid(Auxtag, 1, 5)
TxtPlaca = Auxplaca
Call cmdPes_Click

Exit Sub

trataerr:
    Call TrataErro(app.title, Error, "cmdDarEntrada_Click")


End Sub

Private Sub cmdDarSaida_Click()

On Error GoTo trataerr

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
    fraseaux = "Permanência Acima de 18 horas"
Else
    fraseaux = ""
End If

aux = MyMsgBox(frase, vbOKCancel, "SAIDA MANUAL", fraseaux)

If aux = 1 Then
    frase = ""
    frase = frase & "select * from tb_praia "
    frase = frase & "where iissuer = '" + TxtEmissor + "' and ltag = '" & TxtTag & "' and tsegresso < '20040101'"
    Set rsaux = dbApp.Execute(frase)
    If rsaux.EOF And rsaux.BOF Then
        MsgBoxService "Tag já saiu do Estacionamento "
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

Exit Sub

trataerr:
    Call TrataErro(app.title, Error, "MyMsgBox")
End Sub

Private Sub cmdimprime_Click()

On Error Resume Next

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
' Call ImprimeRelDel(Filename)




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
LBLCAD.Caption = ""
lblPatio.Caption = ""
lblNela.Caption = ""
LblTRN.Caption = ""
lblCadLAtivo.Caption = ""
lblCadL.Caption = ""
lblCadLTipo.Caption = ""


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

On Error GoTo trataerr

Dim aux As Integer
Dim PesqPlaca As Boolean
Dim NaoAchei As Boolean


cmdLimpa.Enabled = True
cmdPes.Enabled = False


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
        MsgBoxService "Placa : " + TxtPlaca + " - Não pode ser pesquisada", vbOKOnly, "Alerta"
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
   'PesqPlaca = True
   lblPatio.Caption = "SIM"
Else
   lblPatio.Caption = "NAO"
End If

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
        frase = " select top 1 iissuer,ltag,cplaca from tb_cadtag where cplaca = '" + TxtPlaca + "' order by iissuer,ltag desc"
        Set rs = dbApp.Execute(frase)
        If Not rs.EOF Then
            If NaoAchei Then Call Achei(rs(0), rs(1), rs(2), NaoAchei)
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
    
    Set rs = dbApp.Execute(frase)
    If Not rs.EOF And Not rs.BOF Then
        If NaoAchei Then Call Achei(rs(0), rs(1), rs(2), NaoAchei)
        PesqPlaca = False
        lblNela.Caption = "SIM"
    Else
        lblNela.Caption = "NAO"
    End If
Else

    'Busca o Tag na listanela
    frase = ""
    frase = frase + " select top 1 iissuer,ltag,cplaca from tb_cadNela "
    frase = frase + " where iissuer = '" + TxtEmissor + "' and ltag = '" + TxtTag + "'"
    frase = frase + " and cst not in (" + gsCodNelaLivre + ")"

    Set rs = dbApp.Execute(frase)
    If Not rs.EOF And Not rs.BOF Then
        If NaoAchei Then Call Achei(rs(0), rs(1), rs(2), NaoAchei)
        PesqPlaca = False
        lblNela.Caption = "SIM"
    Else
        lblNela.Caption = "NAO"
    End If

    'frase = " select count(*) from tb_cadNela where cplaca = '" + TxtPlaca + "'"
'    Set rs = dbApp.Execute(frase)
'    If rs(0) = 0 Then
'        lblNela.Caption = "NAO"
'    End If
'    If rs(0) > 1 Then
'        lblNela.Caption = "---"
'    End If
'    If rs(0) = 1 Then
'
'
'
'        frase = ""
'        frase = frase + " select top 1 iissuer,ltag,cplaca from tb_cadNela "
'        frase = frase + " where  cplaca = '" + TxtPlaca + "'"
'        frase = frase + " and cst not in (" + gsCodNelaLivre + ")"
'        Set rs = dbApp.Execute(frase)
'
'        If Not rs.EOF And Not rs.BOF Then
'            If NaoAchei Then Call Achei(rs(0), rs(1), rs(2), NaoAchei)
'            PesqPlaca = False
'            lblNela.Caption = "SIM"
'
'        Else
'            PesqPlaca = False
'            lblNela.Caption = "NAO"
'
'        End If
'
'
'
'
'    End If
End If

'procura na Lista TRANSACAO
LblTRN.Caption = ""
If Not PesqPlaca Then
    frase = ""
    frase = frase + " select top 1 iissuer,ltag,cplaca from tb_transacao "
    frase = frase + " where iissuer = '" + TxtEmissor + "' and ltag = '" + TxtTag + "' order by tsdataoperacao desc"
Else
    frase = ""
    frase = frase + " select top 1 iissuer,ltag,cplaca from tb_transacao "
    frase = frase + " where  cplaca = '" + TxtPlaca + "' order by tsdataoperacao desc"
End If
Set rs = dbApp.Execute(frase)
If Not rs.EOF And Not rs.BOF Then
    If NaoAchei Then Call Achei(rs(0), rs(1), rs(2), NaoAchei)
    LblTRN.Caption = "SIM"
Else
    LblTRN.Caption = "NAO"
End If

'procura na Tb_usertag
lblCadL.Caption = ""
If Not PesqPlaca Then
    frase = ""
    frase = frase + " select top 1 iissuer,ltag,cplaca from tb_usertag "
    frase = frase + " where iissuer = '" + TxtEmissor + "' and ltag = '" + TxtTag + "'"
Else
    frase = ""
    frase = frase + " select top 1 iissuer,ltag,cplaca from tb_usertag "
    frase = frase + " where  cplaca = '" + TxtPlaca + "'"
End If

Set rs = dbApp.Execute(frase)
If Not rs.EOF And Not rs.BOF Then
    If NaoAchei Then Call Achei(rs(0), rs(1), rs(2), NaoAchei)
    frase = ""
    frase = frase + " select top 1 ctipotarifa, cast(cativo as int) from tb_usertag "
    'frase = frase + " where iissuer = '" + TxtEmissor + "' and ltag = '" + TxtTag + "'"
    frase = frase + " where cplaca = '" + TxtPlaca + "'"
    Set rs = dbApp.Execute(frase)
    lblCadL.Caption = "SIM"
    lblCadLTipo.Caption = Format(rs(0))
    lblCadLAtivo.Caption = Format(rs(1))
Else
    lblCadL.Caption = "NAO"
    lblCadLTipo.Caption = "---"
    lblCadLAtivo.Caption = "---"
End If

If NaoAchei Then
    MsgBoxService "Digite uma Placa ou Tag", vbOKOnly, "Alerta"
    Call cmdLimpa_Click
    Exit Sub
End If

'atualizar grid de historico
Grid1.Clear
Set rshist = Nothing




If Not PesqPlaca Then
    frase = ""
    frase = "exec Pr_PesqTAGPLACA " + TxtEmissor + "," + TxtTag + ",''"
Else
    frase = ""
    frase = "exec Pr_PesqTAGPLACA " + TxtEmissor + ",'','" + TxtPlaca + "'"
End If


' frase = "exec Pr_PesqTAG " + TxtEmissor + "," + TxtTag
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

FraDarSaida.Enabled = False
FraDarSaida.Enabled = False
   
   
AuxEntData = Format(Now(), "DD/MM/YYYY HH:MM:SS")
AuxSaiData = Format(Now(), "DD/MM/YYYY HH:MM:SS")
DTPEntDia.value = Format(Now(), "DD/MM/YYYY HH:MM:SS")
DTPEntHora.value = Format(Now(), "DD/MM/YYYY HH:MM:SS")
DTPSaiDia.value = Format(Now(), "DD/MM/YYYY HH:MM:SS")
DTPSaiHora.value = Format(Now(), "DD/MM/YYYY HH:MM:SS")

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
        txtLast = AuxEntData + " ====> " + RTrim(rshist("Obs")) + " Não Saiu "
        FraDarSaida.Enabled = True
        'atualiza campo de ultima entrada
        DTPEntDia = AuxEntData
        DTPEntHora = DTPEntDia
        'nao pode entrar ou sair antes da ultima entrada
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

Exit Sub

trataerr:
    Call TrataErro(app.title, Error, "cmdPes")

End Sub

Private Sub Command1_Click()
Dim objHTTP As New MSXML2.XMLHTTP
Dim strEnvelope As String
Dim strReturn As String
Dim objReturn As New MSXML2.DOMDocument
Dim dblTax As Double
Dim strQuery As String

'Create the SOAP Envelope
'strEnvelope = _
'"<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:est=""http://service.grupostp.com.br/estacionamento/EstacionamentoCelularFachada/"">" & _
'   "<soapenv:Header/>" & _
'   "<soapenv:Body>" & _
'      "<est:obterTagRequest>" & _
'         "<placa>FKX9673</placa>" & _
'      "</est:obterTagRequest>" & _
'   "</soapenv:Body>" & _
'"</soapenv:Envelope>"

strEnvelope = _
"<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:con=""http://service.grupostp.com.br/ConsultaQualidadeDispositivoFachada/"">" & _
   "<soapenv:Header/>" & _
   "<soapenv:Body>" & _
      "<con:consultarQualidadeLeituraDispositivoRequest>" & _
         "<placaVeiculo>ERE8321</placaVeiculo>" & _
      "</con:consultarQualidadeLeituraDispositivoRequest>" & _
   "</soapenv:Body>" & _
"</soapenv:Envelope>"


'Set up to post to our local server
objHTTP.Open "post", "http://intranet-portal.cgmp-osa.com.br:8898/EstacionamentoCelularFachadaPS", False

'Set a standard SOAP/ XML header for the content-type
objHTTP.setRequestHeader "Content-Type", "text/xml"

'Set a header for the method to be called
objHTTP.setRequestHeader "SOAPAction", "obterTag"

'Make the SOAP call
objHTTP.send strEnvelope

'Get the return envelope
strReturn = objHTTP.responseText

'Load the return envelope into a DOM
objReturn.LoadXml strReturn



End Sub


Private Sub DTPEntDia_change()
DTPEntDia.value = Format(DTPEntDia.value, "DD/MM/YYYY HH:MM:SS")
DTPEntDia.MaxDate = Now()
DTPEntHora.MaxDate = DTPEntDia.MaxDate
DTPEntHora.value = DTPEntDia.value
AuxEntData = DTPEntHora
End Sub
Private Sub DTPSaiDia_change()
DTPSaiDia.value = Format(DTPSaiDia.value, "DD/MM/YYYY HH:MM:SS")
DTPSaiDia.MaxDate = Now()
DTPSaiHora.MaxDate = DTPSaiDia.MaxDate
DTPSaiHora.value = DTPSaiDia.value
AuxSaiData = DTPSaiHora
End Sub
Private Sub DTPEnthora_change()
DTPEntHora.value = Format(DTPEntHora.value, "DD/MM/YYYY HH:MM:SS")
DTPEntDia.MaxDate = Now()
DTPEntHora.MaxDate = DTPEntDia.MaxDate
DTPEntDia.value = DTPEntHora.value
AuxEntData = DTPEntHora
End Sub
Private Sub DTPsaihora_change()
DTPSaiHora.value = Format(DTPSaiHora.value, "DD/MM/YYYY HH:MM:SS")
DTPSaiDia.MaxDate = Now()
DTPSaiHora.MaxDate = DTPSaiDia.MaxDate
DTPSaiDia.value = DTPSaiHora.value
AuxSaiData = DTPSaiHora
End Sub


Private Sub Form_Load()

On Error GoTo trataerr



Me.Top = 10
Me.Left = 10

'Esconder botoes de comandos manuais quando a variavel PermiteManuais estiver setada para zero no ini
If gsPermiteManuais = 0 Then
cmdDarSaida.Visible = False
cmdDarBaixa.Visible = False
cmdDarEntrada.Visible = False
End If


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
   cmbEntrada.ListIndex = 0
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

'On Error GoTo trataerr
trataerr:
Call TrataErro(app.title, Error, "Form_Load")

End Sub


Private Sub cmdsair_Click()

Unload Me

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

