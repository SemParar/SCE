VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRelMov 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório do Movimento"
   ClientHeight    =   9375
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   15930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   15930
   Begin VB.CheckBox chkImprimeTransacaoManuais 
      Caption         =   "Imprime"
      Height          =   315
      Left            =   14640
      TabIndex        =   50
      Top             =   3840
      Width           =   975
   End
   Begin VB.Frame Frame8 
      Caption         =   "Transações Saidas Manuais"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2532
      Left            =   3240
      TabIndex        =   47
      Top             =   3840
      Width           =   12495
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1Man 
         Height          =   2055
         Left            =   120
         TabIndex        =   48
         Top             =   360
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   3625
         _Version        =   393216
         BackColor       =   -2147483633
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
   Begin VB.Frame Frame7 
      Caption         =   "Resumo do Movimento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2532
      Left            =   120
      TabIndex        =   37
      Top             =   3840
      Width           =   2895
      Begin VB.TextBox txtTotMan 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         TabIndex        =   41
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtValorMan 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         TabIndex        =   40
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox TxtZeroMan 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         TabIndex        =   39
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtValMan 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         TabIndex        =   38
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Quantidade Manuais"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "Sem Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   45
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Com Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   44
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   2040
         Width           =   1095
      End
   End
   Begin VB.CheckBox chkImprimeDivisaoEntradas 
      Caption         =   "Imprime"
      Height          =   315
      Left            =   14640
      TabIndex        =   36
      Top             =   6600
      Width           =   975
   End
   Begin VB.CheckBox chkImprimeArquivosGerados 
      Caption         =   "Imprime"
      Height          =   315
      Left            =   8880
      TabIndex        =   35
      Top             =   6600
      Width           =   975
   End
   Begin VB.CheckBox chkImprimeDivisaoValores 
      Caption         =   "Imprime"
      Height          =   315
      Left            =   3240
      TabIndex        =   34
      Top             =   6600
      Width           =   975
   End
   Begin VB.CheckBox chkImprimeTransacao 
      Caption         =   "Imprime"
      Height          =   315
      Left            =   14640
      TabIndex        =   33
      Top             =   960
      Width           =   975
   End
   Begin VB.Frame Frame6 
      Caption         =   "Divisao Entradas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   10200
      TabIndex        =   31
      Top             =   6600
      Width           =   5535
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridEntrada 
         Height          =   1695
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   2990
         _Version        =   393216
         BackColor       =   -2147483633
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
   Begin VB.CommandButton CmdExportarCSV 
      Caption         =   "Exportar"
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
      Left            =   6120
      Picture         =   "frmRelMov.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   120
      Width           =   972
   End
   Begin VB.TextBox txtseqfile 
      BackColor       =   &H8000000F&
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
      Height          =   372
      Left            =   14640
      TabIndex        =   29
      Top             =   8880
      Width           =   1092
   End
   Begin VB.Frame Frame4 
      Caption         =   "Transações Saidas Automaticas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2532
      Left            =   3240
      TabIndex        =   27
      Top             =   960
      Width           =   12495
      Begin VB.CheckBox Check1 
         Caption         =   "Imprime"
         Height          =   315
         Left            =   11400
         TabIndex        =   49
         Top             =   2520
         Width           =   975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
         Height          =   2055
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   3625
         _Version        =   393216
         BackColor       =   -2147483633
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
   Begin VB.Frame Frame5 
      Caption         =   "Divisao Valores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2172
      Left            =   120
      TabIndex        =   25
      Top             =   6600
      Width           =   4215
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
         Height          =   1695
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   2990
         _Version        =   393216
         BackColor       =   -2147483633
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
      Caption         =   "Arquivos Gerados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   4440
      TabIndex        =   23
      Top             =   6600
      Width           =   5655
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdFiles 
         Height          =   1695
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   2990
         _Version        =   393216
         BackColor       =   -2147483633
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
      Height          =   732
      Left            =   8280
      Picture         =   "frmRelMov.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   120
      Width           =   972
   End
   Begin VB.CommandButton CmdRegerar 
      Caption         =   "Regerar"
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
      Left            =   7200
      Picture         =   "frmRelMov.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   120
      Width           =   972
   End
   Begin VB.Frame Frame2 
      Caption         =   " DIGITE A DATA DO MOVIMENTO   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   3732
      Begin VB.TextBox TxtDia 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   17
         Top             =   240
         Width           =   408
      End
      Begin VB.TextBox TxtMes 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         TabIndex        =   16
         Top             =   240
         Width           =   408
      End
      Begin VB.TextBox TxtAno 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3000
         TabIndex        =   15
         Text            =   "2004"
         Top             =   240
         Width           =   612
      End
      Begin VB.Label Label3 
         Caption         =   "DIA : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   372
      End
      Begin VB.Label Label1 
         Caption         =   "MES : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   1200
         TabIndex        =   19
         Top             =   240
         Width           =   492
      End
      Begin VB.Label Label2 
         Caption         =   "ANO : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   2400
         TabIndex        =   18
         Top             =   240
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
      Height          =   732
      Left            =   3960
      Picture         =   "frmRelMov.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   120
      Width           =   972
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
      Height          =   732
      Left            =   5040
      Picture         =   "frmRelMov.frx":0E98
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   120
      Width           =   972
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
      Height          =   732
      Left            =   14640
      Picture         =   "frmRelMov.frx":11A2
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   972
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resumo do Movimento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2532
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2895
      Begin VB.TextBox txtVal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox TxtZero 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         TabIndex        =   7
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox txtTot 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         TabIndex        =   6
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Com Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Sem Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Quantidade Automaticas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   8880
      Width           =   14295
   End
End
Attribute VB_Name = "frmRelMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Filename As String

Private Sub CmdExportarCSV_Click()
On Error GoTo trataerr
Dim Filename As String

Filename = "MOV_" & gsEst_Codigo & "_" & TxtAno & TxtMes & TxtDia & "_IMP" & Format(Now(), "YYYYMMDDHHMMSS") + ".csv"
Call ArqGrid(Filename, Grid1)
MsgBoxService "Arquivo " & Filename & " Exportado "

Exit Sub
'On Error GoTo trataerr
trataerr:
Call TrataErro(App.title, Me.Name, "CmdExportarCSV")

End Sub

Private Sub cmdGerar_Click()
'On Error GoTo trataerr
'Dim aux As Integer

'aux = MyMsgBox("Confirma a Mudança de Data de Geraçao", vbOKCancel, "GERAR ARQUIVO TRN", "Tem Certeza da Operação")

'If aux = 1 Then
'    ' Call Atualiza_NextTRN
'End If
'
'Exit Sub
''On Error GoTo trataerr
'trataerr:
'Call TrataErro(App.title, Me.Name, "cmdGerar_Click")

End Sub

Private Sub CmdRegerar_Click()
On Error GoTo trataerr
Dim aux As Integer

If Val(txtseqfile) > 0 Then
    aux = MyMsgBox("      Confirma a Regerar o Arquivo : " + txtseqfile, vbOKCancel, "REGERAR ARQUIVO TRN", "Tem Certeza da Operação")
    If aux = 1 And Val(txtseqfile) > 0 Then
        Call CriaFileTRN(TxtAno & TxtMes & TxtDia, Val(txtseqfile))
    End If
End If

Exit Sub
'On Error GoTo trataerr
trataerr:
Call TrataErro(App.title, Me.Name, "CMDRegerar")

End Sub

Private Sub Form_Load()

On Error GoTo trataerr

Dim frase As String
Dim rs As New Recordset
rs.CursorType = adOpenStatic

If gbytNivel_Acesso_Usuario <> gintNIVEL_ADMINISTRADOR Then
    cmdGerar.Visible = False
    CmdRegerar.Visible = False
End If

Me.Top = 50
Me.Left = 50


Set rsGeral = Nothing

frase = "select max(tsdatamovimento) from tb_transacao where lseqfile is not null"
Set rs = dbApp.Execute(frase)

If Not rs.BOF And Not rs.EOF And Not IsNull(rs(0)) Then
    TxtDia = Format(CVDate(rs(0)), "dd")
    TxtMes = Format(CVDate(rs(0)), "MM")
    TxtAno = Format(CVDate(rs(0)), "YYYY")
Else
    TxtDia = Format(Date, "dd")
    TxtMes = Format(Date, "MM")
    TxtAno = Format(Date, "YYYY")
End If

chkImprimeDivisaoValores.Value = 1
chkImprimeArquivosGerados.Value = 1
chkImprimeTransacao.Value = 0
chkImprimeDivisaoEntradas.Value = 1

Call Lercmd_Click
 

Exit Sub

'On Error GoTo trataerr
trataerr:
Call TrataErro(App.title, Error, "Form_Load")

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo trataerr
Call ImprimeRelDel(Filename)

Exit Sub

trataerr:
Call TrataErro(App.title, Me.Name, "FormUload")


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
Dim soma As Double



ReDim extra(6)
soma = Val(txtValor) + Val(txtValorMan)

extra(0) = Text1
extra(1) = "Quantidade Automaticas : " + txtTot + " => Zeradas: " + TxtZero + " | Com Valor: " + txtVal
extra(2) = "Valor Automaticas: " + FormatCurrency(Val(txtValor))
extra(3) = "Quantidade Manuais : " + txtTotMan + "  => Zeradas: " + TxtZeroMan + " | Com Valor: " + txtValMan
extra(4) = "Valor Manuais: " + FormatCurrency(Val(txtValorMan))
extra(5) = "Valor Total : " + FormatCurrency(soma)

Filename = "TRN_" & gsEst_Codigo & "_" & TxtAno + TxtMes + TxtDia & "_" + Format(Now, "YYYYMMDDHHMMSS") & ".HTML"
'Filename = "Transacao" & Format(Date, "YYYYMMDD") & ".html"
Call ImprimeHeader(Filename, "Controle de Transações")
Call ImprimeExtra(Filename, extra)
ReDim extra(1)

If chkImprimeDivisaoValores Then
   extra(0) = "Quantidade por Tarifas"
   Call ImprimeExtra(Filename, extra)
   Call Imprimegrid(Filename, Grid2)
End If

If chkImprimeArquivosGerados Then
    extra(0) = "Controle de Arquivos"
    Call ImprimeExtra(Filename, extra)
    Call Imprimegrid(Filename, GrdFiles)
End If

If chkImprimeTransacao Then
    extra(0) = "Transações do Movimento"
    Call ImprimeExtra(Filename, extra)
    Call Imprimegrid(Filename, Grid1)
End If

If chkImprimeTransacaoManuais Then
    extra(0) = "Transações Manuais do Movimento"
    Call ImprimeExtra(Filename, extra)
    Call Imprimegrid(Filename, Grid1Man)
End If


' JAHARA
If chkImprimeDivisaoEntradas Then
    extra(0) = "Qauntidade por Entrada"
    Call ImprimeExtra(Filename, extra)
    Call Imprimegrid(Filename, GridEntrada)
End If
' JAHARA

Call ImprimeFooter(Filename, "Impresso em : " + Format(Now, "DD/MM/YY HH:MM:SS"))
Call ImprimeRel(Filename)

Filename = "TRN_" & gsEst_Codigo & "_" & TxtAno + TxtMes + TxtDia & "_" + Format(Now, "YYYYMMDDHHMMSS") & ".csv"
Call ArqGrid(Filename, Grid1)

End Sub

Private Sub Lercmd_Click()
On Error GoTo trataerr
Dim frase As String
Dim rs As New Recordset
Dim fraseaux As String
Dim aux As Boolean
rs.CursorType = adOpenStatic

Text1 = "Filtrado pelo Dia " + TxtDia + "/" + TxtMes + "/" + TxtAno

Lercmd.Enabled = False

If TxtDia = "00" Then
    aux = True
    fraseaux = " tsdatamovimento is null and "
Else
    aux = False
    fraseaux = " tsdatamovimento = '" & TxtAno + TxtMes + TxtDia & "' and "
End If

'TXTVALOR E TXTVAL AUTOMATICAS
frase = ""
frase = frase & "select count(*),dbo.fpoev2(sum(ivalor))"
frase = frase & "       from tb_transacao ta"
frase = frase & "       where  "
frase = frase & fraseaux
frase = frase & "       ivalor > 0 and "
frase = frase & "       istsaida =1 and "
frase = frase & "       1 = 1 "
Set rs = dbApp.Execute(frase)
txtVal = rs(0)
If rs(0) <> 0 Then
    txtValor = rs(1)
Else
    txtValor = 0
End If

'TXTVALOR E TXTVAL MANUAIS
frase = ""
frase = frase & "select count(*),dbo.fpoev2(sum(ivalor))"
frase = frase & "       from tb_transacao ta"
frase = frase & "       where  "
frase = frase & fraseaux
frase = frase & "       ivalor > 0 and "
frase = frase & "       istsaida =0 and "
frase = frase & "       1 = 1 "
Set rs = dbApp.Execute(frase)
txtValMan = rs(0)
If rs(0) <> 0 Then
    txtValorMan = rs(1)
Else
    txtValorMan = 0
End If


'TXTZERO E TXTTOT AUTOMATICAS
frase = ""
frase = frase & "select count(*) "
frase = frase & "       from tb_transacao ta"
frase = frase & "       where  "
frase = frase & fraseaux
frase = frase & "       ivalor = 0 and "
frase = frase & "       istsaida =1 and "
frase = frase & "       1 = 1 "
Set rs = dbApp.Execute(frase)
TxtZero = rs(0)
txtTot = Val(TxtZero) + Val(txtVal)


'TXTZERO E TXTTOT MANUAIS
frase = ""
frase = frase & "select count(*) "
frase = frase & "       from tb_transacao ta"
frase = frase & "       where  "
frase = frase & fraseaux
frase = frase & "       ivalor = 0 and "
frase = frase & "       istsaida =0 and "
frase = frase & "       1 = 1 "
Set rs = dbApp.Execute(frase)
TxtZeroMan = rs(0)
txtTotMan = Val(TxtZeroMan) + Val(txtValMan)


'GRDFILES AUTOMATICAS
frase = ""
frase = frase & "select "
If aux Then
frase = frase & " 'ND', "
Else
frase = frase & " NroFile,"
End If
frase = frase & " DtOper,"
frase = frase & " QtTotFile,"
frase = frase & " ValFile,"
frase = frase & " Tipo"
frase = frase & "       From"
frase = frase & "      (select isnull(lseqfile,0) as NroFile,"
frase = frase & "       convert(char,tsdataoperacao,3) as DtOper,"
frase = frase & "       count(*) as QtTotFile,"
frase = frase & "       dbo.fpoev2(Sum(ta.ivalor)) As ValFile,"
frase = frase & " CASE istsaida WHEN 1 THEN 'A' Else 'M' END AS 'Tipo'"
frase = frase & "       from tb_transacao ta"
frase = frase & "       where  "
frase = frase & fraseaux
frase = frase & "       ivalor > 0 and "
frase = frase & "       1 = 1 "
frase = frase & "       GROUP BY istsaida,isnull(lseqfile,0),convert(char,tsdataoperacao,3) ) as TBTEMP"
frase = frase & " order by NroFile,dtoper "
Set rs = dbApp.Execute(frase)

Set GrdFiles.DataSource = rs
GrdFiles.TextMatrix(0, 0) = "Sequencial"
GrdFiles.TextMatrix(0, 1) = "Dt Operação "
GrdFiles.TextMatrix(0, 2) = "Qtde      "
GrdFiles.TextMatrix(0, 3) = "Valor(R$)  "
GrdFiles.TextMatrix(0, 4) = "Tipo "
GrdFiles.ColAlignment = flexAlignLeftCenter
GrdFiles.ColAlignment(0) = flexAlignRightCenter
GrdFiles.ColAlignment(2) = flexAlignRightCenter
GrdFiles.ColAlignment(3) = flexAlignRightCenter
GrdFiles.ColAlignment(4) = flexAlignRightCenter
Call FormataGridx(GrdFiles, rs)
GrdFiles.ColWidth(0) = 1200
GrdFiles.ColWidth(2) = 1000



frase = ""
frase = frase & " select"
frase = frase & " dbo.fpoev2(ivalor),"
frase = frase & " str(count(*),10),"
frase = frase & " dbo.fpoev2(sum(ivalor))"
frase = frase & " From tb_transacao"
frase = frase & " where "
frase = frase & fraseaux
frase = frase & " 1 = 1"
frase = frase & " group by dbo.fpoev2(ivalor) with rollup"
Set rs = dbApp.Execute(frase)
Set Grid2.DataSource = rs
Grid2.TextMatrix(0, 0) = "Tarifa       "
Grid2.TextMatrix(0, 1) = "Qtde         "
Grid2.TextMatrix(0, 2) = "Valor(R$)     "
Grid2.ColAlignment = flexAlignRightCenter
Call FormataGridx(Grid2, rs)
Grid2.ColWidth(0) = 1000
Grid2.ColWidth(1) = 800



'base de dados
frase = ""
frase = frase & " select"
frase = frase & " convert(char(8),tsdataoperacao,3) as 'Data Operacao',"
frase = frase & " cPlaca as Placa,"
frase = frase & " right('00000' + cast(iissuer as varchar(5)),5) + '-' + right('0000000000' + cast(ltag as varchar(10)),10) as Tag,"
frase = frase & " convert(char(16),convert(nvarchar(8),tsentrada,3)+ ' ' + convert(nvarchar(5),tsentrada,8) + '-' + "
frase = frase & " replace(replace(istentrada,'0','M'),'1','A')) as Entrada,"
frase = frase & " convert(char(16),convert(nvarchar(8),tssaida,3)+ ' ' + convert(nvarchar(5),tssaida,8) + '-' + "
frase = frase & " replace(replace(istsaida,'0','M'),'1','A')) as Saida,"
' frase = frase & " str(ivalor,8) as Valor,"
frase = frase & " dbo.fpoeV2(ivalor) as Valor,"
frase = frase & " (select cdescricao from tb_pista where ipista = iacesso) as EntPista,"
frase = frase & " (select cdescricao from tb_pista where ipista = isaida) as SaiPista,"
frase = frase & " str(lseqfile,6) as Seq,"
frase = frase & " str(lseqreg,6) as Reg"
frase = frase & " from tb_transacao ta"
frase = frase & " where "
frase = frase & fraseaux
frase = frase & "       istsaida =1 and "
frase = frase & " 1 = 1 "
frase = frase & " order by tsdataoperacao,lseqfile,lseqreg"
Set rs = dbApp.Execute(frase)
Grid1.Clear

Set Grid1.DataSource = rs

Grid1.TextMatrix(0, 0) = "DT Oper    "
Grid1.TextMatrix(0, 1) = "Placa      "
Grid1.TextMatrix(0, 2) = "Tag                 "
Grid1.TextMatrix(0, 3) = "Entrada           "
Grid1.TextMatrix(0, 4) = "Saida             "
Grid1.TextMatrix(0, 5) = "Valor     "
Grid1.TextMatrix(0, 6) = "Pista ENT    "
Grid1.TextMatrix(0, 7) = "Pista SAI    "
Grid1.TextMatrix(0, 8) = "Seq  "
Grid1.TextMatrix(0, 9) = "Reg  "

Grid1.ColAlignment = flexAlignLeftCenter
Grid1.ColAlignment(5) = flexAlignRightCenter
Call FormataGridx(Grid1, rs)

Grid1.Refresh




'GRID DE TRANSACOES MANUAIS
'base de dados
frase = ""
frase = frase & " select"
frase = frase & " convert(char(8),tsdataoperacao,3) as 'Data Operacao',"
frase = frase & " cPlaca as Placa,"
frase = frase & " right('00000' + cast(iissuer as varchar(5)),5) + '-' + right('0000000000' + cast(ltag as varchar(10)),10) as Tag,"
frase = frase & " convert(char(16),convert(nvarchar(8),tsentrada,3)+ ' ' + convert(nvarchar(5),tsentrada,8) + '-' + "
frase = frase & " replace(replace(istentrada,'0','M'),'1','A')) as Entrada,"
frase = frase & " convert(char(16),convert(nvarchar(8),tssaida,3)+ ' ' + convert(nvarchar(5),tssaida,8) + '-' + "
frase = frase & " replace(replace(istsaida,'0','M'),'1','A')) as Saida,"
' frase = frase & " str(ivalor,8) as Valor,"
frase = frase & " dbo.fpoeV2(ivalor) as Valor,"
frase = frase & " (select cdescricao from tb_pista where ipista = iacesso) as EntPista,"
frase = frase & " (select cdescricao from tb_pista where ipista = isaida) as SaiPista,"
frase = frase & " str(lseqfile,6) as Seq,"
frase = frase & " str(lseqreg,6) as Reg"
frase = frase & " from tb_transacao ta"
frase = frase & " where "
frase = frase & fraseaux
frase = frase & "       istsaida =0 and "
frase = frase & " 1 = 1 "
frase = frase & " order by tsdataoperacao,lseqfile,lseqreg"
Set rs = dbApp.Execute(frase)
Grid1Man.Clear

Set Grid1Man.DataSource = rs

Grid1Man.TextMatrix(0, 0) = "DT Oper    "
Grid1Man.TextMatrix(0, 1) = "Placa      "
Grid1Man.TextMatrix(0, 2) = "Tag                 "
Grid1Man.TextMatrix(0, 3) = "Entrada           "
Grid1Man.TextMatrix(0, 4) = "Saida             "
Grid1Man.TextMatrix(0, 5) = "Valor     "
Grid1Man.TextMatrix(0, 6) = "Pista ENT    "
Grid1Man.TextMatrix(0, 7) = "Pista SAI    "
Grid1Man.TextMatrix(0, 8) = "Seq  "
Grid1Man.TextMatrix(0, 9) = "Reg  "

Grid1Man.ColAlignment = flexAlignLeftCenter
Grid1Man.ColAlignment(5) = flexAlignRightCenter
Call FormataGridx(Grid1Man, rs)

Grid1Man.Refresh

' JAHARA - Implementar grid com contagem por ENTRADA

'base de dados
frase = ""
frase = frase & " select tp.cdescricao as EntPista, COUNT(*) quant ,dbo.fpoeV2(SUM(IValor)) total "
frase = frase & " from tb_transacao ta , tb_pista tp"
frase = frase & " where "
frase = frase & fraseaux
frase = frase & " 1 = 1  and ta.Iacesso = tp.IPista"
frase = frase & " group by tp.cdescricao   with rollup"
Set rs = dbApp.Execute(frase)
GridEntrada.Clear

Set GridEntrada.DataSource = rs

GridEntrada.TextMatrix(0, 0) = "Acesso                "
GridEntrada.TextMatrix(0, 1) = "Quant    "
GridEntrada.TextMatrix(0, 2) = "Valor (R$)     "

GridEntrada.ColAlignment = flexAlignLeftCenter
GridEntrada.ColAlignment(1) = flexAlignRightCenter
GridEntrada.ColAlignment(2) = flexAlignRightCenter

Call FormataGridx(GridEntrada, rs)

GridEntrada.Refresh

' JAHARA - Implementar grid com contagem por ENTRADA



Lercmd.Enabled = True

Exit Sub
'On Error GoTo trataerr
trataerr:
Call TrataErro(App.title, Me.Name, "LerCMD")

End Sub

Private Sub saircmd_Click()

Unload Me

End Sub

Private Sub TxtAno_Change()

Lercmd.Enabled = True

End Sub

Private Sub TxtDia_Change()

Lercmd.Enabled = True

End Sub

Private Sub TxtDia_LostFocus()

If Val(TxtDia) > 31 Or Val(TxtDia) < 0 Then TxtDia = "01"

TxtDia = Format(TxtDia, "00")

End Sub

Private Sub TxtMes_Change()

Lercmd.Enabled = True

End Sub

Private Sub TxtMes_LostFocus()

If Val(TxtMes) > 12 Or Val(TxtMes) < 1 Then TxtMes = "1"

TxtMes = Format(TxtMes, "00")

End Sub

