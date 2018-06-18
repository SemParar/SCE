VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTagReenvio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reenvio de Tag"
   ClientHeight    =   6870
   ClientLeft      =   30
   ClientTop       =   855
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   12195
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   855
      Left            =   120
      TabIndex        =   34
      Top             =   5760
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   1508
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtReg 
      Height          =   615
      Left            =   3360
      TabIndex        =   33
      Text            =   "Reg"
      Top             =   4680
      Width           =   3015
   End
   Begin VB.TextBox txtTrn 
      Height          =   615
      Left            =   480
      TabIndex        =   32
      Text            =   "TRN"
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Frame frmValor 
      Caption         =   "Valor"
      Height          =   1575
      Left            =   10920
      TabIndex        =   30
      Top             =   1440
      Width           =   2535
      Begin VB.Label lblValor 
         Caption         =   "Valor"
         Height          =   615
         Left            =   480
         TabIndex        =   31
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "SAIDA"
      Height          =   1575
      Left            =   5160
      TabIndex        =   19
      Top             =   1440
      Width           =   5535
      Begin VB.Frame FraDarSaida 
         Caption         =   "Confirme Saida"
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
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   3492
         Begin MSComCtl2.DTPicker DTPSaiDia 
            Height          =   492
            Left            =   120
            TabIndex        =   24
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
            TabIndex        =   25
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
         Left            =   3960
         TabIndex        =   29
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblIpistaSai 
         Caption         =   "IPista"
         Height          =   375
         Left            =   4200
         TabIndex        =   27
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ENTRADA"
      Height          =   1575
      Left            =   120
      TabIndex        =   18
      Top             =   1440
      Width           =   4815
      Begin VB.Frame FraDarEntrada 
         Caption         =   "Confirme  Entrada"
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
         TabIndex        =   20
         Top             =   360
         Width           =   3492
         Begin MSComCtl2.DTPicker DTPEntHora 
            Height          =   492
            Left            =   1920
            TabIndex        =   21
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
            TabIndex        =   22
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
      Begin VB.Label Label4 
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
         Left            =   3840
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LblIpistaEnt 
         Caption         =   "IPista"
         Height          =   375
         Left            =   3840
         TabIndex        =   26
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pesquisa de STATUS"
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   13335
      Begin VB.Label lblPatio 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPatio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   17
         Top             =   360
         Width           =   855
      End
      Begin VB.Label LblTRN 
         Alignment       =   1  'Right Justify
         Caption         =   "lblTRN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12240
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Registro :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   3
         Left            =   10920
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblNela 
         Alignment       =   1  'Right Justify
         Caption         =   "lblNela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8640
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.Label LBLCAD 
         Alignment       =   1  'Right Justify
         Caption         =   "lblCAd"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Pátio :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Cad Nela :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   1
         Left            =   7200
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Cad Tag :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   0
         Left            =   3600
         TabIndex        =   10
         Top             =   360
         Width           =   1215
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
      Left            =   11040
      Picture         =   "frmTagReenvio.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
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
      Left            =   12360
      Picture         =   "frmTagReenvio.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   972
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
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   10575
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
         Left            =   1560
         TabIndex        =   3
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
         Left            =   4440
         TabIndex        =   2
         Text            =   "tag"
         Top             =   360
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
         Left            =   8400
         TabIndex        =   1
         Text            =   "Placa"
         Top             =   360
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
         Left            =   360
         TabIndex        =   8
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
         Left            =   7200
         TabIndex        =   7
         Top             =   360
         Width           =   735
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
         Left            =   3720
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmTagReenvio"
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
    MsgBox "ESCOLHA UMA PISTA", vbOKOnly, "MENSAGEM"
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

Call cmdLimpa_Click
TxtTag = Mid(Auxtag, 6, 10)
TxtEmissor = Mid(Auxtag, 1, 5)
TxtPlaca = Auxplaca
Call cmdPes_Click

End Sub

Private Sub cmdDarSaida_Click()
Dim aux As String
Dim Auxvalor  As String
Dim rsaux As Recordset

If cmbSaida = "" Then
    MsgBox "ESCOLHA UMA PISTA", vbOKOnly, "MENSAGEM"
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
    If rsaux.EOF = True Or rsaux.BOF = True Then
        MsgBox "Tag nao esta no estacionamento "
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

Filename = "tag_" & gsEst_Codigo & "_" & TxtAno + TxtMes + TxtDia & "_" + Format(Now, "YYYYMMDDHHMMSS") & ".HTML"
'Filename = "TAG_" & TxtTag & "_" & Format(Date, "YYYYMMDD") & ".html"
Call ImprimeHeader(Filename, "Historico de Tag ")
Call Imprimegrid(Filename, Grid1)
Call ImprimeExtra(Filename, extra)
Call ImprimeFooter(Filename, "Impresso em : " + Format(Now, "DD/MM/YY HH:MM:SS"))
Call ImprimeRel(Filename)
Call ImprimeRelDel(Filename)

End Sub

Public Sub cmdLimpa_Click()

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

cmdimprime.Enabled = False


End Sub

Public Sub cmdPes_Click()
Dim aux As Integer
Dim PesqPlaca As Boolean
Dim NaoAchei As Boolean


NaoAchei = True

Auxtag = ""
TxtTag = UCase(LTrim(RTrim(TxtTag)))
TxtPlaca = UCase(LTrim(RTrim(TxtPlaca)))
TxtEmissor = UCase(LTrim(RTrim(TxtEmissor)))
If TxtEmissor = "" Then TxtEmissor = "290"

If TxtTag <> "" Then
    If Val(TxtTag) > 2000000000 Then
        MsgBox "Numero de TAG Invalido - Acima do Limite"
        Call cmdLimpa_Click
        Exit Sub
    End If
    TxtPlaca = ""
    PesqPlaca = False
Else
    'elimina qualquer pesquisa com placa XXX / YYY
    If TxtPlaca = "XXX9999" Or TxtPlaca = "YYY9999" Then
        MsgBox "Placa : " + TxtPlaca + " - Não pode ser pesquisada", vbOKOnly, "Alerta"
        Call cmdLimpa_Click
        Exit Sub
    End If
    If TxtPlaca = "" Then
        MsgBox "DIGITE = Numero de TAG ou Placa", vbOKOnly, "Alerta"
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
LBLCAD.Caption = ""
If Not PesqPlaca Then
    frase = ""
    frase = frase + " select top 1 iissuer,ltag,cplaca from tb_cadtag "
    frase = frase + " where iissuer = '" + TxtEmissor + "' and ltag = '" + TxtTag + "'"
    Set rs = dbApp.Execute(frase)
    If Not rs.EOF Then
        If NaoAchei Then Call Achei(rs(0), rs(1), rs(2), NaoAchei)
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
    If rs(0) > 1 Then
        frase = ""
        frase = frase & " select t.iissuer, t.ltag,"
        frase = frase & " (select 'Lista Nela nr.' + cast(lseqfile as varchar(10)) from tb_cadnela l where l.ltag = t.ltag and l.iissuer = t.iissuer)"
        frase = frase & " from tb_cadtag t"
        frase = frase & " where cplaca = '" + TxtPlaca + "'"
        frase = frase & " order by t.iissuer,t.ltag "
        Set rs = dbApp.Execute(frase)
        rs.MoveFirst
        Do While Not rs.EOF
            msg = msg & rs(0) & " - " & rs(1) & " : " & rs(2) & vbCr
            rs.MoveNext
        Loop
        msg = "Placa : " + TxtPlaca + " - Tem mais que um tag no cadastro. Digite um Tag" & vbCr & vbCr & msg
        MsgBox msg, vbOKOnly, "Alerta"
        LBLCAD.Caption = "NAO"
   End If
   If rs(0) = 1 Then
        frase = " select top 1 iissuer,ltag,cplaca from tb_cadtag where cplaca = '" + TxtPlaca + "'"
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
Else
    frase = ""
    frase = frase + " select top 1 iissuer,ltag,cplaca from tb_cadtag "
    frase = frase + " where  cplaca = '" + TxtPlaca + "'"
    frase = frase + " and cst not in (" + gsCodNelaLivre + ")"
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
    MsgBox "Digite uma Placa ou Tag", vbOKOnly, "Alerta"
    Call cmdLimpa_Click
    Exit Sub
End If

'frase = ""
'frase = frase & " ("
'frase = frase & " select "
'frase = frase & " cplaca as Placa,"
'frase = frase & " cast(iissuer as char(5)) + cast(ltag as char(10)) as Tag,"
'frase = frase & " convert(char(17),convert(varchar(8),tsentrada,3)+ ' ' + convert(varchar(8),tsentrada,8)) as Entrada,"
'frase = frase & " (select cdescricao from tb_pista where ipista = iacesso) as PE,"
'frase = frase & " convert(char(17),replace(convert(varchar(8),tssaida,3)+ ' '+ convert(varchar(8),tssaida,8),'01/01/80 00:00:00',' ')) as Saida,"
'frase = frase & " (select cdescricao from tb_pista where ipista = isaida) as PS,"
'frase = frase & " right('            ' + cast(ivalor as char(10)),10) as Valor,"
'frase = frase & " case istentrada when 0 then 'Ent Man - ' when 1 then 'Ent Aut - ' end +"
'frase = frase & " case istsaida when 0 then 'Sai Man' when 1 then 'Sai Aut' end as Obs,"
'frase = frase & " convert(varchar(30),tsentrada,120)+' - 0 - TRN' as Tab"
'frase = frase & " From tb_transacao"
'frase = frase & " where iissuer='" + TxtEmissor + "' and ltag = '" + TxtTag + "')"
'frase = frase & " Union"
'frase = frase & " ("
'frase = frase & " select "
'frase = frase & " cplaca as Placa,"
'frase = frase & " cast(iissuer as char(5)) + cast(ltag as char(10)) as Tag,"
'frase = frase & " convert(char(17),convert(varchar(8),tsingresso,3)+ ' ' + convert(varchar(8),tsingresso,8)) as Entrada,"
'frase = frase & " (select cdescricao from tb_pista where ipista = ientpista) as PE,"
'frase = frase & " convert(char(17),replace(convert(varchar(8),tsegresso,3)+ ' '+ convert(varchar(8),tsegresso,8),'01/01/80 00:00:00',' ')) as Saida,"
'frase = frase & " (select cdescricao from tb_pista where ipista = isaipista) as PS,"
'frase = frase & " right('            ' + cast(lvalor as char(10)),10) as Valor,"
''frase = frase & " sztexto as Obs,"
'frase = frase & " case ientauto when 0 then 'Ent Man - ' when 1 then 'Ent Aut - ' end as Obs, "
''frase = frase & " case isaiauto when 0 then 'Sai Man' when 1 then 'Sai Aut' end as Obs,"
'frase = frase & " convert(varchar(30),tsingresso,120)+' - 2 - PR' as Tab"
'frase = frase & " From tb_praia"
'frase = frase & " where iissuer='" + TxtEmissor + "' and ltag = '" + TxtTag + "'"
'frase = frase & "  )"
'frase = frase & " Union"
'frase = frase & " ("
'frase = frase & " select"
'frase = frase & " cplaca as placa,"
'frase = frase & " cast(iparam as char(5)) + cast(lparam as char(10)) as tag,"
'frase = frase & " '' as entrada,"
'frase = frase & " (select cdescricao from tb_pista where ipista = lturno)  as PE,"
'frase = frase & " convert(char(17),convert(varchar(8),tsdatahora,3)+ ' ' + convert(varchar(8),tsdatahora,8)) as saida,"
'frase = frase & " ' ' as PS,"
'frase = frase & " ' ' as Valor,"
'frase = frase & " sztexto as obs,"
'frase = frase & " convert(varchar(30),tsdatahora,120)+' - 1 - EV' as Tab"
'frase = frase & " From tb_eventos"
'frase = frase & " Where"
'frase = frase & " iparam = '" + TxtEmissor + "' and"
'frase = frase & " lparam = '" + TxtTag + "' and"
'frase = frase & " icodigo = 40 and"
'frase = frase & " tsdatahora > convert(varchar(8),getdate()-90,112)"
'frase = frase & " )"
'frase = frase & "  order by Tab desc"

frase = "exec pr_PesqTag " + TxtEmissor + "," + TxtTag

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

FraDarSaida.Enabled = False
FraDarSaida.Enabled = False
   
   
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


Me.Top = 10
Me.Left = 10

Call cmdLimpa_Click

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

Private Sub LblIpista_Click()

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

