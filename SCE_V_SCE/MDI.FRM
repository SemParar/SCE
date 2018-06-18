VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDI 
   BackColor       =   &H8000000C&
   Caption         =   "Controle de Cadastros"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   1515
   ClientWidth     =   11880
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar sta_Barra_MDI 
      Align           =   2  'Align Bottom
      DragMode        =   1  'Automatic
      Height          =   465
      Left            =   0
      TabIndex        =   0
      Top             =   3690
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   820
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7048
            MinWidth        =   3529
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3545
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3545
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3545
            MinWidth        =   3528
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   240
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   1440
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgMDI 
      Left            =   120
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuArquivo_Logon 
         Caption         =   "&Logon"
      End
      Begin VB.Menu mnuArquivo_Sair 
         Caption         =   "Sai&r"
      End
   End
   Begin VB.Menu mnuProcessamento 
      Caption         =   "&Processamento"
      Visible         =   0   'False
      Begin VB.Menu mnuProcessamento_Cad_ListaNela 
         Caption         =   "Trata - Lista Nela (LN / LNT)"
      End
      Begin VB.Menu mnuProcessamento_Cad_Tag 
         Caption         =   "Trata - Tags (TG / TGT)"
      End
      Begin VB.Menu mnuProcessamento_file_TRN 
         Caption         =   "Trata - Financeiro (TRN/TRT/TRF)"
      End
      Begin VB.Menu mnuTag_Reenvio 
         Caption         =   "&Reenvio de Transacao"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "&Controles"
      Begin VB.Menu mnuRelatorios_TRN 
         Caption         =   "Transação"
      End
      Begin VB.Menu mnuRelatorios_Rej 
         Caption         =   "Rejeitadas"
      End
      Begin VB.Menu mnuRelatorios_TRF 
         Caption         =   "Financeiro"
      End
      Begin VB.Menu mneOperacao 
         Caption         =   "Operacao"
      End
   End
   Begin VB.Menu mnutag 
      Caption         =   "&Tag"
      Begin VB.Menu mnutagticket 
         Caption         =   "Ticket"
      End
      Begin VB.Menu mnuRelatorios_Pateo 
         Caption         =   "Veiculos no Patio"
      End
      Begin VB.Menu mnuTag_Pesquisar 
         Caption         =   "&Pesquisar"
      End
   End
   Begin VB.Menu mnuEstatisticas 
      Caption         =   "Estatisticas"
      Begin VB.Menu mnuInd 
         Caption         =   "Indice"
      End
      Begin VB.Menu mnuPerm 
         Caption         =   "Permanencia"
      End
      Begin VB.Menu mnuRelatorios_Periodo 
         Caption         =   "Relatorio Periodo"
      End
      Begin VB.Menu mnuRelatorios_Exportar 
         Caption         =   "Exportar Movimento"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuDesc 
      Caption         =   "Descontos"
      Begin VB.Menu mnuDescRel 
         Caption         =   "FNAC"
      End
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "&Config"
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Activate()

    mnuArquivo_Click
    sta_Barra_MDI.Panels(1).Text = gsUser
    sta_Barra_MDI.Panels(2).Text = "                                           "
    sta_Barra_MDI.Panels(3).Text = Date
    sta_Barra_MDI.Panels(4).Text = Time
    sta_Barra_MDI.Panels(5).Text = "E: " + Format(gsEst_Codigo) + " - V: " + Format(app.Major) + "." + Format(app.Minor) + "." + Format(app.Revision)

 End Sub

Private Sub MDIForm_Load()
    If gsUser <> "" Then
        mnuArquivo_Logon.Caption = "Logoff..."
    End If
' controlar menu
    If gbytNivel_Acesso_Usuario <> gintNIVEL_ADMINISTRADOR Then
        mnuDesc.Visible = False
    Else
        mnuDesc.Visible = True
    End If
    
    Me.Caption = " SCE - Sistema de Controle de Estacionamento - " + gsEst_Nome
    
    rsGeral.CursorType = adOpenStatic
    Set rsGeral = Nothing

    sta_Barra_MDI.Panels(3).Text = Date
  
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Dim i As Integer
Dim msg
' Set the message text.

If Not gbTransferStatus Then
    msg = "Você Realmente esta Saindo da Aplicação - SCE "
    '   If MsgBoxService(msg, vbQuestion + vbYesNo, Me.Caption) = vbNo Then Cancel = True
    Call MsgBoxService(msg, vbQuestion + vbYesNo, Me.Caption)
    Exit Sub
End If
MsgBoxService "O SCE está executando uma transferência de arquivos... Favor Aguardar", vbOKOnly, Me.Caption

Cancel = True

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

MySetPar "SCE", "SISTEMA", "SCELN", 0
MySetPar "SCE", "SISTEMA", "SCETG", 0
MySetPar "SCE", "SISTEMA", "SCETR", 0

End Sub

Private Sub mneOperacao_Click()
On Error GoTo trataerr

frmRelOperacional.Show

Exit Sub

trataerr:
Call TrataErro(app.title, Me.Name, "mnuRelatorio_Operacional_Click(")

End Sub

Private Sub mnuArquivo_Click()
    If Trim(gsUser) <> "" Then
        mnuArquivo_Logon.Caption = "Log &Off ... "
        If gbytNivel_Acesso_Usuario = gintNIVEL_ADMINISTRADOR Then
            'x
        End If
    Else
        mnuArquivo_Logon.Caption = "Log O&n"
    End If
End Sub

Private Sub mnuArquivo_Configuracao_Impressao_Click()
    dlgMDI.ShowPrinter
End Sub


Private Sub mnuArquivo_Logon_Click()
    If Trim(gsUser) <> "" Then
        gbytNivel_Acesso_Usuario = 0
        MsgBoxService "Logoff de  " & gsUser & "  efetuado com sucesso ! ", vbOKOnly + vbInformation, "Aviso"
        gsUser = ""
        DoEvents
    Else
        frmLogin.Show vbModal
    End If
    MDIForm_Activate
End Sub

Private Sub mnuArquivo_Sair_Click()
    dbApp.Close
    Unload Me
    End
End Sub


Private Sub mnuConfig_Click()
frmConfig.Show
End Sub

Private Sub mnuDescRel_Click()

If gbytNivel_Acesso_Usuario = gintNIVEL_ADMINISTRADOR Then
    frmRelDesc.Show
Else
    MsgBoxService "Somente Administrator Tem Acesso ao Recurso"
End If
    
End Sub

Private Sub mnuInd_Click()
On Error GoTo trataerr

frmRelInd.Show

Exit Sub

trataerr:
Call TrataErro(app.title, Me.Name, "mnuInd_Click")
End Sub

Private Sub mnuPerm_Click()

On Error GoTo trataerr
frmRelPerm.Show
Exit Sub

trataerr:
Call TrataErro(app.title, Me.Name, "mnuPerm_Click")

End Sub

Private Sub mnuProcessamento_Cad_ListaNela_Click()
On Error GoTo trataerr

'If gssceln < 2 Then
'    MySetPar gsPathIniFile, "SISTEMA", "SCELN", 1
'    Shell gsPath + "sceln.exe"
'End If

Exit Sub

trataerr:
Call TrataErro(app.title, Me.Name, "mnuProcessamento_Cad_ListaNela_Click")

End Sub

Private Sub mnuProcessamento_Cad_Tag_Click()
On Error GoTo trataerr

'If gsscetg < 2 Then
'    MySetPar gsPathIniFile, "SISTEMA", "SCETG", 1
'    Shell gsPath + "scetg.exe"
'End If

Exit Sub

trataerr:
Call TrataErro(app.title, Me.Name, "mnuProcessamento_Cad_Tag_Click")

End Sub

Private Sub mnuProcessamento_file_TRN_Click()
On Error GoTo trataerr

'If gsscetr < 2 Then
'    MySetPar gsPathIniFile, "SISTEMA", "SCETR", 1
'    Shell gsPath + "scetr.exe"
'End If

Exit Sub

trataerr:
Call TrataErro(app.title, Me.Name, "mnuProcessamento_file_TRN_Click")


End Sub


Private Sub mnuRelatorios_Exportar_Click()
On Error GoTo trataerr

frmRelExportar.Show

Exit Sub

trataerr:
Call TrataErro(app.title, Me.Name, "mnuProcessamento_file_TRN_Click")


End Sub

Private Sub mnuRelatorios_Pateo_Click()
On Error GoTo trataerr

frmRelPat.Show
Exit Sub

trataerr:
Call TrataErro(app.title, Me.Name, "mnuProcessamento_file_TRN_Click")


End Sub

Private Sub mnuRelatorios_Periodo_Click()

On Error GoTo trataerr

frmRelDia.Show
Exit Sub

trataerr:
Call TrataErro(app.title, Me.Name, "mnuProcessamento_file_TRN_Click")


End Sub

Private Sub mnuRelatorios_Rej_Click()
On Error GoTo trataerr

frmRelRej.Show
Exit Sub

trataerr:
Call TrataErro(app.title, Me.Name, "mnuProcessamento_file_TRN_Click")


End Sub


Private Sub mnuRelatorios_TRF_Click()

On Error GoTo trataerr
 
frmRelFin.Show
Exit Sub

trataerr:
Call TrataErro(app.title, Me.Name, "mnuProcessamento_file_TRN_Click")


End Sub

Private Sub mnuRelatorios_TRN_Click()
On Error GoTo trataerr

Me.MousePointer = vbHourglass
frmRelMov.Show
Me.MousePointer = vbDefault
Exit Sub

trataerr:
Call TrataErro(app.title, Me.Name, "mnuProcessamento_file_TRN_Click")


End Sub


Private Sub mnuTag_Pesquisar_Click()
On Error GoTo trataerr

frmTagPes.Show

Exit Sub
trataerr:
    Call TrataErro(app.title, Me.Name, "mnuTag_Pesquisar_Click")

End Sub


Private Sub mnutagticket_Click()
On Error GoTo trataerr

frmTagTicket.Show

Exit Sub
trataerr:
    Call TrataErro(app.title, Me.Name, "mnuTag_Pesquisar_Click")

End Sub

Private Sub Timer1_Timer()
Dim i As Long
Dim ponto As String

On Error GoTo timererro

sta_Barra_MDI.Panels(3).Text = Date
sta_Barra_MDI.Panels(4).Text = Time

Dim teste As String

If Not gbTransferStatus And giTransfere <= 0 Then

    If Now > CVDate(gsNextClear) Then
        sta_Barra_MDI.Panels(2).Text = "Limpando Arquivos"
        ponto = "LIMPAARQUIVOS"
        gbTRNOK = False
        Call LimparArquivos
        If gbTRNOK Then
            gsNextClear = CVDate(gsNextClear) + 1
            MySetPar app.title, "SISTEMA", "NextClear", Format(gsNextClear, "DD/MM/YY hh:mm:ss")
        End If
        sta_Barra_MDI.Panels(2).Text = ""
    End If
    
    Call mnuProcessamento_Cad_ListaNela_Click
    DoEvents

    Call mnuProcessamento_Cad_Tag_Click
    DoEvents

    Call mnuProcessamento_file_TRN_Click
    DoEvents
    
    giTransfere = MyGetPar(gsPathIniFile, "SISTEMA", "TempoBusca", 60)
    gbTransferStatus = False

Else
    If gbTransferStatus Or giTransfere < 1 Then
         giTransfere = MyGetPar(gsPathIniFile, "SISTEMA", "TempoBusca", 60)
    End If
    giTransfere = giTransfere - 1
End If


Exit Sub

timererro:
    'Call MyMsgBox("Erro no Timer ---- " & ponto, vbOK, "TIMER SCE", "")
    Call LogErro("Timer", "Erro no Timer ---- " & ponto)
    giTransfere = 10
    End
End Sub



