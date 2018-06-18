VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDI 
   BackColor       =   &H8000000A&
   Caption         =   "SCE - TRN"
   ClientHeight    =   7980
   ClientLeft      =   105
   ClientTop       =   735
   ClientWidth     =   11880
   LinkTopic       =   "MDIForm1"
   Begin MSComctlLib.StatusBar sta_Barra_MDI 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   0
      Top             =   7515
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   820
      ShowTips        =   0   'False
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
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuArquivo_Logon 
         Caption         =   "&Logon"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuArquivo_Sair 
         Caption         =   "Sai&r"
      End
   End
   Begin VB.Menu mnuProcessamento 
      Caption         =   "&Processamento"
      Begin VB.Menu mnuProcessamento_file_TRN 
         Caption         =   "Arquivos  (TRN/TRT/TRF)"
      End
      Begin VB.Menu mnurelmov 
         Caption         =   "Transação"
      End
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
    sta_Barra_MDI.Panels(4).Text = Date
    sta_Barra_MDI.Panels(3).Text = Time
    sta_Barra_MDI.Panels(5).Text = "E: " + Format(gsEst_Codigo) + " - V: " + Format(app.Major) + "." + Format(app.Minor) + "." + Format(app.Revision)
End Sub

Private Sub MDIForm_Load()

If gsUser <> "" Then
   mnuArquivo_Logon.Caption = "Logoff..."
End If

Me.Top = 10
Me.Left = 10

Me.Caption = " SCE - Transacao - " + gsEst_Nome

rsGeral.CursorType = adOpenStatic
Set rsGeral = Nothing

sta_Barra_MDI.Panels(4).Text = Date
Call Atualiza_NextTRN

frmCadRT.Show

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Dim i As Integer
Dim msg
' Set the message text.

If Not gbTransferStatus Then
    msg = "Você Realmente esta Saindo da Aplicação - SCETR "
   ' If MsgBoxService(msg, vbQuestion + vbYesNo, Me.Caption) = vbNo Then Cancel = True
     Call MsgBoxService(msg, vbQuestion + vbYes, Me.Caption)
    Exit Sub
End If
MsgBoxService "O SCETR está executando uma transferência de arquivos... Favor Aguardar", vbOKOnly, Me.Caption

Cancel = True

End Sub

Private Sub MDIForm_Resize()
'    lblData.Left = MDI.Width - (lblData.Width + 150)
'    lblHora.Left = MDI.Width - (lblData.Width + lblHora.Width + 150)
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
Private Sub mnuProcessamento_file_TRN_Click()

frmCadRT.Show

End Sub

Private Sub mnurelmov_Click()

frmRelMov.Show

End Sub

Private Sub Timer1_Timer()
Dim i As Long
Dim ponto As String

On Error GoTo timererro

MDI.sta_Barra_MDI.Panels(3).Text = Date
MDI.sta_Barra_MDI.Panels(4).Text = Time

LerGlobais


If Not gbTransferStatus And giTransfere <= 0 Then
    giTransfere = 20
    DoEvents
    gbTransferStatus = True
    ponto = "INICIO"
    If Now > CVDate(gsNextTrn) Then
        sta_Barra_MDI.Panels(2).Text = "Criando TRN : " + gsNextTrnNr
        ponto = "CALLTRATATRN"
        gbTRNOK = False
        Call TrataTRN
        
        'Gerar TRNAutoExpresso se o parametro estiver setada para 1
        If gsGerarTrnAE = 1 Then
            Call TrataTRNAutoExpresso
        End If
        
        sta_Barra_MDI.Panels(2).Text = ""
    End If
    sta_Barra_MDI.Panels(2).Text = "Tranferindo Arquivo"
    
    ponto = "CGMPRECEBETR"
    Call CGMPRecebeTR
    DoEvents
    
    ponto = "TrataFileGZ"
    Call TrataFileGZ
    DoEvents
    
    ponto = "CGMPENVIA"
    Call CGMPEnvia
    DoEvents
    
    sta_Barra_MDI.Panels(2).Text = "Atualizando TRT e TRF"
    ponto = "TRATAFILERT"
    Call TrataFileRT
    DoEvents
    
    sta_Barra_MDI.Panels(2).Text = ""
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
    MsgBoxService app.title & Error & "TimerMDI_SCETR"
    giTransfere = 10
    gbTransferStatus = False
End Sub


