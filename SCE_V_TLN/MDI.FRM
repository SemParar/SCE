VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDI 
   BackColor       =   &H8000000C&
   Caption         =   "SCE - Atualiza LN"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   2100
   ClientWidth     =   13035
   LinkTopic       =   "MDIForm1"
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   1440
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar sta_Barra_MDI 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   6945
      Width           =   13035
      _ExtentX        =   22992
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   14491
            MinWidth        =   7056
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "16/8/2017"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Object.Tag             =   ""
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
   Begin MSComDlg.CommonDialog dlgMDI 
      Left            =   120
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   240
   End
   Begin VB.Menu mnuArquivo_Sair 
      Caption         =   "Sai&r"
   End
   Begin VB.Menu mnuProcessamento 
      Caption         =   "&Processamento"
      Begin VB.Menu mnuProcessamento_Cad_ListaNela 
         Caption         =   "Cadastro  - Lista Nela (LN / LNT)"
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
    sta_Barra_MDI.Panels(1).Text = gssceln + " " + gsUser
    sta_Barra_MDI.Panels(3).Text = Date
    sta_Barra_MDI.Panels(4).Text = Time
    sta_Barra_MDI.Panels(5).Text = "E: " + Format(gsEst_Codigo) + " - V: " + Format(app.Major) + "." + Format(app.Minor) + "." + Format(app.Revision)


End Sub

Private Sub MDIForm_Load()

Me.Top = 0
Me.Left = 0

Me.Caption = " SCE - Lista Nela - " + gsEst_Nome
MDI.sta_Barra_MDI.Panels(4).Text = Time

If UCase(gsListas) = "EST" Then
        frmCadEst.Show
        Else
        frmCadLN.Show
    End If



End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim i As Integer
Dim msg
' Set the message text.

If Not gbTransferStatus Then
    msg = "Você Realmente esta Saindo da Aplicação - SCELN "
    'If MsgBoxService(msg, vbQuestion + vbYesNo, Me.Caption) = vbNo Then Cancel = True
    Call MsgBoxService(msg, vbQuestion + vbYesNo, Me.Caption)
    Exit Sub
End If
MsgBoxService "O SCELN está executando uma transferência de arquivos... Favor Aguardar", vbOKOnly, Me.Caption

Cancel = True

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Dim i As Integer
Dim msg

dbApp.Close

End Sub

Private Sub mnuArquivo_Sair_Click()
    Unload Me
    End
End Sub

Private Sub mnuProcessamento_Cad_ListaNela_Click()

frmCadLN.Show

End Sub

Private Sub Timer1_Timer()
Dim i As Long
Dim ponto As String
Dim frase As String
Dim rs As New Recordset
rs.CursorType = adOpenStatic

On Error GoTo timererro

MDI.sta_Barra_MDI.Panels(3).Text = Date
MDI.sta_Barra_MDI.Panels(4).Text = Time


If gsListas <> "EST" Then

    If Not gbTransferStatus And giTransfere <= 0 Then
        
        gbTransferStatus = True
        Call LerGlobais
         
        ponto = "CGMPRECEBELN"
        Call CGMPRecebeLN
        DoEvents
        
        'Limpar cadfileCtl
        frase = "delete " + gsPath_DBCAD + ".dbo.TB_CADnelaCtl where lseqfile is null"
        Set rs = dbApp.Execute(frase)
        MDI.sta_Barra_MDI.Panels(2).Text = "Atualizando Lista Nela"
        ponto = "TRATAFILE LN"
        Call TrataFileLN
        MDI.sta_Barra_MDI.Panels(2).Text = ""
        giTransfere = 10
        gbTransferStatus = False
        
    Else
       If gbTransferStatus Or giTransfere < 1 Then
          'retentar daqui a 1 minuto
          giTransfere = gsTempoBusca
          
       End If
       giTransfere = giTransfere - 1
    End If
End If


Exit Sub

timererro:
    Call TrataErro(app.title, Error, "TimerMDI_SCELN")
    giTransfere = gsTempoBusca
    gbTransferStatus = False
    
End Sub
