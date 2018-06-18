VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDI 
   BackColor       =   &H8000000C&
   Caption         =   "SCE Atualiza TAG"
   ClientHeight    =   11190
   ClientLeft      =   3975
   ClientTop       =   1665
   ClientWidth     =   11880
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar sta_Barra_MDI 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   10860
      Width           =   11880
      _ExtentX        =   20955
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
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   12454
            MinWidth        =   7056
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   120
      Top             =   240
   End
   Begin VB.Menu mnuArquivo_Sair 
      Caption         =   "Sai&r"
   End
   Begin VB.Menu mnuProcessamento 
      Caption         =   "&Processamento"
      Begin VB.Menu mnuProcessamento_CadLCT 
         Caption         =   "Cadastro  - LCT (LCT/LCI)"
      End
      Begin VB.Menu mnuProcessamento_CadTAG 
         Caption         =   "Cadastro  - TAG (TAG /TGT)"
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
    sta_Barra_MDI.Panels(1).Text = gsUser
    sta_Barra_MDI.Panels(3).Text = Date
    sta_Barra_MDI.Panels(4).Text = Time
    sta_Barra_MDI.Panels(5).Text = "E: " + Format(gsEst_Codigo) + " - V: " + Format(app.Major) + "." + Format(app.Minor) + "." + Format(app.Revision)
End Sub

Private Sub MDIForm_Load()

Me.Top = 10
Me.Left = 10
Me.Height = 14000

checkFunctions

Me.Caption = " SCE - Cadastro de Tag - " + gsEst_Nome
MDI.sta_Barra_MDI.Panels(4).Text = Time

If UCase(gsListas) = "EST" Then
frmCadEst.Show
Else
'frmCadLCTCD.Show
frmCadCD.Show
End If



End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim i As Integer
Dim msg

If Not gbTransferStatus Then
    msg = "Você Realmente esta Saindo da Aplicação - SCETAG "
    Call MsgBoxService(msg, vbQuestion + vbYesNo, Me.Caption)
    Exit Sub
End If
MsgBoxService "O SCETAG está executando uma transferência de arquivos... Favor Aguardar", vbOKOnly, Me.Caption

Cancel = True

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

dbApp.Close

End Sub

Private Sub mnuArquivo_Sair_Click()
    Unload Me
    End
End Sub


Private Sub mnuProcessamento_Cad_ListaNela_Click()

End Sub

Private Sub mnuConfig_Click()
frmConfig.Show
End Sub

Private Sub mnuProcessamento_CadLCT_Click()
'frmCadLCTCD.Show
End Sub

Private Sub mnuProcessamento_CadTAG_Click()
frmCadCD.Show
End Sub

Private Sub Timer1_Timer()
Dim i As Long
Dim ponto As String
Dim frase As String
Dim rs As New Recordset
rs.CursorType = adOpenStatic

On Error GoTo timererro

LerGlobais

MDI.sta_Barra_MDI.Panels(3).Text = Date
MDI.sta_Barra_MDI.Panels(4).Text = Time
DoEvents

If UCase(gsListas) = "EST" Then
frmCadEst.Show
Else
'frmCadLCTCD.Show
frmCadCD.Show

End If

If Not gbTransferStatus And giTransfere = 0 Then
        gbTransferStatus = True
        
        Call LerGlobais
        
        'ponto = "TRATAFILE GZ"
        'Call TrataFileGZ
        DoEvents
        
        'VERIFICA O TIPO DE LISTA CONFIGURADA
        If UCase(gsListas) = "EST" Then
        
            'EXIBE MENSAGEM DE BUSCA DE ARQUIVOS
            ponto = "CGMPRECEBEEST"
            MDI.sta_Barra_MDI.Panels(2).Text = "Buscando arquivos"
            
            'CHAMA ROTINA PARA BUSCAR ARQUIVOS TAG
            Call CGMPRecebeTG
            DoEvents
            MDI.sta_Barra_MDI.Panels(2).Text = ""
            
            'LIMPA TABELA TB_CADESTCTL
           
            frase = "delete TB_CADESTCTL where lseqfile is null"
            Set rs = dbApp.Execute(frase)
            
            'EXIBE MENSAGEM DE TRATAMENTO DE ARQUIVO
            MDI.sta_Barra_MDI.Panels(2).Text = "Atualizando EST"
            ponto = "TRATAFILE EST"
            
            'CHAMA ROTINA PARA TRATAR ARQUIVO EST
            Call TrataFileEST
            
            DoEvents
        Else
        
            'EXIBE MENSAGEM DE BUSCA DE ARQUIVOS
            ponto = "CGMPRECEBETG"
            MDI.sta_Barra_MDI.Panels(2).Text = "Buscando arquivos"
            
            'CHAMA ROTINA PARA BUSCAR ARQUIVOS TAG
            Call CGMPRecebeTG
            DoEvents
            MDI.sta_Barra_MDI.Panels(2).Text = ""
            
            'LIMPA TABELA TB_CADTAGCTL
            frase = "delete tb_CadtagCtl where lseqfile is null"
            Set rs = dbApp.Execute(frase)
            
            'EXIBE MENSAGEM DE TRATAMENTO DE ARQUIVO
            MDI.sta_Barra_MDI.Panels(2).Text = "Atualizando TAG"
            ponto = "TRATAFILE TAG"
            
            'CHAMA ROTINA PARA TRATAR ARQUIVO TAG
            Call TrataFileTG
            
            DoEvents
            
        End If
        
        
        MDI.sta_Barra_MDI.Panels(2).Text = ""
        giTransfere = gsTempoBusca
        gbTransferStatus = False
Else
   If gbTransferStatus Or giTransfere < 1 Then
      giTransfere = gsTempoBusca
   End If
   giTransfere = giTransfere - 1
End If

Exit Sub

timererro:
    Call TrataErro(app.title, Error, "TimerMDI_SCETG_ETT")
    giTransfere = gsTempoBusca
    gbTransferStatus = False
End Sub
