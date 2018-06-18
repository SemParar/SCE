VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmConfig 
   Caption         =   "Parametros de Configuracao"
   ClientHeight    =   9075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight     =   9075
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnGeracao 
      Caption         =   "Travar Geracao"
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton btnEnvio 
      Caption         =   "Travar Envio"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdAlteraConfig 
      Caption         =   "Alterar Configuracao"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid dg 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   14631
      _Version        =   393216
      AllowUpdate     =   0   'False
      DefColWidth     =   200
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Parametro"
         Caption         =   "Parametro"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Valor"
         Caption         =   "Valor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Validade"
         Caption         =   "Validade"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   3000,189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2505,26
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1995,024
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adogrid 
      Height          =   330
      Left            =   240
      Top             =   10680
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   0
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"frmConfig.frx":0000
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnEnvio_Click()

Dim respostaInputBox As String

respostaInputBox = InputBox("Digite o motivo:", "Resposta Obrigatoria")

If respostaInputBox <> "" Then
gsstrErro_Posicao = "DbOpen"
    If Not DbOpen(gsPath_DS, gsPath_DB) Then
          MsgBoxService "DataBase SCE não pode ser aberto! Aplicação será finalizada!", vbCritical, "Erro"
          End
    End If

If bEnvioTravado = True Then
retorno = FsetParam("SCE", "CGMPTRN", DateTime.Now, "c:\cgmp-edi\1.envio\trn\")
btnEnvio.Caption = "Travar Envio"
bEnvioTravado = False
MsgBox ("Envio de Arquivo Liberado")
'Gravar Evento 53
Call GravaEventos(53, "", "SCE", 0, respostaInputBox, 0, 0)

Else
retorno = FsetParam("SCE", "CGMPTRN", DateTime.Now, "c:\cgmp-edi\1.envio\trn\travado\")
btnEnvio.Caption = "Liberar Envio"
bEnvioTravado = True
MsgBox ("Envio de Arquivo Travado")

'Gravar Evento 52
Call GravaEventos(52, "", "SCE", 0, respostaInputBox, 0, 0)
End If
 
'Grid Refresh
frmConfig.dg.Refresh
frmConfig.adogrid.Refresh
Else
MsgBox ("Necessario digitar o motivo para acao solicitada")
End If

End Sub

Private Sub btnGeracao_Click()

Dim anoref As String
Dim anoatual As String
Dim anopreenchido As String
Dim resp As String

Dim respostaInputBox As String

respostaInputBox = InputBox("Digite o motivo:", "Resposta Obrigatoria")

If respostaInputBox <> "" Then

gsstrErro_Posicao = "DbOpen"
If Not DbOpen(gsPath_DS, gsPath_DB) Then
    MsgBoxService "DataBase SCE não pode ser aberto! Aplicação será finalizada!", vbCritical, "Erro"
    End
End If
    
anoref = Year(DateTime.Now) + 50
anoatual = Year(DateTime.Now)
anopreenchido = Mid$(gsNextTrn, 7, 4)

If bGeracaoTravada = True Then
'Substituir o ano fake pelo ano atual
gsNextTrn = Format(Replace(gsNextTrn, anopreenchido, anoatual), "DD/MM/YYYY HH:mm:ss")
btnGeracao.Caption = "Travar Geracao"
bGeracaoTravada = False
MsgBox ("Geracao de Arquivo Liberado")

'Gravar Evento 51
Call GravaEventos(51, "", "SCE", 0, respostaInputBox, 0, 0)

Else
'Substituir o ano atual pelo ano fake
gsNextTrn = Format(Replace(gsNextTrn, anoatual, anoref), "DD/MM/YYYY HH:mm:ss")
btnGeracao.Caption = "Liberar Geracao"
bGeracaoTravada = True
MsgBox ("Geracao de Arquivo Travada")

'Gravar Evento 50
Call GravaEventos(50, "", "SCE", 0, respostaInputBox, 0, 0)

End If

retorno = FsetParam("SCE", "NextTRN", DateTime.Now, Format(gsNextTrn, "DD/MM/YYYY HH:mm:ss"))

'Grid Refresh
frmConfig.dg.Refresh
frmConfig.adogrid.Refresh

Else
MsgBox ("Necessario digitar o motivo para acao solicitada")
End If

End Sub


Private Sub cmdAlteraConfig_Click()
Dim param As String
Dim value As String

param = Me.dg.Columns(0).value
value = Me.dg.Columns(1).value


frmSalvarConfig.txtParam = param
frmSalvarConfig.txtValue = value
frmSalvarConfig.txtNewValue = value

frmSalvarConfig.Show

End Sub





Private Sub Form_Load()
Dim sBuffer As String
    Dim lSize As Long
    Dim frase As String
    Dim resultdb As Recordset


dg.ScrollBars = dbgVertical

If gbytNivel_Acesso_Usuario <> gintNIVEL_ADMINISTRADOR Then
    cmdAlteraConfig.Visible = False
    btnEnvio.Visible = False
    btnGeracao.Visible = False
End If


'On Error GoTo LoadError
    LerGlobais
    
    If bEnvioTravado = True Then
    btnEnvio.Caption = "Liberar Envio"
    Else: btnEnvio.Caption = "Travar Envio"
    End If
    
    If bGeracaoTravada = True Then
    btnGeracao.Caption = "Liberar Geracao"
    Else: btnGeracao.Caption = "Travar Geracao"
    End If
    
    gsstrErro_Posicao = "DbOpen"
    If Not DbOpen(gsPath_DS, gsPath_DB) Then
          MsgBoxService "DataBase SCE não pode ser aberto! Aplicação será finalizada!", vbCritical, "Erro"
          End
    End If
   

   
   frase = "SELECT SUBSTRING(szParam,16,LEN(SZPARAM)) Parametro,"
   frase = frase + " szValor valor"
   frase = frase + " ,tsvalidade Validade"
   frase = frase + " FROM " + gsPath_DB + ".[dbo].[tb_AppParam] A"
   frase = frase + " where szParam like 'VMSCE%'"
   frase = frase + " AND tsValidade = (SELECT MAX(TSVALIDADE) FROM tb_AppParam V WHERE A.szParam=V.szParam GROUP BY V.szParam )"
   frase = frase + " order by Parametro"
   
   
      With adogrid
   .ConnectionString = "Provider=SQLOLEDB.1;Password=" + gsPWD + ";Persist Security Info=True;User ID=" + gsuid + ";Initial Catalog=" + gsPath_DB + ";Data Source=" + gsPath_DS + ""
   .RecordSource = frase
   .CommandType = adCmdText
   .BOFAction = adDoMoveFirst
   .CommandTimeout = 30
   .CursorType = adOpenStatic
   .EOFAction = adDoMoveLast
   .Height = 330
   .Width = 3495
   .Left = 240
   .Top = 10680
   .Orientation = adHorizontal
   .LockType = adLockOptimistic
   End With
   
   Set dg.DataSource = adogrid
    Screen.MousePointer = vbDefault
    gsstrErro_Posicao = "Centraliza_Form_Left"
    frmLogin.Left = Centraliza_Form_Left(frmLogin)
    frmLogin.Top = Centraliza_Form_Top(frmLogin)
    Exit Sub
    
LoadError:
    TrataErro Me.Caption, gsstrErro_Posicao, "Config"
End Sub
