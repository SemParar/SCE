VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRelExportar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar Dados do Dia"
   ClientHeight    =   5970
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   11430
   Begin VB.CommandButton Saircmd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10200
      Picture         =   "frmRelExportar.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4800
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPData 
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   16449537
      CurrentDate     =   38757
   End
   Begin VB.CommandButton imprimecmd 
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
      Height          =   975
      Left            =   10200
      Picture         =   "frmRelExportar.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Lercmd 
      Caption         =   "Ler Dados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10200
      Picture         =   "frmRelExportar.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdmov 
      Height          =   3132
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   9732
      _ExtentX        =   17171
      _ExtentY        =   5530
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grddia 
      Height          =   852
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   11172
      _ExtentX        =   19711
      _ExtentY        =   1508
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
   Begin VB.Label Label1 
      Caption         =   "Data do Movimento:"
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
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "frmRelExportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim frase As String
Dim rs As New Recordset
Dim rsdia As New Recordset
Dim rsmov As New Recordset
Dim rsent As New Recordset
Dim rssai As New Recordset
Dim Sprevisto As String
Dim Filename As String

Private Sub exportar()
Dim arch As Integer
Dim cont As Integer
Dim Y, x As Integer
Dim file As String
Dim frase As String

If Sprevisto <> "Sem Movimento" Then

arch = FreeFile
file = "SP" & Format(DTPData, "yyyyMMdd") & ".TXT"
Open gsPath_REL & file For Output As arch
 
For Y = 1 To grdmov.Rows - 1
    frase = ""
    If grdmov.TextMatrix(Y, 0) = "01" Then
        frase = frase + grdmov.TextMatrix(Y, 0) + "," + grdmov.TextMatrix(Y, 1) + "," + IIf(grdmov.TextMatrix(Y, 2) = "", "0", grdmov.TextMatrix(Y, 2))
        Print #arch, frase
    End If
    If grdmov.TextMatrix(Y, 0) = "02" Then
        frase = frase + grdmov.TextMatrix(Y, 0) + "," + grdmov.TextMatrix(Y, 1) + "," + IIf(grdmov.TextMatrix(Y, 3) = "", "0", grdmov.TextMatrix(Y, 3)) + "," + IIf(grdmov.TextMatrix(Y, 4) = "", "0", grdmov.TextMatrix(Y, 4)) + "," + IIf(grdmov.TextMatrix(Y, 5) = "", "0", grdmov.TextMatrix(Y, 5)) + "," + IIf(grdmov.TextMatrix(Y, 6) = "", "0", grdmov.TextMatrix(Y, 6))
        Print #arch, frase
    End If
    If grdmov.TextMatrix(Y, 0) = "03" Then
        frase = frase + grdmov.TextMatrix(Y, 0) + "," + IIf(grdmov.TextMatrix(Y, 2) = "", "0", grdmov.TextMatrix(Y, 2) + "," + IIf(grdmov.TextMatrix(Y, 3) = "", "0", grdmov.TextMatrix(Y, 3)) + "," + IIf(grdmov.TextMatrix(Y, 6) = "", "0", grdmov.TextMatrix(Y, 6)))
        frase = frase + "," + Mid(Sprevisto, 1, 1)
        Print #arch, frase
    End If
Next
 
Close arch

End If

End Sub



Private Sub Form_Load()

rs.CursorType = adOpenStatic

Me.Top = 10
Me.Left = 10

frase = ""
frase = frase & "Select max(datamovimento) from tb_trnctl where tipo = 'TR'"
Set rs = dbApp.Execute(frase)
If rs.BOF And rs.EOF Then
    MsgBoxService "Nao tem Arquivos de Transacao Fechados"
Else
    DTPData = Mid(rs(0), 7, 2) & "/" & Mid(rs(0), 5, 2) & "/" & Mid(rs(0), 1, 4)
    DTPData.MaxDate = DTPData
End If
Call Lercmd_Click

End Sub

Private Sub Form_Unload(Cancel As Integer)

Call ImprimeRelDel(Filename)

End Sub



Private Sub imprimecmd_Click()
Dim fraseaux() As String

Filename = "MOV_" & gsEst_Codigo & "_" & TxtAno + TxtMes + TxtDia & "_" + Format(Now, "YYYYMMDDHHMMSS") & ".HTML"

'Filename = "RelMovDia" & Format(DTPData, "YYYYMMDD") & ".html"
ReDim fraseaux(1)
fraseaux(0) = " Dados para o Movimento de : " & Format(DTPData, "DD/MM/yyyy")
Call Lercmd_Click
Call ImprimeHeader(Filename, " Entradas e Saidas " + Sprevisto)
Call ImprimeExtra(Filename, fraseaux)
Call Imprimegrid(Filename, grdmov)
Call ImprimeFooter(Filename, "Impresso em : " + Format(Now, "DD/MM/YY HH:MM:SS"))
Call ImprimeRel(Filename)
Call exportar

End Sub

Private Sub Lercmd_Click()

Set rsdia = Nothing
frase = ""
frase = frase & " select"
frase = frase & "         isnull(convert(char(8),cast(ra.datamovimento as smalldatetime),3),'-') as DtMovimento,"
frase = frase & "         isnull(ra.seqfile,'-') as SeqArquivo,"
frase = frase & "         isnull(str(cast(ra.reginformado as int),7),'-') as QtdeReg ,"
frase = frase & "         isnull(str(cast(ra.totinformado as int)/100.00,10,2),'-') as ValorEnviado ,"
frase = frase & "         isnull(str((select count(*) from tb_relfin1 as rd where " & IIf(ChkZerados, "cast(valor as int) > 0  and ", "") & " ra.seqfile = rd.seqfile),8),'-') as QtdeRejeitada ,"
frase = frase & "         isnull(str(cast(rb.totencontrado as int)/100.00,10,2),'-') as ValorRejeitado,"
frase = frase & "         isnull(str(cast(rb.totinformado as int)/100.00,10,2),'Faltante') as ValorAceito ,"
frase = frase & "         isnull(convert(char(8),cast(rc.datapagamento as smalldatetime),3),'-') as DtPGTO,"
frase = frase & "         isnull(str(cast(rc.valorcgmp as int)/100.00,10,2),'Faltante') as ValorCGMP ,"
frase = frase & "         isnull(str(cast(rc.valorsgmp as int)/100.00,10,2),'Faltante') as ValorSGMP"
frase = frase & " From"
frase = frase & "         tb_trnctl as ra ,"
frase = frase & "         tb_trfctl as rb ,"
frase = frase & "         tb_relfin2 As rc"
frase = frase & " Where"
frase = frase & "         ra.seqfile *= rc.seqfile and"
frase = frase & "         ra.datamovimento = '" & Format(DTPData, "YYYYMMDD") & "' and"
frase = frase & "         ra.tipo = 'TR' and"
frase = frase & "         rb.tipo = 'RF' and"
frase = frase & "         rb.seqfile =* ra.seqfile"
frase = frase & " Order By"
frase = frase & "         ra.seqfile"
Set rsdia = dbApp.Execute(frase)

If rsdia.EOF And rsdia.BOF Then
   Sprevisto = "Sem Movimento"
Else
    rsdia.MoveFirst
    If rsdia(6) = "Faltante" Then
       Sprevisto = "Previsto"
    Else
        Sprevisto = " "
    End If
End If

Set rsdia = dbApp.Execute(frase)

grddia.Clear
Set grddia.DataSource = rsdia

frase = ""
frase = frase & "calcmovlimpa"
dbApp.Execute (frase)

frase = ""
frase = frase & "calcmovent '" & Format(DTPData, "YYYYMMDD") & " " & Format(gsEst_Horario, "hh:mm:ss") & "','"
frase = frase & Format(DTPData + 1, "YYYYMMDD") & " " & Format(gsEst_Horario, "hh:mm:ss") & "' "
'Set rsmov = dbApp.Execute(frase)
dbApp.Execute (frase)

frase = ""
frase = frase & "calcmovsaitol '" & Format(DTPData, "YYYYMMDD") & " " & Format(gsEst_Horario, "hh:mm:ss") & "','"
frase = frase & Format(DTPData + 1, "YYYYMMDD") & " " & Format(gsEst_Horario, "hh:mm:ss") & "' "
'Set rsmov = dbApp.Execute(frase)
dbApp.Execute (frase)

frase = ""
frase = frase & "calcmovsaipgt '" & Format(DTPData, "YYYYMMDD") & " " & Format(gsEst_Horario, "hh:mm:ss") & "','"
frase = frase & Format(DTPData + 1, "YYYYMMDD") & " " & Format(gsEst_Horario, "hh:mm:ss") & "' "
'Set rsmov = dbApp.Execute(frase)
dbApp.Execute (frase)

frase = ""
frase = frase & "calcmovsaivalor '" & Format(DTPData, "YYYYMMDD") & " " & Format(gsEst_Horario, "hh:mm:ss") & "','"
frase = frase & Format(DTPData + 1, "YYYYMMDD") & " " & Format(gsEst_Horario, "hh:mm:ss") & "' "
'Set rsmov = dbApp.Execute(frase)
dbApp.Execute (frase)

Set rsmov = Nothing
frase = ""
frase = frase & "select tipo as TIPO ,np as PISTA,"
frase = frase & "CAST(sum(tent) as VARCHAR(18)) as ENTRADAS, cast(sum(tsai) as varchar(18)) AS SAIDAS ,"
frase = frase & "cast(sum(stol) as varchar(18)) AS TOLERANCIA,cast(sum(spgt) as Varchar(18)) AS PAGAS,"
frase = frase & "cast(sum(valor) as varchar(25)) AS VALOR from movdia "
frase = frase & "group by tipo,np order by tipo "
Set rsmov = dbApp.Execute(frase)
grdmov.Clear
Set grdmov.DataSource = rsmov
    
grdmov.TextMatrix(0, 0) = "Tipo         "
grdmov.TextMatrix(0, 1) = "Pista        "
grdmov.TextMatrix(0, 2) = "Entradas     "
grdmov.TextMatrix(0, 3) = "Saidas       "
grdmov.TextMatrix(0, 4) = "Tolerancia   "
grdmov.TextMatrix(0, 5) = "Pagas        "
grdmov.TextMatrix(0, 6) = "Valor        "
grdmov.Refresh
    
Me.Refresh

Set rs = Nothing

End Sub


Private Sub saircmd_Click()

Unload Me

End Sub


