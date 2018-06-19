VERSION 5.00
Begin VB.Form frmExportaConsumos 
   Caption         =   "Exportação de Consumidores com Consumos"
   ClientHeight    =   1950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2730
   LinkTopic       =   "Form2"
   ScaleHeight     =   1950
   ScaleWidth      =   2730
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   2280
      Top             =   1320
   End
   Begin VB.CommandButton Gerar 
      Caption         =   "Gerar Relatório"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Ano 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Text            =   "2013"
      Top             =   720
      Width           =   1215
   End
   Begin VB.ComboBox Mes 
      Height          =   315
      ItemData        =   "frmExportaConsumos.frx":0000
      Left            =   1080
      List            =   "frmExportaConsumos.frx":0002
      TabIndex        =   0
      Text            =   "1"
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label LAno 
      Caption         =   "Ano"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   495
   End
   Begin VB.Label LMes 
      Caption         =   "Mês"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmExportaConsumos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Verifica se digitou corretamente o ano em que deseja o relatório de consumos
'
'
'
Private Sub Ano_LostFocus()
    If IsNumeric(Me.Ano) = False Then
        MsgBox ("O ano necessita ser um campo numérico")
    End If
End Sub

Private Sub Form_Load()
    Mes.AddItem "1"
    Mes.AddItem "2"
    Mes.AddItem "3"
    Mes.AddItem "4"
    Mes.AddItem "5"
    Mes.AddItem "6"
    Mes.AddItem "7"
    Mes.AddItem "8"
    Mes.AddItem "9"
    Mes.AddItem "10"
    Mes.AddItem "11"
    Mes.AddItem "12"
    varGlobais.pararExecucao = False               'indica que iniciará sem sem informar que deverá parar a execução
    FrmMain.Timer1.Enabled = True                  'habilita o timer
End Sub


Private Sub Form_Unload(Cancel As Integer)
    FrmMain.Timer1.Enabled = False                 'desahabilita o timer
End Sub

' Gera relatório
'
'
'
Private Sub Gerar_Click()
    Dim CAMINHO, SQL As String
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Dim TB_GEOMETRIA As String
    Dim mPROVEDOR As String
    Dim mSERVIDOR As String
    Dim mPORTA As String
    Dim mBANCO As String
    Dim mUSUARIO As String
    Dim Senha As String
    Dim decriptada As String
    Dim conexao As New ADODB.connection
    Dim strConn As String
    Dim nStr As String
    Dim count2 As Integer
    Dim Mes As String
    Dim Ano As String
    Dim consumo As New CConsumo                                                     'funções de conversão de l/s para m3/mês
    Dim nomeArquivo As New CArquivo                                                 'para o usuário selecionar onde será salvo o arquivo
    
    CAMINHO = nomeArquivo.SelecionaDiretorio                                        'solicita ao usuário a seleção de um diretório
    CAMINHO = CAMINHO + "\" + nomeArquivo.prefixo + "Exportação dos consumos nas ligações.txt"  'coloca um prefixo de data e hora em que o arquivo será gerado
    Mes = frmExportaConsumos.Mes.Text
    Ano = frmExportaConsumos.Ano.Text
    If frmCanvas.TipoConexao = 4 Then
        If count2 <> 10 Then
            mSERVIDOR = ReadINI("CONEXAO", "SERVIDOR", App.path & "\CONTROLES\GEOSAN.ini")
            mPORTA = ReadINI("CONEXAO", "PORTA", App.path & "\CONTROLES\GEOSAN.ini")
            mBANCO = ReadINI("CONEXAO", "BANCO", App.path & "\CONTROLES\GEOSAN.ini")
            mUSUARIO = ReadINI("CONEXAO", "USUARIO", App.path & "\CONTROLES\GEOSAN.ini")
            Senha = ReadINI("CONEXAO", "SENHA", App.path & "\CONTROLES\GEOSAN.ini")
            nStr = frmCanvas.FunDecripta(Senha)
            decriptada = frmCanvas.Senha
            strConn = "DRIVER={PostgreSQL Unicode}; DATABASE=" + mBANCO + "; SERVER=" + mSERVIDOR + "; PORT=" + mPORTA + "; UID=" + mUSUARIO + "; PWD=" + nStr + "; ByteaAsLongVarBinary=1;"
            conexao.Open strConn
            count2 = 10
        End If
    End If
    If frmCanvas.TipoConexao <> 4 Then
        SQL = "SELECT * FROM NXGS_V_LIG_COMERCIAL_CONSUMO N INNER JOIN RAMAIS_AGUA_LIGACAO R on R.NRO_LIGACAO / 10=N.NRO_LIGACAO_SEM_DV inner join NXGS_V_LIG_COMERCIAL F on  R.NRO_LIGACAO / 10 = F.NRO_LIGACAO_SEM_DV WHERE N.ANO = " & Ano & " AND N.MES = " & Mes & " ORDER BY R.NRO_LIGACAO ASC"
        rs.Open SQL, Conn, adOpenDynamic, adLockReadOnly
    Else
        'precisa atualizar esta querie para funcionar com o banco Postgres
        SQL = "SELECT * FROM " + """" + "NXGS_V_LIG_COMERCIAL_CONSUMO" + """" + " N INNER JOIN " + """" + "RAMAIS_AGUA_LIGACAO" + """" + " R on CAST(R." + """" + "NRO_LIGACAO" + """" + " AS INTEGER)=N." + """" + "NRO_LIGACAO" + """" + "inner join " + """" + "NXGS_V_LIG_COMERCIAL" + """" + "F ON R." + """" + "NRO_LIGACAO" + """" + "=F." + """" + "NRO_LIGACAO" + """" + " ORDER BY R." + """" + "NRO_LIGACAO" + """" + " ASC"
        rs.Open SQL, conexao, adOpenDynamic, adLockReadOnly
    End If
    frmExportaConsumos.Hide                     'esconde a caixa de diálogo
    If frmCanvas.TipoConexao <> 4 Then
        Open CAMINHO For Output As #1
        Print #1, "CONSUMIDOR;NRO_LIGACAO;CONSUMO MEDIDO (M3/MÊS);CONSUMO MÉDIO (M3/MES); CONSUMO MÉDIO (L/S);MES,ANO"
        Do While Not rs.EOF = True
            DoEvents                                                                'para o VB poder escutar o timer e poder parar o processamento caso a tecla ESC tenha sido pressionada
            If varGlobais.pararExecucao = True Then
                varGlobais.pararExecucao = False
                Screen.MousePointer = vbNormal
                Close #1
                rs.Close
                Exit Sub
            End If
            FrmMain.sbStatusBar.Panels(2).Text = "Consumidor: " & rs!NRO_LIGACAO    'mostra na barra de status o consumidor que está sendo exportado
            Print #1, rs!CONSUMIDOR & ";"; rs!NRO_LIGACAO & ";" & rs!consumo_medido & ";" & Round(consumo.lps2m3mes(rs!consumo_lps), 1) & ";" & Round(rs!consumo_lps, 5) & ";" & rs!Mes & ";" & rs!Ano
            rs.MoveNext
        Loop
        Close #1
        rs.Close
        'MousePointer = vbDefault
    Else
        'precisa verificar esta implementação em Postgres
        Open CAMINHO For Output As #1
        Print #1, "CONSUMIDOR;NRO_LIGACAO;CONSUMO;COORDENADAS ESPACIAIS;MES,ANO"
        Do While Not rs.EOF = True
            Print #1, rs!CONSUMIDOR & ";"; rs!NRO_LIGACAO & ";" & rs!consumo_medido & ";" & rs3!spatial_data & ";" & rs!Mes & ";" & rs!Ano
            rs.MoveNext
        Loop
        Close #1
        rs.Close
    End If
    MsgBox "Arquivo exportado em " & CAMINHO & ".", vbInformation, "Exportação Concluída!"
    Unload Me
    Exit Sub

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        Close #1
        rs.Close
        ErroUsuario.Registra "frmExportaConsumidores", "Gerar_Click", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub
' Configura um timer para caso o usuário selecione a tecla ESC ele pare a execução
'
' varGlobais.pararExecucao - contem a informação que deve ser configurada na rotina que deseja-se cancelar a execução. Lembrando-se de colocar um Doevents antes. Veja o exemplo abaixo
' o intervalo do timer está definido no MDIForm_Load
'
'DoEvents                                                            'para o VB poder escutar o timer e poder parar o processamento caso a tecla ESC tenha sido pressionada
'If varGlobais.pararExecucao = True Then
'    varGlobais.pararExecucao = False
'    Screen.MousePointer = vbNormal
'    Exit Sub
'End If
'
' O timer deve ser habilitado antes de entrar na rotina que requer cálculo intensivo. Veja o exemplo abaixo:
'FrmMain.Timer1.Enabled = True                               'habilita o timer
'
Private Sub Timer1_Timer()
    If GetAsyncKeyState(VK_ESCAPE) Then
        MsgBox ("Comando cancelado.")
        varGlobais.pararExecucao = True
    End If
End Sub
