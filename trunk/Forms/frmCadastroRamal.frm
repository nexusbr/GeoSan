VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCadastroRamal 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de Ramal de Água"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkExecFiltroPorLote 
      Caption         =   "Executar Filtro por Lote"
      Height          =   285
      Left            =   210
      TabIndex        =   43
      Top             =   7860
      Width           =   1980
   End
   Begin VB.CheckBox chkExecPreFiltro 
      Caption         =   "Executar Pré Filtro ao iniciar"
      Height          =   285
      Left            =   210
      TabIndex        =   42
      Top             =   7545
      Width           =   2580
   End
   Begin VB.Frame Frame5 
      Caption         =   "Ligações Fictícias"
      Height          =   1395
      Left            =   150
      TabIndex        =   32
      Top             =   6045
      Width           =   4845
      Begin VB.Frame Frame7 
         Caption         =   "Consumo (médio/ligação)"
         Height          =   1005
         Left            =   2100
         TabIndex        =   35
         Top             =   285
         Width           =   2580
         Begin VB.TextBox txtConsumoFicticia 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1215
            TabIndex        =   38
            Text            =   "0.00"
            ToolTipText     =   "Informe o consumo médio de uma ligação"
            Top             =   435
            Width           =   1140
         End
         Begin VB.OptionButton optMetroCubico 
            Caption         =   "M³/Mês"
            Height          =   285
            Left            =   150
            TabIndex        =   37
            Top             =   285
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.OptionButton optLitrosSegundo 
            Caption         =   "LPS"
            Height          =   285
            Left            =   150
            TabIndex        =   36
            Top             =   615
            Width           =   870
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Quantidade"
         Height          =   1005
         Left            =   480
         TabIndex        =   33
         Top             =   285
         Width           =   1320
         Begin VB.VScrollBar UpDown2 
            Height          =   615
            Left            =   840
            TabIndex        =   44
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox txtQtd 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   180
            TabIndex        =   34
            Text            =   "0"
            Top             =   405
            Width           =   495
         End
      End
   End
   Begin VB.CommandButton cmdConsultarLigacoes 
      Caption         =   "Consultar consumo"
      Height          =   390
      Left            =   5130
      TabIndex        =   31
      Top             =   7740
      Width           =   1740
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   390
      Left            =   6930
      TabIndex        =   30
      Top             =   7740
      Width           =   1065
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Salvar"
      Height          =   390
      Left            =   8040
      TabIndex        =   29
      Top             =   7740
      Width           =   1035
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pré Filtro"
      Height          =   1155
      Left            =   150
      TabIndex        =   25
      Top             =   150
      Width           =   8925
      Begin VB.OptionButton optConsumidor 
         Caption         =   "Consumidor"
         Height          =   225
         Left            =   6255
         TabIndex        =   6
         Top             =   360
         Width           =   2085
      End
      Begin VB.OptionButton optEndereço 
         Caption         =   "Endereço"
         Height          =   225
         Left            =   3945
         TabIndex        =   4
         Top             =   360
         Width           =   2220
      End
      Begin VB.OptionButton optNumLigacao 
         Caption         =   "Inscrição"
         Height          =   240
         Left            =   105
         TabIndex        =   0
         Top             =   345
         Width           =   1800
      End
      Begin VB.CommandButton cmdAtivaFiltro 
         Caption         =   ">>"
         Default         =   -1  'True
         Height          =   360
         Left            =   8415
         TabIndex        =   8
         Top             =   600
         Width           =   405
      End
      Begin VB.OptionButton optInscricao 
         Caption         =   "Ligação / Matrícula"
         Height          =   210
         Left            =   1995
         TabIndex        =   2
         Top             =   360
         Width           =   1740
      End
      Begin VB.TextBox txtConsumidor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6270
         TabIndex        =   7
         Top             =   600
         Width           =   2085
      End
      Begin VB.TextBox txtEndereco 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3930
         TabIndex        =   5
         Top             =   600
         Width           =   2265
      End
      Begin VB.TextBox txtNumLigacao 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   90
         TabIndex        =   1
         Top             =   600
         Width           =   1845
      End
      Begin VB.TextBox txtInscricao 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2010
         TabIndex        =   3
         Top             =   600
         Width           =   1845
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Selecione os consumidores associados ao ramal"
      Height          =   2475
      Left            =   150
      TabIndex        =   17
      Top             =   1620
      Width           =   8925
      Begin MSComctlLib.ListView lvLigacoes 
         Height          =   2055
         Left            =   240
         TabIndex        =   45
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "LigMat"
            Text            =   "Inscrição"
            Object.Width           =   3529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "Insc"
            Text            =   "Ligação / Matricula"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "Ende"
            Text            =   "Endereço"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "Cons"
            Text            =   "Consumidor"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "Tpo"
            Text            =   "Tipo"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdPesquisaLigacoes 
         Caption         =   "Pesquisar Ligações"
         Height          =   375
         Left            =   6450
         TabIndex        =   24
         Top             =   2775
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   270
         Left            =   180
         TabIndex        =   26
         Top             =   2715
         Width           =   2310
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Ramal"
      Height          =   1740
      Left            =   150
      TabIndex        =   18
      Top             =   4185
      Width           =   8940
      Begin VB.TextBox txtProfundidade 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5760
         TabIndex        =   12
         Top             =   660
         Width           =   1845
      End
      Begin VB.TextBox txtComprimentoRamal 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5760
         TabIndex        =   11
         Top             =   300
         Width           =   1845
      End
      Begin VB.Frame Frame2 
         Caption         =   "Posicionamento em relação ao lote"
         Height          =   615
         Left            =   135
         TabIndex        =   19
         Top             =   1035
         Width           =   8685
         Begin VB.OptionButton optEsquerdo 
            Caption         =   "Esquerdo"
            Height          =   225
            Left            =   2745
            TabIndex        =   14
            Top             =   300
            Width           =   975
         End
         Begin VB.OptionButton optCentro 
            Caption         =   "Centro"
            Height          =   225
            Left            =   4995
            TabIndex        =   15
            Top             =   300
            Width           =   855
         End
         Begin VB.OptionButton optDireito 
            Caption         =   "Direito"
            Height          =   225
            Left            =   7155
            TabIndex        =   16
            Top             =   300
            Width           =   1125
         End
         Begin VB.OptionButton optDesconhecido 
            Caption         =   "Desconhecido"
            Height          =   225
            Left            =   450
            TabIndex        =   13
            Top             =   300
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.TextBox txtDistanciaTestada 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2055
         TabIndex        =   9
         Top             =   285
         Width           =   1845
      End
      Begin VB.TextBox txtDistanciaLado 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2055
         TabIndex        =   10
         Top             =   645
         Width           =   1845
      End
      Begin VB.Label Label7 
         Caption         =   "Profundidade"
         Height          =   195
         Left            =   4560
         TabIndex        =   23
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label6 
         Caption         =   "Comprimento"
         Height          =   195
         Left            =   4560
         TabIndex        =   22
         Top             =   315
         Width           =   1080
      End
      Begin VB.Label Label3 
         Caption         =   "Distância Testada"
         Height          =   195
         Left            =   570
         TabIndex        =   21
         Top             =   345
         Width           =   1440
      End
      Begin VB.Label Label5 
         Caption         =   "Distância Lado"
         Height          =   195
         Left            =   570
         TabIndex        =   20
         Top             =   690
         Width           =   1290
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Configurar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2280
      MousePointer    =   1  'Arrow
      TabIndex        =   41
      Top             =   7845
      Width           =   1065
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      FillStyle       =   0  'Solid
      Height          =   240
      Left            =   5790
      Shape           =   3  'Circle
      Top             =   6645
      Width           =   225
   End
   Begin VB.Label Label2 
      Caption         =   "Ramal conectado a rede:"
      Height          =   270
      Left            =   5445
      TabIndex        =   40
      Top             =   6105
      Width           =   3405
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      FillStyle       =   0  'Solid
      Height          =   240
      Left            =   8220
      Shape           =   3  'Circle
      Top             =   6645
      Width           =   225
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   6
      X1              =   5940
      X2              =   8235
      Y1              =   6765
      Y2              =   6765
   End
   Begin VB.Label lblUsuarioData 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   5145
      TabIndex        =   28
      Top             =   7110
      Width           =   3960
   End
   Begin VB.Label lblResultado 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   255
      TabIndex        =   27
      Top             =   1335
      Width           =   7680
   End
   Begin VB.Label lblRede 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Left            =   6210
      TabIndex        =   39
      Top             =   6435
      Width           =   1770
   End
End
Attribute VB_Name = "FrmCadastroRamal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private object_id_ramal As String
Private tcs As TeCanvas, tdbramais As TeDatabase, tdbtrecho As TeDatabase, object_id_lote As String, Object_id_trecho As String
Private rs As ADODB.Recordset
Dim blnCancelar As Boolean
Dim i As Long
Dim j As Long
Dim blnBotaoFechar As Boolean
Dim intKeyAscii As Integer
Dim iniAtivado As String 'ARMAZENA A INFORMACAO SE A PESQUISA POR LOTE ESTA ATIVADA
Dim iniTabela As String
Dim iniREF_IPTU As String
Dim iniREF_NROLIGACAO As String
Dim TB_Ramais As String
Dim TB_Ligacoes As String
Dim TB_comercial As String
Dim PESQUISA As String
Dim VALOR As String
Dim va As String
Dim ve As String
Dim vi As String
Dim vo As String
Dim vu As String
Dim vc As String
Dim vd As String
Dim vm As String
Dim vf As String
Dim count2, count3 As Integer

Public Sub init(TipoRamal As String, m_object_id_ramal As String, m_tcs As TeCanvas, m_tdbramais As TeDatabase, m_tdbtrecho As TeDatabase, m_object_id_lote As String, m_object_id_trecho As String)

On Error GoTo Trata_Erro

    'Verifica se os ramais estão associados aos polígonos dos lotes. Antigamente o GeoSan tinha as ligações associadas aos lotes
    If ReadINI("RamaisFiltroLotes", "Ativado", App.path & "\CONTROLES\GEOSAN.ini") <> "SIM" Then
        'Não, as ligações não estão associadas ao polígono do lote
        Me.chkExecFiltroPorLote.value = 0
    Else
        'Sim, as ligações estão associadas ao polígono do lote
        Me.chkExecFiltroPorLote.value = 1
    End If
      
    'Verifica se existe algum tipo de filtro ativado para ramais
    If ReadINI("RamaisFiltro", "Ativado", App.path & "\CONTROLES\GEOSAN.ini") <> "SIM" Then
        'Não existem filtros ativados para ramais
        Me.chkExecPreFiltro.value = 0
    Else
        'Sim, existem filtros ativados para ramais
        Me.chkExecPreFiltro.value = 1
    End If

    'Verifica se as ligações serão consultadas pela associação das mesmas aos lotes
    VALOR = ReadINI("RAMAIS", "CONSULTAR_LIGAÇÕES", App.path & "\CONTROLES\GEOSAN.ini")
    
    'Caso tenha sido retornado um valor válido sobre a forma de consulta das ligações, se serão pelo lote ou não
    If VALOR = "" Or VALOR <> "SIM" Or VALOR <> "NÃO" Then
        'Salva no arquivo de inicilização do GeoSan que elas não são consultadas pelo lote
        Call WriteINI("RAMAIS", "CONSULTAR_LIGAÇÕES", "NÃO", App.path & "\CONTROLES\GEOSAN.ini")
        
        'Avisa que a consulta das ligações não será pelo polígono do lote
        VALOR = "NÃO"
    End If
   
    'Configura a visulização da caixa de diálogo de ligações pelo lote
    If VALOR = "NÃO" Then 'no caso de SQL consutar ligações está desabilitado
        cmdConsultarLigacoes.Visible = False
    End If

    'MUDA OS NOMES DAS TABELAS EQUIPARANDO COM OS TIPOS
    'Configura o nome das tabelas que serão acessadas, no caso de esgoto e no caso de água
    If TipoRamal = "ESGOTO" Then
      'Tabelas de esgoto a serem consultadas
      TB_Ramais = "RAMAIS_ESGOTO"
      TB_Ligacoes = "RAMAIS_ESGOTO_LIGACAO"
      TB_comercial = "NXGS_V_LIG_COMERCIAL_E"
      Me.Frame5.Visible = False 'CADASTRO DE LIGAÇÕES FICTÍCIAS
      Me.Caption = "Cadastro de Ramal de Esgoto"
      Me.cmdConsultarLigacoes.Visible = False
    Else
      'Tabelas de água a serem consultadas
      TB_Ramais = "RAMAIS_AGUA"                 'Informações sobre os ramais
      TB_Ligacoes = "RAMAIS_AGUA_LIGACAO"       'Informações sobre as ligações de água
      TB_comercial = "NXGS_V_LIG_COMERCIAL"     'Informações sobre os dados das ligações de água. Geralmente uma vista para o sistema comercial
    End If

        'Faz a verificação de que tipo de usuário está acessando o sistema para poder habilitar ou desabilitar a opção de alterar/salvar os dados da caixa de diálogo
        Dim ga As String
        Dim ge As String
        Dim gi As String
        Dim go As String
        Dim gu As String
        
        ga = "USRLOG"
        ge = "USRFUN"
        gi = "SYSTEMUSERS"
        go = "OBJECT_ID_"
    
        'Se o usuário for um visitante ele não pode alterar/salvar os dados da caixa de diálogo, somente pode consultar
        Set rs = New ADODB.Recordset
        
        If frmCanvas.TipoConexao <> 4 Then
            'Se for banco de dados Postgres
            rs.Open "SELECT USRLOG, USRFUN FROM SYSTEMUSERS WHERE USRLOG = '" & strUser & "' ORDER BY USRLOG", Conn, adOpenDynamic, adLockReadOnly
        Else
            'Se for banco de dados SQLServer ou Oracle
            rs.Open "SELECT " + """" + ga + """" + ", " + """" + ge + """" + " FROM " + """" + gi + """" + "  WHERE " + """" + ga + """" + "  = '" & strUser & "' ORDER BY " + """" + ga + """" + " ", Conn, adOpenDynamic, adLockOptimistic
        End If
        
        'Verifica que tipo de usuário é
        If rs.EOF = False Then
            If rs!UsrFun = 3 Or rs!UsrFun = 4 Then
                'É visitante ou visualizador apenas
                Me.cmdConfirmar.Enabled = False 'DESABILITA O BOTÃO SALVAR
            End If
        End If
        rs.Close

    'LoozeXP1.InitIDESubClassing
    object_id_lote = m_object_id_lote
    object_id_ramal = m_object_id_ramal
    Object_id_trecho = m_object_id_trecho
    Set tcs = m_tcs
    Set tdbramais = m_tdbramais
    Set tdbtrecho = m_tdbtrecho
   
    If object_id_ramal <> "" Then ' RAMAL EXISTENTE
        'RETORNA ATRIBUTOS DO RAMAL
        Me.Caption = Me.Caption & " - Cod.: " & object_id_ramal
        Set rs = New ADODB.Recordset
        
        'Como a tabela que contem os ramais não possui os números das ligações, primeiro procura o número dos ramais. Aqui abre a conexão para consultar os ramais
        If frmCanvas.TipoConexao <> 4 Then
            'Procura em um banco Postgres
            rs.Open ("SELECT * FROM  " & TB_Ramais & "  WHERE OBJECT_ID_ = '" & object_id_ramal & "'"), Conn, adOpenForwardOnly, adLockReadOnly
        Else
            'Procura em um banco SQLServer ou Oracle
            rs.Open ("SELECT * FROM  " + """" + TB_Ramais + """" + "  WHERE " + """" + go + """" + "  = '" & object_id_ramal & "'"), Conn, adOpenDynamic, adLockOptimistic
        End If
      
        'Caso tenha encontrado algum ramal
        If rs.EOF = False Then
            'Obtem alguns dados do ramal
            txtDistanciaLado.Text = IIf(IsNull(rs!Distancia_Lado), 0, rs!Distancia_Lado)
            txtDistanciaTestada.Text = IIf(IsNull(rs!Distancia_Testada), 0, rs!Distancia_Testada)
            txtProfundidade.Text = IIf(IsNull(rs!Profundidade_RAMAL), 0, rs!Profundidade_RAMAL)
            txtComprimentoRamal.Text = IIf(IsNull(rs!COMPRIMENTO_RAMAL), 0, rs!COMPRIMENTO_RAMAL)
         
            'Obtem o posicionamento do ramal com relação a frente do lote
            Select Case rs!posicionamento_lote
                Case 1
                    optDesconhecido = True
                Case 2
                    optEsquerdo = True
                Case 3
                    optCentro = True
                Case 4
                    optDireito = True
            End Select
      
            Me.lblRede.Caption = rs!Object_id_trecho
            Me.lblUsuarioData.Caption = "Cadastrado por: " & rs.Fields("USUARIO_LOG").value & " em " & rs.Fields("DATA_LOG").value
        End If
        
        rs.Close
        Set rs = Nothing
        'Agora vamos carregar os dados das ligações para serem apresentados na caixa de diálogo
        CarregaLigacoes
    Else
        Me.lblRede.Caption = ramal_Object_id_trecho
        optDesconhecido = True
    End If
   
    If Me.lvLigacoes.ListItems.count > 0 Then
        'Me.cmdConsultarLigacoes.Enabled = True 'DESATIVADO PARA CORRECÇÃO DAS QUERYS DO SQL SERVER
    Else
        If Me.chkExecPreFiltro.value = 1 Then
            PESQUISA = ReadINI("RamaisFiltro", "PESQUISA", App.path & "\CONTROLES\GEOSAN.ini")
            VALOR = ReadINI("RamaisFiltro", "VALOR", App.path & "\CONTROLES\GEOSAN.ini")
            
            If PESQUISA = "NUM_LIGAÇÃO" Then
                Me.optNumLigacao.value = True
                Me.txtNumLigacao = VALOR
            ElseIf PESQUISA = "INSCRIÇÃO" Then
                    Me.optInscricao.value = True
                    Me.txtInscricao.Text = VALOR
            ElseIf PESQUISA = "ENDEREÇO" Then
                Me.optEndereço.value = True
                Me.txtEndereco.Text = VALOR
            ElseIf PESQUISA = "CONSUMIDOR" Then
                Me.optConsumidor.value = True
                Me.txtConsumidor.Text = VALOR
            End If
            
            Me.cmdConsultarLigacoes.Enabled = False
        End If
                
        'CARREGA OS TEXTOS COM O FILTRO PRÉ DETERMINADO
        If Me.chkExecFiltroPorLote.value = 0 Then
            Carrega_PreFiltro (False)
        Else
            Carrega_PreFiltro (True)
        End If
    End If
    Me.Show vbModal
    'LoozeXP1.EndWinXPCSubClassing
    Exit Sub
    
Trata_Erro:
If Err.Number = 0 Or Err.Number = 20 Then
    Resume Next
ElseIf Err.Number = -2147467259 Then
    PrintErro CStr(Me.Name), "Public Sub Init", CStr(Err.Number), CStr(Err.Description), True
    End
Else
    PrintErro CStr(Me.Name), "Public Sub Init", CStr(Err.Number), CStr(Err.Description), True
End If
End Sub


Private Sub chkExecFiltroPorLote_LostFocus()
   If Me.chkExecFiltroPorLote.value = 1 Then
      'GRAVA NO ARQUIVO INI QUE FOI ATIVADA A PESQUISA
      Call WriteINI("RAMAISFILTROLOTES", "ATIVADO", "SIM", App.path & "\CONTROLES\GEOSAN.INI")
   Else
      'GRAVA NO ARQUIVO INI QUE FOI DESATIVADA A PESQUISA
      Call WriteINI("RAMAISFILTROLOTES", "ATIVADO", "NÃO", App.path & "\CONTROLES\GEOSAN.INI")
   End If
End Sub

Private Sub chkExecPreFiltro_LostFocus()
   If Me.chkExecPreFiltro.value = 1 Then
      'GRAVA NO ARQUIVO INI QUE FOI ATIVADA A PESQUISA
      Call WriteINI("RAMAISFILTRO", "ATIVADO", "SIM", App.path & "\CONTROLES\GEOSAN.INI")
   Else
      'GRAVA NO ARQUIVO INI QUE FOI DESATIVADA A PESQUISA
      Call WriteINI("RAMAISFILTRO", "ATIVADO", "NÃO", App.path & "\CONTROLES\GEOSAN.INI")
   End If
End Sub


Private Sub cmdAtivaFiltro_Click()
    
    If cmdAtivaFiltro.Caption = ">>." Then
        blnCancelar = True                      'VARIÁVEL QUE INTERROMPE A PESQUISA
        cmdAtivaFiltro.Caption = ">>"
    Else
        Me.MousePointer = vbHourglass
        Me.lblResultado.Caption = "Pesquisando..."
        DoEvents
        cmdAtivaFiltro.Caption = ">>."
        
        blnCancelar = False
        
        LimpaLista
        Carrega_PreFiltro (False) ' CARREGA FILTROS SEM PESQUISAR O LOTE
        cmdAtivaFiltro.Caption = ">>" 'REATIVA BOTÃO PARA NOVA PESQUISA
    End If
    
    Me.MousePointer = vbDefault
    
End Sub
Private Function LimpaLista()

    Dim blnExisteSelecionado As Boolean
    blnExisteSelecionado = False
reinicia:
    For i = 1 To lvLigacoes.ListItems.count
        If lvLigacoes.ListItems(i).Checked = False Then
            lvLigacoes.ListItems.Remove (i)
            lvLigacoes.Refresh
            GoTo reinicia
        End If
    Next
    

End Function


Private Function Carrega_PreFiltro(ByVal ComLotes As Boolean)

On Error GoTo Trata_Erro
   
Dim str As String
Dim itmx As ListItem
   
Dim strIni As String
Dim criterio As String

     
Dim TIPO_R As String
Dim TB_Ligacao As String
Dim strLigacoes As String

'HIDROMETRADAS E FICTÍCIAS
  
Dim RS_NRO_LIGACAO As New ADODB.Recordset

Dim ma As String
Dim mi As String
Dim mo As String
Dim mu As String
Dim mb As String
Dim mc As String
Dim md As String
Dim mf As String
Dim mg As String
Dim mh As String
Dim mj As String

ma = "NRO_LIGACAO"
mi = "CLASSIFICACAO_FISCAL"
mo = "ENDERECO"
mu = "CONSUMIDOR"
mb = "COD_LOGRADOURO"
mc = "TIPO"
md = "ECONOMIAS"
mf = "HIDROMETRADO"
mg = "NXGS_V_LIG_COMERCIAL"
mh = "NXGS_V_LIG_COMERCIAL_E"
mj = "RAMAIS_ESGOTO_LIGACAO"

   Set rs = New ADODB.Recordset
If tcs.getCurrentLayer = "RAMAIS_AGUA" Then
       'SELECIONA A TABELA OU VIEW QUE POSSUI DADOS DOS CONSUMIDORES DE AGUA
       TIPO_R = "AGUA"
       TB_Ligacoes = "RAMAIS_AGUA_LIGACAO"
       If frmCanvas.TipoConexao <> 4 Then
       strIni = "SELECT NRO_LIGACAO, CLASSIFICACAO_FISCAL, ENDERECO, CONSUMIDOR, COD_LOGRADOURO as " + """" + "CODLOGRAD" + """" + ", TIPO, ECONOMIAS, HIDROMETRADO FROM NXGS_V_LIG_COMERCIAL"
       Else
       strIni = "SELECT " + """" + ma + """" + ", " + """" + mi + """" + ", " + """" + mo + """" + ", " + """" + mu + """" + ", " + """" + mb + """" + " as " + """" + "CODLOGRAD" + """" + ", " + """" + mc + """" + ", " + """" + md + """" + ", " + """" + mf + """" + " FROM " + """" + mg + """" + ""
       End If
   Else
       'SELECIONA A TABELA OU VIEW QUE POSSUI DADOS DOS CONSUMIDORES DE ESGOTO
       TIPO_R = "ESGOTO"
       TB_Ligacoes = "RAMAIS_ESGOTO_LIGACAO"
       If frmCanvas.TipoConexao <> 4 Then
       strIni = "SELECT NRO_LIGACAO, CLASSIFICACAO_FISCAL, ENDERECO, CONSUMIDOR, COD_LOGRADOURO as " + """" + "CODLOGRAD" + """" + ", TIPO, ECONOMIAS, HIDROMETRADO FROM NXGS_V_LIG_COMERCIAL_E"
  Else
  strIni = "SELECT " + """" + ma + """" + ", " + """" + mi + """" + ", " + """" + mo + """" + ", " + """" + mu + """" + ", " + """" + mb + """" + " as " + """" + "CODLOGRAD" + """" + ", " + """" + mc + """" + ", " + """" + md + """" + ", " + """" + mf + """" + " FROM " + """" + mh + """" + ""
  End If
  
   
End If

   
     
   str = "" 'LIMPA A STRING DE COMANDO
   
   If Me.optNumLigacao.value = True And Trim(Me.txtNumLigacao) <> "" Then
   ' If Me.optInscricao.value = True And Trim(Me.txtInscricao) <> "" Then
      
      'Me.lvLigacoes.SortKey = 0 'SETA O SORT PARA A PRIMEIRA COLUNA E TIRA O ORDER BY DO SELECT
      If frmCanvas.TipoConexao = 1 Then 'SQL
                 str = strIni & " WHERE CLASSIFICACAO_FISCAL LIKE '%" & Me.txtNumLigacao.Text & "%' AND NRO_LIGACAO NOT IN (SELECT NRO_LIGACAO FROM " & TB_Ligacoes & ")"
      ElseIf frmCanvas.TipoConexao = 2 Then
           str = strIni & " A WHERE CLASSIFICACAO_FISCAL LIKE '%" & Me.txtNumLigacao.Text & "%' AND NOT EXISTS (SELECT s.NRO_LIGACAO FROM " & TB_Ligacoes & " ss  INNER JOIN NXGS_V_LIG_COMERCIAL s ON s.NRO_LIGACAO = ss.NRO_LIGACAO)"
      Else
         str = strIni & " A WHERE " + """" + "CLASSIFICACAO_FISCAL" + """" + " LIKE '%" & UCase(Me.txtNumLigacao.Text) & "%' AND NOT EXISTS (SELECT s." + """" + "NRO_LIGACAO" + """" + " FROM " & """" + TB_Ligacoes + """" & " ss  INNER JOIN " + """" + "NXGS_V_LIG_COMERCIAL" + """" + " s ON s." + """" + "NRO_LIGACAO" + """" + " = ss." + """" + "NRO_LIGACAO" + """" + ")"
    ' WritePrivateProfileString "A", "A", str, App.path & "\DEBUG.INI"
     End If
   
   ElseIf Me.optInscricao.value = True And Trim(Me.txtInscricao) <> "" Then
         
      'Me.lvLigacoes.SortKey = 1 'SETA O SORT PARA A SEGUNDA COLUNA E TIRA O ORDER BY DO SELECT
      If frmCanvas.TipoConexao = 1 Then 'SQL
         str = strIni & " WHERE NRO_LIGACAO LIKE '%" & Me.txtInscricao.Text & "%' AND NRO_LIGACAO NOT IN (SELECT NRO_LIGACAO FROM " & TB_Ligacoes & ")"
      ElseIf frmCanvas.TipoConexao = 2 Then ' ORACLE
         str = strIni & " A WHERE NRO_LIGACAO LIKE '%" & Me.txtInscricao.Text & "%' AND NOT EXISTS (SELECT s.NRO_LIGACAO FROM " & TB_Ligacoes & " ss  INNER JOIN NXGS_V_LIG_COMERCIAL s ON s.NRO_LIGACAO = ss.NRO_LIGACAO)"
      Else
       str = strIni & " WHERE " + """" + ma + """" + " LIKE '%" & Val(UCase(Me.txtInscricao.Text)) & "%' AND " + """" + ma + """" + " NOT IN (SELECT " + """" + ma + """" + " FROM " + """" + TB_Ligacoes + """" + ")"
    
     ' MsgBox "ARQUIVO DEBUG SALVO"
' WritePrivateProfileString "A", "A", str, App.path & "\DEBUG.INI"
      End If
   
   ElseIf Me.optEndereço.value = True And Trim(Me.txtEndereco) <> "" Then
     
      'Me.lvLigacoes.SortKey = 2 'SETA O SORT PARA A TERCEIRA COLUNA E TIRA O ORDER BY DO SELECT
      If frmCanvas.TipoConexao = 1 Then 'SQL
         str = strIni & " WHERE upper(ENDERECO) LIKE '%" & (Me.txtEndereco.Text) & "%' AND NRO_LIGACAO NOT IN (SELECT NRO_LIGACAO FROM " & TB_Ligacoes & ")" ' ORDER BY TAM ASC, ENDERECO ASC"
      ElseIf frmCanvas.TipoConexao = 2 Then ' ORACLE ' ORACLE
         'NA CONSULTA ORACLE COM LIKE É NECESSÁRIO COLOCAR O SINAL % NO LUGAR DO ESPAÇO ENTRE PALAVRAS
     criterio = "%" & Replace(Trim(Me.txtEndereco.Text), " ", "%") & "%"
         str = strIni & " A WHERE upper(ENDERECO) LIKE '" & criterio & "' AND NOT EXISTS (SELECT s.NRO_LIGACAO FROM " & TB_Ligacoes & " ss  INNER JOIN NXGS_V_LIG_COMERCIAL s ON s.NRO_LIGACAO = ss.NRO_LIGACAO)"
        Else
      Dim za As String
      
      za = UCase("ENDERECO")
      
       str = strIni & " WHERE " + """" + za + """" + " LIKE '%" & Me.txtEndereco.Text & "%' AND " + """" + ma + """" + " NOT IN (SELECT " + """" + ma + """" + " FROM " + """" + TB_Ligacoes + """" + ")" ' ORDER BY TAM ASC, ENDERECO ASC"
      
      
      End If
      
   ElseIf Me.optConsumidor.value = True And Trim(Me.txtConsumidor.Text) <> "" Then
     
     ' Me.lvLigacoes.SortKey = 3 'SETA O SORT PARA A QUARTA COLUNA E TIRA O ORDER BY DO SELECT
      If frmCanvas.TipoConexao = 1 Then 'SQL
         str = strIni & " WHERE (CONSUMIDOR) LIKE '%" & (Me.txtConsumidor.Text) & "%' AND NRO_LIGACAO NOT IN (SELECT NRO_LIGACAO FROM " & TB_Ligacoes & ")"
       ElseIf frmCanvas.TipoConexao = 2 Then ' ' ORACLE
         'NA CONSULTA ORACLE COM LIKE É NECESSÁRIO COLOCAR O SINAL % NO LUGAR DO ESPAÇO ENTRE PALAVRAS
          criterio = "" & Replace(Trim(Me.txtConsumidor.Text), " ", "%") & "%"
       '  str = strIni & " A WHERE upper(CONSUMIDOR) LIKE '%" & criterio & "' AND NOT EXISTS (SELECT s.NRO_LIGACAO FROM " & TB_Ligacoes & " ss  INNER JOIN NXGS_V_LIG_COMERCIAL s ON s.NRO_LIGACAO = ss.NRO_LIGACAO)"
     
     str = strIni & " A WHERE upper(CONSUMIDOR) LIKE '%" & criterio & "'  AND NRO_LIGACAO NOT IN (SELECT NRO_LIGACAO FROM " & TB_Ligacoes & ")"
     
     
     Else
       Dim ze As String
       ze = UCase("CONSUMIDOR")
        str = strIni & " WHERE " + """" + ze + """" + " LIKE '%" & Me.txtConsumidor.Text & "%' AND " + """" + ma + """" + " NOT IN (SELECT " + """" + ma + """" + " FROM " + """" + TB_Ligacoes + """" + ")"
       
      End If
     ' MsgBox "ARQUIVO DEBUG SALVO"
 'WritePrivateProfileString "A", "A", str, App.path & "\DEBUG.INI"
   
   End If

    

   If Me.optNumLigacao.value = True Then
      PESQUISA = "NUM_LIGAÇÃO"
      VALOR = Me.txtNumLigacao.Text
   ElseIf Me.optInscricao.value = True Then
      PESQUISA = "INSCRIÇÃO"
      VALOR = Me.txtInscricao.Text
   ElseIf Me.optConsumidor.value = True Then
      PESQUISA = "CONSUMIDOR"
      VALOR = Me.txtConsumidor.Text
   ElseIf Me.optEndereço.value = True Then
      PESQUISA = "ENDEREÇO"
      VALOR = Me.txtEndereco.Text
   End If
   
   

'MsgBox "ARQUIVO DEBUG SALVO"
 'WritePrivateProfileString "A", "A", str, App.path & "\DEBUG.INI"
   
        
    ' MsgBox "ARQUIVO DEBUG SALVO"
 'WritePrivateProfileString "A", "A", str, App.path & "\DEBUG.INI"
        
        
   'GRAVA NO ARQUIVO INI AS ULTIMAS CONFIGURAÇÕES DO FILTRO
   Call WriteINI("RAMAISFILTRO", "PESQUISA", PESQUISA, App.path & "\CONTROLES\GEOSAN.INI")
   Call WriteINI("RAMAISFILTRO", "VALOR", VALOR, App.path & "\CONTROLES\GEOSAN.INI")
     
    'FAZ SELECT COM BASE NOS CAMPOS CRIADOS
    i = 0
    Me.lblResultado.Caption = "Localizadas " & i & " referencias"
    
    If str <> "" Then
        Set rs = New ADODB.Recordset
      '  If frmCanvas.TipoConexao <> 4 Then
        
       rs.Open str, Conn, adOpenForwardOnly, adLockOptimistic
       ' Else
       '     rs.Open str, Conn, adOpenForwardOnly, adLockOptimistic
       ' End If
      ' Imprima (str)
        
        'Set RS = ConnSec.execute(str)
        If rs.EOF = False Then
            'CARREGA NO FORM TODAS AS LIGAÇÕES DISPONIVEIS COM BASE NO PRÉ FILTRO
            
            'CARREGA_GRID (RS) 'CHAMA A FUNCAO PASSANDO O RECORDSET
            
            Do While Not rs.EOF And blnCancelar = False
                DoEvents

                'Set itmx = lvLigacoes.ListItems.Add(, , rs.Fields("NRO_LIGACAO").value)
                Set itmx = lvLigacoes.ListItems.Add(, , rs.Fields("CLASSIFICACAO_FISCAL").value)

                itmx.SubItems(1) = IIf(IsNull(rs.Fields("NRO_LIGACAO").value), "", rs.Fields("NRO_LIGACAO").value)

                itmx.SubItems(2) = IIf(IsNull(rs.Fields("ENDERECO").value), "", rs.Fields("ENDERECO").value)
                itmx.SubItems(3) = IIf(IsNull(rs.Fields("CONSUMIDOR").value), "", rs.Fields("CONSUMIDOR").value)

                'incluído para mostrar o tipo da ligação
                itmx.SubItems(4) = IIf(IsNull(rs.Fields("TIPO").value), "", rs.Fields("TIPO").value)

                itmx.Tag = rs.Fields("codlograd").value
                i = i + 1
                Me.lblResultado.Caption = "Mostrando " & i & " de " & i & " referencias encontradas"
               If i >= 500 Then
                  j = i
                  Do While Not rs.EOF And blnCancelar = False
                      rs.MoveNext
                      j = j + 1
                  Loop
                  Me.lblResultado.Caption = "Mostrando " & i & " referencias de " & j & " encontradas"
                  Exit Do
               End If

               rs.MoveNext
            Loop
        End If
saida:
        rs.Close
        Set rs = Nothing
    End If
    
   If ComLotes = True Then ' SE A FUNCAO FOI CHAMADA COM TRUE

   
      If idAutoLote <> "" Then 'idAutoLote EH CARREGADO NO MOUSE MOVE DO TE CANVAS
         
         
'         [RAMAISFILTROLOTES]
'         ATIVADO = SIM
'         REF_IPTU = CADASTRO
'         REF_NROLIGACAO = MATRICULA
'         TABELA_PLANO = LOTES_PREF
'         TABELA_ATRIB = LOTES_PREF_DAEV
         
         
         
         iniTabela = ReadINI("RAMAISFILTROLOTES", "TABELA_ATRIB", App.path & "\CONTROLES\GEOSAN.ini")
         
         iniREF_NROLIGACAO = ReadINI("RAMAISFILTROLOTES", "REF_NROLIGACAO", App.path & "\CONTROLES\GEOSAN.ini")
   
         If iniTabela = "" Or iniREF_NROLIGACAO = "" Then
            MsgBox "A pesquisa automatica por lote precisa ser configurada.", vbInformation, ""
         Else
         
         'PESQUISAR QUAIS NUMEROS DE LIGACAO ESTAO NO LOTE

        If frmCanvas.TipoConexao <> 4 Then
         str = "SELECT " & iniREF_NROLIGACAO & " AS " + """" + "NRO_LIGACAO" + """" + " FROM " & iniTabela
         str = str & " WHERE LOTE_ID = '" & idAutoLote & "' AND " & iniREF_NROLIGACAO & " <> '0'"
         
         Else
         Dim qa As String
         qa = "LOTE_ID"
           str = "SELECT " + """" + iniREF_NROLIGACAO + """" + " AS " + """" + "NRO_LIGACAO" + """" + " FROM '" & iniTabela & "'"
         str = str & " WHERE " + """" + qa + """" + " = '" & idAutoLote & "' AND " + """" + iniREF_NROLIGACAO + """" + " <> '0'"
         
         End If
         
         
         Set rs = New ADODB.Recordset
         
         
         rs.Open str, Conn, adOpenKeyset, adLockOptimistic
         str = ""
         strLigacoes = ""
         If rs.EOF = False Then
            Do While Not rs.EOF
               strLigacoes = strLigacoes & "'" & rs!NRO_LIGACAO & "',"
               rs.MoveNext
            Loop
            strLigacoes = mid(strLigacoes, 1, Len(strLigacoes) - 1) 'TIRA A ULTIMA VIRGULA
         
         
            
            'Me.lvLigacoes.SortKey = 1 'SETA O SORT PARA A SEGUNDA COLUNA DO LIST E TIRA O ORDER BY DO SELECT
            If frmCanvas.TipoConexao = 1 Then 'SQL
               str = strIni & " WHERE NRO_LIGACAO IN (" & strLigacoes & ") AND NRO_LIGACAO NOT IN (SELECT NRO_LIGACAO FROM " & TB_Ligacoes & ")"
            ElseIf frmCanvas.TipoConexao = 2 Then 'SQLORACLE
               str = strIni & " A WHERE NRO_LIGACAO IN (" & strLigacoes & ") AND NOT EXISTS (SELECT NRO_LIGACAO FROM " & TB_Ligacoes & " B WHERE A.NRO_LIGACAO = B.NRO_LIGACAO)"
               Else 'Postgres
               str = strIni & " WHERE " + """" + ma + """" + " IN ('" & strLigacoes & "') AND " + """" + ma + """" + " NOT IN (SELECT " + """" + ma + """" + " FROM " + """" + TB_Ligacoes + """" + ")"
            End If
         
            
         
         
            Set rs = New ADODB.Recordset
            rs.Open str, Conn, ReadOnly, adLockOptimistic
            
               If rs.EOF = False Then
                  
                  Do While Not rs.EOF And blnCancelar = False
                     DoEvents
         
                     'Set itmx = lvLigacoes.ListItems.Add(, , rs.Fields("NRO_LIGACAO").value)
                     Set itmx = lvLigacoes.ListItems.Add(, , rs.Fields("CLASSIFICACAO_FISCAL").value)
         
                     itmx.SubItems(1) = IIf(IsNull(rs.Fields("NRO_LIGACAO").value), "", rs.Fields("NRO_LIGACAO").value)
         
                     itmx.SubItems(2) = IIf(IsNull(rs.Fields("ENDERECO").value), "", rs.Fields("ENDERECO").value)
                     itmx.SubItems(3) = IIf(IsNull(rs.Fields("CONSUMIDOR").value), "", rs.Fields("CONSUMIDOR").value)
         
                     'incluído para mostrar o tipo da ligação
                     itmx.SubItems(4) = IIf(IsNull(rs.Fields("TIPO").value), "", rs.Fields("TIPO").value)
         
                     itmx.Tag = rs.Fields("codlograd").value
                     i = i + 1
                     Me.lblResultado.Caption = "Mostrando " & i & " de " & i & " referencias encontradas"
                     If i >= 500 Then
                        j = i
                        Do While Not rs.EOF And blnCancelar = False
                            rs.MoveNext
                            j = j + 1
                        Loop
                        Me.lblResultado.Caption = "Mostrando " & i & " referencias de " & j & " encontradas"
                        Exit Do
                     End If
         
                     rs.MoveNext
                  Loop
               
               End If
            End If
         End If
      End If
      
   End If
    
   cmdFechar.Caption = "Cancelar"

Trata_Erro:

If Err.Number = 0 Or Err.Number = 20 Then
   Resume Next
ElseIf Err.Number = -2147467259 Then
   
   PrintErro CStr(Me.Name), "Private Function Carrega_PreFiltro()", CStr(Err.Number), CStr(Err.Description), True


Else
   
   PrintErro CStr(Me.Name), "Private Function Carrega_PreFiltro()", CStr(Err.Number), CStr(Err.Description), True


End If

'MsgBox "ARQUIVO DEBUG SALVO"
 'WritePrivateProfileString "A", "A", Err.Description, App.path & "\DEBUG.INI"


End Function












'Private Sub CARREGA_GRID(ByVal rs As Recordset)
'On Error GoTo Trata_Erro
'
'   Do While Not rs.EOF And blnCancelar = False
'      DoEvents
'
'      'Set itmx = lvLigacoes.ListItems.Add(, , rs.Fields("NRO_LIGACAO").value)
'      Set itmx = lvLigacoes.ListItems.Add(, , rs.Fields("CLASSIFICACAO_FISCAL").value)
'
'      itmx.SubItems(1) = IIf(IsNull(rs.Fields("NRO_LIGACAO").value), "", rs.Fields("NRO_LIGACAO").value)
'
'      itmx.SubItems(2) = IIf(IsNull(rs.Fields("ENDERECO").value), "", rs.Fields("ENDERECO").value)
'      itmx.SubItems(3) = IIf(IsNull(rs.Fields("CONSUMIDOR").value), "", rs.Fields("CONSUMIDOR").value)
'
'      'incluído para mostrar o tipo da ligação
'      itmx.SubItems(4) = IIf(IsNull(rs.Fields("TIPO").value), "", rs.Fields("TIPO").value)
'
'      itmx.Tag = rs.Fields("codlograd").value
'      i = i + 1
'      Me.lblResultado.Caption = "Mostrando " & i & " de " & i & " referencias encontradas"
'      If i >= 500 Then
'         j = i
'         Do While Not rs.EOF And blnCancelar = False
'             rs.MoveNext
'             j = j + 1
'         Loop
'         Me.lblResultado.Caption = "Mostrando " & i & " referencias de " & j & " encontradas"
'         Exit Do
'      End If
'
'      rs.MoveNext
'   Loop
'Trata_Erro:
'
'   If Err.Number = 0 Or Err.Number = 20 Then
'      Resume Next
'   Else
'
'      PrintErro CStr(Me.Name), "Private Sub CARREGA_GRID", CStr(Err.Number), CStr(Err.Description), True
'
'   End If
'End Sub


'''Private Sub cmdConfirmar_Click()
'''
'''   On Error GoTo Trata_Erro
'''   Dim intlocalerro As Integer
'''   Dim rsCria As ADODB.Recordset
'''   Dim a As Integer
'''   Dim cgeo As New clsGeoReference
'''   Dim x As Double
'''   Dim y As Double
'''   Dim str As String
'''
'''   Dim strNroL As String 'NÚMERO DA LIGACAO
'''   Dim strInsc As String 'NUMERO DA INSCRIÇÃO
'''   Dim strTipo As String 'TIPO DA LIGACAO
'''   Dim strCons As String 'CONSUMO DA LIGACAO
'''   Dim strEcon As String 'QUANTIDADE DE ECONOMIAS NA LIGAÇÃO
'''   Dim strHidr As String
'''
'''   Set rsCria = New ADODB.Recordset 'recordset utilizado para criar o regitro na tabela
'''   Dim strNroLigaSel As String
'''   'Conn.BeginTrans
'''
'''   Conn.Close
'''   Conn.Open
'''
'''   If object_id_ramal = "" Then 'NOVO RAMAL
'''
'''        str = strUser & Now
'''
'''        'O CAMPO ID DA TABELA É POPULADA COM A AUTO NUMERAÇÃO DA TABELA
'''
'''        'Set rsCria = Conn.execute("SELECT ID FROM RAMAIS_AGUA WHERE OBJECT_ID_ = '" & str & "'")
'''
'''        rsCria.Open "RAMAIS_AGUA", Conn, adOpenKeyset, adLockOptimistic, adCmdTable
'''
'''        'CRIA O RAMAL NA TABELA DE ATRIBUTOS
'''        rsCria.AddNew
'''        rsCria.Fields("OBJECT_ID_").value = str
'''        rsCria.Fields("OBJECT_ID_TRECHO").value = "0"
'''        rsCria.Update
'''
'''        object_id_ramal = rsCria.Fields("ID").value
'''        rsCria.Close
'''
'''        tcs.object_id = object_id_ramal
'''        tcs.saveOnMemory
'''        tcs.SaveInDatabase
'''
'''        'LOCALIZA A COORDENADA DO FINAL DA LINHA DO RAMAL
'''        tdbramais.getPointOfLine 0, object_id_ramal, 1, x, y
'''
'''        'INSERE PONTO NO FINAL DA LINHA DO RAMAL
'''        tdbramais.addPoint object_id_ramal, x, y
'''
'''        tdbramais.getPointOfLine 0, object_id_ramal, 0, x, y
'''        'tdb.setCurrentLayer cgeo.GetLayerOperation(tcs.getCurrentLayer, 1)
'''
'''        Object_id_trecho = ramal_Object_id_trecho 'VARIÁVEL RAMAL_OBJECT_ID_TRECHO CARREGADA NO TCANVAS ON_CLICK
'''
'''   Else
'''
'''      'O RAMAL JA EXISTE, RESETA AS INFORMAÇÕES DE CONSUMIDORES ASSOCIADOS A ELE
'''      Conn.execute "DELETE FROM RAMAIS_AGUA_LIGACAO WHERE OBJECT_ID_ = '" & object_id_ramal & "'"
'''
'''   End If
'''
'''   Dim rsRamal As ADODB.Recordset
'''   Set rsRamal = New ADODB.Recordset
'''
'''   rsRamal.Open "SELECT * FROM RAMAIS_AGUA WHERE ID = " & "'" & object_id_ramal & "'", Conn, adOpenKeyset, adLockOptimistic
'''
'''   If rsRamal.EOF = False Then
'''
'''       'ATUALIZA A TABELA DE RAMAIS COMPLEMENTANDO AS INFORMAÇÕES
'''
'''       rsRamal.Fields("Distancia_Lado").value = IIf(IsNumeric(txtDistanciaLado), txtDistanciaLado, 0)
'''       rsRamal.Fields("Distancia_Testada").value = IIf(IsNumeric(txtDistanciaTestada), txtDistanciaTestada, 0)
'''       rsRamal.Fields("Profundidade_RAMAL").value = IIf(IsNumeric(txtProfundidade), txtProfundidade, 0)
'''       rsRamal.Fields("Comprimento_Ramal").value = IIf(IsNumeric(txtComprimentoRamal), txtComprimentoRamal, 0)
'''
'''       For i = 1 To lvLigacoes.ListItems.Count
'''          If lvLigacoes.ListItems(i).Checked = True Then
'''              rsRamal!cod_lograd = lvLigacoes.ListItems(1).Tag 'PEGA O PRIMEIRO LOGRADOURO SELECIONADO NA LISTA
'''              Exit For
'''          End If
'''       Next
'''
'''      If optDesconhecido Then rsRamal.Fields("posicionamento_lote").value = 1
'''      If optEsquerdo Then rsRamal.Fields("posicionamento_lote").value = 2
'''      If optCentro Then rsRamal.Fields("posicionamento_lote").value = 3
'''      If optDireito Then rsRamal.Fields("posicionamento_lote").value = 4
'''
'''      rsRamal!Object_id_ = object_id_ramal
'''      rsRamal!Object_id_trecho = Object_id_trecho
'''      rsRamal!usuario_log = strUser
'''      rsRamal!DATA_LOG = Format(Now, "DD/MM/YY HH:MM")
'''
'''      rsRamal.Update
'''
'''   End If
'''   rsRamal.Close
'''
'''
'''   'RE-ASSOCIA AS LIGAÇÕES AO RAMAL
'''   strNroLigaSel = ""
'''
'''   For a = 1 To lvLigacoes.ListItems.Count
'''       If lvLigacoes.ListItems(a).Checked Then 'PARA CADA ITEM SELECIONADO NA LISTA
'''          If strNroLigaSel <> "" Then
'''             strNroLigaSel = strNroLigaSel & ",'" & lvLigacoes.ListItems(a).SubItems(1) & "'"
'''          Else
'''             strNroLigaSel = "'" & lvLigacoes.ListItems(a).SubItems(1) & "'"
'''          End If
'''       End If
'''   Next
'''
'''
'''   If strNroLigaSel <> "" Then
'''
'''      'SELECIONA AS INFORMAÇÕES DOS CONSUMIDORES NA NXGS_V_LIG_COMERCIAL
'''
'''      str = "SELECT NRO_LIGACAO, CLASSIFICACAO_FISCAL, COD_LOGRADOURO, "
'''      str = str & "TIPO, ECONOMIAS, HIDROMETRADO FROM NXGS_V_LIG_COMERCIAL WHERE NRO_LIGACAO IN (" & strNroLigaSel & ")"
'''
'''      Set rs = New ADODB.Recordset
'''      rs.Open str, Conn, adOpenDynamic, adLockReadOnly, adCmdText
'''
'''       If rs.EOF = False Then
'''         Do While Not rs.EOF
'''
'''            strNroL = Trim(rs!NRO_LIGACAO)                 'NÚMERO DA LIGACAO
'''
'''            If Len(rs!CLASSIFICACAO_FISCAL) > 0 Then
'''               strInsc = Trim(rs!CLASSIFICACAO_FISCAL)
'''            Else
'''               strInsc = ""  'NUMERO DA INSCRIÇÃO
'''            End If
'''
'''            If Len(rs!tipo) > 0 Then
'''               strTipo = Trim(mid(rs!tipo, 1, 20))
'''            Else
'''               strTipo = ""  'TIPO DA LIGACAO
'''            End If
'''
'''            If Len(rs!ECONOMIAS) > 0 Then
'''               strEcon = Trim(rs!ECONOMIAS)
'''            Else
'''               strEcon = 1  'QUANTIDADE DE ECONOMIAS NA LIGAÇÃO
'''            End If
'''
'''            If UCase(rs!HIDROMETRADO) = "SIM" Then 'IDENTIFICA E ARMAZENA EM LETRA MINÚSCULA
'''               strHidr = "sim"
'''            ElseIf UCase(rs!HIDROMETRADO) = "NAO" Or UCase(rs!HIDROMETRADO) = "NÃO" Then
'''               strHidr = "nao"
'''            Else
'''               strHidr = ""
'''            End If
'''
'''            str = "INSERT INTO RAMAIS_AGUA_LIGACAO (OBJECT_ID_,NRO_LIGACAO,INSCRICAO_LOTE,TIPO,HIDROMETRADO,ECONOMIAS,CONSUMO_LPS) "
'''            str = str & "VALUES ('" & object_id_ramal & "','" & strNroL & "','" & strInsc & "','" & strTipo & "','" & strHidr & "','" & strEcon & "','0')"
'''
'''            Conn.execute (str)
'''            rs.MoveNext
'''
'''         Loop
'''      End If
'''      rs.Close
'''
'''   End If
'''
'''   'INSERINDO RAMAL FICTÍCIO SE ESTE FOI SELECIONADO
'''   If CInt(Me.txtQtd.Text) > 0 Then
'''
'''      'CAPTURA O CONSUMO DIGITADO E CONVERTE SE NECESSÁRIO
'''      If CDbl(Me.txtConsumoFicticia.Text) > 0 Then
'''         If Me.optMetroCubico.value = True Then
'''             'SE FOR METRO CUBICO, CONVERTE PARA LITROS POR SEGUNDO
'''             strCons = Replace(Me.txtConsumoFicticia.Text, ".", ",") * 0.00038580246
'''         Else
'''             strCons = Replace(Me.txtConsumoFicticia.Text, ".", ",")
'''         End If
'''      Else
'''         strCons = 0
'''      End If
'''
'''      strCons = Replace(strCons, ",", ".") 'troca a vírgula pelo ponto
'''
'''      For i = 0 To CInt(Me.txtQtd.Text) - 1
'''
'''         str = "INSERT INTO RAMAIS_AGUA_LIGACAO (OBJECT_ID_,NRO_LIGACAO,INSCRICAO_LOTE,TIPO,HIDROMETRADO,ECONOMIAS,CONSUMO_LPS) "
'''         str = str & "VALUES ('" & object_id_ramal & "','999" & object_id_ramal & i & "','999" & object_id_ramal & i & "','FICTÍCIA','nao','1','" & strCons & "')"
'''
'''         Conn.execute (str)
'''
'''      Next
'''
'''   End If
'''
'''
'''   tcs.plotView
'''   Unload Me
'''
'''
'''Trata_Erro:
'''    If Err.Number = 0 Or Err.Number = 20 Then
'''        Resume Next
'''    ElseIf Err.Number = -2147418113 Then ' Erro geral de rede
'''        MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência. Reinicie o sistema.", vbInformation
'''        Open App.path & "\Controles\GeoSanLog.txt" For Append As #1
'''        Print #1, Now & " " & strUser & " " & Versao_Geo & " - frmCadastroRamalAgua - Private Sub cmdConfirmar_Click() - Local Num: " & intlocalerro & " - Erro Num: " & Err.Number & " - " & Err.Description & " Erro Geral de Rede - Programa foi fechado."
'''        Close #1
'''        End
'''
'''    ElseIf Err.Number = -2147417848 Then ' automation error
'''        'Conn.RollbackTrans
'''        MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência. Reinicie o sistema.", vbInformation
'''        Open App.path & "\Controles\GeoSanLog.txt" For Append As #1
'''        Print #1, Now & " " & strUser & " " & Versao_Geo & " - frmCadastroRamalAgua - Private Sub cmdConfirmar_Click() - Local Num: " & intlocalerro & " - Erro Num: " & Err.Number & " - " & Err.Description & " - Programa foi fechado."
'''        Close #1
'''        End
'''
'''    ElseIf Err.Number = -2147467259 Or mid(Err.Description, 1, 9) = "ORA-03114" Then 'PERDA DE CONEXÃO BANCO SQL OU ORACLE
'''       'Conn.RollbackTrans
'''       MsgBox "Não há conexão ativa com o banco de dados. Contate o Administrador de Rede." & Chr(13) & Chr(13) & "O Geosan será fechado.", vbinformation, "Falha de rede"
'''       Open App.path & "\Controles\GeoSanLog.txt" For Append As #1
'''       Print #1, Now & " " & strUser & " " & Versao_Geo & " - frmCadastroRamalAgua - Private Sub cmdConfirmar_Click() - Não há conexão ativa com a rede. Programa foi fechado."
'''       Close #1
'''       End
'''    ElseIf Err.Number = -2147168227 Then ' MAX TRANSACTIONS EXCEDIDA. FECHAR E REABRIR A CONEXÃO
'''        'MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
'''        Conn.Close
'''        Conn.Open
'''        Resume
'''    Else
''''        If RS Is Nothing Then
''''           If RS.State = 1 Then
''''              RS.Close
''''           End If
''''        End If
'''        tcs.Normal
'''        tcs.Select
'''        'Conn.RollbackTrans
'''
'''        Open App.path & "\Controles\GeoSanLog.txt" For Append As #1
'''        Print #1, Now & " " & strUser & " " & Versao_Geo & " - frmCadastroRamalAgua - Private Sub cmdConfirmar_Click() - Local Num: " & intlocalerro & " - Erro Num: " & Err.Number & " - " & Err.Description
'''        Close #1
'''
'''        MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
'''        Unload Me
'''
'''    End If
'''End Sub

Private Sub cmdConfirmar_Click()

   On Error GoTo Trata_Erro
   Dim intlocalerro As Integer
   Dim rsCria As ADODB.Recordset
   Dim a As Integer
   Dim cgeo As New clsGeoReference
   Dim X As Double
   Dim Y As Double
   Dim str As String
   
   Dim strNroL As String 'NÚMERO DA LIGACAO
   Dim strInsc As String 'NUMERO DA INSCRIÇÃO
   Dim strTipo As String 'TIPO DA LIGACAO
   Dim strCons As String 'CONSUMO DA LIGACAO
   Dim strEcon As String 'QUANTIDADE DE ECONOMIAS NA LIGAÇÃO
   Dim strHidr As String
   
   Set rsCria = New ADODB.Recordset 'recordset utilizado para criar o regitro na tabela
   Dim strNroLigaSel As String
   'Conn.BeginTrans
   
   Conn.Close
   
   Conn.Open
   
    va = "NRO_LIGACAO"
         ve = "CLASSIFICACAO_FISCAL"
         vi = "COD_LOGRADOURO"
         vo = "TIPO"
         vu = "ECONOMIAS"
         vc = "HIDROMETRADO"
         vd = "OBJECT_ID_"
         vf = "CONSUMO_LPS"
          ve = "INSCRICAO_LOTE"
   If object_id_ramal = "" Then 'NOVO RAMAL
      
        str = strUser & Now
        
        'O CAMPO ID DA TABELA É POPULADA COM A AUTO NUMERAÇÃO DA TABELA (TRIGGER NO ORACLE)
 
        Set rsCria = New ADODB.Recordset
      '  rsCria.Open TB_Ramais, Conn, adOpenDynamic, adLockOptimistic
      
      
      If frmCanvas.TipoConexao = 1 Then
      
    rsCria.Open TB_Ramais, Conn, adOpenKeyset, adLockOptimistic
    
    
    ElseIf frmCanvas.TipoConexao = 2 Then
     rsCria.Open "Select * from " + TB_Ramais + "", Conn, adOpenKeyset, adLockOptimistic
    Else
    If count2 <> 10 Then
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
Set rsCria = New ADODB.Recordset

     rsCria.Open "Select * from " + """" + TB_Ramais + """", conexao, adOpenDynamic, adLockOptimistic
    
    End If
        
      rsCria.AddNew
    If frmCanvas.TipoConexao <> 4 Then
    rsCria.Fields("OBJECT_ID_").value = str
      rsCria.Fields("OBJECT_ID_TRECHO").value = "0"
    Else
     ' rsCria.Fields("OBJECT_ID_").value = str
      rsCria.Fields("DATA_LOG").value = str
      rsCria.Fields("OBJECT_ID_TRECHO").value = "0"
      '  rsCria.Fields("OBJECT_ID_").value = "'aaa'"
       ' rsCria.Fields("OBJECT_ID_TRECHO").value = "'ss'"
 End If
        rsCria.Update
        
         If frmCanvas.TipoConexao = 4 Then
         rsCria.Close
          Set rsCria = New ADODB.Recordset
          Dim strin As String
          strin = "Select " + """" + "ID" + """" + "from " + """" + TB_Ramais + """" + " WHERE " + """" + "DATA_LOG" + """" + " ='" + str + "'"
         rsCria.Open strin, Conn, adOpenDynamic, adLockOptimistic
         End If
        
        object_id_ramal = rsCria.Fields("ID").value
        rsCria.Close
        
        tcs.object_id = object_id_ramal
        tcs.saveOnMemory
        tcs.SaveInDatabase
        
        'MsgBox tcs.getCurrentLayer
        
        
        tdbramais.setCurrentLayer TB_Ramais '"RAMAIS_AGUA"
        'RETORNA EM X E Y A COORDENADA DO FINAL DA LINHA
        tdbramais.getPointOfLine 0, object_id_ramal, 1, X, Y
        
        'INSERE PONTO NO FINAL DA LINHA
        tdbramais.addPoint object_id_ramal, X, Y

        
        tdbramais.getPointOfLine 0, object_id_ramal, 0, X, Y
        'tdb.setCurrentLayer cgeo.GetLayerOperation(tcs.getCurrentLayer, 1)
        
        Object_id_trecho = ramal_Object_id_trecho 'VARIÁVEL RAMAL_OBJECT_ID_TRECHO CARREGADA NO TCANVAS ON_CLICK
        
        Dim rsRamal As ADODB.Recordset
        Set rsRamal = New ADODB.Recordset
        ve = "TB_RAMAIS"
        vi = "ID"
        
        intlocalerro = 4
        If frmCanvas.TipoConexao <> 4 Then
        rsRamal.Open "SELECT * FROM  " & TB_Ramais & "  WHERE ID = " & "'" & object_id_ramal & "'", Conn, adOpenKeyset, adLockOptimistic, adCmdText
        Else
        
        rsRamal.Open "SELECT * FROM  " + """" + TB_Ramais + """" + " WHERE " + """" + "ID" + """" + " = " & "'" & object_id_ramal & "'", conexao, adOpenDynamic, adLockOptimistic, adCmdText
        
        End If
        If rsRamal.EOF = False Then
            
            rsRamal.Fields("Distancia_Lado").value = IIf(IsNumeric(txtDistanciaLado), txtDistanciaLado, 0)
            rsRamal.Fields("Distancia_Testada").value = IIf(IsNumeric(txtDistanciaTestada), txtDistanciaTestada, 0)
            rsRamal.Fields("Profundidade_RAMAL").value = IIf(IsNumeric(txtProfundidade), txtProfundidade, 0)
            rsRamal.Fields("Comprimento_Ramal").value = IIf(IsNumeric(txtComprimentoRamal), txtComprimentoRamal, 0)
             
            For i = 1 To lvLigacoes.ListItems.count
               If lvLigacoes.ListItems(i).Checked = True Then
               'MsgBox lvLigacoes.ListItems(1).Tag
                   rsRamal!COD_LOGRAD = Val(lvLigacoes.ListItems(1).Tag) 'PEGA O PRIMEIRO LOGRADOURO SELECIONADO NA LISTA
                   Exit For
               End If
            Next
        
           If optDesconhecido Then rsRamal.Fields("posicionamento_lote").value = 1
           If optEsquerdo Then rsRamal.Fields("posicionamento_lote").value = 2
           If optCentro Then rsRamal.Fields("posicionamento_lote").value = 3
           If optDireito Then rsRamal.Fields("posicionamento_lote").value = 4
           
           rsRamal!Object_id_ = object_id_ramal
           rsRamal!Object_id_trecho = Object_id_trecho
           rsRamal!USUARIO_LOG = strUser
           rsRamal!DATA_LOG = Format(Now, "DD/MM/YY HH:MM")
           
           rsRamal.Update
        
        Else
            Exit Sub
            'Conn.execute "DELETE FROM RAMAIS_AGUA WHERE object_id_ = '" & str & "'"
        End If
        rsRamal.Close
        
        
      strNroLigaSel = ""
        
      For a = 1 To lvLigacoes.ListItems.count
          If lvLigacoes.ListItems(a).Checked Then 'PARA CADA ITEM SELECIONADO NA LISTA
             If strNroLigaSel <> "" Then
                strNroLigaSel = strNroLigaSel & ",'" & lvLigacoes.ListItems(a).SubItems(1) & "'"
             Else
                strNroLigaSel = "'" & lvLigacoes.ListItems(a).SubItems(1) & "'"
             End If
          End If
      Next
        
      
If strNroLigaSel <> "" Then
      If frmCanvas.TipoConexao <> 4 Then 'SQL
       Set rs = New ADODB.Recordset
  
         str = "SELECT NRO_LIGACAO, CLASSIFICACAO_FISCAL, COD_LOGRADOURO, "
         str = str & "TIPO, ECONOMIAS, HIDROMETRADO FROM " & TB_comercial & " WHERE NRO_LIGACAO IN (" & strNroLigaSel & ")"
           rs.Open str, Conn, adOpenDynamic, adLockOptimistic
         Else 'Postgres
   va = "NRO_LIGACAO"
         ve = "CLASSIFICACAO_FISCAL"
         vi = "COD_LOGRADOURO"
         vo = "TIPO"
         vu = "ECONOMIAS"
         vc = "HIDROMETRADO"
         vd = "OBJECT_ID_"
         vf = "CONSUMO_LPS"
         ' ve = "INSCRICAO_LOTE"
    Set rs = New ADODB.Recordset
  
   
        
         str = "SELECT " + """" + va + """" + ", " + """" + ve + """" + ", " + """" + vi + """" + ", "
         str = str & "" + """" + vo + """" + ", " + """" + vu + """" + ", " + """" + vc + """" + " FROM " + """" + TB_comercial + """" + " WHERE " + """" + va + """" + " IN (" & strNroLigaSel & ")"
        rs.Open str, conexao, adOpenDynamic, adLockOptimistic
         End If
       
          'WritePrivateProfileString "A", "A", str, App.path & "\DEBUG.INI"
          
          
          
       '  rs.Open str, Conn, adOpenDynamic, adLockOptimistic
          

          If rs.EOF = False Then
            Do While Not rs.EOF
               
               
               strNroL = Trim(rs!NRO_LIGACAO)                 'NÚMERO DA LIGACAO
               
               If Trim(rs!CLASSIFICACAO_FISCAL) <> "" Then strInsc = Trim(rs!CLASSIFICACAO_FISCAL) Else strInsc = ""  'NUMERO DA INSCRIÇÃO
               If rs!tipo <> "" Then strTipo = Trim(rs!tipo) Else strTipo = ""           'TIPO DA LIGACAO
               If rs!ECONOMIAS <> "" Then strEcon = Trim(rs!ECONOMIAS) Else strEcon = ""  'QUANTIDADE DE ECONOMIAS NA LIGAÇÃO
               If UCase(rs!HIDROMETRADO) = "SIM" Or UCase(rs!HIDROMETRADO) = "NAO" Then strHidr = LCase(rs!HIDROMETRADO) Else strHidr = "" 'ARMAZENA EM LETRA MINÚSCULA
 If frmCanvas.TipoConexao <> 4 Then
               str = "INSERT INTO " & TB_Ligacoes & " (OBJECT_ID_,NRO_LIGACAO,INSCRICAO_LOTE,TIPO,HIDROMETRADO,ECONOMIAS,CONSUMO_LPS) "
               If strTipo = "" Then
               strTipo = 0
               End If
               
               str = str & "VALUES ('" & object_id_ramal & "','" & strNroL & "','" & strInsc & "','" & strTipo & "','" & strHidr & "'," & strEcon & ",0)"
              ' MsgBox str
              Else
              
          va = "NRO_LIGACAO"
         ve = "INSCRICAO_LOTE"
         vi = "COD_LOGRADOURO"
         vo = "TIPO"
         vu = "ECONOMIAS"
         vc = "HIDROMETRADO"
         vd = "OBJECT_ID_"
         vf = "CONSUMO_LPS"
         If strTipo = "" Then
         strTipo = 0
         End If
             If strHidr = "" Then
         strHidr = 0
         End If
              str = "INSERT INTO " + """" + TB_Ligacoes + """" + " (" + """" + vd + """" + "," + """" + va + """" + "," + """" + ve + """" + "," + """" + vo + """" + "," + """" + vc + """" + "," + """" + vu + """" + "," + """" + vf + """" + ") "
               str = str & "VALUES ('" & object_id_ramal & "','" & strNroL & "','" & strInsc & "','" & strTipo & "','" & strHidr & "','" & strEcon & "','0')"
              
              
              End If
               Conn.execute (str)
               rs.MoveNext

            Loop
         End If
         rs.Close
        
      End If
        
      If TB_Ligacoes = "RAMAIS_AGUA_LIGACAO" Then
         SubInsereFicticios
      End If
      
      
        
      
    Else ' RAMAL EXISTENTE
    
     If frmCanvas.TipoConexao = 4 Then
     If count3 <> 10 Then


    
    
      Dim conexao2 As New ADODB.connection
      
      mSERVIDOR = ReadINI("CONEXAO", "SERVIDOR", App.path & "\CONTROLES\GEOSAN.ini")
mPORTA = ReadINI("CONEXAO", "PORTA", App.path & "\CONTROLES\GEOSAN.ini")
mBANCO = ReadINI("CONEXAO", "BANCO", App.path & "\CONTROLES\GEOSAN.ini")
mUSUARIO = ReadINI("CONEXAO", "USUARIO", App.path & "\CONTROLES\GEOSAN.ini")
Senha = ReadINI("CONEXAO", "SENHA", App.path & "\CONTROLES\GEOSAN.ini")
nStr = frmCanvas.FunDecripta(Senha)
decriptada = frmCanvas.Senha
  strConn = "DRIVER={PostgreSQL Unicode}; DATABASE=" + mBANCO + "; SERVER=" + mSERVIDOR + "; PORT=" + mPORTA + "; UID=" + mUSUARIO + "; PWD=" + nStr + "; ByteaAsLongVarBinary=1;"

    conexao2.Open strConn
      
      count3 = 10
      End If
      End If
      Set rs = New ADODB.Recordset
       ' ve = "TB_RAMAIS"
        'vi = "ID"
      If frmCanvas.TipoConexao <> 4 Then 'SQL

      rs.Open "SELECT * FROM  " & TB_Ramais & "  WHERE OBJECT_ID_ ='" & object_id_ramal & "'", Conn, adOpenKeyset, adLockOptimistic
      Else
      rs.Open "SELECT * FROM  " + """" + TB_Ramais + """" + "  WHERE " + """" + vd + """" + " ='" & object_id_ramal & "'", conexao2, adOpenDynamic, adLockOptimistic
      End If
      
      If rs.EOF = False Then
         rs.Fields("Distancia_Lado").value = IIf(IsNumeric(txtDistanciaLado), txtDistanciaLado, 0)
         rs.Fields("Distancia_Testada").value = IIf(IsNumeric(txtDistanciaTestada), txtDistanciaTestada, 0)
         rs.Fields("Profundidade_RAMAL").value = IIf(IsNumeric(txtProfundidade), txtProfundidade, 0)
         rs.Fields("Comprimento_Ramal").value = IIf(IsNumeric(txtComprimentoRamal), txtComprimentoRamal, 0)
         
         For i = 1 To lvLigacoes.ListItems.count
             If lvLigacoes.ListItems(i).Checked = True Then
                 If lvLigacoes.ListItems(1).Tag <> "" Then
                     rs.Fields("cod_lograd").value = lvLigacoes.ListItems(1).Tag 'PEGA O PRIMEIRO LOGRADOURO SELECIONADO NA LISTA
                 End If
                 Exit For
             End If
         Next
         
         If optDesconhecido Then rs.Fields("posicionamento_lote").value = 1
         If optEsquerdo Then rs.Fields("posicionamento_lote").value = 2
         If optCentro Then rs.Fields("posicionamento_lote").value = 3
         If optDireito Then rs.Fields("posicionamento_lote").value = 4
          
         rs.Fields("USUARIO_LOG").value = strUser
         rs.Fields("DATA_LOG").value = Format(Now, "DD/MM/YY HH:MM") ' & "/" & Format(Now, "MM") & "/" & Format(Now, "YY") & " " & Format(Now, "HH") & ":" & Format(Now, "MM")
         
         rs.Update
         rs.Close
      
      End If

      intlocalerro = 6
       If frmCanvas.TipoConexao <> 4 Then 'SQL
      Conn.execute "DELETE FROM " & TB_Ligacoes & " WHERE OBJECT_ID_ = '" & object_id_ramal & "'"
      Else
      Conn.execute "DELETE FROM " + """" + TB_Ligacoes + """" + " WHERE " + """" + vd + """" + " = '" & object_id_ramal & "'"
      End If
      
      strNroLigaSel = ""
        
      For a = 1 To lvLigacoes.ListItems.count
          If lvLigacoes.ListItems(a).Checked Then 'PARA CADA ITEM SELECIONADO NA LISTA
             If mid(lvLigacoes.ListItems(a).SubItems(1), 1, Len(object_id_ramal)) <> object_id_ramal Then
               
               If strNroLigaSel <> "" Then
                  strNroLigaSel = strNroLigaSel & ",'" & lvLigacoes.ListItems(a).SubItems(1) & "'"
               Else
                  strNroLigaSel = "'" & lvLigacoes.ListItems(a).SubItems(1) & "'"
               End If
             
             End If
          End If
      Next
      
      If strNroLigaSel <> "" Then
         
           If frmCanvas.TipoConexao <> 4 Then
         str = "SELECT NRO_LIGACAO, CLASSIFICACAO_FISCAL, COD_LOGRADOURO, "
         str = str & "TIPO, ECONOMIAS, HIDROMETRADO FROM NXGS_V_LIG_COMERCIAL WHERE NRO_LIGACAO IN (" & strNroLigaSel & ")"
         Else
         Dim fg, fh As String
         fg = "COD_LOGRADOURO"
           fh = "NXGS_V_LIG_COMERCIAL"
           
              va = "NRO_LIGACAO"
         vm = "CLASSIFICACAO_FISCAL"
         vi = "COD_LOGRADOURO"
         vo = "TIPO"
         vu = "ECONOMIAS"
         vc = "HIDROMETRADO"
         vd = "OBJECT_ID_"
         vf = "CONSUMO_LPS"
           
           
         str = "SELECT " + """" + va + """" + ", " + """" + vm + """" + ", " + """" + vi + """" + ", "
         str = str & """" + vo + """" + ", " + """" + vu + """" + ", " + """" + vc + """" + " FROM " + """" + fh + """" + " WHERE " + """" + va + """" + " IN (" & strNroLigaSel & ")"
         

        Set rs = New ADODB.Recordset
         
         End If
         
 If frmCanvas.TipoConexao <> 4 Then
         'rs.Open str, Conn, adOpenDynamic, adLockReadOnly, adCmdText 'RECORDSET OBTEM INFORMAÇÕES PARA O INSERT
         rs.Open str, Conn, adOpenDynamic, adLockOptimistic
         Else
         rs.Open str, conexao2, adOpenDynamic, adLockOptimistic
         
         End If
        va = "NRO_LIGACAO"
         ve = "CLASSIFICACAO_FISCAL"
         vi = "COD_LOGRADOURO"
         vo = "TIPO"
         vu = "ECONOMIAS"
         vc = "HIDROMETRADO"
         vd = "OBJECT_ID_"
         vf = "CONSUMO_LPS"
        ve = "INSCRICAO_LOTE"
          
         If rs.EOF = False Then
            Do While Not rs.EOF
            
               strNroL = Trim(rs!NRO_LIGACAO)                                             'NÚMERO DA LIGACAO
               
               If Trim(rs!CLASSIFICACAO_FISCAL) <> "" Then strInsc = Trim(rs!CLASSIFICACAO_FISCAL) Else strInsc = ""  'NUMERO DA INSCRIÇÃO
               If rs!tipo <> "" Then strTipo = Trim(rs!tipo) Else strTipo = ""           'TIPO DA LIGACAO
               If rs!ECONOMIAS <> "" Then strEcon = Trim(rs!ECONOMIAS) Else strEcon = ""  'QUANTIDADE DE ECONOMIAS NA LIGAÇÃO
               If UCase(rs!HIDROMETRADO) = "SIM" Or UCase(rs!HIDROMETRADO) = "NAO" Then strHidr = LCase(rs!HIDROMETRADO) Else strHidr = "" 'ARMAZENA EM LETRA MINÚSCULA
                If frmCanvas.TipoConexao <> 4 Then
               str = "INSERT INTO " & TB_Ligacoes & " (OBJECT_ID_,NRO_LIGACAO,INSCRICAO_LOTE,TIPO,HIDROMETRADO,ECONOMIAS,CONSUMO_LPS) "
               str = str & "VALUES ('" & object_id_ramal & "','" & strNroL & "','" & strInsc & "','" & strTipo & "','" & strHidr & "','" & strEcon & "','0')"
               Else
               
                  If strTipo = "" Then
         strTipo = 0
         End If
             If strHidr = "" Then
         strHidr = 0
         End If
         
         If strInsc = "" Then
         strInsc = 0
         End If
               
               
                str = "INSERT INTO " + """" + TB_Ligacoes + """" + " (" + """" + vd + """" + "," + """" + va + """" + "," + """" + ve + """" + "," + """" + vo + """" + "," + """" + vc + """" + "," + """" + vu + """" + "," + """" + vf + """" + ") "
               str = str & "VALUES ('" & object_id_ramal & "','" & strNroL & "','" & strInsc & "','" & strTipo & "','" & strHidr & "','" & strEcon & "','0')"
  
         
               
               
               End If
               
               Conn.execute (str)
               rs.MoveNext
            
            Loop
         End If
         rs.Close
      End If

      If TB_Ligacoes = "RAMAIS_AGUA_LIGACAO" Then
         SubInsereFicticios
      End If

   End If
   
   Set rs = Nothing
   tcs.plotView
   Unload Me
   

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    ElseIf Err.Number = -2147418113 Then ' Erro geral de rede
        
        PrintErro CStr(Me.Name), "Private Sub cmdConfirmar_Click()", CStr(Err.Number), CStr(Err.Description), True
        End
    
    ElseIf Err.Number = -2147417848 Then ' automation error
    
        PrintErro CStr(Me.Name), "Private Sub cmdConfirmar_Click()", CStr(Err.Number), CStr(Err.Description), True
        End
    
    ElseIf Err.Number = -2147467259 Or mid(Err.Description, 1, 9) = "ORA-03114" Then 'PERDA DE CONEXÃO BANCO SQL OU ORACLE
        
        PrintErro CStr(Me.Name), "Private Sub cmdConfirmar_Click()", CStr(Err.Number), CStr(Err.Description), True
        'End
    
    ElseIf Err.Number = -2147168227 Then ' MAX TRANSACTIONS EXCEDIDA. FECHAR E REABRIR A CONEXÃO
        'MsgBox "Um posssível erro foi identificado:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo GeoSanLog.txt com informações desta ocorrência.", vbInformation
        Conn.Close
        Conn.Open
        Resume
    Else
        
        tcs.Normal
        tcs.Select
        
        PrintErro CStr(Me.Name), "Private Sub cmdConfirmar_Click()", CStr(Err.Number), CStr(Err.Description), True
        Unload Me
        
    End If
   ' MsgBox Err.Description
    
End Sub

Private Sub SubInsereFicticios()
   
   Dim str As String
   Dim strCons As String 'CONSUMO DA LIGACAO
      
   'INSERINDO RAMAL FICTÍCIO SE ESTE FOI SELECIONADO
   If CInt(Me.txtQtd.Text) > 0 Then

      'CAPTURA O CONSUMO DIGITADO E CONVERTE SE NECESSÁRIO
      If CDbl(Me.txtConsumoFicticia.Text) > 0 Then
         If Me.optMetroCubico.value = True Then
             'SE FOR METRO CUBICO, CONVERTE PARA LITROS POR SEGUNDO
             strCons = Replace(Me.txtConsumoFicticia.Text, ".", ",") * 0.00038580246
         Else
             strCons = Replace(Me.txtConsumoFicticia.Text, ".", ",")
         End If
      Else
         strCons = 0
      End If
      
      strCons = Replace(strCons, ",", ".") 'troca a vírgula pelo ponto
      
      For i = 0 To CInt(Me.txtQtd.Text) - 1
         
        If frmCanvas.TipoConexao <> 4 Then 'SQL
         
         
         str = "INSERT INTO " & TB_Ligacoes & " (OBJECT_ID_,NRO_LIGACAO,INSCRICAO_LOTE,TIPO,HIDROMETRADO,ECONOMIAS,CONSUMO_LPS) "
         str = str & "VALUES ('" & object_id_ramal & "','999" & object_id_ramal & i & "','999" & object_id_ramal & i & "','FICTÍCIA','nao','1','" & strCons & "')"
            Else
            Dim jo As String
            Dim ja As String
            Dim je As String
            Dim ji As String
            Dim ju As String
            Dim jb As String
            Dim jc As String
           
            
            jo = "OBJECT_ID_"
            ja = "NRO_LIGACAO"
            je = "INSCRICAO_LOTE"
            ji = "TIPO"
            ju = "HIDROMETRADO"
            jb = "ECONOMIAS"
            jc = "CONSUMO_LPS"
            
            
             str = "INSERT INTO " + """" + TB_Ligacoes + """" + " (" + """" + jo + """" + "," + """" + ja + """" + "," + """" + je + """" + "," + """" + ji + """" + "," + """" + ju + """" + "," + """" + jb + """" + "," + """" + jc + """" + ") "
         str = str & "VALUES ('" & object_id_ramal & "','999" & object_id_ramal & i & "','999" & object_id_ramal & i & "','FICTÍCIA','nao','1','" & strCons & "')"
            End If
            
         Conn.execute (str)
         
      Next
      
   End If

End Sub
'Carrega os dados das ligações para apresentar na caixa de diálogo para o usuário
Private Sub CarregaLigacoes()
    Dim intlocalerro As Integer
    On Error GoTo Trata_Erro
    Dim NRO_LIGACOES As String, INSCRICOES_LOTES As String, msg As String
    Dim rsAssociados As ADODB.Recordset, str As String, itmx As ListItem, a As Integer, Qtde As Integer
    'RECUPERA TODAS AS INSCRICOES DE TODOS LOTE
    'Inicia o processo de recuperar as inscrições/ligações associadas aos lotes. Não temos mais lotes. Hoje as inscrições são armazenadas nos nós nas extremidades dos ramais
    str = GetQueryProcess(3)
    INSCRICOES_LOTES = "''"
    
    If Trim(object_id_lote) = "" Then
        str = Replace(str, "@OBJECT_ID_", "''")
    Else
        str = Replace(str, "@OBJECT_ID_", object_id_lote)
    End If
    
    intlocalerro = 1
    Set rs = Conn.execute(str)
    
    'Vefifica se existe alguma ligação associada ao polígono do lote. Não utilizamos mais polígonos de lotes. Claro que não terá nenhuma
    While Not rs.EOF
        
        If INSCRICOES_LOTES = "''" Then
            INSCRICOES_LOTES = "'" & rs(0).value & "'"
        Else
            INSCRICOES_LOTES = INSCRICOES_LOTES & ",'" & rs(0).value & "'"
        End If
    
        rs.MoveNext
    Wend
    
    rs.Close
    'Finaliza o processo de recuperar as inscrições/ligações associadas aos lotes.
    
    'Inicia a recuperação de todas as ligações associadas ao nó na extremidade do ramal
    'Configura o código de erro para que se ocorra um erro, seja indicado o local onde ele ocorreu (2)
    intlocalerro = 2
    vi = "OBJECT_ID_"
    Set rsAssociados = New ADODB.Recordset
    
    'Seleciona todas as colunas da tabela onde estão os números das ligações de água ou esgoto, associadas aos ramais. Pode existir mais de uma ligação associada a um mesmo ramal
    If frmCanvas.TipoConexao <> 4 Then
        'No caso de ser banco de dados SQLServer ou Oracle
        str = "SELECT * FROM " & TB_Ligacoes & " WHERE OBJECT_ID_ = '" & object_id_ramal & "'"
    Else
        'No caso de se banco de dados Postgres
        str = "SELECT * FROM " + """" + TB_Ligacoes + """" + " WHERE " + """" + vi + """" + " = '" & object_id_ramal & "'"
    End If

    'Abre a conexão com o banco de dados para ver os números das ligações
    rsAssociados.Open str, Conn, adOpenForwardOnly, adLockReadOnly
    NRO_LIGACOES = "''"
   
    'Enquanto houverem ligações associadas ao ramal/nó selecionado
    If rsAssociados.EOF = False Then
        While Not rsAssociados.EOF
            'Se o número da ligação for nulo, ou seja se for a primera vez que estiver lendo a primeira ligação
            If NRO_LIGACOES = "''" Then
                'Armazena ela no vetor de ligações
                NRO_LIGACOES = "'" & rsAssociados.Fields("NRO_LIGACAO").value & "'"
            'Se for a segunda ou próxima ligação, acrescenta a mesma ao vetor de ligações
            Else
                NRO_LIGACOES = NRO_LIGACOES & ",'" & rsAssociados.Fields("NRO_LIGACAO").value & "'"
            End If
            rsAssociados.MoveNext
        Wend
        'Pronto! Agora tenho um vetor com o(s) numero(s) de todas as ligações que eu selecionei (nó) e estão ligadas ao determinado ramal
        'Configura o código de erro para que se ocorra um erro, seja indicado o local onde ele ocorreu (3), indicando que agora iremos para outra fase
        intlocalerro = 3
        'Recupera a querie junto a vista ou tabela que contem a lista de todas as ligações do município, com os dados das mesmas. Geralmente vinda do banco de dados comercial
        str = GetQueryProcess(2)
        'Faz a substituição no local dos números das ligações pelo vetor contendo todos os números de ligações selecionados
        str = Replace(str, "@NRO_LIGACAO", NRO_LIGACOES)
        'Faz a substituição no local com a classificação fiscal da prefeitura com as inscrições dos lotes. Não tem nada para substituir
        str = Replace(str, "@CLASSIFICACAO_FISCAL", INSCRICOES_LOTES)
        'No caso de cadastro de ramais de esgoto, ele necessita informar que a tabela de esgotos é outra
        If TB_Ramais = "RAMAIS_ESGOTO" Then
            str = Replace(str, "NXGS_V_LIG_COMERCIAL", "NXGS_V_LIG_COMERCIAL_E")
        End If
        'MsgBox "ARQUIVO DEBUG SALVO"
        ' WritePrivateProfileString "A", "A", str, App.path & "\DEBUG.INI"
        'MsgBox str
        'Set rs = ConnSec.execute(str)
        'Abre a conexão com o banco de dados para ler as ligações do vetor de ligações
        Set rs = New ADODB.Recordset
        rs.Open str, Conn, adOpenDynamic, adLockOptimistic
        'Enquanto existirem ligações
        While Not rs.EOF
            'CARREGA NO FORM TODAS AS LIGAÇÕES CADASTRADAS
            With lvLigacoes
                'Set itmx = .ListItems.Add(, , rs.Fields("NRO_LIGACAO").value)
                'itmx.SubItems(1) = IIf(IsNull(rs.Fields("CLASSIFICACAO_FISCAL").value), "", rs.Fields("CLASSIFICACAO_FISCAL").value)
                Set itmx = lvLigacoes.ListItems.Add(, , rs.Fields("CLASSIFICACAO_FISCAL").value)
                itmx.SubItems(1) = IIf(IsNull(rs.Fields("NRO_LIGACAO").value), "", rs.Fields("NRO_LIGACAO").value)
                itmx.SubItems(2) = IIf(IsNull(rs.Fields("ENDERECO").value), "", rs.Fields("ENDERECO").value)
                itmx.SubItems(3) = IIf(IsNull(rs.Fields("CONSUMIDOR").value), "", rs.Fields("CONSUMIDOR").value)
                itmx.SubItems(4) = IIf(IsNull(rs.Fields("TIPO").value), "", rs.Fields("TIPO").value)
                rsAssociados.Filter = "NRO_LIGACAO='" & rs.Fields("NRO_LIGACAO").value & "'"
                If Not rsAssociados.EOF Then itmx.Checked = True
                    itmx.Tag = IIf(IsNull(rs.Fields("codlograd").value), "", rs.Fields("codlograd").value)
            End With
            rs.MoveNext
        Wend
        rs.Close
    End If
    
    'CARREGA AS LIGAÇÕES FICTÍCIAS CASO SEJA RAMAIS DE AGUA
    If TB_Ligacoes = "RAMAIS_AGUA_LIGACAO" Then
        If frmCanvas.TipoConexao <> 4 Then
            str = "SELECT * FROM " & TB_Ligacoes & " WHERE NRO_LIGACAO IN (" & NRO_LIGACOES & ") AND TIPO = 'FICTÍCIA'"
        Else
            vi = "NRO_LIGACAO"
            vo = "TIPO"
            If frmCanvas.TipoConexao = 4 Then
                If NRO_LIGACOES = "''" Then
                    NRO_LIGACOES = "'0'"
                End If
            End If
            str = "SELECT * FROM " + """" + TB_Ligacoes + """" + " WHERE " + """" + vi + """" + " IN (" & NRO_LIGACOES & ") AND " + """" + vo + """" + " = 'FICTÍCIA'"
        End If
        Set rs = Conn.execute(str)
        If rs.EOF = False Then
            While Not rs.EOF
               Set itmx = lvLigacoes.ListItems.Add(, , rs.Fields("INSCRICAO_LOTE").value)
               itmx.SubItems(1) = IIf(IsNull(rs.Fields("NRO_LIGACAO").value), "", rs.Fields("NRO_LIGACAO").value)
               itmx.SubItems(2) = ""                                                                                'IIf(IsNull(rs.Fields("ENDERECO").value), "", rs.Fields("ENDERECO").value)
               itmx.SubItems(3) = ""                                                                                'IIf(IsNull(rs.Fields("CONSUMIDOR").value), "", rs.Fields("CONSUMIDOR").value)
               'rsAssociados.Filter = "NRO_LIGACAO='" & rs.Fields("NRO_LIGACAO").value & "'"
               'If Not rsAssociados.EOF Then itmx.Checked = True
               itmx.Checked = True
               itmx.SubItems("4") = "FICTÍCIA"
               Me.txtQtd.Text = CInt(Me.txtQtd.Text) + 1
               Me.optLitrosSegundo.value = True
               Me.txtConsumoFicticia.Text = IIf(IsNull(rs.Fields("CONSUMO_LPS").value), "0.00", rs.Fields("CONSUMO_LPS").value)
               rs.MoveNext
            Wend
            rs.Close
        End If
    End If
    intlocalerro = 4
    rsAssociados.Close
    'CarregaLigacoes_err:
    '   msg = "object_id_lote: '" & object_id_lote & "'"
    '   msg = msg & vbCrLf & "INSCRICOES_LOTES: " & INSCRICOES_LOTES
    '   MsgBox Err.Description & vbCrLf & msg & vbCrLf & str

Trata_Erro:
If Err.Number = 0 Or Err.Number = 20 Then
    Resume Next
Else
    PrintErro CStr(Me.Name), "Private Sub carregaLigacoes()", CStr(Err.Number), CStr(Err.Description), True
End If

End Sub

Private Sub cmdConsultarLigacoes_Click()
On Error GoTo Trata_Erro
   Dim str As String
   Dim j As Integer
   Dim list As ListItem
   frmConsumoLote.lvLigacoes.ListItems.Clear
   For j = 1 To Me.lvLigacoes.ListItems.count
            
      'frmConsumoLote.lvLigacoes.ListItems.Add (1)
      
      If mid(Me.lvLigacoes.ListItems.Item(j), 1, 3) <> "999" Then 'NÃO INSERE FICTÍCIA (COMEÇAM COM 999)
      
         Set list = frmConsumoLote.lvLigacoes.ListItems.Add(, , Me.lvLigacoes.ListItems.Item(j))
            
         list.SubItems(1) = Me.lvLigacoes.ListItems(j).SubItems(1)
         
         list.SubItems(2) = Me.lvLigacoes.ListItems(j).SubItems(2)
         
         list.SubItems(3) = Me.lvLigacoes.ListItems(j).SubItems(3)
      
      End If
   Next
    
   frmConsumoLote.Show (1)
   
Trata_Erro:
If Err.Number = 0 Or Err.Number = 20 Then
   Resume Next
Else
   PrintErro CStr(Me.Name), "cmdConsultarLigacoes_Click()", CStr(Err.Number), CStr(Err.Description), True
End If
   
    
End Sub

Private Sub cmdFechar_Click()
    
    If cmdFechar.Caption = "Fechar" Then
        If object_id_ramal = "" Then
           tcs.Normal
           tcs.Select
        End If
        Unload Me
    Else
        If MsgBox("Deseja cancelar as alterações realizadas?", vbQuestion + vbDefaultButton2 + vbYesNo, "Cancelar") = vbYes Then
             If object_id_ramal = "" Then
                tcs.Normal
                tcs.Select
             End If
             Unload Me
        End If
    End If
   
End Sub

Private Sub cmdPesquisaLigacoes_Click()
   frmCadastroRamalFiltro.init Me, tcs, object_id_ramal
End Sub

Private Function Verifica_Ligacao(index As Integer) As Boolean
On Error GoTo Trata_Erro
   Dim a As Integer, UltimoEndereco As String, rs As ADODB.Recordset, str As String
   For a = 1 To lvLigacoes.ListItems.count
      If lvLigacoes.ListItems(a).Checked Then
         If UltimoEndereco <> "" Then
            'VERIFICA SE TODOS OS LOGRADOUROS SAO O MESMO
            If lvLigacoes.ListItems(a).Tag <> UltimoEndereco Then
               MsgBox "Não possível vincular em um mesmo ramal ligações de logradouros diferentes", vbExclamation
               Exit Function
            End If
            
         End If
         UltimoEndereco = lvLigacoes.ListItems(a).Tag
      End If
   Next
   'VEVIFICA SE A LIGAÇÃO JÁ ESTÁ VINCULADA EM OUTRO RAMAL
   str = GetQueryProcess(10)
   str = Replace(str, "@LAYER", tcs.getCurrentLayer)
   str = Replace(str, "@OBJECT_ID_RAMAL", object_id_ramal)
   str = Replace(str, "@NRO_LIGACAO", lvLigacoes.ListItems(index).Text)
   Set rs = Conn.execute(str)
   If Not rs.EOF Then
      MsgBox "Esta ligação embora esteja vinculada este lote, já está vincula a outro ramal:" & rs(0).value, vbExclamation
      Exit Function
   End If
   rs.Close
   Set rs = Nothing
   Verifica_Ligacao = True

Trata_Erro:
If Err.Number = 0 Or Err.Number = 20 Then
   Resume Next
Else
   
   PrintErro CStr(Me.Name), "Private Sub Verifica_Ligacao", CStr(Err.Number), CStr(Err.Description), True

End If

End Function


Private Sub Form_Activate()
   Me.lvLigacoes.SetFocus
   
End Sub





Private Sub Label4_Click()
   frmCadastroRamalAutoLote.Show 1
End Sub

Private Sub lvLigacoes_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    cmdFechar.Caption = "Cancelar"
End Sub


Private Sub optCentro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdFechar.Caption = "Cancelar"
End Sub


Private Sub optDireito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdFechar.Caption = "Cancelar"
End Sub

Private Sub optEsquerdo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdFechar.Caption = "Cancelar"
End Sub

Private Sub txtComprimentoRamal_KeyPress(KeyAscii As Integer)
    cmdFechar.Caption = "Cancelar"
    testa_letra (KeyAscii)
    KeyAscii = intKeyAscii
End Sub

Private Sub txtConsumoFicticia_LostFocus()
   If IsNumeric(txtConsumoFicticia.Text) = False Then
      MsgBox "Somente números são aceitos para consumo médio.", , ""
      txtConsumoFicticia.SetFocus
   End If
End Sub

Private Sub txtDistanciaLado_KeyPress(KeyAscii As Integer)
    cmdFechar.Caption = "Cancelar"
    testa_letra (KeyAscii)
    KeyAscii = intKeyAscii
End Sub

Private Sub txtDistanciaTestada_KeyPress(KeyAscii As Integer)
    cmdFechar.Caption = "Cancelar"
    testa_letra (KeyAscii)
    KeyAscii = intKeyAscii
End Sub


Private Sub txtInscricao_Change()
    Me.optInscricao.value = True
End Sub

Private Sub txtEndereco_Change()
    Me.optEndereço.value = True
End Sub

Private Sub txtConsumidor_Change()
    Me.optConsumidor.value = True
End Sub

Private Function testa_letra(ByVal KeyAscii As Integer)
'FUNÇÃO QUE VERIFICA SE O CARACTERE DIGITADO É NUMÉRICO OU BACKSPACE OU VIRGULA, CASO CONTRARIO, ANULA
    If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 44 Then
        intKeyAscii = KeyAscii
    Else
        intKeyAscii = 0
    End If
End Function

Private Sub txtNumLigacao_KeyPress(KeyAscii As Integer)
    Me.optNumLigacao.value = True
    testa_letra (KeyAscii)
    KeyAscii = intKeyAscii
End Sub

Private Sub txtProfundidade_KeyPress(KeyAscii As Integer)
    cmdFechar.Caption = "Cancelar"
    testa_letra (KeyAscii)
    KeyAscii = intKeyAscii
End Sub


Private Sub txtQtd_LostFocus() 'LIGAÇÃO FICTÍCIA
   If txtQtd.Text = "" Then
      txtQtd.Text = "0"
   End If
   
   If IsNumeric(txtQtd.Text) = False Then
      txtQtd.SetFocus
   End If
   
End Sub

Private Sub UpDown2_Change() 'LIGAÇÃO FICTÍCIA
   
     txtQtd.Text = UpDown2.value

End Sub



