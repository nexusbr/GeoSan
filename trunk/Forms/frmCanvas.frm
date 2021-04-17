VERSION 5.00
Object = "{18576B0E-A129-4A50-9930-59E18A6FE5E1}#1.0#0"; "TeComCanvas.dll"
Object = "{87AC6DA5-272D-40EB-B60A-F83246B1B8D7}#1.0#0"; "TeComDatabase.dll"
Object = "{9AB389E7-EAED-4DBF-941D-EB86ED1F9A76}#1.0#0"; "TeComConnection.dll"
Object = "{EE78E37B-39BE-42FA-80B7-E525529739F7}#1.0#0"; "TeComViewDatabase.dll"
Begin VB.Form frmCanvas 
   Caption         =   "Mapa"
   ClientHeight    =   5952
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   7680
   Icon            =   "frmCanvas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5952
   ScaleWidth      =   7680
   WindowState     =   2  'Maximized
   Begin TECOMCANVASLibCtl.TeCanvas TCanvas 
      Height          =   2415
      Left            =   3360
      OleObjectBlob   =   "frmCanvas.frx":08CA
      TabIndex        =   8
      Top             =   600
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Left            =   6615
      Top             =   5400
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ajustar Escala"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
      Begin VB.CommandButton cmdConfEscala 
         Caption         =   "OK"
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtEscala 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraRedes 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tamanho das Redes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   780
      Visible         =   0   'False
      Width           =   2175
      Begin VB.TextBox txtRede2 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtRede1 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Segunda"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   760
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Primeira"
         Height          =   255
         Left            =   150
         TabIndex        =   6
         Top             =   270
         Width           =   1935
      End
   End
   Begin VB.Timer TimerSetWorld 
      Interval        =   100
      Left            =   6180
      Top             =   5220
   End
   Begin TeComConnectionLibCtl.TeAcXConnection TeAcXConnection1 
      Left            =   6360
      OleObjectBlob   =   "frmCanvas.frx":08FE
      Top             =   3360
   End
   Begin TECOMDATABASELibCtl.TeDatabase TeDatabaseRamais 
      Left            =   720
      OleObjectBlob   =   "frmCanvas.frx":0922
      Top             =   5400
   End
   Begin TECOMDATABASELibCtl.TeDatabase TeDatabase3 
      Left            =   720
      OleObjectBlob   =   "frmCanvas.frx":0946
      Top             =   4680
   End
   Begin TECOMDATABASELibCtl.TeDatabase TeDatabase2 
      Left            =   480
      OleObjectBlob   =   "frmCanvas.frx":096A
      Top             =   3720
   End
   Begin TECOMDATABASELibCtl.TeDatabase TeDatabase1 
      Left            =   480
      OleObjectBlob   =   "frmCanvas.frx":098E
      Top             =   2640
   End
   Begin TeComViewDatabaseLibCtl.TeViewDatabase TeViewDatabase2 
      Left            =   4200
      OleObjectBlob   =   "frmCanvas.frx":09B2
      Top             =   4560
   End
   Begin TeComViewDatabaseLibCtl.TeViewDatabase TeViewDatabase1 
      Left            =   4080
      OleObjectBlob   =   "frmCanvas.frx":09D6
      Top             =   3600
   End
End
Attribute VB_Name = "frmCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim geo As Variant
Dim tipoDeConexao As String

Dim nStr As String
Public Position_X As Double, Position_Y As Double
Private mUserName As String, ViewName As String
Private xmin, ymin, xmax, ymax, LastEvent As TypeGeometryEvent
Dim Tc As New clsTerraConfig, Tr As New clsTerraLib, LastDocument As String, tempo As Date
Dim lastGpsObjIdPointSelected As String                                                              'guarda o object_id do último ponto GPS selecionado
Dim CLIQUE_RAMAL As Integer
Dim intQtdLinhasNaCoordenada As Integer
Dim postg As Integer
Dim postg2 As Integer
Dim postg3 As Integer
Dim postg4 As Integer
Dim postg5 As Integer
Dim xOld As Double
Dim yOld As Double
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String
Dim f As String
Dim g As String
Dim h As String
Dim i As String
Dim layeratual As String
Dim selec As Long
Dim mPROVEDOR As String
Dim mSERVIDOR As String
Dim mPORTA As String
Dim mBANCO As String
Dim mUSUARIO As String
Dim senha2 As String
Dim decriptada As String
Dim user As String
Dim con As New ADODB.connection
Dim strConn As String
Dim count2 As Integer
Dim conexao As New ADODB.connection
Dim cConsumidor As New clsConsumidorControler
Dim object_id_consumidorSelecionado As Long         'referente ao cadastro de ligações com GPS em campo é a seleção do consumidor que será utilizado para ligar ao ramal
Dim object_id_redeSelecionada As Long               'referente ao trecho de rede selecionado ao qual serão criados os ramais e ligados os consumidores
Dim object_id_ramalAddConsumerSelecionado As Long   'referente ao ramal que foi selecionado, para que os outros consumidores sejam adicionados ao mesmo
Dim object_id_ramalAddConsumerConsumidorSelecionado As Long 'referente ao consumidor selecionado pelo usuário, o qual será adicionado ao ramal já existente

'Constantes utilizadas na função ConvertTwipsToPixels para converter pixel para milímetro
Const WU_LOGPIXELSX = 88
Const WU_LOGPIXELSY = 90

'Insere uma ligação de água do coletor de campo obtida pelo GPS do celular, no ramal selecionado
'
'
'
Private Sub InsereLigacaoNoRamalSelecionado(object_id_ramalSelecionado As Long, object_id_consumidoreSelecionado As Long)
        Dim debugCodigoErro As String
        Dim rsInsereRamaisAguaLigacao As ADODB.Recordset
        Dim rsApagaLigacaoGpsCadastrada As ADODB.Recordset
        Dim strAbreConexaoInsereRamaisAguaLigacao As String
        Dim volumeFaturado As Double
        Dim numeroDaLigacaoComDV As String
        Dim strApagaGeometriaDaLigacaoGpsCadastrada As String
        Dim strInsereLigacao As String
        Dim strApagaLigacaoGpsCadastrada As String
        Dim dataCadastroLigacao As String
        
        '1 - Insere em RAMAIS_AGUA_LIGACAO a ligação selecionada pelo usuário
        On Error GoTo Transacao_Erro
        Conn.BeginTrans
        debugCodigoErro = "0"
        Set rsInsereRamaisAguaLigacao = New ADODB.Recordset
        strAbreConexaoInsereRamaisAguaLigacao = "SELECT NRO_LIGACA, VOL_FATURA FROM NXGS_V_LIG_COMERCIAL_GPS WHERE object_id_272 = " + CStr(object_id_consumidoreSelecionado)
        debugCodigoErro = "1 - Select: " & strAbreConexaoInsereRamaisAguaLigacao
        rsInsereRamaisAguaLigacao.Open strAbreConexaoInsereRamaisAguaLigacao, Conn, adOpenKeyset, adLockOptimistic, adCmdText
        
        '2 - Inicia a atualização de RAMAIS_AGUA_LIGACAO com todos os dados
        dataCadastroLigacao = Now
        If rsInsereRamaisAguaLigacao.EOF = False Then                         'Tem que encontrar a linha em RAMAIS_AGUA que acabou de ser inserida
            volumeFaturado = IIf(IsNull(rsInsereRamaisAguaLigacao!VOL_FATURA), 0, rsInsereRamaisAguaLigacao!VOL_FATURA)
            numeroDaLigacaoComDV = rsInsereRamaisAguaLigacao!NRO_LIGACA
        End If
        strInsereLigacao = "INSERT INTO RAMAIS_AGUA_LIGACAO (OBJECT_ID_,NRO_LIGACAO,CONSUMO_LPS, DATA_LOG, USUARIO_LOG) "
        strInsereLigacao = strInsereLigacao & "VALUES ('" & object_id_ramalSelecionado & "','" & numeroDaLigacaoComDV & "', " & volumeFaturado & ", '" & dataCadastroLigacao & "' , '" & strUser & "' )"
        Conn.execute (strInsereLigacao)
        rsInsereRamaisAguaLigacao.Close

        '3 - Apaga NX GPS
        debugCodigoErro = "2"
        Set rsApagaLigacaoGpsCadastrada = New ADODB.Recordset
        strApagaLigacaoGpsCadastrada = "SELECT NRO_LIGACA FROM NXGS_V_LIG_COMERCIAL_GPS WHERE object_id_272 = " + CStr(object_id_consumidorSelecionado)
        rsApagaLigacaoGpsCadastrada.Open strAbreConexaoInsereRamaisAguaLigacao, Conn, adOpenKeyset, adLockOptimistic, adCmdText
        strApagaLigacaoGpsCadastrada = "DELETE FROM NXGS_V_LIG_COMERCIAL_GPS WHERE object_id_272 = " + CStr(object_id_consumidorSelecionado)
        Conn.execute (strApagaLigacaoGpsCadastrada)
        debugCodigoErro = "3"
        strApagaGeometriaDaLigacaoGpsCadastrada = "DELETE FROM POINTS272 WHERE object_id = " + CStr(object_id_consumidoreSelecionado)
        Conn.execute (strApagaGeometriaDaLigacaoGpsCadastrada)
        rsApagaLigacaoGpsCadastrada.Close
        Conn.CommitTrans
        On Error GoTo Trata_Erro
        debugCodigoErro = "4"
        
        TCanvas.plotView
        Exit Sub

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        ErroUsuario.Registra "frmCanvas", "InsereRamalLigacaoGPS - Código Erro: " & debugCodigoErro, CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
    Exit Sub
    
Transacao_Erro:
    Conn.RollbackTrans
    'Conn.Close
    ErroUsuario.Registra "frmCanvas", "InsereRamalLigacaoGPS - Código Erro: " & debugCodigoErro, CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    Exit Sub
End Sub

'Insere um ramal e a ligação do mesmo, para os consumidores levantados com GPS
'
'
'
Private Sub InsereRamalLigacaoGPS(object_id_ligacaoGPS As Long, object_id_rede As Long)
        On Error GoTo Trata_Erro
        Dim debugCodigoErro As String
        Dim objIdRamalTemporarioDoUsuario As String
        Dim dataCadastroRamal As String
        Dim rsAdicionaNovoAtributoRamalAgua As ADODB.Recordset
        Dim rsConsultaMaximoGeoId As ADODB.Recordset
        Dim rsAtualizaDadosDoRamal As ADODB.Recordset
        Dim rsAtualizaDadosLinhaRamal As ADODB.Recordset
        Dim rsInsereRamaisAguaLigacao As ADODB.Recordset
        Dim rsApagaLigacaoGpsCadastrada As ADODB.Recordset
        Dim idUnicoRamaisAgua As Long
        Dim retornoLinhaPerpendicular As Long
        Dim pontoSobreALinha As Long
        Dim comprimentoDoRamal As Double
        Dim coord_x_NaLinha As Double
        Dim coord_y_NaLinha As Double
        Dim retornoPontoGPS As Long
        Dim coordX_pontoGPS As Double
        Dim coordY_pontoGPS As Double
        Dim stringObject_id_ligacaoGPS As String
        Dim stringObject_id_rede As String
        Dim linhaRamalX(1) As Double, linhaRamalY(1) As Double
        Dim retornoAdicionaLinhaRamal As Long
        Dim strAdicionaNovoRamalAgua As String
        Dim strConsultaMaxGeoIdLinhaRamal As String
        Dim strAtualizaObjIdRamal As String
        Dim strAtualizaDadosDaLinhaRamal As String
        Dim strAbreConexaoInsereRamaisAguaLigacao As String
        Dim strInsereRamaisAguaLigacao As String
        Dim strInsereLigacao As String
        Dim strApagaLigacaoGpsCadastrada As String
        Dim strApagaGeometriaDaLigacaoGpsCadastrada As String
        Dim retornoPontoDoRamalInserido As Long
        Dim geomIdRamal As Long
        Dim volumeFaturado As Double
        Dim numeroDaLigacaoComDV As String
        Dim dataCadastroLigacao As String
        
        '1 - Adiciona linha a tabela de atributos de ramais de água
        debugCodigoErro = "0"
        stringObject_id_ligacaoGPS = CStr(object_id_ligacaoGPS)
        stringObject_id_rede = CStr(object_id_rede)
        dataCadastroRamal = Now
        objIdRamalTemporarioDoUsuario = strUser & dataCadastroRamal
        strAdicionaNovoRamalAgua = "RAMAIS_AGUA"
        Set rsAdicionaNovoAtributoRamalAgua = New ADODB.Recordset
        On Error GoTo Transacao_Erro
        Conn.BeginTrans
        rsAdicionaNovoAtributoRamalAgua.Open strAdicionaNovoRamalAgua, Conn, adOpenKeyset, adLockOptimistic
        rsAdicionaNovoAtributoRamalAgua.AddNew                                                          'Cria uma nova linha na tabela RAMAIS_AGUA
        rsAdicionaNovoAtributoRamalAgua.Fields("OBJECT_ID_").value = objIdRamalTemporarioDoUsuario      'Atualiza o OBJECT_ID_ da tabela RAMAIS_AGUA com o nome do usuário, data e hora (temporáriamente)
        rsAdicionaNovoAtributoRamalAgua.Fields("OBJECT_ID_TRECHO").value = object_id_rede               'Atualiza o OBJECT_ID do trecho de rede de água em RAMAIS_AGUA com zero (temporariamente)
        rsAdicionaNovoAtributoRamalAgua.Fields("DATA_LOG").value = dataCadastroRamal
        rsAdicionaNovoAtributoRamalAgua!USUARIO_LOG = strUser                                           'Salva o nome do usuário
        rsAdicionaNovoAtributoRamalAgua.Update                                                          'Atualiza no banco de dados a tabela RAMAIS_AGUA
        rsAdicionaNovoAtributoRamalAgua.Close
        Conn.CommitTrans
        debugCodigoErro = debugCodigoErro + " " + objIdRamalTemporarioDoUsuario

        '2 - Atualiza os dados de RAMAIS_AGUA inclusive com o OBJECT_ID do trecho de rede e OBJECT_ID do ramal
        On Error GoTo Trata_Erro
        Set rsAtualizaDadosDoRamal = New ADODB.Recordset
        strAtualizaObjIdRamal = "SELECT * FROM RAMAIS_AGUA WHERE OBJECT_ID_ = '" + objIdRamalTemporarioDoUsuario + "'"
        debugCodigoErro = "1" + " - " + strAtualizaObjIdRamal
        On Error GoTo Transacao_Erro
        Conn.BeginTrans
        rsAtualizaDadosDoRamal.Open strAtualizaObjIdRamal, Conn, adOpenKeyset, adLockOptimistic, adCmdText
        If rsAtualizaDadosDoRamal.EOF = False Then                         'Tem que encontrar a linha em RAMAIS_AGUA que acabou de ser inserida
            idUnicoRamaisAgua = rsAtualizaDadosDoRamal.Fields("ID").value  'Obtem o ID da nova linha inserida em RAMAIS_AGUA (que foi gerado automaticamente, para poder depois localizar este ramal e colocar os demais dados na tabela de atributos dele
            debugCodigoErro = debugCodigoErro + " idUnicoRamaisAgua: " + CStr(idUnicoRamaisAgua)
            rsAtualizaDadosDoRamal!Object_id_ = CStr(idUnicoRamaisAgua)          'ID autonumérico da tabela Ramais                            'Agora coloca o OBJECT_ID do ramal correto, o anterior tinha o nome do usuário-data-hora
            debugCodigoErro = debugCodigoErro + " object_id_: " + rsAtualizaDadosDoRamal!Object_id_
            rsAtualizaDadosDoRamal!USUARIO_LOG = strUser                   'Salva o nome do usuário
            debugCodigoErro = debugCodigoErro + " strUser: " + rsAtualizaDadosDoRamal!USUARIO_LOG
            rsAtualizaDadosDoRamal.Update
        End If
        rsAtualizaDadosDoRamal.Close
        Conn.CommitTrans
        On Error GoTo Trata_Erro
        
        '3 - Adiciona geometria da linha de ramal de água
        debugCodigoErro = "2"
        retornoPontoGPS = TeDatabase1.setCurrentLayer("NXGS_V_LIG_COMERCIAL_GPS")
        TeDatabase1.getCenterGeometry 0, stringObject_id_ligacaoGPS, TypeGeometry.points, coordX_pontoGPS, coordY_pontoGPS
        retornoPontoGPS = TeDatabase1.setCurrentLayer("WATERLINES")
        retornoLinhaPerpendicular = TeDatabase1.getMinimumDistance(0, stringObject_id_rede, TypeGeometry.lines, coordX_pontoGPS, coordY_pontoGPS, comprimentoDoRamal, pontoSobreALinha, coord_x_NaLinha, coord_y_NaLinha)
        If retornoLinhaPerpendicular = 1 And coordX_pontoGPS > 0 And coordY_pontoGPS > 0 Then
            linhaRamalX(0) = coord_x_NaLinha
            linhaRamalY(0) = coord_y_NaLinha
            linhaRamalX(1) = coordX_pontoGPS
            linhaRamalY(1) = coordY_pontoGPS
            retornoPontoGPS = TeDatabase1.setCurrentLayer("RAMAIS_AGUA")
            retornoAdicionaLinhaRamal = TeDatabase1.addLine(idUnicoRamaisAgua, linhaRamalX(0), linhaRamalY(0), 2)
        Else
            debugCodigoErro = "2 - stringObject_id_ligacaoGPS = " + stringObject_id_ligacaoGPS + " retornoPerpendicular = " + CStr(retornoLinhaPerpendicular) + " coordX_pontoGPS = " + CStr(coordX_pontoGPS) + " coordY_pontoGPS = " + CStr(coordY_pontoGPS)
            GoTo Trata_Erro 'por algum motivo em poucos casos acontece de ele não pergar a coordenada do ponto do ramal (extremidade do mesmo) foi colocado isso para poder identificar o que está acontecendo
        End If
        
        '4 - Adiciona geometria do ponto ao ramail
        debugCodigoErro = "3"
        If coordX_pontoGPS > 0 And coordY_pontoGPS > 0 Then         'só adiciona o ponto do ramal se existir a coordenada. Este if é para identificar o que pode estar acontecendo em poucos casos quando se cadastra o ramal automaticamente
            retornoPontoDoRamalInserido = TeDatabase1.addPoint(idUnicoRamaisAgua, coordX_pontoGPS, coordY_pontoGPS)
        Else
            debugCodigoErro = "3 - stringObject_id_ligacaoGPS = " + stringObject_id_ligacaoGPS + " retornoPerpendicular = " + CStr(retornoLinhaPerpendicular) + " coordX_pontoGPS = " + CStr(coordX_pontoGPS) + " coordY_pontoGPS = " + CStr(coordY_pontoGPS)
            GoTo Trata_Erro
        End If
        
        '5 - Insere em RAMAIS_AGUA_LIGACAO a ligação selecionada pelo usuário
        On Error GoTo Transacao_Erro
        Conn.BeginTrans
        debugCodigoErro = "4"
        dataCadastroLigacao = Now
        Set rsInsereRamaisAguaLigacao = New ADODB.Recordset
        strAbreConexaoInsereRamaisAguaLigacao = "SELECT NRO_LIGACA, VOL_FATURA FROM NXGS_V_LIG_COMERCIAL_GPS WHERE object_id_272 = " + CStr(object_id_consumidorSelecionado)
        debugCodigoErro = "5 - Select: " & strAbreConexaoInsereRamaisAguaLigacao
        rsInsereRamaisAguaLigacao.Open strAbreConexaoInsereRamaisAguaLigacao, Conn, adOpenKeyset, adLockOptimistic, adCmdText
        'Inicia a atualização de RAMAIS_AGUA com todos os dados
        If rsInsereRamaisAguaLigacao.EOF = False Then                         'Tem que encontrar a linha em RAMAIS_AGUA que acabou de ser inserida
            volumeFaturado = IIf(IsNull(rsInsereRamaisAguaLigacao!VOL_FATURA), 0, rsInsereRamaisAguaLigacao!VOL_FATURA)
            numeroDaLigacaoComDV = rsInsereRamaisAguaLigacao!NRO_LIGACA
        End If
        strInsereLigacao = "INSERT INTO RAMAIS_AGUA_LIGACAO (OBJECT_ID_,NRO_LIGACAO,CONSUMO_LPS, DATA_LOG, USUARIO_LOG) "
        strInsereLigacao = strInsereLigacao & "VALUES ('" & idUnicoRamaisAgua & "','" & numeroDaLigacaoComDV & "', " & volumeFaturado & ", '" & dataCadastroLigacao & "' , '" & strUser & "' )"
        Conn.execute (strInsereLigacao)
        rsInsereRamaisAguaLigacao.Close

        '6 - Apaga NX GPS
        debugCodigoErro = "6"
        Set rsApagaLigacaoGpsCadastrada = New ADODB.Recordset
        strApagaLigacaoGpsCadastrada = "SELECT NRO_LIGACA FROM NXGS_V_LIG_COMERCIAL_GPS WHERE object_id_272 = " + CStr(object_id_consumidorSelecionado)
        rsApagaLigacaoGpsCadastrada.Open strAbreConexaoInsereRamaisAguaLigacao, Conn, adOpenKeyset, adLockOptimistic, adCmdText
        strApagaLigacaoGpsCadastrada = "DELETE FROM NXGS_V_LIG_COMERCIAL_GPS WHERE object_id_272 = " + CStr(object_id_consumidorSelecionado)
        Conn.execute (strApagaLigacaoGpsCadastrada)
        debugCodigoErro = "7"
        strApagaGeometriaDaLigacaoGpsCadastrada = "DELETE FROM POINTS272 WHERE object_id = " + CStr(object_id_consumidorSelecionado)
        Conn.execute (strApagaGeometriaDaLigacaoGpsCadastrada)
        rsApagaLigacaoGpsCadastrada.Close
        Conn.CommitTrans
        On Error GoTo Trata_Erro
        debugCodigoErro = "8"
        
        TCanvas.plotView
        Exit Sub

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        ErroUsuario.Registra "frmCanvas", "InsereRamalLigacaoGPS - Código Erro: " & debugCodigoErro, CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
    Exit Sub
    
Transacao_Erro:
    Conn.RollbackTrans
    'Conn.Close
    ErroUsuario.Registra "frmCanvas", "InsereRamalLigacaoGPS - Código Erro: " & debugCodigoErro, CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    Exit Sub
End Sub


' Converte twips para pixels. No TeCanvas.width a medida é em twips e é necessário converter para pixels para que possa ser
' configurada a tolerância do snap. A tolerância do snap é medida em pixels.
' Esta função retorna o numero de pixels equivalentes.
'
' lngTwips - numero de twips
' lngDirection - 0 = horizontal, outro valor = vertical, se as medidas estão sendo realizadas na horizontal ou vertical
'
Function ConvertTwipsToPixels(lngTwips As Long, lngDirection As Long) As Long
   'Handle to device
   Dim lngDC As Long
   Dim lngPixelsPerInch As Long
   
   Const nTwipsPerInch = 1440
   lngDC = GetDC(0)
   If (lngDirection = 0) Then       'Horizontal
      lngPixelsPerInch = GetDeviceCaps(lngDC, WU_LOGPIXELSX)
   Else                             'Vertical
      lngPixelsPerInch = GetDeviceCaps(lngDC, WU_LOGPIXELSY)
   End If
   lngDC = ReleaseDC(0, lngDC)
   ConvertTwipsToPixels = (lngTwips / nTwipsPerInch) * lngPixelsPerInch
End Function
'teste do código de mover ramais junto com a rede
'
'
Private Sub teste()
    Dim moveRamal As New CCoordIniRamalDistTrecho   'classe para obter a nova coordenada do ramal que foi movido
    Dim linha As CLine2D                            'nova linha do novo ramal após a movimentação do trecho de rede
    Dim objIdTrecho As String                       'objId do trecho inicial antes de ser movido
    Dim objIdRamal As String                        'objId do ramal inicial antes de ser movido
    Dim novoObjIdTrecho As String                   'objId do trecho final depois da movimentação pelo usuário
    Dim objIDsRamais As New CObtemObjIDsRamais
    Dim listObjIDsRamais() As String
    Dim i As Integer
    Dim distIniRamalAntes As Double                 'distância do início do ramal antes de tanto o trecho quanto o ramal serem movidos
    Dim distIniRamalDepois As Double                'distância do início do ramal depois de tanto o trecho quanto o ramal serem movidos
    Dim distEquiv As New CDistanciaEquivalente      'classe para obter a distância do início do ramal ao início do trecho após movido os mesmos
    
    'obtem o objId do trecho a ser movido
    'obtem o objId do trecho movido
    ' Call objIDsRamais.getObjIDs("14064", TeDatabase4, listObjIDsRamais)                             'obtem todos os objIDs dos ramais que estão ligados ao trecho de rede que está sendo movido
    For i = 0 To UBound(listObjIDsRamais)           'enquanto existirem ramais
        'distIniRamalAntes = moveRamal.distancia("14064", listObjIDsRamais(i), TeDatabase4)          'obtem a distância do início do ramal antes de tanto o trecho quanto o ramal serem movidos
        distIniRamalDepois = distEquiv.distanciaRamalDepoisMovido(123.22, 134.44, distIniRamalAntes)
        Set linha = moveRamal.coordsRamal(distIniRamalDepois, "14064", cGeoDatabase.geoDatabase)                 'obtem as novas coordenadas inicial e final do ramal movido após mover o trecho de rede
        'Set linha = moveRamal.linha(objIdTrecho, listObjIDsRamais(i), novoObjIdTrecho, TeDatabase4) 'obtem a nova linha do ramal movido
        'apaga a geometria do ramal
        'desenha a geometria do novo ramal
        'atualiza o objId do novo ramal com o mesmo que o anterior para ligar aos atributos existentes
        i = i + 1
    Next
    
        

    'fim
End Sub
Public Static Function TipoConexao() As String

tipoDeConexao = typeconnection
TipoConexao = tipoDeConexao

End Function



Public Static Function POST() As Integer


POST = postg

End Function

Public Static Function POST2(po3 As Integer) As Integer


postg = po3

End Function

Public Static Function POSTA() As Integer


POSTA = postg2

End Function

Public Static Function POST2A(po2 As Integer) As Integer


postg2 = po2

End Function




Public Static Function POSTB() As Integer


POSTB = postg3

End Function

Public Static Function POST2B(po3 As Integer) As Integer


postg3 = po3

End Function


Public Static Function POSTC() As Integer


POSTC = postg4

End Function

Public Static Function POST2C(po4 As Integer) As Integer


postg4 = po4

End Function


Public Static Function POSTD() As Integer


POSTD = postg5

End Function

Public Static Function POST2D(po5 As Integer) As Integer


postg5 = po5

End Function




Public Static Function Senha() As String

Senha = nStr

End Function
' Esta função é automaticamente sempre chamada quando é solicitada a inicialização de um novo canvas
'
' Conn - conexão realizada
' username - nome do usuário que logou no GeoSan
'
Public Function Init(Conn As ADODB.connection, username As String) As Boolean
    On Error GoTo Trata_Erro
    Dim rs As ADODB.Recordset
    Dim linha As Integer
    
    tipoDeConexao = typeconnection
    If typeconnection <> POSTGRESQL Then
        'se não for Postgresss
        TeViewDatabase1.username = username
        TeViewDatabase1.Provider = typeconnection
        TeViewDatabase1.connection = Conn
        'LoadThemes
        'user = username
        'con = Conn
        TeDatabase1.username = username
        TeDatabase1.Provider = typeconnection
        TeDatabase1.connection = Conn
        TeDatabase2.Provider = typeconnection
        TeDatabase2.connection = Conn
        TeDatabase3.Provider = typeconnection
        TeDatabase3.connection = Conn
        'cGeoDatabase.configura Conn, typeconnection, username  'retirada esta inicialização e movida para a rotina main do arquivo Global.bas para poder realizar consultas com TeDatabase em antes mesmo de ter aberto a vista
        TeDatabaseRamais.Provider = typeconnection              'inicializa a conexão para poder inserir um ramal
        TeDatabaseRamais.connection = Conn
        TCanvas.Provider = typeconnection
        TCanvas.user = username
        TCanvas.connection = Conn ' SE DER ERRO AQUI, VERIFICAR A VERSÃO DA TECOM INSTALADA NA MÁQUINA
        'ViewName = TeViewDatabase1.getActiveView
        If Tc.GetWorldByUser(username, xmin, ymin, xmax, ymax, typeconnection) Then
            TCanvas.setProjection "WATERLINES"
            TCanvas.setWorld CDbl(xmin), CDbl(ymin), CDbl(xmax), CDbl(ymax)
        End If
        '***************************************************
        'incluido em 16/01/2009 - Jonathas
        'Recurso Tecom 3.2 - Configuração do tamanho do ponto do acordo com o zoom
        If ReadINI("MAPA", "FIXAR_ICONE", App.path & "\CONTROLES\GEOSAN.INI") = "SIM" Then
            TCanvas.fixedPoint = True
        Else
            TCanvas.fixedPoint = False
        End If
        '***************************************************
        'DEIXA COMO CURRENT LAYER O RAMAIS AGUA CASO SEJA USUÁRIO VISUALIZADOR
        Set rs = New ADODB.Recordset
        rs.Open "SELECT USRLOG, USRFUN FROM SYSTEMUSERS WHERE USRLOG = '" & strUser & "' ORDER BY USRLOG", Conn, adOpenDynamic, adLockOptimistic
        If rs.EOF = False Then
            If rs!UsrFun = 4 Then 'VISUALIZADOR
                MsgBox "Layer corrente: RAMAIS_AGUA", vbInformation, ""
                TCanvas.setCurrentLayer "RAMAIS_AGUA"
            End If
        End If
        rs.Close
        Me.Show
        TCanvas.plotView            'mostra o mapa na tela
        TCanvas.snapOn = 1          'liga o snap
        mUserName = username
        'Para saber quantos canvas estão abertos...
        If FrmMain.Tag = "" Then
            FrmMain.Tag = 0
        Else
            FrmMain.Tag = Int(FrmMain.Tag) + 1
        End If
    Else
        'se for Postgress
        Dim mPROVEDOR As String
        Dim mSERVIDOR As String
        Dim mPORTA As String
        Dim mBANCO As String
        Dim mUSUARIO As String
        Dim Senha As String
        Dim decriptada As String
        mSERVIDOR = ReadINI("CONEXAO", "SERVIDOR", App.path & "\CONTROLES\GEOSAN.ini")
        mPORTA = ReadINI("CONEXAO", "PORTA", App.path & "\CONTROLES\GEOSAN.ini")
        mBANCO = ReadINI("CONEXAO", "BANCO", App.path & "\CONTROLES\GEOSAN.ini")
        mUSUARIO = ReadINI("CONEXAO", "USUARIO", App.path & "\CONTROLES\GEOSAN.ini")
        Senha = ReadINI("CONEXAO", "SENHA", App.path & "\CONTROLES\GEOSAN.ini")
        nStr = FunDecripta(Senha)
        decriptada = nStr
        Call WriteINI("CONEXAO", "USER", username, App.path & "\CONTROLES\GEOSAN.INI")
        TeAcXConnection1.Open mUSUARIO, decriptada, mBANCO, mSERVIDOR, mPORTA
        TeViewDatabase1.username = username
        TeViewDatabase1.Provider = typeconnection
        TeViewDatabase1.connection = TeAcXConnection1.objectConnection_
        ' TeViewDatabase1.addView ("TESTE2000")
        'TeViewDatabase1.addTheme("WATERLINES", "TESTE2000", "WATERLINES") = True
        TeDatabase1.username = username
        TeDatabase1.Provider = typeconnection
        TeDatabase1.connection = TeAcXConnection1.objectConnection_
        TeDatabase2.Provider = typeconnection
        TeDatabase2.connection = TeAcXConnection1.objectConnection_
        TeDatabase3.Provider = typeconnection
        TeDatabase3.connection = TeAcXConnection1.objectConnection_
        cGeoDatabase.geoDatabase.username = username
        cGeoDatabase.geoDatabase.Provider = typeconnection
        cGeoDatabase.geoDatabase.connection = TeAcXConnection1.objectConnection_
        TeDatabaseRamais.Provider = typeconnection                    'inicializa a conexão para pode inserir um ramal
        TeDatabaseRamais.connection = TeAcXConnection1.objectConnection_
        TCanvas.Provider = typeconnection 'Provider 4 = PostgreSQL
        TCanvas.user = username
        TCanvas.connection = TeAcXConnection1.objectConnection_       'É nessa parte que é setada a conexão com
                                                                      'a TeComConnection. Isso é válido para
                                                                      'todas as outras TeComs. Porém, quando for
                                                                      'realizar as querys de atributos, as mesmas
                                                                      'devem ser feitas pela conexão ado do vb.
                                                                      'Se quiser trabalhar com transação, deve-se
                                                                      'abrir a transação da conexão ado e da
                                                                      'TeComConnection. Ex:
                                                                      'conexao.BeginTrans
                                                                      'TeConnection.beginTransaction
                                                                      'O mesmo vale para o Commit e para o
                                                                      'Rollback.
        'TCanvas.saveOnMemory
        'TCanvas.SaveInDatabase
        If Tc.GetWorldByUser(username, xmin, ymin, xmax, ymax, typeconnection) Then
            TCanvas.setProjection "WATERLINES"
            TCanvas.setWorld CDbl(xmin), CDbl(ymin), CDbl(xmax), CDbl(ymax)
        End If
        '***************************************************
        'incluido em 16/01/2009 - Jonathas
        'Recurso Tecom 3.2 - Configuração do tamanho do ponto do acordo com o zoom
        If ReadINI("MAPA", "FIXAR_ICONE", App.path & "\CONTROLES\GEOSAN.INI") = "SIM" Then
            TCanvas.fixedPoint = True
        Else
            TCanvas.fixedPoint = False
        End If
        '***************************************************
        'DEIXA COMO CURRENT LAYER O RAMAIS AGUA CASO SEJA USUÁRIO VISUALIZADOR
        Set rs = New ADODB.Recordset
        Dim stringconexao As String
        Dim a As String
        Dim b As String
        Dim c As String
        Dim d As String
        Dim e As String
        a = "USRLOG"
        b = "USRFUN"
        c = "SYSTEMUSERS"
        stringconexao = "Select " + """" + a + """" + "," + """" + b + """" + " from " + """" + c + """" + " where " + """" + a + """" + "=" + "'" & strUser & "' ORDER BY " + """" + a + """" + ""
        ' rs.Open stringconexao, Conn, adOpenDynamic, adLockReadOnly
        rs.Open stringconexao, Conn, adOpenDynamic, adLockOptimistic
        If rs.EOF = False Then
            If rs!UsrFun = 4 Then 'VISUALIZADOR
                MsgBox "Layer corrente: RAMAIS_AGUA", vbInformation, ""
                TCanvas.setCurrentLayer "RAMAIS_AGUA"
            End If
        End If
        rs.Close
        Me.Show
        TCanvas.plotView
        TCanvas.snapOn = 1                  'liga o snap
        mUserName = username
        'Para saber quantos canvas estão abertos...
        If FrmMain.Tag = "" Then
            FrmMain.Tag = 0
        Else
            FrmMain.Tag = Int(FrmMain.Tag) + 1
        End If
    End If
    Set cConsumidor.tcs = TCanvas                        'aqui foi feito diferente, para os controles de métodos e evendos do TeCanvas sejam executados diretamente dentro da classe e na classe clsTerralib, fazendo assim uma separação e melhor orientação a objetos
    Set cConsumidor.tdbcon = TeDatabase2                 'seta e TeDatabase2 passa ser valor para a variável cConsumidor.tdbcon
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        ErroUsuario.Registra "frmCanvas", "init", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
        End
    End If
End Function

Private Sub cmdConfEscala_Click()
On Error GoTo Trata_Erro
    If Trim(txtEscala.Text) <> "" And IsNumeric(txtEscala.Text) Then
        TCanvas.setScale Int(txtEscala.Text)
    Else
        MsgBox "Digite um valor numérico para a escala.", vbInformation, "Atenção!"
        txtEscala.SetFocus
    End If
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
   Else
    
      PrintErro CStr(Me.Name), "cmdConfEscala_Click()", CStr(Err.Number), CStr(Err.Description), True
      
   End If
End Sub

Private Sub Form_Activate()
   'TeViewDatabase1.setActiveView
   'TCanvas.v ViewName
   'TeViewDatabase1.connection = Conn

   LoadThemes
   
   LoadToolsBar
   TCanvas_onEndSELECT
End Sub
' Rotina responsável por verificar qual ícone foi selecionada
' ativa o comando selecionado, caso seja desenho de rede, zoom área, etc. Para o programa ficar sabendo o que ele está fazendo
'
'
'
Private Sub LoadToolsBar()
   Dim a As Integer
   For a = 1 To FrmMain.tbToolBar.Buttons.count
      If FrmMain.tbToolBar.Buttons.Item(a).Style = tbrCheck Then FrmMain.tbToolBar.Buttons(a).value = tbrUnpressed
   Next
   Select Case Tr.TerraEvent
      Case tg_SelectObject
         FrmMain.tbToolBar.Buttons("kselection").value = tbrPressed
      Case tg_ZoomArea
         FrmMain.tbToolBar.Buttons("kzoomarea").value = tbrPressed
      Case tg_Pan
         FrmMain.tbToolBar.Buttons("kpan").value = tbrPressed
      Case tg_DrawNetWorkline
         FrmMain.tbToolBar.Buttons("kdrawnetworkline").value = tbrPressed
         'limpa todos os itens editados em memória, as geometrias das listas temporárias e geometrias a serem removidas do banco de dados
         TCanvas.clearEditItens (2)         'limpa linhas
         TCanvas.clearEditItens (4)         'limpa pontos
      Case tg_DrawNetWorkNode
         FrmMain.tbToolBar.Buttons("kinsertnetworknode").value = tbrPressed
      Case tg_MoveNetWorkNode
         FrmMain.tbToolBar.Buttons("kmovenetworknode").value = tbrPressed
   End Select
End Sub



' Entra quando uma tecla é pressionada
'
'
'
Private Sub Form_KeyPress(KeyAscii As Integer)
    With FrmMain
        Select Case KeyAscii
            Case vbKeyDelete
                .tbToolBar_ButtonClick .tbToolBar.Buttons("kdelete")
            Case 19 'vbKeyControl + vbKeyS
                .tbToolBar_ButtonClick .tbToolBar.Buttons("ksave")
            Case 27 'ESC
                TCanvas.Cancel                              'cancela a operação que está sendo realizada e ainda não foi salva
                frmNetWorkLegth.txtLength.Text = 0
                Dim layerCorrente As String
                TCanvas.ToolTipText = ""
                TCanvas.Normal
'                layerCorrente = TCanvas.setCurrentLayer("")
'                layerCorrente = TCanvas.getCurrentLayer()
'                If layerCorrente <> "" Then
'                    TCanvas.Select
'                End If
                Tr.TerraEvent = tg_NoEvent
                TCanvas.clearSelectItens 0
                TCanvas.clearEditItens 0 '.clearEditItens 1: .clearEditItens 2: .clearEditItens 4: .clearEditItens 128
                FrmMain.sbStatusBar.Panels(1).Text = " "
                FrmMain.sbStatusBar.Panels(2).Text = " "
                FrmMain.sbStatusBar.Panels(3).Text = " "
                LoadToolsBar
'            Case 49 'número 1 seta para esquerda (não captura seta)
'                MsgBox ("pressionada seta para esquerda.")
'            Case Else
'                MsgBox ("Pressionado " & KeyAscii)
        End Select
    End With
End Sub


Private Sub Form_Resize()
On Error GoTo Trata_Erro

   If Me.Width > 200 And Me.Height > 200 Then
      TCanvas.Move 100, 100, Me.Width - 200, Me.Height - 200
      TCanvas.plotView
   End If

Trata_Erro:
If Err.Number = 0 Or Err.Number = 20 Then
   Resume Next
Else

   PrintErro CStr(Me.Name), "Private Sub Form_Resize", CStr(Err.Number), CStr(Err.Description), True

End If

End Sub


Private Sub Form_Unload(Cancel As Integer)

On Error GoTo Trata_Erro
   
  FrmMain.Manager1.GridVisibled False
   Tc.SetWorldByUser strUser, CDbl(xmin), CDbl(ymin), CDbl(xmax), CDbl(ymax)
   
  Set Tc = Nothing
   
  On Error Resume Next
   
  ' Set FrmMain.ViewManager1.tcs = Null
 '  Set FrmMain.ViewManager1.tvm = Null
 ' Set FrmMain.ViewManager1.tvw = Null
   
   
   FrmMain.ViewManager1.resetView
'FrmMain.ViewManager1.start

   Unload frmNetWorkLegth
   
   'para saber quantos canvas estão abertos...
   FrmMain.Tag = Int(FrmMain.Tag) - 1


Trata_Erro:
    
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
       
       PrintErro CStr(Me.Name), "Private Sub Form_Unload", CStr(Err.Number), CStr(Err.Description), True
       
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Trata_Erro
   
   TCanvas.zoomArea
   TCanvas.drawpo

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
   Else
    
      PrintErro CStr(Me.Name), "Private Sub Toolbar1_ButtonClick", CStr(Err.Number), CStr(Err.Description), True
    
   End If
End Sub
' Rotina para determinar qual o botão foi selecionado pelo usuário
' Referente as ícones no menu principal do GeoSan
' Sempre que uma ícone na janela principal do GeoSan é selecionada este é o primeiro lugar onde a aplicação entra
'
'
'
Public Sub Tb_SELECT(ByVal Button As String)
    On Error GoTo Trata_Erro 'trata erros
    Dim a As Integer, object_ids As String ' declaração das variáveis a  do tipo integer e object_ids do tipo string
    Dim retval As String ' declaração da variável retval do tipo string
    
    LastEvent = Tr.TerraEvent 'LastEvent recebe o conteúdo de Tr.TerraEvent
    TCanvas.ToolTipText = "" 'em branco
    With TCanvas ' Com o TCanvas
        Select Case Button    'selecione um case
            Case "kselection"
                TCanvas.Normal ' TCanvas da area normal desmarca item 1, item2, item3, item4 e 128
                TCanvas.Select
                Tr.TerraEvent = tg_SelectObject
                .clearEditItens 1: .clearEditItens 2: .clearEditItens 4: .clearEditItens 128
            Case "kplotview" ' plota a vista desmarca item 1, item2, item3, item4 e 128
                TCanvas.plotView
                .clearEditItens 1: .clearEditItens 2: .clearEditItens 4: .clearEditItens 128
            Case "krecompose" 'recompõe a vista desmarca item 1, item2, item3, item4 e 128
                TCanvas.recompose
                .clearEditItens 1: .clearEditItens 2: .clearEditItens 4: .clearEditItens 128
            Case "kzoomarea" ' zoom da area desmarca item 1, item2, item3, item4 e 128
                TCanvas.zoomArea
                Tr.TerraEvent = tg_ZoomArea
                .clearEditItens 1: .clearEditItens 2: .clearEditItens 4: .clearEditItens 128
            Case "kpan" ' recorta plotview desmarcaitem 1, item2, item3, item4 e 128
                TCanvas.pan
                Tr.TerraEvent = tg_Pan
            Case "kundoview" 'retorna a visualização anterior desmarca item 1, item2, item3, item4 e 128
                TCanvas.undoView
                .clearEditItens 1: .clearEditItens 2: .clearEditItens 4: .clearEditItens 128
            Case "kredoview"  'desfaz a última visualização desmarca item 1, item2, item3, item4 e 128
                TCanvas.redoView
                .clearEditItens 1: .clearEditItens 2: .clearEditItens 4: .clearEditItens 128
            Case "KFindCoordenadas" 'final das coordenadas desmarca item 1, item2, item3, item4 e 128
                .clearEditItens 1: .clearEditItens 2: .clearEditItens 4: .clearEditItens 128
                'Declaração das variáveis x,y (verificar se este x e y não estão sendo utilizados em outro lugar, pois mudou para maiúsculas na revisão 75)
                Dim X As Double, Y As Double
                X = InputBox("Informe a Coordena X ") ' entrada da coordenada x
                Y = InputBox("Informe a Coordena Y ") ' entrada da coordenada y
                If X <> 0 And Y <> 0 Then ' se x e y for diferente de zero
                    TCanvas.setWorld X - 50, Y - 50, X + 50, Y + 50 '  'configura as coordenadas mundo a serem utilizadas para desenho
                    TCanvas.plotView ' plota o layer
                End If ' final do if
            Case "KEncontraConsumidor" ' localizar consumidores
                TCanvas.setCurrentLayer "RAMAIS_AGUA" ''configura o plano "RAMAIS_AGUA" como corrente
                frmEncontraConsumidor.Show 1 ' encontra consumidor e adiciona
            Case "KEncontraTexto" ' case encontra texto e adiciona
                frmEncontraTexto.Show 1
            Case "kzoomin" ' zoom menos -
                TCanvas.zoomIn dblFatorZoomMenos
            Case "kzoomout" ' zoom mais +
                TCanvas.zoomOut dblFatorZoomMais
            Case Else
                If TCanvas.getCurrentLayer <> "" Then                    'configura o plano corrente e se for diferente da falta de seleção
                    TeDatabase1.setCurrentLayer TCanvas.getCurrentLayer  'aciona atabela para modificar o plano e configura um plano corrente
                    Set Tr.tcs = TCanvas                                 'seta e TCanvas passa ser valor para a variável Tr.tcs
                    Set Tr.tdb = TeDatabase1                             'seta e TeDatabase1 passa ser valor para a variável Tr.tdb
                    Set Tr.tdbcon = TeDatabase2                          'seta e TeDatabase2 passa ser valor para a variável Tr.tdbcon
                    Set Tr.tdbconref = TeDatabase3                       'seta e TeDatabase3 passa ser valor para a variável Tr.tdbconref
                    Set Tr.CtrlMgr = FrmMain.Manager1                    'CtrlMgr recebe o form.Manager1
                    Select Case Button ' selecione uma das opções
                        Case "kCalcularArea"
                            TCanvas.calculateArea
                            TCanvas.ToolTipText = "" ' se for igual em branco
                        Case "kdrawnetworkline" ' foi selecionada a ícone de desenhar rede de agua (esgoto ou drenagem)
                            TCanvas.clearSelectItens 0                     'desmarca se há item selecionado
                            'é aqui com o comando Tr.DrawNetWorkLine onde é ativado o início do desenho da rede (veja esta rotina na classe clsTerralib em Public Function DrawNetWorkLine)
                            If Tr.DrawNetWorkLine = True Then              'chama a classe drawnetworkline para iniciar o desenho da linha. Public Function DrawNetWorkLine(Optional mback As Boolean) As Boolean
                                frmNetWorkLegth.Init TCanvas, FrmMain
                                FrmMain.ViewManager1.LoadImageSnap Tr.cgeo.GetReferenceLayer(.getCurrentLayer), mOnSnapLock
                                FrmMain.TabStrip1.Tabs(2).Selected = True
                            Else
                                FrmMain.tbToolBar.Buttons("kdrawnetworkline").value = tbrUnpressed
                                .clearEditItens 1: .clearEditItens 2: .clearEditItens 4: .clearEditItens 128
                                Exit Sub
                            End If
                        Case "kmovenetworknode" 'mover nó da rede
                            Tr.MoveNetWorkNode
                        Case "kinsertnetworknode"
                            'fraRedes.Visible = T rue
                            Tr.DrawNetWorkNode
                        Case "kdrawtext"
                            'A implantar
                        Case "kinsertdoc" ' este
                            Tr.DrawPoint: Tr.TerraEvent = tg_DrawGeometrys
                        Case "kdrawramalAddConsumer"
                            TCanvas.Normal ' TCanvas da area normal desmarca item 1, item2, item3, item4 e 128
                            TCanvas.Select
                            Tr.TerraEvent = tg_DrawRamalAddConsumer
                            TCanvas.clearEditItens TypeGeometry.Polyguns: TCanvas.clearEditItens TypeGeometry.lines: TCanvas.clearEditItens TypeGeometry.points: TCanvas.clearEditItens TypeGeometry.texts
                            TCanvas.setCurrentLayer ("RAMAIS_AGUA")
                            FrmMain.sbStatusBar.Panels(1).Text = "Selecione o ramal para adicionar o consumidor"
                            FrmMain.sbStatusBar.Panels(2).Text = " "
                        Case "kdrawramalAuto"
                            TCanvas.Normal ' TCanvas da area normal desmarca item 1, item2, item3, item4 e 128
                            TCanvas.Select
                            Tr.TerraEvent = tg_DrawRamalAuto
                            TCanvas.clearEditItens TypeGeometry.Polyguns: TCanvas.clearEditItens TypeGeometry.lines: TCanvas.clearEditItens TypeGeometry.points: TCanvas.clearEditItens TypeGeometry.texts
                            TCanvas.setCurrentLayer ("WATERLINES")
                            FrmMain.sbStatusBar.Panels(1).Text = "Selecione o trecho de rede em que irá ligar os ramais"
                            FrmMain.sbStatusBar.Panels(2).Text = " "
                        Case "kdrawramal"
                            If ConnSec.State = 1 Then
                                TCanvas.clearSelectItens 0 'desmarca se há item selecionado
                                Tr.DrawRamal: Tr.TerraEvent = tg_DrawRamal
                            Else
                                MsgBox "A conexão com o banco de dados comercial não foi configurada para realizar esta operação.", vbInformation, "Conexão Comercial"
                            End If
                        Case "kdelete"
                            Tr.TerraEvent = tg_SelectObject             'insere um evento de seleção para que ao apagar ele saber que foi selecionado um ponto com link para um documento o qual será apagado
                            Tr.Delete
                        Case "ksearchinnetwork" ''obtem a quantidade de poligonos selecionados em memória
                            If .getSelectCount(lines) = 1 Then
                                Dim Trecho As String
                                Trecho = TCanvas.getSelectObjectId(0, lines) 'CAPTURA O TRECHO SELECIONADO
                                TCanvas.Normal                               'LIMPA A SELEÇÃO DE QUALQUER OBJETO NO MAPA
                                TCanvas.Select
                                object_ids = FrmProcess.FindValvulas(Trecho, TCanvas)   'Tr.CGeo.SELECTRede TCanvas.getSELECTObjectId(0, lines)
                                If object_ids <> "" Then
                                    frmConsumidoresDesabastecidos.Init object_ids
                                End If
                            Else
                                MsgBox "Selecione 1 trecho de rede de agua para esta função.", vbInformation, ""
                            End If
                        Case "kdeclivity"
                            If .getSelectCount(lines) = 1 Then
                                Set Tr.cgeo.tcs = TCanvas
                                Tr.cgeo.GetDeclivity .getCurrentLayer, Tr.cgeo.GetReferenceLayer(.getCurrentLayer), .getSelectObjectId(0, lines)
                            End If
                        Case "ksearchattribute"
                            Tr.SearchGeomtryForAttribute
                        Case "ksave"
                            If cConsumidor.TerraEvent = tg_MoveGpsPoint Then        'este if foi colocado pois é uma melhora no código para separar as ações por classes distintas
                                cConsumidor.SaveInDatabase
                            Else
                                Tr.SaveInDatabase
                                If FrmMain.tbToolBar.Buttons("kdrawnetworkline").value = tbrUnpressed Then
                                    With TCanvas
                                        .Normal
                                        .Select: Tr.TerraEvent = tg_SelectObject
                                        .clearEditItens 1: .clearEditItens 2: .clearEditItens 4: .clearEditItens 128
                                    End With
                                End If
                                'TCanvas.plotView  2013-05-01 - retirado pois após desenhar uma rede ele plotava a vista 3 vezes
                                LoadToolsBar
                            End If
                        Case "kdrawintersection"
                            Tr.DrawInterSection
                        Case "kdrawline"
                        Case "kdrawpoint"
                        Case "kdrawtext"
                        Case "mnuPoligono"
                            'TCanvas.Select True
                            Tr.TerraEvent = 0
                            TCanvas.Normal
                            TCanvas.drawPolygon
                        Case "kMoveVertice"
                            Tr.moveVertice: Tr.TerraEvent = tg_MoveNetWorkVertice       'chama clsTerralib.MoveVertice e informa o evento que está realizando, para iniciar o método de movimentação do vértice da rede e salvar na memória quem são os ramais conectados a mesma
                        Case "kMoveConsumidorGPS"
                            cConsumidor.TerraEvent = tg_MoveGpsPoint                    'informa a classe cConsumidor que agora é um evento de mover um consumidor para outra posição
                            cConsumidor.Move
                    End Select
                Else
                    MsgBox "Nenhum plano está ativo. Selecione antes o plano de informação que deseja realizar esta operação.", vbExclamation
                End If
        End Select
        'comprimento da linha
        If Tr.TerraEvent = tg_DrawNetWorkline Then
            frmNetWorkLegth.Init TCanvas, FrmMain
            Dim Lh As Double
            TCanvas.getLengthOfLastSegmentOfLine Lh
            frmNetWorkLegth.txtLength.Text = Lh
        Else
            Unload frmNetWorkLegth
        End If
        TCanvas_onEndPlotView                   'chama Tcanvas_onEndPlotView para acertar x,y,min e max e a tolerância de localização para desenho de redes
        LoadToolsBar                            'ativa o comando selecionado, caso seja desenho de rede, zoom área, etc. Para o programa ficar sabendo o que ele está fazendo
    End With
    Exit Sub
   
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    ElseIf Err.Number = 13 Then
        Exit Sub
    Else
       ErroUsuario.Registra "frnCanvas", "Tb_SELECT", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub




Private Sub TCanvas_GotFocus()
    TCanvas.setSelPoint 1, 4, vbMagenta         'configura o ponto de seleção para tamanho 1, quadrado (4), e cor magenta
End Sub

Private Sub TCanvas_onArea(ByVal value As Double)

   FrmMain.sbStatusBar.Panels(1).Text = "Área do polígono: " & Format(value, "0.00") & " m²"
   TCanvas.ToolTipText = "Área: " & Format(value, "0.00") & " m²"

End Sub

'Rotina que ao selecionar duplo clique do mouse, vai identificar todas as redes que estão dentro do polígono finalizado.
'
'
Private Sub TCanvas_onDblClick(ByVal Button As Long, ByVal X As Double, ByVal Y As Double)
On Error GoTo Trata_Erro
'XXX - para lembrar que é aqui que ele fecha o poligono de seleção de redes com duplo clique
'A FUNÇÃO DUPLO CLIQUE É UTILIZADA PARA FECHAR UM POLÍGONO QUE ESTÁ SENDO DESENHADO E
'APOS ISSO, INSERIR OS OBJECT_ID_ DAS LINHAS QUE ESTÃO DENTRO OU NA BORDA DO POLÍGONO E O NOME DO
'USUÁRIO QUE FEZ A SELEÇÃO EM UMA TABELA CHAMADA POLIGONO_SELEAO

   If Tr.TerraEvent = tg_DrawRamal Then 'SE ESTA DESENHANDO RAMAL e selecionou duplo click, sai pois não tem que entrar aqui nunca. Este if foi colocado pois foi verificado um bug quando o mouse está quebrado e entrando muito clicks quando o usuário pressiona apenas uma vez ele entra nesta rotina e trava o GeoSan
        Exit Sub
   End If


   Me.MousePointer = vbHourglass

   Dim i As Long
   
   geo = TCanvas.Geometry
   blnPoligonoVirtual = True
   
   TeDatabase1.setCurrentLayer "WATERLINES"
   'CARREGA NA VARIAVEL TOTAL A QUANTIDADE DE LINHAS QUE ESTÃO CONTIDADAS NO POLÍGONO
   lngTotalRedesDentro = TeDatabase1.Within(geo, tpPOLYGONS, tpLINES)
   If lngTotalRedesDentro > 0 Then
      ReDim ArrRedesDentro(lngTotalRedesDentro - 1) 'REDIMENSIONA O ARRAY
      FrmMain.ProgressBar1.Visible = True: FrmMain.ProgressBar1.value = 1: FrmMain.ProgressBar1.Max = lngTotalRedesDentro
      For i = 0 To lngTotalRedesDentro - 1
         DoEvents
         ArrRedesDentro(i) = TeDatabase1.objectIds(i)
         FrmMain.ProgressBar1.value = i + 1
      Next
   Else
      lngTotalRedesDentro = 0
   End If
    
   lngTotalRedesDivisa = TeDatabase1.Crosses(geo, tpPOLYGONS, tpLINES)
   If lngTotalRedesDivisa > 0 Then
      ReDim ArrRedesDivisa(lngTotalRedesDivisa - 1) 'REDIMENSIONA O ARRAY
      FrmMain.ProgressBar1.Visible = True: FrmMain.ProgressBar1.value = 1: FrmMain.ProgressBar1.Max = lngTotalRedesDivisa
      For i = 0 To lngTotalRedesDivisa - 1
         DoEvents
         ArrRedesDivisa(i) = TeDatabase1.objectIds(i)
         FrmMain.ProgressBar1.value = i + 1
      Next
   Else
      lngTotalRedesDivisa = 0
   End If
       
' ###########################################################################################

   TeDatabase1.setCurrentLayer "WATERCOMPONENTS"
   'CARREGA NA VARIAVEL TOTAL A QUANTIDADE DE LINHAS QUE ESTÃO CONTIDADAS NO POLÍGONO
   lngTotalPontosDentro = TeDatabase1.Within(geo, tpPOLYGONS, tpPOINTS)
   If lngTotalPontosDentro > 0 Then
      ReDim ArrPontosDentro(lngTotalPontosDentro - 1) 'REDIMENSIONA O ARRAY
      FrmMain.ProgressBar1.Visible = True: FrmMain.ProgressBar1.value = 1: FrmMain.ProgressBar1.Max = lngTotalPontosDentro
      For i = 0 To lngTotalPontosDentro - 1
         DoEvents
         ArrPontosDentro(i) = TeDatabase1.objectIds(i)
         FrmMain.ProgressBar1.value = i + 1
      Next
   Else
      lngTotalPontosDentro = 0
   End If
    
   lngTotalPontosDivisa = TeDatabase1.Crosses(geo, tpPOLYGONS, tpPOINTS)
   If lngTotalPontosDivisa > 0 Then
      ReDim ArrPontosDivisa(lngTotalPontosDivisa - 1) 'REDIMENSIONA O ARRAY
      FrmMain.ProgressBar1.Visible = True: FrmMain.ProgressBar1.value = 1: FrmMain.ProgressBar1.Max = lngTotalPontosDivisa
      For i = 0 To lngTotalPontosDivisa - 1
         DoEvents
         ArrPontosDivisa(i) = TeDatabase1.objectIds(i)
         FrmMain.ProgressBar1.value = i + 1
      Next
   Else
      lngTotalPontosDivisa = 0
   End If
       
       
' ###########################################################################################
   
   TeDatabase1.setCurrentLayer "RAMAIS_AGUA"
   lngTotalRamaisDentro = TeDatabase1.Within(geo, tpPOLYGONS, tpLINES)
   If lngTotalRamaisDentro > 0 Then
      ReDim ArrRamaisDentro(lngTotalRamaisDentro - 1) 'REDIMENSIONA O ARRAY
      FrmMain.ProgressBar1.Visible = True: FrmMain.ProgressBar1.value = 1: FrmMain.ProgressBar1.Max = lngTotalRamaisDentro
      For i = 0 To lngTotalRamaisDentro - 1
         DoEvents
         ArrRamaisDentro(i) = TeDatabase1.objectIds(i)
         FrmMain.ProgressBar1.value = i + 1
      Next
   Else
      lngTotalRamaisDentro = 0
   End If

    
   lngTotalRamaisDivisa = TeDatabase1.Crosses(geo, tpPOLYGONS, tpLINES)
   If lngTotalRamaisDivisa > 0 Then
      ReDim ArrRamaisDivisa(lngTotalRamaisDivisa - 1) 'REDIMENSIONA O ARRAY
      FrmMain.ProgressBar1.Visible = True: FrmMain.ProgressBar1.value = 1: FrmMain.ProgressBar1.Max = lngTotalRamaisDivisa
      For i = 0 To lngTotalRamaisDivisa - 1
         DoEvents
         ArrRamaisDivisa(i) = TeDatabase1.objectIds(i)
         FrmMain.ProgressBar1.value = i + 1
      Next
   Else
      lngTotalRamaisDivisa = 0
   End If
      
   FrmMain.ProgressBar1.Visible = False
   Me.MousePointer = vbDefault
   
   If lngTotalRedesDentro > 0 Or lngTotalRedesDivisa > 0 Or lngTotalRamaisDentro > 0 Or lngTotalRamaisDivisa > 0 Or lngTotalPontosDentro > 0 Or lngTotalPontosDivisa > 0 Then
       
       frmAtualizarSetores.Show 1
       
   End If
   
   TCanvas.Normal
    
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Or Err.Number = 13 Then
      Resume Next
   Else
      PrintErro CStr(Me.Name), "Private Sub TCanvas_onEndPlotView", CStr(Err.Number), CStr(Err.Description), True
   End If
End Sub
Private Sub TCanvas_onBeginPlotView()
    'MsgBox "Inicio: " & tempo & "Fim: " & Time
End Sub



Private Sub TCanvas_onEndMoveGeometries(ByVal distance As Double, ByVal deltaX As Double, ByVal deltaY As Double)
    cConsumidor.InsereTexto lastGpsObjIdPointSelected
End Sub

' Evento que é disparado quando é terminado de mover o vértice de uma linha, no caso um trecho de rede de água
' Como terminou de mover o vértice da rede, tem agora que salvar a rede na nova posição e recalcular o
' novo posicionamento dos ramais
'
'
'
Private Sub TCanvas_onEndMoveGeometryPoint()
    Dim contTrechos As Integer
    Dim contRamais As Integer
    Dim totalRamais As Integer
    Dim xIniRa As Double
    Dim yIniRa As Double
    
    TCanvas.saveOnMemory                                    'salva na memória a nova posição da rede
    TCanvas.SaveInDatabase                                  'salva no banco de dados a nova posição da rede
    TCanvas.redraw
    'inicia agora a movimentação de todos os ramais associados a esta rede que foi recém salva no banco de dados.
    totalRamais = UBound(ramalMovendo)                      'início para mover os ramais, começa
    For contRamais = 0 To totalRamais
        If ramalMovendo(contRamais).objIdTrecho = varGlobais.objIdTreSelecionado And ramalMovendo(contRamais).objIdRamal <> -1 And ramalMovendo(contRamais).objIdTrecho = varGlobais.objIdTreSelecionado Then
            Dim distIniRamalDepois As Double                'distância do início do ramal depois de tanto o trecho quanto o ramal serem movidos
            Dim moveRamal As New CCoordIniRamalDistTrecho   'classe para obter a coordenada inicial do ramal a uma determinada distância do início do trecho de rede
            Dim distEquiv As New CDistanciaEquivalente      'classe para obter a distância do início do ramal ao início do trecho após movido os mesmos
            Dim retorno As Boolean
            Dim novoComprTrecho As Double
            Dim xRamal(1) As Double, yRamal(1) As Double
            Dim comprimentoRamal As Double                  'comprimento calculado da extensão do ramal
            Dim pontoSobreLinha As Long                     'indica se o ponto de início do ramal ficou ou não sobre a linha

            pontoSobreLinha = True
            cGeoDatabase.geoDatabase.setCurrentLayer ("Waterlines")
            retorno = cGeoDatabase.geoDatabase.getLengthOfLine(varGlobais.objIdTreSelecionado, "", novoComprTrecho)
            distIniRamalDepois = distEquiv.distanciaRamalDepoisMovido(ramalMovendo(contRamais).comprTrecho, novoComprTrecho, ramalMovendo(contRamais).Distancia)
            'moveRamal.coordsRamal distIniRamalDepois, CStr(LINE_ID), cGeoDatabase.geoDatabase       'obtem as novas coordenadas inicial e final do ramal movido após mover o trecho de rede. Desativada, pois foi substituído pelo ponto perpendicular
            retorno = cGeoDatabase.geoDatabase.getMinimumDistance(0, ramalMovendo(contRamais).objIdTrecho, 2, ramalMovendo(contRamais).xHidrom, ramalMovendo(contRamais).yHidrom, comprimentoRamal, pontoSobreLinha, xIniRa, yIniRa)    'obtem a nova coordenada inicial do ramal, perpendicular ao segmento de linha mais próximo
            xRamal(0) = xIniRa
            yRamal(0) = yIniRa
            xRamal(1) = ramalMovendo(contRamais).xHidrom                                            'estas coordenadas foram testadas e estão corretas, bate com a coordenada onde está o ponto (nó) do hidrômetro
            yRamal(1) = ramalMovendo(contRamais).yHidrom
            cGeoDatabase.geoDatabase.setCurrentLayer ("RAMAIS_AGUA")                                'seta o layer em que serão apagadas e adicionadas as geometrias
            cGeoDatabase.geoDatabase.deleteGeometry ramalMovendo(contRamais).geomIdRamal, ramalMovendo(contRamais).objIdRamal, 2
            cGeoDatabase.geoDatabase.addLine ramalMovendo(contRamais).objIdRamal, xRamal(0), yRamal(0), 2
        End If
    Next                                                    'final da movimentação de ramais
    TCanvas.plotView
End Sub

' Evento que ocorre quando é selecionado um ou mais objetos no canvas
' A rotina dentro do evento  carrega as propridades
' na componente manage1(Gerenciador de Propridades) do Form Principal
' Obs: Havendo apenas um objeto selecionado é disparado .LoadDefaultProperties
'      Havendo mais de um objeto selecionado é disparado .LoadComunsObjects
' Autor: Luis CLaudio
' Data: 31/08/06
' Nesta rotina é configurada a escala da tolerância de localização
'
'
'
Private Sub TCanvas_onEndPlotView()
    On Error GoTo Trata_Erro
    Dim MyScale As Double
    Dim pixelsTela As Long                  'número de pixels totais na largura do canvas
    Dim distHorizontal As Double            'distância horizontal em metros do canvas
    Dim tamanhoPixel As Double              'tamanho em metros de um pixel
    Dim tolerancia As Double                'tolerância de localização de extremidadde do drawnetworkline
    Dim toleranciaSnap As Double            'tolerância do snap no canvas
    
    tolerancia = 0.5                        'define a tolerância de localização de uma extremidade de uma rede, mais do que isso ele cria um novo nó
    MyScale = TCanvas.getScale
    TCanvas.getWorld xmin, ymin, xmax, ymax 'obtem as coordenadas do box do canvas no formato mundo
    ViewName = TeViewDatabase1.getActiveView
    'carrega as variáveis globais para o módulo de impressão
    CanvasXmin_ = xmin
    CanvasYmin_ = ymin
    CanvasXmax_ = xmax
    CanvasYmax_ = ymax
    strViewAtiva_ = ViewName
    FrmMain.txtEscala.Text = "1 / " & Round(MyScale, 0)
    If TCanvas.getCurrentLayer <> "" Then
        strLayerAtivo = TCanvas.getCurrentLayer
    Else
        strLayerAtivo = ""
    End If
    TCanvas.ToolTipText = ""
    
    'aqui nas próximas 4 linhas ele irá converter as unidades de medida da janela do canvas (Twips) para pixels
    'e depois irá determinar um valor de tolerância em pixels para o snap, que aceita somente pixels como unidade de medida
    pixelsTela = ConvertTwipsToPixels(TCanvas.Width, 0)                 'obtem o número total de pixels do canvas na horizontal
    distHorizontal = xmax - xmin                                        'obtem a distância em metros na horizontal do canvas
    toleranciaSnap = 1.5 * tolerancia * pixelsTela / distHorizontal           'calcula o número de pixels para a tolerância em metros especificada
    TCanvas.toleranceToSnap(0) = toleranciaSnap                         'seta no canvas a tolerância de snap - 0 = estremidades
    FrmMain.sbStatusBar.Panels(2).Text = "Snap: " & Round(toleranciaSnap, 2)  'mostra a tolerância de snap na barra de status
    'para corrigir o DrawNetWorkLine - Luis
    'aqui é definida a tolerância de localização quando estiver desenhando uma rede (snap)
    'foram inseridas algumas tolerâncias a mais para ver se resolve quando não localiza o nó ou pega o do lado por engano
    'não resolveu e ai colocamos o snap igual a .tolerance do canvas
    If Tr.TerraEvent = 1 Then 'tg_DrawNetWorkline - caso esteja desenhando uma rede muda a tolerância conforme a escala em que o usuário estiver
        With TCanvas
        FrmMain.sbStatusBar.Panels(3).Text = "Tolerância localização Rede: " & Round(tolerancia, 2) 'mostra a tolerância de drawNetworkLine na barra de status
        .tolerance = tolerancia
'            MyScale = .getScale
'            Select Case MyScale
'            Case Is < 10
'                .tolerance = 0.001
'            Case Is < 50
'                .tolerance = 0.005
'            Case Is < 100
'                .tolerance = 0.01
'            Case Is < 200
'                .tolerance = 0.05
'            Case Is < 300
'                .tolerance = 0.075
'            Case Is < 500
'                .tolerance = 0.1
'            Case Is < 1000
'                .tolerance = 0.5
'            Case Is >= 1000
'                .tolerance = 1
'            End Select
        End With
    Else
        TCanvas.tolerance = 1
    End If
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        PrintErro CStr(Me.Name), "Erro na tolerância de localização em Private Sub TCanvas_onEndPlotView", CStr(Err.Number), CStr(Err.Description), True
    End If
End Sub
' Entra nesta rotina quando o usuário termina de selecionar uma geometria.
'
'
'
Private Sub TCanvas_onEndSELECT()
On Error GoTo Trata_Erro

   Dim strDistrito As String
   Dim IdDistrito As Integer

   Dim i As Integer, j As Integer, VarObj As String
   Dim frm As New FrmAssociation                                                            'formulário para a associação de documentos a pontos no mapa
   'este select é para o cadastro de ramais que vem da leitura em campo
   Select Case Tr.TerraEvent       'verifica o comando que está sendo executado
        Case tg_DrawRamalAutoSelecionaConsumidor                                            'aqui o usuário vai selecionando um consumidor após o outro para ligá-los no mesmo trecho de rede selecionado em tg_DrawRamalAuto
            object_id_consumidorSelecionado = TCanvas.getSelectObjectId(0, TypeGeometry.points)
            FrmMain.sbStatusBar.Panels(2).Text = "Ligação: " & str(object_id_consumidorSelecionado)
            InsereRamalLigacaoGPS object_id_consumidorSelecionado, object_id_redeSelecionada
            
        Case tg_DrawRamalAuto                                                               'aqui o usuário seleciona a rede de água em que serão ligados os ramais
            object_id_redeSelecionada = TCanvas.getSelectObjectId(0, TypeGeometry.lines)
            TCanvas.Normal ' TCanvas da area normal desmarca item 1, item2, item3, item4 e 128
            TCanvas.Select
            Tr.TerraEvent = tg_DrawRamalAutoSelecionaConsumidor
            TCanvas.clearEditItens TypeGeometry.Polyguns: TCanvas.clearEditItens TypeGeometry.lines: TCanvas.clearEditItens TypeGeometry.points: TCanvas.clearEditItens TypeGeometry.texts
            TCanvas.setCurrentLayer ("NXGS_V_LIG_COMERCIAL_GPS")
            FrmMain.sbStatusBar.Panels(2).Text = "Rede selecionada: " & str(object_id_redeSelecionada)
            FrmMain.sbStatusBar.Panels(1).Text = "Selecione a ligação de água"
        
        Case tg_DrawRamalAddConsumer                                                        'aqui o usuário terminou de selecionar o ramal para poder em seguida adicionar mais consumidores a este ramal
            object_id_ramalAddConsumerSelecionado = TCanvas.getSelectObjectId(0, TypeGeometry.lines)
            TCanvas.Normal ' TCanvas da area normal desmarca item 1, item2, item3, item4 e 128
            TCanvas.Select
            Tr.TerraEvent = tg_DrawRamalAddConsumerSelecionaConsumidor
            TCanvas.clearEditItens TypeGeometry.Polyguns: TCanvas.clearEditItens TypeGeometry.lines: TCanvas.clearEditItens TypeGeometry.points: TCanvas.clearEditItens TypeGeometry.texts
            TCanvas.setCurrentLayer ("NXGS_V_LIG_COMERCIAL_GPS")
            FrmMain.sbStatusBar.Panels(2).Text = "Ramal selecionado: " & str(object_id_ramalAddConsumerSelecionado)
            FrmMain.sbStatusBar.Panels(1).Text = "Selecione a ligação de água para ligar no ramal"
            
        Case tg_DrawRamalAddConsumerSelecionaConsumidor
            object_id_ramalAddConsumerConsumidorSelecionado = TCanvas.getSelectObjectId(0, TypeGeometry.points)
            FrmMain.sbStatusBar.Panels(2).Text = "Ligação: " & str(object_id_ramalAddConsumerConsumidorSelecionado)

            InsereLigacaoNoRamalSelecionado object_id_ramalAddConsumerSelecionado, object_id_ramalAddConsumerConsumidorSelecionado
        
        Case Else
        
   End Select
   
   With FrmMain.Manager1
      If TCanvas.getSelectCount(lines) Or TCanvas.getSelectCount(points) Or TCanvas.getSelectCount(Polyguns) Then            'retorna quantas geometrias foram selecionadas do tipo linha, ponto ou polígono e com isso verrifica se foram selecionadas uma destas geometrias
         .GridEnabled True: .GridVisibled True
         Select Case Tr.cgeo.GetLayerTypeReference(TCanvas.getCurrentLayer)
            
            Case LayerTypeRefence.Trecho_Rede_Agua, LayerTypeRefence.Trecho_Rede_Drenagem, LayerTypeRefence.Trecho_Rede_esgoto, _
               LayerTypeRefence.Componente_Rede_Agua, LayerTypeRefence.Componente_Rede_Drenagem, LayerTypeRefence.Componente_Rede_Esgoto
               'Verifica a seleção apenas das geometrias 2(linhas) e 4(Pontos)
               varGlobais.objIdTreSelecionado = TCanvas.getSelectObjectId(0, 2)             'obtem o object_id do trecho de rede selecionado. Será utilizado para movimentação do vértice
               For j = 2 To 4 Step 2
                  
                  Dim X As String
                  X = TCanvas.getSelectObjectId(0, 2) ' 2 - representa elemento tipo linha
                  
                  If TCanvas.getSelectCount(j) = 1 Then
                     .LoadDefaultProperties TCanvas.getSelectObjectId(0, j), TCanvas.getCurrentLayer, False
                  
                  ElseIf TCanvas.getSelectCount(j) > 1 Then
                     For i = 0 To TCanvas.getSelectCount(j) - 1
                        With TCanvas
                           VarObj = IIf(i, VarObj & "," & .getSelectObjectId(i, j), .getSelectObjectId(i, j))
                        End With
                     Next
                      'Carrega Prorpiedades Properties Manager
                     .LoadComunsObjects VarObj, TCanvas.getCurrentLayer, FrmMain.mnuMultProperteis.Checked
                  End If
                  
                  FrmMain.TabStrip1.Tabs(2).Selected = True
                  Tr.TerraEvent = tg_SelectObject 'Define o evento de selecao para a classe
               
               Next
            
            Case LayerTypeRefence.DOCUMENTOS
               If TCanvas.getSelectCount(points) = 1 Then
                  If LastDocument <> TCanvas.getSelectObjectId(0, points) Then
                     LastDocument = TCanvas.getSelectObjectId(0, points)
                     frm.Init TCanvas.getSelectObjectId(0, points), TCanvas, TeDatabase1
                     LastDocument = ""
                  Else
                    LastDocument = ""
                  End If
               End If
            
            Case LayerTypeRefence.OUTROS
               If TCanvas.getSelectCount(1) = 1 Then
                  .LoadDefaultProperties TCanvas.getSelectObjectId(0, 1), TCanvas.getCurrentLayer, False
               ElseIf TCanvas.getSelectCount(2) = 1 Then
                  .LoadDefaultProperties TCanvas.getSelectObjectId(0, 2), TCanvas.getCurrentLayer, False
               ElseIf TCanvas.getSelectCount(4) = 1 Then
                  .LoadDefaultProperties TCanvas.getSelectObjectId(0, 4), TCanvas.getCurrentLayer, False
               ElseIf TCanvas.getSelectCount(128) = 1 Then
                  .LoadDefaultProperties TCanvas.getSelectObjectId(0, 128), TCanvas.getCurrentLayer, False
               End If
               FrmMain.TabStrip1.Tabs(2).Selected = True
                       
            Case LayerTypeRefence.Poligonos
                idPoligonSel = TCanvas.getSelectGeoId(0, 1)
                strLayerAtivo = TCanvas.getCurrentLayer
               
            Case LayerTypeRefence.CONSUMIDOR_GPS
                Dim consumidorObject_id As String
                
                consumidorObject_id = TCanvas.getSelectObjectId(0, points)                                          'obtem o object_id do ponto selecionado
                If TCanvas.getSelectCount(points) = 1 Then                                                          'retorna o número de pontos GPS selecionados no Canvas e que estão em memória
                    lastGpsObjIdPointSelected = consumidorObject_id                                                 'é um novo ponto GPS selecionado, então salva o object_id do mesmo na memória
                    FrmMain.sbStatusBar.Panels(1).Text = cConsumidor.ObtemEnderecoCompleto(consumidorObject_id)     'Mostra o endereço na barra de status para o usuário saber o que selecionou
                Else
                    MsgBox ("Selecionou mais de um ponto GPS, selecione novamente.")
                End If
                strLayerAtivo = TCanvas.getCurrentLayer
                
            ' Caso esteja tratando um ramal
            Case LayerTypeRefence.RAMAIS_AGUA, LayerTypeRefence.RAMAIS_ESGOTO
               Set Tr.tcs = TCanvas
               Set Tr.tdbconref = TeDatabase2
               Tr.tdbconref.setCurrentLayer Tr.cgeo.GetLayerOperation(TCanvas.getCurrentLayer, 1)           ' trecho de rede de água. Não temos mais retorno de polígono de lote associado a ligação de água
               If TCanvas.getSelectCount(lines) Then
                  Tr.OnRamal Position_X, Position_Y, TCanvas.getSelectObjectId(0, lines)
               ElseIf TCanvas.getSelectCount(points) Then
                  Tr.OnRamal Position_X, Position_Y, TCanvas.getSelectObjectId(0, points)
               End If
               
               
               Tr.TerraEvent = tg_SelectObject
         End Select
      Else
         .GridEnabled False: .GridVisibled False: FrmMain.TabStrip1.Tabs(1).Selected = True
      End If
   End With
   FrmMain.SizeControls
   Exit Sub

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
   Else
      ErroUsuario.Registra "frmCanvas", "onEndSELECT", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
   End If
End Sub
' Desvia quando encontra um erro
'
' code - Código de identificação do erro
' message - mesnsagem explicativa do erro
'
Private Sub TCanvas_onError(ByVal code As String, ByVal errorMessage As String)
   Select Case code
        Case "Err032"
            'canvas não foi aberto ainda, desconsiderar
        Case "Err068"
            MsgBox "Rede muito próxima." & vbCrLf & vbCrLf & "Mensagem: " & code & " - " & errorMessage
        Case "Err030"
            MsgBox "Selecione um layer antes." & vbCrLf & vbCrLf & "Mensagem: " & code & " - " & errorMessage
        Case "Err028"
            ErroUsuario.Registra "frmCanvas", "TCanvas_onError. Usuário na tela do GeoSan selecionou um comando a apareceu este erro e continuou utilizando o software. Mensagem: " & code & " Descrição: " & errorMessage, CStr(Err.Number), CStr(Err.Description), False, glo.enviaEmails 'não mostra mensagem para o usuário
        Case Else
            ErroUsuario.Registra "frmCanvas", "TCanvas_onError. Usuário na tela do GeoSan selecionou um comando a apareceu este erro e continuou utilizando o software. Mensagem: " & code & " Descrição: " & errorMessage, CStr(Err.Number), CStr(Err.Description), False, glo.enviaEmails 'não mostra mensagem para o usuário
            MsgBox "Não é possível realizar este comando. Mensagem: " & code & " - " & errorMessage
    End Select
End Sub

Private Sub TCanvas_onIntersectionPoint(ByVal X As Double, ByVal Y As Double)
On Error GoTo Trata_Erro
   TeDatabase1.moveNetWorkNodeTo "watercomponents", "WATERLINES", "", , X, Y
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
   Else
    
      PrintErro CStr(Me.Name), "Private Sub TCanvas_onIntersectionPoint", CStr(Err.Number), CStr(Err.Description), True
      
   End If
End Sub
' Evento de captura de tecla utilizada
' A tecla PageUp faz a função de ZOOM OUT(afastamento)
' A tecla PageDown faz a função de ZOOM IN(aproximação)
' É utilizado um arquivo externo (geosan.ini) para armazenar o fator de zoom que será aplicado quando a função for chamada
' Variáveis carregadas no evento MouseMove Position_X e Position_Y possuem coordenadas do mouse
' É feita centralização do mapa no local do ponteiro do mouse antes do zoom
'
' key - código da tecla pressionada
'
Private Sub TCanvas_onKeyPress(ByVal key As Long)
    On Error GoTo Trata_Erro
    Dim retval As String

    'TCanvas.setWorld Position_X - 50, Position_Y - 50, Position_X - 50, Position_Y - 50
    'TCanvas.plotView
    'Dim Scala As Double
    'Scala = TCanvas.getScale
    Select Case key
        Case 27                     'ESC
            Dim layerCorrente As String
            TCanvas.ToolTipText = ""
            TCanvas.Cancel
            TCanvas.Normal
'            layerCorrente = TCanvas.setCurrentLayer(Null)
'            layerCorrente = TCanvas.getCurrentLayer()
'            If layerCorrente <> "" Then
'                TCanvas.Select
'            End If
            Tr.TerraEvent = tg_NoEvent
            TCanvas.clearSelectItens 0
            TCanvas.clearEditItens 0 '.clearEditItens 1: .clearEditItens 2: .clearEditItens 4: .clearEditItens 128
            FrmMain.sbStatusBar.Panels(1).Text = " "
            FrmMain.sbStatusBar.Panels(2).Text = " "
            FrmMain.sbStatusBar.Panels(3).Text = " "
            LoadToolsBar
        Case 33                     'PageUp
            TCanvas.zoomIn dblFatorZoomMenos
            TCanvas.redraw                          'para que o comando de seleção de polígono continue aparecendo
            'TCanvas.zoomIn = Replace(ReadINI("MAPA", "ZOOM_MAIS", App.path & "\CONTROLES\GEOSAN.ini"), ",", ".")
        Case 34                     'PageDown
            TCanvas.zoomOut dblFatorZoomMais
            TCanvas.redraw                          'para que o comando de seleção de polígono continue aparecendo
            'TCanvas.zoomOut = Replace(ReadINI("MAPA", "ZOOM_MENOS", App.path & "\CONTROLES\GEOSAN.ini"), ",", ".")
        Case 46 'DEL
            Tr.Delete
        Case 48                     'zero
            FrmMain.sbStatusBar.Panels(5).Text = "0.00"
            X1 = X1i
            Y1 = Y1i
        Case 87, 119           'W ou w
            TCanvas.verticalPan 50
            TCanvas.redraw                          'para que o comando de seleção de polígono continue aparecendo
        Case 90, 122            'Z ou z
            TCanvas.verticalPan -50
            TCanvas.redraw                          'para que o comando de seleção de polígono continue aparecendo
        Case 65, 97             'A ou a
            TCanvas.horizontalPan -50
            TCanvas.redraw                          'para que o comando de seleção de polígono continue aparecendo
        Case 83, 115            'S ou s
            TCanvas.horizontalPan 50
            TCanvas.redraw                          'para que o comando de seleção de polígono continue aparecendo
        Case 81, 113            'q ou Q              para mover um ponto gps
            cConsumidor.TerraEvent = tg_MoveGpsPoint                    'informa a classe cConsumidor que agora é um evento de mover um consumidor para outra posição
            cConsumidor.Move
    End Select
    Exit Sub

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        ErroUsuario.Registra "frmCanvas", "TCanvas_onKeyPress", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub
' Entra aqui para as teclas especiais como Enter, seta, etc.
'
'
'
Private Sub TCanvas_onKeyUp(ByVal key As Long, ByVal Shift As Long, ByVal ctrl As Long)
    On Error GoTo Trata_Erro
    
    Select Case key
        Case 13         'ENTER
            FrmMain.ActiveForm.Tb_SELECT "ksave"
        Case 39         'CTRL + Seta para direita
            TCanvas.horizontalPan 50
            TCanvas.redraw                          'para que o comando de seleção de polígono continue aparecendo
        Case 37         'CTRL + Seta para esquerda
            TCanvas.horizontalPan -50
            TCanvas.redraw                          'para que o comando de seleção de polígono continue aparecendo
        Case 38         'CTRL + Seta para cima
            TCanvas.verticalPan 50
            TCanvas.redraw                          'para que o comando de seleção de polígono continue aparecendo
        Case 40         'CTRL + Seta para baixo
            TCanvas.verticalPan -50
            TCanvas.redraw                          'para que o comando de seleção de polígono continue aparecendo
    End Select
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        ErroUsuario.Registra "frmCanvas", "TCanvas_onKeyUp", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Sub
' Evento quando estou desenhando uma linha e entro o segundo ponto da mesma
' Este evento está sendo utilizado apenas quando desenho ramais
'
' distance - distância da linha do primeiro ponto até o segundo ponto
'
Private Sub TCanvas_onLine(ByVal distance As Double)
    On Error GoTo Trata_Erro
    If Tr.TerraEvent = tg_DrawRamal Then 'SE ESTA DESENHANDO RAMAL
        Tr.OnRamal Position_X, Position_Y, ""
    End If
    Exit Sub
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        PrintErro CStr(Me.Name), "Private Sub TCanvas_onLine", CStr(Err.Number), CStr(Err.Description), True
    End If
End Sub

'Procedimento que carrega no trevview do main os temas do tcanvas corrente
Private Function LoadThemes() 'ViewName As String)
''On Error GoTo Trata_Erro


   If Not TCanvas Is Nothing Then
      Screen.MousePointer = vbHourglass
      DoEvents
    
      
      If ViewName <> "" Then
      
'TeViewDatabase1.connection = Conn

         TeViewDatabase1.setActiveView ViewName
      Else
         ViewName = TeViewDatabase1.getActiveView
       
      End If

      With FrmMain.ViewManager1
         Set .tcs = TCanvas
       
          
         Set .tvw = TeViewDatabase1
      
         
         Set .mConn = Conn
          

         .Provider = typeconnection
         FrmMain.txtEscala.Text = "1 / " & Round(TCanvas.getScale, 0)
         .start
         Select Case Tr.TerraEvent
            Case tg_DrawNetWorkline
               .LoadImageSnap Tr.cgeo.GetReferenceLayer(TCanvas.getCurrentLayer), mOnSnapLock
         End Select
         .LoadImageSnap TCanvas.getCurrentLayer, mOnSet
      End With
     
      Me.Caption = "Vista: " & TeViewDatabase1.getActiveView
     
         
 
      Screen.MousePointer = vbNormal
      If Tr.TerraEvent = tg_DrawNetWorkline Then
         frmNetWorkLegth.Init TCanvas, FrmMain
         Dim Lh As Double
         TCanvas.getLengthOfLastSegmentOfLine Lh
         frmNetWorkLegth.txtLength.Text = Lh
      Else
         Unload frmNetWorkLegth
      End If
   End If
Trata_Erro:

  '  If Err.Number = 0 Or Err.Number = 20 Then
    '   Resume Next
    'Else
       
    '  PrintErro CStr(Me.Name), "Private Function LoadThemes", CStr(Err.Number), CStr(Err.Description), True
       
  ' End If
End Function
' Entra nesta rotina quando o mouse é pressionado
'
' Button - botão do mouse selecionado
' x, y, z - coordenada em que o mouse foi selecionado
'
Private Sub TCanvas_onMouseDown(ByVal Button As Long, ByVal X As Double, ByVal Y As Double)
    On Error GoTo Trata_Erro
    Dim controlaErro As String              'para indicar onde ocorreu o erro, caso ocorra
    
    controlaErro = "sem erro"
    X1 = 0 'passa as coordenadas para calculo e exibição
    Y1 = 0
    
    Select Case Button  'VERIFICA O BOTÃO DO MOUSE QUE FOI SELECIONADO
        
        Case 0          'SELECIONADO O BOTÃO DA ESQUERDA
            Select Case Tr.TerraEvent       'verifica o comando que está sendo executado
            
                Case tg_DrawNetWorkline     'DESENHANDO REDE
                        controlaErro = "tg_DrawNetWorkline"
                        FrmMain.Manager1.GridEnabled True
                        X1 = X 'passa as coordenadas para calculo e exibição
                        Y1 = Y

                Case tg_MoveNetWorkNode     'MOVENDO REDE
                        controlaErro = "tg_MoveNetWorkNode"
                        FrmMain.Manager1.GridEnabled True
                        X1 = X 'passa as coordenadas para calculo e exibição
                        Y1 = Y

                Case tg_DrawNetWorkNode     'DESENHANDO UM NÓ
                    controlaErro = "tg_DrawNetWorkNode-1"
                    Tr.SaveInDatabase: FrmMain.Manager1.GridEnabled True
                    controlaErro = "tg_DrawNetWorkNode-2"
                    With TCanvas
                        .Normal
                        .Select: Tr.TerraEvent = tg_SelectObject
                        .clearEditItens 1: .clearEditItens 2: .clearEditItens 4: .clearEditItens 128
                    End With
                    controlaErro = "tg_DrawNetWorkNode-3"
                    LoadToolsBar
                
                Case tg_DrawRamal           'DESENHANDO UM RAMAL
                    If UCase(TCanvas.getCurrentLayer) = "RAMAIS_AGUA" Or UCase(TCanvas.getCurrentLayer) = "RAMAIS_ESGOTO" Then
                        'ESTA DESENHANDO RAMAL, CAPTURA O PRIMEIRO CLIQUE DO MOUSE E TESTA SE ESTE CLIQUE
                        'FOI FEITO SOBRE UMA REDE
                        If CLIQUE_RAMAL = 0 Then
                            'VERIFICA SE O LAYER CORRENTE É O DE RAMAIS DE AGUA OU ESGOTO
                            'SE FOR O DE AGUA, SETA O CURRENT LAYER DO TEDATABASE PARA RAMAIS_AGUA
                            'SE FOR O DE ESGOTO, SETA O CURRENT LAYER DO TEDATABASE PARA RAMAIS_ESGOTO
                            controlaErro = "tg_DrawRamal-RAMAIS_AGUA-1"
                            If UCase(TCanvas.getCurrentLayer) = "RAMAIS_AGUA" Then
                                controlaErro = "tg_DrawRamal-RAMAIS_AGUA-2"
                                TeDatabaseRamais.setCurrentLayer "WATERLINES"
                            Else
                                controlaErro = "tg_DrawRamal-RAMAIS_AGUA-3"
                                TeDatabaseRamais.setCurrentLayer "SEWERLINES"
                            End If
                            'VERIFICA SE O USUÁRIO CLICOU SOBRE UMA REDE DE AGUA OU ESGOTO
                            controlaErro = "tg_DrawRamal-RAMAIS_AGUA-4"
                            intQtdLinhasNaCoordenada = TeDatabaseRamais.locateGeometry(X, Y, tpLINES, 1)
                            'intQtdLinhasNaCoordenada = TeDatabaseRamais.locateGeometryXY(x, y, tpLINES)
                            'CASO NÃO, EXIBE MENSAGEM E REINICIA O PROCESSO
                            controlaErro = "tg_DrawRamal-RAMAIS_AGUA-5"
                            If intQtdLinhasNaCoordenada = 0 Then
                                controlaErro = "tg_DrawRamal-RAMAIS_AGUA-6"
                                MsgBox "Inicie o desenho do ramal partindo do trecho de rede.", vbInformation, ""
                                TCanvas.Normal
                                TCanvas.Select
                                CLIQUE_RAMAL = 0
                                TCanvas.clearSelectItens 0 'desmarca se há item selecionado
                                Tr.DrawRamal 'reinicia o processo de cadastramento de ramal
                                Tr.TerraEvent = tg_DrawRamal
                            'CASO HÁ MAIS DE UMA REDE SOB O CLIQUE, EXIBE MENSAGEM E REINICIA O PROCESSO
                            ElseIf intQtdLinhasNaCoordenada > 1 Then
                                controlaErro = "tg_DrawRamal-RAMAIS_AGUA-7"
                                MsgBox "Foi identificado mais de um trecho de rede no local selecionado." & Chr(13) & Chr(13) & "tente novamente.", vbInformation, ""
                                TCanvas.Normal
                                TCanvas.Select
                                CLIQUE_RAMAL = 0
                                TCanvas.clearSelectItens 0 'desmarca se há item selecionado
                                Tr.DrawRamal 'reinicia o processo de cadastramento de ramal
                                Tr.TerraEvent = tg_DrawRamal
                                'CASO SIM, CAPTURA O OBJECT_ID_ DA REDE QUE FOI SELECIONADA E PASSA
                                'PARA A VARIÁVEL QUE VAI SALVAR O RAMAL
                            Else
                                controlaErro = "tg_DrawRamal-RAMAIS_AGUA-8"
                                ramal_Object_id_trecho = TeDatabaseRamais.objectIds(0)
                                'TCanvas.ToolTipText = "Rede: " & ramal_Object_id_trecho
                                'GUARDA A INFORMAÇÃO DE QUE O PRIMEIRO CLIQUE JA FOI DADO PARA DESENHAR O RAMAL
                                CLIQUE_RAMAL = 1
                            End If
                        Else
                            CLIQUE_RAMAL = 0
                        End If
                    End If

                Case Else                   'nenhuma das anteriores
                        controlaErro = "tg_DrawRamal-RAMAIS_AGUA-9"
                        FrmMain.Manager1.GridEnabled False
            End Select

        Case 1          'SELECIONADO O BOTÃO DIREITO DO MOUSE
            Select Case tbrPressed
                Case FrmMain.tbToolBar.Buttons("kdrawnetworkline").value        'usuário selecionou anteriormente que estava desenhando uma rede
                    controlaErro = "tbrPressed-1"
                    'então vamos reiniciar o desenho da rede a partir do início
                    TCanvas.Normal                      'volta o canvas para o estado de visualização
                    'limpa todos os itens editados em memória, as geometrias das listas temporárias e geometrias a serem removidas do banco de dados
                    TCanvas.clearEditItens (2)          'limpa linhas
                    TCanvas.clearEditItens (4)          'limpa pontos
                    controlaErro = "tbrPressed-2"
                    Tr.DrawNetWorkLine                  'ativa novamente o desenho de uma rede a partir do início
                    controlaErro = "tbrPressed-3"
                    Select Case LastEvent               'precisa verificar o que faz e se está passando realmente por este select case
                        Case tg_DrawNetWorkline
                            controlaErro = "tbrPressed-4"
                            Tr.DrawNetWorkLine True
                        Case tg_MoveNetWorkNode
                            controlaErro = "tbrPressed-5"
                            Tr.MoveNetWorkNode True
                    End Select
            End Select
        
        Case Else       'nenhuma das anteriores

    End Select
    Exit Sub
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    ElseIf Err.Number = -2147467259 Then            ' indica que não existe conexão com o banco de dados.
        ErroUsuario.Registra "frmCanvas", "onMouseDown (-2147467259)", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails, controlaErro
    Else
        ErroUsuario.Registra "frmCanvas", "onMouseDown", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails, controlaErro
    End If
End Sub
' Entra nesta rotina quando o usuário move o mouse, independente de qualquer coisa ou botão ter sido selecionado
' Ele sempre entra nesta rotina quando o mouse move dentro da área do mapa
' Atualmente ela está apenas atualizando as coordenadas na barra de status e verificando o comprimento
'
' x, y - coordenadas UTM da posição do mouse na tela
' lat, long - latitude e longitude da posição do mouse na tela
'
Private Sub TCanvas_onMouseMove(ByVal X As Double, ByVal Y As Double, ByVal lat As String, ByVal lon As String)
    On Error GoTo Trata_Erro
    Dim TBP As String
    Dim TBA As String
    Dim pesquisar As Boolean
    Dim dist As Integer
    Dim COMP As Double
    pesquisar = False
    If (xOld - X) > 3 Or (X - xOld) > 3 Then
        xOld = X
        pesquisar = True
        'TCanvas.ToolTipText = ""
    ElseIf (yOld - Y) > 3 Or (Y - yOld) > 3 Then
        yOld = Y
        pesquisar = True
        'TCanvas.ToolTipText = ""
    End If
    If pesquisar = True Then
        'PEGAR O NOME DA TABELA NO GEOSAN.INI
        'acredito que esta rotina não é mais utilizada, pois o GeoSan não mais lê os dados do lote da prefeitura. O consumidor agora é salvo no ramal diretamente
        If UCase(TCanvas.getCurrentLayer) = UCase("RAMAIS_AGUA") Or _
            UCase(TCanvas.getCurrentLayer) = UCase("RAMAIS_ESGOTO") Then
            If ReadINI("RAMAISFILTROLOTES", "ATIVADO", App.path & "\CONTROLES\GEOSAN.INI") = "SIM" Then
                TBP = ReadINI("RAMAISFILTROLOTES", "TABELA_PLANO", App.path & "\CONTROLES\GEOSAN.INI")
                TBA = ReadINI("RAMAISFILTROLOTES", "TABELA_ATRIB", App.path & "\CONTROLES\GEOSAN.INI")
                Call Pesquisa_Dados_Lote(X, Y, lat, lon, TBA, TBP)
            End If
        End If
    End If
    FrmMain.sbStatusBar.Panels(4).Text = "x: " & Round(X, 2) & " - y:" & Round(Y, 2)
    'If X1 <> 0 Then ' SE A VARIAVEL DE PRIMEIRO CLICK ESTIVER ZERADA...
    X1i = X
    Y1i = Y
        COMP = Sqr((Abs(X - X1) ^ 2) + (Abs(Y - Y1) ^ 2))
'        FrmMain.sbStatusBar.Panels(1).Text = "Comprimento da rede: " & Format(COMP, "0.00") & " m"
        FrmMain.sbStatusBar.Panels(5).Text = Format(COMP, "0.00") & " m"
        'TCanvas.ToolTipText = Format(COMP, "0.00") & " m"
    'Else
        'FrmMain.sbStatusBar.Panels(1).Text = ""
    'End If
    Exit Sub
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    ElseIf Err.Number = 11 Then
        Exit Sub
    ElseIf Err.Number = 6 Then  'Está com zoom muito longe fora do permissível, o mapa do usuário está errado
        Exit Sub
    Else
        MsgBox Err.Number
        PrintErro CStr(Me.Name), "Private Sub TCanvas_onMouseMove", CStr(Err.Number), CStr(Err.Description), True
    End If
End Sub

Sub Pesquisa_Dados_Lote(ByVal X As Double, ByVal Y As Double, ByVal lat As String, ByVal lon As String, ByVal TBAtributo As String, ByVal TBPlano As String)

On Error GoTo Trata_Erro
      Dim rs As ADODB.Recordset
      Dim Obj As String, str As String, Mystep As String

      

      'PEGAR O NOME DA TABELA NO GEOSAN.INI
      'saber a tabela de geometrias
      
      



      
   If typeconnection <> 4 Then
      
      
      
      TeDatabase1.connection = Conn
      Else
   TeDatabase2.Provider = typeconnection


      TeDatabase1.connection = TeAcXConnection1.objectConnection_

      End If
      '
      'tabela = "LOTES_PREF"
      If TBPlano <> "" And TBAtributo <> "" Then
      
         TeDatabase1.setCurrentLayer CStr(TBPlano)
         
         If TeDatabase1.locateGeometryXY(X, Y, tpPOLYGONS) = 1 Then
            
            'LOCALIZADA 1 GEOMETRIA DE POLIGONO DE LOTE
            'LOCALIZAR NA TABELA DE ATRIBUTO QUAL IPTU DO LOTE
            
            idAutoLote = TeDatabase1.objectIds(0)
            Dim ba As String
            Dim be As String
            Dim bi As String
             Dim h As String
            ba = "CADASTRO"
            be = TBAtributo
            h = "be"
            bi = "LOTE_ID"
            
            If frmCanvas.TipoConexao <> 4 Then
            
            str = "SELECT CADASTRO AS " + """" + "IPTU" + """" + " FROM " & TBAtributo & " WHERE LOTE_ID = '" & idAutoLote & "'"
            Else
            str = "SELECT " + """" + ba + """" + " AS " + """" + "IPTU" + """" + " FROM " + """" + TBAtributo + """" + " WHERE " + """" + bi + """" + " = '" & idAutoLote & "'"
            End If
            
            Set rs = New ADODB.Recordset
           ' rs.Open str, Conn, adOpenForwardOnly, adLockReadOnly
           rs.Open str, Conn, adOpenDynamic, adLockOptimistic
            If rs.EOF = False Then

                TCanvas.ToolTipText = "IPTU: " & rs!IPTU

            End If
            rs.Close

         Else
         
            TCanvas.ToolTipText = ""
         
         End If
    
End If
      Position_X = X
      Position_Y = Y
      'FrmMain.sbStatusBar.Panels(4).Text = "x: " & Round(x, 2) & " - y:" & Round(y, 2)
      Set rs = Nothing
Exit Sub

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
    
      PrintErro CStr(Me.Name), "Sub Pesqisa_Dados_Lotes", CStr(Err.Number), CStr(Err.Description), True
      
   End If
End Sub

'Private Sub txtEdit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 2 Then 'Botão direito foi pressionado
'    End If
'End Sub

Private Sub TCanvas_onMouseUp(ByVal Button As Long, ByVal X As Double, ByVal Y As Double)
On Error GoTo Trata_Erro
   If Button = 0 Then 'BOTÃO ESQUERDO DO MOUSE
      'PopupMenu
   ElseIf Button = 1 Then
      Dim Lh As Double
'      TCanvas.getLengthOfLastSegmentOfLine Lh 'aqui dá erro qd não tenho o segmento de uma linha
'      frmNetWorkLegth.txtLength.Text = Lh
   End If
   Exit Sub
   
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
       
       PrintErro CStr(Me.Name), "Private Sub TCanvas_onMouseUp", CStr(Err.Number), CStr(Err.Description), True
    
    End If
End Sub
Private Sub TCanvas_onMoveGeometries(ByVal distance As Double, ByVal deltaX As Double, ByVal deltaY As Double)
    'MsgBox ("movendo geometria")        'para teste de movimentação de geometria e verificar quando entra aqui
End Sub
' Entra neste evento quando o usuário selecionou um vértice de uma linha para mover
' Ele vai chegar se não é o inicial ou final, pois estes só podem ser movidos pelo nó
'
' distanceSegment1 - comprimento do primeiro seguimento conectado no vértice da linha
' distanceSegment2 - comprimento do segundo seguimento conectado no vértice da linha
' distanceOldToNewPoint - distância do vértice antes de mover até a nova posição onde foi movido
'
Private Sub TCanvas_onMoveGeometryPoint(ByVal distanceSegment1 As Double, ByVal distanceSegment2 As Double, ByVal distanceOldToNewPoint As Double)
    If distanceSegment2 = 0 Or distanceSegment1 = 0 Then
        varGlobais.moverVertice = False                     'indica que o usuário selecionou a extremidade de uma rede e não pode mover pois é um nó
        MsgBox "Você não pode mover um nó com a ferramenta de mover vértice. Selecione o layer nó e em seguida a ferramenta de mover nós."
    Else
        varGlobais.moverVertice = True                      'é um vértice, não sendo o ponto inicial nem final
    End If
End Sub

Private Sub TCanvas_onPoint(ByVal X As Double, ByVal Y As Double)
On Error GoTo Trata_Erro
   Select Case Tr.TerraEvent
      Case tg_DrawNetWorkNode
         Tr.SaveInDatabase
'      Case tg_DrawPoint
'         Tr.OnPoint x, y
      Case tg_DrawGeometrys
         Tr.OnPoint X, Y
      Case tg_DrawRamal
         Tr.OnRamal X, Y, True
   End Select
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
   Else
      PrintErro CStr(Me.Name), "Private Sub TCanvas_onPoint", CStr(Err.Number), CStr(Err.Description), True
      
   End If

End Sub
'###########################################################################
'ROTINA QUE SALVA OS DADOS VERTORIAIS DAS REDES
'LINE_ID = NOVA LINHA
'NODE_ID1 = NOVO NÓ OU NÓ EXISTENTE
'NODE_ID2 = NOVO NÓ
'###########################################################################
'
'APÓS O SAVEINDATABASE O TECANVAS RETORNA, ATRAVÉS DO MÉTODO OnSaveNetworkLine
'OS CÓDIGOS GEOM_ID DOS PONTOS CRIADOS E O CODIGO LINE_ID DA LINHA CRIADA.
'SENDO NOVOS OU NÃO, OS CÓDIGOS DE PONTOS SÃO RETORNADOS
'SE RETORNADO 0(ZERO) PARA ALGUM PONTO, SIGNIFICA QUE A REDE ESTÁ SENDO MOVIDA
'NESTE CASO ENTÃO DEVE SER EXCLUIDO E REFEITO O TEXTO DAS LINHAS QUE ESTÃO SENDO MOVIDAS
'
' LINE_ID - Id da linha de rede que foi desenhada
' Node_id1 - Id do nó inicial que foi desenhado
' Node_id2 - Id do nó final que foi desenhado
'
Private Sub TCanvas_onSaveNetWorkLine(ByVal LINE_ID As Long, ByVal Node_id1 As Long, ByVal Node_id2 As Long)
    On Error GoTo Trata_Erro
    Dim TbGeometriaLinhas As String
    Dim TbGeometriaPontos As String
    Dim CompCalc As Long
    Dim CompCalc2 As Long
    Dim CompCalc3 As Double
    Dim LayerName As String
    Dim RefLayer As String
    
    a = "LENGTHCALCULATED"
    b = "USUARIO_LOG"
    c = "INSCRICAO_LOTE"
    d = "DATA_LOG"
    e = "OBJECT_ID_"
    f = Replace(Round(CompCalc3, 2), ",", ".")
    g = Format(Now, "DD/MM/YY HH:MM")
    h = RefLayer
    X1 = 0 ' ZERA A COORDENADA DE PRIMEIRO CLIQUE USADA PARA CALCULO DA DISTÂNCIA
    LayerName = TCanvas.getCurrentLayer
    RefLayer = TCanvas.GetReferenceLayer
    If Node_id1 = 0 Or Node_id2 = 0 Then 'ESTA MOVENDO A REDE. Sempre quando move ele entra com os objects_ids dos nós com zero para indicar movendo, vindo apenas
        'Call objIDsRamais.getObjIDs(LINE_ID, TeDatabase4, listObjIDsRamais)                             'obtem todos os objIDs dos ramais que estão ligados ao trecho de rede que está sendo movido

        If RefLayer = "WATERLINES" Then                                 'aqui inicializa o redesenho dos ramais de água na nova posição
            Dim contTrechos As Integer
            Dim contRamais As Integer
            Dim totalRamais As Integer
            Dim xIniRa As Double
            Dim yIniRa As Double
            
            totalRamais = UBound(ramalMovendo)
            'For contTrechos = 0 To varGlobais.totalTrechos
                For contRamais = 0 To totalRamais
                    If ramalMovendo(contRamais).objIdTrecho = LINE_ID And ramalMovendo(contRamais).objIdRamal <> -1 And ramalMovendo(contRamais).objIdTrecho = LINE_ID Then
                        Dim distIniRamalDepois As Double                'distância do início do ramal depois de tanto o trecho quanto o ramal serem movidos
                        Dim moveRamal As New CCoordIniRamalDistTrecho   'classe para obter a coordenada inicial do ramal a uma determinada distância do início do trecho de rede
                        Dim distEquiv As New CDistanciaEquivalente      'classe para obter a distância do início do ramal ao início do trecho após movido os mesmos
                        Dim retorno As Boolean
                        Dim novoComprTrecho As Double
                        Dim xRamal(1) As Double, yRamal(1) As Double
                        Dim comprimentoRamal As Double                  'comprimento calculado da extensão do ramal
                        Dim pontoSobreLinha As Long                  'indica se o ponto de início do ramal ficou ou não sobre a linha
                        
                        pontoSobreLinha = True
                        
                        cGeoDatabase.geoDatabase.setCurrentLayer ("Waterlines")
                        retorno = cGeoDatabase.geoDatabase.getLengthOfLine(LINE_ID, "", novoComprTrecho)
                        distIniRamalDepois = distEquiv.distanciaRamalDepoisMovido(ramalMovendo(contRamais).comprTrecho, novoComprTrecho, ramalMovendo(contRamais).Distancia)
                        'moveRamal.coordsRamal distIniRamalDepois, CStr(LINE_ID), cGeoDatabase.geoDatabase       'obtem as novas coordenadas inicial e final do ramal movido após mover o trecho de rede. Desativada, pois foi substituído pelo ponto perpendicular
                        retorno = cGeoDatabase.geoDatabase.getMinimumDistance(0, ramalMovendo(contRamais).objIdTrecho, 2, ramalMovendo(contRamais).xHidrom, ramalMovendo(contRamais).yHidrom, comprimentoRamal, pontoSobreLinha, xIniRa, yIniRa)    'obtem a nova coordenada inicial do ramal, perpendicular ao segmento de linha mais próximo
                        
                        xRamal(0) = xIniRa
                        yRamal(0) = yIniRa
                        xRamal(1) = ramalMovendo(contRamais).xHidrom                                           'estas coordenadas foram testadas e estão corretas, bate com a coordenada onde está o ponto (nó) do hidrômetro
                        yRamal(1) = ramalMovendo(contRamais).yHidrom
                        cGeoDatabase.geoDatabase.setCurrentLayer ("RAMAIS_AGUA")                                'seta o layer em que serão apagadas e adicionadas as geometrias
                        cGeoDatabase.geoDatabase.deleteGeometry ramalMovendo(contRamais).geomIdRamal, ramalMovendo(contRamais).objIdRamal, 2
                        cGeoDatabase.geoDatabase.addLine ramalMovendo(contRamais).objIdRamal, xRamal(0), yRamal(0), 2
                    End If
                Next
            'Next
        End If
        'finaliza

        'CALCULAR O NOVO COMPRIMENTO DA LINHA E ATUALIZAR NA BASE
        TeDatabase1.setCurrentLayer RefLayer
        'OBTEM NA VARIÁVEL CompCalc O COMPRIMENTO DA LINHA
        TeDatabase1.getLengthOfLine LINE_ID, CStr(LINE_ID), CompCalc3
        If frmCanvas.TipoConexao <> 4 Then
            'ATUALIZAR O COMPRIMENTO DA REDE, USUÁRIO E DATA DE ATUALIZAÇÃO
            Conn.execute ("UPDATE " & RefLayer & " SET LENGTHCALCULATED = " & Replace(Round(CompCalc3, 2), ",", ".") & ", USUARIO_LOG = '" & strUser & "', DATA_LOG = '" & Format(Now, "DD/MM/YY HH:MM") & "' WHERE OBJECT_ID_ = '" & LINE_ID & "'")
        Else
            'MsgBox "UPDATE  " + """" + h + """" + "SET " + """" + a + """" + " =  '" & Replace(Round(CompCalc3, 2), ",", ".") & "', " + """" + b + """" + " = '" & strUser & "', " + """" + d + """" + "= '" & Format(Now, "DD/MM/YY HH:MM") & "' WHERE " + """" + e + """" + " = '" & LINE_ID & "'"
            'UPDATE "DRAINLINES" SET "LENGTHCALCULATED" = CAST(regexp_replace ('34', '3', '1') As Integer), "USUARIO_LOG" = 'Administrador', "DATA_LOG" = 'Format(Now, "DD/MM/YY HH:MM")' WHERE "OBJECT_ID_" = '5'
            Conn.execute ("UPDATE  " + """" + RefLayer + """" + "SET " + """" + a + """" + " =  '" & Replace(Round(CompCalc3), ",", ".") & "', " + """" + b + """" + " = '" & strUser & "', " + """" + d + """" + "= '" & Format(Now, "DD/MM/YY HH:MM") & "' WHERE " + """" + e + """" + " = '" & LINE_ID & "'")
        End If
        'CHAMA O MÉTODO DE EXCLUIR E CRIAR TEXTOS DENTRO DO MÉTODO Tr.CreatNetWorkAttribute
        Tr.CreatNetWorkAttribute LINE_ID, Node_id1, Node_id2, True
        FrmMain.sbStatusBar.Panels(1).Text = "Rede " & LINE_ID & " movida com sucesso."
    Else  'ESTÁ DESENHANDO A REDE
        Dim JaExisteRede As Boolean
        Dim rs As ADODB.Recordset
        JaExisteRede = False
        TbGeometriaLinhas = LCase(TeDatabase1.getRepresentationTableName(TCanvas.getCurrentLayer, tpLINES))
        TbGeometriaPontos = LCase(TeDatabase1.getRepresentationTableName(TCanvas.GetReferenceLayer, tpPOINTS))
                'VERIFICA SE NÃO JA EXISTE UMA REDE COM ESTES MESMOS NÓS INICIAIS E FINAIS
                Set rs = New ADODB.Recordset 'alterado em 20/10/2010
                Dim dt As String
                Dim dm As String
                Dim dg As String
                Dim dv As String
                dv = "OBJECT_ID_"
                dt = "INITIALCOMPONENT"
                dm = "FINALCOMPONENT"
                dg = "d"
                'aqui ele vai verificar se a rede que está sendo desenhada, está sendo desenhada por cima de outra, tanto num sentido quanto no outro
                If frmCanvas.TipoConexao <> 4 Then
                    rs.Open ("SELECT OBJECT_ID_ FROM " & LayerName & " WHERE INITIALCOMPONENT = '" & Node_id1 & "' AND FINALCOMPONENT = '" & Node_id2 & "'"), Conn, adOpenForwardOnly, adLockReadOnly
                Else
                    rs.Open ("SELECT " + """" + dv + """" + " FROM " + """" + LayerName + """" + " WHERE " + """" + dt + """" + " = '" & Node_id1 & "' AND " + """" + dm + """" + " = '" & Node_id2 & "'"), Conn, adOpenDynamic, adLockOptimistic
                End If
                If rs.EOF = False Then
                    JaExisteRede = True
                Else
                    Set rs = New ADODB.Recordset
                    If frmCanvas.TipoConexao <> 4 Then
                        rs.Open ("SELECT OBJECT_ID_ FROM " & LayerName & " WHERE FINALCOMPONENT = '" & Node_id1 & "' AND INITIALCOMPONENT = '" & Node_id2 & "'"), Conn, adOpenForwardOnly, adLockReadOnly
                    Else
                        rs.Open ("SELECT " + """" + dv + """" + " FROM " + """" + LayerName + """" + " WHERE " + """" + dm + """" + " = '" & Node_id1 & "' AND " + """" + dt + """" + " = '" & Node_id2 & "'"), Conn, adOpenDynamic, adLockOptimistic
                    End If
                    If rs.EOF = False Then
                        JaExisteRede = True
                    End If
                End If
                rs.Close
                If JaExisteRede = True Then
                    MsgBox "Já existe uma rede desenhada entre estas 2 peças.", vbExclamation, ""
                    'DELETA GEOMETRIA DE LINHA QUE FOI CRIADA
                    If frmCanvas.TipoConexao <> 4 Then
                        Conn.execute ("DELETE FROM " & TbGeometriaLinhas & " WHERE GEOM_ID = " & LINE_ID)
                    Else
                        Dim ga As String
                        ga = "geom_id"
                        Conn.execute ("DELETE FROM " + """" + TbGeometriaLinhas + """" + " WHERE " + """" + ga + """" + " = '" & LINE_ID & "'")
                    End If
                    FrmMain.sbStatusBar.Panels(1).Text = "Rede " & LINE_ID & " não criada."
                    'SAI DO EVENTO
                    Exit Sub
                End If
                'termina a verificação (e aviso se for o caso) de que a rede está sendo desenhada sobre outra já existente
        a = "tbgeometrialinhas"
        b = "object_id"
        c = "geom_id"
        'ATUALIZA OS OBJECTS_ID COM O MESMO CÓDIGO DO AUTO NUMERADOR
        If frmCanvas.TipoConexao <> 4 Then
            Conn.execute ("UPDATE " & TbGeometriaLinhas & " SET OBJECT_ID = GEOM_ID WHERE GEOM_ID = " & LINE_ID)
            Conn.execute ("UPDATE " & TbGeometriaPontos & " SET OBJECT_ID = GEOM_ID WHERE GEOM_ID = " & Node_id1)
            Conn.execute ("UPDATE " & TbGeometriaPontos & " SET OBJECT_ID = GEOM_ID WHERE GEOM_ID = " & Node_id2)
        Else
            Conn.execute ("UPDATE " + """" & TbGeometriaLinhas & """" + " SET " + """" + "object_id" + """" + " = " + """" + "geom_id" + """" + " WHERE " + """" + "geom_id" + """" + " =  '" & LINE_ID & "'")
            Conn.execute ("UPDATE " + """" & TbGeometriaPontos & """" + " SET " + """" + "object_id" + """" + " = " + """" + "geom_id" + """" + " WHERE " + """" + "geom_id" + """" + " =  '" & Node_id1 & "'")
            Conn.execute ("UPDATE " + """" & TbGeometriaPontos & """" + " SET " + """" + "object_id" + """" + " = " + """" + "geom_id" + """" + " WHERE " + """" + "geom_id" + """" + " =  '" & Node_id2 & "'")
        End If
        Tr.CreatNetWorkAttribute LINE_ID, Node_id1, Node_id2, False
        FrmMain.sbStatusBar.Panels(1).Text = "Rede " & LINE_ID & " salva com sucesso."
    End If
    Exit Sub

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
   Else
      varGlobais.realizaCommit = False                      'pede para voltar tudo o que está fazendo no banco de dados, para traz e não comitar nada
      ErroUsuario.Registra "frmCanvas", "onSaveNetWorkLine", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
   End If
End Sub

Private Sub TCanvas_onSaveNetWorkNode(ByVal node_id As Long, ByVal line1_id As Long, ByVal line2_id As Long)
On Error GoTo Trata_Erro

a = "TBGEOMETRIALINHAS"
b = "object_id"
c = "geom_id"
d = "TBGEOMETRIAPONTOS"



   'AO INSERIR OU MOVER UM NÓ DE REDE EM UMA REDE JA EXISTENTE, ENTRA NESTE MÉTODO
   'O NODE_ID É O CÓDIGO DO NOVO NÓ, LINE1_ID É A LINHA JA EXISTENTE E
   'A LINE2_ID É A GEOMETRIA SALVA PELO TE_CANVAS PARA A NOVA LINHA
   'CASO SEJA RETORNADO 0(ZERO) PARA ALGUMA DAS LINHAS SIGNIFICA QUE O NÓ DE REDE FOI MOVIDO
   
TCanvas.ToolTipText = ""
X1 = 0 ' ZERA A COORDENADA DE PRIMEIRO CLIQUE USADA PARA CALCULO DA DISTÂNCIA

If line1_id = 0 Or line2_id = 0 Then 'O NÓ DE REDE FOI MOVIDO E DEVERÁ SOFRER ALTERAÇÕES SOMENTE SE FOR PEÇA DE ESGOTO

   Tr.CreatNetWorkNode node_id, line1_id, line2_id, True


Else

     Dim TbGeometriaPontos As String
     Dim TbGeometriaLinhas As String
     
   'ATUALIZA O OBJECT_ID DA LINHA RECEM CRIADA NA TABELA LINES
   
   
    If frmCanvas.TipoConexao <> 4 Then
   TbGeometriaLinhas = LCase(TeDatabase1.getRepresentationTableName(TCanvas.getCurrentLayer, tpLINES))
   
      Conn.execute ("UPDATE " & TbGeometriaLinhas & " SET OBJECT_ID = GEOM_ID WHERE GEOM_ID = " & line2_id)
   
     
   
   'ATUALIZA O OBJECT_ID DA POINTS COM O MESMO CÓDIGO DO AUTO NUMERADOR DO TeCanvas
   TbGeometriaPontos = LCase(TeDatabase1.getRepresentationTableName(TCanvas.GetReferenceLayer, tpPOINTS))
   Conn.execute ("UPDATE " & TbGeometriaPontos & " SET OBJECT_ID = GEOM_ID WHERE GEOM_ID = " & node_id)
  Else
    TbGeometriaLinhas = LCase(TeDatabase1.getRepresentationTableName(TCanvas.getCurrentLayer, tpLINES))
   
   
Conn.execute ("UPDATE " + """" + TbGeometriaLinhas + """" + " SET " + """" + b + """" + " = " + """" + c + """" + " WHERE " + """" + c + """" + " = '" & line2_id & " '")
     
   
   'ATUALIZA O OBJECT_ID DA POINTS COM O MESMO CÓDIGO DO AUTO NUMERADOR DO TeCanvas
   TbGeometriaPontos = LCase(TeDatabase1.getRepresentationTableName(TCanvas.GetReferenceLayer, tpPOINTS))
   Conn.execute ("UPDATE " + """" + TbGeometriaPontos + """" + " SET " + """" + b + """" + " = " + """" + c + """" + " WHERE " + """" + c + """" + " = '" & node_id & " '")
   End If
   
   Tr.CreatNetWorkNode node_id, line1_id, line2_id, False
   
   FrmMain.sbStatusBar.Panels(1).Text = "Componente " & node_id & " criado com sucesso."
   
End If
   
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
       
       PrintErro CStr(Me.Name), "Private Sub TCanvas_onSaveNetWorkLine", CStr(Err.Number), CStr(Err.Description), True
       
    End If
End Sub

Private Sub TCanvas_onSnap(ByVal distance1 As Double, ByVal distance2 As Double)
On Error GoTo Trata_Erro
    
    If FrmMain.tbToolBar.Buttons("kinsertnetworknode").value = tbrPressed Then
        Dim xmin As Double, xmax As Double, ymin As Double, ymax As Double
        Call TeDatabase1.getGeometryBox(0, TCanvas.getSelectObjectId(0, 2), tpLINES, xmin, ymin, xmax, ymax)
        If (Position_X >= xmin And Position_X <= xmax) And (Position_Y >= ymin And Position_Y <= ymax) Then
            txtRede1.Text = distance1
            txtRede2.Text = distance2
        End If
    End If
   
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
   Else
      
      PrintErro CStr(Me.Name), "Private Sub TCanvas_onSnap", CStr(Err.Number), CStr(Err.Description), True
   End If
End Sub



' Fica aguardando o usuário fazer alguma coisa
'
'
'
Private Sub TimerSetWorld_Timer()
    On Error GoTo Trata_Erro
    'timer para inicializar o método SetWorld do TeCanvas
    If xWorld > 0 And yWorld > 0 Then
        TCanvas.setWorld xWorld - 100, yWorld - 100, xWorld + 100, yWorld + 100
        If blnLocalizandoConsumidor = True Then
            blnLocalizandoConsumidor = False
            TCanvas.setScale 80
        End If
        xWorld = 0
        TCanvas.plotView
    End If
    If canvasScale > 0 Then
        TCanvas.setScale canvasScale
        canvasScale = 0
    End If
    'TimerSetWorld.Enabled = False
    
Trata_Erro:
End Sub

Public Function FunDecripta(ByVal strDecripta As String) As String


    Dim IntTam As Integer
    Dim i As Integer
    Dim letra As String
    IntTam = Len(strDecripta)
    nStr = ""

    'desconsidera os os numeros de HH-MM-SS
    strDecripta = mid(strDecripta, 6, 5) & mid(strDecripta, 16, 5) & mid(strDecripta, 26, 5) & _
                  mid(strDecripta, 36, 5) & mid(strDecripta, 46, 5) & mid(strDecripta, 56, 200)

    i = 1
    Do While Not i = IntTam - 29
        letra = mid(strDecripta, i, 5)
        Select Case letra
        Case "14334"
            nStr = nStr & "a"
        Case "14212"
            nStr = nStr & "A"
        Case "24334"
            nStr = nStr & "á"
        Case "24134"
            nStr = nStr & "â"
        Case "24234"
            nStr = nStr & "ã"
        Case "24314"
            nStr = nStr & "à"
        Case "24324"
            nStr = nStr & "b"
        Case "14223"
            nStr = nStr & "B"
        Case "11211"
            nStr = nStr & "ç"
        Case "11311"
            nStr = nStr & "Ç"
        Case "13334"
            nStr = nStr & "c"
        Case "14324"
            nStr = nStr & "C"
        Case "24344"
            nStr = nStr & "d"
        Case "14444"
            nStr = nStr & "D"
        Case "12314"
            nStr = nStr & "e"
        Case "21111"
            nStr = nStr & "E"
        Case "24321"
            nStr = nStr & "é"
        Case "32314"
            nStr = nStr & "ê"
        Case "31314"
            nStr = nStr & "f"
        Case "21311"
            nStr = nStr & "F"
        Case "32134"
            nStr = nStr & "g"
        Case "21341"
            nStr = nStr & "G"
        Case "31324"
            nStr = nStr & "h"
        Case "22111"
            nStr = nStr & "H"
        Case "32124"
            nStr = nStr & "i"
        Case "21112"
            nStr = nStr & "I"
        Case "31334"
            nStr = nStr & "í"
        Case "32333"
            nStr = nStr & "ì"
        Case "11314"
            nStr = nStr & "j"
        Case "23122"
            nStr = nStr & "J"
        Case "33134"
            nStr = nStr & "k"
        Case "23411"
            nStr = nStr & "K"
        Case "33314"
            nStr = nStr & "l"
       Case "32222"
            nStr = nStr & "L"
        Case "43423"
            nStr = nStr & "m"
        Case "32111"
            nStr = nStr & "M"
        Case "42423"
            nStr = nStr & "n"
        Case "33221"
            nStr = nStr & "N"
        Case "43234"
            nStr = nStr & "o"
        Case "33233"
            nStr = nStr & "O"
        Case "42444"
            nStr = nStr & "ô"
        Case "43223"
            nStr = nStr & "õ"
        Case "42433"
            nStr = nStr & "ò"
        Case "43231"
            nStr = nStr & "ó"
        Case "22223"
            nStr = nStr & "p"
        Case "33444"
            nStr = nStr & "P"
        Case "43233"
            nStr = nStr & "q"
        Case "34442"
            nStr = nStr & "Q"
        Case "43421"
            nStr = nStr & "r"
        Case "34332"
            nStr = nStr & "R"
        Case "13443"
            nStr = nStr & "s"
        Case "34222"
            nStr = nStr & "S"
        Case "44444"
            nStr = nStr & "t"
        Case "34112"
            nStr = nStr & "T"
        Case "13444"
            nStr = nStr & "u"
        Case "41311"
            nStr = nStr & "U"
        Case "11111"
            nStr = nStr & "ú"
        Case "13243"
            nStr = nStr & "ù"
        Case "11115"
            nStr = nStr & "û"
        Case "13241"
           nStr = nStr & "v"
        Case "41222"
            nStr = nStr & "V"
        Case "12443"
            nStr = nStr & "x"
        Case "41133"
            nStr = nStr & "X"
        Case "13244"
            nStr = nStr & "y"
        Case "42231"
            nStr = nStr & "Y"
        Case "13441"
            nStr = nStr & "w"
        Case "42222"
            nStr = nStr & "W"
        Case "11313"
            nStr = nStr & "z"
        Case "42213"
            nStr = nStr & "Z"
        Case "11312"
            nStr = nStr & "@"
        Case "11114"
            nStr = nStr & "%"
        Case "12341"
            nStr = nStr & "&"
        Case "13343"
            nStr = nStr & "*"
        Case "12342"
            nStr = nStr & "("
        Case "13344"
            nStr = nStr & ")"
        Case "12333"
            nStr = nStr & "$"
        Case "23334"
            nStr = nStr & "!"
        Case "13331"
            nStr = nStr & "#"
        Case "21242"
            nStr = nStr & "?"
        Case "22313"
            nStr = nStr & "1"
        Case "23424"
            nStr = nStr & "2"
        Case "24131"
            nStr = nStr & "3"
        Case "41414"
            nStr = nStr & "4"
        Case "22314"
           nStr = nStr & "5"
        Case "23423"
            nStr = nStr & "6"
        Case "44134"
            nStr = nStr & "7"
        Case "21241"
            nStr = nStr & "8"
       Case "22312"
           nStr = nStr & "9"
       Case "23231"
            nStr = nStr & "0"
        Case "34123"
            nStr = nStr & " "
        Case "14121"
            nStr = nStr & "_"
        Case "14144"
            nStr = nStr & "/"
        Case "12131"
            nStr = nStr & "\"
        Case "12124"
            nStr = nStr & "-"
        Case "21421"
            nStr = nStr & ";"
        Case "21321"
            nStr = nStr & ":"
        Case "14431"
            nStr = nStr & ","
        Case "13421"
            nStr = nStr & "."
        Case "11213"
            nStr = nStr & "+"
        Case "11212"
            nStr = nStr & "="

        Case Else
            MsgBox "Código de criptografia inválido!"
            'mStrDeCriptografa = ""
            Exit Function
        End Select
        i = i + 5
    Loop
  FunDecripta = nStr
    'mStrDeCriptografa = nStr

Exit Function
End Function



