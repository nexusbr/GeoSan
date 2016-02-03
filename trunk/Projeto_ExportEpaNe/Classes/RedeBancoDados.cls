VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RedeBancoDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_oConexao As New ADODB.Connection                          'conexão com o banco de dados para poder ler as linhas das tabelas do mesmo
Private rs_trechosDeRede As New ADODB.Recordset                     'lista de trechos de rede que foram selecionados pela cerca pelo usuário do GeoSan para serem exportados para o Epanet
Private usuarioQueExportouRedes As String                           'nome do usuário que exportou as redes do GeoSan
Private rsCoordinates As New ADODB.Recordset
Private rsPipes As New ADODB.Recordset
Private rsJunctions As New ADODB.Recordset                          'vetor com todas as junções (nós de conexão) que é utilizado depois de completo para exportar para o EPANET
Private rsPumps As New ADODB.Recordset
Private rsValves As New ADODB.Recordset
Private rsReservoirs As New ADODB.Recordset
Private rsTanks As New ADODB.Recordset
Private rsVertices As New ADODB.Recordset                           'Vertices da linha com exceção do inicial e final
Private rsNosTmp As New ADODB.Recordset                             'Criado para armazenar todos os dados de todos nos - Copia do Watercomponenstes/Points
Private rsTrechosExportados As New ADODB.Recordset                  'Criado para armazenar os trechos que já foram exportados
Private RsTrechos As New ADODB.Recordset
Private rsNosExportados As New ADODB.Recordset                       'Criado para armazenar os Nós que já foram exportados
Private rsWaterCompTypes As New ADODB.Recordset                      'RecordSet de referencia para consulta do tipo de componente
Private rs_global As ADODB.Recordset

'Abre a conexão com o banco de dados para obter os trechos de rede que serão exportados para o EPANET
'
'
'
Public Function RsGlobal_abreconexaoComTrechosAExportar()
    Dim str As String
    
    'Agora temos que descobrir todos os trechos de rede (TIPO=1 na tabela POLIGONO_SELECAO) que serão exportados para o Epanet, do usuário ativo
    str = "SELECT * FROM WATERLINES INNER JOIN X_MATERIAL ON WATERLINES.MATERIAL = X_MATERIAL.MATERIALID"
    str = str & " WHERE WATERLINES.OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM POLIGONO_SELECAO WHERE USUARIO = '" & usuario_global & "' AND TIPO = 1)"
    
    'Agora que temos trechos a serem exportados, vamos exportar para o Epanet
    Set rs_global = New ADODB.Recordset
    rs_global.Open str, conn, adOpenDynamic, adLockReadOnly
    conn.CommandTimeout = 300
End Function


'Recebe a conexão com o banco de dados a partir de instanciada a classe e passada a conexão para a mesma, armazenando a conexão até o final da aplicação
'Isto permitirá todas as operações com banco de dados
'
'
'
Property Set Conexao(oConexao As ADODB.Connection)
    Set m_oConexao = oConexao
End Property

'Obtém o número do nó inicial do trecho de rede
'
'
'
Property Get rs_trechosDeRede_numeroNoInicial() As String
    rs_trechosDeRede_numeroNoInicial = rs_trechosDeRede.Fields("InitialComponent").Value
End Property

'Obtém o número do nó final do trecho de rede
'
'
'
Property Get rs_trechosDeRede_numeroNoFinal() As String
    rs_trechosDeRede_numeroNoFinal = rs_trechosDeRede.Fields("FinalComponent").Value
End Property

'Obtem o object_id (numero) do trecho de rede
'
'
'
Property Get rs_trechosDeRede_ObjectIdDoTrecho() As String
    rs_trechosDeRede_ObjectIdDoTrecho = rs_trechosDeRede.Fields("object_id_").Value
End Property

'Abre a conexão com os trechos de rede que serão exportados para o EPANET
'
'
Function IniciaLeituraTrechosRede()
    Dim str As String
    
    str = "SELECT * FROM WATERLINES INNER JOIN X_MATERIAL ON WATERLINES.MATERIAL = X_MATERIAL.MATERIALID"       'Agora temos que descobrir todos os trechos de rede (TIPO=1 na tabela POLIGONO_SELECAO) que serão exportados para o Epanet, do usuário ativo
    str = str & " WHERE WATERLINES.OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM POLIGONO_SELECAO WHERE USUARIO = '" & usuarioQueExportouRedes & "' AND TIPO = 1)"
    rs_trechosDeRede.Open str, m_oConexao, adOpenDynamic, adLockReadOnly                                        'Abre a conexão com o banco de dados para obter os trechos de rede que serão exportados para o EPANET
    conn.CommandTimeout = 300
End Function

'Verifica se terminou de ler todos os trechos de rede selecionados pelo usuário
'
' Rs_trechosDeRede_terminouDeLer - True = Terminou; False = Não terminou de ler todos os trechos
'
Public Function Rs_trechosDeRede_terminouDeLer() As Boolean
    If rs_trechosDeRede.EOF = True Then
        Rs_trechosDeRede_terminouDeLer = True
    Else
        Rs_trechosDeRede_terminouDeLer = False
    End If
End Function

'Obtem a quantidade total de trechos de rede que serão processados
'
'
Function ObtemNumeroTrechosQueSeraoExportados() As Integer
    Dim Rs As ADODB.Recordset
    Dim strtot As String                        'armazena a querie para obter o número total de segmentos de rede que serão exportados para o Epanet (TIPO=1 na tabela POLIGONO_SELECAO)
    Dim totalTrechosExportar As Integer         'número total de trechos de rede de água que serão exportados para o Epanet, disponíveis em (TIPO=1 na tabela POLIGONO_SELECAO)
    
    'Agora temos que descobrir todos os trechos de rede (TIPO=1 na tabela POLIGONO_SELECAO) que serão exportados para o Epanet, do usuário ativo
    strtot = "SELECT COUNT(*) FROM WATERLINES INNER JOIN X_MATERIAL ON WATERLINES.MATERIAL = X_MATERIAL.MATERIALID"
    strtot = strtot & " WHERE WATERLINES.OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM POLIGONO_SELECAO WHERE USUARIO = '" & usuarioQueExportouRedes & "' AND TIPO = 1)"
    Set Rs = New ADODB.Recordset
    'Obtem o número total de trechos de rede a serem exportados para o Epanet os quais são do TIPO=1 na tabela POLIGONO_SELECAO
    Rs.Open strtot, m_oConexao, adOpenDynamic, adLockReadOnly
    totalTrechosExportar = Rs(0).Value      'obtem o número total de trechos de rede que serão exportados
    Set Rs = Nothing
    ObtemNumeroTrechosQueSeraoExportados = totalTrechosExportar
End Function

'Obtem o nome do usuário que está exportando para o EPANET
'
'
Public Function ObtemNomeUsuario()
    Dim retval As String
    Dim usuario As String

    retval = Dir(App.Path & "\Controles\UserLog.txt")
   
    'verifica se o arquivo existe na pasta
    If retval <> "" Then
        'Abre e lê o arquivo para ver que usuário será consultado no polígono selecionado, pois podem existir vários usuários realizando esta operação ao mesmo tempo
        Open App.Path & "\Controles\UserLog.txt" For Input As #3
        Line Input #3, usuario
        Close #3
    Else
        'Avisa e cai fora, pois não dá para executar a operação
        MsgBox "É necessário criar a seleção por polígono.", vbOKOnly + vbInformation, "Mensagem"
        End
    End If
    usuarioQueExportouRedes = usuario
End Function

'Atualiza todas as rugosidades de todas as tubulações, conforme o tipo de material. Foi considerada uma tubulação de 20 anos de idade
'Sempre que alguma linha for alterada na tabela X_Material, estas rugosidades deverão ser revistas
'
'
'
Public Function AtualizaAsRugosidades()
    'Liga a ampulheta no ponteiro do mouse
    MousePointer = vbHourglass
    m_oConexao.Execute ("UPDATE WATERLINES SET ROUGHNESS = 0")                                            'zera todas as rugozidades
    m_oConexao.Execute ("UPDATE WATERLINES SET ROUGHNESS = 111 WHERE MATERIAL = 0")                       'Desconhecido
    m_oConexao.Execute ("UPDATE WATERLINES SET ROUGHNESS = 130 WHERE MATERIAL = 1 AND ROUGHNESS = 0")     'Cerâmico
    m_oConexao.Execute ("UPDATE WATERLINES SET ROUGHNESS = 120 WHERE MATERIAL = 2 AND ROUGHNESS = 0")     'Concreto
    m_oConexao.Execute ("UPDATE WATERLINES SET ROUGHNESS = 110 WHERE MATERIAL = 3 AND ROUGHNESS = 0")     'FF
    m_oConexao.Execute ("UPDATE WATERLINES SET ROUGHNESS = 105 WHERE MATERIAL = 4 AND ROUGHNESS = 0")     'Fibrocimento
    m_oConexao.Execute ("UPDATE WATERLINES SET ROUGHNESS = 90  WHERE MATERIAL = 5 AND ROUGHNESS = 0")     'PEAD
    m_oConexao.Execute ("UPDATE WATERLINES SET ROUGHNESS = 130 WHERE MATERIAL = 6 AND ROUGHNESS = 0")     'PRSV
    m_oConexao.Execute ("UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 7 AND ROUGHNESS = 0")     'PVC branco
    m_oConexao.Execute ("UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 8 AND ROUGHNESS = 0")     'PVC corrugado
    m_oConexao.Execute ("UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 9 AND ROUGHNESS = 0")     'PVC DeFoFo
    m_oConexao.Execute ("UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 10 AND ROUGHNESS = 0")    'PVC Ocre
    m_oConexao.Execute ("UPDATE WATERLINES SET ROUGHNESS = 140 WHERE MATERIAL = 11 AND ROUGHNESS = 0")    'PVC PBA
    'Volta o mouse para o normal
    FrmEPANET.MousePointer = vbDefault
End Function

'Zera todos os materiais onde os mesmos estiverem com valor nulo
'
'
'
Public Function ZeraMateriais()
    m_oConexao.Execute ("UPDATE WATERLINES SET MATERIAL = 0 WHERE MATERIAL IS NULL")                        'Zera todos os materiais de tubulações quando o mesmo não estiver cadastrado
End Function

'Define a estrutura dos vetores que conterão os dados que serão exportados para o Epanet
'
'
'
Sub AbrirEstruturaExporteRede()
    'coordenadas dos nós
    rsCoordinates.Fields.Append "id", adVarChar, 255            'número do nó
    rsCoordinates.Fields.Append "x", adDouble                   'coordenada X
    rsCoordinates.Fields.Append "y", adDouble                   'coordenada Y
    rsCoordinates.Open
    
    'tubulações
    rsPipes.Fields.Append "id", adVarChar, 255                  'número da tubulação
    rsPipes.Fields.Append "node1", adVarChar, 255
    rsPipes.Fields.Append "node2", adVarChar, 255
    rsPipes.Fields.Append "length", adVarChar, 255
    rsPipes.Fields.Append "diameter", adDouble, 255
    rsPipes.Fields.Append "roughness", adDouble, 255
    rsPipes.Fields.Append "minorloss", adVarChar, 255
    rsPipes.Fields.Append "status", adVarChar, 255
    rsPipes.Fields.Append "Description", adVarChar, 255         'incluido em 13/05/2009 Jonathas
    rsPipes.Open
    
    'junções
    rsJunctions.Fields.Append "id", adVarChar, 255
    rsJunctions.Fields.Append "elev", adVarChar, 255
    rsJunctions.Fields.Append "demand", adDouble, 255
    rsJunctions.Fields.Append "pattern", adVarChar, 255
    rsJunctions.Open
    
    'bombas
    rsPumps.Fields.Append "id", adVarChar, 255
    rsPumps.Fields.Append "node1", adVarChar, 255
    rsPumps.Fields.Append "node2", adVarChar, 255
    rsPumps.Fields.Append "parameters", adVarChar, 255
    rsPumps.Fields.Append "potencia", adVarChar, 255
    rsPumps.Open
    
    'válvulas
    rsValves.Fields.Append "id", adVarChar, 255
    rsValves.Fields.Append "node1", adVarChar, 255
    rsValves.Fields.Append "node2", adVarChar, 255
    rsValves.Fields.Append "diameter", adVarChar, 255
    rsValves.Fields.Append "type", adVarChar, 255
    rsValves.Fields.Append "setting", adVarChar, 255
    rsValves.Fields.Append "minorloss", adVarChar, 255
    rsValves.Open
    
    'reservatórios
    rsReservoirs.Fields.Append "ID", adVarChar, 255
    rsReservoirs.Fields.Append "Head", adVarChar, 255
    rsReservoirs.Fields.Append "Pattern", adVarChar, 255
    rsReservoirs.Open
    
    'tanques
    rsTanks.Fields.Append "ID", adVarChar, 255
    rsTanks.Fields.Append "Elevation", adVarChar, 255
    rsTanks.Fields.Append "InitLevel", adVarChar, 255
    rsTanks.Fields.Append "MinLevel", adVarChar, 255
    rsTanks.Fields.Append "MaxLevel", adVarChar, 255
    rsTanks.Fields.Append "Diameter", adVarChar, 255
    rsTanks.Fields.Append "MinVol", adVarChar, 255
    rsTanks.Fields.Append "VolCurve", adVarChar, 255
    rsTanks.Open
    
    'vértices de linhas de tubulações
    rsVertices.Fields.Append "ID", adVarChar, 255               'número da tubulação
    rsVertices.Fields.Append "X-Coord", adDouble
    rsVertices.Fields.Append "Y-Coord", adDouble
    rsVertices.Open
    
    'nós
    rsNosTmp.Fields.Append "ID", adVarChar, 255
    rsNosTmp.Fields.Append "X", adDouble
    rsNosTmp.Fields.Append "Y", adDouble
    rsNosTmp.Fields.Append "Tipo", adInteger
    rsNosTmp.Fields.Append "Padrao", adInteger
    rsNosTmp.Fields.Append "Curva", adInteger
    rsNosTmp.Fields.Append "Diametro", adVarChar, 255
    rsNosTmp.Fields.Append "Cota", adDouble
    rsNosTmp.Fields.Append "NivelMin", adDouble
    rsNosTmp.Fields.Append "NivelMax", adDouble
    rsNosTmp.Fields.Append "VolumeMin", adDouble
    rsNosTmp.Fields.Append "CurvaVol", adDouble
    rsNosTmp.Fields.Append "Parametros", adDouble
    rsNosTmp.Fields.Append "setting", adDouble
    rsNosTmp.Fields.Append "SubTypeValve", adDouble
    rsNosTmp.Fields.Append "demanda", adDouble
    rsNosTmp.Fields.Append "estado", adVarChar, 255
    rsNosTmp.Fields.Append "Description", adVarChar, 255
    rsNosTmp.Open
    
    'lista de trechos exportados
    rsTrechosExportados.Fields.Append "id", adVarChar, 255
    rsTrechosExportados.Open
    
    RsTrechos.Fields.Append "id", adVarChar, 255
    RsTrechos.Open
    
    'lista de nós exportados
    rsNosExportados.Fields.Append "id", adVarChar, 255
    rsNosExportados.Open
End Sub
'Gera um vetor temporário de nós com seus atributos, como o objetivo de facilitar a leitura dos dados dos nós da rede
'Cria uma cópia da query da tabela watercomponents + points para RsNosTMP com todos os nos das tabelas relacionadas
'
'arquivoLog - nome do arquivo de logo onde está sendo exportado o log do Epanet
'
Function CarregaRsNosTMP()

On Error GoTo Trata_Erro
    Dim layer_id As Long
    Dim strSQL As String
    
    layer_id = GetLayerID("WATERCOMPONENTS")
   
    'Gera a query desnormatizada junto aos nos(Watercomponents) para facilitar a leitura dos dados dos mesmos
    'Select a.OBJECT_ID_, X, Y, ID_TYPE, GROUNDHEIGHT, DEMAND, Pattern, SubTypeValve,
    'case when State = 2 then 'Closed' else 'Open' end state FROM (Select OBJECT_ID_, X, Y, ID_TYPE,
    'GROUNDHEIGHT, DEMAND, Pattern FROM watercomponents inner join points2 on object_id_=object_id)
    ' a Left Join (select object_id_,value_ as SubTypeValve from watercomponentsdata  where id_type = 1
    'and id_subtype = 1) b on a.object_id_=b.object_id_  left Join (select object_id_,value_ as State
    'from watercomponentsdata  where id_type = 1 and id_subtype = 2) c on a.object_id_=c.object_id_

    ' * Alguns números acima são variáveis na query a seguir
    'Exemplo de resultado da query:
    'OBJECT_ID_; X; Y; TIPO DE COMPONENTE; COTA; DEMANDA DE CONSUMO;PADRÃO;ESTADO
    '100     289716,2251315639   9110857,324804159   25  0,  0,  0       NULL    Open
    '10000   291963,3551800701   9110854,729955614   0   0,  0,  NULL    NULL    Open
    '10001   291975,6117865313   9110853,035953095   0   0,  0,  NULL    NULL    Open
    '10002   291986,8719209225   9110851,24230337    0   0,  0,  NULL    NULL    Open
    '10003   291991,2563980305   9110857,021841375   0   0,  0,  NULL    NULL    Open

    strSQL = " Select a.OBJECT_ID_"
    strSQL = strSQL & ", x, y, ID_TYPE, GROUNDHEIGHT, DEMAND, Pattern, SubTypeValve, case when State = 2 then 'Closed' else 'Open' end state"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & "(Select OBJECT_ID_"
    strSQL = strSQL & ", X, Y, ID_TYPE, GROUNDHEIGHT, DEMAND, Pattern"
    strSQL = strSQL & " FROM watercomponents inner join points" & layer_id & " on object_id_=object_id) a"
    strSQL = strSQL & " Left Join"
    strSQL = strSQL & " (select object_id_,value_ as SubTypeValve from watercomponentsdata  where id_type = 1 and id_subtype = 1) b"
    strSQL = strSQL & " on a.object_id_=b.object_id_"
    strSQL = strSQL & "  left Join (select object_id_,value_ as State from watercomponentsdata  where id_type = 1 and id_subtype = 2) c"
    strSQL = strSQL & " on a.object_id_=c.object_id_"

    Set rsNos = New ADODB.Recordset
    
    rsNos.Open strSQL, m_oConexao ', adOpenDynamic, adLockReadOnly

    While Not rsNos.EOF
        With rsNosTmp
          .AddNew
          .Fields("ID").Value = rsNos.Fields("OBJECT_ID_").Value
          .Fields("X").Value = rsNos.Fields("x").Value
          .Fields("Y").Value = rsNos.Fields("y").Value
          .Fields("Tipo").Value = IIf(IsNull(rsNos.Fields("id_type").Value), 0, rsNos.Fields("id_type").Value)
          If rsNos.Fields("ID_TYPE").Value = No_Valvulas Then
             Select Case rsNos.Fields("SubTypeValve").Value
                Case 4, 0
                   .Fields("Tipo").Value = 1
                Case Else
                   .Fields("Tipo").Value = 99
             End Select
          End If
            .Fields("Cota").Value = IIf(IsNull(rsNos.Fields("GROUNDHEIGHT").Value), 0, rsNos.Fields("GROUNDHEIGHT").Value)
            .Fields("Demanda").Value = IIf(IsNull(rsNos.Fields("demand").Value), 0, rsNos.Fields("demand").Value)
            .Fields("Padrao").Value = IIf(IsNull(rsNos.Fields("PATTERN").Value), 0, rsNos.Fields("PATTERN").Value)
            .Fields("estado").Value = rsNos.Fields("state").Value
       End With
       rsNos.MoveNext
    Wend

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        Close #2
        Open App.Path & "\LogErroExportEPANET.txt" For Append As #2
        Print #2, Now & "  - ModExporte - Sub CarregaRsNosTMP() - Linha: " & intLinhaCod & " - " & Err.Number & " - " & Err.Description
        Close #2
        MsgBox "Um posssível erro foi identificado na rotina 'CarregaRsNosTMP()':" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo LogErroExportEPANET.txt com informações desta ocorrencia.", vbInformation
        If rsNos.EOF = False Then
            MsgBox "Problema com o nó de rede número: " & rsNos.Fields("OBJECT_ID_").Value
        End If
        'Resume
    End If
End Function

'Verifica se um trecho de rede já foi exportado, pois se já foi, não precisa exportar mais o mesmo
'
'id = número do object_id do trecho de rede
'TrechoCadastrado = True - já foi exportado para exportação para o EPANET
'TrechoCadastrado = False - ainda não foi exportado para a exportação para o EPANET
'
Public Function RsTrechosExportados_verificaSeFoiExportado(id As String) As Boolean
   rsTrechosExportados.Filter = "id='" & id & "'"   'aplica um filtro pelo object_id aos trechos já exportados
   TrechoExportado = Not rsTrechosExportados.EOF    'retorna verdadeiro se o trecho se o trecho ja foi exportado, falso se não
End Function

'Adiciona o Id do trecho de rede que já foi exportado para o vetor de trechos exportados
'Isso é feito para saber quais os trechos de rede que já forma exportados
'
'
'
Public Function RsTrechos_adicionaIdDoTrecho()
    RsTrechos.AddNew                                'Insere no vetor de trechos exportados o trecho processado
    RsTrechos.Fields("id").Value = idLinha
End Function

'Verifica se o nó da rede já foi exportado na lista de nós a serem exportados para o EPANET
'retorna true se foi exportado, false se não foi cadastrado
'
'
Public Function RsNosExportados_verificaSeNoFoiExportado(Object_id_no As String) As Boolean
   rsNosExportados.Filter = "id='" & Object_id_no & "'"
   VerificaSeNoFoiExportado = Not rsNosExportados.EOF
End Function

'Obtem o tipo de nó que estamos lendo (specification), para saber se é:
' JUNCTION, PUMP, RNF, REGISTER, VALVE ou RNV
'Isto é feito pois é necessário saber o que é o nó para dar o tratamento correto na exportação para o EPANET
'
'
'
'Public Function EspecificacaoDoTipoDeNo(Object_id_no As String) As Boolean
'    Dim strTipoComp As String                                                           'Variável que receberá o tipo de componente (Junction, valve,pump...)
'    rsNosTmp.Filter = "id='" & no & "'"                                                 'Seleciona no Cursor rsNosTmp o nó igual o valor de NO
'    If RsNosTmpTerminoDeLer = False Then                                                'Se não terminou de ler todos os nós
'        strTipoComp = ""
'        If blnRsWaterCompTypes = True Then
'            rsWaterCompTypes.MoveFirst
'            Do While Not rsWaterCompTypes.EOF = True
'                If rsNosTmp.Fields("Tipo").Value = rsWaterCompTypes!id_type Then
'                    strTipoComp = rsWaterCompTypes!SPECIFICATION_
'                    Exit Do
'                End If
'                rsWaterCompTypes.MoveNext
'            Loop
'        End If
'End Function

'Obtem o tipo de nó temporário do trecho de rede
'
'
Public Function RsNosTmp_obtemTipoDeNo() As Integer
    RsNosTmp_obtemTipoDeNo = rsNosTmp.Fields("Tipo").Value
End Function

'Verifica se terminou de ler todos os nós de rsNosTmp, que é um vetor temporário com todos
'os nós lidos de WaterComponents e Points2
'
'
'
Private Function RsNosTmpTerminoDeLer() As Boolean
    If rsNosTmp.EOF = True Then
        RsNosTmpTerminouDeLer = True
    Else
        RsNosTmpTerminouDeLer = False
    End If
End Function
'Filtra todos os nós existentes pelo nó selecionado
'
'
'
Public Function RsNosTmp_filtraPorNo(no As String)
     rsNosTmp.Filter = "id='" & no & "'"
End Function

'Verifica se no vetor criado temporáriamente rsNosTmp, se existem nós disponíveis
'isto geralmente é verificado após aplicar um filtro pelo nó que se está procurando
'em RsNosTmp_filtraPorNo
'
'
'
Public Function RsNosTmp_ExistemNos() As Boolean
    If rsNosTmp.EOF = False Then
        RsNosTmp_ExistemNos = True
    Else
        RsNosTmp_ExistemNos = False
    End If
End Function

'Inicializa o vetor com os componentes das redes, bombas, válvulas, etc.
'Por exemplo: curva 45, T, tipo reservatório, etc.
'
'
Public Function RsWaterCompTypes_Inicializa()
    Set rsWaterCompTypes = conn.Execute("SELECT * FROM WATERCOMPONENTSTYPES ORDER BY ID_TYPE")
End Function

'Obtem o tipo de componente
'
'
Private Function RsWaterCompTypes_obtemTipoDeComponente() As String
    RsWaterCompTypes_obtemTipoDeComponente = rsWaterCompTypes!id_type
End Function

'Obtem a especificação da componente
'
'
'
Private Function RsWaterCompTypes_obtemEspecificacao() As String
    RsWaterCompTypes_obtemEspecificacao = rsWaterCompTypes!SPECIFICATION_
End Function

'Move para a próxima componente da rede de água, primeiro nó do trecho de rede de água
'
'
'
Private Function RsWaterCompTypes_moveParaProximoTipoDeComponente()
    rsWaterCompTypes.MoveNext
End Function

'Verifica se já terminou de ler os dois nós da rede de água
'
'
'
Private Function rsWaterCompTypes_TerminouDeLerTiposDeComponentes() As Boolean
    If rsWaterCompTypes.EOF = True Then
        rsWaterCompTypes_TerminouDeLerTiposDeComponentes = True
    Else
        rsWaterCompTypes_TerminouDeLerTiposDeComponentes = False
    End If
End Function

'Verifica se existem parâmetros preenchidos das componentes como trechos, bombas, válvulas, etc.
'
'
'
Public Function RsWaterCompTypes_ExitemParametros() As Boolean
    If rsWaterCompTypes.EOF = False Then
        RsWaterCompTypes_ExitemParametros = True
    Else
        MsgBox "Não será possivel identificar e exportar bombas e válvulas pois a tabela WATERCOMPONENTSTYPES está vazia.", vbExclamation, ""
        RsWaterCompTypes_ExitemParametros = False
    End If
End Function

'Move para o primeiro tipo de componente da rede de água
'
'
'
Private Function rsWaterCompTypes_MoveParaPrimeiroTipoDeComponente()
    rsWaterCompTypes.MoveFirst
End Function

'Obtem a especificação do tipo de componente em questão
'
'
'
Public Function RsWaterCompTypes_tipoDeEspecificacao() As String
    Dim strTipoComp As String
    strTipoComp = ""
    If RsWaterCompTypes_ExitemParametros = True Then                                      'a tabela de componentes não está vazia sem nenhuma linha
        rsWaterCompTypes_MoveParaPrimeiroTipoDeComponente
        Do While Not rsWaterCompTypes_TerminouDeLerTiposDeComponentes()                   'encontra a especificação da componente - enquanto estiver lendo os tipos de componentes de rede existentes
            If RsNosTmp_obtemTipoDeNo = RsWaterCompTypes_obtemTipoDeComponente Then       'compara a componente do nó selecionado com a existente na tabela de tipos de componentes
                strTipoComp = RsWaterCompTypes_obtemEspecificacao                         'pega a especificação pois encontrou a mesma
                Exit Do                                                                   'já encontrou a especificação, sai do loop
            End If
            RsWaterCompTypes_moveParaProximoTipoDeComponente
        Loop
    End If
    RsWaterCompTypes_tipoDeEspecificacao = strTipoComp
End Function

'Adiciona os dados da junção para poder exportar depois para o EPANET
'
'
'
Public Function RsJunctions_adicionaJuncao(no As String, demanda As Double)
    rsJunctions.AddNew
    rsJunctions.Fields("id").Value = no
    rsJunctions.Fields("elev").Value = Format(rsNosTmp("cota").Value, ".0")
    rsJunctions.Fields("demand").Value = demanda
    rsJunctions.Fields("pattern").Value = IIf(rsNosTmp("padrao").Value = 0, "", rsNosTmp("padrao").Value)
End Function

'Adiciona os dados da bomba para poder exportar depois para o EPANET
'
'
'
Public Function RsPumps_adicionaBomba(no As String)
    rsPumps.AddNew
    rsPumps.Fields("id").Value = no
    If conta_no = 1 Then
       rsPumps.Fields("Node1").Value = no          'NESTE TRECHO É DEFINIDO O SENTIDO DA BOMBA
       rsPumps.Fields("Node2").Value = no & "A"    'A ORDEM DE NO E NOA INFLUENCIA NO SENTIDO
    Else
       rsPumps.Fields("Node1").Value = no & "A"
       rsPumps.Fields("Node2").Value = no
       rsPumps.Fields("potencia").Value = "10"
    End If
End Function

'Adiciona os dados da válvula para poder exportar depois para o EPANET
'
'
'
Public Function RsValves_adicionaValvula(no As String)
    rsValves.AddNew
    rsValves.Fields("ID").Value = "V" & no
    rsValves.Fields("Node1").Value = no & "A"
    rsValves.Fields("Node2").Value = no
    rsValves.Fields("Diameter").Value = Replace(rs_global.Fields("internaldiameter").Value, ",", ".") 'adiciona o diâmetro do trecho da tubulação a válvula
    'A próxima rotina foi retirada e substituída pela linha acima, uma vez que ela não estava operante. Não houve entendimento do por que ela foi criada 2013-04-27
    'AddSubItemValves NO 'Adiciona os sub itens para a valvula (setting,type,diameter)
End Function

'Adiciona os dados da valvula2 para poder exportar depois para o EPANET
'Veja que ele converte aqui um ponto em um trecho de rede
'
'
Public Function RsValves2_adicionaValvula(no As String)
    rsPipes.AddNew
    rsPipes.Fields("id").Value = no & "A"
    rsPipes.Fields("node1").Value = no & "A"
    rsPipes.Fields("node2").Value = no
    rsPipes.Fields("length").Value = 0.1
    rsPipes.Fields("diameter").Value = Replace(rs_global.Fields("internaldiameter").Value, ",", ".")
    rsPipes.Fields("roughness").Value = Replace(rs_global.Fields("roughness").Value, ",", ".")
    rsPipes.Fields("status").Value = IIf(rsNosTmp("estado").Value = 0, " ", rsNosTmp("estado").Value)
    If rs_global.Fields("MATERIALNAME").Value <> "" Then
       rsPipes.Fields("Description").Value = rs_global.Fields("MATERIALNAME").Value
    Else
       rsPipes.Fields("Description").Value = ""
    End If
End Function

'Adiciona os dados do registro para poder exportar depois para o EPANET
'Veja que ele converte aqui um ponto em um trecho de rede
'
'
Public Function RsPipes_adicionaRegistro(no As String)
    rsPipes.AddNew
    rsPipes.Fields("ID").Value = no & "A"
    rsPipes.Fields("NODE1").Value = no & "A"
    rsPipes.Fields("NODE2").Value = no
    rsPipes.Fields("LENGTH").Value = 1
    rsPipes.Fields("DIAMETER").Value = Replace(rs_global.Fields("internaldiameter").Value, ",", ".")
    rsPipes.Fields("ROUGHNESS").Value = Replace(rs_global.Fields("roughness").Value, ",", ".")
    rsPipes.Fields("STATUS").Value = IIf(rsNosTmp("ESTADO").Value = 0, " ", rsNosTmp("ESTADO").Value)
    rsPipes.Fields("DESCRIPTION").Value = "REGISTRO"
End Function

'Adiciona o trecho de rede em si
'
'
'
Public Function RsPipes_adicionaTrechoRede(idTrechoRede As String, idNoInicial As String, idNoFinal As String)
    rsPipes.AddNew
    rsPipes.Fields("id").Value = idTrechoRede
    rsPipes.Fields("node1").Value = idNoInicial
    rsPipes.Fields("node2").Value = idNoFinal
    If rs_global.Fields("Length").Value > 0 Then
        rsPipes.Fields("length").Value = Replace(rs_global.Fields("Length").Value, ",", ".")
    Else
        rsPipes.Fields("Length").Value = Replace(rs_global.Fields("LengthCalculated").Value, ",", ".")
    End If
    rsPipes.Fields("diameter").Value = Replace(rs_global.Fields("internaldiameter").Value, ",", ".")
    rsPipes.Fields("roughness").Value = Replace(rs_global.Fields("roughness").Value, ",", ".")
    'rsPipes.Fields("status").Value = IIf(rsNosTmp("estado").Value = 0, " ", rsNosTmp("estado").Value)
    rsPipes.Fields("status").Value = "Open"
    If rs_global.Fields("MATERIALNAME").Value <> "" Then
        rsPipes.Fields("Description").Value = rs_global.Fields("MATERIALNAME").Value
    Else
        rsPipes.Fields("Description").Value = ""
    End If
    'rsPipes.Fields("Description").Value = IIf(rs_global.Fields("MATERIAL").Value <> Null, "", rs_global.Fields("MATERIAL").Value)
End Function

'Adiciona os vértices dos trechos de rede
'
'
'
Public Function RsVertices_adicionaVerticesDosTrechosRede(idTrechoRede As String, coordXVertice As Double, coordYVertice As Double)
    rsVertices.AddNew
    rsVertices.Fields("ID").Value = idTrechoRede
    rsVertices.Fields("X-Coord").Value = coordXVertice
    rsVertices.Fields("Y-Coord").Value = coordYVertice
End Function


'Adiciona os dados do RNV (Reservatório de Nível Variável) para poder exportar para o EPANET
'
'
'
Public Function RsTanks_adicionaReservatorioNivelVariavel(no As String)
    rsTanks.AddNew
    rsTanks.Fields("ID").Value = no
    rsTanks.Fields("Elevation").Value = Format(rsNosTmp("cota").Value, ".0")
    AddSubItemTank no
End Function

'Adiciona os dados do reservatório para poder exportar para o EPANET
'
'
'
Public Function RsReservois_adicionaReservatorio(no As String)
    rsReservoirs.AddNew
    rsReservoirs.Fields("ID").Value = no
    rsReservoirs.Fields("Head").Value = Format(rsNosTmp("cota").Value, ".0")
    rsReservoirs.Fields("Pattern").Value = ""
End Function

'Adiciona as coordenadas geográficas de um nó, para o mesmo poder exportar para o EPANET
'
'
'
Public Function RsCoordinates_adicionaCoordenadas(no As String, x As Double, y As Double)
    rsCoordinates.AddNew
    rsCoordinates.Fields("id").Value = no
    rsCoordinates.Fields("x").Value = x
    rsCoordinates.Fields("y").Value = y
End Function

'Obtem a demanda do nó que está sendo lido
'
'
'
Public Function RsNosTmp_obtemDemandaDoNo()
    RsNosTmp_obtemDemandaDoNo = rsNosTmp("demanda").Value
End Function

'Obtem a coordenada x do nó temporário
'
'
'
Public Function RsNosTmp_coordX(no As String) As Double
    RsNosTmp_coordX = rsNosTmp("x").Value
End Function


'Obtem a coordenada y do nó temporário
'
'
'
Public Function RsNosTmp_coordY(no As String) As Double
    RsNosTmp_coordY = rsNosTmp("y").Value
End Function

'Adiciona informação do nó que foi exportado vetorialmente, preparando para o EPANET
'
'
'
Public Function RsNosExportados_insereNoExportado(no As String)
    rsNosExportados.AddNew
    rsNosExportados.Fields("id").Value = no
End Function

'Fecha o acesso aos vetores temporários
'
'
'
Public Function FechaAcessosAosVetores()
    Set rsNosTmp = Nothing
    
    RsTrechos.Close
    Set RsTrechos = Nothing
    
    rsTrechosExportados.Close
    Set rsTrechosExportados = Nothing
    
    Set rsNos = Nothing
End Function
