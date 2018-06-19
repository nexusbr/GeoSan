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
Private rsCurves As New ADODB.Recordset                             'parêmetros das curvas de bomba
Private rsValves As New ADODB.Recordset
Private rsReservoirs As New ADODB.Recordset                         'São os Reservatórios de Nível Fixo - RNF
Private rsTanks As New ADODB.Recordset                              'São os Reservatórios de Nível Variável - RNV
Private rsVertices As New ADODB.Recordset                           'Vertices da linha com exceção do inicial e final
Private rsNosTmp As New ADODB.Recordset                             'Criado para armazenar todos os dados de todos nos - Copia do Watercomponenstes/Points
Private rsTrechosExportados As New ADODB.Recordset                  'Criado para armazenar os trechos que já foram exportados
Private rsTrechos As New ADODB.Recordset
Private rsNosExportados As New ADODB.Recordset                      'Criado para armazenar os Nós que já foram exportados
Private rsWaterCompTypes As New ADODB.Recordset                     'RecordSet de referencia para consulta do tipo de componente

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

'Move para o próximo trecho de rede selecionado pela cerca no GeoSan
'
'
'
Public Function Rs_trechosDeRede_moveParaProximo()
    rs_trechosDeRede.MoveNext                                              'Move o ponteiro do cursor de trechos a serem processados para a próxima tupla
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
    rsPipes.Fields.Append "Description", adVarChar, 255
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
    rsPumps.Fields.Append "parameters", adVarChar, 255              'Aqui vai ou POWER ou HEAD e os demais parâmetros SPEED e/ou PATTERN. Uma das seguintes propriedades: POWER ou HEAD tem que ser fornecidas para caracterizar a bomba. Utilize a propriedade POWER se não for possível obter uma curva da bomba. As restantes palavras-chave são opcionais.
    rsPumps.Fields.Append "power", adVarChar, 255                   'POWER - valor da potência para uma energia de bombeamento constante, hp (kW)
    rsPumps.Fields.Append "carga", adVarChar, 255
    rsPumps.Fields.Append "vazao", adVarChar, 255
    rsPumps.Fields.Append "curvaRendimento1", adVarChar, 255
    rsPumps.Fields.Append "head", adVarChar, 255                    'HEAD - ID da curva que descreve a vazão vs altura de elevação para a bomba
    rsPumps.Fields.Append "curvaRendimento2", adVarChar, 255
    rsPumps.Fields.Append "speed", adVarChar, 255                   'SPEED - regulação de velocidade (regulação de velocidade usual é 1.0, 0 significa que a bomba está desligada)
    rsPumps.Fields.Append "pattern", adVarChar, 255
    rsPumps.Fields.Append "description", adVarChar, 255             'PATTERN - ID do padrão temporal que descreve a variação da propriedade Regulação de Velocidade (SPEED) com o tempo
    rsPumps.Open
    
    'válvulas
    'PRV (válvula redutora de pressão) Pressão, m (psi)
    'PSV (Válvula Sustentadora de Pressão) Pressão, m (psi)
    'PBV (Válvula de perda de carga fixa) Pressão, m (psi)
    'FCV (válvula reguladora de vazão) Vazão (unidades de vazão)
    'TCV (válvula de controle de perda de carga) Coeficiente de perda de carga singular
    'GPV (válvula genérica) ID da curva de perda de carga
    rsValves.Fields.Append "id", adVarChar, 255
    rsValves.Fields.Append "node1", adVarChar, 255
    rsValves.Fields.Append "node2", adVarChar, 255
    rsValves.Fields.Append "diameter", adVarChar, 255
    rsValves.Fields.Append "type", adVarChar, 255               'tipo de válvula - PRV, PSV, PBV, FCV, TCV, GPV
    rsValves.Fields.Append "setting", adVarChar, 255            'parâmetro de controle na válvula
    rsValves.Fields.Append "minorloss", adVarChar, 255          'Coeficiente de perda de carga singular
    rsValves.Open
    
    'Reservatório de Nível Fixo - RNF
    rsReservoirs.Fields.Append "ID", adVarChar, 255
    rsReservoirs.Fields.Append "Head", adVarChar, 255           'nível da água
    rsReservoirs.Fields.Append "Pattern", adVarChar, 255        'padrão de nível
    rsReservoirs.Fields.Append "Description", adVarChar, 255    'para armazenar o nome do reservatório
    rsReservoirs.Open

    'Reservatório de Nível Variável - RNV
    rsTanks.Fields.Append "ID", adVarChar, 255
    rsTanks.Fields.Append "Elevation", adVarChar, 255           'cota
    rsTanks.Fields.Append "InitLevel", adVarChar, 255           'altura inicial da água
    rsTanks.Fields.Append "MinLevel", adVarChar, 255            'altura mínima da água
    rsTanks.Fields.Append "MaxLevel", adVarChar, 255            'altura máxima da água
    rsTanks.Fields.Append "Diameter", adVarChar, 255            'diâmetro deo reservatório
    rsTanks.Fields.Append "MinVol", adVarChar, 255              'volume mínimo (não está sendo exportado, não utilizado)
    rsTanks.Fields.Append "VolCurve", adVarChar, 255            'curva de volume
    rsTanks.Fields.Append "Description", adVarChar, 255    'para armazenar o nome do reservatório
    rsTanks.Open
    
    'Curvas
    rsCurves.Fields.Append "ID", adVarChar, 255
    rsCurves.Fields.Append "X-Value", adVarChar, 255           'nível da água
    rsCurves.Fields.Append "Y-Value", adVarChar, 255        'padrão de nível
    rsCurves.Open
    
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
    
    rsTrechos.Fields.Append "id", adVarChar, 255
    rsTrechos.Open
    
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
    rsTrechos.AddNew                                'Insere no vetor de trechos exportados o trecho processado
    rsTrechos.Fields("id").Value = idLinha
End Function

'Verifica se o nó da rede já foi exportado na lista de nós a serem exportados para o EPANET
'retorna true se foi exportado, false se não foi cadastrado
'
'
Public Function RsNosExportados_esteNoJaFoiExportado(Object_id_no As String) As Boolean
    rsNosExportados.Filter = "id='" & Object_id_no & "'"
    If rsNosExportados.EOF = True Then
        RsNosExportados_esteNoJaFoiExportado = False
    Else
        RsNosExportados_esteNoJaFoiExportado = True
    End If
End Function

'Obtem o tipo de nó que estamos lendo (specification), para saber se é:
' JUNCTION, PUMP, RNF, REGISTER, VALVE ou RNV
'Isto é feito pois é necessário saber o que é o nó para dar o tratamento correto na exportação para o EPANET
'

'Obtem o tipo de nó temporário do trecho de rede
'
'
Public Function RsNosTmp_obtemTipoDeNo() As Integer
    On Error GoTo Trata_Erro
    RsNosTmp_obtemTipoDeNo = rsNosTmp.Fields("Tipo").Value

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        Close #1
        Close #2
        Open App.Path & "\LogErroExportEPANET.txt" For Append As #2
        
        Print #2, Now & "  - Function RsNosTmp_obtemTipoDeNo - nó com problema: " & rsNosTmp.Filter & " - " & Err.Number & " - " & Err.Description
        Close #2
        MsgBox "Existe um nó inicial ou final na rede 'Waterlines' que não existem em 'Watercomponents' - nó numero: " & rsNosTmp.Filter & " - " & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo LogErroExportEPANET.txt com informações desta ocorrencia.", vbInformation
    End If
    
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
    If RsWaterCompTypes_ExitemParametros = True Then     'a tabela de componentes não está vazia sem nenhuma linha
        rsWaterCompTypes_MoveParaPrimeiroTipoDeComponente
        Do While Not rsWaterCompTypes_TerminouDeLerTiposDeComponentes()                   'encontra a especificação da componente - enquanto estiver lendo os tipos de componentes de rede existentes
            If RsNosTmpTerminoDeLer = False Then                                          'não terminou de ler RsNosTmp
                If RsNosTmp_obtemTipoDeNo = RsWaterCompTypes_obtemTipoDeComponente Then   'compara a componente do nó selecionado com a existente na tabela de tipos de componentes
                    strTipoComp = RsWaterCompTypes_obtemEspecificacao                     'pega a especificação pois encontrou a mesma
                    Exit Do                                                               'já encontrou a especificação, sai do loop
                End If
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
Public Function RsPumps_adicionaBomba(no As String, conta_no As Integer)
    rsPumps.AddNew
    rsPumps.Fields("id").Value = "B" & no
    If conta_no = 1 Then
       rsPumps.Fields("Node1").Value = no          'NESTE TRECHO É DEFINIDO O SENTIDO DA BOMBA
       rsPumps.Fields("Node2").Value = no & "A"    'A ORDEM DE NO E NOA INFLUENCIA NO SENTIDO
    Else
       rsPumps.Fields("Node1").Value = no & "A"
       rsPumps.Fields("Node2").Value = no
    End If
    AddSubItemPumps (no)                            'Adiciona os sub itens para a bomba (curva)
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
    rsValves.Fields("Diameter").Value = Replace(rs_trechosDeRede.Fields("internaldiameter").Value, ",", ".")    'adiciona o diâmetro do trecho da tubulação a válvula
    AddSubItemValves no, rsValves.Fields("Diameter").Value                                                      'Adiciona os sub itens para a valvula (setting,type,diameter)
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
    rsPipes.Fields("diameter").Value = Replace(rs_trechosDeRede.Fields("internaldiameter").Value, ",", ".")
    rsPipes.Fields("roughness").Value = Replace(rs_trechosDeRede.Fields("roughness").Value, ",", ".")
    rsPipes.Fields("status").Value = IIf(rsNosTmp("estado").Value = 0, " ", rsNosTmp("estado").Value)
    If rs_trechosDeRede.Fields("MATERIALNAME").Value <> "" Then
       rsPipes.Fields("Description").Value = rs_trechosDeRede.Fields("MATERIALNAME").Value
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
    rsPipes.Fields("DIAMETER").Value = Replace(rs_trechosDeRede.Fields("internaldiameter").Value, ",", ".")
    rsPipes.Fields("ROUGHNESS").Value = Replace(rs_trechosDeRede.Fields("roughness").Value, ",", ".")
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
    If rs_trechosDeRede.Fields("Length").Value > 0 Then
        rsPipes.Fields("length").Value = Replace(rs_trechosDeRede.Fields("Length").Value, ",", ".")
    Else
        rsPipes.Fields("Length").Value = Replace(rs_trechosDeRede.Fields("LengthCalculated").Value, ",", ".")
    End If
    rsPipes.Fields("diameter").Value = Replace(rs_trechosDeRede.Fields("internaldiameter").Value, ",", ".")
    rsPipes.Fields("roughness").Value = Replace(rs_trechosDeRede.Fields("roughness").Value, ",", ".")
    'rsPipes.Fields("status").Value = IIf(rsNosTmp("estado").Value = 0, " ", rsNosTmp("estado").Value)
    rsPipes.Fields("status").Value = "Open"
    If rs_trechosDeRede.Fields("MATERIALNAME").Value <> "" Then
        rsPipes.Fields("Description").Value = rs_trechosDeRede.Fields("MATERIALNAME").Value
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
    rsTanks.Fields("Elevation").Value = FormataStringNumerica(rsNosTmp("cota").Value, "0.0")
    AdicionaSubItensReservatorioNivelVariavel no
End Function

'Adiciona os dados do reservatório para poder exportar para o EPANET
'
'
'
Public Function RsReservois_adicionaReservatorioNivelFixo(no As String)
    rsReservoirs.AddNew
    rsReservoirs.Fields("ID").Value = no
    AdicionaSubItensReservatorioNivelFixo no
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
    
    rsTrechos.Close
    Set rsTrechos = Nothing
    
    rsTrechosExportados.Close
    Set rsTrechosExportados = Nothing
    
    Set rsNos = Nothing
End Function

'Gera automaticamente o cabeçalho da componente ou trecho a ser exportado para o EPANET
'
'
'
Private Function GeraCabecalhoComPosicaoColunaDescricao(rsGenerico As ADODB.Recordset, stringCabecalho As String) As Integer
    Dim posicaoDaDescricao As Integer
    
    posicaoDaDescricao = 100                                                                                'coloca um valor alto para informar que não foi localizada a descrição
    rsGenerico.MoveFirst
    stringCabecalho = ""
    For A = 0 To rsGenerico.Fields.Count - 1
        If rsGenerico.Fields(A).Name = "Description" Then
            posicaoDaDescricao = A                                                                          'lê a posição onde está a descrição da bomba para colocá-la depois do ; no final
            stringCabecalho = stringCabecalho & rsGenerico.Fields(A).Name & Chr(vbKeyTab) & Chr(vbKeyTab)
        Else
            stringCabecalho = stringCabecalho & rsGenerico.Fields(A).Name & Chr(vbKeyTab) & Chr(vbKeyTab)
        End If
    Next
    GeraCabecalhoComPosicaoColunaDescricao = posicaoDaDescricao
End Function

'Imprime o cabeçalho da componente ou trecho a ser exportado, no arquivo de exportação do EPANET
'
'
'
Private Function ImprimeCabecalho(titulo As String, cabacalho As String)
    Open FrmEPANET.txtArquivo.Text For Append As #1
    Print #1, ""
    Print #1, titulo
    Print #1, ";" & cabacalho
    Close #1
End Function

'Imprime os dados da componente ou trecho a ser exportado, no arquivo de exportação do EPANET
'
'
'
Private Function GravaDadosDoNoOuTrecho(rsGenerico As ADODB.Recordset, posicaoDaDescricao As Integer)
    Dim str As String
    
    str = ""
    Open FrmEPANET.txtArquivo.Text For Append As #1
    While Not rsGenerico.EOF
        For A = 0 To rsGenerico.Fields.Count - 1
            If A <> posicaoDaDescricao Then                                         'pula a descrição para colocá-la somente depois do ;
                str = str & IIf(IsNumeric(rsGenerico.Fields(A).Value), Replace(rsGenerico.Fields(A).Value, ",", "."), rsGenerico.Fields(A).Value) & Chr(vbKeyTab) & Chr(vbKeyTab)
            End If
        Next
        If posicaoDaDescricao <> 100 Then                                            'existe uma descrição válida
            Print #1, str & ";" & rsGenerico.Fields(posicaoDaDescricao).Value       'coloca a string com os dados da bomba, seguidos da descrição da mesma
        Else
            Print #1, str & ";"
        End If
        str = ""
        rsGenerico.MoveNext
    Wend
    Close #1
End Function

'Imprime os dados da bomba a ser exportado, no arquivo de exportação do EPANET
'
'
'
Private Function GravaDadosDaBomba(rsGenerico As ADODB.Recordset)
    Dim str As String
    
    str = ""
    Open FrmEPANET.txtArquivo.Text For Append As #1
    While Not rsGenerico.EOF
        str = rsGenerico.Fields("id").Value & Chr(vbKeyTab) & Chr(vbKeyTab)
        str = str + rsGenerico.Fields("node1").Value & Chr(vbKeyTab) & Chr(vbKeyTab)
        str = str + rsGenerico.Fields("node2").Value & Chr(vbKeyTab) & Chr(vbKeyTab)
        str = str + rsGenerico.Fields("parameters").Value & Chr(vbKeyTab) & Chr(vbKeyTab)
        str = str + ";" + rsGenerico.Fields("description").Value & Chr(vbKeyTab) & Chr(vbKeyTab)
        Print #1, str
        str = ""
        rsGenerico.MoveNext
    Wend
    Close #1
End Function

'Gera o arquivo de saída .INP que o Epanet lê
'Recupera os dados do cursor em memoria e cria o arquivo .INP
'#########################################################################
'SEQUENCIA DA ESTRUTURA QUE DEVE SER GRAVADA PARA O ARQUIVO .INP DO EPANET
'#########################################################################
   ' "TITLE"
   ' "JUNCTIONS"
   ' "RESERVOIRS"
   ' "TANKS"
   ' "PUMPS"
   ' "VALVES"
   ' "PIPES"
   ' "TAGS"
   ' "DEMANDS"
   ' "STATUS"
   ' "PATTERNS"
   ' "CURVES"
   ' "CONTROLS"
   ' "RULES"
   ' "ENERGY"
   ' "EMITTERS"
   ' "SOURCES"
   ' "REACTIONS"
   ' "MIXING"
   ' "REACTIONS"
   ' "TIMES"
   ' "REPORT"
   ' "OPTIONS"
   ' "COORDINATES"
   ' "VERTICES"
   ' "BACKDROP"
   ' "END"
'#########################################################################
'
Public Function GeraArquivo_de_Saida()
    On Error GoTo Trata_Erro
    Dim A As Long
    Dim str As String
    Dim posicaoDaDescricao As Integer                                                   'posição no vetor, da coluna descrição

    Open FrmEPANET.txtArquivo.Text For Output As #1
    Close #1

    rsJunctions.Filter = ""                                                             'grava no arquivo as Junctions
    If rsJunctions.RecordCount > 0 Then
        posicaoDaDescricao = GeraCabecalhoComPosicaoColunaDescricao(rsJunctions, str)
        Call ImprimeCabecalho("[JUNCTIONS]", str)
        Call GravaDadosDoNoOuTrecho(rsJunctions, posicaoDaDescricao)
    End If

    rsReservoirs.Filter = ""                                                            'grava no arquivo as Reservoirs - RNF
    If rsReservoirs.RecordCount > 0 Then
        posicaoDaDescricao = GeraCabecalhoComPosicaoColunaDescricao(rsReservoirs, str)
        Call ImprimeCabecalho("[RESERVOIRS]", str)
        Call GravaDadosDoNoOuTrecho(rsReservoirs, posicaoDaDescricao)
    End If
    
    rsTanks.Filter = ""                                                                 'grava no arquivo as Tanks
    If rsTanks.RecordCount > 0 Then
        posicaoDaDescricao = GeraCabecalhoComPosicaoColunaDescricao(rsTanks, str)
        Call ImprimeCabecalho("[TANKS]", str)
        Call GravaDadosDoNoOuTrecho(rsTanks, posicaoDaDescricao)
    End If
    
    rsPumps.Filter = ""                                                                 'grava no arquivo as Pumps
    If rsPumps.RecordCount > 0 Then
        str = "ID                  Node1               Node2               Parameters"
        Call ImprimeCabecalho("[PUMPS]", str)
        Call GravaDadosDaBomba(rsPumps)
    End If

    rsValves.Filter = ""                                                                'grava no arquivo as Valves
    If rsValves.RecordCount > 0 Then
        posicaoDaDescricao = GeraCabecalhoComPosicaoColunaDescricao(rsValves, str)
        Call ImprimeCabecalho("[VALVES]", str)
        Call GravaDadosDoNoOuTrecho(rsValves, posicaoDaDescricao)
    End If

    With rsValves                                                                       'grava no arquivo os status das válvulas, esta gravação é realizada para facilitar a simulação hidráulica no Epanet, evitando que o usuário tenha que abrir manualmente os registros para iniciar uma simulação completa. Depois ele fecha os registros que achar necessário
        .Filter = ""
        If .RecordCount > 0 Then
            .MoveFirst                                                                  'pula o primeiro registro que é o cabeçalho
            str = "ID                 Status/Setting"
            Call ImprimeCabecalho("[STATUS]", str)
            Open FrmEPANET.txtArquivo.Text For Append As #1
            While Not .EOF
                str = .Fields(0).Value & Chr(vbKeyTab) & Chr(vbKeyTab) & "Open"         'prepara para impressão somente o número da válvula e o status da mesma que está aberta
                Print #1, str & ";"
                str = ""
                .MoveNext                                                               'vai para a próxima válvula
            Wend
            Close #1
        End If
    End With

    rsPipes.Filter = ""                                                                 'exporta tubulações
    If rsPipes.RecordCount > 0 Then
        posicaoDaDescricao = GeraCabecalhoComPosicaoColunaDescricao(rsPipes, str)
        Call ImprimeCabecalho("[PIPES]", str)
        Call GravaDadosDoNoOuTrecho(rsPipes, posicaoDaDescricao)
    End If

    'X_PATTERNS
    Dim MyArray() As String
    Dim rsPatterns As ADODB.Recordset
    If conn.Provider <> "PostgreSQL.1" Then
        Set rsPatterns = conn.Execute("Select * from x_patterns")
    Else
        Set rsPatterns = conn.Execute("Select * from " + """" + "X_PATTERNS" + """" + "")
    End If
    
    If rsPatterns.EOF = False Then                                                      'grava no arquivo as Patterns
        With rsPatterns
            str = "ID       Multipliers"
            Call ImprimeCabecalho("[PATTERNS]", str)
'            Print #1, "[PATTERNS]"
'            Print #1, ";ID" & Chr(vbKeyTab) & Chr(vbKeyTab) & "Multipliers"
            Open FrmEPANET.txtArquivo.Text For Append As #1
            Print #1, ";" & rsPatterns("descricao").Value
            Close #1
            Open FrmEPANET.txtArquivo.Text For Append As #1
            While Not .EOF
                MyArray = Split(rsPatterns("Padrao").Value, ";", 25)
                Print #1, rsPatterns("ID").Value & Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(0), ",", ".") & _
                    Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(1), ",", ".") & _
                    Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(2), ",", ".") & _
                    Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(3), ",", ".") & _
                    Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(4), ",", ".") & _
                    Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(5), ",", ".")
                If MyArray(6) <> "" Then
                    Print #1, rsPatterns("ID").Value & Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(6), ",", ".") & _
                        Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(7), ",", ".") & _
                        Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(8), ",", ".") & _
                        Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(9), ",", ".") & _
                        Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(10), ",", ".") & _
                        Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(11), ",", ".")
                    If MyArray(12) <> "" Then
                        Print #1, rsPatterns("ID").Value & Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(12), ",", ".") & _
                            Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(13), ",", ".") & _
                            Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(14), ",", ".") & _
                            Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(15), ",", ".") & _
                            Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(16), ",", ".") & _
                            Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(17), ",", ".")
                        If MyArray(18) <> "" Then
                            Print #1, rsPatterns("ID").Value & Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(18), ",", ".") & _
                                Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(19), ",", ".") & _
                                Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(20), ",", ".") & _
                                Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(21), ",", ".") & _
                                Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(22), ",", ".") & _
                                Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray(23), ",", ".")
                     End If
                  End If
                End If
                rsPatterns.MoveNext
            Wend
            Close #1
        End With
    End If
    rsPatterns.Close
    Set rsPatterns = Nothing
    
    rsCurves.Filter = ""                                                               'grava no arquivo as Curves
    If rsCurves.RecordCount > 0 Then
        posicaoDaDescricao = GeraCabecalhoComPosicaoColunaDescricao(rsCurves, str)
        Call ImprimeCabecalho("[CURVES]", str)
        Call GravaDadosDoNoOuTrecho(rsCurves, posicaoDaDescricao)
    End If

    rsCoordinates.Filter = ""                                                               'grava no arquivo as Coordinates
    If rsCoordinates.RecordCount > 0 Then
        posicaoDaDescricao = GeraCabecalhoComPosicaoColunaDescricao(rsCoordinates, str)
        Call ImprimeCabecalho("[COORDINATES]", str)
        Call GravaDadosDoNoOuTrecho(rsCoordinates, posicaoDaDescricao)
    End If

    rsVertices.Filter = ""                                                                  'grava no arquivo as Vertices
    If rsVertices.RecordCount > 0 Then
        posicaoDaDescricao = GeraCabecalhoComPosicaoColunaDescricao(rsVertices, str)
        Call ImprimeCabecalho("[VERTICES]", str)
        Call GravaDadosDoNoOuTrecho(rsVertices, posicaoDaDescricao)
    End If

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        Close #1
        Close #2
        Open App.Path & "\LogErroExportEPANET.txt" For Append As #2
        Print #2, Now & "  - Sub GeraArquivo_de_Saida() - Linha: " & intLinhaCod & " - " & Err.Number & " - " & Err.Description
        Close #2
        MsgBox "Um posssível erro foi identificado na rotina 'GeraArquivo_de_Saida()':" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo LogErroExportEPANET.txt com informações desta ocorrencia.", vbInformation
    End If
End Function

'Lê a tabela com as curvas cadastradas pelo usuário e exporta as mesmas para o EPANET
'esta rotina não foi implementada e testada ainda
'
'
'
Private Function LeCurvasEGravaNoArquivoParaExportacaoEPANET()
Dim b As Integer
    Dim MyArray_x() As String
    Dim MyArray_y() As String
    Dim rsCurves As ADODB.Recordset
    Dim str As String
    
    Set rsCurves = conn.Execute("Select * from x_Curves order by tipo")

    rsCurves.Filter = ""
    If rsCurves.EOF = False Then
        With rsCurves
            str = "ID       X-Value     Y-Value"
            Call ImprimeCabecalho("[CURVES]", str)
            Open FrmEPANET.txtArquivo.Text For Append As #1
            For b = 1 To 4
                If b = 1 Then
                    rsCurves.Filter = "Tipo = 'Bomba'"
                    If Not rsCurves.EOF Then Print #1, ";PUMPS:" & rsCurves.Fields("descricao").Value
                ElseIf b = 2 Then
                    rsCurves.Filter = "Tipo = 'Rendimento'"
                    If Not rsCurves.EOF Then Print #1, ";EFFICIENCY:" & rsCurves.Fields("descicao").Value
                ElseIf b = 3 Then
                    rsCurves.Filter = "Tipo = 'Volume'"
                    If Not rsCurves.EOF Then Print #1, ";VOLUME:" & rsCurves.Fields("descicao").Value
                ElseIf b = 4 Then
                    rsCurves.Filter = "Tipo = 'Perda de Carga'"
                    If Not rsCurves.EOF Then Print #1, ";HEADLOSS:" & rsCurves.Fields("descicao").Value
                End If
                intLinhaCod = 33
                While Not .EOF
                    MyArray_x = Split(rsCurves("Coordenada_x").Value, ";", 50)
                    MyArray_y = Split(rsCurves("Coordenada_y").Value, ";", 50)
                    For A = 0 To 49
                        If MyArray_x(A) = "" Then
                            Exit For
                        Else
                            Print #1, .Fields("ID").Value & Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray_x(A), ",", ".") & Chr(vbKeyTab) & Chr(vbKeyTab) & Replace(MyArray_y(A), ",", ".")
                        End If
                    Next
                    rsCurves.MoveNext
                Wend
            Next
            Close #1
        End With
    End If
    rsCurves.Close
    Set rsCurves = Nothing
End Function


'Adiciona os parâmetros entrados em um reservatório de nível variável RNV
'
'
'
Private Sub AdicionaSubItensReservatorioNivelVariavel(id As String)
   Dim Rs As ADODB.Recordset
   Dim CURVE As String
   Dim numeroCurvaBomba As String

   Set Rs = conn.Execute("Select b.eparef, w.value_, b.Description_ from watercomponentsdata w " & _
                         "INNER JOIN watercomponentssubtypes b on b.id_subtype=w.id_subtype and b.id_type=w.id_type " & _
                         "LEFT JOIN WaterComponentsSelections s on s.id_subtype=w.id_subtype and s.id_type=w.id_type and w.value_=s.value_ " & _
                         "where object_id_ = '" & id & "'")
   
   rsTanks.Fields("MaxLevel").Value = FormataStringNumerica("0", "0.0")                                                         'inicializa com zero ou nulo, caso não existem valores preenchidos
   rsTanks.Fields("MinLevel").Value = FormataStringNumerica("0", "0.0")
   rsTanks.Fields("InitLevel").Value = FormataStringNumerica("0", "0.0")
   rsTanks.Fields("Diameter").Value = FormataStringNumerica("0", "0.0")
   rsTanks.Fields("MinVol").Value = FormataStringNumerica("0", "0.0")
   rsTanks.Fields("VolCurve").Value = " "
   rsTanks.Fields("Description").Value = " "
   
   While Not Rs.EOF
      Select Case Rs.Fields("EPAREF").Value
      
         Case "NMAXIMO"
            rsTanks.Fields("MaxLevel").Value = FormataStringNumerica(Rs.Fields("VALUE_").Value, "0.0")
         Case "NMINIMO"
            rsTanks.Fields("MinLevel").Value = FormataStringNumerica(Rs.Fields("VALUE_").Value, "0.0")
         Case "NINICIAL"
            rsTanks.Fields("InitLevel").Value = FormataStringNumerica(Rs.Fields("VALUE_").Value, "0.0")
         Case "DIAMETER"
            rsTanks.Fields("Diameter").Value = FormataStringNumerica(Rs.Fields("VALUE_").Value, "0.0")
         Case "VOLUME"
            rsTanks.Fields("MinVol").Value = FormataStringNumerica(Rs.Fields("VALUE_").Value, "0.0")
         Case "VOLCURVE"
            numeroCurvaBomba = FormataStringNumerica(Rs.Fields("VALUE_").Value, "")
            rsTanks.Fields("VolCurve").Value = IIf(numeroCurvaBomba = "0", " ", "C" & numeroCurvaBomba)                         'pois se for zero tem que deixar em branco, senão pensa que é a curva zero, C0
         Case "DESC"
            rsTanks.Fields("Description").Value = Rs.Fields("VALUE_").Value
      End Select
      Rs.MoveNext
   Wend
   Rs.Close
   Set Rs = Nothing
End Sub

'Adiciona os parâmetros entrados em um reservatório de nível fixo RNF
'
'
'
Private Sub AdicionaSubItensReservatorioNivelFixo(id As String)
    Dim Rs As ADODB.Recordset
    Dim cota As Double
    Dim parametrosPreenchidos As String

   Set Rs = conn.Execute("Select b.eparef, w.value_, b.Description_ from watercomponentsdata w " & _
                         "INNER JOIN watercomponentssubtypes b on b.id_subtype=w.id_subtype and b.id_type=w.id_type " & _
                         "LEFT JOIN WaterComponentsSelections s on s.id_subtype=w.id_subtype and s.id_type=w.id_type and w.value_=s.value_ " & _
                         "where object_id_ = '" & id & "'")
   
   rsReservoirs.Fields("Head").Value = FormataStringNumerica(rsNosTmp("cota").Value, "")          'inicializa com a cota caso não existem valores preenchidos em head
   rsReservoirs.Fields("Pattern").Value = " "
   rsReservoirs.Fields("Description").Value = " "
   
   While Not Rs.EOF
      Select Case Rs.Fields("EPAREF").Value
         Case "HEAD"
            cota = CDbl(rsNosTmp("cota").Value) * CDbl(Rs.Fields("VALUE_").Value)
            rsReservoirs.Fields("Head").Value = FormataStringNumerica(CStr(cota), "0.0")
         Case "PATTERN"
            rsReservoirs.Fields("Pattern").Value = Rs.Fields("VALUE_").Value
         Case "DESC"
            rsReservoirs.Fields("Description").Value = Rs.Fields("VALUE_").Value
      End Select
      Rs.MoveNext
   Wend
   Rs.Close
   Set Rs = Nothing
End Sub

'Atribui o valor de curva para bomba
'
'
Public Function AddSubItemPumps(id As String)
   Dim Rs As ADODB.Recordset
   Dim curvaAltura As Double                            'altura (carga) para gerar a curva automaticamente, caso o usuário possua somente altura e vazão para a bomba
   Dim curvaVazao As Double                             'vazão para gerar a curva automaticamente, caso o usuário possua somente altura e vazão para a bomba
   Dim curva As String                                  'nome da curva da bomba
   Dim speed As String                                  'regulação de velocidade (regulação de velocidade usual é 1.0, 0 significa que a bomba está desligada)
   Dim pattern As String                                'ID do padrão temporal que descreve a variação da propriedade Regulação de Velocidade (SPEED) com o tempo
   Dim power As String                                  'valor da potência para uma energia de bombeamento constante, hp (kW)
   Dim head As String                                   'ID da curva que descreve a vazão vs altura de elevação para a bomba
   
   Set Rs = conn.Execute("Select b.eparef, w.value_, b.description_ from watercomponentsdata w " & _
                         "INNER JOIN watercomponentssubtypes b on b.id_subtype=w.id_subtype and b.id_type=w.id_type " & _
                         "LEFT JOIN WaterComponentsSelections s on s.id_subtype=w.id_subtype and s.id_type=w.id_type and w.value_=s.value_ " & _
                         "where object_id_ = '" & id & "'" & _
                         "order by b.eparef")
                                        
    rsPumps.Fields("parameters").Value = " "
    rsPumps.Fields("power").Value = FormataStringNumerica("9999", "")                                   'inicializa com zero ou nulo, caso não existem valores preenchidos
    rsPumps.Fields("carga").Value = FormataStringNumerica("9999", "")
    rsPumps.Fields("vazao").Value = FormataStringNumerica("9999", "")
    rsPumps.Fields("speed").Value = FormataStringNumerica("9999", "")
    rsPumps.Fields("pattern").Value = ""
    rsPumps.Fields("head").Value = ""
    rsPumps.Fields("description").Value = " "
            
    parametrosPreenchidos = ""
    If AddSubItemPumps_possuiParametroPreenchido("POWER", id) = True Then
        parametrosPreenchidos = "POWER"
    ElseIf AddSubItemPumps_possuiParametroPreenchido("CARGA", id) = True And AddSubItemPumps_possuiParametroPreenchido("VAZAO", id) = True Then
        parametrosPreenchidos = "CARGA"
    ElseIf AddSubItemPumps_possuiParametroPreenchido("CURBOMBA", id) = True Then
        parametrosPreenchidos = "CURBOMBA"
    End If
    
    If parametrosPreenchidos <> "" Then
        Select Case parametrosPreenchidos
            Case "POWER"                                                                             'Usuário só tem a potência da bomba
                Do While Not Rs.EOF
                    Select Case Rs.Fields("EPAREF").Value
                        Case "POWER"
                            power = IIf(Rs.Fields("VALUE_").Value = "0", " ", " POWER " & Rs.Fields("VALUE_").Value)
                        Case "SPEED"
                            speed = IIf(Rs.Fields("VALUE_").Value = "0", " ", " SPEED " & Rs.Fields("VALUE_").Value)
                        Case "PATTERN"
                            pattern = IIf(Rs.Fields("VALUE_").Value = "0", " ", " PATTERN " & Rs.Fields("VALUE_").Value)
                        Case "DESC"
                          rsPumps.Fields("description").Value = Rs.Fields("VALUE_").Value
                    End Select
                    Rs.MoveNext
                Loop
                rsPumps.Fields("parameters").Value = power & speed & pattern
            Case "CARGA"                                                                          'Usuário só tem a carga (altura) e vazão da bomba, podendo ter a curva de rendimento
                Do While Not Rs.EOF
                    Select Case Rs.Fields("EPAREF").Value
                        Case "CARGA"
                            curvaAltura = Rs.Fields("VALUE_").Value
                        Case "VAZAO"
                            curvaVazao = Rs.Fields("VALUE_").Value
                        Case "SPEED"
                            speed = IIf(Rs.Fields("VALUE_").Value = "0", " ", " SPEED " & Rs.Fields("VALUE_").Value)
                        Case "PATTERN"
                            pattern = IIf(Rs.Fields("VALUE_").Value = "0", " ", " PATTERN " & Rs.Fields("VALUE_").Value)
                        Case "DESC"
                          rsPumps.Fields("description").Value = Rs.Fields("VALUE_").Value
                    End Select
                    Rs.MoveNext
                Loop
                If curvaAltura = 0 Or curvaVazao = 0 Then                                           'os parâmetros não foram preenchidos pelo usuário, quando deveriam ser os dois preenchidos para poder gerar a curva, neste caso não exportará a curva da bomba
                    rsPumps.Fields("parameters").Value = " "
                Else                                                                                'foram preenchidos
                    curva = "C" & Replace(str(id), " ", "")
                    AddCurve curva, curvaAltura, curvaVazao
                    rsPumps.Fields("parameters").Value = " HEAD " & curva & speed & pattern
                End If
            Case "CURBOMBA"                                                                         'Usuário tem a curva da bomba, podendo ter a curva de rendimento da bomba também
                Do While Not Rs.EOF
                    Select Case Rs.Fields("EPAREF").Value
                        Case "CURBOMBA"
                           head = IIf(Rs.Fields("VALUE_").Value = "0", " ", " HEAD " & FormataStringNumerica(Rs.Fields("VALUE_").Value, ""))
                        Case "SPEED"
                            speed = IIf(Rs.Fields("VALUE_").Value = "0", " ", " SPEED " & Rs.Fields("VALUE_").Value)
                        Case "PATTERN"
                            pattern = IIf(Rs.Fields("VALUE_").Value = "0", " ", " PATTERN " & Rs.Fields("VALUE_").Value)
                        Case "DESC"
                          rsPumps.Fields("description").Value = Rs.Fields("VALUE_").Value
                    End Select
                    Rs.MoveNext
                Loop
                rsPumps.Fields("parameters").Value = head & speed & pattern
        End Select
    End If
    Rs.Close
    Set Rs = Nothing
End Function

'Verifica quais os parâmetros que foram preenchidos para a bomba
'Opção 1 - Potência (POWER)
'Opção 2 - Altura e Vazão (CARGA, VAZAO, RENDIM)
'Opção 3 - Curva da Bolba (CURBOMBA, CURREND)
'Parâmetro pode receber: POWER, CARGA ou CURBOMBA, correspontentes as opções 1, 2 e 3 respectivamente
'
'
Private Function AddSubItemPumps_possuiParametroPreenchido(parametro As String, id As String) As String
    Dim Rs As ADODB.Recordset

    Set Rs = conn.Execute("Select b.eparef, w.value_, b.description_ from watercomponentsdata w " & _
                         "INNER JOIN watercomponentssubtypes b on b.id_subtype=w.id_subtype and b.id_type=w.id_type " & _
                         "LEFT JOIN WaterComponentsSelections s on s.id_subtype=w.id_subtype and s.id_type=w.id_type and w.value_=s.value_ " & _
                         "where object_id_ = '" & id & "'" & _
                         "order by b.eparef")

    Do While Not Rs.EOF
        Select Case Rs.Fields("EPAREF").Value
            Case parametro
                AddSubItemPumps_possuiParametroPreenchido = True
                Exit Function
        End Select
        Rs.MoveNext
    Loop
    AddSubItemPumps_possuiParametroPreenchido = False
End Function

Private Function AddCurve(id As String, x As Double, y As Double)
    rsCurves.AddNew
    rsCurves.Fields("ID").Value = "C" & id
    rsCurves.Fields("X-Value").Value = x
    rsCurves.Fields("Y-Value").Value = y
End Function

'Adiciona as outras componentes de uma válvula, como diametro, tipo, etc.
'
' id - número do nó
'
' Esta rotina não está funcionando a querie abaixo está com colunas não existentes e pelo
' que foi verificado não retorna linhas. Em nosso entendimento ela não tem função 2013-04-27
'
'
Public Function AddSubItemValves(id As String, diametro As String)
'Atribui os valores especificos para valvula
    Dim Rs As ADODB.Recordset
    Dim PumpDiameter As Double, PumpType As String, PumpSetting As String, PumpMinorLoss As String
    
    Set Rs = conn.Execute("Select b.eparef, w.value_, b.description_ from watercomponentsdata w " & _
        "INNER JOIN watercomponentssubtypes b on b.id_subtype=w.id_subtype and b.id_type=w.id_type " & _
        "LEFT JOIN WaterComponentsSelections s on s.id_subtype=w.id_subtype and s.id_type=w.id_type and w.value_=s.value_ " & _
        "where object_id_ = '" & id & "'")
    While Not Rs.EOF
        Select Case Rs.Fields("EPAREF").Value
            Case "TYPE"                                                             'tipo de válvula - PRV, PSV, PBV, FCV, TCV, GPV
                rsValves.Fields("Type").Value = ObtemTipoDeValvula(Rs.Fields("VALUE_").Value)
            Case "SETTING"                                                          'parametro de controle
                rsValves.Fields("Setting").Value = Rs.Fields("VALUE_").Value
            Case "DIAMETER"                                                         'diametro
                If CDbl(Rs.Fields("VALUE_").Value) <> 0 Then                              'caso o diâmetro da válvula não tenha sido preenchido, então coloque o diâmetro da tubulação
                    rsValves.Fields("Diameter").Value = Rs.Fields("VALUE_").Value
                Else
                    rsValves.Fields("Diameter").Value = diametro
                End If
            Case "MINORLOSS"                                                        'coeficiente de perda de carga na válvula
                rsValves.Fields("MinorLoss").Value = Rs.Fields("VALUE_").Value
        End Select
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
End Function

'Obtem o tipo de válvula a partir do número da mesma
'
'
'
Private Function ObtemTipoDeValvula(tipo As String) As String
    Select Case tipo
        Case "1"
            ObtemTipoDeValvula = "PRV"      'PRV (válvula redutora de pressão) Pressão, m (psi)
        Case "2"
            ObtemTipoDeValvula = "PSV"      'PSV (Válvula Sustentadora de Pressão) Pressão, m (psi)
        Case "3"
            ObtemTipoDeValvula = "PBV"      'PBV (Válvula de perda de carga fixa) Pressão, m (psi)
        Case "4"
            ObtemTipoDeValvula = "FCV"      'FCV (válvula reguladora de vazão) Vazão (unidades de vazão)
        Case "5"
            ObtemTipoDeValvula = "TCV"      'TCV (válvula de controle de perda de carga) Coeficiente de perda de carga singular
        Case "6"
            ObtemTipoDeValvula = "GPV"      'GPV (válvula genérica) ID da curva de perda de carga
    End Select
End Function


'Formata uma string numérica para poder ser salva no banco de dados
' "0" -> "0.0"
' "645,84" -> "645.8"
'
'
'
Private Function FormataStringNumerica(valor As String, formato As String) As String
   Dim val As Double
   Dim valStr As String
   
   val = CDbl(Replace(valor, ".", ","))
   valStr = CStr(val)
   valStr = Format(valStr, formato)
   FormataStringNumerica = Replace(valStr, ",", ".")
End Function
