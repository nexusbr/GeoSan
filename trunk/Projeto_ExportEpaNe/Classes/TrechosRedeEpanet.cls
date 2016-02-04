VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TrechosRedeEpanet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Esta classe é referente ao processamento dos trechos de redes que foram selecionados para exportação para o EPANET
'
'
Option Explicit
Private conBd As ADODB.Connection
Private trechos As New TrechosRedeEpanet
Private usuario_global As String
Private conn_global As ADODB.Connection
Private az As Integer
Private rsWaterCompTypes As New ADODB.Recordset         'RecordSet de referencia para consulta do tipo de componente
Private blnRsWaterCompTypes As Boolean                  'Indicador para informar se a tabela RsWaterCompTypes foi carregada com registros
Private strTipoComp As String                           'Variável que receberá o tipo de componente (Junction, valve,pump...)
Private rsNos As ADODB.Recordset                        'Cursor Temporário que retonará todos os nos da Rede para abastecer o RsNosTMP
Private rsNos2 As Recordset
Private tb As New TeDatabase

'Realiza toda a exportação para o EPANET
'
'
'Procedimento Exporte EPANET recebe como parametros o cursor trazendo todas os trechos
'a serem exportados e o objecto de conexão
'(rsTrechos):É a tabela Waterlines com os filtros de tipo de rede e setor selecionados pelo usuário (TIPO=1 na tabela POLIGONO_SELECAO)
'arquivoLog: nome do arquivo onde está sendo escrito todo o log do sistema
'
'
Function Exporta()

    Dim tipoErro As String                          'indica para o arquivo de log o tipo de erro que pode estar acontecendo
    
    tipoErro = "Erro não localizado"
    'Neste arquivo existe gravado o nome do usuário ativo, que indica que usuário exportou para o Epanet o polígono de seleção
    'Através dele será feita uma pesquisa no banco de dados POLIGONO_SELECAO para ver que redes OBJECT_ID_s serão exportados para o EPANET
    'Lembrando que
    '                TIPO = 0 - Nós
    '                TIPO = 1 - Redes
    '                TIPO = 2 - Ramais

    banco.AtualizaAsRugosidades                         'realiza a atualização, no banco geográfico do GeoSan, de todas as rugosidades a partir do material de cada rede de água
    banco.ZeraMateriais                                 'Zera todos os materiais de tubulações quando o mesmo não estiver cadastrado
    banco.AbrirEstruturaExporteRede
    usuario_global = trechos.ObtemUsuario
    
    banco.RsGlobal_abreconexaoComTrechosAExportar (usuario_global)       'para poder começar a ler os trechos que serão exportados
      
    ExportaParaEpanet                           'inicia a exportação para o EPANET
    
    banco.GeraArquivo_de_Saida                        'Gera o arquivo .INP de saída para o Epanet

    Screen.MousePointer = vbNormal
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
    strtot = strtot & " WHERE WATERLINES.OBJECT_ID_ IN (SELECT OBJECT_ID_ FROM POLIGONO_SELECAO WHERE USUARIO = '" & trechos.ObtemUsuario & "' AND TIPO = 1)"
    Set Rs = New ADODB.Recordset
    'Obtem o número total de trechos de rede a serem exportados para o Epanet os quais são do TIPO=1 na tabela POLIGONO_SELECAO
    Rs.Open strtot, conn, adOpenDynamic, adLockReadOnly
    totalTrechosExportar = Rs(0).Value      'obtem o número total de trechos de rede que serão exportados
    Rs.Close
    Set Rs = Nothing
    ObtemNumeroTrechosQueSeraoExportados = totalTrechosExportar
End Function

'Obtem o nome do usuário que está exportando para o EPANET
'
'
Function ObtemUsuario() As String
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
    usuario_global = usuario
    ObtemUsuario = usuario
End Function

'Exporta para o EPANET
'
'
Function ExportaParaEpanet()
    Dim no As String                                                    'Vaviavel que guadará o nó a ser processado
    Dim contadorTrechos As Integer                                      'para contar quantos trechos está exportando
    Dim mPROVEDOR As String
    Dim mSERVIDOR As String
    Dim mPORTA As String
    Dim mBANCO As String
    Dim mUSUARIO As String
    Dim Senha As String
    Dim decriptada As String
    Dim prov As String
    Dim numeroErro As String                                            'para auxiliar a identificar onde ocorreu o erro
    Dim conta_no As Integer                                             'Variável contador para repetição do processo para o no inicial e final de cada trecho
    Dim Lin_len As Double, x As Double, y As Double                     'Variáveis que retornarão a posição do ponto virtual
    Dim NoI As String, NoF As String                                    'Variáveis que guardarão os nós para inserção do trecho
    Dim retorno As Double, distancia As Double, novoLinLen As Double
    Dim numTotalVerticesNaLinha As Integer
    Dim i As Integer
    Dim teNet As New TECOMNETWORKLib.TeNetwork
    Dim geom_id As Long, rightside As Long, adjust As Long               'Variáveis da biblioteca
    Dim Object_id As String
    Dim xpinter As Double, ypinter As Double, metricValue As Double
    Dim verticeInicial_x As Double, verticeInicial_y As Double, vertice_Y As Double, vertice_X As Double
    Dim idLinha As String
    Dim idNo As String
    
    
    contadorTrechos = 0                                                 'Informa que o contador de trechos exportados é zero
    If (az <> 10) Then
        mSERVIDOR = ReadINI("CONEXAO", "SERVIDOR", App.Path & "\CONTROLES\GEOSAN.ini")
        mPORTA = ReadINI("CONEXAO", "PORTA", App.Path & "\CONTROLES\GEOSAN.ini")
        mBANCO = ReadINI("CONEXAO", "BANCO", App.Path & "\CONTROLES\GEOSAN.ini")
        mUSUARIO = ReadINI("CONEXAO", "USUARIO", App.Path & "\CONTROLES\GEOSAN.ini")
        Senha = ReadINI("CONEXAO", "SENHA", App.Path & "\CONTROLES\GEOSAN.ini")
        prov = ReadINI("CONEXAO", "PROVEDOR", App.Path & "\CONTROLES\GEOSAN.ini")
        decriptada = FunDecripta(Senha)
        az = 10
    End If
    tb.Provider = Provider                                              'Configura o objeto tb(Tecomdatabase) que será usado para retornar para as variáveis lin_len, x e y seus valores para cada trecho
    tb.Connection = conn
    tb.setCurrentLayer "WATERLINES"
    teNet.Provider = 1
    teNet.Connection = conn                                             'configura o componente para conexao com o banco de dados
    teNet.setCurrentLayer "WATERLINES"                                  'seta o plano "WATERLINES" como corrente
    trechos.AbrirEstruturaExporteRede                                   'Abre os cursores que guardarão os objectos do rede(nos,trechos,etc) em memoria para serem gerados em arquivo txt
    banco.RsWaterCompTypes_Inicializa                                   'para ver quais os tipos de parâmetros para redes, bombas, válvulas, etc. e cria um vetor com todos eles
    banco.CarregaRsNosTMP                                               'Carrega RsNosTMP em memoria, tranferes os dados do cursor no servidor para a maquina rsNosTmp = rsNos
        
    Do While Not banco.Rs_trechosDeRede_terminouDeLer = True                    'Percorre todos os trechos da tabela waterlines com a clausura where (setor e tipo de rede), com o polígono de seleção, ativada e cursor iniciando no primeiro registro
        DoEvents
        If Cancelar = True Then
           Return
        End If
        contadorTrechos = contadorTrechos + 1                                   'Incrementa o contador de trechos lidos
        For conta_no = 1 To 2                                                   'para cada um dos nós do trecho d rede, atribui o ID do nó ao nó inicial ou final. Percorre os dois nós do Trecho
            If conta_no = 1 Then                                                'se estamos com o nó inicial do trecho, primeira passagem do FOR
                no = banco.rs_trechosDeRede_numeroNoInicial
            Else                                                                'se estamos com o nó final do trecho, segunda passagem do FOR
                no = banco.rs_trechosDeRede_numeroNoFinal
            End If
            If Not banco.RsTrechosExportados_verificaSeFoiExportado(banco.rs_trechosDeRede_ObjectIdDoTrecho) Then           'verifica se o trecho (selecionado pela cerca do GeoSan) não foi exportado, se não foi exportado, vamos processar
                idLinha = banco.rs_trechosDeRede_ObjectIdDoTrecho               'obtem o OBJECT_ID do trecho de rede (selecionado pela cerca) que estamos exportando para o EPANET
                idNo = no
                If conta_no = 1 Then                                            'é o nó inicial do trecho de rede (primeira passagem), atribui ao nó inicial
                   NoI = no
                Else                                                            'é o nó final do trecho de rede (primeira passagem), atribui ao nó final
                   NoF = no
                End If
                If Not banco.RsNosExportados_verificaSeNoFoiExportado(no) Then  'se o nó não foi exportado ainda
                    banco.RsNosTmp_filtraPorNo (no)                             'filtra para pegar os dados somente do nó que está sendo lido
                    strTipoComp = banco.RsWaterCompTypes_tipoDeEspecificacao    'Obter a especificação para poder processar a exportação conforme o tipo de nó
                    If banco.RsWaterCompTypes_ExitemParametros = True Then
                        Select Case strTipoComp                                 'Seleciona do processo a ser usado para o tipo do nó
                            'Case No_Bombas, No_Valvulas, No_Valvulas_99 'Especial nó
                            'É uma verificação realizada na coluna Specification_ da tabela WaterComponentsTypes
                            'Mudando nesta tabela o valor de Specification_ muda se o nó será exportado como válvula (pequeno trecho) ou não
                            Case "PUMP", "VALVE", "VALVE2", "REGISTER"          'vai dividir uma válvula, registro, etc, que é representado por um nó, em dois nós e um trecho
                                Call ConverteNoEmTrecho(idLinha, Lin_len, conta_no, numTotalVerticesNaLinha, distancia)
                            Select Case strTipoComp
                                Case "PUMP"
                                    Call banco.RsPumps_adicionaBomba(no)        'adiciona os dados da bomba para poder exportar depois para o EPANET
                                    AddSubItemPumps no                          'Adiciona os sub itens para a bomba (curva)
                                Case "VALVE"
                                    Call banco.RsValves_adicionaValvula(no)     'adiciona os dados da válvula para poder exportar depois para o EPANET
                                Case "VALVE2"
                                    Call banco.RsValves2_adicionaValvula(no)    'adiciona os dados da válvula2 para poder exportar depois para o EPANET
                                Case "REGISTER"
                                    Call banco.RsPipes_adicionaRegistro(no)     'adiciona os dados do registro, na verdade um registro é um trecho de rede, para poder depois exportar para o EPANET
                            End Select
                        End Select
                    End If
                    Select Case strTipoComp                                     'Insere o nó processado
                        Case "RNV" 'No_Tanques
                            Call banco.RsTanks_adicionaReservatorioNivelVariavel(no)    'Adiciona os dados do RNV (Reservatório de Nível Variável) para poder exportar para o EPANET
                        Case "RNF" 'No_Reservatorios
                            Call banco.RsReservois_adicionaReservatorio(no)             'Adiciona os dados do reservatório para poder exportar para o EPANET
                        Case Else
                            Call banco.RsJunctions_adicionaJuncao(no, banco.RsNosTmp_obtemDemandaDoNo)  'Adiciona os dados da junção para poder exportar depois para o EPANET
                    End Select
                    Call banco.RsCoordinates_adicionaCoordenadas(no, banco.RsNosTmp_coordX(no), banco.RsNosTmp_coordY(no))  'Adiciona as coordenadas geográficas de um nó, para o mesmo poder exportar para o EPANET
                End If
                banco.RsNosExportados_insereNoExportado (no)                    'Insere o NO no cursor temporário rsNosExportados
            End If
        Next
        'Insere o trecho no cursor temporário rsPipes
        Call banco.RsPipes_adicionaTrechoRede(idLinha, NoI, NoF)
        
        numTotalVerticesNaLinha = tb.getQuantityPointsLine(0, idLinha)                                  'obtem o número de vértices existentes no trecho de rede que foi inserido
        If numTotalVerticesNaLinha > 2 Then                                                             'existem vértice intermediários na linha que necessitam ser considerados no Epanet, pois 2 já existem, o inícial e final
            For i = 1 To numTotalVerticesNaLinha - 2                                                    'percorra do primeiro ao último vértice do trecho de rede
                If tb.getPointOfLine(0, idLinha, i, vertice_X, vertice_Y) Then                          'obtem as coordenadas do vértice
                    Call banco.RsVertices_adicionaVerticesDosTrechosRede(idLinha, vertice_X, vertice_Y)      'insere o vértice
                End If
            Next
        End If

        banco.RsTrechos_adicionaIdDoTrecho                                                              'Adiciona o Id do trecho de rede que já foi exportado para o vetor de trechos exportados
        
        If FrmEPANET.ProgressBar1.Value < FrmEPANET.ProgressBar1.Max Then                               'Atualiza o formulario frmOdometro, para que o usuário saiba que o processamento está andando
            FrmEPANET.ProgressBar1.Value = FrmEPANET.ProgressBar1.Value + 1
        End If
        'banco.RsGlobal_moveParaProximo                                                                  'Move para o próximo trecho de rede selecionado pela cerca no GeoSan
        banco.Rs_trechosDeRede_moveParaProximo
    Loop
End Function

'Converte um ponto (nó) em trecho.
'Isto é realizado quando tempos bomba, válvulas ou registros, pois no EPANET estes elementos são trechos de rede e não nós
'
'
'
Private Function ConverteNoEmTrecho(idLinha As String, Lin_len As Double, conta_no As Integer, numTotalVerticesNaLinha_ As Integer, distancia As Double)
    'Verifica a Direção da Tubulação e recupera um ponto X_NO_VIRT, Y_NO_VIRT a 1/3 da distancia
    ' e insere o no virtual em RsJuntions e RsCoordinates
    tb.getLengthOfLine 0, idLinha, Lin_len
    'carrega em Lin_len o comprimento total da linha
    Dim retorno As Double
    Dim X_Componente As Double 'COORDENADA X DO VERTICE
    Dim Y_Componente As Double 'COORDENADA Y DO VERTICE
    Dim X_Vertice As Double    'COORDENADA X DO VERTICE
    Dim Y_Vertice As Double    'COORDENADA Y DO VERTICE
    Dim VERTICE_LEN As Double  'ARMAZENA O COMPRIMENTO DO VERTICE ANALIZADO
    Dim X_NO_VIRT As Double    'ARMAZENA A COORDENADA X DO NÓ VIRTUAL CASO SEJA NECESSÁRIO
    Dim Y_NO_VIRT As Double    'ARMAZENA A COORDENADA Y DO NÓ VIRTUAL CASO SEJA NECESSÁRIO
    If conta_no = 2 Then ' se for o ponto final pega a distância de 2/3 do comprimento da linha
        numTotalVerticesNaLinha_ = tb.getQuantityPointsLine(0, idLinha) 'retorna número de pontos que compõem a linha. se maior que 2 significa que tem vertices
        If numTotalVerticesNaLinha_ > 2 Then 'existem vértices na linha
            'Pegar o penultimo ponto(vertice)
            'RETORNA A COORDENADA DO ÚLTIMO VERTICE
            retorno = tb.getPointOfLine(0, idLinha, (numTotalVerticesNaLinha_ - 2), X_Vertice, Y_Vertice)
            'RETORNA A COORDENADA DO ÚLTIMO NÓ
            retorno = tb.getPointOfLine(0, idLinha, (numTotalVerticesNaLinha_ - 1), X_Componente, Y_Componente)
            'RETORNA A DISTANCIA ENTRE O ULTIMO NÓ E O ULTIMO VERTICE
            VERTICE_LEN = DistanceBetween(X_Vertice, Y_Vertice, X_Componente, Y_Componente)
            'DISTANCIA = COMPRIMENTO TOTAL DA LINHA - COMPRIMENTO DO ULTIMO VERTICE +
            '2 TERÇOS DA DISTANCIA DO ULTIMO VERTICE
            distancia = (Lin_len - VERTICE_LEN) + (VERTICE_LEN * 0.666666)
            'CARREGA EM X_NO_VIRT E Y_NO_VIRT AS COORDENADAS DE LOCALIZAÇÃO DO PONTO
            'VIRTUAL QUE DEVERÁ SER CRIADO
            tb.getPerpendicularPoint 0, idLinha, distancia, 0, X_NO_VIRT, Y_NO_VIRT
        Else
            'DISTANCIA = 2 TERÇOS DO COMPRIMENTO TOTAL DA LINHA
            distancia = Lin_len * 0.666666
            
            'CARREGA EM X_NO_VIRT E Y_NO_VIRT AS COORDENADAS DE LOCALIZAÇÃO DO PONTO
            'VIRTUAL QUE DEVERÁ SER CRIADO
            tb.getPerpendicularPoint 0, idLinha, distancia, 0, X_NO_VIRT, Y_NO_VIRT
        End If
    Else ' se o ponto for o inicial, pega a distãncia de 1/3 do comprimento da linha
        numTotalVerticesNaLinha_ = tb.getQuantityPointsLine(0, idLinha)
        If numTotalVerticesNaLinha_ > 2 Then 'existem vértices na linha
            'RETORNA A COORDENADA DO PRIMEIRO VERTICE
            retorno = tb.getPointOfLine(0, idLinha, 1, X_Vertice, Y_Vertice)
            'RETORNA A COORDENADA DO PRIMEIRO NÓ
            retorno = tb.getPointOfLine(0, idLinha, 0, X_Componente, Y_Componente)
            'RETORNA EM VERTICE_LEN A DISTANCIA ENTRE O PRIMEIRO VERTICE E O PRIMEIRO NÓ
            VERTICE_LEN = DistanceBetween(X_Vertice, Y_Vertice, X_Componente, Y_Componente)
            'DISTANCIA = COMPRIMENTO TOTAL DA LINHA - COMPRIMENTO DO ULTIMO VERTICE +
            '2 TERÇOS DA DISTANCIA DO ULTIMO VERTICE
            distancia = VERTICE_LEN * 0.33333
            'CARREGA EM X_NO_VIRT E Y_NO_VIRT AS COORDENADAS DE LOCALIZAÇÃO DO PONTO
            'VIRTUAL QUE DEVERÁ SER CRIADO
            tb.getPerpendicularPoint 0, idLinha, distancia, 0, X_NO_VIRT, Y_NO_VIRT
        Else
            'DISTANCIA = 1 TERÇO DO COMPRIMENTO TOTAL DA LINHA
            distancia = Lin_len * 0.333333
            'CARREGA EM X_NO_VIRT E Y_NO_VIRT AS COORDENADAS DE LOCALIZAÇÃO DO PONTO
            'VIRTUAL QUE DEVERÁ SER CRIADO
            tb.getPerpendicularPoint 0, idLinha, distancia, 0, X_NO_VIRT, Y_NO_VIRT
        End If
        ' retorna em x, y a coordenada do ponto inicial que fica a 1/3 do início da linha
        'tb.getPerpendicularPoint 0, .Fields("object_id_").Value, (Lin_len / 3) * 2, 0, X_Componente, Y_Componente ' colocou zero antes de x, y para retornar o ponto na própria linha
    End If
    Call banco.RsJunctions_adicionaJuncao(no, 0#)                                   'adiciona o nó ao vetor de junções para depois poder exportá-lo para o EPANET - coloca a demanda como zero, pois é um nó convertido em trecho por ser válvula, bomba ou registro
    Call banco.RsCoordinates_adicionaCoordenadas(no & "A", X_NO_VIRT, Y_NO_VIRT)    'Adiciona as coordenadas geográficas de um nó, para o mesmo poder exportar para o EPANET - soma a letra A ao numero do nó, pois converteu o mesmo em dois, para representar um trecho, como o EPANET exige
    '###########################################################
    'Alterar a variável NoI ou NoF com o valor NO & "A" para ser usado na inserção do trecho
    'indicando que o trecho a ser inserido usará um nó virtual
    If conta_no = 1 Then
       NoI = no & "A"
    Else
       NoF = no & "A"
    End If
    '######################
    'Cria o Componente do tipo entre o nó virtual e o nó processado
End Function


Function NoCadastrada(Object_id_ As String) As Boolean
   'Verifica se o nó já foi cadastrado
   rsNosExportados.Filter = "id='" & Object_id_ & "'"
   NoCadastrada = Not rsNosExportados.EOF
End Function
'Define a estrutura dos vetores que conterão os dados que serão exportados para o Epanet
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
    
'    Open arquivoLog For Append As #5                                'continua a realizar o log do sistema
'    Print #5, vbCrLf & "CarregaRsNosTMP; Inicia o registro do resultado dos nós temporários"
'
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

'    Print #5, "CarregaRsNosTMP; string conexão: " & strSQL & " conexão: " & conn_global
'    Close #5
'    Open arquivoLog For Append As #5                                'continua a realizar o log do sistema
    Set rsNos = New ADODB.Recordset
    
    'conn.Open
    rsNos.Open strSQL, conn ', adOpenDynamic, adLockReadOnly

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

    'rsNos.Close


     'Set rsNos = Nothing
    'conn.Close
    'AO FINAL DESTA ROTINA FICARÁ EXISTENTE A TABELA DE NÓS COMPLETA DESNORMATIZADA (rsNosTmp)
    'Close #5
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

'
'
'
Private Sub AddSubItemTank(id As String)
   Dim Rs As ADODB.Recordset
   Dim CURVE As String

   Set Rs = conn.Execute("Select b.eparef, w.value_, s.description_ from watercomponentsdata w " & _
                         "INNER JOIN watercomponentssubtypes b on b.id_subtype=w.id_subtype and b.id_type=w.id_type " & _
                         "LEFT JOIN WaterComponentsSelections s on s.id_subtype=w.id_subtype and s.id_type=w.id_type and cast(w.value_ as INT)=s.value_ " & _
                         "where object_id_ = '" & id & "'")
   rsTanks.Fields("MaxLevel").Value = 0
   rsTanks.Fields("MinLevel").Value = 0
   rsTanks.Fields("InitLevel").Value = 0
   rsTanks.Fields("Diameter").Value = 0
   rsTanks.Fields("MinVol").Value = 0
   rsTanks.Fields("VolCurve").Value = " "
   
   While Not Rs.EOF
      Select Case Rs.Fields("EPAREF").Value
      
         Case "NMAXIMO"
            rsTanks.Fields("MaxLevel").Value = Replace(Rs.Fields("VALUE_").Value, ",", ".")
         Case "NMINIMO"
            rsTanks.Fields("MinLevel").Value = Replace(Rs.Fields("VALUE_").Value, ",", ".")
         Case "NINICIAL"
            rsTanks.Fields("InitLevel").Value = Replace(Rs.Fields("VALUE_").Value, ",", ".")
         Case "DIAMETER"
            rsTanks.Fields("Diameter").Value = Replace(Rs.Fields("VALUE_").Value, ",", ".")
         Case "VOLUME"
            rsTanks.Fields("MinVol").Value = Replace(Rs.Fields("VALUE_").Value, ",", ".")
         Case "VOLCURVE"
            rsTanks.Fields("VolCurve").Value = Replace(Rs.Fields("VALUE_").Value, ",", ".")
      End Select
      Rs.MoveNext
   Wend
   Rs.Close
   Set Rs = Nothing
End Sub

'Atribui o valor de curva para bomba
'
'
Private Sub AddSubItemPumps(id As String)
   Dim Rs As ADODB.Recordset
   Dim CURVE As String
   Set Rs = conn.Execute("Select b.eparef, w.value_, s.description_ from watercomponentsdata w " & _
                         "INNER JOIN watercomponentssubtypes b on b.id_subtype=w.id_subtype and b.id_type=w.id_type " & _
                         "LEFT JOIN WaterComponentsSelections s on s.id_subtype=w.id_subtype and s.id_type=w.id_type and cast(w.value_ as INT)=s.value_ " & _
                         "where object_id_ = '" & id & "'")
                         
   While Not Rs.EOF
      Select Case Rs.Fields("EPAREF").Value
         Case "CURVE"
            rsPumps.Fields("Parameters").Value = " HEAD " & Rs.Fields("VALUE_").Value
      End Select
      Rs.MoveNext
   Wend
   Rs.Close
   Set Rs = Nothing
End Sub
'Adiciona as outras componentes de uma válvula, como diametro, tipo, etc.
'
' id - número do nó
'
' Esta rotina não está funcionando a querie abaixo está com colunas não existentes e pelo
' que foi verificado não retorna linhas. Em nosso entendimento ela não tem função 2013-04-27
'
'
Private Sub AddSubItemValves(id As String)
'Atribui os valores especificos para valvula
    Dim Rs As ADODB.Recordset
    Dim PumpDiameter As Double, PumpType As String, PumpSetting As String, PumpMinorLoss As String
    If conn.Provider <> "PostgreSQL.1" Then
        Set Rs = conn.Execute("Select b.eparef, w.value_, s.description_ from watercomponentsdata w " & _
                                "INNER JOIN watercomponentssubtypes b on b.id_subtype=w.id_subtype and b.id_type=w.id_type " & _
                                "LEFT JOIN WaterComponentsSelections s on s.id_subtype=w.id_subtype and s.id_type=w.id_type and cast(w.value_ as INT)=s.value_ " & _
                                "where object_id_ = '" & id & "'")
    Else
        Set Rs = conn.Execute("Select " + """" + "WATERCOMPONENTSSUBTYPES" + """" + "." + """" + "EPAREF" + """" + "," + """" + "WATERCOMPONENTSDATA" + """" + "." + """" + "VALUE_" + """" + "," + """" + "WATERCOMPONENTSSELECTIONS" + """" + "." + """" + "DESCRIPTION_" + """" + " from " + """" + "WATERCOMPONENTSDATA" + """" & _
                                "INNER JOIN " + """" + "WATERCOMPONENTSSUBTYPES" + """" + " on " + """" + "WATERCOMPONENTSSUBTYPES" + """" + "." + """" + "ID_SUBTYPE" + """" + "=" + """" + "WATERCOMPONENTSDATA" + """" + "." + """" + "ID_SUBTYPE" + """" + " and " + """" + "WATERCOMPONENTSSUBTYPES" + """" + "." + """" + "ID_TYPE" + """" + "=" + """" + "WATERCOMPONENTSDATA" + """" + "." + """" + "ID_TYPE" + """" & _
                                "LEFT JOIN " + """" + "WATERCOMPONENTSSELECTIONS" + """" + " on " + """" + "WATERCOMPONENTSSELECTIONS" + """" + "." + """" + "ID_SUBTYPE" + """" + "=" + """" + "WATERCOMPONENTSDATA" + """" + "." + """" + "ID_SUBTYPE" + """" + " and " + """" + "WATERCOMPONENTSSELECTIONS" + """" + "." + """" + "ID_TYPE" + """" + "=" + """" + "WATERCOMPONENTSDATA" + """" + "." + """" + "ID_TYPE" + """" + " and cast(" + """" + "WATERCOMPONENTSDATA" + """" + "." + """" + "VALUE_" + """" + " as INT) =" + """" + "WATERCOMPONENTSSELECTIONS" + """" + "." + """" + "VALUE_" + """" & _
                                "where " + """" + "OBJECT_ID_" + """" + " = '" & id & "'")
    End If
    While Not Rs.EOF
        Select Case Rs.Fields("EPAREF").Value
            Case "TYPE"
                rsValves.Fields("Type").Value = Rs.Fields("DESCRIPTION_").Value
            Case "SETTING"
                rsValves.Fields("Setting").Value = Rs.Fields("VALUE_").Value
            Case "DIAMETER"
                rsValves.Fields("Diameter").Value = Rs.Fields("VALUE_").Value
            Case "NINORLOSS"
                rsValves.Fields("MinorLoss").Value = 0 'implementar
        End Select
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
End Sub

'Verifica se um trecho de rede já foi cadastrado, pois se já foi, não precisa exportar mais o mesmo
'
'id = número do object_id do trecho de rede
'TrechoCadastrado = True - já foi cadastrado para exportação para o EPANET
'TrechoCadastrado = False - ainda não foi cadastrado para exportação para o EPANET
'
Private Function TrechoCadastrado(id As String) As Boolean
   rsTrechosExportados.Filter = "id='" & id & "'"   'aplica um filtro pelo object_id aos trechos já exportados
   TrechoCadastrado = Not rsTrechosExportados.EOF   'retorna verdadeiro se o trecho se o trecho ja foi exportado, falso se não
End Function
