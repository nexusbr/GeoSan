Attribute VB_Name = "ModExporte"
Option Explicit
'Conexão que seja usado em todo o Processo
Private conn As New ADODB.Connection

'Objecto Utilizado para retornar a posicao
'em que seja colocado o nó virtual e os vétices das rede(linhas)
Private tb As New TeDatabase

'Cursor Temporário que retonará todos os nos da Rede
'para abastecer o RsNosTMP
Private rsNos As Recordset

'Criado para armazenar os trechos que já foram exportados
Private rsTrechosExportados As New ADODB.Recordset
'Criado para armazenar os Nós que já foram exportados
Private rsNosExportados As New ADODB.Recordset

'Criado para armazenar todos os dados de todos nos
'Copia do Watercomponenstes/Points
Private rsNosTmp As New ADODB.Recordset

'Criados Para Armazenar os Componentes / Trechos
Private rsCoordinates As New ADODB.Recordset
Private rsPipes As New ADODB.Recordset
Private rsJunctions As New ADODB.Recordset
Private rsPumps As New ADODB.Recordset
Private rsValves As New ADODB.Recordset
Private rsReservoirs As New ADODB.Recordset
Private rsTanks As New ADODB.Recordset
Private rsVertices As New ADODB.Recordset 'Vertices da linha com exceção do inicial e final
Dim az As Integer
'Variavel que guardará o layer_id dos NOS(Watercomponents)
Private layer_id As Integer

'RecordSet de referencia para consulta do tipo de componente
Private rsWaterCompTypes As New ADODB.Recordset

Private strTipoComp As String 'Variável que receberá o tipo de componente (Junction, valve,pump...)

Public intLinhaCod As Integer 'indicador de linha para tratamento de erro
Public Cancelar As Boolean

Dim blnRsWaterCompTypes As Boolean 'Indicador para informar se a tabela RsWaterCompTypes foi carregada com registros
'FUNÇÕES PARA LER E GRAVAR NO ARQUIVO .INI-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nsize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Lê as informações do arquivo de inicialização do GeoSan
'Arquivo=nome do arquivo ini
'Secao=O que esta entre []
'Entrada=nome do que se encontra antes do sinal de igual
Public Function ReadINI(Secao As String, Entrada As String, Arquivo As String)
    Dim retlen As String
    Dim Ret As String
    Ret = String$(255, 0)
    retlen = GetPrivateProfileString(Secao, Entrada, "", Ret, Len(Ret), Arquivo)
    Ret = Left$(Ret, retlen)
    ReadINI = Ret
End Function
'Procedimento Exporte EPANET recebe como parametros o cursor trazendo todas os trechos
'a serem exportados e o objecto de conexão
'(rsTrechos):É a tabela Waterlines com os filtros de tipo de rede e setor selecionados pelo usuário (TIPO=1 na tabela POLIGONO_SELECAO)
'arquivoLog: nome do arquivo onde está sendo escrito todo o log do sistema
Public Sub ExportaEPANet(rsTrechos As ADODB.Recordset, mconn As ADODB.Connection, arquivoLog As String)
    On Error GoTo Trata_Erro
    Dim numeroErro As String                'para auxiliar a identificar onde ocorreu o erro
    Dim contadorTrechos As Integer          'para contar quantos trechos está exportando
    Dim mPROVEDOR As String
    Dim mSERVIDOR As String
    Dim mPORTA As String
    Dim mBANCO As String
    Dim mUSUARIO As String
    Dim Senha As String
    Dim decriptada As String
    Dim nStr As String
    Dim prov As String
    
    Open arquivoLog For Append As #5                                'continua a realizar o log do sistema
    Print #5, vbCrLf & "ExportaEPANet;*************************************************************************************************"  'Indica que começou a nova fase de exportação
    Print #5, vbCrLf & "ExportaEPANet;Inicia a exportação (2a. fase), onde recebe o cursor pronto com os dados para exportar."
    Print #5, vbCrLf & "ExportaEPANet;Querie recebida: " & rsTrechos.Source
    'Informa que o contador de trechos exportados é zero
    contadorTrechos = 0
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
    If prov = "4-PostgreSQL" Then
        FrmEPANET.TeAcXConnection1.Open mUSUARIO, decriptada, mBANCO, mSERVIDOR, mPORTA
    End If
    Set conn = mconn
    Dim NO As String                                'Vaviavel que guadará o nó a ser processado
    Dim conta_no As Integer                         'Variável contador para repetição do processo para o no inicial e final de cada trecho
    Dim Lin_len As Double, x As Double, y As Double 'Variáveis que retornarão a posição do ponto virtual
    Dim NoI As String, NoF As String                'Variáveis que guardarão os nós para inserção do trecho
   
    'Configura o objeto tb(Tecomdatabase) que será usado para retornar para as variáveis lin_len, x e y
    'seus valores para cada trecho
    If conn.Provider <> "PostgreSQL.1" Then
        'caso Oracle ou SQLServer
        tb.Provider = Provider
        tb.Connection = conn
    Else
        'caso Postgres
        tb.Provider = Provider
        tb.Connection = FrmEPANET.TeAcXConnection1.objectConnection_
    End If
    'tb.setCurrentLayer "waterlines"
    'PADRONIZADO O NOME DAS TABELAS PARA LETRA MAIÚSCULA - Jonathas 19/03/09
    tb.setCurrentLayer "WATERLINES"
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    Dim retorno As Double, distancia As Double, novoLinLen As Double
    Dim numTotalVerticesNaLinha, i As Integer
    Dim teNet As New TECOMNETWORKLib.TeNetwork
    'Variáveis da biblioteca
    Dim geom_id As Long, rightside As Long, adjust As Long
    Dim object_id As String
    Dim xpinter As Double, ypinter As Double, metricValue As Double
    Dim verticeInicial_x As Double, verticeInicial_y As Double, vertice_Y As Double, vertice_X As Double
    teNet.Provider = 1
    'configura o componente para conexao com o banco de dados
    teNet.Connection = conn
    'seta o plano "WATERLINES" como corrente
    teNet.setCurrentLayer "WATERLINES"
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    intLinhaCod = 1
    'Abre os cursores que guardarão os objectos do rede(nos,trechos,etc) em memoria
    'para serem gerados em arquivo txt
    AbrirEstruturaExporteRede
    'RECORDSET COM OS TIPOS DE REDES EXISTENTES
    If conn.Provider <> "PostgreSQL.1" Then
        Print #5, vbCrLf & "ExportaEPANet;SELECT * FROM WATERCOMPONENTSTYPES ORDER BY ID_TYPE"
        Set rsWaterCompTypes = conn.Execute("SELECT * FROM WATERCOMPONENTSTYPES ORDER BY ID_TYPE")
    Else
        Print #5, vbCrLf & "ExportaEPANet;SELECT * FROM " + """" + "WATERCOMPONENTSTYPES" + """" + " ORDER BY " + """" + "ID_TYPE" + """" + ""
        Set rsWaterCompTypes = conn.Execute("SELECT * FROM " + """" + "WATERCOMPONENTSTYPES" + """" + " ORDER BY " + """" + "ID_TYPE" + """" + "")
    End If
    If rsWaterCompTypes.EOF = False Then
        blnRsWaterCompTypes = True
    Else
        Print #5, "ExportaEPANet;Não será possivel identificar e exportar bombas e válvulas pois a tabela WATERCOMPONENTSTYPES está vazia."
        MsgBox "Não será possivel identificar e exportar bombas e válvulas pois a tabela WATERCOMPONENTSTYPES está vazia.", vbExclamation, ""
        blnRsWaterCompTypes = False
    End If
    intLinhaCod = 2
    Close #5            'fecha o arquivo de log pois será reaberto na subrotina CarregaRsNosTMP
    'Carrega RsNosTMP em memoria, tranferes os dados do cursor no servidor para a maquina rsNosTmp = rsNos
    CarregaRsNosTMP arquivoLog
    Open arquivoLog For Append As #5                                'continua a realizar o log do sistema para esta rotina
    intLinhaCod = 3
    'With rsTrechos
    'Percorre todos os trechos da tabela waterlines com a clausura where (setor e tipo de rede)
    'ativada e cursor iniciando no primeiro registro
    Print #5, "ExportaEPANet;Inicia o cursor em cada trecho de rede com a querie: " & rsTrechos.Source
    Do While Not rsTrechos.EOF = True
        intLinhaCod = 161
        DoEvents
        If Cancelar = True Then
           Exit Sub
        End If
        
        'Percorre os dois nós do Trecho
        intLinhaCod = 4
        contadorTrechos = contadorTrechos + 1                                   'Incrementa o contador de trechos lidos
        'Imprime o trecho de rede que ele está lendo para exportar para o Epanet
        Print #5, "ExportaEPANet;" & Now & " - " & contadorTrechos & " - Trecho lido: " & rsTrechos!Object_id_

        For conta_no = 1 To 2
            'Verifica a variável conta_no e atribui a variavel NO o valor do nó inicial ou final
            If conta_no = 1 Then 'refere-se ao nó inicial
                Print #5, "ExportaEPANet;" & Now & " - " & contadorTrechos & " - Trecho lido: " & rsTrechos!Object_id_ & " - Nó inicial: " & rsTrechos!InitialComponent
                NO = rsTrechos.Fields("InitialComponent").Value
            Else 'refere-se ao no final
                Print #5, "ExportaEPANet;" & Now & " - " & contadorTrechos & " - Trecho lido: " & rsTrechos!Object_id_ & " - Nó final: " & rsTrechos!FinalComponent
                NO = rsTrechos.Fields("FinalComponent").Value
            End If
            intLinhaCod = 5
            'Verifica se o trecho não foi cadastrado
            If Not TrechoCadastrado(rsTrechos.Fields("object_id_").Value) Then 'se não foi exportado, vamos processar
                Dim idLinha As String
                Dim idNo As String
                idLinha = rsTrechos.Fields("object_id_").Value
                idNo = NO
                '###########################################################
                'Atribui a variável NoI ou NoF com o valor NO para ser usado na inserção do trecho
                'no procedimento inserção de Trecho
                If conta_no = 1 Then
                   NoI = NO
                Else
                   NoF = NO
                End If
                '######################
                'Verifica se o nó não foi cadastrado
                intLinhaCod = 6
                If Not NoCadastrada(NO) Then 'se o nó não foi cadastrado
                    'Seleciona no Cursor rsNosTmp o nó igual o valor de NO
                    rsNosTmp.Filter = "id='" & NO & "'"
                    'Seleciona do processo a ser usado para o tipo do nó
                    intLinhaCod = 7
                    If rsNosTmp.EOF = False Then
                        strTipoComp = ""
                        If blnRsWaterCompTypes = True Then
                            rsWaterCompTypes.MoveFirst
                            Do While Not rsWaterCompTypes.EOF = True
                                If rsNosTmp.Fields("Tipo").Value = rsWaterCompTypes!id_type Then
                                    strTipoComp = rsWaterCompTypes!SPECIFICATION_
                                    Exit Do
                                End If
                                rsWaterCompTypes.MoveNext
                            Loop
                        End If
                        'Neste primeiro select case ele vai dividir uma válvula, registro, etc, que é representado por um nó, em dois nós.
                        Select Case strTipoComp
                            'Case No_Bombas, No_Valvulas, No_Valvulas_99 'Especial nó
                            Case "PUMP", "VALVE", "VALVE2", "REGISTER"
                                'Verifica a Direção da Tubulação e recupera um ponto X_NO_VIRT, Y_NO_VIRT a 1/3 da distancia
                                ' e insere o no virtual em RsJuntions e RsCoordinates
                                tb.getLengthOfLine 0, idLinha, Lin_len
                                'carrega em Lin_len o comprimento total da linha
                                intLinhaCod = 8
                                Dim X_Componente As Double 'COORDENADA X DO VERTICE
                                Dim Y_Componente As Double 'COORDENADA Y DO VERTICE
                                Dim X_Vertice As Double    'COORDENADA X DO VERTICE
                                Dim Y_Vertice As Double    'COORDENADA Y DO VERTICE
                                Dim VERTICE_LEN As Double  'ARMAZENA O COMPRIMENTO DO VERTICE ANALIZADO
                                Dim X_NO_VIRT As Double    'ARMAZENA A COORDENADA X DO NÓ VIRTUAL CASO SEJA NECESSÁRIO
                                Dim Y_NO_VIRT As Double    'ARMAZENA A COORDENADA Y DO NÓ VIRTUAL CASO SEJA NECESSÁRIO
                                If conta_no = 2 Then ' se for o ponto final pega a distância de 2/3 do comprimento da linha
                                    numTotalVerticesNaLinha = tb.getQuantityPointsLine(0, idLinha) 'retorna número de pontos que compõem a linha. se maior que 2 significa que tem vertices
                                    If numTotalVerticesNaLinha > 2 Then 'existem vértices na linha
                                        'Pegar o penultimo ponto(vertice)
                                        'RETORNA A COORDENADA DO ÚLTIMO VERTICE
                                        retorno = tb.getPointOfLine(0, idLinha, (numTotalVerticesNaLinha - 2), X_Vertice, Y_Vertice)
                                        'RETORNA A COORDENADA DO ÚLTIMO NÓ
                                        retorno = tb.getPointOfLine(0, idLinha, (numTotalVerticesNaLinha - 1), X_Componente, Y_Componente)
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
                                    numTotalVerticesNaLinha = tb.getQuantityPointsLine(0, idLinha)
                                    If numTotalVerticesNaLinha > 2 Then 'existem vértices na linha
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
                                intLinhaCod = 9
                                rsJunctions.AddNew
                                rsJunctions.Fields("id").Value = NO & "A"
                                rsJunctions.Fields("elev").Value = Format(rsNosTmp("cota").Value, ".0")
                                rsJunctions.Fields("demand").Value = 0
                                rsJunctions.Fields("pattern").Value = ""
                                rsCoordinates.AddNew
                                rsCoordinates.Fields("id").Value = NO & "A"
                                rsCoordinates.Fields("x").Value = X_NO_VIRT
                                rsCoordinates.Fields("y").Value = Y_NO_VIRT
                                '###########################################################
                                'Alterar a variável NoI ou NoF com o valor NO & "A" para ser usado na inserção do trecho
                                'indicando que o trecho a ser inserido usará um nó virtual
                                If conta_no = 1 Then
                                   NoI = NO & "A"
                                Else
                                   NoF = NO & "A"
                                End If
                                '######################
                                'Cria o Componente do tipo entre o nó virtual e o nó processado
                                 intLinhaCod = 10
                            'Select Case rsNosTmp.Fields("Tipo").Value
                            'Agora adiciona aos recosets
                            Select Case strTipoComp
                                Case "PUMP"
                                    rsPumps.AddNew
                                    rsPumps.Fields("id").Value = NO
                                    If conta_no = 1 Then
                                       rsPumps.Fields("Node1").Value = NO          'NESTE TRECHO É DEFINIDO O SENTIDO DA BOMBA
                                       rsPumps.Fields("Node2").Value = NO & "A"    'A ORDEM DE NO E NOA INFLUENCIA NO SENTIDO
                                    Else
                                       rsPumps.Fields("Node1").Value = NO & "A"
                                       rsPumps.Fields("Node2").Value = NO
                                    End If
                                    AddSubItemPumps NO  'Adiciona os sub itens para a bomba (curva)
                                Case "VALVE"
                                    rsValves.AddNew
                                    rsValves.Fields("ID").Value = "V" & NO
                                    rsValves.Fields("Node1").Value = NO & "A"
                                    rsValves.Fields("Node2").Value = NO
                                    AddSubItemValves NO 'Adiciona os sub itens para a valvula (setting,type,diameter)
                                Case "VALVE2"
                                    rsPipes.AddNew
                                    rsPipes.Fields("id").Value = NO & "A"
                                    rsPipes.Fields("node1").Value = NO & "A"
                                    rsPipes.Fields("node2").Value = NO
                                    rsPipes.Fields("length").Value = 0.1
                                    rsPipes.Fields("diameter").Value = Replace(rsTrechos.Fields("internaldiameter").Value, ",", ".")
                                    rsPipes.Fields("roughness").Value = Replace(rsTrechos.Fields("roughness").Value, ",", ".")
                                    rsPipes.Fields("status").Value = IIf(rsNosTmp("estado").Value = 0, " ", rsNosTmp("estado").Value)
                                    If rsTrechos.Fields("MATERIALNAME").Value <> "" Then
                                       rsPipes.Fields("Description").Value = rsTrechos.Fields("MATERIALNAME").Value
                                    Else
                                       rsPipes.Fields("Description").Value = ""
                                    End If
                                Case "REGISTER"
                                    rsPipes.AddNew
                                    rsPipes.Fields("ID").Value = NO & "A"
                                    rsPipes.Fields("NODE1").Value = NO & "A"
                                    rsPipes.Fields("NODE2").Value = NO
                                    rsPipes.Fields("LENGTH").Value = 1
                                    rsPipes.Fields("DIAMETER").Value = Replace(rsTrechos.Fields("internaldiameter").Value, ",", ".")
                                    rsPipes.Fields("ROUGHNESS").Value = Replace(rsTrechos.Fields("roughness").Value, ",", ".")
                                    rsPipes.Fields("STATUS").Value = IIf(rsNosTmp("ESTADO").Value = 0, " ", rsNosTmp("ESTADO").Value)
                                    rsPipes.Fields("DESCRIPTION").Value = "REGISTRO"
                            End Select
                        End Select
                    End If 'rsNosTmp.EOF = true
                    'Insere o nó processado
                    intLinhaCod = 11
                    'Select Case rsNosTmp.Fields("Tipo").Value
                    Select Case strTipoComp
                        Case "RNV" 'No_Tanques
                            rsTanks.AddNew
                            rsTanks.Fields("ID").Value = NO
                            rsTanks.Fields("Elevation").Value = Format(rsNosTmp("cota").Value, ".0")
                            AddSubItemTank NO
                        Case "RNF" 'No_Reservatorios
                            rsReservoirs.AddNew
                            rsReservoirs.Fields("ID").Value = NO
                            rsReservoirs.Fields("Head").Value = ""
                            rsReservoirs.Fields("Pattern").Value = ""
                        Case Else
                            rsJunctions.AddNew
                            rsJunctions.Fields("id").Value = NO
                            rsJunctions.Fields("elev").Value = Format(rsNosTmp("cota").Value, ".0")
                            rsJunctions.Fields("demand").Value = rsNosTmp("demanda").Value
                            rsJunctions.Fields("pattern").Value = IIf(rsNosTmp("padrao").Value = 0, "", rsNosTmp("padrao").Value)
                    End Select
                    intLinhaCod = 12
                    rsCoordinates.AddNew
                    rsCoordinates.Fields("id").Value = NO
                    rsCoordinates.Fields("x").Value = rsNosTmp("x").Value
                    rsCoordinates.Fields("y").Value = rsNosTmp("y").Value
                End If
                intLinhaCod = 13
                'Insere o NO no cursor temporário rsNosExportados
                rsNosExportados.AddNew
                rsNosExportados.Fields("id").Value = NO
            End If
        Next
        'Insere o trecho no cursor temporário rsPipes
        intLinhaCod = 14
        rsPipes.AddNew
        rsPipes.Fields("id").Value = idLinha
        rsPipes.Fields("node1").Value = NoI
        rsPipes.Fields("node2").Value = NoF
        If rsTrechos.Fields("Length").Value > 0 Then
            rsPipes.Fields("length").Value = Replace(rsTrechos.Fields("Length").Value, ",", ".")
        Else
            rsPipes.Fields("Length").Value = Replace(rsTrechos.Fields("LengthCalculated").Value, ",", ".")
        End If
        rsPipes.Fields("diameter").Value = Replace(rsTrechos.Fields("internaldiameter").Value, ",", ".")
        rsPipes.Fields("roughness").Value = Replace(rsTrechos.Fields("roughness").Value, ",", ".")
        'rsPipes.Fields("status").Value = IIf(rsNosTmp("estado").Value = 0, " ", rsNosTmp("estado").Value)
        rsPipes.Fields("status").Value = "Open"
        If rsTrechos.Fields("MATERIALNAME").Value <> "" Then
            rsPipes.Fields("Description").Value = rsTrechos.Fields("MATERIALNAME").Value
        Else
            rsPipes.Fields("Description").Value = ""
        End If
        'rsPipes.Fields("Description").Value = IIf(rsTrechos.Fields("MATERIAL").Value <> Null, "", rsTrechos.Fields("MATERIAL").Value)
        ' INICIO DA ROTINA DE INSERIR OS VÉRTICES CASO NECESSÁRIO
        numTotalVerticesNaLinha = tb.getQuantityPointsLine(0, idLinha)
        If numTotalVerticesNaLinha > 2 Then ' existem vértice intermediários na linha que necessitam ser considerados no Epanet
            For i = 1 To numTotalVerticesNaLinha - 2 'DO PRIMEIRO AO ULTIMO VÉRTICE, FAÇA
                If tb.getPointOfLine(0, idLinha, i, vertice_X, vertice_Y) Then
                    rsVertices.AddNew
                    rsVertices.Fields("ID").Value = idLinha
                    rsVertices.Fields("X-Coord").Value = vertice_X
                    rsVertices.Fields("Y-Coord").Value = vertice_Y
                End If
            Next
        End If
        ' FIM DA ROTINA DE INSERIR VERTICES
        intLinhaCod = 15
        'AddVertices id
        'Insere no cursor de trechos exportados o trecho processado
        rsTrechosExportados.AddNew
        intLinhaCod = 151
        rsTrechosExportados.Fields("id").Value = idLinha 'rsTrechos.Fields("object_id_").Value
        'Atualiza o formulario frmOdometro
        intLinhaCod = 152
        If FrmEPANET.ProgressBar1.Value < FrmEPANET.ProgressBar1.Max Then
            FrmEPANET.ProgressBar1.Value = FrmEPANET.ProgressBar1.Value + 1
        End If
        'frmOdometro.Caption = FrmEPANET.ProgressBar1.Value & " até " & frmOdometro.ProgressBar1.Max
        intLinhaCod = 153
        'Move o ponteiro do cursor de trechos a serem processados para a próxima tupla
        intLinhaCod = 16
        numeroErro = "Penuntimo trecho lido: " & rsTrechos!Object_id_
        Print #5, "ExportaEPANet;Querie rsTrechos: " & rsTrechos.Source
        rsTrechos.MoveNext
        numeroErro = numeroErro + " Ultimo trecho lido: " & rsTrechos!Object_id_
    Loop
    'End With
    intLinhaCod = 17
    Set rsNosTmp = Nothing
    rsTrechos.Close
    Set rsTrechos = Nothing
    rsTrechosExportados.Close
    Set rsTrechosExportados = Nothing
    Set rsNos = Nothing
    intLinhaCod = 18
    'Gera o arquivo .INP de saída para o Epanet
    GeraArquivo_de_Saida
    intLinhaCod = 19
    Screen.MousePointer = vbNormal
    Print #5, vbCrLf & "ExportaEPANet;*************************************************************************************************"
    Print #5, "ExportaEPANet;Exportação concluída com sucesso!"
    MsgBox "Exportação concluída com sucesso!", vbInformation, "Exporte Epanet"
    Close #5
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Or Err.Number = 3021 Then
        'O código -2147467259 é um erro de timeout expired que aconteceu e quando foram criados os índices foi corrigido.
        'O código 3021 é final de arquivo
        Resume Next
    Else
        Close #5
        Open arquivoLog For Append As #5
        Print #5, vbCrLf & "ExportaEPANet;" & Now & "  - ModExporte - Sub ExportaEPANet(rsTrechos As ADODB.Recordset, mconn As ADODB.Connection) - Linha: " & intLinhaCod & " - " & Err.Number & " - " & Err.Description
        Print #5, "ExportaEPANet;Um posssível erro foi identificado na rotina 'ExportaEPANet':" & Chr(13) & Chr(13) & Err.Description & Chr(13) & numeroErro & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo LogErroExportEPANET.txt com informações desta ocorrencia.", vbInformation
        Close #5
        MsgBox "Um posssível erro foi identificado na rotina 'ExportaEPANet':" & Chr(13) & Chr(13) & Err.Description & Chr(13) & numeroErro & Chr(13) & "Foi gerado na pasta do aplicativo o arquivo LogErroExportEPANET.txt com informações desta ocorrencia.", vbInformation
    End If
End Sub
Public Function DistanceBetween(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double) As Double
  ' Calculate the distance between two points, given their X/Y coordinates.
  
  ' The short version...
  DistanceBetween = Sqr((Abs(X2 - X1) ^ 2) + (Abs(Y2 - Y1) ^ 2))
  
End Function

Sub AddSubItemTank(id As String)
   Dim Rs As ADODB.Recordset
   Dim CURVE As String
   If conn.Provider <> "PostgreSQL.1" Then
   Set Rs = conn.Execute("Select b.eparef, w.value_, s.description_ from watercomponentsdata w " & _
                         "INNER JOIN watercomponentssubtypes b on b.id_subtype=w.id_subtype and b.id_type=w.id_type " & _
                         "LEFT JOIN WaterComponentsSelections s on s.id_subtype=w.id_subtype and s.id_type=w.id_type and cast(w.value_ as INT)=s.value_ " & _
                         "where object_id_ = '" & id & "'")
                         Else
                         
                    
                         
                         
                         
                         
                         
                         
                         Set Rs = conn.Execute("Select " + """" + "WATERCOMPONENTSSUBTYPES" + """" + "." + """" + "EPAREF" + """" + "," + """" + "WATERCOMPONENTSDATA" + """" + "." + """" + "VALUE_" + """" + "," + """" + " WATERCOMPONENTSSELECTIONS" + """" + "." + """" + "DESCRIPTION_" + """" + " from " + """" + "WATERCOMPONENTSDATA" + """" & _
                         "INNER JOIN " + """" + "WATERCOMPONENTSSUBTYPES" + """" + " on " + """" + "WATERCOMPONENTSSUBTYPES" + """" + "." + """" + "ID_SUBTYPE" + """" + "=" + """" + "WATERCOMPONENTSDATA" + """" + "." + """" + "ID_SUBTYPE" + """" + " and " + """" + "WATERCOMPONENTSSUBTYPES" + """" + "." + """" + "ID_TYPE" + """" + "=" + """" + "WATERCOMPONENTSDATA" + """" + "." + """" + "ID_TYPE" + """" & _
                         "LEFT JOIN " + """" + "WATERCOMPONENTSSELECTIONS" + """" + " on " + """" + "WATERCOMPONENTSSELECTIONS" + """" + "." + """" + "ID_SUBTYPE" + """" + "=" + """" + "WATERCOMPONENTSDATA" + """" + "." + """" + "ID_SUBTYPE" + """" + " and " + """" + "WATERCOMPONENTSSELECTIONS" + """" + "." + """" + "ID_TYPE" + """" + "=" + """" + "WATERCOMPONENTSDATA" + """" + "." + """" + "ID_TYPE" + """" + " and cast(" + """" + "WATERCOMPONENTSDATA" + """" + "." + """" + "VALUE_" + """" + " as INT) =" + """" + "WATERCOMPONENTSSELECTIONS" + """" + "." + """" + "VALUE_" + """" & _
                         "where " + """" + "OBJECT_ID_" + """" + " = '" & id & "'")
   
                         
                     
                         
                         
                         
                         End If
   
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
Sub AddSubItemPumps(id As String)
   Dim Rs As ADODB.Recordset
   Dim CURVE As String
    If conn.Provider <> "PostgreSQL.1" Then
   Set Rs = conn.Execute("Select b.eparef, w.value_, s.description_ from watercomponentsdata w " & _
                         "INNER JOIN watercomponentssubtypes b on b.id_subtype=w.id_subtype and b.id_type=w.id_type " & _
                         "LEFT JOIN WaterComponentsSelections s on s.id_subtype=w.id_subtype and s.id_type=w.id_type and cast(w.value_ as INT)=s.value_ " & _
                         "where object_id_ = '" & id & "'")
                         
                         Else
                         
                        Set Rs = conn.Execute("Select " + """" + "WATERCOMPONENTSSUBTYPES" + """" + "." + """" + "EPAREF" + """" + "," + """" + "WATERCOMPONENTSDATA" + """" + "." + """" + "VALUE_" + """" + "," + """" + " WATERCOMPONENTSSELECTIONS" + """" + "." + """" + "DESCRIPTION_" + """" + " from " + """" + "WATERCOMPONENTSDATA" + """" & _
                         "INNER JOIN " + """" + "WATERCOMPONENTSSUBTYPES" + """" + " on " + """" + "WATERCOMPONENTSSUBTYPES" + """" + "." + """" + "ID_SUBTYPE" + """" + "=" + """" + "WATERCOMPONENTSDATA" + """" + "." + """" + "ID_SUBTYPE" + """" + " and " + """" + "WATERCOMPONENTSSUBTYPES" + """" + "." + """" + "ID_TYPE" + """" + "=" + """" + "WATERCOMPONENTSDATA" + """" + "." + """" + "ID_TYPE" + """" & _
                         "LEFT JOIN " + """" + "WATERCOMPONENTSSELECTIONS" + """" + " on " + """" + "WATERCOMPONENTSSELECTIONS" + """" + "." + """" + "ID_SUBTYPE" + """" + "=" + """" + "WATERCOMPONENTSDATA" + """" + "." + """" + "ID_SUBTYPE" + """" + " and " + """" + "WATERCOMPONENTSSELECTIONS" + """" + "." + """" + "ID_TYPE" + """" + "=" + """" + "WATERCOMPONENTSDATA" + """" + "." + """" + "ID_TYPE" + """" + " and cast(" + """" + "WATERCOMPONENTSDATA" + """" + "." + """" + "VALUE_" + """" + " as INT) =" + """" + "WATERCOMPONENTSSELECTIONS" + """" + "." + """" + "VALUE_" + """" & _
                         "where " + """" + "OBJECT_ID_" + """" + " = '" & id & "'")
                         
                         
                         End If
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


Sub AddSubItemValves(id As String)
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

Function TrechoCadastrado(id As String) As Boolean

   'id = número do nó
   'Verifica se o trecho já foi cadastrado
   rsTrechosExportados.Filter = "id='" & id & "'"
   'retorna verdadeiro se o trecho se o trecho ja foi exportado, falso se não
   TrechoCadastrado = Not rsTrechosExportados.EOF
End Function

Function NoCadastrada(Object_id_ As String) As Boolean
   'Verifica se o nó já foi cadastrado
   rsNosExportados.Filter = "id='" & Object_id_ & "'"
   NoCadastrada = Not rsNosExportados.EOF
End Function
'Gera um vetor temporário de nós com seus atributos, como o objetivo de facilitar a leitura dos dados dos nós da rede
'Cria uma cópia da query da tabela watercomponents + points para RsNosTMP com todos os nos das tabelas relacionadas
'
'arquivoLog - nome do arquivo de logo onde está sendo exportado o log do Epanet
'
Sub CarregaRsNosTMP(arquivoLog As String)
On Error GoTo Trata_Erro
    Dim layer_id As Long
    Dim strSQL As String
    
    Open arquivoLog For Append As #5                                'continua a realizar o log do sistema
    Print #5, vbCrLf & "CarregaRsNosTMP; Inicia o registro do resultado dos nós temporários"
    Set rsNos = New ADODB.Recordset
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
    If conn.Provider <> "PostgreSQL.1" Then
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
    Else
        ' layer_id = Trim(str(GetLayerID("WATERCOMPONENTS")))
        Dim aaz As String
        aaz = Trim(str(GetLayerID("WATERCOMPONENTS")))
        strSQL = " Select A" + "." + """" + "OBJECT_ID_" + """" + ","
        strSQL = strSQL + """" + "x" + """" + " ," + """" + "y" + """" + "," + """" + "ID_TYPE" + """" + "," + """" + "INITIALGROUNDHEIGHT" + """" + "," + """" + "DEMAND" + """" + "," + """" + "PATTERN" + """" + "," + """" + "SUBTYPEVALVE" + """" + "," + "case when " + """" + "STATE" + """" + " = '2' then 'Closed' else 'Open' end " + """" + "STATE" + """" + ""
        strSQL = strSQL + "        From"
        strSQL = strSQL + " (Select " + """" + "OBJECT_ID_" + """" + ","
        strSQL = strSQL + """" + "x" + """" + " ," + """" + "y" + """" + "," + """" + "ID_TYPE" + """" + "," + """" + "INITIALGROUNDHEIGHT" + """" + "," + """" + "DEMAND" + """" + "," + """" + "PATTERN" + """" + ""
        strSQL = strSQL + "      from " + """" + "WATERCOMPONENTS" + """" + " inner join " + """" + "points" & aaz + """" + " on " + """" + "WATERCOMPONENTS" + """" + "." + """" + "OBJECT_ID_" + """" + "=" + """" + "points" + aaz + """" + "." + """" + "object_id" + """" + ")A"
        strSQL = strSQL + " Left Join"
        strSQL = strSQL + " (select " + """" + "OBJECT_ID_" + """" + "," + """" + "VALUE_" + """" + " as " + """" + "SUBTYPEVALVE" + """" + " from " + """" + "WATERCOMPONENTSDATA" + """" + "  where " + """" + "ID_TYPE" + """" + " = '1' and " + """" + "ID_SUBTYPE" + """" + " = '1')B"
        strSQL = strSQL + " on  A" + "." + """" + "OBJECT_ID_" + """" + "=B" + "." + """" + "OBJECT_ID_" + """" + ""
        strSQL = strSQL + "  left Join (select " + """" + "OBJECT_ID_" + """" + "," + """" + "VALUE_" + """" + " as " + """" + "STATE" + """" + " from " + """" + "WATERCOMPONENTSDATA" + """" + "  where " + """" + "ID_TYPE" + """" + " = '1' and " + """" + "ID_SUBTYPE" + """" + " = '2')C"
        strSQL = strSQL + "  ON A" + "." + """" + "OBJECT_ID_" + """" + "=C" + "." + """" + "OBJECT_ID_" + """" + ""
        'MsgBox "ARQUIVO DEBUG SALVO"
        'WritePrivateProfileString "A", "A", strSQL, App.Path & "\DEBUG.INI"
    End If
    Print #5, "CarregaRsNosTMP; string conexão: " & strSQL & " conexão: " & conn
    rsNos.Open strSQL, conn
    If conn.Provider <> "PostgreSQL.1" Then
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
    Else
        While Not rsNos.EOF
            With rsNosTmp
                .AddNew
                .Fields("ID").Value = rsNos.Fields("OBJECT_ID_").Value
                .Fields("X").Value = rsNos.Fields("x").Value
                .Fields("Y").Value = rsNos.Fields("y").Value
                .Fields("Tipo").Value = IIf(IsNull(rsNos.Fields("ID_TYPE").Value), 0, rsNos.Fields("ID_TYPE").Value)
                If rsNos.Fields("ID_TYPE").Value = No_Valvulas Then
                    Select Case rsNos.Fields("SUBTYPEVALVE").Value
                        Case 4, 0
                            .Fields("Tipo").Value = 1
                        Case Else
                            .Fields("Tipo").Value = 99
                    End Select
                End If
                .Fields("Cota").Value = IIf(IsNull(rsNos.Fields("INITIALGROUNDHEIGHT").Value), 0, rsNos.Fields("INITIALGROUNDHEIGHT").Value)
                .Fields("Demanda").Value = IIf(IsNull(rsNos.Fields("DEMAND").Value), 0, rsNos.Fields("DEMAND").Value)
                .Fields("Padrao").Value = IIf(IsNull(rsNos.Fields("PATTERN").Value), 0, rsNos.Fields("PATTERN").Value)
                .Fields("estado").Value = rsNos.Fields("state").Value
            End With
            rsNos.MoveNext
        Wend
    End If
    rsNos.Close
    
    Set rsNos = Nothing
    'AO FINAL DESTA ROTINA FICARÁ EXISTENTE A TABELA DE NÓS COMPLETA DESNORMATIZADA (rsNosTmp)
    Close #5
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
End Sub
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
    
    'lista de nós exportados
    rsNosExportados.Fields.Append "id", adVarChar, 255
    rsNosExportados.Open
End Sub
'
'
'
Sub GeraArquivo_de_Saida()
On Error GoTo Trata_Erro
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

    
    Dim A As Long
    Dim str As String
    
    intLinhaCod = 1
    Open FrmEPANET.txtArquivo.Text For Output As #1
       intLinhaCod = 2
       'grava no arquivo as Junctions


   Print #1, "[JUNCTIONS]"
   'CARREGA EM STR O CABEÇALHO
   str = "ID" & Chr(vbKeyTab) & Chr(vbKeyTab)
   str = str & "ELEV" & Chr(vbKeyTab) & Chr(vbKeyTab)
   str = str & "DEMAND" & Chr(vbKeyTab) & Chr(vbKeyTab)
   str = str & "PATTERN" & Chr(vbKeyTab) & Chr(vbKeyTab)
   Print #1, ";" & str
   str = ""

   ' "id", adVarChar, 255
   ' "elev", adVarChar, 255
   ' "demand", adDouble, 255
   ' "pattern", adVarChar, 255


Dim Cota As String
Dim pos As Integer


   With rsJunctions
     .Filter = ""
      While Not .EOF
         If IsNumeric(.Fields("ID").Value) = True Then
              str = .Fields("ID").Value & Chr(vbKeyTab) & Chr(vbKeyTab)
              
              Cota = Replace(.Fields("ELEV").Value, ",", ".")  ' recebe o valor do banco troca a virgula por ponto
              pos = InStr(1, Cota, ".", vbBinaryCompare) + 1   ' localiza a posição do ponto na string e adiciona 1
              Cota = Mid(Cota, 1, pos)                         ' pega 1 casa apos a virgula
              
              str = str & Cota & Chr(vbKeyTab) & Chr(vbKeyTab)
              
              str = str & Replace(.Fields("DEMAND").Value, ",", ".") & Chr(vbKeyTab) & Chr(vbKeyTab)
              str = str & Replace(.Fields("PATTERN").Value, ",", ".") & Chr(vbKeyTab) & Chr(vbKeyTab)
         Else
              str = .Fields("ID").Value & Chr(vbKeyTab) & Chr(vbKeyTab)
              str = str & Replace(.Fields("ELEV").Value, ",", ".") & Chr(vbKeyTab) & Chr(vbKeyTab)
              str = str & "0" & Chr(vbKeyTab) & Chr(vbKeyTab)
              str = str & Replace(.Fields("PATTERN").Value, ",", ".") & Chr(vbKeyTab) & Chr(vbKeyTab)
         End If
         Print #1, str & ";"
         str = ""
         .MoveNext
      Wend
   End With


'ORIGINAL
'       With rsJunctions
'          .Filter = ""
'          intLinhaCod = 3
'          If .RecordCount > 0 Then
'             .MoveFirst
'             intLinhaCod = 4
'             For A = 0 To .Fields.Count - 1
'                str = str & .Fields(A).Name & Chr(vbKeyTab) & Chr(vbKeyTab)
'             Next
'             intLinhaCod = 5
'             Print #1, "[JUNCTIONS]"
'             Print #1, ";" & str
'             str = ""
'             intLinhaCod = 6
'             While Not .EOF
'                For A = 0 To .Fields.Count - 1
'                   str = str & IIf(IsNumeric(.Fields(A).Value), _
'                            Replace(.Fields(A).Value, ",", "."), _
'                            .Fields(A).Value) & Chr(vbKeyTab) & Chr(vbKeyTab)
'                Next
'                Print #1, str & ";"
'                str = ""
'                .MoveNext
'             Wend
'             intLinhaCod = 7
'          End If
'       End With
       
       'grava no arquivo as Reservoirs
       With rsReservoirs
          .Filter = ""
          intLinhaCod = 8
          If .RecordCount > 0 Then
             .MoveFirst
             intLinhaCod = 9
             For A = 0 To .Fields.Count - 1
                str = str & .Fields(A).Name & Chr(vbKeyTab) & Chr(vbKeyTab)
             Next
             Print #1, ""
             Print #1, "[RESERVOIRS]"
             Print #1, ";" & str
             str = ""
             intLinhaCod = 10
             While Not .EOF
                For A = 0 To .Fields.Count - 1
                   str = str & IIf(IsNumeric(.Fields(A).Value), _
                            Replace(.Fields(A).Value, ",", "."), _
                            .Fields(A).Value) & Chr(vbKeyTab) & Chr(vbKeyTab)
                Next
                Print #1, str & ";"
                str = ""
                .MoveNext
             Wend
            intLinhaCod = 11
          End If
       End With
       
       'grava no arquivo as Tanks
       With rsTanks
          .Filter = ""
          intLinhaCod = 12
          If .RecordCount > 0 Then
             .MoveFirst
             intLinhaCod = 13
             For A = 0 To .Fields.Count - 1
                str = str & .Fields(A).Name & Chr(vbKeyTab) & Chr(vbKeyTab)
             Next
             Print #1, ""
             Print #1, "[TANKS]"
             Print #1, ";" & str
             str = ""
             intLinhaCod = 14
             While Not .EOF
                For A = 0 To .Fields.Count - 1
                   str = str & IIf(IsNumeric(.Fields(A).Value), _
                            Replace(.Fields(A).Value, ",", "."), _
                            .Fields(A).Value) & Chr(vbKeyTab) & Chr(vbKeyTab)
                Next
                Print #1, str & ";"
                str = ""
                .MoveNext
             Wend
             intLinhaCod = 15
          End If
       End With
       
       'grava no arquivo as Pumps
       With rsPumps
          .Filter = ""
          intLinhaCod = 16
          If .RecordCount > 0 Then
             .MoveFirst
             intLinhaCod = 17
             For A = 0 To .Fields.Count - 1
                str = str & .Fields(A).Name & Chr(vbKeyTab) & Chr(vbKeyTab)
             Next
             Print #1, ""
             Print #1, "[PUMPS]"
             Print #1, ";" & str
             str = ""
             intLinhaCod = 18
             While Not .EOF
                For A = 0 To .Fields.Count - 1
                   str = str & IIf(IsNumeric(.Fields(A).Value), _
                            Replace(.Fields(A).Value, ",", "."), _
                            .Fields(A).Value) & Chr(vbKeyTab) & Chr(vbKeyTab)
                Next
                Print #1, str & ";"
                str = ""
                .MoveNext
             Wend
             intLinhaCod = 19
          End If
       End With
       
       'grava no arquivo as Valves
       With rsValves
          .Filter = ""
          If .RecordCount > 0 Then
             .MoveFirst
             intLinhaCod = 20
             For A = 0 To .Fields.Count - 1
                str = str & .Fields(A).Name & Chr(vbKeyTab) & Chr(vbKeyTab)
             Next
             Print #1, ""
             Print #1, "[VALVES]"
             Print #1, ";" & str
             str = ""
             intLinhaCod = 21
             While Not .EOF
                For A = 0 To .Fields.Count - 1
                   str = str & IIf(IsNumeric(.Fields(A).Value), _
                            Replace(.Fields(A).Value, ",", "."), _
                            .Fields(A).Value) & Chr(vbKeyTab) & Chr(vbKeyTab)
                Next
                Print #1, str & ";"
                str = ""
                .MoveNext
             Wend
             intLinhaCod = 22
          End If
       End With
       
       'grava no arquivo as Pipes
       
   ' "id", adVarChar, 255
   ' "node1", adVarChar, 255
   ' "node2", adVarChar, 255
   ' "length", adVarChar, 255
   ' "diameter", adDouble, 255
   ' "roughness", adDouble, 255
   ' "minorloss", adVarChar, 255
   ' "status", adVarChar, 255
       
      
         
'      Print #1, ""
'      Print #1, "[PIPES]"
'      'CARREGA EM STR O CABEÇALHO
'      str = "ID" & Chr(vbKeyTab) & Chr(vbKeyTab)
'      str = str & "NODE1" & Chr(vbKeyTab) & Chr(vbKeyTab)
'      str = str & "NODE2" & Chr(vbKeyTab) & Chr(vbKeyTab)
'      str = str & "LENGTH" & Chr(vbKeyTab) & Chr(vbKeyTab)
'      str = str & "DIAMETER" & Chr(vbKeyTab) & Chr(vbKeyTab)
'      str = str & "ROUGHNESS" & Chr(vbKeyTab) & Chr(vbKeyTab)
'      str = str & "MINORLOSS" & Chr(vbKeyTab) & Chr(vbKeyTab)
'      str = str & "STATUS" & Chr(vbKeyTab) & Chr(vbKeyTab)
'
'      Print #1, ";" & str
'      str = ""
'
'      With rsPipes
'          .Filter = ""
'          If .RecordCount > 0 Then
'             .MoveFirst
'
'             While Not .EOF
'
'                  If IsNumeric(.Fields("ID").Value) = True Then
'
'                     str = .Fields("ID").Value & Chr(vbKeyTab) & Chr(vbKeyTab)
'                     str = .Fields("ID").Value & Chr(vbKeyTab) & Chr(vbKeyTab)
'                     str = .Fields("ID").Value & Chr(vbKeyTab) & Chr(vbKeyTab)
'                     str = .Fields("ID").Value & Chr(vbKeyTab) & Chr(vbKeyTab)
'                     str = .Fields("ID").Value & Chr(vbKeyTab) & Chr(vbKeyTab)
'                     str = .Fields("ID").Value & Chr(vbKeyTab) & Chr(vbKeyTab)
'                     str = .Fields("ID").Value & Chr(vbKeyTab) & Chr(vbKeyTab)
'                     str = .Fields("ID").Value & Chr(vbKeyTab) & Chr(vbKeyTab)
'                     str = .Fields("ID").Value & Chr(vbKeyTab) & Chr(vbKeyTab)
'                     str = .Fields("ID").Value & Chr(vbKeyTab) & Chr(vbKeyTab)
'
'                     str = str & "NODE1" & Chr(vbKeyTab) & Chr(vbKeyTab)
'                     str = str & "NODE2" & Chr(vbKeyTab) & Chr(vbKeyTab)
'                     str = str & "LENGTH" & Chr(vbKeyTab) & Chr(vbKeyTab)
'                     str = str & "DIAMETER" & Chr(vbKeyTab) & Chr(vbKeyTab)
'                     str = str & "ROUGHNESS" & Chr(vbKeyTab) & Chr(vbKeyTab)
'                     str = str & "MINORLOSS" & Chr(vbKeyTab) & Chr(vbKeyTab)
'                     str = str & "STATUS" & Chr(vbKeyTab) & Chr(vbKeyTab)
'
'
'                Else 'SE O ID NÃO É NUMÉRICO, IMPRIME TUDO MAS ZERA A DEMANDA
'                     str = "ID" & Chr(vbKeyTab) & Chr(vbKeyTab)
'                     str = str & "NODE1" & Chr(vbKeyTab) & Chr(vbKeyTab)
'                     str = str & "NODE2" & Chr(vbKeyTab) & Chr(vbKeyTab)
'                     str = str & "LENGTH" & Chr(vbKeyTab) & Chr(vbKeyTab)
'                     str = str & "DIAMETER" & Chr(vbKeyTab) & Chr(vbKeyTab)
'                     str = str & "ROUGHNESS" & Chr(vbKeyTab) & Chr(vbKeyTab)
'                     str = str & "MINORLOSS" & Chr(vbKeyTab) & Chr(vbKeyTab)
'                     str = str & "STATUS" & Chr(vbKeyTab) & Chr(vbKeyTab)
'
'                End If
'
'                For A = 0 To .Fields.Count - 1 '
'                   str = str & IIf(IsNumeric(.Fields(A).Value), _
'                            Replace(.Fields(A).Value, ",", "."), _
'                            .Fields(A).Value) & Chr(vbKeyTab) & Chr(vbKeyTab)
'                  If A = 7 Then
'                     str = str & ";" & .Fields(8).Value
'                     Exit For
'                  End If
'
'                Next
'                Print #1, str
'                'Print #1, str & ";"
'                str = ""
'                .MoveNext
'             Wend
'             intLinhaCod = 25
'          End If
'      End With
       
       
       With rsPipes
          .Filter = ""
          intLinhaCod = 23
          If .RecordCount > 0 Then
             .MoveFirst
             intLinhaCod = 24
             For A = 0 To .Fields.Count - 2 'NÃO IMPRIME O NOME DO CAMPO DESCRIPTION
                str = str & .Fields(A).Name & Chr(vbKeyTab) & Chr(vbKeyTab)
             Next
             Print #1, ""
             Print #1, "[PIPES]"
             Print #1, ";" & str
             str = ""
             While Not .EOF
                For A = 0 To .Fields.Count - 1 '
                   str = str & IIf(IsNumeric(.Fields(A).Value), _
                            Replace(.Fields(A).Value, ",", "."), _
                            .Fields(A).Value) & Chr(vbKeyTab) & Chr(vbKeyTab)
                  If A = 7 Then
                     str = str & ";" & .Fields(8).Value
                     Exit For
                  End If

                Next
                Print #1, str
                'Print #1, str & ";"
                str = ""
                .MoveNext
             Wend
             intLinhaCod = 25
          End If
       End With
       
       intLinhaCod = 26
       Dim MyArray() As String
       Dim rsPatterns As ADODB.Recordset
        If conn.Provider <> "PostgreSQL.1" Then
       Set rsPatterns = conn.Execute("Select * from x_patterns")
       Else
        Set rsPatterns = conn.Execute("Select * from " + """" + "X_PATTERNS" + """" + "")
       
       End If
       'grava no arquivo as Patterns
       intLinhaCod = 26
       
       If rsPatterns.EOF = False Then
       
             With rsPatterns
                Print #1, "[PATTERNS]"
                Print #1, ";ID" & Chr(vbKeyTab) & Chr(vbKeyTab) & "Multipliers"
                Print #1, ";" & rsPatterns("descricao").Value
                intLinhaCod = 27
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
                intLinhaCod = 28
             End With
       End If
       
       intLinhaCod = 29
       rsPatterns.Close
       Set rsPatterns = Nothing
       Dim b As Integer
       Dim MyArray_x() As String
       Dim MyArray_y() As String
       Dim rsCurves As ADODB.Recordset
       If conn.Provider <> "PostgreSQL.1" Then
       Set rsCurves = conn.Execute("Select * from x_Curves order by tipo")
       Else
       Set rsCurves = conn.Execute("Select * from " + """" + "X_CURVES" + """" + "")
       End If
       intLinhaCod = 30
       
       'grava no arquivo as Curves
       If rsCurves.EOF = False Then
             intLinhaCod = 31
             With rsCurves
                
                Print #1, "[CURVES]"
                Print #1, ";ID" & Chr(vbKeyTab) & Chr(vbKeyTab) & "X-Value" & Chr(vbKeyTab) & Chr(vbKeyTab) & "Y-Value"
                intLinhaCod = 32
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
                 intLinhaCod = 34
             End With
       End If
       rsCurves.Close
       Set rsCurves = Nothing
       
       'grava no arquivo as Coordinates
        Open App.Path & "\LogErroExportEPANET-histórico.txt" For Append As #4
       intLinhaCod = 35
       With rsCoordinates
          .Filter = ""
          If .RecordCount > 0 Then
             .MoveFirst
             For A = 0 To .Fields.Count - 1
                str = str & .Fields(A).Name & Chr(vbKeyTab) & Chr(vbKeyTab)
             Next
             intLinhaCod = 36
             Print #1, ""
             Print #1, "[COORDINATES]"
             Print #1, ";" & str
             str = ""
             intLinhaCod = 37
             While Not .EOF
                For A = 0 To .Fields.Count - 1
                   str = str & IIf(IsNumeric(.Fields(A).Value), _
                            Replace(.Fields(A).Value, ",", "."), _
                            .Fields(A).Value) & Chr(vbKeyTab) & Chr(vbKeyTab)
                Next
                Print #1, str & ";"
                Print #4, str & ";"
                str = ""
                .MoveNext
             Wend
             intLinhaCod = 38
          End If
       End With
       Close #4
       'grava no arquivo as Vertices
       intLinhaCod = 39
       With rsVertices
          .Filter = ""
          intLinhaCod = 40
          If .RecordCount > 0 Then
             .MoveFirst
             intLinhaCod = 41
             For A = 0 To .Fields.Count - 1
                str = str & .Fields(A).Name & Chr(vbKeyTab) & Chr(vbKeyTab)
             Next
             intLinhaCod = 42
             Print #1, ""
             Print #1, "[VERTICES]"
             Print #1, ";" & str
             str = ""
             intLinhaCod = 43
             While Not .EOF
                For A = 0 To .Fields.Count - 1
                   str = str & IIf(IsNumeric(.Fields(A).Value), _
                            Replace(.Fields(A).Value, ",", "."), _
                            .Fields(A).Value) & Chr(vbKeyTab) & Chr(vbKeyTab)
                Next
                Print #1, str & ";"
                str = ""
                .MoveNext
             Wend
             intLinhaCod = 44
          End If
       End With
    Close #1

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

End Sub

Public Function GetLayerID(LayerName_ As String) As Integer
   Dim Rs As ADODB.Recordset
    If conn.Provider <> "PostgreSQL.1" Then
    Set Rs = conn.Execute("SELECT LAYER_ID FROM TE_LAYER WHERE UPPER(name) ='" & UCase(LayerName_) & "'")
    Else
    
    Set Rs = conn.Execute("Select " + """" + "layer_id" + """" + " from " + """" + "te_layer" + """" + " where " + """" + "name" + """" + "='" & LayerName_ & "'")
    End If
    
    If Rs.EOF = False Then
        GetLayerID = Rs(0).Value
    Else
        MsgBox "Não Localizado o Layer " & UCase(LayerName_)
        End
    End If
    Rs.Close
    Set Rs = Nothing
End Function

Public Function FunDecripta(ByVal strDecripta As String) As String


    Dim IntTam As Integer
    Dim i As Integer
    Dim letra, nStr As String
    IntTam = Len(strDecripta)
    nStr = ""

    'desconsidera os os numeros de HH-MM-SS
    strDecripta = Mid(strDecripta, 6, 5) & Mid(strDecripta, 16, 5) & Mid(strDecripta, 26, 5) & _
                  Mid(strDecripta, 36, 5) & Mid(strDecripta, 46, 5) & Mid(strDecripta, 56, 200)

    i = 1
    Do While Not i = IntTam - 29
        letra = Mid(strDecripta, i, 5)
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

