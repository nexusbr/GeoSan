VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTerraLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String
Dim f As String
Dim g As String
Dim h As String
Dim i As String
Dim j As String
Dim k As String
Dim l As String
Dim m As String
Public Enum TypeGeometryEvent
   tg_NoEvent = 0
   tg_DrawNetWorkline = 1
   tg_DrawNetWorkNode = 2
   tg_MoveNetWorkNode = 3
   tg_SelectObject = 4
   tg_DrawGeometrys = 5
   tg_DrawPoint = 6
   tg_ZoomArea = 7
   tg_Pan = 8
   tg_DrawRamal = 9
   tg_InsertDoc = 10
   tg_MoveNetWorkVertice = 11
   tg_MoveGpsPoint = 12
   tg_DrawRamalAuto = 13                                        'para selecionar o trecho de rede ao qual os ramais serão desenhados
   tg_DrawRamalAutoSelecionaConsumidor = 14                     'para selecionar o consumidor que será associado ao ramal que será desenhado
   tg_DrawRamalAddConsumer = 15                                 'para selecionar o ramal ao qual serão ligados os consumidores
   tg_DrawRamalAddConsumerSelecionaConsumidor = 16              'para selecionar os consumidores quer irão ligar-se aos ramais
End Enum

Private LayerReference As String, rs As ADODB.Recordset
Private mvartcs As TeCanvas 'local copy
Private mvarCtrlMgr As Manager 'local copy
Private mvarCGeo As clsGeoReference 'local copy
Private mvarTerraEvent As TypeGeometryEvent 'local copy

Private mvartdb As TeDatabase 'local copy
Private mvartdbcon As TeDatabase 'local copy
Private mvartdbconref As TeDatabase 'local copy
' Recebe o objeto da conexão TeDatabase
'
Public Property Set tdb(ByVal vData As TeDatabase)
    Set mvartdb = vData
    Set cgeo.tdb = vData
End Property
' Retorna o objeto da conexão TeDatabase
'
Public Property Get tdb() As TeDatabase
    Set tdb = mvartdb
End Property

Public Property Set tdbcon(ByVal vData As TeDatabase)
    Set mvartdbcon = vData
    Set cgeo.tdbcon = vData
End Property

Public Property Get tdbcon() As TeDatabase
    Set tdbcon = mvartdbcon
End Property

Public Property Set tdbconref(ByVal vData As TeDatabase)
    Set mvartdbconref = vData
End Property

Public Property Get tdbconref() As TeDatabase
    Set tdbconref = mvartdbconref
End Property


Public Property Let TerraEvent(ByVal vData As TypeGeometryEvent)
    mvarTerraEvent = vData
End Property

Public Property Get TerraEvent() As TypeGeometryEvent
    TerraEvent = mvarTerraEvent
End Property

Public Property Set cgeo(ByVal vData As clsGeoReference)
    Set mvarCGeo = vData
End Property

Public Property Get cgeo() As clsGeoReference
    Set cgeo = mvarCGeo
End Property

Public Property Set CtrlMgr(ByVal vData As Manager)
    Set mvarCtrlMgr = vData
End Property

Public Property Get CtrlMgr() As Manager
    Set CtrlMgr = mvarCtrlMgr
End Property

Public Property Set tcs(ByVal vData As TeCanvas)
    Set mvartcs = vData
    
End Property

Public Property Get tcs() As TeCanvas
    Set tcs = mvartcs
End Property
' Classe responsável pelo desenho da rede. Aquí é onde é ativado o início do desenho de uma rede
' Retorna verdadeiro se foi ativada corretamente
'
'
Public Function DrawNetWorkLine(Optional mback As Boolean) As Boolean
    On Error GoTo Trata_Erro
    With tcs                                                             'com o objeto teCanvas
        LayerReference = cgeo.GetLayerOperation(.getCurrentLayer, 0)      'obtem o nome do layer de referencia, por exemplo se o principal é waterlines o de referência é watercomponents, sewerlines é sewercomponents, etc.
        If Trim(LayerReference) <> "" Then
            tdbcon.setCurrentLayer LayerReference                          'define como ativo o layer de nós
        Else
            MsgBox "Selecione antes o plano em que deseja desenhar!", vbInformation, "Desenho de Rede"
            DrawNetWorkLine = False
            Exit Function
        End If
        'tdbconref.setCurrentLayer "CEL_PONTO_ATRIB"
        Select Case cgeo.GetLayerTypeReference(.getCurrentLayer)                      'obtem o tipo de layer de referência
            Case LayerTypeRefence.Trecho_Rede_Agua, LayerTypeRefence.Trecho_Rede_Drenagem, LayerTypeRefence.Trecho_Rede_esgoto
                If Not mback Then
                    .object_id = ""
                    .addLayerToSnap cgeo.GetReferenceLayer(tcs.getCurrentLayer)             'obtem o layer de referência e coloca-o como de snap
                    .setReferenceLayer cgeo.GetReferenceLayer(tcs.getCurrentLayer)          'seta o nome do layer de referência
                    CtrlMgr.LoadDefaultProperties 0, tcs.getCurrentLayer, True              'carrega os valores padrão do gerenciador de propriedades
                    CtrlMgr.GridEnabled True: CtrlMgr.GridVisibled True
                End If
                .DrawNetWorkLine
                TerraEvent = tg_DrawNetWorkline
                DrawNetWorkLine = True
            Case Else
                DrawNetWorkLine = False
                MsgBox "Não é possivel desenhar rede neste plano: " & .getCurrentLayer, vbExclamation
                Exit Function
        End Select
    End With
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        PrintErro "clsTerraLib", "Public Function DrawNetWorkLine", CStr(Err.Number), CStr(Err.Description), True
    End If
End Function
' Entra neste evento logo após selecionado um nó da rede para ser movido para uma nova posição.
' Ele checa se foi selecionado mais de um nó pois é permitido mover somente um nó por vez
' Se estiver tudo OK então adiciona o layer de snap, referência e ativa o comendo de mover nó de uma rede
' Entra neste evento antes de salvar no banco, com os dados originais
'
'
Public Function MoveNetWorkNode()
    On Error GoTo Trata_Erro
    Dim layerReferencia As String                                       'nome do layer de referência
    
    layerReferencia = cgeo.GetReferenceLayer(tcs.getCurrentLayer)
    If layerReferencia <> "" Then
        If tcs.getSelectCount(points) - 1 = 0 Then                      'VERIFICA SE SOMENTE 1 PONTO FOI SELECIONADO PARA MOVER
            Dim objIdNoMovendo As Long                                  'object_id do nó que foi selecionado para mover
            Dim trechos As New CLocalizaTrechos                         'instancia objeto que irá localizar os trechos conectados ao nó selecionado
            Dim objIDsTrechosMovendo() As String                        'lista com todos os object_ids dos trechos que serão movidos com o nó (presos no nó)
            Dim contTrecho As Integer                                   'contador de trechos de rede sendo movidos
            Dim listObjIDsRamais() As String                            'lista de todos objIDs dos ramais de um determinado trecho
            Dim listGeomIDsRamais() As String                           'lista de todos os geomIDs dos ramais de um determinado trecho
            Dim contRamal As Integer                                    'contador de ramais conectados em um determinado trecho de rede
            Dim distIniRamalAntes As Double                             'distância do início do ramal antes de tanto o trecho quanto o ramal serem movidos
            Dim distRamalDoInicioTrecho As New CDistRamalInicioTrecho   'clases para calcular a distância do início do ramal antes de tanto o trecho quanto o ramal serem movidos
            Dim objIDsRamais As New CObtemObjIDsRamais
            Dim coordRamal As New CCoordIniFinLinha                     'classe para obter as coordenadas inicial e final do ramal que foi movido
            Dim objIdTrecho As String
            Dim objIdRamal As String
            Dim geomIdRamal As String
            Dim contador As Integer                                     'número de ramais processados
            Dim comprTrecho As Double                                   'comprimento total do trecho de rede antes de mover
            Dim retorno As Boolean
            
            If layerReferencia = "WATERLINES" Then                      'é água, tem que tratar todos os ramais
            'nesta etapa ele vai localizar todos os trechos que serão movidos e os respectivos ramais se houverem e armazenar tudo na memória
                contador = 0
                objIdNoMovendo = tcs.getSelectObjectId(0, 4)                        'obtem o object_id da primeira (0) (e única) geometria do tipo 4=ponto (ver tipos em setActiveGeometry)
                varGlobais.objIdNoSelecionado = objIdNoMovendo                      'guarda o valor do objId do nó da rede que está movendo e foi selecionado, para depois poder atribir a cota ao mesmo
                trechos.obtemTrechos objIdNoMovendo, objIDsTrechosMovendo           'obtem todos os object_ids dos trechos de rede que estão conectados ao nó que está movendo
                varGlobais.objIdTre = Null                                          'zera a variável global, caso contenha algum valor lá
                varGlobais.objIdTre = objIDsTrechosMovendo                          'armazena na variável global todos os object_ids dos trechos de rede que estão conectados ao nó que está movendo
                varGlobais.totalTrechos = UBound(objIDsTrechosMovendo)              'armazena o número total de trechos de rede que estão sendo movidos
                For contTrecho = 0 To UBound(objIDsTrechosMovendo)                  'para cada trecho de rede que estiver sendo movido, começando em zero
                    objIdTrecho = objIDsTrechosMovendo(contTrecho)
                    objIDsRamais.getObjIDs objIdTrecho, listObjIDsRamais, listGeomIDsRamais     'obtem todos os objIDs dos ramais que estão ligados ao trecho de rede que está sendo movido
                    If listObjIDsRamais(0) <> "" Then                               'existem ramais conectados ao trecho
                        For contRamal = 0 To UBound(listObjIDsRamais)               'enquanto existirem ramais
                            contador = contador + 1
                            ReDim Preserve ramalMovendo(contador - 1)
                            objIdRamal = listObjIDsRamais(contRamal)
                            geomIdRamal = listGeomIDsRamais(contRamal)
                            distIniRamalAntes = distRamalDoInicioTrecho.Distancia(objIdTrecho, objIdRamal)  'obtem a distância do início do ramal antes de tanto o trecho quanto o ramal serem movidos
                            ramalMovendo(contador - 1).objIdTrecho = objIdTrecho
                            ramalMovendo(contador - 1).objIdRamal = objIdRamal
                            ramalMovendo(contador - 1).geomIdRamal = geomIdRamal
                            ramalMovendo(contador - 1).Distancia = distIniRamalAntes
                            coordRamal.GetStartEndPointsRamal objIdRamal                                'obtem as coordenadas inicial e final do ramal
                            ramalMovendo(contador - 1).xHidrom = coordRamal.linha.xf                    'pega as coordenadas de onde está o hidrômetro
                            ramalMovendo(contador - 1).yHidrom = coordRamal.linha.yf
                            retorno = cGeoDatabase.geoDatabase.setCurrentLayer("WaterLines")
                            retorno = cGeoDatabase.geoDatabase.getLengthOfLine(objIdTrecho, "", comprTrecho) 'está como geom_id deveria ser na segunda
                            ramalMovendo(contador - 1).comprTrecho = comprTrecho
                        Next                                                        'próximo ramal
                    Else                                                            'não existem ramais
                        contador = contador + 1
                        ReDim Preserve ramalMovendo(contador - 1)
                        ramalMovendo(contador - 1).objIdTrecho = objIdTrecho
                        ramalMovendo(contador - 1).objIdRamal = -1
                    End If
                Next                                                                'próximo trecho
            'final da etapa
            Else                                                                    'é esgoto, não trata ramais de esgoto
                objIdNoMovendo = tcs.getSelectObjectId(0, 4)                        'obtem o object_id da primeira (0) (e única) geometria do tipo 4=ponto (ver tipos em setActiveGeometry)
                varGlobais.objIdNoSelecionado = objIdNoMovendo                      'guarda o valor do objId do nó da rede que está movendo e foi selecionado, para depois poder atribir a cota ao mesmo
            End If
            tcs.addLayerToSnap cgeo.GetReferenceLayer(tcs.getCurrentLayer)
            tcs.setReferenceLayer cgeo.GetReferenceLayer(tcs.getCurrentLayer)
            tcs.MoveNetWorkNode: TerraEvent = tg_MoveNetWorkNode
        Else
            MsgBox "Selecione um ponto para esta operação", vbExclamation
       End If
    Else
        MsgBox "Selecione um plano válido e um ponto, para esta operação", vbExclamation
    End If
    Exit Function

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "clsTerralib", "MoveNetWorkNode", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
   
End Function
' Chamada quando o usuário seleciona para inserir um nó na rede
Public Function DrawNetWorkNode()
On Error GoTo Trata_Erro
   
   With tcs
      If cgeo.GetReferenceLayer(tcs.getCurrentLayer, True) <> "" Then
         If tcs.getSelectCount(lines) = 1 Then
            'CtrlMgr.LoadDefaultProperties 0, CGeo.GetReferenceLayer(.getCurrentLayer), True
            
            tcs.setReferenceLayer cgeo.GetReferenceLayer(.getCurrentLayer)
            
            tcs.addLayerToSnap .getCurrentLayer
            
            tcs.addLayerToSnap cgeo.GetReferenceLayer(.getCurrentLayer)
            
            tcs.insertNetWorkNode False: TerraEvent = tg_DrawNetWorkNode
         ElseIf tcs.getSelectCount(lines) < 1 Then
            MsgBox "Selecione uma rede, para esta operação, depois de um clique na linha", vbExclamation
         ElseIf tcs.getSelectCount(lines) > 1 Then
            MsgBox "Selecione apenas uma rede, para esta operação, depois de um clique na linha", vbExclamation
         End If
      Else
         MsgBox "Selecione um plano válido, para esta operação", vbExclamation
      End If
   End With
   Exit Function

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       ErroUsuario.Registra "clsTerralib", "DrawNetWorkNode", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Function
' Para mover um vértice de um trecho de rede de água. Entra nesta rotina após selecionar a rede
' O objetivo aqui é coletar os dados para que o sistema possa depois saber quais ramais necessitam
' ser recalculados em sua nova posição. São todos salvos em memória
'
'
'
Public Function moveVertice()
    Dim objIdSelecionado As String                                          'objectId do trecho que foi selecionado no qual será movido um ou mais vértices
    Dim contador As Integer                                                 'número de ramais processados
    Dim objIdTrecho As String                                               'objectId do trecho que foi selecionado no qual será movido, mesmo que o anterior, só para ficar igual o código
    Dim objIDsRamais As New CObtemObjIDsRamais
    Dim listObjIDsRamais() As String                                        'lista de todos objIDs dos ramais de um determinado trecho
    Dim listGeomIDsRamais() As String                                       'lista de todos os geomIDs dos ramais de um determinado trecho
    Dim contRamal As Integer                                                'contador de ramais conectados em um determinado trecho de rede
    Dim objIdRamal As String
    Dim geomIdRamal As String
    Dim distIniRamalAntes As Double                                         'distância do início do ramal antes de tanto o trecho quanto o ramal serem movidos
    Dim distRamalDoInicioTrecho As New CDistRamalInicioTrecho               'clases para calcular a distância do início do ramal antes de tanto o trecho quanto o ramal serem movidos
    Dim retorno As Boolean
    Dim coordRamal As New CCoordIniFinLinha                                 'classe para obter as coordenadas inicial e final do ramal que foi movido
    Dim comprTrecho As Double                                               'comprimento total do trecho de rede antes de mover

    If tcs.getCurrentLayer = "WATERLINES" Then
        objIdSelecionado = varGlobais.objIdTreSelecionado                   'para obter o object_id da linha que foi selecionada para mover o vértice da mesma. Esta variável foi configurada em  frmCanvas.TCanvas_onEndSELECT
        tcs.moveGeometryPoint
        'nesta etapa ele vai localizar todos os trechos que serão movidos e os respectivos ramais se houverem e armazenar tudo na memória
        contador = 0
        varGlobais.objIdTre = Null                                          'zera a variável global, caso contenha algum valor lá
        varGlobais.objIdTre = objIdSelecionado                              'só um revisar para ver se nao deveria ser vetor 'armazena na variável global todos os object_ids dos trechos de rede que estão conectados ao nó que está movendo
        varGlobais.totalTrechos = 1                                         'armazena o número total de trechos de rede que estão sendo movidos. Um só pois está movendo o vértice de um trecho
        objIdTrecho = objIdSelecionado
        objIDsRamais.getObjIDs objIdTrecho, listObjIDsRamais, listGeomIDsRamais     'obtem todos os objIDs dos ramais que estão ligados ao trecho de rede que está sendo movido
        If listObjIDsRamais(0) <> "" Then                                   'existem ramais conectados ao trecho
            For contRamal = 0 To UBound(listObjIDsRamais)                   'enquanto existirem ramais
                contador = contador + 1
                ReDim Preserve ramalMovendo(contador - 1)
                objIdRamal = listObjIDsRamais(contRamal)
                geomIdRamal = listGeomIDsRamais(contRamal)
                distIniRamalAntes = distRamalDoInicioTrecho.Distancia(objIdTrecho, objIdRamal)  'obtem a distância do início do ramal antes de tanto o trecho quanto o ramal serem movidos
                ramalMovendo(contador - 1).objIdTrecho = objIdTrecho
                ramalMovendo(contador - 1).objIdRamal = objIdRamal
                ramalMovendo(contador - 1).geomIdRamal = geomIdRamal
                ramalMovendo(contador - 1).Distancia = distIniRamalAntes
                coordRamal.GetStartEndPointsRamal objIdRamal                'obtem as coordenadas inicial e final do ramal
                ramalMovendo(contador - 1).xHidrom = coordRamal.linha.xf    'pega as coordenadas de onde está o hidrômetro
                ramalMovendo(contador - 1).yHidrom = coordRamal.linha.yf
                retorno = cGeoDatabase.geoDatabase.setCurrentLayer("WaterLines")
                retorno = cGeoDatabase.geoDatabase.getLengthOfLine(objIdTrecho, "", comprTrecho) 'está como geom_id deveria ser na segunda
                ramalMovendo(contador - 1).comprTrecho = comprTrecho
            Next                                                            'próximo ramal
        Else                                                                'não existem ramais
            contador = contador + 1
            ReDim Preserve ramalMovendo(contador - 1)
            ramalMovendo(contador - 1).objIdTrecho = objIdTrecho
            ramalMovendo(contador - 1).objIdRamal = -1
        End If
        'final da etapa
    Else
        MsgBox "Selecione um plano válido, para esta operação", vbExclamation
    End If
End Function
Public Function DrawPoint()
On Error GoTo Trata_Erro
   
   If cgeo.GetLayerTypeReference(tcs.getCurrentLayer) = DOCUMENTOS Then
      tcs.DrawPoint
      tcs.addLayerToSnap tcs.getCurrentLayer
   Else
      MsgBox "Selecione um plano válido, para esta operação", vbExclamation
   End If

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
   Else
      PrintErro "clsTerraLib", "Public Function DrawPoint", CStr(Err.Number), CStr(Err.Description), True
   End If

End Function
' Desenha um ramal junto a rede de água ou esgoto
' Estamos utilizando atualmente apenas ramais de água
'
'
Public Function DrawRamal() As Boolean
On Error GoTo Trata_Erro
   
   If cgeo.IsValidLayerOperation(tcs.getCurrentLayer, RAMAIS_AGUA) Or cgeo.IsValidLayerOperation(tcs.getCurrentLayer, RAMAIS_ESGOTO) Then
      Screen.MousePointer = vbHourglass
      tcs.drawLine
      tcs.addLayerToSnap cgeo.GetLayerOperation(tcs.getCurrentLayer, 1)         ' Seta o layer de snap de referência. Para isso tem que buscar esta informação na tabela GS_LAYER_CONFIG_LAYERS. Onde o layer ativo selecionado e o tipo de operação 1 (desenhar ramal) retorne o número do layer de referência para snap. Por exemplo, se for um ramal o layer de referência de snap é o de redes de água (1)
      tdb.setCurrentLayer tcs.getCurrentLayer                                   ' Seta o layer ativo de desenho o que estiver selecionado pelo usuário
      tdbcon.setCurrentLayer cgeo.GetLayerOperation(tcs.getCurrentLayer, 1)
      tdbconref.setCurrentLayer cgeo.GetLayerOperation(tcs.getCurrentLayer, 1)
      FrmMain.sbStatusBar.Panels(1).Text = "Desenho de ramal: Clique em um treho da rede e desenhe o ramal"
      Screen.MousePointer = vbNormal
   Else
      MsgBox "Selecione um tema de plano de ramais", vbExclamation
   End If
   Exit Function
   
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
       
       PrintErro "clsTerraLib", "Public Function DrawRamal", CStr(Err.Number), CStr(Err.Description), True
       
    End If
    Screen.MousePointer = vbNormal
End Function
' Apaga um elemento do banco de dados geográfico
'
'
'
Public Function Delete() As Boolean
    On Error GoTo Trata_Erro
    Dim i As Integer, rs As ADODB.Recordset
    Dim strsql As String
    Dim linhasSelecionadasMemoria As Integer                            'número total de linhas selecionadas pelo usuário que estão na memória
    Dim pontosSelecionadosMemoria As Integer                            'número total de pontos selecionados pelo usuário que estão na memória
    
    With tcs                            'com o objeto do TeCanvas
        Select Case TerraEvent
            Case tg_SelectObject                                                                                                                                'caso tenha selecionado um objeto para apagar faça, se não selecionou, vai para o final e sai
                Select Case cgeo.GetLayerTypeReference(tcs.getCurrentLayer)  ' cgeo.GetTypeReference(tcs.getCurrentLayer)                                       'conforme o número do layer que foi selecionado
                    Case LayerTypeRefence.Trecho_Rede_Agua, LayerTypeRefence.Trecho_Rede_Drenagem, LayerTypeRefence.Trecho_Rede_esgoto                          'selecionou o layer de redes, vai apagar um trecho de rede
                        If MsgBox("Deseja realmente excluir as rede(s) selecionada(s)", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                            linhasSelecionadasMemoria = .getSelectCount(lines)
                            For i = 0 To linhasSelecionadasMemoria - 1
                                'ZERA O OBJECT_ID_TRECHO DO RAMAL QUE ESTÁ CONECTADO A REDE QUE ESTA SENDO EXCLUÍDA
                                Dim RetornaLayer As String
                                RetornaLayer = UCase(cgeo.GetLayerOperation(.getCurrentLayer, 1))
                                If RetornaLayer = "RAMAIS_AGUA" Or RetornaLayer = "RAMAIS_ESGOTO" Then
                                    a = RetornaLayer
                                    b = "+a+"
                                    c = "OBJECT_ID_TRECHO"
                                    If frmCanvas.TipoConexao <> 4 Then
                                        strsql = "UPDATE " & RetornaLayer & " SET OBJECT_ID_TRECHO = '0' WHERE OBJECT_ID_TRECHO = '" & .getSelectObjectId(i, lines) & "'"
                                    Else
                                        strsql = "UPDATE " + """" + a + """" + " SET " + """" + c + """" + " = '0' WHERE " + """" + c + """" + "  = '" & .getSelectObjectId(i, lines) & "'"
                                    End If
                                    Conn.execute (strsql)
                                End If
                                '*****************************************************************************************************
                                'CHAMA O PROCEDIMENTO DE DELETE DE REDES
                                cgeo.DeleteRede .getCurrentLayer, .getSelectObjectId(i, lines)
                            Next
                        End If
                    Case LayerTypeRefence.AMARRACAO                                                                                                             'layer número 10 - caso deseje apagar uma amarração de rede
                        .deleteSelectGeometry
                    Case LayerTypeRefence.DOCUMENTOS                                                                                                            'layer número 9 - caso deseje apagar um documento associado
                        If MsgBox("Deseja realmente excluir o(s) ponto(s) de documentação", 36) = vbYes Then
                            For i = 0 To tcs.getSelectCount(points) - 1
                                If frmCanvas.TipoConexao <> 4 Then
                                    Conn.execute "Delete From X_Files Where Object_id_='" & tcs.getSelectObjectId(i, points) & "'"
                                Else
                                    c = "X_FILES"
                                    d = "OBJECT_ID_"
                                    Conn.execute "Delete From " + """" + c + """" + " Where " + """" + d + """" + "='" & tcs.getSelectObjectId(i, points) & "'"
                                End If
                            Next
                            .deleteSelectGeometry
                        End If
                    Case LayerTypeRefence.RAMAIS_AGUA, LayerTypeRefence.RAMAIS_ESGOTO                                                                           'layer número 7 - caso deseje apagar um ramal de água ou esgoto. Apaga a Geometria com attributo
                        pontosSelecionadosMemoria = .getSelectCount(points)                                                                                     'número total de pontos selecionados pelo usuário que estão na memória
                        linhasSelecionadasMemoria = .getSelectCount(lines)                                                                                      'número total de linhas selecionadas pelo usuário que estão na memória
                        If pontosSelecionadosMemoria > 0 Or linhasSelecionadasMemoria > 0 Then                                                                  'verifica se o usuário selecionou mais de um ramal, pois só pode apagar um ramal por vez
                            If pontosSelecionadosMemoria > 1 Or linhasSelecionadasMemoria > 1 Then
                                MsgBox "Somente é possivel excluir um ramal por vez", vbExclamation
                                Exit Function
                            End If
                            If frmCanvas.TipoConexao <> 4 Then                                                                                                  'caso seja SQLServer ou Oracle
                                Set rs = Conn.execute("SELECT * from " & tcs.getCurrentLayer & "_LIGACAO WHERE OBJECT_ID_='" & IIf(pontosSelecionadosMemoria = 1, .getSelectObjectId(0, points), .getSelectObjectId(0, lines)) & "'")     'para verificar se existem ligações de água ainda associadas ao ramal, pois não posso apagar um ramal se existem hidrômetros associados ao mesmo
                            Else                                                                                                                                'caso seja Postgres
                                Set rs = Conn.execute("SELECT * from " & """" + tcs.getCurrentLayer & "_LIGACAO" + """" + " WHERE " + """" + "OBJECT_ID_" + """" + "='" & IIf(pontosSelecionadosMemoria = 1, .getSelectObjectId(0, points), .getSelectObjectId(0, lines)) & "'")
                            End If
                            If Not rs.EOF Then                                                                                                                  'existem hidrômetros associados/ativos junto ao ramal, não pode apagar
                                MsgBox "Para excluir o ramal é necessário que ele não tenha nenhum hidrômentro associado.", vbInformation, ""
                                Exit Function
                            End If
                            If MsgBox("Deseja realmente excluir o ramal selecionado", 36) = vbYes Then
                                If pontosSelecionadosMemoria = 1 Then                                                                                             'verifica quantas geometrias foram selecionadas, deveria ser apenas uma geometria, a do ramal
                                    For i = 0 To pontosSelecionadosMemoria - 1
                                        a = "OBJECT_ID_"
                                        If frmCanvas.TipoConexao <> 4 Then                                                                                      'se for SQLServer ou Oracle
                                            Conn.execute "Delete From " & tcs.getCurrentLayer & "  Where Object_id_='" & .getSelectObjectId(i, points) & "'"
                                            Conn.execute "Delete From lines" & cgeo.GetLayerID(.getCurrentLayer) & " Where Object_id='" & .getSelectObjectId(i, points) & "'"
                                            Conn.execute "Delete From points" & cgeo.GetLayerID(.getCurrentLayer) & " Where Object_id='" & .getSelectObjectId(i, points) & "'"
                                        Else                                                                                                                    'se for Postgres
                                            Conn.execute "Delete From " & """" + tcs.getCurrentLayer + """" & "  Where " + """" + "OBJECT_ID_" + """" + "='" & .getSelectObjectId(i, points) & "'"
                                            Conn.execute "Delete From " + """" + "lines" & cgeo.GetLayerID(.getCurrentLayer) & """" + " Where " + """" + "object_id" + """" + "='" & .getSelectObjectId(i, points) & "'"
                                            Conn.execute "Delete From " + """" + "points" & cgeo.GetLayerID(.getCurrentLayer) & """" + " Where " + """" + "object_id" + """" + "='" & .getSelectObjectId(i, points) & "'"
                                        End If
                                    Next
                                Else                                                                                                                            'só pode existir um ponto na extremidade do ramal, o qual é a ligação do hidrômetro, portanto se entrar neste else é por que este ponto não foi selecionado, pois já foi testado antes se existia mais de um ponto
                                    MsgBox "Você não selecionou o hidrômetro, somente o ramal. Selecione ambos ou somente o hidrômetro para apagar."
                                End If
                            End If
                        Else
                            MsgBox "Nenhum ramal selecionado para exclusão", vbExclamation
                        End If
                End Select
                tcs.Normal
                tcs.Select
                tcs.plotView
        End Select
    End With
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        PrintErro "clsTerraLib", "Public Function Delete", CStr(Err.Number), CStr(Err.Description), True
    End If
End Function

Public Function OnPoint(X As Double, Y As Double)
On Error GoTo Trata_Erro
   Dim object_id As String, frm As New FrmAssociation
   tdb.setCurrentLayer tcs.getCurrentLayer
   If tdb.locateGeometry(X, Y, tpPOINTS, 1) = 1 Then
      TerraEvent = tg_SelectObject
      tcs.Normal
      tcs.Select
      tcs.addSelectObjectIds tdb.objectIds(0)
      frm.Init tdb.objectIds(0), tcs, tdb
      
   Else
      TerraEvent = tg_DrawGeometrys
      frm.Init "", tcs, tdb, X, Y
      tcs.Normal
      tcs.Select
      TerraEvent = tg_SelectObject
   End If
   
   Set frm = Nothing

Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
   Else
    
      PrintErro "clsTerraLib", "Public Function OnPoint", CStr(Err.Number), CStr(Err.Description), True
      
   End If
   
End Function
' Entra nesta função quando está desenhando ou selecionando o ramal
' Ela é chamada quando entro o segundo click do ramal
'
'
Public Function OnRamal(X As Double, Y As Double, object_id_ramal As String) As Boolean
    On Error GoTo Trata_Erro
    Dim frm As New FrmCadastroRamal, Qtde As Long, a As Long
    Dim object_id_lote As String
    
    tdbconref.setCurrentLayer cgeo.GetLayerOperation(tcs.getCurrentLayer, 1)
    'Verifica se existe um filtro anterior de ramais e chama se já houver filtro anterior
    If ReadINI("RAMAISFILTROLOTES", "ATIVADO", App.path & "\CONTROLES\GEOSAN.INI") = "SIM" Then
        Dim TBP As String
        TBP = ReadINI("RAMAISFILTROLOTES", "TABELA_PLANO", App.path & "\CONTROLES\GEOSAN.INI")
        If TBP <> "" Then
            tdbconref.setCurrentLayer TBP
        End If
    End If
    'Chama a caixa de diálogo de cadastro de ramais
    object_id_lote = ""    ' zera apenas para aparecer na caixa de diálogo
    If UCase(tcs.getCurrentLayer) = "RAMAIS_AGUA" Then
        'ORIGINAL
        frm.Init "AGUA", object_id_ramal, tcs, tdb, tdbcon, object_id_lote, ""
    Else
        frm.Init "ESGOTO", object_id_ramal, tcs, tdb, tdbcon, object_id_lote, ""
    End If
    Exit Function
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    ElseIf Err.Number = -2147467259 Then
        PrintErro "clsTerraLib", "Public Function OnRamal", CStr(Err.Number), CStr(Err.Description), True
    Else
        PrintErro "clsTerraLib", "Public Function OnRamal", CStr(Err.Number), CStr(Err.Description), True
    End If
End Function
' Função chamada para salvar no banco de dados
'
'
'
'
Public Function SaveInDatabase() As Boolean
    On Error GoTo Trata_Erro
    Dim USR As String
    
    With tcs
        Select Case TerraEvent                                                                  'A variavel TerraEvent determina o evento/funcão que o usuário está utilizando
            Case tg_DrawNetWorkline, tg_MoveNetWorkNode, tg_DrawNetWorkNode, tg_DrawGeometrys   'caso o usuario esteja desenhando uma ou mais geometrias, salva as geometrias
                tcs.saveOnMemory
                tcs.SaveInDatabase
                Screen.MousePointer = vbHourglass
                Screen.MousePointer = vbNormal
            Case tg_SelectObject                                                                'Caso o ususario estiver apenas consultando os atribuitos de uma ou mais geometria. Salva apenas os atributos
                Select Case cgeo.GetLayerTypeReference(tcs.getCurrentLayer)
                    Case LayerTypeRefence.OUTROS
                        MsgBox "Somente leitura", vbExclamation
                    Case Else
                        If tcs.getSelectCount(GetGeometrySELECTed()) = 1 Then
                            varGlobais.realizaCommit = True                                     'indica que está tudo ok para realizar commit. Se algo ocorrer de errado a partir do saveInDatabase ou mesmo nas propriedades, volta como falso e ai não faz nada
                            SaveNetWorkAttributesSingle GetGeometrySELECTed()
                        ElseIf tcs.getSelectCount(GetGeometrySELECTed()) > 1 Then
                            varGlobais.realizaCommit = True                                     'indica que está tudo ok para realizar commit. Se algo ocorrer de errado a partir do saveInDatabase ou mesmo nas propriedades, volta como falso e ai não faz nada
                            SaveNetWorkAttributesMultiple GetGeometrySELECTed()
                        End If
                End Select
                .plotView
            Case tg_MoveNetWorkVertice                                                          'está movendo apenas um vértice da rede de água
'                tcs.saveOnMemory
'                tcs.SaveInDatabase
'                tcs.plotView
            Case tg_MoveGpsPoint
                tcs.saveOnMemory
                tcs.SaveInDatabase
                Screen.MousePointer = vbHourglass
                Screen.MousePointer = vbNormal
        End Select
    End With
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    ElseIf mid(Err.Description, 1, 9) = "ORA-03114" Then
        PrintErro "clsTerraLib", "Public Function SaveInDatabase()", CStr(Err.Number), CStr(Err.Description), True
        End
    Else
        PrintErro "clsTerraLib", "Public Function SaveInDatabase()", CStr(Err.Number), CStr(Err.Description), True
    End If
    Screen.MousePointer = vbNormal
End Function

Public Function GetGeometrySELECTed() As TypeGeometry
On Error GoTo Trata_Erro
   If tcs.getSelectCount(Polyguns) > 0 Then
      GetGeometrySELECTed = Polyguns
   ElseIf tcs.getSelectCount(lines) > 0 Then
      GetGeometrySELECTed = lines
   ElseIf tcs.getSelectCount(points) > 0 Then
      GetGeometrySELECTed = points
   ElseIf tcs.getSelectCount(texts) > 0 Then
      GetGeometrySELECTed = texts
   End If
   Exit Function
   
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
      Resume Next
   Else
      PrintErro "clsTerraLib", "Public Function GetGeometrySELECTed()", CStr(Err.Number), CStr(Err.Description), True
      
   End If
End Function
' Salva somente os atributos de um trecho de rede
'
'
'
Public Function SaveNetWorkAttributesSingle(Geometry As TypeGeometry)
    On Error GoTo Trata_Erro
    
    If CtrlMgr.Itens.Item(1).ValueStore >= 0 Then
       CtrlMgr.SaveProperties tcs.getSelectObjectId(0, Geometry), strUser 'inserido dia 25/11/08 Jonathas
       If Geometry = lines Then
          cgeo.UpdateTextsInLines tcs.getCurrentLayer, tcs.getSelectObjectId(0, Geometry)
       Else
          'SALVA TEXTO DE PONTO SOMENTE SE O PONTO FOR UMA PEÇA DE ESGOTO
          If tcs.getCurrentLayer = "SEWERCOMPONENTS" Then
             cgeo.InsertTextInPoint tcs.getSelectObjectId(0, Geometry), tcs.getSelectObjectId(0, Geometry), tcs.getCurrentLayer
          End If
       End If
    End If
    Exit Function
    
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       varGlobais.realizaCommit = False                      'pede para voltar tudo o que está fazendo no banco de dados, para traz e não comitar nada
       ErroUsuario.Registra "clsTerraLib", "SaveNetWorkAttributesSingle", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Function
' Salva os atributos de vários trechos de rede
'
'
'
Public Function SaveNetWorkAttributesMultiple(Geometry As TypeGeometry) As Boolean
    On Error GoTo Trata_Erro
    Dim object_id As String, i As Integer
    If CtrlMgr.Itens.Item(1).ValueStore >= 0 Then
       For i = 0 To tcs.getSelectCount(Geometry) - 1
          object_id = tcs.getSelectObjectId(i, Geometry)
          If object_id <> "" Then
              CtrlMgr.SaveMultProperties object_id, tcs.getCurrentLayer
              If Geometry = lines Then
                 cgeo.UpdateTextsInLines tcs.getCurrentLayer, object_id
              Else
                 cgeo.InsertTextInPoint CLng(object_id), object_id, tcs.getCurrentLayer
                 'cgeo.InsertTextInPoint tcs.getCurrentLayer, object_id
              End If
          End If
       Next
    End If
    Exit Function
   
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
       varGlobais.realizaCommit = False                      'pede para voltar tudo o que está fazendo no banco de dados, para traz e não comitar nada
       ErroUsuario.Registra "clsTerraLib", "SaveNetWorkAttributesMultiple", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Function
' Cria agora a tabela de atributos de rede
'
' LINE_ID - Id da linha a ser criada
' Node_id1 - Nó inicial da linha a ser criado ou atualizado
' Node_id2 - Nó final da linha a ser criado ou atualizado
' Movendo - indica se está criando ou movendo um trecho de rede
'
Public Function CreatNetWorkAttribute(LINE_ID As Long, Node_id1 As Long, Node_id2 As Long, Movendo As Boolean) As Boolean
    On Error GoTo Trata_Erro
    Dim az As String
    Dim object_id As String, CompCalc As Double, a As Integer, ld As String, nd1 As String, nd2 As String
    Dim LayerName As String
    Dim layerReferencia As String
    Dim i As Byte
    Dim No As String
    Dim getZNo As New CObtemZdoNo                           'método para o cálculo da cota z do nó inserido
    Dim zNo As Double                                       'cota z do nó inserido que será calculada
    Dim stringSQL As String
    Dim rsq As New ADODB.Recordset
    Dim TBGeometria As String
    Dim compMax As Double
    Dim getObjIdsNos As New CObtemObjIdsNos                 'para receber os obj_ids dos nós inicial e final
    Dim noInicial As String
    Dim noFinal As String
    Dim retorno As Boolean
    
    If TerraEvent = tg_DrawNetWorkline Then
        LayerName = tcs.getCurrentLayer
    Else
        LayerName = cgeo.GetReferenceLayer(tcs.getCurrentLayer)
    End If
    layerReferencia = cgeo.GetReferenceLayer(LayerName)                         'obtem o nome do layer de atributos de referência WaterComponents ou SwerComponents
    TBGeometria = tdb.getRepresentationTableName(layerReferencia, tpPOINTS)     'obtem o nome do layer de geometrias de referência Points2 ou Points4
    If Movendo = True Then                                                      'SE ESTIVER APENAS MOVENDO, EXCLUI E REFAZ O TEXTO DA LINHA, E ATUALIZA A COTA DO NÓ
        LayerName = tcs.GetReferenceLayer                                       'SETA O LAYER DE REFERENCIA COMO ATIVO
        cgeo.DeleteTextObjectWithInsertText LayerName, LINE_ID, CStr(LINE_ID)
        'atualização da cota do nó
        'Antes de atribuir o Z do nó verifica se o usuário quer que isso seja realizado enquanto ele desenha a rede. Se o MDT não estiver pronto o usuário pode optar por desligar esta opção
        If varGlobais.deveCalcularZNo = True Then
            No = varGlobais.objIdNoSelecionado
            
            If No <> "NULO" Then
                If frmCanvas.TipoConexao <> 4 Then      'caso seja Oracle ou SQLServer
                    stringSQL = "SELECT object_id FROM " & TBGeometria & " WHERE OBJECT_ID = " & No     'Tabela Points2 para água ou Points4 para esgoto
                    rsq.Open stringSQL, Conn, adOpenDynamic, adLockReadOnly
                Else    'caso seja Postgres
                        'implementar
                End If
                If rsq.EOF = False Then ' SE ENCONTROU O REGISTRO, ou seja o nó existe
                    zNo = getZNo.CObtemZFromObjIdNo(No, TBGeometria)                     'obtem a cota z do nó a partir do object_id do mesmo. Passa o nome do layer de geometrias dos nós de água o esgoto
                    If frmCanvas.TipoConexao <> 4 Then
                        stringSQL = "UPDATE " & layerReferencia & " SET GROUNDHEIGHT = " & Replace(Round(zNo, 2), ",", ".") & " where object_id_ = " & No   'atualiza a cota do nó
                        Conn.execute (stringSQL)
                    Else 'Postgres
                        'Implementar
                    End If
                Else                    'nó não existe
                    'deveria ter encontrado o nó - exibir mensagem de erro
                    MsgBox ("Nó " & No & " não foi encontrado.")
                End If
                rsq.Close
                varGlobais.objIdNoSelecionado = "NULO"                      'para rodar só uma vez (a primeira), pois entra nesta rotina várias vezes para cada seguimento de rede
            End If
        End If
        'fim da atualização da cota do nó movido
    Else                    ' ESTÁ DESENHANDO A REDE
        Set rs = New ADODB.Recordset
        'prepara querie para procurar em todas as linhas de rede se já existe alguma com o object_id_ da que será criada
        If frmCanvas.TipoConexao <> 4 Then
            rsq.Open ("SELECT * FROM " & LayerName & " WHERE OBJECT_ID_ = " & LINE_ID), Conn, adOpenDynamic, adLockReadOnly
        Else
            b = "OBJECT_ID_"
            rsq.Open ("SELECT * FROM " + """" + LayerName + """" + " WHERE " + """" + b + """" + " = '" & LINE_ID & "'"), Conn, adOpenDynamic, adLockOptimistic
        End If
        If rsq.EOF = True Then ' SE NÃO ENCONTROU O REGISTRO
            'CRIA UM NOVO REGISTRO NA TABELA DE ATRIBUTOS DE LINHAS COM LINE_ID,OBJECT_ID_ IGUAL AO CÓDIGO DA GEOMETRIA
            ' O PRIMNEIRO LINE ID É NUMERICO E O SEGUNDO É TEXTO
            If frmCanvas.TipoConexao <> 4 Then
                stringSQL = "INSERT INTO " & LayerName & " (LINE_ID,OBJECT_ID_) VALUES (" & LINE_ID & ",'" & LINE_ID & "')"
            Else
                az = "LINE_ID"
                b = "OBJECT_ID_"
                c = "INSCRICAO_LOTE"
                stringSQL = "INSERT INTO  " + """" + LayerName + """" + " ( " + """" + az + """" + "," + """" + b + """" + ") VALUES ('" & LINE_ID & "','" & LINE_ID & "')"
            End If
            Conn.execute (stringSQL)
        Else
            'NÃO DEVERIA TER ENCONTRADO
        End If
        rsq.Close
        'DLL Pm4Manager SALAVA AS PRORPIEDADES DA LINHA NA TABELA DE ATRIBUTOS
        CtrlMgr.SaveProperties CStr(LINE_ID), strUser
        'OBTEM NA VARIÁVEL CompCalc O COMPRIMENTO DA NOVA LINHA
        tdb.getLengthOfLine LINE_ID, CStr(LINE_ID), CompCalc
        'ATUALIZADO O COMPRIMENTO, NÓ INICIAL E NÓ FINAL DA LINHA NA TABELA DE ATRIBUTOS
        compMax = CompCalc  ' Replace(Round(CompCalc, 2), ",", ".")
        If frmCanvas.TipoConexao <> 4 Then
            If compMax > 9999.99 Then
                ' SE O COMPRIMENTO MAXIMO DE UMA REDE FOR ULTRAPASSADO, ELA SERÁ SALVA NO ENTANTO COM O LIMITE MAXIMO DE COMPRIMENTO
                MsgBox "Rede com comprimento máximo '9999.99' ultrapassado.", vbExclamation, "Aviso"
                stringSQL = "UPDATE " & LayerName & " SET LENGTHCALCULATED = 9999.99, INITIALCOMPONENT = '" & Node_id1 & "', FINALCOMPONENT = '" & Node_id2 & "' WHERE OBJECT_ID_ = '" & LINE_ID & "'"
            Else
                stringSQL = "UPDATE " & LayerName & " SET LENGTHCALCULATED = " & Replace(Round(CompCalc, 2), ",", ".") & ", INITIALCOMPONENT = '" & Node_id1 & "', FINALCOMPONENT = '" & Node_id2 & "' WHERE OBJECT_ID_ = '" & LINE_ID & "'"
            End If
        Else
            az = LayerName
            c = "LENGTHCALCULATED"
            d = "INITIALCOMPONENT"
            e = "FINALCOMPONENT"
            f = "OBJECT_ID_"
            g = Round(Replace(Round(CompCalc, 2), ",", "."))
            h = Round(g)
            If compMax > 9999.99 Then
                ' SE O COMPRIMENTO MAXIMO DE UMA REDE FOR ULTRAPASSADO, ELA SERÁ SALVA NO ENTANTO COM O LIMITE MAXIMO DE COMPRIMENTO
                MsgBox "Rede com comprimento máximo '9999.99' ultrapassado.", vbExclamation, "Aviso"
                stringSQL = "UPDATE " + """" + az + """" + " SET " + """" + c + """" + " = '9999.99', " + """" + d + """" + " = '" & Node_id1 & "', " + """" + e + """" + " = '" & Node_id2 & "' WHERE " + """" + f + """" + " = '" & LINE_ID & "'"
            Else
                ' MsgBox "UPDATE " + """" + az + """" + " SET " + """" + c + """" + " = '" + g + "', " + """" + d + """" + " = '" & Node_id1 & "', " + """" + e + """" + " = '" & Node_id2 & "' WHERE " + """" + f + """" + " = '" & line_id & "'"
                stringSQL = "UPDATE " + """" + LayerName + """" + " SET " + """" + c + """" + " = '" + h + "', " + """" + d + """" + " = '" & Node_id1 & "', " + """" + e + """" + " = '" & Node_id2 & "' WHERE " + """" + f + """" + " = '" & LINE_ID & "'"
                'WritePrivateProfileString "A", "A", stringSQL, App.path & "\DEBUG.INI"
            End If
        End If
        Conn.execute (stringSQL)
        'DEPOIS DE CARREGADAS AS INFORMAÇÕES NA TABELA DE ATRIBUTOS, CHAMA O MÉTODO DE CRIAR TEXTOS
        InsertTextAttributesLine LayerName, LINE_ID, CStr(LINE_ID)                                  'aqui entre outras coisas insere as cotas iniciais e finais no trecho de rede de esgoto, quanto for esgoto
        'xxxxxxxxxxx
        'INSERE O ATRIBUTO DO NÓ 1 E NÓ 2 NA TABELA DE ATRIBUTOS COM O SEGUINTE LOOP
        For i = 1 To 2 'EXECUTA 2 VEZES A ROTINA ABAIXO
            If i = 1 Then
                No = Node_id1
            Else
                No = Node_id2
            End If
            Set rs = New ADODB.Recordset
            If frmCanvas.TipoConexao <> 4 Then
                rsq.Open ("SELECT * FROM " & layerReferencia & " WHERE OBJECT_ID_ = '" & No & "'"), Conn, adOpenDynamic, adLockReadOnly
            Else
                az = "OBJECT_ID_"
                rsq.Open ("SELECT * FROM " + """" + layerReferencia + """" + " WHERE " + """" + az + """" + " = '" & No & "'"), Conn, adOpenDynamic, adLockOptimistic
            End If
            If rsq.EOF = True Then ' SE NÃO ENCONTROU O REGISTRO, ou seja o nó não existe ainda
                'Aqui ele atribui o Z ao nó
                'Antes de atribuir o Z do nó verifica se o usuário quer que isso seja realizado enquanto ele desenha a rede. Se o MDT não estiver pronto o usuário pode optar por desligar esta opção
                If varGlobais.deveCalcularZNo = True Then
                    zNo = getZNo.CObtemZFromObjIdNo(No, TBGeometria)                     'obtem a cota z do nó a partir do object_id do mesmo. Passa o nome do layer de geometrias dos nós de água o esgoto
                Else
                    zNo = 0
                End If
                If frmCanvas.TipoConexao <> 4 Then
                    stringSQL = "INSERT INTO " & layerReferencia & " (COMPONENT_ID,OBJECT_ID_, GROUNDHEIGHT) VALUES (" & No & ",'" & No & "', " & Replace(Round(zNo, 2), ",", ".") & ")"
                    Conn.execute (stringSQL)
                Else
                    az = "OBJECT_ID_"
                    b = "COMPONENT_ID"
                    c = "INSCRICAO_LOTE"
                    d = "TIPO"
                    e = "HIDROMETRADO"
                    f = "ECONOMIAS"
                    g = "CONSUMO_LPS"
                    h = "TB_LIGACOES"
                    Conn.execute ("INSERT INTO " + """" + layerReferencia + """" + " (" + """" + b + """" + "," + """" + az + """" + ") VALUES ('" & No & "','" & No & "')")
                End If
                'CRIA UM NOVO REGISTRO NA TABELA DE ATRIBUTOS DE NÓS COM COMPONENT_ID,OBJECT_ID_
                If cgeo.GetTypeText(LayerName) = 2 Then
                    cgeo.InsertTextInPoint CLng(No), CStr(No), layerReferencia, 0, 0
                End If
            End If
            rsq.Close
        Next
        'Rotina que procura Atributos por referencia
        'If FrmMain.mnuLoadAttributeByReference.Checked Then
        '   LoadAttributeByReference Node_id1
        '   LoadAttributeByReference Node_id2
        'End If
    End If
    CreatNetWorkAttribute = True
    Exit Function

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else
        varGlobais.realizaCommit = False                      'pede para voltar tudo o que está fazendo no banco de dados, para traz e não comitar nada
        ErroUsuario.Registra "clsTerraLib", "CreatNetWorkAttribute", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
    End If
End Function

Private Function NewObject(LayerName As String, ID As Long, object_id As String, Geometry As TypeGeometry, Optional compr As Double) As Boolean
On Error GoTo Trata_Erro
   Dim LenghtCalc As Double
   Dim rsNew As ADODB.Recordset
   Dim strsql As String
   
   If Geometry = lines Then


  
     

     If frmCanvas.TipoConexao <> 4 Then
         
    strsql = "INSERT INTO " & LayerName & " (LINE_ID,OBJECT_ID_) VALUES (" & ID & ",'" & ID & "')"
     Else
     a = "OBJECT_ID_"
      b = "LINE_ID"
      c = "INSCRICAO_LOTE"
      
      e = "HIDROMETRADO"
      f = "LENGTHCALCULATED"
     
      strsql = "INSERT INTO " + """" + LayerName + """" + " (" + """" + b + """" + "," + """" + a + """" + ") VALUES ('" & ID & "','" & ID & "')"
     End If


     
      Conn.execute (strsql)
   
      CtrlMgr.SaveProperties object_id, strUser 'variável strUser incluida em 25/11/08 Jonathas
         
         If frmCanvas.TipoConexao <> 4 Then
         
         Conn.execute "Update " & LayerName & " set lengthcalculated= " & Replace(Round(compr, 2), ",", ".") & " where Object_id_ = " & object_id
         Else
         a = "OBJECT_ID_"
      b = "LINE_ID"
      c = "INSCRICAO_LOTE"
      
      e = "HIDROMETRADO"
      f = "LENGTHCALCULATED"
         Conn.execute "Update " + """" + LayerName + """" + " set " + """" + f + """" + " = '" & Replace(Round(compr, 2), ",", ".") & "' where " + """" + a + """" + " = '" & object_id & "'"
         End If
      'CHAMA O MÉTODO DE CRIAR TEXTOS
      InsertTextAttributesLine LayerName, ID, object_id

   
   ElseIf Geometry = points Then
      
     


     If frmCanvas.TipoConexao <> 4 Then
         
    strsql = "INSERT INTO " & LayerName & " (COMPONENT_ID,OBJECT_ID_) VALUES (" & ID & ",'" & ID & "')"
     Else
      a = "OBJECT_ID_"
      b = "COMPONENT_ID"
      c = "INSCRICAO_LOTE"
     
     
     strsql = "INSERT INTO " + """" + LayerName + """" + " (" + """" + b + """" + "," + """" + a + """" + ") VALUES ('" & ID & "','" & ID & "')"
     End If
      
      
      
     
      Conn.execute (strsql)
      If cgeo.GetTypeText(LayerName) = 2 Then
         cgeo.InsertTextInPoint ID, object_id, LayerName, 0, 0
      End If
    
   End If
      
   NewObject = True

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
      
      PrintErro "clsTerraLib", "Private Function NewObject", CStr(Err.Number), CStr(Err.Description), True
       
    End If
End Function
' Insere os atributos de texto da linha no banco de dados
'
' LayerName - nome do layer
' geom_id - da linha
' object_id - da linha
'
Private Function InsertTextAttributesLine(LayerName As String, geom_id As Long, object_id As String) As Boolean
    On Error GoTo Trata_Erro
    Dim Length As Double, Diameter As String, Material As String, CotaIni As Double, CotaFim As Double
    
    a = "X_MATERIAL"
    b = "MATERIAL"
    c = "MATERIALID"
    d = "OBJECT_ID_"
    If frmCanvas.TipoConexao <> 4 Then
        Set rs = Conn.execute("SELECT * From  " & tcs.getCurrentLayer & " left JOIN x_Material on material=materialid where object_id_='" & object_id & "'")
    Else
        'MsgBox "ARQUIVO DEBUG SALVO"
        'WritePrivateProfileString "A", "A", "SELECT * From  " + """" + tcs.getCurrentLayer + """" + " left JOIN " + """" + a + """" + " on" + """" + b + """" + "=" + """" + c + """" + " where " + """" + d + """" + "='" & object_id & "'", App.path & "\DEBUG.INI"
        Set rs = Conn.execute("SELECT * From  " + """" + tcs.getCurrentLayer + """" + " left JOIN " + """" + a + """" + " on" + """" + b + """" + "=" + """" + c + """" + " where " + """" + d + """" + "='" & object_id & "'")
    End If
    If Not rs.EOF Then
        Diameter = Chr$(216) & " " & IIf(IsNull(rs!INTERNALDIAMETER), 0, rs!INTERNALDIAMETER)
        Material = IIf(IsNull(rs!MATERIALNAME), 0, rs!MATERIALNAME)
        Length = IIf(rs!Length = 0, rs!LENGTHCALCULATED, rs!Length)
        If tcs.getCurrentLayer = "SEWERLINES" Then                                                              'aqui deve ser inserida a cota a partir do MDT. Tenho o geom_id e obj_id da rede de esgoto
            CotaIni = rs!INITIALTUBEDEEPNESS
            CotaFim = rs!FINALTUBEDEEPNESS
        Else
            CotaIni = rs!INITIALGROUNDHEIGHT
            CotaFim = rs!FINALGROUNDHEIGHT
        End If
        'insere no banco de dado os textos do trecho de rede
        If cgeo.GetTypeText(tcs.getCurrentLayer) = 1 Then
            cgeo.InsertTextInDatabase geom_id, object_id, Diameter, Material, Length                           'insere sem as cotas iniciais e finais (redes de água)
        Else
            cgeo.InsertTextInDatabase geom_id, object_id, Diameter, Material, Length, True, CotaIni, CotaFim   'insere com as cotas iniciais e finais (redes de esgoto e drenagem)
        End If
    End If
    rs.Close
    Set rs = Nothing

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
      PrintErro "clsTerraLib", "Private Function InsertTextAttributes", CStr(Err.Number), CStr(Err.Description), True
    End If
End Function

Public Function CreatNetWorkNode(ByVal node_id As Long, ByVal line1_id As Long, ByVal line2_id As Long, Movendo As Boolean)

On Error GoTo Trata_Erro
   
   Dim Nd As String, ld1 As String, ld2 As String, object_id As String
   Dim IniComp As String, FimComp As String, CompCalc As Double, LayerName As String, LayerREf As String
   
   LayerREf = tcs.GetReferenceLayer ' cgeo.GetReferenceLayer(LayerName)
   tdbcon.setCurrentLayer LayerREf 'cgeo.GetReferenceLayer(tcs.getCurrentLayer)
   LayerName = tcs.getCurrentLayer
   
   If Movendo = True Then
      
      'VERIFICA SE O TIPO DO PONTO É ESGOTO, CASO SIM ELE RECEBE TEXTO
      If cgeo.GetTypeText(LayerName) = 2 Then
         
         cgeo.InsertTextInPoint node_id, CStr(node_id), LayerName, 0, 0

      End If
   
   Else
      
      'RETORNA O NÓ INICIAL E FINAL DA LINHA EXISTENTE QUE RECEBEU O NOVO NÓ
      Set rs = New ADODB.Recordset
      If frmCanvas.TipoConexao <> 4 Then
      rs.Open ("SELECT INITIALCOMPONENT, FINALCOMPONENT FROM " & LayerName & " WHERE OBJECT_ID_ = '" & line1_id & "'"), Conn, adOpenForwardOnly, adLockReadOnly
      Else
      f = "INITIALCOMPONENT"
      g = "FINALCOMPONENT"
      h = "OBJECT_ID_"
      'Dim aaa As String
      'aaa = "SELECT " + """" + f + """" + ", " + """" + g + """" + " FROM " + """" + LayerName + """" + " WHERE " + """" + h + """" + " = '" & line1_id & "'"

      
   '   MsgBox "ARQUIVO DEBUG SALVO"
' WritePrivateProfileString "A", "A", aaa, App.path & "\DEBUG.INI"

      
      
      
      
      rs.Open ("SELECT " + """" + f + """" + ", " + """" + g + """" + " FROM " + """" + LayerName + """" + " WHERE " + """" + h + """" + " = '" & line1_id & "'"), Conn, adOpenDynamic, adLockOptimistic
      End If
      
      If rs.EOF = False Then
         IniComp = rs.Fields("InitialComponent").value
         FimComp = rs.Fields("FinalComponent").value
      End If
      
      rs.Close
      Set rs = Nothing
   
      Nd = node_id
      'MOMENTO QUE CRIA O ATRIBUTO DO NOVO COMPONENTE DE REDE
      
      'strsql = "INSERT INTO " & LayerName & " (COMPONENT_ID,OBJECT_ID_) VALUES (" & node_id & ",'" & node_id & "')"
      
      
     If frmCanvas.TipoConexao <> 4 Then
         
   Conn.execute ("INSERT INTO " & LayerREf & " (COMPONENT_ID,OBJECT_ID_) VALUES (" & node_id & ",'" & node_id & "')")
     
     Else
      
        a = "OBJECT_ID_"
      b = "COMPONENT_ID"
      c = "LINE_ID"
    
     
   Conn.execute ("INSERT INTO " + """" + LayerREf + """" + " (" + """" + b + """" + "," + """" + a + """" + ") VALUES ('" & node_id & "','" & node_id & "')")
     End If
      
      
      
      
      'VERIFICA SE O TIPO DO PONTO É ESGOTO, CASO SIM ELE RECEBE TEXTO
      If cgeo.GetTypeText(LayerName) = 2 Then
         cgeo.InsertTextInPoint node_id, CStr(node_id), LayerREf, 0, 0
      End If
      
         
      'CRIA ATRIBUTOS PARA A NOVA LINHA
      
      If frmCanvas.TipoConexao <> 4 Then
         
   Conn.execute ("INSERT INTO " & LayerName & " (LINE_ID,OBJECT_ID_) VALUES (" & line2_id & ",'" & line2_id & "')")
   
     Else
      a = "OBJECT_ID_"
      b = "COMPONENT_ID"
      c = "LINE_ID"
      d = "TIPO"
     
   Conn.execute ("INSERT INTO " + """" + LayerName + """" + " (" + """" + c + """" + "," + """" + a + """" + ") VALUES ('" & line2_id & "','" & line2_id & "')")
   
   End If
      
            
      'DLL Pm4Manager SALVA OS ATRIBUTOS
      CtrlMgr.SaveProperties CStr(line2_id), strUser 'variável strUser incluida em 25/11/08 Jonathas
         
      'ATUALIZA O COMPRIMENTO DA NOVA LINHA
      tdb.getLengthOfLine line2_id, CStr(line2_id), CompCalc
      If frmCanvas.TipoConexao <> 4 Then
      Conn.execute "UPDATE " & LayerName & " SET LENGTHCALCULATED = " & Replace(Round(CompCalc, 2), ",", ".") & " WHERE OBJECT_ID_ = '" & line2_id & "'"
      Else
      a = LayerName
b = "+a+"
c = "LENGTHCALCULATED"
d = Round(Replace(Round(CompCalc, 2), ",", "."))
e = "'d'"
f = "OBJECT_ID_"

       Conn.execute "UPDATE " + """" + a + """" + " SET " + """" + c + """" + " = '" & d & "' WHERE " + """" + f + """" + " = '" & line2_id & "'"
      End If
      'ATUALIZA COMPONENTE INICIAL E FINAL DA NOVA LINHA
      If frmCanvas.TipoConexao <> 4 Then
      
      
     
      
      Conn.execute "Update " & LayerName & " set InitialComponent= " & node_id & ", FinalComponent=" & FimComp & " WHERE OBJECT_ID_ = '" & line2_id & "'"
      Else
      
      a = LayerName

c = "LENGTHCALCULATED"


f = "OBJECT_ID_"
g = "NODE_ID"
h = "INITIALCOMPONENT"
i = "FINALCOMPONENT"
j = FimComp

Dim aaa, bbb, ccc As String
aaa = str(j)
bbb = str(node_id)
ccc = str(line2_id)


      
      '  Dim bbbb As String
   ' bbbb = "Update " + """" & LayerName & """" + " set " + """" + "INITIALCOMPONENT" + """" + "=" + bbb
   '   MsgBox "ARQUIVO DEBUG SALVO"
 'WritePrivateProfileString "A", "A", bbbb, App.path & "\DEBUG.INI"

      
      Conn.execute "Update " + """" & LayerName & """" + " set " + """" + "INITIALCOMPONENT" + """" + "='" + bbb + "', " + """" + "FINALCOMPONENT" + """" + "= ' " & aaa & " ' WHERE " + """" + "OBJECT_ID_" + """" + " = '" & ccc & "'"
      

      End If
      'CHAMA O MÉTODO DE CRIAR TEXTOS PARA A NOVA LINHA
      InsertTextAttributesLine LayerName, line2_id, CStr(line2_id)
         

            
      
      'ATUALIZA O COMPRIMENTO DA NOVA LINHA
      tdb.getLengthOfLine line1_id, CStr(line1_id), CompCalc
       If frmCanvas.TipoConexao <> 4 Then
      Conn.execute "Update " & LayerName & " set lengthcalculated= " & Replace(Round(CompCalc, 2), ",", ".") & " where Object_id_ = '" & line1_id & "'"
      Else
      a = LayerName
b = "+a+"
c = "LENGTHCALCULATED"
d = Round(Replace(Round(CompCalc, 2), ",", "."))
e = "'d'"
f = "OBJECT_ID_"
g = "NODE_ID"
h = "INITIALCOMPONENT"
i = "FINALCOMPONENT"
j = IniComp
k = "'j'"
l = node_id
m = "'l'"
      Conn.execute "Update " + """" + a + """" + " set " + """" + c + """" + "='" + d + "' where " + """" + f + """" + " = '" & line1_id & "'"
      End If
      'ATUALIZA COMPONENTE INICIAL E FINAL DA ANTIGA LINHA
        If frmCanvas.TipoConexao <> 4 Then
      Conn.execute "Update " & LayerName & " set InitialComponent= " & IniComp & ", FinalComponent=" & node_id & " WHERE OBJECT_ID_ = '" & line1_id & "'"
      Else
      
                    
      
      a = LayerName
b = "+a+"
c = "LENGTHCALCULATED"
d = Replace(Round(CompCalc, 2), ",", ".")
e = "'d'"
f = "OBJECT_ID_"
g = "NODE_ID"
h = "INITIALCOMPONENT"
i = "FINALCOMPONENT"
j = IniComp
k = "'j'"
l = node_id
m = "'l'"
      
      
      
        Conn.execute "Update " + """" + a + """" + " set " + """" + h + """" + " = '" & IniComp & "', " + """" + i + """" + "='" & node_id & "' WHERE " + """" + f + """" + " = '" & line1_id & "'"
      End If
      'CHAMA O MÉTODO DE EXCLUIR E CRIAR TEXTOS PARA A ANTIGA LINHA
      cgeo.DeleteTextObjectWithInsertText LayerName, line1_id, CStr(line1_id)
         
      

      
      If UCase(LayerName) = "WATERLINES" Then
         'REASSOCIA OS RAMAIS DE AGUA A NOVA REDE PARTINDO DO PRINCÍPIO QUE ELA ESTEJA NO MESMO LOCAL DA ANTIGA
         Set rs = New ADODB.Recordset
         If frmCanvas.TipoConexao <> 4 Then
         rs.Open ("SELECT OBJECT_ID_ FROM RAMAIS_AGUA WHERE OBJECT_ID_TRECHO = '" & line1_id & "'"), Conn, adOpenDynamic, adLockReadOnly
         Else
            a = "RAMAIS_AGUA"
      b = "OBJECT_ID_"
      c = "OBJECT_ID_TRECHO"
        rs.Open ("SELECT " + """" + b + """" + " FROM " + """" + a + """" + " WHERE " + """" + c + """" + " = '" & line1_id & "'"), Conn, adOpenDynamic, adLockOptimistic
         End If

     If frmCanvas.TipoConexao <> 4 Then
         Conn.execute ("DELETE FROM POLIGONO_SELECAO WHERE USUARIO = '" & strUser & "' AND TIPO = '2'")
         Else
a = "POLIGONO_SELECAO"
b = "USUARIO"
c = "TIPO"

         Conn.execute ("DELETE FROM " + """" + a + """" + " WHERE " + """" + b + """" + " = '" & strUser & "' AND " + """" + c + """" + " = '2'")
         End If
         If rs.EOF = False Then
            Do While Not rs.EOF
            
              

     If frmCanvas.TipoConexao <> 4 Then
    
         
     Conn.execute ("INSERT INTO POLIGONO_SELECAO (OBJECT_ID_,USUARIO,TIPO) VALUES ( '" & rs!Object_id_ & "','" & strUser & "',2)")
     
     Else
     a = "POLIGONO_SELECAO"
      b = "OBJECT_ID_"
      c = "USUARIO"
      d = "TIPO"
      e = "HIDROMETRADO"
      f = "ECONOMIAS"
      g = "CONSUMO_LPS"
      h = "TB_LIGACOES"


 Conn.execute ("INSERT INTO " + """" + a + """" + " (" + """" + b + """" + "," + """" + c + """" + "," + """" + d + """" + ") VALUES ( '" & rs!Object_id_ & "','" & strUser & "','2')")
     End If
            
            
              
               rs.MoveNext
            Loop
            
            'CHAMA A FUNÇÃO PARA RELOCALIZAR TRECHOS DOS RAMAIS DE AGUA
            frmAlteraNoPoligono.ATUALIZA_TRECHOS_RAMAIS_AGUA
         
         End If
         rs.Close
        
      End If
   End If
   
'   'Rotina que procura Atributos por referencia
'   If FrmMain.mnuLoadAttributeByReference.Checked Then
'      LoadAttributeByReference node_id
'   End If
   
Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
      
      PrintErro "clsTerraLib", "Public Function CreatNetWorkNode", CStr(Err.Number), CStr(Err.Description), True
       
    End If
End Function

Public Function DistanceBetween(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double) As Double
  ' Calculate the distance between two points, given their X/Y coordinates.
  
  ' The short version...
  DistanceBetween = Sqr((Abs(X2 - X1) ^ 2) + (Abs(Y2 - Y1) ^ 2))
  
End Function


Public Function SearchGeomtryForAttribute()
'   On Error GoTo SearchGeomtryForAttribute_err
'   Dim xmin As Double, ymin As Double, xmax As Double, ymax As Double, Object_id As String
'   Object_id = FrmMsg.init("Entre com o número da inscrição", gsInteger)
'
'   CleanAllGeometries
'   If Object_id <> "" Then
'      With tcs
'         .Normal
'         Set rs = Conn.execute("SELECT Object_id_ From lotes where inscricao='" & Object_id & "'")
'         If Not rs.EOF Then
'            Object_id = IIf(IsNull(rs!Object_id_), "", rs!Object_id_)
'            If Object_id <> "" Then
'               If .addSELECTObjectIds(Object_id) = 1 Then
'                  .getSELECTBox xmin, ymin, xmax, ymax
'                  .setWorld xmin - 1000, ymin - 1000, xmax + 1000, ymax + 1000
'                  .SELECT
'                  .setScale 1000
'               Else
'                  MsgBox "Objecto não encontrado", vbExclamation
'               End If
'            Else
'               MsgBox "Objecto não encontrado", vbExclamation
'            End If
'         Else
'            MsgBox "Número da inscrição não encontrado", vbExclamation
'
'         End If
'         rs.Close
'         Set rs = Nothing
'      End With
'   End If
'   Exit Function
'SearchGeomtryForAttribute_err:
'   MsgBox Err.Description
   Dim frm As New FrmPesquisaCliente
   frm.Init tcs
   Set frm = Nothing
End Function

Public Sub DrawInterSection()
   tcs.drawWithCircleIntersection True
End Sub

Private Sub CleanAllGeometries()
   With tcs
     .clearSelectItens 1
     .clearSelectItens 2
     .clearSelectItens 128
     .clearSelectItens 4
   End With
End Sub


Private Sub Class_Initialize()
   Set cgeo = New clsGeoReference
End Sub

Private Sub Class_Terminate()
   Set cgeo = Nothing
End Sub

Public Function PmNewLine(LayerName As String, Geometry As TypeGeometry, object_id As String) As Boolean
   
   'CRIA UMA NOVA LINHA NA TABELA DE ATRIBUTOS
   
   Dim Object_id_  As String
   Dim rsNew As ADODB.Recordset
   Set rsNew = New ADODB.Recordset
   
   rsNew.Open LayerName, Conn, adOpenKeyset, adLockOptimistic, adCmdTable
   
   rsNew.AddNew
                                          'INSERE MESMO CÓDIGO DO AUTO NUMBER DA TABELA DE GEOMETRIAS
   rsNew("OBJECT_ID_").value = object_id  '                                                       XXXXX strUser & Now ' INSERE VALOR EM OBJECT_ID_ PARA QUE SEJA CRIADO O CÓDIGO AUTO NUMERAÇÃO EM LINE_ID

   rsNew("LINE_ID").value = object_id

   rsNew.Update
   
  
   rsNew.Close
   Set rsNew = Nothing
   
  
   PmNewLine = True

End Function
Public Function PmNewPoint(LayerName As String, Geometry As TypeGeometry, object_id As String) As Boolean
   
   'CRIA UMA NOVA LINHA NA TABELA DE ATRIBUTOS
   
   Dim Object_id_  As String
   Dim rsNew As ADODB.Recordset
   Set rsNew = New ADODB.Recordset
   
   rsNew.Open LayerName, Conn, adOpenKeyset, adLockOptimistic, adCmdTable
   
   rsNew.AddNew
                                          'INSERE MESMO CÓDIGO DO AUTO NUMBER DA TABELA DE GEOMETRIAS
   rsNew("OBJECT_ID_").value = object_id  '                                                       XXXXX strUser & Now ' INSERE VALOR EM OBJECT_ID_ PARA QUE SEJA CRIADO O CÓDIGO AUTO NUMERAÇÃO EM LINE_ID
      
   rsNew.Update
   
   rsNew.Close
   Set rsNew = Nothing
   
   
   PmNewPoint = True

End Function
Private Sub LoadAttributeByReference(geom_id As Long)


' FUNÇÃO ESTÁ DESABILITADA


   On Error GoTo Trata_Erro
   
   If CDbl(geom_id) = 0 Then
      Exit Sub
   End If
   
   Dim rsPoints As New ADODB.Recordset, rsRef As New ADODB.Recordset
   Dim id_Type As Integer, Id_SubType As Integer, value As Integer
   Dim layer_id As Integer, Attr_Link As String
   
   If cgeo.GetLayerAttrib("WATERCOMPONENTS", layer_id, Attr_Link) Then
   If frmCanvas.TipoConexao <> 4 Then
      rsPoints.Open "SELECT * from points" & layer_id & " inner join watercomponents on object_id = object_id_ where geom_id=" & geom_id, Conn, adOpenKeyset, adLockOptimistic, adCmdText
      Else
      a = "points"
      b = "WATERCOMPONENTS"
      c = "OBJECT_ID_"
      d = "obect_id"
      e = "geom_id"
      rsPoints.Open "SELECT * from " + """" + a + layer_id + """" + " inner join " + """" + b + """" + " on " + """" + d + """" + " = " + """" + c + """" + " where " + """" + e + """" + "='" & geom_id & "'", Conn, adOpenKeyset, adLockOptimistic, adCmdText
      End If
      If rsPoints.EOF = False Then
         
         If rsPoints!id_Type > 0 Then
            rsPoints.Close
            Set rsPoints = Nothing
            Exit Sub
         End If
         
         With tdbconref
            
            .setCurrentLayer "WATERCOMPONENTS"
            
            If .locateGeometry(rsPoints!X, rsPoints!Y, 4, 0.7) Then
               If cgeo.GetLayerAttrib("CEL_PONTO_ATRIB", layer_id, Attr_Link) Then
                  rsRef.Open "SELECT * from CEL_PONTO_ATRIB where " & Attr_Link & "=" & .objectIds(0), Conn, adOpenKeyset, adLockOptimistic
                  If GetTypeObjectByReference(rsRef!NOME_CELUL, id_Type, Id_SubType, value) Then
                       If frmCanvas.TipoConexao <> 4 Then
                        Conn.execute "update watercomponents set id_type=" & id_Type & _
                                                                 ", angle = " & Replace(rsRef!ANGULO, ",", ".") & _
                                                                 ", NOME_CELUL = '" & Replace(rsRef!NOME_CELUL, ",", ".") & _
                                                                 "', ORIGEM_CAL = '" & Replace(rsRef!ORIGEM_CAL, ",", ".") & _
                                                                 "', X_ = " & Replace(rsRef!X_, ",", ".") & _
                                                                 ", Y_ = " & Replace(rsRef!Y_, ",", ".") & _
                                                                 ", COR = " & Replace(rsRef!COR, ",", ".") & _
                                                                 ", TAMANHO_X = " & Replace(rsRef!TAMANHO_X, ",", ".") & _
                                                                 ", TAMANHO_Y = " & Replace(rsRef!TAMANHO_Y, ",", ".") & _
                                                                 ", CENT_CEL_X = " & Replace(rsRef!CENT_CEL_X, ",", ".") & _
                                                                 ", CENT_CEL_Y = " & Replace(rsRef!CENT_CEL_Y, ",", ".") & _
                                                                 ", COR_CELULA = " & Replace(rsRef!COR_CELULA, ",", ".") & _
                                                                 ", ESC_CEL_X = " & Replace(rsRef!ESC_CEL_X, ",", ".") & _
                                                                 ", ESC_CEL_Y = " & Replace(rsRef!ESC_CEL_Y, ",", ".") & _
                                                                 " where object_id_ =" & rsPoints!object_id
                                                                 
                                                                 
Else
                                                                 
Dim a1 As String
Dim b1 As String
Dim c1 As String
Dim d1 As String
Dim e1 As String
Dim f1 As String
Dim g1 As String
Dim h1 As String
Dim i1 As String
Dim j1 As String
Dim k1 As String
Dim l1 As String
                                                                 
Dim a2 As String
Dim b2 As String
Dim c2 As String
Dim d2 As String
Dim e2 As String
Dim f2 As String
Dim g2 As String
Dim h2 As String
Dim i2 As String
Dim j2 As String
Dim k2 As String
Dim l2 As String
Dim m2 As String
 Dim m3 As String
 Dim m4 As String
  Dim m5 As String
a = "WATERCOMPONENTS"
b = "ID_TYPE"
c = id_Type
d = "'c'"
e = "ANGLE"
f = Replace(rsRef!ANGULO, ",", ".")
g = "'f'"
h = "NOME_CELUL"
i = Replace(rsRef!NOME_CELUL, ",", ".")
j = "'i'"
k = "X_"
l = "Y_"
m = "COR"
a1 = "TAMANHO_X"
b1 = "TAMANHO_Y"
c1 = "CENT_CEL_X"
d1 = "CENT_CEL_X"
e1 = "ANGLE"
f1 = "CENT_CEL_X"
g1 = "CENT_CEL_Y"
h1 = "COR_CELULA"
i1 = "ESC_CEL_X"
j1 = "ESC_CEL_Y"
k1 = "OBJECT_ID_"
l1 = Replace(rsRef!ORIGEM_CAL, ",", ".")
a2 = Replace(rsRef!X_, ",", ".")
b2 = Replace(rsRef!Y_, ",", ".")
c2 = Replace(rsRef!COR, ",", ".")
d2 = Replace(rsRef!TAMANHO_X, ",", ".")
e2 = Replace(rsRef!TAMANHO_Y, ",", ".")
f2 = Replace(rsRef!CENT_CEL_X, ",", ".")
g2 = Replace(rsRef!CENT_CEL_Y, ",", ".")
h2 = Replace(rsRef!COR_CELULA, ",", ".")
i2 = Replace(rsRef!ESC_CEL_X, ",", ".")
j2 = Replace(rsRef!ESC_CEL_Y, ",", ".")
h2 = rsPoints!object_id
l2 = h2
m2 = "id_type"
m3 = "ORIGEM_CAL"
m4 = l1
m5 = i


Conn.execute "update " + """" + a + """" + " set " + """" + m2 + """" + " ='" & id_Type & _
"', " + """" + e + """" + " = '" & f & _
"'," + """" + h + """" + " =  '" & m5 & _
"', " + """" + m3 + """" + " = '" & l1 & _
"'," + """" + k + """" + " = '" & a2 & _
"', " + """" + l + """" + " = '" & b2 & _
"', " + """" + m + """" + " = '" & c2 & _
"', " + """" + a1 + """" + " = '" & d2 & _
"', " + """" + b1 + """" + " = '" & e2 & _
"', " + """" + i2 + """" + " = '" & f2 & _
"', " + """" + g1 + """" + " = '" & g2 & _
"', " + """" + h1 + """" + " = '" & h2 & _
"', " + """" + i1 + """" + " = '" & i2 & _
"', " + """" + j1 + """" + " = '" & j2 & _
"' where " + """" + k1 + """" + " ='" & l2 & "'"
                                                                 
End If
                                                                 
                                                                 
                                         
                                                                 
                                                                 
                                                                 
                     If Id_SubType > 0 Then
                     
                       
      b = "WATERCOMPONENTSDATA"
      c = "OBECT_ID"
      d = "ID_TYPE"
      e = "ID_SUBTYPE"
      f = "VALUE_"
      


     If frmCanvas.TipoConexao <> 4 Then
         
     Conn.execute "Insert Into WaterComponentsData (object_id_,id_type,id_subtype,value_)" & _
                           "values(" & rsPoints!object_id & "," & id_Type & "," & Id_SubType & "," & value & ")"
     
     Else
     
      Conn.execute "Insert Into " + """" + b + """" + "(" + """" + c + """" + "," + """" + d + """" + "," + """" + e + """" + "," + """" + f + """" + ")" & _
                           "values('" & rsPoints!object_id & "','" & id_Type & "','" & Id_SubType & "','" & value & "')"
     End If
                     
                     
                     
                        
                     End If
                  End If
                  rsRef.Close
               End If
            End If
         End With
      End If
   End If
   

Trata_Erro:
    If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
    Else
        'xxxx mantido do tratamento antigo
        If Not rsRef Is Nothing Then
           If rsRef.State = 1 Then rsRef.Close
           Set rsRef = Nothing
        End If
        If Not rsPoints Is Nothing Then
           If rsPoints.State = 1 Then rsPoints.Close
           Set rsPoints = Nothing
        End If
        'xxxx
      
         PrintErro "clsTerraLib", "Private Sub LoadAttributeByReference", CStr(Err.Number), CStr(Err.Description), True
       
    End If

End Sub

Function GetTypeObjectByReference(Description As String, id_Type As Integer, Id_SubType As Integer, value As Integer) As Boolean
   GetTypeObjectByReference = True
   Select Case Description
      Case "ADAPTA"
         id_Type = 31 'Adaptador
      Case "BOOSTE"
         id_Type = 36 'Booster
      Case "CRUSEM", "CRUZET"
         id_Type = 18 'cruzeta
         Id_SubType = 4
         value = 2
      Case "CUR011"
         id_Type = 18 'cruzeta
         Id_SubType = 4
         value = 7
      Case "CUR022"
         id_Type = 18 'cruzeta
         Id_SubType = 4
         value = 6
      Case "CUR045", "CUR315"
         id_Type = 18 'cruzeta
         Id_SubType = 4
         value = 4
      Case "CUR090", "CUR270"
         id_Type = 18 'cruzeta
         Id_SubType = 4
         value = 3
      Case "CUR337"
         id_Type = 18 'cruzeta
         Id_SubType = 4
         value = 6
      Case "CUR349"
         id_Type = 18 'cruzeta
         Id_SubType = 4
         value = 7
      Case "DESCA1", "DESCAR", "DESCAR"
         id_Type = 23
      Case "GIBAUL"
         id_Type = 34
      Case "HIDRAN"
         id_Type = 2
      Case "HIDROM"
         id_Type = 28
      Case "HIDSUB"
         id_Type = 26
      Case "HIDVAL"
         id_Type = 25
      Case "LUVA", "LUVA__"
         id_Type = 33
      Case "MACROM"
         id_Type = 38
      Case "RED", "REDUCA"
         id_Type = 37
      Case "REGGEN", "REGIST", "RREGIS"
         id_Type = 1
      Case "TAP"
         id_Type = 32
      Case "TE", "TEE001", "TEE002"
         id_Type = 18
         Id_SubType = 4
         value = 1
      Case "VALESF"
         id_Type = 1
         Id_SubType = 1
         value = 2
      Case "VALRED"
         id_Type = 18
         Id_SubType = 5
         value = 5
      Case "VALRET"
         id_Type = 1
         Id_SubType = 1
         value = 4
      Case "VALVA", "VALVUL"
         id_Type = 1
         Id_SubType = 1
         value = 0
      Case "VENTOS"
         id_Type = 22
      Case "VENTUR"
         id_Type = 29
      Case "PLUGUE" ' CAP
         id_Type = 35
      Case "AMARRA", "ARVORE", "BL2", "BOCALE", "BOCALO", "CV", "CXDRE", "CXDREE", _
         "CAV", "CAVAL", "CAVAL_", "$D", "$U", "A3", "ACO", "EPS", _
         "FILTRO", "FROM_B", "INCE", "MADE", "MARC", "NORTE", "PIQUET", "PLUGUE", "POSTE", "PV", "PVDRE", _
         "SETA", "SETDRE", "SETESG"
         GetTypeObjectByReference = False
   End Select
      
End Function



