VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMoveTrechoRedeComRamal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Funç~oes para obter as coordenadas do ramal em sua nova posição, após movida a rede
'
Private dist As Double                              'distância do início do ramal antes de mover até o início do trecho antes de mover
' Calcula a distância entre o ponto inicial do trecho antes de ser movido e o ponto inicial da linha do ramal
'
' distancia - retorna a distância do início do ramal antes de mover até o iníco do trecho de rede antes de mover
' objIdTrecho - do trecho antes de mover
' objIdRamal - do ramal ates de mover
' teDatabaseConn - conexão com o banco de dados geográfico
'
Public Function distancia(objIdTrecho As String, objIdRamal As String) As Double
    ', ByVal teDatabaseConn As TeDatabase
    'Dim a As New CGeometria                             'declaração do objeto como do tipo classe CGeometria e criação do mesmo (New) através de uma auto instância
    Dim t As Integer
    Dim qtdPontos As Integer
    Dim ponto1 As New CPoint2D
    Dim linha1 As New CLine2D                       'linha original do trecho de rede
    Dim linha2 As New CLine2D                       'linha original do ramal, antes de mover, ligado ao trecho de rede
    Dim coordInt2Linhas As New CDuasLinhas
    Dim coordenadasLinha1 As New CCoordIniFinLinha  'coordenadas ao trecho de rede antes de mover
    Dim coordenadasLinha2 As New CCoordIniFinLinha  'coordenadas do ramal de rede antes de mover
    Dim distInicial As New CDistancia2Pontos

    'obtem as coordenadas inicial e final do trecho
    'coordenadasLinha1.conexao = teDatabaseConn
    coordenadasLinha1.GetStartEndPointsTrecho (objIdTrecho)
    Set linha1 = coordenadasLinha1.linha
    'obtem as coordenadas inicial e final do ramal que está conectado ao trecho de rede
    'coordenadasLinha2.conexao = teDatabaseConn
    coordenadasLinha2.GetStartEndPointsRamal (objIdRamal)
    Set linha2 = coordenadasLinha2.linha
    Call coordInt2Linhas.Intersecao2Linhas(linha1, linha2)      'calcula a coordenada de interseção entre as linhas do ramal e do trecho de rede
    distancia = distInicial.Distancia2Pontos(linha1.xi, linha1.xf, coordInt2Linhas.x, coordInt2Linhas.y)  'calcula a distância entre o ponto de interseção do ramal com o trecho de rede e o ponto inicial do trecho de rede
End Function
' Obtem as coordenadas iniciais e finais do ramal após a movimentação do trecho de rede
'
' distancia - nova distância do início do trecho de rede até o início do ramal na nova posição tanto do trecho quanto do ramal
' newObjIdTrecho - do trecho depois que moveu o mesmo
' teDatabaseConn - conexão com o banco de dados geográfico
'
Public Function coordsRamal(distancia As Double, newObjIdTrecho As String, ByVal teDatabaseConn As TeDatabase) As CLine2D
    Dim linha3 As New CLine2D                       'linha final movida do trecho de rede
    Dim linha4 As New CLine2D                       'linha final do ramal, movido, ligado ao trecho de rede
    Dim coordenadasLinha3 As New CCoordIniFinLinha  'coordenadas ao trecho de rede depois de mover
    Dim coordXYdist As New CCoordXYaDistLinha
    
    'obtem as coordenadas inicial e final do trecho de rede na nova posição ao qual foi movido
    coordenadasLinha3.conexao = teDatabaseConn
    coordenadasLinha3.GetStartEndPointsTrecho (newObjIdTrecho)               'calcula as coordenadas inicial e final do novo trecho de rede
    Set linha3 = coordenadasLinha3.linha                                'obtem as coordenadas
    Call coordXYdist.calculaCoord(linha3, distancia)
    coordsRamal.xi = coordXYdist.x                                      'nova coordenada x inicial do ramal movido
    coordsRamal.yi = coordXYdist.y                                      'nova coordenada y inicial do ramal movido
    coordsRamal.xf = linha3.xf                                          'coordenada x final do ramal, onde está o hidrômetro, que não foi movido, mesma que a anterior
    coordsRamal.yf = linha3.yf                                          'coordenada y final do ramal, onde está o hidrômetro, que não foi movido, mesma que a anterior
End Function
