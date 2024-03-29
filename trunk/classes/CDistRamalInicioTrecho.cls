VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDistRamalInicioTrecho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Calcula a dist�ncia entre o ponto inicial do trecho antes de ser movido e o ponto inicial da linha do ramal
'
' distancia - retorna a dist�ncia do in�cio do ramal antes de mover at� o in�co do trecho de rede antes de mover
' objIdTrecho - do trecho antes de mover
' objIdRamal - do ramal ates de mover
' teDatabaseConn - conex�o com o banco de dados geogr�fico
'
Public Function Distancia(objIdTrecho As String, objIdRamal As String) As Double
    ', ByVal teDatabaseConn As TeDatabase
    'Dim a As New CGeometria                             'declara��o do objeto como do tipo classe CGeometria e cria��o do mesmo (New) atrav�s de uma auto inst�ncia
    On Error GoTo Trata_Erro:
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
    'obtem as coordenadas inicial e final do ramal que est� conectado ao trecho de rede
    'coordenadasLinha2.conexao = teDatabaseConn
    coordenadasLinha2.GetStartEndPointsRamal (objIdRamal)
    Set linha2 = coordenadasLinha2.linha
    Set coordInicialRamal = coordenadasLinha2.linha
    coordInt2Linhas.Intersecao2Linhas linha1, linha2       'calcula a coordenada de interse��o entre as linhas do ramal e do trecho de rede
    Distancia = distInicial.Distancia2Pontos(linha1.xi, linha1.yi, coordInt2Linhas.X, coordInt2Linhas.Y)  'calcula a dist�ncia entre o ponto de interse��o do ramal com o trecho de rede e o ponto inicial do trecho de rede
    Exit Function
    
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
   Else
      ErroUsuario.Registra "CDistRamalInicioTrecho", "Distancia", CStr(Err.Number), CStr(Err.Description) + " objIdTrecho: " + objIdTrecho + " objIdRamal: " + objIdRamal, True, glo.enviaEmails
   End If
End Function
