VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCoordXYaDistLinha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Calcula a coordenada x e y a uma distância do ponto inicial de uma linha fornecida a partir de suas
' coordenadas inicial e final
'
'
Private xd As Double        'coordenada x a partir de uma distância do ponto inicial da linha
Private yd As Double        'coordenada y a partir de uma distância do ponto inicial da linha
'Retorna valor de x a uma distância do ponto inicial
'
Property Get x() As Double
    x = xd
End Property
'Retorna valor de x a uma distância do ponto inicial
'
Property Get y() As Double
    y = yd
End Property
' Dada uma linha definida pelas coordenadas finais e iniciais, qual a coordenada x, y sobre a linha a uma dada distância
'
' xi - coordenada x do ponto inicial da linha
' yi - coordenada y do ponto inicial da linha
' xf - coordenada x do ponto final da linha
' yf - coordenada y do ponto final da linha
' distancia - distância do ponto inicial da linha onde serão obtidas as coordenadas xd e yd
'
Public Sub calculaCoord(linha As CLine2D, distancia As Double)
    xd = linha.xi * (1 - d) + linha.xf * distancia
    yd = linha.yi * (1 - d) + linha.yf * distancia
End Sub
