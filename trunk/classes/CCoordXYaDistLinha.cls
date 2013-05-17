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
' The function above takes a startPoint (x,y) and an endPoint (x,y) and a distance (from the endpoint.
' If your distance is negative, the returned point will be beyond the endPoint along the same line.
' If your distance is greater than the distance between startPoint and endPoint, the return point will be before your startpoint,
' but still on the same line. If your distance is a positive number and is less than the distance between startPoint and endPoint,
' the point returned will be on the line segment between the startPoint and endPoint at 'distance' from the endpoint.
'
' xi - coordenada x do ponto inicial da linha
' yi - coordenada y do ponto inicial da linha
' xf - coordenada x do ponto final da linha
' yf - coordenada y do ponto final da linha
' distancia - distância do ponto inicial da linha onde serão obtidas as coordenadas xd e yd
'
Public Sub calculaCoord(linha As CLine2D, Distancia As Double)
    Dim x As Double
    Dim y As Double
    Dim z As Double
    Dim ratio As Double
    Dim deltaX As Double
    Dim deltaY As Double
    Dim xi As Double
    Dim yi As Double
    Dim xf As Double
    Dim yf As Double
    
    xi = linha.xi       'inverte pois a fórmula é para distância entre a partir do ponto final
    yi = linha.yi
    xf = linha.xf
    yf = linha.yf
    
    x = xf - xi
    y = yf - yi
    z = Sqr(x * x + y * y)
    ratio = Distancia / z
    deltaX = x * ratio
    deltaY = y * ratio
    
    xd = xf - deltaX
    yd = yf - deltaY
End Sub
