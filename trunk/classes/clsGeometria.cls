VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGeometria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private conexaoTdb As TeDatabase     'cria uma variável privada e interna a classe
Private xi As Double
Private yi As Double
Private xf As Double
Private yf As Double
' Recebe o objeto da conexão TeDatabase
'
Public Property Let conexao(ByVal vData As TeDatabase)
    Set conexaoTdb = vData
End Property
'Retorna valor
Property Get x() As Integer
    x = Year(Now)
End Property
'  Recebe valor de x inicial
'
Property Let coordX(ByVal xInicial As Double)
    xi = xInicial
End Property
'  Recebe valor de y inicial
'
Property Let coordY(ByVal yInicial As Double)
    yi = yInicial
    yi = xi
End Property
Public Function GetStartEndPoints(objId As String) As Integer
    Dim retorno As Boolean
    
    If conexaoTdb.setCurrentLayer("Waterlines") Then
        retorno = conexaoTdb.getPointOfLine(0, objId, 0, xi, yi)
        retorno = conexaoTdb.getPointOfLine(0, objId, conexaoTdb.getQuantityPointsLine(0, "21") - 1, xf, yf)
    End If
End Function
' Calcula a coordenada x, y da interseção de duas linhas
'
'
'
Public Sub Intersecao2Linhas(L1xi As Double, L1yi As Double, L1xf As Double, L1yf As Double, L2xi As Double, L2yi As Double, L2xf As Double, L2yf As Double)
    Dim Xint As Double
    Dim Yint As Double
    Dim a1 As Double
    Dim b1 As Double
    Dim c1 As Double
    Dim a2 As Double
    Dim b2 As Double
    Dim c2 As Double
    Dim det As Double
    
    a1 = L1yf - L1yi
    b1 = L1xf - L1xi
    c1 = a1 * L1xi + b1 * L1yi
    
    a2 = L2yf - L2yi
    b2 = L2xf - L2xi
    c2 = a2 * L2xi + b2 * L2yi
    
    det = a1 * b2 - a2 * b1
    If det = 0 Then
        MsgBox "Linhas paralelas"
    Else
        Xint = (b2 * c1 - b1 * c2) / det
        Yint = (a1 * c2 - a1 * c1) / det
    End If
End Sub
' Calcula a distância entre dois pontos
'
'
'
Public Function Distancia2Pontos(xi As Double, yi As Double, xf As Double, yf As Double) As Double
    Distancia2Pontos = Sqr((xf - xi) * (xf - xi) + (yf - yi) * (yf - yi))
End Function
' Dada uma linha definida pelas coordenadas finais e iniciais, qual a coordenada x, y sobre a linha a uma dada distância
'
'
'
Public Sub CoordXYaDistLinha(xi As Double, yi As Double, xf As Double, yf As Double, d As Double)
    Dim x As Double
    Dim y As Double
    
    x = xi * (1 - d) + xf * d
    y = yi * (1 - d) + yf * d
End Sub
' Calcula a nova distância do início de um ramal de água após mover a rede
' Retorna a nova distância do ramal em relação ao início da linha
'
' b - comprimento da linha depois de movida
' a - comprimento inicial da linha antes de mover
' d1 - distância do ramal para o início da linha
'
Public Function NovaDistanciaRamal(b As Double, a As Double, d1 As Double) As Double
    NovaDistancia = b * d1 / a
End Function
