VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLine2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private PointX As Double
Private PointY As Double
Private PontoInicial As New CPoint2D
Private PontoFinal As New CPoint2D
'Retorna valor de x inicial
'
Property Get xi() As Double
    xi = PontoInicial.x
End Property
'  Recebe valor de x inicial
'
Property Let xi(ByVal x As Double)
    PontoInicial.x = x
End Property
'Retorna valor de y inicial
'
Property Get yi() As Double
    yi = PontoInicial.y
End Property
'  Recebe valor de y inicial
'
Property Let yi(ByVal y As Double)
    PontoInicial.y = y
End Property
'Retorna valor de x final
'
Property Get xf() As Double
    xf = PontoFinal.x
End Property
'  Recebe valor de x final
'
Property Let xf(ByVal x As Double)
    PontoFinal.x = x
End Property
'Retorna valor de y final
'
Property Get yf() As Double
    yf = PontoFinal.y
End Property
'  Recebe valor de y final
'
Property Let yf(ByVal y As Double)
    PontoFinal.y = y
End Property
Private Sub Line2D()

End Sub
