VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPoint2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private PointX As Double
Private PointY As Double
'Retorna valor de x
'
Property Get X() As Double
    X = PointX
End Property
'  Recebe valor de x
'
Property Let X(ByVal X As Double)
    PointX = X
End Property
'Retorna valor de y
'
Property Get Y() As Double
    Y = PointY
End Property
'  Recebe valor de x
'
Property Let Y(ByVal Y As Double)
    PointY = Y
End Property
Private Sub Point2D()

End Sub

