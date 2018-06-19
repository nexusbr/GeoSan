VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDuasLinhas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Classe para calcular as coordenadas x e y da inteseção de duas linhas
'
Private xInt As Double
Private yInt As Double
'Retorna valor de x da insterseção das duas linhas
'
Property Get X() As Double
    X = xInt
End Property
'Retorna valor de y da insterseção das duas linhas
'
Property Get Y() As Double
    Y = yInt
End Property
' Calcula a coordenada x, y da interseção de duas linhas
'
'
'
Public Sub Intersecao2Linhas(linha1 As CLine2D, linha2 As CLine2D)
    On Error GoTo Trata_Erro:
    Dim a1 As Double
    Dim b1 As Double
    Dim c1 As Double
    Dim a2 As Double
    Dim b2 As Double
    Dim c2 As Double
    Dim det As Double
    
    a1 = linha1.yf - linha1.yi
    b1 = linha1.xf - linha1.xi
    c1 = a1 * linha1.xi + b1 * linha1.yi
    
    a2 = linha2.yf - linha2.yi
    b2 = linha2.xf - linha2.xi
    c2 = a2 * linha2.xi + b2 * linha2.yi
    
    det = a1 * b2 - a2 * b1
    If det = 0 Then
        MsgBox "Linhas paralelas"
    Else
        xInt = (b2 * c1 - b1 * c2) / det
        yInt = (a1 * c2 - a2 * c1) / det
    End If
    Exit Sub
    
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
   Else
      ErroUsuario.Registra "CDuasLinhas", "Intersecao2Linhas", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
   End If
End Sub
