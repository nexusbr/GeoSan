VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDistancia2Pontos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Classe para calcular a distância entre dois pontos
'
'
Private xi As Double
Private yi As Double
Private xf As Double
Private yf As Double
' Calcula a distância entre dois pontos
'
'
'
Public Function Distancia2Pontos(xi As Double, yi As Double, xf As Double, yf As Double) As Double
    On Error GoTo Trata_Erro:
    Distancia2Pontos = Sqr((xf - xi) * (xf - xi) + (yf - yi) * (yf - yi))
    Exit Function
    
Trata_Erro:
   If Err.Number = 0 Or Err.Number = 20 Then
       Resume Next
   Else
      ErroUsuario.Registra "CDistancia2Pontos", "Distancia2Pontos", CStr(Err.Number), CStr(Err.Description), True, glo.enviaEmails
   End If
End Function
