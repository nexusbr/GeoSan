VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDistanciaEquivalente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Classe para calcular a nova distância do ramal, por regra de 3, após o trecho ser movido

' Calcula a nova distância do início de um ramal de água após mover a rede
'
' distanciaRamalDepoisMovido - retorna a nova distância do ramal em relação ao início da linha
' compLinhaAntes - comprimento da linha depois de movida
' compLinhaDepois - comprimento inicial da linha antes de mover
' distanciaRamalAntes - distância do ramal para o início da linha
'
Public Function distanciaRamalDepoisMovido(compLinhaAntes As Double, compLinhaDepois As Double, distanciaRamalAntes As Double) As Double
    NovaDistancia = compLinhaAntes * distanciaRamalAntes / compLinhaDepois
End Function
